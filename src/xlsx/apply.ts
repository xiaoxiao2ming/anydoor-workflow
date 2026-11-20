/**
 * 触发计划的执行器 —— 把 planTrigger 算出的计划真的落进 Workbook。
 *
 * 与 trigger.ts 的分工：
 *   trigger.ts  = 判定（纯逻辑，可单测，不碰 Workbook）
 *   apply.ts    = 执行（写值 / 重算），仍然不碰 IO —— Workbook 由调用方传入
 *
 * 为什么不直接写进 WorkbookEditor：
 *   执行逻辑要能脱离 React 单测（改一格值、重算次数、跳过不存在的工作表…），
 *   放组件里就只能靠手点验证。
 *
 * 纯计算层铁律：不 import electron / node: / document。
 */

import { getSheet, sheetNames, type Workbook } from '@office-kit/xlsx/workbook'
import { setCellByCoord, getCellByCoord } from '@office-kit/xlsx/worksheet'
import { recomputeSheet, type RecomputeOpts } from './recompute'
import { parseA1 } from './a1'
import type { TriggerPlan } from './trigger'

export interface ApplyResult {
  /** 真正写入的格数（跳过了空表/坏坐标的） */
  written: number
  /** 各工作表重算结果 */
  recomputed: Record<string, { formulas: number; updated: number; skippedCrossSheet: number }>
  /** 跳过的写入及原因，面板要如实告诉用户「为什么没写进去」 */
  skipped: { target: string; reason: string }[]
}

export interface ApplyOpts extends RecomputeOpts {
  /**
   * 干跑：不真的写，只回报「会写什么」。
   * UI 用它做「确认再落盘」—— 参数改了先给用户看清单，认可了才真写。
   */
  dryRun?: boolean
}

/**
 * 执行一份计划。
 *
 * 返回值里 skipped 是刻意设计的：写不进去的格子必须**报告**而不是静默吞掉，
 * 否则用户会以为「规则命中了但文件没变」，进而怀疑整套机制。
 */
export function applyTriggerPlan(
  wb: Workbook,
  plan: TriggerPlan,
  opts: ApplyOpts = {}
): ApplyResult {
  const written: string[] = []
  const skipped: { target: string; reason: string }[] = []
  const recomputed: ApplyResult['recomputed'] = {}

  const known = new Set(sheetNames(wb))

  /* ── ① 写值 ── */
  for (const w of plan.writes) {
    if (!known.has(w.sheet)) {
      skipped.push({ target: `${w.sheet}!${w.coord}`, reason: `工作表「${w.sheet}」不存在` })
      continue
    }
    const ws = getSheet(wb, w.sheet)
    if (!ws) {
      skipped.push({ target: `${w.sheet}!${w.coord}`, reason: `工作表「${w.sheet}」取不到` })
      continue
    }
    if (!parseA1(w.coord)) {
      skipped.push({ target: `${w.sheet}!${w.coord}`, reason: `坐标「${w.coord}」不是合法 A1 引用` })
      continue
    }

    if (opts.dryRun) {
      written.push(`${w.sheet}!${w.coord}`)
      continue
    }

    // 已有公式的格子不能直接覆盖值 —— 那样公式就没了，等于静默毁掉用户模型。
    // 遇到这种格：跳过并说明原因，让用户决定（这属于要人工确认的破坏性操作）。
    const cur = getCellByCoord(ws, w.coord)
    const curVal = cur && typeof cur === 'object' ? (cur as { value?: unknown }).value : undefined
    if (curVal && typeof curVal === 'object' && (curVal as { kind?: string }).kind === 'formula') {
      skipped.push({ target: `${w.sheet}!${w.coord}`, reason: '该格是公式格，写值会覆盖公式（已跳过）' })
      continue
    }

    setCellByCoord(ws, w.coord, normalizeForCell(w.value))
    written.push(`${w.sheet}!${w.coord}`)
  }

  /* ── ② 重算 ── */
  const sheetsToCalc = new Set<string>([...plan.recomputeSheets])
  if (plan.recomputeAll) {
    for (const n of known) sheetsToCalc.add(n)
  }
  for (const name of sheetsToCalc) {
    if (!known.has(name)) {
      skipped.push({ target: `重算 ${name}`, reason: `工作表「${name}」不存在` })
      continue
    }
    if (opts.dryRun) {
      recomputed[name] = { formulas: 0, updated: 0, skippedCrossSheet: 0 }
      continue
    }
    try {
      const r = recomputeSheet(wb, name, { resolveUrn: opts.resolveUrn })
      recomputed[name] = { formulas: r.formulas, updated: r.updated, skippedCrossSheet: r.skippedCrossSheet }
    } catch (e) {
      // 重算失败不阻断整条链：保留原缓存值总比整表丢公式好（与 WorkbookEditor.commit 同口径）
      skipped.push({ target: `重算 ${name}`, reason: '重算失败：' + String(e).slice(0, 80) })
    }
  }

  return { written: written.length, recomputed, skipped }
}

/**
 * 写进单元格的值的归一化。
 *
 * 坑：直接 setCellValue 一个对象会产出非法 OOXML（OO 里 cell 只接受标量）。
 * 所以这里把对象/数组转成可读文本，undefined 归一成空串。
 */
export function normalizeForCell(v: unknown): string | number | boolean {
  if (v === null || v === undefined) return ''
  if (typeof v === 'number') return Number.isFinite(v) ? v : ''
  if (typeof v === 'boolean') return v
  if (typeof v === 'string') return v
  if (v instanceof Date) return v.toISOString().slice(0, 10)
  try {
    return JSON.stringify(v)
  } catch {
    return String(v)
  }
}

/**
 * 把源格里的公式文本也读出来（供 UI 展示「这格本来是公式，被覆盖了」）。
 * 单独暴露是因为 UI 要在**执行前**就预警，执行后就晚了。
 */
export function formulaAt(wb: Workbook, sheet: string, coord: string): string | null {
  const ws = getSheet(wb, sheet)
  if (!ws) return null
  const cell = getCellByCoord(ws, coord)
  if (!cell || typeof cell !== 'object') return null
  const v = (cell as { value?: unknown }).value
  if (v && typeof v === 'object' && (v as { kind?: string }).kind === 'formula') {
    return ((v as { formula?: string }).formula) ?? ''
  }
  return null
}
