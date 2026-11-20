/**
 * 工作表内公式重算 —— 编辑一格之后，把本表引用它的公式刷一遍。
 *
 * 为什么放在独立模块：WorkbookEditor 只该管交互；「收集值 → 迭代求值 → 写回缓存」
 * 是纯逻辑，值得单测与复用（将来流程编排的重算也走它）。
 *
 * 为什么多轮迭代：公式可互相引用且**声明顺序 ≠ 依赖顺序** —— 单趟求值会静默算错。
 * 项目铁律（calcTable）用的就是「迭代到无新进展」，这里同理，上限 MAX_ROUNDS 兜底。
 *
 * 第一版边界（刻意收窄）：
 *   · 只处理**同表** A1 引用；公式体含「!」（跨表引用）的保持原缓存值不动，
 *     计数返回，由 UI 提示「跨表引用暂不自动重算」。
 *   · shared 从属格（office-kit 已展开文本的照算；文本为空的保留原样）。
 *   · 解析失败的公式保留原缓存值 —— 不让一个坏公式污染整表。
 *
 * 纯计算层：不碰 electron / node: / document。
 */
import { getSheet } from '@office-kit/xlsx/workbook'
import { getCellByCoord, getValueExtent } from '@office-kit/xlsx/worksheet'
import { setFormula } from '@office-kit/xlsx/cell'
import { evalRowFormula } from './evalFormula'
import { indexToCol } from './a1'
import type { Workbook } from '@office-kit/xlsx/workbook'

const MAX_ROUNDS = 20

/** node:// 引用取值（批次 1-6 · 引用参与求值）—— 由调用方注入（编辑器持有 tables/records） */
export interface RecomputeOpts {
  resolveUrn?: (urn: string) => unknown
}

export interface RecomputeResult {
  /** 本表公式格总数（含跨表跳过的） */
  formulas: number
  /** 含跨表引用（公式体带「!」）、本轮保持不动的数量 */
  skippedCrossSheet: number
  /** 实际重算并写回缓存值的格子数 */
  updated: number
}

/** 收集一个工作表的「坐标 → 值」表：普通格取值，公式格取缓存值。 */
export function collectRefValues(
  wb: Workbook,
  sheetName: string,
  opts?: RecomputeOpts
): Record<string, unknown> {
  const ws = getSheet(wb, sheetName)
  const ref: Record<string, unknown> = {}
  if (!ws) return ref
  const e = getValueExtent(ws)
  if (!e) return ref
  for (let r = e.minRow; r <= e.maxRow; r++) {
    for (let c = e.minCol; c <= e.maxCol; c++) {
      const coord = indexToCol(c) + r
      const cell = getCellByCoord(ws, coord)
      if (!cell || typeof cell !== 'object') continue
      const v = (cell as { value?: unknown }).value
      if (v && typeof v === 'object' && (v as { kind?: string }).kind === 'formula') {
        const fv = v as { cachedValue?: unknown }
        ref[coord] = fv.cachedValue ?? ''
      } else if (
        opts?.resolveUrn &&
        typeof v === 'string' &&
        v.startsWith('node://')
      ) {
        // 引用格：A1 公式指向它时拿到的是上游值而不是 URN 文本（取不到保留原文）
        const rv = opts.resolveUrn(v)
        ref[coord] = rv == null ? v : rv
      } else {
        ref[coord] = v ?? ''
      }
    }
  }
  return ref
}

/**
 * 重算一个工作表内的全部同表公式，并把新缓存值写回（保留公式文本）。
 * 只写「值真的变了」的格子 —— 与 patchSave 配合时，改动面越小越好。
 */
export function recomputeSheet(wb: Workbook, sheetName: string, opts?: RecomputeOpts): RecomputeResult {
  const ws = getSheet(wb, sheetName)
  if (!ws) return { formulas: 0, skippedCrossSheet: 0, updated: 0 }
  const e = getValueExtent(ws)
  if (!e) return { formulas: 0, skippedCrossSheet: 0, updated: 0 }

  // ① 收集值表 + 公式格清单
  const refValues = collectRefValues(wb, sheetName, opts)
  const formulaCells: { coord: string; formula: string }[] = []
  let skippedCrossSheet = 0
  for (let r = e.minRow; r <= e.maxRow; r++) {
    for (let c = e.minCol; c <= e.maxCol; c++) {
      const coord = indexToCol(c) + r
      const cell = getCellByCoord(ws, coord)
      if (!cell || typeof cell !== 'object') continue
      const v = (cell as { value?: unknown }).value
      if (v && typeof v === 'object' && (v as { kind?: string }).kind === 'formula') {
        const f = ((v as { formula?: string }).formula) ?? ''
        if (f) {
          if (f.includes('!')) skippedCrossSheet++
          else formulaCells.push({ coord, formula: f })
        }
      }
    }
  }
  if (!formulaCells.length) return { formulas: 0, skippedCrossSheet, updated: 0 }

  // ② 多轮迭代到不动点
  const changed = new Map<string, unknown>()
  for (let round = 0; round < MAX_ROUNDS; round++) {
    let moved = false
    for (const { coord, formula } of formulaCells) {
      const out = evalRowFormula(formula, refValues, opts?.resolveUrn)
      if (!out.ok) continue // 解析失败：保留原缓存值
      const old = refValues[coord]
      if (String(out.value) !== String(old)) {
        refValues[coord] = out.value
        changed.set(coord, out.value)
        moved = true
      }
    }
    if (!moved) break
  }

  // ③ 写回：保留公式文本，只更新缓存值 —— 只动变过的格子
  let updated = 0
  for (const { coord, formula } of formulaCells) {
    const nv = changed.get(coord)
    if (nv === undefined) continue
    const cell = getCellByCoord(ws, coord)
    if (!cell) continue
    setFormula(
      cell,
      formula,
      { cachedValue: nv as number | string | boolean }
    )
    updated++
  }
  return { formulas: formulaCells.length, skippedCrossSheet, updated }
}
