/**
 * Excel 侧的数据构建：把引擎算好的上游数据按绑定映射翻译成
 * 「参数表 + 引用上下文」，交给 bindXlsxFromParamTable 写盘。
 *
 * 与 docx/buildData.ts 共用 src/ref/context.ts 的上游上下文 ——
 * 两边都不再自己算 ownerOf / columnOf / ctx，否则必然走偏。
 *
 * Excel 侧比 Word 侧多的两件事：
 *   ① Table（行级域）既能放标量也能放数组，所以要按 slot.kind 分流
 *   ② Table 可以带计算列公式（Word 没有公式引擎）
 */

import type { TableSchema, XlsxBinding } from '../core/schema'
import type { Row } from '../core/store'
import { ParamTable, type ParamSpec } from '../ref/paramTable'
import { aggregate, buildUpstream, resolveBindingTarget, type Agg } from '../ref/context'
import type { XlsxFieldSpec } from './field'

export interface BuildXlsxResult {
  table: ParamTable
  /** key → 取值说明，供 UI 显示「这个域取到了什么」 */
  resolved: Record<string, string>
  missing: string[]
  /** Table 的计算列公式，直接喂 extras.tableFormulas */
  formulas: Record<string, Record<string, string>>
}

export function buildXlsxParamTable(
  slots: XlsxFieldSpec[],
  map: Record<string, XlsxBinding>,
  node: TableSchema,
  tables: Record<string, TableSchema>,
  records: Record<string, Row[]>,
): BuildXlsxResult {
  const ub = buildUpstream(node, tables, records)

  const specs: ParamSpec[] = []
  const resolved: Record<string, string> = {}
  const missing: string[] = []
  const formulas: Record<string, Record<string, string>> = {}

  for (const slot of slots) {
    const spec = map[slot.key]

    // ── 手工值 ────────────────────────────────────────────────
    if (spec?.from === 'manual') {
      const v = spec.value ?? ''
      if (slot.kind === 'table') {
        // Table 需要数组；手工填 JSON 数组，且必须是对象数组
        try {
          const arr: unknown = v.trim() ? JSON.parse(v) : []
          const ok = Array.isArray(arr) && arr.every((r) => r !== null && typeof r === 'object')
          specs.push({ key: slot.key, default: ok ? arr : [] })
          resolved[slot.key] = ok ? `手工 ${(arr as unknown[]).length} 行` : '手工值不是对象数组'
          if (!ok) missing.push(slot.key)
        } catch {
          specs.push({ key: slot.key, default: [] })
          resolved[slot.key] = '手工 JSON 解析失败'
          missing.push(slot.key)
        }
      } else {
        specs.push({ key: slot.key, default: v })
        resolved[slot.key] = `手工值：${v.slice(0, 16) || '（空）'}`
      }
      continue
    }

    // ── 解析取值目标：source（URN）> fieldId > 名字自动匹配 ────
    const tgt = resolveBindingTarget(spec, slot.key, ub)

    // 非 node:// 的 URN（外部文件 / API / 全局参数）：本地没有对应字段，
    // 值只能靠 resolveUrn 在参数表解析时取。这里只登记引用，不给默认值 ——
    // 取不到时 trace 会如实标 unresolved，不静默留空。
    if (tgt.by === 'source' && tgt.fieldId === null && tgt.urn && !tgt.urn.startsWith('node://')) {
      specs.push({ key: slot.key, source: tgt.urn })
      resolved[slot.key] = tgt.urn
      continue
    }

    if (!tgt.fieldId) {
      missing.push(slot.key)
      resolved[slot.key] = spec ? '字段已失效' : '未绑定'
      continue
    }
    const fid = tgt.fieldId
    const fname = ub.fields.find((x) => x.id === fid)?.name ?? fid
    const rows = ub.rowsOf(fid)

    // ── 命名区域（标量域）────────────────────────────────────
    if (slot.kind === 'definedName') {
      if (spec?.from === 'rows') {
        missing.push(slot.key)
        resolved[slot.key] = '命名区域只能放标量，放不下整表'
        continue
      }
      const val = aggregate(rows, fid, spec?.agg as Agg | undefined)
      specs.push({ key: slot.key, source: tgt.urn ?? undefined, default: val })
      const suffix = spec?.agg && spec.agg !== 'first' ? `（${spec.agg}）` : ''
      resolved[slot.key] = tgt.urn
        ? `${fname} ← ${tgt.urn}${suffix}${tgt.by === 'auto' ? '（自动匹配）' : ''}`
        : `${fname}（无上游归属，用计算值）`
      continue
    }

    // ── Table（行级循环域）──────────────────────────────────
    const cols = spec?.columns
    const hasCols = !!cols && Object.keys(cols).length > 0
    const projected = hasCols
      ? rows.map((r) => {
          const o: Row = {}
          for (const [tplKey, srcId] of Object.entries(cols as Record<string, string>)) {
            o[tplKey] = r[srcId] ?? ''
          }
          return o
        })
      : rows
    specs.push({ key: slot.key, default: projected })
    resolved[slot.key] = `${projected.length} 行${hasCols ? '（已映射列）' : ''}`

    if (spec?.formulas && Object.keys(spec.formulas).length) {
      // 只保留模板里真实存在的列：写到不存在的列会静默失效，那是最难查的 bug
      const valid = slot.columns ?? []
      const kept: Record<string, string> = {}
      for (const [col, expr] of Object.entries(spec.formulas)) {
        if (valid.includes(col)) kept[col] = expr
        else missing.push(`${slot.key}.${col}（计算列不在模板里）`)
      }
      if (Object.keys(kept).length) formulas[slot.key] = kept
    }
  }

  const table = new ParamTable(specs, ub.ctx)
  // 引用没解析到值的域，一律计为 missing —— 不静默留空
  for (const t of table.trace()) {
    if (!t.ok && !missing.includes(t.key)) missing.push(t.key)
  }

  return { table, resolved, missing, formulas }
}
