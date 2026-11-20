/**
 * 域在工作表上的位置索引 —— 给编辑器画标记用。
 *
 * 用户在「打开 Excel」时最想知道的一件事：**这份模板里，哪些格子是域？**
 * 光有一份绑定清单看不出位置，所以要把 discoverXlsxFields 的结果
 * 换算成「工作表 → 坐标集合」，才能在网格上高亮出来。
 */

import { parseA1, parseRefAsRange, indexToCol } from './a1'
import type { XlsxFieldSpec } from './field'

export interface SheetMarks {
  /** 命名区域指向的单格：坐标 → 域名 */
  named: Map<string, string>
  /** Table 范围：域名 + 边界 */
  tables: { key: string; minRow: number; maxRow: number; minCol: number; maxCol: number; ref: string }[]
}

/** 该坐标是否落在任一 Table 范围内 */
export function inAnyTable(marks: SheetMarks | undefined, row: number, col: number): string | null {
  if (!marks) return null
  for (const t of marks.tables) {
    if (row >= t.minRow && row <= t.maxRow && col >= t.minCol && col <= t.maxCol) return t.key
  }
  return null
}

/** 坐标 → A1（1-based） */
export function coordOf(row: number, col: number): string {
  return indexToCol(col) + row
}

/**
 * 域的引用显示文本。
 *
 * 坑：命名区域的 `ref` 本身就带表名（`汇总表!$B$1`），而 Table 的 `ref` 只有范围（`A1:D2`）。
 * 无脑拼 `${sheet}!${ref}` 就会得到 `汇总表!汇总表!$B$1` —— 重复表名。
 * 所以：ref 已经带 `!` 就不再补前缀。
 */
export function displayRef(sheet: string | undefined, ref: string): string {
  const r = String(ref ?? '')
  if (r.includes('!')) return r
  return sheet ? `${sheet}!${r}` : r
}

/**
 * 把域清单整理成「工作表 → 标记」。
 * 认不出来的 ref 直接跳过（不猜），免得给用户标错位置。
 */
export function buildMarks(fields: XlsxFieldSpec[]): Map<string, SheetMarks> {
  const out = new Map<string, SheetMarks>()
  const get = (sheet: string): SheetMarks => {
    let m = out.get(sheet)
    if (!m) {
      m = { named: new Map(), tables: [] }
      out.set(sheet, m)
    }
    return m
  }

  for (const f of fields) {
    if (f.kind === 'definedName') {
      const sheet = f.sheet
      if (!sheet) continue
      // ref 形如 '汇总表!$B$1' 或 '$B$1'
      const cellRef = String(f.ref).split('!').pop() ?? ''
      const one = parseA1(cellRef)
      if (!one) continue
      get(sheet).named.set(coordOf(one.row, one.col), f.key)
    } else {
      const sheet = f.sheet
      if (!sheet) continue
      const r = parseRefAsRange(f.ref)
      if (!r) continue
      get(sheet).tables.push({ key: f.key, ref: f.ref, ...r })
    }
  }
  return out
}
