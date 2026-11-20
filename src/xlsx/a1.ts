/**
 * A1 坐标工具 —— xlsx 侧的公共小工具。
 *
 * 之前这些函数在 translate.ts 里是私有的，编辑器要用只能复制一份。
 * 抽出来共用，避免两处实现慢慢漂移。
 * 纯函数，零依赖，浏览器/Node 同构。
 */

/** 'A' → 1，'Z' → 26，'AA' → 27 */
export function colToIndex(letters: string): number {
  let n = 0
  for (const ch of letters.toUpperCase()) {
    if (ch < 'A' || ch > 'Z') throw new Error(`非法列字母：${letters}`)
    n = n * 26 + (ch.charCodeAt(0) - 64)
  }
  return n
}

/** 1 → 'A'，26 → 'Z'，27 → 'AA' */
export function indexToCol(n: number): string {
  if (!Number.isInteger(n) || n < 1) throw new Error(`非法列序号：${n}`)
  let s = ''
  while (n > 0) {
    const r = (n - 1) % 26
    s = String.fromCharCode(65 + r) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

/** 'B3' → { row: 3, col: 2, colLetter: 'B' }；非法返回 null */
export function parseA1(a1: string): { row: number; col: number; colLetter: string } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/.exec(String(a1).trim())
  if (!m) return null
  const col = colToIndex(m[1])
  const row = Number(m[2])
  if (row < 1) return null
  return { row, col, colLetter: m[1].toUpperCase() }
}

/** 'A1:C4' → 边界（含 $ 也认）；非法返回 null */
export function parseRange(ref: string): { minRow: number; maxRow: number; minCol: number; maxCol: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7}):\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/.exec(String(ref).trim())
  if (!m) return null
  const c1 = colToIndex(m[1])
  const r1 = Number(m[2])
  const c2 = colToIndex(m[3])
  const r2 = Number(m[4])
  return {
    minRow: Math.min(r1, r2),
    maxRow: Math.max(r1, r2),
    minCol: Math.min(c1, c2),
    maxCol: Math.max(c1, c2),
  }
}

/** 单格 'B3' 也当范围看（B3:B3） */
export function parseRefAsRange(ref: string): ReturnType<typeof parseRange> {
  if (parseRange(ref)) return parseRange(ref)
  const one = parseA1(ref)
  if (!one) return null
  return { minRow: one.row, maxRow: one.row, minCol: one.col, maxCol: one.col }
}

/** 范围 → 所有坐标（上限保护，避免超大范围把内存吃光） */
export function expandRange(ref: string, limit = 200000): string[] {
  const r = parseRefAsRange(ref)
  if (!r) return []
  const out: string[] = []
  for (let row = r.minRow; row <= r.maxRow; row++) {
    for (let col = r.minCol; col <= r.maxCol; col++) {
      if (out.length >= limit) return out
      out.push(indexToCol(col) + row)
    }
  }
  return out
}

/** 列字母偏移：colOffset('A', 2) → 'C' */
export function colOffset(col: string, offset: number): string {
  return indexToCol(colToIndex(col) + offset)
}
