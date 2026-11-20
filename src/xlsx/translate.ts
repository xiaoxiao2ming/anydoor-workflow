/**
 * 公式平移 —— office-kit@0.18.0 的 dist/formula/translate 只有 .d.ts、没有实现文件，
 * `exports` 里也没有 `./formula` 子路径，所以那份 openpyxl Translator 移植拿不到。
 * 这里自己实现一份（零依赖）。
 *
 * 用途：参数表模板行向下复制时，把 A1 式公式平移到目标行，
 * 正确处理 `$` 绝对引用、`1:1` 行范围、`A:A` 列范围、`Sheet!A1` 跨表、命名区域。
 */

const CELL_RE = /(?<![A-Za-z0-9_.])(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})(?![\d(])/g

// 列字母/坐标换算统一走 a1.ts（原先是本文件私有副本，编辑器也要用，抽出来共用）
import { colToIndex, indexToCol } from './a1'

function a1ToRC(a1: string): [number, number] | null {
  const m = String(a1).match(/^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})/)
  if (!m) return null
  return [Number(m[2]), colToIndex(m[1])]
}

function shiftCol(col: string, abs: boolean, d: number): string {
  if (abs || !d) return col
  const i = colToIndex(col) + d
  if (i < 1) throw new Error('translateFormula: 列平移越界')
  return indexToCol(i)
}

function shiftRow(row: string, abs: boolean, d: number): string {
  if (abs || !d) return row
  const r = Number(row) + d
  if (r < 1) throw new Error('translateFormula: 行平移后 < 1')
  return String(r)
}

export interface TranslateOptions {
  /** 目标单元格，如 'B5'；给出则自动推导行列偏移 */
  dest?: string
  rowDelta?: number
  colDelta?: number
}

/**
 * 把定义在 origin 的公式平移到 dest。
 * 保护字符串字面量；`$` 绝对引用不偏移；命名区域/中文表名原样返回。
 */
export function translateFormula(formula: string, origin: string, opts: TranslateOptions = {}): string {
  if (!formula || !formula.startsWith('=')) return formula
  let rd = opts.rowDelta ?? 0
  let cd = opts.colDelta ?? 0
  if (opts.dest) {
    const o = a1ToRC(origin)
    const d = a1ToRC(opts.dest)
    if (!o || !d) throw new Error(`translateFormula: origin/dest 不是合法 A1（${origin} / ${opts.dest}）`)
    rd = d[0] - o[0]
    cd = d[1] - o[1]
  }
  if (!rd && !cd) return formula

  // 拆出字符串字面量，避免误伤（"$B$2" 这种文本不能被当成引用平移）
  const parts: { t: 'code' | 'str'; s: string }[] = []
  const STR = /"([^"]|"")*"/g
  let last = 0
  let m: RegExpExecArray | null
  while ((m = STR.exec(formula))) {
    if (m.index > last) parts.push({ t: 'code', s: formula.slice(last, m.index) })
    parts.push({ t: 'str', s: m[0] })
    last = m.index + m[0].length
  }
  if (last < formula.length) parts.push({ t: 'code', s: formula.slice(last) })

  // 注意：必须在完整的 code 段上跑正则，保留前后字符上下文。
  // 若先按 token 切片再替换，'LOG10(' 的 '(' 就不在切片里了，函数名会被误当成单元格。
  return parts
    .map((p) =>
      p.t === 'str'
        ? p.s
        : p.s
            .replace(
              /(?<![A-Za-z0-9_.])(\$?)(\d+):(\$?)(\d+)(?![\d])/g,
              (_m, d1, r1, d2, r2) => d1 + shiftRow(r1, !!d1, rd) + ':' + d2 + shiftRow(r2, !!d2, rd)
            )
            .replace(
              /(?<![A-Za-z0-9_.])(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9])/g,
              (_m, d1, c1, d2, c2) => d1 + shiftCol(c1, !!d1, cd) + ':' + d2 + shiftCol(c2, !!d2, cd)
            )
            .replace(
              CELL_RE,
              (_m, dl: string, col: string, dr: string, row: string) =>
                (dl || '') + shiftCol(col, !!dl, cd) + (dr || '') + shiftRow(row, !!dr, rd)
            )
    )
    .join('')
}
