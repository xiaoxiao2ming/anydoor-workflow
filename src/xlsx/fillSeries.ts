/**
 * #18 填充柄智能序列 —— 源块数字矩阵的严格等差识别与外推（纯计算层）。
 *
 * 语义（2026-10-07 拍板，对齐 Excel 默认行为）：
 * - 源片段（origin 块在该列/该行的切片）全为有限 number 且 ≥2 个、相邻差全相等
 *   → 等差序列，外推 = first + delta × k（k = 距 origin 上/左缘的格数；
 *   同一公式对源位也成立，源位 k 落在 [0, len) 自然取到源值，无需分周期）。
 * - 否则返回 null → 调用方平铺复制（周期 %oh/%ow，行为同 v1）：
 *   单值（1 → 1,1,1）、非数字（文本/公式/空）、差不等（1,2,4）都算不成序列。
 * - 只认 typeof === 'number'：文本数字 '3' 参与平铺不参与序列（同 Excel 文本语义）。
 * - 公式引用偏移不在本文件：UI 层走 translate.ts 的 translateFormula
 *   （fill 拖动只向右/下扩展，delta ≥ 0；角块公式双向偏移同样安全）。
 * - 角块（r、c 同时超出 origin）不做序列外推：源列/源行切片缺一 → null → 平铺。
 *
 * 纯计算层铁律：不 import electron / node: / 不碰 document。
 */

/** origin 块的数字矩阵：nums[r - r1][c - c1]，非数字/公式/空为 null */
export type NumMatrix = (number | null)[][]

export interface FillOrigin {
  r1: number
  c1: number
  r2: number
  c2: number
}

/** 识别一维片段的严格等差 delta；不成序列返回 null（单值 / 含非数字 / 相邻差不全等） */
export function seriesDelta(seg: (number | null)[]): number | null {
  if (seg.length < 2) return null
  for (const v of seg) {
    if (v == null || !Number.isFinite(v)) return null
  }
  const d = (seg[1] as number) - (seg[0] as number)
  for (let i = 2; i < seg.length; i++) {
    if ((seg[i] as number) - (seg[i - 1] as number) !== d) return null
  }
  return d
}

/**
 * 目标格 (r, c) 的智能序列取值；null = 调用方平铺复制。
 * 纵向扩展带（r > r2 且 c ≤ c2）→ 列切片外推；横向扩展带（c > c2 且 r ≤ r2）→ 行切片外推；
 * 角块 / origin 内 / 不成序列 → null。
 */
export function fillSmartNumber(nums: NumMatrix, o: FillOrigin, r: number, c: number): number | null {
  const inOriginCols = c >= o.c1 && c <= o.c2
  const inOriginRows = r >= o.r1 && r <= o.r2
  if (r > o.r2 && inOriginCols) {
    // 纵向扩展带：列切片 nums[0..oh-1][c - c1]
    const ci = c - o.c1
    const seg = nums.map((row) => row[ci] ?? null)
    const d = seriesDelta(seg)
    if (d != null) return (seg[0] as number) + d * (r - o.r1)
    return null
  }
  if (c > o.c2 && inOriginRows) {
    // 横向扩展带：行切片 nums[r - r1][0..ow-1]
    const ri = r - o.r1
    const row = nums[ri]
    if (!row) return null
    const d = seriesDelta(row)
    if (d != null) return (row[0] as number) + d * (c - o.c1)
    return null
  }
  return null // origin 内 / 角块 → 平铺
}
