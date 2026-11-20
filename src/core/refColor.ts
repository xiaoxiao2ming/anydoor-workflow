/**
 * 参数引用色（单一真源）
 *
 * 立档：2026-10-17 · 用户「能不能对参数引用的地址加个颜色区分」
 *
 * ★ 为什么单独一个文件（而不是放在 ooDomains.ts）：
 *   ooDomains 依赖 jszip（重）。引用面板（RefPanel）只想画个小色块，
 *   若从 ooDomains 引，会把 jszip 拖进面板所在的 chunk。
 *   色算法本身零依赖，拆出来后两边都引同一份 —— **色必须是同一个色**，
 *   否则「面板上的色」和「文档里的底色」对不上，图例就失去意义。
 */

/** 引用格底色（ARGB 浅色系：黑字都清楚，六种差异足够分辨） */
export const REF_PALETTE = [
  'FFFFF2CC', // 浅黄
  'FFCCE5FF', // 浅蓝
  'FFD9F2D9', // 浅绿
  'FFE8DDFF', // 浅紫
  'FFFFE0CC', // 浅橙
  'FFFFD6E7', // 浅粉
]

/**
 * 定义名称 → 底色（djb2 稳定哈希：同一参数永远同一个色）
 * 注意：改动算法会让所有已有文档的颜色变一次，需同步告知用户。
 */
export function refColorOf (name: string): string {
  let h = 5381
  for (let i = 0; i < name.length; i++) h = (((h << 5) + h) ^ name.charCodeAt(i)) >>> 0
  return REF_PALETTE[h % REF_PALETTE.length]
}
