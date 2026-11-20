/**
 * 工作簿编辑动作集 —— WorkbookEditor 工具栏的「结构与格式」操作。
 *
 * 原则：每个动作 = 「改 wb → 返回统计」，不碰 React、不碰保存
 * （保存统一走 WorkbookEditor 的 commit：recompute → patchSave → undo 快照）。
 *
 * 行列插入/删除用「尾部扩一格 + moveRange 整体搬移 + 清空腾出行」组合
 * （office-kit 没有 insertRowAt/insertColAt）。
 * ⚠ 已知限制（v1）：搬移只挪值与样式，**公式引用不自动偏移**
 *   （Excel 插行会把 =B5 调成 =B6，这里不会）—— 有公式的表慎用，后续版本补引用偏移。
 *
 * 纯计算层：不碰 electron / node: / document。
 */
import {
  getCellByCoord,
  getValueExtent,
  setCellByCoord,
  ensureCellByCoord,
  moveRange,
  mergeCells,
  unmergeCells,
  setRangeValues,
  setFreezePanes,
  appendRows,
} from '@office-kit/xlsx/worksheet'
import {
  getSheet,
  sheetNames,
  addWorksheet,
  removeSheet,
  renameSheet,
} from '@office-kit/xlsx/workbook'
import {
  setCellStyle,
  alignCellHorizontal,
  patchCellFont,
  getCellFont,
  getCellFill,
  getCellBorder,
  getCellNumberFormat,
  setRangeStyle,
  colorToHex,
  rgbColor,
  makePatternFill,
  makeBorder,
  makeSide,
  makeFont,
} from '@office-kit/xlsx/styles'
import type { FontPatch, UnderlineStyle, Color, SideStyle } from '@office-kit/xlsx/styles'
import type { Cell } from '@office-kit/xlsx/cell'
import { indexToCol, colToIndex } from './a1'
import type { Workbook } from '@office-kit/xlsx/workbook'

export interface CellRect {
  /** 1-based，含端点，已归一（r1≤r2、c1≤c2） */
  r1: number
  c1: number
  r2: number
  c2: number
}

const rng = (c1: number, r1: number, c2: number, r2: number) =>
  `${indexToCol(c1)}${r1}:${indexToCol(c2)}${r2}`

/** 解析 'B12' → {r:12, c:2} */
export function parseCoord(coord: string): { r: number; c: number } {
  const m = /^([A-Za-z]+)(\d+)$/.exec(coord.trim())
  if (!m) throw new Error('非法坐标：' + coord)
  return { r: Number(m[2]), c: colToIndex(m[1]) }
}

function extentMax(ws: Parameters<typeof getValueExtent>[0]) {
  const e = getValueExtent(ws)
  // 口径必须与网格显示一致（grid 至少渲染 30 行 × 12 列，见 WorkbookEditor 的 grid useMemo）：
  // 否则空表/小表上删列删行会因 maxC/maxR 过小而静默返回 —— 用户点了没反应（实测复现）。
  return {
    maxR: Math.max(e?.maxRow ?? 0, 30),
    maxC: Math.max(e?.maxCol ?? 0, 12),
  }
}

/** 清空一整行（写 null，不删格子本身） */
function clearRow(ws: never | Parameters<typeof setRangeValues>[0], r: number, maxC: number) {
  setRangeValues(
    ws,
    rng(1, r, maxC, r),
    [Array.from({ length: maxC }, () => null)] as never
  )
}

/** 清空一整列 */
function clearCol(ws: Parameters<typeof setRangeValues>[0], c: number, maxR: number) {
  const col: (null)[] = Array.from({ length: maxR }, () => null)
  const grid: (null)[][] = []
  for (let r = 1; r <= maxR; r++) grid.push(col)
  setRangeValues(ws, rng(c, 1, c, maxR), grid as never)
}

/* ────────────── 行 / 列 ────────────── */

/** 在第 r 行上方插入一行（下方整体下移，腾出的行清空） */
export function insertRowAbove(wb: Workbook, sheetName: string, r: number): number {
  const ws = getSheet(wb, sheetName)
  if (!ws) return 0
  const { maxR, maxC } = extentMax(ws)
  appendRows(ws, [Array.from({ length: maxC }, () => undefined)])
  const moved = moveRange(ws, rng(1, r, maxC, maxR), rng(1, r + 1, maxC, maxR + 1))
  clearRow(ws, r, maxC)
  return moved
}

/** 删除第 r 行（下方整体上移，尾行清空） */
export function deleteRow(wb: Workbook, sheetName: string, r: number): number {
  const ws = getSheet(wb, sheetName)
  if (!ws) return 0
  const { maxR, maxC } = extentMax(ws)
  if (maxR <= 1) return 0
  const moved = moveRange(ws, rng(1, r + 1, maxC, maxR), rng(1, r, maxC, maxR - 1))
  clearRow(ws, maxR, maxC)
  return moved
}

/** 在第 c 列左侧插入一列 */
export function insertColLeft(wb: Workbook, sheetName: string, c: number): number {
  const ws = getSheet(wb, sheetName)
  if (!ws) return 0
  const { maxR, maxC } = extentMax(ws)
  // 先把尾列逐行 ensure 出来（moveRange 的目标才落得住）
  for (let r = 1; r <= maxR; r++) ensureCellByCoord(ws, indexToCol(maxC + 1) + r)
  const moved = moveRange(ws, rng(c, 1, maxC, maxR), rng(c + 1, 1, maxC + 1, maxR))
  clearCol(ws, c, maxR)
  return moved
}

/** 删除第 c 列（右侧整体左移，尾列清空） */
export function deleteCol(wb: Workbook, sheetName: string, c: number): number {
  const ws = getSheet(wb, sheetName)
  if (!ws) return 0
  const { maxR, maxC } = extentMax(ws)
  if (maxC <= 1) return 0
  const moved = moveRange(ws, rng(c + 1, 1, maxC, maxR), rng(c, 1, maxC - 1, maxR))
  clearCol(ws, maxC, maxR)
  return moved
}

/* ────────────── 合并 ────────────── */

export function mergeSel(wb: Workbook, sheetName: string, rect: CellRect): string {
  const ws = getSheet(wb, sheetName)
  if (!ws) throw new Error('工作表不存在')
  if (rect.r1 === rect.r2 && rect.c1 === rect.c2) throw new Error('请先扩展选区（Shift+点击）再合并')
  const range = rng(rect.c1, rect.r1, rect.c2, rect.r2)
  mergeCells(ws, range)
  return range
}

export function unmergeSel(wb: Workbook, sheetName: string, rect: CellRect): boolean {
  const ws = getSheet(wb, sheetName)
  if (!ws) return false
  return unmergeCells(ws, rng(rect.c1, rect.r1, rect.c2, rect.r2))
}

/* ────────────── 格式（选区内逐格） ────────────── */

function eachCellInRect(
  wb: Workbook,
  sheetName: string,
  rect: CellRect,
  fn: (cell: Cell) => void,
): number {
  const ws = getSheet(wb, sheetName)
  if (!ws) return 0
  let n = 0
  for (let r = rect.r1; r <= rect.r2; r++) {
    for (let c = rect.c1; c <= rect.c2; c++) {
      const cell = getCellByCoord(ws, indexToCol(c) + r) ?? ensureCellByCoord(ws, indexToCol(c) + r)
      fn(cell)
      n++
    }
  }
  return n
}

export function alignSel(
  wb: Workbook,
  sheetName: string,
  rect: CellRect,
  horiz: 'left' | 'center' | 'right',
): number {
  return eachCellInRect(wb, sheetName, rect, (cell) => alignCellHorizontal(wb, cell, horiz))
}

export function numberFmtSel(wb: Workbook, sheetName: string, rect: CellRect, fmt: string): number {
  return eachCellInRect(wb, sheetName, rect, (cell) => setCellStyle(wb, cell, { numberFormat: fmt }))
}

/**
 * 字体格式（#33 合并写回）：只动调用方显式传入的键，其余字体属性保持现状。
 * v1 缺陷：用 setCellStyle 整只替换 font spec——只传 bold 会把字号/字体名一起重置回默认
 * （office-kit 注释明示「Reach for patchCellFont unless you really mean 'this font and nothing inherited'」）。
 * patchCellFont 语义：显式键生效、undefined 键 = 移除该属性、省略键 = 保持现值。
 */
export function fontSel(
  wb: Workbook,
  sheetName: string,
  rect: CellRect,
  font: { bold?: boolean; italic?: boolean; underline?: boolean; size?: number; color?: string },
): number {
  // FontPatch 属性带 readonly（继承 Font），逐字段赋值会 TS2540 —— 用可变中间类型构造，传入时断言
  const patch: {
    bold?: boolean
    italic?: boolean
    underline?: UnderlineStyle
    size?: number
    color?: { rgb: string }
  } = {}
  if ('bold' in font) patch.bold = font.bold
  if ('italic' in font) patch.italic = font.italic
  // OOXML 的下划线是样式枚举（single/double…），不是布尔
  if ('underline' in font) patch.underline = font.underline ? ('single' as UnderlineStyle) : undefined
  if ('size' in font) patch.size = font.size
  if ('color' in font) patch.color = font.color ? { rgb: font.color } : undefined
  return eachCellInRect(wb, sheetName, rect, (cell) => patchCellFont(wb, cell, patch as FontPatch))
}

/** 读首格当前字体（右键菜单/格式面板的「当前值」来源；#33 字体合并语义配套） */
export function fontOfFirst(wb: Workbook, sheetName: string, rect: CellRect) {
  const ws = getSheet(wb, sheetName)
  if (!ws) return null
  const cell = getCellByCoord(ws, indexToCol(rect.c1) + rect.r1)
  if (!cell) return null
  return getCellFont(wb, cell)
}

/* ────────────── 设置单元格格式对话框（#33 第二块：字体/边框/填充/数字）────────────── */

/** 对话框边框单边状态：style='none' = 无边线；color '#RRGGBB' 或 null（自动色） */
export interface FmtSideState {
  style: SideStyle | 'none'
  color: string | null
}

/** 格式对话框四轴状态（WPS 同款：读首格做初值，确定后四轴一次写入选区） */
export interface FmtState {
  font: {
    name: string
    size: number
    bold: boolean
    italic: boolean
    underline: boolean
    strike: boolean
    /** '#RRGGBB'；null = 自动色（不写 color 键） */
    color: string | null
  }
  /** '#RRGGBB'；null = 无填充 */
  fill: string | null
  border: { top: FmtSideState; bottom: FmtSideState; left: FmtSideState; right: FmtSideState }
  /** 格式代码：'General' 或 '#,##0.00' / '0.00%' / 'yyyy-mm-dd' 等 */
  numFmt: string
}

const DEFAULT_FMT_STATE = (): FmtState => ({
  font: { name: '', size: 11, bold: false, italic: false, underline: false, strike: false, color: null },
  fill: null,
  border: {
    top: { style: 'none', color: null },
    bottom: { style: 'none', color: null },
    left: { style: 'none', color: null },
    right: { style: 'none', color: null },
  },
  numFmt: 'General',
})

/** Color → '#RRGGBB'（theme/auto 色无主题解析不了，回退 null） */
const toHex6 = (color: Color | undefined): string | null => {
  const argb = color ? colorToHex(color) : undefined
  return argb ? '#' + argb.slice(2) : null
}

/**
 * 读首格四轴初值（混合选区显示首格格式，WPS 同语义；表不存在/空格返回默认值）。
 * 空格返回默认值而非 null —— 对话框始终可开（给空选区上格式是合法操作）。
 */
export function fmtOfFirst(wb: Workbook, sheetName: string, rect: CellRect): FmtState {
  const ws = getSheet(wb, sheetName)
  const cell = ws ? getCellByCoord(ws, indexToCol(rect.c1) + rect.r1) : null
  if (!cell) return DEFAULT_FMT_STATE()
  const f = getCellFont(wb, cell)
  const fill = getCellFill(wb, cell)
  const bd = getCellBorder(wb, cell)
  const side = (s: { style?: SideStyle; color?: Color } | undefined): FmtSideState => ({
    style: s?.style ?? 'none',
    color: toHex6(s?.color),
  })
  // 只有纯色填充能进对话框（渐变不在建模内，读作无填充）
  const solid = fill.kind === 'pattern' && fill.patternType === 'solid' ? fill.fgColor : undefined
  return {
    font: {
      name: f.name ?? '',
      size: f.size ?? 11,
      bold: !!f.bold,
      italic: !!f.italic,
      underline: !!f.underline,
      strike: !!f.strike,
      color: toHex6(f.color),
    },
    fill: toHex6(solid),
    border: { top: side(bd.top), bottom: side(bd.bottom), left: side(bd.left), right: side(bd.right) },
    numFmt: getCellNumberFormat(wb, cell) || 'General',
  }
}

/** '#RRGGBB' → 'FFRRGGBB'（对话框状态带 # 前缀，OOXML Color 要纯 hex） */
const argb = (hex6: string) => 'FF' + hex6.replace('#', '')

/**
 * 四轴一次写入选区（格式对话框确定按钮）。setRangeStyle 内部 ensureCell 建格 +
 * 按 spec 轴出补丁 xf 与现有 xf 合并 —— 空选区也能上格式。
 * ⚠ 已知限制（与 WPS 对话框同行为）：font/border 各自整轴注册进样式池，
 * 对话框没建模的属性（vertAlign/scheme/渐变填充/斜边框/保护）确定后会丢。
 */
export function applyFmtSel(wb: Workbook, sheetName: string, rect: CellRect, st: FmtState): void {
  const ws = getSheet(wb, sheetName)
  if (!ws) return
  const fside = (s: FmtSideState) =>
    s.style === 'none'
      ? undefined
      : makeSide({ style: s.style, color: s.color ? rgbColor(argb(s.color)) : undefined })
  setRangeStyle(wb, ws, rng(rect.c1, rect.r1, rect.c2, rect.r2), {
    font: makeFont({
      name: st.font.name || undefined,
      size: st.font.size,
      bold: st.font.bold || undefined,
      italic: st.font.italic || undefined,
      strike: st.font.strike || undefined,
      underline: st.font.underline ? ('single' as UnderlineStyle) : undefined,
      color: st.font.color ? rgbColor(argb(st.font.color)) : undefined,
    }),
    fill: st.fill
      ? makePatternFill({ patternType: 'solid', fgColor: rgbColor(argb(st.fill)) })
      : makePatternFill({ patternType: 'none' }),
    border: makeBorder({
      top: fside(st.border.top),
      bottom: fside(st.border.bottom),
      left: fside(st.border.left),
      right: fside(st.border.right),
    }),
    numberFormat: st.numFmt || 'General',
  })
}

/**
 * 插入单元格（#33）：在选区位置腾出空白，原内容按方向推移，被推出数据区的尾部内容丢弃（Excel 同款语义）。
 * down = 下方单元格下移（只动选区覆盖的列）；right = 右侧单元格右移（只动选区覆盖的行）。
 * 返回移动的格子数。已知限制沿用 v1：公式引用不自动偏移。
 */
export function insertCellShift(
  wb: Workbook,
  sheetName: string,
  rect: CellRect,
  dir: 'down' | 'right',
): number {
  const ws = getSheet(wb, sheetName)
  if (!ws) return 0
  const { maxR, maxC } = extentMax(ws)
  const nR = rect.r2 - rect.r1 + 1
  const nC = rect.c2 - rect.c1 + 1
  let moved = 0
  if (dir === 'down') {
    // 尾部扩 nR 行 → 选区以下（覆盖列范围内）整体下移 → 清空腾出的选区
    const rows: (undefined)[][] = []
    for (let i = 0; i < nR; i++) rows.push(Array.from({ length: maxC }, () => undefined))
    appendRows(ws, rows)
    moved = moveRange(ws, rng(rect.c1, rect.r1, rect.c2, maxR), rng(rect.c1, rect.r1 + nR, rect.c2, maxR + nR))
  } else {
    // 右移：逐行 ensure 尾部 nC 列（同 insertColLeft 对尾列的惯例），选区以右整体右移，尾部溢出丢弃
    for (let r = rect.r1; r <= rect.r2; r++) {
      for (let i = 1; i <= nC; i++) ensureCellByCoord(ws, indexToCol(maxC + i) + r)
    }
    moved = moveRange(ws, rng(rect.c1, rect.r1, maxC, rect.r2), rng(rect.c1 + nC, rect.r1, maxC + nC, rect.r2))
  }
  // 腾出的选区清空（moveRange 不清源——与 insertRowAbove 的 clearRow 同理）
  setRangeValues(
    ws,
    rng(rect.c1, rect.r1, rect.c2, rect.r2),
    Array.from({ length: nR }, () => Array.from({ length: nC }, () => null)) as never,
  )
  return moved
}

/* ────────────── 冻结 ────────────── */

/** 冻结到选中格的上方与左方；传 undefined 取消冻结 */
export function freezeAt(wb: Workbook, sheetName: string, coord: string | undefined): void {
  const ws = getSheet(wb, sheetName)
  if (!ws) return
  setFreezePanes(ws, coord as never)
}

/* ────────────── 排序（#33A：写回 records 本体，不用 ag-grid 内部排序） ────────────── */

/**
 * 按某列对整表行排序（1..maxRow，Excel「数据→排序」同语义）。
 *
 * 实现 = 「置换 + moveRange 循环搬运」：算出新行序（order[dst]=src），把置换分解成
 * 若干个循环，每个循环用一行尾部中转行串行搬完 —— 全程走 moveRange，
 * 值/公式/URN 引用/样式随行保真移动（与插删行同一条搬移链路）。
 *
 * 排序键分组（升序）：数字 < 文本（localeCompare zh）< 空格（永远在尾，Excel 同款）。
 * ⚠ 已知限制（与插删行同口径）：公式引用不自动偏移（=A2 不跟随行移动）。
 *
 * @returns 参与排序的行数（0 = 表不存在）
 */
export function sortRowsByCol(
  wb: Workbook,
  sheetName: string,
  keyCol: number,
  dir: 'asc' | 'desc',
): number {
  const ws = getSheet(wb, sheetName)
  if (!ws) return 0
  const { maxR, maxC } = extentMax(ws)
  if (maxR < 2) return maxR
  // ① 读排序键（公式格取缓存值，与网格显示同口径）
  const keys: unknown[] = []
  for (let r = 1; r <= maxR; r++) {
    const cell = getCellByCoord(ws, indexToCol(keyCol) + r)
    let v: unknown = ''
    if (cell && typeof cell === 'object' && 'value' in cell) {
      const raw = (cell as { value?: unknown }).value
      if (raw && typeof raw === 'object' && (raw as { kind?: string }).kind === 'formula') {
        v = (raw as { cachedValue?: unknown }).cachedValue ?? ''
      } else {
        v = raw ?? ''
      }
    }
    keys.push(v)
  }
  // ② 键分组权值：0=数字 1=文本 2=空（空永远最后，升降序一致）
  const rank = (v: unknown): number => {
    if (v === null || v === undefined || v === '') return 2
    if (typeof v === 'number' && Number.isFinite(v)) return 0
    if (typeof v === 'string' && /^-?\d+(\.\d+)?$/.test(v.trim())) return 0
    return 1
  }
  const numOf = (v: unknown): number => (typeof v === 'number' ? v : Number(String(v).trim()))
  const order: number[] = Array.from({ length: maxR }, (_, i) => i) // 0-based
  order.sort((a, b) => {
    const ra = rank(keys[a])
    const rb = rank(keys[b])
    if (ra !== rb) return ra - rb
    if (ra === 0) {
      const d = numOf(keys[a]) - numOf(keys[b])
      return dir === 'asc' ? d : -d
    }
    if (ra === 1) {
      const d = String(keys[a]).localeCompare(String(keys[b]), 'zh-Hans-CN')
      return dir === 'asc' ? d : -d
    }
    return a - b // 空组保持原行序稳定
  })
  if (order.every((v, i) => v === i)) return maxR // 已经有序，不动格子
  // ③ 尾部扩一行中转
  appendRows(ws, [Array.from({ length: maxC }, () => undefined)])
  const transit = maxR + 1
  // ④ 循环置换搬运：order[dst]=src —— dst 的内容来自 src，链条顺序搬，每循环一次中转
  const visited = new Array<boolean>(maxR).fill(false)
  const moveRow = (src1: number, dst1: number) =>
    moveRange(ws, rng(1, src1, maxC, src1), rng(1, dst1, maxC, dst1))
  for (let i = 0; i < maxR; i++) {
    if (visited[i] || order[i] === i) {
      visited[i] = true
      continue
    }
    moveRow(i + 1, transit) // 救出循环起点 i 的原始内容（它最先被覆盖）
    visited[i] = true
    let cur = i
    while (order[cur] !== i) {
      const src = order[cur]
      moveRow(src + 1, cur + 1) // dst=cur ← src=order[cur]
      visited[src] = true
      cur = src
    }
    moveRow(transit, cur + 1) // 收尾：起点内容流回链尾（cur=最后被腾空的行；落回起点 i 会覆盖刚搬好的数据）
  }
  deleteRow(wb, sheetName, transit)
  return maxR
}

/* ────────────── 粘贴扩展（#33A 选择性粘贴：转置 / 格式快照） ────────────── */

/** TSV 文本 → 二维网格（去尾空行；不转数字——转换发生在落格时） */
export function parseTSVGrid(text: string): string[][] {
  const lines = text.replace(/\r/g, '').split('\n')
  while (lines.length && lines[lines.length - 1] === '') lines.pop()
  return lines.map((l) => l.split('\t'))
}

/** 二维网格转置 */
export function transposeGrid<T>(grid: T[][]): T[][] {
  const nR = grid.length
  const nC = Math.max(...grid.map((g) => g.length), 0)
  const out: T[][] = []
  for (let c = 0; c < nC; c++) {
    const row: T[] = []
    for (let r = 0; r < nR; r++) row.push(grid[r][c])
    out.push(row)
  }
  return out
}

/**
 * 把值网格（string[][]）写入从 startCoord 开始的范围：纯数字串自动转数字（Excel 同款）。
 * pasteTSV 的网格入口版 —— 转置/选择性粘贴在 UI 侧先变换网格再落格。
 */
export function pasteGrid(
  wb: Workbook,
  sheetName: string,
  startCoord: string,
  grid: string[][],
): { rows: number; cols: number } {
  const ws = getSheet(wb, sheetName)
  if (!ws) throw new Error('工作表不存在')
  const start = parseCoord(startCoord)
  const nR = grid.length
  if (!nR) return { rows: 0, cols: 0 }
  const nC = Math.max(...grid.map((g) => g.length))
  const values: (number | string | null)[][] = []
  for (let r = 0; r < nR; r++) {
    const row: (number | string | null)[] = []
    for (let c = 0; c < nC; c++) {
      const raw = grid[r][c] ?? ''
      const t = raw.trim()
      row.push(t !== '' && /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : raw)
      ensureCellByCoord(ws, indexToCol(start.c + c) + (start.r + r))
    }
    values.push(row)
  }
  setRangeValues(ws, rng(start.c, start.r, start.c + nC - 1, start.r + nR - 1), values as never)
  return { rows: nR, cols: nC }
}

/** 格式快照矩阵（选择性粘贴「贴格式」的数据源）：每格四轴 = fmtOfFirst 同口径 */
export function fmtSnapshotOfRect(wb: Workbook, sheetName: string, rect: CellRect): FmtState[][] {
  const out: FmtState[][] = []
  for (let r = rect.r1; r <= rect.r2; r++) {
    const row: FmtState[] = []
    for (let c = rect.c1; c <= rect.c2; c++) {
      row.push(fmtOfFirst(wb, sheetName, { r1: r, c1: c, r2: r, c2: c }))
    }
    out.push(row)
  }
  return out
}

/** 把格式快照矩阵逐格写回（从 startCoord 开始，矩形对齐；1×1 写 = applyFmtSel 复用） */
export function applyFmtSnapshot(
  wb: Workbook,
  sheetName: string,
  startCoord: string,
  snap: FmtState[][],
): number {
  const start = parseCoord(startCoord)
  let n = 0
  for (let r = 0; r < snap.length; r++) {
    for (let c = 0; c < snap[r].length; c++) {
      applyFmtSel(wb, sheetName, { r1: start.r + r, c1: start.c + c, r2: start.r + r, c2: start.c + c }, snap[r][c])
      n++
    }
  }
  return n
}

/* ────────────── 粘贴 TSV ────────────── */

/**
 * 把剪贴板的 TSV 文本（从 WPS/Excel 复制的区域）写入从 startCoord 开始的范围。
 * 纯数字串自动转数字 —— 与 Excel 粘贴行为一致。
 */
export function pasteTSV(
  wb: Workbook,
  sheetName: string,
  startCoord: string,
  text: string,
): { rows: number; cols: number } {
  const ws = getSheet(wb, sheetName)
  if (!ws) throw new Error('工作表不存在')
  const start = parseCoord(startCoord)
  const lines = text.replace(/\r/g, '').split('\n')
  while (lines.length && lines[lines.length - 1] === '') lines.pop()
  if (!lines.length) return { rows: 0, cols: 0 }
  const grid = lines.map((l) => l.split('\t'))
  const nR = grid.length
  const nC = Math.max(...grid.map((g) => g.length))
  // 逐格转换 + ensure
  const values: (number | string | null)[][] = []
  for (let r = 0; r < nR; r++) {
    const row: (number | string | null)[] = []
    for (let c = 0; c < nC; c++) {
      const raw = grid[r][c] ?? ''
      const t = raw.trim()
      row.push(t !== '' && /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : raw)
      ensureCellByCoord(ws, indexToCol(start.c + c) + (start.r + r))
    }
    values.push(row)
  }
  setRangeValues(
    ws,
    rng(start.c, start.r, start.c + nC - 1, start.r + nR - 1),
    values as never
  )
  return { rows: nR, cols: nC }
}

/* ────────────── 工作表管理 ────────────── */

export function addSheetTitled(wb: Workbook, title: string): void {
  if (sheetNames(wb)?.includes(title)) throw new Error('已存在同名工作表：' + title)
  addWorksheet(wb, title)
}

export function renameSheetTitled(wb: Workbook, oldTitle: string, newTitle: string): void {
  if (oldTitle === newTitle) return
  if (sheetNames(wb)?.includes(newTitle)) throw new Error('已存在同名工作表：' + newTitle)
  renameSheet(wb, oldTitle, newTitle)
}

export function removeSheetTitled(wb: Workbook, title: string): void {
  const all = sheetNames(wb) ?? []
  if (all.length <= 1) throw new Error('至少要保留一个工作表')
  removeSheet(wb, title)
}

/** 供「插入公式」用的辅助：把 A1 坐标的值转回公式栏文本 */
export { setCellByCoord, getCellByCoord }
