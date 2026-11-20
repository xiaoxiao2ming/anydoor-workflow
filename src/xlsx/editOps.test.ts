import { describe, it, expect } from 'vitest'
import { createWorkbook, addWorksheet, getSheet } from '@office-kit/xlsx/workbook'
import { setCellFont, getCellFont, setCellBackgroundColor, cellStyleToCss, getCellNumberFormat } from '@office-kit/xlsx/styles'
import { getCellByCoord, setCellByCoord, ensureCellByCoord } from '@office-kit/xlsx/worksheet'
import { fontSel, insertCellShift, fontOfFirst, fmtOfFirst, applyFmtSel, sortRowsByCol, parseTSVGrid, transposeGrid, pasteGrid, fmtSnapshotOfRect, applyFmtSnapshot, type CellRect } from './editOps'

/**
 * #33 右键菜单 v2 底层动作：
 * 1. fontSel 合并写回——v1 用 setCellStyle 整只替换 font spec，只传 bold 会把字号/字体名重置回默认
 *   （实测缺陷）；改 patchCellFont 后只动显式传入的键。
 * 2. insertCellShift——选区腾空、原内容按方向推移、尾部溢出丢弃（Excel 同款语义）。
 */

function mkWb() {
  const wb = createWorkbook()
  addWorksheet(wb, 'Sheet1')
  return wb
}

const set = (wb: ReturnType<typeof createWorkbook>, coord: string, v: unknown) => {
  const ws = getSheet(wb, 'Sheet1')!
  setCellByCoord(ws, coord, v as never)
}
const get = (wb: ReturnType<typeof createWorkbook>, coord: string): unknown => {
  const ws = getSheet(wb, 'Sheet1')!
  const c = getCellByCoord(ws, coord)
  return c && typeof c === 'object' && 'value' in c ? (c as { value?: unknown }).value : undefined
}

describe('fontSel 合并写回（#33：只动显式传入的键）', () => {
  it('预设字号 14 后加粗——字号保持不被重置（v1 缺陷回归）', () => {
    const wb = mkWb()
    const ws = getSheet(wb, 'Sheet1')!
    const cell = ensureCellByCoord(ws, 'B2')!
    setCellFont(wb, cell, { size: 14, name: '宋体' })
    const rect: CellRect = { r1: 2, c1: 2, r2: 2, c2: 2 }
    fontSel(wb, 'Sheet1', rect, { bold: true })
    const f = getCellFont(wb, cell)
    expect(f.bold).toBe(true)
    expect(f.size).toBe(14) // v1 整只替换会把 size 抹掉
    expect(f.name).toBe('宋体')
  })

  it('只传 underline 时 bold/字号都不动', () => {
    const wb = mkWb()
    const ws = getSheet(wb, 'Sheet1')!
    const cell = ensureCellByCoord(ws, 'C3')!
    setCellFont(wb, cell, { bold: true, size: 12 })
    fontSel(wb, 'Sheet1', { r1: 3, c1: 3, r2: 3, c2: 3 }, { underline: true })
    const f = getCellFont(wb, cell)
    expect(f.underline).toBe('single')
    expect(f.bold).toBe(true)
    expect(f.size).toBe(12)
  })

  it('取消加粗（bold:false 显式写回，不误伤其他属性）', () => {
    const wb = mkWb()
    const ws = getSheet(wb, 'Sheet1')!
    const cell = ensureCellByCoord(ws, 'D4')!
    setCellFont(wb, cell, { bold: true, size: 16 })
    fontSel(wb, 'Sheet1', { r1: 4, c1: 4, r2: 4, c2: 4 }, { bold: false })
    const f = getCellFont(wb, cell)
    expect(f.bold).toBe(false)
    expect(f.size).toBe(16)
  })

  it('fontOfFirst 读首格字体', () => {
    const wb = mkWb()
    const ws = getSheet(wb, 'Sheet1')!
    const cell = ensureCellByCoord(ws, 'A1')!
    setCellFont(wb, cell, { bold: true })
    const f = fontOfFirst(wb, 'Sheet1', { r1: 1, c1: 1, r2: 2, c2: 2 })
    expect(f?.bold).toBe(true)
  })
})

describe('insertCellShift（#33 插入单元格：腾空+推移+尾部丢弃）', () => {
  it('down：A 列 1/2/3 选 A2 插入 → A2 空、2/3 下移一行', () => {
    const wb = mkWb()
    set(wb, 'A1', 1)
    set(wb, 'A2', 2)
    set(wb, 'A3', 3)
    insertCellShift(wb, 'Sheet1', { r1: 2, c1: 1, r2: 2, c2: 1 }, 'down')
    expect(get(wb, 'A1')).toBe(1)
    expect(get(wb, 'A2')).toBe(undefined) // 腾空
    expect(get(wb, 'A3')).toBe(2)
    expect(get(wb, 'A4')).toBe(3)
  })

  it('right：行 1/2/3 选 A1 插入 → A1 空、1/2 右移一列', () => {
    const wb = mkWb()
    set(wb, 'A1', 1)
    set(wb, 'B1', 2)
    set(wb, 'C1', 3)
    insertCellShift(wb, 'Sheet1', { r1: 1, c1: 1, r2: 1, c2: 1 }, 'right')
    expect(get(wb, 'A1')).toBe(undefined) // 腾空
    expect(get(wb, 'B1')).toBe(1)
    expect(get(wb, 'C1')).toBe(2)
    expect(get(wb, 'D1')).toBe(3)
  })

  it('2×2 选区 down：整块推移两行，仅选区覆盖的列参与', () => {
    const wb = mkWb()
    set(wb, 'A1', 'a1')
    set(wb, 'B1', 'b1')
    set(wb, 'A2', 'a2')
    set(wb, 'B2', 'b2')
    set(wb, 'A3', 'a3')
    set(wb, 'C1', 'c1') // 选区外不受影响
    insertCellShift(wb, 'Sheet1', { r1: 1, c1: 1, r2: 2, c2: 2 }, 'down')
    expect(get(wb, 'A1')).toBe(undefined)
    expect(get(wb, 'B2')).toBe(undefined)
    expect(get(wb, 'A3')).toBe('a1')
    expect(get(wb, 'B4')).toBe('b2')
    expect(get(wb, 'C1')).toBe('c1') // c 列不在选区列范围，纹丝不动
  })

  it('样式随内容一起推移（down 推移后底色跟走）', () => {
    const wb = mkWb()
    const ws = getSheet(wb, 'Sheet1')!
    set(wb, 'A1', 7)
    const c1 = ensureCellByCoord(ws, 'A1')!
    setCellBackgroundColor(wb, c1, { rgb: 'FFFF00' })
    // 先自证底色写上了（防 API 静默丢导致假断言）
    expect((cellStyleToCss(wb, c1)['background-color'] ?? '').toLowerCase()).toContain('ffff00')
    insertCellShift(wb, 'Sheet1', { r1: 1, c1: 1, r2: 1, c2: 1 }, 'down')
    const moved = ensureCellByCoord(ws, 'A2')!
    expect(get(wb, 'A2')).toBe(7)
    // moveRange 搬值也搬样式（editOps 头注释口径）
    expect((cellStyleToCss(wb, moved)['background-color'] ?? '').toLowerCase()).toContain('ffff00')
  })
})

describe('设置单元格格式对话框读写（#33 第二块：fmtOfFirst / applyFmtSel）', () => {
  it('空格/表外读初值 = 默认态（对话框始终可开）', () => {
    const wb = mkWb()
    const init = fmtOfFirst(wb, 'Sheet1', { r1: 5, c1: 3, r2: 5, c2: 3 })
    expect(init.font.size).toBe(11)
    expect(init.font.bold).toBe(false)
    expect(init.font.color).toBeNull()
    expect(init.fill).toBeNull()
    expect(init.border.top.style).toBe('none')
    expect(init.numFmt).toBe('General')
  })

  it('四轴一次写入 → fmtOfFirst 读回一致', () => {
    const wb = mkWb()
    set(wb, 'B2', 3.14)
    const st = {
      font: { name: '宋体', size: 14, bold: true, italic: true, underline: true, strike: false, color: '#FF0000' },
      fill: '#FFFF00',
      border: {
        top: { style: 'thin' as const, color: '#00B050' },
        bottom: { style: 'none' as const, color: null },
        left: { style: 'none' as const, color: null },
        right: { style: 'none' as const, color: null },
      },
      numFmt: '#,##0.00',
    }
    applyFmtSel(wb, 'Sheet1', { r1: 2, c1: 2, r2: 3, c2: 2 }, st) // B2:B3 两格
    const back = fmtOfFirst(wb, 'Sheet1', { r1: 2, c1: 2, r2: 2, c2: 2 })
    expect(back.font.name).toBe('宋体')
    expect(back.font.size).toBe(14)
    expect(back.font.bold).toBe(true)
    expect(back.font.italic).toBe(true)
    expect(back.font.underline).toBe(true)
    expect(back.font.color).toBe('#FF0000')
    expect(back.fill).toBe('#FFFF00')
    expect(back.border.top.style).toBe('thin')
    expect(back.border.top.color).toBe('#00B050')
    expect(back.numFmt).toBe('#,##0.00')
    // 选区第二格（B3 空格）也拿到格式 —— setRangeStyle ensureCell 建格语义
    const back3 = fmtOfFirst(wb, 'Sheet1', { r1: 3, c1: 2, r2: 3, c2: 2 })
    expect(back3.fill).toBe('#FFFF00')
    expect(back3.numFmt).toBe('#,##0.00')
  })

  it('fill:null 写回 = 清填充；数字格式真落在 cell 上（显示链路数据源）', () => {
    const wb = mkWb()
    const ws = getSheet(wb, 'Sheet1')!
    const cell = ensureCellByCoord(ws, 'A1')!
    // 先自证有填充，再清
    applyFmtSel(wb, 'Sheet1', { r1: 1, c1: 1, r2: 1, c2: 1 }, {
      font: { name: '', size: 11, bold: false, italic: false, underline: false, strike: false, color: null },
      fill: '#FFC000',
      border: { top: { style: 'none', color: null }, bottom: { style: 'none', color: null }, left: { style: 'none', color: null }, right: { style: 'none', color: null } },
      numFmt: 'General',
    })
    expect(fmtOfFirst(wb, 'Sheet1', { r1: 1, c1: 1, r2: 1, c2: 1 }).fill).toBe('#FFC000')
    applyFmtSel(wb, 'Sheet1', { r1: 1, c1: 1, r2: 1, c2: 1 }, {
      font: { name: '', size: 11, bold: false, italic: false, underline: false, strike: false, color: null },
      fill: null,
      border: { top: { style: 'none', color: null }, bottom: { style: 'none', color: null }, left: { style: 'none', color: null }, right: { style: 'none', color: null } },
      numFmt: '0.00%',
    })
    expect(fmtOfFirst(wb, 'Sheet1', { r1: 1, c1: 1, r2: 1, c2: 1 }).fill).toBeNull()
    expect(getCellNumberFormat(wb, cell)).toBe('0.00%')
  })

  it('cellStyleToCss 反映字体加粗与填充（网格 cellStyle 渲染的数据源）', () => {
    const wb = mkWb()
    applyFmtSel(wb, 'Sheet1', { r1: 1, c1: 1, r2: 1, c2: 1 }, {
      font: { name: '', size: 12, bold: true, italic: false, underline: false, strike: false, color: '#FF0000' },
      fill: '#0070C0',
      border: { top: { style: 'none', color: null }, bottom: { style: 'none', color: null }, left: { style: 'none', color: null }, right: { style: 'none', color: null } },
      numFmt: 'General',
    })
    const ws = getSheet(wb, 'Sheet1')!
    const css = cellStyleToCss(wb, ensureCellByCoord(ws, 'A1')!)
    expect(css['font-weight']).toBe('bold')
    expect((css['background-color'] ?? '').toLowerCase()).toContain('0070c0')
  })
})

describe('sortRowsByCol（#33A 排序写回本体：moveRange 循环置换）', () => {
  it('升序：整行数据按键列重排（值跟着行走）', () => {
    const wb = mkWb()
    // B 列做键：3/1/2；A 列是同行数据：三十/一十/二十
    set(wb, 'B1', 3)
    set(wb, 'A1', '三十')
    set(wb, 'B2', 1)
    set(wb, 'A2', '一十')
    set(wb, 'B3', 2)
    set(wb, 'A3', '二十')
    sortRowsByCol(wb, 'Sheet1', 2, 'asc') // 按 B 列升序
    expect(get(wb, 'B1')).toBe(1)
    expect(get(wb, 'A1')).toBe('一十')
    expect(get(wb, 'B2')).toBe(2)
    expect(get(wb, 'A2')).toBe('二十')
    expect(get(wb, 'B3')).toBe(3)
    expect(get(wb, 'A3')).toBe('三十')
  })

  it('降序 = 升序的镜像', () => {
    const wb = mkWb()
    set(wb, 'B1', 3)
    set(wb, 'B2', 1)
    set(wb, 'B3', 2)
    sortRowsByCol(wb, 'Sheet1', 2, 'desc')
    expect(get(wb, 'B1')).toBe(3)
    expect(get(wb, 'B2')).toBe(2)
    expect(get(wb, 'B3')).toBe(1)
  })

  it('空值排尾（键列 B）', () => {
    const wb = mkWb()
    set(wb, 'B2', 5)
    set(wb, 'B3', 1)
    // B1 留空
    sortRowsByCol(wb, 'Sheet1', 2, 'asc')
    expect(get(wb, 'B1')).toBe(1)
    expect(get(wb, 'B2')).toBe(5)
    expect(String(get(wb, 'B3') ?? '')).toBe('') // 空格沉底
  })

  it('数字与文本分组（数字组在前、文本组在后）', () => {
    const wb = mkWb()
    set(wb, 'B1', '钢管')
    set(wb, 'B2', 10)
    set(wb, 'B3', 2)
    sortRowsByCol(wb, 'Sheet1', 2, 'asc')
    expect(get(wb, 'B1')).toBe(2)
    expect(get(wb, 'B2')).toBe(10)
    expect(get(wb, 'B3')).toBe('钢管')
  })

  it('已有序时不搬格子（幂等）', () => {
    const wb = mkWb()
    set(wb, 'B1', 1)
    set(wb, 'B2', 2)
    sortRowsByCol(wb, 'Sheet1', 2, 'asc')
    expect(get(wb, 'B1')).toBe(1)
    expect(get(wb, 'B2')).toBe(2)
  })

  it('样式随行搬移（填充色跟数据走）', () => {
    const wb = mkWb()
    set(wb, 'B1', 3)
    set(wb, 'B2', 1)
    // B1 上黄色填充 → 排序后应跟 3 一起落到第 2 行
    applyFmtSel(wb, 'Sheet1', { r1: 1, c1: 2, r2: 1, c2: 2 }, {
      font: { name: '', size: 11, bold: false, italic: false, underline: false, strike: false, color: null },
      fill: '#FFFF00',
      border: { top: { style: 'none', color: null }, bottom: { style: 'none', color: null }, left: { style: 'none', color: null }, right: { style: 'none', color: null } },
      numFmt: 'General',
    })
    sortRowsByCol(wb, 'Sheet1', 2, 'asc')
    expect(get(wb, 'B1')).toBe(1)
    expect(get(wb, 'B2')).toBe(3)
    expect(fmtOfFirst(wb, 'Sheet1', { r1: 2, c1: 2, r2: 2, c2: 2 }).fill).toBe('#FFFF00')
  })
})

describe('粘贴扩展（#33A 选择性粘贴底座）', () => {
  it('parseTSVGrid + transposeGrid + pasteGrid：转置落格', () => {
    const wb = mkWb()
    const grid = parseTSVGrid('1\t2\n3\t4')
    expect(grid).toEqual([['1', '2'], ['3', '4']])
    pasteGrid(wb, 'Sheet1', 'A1', transposeGrid(grid))
    expect(get(wb, 'A1')).toBe(1)
    expect(get(wb, 'B1')).toBe(3)
    expect(get(wb, 'A2')).toBe(2)
    expect(get(wb, 'B2')).toBe(4)
  })

  it('pasteGrid 纯数字串自动转数字', () => {
    const wb = mkWb()
    pasteGrid(wb, 'Sheet1', 'A1', [['10', 'x2']])
    expect(get(wb, 'A1')).toBe(10)
    expect(get(wb, 'B1')).toBe('x2')
  })

  it('fmtSnapshotOfRect + applyFmtSnapshot：格式快照往返', () => {
    const wb = mkWb()
    set(wb, 'A1', 7)
    applyFmtSel(wb, 'Sheet1', { r1: 1, c1: 1, r2: 1, c2: 1 }, {
      font: { name: '', size: 11, bold: false, italic: false, underline: false, strike: false, color: null },
      fill: '#00FF00',
      border: { top: { style: 'none', color: null }, bottom: { style: 'none', color: null }, left: { style: 'none', color: null }, right: { style: 'none', color: null } },
      numFmt: '0.00%',
    })
    const snap = fmtSnapshotOfRect(wb, 'Sheet1', { r1: 1, c1: 1, r2: 1, c2: 1 })
    applyFmtSnapshot(wb, 'Sheet1', 'C3', snap) // 贴到 C3
    expect(fmtOfFirst(wb, 'Sheet1', { r1: 3, c1: 3, r2: 3, c2: 3 }).fill).toBe('#00FF00')
    expect(fmtOfFirst(wb, 'Sheet1', { r1: 3, c1: 3, r2: 3, c2: 3 }).numFmt).toBe('0.00%')
  })
})
