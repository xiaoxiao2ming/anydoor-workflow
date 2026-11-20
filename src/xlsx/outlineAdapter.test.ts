import { describe, it, expect } from 'vitest'
import { createWorkbook, addWorksheet } from '@office-kit/xlsx/workbook'
import { getSheet, type Workbook } from '@office-kit/xlsx/workbook'
import { setCellByCoord, getCellByCoord } from '@office-kit/xlsx/worksheet'
import type { Worksheet } from '@office-kit/xlsx/worksheet'

import {
  readOutline,
  writeOutline,
  commitToSheet,
  renumberOnly,
  isTitleRow,
  levelCellValue,
  noCellValue,
} from './outlineAdapter'
import { makeRow, makeRowsByLevel } from './outline'

/** 造一张只有 A/B/C 三列的小表，省去业务列的干扰 */
function sheet(name = 'S1'): { wb: Workbook; ws: Worksheet } {
  const wb = createWorkbook()
  const ws = addWorksheet(wb, name)
  return { wb, ws }
}

/** 读某格显示值（公式格取 cachedValue，与 readOutline 内部同口径） */
function cellStr(ws: Worksheet, coord: string): string {
  const c = getCellByCoord(ws, coord)
  const v = (c as { value?: unknown } | undefined)?.value
  if (v && typeof v === 'object' && (v as { kind?: string }).kind === 'formula') {
    return String((v as { cachedValue?: unknown }).cachedValue ?? '')
  }
  return v === undefined || v === null ? '' : String(v)
}

describe('readOutline —— 读表解析', () => {
  it('空表返回空数组，不抛', () => {
    const { ws } = sheet()
    const { rows, clamped } = readOutline(ws)
    expect(rows).toEqual([])
    expect(clamped).toBe(false)
  })

  it('null 工作表安全降级', () => {
    expect(readOutline(null).rows).toEqual([])
    expect(readOutline(undefined).rows).toEqual([])
  })

  it('A 列数字 → level，B 列脏值不认（只以 A 列为准）', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', 1)
    setCellByCoord(ws, 'B1', '999') // 脏编号，应被重算覆盖
    setCellByCoord(ws, 'C1', '总则')
    setCellByCoord(ws, 'A2', 2)
    setCellByCoord(ws, 'C2', '编制依据')
    const { rows } = readOutline(ws)
    expect(rows[0].level).toBe(1)
    expect(rows[0].title).toBe('总则')
    expect(rows[0].no).toBe('1') // 忽略 B 列写的 999
    expect(rows[1].level).toBe(2)
    expect(rows[1].no).toBe('1.1')
  })

  it('A 列 # 字符模式也认（方案 §6 兼容口径）', () => {
    const { ws } = sheet()
    // 逐级下降才合法：1 → 2 → 3。中间跳级会被 normalizeLevels 钳掉（那是另一个用例）
    setCellByCoord(ws, 'A1', '#')
    setCellByCoord(ws, 'C1', '一级')
    setCellByCoord(ws, 'A2', '##')
    setCellByCoord(ws, 'C2', '二级')
    setCellByCoord(ws, 'A3', '###')
    setCellByCoord(ws, 'C3', '三级')
    const { rows, clamped } = readOutline(ws)
    expect(clamped).toBe(false)
    expect(rows.map((r) => r.level)).toEqual([1, 2, 3])
    expect(rows.map((r) => r.no)).toEqual(['1', '1.1', '1.1.1'])
  })

  it('首行不能是深层（normalizeLevels 钳制，outline.ts 既有铁律）', () => {
    // 单行 '###' 无父级可挂 → 钳成 1，这是正确行为，不是 bug
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', '###')
    const { rows, clamped } = readOutline(ws)
    expect(clamped).toBe(true)
    expect(rows[0].level).toBe(1)
    expect(rows[0].no).toBe('1')
  })

  it('A 列垃圾值归 0（普通数据行），不抛', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', 'abc')
    setCellByCoord(ws, 'C1', '不是标题')
    const { rows } = readOutline(ws)
    expect(rows[0].level).toBe(0)
    expect(rows[0].no).toBe('')
  })

  it('跳级（1,3）被钳制为 (1,2) 并回报 clamped', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', 1)
    setCellByCoord(ws, 'A2', 3) // 跳级
    const { rows, clamped } = readOutline(ws)
    expect(clamped).toBe(true)
    expect(rows[1].level).toBe(2)
    expect(rows[1].no).toBe('1.1')
  })

  it('普通行不打断编号：1 / 0 / 1 → 1 / 1.1', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', 1)
    setCellByCoord(ws, 'A2', 0)
    setCellByCoord(ws, 'A3', 1)
    const { rows } = readOutline(ws)
    expect(rows.map((r) => r.no)).toEqual(['1', '', '2'])
  })

  it('业务列（D 起）进 cells，且不含 A/B/C', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', 1)
    setCellByCoord(ws, 'C1', '管段')
    setCellByCoord(ws, 'D1', 'DN100')
    setCellByCoord(ws, 'E1', 0.3)
    const { rows } = readOutline(ws)
    expect(rows[0].cells).toEqual({ D: 'DN100', E: 0.3 })
    expect(rows[0].cells.A).toBeUndefined()
    expect(rows[0].cells.B).toBeUndefined()
    expect(rows[0].cells.C).toBeUndefined()
  })

  it('业务列空值不进 cells（不产生一堆undefined 键）', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', 1)
    setCellByCoord(ws, 'D1', '')
    const { rows } = readOutline(ws)
    expect(Object.keys(rows[0].cells)).toEqual([])
  })

  it('自定义列布局生效', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', '数据') // A 变成业务列
    setCellByCoord(ws, 'B1', 1) // B 变标记列
    setCellByCoord(ws, 'D1', '自定义标题')
    const { rows } = readOutline(ws, { level: 'B', no: 'C', title: 'D' })
    expect(rows[0].level).toBe(1)
    expect(rows[0].title).toBe('自定义标题')
    expect(rows[0].no).toBe('1')
  })

  it('非法列字母回落默认，不抛', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', 1)
    const { rows } = readOutline(ws, { level: '??', no: '', title: '1' })
    expect(rows[0].level).toBe(1) // 回落 A 列
  })

  it('公式格取 cachedValue 当level（否则公式标题会凭空消失）', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', 1)
    // 手工塞一个公式形态的值
    const c = getCellByCoord(ws, 'A1') as { value?: unknown }
    c.value = { kind: 'formula', formula: '1+0', cachedValue: 1 }
    const { rows } = readOutline(ws)
    expect(rows[0].level).toBe(1)
  })
})

describe('writeOutline —— 写回', () => {
  it('level>0 写A/B/C 三列', () => {
    const { ws } = sheet()
    const rows = makeRowsByLevel([1, 2], ['总则', '依据'])
    writeOutline(ws, rows)
    expect(cellStr(ws, 'A1')).toBe('1')
    expect(cellStr(ws, 'B1')).toBe('1')
    expect(cellStr(ws, 'C1')).toBe('总则')
    expect(cellStr(ws, 'A2')).toBe('2')
    expect(cellStr(ws, 'B2')).toBe('1.1')
    expect(cellStr(ws, 'C2')).toBe('依据')
  })

  it('level=0 清空三列（降回普通行不留残影）', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', 2)
    setCellByCoord(ws, 'B1', '1.1')
    setCellByCoord(ws, 'C1', '旧标题')
    const rows = [makeRow({ level: 0 })]
    writeOutline(ws, rows)
    expect(cellStr(ws, 'A1')).toBe('')
    expect(cellStr(ws, 'B1')).toBe('')
    expect(cellStr(ws, 'C1')).toBe('')
  })

  it('不碰业务列（章节编排无权改用户数据）', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'D1', 'DN100')
    setCellByCoord(ws, 'E1', 0.3)
    writeOutline(ws, makeRowsByLevel([1], ['标题']))
    expect(cellStr(ws, 'D1')).toBe('DN100')
    expect(cellStr(ws, 'E1')).toBe('0.3')
  })

  it('null 工作表安全降级', () => {
    expect(() => writeOutline(null, makeRowsByLevel([1]))).not.toThrow()
  })

  it('row.no 是脏值/空串时，写盘仍写正确编号（现算，不信入参）', () => {
    // 这是本轮实测修掉的缺陷：原先直接写 row.no，
    // 而 makeRowsByLevel 造的测试数据没跑过 renumber → 写出去的 B 列全是空
    const { ws } = sheet()
    const rows = makeRowsByLevel([1, 2, 2], ['一', '1-1', '1-2'])
    expect(rows.map((r) => r.no)).toEqual(['', '', '']) // 前提：入参确实是空的
    writeOutline(ws, rows)
    expect(cellStr(ws, 'B1')).toBe('1')
    expect(cellStr(ws, 'B2')).toBe('1.1')
    expect(cellStr(ws, 'B3')).toBe('1.2')
  })

  it('row.no 是过期值时，写盘用重算结果覆盖', () => {
    const { ws } = sheet()
    const rows = makeRowsByLevel([1, 2], ['一', '1-1'])
    rows[0].no = '99' // 改层级前的旧编号
    rows[1].no = '99.1'
    writeOutline(ws, rows)
    expect(cellStr(ws, 'B1')).toBe('1')
    expect(cellStr(ws, 'B2')).toBe('1.1')
  })

  it('写盘不改入参（rows 保持 readonly 语义）', () => {
    const { ws } = sheet()
    const rows = makeRowsByLevel([1, 2], ['一', '1-1'])
    writeOutline(ws, rows)
    expect(rows.map((r) => r.no)).toEqual(['', '']) // 未被写盘副作用污染
  })
})

describe('往返一致性', () => {
  it('写 → 读 → 写 幂等（第二次写完单元格内容完全不变）', () => {
    const { ws } = sheet()
    const rows = makeRowsByLevel([1, 2, 2, 3, 0, 1], ['一', '1-1', '1-2', '1-2-1', '数据', '二'])
    writeOutline(ws, rows)
    const snap1 = ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A6', 'B6', 'C6'].map((c) => cellStr(ws, c))
    const back = readOutline(ws)
    writeOutline(ws, back.rows)
    const snap2 = ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A6', 'B6', 'C6'].map((c) => cellStr(ws, c))
    expect(snap2).toEqual(snap1)
  })

  it('往返后层级与编号完全保持', () => {
    const { ws } = sheet()
    const rows = makeRowsByLevel([1, 2, 2, 3], ['一', '1-1', '1-2', '1-2-1'])
    writeOutline(ws, rows)
    const back = readOutline(ws)
    expect(back.rows.map((r) => r.level)).toEqual([1, 2, 2, 3])
    expect(back.rows.map((r) => r.no)).toEqual(['1', '1.1', '1.2', '1.2.1'])
    expect(back.rows.map((r) => r.title)).toEqual(['一', '1-1', '1-2', '1-2-1'])
  })

  it('B 列被外部改脏后，重读能纠正回来', () => {
    const { ws } = sheet()
    writeOutline(ws, makeRowsByLevel([1, 2], ['一', '1-1']))
    setCellByCoord(ws, 'B2', '99.99') // 模拟用户手改编号
    const back = readOutline(ws)
    expect(back.rows[1].no).toBe('1.1')
  })
})

describe('commitToSheet —— 结构变更收口', () => {
  it('钳制 + 重算 + 写盘三步一次到位', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', 1)
    setCellByCoord(ws, 'A2', 4) // 跳级
    // 绕开 readOutline（它自己会钳制），直接给未规范化的行，验证 commitToSheet 自己也能钳
    const rows = [makeRow({ level: 1 }), makeRow({ level: 4, title: '跳级' })]
    const res = commitToSheet(ws, rows)
    expect(res.clamped).toBe(true)
    expect(res.written).toBe(2)
    expect(cellStr(ws, 'A2')).toBe('2')
    expect(cellStr(ws, 'B2')).toBe('1.1')
  })

  it('幂等：已规范化的行再commit 一次 clamped=false（不重复打扰用户）', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', 1)
    setCellByCoord(ws, 'A2', 2)
    const { rows } = readOutline(ws) // 读的时候已钳制过
    expect(commitToSheet(ws, rows).clamped).toBe(false)
  })

  it('结构变更后写盘的行数与行数组一致', () => {
    const { ws } = sheet()
    const rows = makeRowsByLevel([1, 2, 2], ['一', '1-1', '1-2'])
    expect(commitToSheet(ws, rows).written).toBe(3)
  })

  it('null 工作表返回 0 且不抛', () => {
    const res = commitToSheet(null, makeRowsByLevel([1]))
    expect(res).toEqual({ written: 0, clamped: false })
  })
})

describe('轻量工具', () => {
  it('renumberOnly 只重算内存里的编号、不写盘', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', 1) // 表里是 1
    const { rows } = readOutline(ws)
    rows[0].no = '错误值'
    renumberOnly(rows)
    expect(rows[0].no).toBe('1') // 内存里被纠正
    expect(cellStr(ws, 'A1')).toBe('1') // 表里本来就是 1，未被改动
  })

  it('renumberOnly 绝不碰工作表（折叠展开不该触发全表写盘）', () => {
    const { ws } = sheet()
    setCellByCoord(ws, 'A1', 2) // 故意写个与重算结果不同的值
    setCellByCoord(ws, 'A2', 3)
    const { rows } = readOutline(ws)
    renumberOnly(rows)
    // 表里的值应保持原样，不被 renumberOnly 改写
    expect(cellStr(ws, 'A1')).toBe('2')
    expect(cellStr(ws, 'A2')).toBe('3')
  })

  it('isTitleRow 只认 level>0', () => {
    expect(isTitleRow(makeRow({ level: 1 }))).toBe(true)
    expect(isTitleRow(makeRow({ level: 0 }))).toBe(false)
    expect(isTitleRow(undefined)).toBe(false)
  })

  it('levelCellValue / noCellValue 与 writeOutline 同口径', () => {
    const t = makeRow({ level: 2, no: '1.1' })
    const b = makeRow({ level: 0 })
    expect(levelCellValue(t)).toBe('2')
    expect(noCellValue(t)).toBe('1.1')
    expect(levelCellValue(b)).toBe('')
    expect(noCellValue(b)).toBe('')
  })
})

describe('规模', () => {
  it('1000 行读写都在 20ms 内', () => {
    const { ws } = sheet()
    const levels = Array.from({ length: 1000 }, (_, i) => (i % 7 === 0 ? 1 : i % 3 === 0 ? 2 : 0))
    const rows = makeRowsByLevel(levels)
    const t0 = performance.now()
    writeOutline(ws, rows)
    const { rows: back } = readOutline(ws)
    const cost = performance.now() - t0
    expect(back.length).toBe(1000)
    expect(cost).toBeLessThan(20)
  })
})
