import { writeFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { pickPoints, pointsToDxf, rowsToDxf, tableToDxf, trimToUsedRange, DXF_DEFAULT_OPTIONS } from './dxf'

/* ★ 2026-11-10 Excel转DXF：展点生成器 + 非空边界整表转换 测试 */

const ROWS = [
  { 点号: 'GPS01', X: 45234.567, Y: 31204.5, 高程: 32.1 },
  { 点号: 'GPS02', X: 45300, Y: 31300 },
  { 点号: '坏行', X: 'abc', Y: 31400 },
  { 点号: '', X: 45400, Y: 31500 },
]

describe('pickPoints — 行解析与容错', () => {
  it('数值行提取，坏行跳过并记原因，空点名自动补 P 序号', () => {
    const { points, skipped } = pickPoints(ROWS, { nameCol: '点号', xCol: 'X', yCol: 'Y', zCol: '高程' })
    expect(points).toHaveLength(3)
    expect(points[0]).toEqual({ name: 'GPS01', x: 45234.567, y: 31204.5, z: 32.1 })
    expect(points[2].name).toBe('P3')
    expect(skipped).toHaveLength(1)
    expect(skipped[0]).toContain('第 4 行')
    expect(skipped[0]).toContain('X 列（X）非数值')
  })
})

describe('pointsToDxf — R12 结构', () => {
  const pts = pickPoints(ROWS, { nameCol: '点号', xCol: 'X', yCol: 'Y' }).points

  it('包含 HEADER(R12 AC1009)/图层表/ENTITIES/EOF 段', () => {
    const dxf = pointsToDxf(pts, DXF_DEFAULT_OPTIONS)
    expect(dxf).toContain('AC1009')
    expect(dxf).toContain('CB_POINTS')
    expect(dxf).toContain('CB_TEXTS')
    expect(dxf).toContain('CB_LINES')
    expect(dxf.trimEnd().endsWith('EOF')).toBe(true)
    /* 组码行右对齐 3 字符（AutoCAD 严格要求） */
    expect(dxf.split('\r\n')[0]).toBe('  0')
  })

  it('POINT 与 TEXT 实体数量与点数一致；坐标保留 3 位内精度', () => {
    const dxf = pointsToDxf(pts, DXF_DEFAULT_OPTIONS)
    const cnt = (kw: string) => dxf.split('\r\n' + kw + '\r\n').length - 1
    expect(cnt('POINT')).toBe(3)
    expect(cnt('TEXT')).toBe(3)
    expect(dxf).toContain('45234.567')
    expect(dxf).toContain('GPS01')
  })

  it('flipY 取反 Y；关闭标注则无 TEXT 实体', () => {
    const dxf = pointsToDxf(pts, { ...DXF_DEFAULT_OPTIONS, flipY: true, drawLabel: false })
    expect(dxf).toContain('-31204.5')
    expect(dxf.split('\r\nTEXT\r\n').length - 1).toBe(0)
  })
})

describe('rowsToDxf — 组合入口', () => {
  it('返回 dxf/count/skipped 三元组', () => {
    const r = rowsToDxf(ROWS, { nameCol: '点号', xCol: 'X', yCol: 'Y' }, DXF_DEFAULT_OPTIONS)
    expect(r.count).toBe(3)
    expect(r.skipped).toHaveLength(1)
    expect(r.dxf.length).toBeGreaterThan(200)
  })
})

describe('非空边界整表转 DXF（用户澄清「边界确定是非空边界」）', () => {
  it('trimToUsedRange：过滤全空行/全空列', () => {
    const fields = ['A', 'B', 'C']
    const rows = [
      { A: 'a1', B: 'b1', C: '' },
      { A: '', B: '', C: '' },
      { A: 'a3', B: 'b3', C: '' },
    ]
    const { cols, rows: used } = trimToUsedRange(fields, rows)
    expect(cols).toEqual(['A', 'B'])
    expect(used).toHaveLength(2)
  })

  it('tableToDxf：网格线与表头+数据 TEXT 数量正确', () => {
    const fields = ['点号', 'X', 'Y']
    const rows = [
      { 点号: 'P1', X: '100', Y: '200' },
      { 点号: 'P2', X: '150', Y: '260' },
    ]
    const { dxf, usedRows, usedCols } = tableToDxf(fields, rows)
    expect(usedRows).toBe(2)
    expect(usedCols).toBe(3)
    /* 横线 4 + 纵线 4 = 8 次引用 + 图层表定义 1 = 9 */
    expect(dxf.split('CB_TABLE_LINES').length - 1).toBe(9)
    /* 表头 3 + 数据非空 6 = 9 次引用 + 图层表定义 1 = 10 */
    expect(dxf.split('CB_TABLE_TEXT').length - 1).toBe(10)
    /* 临时调试（用完即删）：写整个 dxf 看真实内容 */
    writeFileSync('.tmp-vt-debug.dxf', dxf)
    expect(dxf.split('\r\nTEXT\r\n').length - 1).toBe(9)
    expect(dxf).toContain('P1')
    expect(dxf).toContain('点号')
    expect(dxf.trimEnd().endsWith('EOF')).toBe(true)
  })

  it('空表 → 空实体但结构完整', () => {
    const { dxf, usedRows } = tableToDxf(['A'], [{ A: '' }])
    expect(usedRows).toBe(0)
    expect(dxf.trimEnd().endsWith('EOF')).toBe(true)
  })
})
