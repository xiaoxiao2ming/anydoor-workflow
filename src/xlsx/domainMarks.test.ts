import { describe, it, expect } from 'vitest'
import { buildMarks, inAnyTable, displayRef } from './domainMarks'
import type { XlsxFieldSpec } from './field'

const F = (o: Partial<XlsxFieldSpec> & { key: string; kind: XlsxFieldSpec['kind'] }): XlsxFieldSpec => ({
  ref: '',
  ...o,
})

describe('displayRef（域引用显示）', () => {
  it('命名区域的 ref 已带表名 → 不再补前缀（回归：曾显示成 汇总表!汇总表!$B$1）', () => {
    expect(displayRef('汇总表', '汇总表!$B$1')).toBe('汇总表!$B$1')
  })

  it('Table 的 ref 只有范围 → 补上表名', () => {
    expect(displayRef('工程量表', 'A1:D2')).toBe('工程量表!A1:D2')
  })

  it('没有表名时原样返回', () => {
    expect(displayRef(undefined, 'A1:D2')).toBe('A1:D2')
  })
})

describe('buildMarks（域位置索引）', () => {
  const fields: XlsxFieldSpec[] = [
    F({ key: '工程名称', kind: 'definedName', ref: '汇总表!$B$1', sheet: '汇总表' }),
    F({ key: '设计流量', kind: 'definedName', ref: '汇总表!$B$2', sheet: '汇总表' }),
    F({ key: '材料表', kind: 'table', ref: 'A1:D2', sheet: '工程量表' }),
  ]

  it('命名区域落到具体坐标', () => {
    const m = buildMarks(fields)
    const s = m.get('汇总表')!
    expect(s.named.get('B1')).toBe('工程名称')
    expect(s.named.get('B2')).toBe('设计流量')
  })

  it('Table 落到范围', () => {
    const m = buildMarks(fields)
    const s = m.get('工程量表')!
    expect(s.tables).toHaveLength(1)
    expect(s.tables[0]).toMatchObject({ key: '材料表', minRow: 1, maxRow: 2, minCol: 1, maxCol: 4 })
  })

  it('inAnyTable 判定范围内外', () => {
    const m = buildMarks(fields)
    const s = m.get('工程量表')
    expect(inAnyTable(s, 1, 1)).toBe('材料表')
    expect(inAnyTable(s, 2, 4)).toBe('材料表')
    expect(inAnyTable(s, 3, 1)).toBe(null) // 超出范围
    expect(inAnyTable(s, 1, 5)).toBe(null) // 超出范围
    expect(inAnyTable(undefined, 1, 1)).toBe(null)
  })

  it('认不出来的 ref 直接跳过，不猜位置', () => {
    const m = buildMarks([
      F({ key: '坏的', kind: 'definedName', ref: '这不是坐标', sheet: '汇总表' }),
      F({ key: '没表名', kind: 'definedName', ref: 'B1', sheet: undefined }),
      F({ key: '坏表', kind: 'table', ref: '乱七八糟', sheet: '工程量表' }),
    ])
    expect(m.get('汇总表')?.named.size ?? 0).toBe(0)
    expect(m.get('工程量表')?.tables.length ?? 0).toBe(0)
  })

  it('带 $ 的 ref 也认', () => {
    const m = buildMarks([F({ key: '长', kind: 'definedName', ref: 'Sheet1!$AA$123', sheet: 'Sheet1' })])
    expect(m.get('Sheet1')!.named.get('AA123')).toBe('长')
  })
})
