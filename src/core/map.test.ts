// 数据对照表（map 节点）计算语义测试
import { describe, it, expect } from 'vitest'
import { applyMapping, computeNode } from './engine'
import { defineField, defineNode, type TableSchema } from './schema'
import type { Row } from '../store'

// 上游：用量清单，只有材料编码
const bill = defineNode('table', '用量清单', [defineField('编码', 'text'), defineField('数量', 'number')], {}, 'bill')
const billRows: Row[] = [
  { 编码: 'A001', 数量: 10 },
  { 编码: 'B007', 数量: 5 },
  { 编码: 'Z999', 数量: 1 }, // 对照表里查不到
]

// 对照表：编码 → 单价 / 单位
const price = defineNode(
  'map',
  '单价对照表',
  [defineField('编码', 'text'), defineField('单价', 'number'), defineField('单位', 'text')],
  { mapKeyField: '编码', mapMatchField: '编码', mapReturnFields: ['单价', '单位'], mapDefault: 'N/A' },
  'price',
)
const priceRows: Row[] = [
  { 编码: 'A001', 单价: 12.5, 单位: 'm' },
  { 编码: 'B007', 单价: 8, 单位: 'kg' },
]

const tables: Record<string, TableSchema> = { bill, price }
const records: Record<string, Row[]> = { bill: billRows, price: priceRows }

describe('数据对照表（map）', () => {
  it('命中补列、未命中填默认值', () => {
    const r = applyMapping(price, priceRows, bill.fields, billRows)
    // 字段：上游 2 列 + 取回 2 列
    expect(r.fields.map((f) => f.id)).toEqual(['编码', '数量', 'map_单价', 'map_单位'])
    expect(r.rows[0]['map_单价']).toBe(12.5)
    expect(r.rows[0]['map_单位']).toBe('m')
    expect(r.rows[1]['map_单价']).toBe(8)
    // 未命中 → 默认值
    expect(r.rows[2]['map_单价']).toBe('N/A')
  })

  it('列名为空时不冲突：重名自动加「（对照）」后缀', () => {
    const r = applyMapping(price, priceRows, bill.fields, billRows)
    expect(r.fields.map((f) => f.name)).toEqual(['编码', '数量', '单价', '单位'])
  })

  it('未配置映射规则时原样透传', () => {
    const bare = defineNode('map', '空对照表', [defineField('键', 'text')], {}, 'bare')
    const r = applyMapping(bare, [{ 键: 'x' }], bill.fields, billRows)
    expect(r.fields.map((f) => f.id)).toEqual(['编码', '数量'])
    expect(r.rows.length).toBe(3)
  })

  it('computeNode(map) 接上游：上游行 + 映射补列', () => {
    const t: Record<string, TableSchema> = {
      ...tables,
      price: { ...price, config: { ...price.config, inputs: ['bill'], mapKeyField: '编码' } },
    }
    const r = computeNode('price', t, records)
    expect(r.rows.length).toBe(3)
    expect(r.fields.map((f) => f.id)).toEqual(['编码', '数量', 'map_单价', 'map_单位'])
    expect(r.rows[0]['map_单价']).toBe(12.5)
    expect(r.rows[2]['map_单位']).toBe('N/A')
  })

  it('computeNode(map) 未接上游：退化为普通对照表，输出自身数据', () => {
    const r = computeNode('price', tables, records)
    expect(r.fields.map((f) => f.id)).toEqual(['编码', '单价', '单位'])
    expect(r.rows.length).toBe(2)
    expect(r.rows[0]['单价']).toBe(12.5)
  })

  it('匹配值带空格也能命中（去空白后比较）', () => {
    const r = applyMapping(price, [{ 编码: ' A001 ', 单价: 99, 单位: 'm' }], bill.fields, billRows)
    expect(r.rows[0]['map_单价']).toBe(99)
  })

  it('对照表同键多行时取第一条', () => {
    const dup = [...priceRows, { 编码: 'A001', 单价: 0.1, 单位: 'cm' }]
    const r = applyMapping(price, dup, bill.fields, billRows)
    expect(r.rows[0]['map_单价']).toBe(12.5)
  })
})
