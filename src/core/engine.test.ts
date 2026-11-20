import { describe, it, expect } from 'vitest'
import { mergeInputs, sortRows, computeNode } from './engine'
import { defineField, defineNode, type TableSchema } from './schema'
import type { Row } from '../store'

const a = defineNode('table', 'A', [
  defineField('name', 'text'),
  defineField('score', 'number', { id: 'score' }),
], {}, 'A')

const rowsA: Row[] = [
  { name: 'b', score: 3 },
  { name: 'a', score: 1 },
  { name: 'c', score: 2 },
]

const tables: Record<string, TableSchema> = { A: a }
const records: Record<string, Row[]> = { A: rowsA }

describe('数据流引擎', () => {
  it('mergeInputs 合并字段与行', () => {
    const m = mergeInputs(tables, records, ['A'])
    expect(m.fields.length).toBe(2)
    expect(m.rows.length).toBe(3)
  })

  it('sortRows 按数值升序', () => {
    const sorted = sortRows(rowsA, a.fields, 'score', 'asc')
    expect(Number(sorted[0].score)).toBe(1)
    expect(Number(sorted[2].score)).toBe(3)
  })

  it('computeNode(sort) 按配置排序且含上游数据', () => {
    tables.S = defineNode('sort', 'S', [], { inputs: ['A'], sortBy: 'score', sortDir: 'asc' }, 'S')
    const r = computeNode('S', tables, records)
    expect(r.rows.map((x) => Number(x.score))).toEqual([1, 2, 3])
    expect(r.fields.map((f) => f.id)).toEqual(['name', 'score'])
  })

  it('computeNode(output) 汇入上游全部行', () => {
    tables.O = defineNode('output', 'O', [], { inputs: ['A'] }, 'O')
    const r = computeNode('O', tables, records)
    expect(r.rows.length).toBe(3)
  })

  it('防环：环里不崩溃', () => {
    // A 的 input 指向 S，S 的 input 指向 A，构成环
    tables.S = defineNode('sort', 'S', [], { inputs: ['A'] }, 'S')
    tables.A = defineNode('sort', 'A2', [], { inputs: ['S'] }, 'A')
    const r = computeNode('A', tables, records)
    expect(Array.isArray(r.rows)).toBe(true)
  })
})
