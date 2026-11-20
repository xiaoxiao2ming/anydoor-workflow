import { describe, it, expect } from 'vitest'
import { defineField, defineTable } from '../core/schema'
import type { TableSchema } from '../core/schema'
import type { Row } from '../core/store'
import type { BindingInfo } from '../docx/binder'
import { aggregate, buildUpstream, resolveBindingTarget, upstreamIdsOf } from './context'
import { buildData } from '../docx/buildData'

/* ── 场景 ──
 *   boq（明细表）  名称 / 数量
 *   price（价目表）名称 / 单价
 *   report（docx） 上游 = boq + price
 */
function fixture() {
  const boq = defineTable('boq', [
    defineField('名称', 'text'),
    defineField('数量', 'number'),
  ])
  const price = defineTable('price', [
    defineField('名称', 'text'),
    defineField('单价', 'number'),
  ])
  const report = defineTable(
    'report',
    [],
    'report',
    'docx',
    { inputs: ['boq', 'price'] },
  )
  const tables: Record<string, TableSchema> = { boq, price, report }
  const records: Record<string, Row[]> = {
    boq: [
      { 名称: '钢管', 数量: 100 },
      { 名称: '阀门', 数量: 12 },
    ],
    price: [
      { 名称: '钢管', 单价: 25.5 },
      { 名称: '阀门', 单价: 260 },
    ],
  }
  return { boq, price, report, tables, records }
}

describe('upstreamIdsOf', () => {
  it('连线优先，其后是面板勾选', () => {
    const t = defineTable('x', [], 'x', 'docx', { linkedInputs: ['b'], inputs: ['a', 'b', 'c'] })
    expect(upstreamIdsOf(t)).toEqual(['b', 'a', 'c'])
  })

  it('没有配置时返回空数组', () => {
    expect(upstreamIdsOf(defineTable('x', []))).toEqual([])
  })

  it('连线与勾选重复时不重复出现', () => {
    const t = defineTable('x', [], 'x', 'docx', { linkedInputs: ['a'], inputs: ['a'] })
    expect(upstreamIdsOf(t)).toEqual(['a'])
  })
})

describe('buildUpstream', () => {
  it('字段取上游并集（保序去重）', () => {
    const { report, tables, records } = fixture()
    const ub = buildUpstream(report, tables, records)
    // boq 的名称/数量 + price 的名称(重复,忽略)/单价
    expect(ub.fields.map((f) => f.name)).toEqual(['名称', '数量', '单价'])
  })

  it('忽略悬空上游（画布上删了节点但配置还留着）', () => {
    const { report, tables, records } = fixture()
    const t: TableSchema = { ...report, config: { inputs: ['boq', '已经不存在的节点'] } }
    const ub = buildUpstream(t, tables, records)
    expect(ub.ids).toEqual(['boq'])
    expect(ub.ownerOf('数量')?.id).toBe('boq')
  })

  it('ownerOf / rowsOf / urnOf 三者对得上', () => {
    const { report, tables, records } = fixture()
    const ub = buildUpstream(report, tables, records)
    expect(ub.ownerOf('单价')).toEqual({ id: 'price', name: '单价' })
    expect(ub.rowsOf('单价')).toHaveLength(2)
    // URN 用字段名而不是内部 id：URN 是给人看、给人写的
    expect(ub.urnOf('单价')).toBe('node://price#单价')
    expect(ub.urnOf('不存在的字段')).toBeNull()
  })

  it('ctx.nodes 同时认 id 和字段名两种键', () => {
    const { report, tables, records } = fixture()
    const ub = buildUpstream(report, tables, records)
    const node = ub.ctx.nodes!.price as Record<string, unknown>
    expect(node['单价']).toBe(25.5)
    expect(Array.isArray(node.rows)).toBe(true)
  })
})

describe('resolveBindingTarget', () => {
  it('按名字自动匹配（未配置绑定）', () => {
    const { report, tables, records } = fixture()
    const ub = buildUpstream(report, tables, records)
    const t = resolveBindingTarget(undefined, '数量', ub)
    expect(t.by).toBe('auto')
    expect(t.fieldId).toBe('数量')
    expect(t.urn).toBe('node://boq#数量')
  })

  it('fieldId 优先于自动匹配', () => {
    const { report, tables, records } = fixture()
    const ub = buildUpstream(report, tables, records)
    // 域名叫「数量」，但显式绑到「单价」
    const t = resolveBindingTarget({ from: 'field', fieldId: '单价' }, '数量', ub)
    expect(t.by).toBe('fieldId')
    expect(t.fieldId).toBe('单价')
    expect(t.urn).toBe('node://price#单价')
  })

  it('source（URN）优先于 fieldId', () => {
    const { report, tables, records } = fixture()
    const ub = buildUpstream(report, tables, records)
    const t = resolveBindingTarget(
      { from: 'field', fieldId: '数量', source: 'node://price#单价' },
      '数量',
      ub,
    )
    expect(t.by).toBe('source')
    expect(t.urn).toBe('node://price#单价')
    // 仍解析出本地字段 id，好让聚合 / 整表取值继续工作
    expect(t.fieldId).toBe('单价')
  })

  it('非 node:// 的 URN（外部文件 / 全局参数）没有本地字段', () => {
    const { report, tables, records } = fixture()
    const ub = buildUpstream(report, tables, records)
    const t = resolveBindingTarget({ from: 'field', source: 'param://审查单位' }, '审查单位', ub)
    expect(t.by).toBe('source')
    expect(t.fieldId).toBeNull()
    expect(t.urn).toBe('param://审查单位')
  })

  it('什么都不匹配 → none', () => {
    const { report, tables, records } = fixture()
    const ub = buildUpstream(report, tables, records)
    expect(resolveBindingTarget(undefined, '不存在的域', ub).by).toBe('none')
  })
})

describe('aggregate', () => {
  const rows: Row[] = [{ v: 1 }, { v: 3 }, { v: 8 }]

  it('各聚合方式', () => {
    expect(aggregate(rows, 'v', 'sum')).toBe(12)
    expect(aggregate(rows, 'v', 'count')).toBe(3)
    expect(aggregate(rows, 'v', 'max')).toBe(8)
    expect(aggregate(rows, 'v', 'min')).toBe(1)
    expect(aggregate(rows, 'v', 'avg')).toBe(4)
    expect(aggregate(rows, 'v', 'first')).toBe(1)
  })

  it('empty 行集 → 空串（不是 undefined，那会让域被跳过）', () => {
    expect(aggregate([], 'v', 'first')).toBe('')
    expect(aggregate([], 'v', 'sum')).toBe('')
  })

  it('值缺失归一到空串', () => {
    expect(aggregate([{}], 'v', 'first')).toBe('')
  })

  it('非数字不参与数值聚合', () => {
    expect(aggregate([{ v: '甲' }, { v: 5 }], 'v', 'sum')).toBe(5)
  })
})

describe('Word 侧接入 URN：buildData + 上游上下文', () => {
  const binding = (kind: BindingInfo['kind'], tag: string, path: string): BindingInfo => ({
    kind,
    path,
    tag,
    part: 'word/document.xml',
    hasRow: false,
    sample: '',
    placeholders: [],
  })

  it('取值为字段时给出 node:// URN 溯源', () => {
    const { report, tables, records } = fixture()
    const ub = buildUpstream(report, tables, records)
    const built = buildData(
      [binding('field', 'field:设计流量', '设计流量')],
      { 'field:设计流量': { from: 'field', fieldId: '数量', agg: 'sum' } },
      ub.fields,
      ub.computed.flatMap((c) => c.rows),
      { ub },
    )
    expect(built.data['设计流量']).toBe(112)
    expect(built.missing).toEqual([])
    const t = built.traces[0]
    expect(t.urn).toBe('node://boq#数量')
    expect(t.from).toBe('node')
    expect(t.ok).toBe(true)
  })

  it('未绑定且名字对不上 → missing + 溯源标记失败', () => {
    const { report, tables, records } = fixture()
    const ub = buildUpstream(report, tables, records)
    const built = buildData([binding('field', 'field:某个域', '某个域')], {}, ub.fields, [], { ub })
    expect(built.missing).toContain('field:某个域')
    expect(built.traces[0].ok).toBe(false)
    expect(built.traces[0].urn).toBe('(未绑定)')
  })

  it('外部引用（param://）解析成功时值来自上下文，失败时标 unresolved', () => {
    const { report, tables, records } = fixture()
    const ub = buildUpstream(report, tables, records)
    // ① 上下文里没有这个参数 → 未解析
    const miss = buildData(
      [binding('field', 'field:审查单位', '审查单位')],
      { 'field:审查单位': { from: 'field', source: 'param://审查单位' } },
      ub.fields,
      [],
      { ub },
    )
    expect(miss.traces[0].ok).toBe(false)
    expect(miss.missing).toContain('field:审查单位')

    // ② 塞进上下文 → 取到值
    const ub2 = buildUpstream(report, tables, records)
    ub2.ctx.params = { 审查单位: '陵水县施工图审查机构' }
    const hit = buildData(
      [binding('field', 'field:审查单位', '审查单位')],
      { 'field:审查单位': { from: 'field', source: 'param://审查单位' } },
      ub2.fields,
      [],
      { ub: ub2 },
    )
    expect(hit.data['审查单位']).toBe('陵水县施工图审查机构')
    expect(hit.traces[0].from).toBe('param')
    expect(hit.traces[0].ok).toBe(true)
  })

  it('each 的 URN 指向哪个上游，就取哪个上游的行', () => {
    const { report, tables, records } = fixture()
    const ub = buildUpstream(report, tables, records)
    const built = buildData(
      [binding('each', 'each:材料表', '材料表')],
      { 'each:材料表': { from: 'rows', source: 'node://price#rows' } },
      ub.fields,
      [],
      { ub },
    )
    const rows = built.data['材料表'] as Row[]
    // price 有 2 行且带「单价」；boq 里没有这个字段
    expect(rows).toHaveLength(2)
    expect(rows[0]['单价']).toBe(25.5)
  })

  it('each 未配置绑定时指向唯一上游的 rows（不标成「未绑定」）', () => {
    const { boq, tables, records } = fixture()
    // 单独造一个只有 1 个上游的 docx 节点
    const doc = defineTable('doc', [], 'doc', 'docx', { inputs: ['boq'] })
    const ub = buildUpstream(doc, { ...tables, doc }, records)
    const built = buildData([binding('each', 'each:明细', '明细')], {}, ub.fields, [], { ub })
    // 明明取了数据，就不能标成未绑定 —— 那是自相矛盾的显示
    expect(built.traces[0].urn).toBe('node://boq#rows')
    expect(built.traces[0].from).toBe('node')
    expect(built.data['明细'] as Row[]).toHaveLength(2)
    void boq
  })

  it('多上游时如实说「全部上游合并」，不假装指向某一个', () => {
    const { report, tables, records } = fixture()
    const ub = buildUpstream(report, tables, records)
    const built = buildData(
      [binding('each', 'each:汇总', '汇总')],
      {},
      ub.fields,
      ub.computed.flatMap((c) => c.rows),
      { ub },
    )
    expect(built.resolved['each:汇总']).toContain('全部上游合并')
    // boq 2 行 + price 2 行
    expect(built.data['汇总'] as Row[]).toHaveLength(4)
  })

  it('不给上游上下文时退回 fieldId 模式（老签名兼容）', () => {
    const fields = [defineField('数量', 'number')]
    const built = buildData(
      [binding('field', 'field:数量', '数量')],
      {},
      fields,
      [{ 数量: 7 }],
    )
    expect(built.data['数量']).toBe(7)
    expect(built.traces[0].urn).toBe('(按名自动匹配)')
  })
})
