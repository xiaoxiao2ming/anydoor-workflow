import { describe, it, expect } from 'vitest'
import { defineTable } from '../core/schema'
import type { Row } from '../core/store'
import {
  resolveNodeFields, readableTokenRanges, readableToUrn, urnToReadable, fxTokenRanges, FX_URN_RE,
} from './fxUrn'

/* ════════════════════════════════════════════════════════════════════
 * 批次 17 回归锁：可读串 ⇄ URN 互转的字段口径。
 *
 * 故障原文：点「数据排序 (2).数量」能插入（面板显示正常），但
 *   ① fx 栏不画胶囊包（data-fx-urn=0）—— 用户看不出这是个引用
 *   ② 提交后单元格存的是 `数据排序 (2)·数量` 原文，不是 node://…
 *      → isPureUrn 判否 → 当普通公式求值 → 联动规则永远不命中
 * 根因：sort/map/output/docx/xlsx 是**汇聚型**节点，自身 schema t.fields 为空
 *      （字段全继承自上游 mergeComputed），而旧代码拿 t.fields 枚举字段。
 * 修复：口径统一到 computeNode 的计算后字段（resolveNodeFields）。
 * ════════════════════════════════════════════════════════════════════ */

/** 一条普通数据表：字段在 schema 里（table/param 型，t.fields 非空） */
const srcTable = () =>
  defineTable('参数表', [{ id: 'price', name: '单价', type: 'number' }], 'param1', 'param')

/** 一条汇聚型节点：schema fields 为空，字段全在上游 —— 这就是 sort 的真实形状 */
const aggTable = (inputs: string[] = ['param1']) =>
  defineTable('数据排序', [], 'sort1', 'sort', { linkedInputs: inputs })

const records: Record<string, Row[]> = {
  param1: [{ id: 'r1', price: 25.5 } as unknown as Row],
  sort1: [] as Row[],
}

describe('resolveNodeFields · 字段口径必须是 computeNode 的计算后结果', () => {
  it('汇聚型节点：schema t.fields 为空，但计算后能拿到上游继承的字段', () => {
    const tables = { param1: srcTable(), sort1: aggTable() }
    // 前置事实：schema 里确实是空的（这正是老口径失效的原因）
    expect(tables.sort1.fields).toHaveLength(0)

    const rf = resolveNodeFields(tables, records)
    // 修复后：sort1 也枚举出「数量」类的字段（来自上游 param1.price）
    expect(Object.keys(rf)).toContain('sort1')
    expect(rf.sort1.fields.length).toBeGreaterThan(0)
    expect(rf.param1.fields.map((f) => f.name)).toContain('单价')
  })

  it('selfId 自身与已删节点被排除', () => {
    const tables = { param1: srcTable(), sort1: aggTable(), gone: { ...srcTable(), id: 'gone', deleted: true } }
    const rf = resolveNodeFields(tables, records, 'sort1')
    expect(Object.keys(rf)).toContain('param1')
    expect(Object.keys(rf)).not.toContain('sort1') // 自身
    expect(Object.keys(rf)).not.toContain('gone') // 已删
  })

  it('tables 为空 → 返回空表而不是抛错', () => {
    expect(resolveNodeFields(null, records)).toEqual({})
    expect(resolveNodeFields(undefined, null)).toEqual({})
  })
})

describe('readableTokenRanges · 可读 token 定位', () => {
  const rf = () => resolveNodeFields({ param1: srcTable(), sort1: aggTable() }, records)

  it('普通表的可读串能命中', () => {
    const r = readableTokenRanges('=参数表·单价', rf())
    expect(r).toHaveLength(1)
    expect(r[0].urn).toBe('node://param1#price')
    expect(r[0].dead).toBe(false)
  })

  it('★ 回归锁：汇聚型节点的可读串也必须命中（旧口径这里返回 0 条）', () => {
    const f = rf()
    const name = f.sort1.fields[0].name
    const s = `=数据排序·${name}`
    const r = readableTokenRanges(s, f)
    expect(r).toHaveLength(1)
    expect(r[0].urn).toBe(`node://sort1#${f.sort1.fields[0].id}`)
  })

  it('长名优先：短 token 是长 token 的前缀时不误吞', () => {
    const t = defineTable('表', [{ id: 'a', name: '值', type: 'number' }], 'n1')
    const t2 = defineTable('表', [{ id: 'b', name: '数值', type: 'number' }], 'n2')
    const f = resolveNodeFields({ n1: t, n2: t2 }, { n1: [], n2: [] })
    // 「表·值」与「表·数值」：长名优先 → 必须先匹「表·数值」
    const r = readableTokenRanges('表·数值', f)
    expect(r).toHaveLength(1)
    expect(r[0].urn).toBe('node://n2#b')
  })

  it('串里没有 · 时不产出 token', () => {
    expect(readableTokenRanges('=B2*C3', rf())).toHaveLength(0)
  })
})

describe('readableToUrn · 提交边界（可读串 → URN）', () => {
  const rf = () => resolveNodeFields({ param1: srcTable(), sort1: aggTable() }, records)

  it('普通 token 转换正确', () => {
    expect(readableToUrn('参数表·单价', rf())).toBe('node://param1#price')
  })

  it('★ 回归锁：汇聚型节点的 token 能转回 URN（旧口径原样返回 → 单元格存进无效文本）', () => {
    const f = rf()
    const fld = f.sort1.fields[0]
    const out = readableToUrn(`数据排序·${fld.name}`, f)
    expect(out).toBe(`node://sort1#${fld.id}`)
    expect(out.startsWith('node://')).toBe(true)
  })

  it('相邻 token 自动补 +（不产生 #pricenode:// 粘连）', () => {
    expect(readableToUrn('参数表·单价+参数表·单价', rf())).toBe('node://param1#price+node://param1#price')
  })

  it('非 token 文本原样保留', () => {
    expect(readableToUrn('SUM(B2:C3)', rf())).toBe('SUM(B2:C3)')
  })

  it('无 · 或空表 → 原样返回', () => {
    expect(readableToUrn('abc', rf())).toBe('abc')
    expect(readableToUrn('参数表·单价', {})).toBe('参数表·单价')
  })

  it('往返一致：可读 → URN → 可读 回到原串', () => {
    const f = rf()
    const urn = readableToUrn('参数表·单价', f)
    expect(urnToReadable(urn, f)).toBe('参数表·单价')
  })
})

describe('urnToReadable · 打开方向（URN → 可读）', () => {
  const rf = () => resolveNodeFields({ param1: srcTable(), sort1: aggTable() }, records)

  it('正常 URN 转可读', () => {
    expect(urnToReadable('node://param1#price', rf())).toBe('参数表·单价')
  })

  it('★ 汇聚型节点的 URN 也转得成可读（否则打开单元格看到一串 node:// 噪音）', () => {
    const f = rf()
    const fld = f.sort1.fields[0]
    expect(urnToReadable(`node://sort1#${fld.id}`, f)).toBe(`数据排序·${fld.name}`)
  })

  it('查不到表/参数 → 保留 URN 原文（不丢信息）', () => {
    expect(urnToReadable('node://nope#x', rf())).toBe('node://nope#x')
    expect(urnToReadable('node://param1#nope', rf())).toBe('node://param1#nope')
  })
})

describe('fxTokenRanges · URN 原文 + 可读 token 合并', () => {
  const rf = () => resolveNodeFields({ param1: srcTable(), sort1: aggTable() }, records)

  it('手打 URN 仍被识别为 token（向下兼容）', () => {
    const r = fxTokenRanges('=node://param1#price', rf())
    expect(r).toHaveLength(1)
    expect(r[0].dead).toBe(false)
    expect(r[0].node).toBe('参数表')
  })

  it('失效 URN 标 dead（节点已删/字段已改名 → 红边胶囊）', () => {
    const r = fxTokenRanges('=node://gone#x', rf())
    expect(r).toHaveLength(1)
    expect(r[0].dead).toBe(true)
  })

  it('可读 + URN 混排：按位置排序且不重叠', () => {
    const r = fxTokenRanges('参数表·单价+node://param1#price', rf())
    expect(r).toHaveLength(2)
    expect(r[0].end).toBeLessThanOrEqual(r[1].start)
  })
})

describe('FX_URN_RE · URN token 边界', () => {
  it('匹配 node://id#field（含中文 id）', () => {
    const m = [...'node://排序_1#数量 +x'.matchAll(FX_URN_RE)]
    expect(m).toHaveLength(1)
    expect(m[0][0]).toBe('node://排序_1#数量')
  })

  it('终止于运算符与括号（不吞后面的公式）', () => {
    const m = [...'node://a#b+C2'.matchAll(FX_URN_RE)]
    expect(m[0][0]).toBe('node://a#b')
  })
})
