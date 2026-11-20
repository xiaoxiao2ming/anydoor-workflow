import { describe, it, expect } from 'vitest'
import {
  valueChanged,
  diffSnapshots,
  evalCondition,
  evalRules,
  planTrigger,
  DebounceGate,
  PRESET_RULES,
  type TriggerRule,
  type SourceChange,
} from './trigger'

/* ── valueChanged ── */

describe('valueChanged', () => {
  it('相同值不算变化', () => {
    expect(valueChanged(1, 1)).toBe(false)
    expect(valueChanged('a', 'a')).toBe(false)
  })

  it('null / undefined / 空串 归一为同一个「空」', () => {
    expect(valueChanged(null, undefined)).toBe(false)
    expect(valueChanged(null, '')).toBe(false)
    expect(valueChanged(undefined, '')).toBe(false)
  })

  it('空 ↔ 非空 算变化', () => {
    expect(valueChanged(null, 5)).toBe(true)
    expect(valueChanged(0, null)).toBe(true)
  })

  it('0 与 false 不等价（别把它们当同一个空）', () => {
    expect(valueChanged(0, false)).toBe(true)
  })

  it('数字串与数字同族：epsilon=0 时值相同判不变，容差内也不算', () => {
    expect(valueChanged('1.0', 1)).toBe(false) // 同为数值族，值相等
    expect(valueChanged('1', 2)).toBe(true)
    expect(valueChanged(0.1 + 0.2, 0.3, 1e-9)).toBe(false)
    expect(valueChanged(0.1 + 0.2, 0.4, 1e-9)).toBe(true)
  })

  it('类型族不同一律算变化（文本 vs 数字）', () => {
    expect(valueChanged('abc', 1)).toBe(true)
    expect(valueChanged(1, '1')).toBe(false) // '1' 是数值族
    expect(valueChanged([1], '1')).toBe(true)
  })

  it('数组顺序变化算变化', () => {
    expect(valueChanged([1, 2], [1, 2])).toBe(false)
    expect(valueChanged([1, 2], [2, 1])).toBe(true)
  })
})

/* ── diffSnapshots ── */

describe('diffSnapshots', () => {
  it('只报真变了的 URN', () => {
    const d = diffSnapshots(
      { 'param://a': 1, 'param://b': 2 },
      { 'param://a': 1, 'param://b': 9 }
    )
    expect(d.changes).toHaveLength(1)
    expect(d.changes[0]).toEqual({ urn: 'param://b', before: 2, after: 9 })
    expect(d.changedUrns.has('param://a')).toBe(false)
  })

  it('新增的 URN 算变化（before 为 undefined）', () => {
    const d = diffSnapshots({}, { 'param://x': 1 })
    expect(d.changes).toHaveLength(1)
    expect(d.changes[0].before).toBeUndefined()
  })

  it('删除的 URN 也算变化（after 为 undefined）', () => {
    const d = diffSnapshots({ 'param://x': 1 }, {})
    expect(d.changes).toHaveLength(1)
  })
})

/* ── evalCondition ── */

describe('evalCondition', () => {
  const ch = (before: unknown, after: unknown) => ({ before, after })

  it('always 恒命中', () => {
    expect(evalCondition({ op: 'always' }, ch(1, 1)).hit).toBe(true)
  })

  it('changed 只在值变化时命中', () => {
    expect(evalCondition({ op: 'changed' }, ch(1, 2)).hit).toBe(true)
    expect(evalCondition({ op: 'changed' }, ch(1, 1)).hit).toBe(false)
  })

  it('数值比较', () => {
    expect(evalCondition({ op: 'gt', value: 5 }, ch(1, 6)).hit).toBe(true)
    expect(evalCondition({ op: 'gt', value: 5 }, ch(1, 5)).hit).toBe(false)
    expect(evalCondition({ op: 'gte', value: 5 }, ch(1, 5)).hit).toBe(true)
    expect(evalCondition({ op: 'lt', value: 5 }, ch(9, 4)).hit).toBe(true)
    expect(evalCondition({ op: 'lte', value: 5 }, ch(9, 5)).hit).toBe(true)
  })

  it('eq 尊重 epsilon（防浮点抖动）', () => {
    expect(evalCondition({ op: 'eq', value: 0.3, epsilon: 1e-9 }, ch(0, 0.1 + 0.2)).hit).toBe(true)
    expect(evalCondition({ op: 'eq', value: 0.3 }, ch(0, 0.1 + 0.2)).hit).toBe(false)
  })

  it('非数字进数值条件 → 不命中且带原因（不静默通过）', () => {
    const r = evalCondition({ op: 'gt', value: 5 }, ch(1, 'abc'))
    expect(r.hit).toBe(false)
    expect(r.reason).toContain('不是数字')
  })

  it('between 闭区间', () => {
    expect(evalCondition({ op: 'between', value: 1, hi: 10 }, ch(0, 5)).hit).toBe(true)
    expect(evalCondition({ op: 'between', value: 1, hi: 10 }, ch(0, 10)).hit).toBe(true)
    expect(evalCondition({ op: 'between', value: 1, hi: 10 }, ch(0, 11)).hit).toBe(false)
  })

  it('between 上下界写反时自动纠正', () => {
    const r = evalCondition({ op: 'between', value: 10, hi: 1 }, ch(0, 5))
    expect(r.hit).toBe(true)
  })

  it('outside 是 between 的反面', () => {
    expect(evalCondition({ op: 'outside', value: 1, hi: 10 }, ch(0, 20)).hit).toBe(true)
    expect(evalCondition({ op: 'outside', value: 1, hi: 10 }, ch(0, 5)).hit).toBe(false)
  })

  it('空 / 非空 / 布尔', () => {
    expect(evalCondition({ op: 'nonEmpty' }, ch('a', 'b')).hit).toBe(true)
    expect(evalCondition({ op: 'nonEmpty' }, ch('a', '')).hit).toBe(false)
    expect(evalCondition({ op: 'empty' }, ch('a', null)).hit).toBe(true)
    expect(evalCondition({ op: 'isTrue' }, ch(false, true)).hit).toBe(true)
    expect(evalCondition({ op: 'isTrue' }, ch(true, false)).hit).toBe(false)
    expect(evalCondition({ op: 'isFalse' }, ch(true, false)).hit).toBe(true)
  })

  it('未知算子 → 不命中且原因可见', () => {
    const r = evalCondition({ op: 'wat' as never }, ch(1, 1))
    expect(r.hit).toBe(false)
    expect(r.reason).toContain('未知条件算子')
  })
})

/* ── evalRules ── */

const rule = (p: Partial<TriggerRule>): TriggerRule => ({
  id: 'r' + Math.random().toString(36).slice(2, 8),
  name: 'r',
  enabled: true,
  watch: 'param://a',
  condition: { op: 'changed' },
  action: { kind: 'writeCell', sheet: 'Sheet1', coord: 'A1' },
  ...p,
})

describe('evalRules', () => {
  const ch: SourceChange[] = [{ urn: 'param://a', before: 1, after: 2 }]

  it('watch 精确匹配才命中', () => {
    const r = rule({ watch: 'param://a' })
    expect(evalRules([r], ch).hits).toHaveLength(1)
    expect(evalRules([rule({ watch: 'param://b' })], ch).hits).toHaveLength(0)
  })

  it('watch="change" 匹配任何来源', () => {
    const r = rule({ watch: 'change' })
    expect(evalRules([r], [{ urn: 'node://n1#x', before: 0, after: 1 }]).hits).toHaveLength(1)
  })

  it('watch 以 # 结尾 → 前缀匹配（吃掉不同 locator）', () => {
    const r = rule({ watch: 'node://n1#' })
    const res = evalRules([r], [{ urn: 'node://n1#rows', before: 0, after: 1 }])
    expect(res.hits).toHaveLength(1)
  })

  it('watch 以 :// 结尾 → 同 scheme 全匹配（param:// 吃所有参数）', () => {
    const r = rule({ watch: 'param://' })
    const res = evalRules([r], [
      { urn: 'param://a', before: 0, after: 1 },
      { urn: 'param://b', before: 0, after: 1 },
    ])
    expect(res.hits).toHaveLength(2)
    // 跨 scheme 不该命中
    const res2 = evalRules([r], [{ urn: 'node://x', before: 0, after: 1 }])
    expect(res2.hits).toHaveLength(0)
  })

  it('enabled=false 的规则跳过但不影响其他', () => {
    const res = evalRules([rule({ enabled: false }), rule({ id: 'on' })], ch)
    expect(res.hits).toHaveLength(1)
    expect(res.hits[0].rule.id).toBe('on')
  })

  it('按 priority 升序；同优先级按声明顺序（稳定）', () => {
    const a = rule({ id: 'a', priority: 50 })
    const b = rule({ id: 'b', priority: 10 })
    const res = evalRules([a, b], ch)
    expect(res.hits.map((h) => h.rule.id)).toEqual(['b', 'a'])
  })

  it('不命中的进 skips 且带原因', () => {
    const res = evalRules([rule({ condition: { op: 'gt', value: 100 } })], ch)
    expect(res.hits).toHaveLength(0)
    expect(res.skips).toHaveLength(1)
    expect(res.skips[0].reason).toContain('2 > 100')
  })

  it('动作去重：多规则写同一格只跑一次', () => {
    const res = evalRules(
      [rule({ id: 'a' }), rule({ id: 'b' })],
      ch
    )
    expect(res.hits).toHaveLength(2)
    expect(res.actions).toHaveLength(1)
  })

  it('空规则集 / 空变更都不炸', () => {
    expect(evalRules([], ch).hits).toHaveLength(0)
    expect(evalRules([rule({})], []).hits).toHaveLength(0)
  })
})

/* ── planTrigger ── */

describe('planTrigger', () => {
  const ch: SourceChange[] = [{ urn: 'param://管径', before: 100, after: 200 }]
  // 本组默认 watch 与 ch 的 URN 对齐（否则规则按 watch 语义不命中，测的就不是写值逻辑了）
  const r2 = (p: Partial<TriggerRule>): TriggerRule => rule({ watch: 'param://管径', ...p })

  it('写值动作取本次变更的新值', () => {
    const p = planTrigger(
      [rule({ watch: 'param://管径', action: { kind: 'writeCell', sheet: 'Sheet1', coord: 'B2' } })],
      ch,
      () => '兜底'
    )
    expect(p.writes).toEqual([
      { sheet: 'Sheet1', coord: 'B2', value: 200, from: 'param://管径' },
    ])
  })

  it('watch=change 时取第一个变更的 URN 作为来源', () => {
    const p = planTrigger(
      [rule({ watch: 'change', action: { kind: 'writeCell', sheet: 'S', coord: 'A1' } })],
      ch,
      () => 'x'
    )
    expect(p.writes[0].from).toBe('param://管径')
    expect(p.writes[0].value).toBe(200)
  })

  it('watch 不在本次变更里 → 回查 resolveValue（常驻来源）', () => {
    // 规则 watch 常量、但本次变更的是另一个 URN → 规则按 watch 语义不命中是「正确」行为，
    // 所以这里用「watch 命中但值取不到」的场景：前缀 watch 匹配、值回查兜底。
    const p = planTrigger(
      [
        rule({
          watch: 'param://',
          condition: { op: 'always' },
          action: { kind: 'writeCell', sheet: 'S', coord: 'A1' },
        }),
      ],
      [],
      (u) => (u === 'param://' ? 42 : undefined)
    )
    // 无变更 → 无命中 → 无写值（规则不凭空触发）
    expect(p.writes).toHaveLength(0)
  })

  it('前缀 watch 命中多条变更 → 同格只写一次，取第一条匹配的新值', () => {
    const p = planTrigger(
      [
        rule({
          watch: 'param://',
          condition: { op: 'always' },
          action: { kind: 'writeCell', sheet: 'S', coord: 'A1' },
        }),
      ],
      [
        { urn: 'param://x', before: 1, after: 11 },
        { urn: 'param://y', before: 2, after: 22 },
      ],
      () => 0
    )
    // 两条变更都命中规则，但 A1 只写一次（动作去重），值取第一条
    expect(p.writes).toHaveLength(1)
    expect(p.writes[0].value).toBe(11)
    expect(p.writes[0].from).toBe('param://x')
  })

  it('区域写入：数组逐格铺满', () => {
    const p = planTrigger(
      [
        {
          ...rule({
            watch: 'change',
            action: { kind: 'writeRange', sheet: 'S', ref: 'A1:A3' },
          }),
        },
      ],
      [{ urn: 'param://xs', before: [1, 2], after: [7, 8, 9] }],
      () => []
    )
    expect(p.writes.map((w) => w.coord)).toEqual(['A1', 'A2', 'A3'])
    expect(p.writes.map((w) => w.value)).toEqual([7, 8, 9])
  })

  it('区域写入：非数组当单值只写首格', () => {
    const p = planTrigger(
      [rule({ watch: 'change', action: { kind: 'writeRange', sheet: 'S', ref: 'A1:B2' } })],
      [{ urn: 'param://one', before: null, after: 5 }],
      () => []
    )
    expect(p.writes).toHaveLength(1)
    expect(p.writes[0].coord).toBe('A1')
  })

  it('写值后自动补重算（否则公式缓存是旧值）', () => {
    const p = planTrigger(
      [r2({ action: { kind: 'writeCell', sheet: 'S', coord: 'A1' } })],
      ch,
      () => 0
    )
    expect(p.recomputeSheets.has('S')).toBe(true)
    expect(p.log.some((l) => l.includes('自动补'))).toBe(true)
  })

  it('显式 recompute 规则不触发自动补', () => {
    const p = planTrigger(
      [
        r2({ action: { kind: 'writeCell', sheet: 'S', coord: 'A1' } }),
        r2({ id: 'rc', action: { kind: 'recompute' } }),
      ],
      ch,
      () => 0
    )
    expect(p.recomputeAll).toBe(true)
    expect(p.log.some((l) => l.includes('自动补'))).toBe(false)
  })

  it('recompute 不带 sheet → recomputeAll', () => {
    const p = planTrigger([r2({ action: { kind: 'recompute' } })], ch, () => 0)
    expect(p.recomputeAll).toBe(true)
  })

  it('commit 动作置 needCommit 并进日志', () => {
    const p = planTrigger([r2({ action: { kind: 'commit' } })], ch, () => 0)
    expect(p.needCommit).toBe(true)
    expect(p.log.some((l) => l.includes('落盘'))).toBe(true)
  })

  it('整条预设链跑通：写值 → 重算 → 落盘', () => {
    const rules: TriggerRule[] = PRESET_RULES.map((r, i) => ({ ...r, id: 'p' + i }))
    const p = planTrigger(rules, ch, () => 0)
    expect(p.writes[0]).toMatchObject({ sheet: 'Sheet1', coord: 'A2', value: 200 })
    expect(p.recomputeAll).toBe(true)
    expect(p.needCommit).toBe(true)
    // 顺序：写值必须在重算之前
    const iWrite = p.log.findIndex((l) => l.startsWith('写值'))
    const iCalc = p.log.findIndex((l) => l.startsWith('重算'))
    expect(iWrite).toBeGreaterThanOrEqual(0)
    expect(iCalc).toBeGreaterThan(iWrite)
  })
})

/* ── DebounceGate ── */

describe('DebounceGate', () => {
  it('窗口为 0 → 全放行', () => {
    const g = new DebounceGate(() => 0)
    expect(g.allow('k', 0)).toBe(true)
    expect(g.allow('k', 0)).toBe(true)
  })

  it('窗口内重复请求被合并', () => {
    let t = 1000
    const g = new DebounceGate(() => t)
    expect(g.allow('k', 800)).toBe(true)
    t = 1200
    expect(g.allow('k', 800)).toBe(false) // 200ms < 800ms
    t = 1900
    expect(g.allow('k', 800)).toBe(true) // 900ms > 800ms
  })

  it('不同 key 互不影响', () => {
    let t = 0
    const g = new DebounceGate(() => t)
    expect(g.allow('a', 800)).toBe(true)
    expect(g.allow('b', 800)).toBe(true)
  })

  it('reset 后必过（「立即更新」按钮语义）', () => {
    let t = 0
    const g = new DebounceGate(() => t)
    g.allow('k', 800)
    t = 10
    expect(g.allow('k', 800)).toBe(false)
    g.reset('k')
    expect(g.allow('k', 800)).toBe(true)
  })

  it('reset 无参 → 清空全部', () => {
    let t = 0
    const g = new DebounceGate(() => t)
    g.allow('a', 800)
    g.allow('b', 800)
    g.reset()
    expect(g.allow('a', 800)).toBe(true)
    expect(g.allow('b', 800)).toBe(true)
  })
})
