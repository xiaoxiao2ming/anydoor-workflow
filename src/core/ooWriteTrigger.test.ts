/**
 * OO 写回触发判定测试
 *
 * 立档：2026-10-06 · 配套 src/core/ooWriteTrigger.ts
 */

import { describe, it, expect } from 'vitest'
import { decideWriteback, diffRefs, shouldWriteOnFirstBind, type Snapshot } from './ooWriteTrigger'
import { ooRefUrn, type OoRef } from './ooRef'

/** ★ 每个用例一套独立数据。
 *  上一版用**模块级共享** data，第一条改了 t1.f1=99 之后没复位，
 *  连带后面两条红掉（"都没变"那��实际 f1 已是 99）。
 *  ⇒ 测试间共享可变状态是最容易伪装成"产品有 bug"的一种坑。
 */
function makeData () {
  const d: Record<string, Record<string, unknown>> = {
    t1: { f1: 12, f2: '甲' },
    t2: { qty: 5 },
  }
  return { d, readNode: (n: string, f: string) => d[n]?.[f] }
}

const REFS: OoRef[] = [
  { nodeId: 't1', fieldId: 'f1', urn: ooRefUrn('t1', 'f1') },
  { nodeId: 't2', fieldId: 'qty', urn: ooRefUrn('t2', 'qty') },
]
const urnOf = (r: OoRef) => r.urn ?? ooRefUrn(r.nodeId, r.fieldId)
const base = (): Snapshot => ({ 'node://t1#f1': { __v: 12 }, 'node://t2#qty': { __v: 5 } })

describe('① 快照比对：只报真正变了的引用', () => {
  it('两个引用，其中一个变 ⇒ 只报那一个', () => {
    const { d, readNode } = makeData()
    d.t1.f1 = 99
    const r = decideWriteback(REFS, urnOf, readNode, base())
    expect(r.write).toBe(true)
    expect(r.changes).toHaveLength(1)
    expect(r.changes[0]).toMatchObject({ urn: 'node://t1#f1', before: 12, after: 99 })
  })

  it('都没变 ⇒ 不写', () => {
    const { readNode } = makeData()          // 全新数据 =与基线一致
    const r = decideWriteback(REFS, urnOf, readNode, base())
    expect(r.write).toBe(false)
    expect(r.why).toContain('未变化')
  })

  it('★ 当前快照每次都刷新（下次比对的基线是最新的）', () => {
    const { readNode } = makeData()
    const r = decideWriteback(REFS, urnOf, readNode, base())
    expect(r.current['node://t1#f1'].__v).toBe(12)
  })

  it('★ 数字串与数字视为同一族（沿用 trigger.ts 既有口径）', () => {
    // 12 → "12" 在 trigger.ts 的 kindOf 里同族（'num'）⇒ 判「没变」，**这是既定设计**：
    //   kindOf: `typeof v === 'string' ? (toNum(v) != null ? 'num' : 'str') : ...`
    //   注释原文：「toNum 会把 false→0、'1.0'→1，若不先拦类型，`0 → false`、`'1.0' → 1`
    //   会被误判成"没变"」—— 数字串被**刻意**归入num 族。
    // ⇒ 这里断言它**不**写回，避免以后有人"顺手改成严格 ==="而与全仓口径分裂。
    const { d, readNode } = makeData()
    d.t1.f1 = '12'
    const r = decideWriteback([REFS[0]], urnOf, readNode, base())
    expect(r.write).toBe(false)
  })

  it('真正跨类型族时算变更（num → bool）', () => {
    const { d, readNode } = makeData()
    d.t1.f1 = true
    const r = decideWriteback([REFS[0]], urnOf, readNode, base())
    expect(r.write).toBe(true)
  })

  it('空 → 有值算变更（null 族→num 族）', () => {
    const { d, readNode } = makeData()
    d.t1.f1 = 0
    const b: Snapshot = { 'node://t1#f1': { __v: null } }
    const r = decideWriteback([REFS[0]], urnOf, readNode, b)
    expect(r.write).toBe(true)
  })
})

describe('★★ 首次建立基线：刻意不写回（反直觉但必要）', () => {
  it('首次（prev=null）不写', () => {
    const { readNode } = makeData()
    const d = decideWriteback(REFS, urnOf, readNode, null)
    expect(d.write).toBe(false)
    expect(d.why).toContain('首次')
  })

  it('但基线会被建立 ⇒ 第二次起就能正常判定', () => {
    const { d, readNode } = makeData()
    const first = decideWriteback(REFS, urnOf, readNode, null)
    expect(first.write).toBe(false)
    // ★ 用它返回的 current 作为下次的 prev
    d.t1.f1 = 77
    const second = decideWriteback(REFS, urnOf, readNode, first.current)
    expect(second.write).toBe(true)
    expect(second.changes[0].after).toBe(77)
  })

  it('why 明确说明理由（用户能看懂为什么没刷新）', () => {
    const { readNode } = makeData()
    const d = decideWriteback(REFS, urnOf, readNode, null)
    expect(d.why).toMatch(/首次建立基线/)
  })

  it('策略开关可查（默认关）', () => {
    expect(shouldWriteOnFirstBind()).toBe(false)
  })
})

describe('③ 边界情形如实报告，不静默', () => {
  it('没有域引用 ⇒ 不写，理由明确', () => {
    const { readNode } = makeData()
    const d = decideWriteback([], urnOf, readNode, {})
    expect(d.write).toBe(false)
    expect(d.why).toContain('没有域引用')
  })

  it('上游节点被删 ⇒ 值为 undefined，仍能判定（不崩）', () => {
    const { readNode } = makeData()
    const b: Snapshot = { 'node://t1#f1': { __v: 12 } }
    const dead: OoRef[] = [{ nodeId: 'gone', fieldId: 'x' }]
    const d = decideWriteback(dead, urnOf, readNode, b)
    expect(d.write).toBe(false)
  })

  it('why 列出变了多少个（便于诊断）', () => {
    const { d, readNode } = makeData()
    const b: Snapshot = { 'node://t1#f1': { __v: 1 }, 'node://t2#qty': { __v: 1 } }
    d.t1.f1 = 2
    d.t2.qty = 2
    const r = decideWriteback(REFS, urnOf, readNode, b)
    expect(r.write).toBe(true)
    expect(r.why).toContain('2 个引用')
  })
})

describe('④ diffRefs 可单独使用（拿当前快照）', () => {
  it('返回 current 供外部存基线', () => {
    const { readNode } = makeData()
    const { current } = diffRefs(REFS, urnOf, readNode, null)
    expect(Object.keys(current).sort()).toEqual(['node://t1#f1', 'node://t2#qty'])
  })
})
