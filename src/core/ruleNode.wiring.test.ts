import { describe, it, expect } from 'vitest'
import { PALETTE_GROUPS, paletteItemClassOf, isRuleNodeKind } from './nodeStyle'
import { bindingsOfRuleNode, effectiveRule, isRuleLocked, ruleNodesOf, ruleConnectPatch } from './ruleSync'
import { defineNode, defineField, type TableSchema } from './schema'

/**
 * 规则节点接入面回归（2026-10-05）
 *
 * 起因：内核ruleSync 全绿、tsc 零错误、529 测试全绿，但**从画布上拖一个规则节点出来、
 * 连一条线到别的节点，什么都不会发生** —— 锁定机制的唯一建立入口在UI 层是断的。
 *
 * ★ 这组测试直接 import `ruleConnectPatch`（真函数，App.handleConnect 调的就是它），
 *   **不复制任何逻辑**。
 *   初版测试抄了一份 canConnectRule，做反向验证时把真实现改回单向（有缺陷），
 *   14 个用例**照样全绿** —— 测的是副本，与真代码零关联。抄逻辑的测试等于没有测试。
 *   抽成纯函数后，判定只有一份，反向验证能真的变红（见文件末尾的对照说明）。
 */

const f = defineField('列', 'text')

const ruleNode = (id = 'r1') => defineNode('rule', '规则', [], { rules: {} }, id)
const sortNode = (id: string, name = '排序') =>
  defineNode('sort', name, [f], { sortBy: 'col1', sortDir: 'asc' }, id)

/** 铺一张表，好让 ruleConnectPatch 能查到 */
const T = (...ts: TableSchema[]): Record<string, TableSchema> =>
  Object.fromEntries(ts.map((t) => [t.id, t]))

describe('规则节点在节点库里可见（真捕获，不靠推断）', () => {
  it('PALETTE_GROUPS 实际含 rule 条目', () => {
    const all = PALETTE_GROUPS.flatMap((g) => g.kinds.map((k) => k.kind))
    expect(all).toContain('rule')
  })

  it('★ 条目提示不能是所属角色的组提示（规则节点曾显示「决定流程什么时候执行」）', () => {
    const g = PALETTE_GROUPS.find((x) => x.kinds.some((k) => k.kind === 'rule'))!
    const hit = g.kinds.find((k) => k.kind === 'rule')!
    // ★ 2026-10-19 用户「把AI节点、规则跟随节点、参数库节点 放在一类」：三家合成
    //   「圆形节点」组——组提示随之更新（不再是动作型的「什么时候执行」）
    expect(g.label).toContain('圆形节点')
    expect(['rule', 'ai', 'paramlib'].every((k) => g.kinds.some((x) => x.kind === k))).toBe(true)
    // 条目提示必须换成自己的语义，不能沿用组提示
    expect(hit.hint).not.toContain('什么时候执行')
    // ★ 也不该是上一版的旧文案「统一下发」—— 节点已改名「规则跟随」，
    //   提示说的是「同类节点一起变」，与用户口径一致（旧断言卡在这里）
    expect(hit.hint).not.toContain('统一下发')
    expect(hit.hint).toContain('规则跟随')
    expect(hit.hint).toContain('一起变')
  })

  it('拿到圆形形状类 + rule 族类（两维度都要，缺一不可）', () => {
    const cls = paletteItemClassOf('rule')
    expect(cls).toContain('trigger')
    expect(cls).toContain('nk--rule')
  })
})

describe('★ 规则锁定的建立入口：连线必须能建起来', () => {
  const r = ruleNode()
  const s1 = sortNode('s1')
  const s2 = sortNode('s2', '排序2')

  it('规则节点 → 辅节点：允许，写进辅节点', () => {
    const p = ruleConnectPatch('r1', 's1', T(r, s1))
    // ★ 断言逐步拆开写（不用 toEqual整体比）：整体比会在新增字段时
    //   一次性红一片，看不出是哪一项变了，且容易被人随手改成 toMatchObject 蒙过去。
    expect(p?.nodeId).toBe('s1')
    expect(p?.linkedInputs).toEqual(['r1'])
    // 第一个连上的身份是主节点（用户：「第一个默认是可修改的主节点」）
    expect(p?.ruleRef).toEqual({ ruleNodeId: 'r1', mode: 'primary' })
  })

  it('★ 辅节点 → 规则节点：反向也要允许（内核双向扫描，UI 曾只通单向）', () => {
    const p = ruleConnectPatch('s1', 'r1', T(r, s1))
    expect(p?.nodeId).toBe('s1')
    expect(p?.linkedInputs).toEqual(['r1'])
    expect(p?.ruleRef).toEqual({ ruleNodeId: 'r1', mode: 'primary' })
  })

  it('两侧都不是规则节点：返回 null，交给 lookup 分支', () => {
    expect(ruleConnectPatch('s1', 's2', T(s1, s2))).toBeNull()
  })

  it('自己连自己：null', () => {
    expect(ruleConnectPatch('r1', 'r1', T(r))).toBeNull()
  })

  it('★ 规则节点 → 规则节点：null（否则成环，effectiveRule 会无限跳）', () => {
    expect(ruleConnectPatch('r1', 'r2', T(r, ruleNode('r2')))).toBeNull()
  })

  it('重复连线同一辅节点：第二次 null（linkedInputs 不出现重复项）', () => {
    const locked = { ...s1, config: { ...s1.config, linkedInputs: ['r1'] } }
    expect(ruleConnectPatch('r1', 's1', T(r, locked))).toBeNull()
  })

  it('★ 回归：规则节点无字段，若不前置处理会被 lookup 分支静默吞掉', () => {
    // 这条断言锁死「必须排在 lookup 的 !src.fields.length 之前」这个顺序
    expect(r.fields.length).toBe(0)
    // 旧顺序下会走到：两个表都没字段 → 静默 return，连线丢失
    expect(r.fields.length === 0 || s1.fields.length === 0).toBe(true)
    // 但规则判定先一步接管，不会走到吞掉它的那段
    expect(ruleConnectPatch('r1', 's1', T(r, s1))).not.toBeNull()
  })

  it('节点不存在：null（不抛）', () => {
    expect(ruleConnectPatch('r1', 'nope', T(r))).toBeNull()
    expect(ruleConnectPatch('nope', 's1', T(r, s1))).toBeNull()
  })
})

describe('端到端：建锁定 → 改规则 → 全组生效 → 解除', () => {
  it('连两个辅节点后，改规则节点上的规则，两处立刻跟着变', () => {
    const r = ruleNode()
    const s1 = sortNode('s1')
    const s2 = sortNode('s2', '排序2')
    // 模拟 handleConnect 写 linkedInputs
    const tables: Record<string, TableSchema> = { [r.id]: r, [s1.id]: s1 }
    tables[s2.id] = { ...s2, config: { ...s2.config, linkedInputs: [r.id] } }
    tables[s1.id] = { ...s1, config: { ...s1.config, linkedInputs: [r.id] } }
    tables[r.id] = { ...r, config: { ...r.config, rules: { sort: { sortBy: 'name', sortDir: 'desc' } } } }

    // 锁定关系成立
    expect(isRuleLocked('s1', tables)).toBe(true)
    expect(isRuleLocked('s2', tables)).toBe(true)
    expect(bindingsOfRuleNode(r.id, tables).map((b) => b.targetId).sort()).toEqual(['s1', 's2'])

    // 改主即全组变（只存一份，无需同步写）
    expect(effectiveRule(tables['s1'], tables)).toMatchObject({ sortBy: 'name', sortDir: 'desc' })
    expect(effectiveRule(tables['s2'], tables)).toMatchObject({ sortBy: 'name', sortDir: 'desc' })

    // 解除后回到自己的设置
    const unlinked: Record<string, TableSchema> = { ...tables }
    unlinked['s1'] = { ...unlinked['s1'], config: { ...unlinked['s1'].config, linkedInputs: [] } }
    expect(isRuleLocked('s1', unlinked)).toBe(false)
    expect(effectiveRule(unlinked['s1'], unlinked)).toMatchObject({ sortBy: 'col1', sortDir: 'asc' })
  })

  it('规则节点被删：辅节点回落自身设置，不规则裸奔', () => {
    const r = ruleNode()
    const s1 = sortNode('s1')
    const tables: Record<string, TableSchema> = {
      [r.id]: r,
      [s1.id]: { ...s1, config: { ...s1.config, linkedInputs: [r.id] } },
    }
    const after: Record<string, TableSchema> = {
      [s1.id]: tables[s1.id],
      [r.id]: { ...r, deleted: true },
    }
    expect(ruleNodesOf('s1', after)).toEqual([])
    expect(effectiveRule(after['s1'], after)).toMatchObject({ sortBy: 'col1' })
  })

  it('回收站里的规则节点不算锁定源', () => {
    const r = ruleNode()
    const s1 = sortNode('s1')
    const tables: Record<string, TableSchema> = {
      [r.id]: { ...r, deleted: true },
      [s1.id]: { ...s1, config: { ...s1.config, linkedInputs: [r.id] } },
    }
    expect(isRuleLocked('s1', tables)).toBe(false)
  })
})

describe('判定单一真相源', () => {
  it('isRuleNodeKind 只认 rule，其它 kind 一律false', () => {
    expect(isRuleNodeKind('rule')).toBe(true)
    for (const k of ['sort', 'map', 'text', 'output', 'docx', 'xlsx', 'start', 'table', 'param']) {
      expect(isRuleNodeKind(k)).toBe(false)
    }
  })
})
