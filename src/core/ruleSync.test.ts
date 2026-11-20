// ruleSync 纯计算层单测
//
// 重点验证「同步不漏」—— 这是整个功能成立的前提。
// 若误用复制式同步，本文件在「改主节点后辅节点是否跟着变」这一组会红。

import { describe, it, expect } from 'vitest'
import {
  bindingsOfRuleNode,
  allBindings,
  ruleNodesOf,
  effectiveRule,
  isRuleLocked,
  ruleFieldsOf,
  validateRule,
  ambiguousLocks,
  RULE_NODE_KIND,
} from './ruleSync'
import type { TableSchema, NodeRuleSet } from './schema'

function t(id: string, kind: string, cfg?: TableSchema['config']): TableSchema {
  return { id, name: id, kind: kind as TableSchema['kind'], fields: [], config: cfg }
}
function tables(...list: TableSchema[]): Record<string, TableSchema> {
  const o: Record<string, TableSchema> = {}
  for (const x of list) o[x.id] = x
  return o
}

describe('bindingsOfRuleNode · 锁定范围沿连线', () => {
  it('锁定以自己为上游的节点（我挂在它下游 → 我是它的规则源）', () => {
    const ts = tables(
      t('r1', 'rule'),
      t('s1', 'sort', { linkedInputs: ['r1'] }),
      t('m1', 'map', { linkedInputs: ['r1'] }),
    )
    const binds = bindingsOfRuleNode('r1', ts)
    expect(binds.map((b) => b.targetId).sort()).toEqual(['m1', 's1'])
  })

  it('锁定把它当上游的节点（主动向下游发规则）', () => {
    const ts = tables(t('r1', 'rule'), t('out1', 'output', { linkedInputs: ['r1'] }))
    expect(bindingsOfRuleNode('r1', ts).map((b) => b.targetId)).toEqual(['out1'])
  })

  it('双向都能锁（上游下游都被锁）', () => {
    const ts = tables(
      t('r1', 'rule'),
      t('a1', 'sort', { linkedInputs: ['r1'] }),
      t('b1', 'sort', { linkedInputs: ['r1'] }),
    )
    // r1→a1；b1 的上游含 r1
    const got = bindingsOfRuleNode('r1', ts).map((b) => b.targetId).sort()
    expect(got).toEqual(['a1', 'b1'])
  })

  it('不锁自己', () => {
    const ts = tables(t('r1', 'rule', { linkedInputs: ['r1'] }))
    expect(bindingsOfRuleNode('r1', ts)).toHaveLength(0)
  })

  it('不锁回收站里的节点（与 edgesOf 一致：不产生边）', () => {
    const ts = tables(
      t('r1', 'rule'),
      { ...t('dead', 'sort', { linkedInputs: ['r1'] }), deleted: true },
    )
    expect(bindingsOfRuleNode('r1', ts)).toHaveLength(0)
  })

  it('规则节点之间不互锁（否则成环）', () => {
    const ts = tables(t('r1', 'rule'), t('r2', 'rule', { linkedInputs: ['r1'] }))
    expect(bindingsOfRuleNode('r1', ts)).toHaveLength(0)
  })

  it('非规则节点调用返回空', () => {
    const ts = tables(t('s1', 'sort'), t('a', 'sort', { linkedInputs: ['s1'] }))
    expect(bindingsOfRuleNode('s1', ts)).toHaveLength(0)
  })

  it('不存在 / 回收站的规则节点返回空', () => {
    expect(bindingsOfRuleNode('nope', tables(t('a', 'sort')))).toHaveLength(0)
    expect(bindingsOfRuleNode('r', tables({ ...t('r', 'rule'), deleted: true }))).toHaveLength(0)
  })

  it('同一节点只被记一次（去重）', () => {
    // linkedInputs 有重复项时不能产生两条锁定
    const ts = tables(t('r1', 'rule'), t('s1', 'sort', { linkedInputs: ['r1', 'r1'] }))
    expect(bindingsOfRuleNode('r1', ts)).toHaveLength(1)
  })
})

describe('effectiveRule · 同步的核心断言', () => {
  const rule: NodeRuleSet = { family: 'sort', scale: 1.5 }

  it('★ 改规则节点 → 全部辅节点立刻跟着变（只存一份，无需同步写）', () => {
    const ts = tables(
      t('r1', 'rule', { rules: { sort: { family: 'sort', scale: 1.2 } } }),
      t('s1', 'sort', { linkedInputs: ['r1'] }),
      t('s2', 'sort', { linkedInputs: ['r1'] }),
    )
    // 初始scale = 1.2
    expect(effectiveRule(ts.s1, ts)?.scale).toBe(1.2)
    expect(effectiveRule(ts.s2, ts)?.scale).toBe(1.2)

    // ★ 只改一处，两处同时变 —— 这就是引用式存储的意义
    const ts2 = tables(
      t('r1', 'rule', { rules: { sort: { family: 'sort', scale: 1.8 } } }),
      t('s1', 'sort', { linkedInputs: ['r1'] }),
      t('s2', 'sort', { linkedInputs: ['r1'] }),
    )
    expect(effectiveRule(ts2.s1, ts2)?.scale).toBe(1.8)
    expect(effectiveRule(ts2.s2, ts2)?.scale).toBe(1.8)
  })

  it('按 kind 取对应桶：sort 节点不吃 map 的规则', () => {
    const ts = tables(
      t('r1', 'rule', { rules: { sort: { scale: 2 }, map: { scale: 0.5 } } }),
      t('s1', 'sort', { linkedInputs: ['r1'] }),
      t('m1', 'map', { linkedInputs: ['r1'] }),
    )
    expect(effectiveRule(ts.s1, ts)?.scale).toBe(2)
    expect(effectiveRule(ts.m1, ts)?.scale).toBe(0.5)
  })

  it('规则项未覆盖时回落自身 config（不是变成 undefined）', () => {
    const ts = tables(
      t('r1', 'rule', { rules: { sort: { family: 'map' } } }), // 只给 family
      t('s1', 'sort', { linkedInputs: ['r1'], sortBy: 'design_flow' } as TableSchema['config']),
    )
    const eff = effectiveRule(ts.s1, ts)
    expect(eff?.family).toBe('map') // 来自规则
    expect((eff as { sortBy?: string })?.sortBy).toBe('design_flow') // 回落自身
  })

  it('规则节点被删 → 回落自身 config，不规则裸奔', () => {
    const ts = tables(
      { ...t('r1', 'rule'), deleted: true },
      t('s1', 'sort', { linkedInputs: ['r1'], family: 'text' } as TableSchema['config']),
    )
    const eff = effectiveRule(ts.s1, ts)
    expect(eff?.family).toBe('text')
  })

  it('未被锁定的节点 → 返回自身 config', () => {
    const ts = tables(t('s1', 'sort', { family: 'text' } as TableSchema['config']))
    expect(effectiveRule(ts.s1, ts)?.family).toBe('text')
  })

  it('规则桶缺该 kind → 回落自身 config', () => {
    const ts = tables(
      t('r1', 'rule', { rules: { map: { scale: 2 } } }), // 只有 map 桶
      t('s1', 'sort', { linkedInputs: ['r1'], family: 'text' } as TableSchema['config']),
    )
    expect(effectiveRule(ts.s1, ts)?.family).toBe('text')
  })

  it('两处读的结果严格一致（同一节点两次读不能不同）', () => {
    const ts = tables(
      t('r1', 'rule', { rules: { sort: { scale: 1.3 } } }),
      t('s1', 'sort', { linkedInputs: ['r1'] }),
    )
    expect(effectiveRule(ts.s1, ts)).toEqual(effectiveRule(ts.s1, ts))
  })
})

describe('isRuleLocked · 只读判定与生效判定同源', () => {
  it('被锁 → true，且与 effectiveRule 口径一致', () => {
    const ts = tables(
      t('r1', 'rule', { rules: { sort: { scale: 1.5 } } }),
      t('s1', 'sort', { linkedInputs: ['r1'] }),
    )
    expect(isRuleLocked('s1', ts)).toBe(true)
    // 判定为锁定的节点，读到的必须是规则节点的值而非自身
    expect(effectiveRule(ts.s1, ts)?.scale).toBe(1.5)
  })

  it('未被锁 → false', () => {
    expect(isRuleLocked('s1', tables(t('s1', 'sort')))).toBe(false)
  })
})

describe('ruleNodesOf / allBindings', () => {
  it('反查锁定来源', () => {
    const ts = tables(t('r1', 'rule'), t('s1', 'sort', { linkedInputs: ['r1'] }))
    expect(ruleNodesOf('s1', ts)).toEqual(['r1'])
  })

  it('全部关系一次算出', () => {
    const ts = tables(
      t('r1', 'rule'),
      t('r2', 'rule'),
      t('s1', 'sort', { linkedInputs: ['r1'] }),
      t('s2', 'map', { linkedInputs: ['r2'] }),
    )
    const m = allBindings(ts)
    expect([...m.keys()].sort()).toEqual(['r1', 'r2'])
    expect(m.get('r1')?.[0].targetId).toBe('s1')
  })

  it('回收站的规则节点不进allBindings', () => {
    const ts = tables({ ...t('r1', 'rule'), deleted: true }, t('s1', 'sort', { linkedInputs: ['r1'] }))
    expect(allBindings(ts).size).toBe(0)
  })
})

describe('validateRule · 规则项缺失要能报出来', () => {
  it('map 缺匹配字段 → 报两条', () => {
    const p = validateRule({}, 'map')
    expect(p.map((x) => x.field)).toContain('mapKeyField')
    expect(p.map((x) => x.field)).toContain('mapMatchField')
  })

  it('sort 缺排序列 → 报', () => {
    expect(validateRule({}, 'sort').map((x) => x.field)).toContain('sortBy')
  })

  it('完整规则 → 无问题', () => {
    expect(validateRule({ mapKeyField: 'a', mapMatchField: 'b' }, 'map')).toHaveLength(0)
  })

  it('缩放越界 → 报', () => {
    expect(validateRule({ scale: 9 }, 'text').map((x) => x.field)).toContain('scale')
    expect(validateRule({ scale: 0.1 }, 'text').map((x) => x.field)).toContain('scale')
  })

  it('未定义的规则项一律放行（不误报）', () => {
    expect(validateRule({}, 'table')).toHaveLength(0)
  })
})

describe('ambiguousLocks · 锁定歧义要能看见', () => {
  it('两个规则节点锁同一节点 → 报歧义', () => {
    const ts = tables(
      t('r1', 'rule'),
      t('r2', 'rule'),
      t('s1', 'sort', { linkedInputs: ['r1', 'r2'] }),
    )
    const amb = ambiguousLocks(ts)
    expect(amb).toHaveLength(1)
    expect(amb[0].targetId).toBe('s1')
    expect(amb[0].ruleNodeIds.sort()).toEqual(['r1', 'r2'])
  })

  it('无歧义 → 空', () => {
    const ts = tables(
      t('r1', 'rule'),
      t('s1', 'sort', { linkedInputs: ['r1'] }),
      t('s2', 'sort', { linkedInputs: ['r1'] }),
    )
    expect(ambiguousLocks(ts)).toHaveLength(0)
  })
})

describe('规则项按 kind 适配（用户：针对不同内容适配不同规则样式体系）', () => {
  it('不同 kind 的专属规则项不同', () => {
    expect(ruleFieldsOf('sort')).toContain('sortBy')
    expect(ruleFieldsOf('sort')).not.toContain('mapKeyField')
    expect(ruleFieldsOf('map')).toContain('mapKeyField')
    expect(ruleFieldsOf('map')).not.toContain('sortBy')
  })

  it('所有 kind 都有通用视觉规则', () => {
    for (const k of ['sort', 'map', 'text', 'table', 'output']) {
      expect(ruleFieldsOf(k)).toEqual(expect.arrayContaining(['family', 'shape', 'scale']))
    }
  })

  it('未知 kind 只给通用视觉，不报错', () => {
    expect(ruleFieldsOf('brandnew')).toEqual(['family', 'shape', 'scale'])
  })
})

describe('常量', () => {
  it('RULE_NODE_KIND 为 rule', () => {
    expect(RULE_NODE_KIND).toBe('rule')
  })
})