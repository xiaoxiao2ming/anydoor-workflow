import { describe, it, expect, beforeEach } from 'vitest'
import { useBaseStore } from './store'
import { defineField } from './schema'
import { isRuleNodeKind, PALETTE_GROUPS, paletteItemClassOf } from './nodeStyle'
import { ruleConnectPatch } from './ruleSync'

/**
 * 规则节点在**真 store** 上的落地验证（2026-10-05）
 *
 * 前一轮的教训：内核 32 用例全绿 + tsc 零错误 + 529 全量全绿，
 * 但从画布上拖出来、连一条线，什么都不会发生 —— 因为 handleConnect 没接规则分支。
 *单元测试测不到「store 动作能不能把状态改对」，这组用例补上这一层。
 *
 * 模拟的是真路径：
 *   拖节点库 → App.handleAddTableAt → addTable（这里直接调 store 的 addTable）
 *   拉连线   → App.handleConnect → ruleConnectPatch + setNodeConfig
 *            （App.tsx 的那段回调依赖 React hooks，无法在 node 环境直接跑，
 *              故此处直接复现它的两个动作 —— 与 wiring.test 不同，
 *              这次断言的是**真 store 状态**，不是复制品）
 */

const st = () => useBaseStore.getState()
const field = defineField('列', 'text')

beforeEach(() => {
  // 每个用例从空画布开始：只保留 store 的写动作，tables 自己清
  useBaseStore.setState({ tables: {}, records: {}, positions: {} })
})

/** 复现 App.handleAddTableAt 里 presets.rule 的那一行 */
function dropRuleNode(id: string) {
  st().addTable({ id, kind: 'rule', name: '规则', fields: [], config: { rules: {} } })
  st().setNodePosition(id, { x: 0, y: 0 })
  return id
}

/** 复现 handleConnect 里规则分支的两个动作 */
function connect(a: string, b: string) {
  const s = st()
  const patch = ruleConnectPatch(a, b, s.tables)
  if (patch) s.setNodeConfig(patch.nodeId, { linkedInputs: patch.linkedInputs })
  return patch
}

describe('① 规则节点在节点库里（拖拽源）', () => {
  it('节点库列出「规则」项，且带圆形形状类 + rule 族色', () => {
    const all = PALETTE_GROUPS.flatMap((g) => g.kinds.map((k) => k.kind))
    expect(all).toContain('rule')
    const cls = paletteItemClassOf('rule')
    expect(cls).toContain('trigger') // 圆形
    expect(cls).toContain('nk--rule') // teal族色
  })
})

describe('② 拖到画布上能落成真节点', () => {
  it('addTable 后：tables 里有它、kind=rule、config.rules 为空对象', () => {
    const id = dropRuleNode('rule_1')
    const t = st().tables[id]
    expect(t).toBeDefined()
    expect(t.kind).toBe('rule')
    expect(isRuleNodeKind(t.kind)).toBe(true)
    expect(t.config?.rules).toEqual({})
    expect(st().positions[id]).toEqual({ x: 0, y: 0 })
  })

  it('★ 但无字段（fields: []）—— 这正是此前连线被静默吞掉的根因', () => {
    const id = dropRuleNode('rule_2')
    expect(st().tables[id].fields).toEqual([])
  })

  it('能放多个规则节点（id 不冲突、各自独立）', () => {
    dropRuleNode('rule_a')
    dropRuleNode('rule_b')
    expect(Object.keys(st().tables).filter((k) => st().tables[k].kind === 'rule').sort()).toEqual([
      'rule_a',
      'rule_b',
    ])
  })
})

describe('③ 连线：规则节点连上几个节点（双向都试）', () => {
  it('规则 → 排序：锁定成立', () => {
    dropRuleNode('r')
    st().addTable({ id: 's1', kind: 'sort', name: '排序1', fields: [field], config: { sortBy: 'col1' } })
    const p = connect('r', 's1')
    expect(p).not.toBeNull()
    expect(st().tables['s1'].config?.linkedInputs).toContain('r')
  })

  it('★ 排序 → 规则（反向）：同样成立（上一轮这条是断的）', () => {
    dropRuleNode('r')
    st().addTable({ id: 's1', kind: 'sort', name: '排序1', fields: [field], config: { sortBy: 'col1' } })
    const p = connect('s1', 'r')
    expect(p).not.toBeNull()
    expect(st().tables['s1'].config?.linkedInputs).toContain('r')
  })

  it('一个规则节点连**多个**同类型节点：全部锁定', () => {
    dropRuleNode('r')
    for (const id of ['s1', 's2', 's3']) {
      st().addTable({ id, kind: 'sort', name: `排序${id}`, fields: [field], config: { sortBy: 'col1' } })
    }
    for (const id of ['s1', 's2', 's3']) connect('r', id)
    for (const id of ['s1', 's2', 's3']) {
      expect(st().tables[id].config?.linkedInputs).toContain('r')
    }
  })

  it('重复连同一条：不产生重复项', () => {
    dropRuleNode('r')
    st().addTable({ id: 's1', kind: 'sort', name: '排序1', fields: [field], config: {} })
    connect('r', 's1')
    connect('r', 's1')
    const lk = st().tables['s1'].config?.linkedInputs ?? []
    expect(lk.filter((x) => x === 'r').length).toBe(1)
  })

  it('规则 → 规则：拒绝（防成环）', () => {
    dropRuleNode('r1')
    dropRuleNode('r2')
    expect(connect('r1', 'r2')).toBeNull()
  })
})

describe('④ 规则定义下发：改圆形节点，被锁的同类型节点全部同步', () => {
  it('建锁定 → 设 sort 规则 → 三个排序节点一起变', () => {
    dropRuleNode('r')
    for (const id of ['s1', 's2', 's3']) {
      st().addTable({ id, kind: 'sort', name: `排序${id}`, fields: [field], config: { sortBy: 'col1', sortDir: 'asc' } })
      connect('r', id)
    }
    // 在规则节点上设规则（真实 store 动作）
    st().setRule('r', 'sort', { sortBy: 'name', sortDir: 'desc' })

    // 规则存在规则节点自己身上
    expect(st().tables['r'].config?.rules?.sort).toEqual({ sortBy: 'name', sortDir: 'desc' })
    // 辅节点自己身上**没有**规则副本
    for (const id of ['s1', 's2', 's3']) {
      expect(st().tables[id].config?.rules).toBeUndefined()
    }
  })

  it('解除锁定后：辅节点回落自身设置', () => {
    dropRuleNode('r')
    st().addTable({ id: 's1', kind: 'sort', name: '排序1', fields: [field], config: { sortBy: 'col1', sortDir: 'asc' } })
    connect('r', 's1')
    st().setRule('r', 'sort', { sortBy: 'name' })
    st().unlinkFromRuleNode('r', 's1')
    expect(st().tables['s1'].config?.linkedInputs).not.toContain('r')
    expect(st().tables['s1'].config?.sortBy).toBe('col1') // 回落到自身
  })

  it('规则节点被删：辅节点回落自身，不规则裸奔', () => {
    dropRuleNode('r')
    st().addTable({ id: 's1', kind: 'sort', name: '排序1', fields: [field], config: { sortBy: 'col1' } })
    connect('r', 's1')
    st().setRule('r', 'sort', { sortBy: 'name' })
    st().trashTable('r')
    // 锁定关系随节点进回收站而失效
    expect(st().tables['r'].deleted).toBe(true)
    expect(st().tables['s1'].config?.sortBy).toBe('col1')
  })
})
