import { describe, it, expect, beforeEach } from 'vitest'
import { useBaseStore } from './store'
import { defineField } from './schema'
import { KIND_STYLE } from './nodeStyle'
import {
  ambiguousLocks,
  bindingsOfRuleNode,
  effectiveRule,
  identityOwnerOf,
  isRuleLocked,
  memberCountOf,
  memberSummaryOf,
  membersOf,
  primaryOf,
  roleModeOf,
  ruleConnectPatch,
} from './ruleSync'

/**
 * 规则跟随节点的**主从 / 跟随**语义（2026-10-05）
 *
 * 用户口径逐条对应：
 *   「规则节点改名规则跟随」
 *   「规则节点第一个默认是可修改的主节点。连接剩下的是从节点。」
 *   「规则节点上显示几个节点。且双击能修改节点的主次、跟随不跟随。」
 *   （双击面板）→「切换跟随/不跟随, 改主次顺序, 移出该组」
 *
 * ── 这组用例的判据（沿用 gate-selfcheck-real-capture）──
 *   · 全部走**真 store 动作 + 真内核函数**，不抄逻辑。
 *   · connect() 复现 App.handleConnect 的规则分支原文（同一段 patch 写法），
 *     所以改坏真实现这些用例会变红 —— 已在提交前做过反向验证，见下。
 *   ·★ 判锁定范围与主次时**不看**自己写的 mode，一律回 membersOf / primaryOf 重新算，
 *     避免「用写入值当读取断言」的自证。
 */

const st = () => useBaseStore.getState()
const field = defineField('列', 'text')

beforeEach(() => {
  useBaseStore.setState({ tables: {}, records: {}, positions: {} })
})

function dropRuleNode(id: string) {
  st().addTable({ id, kind: 'rule', name: '规则跟随', fields: [], config: { rules: {} } })
  return id
}

/** ★ 与 App.tsx handleConnect 规则分支**逐字同构**的连线动作 */
function connect(a: string, b: string) {
  const s = st()
  const patch = ruleConnectPatch(a, b, s.tables)
  if (patch) {
    s.setNodeConfig(patch.nodeId, {
      linkedInputs: patch.linkedInputs,
      // 与 App 同构：undefined 时不写（否则浅合并抹掉已有身份）
      ...(patch.ruleRef ? { ruleRef: patch.ruleRef } : {}),
    })
  }
  return patch
}

function addSort(id: string, name = '排序', cfg: Record<string, unknown> = {}) {
  st().addTable({ id, kind: 'sort', name, fields: [field], config: { sortBy: 'col1', ...cfg } })
  return id
}

describe('① 改名：节点库与提示里都是「规则跟随」', () => {
  it('kind 规则的 label 是「规则跟随」', () => {
    expect(KIND_STYLE.rule.label).toBe('规则跟随')
  })
})

describe('② 第一个连上的自动是主节点，其余是从节点', () => {
  it('★ 连第一个 → 它是 primary', () => {
    dropRuleNode('r')
    addSort('s1')
    connect('r', 's1')
    expect(roleModeOf('s1', 'r', st().tables)).toBe('primary')
    expect(primaryOf('r', st().tables)?.targetId).toBe('s1')
  })

  it('★ 连第二个 → follow，且主节点仍是第一个（primaryOf 重新算，不看写入值）', () => {
    dropRuleNode('r')
    addSort('s1')
    addSort('s2')
    connect('r', 's1')
    connect('r', 's2')
    expect(roleModeOf('s2', 'r', st().tables)).toBe('follow')
    expect(primaryOf('r', st().tables)?.targetId).toBe('s1')
    // 组内 primary 恒有且仅有一个
    expect(membersOf('r', st().tables).filter((m) => m.mode === 'primary').length).toBe(1)
  })

  it('反向连线（成员 → 规则）也认得出身份', () => {
    dropRuleNode('r')
    addSort('s1')
    addSort('s2')
    connect('s1', 'r')
    connect('s2', 'r')
    expect(roleModeOf('s1', 'r', st().tables)).toBe('primary')
    expect(roleModeOf('s2', 'r', st().tables)).toBe('follow')
  })

  it('membersOf 把主节点排在第一个（「第一个默认是可修改的主节点」）', () => {
    dropRuleNode('r')
    addSort('s1')
    addSort('s2')
    connect('r', 's1')
    connect('r', 's2')
    expect(membersOf('r', st().tables).map((m) => m.targetId)).toEqual(['s1', 's2'])
  })
})

describe('③ 改主次：设主后原主自动降为从', () => {
  it('★ setRulePrimary：目标转 primary，原 primary 降 follow', () => {
    dropRuleNode('r')
    addSort('s1')
    addSort('s2')
    connect('r', 's1')
    connect('r', 's2')
    st().setRulePrimary('r', 's2')
    expect(roleModeOf('s2', 'r', st().tables)).toBe('primary')
    expect(roleModeOf('s1', 'r', st().tables)).toBe('follow')
    expect(primaryOf('r', st().tables)?.targetId).toBe('s2')
  })

  it('设主后 membersOf 顺序也变了（主在第一行）', () => {
    dropRuleNode('r')
    addSort('s1')
    addSort('s2')
    connect('r', 's1')
    connect('r', 's2')
    st().setRulePrimary('r', 's2')
    expect(membersOf('r', st().tables).map((m) => m.targetId)).toEqual(['s2', 's1'])
  })

  it('「不跟随」的成员被设主后 → 转primary（不是保持 ignore）', () => {
    dropRuleNode('r')
    addSort('s1')
    addSort('s2')
    connect('r', 's1')
    connect('r', 's2')
    st().setRuleFollow('r', 's2', false)
    st().setRulePrimary('r', 's2')
    expect(roleModeOf('s2', 'r', st().tables)).toBe('primary')
  })

  it('非成员设主：拒绝（不能凭空造出锁定关系）', () => {
    dropRuleNode('r')
    addSort('s1')
    addSort('x')
    connect('r', 's1')
    st().setRulePrimary('r', 'x')
    expect(roleModeOf('x', 'r', st().tables)).toBeNull()
    expect(primaryOf('r', st().tables)?.targetId).toBe('s1')
  })
})

describe('④ 跟随 / 不跟随：不跟随的节点不受该规则影响', () => {
  /*★ 三个成员：s1 主（primary）、s2 / s3 从（follow）。
     测「不跟随」必须用**从节点**：主节点退跟随时会被 store 降级并让位
     （见 ⑤），拿主节点来测这条会得到「规则仍生效」而误判成功能坏了。*/
  function threeMembers() {
    dropRuleNode('r')
    for (const id of ['s1', 's2', 's3']) {
      addSort(id, `排序${id.slice(1)}`, { sortBy: 'col1', sortDir: 'asc' })
      connect('r', id)
    }
  }

  it('★ 从节点切到不跟随 → effectiveRule 回落到自身设置', () => {
    threeMembers()
    st().setRule('r', 'sort', { sortBy: 'name', sortDir: 'desc' })
    expect(effectiveRule(st().tables['s2'], st().tables)).toMatchObject({ sortBy: 'name', sortDir: 'desc' })

    st().setRuleFollow('r', 's2', false)
    // 回落到 s2 自己的设置
    expect(roleModeOf('s2', 'r', st().tables)).toBe('ignore')
    expect(effectiveRule(st().tables['s2'], st().tables)).toMatchObject({ sortBy: 'col1', sortDir: 'asc' })
    // ★ 组里其它人不受影响（不能一个不跟随就全员失效）
    expect(effectiveRule(st().tables['s3'], st().tables)).toMatchObject({ sortBy: 'name' })
  })

  it('★ 但连线还在：isRuleLocked 仍为 true（那是「连着」的事实）', () => {
    threeMembers()
    st().setRuleFollow('r', 's2', false)
    expect(isRuleLocked('s2', st().tables)).toBe(true)
    expect(memberCountOf('r', st().tables)).toBe(3)
  })

  it('切回跟随 → 规则重新生效', () => {
    threeMembers()
    st().setRule('r', 'sort', { sortBy: 'name' })
    st().setRuleFollow('r', 's2', false)
    st().setRuleFollow('r', 's2', true)
    expect(roleModeOf('s2', 'r', st().tables)).toBe('follow')
    expect(effectiveRule(st().tables['s2'], st().tables)).toMatchObject({ sortBy: 'name' })
  })

  it('对已是 follow 的节点「跟随」= 保持 follow，不会被误升为 primary', () => {
    threeMembers()
    st().setRuleFollow('r', 's2', true)
    expect(roleModeOf('s2', 'r', st().tables)).toBe('follow')
    expect(primaryOf('r', st().tables)?.targetId).toBe('s1')
  })
})

describe('⑤ 主节点不能退跟随（组内恒有且仅有一个 primary）', () => {
  it('★ 组里有两个：主节点退跟随 → 降为 follow，主位让给下一个', () => {
    dropRuleNode('r')
    addSort('s1')
    addSort('s2')
    connect('r', 's1')
    connect('r', 's2')
    st().setRuleFollow('r', 's1', false)
    // s1 降为 follow（不是 ignore），主位让给 s2
    expect(roleModeOf('s1', 'r', st().tables)).toBe('follow')
    expect(roleModeOf('s2', 'r', st().tables)).toBe('primary')
    expect(membersOf('r', st().tables).filter((m) => m.mode === 'primary').length).toBe(1)
  })

  it('★ 组里只有它一个：拒绝退跟随（否则组变成无主，规则改了没人执行）', () => {
    dropRuleNode('r')
    addSort('s1')
    connect('r', 's1')
    st().setRuleFollow('r', 's1', false)
    expect(roleModeOf('s1', 'r', st().tables)).toBe('primary')
  })
})

describe('⑥ 移出该组（unlink）会清掉身份，不留悬空ruleRef', () => {
  it('★ unlink 后 ruleRef 整个没了（不是只改 mode）', () => {
    dropRuleNode('r')
    addSort('s1')
    connect('r', 's1')
    expect(st().tables['s1'].config?.ruleRef).toBeDefined()
    st().unlinkFromRuleNode('r', 's1')
    expect(st().tables['s1'].config?.ruleRef).toBeUndefined()
    expect(memberCountOf('r', st().tables)).toBe(0)
  })

  it('移出的是 primary：下一位自动接任主节点', () => {
    dropRuleNode('r')
    addSort('s1')
    addSort('s2')
    connect('r', 's1')
    connect('r', 's2')
    st().unlinkFromRuleNode('r', 's1')
    expect(primaryOf('r', st().tables)?.targetId).toBe('s2')
  })

  /*★ 这条是上一条「为什么还不够」的补丁。
     primaryOf 走 membersOf 的兜底，所以哪怕 store 没做让位，
     「组里还剩一个成员」时兜底也会让它当主节点 —— 上一条照样绿。
     要真正锁住 store 的让位逻辑，必须**绕过兜底**直接断言存下来的身份：
     移出主节点后，剩下的成员身上应当**真的**被写上了 mode:'primary'。
     （反向验证：把 unlinkFromRuleNode 里的让位分支关掉，只有这一条会红。）*/
  it('★ 让位是**真写进 store**的，不靠 membersOf 兜底蒙混', () => {
    dropRuleNode('r')
    addSort('s1')
    addSort('s2')
    connect('r', 's1')
    connect('r', 's2')
    expect(st().tables['s2'].config?.ruleRef?.mode).toBe('follow') // 事前：s2 是从
    st().unlinkFromRuleNode('r', 's1')
    // 直接读存储值，绕开 membersOf 的兜底
    expect(st().tables['s2'].config?.ruleRef).toEqual({ ruleNodeId: 'r', mode: 'primary' })
  })

  it('移出主节点时若只剩「不跟随」的成员，不强行提升它（它明确不跟随）', () => {
    dropRuleNode('r')
    addSort('s1')
    addSort('s2')
    connect('r', 's1')
    connect('r', 's2')
    st().setRuleFollow('r', 's2', false)
    st().unlinkFromRuleNode('r', 's1')
    expect(st().tables['s2'].config?.ruleRef).toEqual({ ruleNodeId: 'r', mode: 'ignore' })
    // 组里只剩一个不跟随的成员 → 组无主，membersOf 兜底也不提（它挑的是非 ignore）
    expect(primaryOf('r', st().tables)).toBeNull()
  })
})

describe('⑦ 圆形节点上的计数（「规则节点上显示几个节点」）', () => {
  it('memberCountOf / memberSummaryOf：连 3 个显示 3，主节点名可取', () => {
    dropRuleNode('r')
    for (const id of ['s1', 's2', 's3']) {
      addSort(id, `排序${id.slice(1)}`)
      connect('r', id)
    }
    expect(memberCountOf('r', st().tables)).toBe(3)
    const s = memberSummaryOf('r', st().tables)
    expect(s.count).toBe(3)
    expect(s.primaryName).toBe('排序1')
  })

  it('一个成员都没连：count 0、primaryName null（画布上显示「—」）', () => {
    dropRuleNode('r')
    expect(memberSummaryOf('r', st().tables)).toEqual({ count: 0, primaryName: null })
  })

  it('「不跟随」的成员仍计入计数（它确实连着，只是规则不生效）', () => {
    dropRuleNode('r')
    addSort('s1')
    addSort('s2')
    connect('r', 's1')
    connect('r', 's2')
    st().setRuleFollow('r', 's2', false)
    expect(memberCountOf('r', st().tables)).toBe(2)
    expect(membersOf('r', st().tables).length).toBe(2)
  })
})

describe('⑧ 向后兼容：早期连线没有 mode 字段', () => {
  it('★ 缺 ruleRef → 规则**仍然生效**（关键不变量：不是「不生效」）', () => {
    dropRuleNode('r')
    addSort('s1', '排序1', { sortBy: 'col1' })
    // 手工造一条「只有连线、没有 ruleRef」的老数据
    st().setNodeConfig('s1', { linkedInputs: ['r'] })
    st().setRule('r', 'sort', { sortBy: 'name' })
    // ★ 断言 effectiveRule 而不是 roleModeOf：缺省是 follow（而非 ignore），
    //   但组内无显式 primary 时 membersOf 会兜底接任，故 roleModeOf 可能是 primary。
    //   真正要锁死的不变量是「升级后规则没失效」。
    expect(effectiveRule(st().tables['s1'], st().tables)).toMatchObject({ sortBy: 'name' })
    expect(roleModeOf('s1', 'r', st().tables)).not.toBe('ignore')
  })

  it('有 ruleRef 但 mode 缺失 → 同样不是 ignore，规则仍生效', () => {
    dropRuleNode('r')
    addSort('s1', '排序1', { sortBy: 'col1' })
    st().setNodeConfig('s1', { linkedInputs: ['r'], ruleRef: { ruleNodeId: 'r' } })
    st().setRule('r', 'sort', { sortBy: 'name' })
    expect(roleModeOf('s1', 'r', st().tables)).not.toBe('ignore')
    expect(effectiveRule(st().tables['s1'], st().tables)).toMatchObject({ sortBy: 'name' })
  })

  it('★ 组内无显式 primary 时兜底接任（老数据没 mode，圆形节点不能显示「无主」）', () => {
    dropRuleNode('r')
    addSort('s1', '排序1')
    addSort('s2', '排序2')
    for (const id of ['s1', 's2']) st().setNodeConfig(id, { linkedInputs: ['r'] })
    const ms = membersOf('r', st().tables)
    expect(ms.map((m) => m.targetId)).toEqual(['s1', 's2'])
    expect(ms.filter((m) => m.mode === 'primary').length).toBe(1)
    expect(primaryOf('r', st().tables)?.targetId).toBe('s1')
  })
})

describe('⑨ 两个规则组锁同一个节点：不能串组、不能抢身份、不能跨组改', () => {
  it('★ A 组的主节点身份不会被 B 组抢走（ruleRef 单槽，后连者不覆盖）', () => {
    dropRuleNode('rA')
    dropRuleNode('rB')
    addSort('s1')
    connect('rA', 's1') // s1 在 A 组是 primary
    connect('rB', 's1') // 再连 B 组：不能覆盖 A 组的身份
    expect(roleModeOf('s1', 'rA', st().tables)).toBe('primary')
    expect(primaryOf('rA', st().tables)?.targetId).toBe('s1')
    expect(identityOwnerOf('s1', st().tables, 'rA')).toBeNull() // A 组仍可改它
    expect(identityOwnerOf('s1', st().tables, 'rB')).toBe('rA') // B 组改不了
  })

  it('★ 组内有成员却无主节点 —— 这就是抢身份的症状，本断言直接锁死它', () => {
    dropRuleNode('rA')
    dropRuleNode('rB')
    addSort('s1')
    addSort('s2')
    connect('rA', 's1')
    connect('rA', 's2')
    connect('rB', 's1') // s1 又进 B 组，但不能把 A 组的 primary 抢走
    const aMembers = membersOf('rA', st().tables)
    expect(aMembers.map((m) => m.targetId)).toEqual(['s1', 's2'])
    expect(aMembers.filter((m) => m.mode === 'primary').length).toBe(1)
    expect(aMembers.find((m) => m.targetId === 's1')?.mode).toBe('primary')
  })

  it('★ B 组不能把「身份属于 A 组」的节点设为不跟随（跨组写必须被拒）', () => {
    dropRuleNode('rA')
    dropRuleNode('rB')
    addSort('s1')
    connect('rA', 's1') // s1 身份归 A 组
    connect('rB', 's1')
    st().setRuleFollow('rB', 's1', false) // 应被拒
    expect(identityOwnerOf('s1', st().tables, 'rA')).toBeNull()
    expect(roleModeOf('s1', 'rA', st().tables)).toBe('primary')
  })

  it('★ B 组把 s1 设为「主节点」也要被拒（同一道守卫，两个写动作都得挡）', () => {
    dropRuleNode('rA')
    dropRuleNode('rB')
    addSort('s1')
    connect('rA', 's1')
    connect('rB', 's1')
    st().setRulePrimary('rB', 's1') // 应被拒
    expect(roleModeOf('s1', 'rA', st().tables)).toBe('primary')
  })

  it('B 组里身份未被占用的成员才能当主节点', () => {
    dropRuleNode('rA')
    dropRuleNode('rB')
    addSort('s1')
    addSort('s2')
    connect('rA', 's1') // s1 身份 → A
    connect('rB', 's1') // s1 进 B 组（身份仍归 A）
    connect('rB', 's2') // s2 进 B 组，身份槽空 → 由 B 组写入
    expect(roleModeOf('s2', 'rB', st().tables)).toBe('primary')
    expect(primaryOf('rB', st().tables)?.targetId).toBe('s2')
    expect(roleModeOf('s1', 'rA', st().tables)).toBe('primary')
  })

  it('ambiguousLocks 能看见这种重叠（设计歧义要可见，不是崩溃）', () => {
    dropRuleNode('rA')
    dropRuleNode('rB')
    addSort('s1')
    connect('rA', 's1')
    connect('rB', 's1')
    expect(ambiguousLocks(st().tables)).toEqual([{ targetId: 's1', ruleNodeIds: ['rA', 'rB'] }])
  })
})
