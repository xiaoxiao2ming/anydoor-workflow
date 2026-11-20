// ════════════════════════════════════════════════════════════════════════
// 连接口角色 —— 「规则线只能连接常规节点顶部的点」
// ════════════════════════════════════════════════════════════════════════
// 用户 2026-10-05 原话：「规则线只能连接常规节点顶部的点」
//
// ── 这组测试为什么必须存在（一次真实踩坑的记录）────────────────────
//上一轮 587 个用例全绿，规则连线功能却带着一个真缺陷上线：
//   规则节点下方口能连到普通节点的**左侧数据流口**，且被记成数据流线。
// 之所以全绿：ruleNode.wiring.test.ts 里那条「反向也要允许」的用例调的是
//   **纯函数 ruleConnectPatch**，而那个函数**根本看不到 handle**——
//   它只拿到两个节点 id。「哪个口挨着哪个口」这件事只有 UI 层知道，
//   而 UI 层零覆盖。本组测试把这个缺口补上：
//   ① 内核层：canConnectPorts 的准入矩阵（快，秒回）
//   ② UI 层：真浏览器真拖拽，证明 handle 的 id 真的挂上了、真的挡住了
// ════════════════════════════════════════════════════════════════════════

import { describe, it, expect } from 'vitest'
import {
  canConnectPorts,
  isRulePortId,
  canJoinRuleGroup,
  RULE_PORT_IN,
  RULE_PORT_IN_2,
  RULE_PORT_OUT,
  RULE_PORT_OUT_2,
} from './ruleSync'
import { defineNode, defineField, type TableSchema } from './schema'

/** 第6 条用例要铺真实节点，kind 参与判定 —— 不能用假对象糊弄 */
const f = defineField('列', 'text')
const t2 = (...ts: TableSchema[]) => Object.fromEntries(ts.map((t) => [t.id, t]))

describe('连接口角色：isRulePortId', () => {
  it('四个规则口 id 都认（上1/上2/下1/下2，2026-10-05 拆 id 后扩到四个）', () => {
    expect(isRulePortId(RULE_PORT_IN)).toBe(true)
    expect(isRulePortId(RULE_PORT_IN_2)).toBe(true)
    expect(isRulePortId(RULE_PORT_OUT)).toBe(true)
    expect(isRulePortId(RULE_PORT_OUT_2)).toBe(true)
  })

  it('★ 数据流口（id 为 null）必须判false —— 这是本次修复的根据', () => {
    // 普通节点的左右数据流口在 Canvas.tsx 里没挂 id，React Flow 传回来就是 null。
    // 若这里错判 true，「规则口↔数据流口」就会被放行，整个修复失效。
    expect(isRulePortId(null)).toBe(false)
    expect(isRulePortId(undefined)).toBe(false)
    expect(isRulePortId('')).toBe(false)
  })

  it('不把任意字符串当规则口（防止 id 拼错时静默放行）', () => {
    // ★ 历史注意：'rule-in-2' 在 2026-10-05 拆上2 id 后成了**正式 id**（转正），
    //   这组「防拼错」哨兵因此换成了尚不存在的编号 —— 若将来下3/上3 落地，
    //   这里要跟着再换，别让哨兵反过来咬合法 id。
    expect(isRulePortId('rule-in-3')).toBe(false)
    expect(isRulePortId('rule-out-3')).toBe(false)
    expect(isRulePortId('Rule-In')).toBe(false)
    expect(isRulePortId('rule_out')).toBe(false)
  })
})

describe('连接口角色：canConnectPorts（用户口径：规则线只能连常规节点顶部的点）', () => {
  it('规则口↔规则口：放行（正向：规则节点下方口 → 普通节点顶部口）', () => {
    expect(canConnectPorts('r1', 's1', RULE_PORT_OUT, RULE_PORT_IN)).toBe(true)
  })

  it('规则口↔规则口：放行（反向：普通节点顶部口 → 规则节点下方口）', () => {
    // 这一条是用户「只能连接上方的点」那句话里另一半诉求：
    // 若这条false，用户想从上方口反向拖下来就会连不上。
    expect(canConnectPorts('s1', 'r1', RULE_PORT_IN, RULE_PORT_OUT)).toBe(true)
  })

  it('★ 规则节点下方口 → 普通节点【左侧数据流口】：必须拒绝（本次修的真缺陷）', () => {
    // 修之前这里是 true：React Flow Strict 只禁同类型对同类型，
    // 规则口与左侧口都是 target，异类型判定下两者可连，
    // 结果 App.handleConnect 把这条线写进 inputs/linkedInputs，
    // 规则节点被当成「输出节点的上游数据源」。
    expect(canConnectPorts('r1', 's1', RULE_PORT_OUT, null)).toBe(false)
  })

  it('★ 反向：普通节点【左侧数据流口】→ 规则节点下方口：同样必须拒绝', () => {
    // 同上，反向也是真缺陷（源若是「开始」节点，会直接把数据流线写进规则节点）
    expect(canConnectPorts('s1', 'r1', null, RULE_PORT_OUT)).toBe(false)
  })

  it('规则节点下方口 → 普通节点【右侧数据流口】：拒绝', () => {
    expect(canConnectPorts('r1', 's1', RULE_PORT_OUT, null)).toBe(false)
  })

  it('数据流口↔数据流口：放行（原有数据流连线一行都不能被影响）', () => {
    // ★ 这条是回归底线：本次只在UI 层开口子，不能顺手把数据流连线堵死。
    // 之前 587 个用例里大量用例依赖数据流连线，若这里false 会一片红。
    expect(canConnectPorts('a1', 'a2', null, null)).toBe(true)
    expect(canConnectPorts('a1', 'a2', undefined, undefined)).toBe(true)
  })

  it('不传 tables：只校验口角色，同类型判定不生效（向后兼容旧调用）', () => {
    // ★ 这条锁住「参数可选」这个设计决定。若哪天有人把 tables 改成必填，
    //   下面那条「数据流口↔数据流口」四参用例会集体变成 undefined，
    //   静默降级成不校验 —— 这条就是防那个坑的哨兵。
    expect(canConnectPorts('r1', 's1', RULE_PORT_OUT, RULE_PORT_IN)).toBe(true)
  })
})

// ════════════════════════════════════════════════════════════════════════
// 拆下1/下2 后的新防线（2026-10-05）：规则口互连必须落在规则节点上
// ════════════════════════════════════════════════════════════════════════
/**
 * ★★ 为什么这组用例必须存在：
 *   常规节点有了底部规则口（下1/下2）之后，「常规节点下1 → 另一常规节点上1」
 *   第一次成为可拖出的连线。canJoinRuleGroup 对「两端都非规则节点」是放行的
 *   （那是它管数据流的合法语义，有专门用例锁着，不能改），
 *   于是这条线会穿透到 ruleConnectPatch —— 它判 null，连线落进**数据流分支**：
 *   从规则口拖出来的线被记成数据流线，口角色与落库语义劈叉。
 *   防线加在 canConnectPorts（UI 准入层），这里就是证明它真拦住的哨兵。
 */
describe('新防线：两端规则口但两端都不是规则节点 → 拒绝', () => {
  // ★ 辅助必须在**本 describe 内**自带一份：第 6 条 describe 里的 rule/tbl
  //   是它回调的局部变量，跨块引用是 ReferenceError（别图省事提出来共用——
  //   那会改动第 6 条块的既有结构，回归面无谓扩大）。
  const rule0 = defineNode('rule', '规则', [], { rules: {} }, 'r1')
  const tbl0 = (id: string) => defineNode('table', id, [f], undefined, id)

  it('★ 常规节点下1 → 常规节点上1：拒绝（否则会劈叉成数据流线）', () => {
    const T = t2(tbl0('s1'), tbl0('s2'))
    expect(canConnectPorts('s1', 's2', RULE_PORT_OUT, RULE_PORT_IN, T)).toBe(false)
  })

  it('★ 常规节点下2 → 常规节点上2：同样拒绝（四个口两两组合都拦）', () => {
    const T = t2(tbl0('s1'), tbl0('s2'))
    expect(canConnectPorts('s1', 's2', RULE_PORT_OUT_2, RULE_PORT_IN_2, T)).toBe(false)
    expect(canConnectPorts('s1', 's2', RULE_PORT_OUT, RULE_PORT_IN_2, T)).toBe(false)
    expect(canConnectPorts('s1', 's2', RULE_PORT_OUT_2, RULE_PORT_IN, T)).toBe(false)
  })

  it('规则节点下1 → 常规节点上2：放行（正经规则线，落点已统一到上2）', () => {
    const T = t2(rule0, tbl0('tb1'))
    expect(canConnectPorts('r1', 'tb1', RULE_PORT_OUT, RULE_PORT_IN_2, T)).toBe(true)
  })

  it('两端查不到节点（拖到空白处）：不拦，交由落库层兜底', () => {
    expect(canConnectPorts('r1', 'ghost', RULE_PORT_OUT, RULE_PORT_IN, t2(rule0))).toBe(true)
  })

  it('不传 tables 时防线不生效（四参旧调用语义不变）', () => {
    expect(canConnectPorts('s1', 's2', RULE_PORT_OUT, RULE_PORT_IN)).toBe(true)
  })
})

// ════════════════════════════════════════════════════════════════════════
// 第 6 条：规则跟随节点只能连接同种类型的节点
// ════════════════════════════════════════════════════════════════════════
/**
 * ★★ 这里记一次真实的判据错误，反向对照必看：
 *
 *   初版把第 6 条实现成 `src.kind !== tgt.kind → 拒绝`。
 *   跑起来才发现**整个规则跟随功能一条线都连不上**。
 *   根因：比较对象搞反了。规则节点自己的 kind 恒为 'rule'
 *   （isRuleNodeKind 就是 kind === 'rule'），被锁节点是 'table'/'sort'/...
 *   两边**永不相等** → 全部拒绝。
 *
 *   正确语义（从 effectiveRule 反推出来）：规则按
 *   `ruleNode.config.rules[kind]` 分桶下发，所以「同种类型」=同一个规则节点
 *   锁定的**成员之间**必须同 kind。规则节点自己不在这个比较里。
 *
 *   这组用例就是那道哨兵：若判据再被改反回 src.kind/tgt.kind，
 *   下面「首个成员不设限」和「同 kind 放行」会同时变红。
 */
describe('第 6 条：规则跟随只能连接同种类型的节点', () => {
  const rule = defineNode('rule', '规则', [], { rules: {} }, 'r1')
  const tbl = (id: string) => defineNode('table', id, [f], undefined, id)
  const sort = (id: string) => defineNode('sort', id, [f], { sortBy: 'col1', sortDir: 'asc' }, id)

  it('★ 组为空时：首个成员不设限，任何 kind 都能当首个（由它定下本组 kind）', () => {
    // 这条同时是「判据没写反」的证据。若实现是 src.kind/tgt.kind 比较，
    // 这里必然 false —— 因为 rule 的 kind 恒为 'rule'，永不等于 'table'。
    const T = t2(rule, tbl('tb1'), sort('so1'))
    expect(canJoinRuleGroup('r1', 'tb1', T)).toBe(true)
    expect(canJoinRuleGroup('r1', 'so1', T)).toBe(true)
  })

  it('★ 组内已有 table 成员：再连table 放行，连 sort 拒绝', () => {
    const T = t2(rule, tbl('tb1'), tbl('tb2'), sort('so1'))
    const linked = { ...T, tb1: { ...T.tb1, config: { linkedInputs: ['r1'] } } }
    expect(canJoinRuleGroup('r1', 'tb2', linked)).toBe(true)
    expect(canJoinRuleGroup('r1', 'so1', linked)).toBe(false)
  })

  it('★ 反向连接同样受约束（规则锁定无向，问方向就漏了一半）', () => {
    const T = t2(rule, tbl('tb1'), sort('so1'), sort('so2'))
    // 组内已有 sort 成员 so1
    const withSort = { ...T, so1: { ...T.so1, config: { linkedInputs: ['r1'] } } }
    // 反向：再连一个 sort → 放行（反向也要能连同 kind 的）
    expect(canJoinRuleGroup('so2', 'r1', withSort)).toBe(true)
    // 反向：连 table → 拒绝（组内已定为 sort）
    expect(canJoinRuleGroup('tb1', 'r1', withSort)).toBe(false)
    // 正向同一条判定，两个方向结论必须一致
    expect(canJoinRuleGroup('r1', 'so2', withSort)).toBe(true)
    expect(canJoinRuleGroup('r1', 'tb1', withSort)).toBe(false)
  })

  it('组内已有 sort 成员：再连 sort 放行，连 table 拒绝（不因谁先连而漂）', () => {
    const T = t2(rule, tbl('tb1'), sort('so1'), sort('so2'))
    const linked = { ...T, so1: { ...T.so1, config: { linkedInputs: ['r1'] } } }
    expect(canJoinRuleGroup('r1', 'so2', linked)).toBe(true)
    expect(canJoinRuleGroup('r1', 'tb1', linked)).toBe(false)
  })

  it('通过 canConnectPorts 走规则口时，同类型约束同样生效（UI 层接线正确）', () => {
    const T = t2(rule, tbl('tb1'), sort('so1'))
    const linked = { ...T, tb1: { ...T.tb1, config: { linkedInputs: ['r1'] } } }
    expect(canConnectPorts('r1', 'tb1', RULE_PORT_OUT, RULE_PORT_IN, linked)).toBe(true)
    expect(canConnectPorts('r1', 'so1', RULE_PORT_OUT, RULE_PORT_IN, linked)).toBe(false)
    // 反向：普通节点顶部口 → 规则节点下方口
    expect(canConnectPorts('so1', 'r1', RULE_PORT_IN, RULE_PORT_OUT, linked)).toBe(false)
  })

  it('回归底线：数据流口↔数据流口不受同类型约束影响', () => {
    // ★ 第 6 条绝不能顺手把正常数据流连线也按 kind 拦掉：
    //   「参数表→ 排序 → 输出」是不同 kind 的正常流水线。
    const T = t2(tbl('tb1'), sort('so1'), rule)
    expect(canConnectPorts('tb1', 'so1', null, null, T)).toBe(true)
    expect(canConnectPorts('so1', 'tb1', null, null, T)).toBe(true)
  })

  it('两端都不是规则节点：canJoinRuleGroup 直接放行（不越权管数据流）', () => {
    const T = t2(tbl('tb1'), sort('so1'))
    expect(canJoinRuleGroup('tb1', 'so1', T)).toBe(true)
  })

  it('规则节点之间不互锁（互锁成环，effectiveRule 会无限跳）', () => {
    const r2 = defineNode('rule', '规则2', [], { rules: {} }, 'r2')
    expect(canJoinRuleGroup('r1', 'r2', t2(rule, r2))).toBe(false)
  })

  it('自己连自己 / 查不到节点：前者拒绝，后者放行交给落库层兜底', () => {
    const T = t2(rule, tbl('tb1'))
    expect(canJoinRuleGroup('r1', 'r1', T)).toBe(false)
    // 拖到空白处时 React Flow 可能给出未落库的 id，这里若拦会出现「连不上」的假死
    expect(canJoinRuleGroup('r1', 'ghost', T)).toBe(true)
  })

  it('回收站里的节点不接收新成员（与 bindingsOfRuleNode 口径一致）', () => {
    const T = t2(rule, { ...tbl('tb1'), deleted: true })
    expect(canJoinRuleGroup('r1', 'tb1', T)).toBe(false)
  })
})
