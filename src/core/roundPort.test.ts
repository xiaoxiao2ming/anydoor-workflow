// ════════════════════════════════════════════════════════════════════════
// AI / 参数库节点的连线准入与计数（2026-10-06）
// ════════════════════════════════════════════════════════════════════════
// 用户原话：
//   「模仿规则跟随节点样式 增加一个AI节点。AI节点的连线放在普通节点的
//     下1锚点位置。」
//   「模仿规则跟随节点样式 增加一个参数库节点。参数库节点的连线放在
//     普通节点的 上2锚点位置。」
//   → 2026-10-06 用户改口径：「参数库连接节点 锚点 上1」（落点由上2 改为上1）
//
// ── 这组测试为什么必须存在 ──────────────────────────────────────────
//   AI/参数库加入后，「两端都是圆形节点专用口」的组合从 1 种（规则线）
//   变成 3 种。canConnectPorts 的旧防线（「两端规则口但两端都非规则节点
//   拒绝」）只认 rule kind —— 若不升级为完整矩阵，AI 线会被旧防线拒死
//   （症状：圆形节点之间能拖出线但全被 isValidConnection 拦下，一条都连不上）；
//   反过来，错配组合（rule↔ai、普通↔普通走专用口）若放行，又会劈叉成
//   语义不明的连线（rulePort.test.ts 记录过同款事故）。
//   triggerLineSemantics 是矩阵本体，这里逐格验证它。
// ════════════════════════════════════════════════════════════════════════

import { describe, it, expect } from 'vitest'
import {
  canConnectPorts,
  isRulePortId,
  triggerLineSemantics,
  roundLineAnchors,
  roundLineKindOf,
  roundLinkCountOf,
  RULE_PORT_IN,
  RULE_PORT_IN_2,
  RULE_PORT_OUT,
  RULE_PORT_OUT_2,
  AI_PORT_IN,
  PARAMLIB_PORT_OUT,
} from './ruleSync'
import { isRoundNodeKind } from './nodeStyle'
import { defineNode, defineField, type TableSchema } from './schema'

const f = defineField('列', 'text')
const t2 = (...ts: TableSchema[]) => Object.fromEntries(ts.map((t) => [t.id, t]))

const rule = defineNode('rule', '规则', [], { rules: {} }, 'r1')
const ai = defineNode('ai', 'AI', [], {}, 'a1')
const paramlib = defineNode('paramlib', '参数库', [], {}, 'p1')
const tbl = (id: string) => defineNode('table', id, [f], undefined, id)

describe('圆形节点判定：isRoundNodeKind', () => {
  it('三家圆形：rule / ai / paramlib 都认', () => {
    expect(isRoundNodeKind('rule')).toBe(true)
    expect(isRoundNodeKind('ai')).toBe(true)
    expect(isRoundNodeKind('paramlib')).toBe(true)
  })
  it('普通节点与未知值不认', () => {
    expect(isRoundNodeKind('table')).toBe(false)
    expect(isRoundNodeKind('param')).toBe(false)
    expect(isRoundNodeKind(undefined)).toBe(false)
  })
})

describe('专用口登记：isRulePortId 必须认 AI/参数库口', () => {
  it('新口 id 都认（漏登记 = 新口被当数据流口，连线被静默拒绝）', () => {
    expect(isRulePortId(AI_PORT_IN)).toBe(true)
    expect(isRulePortId(PARAMLIB_PORT_OUT)).toBe(true)
  })
  it('原有四口仍认（回归底线）', () => {
    expect(isRulePortId(RULE_PORT_IN)).toBe(true)
    expect(isRulePortId(RULE_PORT_IN_2)).toBe(true)
    expect(isRulePortId(RULE_PORT_OUT)).toBe(true)
    expect(isRulePortId(RULE_PORT_OUT_2)).toBe(true)
  })
})

describe('连线语义矩阵：triggerLineSemantics', () => {
  it('AI 线：普通节点下1/下2 → AI 口，正向反向都放行', () => {
    expect(triggerLineSemantics('table', 'ai', RULE_PORT_OUT, AI_PORT_IN)).toBe('ai')
    expect(triggerLineSemantics('table', 'ai', RULE_PORT_OUT_2, AI_PORT_IN)).toBe('ai')
    // 反向拖线（RF 交换 source/target）：AI 口在 sourceHandle 位
    expect(triggerLineSemantics('ai', 'table', AI_PORT_IN, RULE_PORT_OUT)).toBe('ai')
  })

  it('参数库线：参数库口 → 普通节点上1/上2，正向反向都放行', () => {
    expect(triggerLineSemantics('paramlib', 'table', PARAMLIB_PORT_OUT, RULE_PORT_IN)).toBe('paramlib')
    expect(triggerLineSemantics('paramlib', 'table', PARAMLIB_PORT_OUT, RULE_PORT_IN_2)).toBe('paramlib')
    // 反向拖线：普通节点上2 在 sourceHandle 位
    expect(triggerLineSemantics('table', 'paramlib', RULE_PORT_IN_2, PARAMLIB_PORT_OUT)).toBe('paramlib')
  })

  it('规则线照旧：规则节点下1 ↔ 普通节点上1/上2（既有行为零漂移）', () => {
    expect(triggerLineSemantics('rule', 'table', RULE_PORT_OUT, RULE_PORT_IN)).toBe('rule')
    expect(triggerLineSemantics('rule', 'table', RULE_PORT_OUT, RULE_PORT_IN_2)).toBe('rule')
    expect(triggerLineSemantics('table', 'rule', RULE_PORT_IN, RULE_PORT_OUT)).toBe('rule')
  })

  it('★ 错配拒绝：口与节点 kind 对不上（rule↔ai、ai↔paramlib、口位错用）', () => {
    // 规则节点连 AI：两家圆形互连，语义不明（会成环/无意义）
    expect(triggerLineSemantics('rule', 'ai', RULE_PORT_OUT, AI_PORT_IN)).toBeNull()
    // AI 连参数库：同上
    expect(triggerLineSemantics('ai', 'paramlib', AI_PORT_IN, PARAMLIB_PORT_OUT)).toBeNull()
    // AI 口去接普通节点的上口：AI 线只认下1/下2 出线
    expect(triggerLineSemantics('table', 'ai', RULE_PORT_IN, AI_PORT_IN)).toBeNull()
    // 参数库口去接普通节点的下口：参数库线只认上1/上2 落点
    expect(triggerLineSemantics('paramlib', 'table', PARAMLIB_PORT_OUT, RULE_PORT_OUT)).toBeNull()
    // AI 口连到规则落点口：错配
    expect(triggerLineSemantics('table', 'ai', RULE_PORT_OUT, RULE_PORT_IN)).toBeNull()
  })

  it('★ 两端都普通节点 / 两端都圆形：一律 null（沿用旧防线语义）', () => {
    // 普通↔普通走专用口：正是 rulePort.test.ts 记录的「劈叉成数据流线」事故形态
    expect(triggerLineSemantics('table', 'table', RULE_PORT_OUT, RULE_PORT_IN)).toBeNull()
    // 圆形三家两两互连全部拒绝
    expect(triggerLineSemantics('rule', 'rule', RULE_PORT_OUT, RULE_PORT_OUT)).toBeNull()
  })
})

describe('UI 准入：canConnectPorts 的 AI/参数库组合（带 tables 五参）', () => {
  it('普通节点下1 → AI：放行（用户口径「连线放在下1锚点位置」）', () => {
    const T = t2(tbl('tb1'), ai)
    expect(canConnectPorts('tb1', 'a1', RULE_PORT_OUT, AI_PORT_IN, T)).toBe(true)
  })

  it('参数库 → 普通节点上2：放行（用户口径「连线放在上2锚点位置」）', () => {
    const T = t2(paramlib, tbl('tb1'))
    expect(canConnectPorts('p1', 'tb1', PARAMLIB_PORT_OUT, RULE_PORT_IN_2, T)).toBe(true)
  })

  it('★ 圆形互连拒绝：rule→ai、ai→paramlib、paramlib→rule 全拦', () => {
    const T = t2(rule, ai, paramlib, tbl('tb1'))
    expect(canConnectPorts('r1', 'a1', RULE_PORT_OUT, AI_PORT_IN, T)).toBe(false)
    expect(canConnectPorts('a1', 'p1', AI_PORT_IN, PARAMLIB_PORT_OUT, T)).toBe(false)
    expect(canConnectPorts('p1', 'r1', PARAMLIB_PORT_OUT, RULE_PORT_IN_2, T)).toBe(false)
  })

  it('★ 普通↔普通走专用口仍拒绝（旧防线升级后语义不回退）', () => {
    const T = t2(tbl('s1'), tbl('s2'))
    expect(canConnectPorts('s1', 's2', RULE_PORT_OUT, RULE_PORT_IN, T)).toBe(false)
    expect(canConnectPorts('s1', 's2', RULE_PORT_OUT, AI_PORT_IN, T)).toBe(false)
    expect(canConnectPorts('s1', 's2', PARAMLIB_PORT_OUT, RULE_PORT_IN_2, T)).toBe(false)
  })

  it('AI 线无「同种类型」约束：table/sort/map 都能连 AI（与规则线不同）', () => {
    const T = t2(tbl('tb1'), ai)
    expect(canConnectPorts('tb1', 'a1', RULE_PORT_OUT, AI_PORT_IN, T)).toBe(true)
  })

  it('数据流口↔数据流口：不受矩阵影响（回归底线）', () => {
    const T = t2(tbl('tb1'), tbl('tb2'), ai)
    expect(canConnectPorts('tb1', 'tb2', null, null, T)).toBe(true)
  })

  it('节点查不到（拖到空白处）：放行交落库层兜底', () => {
    const T = t2(tbl('tb1'), ai)
    expect(canConnectPorts('tb1', 'ghost', RULE_PORT_OUT, AI_PORT_IN, T)).toBe(true)
    expect(canConnectPorts('ghost', 'a1', RULE_PORT_OUT, AI_PORT_IN, T)).toBe(true)
  })

  it('不传 tables（四参）：AI/参数库组合只校验口角色（旧调用语义不变）', () => {
    expect(canConnectPorts('tb1', 'a1', RULE_PORT_OUT, AI_PORT_IN)).toBe(true)
    expect(canConnectPorts('p1', 'tb1', PARAMLIB_PORT_OUT, RULE_PORT_IN_2)).toBe(true)
  })
})

describe('圆内计数：roundLinkCountOf', () => {
  it('AI 节点：数自己 linkedInputs 里的有效上游', () => {
    const a = { ...ai, config: { linkedInputs: ['tb1', 'tb2'] } }
    const T = t2(a, tbl('tb1'), tbl('tb2'))
    expect(roundLinkCountOf('a1', T)).toBe(2)
  })

  it('参数库节点：数「把我记进 linkedInputs」的有效下游（全表扫，无反向索引）', () => {
    const d1 = { ...tbl('d1'), config: { linkedInputs: ['p1'] } }
    const d2 = { ...tbl('d2'), config: { linkedInputs: ['p1', 'a1'] } }
    const T = t2(paramlib, d1, d2, ai)
    expect(roundLinkCountOf('p1', T)).toBe(2)
  })

  it('双向混合：AI 与参数库两个方向都数（同一节点只可能占一个方向，防御性并集）', () => {
    const a = { ...ai, config: { linkedInputs: ['tb1'] } }
    const d1 = { ...tbl('d1'), config: { linkedInputs: ['a1'] } }
    const T = t2(a, d1, tbl('tb1'))
    expect(roundLinkCountOf('a1', T)).toBe(2)
  })

  it('回收站节点不计入（与 edgesOf 的口径一致：软删除不产生边）', () => {
    const a = { ...ai, config: { linkedInputs: ['tb1', 'tb2'] } }
    const T = t2(a, tbl('tb1'), { ...tbl('tb2'), deleted: true })
    expect(roundLinkCountOf('a1', T)).toBe(1)
  })

  it('无连线 / 节点不存在：返回 0（不抛错）', () => {
    expect(roundLinkCountOf('a1', t2(ai))).toBe(0)
    expect(roundLinkCountOf('ghost', t2(ai))).toBe(0)
  })
})

// ════════════════════════════════════════════════════════════════════════
// 锚点落位：roundLineAnchors（2026-10-06）
// ════════════════════════════════════════════════════════════════════════
// 这组测试的存在理由与上方 canConnectPorts 那组**完全不同**，不能合并：
//   canConnectPorts / triggerLineSemantics 回答「**能不能连**」（能力判定）；
//   roundLineAnchors 回答「**连到哪儿**」（默认落点）。
//   两者刻意分开，所以「准入放行上2」与「默认落上1」可以同时成立 ——
//   下面的「★ 准入 ≠ 落点」一条就是钉这个的。
//
// ★★ 上一轮真实踩过的坑（这也是为什么方向判定放在被测函数内）：
//   第一版签名是 `roundLineAnchors(lineKind, roundIsSource: boolean)`，
//   Canvas 传的是 `srcT.kind !== 'rule' ? false : true`。
//   那个式子只对规则线成立 —— 参数库线的 source 端恰恰不是 rule，
//   于是「参数库在 source 端」（最常见的场景）被判成反向：
//   sourceHandle/targetHandle 整体对调，RF 找不到对应口，线连不出来。
//   症状与「锚点 id 写错」一模一样，但根因在**方向推导**而非 id。
//   现签名只收 srcKind（事实），方向由函数内按线种推。
// ════════════════════════════════════════════════════════════════════════
describe('锚点落位：roundLineAnchors（画布实际连到哪儿）', () => {
  it('★ 参数库线：参数库在 source 端 → 普通节点落上1（用户 2026-10-06 改口径，原为上2）', () => {
    expect(roundLineAnchors('paramlib', 'paramlib')).toEqual({
      sourceHandle: PARAMLIB_PORT_OUT,
      targetHandle: RULE_PORT_IN,
    })
  })

  it('参数库线反向：普通节点在 source 端 → 上1 在 source 位、参数库口在 target 位', () => {
    // 这一条同时是上一轮那个 bug 的**直接回归锁**：
    // 若方向推导又跑回「只看是不是 rule」，本条与上一条会整体对调而失败。
    expect(roundLineAnchors('paramlib', 'table')).toEqual({
      sourceHandle: RULE_PORT_IN,
      targetHandle: PARAMLIB_PORT_OUT,
    })
  })

  it('规则线：仍走上2（与参数库刻意错开，零漂移）', () => {
    expect(roundLineAnchors('rule', 'rule')).toEqual({
      sourceHandle: RULE_PORT_OUT,
      targetHandle: RULE_PORT_IN_2,
    })
    expect(roundLineAnchors('rule', 'table')).toEqual({
      sourceHandle: RULE_PORT_IN_2,
      targetHandle: RULE_PORT_OUT,
    })
  })

  it('AI 线：普通节点下1 → AI 上沿（正反两向都要有明确答案，不许含糊）', () => {
    expect(roundLineAnchors('ai', 'table')).toEqual({
      sourceHandle: RULE_PORT_OUT,
      targetHandle: AI_PORT_IN,
    })
    // 反边（AI 口是 target，真实数据不该出现）也必须给出确定落位，
    // 不能返回空 —— 返回空会让 RF 悄悄退回默认口。
    expect(roundLineAnchors('ai', 'ai')).toEqual({
      sourceHandle: AI_PORT_IN,
      targetHandle: RULE_PORT_OUT,
    })
  })

  it('非圆形专属线 → 空对象（RF 走默认数据流口，不落到圆形口上）', () => {
    expect(roundLineAnchors(null, 'table')).toEqual({})
    expect(roundLineAnchors(null, 'rule')).toEqual({})
  })

  it('★ 准入放行上2，但默认落点必须是上1（这是两个不同概念）', () => {
    // 能力判定：上2 仍然可以手动拖连过去（那是用户的选择，不是 bug）
    expect(triggerLineSemantics('paramlib', 'table', PARAMLIB_PORT_OUT, RULE_PORT_IN_2)).toBe('paramlib')
    expect(triggerLineSemantics('paramlib', 'table', PARAMLIB_PORT_OUT, RULE_PORT_IN)).toBe('paramlib')
    // 默认落点：上1
    expect(roundLineAnchors('paramlib', 'paramlib').targetHandle).toBe(RULE_PORT_IN)
    expect(roundLineAnchors('paramlib', 'paramlib').targetHandle).not.toBe(RULE_PORT_IN_2)
  })

  it('★ 规则线与参数库线在同一节点上必须落在不同的口（否则视觉混淆）', () => {
    // 同一节点既被规则锁定又被参数库供参数时，两条线要能一眼分清。
    const ruleTarget = roundLineAnchors('rule', 'rule').targetHandle
    const paramTarget = roundLineAnchors('paramlib', 'paramlib').targetHandle
    expect(ruleTarget).toBe(RULE_PORT_IN_2)
    expect(paramTarget).toBe(RULE_PORT_IN)
    expect(ruleTarget).not.toBe(paramTarget)
  })

  it('返回的两端都是已登记的专用口（漏登记 isRulePortId = 该端被当数据流口处理）', () => {
    const cases: Array<['rule' | 'ai' | 'paramlib', string]> = [
      ['rule', 'rule'],
      ['rule', 'table'],
      ['ai', 'table'],
      ['ai', 'ai'],
      ['paramlib', 'paramlib'],
      ['paramlib', 'table'],
    ]
    for (const [kind, srcKind] of cases) {
      const a = roundLineAnchors(kind, srcKind)
      expect(isRulePortId(a.sourceHandle)).toBe(true)
      expect(isRulePortId(a.targetHandle)).toBe(true)
    }
  })

  it('★ 边引用的每个口，都必须真实挂在那一端的节点上（口归属表逐格核对）', () => {
    // 口的归属，直接照抄 Canvas.tsx 的 Handle 声明：
    //   RULE_PORT_IN / RULE_PORT_IN_2 / RULE_PORT_OUT / RULE_PORT_OUT_2 → 普通节点（!isRound）
    //   RULE_PORT_OUT        → 规则节点也在底部挂了一个（同位复用）
    //   AI_PORT_IN           → 仅 ai 节点
    //   PARAMLIB_PORT_OUT→ 仅 paramlib 节点
    // 这条断言防的是「给 paramlib 节点派一个 RULE_PORT_IN_2」这类错派：
    // 症状是线连不出来（RF 找不到该 id 的 handle，退回节点中心），
    // 但 tsc 与准入矩阵测试全都察觉不到。
    const OWNER: Record<string, string[]> = {
      [RULE_PORT_IN]: ['table'],
      [RULE_PORT_IN_2]: ['table'],
      [RULE_PORT_OUT]: ['table', 'rule'],
      [RULE_PORT_OUT_2]: ['table'],
      [AI_PORT_IN]: ['ai'],
      [PARAMLIB_PORT_OUT]: ['paramlib'],
    }
    const cases: Array<['rule' | 'ai' | 'paramlib', string]> = [
      ['rule', 'rule'],
      ['rule', 'table'],
      ['ai', 'table'],
      ['ai', 'ai'],
      ['paramlib', 'paramlib'],
      ['paramlib', 'table'],
    ]
    for (const [lineKind, srcKind] of cases) {
      const a = roundLineAnchors(lineKind, srcKind)
      // 圆形专属线的另一端恒为普通节点
      const tgtKind = srcKind === lineKind ? 'table' : lineKind
      expect(OWNER[a.sourceHandle as string]).toContain(srcKind)
      expect(OWNER[a.targetHandle as string]).toContain(tgtKind)
    }
  })

  it('反向边的口在几何上仍是「线从普通节点顶出、到圆形节点底入」', () => {
    // 记录一个容易误判的点：RF 的 Handle type="target/source" 是给
    // **连接校验**（isValidConnection）用的，不是渲染约束。
    // 反向边（普通节点在source 端）把 RULE_PORT_IN_2 放在 sourceHandle 上
    // 是正确的 —— 该 handle 就挂在普通节点顶部，线从那里画出来才对。
    // ★ 抽离前的内联实现在这里返回的是 undefined，
    //   RF 找不到 target 节点的 'rule-in-2' 就退回节点中心画线（视觉上从圆心出发）。
    //   抽离后顺带修掉了这个，但**不是**本次需求要的改动，故单测钉住当前正确行为。
    expect(roundLineAnchors('rule', 'table')).toEqual({
      sourceHandle: RULE_PORT_IN_2,
      targetHandle: RULE_PORT_OUT,
    })
    expect(roundLineAnchors('paramlib', 'table')).toEqual({
      sourceHandle: RULE_PORT_IN,
      targetHandle: PARAMLIB_PORT_OUT,
    })
  })
})

// ════════════════════════════════════════════════════════════════════════
// 判定 → 落位：roundLineKindOf + roundLineAnchors 串起来断言（2026-10-06）
// ════════════════════════════════════════════════════════════════════════
// ★★ 这组是本次改动的**真正护栏**，上面那两组都拦不住下面的错：
//   上一轮把参数库落点从上2 改上1 时，全量 777 个用例全绿 —— 因为
//   「线种判定」和「锚点落位」分处两地、各自内联，谁也没验证过
//   「被判为paramlib 的边，落位一定是上1」。只有把两个函数**串起来跑**，
//   才能断言这条端到端不变量。
//   判据来源：画布真实数据里圆形专属线的两端组合（普通↔rule/ai/paramlib）。
// ════════════════════════════════════════════════════════════════════════
describe('判定→落位 端到端不变量：判为 paramlib 的边必落上1', () => {
  it('★ 参数库在 source 端（落库层唯一真实形态）→ 上1', () => {
    const kind = roundLineKindOf('flow', 'paramlib', 'table')
    expect(kind).toBe('paramlib')
    expect(roundLineAnchors(kind, 'paramlib')).toEqual({
      sourceHandle: PARAMLIB_PORT_OUT,
      targetHandle: RULE_PORT_IN,
    })
  })

  it('参数库在 target 端（反边，真实数据不该出现）→ 仍须给确定落位，不许空', () => {
    const kind = roundLineKindOf('flow', 'table', 'paramlib')
    expect(kind).toBe('paramlib')
    const a = roundLineAnchors(kind, 'table')
    expect(a.sourceHandle).toBe(RULE_PORT_IN)
    expect(a.targetHandle).toBe(PARAMLIB_PORT_OUT)
  })

  it('规则线恒落上2、参数库线恒落上1 —— 正反两向扫，两者永不撞同一个口', () => {
    const combos: Array<[string, string]> = [
      ['rule', 'table'],
      ['table', 'rule'],
      ['paramlib', 'table'],
      ['table', 'paramlib'],
    ]
    const plainTargets = new Map<string, string>()
    for (const [srcKind, tgtKind] of combos) {
      const kind = roundLineKindOf('flow', srcKind, tgtKind)
      const a = roundLineAnchors(kind, srcKind)
      // 只关心「普通节点那一端落在哪个口」
      const plainPort = srcKind === 'table' ? a.sourceHandle : a.targetHandle
      if (kind === 'rule') expect(plainPort).toBe(RULE_PORT_IN_2)
      if (kind === 'paramlib') expect(plainPort).toBe(RULE_PORT_IN)
      plainTargets.set(`${kind}:${srcKind}`, plainPort as string)
    }
    // 一条线的正反两向**应当**得到同一个口（那是同一条线的两端），
    // 所以 4 种组合只对应 2 个「线种:口」取值 —— 若这里出现 4 个，
    // 说明反向时被派了另一个口，那正是本轮修掉的方向 bug。
    expect([...new Set(plainTargets.values())].sort()).toEqual([RULE_PORT_IN, RULE_PORT_IN_2].sort())
  })

  it('判定优先级：rule > ai > paramlib（与 App.handleConnect 分支顺序同源）', () => {
    // 混搭组合在准入层本就被拒，但判定层必须**确定**地给一个答案，
    // 不能随代码顺序漂移。落库层 rule 分支排在最前，故判定也 rule 优先。
    expect(roundLineKindOf('flow', 'rule', 'ai')).toBe('rule')
    expect(roundLineKindOf('flow', 'ai', 'paramlib')).toBe('ai')
    expect(roundLineKindOf('flow', 'rule', 'paramlib')).toBe('rule')
  })

  it('非 flow 边 / 两端都是普通节点 / 节点查不到 → 一律不是圆形专属线', () => {
    expect(roundLineKindOf('assoc', 'rule', 'table')).toBe(null)
    expect(roundLineKindOf('lookup', 'table', 'table')).toBe(null)
    expect(roundLineKindOf('flow', 'table', 'param')).toBe(null)
    // 节点被删/未加载：宁可退回默认数据流口，也不要瞎派一个圆形口
    expect(roundLineKindOf('flow', undefined, 'table')).toBe(null)
    expect(roundLineKindOf('flow', 'paramlib', undefined)).toBe(null)
    expect(roundLineKindOf(undefined, 'paramlib', 'table')).toBe(null)
  })

  it('★ 判定为圆形专属线的合法边，两端口必定落在真实存在的 handle 上（联合不变量）', () => {
    // 把「准入放行 + 判定为圆形专属线」的边全扫一遍，断言派出的口
    // 在 Canvas.tsx 里确实有对应 Handle 声明。
    const OWNER: Record<string, string[]> = {
      [RULE_PORT_IN]: ['table'],
      [RULE_PORT_IN_2]: ['table'],
      [RULE_PORT_OUT]: ['table', 'rule'],
      [RULE_PORT_OUT_2]: ['table'],
      [AI_PORT_IN]: ['ai'],
      [PARAMLIB_PORT_OUT]: ['paramlib'],
    }
    // 各线种的「圆端口 ↔ 普通端口」配对（与 triggerLineSemantics 矩阵同源）
    //   ★ 圆端口 = 挂在**圆形节点**上的那个口。AI 节点上只有 AI_PORT_IN，
    //     RULE_PORT_OUT 挂在普通节点下1（以及规则节点底部），别填反。
    const LEGAL: Array<[string, string, string, string]> = [
      // [线种, 圆端口, 普通端口, 圆形侧 kind]
      ['rule', RULE_PORT_OUT, RULE_PORT_IN_2, 'rule'],
      ['ai', AI_PORT_IN, RULE_PORT_OUT, 'ai'],
      ['paramlib', PARAMLIB_PORT_OUT, RULE_PORT_IN, 'paramlib'],
    ]
    let checked = 0
    for (const [lineKind, roundPort, plainPort, roundKind] of LEGAL) {
      for (const [srcKind, tgtKind] of [[roundKind, 'table'], ['table', roundKind]] as Array<[string, string]>) {
        // 前置条件：这条边必须真的过准入，否则它不存在、断言无意义
        expect(triggerLineSemantics(srcKind, tgtKind, srcKind === 'table' ? plainPort : roundPort, tgtKind === 'table' ? plainPort : roundPort)).toBe(lineKind)
        const kind = roundLineKindOf('flow', srcKind, tgtKind)
        expect(kind).toBe(lineKind)
        const a = roundLineAnchors(kind, srcKind)
        expect(a.sourceHandle).toBeDefined()
        expect(a.targetHandle).toBeDefined()
        expect(OWNER[a.sourceHandle as string]).toContain(srcKind)
        expect(OWNER[a.targetHandle as string]).toContain(tgtKind)
        // 且普通节点那一端必须落在该线种约定的口上
        expect(srcKind === 'table' ? a.sourceHandle : a.targetHandle).toBe(plainPort)
        checked++
      }
    }
    // ★ 不扫混搭圆形组合（rule↔ai、ai↔paramlib 等）：
    //   它们在 triggerLineSemantics 里一律返回 null，准入层就拒了，
    //   真实工程里不存在这种边。而 roundLineKindOf 的优先级短路
    //   （rule > ai > paramlib）会对它们给出一个「判定」，
    //   若拿这个判定去派口就会派到该节点根本没有的口上——
    //   但那条路径走不到，所以不构成需要防的缺陷，不纳入本不变量。
    // 确认扫描真的跑到了东西（防「循环写错导致一条都没检查」的空集假绿）
    expect(checked).toBe(6)
  })
})
