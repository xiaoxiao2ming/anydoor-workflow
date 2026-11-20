// ════════════════════════════════════════════════════════════════════════
// 规则同步层 —— 圆形「规则节点」给连上的节点发规则
// ════════════════════════════════════════════════════════════════════════
// 需求（用户 2026-10-05 原话）：
//   「画布内容中增加个圆形的节点。类似于功能节点。为规则适配节点。
//     对画布中，相同类型的节点进行规则匹配。一旦匹配了，默认不能每个改规则，
//     只能改主流程的节点进行修改。修改后所有相同规则的节点规则同步改。」
//   「圆形节点是对连线的节点进行规则锁定（统一）圆形节点也在画布上。」
//
// ── 为什么规则只存一份（最重要的一条） ─────────────────────────────────
//   最直觉的写法是「主节点改完，把规则复制写进每个辅节点」。
//   本层**刻意不这么写**，因为复制式同步必然漏：
//     · 锁定5 个辅节点 → 5 处写操作，中途抛错就只改了 3 个
//     · 辅节点被删除 → 复制品成了孤儿垃圾
//     · 主节点改规则 → 忘了同步，副本与主本漂移
//   漂移的隐蔽性极高：画面看起来一切正常，只是某处不跟着变，
//   等发现时往往已经跑过很多批数据。
//
//   故：**规则只存在规则节点身上，辅节点只存 ruleRef（指向谁）。**
//   「同步」于是变成「读的时候跳回源取」—— 不是一次会漏的写操作，
//   而是**结构上不可能漏**的读操作。
//
// ── 纯计算层铁律 ──────────────────────────────────────────────────────
// 本文件禁止 import electron / `node:` / 碰 document（同 ref/、xlsx/field.ts）。
// 纯函数，可被 vitest 直接单测，不依赖画布与 Electron 运行时。
// ════════════════════════════════════════════════════════════════════════

import type { TableSchema, NodeRuleSet, NodeKind } from './schema'
import { isRuleNodeKind, isRoundNodeKind } from './nodeStyle'

// ════════════════════════════════════════════════════════════════════════
// 连接口角色（用户 2026-10-05 定：「规则线只能连接常规节点顶部的点」）
// ════════════════════════════════════════════════════════════════════════
/**
 * React Flow 的 handle id。**只给规则口用**，数据流口一律留null。
 *
 * ★ 为什么用 id 而不是靠 `type` 区分（这是本条约束的核心）：
 *   React Flow 的 Strict 模式只禁「同类型对同类型」（target→target、
 *   source→source），异类型一律放行（源码 @xyflow/system isValidHandle：
 *   `connectionMode === Strict ? fromType !== handleType : ...`）。
 *   而规则口是 target、规则节点下方的口也是…… 一旦两类口type 撞在一起，
 *   规则口就会和**左侧数据流口**（target）无差别地被连上：
 *     规则节点下方口 → 普通节点左侧口= 被记成「数据流线」，
 *     若目标是排序/输出还会走 isMergeKind 分支，规则节点变成「输出的上游」。
 *   光看 type 拦不住，必须给规则口一个专属 id 才能在 UI 层区分。
 */
export const RULE_PORT_IN = 'rule-in'
/** 上2 的专属 id（2026-10-05 用户：「统一修改上1的锚点 改为上2位置」）。
 *
 * ★ 为什么上2 不能再与上1 共用 RULE_PORT_IN：
 *   画布连线（Canvas.edgesOf）给规则线显式指 targetHandle 时，
 *   RF 在**同一节点内**按 id 找 handle，找到**第一个**匹配项就落笔——
 *   上1/上2 同 id 的年代，规则线永远落在上1，上2 成了纯摆设。
 *   拆成专属 id 后「落点=上2」才是一条可指名的指令，而不是碰运气的取首。
 */
export const RULE_PORT_IN_2 = 'rule-in-2'
export const RULE_PORT_OUT = 'rule-out'
/** 常规节点底部的「下2」（2026-10-05 用户：「增加节点 下1 下2 的锚点」）。
 *  常规节点的下1 复用 RULE_PORT_OUT（与规则节点出线口同 id 同角色，
 *  RF 的 handle 按**节点内**查找，不同节点同 id 互不干扰）；
 *  下2 是第二个口，必须有自己的 id —— 同节点内同 id 会被 RF 取首。 */
export const RULE_PORT_OUT_2 = 'rule-out-2'

/* ── AI / 参数库节点的专属口（2026-10-06，用户：
      「模仿规则跟随节点样式 增加一个AI节点。AI节点的连线放在普通节点的
       下1锚点位置」「……参数库节点……上2锚点位置」）── */

/** AI 节点的接收口（target，位于 AI 圆上沿）：
 *  连线 = 普通节点下1（RULE_PORT_OUT，source）→ 本口。AI 吃数据。 */
export const AI_PORT_IN = 'ai-in'

/** 参数库节点的出线口（source，位于参数库圆下沿 —— 与规则节点出线口同位）：
 *  连线 = 本口 → 普通节点上1/上2（RULE_PORT_IN / RULE_PORT_IN_2）。参数库供数。
 *  ★ 画布默认落点是**上1**（用户 2026-10-06：「参数库连接节点 锚点 上1」，
 *    原为上2）。与规则线（默认上2）刻意错开，两种线在同一节点上不重叠。 */
export const PARAMLIB_PORT_OUT = 'paramlib-out'

/** 这个 handle id 是不是规则口（忽略 null / undefined / 普通数据流口）。
 *  ★ 新增口 id 时**必须**来这里登记：本函数是 canConnectPorts 的唯一判据，
 *    漏登记 = 新口被当成数据流口，规则线落到它身上会被静默拒绝。
 *  ★ 2026-10-06 起它实际判的是「圆形节点专用口」全集：
 *    规则口四件套 + AI 接收口 + 参数库出线口。函数名沿用旧名不改 ——
 *    改名会波及全部调用点与既有测试，收益为零（语义见本注释）。 */
export function isRulePortId(id: string | null | undefined): boolean {
  return (
    id === RULE_PORT_IN ||
    id === RULE_PORT_IN_2 ||
    id === RULE_PORT_OUT ||
    id === RULE_PORT_OUT_2 ||
    id === AI_PORT_IN ||
    id === PARAMLIB_PORT_OUT
  )
}

/**
 * 这条「两端都是圆形节点专用口」的连线属于哪种语义（2026-10-06）。
 *
 * ★ 为什么抽成纯函数：canConnectPorts 是 UI 准入层，AI/参数库加入后
 *   「两端专用口」的组合从 1 种（规则线）变成 3 种，准入矩阵必须按
 *   「圆形节点在哪端 + 两端口 id」精确判定 —— 内联在 canConnectPorts 里
 *   会重蹈「判定写两处→漂移」的覆辙（同 ruleConnectPatch 抽离的理由）。
 *
 * ★ 方向无关：React Flow 支持从 target 口反向拖出（RF 自动交换 source/target，
 *   既有测试 canConnectPorts('s1','r1',RULE_PORT_IN,RULE_PORT_OUT) 依赖此行为），
 *   故本函数不假设「圆形端一定在 source」，两端 handle 都参与判定。
 *
 * 准入矩阵（口 id ↔ 节点 kind 对）：
 *   · 规则线   ：rule 节点出线口（RULE_PORT_OUT）↔ 普通节点规则落点口（IN/IN_2）
 *   · AI 线    ：AI 接收口（AI_PORT_IN）↔ 普通节点下1/下2（RULE_PORT_OUT/OUT_2）
 *   · 参数库线 ：参数库出线口（PARAMLIB_PORT_OUT）↔ 普通节点上1/上2（IN/IN_2）
 *     —— 画布默认落点为**上1**（见下方注释），上2 仍接受用户手动拖连
 *   · 圆形节点之间互连、两端都普通节点、口与节点 kind 错配 → 一律 null（拒绝）
 *
 * @returns 'rule' | 'ai' | 'paramlib' = 合法的连线语义；null = 拒绝
 */
export function triggerLineSemantics(
  srcKind: string,
  tgtKind: string,
  sourceHandle: string | null | undefined,
  targetHandle: string | null | undefined,
): 'rule' | 'ai' | 'paramlib' | null {
  const roundS = isRoundNodeKind(srcKind)
  const roundT = isRoundNodeKind(tgtKind)
  // 圆形节点之间互连不认（会成环/无语义）；两端都普通节点却走专用口也不认
  // （既有防线：从下1连到上1的普通↔普通连线必须拒绝）
  if (roundS === roundT) return null
  const roundHandle = roundS ? sourceHandle : targetHandle // 圆形节点那端的口
  const plainHandle = roundS ? targetHandle : sourceHandle // 普通节点那端的口
  const roundKind = roundS ? srcKind : tgtKind
  if (roundKind === 'rule') {
    // 规则线：圆形端必须是规则节点出线口（下1），普通端必须是规则落点口（上1/上2）
    const okRound = roundHandle === RULE_PORT_OUT
    const okPlain = plainHandle === RULE_PORT_IN || plainHandle === RULE_PORT_IN_2
    return okRound && okPlain ? 'rule' : null
  }
  if (roundKind === 'ai') {
    // AI 线：普通节点下1/下2 出线 → AI 接收口（用户：「连线放在普通节点的 下1锚点位置」）
    const okRound = roundHandle === AI_PORT_IN
    const okPlain = plainHandle === RULE_PORT_OUT || plainHandle === RULE_PORT_OUT_2
    return okRound && okPlain ? 'ai' : null
  }
  // 参数库线：参数库出线口 → 普通节点上1/上2
  //  ★ 画布实际落点是**上1**（RULE_PORT_IN，用户 2026-10-06：「参数库连接节点 锚点 上1」）。
  //    本函数仍放行上2：准入是「能不能连」的**能力**判定，画布落点是**默认位置**，
  //    两者刻意分开 —— 用户从口上拖线时落在上2 应当被接受（那是用户的选择，不是 bug），
  //    而未指定口的自动连线按上1 落。
  const okRound = roundHandle === PARAMLIB_PORT_OUT
  const okPlain = plainHandle === RULE_PORT_IN || plainHandle === RULE_PORT_IN_2
  return okRound && okPlain ? 'paramlib' : null
}

/**
 * 画布层连线准入（喂给 React Flow 的 `isValidConnection`）。
 *
 * 规则（用户口径「规则线只能连接常规节点顶部的点」+ 2026-10-06 AI/参数库扩充）：
 *   · 两端**都是圆形节点专用口** → 按 triggerLineSemantics 的准入矩阵判定
 *     （规则线沿用 canJoinRuleGroup 的同类型约束；AI/参数库线无同类型约束）
 *   · 一端专用口、一端数据流口 → **拒绝**（这是本条约束修掉的真缺陷）
 *   · 两端都是数据流口 → 放行（原有数据流连线一行不改）
 *
 * ★ 为什么规则口之间「两个方向都通」而不用改成ConnectionMode.Loose：
 *   规则口是 target、规则节点口是 source，异类型在 Strict 下本来就双向合法
 *   （源码见上）。内核 ruleConnectPatch 也是对称处理两个方向的。
 *   改全局连接模式会波及所有数据流连线，风险远大于收益——
 *   **在isValidConnection 里精确开口子，才是把约束写在最小作用域里。**
 *
 * @param source/target 节点 id；sourceHandle/targetHandle 为 handle id
 * @param tables 可选。传入后才校验「同种类型」（用户 2026-10-05 第6 条：
 *   「规则跟随节点只能连接同种类型的节点」）与口↔kind 矩阵。
 *   不传 = 只校验口角色，保持既有调用与测试的语义不变。
 *
 * ★★ 为什么同类型约束必须放在这里、且参数可选：
 *   这个判定**需要 tables 才能算**（要比现有成员的 kind），
 *   而 isValidConnection 的签名只有 Connection 参数 —— 只能由调用方喂进来。
 *   参数可选而非必填：内核已有 20+ 个用例直接调本函数四参版本，
 *   改签名会把它们全变成"忘了传 tables = 不校验"，反而更危险。
 *   故本函数只做**口角色**判定，同类型判定委托给 canJoinRuleGroup。
 */
export function canConnectPorts(
  source: string,
  target: string,
  sourceHandle: string | null | undefined,
  targetHandle: string | null | undefined,
  tables?: Record<string, TableSchema>,
): boolean {
  const s = isRulePortId(sourceHandle)
  const t = isRulePortId(targetHandle)
  // 同类才放行：专用口↔专用口（规则/AI/参数库）、数据流口↔数据流口（原有数据流）
  if (s !== t) return false
  // 两端数据流口 → 原有数据流，一行不改（matrix 只管专用口组合）
  if (!s) return true
  // 两端专用口但没给 tables：只校验到「口角色」，语义判定交落库层兜底
  // （既有四参语义 —— 20+ 个用例依赖，不能破坏）
  if (!tables) return true
  const src = tables[source]
  const tgt = tables[target]
  // 节点查不到（拖到空白处）时不拦，交由落库层兜底，避免假死（既有行为）
  if (!src || !tgt) return true
  // ★ 2026-10-06：AI/参数库加入后，「两端专用口」必须按口↔kind 矩阵精判。
  //   旧防线「两端规则口但两端都非规则节点拒绝」升级为完整矩阵
  //   （普通↔普通走专用口、rule↔ai、ai↔paramlib 等错配都在矩阵里拒绝）。
  const sem = triggerLineSemantics(src.kind, tgt.kind, sourceHandle, targetHandle)
  if (sem === null) return false
  // 只有规则线沿用「同种类型」约束（canJoinRuleGroup：首个成员不设限）；
  // AI/参数库线无此约束 —— AI 可吃任意普通节点，参数库可供给任意普通节点。
  if (sem === 'rule') return canJoinRuleGroup(source, target, tables)
  return true
}

/**
 * 这条 flow 边**是不是**圆形节点专属线、是哪一种（2026-10-06 抽离）。
 *
 * ★ 为什么要跟 roundLineAnchors 一起抽出来：
 *   两者是一对 —— 判定「是哪种线」与判定「连到哪个口」必须同源。
 *   判定此前也是 Canvas.tsx 里的一段内联三元嵌套，没法测；
 *   于是「判定按 rule→ai→paramlib 顺序短路、落位按lineKind 分派」这两处
 *   只能靠人脑对齐。上一轮参数库落点改错方向，正是这种对齐失效的结果。
 *   抽出来后可以断言「凡判定为 paramlib 的边，落位必是上1」这类**跨函数不变量**。
 *
 * 判定顺序 = 落库层 App.handleConnect 的分支优先级（rule → ai → paramlib）：
 *   任一端是 rule → 规则线；否则任一端是 ai → AI 线；
 *   否则任一端是 paramlib → 参数库线；否则不是圆形专属线（null）。
 *
 * @param edgeKind 边类型。非 'flow' 一律返回 null —— 关联边/lookup 边不参与
 * @param srcKind / tgtKind 两端节点 kind。任一端查不到（undefined）→ 返回 null
 *   （节点被删/未加载时宁可退回默认数据流口，也不要瞎派一个圆形口）
 */
export function roundLineKindOf(
  edgeKind: string | undefined,
  srcKind: string | undefined,
  tgtKind: string | undefined,
): 'rule' | 'ai' | 'paramlib' | null {
  if (edgeKind !== 'flow' || !srcKind || !tgtKind) return null
  if (srcKind === 'rule' || tgtKind === 'rule') return 'rule'
  if (srcKind === 'ai' || tgtKind === 'ai') return 'ai'
  if (srcKind === 'paramlib' || tgtKind === 'paramlib') return 'paramlib'
  return null
}

/**
 * 圆形节点专属线在画布上的**两端锚点**（2026-10-06 抽离，供测试直接断言）。
 *
 * ★★ 为什么要抽成纯函数：
 *   锚点落位此前是内联在 Canvas.tsx 的 edgesOf 里的一串 if，
 *   **没有任何测试能碰它** —— 测试只能验 canConnectPorts 的准入矩阵，
 *   而准入回答的是「能不能连」，不是「连到哪儿」。
 *   结果是：把参数库的落点从上2 改成上1，**全量测试依然全绿**，
 *   因为改的是那条无人断言的 if。
 *   这与「样式断言恒真」是同一类病：验的不是用户看到的那件事。
 *   抽出来后，锚点可被单测钉死（见 roundPort.test.ts）。
 *
 * ★ 口径来源（用户原话，逐条对应）：
 *   · 规则线   ：rule 节点下1 → 普通节点**上2**（RULE_PORT_IN_2）
 *   · AI 线    ：普通节点下1 → AI 节点上沿
 *   · 参数库线 ：参数库下沿 → 普通节点**上1**（RULE_PORT_IN）
 *     「参数库连接节点 锚点 上1」（2026-10-06；原为上2）
 *
 *   规则线走上2、参数库走上1 是**刻意错开**：同一节点被规则锁定又被参数库供参数时，
 *   两种线落在不同的口上才分得清谁是谁。叠在一个口上会视觉混淆。
 *
 * @param lineKind 线种（由 roundLineKindOf 判定）；非圆形专属线传 null
 *   → 返回空对象（让 RF 走默认数据流口，而不是落到一个圆形口上）
 * @param srcKind 边**source 端**节点的 kind
 *
 * ★★ 为什么方向由本函数自己判、而不是让调用方传一个 `roundIsSource: boolean`：
 *   第一版签名就是布尔，而 Canvas 的传参写成 `srcT.kind !== 'rule'`。
 *   那个式子只对规则线成立 —— 参数库线里 source 端恰恰**不是** rule，
 *   于是「参数库在 source 端」这个最常见的场景被判成了反向，
 *   线会连到反方向（sourceHandle/targetHandle 整体对调，RF 找不到对应口）。
 *   布尔把「哪一端是圆形节点」这个**事实**藏在了调用方的一个判断里，
 *   换个线种就错，且 TS 不会拦。现在改为只传**事实**（srcKind），
 *   方向由本函数按线种自行推 —— 三段结构对称，新增线种时不会漏改。
 *
 *   注：`lineKind` 的三个取值恰好就是三种圆形节点的 kind 值，
 *   所以「圆形节点是否在 source 端」直接就是 `srcKind === lineKind`。
 */
export function roundLineAnchors(
  lineKind: 'rule' | 'ai' | 'paramlib' | null,
  srcKind: string | undefined,
): { sourceHandle?: string; targetHandle?: string } {
  if (!lineKind) return {}
  // 规则线：规则节点下1（RULE_PORT_OUT）↔ 普通节点上2（RULE_PORT_IN_2）
  if (lineKind === 'rule') {
    return srcKind === 'rule'
      ? { sourceHandle: RULE_PORT_OUT, targetHandle: RULE_PORT_IN_2 }
      : { sourceHandle: RULE_PORT_IN_2, targetHandle: RULE_PORT_OUT }
  }
  // AI 线：普通节点下1 → AI 上沿。方向语义固定（普通节点恒为供数方）。
  // srcKind === 'ai' 属反边（AI 口是 target，用户无法从它起手拖线，
  // 真实数据里不该出现）—— 不硬套方向，按实际两端落位。
  if (lineKind === 'ai') {
    return srcKind === 'ai'
      ? { sourceHandle: AI_PORT_IN, targetHandle: RULE_PORT_OUT }
      : { sourceHandle: RULE_PORT_OUT, targetHandle: AI_PORT_IN }
  }
  // 参数库线：参数库下1（PARAMLIB_PORT_OUT）→ 普通节点上1（RULE_PORT_IN）
  return srcKind === 'paramlib'
    ? { sourceHandle: PARAMLIB_PORT_OUT, targetHandle: RULE_PORT_IN }
    : { sourceHandle: RULE_PORT_IN, targetHandle: PARAMLIB_PORT_OUT }
}

/**
 * 这条规则线该不该建立（用户 2026-10-05 第 6 条：
 * 「规则跟随节点只能连接同种类型的节点」）。
 *
 * ── 判据说明（这里踩过一个把语义搞反的坑，记下来防止重犯）───────────
 *   初版写成 `src.kind !== tgt.kind → 拒绝`。**这是把比较对象搞反了**：
 *   规则节点自己的 kind 恒为 `'rule'`（isRuleNodeKind 就是 `kind === 'rule'`），
 *   被锁节点是 'table'/'param'/'sort'/...，两者**永不相等**
 *   → 规则跟随功能被整体打死，一条线都连不上。
 *
 *   正确语义来自 effectiveRule 的实现：规则按 `ruleNode.config.rules[kind]`
 *   **分桶下发**（见本文件下方 effectiveRule）。所以「同种类型」指的是
 *   **同一个规则节点锁定的成员必须彼此同 kind** ——
 *   一个圆形节点锁了 3 个数据表格，这 3 个共享 rules.table 一份规则，
 *   主节点改了对3 个一起生效；若再锁一个「数据排序」进来，
 *   它读的是 rules.sort 桶，成员的「同步」语义就断了（规则编辑器也得分栏）。
 *
 *   且**第一个成员不设限**：组还空着时，任何 kind 都能当首个，
 *   由它定下这一组的 kind（用户从面板里换成员类型时也不该被卡住）。
 *
 * ── 为什么抽成纯函数（App 与测试共用同一份）──────────────────────────
 *   初版把这条判定分别写在 canConnectPorts 和 App.handleConnect 里，
 *   测试若照抄一份就等于没测（项目里已吃过这个亏，见 ruleConnectPatch 注释）。
 *   落库层必须再兜一道：isValidConnection 只在 UI 交互里生效，
 *   程序化调用 handleConnect 会绕过它。
 *
 * @param tables 必须全量传 —— 「组里已有哪些成员」只能全表扫（无反向索引）
 * @returns true = 允许建立这条规则锁定关系
 */
export function canJoinRuleGroup(
  source: string,
  target: string,
  tables: Record<string, TableSchema>,
): boolean {
  const src = tables[source]
  const tgt = tables[target]
  // 查不到节点 → 放行，交由落库层（ruleConnectPatch / handleConnect 的
  // `if (!src || !tgt) return`）兜底。这里若拦，会出现"拖到空白处连不上"的假死。
  if (!src || !tgt) return true
  // 两端都不是规则节点 → 这条线与规则锁定无关（原有数据流/关联连线，一行不改）
  if (!isRuleNodeKind(src.kind) && !isRuleNodeKind(tgt.kind)) return true
  // 自己连自己
  if (source === target) return false
  // 规则节点之间不互锁：互锁会成环，effectiveRule 解析时无限跳
  if (isRuleNodeKind(src.kind) && isRuleNodeKind(tgt.kind)) return false
  // 规则节点被移入回收站 → 不接受新成员（与 bindingsOfRuleNode 的回收站口径一致）
  const ruleId = isRuleNodeKind(src.kind) ? source : target
  const other = isRuleNodeKind(src.kind) ? tgt : src
  if (tables[ruleId]?.deleted || other.deleted) return false

  // 组内已有成员 → 新成员必须与它们同 kind；组空 → 首个不设限
  const members = bindingsOfRuleNode(ruleId, tables)
  if (!members.length) return true
  return members.every((m) => m.targetKind === other.kind)
}

/** 一条规则锁定关系（已解析） */
export interface RuleBinding {
  /** 规则节点 id（主） */
  ruleNodeId: string
  /** 被锁节点 id（辅） */
  targetId: string
  /** 被锁节点的 kind —— 决定取 rules 的哪个桶 */
  targetKind: string
  /**
   * 组内身份（用户：「规则节点第一个默认是可修改的主节点，
   * 连接剩下的是从节点」+「双击能修改节点的主次、跟随不跟随」）。
   * 缺省按 follow 处理 —— 见 schema.ruleRef.mode 的向后兼容说明。
   */
  mode: 'primary' | 'follow' | 'ignore'
}

/**
 * 各 kind 的规则适配体系 —— 用户：
 * 「针对不同节点的内容 适配不同的规则样式体系」。
 *
 * ★ 为什么集中这张表：
 *   若散在各个 UI 分支里写 if（sort 才能选 sortBy、map 才能选 mapKeyField），
 *   加新 kind 时必漏一处，症状是「规则项显示了但没生效」或反过来。
 *   单一真相源铁律同样适用于「规则项可见性判定」，不只适用于颜色。
 */
export const RULE_FIELDS_OF_KIND: Record<string, (keyof NodeRuleSet)[]> = {
  sort: ['sortBy', 'sortDir'],
  map: ['mapKeyField', 'mapMatchField', 'mapReturnFields'],
  text: ['textStyle'],
  output: ['exportFormat'],
  docx: ['exportFormat'],
  xlsx: ['exportFormat'],
  oosheet: ['exportFormat'],
  ooword: ['exportFormat'],
}

/** 通用视觉规则 —— 任何 kind 都可适配（用户：「针对不同节点的内容适配不同的规则样式体系」） */
export const COMMON_RULE_FIELDS: (keyof NodeRuleSet)[] = ['family', 'shape', 'scale']

/** 某 kind 可适配的全部规则项 = 通用视觉 + 专属行为 */
export function ruleFieldsOf(kind: string): (keyof NodeRuleSet)[] {
  return [...COMMON_RULE_FIELDS, ...(RULE_FIELDS_OF_KIND[kind] ?? [])]
}

/* ────────────────────────── 连线解析 ──────────────────────────
   ★ 关键约束：不另存边表。
     画布连线是从 config.linkedInputs / lookup 字段**派生**的
     （projectFormat.ts 的 edgesOf() 是单一真相源）。
     若这里另建一份lockEdges 存储，就出现两份连线真相，
     删节点/删连线时会不一致——这正是项目里吃过亏的地方。
     故直接复用同样的派生规则。
   ────────────────────────────────────────────────────────── */

/**
 * 推导一个规则节点锁定了哪些节点。
 *
 * 锁定范围 = **沿连线可达的节点**，且排除：
 *   · 自己
 *   · 回收站里的节点（与 edgesOf 一致：不产生边）
 *   · 其它规则节点（规则节点之间不互相锁定，避免环）
 *
 * @param ruleId规则节点 id
 * @param tables 全量表（要互相查表，故必须全量传）
 */
export function bindingsOfRuleNode(ruleId: string, tables: Record<string, TableSchema>): RuleBinding[] {
  const self = tables[ruleId]
  if (!self || !isRuleNodeKind(self.kind) || self.deleted) return []

  const out: RuleBinding[] = []
  const seen = new Set<string>([ruleId])

  /* 身份一律读 **被锁节点自己** 身上的 ruleRef —— 主从/跟随状态属于节点自身，
     不另存成员表（与「不另存边表」同一条理由：避免两份真相）。
     跨规则节点读取时必须校验 ruleNodeId 匹配，否则节点被两个规则组锁定时
     会串组：拿A 组的身份去解释B 组的成员。 */
  const modeOfNode = (t: TableSchema): 'primary' | 'follow' | 'ignore' => {
    const ref = t.config?.ruleRef
    if (!ref || ref.ruleNodeId !== ruleId) return 'follow'
    return ref.mode ?? 'follow'
  }

  // ① 作为连线终点：本节点挂在被锁节点的上游 —— 不构成「锁定谁」。
  //    锁定关系由**规则节点主动向下游发**，故只看成上游。
  for (const srcId of self.config?.linkedInputs ?? []) {
    if (seen.has(srcId)) continue
    const t = tables[srcId]
    if (!t || t.deleted || isRuleNodeKind(t.kind)) continue
    seen.add(srcId)
    out.push({ ruleNodeId: ruleId, targetId: srcId, targetKind: t.kind, mode: modeOfNode(t) })
  }

  // ② 作为连线源：谁把它当上游，谁被锁定。
  //    ★ 必须全量扫，因为「谁连了我」没有反向索引，只能扫。
  for (const t of Object.values(tables)) {
    if (t.deleted || seen.has(t.id) || isRuleNodeKind(t.kind)) continue
    const ups = t.config?.linkedInputs ?? []
    if (!ups.includes(ruleId)) continue
    seen.add(t.id)
    out.push({ ruleNodeId: ruleId, targetId: t.id, targetKind: t.kind, mode: modeOfNode(t) })
  }

  return out
}

/**
 * 解析**全画布**的锁定关系。
 * @returns ruleNodeId → 该规则节点锁定的所有辅节点
 */
export function allBindings(tables: Record<string, TableSchema>): Map<string, RuleBinding[]> {
  const map = new Map<string, RuleBinding[]>()
  for (const t of Object.values(tables)) {
    if (!isRuleNodeKind(t.kind) || t.deleted) continue
    map.set(t.id, bindingsOfRuleNode(t.id, tables))
  }
  return map
}

/**
 * 连线判定：这条连线是不是「规则锁定」，若是，锁定该写进哪个节点的 linkedInputs。
 *
 * ★ 为什么抽到这里（而不是内联在 App.handleConnect 里）：
 *   初版把判定内联在回调里，测试照着回调**抄了一份**canConnectRule，
 *   结果做反向验证时把真实现改回单向（有缺陷），测试**照样 14 个全绿** ——
 *   测的是副本，与真代码零关联。抄逻辑的测试等于没有测试。
 *   现在 App 与测试共用这一个函数，反向验证才能真的变红。
 *
 * ★ 为什么**对称**处理两个方向（`src.kind === 'rule' || tgt.kind === 'rule'`）：
 *   bindingsOfRuleNode 内核是双向扫描的（既扫规则节点自己的 linkedInputs，
 *   也扫辅节点里含规则节点 id 的）。只实现单向 = 内核支持双向、UI 只通一半。
 *   画布上连线本就无向，规则锁定也是，所以不该问方向。
 *
 * @returns 需写入的 { nodeId, linkedInputs, ruleRef }；null = 这条线不归规则锁定管（走别的分支）
 */
export function ruleConnectPatch(
  srcId: string,
  tgtId: string,
  tables: Record<string, TableSchema>,
): {
  nodeId: string
  linkedInputs: string[]
  /**
   * 要写入的组内身份；**undefined = 不要动它**。
   *
   * ★ 为什么要留这个「不写」的出口（单槽设计的必然结果）：
   *   ruleRef 是单槽 —— 一个节点身上只有一条「我在哪个组里是什么身份」。
   *   节点被两个规则组锁定时（本项目认这是设计歧义，见 ambiguousLocks），
   *   若这里无条件覆盖，第二个规则节点会把第一个组的身份**静默抢走**：
   *   A 组的成员表里 s1 的身份从primary 变成 follow（因为 ruleNodeId 不匹配而回落），
   *   A 组于是变成「有成员但无主节点」。
   *   症状：A 组面板显示「暂无主节点」，而 A 组明明还有人在跟着。
   *   所以已被别的组占用时**不写**，让它在A 组按缺省 follow 参与（缺省即 follow，
   *   见 roleModeOf 的向后兼容说明）—— 少一个显式身份不会丢规则，
   *   抢走身份才会。
   */
  ruleRef?: { ruleNodeId: string; mode: 'primary' | 'follow' }
} | null {
  const src = tables[srcId]
  const tgt = tables[tgtId]
  if (!src || !tgt) return null
  if (src.kind !== 'rule' && tgt.kind !== 'rule') return null
  if (srcId === tgtId) return null
  // 规则节点之间不互锁：互锁会成环，effectiveRule 解析时无限跳
  if (src.kind === 'rule' && tgt.kind === 'rule') return null
  const ruleId = src.kind === 'rule' ? srcId : tgtId
  const otherId = src.kind === 'rule' ? tgtId : srcId
  const other = src.kind === 'rule' ? tgt : src
  // 沿 linkedInputs 追加而不另存边表：删除节点时只需清理一处引用，
  // 不会出现「两份连线真相不一致」（与 projectFormat.edgesOf 同源）
  const linked = other.config?.linkedInputs ?? []
  if (linked.includes(ruleId)) return null // 重复连线不写
  // 身份已被别的规则组占用 → 本次只加连线，不碰身份（见上面 ruleRef 的注释）
  const occupied = !!other.config?.ruleRef && other.config.ruleRef.ruleNodeId !== ruleId
  return {
    nodeId: otherId,
    linkedInputs: [ruleId, ...linked],
    // ★ 第一个连上的默认是主节点（用户：「规则节点第一个默认是可修改的主节点。
    //   连接剩下的是从节点」）。判定看的是「连上之前组内是否已有成员」。
    //   不能等连完再补——那样要两次写，且中途失败会留下无主的组。
    ...(occupied
      ? {}
      : {
          ruleRef: {
            ruleNodeId: ruleId,
            mode: bindingsOfRuleNode(ruleId, tables).length === 0 ? ('primary' as const) : ('follow' as const),
          },
        }),
  }
}

/**
 * 某节点在某个规则组里的身份。
 *
 * @param override 显式指定（写操作后回读用）；不传则读节点身上的 ruleRef
 * @returns 'primary' | 'follow' | 'ignore'；未被该规则锁定则 null
 *
 * ★ 缺省是 **follow** 而不是 ignore：
 *   早期版本的连线没有 mode 字段，缺失就当 ignore 的话，
 *   升级后所有已有连线会**静默失去规则** —— 症状是「规则一夜之间全失效」，
 *   且没有任何报错。宁可多跟随（回到旧行为）也不能少跟随。
 */
export function roleModeOf(
  targetId: string,
  ruleNodeId: string,
  tables: Record<string, TableSchema>,
  override?: 'primary' | 'follow' | 'ignore',
): 'primary' | 'follow' | 'ignore' | null {
  // ★ 走 membersOf 而不是直接读 bindingsOfRuleNode：
  //   membersOf 里含「组内无primary 时的兜底接任」。若这里绕开它自己读一份，
  //   就会出现 roleModeOf 说 follow、membersOf 说 primary 的两份真相 ——
  //   症状是面板把某人标成主节点，而 effectiveRule 那边按从节点处理。
  const bind = membersOf(ruleNodeId, tables).find((b) => b.targetId === targetId)
  if (!bind) return null
  if (override) return override
  return bind.mode
}

/**
 * 这个节点的身份槽（ruleRef 单槽）此刻归哪个规则组所有。
 *
 * ★ 这是「单槽设计」必须配套的读接口。ruleRef 只能存一个 ruleNodeId，
 *   所以一个节点被两个规则组锁定时，**只有一个组能改它的身份**。
 *   没有这道判定，写操作会静默抢走另一个组的身份：
 *   在 B 组把 s1 设为「不跟随」→ 写 ruleRef={rB,ignore} → A 组的主节点身份没了，
 *   A 组于是变成「有成员但无主」。症状是 A 组面板忽然显示别人是主节点，
 *   且A 组的规则不再以 s1 为锚 —— 跨组数据损坏，且极难定位。
 *
 * @returns 占用者规则节点 id；未被占用（槽空或就是本组）返回 null
 */
export function identityOwnerOf(
  targetId: string,
  tables: Record<string, TableSchema>,
  /** 传入则排除「占用者就是自己」这一情况（本组的槽当然算可用） */
  forRuleNodeId?: string,
): string | null {
  const ref = tables[targetId]?.config?.ruleRef
  if (!ref?.ruleNodeId) return null
  if (forRuleNodeId && ref.ruleNodeId === forRuleNodeId) return null
  return ref.ruleNodeId
}

/**
 * 组内成员，**主节点排第一个**（用户：「规则节点第一个默认是可修改的主节点」）。
 *
 * ⚠「第一个」不是「最早上线的那个」—— 主节点是**可改的**（双击面板里能换人），
 *   所以必须显式按 mode 排，不能拿插入顺序当主次。
 *
 * ★★ 兜底：组内若**没有**显式 primary，且存在身份未被别的组占用的成员，
 *   把第一个这样的成员当作主节点返回。
 *
 *   为什么需要：节点被两个规则组锁定时（内核判为设计歧义，见 ambiguousLocks），
 *   后连的组里成员身份要么被前一个组占用、要么没写进去，组内就一个显式 primary 都没有。
 *   若不兜底，症状是面板显示「暂无主节点」而组里明明有人跟着，
 *   且圆形节点的规则编辑区仍可点 —— 点了没人执行，正是「改了没反应」。
 *
 *   为什么只挑「未被占用」的：占用中的成员在本组**改不动身份**（见 identityOwnerOf），
 *   把它标成主节点会出现「主节点的按钮全是灰的」—— 比显示无主更难懂。
 *   若全组都被占用，就老实返回无主，由 UI 提示这是重叠锁定。
 *
 *   为什么兜底只在这里做、不写回 store：写回就得在每次连线/取消/删除时同步修，
 *   漏一处就出现「面板说有主、实际存的是 follow」的两份真相。纯派生则永远一致。
 *   effectiveRule 不依赖 primary（只看 ignore），故这个兜底纯影响 UI 呈现。
 *
 * @param includeIgnored 是否把「不跟随」成员也列进来。UI 要显示（要能切回来），
 *        effectiveRule 那类计算不要（它得跳过 ignore）。
 */
export function membersOf(
  ruleNodeId: string,
  tables: Record<string, TableSchema>,
  includeIgnored = true,
): RuleBinding[] {
  const all = bindingsOfRuleNode(ruleNodeId, tables)
  const list = includeIgnored ? all : all.filter((b) => b.mode !== 'ignore')
  /*排序：主节点必须在**第一行**（用户 2026-10-05 第 4 条：
     「双击规则跟随节点 打开的节点列表 主节点位于最上方」）。

     ★ 为什么 rank 相同还要再排 name：bindingsOfRuleNode 的返回顺序取决于
       「linkedInputs 数组顺序」与「Object.values(tables) 的键顺序」，
       两者都不保证稳定。用户看到的主节点不在最上方，根因就是这里：
       只有一个 primary 时 rank 能把它排到第一，但一旦有**两个** primary
       （历史数据/多组串味），谁的 mode 都是 primary，谁在前面完全随机。
       加 name 做次级键让顺序确定下来，主节点至少稳定在最上方那批的第一位。*/
  const sorted = [...list].sort(
    (a, b) => rank(a.mode) - rank(b.mode) || a.targetId.localeCompare(b.targetId),
  )
  if (!sorted.some((b) => b.mode === 'primary')) {
    // 第一个「非 ignore 且身份未被别的组占用」的成员接任。
    // 用 map 生成新对象而非就地改：调用方可能把结果缓存下来，就地改会串味。
    const heirId = sorted.find((b) => b.mode !== 'ignore' && !identityOwnerOf(b.targetId, tables, ruleNodeId))?.targetId
    if (heirId) return sorted.map((b) => (b.targetId === heirId ? { ...b, mode: 'primary' as const } : b))
  }
  return sorted
}
const rank = (m: 'primary' | 'follow' | 'ignore') => (m === 'primary' ? 0 : m === 'follow' ? 1 : 2)

/** 组内主节点（用户：「第一个默认是可修改的主节点」）。无成员时 null。 */
export function primaryOf(
  ruleNodeId: string,
  tables: Record<string, TableSchema>,
): RuleBinding | null {
  return membersOf(ruleNodeId, tables).find((b) => b.mode === 'primary') ?? null
}

/** 该规则节点上已挂了多少个节点 —— 圆形节点正中央显示的数字（「规则节点上显示几个节点」）。 */
export function memberCountOf(ruleNodeId: string, tables: Record<string, TableSchema>): number {
  return bindingsOfRuleNode(ruleNodeId, tables).length
}

/** 圆形节点正中央那一行：「3」或「3 · 主 排序A」。 */
export function memberSummaryOf(
  ruleNodeId: string,
  tables: Record<string, TableSchema>,
): { count: number; primaryName: string | null } {
  const binds = membersOf(ruleNodeId, tables)
  const p = binds.find((b) => b.mode === 'primary')
  return {
    count: binds.length,
    primaryName: p ? (tables[p.targetId]?.name ?? null) : null,
  }
}

/**
 * 反查：一个节点被哪些规则节点锁定。
 * @returns 规则节点 id 列表。无则空数组。
 */
export function ruleNodesOf(targetId: string, tables: Record<string, TableSchema>): string[] {
  const out: string[] = []
  for (const [ruleId, binds] of allBindings(tables)) {
    if (binds.some((b) => b.targetId === targetId)) out.push(ruleId)
  }
  return out
}

/**
 * 取某节点**实际生效**的规则。
 *
 * ★ 优先级：规则节点的规则 > 节点自身 config（自身配置是更底层的默认值）。
 *   规则节点删掉后，节点自动回落用自己的 config，不会规则裸奔。
 *
 * @returns 规则；未被锁定且无配置时返回 null
 */
export function effectiveRule(target: TableSchema, tables: Record<string, TableSchema>): NodeRuleSet | null {
  const owners = ruleNodesOf(target.id, tables)
  if (!owners.length) return target.config ?? null

  // 一节点被多个规则节点锁定时，取后写入的那个（后设优先，与 setNodeConfig 语义一致）。
  // 实际使用中不该出现——两个圆形节点锁同一个节点是设计上的歧义，UI 应给警告。
  const ruleNodeId = owners[owners.length - 1]
  // ★ 显式「不跟随」的节点不受该规则影响（用户：双击可切跟随/不跟随）。
  //   注意这里判的是**生效规则**，不是 isRuleLocked —— 连线还在、只是不跟随，
  //   UI 上仍要显示锁定标记（那是「连着」的事实），但规则不生效。
  if (roleModeOf(target.id, ruleNodeId, tables) === 'ignore') return target.config ?? null
  const ruleNode = tables[ruleNodeId]
  const rule = ruleNode?.config?.rules?.[target.kind]
  if (!rule) return target.config ?? null

  // 规则节点的规则覆盖自身 config，未涉及的项回落自身（浅合并，不做深合并：
  // 深合并会让「删除某个规则项」永远删不掉）
  return { ...target.config, ...rule }
}

/**
 * 某节点是否处于「被锁定只读」状态 —— UI 据此禁用规则编辑区。
 *
 * 只读范围与写入路径必须同源：effectiveRule 走 ruleNodesOf，
 * 这里也走同一个函数。两处各判一次必然漂移。
 */
export function isRuleLocked(targetId: string, tables: Record<string, TableSchema>): boolean {
  return ruleNodesOf(targetId, tables).length > 0
}

/**
 * 校验一份规则对某 kind 是否完整 —— 缺必填项时 UI 应标黄而非静默接受。
 *
 * 只对「选了值就应生效」的字段校验；family/shape/scale 有默认值，不算缺失。
 */
export function validateRule(rule: NodeRuleSet, kind: string): { field: string; msg: string }[] {
  const problems: { field: string; msg: string }[] = []
  const has = (k: keyof NodeRuleSet) => rule[k] !== undefined && rule[k] !== ''

  if (kind === 'map') {
    if (!has('mapKeyField')) problems.push({ field: 'mapKeyField', msg: '对照节点必须选「用上游哪个字段去查」' })
    if (!has('mapMatchField')) problems.push({ field: 'mapMatchField', msg: '对照节点必须选「用本节点哪个字段匹配」' })
  }
  if (kind === 'sort') {
    if (!has('sortBy')) problems.push({ field: 'sortBy', msg: '排序节点必须选排序列' })
    if (!has('sortDir')) problems.push({ field: 'sortDir', msg: '排序节点必须选升降序' })
  }
  if (rule.scale !== undefined && (rule.scale < 0.5 || rule.scale > 2)) {
    problems.push({ field: 'scale', msg: '缩放倍率需在 0.5~2 之间' })
  }
  return problems
}

/**
 * 锁定歧义检测 —— 两个规则节点锁同一个节点。
 * 属设计歧义不是崩溃，但必须能看见，否则「改了没生效」无从排查。
 */
export function ambiguousLocks(tables: Record<string, TableSchema>): { targetId: string; ruleNodeIds: string[] }[] {
  const owners = new Map<string, string[]>()
  for (const [ruleId, binds] of allBindings(tables)) {
    for (const b of binds) {
      const arr = owners.get(b.targetId) ?? []
      arr.push(ruleId)
      owners.set(b.targetId, arr)
    }
  }
  return [...owners.entries()].filter(([, ids]) => ids.length > 1).map(([targetId, ruleNodeIds]) => ({ targetId, ruleNodeIds }))
}

/** 规则节点 kind 常量，避免各处硬编码字符串 */
export const RULE_NODE_KIND: NodeKind = 'rule'

/**
 * 圆形节点（AI / 参数库）的连线计数 —— 圆内正中显示的数字。
 *
 * 数法（与 bindingsOfRuleNode 同一个「双向扫描 linkedInputs」的口径）：
 *   · 作为下游：我的 linkedInputs 里有几个有效上游（AI 吃了几个节点）
 *   · 作为上游：几个有效节点把我记进它们的 linkedInputs（参数库供了几个节点）
 *
 * ★ 为什么不复用 memberSummaryOf：那是规则锁定口径（含主从身份、
 *   排除其它规则节点成员），AI/参数库连线是纯数据流式关系，
 *   没有身份槽 —— 两套语义不能共用一个数法，否则圆内数字口径漂移。
 *
 * 规则节点（rule）不用这个函数 —— 它的圆内计数仍走 memberSummaryOf。
 */
export function roundLinkCountOf(nodeId: string, tables: Record<string, TableSchema>): number {
  const self = tables[nodeId]
  if (!self || self.deleted) return 0
  let n = 0
  // ① 作为下游：我连了谁（AI 的情形）
  for (const srcId of self.config?.linkedInputs ?? []) {
    const t = tables[srcId]
    if (t && !t.deleted) n++
  }
  // ② 作为上游：谁连了我（参数库的情形）。无反向索引，只能全表扫（同 bindingsOfRuleNode）
  for (const t of Object.values(tables)) {
    if (t.deleted || t.id === nodeId) continue
    if ((t.config?.linkedInputs ?? []).includes(nodeId)) n++
  }
  return n
}