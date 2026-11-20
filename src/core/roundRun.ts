// ════════════════════════════════════════════════════════════════════════
// 圆形节点运行动画的状态机（2026-10-06）
// ════════════════════════════════════════════════════════════════════════
// 用户原话（本轮两条）：
//   1「对选中的节点点击动态特效才会有效果」
//   2「增加特效的结束的退出效果 测试的时候3秒结束」
//
// ── 为什么要有这个文件 ──────────────────────────────────────────────
//   上一版的动画条件是 `isRound && !isRule && (nodesFolded || roundRunning)`，
//   埋在 Canvas.tsx 的 TableNode 里。结果点一下播放按钮，
//   **全画布所有 AI/参数库节点一起转圈**，而用户要的是
//   「只有我正点着的那个节点在动」。
//   更麻烦的是「3 秒后进入退出态」这种带时序的逻辑 —— 它一旦内联在组件里，
//   就没有任何测试能碰它：改时长、改状态迁移，全量测试依然全绿。
//   （这与「样式断言恒真」同类：验的不是用户看到的那件事。）
//   故把三件事抽成纯函数：
//     · roundRunStateOf   —— 某个节点此刻该不该播（判定）
//     · 状态机迁移         —— running → exiting → off 的时序（纯逻辑，不碰定时器）
//     · 时长常量           —— 单一真相源，CSS 与逻辑共用
//
// ── 状态机（3 秒是「测试档」时长）────────────────────────────────────
//   off ──播放──> running ──3s 到──> exiting ──0.5s──> off
//    ▲                                │
//    └──────────── 手动停止 ────────────┘
//   ★ 为什么退出态要独立成一个状态，而不是直接置false：
//     直接置 false = 动画元素被立刻移除，用户看到的是「啪」一下消失。
//     必须让 ::before/::after 有机会播完一段淡出+收缩，才叫「退出效果」。
//     这是纯 CSS 做不到的 —— class 一摘，元素就没了，动画无从播完。
//
// ── 为什么时长是常量而不是散在 CSS 里 ────────────────────────────────
//   退出动画的 CSS 时长若写在 index.css、而 JS 里的超时写 500ms，
//   两者一旦漂移就是「class 摘早了，淡出只播一半」。
//   故把时长放这里，CSS 侧用同名变量对齐（见 ROUND_RUN_EXIT_MS 注释）。
// ════════════════════════════════════════════════════════════════════════

/** 运行动画的三个状态 */
export type RoundRunState = 'off' | 'running' | 'exiting'

/** 测试档：播放时长（用户：「测试的时候3秒结束」） */
export const ROUND_RUN_TEST_MS = 3000

/** 退出淡出时长。★ 必须与 index.css 里 rf-round-exit 的 animation 时长一致 */
export const ROUND_RUN_EXIT_MS = 500

/**
 * 某个圆形节点此刻该不该挂运行动画。
 *
 * ★ 用户口径（2026-10-06 第1条）：**只对选中的节点**播。
 *   上一版是 `nodesFolded || roundRunning` —— 折叠或点播放，全画布圆形节点一起亮。
 *   现在三个条件同时成立才播：
 *     ① 是 AI/参数库（**不是** rule —— 用户口径只有 AI/参数库有这套动画）
 *     ② 它正被选中
 *     ③ 状态不是 off（播放中 running，或退出淡出中 exiting）
 *
 * @param kind 节点 kind
 * @param selected 该节点是否被选中
 * @param state 当前运行状态
 * @param folded 是否处于全局折叠态。
 *   ★ 保留这个形参不是为了参与判定，而是**为了把「折叠不再自动播」这件事
 *   钉在签名里**：曾经的 bug 正是 `(nodesFolded || roundRunning)`。
 *   留着它 + 测试里显式传 true 并断言不播，比悄悄删掉形参更防回退 ——
 *   删掉形参后，下次重构很容易又写回`folded || running`。
 */
export function roundRunStateOf(
  kind: string,
  selected: boolean,
  state: RoundRunState,
  folded: boolean,
): RoundRunState {
  void folded // 折叠**不参与**判定（用户 2026-10-06 第1条：只对选中的节点播）
  if (state === 'off') return 'off'
  if (kind === 'rule') return 'off' // 规则节点不挂这套动画（用户口径只有 AI/参数库）
  if (kind !== 'ai' && kind !== 'paramlib') return 'off'
  if (!selected) return 'off'
  return state
}

/**
 * 状态机迁移：给定当前状态与一个事件，返回下一状态。
 * 纯函数，**不碰定时器** —— 定时由调用方（Canvas）按返回结果安排。
 *
 * ★★ 事件必须按「两个不同的时长」分开，不能让调用方拿同一个 `timeout`
 *   去表达 3 秒播放到点与 0.5 秒退出到点：
 *   合并成一种事件后，「exiting 收到 timeout 该干什么」就没有唯一答案了
 *   —— 我第一版就写成了 `cur === 'running' ? 'exiting' : cur`，
 *   结果 exiting 收到退出超时**原地不动**，退出动画永远不结束
 *   （本轮真bug：4 条用例同时变红才抓到）。
 *   事件名各自对应唯一时长，状态机就没有歧义分支。
 *
 * @param ev
 *   · 'play'按钮点播放
 *   · 'runElapsed'  播放时长到（ROUND_RUN_TEST_MS = 3s）
 *   · 'stop'        手动停止
 *   · 'exitElapsed' 退出淡出到（ROUND_RUN_EXIT_MS = 0.5s）
 */
export function nextRoundRunState(
  cur: RoundRunState,
  ev: 'play' | 'runElapsed' | 'stop' | 'exitElapsed',
): RoundRunState {
  switch (ev) {
    case 'play':
      // 已在退出淡出中再点播放 → 回到运行（重新开始，不叠加退出动画）
      return 'running'
    case 'runElapsed':
      // 只有运行中才进退出；已退出/未播的收到保持不变（防重复定时器叠加）
      return cur === 'running' ? 'exiting' : cur
    case 'stop':
      // 手动停止：运行中 → 退出淡出（也要有退出效果，不能啪一下消失）
      return cur === 'running' ? 'exiting' : cur
    case 'exitElapsed':
      // ★ 退出淡出到点 → 必须回 off，否则退出动画永不结束（本轮修的真 bug）
      return cur === 'exiting' ? 'off' : cur
  }
}

/**
 * 该状态是否应挂class（把「挂不挂」与「挂哪个」分开）。
 * 退出态也挂 —— class 摘掉动画就没了，退出效果无从播（见文件头说明）。
 */
export function roundRunClassOf(state: RoundRunState): string {
  if (state === 'running') return ' rf-round-run'
  if (state === 'exiting') return ' rf-round-run rf-round-exit'
  return ''
}

/**
 * 当前有没有「可播动画的选中节点」—— 顶栏按钮据此置灰。
 *
 * ★ 为什么要单独一个函数（而不是在 CanvasRibbon 里写个 `length > 0`）：
 *   按钮置灰是**用户唯一能看到的指引** —— 没置灰而点击又毫无反应时，
 *   用户只会以为按钮坏了。判定一旦和 UI 混在一处，就没法单测「什么时候该亮」。
 *   与 roundRunStateOf 同源（同一份 kind 口径），避免两处各判一遍。
 *
 * @param selectedKinds 选中节点的 kind 列表（RF 的 onSelectionChange 给的是节点数组）
 */
export function hasPlayableSelection(selectedKinds: readonly string[]): boolean {
  return selectedKinds.some((k) => k === 'ai' || k === 'paramlib')
}
