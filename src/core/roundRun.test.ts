// ════════════════════════════════════════════════════════════════════════
// 圆形节点运行动画 · 状态机与判定（2026-10-06）
// ════════════════════════════════════════════════════════════════════════
// 用户原话：
//   1「对选中的节点点击动态特效才会有效果」
//   2「增加特效的结束的退出效果 测试的时候3秒结束」
//
// ── 这组测试防的是上一版的真缺陷 ──────────────────────────────────────
//   旧条件写在 Canvas.tsx 里：`isRound && !isRule && (nodesFolded || roundRunning)`
//   —— 点一下播放按钮，**全画布所有 AI/参数库节点一起转圈**。
//   而用户要的是「只有我点着的那个在动」。
//   这个缺陷能溜过去，是因为它内联在组件里、没有对应测试：
//   全量 793 个用例全绿，而产品行为是错的。
//   本组把「谁该播」钉成断言，改为纯函数后即可直接验。
//
// ★ 另一类要防的：把 `folded` 又写回判定。
//   折叠态自动播是上一版的行为，用户 2026-10-06 第1条把它收窄掉了。
//   所以 roundRunStateOf 特意保留 folded 形参并在测试里显式传 true
//   断言「折叠也不播」—— 删掉形参反而容易让下次重构写回去。
// ════════════════════════════════════════════════════════════════════════

import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

import {
  roundRunStateOf,
  nextRoundRunState,
  roundRunClassOf,
  hasPlayableSelection,
  ROUND_RUN_TEST_MS,
  ROUND_RUN_EXIT_MS,
  type RoundRunState,
} from './roundRun'

describe('时长常量：单一真相源', () => {
  it('测试档播放时长 = 3 秒（用户：「测试的时候3秒结束」）', () => {
    expect(ROUND_RUN_TEST_MS).toBe(3000)
  })
  it('退出淡出时长 > 0 且短于播放时长（退出不能比运行还长）', () => {
    expect(ROUND_RUN_EXIT_MS).toBeGreaterThan(0)
    expect(ROUND_RUN_EXIT_MS).toBeLessThan(ROUND_RUN_TEST_MS)
  })
  it('★ 时长必须与 index.css 的退出动画对齐（漂移 = 淡出只播一半）', () => {
    // 不去读 CSS 文件（本测试环境无 DOM/fs），但把这条约束写成断言：
    // 若有人改了 CSS 里的 rf-round-exit 时长而忘了这里，
    // 至少这条测试提醒他去同步。CSS 侧时长在 rf-round-exit 上有同名注释。
    expect(ROUND_RUN_EXIT_MS).toBe(500)
  })
})

describe('判定：roundRunStateOf（★ 只对选中的节点播）', () => {
  it('★ 选中的 AI 节点在 running 时播', () => {
    expect(roundRunStateOf('ai', true, 'running', false)).toBe('running')
  })
  it('★ 选中的参数库节点在 running 时播', () => {
    expect(roundRunStateOf('paramlib', true, 'running', false)).toBe('running')
  })
  it('★★ 未选中的节点一律不播（这条就是用户第1条，也是旧版的缺陷所在）', () => {
    // 旧版这里会返回 running —— 判据是 (nodesFolded || roundRunning)，
    // 完全不看 selected，于是全画布圆形节点一起转圈。
    expect(roundRunStateOf('ai', false, 'running', false)).toBe('off')
    expect(roundRunStateOf('paramlib', false, 'running', false)).toBe('off')
    // 退出淡出中同理：未选中就不该有淡出动画
    expect(roundRunStateOf('ai', false, 'exiting', false)).toBe('off')
  })

  it('★ 折叠态下即便 running 也不播（第1条把「折叠自动播」收窄掉了）', () => {
    // folded 传 true 是关键：这一条挡住「顺手把 folded || running 写回去」
    expect(roundRunStateOf('ai', true, 'running', true)).toBe('running') // 选中仍然播
    expect(roundRunStateOf('ai', false, 'running', true)).toBe('off') // 没选中不播
    expect(roundRunStateOf('paramlib', false, 'running', true)).toBe('off')
  })

  it('★ 规则节点永不播（用户口径：这套动画只给 AI/参数库）', () => {
    // 即便「选中 + running」也不播。rule 有自己的 rf-flowing 连线流动表达。
    expect(roundRunStateOf('rule', true, 'running', false)).toBe('off')
    expect(roundRunStateOf('rule', true, 'exiting', false)).toBe('off')
  })

  it('普通节点永不播（table/param/sort… 一个都不行）', () => {
    for (const k of ['table', 'param', 'sort', 'merge', 'output', 'compare', 'start', 'word', 'excel', '章节工作簿']) {
      expect(roundRunStateOf(k, true, 'running', false), `${k} 不该播`).toBe('off')
    }
  })

  it('未知 kind 不播（宁可少播，不可给错节点开动画）', () => {
    expect(roundRunStateOf('futureKind', true, 'running', false)).toBe('off')
    expect(roundRunStateOf('', true, 'running', false)).toBe('off')
  })

  it('off 状态下无论选中与否都不播', () => {
    expect(roundRunStateOf('ai', true, 'off', false)).toBe('off')
    expect(roundRunStateOf('ai', false, 'off', false)).toBe('off')
  })
})

describe('状态机：nextRoundRunState（3 秒后进退出，退出后归位）', () => {
  it('★ 完整时序 off → running →(3s 到) exiting →(0.5s) off', () => {
    let s: RoundRunState = 'off'
    s = nextRoundRunState(s, 'play')
    expect(s).toBe('running')
    s = nextRoundRunState(s, 'runElapsed') // 3 秒播放到点
    expect(s).toBe('exiting')
    s = nextRoundRunState(s, 'exitElapsed') // 0.5 秒退出到点
    expect(s).toBe('off')
  })

  it('★ 退出态必须存在，不能从 running 直接回 off（第2条：要有退出效果）', () => {
    // 少了 exiting 这一跳，class 立刻被摘 → 元素移除 → 用户看到「啪」一下消失
    expect(nextRoundRunState('running', 'runElapsed')).toBe('exiting')
  })

  it('手动停止也走退出淡出（不是直接消失）', () => {
    expect(nextRoundRunState('running', 'stop')).toBe('exiting')
  })

  it('重复 stop 幂等：已在 exiting 再 stop 仍是 exiting（不会跳回 running）', () => {
    expect(nextRoundRunState('exiting', 'stop')).toBe('exiting')
    expect(nextRoundRunState('exiting', 'exitElapsed')).toBe('off')
  })

  it('★ 重复事件不会叠加（防多个定时器把状态推乱）', () => {
    // 这是真会发生的：用户连点两次播放 → 两个 3s 定时器 → 第一个 runElapsed
    // 把状态推到 exiting，第二个再补一刀；若实现不幂等就会乱跳。
    const s1 = nextRoundRunState('running', 'runElapsed')
    expect(s1).toBe('exiting')
    // 第二个播放定时器到点：已在 exiting，不该把它推回 running（也不能当退出到点）
    expect(nextRoundRunState(s1, 'runElapsed')).toBe('exiting')
    // 退出到点：回 off
    expect(nextRoundRunState(s1, 'exitElapsed')).toBe('off')
  })

  it('★ exiting 收到 exitElapsed 必须回 off（否则退出动画永不结束）', () => {
    // 本轮真 bug：第一版把两个时长都叫timeout，写成
    // `cur === 'running' ? 'exiting' : cur` —— exiting 收到退出超时**原地不动**，
    // 状态永久卡住，退出动画播完也不消失。4 条用例同时变红才抓到。
    expect(nextRoundRunState('exiting', 'exitElapsed')).toBe('off')
    expect(nextRoundRunState('exiting', 'exitElapsed')).toBe('off')
  })

  it('off 收到任意事件都保持 off（空转事件不产生状态）', () => {
    expect(nextRoundRunState('off', 'runElapsed')).toBe('off')
    expect(nextRoundRunState('off', 'exitElapsed')).toBe('off')
    expect(nextRoundRunState('off', 'stop')).toBe('off')
  })

  it('退出淡出中点播放 → 回到 running（重新开始，不叠加退出）', () => {
    expect(nextRoundRunState('exiting', 'play')).toBe('running')
  })

  it('运行中点播放 → 仍是 running（幂等，不会重置成别的态）', () => {
    expect(nextRoundRunState('running', 'play')).toBe('running')
  })
})

describe('class 映射：roundRunClassOf', () => {
  it('running → 只挂 rf-round-run', () => {
    expect(roundRunClassOf('running')).toBe(' rf-round-run')
  })
  it('★ exiting → 同时挂 rf-round-run 与 rf-round-exit（退出动画需要底层元素还在）', () => {
    const c = roundRunClassOf('exiting')
    expect(c).toContain('rf-round-run')
    expect(c).toContain('rf-round-exit')
  })
  it('off → 空字符串（class 摘干净，元素不残留）', () => {
    expect(roundRunClassOf('off')).toBe('')
  })
  it('★ 三态互不相同的 class（防两个态撞成同一个，退出效果形同虚设）', () => {
    const all = new Set<RoundRunState>(['off', 'running', 'exiting'])
    const classes = [...all].map(roundRunClassOf)
    expect(new Set(classes).size).toBe(3)
  })
})

describe('端到端不变量：选中 + 状态 → 最终 class', () => {
  it('★ 串起判定与 class：选中 AI 播放 → 挂 run；不选中 → 什么都不挂', () => {
    const sel = (selected: boolean) =>
      roundRunClassOf(roundRunStateOf('ai', selected, 'running', false))
    expect(sel(true)).toBe(' rf-round-run')
    expect(sel(false)).toBe('')
  })

  it('★ 全流程走一遍：选中 → 播 → 3s 后退出 → 归位，class 依次正确', () => {
    const selected = true
    let s: RoundRunState = 'off'
    const cls = () => roundRunClassOf(roundRunStateOf('ai', selected, s, false))

    expect(cls()).toBe('')
    s = nextRoundRunState(s, 'play')
    expect(cls()).toBe(' rf-round-run')
    s = nextRoundRunState(s, 'runElapsed')
    // 退出中：仍有底层 class + 退出 class
    expect(cls()).toBe(' rf-round-run rf-round-exit')
    s = nextRoundRunState(s, 'exitElapsed')
    expect(cls()).toBe('')
  })

  it('★ 取消选中立即停（哪怕状态还是 running）——退出效果不挂到别人身上', () => {
    // 反向用例：若判定漏了 selected，退出淡出会跟着动画跑到别的节点上
    expect(roundRunClassOf(roundRunStateOf('ai', false, 'exiting', false))).toBe('')
  })
})

// ════════════════════════════════════════════════════════════════════════
// 时长对齐：JS 常量 vs CSS 动画（真读文件，不是断言常量等于 500）
// ════════════════════════════════════════════════════════════════════════
// ★ 为什么要真读 CSS：
//   只断言 `ROUND_RUN_EXIT_MS === 500` 是**恒真**的 —— 它只锁住了 JS 侧，
//   改 CSS 里的 0.5s → 0.8s 不会有任何测试变红，而后果是
//   「JS 500ms 摘class、动画还要再播 300ms」= 退出效果被砍一半。
//   两处时长是同一语义的两份定义（陷阱 9.40：必然漂移），
//   要么单一真相源，要么用测试把两份钉在一起。
//   这里选后者：CSS 语法上无法 import 常量，但可以**读出来比对**。
//
// 环境无 DOM，用 node:fs 直读 —— vitest 默认 environment 是 node，可读文件。
// ════════════════════════════════════════════════════════════════════════
describe('时长对齐：CSS 退出动画时长必须与 JS 常量一致', () => {
  const cssPath = join(dirname(fileURLToPath(import.meta.url)), '..', 'index.css')
  const css = readFileSync(cssPath, 'utf8')

  it('先确认真的读到了 CSS（防路径写错导致后面全是空断言）', () => {
    expect(css.length).toBeGreaterThan(1000)
    expect(css).toContain('rf-round-exit')
  })

  it('★ rf-round-exit 的 animation 时长 === ROUND_RUN_EXIT_MS', () => {
    // 抓 rf-round-exit 规则块里出现的所有时长（0.5s / 500ms 都认）
    const m = css.match(
      /\.rf-node\.rf-round-run\.rf-round-exit::before\s*\{[^}]*?animation:[^;]*?([\d.]+)s\b/,
    )
    expect(m, '未找到 rf-round-exit::before 的 animation 时长').not.toBeNull()
    const cssSec = Number(m![1])
    const jsSec = ROUND_RUN_EXIT_MS / 1000
    // 允许 0.05s 误差（CSS 写 0.5s 与 0.50s 视为一致）
    expect(Math.abs(cssSec - jsSec)).toBeLessThan(0.05)
  })

  it('★ rf-round-exit 必须与 rf-round-run 同时存在（否则元素被重建，动画从错帧开始）', () => {
    // 选择器写成 `.rf-round-run.rf-round-exit` 而不是单独的 `.rf-round-exit`
    expect(css).toContain('.rf-node.rf-round-run.rf-round-exit::before')
    expect(css).toContain('.rf-node.rf-round-run.rf-round-exit::after')
  })

  it('★ 退出动画必须 forwards 填充（否则最后一帧回到初始，闪一下才消失）', () => {
    expect(css).toMatch(/rf-round-exit-in[^}]*?forwards/)
    expect(css).toMatch(/rf-round-exit-fade[^}]*?forwards/)
  })

  it('★ reduced-motion 下退出也要有覆盖（否则退化为「什么都不发生」）', () => {
    const reduced = css.slice(css.indexOf('@media (prefers-reduced-motion: reduce)'))
    expect(reduced).toContain('rf-round-exit')
  })
})

describe('按钮置灰判定：hasPlayableSelection', () => {
  it('选中 AI / 参数库 → 可播（按钮可点）', () => {
    expect(hasPlayableSelection(['ai'])).toBe(true)
    expect(hasPlayableSelection(['paramlib'])).toBe(true)
  })
  it('★ 只选中普通节点 → 按钮置灰（点了没反应不如置灰）', () => {
    expect(hasPlayableSelection(['table'])).toBe(false)
    expect(hasPlayableSelection(['param', 'sort', 'output'])).toBe(false)
  })
  it('★ 规则节点选中 → 置灰（用户口径：这套动画只给 AI/参数库）', () => {
    expect(hasPlayableSelection(['rule'])).toBe(false)
  })
  it('空选区 → 置灰', () => {
    expect(hasPlayableSelection([])).toBe(false)
  })
  it('多选里含一个 AI 即可播（不必全选）', () => {
    expect(hasPlayableSelection(['table', 'ai', 'rule'])).toBe(true)
  })
  it("空串 kind（节点查不到时 Canvas 传空串）→ 不算可播", () => {
    expect(hasPlayableSelection([''])).toBe(false)
    expect(hasPlayableSelection(['', 'table'])).toBe(false)
  })
  it('★ 与播放判定同源：能播的选中集合，播的时候一定有节点真的挂上 class', () => {
    // 跨函数不变量：hasPlayableSelection 说「可播」时，
    // 至少要有一个 kind 真的会被 roundRunStateOf 判成 running。
    const kinds = ['table', 'rule', 'ai', 'paramlib', 'sort', 'output', '']
    for (const k of kinds) {
      const playable = hasPlayableSelection([k])
      const actuallyRun = roundRunClassOf(roundRunStateOf(k, true, 'running', false)) !== ''
      expect(playable, `kind=${k} 置灰判定与播放判定不一致`).toBe(actuallyRun)
    }
  })
})
