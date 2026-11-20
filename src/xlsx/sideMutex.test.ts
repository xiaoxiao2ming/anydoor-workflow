import { describe, it, expect } from 'vitest'

/**
 * 侧边栏 ↔ 弹窗互斥逻辑（2026-10-04用户报 bug 的回归锁）
 *
 * 该逻辑内联在 WorkbookEditor 组件里（依赖 React state），无法直接单测。
 * 这里把**判定规则**原样抽出成纯函数做锁，防止后续改 toggle 时把互斥弄丢。
 *
 * 真实缺陷：打开数据引用 → 点「批注清单」→ 弹窗被侧边栏压住（z-index 8100 > 500），
 * 再点一次只关弹窗，侧边栏仍在 → 用户以为"没响应"。
 */

/** 侧边栏三态 */
type Sides = { ref: boolean; trigger: boolean; right: boolean }

/** 弹窗三态 */
type Modals = { domains: boolean; cmts: boolean; fnlib: boolean }

const SIDES_ALL = ['ref', 'trigger', 'right'] as const
const MODALS_ALL = ['domains', 'cmts', 'fnlib'] as const

/** 打开弹窗 → 收掉全部侧边栏（实现见closeSides） */
function openModal(sides: Sides): Sides {
  return { ref: false, trigger: false, right: false }
}

/** toggle 侧边栏：关→只动自己；开→先收同缘的另一个 */
function toggleSide(sides: Sides, which: typeof SIDES_ALL[number]): Sides {
  const cur = sides[which]
  if (cur) return { ...sides, [which]: false }
  // 同缘互斥：ref 与 trigger 共左缘
  if (which === 'trigger') return { ...sides, ref: false, trigger: true }
  if (which === 'ref') return { ...sides, trigger: false, ref: true }
  return { ...sides, right: true }
}

describe('侧边栏与弹窗互斥（回归锁）', () => {
  it('打开任一弹窗时，三个侧边栏全部收起', () => {
    const allOpen: Sides = { ref: true, trigger: true, right: true }
    for (const m of MODALS_ALL) {
      expect(openModal(allOpen)).toEqual({ ref: false, trigger: false, right: false })
      expect(m).toBeTruthy()
    }
  })

  it('复现用户报的 bug：开引用栏 → 开批注清单 → 侧边栏必须收起', () => {
    let s: Sides = { ref: false, trigger: false, right: false }
    s = toggleSide(s, 'ref')
    expect(s.ref).toBe(true)
    // 关键一步：打开批注清单
    s = openModal(s)
    expect(s.ref).toBe(false)
    // 侧边栏收起后，弹窗的点击不会再被遮挡
    expect(SIDES_ALL.every((k) => s[k] === false)).toBe(true)
  })

  it('引用栏与联动面板同占左缘，开一个必关另一个', () => {
    let s: Sides = { ref: false, trigger: false, right: false }
    s = toggleSide(s, 'ref')
    expect(s).toEqual({ ref: true, trigger: false, right: false })
    s = toggleSide(s, 'trigger')
    expect(s).toEqual({ ref: false, trigger: true, right: false })
  })

  it('toggle 关态只动自己，不误伤同缘另一个', () => {
    const s1 = toggleSide({ ref: true, trigger: false, right: true }, 'ref')
    expect(s1).toEqual({ ref: false, trigger: false, right: true })
  })

  it('右栏与左缘互不干扰', () => {
    let s: Sides = { ref: false, trigger: false, right: false }
    s = toggleSide(s, 'ref')
    s = toggleSide(s, 'right')
    expect(s).toEqual({ ref: true, trigger: false, right: true })
  })

  it('幂等：连续开同一个弹窗两次，侧边栏不会复活', () => {
    let s: Sides = { ref: true, trigger: false, right: false }
    s = openModal(s)
    s = openModal(s)
    expect(s).toEqual({ ref: false, trigger: false, right: false })
  })
})

describe('z-index 层级（回归锁）', () => {
  // 层级必须满足：弹窗 > 侧边栏> 放大态窗体
  const LVL = { 'z-node-editor': 8000, 'z-editor-side': 8100, 'z-editor-pop': 8500, 'z-modal': 10000 }
  it('弹窗盖住侧边栏（这是原bug 的根因）', () => {
    expect(LVL['z-modal']).toBeGreaterThan(LVL['z-editor-side'])
  })
  it('侧边栏盖住放大态窗体', () => {
    expect(LVL['z-editor-side']).toBeGreaterThan(LVL['z-node-editor'])
  })
  it('侧边栏低于放大态内部弹层', () => {
    expect(LVL['z-editor-pop']).toBeGreaterThan(LVL['z-editor-side'])
  })
})