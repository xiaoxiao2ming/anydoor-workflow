import { describe, expect, it } from 'vitest'
import { matchIntent } from './chatIntent'

/* ★ 2026-11-07 用户连说两次「保存画布文件」「保存整个画布文件」，
 *  对话却回了「已合并 1 份 / 3 页 → 合并-2026-11-07-1325.pdf」。
 *  保存是画布动作，必须命中 save-canvas，而不是被 run/add-node 抢走。 */
describe('matchIntent — 保存画布话术', () => {
  it('「保存画布文件」→ save-canvas', () => {
    expect(matchIntent('保存画布文件', [])?.id).toBe('save-canvas')
  })

  it('「保存整个画布文件」→ save-canvas', () => {
    expect(matchIntent('保存整个画布文件', [])?.id).toBe('save-canvas')
  })

  it('「把画布存一下」→ save-canvas', () => {
    expect(matchIntent('把画布存一下', [])?.id).toBe('save-canvas')
  })

  it('「存盘」→ save-canvas', () => {
    expect(matchIntent('存盘', [])?.id).toBe('save-canvas')
  })

  it('「保存一下」「保存吧」→ save-canvas', () => {
    expect(matchIntent('保存一下', [])?.id).toBe('save-canvas')
    expect(matchIntent('保存吧', [])?.id).toBe('save-canvas')
  })

  it('「保存工程」→ save-canvas', () => {
    expect(matchIntent('保存工程', [])?.id).toBe('save-canvas')
  })

  it('保存动作不会被执行意图抢走（action 就是 saveCanvas）', () => {
    expect(matchIntent('保存画布文件', [])?.action).toBe('saveCanvas')
  })
})

describe('matchIntent — 别把产出文件当成存工程', () => {
  it('「把文档转成PDF」仍是 add-node，不是 save-canvas', () => {
    const hit = matchIntent('把文档转成PDF', [])
    expect(hit?.id).not.toBe('save-canvas')
    expect(hit?.id).toBe('add-node')
  })

  it('「保存成PDF」不该走存工程', () => {
    expect(matchIntent('保存成PDF', [])?.id).not.toBe('save-canvas')
  })

  it('「执行 N071」仍是 run-node', () => {
    expect(matchIntent('执行 N071', [])?.id).toBe('run-node')
  })
})
