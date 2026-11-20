import { describe, expect, it } from 'vitest'
import { routeLlmOpt } from './chatIntent'

/* ★ 2026-11-09 用户「能否做成优化混合模型」：任务感知的模型路由（Anthropic Routing 模式）。
 *  未锁模型时：route/query（短输出格式任务）→ light(3B)；summarize/analyze（质量）→ heavy(9B)；
 *  chat 跟随思考档位（off→light 秒答）；用户锁定具体模型 → 全尊重不传 tier。
 *  注意：localStorage 在 vitest 环境存在（jsdom/happy-dom 或全局 stub），默认未锁定 = 无 model。 */
describe('routeLlmOpt — 混合模型任务路由', () => {
  it('route/query 任务 → light（短输出快任务）', () => {
    expect(routeLlmOpt('route').tier).toBe('light')
    expect(routeLlmOpt('query').tier).toBe('light')
  })

  it('summarize/analyze 任务 → heavy（质量优先）', () => {
    expect(routeLlmOpt('summarize').tier).toBe('heavy')
    expect(routeLlmOpt('analyze').tier).toBe('heavy')
  })

  it('query 任务的 noThink/maxTokens 覆盖仍由调用方合成（不吞 base 外参数）', () => {
    const opt = { ...routeLlmOpt('query'), noThink: true, maxTokens: 256 }
    expect(opt.noThink).toBe(true)
    expect(opt.maxTokens).toBe(256)
    expect(opt.tier).toBe('light')
  })

  it('用户锁定模型时不传 tier（全尊重手选）', () => {
    /* vitest node 环境无 localStorage——stub 一个（currentLlmModelOpt 只用 getItem） */
    const stub = { getItem: () => 'Qwen3.5-9B-Q4_K_M.gguf', setItem: () => {}, removeItem: () => {} }
    ;(globalThis as unknown as { localStorage: unknown }).localStorage = stub
    try {
      const opt = routeLlmOpt('route')
      expect(opt.model).toBe('Qwen3.5-9B-Q4_K_M.gguf')
      expect(opt.tier).toBeUndefined()
    } finally {
      delete (globalThis as unknown as { localStorage?: unknown }).localStorage
    }
  })
})
