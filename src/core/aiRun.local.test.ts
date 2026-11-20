import { afterEach, describe, expect, it } from 'vitest'
import { runAiNode } from './aiRun'
import type { TableSchema } from './schema'

/* ★ 2026-11-09 混合模型收尾：AI 节点本地模式（API 留空 → 走 llm:chat 本地大模型）。
 *  mock globalThis.api.llmChat 验证：零配置可分析、输出写 aiLastOutput、失败写 aiLastError。
 *  afterEach 必须卸载 stub——用例抛异常若不清理会污染后续用例（实测踩过）。 */

type Store = {
  tables: Record<string, TableSchema>
  records: Record<string, Record<string, unknown>[]>
  setNodeConfig: (id: string, patch: Record<string, unknown>) => void
}

function makeStore(): Store {
  const tables: Record<string, TableSchema> = {
    ai1: { id: 'ai1', name: 'AI 分析', kind: 'ai', fields: [], config: {} },
  }
  const patches: Record<string, Record<string, unknown>> = {}
  const store: Store = {
    tables,
    records: {},
    setNodeConfig: (id, patch) => {
      patches[id] = { ...(patches[id] ?? {}), ...patch }
      tables[id] = { ...tables[id], config: { ...(tables[id].config ?? {}), ...patch } }
    },
  }
  ;(store as unknown as { __patches: typeof patches }).__patches = patches
  return store
}

afterEach(() => {
  delete (globalThis as unknown as { api?: unknown }).api
})

describe('runAiNode — 本地模式（API 留空走本地大模型）', () => {
  it('零配置：aiApiBase 空 → llmChat 被调用，输出写 aiLastOutput', async () => {
    let gotPrompt = ''
    ;(globalThis as unknown as { api?: unknown }).api = {
      llmChat: (p: string) => { gotPrompt = p; return { ok: true, text: '监测到数值变化，整体平稳。' } },
    }
    const store = makeStore()
    const r = await runAiNode('ai1', store as unknown as Parameters<typeof runAiNode>[1])
    expect(r.ok).toBe(true)
    expect(typeof gotPrompt).toBe('string')
    const patches = (store as unknown as { __patches: Record<string, Record<string, unknown>> }).__patches
    expect(String(patches.ai1.aiLastOutput)).toContain('平稳')
    expect(patches.ai1.aiLastError).toBeUndefined()
  })

  it('本地推理失败 → 错误写 aiLastError', async () => {
    ;(globalThis as unknown as { api?: unknown }).api = {
      llmChat: () => ({ ok: false, reason: '模型加载失败' }),
    }
    const store = makeStore()
    const r = await runAiNode('ai1', store as unknown as Parameters<typeof runAiNode>[1])
    expect(r.ok).toBe(false)
    const patches = (store as unknown as { __patches: Record<string, Record<string, unknown>> }).__patches
    expect(String(patches.ai1.aiLastError)).toContain('模型加载失败')
  })

  it('缺少推理模块 → 给可读指引', async () => {
    const store = makeStore()
    const r = await runAiNode('ai1', store as unknown as Parameters<typeof runAiNode>[1])
    const patches = (store as unknown as { __patches: Record<string, Record<string, unknown>> }).__patches
    expect(r.ok).toBe(false)
    expect(String(patches.ai1.aiLastError)).toContain('本地模型运行时不可用')
  })
})
