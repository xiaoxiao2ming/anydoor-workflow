import { describe, expect, it } from 'vitest'
import { aiBuildMessages, aiInterpolate, aiUpstreamContentOf, aiVarMapOf } from './aiRun'
import { aiPrefsPatchOf, defineNode } from './schema'
import type { NodeConfig, TableSchema } from './schema'
import type { Row } from './store'

const T = (id: string, kind: TableSchema['kind'], name = id): TableSchema =>
  ({ id, kind, name, fields: [], config: {} }) as TableSchema

describe('aiUpstreamContentOf 上游内容组装', () => {
  it('text 节点 = 原文直出（单行）', () => {
    const tables = { a: T('a', 'text') }
    const records = { a: [{ text: '你好世界' }] as Row[] }
    expect(aiUpstreamContentOf('a', tables, records)).toBe('你好世界')
  })
  it('text 节点多行 = 换行合并', () => {
    const tables = { a: T('a', 'text') }
    const records = { a: [{ text: '第一行' }, { text: '第二行' }] as Row[] }
    expect(aiUpstreamContentOf('a', tables, records)).toBe('第一行\n第二行')
  })
  it('非 text 节点 = rows 的 JSON 兜底', () => {
    const tables = { b: { ...T('b', 'table'), fields: [{ id: 'name', name: '名称', type: 'text' }] } }
    const records = { b: [{ name: '钢管' }] as Row[] }
    expect(aiUpstreamContentOf('b', tables, records)).toContain('钢管')
  })
  it('未知节点 = 空串', () => {
    expect(aiUpstreamContentOf('x', {}, {})).toBe('')
  })
})

describe('aiInterpolate 插值', () => {
  it('命中替换', () => {
    expect(aiInterpolate('请处理：{{内容}}', { 内容: 'ABC' })).toBe('请处理：ABC')
  })
  it('变量名两侧空白等价', () => {
    expect(aiInterpolate('{{ 内容 }}', { 内容: 'X' })).toBe('X')
  })
  it('未命中原样保留（不吞字面花括号）', () => {
    expect(aiInterpolate('a {{不存在}} b', {})).toBe('a {{不存在}} b')
  })
})

describe('aiVarMapOf 变量表', () => {
  it('配置名优先，未配置用上游节点名兜底', () => {
    const tables = { a: T('a', 'text', '原始资料') }
    const records = { a: [{ text: '正文' }] as Row[] }
    const named = { linkedInputs: ['a'], aiInputVars: [{ name: '资料', from: 'a' }] } as NodeConfig
    expect(aiVarMapOf(named, tables, records)).toEqual({ 资料: '正文' })
    const plain = { linkedInputs: ['a'] } as NodeConfig
    expect(aiVarMapOf(plain, tables, records)).toEqual({ 原始资料: '正文' })
  })
})

describe('aiBuildMessages 消息组装', () => {
  const tables = { a: T('a', 'text', '文本 1') }
  const records = { a: [{ text: '你好' }] as Row[] }
  it('空 userPrompt → 上游内容直接当 user 消息（零配置闭环）', () => {
    const cfg = { linkedInputs: ['a'], aiSystemPrompt: '你是文本处理助手。' } as NodeConfig
    expect(aiBuildMessages(cfg, tables, records)).toEqual([
      { role: 'system', content: '你是文本处理助手。' },
      { role: 'user', content: '你好' },
    ])
  })
  it('有 userPrompt → 插值 {{变量}}', () => {
    const cfg = { linkedInputs: ['a'], aiUserPrompt: '请翻译：{{文本 1}}' } as NodeConfig
    const msgs = aiBuildMessages(cfg, tables, records)
    expect(msgs[msgs.length - 1]).toEqual({ role: 'user', content: '请翻译：你好' })
  })
  it('outputNote 拼进 system；无 system 时只有 user', () => {
    const cfg = { linkedInputs: ['a'], aiOutputNote: '只输出 JSON' } as NodeConfig
    const msgs = aiBuildMessages(cfg, tables, records)
    expect(msgs[0]).toEqual({ role: 'system', content: '输出要求：只输出 JSON' })
    const bare = { linkedInputs: ['a'] } as NodeConfig
    expect(aiBuildMessages(bare, tables, records)).toEqual([{ role: 'user', content: '你好' }])
  })
})

describe('aiPrefsPatchOf 偏好匹配', () => {
  it('未配置 + text 上游 → 预填 text-io（文字进文字出）', () => {
    const patch = aiPrefsPatchOf({}, ['text'])
    expect(patch).not.toBeNull()
    expect(patch?.aiSkillId).toBe('text-io')
    expect(patch?.aiOutputMode).toBe('text')
    expect(String(patch?.aiSystemPrompt)).toContain('文本处理助手')
  })
  it('已配置（含 custom）一律不覆盖', () => {
    expect(aiPrefsPatchOf({ aiSkillId: 'custom' }, ['text'])).toBeNull()
    expect(aiPrefsPatchOf({ aiSkillId: 'summarize' }, ['text'])).toBeNull()
  })
  it('无偏好注册的类型 → null', () => {
    expect(aiPrefsPatchOf({}, ['table'])).toBeNull()
    expect(aiPrefsPatchOf({}, [])).toBeNull()
  })
  it('多上游取第一个命中', () => {
    expect(aiPrefsPatchOf({}, ['table', 'text'])?.aiSkillId).toBe('text-io')
  })
})

describe('text 节点预设（样例链路的结构依据）', () => {
  it('text 预设字段 id 为 text（records 行 key 与 aiUpstreamContentOf 对齐）', () => {
    const node = defineNode('text', '文本', [{ id: 'text', name: 'text', type: 'text' }], {}, 't1')
    expect(node.fields[0].id).toBe('text')
  })
})
