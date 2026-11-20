import { describe, expect, it } from 'vitest'
import { matchIntent, parseNodeRefs, resolveNodeRef } from './chatIntent'

/* ★ 2026-11-07 用户「N071节点」→ 模型答「N071节点是什么意思？你能再详细说明一下吗？」：
 *  节点标题写着 [N071]，对话却完全不认这个编号。规则层必须先把编号认出来。 */
describe('parseNodeRefs — 节点编号识别', () => {
  it('N071 / n71 / N0071 都解析成同一个编号 71', () => {
    expect(parseNodeRefs('N071节点')).toEqual([71])
    expect(parseNodeRefs('n71')).toEqual([71])
    expect(parseNodeRefs('N0071 是什么')).toEqual([71])
  })

  it('一句话里多个编号：去重且保持出现顺序', () => {
    expect(parseNodeRefs('把 N012 连到 N005，再接 N012')).toEqual([12, 5])
  })

  it('没有编号、或只是字母 n 时返回空', () => {
    expect(parseNodeRefs('分析下文档节点')).toEqual([])
    expect(parseNodeRefs('n卡驱动')).toEqual([])
    expect(parseNodeRefs('')).toEqual([])
  })
})

describe('matchIntent — 编号话术的路由', () => {
  it('「N071节点」→ node-info（不是闲聊兜底）', () => {
    expect(matchIntent('N071节点', [])?.id).toBe('node-info')
  })

  it('「N071是什么意思」→ node-info', () => {
    expect(matchIntent('N071是什么意思', [])?.id).toBe('node-info')
  })

  it('「执行N071」→ run-node（带编号执行）', () => {
    expect(matchIntent('执行N071', [])?.id).toBe('run-node')
  })

  it('不带编号的普通提问不被编号规则抢走', () => {
    expect(matchIntent('分析下文档节点中的内容并总结', [])?.id).toBe('analyze-node')
  })
})

describe('resolveNodeRef — 画布为空时不误判', () => {
  it('画布上没有对应编号 → null（交给调用方提示现有编号）', () => {
    expect(resolveNodeRef('N999节点')).toBeNull()
  })
})
