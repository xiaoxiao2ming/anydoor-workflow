import { describe, expect, it } from 'vitest'
import { matchIntent } from './chatIntent'

/* ★ 2026-11-08 用户截图实测：「word转PDF」被 LLM 路由瞎猜成 image-to-word，
 *  建出「图片导入 → Word 输出」节点——完全驴唇不对马嘴。
 *  真因：规则层全不命中（add-node 只认「转成PDF」，裸「转PDF」不认），
 *  意图清单里又没有 Word→PDF，模型只能抓最像「转换」的意图猜。
 *  新增 word-to-pdf 意图（action: toPdf）接右键「转为PDF节点」既有链路。 */
describe('matchIntent — Word 转 PDF 话术', () => {
  it('「word转PDF」（用户原话）→ word-to-pdf', () => {
    const hit = matchIntent('word转PDF', [])
    expect(hit?.id).toBe('word-to-pdf')
    expect(hit?.action).toBe('toPdf')
  })

  it('「word转pdf」小写 → word-to-pdf', () => {
    expect(matchIntent('word转pdf', [])?.id).toBe('word-to-pdf')
  })

  it('「把word转成PDF」→ word-to-pdf', () => {
    expect(matchIntent('把word转成PDF', [])?.id).toBe('word-to-pdf')
  })

  it('「word文档转pdf」不被 image-to-word 抢走', () => {
    expect(matchIntent('word文档转pdf', [])?.id).toBe('word-to-pdf')
  })

  it('「docx转PDF」→ word-to-pdf', () => {
    expect(matchIntent('docx转PDF', [])?.id).toBe('word-to-pdf')
  })

  it('「word to pdf」英文语序 → word-to-pdf', () => {
    expect(matchIntent('word to pdf', [])?.id).toBe('word-to-pdf')
  })

  it('「把这个word转换为pdf」→ word-to-pdf', () => {
    expect(matchIntent('把这个word转换为pdf', [])?.id).toBe('word-to-pdf')
  })

  it('「N105 转PDF」（编号引用）→ word-to-pdf，不依赖 LLM 路由（2026-11-09）', () => {
    expect(matchIntent('N105 转PDF', [])?.id).toBe('word-to-pdf')
    expect(matchIntent('把N105转成pdf', [])?.id).toBe('word-to-pdf')
    expect(matchIntent('N105 转成PDF', [])?.action).toBe('toPdf')
  })

  it('「N105」单独问仍是 node-info，不被编号转PDF规则误伤', () => {
    expect(matchIntent('N105', [])?.id).toBe('node-info')
    expect(matchIntent('N105节点是什么', [])?.id).toBe('node-info')
  })
})

describe('matchIntent — 反向与邻近话术不误伤', () => {
  it('「pdf转word」不是 word-to-pdf（方向反了）', () => {
    expect(matchIntent('pdf转word', [])?.id).not.toBe('word-to-pdf')
  })

  it('「把文档转成PDF」（没提 word/docx）仍走 add-node，不是 word-to-pdf', () => {
    expect(matchIntent('把文档转成PDF', [])?.id).toBe('add-node')
  })

  it('「把图片转成word」仍是 image-to-word', () => {
    expect(matchIntent('把图片转成word', [])?.id).toBe('image-to-word')
  })

  it('「识别图片文字」仍是 image-to-word', () => {
    expect(matchIntent('识别图片文字', [])?.id).toBe('image-to-word')
  })
})
