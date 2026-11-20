import { describe, it, expect } from 'vitest'
import { replaceInParagraph, scanParagraph, paragraphBlocks, type PlaceholderHit } from './text'

/**
 * 这些用例全部围绕一个真实风险：PowerPoint 会把占位符拆散到多个 a:r。
 * 拆散后单个 a:t 里的正则匹配全部落空 —— 不报错、静默失效。
 */

const run = (t: string) => `<a:r><a:rPr lang="zh-CN"/><a:t>${t}</a:t></a:r>`
const para = (...runs: string[]) => `<a:p>${runs.join('')}</a:p>`

/**
 * 把段落里所有 a:t 的文字拼出来，作为"用户在 PowerPoint 里看到的内容"。
 * 必须解码实体 —— 拼的是 XML 原文，`&lt;` 在屏幕上显示的是 `<`。
 * （曾经这里没解码，导致「转义正确」被误判成「转义失败」。）
 */
function visible(block: string): string {
  const raw = [...block.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]).join('')
  return raw
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

const resolve = (v: Record<string, string>) => (_tag: string, hit: PlaceholderHit) =>
  hit.path in v ? v[hit.path] : null

describe('占位符扫描', () => {
  it('识别整段未拆散的占位符', () => {
    const hits = scanParagraph(para(run('项目：{{field:项目名}}')))
    expect(hits).toHaveLength(1)
    expect(hits[0]).toMatchObject({ kind: 'field', path: '项目名' })
  })

  it('识别被拆散到 3 个 run 的占位符（核心 case）', () => {
    const block = para(run('{{field:'), run('项目名'), run('}}'))
    const hits = scanParagraph(block)
    expect(hits).toHaveLength(1)
    expect(hits[0].path).toBe('项目名')
  })

  it('同一段落里的多个占位符都能识别', () => {
    const block = para(run('{{field:A}} 至 {{field:B}}'))
    expect(scanParagraph(block).map((h) => h.path)).toEqual(['A', 'B'])
  })

  it('占位符前后有普通文字时不受影响', () => {
    const hits = scanParagraph(para(run('合计金额：{{field:总价}} 元')))
    expect(hits.map((h) => h.path)).toEqual(['总价'])
  })

  it('非占位符的花括号不会被误认', () => {
    expect(scanParagraph(para(run('{ 不是占位符 }')))).toHaveLength(0)
    expect(scanParagraph(para(run('{{ 没有类型 }}')))).toHaveLength(0)
  })
})

describe('替换：未拆散', () => {
  it('就地替换，保留段落结构', () => {
    const block = para(run('项目：{{field:项目名}}'))
    const { block: out, count } = replaceInParagraph(block, resolve({ 项目名: '凤凰社区' }))
    expect(count).toBe(1)
    expect(visible(out)).toBe('项目：凤凰社区')
    expect(out.startsWith('<a:p>')).toBe(true)
  })

  it('多次替换累计计数', () => {
    const block = para(run('{{field:A}}-{{field:B}}'))
    const { count } = replaceInParagraph(block, resolve({ A: '1', B: '2' }))
    expect(count).toBe(2)
  })

  it('值里的 XML 特殊字符被转义', () => {
    const block = para(run('{{field:备注}}'))
    const { block: out } = replaceInParagraph(block, resolve({ 备注: 'a<b & "c"' }))
    expect(visible(out)).toBe('a<b & "c"')
    // 原文里必须是实体，否则 XML 破损
    expect(out).toContain('a&lt;b &amp; &quot;c&quot;')
  })

  it('resolver 返回 null 时原样保留', () => {
    const block = para(run('{{field:未配}}'))
    const { block: out, count } = replaceInParagraph(block, () => null)
    expect(count).toBe(0)
    expect(visible(out)).toBe('{{field:未配}}')
  })
})

describe('替换：跨 run（PowerPoint 真实行为）', () => {
  it('拆成 3 段也能替换，值落在第一个 run', () => {
    const block = para(run('{{field:'), run('项目名'), run('}}'))
    const { block: out, count } = replaceInParagraph(block, resolve({ 项目名: '凤凰社区' }))
    expect(count).toBe(1)
    expect(visible(out)).toBe('凤凰社区')
  })

  it('占位符前后还有别的文字时，前后文字都保住', () => {
    const block = para(run('合计：{{field:'), run('总价'), run('}} 元'))
    const { block: out } = replaceInParagraph(block, resolve({ 总价: '128.5' }))
    expect(visible(out)).toBe('合计：128.5 元')
  })

  it('中间 run 还有占位符之外的文字时不会被误删', () => {
    const block = para(run('{{field:'), run('A}} 与 {{field:'), run('B}}'))
    const { block: out, count } = replaceInParagraph(block, resolve({ A: '一', B: '二' }))
    expect(count).toBe(2)
    expect(visible(out)).toBe('一 与 二')
  })

  it('替换后 a:r 数量不减少（结构保持，避免 PowerPoint 报错）', () => {
    const block = para(run('{{field:'), run('X'), run('}}'))
    const before = (block.match(/<a:r>/g) ?? []).length
    const { block: out } = replaceInParagraph(block, resolve({ X: '值' }))
    expect((out.match(/<a:r>/g) ?? []).length).toBe(before)
  })

  it('跨 run 时值里的特殊字符同样被转义', () => {
    const block = para(run('{{field:'), run('备注'), run('}}'))
    const { block: out } = replaceInParagraph(block, resolve({ 备注: '<b>&' }))
    expect(out).toContain('&lt;b&gt;&amp;')
    expect(out).not.toContain('<b>&')
  })
})

describe('段落切分', () => {
  it('切出 txBody 里的多个 a:p', () => {
    const xml = `<p:txBody><a:bodyPr/><a:p>${run('一')}</a:p><a:p>${run('二')}</a:p></p:txBody>`
    const blocks = paragraphBlocks(xml)
    expect(blocks).toHaveLength(2)
    expect(blocks[1].at).toBeGreaterThan(blocks[0].at)
  })

  it('切出的区间能原样拼回原文', () => {
    const xml = `<p:txBody><a:p>${run('一')}</a:p><a:p>${run('二')}</a:p></p:txBody>`
    const blocks = paragraphBlocks(xml)
    for (const b of blocks) expect(xml.slice(b.at, b.at + b.block.length)).toBe(b.block)
  })
})
