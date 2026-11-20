/**
 * docx 内容控件发现测试
 *
 * 立档：2026-10-29 · 配套 src/core/docxSdtDiscover.ts
 *
 * ★ 用**真实 docx 样本**做主用例 —— 实测发现真实 tag 是
 *   `field:工程名称` / `each:材料表` 这种 **kind:名称** 形态，
 *   不是裸名称（这决定了 toWordSites 要剥前缀）。
 */

import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { discoverDocxSdt, scanSdtInXml, toWordSites } from './docxSdtDiscover'
import { patchDocxSdt } from './ooWriteback'

const SAMPLE = 'sample-template/字段模板.docx'   // 真实文件，含 3 个 sdt

describe('★★ ① 真实 docx：能列出内容控件', () => {
  it('发现全部 sdt（真实样本）', async () => {
    const bytes = new Uint8Array(readFileSync(join(process.cwd(), SAMPLE)))
    const list = await discoverDocxSdt(bytes)
    expect(list.length).toBeGreaterThanOrEqual(3)
    // 实测该文件是 field:工程名称 / field:设计流量 / each:材料表
    const tags = list.map((s) => s.tag)
    expect(tags).toContain('field:工程名称')
    expect(tags.some((t) => t.startsWith('each:'))).toBe(true)
  })

  it('★ 真实 tag 是 kind:名称 形态 ⇒ kind 被正确切出', async () => {
    const bytes = new Uint8Array(readFileSync(join(process.cwd(), SAMPLE)))
    const list = await discoverDocxSdt(bytes)
    const f = list.find((s) => s.tag === 'field:工程名称')
    expect(f).toBeDefined()
    expect(f!.kind).toBe('field')
  })

  it('能读到控件当前文本（供 UI 展示"这域现在是几"）', async () => {
    const bytes = new Uint8Array(readFileSync(join(process.cwd(), SAMPLE)))
    const list = await discoverDocxSdt(bytes)
    // 至少有一个控件能读出文本或空串（不抛）
    for (const s of list) expect(typeof s.text).toBe('string')
  })
})

describe('★★ ② 发现与写回口径一致（最关键）', () => {
  it('★ 发现出来的 tag，patchDocxSdt 一定写得进去', async () => {
    const bytes = new Uint8Array(readFileSync(join(process.cwd(), SAMPLE)))
    const list = await discoverDocxSdt(bytes)
    // 逐个尝试写：hit 必须全为 true
    // ★ 这是本文件最要紧的一条 —— 「发现得到但写不进去」是最隐蔽的割裂
    const jszip = await import('jszip')
    const zip = await jszip.default.loadAsync(bytes)
    let xml = await zip.file('word/document.xml')!.async('string')
    for (const s of list) {
      const r = patchDocxSdt(xml, s.tag, 'ZZ测试ZZ')
      expect(r.hit, `tag=${s.tag} 发现得到却写不进去`).toBe(true)
    }
  })

  it('没有 tag 的控件不产出（不是我们的域）', () => {
    const xml = '<w:document><w:body><w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>'
    expect(scanSdtInXml(xml)).toHaveLength(0)
  })

  it('多个 sdt 时不跨块误匹配（tag 不会张冠李戴）', () => {
    const xml =
      '<w:document><w:body>' +
      '<w:sdt><w:sdtPr><w:tag w:val="A"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>旧A</w:t></w:r></w:p></w:sdtContent></w:sdt>' +
      '<w:sdt><w:sdtPr><w:tag w:val="B"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>旧B</w:t></w:r></w:p></w:sdtContent></w:sdt>' +
      '</w:body></w:document>'
    const list = scanSdtInXml(xml)
    expect(list.map((s) => s.tag)).toEqual(['A', 'B'])
    // 只改B，A 必须保持原样
    const r = patchDocxSdt(xml, 'B', '新B')
    expect(r.xml).toContain('旧A')
    expect(r.xml).toContain('新B')
  })
})

describe('③ toWordSites：把 kind:名称 剥成可匹配的 key', () => {
  it('field:工程名称 ⇒ key= 工程名称（宿主侧域名/字段名不带前缀）', () => {
    const w = toWordSites([{ tag: 'field:工程名称', part: 'word/document.xml', text: '' }])
    expect(w[0].key).toBe('工程名称')
    expect(w[0].tag).toBe('field:工程名称')   // tag 保持原样（写回时用它定位）
  })
  it('无前缀时 key === tag', () => {
    const w = toWordSites([{ tag: '表甲.数量', part: '', text: '' }])
    expect(w[0].key).toBe('表甲.数量')
  })
  it('★ 传进来的应是**已解码**明文（解码在 scanSdtInXml 那层做）', () => {
    // DocxSdtSite.tag 的契约就是**明文**（scanSdtInXml 已反转义过），
    // 所以 toWordSites 不该、也不需要再解一次。
    // （第一版这里喂了 'field:A&amp;B' 原始实体串，期望被解码 ⇒ 报红，
    //   是我把「哪一层负责解码」搞混了，测试写错不是实现错。）
    const w = toWordSites([{ tag: 'field:A&B', part: '', text: '' }])
    expect(w[0].tag).toBe('field:A&B')
    expect(w[0].key).toBe('A&B')
  })
  it('解码发生在 scanSdtInXml：原始实体 → 明文 tag', () => {
    const xml = '<w:document><w:body><w:sdt><w:sdtPr><w:tag w:val="field:A&amp;B"/></w:sdtPr>' +
      '<w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>'
    const list = scanSdtInXml(xml)
    expect(list[0].tag).toBe('field:A&B')   // ← 解码在这一层
  })
})

describe('④ 边界', () => {
  it('不是 zip ⇒ 空列表', async () => {
    expect(await discoverDocxSdt(new Uint8Array([1, 2, 3]))).toEqual([])
  })
  it('没有 document.xml ⇒ 空列表', async () => {
    const JSZip = (await import('jszip')).default
    const zip = new JSZip()
    zip.file('hello.txt', 'x')
    expect(await discoverDocxSdt(await zip.generateAsync({ type: 'uint8array' }))).toEqual([])
  })
})
