import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { sdtifyDocx } from './sdtify'
import { parseTemplate, bind } from './binder'
import { buildData } from './buildData'
import type { FieldDef } from '../core/schema'
import type { Row } from '../core/store'

/**
 * 端到端：模拟「设计院给一份普通 docx → 导入 → 绑定上游数据 → 出稿」。
 * 覆盖的是 UI 接入层（sdtifyDocx / buildData），引擎内部由 binder.test.ts 保证。
 */

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const run = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`
const para = (...r: string[]) => `<w:p>${r.join('')}</w:p>`

/** 设计院视角的普通模板：只有占位符标记，没有任何内容控件 */
const BODY = [
  para(run('{{field:项目名称}}')),
  para(run('本工程管线总长 '), run('{{field:长度}}'), run(' 米。')),
  '<w:tbl><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>' +
    `<w:p>${run('{{each:清单}}{{名称}}')}</w:p></w:tc>` +
    '<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>' +
    `<w:p>${run('{{数量}}')}</w:p></w:tc></w:tr></w:tbl>`,
  para(run('{{rich:结论}}')),
  para(run('{{if:有风险}}风险提示：需支护。')),
].join('')

const DOC = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${BODY}</w:body></w:document>`

async function makeDocx(xml: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  )
  zip.file('word/document.xml', xml)
  return zip.generateAsync({ type: 'uint8array' })
}

const docOf = async (b: Uint8Array) => (await JSZip.loadAsync(b)).file('word/document.xml')!.async('string')

/** 模拟 computeNode 的输出 */
const FIELDS: FieldDef[] = [
  { id: 'f_name', name: '项目名称', type: 'text' },
  { id: 'f_len', name: '长度', type: 'number' },
]
const ROWS: Row[] = [
  { f_name: '示例路改造工程', f_len: 128.4 },
  { f_name: '示例路改造工程', f_len: 216.7 },
]
const DETAIL: Row[] = [
  { 名称: 'DN300 管', 数量: 12 },
  { 名称: 'DN400 管', 数量: 8 },
  { 名称: 'DN500 管', 数量: 5 },
]

describe('端到端：普通 docx → 导入 → 绑定 → 出稿', () => {
  it('sdtifyDocx 把标记原地转成控件，且只改 document.xml', async () => {
    const raw = await makeDocx(DOC)
    const r = await sdtifyDocx(raw)

    expect(r.count).toBe(5)
    expect(r.tags).toContain('field:项目名称')
    expect(r.tags).toContain('each:清单')
    expect(r.tags).toContain('rich:结论')
    expect(r.tags).toContain('if:有风险')
    expect(r.changed).toEqual(['word/document.xml'])

    // [Content_Types].xml 没被动过（非目标部件字节不变）
    const rawCt = await (await JSZip.loadAsync(raw)).file('[Content_Types].xml')!.async('string')
    const outCt = await (await JSZip.loadAsync(r.buffer)).file('[Content_Types].xml')!.async('string')
    expect(outCt).toBe(rawCt)

    // each / if 只是声明标记，不该留在正文里
    const xml = await docOf(r.buffer)
    expect(xml).not.toContain('{{each:清单}}')
    expect(xml).not.toContain('{{if:有风险}}')
    // 行内短占位符要留着，供 each 展开时填充
    expect(xml).toContain('{{名称}}')
    expect(xml).toContain('{{数量}}')
  })

  it('buildData 自动按字段名匹配，未配置的标为 missing', async () => {
    const templated = await sdtifyDocx(await makeDocx(DOC))
    const { bindings } = await parseTemplate(templated.buffer)
    expect(bindings.length).toBe(5)

    const built = buildData(bindings, {}, FIELDS, ROWS)
    // 字段名「项目名称」「长度」与上游字段同名 → 自动匹配
    expect(built.data['项目名称']).toBe('示例路改造工程')
    // 没配 agg 时取首行
    expect(built.data['长度']).toBe(128.4)
    expect(built.missing).toContain('rich:结论')
  })

  it('each 取上游明细行，each 行数随数据变化', async () => {
    const templated = await sdtifyDocx(await makeDocx(DOC))
    const { bindings } = await parseTemplate(templated.buffer)

    // 手工映射：清单 → 上游明细行
    const built = buildData(
      bindings,
      {
        'each:清单': { from: 'rows' },
        'field:项目名称': { from: 'manual', value: '示例路改造工程' },
        'rich:结论': { from: 'manual', value: '<p>满足规范。</p>' },
        'if:有风险': { from: 'manual', value: 'true' },
      },
      FIELDS,
      DETAIL,
    )
    expect((built.data['清单'] as Row[]).length).toBe(3)

    const { buffer } = await bind(templated.buffer, built.data, { stripControls: true })
    const xml = await docOf(buffer)
    expect(xml).toContain('DN300 管')
    expect(xml).toContain('DN400 管')
    expect(xml).toContain('DN500 管')
    // 全部控件都被赋值 → 外壳剥净，交付稿里不留任何 <w:sdt>
    expect(xml).not.toContain('<w:sdt>')
    expect(xml).not.toContain('{{')
  })

  it('聚合 sum 与条件段真值生效', async () => {
    const templated = await sdtifyDocx(await makeDocx(DOC))
    const { bindings } = await parseTemplate(templated.buffer)

    const built = buildData(
      bindings,
      {
        'field:长度': { from: 'field', fieldId: 'f_len', agg: 'sum' },
        'each:清单': { from: 'rows' },
        'rich:结论': {
          from: 'manual',
          value: '<p>满足<b>规范</b></p><ol><li>全线重力流</li></ol>',
        },
        'if:有风险': { from: 'manual', value: 'true' },
      },
      FIELDS,
      DETAIL,
    )

    // 128.4 + 216.7 = 345.1（用 ROWS 的数据，这里传的是 DETAIL，长度字段缺失 → sum 为空串）
    expect(built.data['长度']).toBe('')

    const built2 = buildData(
      bindings,
      { 'field:长度': { from: 'field', fieldId: 'f_len', agg: 'sum' } },
      FIELDS,
      ROWS,
    )
    expect(built2.data['长度']).toBe(345.1)

    const { buffer } = await bind(templated.buffer, built.data, { stripControls: true })
    const xml = await docOf(buffer)
    expect(xml).toContain('风险提示：需支护。')
    expect(xml).toContain('满足')
    expect(xml).toContain('w:numId w:val="101"')
  })

  it('if 为假时整段删除', async () => {
    const templated = await sdtifyDocx(await makeDocx(DOC))
    const { bindings } = await parseTemplate(templated.buffer)
    const built = buildData(bindings, { 'if:有风险': { from: 'manual', value: 'false' } }, FIELDS, ROWS)
    const { buffer } = await bind(templated.buffer, built.data, { stripControls: true })
    const xml = await docOf(buffer)
    expect(xml).not.toContain('风险提示')
  })

  it('each 列映射：模板列名与上游字段名不一致时也能填充', async () => {
    const templated = await sdtifyDocx(await makeDocx(DOC))
    const { bindings } = await parseTemplate(templated.buffer)
    // 模板行内占位符是 {{名称}}/{{数量}}，上游明细字段故意用不同的名字
    const oddRows: Row[] = [
      { seg: 'K1-K2', spec: 'DN300' },
      { seg: 'K2-K3', spec: 'DN400' },
    ]
    // 先确认 placeholders 提取到了
    const each = bindings.find((b) => b.kind === 'each')!
    expect(each.placeholders).toContain('名称')
    expect(each.placeholders).toContain('数量')

    const built = buildData(
      bindings,
      {
        'each:清单': { from: 'rows', columns: { 名称: 'seg', 数量: 'spec' } },
        'field:项目名称': { from: 'manual', value: '映射测试' },
        'rich:结论': { from: 'manual', value: '<p>ok</p>' },
        'if:有风险': { from: 'manual', value: 'true' },
      },
      FIELDS,
      oddRows,
    )
    const { buffer } = await bind(templated.buffer, built.data, { stripControls: true })
    const xml = await docOf(buffer)
    expect(xml).toContain('K1-K2')
    expect(xml).toContain('DN400')
    // 行数正确
    expect((xml.match(/K\d-K\d/g) || []).length).toBe(2)
  })

  it('rich 内 {{字段名}} 动态引用替换为上游值', async () => {
    const templated = await sdtifyDocx(await makeDocx(DOC))
    const { bindings } = await parseTemplate(templated.buffer)
    const built = buildData(
      bindings,
      {
        'rich:结论': { from: 'manual', value: '<p>本工程全长 {{长度}} 米，项目为 {{项目名称}}。</p>' },
        'if:有风险': { from: 'manual', value: 'true' },
      },
      FIELDS,
      ROWS,
    )
    const { buffer } = await bind(templated.buffer, built.data, { stripControls: true })
    const xml = await docOf(buffer)
    expect(xml).toContain('本工程全长 128.4 米，项目为 示例路改造工程。')
    // 不存在的字段保留原文并记入 missing
    const built2 = buildData(
      bindings,
      { 'rich:结论': { from: 'manual', value: '<p>{{不存在的字段}}</p>' } },
      FIELDS,
      ROWS,
    )
    expect(built2.missing.some((m) => m.includes('不存在的字段'))).toBe(true)
  })
})

/* ---------- 富文本图片 ---------- */

// 8x5 的纯色 PNG，够用来验证「字节进部件 + rId 对得上」
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAgAAAAFCAYAAAB4ka1VAAAAHElEQVQI12P4//8/AzYwiFEMDAwMDAwMDAwMDAAsEQMYlQAAAABJRU5ErkJggg=='

describe('富文本图片：落盘与外部降级', () => {
  it('data: URI 图片写进 media / rels / Content_Types，rId 对得上', async () => {
    const templated = await sdtifyDocx(await makeDocx(DOC))
    const { bindings } = await parseTemplate(templated.buffer)
    const built = buildData(
      bindings,
      {
        'rich:结论': {
          from: 'manual',
          value: `<p>现场照片：</p><p><img src="data:image/png;base64,${PNG}" alt="现场照片"></p>`,
        },
        'if:有风险': { from: 'manual', value: 'true' },
      },
      FIELDS,
      ROWS,
    )
    const { buffer } = await bind(templated.buffer, built.data, { stripControls: true })
    const zip = await JSZip.loadAsync(buffer)

    const media = Object.keys(zip.files).filter((f) => /^word\/media\/.+\.png$/.test(f))
    expect(media.length).toBe(1)

    const xml = await zip.file('word/document.xml')!.async('string')
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    const ct = await zip.file('[Content_Types].xml')!.async('string')

    expect(xml).toContain('<w:drawing')
    expect(rels).toContain('/relationships/image')
    expect(ct).toContain('Extension="png"')

    // 关键：drawing 引用的 rId 必须在 rels 里真实存在，否则 Word 报文档损坏
    const rId = /r:embed="(rId\d+)"/.exec(xml)?.[1]
    expect(rId).toBeTruthy()
    expect(rels).toContain(`Id="${rId}"`)
    // 且 rels 指向的正是刚落进 media 的那个部件
    const target = new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(rels)?.[1]
    expect(target).toBeTruthy()
    expect(zip.file(`word/${target}`)).not.toBeNull()
    // alt 文字跟着进了 docPr（读屏与 Word 的「图片格式」面板要用）
    expect(xml).toContain('现场照片')
  })

  it('外链图片降级为可见文字，不静默消失', async () => {
    const templated = await sdtifyDocx(await makeDocx(DOC))
    const { bindings } = await parseTemplate(templated.buffer)
    const built = buildData(
      bindings,
      {
        'rich:结论': {
          from: 'manual',
          value: '<p>见图：</p><p><img src="https://cdn.example.com/a.png" alt="管网示意图"></p><p>以上。</p>',
        },
        'if:有风险': { from: 'manual', value: 'true' },
      },
      FIELDS,
      ROWS,
    )
    const { buffer } = await bind(templated.buffer, built.data, { stripControls: true })
    const xml = await docOf(buffer)
    expect(xml).toContain('管网示意图')
    // 没有 sink 可用 → 不产生任何 media 部件
    const media = Object.keys((await JSZip.loadAsync(buffer)).files).filter((f) => f.startsWith('word/media/'))
    expect(media.length).toBe(0)
  })

  it('回归：只含图片的段落不会被吞掉', async () => {
    // 曾经 <p><img></p> 因为从不调用 ensure() 而让整段凭空消失——静默数据丢失
    const templated = await sdtifyDocx(await makeDocx(DOC))
    const { bindings } = await parseTemplate(templated.buffer)
    const built = buildData(
      bindings,
      {
        'rich:结论': {
          from: 'manual',
          value: `<p>前</p><p><img src="data:image/png;base64,${PNG}"></p><p>后</p>`,
        },
        'if:有风险': { from: 'manual', value: 'true' },
      },
      FIELDS,
      ROWS,
    )
    const { buffer } = await bind(templated.buffer, built.data, { stripControls: true })
    const zip = await JSZip.loadAsync(buffer)
    const xml = await zip.file('word/document.xml')!.async('string')
    // 拆出 rich 控件原本所在的那三段的连续串，确认前后文都在
    const seg = /前[\s\S]{0,4000}后/.exec(xml)?.[0] ?? ''
    expect(seg).toContain('前')
    expect(seg).toContain('后')
    expect(seg).toContain('<w:drawing')
  })
})

/* ---------- 表格跨页（OOXML 真排版指令）---------- */

describe("表格跨页：cantSplit + tblHeader 真写进导出件", () => {
  const CROSS_BODY = [
    para(run('{{field:项目名称}}')),
    '<w:tbl><w:tblGrid>' + Array.from({ length: 2 }, () => '<w:gridCol w:w="1500"/>').join('') + '</w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>' +
      `<w:p>${run('{{each:清单}}{{名称}}')}</w:p></w:tc>` +
      '<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>' +
      `<w:p>${run('{{数量}}')}</w:p></w:tc></w:tr></w:tbl>`,
    para(run('{{rich:结论}}')),
    para(run('{{if:有风险}}风险提示：需支护。')),
  ].join('')

  const CROSS_DOC = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${CROSS_BODY}</w:body></w:document>`

  const mk = () => sdtifyDocx(makeDocx(CROSS_DOC))
  /**
   * 跨页行为（用户 2026-09-28：「根本不能让我做到跨页。甚至表头重复出现」）
   *
   * 这两条是**真写进 OOXML** 的排版指令，Word/WPS 打开导出件即生效：
   *   - <w:cantSplit/>   → 行不被分页拦腰拆开（放不下就整行推到下一页）
   *   - <w:tblHeader/>   → 首行在每页顶部重复（Word 的「标题行重复」）
   *
   * 这里断言到 XML 原文级：只加在正确的位置（trPr 内、tr 的第一个子元素，
   * 这是 OOXML 的强制顺序），且 tblHeader 只出现在首行。
   */
  it('跨页：所有行带 cantSplit，且只有首行带 tblHeader', async () => {
    const templated = await sdtifyDocx(await makeDocx(DOC))
    const { bindings } = await parseTemplate(templated.buffer)
    // 4 行 2 列的表：首行当表头，后 3 行是数据
    const rows = ['名称,数量', 'A,1', 'B,2', 'C,3'].map((l) => {
      const [a, b] = l.split(',')
      return { 名称: a, 数量: b }
    })
    const built = buildData(
      bindings,
      {
        'rich:结论': { from: 'manual', value: '<p>结论</p>' },
        'if:有风险': { from: 'manual', value: 'true' },
      },
      FIELDS,
      rows,
    )
    const { buffer } = await bind(templated.buffer, built.data, { stripControls: true })
    const xml = await docOf(buffer)

    // cantSplit：每个 w:tr 都有（表格行数 ≥ 2）
    const trCount = (xml.match(/<w:tr>/g) ?? []).length
    const cantCount = (xml.match(/<w:cantSplit\/>/g) ?? []).length
    expect(trCount).toBeGreaterThanOrEqual(2)
    expect(cantCount).toBe(trCount)

    // tblHeader：只出现 1 次（仅首行）
    expect((xml.match(/<w:tblHeader\/>/g) ?? []).length).toBe(1)

    // 位置正确性：<w:trPr><w:cantSplit/><w:tblHeader/></w:trPr> 必须紧跟 <w:tr>
    //（OOXML 规定 trPr 是 tr 的第一个子元素；顺序错了 Word 会忽略整块属性）
    expect(xml).toMatch(/<w:tr><w:trPr><w:cantSplit\/><w:tblHeader\/><\/w:trPr>/)
    // 数据行只有 cantSplit，没有 tblHeader
    expect(xml).not.toMatch(/<w:tblHeader\/><w:cantSplit\/>/)
    // 不存在「tblHeader 出现在 tc 之后」的错误结构
    expect(xml).not.toMatch(/<w:tc>[\s\S]{0,80}<w:tblHeader/)
  })

  it('跨页：单行表不加 tblHeader（避免 Word 多画一行）', async () => {
    const templated = await sdtifyDocx(await makeDocx(DOC))
    const { bindings } = await parseTemplate(templated.buffer)
    const built = buildData(
      bindings,
      {
        'rich:结论': { from: 'manual', value: '<p>结论</p>' },
        'if:有风险': { from: 'manual', value: 'true' },
      },
      FIELDS,
      [{ 名称: '仅一行', 数量: '1' }],
    )
    const { buffer } = await bind(templated.buffer, built.data, { stripControls: true })
    const xml = await docOf(buffer)
    expect((xml.match(/<w:cantSplit\/>/g) ?? []).length).toBe((xml.match(/<w:tr>/g) ?? []).length)
    expect(xml).not.toContain('<w:tblHeader/>')
  })

  /**
   * 落盘自证：把产物**写到磁盘再读回来解包**，验跨页属性没在序列化环节丢掉。
   *
   * 为什么必须过一遍磁盘（2026-09-28 教训）：
   *   内存里读 document.xml 只能证明「绑定函数返回的 buffer 对」，
   *   证明不了「这份 buffer 存成文件后还是一份 Word 能正确解析的 docx」——
   *   JSZip 序列化 / 部件压缩 / [Content_Types] 缺失都可能让 Word 打不开或忽略属性。
   *   Word 只要打不开文件，cantSplit/tblHeader 写得再对也没意义。
   */
  it('跨页：产物落盘 → 重新解包，属性仍在且文件可被 Word 解析', async () => {
    const templated = await sdtifyDocx(await makeDocx(DOC))
    const { bindings } = await parseTemplate(templated.buffer)
    const built = buildData(
      bindings,
      {
        'rich:结论': { from: 'manual', value: '<p>结论</p>' },
        'if:有风险': { from: 'manual', value: 'true' },
      },
      FIELDS,
      ['名称,数量', 'A,1', 'B,2', 'C,3'].map((l) => {
        const [a, b] = l.split(',')
        return { 名称: a, 数量: b }
      }),
    )
    const { buffer } = await bind(templated.buffer, built.data, { stripControls: true })

    // 真的写盘再读回（走完整个序列化往返）
    const dir = path.join(os.tmpdir(), 'wrp-crosspage-' + Date.now())
    fs.mkdirSync(dir, { recursive: true })
    const p = path.join(dir, 'out.docx')
    fs.writeFileSync(p, buffer)
    try {
      const back = fs.readFileSync(p)
      const zip = await JSZip.loadAsync(back)
      // Word 必需部件齐全 → 文件是「能打开的 docx」
      expect(zip.file('word/document.xml')).toBeTruthy()
      expect(zip.file('[Content_Types].xml')).toBeTruthy()
      const xml = await zip.file('word/document.xml')!.async('string')
      const trs = xml.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) ?? []
      expect((xml.match(/<w:cantSplit\/>/g) ?? []).length).toBe(trs.length)
      expect((xml.match(/<w:tblHeader\/>/g) ?? []).length).toBe(1)
      expect(xml).toMatch(/<w:trPr><w:cantSplit\/><w:tblHeader\/><\/w:trPr>/)
    } finally {
      fs.rmSync(dir, { recursive: true, force: true })
    }
  })

  /**
   * rich 路径（HTML 表 → html2ooxml.tableXml）也必须产出跨页属性。
   *
   * 为什么单列一条：each 走 binder.withRowProps，rich 走 tableXml ——
   * 两条路径完全独立。2026-09-28 就是只改了 tableXml、each 路径漏了，
   * 测试实测 cantSplit=0 才发现。两条都要进测试，缺一条就会静默回归。
   */
  it('跨页：rich 富文本插入的表格同样带跨页属性', async () => {
    const templated = await sdtifyDocx(await makeDocx(DOC))
    const { bindings } = await parseTemplate(templated.buffer)
    const html =
      '<table><tr><th><p>序号</p></th><th><p>名称</p></th></tr>' +
      '<tr><td><p>1</p></td><td><p>甲</p></td></tr>' +
      '<tr><td><p>2</p></td><td><p>乙</p></td></tr></table>'
    const built = buildData(
      bindings,
      {
        'rich:结论': { from: 'manual', value: html },
        'if:有风险': { from: 'manual', value: 'true' },
      },
      FIELDS,
      [{ 名称: 'a', 数量: '1' }],
    )
    const { buffer } = await bind(templated.buffer, built.data, { stripControls: true })
    const xml = await docOf(buffer)
    // rich 插入的表：3 行，全部 cantSplit，首行 tblHeader
    const richTbl = /<w:tbl>[\s\S]*?<\/w:tbl>/g
    const tbls = xml.match(richTbl) ?? []
    const withHeader = tbls.filter((t) => t.includes('<w:tblHeader/>'))
    expect(withHeader.length).toBeGreaterThanOrEqual(1)
    for (const t of withHeader) {
      const trs2 = t.match(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g) ?? []
      expect((t.match(/<w:cantSplit\/>/g) ?? []).length).toBe(trs2.length)
      expect((t.match(/<w:tblHeader\/>/g) ?? []).length).toBe(1)
    }
  })
})
