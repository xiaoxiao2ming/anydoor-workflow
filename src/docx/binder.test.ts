import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { sdtify } from './sdtify'
import { parseTemplate, bind } from './binder'

/**
 * 用 JSZip 手工构造最小 docx（第三方/设计院视角的普通文档），
 * 走完整链路：sdtify 控件化 → parse 绑定清单 → bind 主路径 → bind 增量路径。
 * 这套断言与独立验证项目 docx-bind-studio 的结论一致。
 */

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const run = (text: string, bold = false, color = '') => {
  const rpr = bold || color ? `<w:rPr>${bold ? '<w:b/>' : ''}${color ? `<w:color w:val="${color}"/>` : ''}</w:rPr>` : ''
  return `<w:r>${rpr}<w:t xml:space="preserve">${text}</w:t></w:r>`
}
const para = (...runs: string[]) => `<w:p>${runs.join('')}</w:p>`

const cell = (text: string, head = false) =>
  `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/>${head ? '<w:shd w:val="clear" w:fill="D9E2F3"/>' : ''}</w:tcPr>${para(run(text, head))}</w:tc>`

/** 模板正文：与真实业务同构 —— 固定文字混排变量、循环表格、富文本、条件段 */
const BODY = [
  para(run('{{field:项目名称}}', true, '185FA5')),
  para(run('编制日期：'), run('{{field:编制日期}}')),
  para(run('本工程管线总长 '), run('{{field:管线总长}}', true, 'A32D2D'), run(' 米，共设检查井 '), run('{{field:检查井数量}}', true, 'A32D2D'), run(' 座。')),
  '<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="7F7F7F"/><w:bottom w:val="single" w:sz="4" w:color="7F7F7F"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>' +
    `<w:tr>${cell('管段编号', true)}${cell('管径', true)}${cell('长度', true)}</w:tr>` +
    `<w:tr>${cell('{{each:管道清单}}{{管段编号}}')}${cell('{{管径}}')}${cell('{{长度}}')}</w:tr>` +
    '</w:tbl>',
  para(run('结论与建议', true)),
  para(run('{{rich:结论与建议}}')),
  para(run('{{if:需风险提示}}风险提示：存在交叉施工。', true, 'BA7517')),
].join('')

const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${BODY}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1134" w:bottom="1440" w:left="1134"/></w:sectPr></w:body></w:document>`

async function makeDocx(documentXml: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  )
  zip.file('word/document.xml', documentXml)
  return zip.generateAsync({ type: 'uint8array' })
}

const dataV1 = {
  项目名称: '示例路市政管道改造工程',
  编制日期: '2026-09-03',
  管线总长: 1286.5,
  检查井数量: 42,
  管道清单: [
    { 管段编号: 'W1-W2', 管径: 'DN300', 长度: 128.4 },
    { 管段编号: 'W2-W3', 管径: 'DN400', 长度: 216.7 },
    { 管段编号: 'W3-W4', 管径: 'DN500', 长度: 342.0 },
  ],
  结论与建议:
    '<p>管线布置<b>满足规范</b>：</p><ol><li>全线重力流；</li><li>埋深 <span style="color:#C00000">3.4 m</span> 需支护。</li></ol><h3>汇总</h3><table><tr><th>管径</th><th>长度</th></tr><tr><td>DN300</td><td>128.4</td></tr></table>',
  需风险提示: true,
}

const dataV2 = {
  ...dataV1,
  项目名称: '示例路市政管道改造工程（修编版）',
  管线总长: 445.1,
  检查井数量: 18,
  管道清单: [
    { 管段编号: 'W1-W2', 管径: 'DN300', 长度: 128.4 },
    { 管段编号: 'W2-W3', 管径: 'DN600', 长度: 316.7 },
  ],
  结论与建议: '<p>取消 DN800 段，<b>总投资下调 18%</b>。</p>',
  需风险提示: false,
}

const count = (s: string, re: RegExp) => (s.match(re) || []).length
const docOf = async (buf: Uint8Array) => (await JSZip.loadAsync(buf)).file('word/document.xml')!.async('string')

describe('sdtify：占位符 → 内容控件', () => {
  it('识别全部 7 个绑定声明并按正确粒度包裹', async () => {
    // 注意：行内的 {{管段编号}}/{{管径}}/{{长度}} 是 each 行内填充符（无 kind 前缀），
    // 由 fillPlaceholders 处理，不生成内容控件，所以是 7 个而非 10 个。
    const { xml, count: n, tags } = sdtify(DOC_XML)
    expect(n).toBe(7)
    expect(tags).toContain('field:项目名称')
    expect(tags).toContain('rich:结论与建议')
    expect(tags).toContain('each:管道清单')
    expect(tags).toContain('if:需风险提示')
    // field 是 run 级：段落里的固定文字不被包进控件
    expect(xml).toMatch(/<w:r><w:t xml:space="preserve">编制日期：<\/w:t><\/w:r>/)
    // each 标记被剔除，不再出现在行内
    expect(xml).not.toContain('{{each:管道清单}}')
    // if 标记被剔除
    expect(xml).not.toContain('{{if:需风险提示}}')
  })
})

describe('binder 主路径：模板 + 数据 → 产物', () => {
  it('v1 全量填充', async () => {
    const raw = await makeDocx(sdtify(DOC_XML).xml)
    const { bindings } = await parseTemplate(raw)
    expect(bindings.length).toBe(7)

    const { buffer, bound } = await bind(raw, dataV1)
    expect(bound).toBe(7)
    const xml = await docOf(buffer)

    expect(xml).toContain('示例路市政管道改造工程')
    expect(xml).toContain('1286.5')
    // 段落固定文字没被吃掉
    expect(xml).toContain('米，共设检查井')
    expect(xml).toContain(' 座。')
    // each 展开 3 行
    expect(count(xml, /W\d-W\d/g)).toBe(3)
    // rich：真列表 + 颜色 + 嵌套表格
    expect(xml).toContain('w:numId w:val="101"')
    expect(xml).toContain('w:val="C00000"')
    expect(count(xml, /<w:tbl>/g)).toBe(2)
    // if 为真保留
    expect(xml).toContain('风险提示：存在交叉施工')
    // numbering 自动注入
    const zip = await JSZip.loadAsync(buffer)
    expect(zip.file('word/numbering.xml')).toBeTruthy()
  })

  it('v2 数据全变（行数 3→2、条件段消失）', async () => {
    const raw = await makeDocx(sdtify(DOC_XML).xml)
    const { buffer } = await bind(raw, dataV2)
    const xml = await docOf(buffer)
    expect(count(xml, /W\d-W\d/g)).toBe(2)
    expect(xml).toContain('445.1')
    expect(xml).toContain('修编版')
    expect(xml).not.toContain('风险提示：存在交叉施工')
    expect(xml).toContain('总投资下调 18%')
  })

  it('stripControls 产出纯净交付稿', async () => {
    const raw = await makeDocx(sdtify(DOC_XML).xml)
    const { buffer } = await bind(raw, dataV1, { stripControls: true })
    const xml = await docOf(buffer)
    expect(xml).not.toContain('<w:sdt>')
    expect(xml).toContain('示例路市政管道改造工程')
  })
})

describe('binder 增量路径：产物 → 产物（只传变化字段）', () => {
  it('field 更新到位，未传的 rich/each/if 原样保留', async () => {
    const raw = await makeDocx(sdtify(DOC_XML).xml)
    const v1 = await bind(raw, dataV1)
    const { buffer } = await bind(v1.buffer, {
      项目名称: '修编二版',
      管线总长: 999.9,
      检查井数量: 30,
    })
    const xml = await docOf(buffer)
    // field 已更新（管线总长出现两处：正文 + 无合计，模板里仅正文一处 → 全部更新）
    expect(xml).toContain('999.9')
    expect(xml).toContain('修编二版')
    // 未传字段原样保留
    expect(xml).toContain('风险提示：存在交叉施工')
    expect(count(xml, /W\d-W\d/g)).toBe(3)
    expect(xml).toContain('w:numId w:val="101"')
  })

  it('each 产物态行数固定并给出警告', async () => {
    const raw = await makeDocx(sdtify(DOC_XML).xml)
    const v1 = await bind(raw, dataV1) // 3 行
    const { warnings } = await bind(v1.buffer, { 管道清单: [{ 管段编号: 'X1' }] }) // 1 行
    expect(warnings.some((w) => w.includes('已是展开态'))).toBe(true)
  })
})
