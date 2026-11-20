/**
 * 生成一份可直接导入试用的示例模板：市政管道工程量报告。
 *
 * 这份模板刻意做成「设计院视角的普通 Word 文档」—— 只有 ｛｛...｝｝ 标记，
 * 没有任何内容控件，用来验证「导入现有模板 → 自动控件化 → 绑定数据」全链路。
 *
 * 运行：node scripts/make-sample-template.mjs
 * 产物：sample-template/市政管道工程量报告-模板.docx
 */

import { mkdir, writeFile } from 'fs/promises'
import { dirname, join } from 'path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'

const HERE = dirname(fileURLToPath(import.meta.url))
const OUT = join(HERE, '..', 'sample-template')

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

const r = (t, o = {}) => {
  const pr = [
    o.b ? '<w:b/>' : '',
    o.i ? '<w:i/>' : '',
    o.color ? `<w:color w:val="${o.color}"/>` : '',
    o.sz ? `<w:sz w:val="${o.sz}"/>` : '',
  ].join('')
  return `<w:r>${pr ? `<w:rPr>${pr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`
}
const p = (runs, o = {}) => {
  const pr = [
    o.style ? `<w:pStyle w:val="${o.style}"/>` : '',
    o.jc ? `<w:jc w:val="${o.jc}"/>` : '',
    o.ind ? `<w:ind w:firstLine="${o.ind}"/>` : '',
  ].join('')
  return `<w:p>${pr ? `<w:pPr>${pr}</w:pPr>` : ''}${Array.isArray(runs) ? runs.join('') : runs}</w:p>`
}

/* ---------- 正文 ---------- */

const body = [
  p(r('市政管道工程量报告', { b: true, sz: '36' }), { style: 'Title', jc: 'center' }),

  p([
    r('工程名称：'),
    r('{{field:工程名称}}', { b: true, color: '185FA5' }),
  ]),
  p([
    r('编制日期：'),
    r('{{field:编制日期}}'),
    r('    编制人：'),
    r('{{field:编制人}}'),
  ]),

  p(r('一、工程概况', { b: true, sz: '28' }), { style: 'Heading1' }),
  p(
    [
      r('本工程管线总长 '),
      r('{{field:管线总长}}', { b: true, color: 'A32D2D' }),
      r(' 米，共设检查井 '),
      r('{{field:检查井数量}}', { b: true, color: 'A32D2D' }),
      r(' 座，总投资约 '),
      r('{{field:总投资}}', { b: true, color: 'A32D2D' }),
      r(' 万元。'),
    ],
    { ind: '480' },
  ),

  p(r('二、管道清单', { b: true, sz: '28' }), { style: 'Heading1' }),
  `<w:tbl><w:tblPr><w:tblW w:w="8500" w:type="dxa"/><w:tblBorders>${[
    'top',
    'left',
    'bottom',
    'right',
    'insideH',
    'insideV',
  ]
    .map((k) => `<w:${k} w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>`)
    .join('')}</w:tblBorders><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
    `<w:tblGrid>${[1600, 2000, 1400, 1400, 2100].map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` +
    `<w:tr>${['序号', '管段编号', '管径', '长度(m)', '备注']
      .map(
        (h) =>
          `<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr>${p(
            r(h, { b: true }),
            { jc: 'center' },
          )}</w:tc>`,
      )
      .join('')}</w:tr>` +
    // 循环行：each 标记 + 行内短占位符
    `<w:tr>${[r('{{序号}}'), r('{{管段编号}}'), r('{{管径}}'), r('{{长度}}'), r('{{备注}}')]
      .map(
        (run) =>
          `<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/></w:tcPr>${p(run, { jc: 'center' })}</w:tc>`,
      )
      .join('')}</w:tr>`.replace(
      '<w:tr>',
      `<w:tr><!-- each -->`,
    ) +
    `</w:tbl>`,
  // 行级 each 标记放在行首单元格里（sdtify 会按 w:tr 包裹，标记本身会被剔除）
  '',

  p(r('三、结论与建议', { b: true, sz: '28' }), { style: 'Heading1' }),
  p(r('{{rich:结论与建议}}'), { ind: '480' }),

  p(r('{{if:需风险提示}}风险提示：本工程存在交叉施工，施工前须与相关单位协调确认。', { b: true, color: 'BA7517' }), {
    ind: '480',
  }),

  p(''),
  p([r('编制：'), r('{{field:编制人}}'), r('        审核：____________')], { jc: 'right' }),
].join('')

// 把 each 标记塞进循环行的第一个单元格（sdtify 认整行，标记位置不重要）
const bodyWithEach = body.replace('<w:tr><!-- each --></w:tr>', '')

/* 重新构造表格：循环行的首格带 each 标记 */
const tbl = body.match(/<w:tbl>[\s\S]*?<\/w:tbl>/)[0]
const rows = [...tbl.matchAll(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g)].map((m) => m[0])
const loopRow = rows[1].replace(
  /(<w:t xml:space="preserve">)/,
  '$1{{each:管道清单}}',
)
const tblFixed = tbl.replace(rows[1], loopRow)
const bodyFinal = bodyWithEach.replace(tbl, tblFixed).replace('<!-- each -->', '')

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>${bodyFinal}<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1134" w:bottom="1440" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="425"/><w:docGrid w:linePitch="312"/></w:sectPr></w:body></w:document>`
  .replace(/\n/g, '')
  // 去掉上面那行 table 后面残留的空段落拼接痕迹
  .replace('</w:tbl><w:p/>', '</w:tbl>')

const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NS}>${p(
  [r('{{field:工程名称}}'), r('  ·  市政管道工程量报告')],
  { jc: 'center' },
)}</w:hdr>`

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体" w:cs="宋体"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="312" w:lineRule="auto" w:after="120"/><w:ind w:firstLineChars="0"/></w:pPr></w:pPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr><w:qFormat/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:qFormat/></w:style>` +
  `</w:styles>`

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`

const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`

const zip = new JSZip()
zip.file('[Content_Types].xml', contentTypes)
zip.file('_rels/.rels', rootRels)
zip.file('word/document.xml', documentXml)
zip.file('word/_rels/document.xml.rels', docRels)
zip.file('word/styles.xml', stylesXml)
zip.file('word/header1.xml', headerXml)

await mkdir(OUT, { recursive: true })
const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } })
const file = join(OUT, '市政管道工程量报告-模板.docx')
await writeFile(file, buf)

const marks = [...documentXml.matchAll(/\{\{(field|rich|each|if):([^{}]+?)\}\}/g)].map((m) => `${m[1]}:${m[2]}`)
console.log('已生成：', file)
console.log('大小：', (buf.length / 1024).toFixed(1), 'KB')
console.log('绑定标记（正文）：', marks.length, '个')
for (const m of marks) console.log('  -', m)
console.log('绑定标记（页眉）：field:工程名称')
console.log('\n用 Word 打开即可看到这是一份普通文档，导入午马未央时会自动控件化。')
