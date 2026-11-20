/**
 * 图片能力探针：回答两个问题
 *  1. 模板里原有的静态图片，经过 sdtify + bind 之后还在不在？
 *  2. 富文本（rich）里写 <img>，能不能变成 docx 里的图片？
 *
 * 跑法：npx vite-node scripts/image-probe.mjs
 *      （必须走 vite-node：源码里的 import 不带扩展名，裸 node 的 ESM 解析会直接报
 *       ERR_MODULE_NOT_FOUND，看起来像"文件丢了"，其实是解析规则不同）
 */
import { readFileSync, writeFileSync } from 'node:fs'
import JSZip from 'jszip'
import { sdtifyDocx } from '../src/docx/sdtify.ts'
import { parseTemplate, bind } from '../src/docx/binder.ts'
import { htmlToOoxml } from '../src/docx/html2ooxml.ts'
import { drawingXml } from '../src/docx/image.ts'
import { samplePng } from './_png.mjs'

const SRC = 'sample-template/市政管道工程量报告-模板.docx'

/* ------------------------------------------------------------------ *
 * 0. 测试图
 * ------------------------------------------------------------------ */
// 320x160 有图案的图（不是 1x1）——1x1 只能证明"管线通了"，
// 看不出缩放比例对不对、宽高比有没有被拉伸。要肉眼验收就得有尺寸有内容。
const PNG_1PX = samplePng(320, 160)
const NATIVE = { w: 320, h: 160 }

/** 把一个真图片部件塞进 docx，并在 document.xml 末尾插入一段引用它的段落 */
async function injectImage(originalBytes, fileName = 'image1.png') {
  const zip = await JSZip.loadAsync(originalBytes)
  const docXml = await zip.file('word/document.xml').async('string')
  const relsXml = await zip.file('word/_rels/document.xml.rels').async('string')

  // 1) 图片本体
  zip.file(`word/media/${fileName}`, PNG_1PX)

  // 2) rels 里加一条关系，取一个不冲突的 rId
  const used = [...relsXml.matchAll(/Id="(rId\d+)"/g)].map((m) => Number(m[1].slice(3)))
  const nextId = `rId${Math.max(0, ...used) + 1}`
  const relEntry =
    `<Relationship Id="${nextId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${fileName}"/>`
  zip.file('word/_rels/document.xml.rels', relsXml.replace('</Relationships>', `${relEntry}</Relationships>`))

  // 3) [Content_Types].xml 声明 png 扩展名
  const ct = await zip.file('[Content_Types].xml').async('string')
  if (!ct.includes('Extension="png"')) {
    zip.file(
      '[Content_Types].xml',
      ct.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>'),
    )
  }

  // 4) 正文末尾插一段带图片的段落。
  //    片段直接复用生产代码的 drawingXml —— 以前这里手抄了一份 OOXML，
  //    结果生产侧修了命名空间、探针这份没修，做出的"模板自带静态图"本身就是坏的，
  //    于是每个下游校验都在报同一个错。手抄 OOXML 是最容易漂移的地方，不要再抄第二份。
  const EMU = 9525 // 1px = 9525 EMU
  // 显示尺寸按原始宽高比 2:1 缩放，方便肉眼确认没有拉伸
  const W = 240
  const H = Math.round((W * NATIVE.h) / NATIVE.w)
  const drawing = `<w:p><w:r>${drawingXml(nextId, W * EMU, H * EMU, 'ProbeImage')}</w:r></w:p>`
  zip.file('word/document.xml', docXml.replace('</w:body>', `${drawing}</w:body>`))

  return { buf: await zip.generateAsync({ type: 'nodebuffer' }), nextId }
}

/* ------------------------------------------------------------------ *
 * 1. 静态图片能否穿过 sdtify + bind
 * ------------------------------------------------------------------ */
console.log('══ 探针 1：模板里原有的静态图片 ══')
const raw = readFileSync(SRC)
const { buf: withImg } = await injectImage(raw)
console.log('注入后 docx 含 media 部件:', Object.keys((await JSZip.loadAsync(withImg)).files).filter((f) => f.startsWith('word/media/')))

const sdt = await sdtifyDocx(new Uint8Array(withImg))
console.log('sdtify 控件数:', sdt.count, '| 改动的部件:', sdt.changed)

const parsed = await parseTemplate(sdt.buffer)
const data = Object.fromEntries(parsed.bindings.map((b) => {
  if (b.kind === 'rich') return [b.path, '<p>富文本正文</p>']
  if (b.kind === 'each') return [b.path, []]
  if (b.kind === 'if') return [b.path, true]
  return [b.path, 'X']
}))
const out = await bind(sdt.buffer, data, { stripControls: true })
const outZip = await JSZip.loadAsync(out.buffer)
const outMedia = Object.keys(outZip.files).filter((f) => f.startsWith('word/media/'))
const outDoc = await outZip.file('word/document.xml').async('string')
console.log('bind 后 media 部件:', outMedia)
console.log('bind 后仍含 w:drawing:', /<w:drawing/.test(outDoc), '| 含 a:blip:', /a:blip/.test(outDoc))
writeFileSync('sample-template/带图片-已填充.docx', out.buffer)
console.log('产物已写出: sample-template/带图片-已填充.docx')

/* ------------------------------------------------------------------ *
 * 2. 富文本里写 <img> 会怎样
 * ------------------------------------------------------------------ */
console.log('\n══ 探针 2：富文本里的 <img>（无 sink，应降级为文字）══')
const html = '<p>文字段落</p><p><img src="data:image/png;base64,iVBORw0KGgo=" alt="示意图"></p><p>结尾</p>'
const ooxml = htmlToOoxml(html)
console.log('段落数:', (ooxml.match(/<w:p[ >]/g) ?? []).length, '(输入 3 段，必须还是 3)')
console.log('含降级文字:', /图片：示意图/.test(ooxml))

/* ------------------------------------------------------------------ *
 * 3. 接上 sink 之后，图片应真正落进 docx
 * ------------------------------------------------------------------ */
console.log('\n══ 探针 3：接上 sink 走完整 bind ══')
const PNG_64 = PNG_1PX.toString('base64')
const richHtml =
  '<p>下面是现场照片：</p>' +
  `<p><img src="data:image/png;base64,${PNG_64}" alt="现场照片"></p>` +
  '<p>外链图：<img src="https://example.com/a.png" alt="外链"></p>' +
  '<p>结尾</p>'

const data2 = Object.fromEntries(
  parsed.bindings.map((b) => {
    if (b.kind === 'rich') return [b.path, richHtml]
    if (b.kind === 'each') return [b.path, []]
    if (b.kind === 'if') return [b.path, true]
    return [b.path, 'X']
  }),
)
const out2 = await bind(sdt.buffer, data2, { stripControls: true })
const z2 = await JSZip.loadAsync(out2.buffer)
const media2 = Object.keys(z2.files).filter((f) => f.startsWith('word/media/') && f.endsWith('.png'))
const doc2 = await z2.file('word/document.xml').async('string')
const rels2 = await z2.file('word/_rels/document.xml.rels').async('string')
const ct2 = await z2.file('[Content_Types].xml').async('string')
// 2 = 探针 1 注入的静态 image1.png + 本次富文本那张；外链图不产生部件
console.log('media 部件:', media2.length, '(应为 2：静态 1 + 富文本 1)')
console.log('document.xml 含 w:drawing:', /<w:drawing/.test(doc2))
console.log('rels 里新增了 image 关系:', /relationships\/image/.test(rels2))
console.log('Content_Types 声明了 png:', ct2.includes('Extension="png"'))
console.log('外链图降级为文字:', /图片：外链/.test(doc2), '(不应静默消失)')
const drew = /r:embed="(rId\d+)"/.exec(doc2)
console.log('drawing 引用的 rId:', drew?.[1], '| rels 里存在该 Id:', drew ? rels2.includes(`Id="${drew[1]}"`) : false)
console.log('段落未被吞掉:', (doc2.match(/<w:p[ >]/g) ?? []).length > 3)
writeFileSync('sample-template/富文本插图-已填充.docx', out2.buffer)
console.log('产物已写出: sample-template/富文本插图-已填充.docx')
