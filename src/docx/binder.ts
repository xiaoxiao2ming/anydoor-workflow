/**
 * docx 绑定引擎
 *
 * 约定（写在内容控件的 tag 里，与 Word 原生概念一致，设计院能在 Word 里看见）：
 *   field:项目名      段落 / 数字，替换文本，保留原 run 的字体、颜色、加粗
 *   rich:结论         富文本，HTML → OOXML，支持粗斜体、列表、链接、表格、图片
 *   each:管道清单     表格循环行，行模板克隆 N 份（行内用 ｛｛字段名｝｝ 短占位符）
 *   if:有风险         条件段落，真则保留，假则整段删除
 *   image:现场照片    整段一张图，值是 data: URI（可来自上游字段）
 *
 * 关键性质：控件外壳常驻 → field 支持「增量更新」（只传变化的字段，其余原样保留）。
 * rich / each / if 的主路径永远是「模板 + 数据 → 产物」：数据变化从模板重生成。
 */

import JSZip from 'jszip'
import { findSdtBlocks, rewrapSdt, unwrapSdt } from './sdt'
import { htmlToOoxml, esc } from './html2ooxml'
import { createImageSink, imageBlockXml } from './image'

export type BindKind = 'field' | 'rich' | 'each' | 'if' | 'image'

export interface BindingInfo {
  kind: BindKind
  path: string
  tag: string
  part: string
  hasRow: boolean
  sample: string
  /** each 专属：循环行内的短占位符名列表（｛｛管径｝｝ → '管径'），供列映射 UI */
  placeholders: string[]
}

export interface BindResult {
  buffer: Uint8Array
  bound: number
  warnings: string[]
}

/** 'field:项目名' → {kind:'field', path:'项目名'} */
export function parseTag(raw: string): { kind: BindKind; path: string } | null {
  const m = /^\s*(field|rich|each|if|image)\s*:\s*([\s\S]+?)\s*$/.exec(String(raw || ''))
  return m ? { kind: m[1] as BindKind, path: m[2] } : null
}

export function getPath(data: unknown, path: string): unknown {
  return String(path)
    .split('.')
    .reduce<unknown>((o, k) => (o == null ? undefined : Array.isArray(o) ? o[Number(k)] : (o as Record<string, unknown>)[k]), data)
}

export function fmtValue(v: unknown): string {
  if (v == null) return ''
  if (v instanceof Date) return v.toISOString().slice(0, 10)
  if (typeof v === 'boolean') return v ? '是' : '否'
  if (typeof v === 'object') return JSON.stringify(v)
  return String(v)
}

const PARTS = ['word/document.xml', 'word/header1.xml', 'word/footer1.xml', 'word/header2.xml', 'word/footer2.xml']

/* ---------- 占位符替换（保留 run 格式） ---------- */

const T_RE = /(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/g

/** 把 XML 中首个 w:t 设为 value，其余清空（保留首个 run 的字体/颜色/加粗） */
function fillText(xml: string, value: unknown): string {
  const text = fmtValue(value)
  let done = false
  const out = xml.replace(T_RE, (_all, open: string, _t: string, close: string) => {
    if (done) return `${open}${close}`
    done = true
    return `${open}${esc(text)}${close}`
  })
  if (!done) {
    // 控件内没有 w:t（空段落等）——补一个
    return xml.replace(/(<w:p(?:\s[^>]*)?>)/, `$1<w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`)
  }
  return out
}

/** 替换 ｛｛key｝｝ 形式的短占位符，key 相对行数据解析 */
function fillPlaceholders(xml: string, row: unknown): string {
  return xml.replace(T_RE, (all, open: string, t: string, close: string) => {
    if (!t.includes('{{')) return all
    const nt = t.replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (_m, key: string) => {
      // 占位符可能是 '名称'，也可能是 sdtify 留下的 'field:名称' —— 后者要先剥掉 kind 前缀
      const raw = key.trim()
      const parsed = parseTag(raw)
      const path = parsed ? parsed.path : raw
      const v = getPath(row, path)
      return v === undefined ? '' : fmtValue(v)
    })
    return `${open}${esc(nt)}${close}`
  })
}

/* ---------- 表格行跨页属性：cantSplit（行不拆开）+ tblHeader（表头每页重复） ---------- */

/**
 * 给一行 <w:tr> 补跨页属性。
 *
 * OOXML 规则（错了 Word 会整块忽略）：
 *   1. <w:trPr> 必须是 <w:tr> 的**第一个**子元素，顺序 w:trPr → w:tc…；
 *   2. <w:cantSplit/> 让本行不被打页拦腰拆开（放不下整行推到下一页）——所有行都要；
 *   3. <w:tblHeader/> 让本行在每页顶部重复出现——**只给首行**，且整表至少 2 行
 *      （单行表加 tblHeader 个别 Word 版本会多画一行）。
 *
 * 幂等：模板行可能已经带 trPr（要合并而不是重复插），也可能不带（要新建）。
 */
function withRowProps(tr: string, isFirst: boolean, total: number): string {
  const wantsHeader = isFirst && total > 1
  const inner: string[] = ['<w:cantSplit/>']
  if (wantsHeader) inner.push('<w:tblHeader/>')
  const want = `<w:trPr>${inner.join('')}</w:trPr>`

  const m = /<w:tr(?:\s[^>]*)?>(<w:trPr(?:\s[^>]*)?>[\s\S]*?<\/w:trPr>)?/.exec(tr)
  if (!m) return tr

  if (m[1]) {
    // 已有 trPr：去掉旧的 cantSplit / tblHeader，再按本次语义重写（避免重复标签）
    const body = m[1]
      .replace(/<w:trPr(?:\s[^>]*)?>/, '')
      .replace(/<\/w:trPr>$/, '')
      .replace(/<w:cantSplit\/>/g, '')
      .replace(/<w:tblHeader\/>/g, '')
    const merged = `<w:trPr>${inner.join('')}${body}</w:trPr>`
    return tr.slice(0, m.index) + tr.slice(m.index).replace(m[1], merged)
  }
  // 没有 trPr：插在 <w:tr> 开标签之后（必须是第一个子元素）
  return tr.replace(/^(<w:tr(?:\s[^>]*)?>)/, `$1${want}`)
}

/* ---------- 列表编号：确保模板里有 numId 100(项目符号) / 101(十进制) ---------- */

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const ABSTRACT_100 =
  `<w:abstractNum w:abstractNumId="100"><w:multiLevelType w:val="hybridMultilevel"/>` +
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="●"/>` +
  `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>` +
  `<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:cs="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>`

const ABSTRACT_101 =
  `<w:abstractNum w:abstractNumId="101"><w:multiLevelType w:val="hybridMultilevel"/>` +
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>` +
  `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>`

const NUM_DEFS =
  ABSTRACT_100 +
  ABSTRACT_101 +
  `<w:num w:numId="100"><w:abstractNumId w:val="100"/></w:num>` +
  `<w:num w:numId="101"><w:abstractNumId w:val="101"/></w:num>`

/** 富文本列表需要 numbering 定义；设计院模板往往没有 —— 自动补齐，缺失时降级为字面符号 */
export async function ensureNumbering(zip: JSZip): Promise<{ bullet: number; decimal: number }> {
  const relsPath = 'word/_rels/document.xml.rels'
  const ctPath = '[Content_Types].xml'
  const numPath = 'word/numbering.xml'
  const existing = zip.file(numPath)

  if (!existing) {
    zip.file(numPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering ${NS}>${NUM_DEFS}</w:numbering>`)
    const ctFile = zip.file(ctPath)
    if (ctFile) {
      const ct = await ctFile.async('string')
      if (!ct.includes('/word/numbering.xml')) {
        zip.file(
          ctPath,
          ct.replace(
            '</Types>',
            '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'
          )
        )
      }
    }
    const relsFile = zip.file(relsPath)
    if (relsFile) {
      const rels = await relsFile.async('string')
      if (!rels.includes('numbering.xml')) {
        zip.file(
          relsPath,
          rels.replace(
            '</Relationships>',
            '<Relationship Id="rIdNumBind" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'
          )
        )
      }
    }
    return { bullet: 100, decimal: 101 }
  }

  const num = await existing.async('string')
  if (!/w:numId="100"/.test(num)) {
    zip.file(numPath, num.replace('</w:numbering>', `${NUM_DEFS}</w:numbering>`))
  }
  return { bullet: 100, decimal: 101 }
}

/* ---------- 解析绑定清单 ---------- */

function extractSample(contentXml: string): string {
  const texts = [...contentXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => m[1])
  return texts.join('').trim().slice(0, 60)
}

/** 导入模板 → 绑定清单（UI 里就是「拖入 docx 后，数据面板自动长出来的字段」） */
export async function parseTemplate(
  input: ArrayBuffer | Uint8Array | Buffer
): Promise<{ bindings: BindingInfo[]; warnings: string[] }> {
  const zip = await JSZip.loadAsync(input)
  const bindings: BindingInfo[] = []
  const warnings: string[] = []
  for (const name of PARTS) {
    const f = zip.file(name)
    if (!f) continue
    const xml = await f.async('string')
    for (const b of findSdtBlocks(xml)) {
      const parsed = parseTag(b.tag)
      if (!parsed) {
        warnings.push(`${name}: 控件 tag「${b.tag}」不符合 field:/rich:/each:/if: 约定，已忽略`)
        continue
      }
      if (parsed.kind === 'each' && !b.hasRow) {
        warnings.push(`${name}: each 控件「${parsed.path}」内部没有表格行，无法循环`)
      }
      // each：抓行内短占位符（列名），供 UI 做「模板列 → 上游字段」映射
      const placeholders =
        parsed.kind === 'each'
          ? [...new Set([...b.contentXml.matchAll(/\{\{\s*([^{}]+?)\s*\}\}/g)].map((m) => m[1].trim()))]
          : []
      bindings.push({
        kind: parsed.kind,
        path: parsed.path,
        tag: b.tag,
        part: name,
        hasRow: b.hasRow,
        sample: extractSample(b.contentXml),
        placeholders,
      })
    }
  }
  return { bindings, warnings }
}

/* ---------- 绑定 ---------- */

export interface BindOptions {
  /** true 时剥掉控件外壳，产出纯净交付稿（Word 里看不到控件框） */
  stripControls?: boolean
}

/**
 * 绑定数据到模板。
 * 数据对象：{ 路径: 值 }；each 的值是对象数组（行内字段按名匹配 ｛｛名字｝｝）。
 * 未提供（undefined）的字段保持原样 —— 这是增量更新的基础。
 */
export async function bind(
  input: ArrayBuffer | Uint8Array | Buffer,
  data: Record<string, unknown>,
  opts: BindOptions = {}
): Promise<BindResult> {
  const zip = await JSZip.loadAsync(input)
  const num = await ensureNumbering(zip)
  const warnings: string[] = []
  let bound = 0

  // 富文本插图用的资源通道：每个部件建一次 slot（rId 从该部件现有关系往后排）
  const images = await createImageSink(zip)

  for (const name of PARTS) {
    const f = zip.file(name)
    if (!f) continue
    // 预建 slot，保证转换过程中 img 回调可用（createImageSink 内部要读 rels，必须提前）
    await images.ensure(name)
    let xml = await f.async('string')
    const blocks = findSdtBlocks(xml)
    if (!blocks.length) continue

    // 从后往前处理，保证下标不失效
    for (let i = blocks.length - 1; i >= 0; i--) {
      const b = blocks[i]
      const parsed = parseTag(b.tag)
      if (!parsed) continue
      const value = getPath(data, parsed.path)
      // 未提供该字段 → 保持原样（增量更新）
      if (value === undefined) continue
      let newContent: string

      switch (parsed.kind) {
        case 'field':
          newContent = fillText(b.contentXml, value)
          break

        case 'rich': {
          // 继承模板段落属性（缩进 / 行距 / 对齐），只换内容不换版式
          const pPr = /<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>/.exec(b.contentXml)
          const html = htmlToOoxml(value == null ? '' : String(value), num, (bytes, ext) =>
            images.add(bytes, ext, name),
          )
          newContent = pPr && pPr.index === 0 ? html.replace(/<w:p>/g, `<w:p>${pPr[0]}`) : html
          break
        }

        case 'each': {
          const rows = Array.isArray(value) ? value : []
          const trs = [...b.contentXml.matchAll(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g)].map((m) => m[0])
          if (!trs.length) {
            warnings.push(`each「${parsed.path}」控件内没有找到表格行，原样保留`)
            newContent = b.contentXml
            break
          }
          if (trs.length === 1) {
            // 模板态：这是唯一能自由改变行数的时机
            newContent = rows
              .map((r, i2) => withRowProps(fillPlaceholders(trs[0], r), i2 === 0, rows.length))
              .join('')
          } else {
            // 产物态：行模板已被上一轮展开覆盖，行数固定，只能逐行更新
            newContent = trs
              .map((tr, i2) => withRowProps(fillPlaceholders(tr, rows[i2] ?? {}), i2 === 0, trs.length))
              .join('')
            if (rows.length !== trs.length) {
              warnings.push(
                `each「${parsed.path}」已是展开态（${trs.length} 行），本次数据 ${rows.length} 行 —— 行数以文档为准。要改变行数，必须从模板重新生成。`
              )
            }
          }
          break
        }

        case 'image': {
          // 值来自上游字段（base64 data: URI）或 UI 选的图；拿不到就降级为文字
          const ok = imageBlockXml(value, (bytes, ext) => images.add(bytes, ext, name))
          if (!ok.includes('<w:drawing')) warnings.push(`image「${parsed.path}」的值不是可识别的图片，已降级为文字`)
          newContent = ok
          break
        }

        case 'if':
          // 为真：保留外壳与内容；为假：整块删除（含控件，因此 if 只能靠模板重生成来恢复）
          if (value) {
            newContent = b.contentXml
            break
          }
          xml = xml.slice(0, b.start) + xml.slice(b.end)
          bound++
          continue

        default:
          continue
      }

      xml = opts.stripControls ? xml.slice(0, b.start) + newContent + xml.slice(b.end) : rewrapSdt(xml, b, newContent)
      bound++
    }

    zip.file(name, xml)
  }

  // 图片关系与 Content_Types 必须在打包前写回，否则 Word 打不开或丢图
  await images.finish()

  const out = await zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
  return { buffer: out, bound, warnings }
}
