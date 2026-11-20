/**
 * OOXML → HTML 转换器（零依赖）
 *
 * 用途：Word 节点的就地编辑（WPS 式）。把 document.xml 的 body 转成 HTML
 * 塞进 contenteditable 编辑面；用户改完经 htmlToOoxml 写回（保留 sectPr）。
 *
 * 编辑对象是「绑定产物」（bind stripControls 后的字节）——与 OnlyOffice「在线编辑」
 * 同一语义：手工精修覆盖在数据绑定结果之上，数据重新绑定后编辑稿作废。
 * 产物里已无 sdt 控件，因此往返是纯内容级的，不存在绑定语义丢失问题。
 *
 * 支持（与 html2ooxml 的 HTML 语法子集严格对齐——它吃的正是这里产出的标签）：
 *   w:p → p / h1-h6（styles.xml 的 heading N / pPr 的 outlineLvl 双路判定）
 *   w:jc → text-align · w:numPr → ul/ol（numbering.xml 查 numFmt，多级转嵌套）
 *   w:r 的 b/i/u/strike/vertAlign/color/highlight/sz → 行内标签
 *   w:t / w:br / w:tab → 文本 / <br> / 空格
 *   w:tbl/w:tr/w:tc → table/tr/td（gridSpan → colspan；嵌套表格递归）
 *   w:drawing（a:blip r:embed）与 w:pict（v:imagedata）→ <img data:URI>
 *   w:sdt / w:smartTag / w:hyperlink → 递归渲染内容
 *   mc:AlternateContent → 只取 mc:Choice 分支（Fallback 弃）
 *
 * 忽略：w:sectPr（单独提取，写回时拼回）、bookmark、proofErr、lastRenderedPageBreak。
 */

import JSZip from 'jszip'
import { imageSize } from './image'

export interface DocxBodyHtml {
  /** body 内容 HTML（不含 sectPr） */
  html: string
  /** body 级 sectPr 原文（写回时原样拼回，保纸张/边距/页码） */
  sectPr: string
  /** w:document 根元素属性原文（xmlns 声明等，写回时复用） */
  rootAttrs: string
}

/* ---------- 词法：XML tokenizer ---------- */

type XTok =
  | { t: 'open'; name: string; attrs: Record<string, string> }
  | { t: 'self'; name: string; attrs: Record<string, string> }
  | { t: 'close'; name: string }
  | { t: 'text'; v: string }

function parseAttrs(raw: string): Record<string, string> {
  const a: Record<string, string> = {}
  const re = /([a-zA-Z_:][-\w:.]*)\s*=\s*("([^"]*)"|'([^']*)')/g
  let m: RegExpExecArray | null
  while ((m = re.exec(raw))) a[m[1]] = m[3] ?? m[4] ?? ''
  return a
}

/** XML 比 HTML 规整得多：无注释容错需求，标签名带命名空间前缀（w:p / a:blip） */
function tokenize(xml: string): XTok[] {
  const out: XTok[] = []
  let i = 0
  while (i < xml.length) {
    const lt = xml.indexOf('<', i)
    if (lt < 0) {
      out.push({ t: 'text', v: xml.slice(i) })
      break
    }
    if (lt > i) out.push({ t: 'text', v: xml.slice(i, lt) })
    if (xml.startsWith('<?', lt)) {
      const e = xml.indexOf('?>', lt)
      i = e < 0 ? xml.length : e + 2
      continue
    }
    const gt = xml.indexOf('>', lt)
    if (gt < 0) {
      out.push({ t: 'text', v: xml.slice(lt) })
      break
    }
    const raw = xml.slice(lt, gt + 1)
    const close = raw[1] === '/'
    // 命名空间前缀里的冒号算进名字：w:p、a:blip、mc:AlternateContent
    const nm = raw.slice(close ? 2 : 1).match(/^[a-zA-Z_][\w.:-]*/)
    const name = nm ? nm[0] : ''
    if (close) out.push({ t: 'close', name })
    else if (raw.endsWith('/>')) out.push({ t: 'self', name, attrs: parseAttrs(raw.slice(1, -2)) })
    else out.push({ t: 'open', name, attrs: parseAttrs(raw.slice(1, -1)) })
    i = gt + 1
  }
  return out
}

/** 取从 i 起、与 name 配对闭合之间的 token（不含首尾标签本身） */
function sliceUntil(toks: XTok[], i: number, name: string | string[]): XTok[] {
  const names = Array.isArray(name) ? name : [name]
  let depth = 1
  const out: XTok[] = []
  while (i < toks.length) {
    const tk = toks[i++]
    if ((tk.t === 'open' || tk.t === 'self') && names.includes(tk.name)) {
      if (tk.t === 'open') depth++
      else continue
    } else if (tk.t === 'close' && names.includes(tk.name) && --depth === 0) break
    out.push(tk)
  }
  return out
}

/** XML 实体反转义（w:t 内容 → 纯文本） */
function unescXml(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_m, d: string) => String.fromCodePoint(Number(d)))
    .replace(/&x([0-9a-fA-F]+);/g, (_m, h: string) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&amp;/g, '&')
}

/** 纯文本 → HTML 转义（与 html2ooxml.esc 同规则） */
const escHtml = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

/* ---------- 资源：rels / styles / numbering ---------- */

interface Ctx {
  zip: JSZip
  /** rId → { target, image }（image=true 表示关系类型是图片） */
  rels: Map<string, { target: string; image: boolean }>
  /** styleId → 样式名（heading 1 / Title / 正文…） */
  styles: Map<string, string>
  /** numId → ilvl=0 的 numFmt（'bullet' | 'decimal' | 其他原文） */
  numFmt: Map<string, string>
  /** media 字节缓存（同一张图在往返里只解一次 base64） */
  media: Map<string, string | null>
}

const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

async function buildCtx(zip: JSZip): Promise<Ctx> {
  const rels = new Map()
  const relsFile = zip.file('word/_rels/document.xml.rels')
  if (relsFile) {
    const txt = await relsFile.async('string')
    for (const m of txt.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = /Id="([^"]+)"/.exec(m[0])?.[1]
      const target = /Target="([^"]+)"/.exec(m[0])?.[1]
      const type = /Type="([^"]+)"/.exec(m[0])?.[1]
      if (id && target) rels.set(id, { target: target.replace(/^\/+/, ''), image: type === REL_IMAGE })
    }
  }

  const styles = new Map<string, string>()
  const basedOn = new Map<string, string>()
  const stylesFile = zip.file('word/styles.xml')
  if (stylesFile) {
    const txt = await stylesFile.async('string')
    for (const m of txt.matchAll(/<w:style\b[^>]*>([\s\S]*?)<\/w:style>/g)) {
      const sid = /w:styleId="([^"]+)"/.exec(m[0])?.[1]
      const name = /<w:name\s+w:val="([^"]+)"/.exec(m[1])?.[1]
      const base = /<w:basedOn\s+w:val="([^"]+)"/.exec(m[1])?.[1]
      if (sid && name) styles.set(sid, name)
      if (sid && base) basedOn.set(sid, base)
    }
  }
  // basedOn 一层跳转：中文模板常把样式名挂在父样式上
  for (const [sid, name] of [...styles]) {
    if (!/heading|title/i.test(name)) {
      const parent = basedOn.get(sid)
      const pn = parent ? styles.get(parent) : undefined
      if (pn) styles.set(sid, pn)
    }
  }

  const numFmt = new Map<string, string>()
  const numFile = zip.file('word/numbering.xml')
  if (numFile) {
    const txt = await numFile.async('string')
    const abstractFmt = new Map<string, string>()
    for (const m of txt.matchAll(/<w:abstractNum\b[^>]*>([\s\S]*?)<\/w:abstractNum>/g)) {
      const aid = /w:abstractNumId="([^"]+)"/.exec(m[0])?.[1]
      if (!aid) continue
      const fmt = /<w:lvl\b[^>]*>([\s\S]*?)<\/w:lvl>/.exec(m[1])
      const f = fmt ? /<w:numFmt\s+w:val="([^"]+)"/.exec(fmt[1])?.[1] : undefined
      abstractFmt.set(aid, f ?? 'bullet')
    }
    for (const m of txt.matchAll(/<w:num\b[^>]*>([\s\S]*?)<\/w:num>/g)) {
      const nid = /w:numId="([^"]+)"/.exec(m[0])?.[1]
      const aid = /<w:abstractNumId\s+w:val="([^"]+)"/.exec(m[1])?.[1]
      if (nid && aid) numFmt.set(nid, abstractFmt.get(aid) ?? 'bullet')
    }
  }

  return { zip, rels, styles, numFmt, media: new Map() }
}

const IMG_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
}

/** rId → data: URI（读 word/ + rels target 的字节转 base64；失败返回 null 走文字降级） */
async function mediaDataUri(ctx: Ctx, rId: string): Promise<{ uri: string; w: number; h: number } | null> {
  if (ctx.media.has(rId)) {
    const hit = ctx.media.get(rId)
    return hit ? { uri: hit, w: 0, h: 0 } : null
  }
  const rel = ctx.rels.get(rId)
  if (!rel || !rel.image) {
    ctx.media.set(rId, null)
    return null
  }
  const path = 'word/' + rel.target.replace(/^\.\//, '')
  const f = ctx.zip.file(path)
  if (!f) {
    ctx.media.set(rId, null)
    return null
  }
  const bytes = await f.async('uint8array')
  let bin = ''
  const CH = 0x8000
  for (let i = 0; i < bytes.length; i += CH) bin += String.fromCharCode(...bytes.subarray(i, i + CH))
  const ext = (path.split('.').pop() ?? 'png').toLowerCase()
  const uri = `data:${IMG_MIME[ext] ?? 'image/png'};base64,${btoa(bin)}`
  ctx.media.set(rId, uri)
  const size = imageSize(bytes)
  return { uri, w: size?.w ?? 0, h: size?.h ?? 0 }
}

/* ---------- 块级结构 ---------- */

interface HtmlPara {
  type: 'para'
  heading: string | null
  align: string | null
  list: { kind: 'ul' | 'ol'; level: number } | null
  /** 行内 HTML（已带格式标签） */
  runs: string
  /** 行距倍数（w:spacing line/240），null = 默认 */
  lineHeight?: number | null
  /** 首行缩进字符数（w:ind firstLine/240），null = 无 */
  textIndent?: number | null
}
interface HtmlTable {
  type: 'table'
  /** 行 → 单元格 → 块（支持嵌套表格） */
  rows: { span: number; colwidth?: number[]; blocks: (HtmlPara | HtmlTable)[] }[][]
  /** w:tblGrid 的列宽（px，@96dpi=twips/15）。Tiptap TableCell 靠 td 的 colwidth 属性
   *  解析进 cell attrs → TableView 自动生成 colgroup 锁定列宽 —— 没有它浏览器 auto
   *  布局自行分配，设计院文档的列宽（如 302px 名称列）会被压掉 42% 导致大面积换行。 */
  cols?: number[]
}
type HtmlBlock = HtmlPara | HtmlTable

const EMPTY_PARA: HtmlPara = { type: 'para', heading: null, align: null, list: null, runs: '' }

/** styles.xml 的样式名 / pPr 的 outlineLvl → 标题级（h1-h6 或 null） */
function headingOf(ctx: Ctx, styleId: string | undefined, outlineLvl: number | null): string | null {
  if (styleId) {
    const name = ctx.styles.get(styleId) ?? ''
    if (/^title$/i.test(name)) return 'h1'
    const m = /heading\s*(\d)/i.exec(name)
    if (m) {
      const n = Math.min(6, Math.max(1, Number(m[1])))
      return `h${n}`
    }
  }
  if (outlineLvl != null && outlineLvl >= 0 && outlineLvl < 6) return `h${outlineLvl + 1}`
  return null
}

/* ---------- 行内：w:r → HTML ---------- */

const ON_OFF = (v: string | undefined) => v !== undefined && /^(0|false|none)$/i.test(v)

interface RunFmt {
  b: boolean
  i: boolean
  u: boolean
  s: boolean
  sub: boolean
  sup: boolean
  color: string | null
  highlight: string | null
  szPt: number | null
  /** 字体族（w:rFonts 的 eastAsia/ascii） */
  font: string | null
}

/** 取 rPr 里某开关标签的 w:val（无属性 = 'on'；不存在 = null） */
function tagVal(rpr: string, name: string): string | null {
  const m = new RegExp(`<${name}\\b([^>]*?)/?>`).exec(rpr)
  if (!m) return null
  const v = /w:val="([^"]*)"/.exec(m[1])?.[1]
  return v === undefined ? 'on' : v
}

const fmtOf = (rpr: string): RunFmt => {
  const isOn = (name: string) => {
    const v = tagVal(rpr, name)
    return v != null && !ON_OFF(v)
  }
  const color = tagVal(rpr, 'w:color')
  const hl = tagVal(rpr, 'w:highlight')
  const va = tagVal(rpr, 'w:vertAlign')
  const sz = /<w:sz\s+w:val="(\d+)"/.exec(rpr)?.[1]
  // 字符底纹（w:shd fill，非命名色高亮的写侧形态）也当突出显示回读
  const shdFill = /<w:shd\b[^>]*w:fill="([0-9A-Fa-f]{6})"/.exec(rpr)?.[1]
  const highlight = (hl && !/^none$/i.test(hl) ? hl : null) ?? (shdFill ? shdFill.toUpperCase() : null)
  // 字体族：优先 eastAsia（中文字体），回退 ascii
  const fontsTag = /<w:rFonts\b[^>]*\/?>/.exec(rpr)?.[0] ?? ''
  const font =
    /w:eastAsia="([^"]+)"/.exec(fontsTag)?.[1] ?? /w:ascii="([^"]+)"/.exec(fontsTag)?.[1] ?? null
  return {
    b: isOn('w:b'),
    i: isOn('w:i'),
    u: isOn('w:u'),
    s: isOn('w:strike'),
    sub: va === 'subscript',
    sup: va === 'superscript',
    color: color && !/^(auto|inherit)$/i.test(color) ? color : null,
    highlight,
    szPt: sz ? Number(sz) / 2 : null,
    font: font && !/^(auto|inherit)$/i.test(font) ? font : null,
  }
}

/** rPr 的格式 → 前后包裹标签（span 样式 + 语义标签，与 htmlToOoxml 的解析习惯对齐） */
function wrapFmt(f: RunFmt, inner: string): string {
  if (!inner) return ''
  const css: string[] = []
  if (f.color) css.push(`color:#${f.color}`)
  // 突出显示两种形态：命名色（yellow）直接用；6 位 hex（w:shd fill 回读）补 #
  if (f.highlight) css.push(`background-color:${/^[0-9A-Fa-f]{6}$/.test(f.highlight) ? '#' + f.highlight : f.highlight}`)
  if (f.szPt) css.push(`font-size:${f.szPt}pt`)
  if (f.font) css.push(`font-family:${f.font}`)
  let out = css.length ? `<span style="${css.join(';')}">${inner}</span>` : inner
  if (f.b) out = `<b>${out}</b>`
  if (f.i) out = `<i>${out}</i>`
  if (f.u) out = `<u>${out}</u>`
  if (f.s) out = `<s>${out}</s>`
  if (f.sub) out = `<sub>${out}</sub>`
  if (f.sup) out = `<sup>${out}</sup>`
  return out
}

/**
 * 行内 token → HTML。w:hyperlink / w:sdt / mc:AlternateContent 递归；
 * w:del（修订删除）内容跳过；w:pict（VML 图）尝试 v:imagedata 的 r:id。
 */
async function inlineHtml(toks: XTok[], ctx: Ctx): Promise<string> {
  let out = ''
  let i = 0
  while (i < toks.length) {
    const tk = toks[i++]
    if (tk.t === 'text') {
      out += escHtml(unescXml(tk.v))
      continue
    }
    if (tk.t === 'close') continue
    // open / self
    switch (tk.name) {
      case 'w:r': {
        const inner = sliceUntil(toks, i, 'w:r')
        i += inner.length + 1
        out += await runHtml(inner, ctx)
        break
      }
      case 'w:hyperlink': {
        const inner = sliceUntil(toks, i, 'w:hyperlink')
        i += inner.length + 1
        const rId = tk.attrs['r:id']
        const anchor = tk.attrs['w:anchor']
        const target = rId ? ctx.rels.get(rId)?.target : anchor ? `#${anchor}` : undefined
        const body = await inlineHtml(inner, ctx)
        out = target && /^https?:/i.test(target) ? `${out}<a href="${target}">${body}</a>` : out + body
        break
      }
      case 'w:sdt':
      case 'w:smartTag': {
        const inner = sliceUntil(toks, i, tk.name)
        i += inner.length + 1
        const content = takeSdtContent(inner)
        out += await inlineHtml(content, ctx)
        break
      }
      case 'w:ins': {
        const inner = sliceUntil(toks, i, 'w:ins')
        i += inner.length + 1
        out += await inlineHtml(inner, ctx)
        break
      }
      case 'w:del': {
        // 修订删除：内容不出现在编辑面（接受 = 删掉）
        const inner = sliceUntil(toks, i, 'w:del')
        i += inner.length + 1
        break
      }
      case 'w:delText':
      case 'w:proofErr':
      case 'w:bookmarkStart':
      case 'w:bookmarkEnd':
      case 'w:lastRenderedPageBreak':
      case 'w:commentRangeStart':
      case 'w:commentRangeEnd':
      case 'w:commentReference':
        break
      case 'mc:AlternateContent': {
        const inner = sliceUntil(toks, i, 'mc:AlternateContent')
        i += inner.length + 1
        // 只取 Choice 分支（新式 drawing）；Fallback（VML）弃
        const choice = findBranch(inner, 'mc:Choice')
        out += choice ? await inlineHtml(choice, ctx) : ''
        break
      }
      case 'w:drawing': {
        const inner = sliceUntil(toks, i, 'w:drawing')
        i += inner.length + 1
        out += await drawingImg(inner, ctx)
        break
      }
      case 'w:pict': {
        const inner = sliceUntil(toks, i, 'w:pict')
        i += inner.length + 1
        out += await pictImg(inner, ctx)
        break
      }
      case 'w:br': {
        const inner = sliceUntil(toks, i, 'w:br')
        i += inner.length
        out += '<br>'
        break
      }
      default:
        // 未知行内元素：跳过整个区间，避免其文本被重复输出
        if (tk.t === 'open') {
          const inner = sliceUntil(toks, i, tk.name)
          i += inner.length + 1
        }
        break
    }
  }
  return out
}

/** 单个 w:r → HTML（rPr 格式 + w:t/w:br/w:tab/w:drawing） */
async function runHtml(toks: XTok[], ctx: Ctx): Promise<string> {
  const rprOpen = toks.findIndex((t) => t.t === 'open' && t.name === 'w:rPr')
  let rpr = ''
  if (rprOpen >= 0) {
    const inner = sliceUntil(toks, rprOpen + 1, 'w:rPr')
    rpr = inner.map((t) => (t.t === 'text' ? '' : serializeToken(t))).join('')
  }
  const f = fmtOf(rpr)
  let content = ''
  let i = 0
  while (i < toks.length) {
    const tk = toks[i++]
    if (tk.t === 'text') {
      // rPr 外的裸文本（少见）按原样保留
      content += escHtml(unescXml(tk.v))
      continue
    }
    if (tk.t === 'close') continue
    if (tk.name === 'w:t') {
      const inner = sliceUntil(toks, i, 'w:t')
      i += inner.length + 1
      content += inner
        .map((t) => (t.t === 'text' ? escHtml(unescXml(t.v)) : ''))
        .join('')
      continue
    }
    if (tk.name === 'w:tab') {
      content += ' '
      continue
    }
    if (tk.name === 'w:br') {
      content += '<br>'
      continue
    }
    if (tk.name === 'w:drawing') {
      const inner = sliceUntil(toks, i, 'w:drawing')
      i += inner.length + 1
      content += await drawingImg(inner, ctx)
      continue
    }
    if (tk.name === 'w:pict') {
      const inner = sliceUntil(toks, i, 'w:pict')
      i += inner.length + 1
      content += await pictImg(inner, ctx)
      continue
    }
    if (tk.name === 'mc:AlternateContent') {
      const inner = sliceUntil(toks, i, 'mc:AlternateContent')
      i += inner.length + 1
      const choice = findBranch(inner, 'mc:Choice')
      if (choice) content += await inlineHtml(choice, ctx)
      continue
    }
    if (tk.t === 'open') {
      const inner = sliceUntil(toks, i, tk.name)
      i += inner.length + 1
    }
  }
  return wrapFmt(f, content)
}

function serializeToken(tk: XTok): string {
  if (tk.t === 'text') return tk.v
  if (tk.t === 'close') return `</${tk.name}>`
  const attrs = Object.entries(tk.attrs)
    .map(([k, v]) => ` ${k}="${v}"`)
    .join('')
  return tk.t === 'self' ? `<${tk.name}${attrs}/>` : `<${tk.name}${attrs}>`
}

/** sdt 内部取 sdtContent（找不到就返回原 token——内容兜底不丢） */
function takeSdtContent(inner: XTok[]): XTok[] {
  const openIdx = inner.findIndex((t) => (t.t === 'open' || t.t === 'self') && t.name === 'w:sdtContent')
  if (openIdx < 0) return inner
  return sliceUntil(inner, openIdx + 1, 'w:sdtContent')
}

/** 在 token 区间里找某分支（如 mc:Choice）的内容 */
function findBranch(toks: XTok[], name: string): XTok[] | null {
  const idx = toks.findIndex((t) => t.t === 'open' && t.name === name)
  if (idx < 0) return null
  return sliceUntil(toks, idx + 1, name)
}

/**
 * w:drawing 内容 → <img>（a:blip r:embed → rels → media → data: URI）。
 *
 * 两种形态：
 *   wp:inline            嵌入型 → 普通 img
 *   wp:anchor + wrapXXX  浮动环绕 → img 带 data-wrap / data-hside + float 样式，
 *                        写回时 htmlToOoxml 据此重建 anchor（往返闭环）
 *
 * 降级铁律（与 htmlToOoxml 对齐）：rId 找不到 rel、media 丢失、解码失败，
 * 一律输出可见占位 —— 不再静默返回空串（旧版用户看到的就是图片凭空消失）。
 */
async function drawingImg(toks: XTok[], ctx: Ctx): Promise<string> {
  const flat = toks.map(serializeToken).join('')
  const rId = /r:embed="([^"]+)"/.exec(flat)?.[1] ?? /r:link="([^"]+)"/.exec(flat)?.[1]
  const hit = rId ? await mediaDataUri(ctx, rId) : null
  // wp:extent cx/cy（EMU）→ px
  const cx = Number(/cx="(\d+)"/.exec(flat)?.[1] ?? 0)
  const cy = Number(/cy="(\d+)"/.exec(flat)?.[1] ?? 0)
  const px = (emu: number) => Math.round(emu / 9525)
  const w = px(cx) || hit?.w || 0
  const h = px(cy) || hit?.h || 0
  if (!hit) {
    // 图形还在但字节取不到：可见占位（尺寸尽量保留，方便用户定位重贴）
    const dim = w ? ` width="${Math.min(w, 540)}" height="${Math.min(h || Math.round(w * 0.6), 400)}"` : ''
    return `<img src="${IMG_MISSING}" alt="图片（原件缺失）" title="图片资源在文档中缺失或不受支持"${dim} data-img-missing="1">`
  }
  const dim = w ? ` width="${w}" height="${h || Math.round(w * 0.6)}"` : ''
  // 浮动环绕回读：wrapSquare → float:left/right；wrapTopAndBottom → 独占一行
  const isAnchor = /<wp:anchor\b/.test(flat)
  if (isAnchor) {
    const topBottom = /<wp:wrapTopAndBottom\b/.test(flat)
    const square = /<wp:wrapSquare\b/.test(flat)
    if (topBottom) {
      return `<img src="${hit.uri}"${dim} alt="" data-wrap="topBottom" style="display:block;margin:6px auto">`
    }
    if (square) {
      // positionH 的 align 决定图靠左/右（与写回侧 drawingAnchorXml 对应）
      const alignM = /<wp:positionH[^>]*>[\s\S]*?<wp:align>([a-z]+)<\/wp:align>/i.exec(flat)
      const hside = alignM?.[1] === 'right' ? 'right' : 'left'
      const side = hside === 'right' ? 'right' : 'left'
      return `<img src="${hit.uri}"${dim} alt="" data-wrap="square" data-hside="${side}" style="float:${side};margin:4px 8px">`
    }
    // 紧密型(wrapTight)等未支持形态按四周型兜底 —— 有环绕总比当 inline 强
    return `<img src="${hit.uri}"${dim} alt="" data-wrap="square" data-hside="left" style="float:left;margin:4px 8px">`
  }
  return `<img src="${hit.uri}"${dim} alt="">`
}

/** media 缺失占位图：1×1 透明 PNG 外框由 CSS 画（编辑面 .wrp-edit img[data-img-missing]） */
const IMG_MISSING =
  'data:image/gif;base64,R0lGODlhAQABAIAAAP///wAAACwAAAAAAQABAAACAkQBADs='

/** w:pict（VML）→ <img>（v:imagedata r:id；v:shape style 里尽量读出宽高） */
async function pictImg(toks: XTok[], ctx: Ctx): Promise<string> {
  const flat = toks.map(serializeToken).join('')
  const rId = /r:id="([^"]+)"/.exec(flat)?.[1]
  const hit = rId ? await mediaDataUri(ctx, rId) : null
  if (!hit) {
    return `<img src="${IMG_MISSING}" alt="图片（原件缺失）" title="图片资源在文档中缺失或不受支持" data-img-missing="1">`
  }
  // VML 的 style="width:213pt;height:160pt" 或 v:shape 的 width/height 属性
  const stM = /style="([^"]*)"/.exec(flat)?.[1] ?? ''
  const wPt = /width\s*:\s*([\d.]+)\s*(pt|px)?/i.exec(stM)
  const hPt = /height\s*:\s*([\d.]+)\s*(pt|px)?/i.exec(stM)
  const toPx = (v: string, u: string | undefined) => Math.round(Number(v) * (u === 'px' || !u ? 1 : 96 / 72))
  const w = wPt ? toPx(wPt[1], wPt[2]) : hit.w
  const h = hPt ? toPx(hPt[1], hPt[2]) : hit.h
  const dim = w ? ` width="${w}" height="${h || Math.round(w * 0.6)}"` : ''
  return `<img src="${hit.uri}"${dim} alt="">`
}

/* ---------- 块级：body tokens → 块序列 ---------- */

async function blocksOf(toks: XTok[], ctx: Ctx): Promise<HtmlBlock[]> {
  const blocks: HtmlBlock[] = []
  let i = 0
  while (i < toks.length) {
    const tk = toks[i++]
    if (tk.t === 'text') {
      const v = tk.v.trim()
      if (v) blocks.push({ ...EMPTY_PARA, runs: escHtml(unescXml(v)) })
      continue
    }
    if (tk.t === 'close') continue
    switch (tk.name) {
      case 'w:p': {
        const inner = sliceUntil(toks, i, 'w:p')
        i += inner.length + 1
        blocks.push(await paraOf(inner, ctx))
        break
      }
      case 'w:tbl': {
        const inner = sliceUntil(toks, i, 'w:tbl')
        i += inner.length + 1
        blocks.push(await tableOf(inner, ctx))
        break
      }
      case 'w:sdt':
      case 'w:smartTag':
      case 'w:customXml': {
        const inner = sliceUntil(toks, i, tk.name)
        i += inner.length + 1
        const content = tk.name === 'w:sdt' ? takeSdtContent(inner) : inner
        blocks.push(...(await blocksOf(content, ctx)))
        break
      }
      default:
        // sectPr / bookmark / proofErr 等块级杂项：整区间跳过
        if (tk.t === 'open') {
          const inner = sliceUntil(toks, i, tk.name)
          i += inner.length + 1
        }
        break
    }
  }
  return blocks
}

async function paraOf(toks: XTok[], ctx: Ctx): Promise<HtmlPara> {
  // pPr 在段落最前；先剥出来拿样式信息
  const pprOpen = toks.findIndex((t) => t.t === 'open' && t.name === 'w:pPr')
  let pprFlat = ''
  let restFrom = 0
  if (pprOpen >= 0) {
    const pprInner = sliceUntil(toks, pprOpen + 1, 'w:pPr')
    pprFlat = pprInner.map(serializeToken).join('')
    restFrom = pprOpen + pprInner.length + 2
  }
  const styleId = /<w:pStyle\s+w:val="([^"]+)"/.exec(pprFlat)?.[1]
  const outlineLvlM = /<w:outlineLvl\s+w:val="(\d+)"/.exec(pprFlat)
  const jc = /<w:jc\s+w:val="([^"]+)"/.exec(pprFlat)?.[1]
  const numId = /<w:numPr>[\s\S]*?<w:numId\s+w:val="([^"]+)"/.exec(pprFlat)?.[1]
  const ilvl = Number(/<w:numPr>[\s\S]*?<w:ilvl\s+w:val="(\d+)"/.exec(pprFlat)?.[1] ?? 0)

  const kind = numId ? (ctx.numFmt.get(numId) === 'decimal' ? 'ol' : 'ul') : null
  const lineM = /<w:spacing\b[^>]*w:line="(\d+)"/.exec(pprFlat)
  const lineRule = /<w:spacing\b[^>]*w:lineRule="([^"]+)"/.exec(pprFlat)?.[1]
  const indM = /<w:ind\b[^>]*w:firstLine="(\d+)"/.exec(pprFlat)
  const runs = await inlineHtml(toks.slice(restFrom), ctx)
  return {
    type: 'para',
    heading: headingOf(ctx, styleId, outlineLvlM ? Number(outlineLvlM[1]) : null),
    align: jc && jc !== 'left' ? (jc === 'both' ? 'justify' : jc) : null,
    list: kind ? { kind, level: Math.min(5, ilvl) } : null,
    runs,
    // lineRule=auto 时 w:line 是 1/240 单位（240=单倍）；exact/atLeast 不用（是固定行高）
    lineHeight: lineM && (!lineRule || lineRule === 'auto') ? Number(lineM[1]) / 240 : null,
    textIndent: indM ? Number(indM[1]) / 240 : null,
  }
}

async function tableOf(toks: XTok[], ctx: Ctx): Promise<HtmlTable> {
  const rows: HtmlTable['rows'] = []
  // w:tblGrid 列宽（px，1px@96dpi = 15 twips）：缺 tblGrid 时 cols 为空 → 不产 colwidth（现状行为）
  const gridTok = toks.findIndex((t) => t.t === 'open' && t.name === 'w:tblGrid')
  let cols: number[] | undefined
  if (gridTok >= 0) {
    const grid = sliceUntil(toks, gridTok + 1, 'w:tblGrid')
    cols = grid
      .map((t) => (t.t === 'self' && t.name === 'w:gridCol' ? Number(t.attrs['w:w'] ?? 0) : null))
      .filter((n): n is number => n != null)
      .map((n) => Math.round((n / 15) * 10) / 10)
    if (!cols.length) cols = undefined
  }
  let colCursor = 0 // 当前格对应 grid 第几列（span 累进）
  let i = 0
  while (i < toks.length) {
    const tk = toks[i++]
    if (tk.t !== 'open' || tk.name !== 'w:tr') continue
    const trInner = sliceUntil(toks, i, 'w:tr')
    i += trInner.length + 1
    const cells: HtmlTable['rows'][number] = []
    let j = 0
    while (j < trInner.length) {
      const c = trInner[j++]
      if (c.t !== 'open' || c.name !== 'w:tc') continue
      const tcInner = sliceUntil(trInner, j, 'w:tc')
      j += tcInner.length + 1
      const tcprOpen = tcInner.findIndex((t) => t.t === 'open' && t.name === 'w:tcPr')
      let span = 1
      if (tcprOpen >= 0) {
        const tcpr = sliceUntil(tcInner, tcprOpen + 1, 'w:tcPr').map(serializeToken).join('')
        span = Math.max(1, Number(/<w:gridSpan\s+w:val="(\d+)"/.exec(tcpr)?.[1] ?? 1))
      }
      // 单元格内容从 tcPr 之后起（tcPr 内的 w:tcW 等不会产出文本，但显式跳过更干净）
      let contentFrom = 0
      if (tcprOpen >= 0) {
        const tcprInner = sliceUntil(tcInner, tcprOpen + 1, 'w:tcPr')
        contentFrom = tcprOpen + tcprInner.length + 2
      }
      // 该格覆盖的 grid 列宽（gridSpan>1 时多段，Tiptap colwidth 属性逗号分隔）
      let colwidth: number[] | undefined
      if (cols) {
        const seg = cols.slice(colCursor, colCursor + span)
        colwidth = seg.length === span ? seg : undefined // grid 不够长（坏表）就放弃，别产错宽
        colCursor += span
      }
      cells.push({ span, colwidth, blocks: await blocksOf(tcInner.slice(contentFrom), ctx) })
    }
    rows.push(cells)
  }
  return { type: 'table', rows, cols }
}

/* ---------- 发射：块序列 → HTML（列表合并 + 标题 + 对齐） ---------- */

function blocksToHtml(blocks: HtmlBlock[]): string {
  let out = ''
  // 列表合成栈：OOXML 的列表是平铺段落（numPr 标记），HTML 要合成嵌套 ul/ol。
  // 栈中每个元素对应一层「<kind><li>…」（li 未闭合）；关一层 = </li></kind>。
  type Open = { kind: 'ul' | 'ol'; level: number }
  const stack: Open[] = []
  const closeOne = () => {
    const top = stack.pop()!
    out += `</li></${top.kind}>`
  }
  const closeAll = () => {
    while (stack.length) closeOne()
  }
  const openList = (l: { kind: 'ul' | 'ol'; level: number }, runs: string) => {
    out += `<${l.kind}><li>${runs}`
    stack.push({ ...l })
  }

  for (const b of blocks) {
    if (b.type === 'table') {
      closeAll()
      out += tableHtml(b)
      continue
    }
    if (!b.list) {
      closeAll()
      const tag = b.heading ?? 'p'
      const css: string[] = []
      if (b.align) css.push(`text-align:${b.align}`)
      if (b.lineHeight) css.push(`line-height:${b.lineHeight}`)
      if (b.textIndent) css.push(`text-indent:${b.textIndent}em`)
      const attr = css.length ? ` style="${css.join(';')}"` : ''
      out += `<${tag}${attr}>${b.runs || '<br>'}</${tag}>`
      continue
    }
    const cur = stack[stack.length - 1]
    if (!cur || b.list.level > cur.level) {
      openList(b.list, b.runs)
      continue
    }
    // 关到能放下的层（level 更浅，或 kind 不同的同层）
    while (stack.length && (stack[stack.length - 1].level > b.list.level || stack[stack.length - 1].kind !== b.list.kind)) {
      closeOne()
    }
    const now = stack[stack.length - 1]
    if (now && now.level === b.list.level) out += `</li><li>${b.runs}`
    else openList(b.list, b.runs)
  }
  closeAll()
  return out
}

function tableHtml(t: HtmlTable): string {
  const rows = t.rows
    .map((cells) => {
      const tds = cells
        .map((c) => {
          const inner = blocksToHtml(c.blocks) || '<p><br></p>'
          // colwidth 是 @tiptap/extension-table-cell 的官方约定属性（parseHTML 按逗号
          // 分裂成 cell attrs.colwidth）→ TableView 据此生成 colgroup 锁定列宽
          const cw = c.colwidth?.length ? ` colwidth="${c.colwidth.map((n) => Math.round(n)).join(',')}"` : ''
          return `<td${cw}${c.span > 1 ? ` colspan="${c.span}"` : ''}>${inner}</td>`
        })
        .join('')
      return `<tr>${tds}</tr>`
    })
    .join('')
  return `<table><tbody>${rows}</tbody></table>`
}

/* ---------- 入口 ---------- */

/**
 * docx 字节 → body HTML + sectPr + 根属性。
 * 编辑面渲染用 html；写回时用 sectPr/rootAttrs 重组 document.xml（见 WordReportPanel）。
 */
export async function docxBodyToHtml(input: Uint8Array): Promise<DocxBodyHtml> {
  const zip = await JSZip.loadAsync(input)
  return docxBodyToHtmlFromZip(zip)
}

/** 已有 zip 实例时避免重复解包（编辑器刷新预览时用） */
export async function docxBodyToHtmlFromZip(zip: JSZip): Promise<DocxBodyHtml> {
  const docFile = zip.file('word/document.xml')
  if (!docFile) throw new Error('不是有效的 docx（缺 word/document.xml）')
  const xml = await docFile.async('string')

  const rootM = /<w:document([^>]*)>/.exec(xml)
  const rootAttrs = rootM ? rootM[1] : ''
  const bodyM = /<w:body[^>]*>([\s\S]*)<\/w:body>/.exec(xml)
  if (!bodyM) throw new Error('document.xml 缺少 w:body')

  // body 级 sectPr：正常在 </w:body> 正前方；兜底取最后一个（避免把段内分节符误当 body 级）
  let sectPr = ''
  let inner = bodyM[1]
  const tail = /<w:sectPr\b[\s\S]*?<\/w:sectPr>\s*$/.exec(inner)
  if (tail) {
    sectPr = tail[0]
    inner = inner.slice(0, tail.index)
  } else {
    const all = [...inner.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)]
    if (all.length) {
      const last = all[all.length - 1]
      sectPr = last[0]
      inner = inner.slice(0, last.index) + inner.slice(last.index + last[0].length)
    }
  }

  const ctx = await buildCtx(zip)
  const blocks = await blocksOf(tokenize(inner), ctx)
  return { html: blocksToHtml(blocks), sectPr, rootAttrs }
}
