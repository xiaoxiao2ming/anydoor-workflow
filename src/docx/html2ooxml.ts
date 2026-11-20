/**
 * 极简 HTML → OOXML 转换器（零依赖）
 *
 * 定位：把富文本编辑器产出的 HTML 转成可以塞进 <w:sdtContent> 的 OOXML 片段。
 * 编辑内核 = Tiptap（ProseMirror），其 HTML 输出形态已纳入解析：
 *   <mark> 高亮 · style="background-color" · style="color" · style="font-size"
 *   style="font-family" · <p style="text-align"> · <p style="line-height">
 *
 * 支持：p/div/h1-h6/blockquote/pre、strong/b、em/i、u、s/del、sub/sup、
 *       span(color/font-size/font-family/background-color)、mark(高亮)、
 *       a(降级为蓝色下划线文本)、br、ul/ol/li、table/tr/td
 * 不做：浮动布局、CSS 盒模型
 *
 * 图片：支持 data: URI（富文本编辑器粘贴出来的就是这种形态），
 * 通过 img 回调交给调用方落进 word/media 并分配 rId；
 * data-wrap 属性（Tiptap 图片环绕）→ wp:anchor + wrapSquare/wrapTopAndBottom；
 * 外链图与注入失败一律降级为可见文字，绝不静默丢弃。
 */

import { decodeDataUri, displaySize, drawingXml, drawingAnchorXml, type ImageWrap } from './image'

const VOID = new Set(['br', 'img', 'hr', 'input', 'meta', 'link', 'wbr', 'col', 'source'])
const SKIP_CONTENT = new Set(['script', 'style', 'head', 'title'])
const HEADINGS: Record<string, string> = {
  h1: 'Heading1', h2: 'Heading2', h3: 'Heading3',
  h4: 'Heading4', h5: 'Heading5', h6: 'Heading6',
}
const INLINE_STYLE: Record<string, 'b' | 'i' | 'u' | 's' | 'sub' | 'sup'> = {
  b: 'b', strong: 'b',
  i: 'i', em: 'i',
  u: 'u', ins: 'u',
  s: 's', del: 's', strike: 's',
  sub: 'sub', sup: 'sup',
}

export const esc = (s: unknown): string =>
  String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

/* ---------- 词法 ---------- */

type Token =
  | { t: 'text'; v: string }
  | { t: 'open'; name: string; attrs: Record<string, string> }
  | { t: 'self'; name: string; attrs: Record<string, string> }
  | { t: 'close'; name: string }

function parseAttrs(raw: string): Record<string, string> {
  const a: Record<string, string> = {}
  const re = /([a-zA-Z_:][-\w:.]*)\s*=\s*("([^"]*)"|'([^']*)'|([^\s"'>]+))/g
  let m: RegExpExecArray | null
  while ((m = re.exec(raw))) a[m[1].toLowerCase()] = (m[3] ?? m[4] ?? m[5] ?? '').trim()
  return a
}

function tokenize(html: string): Token[] {
  const out: Token[] = []
  let i = 0
  while (i < html.length) {
    const lt = html.indexOf('<', i)
    if (lt < 0) {
      if (i < html.length) out.push({ t: 'text', v: html.slice(i) })
      break
    }
    if (lt > i) out.push({ t: 'text', v: html.slice(i, lt) })
    if (html.startsWith('<!--', lt)) {
      const e = html.indexOf('-->', lt)
      i = e < 0 ? html.length : e + 3
      continue
    }
    const gt = html.indexOf('>', lt)
    if (gt < 0) {
      out.push({ t: 'text', v: html.slice(lt) })
      break
    }
    const raw = html.slice(lt, gt + 1)
    const close = raw[1] === '/'
    const nm = raw.slice(close ? 2 : 1).match(/^([a-zA-Z][\w-]*)/)
    const name = nm ? nm[1].toLowerCase() : ''
    if (close) out.push({ t: 'close', name })
    else if (raw.endsWith('/>') || VOID.has(name)) out.push({ t: 'self', name, attrs: parseAttrs(raw) })
    else out.push({ t: 'open', name, attrs: parseAttrs(raw) })
    i = gt + 1
  }
  return out
}

/** 取从 i 起、直到与 tagName 配对闭合之间的 token（不含闭合标签本身） */
function sliceUntil(toks: Token[], i: number, tagName: string | string[]): Token[] {
  const names = Array.isArray(tagName) ? tagName : [tagName]
  let depth = 1
  const out: Token[] = []
  while (i < toks.length) {
    const tk = toks[i++]
    if (tk.t === 'open' && names.includes(tk.name)) depth++
    else if (tk.t === 'close' && names.includes(tk.name) && --depth === 0) break
    else out.push(tk)
  }
  return out
}

/* ---------- 语法：tokens → 块列表 ---------- */

interface Run {
  t: string
  /** 图片：整段 w:drawing XML，有它时忽略 t */
  draw?: string
  b?: number
  i?: number
  u?: number
  s?: number
  sub?: number
  sup?: number
  color?: string | null
  sz?: number | null
  /** 突出显示：命名色（w:highlight）或任意 RGB hex（w:shd 字符底纹） */
  hl?: string | null
  /** 字体族（w:rFonts） */
  ff?: string | null
}
interface ParaBlock {
  /** 判别字段：让 Block 成为可判别联合，避免运行时 typeof 猜测 */
  type: 'para'
  runs: Run[]
  heading: string | null
  list: { kind: 'ul' | 'ol'; level: number; idx: number } | null
  indent: number
  pre: boolean
  /** 段落对齐（Tiptap TextAlign → <p style="text-align:center">） */
  align?: string | null
  /** 行距（Tiptap 行距扩展 → <p style="line-height:1.5">），1.5 表示 150% */
  lineHeight?: number | null
  /** 首行缩进字符数（<p style="text-indent:2em">） */
  textIndent?: number | null
}
interface TableBlock {
  type: 'table'
  rows: TableCellBlock[][]
}
interface TableCellBlock {
  head: boolean
  blocks: Block[]
}
type Block = ParaBlock | TableBlock

interface Style {
  b: number
  i: number
  u: number
  s: number
  sub: number
  sup: number
  color: string | null
  sz: number | null
  hl: string | null
  ff: string | null
}
const newStyle = (): Style => ({ b: 0, i: 0, u: 0, s: 0, sub: 0, sup: 0, color: null, sz: null, hl: null, ff: null })

/**
 * 富文本里的图片落进 docx 的通道。
 * 返回该图片的 rId（调用方写进 rels 并分配），返回 null 表示放弃——转换器随即降级为文字，
 * 绝不静默丢弃。
 */
export type ImageSink = (bytes: Uint8Array, ext: string) => string | null

function parseTokens(toks: Token[], img?: ImageSink): Block[] {
  const blocks: Block[] = []
  const stack: Style[] = []
  const st = newStyle()
  const listStack: { kind: 'ul' | 'ol'; idx: number }[] = []
  const ctx = { heading: null as string | null, pre: false }
  let para: ParaBlock | null = null

  const curList = () => {
    if (!listStack.length) return null
    const top = listStack[listStack.length - 1]
    return { kind: top.kind, level: listStack.length - 1, idx: top.idx }
  }
  const ensure = (): ParaBlock => {
    if (!para) para = { type: 'para', runs: [], heading: ctx.heading, list: curList(), indent: 0, pre: ctx.pre }
    return para
  }
  /**
   * 块级标签（p/div/h1-6）的 style → 段落属性。
   * Tiptap 的把对齐/行距/缩进放在块标签的 style 上（不是内联 span），
   * 晚于 ensure() 调用也会写进已在建的段落。
   */
  const applyBlockStyle = (tk: Extract<Token, { t: 'open' | 'self' }>) => {
    const style = tk.attrs.style || ''
    if (!style) return
    const p = ensure()
    const al = /text-align\s*:\s*(left|center|right|justify)/i.exec(style)
    if (al) p.align = al[1].toLowerCase()
    const lh = /line-height\s*:\s*([\d.]+)(pt|px|em|%)?/i.exec(style)
    if (lh) {
      const v = parseFloat(lh[1])
      const u = (lh[2] || '').toLowerCase()
      // 存成「倍数」：无单位/em = 倍数；% 除 100；pt/px 换算（12pt 基准）
      p.lineHeight = u === '%' ? v / 100 : u === 'pt' ? v / 12 : u === 'px' ? v / 16 : v
    }
    const ti = /text-indent\s*:\s*([\d.]+)(em|pt|px|cm)?/i.exec(style)
    if (ti) {
      const v = parseFloat(ti[1])
      const u = (ti[2] || 'em').toLowerCase()
      // OOXML 首行缩进用「字符数」：em 直接就是字符数，其余换算成 pt 再除以字号(12pt)
      p.textIndent = u === 'em' ? v : u === 'cm' ? (v * 28.35) / 12 : v / 12
    }
  }
  const flush = () => {
    if (para) {
      blocks.push(para)
      para = null
    }
  }
  const emit = (text: string) => {
    if (text === '') return
    ensure().runs.push({ t: text, ...st })
  }

  /**
   * 图片降级铁律：拿不到图也必须留下可见痕迹。
   * 之前 <p><img></p> 会因为从不调用 ensure() 而让整段凭空消失，
   * 用户导出后才发现少了一段——属于静默数据丢失。
   *
   * 环绕属性（编辑面浮层写入）：
   *   data-wrap="inline"                嵌入型（默认，随文字流）
   *   data-wrap="square" data-hside     四周型，hside=left/right
   *   data-wrap="topBottom"             上下型
   */
  const emitImage = (attrs: Record<string, string>, sink?: ImageSink) => {
    const alt = (attrs.alt || '').trim()
    const src = attrs.src || ''
    const decoded = sink ? decodeDataUri(src) : null
    const rId = decoded && sink ? sink(decoded.bytes, decoded.ext) : null

    if (decoded && rId) {
      const { cx, cy } = displaySize(decoded.bytes, attrs)
      const wrap = attrs['data-wrap']
      if (wrap === 'square' || wrap === 'topBottom') {
        const w: ImageWrap = wrap === 'topBottom' ? 'topAndBottom' : 'square'
        const hside = attrs['data-hside'] === 'right' ? 'right' : 'left'
        ensure().runs.push({ t: '', draw: drawingAnchorXml(rId, cx, cy, w, hside, alt || '图片') })
      } else {
        ensure().runs.push({ t: '', draw: drawingXml(rId, cx, cy, alt || '图片') })
      }
      return
    }
    // 非 data: URI（外链图）或注入失败 → 转成文字，至少读者知道这里有张图
    const label = alt || (/^(https?:|file:)/i.test(src) ? src : '未支持的图片')
    emit(`[图片：${label}]`)
  }
  const pushStyle = (tk: Extract<Token, { t: 'open' }>) => {
    stack.push({ ...st })
    const a = tk.attrs
    const key = INLINE_STYLE[tk.name]
    if (key) {
      st[key] = 1
      return
    }
    if (tk.name === 'a') {
      st.u = 1
      if (!st.color) st.color = '0563C1'
      return
    }
    if (tk.name === 'span' || tk.name === 'font') {
      const css = [a.style || '', a.color ? `color:${a.color}` : '', a.size ? `font-size:${a.size}pt` : ''].join(';')
      // 两个颜色正则都用「行首或分号后」断言：否则 background-color:yellow 里的
      // `color:` 会被文字颜色正则误吃（只有高亮没有字色的 run 会变黄字）
      const c = /(?:^|;)\s*color\s*:\s*([^;]+)/i.exec(css)
      if (c) st.color = normColor(c[1].trim())
      const sz = /font-size\s*:\s*([\d.]+)\s*(pt|px|rem|em)?/i.exec(css)
      if (sz) {
        const unit = (sz[2] || 'pt').toLowerCase()
        const pt = unit === 'px' ? +sz[1] * 0.75 : unit === 'rem' || unit === 'em' ? +sz[1] * 12 : +sz[1]
        st.sz = Math.round(pt * 2)
      }
      // 字体族（Tiptap FontFamily 扩展 → style="font-family:..."）
      const ff = /font-family\s*:\s*([^;]+)/i.exec(css)
      if (ff) st.ff = ff[1].trim().replace(/^['"]|['"]$/g, '')
      // 突出显示（Tiptap Highlight 的 <mark> 或 execCommand hiliteColor 的 background-color）。
      // 之前这层信息在导出时整层丢弃 —— 用户标了黄导出变白，属于可见格式丢失。
      const bg = /(?:^|;)\s*background-color\s*:\s*([^;]+)/i.exec(css)
      if (bg) {
        const rgb = normColor(bg[1].trim())
        if (rgb) st.hl = rgb
      }
    }
    // Tiptap 的 Highlight 扩展产出 <mark>（可带 data-color / style background-color）
    if (tk.name === 'mark') {
      const css = [a.style || '', a['data-color'] ? `background-color:${a['data-color']}` : ''].join(';')
      const bg = /(?:^|;)\s*background-color\s*:\s*([^;]+)/i.exec(css)
      const rgb = bg ? normColor(bg[1].trim()) : null
      st.hl = rgb ?? 'FFFF00' // <mark> 无显式色 → 默认黄（HTML 语义默认即黄底）
    }
  }
  const popStyle = () => {
    const prev = stack.pop()
    if (prev) Object.assign(st, prev)
  }

  let i = 0
  while (i < toks.length) {
    const tk = toks[i++]

    if (tk.t === 'text') {
      emit(ctx.pre ? tk.v : tk.v.replace(/\s+/g, ' '))
      continue
    }

    if (tk.t === 'self') {
      if (tk.name === 'br') ensure().runs.push({ t: '\n', ...st })
      else if (tk.name === 'img') emitImage(tk.attrs, img)
      continue
    }

    if (tk.t === 'open') {
      if (SKIP_CONTENT.has(tk.name)) {
        i += sliceUntil(toks, i, tk.name).length + 1
        continue
      }
      switch (tk.name) {
        case 'ul':
          flush()
          listStack.push({ kind: 'ul', idx: 0 })
          break
        case 'ol':
          flush()
          listStack.push({ kind: 'ol', idx: 0 })
          break
        case 'li':
          flush()
          if (listStack.length) listStack[listStack.length - 1].idx++
          break
        case 'p':
        case 'div':
        case 'blockquote':
          flush()
          applyBlockStyle(tk)
          break
        case 'pre':
          flush()
          ctx.pre = true
          break
        case 'table': {
          flush()
          const inner = sliceUntil(toks, i, 'table')
          i += inner.length + 1 // 内部 token + 配对的 </table>
          blocks.push({ type: 'table', rows: parseTableTokens(inner, img) })
          break
        }
        case 'h1': case 'h2': case 'h3': case 'h4': case 'h5': case 'h6':
          flush()
          ctx.heading = HEADINGS[tk.name]
          applyBlockStyle(tk)
          break
      default:
        if (INLINE_STYLE[tk.name] || tk.name === 'a' || tk.name === 'span' || tk.name === 'font' || tk.name === 'mark') pushStyle(tk)
      }
      continue
    }

    // close
    switch (tk.name) {
      case 'ul':
      case 'ol':
        listStack.pop()
        flush()
        break
      case 'li':
        flush()
        break
      case 'pre':
        flush()
        ctx.pre = false
        break
      case 'p':
      case 'div':
      case 'blockquote':
        flush()
        break
      case 'h1': case 'h2': case 'h3': case 'h4': case 'h5': case 'h6':
        flush()
        ctx.heading = null
        break
      default:
        if (stack.length) popStyle()
    }
  }
  flush()
  return blocks
}

function parseTableTokens(inner: Token[], img?: ImageSink): TableCellBlock[][] {
  const rows: TableCellBlock[][] = []
  let i = 0
  while (i < inner.length) {
    const tk = inner[i]
    if (tk.t === 'open' && tk.name === 'tr') {
      const cellsTokens = sliceUntil(inner, i + 1, 'tr')
      rows.push(parseRowTokens(cellsTokens, img))
      i += cellsTokens.length + 2
    } else i++
  }
  return rows
}

function parseRowTokens(rowTokens: Token[], img?: ImageSink): TableCellBlock[] {
  const cells: TableCellBlock[] = []
  let i = 0
  while (i < rowTokens.length) {
    const tk = rowTokens[i]
    if (tk.t === 'open' && (tk.name === 'td' || tk.name === 'th')) {
      const head = tk.name === 'th'
      const cellTokens = sliceUntil(rowTokens, i + 1, ['td', 'th'])
      const blocks = parseTokens(cellTokens, img)
      cells.push({
        head,
        blocks: blocks.length ? blocks : [{ type: 'para', runs: [], heading: null, list: null, indent: 0, pre: false }],
      })
      i += cellTokens.length + 2
    } else i++
  }
  return cells
}

function normColor(c: string | undefined): string | null {
  if (!c) return null
  c = c.trim()
  if (c.startsWith('#')) {
    let h = c.slice(1)
    if (h.length === 3) h = h.split('').map((x) => x + x).join('')
    return h.toUpperCase().padStart(6, '0').slice(0, 6)
  }
  const m = c.match(/rgba?\(([^)]+)\)/)
  if (m) {
    const [r, g, b] = m[1].split(',').map((x) => parseInt(x, 10) || 0)
    return [r, g, b].map((x) => (x || 0).toString(16).padStart(2, '0')).join('').toUpperCase()
  }
  const named: Record<string, string> = { red: 'FF0000', blue: '0000FF', green: '008000', black: '000000', gray: '808080' }
  return named[c.toLowerCase()] || null
}

/* ---------- 发射：块 → OOXML ---------- */

/** w:highlight 官方认可的 16 命名色（OOXML 规范 ST_HighlightColor） */
const HL_NAMED: Record<string, string> = {
  FFFF00: 'yellow',
  '00FF00': 'green',
  '00FFFF': 'cyan',
  FF00FF: 'magenta',
  '0000FF': 'blue',
  FF0000: 'red',
  '00008B': 'darkBlue',
  '008B8B': 'darkCyan',
  '006400': 'darkGreen',
  '8B008B': 'darkMagenta',
  '8B0000': 'darkRed',
  '808000': 'darkYellow',
  FFFFFF: 'white',
  D3D3D3: 'lightGray',
  A9A9A9: 'darkGray',
  '000000': 'black',
}

/** 高亮色 → OOXML：命名色用 w:highlight（Word 突出显示面板原生色），任意 RGB 用 w:shd 字符底纹（WPS/Word 都渲染） */
function highlightXml(rgb: string): string {
  const named = HL_NAMED[rgb.toUpperCase()]
  if (named) return `<w:highlight w:val="${named}"/>`
  return `<w:shd w:val="clear" w:color="auto" w:fill="${rgb}"/>`
}

function runsXml(runs: Run[]): string {
  return runs
    .filter((r) => r.t !== '' || r.draw)
    .map((r) => {
      const pr: string[] = []
      if (r.b) pr.push('<w:b/>')
      if (r.i) pr.push('<w:i/>')
      if (r.u) pr.push('<w:u w:val="single"/>')
      if (r.s) pr.push('<w:strike/>')
      if (r.sub) pr.push('<w:vertAlign w:val="subscript"/>')
      if (r.sup) pr.push('<w:vertAlign w:val="superscript"/>')
      if (r.color) pr.push(`<w:color w:val="${r.color}"/>`)
      if (r.ff) {
        // 中文字体名要用 w:eastAsia，否则 Word 只对西文生效
        const f = esc(r.ff)
        pr.push(`<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}"/>`)
      }
      if (r.hl) pr.push(highlightXml(r.hl))
      if (r.sz) pr.push(`<w:sz w:val="${r.sz}"/><w:szCs w:val="${r.sz}"/>`)
      const rpr = pr.length ? `<w:rPr>${pr.join('')}</w:rPr>` : ''
      const body = r.draw ?? `<w:t xml:space="preserve">${esc(r.t)}</w:t>`
      return `<w:r>${rpr}${body}</w:r>`
    })
    .join('')
}

function paraXml(p: ParaBlock, num: { bullet: number; decimal: number } | null): string {
  const ppr: string[] = []
  let head = ''
  if (p.heading) ppr.push(`<w:pStyle w:val="${p.heading}"/>`)
  if (p.pre) ppr.push('<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/></w:rPr>')
  // 段落级排版（Tiptap TextAlign / 行距 / 首行缩进）
  if (p.align) {
    const jc = p.align === 'justify' ? 'both' : p.align
    ppr.push(`<w:jc w:val="${jc}"/>`)
  }
  if (p.lineHeight && p.lineHeight > 0) {
    // 240 = 单倍行距基准（240 twips 的 line）；auto 模式配 lineRule
    ppr.push(`<w:spacing w:line="${Math.round(p.lineHeight * 240)}" w:lineRule="auto"/>`)
  }
  if (p.textIndent && p.textIndent > 0) {
    // 首行缩进：firstLine 单位 twips，1 字符 ≈ 240 twips（12pt 字号）
    ppr.push(`<w:ind w:firstLine="${Math.round(p.textIndent * 240)}"/>`)
  }
  if (p.list) {
    if (num) {
      ppr.push(
        `<w:numPr><w:ilvl w:val="${p.list.level}"/><w:numId w:val="${p.list.kind === 'ol' ? num.decimal : num.bullet}"/></w:numPr>`
      )
      ppr.push(`<w:ind w:left="${720 * (p.list.level + 1)}" w:hanging="360"/>`)
    } else {
      // 模板无 numbering 定义 → 降级为字面符号，保证信息不丢
      const mark = p.list.kind === 'ol' ? `${p.list.idx}.` : '•'
      head = `<w:r><w:t xml:space="preserve">${esc('　'.repeat(p.list.level) + mark + ' ')}</w:t></w:r>`
    }
  }
  const pr = ppr.length ? `<w:pPr>${ppr.join('')}</w:pPr>` : ''
  return `<w:p>${pr}${head}${runsXml(p.runs)}</w:p>`
}

/**
 * 表格 → OOXML。
 *
 * ★ 跨页行为（2026-09-28 补齐，用户：「根本不能让我做到跨页。甚至表头重复出现」）：
 *
 *  1. `<w:tblHeader/>` —— 写在首行的 `<w:trPr>` 里，声明它为「标题行」。
 *     Word/WPS 打开后，表格跨页时**每页顶部自动重复该行**。
 *     ★ 必须放进 `<w:trPr>`，且 `<w:trPr>` 必须是 `<w:tr>` 的第一个子元素
 *       （OOXML 的 tr 子元素顺序是强制的：trPr → tc…）。
 *     只对**首行**加（用户拍板「首行自动」）：靠 `rowIdx === 0` 判定，
 *     不依赖 `head`（head 来自 th，而用户可能用 td 做表头、或首行不全是 th）。
 *     ★ 表格只有 1 行时不加（没有意义，且个别 Word 版本对单行 tblHeader 会多画一行）。
 *
 *  2. `<w:cantSplit/>` —— 同样写在 `<w:trPr>`，声明「本行不许被分页拆开」。
 *     Word 遇到放不下的行会**整行推到下一页**，正好对齐我们编辑视图的避让语义。
 *     对**所有行**加（用户要「做到跨页」，即整行移动而非拦腰截断）。
 *
 *  这两条是**真写进文件的**，与编辑视图那套「视觉分页」（纸 + 页隙遮断条）
 *  是两套东西 —— 编辑视图是模拟，这里才是 Word 里的真实排版。
 */
function tableXml(t: { rows: TableCellBlock[][] }, num: { bullet: number; decimal: number } | null): string {
  const cols = Math.max(1, ...t.rows.map((r) => r.length))
  const w = Math.floor(9000 / cols)
  const grid = Array.from({ length: cols }, () => `<w:gridCol w:w="${w}"/>`).join('')
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((k) => `<w:${k} w:val="single" w:sz="4" w:space="0" w:color="999999"/>`)
    .join('')
  const rows = t.rows
    .map((cells, rowIdx) => {
      const tds = cells
        .map((c) => {
          const boldPr = c.head ? '<w:pPr><w:rPr><w:b/></w:rPr></w:pPr>' : ''
          const inner = c.blocks.map((b) => (b.type === 'table' ? tableXml(b, num) : paraXml(b, num))).join('') || `<w:p>${boldPr}</w:p>`
          const shaded = c.head ? inner.replace(/<w:p>/g, `<w:p>${boldPr}`) : inner
          const shade = c.head ? '<w:shd w:val="clear" w:color="auto" w:fill="E8E8E8"/>' : ''
          // 不引用 TableGrid 样式：设计院的模板未必定义它，显式边框更可靠
          return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${shade}</w:tcPr>${shaded}</w:tc>`
        })
        .join('')
      // trPr 必须是 tr 的第一个子元素；cantSplit（不拆行）所有行都加
      const trPr: string[] = ['<w:cantSplit/>']
      // tblHeader（跨页重复表头）只加在首行，且至少 2 行才有意义
      if (rowIdx === 0 && t.rows.length > 1) trPr.push('<w:tblHeader/>')
      return `<w:tr><w:trPr>${trPr.join('')}</w:trPr>${tds}</w:tr>`
    })
    .join('')
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${borders}</w:tblBorders></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`
}

/**
 * @param html 富文本 HTML
 * @param num numbering.xml 中的 numId；为 null 时列表降级为字面符号
 * @returns OOXML 片段，可直接放进 <w:sdtContent> 或 <w:tc>
 */
export function htmlToOoxml(
  html: string,
  num: { bullet: number; decimal: number } | null = null,
  img?: ImageSink,
): string {
  if (!html || !String(html).trim()) return '<w:p/>'
  const blocks = parseTokens(tokenize(String(html)), img)
  if (!blocks.length) return '<w:p/>'
  return blocks.map((b) => (b.type === 'table' ? tableXml(b, num) : paraXml(b, num))).join('')
}
