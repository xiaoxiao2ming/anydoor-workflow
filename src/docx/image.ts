/**
 * docx 图片注入
 *
 * 富文本里的 <img> 要落进 Word，必须同时改动三处，缺一个 Word 就打不开或丢图：
 *   1. word/media/xxx.png        —— 图片字节本体
 *   2. word/_rels/document.xml.rels —— 声明 image 关系并分配 rId
 *   3. [Content_Types].xml        —— 声明扩展名对应的 MIME
 *
 * 绑定引擎原本只做 document.xml 的字符串级手术，不知道部件的存在。
 * 这里把「跨部件资源注入」收拢成一个 sink，bind() 只需要在开头建一次、结尾 flush 一次。
 */

import type JSZip from 'jszip'

const EMU_PER_PX = 9525

/* ---------- data: URI 解码 ---------- */

const EXT_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  tif: 'image/tiff',
  tiff: 'image/tiff',
}

/** base64 → 字节。浏览器与 Node 都能用，不依赖 Buffer / atob 之外的东西 */
function b64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64)
  const out = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
  return out
}

export interface DecodedImage {
  bytes: Uint8Array
  ext: string
}

/** 只认 data: URI —— 富文本编辑器粘贴出来的图片就是这种形态，同步可解 */
export function decodeDataUri(src: string): DecodedImage | null {
  const m = /^data:([a-z0-9.+-]+\/[a-z0-9.+-]+)?(;[^,]*)?,(.*)$/is.exec(String(src).trim())
  if (!m) return null
  const [, mime = '', meta = '', payload = ''] = m
  if (!/;base64/i.test(meta)) return null
  let bytes: Uint8Array
  try {
    bytes = b64ToBytes(payload.replace(/\s+/g, ''))
  } catch {
    return null
  }
  if (!bytes.length) return null
  let ext = mime.split('/')[1]?.toLowerCase() ?? ''
  if (ext === 'svg+xml') ext = 'svg'
  if (!ext || !(ext in EXT_MIME)) ext = sniffExt(bytes)
  return { bytes, ext }
}

/** data: URI 没写 MIME 或写了不认识的，按魔数猜 */
function sniffExt(b: Uint8Array): string {
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'png'
  if (b[0] === 0xff && b[1] === 0xd8) return 'jpg'
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'gif'
  if (b[0] === 0x42 && b[1] === 0x4d) return 'bmp'
  if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46) return 'webp'
  return 'png'
}

/* ---------- 同步读图片原始尺寸 ---------- */

/**
 * 从文件头读宽高（px）。读不出来返回 null，调用方用默认尺寸兜底。
 * 同步实现是为了让 htmlToOoxml 保持同步 —— 不为了拿尺寸把整条转换链改成 async。
 */
export function imageSize(bytes: Uint8Array): { w: number; h: number } | null {
  try {
    if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
      // PNG：IHDR 里 width/height 是大端 4 字节
      const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
      return { w: dv.getUint32(16), h: dv.getUint32(20) }
    }
    if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
      const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
      return { w: dv.getUint16(6, true), h: dv.getUint16(8, true) }
    }
    if (bytes[0] === 0xff && bytes[1] === 0xd8) {
      // JPEG：扫各段找 SOF0~SOF15
      let i = 2
      while (i < bytes.length - 9) {
        if (bytes[i] !== 0xff) {
          i++
          continue
        }
        const marker = bytes[i + 1]
        if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
          return { h: (bytes[i + 5] << 8) | bytes[i + 6], w: (bytes[i + 7] << 8) | bytes[i + 8] }
        }
        i += 2 + ((bytes[i + 2] << 8) | bytes[i + 3])
      }
      return null
    }
  } catch {
    return null
  }
  return null
}

/** 显示尺寸：优先 HTML 上的 width/height，其次图片原始尺寸，最后兜底 360px 宽 */
export function displaySize(
  bytes: Uint8Array,
  attrs: Record<string, string>,
): { cx: number; cy: number } {
  const MAX_W = 540 // A4 正文栏宽约 595px，留点边距
  const nat = imageSize(bytes)
  let w = parseFloat(attrs.width ?? '')
  let h = parseFloat(attrs.height ?? '')

  if (!Number.isFinite(w) && !Number.isFinite(h) && nat) {
    w = nat.w
    h = nat.h
  }
  if (!Number.isFinite(w) && Number.isFinite(h) && nat) w = (h * nat.w) / nat.h
  if (!Number.isFinite(h) && Number.isFinite(w) && nat) h = (w * nat.h) / nat.w
  if (!Number.isFinite(w) || w <= 0) w = nat?.w ?? 360
  if (!Number.isFinite(h) || h <= 0) h = nat?.h ?? 270

  // 超出正文宽度按比例缩，避免撑破版心
  if (w > MAX_W) {
    h = (h * MAX_W) / w
    w = MAX_W
  }
  return { cx: Math.round(w * EMU_PER_PX), cy: Math.round(h * EMU_PER_PX) }
}

/* ---------- OOXML 片段 ---------- */

let docPrSeq = 1

/**
 * drawingml 相关命名空间。
 *
 * 片段是字符串级插进 document.xml 的，不能指望根元素 w:document 已经声明过这些前缀：
 * 模板里的 w:document 通常带着全套 xmlns，但精简模板 / 手写片段不一定有。
 * 每个片段自带声明是幂等的（重复声明同一 URI 合法），换来片段可独立解析、便于校验。
 */
const NS_WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
const NS_R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

/** inline 型浮动图（随文字流排版，造价报告里的示意图/截图都用这种） */
export function drawingXml(rId: string, cx: number, cy: number, name = '图片'): string {
  const id = 9000 + docPrSeq++
  const safe = String(name).replace(/[<>&"]/g, '')
  return (
    `<w:drawing><wp:inline ${NS_WP} ${NS_R} distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${id}" name="${safe}"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="${safe}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>` +
    `</wp:inline></w:drawing>`
  )
}

/** 浮动锚定型的换行方式：square = 四周型，topAndBottom = 上下型 */
export type ImageWrap = 'square' | 'topAndBottom'

/**
 * 浮动图（文字环绕）。
 *
 * 与 inline 的区别：inline 是「大字符」随文字流排；anchor 脱离文字流，
 * 文字按 wrap 元素绕排（四周型/上下型）。hsdie 只对 square 有意义（图靠左/右，
 * 文字从另一侧绕过）。锚定相对「段落 + 栏」，跟随锚点段落移动 —— 与 WPS
 * 「四周型环绕」的默认行为一致，不需要用户理解绝对定位。
 */
export function drawingAnchorXml(
  rId: string,
  cx: number,
  cy: number,
  wrap: ImageWrap,
  hside: 'left' | 'right',
  name = '图片',
): string {
  const id = 9000 + docPrSeq++
  const safe = String(name).replace(/[<>&"]/g, '')
  const wrapXml = wrap === 'topAndBottom' ? '<wp:wrapTopAndBottom/>' : '<wp:wrapSquare wrapText="bothSides"/>'
  const hAlign = hside === 'right' ? 'right' : 'left'
  return (
    `<w:drawing><wp:anchor ${NS_WP} ${NS_R} distT="0" distB="0" distL="114300" distR="114300"` +
    ` simplePos="0" relativeHeight="${2 + (id % 2000)}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:align>${hAlign}</wp:align></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    wrapXml +
    `<wp:docPr id="${id}" name="${safe}"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="${safe}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>` +
    `</wp:anchor></w:drawing>`
  )
}

/* ---------- image: 控件 → 段落 ---------- */

const escXml = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

/**
 * image: 控件的绑定结果：整段就是一个 w:drawing。
 *
 * @param value 图片源。目前只认 data: URI（上游字段里存 base64，或 UI 选完图转出来的）。
 * @param add   注入通道；为 null 时无法落盘，走文字降级。
 */
export function imageBlockXml(
  value: unknown,
  add: ((bytes: Uint8Array, ext: string) => string | null) | null,
): string {
  const src = value == null ? '' : String(value).trim()
  if (!src) return '<w:p/>'

  const decoded = add ? decodeDataUri(src) : null
  const rId = decoded && add ? add(decoded.bytes, decoded.ext) : null
  if (decoded && rId) {
    const { cx, cy } = displaySize(decoded.bytes, {})
    return `<w:p><w:r>${drawingXml(rId, cx, cy)}</w:r></w:p>`
  }
  // 与富文本同一条铁律：拿不到图也留可见痕迹，绝不静默丢弃
  const label = /^(https?:|file:)/i.test(src) ? src : src.slice(0, 60)
  return `<w:p><w:r><w:t xml:space="preserve">[图片：${escXml(label)}]</w:t></w:r></w:p>`
}

/* ---------- zip 级 sink ---------- */

const REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

function relsPathOf(part: string): string {
  const i = part.lastIndexOf('/')
  return i < 0 ? `_rels/${part}.rels` : `${part.slice(0, i)}/_rels/${part.slice(i + 1)}.rels`
}

export interface ImageSink {
  /**
   * 为一个部件预建 slot（读它现有的 rels，算出下一个可用 rId）。
   * 必须在转换该部件内容之前调用 —— add 依赖 slot 已存在。
   */
  ensure(part: string): Promise<void>
  /** 落一张图，返回该部件 rels 里的 rId；返回 null 表示放弃 */
  add(bytes: Uint8Array, ext: string, part: string): string | null
  /** 写回所有改动的 rels 与 Content_Types，必须在 generateAsync 之前调用 */
  finish(): Promise<void>
  /** 本次注入的图片数 */
  readonly count: number
}

export async function createImageSink(zip: JSZip): Promise<ImageSink> {
  // 每个部件各自维护：现有 rId 最大值、待写入的关系、扩展名集合
  const per = new Map<string, { next: number; rels: string[] }>()
  const exts = new Set<string>()
  let n = 0

  const slot = async (part: string) => {
    let s = per.get(part)
    if (s) return s
    const rp = relsPathOf(part)
    const f = zip.file(rp)
    const txt = f ? await f.async('string') : ''
    const used = [...txt.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]))
    s = { next: Math.max(900, ...used, 0) + 1, rels: [] }
    per.set(part, s)
    return s
  }

  return {
    get count() {
      return n
    },
    async ensure(part) {
      await slot(part)
    },
    add(bytes, ext, part) {
      if (!EXT_MIME[ext]) return null
      const s = per.get(part)
      if (!s) return null // 没 slot 过说明调用方没初始化，直接放弃
      const rId = `rId${s.next++}`
      const file = `word/media/bind-${part.replace(/\W+/g, '')}-${rId}.${ext}`
      zip.file(file, bytes)
      // rels 里的 Target 相对部件所在目录，部件都在 word/ 下
      s.rels.push(
        `<Relationship Id="${rId}" Type="${REL_TYPE}" Target="${file.slice(file.indexOf('/') + 1)}"/>`,
      )
      exts.add(ext)
      n++
      return rId
    },
    async finish() {
      for (const [part, s] of per) {
        if (!s.rels.length) continue
        const rp = relsPathOf(part)
        const f = zip.file(rp)
        const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        if (f) {
          const txt = await f.async('string')
          zip.file(rp, txt.replace('</Relationships>', `${s.rels.join('')}</Relationships>`))
        } else {
          const NS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
          zip.file(rp, `${head}\n<Relationships ${NS}>${s.rels.join('')}</Relationships>`)
        }
      }
      if (!exts.size) return
      const ctFile = zip.file('[Content_Types].xml')
      if (!ctFile) return
      let ct = await ctFile.async('string')
      let dirty = false
      for (const ext of exts) {
        if (ct.includes(`Extension="${ext}"`)) continue
        ct = ct.replace('</Types>', `<Default Extension="${ext}" ContentType="${EXT_MIME[ext]}"/></Types>`)
        dirty = true
      }
      if (dirty) zip.file('[Content_Types].xml', ct)
    },
  }
}
