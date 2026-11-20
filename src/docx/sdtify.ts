import JSZip from 'jszip'

/**
 * 占位符 → 内容控件：把普通 docx 里的 ｛｛field:xx｝｝ 标记一键转成 Word 内容控件。
 * 设计院给的模板是普通文档，这一步是「导入现有模板」的入口。
 *
 * 包裹粒度（与 Word 原生一致，决定「同一段落里能否混排固定文字与变量」）：
 *   field → run 级内联控件，段落里前后的固定文字原样保留
 *   rich  → 段落级块控件，富文本是多段落，必须整块替换
 *   if    → 段落级块控件，布尔为假时整段删除
 *   image → 段落级块控件，值是图片 data: URI，渲染成 w:drawing
 *   each  → 行级块控件，包住整行以便克隆
 */

const KINDS = 'field|rich|each|if|image'
const TAG_RE = new RegExp(`\\{\\{\\s*(${KINDS})\\s*:\\s*([^{}]+?)\\s*\\}\\}`)
const ANY_RE = new RegExp(`\\{\\{\\s*(?:${KINDS})\\s*:\\s*[^{}]+?\\s*\\}\\}`)
const R_RE = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g
const P_RE = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g
const TR_RE = /<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g

export interface SdtifyResult {
  xml: string
  count: number
  tags: string[]
}

function wrap(target: string, tag: string, id: number, kind: string): string {
  // image 用 Word 原生的 picture 控件语义；rich 用 richText；其余是普通文本控件
  const marker = kind === 'rich' ? '<w:richText/>' : kind === 'image' ? '<w:picture/>' : '<w:text/>'
  const alias = tag.replace(new RegExp(`^(?:${KINDS})\\s*:\\s*`), '')
  return (
    `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:alias w:val="${alias}"/>` +
    `<w:tag w:val="${tag}"/>${marker}</w:sdtPr>` +
    `<w:sdtContent>${target}</w:sdtContent></w:sdt>`
  )
}

export function sdtify(xml: string): SdtifyResult {
  interface Target {
    start: number
    end: number
    tag: string
    body: string
    kind: string
  }
  const targets: Target[] = []
  const taken: { start: number; end: number }[] = []
  const free = (start: number, end: number) => !taken.some((t) => start < t.end && end > t.start)
  const take = (t: Target) => {
    targets.push(t)
    taken.push({ start: t.start, end: t.end })
  }

  const scan = (re: RegExp, kinds: string[]) => {
    re.lastIndex = 0
    for (const m of xml.matchAll(re)) {
      const body = m[0]
      const start = m.index
      const end = start + body.length
      if (!free(start, end)) continue
      const t = TAG_RE.exec(body)
      if (!t || !kinds.includes(t[1])) continue
      const kind = t[1]
      const tag = `${kind}:${t[2]}`
      // each / if / image 只是声明标记，不参与最终显示 → 从内容里剔除
      const cleaned =
        kind === 'each' || kind === 'if' || kind === 'image'
          ? body.replace(ANY_RE, (s) => (s === t[0] ? '' : s))
          : body
      take({ start, end, tag, body: cleaned, kind })
    }
  }

  scan(TR_RE, ['each'])
  scan(P_RE, ['rich', 'if', 'image'])
  scan(R_RE, ['field'])

  // 从后往前替换，避免下标失效
  targets.sort((a, b) => b.start - a.start)
  let out = xml
  let id = 1000
  for (const t of targets) {
    out = out.slice(0, t.start) + wrap(t.body, t.tag, id++, t.kind) + out.slice(t.end)
  }
  return { xml: out, count: targets.length, tags: targets.map((t) => t.tag).reverse() }
}

/** 模板正文与页眉页脚里可能出现的绑定位置 */
const PARTS = ['word/document.xml', 'word/header1.xml', 'word/footer1.xml', 'word/header2.xml', 'word/footer2.xml']

/**
 * zip 级 sdtify：设计院给的模板是普通 docx，正文/页眉里有 ｛｛field:xx｝｝ 标记，
 * 这一步把标记原地换成 Word 内容控件，之后 Word 打开能看到控件、引擎能按 tag 定位。
 *
 * 只在对应部件存在且确实含标记时才改动该部件，其余部件字节不动。
 */
export async function sdtifyDocx(input: ArrayBuffer | Uint8Array | Buffer): Promise<{
  buffer: Uint8Array
  count: number
  tags: string[]
  changed: string[]
}> {
  const zip = await JSZip.loadAsync(input)
  const tags: string[] = []
  const changed: string[] = []

  for (const name of PARTS) {
    const f = zip.file(name)
    if (!f) continue
    const xml = await f.async('string')
    const r = sdtify(xml)
    if (!r.count) continue
    zip.file(name, r.xml)
    tags.push(...r.tags)
    changed.push(name)
  }

  const buffer = await zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
  return { buffer, count: tags.length, tags, changed }
}
