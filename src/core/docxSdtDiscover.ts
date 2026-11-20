/**
 * docx 内容控件（sdt）落点发现
 *
 * 立档：2026-10-29 · 补上 OO 域写回在 **Word 侧**的最后一块
 *
 * 为什么需要：core/ooWriteback.patchDocxSdt 要按 tag 定位内容控件才能写值，
 * 而此前**没有任何地方能列出打开的 docx 里有哪些 sdt**
 * （docx/sdtify.ts 只在「导入时」往里塞，不负责「读出来」）
 * ⇒ discoverDocx 传不进来 ⇒ Word 侧一个域也写不了。
 *
 * 纯计算层：字节进 → sdt 列表出，零 IO。
 */

import JSZip from 'jszip'

/** 一个内容控件的落点 */
export interface DocxSdtSite {
  /** w:tag 的值（写回时的定位键） */
  tag: string
  /** 所属部件（默认 word/document.xml） */
  part: string
  /** 当前文本值（便于 UI 展示"这个域现在是几"） */
  text: string
  /** 控件类型：field/rich/each/if/image（tag 形如 'field:名称' 时取前缀） */
  kind?: string
}

const DOC_PARTS = ['word/document.xml', 'word/header1.xml', 'word/footer1.xml', 'word/footnotes.xml', 'word/endnotes.xml']

/**
 * ★ 列出 docx 里所有内容控件。
 *
 * 匹配方式与 core/ooWriteback.patchDocxSdt **保持一致**（同一套正则思路）——
 * ★ 两处若不一致，会出现「发现得到但写不进去」或「写得进去但发现不到」的割裂。
 */
export async function discoverDocxSdt (bytes: Uint8Array): Promise<DocxSdtSite[]> {
  const out: DocxSdtSite[] = []
  let zip
  try {
    zip = await JSZip.loadAsync(bytes)
  } catch {
    return out
  }
  for (const part of DOC_PARTS) {
    const f = zip.file(part)
    if (!f) continue
    let xml = ''
    try {
      xml = await f.async('string')
    } catch {
      continue
    }
    for (const s of scanSdtInXml(xml, part)) out.push(s)
  }
  return out
}

/** 从一段 document.xml 里扫出全部 sdt（导出以便单测，不必走 zip） */
export function scanSdtInXml (xml: string, part = 'word/document.xml'): DocxSdtSite[] {
  const out: DocxSdtSite[] = []
  // 逐个 <w:sdt> 块取tag（用非贪婪 + 阻断 </w:sdt> 保证不跨块匹配）
  const re = /<w:sdt>(?:(?!<\/w:sdt>)[\s\S])*?<\/w:sdt>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml))) {
    const block = m[0]
    const tag = /<w:tag w:val="([^"]*)"/.exec(block)?.[1]
    if (!tag) continue            // 没有 tag 的控件不是我们的域，跳过
    // 内容控件里的文本：取 sdtContent 下的 w:t
    const cm = /<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/.exec(block)
    let text = ''
    if (cm) {
      const ts = [...cm[1].matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
      text = ts.map((x) => x[1]).join('')
    }
    // tag 形如 'field:名称' 时 kind 是前缀
    const kind = /^([a-z]+):/.exec(tag)?.[1]
    out.push({ tag: unescapeXml(tag), part, text: unescapeXml(text), kind })
  }
  return out
}

/**
 * ★ 把 sdt 列表对齐到 ooRefs（生成 writeback 需要的 wordSites）。
 *
 * 匹配口径与 Excel 侧同源（见 core/ooRefSites）：
 *   urn locator（域名）→ 字段名 → fieldId，
 *   其中 Word 的 tag 若带 `field:` 前缀，优先匹配其后的名称。
 */
export function toWordSites (
  sdtList: DocxSdtSite[],
): { tag: string; key?: string; part?: string }[] {
  return sdtList.map((s) => {
    const raw = s.tag
    const colon = /^([a-z]+):(.+)$/.exec(raw)
    return {
      tag: raw,
      // kind:name 形态下，key 是 name（宿主侧的域名/字段名多不带前缀）
      key: colon ? colon[2] : raw,
      part: s.part,
    }
  })
}

function unescapeXml (s: string): string {
  return String(s)
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}
