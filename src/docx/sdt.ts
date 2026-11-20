/**
 * 内容控件（w:sdt）扫描器
 *
 * 为什么用字符串级扫描而不是 XML 解析：
 *   docx 是 zip 包，document.xml 里除了绑定点还有大量无关内容。
 *   字符串级"外科手术"能保证非目标区域字节不变，格式保真度最高；
 *   XML parse → serialize 的往返会重排属性、丢掉未知扩展，得不偿失。
 */

export interface SdtBlock {
  start: number
  end: number
  tag: string
  contentStart: number
  contentEnd: number
  contentXml: string
  hasRow: boolean
}

/** 从 start 处的 <w:sdt> 出发，返回配对 </w:sdt> 的结束下标 */
function matchSdtEnd(xml: string, start: number): number {
  const re = /<w:sdt(?:\s[^>]*)?>|<\/w:sdt>/g
  re.lastIndex = start
  let depth = 0
  let m: RegExpExecArray | null
  while ((m = re.exec(xml))) {
    if (m[0].startsWith('</')) {
      if (--depth === 0) return m.index + m[0].length
    } else depth++
  }
  return -1
}

const firstMatch = (re: RegExp, s: string): RegExpExecArray | null => {
  re.lastIndex = 0
  return re.exec(s)
}

export function decodeXml(s: string): string {
  return String(s)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/** 扫描一段 XML 里的顶层内容控件（带可识别 tag 的） */
export function findSdtBlocks(xml: string): SdtBlock[] {
  const out: SdtBlock[] = []
  const openRe = /<w:sdt(?:\s[^>]*)?>/g
  let m: RegExpExecArray | null
  while ((m = openRe.exec(xml))) {
    const start = m.index
    const end = matchSdtEnd(xml, start)
    if (end < 0) continue
    const body = xml.slice(start, end)

    const pr = firstMatch(/<w:sdtPr(?:\s[^>]*)?>([\s\S]*?)<\/w:sdtPr>/, body)
    const prXml = pr ? pr[1] : ''
    const tagM = firstMatch(/<w:tag\b[^>]*\bw:val="([^"]*)"/, prXml)
    const tag = tagM ? decodeXml(tagM[1]) : ''
    if (!tag) {
      openRe.lastIndex = end
      continue
    }

    const cOpen = firstMatch(/<w:sdtContent(?:\s[^>]*)?>/, body)
    if (!cOpen) {
      openRe.lastIndex = end
      continue
    }
    const contentStart = start + cOpen.index + cOpen[0].length
    const contentEnd = start + body.indexOf('</w:sdtContent>')
    if (contentEnd < contentStart) {
      openRe.lastIndex = end
      continue
    }

    const contentXml = xml.slice(contentStart, contentEnd)
    out.push({
      start,
      end,
      tag,
      contentStart,
      contentEnd,
      contentXml,
      hasRow: /<w:tr(?:\s[^>]*)?>/.test(contentXml),
    })
    openRe.lastIndex = end
  }
  return out
}

/** 剥掉 sdt 外壳，只留内部内容 */
export function unwrapSdt(block: SdtBlock): string {
  return block.contentXml
}

/** 用新内容重建 sdt（原文档前缀、控件外壳与 sdtPr 全部原样保留） */
export function rewrapSdt(xml: string, block: SdtBlock, newContent: string): string {
  const prefix = xml.slice(0, block.start)
  const head = xml.slice(block.start, block.contentStart) // <w:sdt>...<w:sdtContent>
  const tail = xml.slice(block.contentEnd) // </w:sdtContent>...</w:sdt>
  return prefix + head + newContent + tail
}

/**
 * 剥掉所有内容控件外壳（含无 tag 的），只留 sdtContent 内容。
 * 用途：交给 docx-preview 渲染 —— 它对「行级控件」（sdt 包 w:tr）支持不完整，
 * 会把数据行整行吞掉。剥壳不影响视觉，Word/WPS 打开效果完全一致。
 */
export function stripAllSdt(xml: string): string {
  const openRe = /<w:sdt(?:\s[^>]*)?>/g
  let m: RegExpExecArray | null
  while ((m = openRe.exec(xml))) {
    const start = m.index
    const end = matchSdtEnd(xml, start)
    if (end < 0) break
    const body = xml.slice(start, end)
    const co = body.indexOf('<w:sdtContent')
    if (co < 0) {
      openRe.lastIndex = end
      continue
    }
    const cStart = co + body.slice(co).indexOf('>') + 1
    const cEnd = body.lastIndexOf('</w:sdtContent>')
    const inner = cEnd > cStart ? body.slice(cStart, cEnd) : ''
    xml = xml.slice(0, start) + inner + xml.slice(end)
    openRe.lastIndex = start // 从同位置继续，处理剥出来的嵌套控件
  }
  return xml
}
