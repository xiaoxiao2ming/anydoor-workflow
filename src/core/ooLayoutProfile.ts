/* ★★ 2026-10-30 版式档案（v3 收口）★★
 *  从 docx 字节解包「行距列表 + 分页符段落 + 页面设置」——样式化重建的版式基准。
 *  为什么建档进 config：重建的 lineList 若只看「已存字节」，而已存字节是旧版
 *  （无行距）的重建产物 ⇒ 永远拿不到行距（死循环，N164 实测）。首份正确字节的
 *  版式固化为 ooLineProfile 后，重建永远有可靠来源。
 *  ★ 「导入即建档」（2026-10-30 用户令「彻底解决」）：boot 加载原版字节时若节点
 *  无档案立即提取固化——新文档天生带档案，保存链任何成败都有版式兜底。 */
export type OoLayoutProfile = {
  /** 档案结构版本：4=含底纹/间距/缩进/样式表段距/页脚全维度 */
  v?: number
  lineList: Array<{ line?: number; rule?: string; indFirstLine?: number; before?: number; after?: number; shdFill?: string }>
  pageBreakIdx: number[]
  /** ★ 2026-10-31 样式表默认段距（styles.xml Normal/docDefaults）——正文段无直排 spacing 时的行距来源 */
  defaultSpacing?: { line?: number; rule?: string; before?: number; after?: number }
  /** ★ 2026-10-31 样式表默认字号（Normal 的 w:sz 半磅值）——run 无 size 时的字号来源 */
  defaultFontSize?: number
  /** ★ 2026-10-31 页脚页码（公文「— X —」式）：prefix/suffix 为 PAGE 字段前后文本 */
  footer?: { prefix?: string; suffix?: string }
  sectionProps?: {
    page: { margin: { top: number; right: number; bottom: number; left: number; header?: number; footer?: number } }
    grid?: { linePitch: number; type?: string }
  }
}

export async function extractOoLayoutProfile(bytes: Uint8Array): Promise<OoLayoutProfile | null> {
  try {
    if (bytes.length < 100) return null
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(bytes)
    const xml = await zip.file('word/document.xml')?.async('string')
    if (!xml) return null
    const bodyPart = xml.slice(xml.indexOf('<w:body>'))
    const noTbl = bodyPart.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, '')
    const xmlParas = noTbl.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? []
    const pageBreakIdx: number[] = []
    xmlParas.forEach((px, pi) => { if (px.includes('<w:br w:type="page"/>')) pageBreakIdx.push(pi) })
    const lineList = xmlParas.map((px) => {
      const mSp = px.match(/<w:spacing[^>]*\/>/)
      const line = mSp ? mSp[0].match(/w:line="(\d+)"/)?.[1] : undefined
      const rule = mSp ? mSp[0].match(/w:lineRule="(\w+)"/)?.[1] : undefined
      /* ★ 缩进也入档案：原版直排 firstLine（0=显式顶格/480=2字缩进），重建时档案优先——
         引擎半残时报错值（如统一 640）不再污染版式 */
      const ind = px.match(/<w:ind[^>]*\/>/)
      const indFL = ind ? ind[0].match(/w:firstLine="(-?\d+)"/)?.[1] : undefined
      /* ★ 段落底纹（截图里的粉/蓝高亮）也入档案 */
      const shd = px.match(/<w:shd [^>]*w:fill="([0-9A-Fa-f]{6})"[^>]*\/>/)
      const shdFill = shd && shd[1].toUpperCase() !== 'AUTO' ? shd[1] : undefined
      const before = mSp ? Number(mSp[0].match(/w:before="(\d+)"/)?.[1] ?? 0) || undefined : undefined
      const after = mSp ? Number(mSp[0].match(/w:after="(\d+)"/)?.[1] ?? 0) || undefined : undefined
      return { line: line ? Number(line) : undefined, rule, indFirstLine: indFL !== undefined ? Number(indFL) : undefined, before, after, shdFill }
    })
    const sect = xml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/)?.[0] ?? ''
    /* ★ 样式表默认段距：Normal 段落样式优先，退 docDefaults（正文无直排 spacing 时的真实行距来源） */
    let defaultSpacing: { line?: number; rule?: string; before?: number; after?: number } | undefined
    let defaultFontSize: number | undefined
    /* ★ 页脚页码：footer*.xml 里 PAGE 字段前后的文本（公文「— X —」式）。
       页码是重建最常丢的元素（docx 库 sections 不自动带页脚），档案记 prefix/suffix。 */
    let footer: { prefix?: string; suffix?: string } | undefined
    try {
      const footerFiles = ['word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml']
      for (const ff of footerFiles) {
        const fxml = await zip.file(ff)?.async('string')
        if (!fxml || !fxml.includes('PAGE')) continue
        /* 按文档顺序切：PAGE 字段前的文本 = prefix，后的 = suffix */
        const pagePos = fxml.indexOf('PAGE')
        const before = fxml.slice(0, pagePos)
        const after = fxml.slice(pagePos)
        const collect = (seg: string) => (seg.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) ?? []).map((t) => t.replace(/<[^>]+>/g, '')).join('').trim()
        footer = { prefix: collect(before) || undefined, suffix: collect(after) || undefined }
        break
      }
    } catch { /* 页脚解析失败不影响 */ }
    try {
      const sxml = await zip.file('word/styles.xml')?.async('string')
      if (sxml) {
        const normalSt = sxml.match(/<w:style [^>]*w:default="1"[^>]*w:type="paragraph"[^>]*>[\s\S]*?<\/w:style>/)?.[0]
          || sxml.match(/<w:style [^>]*w:type="paragraph"[^>]*w:default="1"[^>]*>[\s\S]*?<\/w:style>/)?.[0]
          || sxml.match(/<w:style [^>]*w:styleId="(?:Normal|a)"[^>]*>[\s\S]*?<\/w:style>/)?.[0]
        const dd = sxml.match(/<w:docDefaults>[\s\S]*?<\/w:docDefaults>/)?.[0] ?? ''
        const src = (normalSt && normalSt.includes('<w:spacing')) ? normalSt : dd
        const dsp = src.match(/<w:spacing[^>]*\/>/)?.[0]
        if (dsp) {
          const dline = dsp.match(/w:line="(\d+)"/)?.[1]
          const drule = dsp.match(/w:lineRule="(\w+)"/)?.[1]
          const dbefore = dsp.match(/w:before="(\d+)"/)?.[1]
          const dafter = dsp.match(/w:after="(\d+)"/)?.[1]
          if (dline || dbefore || dafter) defaultSpacing = { line: dline ? Number(dline) : undefined, rule: drule, before: dbefore ? Number(dbefore) : undefined, after: dafter ? Number(dafter) : undefined }
          /* ★ 默认字号：Normal 样式的 w:sz（半磅）——run 无显式 size 时的字号来源
             （重建产物没样式层，docx 库默认 11pt 比原版小一大截，14:52 实测 14pt→10pt） */
          const dsz = normalSt ? normalSt.match(/<w:sz w:val="(\d+)"\/>/)?.[1] : undefined
          defaultFontSize = dsz ? Number(dsz) : undefined
        }
      }
    } catch { /* 样式表解析失败不影响 */ }
    const num = (tag: string, attr: string) => Number(sect.match(new RegExp(`<w:${tag}[^>]*w:${attr}="(\\d+)"`))?.[1] ?? 0)
    const gridPitch = Number(sect.match(/<w:docGrid[^>]*w:linePitch="(\d+)"/)?.[1] ?? 0)
    const gridType = sect.match(/<w:docGrid[^>]*w:type="(\w+)"/)?.[1]
    const sectionProps = num('pgMar', 'top')
      ? {
          page: {
            margin: {
              top: num('pgMar', 'top'), right: num('pgMar', 'right'), bottom: num('pgMar', 'bottom'), left: num('pgMar', 'left'),
              ...(num('pgMar', 'header') ? { header: num('pgMar', 'header') } : {}),
              ...(num('pgMar', 'footer') ? { footer: num('pgMar', 'footer') } : {}),
            },
          },
          ...(gridPitch ? { grid: { linePitch: gridPitch, ...(gridType ? { type: gridType } : {}) } } : {}),
        }
      : undefined
    /* 三样全空 = 这份字节本身没有版式信息（纯文本产物）——不建档，等正确字节。
       ★ 2026-10-31 防呆：sectPr 全是 docx 库默认值（四边 1440）也视为无 sectionProps——
       否则坏产物会建出「默认边距档案」，覆盖不了也救不了（726390 实测死循环变体） */
    const isDefaultMargin = sectionProps
      && sectionProps.page.margin.top === 1440 && sectionProps.page.margin.right === 1440
      && sectionProps.page.margin.bottom === 1440 && sectionProps.page.margin.left === 1440
    if ((!lineList.some((x) => x.line) && !pageBreakIdx.length && !sectionProps)
      || (!lineList.some((x) => x.line) && !pageBreakIdx.length && isDefaultMargin)) return null
    if (isDefaultMargin && sectionProps) {
      /* 边距是默认值但有分页符——保留分页，丢掉默认边距 */
      return { v: 4, lineList, pageBreakIdx, sectionProps: undefined, defaultSpacing, defaultFontSize, footer }
    }
    return { v: 4, lineList, pageBreakIdx, sectionProps, defaultSpacing, defaultFontSize, footer }
  } catch {
    return null
  }
}
