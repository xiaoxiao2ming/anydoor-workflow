import { renderAsync } from 'docx-preview'

/**
 * 把 docx 渲染成一份完整的可打印 HTML（样式内联、图片 base64 内联，无外部依赖）。
 *
 * 用途：导出 PDF。走「docx → HTML → Chromium 打印」而不是另外接一套 PDF 生成器，
 * 因为渲染引擎与预览面板是同一个（docx-preview），能严格保证
 * 「预览看到的排版 == PDF 拿到的排版」，不会出现两套渲染结果对不上的情况。
 *
 * 必须在渲染进程调用（依赖 DOM）。
 */
export async function docxToHtml(data: Uint8Array): Promise<string> {
  const body = document.createElement('div')
  const style = document.createElement('div')

  await renderAsync(data as unknown as Blob, body, style, {
    inWrapper: true,
    ignoreWidth: false,
    ignoreHeight: false,
    ignoreFonts: false,
    breakPages: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
    // 图片转 base64：产物 HTML 会落到临时文件再打印，相对路径必然失效
    useBase64URL: true,
    trimXmlDeclaration: true,
    experimental: false,
    // 关键：库会把「灰底 / 30px padding / 页间 30px 间距 / 投影」整段包进
    // @media not print，打印时自动失效 —— 比事后用 !important 覆盖干净得多
    hideWrapperOnPrint: true,
  })

  // breakPages 把每页渲染成一个 <section class="docx">，高度正好等于一页。
  // 显式强制分页，避免浮点舍入让下一页溢出 1px 而多出空白页。
  const printCss = `
    @page { size: A4; margin: 0; }
    html, body { margin: 0; padding: 0; background: #fff; }
    * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .docx-wrapper > section.docx { page-break-after: always; break-after: page; }
    .docx-wrapper > section.docx:last-child { page-break-after: auto; break-after: auto; }
  `

  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>report</title><style>${style.innerHTML}${printCss}</style></head><body>${body.innerHTML}</body></html>`
}
