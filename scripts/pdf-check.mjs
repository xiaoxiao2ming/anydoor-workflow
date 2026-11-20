/**
 * PDF 导出实证：真出一份 PDF 并检查它是不是一份正常的 PDF。
 *
 * Electron 的 printToPDF 与 Chromium 的 page.pdf() 用的是同一套打印管线，
 * 所以这里用 playwright 的 Chromium 验证「docx → HTML → 打印」这段——
 * 排版、分页、图片保真全在这一段，也是唯一会出问题的地方。
 *
 * 跑法：先起 vite（npm run dev:web 或 vite --config vite.web.config.ts --port 5199），
 *      再 node scripts/pdf-check.mjs
 */
import { readFileSync, writeFileSync, statSync } from 'node:fs'
import { chromium } from 'playwright-core'

const BASE = process.env.SMOKE_URL || 'http://127.0.0.1:5199'
const PROBE = BASE + '/scripts/pdf-probe.html'
const OUT = new URL('../sample-template/', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')

const CASES = [
  { file: '市政管道工程量报告-已填充.docx', pdf: '市政管道工程量报告.pdf', minKB: 8 },
  { file: '富文本插图-已填充.docx', pdf: '富文本插图.pdf', minKB: 8 },
]

/** 数 PDF 里的页对象（/Type /Page，不含 /Pages 容器） */
function countPages(buf) {
  const s = buf.toString('latin1')
  return (s.match(/\/Type\s*\/Page[^s]/g) ?? []).length
}

/** 给单个用例套超时，避免一个文件卡死拖垮整轮验证 */
function withTimeout(p, ms, label) {
  return Promise.race([
    p,
    new Promise((_, rej) => setTimeout(() => rej(new Error(`${label} 超时 ${ms}ms`)), ms)),
  ])
}

const run = async () => {
  const browser = await chromium.launch()
  const errors = []

  // 两个页面各司其职：
  //   probe —— 常驻探针页，负责 docxToHtml / XML 校验，全程不导航
  //   print —— 只负责 setContent + page.pdf()
  // 曾经把两者混在一个页面上：setContent 会冲掉定义 window.* 的模块脚本，
  // 导致第二个用例调用已不存在的函数而卡死（表现为"第二个文件永远超时"）。
  const probe = await browser.newPage()
  probe.on('pageerror', (e) => errors.push(String(e)))
  await probe.goto(PROBE, { waitUntil: 'load', timeout: 30000 })
  await probe.waitForFunction(() => window.probeReady === true, null, { timeout: 20000 })

  let print = await browser.newPage()

  let bad = 0
  for (const c of CASES) {
    const src = OUT + c.file
    let sizeKB = 0
    try {
      sizeKB = statSync(src).size / 1024
    } catch {
      console.log(`✗ ${c.file} 不存在，跳过`)
      bad++
      continue
    }
    const b64 = readFileSync(src).toString('base64')

    // ── 先校验 XML：Word 拒绝打开 malformed XML，这一步不过后面全没意义 ──
    const v = await withTimeout(
      probe.evaluate((b) => window.validateDocxB64(b), b64),
      30000,
      'XML 校验',
    )
    if (v.bad.length) {
      console.log(`✗ ${c.file} 有 ${v.bad.length} 个部件 XML 非法：`)
      for (const b of v.bad.slice(0, 5)) console.log(`    ${b.name}: ${b.msg}`)
      bad++
      continue
    }
    console.log(`  XML 校验通过（${v.checked} 个部件）`)

    const html = await withTimeout(
      probe.evaluate((b) => window.docxToHtmlB64(b), b64),
      60000,
      'docxToHtml',
    )

    // 每份文档换一个干净的打印页，避免上一份的样式/图片残留影响分页
    await print.close()
    print = await browser.newPage()
    await print.setContent(html, { waitUntil: 'load' })
    // 等图片解码与 layout 稳定
    await print.waitForTimeout(600)

    const pdf = await withTimeout(
      print.pdf({
        format: 'A4',
        printBackground: true,
        margin: { top: '0', bottom: '0', left: '0', right: '0' },
      }),
      60000,
      '生成 PDF',
    )
    const out = OUT + c.pdf
    writeFileSync(out, pdf)

    const head = pdf.subarray(0, 5).toString('latin1')
    const pages = countPages(pdf)
    const ok = head === '%PDF-' && pdf.length / 1024 >= c.minKB && pages >= 1
    if (!ok) bad++

    console.log(
      `${ok ? '✓' : '✗'} ${c.pdf}  ${(pdf.length / 1024).toFixed(1)}KB  ${pages} 页  ` +
        `头部=${head}  (源 docx ${sizeKB.toFixed(1)}KB)`,
    )
    // 顺手存一份渲染 HTML 便于肉眼比对
    writeFileSync(OUT + c.pdf.replace(/\.pdf$/, '.html'), html)
  }

  if (errors.length) {
    console.log('页面错误', errors.length, '条：')
    for (const e of errors.slice(0, 5)) console.log('  -', e.slice(0, 200))
  }
  console.log(bad ? `\n${bad} 个用例未通过` : '\nPDF 全部生成成功 ✓')

  // browser.close() 在这台机器上偶发挂住（PDF 已写完、日志已打印，就是进程不退出）。
  // 验证结论不依赖清理结果，所以只给它 5 秒，超时就弃掉浏览器进程直接走人——
  // 否则每次跑完都要手动杀，还容易被误判成"脚本卡死"。
  await withTimeout(browser.close(), 5000, '关闭浏览器').catch(() => {})
  process.exit(bad || errors.length ? 2 : 0)
}

run().catch((e) => {
  console.error('PDF 验证失败：', e.message)
  process.exit(1)
})
