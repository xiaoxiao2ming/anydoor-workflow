/**
 * OO Word 保存保真度 · PDF 客观对比工具（2026-10-29 用户指令「转PDF验证/持续迭代」）
 *
 * 原理：两份 docx → 走应用自己的 x2t 管线（/pdfconv，与「转为 PDF 节点」同管道）各转 PDF
 *      → pdfjs 逐页解析（文字行 + 字号 + 字体分布）→ 输出客观差异报告。
 *
 * 用法：node scripts/check-oo-pdf-diff.mjs <参照.docx> <对比.docx>
 *   例：node scripts/check-oo-pdf-diff.mjs "保存留档/2026-10-29-21-05-50_原版" "保存留档/2026-10-29-21-30-08_重建"
 * 退出码：0=文字与主格式一致  1=存在差异（明细见输出）
 */
const POC = 'http://127.0.0.1:10215'
import fs from 'node:fs'
import path from 'node:path'

async function toPdf (docxPath) {
  const buf = fs.readFileSync(docxPath)
  const name = (path.basename(docxPath).replace(/[^\w.-]+/g, '_') || 'x') + '.docx'
  const r = await fetch(POC + '/pdfconv', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: 'fiddiff-' + Date.now().toString(36) + '-' + name, dataBase64: buf.toString('base64') }),
  }).then((x) => x.json())
  if (!r || !r.ok || !r.dataBase64) throw new Error('转 PDF 失败: ' + ((r && r.err) || '无产物'))
  return Buffer.from(r.dataBase64, 'base64')
}

async function analyze (pdfBuf) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(pdfBuf), useSystemFonts: true, isEvalSupported: false }).promise
  const pages = []
  const sizeHist = {}
  const fontHist = {}
  let chars = 0
  let textAll = ''
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p)
    const tc = await page.getTextContent()
    const lines = new Map()
    for (const it of tc.items) {
      if (!it.str) continue
      chars += it.str.length
      textAll += it.str
      const y = Math.round(it.transform[5])
      const key = Math.round(y / 6)
      if (!lines.has(key)) lines.set(key, '')
      lines.set(key, lines.get(key) + it.str)
      const sz = Math.round((it.height || 0) * 10) / 10
      if (sz > 0) sizeHist[sz] = (sizeHist[sz] || 0) + 1
      const fn = (it.fontName || '').replace(/^[A-Z]{6}\+/, '')
      if (fn) fontHist[fn] = (fontHist[fn] || 0) + 1
    }
    const arr = [...lines.entries()].sort((a, b) => b[0] - a[0]).map(([, v]) => v.trim()).filter(Boolean)
    pages.push(arr)
  }
  return { numPages: doc.numPages, pages, chars, sizeHist, fontHist, textAll }
}

function histTop (h, n = 5) {
  return Object.entries(h).sort((a, b) => b[1] - a[1]).slice(0, n).map(([k, v]) => `${k}×${v}`).join(' ')
}

async function main () {
  const [, , aPath, bPath] = process.argv
  if (!aPath || !bPath) { console.error('用法: node scripts/check-oo-pdf-diff.mjs <参照.docx> <对比.docx>'); process.exit(1) }
  console.log('[diff] 参照:', path.basename(aPath))
  console.log('[diff] 对比:', path.basename(bPath))

  console.log('[diff] x2t 转 PDF…')
  const pdfA = await toPdf(aPath)
  const pdfB = await toPdf(bPath)
  console.log(`[diff] A=${pdfA.length}B  B=${pdfB.length}B`)

  console.log('[diff] 解析…')
  const A = await analyze(pdfA)
  const B = await analyze(pdfB)

  let pass = true
  console.log('\n===== 客观指标 =====')
  console.log(`页数      : A=${A.numPages}  B=${B.numPages}${A.numPages !== B.numPages ? '  ✗' : ''}`)
  if (A.numPages !== B.numPages) pass = false
  console.log(`总字符    : A=${A.chars}  B=${B.chars}${A.chars !== B.chars ? '  ✗' : ''}`)
  if (A.chars !== B.chars) pass = false

  const lineDiff = []
  const maxP = Math.max(A.pages.length, B.pages.length)
  for (let p = 0; p < maxP; p++) {
    const la = A.pages[p] || []; const lb = B.pages[p] || []
    const n = Math.max(la.length, lb.length)
    for (let i = 0; i < n; i++) {
      if ((la[i] || '') !== (lb[i] || '')) lineDiff.push(`P${p + 1} 行${i + 1}: A="${(la[i] || '').slice(0, 44)}" | B="${(lb[i] || '').slice(0, 44)}"`)
    }
  }
  /* ★ 折行无关对比：版面密度不同会让折行位置不同（行级必"差"但内容一致）。
     全文拼接去空白后再比——不一致才是真内容差异。 */
  const flatA = A.textAll.replace(/\s+/g, '')
  const flatB = B.textAll.replace(/\s+/g, '')
  const contentSame = flatA === flatB
  if (lineDiff.length) {
    console.log(`\n行级差异 ${lineDiff.length} 处（版面密度/折行不同会造成行级差异，见下方"内容级"判定）:`)
    lineDiff.slice(0, 6).forEach((d) => console.log('  · ' + d))
    if (lineDiff.length > 6) console.log(`  · …共 ${lineDiff.length} 处`)
  } else {
    console.log('\n行级: 全部一致 ✓')
  }
  if (contentSame) {
    console.log('内容级（去空白全文）: 完全一致 ✓ —— 文字保真无缺失（行级差异纯属折行位置）')
  } else {
    pass = false
    // 找出第一处内容差异
    let i = 0; while (i < Math.min(flatA.length, flatB.length) && flatA[i] === flatB[i]) i++
    console.log(`内容级（去空白全文）: ✗ 不一致（A=${flatA.length} 字 B=${flatB.length} 字）`)
    console.log(`  首个差异 @${i}: A="…${flatA.slice(Math.max(0, i - 20), i + 20)}…" | B="…${flatB.slice(Math.max(0, i - 20), i + 20)}…"`)
  }

  console.log(`\n字号分布  : A=[${histTop(A.sizeHist)}]  B=[${histTop(B.sizeHist)}]`)
  const norm = (h) => { const t = Object.values(h).reduce((x, y) => x + y, 0) || 1; const o = {}; for (const [k, v] of Object.entries(h)) o[k] = Math.round((v / t) * 1000) / 1000; return o }
  const szA = JSON.stringify(norm(A.sizeHist)); const szB = JSON.stringify(norm(B.sizeHist))
  if (szA !== szB) { pass = false; console.log('  ✗ 字号分布不一致（字变大/变小会体现在这里）') } else console.log('  ✓ 字号分布一致（占比归一，内容量差异不影响）')
  console.log(`字体分布  : A=[${histTop(A.fontHist)}]  B=[${histTop(B.fontHist)}]`)
  const strip = (h) => JSON.stringify(Object.entries(h).map(([k, v]) => [k.replace(/^g_d\d+_/, ''), v]).sort())
  if (strip(A.fontHist) !== strip(B.fontHist)) { pass = false; console.log('  ✗ 字体分布不一致') } else console.log('  ✓ 字体分布一致（g_dN_ 前缀为文件序号，已归一）')

  console.log('\n=====', pass ? '✅ 保真一致（内容/字号/字体全对齐）' : '❌ 存在保真差异', '=====')
  if (pass && A.numPages !== B.numPages) console.log(`ℹ 版面密度差异：页数 ${A.numPages}→${B.numPages}（页边距/行距未提取，属已知清单项）`)
  process.exit(pass ? 0 : 1)
}

main().catch((e) => { console.error('[diff] 异常:', e); process.exit(1) })
