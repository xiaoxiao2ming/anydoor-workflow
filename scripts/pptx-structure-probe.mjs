/**
 * pptx 结构探针：回答「用 jszip + 字符串手术做 PPT 绑定到底可不可行」
 *
 * 只读取、不修改源文件。探的是四件事：
 *   1. 部件构成（和 docx 比有什么不同）
 *   2. 文字在哪、怎么放的（a:t）
 *   3. **占位符会不会被拆散到多个 run** —— 这是 docx 侧踩过的头号坑，
 *      pptx 若同样存在，扫描器必须先做 run 合并，否则占位符根本匹配不到
 *   4. 表格/图片长什么样
 *
 * 跑法：npx vite-node --config vite.scripts.config.ts scripts/pptx-structure-probe.mjs <文件路径>
 */
import { readFileSync } from 'node:fs'
import JSZip from 'jszip'

const file = process.argv[2]
if (!file) {
  console.error('用法: pptx-structure-probe.mjs <file.pptx>')
  process.exit(1)
}

const buf = readFileSync(file)
const zip = await JSZip.loadAsync(buf)
const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir)

console.log(`文件: ${file}`)
console.log(`大小: ${(buf.length / 1024).toFixed(1)} KB`)

/* ── 1. 部件构成 ── */
const byDir = {}
for (const n of names) {
  const d = n.includes('/') ? n.slice(0, n.lastIndexOf('/')) : '(根)'
  byDir[d] = (byDir[d] ?? 0) + 1
}
console.log('\n══ 部件分布 ══')
for (const [d, c] of Object.entries(byDir).sort((a, b) => b[1] - a[1])) {
  console.log(`  ${c.toString().padStart(3)}  ${d}`)
}

const slides = names.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
slides.sort((a, b) => Number(/\d+/.exec(a)[0]) - Number(/\d+/.exec(b)[0]))
console.log(`\nslide 数: ${slides.length}`)
console.log(`版式(layout): ${names.filter((n) => n.startsWith('ppt/slideLayouts/')).length}`)
console.log(`母版(master): ${names.filter((n) => n.startsWith('ppt/slideMasters/')).length}`)
console.log(`媒体(media): ${names.filter((n) => n.startsWith('ppt/media/')).length}`)

/* ── 2~3. 文字与 run 拆分情况 ── */
let totalText = 0
let totalRuns = 0
const splitSamples = []

for (const s of slides) {
  const xml = await zip.file(s).async('string')
  const runs = [...xml.matchAll(/<a:r>([\s\S]*?)<\/a:r>/g)].map((m) => m[1])
  const texts = runs.map((r) => (/<a:t>([\s\S]*?)<\/a:t>/.exec(r)?.[1] ?? '').replace(/\s+/g, ' '))
  totalRuns += runs.length
  totalText += texts.join('').length

  // 「看起来像占位符的片段」：{{ 或 }} 单独出现在一个 run 里，说明占位符被拆了
  const frag = texts.filter((t) => t.includes('{{') || t.includes('}}') || /^[a-z_]+[:.]/i.test(t))
  if (frag.length) splitSamples.push({ slide: s, frag: frag.slice(0, 6) })
}

console.log('\n══ 文字统计 ══')
console.log(`  a:r 总数 ${totalRuns}，文字总长 ${totalText}`)
console.log(`  含 {{ 或 }} 的 slide: ${splitSamples.length}`)

/* ── 3. run 拆分实证：把连续 run 拼起来看，有没有跨 run 的完整占位符 ── */
console.log('\n══ run 拆分检查（占位符是否被拆散）══')
let splitFound = 0
for (const s of slides.slice(0, 30)) {
  const xml = await zip.file(s).async('string')
  // 逐段取 a:p，把段内所有 a:t 拼起来
  const paras = [...xml.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)].map((m) => m[1])
  for (const p of paras) {
    const parts = [...p.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1])
    if (parts.length < 2) continue
    const joined = parts.join('')
    const full = /\{\{[^}]*\}\}/.exec(joined)
    if (!full) continue
    // 完整占位符是否落在单个 run 里？不是 → 被拆了
    const inOneRun = parts.some((t) => t.includes(full[0]))
    if (!inOneRun) {
      splitFound++
      if (splitFound <= 3) {
        console.log(`  ${s} → 「${full[0]}」分散在 ${parts.length} 个 run: ${JSON.stringify(parts)}`)
      }
    }
  }
}
console.log(splitFound ? `  ⚠ 发现 ${splitFound} 处占位符被拆散到多个 run` : '  未发现拆散（注意：该文件可能本就没有占位符，结论待注入后再测）')

/* ── 4. 表格 / 图片 ── */
console.log('\n══ 结构要素 ══')
let tables = 0
let pics = 0
let charts = 0
for (const s of slides) {
  const xml = await zip.file(s).async('string')
  tables += (xml.match(/<a:tbl>/g) ?? []).length
  pics += (xml.match(/<p:pic>/g) ?? []).length
  charts += (xml.match(/<c:chart\s/g) ?? []).length
}
console.log(`  表格 a:tbl: ${tables} | 图片 p:pic: ${pics} | 图表 c:chart: ${charts}`)

/* ── 5. 抽一页看看形状名（pptx-automizer 靠 name 定位元素）── */
const s1 = slides[0]
if (s1) {
  const xml = await zip.file(s1).async('string')
  const shapeNames = [...xml.matchAll(/<p:cNvPr[^>]*name="([^"]*)"/g)].map((m) => m[1])
  console.log(`\n══ ${s1} 的形状名（前 15）══`)
  console.log('  ' + (shapeNames.slice(0, 15).join(' | ') || '（无）'))
  const firstText = (/<a:t>([\s\S]*?)<\/a:t>/.exec(xml)?.[1] ?? '').slice(0, 60)
  console.log(`  首个 a:t: ${JSON.stringify(firstText)}`)
}

console.log('\n══ 结论提示 ══')
console.log('  若上面「未发现拆散」但该文件本就没有 {{}} 占位符，')
console.log('  结论不成立 —— 必须用注入占位符后的副本再测一次（见 pptx-bind-probe.mjs）。')
