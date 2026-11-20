#!/usr/bin/env node
/**
 * 单个文件的往返**详解**（排查用）
 *
 * 与 check-regress.mjs 的分工：
 *   · check-regress  → 清单全量 + 汇总 + 退出码，做**门禁**（升级依赖前后跑）
 *   · 本脚本         → 一个文件、逐部件列清单、结果落 JSON，做**排查**
 *                      （门禁说「丢了 xl/tables/table1.xml」时，用这个看细节）
 *
 * 用法：node scripts/check-roundtrip.mjs [xlsx 路径]
 *      不带参数时用清单里的第一个文件。
 */
import fs from 'node:fs'
import path from 'node:path'
import { roundtripOne, unzip, categoryOf, isDirEntry, OUT_DIR, loadList } from './lib/roundtrip.mjs'

const list = loadList()
const fallback = list.files?.[0]?.path
const file = process.argv[2] || fallback

if (!file) {
  console.error('用法：node scripts/check-roundtrip.mjs <xlsx 路径>')
  console.error('（清单里也没有文件，先配置 scripts/regress-list.json）')
  process.exit(2)
}

const abs = path.isAbsolute(file) ? file : path.join(process.cwd(), file)
if (!fs.existsSync(abs)) {
  console.error('文件不存在：' + abs)
  process.exit(2)
}

console.log('══════════════════════════════════════════════════════════════')
console.log('  往返详解')
console.log('  ' + abs)
console.log('══════════════════════════════════════════════════════════════')

// ── 源文件结构 ──
const srcBuf = fs.readFileSync(abs)
const before = unzip(srcBuf)
const srcParts = Object.keys(before).filter((n) => !isDirEntry(n))

console.log('')
console.log(`源文件 ${(srcBuf.length / 1024).toFixed(1)} KB · 真文件 ${srcParts.length} 个（另有 ${Object.keys(before).length - srcParts.length} 个目录标记）`)
const byCat = {}
for (const n of srcParts) (byCat[categoryOf(n)] ||= []).push(n)
for (const [k, v] of Object.entries(byCat).sort((a, b) => b[1].length - a[1].length)) {
  console.log('  ' + k.padEnd(14) + String(v.length).padStart(3) + ' 个')
}

console.log('')
console.log('── 往返过程 ──')
const r = await roundtripOne({ path: abs, label: path.basename(abs), origin: 'adhoc' }, { log: (s) => console.log(s), keepOut: true })

if (r.status !== 'ok') {
  console.log('')
  console.log(`❌ 未通过：${r.status}`)
  if (r.openError) {
    console.log(`   分类：${r.openError.kind}`)
    console.log(`   现象：${r.openError.title}`)
    if (r.openError.advice) console.log(`   建议：${r.openError.advice}`)
  }
  if (r.error) console.log(`   原始错误：${r.error}`)
  process.exitCode = 1
  process.exit()
}

// ── 逐部件 ──
const outPath = path.join(OUT_DIR, `roundtrip-${path.basename(abs)}`)
const outBuf = fs.readFileSync(outPath)
const after = unzip(outBuf)

console.log('')
console.log(`产物 ${(outBuf.length / 1024).toFixed(1)} KB · ${outPath}`)
console.log('')
console.log('── 逐部件 ──')
for (const n of srcParts) {
  const mark = !(n in after) ? '❌ 丢失' : before[n].crc === after[n].crc ? '✅ 一致' : '✏️  改动'
  console.log('  ' + mark.padEnd(8) + categoryOf(n).padEnd(12) + n)
}
const added = Object.keys(after).filter((n) => !isDirEntry(n) && !(n in before))
if (added.length) {
  console.log('')
  console.log('  新增部件：')
  for (const n of added) console.log('    ➕ ' + categoryOf(n).padEnd(12) + n)
}

console.log('')
console.log('── 数据完整性 ──')
for (const s of r.cells.nonEmptyDeltas) {
  const flag = s.before === s.after ? '✅' : '⚠️ '
  console.log(`  ${flag} ${s.sheet}: 非空 ${s.before} → ${s.after}`)
}
console.log(`  逐格 ${r.cells.checked} 格，除目标格外差异 ${r.cells.diff} ${r.cells.diff === 0 ? '✅' : '❌'}`)
if (r.cells.diff > 0) {
  for (const d of r.cells.diffSamples) console.log(`     ${d.sheet}!${d.coord}  ${JSON.stringify(d.from)} → ${JSON.stringify(d.to)}`)
}
if (r.target) {
  console.log(`  目标格 ${r.target.sheet}!${r.target.coord}：原值 ${JSON.stringify(r.target.old)} → 哨兵值`)
}

// ── 落盘明细，便于人工复核 ──
const report = {
  src: abs,
  srcBytes: r.srcBytes,
  outBytes: r.outBytes,
  openedRaw: r.openedRaw,
  applied: r.applied,
  sheets: r.sheets,
  target: r.target,
  parts: {
    src: r.parts.src,
    out: r.parts.out,
    same: r.parts.same,
    changed: r.parts.changed,
    added: r.parts.added,
    lost: r.parts.lost,
    criticalLost: r.parts.criticalLost,
  },
  cells: {
    checked: r.cells.checked,
    diff: r.cells.diff,
    diffSamples: r.cells.diffSamples,
    nonEmptyDeltas: r.cells.nonEmptyDeltas,
  },
}
fs.writeFileSync(path.join(OUT_DIR, 'roundtrip-report.json'), JSON.stringify(report, null, 1), 'utf8')

console.log('')
console.log('  明细：sample-template/roundtrip-report.json')
console.log('  产物：sample-template/' + path.basename(outPath))

const bad = r.parts.criticalLost.length > 0 || r.cells.diff > 0
console.log('')
console.log(bad ? '❌ 存在风险项，请人工复核' : '✅ 补丁式往返成立：关键部件零丢失 + 单元格零差异')
process.exitCode = bad ? 1 : 0
