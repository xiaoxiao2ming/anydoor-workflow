#!/usr/bin/env node
/**
 * 真实文件回归门禁 —— 升依赖前后的那把尺子
 *
 * 为什么必须存在：
 *   本项目唯一会**改写用户文件**的依赖是 @office-kit/xlsx。
 *   它的往返（打开 → 改一格 → 存回）必须在真实文件上验证：
 *   所有「解析成自有模型再序列化」的库都会把不认识的部件**静默丢弃**，
 *   用户改一个数字、存回去发现图表没了，这是交付事故级的问题。
 *
 *   而它恰恰是最容易被升级毁掉的东西 —— 一个次要版本就可能收紧/放宽 schema，
 *   让原本能开的文件突然开不了（本机实测过一次：0.18.0 打不开 WPS 导出的文件）。
 *
 * 所以：**升级 @office-kit/* 之后，先跑这个，绿了再合。**
 *
 * 用法：
 *   node scripts/check-regress.mjs             # 跑清单全量（门禁）
 *   node scripts/check-regress.mjs -v          # 附逐文件详细日志
 *   node scripts/check-regress.mjs --file <p>  # 只跑一个文件（临时验证）
 *   node scripts/check-regress.mjs --list      # 只看清单
 *
 * 退出码：0 = 全部通过；1 = 有必须修的问题（未打开 / 关键部件丢失 / 数据差异）
 */
import path from 'node:path'
import { loadList, roundtripOne, categoryOf } from './lib/roundtrip.mjs'

const argv = process.argv.slice(2)
const verbose = argv.includes('-v') || argv.includes('--verbose')
const listOnly = argv.includes('--list')
const fileArg = argv.includes('--file') ? argv[argv.indexOf('--file') + 1] : null

const list = loadList()
let entries = list.files ?? []
if (fileArg) entries = [{ path: fileArg, label: '（命令行指定）', origin: 'adhoc', expectOpen: true }]

if (listOnly) {
  console.log(`清单：${entries.length} 个文件`)
  for (const e of entries) {
    console.log(`  [${e.origin ?? '?'}]${e.expectOpen === false ? '[可选]' : ''} ${e.label ?? e.path}`)
    console.log(`        ${e.path}`)
    if (e.note) console.log(`        ↳ ${e.note}`)
  }
  process.exit(0)
}

console.log('══════════════════════════════════════════════════════════════')
console.log('  真实文件回归门禁')
console.log(`  清单 ${entries.length} 个文件${fileArg ? '（已用 --file 覆盖）' : ' · scripts/regress-list.json'}`)
console.log('══════════════════════════════════════════════════════════════')

/** 一行摘要：打开状态 · 规模 · 部件 · 逐格 */
function summaryLine(r) {
  if (r.status === 'missing') return '样本不在本机（跳过，不计入通过率）'
  if (r.status === 'open-failed') return `打开 ❌ [${r.openError?.kind ?? '?'}] ${r.openError?.title ?? ''}`
  if (r.status === 'error') return `出错 ⚠ ${String(r.error).slice(0, 90)}`
  return [
    `打开 ✅${r.openedRaw ? '（原文即可）' : `（归一化 ${r.applied.length} 处）`}`,
    `工作表 ${r.sheets.length}`,
    `非空 ${r.nonEmpty}`,
    `部件 ${r.parts.src}→${r.parts.out}（一致 ${r.parts.same} / 改动 ${r.parts.changed} / 丢失 ${r.parts.lost.length}）`,
    `逐格 ${r.cells.checked} 差异 ${r.cells.diff}`,
  ].join(' · ')
}

const results = []
for (let i = 0; i < entries.length; i++) {
  const entry = entries[i]
  console.log('')
  console.log(`[${i + 1}/${entries.length}] ${entry.label ?? path.basename(entry.path)}`)
  console.log(`        ${entry.path}`)
  const r = await roundtripOne(entry, { log: verbose ? (s) => console.log(s) : () => {} })
  results.push({ entry, r })
  console.log('        ' + summaryLine(r))
  if (r.status === 'open-failed') {
    console.log(`        ⚠ ${r.openError?.kind}：${r.openError?.title}`)
    if (r.openError?.advice) console.log(`          建议：${r.openError.advice}`)
    if (r.parts?.lost?.length) console.log(`          丢失部件：${r.parts.lost.slice(0, 4).join(', ')}`)
    else if (r.error) console.log(`          原始错误：${String(r.error).slice(0, 140)}`)
  } else if (r.status === 'error') {
    console.log(`        ⚠ ${r.error}`)
  } else if (r.status === 'ok' && r.cells.diff > 0) {
    console.log(`        ⚠ 数据差异 ${r.cells.diff} 处，前几个：`)
    for (const d of r.cells.diffSamples.slice(0, 4)) {
      console.log(`          ${d.sheet}!${d.coord}  ${JSON.stringify(d.from)} → ${JSON.stringify(d.to)}`)
    }
  } else if (r.status === 'ok' && r.parts.criticalLost.length) {
    console.log(`        ⚠ 关键部件丢失：${r.parts.criticalLost.join(', ')}`)
  }
}

/* ────────────────── 汇总 ────────────────── */

console.log('')
console.log('══════════════════════════════════════════════════════════════')
console.log('  汇总')
console.log('══════════════════════════════════════════════════════════════')
console.log(
  '  ' +
    '文件'.padEnd(30) +
    '来源'.padEnd(6) +
    '打开'.padEnd(6) +
    '归一'.padEnd(5) +
    '部件丢'.padEnd(7) +
    '关键'.padEnd(5) +
    '逐格差'.padEnd(7) +
    '数据量'
)
for (const { entry, r } of results) {
  const label = (entry.label ?? path.basename(entry.path)).slice(0, 28)
  const origin = entry.origin ?? '?'
  const open = r.status === 'ok' ? '✅' : r.status === 'missing' ? '—' : '❌'
  const norm = r.status === 'ok' ? String(r.applied.length) : '-'
  const lostN = r.parts ? String(r.parts.lost.length) : '-'
  const critN = r.parts ? String(r.parts.criticalLost.length) : '-'
  const diffN = r.cells ? String(r.cells.diff) : '-'
  const volume = r.status === 'ok' ? `${r.sheets.length}表/${r.nonEmpty}非空/${r.cells.checked}格` : ''
  console.log(
    '  ' +
      label.padEnd(30) +
      origin.padEnd(6) +
      open.padEnd(6) +
      norm.padEnd(5) +
      lostN.padEnd(7) +
      critN.padEnd(5) +
      diffN.padEnd(7) +
      volume
  )
}

/* ────────────────── 判定 ────────────────── */

const failures = []
const skipped = []
for (const { entry, r } of results) {
  const required = entry.expectOpen !== false
  if (r.status === 'missing') {
    // 样本不存在不算失败（本机文件可能被移动过），但要明确说出来，
    // 否则「6 个文件全绿」可能是「其实只跑了 2 个」的假象
    skipped.push(entry.label ?? entry.path)
    continue
  }
  if (r.status === 'open-failed') {
    if (required) failures.push(`${entry.label}：打不开（${r.openError?.kind}）`)
    else skipped.push(`${entry.label}（标记为可选：${r.openError?.kind}）`)
    continue
  }
  if (r.status === 'error') {
    failures.push(`${entry.label}：${r.error}`)
    continue
  }
  if (r.parts.criticalLost.length) {
    failures.push(`${entry.label}：关键部件丢失 ${r.parts.criticalLost.join(', ')}`)
  }
  if (r.cells.diff > 0) {
    failures.push(`${entry.label}：单元格数据差异 ${r.cells.diff} 处`)
  }
}

console.log('')
const ran = results.filter(({ r }) => r.status === 'ok').length
const answered = results.filter(({ r }) => r.status !== 'missing').length
console.log(`  实际运行 ${ran}/${answered} 个（另有 ${results.length - answered} 个样本不在本机）`)
if (skipped.length) {
  console.log(`  跳过：${skipped.join('；')}`)
  console.log('  ⚠ 注意：「全绿」要连同运行数一起看 —— 样本缺失时的绿是空绿。')
}

// 部件丢失/数据差异的分类统计（一张图看清是哪一类问题）
const lostByCat = {}
for (const { r } of results) {
  if (r.status !== 'ok') continue
  for (const n of r.parts.lost) lostByCat[categoryOf(n)] = (lostByCat[categoryOf(n)] ?? 0) + 1
}
if (Object.keys(lostByCat).length) {
  console.log('  丢失部件按类：' + Object.entries(lostByCat).map(([k, v]) => `${k}×${v}`).join('、'))
}

console.log('')
if (failures.length === 0) {
  console.log('  ✅ 通过：可打开 · 关键内容部件零丢失 · 单元格零差异')
  process.exitCode = 0
} else {
  console.log(`  ❌ ${failures.length} 项必须处理：`)
  for (const f of failures) console.log('     · ' + f)
  console.log('')
  console.log('  下一步：把失败文件喂给 src/xlsx/normalize.ts（失败样本已留在')
  console.log('          userData/open-failures 或浏览器 localStorage），补完后重跑本脚本。')
  process.exitCode = 1
}
