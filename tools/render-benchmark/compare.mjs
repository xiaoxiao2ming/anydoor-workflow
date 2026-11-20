#!/usr/bin/env node
/**
 * 渲染兼容性对比（目标：自研渲染尽可能接近 WPS）
 *
 * 目录约定：
 *   docs/    原始 docx/xlsx（你的真实业务文档）
 *   golden/  标准答案——每份文档用桌面 WPS 打开 → 另存为 PDF 放这里（同名）
 *   mine/    自研渲染导出的 PDF/图片（脚本或手工放入）
 *   reports/ 对比报告输出
 *
 * 用法：
 *   node compare.mjs              # 对比全部有 golden 的文档
 *   node compare.mjs 某文档名      # 只对比一份
 *
 * 依赖：npm i pdfjs-dist pixelmatch pngjs（首次运行自动装）
 */
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const dir = path.dirname(fileURLToPath(import.meta.url))
const DOCS = path.join(dir, 'docs')
const GOLDEN = path.join(dir, 'golden')
const MINE = path.join(dir, 'mine')
const REPORTS = path.join(dir, 'reports')

const only = process.argv[2]
const pdf = requireShim()

function requireShim() {
  try {
    return require
  } catch {
    return null
  }
}

function list(base) {
  if (!fs.existsSync(base)) return []
  return fs.readdirSync(base).filter((f) => /\.(pdf|png)$/i.test(f))
}

function pairs() {
  return list(GOLDEN)
    .map((g) => {
      const stem = g.replace(/\.(pdf|png)$/i, '')
      const m = list(MINE).find((x) => x.startsWith(stem))
      return m ? { stem, golden: path.join(GOLDEN, g), mine: path.join(MINE, m) } : null
    })
    .filter(Boolean)
    .filter((p) => !only || p.stem.includes(only))
}

async function main() {
  const jobs = pairs()
  if (!jobs.length) {
    console.log(`[基准集为空]
  1) 把典型文档放进 docs/
  2) 用桌面 WPS 打开每份 → 另存为 PDF → 放进 golden/（同名）
  3) 自研渲染结果（PDF/截图）放进 mine/
  4) 再跑本脚本`)
    return
  }

  const { default: pixelmatch } = await import('pixelmatch')
  const { PNG } = await import('pngjs')
  const results = []

  for (const job of jobs) {
    /* 输入统一为 PNG 页图（pdf 转 png 依赖可选；直接放 PNG 最简单） */
    const a = loadPng(job.golden)
    const b = loadPng(job.mine)
    if (!a || !b) {
      results.push({ stem: job.stem, status: '跳过（缺 PNG，先转图）' })
      continue
    }
    const w = Math.min(a.width, b.width)
    const h = Math.min(a.height, b.height)
    const diff = new PNG({ width: w, height: h })
    const bad = pixelmatch(
      crop(a, w, h).data, crop(b, w, h).data, diff.data,
      w, h, { threshold: 0.12 }
    )
    const pct = ((bad / (w * h)) * 100).toFixed(2)
    const out = path.join(REPORTS, job.stem + '.diff.png')
    fs.writeFileSync(out, PNG.sync.write(diff))
    results.push({ stem: job.stem, pct: pct + '%', status: pct < 2 ? '✅ 优秀' : pct < 5 ? '🟡 可用' : '🔴 需修', diff: out })
  }

  console.log('\n===== WPS 兼容性对比报告 =====')
  for (const r of results) {
    console.log(`${r.status}  ${r.stem}  差异=${r.pct ?? '-'}${r.diff ? '  差异图: ' + path.basename(r.diff) : ''}  ${r.status.startsWith('跳过') ? r.status : ''}`)
  }
  const scored = results.filter((r) => r.pct)
  if (scored.length) {
    const avg = (scored.reduce((s, r) => s + parseFloat(r.pct), 0) / scored.length).toFixed(2)
    console.log(`\n平均差异: ${avg}%   （<2% 优秀 / <5% 可用 / 再高建议修）`)
  }
}

function loadPng(p) {
  try {
    if (!/\.png$/i.test(p)) return null
    return PNG.sync.read(fs.readFileSync(p))
  } catch {
    return null
  }
}
function crop(png, w, h) {
  if (png.width === w && png.height === h) return png
  const out = new PNG({ width: w, height: h })
  PNG.bitblt(png, out, 0, 0, w, h, 0, 0)
  return out
}
main().catch((e) => { console.error(e); process.exit(1) })
