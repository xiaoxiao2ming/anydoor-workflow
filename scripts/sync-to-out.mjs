#!/usr/bin/env node
/* ★ 2026-10-25 用户指令「我要每次重启都是最新版」：
   把最新构建产物（渲染层 dist / 主进程 dist-electron / package.json）
   同步到打包目录 out/client-base/resources/app/ —— 秒级增量覆盖。
   双击启动器（bat）先跑本脚本再启动 exe，保证每次打开都是最新代码。 */
import { cpSync, existsSync, writeFileSync, readFileSync, mkdirSync, readdirSync, copyFileSync, statSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const OUT = process.argv.includes('--out')
  ? join(process.argv[process.argv.indexOf('--out') + 1], 'client-base')
  : join(root, 'out', 'client-base')
const APP = join(OUT, 'resources', 'app')

if (!existsSync(APP)) {
  console.error('[sync] 打包目录不存在（先跑一次完整打包）：', OUT)
  process.exit(1)
}

const t0 = Date.now()
// 渲染层 + 主进程 + 包描述（启动器依赖 package.json main 字段）
for (const dir of ['dist', 'dist-electron']) {
  const src = join(root, dir)
  if (existsSync(src)) cpSync(src, join(APP, dir), { recursive: true })
}
cpSync(join(root, 'package.json'), join(APP, 'package.json'))
// ★ 2026-10-26 19:14 实锤补漏：服务端脚本（app-server/oo-core 等内嵌服务链）也是
//   运行时代码——只同步 dist 的话，改 /pdfconv /docxconv 这类端点打包版永远旧代码
//   （症状=请求落静态托管返回 index.html，前端报「Unexpected token '<' is not valid JSON」）
cpSync(join(root, 'app-server.mjs'), join(APP, 'app-server.mjs'))
if (existsSync(join(root, 'scripts'))) cpSync(join(root, 'scripts'), join(APP, 'scripts'), { recursive: true })
// electron 专用入口文件（pack 脚本同款清单）
for (const f of ['electron/preload.cjs']) {
  const src = join(root, f)
  if (existsSync(src)) cpSync(src, join(APP, f))
}
// main 字段保险（pack 已写对，这里幂等兜底）
try {
  const pj = JSON.parse(readFileSync(join(APP, 'package.json'), 'utf8'))
  if (pj.main !== 'dist-electron/main.js') {
    pj.main = 'dist-electron/main.js'
    writeFileSync(join(APP, 'package.json'), JSON.stringify(pj, null, 2))
  }
} catch { /* ignore */ }

/* ★ 聪明模型运行时（2026-11-07）：node-llama-cpp 及其依赖闭包 + Qwen 模型文件。
   依赖清单动态递归收集（读各包 package.json dependencies）——不再手写，永不缺包。
   存在即跳过（模型 491MB，避免每次全量拷贝） */
const collectDeps = (rootName) => {
  const seen = new Set()
  const walk = (name) => {
    if (seen.has(name)) return
    const pjPath = join(root, 'node_modules', name, 'package.json')
    if (!existsSync(pjPath)) return
    seen.add(name)
    try {
      const pj = JSON.parse(readFileSync(pjPath, 'utf8'))
      for (const d of Object.keys(pj.dependencies || {})) walk(d)
    } catch { /* ignore */ }
  }
  walk(rootName)
  return [...seen]
}
const LLM_PKGS = collectDeps('node-llama-cpp')
let syncedPkgs = 0
for (const p of LLM_PKGS) {
  const src = join(root, 'node_modules', p)
  const dst = join(APP, 'node_modules', p)
  if (existsSync(src) && !existsSync(dst)) {
    cpSync(src, dst, { recursive: true })
    syncedPkgs++
  }
}
/* 模型文件（2026-11-07 修静默失败：此前 out 目录根本没带 models/，打包版聪明模型一直加载不到。
   逐文件判断——out 里缺哪个 .gguf 就拷哪个，已存在的不重拷（大文件 5.7GB 经不起整目录反复 cpSync）。 */
const modelsSrcDir = join(root, 'models')
const modelsDstDir = join(APP, 'models')
if (existsSync(modelsSrcDir)) {
  mkdirSync(modelsDstDir, { recursive: true })
  let copiedModels = 0
  for (const f of readdirSync(modelsSrcDir)) {
    if (!f.toLowerCase().endsWith('.gguf')) continue
    const s = join(modelsSrcDir, f)
    const d = join(modelsDstDir, f)
    if (!existsSync(d)) {
      copyFileSync(s, d)
      copiedModels++
      console.log(`[sync] 模型已拷贝：${f}（${(statSync(s).size / 1e9).toFixed(2)} GB）`)
    }
  }
} else {
  console.log('[sync] 警告：源目录没有 models/，聪明模型将不可用')
}

console.log('[sync] 最新代码已同步到打包目录（' + ((Date.now() - t0) / 1000).toFixed(1) + 's）' + (syncedPkgs ? `，LLM 运行时新装 ${syncedPkgs} 包` : ''))
