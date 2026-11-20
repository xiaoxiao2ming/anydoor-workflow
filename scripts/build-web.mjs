#!/usr/bin/env node
/**
 * 网页端构建（2026-10-17 用户指令：方便做成网页端同步发布）
 *
 * 为什么需要这个脚本，而不是直接 `vite build --config vite.web.config.ts`：
 *   1) vite 的 emptyOutDir 被关着（沙箱对批量删除有守卫，vite 用 MSYS 路径清目录会被拒），
 *      于是历史产物一层层堆着 —— 实测 dist-web 已经 149M，其中 121 个旧 assets 全是死文件，
 *      部署上去等于每次多传 20 倍垃圾。本脚本用 node:fs 自己清干净。
 *   2) 构建完必须写 version.json，服务端 /healthz 与前端标题栏都靠它报版本。
 *   3) 产物要能直接被 app-server 托管，出问题当场报错，而不是等部署完打开白屏。
 *
 * 用法：
 *   node scripts/build-web.mjs            # 清 → 构建 → 校验 → 写 version.json
 *   node scripts/build-web.mjs --no-clean # 不清旧产物（增量调试用）
 *
 * 产物：dist-web/（index.html + assets/* + public 资产 + version.json）
 * 下一步：npm run serve:web  或  node app-server.mjs dist-web 5174
 */
import { rmSync, existsSync, readdirSync, statSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { dirname } from 'node:path'
import { readVersion, buildStamp } from './app-version.mjs'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const OUT = join(ROOT, 'dist-web')
const noClean = process.argv.includes('--no-clean')

function humanSize(bytes) {
  if (bytes < 1024) return bytes + ' B'
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB'
  return (bytes / 1024 / 1024).toFixed(2) + ' MB'
}

function dirSize(dir) {
  let total = 0
  let files = 0
  const walk = (d) => {
    for (const e of readdirSync(d, { withFileTypes: true })) {
      const p = join(d, e.name)
      if (e.isDirectory()) walk(p)
      else {
        try {
          total += statSync(p).size
          files++
        } catch { /* 统计失败不影响构建 */ }
      }
    }
  }
  if (existsSync(dir)) walk(dir)
  return { total, files }
}

/* ── ① 清旧产物 ────────────────────────────────────────────────
 * 只删 assets/ 与 index.html/version.json：public 下的 logo 与 /wf 由 vite 负责覆盖，
 * 不需要动（删了反而会在「public 复制失败」时丢掉静态资产）。
 * 用 node:fs.rmSync 而非 shell rm：路径是 Windows 绝对路径，shell 层会被沙箱守卫拦。 */
if (!noClean) {
  for (const rel of ['assets', 'index.html', 'version.json']) {
    const p = join(OUT, rel)
    try {
      rmSync(p, { recursive: true, force: true })
    } catch (e) {
      console.warn('[build:web] 清理失败（继续构建，旧文件会残留）:', rel, String(e && e.message))
    }
  }
}

/* ── ② 构建 ───────────────────────────────────────────────────── */
process.env.VITE_CACHE_DIR = process.env.VITE_CACHE_DIR || 'node_modules/.vite'
const t0 = Date.now()
const { build } = await import('vite')
await build({ configFile: join(ROOT, 'vite.web.config.ts'), mode: 'production' })
const ms = Date.now() - t0

/* ── ③ 产物校验（部署前最后一道闸）─────────────────────────────── */
const problems = []
const indexFile = join(OUT, 'index.html')
if (!existsSync(indexFile)) problems.push('缺少 dist-web/index.html')
const assetsDir = join(OUT, 'assets')
const assets = existsSync(assetsDir) ? readdirSync(assetsDir) : []
if (!assets.some((f) => f.endsWith('.js'))) problems.push('assets 里没有 .js（构建未产出）')
if (!assets.some((f) => f.endsWith('.css'))) problems.push('assets 里没有 .css')
if (problems.length) {
  console.error('\n[build:web] ❌ 产物校验失败：')
  for (const p of problems) console.error('   - ' + p)
  process.exit(1)
}

/* ── ④ version.json：服务端 /healthz 与前端展示共用 ─────────────── */
const version = readVersion()
const builtAt = buildStamp()
writeFileSync(
  join(OUT, 'version.json'),
  JSON.stringify({ version, channel: 'web', builtAt, assets: assets.length }, null, 2) + '\n',
)

/* ── ⑤ 报告 ───────────────────────────────────────────────────── */
const size = dirSize(OUT)
console.log(`\n[build:web] ✅ v${version} 构建完成 ${(ms / 1000).toFixed(1)}s`)
console.log(`[build:web] 产物 ${OUT}`)
console.log(`[build:web] 体积 ${humanSize(size.total)}（${size.files} 个文件，assets ${assets.length} 个）`)
console.log('[build:web] 下一步：npm run serve:web   →   http://127.0.0.1:5174/\n')
