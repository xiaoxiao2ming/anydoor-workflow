#!/usr/bin/env node
/**
 * 生成「可直接部署」的网页端目录（2026-10-17 网页端同步发布）
 *
 * 产物 release/web/ 是自包含的：只需要 Node 18+，**不需要 npm install**
 * （app-server 只用 node: 内置模块；OnlyOffice 分发在网页版被 APP_DISABLE_OO=1 关掉）。
 *
 * 目录结构：
 *   release/web/
 *     app-server.mjs              静态托管 + /api/pages + /audit-log + /healthz
 *     scripts/oo-offline-service.mjs、oo-core.mjs   （APP_DISABLE_OO=1 时不会被加载，留着只为桌面复用）
 *     dist-web/                   前端产物（index.html + assets + version.json）
 *     package.json                仅声明 type=module 与 start 脚本
 *     README.md                   部署说明（环境变量、探针、数据目录）
 *
 * 用法：
 *   node scripts/build-web.mjs && node scripts/publish-web.mjs
 *   npm run publish:web          # 上面两步
 *   npm run publish:web -- --zip # 额外打一个 zip，便于上传到任意主机
 */
import { mkdirSync, rmSync, cpSync, writeFileSync, existsSync, statSync, readdirSync, renameSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { spawnSync } from 'node:child_process'
import { readVersion, buildStamp } from './app-version.mjs'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const OUT = join(ROOT, 'release', 'web')
const wantZip = process.argv.includes('--zip')
const version = readVersion()
const builtAt = buildStamp()

if (!existsSync(join(ROOT, 'dist-web', 'index.html'))) {
  console.error('[publish:web] ❌ dist-web 还没构建，先跑 npm run build:web')
  process.exit(1)
}

/* ── ① 清空输出目录（只动 release/web 这一块，不碰源码）──
 * ⚠ release/web/server-data 是「在这个目录里直接跑起来」时产生的真实画布数据，
 *   重新发布不能把它删掉。先挪出去，清完再挪回来。 */
const dataDir = join(OUT, 'server-data')
const dataStash = join(ROOT, 'release', '.server-data.stash')
let stashed = false
try {
  if (existsSync(dataDir)) {
    rmSync(dataStash, { recursive: true, force: true })
    renameSync(dataDir, dataStash)
    stashed = true
  }
  rmSync(OUT, { recursive: true, force: true })
} catch (e) {
  console.warn('[publish:web] 清理 release/web 失败:', e && e.message)
}
mkdirSync(join(OUT, 'scripts'), { recursive: true })
if (stashed) {
  try {
    renameSync(dataStash, join(OUT, 'server-data'))
    console.log('[publish:web] 已保留现有画布数据 server-data/（重新发布不丢）')
  } catch (e) {
    console.warn('[publish:web] 恢复 server-data 失败，原数据在', dataStash, e && e.message)
  }
}

/* ── ② 复制最小运行集 ── */
cpSync(join(ROOT, 'dist-web'), join(OUT, 'dist-web'), { recursive: true })
cpSync(join(ROOT, 'app-server.mjs'), join(OUT, 'app-server.mjs'))
for (const f of ['oo-offline-service.mjs', 'oo-core.mjs', 'auth.mjs', 'oo-remote.mjs', '_blank_xlsx.b64', '_blank_docx.b64']) {
  const src = join(ROOT, 'scripts', f)
  if (existsSync(src)) cpSync(src, join(OUT, 'scripts', f))
}

/* ── ③ 运行期元数据 ── */
writeFileSync(
  join(OUT, 'package.json'),
  JSON.stringify(
    {
      name: 'client-base-web',
      version,
      private: true,
      type: 'module',
      description: 'client-base 网页端（画布 + 数据关联底座）',
      main: 'app-server.mjs',
      scripts: {
        start: 'node app-server.mjs',
        dev: 'node app-server.mjs',
      },
      // 零依赖：不要在这里加 dependencies，否则部署时又得装包
      dependencies: {},
      engines: { node: '>=18' },
      builtAt,
    },
    null,
    2,
  ) + '\n',
)

writeFileSync(
  join(OUT, 'README.md'),
  `# client-base 网页端 · v${version}

构建时间：${builtAt}

## 启动

\`\`\`bash
node app-server.mjs
\`\`\`

默认监听 \`127.0.0.1:5174\`；托管平台会注入 \`PORT\`，此时自动监听 \`0.0.0.0\`。

## 环境变量

| 变量 | 默认 | 说明 |
|---|---|---|
| \`PORT\` | 5174 | 监听端口（平台注入） |
| \`HOST\` | 有 PORT 时 0.0.0.0，否则 127.0.0.1 | 监听地址 |
| \`APP_DATA_DIR\` | \`./server-data/pages\` | 画布页数据目录（**必须可写、必须持久化**） |
| \`APP_OFFICE_DIR\` | \`./server-data/office\` | Office 中转目录 |
| \`APP_STATIC_DIR\` | \`./dist-web\` | 静态资源目录 |
| \`APP_DISABLE_OO\` | 未设 | \`1\` = 关闭 OnlyOffice 分发（网页部署请务必设 1） |
| \`APP_PASSWORD\` | 未设 | **访问口令**（2026-10-17）：设置后全站需登录（仿 DSM 登录页）；未设 = 关闭验证。公网/共享网络部署务必设置 |
| \`OFFICE_JWT_SECRET\` | 开发密钥 | OnlyOffice 签名密钥（网页版用不到） |

## 探针

- \`GET /healthz\` → \`{ ok, version, oo, dataDir, uptime }\`
- \`GET /version\` → 追加 node / platform

前端启动后会读 \`/healthz\`；\`oo:false\` 时 Office 节点降级为「只读提示」，
其余能力（画布、表格、参数、公式、.wfc 导入导出）全部可用。

## 数据

画布页写在 \`APP_DATA_DIR\`（默认 \`server-data/pages\`），与桌面端的
\`userData/client-base.sqlite\` **互相独立**，互不自动同步。
要备份就打包这个目录。
`,
)

/* ── ④ 体积报告 ── */
function dirSize(dir) {
  let total = 0
  let files = 0
  for (const e of readdirSync(dir, { withFileTypes: true })) {
    const p = join(dir, e.name)
    if (e.isDirectory()) {
      const s = dirSize(p)
      total += s.total
      files += s.files
    } else {
      total += statSync(p).size
      files++
    }
  }
  return { total, files }
}
const s = dirSize(OUT)
const mb = (n) => (n / 1024 / 1024).toFixed(2) + ' MB'
console.log(`\n[publish:web] ✅ 部署包已生成：${OUT}`)
console.log(`[publish:web] 体积 ${mb(s.total)}（${s.files} 个文件）  版本 v${version}`)
console.log('[publish:web] 本地试跑：cd release/web && node app-server.mjs')

/* ── ⑤ 群晖部署包 ────────────────────────────────────────────────
 * 目标：File Station 上传一个 tar.gz → 解压 → ssh 跑 install.sh 就能用。
 * 结构（顶层目录会被 --strip-components=1 剥掉）：
 *   synology-<ver>/
 *     app/                 ← 与 release/web 完全一致（只读挂进容器）
 *     docker-compose.yml   ← 挂载式，不 build
 *     Dockerfile.synology  ← 想做成单体镜像时的备选（app/ 里也放一份）
 *     install.sh           ← 建目录 / 灌默认画布 / 起容器
 *     .env.example
 *     README.md            ← 群晖部署手册
 */
const synoName = `synology-${version}`
const synoDir = join(ROOT, 'release', synoName)
const DEPLOY_DIR = join(ROOT, 'deploy', 'synology')
rmSync(synoDir, { recursive: true, force: true })
mkdirSync(synoDir, { recursive: true })
cpSync(OUT, join(synoDir, 'app'), { recursive: true })

const synoFiles = [
  ['docker-compose.yml', 'docker-compose.yml'],
  ['.env.example', '.env.example'],
  ['install.sh', 'install.sh'],
  ['README-synology.md', 'README.md'],
  ['Dockerfile.synology', 'Dockerfile.synology'],
]
for (const [src, dst] of synoFiles) {
  const p = join(DEPLOY_DIR, src)
  if (existsSync(p)) cpSync(p, join(synoDir, dst))
  else console.warn(`[publish:web] ⚠ 缺少群晖部署模板 ${src}`)
}
// 路线 C 在 app/ 里直接 build，所以这里再放一份
if (existsSync(join(DEPLOY_DIR, 'Dockerfile.synology'))) {
  cpSync(join(DEPLOY_DIR, 'Dockerfile.synology'), join(synoDir, 'app', 'Dockerfile.synology'))
}

/* 打 tar.gz：优先用系统 tar（Windows 自带 bsdtar、群晖是 GNU tar，互相能解）
 * ⚠ 必须 cwd 到 release/ 传相对文件名：Windows 路径带 `D:`，GNU tar 会当成远程主机
 *   （报 "Cannot connect to D: resolve failed"）。 */
const tarName = `client-base-synology-v${version}.tar.gz`
const tarPath = join(ROOT, 'release', tarName)
rmSync(tarPath, { force: true })
const tarRun = spawnSync('tar', ['-czf', tarName, synoName], {
  cwd: join(ROOT, 'release'),
  encoding: 'utf8',
})
if (tarRun.status === 0 && existsSync(tarPath)) {
  console.log(`[publish:web] 📦 群晖部署包：${tarPath}（${mb(statSync(tarPath).size)}）`)
  console.log('[publish:web]    群晖上用：sudo tar -xzf 该文件 --strip-components=1 -C /volume1/docker/client-base && sudo sh install.sh')
} else {
  console.warn('[publish:web] ⚠ tar 打包失败（可直接上传 release/' + synoName + '/ 整个目录）:', tarRun.stderr || tarRun.error?.message)
}

/* ── ⑥ 可选 zip ── */
if (wantZip) {
  const { default: JSZip } = await import('jszip')
  const { readFileSync } = await import('node:fs')
  const zip = new JSZip()
  const walk = (dir, prefix) => {
    for (const e of readdirSync(dir, { withFileTypes: true })) {
      const p = join(dir, e.name)
      if (e.isDirectory()) walk(p, prefix + e.name + '/')
      else zip.file(prefix + e.name, readFileSync(p))
    }
  }
  walk(OUT, '')
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } })
  const zipPath = join(ROOT, 'release', `client-base-web-v${version}.zip`)
  writeFileSync(zipPath, buf)
  console.log(`[publish:web] 📦 ${zipPath}（${mb(buf.length)}）`)
}
