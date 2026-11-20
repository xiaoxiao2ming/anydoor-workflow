// 集中存储后端：既托管 dist-web 静态资源，又提供 /api/pages 页面增删查改。
// 所有画布数据落盘到 server-data/pages/<id>.json，多浏览器/多设备访问同一份。
// 另提供 /office/* 端点支撑 OnlyOffice Document Server 集成：
//   GET  /office/files/:name   文档服务（DS 拉取 docx/xlsx/pptx）
//   POST /office/callback      DS 保存回调（status=2/6 时下载新文件写回）
//   POST /office/config        下发编辑器配置（含 JWT 签名）
import { createServer } from 'node:http'
import { handleOoRequest, handleOoUpgrade, setOoCorePath } from './scripts/oo-offline-service.mjs'
import { authEnabled, handleAuth, isAuthed } from './scripts/auth.mjs'
import { isOoRemote, handleOoRemote } from './scripts/oo-remote.mjs'
import { readFile, stat, writeFile, mkdir, unlink, readdir, appendFile, rename } from 'node:fs/promises'
import { readFileSync, statSync, watch, existsSync } from 'node:fs'
import { join, extname, normalize, dirname, resolve, basename } from 'node:path'
import { createHmac, timingSafeEqual } from 'node:crypto'
import { fileURLToPath } from 'node:url'

const __dirname = dirname(fileURLToPath(import.meta.url))

/* ★ 向上探测「真实的 scripts/oo-core.mjs」并告知 oo-offline-service（2026-10-06）
 *
 * 为什么需要：本文件在 C 端生产模式下会被 vite-plugin-electron **整个内联**进
 * dist-electron/main.js，于是 `__dirname`（= import.meta.url 的目录）变成 dist-electron/，
 * 而 oo-core.mjs 实际在 <项目根>/scripts/ 下。
 * oo-offline-service 若只按自己的 import.meta.url 推导就会找错文件，
 * 进而让**根路径 `/` 也返回 404**（OO 分发在静态托管之前，异常一路冒泡到最外层 catch）
 * ⇒ Electron 窗口永远停在 about:blank ⇒ C 端出包后打不开。
 *
 * 为什么用「向上探测」而不是直接 join(__dirname, '..', 'scripts')：
 * 源码直跑时 __dirname 就是项目根（join(__dirname,'scripts') 对），
 * 打包后 __dirname 是 dist-electron（需要再退一级）——
 * 两种布局只有「逐级向上找到真实存在的那个文件」才都成立。
 */
function locateOoCore () {
  let dir = __dirname
  for (let i = 0; i < 6; i++) {
    const cand = join(dir, 'scripts', 'oo-core.mjs')
    if (existsSync(cand)) return cand
    const parent = dirname(dir)
    if (parent === dir) break
    dir = parent
  }
  return join(__dirname, 'scripts', 'oo-core.mjs') // 兜底：保留原行为，报错信息里能看出路径
}
const OO_CORE_PATH = locateOoCore()
setOoCorePath(OO_CORE_PATH)
// Electron 主进程可注入覆盖（打包布局变化时的兜底）
if (process.env.OO_CORE_PATH) setOoCorePath(process.env.OO_CORE_PATH)
/* ★ 部署形态开关（2026-10-17 网页端同步发布）
 *   APP_DATA_DIR    数据目录（发布平台通常只给 /data 或 /tmp 可写，项目目录只读）
 *   APP_OFFICE_DIR  Office 中转目录（同上）
 *   APP_DISABLE_OO  1 = 完全不挂 OnlyOffice 分发。
 *     网页版部署（Linux 容器）里没有 DocService（Windows 二进制），
 *     让每个静态请求都先试一次 oo-core 加载既慢又刷错误日志，直接关掉；
 *     前端经 /healthz 得知 oo:false 后，把 OO 节点降级为「只读提示」。
 *   PORT / HOST     平台注入的监听参数（见 startAppServer）
 */
const OO_DISABLED = process.env.APP_DISABLE_OO === '1' || process.env.APP_DISABLE_OO === 'true'
const DEFAULT_DATA_DIR = process.env.APP_DATA_DIR
  ? resolve(process.env.APP_DATA_DIR)
  : join(__dirname, 'server-data', 'pages')
const DEFAULT_OFFICE_DIR = process.env.APP_OFFICE_DIR
  ? resolve(process.env.APP_OFFICE_DIR)
  : join(__dirname, 'server-data', 'office')

// 数据目录由 startAppServer 决定（同一进程内只有一个实例，模块级变量足够）
let DATA_DIR = DEFAULT_DATA_DIR
let OFFICE_DIR = DEFAULT_OFFICE_DIR
const indexFile = () => join(DATA_DIR, 'index.json')

// DS 与本服务之间的共享密钥（两边必须一致；生产环境用环境变量覆盖）
const OFFICE_JWT_SECRET = process.env.OFFICE_JWT_SECRET || 'client-base-office-dev-secret'

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8', // ★ 2026-10-23 pdf.js worker 模块（缺则 fetch dynamically imported module 失败）
  '.css': 'text/css; charset=utf-8',
  '.wasm': 'application/wasm',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.map': 'application/json; charset=utf-8',
  '.ico': 'image/x-icon',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  // office 三件套：DS 拉文档时的 Content-Type 必须正确，否则拒收
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
}

// office 工作目录只认这三类文件（防路径穿越 + 防误读任意文件）
const OFFICE_EXT = new Set(['.docx', '.xlsx', '.pptx'])

// ★ 实时控制台查看器（/devconsole）：Electron 主进程把渲染层（含 OO 编辑器 iframe 内部报错）
//   与主轴 console 全量归集到 .dev-console.log，这里用 SSE 实时推流 + 单页 HTML 查看器呈现，
//   解决 CDP 在沙箱里不稳定时看不到 iframe 内部报错的盲区。路径与 electron/main.ts 保持一致（cwd 根目录）。
const DEV_CONSOLE_LOG = resolve(process.cwd(), '.dev-console.log')
const DEV_CONSOLE_VIEWER = join(__dirname, 'scripts', 'devconsole-viewer.html')

// SSE 推流：先发文件末尾 300 行历史，再用 fs.watch 增量推送新增行（含心跳保活）
function streamDevConsole(req, res) {
  if (res.writableEnded) return
  res.writeHead(200, {
    'Content-Type': 'text/event-stream; charset=utf-8',
    'Cache-Control': 'no-cache, no-transform',
    Connection: 'keep-alive',
    'X-Accel-Buffering': 'no',
  })
  const send = (line) => { try { res.write(`data: ${line}\n\n`) } catch {} }
  try { writeFileSync(DEV_CONSOLE_LOG, '', { flag: 'a' }) } catch {}
  // 历史末尾
  try {
    if (existsSync(DEV_CONSOLE_LOG)) {
      const buf = readFileSync(DEV_CONSOLE_LOG, 'utf8')
      const lines = buf.split('\n').filter(Boolean).slice(-300)
      for (const l of lines) send(l)
    }
  } catch {}
  let offset = (() => { try { return statSync(DEV_CONSOLE_LOG).size } catch { return 0 } })()
  const hb = setInterval(() => { try { res.write(': ping\n\n') } catch {} }, 15000)
  let closed = false
  const onChange = () => {
    if (closed) return
    try {
      const size = statSync(DEV_CONSOLE_LOG).size
      if (size < offset) offset = 0 // 文件被轮转/截断，从头读
      if (size > offset) {
        const chunk = readFileSync(DEV_CONSOLE_LOG, 'utf8', { start: offset })
        offset = size
        const newLines = chunk.split('\n').filter(Boolean)
        for (const l of newLines) send(l)
      }
    } catch {}
  }
  let w
  try { w = watch(DEV_CONSOLE_LOG, onChange) } catch {}
  const cleanup = () => {
    closed = true
    clearInterval(hb)
    try { w && w.close() } catch {}
    try { res.end() } catch {}
  }
  req.on('close', cleanup)
  req.on('error', cleanup)
}

function sendDevConsole(req, res) {
  if (req.url && new URL(req.url, 'http://x').pathname === '/devconsole/stream') {
    streamDevConsole(req, res)
    return true
  }
  try {
    const html = readFileSync(DEV_CONSOLE_VIEWER, 'utf8')
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-cache' })
    res.end(html)
    return true
  } catch (e) {
    sendJson(res, 500, { error: 'devconsole viewer missing', detail: String(e && e.message || e) })
    return true
  }
}

function sendJson(res, code, obj) {
  if (res.writableEnded) return
  const body = JSON.stringify(obj)
  res.setHeader('Content-Type', 'application/json; charset=utf-8')
  res.setHeader('Cache-Control', 'no-cache')
  res.statusCode = code
  res.end(body)
}

async function readBody(req) {
  const chunks = []
  for await (const c of req) chunks.push(c)
  if (!chunks.length) return {}
  try {
    return JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}')
  } catch {
    return {}
  }
}

/* ── 操作审计日志落盘（2026-10-17）────────────────────────────────
 * 与 electron/main.ts 的 auditAppend 保持**完全一致**的落点与命名：
 *   logs/audit-YYYY-MM-DD.log （超过 8MB 滚到 .1）
 * 这样「桌面端」和「网页版」的日志可以直接拼在一起看，排查时不用先问环境。
 */
const AUDIT_MAX = 8 * 1024 * 1024
function auditDir () {
  return join(process.cwd(), 'logs')
}
function auditFilePath (d = new Date()) {
  const p = (n) => String(n).padStart(2, '0')
  return join(auditDir(), `audit-${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}.log`)
}
async function appendAuditLines (lines) {
  try {
    await mkdir(auditDir(), { recursive: true })
    const file = auditFilePath()
    try {
      const st = await stat(file)
      if (st.size > AUDIT_MAX) await rename(file, file + '.1').catch(() => writeFile(file, ''))
    } catch { /* 首次写 */ }
    await appendFile(file, lines.join('\n') + '\n', 'utf8')
  } catch (e) {
    console.error('[audit] 写日志失败:', (e && e.message) || e)
  }
}

/** 服务端事件也进审计日志：会话标记固定 `[srv]`，与浏览器的 `[sXXXXXX]` 区分 */
function authLog (line) {
  const d = new Date()
  const p = (n, w = 2) => String(n).padStart(w, '0')
  const stamp = `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ` +
    `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}.${p(d.getMilliseconds(), 3)}`
  void appendAuditLines([`${stamp} [srv] 鉴权 ${line}`])
}

async function ensureData() {
  await mkdir(DATA_DIR, { recursive: true })
  try {
    await stat(indexFile())
  } catch {
    await writeFile(indexFile(), '[]', 'utf8')
  }
}

async function loadIndex() {
  try {
    const raw = await readFile(indexFile(), 'utf8')
    const arr = JSON.parse(raw)
    return Array.isArray(arr) ? arr : []
  } catch {
    return []
  }
}

async function saveIndex(list) {
  await writeFile(indexFile(), JSON.stringify(list, null, 2), 'utf8')
}

function newId() {
  return 'page_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6)
}

async function handleApi(req, res, url) {
  const parts = url.pathname.split('/').filter(Boolean) // ['api','pages', ':id']
  const seg = parts[1] // 'pages'
  if (seg !== 'pages') return false

  // GET /api/pages -> 列表
  if (req.method === 'GET' && parts.length === 2) {
    const list = await loadIndex()
    sendJson(res, 200, list)
    return true
  }

  // POST /api/pages -> 新建（空画布）
  if (req.method === 'POST' && parts.length === 2) {
    const body = await readBody(req)
    const id = newId()
    const name = (body.name || `画布${await nextNum()}`).toString().slice(0, 60)
    const data = { id, name, tables: {}, records: {}, positions: {} }
    await writeFile(join(DATA_DIR, id + '.json'), JSON.stringify(data, null, 2), 'utf8')
    const list = await loadIndex()
    list.push({ id, name })
    await saveIndex(list)
    sendJson(res, 200, { id, name })
    return true
  }

  // 带 id 的操作
  const id = parts[2]
  if (!id) return false
  const file = join(DATA_DIR, id + '.json')

  // GET /api/pages/:id -> 完整数据
  if (req.method === 'GET') {
    try {
      const raw = await readFile(file, 'utf8')
      sendJson(res, 200, JSON.parse(raw))
    } catch {
      sendJson(res, 404, { error: 'not found' })
    }
    return true
  }

  // PUT /api/pages/:id -> 保存（含可选改名）
  if (req.method === 'PUT') {
    const body = await readBody(req)
    const name = typeof body.name === 'string' ? body.name.slice(0, 60) : undefined
    const data = {
      id,
      name: name ?? id,
      tables: body.tables ?? {},
      records: body.records ?? {},
      positions: body.positions ?? {},
    }
    await writeFile(file, JSON.stringify(data, null, 2), 'utf8')
    if (name !== undefined) {
      const list = await loadIndex()
      const it = list.find((p) => p.id === id)
      if (it) it.name = name
      else list.push({ id, name })
      await saveIndex(list)
    }
    sendJson(res, 200, { ok: true })
    return true
  }

  // DELETE /api/pages/:id
  if (req.method === 'DELETE') {
    try {
      await unlink(file)
    } catch {
      /* 文件不存在也当作成功 */
    }
    const list = await loadIndex()
    await saveIndex(list.filter((p) => p.id !== id))
    sendJson(res, 200, { ok: true })
    return true
  }
  return false
}

async function nextNum() {
  const list = await loadIndex()
  let max = 0
  for (const p of list) {
    const m = /画布(\d+)/.exec(p.name || '')
    if (m) max = Math.max(max, Number(m[1]))
  }
  return max + 1
}

/* ------------------------------------------------------------------ *
 * OnlyOffice 集成（DS 8+ 强制 JWT；这里零依赖实现 HS256，不引第三方库）
 * ------------------------------------------------------------------ */
const b64url = (buf) => Buffer.from(buf).toString('base64url')

function signJwt(payload, secret, expSec = 3600) {
  const header = b64url(JSON.stringify({ alg: 'HS256', typ: 'JWT' }))
  const body = b64url(JSON.stringify({ ...payload, exp: Math.floor(Date.now() / 1000) + expSec }))
  const sig = createHmac('sha256', secret).update(`${header}.${body}`).digest('base64url')
  return `${header}.${body}.${sig}`
}

function verifyJwt(token, secret) {
  if (typeof token !== 'string') return null
  const parts = token.split('.')
  if (parts.length !== 3) return null
  const expect = createHmac('sha256', secret).update(`${parts[0]}.${parts[1]}`).digest()
  const got = Buffer.from(parts[2], 'base64url')
  if (expect.length !== got.length || !timingSafeEqual(expect, got)) return null
  try {
    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'))
    if (typeof payload.exp === 'number' && payload.exp * 1000 < Date.now()) return null
    return payload
  } catch {
    return null
  }
}

/** 编辑器版本键：文件 mtime 变了 key 就变，DS 才会丢弃旧缓存 */
async function officeKey(file) {
  try {
    const s = await stat(join(OFFICE_DIR, file))
    return `${file}-${Math.round(s.mtimeMs)}`
  } catch {
    return `${file}-${Date.now()}`
  }
}

/**
 * /office/* 路由。
 * document.url 与 callbackUrl 用请求的 Host 头拼接 —— 浏览器怎么访问本服务，
 * DS 就怎么回访。C 端场景 DS 在 NAS、要回访电脑局域网 IP，所以必须用局域网 IP
 * 访问 app-server（或请求体传 hostOverride 显式指定）。
 */
async function handleOffice(req, res, url, hostHeader) {
  const parts = url.pathname.split('/').filter(Boolean) // ['office', ...]
  const sub = parts[1]

  // GET /office/files —— 工作目录里的模板清单（只列名与大小，不吐字节）
  // ★ 批次17 注：dev 模式下本文件的 /office/* 分支到不了这里（scripts/oo-core.mjs 的静态分发
  //   会先吞掉未匹配的路径并 return），所以前端实际走的是 /office-samples（见 oo-core.mjs）。
  //   此端点保留给独立 app-server 部署（electron 生产模式 / vite.web.config）使用。
  if (sub === 'files' && req.method === 'GET' && parts.length === 2) {
    try {
      const names = await readdir(OFFICE_DIR)
      const files = []
      for (const n of names) {
        const ext = extname(n).toLowerCase()
        if (!OFFICE_EXT.has(ext)) continue
        const st = await stat(join(OFFICE_DIR, n)).catch(() => null)
        if (st && st.isFile()) files.push({ name: n, ext: ext.slice(1), size: st.size, updatedAt: st.mtimeMs })
      }
      sendJson(res, 200, { dir: OFFICE_DIR, files })
    } catch (e) {
      sendJson(res, 500, { error: String(e && e.message ? e.message : e) })
    }
    return true
  }

  // GET /office/files/:name —— 文档服务
  if (sub === 'files' && req.method === 'GET') {
    const name = basename(decodeURIComponent(parts.slice(2).join('/')))
    const ext = extname(name).toLowerCase()
    if (!name || !OFFICE_EXT.has(ext)) return sendJson(res, 400, { error: 'bad file name' })
    try {
      const data = await readFile(join(OFFICE_DIR, name))
      res.setHeader('Content-Type', MIME[ext])
      res.setHeader('Cache-Control', 'no-cache')
      res.end(data)
    } catch {
      sendJson(res, 404, { error: 'file not found: ' + name })
    }
    return true
  }

  // POST /office/callback —— DS 保存回调。协议：status 2/6 带 url，需下载写回；
  // 任何情况都要回 {"error":0}，否则 DS 认为保存失败并重试。
  if (sub === 'callback' && req.method === 'POST') {
    const body = await readBody(req)
    // JWT 开启时真实负载藏在 token 里（payload 即整个 body 的镜像）
    const payload = body.token ? verifyJwt(body.token, OFFICE_JWT_SECRET) : body
    if (!payload) {
      console.warn('[office] callback JWT 校验失败')
      return sendJson(res, 200, { error: 1 })
    }
    const status = payload.status
    const file = payload.key ? String(payload.key).replace(/-\d+$/, '') : ''
    if ((status === 2 || status === 6) && typeof payload.url === 'string') {
      try {
        const r = await fetch(payload.url)
        if (!r.ok) throw new Error(`download ${r.status}`)
        const buf = Buffer.from(await r.arrayBuffer())
        await writeFile(join(OFFICE_DIR, file), buf)
        console.log(`[office] 保存回写成功: ${file} (${(buf.length / 1024).toFixed(1)}KB, status=${status})`)
      } catch (e) {
        console.error('[office] 保存回写失败', file, e instanceof Error ? e.message : e)
        return sendJson(res, 200, { error: 4 })
      }
    }
    return sendJson(res, 200, { error: 0 })
  }

  // POST /office/upload —— 前端把绑定产物推到工作目录（body: {name, dataBase64}）
  // 供后续「在线编辑」用。上传成功后 files/:name 即可被 DS 拉取。
  if (sub === 'upload' && req.method === 'POST') {
    const body = await readBody(req)
    const name = basename(String(body.name || ''))
    const ext = extname(name).toLowerCase()
    if (!name || !OFFICE_EXT.has(ext)) return sendJson(res, 400, { error: 'bad file name' })
    if (typeof body.dataBase64 !== 'string' || !body.dataBase64)
      return sendJson(res, 400, { error: 'missing dataBase64' })
    const buf = Buffer.from(body.dataBase64, 'base64')
    if (!buf.length) return sendJson(res, 400, { error: 'empty payload' })
    await writeFile(join(OFFICE_DIR, name), buf)
    console.log(`[office] 上传: ${name} (${(buf.length / 1024).toFixed(1)}KB)`)
    return sendJson(res, 200, { ok: true, size: buf.length })
  }

  // POST /office/config —— 下发编辑器配置（token 在服务端签，密钥不出服务端）
  if (sub === 'config' && req.method === 'POST') {
    const body = await readBody(req)
    const file = basename(String(body.file || ''))
    const ext = extname(file).toLowerCase()
    if (!file || !OFFICE_EXT.has(ext)) return sendJson(res, 400, { error: 'bad file name' })
    const dsUrl = String(body.dsUrl || '').replace(/\/+$/, '')
    if (!/^https?:\/\//.test(dsUrl)) return sendJson(res, 400, { error: 'bad dsUrl' })
    /* ★ 2026-10-24 iPad 接入：DS 地址跟随请求方 hostname（iPad 拿到 LAN IP 而非 127.0.0.1） */
    let host = String(body.hostOverride || hostHeader || `127.0.0.1`)
    host = host.replace(/^127\.0\.0\.1(:|$)/, (req.headers.host || '127.0.0.1:10215').split(':')[0] + '$1')
    const base = `http://${host}`
    const docType = ext === '.docx' ? 'word' : ext === '.xlsx' ? 'cell' : 'slide'
    const key = await officeKey(file)
    const config = {
      documentType: docType,
      document: {
        fileType: ext.slice(1),
        key,
        title: file,
        url: `${base}/office/files/${encodeURIComponent(file)}`,
        permissions: { edit: true, download: true, print: true },
      },
      editorConfig: {
        callbackUrl: `${base}/office/callback`,
        lang: 'zh-CN',
        customization: { autosave: true, compactHeader: true, hideRightMenu: true },
      },
    }
    config.token = signJwt(config, OFFICE_JWT_SECRET)
    return sendJson(res, 200, {
      apiJs: `${dsUrl}/web-apps/apps/api/documents/api.js`,
      config,
    })
  }

  sendJson(res, 404, { error: 'unknown office api' })
  return true
}

/**
 * 读部署版本（dist-web/version.json，由 scripts/build-web.mjs 生成）。
 * 读不到不让探针失败 —— 返回 unknown，前端照样能打开。
 */
async function readDeployVersion(root) {
  try {
    const raw = await readFile(join(root, 'version.json'), 'utf8')
    const obj = JSON.parse(raw)
    return { version: String(obj.version || 'unknown'), builtAt: obj.builtAt || null }
  } catch {
    return { version: 'unknown', builtAt: null }
  }
}

/**
 * 启动本地服务：提供 /api/pages 页面增删查改，可同时托管静态资源。
 * 三种用法：
 *   1) CLI：node app-server.mjs [rootDir] [port] [host]（参数省略时读 PORT/HOST 环境变量）
 *   2) Electron 主进程内嵌：startAppServer({ root: 静态目录|null, port: 0 })
 *   3) vite 开发代理：startAppServer({ root: null, port: 0 }) 只提供 API
 * @param {{root?: string|null, port?: number, host?: string, dataDir?: string, officeDir?: string, label?: string}} opts
 *   root    - 静态资源根目录；null 表示只提供 API（静态请求返回 404）
 *   port    - 监听端口；0 表示由系统分配空闲端口
 *   host    - 监听地址；默认 127.0.0.1，检测到 PORT 环境变量时默认 0.0.0.0（容器必需）
 *   dataDir - 页面数据目录，默认 APP_DATA_DIR 或项目内 server-data/pages
 * @returns {Promise<{port: number, dataDir: string, close: () => Promise<void>}>}
 */
export async function startAppServer(opts = {}) {
  const ROOT = opts.root === undefined
    ? (process.env.APP_STATIC_DIR ? resolve(process.env.APP_STATIC_DIR) : join(__dirname, 'dist-web'))
    : opts.root
  /* ★ 平台注入优先（2026-10-17 网页端发布）：
   *   主流托管平台（CloudBase / Railway / Fly / 自建容器）都用 PORT 环境变量指定端口，
   *   并且要求监听 0.0.0.0 才能被健康检查打进来。显式传参仍然优先（开发脚本/Electron 用）。 */
  const envPort = Number(process.env.PORT)
  const PORT = opts.port ?? (Number.isFinite(envPort) && envPort > 0 ? envPort : 5174)
  const HOST = opts.host ?? process.env.HOST ?? '127.0.0.1'
  if (opts.dataDir) DATA_DIR = opts.dataDir
  if (opts.officeDir) OFFICE_DIR = opts.officeDir
  const label = opts.label ?? 'app-server'

  const server = createServer(async (req, res) => {
    // ★ 客户端中断后异步写 → ERR_STREAM_WRITE_AFTER_END 的 'error' 事件无人接会崩进程
    //   （Electron 渲染层中断编辑器加载是常态），静默吞掉（2026-10-10 无头浏览器实测同崩）
    res.on('error', () => {})
    try {
      /* ★ 访问口令（2026-10-17）：路由器已把端口映射到公网，先把门关上。
       *   /healthz 与 /version 放行——容器 healthcheck 和前端能力探测都不带 cookie；
       *   其余一切（页面、静态资源、/api、/poc、/audit-log、/devconsole）都要求先登录。
       *   APP_PASSWORD 未设时 authEnabled()=false，本地开发/桌面端零打扰。 */
      if (authEnabled()) {
        /* ★ 2026-10-24 iPad 接入：本机（Electron 窗口）请求免认证——
           局域网设备（iPad）仍需配对码；否则窗口 loadURL 被登录页拦截误判失败 */
        const ra = req.socket.remoteAddress || ''
        const isLocal = ra === '127.0.0.1' || ra === '::1' || ra === '::ffff:127.0.0.1'
        const u = new URL(req.url || '/', 'http://x')
        if (!isLocal && u.pathname !== '/healthz' && u.pathname !== '/version') {
          const ver = await readDeployVersion(ROOT)
          const handled = await handleAuth({
            req, res, url: u,
            version: ver.version,
            host: String(req.headers.host || ''),
            log: authLog,
          })
          if (handled) return
        }
      }

      // ★ 实时控制台查看器：必须在 OO 分发之前拦截，避免被 handleOoRequest 误吞
      if (req.url && req.url.startsWith('/devconsole')) {
        if (sendDevConsole(req, res)) return
      }

      /* ★ 健康探针与版本（2026-10-17 网页端同步发布）
       *   /healthz 给托管平台做存活检查，同时把「这份部署具备什么能力」告诉前端：
       *     oo:false → 没有 OnlyOffice 文档服务（网页版常态），OO 节点降级为只读提示。
       *   必须放在 OO 分发之前：否则 handleOoRequest 会先吃掉这个路径。 */
      if (req.url === '/healthz' || req.url === '/version') {
        const url = new URL(req.url, 'http://x')
        const ver = await readDeployVersion(ROOT)
        /* 探针放行（容器 healthcheck 不带 cookie），但未登录时只回最小信息：
         * 不暴露数据目录、日志路径、Node/平台版本——这些都是攻击者的地图。 */
        const authed = isAuthed(req)
        sendJson(res, 200, {
          ok: true,
          service: 'client-base',
          version: ver.version,
          channel: 'web',
          builtAt: ver.builtAt || null,
          oo: !OO_DISABLED || isOoRemote(),
          ...(isOoRemote() ? { ooMode: 'remote' } : {}),
          ...(authed ? {
            dataDir: DATA_DIR,
            /* 审计日志文件路径也报出来：验证脚本/运维不用猜 cwd，
             * 部署包里数据目录与日志目录都可能被平台改写。 */
            logFile: auditFilePath(),
            uptime: Math.round(process.uptime()),
            auth: authEnabled() ? 'on' : 'off',
          } : {}),
          ...(authed && url.pathname === '/version' ? { node: process.version, platform: process.platform } : {}),
        })
        return
      }

      /* ★ OO 远程桥（2026-10-17）：群晖 DS（Linux 版 DocumentServer）在场时，
       *   用同一套宿主协议把 /poc、/office-offline/*、/doc/command 桥到远程 DS——
       *   网页版从此能在线编辑工作簿。必须排在本地 OO 分发之前。 */
      if (isOoRemote()) {
        const u = new URL(req.url || '/', 'http://x')
        const handled = await handleOoRemote({
          req, res, url: u,
          OFFICE_DIR,
          host: String(req.headers.host || ''),
          log: authLog,
        })
        if (handled) return
      }

      // ★ OnlyOffice 离线资产/桥/工作区（软件自带，无需手动启动 5199）
      //   注意：handleOoRequest 多数分支（/poc、/office-offline/* 等）同步写完 res 后 return 的是 undefined，
      //   而 /doc 反代是异步写 res。若仅凭「返回真值」判定，/doc 的异步 writeHead 会与后续静态处理器
      //   撞头 → ERR_HTTP_HEADERS_SENT 拖垮整个 dev 进程。故用「已返回真值 或 res 已写过」双重判定。
      //★★ 这里必须**就地兜住**加载失败，不能让异常冒泡：
      //   OO 分发在静态托管**之前**，异常一路冒到最外层 catch 会 404 掉**连根路径 `/` 在内**的所有请求
      //   ⇒ 页面根本加载不出来（实测 C 端生产模式窗口停在 about:blank）。
      //   OO 不可用只该影响 OO 自身的路径，不该拖垮整个站点 —— 故 catch 后继续往下走静态分支。
      if (!OO_DISABLED && req.url && !req.url.startsWith('/api/')) {
        try {
          const ooHandled = await handleOoRequest(req, res)
          if (ooHandled || res.writableEnded || res.headersSent) return
        } catch (e) {
          console.error('[oo] 分发失败，继续走静态分支:', (e && e.message) || e)
        }
      }

      if (req.url && req.url.startsWith('/api/')) {
        const url = new URL(req.url, 'http://x')
        const handled = await handleApi(req, res, url)
        if (handled) return
        sendJson(res, 404, { error: 'unknown api' })
        return
      }

      /* ★ 操作审计日志（2026-10-17 用户指令：每个点击/键盘/反馈都进日志）——
       * 浏览器/网页版没有 Electron IPC，改走同源 HTTP 落盘到同一个 logs/ 目录，
       * 两端的日志格式与文件名完全一致（audit-YYYY-MM-DD.log），排查时不用区分环境。
       * ⚠ readBody() 返回的**已是解析好的对象**（不是 Buffer），直接取 .lines。 */
      if (req.url === '/audit-log' && req.method === 'POST') {
        try {
          const obj = await readBody(req)
          const lines = Array.isArray(obj.lines) ? obj.lines.map(String) : []
          if (lines.length) await appendAuditLines(lines)
          sendJson(res, 200, { ok: true, file: auditFilePath() })
        } catch (e) {
          sendJson(res, 200, { ok: false, err: String(e && e.message ? e.message : e) })
        }
        return
      }

      if (req.url && req.url.startsWith('/office/')) {
        const url = new URL(req.url, 'http://x')
        await handleOffice(req, res, url, req.headers.host)
        return
      }

      // root 为 null 时只做 API 服务（dev 模式下页面由 vite 提供）
      if (ROOT === null) {
        sendJson(res, 404, { error: 'static hosting disabled' })
        return
      }

      let p = normalize(join(ROOT, '.' + new URL(req.url, 'http://x').pathname))
      let s
      try {
        s = await stat(p)
      } catch {
        p = join(ROOT, 'index.html')
        s = await stat(p)
      }
      if (s.isDirectory()) p = join(p, 'index.html')
      const data = await readFile(p)
      if (res.writableEnded) return
      res.setHeader('Content-Type', MIME[extname(p).toLowerCase()] ?? 'application/octet-stream')
      res.setHeader('Cache-Control', 'no-cache')
      res.end(data)
    } catch (e) {
      res.statusCode = 404
      res.end('not found: ' + (e instanceof Error ? e.message : String(e)))
    }
  })

  await ensureData()
  await mkdir(OFFICE_DIR, { recursive: true })
  await new Promise((r) => server.listen(PORT, HOST, r))
  // ★ OO 协作 ws 隧道：/doc/<key>/c/ 的 upgrade 直通 DocService(3798)——编辑器第二条连接全靠它
  server.on('upgrade', (req, socket, head) => {
    if (req.url && req.url.includes('/doc/')) handleOoUpgrade(req, socket, head)
    else socket.destroy()
  })
  const port = server.address().port
  const shownHost = HOST === '0.0.0.0' ? '0.0.0.0 (局域网可达)' : HOST
  console.log(`[${label}] data=${DATA_DIR} office=${OFFICE_DIR} static=${ROOT ?? '(disabled)'} -> http://${shownHost}:${port}/`)
  return {
    port,
    dataDir: DATA_DIR,
    close: () => new Promise((r) => server.close(r)),
  }
}

// 注意：这里刻意不用 top-level await —— 该文件会被 esbuild 打进 Electron 主进程，
// 某些 target 配置（如 chrome87）不支持 top-level await，会直接编译失败。
if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  // CLI 用法：node app-server.mjs [rootDir] [port] [host]
  //   ★ 参数省略时交由 startAppServer 读 PORT / HOST / APP_STATIC_DIR（发布平台注入）。
  //     原实现写死默认值，会把平台的 PORT 覆盖掉 —— 容器里表现为「起来了但探针打不通」。
  // host 传 0.0.0.0 时文档服务暴露到局域网 —— OnlyOffice DS（NAS/其他容器）才能回访
  const argPort = process.argv[3] ? Number(process.argv[3]) : undefined
  startAppServer({
    root: process.argv[2] ?? undefined,
    port: argPort,
    host: process.argv[4] ?? undefined,
  }).catch((e) => {
    console.error('[app-server] 启动失败', e)
    process.exit(1)
  })
}
