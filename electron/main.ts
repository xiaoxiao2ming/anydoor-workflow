import { app, BrowserWindow, ipcMain, dialog, session, Menu, protocol, net , shell, Tray, nativeImage, clipboard } from 'electron'
import { join, dirname, resolve } from 'path'
import { readFile, writeFile, mkdir, unlink, readdir } from 'fs/promises'
import { existsSync, appendFileSync, statSync, writeFileSync, readFileSync, readdirSync } from 'fs'
import { tmpdir } from 'node:os'
// node:http 是**默认导出**（不导出 { http }），写成 import { http } 会TS2305
import http from 'node:http'
import { fileURLToPath } from 'node:url'
import { pathToFileURL } from 'node:url'
import { spawn } from 'node:child_process'
import { startAppServer } from '../app-server.mjs'
import { ensureDocService } from '../scripts/oo-docservice.mjs'

/* 启动日志路径（2026-11-10 脱敏）：跟随 userData（每用户独立目录），不再硬编码本机绝对路径 */
const BOOT_LOG_PATH = (() => { try { return join(app.getPath('userData'), 'clientbase-boot.log') } catch { return 'clientbase-boot.log' } })()

// 清除该变量，保证 electron 以 GUI 模式运行（WorkBuddy 环境会带 ELECTRON_RUN_AS_NODE=1）
delete process.env.ELECTRON_RUN_AS_NODE

/* ★ 2026-10-25 闪退排查：全局异常全量落 C 盘 boot log（含异步拒绝）——
   此前异常只进 stdout（GUI 子系统收不到），崩溃完全不可见 */
process.on('uncaughtException', (e) => {
  try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' UNCAUGHT: ' + String(e?.stack || e).slice(0, 900) + String.fromCharCode(10)) } catch {}
})
process.on('unhandledRejection', (e) => {
  try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' UNHANDLED-REJ: ' + String((e as Error)?.stack || e).slice(0, 900) + String.fromCharCode(10)) } catch {}
})
/* ★ 退出链全家桶：exit 码 / quit 三事件 / 渲染层崩溃 / GPU 崩溃——
   「秒退」从此必有留痕（taskkill /F 硬杀除外，那种查不到 exit 码=9 就是） */
process.on('exit', (code) => {
  try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' process-exit code=' + code + String.fromCharCode(10)) } catch {}
})
const __quitMark = (tag: string) => {
  try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' ' + tag + String.fromCharCode(10)) } catch {}
}
app.on('before-quit', (e) => { __quitMark('before-quit'); (globalThis as { __isQuittingMark?: boolean }).__isQuittingMark = true })
app.on('will-quit', () => __quitMark('will-quit'))
app.on('quit', (_e, code) => __quitMark('app-quit code=' + code))
app.on('child-process-gone', (_e, details) => {
  __quitMark('child-process-gone ' + details.type + ' reason=' + details.reason + ' exit=' + details.exitCode)
  /* ★ GPU 进程崩 → 页面合成失败白/黑屏（窗体还在）——GPU 进程 Electron 会自动重启，
     但已挂页面不一定恢复：1.5 秒后 reload 所有窗口兜底（观察期：若崩溃频发，
     写 userData/gpu-accel.txt 内容 off 可关硬件加速） */
  if (details.type === 'GPU' && details.reason !== 'clean-exit') {
    setTimeout(() => {
      for (const w of BrowserWindow.getAllWindows()) {
        try { w.webContents.reload() } catch { /* ignore */ }
      }
      __quitMark('gpu-crash auto-reload')
    }, 1500)
  }
})

// 禁用代理：本机环境的系统代理会把 localhost 的 dev server 请求也拦掉，
// 导致窗口加载失败（白屏/不可见）。C 端只访问本机服务，不需要代理。
app.commandLine.appendSwitch('no-proxy-server')
// ★ 2026-10-24 用户指令「硬件加速开关」：偏好文件驱动（ready 前读——
//   disableHardwareAcceleration 必须在 app ready 前调用，重启生效）
const __gpuPrefFile = () => { try { return join(app.getPath('userData'), 'gpu-accel.txt') } catch { return 'gpu-accel.txt' } }  /* ★ userData 稳定路径（execPath 相对路径在打包/开发态漂移） */
const __gpuAccelOn = () => { try { return !existsSync(__gpuPrefFile()) || readFileSync(__gpuPrefFile(), 'utf8').trim() !== 'off' } catch { return true } }
if (!__gpuAccelOn()) { app.disableHardwareAcceleration(); app.commandLine.appendSwitch('disable-gpu') }
// OO 排障：每个实例固定开调试端口（CDP 连接用）+ 渲染层 console 转主进程 stdout
// ★ 端口/profile 可经环境变量覆盖（2026-11-09）：多实例调试时 9334 可能被僵尸 socket 占用，
//   第二实例 CB_USER_DATA_DIR=独立目录 + CB_CDP_PORT=9335 即可绕开
//   （requestSingleInstanceLock 按 userData 路径分键；setPath 必须在锁之前）。
const CB_ALT_PROFILE = process.env.CB_USER_DATA_DIR
if (CB_ALT_PROFILE) app.setPath('userData', CB_ALT_PROFILE)
app.commandLine.appendSwitch('remote-debugging-port', process.env.CB_CDP_PORT ?? '9334')

/* ── OnlyOffice 桌面构建私有协议：ascdesktop:// ─────────────────────────
 * vendors 里的 web-apps/sdkjs 是 DesktopEditors 桌面构建：
 *  - 字体文件按 AllFonts.js 里的绝对路径经 ascdesktop://fonts/<path> 加载
 *    （真实 DesktopEditors 由 CEF 壳拦截；纯浏览器无法响应 → 编辑器永远 Loading）
 *  - 这里在壳层接管该协议，按候选目录读本地字体/资源文件直出。 */
protocol.registerSchemesAsPrivileged([
  {
    scheme: 'ascdesktop',
    // standard+secure 必开：缺了 fetch(ascdesktop://...) 直接 TypeError: Failed to fetch
    privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true, bypassCSP: true },
  },
])

const OO_VENDOR_ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', 'vendors', 'onlyoffice')
const OO_FONT_DIRS = [
  'C:/Program Files/ONLYOFFICE/DesktopEditors/fonts/',
  'C:/Windows/Fonts/',
  OO_VENDOR_ROOT.replace(/\\/g, '/') + '/fonts/',
]

/** ascdesktop://fonts/<abs-path> → 读本地文件；找不到时按候选根目录兜底 */
function resolveAscdesktopFile(rawUrl: string): string | null {
  const rest = rawUrl.slice('ascdesktop://fonts/'.length).split('?')[0].split('#')[0]
  const cand = decodeURIComponent(rest).replace(/^\/+/, '')
  if (existsSync(cand)) return cand
  for (const dir of OO_FONT_DIRS) {
    const f = dir + cand.split('/').pop()
    if (existsSync(f)) return f
  }
  return null
}

const FONT_MIME: Record<string, string> = {
  '.ttf': 'font/ttf', '.otf': 'font/otf', '.ttc': 'font/collection', '.woff': 'font/woff', '.woff2': 'font/woff2',
}

// package.json 是 "type": "module"，编译产物为 ESM，没有 CommonJS 的 __dirname。
// 旧代码直接用 __dirname，在 ESM 下是 ReferenceError —— 桌面端因此一直起不来。
const HERE = dirname(fileURLToPath(import.meta.url))

// ★ 实时控制台归集：渲染层（含 OO 编辑器 iframe 内部报错）与主轴 console 全量落同一文件，
//   由 app-server 的 /devconsole（SSE 网页查看器）实时推送，解决 CDP 不可靠时的盲区。
const DEV_CONSOLE = resolve(process.cwd(), '.dev-console.log')
function devConsoleWrite (scope: string, level: string, msg: unknown, src?: string, line?: number) {
  try {
    const entry = JSON.stringify({
      ts: Date.now(), scope, level,
      msg: String(msg == null ? '' : msg).slice(0, 4000),
      src: src ? String(src).slice(-160) : '',
      line: line || 0,
    }) + '\n'
    try { if (statSync(DEV_CONSOLE).size > 6 * 1024 * 1024) writeFileSync(DEV_CONSOLE, readFileSync(DEV_CONSOLE, 'utf8').slice(-1500000)) } catch {}
    appendFileSync(DEV_CONSOLE, entry)
  } catch {}
}
// 主轴 console 也归集（跳过已在 console-message 里按 renderer scope 落盘的 [renderer:*] 行，避免重复）
{
  const FN = ['log', 'warn', 'error', 'info', 'debug'] as const
  type Fn = (typeof FN)[number]
  const _oc = Object.fromEntries(
    FN.map((f) => [f, console[f].bind(console)]),
  ) as Record<Fn, (...a: unknown[]) => void>
  for (const fn of FN) {
    const orig = _oc[fn]
    console[fn] = ((...a: unknown[]) => {
      orig(...a)
      const s = a
        .map((x) => {
          if (typeof x === 'string') return x
          if (x && typeof x === 'object' && 'stack' in x && x.stack) return String(x.stack)
          if (x && typeof x === 'object' && 'message' in x && x.message) return String(x.message)
          return JSON.stringify(x)
        })
        .join(' ')
      if (s.includes('[renderer:')) return
      devConsoleWrite('main', fn, s, '', 0)
    }) as typeof console.log
  }
}

/* ── 操作审计日志落盘（2026-10-17 用户指令）──────────────────────────
 * 「我对软件的每个点击、键盘输入、软件的反馈，都放进日志」。
 * 渲染层（src/core/audit.ts）负责采集与拼行，主进程只做「追加 + 轮转」：
 *   logs/audit-YYYY-MM-DD.log   当天文件，纯文本一行一条，grep 友好
 *   超过 8MB 时保留尾部 2000 行重写（与 scripts/oo-core.mjs 的日志轮转同一套做法）
 * 目录优先放在**进程工作目录**（开发时=项目根，用户一眼能找到）；
 * 不可写（安装目录只读）时退回 userData/logs。
 *
 * ★ 为什么目录创建是**异步**的：桌面端构建把 main.ts 与 app-server/oo-core 一起打进
 *   一个 ESM 文件，给 `fs` 增加新的具名导入会让打包器生成
 *   `import fs, { ... } from "node:fs"` 的混合形态，Electron 内置 Node 20 的
 *   CJS-ESM 互操作在这条路径上会抛 `Cannot read properties of undefined (reading 'exports')`
 *   → 进程直接起不来（实测）。所以这里只用**已导入**的 mkdir（fs/promises）异步建目录，
 *   运行期追加一律走已有的 appendFileSync。 */
let AUDIT_DIR = ''
async function ensureAuditDir(): Promise<string> {
  if (AUDIT_DIR) return AUDIT_DIR
  const cands = [resolve(process.cwd(), 'logs'), join(app.getPath('userData'), 'logs')]
  for (const d of cands) {
    try {
      await mkdir(d, { recursive: true })
      AUDIT_DIR = d
      return d
    } catch { /* 下一个候选 */ }
  }
  AUDIT_DIR = cands[cands.length - 1]
  return AUDIT_DIR
}
function auditFile(d = new Date()): string {
  const dir = AUDIT_DIR || join(process.cwd(), 'logs')
  const p = (n: number) => String(n).padStart(2, '0')
  return join(dir, `audit-${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}.log`)
}
const AUDIT_MAX = 8 * 1024 * 1024
/** 渲染层一批日志（已格式化的整行）→ 追加落盘；失败静默（日志不能反噬主流程） */
function auditAppend(lines: string[]): void {
  try {
    if (!Array.isArray(lines) || !lines.length) return
    if (!AUDIT_DIR) return   // 目录还没就绪（启动最初几毫秒）→ 丢弃，不阻塞主线程
    const file = auditFile()
    try {
      if (statSync(file).size > AUDIT_MAX) {
        // 轮转：只留尾部 2000 行（8MB 读一次，仅在超限时发生）
        const keep = readFileSync(file, 'utf8').split(/\r?\n/).slice(-2000).join('\n')
        writeFileSync(file, keep + '\n')
      }
    } catch { /* 首次写：文件不存在 */ }
    appendFileSync(file, lines.join('\n') + '\n')
  } catch (e) {
    lastAuditError = String(e)
  }
}
let lastAuditError = ''

function dbPath() {
  return join(app.getPath('userData'), 'client-base.sqlite')
}

/**
 * 画布页面数据目录：
 * - 开发时：项目内 server-data/pages（与网页版 npm run dev:web 共用同一份数据，切换不丢）
 * - 交付版：userData/pages（标准的桌面应用数据位置，不受安装目录权限影响）
 *
 * ★★ 判据不能用 `app.isPackaged`（2026-10-06 实测踩坑）：
 *   `isPackaged` 只有在**asar 打包**形态下才是 true；
 *   而本项目的交付形态是**免安装目录**（resources/app/ +复制过来的 electron.exe），
 *   这种形态下 `isPackaged === false` ⇒
 *   数据被写进 `resources/app/server-data/pages` —— 也就是**程序目录里**。
 *   后果：① 换版本/重装 ⇒ 用户画布数据连同旧目录一起丢；
 *        ② 程序放在 C:\Program Files 等受限目录 ⇒ 写入失败（EPERM/EACCES）。
 *   ⇒ 改为「能定位到项目根（有 src 与 node_modules）才算开发态」，
 *     否则一律用 userData。
 */
function isDevLayout (): boolean {
  // 开发态的判据：父目录里同时存在 src/ 与 node_modules/（打包产物里不会有 src）
  return existsSync(join(HERE, '..', 'src')) && existsSync(join(HERE, '..', 'node_modules'))
}

function pagesDir() {
  return isDevLayout()
    ? join(HERE, '..', 'server-data', 'pages')
    : join(app.getPath('userData'), 'pages')
}

/**
 * Word 模板库目录：userData/templates
 * 与 pagesDir 不同 —— 模板是二进制资产且体积较大，不参与网页版数据目录共享，
 * 桌面端与网页版各自独立（网页版走内存降级）。
 */
function templatesDir() {
  return join(app.getPath('userData'), 'templates')
}

/**
 * 打开失败的样本目录：userData/open-failures
 *
 * 每次「解析器读不了某个真实文件」都留一份（JSON 元数据 + 原始字节副本），
 * 用来回答「归一化层还需要补什么」。原始字节是必需的 —— 没有它就只能靠猜；
 * 而且这是用户本机、不是服务器，留副本没有外泄问题。
 */
function failuresDir() {
  return join(app.getPath('userData'), 'open-failures')
}

interface TemplateMeta {
  id: string
  name: string
  size: number
  updatedAt: number
  /** 落盘用的扩展名。老数据没有这个字段 → 读取时按候选逐个探测 */
  ext?: string
}

/**
 * 模板在磁盘上的文件名。
 *
 * 原先写死 `${id}.docx`，加 Excel 模板后这就错了 —— xlsx 字节躺在 .docx 文件里。
 * 功能上还能用（读取按 id），但用户在 userData/templates 里翻文件时会被误导。
 */
const TPL_EXTS = ['docx', 'xlsx', 'doc', 'xls', 'bin']

function extOf(name: string): string {
  const m = /\.([A-Za-z0-9]{1,5})$/.exec(name)
  const e = m?.[1]?.toLowerCase()
  return e && TPL_EXTS.includes(e) ? e : 'bin'
}

/** 按 index 里记的 ext 读；老数据没有 ext，就按候选顺序探测 */
async function readTemplateBytes(id: string, list: TemplateMeta[]): Promise<Uint8Array> {
  const meta = list.find((t) => t.id === id)
  const candidates = [meta?.ext, ...TPL_EXTS].filter(Boolean) as string[]
  for (const ext of candidates) {
    try {
      return new Uint8Array(await readFile(join(templatesDir(), `${id}.${ext}`)))
    } catch {
      /* 试下一个后缀 */
    }
  }
  throw new Error(`模板文件不存在：${id}`)
}

/** 删除该 id 的所有可能后缀文件（换扩展名后旧文件不会残留） */
async function unlinkTemplate(id: string): Promise<void> {
  for (const ext of new Set([...TPL_EXTS])) {
    try {
      await unlink(join(templatesDir(), `${id}.${ext}`))
    } catch {
      /* 不存在就是已经删掉了 */
    }
  }
}

async function readIndex(): Promise<TemplateMeta[]> {
  try {
    return JSON.parse(await readFile(join(templatesDir(), 'index.json'), 'utf8')) as TemplateMeta[]
  } catch {
    return []
  }
}

async function writeIndex(list: TemplateMeta[]) {
  await writeFile(join(templatesDir(), 'index.json'), JSON.stringify(list, null, 2), 'utf8')
}

// ★ 单实例锁：再次启动（双击快捷方式/bat）→ 聚焦已有窗口，服务不重启不闪断
/* ★ 终极印记（模块加载即写，绝对路径）：无此文件=electron 跑的不是本 main */
try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' module loaded from=' + String(import.meta.url).slice(-70) + String.fromCharCode(10)) } catch (eB) { try { console.error('[BOOT] write fail', eB) } catch {} }
/* ★ 2026-10-24：Tray 在 app ready 事件回调创建（官方要求 ready 后） */
app.on('ready', () => {
try {
  /* ★ 2026-10-26 用户指令「任务栏logo 托盘logo统一下」：托盘改用与窗口图标同一设计——
     首选 dist/tray.png（终稿 32px PNG，nativeImage 对 ICO 解码不可靠），
     候选 icons ico → build ico → 旧 logo-tile 兜底 */
  const trayCandidates = [
    join(HERE, '../dist/tray.png'),                  // 终稿 32px PNG（dev 与打包同路径，sync 同步）
    join(HERE, '../icons/client-base-app.ico'),      // 交付包 ico
    join(HERE, '../../build/client-base-app.ico')    // 开发态 ico
  ]
  let trayIcon = nativeImage.createFromPath(trayCandidates[0]).resize({ width: 16, height: 16 })
  for (let i = 1; i < trayCandidates.length && trayIcon.isEmpty(); i++) {
    trayIcon = nativeImage.createFromPath(trayCandidates[i]).resize({ width: 16, height: 16 })
  }
  if (trayIcon.isEmpty()) trayIcon = nativeImage.createFromPath(join(HERE, '../dist/logo-tile.png'))  // 兜底
  const tray = new Tray(trayIcon)
  tray.setToolTip('任意门')
  const trayMenu = Menu.buildFromTemplate([
    { label: '显示主窗口', click: () => { const w = BrowserWindow.getAllWindows()[0]; if (w) { w.show(); w.focus() } } },
    { label: '远程访问设置', click: () => showRemoteAccessDialog() },
    { type: 'separator' },
    { label: '重启应用', click: () => { isQuitting = true; app.relaunch(); app.exit(0) } },
    { label: '彻底退出', click: () => { isQuitting = true; app.quit() } },
  ])
  /* ★ 2026-10-24 用户指令「单击托盘图标 打开窗体」：Windows 上设置 contextMenu
     会屏蔽 click 事件——改双事件：左键=显窗，右键=弹菜单（popUpContextMenu） */
  const showMain = () => { const w = BrowserWindow.getAllWindows()[0]; if (w) { if (w.isMinimized()) w.restore(); w.show(); w.focus() } }
  tray.on('click', showMain)
  tray.on('right-click', () => tray.popUpContextMenu(trayMenu))
  try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' tray created' + String.fromCharCode(10)) } catch (eL) { console.error('[main] tray日志写失败:', eL) }
} catch (eT) {
  try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' tray FAILED: ' + String(eT).slice(0, 200) + String.fromCharCode(10)) } catch {}
}
})

const gotTheLock = app.requestSingleInstanceLock()
try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' lock-ok=' + gotTheLock + String.fromCharCode(10)) } catch {}
if (!gotTheLock) {
  /* ★ 2026-10-25 闪退根治：强杀（taskkill /F）后单例 mutex 短暂 abandoned——
     新实例获取失败即 quit=用户看到的「秒退」。自动 relaunch 重试最多 5 次
     （env 计数防无限循环），锁到手或重试耗尽为止 */
  const retries = parseInt(process.argv.find((a) => a.startsWith('--cb-lock-retry='))?.split('=')[1] ?? '0', 10)
  try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' lock-busy retry=' + retries + String.fromCharCode(10)) } catch {}
  if (retries < 5) {
    app.relaunch({ args: [...process.argv.slice(1).filter((a) => !a.startsWith('--cb-lock-retry=')), '--cb-lock-retry=' + (retries + 1)] })
    app.exit(0)
  } else {
    app.quit()
  }
} else {
  // ★★ 双击 .wfc 的核心链路（2026-10-07 用户：「双击*.wfc文件 打开软件并加载该文件内容」）：
  //   注册表把 .wfc 的 open 命令指到 `client-base.exe "%1"` ⇒ 文件路径进 **新实例的 argv**。
  //   两种情况都要接住：
  //   ① 软件没开：本实例自己的 argv 带路径（pendingWfc，渲染层 ready 后 invoke 取走）
  //   ② 软件已开：新实例因单实例锁退出，但 second-instance 事件把它的 argv 送进来
  //     ⇒ 转发已有窗口渲染层打开，用户体验 =「双击就切到那个画布」
  app.on('second-instance', (_e, argv) => {
    const w = BrowserWindow.getAllWindows()[0]
    if (!w) {
      /* ★ 2026-10-24 修复「打开软件没显示」：旧实例进程活着但窗口已销毁（异常路径）
         ——旧逻辑直接 return，用户看到的是「新实例退了、旧实例无窗」什么都没有。
         自愈：本实例 relaunch 重启（exit 释放单例锁 → 新实例干净建窗）。 */
      console.log('[main] 检测到旧实例窗口已丢失——自动重启自愈')
      try { appendFileSync(join(HERE, 'tray-debug.log'), new Date().toISOString() + ' self-heal relaunch (window lost)' + String.fromCharCode(10)) } catch {}
      isQuitting = true
      app.relaunch()
      app.exit(0)
      return
    }
    if (w.isMinimized()) w.restore()
    w.show(); w.focus()
    // ★★ 唤醒后**校验旧窗口是否真的加载成功**，失败就重载。
    //   实测（2026-10-06）：只要有一个实例的窗口停在加载失败状态
    //   （生产模式 app-server 端口未就绪 ⇒ loadURL 挂起 ⇒ 页面停在 about:blank），
    //   用户之后再双击多少次都只会聚焦到这个坏窗口 —— 表现为「C 端打不开，
    //   关掉重开也不行」，只能去任务管理器手动结束进程。
    //   ⇒ 这里判断加载状态，若异常则 reload()，让「双击」这个动作永远能救回来。
    //   isLoadingFinished 只表示「上一次导航结束」，故还要看 URL 是否还是 about:blank。
    try {
      const wc = w.webContents
      const blank = !wc.getURL() || /^about:/.test(wc.getURL())
      if (blank || wc.isCrashed?.()) {
        console.log('[main] 既有窗口未正常加载，执行 reload')
        wc.reload()
      }
    } catch (e) {
      console.log('[main] 既有窗口状态检查失败，尝试 reload', e)
      try { w.webContents.reload() } catch { /* ignore */ }
    }
    // 双击的 .wfc 路径 → 转给渲染层打开
    const wfc = findWfcArg(argv)
    if (wfc) {
      pendingWfc = wfc
      w.webContents.send('wfc:open', wfc)
    }
  })
}

/** 从 argv 里找 .wfc 文件路径（打包后 argv[0]=exe；dev 下前面还有 electron 路径） */
function findWfcArg (argv: string[]): string | null {
  for (let i = 1; i < argv.length; i++) {
    const a = argv[i]
    if (a && a.toLowerCase().endsWith('.wfc')) {
      try { return resolve(a) } catch { /* ignore */ }
    }
  }
  return null
}

/** 待渲染层来取的 .wfc 路径（首启双击场景；渲染层 ready 后 invoke 取走） */
let pendingWfc: string | null = null
// 首启参数里的 .wfc（窗口创建前先记下）
pendingWfc = findWfcArg(process.argv)

let isQuitting = false
/* ★ 远程访问信息（启动时赋值，托盘「远程访问设置」读取） */
let lanUrlInfo = ''
function showRemoteAccessDialog() {
  dialog.showMessageBox({
    type: 'info',
    title: '远程访问设置',
    message: 'iPad / 手机访问（同一局域网）',
    detail: `地址：${lanUrlInfo || '（启动后生成）'}
配对码：${pairCodeInfo || '（启动后生成）'}

在另一台设备的浏览器打开上方地址，输入配对码即可使用完整画布。
本机（Electron 窗口）无需配对码。`,
    buttons: ['好的'],
  })
}
let pairCodeInfo = ''
app.on('before-quit', () => { isQuitting = true })

// ★ 2026-10-09 01:34 撤销本体启动屏（f11bd9d）：用户明确「波浪展示要放在 OO 打开场景，
//   不是软件本体启动」——波浪+加载条已移至 /poc 页面（oo-core.mjs）的文档加载层。
//   本体启动只保留 backgroundColor=#0d1117（首帧即主题色，防系统白闪）。

async function createWindow() {
  try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' cw-enter' + String.fromCharCode(10)) } catch {}
  // 去掉默认应用菜单（File/Edit/View/...）—— C 端不该露出浏览器式菜单
  Menu.setApplicationMenu(null)
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    // 初始标题只作首帧占位：页面挂载后 App.tsx 会把 document.title 同步成
    // 「当前画布名」，Electron 原生标题栏跟随页面 title（除非 setTitle 覆盖）。
    title: '任意门',
    show: false,
    // ★ 2026-10-09：首帧底色 = 应用主题色（body 同款 #0d1117）。此前无此项，
    //   冷启动时窗口先亮出系统默认底色再变深蓝，观感突兀。
    backgroundColor: '#0d1117',
    // ★ 窗口/任务栏图标（2026-10-07）：水墨定稿 app ico（垫宣纸白砖）。
    //   dev 态取 build/；交付包取 resources/app/icons/（pack 自动拷）。
    //   app ico 缺失时退 shortcut ico（同为白砖+墨圈视觉，仅差角标）；再缺用默认图标不报错。
    icon: (() => {
      const candidates = [
        join(HERE, '..', 'icons', 'client-base-app.ico'),      // 交付包
        join(HERE, '..', 'icons', 'client-base-shortcut.ico'), // 交付包兜底
        join(HERE, '..', '..', 'build', 'client-base-app.ico') // 开发态
      ]
      for (const c of candidates) {
        try { if (existsSync(c)) return c } catch { /* ignore */ }
      }
      return undefined
    })(),
    webPreferences: {
      // ★ preload 用手写 .cjs（不受 type:module 影响）：vite-plugin-electron 0.28.8 输出 ESM
      //   会被 sandbox 以 CJS 加载 → "Cannot use import statement outside a module" → window.api 消失
      preload: join(HERE, '..', 'electron', 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      // ★ 2026-10-23 修「双击 PDF 打不开」：启用 PDFium 插件——Electron 默认禁用
      //   内置 PDF 查看器，iframe/blob URL 加载 PDF 只会空白（应用内 PDF 预览依赖它）
      plugins: true,
    },
  })

  /* ★ 2026-11-07 用户反馈「输入框里右键粘贴没反应」「画布上选文字复制不生效」：
     Electron 没有内建右键菜单，文本输入框右键什么都不会弹（浏览器里那套菜单不存在）。
     这里按上下文给最小菜单：
       · 可编辑区（textarea/input/可编辑单元格）→ 剪切 / 复制 / 粘贴 / 全选（原生 role，行为可靠）
       · 非可编辑但有选中文字 → 复制文字 / 引用到输入框（作为提示词）
       · 其余（画布节点/画布空白，页面自绘了节点菜单；对话气泡也已自绘）→ 不弹，避免两套菜单打架 */
  /* 渲染层自绘右键菜单时会先占位（ctx:handled）——本次 context-menu 不再弹原生菜单，
     避免「应用内菜单 + 系统菜单」两层叠在一起（对话气泡就是这种场景）。 */
  let ctxSkipUntil = 0
  ipcMain.on('ctx:handled', () => { ctxSkipUntil = Date.now() + 120 })

  win.webContents.on('context-menu', (_e, params) => {
    if (Date.now() < ctxSkipUntil) return
    console.log('[main] context-menu', JSON.stringify({ isEditable: params.isEditable, sel: params.selectionText.slice(0, 24), canPaste: params.editFlags.canPaste }))
    const tpl: Electron.MenuItemConstructorOptions[] = []
    const sel = params.selectionText.trim()
    if (params.isEditable) {
      if (sel) tpl.push({ role: 'cut', label: '剪切' })
      tpl.push({ role: 'copy', label: '复制', enabled: params.editFlags.canCopy || !!sel })
      tpl.push({ role: 'paste', label: '粘贴', enabled: params.editFlags.canPaste })
      tpl.push({ type: 'separator' })
      tpl.push({ role: 'selectAll', label: '全选' })
    } else if (sel) {
      /* ★ 2026-11-07 用户「沟通的文字选中后能复制，右键也有菜单，且右键能引用作为提示词工程的一部分」：
         只读文字（对话气泡 / 画布卡片正文 / 日志）右键给两项：
           · 复制文字——主进程直接写系统剪贴板。不用 role:'copy'：只读区没有焦点输入框，
             role 依赖 webContents.copy() 的聚焦语义，实测偶发「按钮点了但剪贴板没变」。
           · 引用到输入框——把选中原文发回渲染层，由对话面板插进提示词输入框（可再编辑/拼接）。
         两项都能拿到「用户当时选中的那段」，不会像气泡菜单那样跟 DOM 选区丢焦点赛跑。 */
      tpl.push({ label: '复制文字', click: () => { clipboard.writeText(params.selectionText) } })
      tpl.push({ label: '引用到输入框（作为提示词）', click: () => { win.webContents.send('chat:quote-text', params.selectionText) } })
    }
    if (!tpl.length) return
    Menu.buildFromTemplate(tpl).popup({ window: win })
  })

  // 菜单没了，刷新快捷键自己补（F5 / Ctrl+R）
  win.webContents.on('before-input-event', (e, input) => {
    const isReload = input.type === 'keyDown' && (input.key === 'F5' || (input.key.toLowerCase() === 'r' && input.control))
    if (isReload) {
      win.webContents.reload()
      e.preventDefault()
      // ★ 主进程拦截的按键不会进渲染层 → 审计日志在这里补一条（否则「我按了 F5 但日志没记录」）
      auditAppend([`${new Date().toLocaleString('zh-CN', { hour12: false })} [main] 键盘 刷新快捷键 | ${input.key}${input.control ? ' (Ctrl+' + input.key + ')' : ''} → 页面重载`])
    }
  })

  /* 窗口级操作审计（2026-10-17）：最大化/最小化/关闭这些动作发生在原生壳上，
   * 渲染层完全看不到，而它们恰恰是「软件没有响应」类问题的第一现场。 */
  const winAudit = (action: string, detail = '') => {
    const d = new Date()
    const p = (n: number, w = 2) => String(n).padStart(w, '0')
    const ts = `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}.${p(d.getMilliseconds(), 3)}`
    auditAppend([`${ts} [main] 窗口 ${action}${detail ? ' | ' + detail : ''}`])
  }
  for (const ev of ['maximize', 'unmaximize', 'minimize', 'restore', 'close', 'focus', 'blur', 'enter-full-screen', 'leave-full-screen']) {
    win.on(ev as Parameters<typeof win.on>[0], () => winAudit(ev))
  }
  // resize 拖动时每帧都触发 → 400ms 去抖，只记最终尺寸（否则一次拉伸几百行）
  let resizeTimer: ReturnType<typeof setTimeout> | null = null
  win.on('resize', () => {
    if (resizeTimer) clearTimeout(resizeTimer)
    resizeTimer = setTimeout(() => {
      try { const [w, h] = win.getSize(); winAudit('resize', `${w}x${h}`) } catch { /* ignore */ }
    }, 400)
  })

  // 窗口生命周期日志：排查"进程在但窗口不可见"
  console.log('[main] 窗口已创建')
  win.webContents.on('did-fail-load', (_e, code, desc, url) => {
    console.error('[main] 页面加载失败', code, desc, url)
  })
  win.webContents.on('render-process-gone', (_e, details) => {
    __quitMark('render-process-gone reason=' + details.reason + ' exit=' + details.exitCode)
    /* ★ 2026-10-25「窗体在内容空白」= 渲染进程崩（exit=-36861 实锤）——
       1 秒后自动重载救活（白屏时内存态已不可达，reload 是唯一出路；
       窗壳留着只会让用户以为软件死了）。循环崩也有 boot log 可查。 */
    if (details.reason === 'crashed' || details.reason === 'oom') {
      setTimeout(() => {
        try {
          if (!win.isDestroyed()) {
            win.webContents.reload()
            __quitMark('render-crash auto-reload')
          }
        } catch { /* 窗口已销毁 */ }
      }, 1000)
    }
  })
  win.webContents.on('did-finish-load', () => {
    console.log('[main] 页面加载完成')
  })
  // 渲染层 console 全量转发（OO 编辑器 iframe 的报错只有这里能看到）
  win.webContents.on('console-message', (_e, level, message, line, sourceId) => {
    const lvl = level >= 3 ? 'error' : level === 2 ? 'warn' : 'log'
    const tag = level >= 3 ? 'ERR' : 'log'
    // 全量落盘（不截断），供 /devconsole 查看器实时读取；同时保留 stdout 短行便于在终端瞄一眼
    devConsoleWrite('renderer', lvl, message, sourceId, line)
    console.log(`[renderer:${tag}] ${String(message).slice(0, 220)} @${String(sourceId).slice(-60)}:${line}`)
  })

  // ★ Electron 内 localhost 常被解析到 IPv6 ::1，而 vite 仅监听 127.0.0.1（IPv4），
  // 导致页面/模块脚本走 localhost 时连接挂死、React 永不渲染（dev 窗口空白、__store 缺失）。
  // 统一改写成 127.0.0.1，与 vite 的 host: '127.0.0.1' 对应。
  const devUrl = (process.env.VITE_DEV_SERVER_URL || '').replace(/localhost/g, '127.0.0.1')
  console.log('[main] devUrl =', devUrl ?? '(无，走生产模式)')
  // 先显示窗口 —— 加载失败也不能让窗口"隐身"（首帧底色由 backgroundColor 提供）
  win.show()
  try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' win-shown' + String.fromCharCode(10)) } catch {}
  try {
    if (devUrl) {
      // 开发模式：页面与 HMR 由 vite dev server 提供，/api 由 vite.config.ts 代理到同进程 app-server
      // ★ 重试循环：vite-plugin-electron 在 vite 完全就绪前就拉起窗口，首载常 ERR_ABORTED →
      //   窗口停在 about:blank（用户看到"拒绝连接"）。失败重试直到挂上。
      for (let i = 0; i < 12; i++) {
        try {
          await win.loadURL(devUrl)
          console.log('[main] dev 页面加载完成')
          break
        } catch (e) {
          console.log(`[main] dev 首载失败（第 ${i + 1} 次）：${String(e).slice(0, 80)}，1s 后重试`)
          await new Promise((r) => setTimeout(r, 1000))
        }
      }
    } else {
      // 生产模式：主进程内嵌 app-server，同时托管 dist 静态资源与 /api
      // 端口传 0 由系统分配，避免与用户已开的网页版服务冲突
      /* ★ 2026-10-24 iPad/跨设备接入（用户指令「iPad 打开画布 做电脑上的内容」）：
         固定端口 10215 + 监听全网卡 + 配对码门禁（复用既有 auth 机制）——
         iPad 浏览器打开 http://<电脑IP>:10215 输码即用（Host-Client 分布式，
         OO/数据引擎全在电脑，iPad 只是界面）。 */
      /* ★ iPad 配对码：首次生成持久化（userData/pair-code.txt），电脑端随处可查 */
      const pairFile = join(app.getPath('userData'), 'pair-code.txt')
      if (!existsSync(pairFile)) {
        const code = String(Math.floor(100000 + Math.random() * 900000))
        writeFileSync(pairFile, code, 'utf8')
      }
      pairCodeInfo = readFileSync(pairFile, 'utf8').trim()
      const pairCode = pairCodeInfo
      process.env.APP_PASSWORD = pairCode
      try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' prod-branch' + String.fromCharCode(10)) } catch {}
      const lanIps: string[] = []
      try {
        const nis = require('node:os').networkInterfaces() as Record<string, { family: string; address: string; internal: boolean }[]>
        for (const list of Object.values(nis)) for (const ni of list ?? []) if (ni.family === 'IPv4' && !ni.internal) lanIps.push(ni.address)
      } catch { /* ignore */ }
      lanUrlInfo = lanIps.length ? 'http://' + lanIps[0] + ':10215/' : 'http://<本机IP>:10215/'
      const lanUrl = lanUrlInfo
      console.log('[main] 跨设备访问: ' + lanUrl + '  配对码: ' + pairCode)
      try { appendFileSync(join(HERE, 'tray-debug.log'), new Date().toISOString() + ' LAN ' + lanUrl + ' code ' + pairCode + String.fromCharCode(10)) } catch {}
      const { port } = await startAppServer({
        root: join(HERE, '..', 'dist'),
        port: 10215,
        host: '0.0.0.0',
        dataDir: pagesDir(),
        label: 'app-server(electron)',
      })
      try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' server-listened port=' + port + String.fromCharCode(10)) } catch {}
      const prodUrl = `http://127.0.0.1:${port}/`
      // ★★ 生产模式同样需要「就绪 + 重试」，不能裸 loadURL（2026-10-06 实测）。
      //   实测症状：app-server 已 listen（日志打出端口号）但 win.loadURL 仍挂住，
      //   窗口永远停在 about:blank —— 而单实例锁会把新启动都引导到这个坏窗口，
      //   用户表现为「双击没反应 / 关掉重开也不行」，只能手动结束进程。
      //   可能诱因：Electron 与本机 IPv6 偏好组合下 127.0.0.1 的连接偶发不建立。
      //   ⇒ 先主动探活（确认服务真的应答 HTTP），再带超时 + 重试地加载；
      //     每次失败都显式 reload，确保不会留下一个"看起来开着但空白"的窗口。
      let loaded = false
      for (let i = 0; i < 10; i++) {
        try {
          // 探活：/api/pages 能应答才算服务真的就绪（只 listen 不够）
          const ping = await new Promise<string>((resolve) => {
            const req = http.get(prodUrl + 'api/pages', (r: import('node:http').IncomingMessage) => {
              r.resume()
              resolve('HTTP ' + r.statusCode)
            })
            req.on('error', (e: Error) => resolve('ERR ' + e.message))
            req.setTimeout(3000, () => { req.destroy(); resolve('TIMEOUT') })
          })
          if (!/^HTTP 2/.test(ping)) throw new Error('探活未通过：' + ping)
          await win.loadURL(prodUrl)
          loaded = true
          try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' prod-loaded' + String.fromCharCode(10)) } catch {}
          console.log('[main] 生产页面加载完成, port=' + port + ' ping=' + ping)
          break
        } catch (e) {
          console.log(`[main] 生产首载失败（第 ${i + 1} 次）：${String(e).slice(0, 100)}，1s 后重试`)
          try { win.webContents.reload() } catch { /* ignore */ }
          await new Promise((r) => setTimeout(r, 1000))
        }
      }
      if (!loaded) {
        // 加载不出来时**别留一个空白窗口骗人** —— 明确告知，用户才知道要查什么
        console.error('[main] 生产页面多次加载仍失败，窗口可能为空白：' + prodUrl)
        try {
          await win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(
            '<meta charset="utf-8"><body style="font:14px/1.6 system-ui;padding:24px">' +
            '<h3>任意门加载失败</h3>' +
            '<p>本地服务地址：<code>' + prodUrl + '</code></p>' +
            '<p>常见原因：dist 缺失或端口被占用。可在项目目录执行 <code>npm run build</code> 后重试。</p>' +
            '</body>',
          ))
        } catch { /* ignore */ }
      }
    }
  } catch (e) {
    // 加载失败也要把窗口亮出来（错误页/空白），不能让窗口"隐身"
    console.error('[main] loadURL 失败（窗口仍显示）', e)
  }

  // ★ 关窗口 = 隐藏（OO 离线服务等继续运行）；彻底退出走 before-quit（任务栏右键/关机）
  win.on('close', (e) => {
    if (!isQuitting) { e.preventDefault(); win.hide() }
  })
  win.on('closed', () => {
    // 仅置空引用，窗口关闭由 window-all-closed 统一处理
    console.log('[main] 窗口已关闭')
  })
}

try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' all-registered' + String.fromCharCode(10)) } catch {}
app.whenReady().then(async () => {
  try { appendFileSync(BOOT_LOG_PATH, new Date().toISOString() + ' whenReady entered' + String.fromCharCode(10)) } catch {}
  /* ★ 启动印记（2026-10-24 托盘排查）：无条件首行写——无此文件=新 main 未跑 */
  try { appendFileSync(join(HERE, 'tray-debug.log'), new Date().toISOString() + ' whenReady entered' + String.fromCharCode(10)) } catch (eR) { console.error('[main] 印记写入失败:', eR) }
  // ★ 单实例锁失败（app.quit() 已排队）就绝不建窗 —— 修「双击出闪退窗」竞态（2026-10-07 实锤）：
  //   quit 与 whenReady 回调赛跑，败者实例仍建了窗 + 起了内嵌服务（日志见多个一次性端口），
  //   随后 quit 把进程带走 = 窗体闪一下消失。锁在 :212 已判过，这里兜住 ready 之后的路径。
  if (!gotTheLock) return
  /* ★ 操作审计日志目录先就绪（异步、不阻塞）：之后窗口事件与渲染层批量日志才能落盘。
   *   失败也不影响启动 —— 日志是旁路。 */
  void ensureAuditDir()
  // ★ DocService(3798) 幂等拉起（2026-10-07 排障根治）：OO 编辑器「标准服务器流」
  //   的 /doc 协作通道必跑依赖，死了编辑器永远 loading（cellDiag 打 editor=无 死循环）。
  //   9-30 手动起过没进启动链，重启机器即丢 → 每次主进程启动顺带 ensure。
  //   不 await：不阻塞建窗；detached 独立进程，主进程退出不连带杀。
  // ★ ensureDocService 门禁（2026-10-07 19:1x 实测修正）：**DS 必须活着**——
  //   编辑器「标准服务器流」auth→open 全走 /doc/，DS 死 = 502 = 永远 loading
  //   （19:17 用户现场实锤，比慢更糟）。DS 活但外部依赖（PG/Redis/RabbitMQ）缺失时
  //   每请求 ~1s 排队（healthcheck 1.0/2.0/3.0s 实锤）——「慢但能开」仍是当前最优态。
  //   OO_DS=0 可显式关闭；装好依赖后 1s 排队消失，速度与可用性两全。
  if (process.env.OO_DS !== '0') {
    ensureDocService().catch((e: unknown) => console.log('[main] DocService ensure 失败:', String(e).slice(0, 120)))
  }
  // ★ 接管 OnlyOffice 桌面构建的字体私有协议（ascdesktop://fonts/<abs-path>）
  protocol.handle('ascdesktop', (req) => {
    const f = resolveAscdesktopFile(req.url)
    console.log('[ascdesktop]', f ? 'HIT ' + f : 'MISS ' + req.url.slice(0, 110))
    if (!f) return new Response('font not found: ' + req.url.slice(0, 120), { status: 404 })
    return net.fetch(pathToFileURL(f).toString()).then((r) => {
      const ext = f.slice(f.lastIndexOf('.')).toLowerCase()
      const headers = new Headers(r.headers)
      headers.set('Content-Type', FONT_MIME[ext] ?? 'application/octet-stream')
      headers.set('Access-Control-Allow-Origin', '*')
      return new Response(r.body, { status: 200, headers })
    })
  })

  // ★ 应用内重启（2026-10-07）：relaunch 排队新进程 + exit 立即退当前。
  //   点 X 是隐藏（保 OO 离线服务活），此前没有任何真·重启入口；OO 服务起不来
  //   / 资源 404 这类「进程态」故障只能靠任务管理器杀。走这里 = 整个进程重生，
  //   内嵌服务随新进程一起重建。
  // ★ 2026-10-08 用户：「重启和关闭按钮做成能彻底退出」—— 重启=杀 DS(全退)再由
  //   新进程重建；关闭=杀 DS(全退)不拉起。共用 killDocService()。
  const killDocService = () => {
    try {
      let dir = HERE
      for (let i = 0; i < 10; i++) {
        const cand = join(dir, 'vendors', 'oo-docservice.pid')
        try { if (existsSync(cand)) { dir = dirname(cand); break } } catch { /* ignore */ }
        dir = join(dir, '..')
      }
      const pidFile = join(dir, 'vendors', 'oo-docservice.pid')
      if (existsSync(pidFile)) {
        const pid = Number(readFileSync(pidFile, 'utf8').trim())
        if (pid > 0) { try { process.kill(pid); console.log('[main] DS(' + pid + ') 已终止') } catch { /* 已死 */ } }
      }
    } catch { /* ignore */ }
  }
  ipcMain.handle('app:relaunch', () => {
    isQuitting = true
    killDocService()
    app.relaunch()
    app.exit(0)
    // ★ 2026-10-08 防僵尸兜底：退出流程若被某句柄/线程阻塞（实锤过：quit 后进程
    //   杀不死、占着单实例锁 → 后续启动全部闪退死锁），3 秒后硬退。
    setTimeout(() => process.exit(0), 3000)
  })

  // ★ 彻底退出（2026-10-07）：与点 X（隐藏）相对的真退出入口。
  //   isQuitting 先置位再 quit —— 绕过 close→hide 拦截（与任务栏右键关闭同一通路，
  //   优雅走完 before-quit → close → window-all-closed 全链）。
  // ★ 2026-10-08 用户：「点击关闭能让软件彻底关闭么？」—— DS 是 detached 独立进程
  //   不随 Electron 退出（此前点关闭后 DS 一直占着 3798）→ quit 前读 pid 文件终止 DS，
  //   应用+内嵌服务全部退出，才是真正的「彻底关闭」。
  ipcMain.handle('app:quit', () => {
    isQuitting = true
    killDocService()
    app.quit()
    setTimeout(() => process.exit(0), 3000)   // ★ 同上：防优雅退出挂死产生僵尸
  })

  /* ★ 本地聪明模型（models/ 目录 GGUF 离线，node-llama-cpp 推理，懒加载——首次点聪明模式才载模型）
     ★ 2026-11-07 用户指令「本地快速删掉，聪明模型保留 新增刚才的」：升到 Qwen3.5-9B（5.7GB）。
       按优先级挑最新：Qwen3.5-9B → qwen2.5-3b → 目录里任意 .gguf（文件缺失时给可读的报错，
       不再静默失败——此前 out 打包目录没带 models/，聪明模式一直在走规则兜底没人发现）。 */
  /* ★ 2026-11-10 用户指令「帮我装个更高端的大模型」：heavy 首选升 Qwen3.8-27B（unsloth 动态量化
     Q3_K_XL 13.2GB，16G 显存内单载可行——比 9B 高两档）；文件缺失时按序回退 9B/3B（existsSync 保护） */
  const LLM_MODEL_PRIORITY = ['Qwen3.8-27B-UD-Q3_K_XL.gguf', 'Qwen3.5-9B-Q4_K_M.gguf', 'qwen2.5-3b-instruct-q4_k_m.gguf']
  /* ★ preferred：用户在对话面板模型菜单里锁定的文件名；不传 = 按优先级自动挑
     ★ 2026-11-09 用户「能否做成优化混合模型」：tier 路由——light 优先找 3B（短输出快任务），
       找不到落回主力（质量优先于速度）；heavy/未指定按现有优先级。 */
  const pickLlmModel = (preferred?: string, tier?: 'light' | 'heavy'): { path: string; name: string } => {
    const dir = join(app.getAppPath(), 'models')
    if (preferred) {
      const p = join(dir, preferred)
      if (!existsSync(p)) throw new Error(`指定的模型不存在：${preferred}（目录 ${dir}）`)
      return { path: p, name: preferred }
    }
    if (tier === 'light') {
      const light = LLM_MODEL_PRIORITY.find((n) => /3b/i.test(n) && existsSync(join(dir, n)))
      if (light) return { path: join(dir, light), name: light }
    }
    for (const name of LLM_MODEL_PRIORITY) {
      const p = join(dir, name)
      if (existsSync(p)) return { path: p, name }
    }
    const any = existsSync(dir) ? readdirSync(dir).filter((f) => f.toLowerCase().endsWith('.gguf')).sort() : []
    if (any.length) return { path: join(dir, any[any.length - 1]), name: any[any.length - 1] }
    throw new Error(`模型文件缺失：${dir} 下没有任何 .gguf（应有 ${LLM_MODEL_PRIORITY[0]}）`)
  }
  /* ★ 2026-11-09 混合模型：会话池化——9B（重任务）与 3B（快任务）**双常驻**，
     按任务分发不再互踢（旧实现切换即 dispose 重载，一次卡几秒）。
     16GB 显存（RX 9070 XT）双载 9B-Q4+3B-Q4 绰绰有余；池上限 2，超限 LRU 逐出。 */
  type LlmEntry = { session: { prompt: (p: string, o?: unknown) => Promise<string>; resetChatHistory?: () => void }; contextRef: { context: { dispose: () => unknown } }; modelRef: { model: { dispose: () => unknown } }; at: number }
  /* ★ 释放顺序（2026-11-10 用户实测 4096 仍爆）：必须先 context.dispose()（KV 显存归 context）
     再 model.dispose()——只 dispose model 时 context 引用不释放 = 显存泄漏（旧模型「逐出」了显存还在） */
  const releaseEntry = (e: LlmEntry | undefined): void => {
    if (!e) return
    try { void e.contextRef.context.dispose() } catch { /* ignore */ }
    try { void e.modelRef.model.dispose() } catch { /* ignore */ }
  }
  const llmSessions = new Map<string, LlmEntry>()
  /* 串行加载链：防止双模型同时 load 造成显存瞬时峰值（前一个失败不阻塞后一个） */
  let loadChain: Promise<unknown> = Promise.resolve()
  const withLoadLock = async <T,>(fn: () => Promise<T>): Promise<T> => {
    const run = loadChain.then(fn, fn)
    loadChain = run.catch(() => {})
    return run
  }
  const ensureLlmSession = async (want?: string, tier?: 'light' | 'heavy'): Promise<{ prompt: (p: string, o?: unknown) => Promise<string>; resetChatHistory?: () => void }> => {
    const picked = pickLlmModel(want, tier)
    const hit = llmSessions.get(picked.name)
    if (hit) { hit.at = Date.now(); return hit.session }
    const entry = await withLoadLock(async () => {
      const again = llmSessions.get(picked.name)
      if (again) return again
      /* ★ 2026-11-10 显存预判（用户实测「context size 8192 too large」）：27B(13.2G) 入池后
         任意双载组合都会超 16G——准入前按模型文件体积估算，超预算（14.5G，给 KV/系统留量）
         从最旧开始逐出，直到放得下。Q4 量化权重显存 ≈ 文件体积 +10%。 */
      const modelGB = (name: string): number => {
        try { return statSync(join(app.getAppPath(), 'models', name)).size / 1e9 } catch { return 0 }
      }
      const VRAM_BUDGET = 14.5
      let poolGB = [...llmSessions.keys()].reduce((a, k) => a + modelGB(k), 0)
      while (llmSessions.size > 0 && poolGB + modelGB(picked.name) > VRAM_BUDGET) {
        const oldest = [...llmSessions.entries()].sort((a, b) => a[1].at - b[1].at)[0]
        if (!oldest) break
        releaseEntry(oldest[1])
        llmSessions.delete(oldest[0])
        poolGB = [...llmSessions.keys()].reduce((a, k) => a + modelGB(k), 0)
        console.log('[llm] 显存预算逐出：' + oldest[0])
      }
      console.log('[llm] 加载模型：' + picked.name)
      const { getLlama, LlamaChatSession } = await import('node-llama-cpp')
      const llama = await getLlama()
      const model = await llama.loadModel({ modelPath: picked.path })
      /* ★ contextSize 必须显式给：Qwen3.5 训练长度 262144，缺省 createContext 会按它分配 KV
         缓存（数 GB），直接把进程挤死（实测 SIGTERM）。8192 = 系统提示+意图清单+2048 回复绰绰有余。
         ★ 显存不足时降级 4096 重试一次（2026-11-10），再失败给可读指引。 */
      let context
      try {
        context = await model.createContext({ contextSize: 8192 })
      } catch {
        try {
          context = await model.createContext({ contextSize: 4096 })
        } catch {
          /* 两次都失败 = 显存被残留占用：给可读指引而不是裸错误 */
          throw new Error(`显存不足以加载 ${picked.name}（即使 4096 上下文）——显存可能被残留会话占用，请重启客户端释放后重试，或改选更小模型`)
        }
      }
      const session = new LlamaChatSession({ contextSequence: context.getSequence() })
      const e: LlmEntry = { session: session as unknown as LlmEntry['session'], contextRef: { context }, modelRef: { model }, at: Date.now() }
      llmSessions.set(picked.name, e)
      /* 池上限 2（9B+3B 正好）：第三个进来逐出最旧的 */
      if (llmSessions.size > 2) {
        const oldest = [...llmSessions.entries()].sort((a, b) => a[1].at - b[1].at)[0]
        if (oldest) {
          releaseEntry(oldest[1])
          llmSessions.delete(oldest[0])
        }
      }
      return e
    })
    return entry.session
  }
  const dropLlmSession = (want?: string, tier?: 'light' | 'heavy') => {
    try {
      const name = want ?? pickLlmModel(undefined, tier).name
      const e = llmSessions.get(name)
      if (e) { releaseEntry(e); llmSessions.delete(name) }
    } catch { /* pick 失败就算了 */ }
  }
  const stripThink = (s: string) => s.replace(/<think>[\s\S]*?<\/think>/g, '').replace(/<\/?think>/g, '').trim()
  /* ★ 硬关思考（2026-11-07 实测）：Qwen3.5 默认在 assistant 轮开头自动开 <think>，
     用 responsePrefix 预填空「空思考块」把它立刻闭合，模型就直接出正文——9B 实测首块 0.0s、总 6.7s。
     代价：思考段不再产生（要看的场合就别开 noThink）。 */
  const NO_THINK_PREFIX = '\n</think>\n\n'
  /* ★ 2026-11-09 温度分任务（用户「怎么优化 使模型更聪明」）：工具/路由 0.1 稳定，
     对话/总结 0.5 自然（0.1 闲聊有复读机感）；调用方经 opts.temperature 传入 */
  const llmOpts = (opts?: { noThink?: boolean; maxTokens?: number; temperature?: number }) => ({
    maxTokens: opts?.maxTokens ?? 2048,
    temperature: opts?.temperature ?? 0.1,
    ...(opts?.noThink ? { responsePrefix: NO_THINK_PREFIX } : {}),
  })
  ipcMain.handle('llm:chat', async (_e, prompt: string, opts?: { model?: string; noThink?: boolean; maxTokens?: number; temperature?: number; tier?: 'light' | 'heavy' }) => {
    try {
      const s = await ensureLlmSession(opts?.model, opts?.tier)
      try { s.resetChatHistory?.() } catch { /* ignore */ }
      /* ★ maxTokens 384 → 2048：Qwen3.5 是思考型模型（思考段无标签、由运行时自动剥），
         实测 512/1024 全耗在思考上出不了正文，2048 稳定出答案（~17s）。
         万一仍超限返回空，给可读提示而不是空字符串。 */
      const raw = await s.prompt(prompt, llmOpts(opts))
      let text = stripThink(raw)
      /* ★ 兜底：正文被思考吃空 → 用 noThink 自动重试一次（用户不必手动「再说一次」） */
      if (!text && !opts?.noThink) {
        /* 重试前也清历史：第一次 prompt 已把本轮问答写进会话，不清会叠加成两轮 */
        try { s.resetChatHistory?.() } catch { /* ignore */ }
        const raw2 = await s.prompt(prompt, llmOpts({ ...opts, noThink: true }))
        text = stripThink(raw2)
      }
      return { ok: true, text: text || '（模型思考超限没有正文——再说一次试试）' }
    } catch (err: unknown) {
      dropLlmSession(opts?.model, opts?.tier)
      return { ok: false, reason: String((err as Error)?.message || err).slice(0, 200) }
    }
  })
  /* ★ 流式对话（2026-11-07 用户「不要结果好了再出」「看不到思考过程」）：
     onResponseChunk 同时携带正文与思考段（chain of thought）——正文 kind:text、思考 kind:thought，
     边生成边经 llm:chunk 事件推给渲染端；返回值仍是完整正文（思考段不含在内）。 */
  ipcMain.handle('llm:chat-stream', async (e, prompt: string, opts?: { model?: string; streamId?: string; noThink?: boolean; maxTokens?: number; tier?: 'light' | 'heavy' }) => {
    try {
      const s = await ensureLlmSession(opts?.model, opts?.tier)
      try { s.resetChatHistory?.() } catch { /* ignore */ }
      const push = (kind: 'text' | 'thought', text: string) => {
        if (!text) return
        if (!e.sender.isDestroyed()) e.sender.send('llm:chunk', { streamId: opts?.streamId ?? '', kind, text })
      }
      /* noThink 时第一块就是预填空 '\n</think>\n\n'，剥掉再上屏，别让用户看到标签 */
      let dropPrefix = opts?.noThink === true
      const pushText = (t: string) => {
        if (dropPrefix) {
          const idx = t.indexOf('</think>')
          if (idx >= 0) { dropPrefix = false; push('text', t.slice(idx + 8)) ; return }
          if (!t.trim()) return
          dropPrefix = false
        }
        push('text', t)
      }
      const raw = await s.prompt(prompt, {
        ...llmOpts(opts),
        onResponseChunk: (chunk: { type?: string; segmentType?: string; text: string }) => {
          if (chunk.type === 'segment' && chunk.segmentType === 'thought') push('thought', chunk.text)
          else pushText(chunk.text)
        },
      })
      let text = stripThink(raw)
      /* ★ 兜底：正文被思考吃空（流式里一个字都没出）→ 用 noThink 自动重试一次 */
      if (!text && !opts?.noThink) {
        dropPrefix = true
        const raw2 = await s.prompt(prompt, {
          ...llmOpts({ ...opts, noThink: true }),
          onResponseChunk: (chunk: { type?: string; segmentType?: string; text: string }) => {
            if (chunk.type === 'segment' && chunk.segmentType === 'thought') push('thought', chunk.text)
            else pushText(chunk.text)
          },
        })
        text = stripThink(raw2)
      }
      return { ok: true, text: text || '（模型思考超限没有正文——再说一次试试）' }
    } catch (err: unknown) {
      dropLlmSession(opts?.model, opts?.tier)
      return { ok: false, reason: String((err as Error)?.message || err).slice(0, 200) }
    }
  })

  /* ★ 联网搜索（2026-11-07 用户「看不到联网引用记录」）：Bing 网页版抓取（免 API key，国内可达），
     返回 [{title,url,snippet}] 给渲染端喂给本地模型总结；引用来源由渲染端展示为链接列表。 */
  /* ★ 网址监控（2026-11-09 用户指令「增加一个网址监控节点…提取、记录、或者输出数据」）：
     主进程 fetch（绕渲染层 CORS，带桌面 UA 与 20s 超时）→ 提取：
     text=cheerio CSS 选择器（空=全正文截断 2000 字）、json=点路径取字段、keyword=关键词包含检测。
     能力集参照 changedetection.io（35.3k stars）：URL 列表/选择器提取/变化标记/历史/导出。 */
  ipcMain.handle('webmon:fetch', async (_e, req: { url: string; mode: 'text' | 'json' | 'keyword'; selector?: string; keyword?: string; jsonPath?: string }) => {
    try {
      const res = await fetch(req.url, {
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 client-base-webmon', 'Accept': '*/*' },
        signal: AbortSignal.timeout(20000),
        redirect: 'follow',
      })
      if (!res.ok) return { ok: false, error: `HTTP ${res.status}` }
      const body = await res.text()
      if (req.mode === 'json') {
        try {
          let data: unknown = JSON.parse(body)
          if (req.jsonPath) {
            for (const seg of req.jsonPath.split('.').filter(Boolean)) {
              if (data == null) break
              const key = /^\d+$/.test(seg) ? Number(seg) : seg
              data = (data as Record<string | number, unknown>)[key as never]
            }
          }
          return { ok: true, value: data == null ? '' : typeof data === 'string' ? data : JSON.stringify(data) }
        } catch { return { ok: false, error: '响应不是有效 JSON' } }
      }
      if (req.mode === 'keyword') {
        const kw = req.keyword ?? ''
        return { ok: true, value: kw && body.includes(kw) ? `出现「${kw}」` : `未出现「${kw}」` }
      }
      const { load } = await import('cheerio')
      const $ = load(body)
      let value = req.selector ? $(req.selector).first().text().replace(/\s+/g, ' ').trim() : $.text().replace(/\s+/g, ' ').trim()
      if (!req.selector && value.length > 2000) value = value.slice(0, 2000)
      if (req.selector && !value) return { ok: false, error: `选择器没匹配到内容：${req.selector}` }
      return { ok: true, value }
    } catch (err: unknown) {
      const msg = String((err as Error)?.message || err)
      return { ok: false, error: /abort|timeout/i.test(msg) ? '请求超时（20s）' : msg.slice(0, 150) }
    }
  })

  ipcMain.handle('web:search', async (_e, query: string) => {
    try {
      const q = String(query || '').trim()
      if (!q) return { ok: false, reason: '搜索词为空' }
      const res = await fetch(`https://www.bing.com/search?q=${encodeURIComponent(q)}&count=8&setlang=zh-hans`, {
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36', 'Accept-Language': 'zh-CN,zh;q=0.9' },
        signal: AbortSignal.timeout(12000),
      })
      const html = await res.text()
      const items: Array<{ title: string; url: string; snippet: string }> = []
      const stripTags = (s: string) => s.replace(/<[^>]+>/g, '').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ').trim()
      const liRe = /<li class="b_algo"[\s\S]*?<\/li>/g
      let li: RegExpExecArray | null
      while ((li = liRe.exec(html)) && items.length < 6) {
        const block = li[0]
        const a = block.match(/<h2[^>]*><a[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a><\/h2>/)
        if (!a) continue
        const url = a[1].startsWith('http') ? a[1] : 'https://www.bing.com' + a[1]
        const title = stripTags(a[2])
        if (!title) continue
        const p = block.match(/<p[^>]*>([\s\S]*?)<\/p>/)
        items.push({ title, url, snippet: p ? stripTags(p[1]).slice(0, 220) : '' })
      }
      if (!items.length) return { ok: false, reason: '搜索结果解析为空（Bing 页面结构可能变了）' }
      return { ok: true, items }
    } catch (err: unknown) {
      return { ok: false, reason: String((err as Error)?.message || err).slice(0, 160) }
    }
  })

  /* 引用来源可点开：系统默认浏览器 */
  ipcMain.handle('shell:openUrl', async (_e, url: string) => {
    if (!/^https?:\/\//i.test(url)) return { ok: false }
    await shell.openExternal(url)
    return { ok: true }
  })

  // 持久化 IPC：读写本地 .sqlite 文件（供将来 sql.js 落盘用）
  ipcMain.handle('db:read', async () => {
    try {
      const buf = await readFile(dbPath())
      return new Uint8Array(buf)
    } catch (e) {
      if ((e as NodeJS.ErrnoException).code === 'ENOENT') return null
      throw e
    }
  })
  ipcMain.handle('db:write', async (_e, data: Uint8Array) => {
    await mkdir(dirname(dbPath()), { recursive: true })
    await writeFile(dbPath(), Buffer.from(data))
  })

  /* ── 系统剪贴板写入（2026-11-07）：气泡右键「复制」走这里 ──
     navigator.clipboard.writeText 在自定义协议/无焦点时会静默失败（catch 里什么也没有），
     表现为「点了复制，Ctrl+V 还是上一次的内容（比如旧截图）」。主进程 clipboard 同步可靠。 */
  ipcMain.handle('clipboard:write', (_e, text: string) => {
    clipboard.writeText(String(text ?? ''))
    return true
  })

  /* ── UI 偏好持久化（2026-10-20 用户「保存最新设置的 背景、UI设置 不要每次打开还是默认」）──
     localStorage 按 源(协议+host+port) 隔离，而 app-server 端口每次启动随机（port:0）
     ⇒ 端口一变 localStorage 就是新源，设置全丢。改存 userData/ui-settings.json：
     preload 启动时 sendSync 同步取一份快照注入 window.api.prefs；渲染层 settings:set 增量保存。 */
  const settingsPath = () => join(app.getPath('userData'), 'ui-settings.json')
  ipcMain.on('settings:get-sync', (e) => {
    try { e.returnValue = JSON.parse(readFileSync(settingsPath(), 'utf8')) } catch { e.returnValue = {} }
  })
  ipcMain.handle('settings:set', async (_e, patch: Record<string, string | null>) => {
    try {
      let cur: Record<string, string> = {}
      try { cur = JSON.parse(readFileSync(settingsPath(), 'utf8')) } catch { /* 首次无文件 */ }
      for (const [k, v] of Object.entries(patch || {})) {
        if (v === null || v === undefined) delete cur[k]
        else cur[k] = String(v)
      }
      await mkdir(dirname(settingsPath()), { recursive: true })
      await writeFile(settingsPath(), JSON.stringify(cur, null, 2))
    } catch (err) { console.error('[settings] 写入失败', err) }
  })

  /* ── 操作审计日志（2026-10-17）────────────────────────────────
   * 渲染层批量提交已格式化的整行；这里只负责追加，不做任何解析。
   * 返回值带路径，宿主首次写入时 console 打一行，用户就知道日志在哪。 */
  ipcMain.handle('audit:write', async (_e, lines: string[]) => {
    await ensureAuditDir()
    auditAppend(lines)
    return { ok: !lastAuditError, file: auditFile(), err: lastAuditError }
  })
  ipcMain.handle('audit:path', async () => {
    await ensureAuditDir()
    return { dir: AUDIT_DIR, file: auditFile() }
  })

  /* ── .wfc 文件关联（2026-10-07：双击 .wfc → 打开软件并加载）──── */
  /** 渲染层 ready 后主动取启动参数里的 .wfc 路径；取走即清（避免下次误触发） */
  ipcMain.handle('wfc:get-launch', () => {
    const p = pendingWfc
    pendingWfc = null
    return p
  })
  /** 📂 系统对话框选 .wfc（filters 限定扩展名；取消返回 null） */
  ipcMain.handle('wfc:pick', async () => {
    const win = BrowserWindow.getFocusedWindow()
    const res = await dialog.showOpenDialog(win ?? ({} as never), {
      title: '打开 .wfc 工程文件',
      filters: [{ name: '工作流画布工程', extensions: ['wfc'] }],
      properties: ['openFile'],
    })
    if (res.canceled || !res.filePaths.length) return null
    return res.filePaths[0]
  })
  /** 读 .wfc 字节（★ 渲染层无磁盘读权限，文件在主进程侧读）。
   *  ⚠ 路径校验：只允许 .wfc 后缀，防止被诱导读取任意文件。 */
  ipcMain.handle('wfc:read', async (_e, path: string) => {
    if (!path || !path.toLowerCase().endsWith('.wfc')) {
      throw new Error('只允许打开 .wfc 文件')
    }
    const buf = await readFile(path)
    return new Uint8Array(buf)
  })

  /* ── Word 模板库：userData/templates ─────────────────────────── */

  ipcMain.handle('tpl:list', async () => readIndex())

  ipcMain.handle('tpl:save', async (_e, name: string, data: Uint8Array) => {
    const dir = templatesDir()
    await mkdir(dir, { recursive: true })
    const list = await readIndex()
    // 同名覆盖：保留原 id，避免历史绑定关系失效
    const hit = list.find((t) => t.name === name)
    const id = hit?.id ?? `t${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`
    const ext = extOf(name)
    // 换过扩展名时先清旧的，避免同一 id 留下两份
    if (hit && hit.ext !== ext) await unlinkTemplate(id)
    await writeFile(join(dir, `${id}.${ext}`), Buffer.from(data))
    const meta: TemplateMeta = { id, name, size: data.byteLength, updatedAt: Date.now(), ext }
    const next = hit ? list.map((t) => (t.id === id ? meta : t)) : [...list, meta]
    await writeIndex(next)
    return meta
  })

  ipcMain.handle('tpl:read', async (_e, id: string) => {
    return readTemplateBytes(id, await readIndex())
  })

  ipcMain.handle('tpl:delete', async (_e, id: string) => {
    await unlinkTemplate(id)
    await writeIndex((await readIndex()).filter((t) => t.id !== id))
  })

  /* ── 打开失败的样本：userData/open-failures ───────────────────── */

  ipcMain.handle('fail:save', async (_e, payload: Record<string, unknown>, bytes: Uint8Array | null) => {
    const dir = failuresDir()
    await mkdir(dir, { recursive: true })
    const at = typeof payload?.at === 'number' ? payload.at : Date.now()
    // 时间戳做名字前缀：目录里按时间排序即可看清「什么时候开始出问题」
    const stamp = new Date(at).toISOString().replace(/[:.]/g, '-')
    const safe = String(payload?.name ?? 'unnamed')
      .replace(/[<>:"|?*\u0000-\u001f\\/]/g, '_')
      .slice(0, 120)
    const stem = `${stamp}_${safe}`
    let bytesPath: string | null = null
    if (bytes && bytes.byteLength) {
      bytesPath = join(dir, `${stem}.xlsx`)
      await writeFile(bytesPath, Buffer.from(bytes))
    }
    await writeFile(join(dir, `${stem}.json`), JSON.stringify({ ...payload, bytesPath }, null, 2), 'utf8')
    return dir
  })

  ipcMain.handle('fail:list', async () => {
    const dir = failuresDir()
    try {
      const names = (await readdir(dir)).filter((n) => n.endsWith('.json'))
      const out: Record<string, unknown>[] = []
      for (const n of names) {
        try {
          out.push(JSON.parse(await readFile(join(dir, n), 'utf8')) as Record<string, unknown>)
        } catch {
          /* 单条坏了不影响其它 */
        }
      }
      return out.sort((a, b) => Number(b.at ?? 0) - Number(a.at ?? 0)).slice(0, 100)
    } catch {
      return []
    }
  })

  /**
   * 各文件类型的对话框配置。
   * 之前 docx 写死在这里，加了 Excel 节点后必须参数化，否则挑 .xlsx 会被过滤器挡掉。
   */
  const FILE_KINDS: Record<string, { pickTitle: string; saveTitle: string; label: string; ext: string }> = {
    docx: { pickTitle: '选择 Word 模板', saveTitle: '导出 Word 报告', label: 'Word 文档', ext: 'docx' },
    xlsx: { pickTitle: '选择 Excel 模板', saveTitle: '导出 Excel 报表', label: 'Excel 工作簿', ext: 'xlsx' },
    pdf: { pickTitle: '选择 PDF 文件', saveTitle: '另存 PDF', label: 'PDF 文件', ext: 'pdf' },
  }
  const kindOf = (k: unknown) => FILE_KINDS[String(k)] ?? FILE_KINDS.docx

  /** 打开系统文件对话框选一个 Office 模板，返回文件名字节（不落库，由渲染层决定是否保存） */
  ipcMain.handle('tpl:pick', async (_e, kind?: string) => {
    const f = kindOf(kind)
    const win = BrowserWindow.getFocusedWindow()
    const res = await dialog.showOpenDialog(win ?? ({} as never), {
      title: f.pickTitle,
      filters: [{ name: f.label, extensions: [f.ext] }],
      properties: ['openFile'],
    })
    if (res.canceled || !res.filePaths.length) return null
    const p = res.filePaths[0]
    const buf = await readFile(p)
    return { name: p.split(/[\\/]/).pop() ?? `template.${f.ext}`, data: new Uint8Array(buf) }
  })

  /** 硬件加速开关（2026-10-24）：偏好文件读写，重启生效 */
  ipcMain.handle('remote:showDialog', () => { showRemoteAccessDialog(); return true })
  /* ★ 2026-10-25 用户指令「先确定位置和名称 保存时直接放进去」：
     系统另存为对话框（选目录+文件名）+ 字节直写选定路径（不经浏览器下载通道） */
  ipcMain.handle('wfc:pickSavePath', async (_e, defaultName: string) => {
    const r = await dialog.showSaveDialog({
      title: '保存工程文件',
      defaultPath: defaultName,
      filters: [{ name: '画布工程', extensions: ['wfc'] }],
      properties: ['createDirectory', 'showOverwriteConfirmation'],
    })
    if (r.canceled || !r.filePath) return null
    return r.filePath
  })
  /* ★ 2026-10-25 用户指令「最后优化一次」：zip/解压移到主进程（fflate 在 Node 跑，
     比渲染层快且完全不冻结 UI）。桌面端保存/打开走这两桥，web 端走 fflate Worker。 */
  ipcMain.handle('project:nativePack', async (_e, files: { path: string; bytes: Uint8Array }[], dirs: string[], savePath?: string) => {
    const { zipSync } = await import('fflate')
    const zipped: Record<string, Uint8Array> = {}
    for (const dir of [...dirs].sort()) zipped[dir + '/'] = new Uint8Array(0)
    for (const f of files) zipped[f.path] = f.bytes
    const out = zipSync(zipped, { level: 1 })
    if (savePath) {
      const { writeFile: wf } = await import('fs/promises')
      await wf(savePath, out)
      return { bytes: null, savedTo: savePath }
    }
    return { bytes: out, savedTo: null }
  })
  ipcMain.handle('project:nativeUnpack', async (_e, zipBytes: Uint8Array) => {
    const { unzipSync } = await import('fflate')
    try {
      const files = unzipSync(zipBytes)
      return { ok: true, files }
    } catch (e) {
      return { ok: false, files: {}, reason: e instanceof Error ? e.message : String(e) }
    }
  })
  /* ★ 2026-10-25 文件夹节点（用户指令「作为读取位置和保存位置」）：
     选目录 / 列文件 / 读文件字节 三桥 */
  ipcMain.handle('folder:pickDir', async () => {
    const r = await dialog.showOpenDialog({ title: '选择文件夹', properties: ['openDirectory', 'createDirectory'] })
    if (r.canceled || !r.filePaths?.[0]) return null
    return r.filePaths[0]
  })
  ipcMain.handle('folder:list', async (_e, dir: string) => {
    try {
      const { readdir, stat } = await import('fs/promises')
      const names = await readdir(dir)
      const out = await Promise.all(names.map(async (name) => {
        try {
          const st = await stat(dir + '/' + name)
          return { name, size: st.size, isDir: st.isDirectory(), mtime: st.mtimeMs }
        } catch { return { name, size: 0, isDir: false, mtime: 0 } }
      }))
      return { ok: true, files: out.filter((f) => !f.isDir).sort((a, b) => b.mtime - a.mtime) }
    } catch (e) {
      return { ok: false, files: [], reason: e instanceof Error ? e.message : String(e) }
    }
  })
  ipcMain.handle('folder:readFile', async (_e, dir: string, name: string) => {
    try {
      const { readFile } = await import('fs/promises')
      const safe = name.replace(/^\.\.\//, '')
      const bytes = await readFile(dir + '/' + safe)
      return { ok: true, bytes: new Uint8Array(bytes) }
    } catch (e) {
      return { ok: false, bytes: new Uint8Array(0), reason: e instanceof Error ? e.message : String(e) }
    }
  })
  ipcMain.handle('wfc:writeFile', async (_e, filePath: string, bytes: Uint8Array) => {
    const { writeFile: wf, mkdir: mk, stat } = await import('fs/promises')
    /* ★ 2026-10-26 二修：目录预建 + 统一 Buffer + 回真实落盘字节数——
       IPC 传大 Uint8Array 时若只回 byteLength，写丢也看不出来 */
    await mk(dirname(filePath), { recursive: true })
    await wf(filePath, Buffer.from(bytes))
    const st = await stat(filePath)
    return { ok: true, size: st.size }
  })

  /* ── PDF转Word 节点：本机 PaddleOCR 引擎外挂通道（2026-10-26）──────────
     引擎=百度开源 PaddleOCR（python -m pip 装在本机），不进安装包。
     ★ 2026-10-26 用户指令「把百度的装入套件的目录下（可卸载）」：
       引擎装 userData/suites/pdf2docx-ocr/venv（独立虚拟环境），卸载=删目录。
     p2w:tmpdir —— 发一次性工作目录（中转 pdf/docx 字节用）
     p2w:exec  —— spawn 本机 python 跑检测/转换；convert 成功时回读产物字节。
     exe 解析顺序：套件内 venv python（装了就优先）→ python → py -3（Windows 装机两种形态兜底）。 */
  const suiteOcrDir = () => join(app.getPath('userData'), 'suites', 'pdf2docx-ocr')
  const suiteVenvPython = () => join(suiteOcrDir(), 'venv', 'Scripts', 'python.exe')
  /* 内嵌 Python（无系统 Python 时安装脚本自动下载的 python312/） */
  const suiteEmbPython = () => join(suiteOcrDir(), 'python312', 'python.exe')

  /* ── 长驻 OCR 引擎进程（2026-10-26 用户问「为什么这么慢」）──────────
     实测：一次性 spawn 每次转换都要 import paddle + 加载 6 个模型（冷 99s / 热 38s）。
     长驻进程加载一次管线，后续单页仅推理（~10s）。行式 JSON 协议（同 run_p2w.py --serve）。
     进程保活至应用退出；引擎崩了下次请求自动重拉。 */
  type P2wSrv = {
    proc: ReturnType<typeof spawn>
    lineBuf: string
    errBuf: string
    ready: boolean
    readyWait: ((ok: boolean) => void)[]
    pending: ((line: string) => void)[]
  }
  let p2wSrv: P2wSrv | null = null
  let p2wChain: Promise<string> = Promise.resolve('')
  /* 全部在册转换子进程（长驻 + 一次性回退）——退出时统一收割 */
  const p2wChildren = new Set<ReturnType<typeof spawn>>()
  const p2wTail = () => (p2wSrv ? p2wSrv.errBuf : '').split('\n').map((s) => s.trim()).filter(Boolean).slice(-3).join(' ｜ ')
  function p2wServeData (d: Buffer | string) {
    if (!p2wSrv) return
    p2wSrv.lineBuf += String(d)
    let i: number
    while ((i = p2wSrv.lineBuf.indexOf('\n')) >= 0) {
      const line = p2wSrv.lineBuf.slice(0, i).trim()
      p2wSrv.lineBuf = p2wSrv.lineBuf.slice(i + 1)
      /* ★ 就绪握手行不占应答槽位——否则首条应答会被 ready 行顶掉，请求永远等不到回话 */
      if (line === 'P2W_READY') {
        p2wSrv.ready = true
        __quitMark('[P2W-SERVE] ready')
        p2wSrv.readyWait.splice(0).forEach((f) => f(true))
        continue
      }
      const cb = p2wSrv.pending.shift()
      if (cb) cb(line)
    }
  }
  /* ★ stderr 必须消费：① 不读会写满管道缓冲区把子进程卡死（937s 超时那次的元凶之一）
     ② 死亡时把尾部几行当原因回传，别再只给「退出码 -1」 */
  function p2wServeErr (d: Buffer | string) {
    if (!p2wSrv) return
    p2wSrv.errBuf = (p2wSrv.errBuf + String(d)).slice(-4000)
  }
  function p2wServeDead (code?: number | null) {
    if (!p2wSrv) return
    const dead = p2wSrv
    p2wSrv = null
    const tail = dead.errBuf.split('\n').map((s) => s.trim()).filter(Boolean).slice(-3).join(' ｜ ')
    __quitMark('[P2W-SERVE] dead code=' + String(code) + ' tail=' + (tail || '(stderr 空)').slice(0, 400))
    dead.readyWait.splice(0).forEach((f) => f(false))
    dead.pending.forEach((cb) => cb(JSON.stringify({ ok: false, err: tail || ('引擎进程退出（退出码 ' + String(code ?? '?') + '）') })))
    dead.pending.length = 0
  }
  function ensureP2wServe (): ReturnType<typeof spawn> {
    if (p2wSrv && !p2wSrv.proc.killed && p2wSrv.proc.exitCode === null) return p2wSrv.proc
    const py = existsSync(suiteVenvPython()) ? suiteVenvPython() : suiteEmbPython()
    const script = join(HERE, '..', 'dist', 'pdf2word', 'run_p2w.py')
    const proc = spawn(py, [script, '--serve'], { windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] })
    const srv: P2wSrv = { proc, lineBuf: '', errBuf: '', ready: false, readyWait: [], pending: [] }
    proc.stdout?.on('data', p2wServeData)
    proc.stderr?.on('data', p2wServeErr)
    proc.stdin?.on('error', () => { /* 进程死后写 stdin 会 EPIPE，忽略即可 */ })
    proc.on('exit', (code) => p2wServeDead(code))
    proc.on('error', () => p2wServeDead(-1))
    p2wSrv = srv
    __quitMark('[P2W-SERVE] spawn py=' + py + ' pyExists=' + existsSync(py) + ' script=' + script + ' scriptExists=' + existsSync(script))
    return proc
  }
  function p2wServeReq (req: { input: string; output: string; lang: string; profile: string }, timeoutMs = 900000): Promise<{ ok: boolean; err?: string; pages?: number }> {
    const run = () => new Promise<string>((resolve) => {
      const srv = p2wSrv
      if (!srv) { resolve(JSON.stringify({ ok: false, err: '引擎未启动' })); return }
      let settled = false
      let bootTo: ReturnType<typeof setTimeout> | null = null
      const done = (line: string) => {
        if (settled) return
        settled = true
        if (bootTo) clearTimeout(bootTo)
        resolve(line)
      }
      /* 引擎冷启动（含 import paddle）通常 5-20s；超过 20s 还没 ready 就判定启动失败 */
      bootTo = setTimeout(() => done(JSON.stringify({ ok: false, err: '引擎启动超时（20s）: ' + (p2wTail() || '无 stderr 输出') })), 20000)
      const send = () => {
        if (settled) return
        if (bootTo) { clearTimeout(bootTo); bootTo = null }
        const to = setTimeout(() => {
          const idx = srv.pending.indexOf(cb)
          if (idx >= 0) srv.pending.splice(idx, 1)
          done(JSON.stringify({ ok: false, err: '转换超时（' + Math.round(timeoutMs / 1000) + 's）——引擎可能卡死，已重置，下次请求自动重拉' }))
          /* ★ 超时=引擎卡死（上次 937s 就是这样挂住的）：杀掉重拉，别留僵尸进程 */
          try { srv.proc.kill() } catch { /* ignore */ }
        }, timeoutMs)
        const cb = (line: string) => { clearTimeout(to); done(line) }
        srv.pending.push(cb)
        srv.proc.stdin?.write(JSON.stringify(req) + '\n')
      }
      if (srv.ready) send()
      else srv.readyWait.push((ok) => { if (ok) send(); else done(JSON.stringify({ ok: false, err: '引擎进程启动即退出: ' + (p2wTail() || '无 stderr 输出') })) })
    })
    p2wChain = p2wChain.then(run, run)
    return p2wChain.then((line) => { try { return JSON.parse(line) } catch { return { ok: false, err: '引擎应答异常: ' + line.slice(0, 120) } } })
  }
  app.on('will-quit', () => {
    try { p2wSrv?.proc.kill() } catch { /* ignore */ }
    /* 一次性转换子进程也要杀（否则重启后仍后台烧 CPU） */
    for (const c of p2wChildren) { try { c.kill() } catch { /* ignore */ } }
    p2wChildren.clear()
  })
  /* ★ 预热：启动 3s 后先把引擎进程拉起来（只 import 不推理，P2W_READY 即完成）——
     ① 首份 PDF 少等一次 python 冷启动 ② 启动失败在 boot log 立即可见 */
  setTimeout(() => {
    try {
      if (existsSync(suiteVenvPython()) || existsSync(suiteEmbPython())) ensureP2wServe()
    } catch { /* ignore */ }
  }, 3000)
  ipcMain.handle('p2w:tmpdir', async () => {
    const dir = join(tmpdir(), 'clientbase-p2w-' + Date.now().toString(36))
    await mkdir(dir, { recursive: true })
    return { dir }
  })
  ipcMain.handle('p2w:exec', async (_e, payload: {
    mode: 'detect' | 'convert'
    input?: string; output?: string; lang?: string; profile?: string; timeoutMs?: number
  }) => {
    const script = join(HERE, '..', 'dist', 'pdf2word', 'run_p2w.py')
    /* 速度档位（2026-10-26 用户问「有没有容差设置」）：pro 精准（缺省）/ std 均衡 / fast 极速 */
    const profile = ['fast', 'std', 'pro'].includes(payload.profile ?? '') ? (payload.profile as string) : 'pro'
    const pyArgs = payload.mode === 'detect'
      ? ['-c', 'import paddleocr,sys;print("PADDLEOCR_OK",getattr(paddleocr,"__version__","?"))']
      : [script, payload.input ?? '', payload.output ?? '', '--lang', payload.lang ?? 'ch', '--profile', profile]
    const tryRun = (exe: string, args: string[]) => new Promise<{ ok: boolean; code: number; stdout: string; stderr: string; exe: string }>((res) => {
      let p
      try {
        p = spawn(exe, args, { windowsHide: true })
      } catch (err) {
        res({ ok: false, code: -1, stdout: '', stderr: String(err), exe })
        return
      }
      let out = ''
      let errTxt = ''
      /* ★ 登记：应用退出时一并杀掉（2026-10-26 教训：回退的一次性转换 2.1GB 后台跑 40 分钟，
         重启 app 杀不掉它——转换子进程必须统一登记在册） */
      p2wChildren.add(p)
      p.on('close', () => p2wChildren.delete(p))
      p.stdout?.on('data', (d) => { out += String(d) })
      p.stderr?.on('data', (d) => { errTxt += String(d) })
      p.on('error', (err) => { errTxt += String(err) })
      p.on('close', (code) => res({ ok: code === 0, code: code ?? -1, stdout: out, stderr: errTxt, exe }))
    })
    /* convert 模式：转换成功后回读产物字节（所有 exe 候选共用——否则套件内分支漏读=产物为空） */
    const finish = async (r: { ok: boolean; code: number; stdout: string; stderr: string; exe: string }) => {
      if (payload.mode === 'convert' && r.ok && payload.output) {
        try {
          const buf = await readFile(payload.output)
          return { ...r, bytes: new Uint8Array(buf) }
        } catch (err) {
          return { ...r, ok: false, stderr: r.stderr + '\n产物读取失败: ' + String(err) }
        }
      }
      return r
    }
    /* ★ convert 走长驻引擎进程（首次请求含管线加载仍慢，之后单页 ~10s）；
       ★ 2026-10-26 二修：长驻通道任何失败都回退一次性 spawn（不再把失败直接甩给用户），
         只有两条路都失败才报错，且错误里带两边的真实原因。 */
    let serveErr = ''
    if (payload.mode === 'convert' && payload.input && payload.output) {
      const out = payload.output
      try {
        ensureP2wServe()
        /* 超时按页数由渲染层动态下发（52 页 900s 固定值太紧），缺省仍 900s */
        const rr = await p2wServeReq(
          { input: payload.input, output: out, lang: payload.lang ?? 'ch', profile },
          Math.max(60000, Math.min(7200000, payload.timeoutMs ?? 900000)),
        )
        if (rr.ok) {
          try {
            const buf = await readFile(out)
            return { ok: true, code: 0, stdout: 'P2W_OK', stderr: '', exe: 'serve', bytes: new Uint8Array(buf) }
          } catch (err) {
            serveErr = '产物读取失败: ' + String(err)
          }
        } else {
          serveErr = String(rr.err ?? '引擎无应答').replace(/\s+/g, ' ').slice(0, 300)
        }
      } catch (err) {
        serveErr = 'serve 通道异常: ' + String(err)
      }
      /* 业务错误（输入不存在等）重试同样失败，直接回，省得再等一遍 */
      if (/输入文件不存在/.test(serveErr)) {
        const m = 'P2W_ERROR: ' + serveErr
        return { ok: false, code: -1, stdout: m, stderr: m, exe: 'serve' }
      }
    }
    /* 套件内 python 优先（venv → 内嵌 python312；用户点了「装入套件目录」就用它，不碰系统环境） */
    const done = async (r: { ok: boolean; code: number; stdout: string; stderr: string; exe: string }) => {
      const f = await finish(r)
      if (!f.ok && serveErr) {
        const detail = (f.stderr || '').replace(/\s+/g, ' ').trim().slice(0, 400) || ('退出码 ' + f.code)
        const m = 'P2W_ERROR: 长驻引擎失败（' + serveErr + '）｜一次性执行: ' + detail
        return { ...f, stdout: m, stderr: m }
      }
      return f
    }
    for (const cand of [suiteVenvPython(), suiteEmbPython()]) {
      if (existsSync(cand)) return await done(await tryRun(cand, pyArgs))
    }
    let r = await tryRun('python', pyArgs)
    if (!r.ok && /ENOENT|not recognized|找不到/i.test(r.stderr)) r = await tryRun('py', ['-3', ...pyArgs])
    return await done(r)
  })

  /* ── 套件中心（2026-10-26）────────────────────────────────────────
     存储目录 userData/suites/（动态套件安装根目录）+ state.json（停用状态）。
     内置套件（随软件分发）走「停用/启用」（引擎字节在 vendors 删不得）；
     动态套件（未来从套件中心下载）装到 SUITES_DIR/<id>/，卸载=删目录。 */
  ipcMain.handle('suite:read', async () => {
    const dir = join(app.getPath('userData'), 'suites')
    await mkdir(dir, { recursive: true })
    let disabled: string[] = []
    try {
      const st = JSON.parse(await readFile(join(dir, 'state.json'), 'utf8'))
      if (Array.isArray(st.disabled)) disabled = st.disabled.filter((x: unknown) => typeof x === 'string')
    } catch { /* 无 state.json = 全启用 */ }
    return { dir, disabled }
  })
  ipcMain.handle('suite:write', async (_e, patch: { disabled?: string[] }) => {
    const dir = join(app.getPath('userData'), 'suites')
    await mkdir(dir, { recursive: true })
    const statePath = join(dir, 'state.json')
    let cur: { disabled?: string[] } = {}
    try { cur = JSON.parse(await readFile(statePath, 'utf8')) } catch { /* ignore */ }
    const next = { ...cur, ...(patch.disabled !== undefined ? { disabled: patch.disabled } : {}) }
    await writeFile(statePath, JSON.stringify(next, null, 2), 'utf8')
    return { ok: true }
  })
  ipcMain.handle('suite:openDir', async () => {
    const dir = join(app.getPath('userData'), 'suites')
    await mkdir(dir, { recursive: true })
    shell.openPath(dir)
    return dir
  })
  /* OCR 套件（百度 PaddleOCR）动态安装/卸载——★ 2026-10-26 用户指令「把百度的装入
     套件的目录下（可卸载）」：引擎装 userData/suites/pdf2docx-ocr/venv（独立 venv，
     不碰系统 Python 环境）；卸载 = 删整个套件目录（干净卸载）。
     前置：本机需有 Python 3.9~3.12（venv 以它为基创建）。 */
  ipcMain.handle('suite:ocrInstall', async () => {
    const dir = suiteOcrDir()
    await mkdir(dir, { recursive: true })
    /* ★ 19:23 实锤升级：用户机器无系统 Python（cmd 里 python 是商店假 shim）——
       安装脚本改 PowerShell：有系统 Python 用 venv 隔离；没有则自动下载
       内嵌 Python 3.12（约 11MB，python.org 官方包）装进套件目录——零系统依赖。
       UTF-8 BOM 必带：PowerShell 5.1 无 BOM 按 ANSI 读，中文全乱码 */
    const ps1 = '\ufeff' + [
      "$ErrorActionPreference = 'Continue'",
      '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12',
      "$dir = Split-Path -Parent $MyInvocation.MyCommand.Path",
      "$venvPy = Join-Path $dir 'venv\\Scripts\\python.exe'",
      "$embPy  = Join-Path $dir 'python312\\python.exe'",
      "Write-Host '[1/5] 查找可用 Python...' -ForegroundColor Cyan",
      '$py = $null',
      'if (Test-Path $venvPy) { $py = $venvPy; Write-Host "  使用套件内 venv（跳过创建）" }',
      'elseif (Test-Path $embPy) { $py = $embPy; Write-Host "  使用套件内内嵌 Python（跳过下载）" }',
      'else {',
      '  foreach ($c in @(@{exe="python"; extra=@()}, @{exe="py"; extra=@("-3")})) {',
      '    & $c.exe @($c.extra + @("-c", "import sys")) 2>$null | Out-Null',
      '    if ($LASTEXITCODE -eq 0) {',
      '      Write-Host ("  找到系统 Python（" + $c.exe + "）——创建套件内 venv 隔离...")',
      '      $venvDir = Join-Path $dir "venv"',
      '      & $c.exe @($c.extra + @("-m", "venv", $venvDir))',
      '      if (Test-Path $venvPy) { $py = $venvPy; break }',
      '    }',
      '  }',
      '}',
      'if (-not $py) {',
      '  Write-Host "  未找到系统 Python —— 自动下载内嵌 Python 3.12（约 11MB，python.org 官方包）..." -ForegroundColor Yellow',
      '  New-Item -ItemType Directory -Force -Path (Join-Path $dir "python312") | Out-Null',
      '  $zip = Join-Path $env:TEMP "python-3.12.8-embed.zip"',
      '  Invoke-WebRequest -Uri "https://www.python.org/ftp/python/3.12.8/python-3.12.8-embed-amd64.zip" -OutFile $zip -UseBasicParsing',
      '  Expand-Archive -Path $zip -DestinationPath (Join-Path $dir "python312") -Force',
      '  $pth = Join-Path $embDir "python312._pth"',
      '  (Get-Content $pth) -replace "#\\s*import site", "import site" | Set-Content $pth',
      '  Write-Host "  引导 pip（get-pip）..."',
      '  $gp = Join-Path $embDir "get-pip.py"',
      '  Invoke-WebRequest -Uri "https://bootstrap.pypa.io/get-pip.py" -OutFile $gp -UseBasicParsing',
      '  & $embPy $gp',
      '  if ($LASTEXITCODE -ne 0) { Write-Host "get-pip 失败" -ForegroundColor Red; Read-Host "按回车退出"; exit 1 }',
      '  $py = $embPy',
      '}',
      "Write-Host '[2/5] 升级 pip...' -ForegroundColor Cyan",
      '& $py -m pip install -U pip',
      "Write-Host '[3/5] 安装百度 PaddleOCR 引擎到套件目录（约 600MB + 文档解析依赖，耐心等待）...' -ForegroundColor Cyan",
      '& $py -m pip install "paddlepaddle==3.2.1" "paddleocr[doc-parser]" python-docx docxcompose',
      'if ($LASTEXITCODE -ne 0) { Write-Host "安装失败——检查网络/代理后重跑本脚本" -ForegroundColor Red; Read-Host "按回车退出"; exit 1 }',
      "Write-Host '[4/5] 验证引擎...' -ForegroundColor Cyan",
      '& $py -c "from paddleocr import PPStructureV3; print(\'engine ok\')"',
      'if ($LASTEXITCODE -ne 0) { Write-Host "验证失败" -ForegroundColor Red; Read-Host "按回车退出"; exit 1 }',
      "Write-Host '[5/5] 完成！' -ForegroundColor Green",
      "Write-Host '回到软件 PDF转word 节点：引擎切到「百」即可转换扫描件。'",
      "Write-Host '注意：首次转换会自动下载模型（约 1-2GB，只下一次，之后离线可用）。'",
      'Read-Host "按回车关闭窗口"',
    ].join('\r\n')
    const psPath = join(dir, 'install_engine.ps1')
    await writeFile(psPath, ps1, 'utf8')
    const { spawn: sp } = await import('node:child_process')
    const child = sp('cmd.exe', ['/c', 'start', '', 'powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-NoExit', '-File', psPath], { detached: true, stdio: 'ignore', windowsHide: false })
    child.unref()
    return { ok: true, msg: '安装窗口已打开（有系统 Python 用 venv；没有自动下载内嵌 Python——引擎全部装在套件目录）' }
  })
  ipcMain.handle('suite:ocrUninstall', async () => {
    const dir = suiteOcrDir()
    if (!existsSync(dir)) return { ok: true, msg: '套件目录不存在（本就未安装）' }
    const { rm } = await import('fs/promises')
    try {
      await rm(dir, { recursive: true, force: true })
      return { ok: true, msg: '已卸载：套件目录已删除（userData/suites/pdf2docx-ocr）' }
    } catch (e) {
      return { ok: false, msg: '卸载失败（目录被占用？关闭正在运行的转换窗口后重试）: ' + String(e).slice(0, 80) }
    }
  })

  /* ★ 2026-10-24 资产导出到文件夹（用户肉眼查重——同指纹同名同大小一目了然） */
  ipcMain.handle('assets:export', async (_e, items: { name: string; bytesB64: string }[]) => {
    const r = await dialog.showSaveDialog({
      title: '选择资产导出位置（会创建文件夹）',
      defaultPath: '资产导出-' + new Date().toISOString().slice(0, 16).replace(/[:T]/g, '-'),
      properties: ['createDirectory'],
    })
    if (r.canceled || !r.filePath) return { ok: false, reason: 'canceled' }
    const dir = r.filePath
    const { mkdirSync, writeFileSync } = await import('fs')
    mkdirSync(dir, { recursive: true })
    let n = 0
    for (const it of items) {
      try {
        writeFileSync(dir + '/' + it.name.replace(/[\/:*?"<>|]/g, '_'), Buffer.from(it.bytesB64, 'base64'))
        n++
      } catch { /* 单文件失败继续 */ }
    }
    return { ok: true, dir, count: n }
  })
  ipcMain.handle('gpu:getAccel', () => __gpuAccelOn())
  ipcMain.handle('gpu:setAccel', (_e, on: boolean) => {
    try { writeFileSync(__gpuPrefFile(), on ? 'on' : 'off', 'utf8'); return true } catch { return false }
  })

  /** 用系统默认软件打开产物：落临时文件 → shell.openPath（2026-10-23 历史右键） */
  ipcMain.handle('shell:openExternalFile', async (_e, name: string, data: Uint8Array) => {
    const safe = String(name || 'file').replace(/[\/:*?"<>|]/g, '_')
    const tmp = join(tmpdir(), `client-base-open-${Date.now().toString(36)}-${safe}`)
    await writeFile(tmp, Buffer.from(data))
    await shell.openPath(tmp)
    return tmp
  })

  /* ★ 2026-10-26 用户指令「所有PDF增加个 用本地浏览器打开」：
     落临时文件 → 找本机浏览器（Edge/Chrome 常见安装路径逐个试）→ 以参数打开
     （浏览器自带 PDF 查看器渲染）；一个都没有时退回系统关联打开。 */
  ipcMain.handle('shell:openInBrowser', async (_e, name: string, data: Uint8Array) => {
    const safe = String(name || 'file').replace(/[\/:*?"<>|]/g, '_')
    const tmp = join(tmpdir(), `client-base-web-${Date.now().toString(36)}-${safe}`)
    await writeFile(tmp, Buffer.from(data))
    const candidates = [
      'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
      'C:/Program Files/Microsoft/Edge/Application/msedge.exe',
      'C:/Program Files/Google/Chrome/Application/chrome.exe',
      'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
      process.env.LOCALAPPDATA ? join(process.env.LOCALAPPDATA, 'Google/Chrome/Application/chrome.exe') : '',
    ].filter(Boolean)
    for (const exe of candidates) {
      if (!existsSync(exe)) continue
      try {
        const child = spawn(exe, ['--single-argument', tmp], { detached: true, stdio: 'ignore', windowsHide: true })
        await new Promise<void>((res) => {
          child.once('spawn', () => res())
          child.once('error', () => res())
          setTimeout(res, 1500)
        })
        if (child.pid) { child.unref(); return tmp }
      } catch { /* 下一个候选 */ }
    }
    await shell.openPath(tmp)
    return tmp
  })


  /** 把产物另存到用户指定位置（kind 决定对话框标题与扩展名过滤） */
  ipcMain.handle('doc:saveAs', async (_e, defaultName: string, data: Uint8Array, kind?: string) => {
    const f = kindOf(kind)
    const win = BrowserWindow.getFocusedWindow()
    const res = await dialog.showSaveDialog(win ?? ({} as never), {
      title: f.saveTitle,
      defaultPath: defaultName,
      filters: [{ name: f.label, extensions: [f.ext] }],
    })
    if (res.canceled || !res.filePath) return null
    await writeFile(res.filePath, Buffer.from(data))
    return res.filePath
  })

  /**
   * HTML → PDF（隐藏窗口打印，Chromium 分页引擎）。
   * 渲染层把 docx-preview 渲染好的完整 HTML 传过来，这里只负责打印，
   * 排版完全由预览的同一套渲染决定 —— 预览即 PDF。
   *
   * 走临时文件而不是 data: URL：报告里内联的 base64 图片很容易让 HTML 涨到几 MB，
   * 而 Chromium 对 data: URL 导航有长度上限，超限会直接白屏。
   */
  ipcMain.handle('doc:pdf', async (_e, html: string, defaultName: string) => {
    const tmp = join(tmpdir(), `client-base-print-${Date.now()}.html`)
    const printWin = new BrowserWindow({
      show: false,
      webPreferences: { offscreen: true, sandbox: false },
    })
    try {
      await writeFile(tmp, html, 'utf8')
      await printWin.loadFile(tmp)
      // 等字体/图片就绪（base64 图片已内联，主要是等 layout 稳定）
      await new Promise((r) => setTimeout(r, 400))
      const pdf = await printWin.webContents.printToPDF({
        pageSize: 'A4',
        printBackground: true,
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
      })
      const win = BrowserWindow.getFocusedWindow()
      const res = await dialog.showSaveDialog(win ?? ({} as never), {
        title: '导出 PDF',
        defaultPath: defaultName,
        filters: [{ name: 'PDF 文档', extensions: ['pdf'] }],
      })
      if (res.canceled || !res.filePath) return null
      await writeFile(res.filePath, Buffer.from(pdf))
      return res.filePath
    } finally {
      printWin.destroy()
      try {
        await unlink(tmp)
      } catch {
        /* 临时文件删不掉无所谓，系统在 %TEMP% 会自己清 */
      }
    }
  })

  try {
    await createWindow()
  } catch (e) {
    console.error('[electron] 启动窗口失败', e)
    app.quit()
  }
})

app.on('window-all-closed', () => {
  // 隐藏窗口模式下不触发；保留 darwin 常驻。Windows 由 before-quit/isQuitting 控制真退出
  if (process.platform !== 'darwin' && isQuitting) app.quit()
})
