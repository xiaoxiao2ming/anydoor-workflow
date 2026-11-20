/**
 * OO 离线服务（app-server 挂载入口）
 *
 * ★ 委托到 ./oo-core.mjs —— 与 poc-server(5199) 共享同一套 handler：
 *   全部九项修复（AllFonts 原样直出 / 确定性协作 key / Editor.bin 反转保存 /
 *   平铺工作区 / /cache 同源代理 / ws /doc 隧道 / 相对 URL / SW 映射 / x-forwarded-host 动态化）。
 *
 * app-server 用法：`if (await handleOoRequest(req, res)) return`
 * 升级隧道：`server.on('upgrade', handleOoUpgrade)`（/doc/<key>/c/ 的 ws）
 */
// ★ 动态加载 oo-core（cache-bust per call）：oo-core.mjs 改动即时生效——app-server 零重启，
//   编辑器会话不被打断（静态 re-export 时代：每次改 oo-core 都得重启 3799，用户窗口全断）
import { pathToFileURL, fileURLToPath } from 'node:url'
import path from 'node:path'
import fs2 from 'node:fs'

/* ★★ oo-core 的真实路径解析（2026-10-06 修 C 端生产模式打不开）
 *
 * 原实现只用 `dirname(fileURLToPath(import.meta.url))` 定位 oo-core.mjs，
 * 这在**源码直跑**（node app-server.mjs）时正确，但在 **C 端生产模式**下必然错：
 *   vite-plugin-electron 会把 app-server.mjs + 本文件**整个内联进**
 *   dist-electron/main.js ⇒ import.meta.url 变成 .../dist-electron/main.js
 *   ⇒ 解析成 dist-electron/oo-core.mjs，而真实文件在 <项目根>/scripts/oo-core.mjs
 *   ⇒ statSync 抛 ENOENT ⇒ loadOoCore 每次请求都失败。
 *
 * 后果比「OO 编辑器打不开」严重得多：app-server 的 OO 分发在**静态托管之前**
 * 拦截所有非 /api/ 请求，异常一路冒泡到最外层 catch ⇒
 *   **连根路径 `/` 都返回 404** ⇒ Electron 窗口永远停在 about:blank ⇒
 *   C 端出包后整个应用打不开。
 *
 * 解法：允许调用方（app-server.mjs，它能拿到项目根）显式指定 ooCorePath；
 * 显式给了就用给的，没给才回退到 import.meta.url 推导（保持源码直跑行为不变）。
 */
let __ooCorePath = null
export function setOoCorePath (absPath) {
  if (absPath) __ooCorePath = absPath
}
function resolveOoCorePath () {
  if (__ooCorePath) return __ooCorePath
  // 环境变量兜底：Electron 主进程知道项目根，可直接注入
  const fromEnv = process.env.OO_CORE_PATH
  if (fromEnv && existsSyncSafe(fromEnv)) return fromEnv
  return path.join(path.dirname(fileURLToPath(import.meta.url)), 'oo-core.mjs')
}
function existsSyncSafe (p) {
  try { fs2.statSync(p); return true } catch { return false }
}

// ★ 快照式热加载 v2（2026-10-01）：签名缓存版
//   v1 每请求都复制+import 一次（~26ms/请求 × 编辑器一次打开几百请求 = 可感知的加载滞后，
//   且每个模块实例永久驻留 V8 → 内存泄漏 + 单线程编译串行排队）
//   v2 按 mtime+size 签名缓存：文件没变 → 直接复用已加载模块（开销≈0）；
//   变了 → 才复制到唯一临时路径 import（热加载语义不变），并发请求用同一份 in-flight promise 防抖
let __snapSeq = 0
let __cachedMod = null
let __cachedSig = ''
let __loading = null
let __loadingSig = ''
async function loadOoCore() {
  const core = resolveOoCorePath()
  let sig
  try {
    const st = fs2.statSync(core)
    sig = st.mtimeMs + ':' + st.size
  } catch (e) { sig = 'missing' }
  if (__cachedMod && __cachedSig === sig) return __cachedMod
  if (!__loading || __loadingSig !== sig) {
    __loadingSig = sig
    __loading = (async () => {
      const tmp = core + '.' + Date.now() + '.' + (__snapSeq++) + '.snap.mjs'
      fs2.copyFileSync(core, tmp)
      try {
        const mod = await import(pathToFileURL(tmp).href)
        __cachedMod = mod
        __cachedSig = sig
        return mod
      } finally {
        setTimeout(() => { try { fs2.unlinkSync(tmp) } catch { /* ignore */ } }, 3000)
      }
    })()
  }
  return __loading
}

export async function handleOoRequest(req, res) {
  const mod = await loadOoCore()
  return mod.handleOoRequest(req, res)
}

export async function handleOoUpgrade(req, socket, head) {
  try {
    const mod = await loadOoCore()
    return mod.handleOoUpgrade(req, socket, head)
  } catch (e) {
    try { socket.destroy() } catch { /* ignore */ }
  }
}
