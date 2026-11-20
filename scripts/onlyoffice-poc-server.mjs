/**
 * OnlyOffice 离线内嵌 POC 服务（5199 独立测试入口）
 *
 *   http://127.0.0.1:5199/poc
 *
 * ★ 核心逻辑全部在 ./oo-core.mjs（与 app-server:3799 共享同一套 handler）。
 *   本文件只负责：起 5199 端口 + 注册 upgrade 隧道。低层调试用；
 *   应用内节点编辑器走 app-server 的同源路由（/poc → vite 代理 → 3799）。
 */
import http from 'node:http'
import { handleOoRequest, handleOoUpgrade } from './oo-core.mjs'
import { ensureDocService } from './oo-docservice.mjs'

const PORT = 5199

const server = http.createServer(async (req, res) => {
  // ★ 客户端中断（导航取消/超时断开）后异步写会 ERR_STREAM_WRITE_AFTER_END，
  //   Response 的 'error' 事件无人接会直接崩进程（2026-10-10 无头浏览器实测崩两次）——
  //   客户端断开是常态不是异常，静默吞掉
  res.on('error', () => {})
  try {
    // ★ 与 app-server 同款守卫（2026-10-10）：handleOoRequest 的流式分支（sendFile）
    //   写完 200 头后 return undefined（falsy）——只看返回值会把「已处理」误判为
    //   「未处理」再补一刀 → 响应体被 nf(outer) 截断/覆盖（无头浏览器实测 wasm 只剩59字节）
    const handled = await handleOoRequest(req, res)
    if (handled || res.writableEnded || res.headersSent) return
    res.statusCode = 404
    res.end('nf(outer) ' + (req.url || '').slice(0, 80))
  } catch (e) {
    try { res.statusCode = 500; res.end(String(e).slice(0, 300)) } catch { /* ignore */ }
  }
})

server.on('upgrade', handleOoUpgrade)

server.listen(PORT, '127.0.0.1', () => {
  console.log(`[oo-poc] http://127.0.0.1:${PORT}/poc`)
  // ★ DocService(3798) 幂等拉起（2026-10-07 排障入链）：/doc 协作通道的必跑依赖，
  //   死了编辑器永远 loading。detached 独立进程，5199 退出不连带杀。
  ensureDocService().catch(() => {})
})
