/* ★ 全局日志总线（2026-10-22 用户指令「画布下方增加 log 显示面板；右上角增加个
 *   和 oo 节点内一致的显示诊断信息」）：
 *   OO 节点内的诊断（addDiag）、全局 JS 错误都汇到这里，LogDock 面板消费。
 *   模块级单例 + 订阅数组；环形缓冲 500 条，新日志前插（与 OO 诊断抽屉同序）。
 *   纯前端内存态，不落盘（落盘链路 /office-log 已有，各写各的）。 */

export type LogKind = 'info' | 'ok' | 'warn' | 'err'

export interface LogEntry {
  at: number
  ts: string
  kind: LogKind
  msg: string
}

const MAX = 500
let seq: LogEntry[] = []
const subs = new Set<(e: LogEntry) => void>()

function tsNow() {
  return new Date().toLocaleTimeString('zh-CN', { hour12: false })
}

/* kind 嗅探：与 OO 诊断的书写习惯对齐（⚠=警告 ✅=成功 ERR/失败/Error=错误） */
export function sniffKind(msg: string): LogKind {
  if (/^⚠|警告|WARN/i.test(msg)) return 'warn'
  if (/^✅|成功|就绪/.test(msg)) return 'ok'
  if (/ERR|失败|Error|异常|❌|FATAL/i.test(msg)) return 'err'
  return 'info'
}

export function pushLog(msg: string, kind?: LogKind): LogEntry {
  const e: LogEntry = { at: Date.now(), ts: tsNow(), kind: kind ?? sniffKind(msg), msg }
  seq = [e, ...seq].slice(0, MAX)
  subs.forEach((fn) => {
    try { fn(e) } catch { /* 订阅者异常不传染 */ }
  })
  return e
}

export function getLogs(): readonly LogEntry[] {
  return seq
}

export function clearLogs() {
  seq = []
  subs.forEach((fn) => {
    try { fn(null as unknown as LogEntry) } catch { /* clear 信号，容忍 */ }
  })
}

export function subscribeLog(fn: (e: LogEntry | null) => void): () => void {
  subs.add(fn as (e: LogEntry) => void)
  return () => subs.delete(fn as (e: LogEntry) => void)
}

/* 全局 JS 错误接入（一次安装）；OO addDiag 由组件侧直接调 pushLog */
let installed = false
export function installGlobalLogHooks() {
  if (installed) return
  installed = true
  window.addEventListener('error', (ev) => {
    pushLog(`[全局] ${ev.message || 'Script error'}`, 'err')
  })
  window.addEventListener('unhandledrejection', (ev) => {
    const r = (ev as PromiseRejectionEvent).reason
    const stack = r instanceof Error ? (r.stack || '').split('\n').slice(1, 4).join(' | ').slice(0, 260) : ''
    pushLog(`[未处理Promise] ${r instanceof Error ? r.message : String(r)}${stack ? ' ← ' + stack : ''}`, 'err')
  })
  pushLog('日志面板就绪（OO 诊断与全局错误在此汇流）', 'ok')
}
