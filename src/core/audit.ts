/**
 * 操作审计日志（2026-10-17 用户指令）
 *
 * 「从现在开始，我对软件的每个点击等操作、键盘输入、软件的反馈，都放入日志。」
 *
 * ── 设计约束 ────────────────────────────────────────────────────────────
 * ① 不能拖慢交互：所有 DOM 事件只做「拼一行字符串 + 入队」，落盘是 250ms 批量异步；
 *    高频事件必须先合并（input 同目标 300ms、wheel 全局 500ms、长按 keydown），
 *    否则鼠标滚滚轮一次就能冲垮日志文件。
 * ② 不能因为日志失败而影响主流程：写盘/网络/IPC 全线 try-catch 静默降级，
 *    最后兜底写内存环形 + localStorage，保证「至少能捞出来」。
 * ③ 要能回答「谁触发的、触发了什么、软件怎么回应的」：
 *    所以每条都带 会话 id + 毫秒时间戳 + 元素指纹（tag#id.class "文案" @坐标），
 *    业务侧再用 auditFeedback() 记「软件反馈」，两者按时间顺序交错即可复盘。
 * ④ 落盘位置统一由宿主决定：
 *    Electron → window.api.auditWrite（主进程 appendFile）
 *    浏览器   → POST /audit-log（app-server 写同一个文件）
 *    都没有   → localStorage 环形缓冲，随时可导出。
 *
 * 日志行格式（纯文本，grep 友好）：
 *   2026-10-17 21:30:01.123 [s1a2b3c] 点击 pointerdown | button#btn-save.btn.primary "保存" @(312,48) 左键 节点=n-12
 */

export type AuditKind =
  | '点击' | '键盘' | '输入' | '滚轮' | '拖拽' | '焦点'
  | '反馈' | '系统' | '错误' | '编辑器'

/** 落盘失败时的本地环形缓冲键（浏览器模式兜底，可人工导出） */
const LOCAL_KEY = 'cb_audit_ring'
const LOCAL_MAX = 4000
/** 内存环形：供「导出最近日志」类按钮与现场诊断读，不参与落盘 */
const RING_MAX = 3000
/** 单行上限：日志是给人看的，超过这个长度的值（如整段 xlsx base64）没有价值 */
const LINE_MAX = 900

const SESSION = 's' + Date.now().toString(36).slice(-6)

/** 只在有 DOM 的环境采集/入队。Node（vitest）里 audit() 直接返回：
 *  否则 store 层的埋点会给测试塞一堆 250ms 定时器（收尾时报 "work after teardown"）。 */
const IS_BROWSER = typeof window !== 'undefined' && typeof document !== 'undefined'

const buf: string[] = []
const ring: string[] = []
let timer: ReturnType<typeof setTimeout> | null = null
let enabled = true
let installed = false
let lastWriteError = ''

/* ── 时间与格式化 ──────────────────────────────────────────────── */

function pad(n: number, w = 2): string {
  return String(n).padStart(w, '0')
}

function stamp(): string {
  const d = new Date()
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ` +
    `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}.${pad(d.getMilliseconds(), 3)}`
}

/** 值清洗：换行折叠成 ⏎（保证「一行一条记录」，日志才能按行 grep）；超长截断 */
function clean(v: unknown): string {
  if (v == null) return ''
  let s: string
  try {
    s = typeof v === 'string' ? v : JSON.stringify(v)
  } catch {
    s = String(v)
  }
  if (s == null) return ''
  s = s.replace(/\r\n?/g, '⏎').replace(/\n/g, '⏎').replace(/\t/g, ' ')
  return s.length > LINE_MAX ? s.slice(0, LINE_MAX) + `…(+${s.length - LINE_MAX})` : s
}

/** 供测试/外部复用：把任意值清洗成「一行日志」里的安全片段 */
export function cleanAuditText(v: unknown): string {
  return clean(v)
}

/* ── 元素指纹 ──────────────────────────────────────────────────── */

/** 只取「能定位到人」的信息：标签 / id / 前 3 个 class / 语义文案 / 坐标 / 所在节点 */
export function describeElement(el: Element | null | undefined): string {
  if (!el) return '-'
  const anyEl = el as Element & { tagName?: string }
  const tag = (anyEl.tagName || '?').toLowerCase()
  const parts: string[] = [tag]

  const id = (el as HTMLElement).id
  if (id) parts.push('#' + id)

  const cls = (el.getAttribute('class') || '').trim().split(/\s+/).filter(Boolean).slice(0, 3)
  if (cls.length) parts.push('.' + cls.join('.'))

  // 表单控件：type / placeholder 往往比 class 更能说明「这是哪个框」
  if (tag === 'input' || tag === 'textarea' || tag === 'select') {
    const type = (el.getAttribute('type') || '').toLowerCase()
    if (type && type !== 'text') parts.push('[type=' + type + ']')
    const ph = el.getAttribute('placeholder')
    if (ph && ph.trim()) parts.push('[placeholder=' + ph.trim().slice(0, 20) + ']')
  }

  // 语义文案优先级：显式审计标签 > testid > aria-label > title > 自身文本
  const attrs = ['data-audit', 'data-testid', 'aria-label', 'title', 'name', 'placeholder']
  let label = ''
  for (const a of attrs) {
    const v = el.getAttribute(a)
    if (v && v.trim()) { label = v.trim(); break }
  }
  if (!label) {
    const txt = (el as HTMLElement).innerText || el.textContent || ''
    label = txt.replace(/\s+/g, ' ').trim()
  }
  if (label) parts.push(' "' + label.slice(0, 30) + (label.length > 30 ? '…' : '') + '"')

  // 光秃秃的 div/input（没有 id、没有 class）无法定位 —— 补一层「它在谁里面」
  if (!id && !cls.length) {
    const host = el.parentElement && el.parentElement.closest('[id], [class]')
    if (host) {
      const hid = host.id ? '#' + host.id : ''
      const hcls = (host.getAttribute('class') || '').trim().split(/\s+/).filter(Boolean).slice(0, 2).join('.')
      parts.push(' ::in=' + host.tagName.toLowerCase() + hid + (hcls ? '.' + hcls : ''))
    }
  }

  // 画布节点 / 连线：React Flow 用 data-id 标识，这是复盘时最关键的一层上下文
  const node = el.closest('.react-flow__node')
  const edge = el.closest('.react-flow__edge')
  if (node) parts.push(' 节点=' + (node.getAttribute('data-id') || '?'))
  else if (edge) parts.push(' 连线=' + (edge.getAttribute('data-id') || '?'))

  return parts.join('')
}

function isPasswordLike(el: Element): boolean {
  if (el.getAttribute('data-audit-mask') != null) return true
  const t = (el.getAttribute('type') || '').toLowerCase()
  return t === 'password'
}

/** 控件当前值（输入类事件用）；密码类一律掩码，只留长度 —— 日志会被分享出去 */
export function valueOf(el: Element | null | undefined): string {
  if (!el) return ''
  const tag = el.tagName.toLowerCase()
  if (tag === 'input') {
    const inp = el as HTMLInputElement
    const type = (inp.type || 'text').toLowerCase()
    if (type === 'checkbox' || type === 'radio') return inp.checked ? '已选中' : '未选中'
    if (type === 'file') return `文件=${(inp.files && inp.files.length) || 0}`
    if (isPasswordLike(el) || type === 'password') return `••••(len=${inp.value.length})`
    return inp.value
  }
  if (tag === 'textarea') {
    const ta = el as HTMLTextAreaElement
    return isPasswordLike(el) ? `••••(len=${ta.value.length})` : ta.value
  }
  if (tag === 'select') {
    const sel = el as HTMLSelectElement
    return Array.from(sel.selectedOptions || []).map((o) => o.text).join(',') || sel.value
  }
  if ((el as HTMLElement).isContentEditable) {
    return (el as HTMLElement).innerText || ''
  }
  return ''
}

/** 按键记法：Ctrl+Shift+S / Enter / F2 —— 修饰键顺序固定，便于统计 */
export function keyLabel(e: KeyboardEvent): string {
  const mods: string[] = []
  if (e.ctrlKey) mods.push('Ctrl')
  if (e.altKey) mods.push('Alt')
  if (e.shiftKey) mods.push('Shift')
  if (e.metaKey) mods.push('Meta')
  const k = e.key === ' ' ? 'Space' : e.key
  return mods.length ? mods.join('+') + '+' + k : k
}

/* ── 入队与落盘 ────────────────────────────────────────────────── */

function pushLine(line: string): void {
  ring.push(line)
  if (ring.length > RING_MAX) ring.splice(0, ring.length - RING_MAX)
  buf.push(line)
  if (buf.length >= 120) { flush(); return }
  if (!timer) timer = setTimeout(() => flush(), 250)
}

function mirrorLocal(lines: string[]): void {
  try {
    const raw = localStorage.getItem(LOCAL_KEY)
    const old: string[] = raw ? (JSON.parse(raw) as string[]) : []
    const merged = old.concat(lines)
    localStorage.setItem(LOCAL_KEY, JSON.stringify(merged.slice(-LOCAL_MAX)))
  } catch { /* 隐私模式 / 配额满：放弃本地镜像，内存 ring 仍在 */ }
}

/**
 * 批量落盘。三条通道按可用性降级：
 * Electron IPC → HTTP（web/dev）→ localStorage。
 * 任何一条失败都往下走，绝不抛。
 */
export function flush(): void {
  if (timer) { clearTimeout(timer); timer = null }
  if (!buf.length) return
  const lines = buf.splice(0, buf.length)
  const w = globalThis as unknown as { api?: { auditWrite?: (l: string[]) => Promise<unknown> } }

  try {
    if (w.api && typeof w.api.auditWrite === 'function') {
      void w.api.auditWrite(lines).catch((e: unknown) => {
        lastWriteError = String(e)
        mirrorLocal(lines)
      })
      return
    }
  } catch (e) {
    lastWriteError = String(e)
  }

  try {
    if (typeof fetch === 'function' && typeof location !== 'undefined' && /^https?:/.test(location.protocol)) {
      void fetch('/audit-log', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ lines }),
      }).then((r) => {
        if (!r.ok) mirrorLocal(lines)
      }).catch(() => mirrorLocal(lines))
      return
    }
  } catch { /* 落到本地镜像 */ }

  mirrorLocal(lines)
}

/* ── 对外 API ──────────────────────────────────────────────────── */

/**
 * 记一条审计日志。
 * @param kind   大类（点击/键盘/输入/反馈…）
 * @param action 动作名（pointerdown / 保存 / 同步域…），业务埋点用短横线英文或中文均可
 * @param detail 详情（自由文本，会清洗成单行）
 */
export function audit(kind: AuditKind | string, action: string, detail = ''): void {
  if (!enabled || !IS_BROWSER) return
  try {
    pushLine(`${stamp()} [${SESSION}] ${kind} ${action}${detail ? ' | ' + clean(detail) : ''}`)
  } catch { /* 日志自身永不影响主流程 */ }
}

/** 软件反馈专用入口（语义化糖）：状态栏变化、脏标记、保存结果、同步结果… */
export function auditFeedback(action: string, detail = ''): void {
  audit('反馈', action, detail)
}

/** 编辑器（OnlyOffice iframe）内动作，由 iframe 通过 postMessage 上报后转发到这里 */
export function auditEditor(action: string, detail = ''): void {
  audit('编辑器', action, detail)
}

/** 调试用：取最近的日志行（内存 ring，不落盘） */
export function auditRecent(n = 200): string[] {
  return ring.slice(-n)
}

export function auditSession(): string {
  return SESSION
}

export function auditStatus(): { enabled: boolean; session: string; buffered: number; lastError: string } {
  return { enabled, session: SESSION, buffered: buf.length, lastError: lastWriteError }
}

/** 运行期开关（排查性能问题时临时关掉；关掉时立即把已有队列刷出去） */
export function setAuditEnabled(v: boolean): void {
  enabled = v
  if (!v) flush()
  else audit('系统', '审计日志已重新开启')
}

/* ── 全局事件采集 ──────────────────────────────────────────────── */

/** input 合并窗口：同目标连续敲键只留一条（带次数与最终值） */
const INPUT_WINDOW = 300
/** wheel 聚合窗口 */
const WHEEL_WINDOW = 500

type PendingInput = { at: number; n: number; el: Element | null; timer: ReturnType<typeof setTimeout> | null }
const pendingInputs = new WeakMap<EventTarget, PendingInput>()

let wheelAcc: { dy: number; dx: number; el: Element | null; n: number; timer: ReturnType<typeof setTimeout> | null } = {
  dy: 0, dx: 0, el: null, n: 0, timer: null,
}
const scrollSeen = new WeakMap<EventTarget, number>()
let lastPointer: { at: number; target: EventTarget | null } = { at: 0, target: null }
const repeatKeys = new WeakMap<EventTarget, { key: string; n: number; at: number }>()

function withPos(el: Element | null): string {
  try {
    const r = el && el.getBoundingClientRect ? el.getBoundingClientRect() : null
    return r ? ` @(${Math.round(r.left)},${Math.round(r.top)})` : ''
  } catch {
    return ''
  }
}

/**
 * 安装全局采集。幂等：重复调用不会重复挂监听（HMR / 多次 render 安全）。
 * @returns 卸载函数（仅测试/热更新需要）
 */
export function installAudit(): () => void {
  if (typeof window === 'undefined' || typeof document === 'undefined') return () => {}
  if (installed) return () => {}
  installed = true

  audit('系统', '审计会话开始', `会话=${SESSION} 页面=${location.pathname} UA=${navigator.userAgent.slice(0, 80)}`)

  const opts = { capture: true, passive: true } as const
  const optsActive = { capture: true } as const

  /* ① 指针按下 —— 「每个点击」的主记录点。
   *    用 pointerdown 而不是 click：click 要等抬起且可能被 preventDefault 吞掉，
   *    而 pointerdown 一定能拿到「用户按了哪个控件」。 */
  const onPointerDown = (ev: Event) => {
    const e = ev as PointerEvent
    lastPointer = { at: Date.now(), target: e.target }
    const el = e.target as Element | null
    const btn = ['左键', '中键', '右键', '侧键1', '侧键2'][e.button] || `button${e.button}`
    audit('点击', 'pointerdown', `${describeElement(el)}${withPos(el)} ${btn} 坐标=(${Math.round(e.clientX)},${Math.round(e.clientY)})`)
  }
  document.addEventListener('pointerdown', onPointerDown, opts)

  /* ② click —— 只补记「没有 pointerdown 前导」的点击（键盘 Enter/Space 激活按钮、
   *    脚本触发），避免每次点击写两行。 */
  const onClick = (ev: Event) => {
    const e = ev as MouseEvent
    const recent = lastPointer.at && Date.now() - lastPointer.at < 700 && lastPointer.target === e.target
    if (recent) return
    const el = e.target as Element | null
    audit('点击', 'click(键盘/脚本触发)', describeElement(el) + withPos(el))
  }
  document.addEventListener('click', onClick, opts)

  const onDblClick = (ev: Event) => {
    const el = ev.target as Element | null
    audit('点击', 'dblclick', describeElement(el) + withPos(el))
  }
  document.addEventListener('dblclick', onDblClick, opts)

  /* ③ 右键菜单 —— 画布/表格里的隐藏操作大多在右键里，排查时经常是断点 */
  const onCtxMenu = (ev: Event) => {
    const el = ev.target as Element | null
    audit('点击', 'contextmenu(右键)', describeElement(el) + withPos(el))
  }
  document.addEventListener('contextmenu', onCtxMenu, opts)

  /* ④ 键盘 —— 全量记，但长按（e.repeat）收敛成「×N」，
   *    输入法组字过程不记（isComposing 时 key=Process，记了也是噪音）。 */
  const onKeyDown = (ev: Event) => {
    const e = ev as KeyboardEvent
    const el = e.target as Element | null
    if (e.isComposing) return
    if (e.key === 'Process') return
    if (e.repeat) {
      const rec = repeatKeys.get(e.target as EventTarget)
      if (rec && rec.key === e.key && Date.now() - rec.at < 1200) { rec.n++; rec.at = Date.now(); return }
      repeatKeys.set(e.target as EventTarget, { key: e.key, n: 2, at: Date.now() })
    }
    audit('键盘', 'keydown', `${keyLabel(e)} 焦点=${describeElement(el)}`)
  }
  document.addEventListener('keydown', onKeyDown, optsActive)

  /* ⑤ 输入 —— 合并窗口内只出一条（含次数与最终值）。
   *    逐字符记会把日志撑爆（打 50 个字 = 50 行），而复盘只需要「改了哪个框、改成什么」。 */
  const flushInput = (target: EventTarget, rec: PendingInput) => {
    const el = rec.el
    const v = valueOf(el)
    const times = rec.n > 1 ? ` ×${rec.n}` : ''
    audit('输入', 'value', `${describeElement(el)}${times} 值="${v.slice(0, 160)}"${v.length > 160 ? '…' : ''}`)
  }
  const onInput = (ev: Event) => {
    const e = ev as InputEvent
    const el = e.target as Element | null
    if (!el) return
    const key = e.target as EventTarget
    const now = Date.now()
    const rec = pendingInputs.get(key)
    if (rec) {
      rec.n++
      rec.at = now
      if (rec.timer) clearTimeout(rec.timer)
      rec.timer = setTimeout(() => { pendingInputs.delete(key); flushInput(key, rec) }, INPUT_WINDOW)
      return
    }
    const fresh: PendingInput = { at: now, n: 1, el, timer: null }
    pendingInputs.set(key, fresh)
    fresh.timer = setTimeout(() => { pendingInputs.delete(key); flushInput(key, fresh) }, INPUT_WINDOW)
  }
  document.addEventListener('input', onInput, opts)

  /* change 单独记：select / checkbox / 文件选择不一定触发 input，
   * 且它代表「提交完成」这个语义，比 input 更适合作为因果链的时间点。 */
  const onChange = (ev: Event) => {
    const el = ev.target as Element | null
    if (!el) return
    const tag = el.tagName.toLowerCase()
    if (tag === 'input' || tag === 'textarea') {
      const t = ((el as HTMLInputElement).type || '').toLowerCase()
      if (t !== 'checkbox' && t !== 'radio' && t !== 'file') return // 交给 input 合并，避免双份
    }
    audit('输入', 'change', `${describeElement(el)} 值="${valueOf(el).slice(0, 160)}"`)
  }
  document.addEventListener('change', onChange, opts)

  /* ⑥ 滚轮 —— 按窗口聚合，否则一次滚轮手势就是几十行 */
  const onWheel = (ev: Event) => {
    const e = ev as WheelEvent
    if (!wheelAcc.timer) {
      wheelAcc = { dy: 0, dx: 0, el: e.target as Element, n: 0, timer: null }
    }
    wheelAcc.dy += e.deltaY
    wheelAcc.dx += e.deltaX
    wheelAcc.n++
    if (wheelAcc.timer) clearTimeout(wheelAcc.timer)
    wheelAcc.timer = setTimeout(() => {
      const a = wheelAcc
      wheelAcc = { dy: 0, dx: 0, el: null, n: 0, timer: null }
      audit('滚轮', 'wheel', `${describeElement(a.el)} ×${a.n} ΔY=${Math.round(a.dy)} ΔX=${Math.round(a.dx)}`)
    }, WHEEL_WINDOW)
  }
  document.addEventListener('wheel', onWheel, opts)

  /* ⑦ 滚动容器 —— 600ms/元素 节流；记录最终 scrollTop/Left，方便复现「滚到哪看不到」 */
  const onScroll = (ev: Event) => {
    const t = ev.target as (EventTarget & { scrollTop?: number; scrollLeft?: number; scrollHeight?: number }) | null
    if (!t || typeof t.scrollTop !== 'number') return
    const now = Date.now()
    const prev = scrollSeen.get(t as EventTarget) || 0
    if (now - prev < 600) return
    scrollSeen.set(t as EventTarget, now)
    const el = t as unknown as Element
    audit('滚轮', 'scroll', `${describeElement(el)} top=${Math.round(t.scrollTop)} left=${Math.round(t.scrollLeft || 0)}`)
  }
  document.addEventListener('scroll', onScroll, opts)

  /* ⑧ 拖拽 —— 画布拖节点、拖文件、拖列宽，都在这里留痕 */
  const onDragStart = (ev: Event) => {
    const e = ev as DragEvent
    const el = e.target as Element | null
    audit('拖拽', 'dragstart', describeElement(el) + withPos(el))
  }
  const onDrop = (ev: Event) => {
    const e = ev as DragEvent
    const el = e.target as Element | null
    const n = e.dataTransfer ? e.dataTransfer.files.length : 0
    audit('拖拽', 'drop', `${describeElement(el)}${withPos(el)} 文件数=${n}`)
  }
  document.addEventListener('dragstart', onDragStart, opts)
  document.addEventListener('drop', onDrop, opts)

  /* ⑨ 焦点 —— 复盘「为什么这次按键没生效」时，焦点位置是第一现场 */
  const onFocusIn = (ev: Event) => {
    const el = ev.target as Element | null
    if (!el || el === document.body) return
    audit('焦点', 'focus', describeElement(el))
  }
  document.addEventListener('focusin', onFocusIn, opts)

  /* ⑩ 软件自身报错 —— 这些是「软件反馈」里最需要留档的一类 */
  const onError = (ev: ErrorEvent) => {
    audit('错误', 'window.onerror', `${ev.message} @${ev.filename || '?'}:${ev.lineno || 0}`)
  }
  const onRejection = (ev: PromiseRejectionEvent) => {
    const r = ev.reason as { message?: string } | string
    audit('错误', 'unhandledrejection', typeof r === 'string' ? r : (r && r.message) || String(r))
  }
  window.addEventListener('error', onError)
  window.addEventListener('unhandledrejection', onRejection)

  /* ⑪ 页面生命周期 —— 关窗/刷新前把队列刷完，否则最后几条操作会丢 */
  const onHidden = () => { if (document.visibilityState === 'hidden') flush() }
  document.addEventListener('visibilitychange', onHidden)
  const onPageHide = () => flush()
  window.addEventListener('pagehide', onPageHide)

  /* ⑫ 原生弹窗 —— 这是最典型的「软件反馈」：保存成功/失败、未保存就关闭的拦截提示
   *    全都走 window.alert/confirm。用包装而不是监听（没有对应事件）。
   *    ⚠ 必须在调用原始函数**之前**记：alert 会阻塞，日志得先入队。 */
  const rawAlert = window.alert
  const rawConfirm = window.confirm
  const rawPrompt = window.prompt
  window.alert = function (msg?: unknown) {
    audit('反馈', '弹窗(alert)', String(msg == null ? '' : msg))
    return rawAlert.call(window, msg as string)
  }
  window.confirm = function (msg?: string) {
    audit('反馈', '询问框(confirm) 弹出', String(msg == null ? '' : msg))
    const r = rawConfirm.call(window, msg)
    audit('反馈', '询问框(confirm) 结果', r ? '确定' : '取消')
    return r
  }
  window.prompt = function (msg?: string, def?: string) {
    audit('反馈', '输入框(prompt) 弹出', String(msg == null ? '' : msg))
    const r = rawPrompt.call(window, msg, def)
    audit('反馈', '输入框(prompt) 结果', r == null ? '取消' : '值="' + r + '"')
    return r
  }

  return () => {
    document.removeEventListener('pointerdown', onPointerDown, opts)
    document.removeEventListener('click', onClick, opts)
    document.removeEventListener('dblclick', onDblClick, opts)
    document.removeEventListener('contextmenu', onCtxMenu, opts)
    document.removeEventListener('keydown', onKeyDown, optsActive)
    document.removeEventListener('input', onInput, opts)
    document.removeEventListener('change', onChange, opts)
    document.removeEventListener('wheel', onWheel, opts)
    document.removeEventListener('scroll', onScroll, opts)
    document.removeEventListener('dragstart', onDragStart, opts)
    document.removeEventListener('drop', onDrop, opts)
    document.removeEventListener('focusin', onFocusIn, opts)
    window.removeEventListener('error', onError)
    window.removeEventListener('unhandledrejection', onRejection)
    document.removeEventListener('visibilitychange', onHidden)
    window.removeEventListener('pagehide', onPageHide)
    window.alert = rawAlert
    window.confirm = rawConfirm
    window.prompt = rawPrompt
    installed = false
  }
}
