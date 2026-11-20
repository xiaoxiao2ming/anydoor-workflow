/**
 * OO 文档预加载管理器（2026-10-07 用户指令：文档节点「预加载按钮」）
 *
 * 流程：预加载 → 后台上传样例/模板 → 隐藏挂载 /poc iframe → documentReady 后
 *       「冻结」（iframe 脱离 DOM：Chromium 暂停其全部 JS/渲染/定时器，现场保留、CPU≈0）
 * 节点打开 → OnlyOfficeNodeEditor 取走冻结 iframe 挂回自己容器（同 document 移动不重载）
 *           → 编辑器从冻结现场原地恢复 → **0 等待**。
 *
 * 状态存取：Map<nodeId, Frozen>（运行时态，不持久化——应用重启后预加载失效属正常，
 * 节点照常走原打开流程）。
 */

export type OoPreloadKind = 'ooword' | 'oosheet' | 'pdf'
export type OoPreloadState = 'idle' | 'loading' | 'frozen'

export interface FrozenEntry {
  iframe: HTMLIFrameElement
  workName: string
  kind: OoPreloadKind
  fileName: string
}

const frozenMap = new Map<string, FrozenEntry>()
const loadingSet = new Set<string>()
let host: HTMLDivElement | null = null

function ensureHost(): HTMLDivElement {
  if (!host) {
    host = document.createElement('div')
    host.id = 'oo-preload-host'
    // 屏幕外常驻容器：iframe 在此完成加载（不可见但不销毁）。
    // 不能 display:none —— 部分浏览器会推迟 iframe 加载；屏幕外 + opacity 0 最稳。
    host.style.cssText = 'position:fixed;left:-99999px;top:0;width:1280px;height:800px;opacity:0;pointer-events:none;overflow:hidden'
    document.body.appendChild(host)
  }
  return host
}

/** 当前预加载状态（节点按钮显隐/文案用） */
export function preloadState(nodeId: string): OoPreloadState {
  if (frozenMap.has(nodeId)) return 'frozen'
  if (loadingSet.has(nodeId)) return 'loading'
  return 'idle'
}

/** 节点打开时取走冻结实例（取走后 Map 删除——实例归编辑器组件所有） */
export function takeFrozen(nodeId: string): FrozenEntry | null {
  const f = frozenMap.get(nodeId)
  if (f) frozenMap.delete(nodeId)
  return f || null
}

/**
 * 重新冻结（2026-10-07 用户：「打开 关闭再打开……不要重复打开」）：
 * 节点编辑器**关闭**时，把运行中的编辑器 iframe 移回隐藏 host（脱离 DOM →
 * Chromium 暂停其全部执行，现场保留）并放回 frozenMap——下次打开原地恢复。
 * ★ 冻结循环的编辑器修改天然保留在冻结现场（DOM/JS 状态都在），不丢失。
 */
export function refreeze(nodeId: string, entry: FrozenEntry): void {
  ensureHost().appendChild(entry.iframe) // append 会自动从编辑器容器摘出
  frozenMap.set(nodeId, entry)
}

/** 是否有冻结实例（编辑器挂载时判断走「恢复」还是「正常加载」） */
export function hasFrozen(nodeId: string): boolean {
  return frozenMap.has(nodeId)
}

/**
 * 等待预加载完成并取走冻结实例（2026-10-07 bug 修复：
 * 预加载「进行中」用户就点开节点 → 编辑器必须**等待接管**而非另开一份文档）。
 * 轮询 200ms；超时返回 null（调用方回退正常加载流程）。
 */
export function waitFrozen(nodeId: string, timeoutMs = 60000): Promise<FrozenEntry | null> {
  return new Promise((resolve) => {
    const t0 = Date.now()
    const tick = () => {
      const f = frozenMap.get(nodeId)
      if (f) { frozenMap.delete(nodeId); resolve(f); return }
      if (!loadingSet.has(nodeId) || Date.now() - t0 > timeoutMs) { resolve(null); return }
      setTimeout(tick, 200)
    }
    tick()
  })
}

/** 等待某 iframe 的 poc 页上报 documentReady（最长 timeoutMs） */
function waitDocumentReady(iframe: HTMLIFrameElement, timeoutMs: number): Promise<boolean> {
  return new Promise((resolve) => {
    const onMsg = (e: MessageEvent) => {
      const d = e.data as { __ooStatus?: string } | null
      if (d && typeof d === 'object' && typeof d.__ooStatus === 'string' && /documentReady/.test(d.__ooStatus)) {
        // 校验消息来自这个 iframe（e.source）
        if (!iframe.contentWindow || e.source === iframe.contentWindow) {
          window.removeEventListener('message', onMsg)
          resolve(true)
        }
      }
    }
    window.addEventListener('message', onMsg)
    setTimeout(() => { window.removeEventListener('message', onMsg); resolve(false) }, timeoutMs)
  })
}


function b64(u8: Uint8Array): string {
  let bin = ''
  const CH = 0x8000
  for (let i = 0; i < u8.length; i += CH) bin += String.fromCharCode(...u8.subarray(i, i + CH))
  return btoa(bin)
}

/**
 * 预加载：后台上传 + 打开 + 冻结。完成后可被 takeFrozen 消费。
 * 字节来源优先级与编辑器 boot 一致：显式 bytes → 模板库 → 默认样例。
 * 抛错 = 预加载失败（按钮回退 idle；节点正常打开流程不受影响）。
 */
export async function preloadOo(opts: {
  nodeId: string
  kind: OoPreloadKind
  bytes?: Uint8Array | null
  fileName?: string
  templateId?: string
}): Promise<void> {
  const { nodeId, kind } = opts
  if (frozenMap.has(nodeId) || loadingSet.has(nodeId)) return
  loadingSet.add(nodeId)
  try {
    // ① 字节来源
    let u8: Uint8Array | null = opts.bytes && opts.bytes.length ? opts.bytes : null
    let fileName = opts.fileName || ''
    let ext = kind === 'oosheet' ? 'xlsx' : 'docx'
    if (!u8 && opts.templateId) {
      const { getTemplateStore } = await import('../docx/templateStore')
      const saved = await getTemplateStore().read(opts.templateId)
      if (saved && saved.length) { u8 = saved; fileName = fileName || '文档' }
    }
    if (!u8) {
      /* ★ 2026-10-12 用户指令「新建节点打开空白新建文件」：无内容不预加载（样例种入已废除）。
       *   空白文档本就秒开，放弃本次预加载即可——节点打开时走 boot 的空白路径（OO 官方新建）。 */
      throw new Error('无内容可预载（新建节点走空白打开）')
    }
    // ② 上传工作区
    const workName = 'oo-' + kind + '-' + Date.now().toString(36) + '.' + ext
    const up = await fetch('/office-offline/open', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: workName, dataBase64: b64(u8) }),
    }).then((x) => x.json())
    if (!up || !up.ok) throw new Error('上传失败: ' + ((up && up.err) || 'unknown'))
    // ③ 隐藏挂载 poc iframe（★ 铺满 host：引擎按初始尺寸建画布，预加载尺寸要与
    //    目标容器同量级，恢复后 resize 才不至于布局错乱）
    const iframe = document.createElement('iframe')
    iframe.src = '/poc?src=' + workName
    iframe.style.cssText = 'width:100%;height:100%;border:none;display:block;background:#87CEEB'
    ensureHost().appendChild(iframe)
    // ④ 等 documentReady
    const ok = await waitDocumentReady(iframe, 45000)
    if (!ok) throw new Error('预加载超时（45s 未 documentReady）')
    // ⑤ 冻结：脱离 DOM——Chromium 暂停其全部执行，现场保留
    frozenMap.set(nodeId, { iframe, workName, kind, fileName })
    iframe.remove()
  } finally {
    loadingSet.delete(nodeId)
  }
}

/** 丢弃冻结实例（节点删除/数据替换时调用，防泄漏） */
export function discardFrozen(nodeId: string): void {
  const f = frozenMap.get(nodeId)
  if (f) { try { f.iframe.remove() } catch { /* ignore */ } frozenMap.delete(nodeId) }
}
