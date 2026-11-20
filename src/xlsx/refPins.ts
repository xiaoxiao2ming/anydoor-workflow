/**
 * 批次 1-4 引用栏改版 · 优先参数 / 星标参数 的全局共享存储。
 *
 * 语义（用户备忘原话）：
 *   「点单元格 → 点『优先参数』→ 收录进所有引用窗体共享的优先区」
 *   —— 收录的是当前选中格引用（node://id#field）对应的 nodeId + fieldId；
 *   优先区 / 星标区跨所有 Excel 节点的引用栏共享（画布级全局，不属于任何单节点）。
 *
 * 存储选型（备忘「store/localStorage 待定」→ 拍板 localStorage）：
 *   - 这是 UI 收藏偏好，不是节点数据 → 不进 store 快照、不进撤销栈、不进 .wmp 工程；
 *   - 与 refpanel:${nodeId} 开关、data-theme 主题等既有 localStorage 惯例一致。
 *
 * 同步：模块级 listener 集（同页多编辑器实例互通）+ window storage 事件
 * （跨窗口/跨 tab 兜底）。localStorage 不可用（node 测试 / 隐私模式）时
 * 退化为纯内存行为：load 返回空、save 只 emit 不落盘。
 */

export type RefPin = { nodeId: string; fieldId: string; ts: number }

/** 优先参数 storage key（收录自选中引用格） */
export const PRIO_KEY = 'refprio:v1'
/** 星标参数 storage key（参数树上手动收藏） */
export const STAR_KEY = 'refstar:v1'

const listeners = new Set<() => void>()

/** 订阅变化（收录/移除/其它实例写入都会触发）；返回退订函数 */
export function subscribePins(fn: () => void): () => void {
  listeners.add(fn)
  return () => {
    listeners.delete(fn)
  }
}

function emit(): void {
  for (const fn of listeners) fn()
}

function store(): Pick<Storage, 'getItem' | 'setItem'> | undefined {
  try {
    if (typeof localStorage !== 'undefined') return localStorage
  } catch {
    /* 读取 localStorage 本身抛错（沙箱）→ 视为不可用 */
  }
  return undefined
}

/** 读取收藏列表；损坏/非法数据逐项过滤，绝不抛错 */
export function loadPins(key: string): RefPin[] {
  const s = store()
  if (!s) return []
  try {
    const raw = s.getItem(key)
    if (!raw) return []
    const arr = JSON.parse(raw) as unknown
    if (!Array.isArray(arr)) return []
    const out: RefPin[] = []
    for (const it of arr) {
      const o = it as Record<string, unknown>
      if (o && typeof o.nodeId === 'string' && o.nodeId && typeof o.fieldId === 'string' && o.fieldId) {
        out.push({ nodeId: o.nodeId, fieldId: o.fieldId, ts: typeof o.ts === 'number' ? o.ts : 0 })
      }
    }
    return out
  } catch {
    return []
  }
}

/** 落盘并广播（空数组也写——「清空收藏」是合法状态） */
export function savePins(key: string, pins: RefPin[]): void {
  const s = store()
  try {
    s?.setItem(key, JSON.stringify(pins))
  } catch {
    /* 配额满 / 隐私模式 → 静默：本次会话内存态仍可用 */
  }
  emit()
}

export function hasPin(pins: RefPin[], nodeId: string, fieldId: string): boolean {
  return pins.some((p) => p.nodeId === nodeId && p.fieldId === fieldId)
}

/** 收录：已存在则不动（不重复、不重排——保住用户手动排序心智），新项插头部（最新在前） */
export function upsertPin(pins: RefPin[], nodeId: string, fieldId: string, ts?: number): RefPin[] {
  if (hasPin(pins, nodeId, fieldId)) return pins
  return [{ nodeId, fieldId, ts: ts ?? Date.now() }, ...pins]
}

/** 移除：返回新数组（不变式：原数组不动） */
export function dropPin(pins: RefPin[], nodeId: string, fieldId: string): RefPin[] {
  return pins.filter((p) => !(p.nodeId === nodeId && p.fieldId === fieldId))
}

// 跨窗口同步兜底：其它窗口/tab 改了收藏 storage（或 clear）→ 本页订阅者刷新。
// 同页多编辑器实例走上面模块级 listeners，无需 storage 事件。
if (typeof window !== 'undefined' && typeof window.addEventListener === 'function') {
  try {
    window.addEventListener('storage', (e) => {
      if (e.key === null || e.key === PRIO_KEY || e.key === STAR_KEY) emit()
    })
  } catch {
    /* 极端环境无 addEventListener → 降级为仅本页同步 */
  }
}
