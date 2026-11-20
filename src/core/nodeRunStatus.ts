/**
 * 节点运行状态 · 轻量事件 store（2026-10-23 用户指令
 * 「在每个节点上方显示关键过程 参照其他画布软件」——n8n/Dify 风格状态徽章）
 *
 * 非持久（内存态，随页面刷新清空）——只服务「运行过程可视化」，
 * 不进 config（不落盘、不触发画布重建）。
 * useSyncExternalStore 订阅，徽章组件按 nodeId 精准重渲染。
 */
import { useSyncExternalStore } from 'react'

export type RunState = 'running' | 'ok' | 'error'
export interface NodeRunStatus {
  state: RunState
  /** 徽章文案（如「转换中…」「PDF 24KB · 412ms」「内容已损坏」） */
  label: string
  at: number
}

const map = new Map<string, NodeRunStatus>()
const listeners = new Set<() => void>()

function emit() {
  for (const l of listeners) l()
}

export function setNodeRunStatus (id: string, state: RunState, label: string) {
  map.set(id, { state, label, at: Date.now() })
  /* ★ 2026-11-09 全局运行状态条（画布底部居中）：顺手记最近一次事件，RunStatusBar 订阅 */
  lastEvent = { id, state, label, at: Date.now() }
  snapVer++
  emit()
}

/* ── 全局运行状态条（2026-11-09 用户指令「在运行的时候 页面最下方 画布居中
 *    显示操作/运行状态 执行出现 执行完毕 淡隐」）────────────────
 *  与节点徽章同源（同一 store）：任何 setNodeRunStatus 调用都驱动底部条。 */
export interface LastRunEvent {
  id: string
  state: RunState
  label: string
  at: number
}
let lastEvent: LastRunEvent | null = null

export function useLastRunEvent (): LastRunEvent | null {
  return useSyncExternalStore(
    subscribe,
    () => lastEvent,
    () => lastEvent,
  )
}

/* ── 主/次分层（2026-11-09 用户指令「处理后往右跑主进程 再上面 次进程在下面」）──
 *  主进程 = 被点执行的节点（runPdfMerge / runNodeToPdf 入口 setPrimaryRunNode）；
 *  次进程 = 其他 running 中的节点（上游转换/装载等子任务）。 */
let primaryId: string | null = null

export function setPrimaryRunNode (id: string | null) {
  primaryId = id
  snapVer++
  emit()
}

export interface RunBarItem { id: string; state: RunState; label: string; at: number }
export interface RunBarSnapshot {
  primary: RunBarItem | null
  /** running 中的次进程（新的在前） */
  secondary: RunBarItem[]
}
const EMPTY_SNAP: RunBarSnapshot = { primary: null, secondary: [] }
let snapVer = 0
let snapCache: { ver: number; snap: RunBarSnapshot } | null = null

function buildSnap (): RunBarSnapshot {
  let primary: RunBarItem | null = null
  const secondary: RunBarItem[] = []
  for (const [id, s] of map) {
    if (primaryId && id === primaryId) primary = { id, ...s }
    else if (s.state === 'running') secondary.push({ id, ...s })
  }
  secondary.sort((a, b) => b.at - a.at)
  return { primary, secondary }
}

export function useRunBarSnapshot (): RunBarSnapshot {
  return useSyncExternalStore(
    subscribe,
    () => {
      if (!snapCache || snapCache.ver !== snapVer) snapCache = { ver: snapVer, snap: buildSnap() }
      return snapCache.snap
    },
    () => EMPTY_SNAP,
  )
}

export function clearNodeRunStatus (id: string) {
  if (map.delete(id)) { snapVer++; emit() }
}

/** ok 状态停留 ms 后自动淡出（running/error 常驻直到下次变更/清除） */
export function setNodeRunStatusAuto (id: string, state: RunState, label: string, okTtlMs = 8000) {
  setNodeRunStatus(id, state, label)
  if (state === 'ok') {
    setTimeout(() => {
      const cur = map.get(id)
      if (cur && cur.state === 'ok' && cur.label === label) clearNodeRunStatus(id)
    }, okTtlMs)
  }
}

function subscribe (cb: () => void) {
  listeners.add(cb)
  return () => listeners.delete(cb)
}

export function useNodeRunStatus (id: string): NodeRunStatus | undefined {
  return useSyncExternalStore(
    subscribe,
    () => map.get(id),
    () => map.get(id),
  )
}
