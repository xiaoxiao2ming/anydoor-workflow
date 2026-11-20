/**
 * 参数引用面板 UI 状态共享（2026-10-11 · 用户指令「最大化的参数引用和常规大小的窗体内容同步」）
 *
 * 背景：同一份「参数引用」有两个宿主——
 *   · 常规窗体 = React 浮卡 RefPanel（窗外左侧，headMax=false）
 *   · 最大化   = /poc 注入 OO 左栏的内嵌面板（批注位，headMax=true）
 * 两处各自维护 q（搜索词）/ openIds（展开的节点）⇒ 切换最大化来回内容对不上。
 *
 * 方案：模块级 store 单例（存活于整个会话，不随组件卸载丢失）+ 双向镜像：
 *   React RefPanel   读写 store，变更经 OnlyOfficeNodeEditor 的订阅推送 __ooRefUi → /poc
 *   /poc 内嵌面板    收 __ooRefUi 应用；自己变更时 postMessage __ooRefUi 回宿主（fromPoc=true
 *                    不再回推，杜绝乒乓）
 */

export interface RefUiState {
  q: string
  openIds: string[]
}

let state: RefUiState = { q: '', openIds: [] }
const subs = new Set<(s: RefUiState) => void>()
/** 防乒乓：/poc 回传的应用过程置 true，subscribers 只读不回推 */
let applyingFromPoc = false

export function getRefUi(): RefUiState {
  return { q: state.q, openIds: [...state.openIds] }
}

export function setRefUi(patch: Partial<RefUiState>, fromPoc = false): void {
  if (typeof patch.q === 'string') state.q = patch.q
  if (Array.isArray(patch.openIds)) state.openIds = [...patch.openIds]
  for (const f of subs) {
    try { f(state) } catch { /* 订阅者异常不扩散 */ }
  }
  applyingFromPoc = fromPoc
}

/** 是否正在应用 /poc 回传（OnlyOfficeNodeEditor 判断要不要回推给 /poc） */
export function isApplyingFromPoc(): boolean {
  return applyingFromPoc
}

export function subscribeRefUi(f: (s: RefUiState) => void): () => void {
  subs.add(f)
  return () => { subs.delete(f) }
}
