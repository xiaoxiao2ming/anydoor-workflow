/**
 * ★ UI 偏好读写入口（2026-10-20）：
 *   桌面端 = userData/ui-settings.json（经 preload 同步快照 + settings:set 增量保存），
 *   网页端 = localStorage 兜底。界面主题/画布背景等一律走这里，别再直连 localStorage。
 */
export function prefGet(key: string): string | null {
  const api = (window as unknown as { api?: { prefs?: Record<string, string> } }).api
  const prefs = api?.prefs
  if (prefs && key in prefs) return prefs[key]
  try { return localStorage.getItem(key) } catch { return null }
}

export function prefSet(key: string, value: string | null) {
  const api = (window as unknown as { api?: { settingsSet?: (patch: Record<string, string | null>) => void } }).api
  try {
    if (value === null) localStorage.removeItem(key)
    else localStorage.setItem(key, value)
  } catch { /* 网页端隐私模式等 */ }
  api?.settingsSet?.({ [key]: value })
}
