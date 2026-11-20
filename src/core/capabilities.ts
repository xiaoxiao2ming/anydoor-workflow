/**
 * 运行环境能力探测（2026-10-17 网页端同步发布）
 *
 * 同一份前端要跑在两种地方：
 *   - Electron 桌面端：带 OnlyOffice 离线文档服务（Windows 二进制 DS + x2t）
 *   - 网页端（浏览器 / 公网部署）：没有 DS，Office 节点无法在线编辑
 *
 * 服务端在 /healthz 里如实报告 `oo` 能力（APP_DISABLE_OO=1 时为 false）。
 * 前端只问一次、全局缓存，OO 节点据此决定「打开编辑器」还是「显示只读提示」，
 * 避免用户在网页端点开一个永远转圈的黑框。
 *
 * 探测失败（老版本服务端没有 /healthz、或纯静态托管）时按「有能力」处理 ——
 * 保持桌面端既有行为，不让新逻辑成为新的失败点。
 */
export interface Caps {
  /** 是否具备 OnlyOffice 在线编辑能力 */
  oo: boolean
  /** 服务端上报的版本（网页端） */
  version: string
  /** 是否探测过（未探测时 OO 按可用处理） */
  probed: boolean
}

let cached: Caps = { oo: true, version: '', probed: false }
let pending: Promise<Caps> | null = null

/** 同步读取当前已知能力（不触发探测） */
export function capsSync(): Caps {
  return cached
}

/** 探测一次并缓存（并发调用共用同一个请求） */
export async function ensureCaps(): Promise<Caps> {
  if (cached.probed) return cached
  if (pending) return pending
  pending = (async () => {
    try {
      const r = await fetch('/healthz', { headers: { Accept: 'application/json' } })
      if (!r.ok) throw new Error('healthz ' + r.status)
      const d = (await r.json()) as { oo?: boolean; version?: string }
      cached = { oo: d.oo !== false, version: String(d.version || ''), probed: true }
    } catch {
      // 服务端没有探针（纯静态托管/旧版本）→ 维持「有能力」，由编辑器自身报错兜底
      cached = { oo: true, version: '', probed: true }
    } finally {
      pending = null
    }
    return cached
  })()
  return pending
}

/** 测试/特殊情况重置（仅诊断脚本使用） */
export function resetCaps(): void {
  cached = { oo: true, version: '', probed: false }
  pending = null
}
