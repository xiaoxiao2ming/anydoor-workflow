/**
 * 版本信息（前端唯一出口）
 *
 * 值由 Vite 构建期 define 注入（scripts/app-version.mjs → versionDefine），
 * 源头是根目录 VERSION —— 桌面端与网页端共用同一个版本号。
 *
 * ⚠ vitest / 纯 Node 环境没有 define 注入，直接引用裸变量会抛 ReferenceError，
 *   所以统一走 typeof 守卫。诊断脚本因此也能读到 '0.0.0-dev' 而不是崩掉。
 */
function injected(name: '__APP_VERSION__' | '__APP_CHANNEL__' | '__APP_BUILD__', fallback: string): string {
  try {
    switch (name) {
      case '__APP_VERSION__':
        return typeof __APP_VERSION__ !== 'undefined' ? __APP_VERSION__ : fallback
      case '__APP_CHANNEL__':
        return typeof __APP_CHANNEL__ !== 'undefined' ? __APP_CHANNEL__ : fallback
      default:
        return typeof __APP_BUILD__ !== 'undefined' ? __APP_BUILD__ : fallback
    }
  } catch {
    return fallback
  }
}

/** 形如 0.00.15（与 VERSION 文件逐字一致） */
export const APP_VERSION = injected('__APP_VERSION__', '0.0.0-dev')
/** 'desktop'（Electron）| 'web'（浏览器 / 网页发布） */
export const APP_CHANNEL = injected('__APP_CHANNEL__', 'dev')
/** 构建时间戳 ISO */
export const APP_BUILD = injected('__APP_BUILD__', '')

/** 标题栏/关于用的单行串：v0.00.15 · web */
export const APP_VERSION_LABEL = `v${APP_VERSION}${APP_CHANNEL ? ' · ' + APP_CHANNEL : ''}`
