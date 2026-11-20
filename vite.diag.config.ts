import { defineConfig, type ConfigEnv } from 'vite'
import type { UserConfig } from 'vite'
import base from './vite.config'

/**
 * 诊断专用 dev 配置（_diag_* 探针用）：
 *   · 端口固定 5174 —— 不与 5173 主 dev server 抢（另一会话/用户窗口在用）
 *   · deps 缓存独立目录 —— 沙箱对「一次删 >50 文件」有 bulk 守卫，
 *     vite 重优化时要清的整目录缓存会直接把 server 起崩；换目录旧缓存自然废弃
 *   · 其余（内嵌 app-server / OO proxy / electron 插件）全部继承主配置
 */
export default defineConfig(async (env: ConfigEnv) => {
  const c = (await (base as unknown as (e: ConfigEnv) => Promise<UserConfig>)(env)) as UserConfig
  return {
    ...c,
    cacheDir: 'node_modules/.vite-cache-diag',
    server: {
      ...(c.server ?? {}),
      port: 5174,
      strictPort: true,
    },
  }
})
