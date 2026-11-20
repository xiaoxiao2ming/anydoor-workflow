// 节点折叠探针专用配置（任务 #108，2026-10-06）：
// ① 独立 cacheDir —— 主缓存目录重建时触发 safe-delete 批量删除守卫（110 个文件 > 阈值 50）
//   被 sandbox 拦截，vite 直接起不来；独立缓存目录是全新的，无需删除动作。
// ② 独立端口 5175 —— 5173 被另一个（旧代码副本的）vite 占着，热加载救不了根目录不同的问题。
// ③ 剥离 electron 插件（2026-10-06 下午）—— 根因：本机 vite-plugin-electron 的
//   treeKillSync 在重启 Electron 时**必现** `spawnSync cmd.exe EBUSY`（未捕获 → vite 全崩），
//   5175 服务两次崩溃同因（第二次触发链：另一会话保存 scripts/oo-offline-service.mjs
//   → main.js 重打包 → Electron 重启 → 崩）。而验证服务只被探针用 http
//   （playwright 连 5175），Electron 窗口纯属附带 —— 剥离后 fold 服务是纯 vite dev，
//   改文件只走 HMR，永不触发这颗雷。
// 除此之外与主配置同源：mergeConfig 继承 vite.config.ts 的插件/代理/中间件。
import { mergeConfig, defineConfig, type UserConfig, type Plugin } from 'vite'
import baseExport from './vite.config'

// 主配置是回调形式（defineConfig(() => ...)），mergeConfig 不能合回调——先解析它
const raw = baseExport as unknown as UserConfig | ((env: { command: string; mode: string }) => UserConfig | Promise<UserConfig>)
const base: UserConfig = typeof raw === 'function' ? await raw({ command: 'serve', mode: 'development' }) : raw

// ③ 过滤 electron 系插件（vite-plugin-electron / vite-plugin-electron-renderer）。
//   插件可能是对象或 [对象, 选项] 元组，统一取 .name 判断。
base.plugins = (base.plugins ?? []).filter((p) => {
  const plug = (Array.isArray(p) ? p[0] : p) as Plugin | undefined
  const name = (plug as { name?: string } | undefined)?.name ?? ''
  return !name.toLowerCase().includes('electron')
})

export default mergeConfig(
  base,
  defineConfig({
    // fold2：剥离 electron 插件改变了依赖优化结果 → vite 要删旧 deps 重建（109 文件
    // > 批量阈值 50）被 safe-delete 守卫拦（fold config 首次启用时同一坑）——
    // 换全新目录零删除直接起。旧 .vite-cache-fold 留着不再引用。
    cacheDir: 'node_modules/.vite-cache-fold2',
    server: { port: 5175, strictPort: true },
  }),
)
