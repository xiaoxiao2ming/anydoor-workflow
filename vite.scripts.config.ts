import { defineConfig } from 'vite'

/**
 * 给 `vite-node scripts/*.mjs` 用的极简配置。
 *
 * 为什么不能复用主 vite.config.ts：主配置挂了 vite-plugin-electron-renderer，
 * 它把 `fs`/`path` 等 Node 内置模块别名到 CJS shim（给渲染进程用）。
 * 在 vite-node 里跑纯 Node 脚本时，这个 shim 会炸：
 *   ReferenceError: require is not defined in ES module scope
 *
 * 脚本只在 Node 里跑，不需要 react / electron / app-server 任何东西。
 * 这里连插件都不装，解析快、出错面小。
 */
export default defineConfig({
  plugins: [],
  // 别去读 .env、别做 html 入口探测，省掉无谓的副作用
  logLevel: 'warn',
})
