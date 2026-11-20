/**
 * 给 vite-node 用的最小配置。
 *
 * 项目主 vite.config.ts 里有 vite-plugin-electron-renderer，它会把 fs / path 等
 * Node 内置模块别名成 .vite-electron-renderer/*.mjs。那套 shim 是给渲染进程用的
 * （内部走 require），在纯 Node 的 ESM 下会直接抛
 * "require is not defined in ES module scope"。
 *
 * 跑 scripts/ 下的命令行脚本时用这份配置，绕开插件别名。
 */
export default { plugins: [] }
