import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import electron from 'vite-plugin-electron'
import renderer from 'vite-plugin-electron-renderer'
import { join } from 'node:path'
import { startAppServer } from './app-server.mjs'
import { versionDefine } from './scripts/app-version.mjs'

// WorkBuddy 自身是 Electron 应用，环境里带着 ELECTRON_RUN_AS_NODE=1，
// 会导致 vite-plugin-electron 拉起的 electron 当纯 Node 跑、不弹窗口。
// 在此清除，保证 `npm run dev` 能弹出真正的 Electron 窗口。
delete process.env.ELECTRON_RUN_AS_NODE

export default defineConfig(async ({ command }) => {
  // 仅开发模式在此内嵌 app-server：vite 只负责页面与 HMR，/api 经 proxy 转发过去，
  // 这样 Electron 窗口与浏览器网页版读写同一份 server-data。
  // 端口传 0 由系统分配，不占用固定端口、不与已运行的网页版服务冲突。
  // build 模式不起服务 —— 否则 listen 会挂住进程导致 vite build 永不退出
  //（生产模式的服务由 electron/main.ts 在主进程内启动）。
  // ★ proxy 项的实际形状：vite 的 ProxyOptions 还支持 ws（WebSocket 透传）与 xfwd（透传原始 host），
  //   下面 /oo-assets、/doc、/cache 都真的在用这两个字段。所以类型必须写全，
  //   否则 tsc 会报 TS2353（对象字面量出现类型里没有的属性）——
  //   ★ 这类错误只卡 C 端（npm run build 会跑 tsc --noEmit），网页端 dev:web 不做类型检查，
  //   属于典型的「两端开发进度不同步」来源。
  type ProxyOpt = { target: string; changeOrigin: boolean; ws?: boolean; xfwd?: boolean }
  let proxy: Record<string, ProxyOpt> | undefined
  let apiPort = 0
  if (command === 'serve') {
    // ★ port 必须是 0（系统分配），不能写死 3799。
    //   app-server.mjs 的契约明写「vite 开发代理：startAppServer({root:null, port:0})」，
    //   且本文件上方注释也写着"端口传 0 由系统分配，不与已运行的网页版服务冲突"，
    //   但下面实际传的是硬编码 3799 —— 代码与自己的注释矛盾。
    //   症状：只要本机已有另一个 dev server 起着（本例 master 侧5173 带着 3799），
    //   本worktree 的 dev server 启动即 EADDRINUSE 崩掉，且崩在 vite 起来之前，
    //   表现为「端口传 5174 也起不来」，很容易误判成端口被占/配置写错。
    const api = await startAppServer({
      root: null,
      dataDir: join(process.cwd(), 'server-data', 'pages'),
      label: 'app-server(dev)',
      port: 0, // ← 契约要求0；实际端口由 api.port 读回
    })
    apiPort = api.port
    // /api（页面数据）与 /office（OnlyOffice 上传/配置/文档服务）都要转过去。
    // 少了 /office，点「在线编辑」会在 vite 上拿 404，表现为"点了没反应"。
    // ★ 下面所有 target 都必须用 api.port 实测值，不能写死端口：
    //   写死后，一旦本机已有另一个 dev server 占用同一端口就EADDRINUSE 崩，
    //   或代理静默指向别的服务（更糟：页面读写的是另一份数据）。
    const oo = `http://127.0.0.1:${apiPort}`
    proxy = {
      // ★ ws+xfwd 必须开：OO 编辑器用 EIO=4 transport=websocket 直连
      //   /oo-assets/doc/<key>/c/（socket 基路径跟 api.js 同前缀）——没 ws 的话 upgrade
      //   在 vite 处被掐（"WebSocket is closed before the connection is established"）；
      //   xfwd 把浏览器原始 host 透传给 oo-core → DS，让 DS 生成的 cache 下载 URL 指回同源。
      '/oo-assets': { target: oo, changeOrigin: true, ws: true, xfwd: true },
      '/work': { target: oo, changeOrigin: true },
      '/bridge.js': { target: oo, changeOrigin: true },
      '/office-offline': { target: oo, changeOrigin: true },
      // ★ OO 编辑器同源化：/poc 页与其依赖全走 dev server → oo（消灭跨源 LocalStartOpen 卡死）
      '/poc': { target: oo, changeOrigin: true },
      '/report': { target: oo, changeOrigin: true },
      '/convert': { target: oo, changeOrigin: true },
      '/dumplog': { target: oo, changeOrigin: true },
      '/fonts': { target: oo, changeOrigin: true },
      '/sysfonts': { target: oo, changeOrigin: true },
      '/cache': { target: oo, changeOrigin: true, xfwd: true },
      // ★ 实时控制台查看器（SSE 推流 + HTML 单页），解决 CDP 不可靠时看不到渲染层/iframe 报错
      '/devconsole': { target: oo, changeOrigin: true },
      // /doc 协作通道：http + ws upgrade 都要（xfwd 让 DocService 生成同源下载 URL）
      '/doc': { target: 'http://127.0.0.1:3798', changeOrigin: true, xfwd: true, ws: true },
      '/api': { target: oo, changeOrigin: false },
      // ★ 能力探测（2026-10-17）：dev 模式下前端 fetch('/healthz') 必须打到 app-server，
      //   否则拿到 SPA fallback 的 HTML、JSON 解析失败 → 每次都记一条「探测失败」日志。
      '/healthz': { target: oo, changeOrigin: false },
      // 用精确正则：'^/version$' 避免把静态文件 /version.json 也转走
      '^/version$': { target: oo, changeOrigin: false },
      // ★ 操作审计日志（2026-10-17）：dev 模式下页面由 vite 提供，
      //   渲染层 POST /audit-log 必须转发到 app-server 才能落盘（否则 404 到 index.html）。
      '/audit-log': { target: oo, changeOrigin: false },

      '/office': { target: oo, changeOrigin: false },
    }
  }

  return {
    // ★ 版本号单一真相源：根目录 VERSION → 前端可直接用 __APP_VERSION__ 显示。
    //   桌面端渠道名 'desktop'，网页端 'web'（vite.web.config.ts 同款）——
    //   两边构建产物一眼能分清出自哪条链路，但版本号始终相同。
    define: versionDefine('desktop'),
    // 依赖预构建缓存走独立目录：WorkBuddy 沙箱对"一次删 >50 个文件"有批量确认守卫，
    // vite 缓存失效时的整目录清理会被拦/超时 —— 换目录后旧缓存自然废弃，无需删除。
    cacheDir: 'node_modules/.vite-cache-oo',
    // ★ dev 预构建 target 必须与 build.target 同步提到 es2022：
    //   mupdf（PDF 内核）用 top-level await，optimizeDeps 默认 chrome87 会直接编译失败
    //   → dev server 起来后秒退（2026-11-09 实测，报 "Top-level await is not available"）。
    optimizeDeps: {
      esbuildOptions: { target: 'es2022' },
    },
    build: {
      // ★ 2026-10-24：target 提到 es2022（Electron 44 = Chromium 130+ 完全支持）——
      //   mupdf wasm 模块用 top-level await，默认 chrome87 目标编译失败。
      target: 'es2022',
      // ★ 同 vite.web.config.ts：emptyOutDir 必须关（2026-10-06 补，主配置一直漏着）。
      //   vite 清空outDir 用整目录删除，而本机沙箱有批量删除守卫
      //   —— dist/assets 动辄两百多个文件，一次删除就被拦：
      //     [safe-delete][SAFE_DELETE_BULK_CONFIRM_REQUIRED] {"count":234,"threshold":50}
      //   症状：`npm run build`（**C 端出包的唯一入口**）在 vite 阶段直接报 error。
      //   ⇒ 典型的又一处「两端不对称」：web 端配置早已这么修，
      //     只有 C 端用的主配置在踩这个坑。
      emptyOutDir: false,
    },
    // ★ CB_NO_ELECTRON=1：纯浏览器调试渲染层（不拉起 Electron 窗口）——
    //   配合 playwright/CDP 排查组件问题；单实例锁会让 dev Electron 在客户端开着时起不来。
    plugins: [
      react(),
      ...(process.env.CB_NO_ELECTRON === '1' ? [] : [
      electron([
        {
          entry: 'electron/main.ts',
          vite: {
            build: {
              outDir: 'dist-electron',
              // ★ node-llama-cpp（聪明模型原生模块）保持 external：运行时从 app/node_modules require，
              //   否则 rollup 解析动态 import 报 invalid id，main.js 构建失败静默回退旧产物
              rollupOptions: { external: [/^node-llama-cpp/, /^onnxruntime-node/] }
            }
          }
        },
        {
          entry: 'electron/preload.ts',
          onstart(options) {
            options.reload()
          },
          vite: {
            build: {
              outDir: 'dist-electron',
              // preload 由 Electron 以 CommonJS 方式加载（sandbox 环境不支持 ESM）。
              // vite-plugin-electron 按 package.json 的 "type" 硬编码输出格式（type:module → es），
              // 所以必须显式覆盖 lib.formats 为 cjs，否则窗口会报：
              // "Unable to load preload script ... Cannot use import statement outside a module"
              lib: {
                entry: 'electron/preload.ts',
                formats: ['cjs'],
                fileName: () => 'preload.js',
              },
            },
          },
        }
      ]),
      ]),
      renderer()
    ],
    server: {
      // 固定端口 5173 时：本机已有另一个 dev server 时会静默递增到 5174/5175…
      // 属预期行为（多个 worktree 并行开发），故strictPort 刻意不开。
      port: 5173,
      // ★ 必须可双栈监听（host: '::' 在 Windows 上同时接受 127.0.0.1 与 [::1]）。
      //   本机 localhost 优先解析到 IPv6 ::1，若只绑 '127.0.0.1'，任何跑在 localhost 源上的
      //   页面（用户浏览器 / OO 编辑器 iframe）拿相对路径下载文档都会 ERR_CONNECTION_REFUSED，
      //   在 OO 编辑器中表现为「Download failed」弹窗（2026-09-30 实锤：dumplog parentOrigin=
      //   http://localhost:<port> 的会话文档下载必挂）。
      host: '::',
      proxy,
    },
  }
})
