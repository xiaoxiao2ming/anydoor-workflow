import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { versionDefine } from './scripts/app-version.mjs'

// 渲染层（网页端 / 浏览器预览）专用配置：
//   构建：npm run build:web  → scripts/build-web.mjs 调这个配置产出 dist-web
//   开发：npm run dev:web    → 页面由 vite 提供，/api /office 代理到独立 app-server
export default defineConfig({
  plugins: [react()],
  cacheDir: process.env.VITE_CACHE_DIR || 'node_modules/.vite',
  // 相对路径：dist-web 可以挂在任意子路径下（反向代理 /s/xxx/ 也能用）
  base: './',
  // ★ 版本号与桌面端同源（根目录 VERSION），渠道标记 web
  define: versionDefine('web'),
  build: {
    outDir: 'dist-web',
    // 不清目录：WorkBuddy 的 safe-delete shim 会拦 vite 的 emptyDir（MSYS 路径被当相对路径拒绝）
    emptyOutDir: false,
  },
  server: {
    host: '127.0.0.1',
    port: 5173,
    // /api 与 /office 都转给独立 app-server。参数是位置式的，root 随便给个目录即可：
    //   node app-server.mjs dist-web 5198
    // 没有 /api，页面初始化 fetch /api/pages 会拿到 SPA fallback 的 index.html，
    // JSON 解析失败、画布永远是空的。
    // 没有 /office，「在线编辑」会上传到 vite 自己身上拿 404 —— 表现为点了没反应。
    //   （app-server 直接托管 dist-web 时是同源，不需要代理；但 vite dev 必须配。）
    proxy: {
      '/api': 'http://127.0.0.1:5198',
      '/office': 'http://127.0.0.1:5198',
      // 操作审计日志（浏览器端写盘通道）：不配这条会 404 到 index.html
      '/audit-log': 'http://127.0.0.1:5198',
    },
  },
})
