import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'
import { AppErrorBoundary } from './ui/AppErrorBoundary'
import { installAudit, audit } from './core/audit'
import { ensureCaps } from './core/capabilities'

import 'ag-grid-community/styles/ag-grid.css'
import 'ag-grid-community/styles/ag-theme-quartz.css'
import './index.css'
import './nodes-office.css'

/* ★ 2026-10-17 用户指令「从现在开始，我对软件的每个点击等操作、键盘输入、软件的反馈，
 *   都放入日志」—— 必须在 render 之前装：首屏加载期间的点击/报错同样要留痕。 */
installAudit()
audit('系统', '应用启动', 'main.tsx 开始挂载 React 根节点')

/* ★ 2026-10-17 网页端同步发布：启动即探测服务端能力（/healthz 的 oo 字段）。
 *   桌面端恒为 true；网页部署（没有 OnlyOffice 文档服务）为 false，
 *   OO 节点会显示只读提示而不是无限转圈。探测失败按有能力处理，不新增失败点。 */
void ensureCaps().then((c) => {
  audit('系统', '环境探测', `Office 在线编辑=${c.oo ? '可用' : '不可用'}${c.version ? ' 服务端版本=v' + c.version : ''}`)
})

/* 把「日志到底写在哪」打到控制台（会进 .dev-console.log / 终端）：
 * 排查时不用再猜路径，用户问「日志在哪」也能直接答。 */
void (window as unknown as { api?: { auditPath?: () => Promise<{ dir: string; file: string }> } })
  .api?.auditPath?.()
  .then((r) => { console.log('[审计] 操作日志文件:', r.file) })
  .catch(() => { /* 网页版走 /audit-log，不在渲染层暴露路径 */ })

// 刻意不用 <React.StrictMode>：开发态 double-mount 会把「格内编辑器 mount 时消费待输入字符」
// 一类一次性 effect 搅成竞态（实测丢首字符/清除时灵时不灵）。本项目副作用集中在事件 handler，
// 不依赖 StrictMode 的副作用体检。
ReactDOM.createRoot(document.getElementById('root')!).render(
  <AppErrorBoundary>
    <App />
  </AppErrorBoundary>
)
audit('系统', '应用已渲染', 'React 根节点提交完成')
