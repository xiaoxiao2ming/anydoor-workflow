import { useEffect, useRef, useState } from 'react'
import { clearLogs, getLogs, subscribeLog, type LogEntry } from './logBus'

/**
 * ★ 画布底部日志面板（2026-10-22 用户指令「画布下方增加log显示面板」）：
 *   数据源 = 全局日志总线 logBus（OO 节点 addDiag 与全局 JS 错误汇流）。
 *   形态仿 OO 诊断抽屉（Consolas 11px、时间戳前缀、收起/清空/复制三钮），
 *   配色走 --c-* 主题变量（与全 UI 主题体系一致，不硬编码深底）。
 *   定位：absolute 挂在 .canvas-wrap 内，不挤占画布 flex 布局。
 *   ★ 开关在顶部工具带（CanvasRibbon 的 LogDockToggle，cr-icon-btn 同款尺寸），
 *     经 localStorage('logdockOpen') + 自定义事件 logdock-toggle 同步——
 *     原右上角浮动钮已按用户指令删除（2026-10-22「下面的按钮放到上面，尺寸协调」）。
 */
export function LogDock() {
  const [logs, setLogs] = useState<readonly LogEntry[]>(getLogs())
  const [open, setOpen] = useState(() => localStorage.getItem('logdockOpen') !== '0')
  const [errCount, setErrCount] = useState(0)
  const seenErr = useRef(0)

  useEffect(() => {
    return subscribeLog((e) => {
      if (e == null) { setLogs(getLogs()); return }
      setLogs(getLogs())
      if (e.kind === 'err') setErrCount(++seenErr.current)
    })
  }, [])

  /* 开关来源改为工具带按钮（logdock-toggle 事件驱动） */
  useEffect(() => {
    const onToggle = (ev: Event) => {
      const next = (ev as CustomEvent<boolean>).detail
      setOpen(next)
      setErrCount(0)
      seenErr.current = 0
    }
    window.addEventListener('logdock-toggle', onToggle)
    return () => window.removeEventListener('logdock-toggle', onToggle)
  }, [])

  /* ★ 2026-10-22 用户「诊断日志打开时，右下角的缩略框不要被遮挡」：
     开合状态写进 .canvas-wrap 的 data-logdock，CSS 据此抬高右下角
     缩放面板（zoom-panel）与缩略图，避开面板高度 */
  useEffect(() => {
    const wrap = document.querySelector('.canvas-wrap')
    if (wrap) wrap.setAttribute('data-logdock', open ? 'open' : 'closed')
    return () => { if (wrap) wrap.removeAttribute('data-logdock') }
  }, [open])

  const copyAll = () => {
    try {
      const text = [...logs].reverse().map((l) => `[${l.ts}] ${l.msg}`).join('\n')
      void navigator.clipboard.writeText(text)
    } catch { /* ignore */ }
  }

  const collapse = () => {
    setOpen(false)
    localStorage.setItem('logdockOpen', '0')
  }

  return (
    open && (
      <div className="logdock" data-testid="logdock">
        <div className="logdock-bar">
          <span className="logdock-title">诊断日志</span>
          <span className="logdock-count">{logs.length}</span>
          <span style={{ flex: 1 }} />
          <button className="logdock-btn" onClick={() => setLogs(getLogs())} title="刷新">刷新</button>
          <button className="logdock-btn" onClick={() => { clearLogs(); setLogs(getLogs()); setErrCount(0); seenErr.current = 0 }}>清空</button>
          <button className="logdock-btn" onClick={copyAll} title="复制全部日志">复制</button>
          <button className="logdock-btn" onClick={collapse} title="收起面板（顶部工具带 ⓘ 重新展开）">收起</button>
        </div>
        <pre className="logdock-body">
          {logs.length === 0
            ? '（暂无日志 —— OO 节点诊断与全局错误会自动汇入）'
            : logs.slice(-100).map((l, i) => (
                <div key={`${l.at}_${i}`} className={`logdock-line k-${l.kind}`}>
                  [{l.ts}] {l.msg}
                </div>
              ))}
        </pre>
      </div>
    )
  )
}
