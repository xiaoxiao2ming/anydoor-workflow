import React from 'react'

/** ★ 2026-11-07 用户反馈「画布软件 打开 不显示」：
 *  此前渲染层任何一个 uncaught 异常都会让 React 卸载整棵树 —— 窗口还在、内容全白，
 *  用户既看不到原因也没有自救入口。这里兜住首屏异常：显示错误摘要 + 重载按钮，
 *  并把堆栈打到控制台（进 .dev-console.log / 审计日志），便于定位而不是盲猜。
 *  注意：只兜「渲染异常」，业务层自己的 try/catch 仍按原样工作。 */
type Props = { children: React.ReactNode }
type State = { error: Error | null }

export class AppErrorBoundary extends React.Component<Props, State> {
  state: State = { error: null }

  static getDerivedStateFromError(error: Error): State {
    return { error }
  }

  componentDidCatch(error: Error, info: React.ErrorInfo) {
    console.error('[界面异常] 渲染失败，已兜底显示排障面板：', error, info.componentStack)
  }

  render() {
    const err = this.state.error
    if (!err) return this.props.children
    const stack = String(err.stack || err.message || err).slice(0, 1200)
    return (
      <div
        style={{
          position: 'fixed',
          inset: 0,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          background: 'var(--c-bg, #1b1d21)',
          color: 'var(--c-text, #e8eaed)',
          font: '13px/1.65 system-ui, -apple-system, "Segoe UI", sans-serif',
          padding: 24,
        }}
      >
        <div style={{ maxWidth: 720, width: '100%' }}>
          <div style={{ fontSize: 17, fontWeight: 600, marginBottom: 8 }}>界面加载出错了</div>
          <div style={{ opacity: 0.72, marginBottom: 14 }}>
            画布数据没有丢失，这是界面渲染时的异常。先把下面的信息反馈出来，或直接重新加载。
          </div>
          <pre
            style={{
              margin: '0 0 16px',
              padding: 12,
              maxHeight: 320,
              overflow: 'auto',
              whiteSpace: 'pre-wrap',
              wordBreak: 'break-all',
              background: 'rgba(127,127,127,.12)',
              border: '1px solid rgba(127,127,127,.25)',
              borderRadius: 8,
              fontSize: 12,
            }}
          >
            {stack}
          </pre>
          <div style={{ display: 'flex', gap: 10 }}>
            <button
              onClick={() => window.location.reload()}
              style={{
                padding: '8px 18px',
                borderRadius: 8,
                border: '1px solid var(--c-accent, #4c8dff)',
                background: 'var(--c-accent, #4c8dff)',
                color: '#fff',
                cursor: 'pointer',
                font: 'inherit',
              }}
            >
              重新加载
            </button>
            <button
              onClick={() => this.setState({ error: null })}
              style={{
                padding: '8px 18px',
                borderRadius: 8,
                border: '1px solid rgba(127,127,127,.4)',
                background: 'transparent',
                color: 'inherit',
                cursor: 'pointer',
                font: 'inherit',
              }}
            >
              忽略并继续
            </button>
          </div>
        </div>
      </div>
    )
  }
}
