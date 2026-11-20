import { useEffect, useRef, useState } from 'react'

/**
 * OnlyOffice 编辑器面板 —— C 端（Electron）与 html 端同一组件。
 *
 * 数据流：
 *   本组件 ──POST /office/config──▶ app-server（签发 JWT，密钥不出服务端）
 *          ◀──{ apiJs, config }────
 *   浏览器动态加载 DS 的 api.js → new DocsAPI.DocEditor(...) → iframe 内嵌编辑器
 *   DS 保存时 POST config.editorConfig.callbackUrl → app-server 下载写回文件
 *
 * 网络铁律：config 里的 document.url / callbackUrl 是给 **DS 容器** 回访用的，
 * 不是给浏览器用的。C 端 DS 跑在 NAS 时，app-server 必须以局域网 IP 访问
 * （host 传 0.0.0.0 启动），否则 DS 拉不到文档、回调也回不来。
 */

export type OfficeDocType = 'word' | 'cell' | 'slide'

declare global {
  interface Window {
    DocsAPI?: new (element: string, config: Record<string, unknown>) => {
      destroyEditor: () => void
    }
  }
}

let scriptSeq = 0

export function OnlyOfficePanel(props: {
  /** 工作文件名（server-data/office 内），如 `报告.docx` */
  file: string
  /** app-server 地址；C 端用局域网地址（如 http://your-app-server:5198），html 端留空走同源 */
  docApiBase?: string
  /** OnlyOffice Document Server 地址，如 http://your-onlyoffice-server:8080 */
  dsUrl: string
  /**
   * 覆盖 config 里 document.url / callbackUrl 的 host 部分。
   * C 端 DS 部署在远端时必填（如 your-app-server:5198）——DS 回访的是它，
   * 不是浏览器用的 localhost。服务端部署（DS 与文档服务同机）时不用传。
   */
  hostOverride?: string
  /** DS 保存回调完成（文件已写回）——在这里刷新 docx-preview 预览 */
  onSaved?: () => void
  onClose?: () => void
}) {
  const holder = useRef<HTMLDivElement>(null)
  const editor = useRef<{ destroyEditor: () => void } | null>(null)
  const [err, setErr] = useState('')
  const [loading, setLoading] = useState(true)
  const [dirty, setDirty] = useState(false)

  useEffect(() => {
    let dead = false
    let script: HTMLScriptElement | null = null
    const mountId = `oo-mount-${++scriptSeq}-${Date.now().toString(36)}`

    async function boot() {
      try {
        setErr('')
        setLoading(true)
        const base = (props.docApiBase ?? '').replace(/\/+$/, '')
        const r = await fetch(`${base}/office/config`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            file: props.file,
            dsUrl: props.dsUrl,
            // C 端 DS 在 NAS 时，浏览器可能用 localhost 访问 app-server，
            // 但 DS 回访必须用局域网 IP —— 由宿主显式传入覆盖
            hostOverride: props.hostOverride,
          }),
        })
        if (!r.ok) throw new Error(`获取编辑器配置失败（${r.status}）——检查 app-server 是否在跑`)
        const cfg = (await r.json()) as { apiJs: string; config: Record<string, unknown> }
        if (dead) return

        // DS 前端 SDK 只需加载一次；加载地址随 dsUrl 变化时换 script
        if (!window.DocsAPI) {
          await new Promise<void>((resolve, reject) => {
            script = document.createElement('script')
            script.src = cfg.apiJs
            script.async = true
            script.onload = () => resolve()
            script.onerror = () =>
              reject(new Error(`Document Server 不可达（${cfg.apiJs}）——检查 dsUrl、DS 是否启动、以及跨机器网络`))
            document.head.appendChild(script)
          })
        }
        if (dead || !holder.current || !window.DocsAPI) return

        // DS 会把 placeholder 整个替换成 iframe，destroy 后无法复用 —— 每次重建
        holder.current.innerHTML = ''
        const mount = document.createElement('div')
        mount.id = mountId
        holder.current.appendChild(mount)

        editor.current = new window.DocsAPI(mountId, {
          ...cfg.config,
          events: {
            onDocumentStateChange: (e: { data?: boolean }) => {
              // 编辑中 data=true；保存完成/无改动 data=false
              if (e?.data === true) setDirty(true)
              else if (e?.data === false) {
                setDirty(false)
                props.onSaved?.()
              }
            },
            onError: () => setErr('编辑器内部报错，详情看 DS 容器日志（docker logs onlyoffice-ds）'),
          },
        })
        setLoading(false)
      } catch (e) {
        if (!dead) {
          setErr(e instanceof Error ? e.message : String(e))
          setLoading(false)
        }
      }
    }

    boot()
    return () => {
      dead = true
      try {
        editor.current?.destroyEditor()
      } catch {
        /* 组件卸载时 DS iframe 可能已自毁 */
      }
      editor.current = null
      script?.remove()
    }
    // props.file/dsUrl 变化即重建编辑器；onSaved 等回调不参与依赖
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.file, props.dsUrl, props.docApiBase, props.hostOverride])

  return (
    <div className="oo-wrap">
      <div className="oo-bar">
        <span className="oo-title">OnlyOffice 编辑 · {props.file}</span>
        {dirty && <span className="oo-dirty">有未保存修改</span>}
        <span className="oo-sp" />
        <button className="oo-btn" onClick={props.onClose}>
          关闭编辑器
        </button>
      </div>
      {err && <div className="oo-err">{err}</div>}
      {loading && !err && <div className="oo-loading">正在连接 Document Server…</div>}
      <div ref={holder} className="oo-holder" />
    </div>
  )
}
