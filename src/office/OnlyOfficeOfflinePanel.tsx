import { useEffect, useRef, useState } from 'react'

/**
 * OnlyOffice 离线编辑器面板（C 端 · 无 Document Server）
 *
 * 架构（对标 ONLYOFFICE 桌面版）：
 *   vendors/onlyoffice/ 静态资产 ←── Electron 主进程本地静态服务（/oo-assets/）
 *   编辑器 iframe ←── api.js（DocsAPI.DocEditor）
 *   docx 读入 ── x2t（wasm，本地）→ 内部格式 → 编辑
 *   保存 ── x2t 回写 docx ── IPC（office:read/write）→ 模板库/绑定产物落盘
 *
 * 数据关联（与自研域体系对接）：
 *   sdt 内容控件 = 我们的 {{field}}/{{each}} 域（sdtify 产物 OnlyOffice 原生识别）
 *   填充走 Automation API：editor.callCommand/executeMethod 读写控件
 *   定制 UI 走 OnlyOffice 插件（侧栏 iframe 承载 RefPanel 数据引用面板）
 *
 * ⚠ 骨架状态：等 vendors/onlyoffice 资产就位（scripts/extract-onlyoffice-assets.mjs）
 *   + Electron 侧 /oo-assets 静态服务后启用。当前不挂进任何 Tab。
 */

export type OnlyOfficeOfflineProps = {
  /** 要编辑的 docx 字节（绑定产物/导入模板） */
  bytes: Uint8Array
  /** 文件名（title/历史 key 用） */
  name: string
  /** 保存回写（x2t 落盘后的 docx 字节）——上层写回模板库/editedBytes */
  onSaved: (bytes: Uint8Array) => void
  onClose?: () => void
}


let scriptSeq = 0

/** 编辑器资产基址：开发=本地静态服务；打包=随 app 资产。由主进程注入 window.__OO_BASE__ */
const assetBase = () => (window as unknown as { __OO_BASE__?: string }).__OO_BASE__ ?? '/oo-assets'

export function OnlyOfficeOfflinePanel(props: OnlyOfficeOfflineProps) {
  const { bytes, name, onSaved, onClose } = props
  const holder = useRef<HTMLDivElement>(null)
  const editor = useRef<{ destroyEditor: () => void } | null>(null)
  const [err, setErr] = useState('')
  const [loading, setLoading] = useState(true)
  const [missing, setMissing] = useState(false)

  useEffect(() => {
    let dead = false
    let script: HTMLScriptElement | null = null
    const mountId = `oo-offline-${++scriptSeq}-${Date.now().toString(36)}`

    async function boot() {
      try {
        setErr('')
        setLoading(true)
        // 1. 资产探测：api.js 可达才继续（不可达 = 资产未抽取，给出明确指引）
        const base = assetBase()
        const probe = await fetch(`${base}/web-apps/apps/api/documents/api.js`, { method: 'HEAD' }).catch(() => null)
        if (!probe || !probe.ok) {
          setMissing(true)
          setLoading(false)
          return
        }
        // 2. docx 字节落盘（主进程写临时工作区，编辑器经 /oo-work/<file> 回读）
        const r = await fetch('/office-offline/open', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name, dataBase64: bytesToB64(bytes) }),
        })
        if (!r.ok) throw new Error(`工作区写入失败（${r.status}）`)
        const { workUrl, key } = await r.json()
        // 3. 加载 api.js 并挂编辑器
        script = document.createElement('script')
        script.src = `${base}/web-apps/apps/api/documents/api.js`
        script.onload = () => {
          if (dead || !window.DocsAPI || !holder.current) return
          editor.current = new window.DocsAPI(mountId, {
            type: 'desktop',
            width: '100%',
            height: '100%',
            events: {
              onDocumentReady: () => setLoading(false),
              onError: (e: unknown) => { setErr(`编辑器错误：${JSON.stringify(e).slice(0, 160)}`); setLoading(false) },
            },
            document: {
              fileType: 'docx',
              key,
              title: name,
              url: workUrl,
              permissions: { edit: true, download: true, print: true },
            },
            editorConfig: {
              mode: 'edit',
              customization: {
                compactHeader: true,
                hideRightMenu: true,
                // 插件位：我们的数据引用侧栏（RefPanel）从这里挂
                plugins: { pluginsData: [] },
              },
            },
          } as unknown as Record<string, unknown>)
        }
        document.head.appendChild(script)
      } catch (e) {
        setErr(e instanceof Error ? e.message : String(e))
        setLoading(false)
      }
    }
    void boot()
    return () => {
      dead = true
      editor.current?.destroyEditor()
      script?.remove()
    }
  }, [bytes, name])

  /** 保存：编辑器内 Ctrl+S / 按钮 → x2t 回写工作区 → 拉回字节 */
  const save = async () => {
    try {
      const r = await fetch('/office-offline/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name }),
      })
      if (!r.ok) throw new Error(`保存回读失败（${r.status}）`)
      const { dataBase64 } = await r.json()
      onSaved(b64ToBytes(dataBase64))
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e))
    }
  }

  return (
    <div className="oo-offline" style={{ position: 'absolute', inset: 0, display: 'flex', flexDirection: 'column' }}>
      <div style={{ display: 'flex', gap: 8, padding: '6px 10px', borderBottom: '1px solid #333' }}>
        <strong>{name}</strong>
        <button className="bar-btn primary" onClick={save}>保存回写</button>
        <button className="bar-btn" onClick={onClose}>关闭</button>
        {loading && <span>编辑器加载中…</span>}
        {err && <span style={{ color: '#e3b341' }}>{err}</span>}
      </div>
      {missing ? (
        <div style={{ padding: 24, lineHeight: 1.8 }}>
          <p>OnlyOffice 离线资产未就位。三步：</p>
          <ol>
            <li>安装 ONLYOFFICE 桌面版（v9.4+）：github.com/ONLYOFFICE/DesktopEditors/releases</li>
            <li>node scripts/extract-onlyoffice-assets.mjs "安装目录"</li>
            <li>重启应用（Electron 主进程挂 /oo-assets 静态服务）</li>
          </ol>
        </div>
      ) : (
        <div ref={holder} id={`oo-offline-mount`} style={{ flex: 1 }} />
      )}
    </div>
  )
}

/* ── base64 工具（与 WordReportPanel openInEditor 同实现） ── */
function bytesToB64(bytes: Uint8Array): string {
  let bin = ''
  const CH = 0x8000
  for (let i = 0; i < bytes.length; i += CH) bin += String.fromCharCode(...bytes.subarray(i, i + CH))
  return btoa(bin)
}
function b64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64)
  const out = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
  return out
}
