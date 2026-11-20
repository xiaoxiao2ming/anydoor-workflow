import { useEffect, useState } from 'react'
import { createPortal } from 'react-dom'
import { OnlyOfficePanel } from './OnlyOfficePanel'

/**
 * OnlyOffice 在线编辑弹层（Word / Excel / PPT 共用）。
 *
 * 原先这段「DS 地址 + 回访地址 + 弹层」是内联在 WordReportPanel 里的，
 * 加 Excel 侧时若不抽出来就是复制一份 —— 两份 DS 配置的 localStorage key、
 * 校验、错误文案迟早会漂移。所以统一收在这里。
 *
 * 网络铁律（见 docs/OnlyOffice集成部署.md）：
 *   config 里的 document.url / callbackUrl 是给 **DS 容器**回访的，不是给浏览器的。
 *   C 端 DS 跑在 NAS 时，"回访地址"必须填本机局域网 IP:5198。
 */

const DS_KEY = 'client-base.office.ds'
const HOST_KEY = 'client-base.office.host'

export function getDsUrl(): string {
  return localStorage.getItem(DS_KEY) ?? ''
}
export function getHostOverride(): string {
  return localStorage.getItem(HOST_KEY) ?? ''
}

export function OnlyOfficeModal({
  file,
  onSaved,
  onClose,
  hint,
}: {
  /** 工作文件名（server-data/office 内），如 `报告.docx` */
  file: string
  /** DS 保存回调完成（文件已写回） */
  onSaved?: () => void
  onClose: () => void
  /** 面板特有的补充说明 */
  hint?: React.ReactNode
}) {
  const [ds, setDs] = useState(getDsUrl)
  const [host, setHost] = useState(getHostOverride)

  // 换个文件（或换节点）时重建，避免残留上一个编辑器的状态
  useEffect(() => {
    setDs(getDsUrl())
    setHost(getHostOverride())
  }, [file])

  return createPortal(
    <div className="oo-modal" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="oo-modal-box">
        <div className="oo-settings">
          <label>
            Document Server
            <input
              value={ds}
              placeholder="http://your-onlyoffice-server:8080（你的 OnlyOffice 服务地址）"
              onChange={(e) => {
                setDs(e.target.value)
                localStorage.setItem(DS_KEY, e.target.value)
              }}
            />
          </label>
          <label>
            回访地址
            <input
              value={host}
              placeholder="C 端必填：本机局域网 IP:5198（DS 容器回访文档与回调用）"
              onChange={(e) => {
                setHost(e.target.value)
                localStorage.setItem(HOST_KEY, e.target.value)
              }}
            />
          </label>
        </div>
        {hint && <div className="oo-hint">{hint}</div>}
        {ds.trim() ? (
          <OnlyOfficePanel
            file={file}
            dsUrl={ds.trim()}
            hostOverride={host.trim() || undefined}
            onSaved={onSaved}
            onClose={onClose}
          />
        ) : (
          <div className="oo-err">
            先填 Document Server 地址。部署方法见 docs/OnlyOffice集成部署.md（NAS Docker 一步起）。
          </div>
        )}
      </div>
    </div>,
    document.body,
  )
}
