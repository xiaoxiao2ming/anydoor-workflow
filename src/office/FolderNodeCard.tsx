/* ─────────────────────────────────────────────────────────────
 * 文件夹节点（2026-10-25 用户指令「作为读取位置和保存位置」）
 *
 * 两用：
 *  - 读模式（folderMode='read'）：列出目录内文件，双击/拖出可加载到画布节点
 *  - 写模式（folderMode='write'）：作为保存输出目录——保存对话框默认指向它
 *
 * config 契约（随画布持久化）：
 *  { folderPath?: string; folderMode?: 'read' | 'write' }
 * ───────────────────────────────────────────────────────────── */
import { useState } from 'react'
import { useBaseStore } from '../core/store'
import { IcFolderOpen, IcRefresh } from '../ui/icons'

interface FolderFile { name: string; size: number; mtime: number }

export function FolderNodeCard({ nodeId }: { nodeId: string }) {
  const cfg = useBaseStore((s) => s.tables[nodeId]?.config) as {
    folderPath?: string
    folderMode?: 'read' | 'write'
  } | undefined
  const path = cfg?.folderPath ?? ''
  const mode: 'read' | 'write' = cfg?.folderMode ?? 'read'
  const [files, setFiles] = useState<FolderFile[] | null>(null)
  const [busy, setBusy] = useState(false)

  const api = (globalThis as {
    api?: {
      folderPickDir?: () => Promise<string | null>
      folderList?: (dir: string) => Promise<{ ok: boolean; files: { name: string; size: number; mtime: number }[]; reason?: string }>
      folderReadFile?: (dir: string, name: string) => Promise<{ ok: boolean; bytes: Uint8Array; reason?: string }>
    }
  }).api

  const pickDir = async () => {
    if (!api?.folderPickDir) { alert('桌面端专用功能（请重启应用获取桥）'); return }
    const d = await api.folderPickDir()
    if (d) useBaseStore.getState().setNodeConfig(nodeId, { folderPath: d })
  }

  const refresh = async () => {
    if (!path || !api?.folderList) return
    setBusy(true)
    try {
      const r = await api.folderList(path)
      if (r.ok) setFiles(r.files.map((f) => ({ name: f.name, size: f.size, mtime: f.mtime })))
      else { setFiles([]); alert('读取目录失败：' + (r.reason ?? '')) }
    } finally { setBusy(false) }
  }

  const extOf = (name: string) => (name.toLowerCase().split('.').pop() ?? '')
  const fmtMB = (b: number) => b > 1024 * 1024 ? (b / 1024 / 1024).toFixed(1) + ' MB' : (b / 1024).toFixed(0) + ' KB'
  const shortPath = path ? (path.length > 38 ? '…' + path.slice(-36) : path) : '未选择文件夹'

  return (
    <div className="pn-root" style={{ gap: 4 }}>
      {/* 头部：路径显示 */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '0 6px' }}>
        <span style={{ fontSize: 12, color: 'var(--c-text2)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', flex: 1 }} title={path || undefined}>
          <span style={{ flexShrink: 0, display: 'inline-flex', color: 'var(--ed-text3, var(--c-text3))' }}><IcFolderOpen size={14} /></span> {shortPath}
        </span>
      </div>
      {/* 模式切换：读/写 */}
      <div style={{ display: 'flex', gap: 4, padding: '0 6px' }}>
        {(['read', 'write'] as const).map((m) => (
          <button key={m} onClick={() => useBaseStore.getState().setNodeConfig(nodeId, { folderMode: m })}
            style={{ fontSize: 11, padding: '2px 10px', borderRadius: 4, cursor: 'pointer',
              border: '1px solid ' + (mode === m ? 'var(--ed-accent-bd, rgba(159,225,203,.42))' : 'var(--c-border)'),
              background: mode === m ? 'var(--ed-accent-bg, rgba(159,225,203,0.14))' : 'transparent',
              color: mode === m ? 'var(--ed-accent, #9fe1cb)' : 'var(--ed-text3, var(--c-text3))' }}>
            {m === 'read' ? '读取位置' : '保存目录'}
          </button>
        ))}
      </div>
      {/* 按钮排（2026-10-25 用户指令「按钮放下面做成正方形」）：31×31 方形圆角 */}
      <div style={{ display: 'flex', gap: 6, padding: '2px 6px' }}>
        <button className="rf-oo-btn" style={{ width: 31, height: 31, padding: 0, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', border: '1px solid var(--c-border)', borderRadius: 6, background: 'transparent', color: 'var(--ed-text3, var(--c-text3))', cursor: 'pointer', flexShrink: 0 }} title="选择文件夹" onClick={() => void pickDir()}><IcFolderOpen size={17} /></button>
        <button className="rf-oo-btn" style={{ width: 31, height: 31, padding: 0, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', border: '1px solid var(--c-border)', borderRadius: 6, background: 'transparent', color: 'var(--ed-text3, var(--c-text3))', cursor: 'pointer', flexShrink: 0 }} title="刷新列表" onClick={() => void refresh()}><IcRefresh size={17} /></button>
      </div>
      {/* 读模式：文件列表 */}
      {mode === 'read' && (
        <>
          <div style={{ fontSize: 11, color: 'var(--ed-text3, var(--c-text3))', padding: '0 6px' }}>{files ? files.length + ' 个文件（可拖出到画布/节点导入）' : ''}</div>
          {files && files.length > 0 && (
            <div style={{ maxHeight: 220, overflowY: 'auto' }}>
              {files.map((f, i) => (
                <div key={f.name + i} className="pn-row nodrag nopan" style={{ margin: '2px 6px' }} title={f.name + '（拖出到画布空白建节点 / 拖到已有节点置入）'}
                  draggable
                  onDragStart={(eD) => {
                    try {
                      eD.dataTransfer.setData('application/x-cb-folder-file', JSON.stringify({ dir: path, name: f.name }))
                      eD.dataTransfer.effectAllowed = 'copy'
                    } catch { /* ignore */ }
                  }}>
                  <span className="pn-idx">{i + 1}</span>
                  <span className={'pm-fmt ' + (extOf(f.name) === 'docx' ? 'pm-fmt-word' : extOf(f.name) === 'xlsx' ? 'pm-fmt-sheet' : extOf(f.name) === 'pdf' ? 'pm-fmt-pdf' : '')} style={extOf(f.name) === 'docx' || extOf(f.name) === 'xlsx' || extOf(f.name) === 'pdf' ? undefined : { background: '#666' }}>{extOf(f.name) === 'docx' ? 'W' : extOf(f.name) === 'xlsx' ? 'X' : extOf(f.name) === 'pdf' ? 'P' : '●'}</span>
                  <span className="pn-nm" style={{ position: 'static', background: 'none', padding: 0, fontSize: 12, color: 'var(--c-text2)' }}>{f.name}</span>
                  <span style={{ marginLeft: 'auto', fontSize: 11, color: 'var(--ed-text3, var(--c-text3))', flexShrink: 0 }}>{fmtMB(f.size)}</span>
                </div>
              ))}
            </div>
          )}
          {files && files.length === 0 && <div style={{ fontSize: 11.5, color: 'var(--ed-text3, var(--c-text3))', padding: '4px 6px' }}>目录为空或不可读</div>}
        </>
      )}
      {/* 写模式：说明 */}
      {mode === 'write' && (
        <div style={{ fontSize: 11.5, color: 'var(--ed-text3, var(--c-text3))', padding: '4px 6px' }}>
          {path ? '保存工程时，另存为对话框将默认打开此目录' : '选择目录后，保存时将默认输出到这里'}
        </div>
      )}
    </div>
  )
}
