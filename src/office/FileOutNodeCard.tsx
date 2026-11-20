/* ─────────────────────────────────────────────────────────────
 * 文件输出节点卡片（2026-10-26 用户指令「输出节点 可以存储文件到固定位置」）
 *
 * 功能：连线上游产物（PDF 文件组 / 合并历史 / 替换页历史）→ 一键保存到固定目录。
 * config 契约（随画布持久化，均为轻量索引——字节不进 config）：
 *  { inputs?: string[]                // 上游节点（连线语义，App.handleConnect 写）
 *    outDir?: string                  // 固定输出目录
 *    lastSaved?: { name; at; size }[] // 最近保存记录（≤10 条，新在前） }
 *
 * 规范依据：节点规范.md 十（视觉）/五（NodeBtnBar 装载条）/十二（设置面板三段式）
 * ───────────────────────────────────────────────────────────── */
import { useState } from 'react'
import { useBaseStore } from '../core/store'
import { getBlob } from '../core/pdfBlobStore'
import { pushLog } from '../ui/logBus'
import { setNodeRunStatus, setNodeRunStatusAuto } from '../core/nodeRunStatus'
import { NodeBtnBar, NodeSqBtn } from '../ui/NodeBtnBar'
import { GrpLockButton, NodeLockButtons } from '../CanvasParts'
import { IcDownload, IcFolderOpen } from '../ui/icons'

type SavedRec = { name: string; at: number; size: number }

const u8FromB64 = (b64: string): Uint8Array => {
  const bin = atob(b64)
  const u8 = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
  return u8
}
const fmtSize = (b: number) => (b > 1024 * 1024 ? (b / 1024 / 1024).toFixed(1) + ' MB' : (b / 1024).toFixed(0) + ' KB')
const fmtTime = (at: number) => {
  const d = new Date(at)
  const p = (n: number) => String(n).padStart(2, '0')
  return `${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`
}
/** 名称确保 .pdf 后缀（上游名可能不带扩展名） */
const ensurePdf = (name: string) => (/\.pdf$/i.test(name) ? name : name + '.pdf')

export function FileOutNodeCard({ nodeId }: { nodeId: string }) {
  const cfg = useBaseStore((s) => s.tables[nodeId]?.config) as
    | { inputs?: string[]; outDir?: string; lastSaved?: SavedRec[] }
    | undefined
  const inputs = cfg?.inputs ?? []
  const dir = cfg?.outDir ?? ''
  const lastSaved = cfg?.lastSaved ?? []
  const [busy, setBusy] = useState(false)

  const api = (globalThis as {
    api?: {
      folderPickDir?: () => Promise<string | null>
      wfcWriteFile?: (filePath: string, bytes: Uint8Array) => Promise<{ ok: boolean; size: number }>
    }
  }).api

  const pickDir = async () => {
    if (!api?.folderPickDir) { setNodeRunStatus(nodeId, 'error', '桌面桥缺失（重启应用）'); return }
    const d = await api.folderPickDir()
    if (d) {
      useBaseStore.getState().pushHistory()
      useBaseStore.getState().setNodeConfig(nodeId, { outDir: d })
      pushLog(`[文件输出] 输出目录已设置：${d}`, 'ok')
    }
  }

  /** 收集上游产物字节（pdfmerge/pdfreplace=最新产物；pdf/pdfslim=全部份按行序） */
  const collectFiles = async (): Promise<{ name: string; bytes: Uint8Array }[]> => {
    const st = useBaseStore.getState()
    const out: { name: string; bytes: Uint8Array }[] = []
    for (const uid of inputs) {
      const up = st.tables[uid]
      if (!up || up.deleted) continue
      const uc = (up.config ?? {}) as Record<string, unknown>
      if (up.kind === 'pdfmerge') {
        const h = ((uc.pdfHistory ?? []) as { id: string; name?: string; b64?: string; at?: number }[])[0]
        if (!h) continue
        const bytes = h.b64 ? u8FromB64(h.b64) : await getBlob('pdfhist-' + h.id + '-' + (h.at ?? 0))
        if (bytes) out.push({ name: ensurePdf(h.name || up.name || 'merged'), bytes })
      } else if (up.kind === 'pdfreplace') {
        const hh = ((uc.replaceHistory ?? []) as { id: string; name?: string; b64?: string }[])[0]
        if (!hh) continue
        const bytes = hh.b64 ? u8FromB64(hh.b64) : await getBlob('rephist-' + hh.id)
        if (bytes) out.push({ name: ensurePdf(hh.name || up.name || 'replaced'), bytes })
      } else if (up.kind === 'pdf' || up.kind === 'pdfslim') {
        const items = (uc.pdfItems ?? []) as { id: string; name: string; inDb?: boolean; b64?: string }[]
        const order = (uc.pdfFileOrder ?? null) as string[] | null
        const ordered = order ? [...items].sort((a, b) => order.indexOf(a.id) - order.indexOf(b.id)) : items
        for (const it of ordered) {
          const bytes = it.b64 ? u8FromB64(it.b64) : await getBlob('pdfitem-' + it.id)
          if (bytes) out.push({ name: ensurePdf(it.name || it.id), bytes })
        }
      }
    }
    return out
  }

  const run = async () => {
    if (busy) return
    if (!dir) { setNodeRunStatus(nodeId, 'error', '未设输出目录（先点文件夹钮）'); return }
    if (!api?.wfcWriteFile) { setNodeRunStatus(nodeId, 'error', '桌面桥缺失（重启应用）'); return }
    setBusy(true)
    setNodeRunStatus(nodeId, 'running', '保存中…')
    const t0 = Date.now()
    try {
      const files = await collectFiles()
      if (!files.length) {
        setNodeRunStatus(nodeId, 'error', '上游无产物（连 PDF/合并/替换节点）')
        return
      }
      const saved: SavedRec[] = []
      const stamp = new Date().toTimeString().slice(0, 8).replace(/:/g, '')
      let fail = 0
      for (const f of files) {
        let name = f.name
        try {
          await api.wfcWriteFile(dir + '/' + name, f.bytes)
        } catch {
          try {
            name = name.replace(/(\.[^.]*)?$/, '-' + stamp + '$1')
            await api.wfcWriteFile(dir + '/' + name, f.bytes)
          } catch { fail++; continue }
        }
        saved.push({ name, at: Date.now(), size: f.bytes.byteLength })
      }
      const st = useBaseStore.getState()
      st.pushHistory()
      st.setNodeConfig(nodeId, { lastSaved: [...saved, ...lastSaved].slice(0, 10) })
      if (!saved.length) setNodeRunStatus(nodeId, 'error', '写入失败（目录不可写？）')
      else {
        setNodeRunStatusAuto(nodeId, 'ok', `${saved.length} 个文件 · ${((Date.now() - t0) / 1000).toFixed(1)}s`)
        pushLog(`[文件输出] 已保存 ${saved.length} 个文件到 ${dir}${fail ? `（${fail} 个失败）` : ''}`, fail ? 'warn' : 'ok')
      }
    } finally {
      setBusy(false)
    }
  }

  const shortPath = dir ? (dir.length > 34 ? '…' + dir.slice(-32) : dir) : ''
  const upstreamNames = inputs.map((uid) => useBaseStore.getState().tables[uid]?.name).filter(Boolean) as string[]

  return (
    <div className="pn-root fo-root">
      {/* 目录行（点击 = 选固定输出目录） */}
      <button
        className="pn-row fo-dir-row nodrag nopan"
        title={dir || '点击选择固定输出目录'}
        onClick={(e) => { e.stopPropagation(); void pickDir() }}
      >
        <span className="fo-dir-ic"><IcFolderOpen size={14} /></span>
        <span className="fo-dir-path">{shortPath || '点击选择输出目录…'}</span>
      </button>
      {/* 上游行（连线即引用，保存时读取） */}
      <div className="pn-row" title={upstreamNames.join('、')}>
        <span className="pn-idx">↑</span>
        <span className="props-file-nm">{inputs.length ? `上游 ${inputs.length} 个：${upstreamNames.slice(0, 2).join('、')}${upstreamNames.length > 2 ? ' 等' : ''}` : '未连上游——把产物节点连到左口'}</span>
      </div>
      {/* ★ 2026-10-26 用户截图指正「排列有点松散」：废除 min-height+弹性占位钉底——
          卡片高度随内容，装载条紧跟内容行（与其他节点按钮行上下衔接同一规则） */}
      <NodeBtnBar className="nb-bar--flat">
        <NodeLockButtons nodeId={nodeId} />
        <GrpLockButton nodeId={nodeId} />
        <NodeSqBtn icon={<IcFolderOpen size={19} />} title="选择输出目录" onClick={() => void pickDir()} />
        <NodeSqBtn icon={<IcDownload size={19} />} title="保存到输出目录" onClick={() => void run()} busy={busy} disabled={!inputs.length} />
      </NodeBtnBar>
      {/* 最近保存（≤5 行，倒序） */}
      {lastSaved.slice(0, 5).map((r) => (
        <div key={r.name + r.at} className="pn-row" title={`${r.name}（${fmtSize(r.size)}）`}>
          <span className="pn-loc pn-loc--in">存</span>
          <span className="props-file-nm">{r.name}</span>
          <span className="fo-saved-meta">{fmtSize(r.size)} · {fmtTime(r.at)}</span>
        </div>
      ))}
    </div>
  )
}
