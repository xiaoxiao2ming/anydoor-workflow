/**
 * 照片/照片组节点 · 行式列表（2026-10-23 用户指令
 * 「增加一个承载照片或者照片组的节点」+「导入后以列表的形式。类似于PDF合并」）
 *
 * 结构（与 PDF 合并节点同语言）：
 *   ① 导入行：📷 多选导入 + 张数徽章
 *   ② 照片行：每行一张（序号 + 小缩略图 + 文件名 + 拖把）
 *      ——拖拽悬停换位（其它行跟随让位）、落位闪绿，排序持久 config.photoOrder（id 序）
 *   行 hover 出删除；照片 base64 存 config.photoItems 随画布走。
 * 下游连 PDF 合并节点：每张图一页（A4 等比适配）——页序 = 行序。
 */
import { useEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { useViewport } from '@xyflow/react'
import { useBaseStore } from '../core/store'
import { putContentAddressed, getBlob, makeThumbnail, hasBlob } from '../core/pdfBlobStore'
import { pushLog } from '../ui/logBus'
import { IcCamera } from '../ui/icons'

export interface PhotoItem { id: string; name: string; type: string; b64?: string; photoDbId?: string; thumbDbId?: string }

const PHOTO_MAX = 24
const PHOTO_MAX_BYTES = 8 * 1024 * 1024

function b64Of (buf: ArrayBuffer): string {
  const u8 = new Uint8Array(buf)
  let s = ''
  const CHUNK = 0x8000
  for (let i = 0; i < u8.length; i += CHUNK) s += String.fromCharCode(...u8.subarray(i, i + CHUNK))
  return btoa(s)
}

/** photoItems 按 photoOrder（id 序）排列——拖拽排序与删除的单一真相源 */
export function orderedPhotos (items: PhotoItem[], order: string[] | undefined | null): PhotoItem[] {
  if (!order || !order.length) return items
  const byId = new Map(items.map((it) => [it.id, it]))
  const head = order.map((id) => byId.get(id)).filter((it): it is PhotoItem => !!it)
  const tail = items.filter((it) => !order.includes(it.id))
  return [...head, ...tail]
}

/** 照片导入共享逻辑（2026-10-23 抽取）：节点卡片拖放 / 📷 按钮 / 画布锁定组行的
 *  PhotoImportButton 三处共用——校验（jpg/png ≤8MB ≤24 张）与持久化单一真相源。 */
export async function importPhotoFiles (
  nodeId: string,
  files: FileList | null,
): Promise<{ added: number; skipped: number; total: number }> {
  const st = useBaseStore.getState()
  const cfg = (st.tables[nodeId]?.config ?? {}) as { photoItems?: PhotoItem[]; photoOrder?: string[] }
  const raw = cfg.photoItems ?? []
  const next: PhotoItem[] = [...raw]
  let skipped = 0
  for (const f of Array.from(files ?? [])) {
    if (next.length >= PHOTO_MAX) { skipped++; continue }
    const isImg = /^image\/(jpeg|png)$/i.test(f.type)
    if (!isImg || f.size > PHOTO_MAX_BYTES) { skipped++; continue }
    try {
      const buf = await f.arrayBuffer()
      next.push({
        id: 'ph-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7),
        name: f.name,
        type: /png/i.test(f.type) ? 'png' : 'jpg',
        ...await (async () => {
          const { hashBytes } = await import('../core/pdfBlobStore')
          const key = 'photo-' + await hashBytes(buf)
          if (await hasBlob(key)) {
            pushLog(`[照片] ${f.name}：与已有文件内容相同，已合并引用（不重复存储）`)
            return { photoDbId: key }
          }
          return { photoDbId: await putContentAddressed('photo', buf) }
        })(),
        thumbDbId: await (async () => {
          const tb = await makeThumbnail(buf)
          const bytes = new Uint8Array(await tb.arrayBuffer())
          const key = 'photothumb-' + await (await import('../core/pdfBlobStore')).hashBytes(bytes)
          const { putBlob } = await import('../core/pdfBlobStore')
          await putBlob(key, bytes)
          return key
        })(),
      })
    } catch { skipped++ }
  }
  st.pushHistory()
  st.setNodeConfig(nodeId, { photoItems: next, photoOrder: (cfg.photoOrder ?? []).concat(next.slice(raw.length).map((it) => it.id)) })
  pushLog(`照片导入：+${next.length - raw.length} 张（共 ${next.length}）${skipped ? `，跳过 ${skipped} 个非 jpg/png 或超 8MB` : ''}`, next.length > raw.length ? 'ok' : 'warn')
  return { added: next.length - raw.length, skipped, total: next.length }
}

/** 导入按钮（2026-10-23 用户指令「放在下面 做成圆角矩形长方形按钮，
 *  和锁定组按钮并排 大小一致」）：31px 高与 GrpLockButton 同语言，
 *  宽度按内容（长方形），渲染在节点首行锁定组旁。 */
export function PhotoImportButton ({ nodeId }: { nodeId: string }) {
  const fileRef = useRef<HTMLInputElement>(null)
  const [busy, setBusy] = useState(false)
  /* ★ 统一数据锁（2026-10-26）：锁定时禁导入（通用 dataLocked || OO 遗留 ooDataLocked） */
  const lcfg = useBaseStore((s) => s.tables[nodeId]?.config) as { dataLocked?: boolean; ooDataLocked?: boolean } | undefined
  const locked = !!(lcfg?.dataLocked || lcfg?.ooDataLocked)
  return (
    <button
      className="pn-import-btn nodrag nopan"
      disabled={busy || locked}
      title={locked ? '数据已锁定：先解锁再导入' : '导入照片（可多选，jpg/png，随画布保存）'}
      onClick={() => fileRef.current?.click()}
    >
      <IcCamera size={19} />
      <input
        ref={fileRef}
        type="file"
        accept="image/jpeg,image/png"
        multiple
        style={{ display: 'none' }}
        onChange={async (e) => {
          setBusy(true)
          try { await importPhotoFiles(nodeId, e.target.files) } finally { setBusy(false); e.target.value = '' }
        }}
      />
    </button>
  )
}

/* ★ 2026-10-24 照片显示 URL 化：photoDbId → object URL（模块级缓存——同指纹同 URL，
   同一文件在多个节点显示也只解码一次）；b64 旧项走 data URL 渐进兼容 */
const photoUrlCache = new Map<string, string>()
function photoSrcOf (it: PhotoItem): string {
  if (it.photoDbId) return photoUrlCache.get(it.photoDbId) || ''
  return it.b64 ? `data:image/${it.type};base64,${it.b64}` : ''
}
function photoSrcThumbOf (it: PhotoItem): string {
  if (it.thumbDbId) return photoUrlCache.get(it.thumbDbId) || photoSrcOf(it)
  return photoSrcOf(it)
}

export function PhotoNodeCard ({ nodeId, lod }: { nodeId: string; lod?: boolean }) {
  const tables = useBaseStore((s) => s.tables)
  /* ★ 统一数据锁（2026-10-26）：锁定时禁 导入/删除/排序（通用字段，OO 遗留兼容） */
  const lcfg = (tables[nodeId]?.config ?? {}) as { dataLocked?: boolean; ooDataLocked?: boolean }
  const dataLocked = !!(lcfg.dataLocked || lcfg.ooDataLocked)
  const cfg = (tables[nodeId]?.config ?? {}) as { photoItems?: PhotoItem[]; photoOrder?: string[] }
  const raw = cfg.photoItems ?? []
  /* ★ 2026-10-24 拖动 LOD：拖动中照片网格降级为占位色块（多图解码合成是拖动重开销） */
  if (lod) {
    const n = raw.length
    return (
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, minWidth: 200 }}>
        {Array.from({ length: Math.min(n, 6) }).map((_, i) => (
          <div key={i} style={{ width: 64, height: 64, borderRadius: 6, background: 'var(--c-bg3, #8884)', border: '1px dashed var(--c-border)' }} />
        ))}
        {n > 6 && <div style={{ alignSelf: 'center', fontSize: 12, color: 'var(--c-text3)' }}>+{n - 6}</div>}
      </div>
    )
  }
  const items = orderedPhotos(raw, cfg.photoOrder ?? null)
  const [msg, setMsg] = useState('')
  /* ★ 照片 URL 异步装载：photoDbId 项从 IndexedDB 取字节 → object URL 进缓存 → 触发重渲 */
  const [urlTick, setUrlTick] = useState(0)
  void urlTick
  useEffect(() => {
    let alive = true
    const load = async () => {
      let changed = false
      for (const it of raw) {
        const key = it.thumbDbId || it.photoDbId
        if (!key || photoUrlCache.has(key)) continue
        try {
          const bytes = await getBlob(key)
          if (!bytes || !alive) continue
          const blob = new Blob([bytes.slice().buffer as ArrayBuffer], { type: `image/${it.type}` })
          photoUrlCache.set(key, URL.createObjectURL(blob))
          changed = true
        } catch { /* 单张失败不阻塞 */ }
      }
      if (changed && alive) setUrlTick((t) => t + 1)
    }
    void load()
    return () => { alive = false }
  }, [raw])
  /* ★ 2026-10-23 用户指令：从资源管理器拖图片文件到节点即导入；
     depth 计数防子元素 enter/leave 抖动，拖入时高亮虚线框 */
  const dragDepth = useRef(0)
  const [dragOver, setDragOver] = useState(false)
  /* 拖拽排序状态（与 PDF 合并行交互同款） */
  const dragIdx = useRef<number | null>(null)
  const [dragging, setDragging] = useState<number | null>(null)
  const [overIdx, setOverIdx] = useState<number | null>(null)
  const [droppedIdx, setDroppedIdx] = useState<number | null>(null)
  /* ★ 拖拽态全局兜底清理（同 PDF 行——dragend 丢失防御） */
  useEffect(() => {
    if (dragging == null) return
    const clear = () => { setDragging(null); setOverIdx(null) }
    window.addEventListener('pointerup', clear)
    window.addEventListener('mouseup', clear)
    window.addEventListener('blur', clear)
    return () => {
      window.removeEventListener('pointerup', clear)
      window.removeEventListener('mouseup', clear)
      window.removeEventListener('blur', clear)
    }
  }, [dragging])

  const stAll = () => useBaseStore.getState()
  const setCfg = (patch: Record<string, unknown>) => stAll().setNodeConfig(nodeId, patch)
  const flash = (m: string) => { setMsg(m); setTimeout(() => setMsg(''), 4000) }
  /* ★ 2026-10-23 用户指令「点击照片节点的每一行，在右边有预览」：
     点击行 → 画布右缘固定预览浮层（大图 + 文件名 + × 关闭；再点行切换） */
  const [preview, setPreview] = useState<PhotoItem | null>(null)
  /* ★ 2026-10-23 用户指令「双击图片打开 类似ooexcel大小」：居中大窗预览（模态，点遮罩/×关）；单击仍是节点右侧小预览 */
  const [previewBig, setPreviewBig] = useState<PhotoItem | null>(null)
  /* ★ 2026-10-23 用户指令「预览不要在上方，在节点右侧」：点击时抓节点 DOM rect，
     预览卡 fixed 定位到节点右缘（rect.right + 12, 与节点顶对齐）。 */
  const [previewPos, setPreviewPos] = useState<{ left: number; top: number }>({ left: 0, top: 0 })
  /* ★ 2026-10-23 用户指令「移动画布关闭预览窗口」：监听 RF 视口（平移/缩放）——
     小预览是 fixed 屏幕坐标，画布一动就与节点错位 → 视口一变即收起 */
  const viewport = useViewport()
  useEffect(() => {
    setPreview(null)
  }, [viewport.x, viewport.y, viewport.zoom])
  /* ★ 预览打开时点卡片外任意处即关（不挡画布操作）：预览是纯浮层无遮罩，
     用户点空白/拖节点/平移画布的第一下 mousedown 就把预览收掉——「不影响下一步操作」。 */
  useEffect(() => {
    if (!preview) return
    const onDoc = (ev: MouseEvent) => {
      const t = ev.target as HTMLElement
      if (!t.closest('.pn-preview') && !t.closest('.pn-row')) setPreview(null)
    }
    document.addEventListener('mousedown', onDoc)
    return () => document.removeEventListener('mousedown', onDoc)
  }, [preview])

  /* 薄封装：共享导入逻辑（2026-10-23 抽取——卡片拖放与锁定组行 PhotoImportButton 同源） */
  const importFiles = async (files: FileList | null) => {
    const r = await importPhotoFiles(nodeId, files)
    if (r.skipped) flash(`跳过 ${r.skipped} 个（仅支持 jpg/png，单张 ≤8MB，最多 ${PHOTO_MAX} 张）`)
  }

  const removeAt = (i: number) => {
    const victim = items[i]
    if (!victim) return
    if (dataLocked) { flash('数据已锁定：先解锁再删除'); return } /* ★ 统一数据锁 */
    stAll().pushHistory()
    stAll().setNodeConfig(nodeId, {
      photoItems: raw.filter((it) => it.id !== victim.id),
      photoOrder: (cfg.photoOrder ?? []).filter((id) => id !== victim.id),
    })
    pushLog(`照片删除：${victim.name}（剩 ${items.length - 1} 张）`)
  }

  /** 拖动中悬停换位：id 序写 config.photoOrder（持久，删除/重排皆稳定） */
  const swapTo = (from: number, to: number) => {
    if (dataLocked) return /* ★ 统一数据锁：锁定禁排序 */
    if (from === to || from < 0 || to < 0 || from >= items.length || to >= items.length) return
    const ids = items.map((it) => it.id)
    const [moved] = ids.splice(from, 1)
    ids.splice(to, 0, moved)
    setCfg({ photoOrder: ids })
    dragIdx.current = to
  }

  return (
    <div
      className={'pn-root nodrag nopan' + (dragOver ? ' pn-root--over' : '')}
      onDragEnter={(e) => { e.preventDefault(); dragDepth.current++; setDragOver(true) }}
      onDragOver={(e) => { e.preventDefault(); try { e.dataTransfer.dropEffect = 'copy' } catch { /* ignore */ } }}
      onDragLeave={() => { dragDepth.current--; if (dragDepth.current <= 0) { dragDepth.current = 0; setDragOver(false) } }}
      onDrop={async (e) => {
        e.preventDefault()
        e.stopPropagation()
        dragDepth.current = 0
        setDragOver(false)
        if (e.dataTransfer.files?.length) {
          if (dataLocked) { flash('数据已锁定：先解锁再导入'); return } /* ★ 统一数据锁 */
          await importFiles(e.dataTransfer.files)
        }
      }}
    >
      {/* ① 张数徽章已删（2026-10-23 用户指令）：数量在标题栏右侧幽灵数字（同节点组语言） */}
      {/* ② 照片行列表：每行一张，拖拽排序（悬停换位跟随让位），hover 出删除 */}
      {items.length === 0 && (
        <div className="pn-empty">导入后可连入 PDF 合并节点（每图一页）</div>
      )}
      {items.map((it, i) => (
        <div
          key={it.id}
          className={
            'pn-row nodrag nopan'
            + (dragging === i ? ' pn-row--drag' : '')
            + (overIdx === i && dragging !== i ? ' pn-row--over' : '')
            + (droppedIdx === i ? ' pn-row--dropped' : '')
          }
          draggable
          onDragStart={(e) => {
            dragIdx.current = i
            setDragging(i)
            try {
              const c = document.createElement('canvas')
              c.width = 1; c.height = 1
              e.dataTransfer.setDragImage(c, 0, 0)
              e.dataTransfer.effectAllowed = 'move'
            } catch { /* 旧浏览器忽略 */ }
          }}
          onDragEnter={() => { setOverIdx(i); if (dragIdx.current != null) swapTo(dragIdx.current, i) }}
          onDragOver={(e) => { e.preventDefault(); try { e.dataTransfer.dropEffect = 'move' } catch { /* ignore */ } }}
          onDrop={(e) => e.preventDefault()}
          onDragEnd={() => {
            const landed = dragIdx.current
            dragIdx.current = null
            setDragging(null)
            setOverIdx(null)
            if (landed != null) {
              setDroppedIdx(landed)
              setTimeout(() => setDroppedIdx((c) => (c === landed ? null : c)), 500)
            }
          }}
          /* 点击（非拖动结束时）→ 右侧大图预览：mousedown/up 位移 <4px 判定为点击
             （HTML5 dnd 拖起后 click 不触发 ✓，但保险再判位移） */
          onClick={(e) => {
            if (dragIdx.current != null) return
            const nodeEl = (e.currentTarget as HTMLElement).closest('.react-flow__node') as HTMLElement | null
            const r = nodeEl?.getBoundingClientRect()
            if (r) {
              const W = 340
              const left = r.right + 12 + W > window.innerWidth ? r.left - W - 12 : r.right + 12
              setPreviewPos({ left: Math.max(8, left), top: Math.max(8, r.top) })
            }
            setPreview((p) => (p?.id === it.id ? null : it))
          }}
          onDoubleClick={(e) => {
            /* ★ 2026-10-23 用户指令「双击打开 类似ooexcel大小」：关小预览、开居中大窗 */
            e.preventDefault()
            setPreview(null)
            setPreviewBig(it)
          }}
          title={`${it.name}（单击右侧预览 · 双击大窗 · 拖动调整顺序 = 合并 PDF 的页序）`}
        >
          <span className="pn-idx">{i + 1}</span>
          <img className="pn-thumb" src={photoSrcThumbOf(it) || undefined} alt={it.name} draggable={false} />
          <span className="pn-nm">{it.name}</span>
          <button
            className="pn-del"
            title="删除这张"
            onClick={(e) => { e.stopPropagation(); removeAt(i) }}
          >×</button>
          <span className="pn-grip">⠿</span>
        </div>
      ))}
      {msg && <div className="pn-msg">{msg}</div>}
      {/* ★ 2026-10-23 用户指令「预览在节点右侧」：fixed 定位到节点右缘（非全屏遮罩）。
          点卡片外空白关闭（document click 一次性监听）。 */}
      {preview && createPortal(
        <div className="pn-preview nodrag nopan" style={{ left: previewPos.left, top: previewPos.top }} onClick={(e) => e.stopPropagation()}>
          <div className="pn-preview-card">
            <div className="pn-preview-bar">
              <span className="pn-preview-nm">{preview.name}</span>
              <button className="pn-preview-close" title="关闭预览" onClick={() => setPreview(null)}>×</button>
            </div>
            <img className="pn-preview-img" src={photoSrcOf(preview) || undefined} alt={preview.name} draggable={false} />
          </div>
        </div>,
        document.body,
      )}
      {/* ★ 2026-10-23 双击大窗预览（用户指令「双击打开 类似ooexcel大小」）：
          全屏暗遮罩 + 居中 min(1180px,92vw)×82vh 大图（同 node-editor/pm-pdfpreview 语言） */}
      {previewBig && createPortal(
        <div
          className="pm-pdfpreview nodrag nopan"
          onClick={() => setPreviewBig(null)}
        >
          <div className="pn-preview-card pm-pdfpreview-card" onClick={(e) => e.stopPropagation()}>
            <div className="pn-preview-bar">
              <span className="pn-preview-nm">🖼 {previewBig.name}</span>
              <button className="pn-preview-close" title="关闭预览" onClick={() => setPreviewBig(null)}>×</button>
            </div>
            <img
              className="pm-preview-img-big"
              src={photoSrcOf(previewBig) || undefined}
              alt={previewBig.name}
              draggable={false}
            />
          </div>
        </div>,
        document.body,
      )}
    </div>
  )
}
