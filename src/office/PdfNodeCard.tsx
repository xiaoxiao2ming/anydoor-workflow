/**
 * PDF 文件节点 · 行式列表（2026-10-23 用户指令「增加一个承载PDF的节点」）
 *
 * 结构与照片节点同语言：
 *   行 = 序号 + 📄 图标 + 文件名 + hover 删除 + 拖把
 *   ——拖拽悬停换位（其它行跟随让位）、落位闪绿，排序持久 config.pdfFileOrder（id 序）
 *   拖系统文件到节点即导入（同照片节点）；base64 存 config.pdfItems 随画布走。
 * 下游连 PDF 合并节点：节点内文件按行序直接拼接（pdf-lib 直转，零转换）。
 */
import { useEffect, useRef, useState, useSyncExternalStore } from 'react'
import { createPortal } from 'react-dom'
import { useBaseStore } from '../core/store'
import { pushLog } from '../ui/logBus'
import { IcMaximize, IcMinimize, IcX, IcDoc, IcHistory, IcDownload, IcRefresh } from '../ui/icons'
import { PdfViewer, type PdfMark } from './PdfViewer'
import { getBlob, putBlob } from '../core/pdfBlobStore'
import { PdfMupdf } from './PdfMupdf'
import { convertToPdf } from './PdfMergeButtons'
import { getTemplateStore } from '../docx/templateStore'
import { setNodeRunStatus, setNodeRunStatusAuto } from '../core/nodeRunStatus'

export interface PdfItem { id: string; name: string; b64?: string; inDb?: boolean; marks?: { id: string; page: number; text: string; rects: { l: number; t: number; w: number; h: number }[] }[] }
/** 「从上游生成」历史条目（2026-10-29）：与 pdfmerge pdfHistory 同口径——字节入库 pdfgenhist-<id>，config 只留索引 */
export interface PdfGenHistEntry { id: string; name: string; at: number; inDb?: boolean }

/* ★ 2026-10-25 用户指令「限制大小」放宽：字节走 IDB 引用（inDb）不进 config——
   b64 内嵌时代的 20MB/24 份保护已过时，放宽到 200MB/100 份（IDB 配额内） */
const PDF_MAX = 100
const PDF_MAX_BYTES = 200 * 1024 * 1024

/* ★ b64Of 已废（2026-10-25）：组件内导入与 Canvas 拖入统一 IDB 模式（putBlob 'pdfitem-'+id），
   b64 字符串拼接 + btoa 大分配是「拖大文件闪退」根因——勿再复活 */

/** pdfItems 按 pdfFileOrder（id 序）排列——拖拽排序与删除的单一真相源 */
export function orderedPdfs (items: PdfItem[], order: string[] | undefined | null): PdfItem[] {
  if (!order || !order.length) return items
  const byId = new Map(items.map((it) => [it.id, it]))
  const head = order.map((id) => byId.get(id)).filter((it): it is PdfItem => !!it)
  const tail = items.filter((it) => !order.includes(it.id))
  return [...head, ...tail]
}

/** PDF 导入共享逻辑（节点拖放 / 首行按钮两处共用） */
export async function importPdfFiles (
  nodeId: string,
  files: FileList | null,
): Promise<{ added: number; skipped: number; total: number }> {
  const st = useBaseStore.getState()
  const cfg = (st.tables[nodeId]?.config ?? {}) as { pdfItems?: PdfItem[]; pdfFileOrder?: string[] }
  const raw = cfg.pdfItems ?? []
  const next: PdfItem[] = [...raw]
  let skipped = 0
  for (const f of Array.from(files ?? [])) {
    const ext = f.name.toLowerCase().split('.').pop() ?? ''
    if (ext !== 'pdf' || f.size > PDF_MAX_BYTES) { skipped++; continue }
    if (next.length >= PDF_MAX) { skipped++; continue }
    try {
      const itId = 'pd-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
      /* ★ 2026-10-25 修「拖大文件闪退」：组件内导入与 Canvas 拖入同款 IDB 模式——
         b64Of 的 O(n²) 字符串拼接 + btoa 大分配把渲染层打崩（render-process-gone 实锤）。
         消费端（预览/合并/OO）早已 getBlob('pdfitem-'+id) + b64 兜底，此处对齐即可 */
      await putBlob('pdfitem-' + itId, new Uint8Array(await f.arrayBuffer()))
      next.push({ id: itId, name: f.name, inDb: true })
    } catch { skipped++ }
  }
  st.pushHistory()
  st.setNodeConfig(nodeId, { pdfItems: next, pdfFileOrder: (cfg.pdfFileOrder ?? []).concat(next.slice(raw.length).map((it) => it.id)) })
  pushLog(`PDF 导入：+${next.length - raw.length} 份（共 ${next.length}）${skipped ? `，跳过 ${skipped} 个非 pdf 或超 200MB` : ''}`, next.length > raw.length ? 'ok' : 'warn')
  return { added: next.length - raw.length, skipped, total: next.length }
}

/** 导入按钮（与 GrpLockButton/PhotoImportButton 同语言：31px 高圆角长方形） */
export function PdfImportButton ({ nodeId }: { nodeId: string }) {
  const fileRef = useRef<HTMLInputElement>(null)
  const [busy, setBusy] = useState(false)
  /* ★ 统一数据锁（2026-10-26）：锁定时禁导入（通用 dataLocked || OO 遗留 ooDataLocked） */
  const lcfg = useBaseStore((s) => s.tables[nodeId]?.config) as { dataLocked?: boolean; ooDataLocked?: boolean } | undefined
  const locked = !!(lcfg?.dataLocked || lcfg?.ooDataLocked)
  return (
    <button
      className="pn-import-btn nodrag nopan"
      disabled={busy || locked}
      title={locked ? '数据已锁定：先解锁再导入' : '导入 PDF（可多选，≤200MB/份，字节存本机库随画布引用）'}
      onClick={() => fileRef.current?.click()}
    >
      📄
      <input
        ref={fileRef}
        type="file"
        accept="application/pdf,.pdf"
        multiple
        style={{ display: 'none' }}
        onChange={async (e) => {
          setBusy(true)
          try { await importPdfFiles(nodeId, e.target.files) } finally { setBusy(false); e.target.value = '' }
        }}
      />
    </button>
  )
}

/* ═══════════ 从上游生成（2026-10-29 用户指令）═══════════
   有上游连线（linkedInputs/inputs，含「转为PDF节点」自动连线）→ 方形⟳钮（共享按钮行）
   点击：逐个上游 OO 文档取最新字节 → /pdfconv x2t 转 PDF 承接本节点；
   旧 pdfItems 逐份归档进 pdfGenHistory——与 PDF 合并节点 pdfHistory 同口径：
   字节入库 pdfgenhist-<id> + config 只留索引（cap 10），🕘钮展开列表（双击外部打开/⤓另存/×删记录） */

let genHistOpen = false
const genHistSubs = new Set<() => void>()
const genHistUI = {
  get: () => genHistOpen,
  toggle: () => { genHistOpen = !genHistOpen; genHistSubs.forEach((f) => f()) },
  sub: (f: () => void) => { genHistSubs.add(f); return () => { genHistSubs.delete(f) } },
}
const apiG = (globalThis as { api?: { openExternalFile?: (name: string, data: Uint8Array) => Promise<string> } }).api
const fmtTimeG = (at: number) => {
  const d = new Date(at)
  const p = (n: number) => String(n).padStart(2, '0')
  return `${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`
}

/** 上游 OO 节点最新字节（ooBytes64 → ooBytesDbId → 模板库，与打开/预加载同优先级） */
async function readUpstreamBytes (uid: string): Promise<{ u8: Uint8Array; name: string } | null> {
  const t = useBaseStore.getState().tables[uid]
  if (!t || t.deleted || (t.kind !== 'ooword' && t.kind !== 'oosheet')) return null
  const c = (t.config ?? {}) as { ooBytes64?: string; ooBytesDbId?: string; ooTemplateId?: string }
  let u8: Uint8Array | null = null
  if (c.ooBytes64) {
    const bin = atob(c.ooBytes64)
    u8 = new Uint8Array(bin.length)
    for (let i = 0; i < bin.length; i++) u8![i] = bin.charCodeAt(i)
  }
  if ((!u8 || !u8.length) && c.ooBytesDbId) { try { u8 = (await getBlob(c.ooBytesDbId)) || null } catch { u8 = null } }
  if ((!u8 || !u8.length) && c.ooTemplateId) { try { u8 = (await getTemplateStore().read(c.ooTemplateId)) || null } catch { u8 = null } }
  if (!u8 || !u8.length) return null
  const base = String(t.name ?? '文档').replace(/\.(docx|xlsx)$/i, '')
  return { u8, name: base + (t.kind === 'oosheet' ? '.xlsx' : '.docx') }
}

/** 生成主流程：旧内容归档 → 上游转 PDF 承接。返回汇总消息（按钮 toast/日志用） */
export async function runPdfGen (nodeId: string): Promise<string> {
  const st = useBaseStore.getState()
  const t = st.tables[nodeId]
  if (!t) return '节点不存在'
  const lcfg = (t.config ?? {}) as { dataLocked?: boolean; ooDataLocked?: boolean }
  if (lcfg.dataLocked || lcfg.ooDataLocked) return '数据已锁定：先解锁再生成'
  const cfg = (t.config ?? {}) as { pdfItems?: PdfItem[]; pdfFileOrder?: string[]; linkedInputs?: string[]; inputs?: string[]; pdfGenHistory?: PdfGenHistEntry[] }
  const upIds: string[] = []
  for (const s of [...(cfg.linkedInputs ?? []), ...(cfg.inputs ?? [])]) if (!upIds.includes(s)) upIds.push(s)
  if (!upIds.length) return '无上游连线——先把 OO 文档节点连过来'
  const raw = cfg.pdfItems ?? []
  const t0 = Date.now()
  /* ★ 2026-10-29 用户指令「包括转化的状态提示」：节点徽章实时报阶段/进度（pdf2word 同语言） */
  if (raw.length) setNodeRunStatus(nodeId, 'running', `归档旧内容 ${raw.length} 份…`)
  /* ① 旧内容逐份归档（字节入库 + 索引，单份失败不拦生成） */
  const newHist: PdfGenHistEntry[] = []
  for (const it of raw) {
    try {
      let bytes: Uint8Array | null = null
      if (it.b64) {
        const bin = atob(it.b64)
        bytes = new Uint8Array(bin.length)
        for (let i = 0; i < bin.length; i++) bytes![i] = bin.charCodeAt(i)
      }
      if (!bytes) bytes = (await getBlob('pdfitem-' + it.id)) || null
      if (!bytes || !bytes.length) continue
      const hid = 'pg' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6)
      await putBlob('pdfgenhist-' + hid, bytes)
      newHist.push({ id: hid, name: it.name, at: Date.now(), inDb: true })
    } catch { /* 归档失败不拦生成 */ }
  }
  /* ② 上游转 PDF 承接为本节点当前内容 */
  const newItems: PdfItem[] = []
  const newOrder: string[] = []
  for (let ui = 0; ui < upIds.length; ui++) {
    const uid = upIds[ui]
    const nm = st.tables[uid]?.name ?? uid
    setNodeRunStatus(nodeId, 'running', `转 PDF 中 ${ui + 1}/${upIds.length} · ${nm}`)
    const src = await readUpstreamBytes(uid)
    if (!src) { pushLog(`[PDF·生成] 跳过「${nm}」（非 OO 文档节点或无内容）`, 'warn'); continue }
    const pdf = await convertToPdf(src.u8, src.name)
    if (!pdf || !pdf.length) { pushLog(`[PDF·生成] 「${nm}」转换失败`, 'err'); continue }
    const pid = 'pd-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
    await putBlob('pdfitem-' + pid, pdf)
    newItems.push({ id: pid, name: src.name.replace(/\.(docx|xlsx)$/i, '') + '.pdf', inDb: true })
    newOrder.push(pid)
    pushLog(`[PDF·生成] ${nm} → ${Math.round(pdf.length / 1024)}KB`, 'ok')
  }
  if (!newItems.length) {
    setNodeRunStatus(nodeId, 'error', '上游无可生成内容')
    return '上游无可生成内容'
  }
  st.pushHistory()
  st.setNodeConfig(nodeId, { pdfItems: newItems, pdfFileOrder: newOrder, pdfGenHistory: [...newHist, ...(cfg.pdfGenHistory ?? [])].slice(0, 10) })
  const sec = ((Date.now() - t0) / 1000).toFixed(1)
  setNodeRunStatusAuto(nodeId, 'ok', `${newItems.length} 份 PDF · ${sec}s`)
  pushLog(`[PDF·生成] 已承接 ${newItems.length} 份最新 PDF${newHist.length ? `，旧 ${newHist.length} 份存入历史` : ''}（${sec}s）`, 'ok')
  return `已生成 ${newItems.length} 份（${sec}s）`
}

/** 方形「⟳ 从上游生成」按钮（Canvas 共享按钮行；有上游连线才显示，31×31 同语言） */
export function PdfGenButton ({ nodeId }: { nodeId: string }) {
  const [busy, setBusy] = useState(false)
  const cfg = useBaseStore((s) => s.tables[nodeId]?.config) as { linkedInputs?: string[]; inputs?: string[]; dataLocked?: boolean; ooDataLocked?: boolean } | undefined
  const upN = (cfg?.linkedInputs ?? []).length + (cfg?.inputs ?? []).length
  if (!upN) return null
  const locked = !!(cfg?.dataLocked || cfg?.ooDataLocked)
  return (
    <button
      className={'pm-sq-btn nodrag nopan' + (busy ? ' pm-sq-btn--busy' : '')}
      disabled={busy || locked}
      title={locked ? '数据已锁定：先解锁再生成' : busy ? '生成中…（进度见节点徽章与日志）' : `从上游生成（${upN} 个上游）：按上游文档最新内容重新转 PDF（旧内容存入历史）`}
      onClick={async (e) => {
        e.stopPropagation()
        if (busy) return
        setBusy(true)
        try {
          const msgR = await runPdfGen(nodeId)
          pushLog(`[PDF·生成] ${msgR}`, 'ok')
        } finally { setBusy(false) }
      }}
    >
      <IcRefresh size={19} />
    </button>
  )
}

/** 历史方形钮（🕘，与 PDF转Word 历史钮同语言；有历史才显示） */
export function PdfGenHistButton ({ nodeId }: { nodeId: string }) {
  const open = useSyncExternalStore(genHistUI.sub, genHistUI.get)
  const cfg = useBaseStore((s) => s.tables[nodeId]?.config) as { pdfGenHistory?: PdfGenHistEntry[] } | undefined
  if (!(cfg?.pdfGenHistory ?? []).length) return null
  return (
    <button
      className={'pm-sq-btn pm-sq-btn--ghost nodrag nopan' + (open ? ' pm-sq-btn--on' : '')}
      title="生成历史（旧内容 PDF 列表）"
      onClick={(e) => { e.stopPropagation(); genHistUI.toggle() }}
    >
      <IcHistory size={19} />
    </button>
  )
}

/** 历史列表（按钮行正下方；双击=默认软件打开，hover ⤓另存 / ×删记录——同 PDF转Word 历史语言） */
export function PdfGenHistoryList ({ nodeId }: { nodeId: string }) {
  const open = useSyncExternalStore(genHistUI.sub, genHistUI.get)
  const cfg = useBaseStore((s) => s.tables[nodeId]?.config) as { pdfGenHistory?: PdfGenHistEntry[] } | undefined
  const hist = cfg?.pdfGenHistory ?? []
  if (!open || !hist.length) return null
  return (
    <div className="pm-hist">
      {hist.map((r) => (
        <div
          key={r.id}
          className="pn-row pm-hist-row nodrag nopan"
          title={`${r.name}（${fmtTimeG(r.at)}）—— 双击用默认软件打开`}
          onDoubleClick={async () => {
            const bytes = await getBlob('pdfgenhist-' + r.id)
            if (!bytes || !apiG?.openExternalFile) { pushLog('[PDF·生成] 打开失败：字节缺失或桌面桥缺失', 'warn'); return }
            try { await apiG.openExternalFile(r.name, bytes) } catch (e) { pushLog(`[PDF·生成] 打开失败：${e instanceof Error ? e.message : String(e)}`, 'err') }
          }}
        >
          <span className="pm-fmt pm-fmt-pdf">P</span>
          <span className="pn-nm">{r.name}</span>
          <span className="fo-saved-meta">{fmtTimeG(r.at)}</span>
          <button
            className="pn-del"
            title="另存为…"
            onClick={async (e) => {
              e.stopPropagation()
              const bytes = await getBlob('pdfgenhist-' + r.id)
              if (!bytes) return
              /* 与 pdfmerge 历史条目同款桥（pdf 分支） */
              const apiS = (window as unknown as { api?: { docSaveAs?: (n: string, d: Uint8Array, k?: string) => Promise<string | null> } }).api
              if (!apiS?.docSaveAs) { pushLog('[PDF·生成] 另存失败：桌面桥缺失（重启应用）', 'warn'); return }
              try {
                const saved = await apiS.docSaveAs(r.name, bytes, 'pdf')
                if (saved) pushLog(`[PDF·生成] 已另存：${saved}`, 'ok')
              } catch (eS) { pushLog(`[PDF·生成] 另存失败：${eS instanceof Error ? eS.message : String(eS)}`, 'err') }
            }}
          >
            <IcDownload size={12} />
          </button>
          <button
            className="pn-del"
            title="删除记录（文件字节保留在本机库）"
            onClick={(e) => {
              e.stopPropagation()
              useBaseStore.getState().pushHistory()
              useBaseStore.getState().setNodeConfig(nodeId, { pdfGenHistory: hist.filter((x) => x.id !== r.id) })
            }}
          >
            <IcX size={12} />
          </button>
        </div>
      ))}
    </div>
  )
}

export function PdfNodeCard ({ nodeId, lod }: { nodeId: string; lod?: boolean }) {
  const tables = useBaseStore((s) => s.tables)
  /* ★ 统一数据锁（2026-10-26）：锁定时禁 导入/删除/排序（通用字段，OO 遗留兼容） */
  const lcfg = (tables[nodeId]?.config ?? {}) as { dataLocked?: boolean; ooDataLocked?: boolean }
  const dataLocked = !!(lcfg.dataLocked || lcfg.ooDataLocked)
  const cfg = (tables[nodeId]?.config ?? {}) as { pdfItems?: PdfItem[]; pdfFileOrder?: string[] }
  const raw = cfg.pdfItems ?? []
  /* ★ 拖动 LOD：拖动中渲染轻量骨架（一行文字），松手恢复完整卡片 */
  if (lod) return <div className="pn-root pn-lod" style={{ height: Math.max(120, raw.length * 34 + 70) }} />
  const items = orderedPdfs(raw, cfg.pdfFileOrder ?? null)
  const [msg, setMsg] = useState('')
  const flash = (m: string) => { setMsg(m); setTimeout(() => setMsg(''), 4000) }
  /* 拖入高亮（depth 计数防子元素 enter/leave 抖动） */
  const dragDepth = useRef(0)
  const [dragOver, setDragOver] = useState(false)
  /* ★ 2026-10-25 用户指令「空态做成导入文件的长条按钮」 */
  const emptyFileRef = useRef<HTMLInputElement>(null)
  /* ★ 2026-10-23 用户指令「PDF节点双击同样也能打开PDF」：双击节点/行 → 居中大窗
     iframe 内嵌查看行序对应份；多份时头部下拉切换 */
  const [viewIdx, setViewIdx] = useState<number | null>(null)
  const [viewMax, setViewMax] = useState(false)
  /* ★ 2026-10-23 修「最大化时内容先偏在一边再闪到中间」：PDF viewer resize 后
     需重排页面居中——重排过程可见产生闪跳。切换时 visibility 藏两帧
     （节点不卸载不重载），重排完成后显示——全程视觉正中。 */
  const [frameHidden, setFrameHidden] = useState(false)
  const toggleMax = () => {
    setViewMax((v) => !v)
    setFrameHidden(true)
    /* ★ 2026-10-23 用户指令「充分优化 最大化不够丝滑 还得 0.1 秒后才居中 偶尔卡卡」：
       PDFium 重排实际需 50~150ms（两帧不够盖住）——延到 ~180ms 隐藏窗口，
       重排全程不可见；显示带 0.15s 淡入渐显掩盖残余——视觉即点即正中。 */
    requestAnimationFrame(() => requestAnimationFrame(() => {
      setTimeout(() => setFrameHidden(false), 150)
    }))
  }
  /* ★ 交互语义：单击=选中/拖动，双击=打开。最大化=纯 CSS 类切换（key 重建 iframe），
     不做卸载重挂/尺寸过渡（PDFium 对此白屏；卸载重挂又致窗体消失——均实测排除）。 */
  /* ★ 2026-10-23 修「双击打不开」：blobUrl 原先在渲染体内 IIFE 创建——任何重渲染
     都换新 URL → iframe 反复重载等于打不开（历史大窗是 state 所以正常）。
     改为打开时创建一次存 state，关闭时撤销。 */
  const [viewBytes, setViewBytes] = useState<Uint8Array | null>(null)
  /* ★ 2026-10-24 PDF.js Viewer 薄壳模式（右键「用 PDF.js Viewer 打开」） */
  const [viewMode, setViewMode] = useState<'default' | 'html' | 'mupdf'>('default')
  const viewBlobUrl = useRef<string | null>(null)
  /* 事件桥：Canvas 右键「用 PDF.js Viewer 打开」→ 本卡片 html 模式预览 */
  useEffect(() => {
    const h = (e: Event) => {
      const d = (e as CustomEvent).detail as { nodeId: string }
      if (d?.nodeId === nodeId) openView(0, 'html')
    }
    window.addEventListener('open-pdf-pjsviewer', h)
    const hM = (e: Event) => {
      const d = (e as CustomEvent).detail as { nodeId: string }
      if (d?.nodeId === nodeId) openView(0, 'mupdf')
    }
    window.addEventListener('open-pdf-mupdf', hM)
    return () => {
      window.removeEventListener('open-pdf-pjsviewer', h)
      window.removeEventListener('open-pdf-mupdf', hM)
    }
  })
  /* ★ 交互语义（用户 18:54 指令）：单击=选中/拖动（无动作），双击=打开（只认 dblclick） */
  const openView = async (i: number, mode: 'default' | 'html' | 'mupdf' = 'default') => {
    const it = items[i]
    if (!it) return
    pushLog('[PDF] 双击打开：' + (it.name ?? '') + (mode !== 'default' ? '（' + (mode === 'html' ? 'PDF.js Viewer' : 'MuPDF') + '）' : ''))
    /* ★ 2026-10-24 b64 出 config：新数据从 IndexedDB 读，旧数据回退 config.b64 */
    let bytes: Uint8Array
    if (it.b64) {
      const bin = atob(it.b64)
      bytes = new Uint8Array(bin.length)
      for (let i2 = 0; i2 < bin.length; i2++) bytes[i2] = bin.charCodeAt(i2)
    } else {
      const fromDb = await getBlob('pdfitem-' + it.id)
      if (!fromDb) { pushLog('[PDF] 文件数据缺失: ' + (it.name ?? it.id), 'warn'); return }
      bytes = fromDb
    }
    /* ★ 2026-10-23 用户指令「最大化最大比例不超过 100% 但得居中」：URL 锁 #zoom=100——
       标准窗/最大化共用同一 URL（切换不重载），页面恒 100%、查看器自动居中 */
    /* ★ 2026-10-24 用户指令「github 最优套件右键打开」：html 模式 = pdf.js 官方查看器
       （/pdfjs-viewer/index.html 薄壳，Firefox 同款渲染管线），blob url 传 file */
    if (mode === 'html') {
      try { if (viewBlobUrl.current) URL.revokeObjectURL(viewBlobUrl.current) } catch { /* ignore */ }
      const bu = URL.createObjectURL(new Blob([bytes as BlobPart], { type: 'application/pdf' }))
      viewBlobUrl.current = bu
      setViewMode('html')
      setViewBytes(bytes)
      setViewIdx(i)
      return
    }
    if (mode === 'mupdf') {
      setViewMode('mupdf')
      setViewBytes(bytes)
      setViewIdx(i)
      return
    }
    setViewMode('default')
    setViewBytes(bytes)
    setViewIdx(i)
  }
  const closeView = () => {
    setViewBytes(null)
    setViewIdx(null)
    setViewMax(false)
  }
  /* 拖拽排序状态（与照片行交互同款） */
  const dragIdx = useRef<number | null>(null)
  const [dragging, setDragging] = useState<number | null>(null)
  const [overIdx, setOverIdx] = useState<number | null>(null)
  const [droppedIdx, setDroppedIdx] = useState<number | null>(null)
  /* ★ 2026-10-23 用户截图「拖拽释放后蓝框残留」防御：dragend 理论上总会触发，
     但个别场景（重渲染时序/失焦/取消）可能丢——拖拽中挂全局
     pointerup/mouseup 双保险，释放即清拖拽态 */
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

  const setCfg = (patch: Record<string, unknown>) => useBaseStore.getState().setNodeConfig(nodeId, patch)

  const importFiles = async (files: FileList | null) => {
    const r = await importPdfFiles(nodeId, files)
    if (r.skipped) flash(`跳过 ${r.skipped} 个（仅支持 pdf，单份 ≤200MB，最多 ${PDF_MAX} 份）`)
  }

  /** 拖动中悬停换位：id 序写 config.pdfFileOrder（持久，删除/重排皆稳定） */
  const swapTo = (from: number, to: number) => {
    if (dataLocked) return /* ★ 统一数据锁：锁定禁排序 */
    if (from === to || from < 0 || to < 0 || from >= items.length || to >= items.length) return
    const ids = items.map((it) => it.id)
    const [moved] = ids.splice(from, 1)
    ids.splice(to, 0, moved)
    setCfg({ pdfFileOrder: ids })
    dragIdx.current = to
  }

  const removeAt = (i: number) => {
    const victim = items[i]
    if (!victim) return
    if (dataLocked) { flash('数据已锁定：先解锁再删除'); return } /* ★ 统一数据锁 */
    useBaseStore.getState().pushHistory()
    setCfg({
      pdfItems: raw.filter((it) => it.id !== victim.id),
      pdfFileOrder: (cfg.pdfFileOrder ?? []).filter((id) => id !== victim.id),
    })
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
      onDoubleClick={() => openView(0)} /* ★ 语义（用户 18:54 指令）：单击=选中/拖动，双击=打开 */
    >
      {/* ① 张数徽章已删（2026-10-23 用户指令）：数量在标题栏右侧幽灵数字（同节点组语言） */}
      {/* ② PDF 行列表：每行一份，拖拽排序（悬停换位跟随让位），hover 出删除 */}
      <input
        ref={emptyFileRef}
        type="file"
        accept="application/pdf,.pdf"
        multiple
        style={{ display: 'none' }}
        onChange={async (e) => { await importFiles(e.target.files); e.target.value = '' }}
      />
      {items.length === 0 && (
        <button className="pn-empty-btn nodrag nopan" title={dataLocked ? '数据已锁定：先解锁再导入' : '点击导入 PDF（可多选，≤200MB/份）'} disabled={dataLocked} onClick={() => emptyFileRef.current?.click()}>
          ＋ 点击导入 PDF（也可拖入）
        </button>
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
          onDoubleClick={(e) => { e.stopPropagation(); openView(i) }}
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
          title={`${it.name}（拖动调整顺序 = 合并 PDF 的拼接顺序）`}
        >
          <span className="pn-idx">{i + 1}</span>
          <span className="pm-fmt pm-fmt-pdf">P</span>
          {/* ★ 2026-10-25 用户指令「增加文件属于外部引用还是软件内部标记」：
              内=字节在软件库（随工程文件走，不易丢）；外=外部引用（可能丢） */}
          <span
            className={'pn-loc ' + ((it.inDb || it.b64) ? 'pn-loc--in' : 'pn-loc--ext')}
            title={(it.inDb || it.b64) ? '软件内部：字节已入库，随工程文件存储' : '外部引用：字节不在软件库，可能丢失'}
          >{(it.inDb || it.b64) ? '内' : '外'}</span>
          <span className="pn-nm">{it.name}</span>
          <button className="pn-del" title="移除这份" onClick={() => removeAt(i)}>×</button>
          <span className="pn-grip">⠿</span>
        </div>
      ))}
      {msg && <div className="pn-msg">{msg}</div>}
      {/* ★ 2026-10-23 双击大窗：iframe 内嵌查看（多份时头部下拉切换；点遮罩/× 关）。
          blobUrl 存 state（打开时创建一次，关闭撤销）——渲染体内创建会随重渲染
          换 URL → iframe 反复重载等于打不开（用户实测根因） */}
      {/* ★ 最大化 = 同容器切 pm-card-max 类（iframe 同节点不重建不重载、
          页码保留；无动画——PDFium 对逐帧 resize 白屏已实测排除）。
          ⤢ 进最大化 / ❐ 还原；× 关闭 */}
      {viewBytes && viewIdx != null && items[viewIdx] && createPortal(
        <div className={'pm-pdfpreview' + (viewMax ? ' pm-preview-full' : '')} onClick={closeView}>
          <div className={'pn-preview-card pm-pdfpreview-card' + (viewMax ? ' pm-card-max' : '')} onClick={(e) => e.stopPropagation()}>
            <div className="pn-preview-bar">
              <span className="pn-preview-nm">📄 {items[viewIdx].name}</span>
              {items.length > 1 && (
                <select
                  className="libpick-select"
                  value={viewIdx}
                  title="切换节点内其它 PDF（行序）"
                  onChange={(e) => openView(Number(e.target.value), viewMode)}
                >
                  {items.map((it, i2) => <option key={it.id} value={i2}>{i2 + 1}. {it.name}</option>)}
                </select>
              )}
              <button className="pn-preview-close" title={viewMax ? '还原' : '最大化'} onClick={toggleMax}>
                {viewMax ? <IcMinimize size={19} /> : <IcMaximize size={19} />}
              </button>
              <button className="pn-preview-close" title="关闭" onClick={closeView}><IcX size={19} /></button>
            </div>
            <div className="pm-preview-body">
              {viewMode === 'html'
                ? <iframe className="pm-preview-frame" src={'/pdfjs-viewer/index.html?file=' + encodeURIComponent(viewBlobUrl.current || '')} title={items[viewIdx!].name} />
                : viewMode === 'mupdf'
                  ? <PdfMupdf data={viewBytes!} dark={document.documentElement.dataset.theme !== 'light'} />
                  : <PdfViewer
                      data={viewBytes!}
                      dark={document.documentElement.dataset.theme !== 'light'}
                      marks={(cfg.pdfItems || [])[viewIdx!]?.marks || []}
                      onSaveMarks={(ms) => {
                        const next = (cfg.pdfItems || []).map((it, ix) => (ix === viewIdx ? { ...it, marks: ms } : it))
                        useBaseStore.getState().pushHistory()
                        useBaseStore.getState().setNodeConfig(nodeId!, { pdfItems: next })
                      }}
                    />}
            </div>
          </div>
        </div>,
        document.body,
      )}
    </div>
  )
}
