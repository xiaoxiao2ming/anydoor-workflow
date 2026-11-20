/* ─────────────────────────────────────────────────────────────
 * PDF转Word 节点卡片（2026-10-26 用户指令「利用百度的开源 增加个节点 pdf2word」；
 * 18:57 用户指令「节点样式仿照PDF替换页节点 节点名称就叫PDF转word」——
 * 卡片结构整体对齐 PdfReplaceNodeCard：分区头锚点 + 词条行 + 行首锚点覆盖）
 *
 * 结构：
 *   分区头「待转换 PDF」（行首锚点 p2w-add = 上游连线追加到组）
 *   词条行 = 行首锚点（p2w-<id> = 上游覆盖本条）+ 序号 + P 方块 + 内外标 + 名称 + 删除
 *   拖拽排序（悬停换位）+ 文件拖入组件内自处理（零弹窗）
 *   缺套件引导行（节点接入规范——能力全缺时显示，点击直达套件中心）
 *   按钮行在 Canvas 共享行（锁组 + 导入 + W执行 + 🕘历史），历史列表在按钮行下
 *
 * 引擎：百度 PaddleOCR（本机外挂）+ x2t 文字版直转——链路见 pdf2wordEngine.ts
 * 存储：词条 p2witem-{id} / 产物 p2whist-{id}，config 只留索引（字节不进 config）
 * ───────────────────────────────────────────────────────────── */
import { useEffect, useRef, useState, useSyncExternalStore } from 'react'
import { Handle, Position, useUpdateNodeInternals } from '@xyflow/react'
import { useBaseStore } from '../core/store'
import { getBlob } from '../core/pdfBlobStore'
import { pushLog } from '../ui/logBus'
import { IcDownload, IcDoc, IcHistory, IcX } from '../ui/icons'
import { NodeSqBtn } from '../ui/NodeBtnBar'
import { hasCapability, suiteCenterUI, useSuiteState } from '../core/suites'
import { orderedP2w, importP2wFiles, runPdf2Word, removeP2wHist, type P2wItem } from './pdf2wordEngine'

/* 历史面板展开态（模块级轻量 store——按钮在 Canvas 共享行、列表在卡片内，跨组件共享） */
let histOpen = false
const histSubs = new Set<() => void>()
export const p2wHistUI = {
  get: () => histOpen,
  toggle: () => { histOpen = !histOpen; histSubs.forEach((f) => f()) },
  sub: (f: () => void) => { histSubs.add(f); return () => { histSubs.delete(f) } },
}

const api = (globalThis as {
  api?: {
    openExternalFile?: (name: string, data: Uint8Array) => Promise<string>
    docSaveAs?: (defaultName: string, data: Uint8Array, kind?: 'docx' | 'xlsx') => Promise<string | null>
  }
}).api

const fmtTime = (at: number) => {
  const d = new Date(at)
  const p = (n: number) => String(n).padStart(2, '0')
  return `${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`
}

/** 导入按钮（Canvas 共享按钮行，与 PdfImportButton 同语言） */
export function Pdf2WordImportButton ({ nodeId }: { nodeId: string }) {
  const fileRef = useRef<HTMLInputElement>(null)
  const [busy, setBusy] = useState(false)
  /* ★ 统一数据锁（2026-10-26）：锁定时禁导入（通用 dataLocked || OO 遗留 ooDataLocked） */
  const lcfg = useBaseStore((s) => s.tables[nodeId]?.config) as { dataLocked?: boolean; ooDataLocked?: boolean } | undefined
  const locked = !!(lcfg?.dataLocked || lcfg?.ooDataLocked)
  return (
    <button
      className="pn-import-btn nodrag nopan"
      disabled={busy || locked}
      title={locked ? '数据已锁定：先解锁再导入' : '导入 PDF（可多选，≤200MB/份）'}
      onClick={() => fileRef.current?.click()}
    >
      <IcDoc size={19} />
      <input
        ref={fileRef}
        type="file"
        accept="application/pdf,.pdf"
        multiple
        style={{ display: 'none' }}
        onChange={async (e) => {
          setBusy(true)
          try { await importP2wFiles(nodeId, e.target.files) } finally { setBusy(false); e.target.value = '' }
        }}
      />
    </button>
  )
}

/** 引擎切换钮（2026-10-26 用户指令「节点上的方形按钮 点击切换引擎」；
 *  19:05 用户指令「切换引擎 显示当前引擎的logo」）：
 *  「W」蓝方块=文字版直转（x2t 内置套件，秒级）/「百」百度品牌蓝方块=百度 OCR。
 *  点击即切换 config.p2wEngine（执行严格按选定引擎）。 */
export function Pdf2WordEngineButton ({ nodeId }: { nodeId: string }) {
  const cfg = useBaseStore((s) => s.tables[nodeId]?.config) as { p2wEngine?: 'x2t' | 'ocr' } | undefined
  const ocr = cfg?.p2wEngine === 'ocr'
  return (
    <NodeSqBtn
      icon={ocr
        ? <span className="pm-fmt pm-fmt-baidu">百</span>
        : <span className="pm-fmt pm-fmt-word">W</span>}
      active={ocr}
      title={ocr
        ? '当前引擎：百度 OCR（扫描件/图片型 PDF，首次转换需下载模型）——点击切换为文字版直转'
        : '当前引擎：文字版直转（x2t，秒级；扫描件转不出文字）——点击切换为百度 OCR'}
      onClick={(e) => {
        e.stopPropagation()
        useBaseStore.getState().pushHistory()
        useBaseStore.getState().setNodeConfig(nodeId, { p2wEngine: ocr ? 'x2t' : 'ocr' })
        pushLog(`[PDF转Word] 引擎已切换：${ocr ? '文字版直转（x2t）' : '百度 OCR'}`, 'ok')
      }}
    />
  )
}

/** 执行转换钮（「W」31×31 同 PdfMergeRunButton 语言；busy 转圈禁点）
 *  ★ 节点接入规范·缺套件态：所需能力（pdf2docx 任意引擎）全缺 → 置灰 + title 引导 */
export function Pdf2WordRunButton ({ nodeId }: { nodeId: string }) {
  const [busy, setBusy] = useState(false)
  useSuiteState() /* 套件停用/启用 → 本钮即时置灰/恢复 */
  const lcfg = useBaseStore((s) => s.tables[nodeId]?.config) as { dataLocked?: boolean; ooDataLocked?: boolean } | undefined
  const locked = !!(lcfg?.dataLocked || lcfg?.ooDataLocked)
  const noSuite = !hasCapability('convert:pdf2docx') && !hasCapability('convert:pdf2docx-ocr')
  const onClick = async () => {
    if (busy || locked || noSuite) return
    setBusy(true)
    try {
      await runPdf2Word(nodeId)
    } finally {
      setBusy(false)
    }
  }
  return (
    <button
      className="pm-sq-btn nodrag nopan"
      title={noSuite
        ? '缺转换套件：文字版直转已停用且 OCR 未安装——打开功能区右侧拼图按钮（套件中心）'
        : locked ? '数据已锁定：先解锁再执行转换' : '转换为 Word（文字版秒转；扫描件自动走 OCR）'}
      disabled={busy || locked || noSuite}
      onClick={onClick}
    >
      {busy ? <span className="rn-spin pm-spin-dark" /> : <span className="pm-sq-txt">W</span>}
    </button>
  )
}

/** 历史钮（🕘 同 PdfReplaceHistButton 语言；有历史才显示） */
export function Pdf2WordHistButton ({ nodeId }: { nodeId: string }) {
  const open = useSyncExternalStore(p2wHistUI.sub, p2wHistUI.get)
  const cfg = useBaseStore((s) => s.tables[nodeId]?.config) as { p2wHistory?: unknown[] } | undefined
  const has = (cfg?.p2wHistory ?? []).length > 0
  if (!has) return null
  return (
    <button
      className={'pm-sq-btn pm-sq-btn--ghost nodrag nopan' + (open ? ' pm-sq-btn--on' : '')}
      title="转换历史（产物 Word 列表）"
      onClick={(e) => { e.stopPropagation(); p2wHistUI.toggle() }}
    >
      <IcHistory size={19} />
    </button>
  )
}

/** 历史列表（按钮行正下方；双击=默认软件打开，hover ⤓另存 / ×删记录） */
export function Pdf2WordHistoryList ({ nodeId }: { nodeId: string }) {
  const open = useSyncExternalStore(p2wHistUI.sub, p2wHistUI.get)
  const cfg = useBaseStore((s) => s.tables[nodeId]?.config) as { p2wHistory?: { id: string; name: string; at: number; inDb?: boolean; b64?: string; engine?: 'x2t' | 'ocr' }[] } | undefined
  const hist = cfg?.p2wHistory ?? []
  const [flash, setFlash] = useState('')
  const flashMsg = (m: string) => { setFlash(m); setTimeout(() => setFlash(''), 3000) }
  if (!open || !hist.length) return null
  const readBytes = async (r: { id: string; name: string; b64?: string }) => {
    if (r.b64) {
      const bin = atob(r.b64)
      const u8 = new Uint8Array(bin.length)
      for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
      return u8
    }
    return getBlob('p2whist-' + r.id)
  }
  return (
    <div className="pm-hist">
      {hist.map((r) => (
        <div
          key={r.id}
          className="pn-row pm-hist-row nodrag nopan"
          title={`${r.name}（${fmtTime(r.at)}${r.engine ? ' · ' + (r.engine === 'x2t' ? '文字版直转' : 'OCR') : ''}）—— 双击用默认软件打开`}
          onDoubleClick={async () => {
            const bytes = await readBytes(r)
            if (!bytes || !api?.openExternalFile) { flashMsg('桌面桥缺失（重启应用）'); return }
            try { await api.openExternalFile(r.name, bytes) } catch (e) { pushLog(`[PDF转Word] 打开失败：${e instanceof Error ? e.message : String(e)}`, 'err') }
          }}
        >
          {/* ★ 2026-10-26 用户指令「转换的历史记录要记录是什么引擎转出来的（带个logo）」：
              引擎 logo 方块（与节点切换钮同语言）——W 蓝=文字版直转 / 百 品牌蓝=百度 OCR */}
          <span
            className={'pm-fmt ' + (r.engine === 'ocr' ? 'pm-fmt-baidu' : 'pm-fmt-word')}
            title={r.engine === 'ocr' ? '百度 OCR 引擎转换' : '文字版直转（x2t）'}
          >{r.engine === 'ocr' ? '百' : 'W'}</span>
          <span className="pn-nm">{r.name}</span>
          <span className="fo-saved-meta">{fmtTime(r.at)}</span>
          <button
            className="pn-del"
            title="另存为…"
            onClick={async (e) => {
              e.stopPropagation()
              const bytes = await readBytes(r)
              if (!bytes || !api?.docSaveAs) { flashMsg('桌面桥缺失（重启应用）'); return }
              const saved = await api.docSaveAs(r.name, bytes, 'docx')
              if (saved) pushLog(`[PDF转Word] 已另存：${saved}`, 'ok')
            }}
          >
            <IcDownload size={12} />
          </button>
          <button
            className="pn-del"
            title="删除记录（文件字节保留在本机库）"
            onClick={(e) => { e.stopPropagation(); removeP2wHist(nodeId, r.id) }}
          >
            <IcX size={12} />
          </button>
        </div>
      ))}
      {flash && <div className="pn-msg">{flash}</div>}
    </div>
  )
}

/* ═══════════ 节点卡片主体（仿 PDF 替换页：分区头锚点 + 词条行） ═══════════ */

export function Pdf2WordNodeCard ({ nodeId, lod }: { nodeId: string; lod?: boolean }) {
  useSuiteState() /* 套件停用/启用 → 缺套件引导行即时显隐 */
  const noSuite = !hasCapability('convert:pdf2docx') && !hasCapability('convert:pdf2docx-ocr')
  const tables = useBaseStore((s) => s.tables)
  const lcfg = (tables[nodeId]?.config ?? {}) as { dataLocked?: boolean; ooDataLocked?: boolean }
  const dataLocked = !!(lcfg.dataLocked || lcfg.ooDataLocked)
  const cfg = (tables[nodeId]?.config ?? {}) as { p2wItems?: P2wItem[]; p2wOrder?: string[] }
  const raw = cfg.p2wItems ?? []
  const items = orderedP2w(raw, cfg.p2wOrder ?? null)
  /* ★ 拖动 LOD：拖动中渲染轻量骨架（一行文字），松手恢复完整卡片 */
  if (lod) return <div className="pr-root pn-lod" style={{ height: Math.max(120, raw.length * 34 + 70) }} />
  const [msg, setMsg] = useState('')
  const flash = (m: string) => { setMsg(m); setTimeout(() => setMsg(''), 4000) }
  const dragDepth = useRef(0)
  const [dragOver, setDragOver] = useState(false)
  const fileRef = useRef<HTMLInputElement>(null)
  const dragIdx = useRef<number | null>(null)
  const [dragging, setDragging] = useState<number | null>(null)
  const [overIdx, setOverIdx] = useState<number | null>(null)
  const [droppedIdx, setDroppedIdx] = useState<number | null>(null)
  /* ★ 锚点动态增删重测量（连线端点不跳——pdfmerge 教训） */
  const updateNodeInternals = useUpdateNodeInternals()
  useEffect(() => { updateNodeInternals(nodeId) }, [raw.length, nodeId, updateNodeInternals])

  const importFiles = async (files: FileList | null) => {
    const r = await importP2wFiles(nodeId, files)
    if (r.skipped) flash(`跳过 ${r.skipped} 个（仅支持 pdf，单份 ≤200MB）`)
  }

  /** 拖动中悬停换位：id 序写 config.p2wOrder（持久，删除/重排皆稳定） */
  const swapTo = (from: number, to: number) => {
    if (dataLocked) return /* ★ 统一数据锁：锁定禁排序 */
    if (from === to || from < 0 || to < 0 || from >= items.length || to >= items.length) return
    const ids = items.map((it) => it.id)
    const [moved] = ids.splice(from, 1)
    ids.splice(to, 0, moved)
    setCfg({ p2wOrder: ids })
    dragIdx.current = to
  }

  const removeAt = (i: number) => {
    const victim = items[i]
    if (!victim) return
    if (dataLocked) { flash('数据已锁定：先解锁再删除'); return } /* ★ 统一数据锁 */
    useBaseStore.getState().pushHistory()
    setCfg({
      p2wItems: raw.filter((it) => it.id !== victim.id),
      p2wOrder: (cfg.p2wOrder ?? []).filter((id) => id !== victim.id),
    })
  }

  const setCfg = (patch: Record<string, unknown>) => useBaseStore.getState().setNodeConfig(nodeId, patch)

  return (
    <div
      className="pr-root"
      onDoubleClick={(e) => e.stopPropagation()}
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
      {/* 分区头（仿替换页 pr-group-head）：行首锚点 = 上游连线追加到本组（Canvas p2w-add 分支） */}
      <div style={{ position: 'relative' }} className={'pr-group-head' + (dragOver ? ' pr-group-head--over' : '')} title="待转换的 PDF（文字版秒转 / 扫描件自动走 OCR）；行首锚点=上游连线追加到本组">
        <Handle
          type="target"
          position={Position.Left}
          id="p2w-add"
          className="nodrag nopan pr-anchor"
          style={{ position: 'absolute', left: -9, top: '50%', transform: 'translateY(-50%)', width: 9, height: 9, background: '#3b82f6', border: '1.5px solid var(--c-bg0)', borderRadius: '50%' }}
          title="连接上游 PDF（PDF/瘦身/合并/替换页产物），追加到待转组"
        />
        待转换 PDF（产物在按钮行 🕘 历史）{dragOver ? ' —— 松手加入' : ''}
      </div>
      <div className={'pr-group' + (dragOver ? ' pr-group--over' : '')}>
        <input
          ref={fileRef}
          type="file"
          accept="application/pdf,.pdf"
          multiple
          style={{ display: 'none' }}
          onChange={async (e) => { await importFiles(e.target.files); e.target.value = '' }}
        />
        {items.length === 0 && (
          <button className="pn-empty-btn nodrag nopan" title={dataLocked ? '数据已锁定：先解锁再导入' : '点击导入 PDF（可多选，≤200MB/份）'} disabled={dataLocked} onClick={() => fileRef.current?.click()}>
            ＋ 点击导入 PDF（也可拖入/连上游）
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
            style={dragging === i ? { opacity: 0.35 } : undefined}
            draggable
            onDragStart={(e) => {
              e.stopPropagation()
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
            title={`${it.name}（拖动调整顺序=转换顺序；行首锚点=上游覆盖本条）`}
          >
            {/* 行首锚点（仿替换页）：上游连线 → 覆盖本条（Canvas p2w-<id> 分支） */}
            <Handle
              type="target"
              position={Position.Left}
              id={'p2w-' + it.id}
              className="nodrag nopan pr-anchor"
              style={{ position: 'absolute', left: -9, top: '50%', transform: 'translateY(-50%)', width: 9, height: 9, background: '#3b82f6', border: '1.5px solid var(--c-bg0)', borderRadius: '50%' }}
              title="连接上游 PDF，覆盖本位置"
            />
            <span className="pn-idx">{i + 1}</span>
            <span className="pm-fmt pm-fmt-pdf">P</span>
            <span
              className={'pn-loc ' + ((it.inDb || it.b64) ? 'pn-loc--in' : 'pn-loc--ext')}
              title={(it.inDb || it.b64) ? '软件内部：字节已入库，随工程文件存储' : '外部引用：字节不在软件库，可能丢失'}
            >{(it.inDb || it.b64) ? '内' : '外'}</span>
            <span className="pn-nm">{it.name}</span>
            <button className="pn-del" title="移除这份" onClick={() => removeAt(i)}>×</button>
          </div>
        ))}
      </div>
      {/* ★ 节点接入规范·缺套件引导：能力全缺时卡片黄条引导进套件中心（可点击） */}
      {noSuite && (
        <button
          className="pn-row nodrag nopan"
          title="点击打开套件中心"
          onClick={(e) => { e.stopPropagation(); suiteCenterUI.open() }}
        >
          <span className="pn-loc pn-loc--ext">⚠</span>
          <span className="props-file-nm">缺转换套件——点击打开套件中心安装/启用</span>
        </button>
      )}
      {msg && <div className="pn-msg">{msg}</div>}
    </div>
  )
}
