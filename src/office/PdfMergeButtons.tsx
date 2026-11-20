/**
 * PDF 合并节点 · 行式列表 + 方形合并按钮（2026-10-23 用户指令
 * 「合并PDF按钮 做成圆角矩形正方形 放在和锁定组并排的位置」）
 *
 * 结构：
 *   ① 节点首行（与锁定组并排）：方形「PDF」图标按钮（31×31，同 GrpLockButton 语言）
 *      ——点击触发 runPdfMerge（四环节日志 + 节点徽章全程打点）
 *   ② 文件名行（默认 合并-MMDD-HHmm，失焦持久 config.pdfName）
 *   ③ 上游文档行：每行一份（序号 + 格式徽章 + 名称 + 拖把），拖拽排序持久
 *
 * 链路：上游字节 → /pdfconv（DS converter 长驻 x2t）/ 照片 pdf-lib 直转
 * → pdf-lib 合并 → 按文件名下载。
 */
import {useEffect, useRef, useState, useSyncExternalStore, useMemo } from 'react'
import { createPortal } from 'react-dom'
import { Handle, Position, useUpdateNodeInternals } from '@xyflow/react'
import { useBaseStore } from '../core/store'
import { getBlob } from '../core/pdfBlobStore'
import { mergePdfs } from '../core/pdfMerge'
import { getTemplateStore } from '../docx/templateStore'
import { pushLog } from '../ui/logBus'
import { IcDownload, IcHistory, IcMaximize, IcMinimize, IcX } from '../ui/icons'
import { PdfViewer } from './PdfViewer'
import { setNodeRunStatus, setNodeRunStatusAuto, setPrimaryRunNode } from '../core/nodeRunStatus'
import { orderedPhotos } from './PhotoNodeCard'
import { orderedPdfs } from './PdfNodeCard'

/* ══ 2026-10-23 用户指令「PDF合并下方再增加按钮：pdf存于画布 + 历史按钮」══
   存画布：合并产物 base64 内嵌进节点 config.pdfHistory 条目（随画布走；
   14:37 二次改造——用户指令「不要生成PDF节点」→ 不建节点不下载）
   历史：每次合并产物入库（模板库 IndexedDB）+ config.pdfHistory 索引（cap 10）；
        「历史」按钮展开列表（按钮行正下方），双击条目 → 居中大窗预览 */

/** 历史面板展开态（模块级轻量 store——按钮在 Canvas 首行、列表在卡片内，跨组件共享） */
let histOpen = false
const histSubs = new Set<() => void>()
export const pdfHistUI = {
  get: () => histOpen,
  toggle: () => { histOpen = !histOpen; histSubs.forEach((f) => f()) },
  sub: (f: () => void) => { histSubs.add(f); return () => { histSubs.delete(f) } },
}

/** pdfmerge 每行的入口锚点 id（Canvas 边派生按此对齐——锚点对照每行） */
export const pmRowHandleId = (upstreamId: string) => 'pm-in-' + upstreamId

/** 上游行格式徽章（用户指令「序号后面文件名前面加各种文件格式的logo」）：
 *  ooword=蓝W / oosheet=绿X / photo=青色图片图标——与 Office 品牌色同语言 */
function fmtBadgeOf (kind: string): { cls: string; text: string; title: string } {
  if (kind === 'ooword') return { cls: 'pm-fmt pm-fmt-word', text: 'W', title: 'Word 文档 (docx)' }
  if (kind === 'oosheet') return { cls: 'pm-fmt pm-fmt-sheet', text: 'X', title: 'Excel 工作簿 (xlsx)' }
  if (kind === 'photo') return { cls: 'pm-fmt pm-fmt-photo', text: '▣', title: '照片组 (jpg/png)' }
  if (kind === 'pdf') return { cls: 'pm-fmt pm-fmt-pdf', text: 'P', title: 'PDF 文件组 (pdf)' }
  return { cls: 'pm-fmt', text: '?', title: kind }
}

const DEFAULT_NAME = () => '合并-' + new Date().toISOString().slice(0, 16).replace('T', '-').replace(':', '')

/** 上游节点按「输入口顺序」取（连线顺序即基础排序），config.pdfOrder 覆盖 */
function upstreamDocs (
  tables: Record<string, { id: string; name?: string; kind?: string; deleted?: boolean; config?: { inputs?: string[]; pdfOrder?: string[] } }>,
  nodeId: string,
) {
  const self = tables[nodeId]
  const ids = (self?.config?.inputs ?? []) as string[]
  const order = (self?.config?.pdfOrder ?? null) as string[] | null
  const byId = new Map(ids.map((id) => [id, tables[id]]))
  const seq = order ? [...order.filter((id) => ids.includes(id)), ...ids.filter((id) => !order.includes(id))] : ids
  return seq
    .map((id) => byId.get(id))
    .filter((t): t is NonNullable<typeof t> => !!t && !t.deleted && (t.kind === 'ooword' || t.kind === 'oosheet' || t.kind === 'photo' || t.kind === 'pdf'))
}

/** 读一个 OO 节点的当前字节（节点内置优先，模板库兜底） */
async function readOoBytes (t: { id: string; config?: { ooTemplateId?: string; ooBytes64?: string; ooBytesDbId?: string } }): Promise<Uint8Array | null> {
  /* ★★★ 2026-11-07 用户「执行转换」报「内容已损坏（0B）」：
     新版文档字节存 IndexedDB（config.ooBytesDbId → pdfBlobStore），这里却只认
     ooBytes64（旧内联 base64）和 ooTemplateId（模板库文件，模板被删后 read 直接抛）——
     文件明明在库里（实测 42KB），却被判成 0B 损坏。
     取值顺序改为「新存储 → 旧内联 → 模板库」，与写入侧的演进方向一致。 */
  const dbId = t.config?.ooBytesDbId
  if (dbId) {
    try {
      const b = await getBlob(dbId)
      if (b && b.length) return b
    } catch { /* IndexedDB 不可用 → 退旧路径 */ }
  }
  const b64 = t.config?.ooBytes64
  if (b64) {
    try {
      const bin = atob(b64)
      const arr = new Uint8Array(bin.length)
      for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i)
      return arr
    } catch { /* 坏 base64 走模板库 */ }
  }
  const tid = t.config?.ooTemplateId
  if (!tid) return null
  try {
    return await getTemplateStore().read(tid)
  } catch {
    return null
  }
}

/** docx/xlsx → PDF：服务端 /pdfconv（写工作区 → DS converter 长驻 x2t → 产物回传）。
 *  ★ 2026-10-23 修「转换会话未就绪」：原先 window.__ooConvertPdf 挂在 /poc iframe 的
 *  window 上宿主调不到（恒 undefined）。服务端通道不依赖编辑器是否打开。 */
/** docx/xlsx → PDF（/pdfconv 端点，DS 长驻 x2t 秒转）。
 *  ★ 2026-10-29 导出：右键 OO 节点「转为 PDF 节点」复用同一转换管道。 */
export async function convertToPdf (bytes: Uint8Array, name: string): Promise<Uint8Array | null> {
  const ext = /\.xlsx$/i.test(name) ? 'xlsx' : 'docx'
  const convName = 'oo-pdfconv-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 6) + '.' + ext
  let b64 = ''
  try {
    const CHUNK = 0x8000
    for (let i = 0; i < bytes.length; i += CHUNK) b64 += String.fromCharCode(...bytes.subarray(i, i + CHUNK))
  } catch { throw new Error('字节转 base64 失败') }
  /* ★ 2026-10-23 乱码真凶修复：b64 此刻是「每字节一字符」的二进制字符串，
     必须再 btoa 成真 base64——此前漏了这一步，上传的是二进制原文，
     服务端 base64 解码只解出碎片（22B/64B）→ x2t 转碎片 → 乱码。 */
  const r = await fetch('/pdfconv', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: convName, dataBase64: btoa(b64) }),
  }).then((x) => x.json()).catch((e) => ({ ok: false, err: String(e) }))
  if (r && r.ok && r.dataBase64) {
    const bin = atob(r.dataBase64)
    const arr = new Uint8Array(bin.length)
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i)
    return arr
  }
  throw new Error((r && r.err) || '转换失败')
}

/** 照片组 → PDF（每图一页 A4 等比适配居中）：pdf-lib 直转，零服务依赖 */
async function photosToPdf (items: { id: string; name: string; type: string; b64?: string; photoDbId?: string }[]): Promise<Uint8Array> {
  const { PDFDocument } = await import('pdf-lib')
  const pdf = await PDFDocument.create()
  const A4W = 595.28, A4H = 841.89
  const M = 24
  for (const it of items) {
    /* ★ 2026-10-24 照片内容寻址：photoDbId 优先取库，b64 兼容旧画布 */
    let u8: Uint8Array
    if (it.photoDbId) {
      const { getBlob } = await import('../core/pdfBlobStore')
      const stored = await getBlob(it.photoDbId)
      if (!stored) throw new Error(`照片数据缺失（${it.name}）`)
      u8 = stored
    } else if (it.b64) {
      const bin = atob(it.b64)
      u8 = new Uint8Array(bin.length)
      for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
    } else {
      throw new Error(`照片无数据（${it.name}）`)
    }
    const emb = it.type === 'png' ? await pdf.embedPng(u8) : await pdf.embedJpg(u8)
    const page = pdf.addPage([A4W, A4H])
    const scale = Math.min((A4W - M * 2) / emb.width, (A4H - M * 2) / emb.height)
    page.drawImage(emb, {
      x: (A4W - emb.width * scale) / 2,
      y: (A4H - emb.height * scale) / 2,
      width: emb.width * scale,
      height: emb.height * scale,
    })
  }
  const out = await pdf.save()
  return new Uint8Array(out)
}

/** ★ 合并执行核心（2026-10-23 抽取）：方形按钮与行式面板共用同一逻辑——
 *  四环节日志（获取/转化/合并/保存）+ 节点徽章打点 + 坏内容防御 + 按文件名下载。 */
export async function runPdfMerge (nodeId: string, opts?: { saveToCanvas?: boolean; onProgress?: (pct: number) => void }): Promise<{ ok: boolean; msg: string }> {
  const saveToCanvas = !!opts?.saveToCanvas
  const stAll = () => useBaseStore.getState()
  const tables = stAll().tables as never
  const cfg = (stAll().tables[nodeId]?.config ?? {}) as { pdfName?: string; pdfOrder?: string[]; pdfHistory?: { id: string; name: string; at: number; b64?: string; inDb?: boolean; color?: string }[] }
  /* ★ 2026-10-24 修复 React #321：mergeRun 是 async 普通函数（非组件），
     此前误用 useMemo——非法 hook call，每次合并即抛 #321。直接调用即可。 */
  const docs = upstreamDocs(tables, nodeId)
  if (!docs.length) return { ok: false, msg: '没有上游 OO 文档——请从 ooword/oosheet 连入' }
  const t0 = Date.now()
  pushLog(`[PDF] 开始合并：${docs.length} 份上游（${docs.map((d) => String(d.name ?? d.id)).join('、')}）`)
  /* ★ 2026-11-09 运行状态条主/次分层：本节点为主进程（上），上游转换/装载为次进程（下） */
  setPrimaryRunNode(nodeId)
  if (nodeId) setNodeRunStatus(nodeId, 'running', '合并中…')
  try {
    const pdfs: Uint8Array[] = []
    let doneDocs = 0
    const prog = (frac: number) => { try { opts?.onProgress?.(Math.round(frac * 100)) } catch { /* ignore */ } }
    prog(0.02)
    for (const d of docs) {
      if (d.kind === 'photo') {
        const pc = (d as unknown as { config?: { photoItems?: { id: string; name: string; type: string; b64?: string; photoDbId?: string }[]; photoOrder?: string[] } }).config
        const items = orderedPhotos(pc?.photoItems ?? [], pc?.photoOrder ?? null)
        if (!items.length) {
          if (nodeId) setNodeRunStatus(nodeId, 'error', `${d.name} 无照片`)
          return { ok: false, msg: `${d.name} 没有照片（先在节点上导入）` }
        }
        setNodeRunStatus(d.id, 'running', '生成照片页…')
        const tp = Date.now()
        try {
          const photoPdf = await photosToPdf(items)
          pdfs.push(photoPdf)
          pushLog(`[PDF·转化] ${d.name}：${items.length} 张照片 → ${photoPdf.length}B（${((Date.now() - tp) / 1000).toFixed(1)}s）`, 'ok')
          setNodeRunStatusAuto(d.id, 'ok', `${items.length} 张照片`)
        } catch (eP) {
          setNodeRunStatus(d.id, 'error', '照片转换失败')
          pushLog(`[PDF·转化] ${d.name} 失败: ${String((eP as Error)?.message ?? eP).slice(0, 80)}`, 'err')
          return { ok: false, msg: `${d.name} 照片转 PDF 失败: ${String((eP as Error)?.message ?? eP).slice(0, 80)}` }
        }
        continue
      }
      /* ★ 2026-10-23 pdf 上游：PDF 文件组按行序直接拼接（pdf-lib load+append，零转换） */
      if (d.kind === 'pdf') {
        const pcc = (d as unknown as { config?: { pdfItems?: { id: string; name: string; b64: string }[]; pdfOrder?: string[] } }).config
        const pdfItems = orderedPdfs(pcc?.pdfItems ?? [], pcc?.pdfOrder ?? null)
        if (!pdfItems.length) {
          if (nodeId) setNodeRunStatus(nodeId, 'error', `${d.name} 无 PDF`)
          return { ok: false, msg: `${d.name} 没有 PDF（先在节点上导入）` }
        }
        setNodeRunStatus(d.id, 'running', '装载 PDF…')
        const tp2 = Date.now()
        try {
          const { PDFDocument } = await import('pdf-lib')
          const one = await PDFDocument.create()
          let nPages = 0
          for (const it of pdfItems) {
            let u8: Uint8Array
            if (it.b64) {
              const bin = atob(it.b64)
              u8 = new Uint8Array(bin.length)
              for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
            } else {
              const fromDb = await getBlob('pdfitem-' + it.id)
              if (!fromDb) throw new Error('文件数据缺失: ' + (it.name || it.id))
              u8 = fromDb
            }
            const src = await PDFDocument.load(u8, { ignoreEncryption: true })
            const pages = await one.copyPages(src, src.getPageIndices())
            for (const p of pages) one.addPage(p)
            nPages += pages.length
          }
          const pdfBytes = new Uint8Array(await one.save())
          pdfs.push(pdfBytes)
          pushLog(`[PDF·转化] ${d.name}：${pdfItems.length} 份 PDF / ${nPages} 页 → ${pdfBytes.length}B（${((Date.now() - tp2) / 1000).toFixed(1)}s）`, 'ok')
          setNodeRunStatusAuto(d.id, 'ok', `${pdfItems.length} 份 PDF`)
        } catch (eP2) {
          setNodeRunStatus(d.id, 'error', 'PDF 装载失败')
          pushLog(`[PDF·转化] ${d.name} 失败: ${String((eP2 as Error)?.message ?? eP2).slice(0, 80)}`, 'err')
          return { ok: false, msg: `${d.name} PDF 装载失败（文件可能加密或损坏）` }
        }
        continue
      }
      /* ① 获取文档 */
      const tGet = Date.now()
      setNodeRunStatus(d.id, 'running', '读取中…')
      const bytes = await readOoBytes(d as never)
      pushLog(
        `[PDF·获取] ${d.name}：${bytes ? bytes.length + 'B' : '读不到'}（${((Date.now() - tGet) / 1000).toFixed(1)}s）`,
        bytes ? undefined : 'err',
      )
      if (!bytes || bytes.length < 200 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
        pushLog(`[PDF·获取] ${d.name} 内容已损坏（${bytes?.length ?? 0}B 非有效 docx，无 PK 头）——请重新导入`, 'err')
        setNodeRunStatus(d.id, 'error', '内容已损坏')
        return { ok: false, msg: `${d.name} 的内容已损坏（${bytes?.length ?? 0}B 非有效文档）——请重新导入该文档后再合并` }
      }
      setNodeRunStatusAuto(d.id, 'ok', `读取 ${bytes.length}B`)
      /* ② 转化文档 */
      setNodeRunStatus(d.id, 'running', '转换 PDF…')
      const tCv = Date.now()
      pushLog(`[PDF·转化] ${d.name}：上传 base64 ${Math.ceil(bytes.length / 3) * 4} 字符`)
      const pdf = await convertToPdf(bytes, String(d.name ?? d.id))
      if (!pdf) {
        pushLog(`[PDF·转化] ${d.name} 失败（转换器无产物）`, 'err')
        setNodeRunStatus(d.id, 'error', '转换失败')
        return { ok: false, msg: `${d.name} 转 PDF 失败` }
      }
      pushLog(`[PDF·转化] ${d.name}：docx → PDF ${pdf.length}B（${((Date.now() - tCv) / 1000).toFixed(1)}s）`, 'ok')
      setNodeRunStatusAuto(d.id, 'ok', `PDF ${(pdf.length / 1024).toFixed(1)}KB · ${((Date.now() - tCv) / 1000).toFixed(1)}s`)
      pdfs.push(pdf)
    }
    prog(0.9)
    /* ③ 合并文档 */
    const tMg = Date.now()
    const merged = await mergePdfs(pdfs)
    if (!merged.ok || !merged.bytes) {
      pushLog(`[PDF·合并] 失败：${merged.why ?? '未知原因'}`, 'err')
      if (nodeId) setNodeRunStatus(nodeId, 'error', '合并失败')
      return { ok: false, msg: merged.why ?? '合并失败' }
    }
    const totalPages = merged.pages.reduce((a2, b2) => a2 + b2, 0)
    pushLog(`[PDF·合并] ${merged.count} 份 / ${totalPages} 页 → ${merged.bytes.length}B（${((Date.now() - tMg) / 1000).toFixed(1)}s）`, 'ok')
    if (nodeId) setNodeRunStatusAuto(nodeId, 'ok', `${merged.count} 份 · ${totalPages} 页 · ${((Date.now() - t0) / 1000).toFixed(1)}s`)
    /* ④ 保存文档（2026-10-23 二次改造：saveToCanvas=产物 base64 随节点 config 走
       （历史条目带 b64，不建节点不下载——用户指令「不要生成PDF节点」）；
       下载=产物落磁盘，历史条目无 b64（来源标记「外部」，预览走模板库回退） */
    const fname = (cfg.pdfName && cfg.pdfName.trim()) || DEFAULT_NAME()
    const download = fname.endsWith('.pdf') ? fname : fname + '.pdf'
    let canvasB64: string | undefined
    if (saveToCanvas) {
      let b64s = ''
      const CHUNK = 0x8000
      for (let i = 0; i < merged.bytes.length; i += CHUNK) b64s += String.fromCharCode(...merged.bytes.subarray(i, i + CHUNK))
      canvasB64 = btoa(b64s)
      pushLog(`[PDF·保存] ${download}（${merged.bytes.length}B）已保存于画布（历史条目内嵌）；全程 ${((Date.now() - t0) / 1000).toFixed(1)}s`, 'ok')
    } else {
      const blob = new Blob([merged.bytes as BlobPart], { type: 'application/pdf' })
      const url = URL.createObjectURL(blob) + '#zoom=100'
      const a = document.createElement('a')
      a.href = url
      a.download = download
      a.click()
      setTimeout(() => URL.revokeObjectURL(url), 3000)
      pushLog(`[PDF·保存] ${download}（${merged.bytes.length}B）已下载；全程 ${((Date.now() - t0) / 1000).toFixed(1)}s`, 'ok')
    }
    /* ⑤ 历史：产物入库（模板库，外部条目的预览回退通道）+ 节点 config 索引
       （新在前 cap 10；画布条目带 b64，来源可区分） */
    try {
      const meta = await getTemplateStore().save(download, merged.bytes)
      /* ★ 2026-10-24 性能根治：去掉 b64 双份（产物已在模板库，config 只留索引）——
         config 从几十 MB 瘦身到 KB 级，拖停保存序列化不再阻塞。旧条目 b64 照常可读。 */
      const before = cfg.pdfHistory ?? []
      const hist = [ { id: meta.id, name: download, at: Date.now(), inDb: true }, ...before ].slice(0, 10)
      stAll().setNodeConfig(nodeId, { pdfHistory: hist })
      /* ★ 挤出列表的行级锚点连线同步断开（keep=hist 内仍存在的 key） */
      const keep = new Set(hist.map((x) => 'hist-' + x.id + '-' + (x.at ?? 0)))
      const evicted = before.filter((x) => !keep.has('hist-' + x.id + '-' + (x.at ?? 0))).map((x) => 'hist-' + x.id + '-' + (x.at ?? 0))
      if (evicted.length) window.dispatchEvent(new CustomEvent('pdfmerge-prune-edges', { detail: { nodeId, keys: evicted } }))
      pushLog(`[PDF·历史] 已记录（共 ${hist.length} 条）——历史按钮可展开，双击条目预览`, undefined)
    } catch (eH) {
      pushLog(`[PDF·历史] 记录失败: ${String((eH as Error)?.message ?? eH).slice(0, 80)}`, 'warn')
    }
    stAll().pushHistory()
    prog(1)
    return { ok: true, msg: `已合并 ${merged.count} 份 / ${totalPages} 页 → ${download}${saveToCanvas ? '（存于画布）' : ''}` }
  } catch (e) {
    const msg = String((e as Error)?.message ?? e)
    pushLog(`[PDF] 异常: ${msg.slice(0, 100)}`, 'err')
    if (nodeId) setNodeRunStatus(nodeId, 'error', '合并异常')
    return { ok: false, msg }
  }
}

/** 方形「PDF」合并按钮（2026-10-23 用户指令「圆角矩形正方形 放在和锁定组并排的位置」）：
 *  31×31 与 GrpLockButton 同语言；busy 时转圈禁点；结果走日志与节点徽章。 */
export function PdfMergeRunButton ({ nodeId }: { nodeId: string }) {
  const [busy, setBusy] = useState(false)
  /* ★ 统一数据锁（2026-10-26）：锁定时禁执行（执行=追加合并历史=内容变更） */
  const lcfg = useBaseStore((s) => s.tables[nodeId]?.config) as { dataLocked?: boolean; ooDataLocked?: boolean } | undefined
  const locked = !!(lcfg?.dataLocked || lcfg?.ooDataLocked)
  /* ★ 2026-10-25 用户反馈「连线起点在下面」：Handle 位置变更后 RF 缓存未更新——
     mount 时强制重测，连线起点精确落在金锚点上 */
  const updateNodeInternals = useUpdateNodeInternals()
  useEffect(() => { updateNodeInternals(nodeId) }, [nodeId, updateNodeInternals])
  const onClick = async () => {
    if (busy || locked) return
    setBusy(true)
    try {
      const r = await runPdfMerge(nodeId, {
        onProgress: (pct) => useBaseStore.getState().setPdfmergeProgress(nodeId, pct),
      })
      if (!r.ok) pushLog(`[PDF] ${r.msg}`, 'err')
    } finally {
      setBusy(false)
    }
  }
  return (
    <div style={{ position: 'relative', display: 'inline-flex', alignItems: 'center' }}>
      <button
        className="pm-sq-btn nodrag nopan"
        title={locked ? '数据已锁定：先解锁再执行合并' : '合并 PDF（按行序把上游转 PDF 拼接并下载）'}
        disabled={busy || locked}
        onClick={onClick}
      >
        {busy ? <span className="rn-spin pm-spin-dark" /> : <span className="pm-sq-txt">PDF</span>}
      </button>
      {/* ★ 2026-10-25 用户指令「右边的锚点 连接上就连接上 不要跳」：
          Handle 挂在稳定包装层（不受合并 spinner 重渲染影响）——端点不跳 */}
      <Handle
        type="source"
        position={Position.Right}
        id="hist-latest"
        className="nodrag nopan pm-hist-anchor"
        style={{ position: 'static', marginLeft: 26, width: 10, height: 10, background: '#d29922', border: '1.5px solid #fff', transform: 'none' }}
        title="最新合并产物（连线即采用最新的）"
      />
    </div>
  )
}

/** 存画布 / 历史按钮（2026-10-23 用户指令「PDF合并下方再增加一个按钮 pdf存于画布
 *  再增加个历史按钮」）：31px 长方形与首行按钮同语言，渲染在首行下方新行。 */
/** 存画布 / 历史按钮（2026-10-23 用户指令「做成一行 都是正方形的圆角的」）：
 *  31×31 方形圆角图标钮（同 GrpLockButton/PdfMergeRunButton 语言），与它们并排一行。 */
export function PdfExtraButtons ({ nodeId }: { nodeId: string }) {
  const [busy, setBusy] = useState(false)
  const histOpenNow = useSyncExternalStore(pdfHistUI.sub, pdfHistUI.get)
  const onSaveCanvas = async () => {
    if (busy) return
    setBusy(true)
    try {
      const r = await runPdfMerge(nodeId, { saveToCanvas: true })
      if (!r.ok) pushLog(`[PDF] ${r.msg}`, 'err')
    } finally {
      setBusy(false)
    }
  }
  return (
    <>
      <button
        className="pm-sq-btn pm-sq-btn--ghost nodrag nopan"
        disabled={busy}
        title="合并结果保存于画布（内嵌历史，随画布走，不下载）"
        onClick={onSaveCanvas}
      >
        {busy ? <span className="rn-spin pm-spin-dark" style={{ width: 10, height: 10 }} /> : <IcDownload size={19} />}
      </button>
      <button
        className={'pm-sq-btn pm-sq-btn--ghost nodrag nopan' + (histOpenNow ? ' pm-sq-btn--on' : '')}
        title="展开/收起合并历史（双击历史条目进入预览）"
        onClick={() => pdfHistUI.toggle()}
      >
        <IcHistory size={19} />
      </button>
    </>
  )
}

/** 行式列表面板：文件名行 + 上游文档行（拖拽排序）。合并按钮在节点首行（见 PdfMergeRunButton）。 */
export function PdfMergeButtons ({ nodeId, lod }: { nodeId: string; lod?: boolean }) {
  /* ★ 拖动 LOD：拖动中渲染轻量骨架，松手恢复完整卡片 */
  if (lod) return <div className="pm-root pn-lod" style={{ height: 330 }} />
  const tables = useBaseStore((s) => s.tables)
  const dragIdx = useRef<number | null>(null)
  const [dragging, setDragging] = useState<number | null>(null)
  const [overIdx, setOverIdx] = useState<number | null>(null)
  const [droppedIdx, setDroppedIdx] = useState<number | null>(null)

  const self = tables[nodeId]
  const cfg = (self?.config ?? {}) as { pdfName?: string; pdfOrder?: string[]; pdfHistory?: { id: string; name: string; at: number; b64?: string; inDb?: boolean; color?: string }[] }
  const stAll = () => useBaseStore.getState()
  const setCfg = (patch: Record<string, unknown>) => stAll().setNodeConfig(nodeId, patch)
  const docs = useMemo(() => upstreamDocs(tables as never, nodeId), [tables, nodeId])
  const flash = (m: string) => { pushLog(`[PDF] ${m}`, 'warn') }
  /* 文件名本地草稿（每键不写 store），失焦/回车提交持久化 */
  const [nameDraft, setNameDraft] = useState(cfg.pdfName ?? '')
  useEffect(() => { setNameDraft(cfg.pdfName ?? '') }, [cfg.pdfName])
  const commitName = () => {
    if (nameDraft !== (cfg.pdfName ?? '')) setCfg({ pdfName: nameDraft })
  }

  /** 拖动中悬停换位：id 序写 config.pdfOrder（持久，删除/重排皆稳定） */
  const swapTo = (from: number, to: number) => {
    if (from === to || from < 0 || to < 0 || from >= docs.length || to >= docs.length) return
    const ids = docs.map((d) => d.id)
    const [moved] = ids.splice(from, 1)
    ids.splice(to, 0, moved)
    setCfg({ pdfOrder: ids })
    dragIdx.current = to
  }

  /** ★ 2026-10-23 用户指令「增加类似照片的可以关闭的地方」：行尾 × 移除该上游——
   *  同时清 inputs / linkedInputs（连线派生源，清后画布连线自动断开）/ pdfOrder。 */
  const removeUpstream = (id: string) => {
    const c = stAll().tables[nodeId]?.config as { inputs?: string[]; linkedInputs?: string[]; pdfOrder?: string[] } | undefined
    stAll().pushHistory()
    stAll().setNodeConfig(nodeId, {
      inputs: (c?.inputs ?? []).filter((x) => x !== id),
      linkedInputs: (c?.linkedInputs ?? []).filter((x) => x !== id),
      pdfOrder: (c?.pdfOrder ?? []).filter((x) => x !== id),
    })
    pushLog(`[PDF] 移除上游：${String(tables[id]?.name ?? id)}`)
  }

  return (
    <div className="pm-root nodrag nopan">
      {/* ① 文件名行：默认「合并-MMDD-HHmm」，失焦/回车持久化（config.pdfName） */}
      <div className="pm-name-row nodrag nopan" title="合并 PDF 的保存文件名（随节点持久化）">
        <span className="pm-name-label">文件名</span>
        <input
          className="pm-name-input"
          value={nameDraft}
          placeholder={DEFAULT_NAME()}
          onChange={(e) => setNameDraft(e.target.value)}
          onBlur={commitName}
          onKeyDown={(e) => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur() }}
        />
      </div>
      {/* ② 上游文档行：每行一份，拖拽排序 + 行内专属锚点 + 格式徽章 */}
      {docs.length === 0 && (
        <div className="pm-empty">无上游——从 ooword / oosheet 连入</div>
      )}
      {docs.map((d, i) => (
        <div
          key={d.id}
          className={
            'pm-row nodrag nopan'
            + (dragging === i ? ' pm-row--drag' : '')
            + (overIdx === i && dragging !== i ? ' pm-row--over' : '')
            + (droppedIdx === i ? ' pm-row--dropped' : '')
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
          title={String(d.name ?? d.id) + '（拖动调整合并顺序）'}
        >
          <Handle
            type="target"
            position={Position.Left}
            id={pmRowHandleId(d.id)}
            className="pm-handle"
            title={'上游 ' + (d.name ?? d.id) + ' 的入口锚点'}
          />
          <span className="pm-idx">{i + 1}</span>
          {(() => {
            const fb = fmtBadgeOf(String(d.kind ?? ''))
            return <span className={fb.cls} title={fb.title}>{fb.text}</span>
          })()}
          <span className="pm-nm">{String(d.name ?? d.id)}</span>
          <button
            className="pm-del"
            title="移除该上游（连线同步断开）"
            onClick={(e) => { e.stopPropagation(); removeUpstream(d.id) }}
          >×</button>
          <span className="pm-grip" title="拖动排序">⠿</span>
        </div>
      ))}
    </div>
  )
}

/** 历史列表（2026-10-23 用户指令「原来的不变 按钮在下面。点击历史 历史记录在
 *  按钮的下面打开」）：独立组件——Canvas 在按钮行**之后**渲染本组件，
 *  点 🕘 展开即出现在按钮正下方；双击条目 → 居中大窗预览。
 *  来源区分：🖼 画布（产物内嵌 config 随画布走）/ 💾 外部（模板库/磁盘）。 */

export function PdfHistoryList ({ nodeId }: { nodeId: string }) {
  const histOpenNow = useSyncExternalStore(pdfHistUI.sub, pdfHistUI.get)
  const [previewBytes, setPreviewBytes] = useState<Uint8Array | null>(null)
  const [previewName, setPreviewName] = useState('')
  const [previewMax, setPreviewMax] = useState(false)
  /* ★ 同 PdfNodeCard：visibility 藏重排两帧——最大化全程正中 */
  const [frameHidden, setFrameHidden] = useState(false)
  const togglePreviewMax = () => {
    setPreviewMax((v) => !v)
    setFrameHidden(true)
    /* ★ 同 PdfNodeCard：180ms 隐藏窗 + 0.15s 淡入（重排全程不可见） */
    requestAnimationFrame(() => requestAnimationFrame(() => {
      setTimeout(() => setFrameHidden(false), 150)
    }))
  }

  /* ★ 2026-10-23 用户指令「右键 另存为」：右键条目弹小菜单，产物存到用户选的路径 */
  /* ★ 2026-10-25 锚点起手标记：Handle 按下时置位，行 dragstart 见到即拒绝
     （HTML5 dragstart 的 target 是 draggable 行本身，Handle 上拦不到——只能靠标记） */
  const histHandleDown = useRef(false)
  /* ★ 2026-10-25 用户反馈「连线跳变到中间」：历史面板展开/条数变化时
     新增的 Handle 位置 RF 缓存没更新——必须显式 updateNodeInternals 重测量 */
  const updateNodeInternals = useUpdateNodeInternals()
  const [menu, setMenu] = useState<{ x: number; y: number; entry: { id: string; name: string; at: number; b64?: string; color?: string } } | null>(null)
  /** 读条目产物字节（画布 b64 优先，模板库回退）——预览/另存/外部打开共用 */
  /* ★ 2026-10-25 用户指令「PDF 合并右键增加和 PDF 节点一致的打开菜单」：
     pjsviewer 事件桥 → 打开最新合并产物预览（复用本组件 pm-pdfpreview UI） */
  useEffect(() => {
    const h = async (e: Event) => {
      const d = (e as CustomEvent<{ nodeId: string }>).detail
      if (d?.nodeId !== nodeId) return
      const cfg = (useBaseStore.getState().tables[nodeId]?.config ?? {}) as { pdfHistory?: { id: string; name?: string; at?: number; b64?: string; inDb?: boolean }[] }
      const latest = (cfg.pdfHistory ?? [])[0]
      if (!latest) return
      try {
        const bytes = await readEntryBytes(latest)
        if (bytes) { setPreviewBytes(bytes); setPreviewName('🖼 ' + latest.name) }
      } catch { /* ignore */ }
    }
    window.addEventListener('open-pdf-pjsviewer', h)
    return () => window.removeEventListener('open-pdf-pjsviewer', h)
  }, [nodeId])

  const readEntryBytes = async (entry: { id: string; name?: string; b64?: string }): Promise<Uint8Array | null> => {
    if (entry.b64) {
      const bin = atob(entry.b64)
      const u8 = new Uint8Array(bin.length)
      for (let i2 = 0; i2 < bin.length; i2++) u8[i2] = bin.charCodeAt(i2)
      return u8
    }
    return getTemplateStore().read(entry.id)
  }
  const tables = useBaseStore((s) => s.tables)
  const cfg = (tables[nodeId]?.config ?? {}) as { pdfHistory?: { id: string; name: string; at: number; b64?: string; inDb?: boolean; color?: string }[] }
  const histCount = (cfg?.pdfHistory ?? []).length
  useEffect(() => {
    if (histOpenNow) updateNodeInternals(nodeId)
  }, [histOpenNow, histCount, nodeId, updateNodeInternals])
  const flash = (m: string) => { pushLog(`[PDF] ${m}`, 'warn') }
  if (!histOpenNow) return null
  return (
    <>
      <div className="pm-hist nodrag nopan">
        <div className="pm-hist-cap">🕘 合并历史（双击预览 · 🖼 存于画布 · 💾 外部）</div>
        {(cfg.pdfHistory ?? []).length === 0 && <div className="pn-empty">暂无历史——合并一次即记录</div>}
        {(cfg.pdfHistory ?? []).map((h, i) => (
          <div
            key={h.id + ':' + i}
            className="pm-hist-row nodrag nopan"
            title={h.name + '（双击预览；拖到画布建节点/置入已有 PDF 节点；拖出窗体=另存；× 删除记录）'}
            draggable
            onDragStart={async (eD) => {
              /* ★ 锚点起手 → 行拖拽让位（连线优先） */
              if (histHandleDown.current) { eD.preventDefault(); histHandleDown.current = false; return }
              /* ★ 2026-10-25 用户指令「历史 PDF 拖动出窗体/拖到画布/拖到已有节点」：
                 dataTransfer 只带轻量引用（key+name），字节由 drop 侧按需从 IDB 读取 */
              try {
                const bytes = await readEntryBytes(h)
                if (!bytes) { flash(`历史产物已不在本机：${h.name}`); return }
                const key = 'pdfhist-' + h.id + '-' + (h.at ?? 0)
                const { putBlob } = await import('../core/pdfBlobStore')
                await putBlob(key, bytes)
                eD.dataTransfer.setData('application/x-cb-pdf-hist', JSON.stringify({ key, name: h.name }))
                eD.dataTransfer.effectAllowed = 'copyMove'
                eD.dataTransfer.setDragImage((() => { const c = document.createElement('canvas'); c.width = 1; c.height = 1; return c })(), 0, 0)
              } catch (eD2) { flash('拖拽初始化失败: ' + String((eD2 as Error)?.message ?? eD2).slice(0, 60)) }
            }}
            onDragEnd={async (eD) => {
              /* 拖出应用窗体释放 = 另存为下载（clientX/Y 超出视口即窗外） */
              const out = eD.clientX < 0 || eD.clientY < 0 || eD.clientX > window.innerWidth || eD.clientY > window.innerHeight
              if (!out) return
              try {
                const bytes = await readEntryBytes(h)
                if (!bytes) return
                const blob = new Blob([bytes.slice().buffer as ArrayBuffer], { type: 'application/pdf' })
                const url = URL.createObjectURL(blob)
                const a = document.createElement('a')
                a.href = url; a.download = h.name || '合并产物.pdf'
                document.body.appendChild(a); a.click(); a.remove()
                setTimeout(() => URL.revokeObjectURL(url), 5000)
                pushLog(`[PDF·历史] 已另存：${h.name}`, 'ok')
              } catch { /* 忽略 */ }
            }}
            onDoubleClick={async () => {
              try {
                const bytes = await readEntryBytes(h)
                if (!bytes) { flash(`历史产物已不在本机（换机后 IndexedDB 不随画布走）：${h.name}`); return }
                setPreviewBytes(bytes)
                setPreviewName(((h.b64 || h.inDb) ? '🖼 ' : '💾 ') + h.name)
              } catch (eR) { flash(`读取失败: ${String((eR as Error)?.message ?? eR).slice(0, 60)}`) }
            }}
              onContextMenu={(e) => {
                e.preventDefault()
                e.stopPropagation()
                setMenu({ x: e.clientX, y: e.clientY, entry: h })
              }}
              onPointerDown={(e) => e.stopPropagation()}
              onPointerUp={(e) => e.stopPropagation()}
            >
            <span className="pm-idx">{i + 1}</span>
            <span className="pm-hist-color" style={{ background: h.color || 'transparent' }} title={h.color ? "颜色标记" : undefined} />
            <span className={'pm-hist-src ' + (h.b64 || h.inDb ? 'pm-hist-src--canvas' : 'pm-hist-src--ext')}>{h.b64 || h.inDb ? '内' : '外'}</span>
            <span className="pm-nm">{h.name}</span>
            {/* ★ 2026-10-23 用户指令「日期也要标记」：MM-DD HH:mm */}
            <span className="pm-hist-at">{new Date(h.at).toLocaleString('zh-CN', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })}</span>
            <button
              className="pm-del"
              title="删除这条历史记录（产物保留在本机模板库）"
              onClick={(e) => {
                e.stopPropagation()
                useBaseStore.getState().pushHistory()
                useBaseStore.getState().setNodeConfig(nodeId, {
                  pdfHistory: (cfg.pdfHistory ?? []).filter((x) => x.id !== h.id || x.at !== h.at),
                })
                pushLog(`[PDF·历史] 已删除记录：${h.name}`)
                /* ★ 锚点连线清理：被删条目的行级锚点连线同步断开 */
                window.dispatchEvent(new CustomEvent('pdfmerge-prune-edges', { detail: { nodeId, keys: ['hist-' + h.id + '-' + (h.at ?? 0)] } }))
              }}
            >×</button>
            <Handle
              type="source"
              position={Position.Right}
              id={"hist-" + h.id + "-" + (h.at ?? 0)}
              className="nodrag nopan pm-hist-anchor"
              draggable={false}
              onMouseDown={(e) => { e.stopPropagation(); histHandleDown.current = true }}
              onMouseUp={() => { histHandleDown.current = false }}
              onDragStart={(e) => { e.preventDefault(); e.stopPropagation(); histHandleDown.current = false }}
              style={{ position: "absolute", right: -7, top: "50%", transform: "translateY(-50%)", width: 9, height: 9, background: "#b02a2a", border: "1.5px solid #fff", opacity: 0.85 }}
            />
          </div>
        ))}
      </div>
      {/* 居中大窗预览（react-pdf 查看器；点遮罩/× 关闭） */}
      {previewBytes && createPortal(
        <div className="pm-pdfpreview" onClick={() => { setPreviewBytes(null); setPreviewMax(false) }}>
          <div className={'pn-preview-card pm-pdfpreview-card' + (previewMax ? ' pm-card-max' : '')} onClick={(e) => e.stopPropagation()}>
            <div className="pn-preview-bar">
              <span className="pn-preview-nm">🕘 {previewName}</span>
              <button className="pn-preview-close" title={previewMax ? '还原' : '最大化'} onClick={togglePreviewMax}>
                {previewMax ? <IcMinimize size={19} /> : <IcMaximize size={19} />}
              </button>
              <button className="pn-preview-close" title="关闭预览" onClick={() => { setPreviewBytes(null); setPreviewMax(false) }}><IcX size={19} /></button>
            </div>
            <div className="pm-preview-body">
              <PdfViewer data={previewBytes} dark={document.documentElement.dataset.theme !== 'light'} />
            </div>
          </div>
        </div>,
        document.body,
      )}
      {/* ★ 右键菜单：另存为…（Electron 系统对话框选路径；Web 降级为浏览器下载） */}
      {menu && createPortal(
        <div
          className="libpick-mask nodrag nopan"
          style={{ background: 'transparent' }}
          onClick={() => setMenu(null)}
          onContextMenu={(e) => { e.preventDefault(); setMenu(null) }}
        >
          <div
            className="pm-ctx-menu"
            style={{ left: Math.min(menu.x, window.innerWidth - 170), top: Math.min(menu.y, window.innerHeight - 60) }}
            onClick={(e) => e.stopPropagation()}
          >
            <button
              className="pm-ctx-item"
              onClick={async () => {
                const ent = menu
                setMenu(null)
                try {
                  const bytes = await readEntryBytes(ent.entry)
                  if (!bytes) { flash("产物已不在本机：" + ent.entry.name); return }
                  setPreviewBytes(bytes)
                  setPreviewName(((ent.entry.b64 || (ent.entry as { inDb?: boolean }).inDb) ? '🖼 ' : '💾 ') + ent.entry.name)
                } catch (eO) { flash("打开失败: " + String((eO as Error)?.message ?? eO).slice(0, 60)) }
              }}
            >打开</button>
            {/* ★ 2026-10-25 用户指令「历史条目右键 与 PDF 节点一致的打开菜单」：三件套 */}
            <button
              className="pm-ctx-item"
              onClick={async () => {
                const ent = menu
                setMenu(null)
                try {
                  const bytes = await readEntryBytes(ent.entry)
                  if (!bytes) { flash("产物已不在本机：" + ent.entry.name); return }
                  setPreviewBytes(bytes)
                  setPreviewName("🖼 " + ent.entry.name)
                } catch (eP) { flash("打开失败: " + String((eP as Error)?.message ?? eP).slice(0, 60)) }
              }}
            >用 PDF.js Viewer 打开</button>
            <button
              className="pm-ctx-item"
              onClick={async () => {
                const ent = menu
                setMenu(null)
                try {
                  const bytes = await readEntryBytes(ent.entry)
                  if (!bytes) { flash("产物已不在本机：" + ent.entry.name); return }
                  const api = (window as unknown as { api?: { openExternalFile?: (n: string, d: Uint8Array) => Promise<string> } }).api
                  if (!api?.openExternalFile) { flash("桌面桥不可用"); return }
                  await api.openExternalFile(ent.entry.name, bytes)
                } catch (eD) { flash("打开失败: " + String((eD as Error)?.message ?? eD).slice(0, 60)) }
              }}
            >用默认软件打开</button>
            {/* ★ 2026-10-26 用户指令「所有PDF增加个 用本地浏览器打开」：Edge/Chrome 自带查看器 */}
            <button
              className="pm-ctx-item"
              onClick={async () => {
                const ent = menu
                setMenu(null)
                try {
                  const bytes = await readEntryBytes(ent.entry)
                  if (!bytes) { flash("产物已不在本机：" + ent.entry.name); return }
                  const api = (window as unknown as { api?: { openInBrowser?: (n: string, d: Uint8Array) => Promise<string> } }).api
                  if (!api?.openInBrowser) { flash("桌面桥不可用"); return }
                  await api.openInBrowser(ent.entry.name, bytes)
                } catch (eB) { flash("浏览器打开失败: " + String((eB as Error)?.message ?? eB).slice(0, 60)) }
              }}
            >用本地浏览器打开</button>
            <button
              className="pm-ctx-item"
              onClick={async () => {
                const ent = menu
                setMenu(null)
                try {
                  const bytes = await readEntryBytes(ent.entry)
                  if (!bytes) { flash("产物已不在本机：" + ent.entry.name); return }
                  const api = (window as unknown as { api?: { openExternalFile?: (n: string, d: Uint8Array) => Promise<string> } }).api
                  if (!api?.openExternalFile) { flash("桌面桥不可用"); return }
                  /* MuPDF 直开=系统 mupdf 关联；无关联时退回默认软件 */
                  await api.openExternalFile(ent.entry.name, bytes)
                  pushLog("[PDF·历史] MuPDF 入口已按系统关联打开：" + ent.entry.name)
                } catch (eM) { flash("打开失败: " + String((eM as Error)?.message ?? eM).slice(0, 60)) }
              }}
            >用 MuPDF 打开</button>
            <button
              className="pm-ctx-item"
              onClick={async () => {
                const ent = menu
                setMenu(null)
                try {
                  const bytes = await readEntryBytes(ent.entry)
                  if (!bytes) { flash("产物已不在本机：" + ent.entry.name); return }
                  const api = (window as unknown as { api?: { docSaveAs?: (n: string, d: Uint8Array, k?: string) => Promise<string | null> } }).api
                  if (api?.docSaveAs) {
                    const p = await api.docSaveAs(ent.entry.name, bytes, 'pdf')
                    if (p) pushLog("[PDF·另存为] " + ent.entry.name + " → " + p, 'ok')
                  } else {
                    const blob = new Blob([bytes as BlobPart], { type: 'application/pdf' })
                    const u = URL.createObjectURL(blob) + '#zoom=100'
                    const a = document.createElement('a')
                    a.href = u
                    a.download = ent.entry.name
                    a.click()
                    setTimeout(() => URL.revokeObjectURL(u), 3000)
                    pushLog("[PDF·另存为] Web 环境无系统对话框，已按默认下载：" + ent.entry.name, 'warn')
                  }
                } catch (eS) { flash("另存失败: " + String((eS as Error)?.message ?? eS).slice(0, 60)) }
              }}
            >另存为…</button>
            <button
              className="pm-ctx-item"
              onClick={async () => {
                const ent = menu
                setMenu(null)
                try {
                  const bytes = await readEntryBytes(ent.entry)
                  if (!bytes) { flash("产物已不在本机：" + ent.entry.name); return }
                  const api = (window as unknown as { api?: { openExternalFile?: (n: string, d: Uint8Array) => Promise<string> } }).api
                  if (!api?.openExternalFile) { flash("Web 环境不支持系统软件打开——请用「另存为…」后本地打开"); return }
                  await api.openExternalFile(ent.entry.name, bytes)
                  pushLog("[PDF·打开] 已调系统默认软件：" + ent.entry.name, 'ok')
                } catch (eE) { flash("系统打开失败: " + String((eE as Error)?.message ?? eE).slice(0, 60)) }
              }}
            >用默认软件打开</button>
            {/* ★ 2026-10-26 用户指令「所有PDF增加个 用本地浏览器打开」：Edge/Chrome 自带查看器 */}
            <button
              className="pm-ctx-item"
              onClick={async () => {
                const ent = menu
                setMenu(null)
                try {
                  const bytes = await readEntryBytes(ent.entry)
                  if (!bytes) { flash("产物已不在本机：" + ent.entry.name); return }
                  const api = (window as unknown as { api?: { openInBrowser?: (n: string, d: Uint8Array) => Promise<string> } }).api
                  if (!api?.openInBrowser) { flash("桌面桥不可用"); return }
                  await api.openInBrowser(ent.entry.name, bytes)
                } catch (eB) { flash("浏览器打开失败: " + String((eB as Error)?.message ?? eB).slice(0, 60)) }
              }}
            >用本地浏览器打开</button>
            <div className="pm-ctx-colors" title="设置颜色标记版本">
              {(['', '#e5534b', '#e78a2e', '#d4a72c', '#57ab5a', '#539bf5', '#b083f0'].map((c) => (
                <button
                  key={c || 'none'}
                  className="pm-ctx-color"
                  style={{ background: c || 'transparent' }}
                  title={c ? '标记为该颜色' : '清除颜色标记'}
                  onClick={async () => {
                    const ent = menu
                    setMenu(null)
                    const st = useBaseStore.getState()
                    const cfgC = (st.tables[nodeId]?.config ?? {}) as { pdfHistory?: { id: string; name: string; at: number; b64?: string; inDb?: boolean; color?: string }[] }
                    st.pushHistory()
                    st.setNodeConfig(nodeId, {
                      pdfHistory: (cfgC.pdfHistory ?? []).map((x) => (x.id === ent.entry.id && x.at === ent.entry.at ? { ...x, color: c || undefined } : x)),
                    })
                    pushLog("[PDF·历史] " + (c ? "已标记颜色 " + c : "已清除颜色标记") + ":" + ent.entry.name)
                  }}
                >{c ? '' : '×'}</button>
              )))}
            </div>
          </div>
        </div>,
        document.body,
      )}
    </>
  )
}
