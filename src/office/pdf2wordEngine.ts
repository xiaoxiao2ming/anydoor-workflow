/* ─────────────────────────────────────────────────────────────
 * PDF转Word 引擎适配层（2026-10-26 用户指令「利用百度的开源 增加个节点 pdf2word」）
 *
 * 引擎 = 百度开源 PaddleOCR（飞桨）PP-Structure 文档解析管线（本机外挂，
 * python -m pip 安装，不进软件安装包——见 public/pdf2word/ 安装说明）。
 *
 * 职责：
 *   ① 词条导入（同 pdf 节点语义：字节走 IDB p2witem-{id}，config 只留索引）
 *   ② 输入收集（词条优先；无词条时收线上游产物——pdf/pdfslim 全部份按行序、
 *      pdfmerge/pdfreplace 最新历史，语义同 fileout collectFiles）
 *   ③ 检测 + 转换（主进程桥 p2w:exec spawn 本机 python；产物字节直读回传）
 *   ④ 产物入库（IDB p2whist-{id}）+ 历史索引（config.p2wHistory cap 10）
 * ───────────────────────────────────────────────────────────── */
import { useBaseStore } from '../core/store'
import { getBlob, putBlob } from '../core/pdfBlobStore'
import { pushLog } from '../ui/logBus'
import { setNodeRunStatus, setNodeRunStatusAuto } from '../core/nodeRunStatus'
import { hasCapability } from '../core/suites'
import { orderedPdfs, type PdfItem } from './PdfNodeCard'

export type P2wItem = { id: string; name: string; b64?: string; inDb?: boolean }
export type P2wHistRec = { id: string; name: string; at: number; inDb?: boolean; source?: string; engine?: 'x2t' | 'ocr' }

const P2W_MAX = 100
const P2W_MAX_BYTES = 200 * 1024 * 1024

const api = (globalThis as {
  api?: {
    p2wTmpdir?: () => Promise<{ dir: string }>
    p2wExec?: (p: { mode: 'detect' | 'convert'; input?: string; output?: string; lang?: string; profile?: string; timeoutMs?: number }) =>
      Promise<{ ok: boolean; code: number; stdout: string; stderr: string; exe: string; bytes?: Uint8Array }>
    wfcWriteFile?: (filePath: string, bytes: Uint8Array) => Promise<{ ok: boolean; size: number }>
  }
}).api

const u8FromB64 = (b64: string): Uint8Array => {
  const bin = atob(b64)
  const u8 = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
  return u8
}
const ensurePdf = (name: string) => (/\.pdf$/i.test(name) ? name : name + '.pdf')
const stripPdf = (name: string) => name.replace(/\.pdf$/i, '')

/** p2wItems 按 p2wOrder（id 序）排列——拖拽排序与删除的单一真相源（复用 pdf 排序语义） */
export function orderedP2w (items: P2wItem[], order: string[] | undefined | null): P2wItem[] {
  return orderedPdfs(items as PdfItem[], order) as P2wItem[]
}

/** 词条导入（节点拖放 / 导入按钮共用）——校验与持久化单一真相源 */
export async function importP2wFiles (nodeId: string, files: FileList | null): Promise<{ added: number; skipped: number; total: number }> {
  if (!files?.length) return { added: 0, skipped: 0, total: 0 }
  const st = useBaseStore.getState()
  const cfg = (st.tables[nodeId]?.config ?? {}) as { p2wItems?: P2wItem[]; p2wOrder?: string[] }
  const raw = cfg.p2wItems ?? []
  const next = [...raw]
  let skipped = 0
  for (const f of Array.from(files)) {
    if (!/\.pdf$/i.test(f.name) || f.size > P2W_MAX_BYTES || next.length >= P2W_MAX) { skipped++; continue }
    const id = 'p2w-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
    await putBlob('p2witem-' + id, new Uint8Array(await f.arrayBuffer()))
    next.push({ id, name: f.name, inDb: true })
  }
  if (next.length === raw.length) return { added: 0, skipped, total: raw.length }
  st.pushHistory()
  st.setNodeConfig(nodeId, { p2wItems: next, p2wOrder: (cfg.p2wOrder ?? []).concat(next.slice(raw.length).map((it) => it.id)) })
  pushLog(`[PDF转Word] 导入：+${next.length - raw.length} 份（共 ${next.length}）${skipped ? `，跳过 ${skipped} 个非 pdf 或超 200MB` : ''}`, next.length > raw.length ? 'ok' : 'warn')
  return { added: next.length - raw.length, skipped, total: next.length }
}

/** 读一份词条字节（IDB 优先 b64 兜底） */
async function readItemBytes (it: P2wItem): Promise<Uint8Array | null> {
  if (it.b64) return u8FromB64(it.b64)
  return getBlob('p2witem-' + it.id)
}

/** 收集待转输入：词条优先；无词条时收线上游（语义同 fileout） */
export async function collectP2wInputs (nodeId: string): Promise<{ items: { name: string; bytes: Uint8Array }[]; source: 'items' | 'upstream' }> {
  const st = useBaseStore.getState()
  const cfg = (st.tables[nodeId]?.config ?? {}) as { p2wItems?: P2wItem[]; p2wOrder?: string[]; inputs?: string[] }
  const items = orderedP2w(cfg.p2wItems ?? [], cfg.p2wOrder ?? null)
  if (items.length) {
    const out: { name: string; bytes: Uint8Array }[] = []
    for (const it of items) {
      const bytes = await readItemBytes(it)
      if (bytes) out.push({ name: ensurePdf(it.name || it.id), bytes })
    }
    return { items: out, source: 'items' }
  }
  const out: { name: string; bytes: Uint8Array }[] = []
  for (const uid of cfg.inputs ?? []) {
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
      const pi = (uc.pdfItems ?? []) as { id: string; name: string; inDb?: boolean; b64?: string }[]
      const order = (uc.pdfFileOrder ?? null) as string[] | null
      const ordered = order ? [...pi].sort((a, b) => order.indexOf(a.id) - order.indexOf(b.id)) : pi
      for (const it of ordered) {
        const bytes = it.b64 ? u8FromB64(it.b64) : await getBlob('pdfitem-' + it.id)
        if (bytes) out.push({ name: ensurePdf(it.name || it.id), bytes })
      }
    }
  }
  return { items: out, source: 'upstream' }
}

/** 引擎检测：本机 python + paddleocr 可用即 true（结果不缓存——装完即时可用） */
export async function detectP2wEngine (): Promise<{ ok: boolean; msg: string }> {
  if (!api?.p2wExec) return { ok: false, msg: '桌面桥缺失（重启应用）' }
  const r = await api.p2wExec({ mode: 'detect' })
  if (r.ok && r.stdout.includes('PADDLEOCR_OK')) return { ok: true, msg: 'PaddleOCR 引擎就绪' }
  if (/ENOENT|not recognized|找不到/i.test(r.stderr)) return { ok: false, msg: '未找到 Python——先装 Python 3.9-3.12（勾选 Add to PATH），再跑 install_engine.bat' }
  return { ok: false, msg: '未安装 PaddleOCR 引擎——双击软件目录 pdf2word/install_engine.bat 一键安装' }
}

/* ★ 2026-10-26 套件「PDF转Word · 文字版直转」（x2t 长驻转换器，约 0.4s/份）：
   /docxconv 与合并链 /pdfconv 同管道，方向反转 outputtype='docx'。
   返回产物字节 + 服务端耗时 took（过程反馈用）。 */
async function convertViaX2t (name: string, bytes: Uint8Array): Promise<{ bytes: Uint8Array; took: number }> {
  const CHUNK = 0x8000
  let bin = ''
  for (let i = 0; i < bytes.length; i += CHUNK) bin += String.fromCharCode(...bytes.subarray(i, i + CHUNK))
  const r = await fetch('/docxconv', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: 'p2w-x2t-' + Date.now().toString(36) + '.pdf', dataBase64: btoa(bin) }),
  }).then((x) => x.json()).catch((e) => ({ ok: false, err: String(e) }))
  if (r && r.ok && r.dataBase64) {
    const parsed = atob(r.dataBase64)
    const arr = new Uint8Array(parsed.length)
    for (let i = 0; i < parsed.length; i++) arr[i] = parsed.charCodeAt(i)
    return { bytes: arr, took: Number(r.took) || 0 }
  }
  throw new Error((r && r.err) || 'x2t 转换失败')
}

/** PDF 页数（过程反馈展示用；解析失败返回 0 不阻塞转换） */
async function pageCountOf (bytes: Uint8Array): Promise<number> {
  try {
    const pdfjs = await import('pdfjs-dist')
    const doc = await pdfjs.getDocument({ data: bytes.slice() }).promise
    const n = doc.numPages
    doc.destroy()
    return n
  } catch { return 0 }
}

/** 执行转换（2026-10-26 用户指令「节点上的方形按钮 点击切换引擎」）：
 *  严格按节点选定引擎（config.p2wEngine，'x2t' 缺省）——不再自动探测路由。
 *  x2t=文字版直转（套件闸门）；ocr=百度 PaddleOCR（先检测套件内 venv 引擎）。 */
/* ★ 并发防抖（2026-10-26 用户截图「明显不正常」）：同一节点重复点击会排两条队，
   前一条超时后回退的一次性转换还在后台烧 CPU，徽章永远 running。 */
const p2wInFlight = new Set<string>()
/* ★ 兜底看门狗：IPC 理论上可能挂死——主链路超时后最多再宽限 120s，强制判败清徽章 */
const withWatchdog = async <T>(p: Promise<T>, ms: number, what: string): Promise<T> => {
  let to: ReturnType<typeof setTimeout> | undefined
  const bomb = new Promise<never>((_, rej) => {
    to = setTimeout(() => rej(new Error(what + ' 超时（' + Math.round(ms / 1000) + 's）——已放弃等待')), ms)
  })
  try {
    return await Promise.race([p, bomb])
  } finally {
    if (to) clearTimeout(to)
  }
}

async function runPdf2WordInner (nodeId: string): Promise<{ ok: boolean; msg: string }> {
  if (!api?.p2wExec || !api?.p2wTmpdir || !api?.wfcWriteFile) {
    setNodeRunStatus(nodeId, 'error', '桌面桥缺失（重启应用）')
    return { ok: false, msg: '桌面桥缺失' }
  }
  const st0 = useBaseStore.getState()
  const cfg0 = (st0.tables[nodeId]?.config ?? {}) as { p2wEngine?: 'x2t' | 'ocr' }
  const engine: 'x2t' | 'ocr' = cfg0.p2wEngine === 'ocr' ? 'ocr' : 'x2t'
  setNodeRunStatus(nodeId, 'running', '收集输入…')
  const { items, source } = await collectP2wInputs(nodeId)
  if (!items.length) {
    const msg = '无输入——导入 PDF 或连线上游'
    setNodeRunStatus(nodeId, 'error', msg)
    return { ok: false, msg }
  }
  /* 套件闸门：选定引擎的套件必须可用 */
  if (engine === 'x2t' && !hasCapability('convert:pdf2docx')) {
    const msg = '文字版直转已停用——套件中心启用或切换 OCR 引擎'
    setNodeRunStatus(nodeId, 'error', msg)
    pushLog('[PDF转Word] ' + msg, 'err')
    return { ok: false, msg }
  }
  let ocrOk = false
  if (engine === 'ocr') {
    setNodeRunStatus(nodeId, 'running', '检测 OCR 引擎…')
    ocrOk = (await detectP2wEngine()).ok
    if (!ocrOk) {
      const msg = 'OCR 引擎未安装——套件中心「装入套件目录」'
      setNodeRunStatus(nodeId, 'error', msg)
      pushLog('[PDF转Word] ' + msg + '（套件中心 → 扫描件 OCR → 装入套件目录；前置：本机有 Python 3.9-3.12）', 'err')
      return { ok: false, msg }
    }
  }
  const t0 = Date.now()
  const { dir } = engine === 'ocr' ? await api.p2wTmpdir() : { dir: '' }
  const st = useBaseStore.getState()
  const cfg = (st.tables[nodeId]?.config ?? {}) as { p2wLang?: 'ch' | 'en'; p2wHistory?: P2wHistRec[]; p2wProfile?: 'fast' | 'std' | 'pro' }
  const lang = cfg.p2wLang === 'en' ? 'en' : 'ch'
  /* 速度档位（2026-10-26 用户问「有没有容差设置」）：pro 精准（缺省）/ std 均衡 / fast 极速 */
  const profile: 'fast' | 'std' | 'pro' = (['fast', 'std', 'pro'] as const).includes(cfg.p2wProfile as 'fast') ? (cfg.p2wProfile as 'fast' | 'std' | 'pro') : 'pro'
  const profileLabel = { pro: '精准', std: '均衡', fast: '极速' }[profile]
  const hist: P2wHistRec[] = [...(cfg.p2wHistory ?? [])]
  let done = 0
  let fail = 0
  for (let i = 0; i < items.length; i++) {
    const it = items[i]
    /* ★ 2026-10-26 用户指令「转换过程 稍微详细一点」：页数 + 阶段 + 逐份耗时全上报 */
    const pages = await pageCountOf(it.bytes)
    setNodeRunStatus(nodeId, 'running', engine === 'ocr'
      ? `OCR 识别中 ${i + 1}/${items.length} · ${it.name}${pages ? ` · ${pages} 页` : ''}`
      : `直转中 ${i + 1}/${items.length} · ${it.name}${pages ? ` · ${pages} 页` : ''}`)
    const histId = 'p2w-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
    const fT0 = Date.now()
    try {
      let bytes: Uint8Array | null = null
      if (engine === 'x2t') {
        const r = await convertViaX2t(it.name, it.bytes)
        bytes = r.bytes
        pushLog(`[PDF转Word] ✓ ${it.name} → docx（${pages} 页 · 引擎转换 ${(r.took / 1000).toFixed(1)}s · 合计 ${((Date.now() - fT0) / 1000).toFixed(1)}s）`, 'ok')
      } else {
        setNodeRunStatus(nodeId, 'running', `OCR 识别中 ${i + 1}/${items.length} · ${it.name}${pages ? ` · ${pages} 页` : ''} · ${profileLabel}档（识别耗时与页数相关）`)
        const inPath = dir + '/in-' + i + '.pdf'
        const outPath = dir + '/out-' + i + '.docx'
        const wr = await api.wfcWriteFile!(inPath, it.bytes)
        /* ★ 2026-10-26 二修：中转文件必须核对字节数——写丢/写空时 python 侧只会回
           「输入文件不存在」，用户侧看到的是「python 退出码 -1」，真因全丢 */
        if (!wr || wr.ok === false) throw new Error('写入中转 PDF 失败（' + inPath + '）')
        if (typeof wr.size === 'number' && wr.size !== it.bytes.length) {
          throw new Error(`写入中转 PDF 不完整（${wr.size}/${it.bytes.length} 字节）`)
        }
        /* ★ 超时按页数动态算（52 页 900s 固定超时太紧）：120s 底 + 30s×页 + 60s 余量，下限 5 分钟 */
        const timeoutMs = Math.max(300000, 120000 + (pages + 2) * 30000)
        const r = await withWatchdog(
          api.p2wExec!({ mode: 'convert', input: inPath, output: outPath, lang, profile, timeoutMs }),
          timeoutMs + 120000,
          'OCR 执行',)
        if (r.ok && r.bytes?.length) bytes = r.bytes
        else {
          const lines = (r.stderr || '').split('\n').map((l) => l.trim()).filter(Boolean)
          const errLine = lines.find((l) => l.includes('P2W_ERROR'))
            || lines[lines.length - 1]
            || ('python 退出码 ' + r.code)
          throw new Error(errLine.replace('P2W_ERROR: ', '').slice(0, 240))
        }
        pushLog(`[PDF转Word] ✓ ${it.name} → docx（${pages} 页 · ${profileLabel}档 OCR ${((Date.now() - fT0) / 1000).toFixed(1)}s）`, 'ok')
      }
      if (bytes?.length) {
        await putBlob('p2whist-' + histId, bytes)
        hist.unshift({ id: histId, name: stripPdf(it.name) + '.docx', at: Date.now(), inDb: true, source, engine })
        done++
      } else throw new Error('产物为空')
    } catch (e) {
      fail++
      pushLog(`[PDF转Word] ✗ ${it.name} 失败（${((Date.now() - fT0) / 1000).toFixed(1)}s）：${e instanceof Error ? e.message : String(e)}`, 'err')
    }
  }
  if (done > 0) {
    st.pushHistory()
    st.setNodeConfig(nodeId, { p2wHistory: hist.slice(0, 10) })
    const sec = ((Date.now() - t0) / 1000).toFixed(1)
    setNodeRunStatusAuto(nodeId, 'ok', `${done} 份 Word · ${sec}s${fail ? `（失败 ${fail}）` : ''}`)
    pushLog(`[PDF转Word] 完成：成功 ${done} / 失败 ${fail}（${sec}s，来源=${source === 'items' ? '词条' : '上游'}，引擎=${engine === 'x2t' ? '文字版直转' : '百度OCR'}）`, fail ? 'warn' : 'ok')
    return { ok: true, msg: `成功 ${done} 份` }
  }
  const msg = fail ? '转换失败（详见日志）' : '无输入'
  setNodeRunStatus(nodeId, 'error', msg.slice(0, 40))
  return { ok: false, msg }
}

/** 对外入口：并发防抖包装（2026-10-26 用户截图「明显不正常」——重复点击把队列搞乱、
 *  徽章永远 running。同一节点同时只允许一次转换在飞。） */
export async function runPdf2Word (nodeId: string): Promise<{ ok: boolean; msg: string }> {
  if (p2wInFlight.has(nodeId)) {
    const msg = '已有一次转换在进行——等它结束（后台完成后历史里会出现），或重启应用终止它'
    pushLog('[PDF转Word] ' + msg + '（重复点击已拦截）', 'warn')
    return { ok: false, msg }
  }
  p2wInFlight.add(nodeId)
  try {
    return await runPdf2WordInner(nodeId)
  } catch (e) {
    /* 任何未捕获异常也要清徽章——running 永远不该悬挂 */
    setNodeRunStatus(nodeId, 'error', ('转换异常：' + (e instanceof Error ? e.message : String(e))).slice(0, 40))
    pushLog(`[PDF转Word] ✗ 未捕获异常：${e instanceof Error ? e.message : String(e)}`, 'err')
    return { ok: false, msg: '转换异常' }
  } finally {
    p2wInFlight.delete(nodeId)
  }
}

/** 删除一条历史记录（字节不删——存储铁律：孤儿数据无害） */
export function removeP2wHist (nodeId: string, id: string) {
  const st = useBaseStore.getState()
  const cfg = (st.tables[nodeId]?.config ?? {}) as { p2wHistory?: P2wHistRec[] }
  const victim = (cfg.p2wHistory ?? []).find((r) => r.id === id)
  st.pushHistory()
  st.setNodeConfig(nodeId, { p2wHistory: (cfg.p2wHistory ?? []).filter((r) => r.id !== id) })
  if (victim) pushLog(`[PDF转Word] 删除记录：${victim.name}`)
}
