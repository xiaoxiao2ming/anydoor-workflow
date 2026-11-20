/**
 * PDF 替换页核心（2026-10-25 用户指令「增加一个PDF替换页的节点」）。
 *
 * 语义：上组（被替换）与下组（替换页）按行配对——
 *   上组第 i 份 PDF 的第 Xi 页 ← 下组第 i 份 PDF 的第 Yi 页。
 * 产物 = 替换后的完整 PDF → unshift 进 replaceHistory（最新在前）。
 * 字节优先 IDB（inDb），b64 兜底（与 pdfHistory 同款双来源）。
 */
import { PDFDocument } from 'pdf-lib'
import { getBlob, putBlob } from './pdfBlobStore'
import { useBaseStore } from './store'

export interface ReplaceEntry { id: string; name: string; b64?: string; inDb?: boolean; page: number }
export interface ReplaceHistEntry { id: string; name: string; at: number; b64?: string; inDb?: boolean; pages?: number[] }

/** b64 → bytes（fetch 原生解码，与 store.ts 同款） */
async function b64ToU8 (b64: string): Promise<Uint8Array> {
  const blob = await (await fetch('data:application/octet-stream;base64,' + b64)).blob()
  return new Uint8Array(await blob.arrayBuffer())
}

/** 双来源读字节：inDb 优先，b64 兜底 */
export async function readReplaceBytes (e: { id: string; b64?: string; inDb?: boolean }): Promise<Uint8Array | null> {
  if (e.inDb) {
    const b = await getBlob(e.id)
    if (b) return b
  }
  if (e.b64) return b64ToU8(e.b64)
  if (!e.inDb) {
    const b = await getBlob(e.id)
    if (b) return b
  }
  return null
}

/**
 * 执行替换：按计划（pairs）逐对执行——每对产出一条历史。
 * pairs 由 analyzePdfReplace 产出（或用户手改页码后的同结构）。
 * 产物名：{原名}_替换页{n}.pdf
 */
export interface ReplacePair { tgtId: string; tgtName: string; tgtPage: number; srcId: string; srcName: string; srcPage: number; sim: number; lowSim?: boolean }

export async function runPdfReplace (nodeId: string, pairs?: ReplacePair[]): Promise<{ ok: boolean; msg: string }> {
  const st = useBaseStore.getState()
  const t = st.tables[nodeId]
  if (!t) return { ok: false, msg: '节点不存在' }
  const cfg = (t.config ?? {}) as {
    replaceTargets?: ReplaceEntry[]
    replaceSources?: ReplaceEntry[]
    replaceHistory?: ReplaceHistEntry[]
  }
  const targets = cfg.replaceTargets ?? []
  const sources = cfg.replaceSources ?? []
  /* 未传计划 → 兜底按当前页码行配对（旧行为） */
  const plan: ReplacePair[] = pairs ?? targets.slice(0, Math.min(targets.length, sources.length)).map((tg, i) => ({
    tgtId: tg.id, tgtName: tg.name, tgtPage: tg.page, srcId: sources[i].id, srcName: sources[i].name, srcPage: sources[i].page, sim: -1,
  }))
  if (!plan.length) return { ok: false, msg: '无替换计划——先匹配或放入文件' }

  const hist: ReplaceHistEntry[] = [...(cfg.replaceHistory ?? [])]
  const t0 = Date.now()
  let done = 0
  const errs: string[] = []

  for (const p of plan) {
    const tg = targets.find((x) => x.id === p.tgtId)
    const sc = sources.find((x) => x.id === p.srcId)
    if (!tg || !sc) { errs.push(`${p.tgtName}：计划项文件已不在组内`); continue }
    try {
      const tb = await readReplaceBytes(tg)
      const sb = await readReplaceBytes(sc)
      if (!tb) { errs.push(`${tg.name}：字节缺失`); continue }
      if (!sb) { errs.push(`${sc.name}：字节缺失`); continue }
      const tp = Math.max(1, Math.floor(p.tgtPage || 1))
      const sp = Math.max(1, Math.floor(p.srcPage || 1))

      const doc = await PDFDocument.load(tb, { ignoreEncryption: true })
      if (tp > doc.getPageCount()) { errs.push(`${tg.name}：页码 ${tp} 超出总页数 ${doc.getPageCount()}`); continue }
      const srcDoc = await PDFDocument.load(sb, { ignoreEncryption: true })
      if (sp > srcDoc.getPageCount()) { errs.push(`${sc.name}：页码 ${sp} 超出总页数 ${srcDoc.getPageCount()}`); continue }

      const [copied] = await doc.copyPages(srcDoc, [sp - 1])
      doc.removePage(tp - 1)
      doc.insertPage(tp - 1, copied)

      const outBytes = await doc.save()
      const hid = 'rephist-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 6)
      const baseName = tg.name.replace(/\.pdf$/i, '')
      const outName = `${baseName}_替换页${tp}.pdf`
      await putBlob(hid, outBytes)
      /* ★ pages=本次执行的全部替换点位——修改前后对比视图/放大浮层的定位切换用
         （考虑多页替换：任一条目打开都能看到整批点位） */
      hist.unshift({ id: hid, name: outName, at: Date.now(), inDb: true, pages: plan.map((p) => Math.max(1, Math.floor(p.tgtPage || 1))) })
      done++
    } catch (e) {
      errs.push(`${tg.name} ← ${sc.name}：${String((e as Error)?.message ?? e).slice(0, 60)}`)
    }
  }

  /* 历史上限 10 条（与 pdfmerge 一致） */
  const trimmed = hist.slice(0, 10)
  st.pushHistory()
  st.setNodeConfig(nodeId, { replaceHistory: trimmed })

  if (!done) return { ok: false, msg: errs.join('；') || '无可执行的替换对' }
  const dur = ((Date.now() - t0) / 1000).toFixed(1)
  const tail = errs.length ? `；跳过：${errs.join('；')}` : ''
  return { ok: true, msg: `${done} 份替换完成，耗时 ${dur}s${tail}` }
}

/* ===== 匹配分析（2026-10-25 用户指令「匹配到哪个文件那一页 相似度多少百分比」）===== */

/** 逐页提取文本 item 序列（pdfjs-dist 动态加载）——序列级相似度/差异定位共用 */
async function extractPageItems (bytes: Uint8Array): Promise<string[][]> {
  const pdfjs = await import('pdfjs-dist')
  const out: string[][] = []
  const doc = await pdfjs.getDocument({ data: bytes.slice() }).promise
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p)
    const tc = await page.getTextContent()
    /* ★ 2026-10-25 用户实锤「红框标记位置不对」根因：pdfjs 的 item 顺序≠视觉顺序
       （不同生成器乱序）——LCS 对乱序序列产生大量伪差异。
       修正：按坐标聚「行」（y 聚类 + 行内 x 排序）——后续 diff/相似度/红框全部基于行 */
    const items = (tc.items as { str?: string; transform?: number[]; width?: number; height?: number }[])
      .map((it) => ({
        str: (it.str ?? '').replace(/\s+/g, ' ').trim(),
        x: it.transform?.[4] ?? 0,
        y: it.transform?.[5] ?? 0,
        h: Math.abs(it.height || it.transform?.[3] || 10),
      }))
      .filter((it) => it.str)
    items.sort((a, b) => (b.y - a.y) || (a.x - b.x)) /* pdf 坐标 y 向上：行从上到下= y 降序 */
    const rows: string[] = []
    let curY = Number.NaN
    let buf = ''
    const flush = () => { if (buf) { rows.push(buf); buf = '' } }
    for (const it of items) {
      /* 行聚类容差 = 字高×0.6（同行基线波动） */
      if (Number.isNaN(curY) || Math.abs(it.y - curY) > it.h * 0.6 + 1) { flush(); curY = it.y }
      buf += (buf ? ' ' : '') + it.str
    }
    flush()
    out.push(rows)
  }
  await doc.destroy()
  return out
}

/** 行级相似度（0-100）：行序列 LCS diff → 1 - 差异行字符占比。
 *  ★ v4（2026-10-25）：v3 item 级 LCS 仍受 item 顺序干扰（pdfjs item 乱序时伪差异）
 *    ——输入改为「视觉行」序列（坐标聚类），diff 段=被修改的行，与红框完全同源。 */
function seqSim (a: string[], b: string[]): number {
  if (!a.length && !b.length) return -1
  if (!a.length || !b.length) return 0
  const segs = seqDiffSegments(a, b)
  let diff = 0
  let total = 0
  for (const t of a) total += t.length
  for (const t of b) total += t.length
  for (const sg of segs) {
    for (let k = sg.aS; k <= sg.aE; k++) diff += a[k].length
    for (let k = sg.bS; k <= sg.bE; k++) diff += b[k].length
  }
  if (!total) return -1
  return Math.max(0, Math.min(100, Math.round(100 * (1 - diff / total))))
}

/** 序列 LCS 差异段（供放大视图红框标注）——返回两侧区间（-1 表示该侧无内容） */
export function seqDiffSegments (a: string[], b: string[]): { aS: number; aE: number; bS: number; bE: number }[] {
  const n = a.length
  const m = b.length
  /* LCS 长度矩阵（item 数几百级，O(n·m) 可承受） */
  const dp: Uint32Array[] = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1))
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1])
    }
  }
  const segs: { aS: number; aE: number; bS: number; bE: number }[] = []
  let i = 0
  let j = 0
  while (i < n && j < m) {
    if (a[i] === b[j]) { i++; j++; continue }
    const aS = i
    const bS = j
    while (i < n && j < m && a[i] !== b[j]) {
      if (dp[i + 1][j] >= dp[i][j + 1]) i++
      else j++
    }
    segs.push({ aS, aE: i - 1, bS, bE: j - 1 })
  }
  if (i < n) segs.push({ aS: i, aE: n - 1, bS: j, bE: j - 1 })
  if (j < m) segs.push({ aS: i, aE: i - 1, bS: j, bE: m - 1 })
  return segs
}

/**
 * 匹配分析：对每个替换来源页，在上组所有文件的每一页中找文本最相似页。
 * 返回计划（含双向页码+相似度）——调用方回填页码输入框+右侧展示+确认执行。
 */
export async function analyzePdfReplace (nodeId: string, tolerance = 0): Promise<{ ok: boolean; pairs: ReplacePair[]; skipped: string[]; msg: string }> {
  const st = useBaseStore.getState()
  const t = st.tables[nodeId]
  if (!t) return { ok: false, pairs: [], skipped: [], msg: '节点不存在' }
  const cfg = (t.config ?? {}) as { replaceTargets?: ReplaceEntry[]; replaceSources?: ReplaceEntry[] }
  const targets = cfg.replaceTargets ?? []
  const sources = cfg.replaceSources ?? []
  if (!targets.length) return { ok: false, pairs: [], skipped: [], msg: '上面组为空——先放入被替换的 PDF' }
  if (!sources.length) return { ok: false, pairs: [], skipped: [], msg: '下面组为空——先放入替换页来源 PDF' }

  /* 预取上组全部页文本（文件多时逐份惰性——v1 全量，图纸规模可承受） */
  const tgtTexts: { entry: ReplaceEntry; texts: string[][] | null }[] = []
  for (const tg of targets) {
    try {
      const b = await readReplaceBytes(tg)
      tgtTexts.push({ entry: tg, texts: b ? await extractPageItems(b) : null })
    } catch { tgtTexts.push({ entry: tg, texts: null }) }
  }

  const pairs: ReplacePair[] = []
  const unanalysable: string[] = []
  for (const sc of sources) {
    try {
      const sb = await readReplaceBytes(sc)
      if (!sb) { unanalysable.push(`${sc.name}：字节缺失`); continue }
      const srcTexts = await extractPageItems(sb)
      for (let sp = 1; sp <= srcTexts.length; sp++) {
        let best: { tg: ReplaceEntry; tp: number; sim: number } | null = null
        for (const { entry, texts } of tgtTexts) {
          if (!texts) continue
          for (let tp = 1; tp <= texts.length; tp++) {
            const sim = seqSim(srcTexts[sp - 1], texts[tp - 1])
            if (sim < 0) continue /* 双方均无文本 */
            if (!best || sim > best.sim) best = { tg: entry, tp, sim }
          }
        }
        if (best) {
          /* ★ 2026-10-25 语义修正（用户令「自从调整后就少了一页」）：容差从「删除线」改「警告线」——
             低于容差的配对**照常进列表**（lowSim 标记，弹窗红徽章），由用户自行判断，
             不再静默砍掉（此前 3 页变 2 页的根因）。真正无法分析的（无文本/缺字节）才跳过 */
          pairs.push({ tgtId: best.tg.id, tgtName: best.tg.name, tgtPage: best.tp, srcId: sc.id, srcName: sc.name, srcPage: sp, sim: best.sim, lowSim: best.sim < tolerance })
        } else {
          unanalysable.push(`${sc.name} 第${sp}页：无文本（扫描件？）——请手动指定页码`)
        }
      }
    } catch (e) {
      unanalysable.push(`${sc.name}：${String((e as Error)?.message ?? e).slice(0, 50)}`)
    }
  }

  if (!pairs.length) return { ok: false, pairs: [], skipped: unanalysable, msg: '未找到可匹配页——' + unanalysable.slice(0, 2).join('；') }
  const tail = unanalysable.length ? `；${unanalysable.length} 项未进列表` : ''
  return { ok: true, pairs, skipped: unanalysable, msg: `${pairs.length} 对匹配完成${tail}` }
}
