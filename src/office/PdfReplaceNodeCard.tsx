/**
 * PDF 替换页节点卡片（2026-10-25 用户指令「增加一个PDF替换页的节点」）。
 *
 * 结构（15:24 用户三条反馈重构）：
 *   PdfReplaceNodeCard —— 上下两组词条（被替换/替换页，页码可改），
 *                          词条边距对齐统一标准（左右 6px/上下 4px）
 *   PdfReplaceRunButton —— 「换」31×31 方形钮，渲染在 Canvas 底部按钮行
 *                          （与 GrpLockButton 并排——用户指令「放在最下面和锁并排」）
 *   PdfReplaceHistList —— 🕘 历史（N）点击展开/再点收起（仿 PDF 合并历史），
 *                          历史词条行：双击 PDF.js 预览 / 右键菜单（打开三件套+
 *                          另存为+删除）/ ⇄ 与原文件视觉 diff（diff-pdf）
 * 执行两段式：匹配分析（pdfjs 文本余弦）→ confirm 列出配对详情 → 确认执行
 * （用户指令「替换的时候页面不要在右侧有变化」——匹配结果只在确认框，不渲染侧列）
 */
import { useState, useEffect, useRef, useSyncExternalStore } from 'react'
import { createPortal } from 'react-dom'
import { Handle, Position, useUpdateNodeInternals } from '@xyflow/react'
import { useBaseStore } from '../core/store'
import { pushLog } from '../ui/logBus'
import { setNodeRunStatusAuto } from '../core/nodeRunStatus'
import { runPdfReplace, analyzePdfReplace, readReplaceBytes, seqDiffSegments, type ReplaceEntry, type ReplaceHistEntry, type ReplacePair } from '../core/pdfReplace'
import { PdfViewer } from './PdfViewer'
import { IcFilter, IcGitCompare, IcHistory, IcMaximize, IcMinimize, IcX } from '../ui/icons'

/* ★ 历史展开状态（模块级外部 store——同 pdfmerge pdfHistUI 模式：
   方形钮在按钮行、列表在按钮行后，跨组件共享） */
let histOpen = false
const histSubs = new Set<() => void>()
const prHistUI = {
  get: () => histOpen,
  toggle: () => { histOpen = !histOpen; histSubs.forEach((f) => f()) },
  sub: (f: () => void) => { histSubs.add(f); return () => { histSubs.delete(f) } },
}

const fmtTime = (at: number) => { const d = new Date(at); return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}` }

/* ═══════════ 执行按钮（Canvas 底部按钮行——与锁并排） ═══════════ */

export function PdfReplaceRunButton ({ nodeId }: { nodeId: string }) {
  const [busy, setBusy] = useState(false)
  /* ★ 统一数据锁（2026-10-26）：锁定时禁执行（执行=追加替换历史=内容变更） */
  const lcfg = useBaseStore((s) => s.tables[nodeId]?.config) as { dataLocked?: boolean; ooDataLocked?: boolean } | undefined
  const locked = !!(lcfg?.dataLocked || lcfg?.ooDataLocked)
  /* ★ 2026-10-25 用户指令「换的按钮左侧增加容差设置」+「点击方形按钮 下方有个左右滑动的」：
     方钮点击 → 下方展开滑条面板。
     ★★ 面板 portal 到 body（fixed 定位）——absolute 挂节点内会被 rf-node-inner
     overflow:hidden 裁掉（点击「没反应」=面板开了但在卡片外被裁，锚点教训同款） */
  const tol = useBaseStore((s) => ((s.tables[nodeId]?.config ?? {}) as { replaceTolerance?: number }).replaceTolerance ?? 60)
  const setTol = (v: number) => {
    const clamped = Math.max(0, Math.min(100, Math.floor(v || 0)))
    useBaseStore.getState().setNodeConfig(nodeId, { replaceTolerance: clamped })
  }
  const tolBtnRef = useRef<HTMLButtonElement>(null)
  const [tolOpen, setTolOpen] = useState(false)
  /* ★ 修「第二次点不出来」：遮罩关闭只清 tolOpen 不清 tolPos——旧 toggle 把「有坐标」
     误判为「已打开」而清坐标，面板条件 (tolOpen && tolPos) 永不成立。
     正解：tolOpen 唯一开关；坐标每次点击都刷新（顺带修拖动画布后位置漂移） */
  const toggleTol = () => {
    const r = tolBtnRef.current?.getBoundingClientRect()
    if (r) setTolPos({ x: r.left, y: r.bottom + 4 })
    setTolOpen((v) => !v)
  }
  const [tolPos, setTolPos] = useState<{ x: number; y: number } | null>(null)
  /* window.confirm 纯文本 → 结构化表格弹窗（被替换页|相似度|匹配来源|缩略图|被替换文件名） */
  const [confirmPairs, setConfirmPairs] = useState<ReplacePair[] | null>(null)
  const [confirmSkip, setConfirmSkip] = useState<string[]>([])

  const run = async () => {
    if (busy || locked) return
    setBusy(true)
    try {
      /* ① 匹配分析：来源页在上组找最相似页（容差过滤低质量配对） */
      const an = await analyzePdfReplace(nodeId, tol)
      if (!an.ok) {
        setNodeRunStatusAuto(nodeId, 'error', an.msg.slice(0, 40))
        pushLog('[PDF·替换] ' + an.msg, 'err')
        return
      }
      /* ② 结构化确认弹窗（表格 + 缩略图对比）+ 跳过详情可见（用户令「替换3页只出2对」要能看到缺的那页去哪了） */
      setConfirmPairs(an.pairs)
      setConfirmSkip(an.skipped)
    } finally {
      setBusy(false)
    }
  }

  const doExecute = async (pairs: ReplacePair[]) => {
    setConfirmPairs(null)
    setBusy(true)
    try {
      const r = await runPdfReplace(nodeId, pairs)
      setNodeRunStatusAuto(nodeId, r.ok ? 'ok' : 'error', r.msg.slice(0, 40))
      pushLog('[PDF·替换] ' + r.msg, r.ok ? 'ok' : 'err')
    } finally {
      setBusy(false)
    }
  }

  return (
    <>
      {/* ★ 容差设置（换钮左侧）：方钮点击 → body 级滑条面板（fixed 贴按钮下缘，防节点裁切） */}
        <button
          ref={tolBtnRef}
          className={'pm-sq-btn pm-sq-btn--ghost nodrag nopan' + (tolOpen ? ' pm-sq-btn--on' : '')}
          title="匹配容差：低于该相似度%的配对将标红警告（不剔除，由你确认）"
          onClick={(e) => { e.stopPropagation(); toggleTol() }}
        >
          <IcFilter size={17} />
        </button>
      {tolOpen && tolPos && createPortal(
        <div
          className="libpick-mask nodrag nopan"
          style={{ background: 'transparent', zIndex: 1000 }}
          onClick={() => setTolOpen(false)}
          onContextMenu={(e) => { e.preventDefault(); setTolOpen(false) }}
        >
          <div
            className="pr-tol-panel"
            style={{ position: 'fixed', left: tolPos.x, top: tolPos.y, margin: 0 }}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="pr-tol-row">
              <span>匹配容差</span>
              <b className={'pr-sim ' + (tol >= 80 ? 'pr-sim--hi' : tol >= 50 ? 'pr-sim--mid' : 'pr-sim--low')}>{tol}%</b>
            </div>
            <input
              type="range"
              min={0}
              max={100}
              step={5}
              value={tol}
              onChange={(e) => setTol(parseInt(e.target.value, 10))}
            />
            <div className="pr-tol-hint"><span>0 · 宽松</span><span>低于容差标红警告</span><span>严格 · 100</span></div>
          </div>
        </div>,
        document.body,
      )}
      <button
        className="pm-sq-btn nodrag nopan"
        title={locked ? '数据已锁定：先解锁再执行替换' : '替换（自动匹配相似页 → 确认 → 执行）'}
        disabled={busy || locked}
        onClick={() => void run()}
      >
        {busy ? <span className="rn-spin pm-spin-dark" /> : <span className="pm-sq-txt">换</span>}
      </button>
      {confirmPairs && createPortal(
        <PdfReplaceConfirm
          nodeId={nodeId}
          pairs={confirmPairs}
          skipped={confirmSkip}
          onCancel={() => { setConfirmPairs(null); pushLog('[PDF·替换] 已取消执行', 'warn') }}
          onConfirm={(pairs) => void doExecute(pairs)}
        />,
        document.body,
      )}
    </>
  )
}

/* ═══════════ 匹配确认弹窗（表格化 + 差异缩略图） ═══════════ */

/** 单页缩略渲染（pdfjs → canvas）。targetW 给定=按宽度适配（对比列用）；否则高 150px */
async function renderPageThumb (bytes: Uint8Array, page: number, canvas: HTMLCanvasElement, targetW?: number): Promise<void> {
  const pdfjs = await import('pdfjs-dist')
  const doc = await pdfjs.getDocument({ data: bytes.slice() }).promise
  const pg = await doc.getPage(Math.max(1, Math.min(page, doc.numPages)))
  const vp1 = pg.getViewport({ scale: 1 })
  const scale = targetW ? targetW / vp1.width : 150 / vp1.height
  const vp = pg.getViewport({ scale })
  canvas.width = Math.floor(vp.width)
  canvas.height = Math.floor(vp.height)
  await pg.render({ canvasContext: canvas.getContext('2d')!, viewport: vp }).promise
}

/** 一对匹配页的并排预览（内嵌确认弹窗下方——用户令「预览在页面下方左右分裂」） */
function ThumbPair ({ nodeId, pair }: { nodeId: string; pair: ReplacePair }) {
  const cT = useRef<HTMLCanvasElement>(null)
  const cS = useRef<HTMLCanvasElement>(null)
  const wrapRef = useRef<HTMLDivElement>(null)
  const [err, setErr] = useState('')
  const [loading, setLoading] = useState(true)
  useEffect(() => {
    let alive = true
    void (async () => {
      try {
        const cfg = ((useBaseStore.getState().tables[nodeId]?.config ?? {}) as { replaceTargets?: ReplaceEntry[]; replaceSources?: ReplaceEntry[] })
        const tg = (cfg.replaceTargets ?? []).find((x) => x.id === pair.tgtId)
        const sc = (cfg.replaceSources ?? []).find((x) => x.id === pair.srcId)
        const tb = tg ? await readReplaceBytes(tg) : null
        const sb = sc ? await readReplaceBytes(sc) : null
        if (!alive) return
        if (!tb || !sb) { setErr('文件数据缺失，无法渲染缩略图'); setLoading(false); return }
        /* ★ 按面板半宽适配（用户令「预览在页面下方左右分裂」+ 窗体加大） */
        const half = Math.max(300, Math.floor((wrapRef.current?.clientWidth ?? 1100) / 2) - 30)
        if (cT.current) await renderPageThumb(tb, pair.tgtPage, cT.current, half)
        if (alive && cS.current) await renderPageThumb(sb, pair.srcPage, cS.current, half)
        if (alive) setLoading(false)
      } catch (e) { if (alive) { setErr(String((e as Error)?.message ?? e).slice(0, 80)); setLoading(false) } }
    })()
    return () => { alive = false }
  }, [nodeId, pair])
  return (
    <div className="pr-ip-body" ref={wrapRef}>
      {err ? <div className="pr-thumbs-err">{err}</div> : (
        <>
          <div className="pr-thumb-col">
            <div className="pr-thumb-cap">被替换页（原）· 第 {pair.tgtPage} 页</div>
            {loading && <div className="pr-thumb-loading" />}
            <canvas ref={cT} style={{ display: loading ? 'none' : undefined }} />
          </div>
          <div className="pr-thumb-vs">vs</div>
          <div className="pr-thumb-col">
            <div className="pr-thumb-cap">替换进来的页 · {pair.srcName} · 第 {pair.srcPage} 页</div>
            {loading && <div className="pr-thumb-loading" />}
            <canvas ref={cS} style={{ display: loading ? 'none' : undefined }} />
          </div>
        </>
      )}
    </div>
  )
}

/* ═══════════ 修改前后整档对比（左右两列 + 同步定位 + 替换点位跳转）═══════════ */

/** 对比列里的单页槽（IntersectionObserver 懒渲染——整档页数多时只有可见页才渲染） */
function CmpPage ({ bytes, page, cur, isPoint, onPick, onZoom }: {
  bytes: Uint8Array
  page: number
  cur: number
  isPoint: boolean
  onPick: (p: number) => void
  onZoom: (p: number) => void
}) {
  const holder = useRef<HTMLDivElement>(null)
  const cRef = useRef<HTMLCanvasElement>(null)
  const [done, setDone] = useState(false)
  useEffect(() => {
    const h = holder.current
    if (!h || done) return
    const io = new IntersectionObserver((ents) => {
      if (ents.some((e) => e.isIntersecting)) {
        io.disconnect()
        void (async () => {
          try {
            if (cRef.current) {
              /* ★ 按列宽适配窗体：测列容器（.pr-cmp-col）宽度——
                 holder 自身宽度随内容塌缩（canvas 未渲染时≈0），测它会得 220px 兜底值 */
              const col = h.closest('.pr-cmp-col')
              const colW = Math.max(220, (col?.clientWidth ?? h.clientWidth) - 30)
              await renderPageThumb(bytes, page, cRef.current, colW)
            }
          } catch { /* 单页失败不阻塞 */ }
          setDone(true)
        })()
      }
    }, { rootMargin: '400px' })
    io.observe(h)
    return () => io.disconnect()
  }, [bytes, page, done])
  return (
    <div
      ref={holder}
      data-page={page}
      className={'pr-cmp-page' + (cur === page ? ' pr-cmp-page--cur' : '') + (isPoint ? ' pr-cmp-page--point' : '')}
      title={`第 ${page} 页${isPoint ? '（替换点）' : ''}——点击两列同时定位`}
      onClick={() => onPick(page)}
    >
      <span className="pr-cmp-pp">
        第 {page} 页{isPoint ? ' · 替换点' : ''}
        {/* ★ 2026-10-25 用户令「上面有个按钮 点击页面具体修改的地方进行适配放大 定位在修改的地方」：
            替换点页角挂放大钮——打开高清差异浮层（红框标注+自动定位修改处） */}
        {isPoint && (
          <button
            className="pr-zoom-btn nodrag nopan"
            title="放大对比——定位到具体修改的地方"
            onClick={(e) => { e.stopPropagation(); onZoom(page) }}
          >
            <IcMaximize size={11} /> 放大
          </button>
        )}
      </span>
      {!done && <div className="pr-thumb-loading" />}
      <canvas ref={cRef} style={{ display: done ? undefined : 'none' }} />
    </div>
  )
}

/** 修改前后对比视图：左=原文件 右=修改后产物；点位条跳转；滚动同步 */
function PdfCompareView ({ leftName, leftBytes, rightName, rightBytes, points, onClose }: {
  leftName: string
  leftBytes: Uint8Array
  rightName: string
  rightBytes: Uint8Array
  points: number[]
  onClose: () => void
}) {
  const [lPages, setLPages] = useState(0)
  const [rPages, setRPages] = useState(0)
  /* ★ 2026-10-25 用户报「页面对照照不到」：此前加载失败被空 catch 吞=静默空白——
     失败原因显式展示（字节缺失/损坏一目了然） */
  const [loadErr, setLoadErr] = useState('')
  const [cur, setCur] = useState(points[0] ?? 1)
  /* ★ 2026-10-25 用户令「不要被动跟随 是同步进行」：滚动同步恢复默认开——
     同步算法改页对页对齐（比例映射会漂移=被动跟随的毛病） */
  const [sync, setSync] = useState(true)
  /* ★ 2026-10-25 用户令「点击页面具体修改的地方进行适配放大 定位在修改的地方」：
     替换点页角「放大」钮 → 高清差异浮层（红框+自动定位） */
  const [zoomPage, setZoomPage] = useState<number | null>(null)
  const lRef = useRef<HTMLDivElement>(null)
  const rRef = useRef<HTMLDivElement>(null)
  const syncing = useRef(false)

  useEffect(() => {
    let alive = true
    void (async () => {
      try {
        const pdfjs = await import('pdfjs-dist')
        const dl = await pdfjs.getDocument({ data: leftBytes.slice() }).promise
        const dr = await pdfjs.getDocument({ data: rightBytes.slice() }).promise
        if (!alive) return
        setLPages(dl.numPages)
        setRPages(dr.numPages)
      } catch (e) {
        if (alive) setLoadErr('页面数据加载失败：' + String((e as Error)?.message ?? e).slice(0, 120) + '（产物/原文件字节可能缺失——重新执行一次替换可重建）')
      }
    })()
    return () => { alive = false }
  }, [leftBytes, rightBytes])

  /* ★ 2026-10-25 用户令「默认定位替换页」：两列页数就绪后自动滚到第一个替换点——
     之前只高亮没滚动（用户截图实锤「没定位到修改的位置」） */
  const didInitScroll = useRef(false)
  useEffect(() => {
    if (didInitScroll.current || !lPages || !rPages) return
    didInitScroll.current = true
    const target = points[0] ?? 1
    setTimeout(() => scrollToPage(target), 120)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lPages, rPages])

  const scrollToPage = (p: number) => {
    setCur(p)
    for (const ref of [lRef, rRef]) {
      const el = ref.current
      if (!el) continue
      const t = el.querySelector(`[data-page="${p}"]`)
      t?.scrollIntoView({ block: 'start', behavior: 'smooth' })
    }
  }
  /* 滚动同步：页对页精确对齐（用户令「不要被动跟随 是同步进行」——
     比例映射 scrollTop 在两列页数/页高不同时会漂移=被动跟随的毛病；
     正确做法：算 a 列顶部可见页码，b 列把同页对齐到同高度） */
  const syncScroll = (from: 'l' | 'r') => {
    if (!sync || syncing.current) return
    syncing.current = true
    const a = from === 'l' ? lRef.current : rRef.current
    const b = from === 'l' ? rRef.current : lRef.current
    if (a && b) {
      const aTop = a.getBoundingClientRect().top
      let page = 1
      for (const s of Array.from(a.querySelectorAll('[data-page]'))) {
        const el = s as HTMLElement
        if (el.getBoundingClientRect().bottom > aTop + 4) { page = Number(el.dataset.page); break }
      }
      const t = b.querySelector(`[data-page="${page}"]`) as HTMLElement | null
      if (t) b.scrollTop += t.getBoundingClientRect().top - b.getBoundingClientRect().top - 4
    }
    requestAnimationFrame(() => { syncing.current = false })
  }
  /** 页码序列 1..N（左右列各自页数） */
  const seq = (n: number) => Array.from({ length: n }, (_, i) => i + 1)

  return createPortal(
    <div className="pr-mask" onClick={onClose}>
      <div className="pr-cmp" onClick={(e) => e.stopPropagation()}>
        <div className="pr-cmp-head">
          <span className="pr-cmp-title" title={leftName}>原文件：{leftName}</span>
          <span className="pr-cmp-title" title={rightName}>修改后：{rightName}</span>
          <label className="pr-cmp-sync" title="一列滚动时另一列跟随">
            <input type="checkbox" checked={sync} onChange={(e) => setSync(e.target.checked)} />同步滚动
          </label>
          <button className="pn-del" title="关闭" onClick={onClose}>×</button>
        </div>
        {points.length > 0 && (
          <div className="pr-cmp-points">
            <span className="pr-cmp-pp">替换点位：</span>
            {points.map((p) => (
              <button key={p} className={'pr-cmp-pt' + (cur === p ? ' pr-cmp-pt--on' : '')} onClick={() => scrollToPage(p)}>
                第 {p} 页
              </button>
            ))}
          </div>
        )}
        {loadErr && <div className="pr-thumbs-err">{loadErr}</div>}
        {!loadErr && (
        <div className="pr-cmp-body">
          <div className="pr-cmp-col" ref={lRef} onScroll={() => syncScroll('l')}>
            {lPages ? seq(lPages).map((p) => (
              <CmpPage key={p} bytes={leftBytes} page={p} cur={cur} isPoint={points.includes(p)} onPick={scrollToPage} onZoom={setZoomPage} />
            )) : <div className="pr-thumb-loading" />}
          </div>
          <div className="pr-cmp-col" ref={rRef} onScroll={() => syncScroll('r')}>
            {rPages ? seq(rPages).map((p) => (
              <CmpPage key={p} bytes={rightBytes} page={p} cur={cur} isPoint={points.includes(p)} onPick={scrollToPage} onZoom={setZoomPage} />
            )) : <div className="pr-thumb-loading" />}
          </div>
        </div>
        )}
        {zoomPage != null && (
          <ZoomDiff
            leftName={leftName}
            leftBytes={leftBytes}
            rightName={rightName}
            rightBytes={rightBytes}
            lPage={zoomPage}
            rPage={zoomPage}
            points={points}
            onClose={() => setZoomPage(null)}
          />
        )}      </div>
    </div>,
    document.body,
  )
}

/* ═══════════ 替换点放大差异对比（高清两页 + 差异文本红框 + 自动定位）═══════════ */

/** 单页渲染 + 提取「视觉行」（坐标聚类：y 行分组 + 行内 x 排序——
 *  ★ pdfjs item 顺序≠视觉顺序，乱序会让 LCS 产生伪差异红框乱套（用户实锤）。
 *    相似度分析（core/pdfReplace）同款行聚类——两侧完全同源） */
async function renderPageGetRows (bytes: Uint8Array, page: number, canvas: HTMLCanvasElement, targetW: number) {
  const pdfjs = await import('pdfjs-dist')
  const doc = await pdfjs.getDocument({ data: bytes.slice() }).promise
  const pg = await doc.getPage(Math.max(1, Math.min(page, doc.numPages)))
  const vp1 = pg.getViewport({ scale: 1 })
  const vp = pg.getViewport({ scale: targetW / vp1.width })
  canvas.width = Math.floor(vp.width)
  canvas.height = Math.floor(vp.height)
  await pg.render({ canvasContext: canvas.getContext('2d')!, viewport: vp }).promise
  const tc = await pg.getTextContent()
  type It = { str: string; rect: MarkRect; yTop: number; xLeft: number }
  const items: It[] = (tc.items as { str?: string; transform?: number[]; width?: number; height?: number }[])
    .map((it) => {
      const tx = pdfjs.Util.transform(vp.transform, it.transform ?? [1, 0, 0, 1, 0, 0])
      const fh = Math.hypot(tx[2], tx[3]) || 10
      return {
        str: it.str ?? '',
        rect: {
          leftPct: (tx[4] / vp.width) * 100,
          topPct: ((tx[5] - fh) / vp.height) * 100,
          wPct: ((it.width || 8) / vp1.width) * 100,
          hPct: (fh / vp.height) * 100,
        },
        yTop: (tx[5] - fh) / vp.height,
        xLeft: tx[4] / vp.width,
      }
    })
    .filter((it) => it.str.trim())
  items.sort((a, b) => (a.yTop - b.yTop) || (a.xLeft - b.xLeft))
  /* 行聚类：top 差 < 行高×0.6+0.15% 视为同一行 */
  const rows: { text: string; rect: MarkRect }[] = []
  let curY = Number.NaN
  let buf = ''
  let rowRects: MarkRect[] = []
  const flush = () => {
    if (!buf) return
    const lft = Math.min(...rowRects.map((r) => r.leftPct))
    const tp = Math.min(...rowRects.map((r) => r.topPct))
    const rt = Math.max(...rowRects.map((r) => r.leftPct + r.wPct))
    const bt = Math.max(...rowRects.map((r) => r.topPct + r.hPct))
    rows.push({ text: buf, rect: { leftPct: lft, topPct: tp, wPct: rt - lft, hPct: bt - tp } })
    buf = ''
    rowRects = []
  }
  for (const it of items) {
    if (Number.isNaN(curY) || Math.abs(it.yTop - curY) > it.rect.hPct * 0.6 + 0.15) { flush(); curY = it.yTop }
    rowRects.push(it.rect)
    buf += (buf ? ' ' : '') + it.str.trim()
  }
  flush()
  void doc
  return { rows }
}

type MarkRect = { leftPct: number; topPct: number; wPct: number; hPct: number }

/** 放大差异对比浮层：fit 整页 + 左上替换点定位（支持多页替换）+ 红框点击放大到修改内容 */
function ZoomDiff ({ leftName, leftBytes, rightName, rightBytes, lPage, rPage, points, onClose }: {
  leftName: string
  leftBytes: Uint8Array
  rightName: string
  rightBytes: Uint8Array
  lPage: number
  rPage: number
  points?: number[]
  onClose: () => void
}) {
  const cL = useRef<HTMLCanvasElement>(null)
  const cR = useRef<HTMLCanvasElement>(null)
  const boxL = useRef<HTMLDivElement>(null)
  const boxR = useRef<HTMLDivElement>(null)
  const [marksL, setMarksL] = useState<(MarkRect | null)[]>([])
  const [marksR, setMarksR] = useState<(MarkRect | null)[]>([])
  const [err, setErr] = useState('')
  const [loading, setLoading] = useState(true)
  /* ★ 多页替换点位：头部左上「定位」按钮组切换（用户令「考虑多个页面替换」） */
  const [page, setPage] = useState(lPage)
  /* ★ 红框点击 → 页内放大到该差异区（transform scale，origin=框中心）；再点还原 */
  const [zoomIdx, setZoomIdx] = useState<number | null>(null)

  useEffect(() => {
    let alive = true
    setLoading(true)
    setMarksL([])
    setMarksR([])
    setZoomIdx(null)
    void (async () => {
      try {
        /* ★ 初始渲染分辨率按「fit 视口」给足清晰度：显示层 max 约束缩到 fit */
        const targetW = Math.max(720, Math.floor(window.innerWidth * 0.42))
        const L = renderPageGetRows(leftBytes, page, cL.current!, targetW)
        const R = renderPageGetRows(rightBytes, page === lPage ? rPage : page, cR.current!, targetW)
        const [l, r] = await Promise.all([L, R])
        if (!alive) return
        /* ★ 行级 LCS diff（2026-10-25 二次修正）：item 乱序 → 按「视觉行」对比——
           每个差异段 = 被修改的行（红框=该行包围盒），与相似度 seqSim 完全同源 */
        const segs = seqDiffSegments(l.rows.map((r) => r.text), r.rows.map((r) => r.text))
        const boxOf = (arr: { rect: MarkRect }[], sIdx: number, eIdx: number): MarkRect | null => {
          const rs = arr.slice(sIdx, eIdx + 1).map((r) => r.rect)
          if (!rs.length) return null
          const lft = Math.min(...rs.map((r) => r.leftPct))
          const tp = Math.min(...rs.map((r) => r.topPct))
          const rt = Math.max(...rs.map((r) => r.leftPct + r.wPct))
          const bt = Math.max(...rs.map((r) => r.topPct + r.hPct))
          return { leftPct: lft, topPct: tp, wPct: rt - lft, hPct: bt - tp }
        }
        if (!alive) return
        setMarksL(segs.map((sg) => boxOf(l.rows, sg.aS, sg.aE)))
        setMarksR(segs.map((sg) => boxOf(r.rows, sg.bS, sg.bE)))
        setLoading(false)
        /* ★ 调试导出（红框位置核对）：行文本/行框/差异段全量进 dev-console——
           用户复现一次即可远端分析（排查记录案例五） */
        console.log('[pr-debug] segs=' + JSON.stringify(segs))
        console.log('[pr-debug] rowsL=' + JSON.stringify(l.rows.map((r) => ({ t: r.text.slice(0, 40), rect: r.rect }))))
        console.log('[pr-debug] rowsR=' + JSON.stringify(r.rows.map((r) => ({ t: r.text.slice(0, 40), rect: r.rect }))))
      } catch (e) {
        if (alive) { setErr(String((e as Error)?.message ?? e).slice(0, 90)); setLoading(false) }
      }
    })()
    return () => { alive = false }
  }, [leftBytes, rightBytes, page, lPage, rPage])

  /* 红框点击 → 两列同时放大到各自对应差异区（同 diff 序号对称），定位框中心 */
  const pickMark = (i: number) => {
    setZoomIdx((v) => (v === i ? null : i))
    setTimeout(() => {
      for (const [marks, box] of [[marksL, boxL], [marksR, boxR]] as [(MarkRect | null)[], React.RefObject<HTMLDivElement | null>][]) {
        const m = marks[i]
        const bx = box.current
        if (!m || !bx) continue
        bx.scrollTop = (m.topPct / 100) * bx.scrollHeight - bx.clientHeight / 2 + (m.hPct / 100) * bx.scrollHeight / 2
        bx.scrollLeft = (m.leftPct / 100) * bx.scrollWidth - bx.clientWidth / 2 + (m.wPct / 100) * bx.scrollWidth / 2
      }
    }, 80)
  }

  const renderCol = (
    canvasRef: React.RefObject<HTMLCanvasElement>,
    boxRef: React.RefObject<HTMLDivElement>,
    marks: (MarkRect | null)[],
    cap: string,
    name: string,
    pg: number,
  ) => (
    <div className="pr-zoom-col">
      <div className="pr-thumb-cap">{cap} · {name} · 第 {pg} 页{marks.length ? ` · ${marks.length} 处差异` : ''}</div>
      <div className="pr-zoom-pagebox" ref={boxRef}>
        {loading && <div className="pr-thumb-loading">渲染中…</div>}
        {/* wrap 只包 canvas：红框百分比相对页面本身（fit 后依旧对齐） */}
        <div
          className={'pr-zoom-wrap' + (zoomIdx != null ? ' pr-zoom-wrap--zoomed' : '')}
          style={zoomIdx != null && marks[zoomIdx] ? { transformOrigin: `${marks[zoomIdx].leftPct}% ${marks[zoomIdx].topPct}%` } : undefined}
        >
          <canvas ref={canvasRef} style={{ display: loading ? 'none' : undefined }} />
          {!loading && marks.map((m, i) => m ? (
            <div
              key={i}
              className={'pr-zoom-mark' + (zoomIdx === i ? ' pr-zoom-mark--on' : '')}
              style={{ left: m.leftPct + '%', top: m.topPct + '%', width: m.wPct + '%', height: m.hPct + '%' }}
              title="点击放大到该处修改"
              onClick={(e) => { e.stopPropagation(); pickMark(i) }}
            />
          ) : null)}
        </div>
      </div>
    </div>
  )

  return createPortal(
    <div className="pr-mask" onClick={onClose}>
      <div className="pr-zoom" onClick={(e) => e.stopPropagation()}>
        <div className="pr-zoom-head">
          {/* ★ 左上定位：多页替换点位切换（用户令「考虑多个页面替换」） */}
          {points && points.length > 0 && (
            <span className="pr-zoom-pts">
              <span className="pr-cmp-pp">定位：</span>
              {points.map((p) => (
                <button key={p} className={'pr-cmp-pt' + (page === p ? ' pr-cmp-pt--on' : '')} onClick={() => setPage(p)}>
                  第 {p} 页
                </button>
              ))}
            </span>
          )}
          <span>修改处放大对比（红框=差异文本，点击红框放大到修改内容）</span>
          <button className="pn-del" title="关闭" onClick={onClose}>×</button>
        </div>
        {err ? <div className="pr-thumbs-err">{err}</div> : (
          <div className="pr-zoom-body">
            {renderCol(cL, boxL, marksL, '原页', leftName, page)}
            <div className="pr-thumb-vs">vs</div>
            {renderCol(cR, boxR, marksR, '替换页', rightName, page)}
          </div>
        )}
      </div>
    </div>,
    document.body,
  )
}

/** 匹配确认弹窗（表格：# / 被替换页 / 相似度 / 匹配来源 / 差异预览 / 被替换文件名） */
function PdfReplaceConfirm ({ nodeId, pairs, skipped, onCancel, onConfirm }: {
  nodeId: string
  pairs: ReplacePair[]
  skipped: string[]
  onCancel: () => void
  onConfirm: (pairs: ReplacePair[]) => void
}) {
  const [thumb, setThumb] = useState<ReplacePair | null>(null)
  /* ★ 2026-10-25 用户令「复制差异情况（那几页有差别）」：整表摘要进剪贴板 */
  const copyDiff = async () => {
    const lines = [
      `PDF 替换对比：共 ${pairs.length} 对`,
      ...pairs.map((p, i) => `${i + 1}. 被替换：${p.tgtName} 第${p.tgtPage}页 ← 替换：${p.srcName} 第${p.srcPage}页（相似 ${p.sim}%）`),
    ]
    try {
      await navigator.clipboard.writeText(lines.join('\n'))
      pushLog(`[PDF·替换] 差异情况已复制（${pairs.length} 对）`, 'ok')
    } catch { pushLog('[PDF·替换] 复制失败（剪贴板不可用）', 'err') }
  }
  /* ★ 用户令「切换页面的按钮 轮询不同的页面对比 上下切换」：预览区头部 上一对/下一对 循环切换 */
  const stepThumb = (dir: 1 | -1) => {
    if (!pairs.length) return
    const i = pairs.indexOf(thumb ?? pairs[0])
    setThumb(pairs[(i + dir + pairs.length) % pairs.length])
  }
  const thumbIdx = thumb ? pairs.indexOf(thumb) : -1
  return createPortal(
    /* ★ 2026-10-25 用户令「不要点击周边就退出」：遮罩点击不再触发取消——退出只认 × / 取消按钮 */
    <div className="pr-mask">
      <div className="pr-confirm" onClick={(e) => e.stopPropagation()}>
        <div className="pr-confirm-head">
          <span>替换对比 · 匹配到 <b>{pairs.length}</b> 对</span>
          <button className="pn-del" title="关闭" onClick={onCancel}>×</button>
        </div>
        <div className="pr-cf-table">
          <div className="pr-cf-row pr-cf-row--head">
            <span>#</span>
            <span>被替换页</span>
            <span>相似度</span>
            <span>匹配来源（替换页）</span>
            <span>差异预览</span>
            <span>被替换文件</span>
          </div>
          {pairs.map((p, i) => (
            <div
              key={i}
              className={'pr-cf-row pr-cf-row--click' + (thumb === p ? ' pr-cf-row--on' : '') + (p.lowSim ? ' pr-cf-row--low' : '')}
              title={p.lowSim ? `⚠ 相似度 ${p.sim}% 低于容差——请人工确认是否为正确配对` : `点击在下方展开左右并排差异预览`}
              onClick={() => setThumb((v) => (v === p ? null : p))}
            >
              <span className="pn-idx">{i + 1}</span>
              <span className="pr-cf-page">第 <b>{p.tgtPage}</b> 页</span>
              <span>
                <span className={'pr-sim ' + (p.sim >= 80 ? 'pr-sim--hi' : p.sim >= 50 ? 'pr-sim--mid' : 'pr-sim--low')}>{p.sim}%</span>
                {p.lowSim && <span className="pr-cf-lowwarn" title={`相似度低于容差，请人工确认`}>⚠低</span>}
              </span>
              <span className="pr-cf-src" title={p.srcName}>
                <span className="pm-fmt pm-fmt-pdf">P</span>
                <span className="pn-nm">{p.srcName}</span>
                <span className="pr-cf-pp">第{p.srcPage}页</span>
              </span>
              <span>
                <button
                  className={'pr-diff' + (thumb === p ? ' pr-diff--on' : '')}
                  title={thumb === p ? '收起下方差异预览' : '在下方展开左右并排差异预览'}
                  onClick={(e) => { e.stopPropagation(); setThumb((v) => (v === p ? null : p)) }}
                >
                  <IcGitCompare size={14} />
                </button>
              </span>
              <span className="pr-cf-tgt" title={p.tgtName}>{p.tgtName}</span>
            </div>
          ))}
          {/* ★ 跳过详情可见（用户令「要替换3页 只出2对」——缺的那页去哪了一目了然） */}
          {skipped.length > 0 && (
            <div className="pr-cf-skip" title={skipped.join('\n')}>
              ⚠ 另有 {skipped.length} 页未进列表：{skipped[0]}
              {skipped.length > 1 ? ` 等 ${skipped.length} 项（调低容差或手改页码后重试）` : ''}
            </div>
          )}
          {/* ★ 差异预览内嵌弹窗下方（左右分裂）+ 上一对/下一对轮询切换 */}
          {thumb && (
            <div className="pr-inline-preview">
              <div className="pr-ip-head">
                <span className="pr-ip-switch">
                  <button className="pr-btn pr-btn--sm" title="上一对" onClick={() => stepThumb(-1)}>▲ 上一对</button>
                  <span>第 {thumbIdx + 1} / {pairs.length} 对</span>
                  <button className="pr-btn pr-btn--sm" title="下一对" onClick={() => stepThumb(1)}>下一对 ▼</button>
                </span>
                <span className="pr-ip-cap">差异预览 · {thumb.tgtName} 第{thumb.tgtPage}页 ⇄ {thumb.srcName} 第{thumb.srcPage}页（相似 {thumb.sim}%）</span>
                <button className="pn-del" title="关闭预览" onClick={() => setThumb(null)}>×</button>
              </div>
              <ThumbPair nodeId={nodeId} pair={thumb} />
            </div>
          )}
        </div>
        <div className="pr-confirm-foot">
          {/* ★ 2026-10-25 用户令「取消和关闭冲突 都放右边 去掉一个」：去关闭，取消归右统
一退出 */}
          <button className="pr-btn pr-btn--go" onClick={() => onConfirm(pairs)}>执行替换（产物进历史）</button>
          <button className="pr-btn pr-btn--tail" onClick={() => void copyDiff()} title="复制哪些页有差别（摘要进剪贴板）">复制差异情况</button>
          <button className="pr-btn" onClick={onCancel}>取消</button>
        </div>
      </div>
    </div>,
    document.body,
  )
}

/* ═══════════ 历史方形钮（按钮行内——与锁/换并排，用户指正「历史按钮是方形的」） ═══════════ */

export function PdfReplaceHistButton ({ nodeId }: { nodeId: string }) {
  const table = useBaseStore((s) => s.tables[nodeId])
  const open = useSyncExternalStore(prHistUI.sub, prHistUI.get)
  const count = (((table?.config ?? {}) as { replaceHistory?: ReplaceHistEntry[] }).replaceHistory ?? []).length
  if (!count) return null
  return (
    <button
      className={'pm-sq-btn pm-sq-btn--ghost nodrag nopan' + (open ? ' pm-sq-btn--on' : '')}
      title={`展开/收起替换历史（${count} 条）`}
      onClick={() => prHistUI.toggle()}
    >
      <IcHistory size={19} />
    </button>
  )
}

/* ═══════════ 历史列表（🕘 方形钮控制展开——列表渲染在按钮行后） ═══════════ */

export function PdfReplaceHistList ({ nodeId }: { nodeId: string }) {
  const table = useBaseStore((s) => s.tables[nodeId])
  const open = useSyncExternalStore(prHistUI.sub, prHistUI.get)
  const [previewBytes, setPreviewBytes] = useState<Uint8Array | null>(null)
  const [previewName, setPreviewName] = useState('')
  const [previewMax, setPreviewMax] = useState(false)
  const [menu, setMenu] = useState<{ x: number; y: number; entry: ReplaceHistEntry } | null>(null)
  /* ★ 整档对比 state——必须在早退 return 之前（条件 hooks = React #310 白屏实锤） */
  const [cmp, setCmp] = useState<{ left: { name: string; bytes: Uint8Array }; right: { name: string; bytes: Uint8Array }; points: number[] } | null>(null)
  if (!table || !open) return null
  const cfg = (table.config ?? {}) as { replaceTargets?: ReplaceEntry[]; replaceHistory?: ReplaceHistEntry[] }
  const history = cfg.replaceHistory ?? []
  if (!history.length) return null

  const flash = (m: string) => pushLog('[PDF·替换] ' + m, 'warn')

  const openPreview = async (h: ReplaceHistEntry) => {
    try {
      const bytes = await readReplaceBytes(h)
      if (!bytes) { flash('产物已不在本机：' + h.name); return }
      setPreviewBytes(bytes)
      setPreviewName(h.name)
    } catch (e) { flash('打开失败：' + String((e as Error)?.message ?? e).slice(0, 60)) }
  }

  /** 与原文件（上组第一份）视觉 diff */
  const diffEntry = async (h: ReplaceHistEntry) => {
    try {
      const src = (cfg.replaceTargets ?? [])[0]
      if (!src) { flash('上组为空——无法确定对比基准（原文件）'); return }
      const out = await readReplaceBytes(h)
      const orig = await readReplaceBytes(src)
      if (!out || !orig) { flash('字节缺失，无法对比'); return }
      const api = (globalThis as unknown as { api?: { pdfDiffView?: (a: { name: string; bytes: Uint8Array }, b: { name: string; bytes: Uint8Array }) => Promise<{ ok: boolean; reason?: string }> } }).api
      if (!api?.pdfDiffView) { flash('桌面桥不可用（需重启应用加载 diff 桥）'); return }
      const r = await api.pdfDiffView({ name: '原文件_' + src.name, bytes: orig }, { name: h.name, bytes: out })
      if (!r.ok) flash('diff 失败：' + (r.reason ?? ''))
    } catch (e) { flash('diff 失败：' + String((e as Error)?.message ?? e).slice(0, 60)) }
  }

  const removeHist = (h: ReplaceHistEntry) => {
    const st = useBaseStore.getState()
    st.pushHistory()
    st.setNodeConfig(nodeId, { replaceHistory: history.filter((x) => x.id !== h.id) })
    pushLog(`[PDF·替换] 已删除历史：${h.name}`)
  }

  /* ★ 2026-10-25 用户指令「修改后对比图 两个文件 左右两列 同时定位到相同位置
     能选择替换的点位进行对比查看」：⇄ = 应用内整档对比（原文件 vs 修改后产物）
     ★★ cmp state 必须在函数顶部（早退 return 之前）——条件 hooks = React #310 白屏实锤 */
  const openCompare = async (h: ReplaceHistEntry) => {
    try {
      const src = (cfg.replaceTargets ?? [])[0]
      if (!src) { flash('上组为空——无法确定对比基准（原文件）'); return }
      const out = await readReplaceBytes(h)
      const orig = await readReplaceBytes(src)
      if (!out || !orig) { flash('字节缺失，无法对比'); return }
      setCmp({ left: { name: src.name, bytes: orig }, right: { name: h.name, bytes: out }, points: h.pages ?? [] })
    } catch (e) { flash('打开对比失败：' + String((e as Error)?.message ?? e).slice(0, 60)) }
  }

  return (
    <>
      {open && (
        <div className="pr-hist">
          {history.map((h, i) => (
            <div
              key={h.id}
              className="pn-row nodrag nopan pm-hist-row"
              title={`${h.name}\n双击预览；右键更多；⇄ 修改前后对比；× 删除`}
              onDoubleClick={() => void openPreview(h)}
              onContextMenu={(e) => { e.preventDefault(); e.stopPropagation(); setMenu({ x: e.clientX, y: e.clientY, entry: h }) }}
              /* ★ 拦右键的 pointer 事件——否则 pointerup 冒泡到画布容器触发节点菜单
                 （「置于底层」与条目菜单同弹——2026-10-25 用户截图实锤；pm-hist-row 同款） */
              onPointerDown={(e) => e.stopPropagation()}
              onPointerUp={(e) => e.stopPropagation()}
            >
              <span className="pn-idx">{i + 1}</span>
              <span className="pm-fmt pm-fmt-pdf">P</span>
              <span
                className={'pn-loc ' + ((h.inDb || h.b64) ? 'pn-loc--in' : 'pn-loc--ext')}
                title={(h.inDb || h.b64) ? '软件内部：字节已入库，可随工程文件存储' : '外部引用：字节不在软件库，可能丢失'}
              >{(h.inDb || h.b64) ? '内' : '外'}</span>
              <span className="pn-nm">{h.name}</span>
              <span className="pr-sz">{fmtTime(h.at)}</span>
              <button className="pr-diff" title="修改前后对比（左右两列同步定位）" onClick={(e) => { e.stopPropagation(); void openCompare(h) }}>
                <IcGitCompare size={14} />
              </button>
              <button className="pn-del" title="删除记录" onClick={(e) => { e.stopPropagation(); removeHist(h) }}>×</button>
            </div>
          ))}
        </div>
      )}
      {/* 预览大窗（复用 PDF 合并 pm-pdfpreview——用户指令「类似PDF合并 可以双击预览」） */}
      {previewBytes && createPortal(
        <div className="pm-pdfpreview" onClick={() => { setPreviewBytes(null); setPreviewMax(false) }}>
          <div className={'pn-preview-card pm-pdfpreview-card' + (previewMax ? ' pm-card-max' : '')} onClick={(e) => e.stopPropagation()}>
            <div className="pn-preview-bar">
              <span className="pn-preview-nm">🕘 {previewName}</span>
              <button className="pn-preview-close" title={previewMax ? '还原' : '最大化'} onClick={() => setPreviewMax((v) => !v)}>
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
      {/* 修改前后整档对比（左右两列 + 点位跳转 + 同步滚动） */}
      {cmp && createPortal(
        <PdfCompareView
          leftName={cmp.left.name}
          leftBytes={cmp.left.bytes}
          rightName={cmp.right.name}
          rightBytes={cmp.right.bytes}
          points={cmp.points}
          onClose={() => setCmp(null)}
        />,
        document.body,
      )}
      {/* 右键菜单（与 PDF 合并历史条目一致：打开/PDF.js/默认软件/MuPDF/另存为/删除） */}
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
            <button className="pm-ctx-item" onClick={() => { const h = menu.entry; setMenu(null); void openPreview(h) }}>用 PDF.js Viewer 打开</button>
            <button
              className="pm-ctx-item"
              onClick={async () => {
                const h = menu.entry
                setMenu(null)
                try {
                  const bytes = await readReplaceBytes(h)
                  if (!bytes) { flash('产物已不在本机：' + h.name); return }
                  const api = (window as unknown as { api?: { openExternalFile?: (n: string, d: Uint8Array) => Promise<string> } }).api
                  if (!api?.openExternalFile) { flash('桌面桥不可用'); return }
                  await api.openExternalFile(h.name, bytes)
                } catch (eD) { flash('打开失败：' + String((eD as Error)?.message ?? eD).slice(0, 60)) }
              }}
            >用默认软件打开</button>
            {/* ★ 2026-10-26 用户指令「所有PDF增加个 用本地浏览器打开」：Edge/Chrome 自带查看器 */}
            <button
              className="pm-ctx-item"
              onClick={async () => {
                const h = menu.entry
                setMenu(null)
                try {
                  const bytes = await readReplaceBytes(h)
                  if (!bytes) { flash('产物已不在本机：' + h.name); return }
                  const api = (window as unknown as { api?: { openInBrowser?: (n: string, d: Uint8Array) => Promise<string> } }).api
                  if (!api?.openInBrowser) { flash('桌面桥不可用'); return }
                  await api.openInBrowser(h.name, bytes)
                } catch (eB) { flash('浏览器打开失败：' + String((eB as Error)?.message ?? eB).slice(0, 60)) }
              }}
            >用本地浏览器打开</button>
            <button
              className="pm-ctx-item"
              onClick={async () => {
                const h = menu.entry
                setMenu(null)
                try {
                  const bytes = await readReplaceBytes(h)
                  if (!bytes) { flash('产物已不在本机：' + h.name); return }
                  const api = (window as unknown as { api?: { openExternalFile?: (n: string, d: Uint8Array) => Promise<string> } }).api
                  if (!api?.openExternalFile) { flash('桌面桥不可用'); return }
                  await api.openExternalFile(h.name, bytes)
                } catch (eM) { flash('打开失败：' + String((eM as Error)?.message ?? eM).slice(0, 60)) }
              }}
            >用 MuPDF 打开</button>
            <button
              className="pm-ctx-item"
              onClick={async () => {
                const h = menu.entry
                setMenu(null)
                try {
                  const bytes = await readReplaceBytes(h)
                  if (!bytes) { flash('产物已不在本机：' + h.name); return }
                  const api = (window as unknown as { api?: { docSaveAs?: (n: string, d: Uint8Array, k?: string) => Promise<string | null> } }).api
                  if (api?.docSaveAs) {
                    const p = await api.docSaveAs(h.name, bytes, 'pdf')
                    if (p) pushLog('[PDF·另存为] ' + h.name + ' → ' + p, 'ok')
                  } else {
                    const blob = new Blob([bytes as BlobPart], { type: 'application/pdf' })
                    const u = URL.createObjectURL(blob) + '#zoom=100'
                    const a = document.createElement('a')
                    a.href = u
                    a.download = h.name
                    a.click()
                    setTimeout(() => URL.revokeObjectURL(u), 3000)
                    pushLog('[PDF·另存为] Web 环境已按默认下载：' + h.name, 'warn')
                  }
                } catch (eS) { flash('另存失败：' + String((eS as Error)?.message ?? eS).slice(0, 60)) }
              }}
            >另存为…</button>
            <button className="pm-ctx-item" onClick={() => { const h = menu.entry; setMenu(null); removeHist(h) }}>删除记录</button>
          </div>
        </div>,
        document.body,
      )}
    </>
  )
}

/* ═══════════ 节点卡片主体（上下两组词条） ═══════════ */

export function PdfReplaceNodeCard ({ nodeId }: { nodeId: string }) {
  const table = useBaseStore((s) => s.tables[nodeId])
  /* ★ 2026-10-25 用户指令「拖进去要有词条贴入提示 就像PDF节点一样 也不要弹窗」：
     组件内自处理拖入（同 PdfNodeCard 模式）——悬停按鼠标 Y 高亮对应组，松手直接贴入 */
  const [dragOver, setDragOver] = useState<'upper' | 'lower' | null>(null)
  const dragDepth = useRef(0)
  /* ★ 2026-10-25 用户指令「节点内的文件可以跨两个类型拖动」：词条可拖到另一组（组间移动） */
  const [draggingId, setDraggingId] = useState<string | null>(null)
  /* ★ 2026-10-25 用户指令「空态做成导入文件的长条按钮」：点击弹文件选择直接落对应组 */
  const fileRef = useRef<HTMLInputElement>(null)
  const [pickUpper, setPickUpper] = useState(true)
  /* ★ 锚点动态增删重测量（连线端点不跳——pdfmerge 教训） */
  const updateNodeInternals = useUpdateNodeInternals()
  const cfgAll = (table?.config ?? {}) as { replaceTargets?: ReplaceEntry[]; replaceSources?: ReplaceEntry[] }
  const entryCount = (cfgAll.replaceTargets ?? []).length + (cfgAll.replaceSources ?? []).length
  useEffect(() => { updateNodeInternals(nodeId) }, [entryCount, nodeId, updateNodeInternals])
  if (!table) return null
  const cfg = (table.config ?? {}) as { replaceTargets?: ReplaceEntry[]; replaceSources?: ReplaceEntry[] }
  const targets = cfg.replaceTargets ?? []
  const sources = cfg.replaceSources ?? []

  /** 拖入落组（无弹窗——用户指令）：按鼠标 Y 判上/下组，字节入库直接贴入 */
  const importFiles = async (files: FileList, upper: boolean) => {
    const pdfs = Array.from(files).filter((f) => (f.name.toLowerCase().split('.').pop() ?? '') === 'pdf')
    if (!pdfs.length) { pushLog('[PDF·替换] 仅支持 PDF 文件', 'warn'); return }
    const { putBlob } = await import('../core/pdfBlobStore')
    const key = upper ? 'replaceTargets' : 'replaceSources'
    const st = useBaseStore.getState()
    st.pushHistory()
    const cur = ((st.tables[nodeId]?.config as Record<string, unknown> ?? {})[key] as ReplaceEntry[]) ?? []
    const adds = await Promise.all(pdfs.map(async (f) => {
      const bid = 'reprepl-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
      await putBlob(bid, new Uint8Array(await f.arrayBuffer()))
      return { id: bid, name: f.name, inDb: true, page: 1 }
    }))
    st.setNodeConfig(nodeId, { [key]: [...cur, ...adds] })
    pushLog(`[PDF·替换] 已加入${upper ? '被替换' : '替换页'}组：${pdfs.map((f) => f.name).join('、')}`, 'ok')
  }

  const setList = (key: 'replaceTargets' | 'replaceSources', list: ReplaceEntry[]) => {
    const st = useBaseStore.getState()
    st.setNodeConfig(nodeId, { [key]: list })
  }
  const removeAt = (key: 'replaceTargets' | 'replaceSources', i: number) => {
    const list = [...(key === 'replaceTargets' ? targets : sources)]
    const [rm] = list.splice(i, 1)
    setList(key, list)
    pushLog(`[PDF·替换] 已移除${key === 'replaceTargets' ? '被替换' : '替换页'}：${rm.name}`)
  }
  const setPage = (key: 'replaceTargets' | 'replaceSources', i: number, page: number) => {
    const list = (key === 'replaceTargets' ? targets : sources).map((x, j) => (j === i ? { ...x, page: Math.max(1, Math.floor(page || 1)) } : x))
    setList(key, list)
  }

  const row = (key: 'replaceTargets' | 'replaceSources', list: ReplaceEntry[]) => (
    list.map((it, i) => (
      <div
        key={it.id}
        className="pn-row nodrag nopan"
        title={`${it.name}（页码 = ${key === 'replaceTargets' ? '被换掉的页' : '取这页替换进去'}；拖动词条可移到另一组；行首锚点可接上游 PDF 自动集成）`}
        draggable
        onDragStart={(e) => {
          e.stopPropagation()
          try {
            e.dataTransfer.effectAllowed = 'move'
            e.dataTransfer.setData('application/x-pr-move', JSON.stringify({ from: key, id: it.id }))
            e.dataTransfer.setData('text/plain', it.name) /* Firefox：至少一个类型才启动拖动 */
          } catch { /* ignore */ }
          setDraggingId(it.id)
          pushLog(`[PDF·替换] 拖动中：${it.name}（${key === 'replaceTargets' ? '被替换' : '替换页'}组 → 拖到另一组松手）`)
        }}
        onDragEnd={() => { setDraggingId(null); setDragOver(null); dragDepth.current = 0 }}
        style={draggingId === it.id ? { opacity: 0.35 } : undefined}
      >
        {/* ★ 2026-10-25 用户指令「每个词条前面设置锚点 上游连接过来默认集成上游文件」：
            行首数据流入口（target）——上游 PDF 节点/PDF 合并金锚点连进来即覆盖本条 */}
        <Handle
          type="target"
          position={Position.Left}
          id={'pr-' + (key === 'replaceTargets' ? 'upper' : 'lower') + '-' + it.id}
          className="nodrag nopan pr-anchor"
          style={{ position: 'absolute', left: -9, top: '50%', transform: 'translateY(-50%)', width: 9, height: 9, background: '#3b82f6', border: '1.5px solid var(--c-bg0)', borderRadius: '50%' }}
          title="连接上游 PDF，自动集成到本位置"
        />
        <span className="pn-idx">{i + 1}</span>
        <span className="pm-fmt pm-fmt-pdf">P</span>
        <span
          className={'pn-loc ' + ((it.inDb || it.b64) ? 'pn-loc--in' : 'pn-loc--ext')}
          title={(it.inDb || it.b64) ? '软件内部：字节已入库，可随工程文件存储' : '外部引用：字节不在软件库，可能丢失'}
        >{(it.inDb || it.b64) ? '内' : '外'}</span>
        <span className="pn-nm">{it.name}</span>
        <span className="pr-page nodrag nopan">
          页<input
            type="number"
            min={1}
            value={it.page}
            draggable={false}
            onChange={(e) => setPage(key, i, parseInt(e.target.value || '1', 10))}
            onClick={(e) => e.stopPropagation()}
          />
        </span>
        <button className="pn-del" draggable={false} title="移除这份" onClick={() => removeAt(key, i)}>×</button>
      </div>
    ))
  )

  return (
    <div
      className="pr-root"
      onDoubleClick={(e) => e.stopPropagation()}
      onDragEnter={(e) => { e.preventDefault(); dragDepth.current++; setDragOver(e.clientY < (e.currentTarget as HTMLElement).getBoundingClientRect().top + (e.currentTarget as HTMLElement).getBoundingClientRect().height / 2 ? 'upper' : 'lower') }}
      onDragOver={(e) => {
        e.preventDefault()
        /* 内部移动（词条跨组）与外部文件导入互斥：按 dataTransfer 类型区分光标语义 */
        const internal = e.dataTransfer.types.includes('application/x-pr-move')
        try { e.dataTransfer.dropEffect = internal ? 'move' : 'copy' } catch { /* ignore */ }
        setDragOver(e.clientY < (e.currentTarget as HTMLElement).getBoundingClientRect().top + (e.currentTarget as HTMLElement).getBoundingClientRect().height / 2 ? 'upper' : 'lower')
      }}
      onDragLeave={() => { dragDepth.current--; if (dragDepth.current <= 0) { dragDepth.current = 0; setDragOver(null) } }}
      onDrop={async (e) => {
        e.preventDefault()
        e.stopPropagation()
        dragDepth.current = 0
        /* ★ drop 即拖动结束：先清拖动态（移动成功后源行卸载，dragend 会丢失——
           不清的话 draggingId 残留，新组里同 id 词条继承半透明灰） */
        setDraggingId(null)
        /* ★ 组别从鼠标坐标现算（不读 dragOver state——拖放高频 setState 时闭包可能是旧值） */
        const rect = (e.currentTarget as HTMLElement).getBoundingClientRect()
        const upper = e.clientY < rect.top + rect.height / 2
        setDragOver(null)
        /* ★ 跨组移动：词条从一组拖到另一组松手落位（引用不变：id/字节/页码原样带走） */
        const mv = e.dataTransfer.getData('application/x-pr-move')
        if (mv) {
          try {
            const { from, id } = JSON.parse(mv) as { from: 'replaceTargets' | 'replaceSources'; id: string }
            const to: 'replaceTargets' | 'replaceSources' = upper ? 'replaceTargets' : 'replaceSources'
            if (from !== to) {
              const src = [...(from === 'replaceTargets' ? targets : sources)]
              const idx = src.findIndex((x) => x.id === id)
              if (idx >= 0) {
                const [moved] = src.splice(idx, 1)
                const dst = [...(to === 'replaceTargets' ? targets : sources), moved]
                useBaseStore.getState().pushHistory()
                useBaseStore.getState().setNodeConfig(nodeId, { [from]: src, [to]: dst })
                pushLog(`[PDF·替换] 已移动到${to === 'replaceTargets' ? '被替换' : '替换页'}组：${moved.name}`, 'ok')
              } else {
                pushLog(`[PDF·替换] 移动失败：源组未找到词条 ${id}`, 'warn')
              }
            } else {
              pushLog('[PDF·替换] 拖回原组，无变化')
            }
          } catch (ex) { pushLog('[PDF·替换] 移动异常：' + String((ex as Error)?.message ?? ex).slice(0, 60), 'warn') }
          return
        }
        if (e.dataTransfer.files?.length) await importFiles(e.dataTransfer.files, upper)
      }}
    >
      <div style={{ position: 'relative' }} className={'pr-group-head' + (dragOver === 'upper' ? ' pr-group-head--over' : '')} title="被替换的 PDF（页码 = 要被换掉的页）；行首锚点=上游连线追加到本组">
        {/* ★ 2026-10-25 回退 portal→行内贴边（getBoundingClientRect 含画布缩放=乱飘根因；
            行内 absolute 相对行定位零缩放依赖——位置永远正确） */}
        <Handle
          type="target"
          position={Position.Left}
          id="pr-add-upper"
          className="nodrag nopan pr-anchor"
          style={{ position: 'absolute', left: -9, top: '50%', transform: 'translateY(-50%)', width: 9, height: 9, background: '#3b82f6', border: '1.5px solid var(--c-bg0)', borderRadius: '50%' }}
          title="连接上游 PDF，追加到被替换组"
        />
        被替换（改页码=被换掉的页）{dragOver === 'upper' ? (draggingId ? ' —— 松手移到本组' : ' —— 松手加入') : ''}
      </div>
      <div className={'pr-group' + (dragOver === 'upper' ? ' pr-group--over' : '')}>
        {targets.length === 0 && (
          <button className="pn-empty-btn nodrag nopan" title="点击导入被替换 PDF（可多选）" onClick={() => { setPickUpper(true); fileRef.current?.click() }}>
            ＋ 点击导入被替换的 PDF
          </button>
        )}
        {row('replaceTargets', targets)}
      </div>
      <div style={{ position: 'relative' }} className={'pr-group-head' + (dragOver === 'lower' ? ' pr-group-head--over' : '')} title="替换页来源 PDF（页码 = 取这页替换进上面的文件）；行首锚点=上游连线追加到本组">
        <Handle
          type="target"
          position={Position.Left}
          id="pr-add-lower"
          className="nodrag nopan pr-anchor"
          style={{ position: 'absolute', left: -9, top: '50%', transform: 'translateY(-50%)', width: 9, height: 9, background: '#3b82f6', border: '1.5px solid var(--c-bg0)', borderRadius: '50%' }}
          title="连接上游 PDF，追加到替换页组"
        />
        替换页（取这页 → 替换进上面）{dragOver === 'lower' ? (draggingId ? ' —— 松手移到本组' : ' —— 松手加入') : ''}
      </div>
      <div className={'pr-group' + (dragOver === 'lower' ? ' pr-group--over' : '')}>
        {sources.length === 0 && (
          <button className="pn-empty-btn nodrag nopan" title="点击导入替换页来源 PDF（可多选）" onClick={() => { setPickUpper(false); fileRef.current?.click() }}>
            ＋ 点击导入替换页来源 PDF
          </button>
        )}
        {row('replaceSources', sources)}
      </div>
      <input
        ref={fileRef}
        type="file"
        accept="application/pdf,.pdf"
        multiple
        style={{ display: 'none' }}
        onChange={async (e) => { const fs0 = e.target.files; if (fs0?.length) await importFiles(fs0, pickUpper); e.target.value = '' }}
      />
    </div>
  )
}
