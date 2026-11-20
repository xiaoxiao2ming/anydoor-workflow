/**
 * PDF 查看器（react-pdf 版）+ 文档智能三件套
 * ★ 2026-10-24 用户指令「需要有识别能力的控件：能识别、标记、定位 pdf 位置」：
 * - 识别：pdf.js getTextContent 全文档文字提取（识别文本层）
 * - 定位：全文搜索 → 结果列表（页码+上下文）→ 点击跳页 + CSS Highlight API 高亮
 * - 标记：选中文本 → 「标记」→ 高亮矩形持久化（比例坐标，缩放不变），标记列表可跳转/删除
 * 基础：双层丝滑缩放（视觉层即时 + 渲染层防抖高清）+ 完整工具条 + 暗色适配
 */
import { useEffect, useMemo, useRef, useState } from 'react'
import { Document, Page, pdfjs } from 'react-pdf'
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import 'react-pdf/dist/Page/AnnotationLayer.css'
import 'react-pdf/dist/Page/TextLayer.css'

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl

/** 标记：比例坐标（相对页宽高 0~1），缩放无关 */
export interface PdfMark {
  id: string
  page: number
  text: string
  rects: { l: number; t: number; w: number; h: number }[]
}

export function PdfViewer ({
  data,
  dark,
  marks,
  onSaveMarks,
}: {
  data: Uint8Array
  dark: boolean
  marks?: PdfMark[]
  onSaveMarks?: (m: PdfMark[]) => void
}) {
  const [numPages, setNumPages] = useState(0)
  const [page, setPage] = useState(1)
  const [pageInput, setPageInput] = useState('1')
  const [renderScale, setRenderScale] = useState(1)
  const [viewScale, setViewScale] = useState(1)
  const [fitWidth, setFitWidth] = useState(true)
  const [err, setErr] = useState('')
  const [loaded, setLoaded] = useState(false)
  const wrapRef = useRef<HTMLDivElement | null>(null)
  const canvasBoxRef = useRef<HTMLDivElement | null>(null)
  const [wrapW, setWrapW] = useState(0)
  const baseWRef = useRef(794)
  const baseHRef = useRef(1123)
  const commitT = useRef<number | undefined>(undefined)

  /* ── 搜索（识别+定位） ── */
  const [query, setQuery] = useState('')
  const [hits, setHits] = useState<{ page: number; ctx: string }[]>([])
  const [hitIdx, setHitIdx] = useState(-1)
  const [searching, setSearching] = useState(false)
  const searchDocRef = useRef<any>(null)

  /* ── 标记 ── */
  const [markBtn, setMarkBtn] = useState<{ x: number; y: number; rects: PdfMark['rects']; text: string } | null>(null)
  const [showMarkList, setShowMarkList] = useState(false)

  const file = useMemo(() => ({ data: data.slice(0) }), [data])

  /* ★ 2026-10-24 用户实测「切换文件后不显示」：data 变化时 loaded/err/page 未重置——
     旧文件的 loaded=true 卡住 loading 提示，新 Document 渲染期间全空白。 */
  useEffect(() => {
    setLoaded(false)
    setErr('')
    setPage(1)
    setPageInput('1')
    setViewScale(1)
    setRenderScale(1)
    setFitWidth(true)
  }, [data])

  /* 全文档文字提取（识别）——缓存 doc 实例供搜索复用 */
  useEffect(() => {
    let dead = false
    try {
      const task = pdfjs.getDocument({ data: data.slice(0) })
      task.promise.then((doc: any) => { if (!dead) searchDocRef.current = doc }).catch(() => {})
    } catch { /* ignore */ }
    return () => { dead = true; searchDocRef.current = null }
  }, [data])

  /* 全文搜索：getTextContent 逐页匹配 → 结果（页码+上下文） */
  const runSearch = async () => {
    const q = query.trim()
    if (!q || !searchDocRef.current) return
    setSearching(true)
    try {
      const doc = searchDocRef.current
      const out: { page: number; ctx: string }[] = []
      for (let p = 1; p <= doc.numPages; p++) {
        const pg = await doc.getPage(p)
        const tc = await pg.getTextContent()
        const text = tc.items.map((it: any) => it.str || '').join('')
        const lower = text.toLowerCase()
        let idx = lower.indexOf(q.toLowerCase())
        while (idx >= 0 && out.length < 200) {
          out.push({ page: p, ctx: text.slice(Math.max(0, idx - 24), idx + q.length + 24) })
          idx = lower.indexOf(q.toLowerCase(), idx + q.length)
        }
        if (out.length >= 200) break
      }
      setHits(out)
      setHitIdx(out.length ? 0 : -1)
      if (out.length) setPage(out[0].page)
    } finally { setSearching(false) }
  }

  /* 当前页文本层高亮（CSS Highlight API——不污染 DOM，Electron 130 支持） */
  useEffect(() => {
    const q = query.trim()
    try { (CSS as any).highlights?.delete('pdf-hits') } catch { /* ignore */ }
    if (!q || !('highlights' in CSS)) return
    const root = canvasBoxRef.current
    if (!root) return
    const spans = root.querySelectorAll('.textLayer span')
    const ranges: Range[] = []
    spans.forEach((sp) => {
      const txt = (sp.textContent || '')
      const lower = txt.toLowerCase()
      let i = lower.indexOf(q.toLowerCase())
      while (i >= 0) {
        try {
          const r = document.createRange()
          r.setStart(sp.firstChild || sp, i)
          r.setEnd(sp.firstChild || sp, i + q.length)
          ranges.push(r)
        } catch { /* ignore */ }
        i = lower.indexOf(q.toLowerCase(), i + q.length)
      }
    })
    if (ranges.length && (CSS as any).highlights) {
      ;(CSS as any).highlights.set('pdf-hits', new (window as any).Highlight(...ranges))
    }
  }, [query, page, loaded, renderScale])

  /* 标记渲染（当前页 marks 的比例坐标 → 绝对定位） */
  const pageMarks = (marks || []).filter((m) => m.page === page)
  const boxW = canvasBoxRef.current?.clientWidth || 0
  const boxH = canvasBoxRef.current?.clientHeight || 0
  const renderMarks = boxW > 0 && boxH > 0

  /* 选区 → 「标记」按钮（mouseup 检测选区在文本层内） */
  const onMouseUp = () => {
    const sel = window.getSelection()
    const txt = sel?.toString().trim() || ''
    const cv = canvasBoxRef.current
    if (!txt || !sel || !cv || sel.rangeCount === 0) { setMarkBtn(null); return }
    const range = sel.getRangeAt(0)
    const cr = range.getBoundingClientRect()
    const box = cv.getBoundingClientRect()
    if (cr.width < 2 || cr.bottom < box.top || cr.top > box.bottom) { setMarkBtn(null); return }
    const rects = Array.from(range.getClientRects())
      .filter((r) => r.width > 2 && r.height > 2)
      .map((r) => ({
        l: (r.left - box.left) / box.width,
        t: (r.top - box.top) / box.height,
        w: r.width / box.width,
        h: r.height / box.height,
      }))
    if (rects.length) setMarkBtn({ x: cr.left - box.left, y: cr.top - box.top - 34, rects, text: txt.slice(0, 300) })
  }

  const saveMark = () => {
    if (!markBtn || !onSaveMarks) { setMarkBtn(null); return }
    const m: PdfMark = {
      id: 'mk-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5),
      page,
      text: markBtn.text,
      rects: markBtn.rects,
    }
    onSaveMarks([...(marks || []), m])
    setMarkBtn(null)
    try { window.getSelection()?.removeAllRanges() } catch { /* ignore */ }
  }

  /* 双层丝滑缩放（视觉层即时 + 渲染层防抖） */
  const setZoom = (z: number, fit = false) => {
    const t = Math.min(4, Math.max(0.25, z))
    if (fit) setFitWidth(true)
    else setFitWidth(false)
    setViewScale(t)
    window.clearTimeout(commitT.current)
    commitT.current = window.setTimeout(() => setRenderScale(t), 250)
  }

  const go = (n: number) => {
    const t = Math.min(numPages, Math.max(1, n))
    setPage(t)
    setPageInput(String(t))
  }

  const onWheel = (e: React.WheelEvent) => {
    if (!e.ctrlKey) return
    e.preventDefault()
    const cur = fitWidth ? fitScale : viewScale
    setZoom(cur * (e.deltaY < 0 ? 1.1 : 0.9))
  }

  const fitScale = useMemo(() => {
    const w = Math.max(240, wrapW - 32)
    return Math.min(4, Math.max(0.25, w / baseWRef.current))
  }, [wrapW])

  useEffect(() => {
    const el = wrapRef.current
    if (!el) return
    const ob = new ResizeObserver((es) => {
      for (const en of es) setWrapW(en.contentRect.width)
    })
    ob.observe(el)
    return () => ob.disconnect()
  }, [])

  const viewW = fitWidth ? Math.max(240, wrapW - 32) : baseWRef.current * viewScale
  const viewH = baseHRef.current * (viewW / baseWRef.current)

  return (
    <div className={'pv-root' + (dark ? ' pv-root--dark' : '')} onWheel={onWheel} onMouseUp={onMouseUp}>
      <div className="pv-bar">
        <button className="pv-btn" disabled={page <= 1} onClick={() => go(page - 1)} title="上一页">‹</button>
        <input
          className="pv-page-input"
          value={pageInput}
          onChange={(e) => setPageInput(e.target.value.replace(/[^\d]/g, ''))}
          onKeyDown={(e) => { if (e.key === 'Enter') go(Number(pageInput) || 1) }}
          onBlur={() => go(Number(pageInput) || 1)}
          title="页码（回车跳转）"
        />
        <span className="pv-page">/ {numPages || '…'}</span>
        <button className="pv-btn" disabled={page >= numPages} onClick={() => go(page + 1)} title="下一页">›</button>
        <span className="pv-sep" />
        <button className="pv-btn" title="缩小" onClick={() => setZoom((fitWidth ? fitScale : viewScale) * 0.85)}>−</button>
        <button className="pv-btn" title="实际大小（100%）" onClick={() => setZoom(1)}>
          {Math.round((fitWidth ? fitScale : viewScale) * 100)}%
        </button>
        <button className="pv-btn" title="放大" onClick={() => setZoom((fitWidth ? fitScale : viewScale) * 1.15)}>＋</button>
        <button className={'pv-btn' + (fitWidth ? ' pv-btn--on' : '')} title="适应宽度" onClick={() => setZoom(fitScale, true)}>⇔</button>
        <span className="pv-sep" />
        {/* ★ 识别+定位：全文搜索 */}
        <input
          className="pv-page-input pv-search-input"
          style={{ width: 110 }}
          placeholder="全文搜索…"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => { if (e.key === 'Enter') runSearch() }}
          title="全文搜索（回车执行）"
        />
        <button className="pv-btn" title="搜索" onClick={runSearch}>🔍</button>
        {hits.length > 0 && (
          <>
            <span className="pv-page">{hitIdx + 1}/{hits.length}</span>
            <button className="pv-btn" title="上一处" onClick={() => { const i = (hitIdx - 1 + hits.length) % hits.length; setHitIdx(i); setPage(hits[i].page) }}>◀</button>
            <button className="pv-btn" title="下一处" onClick={() => { const i = (hitIdx + 1) % hits.length; setHitIdx(i); setPage(hits[i].page) }}>▶</button>
          </>
        )}
        {searching && <span className="pv-page">搜索中…</span>}
        <span style={{ flex: 1 }} />
        {/* ★ 标记列表 */}
        {(marks || []).length > 0 && (
          <button className={'pv-btn' + (showMarkList ? ' pv-btn--on' : '')} title="标记列表" onClick={() => setShowMarkList((v) => !v)}>
            标记 {marks!.length}
          </button>
        )}
      </div>
      <div className={'pv-scroll' + (dark ? ' pv-scroll--dark' : '')} ref={wrapRef} style={{ position: 'relative' }}>
        {/* 标记列表面板 */}
        {showMarkList && (
          <div className="pv-marklist">
            <div className="pv-marklist-head">标记（{(marks || []).length}）<button className="pv-btn" style={{ marginLeft: 'auto' }} onClick={() => setShowMarkList(false)}>×</button></div>
            {(marks || []).length === 0 && <div className="pv-marklist-empty">选中文本后点「标记」按钮创建</div>}
            {(marks || []).map((m) => (
              <div key={m.id} className="pv-markitem" onClick={() => { setPage(m.page); setPageInput(String(m.page)) }} title="点击跳转">
                <span className="pv-markitem-page">P{m.page}</span>
                <span className="pv-markitem-text">{m.text.slice(0, 60)}</span>
                <button className="pv-markitem-del" title="删除标记" onClick={(e) => { e.stopPropagation(); onSaveMarks?.((marks || []).filter((x) => x.id !== m.id)) }}>×</button>
              </div>
            ))}
          </div>
        )}
        {err && <div className="pv-err">PDF 渲染失败：{err}</div>}
        {!err && !loaded && <div className="pv-loading">PDF 加载中…</div>}
        {!err && (
          <Document
            file={file}
            onLoadSuccess={({ numPages: n }) => { setNumPages(n); setLoaded(true) }}
            onLoadError={(e) => setErr(String(e?.message ?? e).slice(0, 200))}
            loading={null}
            error={null}
          >
            <div
              className="pv-page-box"
              ref={canvasBoxRef}
              style={{ width: viewW, height: viewH, transition: 'width 0.18s ease, height 0.18s ease', position: 'relative' }}
            >
              {/* 标记高亮层（比例坐标还原） */}
              {renderMarks && pageMarks.map((m) =>
                m.rects.map((r, ri) => (
                  <div
                    key={m.id + '-' + ri}
                    className="pv-mark-hl"
                    style={{ left: r.l * boxW, top: r.t * boxH, width: r.w * boxW, height: r.h * boxH }}
                    title={m.text.slice(0, 80)}
                  />
                ))
              )}
              {/* 选区标记按钮 */}
              {markBtn && (
                <button
                  className="pv-markbtn"
                  style={{ left: Math.max(4, markBtn.x), top: Math.max(4, markBtn.y) }}
                  onClick={saveMark}
                  title="高亮标记此处"
                >
                  标记
                </button>
              )}
              <Page
                pageNumber={page}
                scale={renderScale}
                renderTextLayer
                renderAnnotationLayer={false}
                onLoadSuccess={(pg: any) => {
                  const vp = pg.getViewport({ scale: 1 })
                  baseWRef.current = vp.width
                  baseHRef.current = vp.height
                }}
                loading={null}
              />
            </div>
          </Document>
        )}
      </div>
    </div>
  )
}
