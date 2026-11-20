/**
 * PDF 专业渲染查看器（MuPDF wasm，2026-10-24 用户指令「专业渲染引擎」第四入口）：
 * - MuPDF = Foxit 旗下专业引擎（AA 渲染业界第一梯队，Firefox/pdf.js 之外最优开源内核）
 * - canvas 直绘（getPixels → ImageData → putImageData），缩放 2x 超采样保清晰
 * - 工具条：翻页/页码/缩放 ±/整页适应（复用 pv-* 样式语言）
 */
import { useEffect, useRef, useState } from 'react'
import * as mupdf from 'mupdf'

export function PdfMupdf ({ data, dark }: { data: Uint8Array; dark: boolean }) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const docRef = useRef<mupdf.PDFDocument | null>(null)
  const [page, setPage] = useState(1)
  const [pageInput, setPageInput] = useState('1')
  const [numPages, setNumPages] = useState(0)
  const [zoom, setZoom] = useState(1)
  const [err, setErr] = useState('')
  /* ★ 2026-10-24 用户反馈：手型拖拽模式（✋ 切换，抓手平移） */
  const [panMode, setPanMode] = useState(true)   /* ★ 默认手型拖拽（用户指令） */
  const scrollRef = useRef<HTMLDivElement | null>(null)
  const panRef = useRef<{ on: boolean; x: number; y: number }>({ on: false, x: 0, y: 0 })

  /* 打开文档（bytes 副本传入 wasm，不污染调用方） */
  useEffect(() => {
    let dead = false
    try {
      const doc = mupdf.Document.openDocument(data.slice(0), 'application/pdf') as mupdf.PDFDocument
      docRef.current = doc
      if (!dead) { setNumPages(doc.countPages()); setPage(1); setPageInput('1') }
    } catch (eO) { if (!dead) setErr(String(eO).slice(0, 140)) }
    return () => { dead = true; docRef.current = null }
  }, [data])

  /* 渲染当前页：2x 超采样 → putImageData（不透明 RGB→RGBA 需 alpha 通道，toPixmap alpha=false 时 pixels 为灰度/RGB——
     MuPDF Pixmap 默认带 alpha（n=4）当 colorspace RGB + alpha=true；这里 alpha=true 保证 RGBA 排布） */
  useEffect(() => {
    const doc = docRef.current
    const cv = canvasRef.current
    if (!doc || !cv || err || numPages <= 0) return
    try {
      const pg = doc.loadPage(page - 1)
      const m = mupdf.Matrix.scale(zoom * 2, zoom * 2)
      const pix = pg.toPixmap(m, mupdf.ColorSpace.DeviceRGB, true, false)
      cv.width = pix.getWidth()
      cv.height = pix.getHeight()
      cv.style.width = Math.round(pix.getWidth() / 2) + 'px'
      cv.style.height = Math.round(pix.getHeight() / 2) + 'px'
      const ctx = cv.getContext('2d')
      if (ctx) ctx.putImageData(new ImageData(new Uint8ClampedArray(pix.getPixels()), pix.getWidth(), pix.getHeight()), 0, 0)
    } catch (eR) { if (!String(eR).includes('dead')) setErr(String(eR).slice(0, 140)) }
  }, [page, zoom, err, numPages])

  const go = (n: number) => {
    const t = Math.min(numPages, Math.max(1, n))
    setPage(t)
    setPageInput(String(t))
    if (scrollRef.current) scrollRef.current.scrollTop = 0  // ★ 翻页回页顶
  }

  /* ★ 2026-10-24 用户指令：Ctrl+滚轮缩放（±10% 步进） */
  const onWheel = (e: React.WheelEvent) => {
    if (!e.ctrlKey) return
    e.preventDefault()
    setZoom((z) => Math.min(4, Math.max(0.25, z * (e.deltaY < 0 ? 1.1 : 0.9))))
  }

  return (
    <div className={'pv-root' + (dark ? ' pv-root--dark' : '')} onWheel={onWheel}>
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
        <button className="pv-btn" title="缩小" onClick={() => setZoom((z) => Math.max(0.25, z / 1.2))}>−</button>
        <button className="pv-btn" title="100%" onClick={() => setZoom(1)}>{Math.round(zoom * 100)}%</button>
        <button className="pv-btn" title="放大" onClick={() => setZoom((z) => Math.min(4, z * 1.2))}>＋</button>
        <button className="pv-btn" title="整页适应" onClick={() => setZoom(1)}>⤢</button>
        <button className="pv-btn" title="手型拖拽（点击切换）" style={panMode ? { borderColor: '#6b8fc9', color: '#6b8fc9' } : undefined} onClick={() => setPanMode((v) => !v)}>✋</button>
      </div>
      <div
        ref={scrollRef}
        className={'pv-scroll' + (dark ? ' pv-scroll--dark' : '')}
        style={{ display: 'flex', justifyContent: 'center', overflow: 'auto', cursor: panMode ? (panRef.current.on ? 'grabbing' : 'grab') : undefined }}
        onMouseDown={(e) => {
          if (!panMode) return
          panRef.current = { on: true, x: e.clientX, y: e.clientY }
          e.preventDefault()
        }}
        onMouseMove={(e) => {
          if (!panMode || !panRef.current.on || !scrollRef.current) return
          scrollRef.current.scrollLeft -= e.clientX - panRef.current.x
          scrollRef.current.scrollTop -= e.clientY - panRef.current.y
          panRef.current.x = e.clientX
          panRef.current.y = e.clientY
        }}
        onMouseUp={() => { panRef.current.on = false }}
        onMouseLeave={() => { panRef.current.on = false }}
      >
        {err && <div className="pv-err">MuPDF 渲染失败：{err}</div>}
        {!err && numPages <= 0 && <div className="pv-loading">MuPDF 加载中…</div>}
        {!err && <canvas ref={canvasRef} style={{ imageRendering: 'auto' }} />}
      </div>
    </div>
  )
}
