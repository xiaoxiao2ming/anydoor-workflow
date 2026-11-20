/**
 * 画布性能监控 HUD（2026-10-24 用户指令「对 JS 执行和 React 渲染管线监控」）：
 * - FPS：rAF 每秒帧计数
 * - 帧耗时：rAF 间隔滑动平均（>16.7ms 即掉帧）
 * - 长任务：PerformanceObserver longtask（>50ms 的主线程阻塞）
 * - JS 堆：performance.memory.usedJSHeapSize（Chromium）
 * - 渲染变更率：onNodesChange 每秒触发次数（React 管线压力）
 * 视图菜单「性能监控」开关控制显隐，点击面板折叠/展开。
 */
import { useEffect, useRef, useState } from 'react'

export function PerfHud ({ changeCount }: { changeCount: number }) {
  const [open, setOpen] = useState(false)
  const [fps, setFps] = useState(60)
  const [frameMs, setFrameMs] = useState(16.7)
  const [heap, setHeap] = useState(0)
  const [longTasks, setLongTasks] = useState(0)
  const [longMs, setLongMs] = useState(0)
  const [cps, setCps] = useState(0)
  const changeRef = useRef(changeCount)
  changeRef.current = changeCount
  /* ★ 2026-10-24 用户指令「测试节点坐标 标记毫秒差异 验证卡顿」：拖动采样器 */
  const [sampling, setSampling] = useState(false)
  const samplesRef = useRef<{ t: number; x: number; y: number }[]>([])
  const samplingRef = useRef(false)
  const maxDtRef = useRef(0)
  const slowRef = useRef(0)
  const [sampleReport, setSampleReport] = useState('')

  /* rAF 循环：FPS + 帧耗时 */
  useEffect(() => {
    let raf = 0
    let last = performance.now()
    let acc = 0
    let frames = 0
    let costAcc = 0
    const loop = (t: number) => {
      const dt = t - last
      last = t
      if (samplingRef.current) {
        const sel = document.querySelector('.react-flow__node.selected') as HTMLElement | null
        if (sel) {
          const r = sel.getBoundingClientRect()
          const arr = samplesRef.current
          const prev = arr[arr.length - 1]
          arr.push({ t, x: Math.round(r.left), y: Math.round(r.top) })
          if (arr.length > 4000) arr.shift()
          if (prev && t - prev.t > 4) {
            const moveD = Math.hypot(r.left - prev.x, r.top - prev.y)
            if (moveD > 0.5) {
              const gap = t - prev.t
              if (gap > maxDtRef.current) maxDtRef.current = gap
              if (gap > 33) slowRef.current++
            }
          }
        }
      }
      frames++
      acc += dt
      if (dt > 16.7) costAcc += dt - 16.7
      if (acc >= 1000) {
        setFps(Math.round((frames * 1000) / acc))
        setFrameMs(Math.round((costAcc / frames) * 10) / 10)
        frames = 0; acc = 0; costAcc = 0
      }
      raf = requestAnimationFrame(loop)
    }
    raf = requestAnimationFrame(loop)
    return () => cancelAnimationFrame(raf)
  }, [])

  /* 长任务监听 + 堆内存 + 变更率采样（1s） */
  useEffect(() => {
    let lt = 0
    let ltm = 0
    let obs: PerformanceObserver | null = null
    try {
      obs = new PerformanceObserver((list) => {
        for (const en of list.getEntries()) { lt++; ltm += en.duration }
      })
      obs.observe({ entryTypes: ['longtask'] })
    } catch { /* 老内核无 longtask */ }
    const iv = setInterval(() => {
      setLongTasks(lt); setLongMs(Math.round(ltm))
      try {
        const mem = (performance as unknown as { memory?: { usedJSHeapSize: number } }).memory
        if (mem) setHeap(Math.round(mem.usedJSHeapSize / 1024 / 1024))
      } catch { /* ignore */ }
      setCps(changeRef.current)
    }, 1000)
    return () => { obs?.disconnect(); clearInterval(iv) }
  }, [])

  const heapColor = heap > 800 ? '#ff6b6b' : heap > 400 ? '#ffb74d' : '#81c784'
  const fpsColor = fps >= 55 ? '#81c784' : fps >= 30 ? '#ffb74d' : '#ff6b6b'

  if (!open) {
    return (
      <div
        className="perf-hud perf-hud--mini"
        title="画布性能监控（点击展开）"
        onClick={() => setOpen(true)}
      >
        <span style={{ color: fpsColor }}>{fps} FPS</span>
      </div>
    )
  }
  return (
    <div className="perf-hud" onClick={(e) => { if ((e.target as HTMLElement).closest('.perf-hud-btn')) return; setOpen(false) }} title="点击折叠">
      <div className="perf-hud-row"><span>FPS</span><span style={{ color: fpsColor }}>{fps}</span></div>
      <div className="perf-hud-row">
        <button
          className="perf-hud-btn"
          onClick={(e) => {
            e.stopPropagation()
            if (samplingRef.current) {
              samplingRef.current = false
              const arr = samplesRef.current
              if (arr.length < 3) { setSampleReport('采样不足——请拖动节点'); return }
              const gaps = []
              for (let i = 1; i < arr.length; i++) {
                const g = arr[i].t - arr[i - 1].t
                if (g > 4 && g < 500) gaps.push(g)
              }
              const avg = Math.round(gaps.reduce((a, b) => a + b, 0) / gaps.length)
              const mx = Math.max(...gaps)
              const over33 = gaps.filter((g) => g > 33).length
              const over100 = gaps.filter((g) => g > 100).length
              const verdict = over100 > 0 ? '明显卡顿' : over33 > 0 ? '偶有掉帧' : '流畅'
              const rep = '拖动采样 ' + arr.length + ' 帧: 平均 ' + avg + 'ms/帧, 最大间隔 ' + mx + 'ms, >33ms ' + over33 + ' 次, >100ms ' + over100 + ' 次 → ' + verdict
              setSampleReport(rep)
              try {
                const w = window as unknown as { pushLog?: (m: string) => void }
                import('../ui/logBus').then((lb) => lb.pushLog('[canvas] ' + rep, over100 > 0 ? 'warn' : 'ok'))
              } catch { /* ignore */ }
              setSampling(false)
            } else {
              samplesRef.current = []
              maxDtRef.current = 0
              slowRef.current = 0
              samplingRef.current = true
              setSampleReport('采样中…请拖动节点')
              setSampling(true)
            }
          }}
        >
          {sampling ? '■ 停止采样' : '● 开始采样'}
        </button>
      </div>
      {sampleReport && <div className="perf-hud-row" style={{ color: '#81c784', whiteSpace: 'normal', maxWidth: 240 }}>{sampleReport}</div>}
      <div className="perf-hud-row"><span>帧超时</span><span style={{ color: frameMs > 8 ? '#ffb74d' : '#81c784' }}>{frameMs} ms/帧</span></div>
      <div className="perf-hud-row"><span>长任务</span><span>{longTasks} 次 / {longMs} ms</span></div>
      <div className="perf-hud-row"><span>JS 堆</span><span style={{ color: heapColor }}>{heap} MB</span></div>
      <div className="perf-hud-row"><span>渲染变更/s</span><span>{cps}</span></div>
    </div>
  )
}
