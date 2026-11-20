/* ★ 2026-11-01 连线显示效果选择器 v2（用户指令）：
 *  —— 圆形按钮，右下角与「打开对话框」按钮同一横线（bottom 27）
 *  —— 修黑主题下拉文字不可见：不再用半定义的 CSS 变量，双主题配色自检
 *  菜单从按钮向上弹出；原效果 + archify 预设 + 关闭，localStorage 持久化 */
import { useEffect, useRef, useState } from 'react'

export type EdgeFx = 'pause' | 'always' | 'off' | 'pulse' | 'blueprint' | 'crawl' | 'breathe' | 'echo'

const ITEMS: Array<{ v: EdgeFx; label: string; desc: string }> = [
  { v: 'pause', label: '原效果', desc: '选中/拖动时流动' },
  { v: 'always', label: '原效果（始终流动）', desc: '连线持续流动' },
  { v: 'pulse', label: '脉冲信号流', desc: '亮点沿边扫过 · 带残影' },
  { v: 'blueprint', label: '蓝图轨迹', desc: '亮点慢扫留轨迹 · 工程图纸感' },
  { v: 'crawl', label: '亮色爬行', desc: '虚线匀速爬行 · 常态数据流' },
  { v: 'breathe', label: '呼吸暗线', desc: '虚线周期明暗 · 待机概览' },
  { v: 'echo', label: '回声流水', desc: '多条边依次亮点 · 显示先后' },
  { v: 'off', label: '关闭效果', desc: '静态连线' },
]

export function EdgeFxPicker({ value, onChange }: { value: EdgeFx; onChange: (v: EdgeFx) => void }) {
  const [open, setOpen] = useState(false)
  const [isLight, setIsLight] = useState(document.documentElement.getAttribute('data-theme') === 'light')
  const wrapRef = useRef<HTMLDivElement | null>(null)

  useEffect(() => {
    const onTheme = () => setIsLight(document.documentElement.getAttribute('data-theme') === 'light')
    window.addEventListener('app-theme-change', onTheme)
    const obs = new MutationObserver(onTheme)
    obs.observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] })
    return () => {
      window.removeEventListener('app-theme-change', onTheme)
      obs.disconnect()
    }
  }, [])
  /* 点外部关闭 */
  useEffect(() => {
    if (!open) return
    const onDown = (e: MouseEvent) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) setOpen(false)
    }
    /* ★ 捕获阶段监听：React Flow 画布的 pointer 拖拽不冒泡也能收到，点其他位置必关 */
    document.addEventListener('mousedown', onDown, true)
    document.addEventListener('pointerdown', onDown, true)
    return () => {
      document.removeEventListener('mousedown', onDown, true)
      document.removeEventListener('pointerdown', onDown, true)
    }
  }, [open])

  /* ★ 动态贴缩略图左侧：量取 .react-flow__minimap 实际位置（不再猜像素） */
  const [mmPos, setMmPos] = useState<{ left: number; bottom: number } | null>(null)
  useEffect(() => {
    const place = () => {
      const mm = document.querySelector('.react-flow__minimap')
      if (!mm) return
      const r = mm.getBoundingClientRect()
      if (r.width < 10) return
      setMmPos({ left: Math.max(8, r.left - 46), bottom: Math.max(16, window.innerHeight - r.bottom) })
    }
    place()
    const t1 = setTimeout(place, 600)
    const t2 = setTimeout(place, 1800)
    window.addEventListener('resize', place)
    return () => { clearTimeout(t1); clearTimeout(t2); window.removeEventListener('resize', place) }
  }, [])

  /* 双主题配色 */
  const T = isLight
    ? { btnBg: '#ffffff', btnBorder: '#c5ccd6', btnColor: '#5a6572', menuBg: '#ffffff', menuBorder: '#d5dbe3', itemText: '#333a42', itemSub: '#8a94a2', itemHover: 'rgba(120,140,180,0.12)', itemActive: 'rgba(77,125,253,0.12)', shadow: '0 10px 32px rgba(30,40,60,0.2)' }
    : { btnBg: 'rgba(40,44,52,0.92)', btnBorder: 'rgba(140,160,190,0.4)', btnColor: '#c8d4e6', menuBg: 'rgba(40,44,52,0.98)', menuBorder: 'rgba(140,160,190,0.35)', itemText: '#dfe7f3', itemSub: '#8fa0bd', itemHover: 'rgba(140,160,190,0.14)', itemActive: 'rgba(77,125,253,0.2)', shadow: '0 10px 32px rgba(0,0,0,0.5)' }

  return (
    <div
      ref={wrapRef}
      style={{
        position: 'fixed', zIndex: 260, /* ★ 缩略图正左侧（动态量取缩略图实际位置） */
        ...(mmPos ? { left: mmPos.left, bottom: mmPos.bottom } : { right: 300, bottom: 24 }),
        display: 'flex', flexDirection: 'column', alignItems: 'flex-end',
      }}
      onClick={(e) => e.stopPropagation()}
    >
      {open && (
        <div
          style={{
            /* ★ 菜单在按钮左侧弹出（不遮缩略图）；贴底向上展开 */
            position: 'absolute', right: 44, bottom: 0, width: 236,
            background: T.menuBg, border: `1px solid ${T.menuBorder}`,
            borderRadius: 12, boxShadow: T.shadow, padding: 5,
          }}
        >
          <div style={{ fontSize: 11, color: T.itemSub, padding: '4px 10px 6px' }}>连线显示效果</div>
          {ITEMS.map((it) => (
            <div
              key={it.v}
              onClick={() => { onChange(it.v); setOpen(false) }}
              style={{
                padding: '7px 10px', borderRadius: 8, cursor: 'pointer', fontSize: 13,
                color: T.itemText,
                background: value === it.v ? T.itemActive : 'transparent',
                fontWeight: value === it.v ? 600 : 400,
              }}
              onMouseEnter={(e) => { (e.target as HTMLElement).style.background = value === it.v ? T.itemActive : T.itemHover }}
              onMouseLeave={(e) => { (e.target as HTMLElement).style.background = value === it.v ? T.itemActive : 'transparent' }}
            >
              <div>{it.label}{value === it.v ? ' ✓' : ''}</div>
              <div style={{ fontSize: 10.5, opacity: 0.6, marginTop: 1 }}>{it.desc}</div>
            </div>
          ))}
        </div>
      )}
      {/* 圆形按钮（与对话框展开钮同一横线） */}
      <button
        title="连线显示效果"
        onClick={() => setOpen((o) => !o)}
        style={{
          width: 34, height: 34, borderRadius: '50%', cursor: 'pointer',
          background: open ? T.itemActive : T.btnBg,
          border: `1px solid ${open ? 'rgba(77,125,253,0.55)' : T.btnBorder}`,
          color: T.btnColor, fontSize: 15,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          boxShadow: open ? T.shadow : '0 4px 16px rgba(30,40,60,0.16)',
          transition: 'background 0.15s, box-shadow 0.15s, border-color 0.15s',
        }}
      >⚡</button>
    </div>
  )
}
