import { useEffect, useLayoutEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import { useRunBarSnapshot, type RunBarItem } from '../core/nodeRunStatus'
import { useBaseStore } from '../core/store'

/**
 * 画布底部居中运行状态条（2026-11-09 用户指令终版：
 * 「进入的时候从左侧渐进 退出的时候往右侧滑动渐隐」+「增加就重算边界 另一个让路」）
 *
 * - Collapse 三态：enter（列宽 0fr + 从左 -14px）→ open（1fr + 归位）→ exit（收缩 0fr + 向右 +14px 渐隐）
 * - 宽度用 grid 0fr↔1fr 过渡（对实际宽度做动画，无 max-width 死区），其余条目平滑让路、整行居中
 * - 主条 2.6s/6.5s 停留后 exit 收缩卸载；次条完成进幽灵队列 exit 后卸载
 * - 数据源：nodeRunStatus store；主节点由 runPdfMerge/runNodeToPdf 入口 setPrimaryRunNode 标记
 */

type CollapseState = 'enter' | 'open' | 'exit'

/* ★ 2026-11-09 用户「渐进入的明显一点」：位移 80px + 0.38s ease-out（滑入减速感），
   opacity 过渡 0.45s 拖尾更长——先看到滑入，渐显收尾更明显 */
const COLLAPSE_MAP: Record<CollapseState, { cols: string; x: number; o: number }> = {
  enter: { cols: '0fr', x: -80, o: 0 },
  open: { cols: '1fr', x: 0, o: 1 },
  exit: { cols: '0fr', x: 80, o: 0 },
}

function Collapse({ state, fade = 1, children }: { state: CollapseState; fade?: number; children: ReactNode }) {
  const s = COLLAPSE_MAP[state]
  return (
    <span
      style={{
        display: 'grid',
        gridTemplateColumns: s.cols,
        opacity: state === 'open' ? fade : s.o,
        transition: 'grid-template-columns .38s cubic-bezier(.22,1,.36,1), opacity .45s ease',
        pointerEvents: 'none',
      }}
    >
      <span
        style={{
          overflow: 'hidden', minWidth: 0, display: 'flex', justifyContent: 'flex-start',
          transform: `translateX(${s.x}px)`,
          transition: 'transform .38s cubic-bezier(.22,1,.36,1)',
        }}
      >
        {children}
      </span>
    </span>
  )
}

/** 挂载首帧 enter，双 rAF 后进入 open（保证过渡有起点） */
function useSlideIn(): boolean {
  const [open, setOpen] = useState(false)
  useLayoutEffect(() => {
    let id2 = 0
    const id1 = requestAnimationFrame(() => { id2 = requestAnimationFrame(() => setOpen(true)) })
    return () => { cancelAnimationFrame(id1); cancelAnimationFrame(id2) }
  }, [])
  return open
}

export function RunStatusBar() {
  const snap = useRunBarSnapshot()
  const [primaryGone, setPrimaryGone] = useState(true)
  const [primaryExit, setPrimaryExit] = useState(false)
  const [ghosts, setGhosts] = useState<RunBarItem[]>([])
  const prevSec = useRef<RunBarItem[]>([])

  const primary = snap.primary

  useEffect(() => {
    if (!primary) { setPrimaryGone(true); setPrimaryExit(false); return }
    setPrimaryGone(false)
    setPrimaryExit(false)
    if (primary.state === 'running') return
    const ttl = primary.state === 'ok' ? 2600 : 6500
    const t1 = window.setTimeout(() => setPrimaryExit(true), ttl)
    const t2 = window.setTimeout(() => setPrimaryGone(true), ttl + 420)
    return () => { window.clearTimeout(t1); window.clearTimeout(t2) }
  }, [primary?.id, primary?.state])

  useEffect(() => {
    const cur = snap.secondary
    const gone = prevSec.current.filter((p) => !cur.some((c) => c.id === p.id))
    prevSec.current = cur
    if (!gone.length) return
    setGhosts((g) => [...g.filter((x) => !gone.some((gg) => gg.id === x.id)), ...gone])
    const t = window.setTimeout(() => setGhosts((g) => g.filter((x) => !gone.some((gg) => gg.id === x.id))), 420)
    return () => window.clearTimeout(t)
  }, [snap.secondary])

  /* ★ 只订阅稳定引用（tables）——selector 返回新函数会触发 React #185 无限重渲染（实测） */
  const tables = useBaseStore((s) => s.tables)
  const nameOfId = (id: string): string => {
    const t = tables[id]
    if (!t) return ''
    const no = String((t.config as { nodeNo?: string } | undefined)?.nodeNo ?? '')
    return (no ? no + ' ' : '') + t.name
  }
  const nameOf = primary ? nameOfId(primary.id) : ''

  const showPrimary = !!primary && !primaryGone
  if (!showPrimary && snap.secondary.length === 0 && ghosts.length === 0) return null
  const st = primary?.state ?? 'running'

  const itemBase: CSSProperties = {
    display: 'inline-flex', alignItems: 'center', gap: 7,
    padding: '7px 13px',
    background: 'var(--ed-bg2, var(--c-bg1))',
    border: '1px solid var(--ed-border, var(--c-border))',
    borderRadius: 8, boxShadow: '0 6px 20px rgba(0,0,0,0.45)',
    fontSize: 12.5, color: 'var(--ed-text2, var(--c-text2))', whiteSpace: 'nowrap',
  }
  const dot = (color: string, glow?: boolean): CSSProperties => ({
    width: 9, height: 9, borderRadius: '50%', flexShrink: 0, background: color,
    ...(glow ? { boxShadow: `0 0 6px ${color}` } : {}),
  })
  const spinDot = (size: number): CSSProperties => ({
    width: size, height: size, borderRadius: '50%', flexShrink: 0,
    border: '2px solid var(--c-blue, #3b82f6)', borderTopColor: 'transparent',
    animation: 'rf-round-spin 1.1s linear infinite',
  })

  return (
    <div
      style={{
        position: 'absolute', bottom: 24, left: '50%', transform: 'translateX(-50%)',
        zIndex: 210, pointerEvents: 'none',
        display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6,
        maxWidth: '86%',
      }}
      data-testid="run-status-bar"
    >
      {/* 上层：主进程（合并的执行与结果，大号条目，从左滑入/向右滑出） */}
      {showPrimary && primary && (
        <PrimaryCollapse exitFlag={primaryExit}>
          <div style={itemBase}>
          {st === 'running' && <span style={spinDot(9)} />}
          {st === 'ok' && (
            /* ★ 2026-11-09 用户指令「最后的成功了 来个动态的打钩」：SVG 描边动画——圆圈画出→勾划出 */
            <svg
              className="runbar-check"
              width={18} height={18} viewBox="0 0 24 24"
              style={{ flexShrink: 0, filter: 'drop-shadow(0 0 4px rgba(34,197,94,.45))' }}
            >
              <circle cx={12} cy={12} r={10} fill="rgba(34,197,94,.12)" stroke="var(--c-green, #22c55e)" strokeWidth={2.2} />
              <path d="M7 12.5 L10.5 16 L17 9" fill="none" stroke="var(--c-green, #22c55e)" strokeWidth={2.4} strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          )}
          {st === 'error' && (
            /* ★ 2026-11-09 用户指令「失败 超时等同等匹配的效果」：与打钩对称——红圈画出→两笔叉划出 */
            <svg
              className="runbar-cross"
              width={18} height={18} viewBox="0 0 24 24"
              style={{ flexShrink: 0, filter: 'drop-shadow(0 0 4px rgba(239,68,68,.45))' }}
            >
              <circle cx={12} cy={12} r={10} fill="rgba(239,68,68,.12)" stroke="var(--c-red, #ef4444)" strokeWidth={2.2} />
              <path className="l1" d="M9 9 L15 15" fill="none" stroke="var(--c-red, #ef4444)" strokeWidth={2.4} strokeLinecap="round" />
              <path className="l2" d="M15 9 L9 15" fill="none" stroke="var(--c-red, #ef4444)" strokeWidth={2.4} strokeLinecap="round" />
            </svg>
          )}
            <span>
              {nameOf ? <b style={{ color: 'var(--ed-text1, var(--c-text1))', marginRight: 6 }}>{nameOf}</b> : null}
              {primary.label}
            </span>
          </div>
        </PrimaryCollapse>
      )}
      {/* 下层：次进程（上游子任务）左右挨着居中；完成进幽灵队列向右滑出后卸载 */}
      {(snap.secondary.length > 0 || ghosts.length > 0) && (
        <div style={{ display: 'flex', flexDirection: 'row', alignItems: 'center', justifyContent: 'center', gap: 6, overflow: 'hidden' }}>
          {snap.secondary.map((it) => (
            <SecondaryItem key={it.id} item={it} itemBase={itemBase} nameOfId={nameOfId} spinDot={spinDot} />
          ))}
          {ghosts.map((it) => (
            <GhostItem key={'g-' + it.id} item={it} itemBase={itemBase} nameOfId={nameOfId} dot={dot} onDone={() => setGhosts((g) => g.filter((x) => x.id !== it.id))} />
          ))}
        </div>
      )}
    </div>
  )
}

/** 主条目三态：挂载 enter → 双 rAF open；exitFlag=true 时 exit（向右滑出收缩） */
function PrimaryCollapse({ exitFlag, children }: { exitFlag: boolean; children: ReactNode }) {
  const entered = useSlideIn()
  const state: CollapseState = exitFlag ? 'exit' : entered ? 'open' : 'enter'
  return <Collapse state={state}>{children}</Collapse>
}

function SecondaryItem({ item, itemBase, nameOfId, spinDot }: {
  item: RunBarItem
  itemBase: CSSProperties
  nameOfId: (id: string) => string
  spinDot: (size: number) => CSSProperties
}) {
  const open = useSlideIn()
  const name = nameOfId(item.id)
  return (
    <Collapse state={open ? 'open' : 'enter'} fade={0.92}>
      <span style={{ ...itemBase, fontSize: 11.5, padding: '5px 10px' }}>
        <span style={spinDot(7)} />
        {name ? <b style={{ color: 'var(--ed-text1, var(--c-text1))', marginRight: 5 }}>{name}</b> : null}
        {item.label}
      </span>
    </Collapse>
  )
}

/** 完成的次条目：挂载即 open（视觉无缝衔接），60ms 后 exit（向右滑出收缩），播完卸载 */
function GhostItem({ item, itemBase, nameOfId, dot, onDone }: {
  item: RunBarItem
  itemBase: CSSProperties
  nameOfId: (id: string) => string
  dot: (color: string, glow?: boolean) => CSSProperties
  onDone: () => void
}) {
  const toExit = useRef(onDone)
  toExit.current = onDone
  const [exiting, setExiting] = useState(false)
  /* ★ 60ms 后才开始 exit（此时已 open 渲染过一帧，过渡有起点）；
     onDone 必须等收缩过渡播完（60+360ms）再调——提前调会把条目直接卸载，滑出动画被腰斩（实测）。 */
  useLayoutEffect(() => {
    const id1 = window.setTimeout(() => setExiting(true), 60)
    const id2 = window.setTimeout(() => toExit.current(), 60 + 420)
    return () => { window.clearTimeout(id1); window.clearTimeout(id2) }
  }, [])
  const name = nameOfId(item.id)
  return (
    <Collapse state={exiting ? 'exit' : 'open'} fade={0.92}>
      <span style={{ ...itemBase, fontSize: 11.5, padding: '5px 10px' }}>
        <span style={dot('var(--c-green, #22c55e)')} />
        {name ? <b style={{ color: 'var(--ed-text1, var(--c-text1))', marginRight: 5 }}>{name}</b> : null}
        {item.label}
      </span>
    </Collapse>
  )
}
