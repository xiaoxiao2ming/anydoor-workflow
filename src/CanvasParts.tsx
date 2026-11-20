/* Canvas 侧子组件（2026-10-22 拆分自 Canvas.tsx 巨石文件——每个组件独立可测、
   改动隔离；Canvas.tsx 只保留画布编排与事件处理）。 */
import { useRef, useState, useMemo, useEffect, type CSSProperties, type ReactNode } from 'react'
import { createPortal } from 'react-dom'
import { useReactFlow, useStore, useViewport, NodeResizer, useNodeId } from '@xyflow/react'
import { useBaseStore } from './core/store'
import { groupFitRect, isRuleNodeKind, isRoundNodeKind, styleOf } from './core/nodeStyle'
import { IcUnlock, IcLock, IcTrash, IcShrink, IcGrid, IcFold, IcMaximize, IcMinimize, IcMapPin } from './ui/icons'
import { NodeSqBtn } from './ui/NodeBtnBar'
/* 锁定组：主节点 + 被它连上的圆形节点（从 Canvas.tsx 随迁的纯函数） */
function groupMembersOf(regId: string, t: Record<string, { id?: string; kind?: string; config?: { linkedInputs?: string[] } } | undefined>): string[] {
  const reg = t[regId]
  if (!reg || isRoundNodeKind(reg.kind)) return [regId]
  const out = [regId]
  const regLi = (reg.config?.linkedInputs ?? []) as string[]
  for (const n of Object.values(t)) {
    if (!n || isRoundNodeKind(n.kind)) continue
    const li = (n.config?.linkedInputs ?? []) as string[]
    if (li.includes(regId) || regLi.includes(n.id as string)) out.push(n.id as string)
  }
  return out
}

import type { Row } from './core/store'
import type { FieldDef } from './core/schema'
import { pushLog } from './ui/logBus'

export function TestTableDialog({ nodeId, onClose }: { nodeId: string; onClose: () => void }) {
  const t = useBaseStore((s) => s.tables[nodeId])
  const rows = useBaseStore((s) => s.records[nodeId]) as Row[] | undefined
  const fields = t?.fields ?? []
  const write = (nextFields: typeof fields, nextRows: Row[]) =>
    useBaseStore.getState().setTableData(nodeId, nextFields, nextRows)
  const renameField = (i: number, newName: string) => {
    const old = fields[i]
    if (!old || !newName.trim()) return
    const nm = newName.trim()
    const f2 = fields.map((f, j) => (j === i ? { ...f, name: nm } : f))
    const r2 = (rows ?? []).map((r) => {
      const c: Row = {}
      Object.keys(r).forEach((k) => { c[k === old.name ? nm : k] = r[k] })
      return c
    })
    write(f2, r2)
  }
  const setType = (i: number, type: 'text' | 'number') => {
    const f = fields[i]
    if (!f) return
    const f2 = fields.map((x, j) => (j === i ? { ...x, type } : x))
    const r2 = (rows ?? []).map((r) => {
      const c = { ...r }
      const v = c[f.name]
      c[f.name] = type === 'number' ? (isNaN(parseFloat(String(v ?? ''))) ? 0 : parseFloat(String(v))) : String(v ?? '')
      return c
    })
    write(f2, r2)
  }
  const setValue = (i: number, v: string) => {
    const f = fields[i]
    if (!f) return
    const val = f.type === 'number' ? (v === '' ? 0 : (parseFloat(v) || 0)) : v
    const base = (rows ?? []).length ? rows! : [{}]
    const r2 = base.map((r, j) => (j === 0 ? { ...r, [f.name]: val } : r))
    write(fields, r2)
  }
  const addField = () => {
    let n = fields.length + 1
    let nm = '参数' + n
    while (fields.some((f) => f.name === nm)) { n++; nm = '参数' + n }
    write(
      [...fields, { id: nm, name: nm, type: 'text' } as FieldDef],
      [...(rows ?? []), {}],
    )
  }
  const delField = (i: number) => {
    const f = fields[i]
    if (!f) return
    const f2 = fields.filter((_, j) => j !== i)
    const r2 = (rows ?? []).map((r) => {
      const c = { ...r }
      delete c[f.name]
      return c
    })
    write(f2, r2)
  }
  const row0: Row = rows?.[0] ?? {}
  const cellStyle: CSSProperties = {
    width: '100%', boxSizing: 'border-box', padding: '4px 6px',
    background: '#0d1117', color: '#e8edf2', border: '1px solid var(--c-border)',
    borderRadius: 4, fontSize: 12, outline: 'none',
  }
  return createPortal(
    <div
      style={{
        position: 'fixed', inset: 0, zIndex: 300, background: 'rgba(1,4,9,.6)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
      }}
      onClick={onClose}
    >
      <div
        style={{
          width: 470, maxWidth: '92vw', boxSizing: 'border-box',
          background: 'var(--c-bg1)', color: '#e8edf2', border: '1px solid var(--c-border)',
          borderRadius: 10, boxShadow: '0 16px 44px rgba(0,0,0,.45)', padding: '14px 16px',
          font: '12.5px/1.6 "Microsoft YaHei",system-ui,sans-serif',
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
          <b style={{ fontSize: 13 }}>数据表节点（测试）· 参数编辑{t ? '｜' + t.name : ''}</b>
          <button
            style={{ background: 'transparent', border: 0, color: 'var(--c-text3)', cursor: 'pointer', fontSize: 15 }}
            title="关闭"
            onClick={onClose}
          >✕</button>
        </div>
        <div style={{ display: 'flex', gap: 8, padding: '0 2px 4px', color: 'var(--c-text3)', fontSize: 11 }}>
          <span style={{ flex: 1.2 }}>参数名</span>
          <span style={{ flex: 0.9 }}>数据格式</span>
          <span style={{ flex: 1 }}>值（第一行）</span>
          <span style={{ width: 22 }} />
        </div>
        {fields.map((f, i) => (
          <div key={f.id} style={{ display: 'flex', gap: 8, marginBottom: 6, alignItems: 'center' }}>
            <input
              style={{ ...cellStyle, flex: 1.2 }}
              value={f.name}
              onChange={(e) => renameField(i, e.target.value)}
            />
            <select
              style={{ ...cellStyle, flex: 0.9 }}
              value={f.type === 'number' ? 'number' : 'text'}
              onChange={(e) => setType(i, e.target.value === 'number' ? 'number' : 'text')}
            >
              <option value="text">文本</option>
              <option value="number">数字</option>
            </select>
            <input
              style={{ ...cellStyle, flex: 1 }}
              type={f.type === 'number' ? 'number' : 'text'}
              value={String(row0[f.name] ?? '')}
              onChange={(e) => setValue(i, e.target.value)}
            />
            <button
              style={{ width: 22, height: 22, background: 'transparent', border: '1px solid var(--c-border)', borderRadius: 4, color: '#f85149', cursor: 'pointer' }}
              title="删除该参数"
              onClick={() => delField(i)}
            >✕</button>
          </div>
        ))}
        <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 10 }}>
          <button
            style={{ padding: '5px 12px', background: 'var(--c-line)', color: '#e8edf2', border: '1px solid var(--c-border)', borderRadius: 6, cursor: 'pointer', fontSize: 12 }}
            onClick={addField}
          >＋ 添加参数</button>
          <button
            style={{ padding: '5px 14px', background: '#1f6feb', color: '#fff', border: 0, borderRadius: 6, cursor: 'pointer', fontSize: 12 }}
            onClick={onClose}
          >完成</button>
        </div>
        <div style={{ marginTop: 8, color: 'var(--c-text4)', fontSize: 11 }}>
          即改即存（随画布保存）；OO Excel 引用面板点选参数名即可引入。
        </div>
      </div>
    </div>,
    document.body,
  )
}

export function GrpLockButton({ nodeId }: { nodeId: string }) {
  const allTables = useBaseStore((s) => s.tables)
  const members = useMemo(() => {
    const t0 = allTables[nodeId]
    if (!t0 || isRoundNodeKind(t0.kind)) return []
    return groupMembersOf(nodeId, allTables).filter((id) => id !== nodeId)
  }, [allTables, nodeId])
  const grpLocked = !!((allTables[nodeId]?.config as { grpLocked?: boolean } | undefined)?.grpLocked)
  const has = members.length > 0
  /* ★ 与 OoNodeButtons/标题栏同语言：31px、线性 SVG currentColor、激活主题色淡底 */
  const lockStyle: CSSProperties = {
    width: 31, height: 31, padding: 0,
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    cursor: 'pointer', borderRadius: 6, flexShrink: 0,
    border: grpLocked ? '1px solid #d29922' : '1px solid var(--c-border)',
    background: grpLocked ? 'rgba(210,153,34,0.18)' : 'transparent',
    color: !has ? 'var(--c-text4)' : grpLocked ? '#d29922' : 'var(--ed-text3, var(--c-text3))',
  }
  return (
    <button
      className="rf-oo-btn"
      style={lockStyle}
      title={has
        ? (grpLocked
          ? '锁定组：开 —— 拖动组内任何节点，整组一起移动（点击解锁）'
          : '锁定组：关 —— 点击锁定本节点与圆形节点的相对位置')
        : '暂无相连的圆形节点 —— 先把 AI / 参数库 / 规则跟随 拖到本节点上连线'}
      onClick={(ev) => {
        ev.stopPropagation()
        if (!has) return
        useBaseStore.getState().setNodeConfig(nodeId, { grpLocked: !grpLocked })
      }}
    >
      {grpLocked ? <IcLock size={19} /> : <IcUnlock size={19} />}
    </button>
  )
}

/* ★ 统一节点锁（2026-10-26 全节点推广，用户备忘「节点锁定统一」）：
   - 数据锁 config.dataLocked：防内容变动（OO 只读打开 / 文件族禁导入/删除/排序/执行）
   - 位置锁 config.posLocked：防拖动（画布 node.draggable=false，全 kind 生效）
   - OO 遗留字段 ooDataLocked/ooPosLocked 读取兼容（任一为真即视为锁定）；
     写入一律走新字段（旧数据首次点击后自然迁移）
   - 样式统一 = NodeSqBtn（pm-sq-btn 31×31 圆角 6，激活 --on 蓝框）——
     与 OO 双钮（OoNodeButtons）同语义同顺序：数据锁在前、位置锁在后 */
export function NodeLockButtons({ nodeId }: { nodeId: string }) {
  const cfg = useBaseStore((s) => s.tables[nodeId]?.config) as
    | { dataLocked?: boolean; posLocked?: boolean; ooDataLocked?: boolean; ooPosLocked?: boolean }
    | undefined
  const dataL = !!(cfg?.dataLocked || cfg?.ooDataLocked)
  const posL = !!(cfg?.posLocked || cfg?.ooPosLocked)
  const tog = (patch: Record<string, boolean>) => useBaseStore.getState().setNodeConfig(nodeId, patch)
  return (
    <>
      <NodeSqBtn
        icon={dataL ? <IcLock size={19} /> : <IcUnlock size={19} />}
        active={dataL}
        title={dataL
          ? '数据已锁定：内容防变动（导入/删除/排序/执行均停用——点击解锁）'
          : '数据锁定：节点内容防变动（OO 文档只读打开；重开保持）'}
        onClick={(ev) => { ev.stopPropagation(); tog({ dataLocked: !dataL }) }}
      />
      <NodeSqBtn
        icon={<IcMapPin size={19} />}
        active={posL}
        title={posL
          ? '位置已锁定：节点不可拖动（点击解锁）'
          : '位置锁定：节点在画布上不可拖动（防误拖；重开保持）'}
        onClick={(ev) => { ev.stopPropagation(); tog({ posLocked: !posL }) }}
      />
    </>
  )
}

/* ★ 对齐追踪线覆盖层（2026-10-21）：flow 坐标 → 屏幕坐标（viewport 变换） */export function AlignmentGuides({
  guides,
}: {
  guides: { v: { x: number; y0: number; y1: number }[]; h: { y: number; x0: number; x1: number }[] }
}) {
  const { x: vx, y: vy, zoom } = useViewport()
  return (
    <svg className="align-guides">
      {guides.v.map((g, i) => (
        <line key={'v' + i} x1={g.x * zoom + vx} y1={g.y0 * zoom + vy} x2={g.x * zoom + vx} y2={g.y1 * zoom + vy} />
      ))}
      {guides.h.map((g, i) => (
        <line key={'h' + i} x1={g.x0 * zoom + vx} y1={g.y * zoom + vy} x2={g.x1 * zoom + vx} y2={g.y * zoom + vy} />
      ))}
    </svg>
  )
}


/* ★ 便利贴内容区（2026-10-22 用户指令「可以直接编辑内容 可以拖动尺寸」）：
   单击/聚焦 → 就地 textarea 编辑（blur 保存进 records）；
   NodeResizer 常驻（角标小把手，不用选组）拖右下角调尺寸。 */
export function NoteStickyBody({ nodeId, value }: { nodeId: string; value?: string }) {
  const [editing, setEditing] = useState(false)
  const [draft, setDraft] = useState('')
  const taRef = useRef<HTMLTextAreaElement | null>(null)
  const stAll = () => useBaseStore.getState()

  const commit = () => {
    setEditing(false)
    const cur = value ?? ''
    if (draft !== cur) {
      const st = stAll()
      st.pushHistory()
      st.setTableData(nodeId, [{ id: 'text', name: 'text', type: 'text' }], [{ text: draft }])
    }
  }

  if (editing) {
    return (
      <textarea
        ref={taRef}
        className="rf-stickynote-input"
        value={draft}
        autoFocus
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          e.stopPropagation()
          if (e.key === 'Escape') { setDraft(value ?? ''); setEditing(false) }
          if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) commit()
        }}
        rows={4}
      />
    )
  }
  return (
    <div
      className="rf-stickynote"
      title="便利贴 · 双击编辑文字 · 拖右缘把手调宽度"
      onDoubleClick={(e) => {
        /* ★ 双击进入就地编辑（2026-10-22 用户指令「双击在便利贴内编辑文字」）；
           单击保留给 RF 拖拽节点 */
        e.stopPropagation()
        if (!nodeId) return
        setDraft(value ?? '')
        setEditing(true)
        setTimeout(() => { taRef.current?.focus(); taRef.current?.select() }, 0)
      }}
    >
      <span className="rf-stickynote-text">{value}</span>
    </div>
  )
}

export function GroupNodeView({ nodeId, name }: { nodeId: string; name: string }) {
  const rf = useReactFlow()
  const rfWrapperRef = useRef<HTMLDivElement | null>(null)
  const rfWrapper = () => (rfWrapperRef.current?.closest(".react-flow") as HTMLElement | null)
  const cfg = useBaseStore((s) => s.tables[nodeId]?.config) as
    | { groupW?: number; groupH?: number; members?: string[]; groupFolded?: boolean; groupPrev?: { x: number; y: number; w: number; h: number } }
    | undefined
  const w = cfg?.groupW ?? 360
  const h = cfg?.groupH ?? 240
  const allMembers = (cfg?.members ?? []) as string[]
  /* ★ 2026-10-22 用户「删除节点后计数不实时更新」：成员按存活过滤——
     删除（软删）即从计数与组行为中消失，恢复删除又自动回来。 */
  const stTables = useBaseStore((s) => s.tables)
  const members = allMembers.filter((id) => {
    const t = stTables[id]
    return t && !t.deleted
  })
  const folded = cfg?.groupFolded === true
  const maxed = !!cfg?.groupPrev
  const stAll = () => useBaseStore.getState()

  /** 成员包围盒（flow 坐标，含默认尺寸） */
  const membersBBox = () => {
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity
    for (const m of members) {
      const ino = rf.getInternalNode(m)
      const p = ino?.internals?.positionAbsolute ?? { x: 0, y: 0 }
      const mw = ino?.measured?.width ?? 220
      const mh = ino?.measured?.height ?? 120
      x0 = Math.min(x0, p.x); y0 = Math.min(y0, p.y)
      x1 = Math.max(x1, p.x + mw); y1 = Math.max(y1, p.y + mh)
    }
    return x0 === Infinity ? null : { x: x0, y: y0, w: x1 - x0, h: y1 - y0 }
  }
  /** 成员网格重排：组内容区内均匀网格铺开（成员自身尺寸不变） */
  const gridMembers = (gx: number, gy: number, gw: number, gh: number) => {
    if (!members.length) return
    const st = stAll()
    st.pushHistory()
    const head = 34, pad = 26
    const innerW = gw - pad * 2
    const innerH = gh - head - pad * 2
    const cols = Math.max(1, Math.ceil(Math.sqrt(members.length)))
    const rows = Math.max(1, Math.ceil(members.length / cols))
    const cw = innerW / cols
    const ch = innerH / rows
    ;[...members].forEach((id, i) => {
      const c = i % cols
      const r = Math.floor(i / cols)
      const ino = rf.getInternalNode(id)
      const mw = ino?.measured?.width ?? 220
      const mh = ino?.measured?.height ?? 120
      const px = Math.round(gx + pad + c * cw + (cw - mw) / 2)
      const py = Math.round(gy + head + pad * 0.6 + r * ch + (ch - mh) / 2)
      st.setNodePosition(id, { x: px, y: py })
    })
  }

  /* ── 六钮 ── */
  const btn = (title: string, onClick: () => void, icon: React.ReactNode) => (
    <button
      className="group-node-btn"
      title={title}
      onPointerDown={(e) => e.stopPropagation()}
      onClick={(e) => { e.stopPropagation(); onClick() }}
    >
      {icon}
    </button>
  )
  const actDisband = () => {
    /* 解散：只删组节点（成员本来就在画布上，原地保留） */
    const st = stAll()
    st.pushHistory()
    st.trashTable(nodeId)
  }
  const actDeleteAll = () => {
    /* 删组及成员：组 + 全部成员进回收站 */
    const st = stAll()
    st.pushHistory()
    st.trashTables([...members, nodeId])
  }
  const actFit = () => {
    /* 适应大小：组框收缩到刚好包住所有成员（成员不动；统一 groupFitRect 约定） */
    const bb = membersBBox()
    if (!bb) return
    const st = stAll()
    st.pushHistory()
    const fit = groupFitRect(bb)
    st.setNodePosition(nodeId, { x: fit.x, y: fit.y })
    st.setNodeConfig(nodeId, { groupW: fit.w, groupH: fit.h })
  }
  const actRearrange = () => {
    /* 重排：成员在组内容区网格铺开（成员尺寸不变） */
    const gn = rf.getInternalNode(nodeId)
    const gx = gn?.internals?.positionAbsolute?.x ?? 0
    const gy = gn?.internals?.positionAbsolute?.y ?? 0
    gridMembers(gx, gy, w, h)
  }
  const actFold = () => {
    const st = stAll()
    st.pushHistory()
    st.setNodeConfig(nodeId, { groupFolded: !folded })
  }
  const actMaximize = () => {
    const st = stAll()
    st.pushHistory()
    if (maxed && cfg?.groupPrev) {
      /* 还原 */
      st.setNodePosition(nodeId, { x: cfg.groupPrev.x, y: cfg.groupPrev.y })
      st.setNodeConfig(nodeId, { groupW: cfg.groupPrev.w, groupH: cfg.groupPrev.h, groupPrev: undefined })
      return
    }
    /* 放大：铺满画布可视区（成员尺寸不变，网格重排铺开），再 fitView 到组窗体 */
    const el = rfWrapper()
    const r = el?.getBoundingClientRect()
    if (!r) return
    const tl = rf.screenToFlowPosition({ x: r.left + 24, y: r.top + 88 })
    const br = rf.screenToFlowPosition({ x: r.right - 24, y: r.bottom - 24 })
    const self = rf.getInternalNode(nodeId)
    const prev = {
      x: self?.internals?.positionAbsolute?.x ?? 0,
      y: self?.internals?.positionAbsolute?.y ?? 0,
      w,
      h,
    }
    st.setNodePosition(nodeId, { x: Math.round(tl.x), y: Math.round(tl.y) })
    st.setNodeConfig(nodeId, {
      groupW: Math.round(br.x - tl.x),
      groupH: Math.round(br.y - tl.y),
      groupPrev: prev,
    })
    gridMembers(Math.round(tl.x), Math.round(tl.y), Math.round(br.x - tl.x), Math.round(br.y - tl.y))
    setTimeout(() => rf.fitView({ nodes: [{ id: nodeId }], padding: 0.06, duration: 320, maxZoom: 1 }), 60)
  }

  const grpSelected = useStore((s) => !!s.nodeLookup.get(nodeId)?.selected)
  /* ★ 计数下拉（2026-10-22 用户指令「点击下拉选择，下拉还有一个全部选择」） */
  const [countOpen, setCountOpen] = useState(false)
  /* ★ 2026-10-22 用户「拖动节点组或点击空白时取消下拉」：
     ① 监听画布 pointerdown（capture）——点空白/点其它节点即收起；
     ② 组节点开始拖动（RF NodeDrag 事件）即收起。 */
  useEffect(() => {
    if (!countOpen) return
    /* 点菜单自身不关（capture 阶段判 target 是否在菜单 DOM 内）；
       点空白/点其它节点/点画布任意处 → 收起。 */
    const close = (ev: PointerEvent) => {
      const t = ev.target as HTMLElement | null
      if (t && typeof t.closest === 'function' && t.closest('.group-count-menu, .group-node-count-btn')) return
      setCountOpen(false)
    }
    window.addEventListener('pointerdown', close, true)
    return () => window.removeEventListener('pointerdown', close, true)
  }, [countOpen])
  /* 组开始拖动 → 收起（订阅 RF store：任意节点 drag 开始标志） */
  const anyDragging = useStore((s) => s.nodes.some((n) => n.dragging))
  useEffect(() => { if (anyDragging) setCountOpen(false) }, [anyDragging])
  const [countPos, setCountPos] = useState<{ left: number; top: number }>({ left: 0, top: 0 })
  /** 选中给定集合（+组自身），其余取消——与画布框选同款 RF 选中态 */
  const selectIds = (ids: Set<string>) => {
    try {
      rf.setNodes((nds) =>
        nds.map((n) => {
          const sel = ids.has(n.id) || n.id === nodeId
          return sel !== !!n.selected ? { ...n, selected: sel } : n
        }),
      )
    } catch { /* 选中失败静默 */ }
  }
  /* ★ 组内有成员被选中 → 隐藏缩放手柄（2026-10-21 用户「拖到组左/上边缘被卡住」）：
     组选中态下沿边的手柄会截胡成员边缘的按压（误启动组缩放，成员拖不动）。
     隐藏后成员可正常拖拽；点组标题栏重新选中组 → 手柄回归。 */

  return (
    <div className="group-node" style={{ width: '100%', height: '100%' }} ref={rfWrapperRef}>
      <div className="group-node-head" title={`节点组 · ${members.length} 个成员（拖标题栏移动整组）`}>
        <span className="group-node-title">{name}</span>
        {btn('解散节点组（节点放回画布）', actDisband, <IcUnlock size={14} />)}
        {btn('删除节点组及节点（进回收站）', actDeleteAll, <IcTrash size={14} />)}
        {btn('适应大小：组框自动包住所有成员（拖入新节点也会自动扩展）', actFit, <IcShrink size={14} />)}
        {btn('重排：组内成员网格铺开', actRearrange, <IcGrid size={14} />)}
        {btn(folded ? '展开组内成员' : '折叠组内成员', actFold, <IcFold size={14} />)}
        {btn(maxed ? '还原组窗体' : '放大组窗体（成员尺寸不变，自动铺开）', actMaximize, maxed ? <IcMinimize size={14} /> : <IcMaximize size={14} />)}
        {/* ★ 计数=下拉选择器（2026-10-22 用户指令「点击下拉选择，下拉还有一个
            全部选择」）：点数字展开成员清单，点某项=仅选中该成员；首项「全部」=
            选中组内全部成员。 */}
        <div className="group-count-wrap" style={{ position: 'relative' }}>
          <button
            className="group-node-count group-node-count-btn"
            title={'点击展开成员清单（共 ' + members.length + ' 个）'}
            onPointerDown={(e) => e.stopPropagation()}
            onClick={(e) => {
              e.stopPropagation()
              /* ★ fixed 定位坐标：按钮矩形右下角；右侧空间不足 → 弹左侧 */
              const r = (e.currentTarget as HTMLElement).getBoundingClientRect()
              const vw = window.innerWidth
              const menuW = 190
              setCountPos({ left: r.right + menuW > vw - 8 ? Math.max(8, r.left - menuW) : r.right, top: r.bottom + 4 })
              setCountOpen((v) => !v)
            }}
          >
            {members.length}
          </button>
          {countOpen && createPortal(
            /* ★ 2026-10-22 修「下拉跑远」：fixed 定位在 RF viewport（transform 祖先）
               内会退化为相对 transform 定位 → 屏幕 rect 被二次映射跑飞。
               portal 到 body（无 transform 祖先）→ fixed 坐标即真实视口坐标，
               菜单贴着按钮右侧弹出。 */
            <div
              className="group-count-menu"
              style={{ position: 'fixed', left: countPos.left, top: countPos.top }}
              onPointerDown={(e) => e.stopPropagation()}
              onDoubleClick={(e) => e.stopPropagation()}
            >
              <button
                className="group-count-item group-count-item--all"
                onClick={(e) => {
                  e.stopPropagation()
                  selectIds(new Set(members))
                  setCountOpen(false)
                }}
              >
                全部选择（{members.length}）
              </button>
              {members.map((id) => (
                <button
                  key={id}
                  className="group-count-item"
                  onClick={(e) => {
                    e.stopPropagation()
                    selectIds(new Set([id]))
                    setCountOpen(false)
                  }}
                >
                  {String(stTables[id]?.name ?? id)}
                </button>
              ))}
            </div>,
            document.body,
          )}
        </div>
      </div>
      <NodeResizer
        nodeId={nodeId}
        isVisible={grpSelected && !folded}
        minWidth={220}
        minHeight={140}
        lineClassName="group-resize-line"
        handleClassName="group-resize-handle"
        onResize={(_e: unknown, params: { x: number; y: number; width: number; height: number }) => {
          /* ★ 2026-10-22 用户「拖组边框让节点向内推」：缩放中实时把越界成员
             推进组内（贴边停）——经 rf.updateNode 走 RF 内部态（store 通道
             会被重建的 existing.position 顶掉）。左界=组左缘，上界=标题栏底线。 */
          const left = params.x
          const top = params.y + 34
          const right = params.x + params.width
          const bottom = params.y + params.height
          for (const id of members) {
            const ino = rf.getInternalNode(id)
            if (!ino) continue
            const p = ino.internals.positionAbsolute
            const mw = ino.measured?.width ?? 220
            const mh = ino.measured?.height ?? 120
            const nx = Math.min(Math.max(p.x, left), right - mw)
            const ny = Math.min(Math.max(p.y, top), bottom - mh)
            if (Math.abs(nx - p.x) > 0.5 || Math.abs(ny - p.y) > 0.5) {
              rf.updateNode(id, (nd: { position?: { x: number; y: number } }) => ({
                position: { x: nx, y: ny },
              }))
            }
          }
        }}
        onResizeEnd={(_e: unknown, params: { x: number; y: number; width: number; height: number }) => {
          const st = stAll()
          st.pushHistory()
          st.setNodeConfig(nodeId, { groupW: Math.round(params.width), groupH: Math.round(params.height) })
          st.setNodePosition(nodeId, { x: Math.round(params.x), y: Math.round(params.y) })
          /* ★ 被推过的成员最终位置落库（RF 内部态 → store 持久化） */
          for (const id of members) {
            const ino = rf.getInternalNode(id)
            const p = ino?.internals?.positionAbsolute
            if (p) st.setNodePosition(id, { x: Math.round(p.x), y: Math.round(p.y) })
          }
          /* ★ 2026-10-22 用户指令「修改节点组大小…也在诊断日志里」 */
          pushLog(`节点组 ${st.tables[nodeId]?.name ?? nodeId} 大小 → ${Math.round(params.width)}×${Math.round(params.height)}，位置 → (${Math.round(params.x)}, ${Math.round(params.y)})`)
        }}
      />
    </div>
  )
}
