import { useEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import type { TableSchema } from '../core/schema'
import type { Row } from '../core/store'
import { computeNode } from '../core/engine'
import { capsuleKindClass } from '../ref/capsule'
import { PARAMLIB_ITEMS } from '../ParamlibPanel'
import { useBaseStore } from '../core/store'
import { refColorOf } from '../core/refColor'
import { domainRawKey, domainNameOf } from '../core/refKey'
import {
  dropPin,
  hasPin,
  loadPins,
  PRIO_KEY,
  savePins,
  STAR_KEY,
  subscribePins,
  upsertPin,
  type RefPin,
} from './refPins'
import { getRefUi, setRefUi, subscribeRefUi } from './refPanelUi'

/**
 * 属性引用栏（画布级侧栏，portal 到 body · fixed 贴屏幕左缘）：
 *   · 索引画布上所有节点（一条一条可折叠）与节点生成的参数（字段，含首行值预览）
 *   · 点击字段 = 当前选中格写入引用（node://<节点>#<字段>），同时登记 linkedInputs —— 画布自动出「数据流」连线
 *   · 非最大化时独立于节点窗体贴最左；最大化后窗体占满全屏，视觉上即「集成在窗体内」
 *
 * 批次 1-4 引用栏改版（原待办 D）：
 *   · 栏内「节点下拉」直接选节点 → 展开/切换该节点参数（不用回画布找节点）
 *   · 第一行「优先参数」：选中引用格 → 点异色标签 → 收录进所有引用栏共享的优先区（refPins localStorage）
 *   · 第二行节点参数树（现有）+ 参数行 ☆ 星标
 *   · 第三行「星标参数」：树上星标的参数汇总（同为全局共享）
 */

const KIND_LABEL: Record<string, string> = {
  table: '表',
  param: '参数表',
  lookup: '对照表',
  sort: '排序',
  output: '输出',
  docx: 'Word',
  xlsx: 'Excel',
  text: '文本',
  start: '开始',
  map: '对照',
}

export function RefPanel(props: {
  tables: Record<string, TableSchema>
  records: Record<string, Row[]>
  selfId: string
  /** 当前选中格的引用（用于高亮对应字段项 + 自动展开节点） */
  activeRef: { nodeId: string; fieldId: string } | null
  onPick: (nodeId: string, nodeIdName: string, fieldId: string, fieldName: string, value?: string, mode?: string) => void
  /** 参数库内容点选（2026-10-23）：value=纯值插入；ref=链接（走 paramlib 节点真引用链） */
  onPickLibValue?: (fieldName: string, value: string, mode: 'value' | 'ref', libNodeId?: string) => void
  onClose: () => void
  /** 位置/尺寸/圆角（由 WorkbookEditor 按 普通视图=窗体外卡片 / 最大化=嵌入窗体左缘 计算） */
  panelStyle?: React.CSSProperties
  /** 备忘 · 进/退场动画类（普通视图一退一进；最大化不加） */
  animClass?: string
  /** 左栏内部布局变化（展开/折叠/搜索）→ 通知父级重算连线 */
  onLayoutChange?: () => void
}) {
  // ★ UI 状态与 /poc 内嵌面板共享（2026-10-11 用户指令「最大化/常规内容同步」）：
  //   初始化从共享 store 读（最大化来回不丢）；变更写回 store。
  const [q, setQ] = useState(() => getRefUi().q)
  const [openIds, setOpenIds] = useState<Set<string>>(() => new Set(getRefUi().openIds))
  useEffect(() => {
    setRefUi({ q })
  }, [q])
  useEffect(() => {
    setRefUi({ openIds: [...openIds] })
  }, [openIds])
  // /poc 内嵌面板变更回传 → 本浮卡跟随（两边同屏时视觉一致）
  useEffect(
    () =>
      subscribeRefUi((s) => {
        setQ((cur) => (cur === s.q ? cur : s.q))
        setOpenIds((cur) => {
          const next = new Set(s.openIds)
          if (cur.size === next.size && [...next].every((x) => cur.has(x))) return cur
          return next
        })
      }),
    [],
  )
  // 批次 1-4：优先/星标收藏（全局共享）+ 节点下拉选中值
  const [prio, setPrio] = useState<RefPin[]>(() => loadPins(PRIO_KEY))
  const [star, setStar] = useState<RefPin[]>(() => loadPins(STAR_KEY))
  const [selNode, setSelNode] = useState('')

  // 左栏布局变化 → 通知父级重算连线端点
  const notifyLayout = () => props.onLayoutChange?.()

  // 收录/移除/其它编辑器实例写入 → 重读（模块级订阅；跨窗口 storage 兜底在 refPins 内部挂）
  useEffect(
    () =>
      subscribePins(() => {
        setPrio(loadPins(PRIO_KEY))
        setStar(loadPins(STAR_KEY))
      }),
    [],
  )

  // 选中引用变化 → 自动展开对应节点（一次性；不强制展开，节点仍可手动折叠）+ 下拉跟随
  useEffect(() => {
    if (props.activeRef && !openIds.has(props.activeRef.nodeId)) {
      setOpenIds((s) => new Set(s).add(props.activeRef!.nodeId))
      notifyLayout()
    }
    setSelNode(props.activeRef?.nodeId ?? '')
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.activeRef?.nodeId])

  const nodes = useMemo(() => {
    const list: { id: string; name: string; kind: string; fields: { id: string; name: string; color: string }[]; firstVals: Record<string, string> }[] = []
    for (const t of Object.values(props.tables)) {
      if (t.deleted || t.id === props.selfId) continue
      let fields: { id: string; name: string; color: string }[] = []
      let firstVals: Record<string, string> = {}
      try {
        const comp = computeNode(t.id, props.tables, props.records)
        fields = comp.fields.map((f) => ({
          id: f.id, name: f.name,
          // 与文档内引用格底色同源（core/refColor 的同一函数 + 同一名称口径）
          color: refColorOf(domainNameOf(domainRawKey(t.name, f.name))),
        }))
        const r0 = comp.rows[0] as Record<string, unknown> | undefined
        for (const f of fields) {
          const v = r0?.[f.id]
          firstVals[f.id] = v == null ? '' : String(v)
        }
      } catch {
        /* 节点计算失败（缺上游等）→ 只列名字 */
      }
      list.push({ id: t.id, name: t.name, kind: t.kind ?? 'table', fields, firstVals })
    }
    return list
  }, [props.tables, props.records, props.selfId])

  // 收藏项 → 节点/字段信息索引（含失效判定）
  const nodeIdx = useMemo(() => new Map(nodes.map((n) => [n.id, n])), [nodes])

  const ql = q.trim().toLowerCase()
  const visible = nodes
    .map((n) => {
      if (!ql) return n
      const hitNode = n.name.toLowerCase().includes(ql)
      const fs = n.fields.filter((f) => f.name.toLowerCase().includes(ql) || f.id.toLowerCase().includes(ql))
      if (hitNode) return n
      if (fs.length) return { ...n, fields: fs }
      return null
    })
    .filter((x): x is NonNullable<typeof x> => !!x)

  const toggle = (id: string) =>
    setOpenIds((s) => {
      const next = new Set(s)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      return next
    })

  // ── 批次 1-4：收藏 toggle（savePins 内部 emit → 订阅回调重读 state，无需手动 set）──
  const togglePin = (kind: 'prio' | 'star', nodeId: string, fieldId: string) => {
    const key = kind === 'prio' ? PRIO_KEY : STAR_KEY
    const cur = kind === 'prio' ? prio : star
    savePins(key, hasPin(cur, nodeId, fieldId) ? dropPin(cur, nodeId, fieldId) : upsertPin(cur, nodeId, fieldId))
  }

  // 优先参数标签语义：选中引用格 → toggle 收录；无选中格 → disabled
  const prioTarget = props.activeRef
  const prioPinned = prioTarget ? hasPin(prio, prioTarget.nodeId, prioTarget.fieldId) : false

  /** 收藏行渲染：胶囊 + 首行值 + × 移除；点击行 = 写入选中格引用（与参数树同语义）；失效红边 */
  const renderPin = (p: RefPin, kind: 'prio' | 'star') => {
    const n = nodeIdx.get(p.nodeId)
    const f = n?.fields.find((x) => x.id === p.fieldId)
    const dead = !n || !f
    const nodeName = n?.name ?? p.nodeId
    const fieldName = f?.name ?? p.fieldId
    const sample = n?.firstVals[p.fieldId] ?? ''
    return (
      <div
        key={`${p.nodeId}#${p.fieldId}`}
        className={`ref-pin${dead ? ' dead' : ''}`}
        data-node={p.nodeId}
        data-field={p.fieldId}
        // 不抢焦点：编辑中点收藏，引用要插进当前编辑器
        onMouseDown={(e) => e.preventDefault()}
        title={dead ? '来源已失效（节点或参数已删除）' : `双击参数名=插入引用地址；双击数值=插入数值（node://${p.nodeId}#${p.fieldId}）`}
        onDoubleClick={(e) => {
          // ★ 2026-10-11：同上，阻断 portal 冒泡（收藏条双击不再穿透祖先处理器）
          e.stopPropagation()
          if (!dead) props.onPick(p.nodeId, nodeName, p.fieldId, fieldName, String(sample || ''), 'ref')
        }}
      >
        <span
          className={`ref-capsule ${capsuleKindClass(n?.kind ?? '')}${dead ? ' dead' : ''}`}
          title={dead ? '来源已失效' : `双击插入引用地址：${nodeName}.${fieldName}`}
          style={{ cursor: 'pointer' }}
          onDoubleClick={(e) => {
            e.stopPropagation()
            if (!dead) props.onPick(p.nodeId, nodeName, p.fieldId, fieldName, String(sample || ''), 'ref')
          }}
        >
          <i className="cap-dot" aria-hidden="true" />
          <span className="cap-node">{nodeName}</span>
          <span className="cap-field">{dead ? '失效' : fieldName}</span>
        </span>
        {!dead && sample && (
          <span
            className="ref-fval"
            title={`双击插入数值：${sample}`}
            style={{ cursor: 'pointer' }}
            onDoubleClick={(e) => {
              e.stopPropagation()
              props.onPick(p.nodeId, nodeName, p.fieldId, fieldName, String(sample || ''), 'value')
            }}
          >
            {sample}
          </span>
        )}
        <span
          className="ref-pin-x"
          title={kind === 'prio' ? '移出优先参数' : '移除星标'}
          onMouseDown={(e) => e.preventDefault()}
          onClick={(e) => {
            e.stopPropagation()
            togglePin(kind, p.nodeId, p.fieldId)
          }}
        >
          ×
        </span>
      </div>
    )
  }

  // 显示/隐藏由编辑器左下角的「属性引用」开关控制（wb-foot），面板内只保留 × 关闭

  return createPortal(
    <div className={`ref-panel${props.animClass ? ` ${props.animClass}` : ''}`} style={props.panelStyle}>
      <div className="ref-head">
        <span className="ref-title">节点数据引用</span>
        <span className="ref-sp" />
        <button className="ref-x" title="收起属性引用栏" onClick={props.onClose}>
          ×
        </button>
      </div>
      {/* ★ 2026-10-23 用户指令「数据引用最上方 切换参数库数据选择 优先选连接的参数库」：
          优先取与主节点连线（linkedInputs）的 paramlib 节点所关联的库；
          选择写 config.paramLibSel（随节点持久）——下游灌数后续接 */}
      {(() => {
        const self = props.tables[props.selfId]
        const linked = ((self?.config?.linkedInputs ?? []) as string[])
          .map((id) => props.tables[id])
          .find((t) => t && t.kind === 'paramlib' && !t.deleted)
        const linkedLib = ((linked?.config ?? {}) as { libName?: string }).libName
        const cur = linkedLib ?? ((self?.config ?? {}) as { paramLibSel?: string }).paramLibSel ?? ''
        const st = useBaseStore.getState()
        return (
          <div className="ref-paramlib" title={linkedLib ? '已优先选中与主节点连接的参数库' : '选择参数库数据'}>
            <span className="ref-paramlib-label">📚 参数库</span>
            <SearchSelect
              value={cur}
              placeholder="（未选择）"
              title="选择参数库（可搜索）"
              options={[
                { value: '', label: '（未选择）' },
                ...PARAMLIB_ITEMS.map((it) => ({
                  value: it.name,
                  label: it.name + "（" + it.rows.length + " 条）" + (it.name === linkedLib ? ' · 已连接' : ''),
                })),
              ]}
              onChange={(v) => {
                st.pushHistory()
                st.setNodeConfig(props.selfId, { paramLibSel: v })
              }}
            />
          </div>
        )
      })()}
      {/* ★ 2026-10-23 用户指令「选择参数库内的内容 然后关联进excel」：
          选中库后展开内容列表（搜索 + 前 80 条）——点单元格 = 该值插入 Excel 光标处；
          插入走 onPickLibValue（NodeEditor 侧 insertRef 纯文本模式，不建域关联） */}
      {(() => {
        const self = props.tables[props.selfId]
        const entry = PARAMLIB_ITEMS.find((it) => it.name === ((self?.config ?? {}) as { paramLibSel?: string }).paramLibSel)
          ?? PARAMLIB_ITEMS.find((it) => it.name === ((self?.config ?? {}) as { libName?: string }).libName)
        if (!entry) return null
        const linkedPl = ((self?.config?.linkedInputs ?? []) as string[])
          .map((id) => props.tables[id])
          .find((t) => t && t.kind === 'paramlib' && !t.deleted)
        return <LibSelect entry={entry} libNodeId={linkedPl?.id} onPickCell={props.onPickLibValue} />
      })()}
      <select
        className="ref-nodesel"
        data-refsel="1"
        value={selNode}
        title="直接选择节点 → 展开该节点参数"
        onChange={(e) => {
          const id = e.target.value
          setSelNode(id)
          // 指定/切换该节点参数：只展开所选节点（其余收起），树内一目了然
          if (id) {
            setOpenIds(new Set([id]))
            notifyLayout()
          }
        }}
      >
        <option value="">所有节点（点 ▾ 指定）</option>
        {nodes.map((n) => (
          <option key={n.id} value={n.id}>
            {n.name}
          </option>
        ))}
      </select>
      {/* 第一行 · 优先参数（全局共享优先区） */}
      <div className="ref-sec ref-sec-prio" data-refsec="prio">
        <div className="ref-sec-head">
          <button
            className={`ref-sec-tag prio${prioPinned ? ' on' : ''}`}
            disabled={!prioTarget}
            title={
              prioTarget
                ? prioPinned
                  ? '点击移出优先参数'
                  : '点击收录当前选中格的引用（所有引用栏共享）'
                : '先点选一个引用格，再点此收录'
            }
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              if (prioTarget) togglePin('prio', prioTarget.nodeId, prioTarget.fieldId)
            }}
          >
            ◆ 优先参数
          </button>
          {prio.length > 0 && <span className="ref-sec-cnt">{prio.length}</span>}
        </div>
        {prio.length === 0 ? (
          <div className="ref-sec-empty">选中引用格后点「◆ 优先参数」收录</div>
        ) : (
          <div className="ref-pinlist">{prio.map((p) => renderPin(p, 'prio'))}</div>
        )}
      </div>
      <input
        className="ref-search"
        placeholder="搜索节点 / 参数…"
        value={q}
        onChange={(e) => { setQ(e.target.value); notifyLayout() }}
      />
      {/* 第二行 · 节点参数树（现有）+ 参数行 ☆ 星标 */}
      <div className="ref-list">
        {visible.length === 0 && <div className="ref-empty">画布上还没有其它节点</div>}
        {visible.map((n) => {
          // open 只由 openIds/搜索驱动 —— 连过线的节点（activeRef）同样允许折叠
          const open = openIds.has(n.id) || !!q
          const isActiveNode = props.activeRef?.nodeId === n.id
          return (
            <div key={n.id} className="ref-node">
              <button
                className={`ref-node-row${isActiveNode ? ' on' : ''}`}
                data-node={n.id}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => { toggle(n.id); notifyLayout() }}
                title={n.id}
              >
                <span className="ref-caret">{open ? '▾' : '▸'}</span>
                <span className="ref-node-name">{n.name}</span>
                <span className="ref-kind">{KIND_LABEL[n.kind] ?? n.kind}</span>
              </button>
              {open && (
                <div className="ref-fields">
                  {n.fields.length === 0 && <div className="ref-field-empty">（无参数）</div>}
                  {n.fields.map((f) => {
                    const isActive = isActiveNode && props.activeRef?.fieldId === f.id
                    const sample = n.firstVals[f.id] ?? ''
                    const starred = hasPin(star, n.id, f.id)
                    return (
                      <div
                        key={f.id}
                        className={`ref-field${isActive ? ' on' : ''}`}
                        data-node={n.id}
                        data-field={f.id}
                        // 不抢焦点：= 编辑中点参数，引用要插进当前编辑器，焦点不能走
                        onMouseDown={(e) => e.preventDefault()}
                        title={`双击参数名=插入引用地址；双击数值=插入数值（node://${n.id}#${f.id}）`}
                        onDoubleClick={(e) => {
                          // ★ 2026-10-11：portal 事件沿 React 组件树冒泡，不阻断会触发祖先
                          //   的双击处理器（用户实测：双击参数 → 节点引用面板开合跳动）
                          e.stopPropagation()
                          props.onPick(n.id, n.name, f.id, f.name, String(sample || ''), 'ref')
                        }}
                      >
                        <span
                          className={`ref-starbtn${starred ? ' on' : ''}`}
                          data-star="1"
                          title={starred ? '取消星标' : '星标参数'}
                          onMouseDown={(e) => e.preventDefault()}
                          onClick={(e) => {
                            e.stopPropagation()
                            togglePin('star', n.id, f.id)
                          }}
                        >
                          {starred ? '★' : '☆'}
                        </span>
                        {/* 批次 8 · 行内嵌两段式胶囊：左=节点名段 · 右=参数名段（值预览仍在行右侧） */}
                        <span
                          className={`ref-capsule ${capsuleKindClass(n.kind)}`}
                          title={`双击插入引用地址：${n.name}.${f.name}`}
                          style={{ cursor: 'pointer' }}
                          onDoubleClick={(e) => {
                            e.stopPropagation()
                            props.onPick(n.id, n.name, f.id, f.name, String(sample || ''), 'ref')
                          }}
                        >
                          <i className="cap-dot" aria-hidden="true" />
                          <span className="cap-node">{n.name}</span>
                          <span className="cap-field">{f.name}</span>
                        </span>
                        {/* ★ 2026-10-17 用户指令「对参数引用的地址加个颜色区分」：
                            色块 = 插入引用后文档里该格会带的底色（与 core/refColor 同一算法，
                            面板上看到什么色，文档里就是什么色）。 */}
                        <i
                          className="ref-swatch"
                          data-color={f.color}
                          title={`文档中该参数的引用格底色 #${String(f.color).slice(-6)}`}
                          style={{
                            width: 9,
                            height: 9,
                            flex: '0 0 auto',
                            borderRadius: 2,
                            background: '#' + String(f.color).slice(-6),
                            border: '1px solid rgba(128,128,128,.45)',
                            alignSelf: 'center',
                          }}
                        />
                        {sample && (
                          <span
                            className="ref-fval"
                            title={`双击插入数值：${sample}`}
                            style={{ cursor: 'pointer' }}
                            onDoubleClick={(e) => {
                              e.stopPropagation()
                              props.onPick(n.id, n.name, f.id, f.name, String(sample || ''), 'value')
                            }}
                          >
                            {sample}
                          </span>
                        )}
                      </div>
                    )
                  })}
                </div>
              )}
            </div>
          )
        })}
      </div>
      {/* 第三行 · 星标参数（全局共享） */}
      <div className="ref-sec ref-sec-star" data-refsec="star">
        <div className="ref-sec-head">
          <span className="ref-sec-tag star">☆ 星标参数</span>
          {star.length > 0 && <span className="ref-sec-cnt">{star.length}</span>}
        </div>
        {star.length === 0 ? (
          <div className="ref-sec-empty">参数行点 ☆ 收藏常用参数</div>
        ) : (
          <div className="ref-pinlist">{star.map((p) => renderPin(p, 'star'))}</div>
        )}
      </div>
      <div className="ref-tip">点击参数 → 写入选中格引用，画布自动连线</div>
    </div>,
    document.body,
  )
}

/**
 * 备忘 · 右侧「数据预选追加」面板（画布级侧栏，镜像左栏）：
 * 勾选上游节点 = 把它的数据追加进本节点的数据流来源（config.inputs）；取消 = 移除。
 * 勾选后展开该节点的参数（含首行值预览），确认「预选了什么」。
 */
export function RightPanel(props: {
  tables: Record<string, TableSchema>
  records: Record<string, Row[]>
  selfId: string
  /** 当前已追加为数据流来源的节点 id（config.inputs） */
  inputs: string[]
  onToggle: (srcId: string, on: boolean) => void
  onClose: () => void
  panelStyle?: React.CSSProperties
  animClass?: string
}) {
  const nodes = useMemo(() => {
    const list: { id: string; name: string; kind: string; fields: { id: string; name: string; color: string }[]; firstVals: Record<string, string> }[] = []
    for (const t of Object.values(props.tables)) {
      if (t.deleted || t.id === props.selfId) continue
      let fields: { id: string; name: string; color: string }[] = []
      const firstVals: Record<string, string> = {}
      try {
        const comp = computeNode(t.id, props.tables, props.records)
        fields = comp.fields.map((f) => ({
          id: f.id, name: f.name,
          // 与文档内引用格底色同源（core/refColor 的同一函数 + 同一名称口径）
          color: refColorOf(domainNameOf(domainRawKey(t.name, f.name))),
        }))
        const r0 = comp.rows[0] as Record<string, unknown> | undefined
        for (const f of fields) firstVals[f.id] = r0?.[f.id] == null ? '' : String(r0?.[f.id])
      } catch {
        /* 节点计算失败（缺上游等）→ 只列名字 */
      }
      list.push({ id: t.id, name: t.name, kind: t.kind ?? 'table', fields, firstVals })
    }
    return list
  }, [props.tables, props.records, props.selfId])

  return createPortal(
    <div className={`ref-panel ref-panel-right${props.animClass ? ` ${props.animClass}` : ''}`} style={props.panelStyle}>
      <div className="ref-head">
        <span className="ref-title">数据预选追加</span>
        <span className="ref-sp" />
        <button className="ref-x" title="收起数据预选追加" onClick={props.onClose}>
          ×
        </button>
      </div>
      <div className="ref-list">
        {nodes.length === 0 && <div className="ref-empty">画布上还没有其它节点</div>}
        {nodes.map((n) => {
          const on = props.inputs.includes(n.id)
          return (
            <div key={n.id} className="ref-node">
              <label className={`ref-node-row ref-node-check${on ? ' on' : ''}`} data-node={n.id} title={on ? '已在数据流来源中' : '勾选 = 追加进本节点数据流'}>
                <input
                  type="checkbox"
                  checked={on}
                  onChange={(e) => props.onToggle(n.id, e.target.checked)}
                />
                <span className="ref-node-name">{n.name}</span>
                <span className="ref-kind">{KIND_LABEL[n.kind] ?? n.kind}</span>
              </label>
              {on && (
                <div className="ref-fields">
                  {n.fields.length === 0 && <div className="ref-field-empty">（无参数）</div>}
                  {n.fields.map((f) => (
                    <div key={f.id} className="ref-field ref-field-ro">
                      <span className="ref-fname">{f.name}</span>
                      {n.firstVals[f.id] && <span className="ref-fval">{n.firstVals[f.id]}</span>}
                    </div>
                  ))}
                </div>
              )}
            </div>
          )
        })}
      </div>
      <div className="ref-tip">勾选节点 → 追加进数据流来源（可再取消）</div>
    </div>,
    document.body,
  )
}


/** 参数库内容下拉（2026-10-23 用户澄清「内容是一个一个拉拉选择 不是把表列出来；
 *  选择参数库 然后下拉选择插入内容」）：内容下拉（每行一条摘要），选中即**链接**
 *  （ref 模式定义名称活引用，走 paramlib 节点真引用链）。 */
/** 可搜索下拉（2026-10-23 用户指令「选择的时候 提供一个搜索 下拉选择的栏」）：
 *  收起态 = 按钮样式；点开 = 搜索框 + 过滤列表（portal fixed，锚按钮矩形）。
 *  点外/ESC 关闭。库选择与内容选择两处共用。 */
function SearchSelect ({
  value,
  options,
  onChange,
  placeholder,
  title,
}: {
  value: string
  options: { value: string; label: string }[]
  onChange: (v: string) => void
  placeholder?: string
  title?: string
}) {
  const [open, setOpen] = useState(false)
  const [q, setQ] = useState('')
  const [pos, setPos] = useState({ left: 0, top: 0, width: 240 })
  const btnRef = useRef<HTMLButtonElement | null>(null)
  const kw = q.trim().toLowerCase()
  const filtered = useMemo(
    () => (kw ? options.filter((o) => o.label.toLowerCase().includes(kw)) : options),
    [options, kw],
  )
  const curLabel = options.find((o) => o.value === value)?.label ?? placeholder ?? '（未选择）'
  useEffect(() => {
    if (!open) return
    const onDown = (e: MouseEvent) => {
      const t = e.target as HTMLElement
      if (!t.closest('.ssearch-panel') && !t.closest('.ssearch-btn')) setOpen(false)
    }
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') setOpen(false) }
    document.addEventListener('mousedown', onDown)
    document.addEventListener('keydown', onKey)
    return () => {
      document.removeEventListener('mousedown', onDown)
      document.removeEventListener('keydown', onKey)
    }
  }, [open])
  return (
    <>
      <button
        ref={btnRef}
        className="ref-paramlib-sel ssearch-btn"
        title={title}
        onClick={() => {
          const r = (btnRef.current as HTMLElement).getBoundingClientRect()
          setPos({ left: r.left, top: r.bottom + 4, width: Math.max(r.width, 240) })
          setQ('')
          setOpen((v) => !v)
        }}
      >
        <span className="ssearch-btn-label">{curLabel}</span>
        <span className="ssearch-caret">▾</span>
      </button>
      {open && createPortal(
        <div className="ssearch-mask" onClick={() => setOpen(false)}>
          <div
            className="ssearch-panel"
            style={{ left: Math.min(pos.left, window.innerWidth - pos.width - 8), top: pos.top, width: pos.width }}
            onClick={(e) => e.stopPropagation()}
          >
            <input
              autoFocus
              className="ssearch-q"
              value={q}
              placeholder="搜索…"
              onChange={(e) => setQ(e.target.value)}
            />
            <div className="ssearch-list">
              {filtered.length === 0 && <div className="ssearch-empty">无匹配项</div>}
              {filtered.map((o) => (
                <div
                  key={o.value}
                  className={'ssearch-item' + (o.value === value ? ' ssearch-item--on' : '')}
                  onClick={() => { onChange(o.value); setOpen(false) }}
                >
                  {o.label}
                </div>
              ))}
            </div>
          </div>
        </div>,
        document.body,
      )}
    </>
  )
}

function LibSelect ({
  entry,
  libNodeId,
  onPickCell,
}: {
  entry: { name: string; fields: string[]; rows: string[][] }
  libNodeId?: string
  onPickCell?: (fieldName: string, value: string, mode: 'value' | 'ref', libNodeId?: string) => void
}) {
  const summary = (r: string[]) => {
    const parts = [r[1] || r[0], r[4] || r[3]].filter(Boolean)
    return parts.length ? parts.join(' / ') : (r.filter(Boolean)[0] || '(空)')
  }
  return (
    <div className="ref-paramlib" style={{ borderBottom: 'none' }}>
      <span className="ref-paramlib-label">内容</span>
      <SearchSelect
        value=""
        placeholder={'下拉选择插入内容（共 ' + entry.rows.length + ' 条）…'}
        title="选择要插入的参数内容——选中即链接至当前 Excel/Word（可搜索）"
        options={entry.rows.map((r, i) => ({ value: String(i), label: summary(r) }))}
        onChange={(v) => {
          const idx = Number(v)
          if (!Number.isFinite(idx) || idx < 0 || idx >= entry.rows.length) return
          const r = entry.rows[idx]
          /* ★ 2026-10-23 用户指令「引用的参数库 直接引用名字就好了」：引用名 = 名称本身（不带库名前缀/行号） */
          const fieldName = r[1] || r[0] || ('条目' + (idx + 1))
          onPickCell?.(fieldName, summary(r), 'ref', libNodeId)
        }}
      />
    </div>
  )
}
