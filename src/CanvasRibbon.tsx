import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react'
import type { NodeKind } from './core/schema'
import { IcSun, IcMoon, IcImage, IcMagnet } from './ui/icons'
import { SuiteCenter } from './ui/SuiteCenter'
import { DEFAULT_BGS, readCanvasBg, saveCanvasBg } from './ui/canvasBg'
import { prefGet, prefSet } from './ui/prefs'

/**
 * 画布顶部 UI —— OnlyOffice（Word/Excel）经典三段式：
 *   ① 标题行（品牌 + 文档名 + 右侧状态/缩放/回收站）
 *   ② 菜单标签行（文件 / 开始 / 插入 / 视图，再点一次收起）
 *   ③ 工具带（按标签切换的分组按钮）
 *
 * 设计约束（沿用画布既有约定）：
 * - 根节点保留 `.canvas-topbar` 类名 —— Canvas 的指针守卫（onPointerDown / 上下文菜单）
 *   靠 `closest('.canvas-topbar')` 排除本区域，类名不能改。
 * - 只做「摆放与分组」，所有行为都透传 Canvas 既有 handler，不复制业务逻辑。
 */

type TabKey = 'file' | 'home' | 'insert' | 'view'

const TABS: { key: TabKey; label: string }[] = [
  { key: 'file', label: '文件' },
  { key: 'home', label: '开始' },
  { key: 'insert', label: '插入' },
  { key: 'view', label: '视图' },
]

/** 插入页可建的节点（kind 集合与 core/schema.ts 的 NodeKind 保持一致） */
const INSERT_KINDS: { kind: NodeKind; label: string; hint: string }[] = [
  { kind: 'webmon', label: '网址监控', hint: '轮询网页/接口，提取数据记录变化' },
  { kind: 'xlsx2dxf', label: 'Excel转DXF', hint: '坐标表展点成图（点/标注/连线），生成 DXF' },
  { kind: 'start', label: '开始', hint: '手动 / 定时 / Webhook 触发' },
  { kind: 'table', label: '数据表', hint: '元数据表，可增字段' },
  { kind: 'map', label: '数据对照表', hint: '按键匹配给下游补列' },
  { kind: 'text', label: '文本', hint: '单值文本，可被引用' },
  { kind: 'docx', label: 'Word 文档', hint: '模板域 w:sdt 回写' },
  { kind: 'xlsx', label: 'Excel 工作簿', hint: '命名区域 / 表格作域' },
  { kind: 'ooword', label: 'OO 文档', hint: '内嵌 OnlyOffice（Word 内核）' },
  { kind: 'oosheet', label: 'OO 工作簿', hint: '内嵌 OnlyOffice（Excel 内核）' },
  { kind: 'chapter', label: '章节工作簿', hint: 'Excel 全量 + 标题分级与目录大纲' },
  { kind: 'ai', label: 'AI 节点', hint: '圆形 · 普通节点下1出线连它，接收数据做 AI 处理' },
  { kind: 'paramlib', label: '参数库', hint: '圆形 · 出线连普通节点上1口，为它供参数' },
]

/** 16px 线性图标，统一 currentColor 跟随主题 */
const ICONS: Record<string, string> = {
  plus: 'M11 5h2v6h6v2h-6v6h-2v-6H5v-2h6z',
  minus: 'M5 11h14v2H5z',
  fit: 'M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5',
  trash: 'M9 3h6l1 2h4v2H4V5h4zM6 8h12l-1 12H7z',
  copy: 'M8 3h9v3h3v12H8zM5 6H3v15h11v-2',
  paste: 'M9 3h6v3H9zM5 5h3v3h8V5h3v16H5z',
  del: 'M6 7h12l-1 13H7zM9 7V4h6v3M10 10v7M14 10v7',
  layout: 'M4 5h16v4H4zM4 11h7v8H4zM13 11h7v8h-7z',
  /* 自动对齐：网格 + 左侧一条竖直基准线 + 三个右对齐的小横杠。
     与 layout（自动排列：三块面板）刻意不同形 —— 两个按钮语义不同，
     图标长得像会让人以为是同一个功能的两个入口。 */
  align: 'M3 4v16M7 7h10v3H7zM7 14h6v3H7z',
  undo: 'M9 7H5V3M5.2 7.2A8 8 0 1 1 4 12',
  redo: 'M15 7h4V3M18.8 7.2A8 8 0 1 0 20 12',
  /* 折叠/展开：chevron。收起时朝上（点它把条放出来），展开时朝下（点它把条收走） */
  collapse: 'M5 15l7-7 7 7',
  expand: 'M5 9l7 7 7-7',
  /* 节点折叠：四角括号 + 中心实心小方块 —— 「整块内容向中心收成一个方块」 */
  foldnode: 'M4 4h3M4 4v3M20 4h-3M20 4v3M4 20h3M4 20v-3M20 20h-3M20 20v-3M9 9.5h6v5H9z',
  marquee: 'M4 4h3M4 4v3M20 4h-3M20 4v3M4 20h3M4 20v-3M20 20h-3M20 20v-3M9 4h6M9 20h6M4 9v6M20 9v6',
  play: 'M7 4l12 8-12 8z',
  page: 'M6 3h8l4 4v14H6zM14 3v5h5',
  sheet: 'M4 5h16v14H4zM4 10h16M9 5v14',
  rename: 'M4 20h4L20 8l-4-4L4 16z',
  /* 诊断日志开关（2026-10-22 用户「下面的按钮放到上面，尺寸与上面协调」）：
     圆圈 + i，与 OO 节点标题栏诊断抽屉 ⓘ 同语义 */
  info: 'M12 21a9 9 0 1 1 0-18 9 9 0 0 1 0 18zM12 11v5M12 7.6v.5',
  /* 全局特效开关（2026-10-23 用户「右上角增加一个全局特效设置按钮」）：
     四角星 + 小星（sparkle） */
  fx: 'M12 3l2 5.5L19.5 10.5 14 12.5 12 18l-2-5.5L4.5 10.5 10 8.5zM18.6 15.4l.8 2.2 2.2.8-2.2.8-.8 2.2-.8-2.2-2.2-.8 2.2-.8z',
}

function Ico({ name, size = 16 }: { name: keyof typeof ICONS | string; size?: number }) {
  const d = ICONS[name] || ICONS.page
  return (
    <svg viewBox="0 0 24 24" width={size} height={size} aria-hidden="true">
      <path d={d} fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  )
}

/** ★ 2026-10-23 用户指令「右上角增加一个全局特效设置按钮」：
 *  打开/关闭编辑器的 app 式展开/收缩动效全局开关（localStorage 'uiFx'，默认开）。
 *  与诊断开关同排同款（cr-icon-btn）。 */
function FxToggle() {
  const [on, setOn] = useState(() => localStorage.getItem('uiFx') !== '0')
  return (
    <button
      className={`cr-icon-btn${on ? ' cr-icon-btn-on' : ''}`}
      title={on
        ? '界面特效：开（双击打开节点=app 式展开，关闭=收缩回节点）。点击关闭'
        : '界面特效：关（打开/关闭编辑器无动画）。点击开启'}
      onClick={() => {
        const next = !on
        setOn(next)
        localStorage.setItem('uiFx', next ? '1' : '0')
      }}
    >
      <Ico name="fx" />
    </button>
  )
}

/** ★ 2026-10-22 用户指令「下面的按钮放到上面 和上面的尺寸保持协调」：
    画布底部日志面板（LogDock）的开关从画布右上浮动钮挪进顶部工具带，
    用 cr-icon-btn 同款尺寸样式；开合经自定义事件 logdock-toggle 通知
    LogDock，状态持久 localStorage 'logdockOpen'（默认开）。 */
function LogDockToggle() {
  const [on, setOn] = useState(() => localStorage.getItem('logdockOpen') !== '0')
  return (
    <button
      className={`cr-icon-btn${on ? ' cr-icon-btn-on' : ''}`}
      title={on
        ? '诊断日志：开（OO 节点诊断与全局错误汇流，显示在画布底部面板）。点击收起'
        : '诊断日志：关（右下角缩略图/缩放面板回到原位）。点击展开'}
      onClick={() => {
        const next = !on
        setOn(next)
        localStorage.setItem('logdockOpen', next ? '1' : '0')
        window.dispatchEvent(new CustomEvent('logdock-toggle', { detail: next }))
      }}
    >
      <Ico name="info" />
    </button>
  )
}

/** ★ 2026-10-09 OO 启动预热开关（用户需求「设置里加个开关」——应用无常设设置页，
    放画布顶栏常显区，与回收站同级）。开 = 应用启动后画布含 OO 节点时后台调
    /warmup（拉 DS + 焐热引擎资源缓存，零转换任务）。状态持久 localStorage 'oo-warmup'。 */
function OoWarmupToggle() {
  const [on, setOn] = useState(() => localStorage.getItem('oo-warmup') !== '0')
  return (
    <button
      className={`cr-icon-btn${on ? ' cr-icon-btn-on' : ''}`}
      title={on
        ? 'OO 启动预热：开（启动时后台焐热文档引擎，首次打开 OO 文档更快）。点击关闭'
        : 'OO 启动预热：关（首个 OO 文档将现场冷加载）。点击开启'}
      onClick={() => {
        const v = on ? '0' : '1'
        localStorage.setItem('oo-warmup', v)
        setOn(!on)
      }}
    >
      <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
        <path d="M2 9c3-3 5 3 8 0s5 3 8 0M2 15c3-3 5 3 8 0s5 3 8 0" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
      </svg>
    </button>
  )
}

/* ★ 主题切换（2026-10-19 用户「主题也加一个按钮在这里」）：
   与 OO 编辑器的主题开关同一存储（localStorage.editorTheme）+ 同一应用方式
   （documentElement.dataset.theme）——CSS 变量随 root 即时切换。 */
/* ★ 背景图选择器（2026-10-19 用户指令「增加背景图片设置 找一批默认图片」）：
   下拉 8 张默认图缩略图 + 无背景；选择即应用（localStorage.canvasBg） */
function BgPicker() {
  const [open, setOpen] = useState(false)
  const [cur, setCur] = useState<string | null>(() => readCanvasBg())
  const pick = (url: string | null) => {
    saveCanvasBg(url)
    setCur(url)
  }
  useEffect(() => {
    if (!open) return
    const close = () => setOpen(false)
    document.addEventListener('click', close)
    return () => document.removeEventListener('click', close)
  }, [open])
  return (
    <span style={{ position: 'relative', display: 'inline-flex' }} onClick={(e) => e.stopPropagation()}>
      <button
        className={`cr-icon-btn${open ? ' cr-icon-btn-on' : ''}${cur ? ' has-bg' : ''}`}
        style={cur ? { border: '1px solid #d29922' } : undefined}
        title="画布背景：从默认图库选择"
        onClick={() => setOpen((v) => !v)}
      >
        <IcImage size={19} />
      </button>
      {open && (
        <div className="bg-pop">
          <div className="bg-pop-h">画布背景</div>
          {(['暗色', '卡通', '商务', '清新', '活力', '夜空'] as const).map((cat) => {
            const items = DEFAULT_BGS.filter((b) => b.cat === cat)
            if (!items.length) return null
            return (
              <div key={cat} className="bg-cat">
                <div className="bg-cat-h">{cat}</div>
                <div className="bg-grid">
                  {items.map((b) => (
                    <button
                      key={b.id}
                      className={`bg-thumb${cur === b.url ? ' on' : ''}`}
                      title={b.name}
                      onClick={() => pick(b.url)}
                    >
                      <img src={b.url} alt={b.name} draggable={false} />
                    </button>
                  ))}
                </div>
              </div>
            )
          })}
          <button className="bg-none" onClick={() => pick(null)}>无背景（恢复默认深色）</button>
        </div>
      )}
    </span>
  )
}

/* ★ 主题选择（2026-10-20 用户「除了白色黑色 增加几个UI主题 比如粉色、商务灰」）：
   data-theme = dark|light（明暗），data-tint = ''|pink|gray（light 之上的色调）。
   应用入口统一走 applyTheme（写 localStorage + root 属性 + 派发事件）。 */
type ThemeKind = 'dark' | 'light' | 'pink' | 'gray'

export function applyTheme(kind: ThemeKind) {
  const theme = kind === 'dark' ? 'dark' : 'light'
  const tint = kind === 'pink' ? 'pink' : kind === 'gray' ? 'gray' : ''
  document.documentElement.dataset.theme = theme
  if (tint) document.documentElement.dataset.tint = tint
  else delete document.documentElement.dataset.tint
  prefSet('editorTheme', theme)
  prefSet('editorTint', tint || null)
  window.dispatchEvent(new CustomEvent('app-theme-change', { detail: theme }))
}

const THEME_OPTIONS: { kind: ThemeKind; label: string; dot: string }[] = [
  { kind: 'dark', label: '黑夜', dot: '#0d1117' },
  { kind: 'light', label: '白天', dot: '#ffffff' },
  { kind: 'pink', label: '粉色', dot: '#f9a8c9' },
  { kind: 'gray', label: '商务灰', dot: '#9aa0a6' },
]

function currentThemeKind(): ThemeKind {
  const tint = prefGet('editorTint')
  if (tint === 'pink' || tint === 'gray') return tint
  return (prefGet('editorTheme') as ThemeKind) || 'dark'
}

function ThemeToggle() {
  const [open, setOpen] = useState(false)
  const [kind, setKind] = useState<ThemeKind>(() => currentThemeKind())
  useEffect(() => {
    /* 启动即应用上次选择（含 tint） */
    applyTheme(currentThemeKind())
  }, [])
  useEffect(() => {
    if (!open) return
    const close = () => setOpen(false)
    document.addEventListener('click', close)
    return () => document.removeEventListener('click', close)
  }, [open])
  const pick = (k: ThemeKind) => {
    setKind(k)
    applyTheme(k)
    setOpen(false)
  }
  return (
    <span style={{ position: 'relative', display: 'inline-flex' }} onClick={(e) => e.stopPropagation()}>
      <button
        className={`cr-icon-btn${open ? ' cr-icon-btn-on' : ''}`}
        title="界面主题：黑夜 / 白天 / 粉色 / 商务灰"
        onClick={() => setOpen((v) => !v)}
      >
        {kind === 'dark' ? <IcMoon size={19} /> : <IcSun size={19} />}
      </button>
      {open && (
        <div className="bg-pop" style={{ width: 150 }}>
          <div className="bg-pop-h">界面主题</div>
          {THEME_OPTIONS.map((o) => (
            <button
              key={o.kind}
              className="bg-none"
              style={{
                display: 'flex', alignItems: 'center', gap: 8, textAlign: 'left',
                marginBottom: 4, color: kind === o.kind ? 'var(--c-blue)' : 'var(--c-text2)',
                fontWeight: kind === o.kind ? 600 : 400,
              }}
              onClick={() => pick(o.kind)}
            >
              <span style={{ width: 14, height: 14, borderRadius: 7, background: o.dot, border: '1px solid var(--c-border)', flexShrink: 0 }} />
              {o.label}
            </button>
          ))}
        </div>
      )}
    </span>
  )
}

/* ★ 主题切换（2026-10-19 用户「主题也加一个按钮在这里」）：
   与 OO 编辑器的主题开关同一存储（localStorage.editorTheme）+ 同一应用方式
   （documentElement.dataset.theme）——CSS 变量随 root 即时切换。 */

/** 工具带分组：上排图标按钮 + 下排组名（OO 的组标题样式） */
function Group({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="cr-group">
      <div className="cr-group-btns">{children}</div>
      <div className="cr-group-label">{label}</div>
    </div>
  )
}

function Tool({
  icon,
  label,
  title,
  on,
  onClick,
}: {
  icon: string
  label: string
  title?: string
  on?: boolean
  onClick: () => void
}) {
  return (
    <button className={`cr-tool${on ? ' cr-tool-on' : ''}`} title={title || label} onClick={onClick}>
      <Ico name={icon} />
      <span>{label}</span>
    </button>
  )
}

export interface CanvasRibbonProps {
  pages: { id: string; name: string }[]
  activePageId: string | null
  zoomPct: number
  stats: { total: number; selected: number }
  marqueeMode: boolean
  trashOpen: boolean
  trashCount: number
  canUndo: boolean
  canRedo: boolean
  needCenter: boolean
  /** 常规节点全局折叠态（store.nodesFolded，瞬态视图态）——按钮 on 显示用 */
  nodesFolded: boolean
  /** ★ 对齐磁吸开关（2026-10-21） */
  alignMagnet?: boolean
  onToggleAlignMagnet?: () => void
  /** 一键折叠/展开全部常规节点（2026-10-06 用户：「点击菜单的一个按钮 都变成折叠状态」） */
  onToggleNodesFold: () => void
  /**
   * 圆形节点（AI/参数库）运行动画状态（store.roundRunState，瞬态三态）——按钮 on 显示用。
   * 2026-10-06 用户：「增加特效的结束的退出效果 测试的时候3秒结束」——
   * 故不是布尔而是三态：'off' | 'running' | 'exiting'。
   * ★ 按钮的 on 高亮只认 running：exiting 是淡出中，已经在「结束」了，
   *   还亮着会让用户以为还能再点。
   */
  roundRunState: 'off' | 'running' | 'exiting'
  /** 当前没选中任何 AI/参数库 —— 播放按钮置灰（判定在 core/roundRun.hasPlayableSelection） */
  selNoneHint: boolean
  /** 播放/停止选中圆形节点的运行动画（时序与定时器在 Canvas 侧，迁移本体在 core/roundRun） */
  onToggleRoundRun: () => void
  onUndo: () => void
  /* ★ 合并标题栏插槽（2026-10-19）：App 的 logo/画布名/版本（左）+ 保存/打开/重启/关闭（右） */
  ribbonSlots?: { current: { left?: ReactNode; right?: ReactNode } } | null
  onRedo: () => void
  onGlobalPreview: () => void
  onSwitchPage: (id: string) => void
  onNewPage: () => void
  onRenamePage: () => void
  onDeletePage: () => void
  onZoomIn: () => void
  onZoomOut: () => void
  onZoomReset: () => void
  onFitView: () => void
  onToggleMarquee: () => void
  onAutoLayout: () => void
  /**
   * 自动对齐（2026-10-06 用户：「增加一个自动对齐按钮」）——
   * ★ 与 onAutoLayout 是**两个独立按钮**（用户：「自动对齐 和 自动排列 两个按钮
   *   一个是对齐 一个是彻底重排」）：
   *     自动对齐 = 网格归位，**保持**相对关系（吸附 + 消重叠 + 拉直同行）
   *     自动排列 = 按上下游分层，**改变**相对关系
   *   作用于**全部**节点，不需要先框选。
   */
  onAutoAlign: () => void
  onRunWorkflow: () => void
  onCopy: () => void
  onPaste: () => void
  onDelete: () => void
  onAddNode: (kind: NodeKind) => void
  onToggleTrash: () => void
}

export function CanvasRibbon(props: CanvasRibbonProps) {
  const [tab, setTab] = useState<TabKey | null>('home')
  /* 功能区折叠（2026-10-06 用户重定语义：「这个不折叠。点击折叠后 只折叠菜单下面的功能区
     折叠状态中 点击选择功能后操作的时候自动折叠」）——菜单标签行（文件/开始/插入/视图）
     永不折叠，folded 只管工具带在不在：
     - folded=true：工具带不渲染（页签行还在，画布只让出工具带那一截高度）
     - 折叠态点页签：tab 置值 → 工具带临时弹出；点了功能按钮/下拉后容器委托置 tab=null → 自动收回
     - 再点折叠按钮（此时 chevron 朝下）→ 恢复常驻 */
  const [folded, setFolded] = useState(false)
  const rootRef = useRef<HTMLDivElement | null>(null)

  // 菜单通栏：把自己的高度写进 --ribbon-h，外层 .layout 用它把画布与竖向节点库一起顶下去。
  // 用 ResizeObserver 而不是写死数值 —— 工具带可折叠（标题行+标签行 ≈ 68px / 展开 ≈ 138px），
  // 页签切换也会改高度，写死会在折叠后留一条空白。
  useLayoutEffect(() => {
    const el = rootRef.current
    if (!el) return
    const apply = () => {
      const h = Math.round(el.getBoundingClientRect().height)
      document.documentElement.style.setProperty('--ribbon-h', `${h}px`)
    }
    apply()
    const ro = new ResizeObserver(apply)
    ro.observe(el)
    return () => {
      ro.disconnect()
      document.documentElement.style.removeProperty('--ribbon-h')
    }
  }, [])
  const {
    pages,
    activePageId,
    zoomPct,
    stats,
    marqueeMode,
    trashOpen,
    trashCount,
    canUndo,
    canRedo,
    needCenter,
    nodesFolded,
    alignMagnet,
    onToggleAlignMagnet,
    onToggleNodesFold,
    roundRunState,
    selNoneHint,
    onToggleRoundRun,
    onUndo,
    onRedo,
    onGlobalPreview,
    onSwitchPage,
    onNewPage,
    onRenamePage,
    onDeletePage,
    onZoomIn,
    onZoomOut,
    onZoomReset,
    onFitView,
    onToggleMarquee,
    onAutoLayout,
    onAutoAlign,
    onRunWorkflow,
    onCopy,
    onPaste,
    onDelete,
    onAddNode,
    onToggleTrash,
  } = props

  // 注：画布名（activeName）不在这里渲染 —— 已与品牌名合并进 app 标题栏一行，
  // 避免同一个名字在界面上出现两次。页面下拉/重命名仍在「文件」页签里。

  return (
    // 类名canvas-topbar 必须保留：Canvas 指针守卫靠它排除本区域
    <div
      ref={rootRef}
      className="canvas-topbar cr-root nodrag nopan nowheel"
      onPointerDown={(e) => e.stopPropagation()}
    >
      {/* ── 折叠态（2026-10-06 重定）：页签行常驻不折叠，只收下面的工具带 ——
          不再需要独立展开入口（页签行本身就带折叠按钮），窄把手方案作废。 ── */}
          <div className="cr-tabrow">
            {/* ★ 合并标题栏（2026-10-19）：logo/画布名/版本 在最左 */}
            {props.ribbonSlots?.current?.left}
            <div className="cr-history">
              <button className="cr-icon-btn" title="撤销 (Ctrl+Z)" disabled={!canUndo} onClick={onUndo}>
                <Ico name="undo" />
              </button>
              <button className="cr-icon-btn" title="重做 (Ctrl+Y)" disabled={!canRedo} onClick={onRedo}>
                <Ico name="redo" />
              </button>
            </div>
            {/* 折叠按钮紧贴撤销/重做：两者都是「对整条界面做操作」，和右侧的缩放/回收站
            （对画布内容做操作）分属两类。只收工具带、页签行不动（用户 2026-10-06）。
            收起时 chevron 朝上、折叠时朝下，方向即状态。 */}
            <button
              className="cr-icon-btn cr-fold-btn"
              title={folded ? '展开功能区' : '折叠功能区（只留菜单行）'}
              onClick={() => {
                if (folded) {
                  setFolded(false)
                  setTab((t) => t ?? 'home') // 展开时恢复原页签；之前已收回则给默认页签，避免展开后空一截
                } else {
                  setFolded(true)
                  setTab(null) // 折叠 = 收工具带；tab 清掉，折叠态下工具带靠「点页签临时弹出」再出现
                }
              }}
            >
              <Ico name={folded ? 'expand' : 'collapse'} />
            </button>
          {TABS.map((t) => (
            <button
              key={t.key}
              className={`cr-tab${tab === t.key ? ' cr-tab-on' : ''}`}
              onClick={() => setTab((cur) => (cur === t.key ? null : t.key))}
              title={`${t.label}（再点一次收起工具带）`}
            >
              {t.label}
            </button>
          ))}
          <div className="cr-right">
            {/* ★ 2026-10-24 用户指令「三个放缩放比例上面」：三胶囊迁往右下角 zoom-panel 上方 */}
            <button
              className={`cr-icon-btn${trashOpen ? ' cr-icon-btn-on' : ''}`}
              title={`回收站${trashCount ? `（${trashCount} 项）` : ''}`}
              onClick={onToggleTrash}
            >
              <Ico name="trash" />
            </button>
            <button
              className={`cr-icon-btn${alignMagnet ? ' cr-icon-btn-on' : ''}`}
              title="对齐磁吸：拖动节点时自动吸附对齐线（点击开/关）"
              onClick={onToggleAlignMagnet}
            >
              <IcMagnet size={19} />
            </button>
            <BgPicker />
            <ThemeToggle />
            <OoWarmupToggle />
            {/* ★ 2026-10-26 19:08 用户反馈「日志按钮去哪了」：套件中心钮原放这里把 cr-right
                挤宽、尾部日志/特效钮被右侧核心钮组盖住——挪进标题栏 ⋯ 折叠池（App.tsx
                actionButtons 'suites'，低频钮正确归宿）。SuiteCenter 弹窗挂载保留在本组件。 */}
            <SuiteCenter />
            {/* ★ 节点折叠按钮放**顶栏常显区**（2026-10-06 用户「没看到按钮」——
                初版放在视图页签工具带里，藏得深找不着。顶栏任何页签下都在，
                与回收站同级：两者都是「对画布内容做一键操作」。 */}
            <button
              className={`cr-icon-btn${nodesFolded ? ' cr-icon-btn-on' : ''}`}
              title={
                nodesFolded
                  ? '展开节点：所有常规节点恢复完整卡片'
                  : '折叠节点：所有常规节点收成正方形（再点恢复）'
              }
              onClick={onToggleNodesFold}
            >
              <Ico name="foldnode" />
            </button>
            {/* ★ 圆形节点运行动画测试按钮（2026-10-06 用户两条：
                1「对选中的节点点击动态特效才会有效果」
                2「增加特效的结束的退出效果 测试的时候3秒结束」）
                —— 放顶栏常显区紧挨节点折叠，两者都是「对画布内容做一键操作」。

                ★ 只对**选中**的圆形节点生效（旧实现是全画布圆形节点一起转圈）。
                ★ 退出中（exiting）不高亮：已经在淡出了，还亮着会让人以为能再点。
                ★ 未选中任何 AI/参数库 时 disabled —— 点了不会有任何反应，
                  与其让人白按一次，不如直接置灰并在 title 里说明怎么用。 */}
            <button
              className={`cr-icon-btn${roundRunState === 'running' ? ' cr-icon-btn-on' : ''}`}
              title={
                roundRunState === 'running'
                  ? '停止运行动画（淡出退出）'
                  : roundRunState === 'exiting'
                    ? '运行动画正在淡出退出…'
                    : selNoneHint
                      ? '请先点选一个 AI / 参数库 节点，再点此按钮播放运行动画'
                      : '播放运行动画：选中节点转圈 3 秒后淡出退出（测试档）'
              }
              disabled={selNoneHint}
              onClick={onToggleRoundRun}
            >
              <Ico name="play" />
            </button>
            {/* ★ 诊断日志开关（2026-10-22）：画布底部 LogDock 的开关挪进工具带，
                与本排图标钮同尺寸同风格 */}
            <LogDockToggle />
            {/* ★ 全局特效开关（2026-10-23 用户指令）：编辑器打开/收缩动效总开关 */}
            <FxToggle />
          </div>
          {/* ★ 合并标题栏（2026-10-19）：保存/打开/重启/关闭 在最右 */}
          {props.ribbonSlots?.current?.right}
        </div>

        {/* 工具带。折叠态（folded）下这里是「临时弹出」：点功能按钮/下拉选完即自动收回——
            容器委托监听而非逐个 Tool 注入，以后新增功能按钮零维护（2026-10-06 用户指令）。 */}
        {tab && (
          <div
            className="cr-ribbon"
            onClick={(e) => {
              if (!folded) return
              const el = e.target as HTMLElement
              if (el.closest('button, select')) setTab(null)
            }}
          >
            {tab === 'file' && (
              <>
                <Group label="画布页">
                  <div className="cr-combo">
                    <select
                      className="cr-page-select"
                      value={activePageId ?? ''}
                      onChange={(e) => onSwitchPage(e.target.value)}
                      title="切换画布页"
                    >
                      {pages.map((p) => (
                        <option key={p.id} value={p.id}>
                          {p.name}
                        </option>
                      ))}
                    </select>
                  </div>
                  <Tool icon="plus" label="新建" title="新建画布页" onClick={onNewPage} />
                  <Tool icon="rename" label="重命名" title="重命名当前画布页" onClick={onRenamePage} />
                  <Tool icon="trash" label="删除" title="删除当前画布页" onClick={onDeletePage} />
                </Group>
                <Group label="数据">
                  <Tool
                    icon="trash"
                    label="回收站"
                    title="打开回收站（已删除节点可还原）"
                    on={trashOpen}
                    onClick={onToggleTrash}
                  />
                  <Tool icon="copy" label="复制选中" title="复制选中节点 (Ctrl+C)" onClick={onCopy} />
                  <Tool icon="paste" label="粘贴" title="粘贴节点 (Ctrl+V)" onClick={onPaste} />
                  <Tool icon="del" label="删除选中" title="删除选中节点 (Delete)" onClick={onDelete} />
                </Group>
              </>
            )}

            {tab === 'home' && (
              <>
                <Group label="剪贴板">
                  <Tool icon="copy" label="复制" title="复制选中节点 (Ctrl+C)" onClick={onCopy} />
                  <Tool icon="paste" label="粘贴" title="粘贴节点 (Ctrl+V)" onClick={onPaste} />
                  <Tool icon="del" label="删除" title="删除选中节点 (Delete)" onClick={onDelete} />
                </Group>
                <Group label="排列与运行">
                  {/* ★ 两个按钮并存、语义不同（2026-10-06 用户：「自动对齐 和 自动排列
                      两个按钮 一个是对齐 一个是彻底重排」）。title 必须说清区别 ——
                      用户分不清就会连点，点「自动对齐」发现上下游全打乱就以为坏了。 */}
                  <Tool
                    icon="align"
                    label="自动对齐"
                    title="自动对齐：全部节点吸附到网格、消除重叠、同行拉齐（保持相对位置不变）"
                    onClick={onAutoAlign}
                  />
                  <Tool
                    icon="layout"
                    label="自动排列"
                    title="自动排列：按上下游关系彻底重排为左→右分层（会改变相对位置）"
                    onClick={onAutoLayout}
                  />
                  <Tool
                    icon="marquee"
                    label="框选"
                    title="框选模式：开启后左键拖拽框选节点（再次点击关闭）"
                    on={marqueeMode}
                    onClick={onToggleMarquee}
                  />
                  <Tool icon="play" label="运行" title="手动执行工作流（按数据流高亮执行顺序）" onClick={onRunWorkflow} />
                </Group>
                <Group label="视图">
                  <Tool icon="fit" label="适应窗口" title="适应画布（最大 100%）" onClick={onFitView} />
                  <Tool icon="minus" label="缩小" title="缩小" onClick={onZoomOut} />
                  <Tool icon="plus" label="放大" title="放大" onClick={onZoomIn} />
                </Group>
              </>
            )}

            {tab === 'insert' && (
              <Group label="插入节点（落在视口中心）">
                {INSERT_KINDS.map((k) => (
                  <Tool
                    key={k.kind}
                    icon={k.kind === 'docx' || k.kind === 'ooword' ? 'page' : k.kind === 'xlsx' || k.kind === 'oosheet' || k.kind === 'chapter' ? 'sheet' : 'plus'}
                    label={k.label}
                    title={`${k.label} —— ${k.hint}`}
                    onClick={() => onAddNode(k.kind)}
                  />
                ))}
              </Group>
            )}

            {tab === 'view' && (
              <>
                <Group label="缩放">
                  <Tool icon="minus" label="缩小" title="缩小" onClick={onZoomOut} />
                  <button className="cr-zoom-val" title="点击恢复 100%" onClick={onZoomReset}>
                    {zoomPct}%
                  </button>
                  <Tool icon="plus" label="放大" title="放大" onClick={onZoomIn} />
                  <Tool icon="fit" label="适应窗口" title="适应画布（最大 100%）" onClick={onFitView} />
                </Group>
                <Group label="画布交互">
                  <Tool
                    icon="fit"
                    label="全局预览"
                    title="将画布内容居中（内容偏离中心时高亮）"
                    on={needCenter}
                    onClick={onGlobalPreview}
                  />
                  {/* 节点折叠按钮已上移到顶栏常显区（cr-right，回收站旁）——
                      这里不再重复放：同一动作两个入口会让「on 态」出现两处，
                      状态一致性多一份维护成本，且用户口径就是「菜单的一个按钮」。 */}
                  <Tool
                    icon="marquee"
                    label="框选"
                    title="框选模式：开启后左键拖拽框选节点"
                    on={marqueeMode}
                    onClick={onToggleMarquee}
                  />
                  <Tool icon="align" label="自动对齐" title="自动对齐：全部节点吸附到网格、消除重叠、同行拉齐（保持相对位置）" onClick={onAutoAlign} />
                  <Tool icon="layout" label="自动排列" title="自动排列：按上下游关系彻底重排（改变相对位置）" onClick={onAutoLayout} />
                  <Tool icon="play" label="运行工作流" title="手动执行工作流" onClick={onRunWorkflow} />
                </Group>
              </>
            )}
          </div>
        )}
    </div>
  )
}
