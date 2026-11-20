import { useCallback, useEffect, useMemo, useRef, useState, type MouseEvent as ReactMouseEvent, type ReactNode } from 'react'
import type { XYPosition } from '@xyflow/react'
import { useBaseStore } from './core/store'
import type { NodePos } from './core/store'
import { computeNode } from './core/engine'
import { defineField, defineNode, aiPrefsPatchOf, type NodeKind, type TableSchema, type FieldDef, type NodeConfig, type NodeRuleSet } from './core/schema'
import { parseFile, exportRows } from './core/io'
import { PALETTE_GROUPS, paletteItemClassOf, styleOf, isMergeKind, isRoundNodeKind } from './core/nodeStyle'
import { ParamlibPanel } from './ParamlibPanel'
import { ruleConnectPatch, canJoinRuleGroup } from './core/ruleSync'
import { MATERIALS, BILL, MATERIALS_SEED, BILL_SEED } from './core/seed'
import { Canvas, type CopyItem } from './Canvas'
import { PopupNodeEditor, hasPopupForm } from './PopupNodeEditor'
import { IcDatabase, IcFolderOpen, IcRotateCcw, IcX, IcSaveAs, IcClear, IcGrid, IcChip, IcMove, IcGlobe, IcMaximize, IcMinimize, IcPuzzle } from './ui/icons'
import { suiteCenterUI } from './core/suites'
import { prefGet, prefSet } from './ui/prefs'
import { TableForm } from './TableForm'
import { PropertiesPanel } from './PropertiesPanel'
import { NodeEditor } from './NodeEditor'
import { RuleMembersPanel } from './RuleMembersPanel'
import { saveCanvasAsWfc, collectAssetManifest, type WfcSaveStep, type AssetManifest } from './wfcSave'
import { openCanvasFromWfcBytes, openCanvasFromWfcFile } from './wfcOpen'
import { audit } from './core/audit'
import { LogDock } from './ui/LogDock'
import { installGlobalLogHooks, pushLog } from './ui/logBus'
import { APP_VERSION, APP_CHANNEL, APP_BUILD, APP_VERSION_LABEL } from './core/version'

/** ★ 2026-11-07 对话「保存画布文件」→ 保存链路的阶段回执（对话侧据此回填气泡）。
 *  stage：cancel 用户取消 / open 已调起保存面板（等点确认）/ done 已落盘 / error 失败 */
function reportChatSave(ok: boolean, stage: 'cancel' | 'open' | 'done' | 'error', msg: string) {
  window.dispatchEvent(new CustomEvent('chat-save-result', { detail: { ok, stage, msg } }))
}

// 默认名自动编号：拖入同类型节点时，若已存在同名默认节点则追加 (2)、(3)...
function uniqueName(base: string, tables: Record<string, TableSchema>): string {
  const taken = new Set(Object.values(tables).map((t) => t.name))
  if (!taken.has(base)) return base
  let i = 2
  while (taken.has(`${base} (${i})`)) i++
  return `${base} (${i})`
}

// Electron preload 暴露的重启通道（preload.cjs relaunch → main app:relaunch）。
// 浏览器跑 dev:web 时 window.api 不存在 → undefined，标题栏重启按钮随之隐藏。
const electronRelaunch = (window as unknown as { api?: { relaunch?: () => Promise<void> } }).api?.relaunch
// Electron preload 暴露的彻底退出通道（preload.cjs quit → main app:quit）。
const electronQuit = (window as unknown as { api?: { quit?: () => Promise<void> } }).api?.quit

export default function App() {
  const ready = useBaseStore((s) => s.ready)
  const init = useBaseStore((s) => s.init)
  const tables = useBaseStore((s) => s.tables)
  const pages = useBaseStore((s) => s.pages)
  const activePageId = useBaseStore((s) => s.activePageId)
  /** 当前画布页名（标题栏与画布名合并显示用） */
  const activePageName = pages.find((p) => p.id === activePageId)?.name ?? '未命名画布'
  /** 重命名当前画布页（标题栏画布名点击触发） */
  const renameActivePage = useCallback(() => {
    if (!activePageId) return
    const name = window.prompt('重命名画布', activePageName)
    if (name) useBaseStore.getState().renamePage(activePageId, name)
  }, [activePageId, activePageName])

  /* 文档标题 = 品牌 + 当前画布名。
     为什么必须同步 document.title：宿主（WorkBuddy 预览面板、浏览器标签页、Electron
     窗口）显示的那一行标题读的是 document.title，而它原先在 index.html 里写死成
     「客户端底座 client-base」—— 换画布后宿主标题纹丝不动，和应用内标题栏
     （.app-docname，已显示画布名）对不上，同一屏出现两个不一致的名字。 */
  useEffect(() => {
    document.title = activePageName
  }, [activePageName])
  /* ★ 全局日志钩子（2026-10-22 用户指令「画布下方增加log显示面板」）：
     window.onerror / unhandledrejection 汇入 logBus，供底部 LogDock 展示 */
  useEffect(() => { installGlobalLogHooks() }, [])
  /* ★ 套件中心状态初始化（2026-10-26）：读停用表——节点能力闸门依赖它 */
  useEffect(() => { void import('./core/suites').then((m) => m.initSuites()) }, [])
  useEffect(() => {
    const foldOff = () => setFoldMenu(false)
    window.addEventListener('cb-drag-start', foldOff)
    document.addEventListener('pointerdown', (e) => {
      const t = e.target as HTMLElement | null
      if (foldMenuRef.current && !foldMenuRef.current.contains(t) && !(t?.closest?.('[title*="更多操作"]'))) setFoldMenu(false)
    })
    const on = () => { setAnyDragging(true); document.body.classList.add('cb-dragging') }
    const off = () => { setAnyDragging(false); document.body.classList.remove('cb-dragging') }
    window.addEventListener('cb-drag-start', on)
    window.addEventListener('cb-drag-end', off)
    /* ★ 保存耗时实时刷新 */
    const si = window.setInterval(() => {
      if (!savingWfcRef.current) return
      setSaveElapsed((Date.now() - saveStartRef.current) / 1000)
    }, 100)
    saveElapsedTimerClear = () => window.clearInterval(si)
    return () => { window.removeEventListener('cb-drag-start', on); window.removeEventListener('cb-drag-end', off); document.body.classList.remove('cb-dragging') }
  }, [])
  /* ★ 2026-10-24 版本戳：当前渲染层构建指纹进日志——用户与我共同可验证 */
  useEffect(() => {
    try {
      const el = document.querySelector('script[src*="assets/index-"]') as HTMLScriptElement | null
      const src = el?.src || location.href
      const m = src.match(/index-([A-Za-z0-9_-]+)\.js/)
      pushLog('[版本] 渲染层构建指纹: ' + (m ? m[1].slice(0, 8) : '未知'), 'ok')
    } catch { /* ignore */ }
  }, [])
  const records = useBaseStore((s) => s.records)
  const positions = useBaseStore((s) => s.positions)
  const addTable = useBaseStore((s) => s.addTable)
  const addField = useBaseStore((s) => s.addField)
  const updateField = useBaseStore((s) => s.updateField)
  const removeField = useBaseStore((s) => s.removeField)
  const renameTable = useBaseStore((s) => s.renameTable)
  const duplicateTable = useBaseStore((s) => s.duplicateTable)
  const removeTable = useBaseStore((s) => s.removeTable)
  const trashTable = useBaseStore((s) => s.trashTable)
  const trashTables = useBaseStore((s) => s.trashTables)
  const restoreTable = useBaseStore((s) => s.restoreTable)
  const restoreLastTrash = useBaseStore((s) => s.restoreLastTrash)
  const purgeTable = useBaseStore((s) => s.purgeTable)
  const emptyTrash = useBaseStore((s) => s.emptyTrash)
  const setNodePosition = useBaseStore((s) => s.setNodePosition)
  const setNodePositions = useBaseStore((s) => s.setNodePositions)
  const setNodeConfig = useBaseStore((s) => s.setNodeConfig)
  const setRule = useBaseStore((s) => s.setRule)
  const clearRule = useBaseStore((s) => s.clearRule)
  const unlinkFromRuleNode = useBaseStore((s) => s.unlinkFromRuleNode)
  const setRulePrimary = useBaseStore((s) => s.setRulePrimary)
  const setRuleFollow = useBaseStore((s) => s.setRuleFollow)
  const setTableKind = useBaseStore((s) => s.setTableKind)
  const setTableData = useBaseStore((s) => s.setTableData)
  const addRow = useBaseStore((s) => s.addRow)
  const removeRow = useBaseStore((s) => s.removeRow)
  const setCell = useBaseStore((s) => s.setCell)

  // 撤销：每次写操作前先快照当前状态入栈（store.undo 可回退到该快照）
  const push = () => useBaseStore.getState().pushHistory()

  const [selected, setSelected] = useState<string>('')
  const [rightTab, setRightTab] = useState<'schema' | 'data'>('schema')
  const [propsWidth, setPropsWidth] = useState(320)
  /* ★ 属性栏悬浮框顶部让位：量功能区(.cr-ribbon)高度写入 --ribbon-h（500ms 低频轮询覆盖展开/收起） */
  useEffect(() => {
    const upd = () => {
      const rib = document.querySelector('.cr-ribbon')
      const h = rib ? rib.getBoundingClientRect().height : 0
      document.documentElement.style.setProperty('--ribbon-h', Math.max(56, Math.round(h)) + 'px')
    }
    upd()
    const iv = setInterval(upd, 500)
    return () => clearInterval(iv)
  }, [])
  /** 双击 docx/xlsx 节点后打开的放大态编辑器；空串表示关闭 */
  const [editorNode, setEditorNode] = useState<string>('')
  /** 画布展开/折叠：隐藏左侧节点库，画布占满（C 端常用的「专注模式」） */
  const [paletteCollapsed, setPaletteCollapsed] = useState(false)
  /* ★ 2026-10-23 用户指令「增加一个展开参数库的侧边 和展开节点库并列」：
     参数库侧板独立开关（折叠条与节点库按钮并列；面板列出内置参数库条目） */
  const [paramlibOpen, setParamlibOpen] = useState(false)
  /** 双击圆形「规则跟随」节点后打开的成员管理面板；空串表示关闭
   *  （用户：「规则节点上显示几个节点。且双击能修改节点的主次、跟随不跟随。」） */
  const [ruleMembersNode, setRuleMembersNode] = useState('')
  /** .wfc 保存进行中（防重复点击；按钮同步显示「保存中…」） */
  const [savingWfc, setSavingWfc] = useState(false)
  /* ★ 2026-10-24 保存两段流：先出文件清单（用户确认）→ 再存储进度 */
  const [saveManifest, setSaveManifest] = useState<AssetManifest | null>(null)
  /* ★ 2026-10-25 用户令「保存后加一个折叠打开的按钮」：完成态不关弹窗——
     一条完成条（默认折叠），展开/收起清单；problems 也在弹窗内可见 */
  const [saveDone, setSaveDone] = useState<{ fileName: string; kb: string; problems: string[]; dedup: number } | null>(null)
  const [saveDoneOpen, setSaveDoneOpen] = useState(false)
  const [saveAwaitName, setSaveAwaitName] = useState('')
  /* ★ 历史记录文件是否存入 .wfc（默认不勾——用户指令） */
  /* ★ 2026-10-25 用户令「历史文件写明属于哪个节点 可展开勾选 或选择最新版本」：
     细粒度树形勾选替代全局开关——key = `${nodeId}:${hId}`；默认全选 */
  const [histSel, setHistSel] = useState<Record<string, boolean>>({})
  const [histOpen, setHistOpen] = useState<Record<string, boolean>>({})
  /* ★ 孤立节点可选不保存（用户令）：key=nodeId（true=本次保存剔除，默认 false=保存） */
  const [orphanSkip, setOrphanSkip] = useState<Record<string, boolean>>({})
  const [orphanOpen, setOrphanOpen] = useState<Record<string, boolean>>({})
  /* ★ 文件清单节点卡片折叠（用户令「全部节点也增加折叠的能力」）——默认展开 */
  const [filesOpen, setFilesOpen] = useState<Record<string, boolean>>({})
  /* ★ 保存面板 tab：files 清单 / orphan 孤立节点 / history 历史节点 */
  const [saveTab, setSaveTab] = useState<'files' | 'orphan' | 'history'>('files')
  /* ★ 2026-10-24 保存进度条（用户指令「保存进度条 每一步展示」） */
  const [saveStep, setSaveStep] = useState<{ label: string; pct: number } | null>(null)
  /* ★ 2026-10-25 保存耗时显示（用户指令「从点击保存到正式发布 提示个时间」） */
  const [saveElapsed, setSaveElapsed] = useState(0)
  const saveStartRef = useRef(0)
  const savingWfcRef = useRef(false)
  let saveElapsedTimerClear: (() => void) | null = null
  /* ★ 2026-10-24 用户指令「视图菜单里增加设置选项」：顶部 ribbon「视图」下拉——
     画布设置（连线动画三态），localStorage 共享 + 事件同步给 Canvas */
  const [viewMenuOpen, setViewMenuOpen] = useState(false)
  const [edgeAnimUi, setEdgeAnimUi] = useState<'pause' | 'always' | 'off'>(() => {
    try { return (localStorage.getItem('edgeAnimMode') as 'pause' | 'always' | 'off') || 'pause' } catch { return 'pause' }
  })
  const [dragLodUi, setDragLodUi] = useState<'wysiwyg' | 'outline'>(() => {
    try { return prefGet('dragLod') === 'outline' ? 'outline' : 'wysiwyg' } catch { return 'wysiwyg' }
  })
    const [magnetUi, setMagnetUi] = useState<boolean>(() => {
    try { return prefGet('alignMagnet') === '1' } catch { return false }
  })
  const [gpuOn, setGpuOn] = useState<boolean | null>(null)
  useEffect(() => {
    const api = (window as unknown as { api?: { gpuGetAccel?: () => Promise<boolean> } }).api
    api?.gpuGetAccel?.().then((v) => setGpuOn(v)).catch(() => {})
  }, [])
  const toggleGpu = async () => {
    /* ★ 2026-10-24 用户反馈「不能点」：点击立即变 UI（乐观更新）+ try/catch——
       主进程/preload 旧（无 handler）时报错可见，不再静默无反应 */
    const nv = !(gpuOn ?? true)
    setGpuOn(nv)
    pushLog('[gpu] 硬件加速 → ' + (nv ? '开' : '关') + '（重启应用后生效）', nv ? 'ok' : 'warn')
    try {
      const api = (window as unknown as { api?: { gpuSetAccel?: (on: boolean) => Promise<boolean> } }).api
      if (!api?.gpuSetAccel) { pushLog('[gpu] 桌面桥不可用——请完全退出应用后重新启动（主进程需加载新版）', 'warn'); setGpuOn(!nv); return }
      await api.gpuSetAccel(nv)
    } catch (eT) {
      pushLog('[gpu] 写入失败: ' + String(eT).slice(0, 80) + '——请完全退出应用后重新启动', 'warn')
      setGpuOn(!nv)
    }
  }
    /* ★ 2026-10-24 用户指令「网页打开 增加全屏按钮」：Web 模式沉浸式 */
  const [isFull, setIsFull] = useState(false)
  /* ★ 2026-10-24 全局拖动态（事件桥）：任意节点拖动中 → OO 编辑器 iframe 降级（拖动卡根因） */
  const [anyDragging, setAnyDragging] = useState(false)
  /* ★ 2026-10-24 动作区溢出折叠：foldVisible=可见的可折叠按钮数，窗口变窄时递减
     （左侧先收进 ⋯ 菜单），余量恢复；检测基准=窗口右缘（动作区 shrink:0 尺寸恒定） */
  const [foldVisible, setFoldVisible] = useState(6)
  const [foldMenu, setFoldMenu] = useState(false)
  const foldMenuRef = useRef<HTMLDivElement | null>(null)
  const foldRef = useRef<HTMLDivElement | null>(null)
  const foldVcRef = useRef(foldVisible)
  useEffect(() => {
    const check = () => {
      const el = foldRef.current
      if (!el) return
      const rect = el.getBoundingClientRect()
      const vw = window.innerWidth
      const overflowPx = rect.right - (vw - 8)   /* 容差 8px：贴合视口即触发（1035px 死区实测修正） */
      const maxFold = 6
      let vc = foldVcRef.current
      if (overflowPx > 0) {
        const need = Math.ceil(overflowPx / 37)
        vc = Math.max(1, vc - need)
        if (vc !== foldVcRef.current) { foldVcRef.current = vc; setFoldVisible(vc) }
      } else if (overflowPx < -40 && vc < maxFold) {
        const back = Math.min(maxFold - vc, Math.ceil((-overflowPx - 40) / 37))
        if (back > 0) { vc += back; foldVcRef.current = vc; setFoldVisible(vc) }
      }
    }
    /* ★ 2026-10-24 修复过折叠：resize 拖动中每帧触发连续递减——节流 150ms，
       停手后按最终宽度一次折叠到位（防拖快了按钮被「删」太多） */
    let debT: ReturnType<typeof setTimeout> | undefined
    const debounced = () => { if (debT) clearTimeout(debT); debT = setTimeout(check, 150) }
    window.addEventListener('resize', debounced)
    const t = setTimeout(check, 300)
    return () => { window.removeEventListener('resize', debounced); clearTimeout(debT); clearTimeout(t) }
  }, [])
  /* ★ 2026-10-24 用户「窗体变小 从左侧开始隐藏 右侧不遮挡」：ResizeObserver 溢出检测，
     按版本→文档名→logo 顺序渐进隐藏左侧，有余量反向恢复（滞回防抖动） */
  const brandRef = useRef<HTMLDivElement | null>(null)
  const [hideLevel, setHideLevel] = useState(0)
  const hideLevelRef = useRef(0)
  const checkBrand = () => {
    const el = brandRef.current
    if (!el) return
    const actions = el.querySelector('.app-titlebar-actions') as HTMLElement | null
    if (!actions) return
    const overflow = actions.getBoundingClientRect().right > el.getBoundingClientRect().right - 2
    const roomy = actions.getBoundingClientRect().right < el.getBoundingClientRect().right - 60
    let lv = hideLevelRef.current
    if (overflow && lv < 3) { lv++; hideLevelRef.current = lv; setHideLevel(lv) }
    else if (roomy && lv > 0) { lv--; hideLevelRef.current = lv; setHideLevel(lv) }
  }
  useEffect(() => {
    const el = brandRef.current
    if (!el) return
    const ro = new ResizeObserver(checkBrand)
    ro.observe(el)
    const t = setTimeout(checkBrand, 300)
    return () => { ro.disconnect(); clearTimeout(t) }
  }, [])
  useEffect(() => { const t = setTimeout(checkBrand, 60); return () => clearTimeout(t) }, [hideLevel])
  useEffect(() => {
    const h = () => setIsFull(!!document.fullscreenElement)
    document.addEventListener('fullscreenchange', h)
    return () => document.removeEventListener('fullscreenchange', h)
  }, [])
  const toggleFull = () => {
    const d = document as Document & { webkitExitFullscreen?: () => void }
    const el = document.documentElement as HTMLElement & { webkitRequestFullscreen?: () => void }
    if (document.fullscreenElement) { (d.exitFullscreen ? d.exitFullscreen() : d.webkitExitFullscreen?.()) }
    else if (el.requestFullscreen) el.requestFullscreen()
    else if (el.webkitRequestFullscreen) el.webkitRequestFullscreen()
  }
    const [perfHudUi, setPerfHudUi] = useState<boolean>(() => {
    try { return prefGet('perfHud') === '1' } catch { return false }
  })
    const setMagnet = (on: boolean) => {
    try { prefSet('alignMagnet', on ? '1' : '0') } catch { /* ignore */ }
    window.dispatchEvent(new CustomEvent('align-magnet-changed', { detail: on ? '1' : '0' }))
    setMagnetUi(on)
  }
    const setEdgeAnim = (m: 'pause' | 'always' | 'off') => {
    try { localStorage.setItem('edgeAnimMode', m) } catch { /* ignore */ }
    window.dispatchEvent(new CustomEvent('edge-anim-changed', { detail: m }))
    setEdgeAnimUi(m)
  }
  /* ★ 合并标题栏插槽（2026-10-19）：原 .app-titlebar 内容改由 Canvas→CanvasRibbon 页签行渲染 */
  const ribbonSlots = useRef<{ left?: ReactNode; right?: ReactNode }>({})
  /** ★ 弹窗形态（2026-10-19）：双击常规数据节点打开的居中弹窗编辑器 */
  const [popupNode, setPopupNode] = useState('')
  /** 保存当前画布全部内容为 .wfc 工程文件（2026-10-06 用户：
      「点击保存能保存当前画布所有内容为 wfc」）—— 编排全在 wfcSave.ts，这里只接线。 */
  const saveWfc = useCallback(async () => {
    if (savingWfc) {
      reportChatSave(false, 'error', '正在保存中，等这次存完再点')
      return
    }
    /* ★ 2026-10-25 用户指令「先确定位置和名称 保存时直接放进去」：
       第一步=系统另存为对话框（选目录+文件名），取消则中止 */
    const apiD = (globalThis as { api?: { wfcPickSavePath?: (d: string) => Promise<string | null> } }).api
    const defName = (activePageName || '未命名画布').replace(/[\\/:*?"<>|]/g, '_') + '.wfc'
    /* ★ 写模式 folder 节点作为保存默认目录（用户指令「作为存储目录」闭环） */
    let folderDefault = ''
    for (const t of Object.values(useBaseStore.getState().tables)) {
      const c = (t?.config ?? {}) as { folderPath?: string; folderMode?: string }
      if (c.folderPath && c.folderMode === 'write') { folderDefault = c.folderPath; break }
    }
    const defFullPath = folderDefault ? (folderDefault + '\\' + defName) : defName
    let savePath: string | null = null
    if (apiD?.wfcPickSavePath) {
      savePath = await apiD.wfcPickSavePath(defFullPath)
      if (!savePath) {
        audit('系统', '保存工程(.wfc) 取消', '用户取消另存为')
        reportChatSave(false, 'cancel', '已取消保存（保存位置对话框点了取消）')
        return
      }
      setSaveStep({ label: '保存位置：' + savePath + (folderDefault ? '（来自文件夹节点）' : '') + ' —— 扫描文件清单…', pct: 0 })
    } else {
      setSaveStep({ label: '保存位置：Downloads/' + defName + ' —— 扫描文件清单…', pct: 0 })
    }
    savingWfcRef.current = true
    setSavingWfc(true)
    saveStartRef.current = Date.now()
    setSaveElapsed(0)
    const saveLoc = 'Downloads/' + (activePageName || '未命名画布').replace(/[\\/:*?"<>|]/g, '_') + '.wfc'
    setSaveStep({ label: '保存位置：' + saveLoc + ' —— 扫描文件清单…', pct: 0 })
    audit('系统', '保存工程(.wfc) 开始', '画布=' + activePageName)
    try {
      /* ★ 2026-10-24 保存两段流第一步：扫描文件清单（用户确认后才存储） */
      const mf = await collectAssetManifest()
      setSaveManifest(mf)
      setSaveAwaitName(savePath || activePageName)
      setSavingWfc(false)
      setSaveStep(null)
      reportChatSave(true, 'open', '已调起保存面板：核对文件清单后点「开始保存」完成落盘（与工具栏保存同一条链路）。')
      return
    } catch (eScan: unknown) {
      setSavingWfc(false)
      setSaveStep(null)
      const msg = eScan instanceof Error ? eScan.message : String(eScan)
      reportChatSave(false, 'error', '扫描文件清单失败：' + msg)
      window.alert('扫描文件清单失败：' + msg)
      return
    }
    /* 清单确认后由 doSaveWfc 执行实际存储（两段流） */
  }, [savingWfc, activePageName])

  /** ★ 实际存储执行（清单确认后）——原 saveWfc 存储段迁移至此 */
  /* ★ 孤立节点/历史节点派生（保存面板 tab 用——edgeLinks 来自 Canvas 同步） */
  const orphanNodes = useMemo(() => {
    const st = useBaseStore.getState()
    return Object.values(st.tables).filter((t) => t && !t.deleted && !st.edgeLinks.includes(t.id)).map((t) => ({ id: t.id, name: t.name ?? t.id }))
  }, [saveManifest])
  const histNodes = useMemo(() => {
    const st = useBaseStore.getState()
    const out: { id: string; name: string; count: number; mb: number; entries: { key: string; hid: string; name: string; at: number; mb: number }[] }[] = []
    for (const t of Object.values(st.tables)) {
      const h = (t?.config as { pdfHistory?: { id: string; name?: string; at?: number }[] } | undefined)?.pdfHistory
      if (!h?.length) continue
      const entries = h.map((e, ei) => ({
        key: t.id + ':' + e.id,
        hid: e.id,
        name: e.name ?? `产物-${ei + 1}`,
        at: e.at ?? 0,
        mb: 15.6,
      }))
      out.push({ id: t.id, name: t.name ?? t.id, count: entries.length, mb: Math.round(entries.length * 15.6 * 10) / 10, entries })
    }
    return out
  }, [saveManifest])

  /* ★ 保存弹窗类型方块（kind→色类/字母，三tab统一） */
  const kindFmt = (kind: string) => {
    const map: Record<string, { c: string; ch: string }> = {
      pdf: { c: 'pm-fmt-pdf', ch: 'P' },
      pdfmerge: { c: 'pm-fmt-pdf', ch: 'M' },
      pdfreplace: { c: 'pm-fmt-pdf', ch: 'R' },
      ooword: { c: 'pm-fmt-word', ch: 'W' },
      oosheet: { c: 'pm-fmt-sheet', ch: 'X' },
      photo: { c: 'pm-fmt-photo', ch: '照' },
    }
    return map[kind] ?? { c: '', ch: (kind[0] || '?').toUpperCase() }
  }

  const histSelCount = histNodes.reduce((a, n) => a + n.entries.filter((e) => histSel[e.key] !== false).length, 0)
  const histSelMB = Math.round(histNodes.reduce((a, n) => a + n.entries.filter((e) => histSel[e.key] !== false).length, 0) * 15.6 * 10) / 10
  const skipCount = Object.values(orphanSkip).filter(Boolean).length

  /* saveManifest 打开时：历史默认全选 + 全收起 */
  useEffect(() => {
    if (!saveManifest) return
    const sel: Record<string, boolean> = {}
    for (const n of histNodes) for (const e of n.entries) sel[e.key] = true
    setHistSel(sel)
    setHistOpen({})
    setOrphanSkip({})
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [saveManifest])

  const doSaveWfc = useCallback(async (pageName: string, includeHistory: boolean | Set<string>, savePath?: string, excludeNodeIds?: string[]) => {
    if (savingWfc) return
    setSavingWfc(true)
    if (!saveStartRef.current) { saveStartRef.current = Date.now(); setSaveElapsed(0) }
    setSaveStep({ label: '准备存储…', pct: 1 })
    audit('系统', '保存工程(.wfc) 开始', '画布=' + pageName)
    try {
      const api = (globalThis as { api?: { wfcWriteFile?: (p: string, b: Uint8Array) => Promise<{ ok: boolean }> } }).api
      const writeTo = savePath && api?.wfcWriteFile ? savePath : undefined
      const r = await saveCanvasAsWfc(pageName, (st) => setSaveStep(st), { includeHistory, writeTo, excludeNodeIds })
      const kb = (r.bytes / 1024).toFixed(1)
      audit('系统', '保存工程(.wfc) 成功', (savePath ?? r.fileName) + ' ' + kb + 'KB 未入包资产=' + r.problems.length + ' 资产去重=' + (r.dedupInfo?.count ?? 0))
      /* ★ 完成态：留在弹窗显示折叠完成条（清单可展开核对），不再 alert 打断 */
      setSaveDone({ fileName: savePath ?? r.fileName, kb, problems: r.problems, dedup: r.dedupInfo?.count ?? 0 })
      setSaveDoneOpen(false)
      reportChatSave(true, 'done', `已保存 ${savePath ?? r.fileName}（${kb} KB）`)
    } catch (eU: unknown) {
      audit('错误', '保存工程(.wfc) 失败', eU instanceof Error ? eU.message : String(eU))
      const msg = eU instanceof Error ? eU.message : String(eU)
      reportChatSave(false, 'error', '保存失败：' + msg)
      window.alert(`保存失败：${msg}`)
    } finally {
      setSavingWfc(false)
      savingWfcRef.current = false
      setSaveElapsed((Date.now() - saveStartRef.current) / 1000)
      setSaveStep(null)
      saveElapsedTimerClear?.()
    }
  }, [savingWfc])

  /** ★ 另存为（2026-10-22 用户指令「保存按钮换成另存为」）：与保存同链路，
   *  但允许输入新文件名（默认当前画布名）——不覆盖原文件。 */
  const saveWfcAs = useCallback(async () => {
    if (savingWfc) return
    const name = window.prompt('另存为（文件名，不含扩展名）：', activePageName || '画布')
    if (!name) return
    setSavingWfc(true)
    setSaveStep({ label: '另存位置：Downloads/' + name + '.wfc —— 确认后开始存储', pct: 0 })
    audit('系统', '另存工程(.wfc) 开始', '画布=' + name)
    try {
      const r = await saveCanvasAsWfc(name, (st) => setSaveStep(st))
      const kb = (r.bytes / 1024).toFixed(1)
      audit('系统', '另存工程(.wfc) 成功', r.fileName + ' ' + kb + 'KB')
      window.alert(`已另存 ${r.fileName}（${kb} KB）`)
    } catch (e) {
      window.alert(`另存失败：${e instanceof Error ? e.message : String(e)}`)
    } finally {
      setSavingWfc(false)
      setSaveStep(null)
    }
  }, [savingWfc, activePageName])

  /* ★ 2026-11-07 对话「保存画布文件」：走工具栏保存的同一条链路（打开面板 → 清单确认 → 落盘）。
     不在这里另写一套存盘逻辑——保存的真相只有 wfcSave.ts 一条。 */
  useEffect(() => {
    const h = () => { void saveWfc() }
    window.addEventListener('chat-save-canvas', h)
    return () => window.removeEventListener('chat-save-canvas', h)
  }, [saveWfc])

  /** ★ 清空画布（2026-10-22 用户指令「新增清空按钮」）：所有节点进回收站
   *  （软删除——可从回收站整批找回，不是不可逆销毁）。 */
  const clearCanvas = useCallback(() => {
    const st = useBaseStore.getState()
    const alive = Object.values(st.tables).filter((t) => !t.deleted)
    if (!alive.length) { window.alert('画布已是空的'); return }
    if (!window.confirm(`清空画布：${alive.length} 个节点全部移入回收站？\n（可从回收站恢复）`)) return
    st.trashTables(alive.map((t) => t.id))
    audit('系统', '清空画布', alive.length + ' 个节点入回收站')
  }, [])

  /** 打开 .wfc（2026-10-07 用户：「双击*.wfc文件 打开软件并加载该文件内容」）。
   *  入口三个：① 双击文件（main argv → pendingWfc → getLaunchWfc）
   *           ② 运行中再双击（second-instance → onOpenWfc）
   *           ③ 标题栏 📂 按钮（pickWfc 系统对话框）
   *  三条路最终都走 openCanvasFromWfcFile —— 作为新画布页导入，不覆盖已开的画布。 */
  const [openingWfc, setOpeningWfc] = useState(false)
  const openWfc = useCallback(async (path: string) => {
    if (openingWfc) return
    setOpeningWfc(true)
    try {
      const api = (window as unknown as { api?: { readWfc?: (p: string) => Promise<Uint8Array> } }).api
      const bytes = await api?.readWfc?.(path)
      if (!bytes) throw new Error('读取文件失败')
      const r = await openCanvasFromWfcBytes(bytes, path.replace(/.*[/\\]/, '').replace(/\.wfc$/i, ''))
      const kb = (bytes.length / 1024).toFixed(1)
      window.alert(
        r.problems.length
          ? `已打开「${r.name}」（${r.nodes} 节点，${kb} KB）\n\n⚠ 部分资产缺失：\n${r.problems.join('\n')}`
          : `已打开「${r.name}」（${r.nodes} 节点，${kb} KB）`,
      )
    } catch (e) {
      window.alert(`打开 .wfc 失败：${e instanceof Error ? e.message : String(e)}`)
    } finally {
      setOpeningWfc(false)
    }
  }, [openingWfc])

  // ① 启动参数带 .wfc（双击文件首启）+ ② 运行中再双击（second-instance 转发）
  //   ⚠ 只在 C 端（window.api 存在）接线；浏览器无此概念。
  useEffect(() => {
    const api = (window as unknown as {
      api?: {
        getLaunchWfc?: () => Promise<string | null>
        onOpenWfc?: (cb: (path: string) => void) => () => void
      }
    }).api
    if (!api?.getLaunchWfc || !api.onOpenWfc) return
    void api.getLaunchWfc().then((p) => { if (p) void openWfc(p) })
    return api.onOpenWfc((p) => { void openWfc(p) })
    // openWfc 依赖 openingWfc ⇒ 用 ref 防抖会导致闭包复杂化；
    // 这里只在 mount 时挂一次，openWfc 内部自带防重复
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const startResize = useCallback((e: ReactMouseEvent) => {    e.preventDefault()
    const onMove = (ev: MouseEvent) => {
      const w = Math.min(760, Math.max(240, window.innerWidth - ev.clientX))
      setPropsWidth(w)
    }
    const onUp = () => {
      document.removeEventListener('mousemove', onMove)
      document.removeEventListener('mouseup', onUp)
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
    }
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
    document.addEventListener('mousemove', onMove)
    document.addEventListener('mouseup', onUp)
  }, [])

  useEffect(() => {
    init()
    /* ★ 重启恢复提示（2026-10-22 用户指令「关闭前画布不空，重启默认提示
       是否恢复，还是打开空白」）：init 完成后看当前画布是否非空——
       非空即上次未清空的现场，弹选择框（恢复=直接用；空白=清空入回收站）。
       ★ 只在会话首启弹一次（sessionStorage 标记），切画布页不再打扰。 */
    if (!sessionStorage.getItem('recoveryAsked')) {
      sessionStorage.setItem('recoveryAsked', '1')
      setTimeout(() => {
        try {
          const st = useBaseStore.getState()
          if (!st.ready) return
          const alive = Object.values(st.tables).filter((t) => !t.deleted)
          if (!alive.length) return
          /* ★ 2026-10-22 用户指令「样式进行重新设计。符合主题」：
             系统 confirm 弹窗（白壳系统标题栏）与应用主题完全脱节——
             换应用内主题化恢复弹窗（RecoveryDialog，--c-* 变量），逻辑不变：
             恢复=直接用现场；空白=清空入回收站可找回 */
          setRecoveryAsk(alive.length)
        } catch { /* 恢复提示失败不影响启动 */ }
      }, 600)
    }
  }, [init])

  // 选中态规范化（合并原「默认选中第一张表」+「选中项失效则清空」两个 effect）：
  // 原实现两个 effect 互相拉扯——A 选中 Object.keys(tables)[0]，若该节点已 deleted，
  // B 立刻把它清空，A 下一帧又选中它，形成每帧死循环 → 右侧属性栏疯狂闪烁/抖动。
  // 现改为：① 只在首次数据就绪时自动选中第一个「未删除」节点；② 选中项被删除后置空且不再自动填充。
  /* ★ 重启恢复弹窗（2026-10-22 主题化重设计）：>0 = 待询问（值为上次现场节点数）。
     原 window.confirm 换成应用内 RecoveryDialog（--c-* 主题变量） */
  const [recoveryAsk, setRecoveryAsk] = useState(0)
  const didAutoSelect = useRef(false)
  useEffect(() => {
    if (!ready) return
    const cur = selected ? tables[selected] : undefined
    if (selected && (!cur || cur.deleted)) {
      setSelected('')
      return
    }
    if (!selected && !didAutoSelect.current) {
      const first = Object.values(tables).find((t) => !t.deleted)
      if (first) {
        didAutoSelect.current = true
        setSelected(first.id)
      }
    }
  }, [ready, selected, tables])

  const schema = selected ? tables[selected] : undefined

  // 编辑器对应的节点被删进回收站 / 彻底删除时，自动关闭编辑器
  // （否则会留下一个指向不存在节点的浮层）
  useEffect(() => {
    if (!editorNode) return
    const t = tables[editorNode]
    if (!t || t.deleted) setEditorNode('')
  }, [editorNode, tables])

  // 任意节点的"展示结果"：table/param 取公式+lookup 算好的表；
  // sort 取上游汇入后排序；output 取上游汇入结果。
  const computed = useMemo(
    () => (schema ? computeNode(schema.id, tables, records) : { fields: [], rows: [] }),
    [schema, tables, records],
  )
  // 排序/输出节点的字段来自上游，用计算出的字段覆盖给数据表格
  const viewSchema = schema ? { ...schema, fields: computed.fields } : undefined

  // 连线语义：
  //  - 拖到 sort/output 节点 → 上游行汇入其 config.inputs（数据流）
  //  - 其它 → 在源头表建 lookup 字段引用目标表（关联）
  const handleConnect = useCallback(
    (srcId: string, tgtId: string) => {
      push()
      const src = tables[srcId]
      const tgt = tables[tgtId]
      if (!src || !tgt) return
      // 开始节点是工作流的触发起点：拖到任意节点都视为「触发/数据流连线」，
      // 让 开始 能连到普通表/参数表/排序/输出，并参与执行顺序与自动布局（之前对表/参数会静默失败）。
      if (src.kind === 'start') {
        if (srcId === tgtId) return
        const inputs = tgt.config?.inputs ?? []
        const linked = tgt.config?.linkedInputs ?? []
        setNodeConfig(tgtId, {
          inputs: inputs.includes(srcId) ? inputs : [srcId, ...inputs],
          linkedInputs: linked.includes(srcId) ? linked : [srcId, ...linked],
        })
        return
      }
      /* AI / 参数库节点（2026-10-06，用户：「模仿规则跟随节点样式 增加一个AI节点。
       * AI节点的连线放在普通节点的 下1锚点位置」「……参数库节点……上2锚点位置」）
       *
       * ★ 必须排在 isMergeKind / 规则分支**之前**（同规则分支当年排到最前的理由）：
       *   参数库线连到 sort/output 时若先命中 isMergeKind，会被多写一个 inputs
       *   （那是数据流口径）；AI 节点 fields=[] 走到 lookup 分支会静默 return。
       *   UI 准入已由 canConnectPorts 的 triggerLineSemantics 矩阵把过关，
       *   这里按同一口径兜底 + 落库。
       *
       * 落库 = 数据流同构（linkedInputs），不另存边表（与 projectFormat.edgesOf 同源）：
       *   · AI 线（普通节点下1 → AI）：AI 是数据消费者，上游记在 **AI 的 linkedInputs**
       *   · 参数库线（参数库 → 普通节点上2）：参数库是提供者，记在 **普通节点的 linkedInputs**
       *   渲染层（Canvas 边派生）按两端 kind 识别线种并显式指锚点 —— 与规则线同招。
       */
      if (tgt.kind === 'ai') {
        if (srcId === tgtId) return
        // 圆形节点之间互连不认（canConnectPorts 已拦，程序化调用走这里兜底）。
        // src.kind 已被上方 start 分支收窄，无需再判 start。
        if (isRoundNodeKind(src.kind)) return
        const linked = tgt.config?.linkedInputs ?? []
        if (linked.includes(srcId)) return // 重复连线不写
        // 偏好匹配（2026-10-06 用户：「针对不同类型的节点。ai接入后匹配不同的偏好」）：
        // AI 未配置过技能时按上游类型预填（text → 文字进文字出，注册表 AI_KIND_PREFS）；
        // 已配置（aiSkillId 非空，含 custom）一律不覆盖。合并进同一个 patch =
        // 一条历史快照（连线 + 偏好预填算一次动作，undo 一次全退）。
        const prefs = aiPrefsPatchOf(tgt.config ?? {}, [src.kind])
        setNodeConfig(tgtId, { linkedInputs: [srcId, ...linked], ...(prefs ?? {}) })
        return
      }
      if (src.kind === 'paramlib') {
        if (srcId === tgtId) return
        // 参数库线的落点必须是普通节点（圆形/开始节点不收）
        if (isRoundNodeKind(tgt.kind) || tgt.kind === 'start') return
        const linked = tgt.config?.linkedInputs ?? []
        if (linked.includes(srcId)) return
        setNodeConfig(tgtId, { linkedInputs: [srcId, ...linked] })
        return
      }
      // 汇入型节点：排序 / 输出 / 对照 / Word / Excel / 章节工作簿 —— 连线 = 上游数据流
      // 判定走 nodeStyle 单一真相源，与 core/engine.ts 同源（两处曾各写一份，易漏）
      if (isMergeKind(tgt.kind)) {
        const inputs = tgt.config?.inputs ?? []
        const linked = tgt.config?.linkedInputs ?? []
        // 连线 = 已连接来源（优先、有绿色线）；prepend 保证排在前面
        setNodeConfig(tgtId, {
          inputs: inputs.includes(srcId) ? inputs : [srcId, ...inputs],
          linkedInputs: linked.includes(srcId) ? linked : [srcId, ...linked],
        })
        return
      }
      /* 规则节点 ⟷ 被它锁定的辅节点：连线 = 规则锁定关系（2026-10-05）
       *
       * ★★ 必须排在**所有**其它分支之前（2026-10-05 修正）。
       *   原顺序把 isMergeKind 分支放在规则分支之前，导致：
       *     · 规则线连到「数据排序」时，命中 isMergeKind → 被写成**数据流**，
       *       规则锁定根本没建立，且症状是「线连上了但主从没生效」；
       *     · 连到「数据输出」同理。
       *   规则线与数据流线在 UI 上就长不同（className/颜色），行为却一样 ——
       *   这正是用户报的「主节点与其他节点连线样式不一致」的根因：
       *   样式走了 rule 线，语义却走了 flow 分支。
       *   判据放在最前面：只要任一端是规则节点，且两端都不是 start，就归规则锁定管。
       *
       *判定本身在 core/ruleSync.ruleConnectPatch（纯函数、对称处理两个方向），
       *   不内联在此：初版内联在这里时，测试抄了一份逻辑，改坏真实现测试仍全绿。*/
      const srcIsRule = src.kind === 'rule'
      const tgtIsRule = tgt.kind === 'rule'
      if (srcIsRule || tgtIsRule) {
        // 同类型约束（用户 2026-10-05 第 6 条：「规则跟随节点只能连接同种类型的节点」）
        // 判定走 core/ruleSync.canJoinRuleGroup（纯函数），与 UI 层
        // isValidConnection 共用同一份 —— 两处各写一份必然会漂移。
        // 这里必须再兜一道：isValidConnection 只在 UI 交互里生效，
        // 程序化调用 handleConnect 会绕过它。
        if (!canJoinRuleGroup(srcId, tgtId, tables)) return
        const rulePatch = ruleConnectPatch(srcId, tgtId, tables)
        if (rulePatch) {
          // 一次写完：连线 + 身份（primary/follow）。
          // ★ 必须同一个 setNodeConfig —— 分两次写会在中途失败时留下
          //   「连上了但没身份」的节点，而 group 里可能一个 primary 都没有。
          // ★ ruleRef 可能是 undefined（该节点已被另一个规则组占用身份，
          //   见 ruleConnectPatch 的注释）—— 这时**不能**把 ruleRef 写进 patch，
          //   否则 setNodeConfig 的浅合并会把已有身份覆盖成 undefined。
          //   故用条件展开而不是写 ruleRef: rulePatch.ruleRef。
          setNodeConfig(rulePatch.nodeId, {
            linkedInputs: rulePatch.linkedInputs,
            ...(rulePatch.ruleRef ? { ruleRef: rulePatch.ruleRef } : {}),
          })
          return
        }
        // ruleConnectPatch 返回 null（自己连自己/规则节点间互连/重复连线）：
        // 不写任何东西，也不 fallthrough 到数据流分支 —— 否则规则线会被
        // 记成数据流，规则节点变成「输出的上游」，正是本机制要避免的。
        return
      }
      // 传统 lookup 关联（需两表都有字段）
      if (!src.fields.length || !tgt.fields.length) return
      addField(srcId, defineField(`${tgt.name}引用`, 'lookup', {
        lookup: {
          table: tgtId,
          sourceField: src.fields[0].id,
          matchField: tgt.fields[0].id,
          returnField: tgt.fields[0].id,
        },
      }))
    },
    [tables, addField, setNodeConfig],
  )

  const handleEdgeRemove = useCallback(
    (edgeId: string) => {
      if (edgeId.startsWith('flow:')) {
        const [, rest] = edgeId.split('flow:')
        const [srcId, tgtId] = rest.split('->')
        // 删除连线需确认（用户备忘）：误触 × 直接拆数据流代价太大
        if (!confirm(`删除数据流连线「${tables[srcId]?.name ?? srcId} → ${tables[tgtId]?.name ?? tgtId}」？`)) return
        push()
        const t = tables[tgtId]
        if (t) {
          setNodeConfig(tgtId, {
            inputs: (t.config?.inputs ?? []).filter((id) => id !== srcId),
            linkedInputs: (t.config?.linkedInputs ?? []).filter((id) => id !== srcId),
          })
        }
        return
      }
      const [srcAndField] = edgeId.split('->')
      const [srcId, fieldId] = srcAndField.split('.')
      if (srcId && fieldId) {
        if (!confirm(`删除关联引用「${tables[srcId]?.name ?? srcId}.${fieldId}」？`)) return
        push()
        removeField(srcId, fieldId)
      }
    },
    [tables, removeField, setNodeConfig],
  )

  // 从面板拖入建节点（返回新 id，Canvas 据此定位）；kind 决定节点功能
  const handleAddTableAt = useCallback(
    (pos: XYPosition, kind: NodeKind = 'table'): string => {
      const id = `${kind}_${Date.now()}`
      const presets: Record<NodeKind, TableSchema> = {
        table: defineNode('table', '新表', [defineField('name', 'text')], {}, id),
        param: defineNode('param', '数据参数表', [defineField('name', 'text')], {}, id),
        // ★ 数据表节点：OO Excel 引用调试专用——默认两个字段（数字 123 / 456，用户 0:47 指定），
        //   拖出来就能在引用面板点选、走「引用更新策略 / 原生标记」全链验证。
        test_table: defineNode(
          'test_table',
          '数据表节点',
          [defineField('参数1', 'number'), defineField('参数2', 'number')],
          {},
          id,
        ),
        sort: defineNode('sort', '数据排序', [], {}, id),
        output: defineNode('output', '数据输出', [], { exportFormat: 'csv' }, id),
        start: defineNode('start', '开始', [], { trigger: 'manual' }, id),
        // 数据对照表：自身是「键 → 值…」的对照规则表，接上上游后按匹配键给每行补列
        map: defineNode(
          'map',
          '数据对照表',
          [defineField('键', 'text'), defineField('值', 'text')],
          { mapMatchField: '键', mapReturnFields: ['值'], mapDefault: '' },
          id,
        ),
        // Word 文档节点：模板里的内容控件（w:sdt）= 域，上游数据灌进域后出交付稿
        docx: defineNode('docx', 'Word 文档', [], { docxMap: {} }, id),
        // Excel 工作簿节点：命名区域 / 表格 = 域，上游数据灌进域后出工作簿
        xlsx: defineNode('xlsx', 'Excel 工作簿', [], { xlsxMap: {} }, id),
        // ★ OnlyOffice 节点：离线内嵌 OnlyOffice 内核（WPS 级编辑），字节直存
        ooword: defineNode('ooword', 'OO 文档', [], {}, id),
        oosheet: defineNode('oosheet', 'OO 工作簿', [], {}, id),
        // 章节工作簿：Excel 全量功能 + 章节层级/目录大纲（A 列 level / B 列编号 / C 列标题）
        chapter: defineNode('chapter', '章节工作簿', [], { xlsxMap: {} }, id),
        // 文本节点：一个最简单的值（数字或一串文字），可被其它节点引用
        text: defineNode('text', '文本', [defineField('text', 'text')], {}, id),
        // 规则节点（圆形）：不干活，只给连上的同类节点统一下发规则。
        // 无字段、无记录 —— 它的价值全在 config.rules（按 kind 分桶）。
        rule: defineNode('rule', '规则跟随', [], { rules: {} }, id),
        // AI 节点（圆形，2026-10-06）：吃数据 —— 普通节点下1 出线连进来，
        // 上游记在它自己的 linkedInputs。无字段，连线语义见 handleConnect AI 分支。
        ai: defineNode('ai', 'AI', [], {}, id),
        // 参数库节点（圆形，2026-10-06）：供数 —— 出线连到普通节点上2，
        // 下游普通节点把它记进 linkedInputs。无字段。
        paramlib: defineNode('paramlib', '参数库', [], {}, id),
        // 节点组（2026-10-21）：视觉容器——框住多个节点（config.members 成员清单，
        // groupW/groupH 边框尺寸）。无字段、不进计算引擎。
        group: defineNode('group', '节点组', [], { groupW: 360, groupH: 240, members: [] }, id),
        // 便利贴/标注（2026-10-22）：亮黄便签，双击改文字
        note: defineNode('note', '便利贴', [defineField('text', 'text')], {}, id),
        // PDF 生成（2026-11-07 由「PDF合并」更名）：上游文档/图片按序生成一份 PDF
        pdfmerge: defineNode('pdfmerge', 'PDF生成', [], {}, id),
        photo: defineNode('photo', '照片', [], {}, id),
        webmon: defineNode('webmon', '网址监控', [], { webmonTargets: [], webmonIntervalSec: 0 }, id),
        xlsx2dxf: defineNode('xlsx2dxf', 'Excel转DXF', [], { xlsx2dxf: { drawPoint: true, drawLabel: true, drawLine: false, textHeight: 3.5, labelOffset: 1.5 } }, id),
        pdf: defineNode('pdf', 'PDF', [], {}, id),
        folder: defineNode('folder', '文件夹', [], {}, id),
        pdfreplace: defineNode('pdfreplace', 'PDF替换页', [], {}, id),
        pdfslim: defineNode('pdfslim', 'PDF瘦身', [], {}, id),
        pdf2word: defineNode('pdf2word', 'PDF转word', [], {}, id),
        fileout: defineNode('fileout', '文件输出', [], {}, id),
      }
    const base = presets[kind].name
    const name = uniqueName(base, tables)
    push()
    addTable({ ...presets[kind], name })
    // 文本节点：给默认内容（引用与显示都有东西）
    if (kind === 'text') {
      setTableData(id, presets.text.fields, [{ text: '文本' }])
    }
    // 数据表节点（测试）：种一行数据，引用面板点选立刻有值可读
    if (kind === 'test_table') {
      setTableData(id, presets.test_table.fields, [{ 参数1: 123, 参数2: 456 }])
    }
    // 便利贴：默认一句话
    if (kind === 'note') {
      setTableData(id, presets.note.fields, [{ text: '备注…' }])
    }
    // 立即落位：拖入坐标写入 store，保证撤销可还原、切换页后位置不丢失
    setNodePosition(id, pos)
    setSelected(id)
    setRightTab('schema')
    return id
  },
  [addTable, setNodePosition, tables, setTableData],
)

  /* ★ 节点组（2026-10-21）：把选中节点框进一个组容器（包围盒 + 内边距） */
  const handleCreateGroup = useCallback(
    (ids: string[], rect: { x: number; y: number; w: number; h: number }) => {
      push()
      const pad = 24
      const head = 34
      const gid = handleAddTableAt(
        { x: Math.round(rect.x - pad), y: Math.round(rect.y - head - pad / 2) },
        'group',
      )
      setNodeConfig(gid, {
        groupW: Math.round(rect.w + pad * 2),
        groupH: Math.round(rect.h + head + pad),
        members: ids,
      })
      setSelected(gid)
    },
    [handleAddTableAt, setNodeConfig, setSelected, push],
  )


  // 数据参数表：从文件导入，整体覆盖字段+记录
  const handleImport = useCallback(
    async (file: File) => {
      if (!selected) return
      try {
        const { fields, rows } = await parseFile(file)
        push()
        setTableData(selected, fields, rows)
      } catch (err) {
        window.alert('导入失败：' + (err as Error).message)
      }
    },
    [selected, setTableData],
  )

  // 数据输出节点：把当前计算结果的字段+行导出为选定格式
  const handleExport = useCallback(async () => {
    if (!schema) return
    const { fields, rows } = computeNode(schema.id, tables, records)
    const fmt = schema.config?.exportFormat ?? 'csv'
    try {
      await exportRows(schema.name, fields, rows, fmt)
    } catch (err) {
      window.alert('导出失败：' + (err as Error).message)
    }
  }, [schema, tables, records])

  const handleDeleteTable = useCallback(() => {
    if (!selected) return
    push()
    trashTable(selected)
    setSelected((s) => (s === selected ? '' : s))
  }, [selected, trashTable])

  const handleDuplicateTable = useCallback(() => {
    if (!selected) return
    push()
    const newId = duplicateTable(selected)
    setSelected(newId)
  }, [selected, duplicateTable])

  // 单个节点移入回收站（已选中态的清理交给调用方统一 push）
  const trashOne = (id: string) => {
    trashTable(id)
    setSelected((s) => (s === id ? '' : s))
  }

  // 删除单节点：移入回收站（软删除，保留数据可恢复）
  const handleDeleteNode = useCallback(
    (id: string) => {
      push()
      trashOne(id)
    },
    [trashTable],
  )

  const handleDeleteSelectedNodes = useCallback(
    (ids: string[]) => {
      if (!ids.length) return
      push()
      // 一次批量删除作为同一个「删除批次」，便于在回收站里「逐步恢复上一次删除」整批找回
      trashTables(ids)
    },
    [trashTables],
  )

  // 复制节点（带 +24,+24 偏移，避免完全重叠）
  const handleDuplicateNode = useCallback(
    (id: string) => {
      push()
      const newId = duplicateTable(id)
      const p = positions[id]
      if (p) setNodePosition(newId, { x: p.x + 24, y: p.y + 24 })
      setSelected(newId)
    },
    [duplicateTable, positions, setNodePosition],
  )

  // 拆离数据流连线：从目标节点的 inputs/linkedInputs 中移除该来源（删除连线需确认——用户备忘）
  const handleDetachFlow = useCallback(
    (src: string, tgt: string) => {
      const st = useBaseStore.getState()
      const sName = st.tables[src]?.name ?? src
      const tName = st.tables[tgt]?.name ?? tgt
      if (!confirm(`拆离数据流连线「${sName} → ${tName}」？`)) return
      push()
      const t = st.tables[tgt]
      if (!t) return
      setNodeConfig(tgt, {
        inputs: (t.config?.inputs ?? []).filter((x) => x !== src),
        linkedInputs: (t.config?.linkedInputs ?? []).filter((x) => x !== src),
      })
    },
    [setNodeConfig],
  )

  // 拆离 lookup 连线：删除源表上的 lookup 字段（删除连线需确认——用户备忘）
  const handleDetachLookup = useCallback(
    (srcTable: string, fieldId: string) => {
      const st = useBaseStore.getState()
      const f = st.tables[srcTable]?.fields.find((x) => x.id === fieldId)
      if (!confirm(`删除关联字段「${st.tables[srcTable]?.name ?? srcTable}.${f?.name ?? fieldId}」？`)) return
      push()
      removeField(srcTable, fieldId)
    },
    [removeField],
  )

  // 粘贴节点：为选中节点创建副本（新 id + 偏移 + 重命名），并重映射内部数据流连线
  // ★ 2026-10-22 用户指令「节点组粘贴携带成员；三种模式」：
  //   mode='group'（默认）：组+成员（数据齐全）
  //   mode='group-empty'：组+成员（成员数据清空——只要结构不要内容）
  //   mode='flat'：不建组，只铺成员节点
  const handlePasteNodes = useCallback(
    (items: CopyItem[], mode?: 'group' | 'group-empty' | 'flat'): string[] => {
      const cur = useBaseStore.getState()
      const idMap: Record<string, string> = {}
      const newIds: string[] = []
      /* 组条目展开成员。★ 2026-10-22 防重：copySelection 现在复制时就装好成员
         （组框在前、成员紧随），剪贴板可能已含成员条目——展开时按 seen 跳过，
         只补缺失的（兼容旧剪贴板内容）。顺序保持剪贴板原序，与返回的 newIds
         一一对应（Canvas.pasteSelection 按 items[i] 配对计算粘贴落点）。 */
      const expanded: CopyItem[] = []
      const seen = new Set<string>()
      for (const it of items) {
        if (it.table.kind === 'group') {
          const members = ((it.table.config as { members?: string[] }).members ?? []) as string[]
          const alive = members.filter((m) => cur.tables[m] && !cur.tables[m].deleted)
          if (mode === 'flat') {
            for (const m of alive) {
              if (seen.has(m)) continue
              seen.add(m)
              expanded.push({ id: m, table: cur.tables[m], record: cur.records[m] ?? [], pos: cur.positions[m] })
            }
            continue
          }
          expanded.push(it)
          seen.add(it.id)
          for (const m of alive) {
            if (seen.has(m)) continue
            seen.add(m)
            expanded.push({ id: m, table: cur.tables[m], record: mode === 'group-empty' ? [] : (cur.records[m] ?? []), pos: cur.positions[m] })
          }
        } else if (!seen.has(it.id)) {
          seen.add(it.id)
          expanded.push(it)
        }
      }
      items = expanded
      items.forEach((it, i) => {
        idMap[it.id] = `${it.table.kind}_${Date.now()}_${i}`
      })
      items.forEach((it, i) => {
        const newId = idMap[it.id]
        const remap = (x: string) => idMap[x] ?? x
        const rawLinked = (it.table.config?.linkedInputs ?? []).map(remap)
        const rawInputs = (it.table.config?.inputs ?? []).map(remap)
        const rawMembers = ((it.table.config as { members?: string[] }).members ?? []) as string[]
        const t: TableSchema = {
          ...it.table,
          id: newId,
          name: uniqueName(it.table.name, cur.tables),
          config: {
            ...it.table.config,
            linkedInputs: rawLinked,
            inputs: rawInputs,
            /* 组成员 → 新成员 id（旧成员被复制的副本接管） */
            ...(it.table.kind === 'group' ? { members: rawMembers.map(remap) } : {}),
          },
        }
        push()
        addTable(t)
        ;(it.record ?? []).forEach((r) => addRow(newId, r))
        const base = it.pos ?? { x: 0, y: 0 }
        /* ★ 2026-10-22 修「粘贴后成员相对位置乱」：整组统一平移 dx=dy=40——
           原 +40+i*30 随索引递增把成员逐个推远（相对布局打散，截图实锤）。
           与 Canvas.pasteSelection 的统一平移口径一致。 */
        setNodePosition(newId, { x: base.x + 40, y: base.y + 40 })
        newIds.push(newId)
      })
      return newIds
    },
    [addTable, addRow, setNodePosition],
  )

  // 自动排列：一次性把分层后的坐标写入 store（并压入一次撤销栈）
  const handleApplyLayout = useCallback(
    (pos: Record<string, NodePos>) => {
      push()
      Object.entries(pos).forEach(([id, p]) => setNodePosition(id, p))
      /* ★ 2026-10-22 用户指令「坐标等也在诊断日志里」 */
      pushLog(`自动排列 ${Object.keys(pos).length} 个节点`)
    },
    [setNodePosition],
  )

  // 属性面板 / 表单的写操作也需进历史栈（撤销可回退）
  const pushRename = (id: string, name: string) => { push(); renameTable(id, name) }
  const pushAddField = (tid: string, f: FieldDef) => { push(); addField(tid, f) }
  const pushUpdateField = (tid: string, fid: string, p: Partial<FieldDef>) => { push(); updateField(tid, fid, p) }
  const pushRemoveField = (tid: string, fid: string) => { push(); removeField(tid, fid) }
  const pushSetNodeConfig = (tid: string, p: Partial<NodeConfig>) => { push(); setNodeConfig(tid, p) }
  /* 规则锁定：与 setNodeConfig 同惯例——store 动作不自压历史，统一由这里 push()。
     若 store 里也自压，一次操作会压两个相同快照，表现为「撤销一次没反应」。 */
  const pushSetRule = (rid: string, kind: string, p: Partial<NodeRuleSet>) => { push(); setRule(rid, kind, p) }
  const pushClearRule = (rid: string, kind: string) => { push(); clearRule(rid, kind) }
  const pushUnlinkRule = (rid: string, tid: string) => { push(); unlinkFromRuleNode(rid, tid) }
  /* 主次 / 跟随切换：同样由App 层 push() 统一压历史（store 动作不自压）。
     这两个是「成员管理面板」的唯一写入口 —— 面板本身不碰store。 */
  const pushSetRulePrimary = (rid: string, tid: string) => { push(); setRulePrimary(rid, tid) }
  const pushSetRuleFollow = (rid: string, tid: string, follow: boolean) => { push(); setRuleFollow(rid, tid, follow) }

  /* 双击圆形规则节点 → 打开成员管理面板。
     useCallback 固定引用：Canvas 侧虽用 ref 兜住了，但仍让引用稳定，
     避免无谓地让下游 memo 失效。顺带把该节点也选上，
     这样浮层关掉后右侧属性栏正好停在这个规则节点上（接着改规则）。 */
  const openRuleMembers = useCallback((ruleNodeId: string) => {
    setSelected(ruleNodeId)
    setRightTab('schema')
    setRuleMembersNode(ruleNodeId)
  }, [])
  /** OO 节点类型切换（按导入扩展名：docx→ooword / xlsx→oosheet；画布徽章/面板/样式全部跟随） */
  const switchTableKind = (tid: string, kind: NodeKind) => { push(); setTableKind(tid, kind) }
  // 对照规则表的编辑：改单元格不进撤销栈（否则每敲一个字符压一次快照，会很快挤掉真正的操作历史），
  // 只有增删行这类结构性改动才进栈。
  const pushAddRow = (tid: string) => { push(); addRow(tid, {}) }
  const pushRemoveRow = (tid: string, idx: number) => { push(); removeRow(tid, idx) }

  // 调试钩子：浏览器控制台可访问 window.__store 看实时状态（含 positions）
  ;(window as unknown as { __store?: typeof useBaseStore }).__store = useBaseStore

  if (!ready) {
    return <div style={{ padding: 24, color: 'var(--c-text3)', fontFamily: 'system-ui' }}>加载底座…</div>
  }

  return (
    <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
      {/* ── app 标题栏（C 端）：标识 + 当前画布名 ──
          撤销/重做/回收站仍在菜单带内：那三项是 Canvas 内部 state（历史栈、回收站开关），
          提到 App 层需要把状态上移，会把画布不变量泄漏到布局组件 —— 不做。

          这里不再挂写死的品牌名「数据关联底座」（用户：「数据关联底座这个名字去掉
          放在最上方标题栏」）：品牌名是产品级常量、每个画布都一样，而标题栏是
          「当前打开的是哪个画布」的唯一出口，同屏摆两个名字只会让人分不清哪个是
          当前的。产品名保留在宿主窗口/标签标题（document.title 的初始占位）里。*/}
      {/* ★ 2026-10-19 用户「这两行合并」：原 .app-titlebar 独立行取消，logo/画布名/版本
          与 保存/打开/重启/关闭 作为插槽并进 CanvasRibbon 的页签行（单行通栏）。 */}
      {(() => {
        /* ★ 动作区按钮定义（fixed=核心按钮永不折叠；其余窄窗时从左收进 ⋯ 菜单） */
        /* ★ 2026-10-24 用户指令「从回收站、对齐磁吸顺序进入折叠区」：这两项排折叠池最前 */
        const actionButtons: { id: string; icon: ReactNode; title: string; busy?: boolean; disabled?: boolean; fixed?: boolean; onClick: () => void }[] = [
                    { id: 'trash', icon: <IcClear size={17} />, title: '回收站（可从回收站恢复）', onClick: () => window.dispatchEvent(new CustomEvent('toggle-trash-panel')) },
          { id: 'magnet', icon: <IcMove size={17} />, title: '对齐磁吸：拖动节点时自动吸附对齐线（点击开/关）', busy: magnetUi, onClick: () => setMagnet(!magnetUi) },
          /* ★ 2026-10-26 套件中心入口在折叠池（低频钮）——原放 ribbon 常显区把日志钮挤出可视区（19:08 用户反馈） */
          { id: 'suites', icon: <IcPuzzle size={17} />, title: '套件中心：查看/安装/停用节点功能套件（PDF转word 引擎等）', onClick: () => suiteCenterUI.open() },
{ id: 'gpu', icon: <IcChip size={17} />, title: gpuOn === false ? '硬件加速：关（点击开启，重启生效）' : '硬件加速：开（点击关闭，重启生效）', busy: gpuOn === false, onClick: toggleGpu },
          { id: 'draglod', icon: <IcMove size={17} />, title: '拖动节点显示：' + (dragLodUi === 'outline' ? '只显示轮廓（点击切换为整体所见即所得）' : '整体所见即所得（点击切换为只显示轮廓）'), busy: dragLodUi === 'outline', onClick: () => {
            const nv = dragLodUi === 'outline' ? 'wysiwyg' : 'outline'
            try { prefSet('dragLod', nv) } catch { /* ignore */ }
            window.dispatchEvent(new CustomEvent('drag-lod-changed', { detail: nv }))
            setDragLodUi(nv)
          } },
          { id: 'remote', icon: <IcGlobe size={17} />, title: '远程访问设置：iPad/手机接入的地址与配对码', onClick: () => { const api = (window as unknown as { api?: { showRemoteDialog?: () => Promise<boolean> } }).api; api?.showRemoteDialog?.() } },
          ...(APP_CHANNEL === 'web' ? [{ id: 'full', icon: (isFull ? <IcMinimize size={17} /> : <IcMaximize size={17} />), title: isFull ? '退出全屏' : '全屏（沉浸式画布）', busy: isFull, onClick: toggleFull }] : []),
          { id: 'view', icon: <IcGrid size={17} />, title: '视图：画布显示设置（连线流动动画等）', fixed: true, onClick: () => {
            setEdgeAnimUi((() => { try { return (localStorage.getItem('edgeAnimMode') as 'pause' | 'always' | 'off') || 'pause' } catch { return 'pause' } })())
            setViewMenuOpen((v) => !v)
          } },
          { id: 'save', icon: <IcDatabase size={17} />, title: '保存 .wfc：当前画布全部内容存为工程文件（覆盖同名文件）', busy: savingWfc, disabled: savingWfc, fixed: true, onClick: saveWfc },
          { id: 'saveas', icon: <IcSaveAs size={17} />, title: '另存为 .wfc：输入新文件名另存一份（不覆盖原文件）', busy: savingWfc, disabled: savingWfc, fixed: true, onClick: saveWfcAs },
          { id: 'clear', icon: <IcClear size={17} />, title: '清空画布：全部节点移入回收站（可从回收站恢复）', onClick: clearCanvas },
          { id: 'open', icon: <IcFolderOpen size={19} />, title: '打开 .wfc 工程文件：作为新画布页导入', disabled: openingWfc, fixed: true, onClick: () => {
            const api = (window as unknown as { api?: { pickWfc?: () => Promise<string | null> } }).api
            if (api?.pickWfc) {
              void api.pickWfc().then((p) => { if (p) void openWfc(p) })
              return
            }
            const inp = document.createElement('input')
            inp.type = 'file'
            inp.accept = '.wfc'
            inp.onchange = async () => {
              const f = inp.files?.[0]
              if (!f) return
              try {
                const r = await openCanvasFromWfcFile(f)
                window.alert(`已打开「${r.name}」（${r.nodes} 节点）`)
              } catch (e) {
                window.alert(`打开 .wfc 失败：${e instanceof Error ? e.message : String(e)}`)
              }
            }
            inp.click()
          } },
          ...(typeof electronRelaunch === 'function' ? [{ id: 'relaunch', icon: <IcRotateCcw size={19} />, title: '重启应用：整个进程退出后重新拉起（内嵌 OO 服务随新进程重建）', fixed: true, onClick: () => void electronRelaunch() }] : []),
          ...(typeof electronQuit === 'function' ? [{ id: 'quit', icon: <IcX size={17} />, title: '彻底关闭：进程与内嵌服务全部退出（区别于点 X 仅隐藏窗口）', fixed: true, onClick: () => void electronQuit() }] : []),
          { id: 'exportassets', icon: <IcFolderOpen size={17} />, title: '导出资产到文件夹：把画布全部文件字节导出（查重用——同内容同指纹同大小）', onClick: async () => {
            try {
              const { listAllBlobs } = await import('./core/pdfBlobStore')
              const items = await listAllBlobs()
              const api = (window as unknown as { api?: { assetsExport?: (items: { name: string; bytesB64: string }[]) => Promise<{ ok: boolean; dir?: string; count?: number; reason?: string }> } }).api
              if (!api?.assetsExport) { window.alert('桌面端暂不支持（请重启应用获取最新桥）'); return }
              const r = await api.assetsExport(items)
              if (r.ok) window.alert('已导出 ' + (r.count ?? 0) + ' 个文件到：\n' + (r.dir ?? ''))
            } catch (e) { window.alert('导出失败：' + String(e)) }
          } },
        ]
        const ribbonTitleLeft = (
          <>
        {/* ★ 22:4x 用户：「软件的左上角的 logo 也和快捷方式一样」—— 天蓝圆角砖墨圈 logo 图 */}
        <img className={"app-mark" + (hideLevel >= 3 ? " brand-hide-logo" : "")} src="/logo-tile.png" alt="ClientBase" draggable={false} />
        <div ref={brandRef} style={{ display: 'flex', alignItems: 'center', gap: 7, flexWrap: 'nowrap', minWidth: 0, flex: 1, overflow: 'hidden' }}>
        <button className={"app-docname" + (hideLevel >= 2 ? " brand-hide-doc" : "")} title="当前画布页 · 点击重命名" onClick={renameActivePage}>
          {activePageName}
        </button>
        {/* ★ 版本号（2026-10-17）：桌面端与网页端同源（根目录 VERSION），
            网页发布后一眼能确认线上跑的是哪一版，不用问也不用猜。 */}
        <span
          className={"app-ver" + (hideLevel >= 1 ? " brand-hide-ver" : "")}
          title={`版本 ${APP_VERSION}\n渠道 ${APP_CHANNEL}\n构建 ${APP_BUILD || '未知'}${APP_CHANNEL === 'web' ? '\n网页版：数据存在服务器 server-data/pages，Office 在线编辑不可用' : ''}`}
        >
          {APP_VERSION_LABEL}
        </span>
        {/* 退出登录（2026-10-17 网页端口令）：只有网页端有门禁，桌面端不显示。
            走 /logout 清 cookie 再回登录页，服务端处理，不依赖前端状态。 */}
        {APP_CHANNEL === 'web' && (
          <a className="app-logout" href="/logout" title="退出登录（清除本设备的访问凭证）">退出</a>
        )}
        {/* 标题栏动作区（2026-10-07 用户：「做成方形按钮放在右侧」）：
            四钮方形 32×32 聚拢靠右；完整语义在悬浮 title。 */}
        </div>
        </>
        )


        const ribbonTitleRight = (
        <div className="app-titlebar-actions" style={{ position: 'relative', flexShrink: 0 }} ref={foldRef}>
          {/* ★ 2026-10-24 用户指令「右侧按钮优先显示 左边的按钮折叠」：
              可折叠按钮池按 foldVisible 截断显示，折叠项收进 ⋯ 菜单；
              fixed 核心按钮（视图/保存/另存/打开/重启/退出）永不折叠。 */}
          {(() => {
            const renderBtn = (b: typeof actionButtons[number]) => (
              <button
                key={b.id}
                className={'app-savebtn' + (b.busy ? ' is-busy' : '')}
                disabled={b.disabled}
                title={b.title}
                onClick={b.onClick}
              >
                {b.icon}
              </button>
            )
            const pool = actionButtons.filter((b) => !b.fixed)
            const visible = pool.slice(-foldVisible || undefined)
            const folded = foldVisible > 0 ? pool.slice(0, -foldVisible) : []
            const fixeds = actionButtons.filter((b) => b.fixed)
            /* ★ 2026-10-24 用户指令「⋯ 放在视图左边 固定不动」：⋯+核心按钮包成
                shrink:0 不可分割组（右锚定）——visible 收缩时 ⋯ 位置恒定不漂移 */
            return (
              <>
                {visible.map(renderBtn)}
                <div style={{ display: 'flex', alignItems: 'center', gap: 'var(--btn-gap)', flexShrink: 0 }}>
                  {folded.length > 0 && (
                    <button
                      className={'app-savebtn' + (foldMenu ? ' is-busy' : '')}
                      title="更多操作（点击展开/收起）"
                      onClick={() => setFoldMenu((v) => !v)}
                    >
                      <span className="app-savebtn-more">⋯</span>
                    </button>
                  )}
                  {fixeds.map(renderBtn)}
                </div>
                {foldMenu && folded.length > 0 && (
                  <div ref={foldMenuRef} style={{ position: 'absolute', top: '100%', left: 0, right: 0, display: 'flex', flexWrap: 'wrap', gap: 6, padding: '6px 8px', background: 'var(--ed-bg2, var(--c-bg1))', border: '1px solid var(--ed-border, var(--c-border))', zIndex: 50 }}>
                    {folded.map(renderBtn)}
                  </div>
                )}
              </>
            )
          })()}
          {/* ★ 2026-10-25 用户指令「保存wfc菜单 参照PDF替换页确认弹窗重新设计」：
              右侧抽屉 → 居中大窗（pr-confirm 同款语言）：表格化清单/勾选联动按钮样式/底栏重排 */}
          {saveManifest && (
            <div className="pr-mask">
              <div className="pr-confirm pr-confirm--save" onClick={(e) => e.stopPropagation()}>
                <div className="pr-confirm-head">
                  <span className="pr-save-title">保存工程<span className="pr-save-page">{activePageName || '未命名画布'}</span></span>
                  <span className="pr-save-total">{saveManifest.totalN} 文件 · {saveManifest.totalMB} MB · {saveManifest.groups.length} 节点</span>
                  <button className="pn-del" title="关闭" onClick={() => { setSaveManifest(null); setSaveDone(null); setSavingWfc(false) }}>×</button>
                </div>
                {saveDone && (
                  <div className="pr-save-done">
                    <span className="pr-save-done-ic">✓</span>
                    <span className="pr-save-done-txt">已保存 {saveDone.fileName}（{saveDone.kb} KB{saveDone.dedup ? ` · 去重${saveDone.dedup}` : ''}）</span>
                    <button className="pr-btn pr-btn--sm" onClick={() => setSaveDoneOpen((v) => !v)}>
                      {saveDoneOpen ? '收起清单 ▴' : '展开清单 ▾'}
                    </button>
                  </div>
                )}
                {saveDone && saveDone.problems.length > 0 && saveDoneOpen && (
                  <div className="pr-cf-skip" style={{ margin: '0 16px', whiteSpace: 'normal' }} title={saveDone.problems.join('\n')}>
                    ⚠ 未入包 {saveDone.problems.length} 项：{saveDone.problems[0]}
                  </div>
                )}
                {(!saveDone || saveDoneOpen) && (
                <>
                <div className="pr-cf-tabs">
                  {([['files', '文件清单'], ['orphan', '孤立节点'], ['history', '历史文件']] as const).map(([k, label]) => (
                    <button key={k} className={'pr-cf-tab' + (saveTab === k ? ' pr-cf-tab--on' : '')} onClick={() => setSaveTab(k)}>
                      {label}{k === 'orphan' && orphanNodes.length ? ` ${orphanNodes.length}` : ''}{k === 'history' && histNodes.length ? ` ${histNodes.length}` : ''}
                    </button>
                  ))}
                </div>
                <div className="pr-cf-table">
                  {saveTab === 'files' && (
                    <div className="pr-save-list">
                      {saveManifest.groups.map((g) => {
                        const open = filesOpen[g.nodeId] !== false
                        return (
                        <div key={g.nodeId} className="pr-save-card">
                          <div className="pr-save-card-head" style={{ cursor: 'pointer' }} onClick={() => setFilesOpen((v) => ({ ...v, [g.nodeId]: open ? false : true }))}>
                            <span className="pm-fmt pm-fmt-pdf">P</span>
                            <span className="pr-save-name">{g.node}</span>
                            <span className="pr-save-meta">{g.entries.length} 项 · {g.totalMB} MB</span>
                            <span style={{ fontSize: 10, color: 'var(--c-text4)' }}>{open ? '▼' : '▶'}</span>
                          </div>
                          {open && g.entries.map((e, i) => (
                            <div key={i} className="pr-save-file">
                              <span className="pr-save-dots" />
                              <span className="pn-loc pn-loc--in" title="软件内部：字节随工程文件存储">内</span>
                              <span className="pr-save-name">{e.dup ? '↪ ' : ''}{e.name}</span>
                              <span className="pr-save-meta">{e.dup ? '去重' : e.sizeMB + ' MB'}</span>
                            </div>
                          ))}
                        </div>
                        )
                      })}
                      {saveManifest.groups.length === 0 && <div style={{ fontSize: 12, color: 'var(--ed-text3, var(--c-text3))', padding: '12px 4px' }}>画布内没有文件资产</div>}
                    </div>
                  )}
                  {saveTab === 'orphan' && (
                    <div className="pr-save-list">
                      <div style={{ fontSize: 11.5, color: 'var(--ed-text3, var(--c-text3))', padding: '2px 4px 8px' }}>没有连线的节点——去掉勾选 = 本次保存不包含它（画布不受影响）</div>
                      {orphanNodes.length === 0 && <div style={{ fontSize: 12, color: 'var(--ed-text3, var(--c-text3))', padding: '4px' }}>没有孤立节点 ✅</div>}
                      {orphanNodes.map((n) => {
                        const t = useBaseStore.getState().tables[n.id]
                        const kf = kindFmt(t?.kind ?? '')
                        const skipped = !!orphanSkip[n.id]
                        const open = !!orphanOpen[n.id]
                        /* ★ 2026-10-25 用户令「孤立节点也能展开」：按 kind 提取内部文件词条 */
                        const cfg = (t?.config ?? {}) as { pdfItems?: { name: string }[]; replaceTargets?: { name: string }[]; replaceSources?: { name: string }[]; photoItems?: { name: string }[] }
                        const inner: string[] = t?.kind === 'pdf'
                          ? (cfg.pdfItems ?? []).map((x) => x.name)
                          : t?.kind === 'pdfreplace'
                            ? [...(cfg.replaceTargets ?? []).map((x) => '被替换：' + x.name), ...(cfg.replaceSources ?? []).map((x) => '替换页：' + x.name)]
                            : t?.kind === 'photo'
                              ? (cfg.photoItems ?? []).map((x) => x.name)
                              : []
                        return (
                          <div key={n.id} className={'pr-save-card' + (skipped ? ' is-dim' : '')}>
                            <div className="pr-save-card-head pr-save-check">
                              <label style={{ display: 'flex', alignItems: 'center', gap: 8, flex: 1, minWidth: 0, cursor: 'pointer' }}>
                                <input
                                  type="checkbox"
                                  checked={!skipped}
                                  onChange={(ev) => setOrphanSkip((v) => ({ ...v, [n.id]: !ev.target.checked }))}
                                />
                                <span className={'pm-fmt ' + kf.c}>{kf.ch}</span>
                                <span className="pr-save-name">{n.name}</span>
                              </label>
                              <span className="pr-save-meta">{skipped ? '本次不保存' : '随本次保存'}</span>
                              <span
                                style={{ fontSize: 10, color: 'var(--c-text4)', cursor: 'pointer', padding: '2px 4px' }}
                                title="展开内部文件"
                                onClick={(e) => { e.preventDefault(); setOrphanOpen((v) => ({ ...v, [n.id]: !open })) }}
                              >{open ? '▼' : '▶'}</span>
                            </div>
                            {open && (
                              inner.length ? inner.map((nm, ii) => (
                                <div key={ii} className="pr-save-file">
                                  <span className="pr-save-dots" />
                                  <span className="pr-save-name">{nm}</span>
                                </div>
                              )) : <div className="pr-save-file"><span className="pr-save-dots" /><span className="pr-save-name" style={{ color: 'var(--c-text4)' }}>无内部文件词条</span></div>
                            )}
                          </div>
                        )
                      })}
                    </div>
                  )}
                  {saveTab === 'history' && (
                    <div className="pr-save-list">
                      <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '2px 4px 8px', fontSize: 11.5, color: 'var(--ed-text3, var(--c-text3))' }}>
                        <span>展开逐条勾选，勾选的随 .wfc 打包</span>
                        <button className="pr-btn pr-btn--sm" onClick={() => setHistSel((v) => { const nv = { ...v }; for (const n of histNodes) for (const e of n.entries) nv[e.key] = true; return nv })}>全选</button>
                        <button className="pr-btn pr-btn--sm" onClick={() => setHistSel((v) => { const nv = { ...v }; for (const n of histNodes) for (const e of n.entries) nv[e.key] = false; return nv })}>全不选</button>
                        <button className="pr-btn pr-btn--sm" title="每个节点只勾选最新一条" onClick={() => setHistSel((v) => {
                          const nv = { ...v }
                          for (const n of histNodes) for (const e of n.entries) nv[e.key] = false
                          for (const n of histNodes) { const latest = n.entries[0]; if (latest) nv[latest.key] = true }
                          return nv
                        })}>仅最新版本</button>
                      </div>
                      {histNodes.length === 0 && <div style={{ fontSize: 12, color: 'var(--ed-text3, var(--c-text3))', padding: '4px' }}>没有历史记录节点</div>}
                      {histNodes.map((n) => {
                        const open = !!histOpen[n.id]
                        const selCount = n.entries.filter((e) => histSel[e.key] !== false).length
                        return (
                          <div key={n.id} className="pr-save-card">
                            <div className="pr-save-card-head" style={{ cursor: 'pointer' }} onClick={() => setHistOpen((v) => ({ ...v, [n.id]: !open }))}>
                              <span className="pm-fmt pm-fmt-pdf">P</span>
                              <span className="pr-save-name">{n.name}</span>
                              <span className="pr-save-meta">{selCount}/{n.count} 条 · {(selCount * 15.6).toFixed(1)} MB</span>
                              <span style={{ fontSize: 10, color: 'var(--c-text4)' }}>{open ? '▼' : '▶'}</span>
                            </div>
                            {open && n.entries.map((e, ei) => (
                              <label key={e.key} className={'pr-save-file pr-save-check'}>
                                <input type="checkbox" checked={histSel[e.key] !== false}
                                  onChange={(ev) => setHistSel((v) => ({ ...v, [e.key]: ev.target.checked }))} />
                                <span className="pr-save-dots" />
                                <span className="pm-fmt pm-fmt-pdf" title="PDF 产物">P</span>
                                <span className="pn-loc pn-loc--in" title="软件内部：字节随工程文件存储">内</span>
                                <span className="pr-save-name">{e.name}</span>
                                <span className="pr-save-meta">第 {n.entries.length - ei} 新 · {e.mb} MB</span>
                              </label>
                            ))}
                          </div>
                        )
                      })}
                    </div>
                  )}
                </div>
                </>
                )}
                <div className="pr-confirm-foot" style={{ flexDirection: 'column', alignItems: 'stretch', gap: 8 }}>
                  {savingWfc && saveStep ? (
                    <>
                      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: 'var(--ed-text2, var(--c-text2))' }}>
                        <span>{saveStep.label}</span>
                        <span>{saveStep.pct}% · {saveElapsed.toFixed(1)}s</span>
                      </div>
                      <div style={{ height: 6, borderRadius: 3, background: 'var(--c-bg3, #8882)', overflow: 'hidden' }}>
                        <div style={{ width: saveStep.pct + '%', height: '100%', borderRadius: 3, background: 'var(--c-blue, #3b82f6)', transition: 'width .25s ease' }} />
                      </div>
                    </>
                  ) : (
                    <>
                      {/* ★ 勾选联动按钮（细粒度 histSel 统计——含历史一眼可辨） */}
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <span className="pr-save-sum">
                          {saveManifest.totalN} 文件{histSelCount ? ` · 历史 ${histSelCount} 条` : ''}{skipCount ? ` · 剔除 ${skipCount} 节点` : ''}
                        </span>
                        <button className="pr-btn" onClick={() => { setSaveManifest(null); setSaveDone(null); setSavingWfc(false) }}>{saveDone ? '关闭' : '取消'}</button>
                        {!saveDone && <span style={{ flex: 1 }} />}
                        {!saveDone && <button
                          className={'pr-btn' + (histSelCount ? ' pr-btn--go' : ' pr-btn--hist-off')}
                          style={{ fontWeight: 600 }}
                          onClick={() => {
                            const n = saveAwaitName || activePageName
                            const sel = new Set(Object.entries(histSel).filter(([, v]) => v).map(([k]) => k))
                            const skip = Object.entries(orphanSkip).filter(([, v]) => v).map(([k]) => k)
                            setSaveManifest(null)
                            setSaveDone(null)
                            void doSaveWfc(n, sel, saveAwaitName || undefined, skip)
                          }}
                        >
                          {histSelCount
                            ? `开始保存 · 含历史（${histSelCount} 条 / ${histSelMB.toFixed(1)} MB）`
                            : '开始保存 · 不含历史'}
                        </button>}
                      </div>
                    </>
                  )}
                </div>
              </div>
            </div>
          )}
          {savingWfc && saveStep && (
            <div style={{ position: 'fixed', bottom: 24, left: '50%', transform: 'translateX(-50%)', zIndex: 200, padding: '10px 16px', background: 'var(--ed-bg2, var(--c-bg1))', border: '1px solid var(--ed-border, var(--c-border))', borderRadius: 8, boxShadow: '0 6px 20px rgba(0,0,0,0.45)', minWidth: 300 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: 'var(--ed-text2, var(--c-text2))', marginBottom: 6 }}>
                <span>正在保存 {activePageName || '画布'}</span>
                <span>{saveStep.pct}% · {saveElapsed.toFixed(1)}s</span>
              </div>
              <div style={{ height: 6, borderRadius: 3, background: 'var(--c-bg3, #8882)', overflow: 'hidden' }}>
                <div style={{ width: saveStep.pct + '%', height: '100%', borderRadius: 3, background: 'var(--c-blue, #3b82f6)', transition: 'width .25s ease' }} />
              </div>
              <div style={{ fontSize: 11.5, color: 'var(--ed-text3, var(--c-text3))', marginTop: 5 }}>{saveStep.label}</div>
            </div>
          )}
          {viewMenuOpen && (
            <div className="view-menu">
              <div className="view-menu-head">视图设置</div>
              {(['pause', 'always', 'off'] as const).map((m) => (
                <div key={m} className={'ctx-item' + (edgeAnimUi === m ? ' is-on' : '')} onClick={() => { setEdgeAnim(m); setViewMenuOpen(false) }}>
                  连线动画：{m === 'pause' ? '拖动时暂停' : m === 'always' ? '始终播放' : '关闭'}{edgeAnimUi === m ? ' ✓' : ''}
                </div>
              ))}
              <div className="ctx-item" onClick={() => setMagnet(!magnetUi)} title="拖动节点时的对齐追踪线与磁吸吸附">
                磁吸对齐：{magnetUi ? '开 ✓' : '关'}
              </div>
              <div className="ctx-item" onClick={() => {
                const nv = dragLodUi === 'outline' ? 'wysiwyg' : 'outline'
                try { prefSet('dragLod', nv) } catch { /* ignore */ }
                window.dispatchEvent(new CustomEvent('drag-lod-changed', { detail: nv }))
                setDragLodUi(nv)
              }} title="拖动节点时的显示：整体所见即所得 / 只显示轮廓">
                拖动节点显示：{dragLodUi === 'outline' ? '只显示轮廓' : '整体所见即所得'} ↻
              </div>
              <div className="ctx-item" onClick={() => {
                const nv = !perfHudUi
                try { prefSet('perfHud', nv ? '1' : '0') } catch { /* ignore */ }
                window.dispatchEvent(new CustomEvent('perf-hud-changed', { detail: nv ? '1' : '0' }))
                setPerfHudUi(nv)
              }} title="画布性能监控 HUD（FPS/帧耗时/长任务/JS 堆）">
                性能监控：{perfHudUi ? '开 ✓' : '关'}
              </div>
            </div>
          )}
        </div>
        )
        ribbonSlots.current = { left: ribbonTitleLeft, right: ribbonTitleRight }
        return null
      })()}
      <div className="layout">
        {/* 左侧竖排图标条：折叠/展开画布（位置与 OO 左侧批注栏一致）。
            CSS 压width 只能让容器塌到 0，内部 flex 子项仍会撑出 ~24px（实测 palette w=0
            而子块 w=24，画布左边漏一条空白带），所以折叠时整块不渲染最干净。 */}
        {/* ★ 2026-10-19 用户「展开的时候不要有个侧边空条」：
            展开态不再渲染 side-strip 空条——折叠按钮挪到面板右上角（内嵌）；
            折叠态保留窄条 + 展开按钮。 */}
        {paletteCollapsed && (
          <div className="side-strip">
            <button
              className="side-strip-btn is-folded"
              onClick={() => setPaletteCollapsed(false)}
              title="展开节点库"
            >
              <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true">
                <path d="M9 6l6 6-6 6" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </button>
            {/* ★ 2026-10-23 用户指令「参数库侧边 和展开节点库并列」 */}
            <button
              className={'side-strip-btn is-folded' + (paramlibOpen ? ' side-strip-btn--on' : '')}
              onClick={() => setParamlibOpen((v) => !v)}
              title="展开参数库"
            >
              <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true">
                <path d="M4 17l6-5-6-5M12 19h8" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </button>
          </div>
        )}
        {paletteCollapsed ? (
          <aside className="palette palette--folded" aria-hidden="true" />
        ) : (
        <aside className="palette" style={{ position: 'relative' }}>
          <button
            className="side-strip-btn palette-fold-inline"
            onClick={() => setPaletteCollapsed(true)}
            title="折叠节点库（画布占满）"
          >
            <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true">
              <path d="M15 6l-6 6 6 6" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </button>
          {/* ★ 2026-10-23 参数库入口（palette 展开态——与折叠按钮并排） */}
          <button
            className={'side-strip-btn palette-fold-inline' + (paramlibOpen ? ' side-strip-btn--on' : '')}
            style={{ right: 44 }}
            onClick={() => setParamlibOpen((v) => !v)}
            title="展开参数库"
          >
            <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true">
              <path d="M4 17l6-5-6-5M12 19h8" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </button>
          <div className="sidebar-section">
            <p className="sidebar-label">添加节点</p>
            {/* 节点库按「七类角色」折叠分组 —— 数据来自 core/nodeStyle.ts 的规则表。
                新增一种 kind 只需在规则表加一行，这里自动出现对应分组与条目。
                形状 = 角色（SHAPE_CLASS），颜色 = 族（--nk-*），两处都从规则表派生。
                分组容器 .palette-grid 为单列铺满，与上一版侧栏观感一致。 */}
            {PALETTE_GROUPS.map((g) => (
              <div key={g.role} className="palette-group">
                <p className="sidebar-label">
                  {g.label}
                  <span className="palette-count">{g.kinds.length}</span>
                </p>
                <div className="palette-grid">
                  {g.kinds.map((k) => (
                    <div
                      key={k.kind}
                      className="palette-item"
                      draggable
                      onDragStart={(e) => {
                        e.dataTransfer.setData('application/reactflow', k.kind)
                        e.dataTransfer.effectAllowed = 'move'
                      }}
                      title={k.hint}
                    >
                      <div className={`palette-shape ${paletteItemClassOf(k.kind)}`} />
                      <span>{styleOf(k.kind).label}</span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
          {/* ★ 2026-10-19 用户「这个删掉」：侧栏「操作」区（复制当前表/删除当前表）移除——
              同名功能已在节点右键菜单/快捷键里，侧栏不再重复放入口 */}
        </aside>
        )}
        {/* ★ 2026-10-23 参数库侧板（与节点库并列；点击条目建 param 节点并灌内置数据） */}
        {paramlibOpen && (
          <ParamlibPanel
            onClose={() => setParamlibOpen(false)}
            onAdd={(name, fields, rows) => {
              const id = handleAddTableAt({ x: 260, y: 140 }, 'param')
              setTableData(id, fields, rows)
              useBaseStore.getState().renameTable(id, name)
              pushLog(`参数库：已添加「${name}」（${rows.length} 行 × ${fields.length} 列）`, 'ok')
            }}
          />
        )}

        <main className="canvas-wrap">
          <Canvas
            ribbonSlots={ribbonSlots}
            tables={tables}
            selected={selected}
            positions={positions}
            onSelect={(id) => {
              setSelected(id)
              setRightTab('schema')
            }}
            onConnect={handleConnect}
            onEdgeRemove={handleEdgeRemove}
            onAddTableAt={handleAddTableAt}
            onNodeDragStop={(id, pos) => {
              setNodePosition(id, pos)
            }}
            onNodeDragStopMulti={(updates) => {
              setNodePositions(updates)
            }}
            onDeleteNode={handleDeleteNode}
            onDuplicateNode={handleDuplicateNode}
            onDeleteSelectedNodes={handleDeleteSelectedNodes}
            onDetachFlow={handleDetachFlow}
            onDetachLookup={handleDetachLookup}
            onPasteNodes={handlePasteNodes}
            onApplyLayout={handleApplyLayout}
            onRestoreTable={restoreTable}
            onRestoreLastTrash={restoreLastTrash}
            onPurgeTable={purgeTable}
            onEmptyTrash={emptyTrash}
            onOpenNode={(id) => {
              setSelected(id)
              setEditorNode(id)
            }}
            onOpenPopupNode={(id) => {
              setSelected(id)
              setPopupNode(id)
            }}            onCreateGroup={handleCreateGroup}
            onOpenRuleMembers={openRuleMembers}
          />
          {/* ★ 画布底部日志面板 + 右上角诊断钮（2026-10-22 用户指令）：
              absolute 挂在 .canvas-wrap（relative）内——横跨画布宽度（不含右栏），
              不参与 React Flow 变换层；与 OO 节点内诊断抽屉同源同形态。 */}
          <LogDock />
        </main>

        {/* ★ 属性栏改右上角圆角悬浮框（2026-11-01 用户指令），不再占右栏挤画布 */}

        {selected && (
        <aside className="props-panel prop-float">
          <div className="props-tabs">
            <button
              className={`tab${rightTab === 'schema' ? ' tab-active' : ''}`}
              onClick={() => setRightTab('schema')}
            >
              属性
            </button>
            <button
              className={`tab${rightTab === 'data' ? ' tab-active' : ''}`}
              onClick={() => setRightTab('data')}
            >
              数据
            </button>
          </div>
          <div className="props-content">
            {rightTab === 'schema' ? (
              <PropertiesPanel
                table={schema}
                tables={tables}
                records={records}
                onRename={pushRename}
                onAddField={pushAddField}
                onUpdateField={pushUpdateField}
                onRemoveField={pushRemoveField}
                onSetNodeConfig={pushSetNodeConfig}
                onSetRule={pushSetRule}
                onClearRule={pushClearRule}
                onUnlinkRuleNode={pushUnlinkRule}
                onImport={handleImport}
                onExport={handleExport}
                rows={schema ? records[schema.id] ?? [] : []}
                onAddRow={pushAddRow}
                onSetCell={setCell}
                onRemoveRow={pushRemoveRow}
              />
            ) : viewSchema ? (
              <TableForm
                schema={viewSchema}
                rows={computed.rows}
                onCellEdit={(idx, fid, v) => {
                  if (schema && (schema.kind === 'table' || schema.kind === 'param')) {
                    push()
                    setCell(schema.id, idx, fid, v)
                  }
                }}
              />
            ) : (
              <div className="props-empty">无节点</div>
            )}
          </div>
        </aside>
        )}
      </div>

      {ruleMembersNode && tables[ruleMembersNode] && (
        <RuleMembersPanel
          ruleNodeId={ruleMembersNode}
          tables={tables}
          onClose={() => setRuleMembersNode('')}
          onSetPrimary={pushSetRulePrimary}
          onSetFollow={pushSetRuleFollow}
          onUnlink={pushUnlinkRule}
        />
      )}

      {editorNode && tables[editorNode] && (
        <NodeEditor
          anyDragging={anyDragging}
          nodeId={editorNode}
          tables={tables}
          records={records}
          onClose={() => setEditorNode('')}
          onSetNodeConfig={pushSetNodeConfig}
          onSwitchKind={switchTableKind}
          onRename={pushRename}
        />
      )}

      {popupNode && tables[popupNode] && hasPopupForm(tables[popupNode].kind) && (
          <PopupNodeEditor
            nodeId={popupNode}
            tables={tables}
            records={records}
            onClose={() => setPopupNode('')}
            onRename={pushRename}
            onAddField={pushAddField}
            onUpdateField={pushUpdateField}
            onRemoveField={pushRemoveField}
            onAddRow={pushAddRow}
            onSetCell={setCell}
            onRemoveRow={pushRemoveRow}
            onPushHistory={push}
          />
        )}

        {/* ★ 重启恢复弹窗（2026-10-22 主题化重设计，替代系统 confirm）：
            恢复=直接用现场（关弹窗即可，画布本就是上次内容）；
            空白=全部节点移入回收站可找回 */}
        {recoveryAsk > 0 && (
          <div className="rc-mask" onClick={(e) => e.stopPropagation()}>
            <div className="rc-dialog" role="dialog" aria-modal="true">
              <div className="rc-head">
                <span className="rc-badge">↺</span>
                <span>检测到上次画布有未清空的内容</span>
              </div>
              <div className="rc-body">
                <p className="rc-line">
                  上次退出时画布上有 <b>{recoveryAsk}</b> 个节点，可直接恢复现场。
                </p>
                <div className="rc-opts">
                  <button
                    className="rc-btn rc-btn-primary"
                    onClick={() => setRecoveryAsk(0)}
                  >
                    恢复上次画布
                  </button>
                  <button
                    className="rc-btn"
                    onClick={() => {
                      try {
                        const st = useBaseStore.getState()
                        const alive = Object.values(st.tables).filter((t) => !t.deleted)
                        if (alive.length) {
                          st.trashTables(alive.map((t) => t.id))
                          try { void useBaseStore.getState().saveNow() } catch { /* ignore */ }
                        }
                      } catch { /* ignore */ }
                      setRecoveryAsk(0)
                    }}
                  >
                    打开空白画布<span className="rc-sub">（内容移入回收站，可找回）</span>
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
    </div>
  )
}
