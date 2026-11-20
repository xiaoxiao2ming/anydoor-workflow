import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { extractOoLayoutProfile, type OoLayoutProfile } from './core/ooLayoutProfile'
import { createPortal } from 'react-dom'
import { prefGet, prefSet } from './ui/prefs'
import { editorOpenFrom } from './ui/editorOpenFrom'
import { IcFolderOpen, IcDatabase, IcDownload, IcHistory, IcPanelLeft, IcPanelRight, IcInfo, IcMark, IcGrid, IcRefresh, IcSun, IcMoon, IcMaximize, IcMinimize, IcX } from './ui/icons'
import { createWorkbook, addWorksheet } from '@office-kit/xlsx/workbook'
import { workbookToBytes } from '@office-kit/xlsx/io'
import type { NodeConfig, TableSchema } from './core/schema'
import { isWorkbookNodeKind } from './core/nodeStyle'
import type { Row } from './core/store'
import { computeNode } from './core/engine'
import { ooRefPatchOf } from './core/ooRef'
import { injectDomains, buildDomainEntries, refColorOf, domainRawKey, domainNameOf, extractRefMarks, diffDomainValues, stripFullCalcFlag } from './core/ooDomains'
import { useOoRefRefresh } from './office/useOoRefRefresh'
import { discoverXlsxFieldsFromBytes } from './xlsx/field'
import { discoverDocxSdt, toWordSites } from './core/docxSdtDiscover'
import { WordReportPanel } from './docx/WordReportPanel'
import { XlsxReportPanel } from './xlsx/XlsxReportPanel'
import { WorkbookEditor } from './xlsx/WorkbookEditor'
import { RefPanel, RightPanel } from './xlsx/RefPanel'
import { importDocxTemplate } from './docx/importTemplate'
import { OnlyOfficeNodeEditor, type OoEditorHandle, type OOKind } from './office/OnlyOfficeNodeEditor'
import { decideMainChannel, rowsHaveContent, rowStats } from './office/saveChannel'
import { pushLog } from './ui/logBus'
import { discardFrozen } from './office/ooPreload'
import { auditFeedback } from './core/audit'
import { useBaseStore } from './core/store'
import {
  getTemplateStore,
  pickTemplateFile,
  saveDocAs,
  seedTemplatesFromDisk,
  type TemplateMeta,
} from './docx/templateStore'

/**
 * 「态 B」放大态编辑器 —— 双击 docx / xlsx 节点后打开。
 *
 * 为什么一定要 Portal 到 document.body：
 *   React Flow 的画布是一个带 translate() + scale() 的 transform 子树。
 *   编辑器若留在节点 DOM 里，会同时踩五个坑：
 *     ① canvas 位图在 CSS scale 下必然模糊
 *     ② position:fixed 被 transform 祖先劫持，弹层跑到画布坐标系
 *     ③ 命中测试要除以 zoom，坐标算不准
 *     ④ 输入框里按退格 → 冒泡到画布 → 整个节点被删
 *     ⑤ 每个节点一份编辑器实例，内存翻倍
 *   Portal 出去后，这五个问题一次全消失 —— 所以这里不做「节点内联编辑器」。
 *
 * 节点小框（态 A）只负责「看一眼 + 改单值」，重活一律进这里。
 */

const KIND_TITLE: Record<string, string> = {
  docx: 'Word 文档',
  xlsx: 'Excel 工作簿',
  ooword: 'OO 文档',
  oosheet: 'OO 工作簿',
  ooslide: 'OO 演示',
  oopdf: 'OO PDF',
  chapter: '章节工作簿',
}

/* ── 外部数据引用更新策略（标题栏方形按钮的设置卡共用，2026-10-18）────────── */
type RefUpdMode = 'off' | 'ask' | 'auto'
const REF_UPD_LABEL: Record<RefUpdMode, string> = { off: '不检测', ask: '检测后提示', auto: '自动更新' }
const REF_UPD_DESC: Record<RefUpdMode, string> = {
  off: '打开时不检查外部关联，按保存内容显示（数字可能不是最新）。',
  ask: '打开时检查外部关联，发现数据有变化才弹窗问你。',
  auto: '打开时直接用画布最新值更新，不弹窗不打扰。',
}
/** 徽章短文案（标题栏左上角 logo 位） */
const BADGE_TEXT: Record<string, string> = {
  docx: 'Word',
  xlsx: 'Excel',
  ooword: 'Word',
  oosheet: 'Excel',
  ooslide: 'PPT',
  oopdf: 'PDF',
  chapter: '章节',
}

type Tab = 'main' | 'book' | 'data'

/**
 * 「走 WorkbookEditor 工作簿路径」的节点类型。
 * 判定走 core/nodeStyle 的单一真相源（isWorkbookNodeKind），
 * 勿在本文件另写一份 —— 本地副本曾与 Canvas/PropertiesPanel 漂移。
 * chapter（章节工作簿）= Excel 全量功能 + 外挂章节编排层，差别只在 outline 插槽。
 */
const isWorkbookKind = isWorkbookNodeKind

interface Props {
  nodeId: string
  tables: Record<string, TableSchema>
  records: Record<string, Row[]>
  onClose: () => void
  onSetNodeConfig: (tableId: string, patch: Partial<NodeConfig>) => void
  /** OO 节点类型切换（按导入文件扩展名自动切 ooword/oosheet；可缺省——Trash 预览等场景不传） */
  onSwitchKind?: (tableId: string, kind: 'ooword' | 'oosheet') => void
  onRename: (tableId: string, name: string) => void

  /** 拖动中：OO iframe 降级为占位（拖动合成卡顿根治） */
  anyDragging?: boolean
}

/** Uint8Array → base64（分块避免 apply 栈溢出）——OO 内容随画布节点配置持久化用 */
function u8ToB64(u8: Uint8Array): string {
  let bin = ''
  for (let i = 0; i < u8.length; i += 0x8000) bin += String.fromCharCode.apply(null, Array.from(u8.subarray(i, i + 0x8000)) as unknown as number[])
  return btoa(bin)
}
/** base64 → 字节（打开历史记录用，2026-10-19） */
function b64ToU8(b64: string): Uint8Array<ArrayBuffer> {
  const bin = atob(b64)
  const u8 = new Uint8Array(new ArrayBuffer(bin.length))
  for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
  return u8
}

export function NodeEditor({ nodeId, tables, records, onClose: rawOnClose, onSetNodeConfig, onSwitchKind, onRename, anyDragging }: Props) {
  /* ★ 2026-10-23 用户指令「关闭节点也做个特效」：关闭走收缩动画（打开的逆向，
   *   窗体收回到节点位置）。遮蔽 onClose prop——组件内全部关闭路径
   *   （× / 不保存 / 保存并关闭 / handleOoClose）自动先播动画再卸载，零改动点。
   *   全局特效开关 uiFx（工具带按钮，默认开）关闭时直通不播。 */
  const closingRef = useRef(false)
  const closeWithAnim = (doClose: () => void) => {
    if (closingRef.current) { doClose(); return }
    let fxOn = true
    try { fxOn = localStorage.getItem('uiFx') !== '0' } catch { /* ignore */ }
    if (!fxOn) { doClose(); return }
    try {
      const el = nodeId ? (document.querySelector(`.react-flow__node[data-id="${nodeId}"]`) as HTMLElement | null) : null
      const r = el?.getBoundingClientRect()
      const win = document.querySelector('.node-editor-mask .node-editor') as HTMLElement | null
      const mask = document.querySelector('.node-editor-mask') as HTMLElement | null
      if (r && r.width > 0 && win && mask) {
        const wr = win.getBoundingClientRect()
        const dx = r.x + r.width / 2 - (wr.x + wr.width / 2)
        const dy = r.y + r.height / 2 - (wr.y + wr.height / 2)
        const s = Math.max(0.05, Math.min(r.width / wr.width, r.height / wr.height))
        closingRef.current = true
        /* ★ 分镜（2026-10-23 用户指令「关闭 侧边栏先从两侧滑出，节点再丝滑退出」）：
           ① 左右侧栏挂 leaving（各自 ref-panel-anim-out 滑出）→ ② 240ms 后
           窗体收缩回节点 → ③ 270ms 后真正卸载。总关闭 ~510ms。 */
        setOoRefLeaving(true)
        setOoRightLeaving(true)
        const shrink = () => {
          win.style.setProperty('--ne-dx', `${Math.round(dx)}px`)
          win.style.setProperty('--ne-dy', `${Math.round(dy)}px`)
          win.style.setProperty('--ne-s', String(s))
          win.classList.add('ne-app-close')
          mask.classList.add('ne-mask-fadeout')
          setTimeout(() => { doClose() }, 270)
        }
        setTimeout(shrink, 240)
        return
      }
    } catch { /* 动画失败直通关闭，不影响功能 */ }
    doClose()
  }
  const onClose = () => closeWithAnim(rawOnClose)
  const table = tables[nodeId]
  const kind = table?.kind
  /** 默认页签：记住上次页签（备忘：节点自动保存上次的配置）；没存过 → Excel 已绑模板直接进「打开工作簿」，其余进「绑定与产出」 */
  const [tab, setTab] = useState<Tab>(() => {
    try {
      const saved = localStorage.getItem(`ne-tab:${nodeId}`) as Tab | null
      if (saved === 'main' || saved === 'data' || (saved === 'book' && tables[nodeId]?.config?.xlsxTemplateId)) {
        return saved
      }
    } catch {
      /* 隐私模式读不了就算了 */
    }
    return isWorkbookKind(kind) && tables[nodeId]?.config?.xlsxTemplateId ? 'book' : 'main'
  })
  /** 备忘 · 页签切换即持久化（下次打开恢复现场） */
  useEffect(() => {
    try {
      localStorage.setItem(`ne-tab:${nodeId}`, tab)
    } catch {
      /* 存不了就算了 */
    }
  }, [nodeId, tab])
  /* ★ 2026-10-23 用户指令「双击节点打开加动效，像打开 app 一样苹果那种丝滑展开」：
     挂载瞬间消费双击节点的屏幕 rect（editorOpenFrom，Canvas 双击时写入），
     计算窗体中心与节点中心的位移 + 缩放比，挂 .ne-app-open 播展开动画。
     rect 缺失（如快捷键打开）退普通出现，不影响功能。 */
  useEffect(() => {
    const from = editorOpenFrom.rect
    editorOpenFrom.rect = null
    /* 全局特效开关（2026-10-23 工具带 ✨ 按钮）：关闭时打开也直通 */
    let fxOn = true
    try { fxOn = localStorage.getItem('uiFx') !== '0' } catch { /* ignore */ }
    /* ★ 分镜（2026-10-23 用户指令「侧边栏后面从两侧滑动进入」）：
       窗体展开 ~300ms 后两侧栏才开始渲染（各自 ref-panel-anim-in 滑入）；
       fx 关或无 rect 时侧栏立即就绪（无分镜）。 */
    if (!fxOn) { setNeSidesReady(true); return }
    const sidesT = setTimeout(() => setNeSidesReady(true), from && from.w ? 300 : 0)
    if (!from || !from.w) return
    const mask = document.querySelector('.node-editor-mask') as HTMLElement | null
    const win = mask?.querySelector('.node-editor') as HTMLElement | null
    if (!win || !mask) return
    const r = win.getBoundingClientRect()
    if (!r.width || !r.height) return
    const dx = from.x + from.w / 2 - (r.x + r.width / 2)
    const dy = from.y + from.h / 2 - (r.y + r.height / 2)
    /* 起始缩放：节点/窗体的短边比（clamp 防极端值），透视上「从节点里长出来」 */
    const s = Math.max(0.05, Math.min(from.w / r.width, from.h / r.height))
    win.style.setProperty('--ne-dx', `${Math.round(dx)}px`)
    win.style.setProperty('--ne-dy', `${Math.round(dy)}px`)
    win.style.setProperty('--ne-s', String(s))
    win.classList.add('ne-app-open')
    mask.classList.add('ne-mask-fade')
    const t = setTimeout(() => {
      win.classList.remove('ne-app-open')
      mask.classList.remove('ne-mask-fade')
      win.style.removeProperty('--ne-dx')
      win.style.removeProperty('--ne-dy')
      win.style.removeProperty('--ne-s')
    }, 420)
    return () => { clearTimeout(t); clearTimeout(sidesT) }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])
  const [draftName, setDraftName] = useState(table?.name ?? '')
  /** 工作簿被就地编辑后自增，用来让绑定面板重新发现域（id 没变但内容变了） */
  const [bookVersion, setBookVersion] = useState(0)
  const [bookBytes, setBookBytes] = useState<Uint8Array | null>(null)
  const [bookErr, setBookErr] = useState('')
  const [busy, setBusy] = useState('')
  /** 模板库历史（xlsx 头部下拉 + 空态历史列表） */
  const [libs, setLibs] = useState<TemplateMeta[]>([])
  /** 最大化：标题栏方形按钮与面板 Ribbon 共用这一个状态 */
  const [headMax, setHeadMax] = useState(false)
  /* ── OO 节点（ooword/oosheet）标题栏文件操作 + 左右面板（与 Excel 同一套）── */
  const ooRef = useRef<OoEditorHandle>(null)
  /** OO 关闭确认：文档有未保存修改时弹三选（保存并关闭/不保存/取消） */
  const [ooCloseAsk, setOoCloseAsk] = useState(false)
  const [ooCloseBusy, setOoCloseBusy] = useState(false)
  const handleOoClose = async () => {
    try {
      const mod = ooRef.current ? await ooRef.current.isModified() : false
      if (mod) { setOoCloseAsk(true); return }
    } catch { /* 状态拿不到就直接关 */ }
    onClose()
  }
  const ooFileRef = useRef<HTMLInputElement>(null)
  /* ★ 2026-10-23 pdf 节点进 OO 链（OO 9.4 原生 PDF 编辑） */
  const isOO = kind === 'ooword' || kind === 'oosheet' || kind === 'pdf'
  const [ooSaving, setOoSavling] = useState(false)
  /** 左栏「节点数据引用」：开合状态按节点记忆（与 Excel 的 refpanel:<id> 同键风格） */
  const [ooRefOpen, setOoRefOpen] = useState(() => {
    try { return localStorage.getItem(`ooref:${nodeId}`) !== '0' } catch { return true }
  })
  useEffect(() => {
    try { localStorage.setItem(`ooref:${nodeId}`, ooRefOpen ? '1' : '0') } catch { /* ignore */ }
  }, [ooRefOpen, nodeId])
  const [ooRightOpen, setOoRightOpen] = useState(false)
  /** OO 诊断抽屉：标题栏 ⓘ 切换，默认关闭（用户 2026-09-30 要求） */
  const [ooDiagOpen, setOoDiagOpen] = useState(false)
  /** 引用格着色开关（2026-10-18 用户指令：标题栏诊断信息右侧加按钮）。
   *  ★ 改版：纯视觉开关 —— 颜色不写文件，切换即时生效不重开（见 toggleRefMark）。 */
  const [refMarkOn, setRefMarkOn] = useState<'1' | '0'>(() => {
    try { return (localStorage.getItem('refMarkColors') as '1' | '0') || '1' } catch { return '1' }
  })
  const refMarkOnRef = useRef<'1' | '0'>(refMarkOn)
  /** 引用格覆盖层数据（打开文档注入时提取：格地址 + 颜色 + 表格几何），供 refMarks prop */
  const [refMarksData, setRefMarksData] = useState<{ marks: { sheet: string; ref: string; color: string }[]; geom: unknown } | null>(null)

  /* ══ 外部关联引用 · 更新策略（2026-10-18 用户指令）══════════════════════════
   * 用户原话：「打开之后检测是否有外部关联。在上方添加个按钮。点击打开设置
   *   是否自动更新 1、不检测 2、检测后弹窗提示 3、自动更新不提示。
   *   根据这些判断决定是否更新 ooexcel 内的外部的数据引用的更新内容」
   * ★ 2026-10-18 晚：设置入口从 /poc 工具栏**挪到宿主标题栏方形按钮**（用户指令
   *   「引用更新设置 放在标题栏 方形按钮上」）——/poc 顶部按钮随 tab 切换会丢、
   *   且藏得深；标题栏按钮与 💾/ⓘ/着色开关同一排，任何状态都可见。
   * ★ 2026-10-18 深夜：模式**随节点配置持久化**（用户指令「设置好 保存最后的状态」）——
   *   写进 config.ooRefUpdMode 随画布文件走（换机/清缓存不丢）；
   *   localStorage 只作旧数据回退。
   *
   * 三种模式：
   *   off  不检测 —— 打开时原样加载，不比对、不更新（表格显示上次保存的数字）
   *   ask  检测后弹窗提示（默认）—— 打开时比对域值，有差异就弹窗问「是否更新」
   *   auto 自动更新不提示 —— 打开前直接注入（旧行为），运行期上游变更也静默刷新
   * ⚠ 模式必须走 ref：打开文档的 boot effect 只跑一次，依赖 state 会拿到挂载时的旧闭包。
   */
  const REF_UPD_KEY = 'ooRefUpdateMode'
  const [refUpdMode, setRefUpdMode] = useState<RefUpdMode>(() => {
    try {
      const cfg = (useBaseStore.getState().tables[nodeId]?.config ?? {}) as NodeConfig
      if (cfg.ooRefUpdMode === 'off' || cfg.ooRefUpdMode === 'ask' || cfg.ooRefUpdMode === 'auto') return cfg.ooRefUpdMode
    } catch { /* ignore */ }
    try {
      const v = localStorage.getItem(REF_UPD_KEY)
      if (v === 'off' || v === 'ask' || v === 'auto') return v
    } catch { /* ignore */ }
    return 'ask'
  })
  /** 统一的模式写入口：state + ref + localStorage（回退）+ 节点配置（随画布持久化） */
  const applyRefUpdMode = useCallback((m: RefUpdMode) => {
    setRefUpdMode(m)
    try { localStorage.setItem(REF_UPD_KEY, m) } catch { /* ignore */ }
    try { onSetNodeConfig(nodeId, { ooRefUpdMode: m }) } catch { /* ignore */ }
    try { logS('REF-POLICY', '节点 ' + nodeId + ' 更新策略 → ' + m + '（已随节点保存）') } catch { /* ignore */ }
  }, [nodeId, onSetNodeConfig])
  const refUpdModeRef = useRef<RefUpdMode>(refUpdMode)
  useEffect(() => { refUpdModeRef.current = refUpdMode }, [refUpdMode])
  /** 待用户确认的更新提问（由 /poc 弹窗渲染；点「立即更新」回 __ooRefApply；items 供弹窗列新旧值） */
  const [refAsk, setRefAsk] = useState<{ id: number; count: number; names: string[]; items?: { name: string; was: string; now: string }[]; trigger: 'open' | 'live' } | null>(null)
  /** 最近一次检测结果（打开时比对得到），/poc 设置卡里显示「共 N 处 / M 处待更新」 */
  const [refInfo, setRefInfo] = useState<{ total: number; changed: number; at: number } | null>(null)
  /** ★ 标题栏「引用更新」设置卡开合（2026-10-18 用户指令「放在标题栏 方形按钮上」） */
  const [refUpdPopOpen, setRefUpdPopOpen] = useState(false)
  const [refUpdPopPos, setRefUpdPopPos] = useState<{ top: number; right: number }>({ top: 44, right: 16 })
  const refUpdBtnRef = useRef<HTMLButtonElement>(null)
  useEffect(() => {
    if (!refUpdPopOpen) return
    const close = () => setRefUpdPopOpen(false)
    document.addEventListener('click', close)
    return () => document.removeEventListener('click', close)
  }, [refUpdPopOpen])
  /** ★ 打开历史（2026-10-19 用户指令「打开文件左侧加个打开历史记录的按钮 点击下拉最近打开的文件」）：
   *  条目存节点配置 ooHistory（字节随存，≤3 条、单条 ≤4MB）；点条目 = 重开该文件。 */
  const [ooHistOpen, setOoHistOpen] = useState(false)
  const [ooHistPos, setOoHistPos] = useState<{ top: number; left: number }>({ top: 44, left: 16 })
  const ooHistBtnRef = useRef<HTMLButtonElement>(null)
  useEffect(() => {
    if (!ooHistOpen) return
    const close = () => setOoHistOpen(false)
    document.addEventListener('click', close)
    return () => document.removeEventListener('click', close)
  }, [ooHistOpen])
  const ooHistItems = ((table?.config ?? {}) as NodeConfig).ooHistory ?? []
  /** 记录一条打开历史（同名去重置顶；≤3 条；>4MB 不记防配置膨胀） */
  const recordOoHistory = useCallback((name: string, buf: Uint8Array) => {
    try {
      if (buf.length > 4 * 1024 * 1024) return
      const cfg = (useBaseStore.getState().tables[nodeId]?.config ?? {}) as NodeConfig
      const hist = (cfg.ooHistory ?? []).filter((h) => h.name !== name)
      hist.unshift({ name, ts: Date.now(), b64: u8ToB64(buf) })
      onSetNodeConfig(nodeId, { ooHistory: hist.slice(0, 3) })
    } catch { /* 历史记录失败不影响打开 */ }
  }, [nodeId, onSetNodeConfig])
  /** 点历史条目：重开该文件（导入即入库同款：先落 ooBytes64 再打开） */
  const openHistItem = useCallback((h: { name: string; b64: string }) => {
    try {
      const bytes = b64ToU8(h.b64)
      const ext = h.name.toLowerCase().endsWith('.docx') ? 'ooword' : 'oosheet'
      if (ext !== kind && onSwitchKind) onSwitchKind(nodeId, ext)
      void persistOoBytes(bytes, '历史打开 ' + h.name)
      void ooRef.current?.openFile(new File([bytes], h.name), ext)
      setOoHistOpen(false)
    } catch (e) { console.error('[oo] 历史打开失败', e) }
  }, [kind, nodeId, onSwitchKind])
  /** 标题栏「引用更新」按钮的提示文案（含当前模式与最近检测结果） */
  const refUpdBtnTitle = '外部数据引用更新：' + REF_UPD_LABEL[refUpdMode]
    + (refInfo
      ? '（本次打开检测 ' + refInfo.total + ' 处' + (refInfo.changed ? '，' + refInfo.changed + ' 处待更新）' : '，数据已是最新）')
      : '（打开文档后自动检测）')
    + ' —— 点击设置'
  /** 高亮行列开关（2026-10-18 用户指令：选中格所在行/列加底色，阅读长表不串行）。
   *  纯视觉覆盖层（/poc 侧注入，不改文档）；默认关，记忆在 localStorage。 */
  const [rowColHi, setRowColHi] = useState<'1' | '0'>(() => {
    try { return (localStorage.getItem('rowColHighlight') as '1' | '0') || '0' } catch { return '0' }
  })
  /** OO 常用样例下拉（用户 2026-10-01）：主键=按节点类型加载默认样例，▾=选全部；/work/<name> 直读工作区 */
  const [ooSampleOpen, setOoSampleOpen] = useState(false)
  const OO_SAMPLES: { group: string; items: { name: string; ext: 'xlsx' | 'docx' }[] }[] = [
    { group: 'Excel 样例', items: [{ name: 'Excel 工作簿.xlsx', ext: 'xlsx' }] },
    { group: 'Word 样例', items: [{ name: 'Word 测试.docx', ext: 'docx' }, { name: '测试报告.docx', ext: 'docx' }] },
  ]
  const ooLoadSample = useCallback(async (name: string, ext: 'xlsx' | 'docx') => {
    try {
      const r = await fetch('/office-samples/' + encodeURIComponent(name))
      if (!r.ok) throw new Error('HTTP ' + r.status)
      const buf = await r.arrayBuffer()
      const target = ext === 'xlsx' ? 'oosheet' : 'ooword'
      if (target !== kind && onSwitchKind) onSwitchKind(nodeId, target)
      await ooRef.current?.openFile(new File([buf], name), ext)
    } catch (e) {
      console.error('[oo] 样例加载失败', e)
    }
  }, [kind, nodeId, onSwitchKind])
  useEffect(() => {
    if (!ooSampleOpen) return
    const close = () => setOoSampleOpen(false)
    document.addEventListener('click', close)
    return () => document.removeEventListener('click', close)
  }, [ooSampleOpen])
  /** OO 当前打开的文件名（标题栏展示） */
  const [ooOpenedName, setOoOpenedName] = useState('')
  /** ★ 2026-10-17 未保存修改（用户指令「做是否修改了的判断 然后提示保存」）：
   *   由 /poc 状态变化主动推来（onDirtyChange），驱动标题栏「● 未保存」+ 💾 高亮 + 关闭拦截。 */
  const [ooDirty, setOoDirty] = useState(false)
  /** 脏标记镜像（异步链路/订阅回调里读 state 会读到旧值，必须走 ref） */
  const ooDirtyRef = useRef(false)
  /** 换节点 = 新文档，脏状态从零开始（否则上一个节点的「未保存」会串到下一个） */
  useEffect(() => { setOoDirty(false); ooDirtyRef.current = false }, [nodeId, kind])
  /** 可引用的画布参数树（供 /poc 内嵌参数引用面板渲染） */
  const ooRefItems = useMemo(() => {
    const out: { id: string; name: string; fields: { id: string; name: string; value: string; color: string }[] }[] = []
    for (const t of Object.values(tables)) {
      if (t.deleted || t.id === nodeId) continue
      try {
        const comp = computeNode(t.id, tables, records)
        const r0 = comp.rows[0] as Record<string, unknown> | undefined
        out.push({
          id: t.id, name: t.name,
          fields: comp.fields.map((f) => {
            const v = r0?.[f.id]
            // ★ 2026-10-17：色块与文档内引用格底色同源（同一个 refColorOf + 同一套名称口径），
            //   面板上看到什么色，文档里就是什么色 —— 只算一次，oo-core 只负责显示。
            const color = refColorOf(domainNameOf(domainRawKey(t.name, f.name)))
            return { id: f.id, name: f.name, value: v == null ? '' : String(v), color }
          }),
        })
      } catch { /* 单节点计算失败不拖垮整体 */ }
    }
    return out
  }, [tables, records, nodeId])

  /** 域外写回（2026-10-06）：把 ooRefs 的当前值写回 OO 文件。
   *  ★ 落点解析在 core/ooRefSites（字段名 ⇄ 命名区域名的口径转换也在那里）。 */
  /** 编辑器内浮层提示（2026-10-18）：着色开关/刷新这类「点了要看得见反馈」的动作走这里。
   *  过去结果只写进参数引用面板的 refreshMsg —— 面板没开时用户完全看不到，误以为按钮坏了。
   *  ★ 同日用户指令：反馈只保留居中 toast 一处（原 refreshMsg 还会让 /poc 左下角再弹一条，重复）。 */
  const [neToast, setNeToast] = useState('')
  const neToastTimer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const flashToast = useCallback((msg: string) => {
    if (!msg) return
    setNeToast(msg)
    if (neToastTimer.current) clearTimeout(neToastTimer.current)
    neToastTimer.current = setTimeout(() => setNeToast(''), 3600)
  }, [])
  useEffect(() => () => { if (neToastTimer.current) clearTimeout(neToastTimer.current) }, [])
  const ooRefresh = useOoRefRefresh({
    nodeId,
    table,
    tables,
    records,
    readTemplate: async (id) => getTemplateStore().read(id),
    // Excel侧域发现：命名区域 / Table 都由它产出（key = 命名区域名）
    discoverXlsx: async (bytes) => {
      const r = await discoverXlsxFieldsFromBytes(bytes)
      return r.map((f) => ({ key: f.key, ref: f.ref, sheet: f.sheet }))
    },
    // Word 侧域发现：从真实 docx 里列出内容控件（w:sdt），剥掉 kind: 前缀作匹配键
    discoverDocx: async (bytes) => toWordSites(await discoverDocxSdt(bytes)),
    // 报告走 OO 面板的 onStatus（面板内部 say() ⇒ 走居中 toast/诊断，2026-10-18 起不再有独立状态栏文案）
    report: (msg) => { try { flashToast(msg) } catch { /* ignore */ } },
  })
  /** 退场动画一退一进（与 Excel 的 ref-panel-anim-* 同款） */
  const [ooRefLeaving, setOoRefLeaving] = useState(false)
  const [ooRightLeaving, setOoRightLeaving] = useState(false)
  /* ★ 2026-10-23 用户指令「打开节点 侧边栏后面从两侧滑动进入；关闭 侧边栏先滑出再丝滑退出」：
     分镜编排状态——sidesReady=false 时左右栏不渲染（窗体先展开），
     mount 后 ~300ms 置 true 触发两侧滑入；关闭时先 leaving 再收缩窗体。 */
  const [neSidesReady, setNeSidesReady] = useState(false)
  /** 面板定位：贴 .node-editor 窗体外缘（普通）或嵌入左右缘（最大化） */
  const [ooRect, setOoRect] = useState<{ left: number; top: number; right: number; height: number } | null>(null)
  /** ★ 宽度低于「开卡时基线的一半」（用户 2026-10-01：拖到页面一半）→ 引用卡整体隐藏，宽度恢复自动回来 */
  const [ooRefWidthHidden, setOoRefWidthHidden] = useState(false)
  const ooHalfBaseRef = useRef(0)
  useEffect(() => {
    if (!isOO) return
    const measure = () => {
      const el = document.querySelector('.node-editor') as HTMLElement | null
      if (!el) return
      const r = el.getBoundingClientRect()
      setOoRect({ left: r.left, top: r.top, right: r.right, height: r.height })
      // ★ 让位压缩（Excel .ne-refroom 同款）：普通态 + 引用卡开 + 空间不够 → 不准叠在 OO 上，
      //   改为右移压缩窗体宽度（判定按 mask「未让位」几何算，防加类后 rect 变化导致条件翻转振荡）
      //   压缩下限：窗口拖到基线一半 → 卡整体隐藏、窗体恢复；拖宽过基线 → 基线上调
      try {
        if (!headMax && (ooRefOpen || ooRightOpen)) {
          const maskEl = document.querySelector('.node-editor-mask') as HTMLElement | null
          if (maskEl) {
            const mr = maskEl.getBoundingClientRect()
            const pad = 28
            const contentW = mr.width - pad * 2
            const naturalW = Math.min(1180, contentW)
            const unshiftedLeft = mr.left + pad + (contentW - naturalW) / 2
            if (!ooHalfBaseRef.current || window.innerWidth > ooHalfBaseRef.current) ooHalfBaseRef.current = window.innerWidth
            const tooNarrow = window.innerWidth <= ooHalfBaseRef.current / 2
            /* ★ 2026-10-19 右栏同款让位（用户「没有和左侧一样 让ooexcel主窗体退让」）：
               左栏要窗体左侧空间、右栏要右侧空间，各自按需让位。 */
            const needRoom = ooRefOpen && !tooNarrow && unshiftedLeft - 242 < 8
            const needRoomR = ooRightOpen && !tooNarrow && unshiftedLeft + naturalW + 242 > window.innerWidth - 8
            el.classList.toggle('ne-refroom', needRoom)
            el.classList.toggle('ne-refroom-r', needRoomR)
            setOoRefWidthHidden(tooNarrow)
            return
          }
        }
        el.classList.remove('ne-refroom')
        el.classList.remove('ne-refroom-r')
        setOoRefWidthHidden(false)
      } catch { /* 几何兜底失败不致命 */ }
    }
    measure()
    // 最大化/还原切换：立即双次测量（rAF 等布局稳定后补测，消除约 1s 的位置延迟）
    const raf = requestAnimationFrame(() => { measure(); setTimeout(measure, 120) })
    window.addEventListener('resize', measure)
    const t = setInterval(measure, 300) // 窗口变化兜底（原 1200ms 有可感延迟，用户实测反馈）
    return () => { cancelAnimationFrame(raf); window.removeEventListener('resize', measure); clearInterval(t) }
  }, [isOO, headMax, ooRefOpen, ooRightOpen])

  /** OO 💾：编辑器字节 → 模板库入库绑定（ooTemplateId）→ 重开节点自动恢复 */
  /** ★ 2026-10-11 保存链重构：/poc 引擎直读值数组 → 服务端构建 xlsx → 字节入节点配置 */
  const saveDoneRef = useRef(false)
  const valuesArrivedRef = useRef(false)
  /* ★★ 2026-10-17「一步一步查」：保存会话 ID ★★
   *   过去所有日志平铺在 oo-events.log 里，连续点两次保存就分不清哪行属于哪次；
   *   一次保存 20+ 行，跨秒交错，排查只能靠猜时间。现在点 💾 时生成一个短 sid，
   *   随 requestSave 下发到 /poc，链路上每一行都带同一个 [sid +Nms] 前缀 ——
   *   读取脚本可以精确切出「这一次保存」并按耗时看出卡在哪一步。 */
  const ooSidRef = useRef('')
  /** 保存起点（点 💾 的时刻），用于 S16 算整链耗时 */
  const oSaveT0Ref = useRef(0)
  const newSaveSid = () => {
    ooSidRef.current = 'S' + Date.now().toString(36).slice(-6)
    oSaveT0Ref.current = Date.now()
    return ooSidRef.current
  }
  /** 统一日志出口：自动带 sid，detail 放宽到 700（要放得下「行×列｜非空格｜首行」） */
  const logS = useCallback((type: string, detail: string) => {
    try {
      void fetch('/office-log', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type, detail: String(detail).slice(0, 700), sid: ooSidRef.current }),
      }).catch(() => { /* 日志失败不影响主流程 */ })
    } catch { /* ignore */ }
    /* ★ 2026-10-17 操作审计：节点侧的软件反馈（保存/S14/REF-SYNC…）同步进 logs/audit-*.log，
     *   与用户点击、键盘输入按同一时间线排列 —— 复盘时「我点了什么→软件做了什么」一眼对上。 */
    auditFeedback(type, String(detail).slice(0, 500))
  }, [])
  /** 字节落盘（两条保存通道共用）：节点配置 ooBytes64 → 模板库副本 → 作废冻结实例 → 画布落盘 → 日志 */
  const persistOoBytes = useCallback(async (bytes: Uint8Array, via: string, savedPath?: string) => {
    const ext = kind === 'oosheet' ? 'xlsx' : 'docx'
    const name2 = (table?.name || 'OO 文档') + ext
    /* ★ 2026-10-29 保真对比底料：覆盖 ooBytes64 前把旧字节留入 IndexedDB（ooprev-<id>）——
       右键「保真对比（PDF 节点）」用它生成 保存前 vs 保存后 并排 PDF 逐页目检 */
    try {
      const oldB64 = (useBaseStore.getState().tables[nodeId]?.config as { ooBytes64?: string } | undefined)?.ooBytes64
      if (oldB64) {
        const bin = atob(oldB64)
        const old = new Uint8Array(bin.length)
        for (let i = 0; i < bin.length; i++) old[i] = bin.charCodeAt(i)
        if (old.length) {
          const mP = await import('./core/pdfBlobStore')
          await mP.putBlob('ooprev-' + nodeId, old)
        }
      }
    } catch { /* 底料留取失败不影响保存 */ }
    /* ★ 2026-10-11 修「保存后重开没有」：节点配置(ooBytes64=主通道)先落——
     *   实测模板库(IndexedDB)写入阻塞会跳过 onSetNodeConfig，内容进不了节点配置。
     *   改为：节点配置先行，模板库副本降级异步，失败不阻断。 */
    /* ★ 2026-10-31 版式档案建档：来源优先级制——boot「导入即建档」的字节=原版
       （60/70 直排/缩进/边距真实值），persist 的字节=保存产物（含引擎折算错值
       384/256——数量打分反而更高，把好档案覆盖掉，726390/855797 实测翻车）。
       ⇒ 已有档案（无论来源）一律不被保存产物覆盖；仅无档案时兜底建档。 */
    let layoutPatch: Record<string, unknown> = {}
    try {
      if (kind === 'ooword') {
        const cfg0 = useBaseStore.getState().tables[nodeId]?.config as { ooLineProfile?: OoLayoutProfile } | undefined
        if (!cfg0?.ooLineProfile) {
          const prof = await extractOoLayoutProfile(bytes)
          if (prof && (prof.lineList.some((x) => x.line) || prof.pageBreakIdx.length || prof.sectionProps)) {
            layoutPatch = { ooLineProfile: prof }
            try { logS('S14-P', '版式档案兜底已建: line=' + prof.lineList.filter((x) => x.line).length + ' 分页=' + prof.pageBreakIdx.length + ' 页面=' + (prof.sectionProps ? '√' : '默认')) } catch { /* ignore */ }
          }
        }
      }
    } catch { /* 建档失败不影响保存 */ }
    onSetNodeConfig(nodeId, { ooBytes64: u8ToB64(bytes), ooTemplateName: name2, ...layoutPatch } as never)
    try { logS('S14', '节点 ' + nodeId + ' 已写入 ooBytes64=' + bytes.length + 'B 通道=' + via) } catch (eL14) {}
    getTemplateStore().save(name2, bytes).then((meta) => {
      onSetNodeConfig(nodeId, { ooTemplateId: meta.id })
    }).catch((e) => console.error('[oo] 模板库副本失败（不影响节点内容）:', e))
    /* ★ 2026-10-11 用户报「没打开保存的」：保存后作废该节点的冻结实例——
     *   否则重开节点时 boot 前两步会接管保存前的旧冻结现场，绕过 ooBytes64 加载分支 */
    try { discardFrozen(nodeId) } catch { /* 无冻结则忽略 */ }
    /* ★ 2026-10-11 用户指令「随画布文件保存」：OO 保存后立即触发画布落盘——
     *   否则节点配置只在内存，不点应用保存就彻底关闭 = 内容丢失 */
    try {
      const r = useBaseStore.getState().saveNow() as unknown
      if (r && typeof (r as Promise<unknown>).then === 'function') {
        void (r as Promise<unknown>)
          .then(() => { try { logS('S15', '画布落盘完成（saveNow Promise resolved）') } catch (eS15) { /* 忽略 */ } })
          .catch((e) => { try { logS('S15-E', '❌ 画布落盘失败: ' + String(e).slice(0, 160)) } catch (eS15b) { /* 忽略 */ } })
      } else {
        try { logS('S15', '画布落盘已触发（同步返回）') } catch (eS15c) { /* 忽略 */ }
      }
    } catch (eSN) {
      console.error('[oo] 画布落盘失败', eSN)
      try { logS('S15-E', '❌ 画布落盘抛错: ' + String(eSN).slice(0, 160)) } catch (eS15d) { /* 忽略 */ }
    }
    /* ★ 2026-10-17 用户指令「换个保存的地方，看看保存的是不是我修改的」：
     *   每次保存额外留一份带时间戳的文件到 sample-template/保存留档/，
     *   并把【内容摘要】（非空单元格数 + 前几项文本）直接写进成功日志——
     *   用户不用打开文件就能看出存的是不是刚才改的东西。 */
    let archivePath = ''
    let digestText = ''
    try {
      const ar = await fetch('/office-offline/archive', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: name2, dataBase64: u8ToB64(bytes), via }),
      }).then((x) => x.json())
      if (ar && ar.ok) {
        archivePath = String(ar.path || '')
        digestText = (ar.digest && ar.digest.text) ? String(ar.digest.text) : ''
      }
    } catch (eAR) { /* 留档失败不阻断主流程 */ }
    try { ooRef.current?.ackSaved() } catch (eAS) { /* 忽略 */ }
    /* ★ 2026-10-17 用户指令：保存成功要把「保存路径」写进日志 + 诊断抽屉能看见 */
    const okMsg = '✅ 保存成功（' + via + '）｜' + bytes.length + 'B 已写入节点并随画布落盘'
      + (archivePath ? '｜留档 ' + archivePath : '')
      + (digestText ? '｜内容摘要：' + digestText : '')
    try { ooRef.current?.opLog(okMsg) } catch (eOL3) { /* 忽略 */ }
    try { logS('SAVE-NODE-OK', '节点 ' + nodeId + ' 通道=' + via + ' 留档=' + (archivePath || '-') + ' ooBytes64=' + bytes.length + 'B 已写入 + 画布已落盘 ' + digestText) } catch (eL4) { /* 忽略 */ }
    /* ★ 2026-10-17 S16：整链总耗时（从点 💾 算起）——
     *   以后「保存慢」不用再靠体感，日志里直接有数字。 */
    try {
      logS('S16', '保存完成 总耗时=' + (oSaveT0Ref.current ? Date.now() - oSaveT0Ref.current : -1) + 'ms 通道=' + via + ' 产物=' + bytes.length + 'B')
    } catch (eS16) { /* 忽略 */ }
  }, [nodeId, kind, table?.name, onSetNodeConfig])
  /* ══════════════════════════════════════════════════════════════════════
   * 域外引用 → 文档（2026-10-17 用户：「改了引用的节点的参数 为什么 oo 节点中的
   * 数据没有更新」）
   *
   * 查清的三处断点（都有实据）：
   *   ① 自动触发**从未接线**：旧代码把 ooRefConsumers.register 写在
   *      NodeEditor.stashbak.bak 里，现行 NodeEditor.tsx 既没 import 也没 register
   *      ⇒ 上游改了不会有任何东西被调用。
   *   ② 就算调了，写回目标是**模板库副本**，而节点显示读的是 config.ooBytes64
   *      ⇒ 写「成功」了，节点内容不变。
   *   ③ 更致命：提取通道（当前唯一可用保存链）把文件重建成「单表 Sheet1 + 值 + 公式」，
   *      设计里的隐藏表『域数据』与定义名称全被吃掉。实测该节点 ooBytes64（3609B）：
   *      sheets=['Sheet1']、definedName=0，而格子里还留着 `<f>数据对照表_键</f>`
   *      ⇒ 文件里已经没有「域」，刷新域值找不到落点，公式也永远解析不出新值。
   *
   * 修法（改③为根，① 顺带补齐）：
   *   打开文档前按 ooRefs 的**当前值**把『域数据』+ 定义名称补回去（core/ooDomains），
   *   并把引用格的缓存值一并改成当前值 —— 引擎重算与不重算都显示对的数。
   *   纯补丁：公式、样式、其它表全部原样保留。
   * ★ 为什么是「打开时注入」而不是「写回时改模板」：
   *   每次打开都按现值重造，天然幂等；提取通道下次把表吃掉也无所谓。
   * ══════════════════════════════════════════════════════════════════════ */

  /** 引用值条目：按**实时 store**（不是 props 快照）取，避免通知到达时数据还没刷新 */
  const buildRefEntries = useCallback(() => {
    const st = useBaseStore.getState()
    const cfg = (st.tables[nodeId]?.config ?? {}) as NodeConfig
    const refs = (cfg.ooRefs ?? []) as { nodeId: string; fieldId: string }[]
    if (!refs.length) return []
    const items = refs.map((r) => {
      const nt = st.tables[r.nodeId]
      const f = (nt?.fields ?? []).find((x) => x.id === r.fieldId)
      let v: unknown
      try {
        const comp = computeNode(r.nodeId, st.tables, st.records)
        v = (comp.rows[0] as Record<string, unknown> | undefined)?.[r.fieldId]
      } catch { v = undefined }
      return { nodeName: nt?.name || r.nodeId, fieldName: f?.name || r.fieldId, value: v }
    })
    return buildDomainEntries(items)
  }, [nodeId])

  /**
   * 打开前注入（给 OnlyOfficeNodeEditor 的回调）：无引用 / Word 侧原样返回。
   *
   * ★ 2026-10-18 起受「引用更新」策略控制（forOpen=true 时）：
   *   off  → 不比对、不注入，按保存内容打开（button 上写「不检测」）
   *   ask  → 先 diffDomainValues 比对，有差异就**不注入**并弹窗询问（Excel 的
   *          「此工作簿含外部链接，是否更新」同款行为）；无差异照常打开
   *   auto → 直接注入（旧行为，不打扰用户）
   *   刷新域值（forOpen 缺省）永远真注入 —— 那是用户明确点了「更新」。
   */
  const injectRefDomains = useCallback(async (u8: Uint8Array, opts?: { forOpen?: boolean }): Promise<{ bytes: Uint8Array; note?: string; changed?: { name: string; was: string; now: string }[] }> => {
    try {
      if (kind !== 'oosheet') return { bytes: u8, note: '' }
      /* ★ 打开路径（含不检测/检测后提示）先摘掉 fullCalcOnLoad：旧文件残留的该标志
       *   会让引擎把刚打开的文档标成「已修改」⇒ 没改任何内容、关闭却提示保存。 */
      let base = u8
      if (opts?.forOpen) {
        try { base = await stripFullCalcFlag(u8) } catch { base = u8 }
      }
      const entries = buildRefEntries()
      if (!entries.length) return { bytes: base, note: '' }
      const mode = refUpdModeRef.current
      let out = base
      let note = ''
      let skipped = false
      let changed: { name: string; was: string; now: string }[] | undefined
      if (opts?.forOpen && mode !== 'auto') {
        /* 检测（也是用户说的「打开之后检测是否有外部关联」）：
         * 只有真的算出差异才打扰用户；没差异就安静打开。
         * ★ 2026-10-18 深夜口径改版：diff 直接比对公式格缓存（用户看到的数）——
         *   域表缺失但缓存=当前值 → 不弹窗，静默重建（修复「没改参数也提示」）。 */
        const d = await diffDomainValues(base, entries)
        changed = d.changed
        setRefInfo({ total: d.total, changed: d.changed.length, at: Date.now() })
        if (mode === 'off') {
          skipped = true
          note = '外部数据更新：不检测（共 ' + d.total + ' 处引用，按保存内容打开）'
        } else if (d.changed.length) {
          skipped = true
          setRefAsk({
            id: Date.now(),
            count: d.changed.length,
            names: d.changed.slice(0, 6).map((c) => c.name),
            items: d.changed.slice(0, 6),
            trigger: 'open',
          })
          note = '检测到 ' + d.changed.length + '/' + d.total + ' 处外部数据有更新 → 等待用户确认'
        } else if (!d.hasSheet) {
          /* 值无差异但域表缺失（保存通道剥过）→ 静默重建，不打扰（显示值不会变） */
          skipped = false
          try { logS('REF-POLICY', '节点 ' + nodeId + ' 域表缺失但值无差异 → 静默重建（共 ' + d.total + ' 处）') } catch { /* 忽略 */ }
        } else {
          skipped = true
          note = '外部数据已是最新（' + d.total + ' 处，无需更新）'
        }
        try { logS('REF-POLICY', '节点 ' + nodeId + ' 打开检测 mode=' + mode + ' 共=' + d.total + ' 差异=' + d.changed.length) } catch { /* 忽略 */ }
      }
      if (!skipped) {
        if (!changed) {
          /* 刷新/auto 路径：补一次 diff，供「更新内容」提示与诊断抽屉 */
          try { changed = (await diffDomainValues(base, entries)).changed } catch { changed = undefined }
        }
        const r = await injectDomains(base, entries)
        out = r.bytes
        /* auto 模式不比对，但仍告诉按钮「共 N 处引用」，用户点开设置卡能看清当前状态 */
        if (opts?.forOpen) setRefInfo({ total: entries.length, changed: 0, at: Date.now() })
        note = r.why
          ? '域注入跳过：' + r.why
          : '域数据 ' + r.injected + ' 条' + (r.createdSheet ? '（新建隐藏表）' : '') + '，公式缓存重算 ' + (r.cachePatched || 0) + ' 格'
            + (r.markedCells ? '，引用格着色 ' + r.markedCells + ' 格' : '')
            + (r.recalcNotes?.length ? '（' + r.recalcNotes.slice(0, 3).join(' ; ') + '）' : '')
        try {
          logS('REF-INJECT', '节点 ' + nodeId + ' ' + note + (r.sites?.length ? '｜' + r.sites.slice(0, 4).join(' ; ') : ''))
        } catch (eRI0) { /* 忽略 */ }
      }
      /* ★ 2026-10-18 着色改版：顺带提取「引用格清单 + 表格几何」给覆盖层
       *   （颜色不再写进文件，开关纯视觉 —— 见 toggleRefMark 注释）。
       *   ⚠ 不注入时也要提取：覆盖层数据来自原始字节里的公式，与域表是否存在无关。 */
      try {
        const ex = await extractRefMarks(out, entries)
        if (ex.marks.length) setRefMarksData({ marks: ex.marks, geom: ex.geom })
      } catch { /* 覆盖层数据失败不影响打开 */ }
      return { bytes: out, note, changed }
    } catch (e) {
      const msg = '域注入失败：' + String((e as Error)?.message ?? e).slice(0, 120)
      try { logS('REF-INJECT-E', msg) } catch (eRI1) { /* 忽略 */ }
      return { bytes: u8, note: msg }
    }
  }, [kind, buildRefEntries, logS, nodeId])

  const applyOoDirty = useCallback((v: boolean) => {
    ooDirtyRef.current = v
    setOoDirty(v)
  }, [])

  /** 刷新中（防连点重复注入） */
  const [refSyncing, setRefSyncing] = useState(false)
  /** 「刷新域值」：读节点当前字节 → 注入现值 → 落盘 → 重开文档（手动/自动/弹窗确认共用）。
   *  why='apply'：用户在「检测后提示」弹窗里点了「立即更新」——**绕过脏文档保护**：
   *  打开含 fullCalcOnLoad 的文件时引擎重算会把文档标成「已修改」，若按普通刷新拦截，
   *  用户点更新就永远被「有未保存修改」挡住（2026-10-19 用户实测「点击更新 为什么没更新」）。
   *  弹窗文案已写明「未保存的修改请先保存」，用户确认即视为接受重开丢弃。 */
  const refreshRefDomains = useCallback(async (why: 'manual' | 'auto' | 'toggle' | 'apply'): Promise<boolean> => {
    /* 结果一律走居中 toast（auto 静默——后台同步不弹）；原 /poc 左下角提示已停用（重复） */
    const say = (msg: string) => { if (why !== 'auto') flashToast(msg) }
    if (!isOO || kind !== 'oosheet') { say('仅 Excel 类 OO 节点支持刷新域值'); return false }
    if (refSyncing) { say('正在刷新中，请稍候…'); return false }
    try { logS('REF-SYNC', '节点 ' + nodeId + ' 触发=' + why + ' 脏=' + ooDirtyRef.current) } catch { /* 忽略 */ }
    if (ooDirtyRef.current && why !== 'apply') {
      say('⚠ 文档有未保存修改 —— 先点 💾 保存，再切换/刷新')
      try { logS('REF-SKIP', '节点 ' + nodeId + ' 跳过刷新：文档有未保存修改 触发=' + why) } catch { /* 忽略 */ }
      return false
    }
    const st = useBaseStore.getState()
    const cfg = (st.tables[nodeId]?.config ?? {}) as NodeConfig & { ooBytes64?: string }
    let cur: Uint8Array | null = null
    if (!cur && cfg.ooBytes64) {
      try {
        const bin = atob(cfg.ooBytes64)
        const u8 = new Uint8Array(bin.length)
        for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
        if (u8.length) cur = u8
      } catch { cur = null }
    }
    if (!cur && cfg.ooTemplateId) {
      try { cur = await getTemplateStore().read(cfg.ooTemplateId) } catch { cur = null }
    }
    /* ★ 2026-10-18 用户反馈「引用格着色点了没开没关」：节点没存过字节时（新建后一直在编辑器里改、
     *   还没点过保存），ooBytes64 与模板库都为空 ⇒ 旧逻辑直接判「没有可刷新内容」，
     *   开关看着像坏的（提示还只在参数引用面板里，用户看不到）。
     *   这里补一条：向当前打开的编辑器要一次工作区字节（与 ⤓ 导出同一条通道），
     *   拿得到就正常重注入——用户不必先存一次再切。 */
    if (!cur && ooRef.current?.currentBytes) {
      try {
        const live = await ooRef.current.currentBytes()
        if (live && live.length) {
          cur = live
          try { logS('REF-SYNC', '节点 ' + nodeId + ' 节点无字节，改取当前工作区 ' + live.length + 'B') } catch { /* 忽略 */ }
        }
      } catch { /* 取不到就走下面的「无内容」分支 */ }
    }
    if (!cur || !cur.length) {
      say('⚠ 没有可用的文档内容 —— 先 📂 打开文件或 💾 保存一次')
      try { logS('REF-SYNC-E', '节点 ' + nodeId + ' 无可用字节（ooBytes64 / 模板库 / 当前工作区都为空）') } catch { /* 忽略 */ }
      return false
    }
    setRefSyncing(true)
    try {
      const r = await injectRefDomains(cur)
      if (!r.bytes || r.bytes === cur) {
        say(r.note || '没有可注入的引用值')
        return false
      }
      await persistOoBytes(r.bytes, '引用刷新（' + (why === 'auto' ? '上游变更' : why === 'apply' ? '弹窗确认更新' : '手动') + '）')
      /* ★ 2026-10-18 用户指令「更新数据后要在提示里 提示更新的内容 诊断信息也要体现」：
       *   toast 与诊断抽屉（ⓘ）都列出「名字：旧值 → 新值」明细。 */
      const chg = r.changed ?? []
      const detail = chg.length
        ? chg.slice(0, 5).map((c) => c.name + '：' + c.was + ' → ' + c.now).join('；') + (chg.length > 5 ? ' 等' : '')
        : ''
      say('✅ 已同步引用值：' + (r.note || '') + (detail ? '｜更新内容：' + detail : '（值无变化）'))
      try {
        ooRef.current?.opLog('🔄 引用更新明细（' + chg.length + ' 处值变化）：' + (detail || '无值变化（仅重建域数据表）'))
      } catch { /* 编辑器未挂载 */ }
      try {
        logS('REF-DETAIL', '节点 ' + nodeId + ' 更新明细：' + (chg.map((c) => c.name + ' ' + c.was + '→' + c.now).join(' ; ') || '无值变化'))
      } catch { /* ignore */ }
      try {
        if (!ooRef.current?.openFile) throw new Error('编辑器未挂载')
        const ab = r.bytes.buffer.slice(r.bytes.byteOffset, r.bytes.byteOffset + r.bytes.byteLength) as ArrayBuffer
        await ooRef.current.openFile(new File([ab], (table?.name || 'OO 文档') + '.xlsx'))
      } catch {
        /* 没重开 ≠ 没生效：字节已落盘，下次打开节点就是新样式。别让用户以为白点了。 */
        say('已写入节点字节，但没能自动重开文档 —— 重新打开节点即可看到效果')
      }
      return true
    } finally {
      setRefSyncing(false)
    }
  }, [isOO, kind, nodeId, injectRefDomains, persistOoBytes, table?.name, refSyncing, flashToast])

  /* 自动同步：上游引用节点的记录一变 ⇒ 重注入（500ms 去抖，连续改只跑一次）
   * ★ 用 zustand subscribe 而不是逐个编辑入口挂回调：编辑入口有 4 处以上
   *   （表格内联、属性面板、TableForm、将来的宏），挂漏一处就是「这次没更新」。 */
  const refsKey = useMemo(() => {
    const refs = ((table?.config as NodeConfig | undefined)?.ooRefs ?? []) as { nodeId: string; fieldId: string }[]
    return refs.map((r) => r.nodeId + '#' + r.fieldId).sort().join('|')
  }, [table?.config])

  /* ★ 订阅用「函数 ref」而不是把 refreshRefDomains 放进 deps：
   *   实测踩坑 —— 上游 setCell 会同时让 App 重渲染 ⇒ persistOoBytes/refreshRefDomains
   *   的依赖（onSetNodeConfig 等）身份变化 ⇒ effect 重跑 ⇒ cleanup 里的 clearTimeout
   *   把刚排上的 500ms 去抖计时器**清掉**，于是「收到通知但什么也没发生」。
   *   日志里的形态：REF-WATCH 紧接着又打一条（重订阅），但永远等不到 REF-SYNC。 */
  /** 「立即更新」：用户点了 /poc 弹窗里的按钮 → 走 'apply' 刷新（注入 + 落盘 + 重开）。
   *  ★ why='apply' 绕过脏文档保护（见 refreshRefDomains 注释：fullCalcOnLoad 重算会把
   *    刚打开的文档标成已修改，普通刷新会被「有未保存修改」拦死 → 用户点更新没反应）。 */
  const applyRefAsk = useCallback(async () => {
    setRefAsk(null)
    try { logS('REF-APPLY', '节点 ' + nodeId + ' 用户确认立即更新外部数据（弹窗） 脏=' + ooDirtyRef.current) } catch { /* ignore */ }
    const ok = await refreshRefDomains('apply')
    if (!ok) {
      /* 刷新链任何一环拒绝都必须让用户看见原因（各分支自带 toast，这里兜底一条） */
      try { logS('REF-APPLY-E', '节点 ' + nodeId + ' 刷新链未执行成功（原因见 REF-SYNC/REF-SKIP/REF-SYNC-E）') } catch { /* ignore */ }
    }
  }, [nodeId, refreshRefDomains])

  /** /poc 设置卡切换模式 → 走统一写入口（state + localStorage + 节点配置持久化） */
  const setRefUpdModeFromEditor = useCallback((m: string) => {
    if (m === 'off' || m === 'ask' || m === 'auto') applyRefUpdMode(m)
  }, [applyRefUpdMode])

  /** 运行期（文档已打开）上游变更 → 按当前模式分流：
   *   off  → 只记日志，绝不改文档
   *   ask  → 比对当前工作区字节，有差异就弹窗（不静默重开，避免吞掉用户没保存的滚动/选中）
   *   auto → 静默刷新（旧行为） */
  const liveRefCheck = useCallback(async () => {
    const mode = refUpdModeRef.current
    if (mode === 'off') {
      try { logS('REF-POLICY', '节点 ' + nodeId + ' 上游变更：不检测，已忽略') } catch { /* ignore */ }
      return
    }
    if (mode === 'auto') { await refreshRefDomains('auto'); return }
    if (ooDirtyRef.current) {
      flashToast('⚠ 文档有未保存修改 —— 外部数据更新已暂缓，保存后再试')
      return
    }
    try {
      const live = await ooRef.current?.currentBytes?.()
      const entries = buildRefEntries()
      if (!live || !live.length || !entries.length) return
      const d = await diffDomainValues(live, entries)
      setRefInfo({ total: d.total, changed: d.changed.length, at: Date.now() })
      if (!d.changed.length) return
      setRefAsk({
        id: Date.now(),
        count: d.changed.length,
        names: d.changed.slice(0, 6).map((c) => c.name),
        items: d.changed.slice(0, 6),
        trigger: 'live',
      })
      try { logS('REF-POLICY', '节点 ' + nodeId + ' 上游变更：' + d.changed.length + ' 处差异 → 弹窗询问') } catch { /* ignore */ }
    } catch { /* 比对失败就当没变化，不打扰 */ }
  }, [nodeId, refreshRefDomains, buildRefEntries, flashToast])
  const liveRefRef = useRef(liveRefCheck)
  useEffect(() => { liveRefRef.current = liveRefCheck }, [liveRefCheck])

  /** 引用格着色开关（2026-10-18 用户指令改版：**纯视觉开关**）
   *  旧实现把颜色写进 xlsx 再重载文档 —— 用户实测「点开关画面不变、要手动重开」，
   *  且重开会丢滚动位置/选中格/撤销栈，被明确否决（「不是只是颜色变了」）。
   *  新实现：颜色不写文件；开关只翻状态并经 refMarks prop 推给 /poc 画覆盖层
   *  （与行列高亮同机制），即时生效、不重开、不脏文档。 */
  const toggleRefMark = useCallback(() => {
    const nv: '1' | '0' = refMarkOnRef.current === '1' ? '0' : '1'
    refMarkOnRef.current = nv
    setRefMarkOn(nv)
    try { localStorage.setItem('refMarkColors', nv) } catch { /* ignore */ }
    try { logS('REF-MARK', '节点 ' + nodeId + ' 引用格着色 → ' + (nv === '1' ? '开' : '关') + '（视觉层）') } catch { /* ignore */ }
    flashToast(nv === '1' ? '🎨 引用格着色：开' : '🎨 引用格着色：关')
  }, [nodeId, flashToast])

  /** 高亮行列开关（2026-10-18 用户指令）：只翻状态 + 推消息给 /poc（纯视觉层，不动文档） */
  const toggleRowColHi = useCallback(() => {
    setRowColHi((prev) => {
      const nv: '1' | '0' = prev === '1' ? '0' : '1'
      try { localStorage.setItem('rowColHighlight', nv) } catch { /* ignore */ }
      try { logS('HI-RC', '节点 ' + nodeId + ' 高亮行列 → ' + (nv === '1' ? '开' : '关')) } catch { /* ignore */ }
      return nv
    })
  }, [nodeId])

  useEffect(() => {
    if (!isOO || kind !== 'oosheet') return
    const refs = ((table?.config as NodeConfig | undefined)?.ooRefs ?? []) as { nodeId: string }[]
    const upstream = [...new Set(refs.map((r) => r.nodeId))]
    if (!upstream.length) return
    let timer: ReturnType<typeof setTimeout> | null = null
    try { logS('REF-WATCH', '节点 ' + nodeId + ' 开始监听上游 ' + upstream.join(',') + '（' + refs.length + ' 个引用）') } catch { /* 忽略 */ }
    const unsub = useBaseStore.subscribe((s, prev) => {
      if (!s || !prev || s.records === prev.records) return
      if (!upstream.some((id) => s.records[id] !== prev.records[id])) return
      if (timer) clearTimeout(timer)
      timer = setTimeout(() => {
        timer = null
        void liveRefRef.current().catch((e) => {
          try { logS('REF-SYNC-E', '自动同步异常: ' + String((e as Error)?.message ?? e).slice(0, 200)) } catch { /* 忽略 */ }
        })
      }, 500)
    })
    return () => { unsub(); if (timer) clearTimeout(timer) }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOO, kind, refsKey])

  const ooSaveFromValues = useCallback(async (payload: { name: string; rows: unknown[][]; mode?: string }) => {
      /* ★ 2026-10-17 用户指令「点保存要看得见反馈」：
       *   过去 rows 为空时这里静默 return，用户点了保存毫无动静（"还是没法保存"的经典体感）。
       *   现在一律先给一条可见日志，区分「无内容可存 / 正在保存 / 成功 / 失败」四种结果。 */
      if (!isOO || !nodeId || !payload) { saveDoneRef.current = true; return }
      try { ooRef.current?.opLog('💾 正在保存…') } catch (eOL1) {}
      try { logS('SAVE-BTN', '节点 ' + nodeId + ' 开始保存，目标=' + payload.name + ' 提取行数=' + (payload.rows ? payload.rows.length : 0) + ' sid=' + ooSidRef.current) } catch (eL7) {}
      /* ★ 2026-10-17 S9a：收到提取值的「形状」——只看行数会误判。
       *   行数 2 非空 0 = 读到了区域却取不到值；非空 4 公式 1 = 真读到内容且公式在。 */
      try {
        const st = rowStats(payload.rows)
        logS('S9a', '提取内容形状: ' + st.rows + '行×' + st.cols + '列｜非空格=' + st.vals + ' 公式=' + st.formulas + '｜首行=' + st.firstRow
          + (st.rows > 0 && st.vals === 0 ? '｜⚠ 有行全空' : ''))
      } catch (eS9a) { logS('S9a-E', '统计异常: ' + String(eS9a).slice(0, 120)) }
    setOoSavling(true)
    try {
      /* ★★ 2026-10-30 顺序重构（根治「forcesave 锁死 callCommand」）★★
       *  实测：forcesave 期间/之后引擎 callCommand 全部静默失效（fn 不执行），
       *  styledExtract 3 连败 → 掉进纯文本降级（样式全丢，10:44 留档实锤）。
       *  改为：ooword 且 dirty=true 时（主通道必被硬门拒）根本不 forcesave——
       *  引擎空闲态先做样式化提取（一次成功），提取成功即跳过主通道。 */
      let skipMainChannel = false
      let styledPre: unknown[] | null = null
      if (kind === 'ooword') {
        let dirtyPre = true
        try { dirtyPre = await ooRef.current!.isModified() } catch { dirtyPre = true }
        if (dirtyPre) {
          try { logS('S11-PRE', 'dirty=true → 跳过 forcesave，空闲态先样式化提取') } catch { /* ignore */ }
          try {
            const se = await ooRef.current?.styledExtract?.() ?? null
            styledPre = se?.paras ?? null
            if (styledPre) {
              skipMainChannel = true
              try { logS('S11-PRE2', '空闲态提取成功: ' + styledPre.length + ' 段（跳过主通道）') } catch { /* ignore */ }
            }
          } catch { /* 提取失败走原流程 */ }
        }
      }
      /* ★ 2026-10-17 主通道改为「OO 自身通道」（用户拍板：要保住原样式）：
       *   asc_Save 推送变更 → DocService forcesave → 轮询 x2t 产物，由 OO 亲自把文档转回 xlsx/docx。
       *   ★ 为什么它当主通道：提取通道是「读值数组 → 服务端重建」，重建出来的 xlsx 只有
       *     Sheet1 + 值 + 公式，字体/颜色/边框/列宽/合并/多表/图表全丢；OO 通道则原样导出，格式全保留。 */
      let bytes: Uint8Array | null = null
      let via = ''
      let savedPath = ''
      /* ★ 2026-10-30 版式档案提升到共享作用域：样式化重建**和纯文本降级**都要用——
         styledExtract 全挂时掉纯文本，段落版式（行距/分页/页面设置）仍可从档案补回
         （11:46 留档实锤：纯文本产物默认边距无行距，用户看到「段落对不起来」）。 */
      const layoutProfShared = (useBaseStore.getState().tables[nodeId!]?.config as { ooLineProfile?: OoLayoutProfile } | undefined)?.ooLineProfile
      const sharedLineList = layoutProfShared?.lineList ?? []
      const sharedPageBreakIdx: number[] = layoutProfShared?.pageBreakIdx ?? []
      const sharedSectionProps = layoutProfShared?.sectionProps
      const sharedDefaultSpacing = layoutProfShared?.defaultSpacing
      const sharedDefaultFontSize = layoutProfShared?.defaultFontSize
      const sharedFooter = layoutProfShared?.footer
      try { if (sharedLineList.some((x) => x.line)) logS('S12-P0', '版式档案可用(共享): line=' + sharedLineList.filter((x) => x.line).length) } catch { /* ignore */ }
      try {
        if (skipMainChannel) {
          try { logS('S11-SKIP', '主通道跳过（已有空闲态提取结果）') } catch { /* ignore */ }
        } else {
        const fb = await ooRef.current?.saveBack()
        /* ★★ 2026-10-17 假成功拦截（这是「还是不行」的真凶）★★
         *   Editor.bin 是「打开时用 x2t 从源文件转出来的」，本身不含编辑内容（见 oo-core 959-965 行）。
         *   所以主通道转回来的产物 == 打开时的原文件；源是空白新建时就是一张空表。
         *   原实现只要「有字节」就认为主通道成功 —— 实测 18:59 那次：提取通道已拿到 123，
         *   主通道却回了个 6258B 的空表把它整个覆盖，用户重开当然是空白。
         *   两道校验：①产物内容为空 ②产物与打开时字节完全相同（=没吃到编辑，且确实改过）。 */
        if (fb && fb.bytes && fb.bytes.length) {
          const extractReal = rowsHaveContent(payload.rows)
          /* ★ 2026-10-29 dirty 无条件问（不再只在 sameAsSource 时）——它是主通道硬门：
             有未保存编辑 ⇒ 产物不可能含编辑（forcesave 架构性 400），必须走样式化重建 */
          let dirty = true
          try { dirty = await ooRef.current!.isModified() } catch { dirty = true }
          const verdict = decideMainChannel(fb, payload.rows, dirty)
          if (!verdict.use && verdict.reject) {
            try { logS('S11b', '主通道判定为假成功→丢弃：' + verdict.reject + '；产物=' + fb.bytes.length + 'B → 改用提取通道') } catch (eL11b) { /* 忽略 */ }
            try { ooRef.current?.opLog('⚠ 主通道产物无效（' + verdict.reject + '），改用提取通道保住内容') } catch (eOL11b) { /* 忽略 */ }
          } else if (verdict.use) {
            bytes = fb.bytes
            via = 'OO 自身通道（保留样式）'
            /* ★ 2026-10-17 S11c：裁决通过也要留痕——这样才能区分
             *   「主通道根本没出产物（S11-E）」和「出了产物且被判可用」。 */
            try {
              logS('S11c', '主通道裁决通过: 产物=' + fb.bytes.length + 'B｜' + (fb.digest && fb.digest.text ? fb.digest.text : '无摘要')
                + '｜sameAsSource=' + (fb.sameAsSource ? 'yes' : 'no') + '｜dirty=' + (dirty ? 'yes' : 'no')
                + '｜提取有内容=' + (extractReal ? 'yes' : 'no') + ' → 采用主通道')
            } catch (eS11c) { /* 忽略 */ }
          }
        }
        } /* ← skipMainChannel else 包裹结束 */
      } catch (eFB) {
        try { ooRef.current?.opLog('⚠ OO 自身通道异常，回退提取通道：' + String(eFB).slice(0, 100)) } catch (eOL9) { /* 忽略 */ }
      }
      /* ★ 2026-10-29 移除「假成功后 1.2s 延时重取主通道」：forcesave 架构性 400（无会话），
         重试取回的仍是同一份不含编辑的产物，纯浪费 2~3s（「正在保存文件卡住」来源之一）。
         dirty 硬门 + 样式化重建已接管编辑捕获。 */
      /* 兜底：OO 通道没拿到 → 用提取到的值数组重建（至少保住值与公式，样式会丢）
       * ★ 2026-10-22 修 -85：仅 oosheet 可走（值数组重建只产 xlsx）——
       *   ooword 走这里会把 xlsx 字节存进 .docx，下次打开必报「内容与扩展名不匹配」。 */
      if (!bytes && payload.rows && payload.rows.length && kind === 'oosheet') {
        try { logS('S12', '降级提取通道：rows=' + payload.rows.length + ' 目标=' + payload.name) } catch (eL12) {}
        const r = await fetch('/office-offline/saveFromValues', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: payload.name, rows: payload.rows }),
        }).then((x) => x.json())
        if (r && r.ok && r.dataBase64) {
          const bin = atob(r.dataBase64)
          const b2 = new Uint8Array(bin.length)
          for (let i = 0; i < bin.length; i++) b2[i] = bin.charCodeAt(i)
          bytes = b2
          via = '提取通道（仅值+公式，样式丢失）'
          savedPath = r.path ? String(r.path) : ''
          try { logS('S13', '提取通道产物=' + b2.length + 'B path=' + (savedPath || '-')) } catch (eL13) {}
        } else {
          try { logS('S12-E', '提取通道失败 ok=' + (r && r.ok) + ' err=' + ((r && r.err) || '-')) } catch (eL12e) {}
        }
      }
      /* ★ 2026-10-23 修「ooword 保存并关闭后重开还是空白」：主通道架构性无产物
       *   （自研 socket 无 DS 会话，forcesave 恒 400、Editor.bin 不含编辑内容），
       *   0.00.32 又禁了值数组重建（产 xlsx 污染 docx 报 -85）⇒ docx 两条路全断。
       *   现补 docx 专用兜底：按提取到的段落文本用 docx 库重建
       *   （/poc readFn 对 .docx 提取 rows=[[段落文本],...]）——保文字，
       *   样式丢失，与 oosheet 提取通道同一取舍。 */
      if (!bytes && payload.rows && payload.rows.length && kind === 'ooword') {
        /* ★★ 2026-10-31 原生保存链（v4 主通道）★★：引擎 asc_DownloadAs(DOCX) 导出
           「当前修改状态」——OO 自己序列化=格式 100%（底纹/行距/缩进/表格全保留），
           架构性替代「提取+重建」（提取协议永远追不上引擎的样式维度）。 */
        try {
          if (!skipMainChannel && ooRef.current) {
            try { logS('S11N', '原生导出通道尝试（asc_DownloadAs DOCX）…') } catch { /* ignore */ }
            const pocWin = (document.querySelector('iframe') as HTMLIFrameElement | null)?.contentWindow
            if (pocWin) {
              const nativeBytes = await new Promise<Uint8Array | null>((resolve) => {
                const timer = setTimeout(() => { cleanup(); resolve(null) }, 35000)
                const onMsg = (e: MessageEvent) => {
                  const d = e.data as Record<string, unknown> | null
                  if (!d || typeof d !== 'object') return
                  if ('__ooNativeSaveData' in d) {
                    const b64 = String(d.__ooNativeSaveData)
                    const bin = atob(b64)
                    const u8 = new Uint8Array(bin.length)
                    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
                    cleanup(); resolve(u8)
                  } else if ('__ooNativeSaveErr' in d) {
                    try { logS('S11N-E', '原生导出失败: ' + String(d.__ooNativeSaveErr).slice(0, 120)) } catch { /* ignore */ }
                    cleanup(); resolve(null)
                  }
                }
                const cleanup = () => { clearTimeout(timer); window.removeEventListener('message', onMsg) }
                window.addEventListener('message', onMsg)
                pocWin.postMessage({ __ooNativeSaveAsk: 1 }, '*')
              })
              if (nativeBytes && nativeBytes.length > 500) {
                bytes = nativeBytes
                via = '原生导出通道（OO 引擎序列化，格式 100%）'
                savedPath = ''
                try { logS('S11N-OK', '原生导出产物=' + nativeBytes.length + 'B——保存链 v4 主通道命中') } catch { /* ignore */ }
                try { logS('S13', '原生导出产物=' + bytes.length + 'B 段落=OO序列化（格式100%）') } catch (eW1) { /* ignore */ }
              }
            }
          }
        } catch (eN) {
          try { logS('S11N-E', '原生导出通道异常: ' + String(eN).slice(0, 120)) } catch { /* ignore */ }
        }
        /* ★★ 2026-10-29 样式化重建（兜底）：callCommand 直读引擎文档模型——
           段落文字 + 字体/字号/加粗/斜体/颜色/对齐/样式名 + 表格，docx 库按格式重建。 */
        try { logS('S12a', '尝试样式化提取（callCommand 文档模型）…') } catch (eW5) { /* ignore */ }
        let styled: unknown[] | null = skipMainChannel ? styledPre : null
        try {
          /* ★ 2026-10-30 失败重试：forcesave 刚结束引擎忙，styledExtract（callCommand）
             首次常超时返回 null——实测 10:35 那轮直接掉进纯文本降级（样式全丢）。
             重试 2 次（间隔 3s）等引擎闲下来，能救回绝大多数场景。 */
          for (let attempt = 1; attempt <= 3 && !styled; attempt++) {
            if (attempt > 1) {
              /* ★ 2026-10-31 重试间隔 3s→30s：引擎文档模型初始化慢时 callCommand
                 需要几十秒才就绪（截图症状=字号变小/底纹丢失的降级保存），给足窗口 */
              try { logS('S12a-R', '提取第 ' + attempt + ' 次重试（等引擎初始化 30s）…') } catch { /* ignore */ }
              await new Promise((r) => setTimeout(r, 30000))
            }
            const se = await ooRef.current?.styledExtract?.() ?? null
            styled = se?.paras ?? null
            if (se?.meta) { try { logS('S12a-P', '探针: ' + se.meta.slice(0, 400)) } catch (eWM) { /* ignore */ } }
            if (!styled && se?.err) { try { logS('S12a-E', '样式化提取失败(第' + attempt + '次): ' + se.err) } catch (eWE) { /* ignore */ } }
          }
        } catch (eSC) { try { logS('S12a-E', '样式化提取通道异常: ' + String(eSC).slice(0, 120)) } catch (eWD) { /* ignore */ } }
        if (Array.isArray(styled) && styled.length) {
          try {
            const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, HeadingLevel, PageBreak } = await import('docx')
            /* ★ 2026-10-30 v3 行距：从已存字节（原版 docx）解包 document.xml，
               按顶层段落序号提取 w:spacing line/lineRule——引擎 GetSpacingLine 不可靠
               （探针返回 null），而版式主差异恰是原版直排行距 line=576 exact。
               引擎侧提取到的值（spLine）优先，缺失时按序号对齐兜底。 */
            let lineList: Array<{ line?: number; rule?: string; indFirstLine?: number; before?: number; after?: number; shdFill?: string }> = sharedLineList.slice()
            const pageBreakIdx: number[] = [...sharedPageBreakIdx]
            let sectionProps: { page: { margin: { top: number; right: number; bottom: number; left: number; header?: number; footer?: number } }; grid?: { linePitch: number; type?: string } } | undefined = sharedSectionProps
            let defaultSpacing: { line?: number; rule?: string; before?: number; after?: number } | undefined = sharedDefaultSpacing
            /* ★ v3 收口：版式档案（首份原件固化）优先——已存字节若是旧版无行距的
               重建产物，从它解包永远拿不到行距（死循环）。档案里有就直接用。 */
            const layoutProf = (useBaseStore.getState().tables[nodeId!]?.config as { ooLineProfile?: OoLayoutProfile } | undefined)?.ooLineProfile
            if (layoutProf?.lineList?.length) {
              lineList = layoutProf.lineList
              layoutProf.pageBreakIdx?.forEach((i) => pageBreakIdx.push(i))
              if (layoutProf.sectionProps) sectionProps = layoutProf.sectionProps
              try { logS('S12a-P2', '使用版式档案: line=' + lineList.filter((x) => x.line).length + ' 分页=' + pageBreakIdx.length + ' 页面=' + (layoutProf.sectionProps ? '√' : '默认')) } catch { /* ignore */ }
            }
            /* ★ 2026-10-30 关键修复：档案命中时**跳过**已存字节解包——否则解包的
               「旧版纯文本产物」会覆盖档案的 lineList/sectionProps（N164 实测翻车：
               分页符用 push 幸存、行距被 = 覆盖丢失） */
            try { (window as unknown as Record<string, unknown>).__ooS12Debug = { shared: sharedLineList.length, sharedLine: sharedLineList.filter((x) => x.line).length, layoutProf: !!layoutProf, layoutProfLine: layoutProf?.lineList?.length ?? -1, lineList: lineList.length, sect: !!sectionProps, sample3: JSON.stringify(lineList.slice(0, 6)).slice(0, 300) } } catch { /* ignore */ }
            if (!(layoutProf?.lineList?.length) && !sharedLineList.some((x) => x.line)) try {
              const curB64 = (useBaseStore.getState().tables[nodeId!]?.config as { ooBytes64?: string } | undefined)?.ooBytes64
              if (curB64) {
                const bin = atob(curB64)
                const storedBytes = new Uint8Array(bin.length)
                for (let ib = 0; ib < bin.length; ib++) storedBytes[ib] = bin.charCodeAt(ib)
                const JSZip = (await import('jszip')).default
                const zip = await JSZip.loadAsync(storedBytes)
                const xml = await zip.file('word/document.xml')?.async('string')
                if (xml) {
                  const bodyPart = xml.slice(xml.indexOf('<w:body>'))
                  const noTbl = bodyPart.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, '')
                  const xmlParas = noTbl.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? []
                  xmlParas.forEach((px, pi) => { if (px.includes('<w:br w:type="page"/>')) pageBreakIdx.push(pi) })
                  lineList = xmlParas.map((px) => {
                    const mSp = px.match(/<w:spacing[^>]*\/>/)
                    if (!mSp) return {}
                    const line = mSp[0].match(/w:line="(\d+)"/)?.[1]
                    const rule = mSp[0].match(/w:lineRule="(\w+)"/)?.[1]
                    return { line: line ? Number(line) : undefined, rule }
                  })
                  /* ★ v3b 页面设置：原版 sectPr 的页边距 + 文档网格（默认 1440/无网格 = 版式漂移源） */
                  const sect = xml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/)?.[0] ?? ''
                  const num = (tag: string, attr: string) => Number(sect.match(new RegExp(`<w:${tag}[^>]*w:${attr}="(\\d+)"`))?.[1] ?? 0)
                  const gridPitch = Number(sect.match(/<w:docGrid[^>]*w:linePitch="(\d+)"/)?.[1] ?? 0)
                  const gridType = sect.match(/<w:docGrid[^>]*w:type="(\w+)"/)?.[1]
                  if (num('pgMar', 'top')) {
                    sectionProps = {
                      page: { margin: { top: num('pgMar', 'top'), right: num('pgMar', 'right'), bottom: num('pgMar', 'bottom'), left: num('pgMar', 'left'), ...(num('pgMar', 'header') ? { header: num('pgMar', 'header') } : {}), ...(num('pgMar', 'footer') ? { footer: num('pgMar', 'footer') } : {}) } },
                      grid: gridPitch ? { linePitch: gridPitch, ...(gridType ? { type: gridType as 'default' | 'lines' | 'linesAndChars' } : {}) } : undefined,
                    }
                  }
                  try { logS('S12a-X', '原版对齐: 顶层段=' + lineList.length + ' 有line=' + lineList.filter((x) => x.line).length + ' 分页=' + pageBreakIdx.length + ' 页边距=' + (sectionProps ? '√' : '默认')) } catch (eX1) { /* ignore */ }
                }
              }
            } catch (eLx) { try { logS('S12a-XE', '行距对齐失败: ' + String(eLx).slice(0, 100)) } catch { /* ignore */ } }
            const ALIGN: Record<string, (typeof AlignmentType)[keyof typeof AlignmentType]> = {
              center: AlignmentType.CENTER, right: AlignmentType.RIGHT, both: AlignmentType.JUSTIFIED, justify: AlignmentType.JUSTIFIED,
            }
            const HEADING = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6] as const
            const mkRuns = (runs: Array<{ text?: string; bold?: boolean; italic?: boolean; underline?: boolean; size?: number; font?: string; color?: string }>) =>
              (runs ?? []).map((r) => new TextRun({
                text: String(r.text ?? ''),
                bold: !!r.bold,
                italics: !!r.italic,
                underline: r.underline ? {} : undefined,
                /* ★ sdkjs GetFontSize 返回的就是半磅值（21:24 实测：原 w:sz=32 → 重建 64 翻倍翻车）
                   ——docx 库 size 同为半磅，直接透传，不乘 2
                   ★ 2026-10-31 run 无显式字号时用样式表默认字号（Normal sz，如公文正文 14pt）
                   ——重建产物无样式层，docx 库默认 11pt 比原版小一大截（14:52 实测渲染高 14→10） */
                size: typeof r.size === 'number' && r.size > 0 ? Math.round(r.size) : (sharedDefaultFontSize || undefined),
                font: typeof r.font === 'string' && r.font ? r.font : undefined,
                color: typeof r.color === 'string' && /^[0-9A-Fa-f]{6}$/.test(r.color) ? r.color : undefined,
              }))
            let topParaIdx = -1 /* 顶层段落序号（跳过表格）——与 lineList 对齐 */
            const children = styled.map((raw, sIdx) => {
              const p = raw as { t?: string; align?: string; style?: string; indLeft?: number; indFirstLine?: number; spBefore?: number; spAfter?: number; spLine?: number; spLineRule?: string; runs?: Array<Record<string, unknown>>; rows?: string[][] }
              if (p.t === 'tbl' && Array.isArray(p.rows)) {
                const rows = p.rows.map((cells) => new TableRow({
                  children: cells.map((c) => new TableCell({ children: [new Paragraph(String(c ?? ''))] })),
                }))
                return new Table({ rows: rows.length ? rows : [new TableRow({ children: [new TableCell({ children: [new Paragraph('')] })] })], width: { size: 100, type: WidthType.PERCENTAGE } })
              }
              topParaIdx++
              const xmlLine = lineList[topParaIdx] ?? {}
              const align = p.align ? ALIGN[p.align] : undefined
              const hm = String(p.style ?? '').match(/^(?:标题|Heading)\s*(\d)/i)
              const heading = hm ? HEADING[Math.min(Number(hm[1]) - 1, HEADING.length - 1)] : undefined
              /* 标题样式（非数字级，如 Title）→ 居中兜底（样式带走的对齐由直排补回） */
              const isTitleStyle = /^(?:标题|Title)$/i.test(String(p.style ?? ''))
              const runs = mkRuns((p.runs ?? []) as Array<Record<string, unknown>>)
              /* ★ 首行缩进「2 字符」启发式：中文公文首行缩进恒为 2 字宽——
                 twips = 最大字号(半磅) × 20（sizeHalf/2 pt × 2 字 × 20 twips/pt）。
                 提取端单位不确定（builder 距离量纲版本差异），按内容量级钳制。 */
              let indentOpt: { firstLine?: number; left?: number } | undefined
              const maxSz = (p.runs ?? []).reduce((m, r) => Math.max(m, typeof r.size === 'number' && r.size > 0 ? r.size : 0), 0)
              /* ★ 2026-10-30 档案缩进优先：原版直排 firstLine（0=显式顶格/480=2字）最权威——
                 引擎半残时报错值（如统一 640）不再污染版式 */
              const xmlIndFL = lineList[topParaIdx]?.indFirstLine
              if (typeof xmlIndFL === 'number' && xmlIndFL !== 0) {
                const maxSz0 = (p.runs ?? []).reduce((m, r) => Math.max(m, typeof r.size === 'number' && r.size > 0 ? r.size : 0), 0)
                /* 原版 480 = 2 字宽（240/字）→ twips = maxSz(半磅)/2 * 2字 * 20 = maxSz*20*(xmlIndFL/480) */
                indentOpt = { firstLine: maxSz0 > 0 ? Math.round((maxSz0 * 20 * xmlIndFL) / 480) : Math.round(xmlIndFL) }
              } else if (xmlIndFL === 0 || xmlIndFL === undefined) {
                /* 档案无 ind 元素 = 段落继承 Normal（该体系无 firstLine）= 顶格；
                   不再放行引擎值（560/640 污染源） */
                indentOpt = { firstLine: 0 }
              } else if (typeof p.indFirstLine === 'number' && p.indFirstLine > 0) {
                /* ★ 2026-10-30 修：只对正缩进生效——undefined/负值(悬挂)不加，
                   否则无缩进段落被误加 2 字缩进 → 每行少 2 字 → 折行增多 → 版式漂移（实测 2 页 vs 3 页） */
                const fl = maxSz > 0 ? Math.round(maxSz * 20) : 640
                indentOpt = { firstLine: fl }
              }
              if (typeof p.indLeft === 'number' && p.indLeft > 0) {
                const left = p.indLeft <= 2000 ? Math.round(p.indLeft) : 0 /* 已近 twips 量级则直用 */
                if (left > 0) indentOpt = { ...(indentOpt ?? {}), left }
              }
              const spBefore0 = typeof xmlLine.before === 'number' ? xmlLine.before : (typeof p.spBefore === 'number' ? p.spBefore : undefined)
              const spBefore = typeof spBefore0 === 'number' && spBefore0 > 0 ? Math.min(Math.round(spBefore0), 600) : undefined
              const spAfter0 = typeof xmlLine.after === 'number' ? xmlLine.after : (typeof p.spAfter === 'number' ? p.spAfter : undefined)
              const spAfter = typeof spAfter0 === 'number' && spAfter0 > 0 ? Math.min(Math.round(spAfter0), 600) : undefined
              /* ★ 2026-10-30 v3 行距还原（版式主差异源）：原版直排 line=576 exact——16pt 字撑
               *  28.8pt 行格，视觉行间空隙大；不写则重建变单倍行距挤成 2 页（截图对比实锤）。
               *  GetSpacingLineRule 返回值小写化后匹配 docx 库 LineRuleType（exact/atLeast/auto）。 */
              const spLine = typeof p.spLine === 'number' && p.spLine > 0 ? Math.min(Math.round(p.spLine), 3000) : undefined
              const lrRaw = String(p.spLineRule ?? '').toLowerCase()
              const spLineRule = lrRaw.includes('exact') ? 'exact' as const : lrRaw.includes('least') ? 'atLeast' as const : lrRaw.includes('auto') ? 'auto' as const : undefined
              /* XML 对齐兜底：引擎没取到行距时用原版 XML 的（段落序号对齐） */
              const finLine = (typeof xmlLine.line === 'number' && xmlLine.line > 0) ? xmlLine.line : spLine /* ★ 档案(原版XML)优先——引擎 line/rule 语义在 auto 下错乱 */
              /* ★ 2026-10-30 rule 优先级修正：档案（原版 XML 直排值）最权威——引擎
                 GetSpacingLineRule 实测对 exact 行距返回 "auto"（错误语义，auto 下
                 line=576 会被解释为 2.4 倍行距，比原版还大），不能让它覆盖档案值。 */
              const finRule = (xmlLine.rule as 'exact' | 'atLeast' | 'auto' | undefined) ?? (spLineRule !== 'auto' ? spLineRule : undefined) ?? 'auto'
              /* ★ 2026-10-31 样式表默认段距：段落无直排 spacing 时应用档案的
                 Normal/docDefaults 行距（正文 306 auto≈24pt 与原版渲染一致） */
              /* ★ 2026-10-31 样式表默认段距：finLine 缺失时也兜底（段落有 before/after
                 但行距来自 Normal 样式 306 auto——原版渲染 24pt 行高） */
              const finLineEff = finLine != null ? finLine : (defaultSpacing?.line != null ? defaultSpacing.line : undefined)
              const spacing = spBefore != null || spAfter != null || finLineEff != null
                ? { before: spBefore, after: spAfter, line: finLineEff, lineRule: finLineEff ? (finRule || (defaultSpacing?.rule as never)) : undefined }
                : (defaultSpacing && (defaultSpacing.line != null || defaultSpacing.before != null || defaultSpacing.after != null)
                  ? { before: defaultSpacing.before, after: defaultSpacing.after, line: defaultSpacing.line, lineRule: defaultSpacing.rule as never }
                  : undefined)
              /* ★ 2026-10-31 档案底纹（粉/蓝段落高亮）应用——styled 提取协议无底纹字段，
                 从档案直排 shd 补（截图「粉色没了」的直接解） */
              const shdFill = (xmlLine as { shdFill?: string }).shdFill
              const shading = shdFill ? { type: 'clear' as const, color: 'auto', fill: shdFill } : undefined
              if (pageBreakIdx.includes(topParaIdx)) runs.push(new PageBreak())
              if (heading) return new Paragraph({ heading, alignment: align ?? (isTitleStyle ? AlignmentType.CENTER : undefined), indent: indentOpt, spacing, ...(shading ? { shading } : {}), children: runs.length ? runs : [new TextRun('')] })
              return new Paragraph({ alignment: align ?? (isTitleStyle ? AlignmentType.CENTER : undefined), indent: indentOpt, spacing, ...(shading ? { shading } : {}), children: runs.length ? runs : [new TextRun('')] })
            })
                        /* ★ 2026-10-31 页脚页码重建（公文「— X —」式）——docx 库 sections 不自动带页脚 */
            const { Footer, PageNumber, AlignmentType: AT } = await import('docx')
            const footersCfg = sharedFooter
              ? { default: new Footer({ children: [new Paragraph({ alignment: AT.CENTER, children: [
                  ...(sharedFooter.prefix ? [new TextRun(sharedFooter.prefix)] : []),
                  new TextRun({ children: [PageNumber.CURRENT] }),
                  ...(sharedFooter.suffix ? [new TextRun(sharedFooter.suffix)] : []),
                ] })] }) }
              : undefined
            const docS = new Document({ sections: [{ ...(sectionProps ? { properties: sectionProps as never } : {}), ...(footersCfg ? { footers: footersCfg } : {}), children: children.length ? children : [new Paragraph('')] }] })
            const blobS = await Packer.toBlob(docS)
            bytes = new Uint8Array(await blobS.arrayBuffer())
            via = '提取通道（样式化重建：字体/加粗/颜色/对齐/表格保留）'
            try { logS('S13a', '样式化重建产物=' + bytes.length + 'B 块=' + styled.length) } catch (eW6) { /* ignore */ }
          } catch (eSR) {
            try { logS('S12a-E', '样式化重建失败: ' + String(eSR).slice(0, 140)) } catch (eW7) { /* ignore */ }
          }
        }
        /* ★ 2026-10-29 沿用守卫（排在样式化重建**之后**——21:21 实测教训：沿用在前会把
           bytes 填上，样式化提取根本没机会跑，编辑永远存不进去）：
           样式化提取也失败时才沿用已存 ooBytes64（保样式、丢本次编辑），并明确警告。 */
        if (!bytes) {
          try {
            /* ★ 实时读 store（避开 useCallback 旧闭包——依赖只有 table?.name） */
            const curB64 = (useBaseStore.getState().tables[nodeId!]?.config as { ooBytes64?: string } | undefined)?.ooBytes64
            if (curB64) {
              const bin = atob(curB64)
              const cur = new Uint8Array(bin.length)
              for (let i = 0; i < bin.length; i++) cur[i] = bin.charCodeAt(i)
              if (cur.length) {
                bytes = cur
                via = '沿用已保存内容（样式化提取不可用，样式版本已保留）'
                try { logS('S11d', '样式化提取未出产物 → 沿用已存 ooBytes64（' + cur.length + 'B，保样式）') } catch (eL11d) { /* 忽略 */ }
                try { ooRef.current?.opLog('⚠ 样式化提取不可用——已保留上次带样式版本；本次文字改动未保存') } catch (eOLd) { /* 忽略 */ }
                pushLog('⚠ [OO保存] 样式化提取不可用，已保留上次带样式版本——本次文字改动未保存', 'warn')
              }
            }
          } catch { /* 沿用失败继续走纯文本兜底 */ }
        }
        if (!bytes) {
          try { logS('S12', '降级提取通道(docx段落，纯文本)：rows=' + payload.rows.length + ' 目标=' + payload.name) } catch (eW0) { /* ignore */ }
          try {
            /* ★ 2026-10-30 纯文本降级也补版式：文字拿不到样式，但档案里有
               行距/分页/页面设置——11:46 实测纯文本产物默认边距无行距，用户看到
               「段落对不起来」的最后一环。
               ★ 2026-10-31 再补文字样式兜底（用户「一保存啥都没了」）：从上次保存
               的产物 XML 提取每段首 run 的字体/字号/加粗，应用到降级重建的段落——
               样式化提取挂掉时字号字体不再丢失（截图症状：正文变小）。 */
            const { Document, Packer, Paragraph, TextRun, PageBreak } = await import('docx')
            /* 上次产物 XML（ooBytes64 优先）→ 每段首 run 的 rPr */
            let prevRunStyles: Array<{ font?: string; size?: number; bold?: boolean }> = []
            try {
              const cfgPrev = useBaseStore.getState().tables[nodeId!]?.config as { ooBytes64?: string } | undefined
              const prevB64 = cfgPrev?.ooBytes64
              if (prevB64) {
                const pbin = atob(prevB64)
                const pu8 = new Uint8Array(pbin.length)
                for (let ip = 0; ip < pbin.length; ip++) pu8[ip] = pbin.charCodeAt(ip)
                const PZip = (await import('jszip')).default
                const pzip = await PZip.loadAsync(pu8)
                const pxml = await pzip.file('word/document.xml')?.async('string')
                if (pxml) {
                  const pbody = pxml.slice(pxml.indexOf('<w:body>'))
                  const pNoTbl = pbody.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, '')
                  const pParas = pNoTbl.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? []
                  prevRunStyles = pParas.map((px) => {
                    const rPr = px.match(/<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0] ?? ''
                    const font = rPr.match(/w:eastAsia="([^"]+)"/)?.[1]
                    const szHalf = rPr.match(/<w:sz w:val="(\d+)"\/>/)?.[1]
                    const bold = rPr.includes('<w:b/>')
                    return { font, size: szHalf ? Number(szHalf) * 2 : undefined, bold: bold || undefined }
                  })
                }
              }
            } catch { /* 上次产物解析失败不影响降级主流程 */ }
            let pi2 = -1
            const paras = payload.rows
              .map((r) => String((Array.isArray(r) ? r[0] : r) ?? ''))
              .map((t) => {
                pi2++
                const xl = sharedLineList[pi2] ?? {}
                const spc = typeof xl.line === 'number' && xl.line > 0 ? { line: xl.line, lineRule: (String(xl.rule ?? '').toLowerCase().includes('exact') ? 'exact' as const : String(xl.rule ?? '').toLowerCase().includes('least') ? 'atLeast' as const : 'auto' as const) } : (sharedDefaultSpacing && (sharedDefaultSpacing.line != null || sharedDefaultSpacing.before != null) ? { line: sharedDefaultSpacing.line, lineRule: sharedDefaultSpacing.rule as never, before: sharedDefaultSpacing.before, after: sharedDefaultSpacing.after } : undefined)
                const prev = prevRunStyles[pi2] ?? {}
                const textRun = new TextRun({
                  text: t,
                  ...(prev.font ? { font: prev.font } : {}),
                  ...(prev.size ? { size: prev.size } : {}),
                  ...(prev.bold ? { bold: true } : {}),
                })
                const kids: Array<InstanceType<typeof TextRun> | InstanceType<typeof PageBreak>> = [textRun]
                if (sharedPageBreakIdx.includes(pi2)) kids.push(new PageBreak())
                /* ★ 2026-10-31 降级也应用档案底纹（粉/蓝段落高亮）——两条保存链格式对齐 */
                const shdFill = (xl as { shdFill?: string }).shdFill
                const shading = shdFill ? { type: 'clear' as const, color: 'auto', fill: shdFill } : undefined
                return new Paragraph(t || spc || sharedPageBreakIdx.includes(pi2) || shading ? { children: kids, spacing: spc, ...(shading ? { shading } : {}) } : {})
              })
                        const { Footer: Ftr, PageNumber: PN, AlignmentType: AT2 } = await import('docx')
            const footersCfg2 = sharedFooter
              ? { default: new Ftr({ children: [new Paragraph({ alignment: AT2.CENTER, children: [
                  ...(sharedFooter.prefix ? [new TextRun(sharedFooter.prefix)] : []),
                  new TextRun({ children: [PN.CURRENT] }),
                  ...(sharedFooter.suffix ? [new TextRun(sharedFooter.suffix)] : []),
                ] })] }) }
              : undefined
            const doc = new Document({ sections: [{ ...(sharedSectionProps ? { properties: sharedSectionProps as never } : {}), ...(footersCfg2 ? { footers: footersCfg2 } : {}), children: paras.length ? paras : [new Paragraph('')] }] })
            const blob = await Packer.toBlob(doc)
            bytes = new Uint8Array(await blob.arrayBuffer())
            via = '提取通道（docx 段落文本+档案版式）'
            try { logS('S13', 'docx 降级产物=' + bytes.length + 'B 段落=' + paras.length + '（版式自档案: line=' + sharedLineList.filter((x) => x.line).length + '）') } catch (eW1) { /* ignore */ }
          } catch (eW) {
            try { logS('S12-E', 'docx 提取通道失败: ' + String(eW).slice(0, 140)) } catch (eW2) { /* ignore */ }
          }
        }
      }
      if (!bytes || !bytes.length) {
        const why = '两条通道都没取到内容（OO 未生成产物，且提取为空）'
        try { ooRef.current?.opLog('❌ 保存失败：' + why + '。请确认单元格已提交（按 Enter 或点别的格子）再保存') } catch (eOL2) {}
        try { logS('SAVE-NODE-FAIL', '节点 ' + nodeId + ' ' + why) } catch (eL8) {}
        saveDoneRef.current = true
        return
      }
      const ext = kind === 'oosheet' ? 'xlsx' : 'docx'
      const name2 = (table?.name || 'OO 文档') + ext
      /* ★ ⤓ 导出模式：拿到的字节（优先 OO 原样导出）→ 本地下载 */
      if (payload.mode === 'export') {
        void saveDocAs(name2, bytes, ext)
        try { ooRef.current?.opLog('⤓ 已导出到本地：' + name2 + '（' + bytes.length + 'B，' + via + '）') } catch (eOL6) { /* 忽略 */ }
        saveDoneRef.current = true
        return
      }
      await persistOoBytes(bytes, via, savedPath)
      /* ★★ 2026-10-30 用户指令「把需要手点的环节全部打通」：保存成功 → 自动重开文档 ★★
       *  之前只有主通道/引用刷新有重开，样式化重建路径保存后编辑器仍显示旧内容——
       *  用户看到的永远是「没保存上」。现在统一：字节入库后立即用已存字节重开，所见=所存。 */
      if (payload.mode === 'node') {
        try {
          ooRef.current?.opLog('⟳ 保存完成——自动重开文档（读回已存字节）…')
          const ab2 = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
          await ooRef.current?.openFile(new File([ab2], name2))
          try { ooRef.current?.ackSaved() } catch { /* 清脏失败不影响 */ }
          try { logS('S16', '保存后自动重开完成: ' + bytes.length + 'B（' + via + '）') } catch { /* ignore */ }
        } catch (eR2) {
          try { ooRef.current?.opLog('⚠ 字节已入库，但自动重开失败——重新打开节点即可看到最新内容') } catch { /* ignore */ }
          pushLog('⚠ [OO保存] 已写入节点字节，自动重开失败——重新打开节点可见最新内容', 'warn')
        }
      }
    } catch (e) {
      try { ooRef.current?.opLog('❌ 保存异常：' + String(e).slice(0, 120)) } catch (eOL4) {}
      try { logS('SAVE-NODE-FAIL', '保存异常: ' + String(e).slice(0, 200)) } catch (eL5) {}
      console.error('[oo] 保存异常', e)
    } finally {
      setOoSavling(false)
    }
  }, [isOO, nodeId, kind, table?.name, onSetNodeConfig, persistOoBytes])

  const ooSaveBack = useCallback(async () => {
    if (!isOO || ooSaving) return
    setOoSavling(true)
    try {
      const fb = await ooRef.current?.saveBack()
      if (!fb || !fb.bytes || !fb.bytes.length) return
      /* ★ 2026-10-17 同假成功拦截：空产物绝不写节点（否则会把已有内容清空） */
      const d = fb.digest
      if (d && ((d.kind === 'xlsx' && !d.cells) || (d.kind === 'docx' && !d.chars))) {
        try { ooRef.current?.opLog('⚠ 忽略空产物（' + (d.text || '') + '）') } catch (eOLC) { /* 忽略 */ }
        return
      }
      const bytes = fb.bytes
      const ext = kind === 'oosheet' ? 'xlsx' : 'docx'
      const name = (table?.name || 'OO 文档') + ext
      const meta = await getTemplateStore().save(name, bytes)
      // ★ 2026-10-11 用户指令「保存在画布节点文件，随画布文件保存」：字节 base64 进节点配置，
      //   真正随画布项目文件走（IndexedDB 模板库仅作副本，换机/清缓存不丢）
      onSetNodeConfig(nodeId, { ooTemplateId: meta.id, ooTemplateName: name, ooBytes64: u8ToB64(bytes) })
      // ★ 2026-10-11 用户报「没打开保存的」：保存后作废该节点的冻结实例——
      //   否则重开节点时 boot 前两步会接管保存前的旧冻结现场，绕过 ooBytes64 加载分支
      try { discardFrozen(nodeId) } catch { /* 无冻结实例则忽略 */ }
    } catch (e) {
      console.error('[oo] 保存失败', e)
    } finally {
      setOoSavling(false)
    }
  }, [isOO, ooSaving, kind, table?.name, nodeId, onSetNodeConfig])

  /** OO ⤓：导出当前编辑器字节 */
  /* ★ 2026-10-11 保存链重构：⤓ 也走新链（引擎直读→服务端构建→本地下载），不再经坏 DS 链 */
  const ooExport = useCallback(async () => {
    ooRef.current?.requestSave('export')
  }, [])
  /** 编辑器主题（暗黑/明媚），localStorage 记住选择；同步到 <html> 供 portal 元素（引用栏）取 CSS 变量 */
  const [theme, setTheme] = useState<'dark' | 'light'>(() => (prefGet('editorTheme') as 'dark' | 'light') || 'dark')
  useEffect(() => {
    document.documentElement.dataset.theme = theme
  }, [theme])
  const toggleTheme = () => {
    setTheme((t) => {
      const next = t === 'dark' ? 'light' : 'dark'
      prefSet('editorTheme', next)
      return next
    })
  }

  useEffect(() => {
    getTemplateStore()
      .list()
      .then(setLibs)
      .catch(() => setLibs([]))
  }, [bookVersion])

  // 最大化：class 是唯一真相（头部方形按钮 / 面板 Ribbon 右上角都只调 setHeadMax）
  useEffect(() => {
    const el = document.querySelector('.node-editor')
    el?.classList.toggle('ne-max', headMax)
    el?.parentElement?.classList.toggle('ne-max-pad', headMax)
    return () => {
      el?.classList.remove('ne-max')
      el?.parentElement?.classList.remove('ne-max-pad')
    }
  }, [headMax, nodeId])

  useEffect(() => {
    setDraftName(tables[nodeId]?.name ?? '')
  }, [nodeId, tables])

  // 刻意不提供「Esc 关闭」与「点遮罩关闭」：编辑工作簿时 Esc 要留给公式/弹层自己用，
  // 点边缘误触直接丢编辑现场代价太大 —— 关闭只走右上角 ×。

  const computed = useMemo(
    () => (table ? computeNode(table.id, tables, records) : { fields: [], rows: [] }),
    [table, tables, records],
  )

  const templateId = table?.config?.xlsxTemplateId
  const templateName = table?.config?.xlsxTemplateName ?? '工作簿.xlsx'

  // 载入工作簿字节（离线查看/编辑走这条路）
  useEffect(() => {
    let dead = false
    if (!isWorkbookKind(kind) || !templateId) {
      setBookBytes(null)
      return
    }
    ;(async () => {
      try {
        setBusy('读取工作簿…')
        // ★ 批次17 两处修复（实测断点：打开节点只见「模板已不在模板库中」，磁盘文件其实都在）：
        //   ① 先把工作目录里的模板补进模板库（幂等，按 name 去重，绝不覆盖已有条目）
        //   ② 再读工作簿 —— 种子在前，磁盘上已有的同名模板能被自动救回，不必回画布重建节点
        //   ③ 无论成败都刷新一次 libs，让标题栏模板下拉有内容可重选（原先只在用户导入后刷新）
        try {
          await seedTemplatesFromDisk()
        } catch {
          /* 磁盘模板不可达（纯静态托管等）→ 静默降级，走原有报错路径 */
        }
        const list = await getTemplateStore().list().catch(() => null)
        if (!dead && list) setLibs(list)
        const b = await getTemplateStore().read(templateId)
        if (dead) return
        setBookBytes(b)
        setBookErr(b ? '' : '模板已不在模板库中')
      } catch (e) {
        if (!dead) setBookErr(e instanceof Error ? e.message : String(e))
      } finally {
        if (!dead) setBusy('')
      }
    })()
    return () => {
      dead = true
    }
  }, [kind, templateId, bookVersion])

  /** 工作簿里的改动 → 写回模板库（原地更新同一个 id）→ 让绑定面板重新发现域 */
  const saveBookToStore = useCallback(
    async (next: Uint8Array) => {
      if (!templateId) return
      setBookBytes(next)
      try {
        await getTemplateStore().save(templateName, next)
        setBookVersion((v) => v + 1)
      } catch (e) {
        setBookErr('写回模板库失败：' + (e instanceof Error ? e.message : String(e)))
      }
    },
    [templateId, templateName],
  )

  /** 标题栏「打开文件」：选一个 .xlsx 存入模板库并绑定（工作簿字节由载入 effect 自动重读） */
  const openXlsxFile = async () => {
    try {
      const picked = await pickTemplateFile('xlsx')
      if (!picked) return
      setBusy('保存模板…')
      const meta = await getTemplateStore().save(picked.name, picked.data)
      setLibs(await getTemplateStore().list().catch(() => libs))
      onSetNodeConfig(nodeId, { xlsxTemplateId: meta.id, xlsxTemplateName: meta.name })
      setBusy('')
    } catch (e) {
      setBookErr('导入模板失败：' + (e instanceof Error ? e.message : String(e)))
      setBusy('')
    }
  }

  /** 空态「创建空白文档」：@office-kit/xlsx 造一个只有 Sheet1 的空工作簿入库（2026-10-22 去 SheetJS/CVE-SDL-2024-0002 系） */
  const createBlankXlsx = async () => {
    try {
      setBusy('创建空白工作簿…')
      const wb = createWorkbook()
      addWorksheet(wb, 'Sheet1')
      const data = await workbookToBytes(wb)
      const meta = await getTemplateStore().save('未命名.xlsx', data)
      setLibs(await getTemplateStore().list().catch(() => libs))
      onSetNodeConfig(nodeId, { xlsxTemplateId: meta.id, xlsxTemplateName: meta.name })
      setBusy('')
    } catch (e) {
      setBookErr('创建失败：' + (e instanceof Error ? e.message : String(e)))
      setBusy('')
    }
  }

  /** 「导入文件名」设置：只改节点上的显示名（导出文件名跟着它走） */
  const renameTpl = () => {
    const n = prompt('导入文件名：', templateName)
    if (n && n.trim()) onSetNodeConfig(nodeId, { xlsxTemplateName: n.trim() })
  }

  /* ── Word 节点标题栏文件操作（与 Excel 的 ne-fileops 同位同款）── */
  const docxTemplateId = table?.config?.docxTemplateId
  const docxTemplateName = table?.config?.docxTemplateName ?? '文档.docx'

  /** Word 模板历史下拉：换绑历史文档 */
  const pickDocxLib = (id: string) => {
    const hit = libs.find((l) => l.id === id)
    if (hit) onSetNodeConfig(nodeId, { docxTemplateId: hit.id, docxTemplateName: hit.name })
  }
  /** Word 模板改名（只改节点显示名） */
  const renameDocxTpl = () => {
    const n = prompt('导入文件名：', docxTemplateName)
    if (n && n.trim()) onSetNodeConfig(nodeId, { docxTemplateName: n.trim() })
  }
  /** 标题栏 📂：选 .docx → 入库绑定（与面板「导入模板」共享同一份逻辑）。
   *  无标记也直接打开（2026-09-27 用户拍板：不弹提示不拦截）——原文档可预览/就地编辑/导出；
   *  要连数据时再加 ｛｛field:x｝｝ 标记重新导入。 */
  const openDocxFile = async () => {
    try {
      const picked = await pickTemplateFile()
      if (!picked) return
      setBusy('正在打开…')
      const { meta } = await importDocxTemplate(picked)
      setLibs(await getTemplateStore().list().catch(() => libs))
      onSetNodeConfig(nodeId, { docxTemplateId: meta.id, docxTemplateName: meta.name })
      setBusy('')
    } catch (e) {
      window.alert('导入模板失败：' + (e instanceof Error ? e.message : String(e)))
      setBusy('')
    }
  }

  if (!table || !kind) return null

  const commitName = () => {
    const n = draftName.trim()
    if (n && n !== table.name) onRename(table.id, n)
    else setDraftName(table.name)
  }

  // 备忘 · 页签换方形图标按钮（节省标题栏横向空间；完整名称在 title 悬停提示，data-tab 供回归脚本定位）
  const tabs: { id: Tab; label: string; icon: string }[] = [{ id: 'main', label: '绑定与产出', icon: '⇄' }]
  if (isWorkbookKind(kind)) tabs.push({ id: 'book', label: '打开工作簿', icon: '▦' })
  tabs.push({ id: 'data', label: '输入数据', icon: '⌨' })

  return createPortal(
    // 遮罩不做「点边缘关闭」—— 误触代价大于便利，关闭只认 × 按钮
    <div className="node-editor-mask">
      {ooCloseAsk && (
        <div style={{ position: 'absolute', inset: 0, zIndex: 60, display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'rgba(0,0,0,0.5)' }}>
          <div style={{ background: 'var(--ed-bg1, #1a1d22)', border: '1px solid var(--ed-border, #333)', borderRadius: 10, padding: '18px 20px', minWidth: 300, boxShadow: '0 12px 40px rgba(0,0,0,.6)' }}>
            <div style={{ fontWeight: 'bold', marginBottom: 10 }}>文档有未保存的修改</div>
            <div style={{ color: '#999', fontSize: 12, marginBottom: 14 }}>关闭前要保存到节点吗？</div>
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
              <button className="bar-btn" disabled={ooCloseBusy} onClick={() => setOoCloseAsk(false)}>取消</button>
              <button className="bar-btn" disabled={ooCloseBusy} onClick={() => { setOoCloseAsk(false); onClose() }}>不保存</button>
              <button
                className="bar-btn primary"
                disabled={ooCloseBusy}
                onClick={async () => {
                  setOoCloseBusy(true)
                  saveDoneRef.current = false
                  valuesArrivedRef.current = false
                  ooRef.current?.requestSave('node')
                  const t0 = Date.now()
                  while (!valuesArrivedRef.current && Date.now() - t0 < 3000) { await new Promise((r) => setTimeout(r, 200)) }
                  setOoCloseAsk(false)
                  onClose()
                }}
              >{ooCloseBusy ? '保存中…' : '保存并关闭'}</button>
            </div>
          </div>
        </div>
      )}
      <div className="node-editor" role="dialog" aria-label={`${KIND_TITLE[kind] ?? kind} 编辑器`} data-theme={theme}>
        <header className="ne-head">
          <span className={`rf-badge rf-badge--${kind}`}>{BADGE_TEXT[kind] ?? kind}</span>
          <input
            className="ne-name"
            value={draftName}
            onChange={(e) => setDraftName(e.target.value)}
            onBlur={commitName}
            onKeyDown={(e) => {
              if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
            }}
            title="节点名称（回车确认）"
          />
          <span className="ne-kind">{KIND_TITLE[kind] ?? kind}</span>
        {isOO ? (
          <span className="ne-count" title="当前打开的文件">
            {ooOpenedName || '未打开文件'}
          </span>
        ) : (
          <span className="ne-count">
            {computed.fields.length} 列 · {computed.rows.length} 行
          </span>
        )}
        {/* ★ 2026-10-17 未保存提示：有改动就常驻显示，点它等于点保存 */}
        {isOO && ooDirty && (
          <span
            className="ne-dirty"
            role="button"
            tabIndex={0}
            title="有未保存的修改 · 点击保存"
            style={{ cursor: 'pointer', color: '#e0a04a', fontSize: 12, whiteSpace: 'nowrap', display: 'inline-flex', alignItems: 'center', gap: 4 }}
            onClick={() => { const sid = newSaveSid(); logS('S1', '点击 ● 未保存 触发保存 节点=' + nodeId + ' sid=' + sid); ooRef.current?.requestSave('node', sid) }}
            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); const sid = newSaveSid(); logS('S1', '键盘触发未保存提示保存 节点=' + nodeId + ' sid=' + sid); ooRef.current?.requestSave('node', sid) } }}
          >
            <span style={{ width: 6, height: 6, borderRadius: 3, background: '#e0a04a', display: 'inline-block' }} />
            未保存
          </span>
        )}

        {/* Excel 节点：文件级操作收进标题栏（历史下拉 + 打开文件 + 保存 + 生成 Excel + 改名） */}
        {isWorkbookKind(kind) && (
          <div className="ne-fileops">
            <select
              className="ne-libsel"
              value={templateId ?? ''}
              title="模板历史（默认显示当前导入文件名）"
              onChange={(e) => {
                const hit = libs.find((l) => l.id === e.target.value)
                if (hit) onSetNodeConfig(nodeId, { xlsxTemplateId: hit.id, xlsxTemplateName: hit.name })
              }}
            >
              <option value="">{templateId ? '（未命名）' : '未选择模板'}</option>
              {libs.map((l) => (
                <option key={l.id} value={l.id}>
                  {l.id === templateId ? table?.config?.xlsxTemplateName || l.name : l.name}
                </option>
              ))}
            </select>
            <button className="ne-iconbtn" title="设置导入文件名" onClick={renameTpl}>
              ✎
            </button>
            <button
              className="ne-iconbtn"
              title="打开一个 .xlsx 作为本节点的工作簿"
              onClick={() => void openXlsxFile()}
            >
              📂
            </button>
            <button
              className="ne-iconbtn"
              disabled={!bookBytes}
              title="把当前工作簿写回模板库"
              onClick={() => bookBytes && void saveBookToStore(bookBytes)}
            >
              💾
            </button>
            <button
              className="ne-iconbtn"
              disabled={!bookBytes}
              title="生成 Excel 文件（另存到本地）"
              onClick={() => bookBytes && void saveDocAs(templateName, bookBytes, 'xlsx')}
            >
              ⤓
            </button>
            <button
              className="ne-iconbtn"
              disabled={!bookBytes}
              title="保存副本（另存一份「文件名-副本」到本地，不影响当前模板）"
              onClick={() =>
                bookBytes &&
                void saveDocAs(
                  templateName.replace(/(\.xlsx)?$/i, '') + '-副本.xlsx',
                  bookBytes,
                  'xlsx',
                )
              }
            >
              ⧉
            </button>
          </div>
        )}

        {/* Word 节点：文件级操作收进标题栏（历史下拉 + 改名 + 打开模板）—— 与 Excel 的 ne-fileops 同位同款；
            生成/导出仍在 Ribbon 与面板底栏（职责不重复落地） */}
        {kind === 'docx' && (
          <div className="ne-fileops">
            <select
              className="ne-libsel"
              value={docxTemplateId ?? ''}
              title="模板历史（默认显示当前导入文件名）"
              onChange={(e) => pickDocxLib(e.target.value)}
            >
              <option value="">{docxTemplateId ? '（未命名）' : '未选择模板'}</option>
              {libs.map((l) => (
                <option key={l.id} value={l.id}>
                  {l.id === docxTemplateId ? table?.config?.docxTemplateName || l.name : l.name}
                </option>
              ))}
            </select>
            <button className="ne-iconbtn" title="设置导入文件名" onClick={renameDocxTpl}>
              ✎
            </button>
            <button
              className="ne-iconbtn"
              title="打开一个 .docx 作为本节点的模板（自动把 ｛｛field:x｝｝ 标记转成内容控件）"
              onClick={() => void openDocxFile()}
            >
              📂
            </button>
          </div>
        )}

        {/* OO 节点：文件级操作收进标题栏（📂 打开 + 💾 保存回写 + ⤓ 导出）—— 与 Excel/Word 的 ne-fileops 同位同款 */}
        {isOO && (
        <div className="ne-fileops">
          {/* ★ 打开历史（2026-10-19 用户指令：打开文件左侧加历史按钮，下拉最近打开的文件） */}
          {isOO && (
            <span style={{ position: 'relative', display: 'inline-flex' }} onClick={(e) => e.stopPropagation()}>
              <button
                ref={ooHistBtnRef}
                className={`ne-iconbtn${ooHistOpen ? ' is-on' : ''}${ooHistItems.length ? '' : ' is-empty'}`}
                title="打开历史记录：最近导入的文件（点击条目直接重开）"
                onClick={() => {
                  const el = ooHistBtnRef.current
                  if (el) {
                    const r = el.getBoundingClientRect()
                    setOoHistPos({ top: Math.round(r.bottom + 6), left: Math.max(8, Math.round(r.left)) })
                  }
                  setOoHistOpen((v) => !v)
                }}
              >
                <IcHistory size={19} />
              </button>
              {ooHistOpen && (
                <div
                  className="ne-refupd-pop"
                  style={{ top: ooHistPos.top, left: ooHistPos.left, right: 'auto', width: 260 }}
                  onClick={(e) => e.stopPropagation()}
                >
                  <div style={{ fontWeight: 600, marginBottom: 6 }}>打开历史</div>
                  {ooHistItems.length === 0 && (
                    <div style={{ color: 'var(--ed-text3)', fontSize: 12, padding: '4px 0 2px' }}>
                      暂无记录 —— 用左侧「打开文件」导入后出现在这里
                    </div>
                  )}
                  {ooHistItems.map((h) => (
                    <div
                      key={h.name + h.ts}
                      onClick={() => openHistItem(h)}
                      style={{
                        display: 'flex', justifyContent: 'space-between', gap: 8, alignItems: 'center',
                        padding: '6px 8px', borderRadius: 6, cursor: 'pointer', fontSize: 12,
                      }}
                      onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.background = 'var(--ed-bg3)' }}
                      onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.background = 'transparent' }}
                      title={h.name}
                    >
                      <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{h.name}</span>
                      <span style={{ color: 'var(--ed-text3)', flexShrink: 0, fontSize: 11 }}>
                        {new Date(h.ts).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                      </span>
                    </div>
                  ))}
                  <div className="ne-refupd-ft">条目随节点保存（最多 3 条）；点击即重开该文件。</div>
                </div>
              )}
            </span>
          )}
          <input
            ref={ooFileRef}
              type="file"
              accept='.docx,.xlsx'
              style={{ display: 'none' }}
              onChange={(e) => {
                const f = e.target.files && e.target.files[0]
                if (f) {
                  // ★ 按扩展名自动切换 OO 节点类型（docx→ooword / xlsx→oosheet；徽章/面板/样式全跟随）
                  const ext = f.name.split('.').pop()?.toLowerCase()
                  const target = ext === 'xlsx' ? 'oosheet' : ext === 'docx' ? 'ooword' : null
                  if (target && target !== kind && onSwitchKind) onSwitchKind(nodeId, target)
                  /* ★ 2026-10-19 用户反馈「新导入的文件也提示保存 即使没修改」：
                   *   导入 = 已入库——字节立即写进节点配置（ooBytes64），不存在
                   *   「关了就丢」的问题 ⇒ 打开后不该再有保存提示。 */
                  void f.arrayBuffer().then((buf) => {
                    const u8 = new Uint8Array(buf)
                    void persistOoBytes(u8, '导入打开 ' + f.name)
                    recordOoHistory(f.name, u8)
                  })
                  void ooRef.current?.openFile(f, ext === 'xlsx' || ext === 'docx' ? ext : undefined)
                }
                e.target.value = ''
              }}
            />
            <button
              className="ne-iconbtn"
              title='打开 docx / xlsx（自动切换 OO Word / OO Excel）'
              onClick={() => ooFileRef.current?.click()}
            >
              {/* ★ 2026-10-19：📂 emoji 是彩色字形（黄棕），深色排里突兀——换描边文件夹
                  （currentColor，与相邻 SVG 同语言，随主题取色）。 */}
              <IcFolderOpen size={19} />
            </button>
            {/* ★ 2026-10-11 用户指令：样例选择器（📚/▾）移除——空节点由服务端自动种默认样例；
                原位置由 💾 保存承担（保存回写节点文件，随画布保存） */}
            <button
              className={`ne-iconbtn${ooDirty ? ' is-dirty' : ''}`}
              disabled={ooSaving}
              title={ooDirty ? '有未保存的修改 —— 点击保存到本节点（随画布文件保存，重开自动恢复）' : '保存：把编辑器当前内容保存回本节点（随画布文件保存，重开自动恢复）'}
              onClick={() => {
                /* ★ 2026-10-17 步骤埋点 S1：整条保存链的起点，同时生成会话 ID。
                 *   没有 S1 就说明按钮没被点到（或事件没绑上），后面所有步骤缺失都不用再查。
                 *   sid 随 requestSave 下发到 /poc，S2…S14 全线共用，日志可按会话精确切开。 */
                const sid = newSaveSid()
                try { ooRef.current?.opLog('💾 已点击保存：向编辑器请求当前内容…') } catch (eOL5) { /* 忽略 */ }
                logS('S1', '点击 💾 节点=' + nodeId + ' 文档=' + (table?.name || '-') + ' sid=' + sid + ' 通道起点')
                ooRef.current?.requestSave('node', sid)
              }}
            >
              {/* 💾 是 emoji，走 emoji 字体、渲染宽度与旁边线性 SVG 不一致 —— 同排参差。换成描边盘。 */}
              <IcDatabase size={19} />
              {/* 有未保存修改时左上角红点（与标题栏「未保存」同一信号源） */}
              {ooDirty && (
                <span style={{ position: 'absolute', top: 2, right: 2, width: 6, height: 6, borderRadius: 3, background: '#e0a04a' }} />
              )}
            </button>
            <button className="ne-iconbtn" title="导出当前编辑器内容到本地" onClick={() => void ooExport()}>
              {/* ⤓ 同样换 SVG：箭头 + 底线，几何居中 */}
              <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                <path d="M8 2.4v7.4" />
                <path d="M5.2 7.1L8 9.9l2.8-2.8" />
                <path d="M2.8 12.4h10.4" />
              </svg>
            </button>
            <button
              className={`ne-iconbtn${ooRefOpen ? ' is-on' : ''}`}
              title={ooRefOpen ? '折叠节点数据引用（左栏）' : '显示节点数据引用（左栏）'}
              onClick={() => setOoRefOpen((v) => !v)}
            >
              {/* 左栏图标：方框 + 左侧实心竖条。★ 用 SVG 而非 ◧/▣——
                  这类符号字形不对称、内置留白偏一边，同排里看着没居中（2026-10-04 用户报）。
                  几何自算：外框 1..15（14px），竖条 x=3..7，右侧留白 8..15，视觉重心居中。 */}
              <IcPanelLeft size={19} />
            </button>
            <button
              className={`ne-iconbtn${ooRightOpen ? ' is-on' : ''}`}
              title={ooRightOpen ? '折叠数据预选（右栏）' : '显示数据预选（右栏）'}
              onClick={() => setOoRightOpen((v) => !v)}
            >
              {/* 右栏图标：与左栏镜像 —— 竖条在右侧，几何与左边完全对称。 */}
              <IcPanelRight size={19} />
            </button>
            <button
              className={`ne-iconbtn${ooDiagOpen ? ' is-on' : ''}`}
              title={ooDiagOpen ? '收起诊断信息' : '显示诊断信息'}
              onClick={() => setOoDiagOpen((v) => !v)}
            >
              {/* ⓘ 换成 SVG：ⓘ 字形带外圈且重心偏左上，与旁边线性图标不齐 */}
              <IcInfo size={19} />
            </button>
            {kind === 'oosheet' && (
              <button
                className={`ne-iconbtn${refMarkOn === '1' ? ' is-on' : ''}`}
                title={refMarkOn === '1'
                  ? '引用格着色：开 —— 点击关闭（即时生效，不改文档）'
                  : '引用格着色：关 —— 点击开启（即时生效，不改文档）'}
                onClick={toggleRefMark}
              >
                {/* 引用底色开关：描边圆角框常驻；开=中心小方块（与「高亮行列」同语言），
                    关=斜杠。2026-10-19 用户「重新设计」——弃整块填充（深色下一坨重色）。 */}
                <IcMark size={19} off={refMarkOn !== '1'} />
              </button>
            )}
            {kind === 'oosheet' && (
              <button
                className={`ne-iconbtn${rowColHi === '1' ? ' is-on' : ''}`}
                title={rowColHi === '1'
                  ? '高亮行列：开 —— 选中格所在行/列加底色；点击关闭'
                  : '高亮行列：关 —— 点击开启（选中格所在行/列加底色，看长表不串行）'}
                onClick={toggleRowColHi}
              >
                {/* 十字高亮图标：3×3 网格，中行中列填充 —— 「行列同时亮」的几何直译。 */}
                <IcGrid size={19} />
              </button>
            )}
            {/* ★ 外部数据引用更新策略（2026-10-18 用户指令「引用更新设置 放在标题栏 方形按钮上」）：
                原入口在 /poc 工具栏顶部 —— 随 tab 切换会丢、藏在功能区里不好找。
                挪到标题栏与 💾/ⓘ/着色开关同一排；auto=主色高亮，off=灰。 */}
            {kind === 'oosheet' && (
              <span
                className="ne-refupd-wrap"
                onClick={(e) => e.stopPropagation()}
              >
                <button
                  ref={refUpdBtnRef}
                  className={`ne-iconbtn${refUpdMode === 'auto' ? ' is-on' : ''}${refUpdMode === 'off' ? ' is-refupd-off' : ''}`}
                  title={refUpdBtnTitle}
                  onClick={() => {
                    /* ★ 弹层用 fixed 定位：按按钮当下 rect 算 top/right（.ne-head overflow:hidden
                     *   会裁 absolute 子层，fixed 逃出裁剪；mask 全屏不构成可见裁剪） */
                    const el = refUpdBtnRef.current
                    if (el) {
                      const r = el.getBoundingClientRect()
                      setRefUpdPopPos({ top: Math.round(r.bottom + 6), right: Math.max(8, Math.round(window.innerWidth - r.right)) })
                    }
                    setRefUpdPopOpen((v) => !v)
                  }}
                >
                  {/* 环形双箭头（刷新）+ 中心格：表达「外部数据 → 表格」的同步语义 */}
                  <IcRefresh size={19} />
                </button>
                {refUpdPopOpen && (
                  <div className="ne-refupd-pop" style={{ top: refUpdPopPos.top, right: refUpdPopPos.right }}>
                    <div className="ne-refupd-hd">外部数据引用更新</div>
                    <div className="ne-refupd-info">
                      {refInfo
                        ? '本次打开：' + refInfo.total + ' 处外部关联' + (refInfo.changed ? '，' + refInfo.changed + ' 处数据有更新' : '，数据已是最新')
                        : '本次打开：未检测到外部关联（或文档尚未打开）'}
                    </div>
                    {(['off', 'ask', 'auto'] as RefUpdMode[]).map((m) => (
                      <button
                        key={m}
                        className={`ne-refupd-row${refUpdMode === m ? ' is-cur' : ''}`}
                        onClick={() => {
                          applyRefUpdMode(m)
                          flashToast('外部数据更新：' + REF_UPD_LABEL[m] + '（已保存）')
                        }}
                      >
                        <span className="ne-refupd-dot" />
                        <span className="ne-refupd-name">{REF_UPD_LABEL[m]}</span>
                        <span className="ne-refupd-desc">{REF_UPD_DESC[m]}</span>
                      </button>
                    ))}
                    <div className="ne-refupd-ft">设置立即保存；打开文档与上游改数时按此执行。</div>
                  </div>
                )}
              </span>
            )}
          </div>
        )}

        {/* ★ 2026-10-19 用户指令「这两个按钮去掉」：OO 节点不需要页签（⇄绑定与产出 /
            ⌨输入数据）——编辑器全屏化，标题栏少一排干扰。表/参数表等节点保留。 */}
        {!isOO && (
          <div className="ne-tabs">
            {tabs.map((t) => (
              <button
                key={t.id}
                className={`tab${tab === t.id ? ' tab-active' : ''}`}
                data-tab={t.id}
                title={t.label}
                onClick={() => setTab(t.id)}
              >
                {t.icon}
              </button>
            ))}
          </div>
        )}
        {/* ★ 2026-10-19 用户「按钮排列向右靠齐 间距统一」：主题/最大化/关闭收进同一组
            （组内 4px 与文件操作组一致），OO 页签隐藏后整排按钮作为一组贴右。 */}
        {(isOO || isWorkbookKind(kind) || kind === 'docx') && (
          <div className="ne-headtail">
            <button
              className="ne-iconbtn"
              title={theme === 'dark' ? '切换到明媚模式（白底）' : '切换到暗黑模式'}
              onClick={toggleTheme}
            >
              {/* ★ 2026-10-19 重设计：☀/🌙 emoji 是彩色字形（橙太阳），换线性 SVG 随主题取色 */}
              {theme === 'dark' ? <IcSun size={19} /> : <IcMoon size={19} />}
            </button>
            {(isOO || isWorkbookKind(kind) || kind === 'docx') && (
              <button
                className="ne-maxbtn"
                title={headMax ? '还原' : '最大化'}
                onClick={() => setHeadMax((m) => !m)}
              >
                {headMax ? <IcMinimize size={17} /> : <IcMaximize size={17} />}
              </button>
            )}
            <button className="ne-close" onClick={isOO ? handleOoClose : onClose} title="关闭">
              ×
            </button>
          </div>
        )}
      </header>

        {/* Word 的功能区（wrp）与 Excel 的（wb）都直接贴住标题栏；OO 节点 iframe 四周全贴边
            （最大化时参数引用/数据预选以 flex 子元素嵌入左右缘，不再用 padding 补位） */}
        <div
          className={`ne-body${tab === 'book' || kind === 'docx' ? ' ne-body--flush' : ''}${isOO ? ' ne-body--flush ne-body--oo' : ''}`}
        >
          {tab === 'main' &&
            (isOO ? (
              <div style={{ display: 'flex', width: '100%', height: '100%', minHeight: 0 }}>
                <div style={{ flex: 1, minWidth: 0, position: 'relative' }}>
                  <OnlyOfficeNodeEditor
                    ref={ooRef}
                    anyDragging={anyDragging}
                    kind={kind}
                    nodeId={nodeId}
                    /* ★ 2026-10-26 统一锁：通用 dataLocked || 遗留 ooDataLocked */
                    dataLocked={!!((table?.config as { dataLocked?: boolean; ooDataLocked?: boolean } | undefined)?.dataLocked || (table?.config as { ooDataLocked?: boolean } | undefined)?.ooDataLocked)}
                    bytes={null}
                    name={table?.name || 'OO 文档'}
                    ooTemplateId={table?.config?.ooTemplateId}
                    ooBytes64={(table?.config as { ooBytes64?: string } | undefined)?.ooBytes64}
                    ooBytesDbId={(table?.config as { ooBytesDbId?: string } | undefined)?.ooBytesDbId}
                    /* ★ 打开文档前把「域数据 + 定义名称 + 引用格缓存」按现值补进去 */
                    injectRefs={injectRefDomains}
                    diagOpen={ooDiagOpen}
                    theme={theme}
                    maximized={headMax}
                    onOpenedName={setOoOpenedName}
                    ooRefItems={ooRefItems}
                    refOpen={ooRefOpen}
                    onToggleRef={() => { try { ooRef.current?.opLog?.('[面板] ' + (ooRefOpen ? '关闭' : '打开') + '参数引用') } catch { /* ignore */ } setOoRefOpen((v) => !v) }}
                    // ★ 域外写回接线（2026-10-06）：/poc 的「刷新域值」按钮 → 宿主执行写回
                    // ★ 2026-10-17：改为「注入域数据 + 引用格缓存」并把结果落回节点字节（原先是写模板库副本，节点永远读不到）
                    onRefreshRef={() => {
                      /* Word 侧（内容控件）仍走旧的模板写回管线；Excel 走新的域注入 */
                      if (kind === 'ooword') { void ooRefresh.refresh(); return }
                      void refreshRefDomains('manual')
                    }}
                    refreshBusy={ooRefresh.busy || refSyncing}
                    rowColHi={rowColHi}
                    /* ★ 2026-10-18 外部引用更新策略：模式 / 待确认提问 / 检测结果 / 回传通道 */
                    ooRefUpdMode={refUpdMode}
                    ooRefAsk={refAsk}
                    ooRefInfo={refInfo}
                    onSetRefUpdMode={setRefUpdModeFromEditor}
                    onApplyRefAsk={() => { void applyRefAsk() }}
                    refMarks={refMarksData ? { on: refMarkOn, marks: refMarksData.marks, geom: refMarksData.geom } : undefined}
                    onPickRef={(srcId, fieldId, fieldName, value, mode) => {
                      // ★ 用户操作可见化（2026-10-10 用户指令）：点击参数进诊断抽屉
                      try { ooRef.current?.opLog?.('[操作] 点击参数（内嵌面板·最大化）: ' + srcId + '.' + (fieldName || fieldId)) } catch { /* ignore */ }
                      // ★ 点选即形成关联（2026-10-06 用户：「我先不做线，
                      //   我做成点击后能引用参数，然后形成关联」）。
                      //   原来这里只写 ooRefs（域清单）⇒ 写了没人消费、也不形成边。
                      //   现在经 ooRefPatchOf 一次算出 ooRefs + linkedInputs：
                      //     · ooRefs       = 域清单（带标准 URN）
                      //     · linkedInputs = 数据流上游 ⇒ edgesOf 能派生出边
                      //        （edgesOf 是画布/工程格式的单一真相源，
                      //          这里写了 ⇒ 连线、工程文件、触发引擎三处同时通）
                      //   ★ 必须合并成一次 setNodeConfig：分两次写会在中途失败时
                      //     留下半截状态，且撤销栈多出一个无意义快照。
                      const patch = ooRefPatchOf(table?.config, srcId, fieldId)
                      onSetNodeConfig(nodeId, patch)
                      try { ooRef.current?.opLog?.('[关联] 已落库 ooRefs+linkedInputs（画布边同步）') } catch { /* ignore */ }
                      // ★ 最大化路径也要把引用公式写进编辑器（2026-10-09 用户实测
                      //   「点参数引用没任何反应」的另一半根因：此前只记关联不写入，
                      //   与普通模式浮卡的 onPick 行为对齐）
                      try {
                        const src = (ooRefItems ?? []).find((x) => x.id === srcId)
                        const f = src?.fields.find((x) => x.id === fieldId)
                        ooRef.current?.insertRef?.({
                          nodeId: srcId,
                          nodeName: src?.name,
                          fieldId,
                          fieldName: fieldName || fieldId,
                          value: value != null && value !== '' ? value : f?.value,
                          mode: mode || 'value',
                        })
                      } catch { /* 编辑器未就绪则跳过 */ }
                    }}
                    onSaved={(bytes) => {
                      // 保存回写由 ooSaveBack 统一处理（入库绑定 ooTemplateId）
                      try { localStorage.setItem('oo-last-saved-' + table?.id, String(bytes.length)) } catch {}
                    }}
                    /* ★ 2026-10-11 保存链重构：/poc 引擎直读值数组回传 → 服务端构建 xlsx → 入节点配置 */
                    saving={ooSaving} onSaveValues={(payload) => { valuesArrivedRef.current = true; void ooSaveFromValues(payload) }}
                    onDirtyChange={applyOoDirty}
                    onClose={onClose}
                  />
                  {/* 浮层提示：着色开关/域刷新的结果（面板收起时也能看见） */}
                  {neToast && <div className="ne-toast">{neToast}</div>}
                </div>
              </div>
            ) : kind === 'docx' ? (
              <WordReportPanel
                table={table}
                tables={tables}
                records={records}
                onSetNodeConfig={onSetNodeConfig}
                editorMode
                maximized={headMax}
                onToggleMax={() => setHeadMax((m) => !m)}
                onTemplatesChanged={() => {
                  void getTemplateStore()
                    .list()
                    .then(setLibs)
                    .catch(() => undefined)
                }}
              />
            ) : (
              <XlsxReportPanel
                table={table}
                tables={tables}
                records={records}
                onSetNodeConfig={onSetNodeConfig}
                refreshKey={bookVersion}
              />
            ))}

          {tab === 'book' &&
            (bookErr ? (
              <div className="wb-err">
                <b>读不到工作簿</b>
                <div className="wb-err-msg">{bookErr}</div>
                {/* ★ 批次17：模板丢失不是死胡同 —— 给三条可执行出路（重选已有模板 / 重新导入 / 建空白）。
                    原先此处只有一句报错，用户被卡死在「必须回画布重建节点」的死角。 */}
                <div className="wb-err-actions">
                  {libs.length > 0 && (
                    <select
                      className="ne-libsel"
                      value=""
                      title="从模板库重选一个工作簿绑定到本节点"
                      onChange={(e) => {
                        const hit = libs.find((l) => l.id === e.target.value)
                        if (hit) onSetNodeConfig(nodeId, { xlsxTemplateId: hit.id, xlsxTemplateName: hit.name })
                      }}
                    >
                      <option value="">从模板库重选…</option>
                      {libs.map((l) => (
                        <option key={l.id} value={l.id}>
                          {l.name}
                        </option>
                      ))}
                    </select>
                  )}
                  <button className="bar-btn primary" onClick={() => void openXlsxFile()}>
                    📂 重新导入…
                  </button>
                  <button className="bar-btn" onClick={() => void createBlankXlsx()}>
                    ＋ 建空白工作簿
                  </button>
                </div>
              </div>
            ) : !templateId ? (
              /* 未导入模板：整页空态，三个入口 */
              <div className="wb-hero">
                <div className="wb-hero-title">打开一个 Excel 工作簿</div>
                <div className="wb-hero-actions">
                  <button className="bar-btn primary" onClick={() => void openXlsxFile()}>
                    📂 导入模板…
                  </button>
                  <button className="bar-btn" onClick={() => void createBlankXlsx()}>
                    ＋ 创建空白文档
                  </button>
                </div>
                {libs.length > 0 && (
                  <div className="wb-hero-history">
                    <div className="wb-hero-h">历史文件</div>
                    {libs.map((l) => (
                      <button
                        key={l.id}
                        className="wb-hero-item"
                        onClick={() => onSetNodeConfig(nodeId, { xlsxTemplateId: l.id, xlsxTemplateName: l.name })}
                      >
                        {l.name}
                      </button>
                    ))}
                  </div>
                )}
                <div className="wb-hero-hint">导入、从历史选择或创建空白文档后，这里就是完整的工作簿编辑器。</div>
              </div>
            ) : !bookBytes ? (
              <div className="ne-empty">{busy || '正在打开…'}</div>
            ) : (
              <WorkbookEditor
                bytes={bookBytes}
                name={templateName}
                nodeId={nodeId}
                tables={tables}
                records={records}
                onSetNodeConfig={onSetNodeConfig}
                maximized={headMax}
                onToggleMax={() => setHeadMax((m) => !m)}
                onChange={(next) => void saveBookToStore(next)}
                onSave={(b) => void saveDocAs(templateName, b, 'xlsx')}
                theme={theme}
                /**
                 * ★ 章节编排插槽 —— 章节节点与普通 Excel 节点的**唯一**差别。
                 *   xlsx 节点不传 → 内核行为与从前逐字节一致（零回归）。
                 *   chapter 传配置 → 多出「审阅 → 章节编排」按钮与章节窗体，
                 *   而列布局、只读态都从节点配置里取，不写死在代码中。
                 */
                outline={
                  kind === 'chapter'
                    ? { cols: table?.config?.chapterCols, readOnly: false }
                    : undefined
                }
              />
            ))}

          {tab === 'data' && <InputDataView fields={computed.fields} rows={computed.rows} />}
        </div>

        {/* ── OO 节点左栏浮卡（普通模式：窗外左侧，Excel 同款）；最大化时参数引用走 /poc 内嵌面板（批注位）── */}
        {isOO && tab === 'main' && !headMax && ooRefOpen && neSidesReady && !ooRefWidthHidden && ooRect && (
          <RefPanel
            tables={tables}
            records={records}
            selfId={nodeId}
            activeRef={null}
            onPick={(srcId, srcName, fieldId, fieldName, value, mode) => {
              // ★ 用户操作可见化（2026-10-10 用户指令）：浮卡点击参数进诊断抽屉
              try { ooRef.current?.opLog?.('[操作] 点击参数（浮卡）: ' + (srcName || srcId) + '.' + fieldName) } catch { /* ignore */ }
              // ★ 同上：一次 patch 同时落 ooRefs + linkedInputs（形成关联）。
              //   另插一条可读文本（=节点名.字段名）进编辑器，让它像公式那样参与计算。
              const patch = ooRefPatchOf(table?.config, srcId, fieldId)
              onSetNodeConfig(nodeId, patch)
              try { ooRef.current?.opLog?.('[关联] 已落库 ooRefs+linkedInputs（画布边同步）') } catch { /* ignore */ }
              try {
                const src = (ooRefItems ?? []).find((x) => x.id === srcId)
                const f = src?.fields.find((x) => x.id === fieldId)
                ooRef.current?.insertRef?.({
                  nodeId: srcId,
                  nodeName: srcName || src?.name,
                  fieldId,
                  fieldName,
                  value: value != null && value !== '' ? value : f?.value,
                  mode: mode || 'value',
                })
              } catch { /* 编辑器未就绪则跳过 */ }
            }}
            onPickLibValue={(fieldName, value, mode, libNodeId) => {
              /* ★ 2026-10-23 用户指令「双击可以 链接至ooexcel或者ooword」：
                 ref 模式 = 真链接（ooRefPatchOf 落 ooRefs+linkedInputs → 画布边 +
                 insertRef 定义名称活引用，公式 =库名.字段）；value 模式 = 纯值插入 */
              try {
                if (mode === 'ref' && libNodeId) {
                  const patch = ooRefPatchOf(table?.config, libNodeId, fieldName)
                  onSetNodeConfig(nodeId, patch)
                  try { ooRef.current?.opLog?.('[参数库·链接] 已落库 ooRefs+linkedInputs（' + libNodeId + '）') } catch { /* ignore */ }
                }
                ooRef.current?.insertRef?.({
                  nodeId: libNodeId || 'paramlib-lib',
                  nodeName: '参数库',
                  fieldId: fieldName,
                  fieldName,
                  value,
                  mode: mode || 'value',
                })
                try { ooRef.current?.opLog?.('[参数库] ' + (mode === 'ref' ? '已链接 ' : '已插入 ') + fieldName + '：' + String(value).slice(0, 40)) } catch { /* ignore */ }
              } catch { /* 编辑器未就绪则跳过 */ }
            }}
            onClose={() => {
              setOoRefLeaving(true)
              // 立即撤让位类：窗体扩展动画与卡片退场动画同步（不等 ≤300ms 的下一轮测量）
              try { document.querySelector('.node-editor')?.classList.remove('ne-refroom') } catch { /* noop */ }
              setTimeout(() => { setOoRefLeaving(false); setOoRefOpen(false) }, 240)
            }}
            animClass={ooRefLeaving ? 'ref-panel-anim-out' : 'ref-panel-anim-in'}
            panelStyle={{
              position: 'fixed',
              left: Math.max(8, ooRect.left - 240),
              top: 28,
              height: 'calc(100vh - 56px)',
              width: 230,
              borderRadius: 10,
              border: '1px solid var(--ed-border)',
              zIndex: 8100,
              transition: 'left .2s ease',
            }}
          />
        )}
        {/* ── OO 节点右栏浮卡（普通模式：窗外右侧）── */}
        {/* ★ 2026-10-19：与左栏同款——窗口过窄（ooRefWidthHidden）时右栏整体隐藏 */}
        {isOO && tab === 'main' && !headMax && ooRightOpen && neSidesReady && !ooRefWidthHidden && ooRect && (
          <RightPanel
            tables={tables}
            records={records}
            selfId={nodeId}
            inputs={table?.config?.inputs ?? []}
            onToggle={(srcId, on) => {
              const cur = table?.config?.inputs ?? []
              const next = on ? (cur.includes(srcId) ? cur : [...cur, srcId]) : cur.filter((x) => x !== srcId)
              onSetNodeConfig(nodeId, { inputs: next })
            }}
            onClose={() => {
              setOoRightLeaving(true)
              setTimeout(() => { setOoRightLeaving(false); setOoRightOpen(false) }, 240)
            }}
            animClass={headMax ? undefined : ooRightLeaving ? 'ref-panel-anim-out-r' : 'ref-panel-anim-in-r'}
            panelStyle={{
              position: 'fixed',
              /* ★ 2026-10-19 修「自定义窗体大小时右栏显示不全」：普通态面板挂在窗体右侧
                 +10px，窗体靠屏幕右缘时直接出屏。按视口夹取——放不下时贴屏幕右缘
                 （最多盖住窗体右缘一点，但保证完整可见）。左栏有让位压缩逻辑所以没这问题。 */
              left: (() => {
                const want = ooRect.right + 10
                const maxX = window.innerWidth - 230 - 8
                return Math.max(8, Math.min(want, maxX))
              })(),
              top: ooRect.top,
              height: ooRect.height,
              width: 230,
              borderRadius: headMax ? 0 : 10,
              border: '1px solid var(--ed-border)',
              zIndex: 8100,
            }}
          />
        )}
      </div>
    </div>,
    document.body,
  )
}

/** 只读的输入数据视图：让用户在编辑器里就能确认「上游到底喂了什么」 */
function InputDataView({ fields, rows }: { fields: { id: string; name: string }[]; rows: Row[] }) {
  if (!fields.length) {
    return (
      <div className="ne-empty">
        上游还没有数据。把别的节点的蓝点拖到本节点左侧，或在属性面板勾选数据来源。
      </div>
    )
  }
  return (
    <div className="ne-table-wrap">
      <table className="ne-table">
        <thead>
          <tr>
            {fields.map((f) => (
              <th key={f.id}>{f.name}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.slice(0, 200).map((r, i) => (
            <tr key={i}>
              {fields.map((f) => (
                <td key={f.id}>{fmt(r[f.id])}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      {rows.length > 200 && (
        <div className="ne-more">
          仅显示前 200 行（共 {rows.length} 行）
        </div>
      )}
    </div>
  )
}

function fmt(v: unknown): string {
  if (v === undefined || v === null) return ''
  if (typeof v === 'number') return Number.isFinite(v) ? String(Math.round(v * 1e6) / 1e6) : String(v)
  return String(v)
}
