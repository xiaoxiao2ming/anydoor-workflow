import { memo, useEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import { createPortal } from 'react-dom'
import { preloadOo, preloadState, type OoPreloadState } from './office/ooPreload'
import {
  ReactFlow,
  Background,
  MiniMap,
  Panel,
  Handle,
  Position,
  BaseEdge,
  EdgeLabelRenderer,
  getBezierPath,
  useNodesState,
  useEdgesState,
  useReactFlow,
  useViewport,
  NodeResizer,
  useNodeId,
  useStore,
  type Node,
  type Edge,
  type EdgeProps,
  type XYPosition,
  type NodeTypes,
  type EdgeTypes,
  type Connection,
  type ReactFlowInstance,
  type NodeChange,
} from '@xyflow/react'
import '@xyflow/react/dist/style.css'
import { CanvasRibbon } from './CanvasRibbon'
import { CanvasChatDock } from './CanvasChatDock'
import { EdgeFxPicker, type EdgeFx } from './EdgeFxPicker'
import type { TableSchema, NodeKind, FieldDef } from './core/schema'
import type { Row, NodePos } from './core/store'
import { useBaseStore } from './core/store'
import { openCanvasFromWfcFile } from './wfcOpen'
import { FolderNodeCard } from './office/FolderNodeCard'
import { FileOutNodeCard } from './office/FileOutNodeCard'
import { Pdf2WordNodeCard, Pdf2WordImportButton, Pdf2WordEngineButton, Pdf2WordRunButton, Pdf2WordHistButton, Pdf2WordHistoryList } from './office/Pdf2WordNodeCard'
import { PdfReplaceNodeCard, PdfReplaceRunButton, PdfReplaceHistButton, PdfReplaceHistList } from './office/PdfReplaceNodeCard'
import { putBlob as putBlobStandalone, getBlob as getBlobHist } from './core/pdfBlobStore'
import { putContentAddressed } from './core/pdfBlobStore'
import { audit } from './core/audit'
import { IcZap, IcLock, IcUnlock, IcMapPin, IcRefresh, IcBan, IcHistory, IcTrash, IcShrink, IcFold, IcGrid, IcMaximize, IcMinimize } from './ui/icons'
import { hasPopupForm } from './PopupNodeEditor'
import { applyCanvasBg, readCanvasBg } from './ui/canvasBg'
/* ★ 2026-10-22 拆分：四组件移居 CanvasParts.tsx（改动隔离，防巨石回归） */
import { TestTableDialog, GrpLockButton, NodeLockButtons, AlignmentGuides, GroupNodeView, NoteStickyBody } from './CanvasParts'
import { PdfMergeButtons } from './office/PdfMergeButtons' 
import { prefGet, prefSet } from './ui/prefs'
import { layoutLayers, executionOrder } from './core/graph'
import { alignLayout, type AlignSize } from './core/align'
// 连线派生的单一真相源：与工程文件落盘（core/projectFormat）共用，见该文件顶部说明
import { edgesOf } from './core/projectFormat'
import { pmRowHandleId, PdfMergeRunButton, PdfExtraButtons, PdfHistoryList, convertToPdf } from './office/PdfMergeButtons'
import { PhotoNodeCard, PhotoImportButton } from './office/PhotoNodeCard'
import { PdfNodeCard, PdfImportButton, PdfGenButton, PdfGenHistButton, PdfGenHistoryList } from './office/PdfNodeCard'
import { ParamlibContentDialog } from './ParamlibContentDialog'
import { stripFullCalcFlag } from './core/ooDomains'
import { useNodeRunStatus, setNodeRunStatus, setNodeRunStatusAuto, setPrimaryRunNode } from './core/nodeRunStatus'
import { badgeClassOf, familyClassOf, styleOf, isWorkbookNodeKind, isRuleNodeKind, isRoundNodeKind, isGroupNodeKind, isMemberOfFoldedGroup, isMergeKind, GROUP_PAD, GROUP_HEAD, groupFitRect, unionRect } from './core/nodeStyle'
import { isRuleLocked, memberSummaryOf, membersOf, ruleNodesOf, roundLinkCountOf } from './core/ruleSync'
// 连接口角色：id 常量 + 准入判定。判定本体在 ruleSync（纯函数、可单测），
// 这里只做接线—— 与 ruleConnectPatch 同一个「不内联判定」的理由。
import { RULE_PORT_IN, RULE_PORT_IN_2, RULE_PORT_OUT, RULE_PORT_OUT_2, AI_PORT_IN, PARAMLIB_PORT_OUT, canConnectPorts, isRulePortId, roundLineAnchors, roundLineKindOf } from './core/ruleSync'
// 运行动画的判定与状态机（纯函数、单测在 core/roundRun.test.ts）：
// 「谁该播」（roundRunStateOf，看 selected）与「挂哪个 class」（roundRunClassOf）
// 都在 core，这里只做接线 —— 与 ruleConnectPatch / roundLineAnchors 同一个理由。
import { roundRunStateOf, roundRunClassOf, nextRoundRunState, hasPlayableSelection, ROUND_RUN_TEST_MS, ROUND_RUN_EXIT_MS } from './core/roundRun'
import { InlineCapsules } from './ref/InlineCapsules'
import { buildRefMark, offsetAfterInsert, plainLength } from './ref/inline'
import { pushLog } from './ui/logBus'
import { RunStatusBar } from './ui/RunStatusBar'
import { WebmonNodeCard } from './office/WebmonNodeCard'
import { Xlsx2dxfNodeCard } from './office/Xlsx2dxfNodeCard'
import { parseFile } from './core/io'
import { useWebmonPoller } from './core/webmon'
import { getTemplateStore } from './docx/templateStore'
import { putBlob, getBlob } from './core/pdfBlobStore'
import { PerfHud } from './office/PerfHud'
import { editorOpenFrom } from './ui/editorOpenFrom'
import { intentById } from './core/chatIntent'
import { runPdfMerge } from './office/PdfMergeButtons'

/** 胶囊拖放的 dataTransfer MIME（与节点拖放 application/reactflow 区分开） */
const CAP_DRAG_MIME = 'application/capsule-mark'

/** 双击可进入编辑器的节点类型（态 B 放大态） */
/* ★ 2026-10-23 pdf 移出 EDITABLE_KINDS：双击=iframe 秒开预览（高频）；
   OO 编辑走右键菜单（转换 10~30s，知情选择） */
const EDITABLE_KINDS: NodeKind[] = ['docx', 'xlsx', 'ooword', 'oosheet', 'chapter']

/* ── OO 文档节点三按钮（2026-10-07 用户指令）────────────────────────
   ⚡预加载：后台上传样例/模板 + 打开 + 冻结（iframe 脱离 DOM，现场保留 CPU≈0）
             → 打开节点时从冻结现场原地恢复（0 等待，见 office/ooPreload.ts）
   🔒数据：config.ooDataLocked → poc 以 mode=view 打开（编辑器只读，防误编辑）
   📍位置：config.ooPosLocked → React Flow node.draggable=false（防误拖）  */
/** base64 → Uint8Array（OO 节点内置内容 ooBytes64 解码用，2026-10-11） */
/* ★ 2026-10-24 性能根治：OO 文档 bytes 入 IndexedDB（config 只留 ooBytesDbId 索引）——
   ooBytes64 曾以 MB 级 base64 挂在 config，拖停保存的全量序列化被它拖死 */
async function stashOoBytes (dbKey: string, bytes: Uint8Array): Promise<string> {
  await putBlob(dbKey, bytes)
  return dbKey
}
async function b64FromDbId (dbId: string): Promise<string | null> {
  try {
    const u8 = await getBlob(dbId)
    if (!u8) return null
    let bin = ''
    const CH = 0x8000
    for (let i = 0; i < u8.length; i += CH) bin += String.fromCharCode(...u8.subarray(i, i + CH))
    return btoa(bin)
  } catch { return null }
}
function b64ToU8(b64: string): Uint8Array {
  const bin = atob(b64)
  const u8 = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
  return u8
}
/** Uint8Array → base64（拖放文件导入 2026-10-23）。
 *  ★ 分块 String.fromCharCode 后必须 btoa——直接返回二进制字符串当 base64
 *    就是今天 /pdfconv 乱码的终极真凶（对账日志实锤），此处一步到位。 */
function b64FromU8(u8: Uint8Array): string {
  let bin = ''
  const CHUNK = 0x8000
  for (let i = 0; i < u8.length; i += CHUNK) bin += String.fromCharCode(...u8.subarray(i, i + CHUNK))
  return btoa(bin)
}

/* ★ 2026-11-08 对话「word转PDF」：与右键「转为PDF节点」（2026-10-29 用户指令）同一条链路，
   抽成模块函数共用：取 OO 文档最新字节（ooBytes64 → ooBytesDbId → 模板库，与预加载同优先级）
   → convertToPdf（x2t 秒转）→ 右侧新建 PDF 节点装入产物 + linkedInputs 连线。
   右键调用只看日志；对话调用把返回值回派 chat-topdf-result 填气泡。 */
async function runNodeToPdf(
  nodeId: string,
  opts: {
    onAddTableAt: (pos: XYPosition, kind?: NodeKind) => string
    pendingPos: { current: Record<string, XYPosition> }
    nodesRef: { current: Array<{ id: string; position: XYPosition; measured?: { width?: number } }> }
  },
): Promise<{ ok: boolean; msg: string }> {
  const st = useBaseStore.getState()
  const t = st.tables[nodeId]
  if (!t || t.deleted) return { ok: false, msg: '节点不存在（可能已被删除）' }
  if (t.kind !== 'ooword' && t.kind !== 'oosheet') {
    return { ok: false, msg: `「${t.name}」是${styleOf(t.kind).label}节点，只有 Word/Excel 文档节点能转 PDF` }
  }
  const cfg = (t.config ?? {}) as { ooBytes64?: string; ooBytesDbId?: string; ooTemplateId?: string }
  let ooB = cfg.ooBytes64
  if (!ooB) {
    const dbId = cfg.ooBytesDbId
    if (dbId) { const r = await b64FromDbId(dbId); if (r) ooB = r }
  }
  let u8: Uint8Array | null = ooB ? b64ToU8(ooB) : null
  if (!u8 || !u8.length) {
    try { u8 = (await getTemplateStore().read(cfg.ooTemplateId ?? '')) || null } catch { u8 = null }
  }
  if (!u8 || !u8.length) {
    pushLog('[OO→PDF] 节点无可转换内容（未导入/未保存）', 'warn')
    /* ★ 2026-11-09 操作必有反馈：无内容也上状态条（error 态），别让执行「无声无息」 */
    setPrimaryRunNode(nodeId)
    setNodeRunStatus(nodeId, 'error', '没有可转换内容——先导入或保存一次')
    return { ok: false, msg: `「${t.name}」还没有可转换内容——先双击节点导入或编辑保存一次，再转 PDF。` }
  }
  const t0 = Date.now()
  const baseName = String(t.name ?? '文档').replace(/\.(docx|xlsx)$/i, '')
  pushLog(`[OO→PDF] 转换中：${baseName}（${Math.round(u8.length / 1024)}KB）…`)
  /* ★ 2026-11-09 运行状态条：转 PDF 也作为主进程上条 */
  setPrimaryRunNode(nodeId)
  setNodeRunStatus(nodeId, 'running', `转换 PDF…`)
  try {
    const pdf = await convertToPdf(u8, baseName + (t.kind === 'oosheet' ? '.xlsx' : '.docx'))
    if (!pdf || !pdf.length) throw new Error('转换产物为空')
    const src = opts.nodesRef.current.find((n) => n.id === nodeId)
    const p = { x: Math.round((src?.position.x ?? 0) + (src?.measured?.width ?? 220) + 60), y: Math.round(src?.position.y ?? 0) }
    const stP = useBaseStore.getState()
    stP.pushHistory()
    const newId = opts.onAddTableAt(p, 'pdf')
    opts.pendingPos.current[newId] = p
    const pid = 'pd-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
    await putBlobStandalone('pdfitem-' + pid, pdf)
    stP.setNodeConfig(newId, { pdfItems: [{ id: pid, name: baseName + '.pdf', inDb: true }], pdfFileOrder: [pid], linkedInputs: [nodeId] })
    stP.renameTable(newId, baseName)
    const msg = `已转成 PDF 节点「${baseName}」并连线自「${t.name}」：${Math.round(pdf.length / 1024)}KB（${((Date.now() - t0) / 1000).toFixed(1)}s）`
    pushLog(`[OO→PDF] ${msg}`, 'ok')
    setNodeRunStatusAuto(nodeId, 'ok', `PDF ${Math.round(pdf.length / 1024)}KB（${((Date.now() - t0) / 1000).toFixed(1)}s）`)
    return { ok: true, msg }
  } catch (eP) {
    const msg = `转换失败：${String((eP as Error)?.message ?? eP).slice(0, 120)}`
    pushLog(`[OO→PDF] ${msg}`, 'err')
    setNodeRunStatus(nodeId, 'error', '转换失败')
    return { ok: false, msg }
  }
}

/**
 * ★ 2026-10-29 用户规则：OO 空白节点可拖入 / 右键导入外部文件作为文档。
 * 空白节点（无 ooTemplateId/ooBytes64/ooBytesDbId）直接导入；已有内容则弹确认替换；
 * 按扩展名自动切换 ooword/oosheet（与节点标题栏导入同一口径）。
 * 字节落 IndexedDB（ooBytesDbId），与画布节点 config 同走既有加载链路，重开自动恢复。
 */
async function importOoFileToNode(
  target: { id: string; name: string; kind?: string; config?: object } | undefined,
  files: File[],
): Promise<void> {
  if (!target) return
  const st = useBaseStore.getState()
  for (const f of files) {
    const ext = (f.name.split('.').pop() ?? '').toLowerCase()
    if (ext !== 'docx' && ext !== 'xlsx') continue
    const wantKind = ext === 'docx' ? 'ooword' : 'oosheet'
    const cur = st.tables[target.id]
    const cfg = (cur?.config ?? {}) as { ooTemplateId?: string; ooBytes64?: string; ooBytesDbId?: string }
    const hasContent = !!(cfg.ooTemplateId || cfg.ooBytes64 || cfg.ooBytesDbId)
    if (hasContent && !window.confirm(`将「${f.name}」替换「${target.name}」当前文档？`)) continue
    /* 按扩展名切 OO 节点类型（docx→ooword / xlsx→oosheet；徽章/面板/样式全跟随） */
    if (wantKind !== cur?.kind) { st.pushHistory(); st.setTableKind(target.id, wantKind) }
    try {
      const bytes = new Uint8Array(await f.arrayBuffer())
      st.pushHistory()
      st.setNodeConfig(target.id, {
        ooTemplateName: f.name,
        ooBytesDbId: await stashOoBytes('oobytes-' + target.id, bytes),
        ooTemplateId: 'tpl-' + Date.now().toString(36),
      })
      pushLog(`[OO·导入] 已${hasContent ? '替换' : '导入'}「${target.name}」：${f.name}`, 'ok')
    } catch (e) {
      pushLog(`[OO·导入] 失败：${f.name} — ${String((e as Error)?.message ?? e).slice(0, 80)}`, 'err')
    }
  }
}

/* ★ 2026-10-26 右键「用默认软件/本地浏览器打开」共用取字节（pdf=词条最新字节——
   原 inline 逻辑抽取；pdfmerge=合并历史最新条目）。返回 null=无可打开内容。 */
async function pdfOpenTargetOf (
  tables: Record<string, { kind?: string; config?: object } | undefined>,
  id: string,
): Promise<{ name: string; bytes: Uint8Array } | null> {
  const t = tables[id]
  if (!t) return null
  if (t.kind === 'pdfmerge') {
    const cfgM = (t.config ?? {}) as { pdfHistory?: { name?: string; b64?: string; id: string; at?: number }[] }
    const latest = (cfgM.pdfHistory ?? [])[0]
    if (!latest) return null
    const { getBlob: gb } = await import('./core/pdfBlobStore')
    const bytes = latest.b64 ? b64ToU8(latest.b64) : await gb('pdfhist-' + latest.id + '-' + (latest.at ?? 0))
    return bytes ? { name: latest.name ?? '合并产物.pdf', bytes } : null
  }
  const cfg0 = (t.config ?? {}) as { pdfItems?: { name: string; b64?: string; inDb?: boolean; id: string }[] }
  const first = (cfg0.pdfItems ?? [])[0]
  if (!first) return null
  const { getBlob: gb2 } = await import('./core/pdfBlobStore')
  const bytes = first.b64 ? b64ToU8(first.b64) : await gb2('pdfitem-' + first.id)
  return bytes ? { name: first.name || '文档.pdf', bytes } : null
}
function OoNodeButtons({ nodeId, kind, compact }: { nodeId: string; kind: NodeKind; compact?: boolean }) {
  const [preloadSt, setPreloadSt] = useState<OoPreloadState>(() => preloadState(nodeId))
  // config 订阅式读取（store 变化 → 按钮/节点卡片同步刷新）
  // ★ 2026-10-26 统一锁：读写迁移到通用 dataLocked/posLocked（ooDataLocked/ooPosLocked 读取兼容）
  const cfg = useBaseStore((s) => s.tables[nodeId]?.config) as { ooDataLocked?: boolean; ooPosLocked?: boolean; dataLocked?: boolean; posLocked?: boolean; ooRefUpdMode?: string } | undefined
  const dataLocked = !!(cfg?.dataLocked || cfg?.ooDataLocked)
  const posLocked = !!(cfg?.posLocked || cfg?.ooPosLocked)
  /* ★ 外部数据引用更新策略（2026-10-18 用户指令「这个更新设置 同样放在节点上 保持数据一致」）：
   *  与编辑器标题栏设置卡共用同一个 config.ooRefUpdMode；节点上点一下循环切换。
   *  读取回退 localStorage（旧数据），写入同时落 localStorage 与节点配置。 */
  const REF_UPD_LABEL: Record<string, string> = { off: '不检测', ask: '检测后提示', auto: '自动更新' }
  const updMode = ((): 'off' | 'ask' | 'auto' => {
    const c = cfg?.ooRefUpdMode
    if (c === 'off' || c === 'ask' || c === 'auto') return c
    try {
      const v = localStorage.getItem('ooRefUpdateMode')
      if (v === 'off' || v === 'ask' || v === 'auto') return v
    } catch { /* ignore */ }
    return 'ask'
  })()
  const cycleUpd = () => {
    const next: 'off' | 'ask' | 'auto' = updMode === 'off' ? 'ask' : updMode === 'ask' ? 'auto' : 'off'
    toggle({ ooRefUpdMode: next })
    try { localStorage.setItem('ooRefUpdateMode', next) } catch { /* ignore */ }
  }
  const doPreload = async () => {
    if (preloadSt !== 'idle') return
    setPreloadSt('loading')
    try {
      const t = useBaseStore.getState().tables[nodeId]
      const tplId = (t?.config as { ooTemplateId?: string } | undefined)?.ooTemplateId
      /* ★ 2026-10-11：预加载优先用节点内置内容（ooBytes64，随画布走）——模板库（IndexedDB）读不到也不退样例 */
      let ooBytes = (t?.config as { ooBytes64?: string } | undefined)?.ooBytes64
      if (!ooBytes) {
        const dbId = (t?.config as { ooBytesDbId?: string } | undefined)?.ooBytesDbId
        if (dbId) { const r = await b64FromDbId(dbId); if (r) ooBytes = r }
      }
      await preloadOo({ nodeId, kind: kind as 'ooword' | 'oosheet', templateId: tplId, bytes: ooBytes ? b64ToU8(ooBytes) : null })
      setPreloadSt('frozen')
    } catch (e) {
      setPreloadSt('idle')
      console.warn('[ooPreload] 预加载失败（节点正常打开不受影响）:', e)
    }
  }
  const toggle = (patch: Record<string, boolean | string>) => useBaseStore.getState().setNodeConfig(nodeId, patch)
  // ★ 2026-10-19 用户「大小颜色也按照刚才改的」：31×31（与标题栏一致），图标一律
  //   线性 SVG + currentColor（弃彩色 emoji），激活=主题色淡底（与 is-on 同语言）
  const sq = (active: boolean, danger = false): CSSProperties => ({
    width: compact ? 24 : 31, height: compact ? 24 : 31, padding: 0,
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    lineHeight: 1, cursor: 'pointer', borderRadius: 6,
    border: active ? `1px solid ${danger ? '#d29922' : 'var(--ed-accent-bd, rgba(159,225,203,.42))'}` : '1px solid var(--c-border)',
    background: active ? (danger ? 'rgba(210,153,34,0.18)' : 'var(--ed-accent-bg, rgba(159,225,203,0.14))') : 'transparent',
    color: active ? (danger ? '#d29922' : 'var(--ed-accent, #9fe1cb)') : 'var(--ed-text3, var(--c-text3))',
    flexShrink: 0,
  })
  return (
    <div className="rf-field rf-oo-btns" style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
      <button
        className="rf-oo-btn"
        style={sq(preloadSt === 'frozen')}
        title={preloadSt === 'frozen'
          ? '已预加载（冻结待命）—— 打开节点 0 等待'
          : preloadSt === 'loading'
            ? '后台加载中…'
            : '预加载：后台打开文档并冻结，打开节点 0 等待'}
        onClick={() => void doPreload()}
      >
        <IcZap size={compact ? 14 : 19} />
      </button>
      <button
        className="rf-oo-btn"
        style={sq(dataLocked)}
        title={dataLocked ? '数据已锁定：文档只读（点击解锁）' : '数据锁定：文档以只读模式打开（防误编辑；重开保持）'}
        onClick={() => toggle({ dataLocked: !dataLocked })}
      >
        {dataLocked ? <IcLock size={compact ? 14 : 19} /> : <IcUnlock size={compact ? 14 : 19} />}
      </button>
      <button
        className="rf-oo-btn"
        style={sq(posLocked, true)}
        title={posLocked ? '位置已锁定：节点不可拖动（点击解锁）' : '位置锁定：节点在画布上不可拖动（防误拖；重开保持）'}
        onClick={() => toggle({ posLocked: !posLocked })}
      >
        <IcMapPin size={compact ? 14 : 19} />
      </button>
      {kind === 'oosheet' && (
        <button
          className="rf-oo-btn"
          style={sq(updMode !== 'off')}
          title={'外部数据引用更新：' + (REF_UPD_LABEL[updMode] || updMode)
            + '（随节点保存）—— 点击切换：不检测 → 检测后提示 → 自动更新'}
          onClick={cycleUpd}
        >
          {updMode === 'off' ? <IcBan size={compact ? 14 : 19} /> : <IcRefresh size={compact ? 14 : 19} />}
        </button>
      )}
      {/* ★ 锁定组按钮排进行尾（2026-10-19 用户「按钮放在一行」） */}
      <GrpLockButton nodeId={nodeId} />
    </div>
  )
}

/* 折叠态 logo 的字（2026-10-06 二轮迭代，用户：「不要只写一个字 参数表就原来大小的
   底部圆角矩形颜色填充尺寸 logo字体也一致」）：
   不再维护单字映射表 FOLD_LOGO —— 直接复用展开态徽章的整套 DOM+类名+文字源
   （.rf-badge + badgeClassOf + styleOf(kind).label），字体/圆角/颜色填充/尺寸
   与展开态天然一致，不存在两套样式漂移。「参数表」就是 table 的 label，
   用户原话点名了它。 */

interface TableNodeData {
  name: string
  kind: NodeKind
  /** 节点 id：data 里未必带（React Flow 把它放在节点对象的 id 上），故正文靶用 useNodeId()取 */
  id?: string
  fields: { id: string; name: string; type: string }[]
  /** docx/xlsx 节点：模板/绑定概要，显示在节点卡片上（如 "模板已绑 · 3 域"） */
  boundHint?: string
  /** ★ 节点组（2026-10-21）：拖动节点悬停在本组上 → 高亮 + 提示「置入节点至节点组」 */
  groupHint?: boolean
  /** text 节点：当前内容（数字或一串文字） */
  textValue?: string
  /** text 节点：显示格式（备忘 · WPS 风格） */
  textStyle?: { bold?: boolean; italic?: boolean; underline?: boolean; fontSize?: number; color?: string; align?: string }
  /**
   * 规则跟随节点双击 → 打开成员管理面板。
   * 走 data 而非 prop：nodeTypes 是模块级常量，拿不到组件 prop。
   */
  onOpenRuleMembers?: OpenRuleMembers
  selected: boolean
  running: boolean
  /** ★ 统一锁（2026-10-26）：数据锁或位置锁任一开启 → rf-node--locked（合成层提升+头部锁标） */
  locked?: boolean
  [k: string]: unknown
}

function defaultPos(i: number): XYPosition {
  return { x: 40 + (i % 3) * 280, y: 40 + Math.floor(i / 3) * 180 }
}

/**
 * docx/xlsx 节点卡片上的一行概要：模板是否已绑、绑了几个「域」。
 * 只做展示，不参与计算 —— 真正的绑定关系在 config 里。
 */
function boundHintOf(t: TableSchema): string {
  if (t.kind === 'docx') {
    if (!t.config?.docxTemplateId) return '未绑模板'
    return `${t.config.docxTemplateName || '模板'} · ${Object.keys(t.config.docxMap ?? {}).length} 域`
  }
  /* ★ 2026-10-22 补 OO 系列分支（用户问「未绑定什么意思」引出的显示 bug）：
     oosheet/ooword 走 ooTemplateId 绑定，原先不在 WORKBOOK_KINDS 也不在 docx 分支，
     boundHint 返回空串 → 卡片底部永远显示兜底文案「未绑定」，绑定后也不变。 */
  if (t.kind === 'oosheet' || t.kind === 'ooword') {
    const cfg = t.config as { ooTemplateId?: string; ooTemplateName?: string } | undefined
    if (!cfg?.ooTemplateId) return t.kind === 'ooword' ? '未绑文档' : '未绑工作簿'
    return cfg.ooTemplateName || '已绑定'
  }
  /* ★ 2026-10-23 用户指令「照片大于等于2 就标记为照片组」：1 张=照片，≥2=照片组 */
  if (t.kind === 'photo') {
    const n = ((t.config as { photoItems?: unknown[] } | undefined)?.photoItems ?? []).length
    if (!n) return '空'
    return n >= 2 ? `照片组 · ${n} 张` : '1 张照片'
  }
  /* ★ 2026-10-23 PDF 文件节点：同照片语言（1 份=PDF，≥2=PDF 组） */
  if (t.kind === 'fileout') {
    const cfgF = (t.config ?? {}) as { outDir?: string; lastSaved?: unknown[] }
    if (!cfgF.outDir) return '未设输出目录'
    return (cfgF.lastSaved ?? []).length ? `已存 ${(cfgF.lastSaved ?? []).length} 次` : '目录已设'
  }
  if (t.kind === 'pdf' || t.kind === 'pdfslim') {
    const n = ((t.config as { pdfItems?: unknown[] } | undefined)?.pdfItems ?? []).length
    if (!n) return '空'
    return n >= 2 ? `PDF 组 · ${n} 份` : '1 份 PDF'
  }
  // 章节工作簿与 Excel 共用模板库绑定，故同一套提示（单一真相源 isWorkbookNodeKind）
  if (isWorkbookNodeKind(t.kind)) {
    if (!t.config?.xlsxTemplateId) return t.kind === 'chapter' ? '未绑工作簿·无章节' : '未绑工作簿'
    return `${t.config.xlsxTemplateName || '工作簿'} · ${Object.keys(t.config.xlsxMap ?? {}).length} 域`
  }
  return ''
}

/** 规则跟随节点的成员管理面板：双击圆形节点时打开（App 层持有 state） */
export type OpenRuleMembers = (ruleNodeId: string) => void

/** 圆圈下方最多显示几个成员名标签，超出的折成「+N」。
 *  4 是权衡值：3 个常见（主 + 2 从）能全显示，5 个以上不至于糊住画布。
 *  ——完整名单在 title 悬停与双击面板里都能看到，不丢信息。 */
const RULE_TAGS_MAX = 4

/* ★ 数据表节点（测试）· 参数编辑弹窗（2026-10-19 用户指令）：
   「点击弹窗 左边是参数名 中间是数据格式 右边是数字」——三列布局：
   参数名（可改名，行键同步换）/ 数据格式（文本|数字下拉）/ 值（按格式解析）。
   即改即存（setTableData → 画布落盘），引用面板与 OO 引用链读到的就是这里的数据。 */
/* ★ 锁定组（2026-10-19 用户指令）：主节点 + 连着它的圆形节点（AI/参数库/规则跟随）。
   成员判定双向：AI 的 linkedInputs 含主节点（AI 吃主节点数）；
   主节点 linkedInputs 含参数库/规则（它们供数/发规则）。 */
function groupMembersOf(regId: string, t: Record<string, TableSchema | undefined>): string[] {
  const reg = t[regId]
  if (!reg || isRoundNodeKind(reg.kind)) return [regId]
  const out = [regId]
  const regLi = (reg.config?.linkedInputs ?? []) as string[]
  for (const n of Object.values(t)) {
    if (!n || n.deleted || !isRoundNodeKind(n.kind)) continue
    const li = (n.config?.linkedInputs ?? []) as string[]
    if (li.includes(regId) || regLi.includes(n.id)) out.push(n.id)
  }
  return out
}


/* ★ 锁定组按钮（2026-10-19）：主节点上显示——锁定后整组一起移动。
   ★ 常驻渲染（用户「刚生成的按钮没了」——不再按有无连线隐藏）：
   没连圆形节点时置灰禁用（is-empty），连上后自动可用。 */

/* ★ 节点组本体（2026-10-21；四轮 20:00 换 RF 原生 NodeResizer）：
   - 拖动整组只认标题栏（node.dragHandle = .group-node-head）
   - 缩放换 RF 原生 <NodeResizer>：拖边框/四角**原生跟手**（此前自写条不跟鼠标），
     isVisible 仅组选中时显示；松手把最终位置/尺寸落库（config.groupW/H + pos）
   - 六钮：解散 / 删组及成员 / 适应大小 / 网格重排 / 折叠成员 / 放大还原 */


/** ★ PDF 拖入空白区五选一（2026-10-26 用户指令）：
 *  1 创建PDF节点 / 2 创建PDF合并节点 / 3 创建PDF瘦身节点 / 4 创建PDF替换页节点 / 5 什么都不做（误操作）。
 *  遮罩点击 / × / Esc = 取消；数字键 1-5 快选。样式走 pr-confirm 语言（pd-choose）。 */
/** 拖入空白区的文件分组 + 用户可选动作（2026-10-26 全格式参照 PDF 规范） */
type DropGroups = { photos: File[]; words: File[]; sheets: File[]; pdfs: File[] }
type DropChoice = 'photos' | 'words' | 'sheets' | 'pdf' | 'merge' | 'slim' | 'replace' | 'all' | 'none' | 'words-open' | 'sheets-open' | 'sheets-dxf'

/** 按扩展名/类型分类拖入文件（照片 8MB / PDF 200MB 上限，与旧导入一致） */
function classifyDrop (files: FileList): { g: DropGroups; rejected: string[] } {
  const g: DropGroups = { photos: [], words: [], sheets: [], pdfs: [] }
  const rejected: string[] = []
  for (const f of Array.from(files)) {
    const ext = f.name.toLowerCase().split('.').pop() ?? ''
    if (f.type === 'image/jpeg' || f.type === 'image/png' || ext === 'jpg' || ext === 'jpeg' || ext === 'png') {
      if (f.size <= 8 * 1024 * 1024) g.photos.push(f)
      else rejected.push(`${f.name}（超 8MB）`)
    } else if (ext === 'docx') g.words.push(f)
    else if (ext === 'xlsx') g.sheets.push(f)
    else if (ext === 'pdf') {
      if (f.size <= 200 * 1024 * 1024) g.pdfs.push(f)
      else rejected.push(`${f.name}（超 200MB）`)
    } else rejected.push(`${f.name}（${ext || '未知'} 类型）`)
  }
  return { g, rejected }
}

/** ★ 拖入空白区选择对话框（2026-10-26 用户指令「PDF五选一」+「其他格式同样参照该规范」）：
 *  每种文件类型一行建节点动作（PDF 四动作），多类型混合时追加「全部导入」；
 *  遮罩点击 / × / Esc = 什么都不做（误操作取消）；数字键 1-N 快选。样式 pd-choose。 */
function DropChooser ({ g, rejected, onPick }: {
  g: DropGroups
  rejected: string[]
  onPick: (c: DropChoice) => void
}) {
  const rows: { k: DropChoice; title: string; desc: string; none?: boolean }[] = []
  if (g.photos.length) rows.push({ k: 'photos', title: '创建照片节点', desc: `载入 ${g.photos.length} 张图片（每图一页，可连 PDF 生成）` })
  if (g.words.length) {
    rows.push({ k: 'words', title: '创建 Word 文档节点', desc: `${g.words.length} 份，每份一个 OO 文档节点` })
    /* ★ 2026-10-29 用户指令「中间加个创建word节点并打开」：建完立即进入编辑 */
    rows.push({ k: 'words-open', title: '创建 Word 文档节点并打开', desc: `${g.words.length} 份，建完立即打开第 1 份进入编辑` })
  }
  if (g.sheets.length) {
    rows.push({ k: 'sheets', title: '创建 Excel 工作簿节点', desc: `${g.sheets.length} 份，每份一个 OO 工作簿节点` })
    rows.push({ k: 'sheets-open', title: '创建 Excel 工作簿节点并打开', desc: `${g.sheets.length} 份，建完立即打开第 1 份进入编辑` })
    /* ★ 2026-11-10 用户指令「拖动 xlsx 等文件 直接放在上」：坐标表一键展点 */
    rows.push({ k: 'sheets-dxf', title: '创建 Excel转DXF 节点', desc: '第 1 份作为坐标数据载入（点号/X/Y 列映射后展点成图）' })
  }
  if (g.pdfs.length) {
    rows.push({ k: 'pdf', title: '创建 PDF 节点', desc: `载入 ${g.pdfs.length} 份 PDF（文件组，可再连合并）` })
    rows.push({ k: 'merge', title: '创建 PDF 生成节点', desc: `载入 ${g.pdfs.length} 份 PDF 并自动连入生成（拼接）` })
    rows.push({ k: 'slim', title: '创建 PDF 瘦身节点', desc: `载入 ${g.pdfs.length} 份 PDF（压缩体积）` })
    rows.push({ k: 'replace', title: '创建 PDF 替换页节点', desc: `载入 ${g.pdfs.length} 份到被替换组（页码 = 被换的页）` })
  }
  const multi = [g.photos.length, g.words.length, g.sheets.length, g.pdfs.length].filter((n) => n > 0).length > 1
  if (multi) rows.push({ k: 'all', title: '全部导入', desc: '以上各类各建各的节点，横向排列' })
  rows.push({ k: 'none', title: '什么都不做', desc: '误操作，取消本次拖入', none: true })
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') { onPick('none'); return }
      const i = Number(e.key)
      if (i >= 1 && i <= rows.length) onPick(rows[i - 1].k)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onPick, rows.length])
  const parts = [
    g.photos.length && `照片 ${g.photos.length} 张`,
    g.words.length && `Word ${g.words.length} 份`,
    g.sheets.length && `Excel ${g.sheets.length} 份`,
    g.pdfs.length && `PDF ${g.pdfs.length} 份`,
    rejected.length && `不支持 ${rejected.length} 个`,
  ].filter(Boolean) as string[]
  const nameList = [g.photos, g.words, g.sheets, g.pdfs].flat().map((f) => f.name)
  return createPortal(
    <div className="pd-choose-mask nodrag nopan" onClick={() => onPick('none')}>
      <div className="pd-choose" onClick={(e) => e.stopPropagation()}>
        <div className="pd-choose-head">
          <span>检测到 {parts.join('、')}</span>
          <button className="pd-choose-x" title="关闭（Esc）" onClick={() => onPick('none')}>×</button>
        </div>
        <div className="pd-choose-names" title={nameList.join('、')}>
          {nameList.slice(0, 3).join('、')}{nameList.length > 3 ? ` 等 ${nameList.length} 个` : ''}
        </div>
        {rows.map((o, i) => (
          <button key={o.k} className={'pd-choose-opt' + (o.none ? ' pd-choose-opt--none' : '')}
            onClick={() => onPick(o.k)}>
            <span className="pd-choose-key">{i + 1}</span>
            <span className="pd-choose-t"><b>{o.title}</b><i>{o.desc}</i></span>
          </button>
        ))}
        <div className="pd-choose-hint">Esc 或点空白处取消 · 数字键快选</div>
      </div>
    </div>,
    document.body
  )
}

/** 节点顶部运行状态徽章（n8n/Dify 风格，2026-10-23 用户指令）：
 *  运行中蓝底 spinner、成功绿底（文案含耗时）、失败红底（错误摘要）。
 *  状态源 core/nodeRunStatus（内存态，ok 8s 自动淡出）。 */
function RunBadge ({ id }: { id: string }) {
  const st = useNodeRunStatus(id)
  /* ★ 2026-10-26 用户指令「状态信息 增加一个时间信息」：running 态实时耗时
     （每 200ms 跳动，X.Xs，起点=状态置位时刻；完成态自带耗时定格不重复加） */
  const [, forceTick] = useState(0)
  useEffect(() => {
    if (st?.state !== 'running') return
    const t = setInterval(() => forceTick((n) => n + 1), 200)
    return () => clearInterval(t)
  }, [st?.state, st?.at])
  if (!st) return null
  const elapsed = st.state === 'running' ? ((Date.now() - st.at) / 1000).toFixed(1) + 's' : ''
  return (
    <div className={'rn-badge nodrag nopan' + ' rn-badge--' + st.state} title={st.label}>
      {st.state === 'running' && <span className="rn-spin" />}
      <span className="rn-label">{st.label}</span>
      {elapsed && <span className="rn-elapsed">{elapsed}</span>}
    </div>
  )
}

function TableNodeInner({ data, selected }: { data: TableNodeData; selected?: boolean }) {
  const isSel = selected || data.selected
  /* ★ 数据表节点（测试）：点「✎ 参数编辑」开三列弹窗（参数名/数据格式/值） */
  const [ttDlgOpen, setTtDlgOpen] = useState(false)
  /* ★ 2026-10-23 用户指令「拖 word/excel 到相应节点：空白直接导入，有文件提示替换」：
     ooword 接 .docx / oosheet 接 .xlsx——depth 计数防子元素抖动，拖入高亮虚线框；
     类型不匹配接管 drop 并日志说明（防冒泡到画布误建新节点）。 */
  const ooDropDepth = useRef(0)
  const [ooDropHi, setOoDropHi] = useState(false)
  const isOOFileNode = data.kind === 'ooword' || data.kind === 'oosheet'
  // 备忘 2026-09-27 ⑤：fx/FCE 悬停胶囊 → 来源节点画布边框亮显（store 瞬态 fxHoverNode 桥）
  const selfId = useNodeId()
  /* ★ 2026-10-23 用户指令「双击圆形参数库节点 打开参数库选择」：
     选择后 libName 写 config——圆内显示已选库名 */
  const [libPickOpen, setLibPickOpen] = useState(false)
  const isParamlib = data.kind === 'paramlib'
  const onParamlibDoubleClick =
    isParamlib && selfId != null ? () => setLibPickOpen(true) : undefined
  const hoverNode = useBaseStore((s) => s.fxHoverNode)
  // 行内胶囊的节点名/失效态需要全量节点表（批次 12）：直接读 store，
  // 不给 TableNode 加 prop —— 节点渲染不应为胶囊反查而依赖整个 tables 字典。
  const allTables = useBaseStore((s) => s.tables)
  const hoverGlow = hoverNode != null && hoverNode === selfId
  // 圆形节点（2026-10-06 起三家）：rule 规则跟随 / ai / paramlib 参数库。
  //  · rf-node--rule        圆形本体（三家共用一套圆形结构，族色靠 kind 类覆盖）
  //  · rf-node--ai/paramlib 各自族色（CSS 变量覆盖，见 index.css）
  //  · rf-node--rule-locked 被别的规则节点锁定 → 加外环，说明「规则不由我决定」
  //判定走 nodeStyle 单一真相源，不在此处另写「谁是圆形」的 if。
  /* ★ 锁定组：主节点上有圆形节点连着时显示按钮（OO 在 OoNodeButtons 行尾，
     非 OO 在上方单钮行）——锁定后整组一起移动 */
  const isRule = isRuleNodeKind(data.kind)
  const isRound = isRoundNodeKind(data.kind)
  /* ★ terminal 节点（终点：pdfmerge/note/output 等）——不渲染出线口 */
  const pmProg = useBaseStore((s0) => s0.pdfmergeProgress[selfId ?? ''])
    const isTerminal = !!(styleOf(data.kind) as { terminal?: boolean } | undefined)?.terminal
  // selfId 可能为 null（节点尚未挂载的合法态），必须先判空再拿它索引/传参
  const lockedByRule = !isRound && selfId != null && isRuleLocked(selfId, allTables)
  /* 指令①（2026-10-05 用户：「点击规则跟随 与该节点跟随 节点连接的节点 的锚点
     在点击选中该节点的时候没显示」）：选中规则节点 → **它的成员节点**的规则口
     （上1/上2/下1/下2）显形。
     判定链：selRuleNode（store 瞬态桥，Canvas.onSelectionChange 写）
       → ruleNodesOf（ruleSync 单一真相源：锁定我的规则节点是谁）
       → 我是不是它的成员。
     ★ 显形走 className（rf-rule-ports-show）而非改 opacity 内联样式：
       与 rf-flowing 同路 —— 瞬态只动一个 class，不碰 nodes/edges 状态，
       也就不会触发重建、不会把别的动画重启（那是踩过的回归根因）。 */
  const selRuleId = useBaseStore((s) => s.selRuleNode)
  const showRulePorts =
    !isRound && selfId != null && selRuleId != null && ruleNodesOf(selfId, allTables).includes(selRuleId)
  /* 全局折叠态（2026-10-06 用户：「每个常规节点都折叠成正方形」）：
     CanvasRibbon 的按钮写 store.nodesFolded，这里只读 —— 常规节点挂
     .rf-node--folded 收成 64px 正方形（样式全在 CSS，与 rf-node--rule 同一套约定：
     组件只挂类名，UI 只读不写样式）。规则节点不折叠 —— 它已是圆形特殊形态。 */
  const nodesFolded = useBaseStore((s) => s.nodesFolded)
  /* ★ 节点组折叠（2026-10-21）：组头 groupFolded=true → 组内成员同样折叠成方块 */
  const inFoldedGroup =
    !isRound && !isGroupNodeKind(data.kind) && selfId != null && isMemberOfFoldedGroup(selfId, allTables)
  const foldedCls = (!isRound && nodesFolded) || inFoldedGroup ? ' rf-node--folded' : '' 
  /* ★ 节点组（2026-10-21）：组节点走专属渲染（透明框 + 标题栏 + 调大小把手），
     不吃常规节点卡的头/字段/锚点/按钮——早退分支在全部 hooks 之后（合法）。 */
  const isGroupN = data.kind === 'group'
  const groupCfg = useBaseStore((s) => (isGroupN && selfId ? s.tables[selfId]?.config : undefined)) as
    | { groupW?: number; groupH?: number; members?: string[] }
    | undefined
  const groupW = groupCfg?.groupW ?? 360
  const groupH = groupCfg?.groupH ?? 240
  /* ★ 组选中「次一级」判定（2026-10-21 用户指令）：组选中后点选了组内成员
     → 组降为次要选中态（虚线淡化），把视觉焦点让给成员，无缝拖拽编辑 */
  const memberSel = useStore((s) =>
    isGroupN && selfId != null
      ? s.nodes.some(
          (nd) => nd.selected && (((groupCfg?.members ?? []) as string[]).includes(nd.id)),
        )
      : false,
  )

    /* 圆形节点（AI/参数库）运行动画（2026-10-06 用户两条：
     1「对选中的节点点击动态特效才会有效果」
      2「增加特效的结束的退出效果 测试的时候3秒结束」）
     - 折叠时圆形节点**不进折叠两行分支**（!isRound 条件）—— 圆内 rf-rule-in
       结构原样不动（「样式内容先不变」），只叠加旋转光环表示「运行中」
     - ★ 判定抽到 core/roundRun.roundRunStateOf（纯函数，有单测）：
       **必须是「选中 + running/exiting」**，折叠**不参与**判定。
       旧实现是 `(nodesFolded || roundRunning)`，点一下播放按钮全画布圆形节点
       一起转圈 —— 那正是用户第1条要修的。
     - 退出态也要挂 rf-round-run（外加 rf-round-exit）：class 一摘元素就没了，
       动画无从播完，「啪」一下消失不算退出效果。
     - rule 节点不挂 —— 用户口径只有 AI/参数库（它有自己的 rf-flowing 表达） */
  const roundRunState = useBaseStore((s) => s.roundRunState)
  const roundRunCls = roundRunClassOf(roundRunStateOf(data.kind, isSel, roundRunState, nodesFolded))
  /* 圆形节点（ai/paramlib）圆内正中的连线计数：AI = 吃了几个上游，
     参数库 = 供了几个下游。数法走 ruleSync.roundLinkCountOf（单一真相源）。
     rule 节点不走这个 —— 它的圆内计数是成员数（memberSummaryOf，含主从语义）。 */
  const roundLinkCount = useMemo(
    () => (isRound && !isRule && selfId != null ? roundLinkCountOf(selfId, allTables) : 0),
    [isRound, isRule, selfId, allTables],
  )
  /* 被连线引用数（2026-10-06 用户：「下方数字 是 有多少 参数 被连线引用多少 两组数字」）。
     口径：我作为数据上游被下游引用的连线条数（edges 里 source===我）。
     走 React Flow 内部 store 的 useStore 取边表 —— edges 由 Canvas 的
     useEdgesState 管，TableNode 拿不到组件 props（nodeTypes 模块级常量），
     节点内取边表只有这条路。selector 返回原始 number 而非新数组：
     zustand 用 Object.is 比较，返回值稳定才不会连环重渲染。 */
  const refCount = useStore((s) =>
    selfId != null ? s.edges.reduce((n, e) => n + (e.source === selfId ? 1 : 0), 0) : 0,
  )
  /* 规则跟随节点圆心的计数（用户：「规则节点上显示几个节点」）。
     走 memberSummaryOf（ruleSync 内核）而非在这里现数一遍 ——
     数法一旦两处就漂移，尤其「不跟随」的成员算不算进来有歧义。 */
  const memberSummary = useMemo(
    () => (isRule && selfId != null ? memberSummaryOf(selfId, allTables) : { count: 0, primaryName: null }),
    [isRule, selfId, allTables],
  )
  /* 圆圈下方要显示的成员名列表（用户：「连接上了，就显示节点」）。
     直接复用 membersOf —— 它已按「主→从→不跟随」排好序（membersOf 内注释：
     主节点必须在第一行）。这里不自己排序，避免两处数法漂移。 */
  const memberList = useMemo(
    () => (isRule && selfId != null ? membersOf(selfId, allTables) : []),
    [isRule, selfId, allTables],
  )
  /* 规则跟随节点的双击：打开成员管理面板（改主次 / 跟随不跟随 / 移出）。
     回调走 data 传入而非 prop —— nodeTypes 是**模块级常量**，
     拿不到组件 prop（除非用 context 多绕一层）。
     这是 React Flow 的官方做法：放data，随节点一起进ReactFlow 状态。 */
  const onRuleDoubleClick =
    isRule && selfId != null && typeof data.onOpenRuleMembers === 'function'
      ? () => (data.onOpenRuleMembers as OpenRuleMembers)(selfId as string)
      : undefined
  /* ★ 2026-10-19 用户指令：数据表节点（测试）双击打开参数编辑弹窗（去掉卡片按钮） */
  const onTestTableDoubleClick =
    data.kind === 'test_table' && selfId != null ? () => setTtDlgOpen(true) : undefined
  if (isGroupN) {
    return (
      <div
        className={`rf-node nk--group rf-groupbox${isSel ? ' rf-node-sel' : ''}${isSel && memberSel ? ' rf-groupbox--secondary' : ''}`}
        onDoubleClick={(e) => e.stopPropagation()}
      >
        {selfId != null ? <GroupNodeView nodeId={selfId} name={data.name} /> : null}
        {data.groupHint && (
          <div className="group-node-hintmask">
            <span>置入节点至节点组</span>
          </div>
        )}
      </div>
    )
  }

  return (
    <div
      /* rf-flowing：规则节点被选中时打上，**由 CSS 用它驱动连线的流动动画**
         （`.rf-flowing .edge-rule--flow path { animation-play-state: running }`）。
         为什么走 className 而不是把「是否流动」塞进 edges 的 className：
         那样点选会触发 setEdges 全表重建 → RF 重挂所有 <path> →
         所有线的动画**同步重启**，肉眼看成「其他工作流全停」。
         这个类只改节点自己一个属性，边一个字节都不动。 */
      className={`rf-node ${familyClassOf(data.kind)}${isSel ? ' rf-node-sel' : ''}${isRound && isSel ? ' rf-flowing' : ''}${data.running ? ' rf-node--running' : ''}${data.locked ? ' rf-node--locked' : ''}${hoverGlow ? ' rf-node-hoverglow' : ''}${isRound ? ' rf-node--rule' : ''}${isRound && !isRule ? ` rf-node--${data.kind}` : ''}${lockedByRule ? ' rf-node--rule-locked' : ''}${showRulePorts ? ' rf-rule-ports-show' : ''}${foldedCls}${roundRunCls}${ooDropHi ? ' rf-node--oover' : ''}`}
      onDragEnter={(e) => {
        if (!isOOFileNode || !e.dataTransfer.types.includes('Files')) return
        e.preventDefault(); e.stopPropagation()
        ooDropDepth.current++; setOoDropHi(true)
      }}
      onDragOver={(e) => {
        if (!isOOFileNode || !e.dataTransfer.types.includes('Files')) return
        e.preventDefault(); e.stopPropagation()
        try { e.dataTransfer.dropEffect = 'copy' } catch { /* ignore */ }
      }}
      onDragLeave={() => {
        if (!isOOFileNode) return
        ooDropDepth.current--; if (ooDropDepth.current <= 0) { ooDropDepth.current = 0; setOoDropHi(false) }
      }}
      onDrop={async (e) => {
        if (!isOOFileNode || !selfId) return
        const dt = e.dataTransfer
        if (!dt.files?.length) return
        e.preventDefault(); e.stopPropagation()
        ooDropDepth.current = 0; setOoDropHi(false)
        const accept = data.kind === 'ooword' ? '.docx' : '.xlsx'
        const hit = Array.from(dt.files).filter((f) => f.name.toLowerCase().endsWith(accept))
        if (!hit.length) {
          pushLog(`拖入文件与节点类型不匹配（${data.kind === 'ooword' ? 'OO 文档接受 .docx' : 'OO 工作簿接受 .xlsx'}）——已忽略`, 'warn')
          return
        }
        const f = hit[0]
        const st = useBaseStore.getState()
        const cfg = (st.tables[selfId]?.config ?? {}) as { ooBytes64?: string; ooBytesDbId?: string; ooTemplateId?: string }
        const hasContent = !!cfg.ooBytes64 || !!cfg.ooBytesDbId || !!cfg.ooTemplateId
        if (hasContent && !window.confirm(`「${f.name}」将替换「${data.name}」的当前内容，是否继续？`)) return
        let buf: Uint8Array<ArrayBufferLike> = new Uint8Array(await f.arrayBuffer())
        if (data.kind === 'oosheet') {
          try { buf = await stripFullCalcFlag(buf) } catch { /* 摘除失败不影响导入 */ }
        }
        st.pushHistory()
        st.setNodeConfig(selfId, {
          ooTemplateId: 'tpl-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7),
          ooTemplateName: f.name,
          ooBytesDbId: await stashOoBytes('oobytes-' + selfId, buf),
        })
        st.renameTable(selfId, f.name.replace(/\.[^.]+$/, ''))
        pushLog(`拖放导入：${f.name} → ${data.name}${hasContent ? '（已替换原内容）' : ''}`, 'ok')
      }}
      onDoubleClick={onRuleDoubleClick ?? onTestTableDoubleClick ?? onParamlibDoubleClick}
      title={
        isRule
          ? `${data.name} · 规则跟随（双击管理成员）`
          : isRound
            ? `${data.name} · ${styleOf(data.kind).label}（连线 ${roundLinkCount} 条）`
            : nodesFolded
              ? `${data.name}（${data.fields.length} 参数 · 被连线引用 ${refCount} 条）`
              : undefined
      }
    >
      {/* ── 数据流连接口：左进右出（2026-10-05 恢复原位）────────────────
         * ★★ 恢复记录（这个方向别再搞反）：
         *   用户第 3 条说「节点连线位置位于每个节点的顶部，而不是左侧」，
         *   我把**数据流口**也从 Left/Right 搬到了 Top —— **理解错了对象**。
         *   用户澄清：「我要的不是节点左右的锚点隐藏。我说的是节点上下的锚点，
         *   节点左右的锚点恢复成老样子」。
         *   → 「上下」说的是**规则跟随口**（普通节点上方 1/3、2/3 处，
         *     规则节点下方出来），它本来就在上/下，从头到尾没动过；
         *   → 「左右」是**数据流口**（target 左、source 右），要恢复原样。
         *   实测确认改动前就是 Left/Right（git show 8efbc85^ → Left + Right）。
         *
         * 两类口分处不同边，天然不重叠，不必再算百分比错开：
         *   规则口：上沿 33% / 66%（方形青绿，规则线专用）
         *   数据流口：左沿 target / 右沿 source（圆形灰，数据流专用） */}
      {/* ★ 运行状态徽章（2026-10-23 用户指令「节点上方显示关键过程 参照其他画布软件」）：
          n8n/Dify 风格——运行中 spinner / 成功 ✓ 耗时 / 失败 ✗ 摘要，
          悬浮节点顶缘中央，非持久内存态（core/nodeRunStatus） */}
      {selfId != null && <RunBadge id={selfId} />}
      {!isRound && (
        <>
          {/* ★ 2026-10-23 用户指令：PDF 合并节点的入口锚点对齐标题名字中间
             （head 30px = 6px padding×2 + 18px 行高 → 文字中心 15px） */}
          {/* ★ 2026-10-25 用户指令「只要pdf节点连接到pdf替换页节点…建立连接 文件名也同步落在pdf替换页节点」：
              pdfreplace 恢复左1 数据口（此前「左边的中间节点不要了」作废）——
              连到节点本体=上游 PDF 首份自动集成到被替换组 + 建线；
              精确落组仍走组级/词条级锚点（pr-add-/pr-upper-） */}
          <Handle type="target" position={Position.Left} className="rf-handle rf-handle--flow rf-port-no-l1" style={data.kind === 'pdfmerge' || data.kind === 'pdfreplace' ? { top: 15 } : undefined} title={data.kind === 'pdfreplace' ? '左1 · 连到节点=上游 PDF 首份自动落入被替换组（精确落组用词条/组锚点）' : '左1 · 数据流入口（上游 → 本节点）'} />
          {/* ★ 2026-10-25 用户指令「去掉中间的锚点 只留 PDF 右侧的 和每个词条后面的」：
              pdfmerge 是终点节点——自身数据流输出口不再渲染（产物引用走金锚点/行级锚点） */}
          {data.kind !== 'pdfmerge' && (
            <Handle type="source" position={Position.Right} className="rf-handle rf-handle--flow rf-port-no-r1" title="右1 · 数据流出口（本节点 → 下游）" />
          )}
        </>
      )}
      {/*
        ★ 顶部两个「规则跟随」专用连接口（2026-10-05，用户口径：
        「每个节点的节点在上方增加一个方便规则跟随节点连接」
          +「上方新增两个，三分之一的位置和三分之二的位置」
          +「规则线只能连接常规节点顶部的点」）

        为什么两个而不是一个：主节点与从节点要在**画布上一眼可分**——
        主节点（本组里可改规则的那个）用 1/3 位，从节点用 2/3 位。
        只有一个口的话主/从只能打开面板才知道，视觉上区分不出来。

        ★★为什么必须给这两个口挂 id（RULE_PORT_IN），而不能只靠 type区分：
          React Flow Strict 模式**只禁同类型对同类型**（源码 @xyflow/system
          isValidHandle：`connectionMode === Strict ? fromType !== handleType : ...`），
          异类型一律放行。这两个口是 target，和**左侧数据流口**（也是 target）
          在React Flow 眼里毫无区别—— 于是规则节点下方的口能连到左侧口，
          App.handleConnect 会把它当数据流线写进 inputs/linkedInputs，
          症状是「规则节点变成输出的上游」，且毫无报错。
          挂上专属 id 后，由canConnectPorts 在 UI 层把两类口隔开。
          （此前 587 个用例全绿却漏掉这个，正是因为测试只调ruleConnectPatch
            这个纯函数，从没验证过「哪个口挨着哪个口」。）
      */}
      {!isRound && (
        <>
          <Handle
            id={RULE_PORT_IN}
            type="target"
            position={Position.Top}
            className="rf-handle rf-handle--rule rf-port-no-t1"
            style={{ left: '33%' }}
            title="上1 · 规则跟随口（主节点位）"
          />
          <Handle
            /* 上2 用**专属 id**（RULE_PORT_IN_2），不再与上1 共用：
               同节点内同 id 时 RF 取第一个匹配，连线永远落上1、上2 成摆设。
               2026-10-05 用户：「统一修改上1的锚点 改为上2位置」——
               画布连线由此显式指 RULE_PORT_IN_2（见 edgesOf 的 targetHandle）。 */
            id={RULE_PORT_IN_2}
            type="target"
            position={Position.Top}
            className="rf-handle rf-handle--rule rf-port-no-t2"
            style={{ left: '66%' }}
            title="上2 · 规则跟随口（从节点位 · 规则线统一落点）"
          />
        </>
      )}
      {/* ── 常规节点底部的 下1/下2（2026-10-05 用户：「增加节点 下1 下2 的锚点」）──
          与上1/上2 对称：33% / 66% 两个位，同为规则口（isRulePortId 认）。
          · 下1 复用 RULE_PORT_OUT（与规则节点出线口同 id —— RF 的 handle
            按**节点内**查找，不同节点同 id 互不干扰）；
          · 下2 挂专属 RULE_PORT_OUT_2（同节点内同 id 会被 RF 取首）。
          ★ 口角色都是 source（出线），与上1/上2 的 target（落点）配对：
            规则口 source ↔ 规则口 target 才是合法规则线（canConnectPorts）。
          ★ 常规节点下1 → 常规节点上1 这类「两端规则口、两端都非规则节点」的
            连线由 canConnectPorts 新防线拒绝（见 ruleSync 注释），不会劈叉成数据流线。
          显隐：走 rf-handle--rule 体系 —— 常态隐藏、拖线中显形、
          选中规则节点时成员节点显形（rf-rule-ports-show）。 */}
      {!isRound && (
        <>
          {/* ★ 2026-10-25 用户指令「下面的锚点去掉」：pdfmerge 是终点节点（只进不出），
              底部两个规则出线口对它无意义——不再渲染（也避免误连跳变） */}
          {!isTerminal && (
            <>
              <Handle
                id={RULE_PORT_OUT}
                type="source"
                position={Position.Bottom}
                className="rf-handle rf-handle--rule rf-port-no-b1"
                style={{ left: '33%' }}
                title="下1 · 规则跟随出线口（主节点位）"
              />
              <Handle
                id={RULE_PORT_OUT_2}
                type="source"
                position={Position.Bottom}
                className="rf-handle rf-handle--rule rf-port-no-b2"
                style={{ left: '66%' }}
                title="下2 · 规则跟随出线口（从节点位）"
              />
            </>
          )}
        </>
      )}
      {isRule && selfId != null && memberList.length > 0 && (
        /* ★★ 连上就显示节点名（用户：「连接上了，就显示节点」）
         *
         * 为什么不直接把名字塞进圆里：圆圈直径被用户要求压到 64px
         * （「规则跟随节点 的节点大小太大了 变小点」），圆内只放得下计数。
         * 所以成员名做成**挂在圆圈上方的标签列**——
         *   ① 不动圆圈尺寸，守住上一条口径；
         *   ② 不占圆内空间，计数仍然一眼可见；
         *   ③ .rf-node 是 overflow:visible，标签溢出不被裁切。
         *
         * ★★ 2026-10-05 用户第 1 条：「放在上面（现在是下面）
         *   当我点击规则跟随节点才浮动出来」
         *   → 两处改动：位置由下改上；**默认隐藏，点中该规则节点才显形**。
         *   显隐挂在 `.rf-node-sel` 上（CSS 侧控制，不新增 React state）：
         *   选中态本来就算在 isSel 里，不新增状态源，
         *   也就不存在「state 与 class 不同步」这类漂移。
         *   常态隐藏的另一个好处：连了 10 个成员的规则节点不会在画布上
         *   长期糊着一列标签，视觉噪音降到选中时才付。
         *
         * 标签内容走 memberList（ruleSync 派生，非此处现数一遍）——
         * 数法一旦两处就漂移（同 RuleMembersPanel 的道理）。
         * 只显示前 MAX 条，多的折成「+N」：连 20 个节点时把画布糊满，
         * 反而看不出这是哪个规则节点。 */
        <div className="rf-rule-tags" title={memberList.map((m) => allTables[m.targetId]?.name ?? m.targetId).join('\n')}>
          {memberList.slice(0, RULE_TAGS_MAX).map((m) => (
            <span key={m.targetId} className={`rf-rule-tag rf-rule-tag--${m.mode}`}>
              <i className="rf-rule-tag-dot" />
              {allTables[m.targetId]?.name ?? m.targetId}
            </span>
          ))}
          {memberList.length > RULE_TAGS_MAX && (
            <span className="rf-rule-tag rf-rule-tag--more">+{memberList.length - RULE_TAGS_MAX}</span>
          )}
        </div>
      )}
      <div className="rf-node-inner">
        {/* 折叠态：不渲染「头+体」结构，改两行。
            上行 logo（2026-10-06 二轮，用户：「不要只写一个字 参数表就原来大小的
            底部圆角矩形颜色填充尺寸 logo字体也一致」）—— rf-fold-logo-row 只是
            居中行盒，里面的徽章与展开态**同一个类串**（rf-badge + badgeClassOf），
            圆角/填充/尺寸/字体零漂移。
            下行两组数字（2026-10-06 二修：「X参 X引 改成 X/X 引用数量的颜色
            要鲜艳一点」）—— 「参数数/引用数」斜杠分隔，引用数用规则青绿。
            节点全名挪到 wrapper title，悬停可见。 */}
        {/* ★ 条件是 !isRound 不是 !isRule（2026-10-06 用户：「圆形节点折叠的时候
             AI 和 参数库 折叠的时候 样式内容先不变」）—— AI/参数库是圆形，
             全局折叠时圆内 rf-rule-in 结构原样保持，不换成这两行折叠结构；
             它们的「运行中」表达走 rf-round-run 旋转光环（见 roundRunCls）。
             （表达式括号内不能放花括号注释 —— #107 同款坑。） */}
        {!isRound && nodesFolded ? (
          <>
            <span className="rf-fold-logo-row">
              <span className={`rf-badge ${badgeClassOf(data.kind)}`}>{styleOf(data.kind).label}</span>
            </span>
            <span
              className="rf-fold-count"
              title={`参数 ${data.fields.length} 个 · 被连线引用 ${refCount} 条（作为数据上游）`}
            >
              <b>{data.fields.length}</b>
              <em>/</em>
              <b className="rf-fold-ref">{refCount}</b>
            </span>
          </>
        ) : (
          <>
            <div className="rf-node-head" style={(data.kind === 'pdfmerge' && selfId != null && (pmProg ?? 0) > 0) ? { backgroundImage: `linear-gradient(90deg, rgba(59, 130, 246, 0.5) ${pmProg}%, transparent ${pmProg}%)` } : undefined}>

          {isRound && selfId != null ? (
            /* 圆内两行（用户第 3 条：「数字居中缩小往下移动，上面及以上规则跟随 四个字」）：
               上行小字交代这是什么节点，下行才是计数。
               改前只有孤零零一个大数字，用户看不出这个圆是什么；
               而「规则跟随」四字与22px 数字硬塞在一行会互相压——
               所以拆成两行、数字缩到 13px 并整体下移。
               2026-10-06 起 AI/参数库模仿同一结构（用户：「模仿规则跟随节点样式」）：
               上行 = 各自 label（'AI' / '参数库'），下行 = 连线数
               （roundLinkCountOf：AI 数上游、参数库数下游 —— 数法见 ruleSync 注释）。
               rule 的计数是成员数（memberSummaryOf，含主从语义），两套口径不混用。 */
            <span className="rf-rule-in" title={isRule ? `${data.name}（双击管理成员）` : isParamlib ? `${data.name} · 参数库（双击选择参数库）` : `${data.name} · ${styleOf(data.kind).label}（连线 ${roundLinkCount} 条）`}>
              <span className="rf-rule-in-label">{isRule ? '规则跟随' : styleOf(data.kind).label}</span>
              {/* ★ 2026-10-23 用户指令「双击圆形参数库 打开参数库选择」：圆内显示已选库名 */}
              {isParamlib && (() => {
                const libName = ((allTables[selfId ?? '']?.config ?? {}) as { libName?: string }).libName
                return libName ? <span className="rf-rule-lib">{libName}</span> : null
              })()}
              <span className="rf-rule-count">
                {isRule ? (memberSummary.count > 0 ? memberSummary.count : '—') : roundLinkCount > 0 ? roundLinkCount : '—'}
              </span>
            </span>
          ) : (
            <>
              <span className={`rf-badge ${badgeClassOf(data.kind)}`}>{styleOf(data.kind).label}</span>
              <span className="rf-node-title" title={data.name}>{data.name}</span>
              {/* ★ 2026-10-23 用户指令「数量做成和节点组中节点数量一样的表达样式」：
                  photo/pdf 标题栏右侧幽灵数字（同 .group-node-count 22×22 幽灵钮语言）；
                  卡片内徽章行已删（与改名后的标题重复） */}
              {(() => {
                if (!selfId) return null
                const k = data.kind
                /* ★ 2026-10-23 用户指令「PDF合并按节点组样式 实时更新」：pdfmerge 计数 =
                   当前上游份数（config.inputs），allTables 订阅 → 连线/断开/删除即时刷新 */
                if (k === 'pdfmerge') {
                  const c = allTables[selfId]?.config as { inputs?: string[] } | undefined
                  const n = (c?.inputs ?? []).length
                  if (!n) return null
                  return <span className="rf-node-count" title={`${n} 份上游（实时）`}>{n}</span>
                }
                if (k === 'fileout') {
                  const ls = ((allTables[selfId]?.config as { lastSaved?: unknown[] } | undefined)?.lastSaved ?? [])
                  if (!ls.length) return null
                  return <span className="rf-node-count" title={`已保存 ${ls.length} 次`}>{ls.length}</span>
                }
                if (k !== 'photo' && k !== 'pdf' && k !== 'pdfslim' && k !== 'pdf2word') return null
                const c = allTables[selfId]?.config as { photoItems?: unknown[]; pdfItems?: unknown[]; p2wItems?: unknown[] } | undefined
                const n = k === 'photo' ? (c?.photoItems ?? []).length : k === 'pdf2word' ? (c?.p2wItems ?? []).length : (c?.pdfItems ?? []).length
                if (!n) return null
                return <span className="rf-node-count" title={k === 'photo' ? `${n} 张照片` : k === 'pdf2word' ? `${n} 份待转 PDF` : `${n} 份 PDF`}>{n}</span>
              })()}
              {/* ★ 统一锁标识（2026-10-26）：锁定节点头部小锁标——一眼看出该节点已锁 */}
              {data.locked && (
                <span className="rf-node-lockmark" title="已锁定（数据/位置）——右键节点可解锁">
                  <IcLock size={11} />
                </span>
              )}
            </>
          )}
        </div>
        <div className="rf-node-body">
          {data.kind === 'note' ? (
            <>
              <NoteStickyBody nodeId={selfId ?? ''} value={data.textValue} />
              {/* ★ 锤子便签式：右缘宽度把手（高度随内容自动，无需拖） */}
              {selfId ? (
                <div
                  className="sticky-width-handle"
                  title="拖动调整宽度"
                  onPointerDown={(e) => {
                    e.stopPropagation()
                    const el = e.currentTarget as HTMLElement
                    el.setPointerCapture(e.pointerId)
                    const st = useBaseStore.getState()
                    const startX = e.clientX
                    const startW = ((st.tables[selfId ?? '']?.config as { noteW?: number } | undefined)?.noteW ?? 190)
                    const rect = (el.closest('.group-node, .rf-node') as HTMLElement | null)?.getBoundingClientRect()
                    const scale = rect && startW ? rect.width / startW : 1
                    const onMove = (ev: PointerEvent) => {
                      const newW = Math.round(Math.min(560, Math.max(120, startW + (ev.clientX - startX) / (scale || 1))))
                      st.setNodeConfig(selfId, { noteW: newW })
                    }
                    const onUp = () => {
                      el.removeEventListener('pointermove', onMove)
                      el.removeEventListener('pointerup', onUp)
                      st.pushHistory()
                    }
                    el.addEventListener('pointermove', onMove)
                    el.addEventListener('pointerup', onUp)
                  }}
                />
              ) : null}
            </>
          ) : data.kind === 'text' ? (
            <div
              className="rf-field rf-textnode"
              title="双击编辑内容"
              /*胶囊拖放靶：面板里的参数胶囊拖进来时，落点由指针在这个元素里的横向位置决定 */
              data-cap-drop={selfId}
            >
              <span
                className="rf-textnode-val"
                style={
                  data.textStyle
                    ? {
                        display: 'block',
                        width: '100%',
                        fontWeight: data.textStyle.bold ? 700 : undefined,
                        fontStyle: data.textStyle.italic ? 'italic' : undefined,
                        textDecoration: data.textStyle.underline ? 'underline' : undefined,
                        fontSize: data.textStyle.fontSize ? `${data.textStyle.fontSize}px` : undefined,
                        color: data.textStyle.color || undefined,
                        textAlign: (data.textStyle.align as React.CSSProperties['textAlign']) || undefined,
                      }
                    : undefined
                }
              >
                <InlineCapsules
                  value={data.textValue}
                  nodeNameOf={(id) => allTables[id]?.name}
                  deadOf={(id) => !allTables[id]}
                />
              </span>
            </div>
          ) : (
            <>
              {/* ── OO 文档节点三按钮（2026-10-07 用户指令）：
                  「放在标题下方第一行靠左位置」—— 按钮条为 body 首行，左对齐。
                  预加载（后台打开+冻结，点开 0 等待）/ 数据锁定（只读）/ 位置锁定（不可拖）── */}
              {/* ★ 2026-10-25 OO 节点词条化（用户指令「仿照 excel 节点 相同尺寸 词条样式」）：
                  功能钮移入词条行尾（24×24 紧凑内联），行样式与 PDF 节点词条同款 */}
              {/* ★ PDF 合并节点（2026-10-22）：上游 OO 文档按序转 PDF 拼接。
                  2026-10-23 用户澄清「原来的不变 按钮在下面；点历史 历史记录在
                  按钮的下面打开」——DOM 顺序固定：卡片内容 → 按钮行 → 历史列表 */}
              {data.kind === 'pdfmerge' && selfId && (
                <PdfMergeButtons nodeId={selfId} lod={!!(data as { lodDragging?: boolean }).lodDragging} />
              )}
              {/* ★ 照片/照片组节点（2026-10-23 用户指令）：承载照片连入 PDF 合并 */}
              {data.kind === 'photo' && selfId && (
                <PhotoNodeCard nodeId={selfId} lod={!!(data as { lodDragging?: boolean }).lodDragging} />
              )}
              {/* ★ PDF 替换页节点（2026-10-25 用户指令「上下两组 执行替换 有历史」） */}
              {data.kind === 'pdfreplace' && selfId && (
                <PdfReplaceNodeCard nodeId={selfId} />
              )}
              {/* ★ 文件夹节点（2026-10-25 用户指令「作为读取位置和保存位置」） */}
              {data.kind === 'folder' && selfId && (
                <FolderNodeCard nodeId={selfId} />
              )}
              {/* ★ PDF 文件节点（2026-10-23 用户指令）：承载 PDF 文件组连入 PDF 合并 */}
              {/* ★ PDF 瘦身节点（2026-10-26）：结构同 PDF 文件节点（pdfItems/pdfFileOrder），复用同一卡片 */}
              {(data.kind === 'pdf' || data.kind === 'pdfslim') && selfId && (
                <PdfNodeCard nodeId={selfId} lod={!!(data as { lodDragging?: boolean }).lodDragging} />
              )}
              {/* ★ 文件输出节点（2026-10-26 用户指令「输出节点 可以存储文件到固定位置」）：
                  连线上游产物 → 一键保存到固定目录（outDir 随画布走） */}
              {data.kind === 'fileout' && selfId && (
                <FileOutNodeCard nodeId={selfId} />
              )}
              {/* ★ 网址监控（2026-11-09 用户指令「增加一个网址监控节点…提取、记录、或者输出数据」） */}
              {data.kind === 'webmon' && selfId && (
                <WebmonNodeCard nodeId={selfId} />
              )}
              {/* ★ Excel转DXF（2026-11-10 用户指令「增加一个 excel转dxf 节点」） */}
              {data.kind === 'xlsx2dxf' && selfId && (
                <Xlsx2dxfNodeCard nodeId={selfId} />
              )}
              {/* ★ PDF转Word 节点（2026-10-26 用户指令「利用百度的开源」）：
                  词条 + 上游行 + 转换按钮行（共享行）+ 历史列表 */}
              {data.kind === 'pdf2word' && selfId && (
                <Pdf2WordNodeCard nodeId={selfId} lod={!!(data as { lodDragging?: boolean }).lodDragging} />
              )}
              {/* ★ 非OO常规节点按钮行（保持节点底部——「原来的不变」） */}
              {!isRound && selfId && data.kind !== 'ooword' && data.kind !== 'oosheet' && data.kind !== 'fileout' && (
                <div className="rf-field rf-oo-btns" style={{ display: 'flex', gap: 4, alignItems: 'center', justifyContent: data.kind === 'pdfreplace' ? 'flex-start' : undefined }}>
                  {/* ★ 统一节点锁（2026-10-26）：数据锁+位置锁——全节点类型同款控件 */}
                  <NodeLockButtons nodeId={selfId} />
                  <GrpLockButton nodeId={selfId} />
                  {/* ★ 2026-10-23 用户指令「做成一行 都是正方形的圆角的」：
                      锁定组+存画布+历史+合并 四钮一行（31×31 方形圆角同语言） */}
                  {data.kind === 'pdfmerge' && <PdfExtraButtons nodeId={selfId} />}
                  {data.kind === 'pdfmerge' && <PdfMergeRunButton nodeId={selfId} />}
                  {/* ★ 2026-10-23 用户指令：照片导入按钮移到首行锁定组旁（同尺寸长方形） */}
                  {data.kind === 'photo' && <PhotoImportButton nodeId={selfId} />}
                  {/* ★ 2026-10-23 PDF 文件节点同款：导入按钮与锁定组并排 */}
                  {(data.kind === 'pdf' || data.kind === 'pdfslim') && <PdfImportButton nodeId={selfId} />}
                  {/* ★ 2026-10-29 用户指令「连接上游后出现生成按钮」：⟳生成 + 🕘历史 进方形按钮行
                      （31×31 与锁定组/合并钮同语言；生成钮有上游才显示，历史钮有历史才显示） */}
                  {data.kind === 'pdf' && <PdfGenButton nodeId={selfId} />}
                  {data.kind === 'pdf' && <PdfGenHistButton nodeId={selfId} />}
                  {/* ★ 2026-10-25 PDF 替换页：执行按钮与锁定组并排（用户指令「放在最下面和那个锁并排」） */}
                  {data.kind === 'pdfreplace' && <PdfReplaceHistButton nodeId={selfId} />}
                  {data.kind === 'pdfreplace' && <PdfReplaceRunButton nodeId={selfId} />}
                  {/* ★ PDF转Word（2026-10-26）：导入 → 引擎切换 → 历史 → 执行 */}
                  {data.kind === 'pdf2word' && <Pdf2WordImportButton nodeId={selfId} />}
                  {data.kind === 'pdf2word' && <Pdf2WordEngineButton nodeId={selfId} />}
                  {data.kind === 'pdf2word' && <Pdf2WordHistButton nodeId={selfId} />}
                  {data.kind === 'pdf2word' && <Pdf2WordRunButton nodeId={selfId} />}
                </div>
              )}
              {/* ★ 2026-10-23 用户指令「点击历史 历史记录在按钮的下面打开」：
                  历史列表渲染在按钮行之后——🕘 展开即出现在按钮正下方 */}
              {data.kind === 'pdfmerge' && selfId && (
                <PdfHistoryList nodeId={selfId} />
              )}
              {/* ★ 2026-10-25 PDF 替换页历史：🕘 点击展开/收起（同 PDF 合并历史位置） */}
              {data.kind === 'pdfreplace' && selfId && (
                <PdfReplaceHistList nodeId={selfId} />
              )}
              {/* ★ 2026-10-26 PDF转Word 历史：🕘 展开在按钮行正下方（产物 Word 列表） */}
              {data.kind === 'pdf2word' && selfId && (
                <Pdf2WordHistoryList nodeId={selfId} />
              )}
              {/* ★ 2026-10-29 PDF 文件节点生成历史：🕘 展开在按钮行正下方（旧内容 PDF 列表） */}
              {data.kind === 'pdf' && selfId && (
                <PdfGenHistoryList nodeId={selfId} />
              )}
              {data.fields.map((f) => (
                <div key={f.id} className="rf-field">
                  <span>{f.name}</span>
                  <span className="rf-type">{f.type}</span>
                </div>
              ))}
              {/* ★ 2026-10-25 OO 节点词条化（用户指令「仿照 PDF 词条 logo 用合并PDF对应的logo」）：
                  词条行 = 序号 + pm-fmt-word/sheet 方块（W蓝/S绿 同 P 红规格）+ 文件名；
                  功能钮恢复独立一行（31×31，按钮不混词条） */}
              {(data.kind === 'ooword' || data.kind === 'oosheet') && selfId && (
                <div className="pn-row oo-entry nodrag nopan" title="OO 文档（双击进入编辑）">
                  <span className="pn-idx">1</span>
                  <span className={'pm-fmt ' + (data.kind === 'oosheet' ? 'pm-fmt-sheet' : 'pm-fmt-word')}>{data.kind === 'oosheet' ? 'X' : 'W'}</span>
                  <span className="pn-nm" style={{ position: 'static', background: 'none', padding: 0, fontSize: 12, color: 'var(--c-text2)' }}>{data.boundHint || (data.kind === 'ooword' ? '未绑定文档' : '未绑定工作簿')}</span>
                </div>
              )}
              {(data.kind === 'ooword' || data.kind === 'oosheet') && selfId && (
                <OoNodeButtons nodeId={selfId} kind={data.kind} />
              )}
              {(EDITABLE_KINDS as string[]).includes(data.kind) && data.kind !== 'ooword' && data.kind !== 'oosheet' && (
                <div className="rf-field rf-field--hint" title="双击节点进入编辑">
                  <span>双击进入编辑</span>
                  <span className="rf-type">{data.boundHint || '未绑定'}</span>
                </div>
              )}
            </>
          )}
          </div>
        </>
        )}
      </div>
      {/* ★ 2026-10-23 用户指令「双击参数库 打开窗体 选择 参数库内容」：
          居中大窗体（AG Grid 全量虚拟滚动）——选库 + 单击行选中参数内容，
          写 config.libName / libSelectedRow，圆内显示库名 */}
      {libPickOpen && selfId && (
        <ParamlibContentDialog nodeId={selfId} onClose={() => setLibPickOpen(false)} />
      )}
      {/* 右侧数据流口已移除（2026-10-05 用户第 3 条：数据流口移到顶部）。
          规则节点仍只在下方有一个口（rule-out），规则线不从顶部出入。*/}
      {isRule && (
        <Handle
          /* ★ 下方这个口挂 RULE_PORT_OUT —— 与普通节点顶部的 RULE_PORT_IN 配对，
             两者同为「规则口」才允许相连（canConnectPorts）。
             它绝不与数据流口混用：普通节点的左右口没有 id，isRulePortId 判 false。 */
          id={RULE_PORT_OUT}
          type="source"
          position={Position.Bottom}
          className="rf-handle rf-handle--rule-out rf-port-no-b1"
          title="下1 · 规则跟随线出线（连到要跟随的节点）"
        />
      )}
      {/* ── AI 节点（圆形）的接收口（2026-10-06，用户：「AI节点的连线放在
          普通节点的 下1锚点位置」）──
          方向：普通节点下1（RULE_PORT_OUT，source）→ 本口（target，AI 吃数据）。
          口挂在 AI 圆**上沿**：线从普通节点底部出、落 AI 顶部，贝塞尔最顺。
          类挂 rf-handle--rule-out 复用「常显 + 尺寸」体系（圆形节点唯一的口
          必须常显，否则无从起手拖线），族色差异由 rf-handle--ai 覆盖（CSS）。
          准入：canConnectPorts 的 triggerLineSemantics 矩阵（'ai' 组合）。 */}
      {data.kind === 'ai' && (
        <Handle
          id={AI_PORT_IN}
          type="target"
          position={Position.Top}
          className="rf-handle rf-handle--rule-out rf-handle--ai"
          title="上 · AI 接收口（普通节点下1 出线连到这里）"
        />
      )}
      {/* ── 参数库节点（圆形）的出线口（2026-10-06，用户：「参数库连接节点 锚点 上1」）──
          方向：本口（source，参数库供数）→ 普通节点上1（RULE_PORT_IN）。
          口挂参数库圆**下沿** —— 与规则节点出线口同位，完全模仿其形态。
          准入：triggerLineSemantics 矩阵（'paramlib' 组合）。 */}
      {data.kind === 'paramlib' && (
        <Handle
          id={PARAMLIB_PORT_OUT}
          type="source"
          position={Position.Bottom}
          className="rf-handle rf-handle--rule-out rf-handle--paramlib"
          title="下1 · 参数库出线（连到普通节点的上1口）"
        />
      )}
      {/* ★ 数据表节点（测试）：参数编辑弹窗（Portal 到 body，避免节点卡片 overflow 裁剪） */}
      {ttDlgOpen && selfId && (
        <TestTableDialog nodeId={selfId} onClose={() => setTtDlgOpen(false)} />
      )}
    </div>
  )
}

function AssocEdge(props: EdgeProps) {
  const { id, sourceX, sourceY, targetX, targetY, sourcePosition, targetPosition, data } = props
  const [path, labelX, labelY] = getBezierPath({
    sourceX,
    sourceY,
    targetX,
    targetY,
    sourcePosition,
    targetPosition,
  })
  const d = data as { onRemove?: (id: string) => void; label?: string; hovered?: boolean } | undefined
  const onRemove = d?.onRemove
  const label = d?.label
  // 删除按钮默认隐藏：仅当鼠标悬停在线条上（由 Canvas 通过 data.hovered 传入）或悬停在按钮本身时才显示；
  // 单击后弹出确认（确认逻辑在 App.handleEdgeRemove，按连线类型给名字）才真正删除。
  const [btnHover, setBtnHover] = useState(false)
  const show = !!d?.hovered || btnHover
  const onBtnClick = (e: React.MouseEvent) => {
    e.stopPropagation()
    onRemove?.(id)
  }
  return (
    <>
      {/* ★ 效果层：pathLength=1 归一化轨道（archify 三招——效果选择器激活时亮点沿真实路径扫过，
         长短边速度/比例一致；默认透明不影响原连线） */}
      <path d={path} pathLength={1} className="edge-fx-track" fill="none" />
      <BaseEdge id={id} path={path} />
      <EdgeLabelRenderer>
        <div
          data-edge-id={id}
          style={{
            position: 'absolute',
            transform: `translate(-50%, -50%) translate(${labelX}px, ${labelY}px)`,
            // 容器本身不拦截指针事件，否则会挡住下方线条、导致 onEdgeMouseEnter 收不到；
            // 仅「删除按钮」在显示时（show）才接管点击/悬停。
            pointerEvents: 'none',
            display: 'flex',
            gap: 4,
            alignItems: 'center',
          }}
        >
          <button
            className="edge-rm"
            style={{ opacity: show ? 1 : 0, pointerEvents: show ? 'all' : 'none' }}
            onMouseEnter={() => setBtnHover(true)}
            onMouseLeave={() => setBtnHover(false)}
            onClick={onBtnClick}
            title="删除该连线"
          >
            ×
          </button>
          {label && <span className="edge-tag" style={{ pointerEvents: 'none' }}>{label}</span>}
        </div>
      </EdgeLabelRenderer>
    </>
  )
}

/* ★ 2026-10-23 拖动性能优化：TableNode memo 化——拖动时未拖节点 data 引用不变，
   memo 直接跳过重渲（87 节点画布拖动卡顿的大头） */
const TableNode = memo(TableNodeInner)
const nodeTypes = { table: TableNode } satisfies NodeTypes
const edgeTypes = { assoc: AssocEdge } satisfies EdgeTypes
const fitViewOptions = { padding: 0.2, maxZoom: 1 }

export interface CopyItem {
  id: string
  table: TableSchema
  record: Row[]
  pos?: NodePos
}

// 缩放控制面板：固定在右下角、缩略图上方，悬停显示更清晰，含当前缩放比例
function ZoomPanel() {
  const { zoomIn, zoomOut, fitView, zoomTo, updateNode } = useReactFlow()
  const { zoom } = useViewport()
  return (
    <Panel position="bottom-right" className="zoom-panel" style={{ bottom: 170 }}>
      <button
        className="zoom-btn"
        onClick={() => zoomOut({ duration: 200 })}
        title="缩小"
        aria-label="缩小"
      >
        <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true">
          <path d="M5 11h14v2H5z" fill="currentColor" />
        </svg>
      </button>
      <span
        className="zoom-val"
        title="当前缩放比例 · 双击恢复 100%"
        onDoubleClick={() => zoomTo(1, { duration: 200 })}
      >
        {Math.round(zoom * 100)}%
      </span>
      <button
        className="zoom-btn"
        onClick={() => zoomIn({ duration: 200 })}
        title="放大"
        aria-label="放大"
      >
        <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true">
          <path d="M11 5h2v6h6v2h-6v6h-2v-6H5v-2h6z" fill="currentColor" />
        </svg>
      </button>
      <button
        className="zoom-fit"
        onClick={() => fitView({ padding: 0.18, duration: 200, maxZoom: 1 })}
        title="适应画布（最大 100%）"
      >
        <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
          <path
            d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
        <span>适应</span>
      </button>
    </Panel>
  )
}

export function Canvas({
  tables,
  selected,
  positions,
  onSelect,
  onConnect,
  onEdgeRemove,
  onAddTableAt,
  onNodeDragStop,
  onNodeDragStopMulti,
  ribbonSlots,
  onDeleteNode,
  onDuplicateNode,
  onDeleteSelectedNodes,
  onDetachFlow,
  onDetachLookup,
  onPasteNodes,
  onApplyLayout,
  onRestoreTable,
  onRestoreLastTrash,
  onPurgeTable,
  onEmptyTrash,
  onOpenNode,
  onOpenPopupNode,
  onCreateGroup,
  onOpenRuleMembers,
}: {
  tables: Record<string, TableSchema>
  selected: string
  positions: Record<string, NodePos>
  onSelect: (id: string) => void
  onConnect: (sourceId: string, targetId: string) => void
  onEdgeRemove: (edgeId: string) => void
  onAddTableAt: (pos: XYPosition, kind?: NodeKind) => string
  onNodeDragStop: (id: string, pos: XYPosition) => void
  onNodeDragStopMulti: (updates: { id: string; pos: XYPosition }[]) => void
  /* ★ 合并标题栏插槽（2026-10-19）：App 的 logo/画布名/版本/四钮 渲染进页签行 */
  ribbonSlots?: { current: { left?: ReactNode; right?: ReactNode } } | null
  onDeleteNode: (id: string) => void
  onDuplicateNode: (id: string) => void
  onDeleteSelectedNodes: (ids: string[]) => void
  onDetachFlow: (src: string, tgt: string) => void
  onDetachLookup: (srcTable: string, fieldId: string) => void
  onPasteNodes: (items: CopyItem[], mode?: 'group' | 'group-empty' | 'flat') => string[]
  onApplyLayout: (positions: Record<string, NodePos>) => void
  onRestoreTable: (id: string) => void
  onRestoreLastTrash: () => number
  onPurgeTable: (id: string) => void
  onEmptyTrash: () => void
  /** 双击可编辑节点（docx/xlsx）时打开放大态编辑器 */
  onOpenNode?: (id: string) => void
  /** ★ 弹窗形态（2026-10-19）：双击常规数据节点 → 打开居中弹窗编辑器 */
  onOpenPopupNode?: (id: string) => void
  /** ★ 节点组（2026-10-21）：Ctrl+G 选中建组（rect=成员包围盒） */
  onCreateGroup?: (ids: string[], rect: { x: number; y: number; w: number; h: number }) => void
  /** 双击圆形「规则跟随」节点时打开成员管理面板（改主次 / 跟随不跟随 / 移出组） */
  onOpenRuleMembers?: OpenRuleMembers
}) {
  /* ★ 2026-10-25 wfc 拖入载入进度（用户指令「显示载入进度」） */
  const [wfcLoadStep, setWfcLoadStep] = useState<{ label: string; pct: number } | null>(null)
  /* ★ 载入耗时显示（用户指令「从拖入文件到完整展开 提示个时间」） */
  const [wfcLoadElapsed, setWfcLoadElapsed] = useState(0)
  const wfcLoadStartRef = useRef(0)
  const instRef = useRef<ReactFlowInstance<Node<TableNodeData>, Edge> | null>(null)
  /* ★ 2026-10-24 用户指令「连线动画设置：拖动时暂停或不暂停」三态 */
  /* ★ 2026-10-24 用户指令「对 JS 执行和 React 渲染管线监控」：性能 HUD */
  const [perfHud, setPerfHud] = useState(() => prefGet('perfHud') === '1')
  const changeCountRef = useRef(0)
  const [changeCount, setChangeCount] = useState(0)
  useEffect(() => {
    const h = (e: Event) => {
      const on = (e as CustomEvent).detail === '1'
      setPerfHud(on)
    }
    window.addEventListener('perf-hud-changed', h)
    return () => window.removeEventListener('perf-hud-changed', h)
  }, [])
  useEffect(() => {
    if (!perfHud) return
    const base = changeCountRef.current
    const iv = setInterval(() => setChangeCount(changeCountRef.current - base), 1000)
    return () => clearInterval(iv)
  }, [perfHud])
    const [edgeAnim, setEdgeAnim] = useState<EdgeFx>(() => {
    try { return (localStorage.getItem('edgeAnimMode') as EdgeFx) || 'pause' } catch { return 'pause' }
  })
  const [interacting, setInteracting] = useState(false)
  const cycleEdgeAnim = () => {
    const next = edgeAnim === 'pause' ? 'always' : edgeAnim === 'always' ? 'off' : 'pause'
    setEdgeAnim(next)
    try { localStorage.setItem('edgeAnimMode', next) } catch { /* ignore */ }
    pushLog('[canvas] 连线动画 → ' + (next === 'pause' ? '拖动时暂停' : next === 'always' ? '始终播放' : '关闭'))  /* App「视图」下拉 ↔ 右键循环 双入口同步 */
  useEffect(() => {
    const h = (e: Event) => setEdgeAnim((e as CustomEvent).detail || 'pause')
    window.addEventListener('edge-anim-changed', h)
    return () => window.removeEventListener('edge-anim-changed', h)
  }, [])
  useEffect(() => {
    const h2 = (e: Event) => {
      const on = (e as CustomEvent).detail === '1'
      setAlignMagnet(on)
      if (!on) { guideSigRef.current = ''; setGuides({ v: [], h: [] }) }
    }
    window.addEventListener('align-magnet-changed', h2)
    return () => window.removeEventListener('align-magnet-changed', h2)
  }, [])
  /* App「视图」下拉 ↔ 右键循环 双入口同步 */
  useEffect(() => {
    const h = (e: Event) => setEdgeAnim((e as CustomEvent).detail || 'pause')
    window.addEventListener('edge-anim-changed', h)
    return () => window.removeEventListener('edge-anim-changed', h)
  }, [])
  }
  const pendingPos = useRef<Record<string, XYPosition>>({})
  /* ★ 2026-10-23 用户指令「拖拽文件到空白位置→按类型提示是否建节点并导入」。
     分类规则：jpg/png→photo（多张归一个照片节点）；docx→ooword（每份一节点）；
     xlsx→oosheet（每份一节点，stripFullCalc 摘除导入重算标志）；其余类型拒绝并说明。
     确认框列出分组摘要，取消则什么都不发生。 */
  /* ★ 2026-10-26 全格式拖入空白区（用户指令「其他格式 同样参照该规范」）：
     classifyDrop 分类 → DropChooser 弹窗选动作 → applyDropPick 执行。
     组导入器复用原 dropImportFiles 逻辑（照片 b64 / PDF 走 IDB / OO 走 stashOoBytes）。 */
  const importPhotos = async (fs: File[], basePos: { x: number; y: number }) => {
    const st = useBaseStore.getState()
    const id = onAddTableAt(basePos, 'photo')
    pendingPos.current[id] = basePos
    const items: { id: string; name: string; type: string; b64: string }[] = []
    for (const f of fs) {
      try {
        const buf = await f.arrayBuffer()
        items.push({
          id: 'ph-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7),
          name: f.name,
          type: /png/i.test(f.type) || /\.png$/i.test(f.name) ? 'png' : 'jpg',
          b64: b64FromU8(new Uint8Array(buf)),
        })
      } catch { /* 单张失败跳过 */ }
    }
    st.setNodeConfig(id, { photoItems: items, photoOrder: items.map((it) => it.id) })
    st.renameTable(id, fs.length > 1 ? `照片组（${fs.length} 张）` : (fs[0]?.name.replace(/\.[^.]+$/, '') || '照片'))
  }
  const importPdfsInto = async (kind: 'pdf' | 'pdfslim', fs: File[], basePos: { x: number; y: number }) => {
    const st = useBaseStore.getState()
    const id = onAddTableAt(basePos, kind)
    pendingPos.current[id] = basePos
    const items: { id: string; name: string; inDb: true }[] = []
    for (const f of fs) {
      try {
        const u8 = new Uint8Array(await f.arrayBuffer())
        const itId = 'pd-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
        /* ★ 2026-10-24 性能根治：bytes 入 IndexedDB，config 只留索引（b64 出 state） */
        await putBlob('pdfitem-' + itId, u8)
        items.push({ id: itId, name: f.name, inDb: true })
      } catch { /* 单份失败跳过 */ }
    }
    st.setNodeConfig(id, { pdfItems: items, pdfFileOrder: items.map((it) => it.id) })
    st.renameTable(id, fs.length > 1 ? `PDF 组（${fs.length} 份）` : (fs[0]?.name.replace(/\.pdf$/i, '') || 'PDF'))
    return { id, items }
  }
  const importOOFiles = async (fs: File[], kind: 'ooword' | 'oosheet', basePos: { x: number; y: number }) => {
    const st = useBaseStore.getState()
    let ox = 0
    const ids: string[] = []
    for (const f of fs) {
      const id = onAddTableAt({ x: basePos.x + ox, y: basePos.y }, kind)
      pendingPos.current[id] = { x: basePos.x + ox, y: basePos.y }
      ids.push(id)
      let buf: Uint8Array<ArrayBufferLike> = new Uint8Array(await f.arrayBuffer())
      if (kind === 'oosheet') {
        try { buf = await stripFullCalcFlag(buf) } catch { /* 摘除失败不影响导入 */ }
      }
      st.setNodeConfig(id, {
        ooTemplateId: 'tpl-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7),
        ooTemplateName: f.name,
        ooBytesDbId: await stashOoBytes(id, buf),
      })
      st.renameTable(id, f.name.replace(/\.[^.]+$/, ''))
      ox += 300
    }
    return ids
  }
  const applyDropPick = async (choice: DropChoice) => {
    const pick = dropPick
    setDropPick(null)
    if (!pick || choice === 'none') return
    const { g, pos } = pick
    const st = useBaseStore.getState()
    st.pushHistory()
    let ox = 0
    const done: string[] = []
    /* ★ 2026-10-29「创建并打开」：收集新建节点，建完立即打开第 1 份进入编辑 */
    const openIds: string[] = []
    if ((choice === 'photos' || choice === 'all') && g.photos.length) {
      await importPhotos(g.photos, { x: pos.x + ox, y: pos.y }); ox += 300
      done.push(`照片 ${g.photos.length} 张`)
    }
    if ((choice === 'words' || choice === 'words-open' || choice === 'all') && g.words.length) {
      const ids = await importOOFiles(g.words, 'ooword', { x: pos.x + ox, y: pos.y }); ox += 300 * g.words.length
      done.push(`Word ${g.words.length} 份${choice === 'words-open' ? '（并打开）' : ''}`)
      if (choice === 'words-open' && ids.length) openIds.push(ids[0])
    }
    if ((choice === 'sheets' || choice === 'sheets-open' || choice === 'all') && g.sheets.length) {
      const ids = await importOOFiles(g.sheets, 'oosheet', { x: pos.x + ox, y: pos.y }); ox += 300 * g.sheets.length
      done.push(`Excel ${g.sheets.length} 份${choice === 'sheets-open' ? '（并打开）' : ''}`)
      if (choice === 'sheets-open' && ids.length) openIds.push(ids[0])
    }
    if ((choice === 'pdf' || choice === 'slim' || choice === 'all') && g.pdfs.length) {
      const kind = choice === 'slim' ? 'pdfslim' : 'pdf'
      await importPdfsInto(kind, g.pdfs, { x: pos.x + ox, y: pos.y }); ox += 300
      done.push(`PDF ${g.pdfs.length} 份${choice === 'slim' ? '（瘦身）' : ''}`)
    }
    if (choice === 'merge' && g.pdfs.length) {
      const { id: pid, items } = await importPdfsInto('pdf', g.pdfs, pos)
      const mpos = { x: pos.x + 340, y: pos.y }
      const mid = onAddTableAt(mpos, 'pdfmerge')
      pendingPos.current[mid] = mpos
      onConnect(pid, mid) /* 统一连线语义：写 merge inputs + 审计（边表由 config 同构派生） */
      pushLog(`[PDF·合并] 已建 PDF 节点（${items.length} 份）并自动连入合并节点`, 'ok')
      done.push(`PDF ${g.pdfs.length} 份 → 合并节点`)
    }
    if (choice === 'sheets-dxf' && g.sheets.length) {
      /* ★ 2026-11-10：第 1 份 Excel 解析后直接落为节点自身 records——卡片列映射后即可展点 */
      const f = g.sheets[0]
      const id = onAddTableAt(pos, 'xlsx2dxf')
      pendingPos.current[id] = pos
      try {
        const parsed = await parseFile(f)
        st.setTableData(id, parsed.fields, parsed.rows)
        st.renameTable(id, f.name.replace(/\.(xlsx|xls|csv)$/i, ''))
        done.push(`Excel转DXF 1 份（载入 ${parsed.rows.length} 行）`)
      } catch (e) {
        done.push(`Excel转DXF 建节点成功但数据解析失败：${String((e as Error)?.message ?? e).slice(0, 60)}`)
      }
    }
    if (choice === 'replace' && g.pdfs.length) {
      const id = onAddTableAt(pos, 'pdfreplace')
      pendingPos.current[id] = pos
      const adds = await Promise.all(g.pdfs.map(async (f) => {
        const bid = 'reprepl-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
        const bytes = new Uint8Array(await f.arrayBuffer())
        await putBlobStandalone(bid, bytes)
        return { id: bid, name: f.name, inDb: true, page: 1 }
      }))
      st.setNodeConfig(id, { replaceTargets: adds })
      st.renameTable(id, g.pdfs.length > 1 ? `替换（${g.pdfs.length} 份）` : (g.pdfs[0]?.name.replace(/\.pdf$/i, '') || 'PDF替换'))
      done.push(`替换页 ${g.pdfs.length} 份（被替换组）`)
    }
    if (done.length) pushLog(`拖放导入：${done.join('、')} → 节点已创建`, 'ok')
    /* ★ 2026-10-29「创建并打开」：编辑器为单例（editorNode 单值），多份时打开第 1 份 */
    if (openIds.length) {
      onOpenNode?.(openIds[0])
      pushLog(`[拖放·打开] 已打开「${useBaseStore.getState().tables[openIds[0]]?.name ?? openIds[0]}」进入编辑`, 'ok')
    }
    if (pick.rejected.length) pushLog(`拖放导入跳过 ${pick.rejected.length} 个不支持文件：${pick.rejected.slice(0, 3).join('、')}${pick.rejected.length > 3 ? ' 等' : ''}`, 'warn')
  }
  const rightRef = useRef<{ active: boolean; moved: boolean; startX: number; startY: number; onNode: boolean; nodeId?: string }>({
    active: false,
    moved: false,
    startX: 0,
    startY: 0,
    onNode: false,
  })
  const leftRef = useRef<{ active: boolean; moved: boolean; startX: number; startY: number }>({
    active: false,
    moved: false,
    startX: 0,
    startY: 0,
  })
  const zRef = useRef<Record<string, number>>({})
  const dragStartPos = useRef<Record<string, XYPosition>>({})
  /* ★ 2026-10-24 用户指令「选择之前整体所见即所得状态 和 拖动时只有轮廓的状态」：
     拖动显示模式 wysiwyg（所见即所得，默认）/ outline（只显示轮廓） */
  const [dragLod, setDragLod] = useState<'wysiwyg' | 'outline'>(() => (prefGet('dragLod') === 'outline' ? 'outline' : 'wysiwyg'))
  const dragLodRef = useRef(dragLod)
  dragLodRef.current = dragLod
  useEffect(() => {
    const h = (e: Event) => {
      const m = (e as CustomEvent).detail === 'outline' ? 'outline' : 'wysiwyg'
      setDragLod(m)
      prefSet('dragLod', m)
    }
    window.addEventListener('drag-lod-changed', h)
    return () => window.removeEventListener('drag-lod-changed', h)
  }, [])
    /* ★ 2026-10-24 用户指令「实时监测」：拖动自动采样——开始记录/逐帧坐标+毫秒/
     松手分析，卡顿（>33ms≥3 次）自动进日志，流畅静默不打扰 */
  const dragSamplesRef = useRef<{ t: number; x: number; y: number }[]>([])
  /* ★ 拖动节点悬停在组上 → 该组高亮+提示「置入节点至节点组」（2026-10-21） */
  const [dropHintId, setDropHintId] = useState<string | null>(null)
  /* ★ 对齐磁吸（2026-10-21 用户指令）：PPT/PS 式对齐追踪线 + 磁吸吸附 */
  const [guides, setGuides] = useState<{
    v: { x: number; y0: number; y1: number }[]
    h: { y: number; x0: number; x1: number }[]
  }>({ v: [], h: [] })
  const guideSigRef = useRef('')
  /* ★ 对齐磁吸开关（2026-10-21 用户指令）：默认开，prefs 持久化 */
  const [alignMagnet, setAlignMagnet] = useState(() => prefGet('alignMagnet') === '1')  /* ★ 2026-10-24 默认关闭（用户指令） */
  const toggleAlignMagnet = () =>
    setAlignMagnet((v) => {
      const nv = !v
      prefSet('alignMagnet', nv ? '1' : '0')
      if (!nv) {
        guideSigRef.current = ''
        setGuides({ v: [], h: [] })
      }
      return nv
    })
  /* ★ Ctrl+G 建组（2026-10-21 用户指令）：选中 ≥2 个非组节点 → 包围盒建组 */
  useEffect(() => {
    const h = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.key.toLowerCase() !== 'g') return
      const sels = nodesRef.current.filter(
        (x) => x.selected && !isGroupNodeKind(tables[x.id]?.kind),
      )
      if (sels.length < 2) return
      e.preventDefault()
      let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity
      for (const n of sels) {
        const w = n.measured?.width ?? 220
        const hh = n.measured?.height ?? 120
        x0 = Math.min(x0, n.position.x)
        y0 = Math.min(y0, n.position.y)
        x1 = Math.max(x1, n.position.x + w)
        y1 = Math.max(y1, n.position.y + hh)
      }
      onCreateGroup?.(sels.map((x) => x.id), { x: x0, y: y0, w: x1 - x0, h: y1 - y0 })
    }
    window.addEventListener('keydown', h)
    return () => window.removeEventListener('keydown', h)
  }, [onCreateGroup, tables])
  /* ★ 画布背景：挂载即应用上次选择的默认背景图（2026-10-19）；
     2026-10-22 起罩色跟主题走——主题切换（app-theme-change）后重应用当前背景图 */
  useEffect(() => {
    applyCanvasBg(readCanvasBg())
    const onThemeChange = () => applyCanvasBg(readCanvasBg())
    window.addEventListener('app-theme-change', onThemeChange)
    return () => window.removeEventListener('app-theme-change', onThemeChange)
  }, [])
  /* ★ 拖动悬停的组高亮标记写进 data（只更新变化的组节点，不惊动其它节点） */
  useEffect(() => {
    setNodes((nds) =>
      nds.map((n) => {
        if (n.data?.kind !== 'group') return n
        const want = n.id === dropHintId
        return n.data.groupHint === want ? n : { ...n, data: { ...n.data, groupHint: want } }
      }),
    )
  }, [dropHintId])

  /* ★ 锁定组拖拽（2026-10-19 用户指令「增加一个锁定圆形节点相对位置的按钮（锁定组）
     放在主节点上。锁定后拖动任何锁定的组的任何节点 整个组会跟着移动」）：
     拖起时若属于已锁定组 → 快照整组成员位置；拖动中按同一增量平移其它成员；
     拖停时整组一次性提交 store（防止重建 effect 用旧位置弹回）。 */
  const grpSnap = useRef<{ ids: string[]; base: Record<string, XYPosition> } | null>(null)
  const resolveGrp = (id: string): string[] | null => {
    const t = useBaseStore.getState().tables
    const k = t[id]?.kind
    /* ★ 节点组：拖组节点 = 整组（组 + config.members）一起移动 */
    if (k === 'group') {
      const m = (t[id]?.config?.members ?? []) as string[]
      return m.length ? [id, ...m] : null
    }
    let regId = id
    if (isRoundNodeKind(k)) {
      const li = (t[id]?.config?.linkedInputs ?? []) as string[]
      regId = li.find((x) => t[x] && !isRoundNodeKind(t[x]?.kind))
        ?? Object.values(t).find((tt) => !isRoundNodeKind(tt.kind) && ((tt.config?.linkedInputs ?? []) as string[]).includes(id))?.id
        ?? ''
      if (!regId) return null
    }
    const reg = t[regId]
    if (!reg || !(reg.config as { grpLocked?: boolean })?.grpLocked) return null
    const ids = groupMembersOf(regId, t)
    return ids.length > 1 ? ids : null
  }
  const clipboardRef = useRef<CopyItem[]>([])
  /* ★ 最近一次 mousemove 屏幕坐标（Ctrl+V 粘贴到光标位置用） */
  const lastMouseRef = useRef<{ x: number; y: number } | null>(null)
  useEffect(() => {
    const onMm = (e: MouseEvent) => { lastMouseRef.current = { x: e.clientX, y: e.clientY } }
    window.addEventListener('mousemove', onMm)
    return () => window.removeEventListener('mousemove', onMm)
  }, [])

  const [nodes, setNodes, onNodesChange] = useNodesState<Node<TableNodeData>>([])
  // ★ 2026-10-08 晚：**自动预加载已撤**——用户实测「刚来时的版本秒开，后面老是遇到问题」：
  //   自动预热会在 DS 转换器刚就绪时轰炸队列（runner 被压死 → 全部文档卡死，18:35 事故实锤）。
  //   回到最初行为：打开就是正常流（DS 健康时本来就秒开）；⚡按钮保留作**手动**预加载入口。
  //   兜底层（/poc 就绪门禁 / 90s 自救 / 心跳守护 / x2t 缓存）全部保留——只兜底不惹事。
  // ★ 2026-10-09 OO 启动预热（用户需求「设置里加个开关，画布里有 oo 系列就冷启动」）：
  //   两阶段，开关 = 顶栏波浪按钮（localStorage 'oo-warmup'）：
  //   ① 环境预热：fetch /warmup——拉 DS + 顺序预读引擎大文件焐热 OS 缓存，**零转换任务**，
  //      与被撤的「自动预加载」（18:35 压死队列事故）完全两码事；当时撤的真根因 lcid 已修复。
  //   ② 引擎全热：对画布**第一个** OO 节点复用 ⚡ 冻结预载（后台真开一次 + 冻结 iframe，
  //      CPU≈0）——Chromium 磁盘缓存 + V8 编译全部焐热（这两层只读文件焐不到），
  //      该节点之后点开 = 冻结现场原地恢复 0 等待，同会话其他文档也吃到全热引擎。
  //   只预热一个节点，不轰炸队列；触发条件 = nodes 里真的出现 OO 节点（画布加载完）。
  const ooWarmupFired = useRef(false)
  useEffect(() => {
    if (ooWarmupFired.current) return
    const ooNodes = nodes.filter((n) => n.data?.kind === 'ooword' || n.data?.kind === 'oosheet')
    if (!ooNodes.length) return
    ooWarmupFired.current = true
    const t = setTimeout(async () => {
      try {
        if (localStorage.getItem('oo-warmup') === '0') return
        // ① 环境预热（DS + OS 缓存）
        try {
          const j = await (await fetch('/warmup')).json()
          console.log('[oo-warmup] 环境预热', JSON.stringify(j))
        } catch { /* 失败静默：首开走正常路径 */ }
        // ② 首个 OO 节点后台真开 + 冻结（⚡ 同款机制）
        const first = ooNodes[0]
        const nid = String(first.id)
        if (preloadState(nid) === 'idle') {
          try {
            const cfg0 = useBaseStore.getState().tables[nid]?.config as { ooTemplateId?: string; ooBytes64?: string; ooBytesDbId?: string } | undefined
            const tplId = cfg0?.ooTemplateId
            /* ★ 2026-10-11：预加载优先用节点内置内容（ooBytes64）——真正随画布走 */
            let ooB = cfg0?.ooBytes64
            if (!ooB) {
              const dbId = cfg0?.ooBytesDbId
              if (dbId) { const r = await b64FromDbId(dbId); if (r) ooB = r }
            }
            await preloadOo({ nodeId: nid, kind: first.data!.kind as 'ooword' | 'oosheet', templateId: tplId, bytes: ooB ? b64ToU8(ooB) : null })
            console.log('[oo-warmup] 首节点已预载冻结:', nid)
          } catch (e) { console.warn('[oo-warmup] 首节点预载失败（不影响正常打开）:', e) }
        }
      } catch { /* ignore */ }
    }, 3000)   // 画布渲染稳定后再预热，不抢首屏
    return () => clearTimeout(t)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nodes])
  // 始终指向最新的 nodes（避免键盘/复制等回调捕获到过期的 nodes 闭包）
  const nodesRef = useRef(nodes)
  nodesRef.current = nodes
  /* 撤销/重做后把 store 的位置同步进 RF 内部（见 onKey 里的说明）。
     用 ref 持有 setNodes，避免按键监听 effect（[] 依赖）捕获过期引用。 */
  const setNodesRef = useRef(setNodes)
  setNodesRef.current = setNodes
  const syncRfFromStore = (fn: () => void) => {
    fn()
    const pos = useBaseStore.getState().positions
    setNodesRef.current((nds) =>
      nds.map((n) => (pos[n.id] ? { ...n, position: pos[n.id] } : n)),
    )
  }
  const [edges, setEdges, onEdgesChange] = useEdgesState<Edge>([])
  /* ★ 2026-11-07 对话加节点需要「当前边」判断流程尾，而 effect 只依赖 setEdges（闭包里的 edges 会过期） */
  const edgesRef = useRef(edges)
  edgesRef.current = edges
  /* ★ 2026-10-26 用户指令「其他格式同样参照该规范」：全格式拖入空白区选择对话框，null=关 */
  const [dropPick, setDropPick] = useState<{ g: DropGroups; rejected: string[]; pos: { x: number; y: number } } | null>(null)
  /** ★ 2026-10-29 用户规则：右键 OO 节点「导入文档…」的隐藏文件选择框 + 目标节点 id */
  const ooImportRef = useRef<HTMLInputElement>(null)
  const ooImportTargetRef = useRef<string | null>(null)
  /* ★ 2026-10-25 连接关系同步 store（孤立节点计算用——保存面板） */
  useEffect(() => {
    const ids: string[] = []
    for (const e of edges) { ids.push(e.source); ids.push(e.target) }
    useBaseStore.getState().setEdgeLinks([...new Set(ids)])
  }, [edges])
  /* ★ 2026-10-25 用户指令「把下面的去掉」：pdfmerge 自身输出口（r1）已移除，
     引用它的遗留连线自动断开（不再飘在节点角落） */
  useEffect(() => {
    const st = useBaseStore.getState()
    const dead = edges.filter((x) => x.sourceHandle === 'rf-port-no-r1' && st.tables[x.source]?.kind === 'pdfmerge')
    if (dead.length) setEdges((es) => es.filter((x) => !dead.includes(x)))
  }, [edges, setEdges])
  /* ★ 2026-11-01 对话意图 → 自动建工作流。
     ★ 2026-11-07 用户「点击也没创建节点」：原先写死 `intentId !== 'image-to-word'` 直接 return，
       而 add-node 这类意图 nodes 为空、按钮又已变「✓ 已在画布创建」——看着成功其实什么都没建。
       现改为通用实现：按 INTENTS 里的 nodes 计划建节点并串联（图片到 Word 额外导入图片），
       没有节点计划的意图明确记日志，不再假装成功。 */
  useEffect(() => {
    const h = async (e: Event) => {
      const d = (e as CustomEvent<{ intentId: string; files: File[] }>).detail
      const it = d?.intentId ? intentById(d.intentId) : null
      if (!it) return
      if (!it.nodes.length) {
        pushLog(`[对话] 「${it.title}」没有节点计划，未创建节点（可改用「加个XX节点」的说法）`, 'warn')
        return
      }
      const stamp = Date.now()
      /* 落点：视口中心起，横向依次排开（不写位置会堆在默认点，被已有节点压住＝「看不见」） */
      const inst = instRef.current
      const el = wrapRef.current
      let origin: { x: number; y: number } | undefined
      if (inst && el) {
        const r = el.getBoundingClientRect()
        origin = inst.screenToFlowPosition({ x: r.x + r.width / 2 - 220, y: r.y + r.height / 2 - 60 })
      }
      /* ★ 2026-11-07 实测：同一 tick 连建 2 个节点，第二个在 React Flow 里迟迟不 measured
         → style 带 visibility:hidden（看着像"没创建"）。逐个建、间隔一帧，稳定可见。 */
      const ids: string[] = []
      for (let i = 0; i < it.nodes.length; i++) {
        const n = it.nodes[i]
        const id = `chat-${it.id}-${stamp}-${i}`
        ids.push(id)
        const s2 = useBaseStore.getState()
        s2.addTable({ id, name: n.name, kind: n.kind, fields: [], ...(n.config ? { config: n.config } : {}) } as never)
        if (origin) s2.setNodePosition(id, { x: Math.round(origin.x + i * 320), y: Math.round(origin.y) })
        if (i < it.nodes.length - 1) await new Promise((r) => setTimeout(r, 90))
      }
      /* ★ 连线写 linkedInputs（派生真相），不 setEdges；延后一帧等节点 measured，
         否则首帧边会被 React Flow 跳过（节点有了线却不显示）。理由见下方 chat-add-node 注释 */
      if (ids.length > 1) {
        setTimeout(() => {
          ids.slice(1).forEach((id, i) => {
            const cur = (useBaseStore.getState().tables[id]?.config ?? {}) as { linkedInputs?: string[] }
            const linked = cur.linkedInputs ?? []
            if (!linked.includes(ids[i])) useBaseStore.getState().setNodeConfig(id, { linkedInputs: [ids[i], ...linked] })
          })
        }, 120)
      }
      await new Promise((r) => setTimeout(r, 150))
      /* 「图片 → Word」：把图片导进第一个节点（图片导入） */
      if (it.id === 'image-to-word' && d.files?.length) {
        const t = useBaseStore.getState().tables[ids[0]]
        if (t) await importOoFileToNode(t, d.files)
      }
      pushLog(`[对话] 「${it.title}」已创建：${it.nodes.map((n) => n.name).join(' → ')}` + (d.files?.length ? `，导入 ${d.files.length} 个文件` : ''), 'ok')
    }
    window.addEventListener('chat-create-workflow', h as EventListener)
    return () => window.removeEventListener('chat-create-workflow', h as EventListener)
  }, [setEdges])
  /* ★ 2026-11-07 对话「添加节点」：通用加节点+自动连线到流程尾部（无下游的节点）。
     ★ 用户指令「要接一个 pdf 合并 而不是 PDF 节点，PDF 节点是承载 PDF 文件的」：
       说「输出为PDF / 生成PDF」现在落到 pdfmerge（PDF生成）。pdfmerge 是**汇入型**节点，
       上游真相是 config.inputs / linkedInputs（卡片按它列词条）——只加一条边不写配置，
       卡片会显示「无上游」；故与 App.handleConnect 的 isMergeKind 分支同口径落库。 */
  useEffect(() => {
    const h = (e: Event) => {
      const d = (e as CustomEvent<{ kind: string; name: string }>).detail
      if (!d?.kind) return
      const st = useBaseStore.getState()
      const newId = 'chat-node-' + Date.now()
      /* 找流程尾：有出线（source）的节点不算尾；多尾取 id 最大（最新） */
      const es = edgesRef.current
      const hasOut = new Set(es.map((x) => x.source))
      const tails = Object.keys(st.tables).filter((id) => st.tables[id] && !st.tables[id].deleted && !hasOut.has(id) && id !== newId)
      const tail = tails.sort().pop()
      const base = tail ? st.positions[tail] : undefined
      /* ★★★ 2026-11-07 用户「点击也没创建节点」：此前 addTable 不写位置 → 新节点落在默认/原点，
         被已有节点压住或跑出视口，看上去就是「没创建」。这里给明确落点：
         流程尾右侧一格；没有尾就放视口中心。 */
      let pos: { x: number; y: number } | undefined
      if (base) pos = { x: base.x + 320, y: base.y + 10 }
      else {
        const inst = instRef.current
        const el = wrapRef.current
        if (inst && el) {
          const r = el.getBoundingClientRect()
          pos = inst.screenToFlowPosition({ x: r.x + r.width / 2 - 110, y: r.y + r.height / 2 - 60 })
        }
      }
      st.addTable({ id: newId, name: d.name, kind: d.kind, fields: [] } as never)
      if (pos) st.setNodePosition(newId, { x: Math.round(pos.x), y: Math.round(pos.y) })
      if (!tail) {
        pushLog('[对话] 已添加节点「' + d.name + '」（画布为空，未连线）', 'ok')
        return
      }
      /* ★★★ 2026-11-07 连线**唯一真相** = 下游节点的 config.linkedInputs（projectFormat.edgesOf 派生）。
         以前只 setEdges 加临时边，下一次 tables 变化全表重建就被冲掉 → 「节点建了但没有线」。
         延后一帧再写：首帧节点还没 measured，边会被 RF 跳过不渲染（重启后才出现，实测复现）。 */
      setTimeout(() => {
        const cfg = (useBaseStore.getState().tables[newId]?.config ?? {}) as { inputs?: string[]; linkedInputs?: string[] }
        const inputs = cfg.inputs ?? []
        const linked = cfg.linkedInputs ?? []
        if (linked.includes(tail)) return
        useBaseStore.getState().setNodeConfig(newId, {
          ...(isMergeKind(d.kind) ? { inputs: inputs.includes(tail) ? inputs : [tail, ...inputs] } : {}),
          linkedInputs: [tail, ...linked],
        })
      }, 120)
      pushLog('[对话] 已添加节点「' + d.name + '」并连线到流程尾部', 'ok')
    }
    window.addEventListener('chat-add-node', h as EventListener)
    return () => window.removeEventListener('chat-add-node', h as EventListener)
  }, [setEdges])
  /* ★ 2026-11-07 用户「执行转换」：对话框只说"已创建"不算完成——这里真正跑一遍输出节点。
     目前接的是 PDF生成（runPdfMerge，与卡片上的「生成 PDF」同一函数）；结果回派 chat-run-result 给对话框。
     选点规则：指定了 nodeId 就用它；否则取「有上游的 pdfmerge」，多个取最新。 */
  useEffect(() => {
    const h = async (e: Event) => {
      const d = (e as CustomEvent<{ nodeId?: string }>).detail ?? {}
      const st = useBaseStore.getState()
      let id = d.nodeId && st.tables[d.nodeId] && !st.tables[d.nodeId].deleted ? d.nodeId : ''
      if (!id) {
        const cands = Object.entries(st.tables).filter(([, t]) => t && !t.deleted && t.kind === 'pdfmerge')
        if (!cands.length) {
          window.dispatchEvent(new CustomEvent('chat-run-result', { detail: { ok: false, msg: '画布上还没有「PDF生成」节点——先说「把文档转成PDF」，我建一个再执行。' } }))
          return
        }
        const withUp = cands.filter(([, t]) => {
          const c = (t.config ?? {}) as { inputs?: string[]; linkedInputs?: string[] }
          return (c.inputs?.length ?? 0) > 0 || (c.linkedInputs?.length ?? 0) > 0
        })
        const pool = (withUp.length ? withUp : cands).map(([k]) => k).sort()
        /* ★ 2026-11-07 用户说「保存画布文件」被误派成执行 → 回了「已合并 1 份 / 3 页」。
           真因之一是这里：没有指定节点时**闷头挑一个**跑，用户根本不知道跑的是哪个。
           多个候选就问一句，别再替用户做决定。 */
        if (pool.length > 1) {
          const names = pool
            .map((k) => {
              const t = st.tables[k]
              const no = String((t?.config as { nodeNo?: string } | undefined)?.nodeNo ?? '')
              return `${no || '（无编号）'} ${t?.name ?? k}`
            })
            .join('、')
          window.dispatchEvent(new CustomEvent('chat-run-result', {
            detail: { ok: false, msg: `画布上有 ${pool.length} 个「PDF生成」节点（${names}）。要跑哪一个？说「执行 N0xx」我就跑那个。` },
          }))
          return
        }
        id = pool[pool.length - 1]
      }
      const r = await runPdfMerge(id)
      /* 对话框若已关掉监听也不会报错（事件没人接而已）；节点上的运行徽章与日志照常打点 */
      const ran = st.tables[id]
      const ranNo = String((ran?.config as { nodeNo?: string } | undefined)?.nodeNo ?? '')
      window.dispatchEvent(new CustomEvent('chat-run-result', {
        detail: { ...r, msg: `${ranNo ? ranNo + ' ' : ''}${ran?.name ?? ''}：${r.msg ?? ''}` },
      }))
    }
    window.addEventListener('chat-run-node', h as EventListener)
    return () => window.removeEventListener('chat-run-node', h as EventListener)
  }, [])
  /* ★ 网址监控轮询器：应用打开期间按各节点间隔自动检查（2026-11-09） */
  useWebmonPoller()
    /* ★ 2026-11-08 对话「word转PDF」：ChatDock 识别意图后派 chat-to-pdf-node（带 nodeId），
     走右键「转为PDF节点」同一条链路（runNodeToPdf），结果回派 chat-topdf-result 填对话气泡。 */
  useEffect(() => {
    const h = async (ev: Event) => {
      const d = (ev as CustomEvent<{ nodeId?: string }>).detail ?? {}
      if (!d.nodeId) return
      const r = await runNodeToPdf(d.nodeId, { onAddTableAt, pendingPos, nodesRef })
      window.dispatchEvent(new CustomEvent('chat-topdf-result', { detail: r }))
    }
    window.addEventListener('chat-to-pdf-node', h as EventListener)
    return () => window.removeEventListener('chat-to-pdf-node', h as EventListener)
  }, [onAddTableAt])
  /* ★ 2026-11-07 用户「N071节点」：对话框按编号回答时，画布要同步选中并居中——
     否则用户只看到一段文字，不知道说的是哪个框。
     用 setCenter 而不是 fitView：fitView 会把缩放拉满，反而看不清单个节点。 */
  useEffect(() => {
    const h = (e: Event) => {
      const id = (e as CustomEvent<{ nodeId?: string }>).detail?.nodeId
      if (!id) return
      const st = useBaseStore.getState()
      if (!st.tables[id] || st.tables[id].deleted) return
      setNodes((nds) => nds.map((n) => ({ ...n, selected: n.id === id, data: { ...n.data, selected: n.id === id } })))
      onSelect(id)
      const inst = instRef.current
      const nd = nodesRef.current.find((n) => n.id === id)
      const pos = st.positions[id] ?? nd?.position
      if (!inst || !pos) return
      const w = nd?.measured?.width ?? 240
      const hh = nd?.measured?.height ?? 120
      inst.setCenter(pos.x + w / 2, pos.y + hh / 2, { zoom: 1, duration: 420 })
    }
    window.addEventListener('chat-focus-node', h as EventListener)
    return () => window.removeEventListener('chat-focus-node', h as EventListener)
  }, [onSelect, setNodes])
  /* ★ 2026-10-25 合并历史锚点连线清理：历史条目删除/被挤出列表时，对应行级锚点的连线断开 */
  useEffect(() => {
    const h = (e: Event) => {
      const d = (e as CustomEvent<{ nodeId: string; keys: string[] }>).detail
      if (!d?.keys?.length) return
      setEdges((es) => es.filter((x) => !(x.source === d.nodeId && x.sourceHandle && d.keys.includes(x.sourceHandle))))
    }
    window.addEventListener('pdfmerge-prune-edges', h)
    return () => window.removeEventListener('pdfmerge-prune-edges', h)
  }, [setEdges])
  /* ★ 拖动变更 rAF 合并缓冲（见 onNodesChange 内注释） */
  const pendingNodeChanges = useRef<NodeChange<Node<TableNodeData>>[]>([])
  const nodeChangeScheduled = useRef(false)
  const [menu, setMenu] = useState<{ x: number; y: number; kind: 'node' | 'pane'; nodeId?: string; flowX?: number; flowY?: number } | null>(null)
  /* ★ 2026-10-23 用户指令「右键节点 以此为基准 重新排列：左侧按连接顺序、右侧也是」：
     以右键节点为锚，上游往左、下游往右分层重排（拓扑层级松弛），层内按原 y 堆叠；
     圆形辅助节点（AI/参数库/规则跟随）与组节点不参与，基准节点原地不动。 */
  const rerangeFromNode = (baseId: string) => {
    const st = useBaseStore.getState()
    const tables = st.tables
    const nds = nodesRef.current
    const base = nds.find((n) => n.id === baseId)
    if (!base) return
    /* 参与者：与基准连通的普通节点（排除圆/组/start 与自身） */
    const partake = new Set<string>()
    const isPart = (id: string) => {
      const t = tables[id]
      return !!t && !t.deleted && !isRoundNodeKind(t.kind) && !isGroupNodeKind(t.kind) && t.kind !== 'start'
    }
    /* 连通集（双向 BFS） */
    const adj = new Map<string, string[]>()
    for (const t of Object.values(tables)) {
      if (!t || t.deleted || !isPart(t.id)) continue
      for (const pe of edgesOf(t, tables)) {
        if (!isPart(pe.source) || !isPart(pe.target)) continue
        if (!adj.has(pe.source)) adj.set(pe.source, [])
        if (!adj.has(pe.target)) adj.set(pe.target, [])
        adj.get(pe.source)!.push(pe.target)
        adj.get(pe.target)!.push(pe.source)
      }
    }
    if (!isPart(baseId)) { pushLog('重排：基准需为普通节点', 'warn'); return }
    const q = [baseId]; partake.add(baseId)
    while (q.length) { const cur = q.shift()!; for (const nb of adj.get(cur) ?? []) if (!partake.has(nb)) { partake.add(nb); q.push(nb) } }
    /* 层级松弛：base=0，上游取 min、下游取 max，迭代至稳定 */
    const level = new Map<string, number>([[baseId, 0]])
    for (let round = 0; round < partake.size + 2; round++) {
      let changed = false
      for (const t of Object.values(tables)) {
        if (!t || t.deleted || !isPart(t.id)) continue
        for (const pe of edgesOf(t, tables)) {
          if (!partake.has(pe.source) || !partake.has(pe.target)) continue
          const ls = level.get(pe.source); const lt = level.get(pe.target)
          if (ls != null && (lt == null || lt < ls + 1)) { level.set(pe.target, ls + 1); changed = true }
          if (lt != null && (ls == null || ls > lt - 1)) { level.set(pe.source, lt - 1); changed = true }
        }
      }
      if (!changed) break
    }
    /* ★ 2026-10-23 用户指令「测算每个节点的坐标和尺寸 确保上下节点距离
       不遮挡也不太远」：逐层实测层内最大实际宽高——
       行距 = 层内最大高 + 24（紧凑不遮挡）；列距按前层最大宽 + 40 逐层累进。
       层内排序 = barycenter（靠基准侧邻居平均 y，=连接顺序少交叉），
       分布 = 围绕基准水平中线均匀对称（N 个 = 上下各 N/2），逐层递归。 */
    const layers = new Map<number, string[]>()
    for (const [id, lv] of level) { if (!layers.has(lv)) layers.set(lv, []); layers.get(lv)!.push(id) }
    const nodeById = (id: string) => nds.find((n) => n.id === id)
    const baseH = base.measured?.height ?? 120
    const baseW = base.measured?.width ?? 220
    const baseCenterY = base.position.y + baseH / 2
    const GAP_V = 24
    const GAP_H = 40
    /* 层实测盒（层内最大宽/高） */
    const boxOf = (lv: number) => {
      let w = 0; let h = 0
      for (const id of layers.get(lv) ?? []) {
        const n = nodeById(id)
        if (!n) continue
        w = Math.max(w, n.measured?.width ?? 220)
        h = Math.max(h, n.measured?.height ?? 120)
      }
      return { w, h }
    }
    /* X 逐层累进：正方向 = 前层X + 前层宽 + GAP_H；负方向 = 后层X - 本层宽 - GAP_H */
    const xOf = new Map<number, number>([[0, base.position.x]])
    const posLv = [...new Set([...level.values()])].sort((a, b) => a - b)
    for (const lv of posLv.filter((v) => v > 0)) {
      const prev = xOf.get(lv - 1)!
      xOf.set(lv, prev + boxOf(lv - 1).w + GAP_H)
    }
    for (const lv of posLv.filter((v) => v < 0).sort((a, b) => b - a)) {
      const next = xOf.get(lv + 1)!
      xOf.set(lv, next - boxOf(lv).w - GAP_H)
    }
    /* 逐层布局（|level| 从小到大）：barycenter 排序 + 中线对称分布 */
    const centerYOf = new Map<string, number>([[baseId, baseCenterY]])
    const pos = new Map<string, { x: number; y: number }>()
    pos.set(baseId, { x: base.position.x, y: base.position.y })
    const layerNums = [...new Set([...level.values()])].sort((a, b) => Math.abs(a) - Math.abs(b))
    for (const lv of layerNums) {
      if (lv === 0) continue
      const toward = lv - Math.sign(lv)
      const rowH = boxOf(lv).h + GAP_V
      const x = xOf.get(lv)!
      const ids = (layers.get(lv) ?? []).map((id) => nodeById(id)!).filter(Boolean)
      /* ★ 2026-10-23 用户实测「没按锚点上下的排序顺序排列」：
         排序键 = 该节点连向靠基准侧节点的边，其基准侧锚点（handle）的实测
         垂直位置（DOM 量锚中心 → flow 坐标）——锚 1（第 1 行）连的排最上，
         锚 2 第二……即「锚点行序 = 排列序」。无锚/量不到退回 barycenter。 */
      const rfEdges = instRef.current?.getEdges?.() ?? []
      const s2f = instRef.current?.screenToFlowPosition?.bind(instRef.current)
      const anchorYOf = (curId: string, nbId: string): number | null => {
        for (const e of rfEdges) {
          const hit = (e.source === nbId && e.target === curId) || (e.source === curId && e.target === nbId)
          if (!hit) continue
          const nbSide = e.source === nbId ? 'source' : 'target'
          const hid = nbSide === 'source' ? e.sourceHandle : e.targetHandle
          const nid = nbSide === 'source' ? e.source : e.target
          if (!hid || !nid || !s2f) break
          const el = document.querySelector('.react-flow__node[data-id="' + nid + '"] .react-flow__handle[data-handleid="' + hid + '"]') as HTMLElement | null
          if (!el) continue
          const r = el.getBoundingClientRect()
          return s2f({ x: r.x + r.width / 2, y: r.y + r.height / 2 }).y
        }
        return null
      }
      /* ★ 2026-10-24 用户反馈「根据接入位置上下排 不让连线交叉」（排序仍不对）：
         全图 barycenter 迭代（Sugiyama 简化）——每节点取连通邻居平均 y，3 轮收敛。
         层内顺序由全图连接关系决定（而非单层锚点/旧坐标），交叉最少化。 */
      const iterY = new Map<string, number>()
      for (const id of partake) {
        const n0 = nds.find((x) => x.id === id)
        iterY.set(id, n0 ? n0.position.y : 0)
      }
      for (let round = 0; round < 3; round++) {
        const next = new Map(iterY)
        for (const id of partake) {
          const nbs = (adj.get(id) || []).filter((nb) => iterY.has(nb))
          if (nbs.length) next.set(id, nbs.reduce((a, b) => a + (iterY.get(b) || 0), 0) / nbs.length)
        }
        for (const [k2, v2] of next) iterY.set(k2, v2)
      }
      const keyed = ids.map((n) => {
        const h = n.measured?.height ?? 120
        const selfCy = n.position.y + h / 2
        const nbCys = (adj.get(n.id) ?? [])
          .filter((nb) => level.get(nb) === toward && centerYOf.has(nb))
          .map((nb) => centerYOf.get(nb)!)
        const bary = nbCys.length ? nbCys.reduce((a, b) => a + b, 0) / nbCys.length : selfCy
        /* 锚 y 优先（多邻居取最小锚 y=最靠上的连接） */
        let key: number | null = null
        for (const nb of adj.get(n.id) ?? []) {
          if (level.get(nb) !== toward || !centerYOf.has(nb)) continue
          const ay = anchorYOf(n.id, nb)
          if (ay != null) key = key == null ? ay : Math.min(key, ay)
        }
        return { id: n.id, h, key: iterY.get(n.id) ?? key ?? bary }
      }).sort((a, b) => a.key - b.key)
      /* ★ 2026-10-23 用户指令「前面的前面/后面的后面 就依前面/后面为基准计算位置」：
         分布中心 = 本层各节点与前面层连接点的平均 y（=bary 均值，随前面层走），
         不再固定基准中线——逐层链式参照，层间紧凑不错位。 */
      const distCenter = keyed.length
        ? keyed.reduce((acc, k) => acc + k.key, 0) / keyed.length
        : baseCenterY
      keyed.forEach((k, i) => {
        const cy = distCenter + (i - (keyed.length - 1) / 2) * rowH
        centerYOf.set(k.id, cy)
        pos.set(k.id, { x, y: cy - k.h / 2 })
      })
    }
    /* ★ 2026-10-23 用户指令「带附属圆形节点 也要排列位置 计算附属节点位置
       并计算占用尺寸 并躲避开」：附属圆参与布局——排在主节点正上方串列
       （x 对齐主节点中心、自下而上堆叠），每放一个即计入占用盒；
       放置时对全部已占盒避让（先左右让位、再上移），全程无重叠。 */
    const boxes = new Map<string, { x: number; y: number; w: number; h: number }>()
    for (const [id, np] of pos) {
      const n = nds.find((x) => x.id === id)!
      boxes.set(id, { x: np.x, y: np.y, w: n?.measured?.width ?? 220, h: n?.measured?.height ?? 120 })
    }
    const hitAny = (x: number, y: number, w: number, h: number, skip?: string) => {
      for (const [bid, b] of boxes) {
        if (skip && bid === skip) continue
        if (x < b.x + b.w && x + w > b.x && y < b.y + b.h && y + h > b.y) return true
      }
      return false
    }
    const roundShift = new Map<string, { x: number; y: number }>()
    /* 按主节点分组附属圆（主节点 = linkedInputs 里第一个在 pos 中的 id） */
    const hostRounds = new Map<string, { id: string; w: number; h: number }[]>()
    for (const t of Object.values(tables)) {
      if (!t || t.deleted || !isRoundNodeKind(t.kind)) continue
      const rn = nds.find((n) => n.id === t.id)
      if (!rn) continue
      const li = ((t.config ?? {}) as { linkedInputs?: string[] }).linkedInputs ?? []
      const host = li.find((hid) => pos.has(hid))
      if (!host) continue
      if (!hostRounds.has(host)) hostRounds.set(host, [])
      hostRounds.get(host)!.push({ id: t.id, w: rn.measured?.width ?? 90, h: rn.measured?.height ?? 90 })
    }
    /* 逐主节点放置附属串：主节点顶上方，x 对齐主节点中心，自下而上；避让=左右让位→上移 */
    for (const [host, rounds] of hostRounds) {
      const hb = boxes.get(host)!
      const hostCx = hb.x + hb.w / 2
      let cursorY = hb.y - 16
      for (const r of rounds) {
        let x = hostCx - r.w / 2
        let y = cursorY - r.h
        /* 避让：先左右 ±(w+16) 试位，再整体上移一行，最多 12 次 */
        let tries = 0
        while (hitAny(x, y, r.w, r.h, r.id) && tries < 12) {
          tries++
          const dir = tries % 2 === 1 ? 1 : -1
          const step = Math.ceil(tries / 2) * (r.w + 16)
          const tx = hostCx - r.w / 2 + dir * step
          if (!hitAny(tx, y, r.w, r.h, r.id)) { x = tx; break }
          if (tries % 2 === 0) y -= r.h + 12
        }
        roundShift.set(r.id, { x, y })
        boxes.set(r.id, { x, y, w: r.w, h: r.h })
        cursorY = y - 12
      }
    }
    st.pushHistory()
    /* ★ 2026-10-23 用户实测「自动排列的布局不会保存」：重排此前只改 RF 画布层
       （setNodes），没写回数据层 tables[].position → 刷新即回旧位。
       同步走 setNodePositions（与拖停落库同通道），附属圆一并落库。 */
    const updates: { id: string; pos: { x: number; y: number } }[] = []
    for (const [id, np] of pos) if (id !== baseId) updates.push({ id, pos: np })
    for (const [id, rp] of roundShift) updates.push({ id, pos: rp })
    /* ★ 2026-10-24 用户指令「拖动不被影响、重排按最新连接避免交叉」：重排收尾——
       对每个 PDF 合并节点，行序写入 config.pdfOrder（按重排后上游 y）。
       平时拖动源节点不触行序（渲染读 pdfOrder 稳定），交叉由重排统一消除。 */
    try {
      const posMap = new Map<string, { x: number; y: number }>()
      for (const u of updates) posMap.set(u.id, u.pos)
      const mergeIds = [...partake].filter((id) => tables[id]?.kind === 'pdfmerge')
      for (const mid of mergeIds) {
        /* ★ 修正（11:19 用户截图逆序实锤）：adj 是正向邻接（source→target），
           adj.get(mid) 取到的是下游——合并节点的上游要反向查：
           所有 adj.get(src) 包含 mid 的 src 才是上游。之前写进了错误的集合。 */
        const ups: { id: string; y: number }[] = []
        for (const src of partake) {
          if (!(adj.get(src) || []).includes(mid)) continue
          const p = posMap.get(src)
          if (p) ups.push({ id: src, y: p.y })
        }
        if (ups.length >= 2) {
          ups.sort((a, b) => a.y - b.y)
          st.setNodeConfig(mid, { pdfOrder: ups.map((u) => u.id) })
        }
      }
    } catch { /* 行序写入失败不影响重排本体 */ }
    st.setNodePositions(updates)
    setNodes((nds2) => nds2.map((n) => {
      if (pos.has(n.id) && n.id !== baseId) return { ...n, position: pos.get(n.id)! }
      const rs = roundShift.get(n.id)
      if (rs) return { ...n, position: rs }
      return n
    }))
    pushLog("重排：以「" + (st.tables[baseId]?.name ?? baseId) + "」为基准，共 " + (pos.size - 1) + " 个连通节点已分层重排", 'ok')
  }
  const [marquee, setMarquee] = useState<{ x: number; y: number; w: number; h: number } | null>(null)
  /* ★ 2026-10-23 用户澄清「我说的是画布：最大化要适应画布 最大比例不超过
     100% 且居中」——窗口尺寸变化（最大化/还原/拖边缘）后画布自动 fitView：
     内容整体适配视口并居中，缩放封顶 100%（不放大超过原始大小）。
     防抖 150ms（拖窗口边缘连续触发），300ms 平滑过渡。 */
  useEffect(() => {
    /* ★ 2026-10-23 用户指令「只有最大化的时候生效 拖动的时候不要」：
       用尺寸跳变阈值区分——最大化/还原=单次大跳变（≥200px）触发 fitView；
       拖窗口边缘=连续小步变化（几十 px/次）不干预。 */
    let prevW = window.innerWidth
    let prevH = window.innerHeight
    let t: number | undefined
    const onResize = () => {
      const dw = Math.abs(window.innerWidth - prevW)
      const dh = Math.abs(window.innerHeight - prevH)
      prevW = window.innerWidth
      prevH = window.innerHeight
      if (Math.max(dw, dh) < 200) return
      window.clearTimeout(t)
      t = window.setTimeout(() => {
        instRef.current?.fitView({ padding: 0.2, maxZoom: 1, duration: 300 })
      }, 150)
    }
    window.addEventListener('resize', onResize)
    return () => {
      window.removeEventListener('resize', onResize)
      window.clearTimeout(t)
    }
  }, [])
  const [marqueeMode, setMarqueeMode] = useState(false)
  const [trashOpen, setTrashOpen] = useState(false)
  useEffect(() => {
    const h = () => setTrashOpen((v) => !v)
    window.addEventListener('toggle-trash-panel', h)
    return () => window.removeEventListener('toggle-trash-panel', h)
  }, [])
  const [runningIds, setRunningIds] = useState<string[]>([])
  // 当前悬停的连线 id：用于控制删除按钮仅在鼠标位于线条上时才显示
  const [hoverEdge, setHoverEdge] = useState<string | null>(null)
  /* 是否正在拉连线 + 拖的是哪种线（2026-10-06 用户：「从节点左边 接 节点右边的
     锚点连线的时候 上下的锚点不用显示」）：
       'rule' = 起点是规则口（isRulePortId）→ 规则线拖拽，上1/上2 显形（原行为）
       'data' = 起点是数据流口（左口/右口）→ 上下的规则口全部隐藏
     判据用 handleId（onConnectStart 的第二参数），比「连到哪」早——起手那一刻
     就该定显隐，而不是等落到目标口。
     ★ 为什么要自己维护、而不是用 RF 的状态：实测拖线全程对比
       `.react-flow` / `body` / `.react-flow__pane` 的 class 与 dataset，
       三个元素在「连线前 / 连线中 / 连线后」完全一致，RF 没有任何字段
       随连接过程变化（源码里 connectionInProgress 只进 React props，不落DOM）。
       所以只能由 onConnectStart/End 维护，打成 .rf-connecting / .rf-connecting-data 供 CSS 用。
     ★ 为什么不能只靠 connectingto/connectingfrom：RF 只给**连线两端**加这两个类，
       画布上其它可连节点的口仍是隐形的，用户看不到「还有哪些能连」。*/
  const [connectKind, setConnectKind] = useState<null | 'data' | 'rule'>(null)
  // 离开连线时延迟清除悬停：避免「删除按钮出现在光标下 → 触发连线 mouseleave → 按钮立即隐藏」的死锁，
  // 否则按钮永远点不到（鼠标移到按钮上的瞬间连线已 unhover，按钮被 pointer-events:none 挡住收不到 hover）
  const edgeLeaveTimer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const wrapRef = useRef<HTMLDivElement>(null)
  const [needCenter, setNeedCenter] = useState(false)
  // Shift 按住时启用「Shift+左键拖拽框选」，并临时关闭 RF 平移以避免冲突
  const [shiftDown, setShiftDown] = useState(false)
  // 顶部 ribbon 要显示缩放比例，而 useViewport 只能在 <ReactFlow> 子树里用（Canvas 本身不在 provider 内），
  // 所以复用 onMoveEnd 事件把 zoom 同步到 state。
  const [zoomPct, setZoomPct] = useState(100)
  const pastLen = useBaseStore((s) => s.past.length)
  const records = useBaseStore((s) => s.records)
  // 全局折叠态：给 CanvasRibbon 的按钮做 on 显示与 toggle（TableNode 侧另有一份自己的订阅）
  const nodesFolded = useBaseStore((s) => s.nodesFolded)
  // 圆形节点运行动画：同上，CanvasRibbon 测试按钮的 on 显示用（TableNode 另有自己的订阅）
  const roundRunState = useBaseStore((s) => s.roundRunState)
  /* 运行动画的定时器句柄（2026-10-06 第2条）。
     放 ref 不放 state：它不需要触发渲染，且每次调度都要先清掉旧的
     —— 连点两次播放会挂两个 3s 定时器，第二个到点时状态已在 exiting，
     重复推进会让退出动画卡住不消失。
     ★卸载时必须清理：否则组件卸载后定时器仍会 setState（跨页面/热更新时
     报「can't update unmounted component」）。 */
  /* 顶栏播放按钮的置灰判定：有没有选中的 AI/参数库（core/roundRun 纯函数，
     与 TableNode 的播放判定同一份 kind 口径）。
     走 RF 的 onSelectionChange 拿选中节点 —— nodes 是 React Flow 的受控状态，
     这里读的是 Canvas 自己的 nodes（已含 selected 标记）。 */
  const selNoneHint = !hasPlayableSelection(nodes.filter((n) => n.selected).map((n) => tables[n.id]?.kind ?? ''))
  const roundRunTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  useEffect(
    () => () => {
      if (roundRunTimerRef.current) {
        clearTimeout(roundRunTimerRef.current)
        roundRunTimerRef.current = null
      }
    },
    [],
  )
  const futureLen = useBaseStore((s) => s.future.length)
  const pages = useBaseStore((s) => s.pages)
  const activePageId = useBaseStore((s) => s.activePageId)
  const trashStack = useBaseStore((s) => s.trashStack)
  const fittingRef = useRef(false)
  /* 「打开成员管理面板」回调走 ref 而不是直接进节点重建 effect 的依赖：
     它是 App 层传进来的函数，若直接进依赖且调用方写成内联箭头函数，
     每次父组件重渲染都是新引用 → 触发 nodes 整体重建 → 拖拽中的节点被弹回。
     用 ref 持有最新值，重建时读取，既拿到最新回调又不影响依赖。*/
  const openRuleMembersRef = useRef<OpenRuleMembers | undefined>(onOpenRuleMembers)
  openRuleMembersRef.current = onOpenRuleMembers

  // 结构变化（表增删/改名/选中）时重建 nodes，保留已有位置与已选中的多选
  useEffect(() => {
    setNodes((prev) => {
      const prevById = new Map(prev.map((n) => [n.id, n]))
      const prevPos = new Map(prev.map((n) => [n.id, n.position]))
      const prevSel = new Map(prev.map((n) => [n.id, n.selected]))
      // 软删除（回收站）的节点不渲染
      return Object.values(tables)
        .filter((t) => !t.deleted)
        .map((t, i) => {
        const existing = prevById.get(t.id)
        const pending = pendingPos.current[t.id]
        if (pending) delete pendingPos.current[t.id]
        // 优先沿用已渲染节点的位置（含 React Flow 测量/拖拽结果），避免重建把位置弹回
        const pos = existing?.position ?? pending ?? positions[t.id] ?? prevPos.get(t.id) ?? defaultPos(i)
        const isSel = pending ? true : (prevSel.get(t.id) ?? false)
        return {
          // 展开保留已有节点对象：React Flow 测量后会写入 measured 等内部字段，
          // 若整体重建为新对象会把这些字段丢掉，导致节点永久卡在 visibility:hidden（测量中）。
          ...existing,
          id: t.id,
          type: 'table',
          position: pos,
          /* ★ 节点组（2026-10-21）：z 跟随自身选中态——
             未选中 = -1 沉底（成员卡片在组框之上，点成员直达成员）；
             选中 = 1 抬到成员之上（缩放条亮起，拖边框调大小不被成员边缘挡住） */
          /* ★ 组 z 分层（2026-10-22 用户「拉出去拉回来 边界老是变」）：
             未选中 = -1 沉底（成员清晰、手柄隐藏不碍事）；
             选中 = 1 抬升（NodeResizer 手柄在成员之上，可拖；RF 拖拽期间
             组尺寸不会再被重建 effect 的 existing.position 顶掉——
             这正是「拉出去再拉回来边界老变」的根因之一）。 */
          zIndex: t.kind === 'group' ? (isSel ? 1 : -1) : zRef.current[t.id] ?? 0,
          /* ★ 2026-10-21 用户「选中组后点不到组内节点」：RF 节点包装层默认
             pointer-events:all——组框体虽穿透，包装层仍拦截点击。组节点包装层
             置 none，标题栏/缩放条子元素自带 auto 恢复交互。 */
          style: t.kind === 'group' ? { pointerEvents: 'none' } : undefined,
          /* ★ 节点组初始/受控尺寸（RF v12 顶层 width/height——NodeResizer 原生改写
             的就是这两个字段，与缩放不再打架；持久化仍在 config.groupW/H） */
          ...(t.kind === 'group'
            ? {
                width: ((t.config as { groupW?: number }).groupW ?? 360),
                height: ((t.config as { groupH?: number }).groupH ?? 240),
              }
            : t.kind === 'note'
              ? {
                  /* ★ 2026-10-22 锤子便签式：宽度受控（拖右缘调），高度=内容自动撑开 */
                  width: ((t.config as { noteW?: number }).noteW ?? 190),
                }
              : {}),
          selected: isSel,
          // ★ 位置锁定（2026-10-07 用户指令；2026-10-26 统一推广全 kind）：
          //    config.posLocked（通用）|| ooPosLocked（OO 遗留兼容）→ 节点不可拖动
          draggable: !((t.config as { ooPosLocked?: boolean; posLocked?: boolean } | undefined)?.ooPosLocked
            || (t.config as { posLocked?: boolean } | undefined)?.posLocked),
          // ★ 节点组（2026-10-21）：只有标题栏能拖动整组——框体/边框留给缩放
          ...(t.kind === 'group' ? { dragHandle: '.group-node-head' } : {}),
          data: {
            ...existing?.data,
            name: ((t.config as { nodeNo?: string } | undefined)?.nodeNo ? ("[" + (t.config as { nodeNo?: string }).nodeNo + "] ") : "") + t.name,
            kind: t.kind,
            fields: t.fields.map((f) => ({ id: f.id, name: f.name, type: f.type })),
            boundHint: boundHintOf(t),
            textValue:
              t.kind === 'text'
                ? String(
                    (records[t.id]?.[0] as Record<string, unknown> | undefined)?.text ?? '',
                  ) + ((records[t.id]?.length ?? 0) > 1 ? `（共 ${records[t.id]?.length} 条）` : '')
                : undefined,
            textStyle: t.config?.textStyle,
            selected: isSel,
            running: false,
            // ★ 统一锁（2026-10-26）：任一锁定（数据/位置，含 OO 遗留字段）→ rf-node--locked
            //   （合成层常驻提升 will-change——锁定越多画布越流畅）+ 头部锁标
            locked: !!((t.config as { ooPosLocked?: boolean; posLocked?: boolean } | undefined)?.ooPosLocked
              || (t.config as { posLocked?: boolean } | undefined)?.posLocked
              || (t.config as { ooDataLocked?: boolean; dataLocked?: boolean } | undefined)?.ooDataLocked
              || (t.config as { dataLocked?: boolean } | undefined)?.dataLocked),
            // 规则跟随节点双击 → 成员管理面板。走 ref 读最新值（见 openRuleMembersRef 的注释）
            onOpenRuleMembers: openRuleMembersRef.current,
          },
        }
      })
    })
  }, [tables, positions, setNodes])

  // 运行高亮：根据 runningIds 单独更新节点 flag，不触发整体重建
  useEffect(() => {
    setNodes((nds) =>
      nds.map((n) => ({ ...n, data: { ...n.data, running: runningIds.includes(n.id) } })),
    )
  }, [runningIds, setNodes])

  // 跟踪 Shift 按键状态：用于「Shift+左键拖拽框选」
  useEffect(() => {
    const onDown = (e: KeyboardEvent) => { if (e.key === 'Shift') setShiftDown(true) }
    const onUp = (e: KeyboardEvent) => { if (e.key === 'Shift') setShiftDown(false) }
    const onBlur = () => setShiftDown(false)
    window.addEventListener('keydown', onDown)
    window.addEventListener('keyup', onUp)
    window.addEventListener('blur', onBlur)
    return () => {
      window.removeEventListener('keydown', onDown)
      window.removeEventListener('keyup', onUp)
      window.removeEventListener('blur', onBlur)
    }
  }, [])

  // 切换画布页后自动 fitView，保证新页内容立即可见（修复切换后内容时显时不显）
  useEffect(() => {
    if (!activePageId) return
    const t = setTimeout(() => {
      const inst = instRef.current
      if (inst && nodesRef.current.length) inst.fitView({ padding: 0.2, duration: 300, maxZoom: 1 })
    }, 90)
    return () => clearTimeout(t)
  }, [activePageId])

  // edges：lookup 字段派生 + 数据流连线
  //★ 派生逻辑已抽到core/projectFormat 的 edgesOf（工程文件落盘要用同一份）。
  //   这里只负责把 ProjectEdge 转成 React Flow 的 Edge —— 不再重复实现判定，
  //   否则两边漂移的症状是「保存后重开连线全没了」，且只在跨进程重开时暴露。
  //
  /*★★★ 铁律（2026-10-05 用户实测报回归后立的）：
   * 「点一下规则节点 → 其他工作流线全停，取消选中才恢复」。
   *
   * 根因：这个 useEffect 的依赖里曾放进去「哪些规则节点被选中」。
   * 选中一变 → setEdges(全表重建) → RF 收到一批**新边对象**，
   * 把每条边的 <path> 重挂 → CSS 动画**同步从头开始**。
   * 多条线一起重启，肉眼看就是「全停住了，然后一起重新走」。
   *
   * ⇒ 教训：**派生边样式时，别把「随选中变化的瞬态」放进重建依赖。**
   *   瞬态（选中/悬停/流动）一律走 CSS class，由祖先标记驱动；
   *   只有「边集合本身变了」（tables 改、lookup 增删）才重建边表。
   *   规则线的流动现在就是这么做的，见 rf-node 那段 rf-flowing。 */
useEffect(() => {
    const list: Edge[] = []
    for (const t of Object.values(tables)) {
      for (const pe of edgesOf(t, tables)) {
        /* 规则线要与数据流线**在样式上分开**（用户 2026-10-05 第 2 条：
         * 「主节点的样式和其他跟随节点的样式不一致」）。
         *
         * ★ 根因不在 CSS，在此之前：规则锁定关系存的是被锁节点 linkedInputs 里
         *   的规则节点 id，与数据流边**共用同一条 flow 边**—— 所以以前
         *   渲染时二者完全一样（都拿到 edge-flow 这个绿色虚线 class），
         *   根本没有"规则线样式"可言，自然也分不出主/从。
         *   判据：这条 flow 边的任一端是规则节点 ⇒ 它是规则线。
         *   （落库侧已保证规则线只会是「规则节点 ↔ 同类型普通节点」，
         *     所以这里不会与数据流线混淆。）
         */
        const srcT = tables[pe.source]
        const tgtT = tables[pe.target]
        /* 圆形节点专属线的线种判定（2026-10-06 抽离到 core/ruleSync.roundLineKindOf）：
         * flow 边任一端是圆形 kind ⇒ 它不是普通数据流线。
         * 判据与落库层（App.handleConnect 分支）同源：任一端是 rule → 规则线；
         * 任一端是 ai → AI 线；任一端是 paramlib → 参数库线。
         * ★ 与锚点落位（roundLineAnchors）同在ruleSync，两者可被同一条测试
         *   串起来断言「判定为 paramlib ⇒ 必落上1」，不必再靠人脑对齐两段代码。 */
        const lineKind = roundLineKindOf(pe.kind, srcT?.kind, tgtT?.kind)
        const isRuleLine = lineKind === 'rule'
        /* 主/从身份读**被锁节点自己**的 ruleRef（与内核 membersOf 同源，
         * 不另存一份「谁是主」——否则改主节点后画布线型不跟着变）。*/
        const other = srcT?.kind === 'rule' ? tgtT : srcT
        const isPrimary = isRuleLine && other?.config?.ruleRef?.mode === 'primary'
        /*规则线恒定带 edge-rule--flow；「此刻是否真的在动」交给 CSS 依据
         *祖先的 .rf-flowing 标记判断。点选不进这里 ⇒ 边对象不重建、
         * 其它线的动画不会被重挂重启（这正是上一版的回归根因）。
         * AI/参数库线同构：edge-ai/edge-paramlib + --flow，族色各随其 kind。*/
        const cls = lineKind === 'rule'
          ? isPrimary
            ? 'edge-rule edge-rule--primary edge-rule--flow'
            : 'edge-rule edge-rule--follow edge-rule--flow'
          : lineKind === 'ai'
            ? 'edge-ai edge-ai--flow'
            : lineKind === 'paramlib'
              ? 'edge-paramlib edge-paramlib--flow'
              : pe.kind === 'flow'
                ? 'edge-flow'
                : undefined
        const tag = lineKind === 'rule' ? (isPrimary ? '规则·主' : '规则·跟随') : pe.kind === 'flow' ? '数据流' : ''
        /* ★★★ 圆形节点专属线必须显式指定两端锚点，否则画出来会连到左1/右1。
         *
         * 根因（规则线 2026-10-05 实测教训，AI/参数库线同构）：React Flow 在
         * edge 上没拿到 sourceHandle/targetHandle 时，会挑节点的默认边 ——
         * 对我们这套口就是「左侧 target / 右侧 source」，也就是数据流口的位置。
         * 于是线虽然配色对了，**走的却是左1→右1** —— 与「连在规则口上」的
         * 口径直接冲突。之所以当年没暴露：配色和锚点位置是两件事，只验
         * stroke 就 PASS 了，从没量过边的实际起终点坐标。
         *   → 铁律：改连线外观时，**样式和锚点位置要分别断言**，
         *     前者对了不代表后者对。
         *
         * 各线种的两端（与 canConnectPorts 准入矩阵、App.handleConnect 分支三方一致）：
         *   规则线   ：rule 节点下1（RULE_PORT_OUT）→ 普通节点上2（RULE_PORT_IN_2）
         *     ——上2 有专属 id（2026-10-05 拆开；早年上1/上2 同 id，RF 取第一个
         *       匹配项，规则线永远落上1，上2 纯摆设）。
         *   AI 线    ：普通节点下1（RULE_PORT_OUT）→ AI 上沿（AI_PORT_IN）
         *   参数库线 ：参数库下沿（PARAMLIB_PORT_OUT）→ 普通节点上1（RULE_PORT_IN）
         *     —— 2026-10-06 用户改口径：「参数库连接节点 锚点 上1」，
         *       原为上2（RULE_PORT_IN_2）。★ 注意这与**规则线**（仍走上2）不同，
         *       两者落点刻意错开：规则线锁规则、参数库供参数，
         *       同落一个口会让两种线在同节点上重叠，视觉上分不清谁是谁。 */
        let sourceHandle: string | undefined
        let targetHandle: string | undefined
        if (lineKind) {
          // ★ 锚点落位已抽离到 core/ruleSync.roundLineAnchors（纯函数，有单测钉死）。
          //   此前这段是一串内联 if，没有任何测试能碰到它 ——
          //   改落点时全量测试依然全绿（改的是那条无人断言的 if）。
          //   口径：规则线→上2、AI 线→普通下1、参数库线→上1（用户 2026-10-06）。
          //   ★ 只传 srcKind（事实），不传「圆形节点是不是 source」这种推导布尔：
          //     推导留在被测函数里，换线种时不会在调用方悄悄算反。
          const anchors = roundLineAnchors(lineKind, srcT.kind)
          sourceHandle = anchors.sourceHandle
          targetHandle = anchors.targetHandle
        } else if (pe.kind === 'flow' && tables[pe.target]?.kind === 'pdfmerge') {
          /* ★ 2026-10-23 用户指令「锚点没有对照到每行」：PDF 合并节点的上游连线
             落到对应行内嵌的专属锚点（pm-in-<上游id>）——连线与行一一对照，
             行拖动换序时锚点随行走、连线端点自动跟随。 */
          targetHandle = pmRowHandleId(pe.source)
        }
        list.push({
          id: pe.id,
          source: pe.source,
          target: pe.target,
          sourceHandle,
          targetHandle,
          label: lineKind ? '' : pe.kind === 'flow' ? '数据流' : (tables[pe.source]?.fields.find((f) => f.id === pe.fieldId)?.name ?? ''),
          type: 'assoc',
          /* 圆形节点专属线（规则/AI/参数库）默认不流动：常驻动画会让人以为
           * 在"传输"，实际它们只做锁定/关联。只有**选中圆形节点时**才流动
           * （--flow 类 + CSS，见上方 flowing 判定）。
           * 走 RF 内建 animated 会带上自己的 dash 样式，和 primary 的实线冲突，
           * 所以动画统一在 CSS 里做。 */
          /* ★ 2026-10-24 性能：animated=每帧改 dashoffset→SVG 每帧 CPU 重绘——
             边多+平移时叠加成卡顿主因。静态虚线由 .edge-flow CSS 提供（视觉不变），
             GPU 合成负责平移，零动画开销。 */
          animated: false,
          className: cls,
          data: {
            onRemove: onEdgeRemove,
            label: isRuleLine ? undefined : pe.kind === 'flow' ? '数据流' : undefined,
            hovered: hoverEdge === pe.id,
          },
        })
      }
    }
    setEdges(list)
  }, [tables, onEdgeRemove, setEdges, hoverEdge])

  const buildDetach = (nodeId: string): { key: string; label: string; run: () => void }[] => {
    const items: { key: string; label: string; run: () => void }[] = []
    const t = tables[nodeId]
    if (!t) return items
    const flowSrc = t.config?.linkedInputs ?? []
    flowSrc.forEach((src) => {
      items.push({ key: `fin-${src}`, label: `拆离：${tables[src]?.name ?? src} → 本节点`, run: () => onDetachFlow(src, nodeId) })
    })
    for (const other of Object.values(tables)) {
      if (other.id === nodeId) continue
      if ((other.config?.linkedInputs ?? []).includes(nodeId)) {
        items.push({ key: `fout-${other.id}`, label: `拆离：本节点 → ${other.name}`, run: () => onDetachFlow(nodeId, other.id) })
      }
    }
    for (const other of Object.values(tables)) {
      other.fields.forEach((f) => {
        if (f.type === 'lookup' && f.lookup?.table === nodeId) {
          items.push({ key: `lkin-${other.id}.${f.id}`, label: `拆离：${other.name}.${f.name} → 本节点`, run: () => onDetachLookup(other.id, f.id) })
        }
      })
    }
    ;(t.fields ?? []).forEach((f) => {
      if (f.type === 'lookup' && f.lookup?.table) {
        items.push({
          key: `lkout-${f.id}`,
          label: `拆离：本节点.${f.name} → ${tables[f.lookup.table]?.name ?? f.lookup.table}`,
          run: () => onDetachLookup(nodeId, f.id),
        })
      }
    })
    return items
  }

  const selectedIds = () => nodesRef.current.filter((n) => n.selected).map((n) => n.id)

  const bringToFront = (id: string) => {
    const max = Object.values(zRef.current).reduce((m, v) => Math.max(m, v), 0)
    zRef.current[id] = max + 1
    setNodes((nds) => nds.map((n) => (n.id === id ? { ...n, zIndex: zRef.current[id] } : n)))
  }
  const sendToBack = (id: string) => {
    const min = Object.values(zRef.current).reduce((m, v) => Math.min(m, v), 0)
    zRef.current[id] = min - 1
    setNodes((nds) => nds.map((n) => (n.id === id ? { ...n, zIndex: zRef.current[id] } : n)))
  }

  const selectInRect = (x: number, y: number, w: number, h: number) => {
    const inst = instRef.current
    if (!inst) return
    setNodes((nds) =>
      nds.map((n) => {
        const tl = inst.flowToScreenPosition(n.position)
        const wn = n.measured?.width ?? 200
        const hn = n.measured?.height ?? 80
        const hit = tl.x < x + w && tl.x + wn > x && tl.y < y + h && tl.y + hn > y
        return { ...n, selected: hit, data: { ...n.data, selected: hit } }
      }),
    )
  }

  const onPointerDown = (e: React.PointerEvent) => {
    // 连线手柄：交给 React Flow 自己处理。若走下面的「选中节点」分支会触发 setNodes 重建，
    // 把 RF 正在建立的连线冲掉，导致「拖动 handle 连不上」。这里直接放行、不做任何选中/重建。
    if ((e.target as HTMLElement).closest('.react-flow__handle')) return
    if (e.button === 2) {
      const target = e.target as HTMLElement
      const nodeEl = target.closest('.react-flow__node')
      rightRef.current = {
        active: true,
        moved: false,
        startX: e.clientX,
        startY: e.clientY,
        onNode: !!nodeEl,
        nodeId: nodeEl?.getAttribute('data-id') || undefined,
      }
      try {
        e.currentTarget.setPointerCapture(e.pointerId)
      } catch {
        /* ignore */
      }
      return
    }
    if (e.button !== 0) return
    const target = e.target as HTMLElement

    // 普通左键（非框选/非 Shift）：在 pointerdown 阶段就立即选中节点。
    // 这样即便快速点击被 React Flow 误判为 1px 微拖拽而吞掉 onNodeClick，选中状态也一定生效。
    if (!marqueeMode && !e.shiftKey) {
      const nodeEl = target.closest('.react-flow__node')
      if (nodeEl) {
        const id = nodeEl.getAttribute('data-id') || undefined
        if (id) {
          const additive = e.ctrlKey || e.metaKey
          // 以「点击前」的选中态为基准，避免 React Flow 默认单击先取消其它节点导致 Ctrl 多选失效
          const prevSelMap = new Map(nodesRef.current.map((x) => [x.id, x.selected]))
          setNodes((nds) =>
            nds.map((nd) => {
              const wasSel = prevSelMap.get(nd.id) ?? false
              // 按住 Ctrl/Cmd：切换该节点选中态并保留其它已选（多选）；再点一次取消
              // 未按住：仅选中被点节点（最新一个），其余取消
              const sel = additive ? (nd.id === id ? !wasSel : wasSel) : nd.id === id
              return { ...nd, selected: sel, data: { ...nd.data, selected: sel } }
            }),
          )
          onSelect(id)
        }
        return
      }
      // 点击空白画布：清空选中（与「点击空白取消选中」的直觉一致）
      if (
        target.closest('.react-flow__pane') &&
        !target.closest('.canvas-tools') &&
        !target.closest('.canvas-topbar')
      ) {
        if (nodesRef.current.some((x) => x.selected)) {
          setNodes((nds) =>
            nds.map((nd) => ({ ...nd, selected: false, data: { ...nd.data, selected: false } })),
          )
          onSelect('')
        }
      }
      return
    }

    // 框选 / Shift 框选
    if (marqueeMode || e.shiftKey) {
      // 落在节点上：交给 React Flow 拖拽
      if (target.closest('.react-flow__node')) return
      // 落在工具栏/顶栏/菜单/计数浮标等 UI 覆盖层上：不启动框选，让按钮点击正常触发
      if (
        target.closest('.canvas-tools') ||
        target.closest('.canvas-topbar') ||
        target.closest('.ctx-menu') ||
        target.closest('.canvas-statbar') ||
        target.closest('.zoom-panel') ||
        target.closest('.trash-panel') ||
        target.closest('.react-flow__minimap')
      ) {
        return
      }
      leftRef.current = { active: true, moved: false, startX: e.clientX, startY: e.clientY }
      try {
        e.currentTarget.setPointerCapture(e.pointerId)
      } catch {
        /* ignore */
      }
    }
  }

  const onPointerMove = (e: React.PointerEvent) => {
    const l = leftRef.current
    if (l.active) {
      const dx = e.clientX - l.startX
      const dy = e.clientY - l.startY
      if (!l.moved && Math.hypot(dx, dy) > 5) l.moved = true
      if (l.moved) {
        const x = Math.min(l.startX, e.clientX)
        const y = Math.min(l.startY, e.clientY)
        setMarquee({ x, y, w: Math.abs(dx), h: Math.abs(dy) })
        selectInRect(x, y, Math.abs(dx), Math.abs(dy))
      }
      return
    }
    const r = rightRef.current
    if (!r.active) return
    const dx = e.clientX - r.startX
    const dy = e.clientY - r.startY
    if (!r.moved && Math.hypot(dx, dy) > 5) r.moved = true
    if (r.moved) {
      const x = Math.min(r.startX, e.clientX)
      const y = Math.min(r.startY, e.clientY)
      const w = Math.abs(dx)
      const h = Math.abs(dy)
      setMarquee({ x, y, w, h })
      selectInRect(x, y, w, h)
    }
  }

  const onPointerUp = (e: React.PointerEvent) => {
    const l = leftRef.current
    if (l.active) {
      try {
        e.currentTarget.releasePointerCapture(e.pointerId)
      } catch {
        /* ignore */
      }
      leftRef.current = { ...l, active: false }
      if (l.moved) setMarquee(null)
      return
    }
    const r = rightRef.current
    if (!r.active) return
    try {
      e.currentTarget.releasePointerCapture(e.pointerId)
    } catch {
      /* ignore */
    }
    rightRef.current = { ...r, active: false }
    if (r.moved) {
      setMarquee(null)
      return
    }
    /* ★ 2026-11-07 用户反馈「在画布上选文字后右键，复制到的是节点」：
       存在文字选区时，右键的语义是「复制这段文字」——不弹节点/画布自绘菜单，
       放行给主进程的原生「复制文字」菜单（见 main.ts context-menu）。 */
    if (window.getSelection()?.toString().trim()) { setMenu(null); return }
    if (r.onNode && r.nodeId) {
      onSelect(r.nodeId)
      setMenu({ x: Math.min(e.clientX, window.innerWidth - 190), y: Math.min(e.clientY, window.innerHeight - 260), kind: 'node', nodeId: r.nodeId })
    } else {
      const fp = instRef.current?.screenToFlowPosition?.({ x: e.clientX, y: e.clientY })
      setMenu({ x: Math.min(e.clientX, window.innerWidth - 190), y: Math.min(e.clientY, window.innerHeight - 260), kind: 'pane', flowX: fp?.x, flowY: fp?.y })
    }
  }

  useEffect(() => {
    if (!menu) return
    const onDown = (e: PointerEvent) => {
      const t = e.target as HTMLElement
      if (t.closest('.ctx-menu') || t.closest('.canvas-tools') || t.closest('.canvas-topbar')) return
      setMenu(null)
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setMenu(null)
    }
    window.addEventListener('pointerdown', onDown)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('pointerdown', onDown)
      window.removeEventListener('keydown', onKey)
    }
  }, [menu])

  // 键盘 Delete/Backspace 删除选中节点（输入框聚焦时不触发）
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Delete' && e.key !== 'Backspace') return
      const ae = document.activeElement as HTMLElement | null
      if (ae && (ae.tagName === 'INPUT' || ae.tagName === 'TEXTAREA' || ae.tagName === 'SELECT' || ae.isContentEditable)) return
      const ids = selectedIds()
      if (ids.length) {
        e.preventDefault()
        onDeleteSelectedNodes(ids)
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [nodes, selected, onDeleteSelectedNodes])

  // 框选模式中按 ESC 退出框选状态（关闭模式并清掉残留的框选矩形）
  useEffect(() => {
    if (!marqueeMode) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault()
        setMarqueeMode(false)
        setMarquee(null)
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [marqueeMode])

  // Ctrl+C / Ctrl+V 复制粘贴选中节点
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const ae = document.activeElement as HTMLElement | null
      if (ae && (ae.tagName === 'INPUT' || ae.tagName === 'TEXTAREA' || ae.tagName === 'SELECT' || ae.isContentEditable)) return
      if (!(e.ctrlKey || e.metaKey)) return
      const k = e.key.toLowerCase()
      /* ★ 2026-11-07 用户反馈「画布上选文字 Ctrl+C 复制不到」：
         画布节点/属性面板里的文字被选中时，Ctrl+C 是「复制文字」，不是「复制节点」。
         节点复制改由右键菜单或未选文字时按 Ctrl+C 触发，避免抢走系统剪贴板
         （抢走后系统剪贴板还是上一次的截图，Ctrl+V 就成了图片）。 */
      const selText = window.getSelection()?.toString().trim() ?? ''
      if (k === 'c' && selText) return
      if (k === 'c') {
        e.preventDefault()
        copySelection()
      } else if (k === 'v') {
        e.preventDefault()
        /* ★ 2026-10-22 用户指令「Ctrl 粘贴到光标位置；无光标位置时右下偏移」 */
        const m = lastMouseRef.current
        if (m) {
          const fp = instRef.current?.screenToFlowPosition?.({ x: m.x, y: m.y })
          if (fp) { pasteSelection({ x: fp.x, y: fp.y }); return }
        }
        pasteSelection()
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [tables, positions, onPasteNodes])

  /** 装剪贴板（2026-10-22 抽公共）：ids → 条目；组条目后紧跟其存活成员（covered 去重）。
   *  copySelection（多选/快捷键）与右键单节点「复制」共用——右键复制组若只装组框 1 项，
   *  pasteSelection 的 items 与展开后的 ids 数量不配对，成员落点全错（同 23:20 根因）。 */
  const copyIdsToClipboard = (wantIds: string[]) => {
    if (!wantIds.length) return
    const st = useBaseStore.getState()
    const picked = wantIds
      .map((id) => ({ id, table: st.tables[id], record: st.records[id] ?? [], pos: st.positions[id] }))
      .filter((it) => it.table)
    const covered = new Set(picked.filter((it) => it.table!.kind !== 'group').map((it) => it.id))
    const items: CopyItem[] = []
    for (const it of picked) {
      items.push(it as CopyItem)
      if (it.table!.kind === 'group') {
        const members = ((it.table!.config as { members?: string[] }).members ?? []) as string[]
        for (const m of members) {
          if (covered.has(m)) continue
          const t = st.tables[m]
          if (!t || t.deleted) continue
          covered.add(m)
          items.push({ id: m, table: t, record: st.records[m] ?? [], pos: st.positions[m] } as CopyItem)
        }
      }
    }
    clipboardRef.current = items
  }

  /* ★ 2026-10-22 用户指令「拖动节点 修改节点组大小 坐标等也在诊断日志里」：
     拖动结束（落库时）记一条——只在 dragstop 记，mousemove 帧不记（会刷屏） */
  const logDragMove = (updates: { id: string; pos: { x: number; y: number } }[]) => {
    if (!updates.length) return
    const first = updates[0]
    const nm = tables[first.id]?.name ?? first.id
    pushLog(updates.length === 1
      ? `拖动 ${nm} → (${first.pos.x}, ${first.pos.y})`
      : `拖动 ${updates.length} 个节点（${nm} 等）→ (${first.pos.x}, ${first.pos.y})`)
  }

  const copySelection = () => {
    copyIdsToClipboard(selectedIds())
  }

  /** 粘贴：at 给定（flow 坐标）时首节点贴点位、整组保持相对布局（PPT 式）；
   *  无 at 时右下偏移 40px（旧行为）。 */
  const pasteSelection = (at?: { x: number; y: number }, mode?: 'group' | 'group-empty' | 'flat') => {
    const items = clipboardRef.current
    if (!items.length) return
    const ids = onPasteNodes(items, mode)
    if (at && items.length) {
      const base = items[0].pos ?? { x: 0, y: 0 }
      const dx = at.x - base.x
      const dy = at.y - base.y
      ids.forEach((newId, i) => {
        const p = items[i].pos
        pendingPos.current[newId] = { x: Math.round((p?.x ?? 0) + dx), y: Math.round((p?.y ?? 0) + dy) }
      })
      return
    }
    /* ★ 2026-10-22 修「粘贴后成员还是进旧组」：偏移必须整组统一 dx=dy=40——
       原先 +40+i*30 随索引递增把成员挤出新组框（相对布局打散），
       散落的成员叠进旧组区域，下次拖动即被中心判定吸进旧组。 */
    ids.forEach((newId, i) => {
      const p = items[i].pos
      pendingPos.current[newId] = { x: (p?.x ?? 0) + 40, y: (p?.y ?? 0) + 40 }
    })
  }

  const isCentered = () => {
    const inst = instRef.current
    const wrap = wrapRef.current
    if (!inst || !wrap) return true
    if (!nodes.length) return true
    const { x, y, zoom } = inst.getViewport()
    const vpW = wrap.clientWidth / zoom
    const vpH = wrap.clientHeight / zoom
    const vcx = -x + vpW / 2
    const vcy = -y + vpH / 2
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity
    for (const n of nodes) {
      const w = n.measured?.width ?? 200
      const h = n.measured?.height ?? 80
      minX = Math.min(minX, n.position.x)
      minY = Math.min(minY, n.position.y)
      maxX = Math.max(maxX, n.position.x + w)
      maxY = Math.max(maxY, n.position.y + h)
    }
    const cx = (minX + maxX) / 2
    const cy = (minY + maxY) / 2
    const m = 30
    const within = minX >= -x + m && maxX <= -x + vpW - m && minY >= -y + m && maxY <= -y + vpH - m
    return within && Math.abs(cx - vcx) < 24 && Math.abs(cy - vcy) < 24
  }

  const handleFit = () => {
    const inst = instRef.current
    if (!inst) return
    fittingRef.current = true
    inst.fitView({ padding: 0.2, duration: 300, maxZoom: 1 })
    setNeedCenter(false)
    setTimeout(() => { fittingRef.current = false }, 420)
  }

  /**
   * 胶囊拖放落点插入（批次 12 · ①的拖拽入口）。
   *
   * 落点算法：用 caretRangeFromPoint 拿到指针所在位置的文本节点与字符偏移，
   * 再把「纯文本视图坐标」换算成「带标记原文坐标」（offsetAfterInsert）。
   * 命中的靶是 text 节点正文 [data-cap-drop]；不在靶上则忽略（不误插到别处）。
   */
  const insertCapsuleAtPointer = (mark: string, clientX: number, clientY: number) => {
    const doc = document as Document & {
      caretRangeFromPoint?: (x: number, y: number) => Range | null
      caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null
    }
    let range: Range | null = null
    try {
      range = doc.caretRangeFromPoint?.(clientX, clientY) ?? null
    } catch {
      /* 浏览器不支持时走下面的兜底 */
    }
    // 定位靶元素：★ 必须用 closest('[data-cap-drop]') —— range.startContainer
    //   往往是正文里的 <span>/文本节点，parentElement 不是靶本身，
    //   直接取 attribute 会拿到 null 而提前 return（这坑踩过一次）。
    const caretEl = range?.startContainer?.parentElement ?? null
    const host = (caretEl?.closest?.('[data-cap-drop]') ??
      document.elementFromPoint(clientX, clientY)?.closest?.('[data-cap-drop]')) as HTMLElement | null
    const nodeId = host?.getAttribute('data-cap-drop')
    if (!nodeId) return
    const store = useBaseStore.getState()
    const rows = store.records[nodeId] ?? []
    const cur = String((rows[0] as Record<string, unknown> | undefined)?.text ?? '')
    // 算纯文本视图下的插入位
    let plain = plainLength(cur)
    if (range && host?.contains(range.startContainer)) {
      // 取该点之前的纯文本字符数（把 range 之前的片段克隆出来量长度，标记会被 plainLength 折算）
      const holder = document.createElement('div')
      holder.appendChild(range.cloneContents())
      plain = plainLength(holder.textContent ?? '')
    }
    const at = offsetAfterInsert(cur, plain)
    const nextText = cur.slice(0, at) + mark + cur.slice(at)
    store.pushHistory()
    const nextRows: Row[] = rows.length
      ? (rows.slice() as Row[])
      : ([{ text: '' }] as Row[])
    nextRows[0] = { ...(nextRows[0] as Record<string, unknown>), text: nextText }
    store.setTableData(nodeId, [{ id: 'text', name: '文本', type: 'text' }], nextRows)
  }

  /**
   * 给所有 text 节点正文挂原生 drop 监听（批次 12 · ① 拖拽入口）。
   *
   * 为什么不用 React Flow 的 onDrop：那个 prop 挂在 RF 内部容器上并带自己的守卫，
   * 落点无法精确到「文字里的第几个字符」；而且面板胶囊的 drop 与节点拖放是两种语义，
   * 混在一个 handler 里会互相干扰。原生监听直接绑在 [data-cap-drop] 上，落点最准。
   *
   * 节点会随增删重渲染，故用事件委托挂在稳定的包装层上，而不是逐节点绑定。
   */
  useEffect(() => {
    const wrap = wrapRef.current
    if (!wrap) return
    const onDrop = (ev: DragEvent) => {
      const mark = ev.dataTransfer?.getData(CAP_DRAG_MIME)
      if (!mark) return
      const target = ev.target as HTMLElement | null
      const host = target?.closest?.('[data-cap-drop]') as HTMLElement | null
      if (!host) return
      ev.preventDefault()
      ev.stopPropagation() // 别让 RF 把这次 drop 当成"拖入新节点"
      insertCapsuleAtPointer(mark, ev.clientX, ev.clientY)
    }
    const onOver = (ev: DragEvent) => {
      if (!ev.dataTransfer?.types?.includes(CAP_DRAG_MIME)) return
      const host = (ev.target as HTMLElement | null)?.closest?.('[data-cap-drop]') as HTMLElement | null
      if (!host) return
      ev.preventDefault()
      ev.dataTransfer.dropEffect = 'copy'
      host.classList.add('rf-cap-drop-hot')
    }
    const onLeave = (ev: DragEvent) => {
      const host = (ev.target as HTMLElement | null)?.closest?.('[data-cap-drop]') as HTMLElement | null
      host?.classList.remove('rf-cap-drop-hot')
    }
    // ★ 必须在**捕获阶段**监听并stopPropagation：React Flow 在冒泡阶段接管 drop
    //   来实现「拖入新节点」，等冒泡到我们这儿时它已经把事办了。
    wrap.addEventListener('drop', onDrop, true)
    wrap.addEventListener('dragover', onOver, true)
    wrap.addEventListener('dragleave', onLeave, true)
    // 自检：挂上后打一条诊断，便于排查「监听没生效」这类问题
    try { (window as unknown as { __capDropHooked?: boolean }).__capDropHooked = true } catch { /* noop */ }
    return () => {
      wrap.removeEventListener('drop', onDrop, true)
      wrap.removeEventListener('dragover', onOver, true)
      wrap.removeEventListener('dragleave', onLeave, true)
    }
    // 依赖 nodes.length：首次挂载时 wrapRef 可能还没就绪，节点数变化时重试挂载
  }, [nodes.length, insertCapsuleAtPointer])

  const doUndo = () => {
    useBaseStore.getState().undo()
    setTimeout(() => setNeedCenter(!isCentered()), 80)
  }
  const doRedo = () => {
    useBaseStore.getState().redo()
    setTimeout(() => setNeedCenter(!isCentered()), 80)
  }

  // Ctrl/Cmd+Z 撤销，Ctrl/Cmd+Shift+Z 或 Ctrl+Y 重做
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const ae = document.activeElement as HTMLElement | null
      if (ae && (ae.tagName === 'INPUT' || ae.tagName === 'TEXTAREA' || ae.tagName === 'SELECT' || ae.isContentEditable)) return
      if (!(e.ctrlKey || e.metaKey)) return
      const k = e.key.toLowerCase()
      if (k === 'z') {
        e.preventDefault()
        // ★ 撤销/重做必须**同时同步 RF 内部**，否则画布纹丝不动：
        //   store 的 undo/redo 只改 store，而重建 effect 取位置时
        //   `existing?.position`永远优先（见 applyPositions 注释）⇒ 撤销形同虚设。
        //   这与「自动排列一直没生效」是同一个既有缺陷的另一个症状。
        if (e.shiftKey) syncRfFromStore(useBaseStore.getState().redo)
        else syncRfFromStore(useBaseStore.getState().undo)
      } else if (k === 'y') {
        e.preventDefault()
        syncRfFromStore(useBaseStore.getState().redo)
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])

  /* 批量落位：**RF 内部 + store 两侧都写**。
   *
   * ★ 这是本轮修掉的既有缺陷（2026-10-06浏览器验收抓出）：
   *   原来的 onApplyLayout 只写 store（push() + setNodePosition），
   *   但 Canvas 的重建 effect（见下）取位置是
   *       existing?.position ?? pending ?? positions[t.id] ?? ...
   *   ——`existing`（RF 内部节点对象）只要存在就永远优先，store 的新值被压掉。
   *   实测证据：把 store.positions 改成 {4000,4000}，DOM 纹丝不动。
   *   ⇒ 症状：「自动排列」这个早就存在的按钮其实**一直没生效**，
   *     而单测（测纯函数 layoutLayers）与它无关，全绿。
   *
   * pendingPos 只对**新建**节点有效（existing 不存在时才能被读到，见 onDrop），
   * 对已存在的节点无效 —— 所以必须直接 setNodes 写 RF 内部。
   *
   * 两边都要写，缺一不可：
   *   只写 RF→ 刷新/翻页后从 store 恢复时位置回退
   *   只写 store → 当前画布不动（本缺陷）
   */
  const applyPositions = (pos: Record<string, NodePos>) => {
    const ids = new Set(Object.keys(pos))
    setNodes((nds) =>
      nds.map((n) => (ids.has(n.id) ? { ...n, position: pos[n.id] } : n)),
    )
    onApplyLayout(pos)
    setTimeout(() => setNeedCenter(!isCentered()), 120)
  }

  const autoLayout = () => {
    applyPositions(layoutLayers(tables))
  }

  /* 自动对齐（2026-10-06 用户：「增加一个自动对齐按钮」）——
   *
   * ★ 与上面 autoLayout（自动排列）是**两个独立按钮、两套语义**（用户口径：
   *   「自动对齐 和 自动排列 两个按钮 一个是对齐 一个是彻底重排」）：
   *     自动对齐 = 网格归位，**保持**节点相对关系（只吸附+消重叠+拉直同行）
   *     自动排列 = 按上下游分层，**改变**相对关系
   *   本体在 core/align.alignLayout（纯函数，30 测试）。
   *
   * 两处必须注意：
   * ① 位置取 **nodesRef（React Flow 权威位置，含拖拽结果）**，不是 store 的 positions
   *    —— 后者可能落后于用户刚拖完的位置，据它算会把节点拽回旧处。
   * ② 尺寸取**实测 measured**（圆形约 64×64、常规约 320×120，差 5 倍）。
   *    不喂实测就会用 align 的保守默认 220×90 算间距，把窄节点推得太远
   *    ——「对齐」就变成「拉散」。measured 为 undefined（刚插入未渲染）时才回落默认。
   *    ★ measured.width/height 各自都是 `number | undefined`（@xyflow/system
   *      的定义如此），故用 ?? 回落；口径与既有 selectInRect / isCentered 一致。 */
  const autoAlign = () => {
    const live = nodesRef.current.filter((n) => !tables[n.id]?.deleted)
    if (!live.length) return
    const cur: Record<string, NodePos> = {}
    const size: Record<string, AlignSize> = {}
    for (const n of live) {
      cur[n.id] = n.position
      const mw = n.measured?.width ?? 200
      const mh = n.measured?.height ?? 90
      size[n.id] = { w: mw, h: mh }
    }
    // ★ 走 applyPositions：RF 内部 + store 两侧都写（只写 store 画布不动的根因，
    //   见 applyPositions 注释）。onApplyLayout 单独用是不生效的。
    applyPositions(alignLayout(live.map((n) => n.id), cur, size))
  }

  const runWorkflow = () => {
    const order = executionOrder(tables)
    if (!order.length) return
    setRunningIds([order[0]])
    order.slice(1).forEach((id, i) => setTimeout(() => setRunningIds((prev) => [...prev, id]), (i + 1) * 480))
    setTimeout(() => setRunningIds([]), order.length * 480 + 900)
  }

  const run = (action: () => void) => {
    /* ★ 2026-10-22 try/finally：action 抛错也必须关菜单——
       否则右键菜单残留（用户实测「粘贴之后右键的面板没消失」） */
    try {
      action()
    } finally {
      setMenu(null)
    }
  }

  /* ── 顶部 ribbon（CanvasRibbon）透传用的行为，全部复用既有逻辑，不另造状态 ── */

  /** 缩放组：直接走 instRef，行为与右下角 ZoomPanel 一致 */
  const zoomBy = (dir: 'in' | 'out') => {
    const inst = instRef.current
    if (!inst) return
    if (dir === 'in') inst.zoomIn({ duration: 200 })
    else inst.zoomOut({ duration: 200 })
  }
  const zoomReset = () => instRef.current?.zoomTo(1, { duration: 200 })
  const fitAll = () => instRef.current?.fitView({ padding: 0.18, duration: 200, maxZoom: 1 })

  /** 插入页：在视口中心（避开顶部 ribbon 下沿）建节点，走既有 onAddTableAt */
  const addNodeAtCenter = (kind: NodeKind) => {
    audit('画布', '新增节点', '类型=' + kind)
    const inst = instRef.current
    const el = wrapRef.current
    if (!inst || !el) {
      onAddTableAt({ x: 0, y: 0 }, kind)
      return
    }
    const r = el.getBoundingClientRect()
    const pos = inst.screenToFlowPosition({ x: r.left + r.width / 2, y: r.top + r.height / 2 + 72 })
    onAddTableAt(pos, kind)
  }

  const trashCount = Object.values(tables).filter((t) => t.deleted).length
  const detachItems = menu?.kind === 'node' && menu.nodeId ? buildDetach(menu.nodeId) : []

  return (
    <div
      ref={wrapRef}
      className={`rf-canvas${marqueeMode ? ' marquee-on' : ''}${connectKind === 'rule' ? ' rf-connecting' : ''}${connectKind === 'data' ? ' rf-connecting-data' : ''}`}
      /* ★ 有文字选区时放行 contextmenu（不 preventDefault），交给主进程弹「复制文字」 */
      onContextMenu={(e) => { if (!window.getSelection()?.toString().trim()) e.preventDefault() }}
      onPointerDown={onPointerDown}
      onPointerMove={onPointerMove}
      onPointerUp={onPointerUp}
      /* ★ 2026-10-25 用户指令「拖动 wfc 文件 提示舍弃原先的打开新的 显示载入进度」：
         .wfc 拖入画布 → 确认框 → 载入进度浮层 → 替换当前画布 */
      onDragOver={async (e) => {
        if (e.dataTransfer.types.includes('Files') || e.dataTransfer.types.includes('application/x-cb-pdf-hist') || e.dataTransfer.types.includes('application/x-cb-folder-file')) {
          e.preventDefault()
          try { e.dataTransfer.dropEffect = e.dataTransfer.types.includes('application/x-cb-pdf-hist') ? 'move' : 'copy' } catch { /* ignore */ }
        }
      }}
      onDrop={async (e) => {
        if (wfcLoadStep) { e.preventDefault(); return }
        const dt = e.dataTransfer
        /* ★ 2026-10-25 历史产物拖到画布空白 = 创建 PDF 节点并加载（用户指令） */
        if (dt.types.includes('application/x-cb-pdf-hist')) {
          e.preventDefault(); e.stopPropagation()
          try {
            const ref = JSON.parse(dt.getData('application/x-cb-pdf-hist')) as { key: string; name: string }
            const { getBlob } = await import('./core/pdfBlobStore')
            const bytes = await getBlobHist(ref.key)
            if (!bytes) { pushLog(`[PDF·历史] 字节缺失：${ref.name}`, 'err'); return }
            const st = useBaseStore.getState()
            const nodeEl = (e.target as HTMLElement).closest('.react-flow__node')
            const targetId = nodeEl?.getAttribute('data-id') ?? ''
            const targetT = targetId ? st.tables[targetId] : undefined
            if (targetT && targetT.kind === 'pdf') {
              if (!window.confirm(`将「${ref.name}」置入节点「${targetT.name}」的文件组？`)) return
              const pid2 = 'pd-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
              await putBlobStandalone('pdfitem-' + pid2, bytes)
              st.pushHistory()
              const cfgN = (targetT.config ?? {}) as { pdfItems?: { id: string; name: string; inDb?: boolean }[]; pdfFileOrder?: string[] }
              st.setNodeConfig(targetId, { pdfItems: [...(cfgN.pdfItems ?? []), { id: pid2, name: ref.name, inDb: true }], pdfFileOrder: [...(cfgN.pdfFileOrder ?? []), pid2] })
              pushLog(`[PDF·历史] 已置入节点「${targetT.name}」：${ref.name}`, 'ok')
              return
            }
            if (!window.confirm(`在画布上创建 PDF 节点并加载「${ref.name}」？`)) return
            const newId = 'pdf-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 6)
            const pos = (instRef.current?.screenToFlowPosition?.({ x: e.clientX, y: e.clientY })) ?? { x: e.clientX, y: e.clientY }
            const pid = 'pd-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
            await putBlobStandalone('pdfitem-' + pid, bytes)
            const st2 = useBaseStore.getState()
            st2.pushHistory()
            st.addTable({ id: newId, name: ref.name.replace(/\.pdf$/i, ''), kind: 'pdf', fields: [], config: { pdfItems: [{ id: pid, name: ref.name, inDb: true }], pdfFileOrder: [pid] } })
            st.setNodePosition(newId, pos)
            pushLog(`[PDF·历史] 已创建节点并加载：${ref.name}`, 'ok')
          } catch (eH) { pushLog(`[PDF·历史] 创建节点失败：${String((eH as Error)?.message ?? eH).slice(0, 80)}`, 'err') }
          return
        }
        /* ★ 2026-10-25 文件夹节点文件拖出（用户指令「支持拖动导入」）：
           读字节（folder:readFile 桥）→ PDF/OO 节点上=置入；画布空白=按扩展名建节点 */
        if (dt.types.includes('application/x-cb-folder-file')) {
          e.preventDefault(); e.stopPropagation()
          try {
            const ref = JSON.parse(dt.getData('application/x-cb-folder-file')) as { dir: string; name: string }
            const apiF = (globalThis as { api?: { folderReadFile?: (dir: string, name: string) => Promise<{ ok: boolean; bytes: Uint8Array; reason?: string }> } }).api
            if (!apiF?.folderReadFile) { pushLog('[文件夹] 读取桥不可用', 'err'); return }
            const r = await apiF.folderReadFile(ref.dir, ref.name)
            if (!r.ok || !r.bytes.byteLength) { pushLog('[文件夹] 读取失败：' + (r.reason ?? ref.name), 'err'); return }
            const bytes = r.bytes
            const stF = useBaseStore.getState()
            const ext = ref.name.toLowerCase().split('.').pop() ?? ''
            const nodeEl = (e.target as HTMLElement).closest('.react-flow__node')
            const targetId = nodeEl?.getAttribute('data-id') ?? ''
            const targetT = targetId ? stF.tables[targetId] : undefined
            if (targetT && targetT.kind === 'pdf' && ext === 'pdf') {
              /* 拖到 PDF 节点 = 置入文件组 */
              if (!window.confirm('将「' + ref.name + '」置入节点「' + targetT.name + '」？')) return
              const pid = 'pd-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
              await putBlobStandalone('pdfitem-' + pid, bytes)
              stF.pushHistory()
              const cfgP = (targetT.config ?? {}) as { pdfItems?: { id: string; name: string; inDb?: boolean }[]; pdfFileOrder?: string[] }
              stF.setNodeConfig(targetId, { pdfItems: [...(cfgP.pdfItems ?? []), { id: pid, name: ref.name, inDb: true }], pdfFileOrder: [...(cfgP.pdfFileOrder ?? []), pid] })
              pushLog('[文件夹] 已置入「' + targetT.name + '」：' + ref.name, 'ok')
            } else if ((ext === 'docx' || ext === 'xlsx') && (!targetT || targetT.kind === 'ooword' || targetT.kind === 'oosheet')) {
              /* docx/xlsx 拖到空白 = 建 OO 节点；拖到 OO 节点 = 替换内容 */
              const wantKind = ext === 'docx' ? 'ooword' : 'oosheet'
              if (targetT && (targetT.kind === 'ooword' || targetT.kind === 'oosheet')) {
                if (!window.confirm('将「' + ref.name + '」替换「' + targetT.name + '」的内容？')) return
                stF.pushHistory()
                stF.setNodeConfig(targetId, { ooTemplateName: ref.name, ooBytesDbId: await stashOoBytes('oobytes-' + targetId, bytes), ooTemplateId: 'tpl-' + Date.now().toString(36) })
                pushLog('[文件夹] 已替换「' + targetT.name + '」：' + ref.name, 'ok')
              } else {
                if (!window.confirm('创建' + (ext === 'docx' ? ' OO 文档' : ' OO 工作簿') + '节点并加载「' + ref.name + '」？')) return
                const newId = wantKind + '-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 6)
                const pos = (instRef.current?.screenToFlowPosition?.({ x: e.clientX, y: e.clientY })) ?? { x: e.clientX, y: e.clientY }
                stF.pushHistory()
                stF.addTable({ id: newId, name: ref.name.replace(/\.[^.]+$/, ''), kind: wantKind, fields: [], config: { ooTemplateName: ref.name, ooBytesDbId: await stashOoBytes('oobytes-' + newId, bytes), ooTemplateId: 'tpl-' + Date.now().toString(36) } })
                stF.setNodePosition(newId, pos)
                pushLog('[文件夹] 已创建节点并加载：' + ref.name, 'ok')
              }
            } else if (!targetT) {
              /* 其它类型拖空白 = 建 PDF 节点兜底 */
              if (!window.confirm('在画布上创建 PDF 节点并加载「' + ref.name + '」？')) return
              const newId = 'pdf-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 6)
              const pos = (instRef.current?.screenToFlowPosition?.({ x: e.clientX, y: e.clientY })) ?? { x: e.clientX, y: e.clientY }
              const pid = 'pd-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
              await putBlobStandalone('pdfitem-' + pid, bytes)
              stF.pushHistory()
              stF.addTable({ id: newId, name: ref.name.replace(/\.[^.]+$/, ''), kind: 'pdf', fields: [], config: { pdfItems: [{ id: pid, name: ref.name, inDb: true }], pdfFileOrder: [pid] } })
              stF.setNodePosition(newId, pos)
              pushLog('[文件夹] 已创建 PDF 节点：' + ref.name, 'ok')
            } else {
              pushLog('[文件夹] 文件类型(.' + ext + ')与目标节点不匹配——拖到空白处或对应类型节点', 'warn')
            }
          } catch (eF) { pushLog('[文件夹] 导入失败：' + String((eF as Error)?.message ?? eF).slice(0, 80), 'err') }
          return
        }
        /* ★ 2026-10-25 历史产物拖到已有 PDF 节点上 = 置入该节点（用户指令） */
        if (dt.types.includes('application/x-cb-pdf-hist')) {
          e.preventDefault(); e.stopPropagation()
          const st2 = useBaseStore.getState()
          try {
            const ref = JSON.parse(dt.getData('application/x-cb-pdf-hist')) as { key: string; name: string }
            const nodeEl = (e.target as HTMLElement).closest('.react-flow__node')
            const nodeId = nodeEl?.getAttribute('data-id') ?? ''
            const t = nodeId ? useBaseStore.getState().tables[nodeId] : undefined
            if (!t || t.kind !== 'pdf') { pushLog(`[PDF·历史] 请拖到 PDF 文件节点上置入，或拖到空白处创建新节点`, 'warn'); return }
            if (!window.confirm(`将「${ref.name}」置入节点「${t.name}」的文件组？`)) return
            const bytes = await getBlobHist(ref.key)
            if (!bytes) { pushLog(`[PDF·历史] 字节缺失：${ref.name}`, 'err'); return }
            const pid = 'pd-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
            await putBlobStandalone('pdfitem-' + pid, bytes)
            st2.pushHistory()
            const cfgN = (t.config ?? {}) as { pdfItems?: { id: string; name: string; inDb?: boolean }[]; pdfFileOrder?: string[] }
            const items = [...(cfgN.pdfItems ?? []), { id: pid, name: ref.name, inDb: true }]
            const order = [...(cfgN.pdfFileOrder ?? []), pid]
            const base = useBaseStore.getState()
            useBaseStore.setState({ tables: { ...base.tables, [nodeId]: { ...t, config: { ...cfgN, pdfItems: items, pdfFileOrder: order } } } })
            pushLog(`[PDF·历史] 已置入节点「${t.name}」：${ref.name}`, 'ok')
          } catch (eH) { pushLog(`[PDF·历史] 置入失败：${String((eH as Error)?.message ?? eH).slice(0, 80)}`, 'err') }
          return
        }
        /* ★ 2026-10-25 文件夹拖入画布（用户指令「文件夹拖进来怎么没反应」）：
           浏览器安全限制拿不到拖入目录的完整路径——建 folder 节点后弹系统选目录框补上 */
        if (dt.types.includes('Files')) {
          try {
            const it0 = Array.from(dt.items ?? []).find((x) => x.kind === 'file')
            const entry = (it0 as DataTransferItem | undefined)?.webkitGetAsEntry?.()
            if (entry?.isDirectory) {
              e.preventDefault(); e.stopPropagation()
              const stD = useBaseStore.getState()
              const apiD = (globalThis as { api?: { folderPickDir?: () => Promise<string | null> } }).api
              /* 目标判定：已有文件夹节点上释放 = 替换该节点目录；空白 = 建新节点 */
              const nodeElD = (e.target as HTMLElement).closest('.react-flow__node')
              const targetIdD = nodeElD?.getAttribute('data-id') ?? ''
              const targetT = targetIdD ? stD.tables[targetIdD] : undefined
              if (targetT && targetT.kind === 'folder') {
                if (!window.confirm('将文件夹节点「' + targetT.name + '」指向「' + entry.name + '」？\n（浏览器安全限制需在系统对话框中确认路径）')) return
                if (apiD?.folderPickDir) {
                  const d = await apiD.folderPickDir()
                  if (d) { stD.pushHistory(); stD.setNodeConfig(targetIdD, { folderPath: d }); pushLog('[文件夹] 节点已指向：' + d, 'ok') }
                }
                return
              }
              if (!window.confirm('检测到文件夹「' + entry.name + '」。\n创建文件夹节点并选择该目录？（浏览器安全限制需手动确认路径）')) return
              const newId = 'folder-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 6)
              const pos = (instRef.current?.screenToFlowPosition?.({ x: e.clientX, y: e.clientY })) ?? { x: e.clientX, y: e.clientY }
              stD.pushHistory()
              stD.addTable({ id: newId, name: entry.name, kind: 'folder', fields: [], config: { folderMode: 'read' } })
              stD.setNodePosition(newId, pos)
              pushLog('[文件夹] 已创建节点：' + entry.name + '——请选择目录', 'ok')
              if (apiD?.folderPickDir) {
                const d = await apiD.folderPickDir()
                if (d) stD.setNodeConfig(newId, { folderPath: d })
              }
              return
            }
          } catch { /* 非目录或检测失败→走原文件导入逻辑 */ }
        }
        if (!dt.types.includes('Files')) return
        const wfc = Array.from(dt.files).find((f) => f.name.toLowerCase().endsWith('.wfc'))
        if (!wfc) return
        e.preventDefault(); e.stopPropagation()
        const okGo = window.confirm(`将舍弃当前画布的全部内容，打开「${wfc.name}」。\n\n确定打开新的画布吗？（当前内容可用保存/另存为先行备份）`)
        if (!okGo) return
        /* ★ 计时起点 + 实时刷新 interval（完成/失败时清理） */
        wfcLoadStartRef.current = Date.now()
        setWfcLoadElapsed(0)
        const loadTimer = window.setInterval(() => setWfcLoadElapsed((Date.now() - wfcLoadStartRef.current) / 1000), 100)
        const stopLoadTimer = () => window.clearInterval(loadTimer)
        setWfcLoadStep({ label: '读取文件…', pct: 10 })
        try {
          await openCanvasFromWfcFile(wfc, (st) => setWfcLoadStep(st))
          setWfcLoadElapsed((Date.now() - wfcLoadStartRef.current) / 1000)
          setWfcLoadStep({ label: '载入完成', pct: 100 })
          ;(window as unknown as { __wfcLoadTimerClear?: () => void }).__wfcLoadTimerClear?.()
          setTimeout(() => setWfcLoadStep(null), 2500)
        } catch (eW) {
          ;(window as unknown as { __wfcLoadTimerClear?: () => void }).__wfcLoadTimerClear?.()
          setWfcLoadStep(null)
          window.alert(`打开 .wfc 失败：${eW instanceof Error ? eW.message : String(eW)}`)
        }
      }}
      /* 双击滚轮（中键）= 适应画布（2026-10-06 用户指令）。button===1 是中键，
         detail 是 mousedown 的同键连击计数（系统双击间隔内第二下 = 2）；
         preventDefault 顺手阻掉浏览器中键自动滚动。
         复用 fitAll（与右下「适应」按钮 / 顶栏「适应窗口」同一逻辑）。
         ★ 必须用 Capture（捕获阶段）：React Flow 底下的 d3-zoom 在 pane 上
           监听 mousedown 且 stopImmediatePropagation，冒泡链在 pane 就断了，
           普通 onMouseDown（冒泡委托到 root）永远收不到 —— 探针实测：
           capture 监听能抓到 detail=2，冒泡监听收不到任何事件。
         无冲突：onPointerDown 对非左键直接 return（见上）；RF 的
         zoomOnDoubleClick 只认左键双击，中键双击本就不触发它。 */
      onMouseDownCapture={(e) => {
        if (e.button === 1 && e.detail === 2) {
          e.preventDefault()
          fitAll()
        }
      }}
    >
      {wfcLoadStep && (
        <div style={{ position: 'absolute', bottom: 24, left: '50%', transform: 'translateX(-50%)', zIndex: 200, padding: '10px 16px', background: 'var(--ed-bg2, var(--c-bg1))', border: '1px solid var(--ed-border, var(--c-border))', borderRadius: 8, boxShadow: '0 6px 20px rgba(0,0,0,0.45)', minWidth: 300 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: 'var(--ed-text2, var(--c-text2))', marginBottom: 6 }}>
            <span>正在载入 .wfc 画布</span>
            <span>{wfcLoadStep.pct}% · {wfcLoadElapsed.toFixed(1)}s</span>
          </div>
          <div style={{ height: 6, borderRadius: 3, background: 'var(--c-bg3, #8882)', overflow: 'hidden' }}>
            <div style={{ width: wfcLoadStep.pct + '%', height: '100%', borderRadius: 3, background: 'var(--c-blue, #3b82f6)', transition: 'width .25s ease' }} />
          </div>
          <div style={{ fontSize: 11.5, color: 'var(--ed-text3, var(--c-text3))', marginTop: 5 }}>{wfcLoadStep.label}</div>
        </div>
      )}
      {/* ★ 2026-11-09 用户指令「运行的时候 页面最下方 画布居中 显示操作/运行状态 执行出现 执行完毕 淡隐」：
          与节点徽章同源（nodeRunStatus store），running 常驻 / ok 2.6s 淡隐 / error 6.5s 淡隐。 */}
      <RunStatusBar />
      {dropPick && createPortal(
        <DropChooser
          g={dropPick.g}
          rejected={dropPick.rejected}
          onPick={applyDropPick}
        />,
        document.body
      )}
      {/* ★ 2026-10-29 用户规则：右键 OO 节点「导入文档…」复用此隐藏选择框——导入目标记在 ooImportTargetRef */}
      <input
        ref={ooImportRef}
        type="file"
        accept=".docx,.xlsx"
        style={{ display: 'none' }}
        onChange={(e) => {
          const f = e.target.files && e.target.files[0]
          const id = ooImportTargetRef.current
          if (f && id) {
            const t = useBaseStore.getState().tables[id]
            void importOoFileToNode(t, [f])
          }
          e.target.value = ''
        }}
      />
      <ReactFlow
        className={'edge-anim-' + edgeAnim + (interacting ? ' canvas-interacting' : '')}
        onMoveStart={() => setInteracting(true)}
        nodes={nodes}
        edges={edges}
        onNodesChange={(rawChanges) => {
        /* ★ 性能 HUD：渲染变更计数（每秒采样显示） */
        changeCountRef.current++
          /* ★ 2026-10-20 用户「整个软件节点、画布显示效果略有模糊」：
             React Flow 用 transform: translate(x px, y px) 摆节点，坐标带小数 =
             亚像素渲染 → 文字/边发糊。这里把所有位置变更取整到整数像素。 */
          /* ★ 组内成员 左/上钳制（2026-10-21 用户「拖到左边和上边 就卡在节点组的
             左侧和上侧」）：成员位置变更直接钳在组的左/上边缘——对位置变更本身
             钳制，与指针拖拽零冲突不卡手；右/下仍由 onNodeDrag 实时扩展。 */
          onNodesChange(
            rawChanges.map((c) => {
              if (c.type !== 'position' || !c.position) return c
              let x = c.position.x
              let y = c.position.y
              const kind = tables[c.id]?.kind
              if (kind && kind !== 'group') {
                for (const g of Object.values(tables)) {
                  if (!g || g.kind !== 'group' || g.deleted) continue
                  const members = ((g.config as { members?: string[] }).members ?? []) as string[]
                  if (!members.includes(c.id)) continue
                  const gn = nodesRef.current.find((nd) => nd.id === g.id)
                  if (gn) {
                    x = Math.max(x, gn.position.x)
                    y = Math.max(y, gn.position.y + 34)
                  }
                  break
                }
              }
              return { ...c, position: { x: Math.round(x), y: Math.round(y) } }
            })
          )

          /* ★ 2026-10-23 性能（用户实测「拖动多了还是卡」二刀）：高刷鼠标
             125~1000Hz 拖动时每秒触发几百次 onNodesChange——每次直通 setNodes
             = 每秒上百次全画布渲染。rAF 合并：position/dimensions 类变更
             缓冲到下一动画帧一次性应用（每帧最多一次渲染）；
             select/remove 等低频变更直通不受影响。 */
          const hasStreaming = rawChanges.some((c) => c.type === 'position' || c.type === 'dimensions')
          if (!hasStreaming) { onNodesChange(rawChanges); return }
          pendingNodeChanges.current.push(...rawChanges)
          if (nodeChangeScheduled.current) return
          nodeChangeScheduled.current = true
          requestAnimationFrame(() => {
            nodeChangeScheduled.current = false
            const batch = pendingNodeChanges.current
            pendingNodeChanges.current = []
            if (batch.length) onNodesChange(batch)
          })
        }}
        onEdgesChange={onEdgesChange}
        nodeTypes={nodeTypes}
        edgeTypes={edgeTypes}
        // 选中已在 onPointerDown 阶段处理（pointerdown 比 onNodeClick 更可靠，不会被微拖拽吞掉）。
        // 这里仅同步右侧面板选中项，不再改 nodes 选中态，避免与 pointerdown 逻辑重复触发。
        onNodeClick={(_, n) => onSelect(n.id)}
        // 双击 = 进入「态 B 放大态」编辑器。只有 docx/xlsx 这类有编辑面的节点响应；
        // 其余节点双击不做事（保持与单击一致，避免误触弹出面板）。
        onNodeDoubleClick={(_, n) => {
          const kind = (n.data as TableNodeData | undefined)?.kind
          // 文本节点：双击直接编辑内容（prompt，回车保存）；note 已就地编辑
          if (kind === 'text') {
            const cur = String((n.data as TableNodeData).textValue ?? '')
            const next = window.prompt('便利贴内容：', cur)
            if (next !== null && next !== cur) {
              useBaseStore.getState().pushHistory()
              useBaseStore
                .getState()
                .setTableData(n.id, [{ id: 'text', name: 'text', type: 'text' }], [{ text: next }])
            }
            return
          }
          if (false) {
            const cur = String((n.data as TableNodeData).textValue ?? '')
            const next = window.prompt('文本内容（数字或一串文字）：', cur)
            if (next !== null && next !== cur) {
              useBaseStore.getState().pushHistory()
              useBaseStore
                .getState()
                .setTableData(n.id, [{ id: 'text', name: '文本', type: 'text' }], [{ text: next }])
            }
            return
          }
          /* ★ 2026-10-23 用户指令「双击节点打开加动效，像打开 app 一样苹果那种丝滑展开」：
             记录节点屏幕 rect → 编辑器挂载时从该位置放大展开（editorOpenFrom 暂存） */
          try {
            const el = document.querySelector(`.react-flow__node[data-id="${n.id}"]`) as HTMLElement | null
            const r = el?.getBoundingClientRect()
            if (r && r.width > 0) editorOpenFrom.rect = { x: r.x, y: r.y, w: r.width, h: r.height }
            else editorOpenFrom.rect = null
          } catch { /* rect 取不到就退普通打开，动效失败不影响功能 */ }
          if (kind && EDITABLE_KINDS.includes(kind)) onOpenNode?.(n.id)
          /* ★ 弹窗形态（2026-10-19）：常规数据节点双击 → 居中弹窗编辑器（OO 系交互） */
          else if (kind && hasPopupForm(kind)) onOpenPopupNode?.(n.id)
        }}
        onNodeDragStart={(_, n) => {
          setInteracting(true)
          dragSamplesRef.current = []
          /* ★ 拖动 LOD：仅 outline 模式打标记（wysiwyg=所见即所得不降级） */
          if (dragLodRef.current === 'outline') setNodes((nds) => nds.map((x) => x.id === n.id ? { ...x, data: { ...x.data, lodDragging: true } } : x))
          window.dispatchEvent(new CustomEvent('cb-drag-start'))
          dragStartPos.current[n.id] = n.position
          /* ★ 锁定组：拖起时若属于已锁定组，快照整组起始位置 */
          const g = resolveGrp(n.id)
          grpSnap.current = g
            ? { ids: g, base: Object.fromEntries(nodesRef.current.filter((x) => g.includes(x.id)).map((x) => [x.id, { ...x.position }])) }
            : null
        }}
        onNodeDrag={(_, n) => {
          /* ★ 自动采样：直接用 onNodeDrag 参数的 flow 坐标——零 DOM 读取
             （v1 用 querySelector+rect 每帧强制布局，layout thrashing 反而加重卡顿） */
          {
            const arrS = dragSamplesRef.current
            arrS.push({ t: performance.now(), x: Math.round(n.position.x), y: Math.round(n.position.y) })
            if (arrS.length > 3000) arrS.shift()
          }
          /* ★ 对齐磁吸（2026-10-21 用户指令「对节点、节点组 等类似ppt排版、ps对齐 一样
             带对齐的磁吸的追踪」）：拖动中与其它节点/组比对 左/中/右·上/中/下 九条对齐线，
             阈值内吸附并画追踪线；对组节点同样生效（成员随组移动）。 */
          if (alignMagnet) {
            /* ★ 对齐目标排除（2026-10-21 用户两条指令）：
               1. 拖节点组 → 组内成员不参与（成员随组等距移动，参与只是噪声线）
               2. 拖组成员 → 组节点不参与（不与组边缘产生磁吸对齐） */
            const draggedKind = tables[n.id]?.kind
            const groupMembers =
              draggedKind === 'group'
                ? (((tables[n.id]?.config ?? {}) as { members?: string[] }).members ?? [])
                : []
            const dragMemberOfGroup = Object.values(tables).some(
              (g) => g && g.kind === 'group' && !g.deleted && (((g.config as { members?: string[] }).members ?? []) as string[]).includes(n.id),
            )
            const dmeas = { width: n.measured?.width ?? 220, height: n.measured?.height ?? 120 }
            const d = {
              x0: n.position.x,
              xc: n.position.x + dmeas.width / 2,
              x1: n.position.x + dmeas.width,
              y0: n.position.y,
              yc: n.position.y + dmeas.height / 2,
              y1: n.position.y + dmeas.height,
            }
            const thr = 8 / (instRef.current?.getViewport?.().zoom || 1)
            type Cand = { delta: number; a: number; b0: number; b1: number }
            const V: Cand[] = []
            const H: Cand[] = []
            for (const o of nodesRef.current) {
              if (o.id === n.id || groupMembers.includes(o.id)) continue
              if (dragMemberOfGroup && o.data?.kind === 'group') continue
              const ow = o.measured?.width ?? 220
              const oh = o.measured?.height ?? 120
              const oo = {
                x0: o.position.x,
                xc: o.position.x + ow / 2,
                x1: o.position.x + ow,
                y0: o.position.y,
                yc: o.position.y + oh / 2,
                y1: o.position.y + oh,
              }
              for (const [dRef, oRef] of [
                [d.x0, oo.x0], [d.xc, oo.xc], [d.x1, oo.x1],
                [d.x0, oo.xc], [d.x1, oo.xc], [d.xc, oo.x0], [d.xc, oo.x1],
              ] as const) {
                const delta = oRef - dRef
                if (Math.abs(delta) <= thr) V.push({ delta, a: oRef, b0: Math.min(d.y0, oo.y0) - 12, b1: Math.max(d.y1, oo.y1) + 12 })
              }
              for (const [dRef, oRef] of [
                [d.y0, oo.y0], [d.yc, oo.yc], [d.y1, oo.y1],
                [d.y0, oo.yc], [d.y1, oo.yc], [d.yc, oo.y0], [d.yc, oo.y1],
              ] as const) {
                const delta = oRef - dRef
                if (Math.abs(delta) <= thr) H.push({ delta, a: oRef, b0: Math.min(d.x0, oo.x0) - 12, b1: Math.max(d.x1, oo.x1) + 12 })
              }
            }
            const near = (arr: Cand[]): Cand | null =>
              arr.reduce<Cand | null>((b, c) => (!b || Math.abs(c.delta) < Math.abs(b.delta) ? c : b), null)
            const bv = near(V)
            const bh = near(H)
            const dx = bv?.delta ?? 0
            const dy = bh?.delta ?? 0
            if (dx || dy) {
              /* ★ 2026-10-21 用户「磁吸在组内容易拖出来」：吸附位移先过组钳制——
                 组内成员吸附后的位置不得越过组左/上边缘（标题栏底线），
                 越界则按组边缘截断该轴的吸附。 */
              let nx2 = (n.position.x ?? 0) + dx
              let ny2 = (n.position.y ?? 0) + dy
              if (dragMemberOfGroup) {
                for (const g of Object.values(tables)) {
                  if (!g || g.kind !== 'group' || g.deleted) continue
                  const members2 = ((g.config as { members?: string[] }).members ?? []) as string[]
                  if (!members2.includes(n.id)) continue
                  const gn2 = nodesRef.current.find((x) => x.id === g.id)
                  if (gn2) {
                    /* 左/上：吸附位置不得越过组边缘（标题栏底线） */
                    nx2 = Math.max(nx2, gn2.position.x)
                    ny2 = Math.max(ny2, gn2.position.y + 34)
                    /* ★ 右/下：吸附位置超出组框 → 组框先长到位再吸附
                       （否则吸附位置在框外 8px，dragstop 中心判定会把成员移出组） */
                    const gW2 = ((g.config as { groupW?: number }).groupW ?? 360)
                    const gH2 = ((g.config as { groupH?: number }).groupH ?? 240)
                    const mw2 = n.measured?.width ?? 220
                    const mh2 = n.measured?.height ?? 120
                    const newW2 = Math.max(gW2, Math.round(nx2 + mw2 - gn2.position.x))
                    const newH2 = Math.max(gH2, Math.round(ny2 + mh2 - gn2.position.y))
                    if (newW2 > gW2 || newH2 > gH2) {
                      useBaseStore.getState().setNodeConfig(g.id, { groupW: newW2, groupH: newH2 })
                    }
                  }
                  break
                }
              }
              instRef.current?.updateNode?.(n.id, (nd: { position?: { x: number; y: number } }) => ({
                position: { x: nx2, y: ny2 },
              }))
            }
            const sig2 = (bv ? 'v' + bv.a + ':' + bv.b0 : '') + (bh ? 'h' + bh.a + ':' + bh.b0 : '')
            if (sig2 !== guideSigRef.current) {
              guideSigRef.current = sig2
              setGuides({
                v: bv ? [{ x: bv!.a, y0: bv!.b0, y1: bv!.b1 }] : [],
                h: bh ? [{ y: bv!.a, x0: bv!.b0, x1: bv!.b1 }] : [],
              })
            }
          }
          /* ★ 拖组内成员 右/下边缘出界 → 组框实时扩展（2026-10-21 用户指令）：
             仅向右/下生长（不虚增对侧、不改组位置）；左/上由 onNodesChange 钳制。 */
          {
            const draggedKind = tables[n.id]?.kind
            if (draggedKind && draggedKind !== 'group') {
              const st = useBaseStore.getState()
              for (const g of Object.values(st.tables)) {
                if (!g || g.kind !== 'group' || g.deleted) continue
                const members = ((g.config as { members?: string[] }).members ?? []) as string[]
                if (!members.includes(n.id)) continue
                const gn = nodesRef.current.find((x) => x.id === g.id)
                if (!gn) continue
                const gW = ((g.config as { groupW?: number }).groupW ?? 360)
                const gH = ((g.config as { groupH?: number }).groupH ?? 240)
                const mw = n.measured?.width ?? 220
                const mh = n.measured?.height ?? 120
                const newW = Math.max(gW, Math.round(n.position.x + mw - gn.position.x))
                const newH = Math.max(gH, Math.round(n.position.y + mh - gn.position.y))
                if (newW > gW || newH > gH) {
                  st.setNodeConfig(g.id, { groupW: newW, groupH: newH })
                }
                break
              }
            }
          }
          /* ★ 拖动节点悬停在组上 → 高亮提示「置入节点至节点组」（2026-10-21）；
             已是组成员的节点拖动不提示（已被组管理）。 */
          {
            const draggedKind = tables[n.id]?.kind
            const inAnyGroup = Object.values(tables).some(
              (g) => g && g.kind === 'group' && !g.deleted && (((g.config as { members?: string[] }).members ?? []) as string[]).includes(n.id),
            )
            if (draggedKind && draggedKind !== 'group' && !inAnyGroup) {
              const cx = n.position.x + (n.measured?.width ?? 220) / 2
              const cy = n.position.y + (n.measured?.height ?? 120) / 2
              let hit: string | null = null
              for (const g of Object.values(tables)) {
                if (!g || g.kind !== 'group' || g.deleted) continue
                const gn = nodesRef.current.find((x) => x.id === g.id)
                if (!gn) continue
                const gw = ((g.config as { groupW?: number }).groupW ?? 360)
                const gh = ((g.config as { groupH?: number }).groupH ?? 240)
                if (cx >= gn.position.x && cx <= gn.position.x + gw && cy >= gn.position.y && cy <= gn.position.y + gh) {
                  hit = g.id
                  break
                }
              }
              setDropHintId((prev) => (prev === hit ? prev : hit))
            }
          }
          /* ★ 锁定组：拖动中按同一增量平移其它成员（base 固定为拖起时位置）
             ★ 2026-10-23 修「拖组时成员跟随慢半拍」：跟随从 setNodes（React state
             ——落后被拖节点一个渲染调度周期，且每帧全数组 map 协调）改为
             instRef.updateNode 直更 RF 内部态——与组缩放推成员（0.00.31）同一通道，
             成员与组框同帧；最终位置仍经 change 事件回流 React state（dragstop 的
             nodesRef 读取不受影响，磁吸吸附已是同款先例）。 */
          const snap = grpSnap.current
          if (!snap || !snap.ids.includes(n.id)) return
          const start = snap.base[n.id]
          if (!start) return
          const dx = n.position.x - start.x
          const dy = n.position.y - start.y
          if (!dx && !dy) return
          const rf = instRef.current
          if (rf && typeof rf.updateNode === 'function') {
            for (const id of snap.ids) {
              if (id === n.id) continue
              const b = snap.base[id]
              if (!b) continue
              rf.updateNode(id, (nd: { position?: { x: number; y: number } }) => ({
                position: { x: b.x + dx, y: b.y + dy },
              }))
            }
            return
          }
          /* 兜底：无实例时退 React state 更新（旧行为） */
          setNodes((nds) => nds.map((x) => {
            if (x.id === n.id || !snap.ids.includes(x.id)) return x
            const b = snap.base[x.id]
            if (!b) return x
            return { ...x, position: { x: b.x + dx, y: b.y + dy } }
          }))
        }}
        onlyRenderVisibleElements
                onNodeDragStop={(_, n) => {
          if (n.data?.lodDragging) {
            setNodes((nds) => nds.map((x) => x.id === n.id ? { ...x, data: { ...x.data, lodDragging: false } } : x))
          }
          window.dispatchEvent(new CustomEvent('cb-drag-end'))
          /* ★ 自动分析：>33ms（掉帧）≥3 次才报告，流畅静默不打扰 */
          try {
            const arr2 = dragSamplesRef.current
            if (arr2.length >= 10) {
              const gaps: number[] = []
              for (let i3 = 1; i3 < arr2.length; i3++) {
                const g = arr2[i3].t - arr2[i3 - 1].t
                if (g > 4 && g < 500) gaps.push(g)
              }
              const avg = Math.round(gaps.reduce((a, b) => a + b, 0) / Math.max(1, gaps.length))
              const mx = gaps.length ? Math.max(...gaps) : 0
              const over33 = gaps.filter((g) => g > 33).length
              const over100 = gaps.filter((g) => g > 100).length
              if (over33 >= 3) {
                pushLog('[canvas] 拖动卡顿: ' + (n.data?.name || n.id) + ' 平均 ' + avg + 'ms/帧, 最大 ' + mx + 'ms, >33ms ' + over33 + ' 次, >100ms ' + over100 + ' 次', over100 > 0 ? 'err' : 'warn')
              }
            }
          } catch { /* ignore */ }
          dragSamplesRef.current = []
          /* ★ 2026-10-24 用户指令「每步操作记录 log：节点坐标/边框/锚点/连线位置」——拖停快照 */
          try {
            const s2fS = instRef.current?.screenToFlowPosition?.bind(instRef.current)
            const elN = document.querySelector(`.react-flow__node[data-id="${n.id}"]`) as HTMLElement | null
            if (elN && s2fS) {
              const rN = elN.getBoundingClientRect()
              const p1 = s2fS({ x: rN.left, y: rN.top })
              const p2 = s2fS({ x: rN.right, y: rN.bottom })
              const anchors = Array.from(elN.querySelectorAll('.react-flow__handle')).map((h) => {
                const hh = h as HTMLElement
                const hr = hh.getBoundingClientRect()
                const ap = s2fS({ x: hr.left + hr.width / 2, y: hr.top + hr.height / 2 })
                const hid = hh.getAttribute('data-handleid') || '?'
                return hid + ':(' + ap.x.toFixed(0) + ',' + ap.y.toFixed(0) + ')'
              }).join(' ')
              pushLog('[canvas] 拖停 ' + (n.data?.name || n.id) + ' 节点=(' + p1.x.toFixed(0) + ',' + p1.y.toFixed(0) + ')~(' + p2.x.toFixed(0) + ',' + p2.y.toFixed(0) + ') 边框=' + (p2.x - p1.x).toFixed(0) + 'x' + (p2.y - p1.y).toFixed(0) + ' 锚=[' + anchors + ']')
            }
          } catch { /* 快照失败不影响操作 */ }
          const start = dragStartPos.current[n.id]
          const moved = !start || start.x !== n.position.x || start.y !== n.position.y
          const snap = grpSnap.current
          grpSnap.current = null
          setDropHintId(null)
          guideSigRef.current = ''
          setGuides({ v: [], h: [] })
          /* ★ 拖节点进/出节点组（2026-10-21）：中心点落点判定——组框内=入组，组框外=移出。
             ★ 组框尺寸在成员拖动过程中恒定不变（用户 21:40 指令）；整理用「适应大小」钮。 */
          if (moved) {
            const draggedKind = tables[n.id]?.kind
            if (draggedKind && draggedKind !== 'group') {
              const cst = useBaseStore.getState()
              const cx = n.position.x + (n.measured?.width ?? 220) / 2
              const cy = n.position.y + (n.measured?.height ?? 120) / 2
              const gnode = new Map(nodesRef.current.map((x) => [x.id, x]))
              let changed = false
              for (const g of Object.values(cst.tables)) {
                if (!g || g.kind !== 'group' || g.deleted) continue
                const gn = gnode.get(g.id)
                if (!gn) continue
                const gw = ((g.config as { groupW?: number }).groupW ?? 360)
                const gh = ((g.config as { groupH?: number }).groupH ?? 240)
                const inside =
                  cx >= gn.position.x && cx <= gn.position.x + gw &&
                  cy >= gn.position.y && cy <= gn.position.y + gh
                const cur = ((g.config as { members?: string[] }).members ?? []) as string[]
                const has = cur.includes(n.id)
                if (inside && !has) {
                  cst.pushHistory()
                  cst.setNodeConfig(g.id, { members: [...cur, n.id] })
                  changed = true
                } else if (!inside && has) {
                  cst.pushHistory()
                  cst.setNodeConfig(g.id, { members: cur.filter((m) => m !== n.id) })
                  changed = true
                }
              }
              if (changed) audit('节点组', '成员调整', n.data.name)
            }
          }
          if (snap && snap.ids.includes(n.id)) {
            const updates = nodesRef.current
              .filter((x) => snap.ids.includes(x.id))
              .map((x) => ({ id: x.id, pos: { x: Math.round(x.position.x), y: Math.round(x.position.y) } }))
            useBaseStore.getState().pushHistory()
            logDragMove(updates)
            onNodeDragStopMulti(updates)
            return
          }
          if (!moved) return
          const sels = nodesRef.current.filter((x) => x.selected)
          const updates = sels.map((x) => ({ id: x.id, pos: { x: Math.round(x.position.x), y: Math.round(x.position.y) } }))
          useBaseStore.getState().pushHistory()
          logDragMove(updates)
          onNodeDragStopMulti(updates)
        }}
        onConnect={(conn: Connection) => {
          /* ★ 2026-10-25 用户指令「连上就是引用 不跳转不置入」：
             合并历史锚点连线 = 纯引用关系记录——不修改目标节点内容 */
          if (conn.source && conn.target && conn.sourceHandle?.startsWith('hist-')) {
            const stQ = useBaseStore.getState()
            const nmQ = (id: string) => (stQ.tables[id] ? stQ.tables[id].name + '(' + id + ')' : id)
            pushLog('[PDF·引用] ' + nmQ(conn.source) + ' → ' + nmQ(conn.target) + '（引用合并产物）', 'ok')
            audit('连线', '合并产物引用建立', nmQ(conn.source) + ' → ' + nmQ(conn.target))
          }
          /* ★ 2026-10-25 用户指令「替换页和被替换页每个词条前面设置锚点
             上游连接过来就默认集成上游的文件」：pr- 锚点连线 → 读上游文件
             字节 → 覆盖该词条（上游=PDF 节点首份/PDF 合并最新产物） */
          if (conn.source && conn.target && conn.targetHandle?.startsWith('pr-add-')) {
            /* ★ 2026-10-25 用户指令「在被替换和替换页行各增加个锚点 拖到这两个点都会加入到其中」：
               组级锚点 = 上游文件追加到该组末尾（新词条，非覆盖） */
            void (async () => {
              try {
                const stQ = useBaseStore.getState()
                const srcT = stQ.tables[conn.source]
                const tgtT = stQ.tables[conn.target]
                if (!srcT || tgtT?.kind !== 'pdfreplace') return
                const cfgS = (srcT.config ?? {}) as { pdfItems?: { id: string; name: string; b64?: string; inDb?: boolean }[]; pdfHistory?: { id: string; name: string; b64?: string; inDb?: boolean }[] }
                const pick = srcT.kind === 'pdf'
                  ? (cfgS.pdfItems ?? [])[0]
                  : srcT.kind === 'pdfmerge'
                    ? (cfgS.pdfHistory ?? [])[0]
                    : undefined
                if (!pick) { pushLog('[PDF·替换] 上游「' + srcT.name + '」没有可集成的 PDF 文件', 'err'); return }
                const upper = conn.targetHandle === 'pr-add-upper'
                const key = upper ? 'replaceTargets' : 'replaceSources'
                const st2 = useBaseStore.getState()
                st2.pushHistory()
                const cur = ((st2.tables[conn.target!]?.config as Record<string, unknown> ?? {})[key] as { id: string; name: string; b64?: string; inDb?: boolean; page: number }[]) ?? []
                st2.setNodeConfig(conn.target!, { [key]: [...cur, { id: pick.id, name: pick.name, inDb: true, page: 1 }] })
                pushLog(`[PDF·引用] 上游 PDF 已引用到${upper ? '被替换' : '替换页'}组位置：${pick.name}（第 ${cur.length + 1} 项）`, 'ok')
                audit('连线', '替换节点组级集成', pick.name + ' → ' + tgtT.name)
              } catch (eA) { pushLog('[PDF·替换] 追加失败：' + String((eA as Error)?.message ?? eA).slice(0, 80), 'err') }
            })()
            /* ★ 2026-10-25 修「连线后没有线」：此前的 return 吞掉了建线——
               集成词条与建立连线并行（async 集成不阻塞），线照常落 */
          }
          if (conn.source && conn.target && conn.targetHandle?.startsWith('pr-') && !conn.targetHandle.startsWith('pr-add-')) {
            void (async () => {
              try {
                const stQ = useBaseStore.getState()
                const srcT = stQ.tables[conn.source]
                const tgtT = stQ.tables[conn.target]
                if (!srcT || tgtT?.kind !== 'pdfreplace') return
                const cfgS = (srcT.config ?? {}) as { pdfItems?: { id: string; name: string; b64?: string; inDb?: boolean }[]; pdfHistory?: { id: string; name: string; b64?: string; inDb?: boolean }[] }
                const pick = srcT.kind === 'pdf'
                  ? (cfgS.pdfItems ?? [])[0]
                  : srcT.kind === 'pdfmerge'
                    ? (cfgS.pdfHistory ?? [])[0]
                    : undefined
                if (!pick) { pushLog('[PDF·替换] 上游「' + srcT.name + '」没有可集成的 PDF 文件', 'err'); return }
                const upper = conn.targetHandle!.startsWith('pr-upper-')
                const key = upper ? 'replaceTargets' : 'replaceSources'
                const listKey = (key === 'replaceTargets' ? (tgtT.config as { replaceTargets?: { id: string }[] }).replaceTargets : (tgtT.config as { replaceSources?: { id: string }[] }).replaceSources) ?? []
                const entryId = conn.targetHandle!.split('-').slice(2).join('-')
                const idx = listKey.findIndex((x) => x.id === entryId)
                if (idx < 0) { pushLog('[PDF·替换] 词条已不存在（锚点失效）', 'err'); return }
                const st2 = useBaseStore.getState()
                st2.pushHistory()
                const cur = ((st2.tables[conn.target!]?.config as Record<string, unknown> ?? {})[key] as { id: string; name: string; b64?: string; inDb?: boolean; page: number }[]) ?? []
                const next = cur.map((x) => (x.id === entryId ? { ...x, id: pick.id, name: pick.name, inDb: true, b64: undefined } : x))
                st2.setNodeConfig(conn.target!, { [key]: next })
                pushLog(`[PDF·替换] 上游已集成：${pick.name} → ${upper ? '被替换' : '替换页'}组第 ${idx + 1} 项`, 'ok')
                audit('连线', '替换节点上游集成', pick.name + ' → ' + tgtT.name)
              } catch (eP) { pushLog('[PDF·替换] 集成失败：' + String((eP as Error)?.message ?? eP).slice(0, 80), 'err') }
            })()
          }
          /* ★ 2026-10-26 PDF转Word 节点（仿替换页锚点语义）：
             p2w-add 组级锚点 = 上游 PDF 追加到待转组末尾；
             p2w-<词条id> 行级锚点 = 上游 PDF 覆盖该词条（引用 id/字节语义同替换页） */
          if (conn.source && conn.target && conn.targetHandle?.startsWith('p2w-')) {
            void (async () => {
              try {
                const stQ = useBaseStore.getState()
                const srcT = stQ.tables[conn.source]
                const tgtT = stQ.tables[conn.target]
                if (!srcT || tgtT?.kind !== 'pdf2word' || !conn.targetHandle) return
                const cfgS = (srcT.config ?? {}) as {
                  pdfItems?: { id: string; name: string; b64?: string; inDb?: boolean }[]
                  pdfHistory?: { id: string; name: string; b64?: string; inDb?: boolean }[]
                  replaceHistory?: { id: string; name: string; b64?: string; inDb?: boolean }[]
                }
                const pick = (srcT.kind === 'pdf' || srcT.kind === 'pdfslim')
                  ? (cfgS.pdfItems ?? [])[0]
                  : srcT.kind === 'pdfmerge'
                    ? (cfgS.pdfHistory ?? [])[0]
                    : srcT.kind === 'pdfreplace'
                      ? (cfgS.replaceHistory ?? [])[0]
                      : undefined
                if (!pick) { pushLog('[PDF转Word] 上游「' + srcT.name + '」没有可集成的 PDF 文件', 'err'); return }
                const st2 = useBaseStore.getState()
                st2.pushHistory()
                const cur = ((st2.tables[conn.target!]?.config as { p2wItems?: { id: string; name: string; b64?: string; inDb?: boolean }[] } ?? {}).p2wItems ?? [])
                if (conn.targetHandle === 'p2w-add') {
                  st2.setNodeConfig(conn.target!, { p2wItems: [...cur, { id: pick.id, name: pick.name, inDb: true }] })
                  pushLog(`[PDF转Word] 上游已追加到待转组：${pick.name}（第 ${cur.length + 1} 项）`, 'ok')
                } else {
                  const entryId = conn.targetHandle.slice(4)
                  if (!cur.some((x) => x.id === entryId)) { pushLog('[PDF转Word] 词条已不存在（锚点失效）', 'err'); return }
                  st2.setNodeConfig(conn.target!, { p2wItems: cur.map((x) => (x.id === entryId ? { id: pick.id, name: pick.name, inDb: true } : x)) })
                  pushLog(`[PDF转Word] 上游已覆盖词条：${pick.name}`, 'ok')
                }
                /* order 同步：新词条追加 order；覆盖时 order 不变 */
                if (conn.targetHandle === 'p2w-add') {
                  const ord = ((st2.tables[conn.target!]?.config as { p2wOrder?: string[] } ?? {}).p2wOrder ?? [])
                  st2.setNodeConfig(conn.target!, { p2wOrder: [...ord, pick.id] })
                }
              } catch (eW) { pushLog('[PDF转Word] 集成失败：' + String((eW as Error)?.message ?? eW).slice(0, 80), 'err') }
            })()
          }
          /* ★ 2026-10-25 用户令「只要pdf节点连接到pdf替换页节点 如果pdf节点带PDF文件
             则建立连接 文件名也同步落在pdf替换页节点」：连到节点级左1 口
             （targetHandle 为空）→ 上游 PDF 首份自动追加到被替换组（精确落组走锚点） */
          if (conn.source && conn.target && !conn.targetHandle) {
            const stQ = useBaseStore.getState()
            const srcT = stQ.tables[conn.source]
            const tgtT = stQ.tables[conn.target]
            if (srcT && tgtT?.kind === 'pdfreplace') {
              void (async () => {
                try {
                  const cfgS = (srcT.config ?? {}) as { pdfItems?: { id: string; name: string; b64?: string; inDb?: boolean }[]; pdfHistory?: { id: string; name: string; b64?: string; inDb?: boolean }[] }
                  const pick = srcT.kind === 'pdf'
                    ? (cfgS.pdfItems ?? [])[0]
                    : srcT.kind === 'pdfmerge'
                      ? (cfgS.pdfHistory ?? [])[0]
                      : undefined
                  if (!pick) { pushLog('[PDF·替换] 上游「' + srcT.name + '」没有可集成的 PDF 文件', 'err'); return }
                  const st2 = useBaseStore.getState()
                  st2.pushHistory()
                  const cur = ((st2.tables[conn.target!]?.config as Record<string, unknown> ?? {}).replaceTargets as { id: string; name: string; b64?: string; inDb?: boolean; page: number }[]) ?? []
                  st2.setNodeConfig(conn.target!, { replaceTargets: [...cur, { id: pick.id, name: pick.name, inDb: true, page: 1 }] })
                  pushLog(`[PDF·引用] 上游 PDF 已引用到被替换组：${pick.name}（第 ${cur.length + 1} 项）`, 'ok')
                  audit('连线', '替换节点默认集成', pick.name + ' → ' + tgtT.name)
                } catch (eD) { pushLog('[PDF·替换] 默认集成失败：' + String((eD as Error)?.message ?? eD).slice(0, 80), 'err') }
              })()
            }
          }
          if (conn.source && conn.target) {
            onConnect(conn.source, conn.target)
            /* ★ 2026-10-17 操作审计：连线是画布上最关键的改动（决定数据流向），
             *   记下来才能把「改了参数为什么下游没更新」和「线到底连没连上」对上。 */
            const nm = (id: string) => (tables[id] ? tables[id].name + '(' + id + ')' : id)
            audit('连线', '建立连线', nm(conn.source) + ' → ' + nm(conn.target) +
              (conn.sourceHandle || conn.targetHandle ? ' 端口=' + (conn.sourceHandle || '-') + '→' + (conn.targetHandle || '-') : ''))
            /* ★ 连线几何快照：延迟采样新边的 svg path（d=贝塞尔路径点） */
            setTimeout(() => {
              try {
                const es2 = instRef.current?.getEdges?.() ?? []
                const e2 = es2.find((x) => x.source === conn.source && x.target === conn.target)
                if (e2?.id) {
                  const pe = document.querySelector('.react-flow__edge[data-id="' + e2.id + '"] path')
                  if (pe) pushLog('[canvas] 连线路径 ' + nm(conn.source) + '→' + nm(conn.target) + ' d=' + (pe.getAttribute('d') || '').slice(0, 180))
                }
              } catch { /* ignore */ }
            }, 200)
          }
        }}
        /* 连线中 → 常规节点顶部的上1/上2 显形。
           用户口径（2026-10-05 本轮）：「上1 和 上2 在一般情况 会隐藏
           只有连接的时候 也就是拖拽连线的时候」—— 即常态隐藏、拖线时显形。
           RF 本身不把「连接进行中」落到任何 DOM 类名/属性上（实测验证），
           所以只能自己维护状态并打成 .rf-connecting，见 connecting state 的注释。

           ★ 为什么下1（规则节点出线口）**不**跟着隐藏：
             它是规则线唯一的起点，用户得先看得见它才知道从哪儿能拖。
             只隐藏落点侧、不隐藏起点侧，是「起点可见、落点随拖线浮现」。 */
        /* 连线显隐（2026-10-06 用户口径二修）：
           「从节点左边 接 节点右边的锚点连线的时候 上下的锚点不用显示」——
           起点是规则口（下1 出线口）→ 拖的是规则线 → 上1/上2 显形（rf-connecting）；
           起点是数据流口（左/右圆口）→ 拖的是数据流线 → 上/下口全部不显示
           （rf-connecting-data，CSS 强藏）。显隐判据在起手那一刻就定，
           不等落到目标口。RF 本身不把「连接进行中」落到任何 DOM 类名/属性上
           （实测验证），所以只能自己维护状态打类。

           ★ 为什么下1（规则节点出线口）在规则拖线时**不**跟着隐藏：
             它是规则线唯一的起点，用户得先看得见它才知道从哪儿能拖。
             数据流拖线时它也被藏（rf-connecting-data 规则）——那时它没有意义。 */
        onConnectStart={(_, { handleId }) =>
          setConnectKind(isRulePortId(handleId ?? '') ? 'rule' : 'data')
        }
        onConnectEnd={() => setConnectKind(null)}
        /* 指令①的写入端：选中集合变化 → 取其中被选中的规则节点写入 store 瞬态桥
           （selRuleNode）。成员节点的 TableNode 读它给规则口挂 rf-rule-ports-show。
           ★ kind 从 tables 查而不是从 node.type：node.type 恒为 'table'
             （nodeTypes 只注册了一种渲染器），节点身份只认 schema.kind。
           ★ 属性列表里的注释只能用不带花括号的块注释：花括号开头会被
             JSX 解析成 spread 属性（待展开对象），直接语法错误 TS1005。 */
        onSelectionChange={(sel) => {
          const rule = sel.nodes.find((n) => tables[n.id]?.kind === 'rule')
          useBaseStore.getState().setSelRuleNode(rule ? rule.id : null)
          /* ★ 组 z 跟随选中（2026-10-22「拉出去拉回来边界老变」）：
             builder 只在 tables 变化时重建——点击选组不触发重建 → z 不切换。
             这里实时同步：组选中→1（手柄可拖），取消→-1（成员清晰）。 */
          setNodes((nds) =>
            nds.map((n) => {
              if (n.data?.kind !== 'group') return n
              const want = n.selected ? 1 : -1
              return n.zIndex === want ? n : { ...n, zIndex: want }
            }),
          )
        }}
        /* ★★ 规则线只走「规则口↔规则口」（用户：「规则线只能连接常规节点顶部的点」）
         *
         * 修掉的是真缺陷，不是洁癖：React Flow Strict 模式只禁同类型对同类型，
         * 规则口（target）和左侧数据流口（target）在它眼里毫无区别，
         * 于是「规则节点下方口 → 普通节点左侧口」是**合法**的，
         * 且会真的写进 inputs/linkedInputs —— 规则节点被记成了「输出的上游」。
         * 反向（普通节点左侧口 → 规则节点下方口）同样合法。
         *
         * 判据是 handle id 而不是节点 kind：光看 kind 拦不住「落在旁边的数据流口」
         * 这类错连，必须精确到「这个口是不是规则口」。
         */
        isValidConnection={(c) =>
          canConnectPorts(c.source, c.target, c.sourceHandle, c.targetHandle, tables)
        }
        onEdgeMouseEnter={(_, e) => {
          if (edgeLeaveTimer.current) clearTimeout(edgeLeaveTimer.current)
          setHoverEdge(e.id)
        }}
        onEdgeMouseLeave={(_, e) => {
          if (edgeLeaveTimer.current) clearTimeout(edgeLeaveTimer.current)
          edgeLeaveTimer.current = setTimeout(() => {
            setHoverEdge((c) => (c === e.id ? null : c))
          }, 160)
        }}
        onInit={(inst) => {
          instRef.current = inst
        }}
        onMoveEnd={() => {
          setInteracting(false)
          const inst = instRef.current
          if (inst) setZoomPct(Math.round(inst.getZoom() * 100))
          if (!fittingRef.current) setNeedCenter(!isCentered())
        }}
        onDrop={async (e) => {
          e.preventDefault()
          const inst = instRef.current
          if (!inst) return
          // 节点拖放建新节点（胶囊拖放由 text 正文自己的原生 drop 处理，见下方 useEffect）
          /* ★ 2026-10-23 修「拖行排序总是创建新表」：行拖拽（PDF 合并节点的
             排序 HTML5 dnd）在行外松手会落到画布 pane——dataTransfer 里没有
             'application/reactflow'（那是调色板拖出的标记），原 fallback 'table'
             兜底直接建了新表。非调色板来源一律不建节点。 */
          const rawKind = e.dataTransfer.getData('application/reactflow') as NodeKind | ''
          /* ★ 2026-10-23 用户指令「拖拽到空白位置则根据文件类型给提示，是否添加节点并导入」：
             非调色板来源若有系统文件 → 弹确认按类型建节点导入（照片/Word/Excel）；
             其余（行拖拽排序等无文件来源）维持忽略。 */
          if (!rawKind) {
            const files = e.dataTransfer.files
            if (files && files.length) {
              /* ★ 2026-10-25 PDF 替换页节点拖入（用户指令「上面被替换/下面替换页」）：
                 PDF 文件落到 pdfreplace 节点上 → 按鼠标 Y 判上/下组（上=被替换/下=替换页） */
              const nodeElR = (e.target as HTMLElement).closest('.react-flow__node')
              const rid = nodeElR?.getAttribute('data-id') ?? ''
              const rt = rid ? useBaseStore.getState().tables[rid] : undefined
              const pdfFiles = Array.from(files).filter((f) => (f.name.toLowerCase().split('.').pop() ?? '') === 'pdf')
              /* ★ 2026-10-29 用户规则：OO 节点（尤其空白节点）上松手拖入 docx/xlsx → 直接作为该节点文档，
                 不再弹通用选择框（通用框用于空白画布/其它节点）。空白直接导入；已有内容询问替换。 */
              if (nodeElR && rt && (rt.kind === 'ooword' || rt.kind === 'oosheet')) {
                const ooFiles = Array.from(files).filter((f) => /\.(docx|xlsx)$/i.test(f.name))
                if (ooFiles.length) {
                  e.preventDefault()
                  await importOoFileToNode(rt, ooFiles)
                  return
                }
              }
              /* ★ 2026-10-26 用户指令「PDF拖入空白区域弹选择对话框」+「其他格式同样参照该规范」：
                 全格式拖入空白区 → 分类后弹选择对话框（各类型建对应节点 + 误操作取消）——
                 仅空白区；节点上松手仍走各节点组件内处理 */
              if (!nodeElR) {
                const dp0 = inst.screenToFlowPosition({ x: e.clientX, y: e.clientY })
                const cls = classifyDrop(files)
                const hasAny = cls.g.photos.length || cls.g.words.length || cls.g.sheets.length || cls.g.pdfs.length
                if (hasAny) {
                  setDropPick({ g: cls.g, rejected: cls.rejected, pos: { x: Math.round(dp0.x), y: Math.round(dp0.y) } })
                } else {
                  pushLog(`拖入文件均不支持导入：${cls.rejected.slice(0, 4).join('、')}${cls.rejected.length > 4 ? ' 等' : ''}（支持 jpg/png、docx、xlsx、pdf）`, 'warn')
                }
                return
              }
              if (rt && rt.kind === 'pdfreplace' && pdfFiles.length) {
                const rect = nodeElR!.getBoundingClientRect()
                const toUpper = e.clientY < rect.top + rect.height / 2
                const key = toUpper ? 'replaceTargets' : 'replaceSources'
                const label = toUpper ? '被替换' : '替换页'
                const stR = useBaseStore.getState()
                stR.pushHistory()
                const cfgR = (rt.config ?? {}) as { replaceTargets?: unknown[]; replaceSources?: unknown[] }
                const cur = (key === 'replaceTargets' ? cfgR.replaceTargets : cfgR.replaceSources) ?? []
                const adds = await Promise.all(pdfFiles.map(async (f) => {
                  const bid = 'reprepl-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
                  const bytes = new Uint8Array(await f.arrayBuffer())
                  await putBlobStandalone(bid, bytes)
                  return { id: bid, name: f.name, inDb: true, page: 1 }
                }))
                stR.setNodeConfig(rid, { [key]: [...cur, ...adds] })
                pushLog(`[PDF·替换] 已加入${label}组：${pdfFiles.map((f) => f.name).join('、')}`, 'ok')
                return
              }
              /* ★ 2026-10-26 落在未自处理节点上（如数据表）→ 同样走选择对话框（统一规范） */
              const dropPos = inst.screenToFlowPosition({ x: e.clientX, y: e.clientY })
              const cls2 = classifyDrop(files)
              const hasAny2 = cls2.g.photos.length || cls2.g.words.length || cls2.g.sheets.length || cls2.g.pdfs.length
              if (hasAny2) {
                setDropPick({ g: cls2.g, rejected: cls2.rejected, pos: { x: Math.round(dropPos.x), y: Math.round(dropPos.y) } })
              } else {
                pushLog(`拖入文件均不支持导入：${cls2.rejected.slice(0, 4).join('、')}${cls2.rejected.length > 4 ? ' 等' : ''}（支持 jpg/png、docx、xlsx、pdf）`, 'warn')
              }
            }
            return
          }
          const kind = rawKind
          const raw = inst.screenToFlowPosition({ x: e.clientX, y: e.clientY })
          const pos = { x: Math.round(raw.x), y: Math.round(raw.y) }
          const newId = onAddTableAt(pos, kind)
          pendingPos.current[newId] = pos
          /* ★ 2026-10-23 参数库条目拖出：带库名标记 → 建后 rename（否则落默认名「参数库」） */
          const libName = e.dataTransfer.getData('application/x-paramlib-name')
          if (libName) useBaseStore.getState().renameTable(newId, libName)
          /* ★ 2026-10-19 用户指令：圆形节点（AI/规则跟随/参数库）拖到**常规节点上**释放
           *   → 直接连线（省掉从锚点拉线），并按连线语义放到常规节点的相对位置：
           *     · AI（普通节点下1 → AI）→ 放到常规节点**下方**
           *     · 参数库（参数库下沿 → 常规上1）→ 放到常规节点**上方**
           *     · 规则跟随（下1出线 → 常规顶部）→ 放到常规节点**上方偏左**
           *   连线方向走 onConnect（App.handleConnect 语义分流：linkedInputs/规则锁定）。 */
          if (isRoundNodeKind(kind)) {
            const hit = nodes.find((n) => {
              if (n.id === newId) return false
              const k = tables[n.id]?.kind
              if (!k || isRoundNodeKind(k) || k === 'start') return false
              const w = n.measured?.width ?? 220
              const h = n.measured?.height ?? 120
              return pos.x >= n.position.x && pos.x <= n.position.x + w && pos.y >= n.position.y && pos.y <= n.position.y + h
            })
            if (hit) {
              const tp = hit.position
              /* ★ 2026-10-19 用户指令「根据节点锚点来定相对位置」：
               *   不再用写死偏移——直接读 RF 内部的 handleBounds（锚点在节点内的
               *   真实坐标），把圆形节点摆到它要连的那个锚点旁边：
               *     · AI（连常规节点下1=底部 source 口）→ 摆在**下1 正下方**
               *     · 参数库/规则跟随（连上1=顶部 target 口）→ 摆在**上1 正上方**
               *   读不到锚点时退回写死偏移兜底。 */
              type HB = { id?: string; position: string; x: number; y: number; width: number; height: number }
              /* RF v12：handleBounds 直接挂在 node 对象上（测量后由 RF 填充） */
              const nint = hit as unknown as
                { handleBounds?: { source?: HB[]; target?: HB[] } } | undefined
              const pickHB = (arr?: HB[], pos?: string): HB | undefined =>
                (arr ?? []).find((h) => (pos ? h.position === pos : true))
              const b1 = pickHB(nint?.handleBounds?.source, 'bottom') // 常规节点 下1
              const t1 = pickHB(nint?.handleBounds?.target, 'top') // 常规节点 上1
              let rel: { x: number; y: number }
              if (kind === 'ai' && b1) {
                rel = { x: tp.x + b1.x + b1.width / 2 - 45, y: tp.y + b1.y + b1.height + 30 }
              } else if (kind !== 'ai' && t1) {
                rel = { x: tp.x + t1.x + t1.width / 2 - 45, y: tp.y + t1.y - 110 }
              } else if (kind === 'ai') {
                rel = { x: tp.x + ((hit.measured?.width ?? 220) - 90) / 2, y: tp.y + (hit.measured?.height ?? 120) + 40 }
              } else {
                rel = kind === 'paramlib' ? { x: tp.x + 30, y: tp.y - 150 } : { x: tp.x - 50, y: tp.y - 150 }
              }
              pendingPos.current[newId] = rel
              useBaseStore.getState().setNodePosition(newId, rel)
              /* ★ App 侧 handleConnect 读的 tables 闭包在本 tick 还是旧的
                 （刚 addTable 的新节点不在里面 ⇒ src/tgt 查不到 → 静默 return，
                 AI 能连是碰上状态已落定的情形）——延一拍等状态落定再连线。 */
              setTimeout(() => {
                try {
                  if (kind === 'ai') onConnect(hit.id, newId)
                  else onConnect(newId, hit.id)
                } catch (eC0) { console.warn('[drop-connect]', eC0) }
              }, 0)
              audit('连线', '拖放直连', (tables[newId]?.name || kind) + ' → ' + (tables[hit.id]?.name || hit.id))
            }
          }
          /* ★ 落点在节点组框内 → 自动入组（2026-10-21 用户指令）。
             以落点（节点左上角）判定；组矩形 = 组节点位置 + config.groupW/groupH。 */
          {
            const st = useBaseStore.getState()
            for (const g of Object.values(st.tables)) {
              if (!g || g.kind !== 'group' || g.deleted) continue
              const gn = nodesRef.current.find((x) => x.id === g.id)
              if (!gn) continue
              const gw = ((g.config as { groupW?: number }).groupW ?? 360)
              const gh = ((g.config as { groupH?: number }).groupH ?? 240)
              const inside =
                pos.x >= gn.position.x && pos.x <= gn.position.x + gw &&
                pos.y >= gn.position.y && pos.y <= gn.position.y + gh
              if (inside) {
                const cur = ((g.config as { members?: string[] }).members ?? []) as string[]
                if (!cur.includes(newId)) st.setNodeConfig(g.id, { members: [...cur, newId] })
                /* ★ 自动扩展边界（2026-10-21 用户指令）：统一 groupFitRect 约定——
                   新节点矩形（未测量先估 220×120）适配后与现有组框取并集 */
                {
                  const gW = ((g.config as { groupW?: number }).groupW ?? 360)
                  const gH = ((g.config as { groupH?: number }).groupH ?? 240)
                  const cur = { x: gn.position.x, y: gn.position.y, w: gW, h: gH }
                  /* ★ 拖入新节点与「适应大小」同一适配约定（groupFitRect，含内边距）——
                     终点几何与按钮整理结果一致（2026-10-21 用户「同样的逻辑适配」） */
                  const fit = groupFitRect({ x: pos.x, y: pos.y, w: 220, h: 120 })
                  const u = unionRect(cur, fit)
                  if (u.w > gW + 0.5 || u.h > gH + 0.5 || u.x < cur.x - 0.5 || u.y < cur.y - 0.5) {
                    st.setNodePosition(g.id, { x: u.x, y: u.y })
                    st.setNodeConfig(g.id, { groupW: u.w, groupH: u.h })
                  }
                }
                break
              }
            }
          }
          setDropHintId(null)
        }}
        onDragOver={(e) => {
          e.preventDefault()
          /* ★ 调色板拖节点悬停在节点组上 → 同样高亮提示「置入节点至节点组」
             （2026-10-21 用户指令；释放落组逻辑在下方 onDrop 的入组判定） */
          {
            const inst = instRef.current
            if (!inst) return
            const raw = inst.screenToFlowPosition({ x: e.clientX, y: e.clientY })
            let hit: string | null = null
            for (const g of Object.values(tables)) {
              if (!g || g.kind !== 'group' || g.deleted) continue
              const gn = nodesRef.current.find((x) => x.id === g.id)
              if (!gn) continue
              const gw = ((g.config as { groupW?: number }).groupW ?? 360)
              const gh = ((g.config as { groupH?: number }).groupH ?? 240)
              if (raw.x >= gn.position.x && raw.x <= gn.position.x + gw && raw.y >= gn.position.y && raw.y <= gn.position.y + gh) {
                hit = g.id
                break
              }
            }
            setDropHintId((prev) => (prev === hit ? prev : hit))
          }
        }}
        panOnDrag={!marqueeMode && !shiftDown}
        selectionKeyCode={null}
        multiSelectionKeyCode={null}
        // 关闭 React Flow 自带的「点击选中」，改由我们在 pointerdown 阶段统一处理，
        // 这样既保证快速点击不丢响应，又避免其默认行为覆盖 Ctrl 多选逻辑。
        elementsSelectable={false}
        /* ★ 2026-10-21：关闭 RF「选中抬升」——组框选中后 zIndex 不能抬到成员之上，
           否则组缩放条再次盖住成员边缘（拖动被阻挡的问题复发） */
        elevateNodesOnSelect={false}
        // 容忍 5px 内的指针抖动，减少把「快速点击」误判为「微小拖拽」的概率
        nodeDragThreshold={5}
        // 缩放下限：RF 默认 minZoom=0.5 —— 用户「缩放最小怎么是50% 最小的比例稍微小点」→ 放宽到 25%。
        // zoomOut 按钮 / 滚轮缩小 / fitView 全部自动受这个钳制，ZoomPanel 与顶栏无需各写各的。
        minZoom={0.25}
        fitView
        /* ★ 2026-10-19 用户「第一次拖入的节点不要200% 还是100%」：初始/首节点适配
         *   不许放大超过 100%（fitView 不限 maxZoom 时单节点会被放大到 200%+） */
        fitViewOptions={fitViewOptions}
        proOptions={{ hideAttribution: true }}
      >
        <Background />
        {perfHud && <PerfHud changeCount={changeCount} />}
        <ZoomPanel />
        {/* ★ 2026-10-24 用户指令「三个放缩放比例上面」：节点/已选/缩放胶囊（右下角缩放控件上方） */}
        <Panel position="bottom-right" style={{ bottom: 214, right: 12, display: 'flex', gap: 6, zIndex: 5 }}>
          <span className="cr-chip" title="画布节点总数">节点 {nodes.length}</span>
          <span className={'cr-chip' + (nodes.some((n) => n.selected) ? ' cr-chip--on' : '')} title="当前选中的节点数">已选 {nodes.filter((n) => n.selected).length}</span>
          <button className="cr-chip cr-chip--btn" title="点击恢复 100% 缩放" onClick={() => { const inst = instRef.current; if (inst) inst.setViewport({ x: 0, y: 0, zoom: 1 }) }}>
            {zoomPct}%
          </button>
        </Panel>
        <MiniMap pannable zoomable style={{ width: 200, height: 132 }} />
        <AlignmentGuides guides={guides} />
        {/* ★ 连线显示效果选择器（画布右上角内侧，避开窗口控件） */}
        <EdgeFxPicker value={edgeAnim} onChange={setEdgeAnim} />
      </ReactFlow>

      {/* ★ 2026-10-31 画布底部工作流对话框（智能助手主页式交互）：点开始/拖文件 → 展开对话讨论工作流 */}
      <CanvasChatDock />

      {/* 顶部 UI：OnlyOffice 三段式 ribbon（标题行 / 菜单标签 / 工具带） */}
      <CanvasRibbon
        ribbonSlots={ribbonSlots}
        pages={pages}
        activePageId={activePageId}
        zoomPct={zoomPct}
        stats={{ total: nodes.length, selected: nodes.filter((n) => n.selected).length }}
        marqueeMode={marqueeMode}
        trashOpen={trashOpen}
        trashCount={trashCount}
        canUndo={pastLen > 0}
        canRedo={futureLen > 0}
        needCenter={needCenter}
        nodesFolded={nodesFolded}
        roundRunState={roundRunState}
        selNoneHint={selNoneHint}
        onToggleNodesFold={() => {
          const next = !useBaseStore.getState().nodesFolded
          useBaseStore.getState().setNodesFolded(next)
        }}
        onToggleRoundRun={() => {
          /* 三态时序（用户 2026-10-06 第2条：「增加特效的结束的退出效果
             测试的时候3秒结束」）：
               off/running --播放--> running --3s 到--> exiting --0.5s--> off
               running --手动停止--> exiting（也要淡出，不能啪一下消失）
             迁移本体在 core/roundRun.nextRoundRunState（纯函数、单测钉死），
             这里只负责**排定时器**—— 定时器是唯一必须在组件侧的东西。
             ★ 每次调度前先 clearTimeout：连点两次播放会挂两个 3s 定时器，
               第二个到点时状态已在 exiting，重复推进会让退出动画卡住不消失
               （nextRoundRunState 对重复事件幂等，但定时器仍该清掉）。 */
          const st = useBaseStore.getState()
          const cur = st.roundRunState
          const ev = cur === 'running' ? 'stop' : 'play'
          const next = nextRoundRunState(cur, ev)
          st.setRoundRunState(next)
          if (roundRunTimerRef.current) {
            clearTimeout(roundRunTimerRef.current)
            roundRunTimerRef.current = null
          }
          if (next === 'running') {
            roundRunTimerRef.current = setTimeout(() => {
              roundRunTimerRef.current = null
              const s2 = useBaseStore.getState()
              const st2 = nextRoundRunState(s2.roundRunState, 'runElapsed')
              s2.setRoundRunState(st2)
              if (st2 === 'exiting') {
                // 退出淡出到点 → 回off（这一跳缺了动画永不结束，见 roundRun 注释）
                roundRunTimerRef.current = setTimeout(() => {
                  roundRunTimerRef.current = null
                  const s3 = useBaseStore.getState()
                  s3.setRoundRunState(nextRoundRunState(s3.roundRunState, 'exitElapsed'))
                }, ROUND_RUN_EXIT_MS)
              }
            }, ROUND_RUN_TEST_MS)
          } else if (next === 'exiting') {
            roundRunTimerRef.current = setTimeout(() => {
              roundRunTimerRef.current = null
              const s3 = useBaseStore.getState()
              s3.setRoundRunState(nextRoundRunState(s3.roundRunState, 'exitElapsed'))
            }, ROUND_RUN_EXIT_MS)
          }
        }}
        onUndo={doUndo}
        onRedo={doRedo}
        onGlobalPreview={handleFit}
        onSwitchPage={(id) => useBaseStore.getState().switchPage(id)}
        onNewPage={() => {
          const name = window.prompt('新画布名称', `画布${pages.length + 1}`)
          if (name !== null) useBaseStore.getState().newPage(name)
        }}
        onRenamePage={() => {
          const cur = pages.find((p) => p.id === activePageId)
          const name = window.prompt('重命名画布', cur?.name ?? '')
          if (name) useBaseStore.getState().renamePage(activePageId as string, name)
        }}
        onDeletePage={() => {
          if (activePageId && window.confirm('删除当前画布页？此操作不可撤销。')) {
            useBaseStore.getState().deletePage(activePageId)
          }
        }}
        onZoomIn={() => zoomBy('in')}
        onZoomOut={() => zoomBy('out')}
        onZoomReset={zoomReset}
        onFitView={fitAll}
        onToggleMarquee={() => setMarqueeMode((v) => !v)}
        onAutoLayout={autoLayout}
        onAutoAlign={autoAlign}
        onRunWorkflow={runWorkflow}
        onCopy={() => run(copySelection)}
        onPaste={() => run(pasteSelection)}
        onDelete={() =>
          run(() => {
            const ids = selectedIds()
            if (ids.length) onDeleteSelectedNodes(ids)
          })
        }
        onAddNode={addNodeAtCenter}
        onToggleTrash={() => setTrashOpen((v) => !v)}
      />

      {trashOpen && (() => {
        const trashed = Object.values(tables).filter((t) => t.deleted)
        return (
          <div className="trash-panel" onPointerDown={(e) => e.stopPropagation()}>
            <div className="trash-head">
              <span className="trash-title">回收站{trashed.length > 0 ? `（${trashed.length}）` : ''}</span>
              <div className="trash-head-actions">
                {(() => {
                  const last = trashStack[trashStack.length - 1]
                  const lastCount = last ? last.length : 0
                  return (
                    lastCount > 0 && (
                      <button
                        className="trash-link restore-last"
                        title="逐步恢复：恢复最近一次删除的全部内容（常用于批量删除后整批找回）"
                        onClick={() => onRestoreLastTrash()}
                      >
                        恢复上一次删除（{lastCount}）
                      </button>
                    )
                  )
                })()}
                {trashed.length > 0 && (
                  <button
                    className="trash-link danger"
                    title="永久删除回收站内全部节点"
                    onClick={() => {
                      if (window.confirm('清空回收站？所有已删除节点将被永久移除，不可恢复。')) {
                        onEmptyTrash()
                        setTrashOpen(false)
                      }
                    }}
                  >
                    清空回收站
                  </button>
                )}
                <button className="trash-link" title="收起" onClick={() => setTrashOpen(false)}>
                  收起
                </button>
              </div>
            </div>
            <div className="trash-body">
              {trashed.length === 0 ? (
                <div className="trash-empty">回收站为空</div>
              ) : (
                trashed.map((t) => (
                  <div key={t.id} className="trash-item">
                    <span className={`rf-badge ${badgeClassOf(t.kind)} trash-badge-k`}>{styleOf(t.kind).label}</span>
                    <span className="trash-name" title={t.name}>
                      {t.name}
                    </span>
                    <div className="trash-item-actions">
                      <button className="trash-restore" title="恢复到画布" onClick={() => onRestoreTable(t.id)}>
                        恢复
                      </button>
                      <button
                        className="trash-purge"
                        title="永久删除"
                        onClick={() => {
                          if (window.confirm(`永久删除「${t.name}」？此操作不可恢复。`)) onPurgeTable(t.id)
                        }}
                      >
                        彻底删除
                      </button>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        )
      })()}

      {marquee && <div className="sel-marquee" style={{ left: marquee.x, top: marquee.y, width: marquee.w, height: marquee.h }} />}

      {menu && (
        <div className="ctx-menu" style={{ left: menu.x, top: menu.y }}>
          {menu.kind === 'pane' && (
            <div className="ctx-item" onClick={() => run(() => setNodes((nds) => nds.map((n) => ({ ...n, selected: true }))))}>
              全选
            </div>
          )}
          {menu.kind === 'pane' && <div className="ctx-sep" />}
          {menu.kind === 'node' &&
            menu.nodeId &&
            EDITABLE_KINDS.includes(tables[menu.nodeId]?.kind) && (
              <>
                <div className="ctx-item" onClick={() => run(() => {
                const k = tables[menu.nodeId as string]?.kind
                if (hasPopupForm(k)) onOpenPopupNode?.(menu.nodeId as string)
                else onOpenNode?.(menu.nodeId as string)
              })}>
                  打开编辑器
                </div>
                <div className="ctx-sep" />
              </>
            )}
          {/* ★ 2026-10-22 用户指令「OO 节点的菜单 统一加入 预加载（如有）数据锁定、
              位置锁定、外部数据引用更新 等」：节点卡片按钮行的四项能力进右键菜单。
              状态从 store 现读（与节点按钮同字段同语义）；预加载需已绑定内容（「如有」），
              无内容置灰；外部数据引用更新仅 oosheet（与按钮行一致）。 */}
          {menu.kind === 'node' &&
            menu.nodeId &&
            ['ooword', 'oosheet'].includes(tables[menu.nodeId]?.kind) &&
            (() => {
              const nid = menu.nodeId as string
              const k = tables[nid]?.kind as 'ooword' | 'oosheet'
              const cfg = (tables[nid]?.config ?? {}) as {
                ooDataLocked?: boolean; ooPosLocked?: boolean; dataLocked?: boolean; posLocked?: boolean
                ooRefUpdMode?: string; ooTemplateId?: string; ooBytes64?: string; ooBytesDbId?: string
              }
              /* ★ 2026-10-26 统一锁：读写走通用 dataLocked/posLocked（遗留字段读取兼容） */
              const dataLocked = !!(cfg.dataLocked || cfg.ooDataLocked)
              const posLocked = !!(cfg.posLocked || cfg.ooPosLocked)
              const hasContent = !!cfg.ooTemplateId || !!cfg.ooBytes64 || !!cfg.ooBytesDbId
              const frozen = preloadState(nid) === 'frozen'
              const REF_LABEL: Record<string, string> = { off: '不检测', ask: '检测后提示', auto: '自动更新' }
              let updMode = cfg.ooRefUpdMode
              if (updMode !== 'off' && updMode !== 'ask' && updMode !== 'auto') {
                try {
                  const v = localStorage.getItem('ooRefUpdateMode')
                  updMode = v === 'off' || v === 'ask' || v === 'auto' ? v : 'ask'
                } catch { updMode = 'ask' }
              }
              const setCfg = (patch: Record<string, boolean | string>) =>
                useBaseStore.getState().setNodeConfig(nid, patch)
              const cycleUpd = () => {
                const next = updMode === 'off' ? 'ask' : updMode === 'ask' ? 'auto' : 'off'
                setCfg({ ooRefUpdMode: next })
                try { localStorage.setItem('ooRefUpdateMode', next) } catch { /* ignore */ }
              }
              const doPreload = async () => {
                try {
                  let ooB = cfg.ooBytes64
                  if (!ooB) {
                    const dbId = cfg.ooBytesDbId
                    if (dbId) { const r = await b64FromDbId(dbId); if (r) ooB = r }
                  }
                  await preloadOo({
                    nodeId: nid, kind: k,
                    templateId: cfg.ooTemplateId ?? undefined,
                    bytes: ooB ? b64ToU8(ooB) : null,
                  })
                  pushLog(`预加载完成：${tables[nid]?.name ?? nid}（后台冻结待命，打开 0 等待）`, 'ok')
                } catch (e) {
                  pushLog(`预加载失败（节点正常打开不受影响）: ${e instanceof Error ? e.message : String(e)}`, 'warn')
                }
              }
              /* ★ 2026-10-29 用户指令「文档右键加个转为PDF节点」→ 2026-11-08 抽成模块级
                 runNodeToPdf（对话「word转PDF」共用同一条链路），这里只负责调它。 */
              const doToPdf = () => { void runNodeToPdf(nid, { onAddTableAt, pendingPos, nodesRef }) }
              /* ★ 2026-10-29 用户指令「转PDF验证」：保真对比——保存前(ooprev-)/后(当前)各转
                 PDF，pdf-lib 拼成左右并排对比 PDF（每页左=保存前 右=保存后），建 PDF 节点逐页目检。
                 底料由 persistOoBytes 在每次覆盖 ooBytes64 前自动留取（ooprev-<id>）。 */
              const doFidelity = async () => {
                try {
                  const live = (useBaseStore.getState().tables[nid]?.config ?? {}) as { ooBytes64?: string; ooBytesDbId?: string }
                  let curU8: Uint8Array | null = null
                  if (live.ooBytes64) {
                    const b = atob(live.ooBytes64)
                    curU8 = new Uint8Array(b.length)
                    for (let i = 0; i < b.length; i++) curU8![i] = b.charCodeAt(i)
                  }
                  if ((!curU8 || !curU8.length) && live.ooBytesDbId) {
                    try { const m = await import('./core/pdfBlobStore'); curU8 = (await m.getBlob(live.ooBytesDbId)) || null } catch { /* ignore */ }
                  }
                  const mP = await import('./core/pdfBlobStore')
                  const prevU8 = await mP.getBlob('ooprev-' + nid)
                  if (!prevU8 || !curU8 || !curU8.length) { pushLog('[保真对比] 缺少对比底料（先保存一次，产生 保存前/后 两个版本）', 'warn'); return }
                  const baseN = String(tables[nid]?.name ?? '文档').replace(/\.(docx|xlsx)$/i, '')
                  pushLog(`[保真对比] 转 PDF：前 ${Math.round(prevU8.length / 1024)}KB / 后 ${Math.round(curU8.length / 1024)}KB …`)
                  const extN = k === 'oosheet' ? '.xlsx' : '.docx'
                  const pdfPrev = await convertToPdf(prevU8, baseN + extN)
                  const pdfCur = await convertToPdf(curU8, baseN + extN)
                  if (!pdfPrev || !pdfCur) { pushLog('[保真对比] 转换失败', 'err'); return }
                  const { PDFDocument } = await import('pdf-lib')
                  const outDoc = await PDFDocument.create()
                  const d1 = await PDFDocument.load(pdfPrev)
                  const d2 = await PDFDocument.load(pdfCur)
                  const W = 1190.55, H = 841.89, half = W / 2
                  const n = Math.max(d1.getPageCount(), d2.getPageCount())
                  for (let i = 0; i < n; i++) {
                    const page = outDoc.addPage([W, H])
                    const pairs: Array<[typeof d1, number, number]> = [[d1, i, 0], [d2, i, half]]
                    for (const [doc, pi, x0] of pairs) {
                      if (pi >= doc.getPageCount()) continue
                      const pg = doc.getPage(pi)
                      const sz = pg.getSize()
                      const sc = Math.min(half / sz.width, H / sz.height)
                      const emb = await outDoc.embedPage(pg)
                      page.drawPage(emb, { x: x0 + (half - sz.width * sc) / 2, y: (H - sz.height * sc) / 2, width: sz.width * sc, height: sz.height * sc })
                    }
                  }
                  const merged = await outDoc.save()
                  const stP = useBaseStore.getState()
                  stP.pushHistory()
                  const srcN = nodesRef.current.find((nn) => nn.id === nid)
                  const p2 = { x: Math.round((srcN?.position.x ?? 0) + (srcN?.measured?.width ?? 220) + 60), y: Math.round(srcN?.position.y ?? 0) }
                  const newId = onAddTableAt(p2, 'pdf')
                  pendingPos.current[newId] = p2
                  const pid = 'pd-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7)
                  await putBlobStandalone('pdfitem-' + pid, merged)
                  stP.setNodeConfig(newId, { pdfItems: [{ id: pid, name: baseN + '-保真对比.pdf', inDb: true }], pdfFileOrder: [pid] })
                  stP.renameTable(newId, baseN + '-保真对比')
                  pushLog(`[保真对比] 已生成对比 PDF 节点（${n} 页，左=保存前 右=保存后）：${Math.round(merged.length / 1024)}KB`, 'ok')
                } catch (eF) {
                  pushLog(`[保真对比] 失败：${String((eF as Error)?.message ?? eF).slice(0, 120)}`, 'err')
                }
              }
              return (
                <>
                  <div
                    className={'ctx-item' + (hasContent ? '' : ' ctx-item-dis')}
                    title={hasContent
                      ? (frozen ? '已预加载（冻结待命）——再次点击无副作用，打开节点 0 等待'
                              : '预加载：后台打开文档并冻结，打开节点 0 等待')
                      : '预加载（如有）：节点尚未绑定内容——先双击进入编辑导入/创建文档'}
                    onClick={() => { if (hasContent) run(() => { void doPreload() }) }}
                  >
                    预加载（后台打开+冻结）{frozen ? ' ✓' : ''}
                  </div>
                  <div
                    className="ctx-item"
                    title="数据锁定：文档以只读模式打开（防误编辑；重开保持）"
                    onClick={() => run(() => setCfg({ dataLocked: !dataLocked }))}
                  >
                    数据锁定{dataLocked ? ' ✓' : ''}（文档只读）
                  </div>
                  <div
                    className="ctx-item"
                    title="位置锁定：节点在画布上不可拖动（防误拖；重开保持）"
                    onClick={() => run(() => setCfg({ posLocked: !posLocked }))}
                  >
                    位置锁定{posLocked ? ' ✓' : ''}（不可拖动）
                  </div>
                  {k === 'oosheet' && (
                    <div
                      className="ctx-item"
                      title={'外部数据引用更新：' + (REF_LABEL[updMode] || updMode)
                        + '（随节点保存）—— 点击切换：不检测 → 检测后提示 → 自动更新'}
                      onClick={() => run(() => cycleUpd())}
                    >
                      外部数据引用更新：{REF_LABEL[updMode] || updMode}
                    </div>
                  )}
                  <div
                    className="ctx-item"
                    title="导入文档：从本地选择 docx / xlsx 作为本节点文档（空白节点直接导入，已有内容则替换）"
                    onClick={() => { ooImportTargetRef.current = nid; setMenu(null); ooImportRef.current?.click() }}
                  >
                    导入文档…
                  </div>
                  <div
                    className={'ctx-item' + (hasContent ? '' : ' ctx-item-dis')}
                    title={hasContent
                      ? '转为 PDF 节点：当前文档经 x2t 转 PDF，在节点右侧新建一个 PDF 节点装入并自动连线'
                      : '转为 PDF 节点：节点尚未绑定内容——先导入/保存文档'}
                    onClick={() => { if (hasContent) run(() => { void doToPdf() }) }}
                  >
                    转为 PDF 节点
                  </div>
                  <div
                    className={'ctx-item' + (hasContent ? '' : ' ctx-item-dis')}
                    title={hasContent
                      ? '保真对比：把「保存前(上一版)/保存后(当前)」各转 PDF 左右并排生成对比节点，逐页目检版式差异'
                      : '保真对比：先保存一次产生前后版本'}
                    onClick={() => { if (hasContent) run(() => { void doFidelity() }) }}
                  >
                    保真对比（PDF 节点）
                  </div>
                  <div className="ctx-sep" />
                </>
              )
            })()}
          {/* ★ 统一节点锁（2026-10-26 全节点推广）：非 OO 节点的右键锁入口——
              OO 走上方专属块；其余全 kind 通用字段 dataLocked/posLocked
              （text/表格/圆形等无按钮行节点由此获得锁入口） */}
          {menu.kind === 'node' &&
            menu.nodeId &&
            !['ooword', 'oosheet'].includes(tables[menu.nodeId]?.kind ?? '') &&
            (() => {
              const nid = menu.nodeId as string
              const lcfg = (tables[nid]?.config ?? {}) as { dataLocked?: boolean; posLocked?: boolean; ooDataLocked?: boolean; ooPosLocked?: boolean }
              const dataLocked = !!(lcfg.dataLocked || lcfg.ooDataLocked)
              const posLocked = !!(lcfg.posLocked || lcfg.ooPosLocked)
              return (
                <>
                  <div
                    className="ctx-item"
                    title="数据锁定：节点内容防变动（导入/删除/排序/执行停用；重开保持）"
                    onClick={() => run(() => useBaseStore.getState().setNodeConfig(nid, { dataLocked: !dataLocked }))}
                  >
                    数据锁定{dataLocked ? ' ✓' : ''}（防内容变动）
                  </div>
                  <div
                    className="ctx-item"
                    title="位置锁定：节点在画布上不可拖动（防误拖；重开保持）"
                    onClick={() => run(() => useBaseStore.getState().setNodeConfig(nid, { posLocked: !posLocked }))}
                  >
                    位置锁定{posLocked ? ' ✓' : ''}（不可拖动）
                  </div>
                  <div className="ctx-sep" />
                </>
              )
            })()}
          <div
            className="ctx-item"
            onClick={() =>
              run(() => {
                if (menu.kind === 'node' && menu.nodeId) onDeleteNode(menu.nodeId)
                else onDeleteSelectedNodes(selectedIds())
              })
            }
          >
            删除
          </div>
          <div
            className="ctx-item"
            title="存入剪贴板：右键空白处「粘贴」或 Ctrl+V 贴到光标位置"
            onClick={() =>
              run(() => {
                /* ★ 2026-10-22 用户指令「点击节点有复制，点空白右键粘贴到右键位置」：
                   复制=存剪贴板（不再立即偏移粘贴） */
                if (menu.kind === 'node' && menu.nodeId && !tables[menu.nodeId]?.deleted) {
                  /* ★ 2026-10-22 右键「复制」也走公共装剪贴板逻辑——
                     组节点会连带成员一起装（与 copySelection 同一口径）。
                     原先这里只装组框 1 项，粘贴时 items/ids 配对越界，
                     成员落点全错、且新组框内没有成员（用户截图实锤）。 */
                  copyIdsToClipboard([menu.nodeId])
                } else {
                  copySelection()
                }
              })
            }
          >
            复制
          </div>
          {menu.kind === 'pane' && clipboardRef.current.length > 0 && (() => {
            /* ★ 剪贴板含组节点 → 三种粘贴模式（2026-10-22 用户指令）；
               普通节点 → 直接粘贴。 */
            const hasGroup = clipboardRef.current.some((it) => it.table?.kind === 'group')
            const at = menu.flowX != null && menu.flowY != null ? { x: menu.flowX, y: menu.flowY } : undefined
            if (!hasGroup) {
              return (
                <div
                  className="ctx-item"
                  title="粘贴到此处（整组保持相对布局）"
                  onClick={() => run(() => pasteSelection(at))}
                >
                  粘贴
                </div>
              )
            }
            return (
              <>
                <div className="ctx-item" onClick={() => run(() => pasteSelection(at, 'group'))}>
                  粘贴 · 组+成员（含数据）
                </div>
                <div className="ctx-item" title="复制组结构但成员数据清空" onClick={() => run(() => pasteSelection(at, 'group-empty'))}>
                  粘贴 · 组+成员（清空数据）
                </div>
                <div className="ctx-item" title="不建组，成员节点直接铺在画布上" onClick={() => run(() => pasteSelection(at, 'flat'))}>
                  粘贴 · 仅成员节点
                </div>
              </>
            )
          })()}
          <div className="ctx-sep" />
          <div
            className="ctx-item"
            onClick={() =>
              run(() => {
                if (menu.kind === 'node' && menu.nodeId) bringToFront(menu.nodeId)
                else selectedIds().forEach((id) => bringToFront(id))
              })
            }
          >
            置于顶层
          </div>
          {/* ★ 2026-10-23 用户指令「别的 pdf 套件都秒开」：双击=iframe 秒开预览；
             编辑（OO 转换 10~30s）走右键知情选择 */}
          {menu.kind === 'node' && menu.nodeId && tables[menu.nodeId]?.kind === 'pdf' && (
            <div
              className="ctx-item"
              title="OnlyOffice 打开并编辑（需转换，大文件 10~30 秒）"
              onClick={() => {
                const id = menu.nodeId as string
                setMenu(null)
                onOpenNode?.(id)
              }}
            >
              用 OO 编辑此 PDF
            </div>
          )}
          {menu.kind === 'node' && menu.nodeId && tables[menu.nodeId]?.kind === 'pdfmerge' && (
            <div
              className="ctx-item"
              title="OnlyOffice 打开并编辑（需转换，大文件 10~30 秒）"
              onClick={() => {
                const id = menu.nodeId as string
                setMenu(null)
                onOpenNode?.(id)
              }}
            >
              用 OO 编辑此 PDF
            </div>
          )}
          {/* ★ 2026-10-24 用户指令「再配置一个 github 最优套件 右键菜单打开」：
             pdf.js 官方查看器（Firefox 同款渲染管线），事件桥触发 PdfNodeCard 预览 */}
          {menu.kind === 'node' && menu.nodeId && (tables[menu.nodeId]?.kind === 'pdf' || tables[menu.nodeId]?.kind === 'pdfmerge') && (
            <div
              className="ctx-item"
              title="pdf.js 官方查看器（Firefox 同款渲染管线，整页适应+连续滚动）"
              onClick={() => {
                const id = menu.nodeId as string
                setMenu(null)
                window.dispatchEvent(new CustomEvent('open-pdf-pjsviewer', { detail: { nodeId: id } }))
              }}
            >
              用 PDF.js Viewer 打开
            </div>
          )}
          {/* ★ 2026-10-24 用户指令：右键「用默认软件打开」——IPC 落临时文件 → shell.openPath */}
          {menu.kind === 'node' && menu.nodeId && (tables[menu.nodeId]?.kind === 'pdf' || tables[menu.nodeId]?.kind === 'pdfmerge') && (
            <div
              className="ctx-item"
              title="用系统默认的 PDF 程序打开（WPS/Adobe 等）"
              onClick={async () => {
                const id = menu.nodeId as string
                setMenu(null)
                try {
                  /* ★ 2026-10-26 重构：取字节走 pdfOpenTargetOf 共用函数（浏览器打开项同源） */
                  const hit = await pdfOpenTargetOf(tables, id)
                  if (!hit) return
                  const api = (window as unknown as { api?: { openExternalFile?: (n: string, d: Uint8Array) => Promise<string> } }).api
                  if (!api?.openExternalFile) { console.error('[pdf] 桌面桥不可用'); return }
                  await api.openExternalFile(hit.name, hit.bytes)
                } catch (e) { console.error('[pdf] 默认软件打开失败', e) }
              }}
            >
              用默认软件打开
            </div>
          )}
          {/* ★ 2026-10-26 用户指令「所有PDF增加个 用本地浏览器打开」：Edge/Chrome 自带
              PDF 查看器渲染（主进程找本机浏览器，找不到退系统关联） */}
          {menu.kind === 'node' && menu.nodeId && (tables[menu.nodeId]?.kind === 'pdf' || tables[menu.nodeId]?.kind === 'pdfmerge') && (
            <div
              className="ctx-item"
              title="用本机浏览器（Edge/Chrome）打开——自带 PDF 查看器"
              onClick={async () => {
                const id = menu.nodeId as string
                setMenu(null)
                try {
                  const hit = await pdfOpenTargetOf(tables, id)
                  if (!hit) return
                  const api = (window as unknown as { api?: { openInBrowser?: (n: string, d: Uint8Array) => Promise<string> } }).api
                  if (!api?.openInBrowser) { console.error('[pdf] 桌面桥不可用'); return }
                  await api.openInBrowser(hit.name, hit.bytes)
                } catch (e) { console.error('[pdf] 浏览器打开失败', e) }
              }}
            >
              用本地浏览器打开
            </div>
          )}
          {/* ★ 2026-10-24 用户指令「专业渲染引擎第四入口」：MuPDF（Foxit 旗下引擎，
             渲染质量业界第一梯队）canvas 直绘 2x 超采样 */}
          {menu.kind === 'node' && menu.nodeId && tables[menu.nodeId]?.kind === 'pdf' && (
            <div
              className="ctx-item"
              title="MuPDF 专业引擎（canvas 直绘，渲染质量顶级）"
              onClick={() => {
                const id = menu.nodeId as string
                setMenu(null)
                window.dispatchEvent(new CustomEvent('open-pdf-mupdf', { detail: { nodeId: id } }))
              }}
            >
              用 MuPDF 打开
            </div>
          )}
                    {/* ★ 2026-10-23 用户指令「右键节点 以此为基准 重新排列：左按连接顺序、右侧也是」 */}
          {menu.kind === 'node' && menu.nodeId && (
            <div
              className="ctx-item"
              title="以该节点为锚：上游往左、下游往右按拓扑层级重排（可撤销）"
              onClick={() => {
                const id = menu.nodeId as string
                setMenu(null)
                rerangeFromNode(id)
              }}
            >
              以此为基准 重新排列
            </div>
          )}
          {menu.kind === 'node' && menu.nodeId && (
            <div className="ctx-item" title="连线流动动画：拖动时暂停（默认）/ 始终播放 / 关闭" onClick={() => { setMenu(null); cycleEdgeAnim() }}>
              连线动画：{edgeAnim === 'pause' ? '拖动时暂停' : edgeAnim === 'always' ? '始终播放' : '关闭'} ↻
            </div>
          )}
          <div
            className="ctx-item"
            onClick={() =>
              run(() => {
                if (menu.kind === 'node' && menu.nodeId) sendToBack(menu.nodeId)
                else selectedIds().forEach((id) => sendToBack(id))
              })
            }
          >
            置于底层
          </div>
          {detachItems.length > 0 && <div className="ctx-sep" />}
          {detachItems.map((it) => (
            <div key={it.key} className="ctx-item ctx-detach-item" title={it.label} onClick={() => run(it.run)}>
              {it.label}
            </div>
          ))}
        </div>
      )}
    </div>
  )
}
