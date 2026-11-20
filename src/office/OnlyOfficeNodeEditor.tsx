import { forwardRef, useCallback, useEffect, useImperativeHandle, useRef, useState } from 'react'
import { getTemplateStore } from '../docx/templateStore'
import { hasFrozen, takeFrozen, waitFrozen, preloadState, refreeze } from './ooPreload'
import { getRefUi, setRefUi, subscribeRefUi, isApplyingFromPoc } from '../xlsx/refPanelUi'
import { ensureCaps } from '../core/capabilities'
/* ★ 2026-10-17 用户指令：编辑器内的点击/键盘/反馈也进操作日志。
 *   /poc 与 frameEditor 两层的采集结果经 postMessage(__ooAudit) 汇总到这里，
 *   由 auditEditor() 统一落到 logs/audit-YYYY-MM-DD.log（与宿主操作同一条时间线）。 */
import { audit, auditFeedback } from '../core/audit'
import { stripFullCalcFlag } from '../core/ooDomains'
import { useBaseStore } from '../core/store'
import { pushLog } from '../ui/logBus'

/**
 * OnlyOffice 节点编辑器（ooword / oosheet）—— 纯 iframe 版（v2）
 *
 * 架构：所有复杂度（DocsAPI、桥、轮询、推 bin、x2t）封在 POC 页面
 * （5199/poc）内部自成体系——React 零介入，节点里只嵌一个 iframe。
 *
 * 数据流：节点字节 → POST 5199/office-offline/open 写工作区
 *        → iframe src=5199/poc?src=<workName> → 编辑器全链路（iframe 内自成体系）
 *
 * 标题栏约定（与 Excel 节点 ne-head 一致）：
 *   📂 打开 / 💾 保存回写 / ⤓ 导出 三个动作由 NodeEditor 的 ne-fileops 触发，
 *   经 ref 句柄调进来 —— 本组件不再渲染自己的工具栏行。
 *
 * 运行前提：`node scripts/onlyoffice-poc-server.mjs`（5199）保持运行。
 */

export type OOKind = 'ooword' | 'oosheet' | 'pdf'

const KIND_CFG: Record<OOKind, { ext: string; title: string }> = {
  ooword: { ext: 'docx', title: 'OO 文档' },
  oosheet: { ext: 'xlsx', title: 'OO 工作簿' },
  /* ★ 2026-10-23 用户指令「换能编辑 PDF 的套件」：OO 9.4 原生 PDF 编辑（≥7.5）——
     pdf 节点走 OO 链（x2t pdf→内部格式，可编辑保存回 pdf） */
  pdf: { ext: 'pdf', title: 'PDF 文档' },
}

/** 服务端 __ooDigest 回传的内容摘要（定义收口到 saveChannel，避免两处漂移） */
import { digestSame, type OoDigest } from './saveChannel'
export type { OoDigest }

export interface OoEditorHandle {
  /** 标题栏 📂：读取本地文件 → 上传工作区 → 编辑器打开 */
  openFile: (f: File, extOverride?: string) => Promise<void>
  /** 关闭前询问：文档是否有未保存修改（走 poc 的 isDocumentModified 桥） */
  /** ★ 保存完成回执：清 /poc 的脏标志（用户保存过就不该再提示） */
  ackSaved: () => void
  isModified: () => Promise<boolean>

  /** ★ 2026-10-11 保存链重构：请求 /poc 引擎直读值数组（新保存链入口，触发 __ooSaveValues 回传） */
  /** ★ 2026-10-17：sid 由宿主在点 💾 时生成，随请求下发，用于把整条链的日志串起来 */
  requestSave: (mode?: 'node' | 'export', sid?: string) => void
  /** 旧保存链（DS forcesave——已证协作通道空转，保留仅作回退）
   *  ★ 2026-10-17：不再直接返回字节，改为带回「产物内容摘要 + 是否等于打开时的原文件」，
   *    让调用方能判定这次主通道到底有没有真的存到东西（此前只要有字节就当成功，
   *    空表 6258B 把提取通道的正确内容覆盖掉了）。 */
  saveBack: () => Promise<{ bytes: Uint8Array; digest: OoDigest | null; sameAsSource: boolean } | null>
  /** ★ 2026-10-29 样式化提取：callCommand 直读引擎文档模型（段落+run 格式 JSON，表格拆行）——
   *  ooword 保存链 forcesave 400 / Editor.bin 静态（打开时从源文件转出，永不更新），
   *  这是唯一能拿到「编辑后内容 + 字体格式」的通道。paras=null 时 err 带失败原因（诊断排障用）。 */
  styledExtract: () => Promise<{ paras: unknown[] | null; err?: string; meta?: string } | null>
  /** 当前工作区字节（⤓ 导出用） */
  currentBytes: () => Promise<Uint8Array | null>
  /** 插入引用公式文本（批次 12·③）：往光标处写 `=节点名.字段名`，参与公式计算 */
  insertRef: (ref: {
    nodeId: string
    nodeName?: string
    fieldId?: string
    fieldName?: string
    value?: unknown
    mode?: string
  }) => void
  /** 宿主操作写诊断抽屉（2026-10-10 用户指令「日志里增加我点击的操作」） */
  opLog: (msg: string) => void
}

export const OnlyOfficeNodeEditor = forwardRef<OoEditorHandle, OnlyOfficeNodeEditorProps>(function OnlyOfficeNodeEditor(
  { anyDragging, kind, nodeId, dataLocked, bytes, name, ooTemplateId, ooBytes64, ooBytesDbId, injectRefs, diagOpen, theme, maximized, refOpen, ooRefItems, onToggleRef, onSaveValues, saving, onPickRef, onOpenedName, onStatus, onRefreshRef, refreshBusy, rowColHi, refMarks, ooRefUpdMode, ooRefAsk, ooRefInfo, onSetRefUpdMode, onApplyRefAsk, onSaved, onDirtyChange, onClose },
  ref,
) {
  const cfg = KIND_CFG[kind]
  /* ★ 保存完成后显示「已保存」2.5 秒再消失（用户指令：不要消失太快） */
  useEffect(() => {
    if (prevSavingRef.current && !saving) {
      setSavedFlash(true)
      const t = setTimeout(() => setSavedFlash(false), 2500)
      prevSavingRef.current = !!saving
      return () => clearTimeout(t)
    }
    prevSavingRef.current = !!saving
  }, [saving])
  const [src, setSrc] = useState('')
  /* ★ 2026-10-18 用户实测「点着色开关后颜色不变，重新打开节点才变」：
   *   字节层已改对（toast 报「去色 8 格」），但桌面 Electron 下只改 iframe 的 src 属性
   *   并不会让新文档真正加载 —— 复用同一个 iframe 时导航被吞，画面停在旧文档。
   *   这里加一个递增序号当 React key：每次上传新工作区都把 iframe 元素整体重建，
   *   引擎一定重新加载（等于自动帮用户做一次「关掉重开」）。 */
  const [srcSeq, setSrcSeq] = useState(0)
  const [blankHint, setBlankHint] = useState(false)
  /* ★ 2026-10-17 网页端降级：服务端 /healthz 报 oo:false（没有 OnlyOffice 文档服务）
   *   → 不加载编辑器，直接给只读说明。避免网页版用户点开一个永远转圈的黑框。 */
  const [ooOff, setOoOff] = useState(false)
  const [savedFlash, setSavedFlash] = useState(false)
  const [recentFiles, setRecentFiles] = useState<{ name: string; mtime: number }[] | null>(null)
  const prevSavingRef = useRef(false)
  const [status, setStatus] = useState('准备…')
  const [diag, setDiag] = useState('')
  const [diagCollapsed, setDiagCollapsed] = useState(false) // ★ 诊断内容折叠（标题行保留）
  // ★ 预加载冻结实例（2026-10-07）：命中时 JSX 不渲染新 iframe，把冻结 iframe 挂回容器原地恢复
  const [attachedFrozen, setAttachedFrozen] = useState<HTMLIFrameElement | null>(null)
  const blankFileRef = useRef<HTMLInputElement | null>(null)
  const containerRef = useRef<HTMLDivElement | null>(null)
  const /** 当前工作区文件名（保存/导出时从 5199 拉字节用） */
    workRef = useRef('')
  /** ★ 2026-10-17：打开时上传的原始字节指纹 + 服务端摘要。
   *  保存时若主通道产物指纹与之相同 ⇒ 产物就是原文件本身（Editor.bin 不含编辑内容），
   *  据此把「假成功」拦下来，改走提取通道。 */
  const srcSigRef = useRef('')
  const srcDigestRef = useRef<OoDigest | null>(null)
  /** ★ 2026-10-17 保存会话 ID（宿主点 💾 时生成下发），宿主侧日志统一带它 */
  const saveSidRef = useRef('')
  const iframeRef = useRef<HTMLIFrameElement | null>(null)
  // 主题只在首次挂载时进 src；后续切换走 setTheme 消息（不重挂 iframe、不重载文件）
  const initialThemeRef = useRef<'dark' | 'light'>(theme === 'dark' ? 'dark' : 'light')
  const ooRefItemsRef = useRef<{ id: string; name: string; fields: { id: string; name: string; value: string }[] }[]>([])
  const modAskRef = useRef<((v: boolean) => void) | null>(null)
  /* ★ 2026-10-17 未保存修改状态（用户指令「做是否修改了的判断 然后提示保存」）：
   *   来源 = /poc 主动推的 __ooDirty（onDocumentStateChange / 键盘输入 / 保存回执）。
   *   ref 供 isModified() 同步读，state 供 UI/宿主渲染。 */
  const dirtyRef = useRef(false)
  const applyDirtyRef = useRef<(v: boolean, why: string) => void>(() => { /* 挂载后替换 */ })
  // ★ 冻结循环（2026-10-07 用户：「打开 关闭再打开……不要重复打开」）：
  //   预加载来源的实例，节点编辑器**关闭**时重新冻结（移回隐藏 host + 放回 frozenMap），
  //   再打开时原地恢复——文档实例跨开关循环存活，修改保留在冻结现场。
  const adoptedRef = useRef<{ fileName: string } | null>(null)

  /* ★ 章节目录侧栏（2026-10-09）：大纲/跳转/拖拽重排（仅 ooword + 最大化显示）。
     链路：面板 → iframe(/poc) → cbref 插件 callCommand → 结果 __chaptersData/__chaptersMoved 回传。
     拖拽语义：拖标题 = 移动「该标题段 + 其全部子内容」（到下一个标题前的区间），
     与 Word 大纲视图的语义一致。编号用 OO 原生多级列表（重排自动刷新），本侧不碰文本。 */
  const [chapters, setChapters] = useState<{ i: number; level: number; text: string }[]>([])
  const [chTotal, setChTotal] = useState(0)
  const [chapOpen, setChapOpen] = useState(false)
  const [chErr, setChErr] = useState('')
  const chapDragFrom = useRef<number | null>(null)
  const [chapDragOver, setChapDragOver] = useState<number | null>(null)
  const cursorSeenRef = useRef<string>('')
  /** 审计日志用的状态消息去重表：key=消息前 40 字符 → 最近一次落盘时间 + 累计条数 */
  const statusSeenRef = useRef(new Map<string, { at: number; n: number }>())
  /* ★ 2026-10-22 最后选中格持久化：光标防抖写定时器 / 恢复只执行一次标记 */
  const selSaveTimer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const selRestoredRef = useRef(false)
  /* 卸载兜底：防抖没跑就关闭（未到 1s）也把最后光标写进 config */
  useEffect(() => () => {
    if (selSaveTimer.current) clearTimeout(selSaveTimer.current)
    const last = cursorSeenRef.current
    if (last && /^[A-Z]+[0-9]+$/.test(last)) {
      try { useBaseStore.getState().setNodeConfig(nodeId!, { ooLastSel: last }) } catch { /* ignore */ }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // ★ 冻结恢复模式下 iframeRef.current 为 null（冻结 iframe 命令式挂回，不经 JSX ref）。
  //   所有 React→编辑器 postMessage 一律经此取窗口：ref 优先、容器实查兜底。
  //   （2026-10-09 章节命令、引用插入两次踩同一坑，2026-10-10 统一收口）
  const getEditorWindow = (): Window | null =>
    iframeRef.current?.contentWindow
    || containerRef.current?.querySelector('iframe')?.contentWindow
    || null

  const sendChapterCmd = (op: string, extra?: Record<string, unknown>) => {
    const win = getEditorWindow()
    if (!win) { setChErr('编辑器 iframe 未就绪'); return }
    setChErr('取大纲中…')
    try { win.postMessage({ __chapterCmd: { op, ...(extra || {}) } }, '*') } catch { /* ignore */ }
    // 2.5s 无回传 → 明确报「插件未响应」，不再让用户面对无声的空列表
    window.setTimeout(() => {
      setChErr((prev) => (prev === '取大纲中…' ? '插件未响应（cbref autostart 未生效？）' : prev))
    }, 2500)
  }
  const refreshChapters = () => { if (kind === 'ooword') sendChapterCmd('getChapters') }

  /* ★ 2026-10-17 用户指令「每个步骤都监听、写在日志里，别再让我贴日志」：
   *   addDiag 原先只进抽屉（组件内 state），进程一关就没了，我排查时什么都看不到。
   *   现在每条诊断同时落盘到 oo-events.log（/office-log），并做两级降噪：
   *   ① [光标] 类高频消息不落盘（一次点选就是一条，会淹没真正的步骤）；
   *   ② 同一条消息 800ms 内只落一次（重试/轮询会重复打同一句）。
   *   关键步骤另有固定编号（S1/S10/S11…），读取脚本按编号判定缺环。 */
  const diagLogRef = useRef<{ msg: string; at: number }>({ msg: '', at: 0 })
  const addDiag = (msg: string) => {
    const ts = new Date().toLocaleTimeString('zh-CN', { hour12: false })
    setDiag((prev) => `[${ts}] ${msg}\n` + prev)
    /* ★ 2026-10-22 用户指令「右上角增加个和oo节点内一致的显示诊断信息」：
       诊断同步汇入全局日志总线 → 画布底部 LogDock 面板可见（不用开节点内抽屉） */
    try { pushLog(msg) } catch { /* 日志总线异常不影响诊断主流程 */ }
    try {
      if (/^\[光标\]/.test(msg)) return
      const now = Date.now()
      if (diagLogRef.current.msg === msg && now - diagLogRef.current.at < 800) return
      diagLogRef.current = { msg, at: now }
      void fetch('/office-log', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type: 'DIAG', detail: msg.slice(0, 700), sid: saveSidRef.current }),
      }).catch(() => { /* 日志失败不影响主流程 */ })
    } catch { /* ignore */ }
  }
  /* 关键步骤专用落盘点：type 直接用步骤编号（S1/S10/S11…），读取脚本按编号查缺环 */
  const logStep = (step: string, detail: string) => {
    try {
      void fetch('/office-log', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type: step, detail: String(detail).slice(0, 700), sid: saveSidRef.current }),
      }).catch(() => { /* ignore */ })
    } catch { /* ignore */ }
    /* ★ 2026-10-17：同一条「软件反馈」也进操作审计日志 ——
     *   这样「我点保存之后发生了什么」可以在一个文件里按时间顺序看完，
     *   不用在 oo-events.log 和 audit-*.log 之间来回对时。 */
    auditFeedback(step, String(detail).slice(0, 500))
  }
  const say = (s: string) => {
    setStatus(s)
    onStatus?.(s)
  }

  /* ★ 2026-10-17 未保存修改的唯一写入口（/poc 推来 / 本地保存完成时也走它）：
   *   状态翻转才动 UI，避免每次按键都重渲染宿主。 */
  applyDirtyRef.current = (v: boolean, why: string) => {
    if (dirtyRef.current === v) return
    dirtyRef.current = v
    try { onDirtyChange?.(v) } catch { /* ignore */ }
    addDiag(v ? '● 有未保存修改（' + why + '）' : '○ 未保存修改已清（' + why + '）')
    logStep('DIRTY', (v ? '有未保存修改' : '已无未保存修改') + ' 来源=' + why)
  }

  /* ★ 2026-10-17 宿主轮询兜底（用户反馈「保存后再改，不提示未保存」）：
   *   保存后自动重开文档会把 keydown 监听甩在旧文档上、引擎状态事件也可能送不到 ——
   *   推送链一旦断链，宿主就永远不知道文档已脏。每 2s 主动问一次（ask 通道，
   *   /poc 应答取「本地标志 OR 引擎直查」），发现真脏但宿主没亮 → 强制点亮（自愈）。 */
  useEffect(() => {
    const askTick = async () => {
      if (dirtyRef.current) return
      const w = getEditorWindow()
      if (!w) return
      const m = await new Promise<boolean>((resolve) => {
        modAskRef.current = resolve
        try { w.postMessage({ __ooIsModifiedAsk: 1 }, '*') } catch { resolve(false) }
        setTimeout(() => { if (modAskRef.current) { modAskRef.current = null; resolve(false) } }, 600)
      })
      if (m && !dirtyRef.current) applyDirtyRef.current(true, '宿主轮询兜底')
    }
    const t = window.setInterval(() => { void askTick() }, 2000)
    return () => window.clearInterval(t)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // ★ /poc 页就绪/异常状态经 postMessage(__ooStatus) 上报 → 记入诊断日志（ⓘ 抽屉查看）
  useEffect(() => {
    const onMsg = (e: MessageEvent) => {
      const d = e.data as { __ooStatus?: string; __ooToggleRef?: number; __ooRefDataRequest?: number; __ooIsModified?: boolean; __ooDirty?: boolean; __ooDirtyWhy?: string; __ooRefPick?: { nodeId: string; fieldId: string; fieldName: string; value?: string; mode?: string }; __ooCursor?: string; __ooAudit?: { kind?: string; action?: string; detail?: string } } | null
      if (d && typeof d === 'object') {
        /* ★ 2026-10-17 操作审计：编辑器内部（/poc + frameEditor）采集的点击/键盘/输入/滚轮，
         *   在这里落地。⚠ 必须放在所有 if 之前：审计不该因为某个分支提前 return 而丢。 */
        const oa = d.__ooAudit
        if (oa && typeof oa === 'object') {
          // kind 由采集侧给（点击/键盘/输入/滚轮/焦点/系统），不要一律写成「编辑器」
          audit(String(oa.kind || '编辑器'), String(oa.action || '?'), String(oa.detail || ''))
        }
        if (typeof d.__ooStatus === 'string') {
          addDiag('[poc] ' + d.__ooStatus)
          /* ★ 2026-10-22 用户指令「ooexcel 中也保存最后选中单元格的位置等信息」：
           *   光标跟踪注册成功 = 引擎/插件已就绪 → 回放 config.ooLastSel
           *   （restoreSel 命令：选中该格并滚动到可见）。只回放一次。 */
          if (d.__ooStatus.indexOf('光标事件注册成功') >= 0 && !selRestoredRef.current) {
            selRestoredRef.current = true
            try {
              const cfg0 = useBaseStore.getState().tables[nodeId!]?.config as { ooLastSel?: string } | undefined
              const last = cfg0?.ooLastSel
              if (last && /^[A-Z]+[0-9]+$/.test(last)) {
                setTimeout(() => {
                  const win = getEditorWindow()
                  if (!win) return
                  win.postMessage({ __cbRefCmd: 'restoreSel', cell: last }, '*')
                  addDiag('[恢复] 选中格 → ' + last)
                }, 600)
              }
            } catch { /* 恢复失败不影响打开 */ }
          }
          /* ★ 编辑器反馈（loading/就绪/报错）也在审计日志里留痕，但要降噪：
           *   cellDiag 一类**周期性诊断**每秒都会重发（内容几乎不变），
           *   全量写进去一分钟就几十行，把真正的操作淹没掉。
           *   ⇒ 同一条消息（取前 40 字符做键）30 秒内只落一次，并标注折叠了多少条。 */
          const key = d.__ooStatus.slice(0, 40)
          const now = Date.now()
          const rec = statusSeenRef.current.get(key)
          if (!rec) {
            statusSeenRef.current.set(key, { at: now, n: 1 })
            audit('编辑器', '状态', d.__ooStatus)
          } else {
            rec.n++
            if (now - rec.at > 30000) {
              audit('编辑器', '状态', d.__ooStatus + '（同类消息 30s 内 ' + rec.n + ' 条，已折叠）')
              statusSeenRef.current.set(key, { at: now, n: 1 })
            }
          }
        }
        // ★ 用户操作可见化（2026-10-10 用户指令「日志里增加我点击的操作」）：
        //   点选单元格 → 光标回执（去重，只在换格时记一条）
        const cur = d.__ooCursor
        if (typeof cur === 'string' && cur && cur !== cursorSeenRef.current) {
          cursorSeenRef.current = cur
          addDiag('[光标] 选中 ' + cur)
          audit('编辑器', '选中单元格', cur)
        }
        /* ★ 2026-10-22 用户指令「ooexcel 中也保存最后选中单元格的位置等信息」：
         *   光标变化 1s 防抖写 config.ooLastSel（A1 坐标）——打开节点时恢复选中
         *   并滚动可见（插件 restoreSel 命令）。setNodeConfig 不压撤销栈（光标移动
         *   不该进撤销历史）。卸载时兜底再写一次当前值。 */
        if (typeof cur === 'string' && /^[A-Z]+[0-9]+$/.test(cur)) {
          if (selSaveTimer.current) clearTimeout(selSaveTimer.current)
          const nid = nodeId!
          selSaveTimer.current = setTimeout(() => {
            try { useBaseStore.getState().setNodeConfig(nid, { ooLastSel: cur }) } catch { /* ignore */ }
          }, 1000)
        }
        if (d.__ooToggleRef) onToggleRef?.()
        /* ★ Ctrl+S（用户指令 2026-10-11）：/poc 转发的保存请求 → 宿主保存回写节点 */
        /* ★ 2026-10-11 保存链重构：/poc 引擎直读值数组回传，宿主按值构建 xlsx 入节点配置 */
        const sv = (d as { __ooSaveValues?: { name: string; rows: unknown[][]; mode?: string } }).__ooSaveValues
        /* ★ 2026-10-17 用户指令「每步都落盘」：S9 是宿主收到提取值的唯一入口。
         *   空 rows 原先被静默丢弃（点保存毫无反应），现在明确记一条 S9-E 说明为何没存。 */
        if (sv) {
          const n = sv.rows ? sv.rows.length : 0
          if (n > 0) {
            logStep('S9', '宿主收到提取值 name=' + sv.name + ' rows=' + n + ' mode=' + (sv.mode || 'node'))
            onSaveValues?.({ name: sv.name, rows: sv.rows, mode: sv.mode })
          } else {
            logStep('S9-E', '提取到 0 行，跳过保存（避免把空表覆盖到已有内容） name=' + sv.name)
            addDiag('⚠ 提取到 0 行内容，本次未保存（防止覆盖已有内容）。请确认单元格已按 Enter 提交后再点 💾')
          }
        }
        setBlankHint(false)
        // ★「刷新域值」按钮：/poc 只发请求，写回由宿主执行（它才有模板/补丁/存储能力）
        if ((d as { __ooRefreshRef?: number }).__ooRefreshRef) {
          addDiag('[poc] 收到刷新域值请求')
          onRefreshRef?.()
        }
        /* ★ 2026-10-18 外部引用更新策略：/poc 设置卡改策略 / 弹窗点「立即更新」 */
        const rum = (d as { __ooRefUpdModeSet?: string }).__ooRefUpdModeSet
        if (typeof rum === 'string') {
          addDiag('[poc] 外部数据更新策略 → ' + rum)
          onSetRefUpdModeRef.current?.(rum)
        }
        if ((d as { __ooRefApply?: number }).__ooRefApply != null) {
          addDiag('[poc] 用户确认立即更新外部数据')
          onApplyRefAskRef.current?.()
        }
        if (d.__ooRefDataRequest) { try { getEditorWindow()?.postMessage({ __ooRefData: ooRefItemsRef.current }, '*') } catch { /* ignore */ } }
        if (typeof d.__ooIsModified === 'boolean') { const r = modAskRef.current; modAskRef.current = null; r?.(d.__ooIsModified) }
        /* ★ 2026-10-17 未保存修改（用户指令「做是否修改了的判断 然后提示保存」）：
         *   /poc 状态翻转时主动推，宿主缓存 → 标题栏亮「未保存」+ 💾 高亮 + 关闭拦截。 */
        if (typeof d.__ooDirty === 'boolean') applyDirtyRef.current(d.__ooDirty, String(d.__ooDirtyWhy || 'poc推送'))
        const pick = (d as { __ooRefPick?: { nodeId: string; fieldId: string; fieldName: string; value?: string; mode?: string } }).__ooRefPick
        if (pick) onPickRef?.(pick.nodeId, pick.fieldId, pick.fieldName, pick.value, pick.mode)
        // ★ /poc 内嵌参数引用面板 UI 变更回传（2026-10-11 最大化/常规内容同步）
        const rui = (d as { __ooRefUi?: { q?: string; openIds?: string[] } }).__ooRefUi
        if (rui && (typeof rui.q === 'string' || Array.isArray(rui.openIds))) {
          setRefUi({ q: rui.q, openIds: rui.openIds }, true)
        }
        // ★ 章节目录回传（2026-10-09）：大纲数据 / 移动结果
        const cd = (d as { __chaptersData?: { chapters?: { i: number; level: number; text: string }[]; total?: number; err?: string } }).__chaptersData
        if (cd) {
          setChapters(cd.chapters || [])
          setChTotal(cd.total || 0)
          setChErr(cd.err ? '取大纲失败: ' + cd.err : '')
        }
        const cm = (d as { __chaptersMoved?: { ok?: boolean; err?: string } }).__chaptersMoved
        if (cm) {
          if (cm.ok) setTimeout(refreshChapters, 350)
          else { setChErr('章节移动失败: ' + (cm.err || '未知')); addDiag('[chapters] 移动失败 ' + JSON.stringify(cm)) }
        }
      }
    }
    window.addEventListener('message', onMsg)
    return () => window.removeEventListener('message', onMsg)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => { ooRefItemsRef.current = ooRefItems ?? [] }, [ooRefItems])

  // ★ 文档就绪自动拉一次章节大纲（word only）；侧栏打开时也补拉（防错过就绪时机）
  useEffect(() => {
    if (kind !== 'ooword') return
    if (/documentReady/.test(status)) refreshChapters()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [status, kind])
  useEffect(() => {
    if (kind === 'ooword' && chapOpen) refreshChapters()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [chapOpen, kind])

  // ★ 写回结果反馈已统一走宿主居中 toast（.ne-toast，2026-10-18 用户指令「左下角提示和居中提示冲突」）：
  //   原实现经 __ooRefreshMsg 让 /poc 左下角再显示一遍 —— 同一条消息两处叠加。此通道停用。

  // 写回进行中 ⇒ 按钮置灰，防重复点击
  useEffect(() => {
    try { getEditorWindow()?.postMessage({ __ooRefreshState: refreshBusy === true }, '*') } catch { /* ignore */ }
  }, [refreshBusy])

  // ★ 高亮行列开关推送（2026-10-18 用户指令）：状态在宿主（localStorage），这里只转发 /poc。
  //   /poc 页随文档重开而重建，旧消息会丢 —— 所以 prop 变化、src 变化、documentReady
  //   三个时机都要补推（延迟两次防时序缝隙）。
  const rowColHiRef = useRef(rowColHi)
  useEffect(() => { rowColHiRef.current = rowColHi }, [rowColHi])
  const pushRowColHi = useCallback(() => {
    try { getEditorWindow()?.postMessage({ __ooRowColHi: rowColHiRef.current === '1' }, '*') } catch { /* ignore */ }
  }, [])
  useEffect(() => {
    pushRowColHi()
    const t1 = setTimeout(pushRowColHi, 1500)
    const t2 = setTimeout(pushRowColHi, 4000)
    return () => { clearTimeout(t1); clearTimeout(t2) }
  }, [rowColHi, src, pushRowColHi])
  useEffect(() => {
    if (!/documentReady/.test(status)) return
    pushRowColHi()
    const t = setTimeout(pushRowColHi, 800)
    return () => clearTimeout(t)
  }, [status, pushRowColHi])

  /* ══ 外部关联引用更新策略（2026-10-18 用户指令）══════════════════════════
   * 用户原话：「打开之后检测是否有外部关联。在上方添加个按钮。点击打开设置
   *   是否自动更新 1、不检测 2、检测后弹窗提示 3、自动更新不提示」
   * 分工：检测与写回都在宿主（它有 xlsx 字节/域表/模板库）；这里只搬消息——
   *   宿主 → /poc：__ooRefUpdMode（当前策略）、__ooRefAsk（要不要更新？）、__ooRefUpdInfo（检测结果）
   *   /poc → 宿主：__ooRefUpdModeSet（改了策略）、__ooRefApply（点了「立即更新」）
   * 时机与高亮行列一致：prop 变化 + documentReady + 延迟补推（/poc 会随文档重开重建）。 */
  const refUpdModeRef = useRef(ooRefUpdMode)
  useEffect(() => { refUpdModeRef.current = ooRefUpdMode }, [ooRefUpdMode])
  const refAskRef = useRef(ooRefAsk)
  useEffect(() => { refAskRef.current = ooRefAsk ?? null }, [ooRefAsk])
  const refInfoRef = useRef(ooRefInfo)
  useEffect(() => { refInfoRef.current = ooRefInfo ?? null }, [ooRefInfo])
  const pushRefUpdState = useCallback(() => {
    const w = getEditorWindow()
    if (!w) return
    try { w.postMessage({ __ooRefUpdMode: refUpdModeRef.current || 'ask' }, '*') } catch { /* ignore */ }
    const info = refInfoRef.current
    if (info) { try { w.postMessage({ __ooRefUpdInfo: info }, '*') } catch { /* ignore */ } }
    const ask = refAskRef.current
    if (ask) { try { w.postMessage({ __ooRefAsk: ask }, '*') } catch { /* ignore */ } }
  }, [])
  useEffect(() => {
    pushRefUpdState()
    const t1 = setTimeout(pushRefUpdState, 1500)
    const t2 = setTimeout(pushRefUpdState, 4000)
    return () => { clearTimeout(t1); clearTimeout(t2) }
  }, [ooRefUpdMode, ooRefAsk, ooRefInfo, src, pushRefUpdState])
  useEffect(() => {
    if (!/documentReady/.test(status)) return
    pushRefUpdState()
    const t = setTimeout(pushRefUpdState, 900)
    return () => clearTimeout(t)
  }, [status, pushRefUpdState])
  /** 回调走 ref：message 监听只挂一次（deps []），直接闭包会拿到挂载时的旧函数 */
  const onSetRefUpdModeRef = useRef(onSetRefUpdMode)
  useEffect(() => { onSetRefUpdModeRef.current = onSetRefUpdMode }, [onSetRefUpdMode])
  const onApplyRefAskRef = useRef(onApplyRefAsk)
  useEffect(() => { onApplyRefAskRef.current = onApplyRefAsk }, [onApplyRefAsk])

  // ★ 引用格覆盖层推送（2026-10-18 着色开关改版）：颜色不写文件，宿主算好
  //   「格地址 + 颜色 + 表格几何」发过来，/poc 画纯视觉色块（开关即时生效、不重开）。
  //   推送时机与行列高亮同款：prop 变化 / src 变化 / documentReady。
  const refMarksRef = useRef(refMarks)
  useEffect(() => { refMarksRef.current = refMarks }, [refMarks])
  const pushRefMarks = useCallback(() => {
    const rm = refMarksRef.current
    try {
      getEditorWindow()?.postMessage({
        __ooRefMarks: {
          on: !!rm && rm.on === '1',
          marks: rm?.marks ?? [],
          geom: rm?.geom ?? null,
        },
      }, '*')
    } catch { /* 编辑器未就绪 */ }
  }, [])
  useEffect(() => {
    pushRefMarks()
    const t1 = setTimeout(pushRefMarks, 1500)
    const t2 = setTimeout(pushRefMarks, 4000)
    return () => { clearTimeout(t1); clearTimeout(t2) }
  }, [refMarks, src, pushRefMarks])
  useEffect(() => {
    if (!/documentReady/.test(status)) return
    pushRefMarks()
    const t = setTimeout(pushRefMarks, 800)
    return () => clearTimeout(t)
  }, [status, pushRefMarks])

  // ★ 插入引用（批次 12 · ③ A 方案）：往光标处写一条 =节点名.字段名 的引用公式文本。
  //   句柄挂在下方那个 useImperativeHandle 里（同名 ref 只能有一个 useImperativeHandle），
  //   这里只备一个转发函数供其调用。
  const insertRefDispatch = useRef<(ref: { nodeId: string; nodeName?: string; fieldId?: string; fieldName?: string; value?: unknown }) => void>(
    () => { /* 编辑器未就绪时静默忽略 */ },
  )
  useEffect(() => {
    insertRefDispatch.current = (ref) => {
      // 冻结恢复兜底已统一收口到 getEditorWindow()；发送结果写诊断抽屉，不再无声失败
      const win = getEditorWindow()
      if (!win) { addDiag('insert: 编辑器 iframe 未就绪，引用未发送'); return }
      try {
        win.postMessage({ __ooRefInsert: ref }, '*')
        addDiag('insert: 已发送 ' + (ref.nodeName || ref.nodeId) + '.' + (ref.fieldName || ''))
      } catch { /* ignore */ }
    }
  }, [])
  // ★ 面板开关状态回传 /poc（工具栏「参数引用」按钮高亮用）——只在可见性/主题变化时推送。
  //   ★ 2026-10-11 用户指令「双击参数会触发参数引用窗体的开关，去掉」：原先本 effect 依赖
  //   ooRefItems——每次双击参数落库都会重发 __ooRefSet/__ooRefState，把 /poc 内嵌面板
  //   开关一遍。现拆分：开关类消息只随 [refOpen, maximized, theme] 走，数据变化只推数据。
  useEffect(() => {
    try {
      getEditorWindow()?.postMessage({ __ooMaximized: maximized === true }, '*')
      getEditorWindow()?.postMessage({ __ooSetTheme: theme === 'dark' ? 'theme-dark' : 'theme-light' }, '*')
      getEditorWindow()?.postMessage({ __ooRefState: refOpen === true }, '*')
      getEditorWindow()?.postMessage({ __ooRefSet: refOpen === true && maximized === true }, '*')
      // ★ UI 状态同步推送（搜索词 + 展开集，最大化↔常规内容一致）
      getEditorWindow()?.postMessage({ __ooRefUi: getRefUi() }, '*')
      // ★ 节点类型：内核据此分流 API —— oosheet 走 callCommand 写单元格，ooword 走 asc_AddText 写正文
      getEditorWindow()?.postMessage({ __ooKind: kind }, '*')
    } catch { /* ignore */ }
  }, [refOpen, theme, maximized, kind])
  // ★ 参数数据推送（就绪后延迟补发两次，防编辑器未就绪丢消息）——纯数据，不含开关指令
  useEffect(() => {
    try { getEditorWindow()?.postMessage({ __ooRefData: ooRefItems ?? [] }, '*') } catch { /* ignore */ }
    const t1 = setTimeout(() => { try { getEditorWindow()?.postMessage({ __ooRefData: ooRefItemsRef.current }, '*') } catch { /* ignore */ } }, 3000)
    const t2 = setTimeout(() => { try { getEditorWindow()?.postMessage({ __ooRefData: ooRefItemsRef.current }, '*') } catch { /* ignore */ } }, 7000)
    return () => { clearTimeout(t1); clearTimeout(t2) }
  }, [ooRefItems, src])

  // ★ React 浮卡侧 UI 变更 → 推给 /poc 内嵌面板（/poc 回传应用期间不回推，防乒乓）
  useEffect(
    () =>
      subscribeRefUi((s) => {
        if (isApplyingFromPoc()) return
        try { getEditorWindow()?.postMessage({ __ooRefUi: { q: s.q, openIds: s.openIds } }, '*') } catch { /* ignore */ }
      }),
    [],
  )

  /** 上传字节到 OO 工作区并打开编辑器 */
  const uploadAndOpen = async (fileName: string, u8: Uint8Array, extOverride?: string, srcExtra = '') => {
    const workName = 'oo-' + kind + '-' + Date.now().toString(36) + '.' + (extOverride || cfg.ext)
    /* ★ 2026-10-18 修复「引用格着色成功但画面不直接变色」：冻结恢复模式（attachedFrozen 非空，
     *   二次打开同节点命中预加载）下 JSX 的 src iframe 不渲染 —— setSrc 新地址没人消费，
     *   文档永远不重载，用户看到的还是旧文件。打开新工作区前必须解除冻结；
     *   adoptedRef 一并清空，卸载时不会把旧现场再 refreeze 收藏。 */
    setAttachedFrozen(null)
    adoptedRef.current = null
    const r = await fetch('/office-offline/open', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: workName, dataBase64: b64(u8) }),
    }).then((x) => x.json())
    if (!r.ok) throw new Error(r.err || '上传失败')
    workRef.current = workName
    /* ★ 2026-10-17 记录「打开时的原始字节」——保存时用它识别主通道假成功 */
    srcSigRef.current = byteSig(u8)
    srcDigestRef.current = ((r as { digest?: OoDigest }).digest) || null
    /* ★ 2026-10-17 用户指令「打开的内容也写进日志」：把服务端算出的内容摘要一并显示 + 落盘，
     *   这样「打开是空白」能一眼分清：文件本身空（上次保存存空了）vs 渲染没出来。 */
    const dg = (r as { digest?: { text?: string; cells?: number; chars?: number } } | null)?.digest
    addDiag(`✅ 已打开: ${fileName} → ${workName} (${u8.length}B)${dg && dg.text ? '｜' + dg.text : ''}`)
    try {
      void fetch('/office-log', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ type: 'OPEN-CONTENT', detail: '节点 ' + (nodeId || '-') + ' 文件=' + fileName + ' 工作区=' + workName + ' 大小=' + u8.length + 'B ' + (dg && dg.text ? dg.text : '') }) })
    } catch (eOL7) { /* 忽略 */ }
    setSrc(`/poc?src=${workName}${srcExtra}`)
    /* ★ 递增 key → iframe 元素重建 → 引擎必定加载新工作区（见 srcSeq 定义处注释） */
    setSrcSeq((n) => n + 1)
    say('已打开 ' + fileName)
    onOpenedName?.(fileName)
  }

  // 初始引导：节点绑定过模板库 → 读模板字节；否则空文档
  useEffect(() => {
    let dead = false
    /** 接管冻结实例（2026-10-07 bug 修复：预加载完成/进行中都要接管，**绝不另开一份**） */
    const adopt = (f: { iframe: HTMLIFrameElement; workName: string; fileName: string }) => {
      setAttachedFrozen(f.iframe)
      workRef.current = f.workName
      adoptedRef.current = { fileName: f.fileName }   // ★ 关闭时按此重新冻结
      setSrc('')            // 不走 JSX src iframe（挂回的是冻结实例本体）
      say('已从预加载恢复: ' + (f.fileName || name || cfg.title))
      onOpenedName?.(f.fileName || name || '')
      addDiag('⚡ 预加载命中：编辑器从冻结现场恢复（0 等待）')
    }
    /** ★ 打开前注入域外引用值（无引用/失败都原样放行，绝不阻断打开） */
    const withRefs = async (u8: Uint8Array): Promise<Uint8Array> => {
      if (!injectRefs) return u8
      try {
        const r = await injectRefs(u8, { forOpen: true })
        if (r && r.note) addDiag('🔄 ' + r.note)
        return r && r.bytes && r.bytes.length ? r.bytes : u8
      } catch (e) {
        addDiag('⚠ 域注入异常（按原文件打开）：' + String(e).slice(0, 80))
        return u8
      }
    }
    async function boot() {
      try {
        /* ★ 网页端无 Office 服务时不做任何打开动作：直接给只读提示。
         *   注意必须在「预加载命中」之前判断——否则会接管一个永远连不上 DS 的冻结实例。 */
        const caps = await ensureCaps()
        if (!caps.oo) {
          if (dead) return
          setOoOff(true)
          say('网页版不支持在线 Office 编辑')
          addDiag('ℹ 服务端 /healthz 报 oo=false（未部署 OnlyOffice 文档服务）→ 本次不加载编辑器')
          auditFeedback('Office 不可用', '网页版无 OnlyOffice 文档服务，已降级为只读提示')
          return
        }
        // ★ 非预加载来源（正常上传/样例打开）：清掉 adopt 标记——关闭时按原样销毁
        adoptedRef.current = null
        // ★ 预加载命中（已完成）：立即接管
        if (nodeId && hasFrozen(nodeId)) {
          const f = takeFrozen(nodeId)
          if (f) { adopt(f); return }
        }
        // ★ 预加载进行中（2026-10-07 bug：原实现不等接管而另开一份文档——
        //   用户看到「重新打开」，且留下孤儿冻结实例）：等待预加载完成后接管
        if (nodeId && preloadState(nodeId) === 'loading') {
          say('后台预加载进行中——完成后将直接接管（不重复打开）…')
          addDiag('⚡ 检测到预加载进行中，等待接管…')
          const f = await waitFrozen(nodeId, 60000)
          if (dead) return
          if (f) { adopt(f); return }
          addDiag('⚠ 预加载等待超时/失败，转正常打开流程')
        }
        if (bytes && bytes.length) {
          await uploadAndOpen(name || cfg.title, await withRefs(bytes))
          return
        }
        /* ★ 2026-10-11 用户指令「保存过就加载保存过的，不再加载样例」：节点内置内容（base64，
         *   随画布节点配置持久化）优先于模板库——真正随画布文件走，换机/清缓存不丢。 */
        if (ooBytes64) {
          try {
            const bin = atob(ooBytes64)
            const u8 = new Uint8Array(bin.length)
            for (let i2 = 0; i2 < bin.length; i2++) u8[i2] = bin.charCodeAt(i2)
            if (u8.length) {
              /* ★ 2026-10-30 导入即建档（用户令「彻底解决版式乱」）：boot 时字节=
                 导入原件（新节点）或上次产物。节点无档案时立即从当前字节提取固化——
                 新文档天生带正确档案，之后保存链任何成败都有版式兜底。
                 （原「保存时建档」会被失败产物建出 0 行距坏档案 → 永不升级死循环） */
              try {
                const nid = nodeId as string
                const st0 = useBaseStore.getState()
                const cfg0 = st0.tables[nid]?.config as { ooLineProfile?: unknown } | undefined
                if (!cfg0?.ooLineProfile) {
                  const { extractOoLayoutProfile } = await import('../core/ooLayoutProfile')
                  const prof = await extractOoLayoutProfile(u8)
                  if (prof && prof.lineList.some((x) => x.line)) {
                    st0.setNodeConfig(nid, { ooLineProfile: prof })
                    if (!dead) addDiag('📐 版式档案已建（导入即建）: 行距段=' + prof.lineList.filter((x) => x.line).length + ' 分页=' + prof.pageBreakIdx.length + ' 页面=' + (prof.sectionProps ? '√' : '默认'))
                  }
                }
              } catch { /* 建档失败不影响 boot */ }
              const inj = await withRefs(u8)
              await uploadAndOpen(name || cfg.title, inj)
              addDiag('✅ 已加载节点保存的内容（' + u8.length + 'B' + (inj.length !== u8.length ? ' → 注入引用后 ' + inj.length + 'B' : '') + '，随画布存储）')
              return
            }
          } catch (eB64) {
            if (!dead) addDiag('⚠ 节点内置内容解码失败，转模板库: ' + String(eB64).slice(0, 60))
          }
        }
        /* ★ 2026-10-29 补 OO 内容 IndexedDB 索引加载分支（schema 里早写了「优先级同
         *   ooBytes64」，boot 却一直没实现）——拖入导入/文件夹拖入只落 ooBytesDbId
         *   + 合成 tpl id（模板库无文件），没有这个分支重开必炸 tpl:read。 */
        if (ooBytesDbId) {
          try {
            const { getBlob } = await import('../core/pdfBlobStore')
            const u8d = await getBlob(ooBytesDbId)
            if (dead) return
            if (u8d && u8d.length) {
              /* ★ 2026-10-31 拖入导入走本分支（非 ooBytes64）——建档逻辑同样必须挂这里，
                 否则拖入的文档永远没有档案（用户实测：拖入→改→保存→版式乱 复现） */
              try {
                const st0 = useBaseStore.getState()
                const nid = nodeId as string
                const cfg0 = st0.tables[nid]?.config as { ooLineProfile?: unknown } | undefined
                if (!cfg0?.ooLineProfile) {
                  const { extractOoLayoutProfile } = await import('../core/ooLayoutProfile')
                  const prof = await extractOoLayoutProfile(u8d)
                  if (prof && (prof.lineList.some((x) => x.line) || prof.sectionProps)) {
                    st0.setNodeConfig(nid, { ooLineProfile: prof })
                    if (!dead) addDiag('📐 版式档案已建（拖入导入即建）: 段=' + prof.lineList.length + ' 分页=' + prof.pageBreakIdx.length + ' 页面=' + (prof.sectionProps ? '√' : '默认'))
                  }
                }
              } catch { /* 建档失败不影响 boot */ }
              const inj = await withRefs(u8d)
              await uploadAndOpen(name || cfg.title, inj)
              addDiag('✅ 已加载节点保存的内容（IndexedDB ' + u8d.length + 'B，随画布存储）')
              return
            }
          } catch (eDb) {
            if (!dead) addDiag('⚠ IndexedDB 内容读取失败，转模板库: ' + String(eDb).slice(0, 60))
          }
        }
        if (ooTemplateId) {
          /* ★ 2026-10-29 模板库缺文件不再炸 boot：tpl:read 抛「模板文件不存在」只降级——
             继续走空白兜底（真空白 docx/xlsx），不再卡「打开失败」+ /poc 默认 lidu.docx -4 */
          let saved: Uint8Array | null = null
          try { saved = await getTemplateStore().read(ooTemplateId) } catch { saved = null }
          if (dead) return
          if (saved && saved.length) {
            await uploadAndOpen(name || cfg.title, await withRefs(saved))
            return
          }
          if (!dead) addDiag('⚠ 模板库缺文件（' + ooTemplateId + '）——转节点内置内容/空白兜底')
        }
        if (dead) return
        /* ★ 2026-10-11 用户指令：未导入、未保存过的 OO 节点，首次打开 = 空白文档。
         *   默认样例自动种入已移除（要样例用 📂 导入本地文件）。 */
        if (!dead) addDiag('未导入/未保存过——空白文档开始（📂 可导入本地文件）')
        /* ★ 2026-10-12 用户指令「新建节点打开空白新建文件」：工作表节点的空白改为
         *   服务端构建的真空白 xlsx——不再依赖 /poc 空源（DS 对未知键默认给 Word 样例）。 */
        if (kind === 'oosheet' && !dead) {
          try {
            const blankName = 'oo-oosheet-blank-' + Date.now().toString(36) + '.xlsx'
            const rb = await fetch('/office-offline/saveFromValues', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              /* ★ 2026-10-17：mode 只用于服务端日志区分来源。原先不传 ⇒ 日志里
               *   「新建空白节点造工作簿」和「用户点保存」长得一模一样（都是 模式=node 行数=0），
               *   排查时极易误判成「保存链在跑」——本轮就被骗过一次。 */
              body: JSON.stringify({ name: blankName, rows: [], mode: 'blank' }),
            }).then((x) => x.json())
            if (rb && rb.ok && rb.dataBase64) {
              const bin2 = atob(rb.dataBase64)
              const u82 = new Uint8Array(bin2.length)
              for (let i2 = 0; i2 < bin2.length; i2++) u82[i2] = bin2.charCodeAt(i2)
              await uploadAndOpen('空白工作簿.xlsx', await withRefs(u82))
              return
            }
          } catch (eBL) { if (!dead) addDiag('⚠ 空白工作簿构建失败，转空白页: ' + String(eBL).slice(0, 60)) }
        }
        /* ★ 2026-10-23 PDF 节点（kind='pdf'）：数据桥——pdfItems[0] 的真实字节
           直接上传进 OO 9.4 编辑（原生 pdf 编辑）。此前 pdf 节点不命中 ooword
           空白分支 → 落 /poc 默认 lidu.docx → -4（用户实测打不开的根因）。 */
        if (kind === 'pdf' && !dead) {
          const items = ((useBaseStore.getState().tables[nodeId!]?.config ?? {}) as { pdfItems?: { id?: string; name: string; b64?: string; inDb?: boolean }[] }).pdfItems ?? []
          const first = items[0]
          if (first?.b64 || (first as { inDb?: boolean }).inDb) {
            let u8p: Uint8Array
            if (first.b64) {
              const bin = atob(first.b64)
              u8p = new Uint8Array(bin.length)
              for (let i = 0; i < bin.length; i++) u8p[i] = bin.charCodeAt(i)
            } else {
              const { getBlob } = await import('../core/pdfBlobStore')
              const fromDb = await getBlob('pdfitem-' + first.id)
              if (!fromDb) { say('文件数据缺失: ' + (first.name || first.id)); return }
              u8p = fromDb
            }
            /* ★ 2026-10-23 用户实测「打开前卡一分钟才能动」：OO 把 PDF 解析成可编辑
               格式是固有成本（5MB 级需 10~30s，二次打开走缓存会快）——明确进度
               提示，避免误判卡死。 */
            say('正在把 PDF 转换为可编辑格式…（大文件需 10~30 秒，属正常解析耗时）')
            addDiag('⏳ PDF 转换中: ' + (first.name || '文档.pdf') + ' (' + Math.round(u8p.length / 1024) + 'KB)')
            await uploadAndOpen(first.name || '文档.pdf', u8p, 'pdf', '&docType=pdfviewer')
            return
          }
          if (!dead) { addDiag('⚠ PDF 节点未导入内容——先拖入本地 PDF 文件'); say('空白 PDF（📂 拖入本地 PDF 开始）') }
          workRef.current = ''
          setBlankHint(true)
          return
        }
        /* ★ 2026-10-22 修 ooword 打不开（errorCode -4 下载失败）：空白兜底与 oosheet 对齐——
         *   原 setSrc('/poc?src=') 空 src 时 /poc 默认加载 /work/lidu.docx（老机器工作区
         *   残留文件，本机没有）→ DocService 下载 404 → -4。现在宿主用 docx 库生成
         *   真空白文档走 uploadAndOpen（与 oosheet 的 saveFromValues 同语言，零 /work 依赖）。 */
        if (kind === 'ooword' && !dead) {
          let u8w: Uint8Array | null = null
          /* 主路：docx 库（Packer 浏览器必须用 toBlob——toBuffer 是 Node API） */
          try {
            const { Document, Packer, Paragraph } = await import('docx')
            const doc = new Document({ sections: [{ children: [new Paragraph({ children: [] })] }] })
            const blob = await Packer.toBlob(doc)
            u8w = new Uint8Array(await blob.arrayBuffer())
          } catch (eDW) {
            /* ★ 2026-10-23 兜底：docx 库 chunk 加载失败（动态 import 病）时用
               JSZip 手造最小 OOXML（合法结构：ContentTypes/.rels/document.xml）。
               ★ 失败文案双写 pushLog（进日志流）+ addDiag（诊断抽屉）——
                 此前只 addDiag，日志流看不到 → -4 排查盲飞一轮。 */
            pushLog('[ooword] 空白文档构建失败（docx 库）: ' + String(eDW).slice(0, 120), 'err')
            try {
              const JSZip = (await import('jszip')).default
              const zip = new JSZip()
              zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
              zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
              zip.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>')
              u8w = await zip.generateAsync({ type: 'uint8array' })
            } catch (eJZ) {
              pushLog('[ooword] 空白文档构建失败（JSZip 兜底也失败）: ' + String(eJZ).slice(0, 120), 'err')
              if (!dead) addDiag('⚠ 空白文档构建失败: ' + String(eJZ).slice(0, 80))
            }
          }
          if (u8w) {
            await uploadAndOpen('空白文档.docx', await withRefs(u8w))
            return
          }
        }
        if (!dead) setBlankHint(true)
        workRef.current = ''
        setSrc(`/poc?src=`)
        say('空白文档（📂 打开本地文件开始）')
        onOpenedName?.('')
      } catch (e) {
        if (!dead) {
          say('打开失败: ' + e)
          addDiag('❌ ' + String(e))
        }
      }
    }
    void boot()
    return () => {
      dead = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [kind, ooTemplateId, nodeId])

  // ★ 冻结实例挂回：attachedFrozen 就绪后 append 到容器（同 document 移动不触发重载）
  //   ★ 2026-10-07 bug 修复「预加载完成要根据打开的窗口适配」：
  //     ① 铺满容器（预加载时在屏幕外 host 里以固定尺寸加载）
  //     ② 挂回后向引擎窗口 dispatch resize——OO 引擎监听 window resize 重算画布布局，
  //        两次延迟触发（等 React 布局稳定），否则引擎保持预加载时的旧尺寸。 */
  useEffect(() => {
    const f = attachedFrozen
    if (!f || !containerRef.current) return
    containerRef.current.appendChild(f)
    f.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;border:none;display:block;background:#87CEEB'
    iframeRef.current = f
    const fire = () => { try { f.contentWindow?.dispatchEvent(new Event('resize')) } catch { /* ignore */ } }
    const t1 = setTimeout(fire, 120)
    const t2 = setTimeout(fire, 600)
    /* ★ 2026-10-18：解除冻结时把旧 iframe 元素从容器里摘掉（uploadAndOpen 会 setAttachedFrozen(null)
     *   触发本 cleanup）——否则它作为后 append 的兄弟节点压在新 src iframe 上面，画面还是旧的。 */
    return () => {
      clearTimeout(t1); clearTimeout(t2)
      try { if (f.parentNode) f.parentNode.removeChild(f) } catch { /* ignore */ }
      /* iframeRef 不在此清空：卸载路径的 refreeze effect（声明在后、清理也在后）
       * 还要靠它拿到冻结 iframe 收回隐藏 host；detach 路径新 iframe 挂载时 ref 会自行覆盖。 */
    }
  }, [attachedFrozen])

  // ★ 关闭节点（组件卸载）：预加载来源的实例**重新冻结**收好——
  //   下次打开原地恢复，绝不重复打开（2026-10-07 用户指令）。
  //   ⚠ effect cleanup 在 React 移除 DOM 之前执行——此刻把 iframe 移回隐藏 host，
  //   引擎现场得以存活；否则随容器一起销毁。
  useEffect(() => {
    return () => {
      if (!nodeId || !adoptedRef.current) return
      const f = iframeRef.current
      const wn = workRef.current
      if (f && wn) {
        try { refreeze(nodeId, { iframe: f, workName: wn, kind, fileName: adoptedRef.current.fileName }) } catch { /* ignore */ }
      }
    }
  }, [nodeId, kind])

  useImperativeHandle(ref, () => ({
    /** 插入引用公式文本（批次 12·③ A 方案）：=节点名.字段名 */
    insertRef: (r: { nodeId: string; nodeName?: string; fieldId?: string; fieldName?: string; value?: unknown; mode?: string }) =>
      insertRefDispatch.current(r),
    /** 宿主操作日志（2026-10-10）：点参数/关联落库等 → 诊断抽屉 */
    opLog: (msg: string) => addDiag(msg),
    ackSaved: () => {
      const ackWin = getEditorWindow()
      try { ackWin?.postMessage({ __ooSavedAck: 1 }, '*') } catch { /* 忽略 */ }
      /* ★ 2026-10-17 双保险：保存后自动重开文档时 postMessage 可能打在正被替换的窗口上
       *   丢失 → /poc 的脏标志留在 true → 用户再改因「状态没翻转」永远不推 → 不提示。
       *   同源可直接调函数，绕过消息通道（实测 out2 日志：20:45:38 保存后无 savedAck 记录）。 */
      try {
        const ackAny = ackWin as unknown as { __setOoDirty?: (v: boolean, why: string) => void } | null
        ackAny?.__setOoDirty?.(false, 'savedAck')
      } catch { /* 忽略 */ }
      /* 本地也立即清：不依赖 /poc 回推的时序，保存成功后 UI 马上变干净 */
      applyDirtyRef.current(false, '本机保存完成')
    },
    isModified: () => {
      /* ★ 2026-10-17：优先用 /poc 主动推来的缓存 —— 即时、无超时、不会假阴性。
       *   缓存说「没有」时才问一次兜底（400ms），防止推送丢失导致漏提示。 */
      if (dirtyRef.current) return Promise.resolve(true)
      return new Promise<boolean>((resolve) => {
        modAskRef.current = resolve
        try { getEditorWindow()?.postMessage({ __ooIsModifiedAsk: 1 }, '*') } catch { resolve(false) }
        setTimeout(() => { if (modAskRef.current) { modAskRef.current = null; resolve(dirtyRef.current) } }, 400)
      })
    },
    openFile: async (f: File, extOverride?: string) => {
      say('读取 ' + f.name + '…')
      addDiag(`📂 选择文件: ${f.name} (${Math.round(f.size / 1024)}KB)`)
      try {
        let buf: Uint8Array = new Uint8Array(await f.arrayBuffer())
        /* ★ 2026-10-19 导入的 xlsx 摘掉 fullCalcOnLoad：外部文件常带该标志，引擎
         *   加载后自行重算 ⇒ 刚导入就「有未保存修改」。缓存由导入文件自带（正确值）。 */
        if (kind === 'oosheet' || extOverride === 'xlsx') {
          try { buf = await stripFullCalcFlag(buf) } catch { /* 摘除失败不影响打开 */ }
        }
        await uploadAndOpen(f.name, buf, extOverride)
      } catch (e) {
        say('打开失败: ' + e)
        addDiag('❌ 打开失败: ' + e)
      }
    },
    requestSave: (mode, sid) => {
      /* ★ 2026-10-17：把 sid 一并下发，/poc 侧后续所有 __lg 都会带上它 */
      if (sid) saveSidRef.current = String(sid)
      try {
        getEditorWindow()?.postMessage({ __ooSaveReq: 1, __ooSaveMode: mode || 'node', __ooSid: saveSidRef.current }, '*')
        logStep('S2-REQ', '已下发 __ooSaveReq mode=' + (mode || 'node') + ' sid=' + (saveSidRef.current || '-') + ' 窗口=' + (getEditorWindow() ? 'ok' : '❌ 未就绪'))
      } catch (eRS) { logStep('S2-REQ-E', '❌ postMessage 失败: ' + String(eRS).slice(0, 120)) }
    },
    saveBack: async () => {
      if (!workRef.current) {
        addDiag('⚠ 没有已打开的工作区文件，无从保存')
        logStep('S10-E', '❌ 主通道无法启动：workRef 为空（编辑器没打开过工作区文件）')
        return null
      }
      logStep('S10', '主通道(OO自身)开始 工作区=' + workRef.current)
      try {
        // ★ 保存链 v3：① 深帧(DocsAPI内层) asc_Save 推送变更 ② DS forcesave（正确 key）③ 轮询取新产物
        try {
          const fr = getEditorWindow() as { document?: Document; __docKey?: string } | null
          const deep = fr && fr.document ? (fr.document.querySelector('iframe[name="frameEditor"]') as HTMLIFrameElement | null) : null
          const deepWin = deep && deep.contentWindow as { editor?: { asc_Save?: () => void; isDocumentModified?: () => boolean }, __docKey?: string } | null
          const api = deepWin && deepWin.editor
          if (api && typeof api.asc_Save === 'function') {
            api.asc_Save()
            logStep('S10s', 'asc_Save 已调用（通知引擎把改动序列化给服务端）')
          } else {
            logStep('S10s-E', '⚠ 深帧无 editor.asc_Save（deep帧=' + !!deep + ' editor=' + (api ? '有' : '无') + '）——此步跳过，不影响后面 forcesave')
          }
          ;(window as { __ooDeepWin?: unknown }).__ooDeepWin = deepWin
        } catch (eS10s) { logStep('S10s-E', '❌ asc_Save 异常: ' + String(eS10s).slice(0, 160)) }
        try {
          const fr = getEditorWindow() as { document?: Document; __docKey?: string } | null
          const deep = fr && fr.document ? (fr.document.querySelector('iframe[name="frameEditor"]') as HTMLIFrameElement | null) : null
          const deepWin = deep && deep.contentWindow as { __docKey?: string } | null
          /* ★ 2026-10-17 修复：__docKey 由 oo-core 挂在 /poc 页 window 上（注释原话「宿主经此读取」），
           *   原实现却去 deep 帧读 ⇒ 读到陈旧值（无头实测 forcesave 发给了 oo-lidu.docx，
           *   当前工作区是 oo-oosheet-*.xlsx ⇒ DS 永远不出产物 ⇒ OO 通道恒「轮询超时」）。 */
          const docKey = (fr && fr.__docKey) || (deepWin && deepWin.__docKey)
          if (docKey) {
            logStep('S10a', 'forcesave key=' + docKey + ' (poc页=' + !!(fr && fr.__docKey) + ' deep帧=' + !!(deepWin && deepWin.__docKey) + ' 工作区=' + workRef.current + ')')
            const __fs0 = Date.now()
            const __fr2 = await fetch('/doc/command', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ c: 'forcesave', key: docKey }),
            }).catch((e) => { logStep('S10a-E', '❌ forcesave 请求异常: ' + String(e).slice(0, 140)); return null })
            /* ★ 2026-10-17：把 DS 的应答体也记下来。实测恒为 400 ——
             *   因为编辑器走的是「自研 socket 假扮 DocService」，DS 侧根本没有这个 key 的会话，
             *   command 接口自然拒绝。这是主通道产不出内容的上游原因，不是本机的偶发故障。 */
            let __fb = ''
            try { __fb = __fr2 ? (await __fr2.text()).slice(0, 240) : '' } catch { __fb = '(响应体读取失败)' }
            logStep('S10a', 'forcesave 应答 http=' + (__fr2 ? __fr2.status : '请求失败') + ' 耗时=' + (Date.now() - __fs0) + 'ms' + (__fb ? ' 应答体=' + __fb : '') + (__fr2 && __fr2.status === 400 ? ' ← DS 无此 key 会话（自研 socket 不建会话），主通道注定无产物' : ''))
            addDiag('⟳ 已触发 forcesave（' + docKey + '）')
          } else {
            logStep('S10a-E', '❌ 拿不到 docKey（poc页与deep帧都没有）→ forcesave 不会发出')
          }
        } catch (eS10a) { logStep('S10a-E', '❌ docKey/forcesave 段异常: ' + String(eS10a).slice(0, 160)) }
        /* ★ 2026-10-17 轮询再砍：Editor.bin 是本机 x2t 从**源文件**转出来的（见 oo-core 965-972 行
         *   「Editor.bin 已就位」门禁），里面从不包含编辑内容 ⇒ 产物恒等于打开时的原文件。
         *   这条通道在架构上就存不到东西，等 6 轮 8 秒纯属让用户干等。
         *   实测真实转换只 0.4s：1 秒内没产物就永远不会有了，立即降级提取通道。 */
        const waits = [400, 600]
        for (let i = 0; i < waits.length; i++) {
          await new Promise((r) => setTimeout(r, waits[i]))
          const __t0 = Date.now()
          let r: { ok?: boolean; err?: string; dataBase64?: string; digest?: OoDigest } | null = null
          let __err = ''
          try {
            r = await fetch('/office-offline/save', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ name: workRef.current }),
            }).then((x) => x.json())
          } catch (eF) { __err = String(eF).slice(0, 120) }
          const __ms = Date.now() - __t0
          /* ★ 2026-10-17：每次轮询都要留痕（含失败原因与耗时）——
           *   「主通道为什么没产物」过去只能看到最后一句 S11-E，
           *   分不清是接口 500、文件不存在、还是转了但为空。 */
          logStep('S10b', '轮询#' + (i + 1) + ' 耗时=' + __ms + 'ms ok=' + (r ? r.ok : '请求异常') + (__err ? ' 请求错误=' + __err : '') + (r && r.err ? ' 服务端err=' + String(r.err).slice(0, 120) : '') + (r && r.dataBase64 ? ' 产物=' + Math.round(String(r.dataBase64).length * 0.75) + 'B' : ' 无产物'))
          if (r && r.ok && r.dataBase64) {
            const u8 = b64ToU8(r.dataBase64)
            const dg: OoDigest | null = (r && r.digest) || null
            /* ★ 2026-10-17：产物指纹与「打开时原始字节」一致 ⇒ 转回来的还是原文件本身 */
            const same = !!srcSigRef.current && byteSig(u8) === srcSigRef.current
            addDiag(`💾 取回编辑产物 ${u8.length}B（第 ${i + 1} 次轮询）${dg && dg.text ? '｜' + dg.text : ''}${same ? '｜⚠ 与打开时字节完全相同' : ''}`)
            logStep('S11', '主通道出产物 ' + u8.length + 'B 第' + (i + 1) + '次轮询｜' + (dg && dg.text ? dg.text : '无摘要') + '｜sameAsSource=' + (same ? 'yes' : 'no') + '｜打开时指纹=' + (srcSigRef.current || '（无，可能来自冻结恢复）') + '｜打开时内容=' + (srcDigestRef.current && srcDigestRef.current.text ? srcDigestRef.current.text : '（无）') + '｜内容相同=' + (digestSame(dg, srcDigestRef.current) ? 'YES←这次编辑没进产物' : 'no'))
            onSaved(u8)
            return { bytes: u8, digest: dg, sameAsSource: same, srcDigest: srcDigestRef.current }
          }
        }
        addDiag('❌ 轮询超时：DS 未生成新产物')
        logStep('S11-E', '主通道失败：轮询' + waits.length + '次无产物（详见各 S10b 的 ok/err）→ 降级提取通道')
        return null
      } catch (e) {
        addDiag('❌ 保存失败: ' + e)
        logStep('S10-E', '❌ 主通道异常: ' + String(e).slice(0, 200) + ' | ' + String((e as Error)?.stack || '').slice(0, 240))
        return null
      }
    },
    currentBytes: async () => {
      if (!workRef.current) return null
      try {
        const r = await fetch('/office-offline/save', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: workRef.current }),
        }).then((x) => x.json())
        return r.ok ? b64ToU8(r.dataBase64) : null
      } catch {
        return null
      }
    },
    /* ★ 2026-10-29 样式化提取：postMessage 问 /poc → poc 取深帧引擎 callCommand 读文档模型
     *  （段落文字+run 格式 JSON）→ parent.postMessage 回 __ooStyledExtractData。5s 超时判 null。 */
    styledExtract: () => new Promise<{ paras: unknown[] | null; err?: string; meta?: string } | null>((resolve) => {
      const w = getEditorWindow()
      if (!w) { resolve({ paras: null, err: 'host:编辑器窗口不可达' }); return }
      let done = false
      const onMsg = (e: MessageEvent) => {
        const d = e.data as { __ooStyledExtractData?: unknown; __ooStyledExtractErr?: string; __ooStyledExtractMeta?: unknown } | null
        if (d && typeof d === 'object' && '__ooStyledExtractData' in d) {
          done = true
          window.removeEventListener('message', onMsg)
          const arr = (d as { __ooStyledExtractData?: unknown }).__ooStyledExtractData
          let meta: string | undefined
          try { const m = (d as { __ooStyledExtractMeta?: unknown }).__ooStyledExtractMeta; if (m) meta = JSON.stringify(m) } catch { /* ignore */ }
          resolve({ paras: Array.isArray(arr) ? (arr as unknown[]) : null, err: (d as { __ooStyledExtractErr?: string }).__ooStyledExtractErr, meta })
        }
      }
      window.addEventListener('message', onMsg)
      try { w.postMessage({ __ooStyledExtractAsk: 1 }, '*') } catch (eP) { window.removeEventListener('message', onMsg); resolve({ paras: null, err: 'host:postMessage 异常 ' + String(eP).slice(0, 80) }); return }
      setTimeout(() => { if (!done) { window.removeEventListener('message', onMsg); resolve({ paras: null, err: 'host:5s 超时（poc 未应答）' }) } }, 5000)
    }),
  }))

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', minHeight: 0 }}>
      <div ref={containerRef} style={{ flex: 1, minHeight: 0, position: 'relative' }}>
        {(saving || savedFlash) && (
          /* ★ 保存中=「💾 保存中…」；完成后=「✓ 已保存」停留 2.5s（用户指令：不要消失太快）。不可选择 */
          <div style={{ position: "absolute", bottom: 16, left: "50%", transform: "translateX(-50%)", zIndex: 35, background: "rgba(22,27,34,.72)", color: "#e8edf2", padding: "6px 18px", borderRadius: 16, fontSize: 12, pointerEvents: "none", border: "1px solid " + (saving ? "rgba(88,166,255,.35)" : "rgba(63,185,80,.45)"), backdropFilter: "blur(4px)" }}>
            {saving ? '💾 保存中…（写入节点与画布）' : '✓ 已保存'}
          </div>
        )}
        {ooOff && (
          /* ★ 2026-10-17 网页端降级提示（服务端 oo:false）。
             不做成 modal：用户还要能拖动画布、看别的节点，这里只占住编辑器区域。 */
          <div style={{ position: 'absolute', inset: 0, zIndex: 40, background: 'rgba(0,0,0,.28)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }}>
            <div style={{ maxWidth: 420, background: 'var(--ed-bg1, #161b22)', border: '1px solid var(--ed-border, #30363d)', borderRadius: 12, padding: '18px 20px', boxShadow: '0 12px 40px rgba(0,0,0,.45)' }}>
              <div style={{ fontSize: 14, fontWeight: 600, marginBottom: 8 }}>
                🌐 浏览器版不支持在线 {kind === 'oosheet' ? '表格' : '文档'} 编辑
              </div>
              <div style={{ fontSize: 12.5, lineHeight: 1.7, opacity: 0.78 }}>
                当前部署没有 OnlyOffice 文档服务（它是随桌面端安装的 Windows 组件）。
                <br />
                节点的文档内容仍随画布一起保存；要编辑请用桌面端打开本画布，或把文件导出后在本地改。
              </div>
              <div style={{ marginTop: 12, fontSize: 11.5, opacity: 0.5 }}>
                其余能力（画布、表格节点、参数、公式、.wfc 导入导出）不受影响
              </div>
            </div>
          </div>
        )}
        {blankHint && (
          /* ★ 2026-10-12 用户指令：空白打开时给出「打开或新建」选择（按类型文案） */
          <div style={{ position: "absolute", inset: 0, zIndex: 40, background: "rgba(0,0,0,.35)", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <div style={{ background: "var(--ed-bg1, #161b22)", border: "1px solid var(--ed-border, #30363d)", borderRadius: 12, padding: "20px 24px", minWidth: 320, boxShadow: "0 12px 40px rgba(0,0,0,.5)" }}>
              <div style={{ fontSize: 14, fontWeight: 600, marginBottom: 6 }}>
                📄 空白新建{kind === 'oosheet' ? '工作簿' : '文档'}
              </div>
              <div style={{ fontSize: 12, opacity: 0.75, marginBottom: 14 }}>此节点尚未保存过内容——选择打开本地文件，或新建空白开始编辑</div>
              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={() => { setBlankHint(false); blankFileRef.current?.click() }} style={{ flex: 1, padding: "9px 0", borderRadius: 8, border: "1px solid var(--ed-border, #30363d)", background: "rgba(88,166,255,.15)", color: "inherit", cursor: "pointer", fontSize: 13 }}>📂 打开本地文件</button>
                <button onClick={async () => { try { const r = await fetch('/office-recent?ext=' + (kind === 'oosheet' ? 'xlsx' : 'docx')).then((x) => x.json()); setRecentFiles(r && r.files ? r.files : []) } catch (eR) { setRecentFiles([]) } }} style={{ flex: 1, padding: "9px 0", borderRadius: 8, border: "1px solid var(--ed-border, #30363d)", background: "rgba(139,148,158,.12)", color: "inherit", cursor: "pointer", fontSize: 13 }}>🕘 最近使用文件</button>
                <button onClick={() => setBlankHint(false)} style={{ flex: 1, padding: "9px 0", borderRadius: 8, border: "1px solid var(--ed-border, #30363d)", background: "transparent", color: "inherit", cursor: "pointer", fontSize: 13 }}>📄 新建空白开始</button>
        {recentFiles && recentFiles.length > 0 && (
          <div style={{ marginTop: 10, maxHeight: 150, overflowY: "auto", border: "1px solid var(--ed-border, #30363d)", borderRadius: 8 }}>
            {recentFiles.map((f) => (
              <div key={f.name} onClick={async () => { try { const r = await fetch("/office-offline/openrecent", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ file: f.name }) }).then((x) => x.json()); if (r && r.ok && r.dataBase64) { const bin = atob(r.dataBase64); const u8 = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i); setBlankHint(false); await uploadAndOpen(f.name, u8) } } catch (eO) { addDiag("❌ 最近文件加载失败: " + String(eO).slice(0, 60)) } }} style={{ padding: "7px 10px", borderBottom: "1px solid rgba(48,54,61,.5)", cursor: "pointer", fontSize: 12, display: "flex", justifyContent: "space-between", gap: 8 }}>
                <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>{f.name}</span>
                <span style={{ opacity: 0.6, flexShrink: 0 }}>{new Date(f.mtime).toLocaleString("zh-CN", { hour12: false }).slice(5, 16)}</span>
              </div>
            ))}
          </div>
        )}
              </div>
            </div>
          </div>
        )}
        <input ref={blankFileRef} type="file" accept=".xlsx,.docx" style={{ display: "none" }}
          onChange={async (e) => { const f = e.target.files && e.target.files[0]; if (!f) return; setBlankHint(false); try { const buf = new Uint8Array(await f.arrayBuffer()); await uploadAndOpen(f.name, buf) } catch (eI) { addDiag("❌ 导入失败: " + String(eI).slice(0, 80)) } e.target.value = "" }} />
        {/* ★ 数据锁定（2026-10-07）：dataLocked=true → poc 以 mode=view 打开（编辑器只读）。
            锁定切换会改 src → iframe 重挂（热缓存下 ~1.3s，可接受）。
            attachedFrozen（预加载恢复）时不渲染这个 JSX iframe——挂回的是冻结实例本体。 */}
        {!attachedFrozen && (
          <>
          {anyDragging && (
            <div style={{ position: 'absolute', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#87CEEB', color: '#0b5', fontSize: 13 }}>
              拖动中…（松手恢复编辑器）
            </div>
          )}
          <iframe
            key={'oo-frame-' + srcSeq}
            ref={iframeRef}
            src={(src || '/poc?src=') + (initialThemeRef.current === 'dark' ? '&theme=dark' : '&theme=light') + (dataLocked ? '&mode=view' : '&mode=edit')}
            title="OnlyOffice"
            style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', border: 'none', background: '#87CEEB', visibility: anyDragging ? 'hidden' : 'visible' }}
          />
          </>
        )}
        {attachedFrozen && dataLocked && (
          /* 冻结恢复 + 锁定的组合：poc 已以 edit 模式打开，v1 用拦截层挡鼠标输入
             （键盘防护不完整——完整只读需解锁后重开，走 mode=view 路径） */
          <div style={{ position: 'absolute', inset: 0, zIndex: 9, background: 'transparent', pointerEvents: 'auto' }} />
        )}
        {status && !src && !attachedFrozen && <div style={{ position: 'absolute', inset: 0, padding: 24, color: '#888', pointerEvents: 'none' }}>{status}</div>}
        {/* ★ 章节目录侧栏（2026-10-09）：仅 word + 最大化显示，左缘叠加。
            点击标题跳转；拖拽标题重排章节（到条目前的插入线为准）；↻ 重取大纲。
            编号：请对标题挂 OO 原生多级编号（开始→多级编号），拖拽后自动刷新。 */}
        {kind === 'ooword' && maximized && (
          <div style={{ position: 'absolute', left: 0, top: 0, bottom: 30, zIndex: 8, width: chapOpen ? 232 : 26, pointerEvents: 'none' }}>
            <button
              onClick={() => setChapOpen(!chapOpen)}
              title={chapOpen ? '收起章节目录' : '章节目录（大纲/跳转/拖拽排序）'}
              style={{
                position: 'absolute', left: 0, top: 0, pointerEvents: 'auto', width: 26, height: 72,
                border: '1px solid #2a3138', borderRadius: '0 4px 4px 0', borderLeft: 'none',
                background: chapOpen ? '#1a222b' : 'rgba(20,26,33,.92)', color: chapOpen ? '#58a6ff' : '#8b949e',
                cursor: 'pointer', fontSize: 12, lineHeight: 1.3, letterSpacing: 2,
                writingMode: 'vertical-lr', padding: '6px 0',
              }}
            >章节</button>
            {chapOpen && (
              <div style={{ position: 'absolute', left: 26, top: 0, bottom: 0, right: 6, pointerEvents: 'auto', background: 'rgba(20,26,33,.97)', borderRight: '1px solid #2a3138', display: 'flex', flexDirection: 'column', color: '#c9d1d9', borderRadius: '0 6px 6px 0', overflow: 'hidden' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '6px 8px', borderBottom: '1px solid #2a3138' }}>
                  <strong style={{ fontSize: 12 }}>章节</strong>
                  <span style={{ flex: 1 }} />
                  <button className="bar-btn" style={{ fontSize: 11, padding: '1px 6px' }} title="重新读取大纲" onClick={refreshChapters}>↻</button>
                </div>
                {chErr && <div style={{ padding: '4px 8px', color: '#e3b341', fontSize: 11 }}>{chErr}</div>}
                <div style={{ flex: 1, overflow: 'auto', minHeight: 0 }}>
                  {chapters.map((c, k) => (
                    <div
                      key={k}
                      draggable
                      onDragStart={() => { chapDragFrom.current = k }}
                      onDragOver={(e) => { e.preventDefault(); setChapDragOver(k) }}
                      onDragLeave={() => { if (chapDragOver === k) setChapDragOver(null) }}
                      onDrop={(e) => {
                        e.preventDefault()
                        const from = chapDragFrom.current
                        setChapDragOver(null)
                        chapDragFrom.current = null
                        if (from == null || from === k) return
                        const fromI = chapters[from].i
                        const count = (from + 1 < chapters.length ? chapters[from + 1].i : chTotal) - fromI
                        const to = chapters[k].i
                        if (to === fromI || (to > fromI && to < fromI + count)) return
                        sendChapterCmd('moveChapter', { from: fromI, count, to })
                      }}
                      onClick={() => sendChapterCmd('gotoChapter', { i: c.i })}
                      title={c.text}
                      style={{
                        padding: '5px 8px 5px ' + (8 + (c.level - 1) * 13) + 'px', fontSize: 12, lineHeight: 1.5, cursor: 'pointer',
                        borderBottom: '1px solid rgba(255,255,255,.05)',
                        background: chapDragOver === k ? 'rgba(88,166,255,.16)' : 'transparent',
                        boxShadow: chapDragOver === k ? 'inset 0 2px 0 #58a6ff' : 'none',
                        whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', userSelect: 'none',
                      }}
                    >{c.text || '（空标题）'}</div>
                  ))}
                  {!chapters.length && (
                    <div style={{ padding: 12, color: '#66707c', fontSize: 11, lineHeight: 1.8 }}>
                      未发现 Heading 标题。<br />把段落样式设为「标题 1/2/3」后点 ↻。<br />提示：标题挂「多级编号」后，拖拽排序编号自动刷新。
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        )}
        {/* 诊断抽屉：标题栏 ⓘ 切换，默认关闭（diagOpen 由 NodeEditor 传入） */}
        {diagOpen && (
          <div
            style={{
              position: 'absolute', left: 0, right: 0, bottom: 30,
              borderTop: '1px solid #333', background: 'rgba(20,20,20,0.96)', zIndex: 5,
            }}
          >
            <div style={{ display: 'flex', gap: 8, padding: '3px 10px', alignItems: 'center' }}>
              <span style={{ color: '#888', fontSize: 11 }}>{status}</span>
              <span style={{ flex: 1 }} />
              <button
                className="bar-btn"
                style={{ fontSize: 11, padding: '2px 8px' }}
                onClick={() => setDiagCollapsed((v) => !v)}
              >{diagCollapsed ? '展开诊断' : '收起诊断'}</button>
              <button
                className="bar-btn"
                style={{ fontSize: 11, padding: '2px 8px' }}
                onClick={() => { setDiag(''); try { say('诊断已清空') } catch { /* ignore */ } }}
              >清空诊断</button>
              <button
                className="bar-btn"
                style={{ fontSize: 11, padding: '2px 8px' }}
                onClick={() => { try { navigator.clipboard.writeText(diag); say('诊断已复制') } catch { /* ignore */ } }}
              >复制诊断</button>
            </div>
            {!diagCollapsed && (
              <pre
                style={{ margin: 0, padding: '4px 10px', maxHeight: 120, overflow: 'auto', fontSize: 11, lineHeight: 1.5, fontFamily: 'Consolas, monospace', color: '#9cdcaa', userSelect: 'text', cursor: 'text', whiteSpace: 'pre-wrap' }}
              >{diag || '（暂无诊断信息）'}</pre>
            )}
          </div>
        )}
      </div>
    </div>
  )
})

export interface OnlyOfficeNodeEditorProps {
  kind: OOKind
  /** 拖动中：iframe 降级为占位（跨进程合成层每帧跟随=拖动卡根因） */
  anyDragging?: boolean
  /** 节点 id：预加载冻结实例按 nodeId 关联（预加载命中 → 挂回冻结现场） */
  nodeId?: string
  /** OO 数据锁定：true = 编辑器以只读模式打开（mode=view，防误编辑） */
  dataLocked?: boolean
  bytes?: Uint8Array | null
  name: string
  /** 节点绑定的模板库 id（保存回写后存在 → 重开节点自动恢复） */
  ooTemplateId?: string
  /** ★ 节点内置内容（base64）：随画布节点配置持久化——优先级高于模板库（2026-10-11 用户指令「保存在画布节点文件随画布保存」） */
  ooBytes64?: string
  /** ★ 2026-10-29 节点内容 IndexedDB 索引（b64 出 config 性能根治）：加载优先级同 ooBytes64——
   *  拖入导入/文件夹拖入链路只落此字段（ooBytesDbId + 合成 tpl id），boot 必须认它，
   *  否则重开节点掉进模板库 → tpl:read 抛「模板文件不存在」→ 打开失败（今日实测翻车链）。 */
  ooBytesDbId?: string
  /**
   * ★ 2026-10-17 域外引用注入（用户：「改了引用的节点的参数 为什么 oo 节点中的数据没有更新」）：
   *   打开文档前由宿主把 ooRefs 的**当前值**写进文件 ——
   *   补回隐藏表『域数据』+ 定义名称 + 引用格缓存。
   *   ★ 必须在**打开前**做：内容是 xlsx 字节，进来之后再改就要重载文档。
   */
  injectRefs?: (bytes: Uint8Array, opts?: { forOpen?: boolean }) => Promise<{ bytes: Uint8Array; note?: string; changed?: { name: string; was: string; now: string }[] }>
  /** 诊断抽屉开关（标题栏 ⓘ 切换，默认关闭） */
  diagOpen?: boolean
  /** 宿主应用主题（dark/light）：决定 OO 编辑器 uiTheme，切换时重挂 iframe */
  theme?: 'dark' | 'light'
  /** 节点是否最大化（决定编辑器内参数引用按钮显隐，普通模式隐藏） */
  maximized?: boolean
  /** 参数引用面板开合（宿主持有状态；回传 /poc 按钮高亮） */
  refOpen?: boolean
  /** 可引用的画布参数树（宿主 computeNode 计算） */
  ooRefItems?: { id: string; name: string; fields: { id: string; name: string; value: string }[] }[]
  /** 面板内点选参数 → 宿主登记 config.ooRefs */
  onPickRef?: (srcId: string, fieldId: string, fieldName: string, value?: string, mode?: string) => void
  /** 打开的文件名上报（宿主显示在标题栏） */
  onOpenedName?: (name: string) => void
  /** /poc 工具栏按钮点击 → 宿主切参数引用面板 */
  onToggleRef?: () => void
  /** ★ Ctrl+S（2026-10-11）：编辑器内 Ctrl+S → 宿主执行保存回写 */
  /** ★ 2026-10-11 保存链重构：/poc 引擎直读的值数组（宿主据此构建 xlsx 入节点配置） */
  saving?: boolean
  onSaveValues?: (payload: { name: string; rows: unknown[][]; mode?: string }) => void
  /** 「刷新域值」—— 把 ooRefs 的当前值写回 OO 文件（2026-10-06）
   *  ★ 写回必须由**宿主**发起：它要读模板、解析落点、调补丁写回、存库，
   *    而/poc（编辑器内）拿不到这些能力。 */
  onRefreshRef?: () => void
  /** 写回是否在进行中（按钮置灰，避免重复点击） */
  refreshBusy?: boolean
  /** ★ 高亮行列开关（2026-10-18 用户指令）：宿主持有状态（localStorage），这里只转发 /poc */
  rowColHi?: '1' | '0'
  /** ★ 2026-10-18 着色开关改版：引用格覆盖层数据（宿主从 xlsx 提取，纯视觉、不写文件）
   *   marks: 命中引用格 [{sheet, ref, color}]；geom: 列宽/行高几何；on: 当前开关。 */
  refMarks?: { on: '1' | '0'; marks: { sheet: string; ref: string; color: string }[]; geom: unknown }
  /** ★ 2026-10-18 外部关联引用更新策略（不检测 / 检测后提示 / 自动更新），由 /poc 顶部按钮切换 */
  ooRefUpdMode?: 'off' | 'ask' | 'auto'
  /** 待用户确认的更新提问：/poc 弹窗渲染，点「立即更新」回传 __ooRefApply */
  ooRefAsk?: { id: number; count: number; names: string[]; items?: { name: string; was: string; now: string }[]; trigger: 'open' | 'live' } | null
  /** 最近一次检测结果（共 N 处 / M 处待更新），/poc 设置卡显示 */
  ooRefInfo?: { total: number; changed: number; at: number } | null
  onSetRefUpdMode?: (mode: string) => void
  onApplyRefAsk?: () => void
  onStatus?: (s: string) => void
  onSaved: (bytes: Uint8Array) => void
  /** ★ 2026-10-17 未保存修改状态上报（/poc 主动推 __ooDirty）：
   *   宿主据此在标题栏亮「未保存」、高亮 💾、关闭时拦截。 */
  onDirtyChange?: (dirty: boolean) => void
  onClose?: () => void
}

function b64(u8: Uint8Array): string {
  let bin = ''
  const CH = 0x8000
  for (let i = 0; i < u8.length; i += CH) bin += String.fromCharCode(...u8.subarray(i, i + CH))
  return btoa(bin)
}
function b64ToU8(b64s: string): Uint8Array {
  const bin = atob(b64s)
  const out = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
  return out
}

/** ★ 2026-10-17 假成功判定用：字节指纹（FNV-1a 32 位 + 长度）。
 *  用途：主通道产物与「打开时上传的原始字节」完全一致 ⇒ 产物就是原文件，
 *  一个字节都没吃到编辑 ⇒ 必然是假成功，不能拿它覆盖提取通道的正确内容。 */
function byteSig(u8: Uint8Array): string {
  let h = 0x811c9dc5
  for (let i = 0; i < u8.length; i++) { h ^= u8[i]; h = Math.imul(h, 0x01000193) >>> 0 }
  return u8.length + ':' + h.toString(16)
}
