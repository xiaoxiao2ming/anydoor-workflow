import { useCallback, useEffect, useMemo, useRef, useState, type MouseEvent as ReactMouseEvent } from 'react'
import { createPortal } from 'react-dom'
import { Document, Packer, Paragraph } from 'docx'
import JSZip from 'jszip'
import { Ribbon, type RibbonTab } from '../office/Ribbon'
import type { TableSchema, NodeConfig, DocxBinding, FieldDef } from '../core/schema'
import type { Row } from '../core/store'
import { buildUpstream } from '../ref/context'
import { parseTemplate, bind, ensureNumbering, type BindingInfo } from './binder'
import { importDocxTemplate } from './importTemplate'
import { buildData } from './buildData'
import { getTemplateStore, pickTemplateFile, saveDocAs, savePdfAs, type TemplateMeta } from './templateStore'
import { DocxPreview } from './DocxPreview'
import { RichEditor } from './RichEditor'
import { TiptapDocEditor, cleanHtml, type Editor } from './TiptapDocEditor'
import { docxToHtml } from './toHtml'
import { htmlToOoxml } from './html2ooxml'
import { docxBodyToHtmlFromZip } from './ooxml2html'
import { createImageSink } from './image'
import { loadEditedDocx, saveEditedDocx, clearEditedDocx } from './editStore'
import { OnlyOfficeModal } from '../office/OnlyOfficeModal'
import { RefPanel, RightPanel } from '../xlsx/RefPanel'
import type { ResolveTrace } from '../ref/paramTable'

/**
 * 溯源里「来源类型」的中文说法。
 * 与 Excel 侧 xrp-trace 刻意用同一套措辞 —— 同一个值在两种产物里
 * 不能有两种叫法，否则用户会以为是不同的东西。
 */
const SOURCE_LABEL: Record<string, string> = {
  override: '手填',
  default: '计算默认',
  node: '画布节点',
  file: '外部文件',
  api: '网络接口',
  param: '全局参数',
  const: '常量',
  unresolved: '未解析',
}

/** 溯源里值的短显示：表格给行数，对象给 JSON 截断，其余取字符串 */
function fmtTraceValue(v: unknown): string {
  if (v === undefined) return '—'
  if (v === null) return '（空）'
  if (Array.isArray(v)) return `${v.length} 行`
  if (typeof v === 'object') return JSON.stringify(v).slice(0, 40)
  const s = String(v)
  return s.length > 40 ? s.slice(0, 40) + '…' : s
}

/**
 * output 节点的「Word 报告」区块。
 *
 * 四条需求的落点：
 *   导入现有模板文档 → 导入 .docx，自动 sdtify 成可绑定模板
 *   文档内数据动态连接 → 绑定清单 + 字段映射，数据变了产物跟着变
 *   富文本/表格/段落   → rich / each / if / field 四类控件
 *   所见即所得        → 同一份字节既进预览也进导出
 */

const KIND_LABEL: Record<BindingInfo['kind'], string> = {
  field: '文本',
  rich: '富文本',
  each: '表格循环',
  if: '条件段',
  image: '图片',
}
const KIND_CLS: Record<BindingInfo['kind'], string> = {
  field: 'bk-field',
  rich: 'bk-rich',
  each: 'bk-each',
  if: 'bk-if',
  image: 'bk-image',
}
/** 手工插图进页面 JSON，超过这个体积要拦一下（base64 会膨胀 33%） */
const MAX_IMAGE_BYTES = 2 * 1024 * 1024
const AGGS: { v: NonNullable<DocxBinding['agg']>; t: string }[] = [
  { v: 'first', t: '首行' },
  { v: 'sum', t: '合计' },
  { v: 'count', t: '计数' },
  { v: 'max', t: '最大' },
  { v: 'min', t: '最小' },
  { v: 'avg', t: '平均' },
]

export function WordReportPanel({
  table,
  tables,
  records,
  onSetNodeConfig,
  compact,
  editorMode,
  maximized,
  onToggleMax,
  onTemplatesChanged,
}: {
  table: TableSchema
  tables: Record<string, TableSchema>
  records: Record<string, Row[]>
  onSetNodeConfig: (tableId: string, patch: Partial<NodeConfig>) => void
  /** 属性面板紧凑态：功能区默认收起、无最大化 */
  compact?: boolean
  /** 放大编辑器态（NodeEditor 双击进入）：启用窗体架构 —— 左右两栏 + 最大化布局联动 + 空态 hero */
  editorMode?: boolean
  /** 最大化（NodeEditor 标题栏 headMax 持有，与 Excel WorkbookEditor 同一状态源） */
  maximized?: boolean
  onToggleMax?: () => void
  /** 模板库变化通知（导入/创建空白后让 NodeEditor 标题栏历史下拉刷新） */
  onTemplatesChanged?: () => void
}) {
  const config = table.config ?? {}
  const templateId = config.docxTemplateId
  const map = config.docxMap ?? {}

  const [tplBytes, setTplBytes] = useState<Uint8Array | null>(null)
  const [bindings, setBindings] = useState<BindingInfo[]>([])
  const [parseWarn, setParseWarn] = useState<string[]>([])
  /** 非阻断提示：模板载入了但有话要说（如没有绑定标记，按原文档处理） */
  const [tplWarn, setTplWarn] = useState('')
  const [preview, setPreview] = useState<Uint8Array | null>(null)
  const [bindWarn, setBindWarn] = useState<string[]>([])
  const [busy, setBusy] = useState('')
  const [err, setErr] = useState('')
  const [libs, setLibs] = useState<TemplateMeta[]>([])
  const [tab, setTab] = useState<'bind' | 'preview'>(editorMode ? 'preview' : 'bind')
  const [zoom, setZoom] = useState(false)

  /* ── WPS 式就地编辑（editorMode 专属）──────────────────────────
   * 编辑对象 = 绑定产物（preview 字节，sdt 已剥离）——与 OnlyOffice「在线编辑」同语义：
   * 手工精修覆盖在数据绑定结果之上，数据重新绑定后编辑稿作废（以数据为准）。
   * 编辑稿按 nodeId 持久化到 IndexedDB（editStore），重开编辑器可恢复；
   * 模板字节永不被编辑稿覆盖 —— 否则 parseTemplate 找不到控件，数据连接就断了。 */
  const [editedBytes, setEditedBytes] = useState<Uint8Array | null>(null)
  const [editRestored, setEditRestored] = useState(false)
  const editPaneRef = useRef<HTMLDivElement>(null)
  /** Tiptap 编辑器实例（命令层入口） */
  const editorRef = useRef<Editor | null>(null)
  /**
   * 编辑器实例（state 版）—— 仅用于「选区格式同步」的 effect 依赖。
   * 不能只用 editorRef：ref 赋值不触发重渲染，父组件的 effect 无法感知实例就绪，
   * 结果 selectionUpdate 永远订阅不上（2026-09-28 实测 S3 hasOn=false 的根因）。
   */
  const [editorInst, setEditorInst] = useState<Editor | null>(null)
  /** 当前编辑基础字节：编辑稿或绑定产物（写回时在它上面换 body） */
  const editBaseRef = useRef<Uint8Array | null>(null)
  editBaseRef.current = editedBytes ?? preview
  /** 编辑面的 OOXML 元信息（写回时重组 document.xml 用） */
  const editMetaRef = useRef<{ sectPr: string; rootAttrs: string }>({ sectPr: '', rootAttrs: '' })
  const editTimerRef = useRef<number | null>(null)
  const statsTimerRef = useRef<number | null>(null)
  const pendingHtmlRef = useRef<string | null>(null)
  /** onEditInput 的 ref 镜像：插图/环绕等定义在它之前的回调经此调用（避开 use-before-declare） */
  const onEditInputRef = useRef<() => void>(() => {})
  /** Ribbon 按钮点击会让编辑面失焦丢选区 —— selectionchange 时把编辑面内的选区存起来 */
  const selRangeRef = useRef<Range | null>(null)
  /**
   * 选区格式快照（WPS「同步选区格式」）：Ribbon 按钮据此显示按下态。
   * 不做成 8 个独立 state —— 一个对象一次更新，避免多次 setState 造成的渲染抖动。
   * 由 Tiptap 的 selectionUpdate 驱动（见下方 effect）。
   */
  const [selFmt, setSelFmt] = useState({
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    subscript: false,
    superscript: false,
    bulletList: false,
    orderedList: false,
    align: 'left' as 'left' | 'center' | 'right' | 'justify',
    block: 'p' as 'p' | 'h1' | 'h2' | 'h3',
  })
  /** flush 落库时读当前模板 id（稿带 tplId —— 换模板后旧稿自动失效）；经 ref 读避免进 deps */
  const tplIdRef = useRef(templateId ?? '')
  tplIdRef.current = templateId ?? ''

  /* ── OnlyOffice 在线编辑 ── */
  const [editing, setEditing] = useState<string | null>(null)
  /** 插图环节的提示按 tag 存 —— 一个控件选错图不该污染整个面板 */
  const [imgWarn, setImgWarn] = useState<Record<string, string>>({})
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null)

  /* ── WPS 式状态栏：纸张缩放 + 页数/字数统计 ── */
  const [zoomPct, setZoomPct] = useState(100)
  const [docStats, setDocStats] = useState({ pages: 1, words: 0 })

  /* ── 页面交互模型（单一真相源，WPS 借鉴）：
   *   标尺 / 缩略图导航 / 缩放 / 滚动跟随 / 视图模式 全部读它、双向写回它。
   *   pages：每页 {top,height}（未缩放 CSS px，由 period 派生）；
   *   sectMargin：页边距（px），标尺可拖改 → 真写 OOXML sectPr；
   *   tabStops：制表位（相对左页边距的 px），标尺点击添加；
   *   viewMode：页面 / Web版式 / 草稿。 */
  const [viewMode, setViewMode] = useState<'page' | 'web' | 'draft'>('page')
  const [sectMargin, setSectMargin] = useState({ l: 120, r: 120, t: 96, b: 96 })
  const [pageModel, setPageModel] = useState<{ pages: { top: number; height: number }[] }>({ pages: [] })
  /** 制表位（相对左页边距的 px，标尺点击添加）；首行/悬挂缩进（px，来自当前段） */
  const [tabStops, setTabStops] = useState<number[]>([])
  const [indentState, setIndentState] = useState({ first: 0, hang: 0 })
  /** 选区浮工具条 / 右键菜单位置（视口坐标）；null = 隐藏 */
  const [floatBar, setFloatBar] = useState<{ x: number; y: number } | null>(null)
  const [ctxMenu, setCtxMenu] = useState<{ x: number; y: number } | null>(null)

  /* ── 上游数据：走引擎，不从 UI 视口抓 ──────────────────────
     统一走 src/ref/context.ts 的 buildUpstream —— 与 Excel 侧同一套实现。
     原先这里只读 config.inputs、漏了连线（config.linkedInputs），
     结果是「拖线连上的上游」在 Word 面板里看不到数据、Excel 面板里却看得到。*/
  const ub = useMemo(() => buildUpstream(table, tables, records), [table, tables, records])
  const upstream = useMemo(() => {
    // 单上游时直接用它的行；多上游时拼接 —— 与 buildDataFromNode 的语义一致
    const rows = ub.computed.length === 1 ? ub.computed[0].rows : ub.computed.flatMap((c) => c.rows)
    return { fields: ub.fields, rows }
  }, [ub])

  /* ── 载入模板库与当前模板 ────────────────────────────────── */
  const refreshLibs = useCallback(async () => {
    try {
      setLibs(await getTemplateStore().list())
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e))
    }
  }, [])

  useEffect(() => {
    void refreshLibs()
  }, [refreshLibs])

  useEffect(() => {
    let alive = true
    // 生命周期门闩重置：模板 parse + 恢复检查完成前，bind effect 不判 changed（防初始化噪声误杀恢复稿）
    editBootRef.current = false
    editFpRef.current = null
    setEditedBytes(null)
    setEditRestored(false)
    if (!templateId) {
      setTplBytes(null)
      setBindings([])
      return
    }
    setBusy('正在载入模板…')
    // 模板字节与历史编辑稿并行载入；稿带 tplId —— 换过模板的旧稿在此自动失效
    Promise.all([getTemplateStore().read(templateId), loadEditedDocx(table.id)])
      .then(async ([bytes, saved]) => {
        if (!alive) return
        if (!bytes) {
          setErr('模板文件已丢失，请重新导入')
          return
        }
        setTplBytes(bytes)
        // 真实页边距：设计院文档左右常不对称（如李渡 92.6/56.7px），不读就按 120/120 假边距
        // 排版 —— 版心 554px 装不下 643px 的表格（tblW），列宽/行高/页数全部失真
        try {
          const z2 = await JSZip.loadAsync(bytes)
          const docXml = (await z2.file('word/document.xml')?.async('string')) ?? ''
          const sect = /<w:sectPr[\s\S]*?<\/w:sectPr>/.exec(docXml)?.[0]
          if (sect) {
            const px = (k: string) => {
              const m = new RegExp(`<w:pgMar[^>]*w:${k}="(-?[\\d.]+)"`).exec(sect)
              return m ? Math.round((Number(m[1]) / 15) * 10) / 10 : null // twips→px @96dpi
            }
            const t = px('top')
            const b2 = px('bottom')
            const l = px('left')
            const r = px('right')
            if (l != null || r != null || t != null || b2 != null) {
              setSectMargin({ t: t ?? 96, b: b2 ?? 96, l: l ?? 120, r: r ?? 120 })
            }
          }
        } catch {
          /* 边距读取失败 → 沿用默认 120/96 */
        }
        const r = await parseTemplate(bytes)
        if (!alive) return
        setBindings(r.bindings)
        setParseWarn(r.warnings)
        if (saved && saved.tplId === templateId) {
          // 就地编辑稿恢复（仅编辑器态）：稿的基准指纹由下一轮 bind effect 记录（此时 built 已无初始化噪声）
          if (editorMode) {
            setEditedBytes(saved.bytes)
            setEditRestored(true)
            editBootRef.current = true
          }
        } else if (saved) {
          void clearEditedDocx(table.id) // 别的模板的旧稿：清掉死数据
        }
      })
      .catch((e: unknown) => alive && setErr(e instanceof Error ? e.message : String(e)))
      .finally(() => alive && setBusy(''))
    return () => {
      alive = false
    }
  }, [editorMode, templateId, table.id])

  /* ── 数据 → 绑定值 ───────────────────────────────────────── */
  const built = useMemo(
    () => buildData(bindings, map, upstream.fields, upstream.rows, { ub }),
    [bindings, map, upstream, ub],
  )

  /* ── 重绑（debounce 250ms）→ 预览 ────────────────────────── */
  /** 上次 bind 的输入指纹：changed = 输入真的变了 → 作废编辑稿（以数据为准，与 OnlyOffice 同语义） */
  const lastBindRef = useRef<{ tpl: unknown; built: unknown } | null>(null)
  /** 恢复稿生命周期：editBootRef = 恢复稿已到达、等下一轮 bind 记基准；editFpRef = 恢复稿的基准 built 指纹 */
  const editBootRef = useRef(false)
  const editFpRef = useRef<unknown>(null)
  const editedRef = useRef<Uint8Array | null>(null)
  editedRef.current = editedBytes
  const editRestoredRef = useRef(false)
  editRestoredRef.current = editRestored
  useEffect(() => {
    if (timer.current) clearTimeout(timer.current)
    // 0 绑定也预览：普通文档（无标记）走 bind 原样输出 —— 「打开就能看」与 Excel 节点一致
    if (!tplBytes) {
      setPreview(null)
      return
    }
    if (editBootRef.current) {
      // 恢复稿到达后的第一轮：本轮 built 即稿的基准（模板 parse 已随恢复一并完成，无初始化噪声），只记不判
      editBootRef.current = false
      editFpRef.current = built
      lastBindRef.current = { tpl: tplBytes, built }
    } else if (editFpRef.current != null) {
      // 恢复稿存活期：数据/模板相对稿的基准变了 → 稿作废
      const changed = editFpRef.current !== built || lastBindRef.current?.tpl !== tplBytes
      lastBindRef.current = { tpl: tplBytes, built }
      if (changed) {
        setEditedBytes(null)
        setEditRestored(false)
        editFpRef.current = null
        void clearEditedDocx(table.id)
      }
    } else {
      // 无稿期：普通指纹判定（首次挂载 lastBindRef 为 null → 只记不判）
      const changed = lastBindRef.current != null && (lastBindRef.current.tpl !== tplBytes || lastBindRef.current.built !== built)
      lastBindRef.current = { tpl: tplBytes, built }
      if (changed && (editedRef.current || editRestoredRef.current)) {
        setEditedBytes(null)
        setEditRestored(false)
        void clearEditedDocx(table.id)
      }
    }
    timer.current = setTimeout(() => {
      // 预览与导出用同一份参数：stripControls 剥掉控件外壳，预览即终稿
      bind(tplBytes, built.data, { stripControls: true })
        .then((r) => {
          setPreview(r.buffer)
          setBindWarn(r.warnings)
        })
        .catch((e: unknown) => setErr(e instanceof Error ? e.message : String(e)))
    }, 250)
    return () => {
      if (timer.current) clearTimeout(timer.current)
    }
  }, [tplBytes, bindings, built, table.id])

  /* ── 导入模板 ────────────────────────────────────────────── */
  const importTemplate = useCallback(async () => {
    setErr('')
    const picked = await pickTemplateFile()
    if (!picked) return
    try {
      setBusy('正在分析模板…')
      setTplWarn('')
      // 设计院给的是普通文档 → 把 ｛｛field:x｝｝ 标记原地转成 Word 内容控件（共享逻辑，与标题栏 📂 同一份）
      const { meta, noMarks, tags } = await importDocxTemplate(picked)
      if (noMarks) {
        // 没有绑定标记 ≠ 不能用：与 Excel 节点同一哲学 —— 原样载入，可预览可导出；
        // 要连数据时在 Word 里加 ｛｛field:字段名｝｝ 等标记再重新导入即可。
        onSetNodeConfig(table.id, { docxTemplateId: meta.id, docxTemplateName: meta.name })
        await refreshLibs()
        onTemplatesChanged?.()
        setTplWarn(
          `未找到绑定标记——已按原文档载入（可预览/导出）。要连数据：加 ｛｛field:字段名｝｝ 后重新导入。`,
        )
        setBusy('')
        return
      }
      onSetNodeConfig(table.id, { docxTemplateId: meta.id, docxTemplateName: meta.name })
      await refreshLibs()
      onTemplatesChanged?.()
      setBusy('')
      // eslint-disable-next-line no-console
      console.log('[docx] 模板已控件化', meta.name, tags)
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e))
      setBusy('')
    }
  }, [onSetNodeConfig, table.id, refreshLibs, onTemplatesChanged])

  /* ── 导出（就地编辑稿优先：用户精修过的内容就是终稿）────────── */
  const exportDoc = useCallback(async () => {
    if (!tplBytes) return
    try {
      setBusy('正在生成…')
      const bytes =
        editedBytes ??
        (await bind(tplBytes, built.data, { stripControls: config.docxStripControls !== false })).buffer
      const name = `${table.name || '报告'}-${new Date().toISOString().slice(0, 10)}.docx`
      await saveDocAs(name, bytes)
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy('')
    }
  }, [tplBytes, editedBytes, built, config.docxStripControls, table.name])

  /* ── 导出 PDF：直接吃预览那份字节，保证「看到的就是拿到的」 ── */
  const exportPdf = useCallback(async () => {
    if (!tplBytes) return
    try {
      setBusy('正在生成 PDF…')
      const bytes = editedBytes ?? preview ?? (await bind(tplBytes, built.data, { stripControls: true })).buffer
      const html = await docxToHtml(bytes)
      const name = `${table.name || '报告'}-${new Date().toISOString().slice(0, 10)}.pdf`
      await savePdfAs(name, html)
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy('')
    }
  }, [tplBytes, editedBytes, preview, built, table.name])

  /* ── OnlyOffice 在线编辑 ──────────────────────────────────
   * 流程：绑定产物上传 app-server 工作目录 → DS 拉取编辑 → 保存回调写回
   * → 前端重新下载 setPreview。手工精修的内容在预览里保留展示；
   * 注意：之后若画布数据再变，绑定重建会以数据为准覆盖编辑稿。 */
  const openInEditor = useCallback(async () => {
    if (!tplBytes) return
    try {
      setBusy('正在上传到编辑工作区…')
      const bytes = preview ?? (await bind(tplBytes, built.data, { stripControls: true })).buffer
      let bin = ''
      const CH = 0x8000
      for (let i = 0; i < bytes.length; i += CH)
        bin += String.fromCharCode(...bytes.subarray(i, i + CH))
      const name = `${table.name || '报告'}.docx`
      const r = await fetch('/office/upload', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, dataBase64: btoa(bin) }),
      })
      if (!r.ok) throw new Error(`上传失败（${r.status}）——app-server 是否在运行？`)
      setEditing(name)
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy('')
    }
  }, [tplBytes, preview, built, table.name])

  /** DS 保存回调完成（文件已写回）→ 重新下载刷新预览，所见即所得 */
  const onOfficeSaved = useCallback(async () => {
    if (!editing) return
    try {
      const r = await fetch(`/office/files/${encodeURIComponent(editing)}`)
      if (!r.ok) return
      setPreview(new Uint8Array(await r.arrayBuffer()))
      setTab('preview')
    } catch {
      /* 保存回写已由服务端确认，刷新失败静默即可 */
    }
  }, [editing])

  const setBinding = (tag: string, patch: Partial<DocxBinding>) => {
    const cur = map[tag] ?? { from: 'manual' as const }
    onSetNodeConfig(table.id, { docxMap: { ...map, [tag]: { ...cur, ...patch } } })
  }

  /**
   * 选一张本地图 → data: URI 存进绑定值。
   *
   * 为什么必须转成 data: URI：绑定值要随节点配置一起存进工程 JSON，
   * 本地绝对路径换台机器就失效。base64 膨胀 33%，所以按解码后字节数设上限。
   * 引擎侧只认 data: URI，外链一律降级为可见文字。
   */
  const pickImage = async (tag: string, file: File) => {
    setImgWarn((m) => ({ ...m, [tag]: '' }))
    if (file.size > MAX_IMAGE_BYTES) {
      setImgWarn((m) => ({
        ...m,
        [tag]: `图片 ${(file.size / 1024 / 1024).toFixed(1)}MB 超过 ${MAX_IMAGE_BYTES / 1024 / 1024}MB 上限，请压缩后再插`,
      }))
      return
    }
    try {
      const buf = await file.arrayBuffer()
      const bytes = new Uint8Array(buf)
      let bin = ''
      for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i])
      const uri = `data:${file.type || 'image/png'};base64,${btoa(bin)}`
      setBinding(tag, { from: 'manual', value: uri })
    } catch (e) {
      setImgWarn((m) => ({ ...m, [tag]: `读取图片失败：${e instanceof Error ? e.message : e}` }))
    }
  }

  /** ── 最大化：状态由 NodeEditor 标题栏持有（props.max），面板内不再自持 ── */

  /* ── 窗体架构（对齐 Excel WorkbookEditor）：左栏=节点数据引用 · 右栏=数据预选追加 ──
     交互完全镜像 Excel 侧：开关持久化 / 一退一进动画 / 几何测量（常规=窗体外卡片、最大化=嵌入窗体缘）
     / 两栏全关自动最大化。localStorage key 加 wrp- 前缀，避免与 Excel 编辑器互相串台。 */
  const [activeTag, setActiveTag] = useState<string>('')
  /** 左栏点参数但还没有选中域时的轻提示（不清 except 下一次成功） */
  const [pickHint, setPickHint] = useState('')
  const [refPanelOpen, setRefPanelOpen] = useState(() => {
    if (!editorMode) return false
    try {
      const saved = localStorage.getItem(`wrp-refpanel:${table.id}`)
      if (saved !== null) return saved === '1'
    } catch {
      /* 读不了就算了 */
    }
    return true
  })
  const refPanelOpenRef = useRef(refPanelOpen)
  refPanelOpenRef.current = refPanelOpen
  useEffect(() => {
    // 串台守卫：compact 实例（属性面板）不得写盘 —— 它先挂载，会把 "0" 写进与
    // 编辑器实例共享的 key，编辑器打开时读到 "0" 就不渲染左栏（2026-09-27 实测 bug）
    if (!editorMode) return
    try {
      localStorage.setItem(`wrp-refpanel:${table.id}`, refPanelOpen ? '1' : '0')
    } catch {
      /* 存不了就算了 */
    }
  }, [refPanelOpen, table.id, editorMode])

  const [rightOpen, setRightOpen] = useState(() => {
    if (!editorMode) return false
    try {
      return localStorage.getItem(`wrp-rightpanel:${table.id}`) === '1'
    } catch {
      return false
    }
  })
  const rightOpenRef = useRef(rightOpen)
  rightOpenRef.current = rightOpen
  useEffect(() => {
    // 同上：compact 实例不写盘，防止覆盖编辑器实例的开关状态
    if (!editorMode) return
    try {
      localStorage.setItem(`wrp-rightpanel:${table.id}`, rightOpen ? '1' : '0')
    } catch {
      /* 存不了就算了 */
    }
  }, [rightOpen, table.id, editorMode])

  /** 一退一进动画（普通视图）：关闭播退场动画再卸载，打开播入场动画；最大化直接切换 —— 同 Excel 左栏 */
  const [panelMounted, setPanelMounted] = useState(refPanelOpen)
  const [panelLeaving, setPanelLeaving] = useState(false)
  const panelLeaveTimerRef = useRef<number | null>(null)
  useEffect(() => {
    if (panelLeaveTimerRef.current != null) {
      window.clearTimeout(panelLeaveTimerRef.current)
      panelLeaveTimerRef.current = null
    }
    if (refPanelOpen) {
      setPanelMounted(true)
      setPanelLeaving(false)
    } else if (maximized) {
      setPanelMounted(false)
      setPanelLeaving(false)
    } else if (panelMounted) {
      setPanelLeaving(true)
      panelLeaveTimerRef.current = window.setTimeout(() => {
        panelLeaveTimerRef.current = null
        setPanelMounted(false)
        setPanelLeaving(false)
      }, 240)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refPanelOpen, maximized])
  const [rightMounted, setRightMounted] = useState(rightOpen)
  const [rightLeaving, setRightLeaving] = useState(false)
  const rightLeaveTimerRef = useRef<number | null>(null)
  useEffect(() => {
    if (rightLeaveTimerRef.current != null) {
      window.clearTimeout(rightLeaveTimerRef.current)
      rightLeaveTimerRef.current = null
    }
    if (rightOpen) {
      setRightMounted(true)
      setRightLeaving(false)
    } else if (maximized) {
      setRightMounted(false)
      setRightLeaving(false)
    } else if (rightMounted) {
      setRightLeaving(true)
      rightLeaveTimerRef.current = window.setTimeout(() => {
        rightLeaveTimerRef.current = null
        setRightMounted(false)
        setRightLeaving(false)
      }, 240)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rightOpen, maximized])
  useEffect(
    () => () => {
      if (panelLeaveTimerRef.current != null) window.clearTimeout(panelLeaveTimerRef.current)
      if (rightLeaveTimerRef.current != null) window.clearTimeout(rightLeaveTimerRef.current)
    },
    [],
  )
  /** 两栏都退出 → 自动进入最大化（下降沿触发一次；还原后不再自动循环）—— 同 Excel */
  const bothClosedRef = useRef(!refPanelOpen && !rightOpen)
  useEffect(() => {
    const both = !refPanelOpen && !rightOpen
    if (both && !bothClosedRef.current && !maximized && onToggleMax) onToggleMax()
    bothClosedRef.current = both
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refPanelOpen, rightOpen, maximized])

  /** 面板几何（viewport 坐标）：常规视图=窗体外圆角卡片 / 最大化=嵌入窗体左右缘（top 从 Ribbon 底开始） */
  const [panelPos, setPanelPos] = useState<{ left: number; top: number; height: number; radius: number; border: boolean } | null>(null)
  const [rightPos, setRightPos] = useState<{ left: number; top: number; height: number; radius: number; border: boolean } | null>(null)
  /** 退场动画期间 pos 会被 measure 清空 —— 渲染用最后一次位置，避免面板跳位 */
  const panelPosLastRef = useRef<{ left: number; top: number; height: number; radius: number; border: boolean } | null>(null)
  if (panelPos) panelPosLastRef.current = panelPos
  const posForRender = panelPos ?? panelPosLastRef.current
  const rightPosLastRef = useRef<{ left: number; top: number; height: number; radius: number; border: boolean } | null>(null)
  if (rightPos) rightPosLastRef.current = rightPos
  const rightPosForRender = rightPos ?? rightPosLastRef.current
  useEffect(() => {
    if (!editorMode) return
    const measure = () => {
      const ne = document.querySelector('.node-editor') as HTMLElement | null
      if (!ne) {
        setPanelPos(null)
        setRightPos(null)
        return
      }
      const nr = ne.getBoundingClientRect()
      const isMax = maximized || ne.classList.contains('ne-max')
      // ★ .rb 必须在 node-editor 内找 —— 页面上同时存在属性面板 compact 实例的 Ribbon
      //   （DOM 顺序在前），全局 querySelector 会命中它，最大化时 top 全错（2026-09-27 实测）
      const rbBottom = () => {
        const rb = ne.querySelector('.rb')
        return rb ? rb.getBoundingClientRect().bottom : nr.top
      }
      // 右栏几何（镜像左栏）：最大化=嵌入右缘；普通=窗体外右侧（屏宽不够则贴屏右）
      if (isMax) {
        setRightPos({ left: nr.right - 230, top: rbBottom(), height: nr.bottom - rbBottom(), radius: 0, border: false })
      } else if (!rightOpenRef.current) {
        setRightPos(null)
      } else {
        setRightPos({
          left: Math.min(nr.right + 12, window.innerWidth - 230 - 8),
          top: nr.top,
          height: nr.height,
          radius: 10,
          border: true,
        })
      }
      if (isMax) {
        // 最大化：left=窗体左缘；top=本编辑器功能区（.rb）底部 —— 引用栏不占 Ribbon 那一行，
        // 高度收到 foot 上缘，底栏按钮不被盖住（可继续点击切换/开关）
        ne.classList.remove('ne-refroom')
        const foot = ne.querySelector('.wrp-foot') as HTMLElement | null
        const footTop = foot ? foot.getBoundingClientRect().top : nr.bottom
        setPanelPos({ left: nr.left, top: rbBottom(), height: Math.max(120, footTop - rbBottom()), radius: 0, border: false })
      } else if (!refPanelOpenRef.current) {
        ne.classList.remove('ne-refroom')
        setPanelPos(null)
      } else {
        // 常规：窗体外左侧，边缘与窗体一致（圆角 10px、1px 边框）
        // 判定基于 mask 内容宽的「未让位」几何 —— 不能用当前 rect（加了让位类后 rect 变化会导致条件翻转振荡）
        const w = 230
        const maskR = document.querySelector('.node-editor-mask')?.getBoundingClientRect()
        const pad = 28 // .node-editor-mask padding
        const contentW = (maskR ? maskR.width : window.innerWidth) - pad * 2
        const naturalW = Math.min(1180, contentW)
        const unshiftedLeft = (maskR ? maskR.left + pad : pad) + (contentW - naturalW) / 2
        if (unshiftedLeft - w - 12 < 8) {
          // 空间不够：引用栏保持贴屏左（不叠在 Word 上），窗体右移让位、宽度收缩
          ne.classList.add('ne-refroom')
          setPanelPos({ left: 8, top: nr.top, height: nr.height, radius: 10, border: true })
        } else {
          ne.classList.remove('ne-refroom')
          setPanelPos({ left: Math.max(8, unshiftedLeft - w - 12), top: nr.top, height: nr.height, radius: 10, border: true })
        }
      }
    }
    measure()
    // Ribbon 收起/展开、窗体尺寸变化都会改变参照 —— ResizeObserver 兜底
    const ro = new ResizeObserver(() => measure())
    const ne2 = document.querySelector('.node-editor')
    const rb2 = ne2?.querySelector('.rb')
    if (ne2) ro.observe(ne2)
    if (rb2) ro.observe(rb2)
    window.addEventListener('resize', measure)
    const t = setTimeout(measure, 60)
    return () => {
      ro.disconnect()
      window.removeEventListener('resize', measure)
      clearTimeout(t)
    }
  }, [editorMode, maximized, refPanelOpen, rightOpen])

  /** 左栏点参数 → 给当前选中域设置取值字段（from:'field'），来源节点并入数据流（画布自动连线） */
  const pickFieldForActive = (srcId: string, _srcName: string, fieldId: string, _fieldName: string) => {
    if (!activeTag) {
      setPickHint('先在「绑定」清单里点选一个域，再到这里点参数')
      return
    }
    const b = bindings.find((x) => x.tag === activeTag)
    if (!b) return
    if (b.kind !== 'field' && b.kind !== 'if') {
      setPickHint(`「${b.path}」是${KIND_LABEL[b.kind]}域，取值方式请在清单里配置`)
      return
    }
    setPickHint('')
    const cur = map[activeTag] ?? { from: 'manual' as const }
    const inputs = config.inputs ?? []
    const nextInputs = inputs.includes(srcId) ? inputs : [...inputs, srcId]
    onSetNodeConfig(table.id, {
      docxMap: { ...map, [activeTag]: { ...cur, from: 'field', fieldId } },
      inputs: nextInputs,
    })
  }

  /** 空态「创建空白文档」：docx 库造一个最小空白文档入库（与 Excel 的 createBlankXlsx 对称） */
  const createBlankDocx = async () => {
    try {
      setBusy('创建空白文档…')
      const doc = new Document({ sections: [{ children: [new Paragraph('')] }] })
      const blob = await Packer.toBlob(doc)
      const data = new Uint8Array(await blob.arrayBuffer())
      const meta = await getTemplateStore().save('未命名.docx', data)
      await refreshLibs()
      onSetNodeConfig(table.id, { docxTemplateId: meta.id, docxTemplateName: meta.name })
      onTemplatesChanged?.()
      setBusy('')
    } catch (e) {
      setErr('创建失败：' + (e instanceof Error ? e.message : String(e)))
      setBusy('')
    }
  }

  /** ── WPS 式就地编辑：装载 / 写回 / 格式命令 ────────────────── */

  /** 编辑基础字节（editedBytes ?? preview）→ HTML 交给 Tiptap 渲染。
   *  Tiptap 内部维护文档模型，只需把解析出的 HTML 作为 value 传入；
   *  重设时机由 TiptapDocEditor 内部把控（聚焦时不打断输入）。 */
  const [editorHtml, setEditorHtml] = useState('')
  useEffect(() => {
    if (!editorMode || tab !== 'preview') return
    const base = editBaseRef.current
    if (!base) return
    let alive = true
    void (async () => {
      try {
        const zip = await JSZip.loadAsync(base)
        const meta = await docxBodyToHtmlFromZip(zip)
        if (!alive) return
        editMetaRef.current = { sectPr: meta.sectPr, rootAttrs: meta.rootAttrs }
        setEditorHtml(meta.html || '<p></p>')
        window.setTimeout(computeStats, 80)
      } catch (e) {
        if (alive) setErr('文档解析失败：' + (e instanceof Error ? e.message : String(e)))
      }
    })()
    return () => {
      alive = false
    }
  }, [editorMode, tab, preview, editedBytes, templateId])

  /** 状态栏统计：页数按 A4 全高 1123px 估算（zoom 会改变 layout，需除回），字数取可见文本（去空白） */
  /* ── 分页常量（computeStats / 纸张层 / 滚动翻页共用，必须同源）──
     页面推进节奏 = 页高 + 页隙；滚动越过分页线即翻页（与 WPS 滚动翻页手感一致） */
  const PAGE_H = 1123
  const PAGE_GAP = 24

  /**
   * 标尺刻度常量（与纸张几何同源，别写魔法数）。
   *
   * A4 宽 21cm，@96dpi → 794px（21 × 37.7953 = 793.7，取整 794）。
   * 1cm = 37.7953px；左右页边距各 120px ≈ 3.175cm ≈ 3.18cm（Word 默认「普通」是 3.18cm）。
   * 可排版文本带宽 = 794 − 240 = 554px = 14.66cm。
   *
   * ★ 旧版错的不是「常量」，而是把 37.8 用在背景刻度、75.6 用在数字，
   *   又额外加 transform:translateX(2px) + border 2px → 数字实测偏 5px，
   *   且数字压在两条刻度线**之间**而不是线上。现在刻度与数字都由同一常量推导。
   */
  const RULER_PX_PER_CM = 37.7953
  /** 标尺刻度区覆盖的厘米数（取整到偶数，只标可排版文本带）*/
  const RULER_CM = 14

  /** 状态栏统计：页数按 A4 全高 1123px 估算（zoom 会改变 layout，需除回），字数取可见文本（去空白） */
  const computeStats = useCallback(() => {
    const host = editPaneRef.current
    if (!host) return
    const k = zoomPct / 100
    const period = PAGE_H + PAGE_GAP
    const pages = Math.max(1, Math.ceil(host.scrollHeight / k / 1123))
    const words = (host.textContent ?? '').replace(/\s+/g, '').length
    setDocStats((s) => (s.pages === pages && s.words === words ? s : { pages, words }))
    // 页面模型：每页 {top,height}（未缩放 CSS px，由 period 派生），供标尺/缩略图/导航统一消费
    setPageModel((pm) => {
      const next = Array.from({ length: pages }, (_, i) => ({ top: i * period, height: PAGE_H }))
      if (pm.pages.length === next.length && next.every((p, i) => pm.pages[i]?.top === p.top)) return pm
      return { pages: next }
    })
    // 内容变了 → 纸张层页数也要跟着变（否则长文档在第 1 页之后就没有背景纸）
    refreshContentPagesRef.current()
  }, [zoomPct])

  const [curPage, setCurPage] = useState(1)
  const editWrapRef = useRef<HTMLDivElement>(null)

  /**
   * 纸张最小高度：至少铺满可视区，否则内容少时纸张下方是一片死白
   * （用户 2026-09-28「纸张下方整个区域是空的」）。
   * 取「视口可用高」与「一张 A4」的较大值，再向上取整到「页高+页隙」的整数倍 ——
   * 保证纸张下沿永远落在页隙上，视觉上不会出现「半张纸」。
   * zoom 会把纸张视觉缩小，故这里要除回（布局高度是未缩放的）。
   */
  /**
   * 纸张页数 = 内容页数与视口页数的较大值。
   * 关键：内容高度必须由「内容变化」驱动重算 —— 只靠 ResizeObserver 不够
   *（.wrp-edit 的 rect 被 minHeight 撑住，ResizeObserver 观察不到内部内容变长）。
   * 所以在 computeStats（每次 flushEdit 后跑）里同步刷新 contentPages。
   */
  const [contentPages, setContentPages] = useState(1)
  const [visPages, setVisPages] = useState(1)
  const period = PAGE_H + PAGE_GAP
  const paperPages = Math.max(1, contentPages, visPages)
  const paperMinH = paperPages * period - PAGE_GAP // 末页不画页隙

  /**
   * 表格行跨页避让（Word/WPS 行为：一行不被页边界拆开）。
   *
   * 为什么需要 JS：分页是「视觉层」—— 页隙由 .wrp-pages 的背景纸提供，
   * 不是浏览器真实分页，所以 CSS `break-inside: avoid` / `page-break-inside` 完全不生效。
   * 做法：量出每个表格行的垂直范围，若与某个页隙区间（[N*period-GAP, N*period)）相交，
   * 就给该行一个 margin-top，把整行推到页隙之后。
   *
   * ★ 三个必须守住的坑（都是实测踩出来的）：
   *
   * 1. 坐标系必须用 offset 链，不能用 getBoundingClientRect。
   *    .wrp-edit 挂了 zoom，getBoundingClientRect 返回缩放后的视觉值，
   *    而 period(1147) 是未缩放的 CSS 值 —— 混用会让判断整体失效。
   *
   * 2. ★ 行的绝对偏移**只能**用「tr → edit 的完整 offsetTop 链」，不能做任何加减。
   *    实测这个 DOM 里：
   *      - tr.offsetTop  = 0, 42, 84 …（相对 **table**，因为 table 是 offsetParent 的实际锚点）
   *      - tr.offsetParent = **null**（表行的 offsetParent 为空，offsetTop 直接相对表格）
   *      - table.offsetTop = 1602（相对 .tiptap 内容盒）
   *    两个 offsetTop 参照系完全不同，做 `tr.offsetTop - table.offsetTop` 得到负数。
   *    而「tr → edit」逐级累加得到 1717, 1759 … 已经是正确的绝对偏移
   *    （TR:0 + TABLE:1602 + TIPTAP:115 = 1717）。
   *    前几版算法的三种错误都在这里：
   *      a) 只累加到 table 就停 → 得到 0,42,84，全部落第 0/1 页，永远判不到页隙；
   *      b) 完整链 + tblTop 相加 → 1717 被算成 3434，判到第 3 页，推错位置；
   *      c) 用 getBoundingClientRect → zoom 缩放后与 period 不同量纲，整体失效。
   *    正解：完整链，**不加** tblTop。
   *
   * 3. ★ 推挤量不能写成行内 style，也不能加 class。
   *    tr 所在的 DOM 子树由 ProseMirror 拥有：任何 `setContent`
   *    （TiptapDocEditor 的 value 同步 effect 会调）或事务重绘都会从文档模型
   *    重建这部分 DOM，把行内 style / 追加的 class **全部抹掉**。
   *    实测：`__wrpAvoid()` 调用后立刻查到 2 行被推，但同一次 React 提交
   *    触发的 setContent 之后，推挤行数又变回 0 —— 白做。
   *    正解：把「第 N 行 → 推多少 px」的映射存进 React state，
   *    用一条注入 <style> 的 CSS 规则表达（选择器用 :nth-child），
   *    规则挂在 .wrp-edit 上，ProseMirror 重建 tr 也影响不到它。
   *
   * 推挤量额外加 PUSH_EPS：只推到 gapEnd 会让行首恰好压在新一页的页边距线上
   *（视觉上贴着纸顶），加一点余量让行首进入安全区。
   */
  /**
   * 推挤安全余量。
   *
   * ★ 为什么要留够（而不是正好推到 gapEnd）：推挤的实现是「让上一行变高」，
   *   表格重排后本行的落点会有 1~3px 的取整/边框塌陷误差。
   *   实测：按 `gapEnd - top`（=30）推，本行落到 2292，而页隙结束在 2294 ——
   *   仍差 2px 压在页隙上。留 6px 余量即可稳定落在页隙之后，
   *   视觉上也就是行首离纸顶多几像素，与 Word 的观感一致。
   */
  const PUSH_EPS = 6

  /**
   * ★ 页底安全线（本轮根因修复的核心）。
   *
   * 用户的真实反馈是「分页还是不对　表格现在都是跨页的」——
   * 上轮只挡「页隙带 [paperBottom-GAP, paperBottom)」，诊断报 PASS，
   * 但那是因为「只有整行跨进页隙带才算被切」这个判定口径本身是错的。
   *
   * 实测页面几何（_diag_tblcross_margin.mjs）：
   *   A4 高 1123；.wrp-edit padding 上下各 96 → 第 1 页可排版文本带 = [96, 1027]。
   *   页隙带 = [1099, 1123)。于是 (1027, 1099) 之间有 72px 的「无人区」：
   *   落在这一带的行不压页隙带 → 算法放行 → 但纸的下沿在 1123，
   *   这一带的行**压在纸的下边距里**，肉眼就是「表格被页面切断」。
   *
   * 正解：危险带应是**页面可排版文本带的下界**，而不是页隙带的起点。
   *   Word 里表格在页边界断开时，断点发生在**下边距之上**：
   *   最后一行只要触到 pageBottom - PAD_BOT 就必须整行搬到下一页。
   *   所以判定带上界 = pageBottom - PAD_BOT（1027），下界沿用页隙结束（1123）。
   *
   * 于是「页底越界带」= [pageBottom - PAD_BOT, pageBottom) = [1027, 1123)，
   * 覆盖了原来的页隙带 [1099, 1123)，两者合并成一条连续的危险带。
   */
  const PAGE_PAD_BOT = 96
  /** 页面上边距（.wrp-edit 的 padding-top）—— 判断「整表能否装进一页」要用 */
  const PAGE_PAD_TOP = 96

  /**
   * 避让计划条目。
   *
   * `row` 语义（决定 <style> 选择器怎么写）：
   *   - `row >= 1` → 行级/表格内推挤：撑 `tbody > tr:nth-child(row-1)` 的 padding-bottom
   *                  （row=1 时改撑第一行 padding-top）。
   *   - `row === 0` → 整表搬迁：撑**表格前一个兄弟块**的 padding-bottom，
   *                  选择器用 `prevKind` + `prevIndex` 定位（见 avoidCss）。
   */
  type AvoidItem = {
    tbl: number
    row: number
    push: number
    gap?: number
    /** row===0 时：前一个兄弟块的标签名（小写）与它在同类兄弟中的序号 */
    prevTag?: string
    prevIndex?: number
  }

  /**
   * 避让计划：`[{ 表格序号, tbody 内行序号(1-based, 含表头行) | 0=整表搬迁, 推挤量 px }]`。
   * 用 state 存是为了让「重算避让」能触发一次 React 渲染，
   * 把规则注入为 <style>（见下方 avoidCss），从而不被 ProseMirror 抹掉。
   */
  const [avoidPlan, setAvoidPlan] = useState<AvoidItem[]>([])

  /**
   * ★ 计算避让计划的唯一入口（单一真相源）。
   *
   * 为什么必须「从基线一次算完」而不是「迭代叠加」：
   * 之前用「每轮量 → 追加新命中行」的迭代写法，实测产出了自相矛盾的计划 ——
   * 基线里压在页隙上的是行 14（2222~2264 与 2264~2306），
   * 但迭代后却给「不挨页隙的行 13」推了 62px、给「真正压页隙的行 14」推了 32px，
   * 两处推挤叠加后行 14 反而钉在 2277（页隙正中）。原因是每轮的命中判定
   * 都跑在「上一轮推挤已生效」的布局上，基线漂移 → 行号与推挤量互相污染。
   *
   * 正解：先把推挤全清（回到基线），在**同一次基线测量**里，
   * 对每张表只找**第一个**与页隙相交的行，按「推到该页隙之后」算量。
   * 后续行会被这一处推挤自然带到下一页，不需要也不能再各自推 ——
   * 表格是连续流，逐行各推会把表格撑裂。
   */
  const computeAvoidPlan = useCallback(() => {
    const wrap = editWrapRef.current
    if (!wrap) return [] as AvoidItem[]
    const edit = wrap.querySelector<HTMLElement>('.wrp-edit')
    if (!edit) return [] as AvoidItem[]
    const period = PAGE_H + PAGE_GAP

    /** 元素到 .wrp-edit 顶部的偏移（CSS px，未受 zoom 影响）。
     *  完整走 offsetParent 链直到 .wrp-edit（它是 relative，构成定位祖先，链在此终止）。 */
    const relToEdit = (el: HTMLElement) => {
      let y = 0
      let n: HTMLElement | null = el
      while (n && n !== edit) {
        y += n.offsetTop
        n = n.offsetParent as HTMLElement | null
      }
      return y
    }

    const out: AvoidItem[] = []
    const tables = [...edit.querySelectorAll<HTMLTableElement>('table')]
    tables.forEach((tbl, ti) => {
      /**
       * ★★ 整表不跨页（用户 2026-09-28 明确要求：「整张表不跨页」）★★
       *
       * 与 Word 原生行为不同：Word 允许表格在页边界断开、下半截落到下一页；
       * 用户要的是**整张表留在同一页**，当前页装不下就整表搬到下一页。
       *
       * 判据：表格的「行盒总高」是否超过当前页剩余的文本带高度。
       *   当前页文本带下界 = pageBottom - PAGE_PAD_BOT
       *     pageBottom = (floor(tblTop / period) + 1) * period - PAGE_GAP
       *   若 tblTop + tblH > 文本带下界 → 整表装不下 → 把表格推到下一页页首。
       *
       * 推挤量的表达方式与行级避让完全一致（都是「撑高前一个块」）：
       *   - 表格在文档里有前一个兄弟块（p/h/ul…）→ 撑它的 padding-bottom；
       *   - 表格是第一个块 → 撑表格自身第一行的 padding-top。
       *   两者都通过同一个 avoidPlan 通道下发到 <style>，
       *   所以选择器体系无需新增一类（见 avoidCss 的 tableRowOnly 处理）。
       *
       * ★ 与行级避让的优先级：整表避让优先。
       *   整表能装下 → 不动它；装不下 → 整表搬走，这一处不再做行级断开判定
       *   （搬走后整表落在新页页首，除非表本身比一页还高）。
       *   若表本身高过一整页，整表避让无解，退回逐页行级断开（Word 行为）。
       */
      const tblTop = relToEdit(tbl)
      const tblH = tbl.offsetHeight
      const pageIdx = Math.floor(tblTop / period)
      const pageBottom = (pageIdx + 1) * period - PAGE_GAP
      const textBandBottom = pageBottom - PAGE_PAD_BOT
      const pageTextHeight = PAGE_H - PAGE_PAD_TOP - PAGE_PAD_BOT
      if (tblTop + tblH > textBandBottom + 0.5 && tblH <= pageTextHeight) {
        // 整表搬到下一页页首：目标 top = pageBottom + PAGE_GAP + PAGE_PAD_TOP
        const nextPageTextTop = pageBottom + PAGE_GAP + PAGE_PAD_TOP
        const push = Math.round(nextPageTextTop - tblTop + PUSH_EPS)
        if (push <= 0) return
        const prev = tbl.previousElementSibling as HTMLElement | null
        if (prev && prev.tagName !== 'TABLE') {
          /**
           * 撑**前一个兄弟块**的 padding-bottom。
           * 选择器用 `:nth-of-type(k)` 定位 —— k = 该块在「同标签兄弟」里的序号
           * （CSS nth-of-type 只按标签名计数，所以同名后续兄弟的序号会一起后移，
           *   用「块自身在同类兄弟中的序号」正好命中它）。
           */
          const tag = prev.tagName.toLowerCase()
          const sameTag = [...(tbl.parentElement?.children ?? [])].filter(
            (el) => el.tagName.toLowerCase() === tag,
          )
          const prevIndex = sameTag.indexOf(prev) + 1
          if (prevIndex > 0) {
            out.push({ tbl: ti + 1, row: 0, push, gap: pageIdx + 1, prevTag: tag, prevIndex })
          }
        } else {
          // 表格是首块（或前一个也是表格）：撑表格自身第一行的 padding-top
          out.push({ tbl: ti + 1, row: 1, push, gap: pageIdx + 1 })
        }
        return // 整表已搬走，这一处不再做行级断开判定
      }

      // ★ 行号必须以 tbody 为参照 —— 必须与注入的 `tbody > tr:nth-child(N)` 同源
      //   （见 avoidCss 注释：把 thead 行算进来会让选择器整体错位一行）。
      const tbody = tbl.querySelector<HTMLTableSectionElement>('tbody')
      const rows = tbody ? [...tbody.querySelectorAll<HTMLElement>(':scope > tr')] : []
      /** 每张表每处页底最多产出一个「断开点」（该页底上第一个越界的行）。
       *  为什么是「每处页底」而不是「每张表」：一张长表可能跨 5 页，
       *  页 1/2/3/4 各有一个断开点，每一处都要各自让开。 */
      const doneGaps = new Set<number>()
      for (let ri = 0; ri < rows.length; ri++) {
        const tr = rows[ri]
        const absTop = relToEdit(tr)
        const bottom = absTop + tr.offsetHeight
        /**
         * ★ 危险带 = 页面的「页底越界带」[pageBottom - PAGE_PAD_BOT, pageBottom)。
         *
         * 不是原来的页隙带 [pageBottom - PAGE_GAP, pageBottom)：
         *   页面的可排版文本带是 [PAGE_PAD_TOP, pageBottom - PAGE_PAD_BOT]，
         *   页隙带起点（pageBottom - PAGE_GAP）比文本带下界（pageBottom - PAGE_PAD_BOT）
         *   低 72px —— 这 72px 里落下的行既不压页隙带（老算法放行）、
         *   又已经越出文本带（肉眼看到表格压在纸下沿）→ 用户说「表格跨页」。
         *
         * 断开判定用**内容区**（扣掉推挤 padding）还是行盒？
         *   用行盒（box）：只要行的**外框**越出文本带就搬走整行。
         *   用内容区会让「被推挤撑高的行」重新落回带内，导致同一处反复推。
         *   推挤实现是撑高**上一行**，本行外框因此整体下移，行盒判定是自洽的。
         */
        for (let n = Math.max(1, Math.floor(absTop / period)); n <= Math.floor(bottom / period) + 1; n++) {
          const pageBottom = n * period - PAGE_GAP
          const dangerStart = pageBottom - PAGE_PAD_BOT
          const dangerEnd = pageBottom
          if (bottom > dangerStart && absTop < dangerEnd) {
            if (!doneGaps.has(n)) {
              doneGaps.add(n)
              /**
               * ★★ 推挤目标 = **下一页文本带顶端**，不是 dangerEnd（本轮根因修复）。
               *
               * 上一版用 `dangerEnd - absTop`（dangerEnd = pageBottom = 1123），
               * 本意「把行推到页底之下」。但页底（1123）**不是**下一页的起点 ——
               * 下一页文本带从 `pageBottom + PAGE_GAP + PAGE_PAD_TOP` 才开始：
               *   1123（页底）→ 1147（页隙结束，period=1147）→ 1243（+上边距 96）。
               *   于是推到 1123+EPS 的行**正好停在 24px 页隙里**（实测 row18 落在 1125，
               *   页隙是 [1123,1147)）—— 肉眼就是「表格被切断在页缝上」，
               *   用户反馈「跨页断行没搞好」。
               *
               * 正解：与**整表避让**用同一个目标（见上方 nextPageTextTop），
               * 让被推的行落进下一页的可排版文本带内：pageBottom + PAGE_GAP + PAGE_PAD_TOP。
               * 这样推完后行首在下一页纸顶下方 96px（页上边距）处，与 Word 观感一致。
               */
              const nextPageTextTop = pageBottom + PAGE_GAP + PAGE_PAD_TOP
              out.push({ tbl: ti + 1, row: ri + 1, push: Math.round(nextPageTextTop - absTop + PUSH_EPS), gap: n })
            }
            break
          }
        }
      }
    })
    return out
  }, [PAGE_H, PAGE_GAP, PAGE_PAD_BOT])

  /** 清掉测量期可能残留的临时行内样式（计划由 <style> 承载，不在这里） */
  const clearInlinePush = useCallback(() => {
    const edit = editWrapRef.current?.querySelector<HTMLElement>('.wrp-edit')
    if (!edit) return
    // 整表搬迁：被撑高的是表格**前一个兄弟块**（p/h/ul…），同样要清
    edit.querySelectorAll<HTMLElement>('.ProseMirror > [data-wrp-push]').forEach((el) => {
      if (el.tagName === 'TABLE') return
      el.removeAttribute('data-wrp-push')
      el.style.paddingBottom = ''
      delete el.dataset.wrpPad
    })
    edit.querySelectorAll<HTMLElement>('tr[data-wrp-push]').forEach((tr) => {
      tr.removeAttribute('data-wrp-push')
      tr.querySelectorAll<HTMLElement>('td, th').forEach((c) => {
        c.style.paddingTop = ''
        c.style.paddingBottom = ''
      })
    })
    // 兜底：清掉任何残留的推挤内联 padding（防止某轮中断后留下脏数据）
    edit.querySelectorAll<HTMLElement>('table td[style], table th[style]').forEach((c) => {
      const el = c as HTMLElement
      if (el.dataset.wrpPad === '1') {
        el.style.paddingTop = ''
        el.style.paddingBottom = ''
        delete el.dataset.wrpPad
      }
    })
  }, [])

  /**
   * 重算避让计划 —— **逐处断开点收敛**。
   *
   * ★ 为什么必须收敛（而不是一次算完）：一张表可能跨好几种页。
   *   实测 90 行长表：跨 5 页，在页 1/2/3/4 各有 6 行被切开。
   *   只推每一处「当前第一个断开点」的话，推完页 1 之后页 2/3/4 的断开点
   *   一个都没处理（实测：收敛前 6 处 → 单次计算后仍剩 4 处）。
   *   所以要对每一处未解决的断开点各推一次，循环到没有为止。
   *
   * ★ 但每轮必须建立在「上一轮推挤已经生效」的真实布局上，且**不能**在
   *   测量时把已生效的推挤摘掉 —— 摘掉会让每轮都只看到「第一个」断开点，
   *   于是永远在推同一行（这正是之前那版「越收敛越错」的根因：
   *   它用 setState 提交后又在测量时 disable 了样式，把推挤效果抹平了）。
   *   正确做法：把计划**累积**成一个数组，每轮开始时先用行内 style 如实重放
   *   当前累积计划（仅仅为了测量），量出下一个未处理的断开点，追加进计划，
   *   再进入下一轮；全部算完后一次 setState 用 <style> 表达最终结果。
   *
   *   用行内 style 做「测量用重放」是安全的：它只活在函数内，
   *   每轮开头先清掉、重新按累积计划写一遍，不受 ProseMirror 何时重绘影响
   *  （measure 是同步的，且在同一个任务里完成，不存在被外部重绘插入的窗口）。
   */
  const avoidTableRowSplit = useCallback(() => {
    const wrap = editWrapRef.current
    if (!wrap) return
    const edit = wrap.querySelector<HTMLElement>('.wrp-edit')
    if (!edit) return

    clearInlinePush()

    /** 测量用：把累积计划按行内 style 重放到 DOM 上（只影响本次测量） */
    const replay = (plan: AvoidItem[]) => {
      clearInlinePush()
      const tables = [...edit.querySelectorAll<HTMLTableElement>('table')]
      for (const item of plan) {
        const tbl = tables[item.tbl - 1]
        if (!tbl) continue
        if (item.row === 0) {
          // 整表搬迁：撑前一个兄弟块的 padding-bottom（与最终 <style> 表达一致）
          const prev = tbl.previousElementSibling as HTMLElement | null
          if (!prev || prev.tagName === 'TABLE') continue
          prev.style.paddingBottom = `${item.push}px`
          prev.dataset.wrpPad = '1'
          prev.dataset.wrpPush = '1'
          continue
        }
        const tbody = tbl.querySelector<HTMLTableSectionElement>('tbody')
        const rows = tbody ? [...tbody.querySelectorAll<HTMLElement>(':scope > tr')] : []
        // 推挤实现 = 让「上一行」变高（padding-bottom）；首行则自身 padding-top
        const target = item.row > 1 ? rows[item.row - 2] : rows[0]
        if (!target) continue
        const cells = target.querySelectorAll<HTMLElement>('td, th')
        cells.forEach((c) => {
          if (item.row > 1) c.style.paddingBottom = `${item.push}px`
          else c.style.paddingTop = `${item.push}px`
          c.dataset.wrpPad = '1'
        })
        target.dataset.wrpPush = '1'
      }
      void edit.offsetHeight // 强制重排，确保下一轮测量读到的是重放后的布局
    }

    /**
     * 每个「表格号 + 页隙号」只处理一次。
     *
     * ★ 这是收敛循环不跑飞的关键，也是踩出来的血泪教训：
     *   推挤的实现是「让上一行变高」→ 表格下移。若不加限制地循环，
     *   就会变成「行 27 压页隙 2 → 推它 → 行 26 又压页隙 2 → 推它 →
     *   行 25 又压页隙 2 …」的无限回退（实测：24 轮里把行 5..27 全部
     *   按 42px 步长一路推下去，推挤量长到 983px，页面被撑爆）。
     *   原因：把某行推过页隙后，页隙位置不变，而表格整体以下的行都往下挪了
     *   —— 于是**紧邻页隙上方**又出现一个新的「压隙行」。
     *   但这在 Word 语义上是错的：一处页断只需要一次搬迁，
     *   搬迁后该页隙上下不再有行跨它（行 27 已在页隙之后，
     *   行 26 往上属于上一页，它压不压页隙由**上一页**的容量决定，
     *   而不是由这次搬迁决定）。
     *   所以：一个页隙修一次，修完就标记为已处理，不再对同一页隙做判定。
     */
    const handledGaps = new Set<string>()
    const plan: AvoidItem[] = []

    for (let round = 0; round < 24; round++) {
      replay(plan)
      const found = computeAvoidPlan().filter((f) => !handledGaps.has(`${f.tbl}:${f.gap}`))
      const next = found[0]
      if (!next) break
      handledGaps.add(`${next.tbl}:${next.gap}`)
      plan.push(next)
    }

    replay([]) // 测量结束，清掉行内重放（最终结果由 <style> 表达）

    setAvoidPlan((prev) => {
      const same =
        prev.length === plan.length &&
        prev.every((p, i) => {
          const q = plan[i]
          return q && p.tbl === q.tbl && p.row === q.row && p.push === q.push
        })
      return same ? prev : plan
    })
  }, [computeAvoidPlan, clearInlinePush])

  /**
   * 避让规则 → 一条注入的 <style>。
   *
   * 为什么不直接写行内 style：tr 的 DOM 由 ProseMirror 拥有，
   * setContent/事务重绘会从文档模型重建它，抹掉一切外部写入（实测确认）。
   * CSS 规则挂在 .wrp-edit（React 层拥有、ProseMirror 不动）上，
   * 命中目标行 —— 重建 tr 也照样命中。
   *
   * ★ 选择器必须用 `tbody > tr:nth-child(N)`，不能用 `tr:nth-child(N)`。
   *   ProseMirror 的表格结构是 `<table><tbody><tr>…`（表头行也在 tbody 里，
   *   由 th 单元格区分，而不是单独的 thead 元素），但**导入的 OOXML 可能生成
   *   `<table><thead><tr>…</thead><tbody><tr>…`**。
   *   用裸 `tr:nth-child(N)` 时，N 在「tr 的兄弟集合」里数 ——
   *   若存在 thead，tbody 里的第一行是 nth-child(1)，与
   *   我们「行号以 tbody 为参照」的编号差一位。
   *   实测症状：算法正确选中了「压页隙的那一行」，但注入的规则打到了它的前一行
   *  （量到行 14 压页隙，规则却写 nth-child(13)），推错位置、页隙没被让开。
   *
   * ★ 推挤必须用「上一行的 padding-bottom」而非「本行的 padding-top」。
   *   第二个隐蔽的坑（比 margin 那个更绕）：
   *   - tr 的 margin 在 border-collapse: collapse 表格里被规范强制忽略 ——
   *     getComputedStyle 会乖乖返回 "32px"，但 offsetTop 一点不动（实测确认）。
   *   - 改用本行 padding-top 确实生效了（行高 42 → 70，表格总高 +28），
   *     但 **padding 是把行的「内容区」下推，行的顶边仍锚在原处** ——
   *     于是行的 offsetTop 不变、offsetHeight 变大，行的范围变成 2264~2334，
   *     依旧跨在页隙 [2270,2294] 上，只是从「拦腰截断」变成「整体压住」。
   *   正解：给**上一行**加 padding-bottom。上一行变高 → 本行顶边被真实下移，
   *   本行整体落到 gapEnd 之后，行的起止都不再与页隙相交。
   *   若被推的是表格第一行（无上一行），则退回用本行 padding-top（此时页隙在
   *   表格之前，把表格整体下推即可，行首已经在页隙之后，不会相交）。
   */
  const avoidCss = avoidPlan.length
    ? avoidPlan
        .map((p) => {
          const t = `.wrp-edit .ProseMirror table:nth-of-type(${p.tbl}) tbody`
          if (p.row === 0) {
            /**
             * 整表搬迁：撑**前一个兄弟块**的 padding-bottom。
             * 选择器 `.ProseMirror > p:nth-of-type(k)` —— nth-of-type 按标签计数，
             * prevIndex 就是该块在同标签兄弟中的序号（computeAvoidPlan 里算好）。
             * ★ 不能用 `.ProseMirror > :nth-child(k)`：nth-child 按全部子元素计数，
             *   而 prevIndex 是按标签计数的（表格前的段落序号）。
             */
            const sel = `.wrp-edit .ProseMirror > ${p.prevTag}:nth-of-type(${p.prevIndex})`
            return `${sel} { padding-bottom: ${p.push}px !important; }`
          }
          if (p.row > 1) {
            // 推「上一行」的 padding-bottom → 本行整体下移
            const prev = `${t} > tr:nth-child(${p.row - 1})`
            return (
              `${prev} > td, ${prev} > th { padding-bottom: ${p.push}px !important; }`
            )
          }
          // 表格第一行：整表下推（padding-top 加在表的第一个单元格上）
          const first = `${t} > tr:nth-child(1)`
          return `${first} > td, ${first} > th { padding-top: ${p.push}px !important; }`
        })
        .join('\n')
    : ''

  /**
   * 推挤行的视觉修饰 —— ★ 已废弃为空字符串（2026-09-28 第二轮修正）。
   *
   * 上一版这里去掉「承载 padding 的那一行」的 border-bottom 并令其背景透明，
   * 想让页底那段留白"看起来像干净的纸"。但实测（scripts/_tmp_r10.mjs）证明这是**错的**：
   *
   *   avoidCss 给 tr:nth-child(k) 加 padding-bottom（k = 被推行 - 1），
   *   而 splitCss 也针对同一个 tr:nth-child(k) 去边框 —— 于是**承载留白的那一行
   *   连同它自己的文字一起丢了边框**：文字留在页 1 上，却没有表格线包着，
   *   视觉上就是"一行孤儿文字"，用户读成「没有跨页断行」。
   *
   * 正解：**什么都不用改**。承载留白的那一行本来就该保持完整表格线；
   * 页与页之间那条 24px 的缝由 .wrp-cutter（不透明遮断条，z-index 高于内容层）
   * 统一盖住 —— 横跨页底的留白自然被遮断，不需要逐行去边框。
   *
   * 保留这个常量而不是删掉引用，是为了留下「为什么不能按行去边框」的证据。
   */
  const splitCss = ''

  /**
   * 需要遮断的页隙条数 = 「内容真正跨越了几个页隙」。
   *
   * ★ 为什么不用 bgPageCount：纸张数由 contentPages/visPages 合成，
   *   而 visPages 是「视口铺满」的兜底 —— 内容少、视口高时纸张会多铺几张，
   *   但那些纸之间的页隙**下面没有内容**，不需要遮断条（多盖会挡住纸面）。
   *   反过来内容高时纸张数又依赖 contentPages 的刷新时机（异步），
   *   拿它算遮断条数会漏（实测：内容跨 3 页只生成 1 条）。
   *   所以这里直接按**内容高度**算，与纸张层解耦。
   */
  const [cutterCount, setCutterCount] = useState(0)

  /** 内容变化时刷新「内容页数」（挂在 computeStats 里调用） */
  const refreshContentPages = useCallback(() => {
    const wrap = editWrapRef.current
    if (!wrap) return
    const p = PAGE_H + PAGE_GAP
    // 先做表格行避让（会改变布局高度），再量内容高度
    avoidTableRowSplit()
    const pm = wrap.querySelector<HTMLElement>('.wrp-edit .ProseMirror')
    const contentH = pm ? pm.scrollHeight : 0
    // 内容高 + 上下页边距(96*2) = 纸张内容区真实所需高
    const n = Math.max(1, Math.ceil((contentH + 192) / p))
    setContentPages((prev) => (prev === n ? prev : n))
    // 遮断条：内容跨过的页隙数 = 内容页数 - 1（与纸张层无关）
    const c = Math.max(0, n - 1)
    setCutterCount((prev) => (prev === c ? prev : c))
    // 影子表头：内容高度/推挤量变了以后，每页页顶要不要重复表头也要重算
    refreshShadowHeadersRef.current()
  }, [PAGE_H, PAGE_GAP, avoidTableRowSplit])

  /** ref 镜像：computeStats 定义在 refreshContentPages 之前，经此调用（避开 use-before-declare） */
  const refreshContentPagesRef = useRef<() => void>(() => {})
  refreshContentPagesRef.current = refreshContentPages

  /* ──────────────────────────────────────────────────────────────────────
   * 影子表头（视觉层，不写回文档模型）
   *
   * 用户的诉求是「跨页 + 表头每页重复」（2026-09-28）。导出件（.docx）里这是
   * OOXML 真排版：<w:tblHeader/> 让 Word 自己每页重画表头。
   * 但编辑视图是一整条**连续流** + 手搓的「假分页」白纸 —— 浏览器根本不会分页，
   * 所以表头只会出现在表格最上面一次，跨页后下一页页顶是空的。
   *
   * 做法：找出「首行落在第 P 页、而后续行落到第 P 页之后」的表格，
   * 为它在该表跨过的每个后续页页顶克隆一行表头（用 innerHTML 复刻，视觉同款）。
   *   位置 top = (n-1) * period + PAGE_PAD_TOP   —— 即第 n 页的可排版文本带顶端。
   *   这正好补上「表头是否随页面重复」这一块，与导出件行为对齐。
   *
   * ★ 三个约束：
   *   1. 必须放在 .wrp-pages（z-index 0）与 .wrp-edit（z-index 1）**之间**，
   *      z-index 置中：否则会被内容盖住（z < 1）或反过来盖住正文（z > 1）。
   *   2. pointer-events: none —— 它是纯装饰，不能吃掉编辑区的点击。
   *   3. 克隆行的 DOM 不归 ProseMirror 管，所以必须克隆成**独立块**（不能塞回表里），
   *      并且每次重算都要整批重建（避免脏行残留）。
   * ────────────────────────────────────────────────────────────────────── */
  /** [{ 表格序号, 页号 }, …] —— 页号 n 表示要在第 n 页页顶画一次表头 */
  const [shadowHeaders, setShadowHeaders] = useState<{ tbl: number; page: number }[]>([])

  const computeShadowHeaders = useCallback(() => {
    const wrap = editWrapRef.current
    if (!wrap) return [] as { tbl: number; page: number }[]
    const edit = wrap.querySelector<HTMLElement>('.wrp-edit')
    if (!edit) return [] as { tbl: number; page: number }[]
    const period = PAGE_H + PAGE_GAP

    const relToEdit = (el: HTMLElement) => {
      let y = 0
      let n: HTMLElement | null = el
      while (n && n !== edit) {
        y += n.offsetTop
        n = n.offsetParent as HTMLElement | null
      }
      return y
    }

    const out: { tbl: number; page: number }[] = []
    const tables = [...edit.querySelectorAll<HTMLTableElement>('table')]
    tables.forEach((tbl, ti) => {
      const tbody = tbl.querySelector<HTMLTableSectionElement>('tbody')
      const rows = tbody ? [...tbody.querySelectorAll<HTMLElement>(':scope > tr')] : []
      // 只处理「有表头行 + 至少 2 行」的表（与导出件 tblHeader 的判定一致：
      // 单行表不加 tblHeader，避免 Word 多画一行；这里同样不画影子表头）
      if (rows.length < 2) return
      const head = rows[0]
      const headTop = relToEdit(head)
      const headBottom = headTop + head.offsetHeight
      const firstPage = Math.floor(headTop / period)
      // 表格最后一行落在第几页 —— 中间每一页都要重复表头
      const last = rows[rows.length - 1]
      const lastBottom = relToEdit(last) + last.offsetHeight
      const lastPage = Math.floor((lastBottom - 1) / period)
      for (let p = firstPage + 1; p <= lastPage; p++) {
        const pageTop = p * period + PAGE_PAD_TOP
        // 表头行本身已经跨过这个页顶（即表头就被推到这页开头）→ 不需要影子
        if (headBottom > pageTop) continue
        out.push({ tbl: ti + 1, page: p + 1 }) // 对外用 1-based 页号
      }
    })
    return out
  }, [PAGE_H, PAGE_GAP, PAGE_PAD_TOP])

  /** 重算影子表头（由 refreshContentPages 触发，与避让同一时机） */
  const refreshShadowHeaders = useCallback(() => {
    const next = computeShadowHeaders()
    setShadowHeaders((prev) =>
      prev.length === next.length && prev.every((p, i) => p.tbl === next[i].tbl && p.page === next[i].page)
        ? prev
        : next,
    )
  }, [computeShadowHeaders])

  /** ref 镜像：refreshContentPages 里经此调用影子表头重算（避开 use-before-declare） */
  const refreshShadowHeadersRef = useRef<() => void>(() => {})
  refreshShadowHeadersRef.current = refreshShadowHeaders

  /**
   * 诊断钩子：把「真实实现」暴露到 window，供 scripts/_diag_tablecross.mjs 直接驱动。
   *
   * 为什么需要它：诊断脚本不能自己复刻一份算法 —— 那样验的是脚本，不是产品。
   * 之前几版诊断「避让后仍有行跨页」正是因为脚本内复刻的算法与真实实现不同源。
   * 生产环境无副作用（只在编辑器挂载时写一个引用）。
   */
  useEffect(() => {
    if (!editorMode || tab !== 'preview') return
    const w = window as unknown as {
      __wrpAvoid?: () => void
      __wrpClearAvoid?: () => void
      __wrpEditor?: Editor | null
      __wrpSetHtml?: (html: string) => void
    }
    w.__wrpAvoid = () => {
      avoidTableRowSplit()
      refreshContentPages()
    }
    /** 清掉避让计划，回到「避让前」基线，供诊断做 A/B 对比。
     *  注意：真正的推挤由 state 驱动的 <style> 表达（setAvoidPlan([]) 即回基线），
     *  行内 style 只是测量期的临时产物，computeAvoidPlan 每次都会先清干净。 */
    w.__wrpClearAvoid = () => {
      clearInlinePush()
      setAvoidPlan([])
    }
    /**
     * 内容注入钩子：让诊断脚本用「与产品同一条路径」（Tiptap setContent）
     * 造数据，而不是直接改 DOM —— 直接改 DOM 会被 value 同步 effect 抹掉，
     * 量出来的几何不是产品真实状态。
     */
    w.__wrpSetHtml = (html: string) => {
      const e = editorRef.current
      if (e) e.commands.setContent(html || '<p></p>', false)
      // 注入后立刻刷新页数/遮断条（生产路径由 computeStats → flushEdit 触发）
      requestAnimationFrame(() => refreshContentPages())
    }
    return () => {
      delete w.__wrpAvoid
      delete w.__wrpClearAvoid
      delete w.__wrpSetHtml
      delete w.__wrpEditor
    }
  }, [editorMode, tab, avoidTableRowSplit, refreshContentPages, clearInlinePush])

  /** 把编辑器实例暴露出去（诊断脚本需要 setContent），随实例就绪更新 */
  useEffect(() => {
    const w = window as unknown as { __wrpEditor?: Editor | null }
    w.__wrpEditor = editorInst
    return () => {
      w.__wrpEditor = null
    }
  }, [editorInst])

  /** 视口变化时刷新「视口页数」（纸张至少铺满可视区，否则下方露灰底） */
  useEffect(() => {
    const wrap = editWrapRef.current
    if (!wrap) return
    const measure = () => {
      const k = Math.max(0.1, zoomPct / 100)
      // 47 = wrap padding-top(10) + 标尺(17+8) + wrap padding-bottom(12)
      const n = Math.max(1, Math.ceil((wrap.clientHeight - 47) / k / (PAGE_H + PAGE_GAP)))
      setVisPages((prev) => (prev === n ? prev : n))
    }
    measure()
    const ro = new ResizeObserver(measure)
    ro.observe(wrap)
    return () => ro.disconnect()
  }, [zoomPct, editorMode])

  /** 编辑器挂载/内容初始就绪后先算一次内容页数（否则首屏纸张数不准） */
  useEffect(() => {
    if (!editorInst) return
    refreshContentPages()
  }, [editorInst, refreshContentPages])

  /**
   * 背景纸张层渲染几张 A4 纸 = 纸张层自己的页数。
   * 必须与 paperMinH 严格同源，否则纸张层比内容层高会撑大 wrap 的 scrollHeight
   * （绝对定位元素同样产生 scrollable overflow），导致滚动条变长 + 滚到底露灰底。
   */
  const bgPageCount = paperPages

  const onWrapScroll = useCallback(() => {
    const wrap = editWrapRef.current
    if (!wrap) return
    const k = zoomPct / 100
    setCurPage(Math.min(docStats.pages, Math.floor(wrap.scrollTop / ((PAGE_H + PAGE_GAP) * k)) + 1))
    // 滚动时环绕浮层坐标会漂移 —— 顺手收起
    setActiveImg(null)
    setWrapPop(null)
  }, [zoomPct, docStats.pages])

  /* ── WPS 借鉴：缩放适配（页宽 / 整页）+ Ctrl+滚轮 ──
   *   fitWidth：纸宽 794 铺满可视区宽度；fitPage：一页高铺满可视区高度。
   *   Ctrl+滚轮走原生非被动监听（React 的 onWheel 是 passive，preventDefault 无效）。 */
  const fitWidth = useCallback(() => {
    const wrap = editWrapRef.current
    if (!wrap) return
    const avail = Math.max(200, wrap.clientWidth - 24)
    setZoomPct(Math.max(50, Math.min(300, Math.round((avail / 794) * 100))))
  }, [])
  const fitPage = useCallback(() => {
    const wrap = editWrapRef.current
    if (!wrap) return
    const avail = Math.max(200, wrap.clientHeight - 40)
    setZoomPct(Math.max(50, Math.min(300, Math.round((avail / (PAGE_H + PAGE_PAD_TOP + PAGE_PAD_BOT)) * 100))))
  }, [])
  useEffect(() => {
    const wrap = editWrapRef.current
    if (!wrap) return
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return
      e.preventDefault()
      setZoomPct((z) => Math.max(50, Math.min(300, z + (e.deltaY < 0 ? 10 : -10))))
    }
    wrap.addEventListener('wheel', onWheel, { passive: false })
    return () => wrap.removeEventListener('wheel', onWheel)
  }, [])

  /* ── WPS 借鉴：标尺交互 ──
   *   左/右页边距块可拖 → 改 sectMargin（视觉 + 状态，导出经 sectPr 落盘）；
   *   首行/悬挂缩进三角可拖 → 改当前段 margin-left / text-indent（走 ProseMirror transaction，真写 OOXML）；
   *   点刻度区 → 加制表位（tabStops 状态，标尺渲染标记）。
   *   拖拽用 window 级 pointermove/up（避免鼠标移出标尺丢事件），dx 除回 zoom。 */
  const rulerDragRef = useRef<{ kind: 'l' | 'r' | 'first' | 'hang'; startX: number; startVal: number } | null>(null)
  const onMarginDown = (side: 'l' | 'r') => (e: React.PointerEvent) => {
    e.preventDefault()
    e.stopPropagation()
    const startX = e.clientX
    const startVal = sectMargin[side]
    rulerDragRef.current = { kind: side, startX, startVal }
    const move = (ev: PointerEvent) => {
      const dx = (ev.clientX - startX) / (zoomPct / 100)
      const v = Math.max(0, Math.min(300, Math.round(startVal + dx)))
      setSectMargin((m) => ({ ...m, [side]: v }))
    }
    const up = () => {
      window.removeEventListener('pointermove', move)
      window.removeEventListener('pointerup', up)
      rulerDragRef.current = null
    }
    window.addEventListener('pointermove', move)
    window.addEventListener('pointerup', up)
  }
  const onIndentDown = (kind: 'first' | 'hang') => (e: React.PointerEvent) => {
    e.preventDefault()
    e.stopPropagation()
    const startX = e.clientX
    const startVal = indentState[kind]
    rulerDragRef.current = { kind, startX, startVal }
    const move = (ev: PointerEvent) => {
      const dx = (ev.clientX - startX) / (zoomPct / 100)
      const v = Math.max(0, Math.min(554, Math.round(startVal + dx)))
      setIndentState((s) => ({ ...s, [kind]: v }))
      const ed = editorInst
      if (!ed) return
      if (kind === 'first') ed.chain().focus().updateAttributes('paragraph', { style: `text-indent:${v}px` }).run()
      else ed.chain().focus().updateAttributes('paragraph', { style: `margin-left:${v}px` }).run()
    }
    const up = () => {
      window.removeEventListener('pointermove', move)
      window.removeEventListener('pointerup', up)
      rulerDragRef.current = null
    }
    window.addEventListener('pointermove', move)
    window.addEventListener('pointerup', up)
  }
  const onRulerScaleClick = (e: React.MouseEvent) => {
    const scale = (e.currentTarget as HTMLElement).getBoundingClientRect()
    const x = (e.clientX - scale.left) / (zoomPct / 100)
    if (x < 0 || x > scale.width / (zoomPct / 100)) return
    setTabStops((ts) => (ts.includes(Math.round(x)) ? ts : [...ts, Math.round(x)].sort((a, b) => a - b)))
  }

  /** 缩略图导航：点页跳转变焦（scrollIntoView 当前页 top） */
  const onThumbClick = (pageTop: number) => {
    const wrap = editWrapRef.current
    if (!wrap) return
    wrap.scrollTo({ top: pageTop * (zoomPct / 100) - 12, behavior: 'smooth' })
    editorInst?.commands.focus('start')
  }

  /** 右键上下文菜单（WPS 借鉴）：阻止默认，定位自绘菜单 */
  const onEditContextMenu = (e: React.MouseEvent) => {
    if (!editorMode) return
    e.preventDefault()
    setCtxMenu({ x: e.clientX, y: e.clientY })
  }
  const ctxRun = (fn: () => void) => {
    setCtxMenu(null)
    fn()
  }

  /** 选区浮工具条：selectionUpdate 时若选区非空且落在编辑器内，定位浮条 */
  const placeFloatBar = useCallback(() => {
    const ed = editorInst
    if (!ed || !ed.isFocused) {
      setFloatBar(null)
      return
    }
    const { from, to } = ed.state.selection
    if (from === to) {
      setFloatBar(null)
      return
    }
    try {
      const start = ed.view.coordsAtPos(from)
      const end = ed.view.coordsAtPos(to)
      const y = Math.min(start.top, end.top) - 8
      const x = (start.left + end.right) / 2
      setFloatBar({ x, y })
    } catch {
      setFloatBar(null)
    }
  }, [editorInst])

  /* ── 插入图片（就地编辑）：选本地图 → 等比压到版心宽内 → insertImage 写入编辑面 ──
     图片插入后走既有写回链路（htmlToOoxml → createImageSink 落 media），
     width/height 属性显式写上 —— 导出尺寸以此为准，不受显示层 max-width 影响。 */
  // 版心宽 = 794 - 左右页边距（sectMargin 由导入文档的 sectPr 初始化，拖标尺可改；
  // 李渡文档 92.6/56.7 → 644.7，设计院表格 tblW 643px 正好放满 —— 固定 554 会装不下）
  const EDIT_MAX_W = Math.max(320, 794 - sectMargin.l - sectMargin.r)
  const imgFileRef = useRef<HTMLInputElement>(null)
  const insertImage = useCallback(() => {
    imgFileRef.current?.click()
  }, [])
  const onImgPicked = useCallback(
    async (file: File) => {
      if (file.size > MAX_IMAGE_BYTES) {
        setErr(`图片 ${(file.size / 1024 / 1024).toFixed(1)}MB 超过 ${MAX_IMAGE_BYTES / 1024 / 1024}MB 上限，请压缩后再插`)
        return
      }
      try {
        const buf = await file.arrayBuffer()
        let bin = ''
        const bytes = new Uint8Array(buf)
        const CH = 0x8000
        for (let i = 0; i < bytes.length; i += CH) bin += String.fromCharCode(...bytes.subarray(i, i + CH))
        const uri = `data:${file.type || 'image/png'};base64,${btoa(bin)}`
        // 先解码拿原始尺寸：超版心宽的等比缩，显式 width/height 写进 node attrs（导出尺寸与显示一致）
        const probe = new Image()
        probe.src = uri
        await probe.decode().catch(() => undefined)
        const natW = probe.naturalWidth || 600
        const natH = probe.naturalHeight || 400
        const scale = natW > EDIT_MAX_W ? EDIT_MAX_W / natW : 1
        const w = Math.round(natW * scale)
        const h = Math.round(natH * scale)
        // Tiptap：setImage 插入节点并带上 width/height/dataWrap（node attrs 会序列化进 HTML，
        // 导出时 html2ooxml 按 attr 生成 inline 或 anchor）
        editorRef.current
          ?.chain()
          .focus()
          .setImage({ src: uri, width: w, height: h, dataWrap: 'inline' } as never)
          .run()
        onEditInputRef.current()
      } catch (e) {
        setErr('插入图片失败：' + (e instanceof Error ? e.message : String(e)))
      }
    },
    [],
  )

  /* ── 图片环绕浮层：点选编辑面里的图片浮出，四档环绕（嵌入/四周左/四周右/上下）──
     浮层挂 editzone（不滚动锚），坐标 = 图片 rect 相对 editzone；滚动即收起（onWrapScroll）。
     监听走 JSX onMouseDown 而非 useEffect addEventListener —— 编辑面挂载晚于 effect 首跑
     （模板载入中 host 为 null），原生监听会永久漏绑（2026-09-27 实测浮层不出现）。 */
  const [activeImg, setActiveImg] = useState<HTMLImageElement | null>(null)
  const [wrapPop, setWrapPop] = useState<{ x: number; y: number } | null>(null)
  const editZoneRef = useRef<HTMLDivElement>(null)
  const closeWrapPop = useCallback(() => {
    setActiveImg(null)
    setWrapPop(null)
  }, [])
  const onEditMouseDown = useCallback((e: ReactMouseEvent) => {
    const t = e.target as HTMLElement
    if (t.tagName !== 'IMG') {
      closeWrapPop()
      return
    }
    if (t.hasAttribute('data-img-missing')) return // 占位图不给环绕
    const zone = editZoneRef.current
    const img = t as HTMLImageElement
    setActiveImg(img)
    if (!zone) return
    const zoneR = zone.getBoundingClientRect()
    const r = img.getBoundingClientRect()
    // 浮层贴图片右上角；越界收回到编辑区内
    setWrapPop({
      x: Math.min(r.right - zoneR.left, zoneR.width - 160),
      y: Math.max(0, r.top - zoneR.top - 44),
    })
  }, [closeWrapPop])

  /** 应用环绕：写 Tiptap image node 的 dataWrap/dataHside 属性
   *  → getHTML() 输出 data-wrap/data-hside 属性 → html2ooxml 决定 inline 还是 anchor+wrap。
   *  float 视觉由 CSS 按 data-wrap 属性驱动（不再直改 DOM style，避免被 Tiptap 重渲染覆盖） */
  const applyWrap = useCallback(
    (wrap: 'inline' | 'square' | 'topBottom', hside: 'left' | 'right' = 'left') => {
      const e = editorRef.current
      if (!e || !activeImg) return
      // 选中该图片节点：按 src 匹配（同一文档内 src 唯一，data URI 即指纹）
      const src = activeImg.getAttribute('src') ?? ''
      let targetPos: number | null = null
      e.state.doc.descendants((node, pos) => {
        if (node.type.name === 'image' && node.attrs.src === src) {
          targetPos = pos
          return false
        }
        return true
      })
      if (targetPos == null) return
      e.chain()
        .focus()
        .setNodeSelection(targetPos)
        .updateAttributes('image', {
          dataWrap: wrap,
          dataHside: wrap === 'square' ? hside : null,
        })
        .run()
      closeWrapPop()
      onEditInputRef.current()
    },
    [activeImg, closeWrapPop],
  )

  /** HTML → docx 字节：全量重建 body（保留 sectPr / 根属性 / 其余部件）。
   *  底座必须用模板字节而非上一稿：body 每次从编辑面 HTML 全量重写，编辑面里的图
   *  （含模板原有图，装载时已转 data: URI）全部经 sink 重新注入 —— 若在旧稿上重建，
   *  旧稿里上一轮注入的 media 会残留累积，每编辑一次多一份（2026-09-27 实测 media 翻倍）。 */
  const buildEditedDocx = useCallback(
    async (html: string): Promise<Uint8Array> => {
      const base = tplBytes ?? editBaseRef.current
      if (!base) throw new Error('没有可编辑的文档')
      const zip = await JSZip.loadAsync(base)
      const num = await ensureNumbering(zip)
      const images = await createImageSink(zip)
      await images.ensure('word/document.xml')
      const bodyXml = htmlToOoxml(html, num, (bytes, ext) => images.add(bytes, ext, 'word/document.xml'))
      const { sectPr, rootAttrs } = editMetaRef.current
      const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      const xml = `${head}<w:document${rootAttrs}><w:body>${bodyXml}${sectPr}</w:body></w:document>`
      zip.file('word/document.xml', xml)
      await images.finish()
      return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } })
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [tplBytes],
  )

  /** 防抖写回：onInput → 600ms → 生成字节 → 状态 + IndexedDB。卸载时 flush。 */
  const flushEdit = useCallback(async () => {
    const html = pendingHtmlRef.current
    if (html == null) return
    pendingHtmlRef.current = null
    try {
      const bytes = await buildEditedDocx(html)
      setEditedBytes(bytes)
      void saveEditedDocx(table.id, tplIdRef.current, bytes)
      computeStats()
    } catch (e) {
      setErr('保存编辑失败：' + (e instanceof Error ? e.message : String(e)))
    }
  }, [buildEditedDocx, table.id, computeStats])

  const onEditInput = useCallback(() => {
    pendingHtmlRef.current = editorRef.current ? cleanHtml(editorRef.current.getHTML()) : editorHtml
    if (editTimerRef.current) window.clearTimeout(editTimerRef.current)
    editTimerRef.current = window.setTimeout(() => void flushEdit(), 600)
    // 页数实时跟进（跨页打字时页隙/页码跟着变，不等 flush）
    if (statsTimerRef.current) window.clearTimeout(statsTimerRef.current)
    statsTimerRef.current = window.setTimeout(computeStats, 200)
  }, [flushEdit, computeStats, editorHtml])
  onEditInputRef.current = onEditInput

  // 卸载时把未落盘的编辑 flush 掉（fire-and-forget：IndexedDB 写入在组件卸载后仍会完成）
  useEffect(
    () => () => {
      if (editTimerRef.current != null) window.clearTimeout(editTimerRef.current)
      if (statsTimerRef.current != null) window.clearTimeout(statsTimerRef.current)
      const html = pendingHtmlRef.current
      if (html == null) return
      pendingHtmlRef.current = null
      void buildEditedDocx(html)
        .then((bytes) => saveEditedDocx(table.id, tplIdRef.current, bytes))
        .catch(() => {})
    },
    [buildEditedDocx, table.id],
  )

  /** 放弃编辑稿：回到绑定产物 */
  const discardEdit = useCallback(() => {
    setEditedBytes(null)
    setEditRestored(false)
    editBootRef.current = false
    editFpRef.current = null
    void clearEditedDocx(table.id)
  }, [table.id])

  /**
   * 同步选区格式 → Ribbon 按下态（WPS 质感的关键：
   * 光标停在粗体字上时，「B」按钮应该是按下的。
   * 没有这个，功能区就是一堆「看起来能点但不知道当前状态」的按钮）。
   *
   * 订阅 Tiptap 的 selectionUpdate —— 光标移动、选区变化、命令执行后都会触发。
   * 只比对变化后再 setState，避免每次光标移动都重渲染整个面板。
   */
  useEffect(() => {
    const e = editorInst
    if (!e) return
    const sync = () => {
      const next = {
        bold: e.isActive('bold'),
        italic: e.isActive('italic'),
        underline: e.isActive('underline'),
        strike: e.isActive('strike'),
        subscript: e.isActive('subscript'),
        superscript: e.isActive('superscript'),
        bulletList: e.isActive('bulletList'),
        orderedList: e.isActive('orderedList'),
        align: (e.isActive({ textAlign: 'center' })
          ? 'center'
          : e.isActive({ textAlign: 'right' })
            ? 'right'
            : e.isActive({ textAlign: 'justify' })
              ? 'justify'
              : 'left') as 'left' | 'center' | 'right' | 'justify',
        block: (e.isActive('heading', { level: 1 })
          ? 'h1'
          : e.isActive('heading', { level: 2 })
            ? 'h2'
            : e.isActive('heading', { level: 3 })
              ? 'h3'
              : 'p') as 'p' | 'h1' | 'h2' | 'h3',
      }
      setSelFmt((prev) => {
        for (const k of Object.keys(next) as (keyof typeof next)[]) {
          if (prev[k] !== next[k]) return next
        }
        return prev
      })
      // 缩进三角跟随当前段（标尺可拖改的回显）
      const pstyle = (e.getAttributes('paragraph').style as string) || ''
      const ti = parseFloat(pstyle.match(/text-indent:\s*([\d.]+)px/)?.[1] || '0') || 0
      const ml = parseFloat(pstyle.match(/margin-left:\s*([\d.]+)px/)?.[1] || '0') || 0
      setIndentState((s) => (s.first === ti && s.hang === ml ? s : { first: ti, hang: ml }))
      // 选区浮工具条 + 滚动跟随光标（WPS 借鉴）
      placeFloatBar()
      const wrap = editWrapRef.current
      if (wrap && e.isFocused) {
        try {
          const caret = e.view.coordsAtPos(e.state.selection.head)
          const wr = wrap.getBoundingClientRect()
          const k = zoomPct / 100
          if (caret.bottom > wr.bottom - 8) wrap.scrollTop += (caret.bottom - (wr.bottom - 8)) / k
          else if (caret.top < wr.top + 8) wrap.scrollTop -= (wr.top + 8 - caret.top) / k
        } catch {
          /* 事务中 head 越界属正常瞬时态，忽略 */
        }
      }
    }
    sync()
    e.on('selectionUpdate', sync)
    e.on('transaction', sync)
    return () => {
      e.off('selectionUpdate', sync)
      e.off('transaction', sync)
    }
    // editorInst 由 TiptapDocEditor 的 onReady 提供 —— 实例就绪（或卸载置空）时重新订阅
  }, [editorInst])

  /**
   * Ribbon 格式命令 → Tiptap 命令层。
   * 相对旧的 execCommand 方案：
   *   - 走 ProseMirror transaction，undo/redo 可靠、选区语义正确、浏览器间一致
   *   - 字号/字体是「精确值」而非 1-7 档离散映射
   *   - chain().focus() 自动处理选区，不需要手动存取 range
   */
  const execFmt = useCallback(
    (cmd: string, arg?: string) => {
      const e = editorRef.current
      if (!e) return
      const c = e.chain().focus()
      switch (cmd) {
        case 'bold': c.toggleBold().run(); break
        case 'italic': c.toggleItalic().run(); break
        case 'underline': c.toggleUnderline().run(); break
        case 'strikeThrough': c.toggleStrike().run(); break
        case 'subscript': c.toggleSubscript().run(); break
        case 'superscript': c.toggleSuperscript().run(); break
        case 'insertUnorderedList': c.toggleBulletList().run(); break
        case 'insertOrderedList': c.toggleOrderedList().run(); break
        case 'justifyLeft': c.setTextAlign('left').run(); break
        case 'justifyCenter': c.setTextAlign('center').run(); break
        case 'justifyRight': c.setTextAlign('right').run(); break
        case 'justifyFull': c.setTextAlign('justify').run(); break
        case 'foreColor': c.setColor(arg ?? '#000000').run(); break
        case 'hiliteColor': c.setHighlight({ color: arg ?? '#ffff00' }).run(); break
        case 'removeFormat': c.unsetAllMarks().clearNodes().run(); break
        case 'undo': c.undo().run(); break
        case 'redo': c.redo().run(); break
        case 'insertHorizontalRule': c.setHorizontalRule().run(); break
        case 'formatBlock': {
          // arg 形如 '<h2>' / '<p>' / '<blockquote>'
          const tag = (arg ?? '').replace(/[<>]/g, '').toLowerCase()
          if (tag === 'p') c.setParagraph().run()
          else if (/^h[1-6]$/.test(tag)) c.toggleHeading({ level: Number(tag[1]) as 1 | 2 | 3 | 4 | 5 | 6 }).run()
          else if (tag === 'blockquote') c.toggleBlockquote().run()
          break
        }
        default:
          // 未映射的命令：静默忽略（避免抛错打断用户操作）
          break
      }
      onEditInputRef.current()
    },
    [editorRef],
  )

  /** 字号（精确 pt）：直接走 TextStyle 的 fontSize 属性 → 产出 style="font-size:XXpt"
   *  → html2ooxml 认 style.font-size → w:sz 双写（WPS 中文字号：小初36/一号26/…/五号10.5/小五9） */
  const applyFontSize = useCallback(
    (pt: number) => {
      const e = editorRef.current
      if (!e) return
      e.chain().focus().setMark('textStyle', { fontSize: `${pt}pt` }).run()
      onEditInputRef.current()
    },
    [editorRef],
  )

  /** 字体族（Tiptap FontFamily 扩展） */
  const applyFontFamily = useCallback(
    (family: string) => {
      const e = editorRef.current
      if (!e) return
      if (!family) e.chain().focus().unsetFontFamily().run()
      else e.chain().focus().setFontFamily(family).run()
      onEditInputRef.current()
    },
    [editorRef],
  )

  /** 行距：段落级 line-height（走自定义扩展的 node attr） */
  const applyLineHeight = useCallback(
    (lh: string) => {
      const e = editorRef.current
      if (!e) return
      e.chain().focus().updateAttributes('paragraph', { lineHeight: lh || null }).run()
      onEditInputRef.current()
    },
    [editorRef],
  )

  // 缩放改变 layout（CSS zoom 影响元素尺寸）→ 页数估算跟着刷新
  useEffect(() => {
    if (!editorMode || tab !== 'preview') return
    const t = window.setTimeout(computeStats, 80)
    return () => window.clearTimeout(t)
  }, [zoomPct, computeStats, editorMode, tab])


  /** WPS 风格功能区页签。「开始」页在编辑器态直接驱动就地编辑面（execCommand）；
   *  紧凑态（属性面板）没有编辑面，保持 disabled。 */
  const NO_EDIT = '就地编辑在 Word 文档窗口中使用（双击节点打开）'
  /** 标题样式命令（formatBlock） */
  const blockFmt = (v: string) => execFmt('formatBlock', `<${v}>`)
  const ribbonTabs: RibbonTab[] = [
    {
      label: '文件',
      groups: [
        {
          header: '模板',
          items: [{ kind: 'btn', label: '导入模板', glyph: '📂', onClick: () => void importTemplate() }],
        },
        {
          header: '导出',
          items: [
            { kind: 'btn', label: '导出 Word', glyph: '📄', onClick: () => void exportDoc() },
            { kind: 'btn', label: '导出 PDF', glyph: '🖨', onClick: () => void exportPdf() },
          ],
        },
      ],
    },
    {
      label: '开始',
      groups: [
        {
          header: '字体',
          items: [
            {
              kind: 'select',
              value: '宋体',
              disabled: !editorMode,
              title: editorMode ? '字体' : NO_EDIT,
              // 88：够放最长的「微软雅黑」(4字=50) + padding(15) + 边框(2) + 箭头预留(24) ≈ 91…
              // 实测 Edge 箭头预留 24 时 4 字需 91，故取 92 保底
              width: 92,
              options: [
                { value: '', label: '默认字体' },
                { value: '宋体', label: '宋体' },
                { value: '黑体', label: '黑体' },
                { value: '仿宋', label: '仿宋' },
                { value: '楷体', label: '楷体' },
                { value: '微软雅黑', label: '微软雅黑' },
                { value: 'Times New Roman', label: 'Times New Roman' },
                { value: 'Arial', label: 'Arial' },
              ],
              onChange: (v) => applyFontFamily(v),
            },
            {
              kind: 'select',
              value: '10.5',
              disabled: !editorMode,
              title: editorMode ? '字号' : NO_EDIT,
              width: 72,
              options: [
                { value: '36', label: '小初' },
                { value: '26', label: '一号' },
                { value: '22', label: '二号' },
                { value: '18', label: '小二' },
                { value: '16', label: '三号' },
                { value: '14', label: '四号' },
                { value: '12', label: '小四' },
                { value: '10.5', label: '五号' },
                { value: '9', label: '小五' },
              ],
              onChange: (v) => applyFontSize(parseFloat(v)),
            },
            'sep',
            // WPS 字形组：B 粗 / I 斜 / U 下划线 / S 删除线 —— 用真实字形而非普通字母
            // active = 当前选区已应用该格式 → 渲染按下态（WPS 同步选区格式）
            { kind: 'btn', label: '', glyph: 'B', glyphStyle: { fontWeight: 800, fontFamily: 'Georgia, serif' }, active: selFmt.bold, disabled: !editorMode, title: editorMode ? '加粗' : NO_EDIT, onClick: () => execFmt('bold') },
            { kind: 'btn', label: '', glyph: 'I', glyphStyle: { fontStyle: 'italic', fontFamily: 'Georgia, serif', fontWeight: 600 }, active: selFmt.italic, disabled: !editorMode, title: editorMode ? '斜体' : NO_EDIT, onClick: () => execFmt('italic') },
            { kind: 'btn', label: '', glyph: 'U', glyphStyle: { textDecoration: 'underline', textUnderlineOffset: '3px', fontFamily: 'Georgia, serif', fontWeight: 600 }, active: selFmt.underline, disabled: !editorMode, title: editorMode ? '下划线' : NO_EDIT, onClick: () => execFmt('underline') },
            { kind: 'btn', label: '', glyph: 'S', glyphStyle: { textDecoration: 'line-through', fontFamily: 'Georgia, serif', fontWeight: 600 }, active: selFmt.strike, disabled: !editorMode, title: editorMode ? '删除线' : NO_EDIT, onClick: () => execFmt('strikeThrough') },
            { kind: 'btn', label: '', glyph: 'X₂', glyphStyle: { fontSize: 19, fontWeight: 600 }, active: selFmt.subscript, disabled: !editorMode, title: editorMode ? '下标' : NO_EDIT, onClick: () => execFmt('subscript') },
            { kind: 'btn', label: '', glyph: 'X²', glyphStyle: { fontSize: 19, fontWeight: 600 }, active: selFmt.superscript, disabled: !editorMode, title: editorMode ? '上标' : NO_EDIT, onClick: () => execFmt('superscript') },
            'sep',
            {
              kind: 'select',
              value: '',
              disabled: !editorMode,
              title: editorMode ? '文字颜色' : NO_EDIT,
              // 92 = 「文字颜色」四字(50) + 左右 padding(15) + 边框(2) + Edge 下拉箭头预留(24)
              width: 92,
              options: [
                { value: '', label: '文字颜色' },
                { value: '#C00000', label: '红色' },
                { value: '#185FA5', label: '蓝色' },
                { value: '#0F6E56', label: '绿色' },
                { value: '#000000', label: '黑色' },
              ],
              onChange: (v) => v && execFmt('foreColor', v),
            },
            {
              kind: 'select',
              value: '',
              disabled: !editorMode,
              title: editorMode ? '突出显示' : NO_EDIT,
              // 同上：四字 + 箭头预留意，防中文被裁成「突出示」
              width: 92,
              options: [
                { value: '', label: '突出显示' },
                { value: '#FFFF00', label: '黄色' },
                { value: '#C6E0B4', label: '绿色' },
                { value: '#F8CBAD', label: '橙色' },
              ],
              onChange: (v) => v && execFmt('hiliteColor', v),
            },
            { kind: 'btn', glyph: '✕', label: '清除', glyphStyle: { fontSize: 17 }, disabled: !editorMode, title: editorMode ? '清除选区格式' : NO_EDIT, onClick: () => execFmt('removeFormat') },
          ],
        },
        {
          header: '段落',
          items: [
            { kind: 'btn', label: '', glyph: '⯇', glyphStyle: { fontSize: 20 }, active: selFmt.align === 'left', disabled: !editorMode, title: editorMode ? '左对齐' : NO_EDIT, onClick: () => execFmt('justifyLeft') },
            { kind: 'btn', label: '', glyph: '≡', glyphStyle: { fontSize: 20 }, active: selFmt.align === 'center', disabled: !editorMode, title: editorMode ? '居中对齐' : NO_EDIT, onClick: () => execFmt('justifyCenter') },
            { kind: 'btn', label: '', glyph: '⯈', glyphStyle: { fontSize: 20 }, active: selFmt.align === 'right', disabled: !editorMode, title: editorMode ? '右对齐' : NO_EDIT, onClick: () => execFmt('justifyRight') },
            // 两端对齐用 ▤（填满的方块）与居中的 ≡（三短横）拉开辨识度 ——
            // 原用 ≣ 与 ≡ 几乎无法区分（2026-09-28 截图复核）
            { kind: 'btn', label: '', glyph: '▤', glyphStyle: { fontSize: 20 }, active: selFmt.align === 'justify', disabled: !editorMode, title: editorMode ? '两端对齐' : NO_EDIT, onClick: () => execFmt('justifyFull') },
            'sep',
            // 列表用「点行 / 数字行」的排版符号，避免和「两端对齐」的横线语义撞车
            { kind: 'btn', label: '', glyph: '⁝', glyphStyle: { fontSize: 26, lineHeight: '22px' }, active: selFmt.bulletList, disabled: !editorMode, title: editorMode ? '无序列表' : NO_EDIT, onClick: () => execFmt('insertUnorderedList') },
            { kind: 'btn', label: '', glyph: '№', glyphStyle: { fontSize: 15, fontWeight: 600 }, active: selFmt.orderedList, disabled: !editorMode, title: editorMode ? '有序列表' : NO_EDIT, onClick: () => execFmt('insertOrderedList') },
            'sep',
            {
              kind: 'select',
              value: '',
              disabled: !editorMode,
              title: editorMode ? '行距' : NO_EDIT,
              width: 78,
              options: [
                { value: '', label: '行距' },
                { value: '1', label: '单倍' },
                { value: '1.25', label: '1.25 倍' },
                { value: '1.5', label: '1.5 倍' },
                { value: '1.75', label: '1.75 倍' },
                { value: '2', label: '2 倍' },
              ],
              onChange: (v) => applyLineHeight(v),
            },
          ],
        },
        {
          header: '样式',
          items: [
            // WPS 样式库是「缩略块预览」；这里用等宽缩略字形替代，比纯文字更像样式库
            { kind: 'btn', label: '正文', glyph: '¶', glyphStyle: { fontSize: 17 }, active: selFmt.block === 'p', disabled: !editorMode, title: editorMode ? '正文段落' : NO_EDIT, onClick: () => blockFmt('p') },
            { kind: 'btn', label: '标题 1', glyph: 'H₁', glyphStyle: { fontSize: 16, fontWeight: 700 }, active: selFmt.block === 'h1', disabled: !editorMode, title: editorMode ? '标题 1' : NO_EDIT, onClick: () => blockFmt('h1') },
            { kind: 'btn', label: '标题 2', glyph: 'H₂', glyphStyle: { fontSize: 15, fontWeight: 700 }, active: selFmt.block === 'h2', disabled: !editorMode, title: editorMode ? '标题 2' : NO_EDIT, onClick: () => blockFmt('h2') },
            { kind: 'btn', label: '标题 3', glyph: 'H₃', glyphStyle: { fontSize: 14, fontWeight: 700 }, active: selFmt.block === 'h3', disabled: !editorMode, title: editorMode ? '标题 3' : NO_EDIT, onClick: () => blockFmt('h3') },
          ],
        },
        // 「撤销 / 重做」不放在「开始」页签 —— WPS 里它们在行1 的快速访问栏。
        // 放这里会让「开始」页签多出 141px，导致 1180 视口下「样式」组被挤出屏幕
        // （2026-09-28 实测 scrollW 1327 > clientW 1180）。见下方 Ribbon quick 属性。
      ],
    },
    {
      label: '插入',
      groups: [
        {
          header: '域（绑定上游字段）',
          items: upstream.fields.slice(0, 12).map((f) => ({
            kind: 'btn' as const,
            label: f.name,
            glyph: '｛｝',
            // 就地编辑的对象是绑定产物（标记已被数据替换），再插标记没有意义 —— 在模板里加标记后重新导入
            disabled: editorMode,
            title: editorMode ? '就地编辑基于绑定产物；要插入域标记请在模板文档里加好后重新导入' : '插入 ｛｛field:' + f.name + '｝｝ 绑定标记',
            onClick: () => {
              const cur = String(map[f.id]?.from === 'manual' ? map[f.id]?.value ?? '' : '')
              setBinding(f.id, { from: 'manual', value: '｛｛field:' + f.name + '｝｝' + cur })
            },
          })),
        },
        {
          header: '对象',
          items: [
            {
              kind: 'btn',
              label: '图片',
              glyph: '🖼',
              disabled: !editorMode,
              title: editorMode ? '插入本地图片（嵌入型，插入后点图片可设环绕）' : NO_EDIT,
              onClick: insertImage,
            },
          ],
        },
      ],
    },
    {
      label: '审阅',
      groups: [
        {
          header: '溯源',
          items: [
            { kind: 'btn', label: '域取值来源', onClick: () => setTab('bind') },
            { kind: 'btn', label: '绑定告警', onClick: () => setTab('bind') },
          ],
        },
      ],
    },
    {
      label: '视图',
      groups: [
        {
          header: '页签',
          items: [
            { kind: 'btn', label: tab === 'bind' ? '✓ 绑定与产出' : '　绑定与产出', onClick: () => setTab('bind') },
            { kind: 'btn', label: tab === 'preview' ? '✓ 文档编辑' : '　文档编辑', onClick: () => setTab('preview') },
          ],
        },
      ],
    },
  ]

  const showHero = !!editorMode && !templateId

  /**
   * 影子表头的 HTML 来源：直接读**当前编辑区里**对应表格的首行。
   *
   * 为什么不缓存一份 state：表头里的文字会被就地编辑改动，
   * 每次渲染都从 DOM 现读，影子就跟真表头始终一致（改一个字两处同时变）。
   * 读不到（表格序号越界 / DOM 还没渲染）就返回空串 → 那个影子头自动消失，
   * 不会留下半截脏行。
   *
   * ★ 必须连**每列的实测宽度**一起克隆（用 colgroup + 百分比）：
   *   表头克隆放在独立表里，若只给 width:100% + table-layout:fixed，
   *   各列会均分 —— 与真表（内容驱动的列宽）错位，影子看着像"另一张表"。
   *   实测：真表列宽 [40,128,80,96,220]，均分后变成每列 ~113，文字四处飘。
   */
  function shadowHeadHtml(tblNo: number): string {
    const edit = editWrapRef.current?.querySelector<HTMLElement>('.wrp-edit')
    if (!edit) return ''
    const tbl = [...edit.querySelectorAll<HTMLTableElement>('table')][tblNo - 1]
    if (!tbl) return ''
    const tbody = tbl.querySelector<HTMLTableSectionElement>('tbody')
    const rows = tbody ? [...tbody.querySelectorAll<HTMLElement>(':scope > tr')] : []
    const head = rows[0]
    if (!head) return ''
    // 真表总宽作分母，逐列换算百分比 → 克隆表各列与真表严格对齐
    const total = tbl.offsetWidth || 1
    const cols = [...head.querySelectorAll<HTMLElement>(':scope > th, :scope > td')]
    const colgroup = cols.map((c) => `<col style="width:${((c.offsetWidth / total) * 100).toFixed(3)}%" />`).join('')
    return `<table><colgroup>${colgroup}</colgroup><tbody>${head.outerHTML}</tbody></table>`
  }

  /** 中段内容：放大态未绑模板 → 整页空态三入口（与 Excel 的 wb-hero 同款）；否则常规内容 */
  const bodyContent = showHero ? (
    <div className="wb-hero">
      <div className="wb-hero-title">打开一个 Word 文档</div>
      {pickHint && <div className="wrp-warn">{pickHint}</div>}
      <div className="wb-hero-actions">
        <button className="bar-btn primary" onClick={importTemplate} disabled={!!busy}>
          📂 导入模板…
        </button>
        <button className="bar-btn" onClick={() => void createBlankDocx()} disabled={!!busy}>
          ＋ 创建空白文档
        </button>
      </div>
      {busy && <div className="wrp-busy">{busy}</div>}
      {libs.length > 0 && (
        <div className="wb-hero-history">
          <div className="wb-hero-h">历史文件</div>
          {libs.map((l) => (
            <button
              key={l.id}
              className="wb-hero-item"
              onClick={() => onSetNodeConfig(table.id, { docxTemplateId: l.id, docxTemplateName: l.name })}
            >
              {l.name}
            </button>
          ))}
        </div>
      )}
      <div className="wb-hero-hint">导入、从历史选择或创建空白文档后，这里就是文档编辑器：绑定域、预览、导出。</div>
    </div>
  ) : (
    <>
      {/* 编辑器态去头去尾（用户拍板：这些入口 Ribbon/标题栏都有了），紧凑态保留原布局 */}
      {!editorMode && (
        <div className="wrp-head">
          <span className="wrp-title">Word 报告</span>
          <div className="wrp-tabs">
            <button className={tab === 'bind' ? 'on' : ''} onClick={() => setTab('bind')}>
              绑定
            </button>
            <button className={tab === 'preview' ? 'on' : ''} onClick={() => setTab('preview')}>
              预览
            </button>
          </div>
        </div>
      )}

      {err && <div className="wrp-err">{err}</div>}
      {editorMode && pickHint && <div className="wrp-warn">{pickHint}</div>}
      {tplWarn && (
        <div className="wrp-warn">
          <span>{tplWarn}</span>
          <button className="wrp-warn-x" title="知道了" onClick={() => setTplWarn('')}>
            ×
          </button>
        </div>
      )}

      {!editorMode && (
        <div className="wrp-row">
          {templateId ? (
            <>
              <span className="wrp-tpl" title={config.docxTemplateName}>
                {config.docxTemplateName ?? '（未命名模板）'}
              </span>
              <button className="bar-btn" onClick={importTemplate} disabled={!!busy}>
                换模板
              </button>
            </>
          ) : (
            <button className="bar-btn primary" onClick={importTemplate} disabled={!!busy}>
              导入 Word 模板
            </button>
          )}
          {libs.length > 1 && (
            <select
              value={templateId ?? ''}
              onChange={(e) => {
                const hit = libs.find((l) => l.id === e.target.value)
                if (hit) onSetNodeConfig(table.id, { docxTemplateId: hit.id, docxTemplateName: hit.name })
              }}
            >
              {libs.map((l) => (
                <option key={l.id} value={l.id}>
                  {l.name}
                </option>
              ))}
            </select>
          )}
        </div>
      )}

      {!editorMode && !templateId && (
        <p className="hint">
          导入设计院给的 .docx，用 <code>{'{{field:字段名}}'}</code> / <code>{'{{rich:名称}}'}</code> /
          <code>{'{{each:名称}}'}</code> / <code>{'{{if:名称}}'}</code> 标记要连接数据的位置，导入时会自动转成 Word
          内容控件。不加标记也可以直接编辑/导出。
        </p>
      )}

      {busy && <div className="wrp-busy">{busy}</div>}

      {tab === 'preview' ? (
        editorMode ? (
          /* ── WPS 式就地编辑面：标尺 + A4 分页纸 + 页边距，contenteditable 直改文档 ──
             editzone 不滚动（editbar/环绕浮层悬浮锚定它）；滚动发生在 wrp-edit-wrap 内部。
             分页：正文层透明，wrp-pages 背景层按页高铺白纸+页隙 —— 内容越过分页线即视觉换页 */
          <div className="wrp-editzone" ref={editZoneRef} onClick={() => setCtxMenu(null)}>
            {(editedBytes || editRestored) && (
              <div className="wrp-editbar">
                <span>编辑稿模式 · 数据重绑后将重新生成</span>
                <button className="bar-btn" onClick={discardEdit} title="放弃手工修改，回到绑定产物">
                  放弃
                </button>
              </div>
            )}
            <div className="wrp-doc-area">
              {/* 缩略图导航窗格（WPS 借鉴）：左列页，点页 scrollIntoView 跳转变焦，当前页高亮 */}
              {viewMode !== 'draft' && (
                <div className="wrp-thumbs">
                  {pageModel.pages.map((p, i) => (
                    <button
                      key={i}
                      type="button"
                      className={`wrp-thumb${curPage === i + 1 ? ' is-cur' : ''}`}
                      style={{ height: `${Math.max(30, (p.height / PAGE_H) * 132)}px` }}
                      onClick={() => onThumbClick(p.top)}
                      title={`第 ${i + 1} 页`}
                    >
                      <span>{i + 1}</span>
                    </button>
                  ))}
                </div>
              )}
              <div className="wrp-edit-wrap" ref={editWrapRef} data-view={viewMode} onScroll={onWrapScroll} onContextMenu={onEditContextMenu}>
                {/* ── 画布层：唯一挂 zoom 的地方 ──
                    zoom 只能挂在这一层：层间偏移（top:28 等）与层内尺寸处在同一缩放坐标系，
                    任何 zoom 下各层起点互相对齐。曾把 zoom 分别挂在五层上、
                    层间偏移却是不缩放的 px 常量 → zoom≠100% 时纸与正文错开 (1123+8)·zoom px
                    （scripts/_diag_zoomfit.mjs 缩放矩阵实测：50% 时 -572px、150% 时 -1696px）。 */}
                <div className="wrp-canvas" style={{ zoom: zoomPct / 100 }}>
                {/* 垂直标尺（WPS 借鉴）：上/下页边距块 + 1cm 刻度。
                    绝对定位在纸左侧留白槽（left = 50% − 397 − 24），不占文档流 ——
                    曾是 column flex 的流式子元素，独占 1123px 整行，把水平标尺/正文
                    压到一屏之外、纸与正文分离 1134px（用户看到的是一张没有字的大白纸）。 */}
                {viewMode !== 'draft' && (
                  <div className="wrp-ruler wrp-ruler-v" aria-hidden>
                    <i className="wrp-ruler-mar" style={{ height: sectMargin.t }} title="上页边距" />
                    <div className="wrp-ruler-scale-v">
                      {Array.from({ length: 28 }, (_, i) => (
                        <i key={`tv${i}`} className={`wrp-ruler-tick-v${(i + 1) % 2 === 0 ? ' is-major' : ''}`} style={{ top: `${(i + 1) * RULER_PX_PER_CM}px` }} />
                      ))}
                    </div>
                    <i className="wrp-ruler-mar" style={{ height: sectMargin.b }} title="下页边距" />
                  </div>
                )}
                {/* 水平标尺（可拖，WPS 借鉴）：左/右页边距块 + 刻度 + 制表位 + 首行/悬挂缩进三角 */}
                <div className="wrp-ruler wrp-ruler-h">
                  <i
                    className="wrp-ruler-mar wrp-mar-l"
                    style={{ width: sectMargin.l }}
                    title={`左页边距 ${(sectMargin.l / RULER_PX_PER_CM).toFixed(2)}cm`}
                    onPointerDown={onMarginDown('l')}
                  />
                  {/* 刻度区不再设 marginLeft/marginRight：左/右页边距已由 .wrp-mar-l / .wrp-mar-r
                      块承担，scale 用 flex:1 自动填满中间 554px。否则会与页边距块叠加 120px，
                      导致 scale 被压到 434px、14cm 刻度溢出被裁，标尺与文本带错位（A2 回归）。 */}
                  <div className="wrp-ruler-scale" onClick={onRulerScaleClick}>
                    {Array.from({ length: RULER_CM }, (_, i) => (
                      <i key={`t${i}`} className={`wrp-ruler-tick${(i + 1) % 2 === 0 ? ' is-major' : ''}`} style={{ left: `${(i + 1) * RULER_PX_PER_CM}px` }} />
                    ))}
                    {Array.from({ length: Math.floor(RULER_CM / 2) }, (_, i) => (
                      <b key={`n${i}`} style={{ left: `${(i + 1) * 2 * RULER_PX_PER_CM}px` }}>
                        {(i + 1) * 2}
                      </b>
                    ))}
                    {tabStops.map((x, i) => (
                      <i key={`tab${i}`} className="wrp-ruler-tab" style={{ left: `${x}px` }} />
                    ))}
                    <i className="wrp-indent wrp-indent-first" style={{ left: `${indentState.first}px` }} title="首行缩进" onPointerDown={onIndentDown('first')} />
                    <i className="wrp-indent wrp-indent-hang" style={{ left: `${indentState.hang}px` }} title="悬挂缩进" onPointerDown={onIndentDown('hang')} />
                  </div>
                  <i
                    className="wrp-ruler-mar wrp-mar-r"
                    style={{ width: sectMargin.r }}
                    title={`右页边距 ${(sectMargin.r / RULER_PX_PER_CM).toFixed(2)}cm`}
                    onPointerDown={onMarginDown('r')}
                  />
                </div>
              {/* ── 一屏多页背景层（WPS 观感）──
                  绝对定位在内容之下，按 A4 周期铺白纸；内容层 .wrp-edit 透明浮在其上。
                  这样「内容少」时纸也铺满可视区，「内容多」时自动长出下一页。
                  纯视觉层：不参与导出（导出走 OOXML 连续流）。 */}
              <div className="wrp-pages" aria-hidden>
                {Array.from({ length: bgPageCount }, (_, i) => (
                  <div key={i} className="wrp-page">
                    <span className="wrp-page-no">{i + 1}</span>
                  </div>
                ))}
              </div>
              {/* ── 影子表头（跨页时每页页顶重复表头行）──
                  纯视觉层：克隆表格首行（连列宽）画在页顶，不写回文档模型。
                  对应导出件里的 <w:tblHeader/>（见 html2ooxmll tableXml / binder withRowProps）。
                  夹在 .wrp-pages(0) 与 .wrp-edit(1) 之间，z-index:2（高于内容，低于遮断条）。 */}
              {shadowHeaders.length > 0 && (
                <div className="wrp-shadow-heads" aria-hidden>
                  {shadowHeaders.map((s, i) => {
                    const html = shadowHeadHtml(s.tbl)
                    if (!html) return null
                    return (
                      <div
                        key={`${s.tbl}-${s.page}-${i}`}
                        className="wrp-shadow-head"
                        style={{ top: `${(s.page - 1) * period + PAGE_PAD_TOP}px` }}
                        // 内容来自本产品自己的编辑器 DOM（clone 出来的真表头行），非外部输入
                        dangerouslySetInnerHTML={{ __html: html }}
                      />
                    )
                  })}
                </div>
              )}
              {/* ── 页隙遮断条（切页时把穿过页隙的内容"切断"）──
                  ★ 为什么必须有：内容层是一整条连续流（表格/段落不会被浏览器分页），
                    且 .wrp-edit 在 .wrp-pages 之上（z-index 1 > 0）。
                    于是表格的单元格边框会**连续穿过页隙**，视觉上把两页"缝合"在一起
                    （用户 2026-09-28 截图：「为什么表格两页连在一起」）。
                  这**不是表格结构问题**，也不需要按页拆表（用户拍板：展示分页、不真拆表）——
                  只要在每处页隙上盖一条与页隙等高的不透明条，把穿过来的内容遮住即可。
                  条高 = PAGE_GAP，位置 = 每页底部；z-index 高于内容层。
                  纯视觉层：不写回文档模型，导出仍是原来的连续表。 */}
              <div className="wrp-cutters" aria-hidden>
                {Array.from({ length: cutterCount }, (_, i) => (
                  <div
                    key={i}
                    className="wrp-cutter"
                    style={{ top: `${(i + 1) * period - PAGE_GAP}px` }}
                  />
                ))}
              </div>
              {/* 表格行跨页避让规则（注入式）：
                  挂在 .wrp-edit 作用域，ProseMirror 重建 tr 也不会失效。
                  data-wrp-avoid 仅作标记，便于诊断脚本区分这条规则与其他样式。
                  见 avoidTableRowSplit / avoidCss 的注释。
                  注：不再注入「去边框」的修饰样式（splitCss 恒为空）——
                  页隙由 .wrp-cutters 统一遮断，逐行去边框反而会把承载留白的那一行
                  变成没有表格线的孤儿文字（用户 2026-09-28：「为什么没有跨页断行」）。 */}
              {avoidCss && <style data-wrp-avoid="">{avoidCss}</style>}
              {/* 编辑内容层：只负责排版与页边距，背景交由 .wrp-pages（透明浮层）
                  编辑内核 = Tiptap（ProseMirror）：格式命令走 transaction，undo/选区可靠 */}
              <div
                ref={editPaneRef}
                className="wrp-edit"
                style={{
                  minHeight: `${paperMinH}px`,
                  padding: `${sectMargin.t}px ${sectMargin.r}px ${sectMargin.b}px ${sectMargin.l}px`,
                }}
                onMouseDown={onEditMouseDown}
              >
                <TiptapDocEditor
                  value={editorHtml}
                  onChange={(html) => {
                    pendingHtmlRef.current = html
                    onEditInputRef.current()
                  }}
                  onBlur={() => void flushEdit()}
                  editorRef={editorRef}
                  onReady={setEditorInst}
                />
              </div>
                </div>{/* /.wrp-canvas（唯一 zoom 承载层） */}
            </div>
            </div>
            {/* 选区浮工具条 + 右键上下文菜单（WPS 借鉴） */}
            {floatBar && editorMode && (
              <div className="wrp-float-bar" style={{ left: floatBar.x, top: floatBar.y }}>
                <button className={selFmt.bold ? 'is-on' : ''} onClick={() => execFmt('bold')} title="加粗"><b>B</b></button>
                <button className={selFmt.italic ? 'is-on' : ''} onClick={() => execFmt('italic')} title="斜体"><i>I</i></button>
                <button className={selFmt.underline ? 'is-on' : ''} onClick={() => execFmt('underline')} title="下划线"><u>U</u></button>
                <button onClick={() => execFmt('strikeThrough')} title="删除线"><s>S</s></button>
                <span className="wrp-fb-sep" />
                <button onClick={() => setTab('bind')} title="去绑定页做 AI 局部改写">AI 改写</button>
              </div>
            )}
            {ctxMenu && editorMode && (
              <div className="wrp-ctx-menu" style={{ left: ctxMenu.x, top: ctxMenu.y }} onMouseDown={(e) => e.stopPropagation()}>
                <button onClick={() => ctxRun(() => { editorInst?.chain().focus().run(); document.execCommand('cut') })}>剪切</button>
                <button onClick={() => ctxRun(() => { editorInst?.chain().focus().run(); document.execCommand('copy') })}>复制</button>
                <button onClick={() => ctxRun(() => { editorInst?.chain().focus().run(); document.execCommand('paste') })}>粘贴</button>
                <span className="wrp-ctx-sep" />
                <button onClick={() => setTab('bind')}>字体 / 段落…</button>
                <button onClick={() => ctxRun(() => void insertImage())}>插入图片</button>
              </div>
            )}
            {/* 图片环绕浮层：挂不滚动的 editzone（editbar 同款锚），滚动即收起 */}
            {activeImg && wrapPop && (
              <div className="wrp-wrap-pop" style={{ left: wrapPop.x, top: wrapPop.y }}>
                <span className="wrp-wrap-pop-t">环绕</span>
                <button title="嵌入型" onClick={() => applyWrap('inline')}>嵌</button>
                <button title="四周型（左）" onClick={() => applyWrap('square', 'left')}>左绕</button>
                <button title="四周型（右）" onClick={() => applyWrap('square', 'right')}>右绕</button>
                <button title="上下型" onClick={() => applyWrap('topBottom')}>上下</button>
              </div>
            )}
            <input
              ref={imgFileRef}
              type="file"
              accept="image/*"
              hidden
              onChange={(e) => {
                const f = e.target.files?.[0]
                if (f) void onImgPicked(f)
                e.target.value = ''
              }}
            />
          </div>
        ) : (
          <div className="wrp-preview-wrap">
            <DocxPreview data={preview} className="wrp-preview" />
            <button className="bar-btn wrp-zoom-btn" onClick={() => setZoom(true)} disabled={!preview}>
              放大预览
            </button>
          </div>
        )
      ) : (
        <>
          {upstream.rows.length > 0 && (
            <div className="wrp-meta">
              上游汇入 {upstream.fields.length} 个字段 / {upstream.rows.length} 行
            </div>
          )}

          {/*
            取值溯源：送审时被问「这个数哪来的」，这里能逐个域答出来。
            与 Excel 侧对称 —— 复用的是同一条 ResolveTrace。
          */}
          {built.traces.length > 0 && (
            <details className="wrp-trace">
              <summary>
                取值溯源（{built.traces.length} 个域
                {built.traces.some((t) => !t.ok) ? `，${built.traces.filter((t) => !t.ok).length} 个未取到` : ''}）
              </summary>
              <div className="wrp-trace-body">
                {built.traces.map((t: ResolveTrace) => (
                  <div key={t.key} className={`wrp-trace-row${t.ok ? '' : ' bad'}`}>
                    <span className="wrp-trace-key" title={t.key}>
                      {t.key}
                    </span>
                    <span className="wrp-trace-from">{SOURCE_LABEL[t.from] ?? t.from}</span>
                    <span className="wrp-trace-urn" title={t.urn}>
                      {t.urn}
                    </span>
                    <span className="wrp-trace-val">{fmtTraceValue(t.value)}</span>
                  </div>
                ))}
              </div>
            </details>
          )}

          {parseWarn.length > 0 && <div className="wrp-warn">{parseWarn.join('；')}</div>}
          {bindWarn.length > 0 && <div className="wrp-warn">{bindWarn.join('；')}</div>}

          <div className="wrp-list">
            {bindings.length === 0 && templateId && <div className="wrp-empty">该模板没有可绑定的控件</div>}
            {bindings.map((b) => {
              const spec = map[b.tag]
              const from = spec?.from ?? (b.kind === 'each' ? 'rows' : 'auto')
              return (
                <div
                  key={b.tag}
                  className={`wrp-item${editorMode && activeTag === b.tag ? ' wrp-item--active' : ''}`}
                  onClick={editorMode ? () => setActiveTag((cur) => (cur === b.tag ? '' : b.tag)) : undefined}
                >
                  <div className="wrp-item-head">
                    <span className={`wrp-kind ${KIND_CLS[b.kind]}`}>{KIND_LABEL[b.kind]}</span>
                    <span className="wrp-name" title={b.tag}>
                      {b.path}
                    </span>
                    <span className="wrp-resolved">{built.resolved[b.tag] ?? ''}</span>
                  </div>

                  {b.kind === 'each' ? (
                    <div className="wrp-each">
                      <div className="wrp-ctl">取上游全部行（{upstream.rows.length} 行）</div>
                      {b.placeholders.length > 0 && (
                        <div className="wrp-cols">
                          {b.placeholders.map((ph) => {
                            const col = map[b.tag]?.columns ?? {}
                            const auto = upstream.fields.find((f) => f.name === ph || f.id === ph)
                            return (
                              <div key={ph} className="wrp-col">
                                <span className="wrp-col-name" title={ph}>
                                  {ph}
                                </span>
                                <select
                                  value={col[ph] ?? ''}
                                  onChange={(e) => {
                                    const next = { ...col }
                                    if (e.target.value) next[ph] = e.target.value
                                    else delete next[ph]
                                    setBinding(b.tag, { from: 'rows', columns: next })
                                  }}
                                >
                                  <option value="">{auto ? `自动：${auto.name}` : '（未匹配）'}</option>
                                  {upstream.fields.map((f) => (
                                    <option key={f.id} value={f.id}>
                                      {f.name}
                                    </option>
                                  ))}
                                </select>
                              </div>
                            )
                          })}
                        </div>
                      )}
                    </div>
                  ) : b.kind === 'rich' ? (
                    <div className="wrp-rich-wrap">
                      <RichEditor
                        value={spec?.value ?? ''}
                        placeholder="输入正文，支持加粗、列表、标题、颜色；可写 {{字段名}} 引用上游数据"
                        onChange={(html) => setBinding(b.tag, { from: 'manual', value: html })}
                      />
                      {upstream.fields.length > 0 && (
                        <select
                          className="wrp-insert"
                          value=""
                          onChange={(e) => {
                            if (!e.target.value) return
                            const f = upstream.fields.find((x) => x.id === e.target.value)
                            if (f) setBinding(b.tag, { from: 'manual', value: (spec?.value ?? '') + `{{${f.name}}}` })
                            e.target.value = ''
                          }}
                        >
                          <option value="">插入字段…</option>
                          {upstream.fields.map((f) => (
                            <option key={f.id} value={f.id}>
                              {f.name}
                            </option>
                          ))}
                        </select>
                      )}
                    </div>
                  ) : b.kind === 'if' ? (
                    <div className="wrp-ctl">
                      <select
                        value={from}
                        onChange={(e) => {
                          const v = e.target.value
                          setBinding(b.tag, v === 'auto' ? { from: 'manual', value: '' } : { from: v as DocxBinding['from'] })
                        }}
                      >
                        <option value="auto">按同名上游字段判断</option>
                        <option value="field">指定字段</option>
                        <option value="manual">手工开关</option>
                      </select>
                      {from === 'field' && (
                        <select
                          value={spec?.fieldId ?? ''}
                          onChange={(e) => setBinding(b.tag, { from: 'field', fieldId: e.target.value })}
                        >
                          <option value="">（选字段）</option>
                          {upstream.fields.map((f) => (
                            <option key={f.id} value={f.id}>
                              {f.name}
                            </option>
                          ))}
                        </select>
                      )}
                      {from === 'manual' && (
                        <select
                          value={spec?.value === 'true' ? 'true' : 'false'}
                          onChange={(e) => setBinding(b.tag, { from: 'manual', value: e.target.value })}
                        >
                          <option value="true">显示这段</option>
                          <option value="false">隐藏这段</option>
                        </select>
                      )}
                    </div>
                  ) : b.kind === 'image' ? (
                    <div className="wrp-img">
                      <div className="wrp-img-row">
                        <select
                          value={from}
                          onChange={(e) => {
                            const v = e.target.value
                            setBinding(
                              b.tag,
                              v === 'auto' ? { from: 'manual', value: '' } : { from: v as DocxBinding['from'] },
                            )
                          }}
                        >
                          <option value="manual">本地选图</option>
                          <option value="field">取上游字段</option>
                        </select>
                        {from === 'field' && (
                          <select
                            value={spec?.fieldId ?? ''}
                            onChange={(e) => setBinding(b.tag, { from: 'field', fieldId: e.target.value })}
                          >
                            <option value="">（选字段，值需为 data: URI）</option>
                            {upstream.fields.map((f) => (
                              <option key={f.id} value={f.id}>
                                {f.name}
                              </option>
                            ))}
                          </select>
                        )}
                      </div>

                      {from === 'manual' && (
                        <div className="wrp-img-row">
                          <label className="bar-btn wrp-img-pick">
                            选择图片…
                            <input
                              type="file"
                              accept="image/*"
                              hidden
                              onChange={(e) => {
                                const f = e.target.files?.[0]
                                if (f) void pickImage(b.tag, f)
                                e.target.value = ''
                              }}
                            />
                          </label>
                          {spec?.value && (
                            <button
                              className="bar-btn"
                              onClick={() => setBinding(b.tag, { from: 'manual', value: '' })}
                            >
                              清除
                            </button>
                          )}
                        </div>
                      )}

                      {/* 缩略图：只认 data: URI，外链在引擎侧会降级成文字，这里也如实显示 */}
                      {/^data:image\//i.test(String(spec?.value ?? '')) ? (
                        <div className="wrp-img-thumb">
                          <img src={String(spec?.value)} alt="" />
                          <span>{Math.round((String(spec?.value).length * 0.75) / 1024)} KB</span>
                        </div>
                      ) : (
                        spec?.value && (
                          <div className="wrp-img-thumb">
                            <span>非 data: URI，导出时会降级为文字占位</span>
                          </div>
                        )
                      )}

                      {imgWarn[b.tag] && <div className="wrp-warn">{imgWarn[b.tag]}</div>}
                    </div>
                  ) : (
                    <div className="wrp-ctl">
                      <select
                        value={from}
                        onChange={(e) => {
                          const v = e.target.value
                          setBinding(b.tag, v === 'auto' ? { from: 'manual', value: '' } : { from: v as DocxBinding['from'] })
                        }}
                      >
                        <option value="auto">按字段名自动匹配</option>
                        <option value="field">指定字段</option>
                        <option value="manual">手工填写</option>
                      </select>
                      {from === 'field' && (
                        <>
                          <select
                            value={spec?.fieldId ?? ''}
                            onChange={(e) => setBinding(b.tag, { from: 'field', fieldId: e.target.value })}
                          >
                            <option value="">（选字段）</option>
                            {upstream.fields.map((f) => (
                              <option key={f.id} value={f.id}>
                                {f.name}
                              </option>
                            ))}
                          </select>
                          <select
                            value={spec?.agg ?? 'first'}
                            onChange={(e) =>
                              setBinding(b.tag, { agg: e.target.value as DocxBinding['agg'] })
                            }
                          >
                            {AGGS.map((a) => (
                              <option key={a.v} value={a.v}>
                                {a.t}
                              </option>
                            ))}
                          </select>
                        </>
                      )}
                      {from === 'manual' && (
                        <input
                          className="wrp-val"
                          value={spec?.value ?? ''}
                          placeholder="取值"
                          onChange={(e) => setBinding(b.tag, { from: 'manual', value: e.target.value })}
                        />
                      )}
                    </div>
                  )}

                  {b.sample && <div className="wrp-sample">模板原文：{b.sample}</div>}
                </div>
              )
            })}
          </div>
        </>
      )}
    </>
  )

  return (
    <div className={`wrp${editorMode ? ' wrp--editor' : ''}`}>
      {/* ── WPS 风格功能区（与 Excel 编辑器共用 Ribbon 壳，样式一致）── */}
      <Ribbon
        tabs={ribbonTabs}
        initialTab="开始"
        collapsedByDefault={compact}
        /* 行1 快速访问栏：撤销 / 重做方形按钮（WPS 同位；放这里才不挤占「开始」页签宽度） */
        quick={
          <>
            <button
              className="rb-quickbtn"
              title={editorMode ? '撤销' : NO_EDIT}
              disabled={!editorMode}
              onClick={() => execFmt('undo')}
            >
              ↶
            </button>
            <button
              className="rb-quickbtn"
              title={editorMode ? '重做' : NO_EDIT}
              disabled={!editorMode}
              onClick={() => execFmt('redo')}
            >
              ↷
            </button>
          </>
        }
      />

      {/* 放大态：中段包进 wrp-main（左右让位，与 Excel 的 wb-main 同构）；紧凑态（属性面板）原样流式 */}
      {editorMode ? (
        <div className="wrp-main">
          <div
            className="wrp-col"
            style={{
              marginLeft: maximized && refPanelOpen && panelPos ? 230 : 0,
              marginRight: maximized && rightOpen && rightPos ? 230 : 0,
            }}
          >
            {bodyContent}
          </div>
        </div>
      ) : (
        bodyContent
      )}

      {/* 底栏：最大化时随左右栏让位 —— 按钮不被引用栏盖住（2026-09-27 用户反馈） */}
      <div
        className="wrp-foot"
        style={{
          marginLeft: maximized && refPanelOpen && panelPos ? 230 : undefined,
          marginRight: maximized && rightOpen && rightPos ? 230 : undefined,
        }}
      >
        {/* 备忘 · 左右两栏开关（放大态才渲染）—— 与 Excel 底栏同款「两侧放置」 */}
        {editorMode && (
          <button
            className="bar-btn"
            title={refPanelOpen ? '折叠节点数据引用' : '显示节点数据引用'}
            style={
              refPanelOpen
                ? { background: 'var(--ed-bg3)', color: 'var(--ed-accent)', borderColor: 'var(--ed-accent)', fontWeight: 600 }
                : undefined
            }
            onClick={() => setRefPanelOpen((v) => !v)}
          >
            ☰ 数据引用
          </button>
        )}
        <label className="wrp-check">
          <input
            type="checkbox"
            checked={config.docxStripControls !== false}
            onChange={(e) => onSetNodeConfig(table.id, { docxStripControls: e.target.checked })}
          />
          导出时去掉控件框（交付稿）
        </label>
        {/* 编辑器态：导出/在线编辑按钮已并入 Ribbon（文件页签 + 就地编辑面），底栏只留编辑稿状态 */}
        {editorMode && editedBytes && (
          <button className="bar-btn" onClick={discardEdit} title="放弃手工修改，回到绑定产物">
            放弃编辑稿
          </button>
        )}
        {/* WPS 式状态栏：页码 / 字数 / 视图模式 / 缩放（编辑器态文档页） */}
        {editorMode && tab === 'preview' && (
          <span className="wrp-status">
            <span className="wrp-status-item">
              第 {curPage} 页，共 {docStats.pages} 页
            </span>
            <span className="wrp-status-item">{docStats.words} 字</span>
            {/* 视图模式切换：页面 / Web版式 / 草稿（WPS 借鉴） */}
            <span className="wrp-status-view">
              {(['page', 'web', 'draft'] as const).map((m) => (
                <button
                  key={m}
                  className={viewMode === m ? 'is-on' : ''}
                  onClick={() => setViewMode(m)}
                  title={m === 'page' ? '页面视图' : m === 'web' ? 'Web 版式（连续）' : '草稿'}
                >
                  {m === 'page' ? '页' : m === 'web' ? '网' : '稿'}
                </button>
              ))}
            </span>
            <span className="wrp-status-zoom">
              <button onClick={() => setZoomPct((z) => Math.max(50, z - 10))} title="缩小">
                −
              </button>
              <input
                type="range"
                min={50}
                max={300}
                step={10}
                value={zoomPct}
                onChange={(e) => setZoomPct(Number(e.target.value))}
                title="缩放滑块"
              />
              <b
                title="点击恢复 100%"
                onClick={() => setZoomPct(100)}
                style={{ cursor: 'pointer' }}
              >
                {zoomPct}%
              </b>
              <button className="wrp-status-fit" onClick={fitWidth} title="页宽适配">
                页宽
              </button>
              <button className="wrp-status-fit" onClick={fitPage} title="整页适配">
                整页
              </button>
            </span>
          </span>
        )}
        {!editorMode && (
          <>
            <span className="wrp-sp" />
            <div className="wrp-actions">
              <button className="bar-btn" onClick={openInEditor} disabled={!tplBytes || !!busy}>
                在线编辑
              </button>
              <button className="bar-btn" onClick={exportPdf} disabled={!tplBytes || !!busy}>
                导出 PDF
              </button>
              <button className="bar-btn export-btn" onClick={exportDoc} disabled={!tplBytes || !!busy}>
                导出 Word
              </button>
            </div>
          </>
        )}
        <span className="wrp-sp" />
        {editorMode && (
          <button
            className="bar-btn"
            title={rightOpen ? '折叠数据预选追加' : '显示数据预选追加'}
            style={
              rightOpen
                ? { background: 'var(--ed-bg3)', color: 'var(--ed-accent)', borderColor: 'var(--ed-accent)', fontWeight: 600 }
                : undefined
            }
            onClick={() => setRightOpen((v) => !v)}
          >
            ☰ 数据预选
          </button>
        )}
      </div>

      {/* 放大预览：portal 到 body，避开属性面板的 overflow 裁剪 */}
      {zoom &&
        createPortal(
          <div className="wrp-overlay" onClick={() => setZoom(false)}>
            <div className="wrp-overlay-inner" onClick={(e) => e.stopPropagation()}>
              <div className="wrp-overlay-bar">
                <span>{config.docxTemplateName ?? '报告预览'}</span>
                <span className="wrp-overlay-hint">预览即导出结果 · A4 分页</span>
                <button className="bar-btn" onClick={() => setZoom(false)}>
                  关闭
                </button>
              </div>
              <DocxPreview data={preview} className="wrp-preview big" />
            </div>
          </div>,
          document.body,
        )}
      {/* OnlyOffice 在线编辑：弹层内部 portal 到 body，编辑器需要完整视口高度 */}
      {editing && (
        <OnlyOfficeModal
          file={editing}
          onSaved={onOfficeSaved}
          onClose={() => setEditing(null)}
          hint="DS 编辑的是绑定产物；手工精修的内容在保存回写后仍按 w:sdt 区域合并，其余字节保留。"
        />
      )}

      {/* 左栏 · 节点数据引用（portal · 常规=窗体外圆角卡片 / 最大化=嵌入窗体左缘）—— 与 Excel 编辑器同款 */}
      {editorMode && panelMounted && posForRender && (
        <RefPanel
          tables={tables}
          records={records}
          selfId={table.id}
          activeRef={null}
          onPick={pickFieldForActive}
          onClose={() => setRefPanelOpen(false)}
          animClass={maximized ? undefined : panelLeaving ? 'ref-panel-anim-out' : 'ref-panel-anim-in'}
          panelStyle={{
            position: 'fixed',
            left: posForRender.left,
            top: posForRender.top,
            height: posForRender.height,
            width: 230,
            borderRadius: posForRender.radius,
            border: posForRender.border ? '1px solid var(--ed-border)' : undefined,
            zIndex: 8100,
          }}
        />
      )}
      {/* 右栏 · 数据预选追加（镜像左栏：常规=窗体外右侧卡片 / 最大化=嵌入右缘） */}
      {editorMode && rightMounted && rightPosForRender && (
        <RightPanel
          tables={tables}
          records={records}
          selfId={table.id}
          inputs={config.inputs ?? []}
          onToggle={(srcId, on) => {
            const cur = config.inputs ?? []
            const next = on ? (cur.includes(srcId) ? cur : [...cur, srcId]) : cur.filter((x) => x !== srcId)
            onSetNodeConfig(table.id, { inputs: next })
          }}
          onClose={() => setRightOpen(false)}
          animClass={maximized ? undefined : rightLeaving ? 'ref-panel-anim-out' : 'ref-panel-anim-in'}
          panelStyle={{
            position: 'fixed',
            left: rightPosForRender.left,
            top: rightPosForRender.top,
            height: rightPosForRender.height,
            width: 230,
            borderRadius: rightPosForRender.radius,
            border: rightPosForRender.border ? '1px solid var(--ed-border)' : undefined,
            zIndex: 8100,
          }}
        />
      )}
    </div>
  )
}
