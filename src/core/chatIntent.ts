/* ─────────────────────────────────────────────────────────────
 * 对话意图引擎 · 模块化整理（2026-11-07，从 CanvasChatDock 抽出）
 *
 * 组成：
 *   - INTENTS：意图注册表（规则 patterns + 特殊动作 action）
 *   - matchIntent / intentById：规则匹配
 *   - collectCanvasDigest：画布节点数据摘要（喂模型）
 *   - fetchWeather：联网查天气（Open-Meteo，唯一联网点）
 *   - buildLlmPrompt：聪明模式 prompt 构造（意图路由规则+few-shot）
 *   - parseLlmIntent：LLM 输出的意图 JSON 解析（白名单校验）
 *
 * 套件关系（docs/30/31）：
 *   - 本模块的 LLM 路由依赖「smart-llm-qwen」套件提供的能力 llm:chat
 *   - 套件停用 → 调用方回落「本地快速」（bge 语义 + 规则）
 *
 * 扩展新意图：往 INTENTS 加一项；语义库同步 src/core/intentEmbed.ts 的 INTENT_DESCS。
 * ───────────────────────────────────────────────────────────── */
import { useBaseStore } from './store'
import { getBlob } from './pdfBlobStore'
import { styleOf, PALETTE_GROUPS } from './nodeStyle'

/** docx 字节 → 正文文字（jszip 解 word/document.xml，剥 w:t 标签，截断防超长） */
async function extractDocxText(bytes: Uint8Array, maxLen: number): Promise<string> {
  try {
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(bytes)
    const xml = await zip.file('word/document.xml')?.async('string')
    if (!xml) return ''
    return [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('').slice(0, maxLen)
  } catch { return '' }
}

/** ★ ChatIntent 类型（引擎模块定义，UI 层反向引用——避免循环依赖） */
export type ChatIntent = {
  id: string
  title: string
  patterns: RegExp[]
  fileTest?: (names: string[]) => boolean
  nodes: Array<{ kind: string; name: string; config?: Partial<import('./schema').NodeConfig> }>
  reply: string
  /** ★ 特殊动作：'analyze' = 读画布数据+模型总结；'weather' = 联网查天气；'open' = 用系统程序打开文档（唯一联网点=天气）；
   *  'run' = 真正执行当前工作流（如 PDF生成）——用户「已在画布创建…」之后接着说「执行转换」；
   *  'saveCanvas' = 保存画布工程（.wfc），与工具栏「保存」同一条链路 */
  action?: 'analyze' | 'weather' | 'open' | 'addNode' | 'run' | 'nodeInfo' | 'saveCanvas' | 'webSearch' | 'toPdf'
}

export const INTENTS: ChatIntent[] = [
  {
    id: 'word-to-pdf',
    title: 'Word → PDF',
    /* ★ 2026-11-08 用户「word转PDF」被 LLM 路由猜成 image-to-word（建出「图片导入→Word输出」）：
       规则层全不命中（add-node 只认「转成PDF」，裸「转PDF」不认）+ 意图清单里没有 Word→PDF，
       模型只能抓最像「转换」的意图瞎猜。而能力早已有——画布右键 OO 文档「转为PDF节点」（x2t 秒转），
       这里把对话层接上。必须排在 image-to-word 之前：其 pattern「(word|docx).{0,6}文档」会抢走「word文档转pdf」。 */
    patterns: [
      /(word|docx)\s*(文档)?\s*转\s*(换)?\s*(成|为|到)?\s*pdf/i,
      /(word|docx)\s*(to|2)\s*pdf/i,
      /* ★ 2026-11-09 「N105 转PDF」「把N105转成pdf」：编号引用的说法必须由规则直接接住——
         之前只靠 LLM 兜底路由，浏览器版/模型不可用时整句落空（回复"聪明模型运行时不可用"）。 */
      /[Nn]\s?0*\d{1,4}\s*(转|导|变|生成|输出)\s*(换)?\s*(成|为|到)?\s*pdf/i,
    ],
    nodes: [],
    reply: '',
    action: 'toPdf',
  },
  {
    id: 'webmon-watch',
    title: '网址监控工作流',
    /* ★ 2026-11-09 用户「你在对话框输入文字 让AI搭建出来」：一句话搭监控→AI分析链。
       patterns 刻意避开「加个XX节点」话术（那是 add-node 单节点语义），这里只接
       「工作流/链/系统」组合与「监控+变化+AI分析」的完整意图。 */
    patterns: [
      /(搭建|创建|做|搭)一?个?.{0,10}(网址|网页|网站).{0,4}(监控|监测).{0,10}(工作流|链|系统|流程)/i,
      /(网址|网页|网站).{0,4}(监控|监测).{0,14}(ai|AI|分析|分析一下)/,
      /监控.{0,12}(网址|网页|网站|页面).{0,14}(变化|更新).{0,12}(ai|分析|提醒|通知|自动)/i,
    ],
    nodes: [
      { kind: 'webmon', name: '网址监控', config: { webmonIntervalSec: 60 } },
      { kind: 'ai', name: 'AI 分析', config: {
        /* ★ 变量名=上游节点名（aiVarMapOf 兜底口径）——{{网址监控}} 自动插值为监控最新值 */
        aiUserPrompt: '以下是网址监控节点的最新数据：\n{{网址监控}}\n\n请分析上述数据的变化，指出关键变化并给出简短结论（150字内）。',
        aiTemperature: 0.3,
        aiMaxTokens: 1024,
      } },
    ],
    reply: '识别到「网址监控工作流」需求：将创建「网址监控 → AI 分析」链路——网址监控已默认 60 秒轮询（URL 在节点卡片上添加），检测到页面变化时自动触发 AI 分析（AI 节点需在属性面板填 API 设置）。',
  },
  {
    id: 'image-to-word',
    title: '图片 → Word',
    patterns: [/图[片像].{0,8}(转|生成|做成|变成).{0,6}(word|docx)/i, /(word|docx).{0,6}文档/i, /\bocr\b/i, /识别.{0,6}(文字|文本)/i],
    fileTest: (names) => names.some((n) => /\.(png|jpe?g|webp|bmp)$/i.test(n)),
    nodes: [
      { kind: 'photo', name: '图片导入' },
      { kind: 'docx', name: 'Word 输出' },
    ],
    reply: '识别到「图片转 Word」需求：将创建「图片导入 → Word 输出」工作流，图片自动导入并用本地 OCR 识别为可编辑文档。',
  },
  {
    id: 'analyze-node',
    title: '分析画布节点',
    patterns: [/分析[\s\S]{0,14}(节点|画布|文档|内容)/, /总结[\s\S]{0,14}(节点|画布|内容|文档)/, /(节点|画布|文档|内容)[\s\S]{0,16}(分析|总结)/, /画布.{0,8}(有|包含).{0,8}(什么|哪些)/, /(读|看)一下.{0,6}(节点|画布)/],
    nodes: [],
    reply: '',
    action: 'analyze',
  },
  {
    id: 'node-info',
    title: '查看节点',
    /* ★ 2026-11-07 用户「N071节点」→ 模型答「N071节点是什么意思」：
       节点卡片上明明写着 [N071]，对话却完全不认这个编号。这里先按规则命中（规则优先于 LLM），
       用编号定位节点并复述它是什么、吃什么、产出什么。 */
    patterns: [
      /^\s*\[?[Nn]0*\d{1,4}\]?\s*(节点?|node)?\s*[?？]?\s*$/,
      /\[?[Nn]0*\d{1,4}\]?.{0,10}(是什么|什么意思|是啥|啥意思|介绍|说明|详情|干嘛|干什么|干啥)/,
      /(节点|node)\s*\[?[Nn]0*\d{1,4}\]?/i,
    ],
    nodes: [],
    reply: '',
    action: 'nodeInfo',
  },
  {
    id: 'open-doc',
    title: '打开文档节点',
    patterns: [/打开.{0,8}(文档|word|docx)/, /(查看|看下|看看).{0,8}(已有|画布).{0,6}文档/],
    nodes: [],
    reply: '',
    action: 'open',
  },
  {
    id: 'query-weather',
    title: '查询天气',
    patterns: [/天气/, /气温|温度.{0,4}(多少|怎么样)?[?？]?$/, /(下不下|会不下)?雨/, /(冷不冷|热不热)/],
    nodes: [],
    reply: '',
    action: 'weather',
  },
  {
    id: 'web-search',
    title: '联网搜索',
    /* ★ 2026-11-07 用户「看不到联网引用记录」：新增通用联网搜索意图——Bing 搜索 → 模型带引用总结，
       气泡末尾展示来源链接列表。排在 weather 之后（「查天气」先归天气，避免抢跑）。 */
    patterns: [/^(帮我|请|给)?(联网|上网|网上)?(搜索|搜一下|搜一搜|查一下|查查|查一查|查搜)/, /(最新|今天|现在).{0,6}(消息|新闻|动态|行情|进展)/, /搜(个|一下)?\s*["'「]?[^"」]{2,}/],
    nodes: [],
    reply: '',
    action: 'webSearch',
  },
  {
    id: 'save-canvas',
    title: '保存画布文件',
    /* ★ 2026-11-07 用户说「保存画布文件」「保存整个画布文件」，对话却回了「已合并 1 份 / 3 页」——
       一句话里没有「执行」，落到了 run 动作（无编号 → 画布侧闷头挑一个 PDF生成 节点跑）。
       保存是**画布动作**，不是节点执行：这里给规则高优先级，命中就走保存链路，
       与工具栏「保存」完全同一条（选位置 → 清单确认 → 落盘）。
       注意别抢「导出为 PDF」「保存成 PDF 文件」——那些是产出文件，不是存工程。 */
    patterns: [
      /(保存|存储|存盘|存档|存入)\s*[一下]*\s*(整个|当前|这个|本)?\s*(画布|工程|工作流|项目)/,
      /(画布|工程|工作流|项目)\s*[一下]*(保存|存盘|存档|存下来|存一下)/,
      /(保存|另存)[^，。]{0,4}\.?(wfc)\b/i,
      /^\s*(保存|存盘|存档|save)\s*(一下|下|吧|啊|呀|呗)*\s*$/i,
    ],
    nodes: [],
    reply: '',
    action: 'saveCanvas',
  },
  {
    id: 'run-node',
    title: '执行工作流',
    /* ★ 2026-11-07 用户「已在画布创建 PDF生成 节点」之后说「执行转换」却没反应：
       这句话没有任何意图承接，落到闲聊兜底。这里补「执行/运行/开始/跑」这一组指令。 */
    patterns: [/^\s*(执行|运行|开始|跑|启动)\s*(转换|生成|合并|工作流|吧|一下)*\s*$/i, /(执行|运行|开始|跑)(一下|起来)?\s*(转换|生成|合并|工作流)/i, /(马上|立即|现在)(执行|生成|转换|开始)/, /(执行|运行|跑|开始)\s*\[?[Nn]0*\d{1,4}\]?/],
    nodes: [],
    reply: '',
    action: 'run',
  },
  {
    id: 'add-node',
    title: '添加节点',
    /* ★ 2026-11-07 用户「将文档节点的文件转换成PDF」没建节点：
       规则只认「输出为 pdf」，不认「转换成 PDF」→ 规则未命中，落到 LLM 猜成 add-node 意图，
       而 add-node 的 nodes 为空 → 卡片「一键创建」什么也建不出来。这里补「转/导/生/做成」。 */
    patterns: [/(增加|添加|新建|创建|加)个?.{0,10}节点/, /后面增加/, /输出为\s*(pdf|word|docx|excel)/i, /(转|导|生|做)(换)?(成|为)\s*(pdf|word|docx|excel)/i],
    nodes: [],
    reply: '',
    action: 'addNode',
  },
]

/** ★ 节点类型别名表：用户说法 → kind + 默认名（通用「加节点」用）
 *  ★ 2026-11-07 用户指令「要接一个 pdf 合并 而不是 PDF 节点，PDF 节点是承载 PDF 文件的」：
 *    - 「输出为 PDF / 生成 PDF / 出个 PDF」= 产出 PDF 的动作 → **pdfmerge（PDF生成）**
 *    - `pdf` kind 只承载**已有** PDF 文件（拖入/导入），不再承接「输出为PDF」这类话术
 *    - pdfmerge 的中文名同步改为「PDF生成」（用户：PDF合并节点 名称改成PDF生成） */
export const NODE_ALIASES: Array<{ match: RegExp; kind: string; name: string }> = [
  { match: /输出为\s*pdf|pdf\s*输出|输出\s*pdf|生成\s*pdf|pdf\s*生成|导出\s*pdf|出个\s*pdf|做个\s*pdf|转(换)?(成|为)\s*pdf/i, kind: 'pdfmerge', name: 'PDF生成' },
  { match: /输出为\s*(word|docx)|(word|docx)\s*输出/i, kind: 'docx', name: 'Word 输出' },
  { match: /pdf\s*合并|合并\s*pdf/i, kind: 'pdfmerge', name: 'PDF生成' },
  { match: /pdf\s*瘦身|压缩\s*pdf|瘦身/i, kind: 'pdfslim', name: 'PDF 瘦身' },
  { match: /pdf\s*转\s*word|pdf转word/i, kind: 'pdf2word', name: 'PDF 转 Word' },
  { match: /替换页|pdf\s*替换/i, kind: 'pdfreplace', name: 'PDF 替换页' },
  { match: /文件夹|目录/, kind: 'folder', name: '文件夹' },
  { match: /网址监控|网页监控|监控(网址|网页|网站)|站点监控/, kind: 'webmon', name: '网址监控' },
  { match: /(excel|xlsx|表格|坐标(表)?)\s*(转|生成|导出|输出)\s*(dxf|cad)|\s*dxf\s*(转|来自).{0,4}(excel|表格)|坐标展点|展点/i, kind: 'xlsx2dxf', name: 'Excel转DXF' },
  { match: /照片|图片导入|图片组/, kind: 'photo', name: '图片导入' },
  { match: /excel|表格/i, kind: 'oosheet', name: 'Excel 表格' },
  { match: /(word|docx)\s*文档/i, kind: 'ooword', name: 'Word 文档' },
  /* 承载型：只有明确「PDF 文件 / 导入 PDF / 放 PDF」才建 pdf 容器节点 */
  { match: /(pdf\s*文件|导入\s*pdf|放\s*pdf|已有\s*pdf|承载\s*pdf)/i, kind: 'pdf', name: 'PDF 文件' },
  { match: /pdf/i, kind: 'pdf', name: 'PDF 文件' },
  { match: /文本/, kind: 'text', name: '文本' },
  { match: /规则/, kind: 'rule', name: '规则' },
  { match: /\bai\b|ai节点/i, kind: 'ai', name: 'AI' },
  { match: /参数库/, kind: 'paramlib', name: '参数库' },
  { match: /章节/, kind: 'chapter', name: '章节工作簿' },
]

/** 从用户话里解析要添加的节点类型（匹配别名表） */
export function matchNodeAlias(text: string): { kind: string; name: string } | null {
  for (const a of NODE_ALIASES) {
    if (a.match.test(text)) return { kind: a.kind, name: a.name }
  }
  return null
}

export function matchIntent(msg: string, fileNames: string[]): ChatIntent | null {
  for (const it of INTENTS) {
    if (it.fileTest && it.fileTest(fileNames)) return it
    if (it.patterns.some((p) => p.test(msg))) return it
  }
  return null
}

export function intentById(id: string): ChatIntent | null {
  return INTENTS.find((x) => x.id === id) ?? null
}

/* ── 节点编号引用（N071）───────────────────────────────────────
 * 节点卡片标题就是「[N071] 名字」，所以用户很自然地用编号指代节点。
 * 编号来自 store 的 config.nodeNo（N001 起，创建/复制时分配）。 */
const NODE_REF_RE = /[Nn]\s?0*(\d{1,4})(?![\d])/g

/** 从一句话里取出所有节点编号（去重、按出现顺序），返回数字数组 [71] */
export function parseNodeRefs(text: string): number[] {
  const out: number[] = []
  for (const m of text.matchAll(NODE_REF_RE)) {
    const n = Number(m[1])
    if (Number.isFinite(n) && !out.includes(n)) out.push(n)
  }
  return out
}

/** 按编号找节点（N071 / n71 / N0071 都认；只看未删除节点） */
export function findNodeByRef(ref: number | string): { id: string; no: string; name: string; kind: string } | null {
  const want = typeof ref === 'number' ? ref : Number(String(ref).replace(/\D/g, ''))
  if (!Number.isFinite(want)) return null
  const st = useBaseStore.getState()
  for (const t of Object.values(st.tables)) {
    if (!t || t.deleted) continue
    const no = String((t.config as { nodeNo?: string } | undefined)?.nodeNo ?? '')
    const m = no.match(/^[Nn]0*(\d+)$/)
    if (m && Number(m[1]) === want) return { id: t.id, no, name: t.name, kind: t.kind }
  }
  return null
}

/** 句子里第一个能对上画布节点的编号引用 */
export function resolveNodeRef(text: string): { id: string; no: string; name: string; kind: string } | null {
  for (const n of parseNodeRefs(text)) {
    const hit = findNodeByRef(n)
    if (hit) return hit
  }
  return null
}

/** 画布节点编号清单（喂模型 / 找不到节点时提示用） */
export function listNodeRefs(limit = 12): string {
  const st = useBaseStore.getState()
  const rows = Object.values(st.tables)
    .filter((t) => t && !t.deleted)
    .map((t) => ({ t, no: String((t.config as { nodeNo?: string } | undefined)?.nodeNo ?? '') }))
    .filter((x) => x.no)
    .sort((a, b) => a.no.localeCompare(b.no))
  if (!rows.length) return '（画布上还没有编号节点）'
  const show = rows.slice(0, limit).map(({ t, no }) => {
    const c = (t.config ?? {}) as { inputs?: string[]; linkedInputs?: string[] }
    const ins = Math.max(c.inputs?.length ?? 0, c.linkedInputs?.length ?? 0)
    return `${no} ${t.name}${ins ? `（${ins} 个输入）` : ''}`
  })
  return show.join('；') + (rows.length > limit ? `；…共 ${rows.length} 个` : '')
}

/** 单个节点的中文说明（对话里回答「N071 是什么」） */
export function describeNodeByRef(ref: number | string): string | null {
  const hit = findNodeByRef(ref)
  return hit ? describeNodeById(hit.id) : null
}

/** 按节点 id 生成中文说明（对话与面板共用；找不到返回 null） */
export function describeNodeById(id: string): string | null {
  const st = useBaseStore.getState()
  const t = st.tables[id]
  if (!t || t.deleted) return null
  const no = String((t.config as { nodeNo?: string } | undefined)?.nodeNo ?? '')
  const hit = { id, no: no || '（无编号）', name: t.name }
  const cfg = (t.config ?? {}) as { inputs?: string[]; linkedInputs?: string[]; ooTemplateName?: string; ooBytesDbId?: string; ooBytes64?: string }
  const label = styleOf(t.kind).label
  const ups = (cfg.inputs?.length ? cfg.inputs : cfg.linkedInputs) ?? []
  const recs = st.records[t.id] ?? []
  const lines: string[] = [`「${hit.no}」是画布上的「${t.name}」节点（类型：${label}）。`]
  if (ups.length) {
    const names = ups.map((id) => st.tables[id]?.name ?? '（已删除节点）')
    const brief = names.slice(0, 4).join('、') + (names.length > 4 ? ` 等 ${names.length} 个` : '')
    lines.push(`· 输入：${ups.length} 个 —— ${brief}`)
  } else {
    lines.push('· 输入：暂无上游（还没连线/还没放内容）')
  }
  if (recs.length) lines.push(`· 数据：${recs.length} 条记录`)
  if (cfg.ooBytesDbId || cfg.ooBytes64) lines.push(`· 文档：已导入${cfg.ooTemplateName ? `「${cfg.ooTemplateName}」` : '内容'}`)
  const pos = st.positions[hit.id]
  if (pos) lines.push(`· 位置：画布 (${Math.round(pos.x)}, ${Math.round(pos.y)})`)
  const runnable = t.kind === 'pdfmerge'
  lines.push(runnable ? '要跑就点卡片上的「生成 PDF」，或直接说「执行 ' + hit.no + '」。' : '要处理它，直接说需求即可（如「分析 ' + hit.no + ' 的内容」）。')
  return lines.join('\n')
}

/** ★ 画布节点数据摘要（喂模型分析用；每节点截断防超长，最多 20 节点/3500 字） */
/** ★ 文档节点清单（OO word/sheet，含是否已导入内容） */
export function listDocNodes(): Array<{ id: string; name: string; kind: string; hasDoc: boolean; template: string }> {
  const st = useBaseStore.getState()
  return Object.values(st.tables)
    .filter((t) => !t.deleted && (t.kind === 'ooword' || t.kind === 'oosheet'))
    .map((t) => {
      const cfg = (t.config ?? {}) as { ooBytesDbId?: string; ooBytes64?: string; ooTemplateName?: string }
      return { id: t.id, name: t.name, kind: t.kind, hasDoc: !!(cfg.ooBytesDbId || cfg.ooBytes64), template: cfg.ooTemplateName ?? '' }
    })
}

type LlmChunk = { kind: 'text' | 'thought'; text: string }
type LlmStreamApi = (
  p: string,
  o: { model?: string; noThink?: boolean; maxTokens?: number },
  cb: (part: LlmChunk) => void,
) => Promise<{ ok: boolean; text?: string; reason?: string }>
type LlmApi = {
  llmChat?: (p: string, o?: { model?: string; noThink?: boolean; maxTokens?: number }) => Promise<{ ok: boolean; text?: string; reason?: string }>
  llmChatStream?: LlmStreamApi
}

/** ★ 当前用户锁定的本地模型文件名（对话面板模型菜单写入 localStorage 'llm-model'）。
 *  'auto'/缺失 = 不指定，主进程按优先级挑（Qwen3.5-9B → qwen2.5-3b）。 */
export function currentLlmModelOpt(): { model?: string } | undefined {
  try {
    const v = localStorage.getItem('llm-model')
    return v && v !== 'auto' ? { model: v } : undefined
  } catch { return undefined }
}

/** ★ 思考强度（2026-11-07 用户「选择模型的时候可以选择思考强度」，
 *  对话面板模型菜单写入 localStorage 'llm-thinking'）：
 *  off=不思考（主进程预填空 <think> 硬关思考，秒出正文，适合快问快答）；
 *  normal=常规思考（2048 token，思考可见）；deep=深度思考（4096 token，想得更久答得更全）。 */
export type ThinkingLevel = 'off' | 'normal' | 'deep'
export function currentThinkingLevel(): ThinkingLevel {
  try {
    const v = localStorage.getItem('llm-thinking')
    return v === 'off' || v === 'deep' ? v : 'normal'
  } catch { return 'normal' }
}
/** 模型 + 思考强度合成一次 LLM 调用的完整参数 */
export function currentLlmFullOpt(): { model?: string; noThink?: boolean; maxTokens?: number } {
  const base = currentLlmModelOpt() ?? {}
  const lv = currentThinkingLevel()
  if (lv === 'off') return { ...base, noThink: true, maxTokens: 1024 }
  if (lv === 'deep') return { ...base, maxTokens: 4096 }
  return { ...base, maxTokens: 2048 }
}

/* ── 优化混合模型路由（2026-11-09 用户「能否做成优化混合模型」）─────────────
 * Anthropic Routing 模式落地：简单/短输出任务走轻量 3B（快），复杂/质量任务走主力 9B。
 * 主进程会话池双常驻（互不 dispose），分发零切换成本。
 * 分层规则（用户未锁定模型时）：
 *   route/query（意图 JSON / 搜索词提取，短输出格式任务）→ light(3B)
 *   summarize/analyze（联网总结 / 文档分析，质量任务）→ heavy(9B)
 *   chat（通用对话）跟随思考档位：off→light 秒答，normal/deep→heavy
 * 用户在模型菜单锁定了具体模型 → 全尊重（不传 tier）。 */
export type LlmTask = 'route' | 'query' | 'chat' | 'summarize' | 'analyze'
export function routeLlmOpt(task: LlmTask): { model?: string; noThink?: boolean; maxTokens?: number; temperature?: number; tier?: 'light' | 'heavy' } {
  const base = currentLlmFullOpt()
  if (currentLlmModelOpt()?.model) return base
  const tier: 'light' | 'heavy' =
    task === 'route' || task === 'query' || (task === 'chat' && currentThinkingLevel() === 'off') ? 'light' : 'heavy'
  /* ★ 温度分任务（2026-11-09「怎么优化 使模型更聪明」）：对话/总结 0.5 自然，
     工具任务（route/query/analyze）0.1 稳定输出格式 */
  const temperature = task === 'chat' || task === 'summarize' ? 0.5 : 0.1
  return { ...base, tier, temperature }
}

/** ★ 软件能力卡（2026-11-09「怎么优化 使模型更聪明」）：从单一真相源 PALETTE_GROUPS 生成
 *  全部节点能力清单注入 system——模型「知道」自己是什么软件、有什么节点，回答操作类问题
 *  不再瞎猜（ACI 原则：工具描述质量 > 提示词话术）。 */
export function softwareCapabilityCard(): string {
  try {
    const lines = PALETTE_GROUPS.map((g) => `【${g.label}】${g.kinds.map((k) => k.label).join('、')}`)
    return `本软件是「午马未央」画布工作台：把节点拖上画布、连线组成工作流，双击节点编辑内容，节点卡片上有执行按钮。可用节点：\n${lines.join('\n')}`
  } catch { return '' }
}

/** ★ 取文档节点字节（系统程序打开用） */
export async function getDocNodeBytes(id: string): Promise<{ name: string; bytes: Uint8Array | null } | null> {
  const st = useBaseStore.getState()
  const t = st.tables[id]
  if (!t) return null
  const cfg = (t.config ?? {}) as { ooBytesDbId?: string; ooBytes64?: string; ooTemplateName?: string }
  let bytes: Uint8Array | null = null
  try {
    bytes = cfg.ooBytes64 ? Uint8Array.from(atob(cfg.ooBytes64), (ch) => ch.charCodeAt(0)) : cfg.ooBytesDbId ? ((await getBlob(cfg.ooBytesDbId)) ?? null) : null
  } catch { bytes = null }
  return { name: cfg.ooTemplateName || t.name, bytes }
}

/** ★ 深度分析单个文档节点：取正文（前 2500 字）→ Qwen 深入总结 */
export async function analyzeDocNode(
  id: string,
  api: { api?: LlmApi },
  onChunk?: (part: { kind: 'text' | 'thought'; text: string; accText: string; accThought: string }) => void,
): Promise<string> {
  const st = useBaseStore.getState()
  const t = st.tables[id]
  if (!t) return '节点不存在（可能已被删除）'
  const cfg = (t.config ?? {}) as { ooBytesDbId?: string; ooBytes64?: string; ooTemplateName?: string }
  let bytes: Uint8Array | null = null
  try {
    bytes = cfg.ooBytes64
      ? Uint8Array.from(atob(cfg.ooBytes64), (ch) => ch.charCodeAt(0))
      : cfg.ooBytesDbId ? ((await getBlob(cfg.ooBytesDbId)) ?? null) : null
  } catch { bytes = null }
  if (!bytes?.length) return `节点「${t.name}」还没有导入文档内容（双击卡片或用导入按钮放入 docx）`
  const text = await extractDocxText(bytes, 2500)
  if (!text) return `节点「${t.name}」的文档内容未能提取（复杂格式或非 docx 文件）`
  const llm = api.api ?? {}
  if (!llm.llmChat) return '聪明模型运行时不可用（缺少本地推理模块）。'
  /* ★ 2026-11-07 用户「还是没显示思考过程」：文档分析以前走非流式 llmChat + noThink
     （为修「思考超限没有正文」），于是思考段被整段关掉、界面上只有「正在分析…」。
     现在改走流式：思考段实时回传（气泡里能看直播），正文也逐字上屏；
     预算给到 3000 —— 主进程在正文仍被思考吃空时会自动用 noThink 重试一次兜底。 */
  const prompt = `以下是文档《${cfg.ooTemplateName || t.name}》的正文内容：\n${text}\n\n请深入分析这篇文档：说明主题、核心要点（分条）、结构与结论。用中文，250字内。`
  let r: { ok: boolean; text?: string; reason?: string }
  if (llm.llmChatStream) {
    let accText = ''
    let accThought = ''
    r = await llm.llmChatStream(prompt, routeLlmOpt('analyze'), (part) => {
      if (part.kind === 'thought') accThought += part.text
      else accText += part.text
      onChunk?.({ kind: part.kind, text: part.text, accText, accThought })
    })
  } else {
    r = await llm.llmChat(prompt, { ...routeLlmOpt('analyze'), noThink: true })
  }
  return r.ok ? (r.text ?? '').slice(0, 650) : `模型分析失败：${r.reason}`
}

/** ★ 画布摘要（async）：表节点带记录示例；OO 文档节点带 docx 正文摘录（前 900 字） */
export async function collectCanvasDigest(): Promise<string> {
  const st = useBaseStore.getState()
  const parts: string[] = []
  for (const t of Object.values(st.tables)) {
    if (t.deleted) continue
    const recs = st.records[t.id] ?? []
    const cfg = (t.config ?? {}) as { ooBytesDbId?: string; ooBytes64?: string; ooTemplateName?: string }
    let docText = ''
    const hasDoc = (t.kind === 'ooword' || t.kind === 'oosheet') && !!(cfg.ooBytesDbId || cfg.ooBytes64)
    if (t.kind === 'ooword' && hasDoc) {
      try {
        const bytes = cfg.ooBytes64
          ? Uint8Array.from(atob(cfg.ooBytes64), (ch) => ch.charCodeAt(0))
          : (await getBlob(cfg.ooBytesDbId!)) ?? null
        if (bytes?.length) docText = await extractDocxText(bytes, 900)
      } catch { /* ignore */ }
    }
    const sample = recs.slice(0, 3).map((r) => JSON.stringify(r).slice(0, 150)).join('；')
    const docPart = docText ? `，文档《${cfg.ooTemplateName ?? ''}》正文摘录：「${docText}」…` : (hasDoc ? '（含文档，正文未能提取）' : '')
    const no = String((t.config as { nodeNo?: string } | undefined)?.nodeNo ?? '')
    parts.push(`${no ? '[' + no + '] ' : ''}[${t.kind}] ${t.name}（${recs.length} 条记录）${sample ? '示例：' + sample : ''}${docPart}`)
    if (parts.length >= 20) break
  }
  return parts.join('\n').slice(0, 5500) || '（画布为空）'
}

/* ── 天气（Open-Meteo 免费无 key；唯一联网点，其余功能保持离线）────── */
const MAJOR_CITIES = ['北京', '上海', '天津', '重庆', '广州', '深圳', '成都', '杭州', '南京', '武汉', '西安', '苏州', '长沙', '郑州', '青岛', '大连', '厦门', '宁波', '合肥', '福州', '昆明', '贵阳', '哈尔滨', '沈阳', '石家庄', '太原', '南昌', '济南', '兰州', '西宁', '海口', '南宁', '香港', '澳门']
const WMO_ZH: Record<number, string> = { 0: '晴', 1: '基本晴', 2: '局部多云', 3: '阴', 45: '雾', 48: '雾凇', 51: '小毛雨', 53: '毛雨', 55: '大毛雨', 61: '小雨', 63: '中雨', 65: '大雨', 66: '冻雨', 71: '小雪', 73: '中雪', 75: '大雪', 80: '阵雨', 81: '强阵雨', 82: '暴雨', 95: '雷暴', 96: '雷暴伴冰雹' }

export function extractCity(text: string): string | null {
  const hit = MAJOR_CITIES.find((c) => text.includes(c))
  if (hit) return hit
  const stripped = text.replace(/(今天|明天|后天|现在|当前|目前|早上|下午|晚上|的)/g, '')
  const m = stripped.match(/([\u4e00-\u9fa5]{2,3}?)(?:市)?(?:天气|气温|温度|下雨|降雨)/)
  return m ? m[1] : null
}

/** 查天气：返回格式化文本（内部联网 geocoding + forecast） */
export async function fetchWeather(text: string): Promise<string> {
  let city = extractCity(text)
  const suffix = city ? '' : '（未识别到城市，默认查上海）'
  if (!city) city = '上海'
  const geo = await fetch(`https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(city)}&count=1&language=zh`).then((r) => r.json())
  const loc = geo.results?.[0]
  if (!loc) throw new Error(`没找到城市「${city}」，试试带上市名，比如「北京天气怎么样」`)
  const w = await fetch(`https://api.open-meteo.com/v1/forecast?latitude=${loc.latitude}&longitude=${loc.longitude}&current=temperature_2m,weather_code,wind_speed_10m&daily=temperature_2m_max,temperature_2m_min,precipitation_probability_max&timezone=auto&forecast_days=2`).then((r) => r.json())
  const c = w.current
  const d = w.daily
  const desc = WMO_ZH[c.weather_code] ?? '未知'
  return `${suffix ? suffix.replace(/^[（(]/, '（') : ''}${loc.name}当前：${desc}，${Math.round(c.temperature_2m)}°C，风速 ${Math.round(c.wind_speed_10m)} km/h。\n今日 ${Math.round(d.temperature_2m_min[0])}~${Math.round(d.temperature_2m_max[0])}°C，降水概率 ${d.precipitation_probability_max[0] ?? 0}%。\n明日 ${Math.round(d.temperature_2m_min[1])}~${Math.round(d.temperature_2m_max[1])}°C，降水概率 ${d.precipitation_probability_max[1] ?? 0}%。\n\n来源：Open-Meteo（open-meteo.com）`
}

/** ★ 从用户话里抽取搜索关键词（web-search 意图）。
 *  ★ 2026-11-08 用户截图实测：上文聊「根据黄历」，用户接「你联网查询」——
 *  旧版剥不掉「你」这个主语，拿着「你联网查询」去 Bing 搜，返回「你」字的词典解释（答非所问）。
 *  现在：人称/指令词循环剥净（stripQueryPhrases）；剥完是空/纯虚词 → 返回 ''，
 *  由调用方（CanvasChatDock）结合最近对话补全搜索词，而不是瞎搜。 */
const WEB_STRIP_RE = /^(帮我|请|给|麻烦|你|您|我|咱)?(麻烦你|请你|帮我|给我|去|帮我)?(联网|上网|在网上|网上)?(搜索|搜一下|搜一搜|搜搜|查一下|查一查|查查|查搜|查询)(一下)?(关于)?/i

/** 循环剥离搜索指令词（人称 + 连接词 + 搜索动词），直到剥不动为止 */
export function stripQueryPhrases(text: string): string {
  let q = text.trim()
  let prev = ''
  while (q !== prev) {
    prev = q
    q = q.replace(WEB_STRIP_RE, '').trim()
  }
  return q
}

/** 剥完指令词后剩下的算不算有效搜索词（纯代词/语气词/疑问词 → 无效，交给上下文补全） */
export function isUsableQuery(q: string): boolean {
  const s = q.trim()
  if (s.length < 2) return false
  return !/^(你|我|他|她|它|什么|怎么|怎样|怎么样|吗|呢|吧|啊|的|了|么)+[?？!！。~～]*$/.test(s)
}

export function extractWebQuery(text: string): string {
  /* 引号包起来的直接用引号内容 */
  const quoted = text.match(/["'「『]([^"'」』]{2,})["'」』]/)
  if (quoted) return quoted[1].trim()
  const q = stripQueryPhrases(text).replace(/[?？!！。]+$/, '').replace(/^(个|下|些)\s*/, '').trim()
  /* ★ 剥空/纯虚词返回 ''（旧版返回原文，导致拿「你联网查询」整句去搜） */
  return isUsableQuery(q) ? q : ''
}

/** ★ 搜索词缺失时的兜底：从最近对话的用户消息里找主题（剥「根据/按照」等引导词）。
 *  「你联网查询」+ 上文「根据黄历」→ 「黄历」。找不到返回 ''。 */
export function fallbackQueryFromTurns(turns: Array<{ role: string; text: string }>): string {
  const users = turns.filter((t) => t.role === 'user').map((t) => t.text.trim()).filter(Boolean)
  for (let i = users.length - 1; i >= 0; i--) {
    const s = stripQueryPhrases(users[i])
      .replace(/^(根据|按照|按|依|看|就|那|按着)\s*/, '')
      .replace(/[?？!！。]+$/, '')
      .trim()
    if (isUsableQuery(s)) return s
  }
  return ''
}

/* ── 聪明模式 LLM 路由 ─────────────────────────────────────────── */

/** 构造意图路由 prompt（意图 JSON vs 自由回答，含 few-shot 与边界规则） */
export function buildLlmPrompt(text: string): string {
  const list = INTENTS.filter((x) => x.action !== 'weather').map((x) => `- ${x.id}：${x.title}${x.action === 'analyze' ? '（自动分析画布节点内容并总结）' : ''}`).join('\n')
  return `你是画布工作流助手的意图路由器。可用意图：\n${list}\n\n画布上的节点（用户说「N0xx」时指的就是这个编号）：\n${listNodeRefs()}\n\n判断规则：\n- 用户提到某个编号（如 N071），就是画布上那个节点；不要说不认识这个编号\n- 用户想执行上述某个工作流，或想分析/总结画布上已有节点的内容 → 只输出 {"intent":"<意图id>"}\n- 用户想打开/查看已有文档节点 → 直接回答：「双击节点卡片即可打开文档」\n- 其他任何情况（提问/闲聊/打招呼/问能力）→ 直接用中文友好回答；无法确定时，说明你是工作流助手并列出能做的事，禁止输出"用户输入不明确"这类话\n\n示例：\n用户：把这张截图弄成能编辑的文档 → {"intent":"image-to-word"}\n用户：word转PDF → {"intent":"word-to-pdf"}\n用户：搭一个网址监控工作流，网页变化后让AI分析 → {"intent":"webmon-watch"}\n用户：几张扫描件合成一个文件 → {"intent":"pdf-merge"}\n用户：分析下文档节点中的内容并总结 → {"intent":"analyze-node"}
用户：N071节点 → {"intent":"node-info"}\n用户：保存画布文件 → {"intent":"save-canvas"}\n用户：打开文档的word → 双击节点卡片即可打开文档\n用户：你好，你能做什么 → 我是工作流助手，可以帮你把图片转成 Word、生成 PDF、汇总 Excel、分析画布节点内容\n\n用户输入：${text}\n你的输出：`
}

/** 解析 LLM 输出：意图 JSON 白名单校验 + 对话模板清洗（小模型爱照抄「用户：」「我：」示例格式） */
export function parseLlmIntent(reply: string): { intentId: string | null; cleanText: string } {
  const m = reply.match(/\{\s*"intent"\s*:\s*"([\w-]+)"\s*\}/)
  const id = m ? m[1] : null
  const known = id ? INTENTS.some((x) => x.id === id) : false
  let clean = reply.replace(/\{\s*"intent"[\s\S]*?\}/g, '').trim()
  /* 清洗对话模板残留：取最后一个「我：」之后的实质内容；剥「用户：xxx」行 */
  if (/[我我方]：/.test(clean)) {
    const segs = clean.split(/(?:^|\n)\s*(?:我|助手|AI)：/)
    const last = segs[segs.length - 1]?.trim()
    if (last) clean = last
  }
  clean = clean.split('\n').filter((l) => !/^\s*用户[：:]/.test(l)).join('\n').trim()
  return { intentId: known ? id : null, cleanText: clean.slice(0, 300) }
}
