/* ★ 2026-10-31 画布工作流对话 v7（白主题适配，用户 5 项反馈）：
 *  1. 展开不再被 ribbon 遮挡（zIndex 提到 250）
 *  2. 适配白色画布主体——面板白底浅灰边、深色文字、OO 工具栏同风格
 *  3. 收起动画柔化——收起前内容先淡出 0.18s
 *  4. 圆角矩形按钮统一 OO 白色工具栏风格；📂 换成窗体右上角同款文件夹图标（内联 SVG）
 *  5. 输入框文字左右 padding 12（placeholder/光标不贴边）
 *  展开动画保持：▶ 圆扩展→胶囊向右（图标随行到右端）→向上伸展至窗体大小 */
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import { useBaseStore } from './core/store'
import { hasCapability } from './core/suites'
import { INTENTS, matchIntent, intentById, collectCanvasDigest, fetchWeather, buildLlmPrompt, parseLlmIntent, listDocNodes, analyzeDocNode, getDocNodeBytes, matchNodeAlias, resolveNodeRef, describeNodeById, listNodeRefs, currentLlmModelOpt, currentLlmFullOpt, currentThinkingLevel, extractWebQuery, isUsableQuery, fallbackQueryFromTurns, routeLlmOpt, softwareCapabilityCard } from './core/chatIntent'
import { buildContextBlock, needsMoreInfo, nowLine, isDateTimeQuestion } from './core/chatContext'
import type { ChatIntent } from './core/chatIntent'

type ChatMsg = {
  /** ★ 稳定 id：流式更新按 id 定位气泡（按文本匹配会在文本变化后失效） */
  id?: string
  role: 'user' | 'assistant'
  text: string
  files?: string[]
  /** ★ 意图卡片（本地轻量意图引擎命中后附带）：一键创建工作流 */
  intent?: { id: string; title: string; plan: string; accepted: boolean }
  /** ★ 文档选择菜单（多文档节点分析时）：点了分析对应文档 */
  docChoices?: Array<{ nodeId: string; label: string; mode?: 'analyze' | 'open'}>
  /** ★ 思考过程（Qwen3.5 思考段，流式收集；默认折叠可展开） */
  thought?: string
  /** ★ 正在思考/生成中：true 时思考块自动展开直播，结束后自动收起 */
  thinking?: boolean
  /** ★ 联网引用来源（web-search 意图）：气泡末尾的链接列表 */
  sources?: Array<{ title: string; url: string }>
}

/* ★ 本地轻量意图引擎（2026-11-01 用户指令「本地轻量模型判断匹配简单工作流」）：
   规则+模式匹配（零外部依赖）；命中即出意图卡片，一键在画布创建工作流节点。
   扩展方式：往 INTENTS 数组加一项（patterns/fileTest/nodes/reply）。 */

/** 按意图 id 取 ChatIntent（语义匹配命中后回查） */
type Phase = 'closed' | 'pill' | 'grow' | 'open'

/* OO 工具栏同款文件夹图标（线条风） */
const FolderIcon = () => (
  <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" />
  </svg>
)
/* 回形针图标（用户指定样式：斜别针线性） */
const PaperclipIcon = () => (
  <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48" />
  </svg>
)

/* ★ 本地模型菜单选项（models/ 目录实际存在的文件；'auto' = 主进程按优先级挑） */
const LLM_MODEL_OPTIONS = [
  { file: 'Qwen3.8-27B-UD-Q3_K_XL.gguf', icon: '🏆', label: '高端模型', desc: 'Qwen3.8-27B · 13.2GB 离线 · 最聪明（首次加载约 30-60 秒，需 16G 显存）' },
  { file: 'Qwen3.5-9B-Q4_K_M.gguf', icon: '🧠', label: '聪明模型', desc: 'Qwen3.5-9B · 5.7GB 离线 · 更聪明（约 8-15 秒）' },
  { file: 'qwen2.5-3b-instruct-q4_k_m.gguf', icon: '⚡', label: '轻量模型', desc: 'Qwen2.5-3B · 2GB 离线 · 更快（约 3-5 秒）' },
]
const llmModelMeta = (file: string) => LLM_MODEL_OPTIONS.find((m) => m.file === file) ?? LLM_MODEL_OPTIONS[0]
/* ★ 思考强度三档（与 chatIntent.currentLlmFullOpt 的 localStorage 'llm-thinking' 约定一致） */
const THINKING_LEVELS: Array<{ id: 'off' | 'normal' | 'deep'; label: string; desc: string }> = [
  { id: 'off', label: '⚡ 不思考', desc: '跳过思考直接回答 · 最快' },
  { id: 'normal', label: '🧠 常规思考', desc: '思考过程可见 · 速度与质量均衡' },
  { id: 'deep', label: '🔬 深度思考', desc: '更大思考预算 · 想得更久答得更全' },
]

export function CanvasChatDock() {
  const [phase, setPhase] = useState<Phase>('open')
  const [input, setInput] = useState('')
  const [files, setFiles] = useState<string[]>([])
  const [fileObjs, setFileObjs] = useState<File[]>([])
  const [dragOver, setDragOver] = useState(false)
  const [msgs, setMsgs] = useState<ChatMsg[]>([])
  /* ★ 最新 msgs 的快照（send 里拼上下文用：state 在闭包里是旧的，ref 才是最新） */
  const msgsRef = useRef<ChatMsg[]>([])
  useEffect(() => { msgsRef.current = msgs }, [msgs])
  const [stretch, setStretch] = useState(false)
  /* ★ 左右停靠：标题栏拖动，x<50% 停左、否则停右（松手记忆到 localStorage） */
  const [dock, setDock] = useState<'l' | 'r'>(() => {
    try { return localStorage.getItem('cb-dock-side') === 'r' ? 'r' : 'l' } catch { return 'l' }
  })
  /* ★ 2026-11-07 拖拽换边：整段视觉走「直接改 DOM」，不再 setState。
     原因（用户「拖的时候怎么卡住了」）：pointermove 每帧 setState，会把整个对话面板
     （消息列表 + 输入框 + 意图卡片）重渲染一遍，帧率直接塌掉。现在拖动期间 React 零渲染，
     只有松手 setDock 一次。落点幽灵框与面板同 rect，保证「提示 = 实际框体」。 */
  const ghostRef = useRef<HTMLDivElement | null>(null)
  const ghostLabelRef = useRef<HTMLDivElement | null>(null)
  const [contentFade, setContentFade] = useState(false)
  /* ★ 2026-11-07 用户「怎么这么窄」：面板宽度可拖拽（左右边缘手柄）+ 记忆。
     0 = 用默认宽度（min(600px, 56%)）；拖动后写 localStorage，下次启动沿用。 */
  const [panelW, setPanelW] = useState<number>(() => {
    try { return Number(localStorage.getItem('cb-dock-w') || 0) || 0 } catch { return 0 }
  })
  /* 右键菜单（复制/引用）
     ★ 2026-11-07 用户「右键复制位置错位」：不再直接用鼠标坐标定位——
     有选中文字时锚在**选区矩形**下方（用户右键的意图是操作那段文字），
     只有完全没选区时才退回鼠标位置。anchor 为 null 表示无选区。 */
  const [ctxMenu, setCtxMenu] = useState<{
    x: number
    y: number
    anchor: { left: number; top: number; right: number; bottom: number; width: number } | null
    text: string
    fromSel?: boolean
  } | null>(null)
  const menuRef = useRef<HTMLDivElement | null>(null)
  const [menuPos, setMenuPos] = useState<{ left: number; top: number } | null>(null)
  const [, tick] = useState(0)
  /* ★ 2026-10-31 主题跟随：root data-theme（light/dark）+ app-theme-change 事件 */
  const [isLight, setIsLight] = useState(document.documentElement.getAttribute('data-theme') === 'light')
  const flowRef = useRef<HTMLDivElement | null>(null)
  const panelRef = useRef<HTMLDivElement | null>(null)
  const taRef = useRef<HTMLTextAreaElement | null>(null)
  const fileRef = useRef<HTMLInputElement | null>(null)
  const fileRef2 = useRef<HTMLInputElement | null>(null)

  useEffect(() => {
    const onTheme = () => setIsLight(document.documentElement.getAttribute('data-theme') === 'light')
    window.addEventListener('app-theme-change', onTheme)
    const obs = new MutationObserver(onTheme)
    obs.observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] })
    const onR = () => tick((t) => t + 1)
    window.addEventListener('resize', onR)
    return () => {
      window.removeEventListener('app-theme-change', onTheme)
      window.removeEventListener('resize', onR)
      obs.disconnect()
    }
  }, [])
  useEffect(() => {
    if (flowRef.current) flowRef.current.scrollTop = flowRef.current.scrollHeight
  }, [msgs])

  /* 用户消息入列（不含旧引导——意图链接管回复） */
  const pushUserOnly = useCallback((text: string, fs: string[]) => {
    setMsgs((m) => [...m, { role: 'user', text, files: fs.length ? fs : undefined }])
  }, [])

  /* ★ 2026-11-07 用户指令「本地快速删掉，聪明模型保留 新增刚才的」：
     不再有 fast/smart 双模式——一律走本地 GGUF 模型，
     规则匹配仍在 LLM 之前（明确的正则比模型快且准）。 */
  /* ★ 2026-11-07 用户「恢复」模型切换：🧠 徽标重新可点，弹菜单在 Qwen3.5-9B（聪明）/ Qwen2.5-3B（轻快）之间切换，
     选择持久化 localStorage 'llm-model'，主进程按指定文件加载（切换时释放旧模型）。 */
  const [llmModel, setLlmModel] = useState<string>(() => {
    try { return localStorage.getItem('llm-model') ?? LLM_MODEL_OPTIONS[0].file } catch { return LLM_MODEL_OPTIONS[0].file }
  })
  useEffect(() => { try { localStorage.setItem('llm-model', llmModel) } catch { /* ignore */ } }, [llmModel])
  /* ★ 思考强度（2026-11-07 用户「选择模型的时候可以选择思考强度」）：
     off=硬关思考秒出正文；normal=常规思考；deep=深度思考（更大 token 预算） */
  const [thinkingLv, setThinkingLv] = useState<'off' | 'normal' | 'deep'>(() => {
    try { const v = localStorage.getItem('llm-thinking'); return v === 'off' || v === 'deep' ? v : 'normal' } catch { return 'normal' }
  })
  useEffect(() => { try { localStorage.setItem('llm-thinking', thinkingLv) } catch { /* ignore */ } }, [thinkingLv])
  const [modelMenu, setModelMenu] = useState(false)
  /* ★ 思考直播滚动：思考内容在长高，块内部要始终显示最底下（最新）的内容 */
  const thoughtBoxRef = useRef<HTMLDivElement | null>(null)
  useEffect(() => {
    const el = thoughtBoxRef.current
    if (el) el.scrollTop = el.scrollHeight
  }, [msgs])
  /* ★ 流式气泡的 msgs 下标（send() 期间唯一，完成即定稿） */
  const streamIdxRef = useRef(-1)

  /* ★ onChunk 流式回调（2026-11-07 用户「不要结果好了再出」「看不到思考过程」）：
     每次回调给 { thoughtAcc, textAcc } 累计正文/累计思考段；不传则退化为一次性返回 */
  const smartChat = useCallback(async (text: string, ctxBlock = '', onChunk?: (st: { thoughtAcc: string; textAcc: string }) => void): Promise<{ intentHit: ChatIntent | null; reply: string }> => {
    /* ★ 套件闸门（docs/30）：依赖「smart-llm-qwen」套件的 llm:chat 能力，停用即回落 */
    if (!hasCapability('llm:chat')) {
      return { intentHit: null, reply: '聪明模型套件已停用。到「套件中心」重新启用「对话助手 · 本地聪明模型」。' }
    }
    const list = INTENTS.filter((x) => x.id !== 'analyze-node').map((x) => `- ${x.id}：${x.title}`).join('\n')
    /* ★ 2026-11-09 软件能力卡注入（「怎么优化 使模型更聪明」）：模型知道自己是什么软件、有什么节点 */
    const sys = `你是工作流意图路由器。${nowLine()}。${softwareCapabilityCard()}\n\n可用意图：\n${list}\n\n画布上现有的节点（标题里的 [N0xx] 就是它们的编号，用户说编号时指的就是这些节点）：\n${listNodeRefs()}\n\n${ctxBlock}判断规则：\n- 用户提到编号（N071 之类）就是在说那个节点，绝对不要回答“不知道这个编号”\n- 用户想执行上述某个工作流 → 只输出 {"intent":"<意图id>"}\n- 其他任何情况（提问/闲聊/打招呼/问能力）→ 直接用中文回答，绝对不要输出 JSON\n- 若当前问题是上文的追问（“穿多少衣服”“那明天呢”之类），**必须结合【最近对话】里的数据直接回答**，不要说“无法确定/请提供信息”，也不要让用户再去查一遍\n- 若【最近对话】确实缺少必要信息，先用一句话说明缺什么，再给出你能给的建议\n- 与【最近对话】无关（换了话题）时忽略它，按当前问题回答\n- 与画布/文档工作流无关的日常问题（天气、穿衣、闲聊）：直接自然地回答，**绝对不要提“意图/节点/工作流/query-xxx”，也不要把用户往功能上引**\n- 问日期/星期/时间类问题必须以上面给出的当前时间为准推算，绝对不要凭记忆编日期\n\n示例：\n用户：把这张截图弄成能编辑的文档 → {"intent":"image-to-word"}\n用户：几张扫描件合成一个文件 → {"intent":"pdf-merge"}\n用户：N071节点 → {"intent":"node-info"}\n用户：你好，你能做什么 → 我是工作流助手，可以帮你把图片转成 Word、生成 PDF、汇总 Excel 表格\n\n用户输入：${text}\n你的输出：`
    const api = (window as unknown as {
      api?: {
        llmChat?: (p: string, o?: { model?: string }) => Promise<{ ok: boolean; text?: string; reason?: string }>
        llmChatStream?: (p: string, o: { model?: string }, onChunk: (part: { kind: 'text' | 'thought'; text: string }) => void) => Promise<{ ok: boolean; text?: string; reason?: string }>
      }
    }).api
    if (!api?.llmChat) return { intentHit: null, reply: '聪明模型运行时不可用（缺少本地推理模块）。' }
    /* ⓪ 身份问题直接答 */
    if (/什么模型|你是谁|你是哪个|怎么实现的|哪个版本/.test(text.trim())) {
      const cur = LLM_MODEL_OPTIONS.find((m) => m.file === (currentLlmModelOpt()?.model ?? LLM_MODEL_OPTIONS[0].file)) ?? LLM_MODEL_OPTIONS[0]
      return { intentHit: null, reply: `我是运行在你电脑上的本地模型：${cur.desc}（完全离线，node-llama-cpp 推理），配合规则引擎负责识别工作流意图和回答画布相关问题。数据不出你的电脑。` }
    }
    /* ★ 闲聊快筛（0.5B/1.5B 都爱编 JSON）：非工作流话题直接回答，不给模型编造机会 */
    if (/^(你好|您好|嗨|hi|hello|谢谢|感谢|今天天气|你是谁|你叫什么)/i.test(text.trim())) {
      return { intentHit: null, reply: '你好！我是画布的工作流助手，可以帮你：把图片转成 Word、合并 PDF、汇总 Excel 表格。拖个文件进来或直接说需求就行。' }
    }
    /* ★ 纯日期/时间问题本地直答（2026-11-08 用户「天气、日期这类都不联网的么」）：
     *  不联网也不耗模型——直接读系统时间。放给模型答反而危险：训练截止后它不知道今天几号，
     *  实测会烧掉 7000+ 字思考再编一个错日期（截图 2026-10-21）。
     *  「明天几号」这类相对推算不拦，交给注入了当前时间的模型去算。 */
    if (isDateTimeQuestion(text)) {
      return { intentHit: null, reply: `${nowLine()}。` }
    }
    /* ★ 流式优先：onResponseChunk 同时吐思考段与正文，分开累计后回调；旧 preload 没有流式能力时退化为一次性 llmChat */
    let accText = ''
    let accThought = ''
    const r = api.llmChatStream
      ? await api.llmChatStream(sys, routeLlmOpt('chat'), (p) => {
          if (p.kind === 'thought') accThought += p.text
          else accText += p.text
          onChunk?.({ thoughtAcc: accThought, textAcc: accText })
        })
      : await api.llmChat(sys, routeLlmOpt('chat'))
    if (!r.ok) return { intentHit: null, reply: `聪明模型出错：${r.reason}` }
    const reply: string = r.text ?? ''
    const m = reply.match(/\{\s*"intent"\s*:\s*"([\w-]+)"\s*\}/)
    if (m && INTENTS.some((x) => x.id === m[1])) {
      const intent = intentById(m[1])!
      return { intentHit: intent, reply: `${intent.reply}（聪明模型判断）` }
    }
    return { intentHit: null, reply: reply.slice(0, 300).replace(/\{\s*"intent"[\s\S]*?\}/g, '').trim() || '（模型未返回内容）' }
  }, [])

  /** ★ 收集画布节点数据摘要（喂给 Qwen 分析用；每节点截断防超长） */
  const collectCanvasDigest = useCallback((): string => {
    const st = useBaseStore.getState()
    const parts: string[] = []
    for (const t of Object.values(st.tables)) {
      if (t.deleted) continue
      const recs = st.records[t.id] ?? []
      const sample = recs.slice(0, 3).map((r) => JSON.stringify(r).slice(0, 180)).join('；')
      const no = String((t.config as { nodeNo?: string } | undefined)?.nodeNo ?? '')
      parts.push(`${no ? '[' + no + '] ' : ''}[${t.kind}] ${t.name}（${recs.length} 条记录）示例：${sample || '（空）'}`)
      if (parts.length >= 20) break
    }
    return parts.join('\n').slice(0, 3500) || '（画布为空）'
  }, [])

  /** ★ 查天气（Open-Meteo 免费无 key；用户问天气才联网，其余全离线） */
  const runWeather = useCallback(async (text: string) => {
    setMsgs((m) => [...m, { role: 'assistant', text: '正在查询天气…' }])
    /* 城市提取三步：① 常见城市白名单精确命中（最准）② 剥时间词后非贪婪正则 ③ 兜底上海 */
    const MAJOR = ['北京', '上海', '天津', '重庆', '广州', '深圳', '成都', '杭州', '南京', '武汉', '西安', '苏州', '长沙', '郑州', '青岛', '大连', '厦门', '宁波', '合肥', '福州', '昆明', '贵阳', '哈尔滨', '沈阳', '石家庄', '太原', '南昌', '济南', '兰州', '西宁', '海口', '南宁', '香港', '澳门']
    let city = MAJOR.find((c) => text.includes(c)) ?? null
    if (!city) {
      const stripped = text.replace(/(今天|明天|后天|现在|当前|目前|早上|下午|晚上|的)/g, '')
      const m = stripped.match(/([\u4e00-\u9fa5]{2,3}?)(?:市)?(?:天气|气温|温度|下雨|降雨)/)
      city = m ? m[1] : null
    }
    const suffix = city ? '' : '（未识别到城市，默认查上海）'
    if (!city) city = '上海'
    try {
      const geo = await fetch(`https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(city)}&count=1&language=zh`).then((r) => r.json())
      const loc = geo.results?.[0]
      if (!loc) {
        setMsgs((m) => m.map((x) => x.text === '正在查询天气…' ? { ...x, text: `没找到城市「${city}」，试试带上市名，比如「北京天气怎么样」` } : x))
        return
      }
      const w = await fetch(`https://api.open-meteo.com/v1/forecast?latitude=${loc.latitude}&longitude=${loc.longitude}&current=temperature_2m,weather_code,wind_speed_10m&daily=temperature_2m_max,temperature_2m_min,precipitation_probability_max&timezone=auto&forecast_days=2`).then((r) => r.json())
      const wmo: Record<number, string> = { 0: '晴', 1: '基本晴', 2: '局部多云', 3: '阴', 45: '雾', 48: '雾凇', 51: '小毛雨', 53: '毛雨', 55: '大毛雨', 61: '小雨', 63: '中雨', 65: '大雨', 66: '冻雨', 71: '小雪', 73: '中雪', 75: '大雪', 80: '阵雨', 81: '强阵雨', 82: '暴雨', 95: '雷暴', 96: '雷暴伴冰雹' }
      const c = w.current
      const d = w.daily
      const desc = wmo[c.weather_code] ?? '未知'
      const line = `${loc.name}当前：${desc}，${Math.round(c.temperature_2m)}°C，风速 ${Math.round(c.wind_speed_10m)} km/h。\n今日 ${Math.round(d.temperature_2m_min[0])}~${Math.round(d.temperature_2m_max[0])}°C，降水概率 ${d.precipitation_probability_max[0] ?? 0}%。\n明日 ${Math.round(d.temperature_2m_min[1])}~${Math.round(d.temperature_2m_max[1])}°C，降水概率 ${d.precipitation_probability_max[1] ?? 0}%。`
      setMsgs((m) => m.map((x) => x.text === '正在查询天气…' ? { ...x, text: line } : x))
    } catch (e) {
      setMsgs((m) => m.map((x) => x.text === '正在查询天气…' ? { ...x, text: `天气查询失败（需要联网）。${String((e as Error)?.message || e).slice(0, 80)}` } : x))
    }
  }, [])

  /** ★ 用系统默认程序（WPS/Word）打开文档节点：单文档直接开；多个出选择菜单 */
  const runDocOpen = useCallback(async (text: string) => {
    setMsgs((m) => [...m, { role: 'assistant', text: '正在打开文档…' }])
    const api = (window as unknown as { api?: { openExternalFile?: (name: string, data: Uint8Array) => Promise<string> } }).api
    const docs = listDocNodes()
    const clean = text.replace(/(打开|查看|看下|看看|帮我|的|请)/g, '').trim()
    const named = clean ? docs.filter((d) => d.name.includes(clean) || d.template.includes(clean) || clean.includes(d.name)) : []
    const target = named.length === 1 ? named[0] : null
    if (docs.length === 0) {
      setMsgs((m) => m.map((x) => x.text === '正在打开文档…' ? { ...x, text: '画布上还没有文档节点——先导入一个 docx，我就能帮你打开。' } : x))
      return
    }
    if (!target && docs.length > 1) {
      setMsgs((m) => m.map((x) => x.text === '正在打开文档…' ? {
        ...x,
        text: `画布上有 ${docs.length} 个文档节点，要打开哪一个？`,
        docChoices: docs.map((d) => ({ nodeId: d.id, label: `📄 ${d.name}${d.template ? '（' + d.template + '）' : ''}`, mode: 'open' as const })),
      } : x))
      return
    }
    const doc = target ?? docs[0]
    const { name, bytes } = (await getDocNodeBytes(doc.id)) ?? { name: doc.name, bytes: null }
    if (!bytes?.length) {
      setMsgs((m) => m.map((x) => x.text === '正在打开文档…' ? { ...x, text: `节点「${doc.name}」还没有导入文档内容。` } : x))
      return
    }
    if (!api?.openExternalFile) {
      setMsgs((m) => m.map((x) => x.text === '正在打开文档…' ? { ...x, text: '打开功能不可用（缺少系统调用接口）。' } : x))
      return
    }
    await api.openExternalFile(name, bytes)
    setMsgs((m) => m.map((x) => x.text === '正在打开文档…' ? { ...x, text: `已用系统默认程序（WPS/Word）打开「${name}」。` } : x))
  }, [])

  /** ★ 文档分析统一入口（runAnalyze / 选择菜单共用）。
      2026-11-07 用户「还是没显示思考过程」：以前这里走非流式 llmChat（且强制关思考），
      界面上只有一句「正在分析…」。现在走流式，思考段实时写进同一条气泡（可展开回看）。 */
  const analyzeInto = useCallback(async (nodeId: string, name: string, mid: string) => {
    const api = (window as unknown as {
      api?: {
        llmChat?: (p: string, o?: { model?: string; noThink?: boolean; maxTokens?: number }) => Promise<{ ok: boolean; text?: string; reason?: string }>
        llmChatStream?: (p: string, o: { model?: string; noThink?: boolean; maxTokens?: number }, cb: (part: { kind: 'text' | 'thought'; text: string }) => void) => Promise<{ ok: boolean; text?: string; reason?: string }>
      }
    }).api
    const patch = (p: { text?: string; thought?: string; thinking?: boolean }) =>
      setMsgs((m) => m.map((x) => (x.id === mid ? { ...x, ...p } : x)))
    patch({ text: `正在分析「${name}」…`, thinking: true })
    const s = await analyzeDocNode(nodeId, { api }, (part) => {
      patch({ text: part.accText.trim() || '（思考中…）', thought: part.accThought || undefined })
    })
    patch({ text: `【${name}】\n${s}`, thinking: false })
  }, [])

  /** ★ 文档分析：命中 analyze 时——0 文档提示；1 个直接分析；多个出选择菜单；名称匹配自动选 */
  const runAnalyze = useCallback(async (text: string) => {
    const mid = 'a' + Date.now().toString(36)
    setMsgs((m) => [...m, { id: mid, role: 'assistant', text: '正在读取画布数据并分析…' }])
    const patch = (p: { text?: string; thought?: string; docChoices?: ChatMsg['docChoices'] }) =>
      setMsgs((m) => m.map((x) => (x.id === mid ? { ...x, ...p } : x)))
    const api = (window as unknown as { api?: { llmChat?: (p: string) => Promise<{ ok: boolean; text?: string; reason?: string }> } }).api
    const docs = listDocNodes()
    /* 名称匹配：用户句子里提到的节点名/文档名（去掉常见虚词后包含匹配） */
    const clean = text.replace(/(分析下|分析|总结一下|总结|一下|内容|的|请|帮我)/g, '').trim()
    const named = clean ? docs.filter((d) => d.name.includes(clean) || d.template.includes(clean) || clean.includes(d.name)) : []
    const target = named.length === 1 ? named[0] : null
    if (!api?.llmChat) {
      const digest = await collectCanvasDigest()
      patch({ text: `画布节点概况：\n${digest.slice(0, 600)}` })
      return
    }
    if (docs.length === 0) {
      patch({ text: '画布上还没有文档节点——先导入一个 docx（拖文件或双击卡片导入），我就能帮你分析内容。' })
      return
    }
    if (target) { await analyzeInto(target.id, target.name, mid); return }
    if (docs.length === 1) { await analyzeInto(docs[0].id, docs[0].name, mid); return }
    /* 多个文档节点：出选择菜单（带各节点按钮） */
    patch({
      text: `画布上有 ${docs.length} 个文档节点，要分析哪一个？`,
      docChoices: docs.map((d) => ({ nodeId: d.id, label: `📄 ${d.name}${d.template ? '（' + d.template + '）' : ''}` })),
    })
  }, [analyzeInto])

  /** ★ 执行选定文档的深度分析（选择菜单点击后） */
  const runDocChoice = useCallback(async (nodeId: string, mode?: 'analyze' | 'open') => {
    const api = (window as unknown as { api?: { llmChat?: (p: string) => Promise<{ ok: boolean; text?: string; reason?: string }>; openExternalFile?: (name: string, data: Uint8Array) => Promise<string> } }).api
    const name = listDocNodes().find((d) => d.id === nodeId)?.name ?? '文档'
    if (mode === 'open') {
      const { bytes } = (await getDocNodeBytes(nodeId)) ?? { bytes: null }
      if (!bytes?.length) { setMsgs((m) => [...m, { role: 'assistant', text: `「${name}」还没有导入文档内容。` }]); return }
      if (!api?.openExternalFile) { setMsgs((m) => [...m, { role: 'assistant', text: '打开功能不可用。' }]); return }
      await api.openExternalFile(name, bytes)
      setMsgs((m) => [...m, { role: 'assistant', text: `已用系统默认程序（WPS/Word）打开「${name}」。` }])
      return
    }
    setMsgs((m) => [...m, { role: 'user', text: `分析「${name}」` }])
    const mid = 'a' + Date.now().toString(36)
    setMsgs((m) => [...m, { id: mid, role: 'assistant', text: `正在分析「${name}」…` }])
    await analyzeInto(nodeId, name, mid)
  }, [analyzeInto])

  /** ★ 2026-11-07 用户「执行转换」：交给画布真正跑一遍输出节点（PDF生成），结果回气泡。
   *  不在这里直接调 runPdfMerge——运行属于画布职责（节点运行徽章/日志都在那边），
   *  对话框只发指令、收结果。 */
  const runWorkflowNow = useCallback((nodeId?: string, nodeLabel?: string) => {
    const tip = nodeLabel ? `正在执行「${nodeLabel}」…` : '正在执行「PDF生成」…'
    setMsgs((m) => [...m, { role: 'assistant', text: tip }])
    const onRes = (ev: Event) => {
      const d = (ev as CustomEvent<{ ok: boolean; msg: string }>).detail
      window.removeEventListener('chat-run-result', onRes as EventListener)
      setMsgs((m) => [...m.filter((x) => x.text !== tip), { role: 'assistant', text: (d?.ok ? '✅ ' : '⚠️ ') + (d?.msg ?? '执行结束（无返回）') }])
    }
    window.addEventListener('chat-run-result', onRes as EventListener)
    /* 兜底摘监听：跑大文档可能几分钟，超时也不该永久挂着 */
    setTimeout(() => window.removeEventListener('chat-run-result', onRes as EventListener), 300000)
    window.dispatchEvent(new CustomEvent('chat-run-node', { detail: nodeId ? { nodeId } : {} }))
  }, [])

  /** ★ 2026-11-07 用户「N071节点」：按编号回答这个节点是什么，并在画布上定位它。
   *  编号来自节点标题 [N071]，用户习惯直接拿编号说话，不该回「这是什么意思」。 */
  const runNodeInfo = useCallback((text: string) => {
    setMsgs((m) => [...m, { role: 'assistant', text: '正在查看节点…' }])
    const hit = resolveNodeRef(text)
    const swap = (t: string) => setMsgs((m) => m.map((x) => (x.text === '正在查看节点…' ? { ...x, text: t } : x)))
    if (!hit) {
      swap(`画布上找不到这个编号的节点。现有节点：\n${listNodeRefs()}`)
      return
    }
    const desc = describeNodeById(hit.id)
    swap(desc ?? `节点「${hit.no}」读取失败（可能刚被删除）`)
    /* 画布侧：选中并居中，让用户一眼看到说的是哪一个 */
    window.dispatchEvent(new CustomEvent('chat-focus-node', { detail: { nodeId: hit.id } }))
  }, [])

  /** ★ 2026-11-07 用户「保存画布文件」→ 对话回了「已合并 1 份 / 3 页」：
   *  保存是画布动作，不该走节点执行。这里派发 chat-save-canvas，
   *  由 App 走工具栏同一条保存链路（选位置 → 清单确认 → 落盘），
   *  并把每个阶段的结果回填气泡（取消/已调起面板/已保存/失败）。 */
  const runSaveCanvas = useCallback(() => {
    const tip = '正在保存画布…'
    setMsgs((m) => [...m, { role: 'assistant', text: tip }])
    let first = true
    const onRes = (ev: Event) => {
      const d = (ev as CustomEvent<{ ok: boolean; msg: string; stage?: 'cancel' | 'open' | 'done' | 'error' }>).detail
      const text = (d?.ok ? (d.stage === 'done' ? '✅ ' : '') : '⚠️ ') + (d?.msg ?? '保存结束（无返回）')
      setMsgs((m) => (first ? m.map((x) => (x.text === tip ? { ...x, text } : x)) : [...m, { role: 'assistant', text }]))
      first = false
      if (d?.stage === 'done' || d?.stage === 'cancel' || d?.stage === 'error') {
        window.removeEventListener('chat-save-result', onRes as EventListener)
      }
    }
    window.addEventListener('chat-save-result', onRes as EventListener)
    /* 面板挂着等用户点确认是正常的，给足 10 分钟；点完/取消会提前摘监听 */
    setTimeout(() => window.removeEventListener('chat-save-result', onRes as EventListener), 600000)
    window.dispatchEvent(new Event('chat-save-canvas'))
  }, [])

  /** ★ 2026-11-08 用户「word转PDF」曾被模型瞎猜成「图片转Word」建错节点：
   *  能力早已有（右键 OO 文档「转为PDF节点」，x2t 秒转），这里接对话层——
   *  选目标节点（编号引用 > 画布唯一 OO 文档）→ 派 chat-to-pdf-node 走右键同一条链路，
   *  结果由 chat-topdf-result 回填气泡。没有文档节点就引导，不瞎建。 */
  const runWordToPdf = useCallback((text: string) => {
    const docKinds = ['ooword', 'oosheet']
    const say = (t: string) => setMsgs((m) => [...m, { role: 'assistant', text: t }])
    const ref = resolveNodeRef(text)
    if (ref && !docKinds.includes(ref.kind)) {
      say(`「${ref.no}」是「${ref.name}」节点（${ref.kind === 'pdf' || ref.kind === 'pdfmerge' ? '本身就是 PDF' : '不是 Word/Excel 文档'}），只有 Word/Excel 文档节点能转 PDF。`)
      return
    }
    let target = ref
    if (!target) {
      const st = useBaseStore.getState()
      const docs = Object.values(st.tables).filter((t) => t && !t.deleted && docKinds.includes(t.kind))
      if (docs.length > 1) {
        const names = docs.map((t) => {
          const no = String((t.config as { nodeNo?: string } | undefined)?.nodeNo ?? '')
          return `${no || '（无编号）'} ${t.name}`
        }).join('、')
        say(`画布上有 ${docs.length} 个文档节点（${names}）。要转哪一个？说「N0xx 转PDF」我就转那个。`)
        return
      }
      if (docs.length === 1) {
        const t0 = docs[0]
        const no = String((t0.config as { nodeNo?: string } | undefined)?.nodeNo ?? '')
        target = { id: t0.id, no, name: t0.name, kind: t0.kind }
      }
    }
    if (!target) {
      say('画布上还没有 Word/Excel 文档节点。先把 docx/xlsx 导入成「OO 文档」节点（插入页拖入，或双击节点导入），再转 PDF；也可以直接右键文档节点选「转为PDF节点」。')
      return
    }
    const tip = `正在把「${target.name}」转成 PDF…`
    setMsgs((m) => [...m, { role: 'assistant', text: tip }])
    const onRes = (ev: Event) => {
      const d = (ev as CustomEvent<{ ok: boolean; msg: string }>).detail
      window.removeEventListener('chat-topdf-result', onRes as EventListener)
      setMsgs((m) => [...m.filter((x) => x.text !== tip), { role: 'assistant', text: (d?.ok ? '✅ ' : '⚠️ ') + (d?.msg ?? '转换结束（无返回）') }])
    }
    window.addEventListener('chat-topdf-result', onRes as EventListener)
    /* 兜底摘监听：大文档 x2t 转换可能要一会儿，超时也不该永久挂着 */
    setTimeout(() => window.removeEventListener('chat-topdf-result', onRes as EventListener), 300000)
    window.dispatchEvent(new CustomEvent('chat-to-pdf-node', { detail: { nodeId: target.id } }))
  }, [])

  /** ★ 意图命中统一分流：analyze/weather 动作直接执行；其他出卡片 */
  const runIntent = useCallback((hit: ChatIntent, fileObjs: File[], text: string) => {
    /* ★ 纯回复意图（nodes 空且无动作）：直接文本回答，不出卡片 */
    if (!hit.nodes.length && !hit.action) {
      setTimeout(() => { setMsgs((m) => [...m, { role: 'assistant', text: hit.reply }]) }, 250)
      return
    }
    if (hit.action === 'analyze') {
      setTimeout(() => { void runAnalyze(text) }, 100)
      return
    }
    if (hit.action === 'weather') {
      setTimeout(() => { void runWeather(text) }, 100)
      return
    }
    if (hit.action === 'webSearch') {
      /* ★ 联网搜索 + 引用（2026-11-07 用户「看不到联网引用记录」）：Bing 抓取 → 本地模型带编号总结 →
         气泡末尾渲染「引用来源」链接列表（点击系统默认浏览器打开） */
      setTimeout(async () => {
        const api = (window as unknown as {
          api?: {
            webSearch?: (q: string) => Promise<{ ok: boolean; items?: Array<{ title: string; url: string; snippet: string }>; reason?: string }>
            llmChatStream?: (p: string, o: { model?: string; noThink?: boolean; maxTokens?: number }, cb: (part: { kind: 'text' | 'thought'; text: string }) => void) => Promise<{ ok: boolean; text?: string; reason?: string }>
            llmChat?: (p: string, o?: { model?: string; noThink?: boolean; maxTokens?: number }) => Promise<{ ok: boolean; text?: string; reason?: string }>
          }
        }).api
        if (!api?.webSearch) {
          setMsgs((m) => [...m, { role: 'assistant', text: '联网搜索不可用（缺少搜索模块）。' }])
          return
        }
        /* ★ 2026-11-08 用户截图：上文聊「根据黄历」，用户接「你联网查询」→ 旧版拿整句去搜，
           返回「你」字的词典解释。现在 extractWebQuery 剥净人称/指令词后为空 → 结合最近对话补全：
           本地模型提取搜索词（noThink 快）→ 启发式回退（最近用户消息剥引导词）→ 都没有就诚实澄清。 */
        let q = extractWebQuery(text)
        setMsgs((m) => { streamIdxRef.current = m.length; return [...m, { role: 'assistant', text: q ? `正在联网搜索「${q}」…` : '正在确认要查什么…' }] })
        const put = (patch: { text?: string; thought?: string; sources?: Array<{ title: string; url: string }> }) => setMsgs((m) => { const c = [...m]; const i = streamIdxRef.current; if (i >= 0 && i < c.length) c[i] = { ...c[i], ...patch }; return c })
        if (!q) {
          const turns = msgsRef.current.slice(-7).map((m) => ({ role: m.role, text: m.text }))
          turns.push({ role: 'user', text })
          /* ① 本地模型从上下文提取搜索词 */
          try {
            if (api.llmChat) {
              const ctx = turns.map((t) => `${t.role === 'user' ? '用户' : '助手'}：${t.text.slice(0, 120)}`).join('\n')
              const rp = await api.llmChat(`根据以下最近对话，判断用户现在想让联网查询什么。只输出搜索关键词本身（不超过12个字），不要任何解释和标点。\n\n${ctx}\n\n搜索关键词：`, { ...routeLlmOpt('query'), noThink: true, maxTokens: 256 })
              const kw = (rp.ok ? (rp.text ?? '') : '').split('\n')[0].replace(/["'「」『』。，,\s]/g, '').slice(0, 24)
              if (kw && isUsableQuery(kw)) q = kw
            }
          } catch { /* 模型不可用走启发式 */ }
          /* ② 启发式回退：最近用户消息剥引导词取主题 */
          if (!q) q = fallbackQueryFromTurns(turns)
          /* ③ 真没有可查的 → 澄清，不瞎搜 */
          if (!q) {
            put({ text: '想联网查什么？直接说「搜 XXX」就好，比如「搜 成都 黄历」。「XXX」就是要查的内容。' })
            return
          }
          put({ text: `正在联网搜索「${q}」…` })
        }
        const sr = await api.webSearch(q)
        if (!sr.ok || !sr.items?.length) {
          put({ text: `联网搜索失败：${sr.reason ?? '没有结果'}。可以直接把关键词再描述清楚一点试试。` })
          return
        }
        const items = sr.items
        put({ text: `搜到 ${items.length} 条结果，正在阅读总结…` })
        /* ★ 总结带上用户最近关心的事：搜「黄历」时结合「今天适合理发」语境作答，不再孤立解释词条 */
        const recentCtx = msgsRef.current.filter((m) => m.role === 'user').slice(-3).map((m) => m.text.slice(0, 60)).join(' / ')
        const sys = `以下是从 Bing 搜索「${q}」得到的结果：\n\n${items.map((it, i) => `[${i + 1}] ${it.title}\n${it.snippet}`).join('\n\n')}\n\n请基于以上搜索结果用中文回答用户的问题（300字内），关键信息后面标注来源编号（如 [1][2]）；搜索结果不足以回答的部分如实说明，禁止编造。\n\n用户的问题：${q}${recentCtx ? `\n（用户最近关心的事，回答时贴合这个语境：${recentCtx}）` : ''}`
        let accText = ''
        let accThought = ''
        /* 搜索总结同样关思考（noThink）：结果条目已给足上下文，思考只会拖慢并可能吃满预算 */
        const r = (api.llmChatStream
          ? await api.llmChatStream(sys, { ...routeLlmOpt('summarize'), noThink: true, maxTokens: 1024 }, (p) => {
              if (p.kind === 'thought') accThought += p.text
              else accText += p.text
              put({ text: accText.trim() || '正在总结…', thought: accThought || undefined })
            })
          : await api.llmChat?.(sys, { ...routeLlmOpt('summarize'), noThink: true, maxTokens: 1024 })) ?? { ok: false, reason: '缺少本地推理模块' }
        put({
          text: r.ok ? (r.text || '（模型没有返回内容，可重试）') : `搜到了结果，但模型总结失败：${r.reason}`,
          sources: items.map((it) => ({ title: it.title, url: it.url })),
        })
      }, 60)
      return
    }
    if (hit.action === 'open') {
      setTimeout(() => { void runDocOpen(text) }, 100)
      return
    }
    if (hit.action === 'nodeInfo') {
      setTimeout(() => { runNodeInfo(text) }, 100)
      return
    }
    if (hit.action === 'saveCanvas') {
      setTimeout(() => { runSaveCanvas() }, 100)
      return
    }
    if (hit.action === 'toPdf') {
      setTimeout(() => { runWordToPdf(text) }, 100)
      return
    }
    if (hit.action === 'run') {
      const ref = resolveNodeRef(text)
      setTimeout(() => { runWorkflowNow(ref?.id, ref ? `${ref.no} ${ref.name}` : undefined) }, 100)
      return
    }
    if (hit.action === 'addNode') {
      const alias = matchNodeAlias(text)
      /* ★ 2026-11-07 「N071 生成PDF」不要再建一个 PDF生成：说的就是那个编号的节点。
         同类型 → 直接跑它；不同类型 → 说清它是什么，别擅自改人家节点。 */
      const ref = resolveNodeRef(text)
      if (ref && alias) {
        const st = useBaseStore.getState()
        const refTable = st.tables[ref.id]
        if (refTable && refTable.kind === alias.kind) {
          setTimeout(() => {
            setMsgs((m) => [...m, { role: 'assistant', text: `「${ref.no}」就是「${ref.name}」节点，直接执行它（不新建）。` }])
            runWorkflowNow(ref.id, `${ref.no} ${ref.name}`)
          }, 150)
          return
        }
        if (refTable) {
          setTimeout(() => {
            setMsgs((m) => [...m, { role: 'assistant', text: `「${ref.no}」是「${ref.name}」节点（类型不是${alias.name}）。要新建${alias.name}，请说「加个${alias.name}节点」；要处理这个节点，直接说需求。` }])
          }, 150)
          return
        }
      }
      if (!alias) {
        setTimeout(() => { setMsgs((m) => [...m, { role: 'assistant', text: '想添加哪种节点？支持：PDF生成、Word输出、PDF瘦身、PDF转Word、替换页、文件夹、图片导入、Excel表格、Word文档、PDF文件、文本、规则、AI、参数库、章节工作簿。' }]) }, 250)
        return
      }
      setTimeout(() => {
        window.dispatchEvent(new CustomEvent('chat-add-node', { detail: { kind: alias.kind, name: alias.name } }))
        /* ★ 2026-11-07 用户「执行转换 没执行呢」：说「转换成PDF」是动作语义，建完就该跑；
           「加个XX节点」是纯搭建，不自动跑（避免打断搭流程）。判据看措辞，不看有没有"节点"二字。
           延迟 600ms：节点落库 → 连线写 linkedInputs → 上游 measurable，否则 runPdfMerge 读不到上游。 */
        const wantsRun = /(转|生成|导出|执行|开始|做)/.test(text) && !/(增加|添加|新建|创建|加)个?.{0,8}节点/.test(text)
        setMsgs((m) => [...m, { role: 'assistant', text: '已在画布创建「' + alias.name + '」节点' + (wantsRun ? '，开始执行。' : '，并连线到当前流程尾部。') }])
        if (wantsRun && alias.kind === 'pdfmerge') setTimeout(() => { runWorkflowNow() }, 600)
      }, 150)
      return
    }
    const key = hit.id + '~' + Date.now()
    if (fileObjs.length) pendingIntentFiles.current.set(key, fileObjs)
    setTimeout(() => {
      setMsgs((m) => [...m, { role: 'assistant', text: hit.reply, intent: { id: key, title: hit.title, plan: hit.nodes.map((n) => n.name).join(' → '), accepted: false } }])
    }, 300)
  }, [runAnalyze, runWeather, runWorkflowNow, runNodeInfo, runSaveCanvas])

  const send = useCallback(() => {
    if (!input.trim() && !files.length) return
    pushUserOnly(input.trim() || '（基于拖入的文件搭工作流）', files)
    const sentText = input.trim()
    const sentFileObjs = fileObjs
    /* ★ 规则匹配永远优先：明确的正则比 LLM 判断更快更准，LLM 只兜模糊表述 */
    const ruleHit = matchIntent(input, files)
    if (ruleHit) {
      runIntent(ruleHit, fileObjs, input)
      setInput('')
      setFiles([])
      setFileObjs([])
      if (taRef.current) taRef.current.style.height = '36px'
      return
    }
    /* ★ 2026-11-07 本地快速模式已删——规则没命中的，一律交给本机聪明模型（Qwen3.5-9B） */
    setInput('')
    setFiles([])
    setFileObjs([])
    if (taRef.current) taRef.current.style.height = '36px'
    setTimeout(async () => {
      /* ★ 流式气泡（2026-11-07 用户「不要结果好了再出」「看不到思考过程」）：
         思考段实时进「💭 思考过程」折叠块（思考中自动展开），正文逐字上屏；
         模型输出意图 JSON（以 { 开头）时正文不上屏（保持占位），等路由结果决定动作 */
      setMsgs((m) => { streamIdxRef.current = m.length; return [...m, { role: 'assistant', text: '（思考中…）', thinking: true }] })
      const put = (patch: { text?: string; thought?: string; thinking?: boolean; sources?: Array<{ title: string; url: string }> }) => setMsgs((m) => { const c = [...m]; const i = streamIdxRef.current; if (i >= 0 && i < c.length) c[i] = { ...c[i], ...patch }; return c })
      const turnList = msgsRef.current.map((m) => ({ role: m.role, text: m.text }))
      const ctx = buildContextBlock(turnList, sentText)
      /* 信息明显不足（没历史 + 极短追问）时让模型先问清，而不是硬猜 */
      const block = ctx.block + (needsMoreInfo(turnList, sentText) ? '【注意】用户这句话信息很少，先用一句话问清他要做什么，不要凭空猜测。\n' : '')
      const r = await smartChat(sentText, block, (st) => {
        const shown = st.textAcc.trimStart()
        put({ text: shown.startsWith('{') || !shown ? '（思考中…）' : shown, thought: st.thoughtAcc || undefined })
      })
      /* 意图命中：撤掉流式占位气泡，交给 runIntent（addNode 直接建节点，多节点的意图才出卡片） */
      if (r.intentHit) {
        setMsgs((m) => m.filter((_, i) => i !== streamIdxRef.current))
        streamIdxRef.current = -1
        runIntent(r.intentHit, sentFileObjs, sentText)
        return
      }
      put({ text: r.reply, thinking: false })
    }, 60)
  }, [input, files, fileObjs, pushUserOnly, smartChat, runIntent])

  /* 一键创建工作流：把意图+文件交给画布（Canvas 监听建节点/连线/导文件） */
  const acceptIntent = useCallback((intentId: string) => {
    setMsgs((m) => m.map((x) => (x.intent && x.intent.id === intentId ? { ...x, intent: { ...x.intent, accepted: true } } : x)))
    const fs = pendingIntentFiles.current.get(intentId) ?? []
    pendingIntentFiles.current.delete(intentId)
    window.dispatchEvent(new CustomEvent('chat-create-workflow', { detail: { intentId: intentId.split('~')[0], files: fs } }))
  }, [])
  const pendingIntentFiles = useRef(new Map<string, File[]>())

  /* ★ 文件进入 = WorkBuddy 式（2026-11-07 用户指令）：直接出文件气泡+意图自动分析，
     不再走输入栏小胶囊 */
  const ingestFiles = useCallback((fl: FileList | File[] | null) => {
    if (!fl || !('length' in fl) || !fl.length) return
    const arr: File[] = Array.from(fl)
    const names = arr.map((f) => f.name)
    setMsgs((m) => [...m, { role: 'user', text: '', files: names }])
    const hit = matchIntent('', names)
    if (hit) {
      const key = hit.id + '~' + Date.now()
      pendingIntentFiles.current.set(key, arr)
      setTimeout(() => {
        setMsgs((m) => [...m, { role: 'assistant', text: hit.reply, intent: { id: key, title: hit.title, plan: hit.nodes.map((n) => n.name).join(' → '), accepted: false } }])
      }, 300)
    } else {
      setTimeout(() => {
        setMsgs((m) => [...m, { role: 'assistant', text: `已收到 ${arr.length} 个文件（${names.join('、')}）。说说要对它做什么？比如「转成 word」「合并成一个」` }])
      }, 300)
    }
  }, [])

  const taGrow = (el: HTMLTextAreaElement) => {
    /* 只按显式换行符计数——1 行恒 36px，长句折行内部滚动 */
    const lines = Math.max(1, Math.min(el.value.split('\n').length, 5))
    el.style.height = (lines * 20 + 16) + 'px'
  }

  /* ★ 双主题配色：浅色（GitHub Light 风，白画布）/ 暗色（玻璃风） */
  const T = isLight ? {
    panelBg: 'rgba(255,255,255,0.97)', panelBorder: '#d9dfe7', shadow: '0 12px 40px rgba(30,40,60,0.16)',
    titleText: '#3a4149', bodyText: '#333a42', subText: '#5a6572',
    bubbleUser: 'rgba(77,125,253,0.1)', bubbleUserBorder: 'rgba(77,125,253,0.3)',
    bubbleAi: '#f2f5f8', bubbleAiBorder: '#e3e8ee',
    chipBg: 'rgba(77,125,253,0.08)', chipBorder: 'rgba(77,125,253,0.3)', chipText: '#3a6fd8',
    inputBg: '#f5f7fa', inputBorder: '#e3e8ee', inputText: '#2a3138',
    btnBg: '#ffffff', btnBorder: '#d5dbe3', btnColor: '#5a6572',
    sep: '#e8ecf1', triBg: '#ffffff', triBorder: '#c5ccd6', triShadow: '0 4px 16px rgba(30,40,60,0.18)',
  } : {
    panelBg: 'rgba(28,30,36,0.96)', panelBorder: 'rgba(120,140,180,0.3)', shadow: '0 10px 34px rgba(0,0,0,0.45)',
    titleText: '#e6ecf5', bodyText: '#dfe7f3', subText: '#9fb4d8',
    bubbleUser: 'rgba(77,125,253,0.22)', bubbleUserBorder: 'rgba(110,160,255,0.35)',
    bubbleAi: 'rgba(120,140,180,0.12)', bubbleAiBorder: 'rgba(120,140,180,0.22)',
    chipBg: 'rgba(110,170,255,0.16)', chipBorder: 'rgba(110,170,255,0.4)', chipText: '#aecdf5',
    inputBg: 'transparent', inputBorder: 'transparent', inputText: '#e6ecf5',
    btnBg: 'rgba(120,140,180,0.12)', btnBorder: 'rgba(120,140,180,0.35)', btnColor: '#9fb4d8',
    sep: 'rgba(120,140,180,0.16)', triBg: 'rgba(120,140,180,0.12)', triBorder: 'rgba(120,140,180,0.35)', triShadow: '0 4px 14px rgba(0,0,0,0.3)',
  }

  /* ★ 2026-11-07 用户「右键能引用作为提示词工程的一部分」：
     右键瞬间 DOM 选区可能已被 mousedown 清掉（不同内核/DOM 结构下都会出现），
     只读 getSelection() 会拿到空串 → 菜单退化成「整条消息」，用户选的那句白选。
     这里持续记住最后一次非空选区（连同选区起点所在节点），
     右键时据此判断「这次右键是不是冲着我刚选的那段来的」。 */
  const lastSelRef = useRef<{ text: string; node: Node | null; rect: { left: number; top: number; right: number; bottom: number; width: number } | null }>({ text: '', node: null, rect: null })
  useEffect(() => {
    const onSel = () => {
      const s = window.getSelection()
      if (!s || s.isCollapsed) return
      const t = s.toString().trim()
      if (!t) return
      /* ★ 记住选区矩形：右键菜单据此贴着那段文字弹，而不是跟着鼠标乱跑 */
      let rect: { left: number; top: number; right: number; bottom: number; width: number } | null = null
      try {
        const r = s.rangeCount > 0 ? s.getRangeAt(0).getBoundingClientRect() : null
        if (r && (r.width > 0 || r.height > 0)) rect = { left: r.left, top: r.top, right: r.right, bottom: r.bottom, width: r.width }
      } catch { /* 跨节点选区偶发抛错，忽略 */ }
      lastSelRef.current = { text: t, node: s.anchorNode, rect }
    }
    document.addEventListener('selectionchange', onSel)
    return () => document.removeEventListener('selectionchange', onSel)
  }, [])

  /* 菜单开着时按 Esc 关掉（右键菜单的通用习惯） */
  useEffect(() => {
    if (!ctxMenu) return
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') setCtxMenu(null) }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [ctxMenu])

  /* 展开三段：closed(圆钮) → pill(胶囊向右延伸) → grow(向上伸展至窗体) → open */
  const expand = useCallback(() => {
    setContentFade(false) /* ★ 展开时清除折叠时的淡出标记（否则内容永久 opacity 0 不显示） */
    setPhase('pill')
    setTimeout(() => setPhase('grow'), 380)
    setTimeout(() => setPhase('open'), 780)
  }, [])

  /* ★ 2026-11-07 用户「选中右键还是没菜单」：打开「复制/引用」菜单（气泡右键、面板空白右键共用）。
     取值优先级：右键当下的选区 → 最近一次选区快照（须在同一作用域内）→ fallback。
     同时 ctxHandled() 让主进程跳过本次系统菜单，避免两层菜单叠加。
     scope 传气泡元素时，只认落在该气泡内的快照，防止引用到别的消息。 */
  const openQuoteMenu = useCallback((e: React.MouseEvent, fallback = '', scope?: HTMLElement | null) => {
    e.preventDefault()
    e.stopPropagation()
    const sel = window.getSelection()
    const live = sel?.toString().trim() ?? ''
    const snap = lastSelRef.current
    const box = scope ?? panelRef.current
    const inScope = !!snap.node && !!box?.contains(snap.node)
    const picked = live || (inScope ? snap.text : '')
    const text = (picked || fallback).trim()
    if (!text) return /* 空白处又没选任何东西：不弹菜单（也别让画布菜单接手） */
    /* ★ 锚点（2026-11-07「右键复制位置错位」）：
       选区矩形 → 快照矩形 → 气泡矩形 → 鼠标点，逐级降级。
       菜单永远贴着「你要操作的那段内容」，不再跟着鼠标飞到面板别处。 */
    let anchor: { left: number; top: number; right: number; bottom: number; width: number } | null = null
    if (live && sel && sel.rangeCount > 0) {
      try {
        const r = sel.getRangeAt(0).getBoundingClientRect()
        if (r.width > 0 || r.height > 0) anchor = { left: r.left, top: r.top, right: r.right, bottom: r.bottom, width: r.width }
      } catch { /* ignore */ }
    }
    if (!anchor && picked && inScope && snap.rect) anchor = snap.rect
    if (!anchor && scope) {
      const r = scope.getBoundingClientRect()
      if (r.width > 0 || r.height > 0) anchor = { left: r.left, top: r.top, right: r.right, bottom: r.bottom, width: r.width }
    }
    ;(window as unknown as { api?: { ctxHandled?: () => void } }).api?.ctxHandled?.()
    setCtxMenu({ x: e.clientX, y: e.clientY, anchor, text, fromSel: !!picked })
  }, [])

  /* 菜单定位：先按锚点算一个方向，再用真实尺寸夹进视口（不猜宽高，避免溢出屏幕） */
  useLayoutEffect(() => {
    if (!ctxMenu) { setMenuPos(null); return }
    const el = menuRef.current
    const w = el?.offsetWidth ?? 200
    const h = el?.offsetHeight ?? 96
    const GAP = 6
    const EDGE = 8
    const a = ctxMenu.anchor
    let left = ctxMenu.x
    let top = ctxMenu.y + GAP
    if (a) {
      left = a.left + a.width / 2 - w / 2
      top = a.bottom + GAP
      if (top + h > window.innerHeight - EDGE) top = Math.max(EDGE, a.top - h - GAP)
    }
    left = Math.min(Math.max(EDGE, left), Math.max(EDGE, window.innerWidth - w - EDGE))
    top = Math.min(Math.max(EDGE, top), Math.max(EDGE, window.innerHeight - h - EDGE))
    setMenuPos((p) => (p && p.left === left && p.top === top ? p : { left, top }))
  }, [ctxMenu])

  /* ★ 2026-11-07 用户「右键能引用作为提示词工程的一部分」：
     把选中的原文插进输入框——插在光标处（用户通常要在这段前后补自己的话），
     不再一律追加到末尾。 */
  const insertQuote = useCallback((raw: string) => {
    const txt = String(raw ?? '').replace(/\s+$/, '').trim()
    if (!txt) return
    const ta = taRef.current
    const focused = !!ta && document.activeElement === ta
    const pos = focused ? (ta!.selectionStart ?? null) : null
    setInput((v) => {
      const piece = '「' + txt + '」'
      if (pos == null || pos > v.length) return (v ? v.replace(/\s+$/, '') + '\n' : '') + piece
      const before = v.slice(0, pos)
      const after = v.slice(pos)
      return before + (before && !/\s$/.test(before) ? ' ' : '') + piece + after
    })
    setTimeout(() => {
      const el = taRef.current
      if (!el) return
      el.focus()
      const end = el.value.length
      try { el.setSelectionRange(end, end) } catch { /* ignore */ }
      taGrow(el)
    }, 20)
  }, [])

  /* ★ 2026-11-07 主进程右键菜单「引用到输入框（作为提示词）」→ 插进输入框并聚焦；
     面板收起时先展开，否则用户看不到引用落到哪了。 */
  useEffect(() => {
    const api = (window as unknown as { api?: { onQuoteText?: (cb: (t: string) => void) => (() => void) } }).api
    if (!api?.onQuoteText) return
    const off = api.onQuoteText((t) => {
      if (phase === 'closed') expand()
      insertQuote(t)
    })
    return off
  }, [phase, expand, insertQuote])

  /* 收起柔化：内容先淡出 0.18s，再开始缩回 */
  const collapse = useCallback(() => {
    setContentFade(true)
    setTimeout(() => {
      setPhase('grow')
      setTimeout(() => setPhase('pill'), 380)
      setTimeout(() => setPhase('closed'), 780)
    }, 180)
  }, [])

  const collapsed = phase === 'closed'
  const H = window.innerHeight
  /* ★ 功能区（.cr-ribbon）展开时让位：面板顶部 = 功能区底边 + 18（与左/下同距）。
     功能区收起（高度 0/隐藏）时忽略它，顶部间距即 18。
     ★ 坐标系：父容器用 .react-flow 的 parentElement（渲染时 DOM 已确定，
     不依赖 panelRef 时序——ref 在首次渲染为 null 会把面板推出屏幕外，实测 top=1237） */
  const rib = typeof document !== 'undefined' ? document.querySelector('.cr-ribbon') : null
  const parentEl = (document.querySelector('.react-flow')?.parentElement as HTMLElement | null) ?? null
  const parentTop = parentEl ? parentEl.getBoundingClientRect().top : 0
  const parentH = parentEl ? parentEl.offsetHeight : H
  const ribRect = rib ? rib.getBoundingClientRect() : null
  const ribVisible = !!(ribRect && ribRect.height > 1 && ribRect.bottom > 0)
  const ribBottom = ribVisible ? ribRect.bottom : parentTop
  const safeTop = Math.max(18, Math.round(ribBottom - parentTop) + 18)
  const xSide: React.CSSProperties = dock === 'l' ? { left: 18 } : { right: 18 }
  /* 面板宽度：用户拖过用固定 px，否则默认 min(600px, 56%)（原 460px/44% 太窄，意图卡片被压成一列） */
  const dockW: number | string = panelW || 'min(600px, 56%)'
  const frame: React.CSSProperties = phase === 'closed'
    ? { ...xSide, bottom: 27, width: 34, height: 34, borderRadius: 17 }
    : phase === 'pill'
      ? { ...xSide, bottom: 27, width: dockW, height: 34, borderRadius: 17 }
      : stretch
        ? { ...xSide, top: safeTop, width: dockW, height: parentH - safeTop - 18, borderRadius: 14 }
        : { ...xSide, bottom: 18, width: dockW, maxHeight: parentH - safeTop - 18, borderRadius: 14 }

  /* ★ 拖左右边缘改宽度（面板贴左时拖右缘，贴右时拖左缘）；松手记忆到 localStorage */
  const startResize = useCallback((e: React.PointerEvent) => {
    e.preventDefault()
    e.stopPropagation()
    const el = e.currentTarget as HTMLElement
    el.setPointerCapture(e.pointerId)
    const startX = e.clientX
    const startW = panelRef.current?.getBoundingClientRect().width ?? 460
    const dir = dock === 'l' ? 1 : -1
    let latest = Math.round(startW)
    const onMove = (ev: PointerEvent) => {
      latest = Math.max(340, Math.min(900, Math.round(startW + (ev.clientX - startX) * dir)))
      setPanelW(latest)
    }
    const onUp = () => {
      el.removeEventListener('pointermove', onMove)
      el.removeEventListener('pointerup', onUp)
      try { localStorage.setItem('cb-dock-w', String(latest)) } catch { /* 隐私模式 */ }
    }
    el.addEventListener('pointermove', onMove)
    el.addEventListener('pointerup', onUp)
  }, [dock])

  /* 拖动：标题栏按住，鼠标在窗口左半→停左、右半→停右（沿边停靠，上下位置不变）
     ★ 2026-11-07「拖的时候卡住了」：全程直接写 DOM，不走 React 状态；
       · 位移超过 5px 才算真拖（纯点击仍可正常选字/打字）；
       · 真拖后：body 加 .cb-dock-dragging（全局禁选 + grabbing）、清掉已有选区；
       · 落点幽灵框 = 面板自己的 getBoundingClientRect（同 left/top/宽/高）——
         所以提示框和实际框体严格对齐，不再靠 safeTop/offsetHeight 另算一套；
       · stopPropagation：不让画布把这次 pointerdown 当成框选/平移起点。 */
  const startDrag = useCallback((e: React.PointerEvent) => {
    e.stopPropagation()
    const panel = panelRef.current
    if (!panel) return
    const startX = e.clientX
    const startY = e.clientY
    /* ★ 2026-11-07「定位和停靠位置还是不一致，没考虑侧边栏」：
       不再自己算 left/right 像素（侧栏、画布缩放、容器内边距都会让算法跑偏）。
       幽灵框和面板**同一个包含块 + 同一套定位规则**（left/right:18），
       宽高/纵向位置直接抄面板的 offset 值 —— 浏览器用同一个基准计算，天然对齐。 */
    const op = (panel.offsetParent as HTMLElement | null) ?? document.body
    const opW = op.clientWidth
    const baseLeft = panel.offsetLeft
    const pw = panel.offsetWidth
    const ph = panel.offsetHeight
    const pt = panel.offsetTop
    /* 跟手范围：容器左侧 18 到右侧 18（以前死死夹在 ±120px，用户「拖一点就拖不动了」就是这个） */
    const minDx = 18 - baseLeft
    const maxDx = opW - 18 - pw - baseLeft
    const ghost = ghostRef.current
    const label = ghostLabelRef.current
    let active = false
    let cancelled = false
    let side: 'l' | 'r' = dock
    let dx = 0

    /* 直接写 DOM：不 setState（否则整个面板每帧重渲染 → 卡）；
       也不套 requestAnimationFrame（窗口失焦/不可见时 rAF 会被节流甚至暂停，
       那样拖起来就像「卡住」）。pointermove 本身已由浏览器按帧合并，开销极低。 */
    const paint = () => {
      if (ghost) {
        ghost.style.display = 'block'
        ghost.style.top = `${pt}px`
        ghost.style.width = `${pw}px`
        ghost.style.height = `${ph}px`
        ghost.style.left = side === 'l' ? '18px' : 'auto'
        ghost.style.right = side === 'r' ? '18px' : 'auto'
      }
      if (label) label.textContent = side === 'l' ? '← 松手停靠左侧' : '松手停靠右侧 →'
      panel.style.transform = `translateX(${dx}px)`
    }

    const cleanup = () => {
      window.removeEventListener('pointermove', move)
      window.removeEventListener('pointerup', up)
      window.removeEventListener('pointercancel', up)
      window.removeEventListener('keydown', onKey, true)
      window.removeEventListener('blur', up)
      document.body.classList.remove('cb-dock-dragging')
      if (ghost) ghost.style.display = 'none'
      /* 松手直接落位：先关过渡再清位移，避免「先回旧侧、再滑到新侧」的二段跳。
         之后由 setDock 改 left/right（同一时刻已在落点上，视觉零跳动）。 */
      panel.style.transition = 'none'
      panel.style.transform = ''
      panel.style.boxShadow = ''
      panel.style.cursor = ''
      window.setTimeout(() => { if (panelRef.current === panel) panel.style.transition = '' }, 32)
      active = false
    }
    const move = (ev: PointerEvent) => {
      if (!active) {
        if (Math.abs(ev.clientX - startX) < 5 && Math.abs(ev.clientY - startY) < 5) return
        active = true
        document.body.classList.add('cb-dock-dragging')
        window.getSelection()?.removeAllRanges()
        panel.style.boxShadow = '0 22px 52px rgba(0,0,0,0.36)'
        panel.style.cursor = 'grabbing'
      }
      /* 目标侧按「面板容器内的中点」判定，不再用 window.innerWidth（有侧栏时中点会偏） */
      side = ev.clientX < (op.getBoundingClientRect().left + opW / 2) ? 'l' : 'r'
      /* 跟手不再夹在 ±120：限制在容器内可停靠的范围（拖到边即止，不是拖一点就锁死） */
      dx = Math.max(minDx, Math.min(maxDx, ev.clientX - startX))
      paint()
    }
    /* Esc 取消本次换边（面板回原位）；pointercancel / 窗口失焦同样收尾，
       避免 body 一直挂着 .cb-dock-dragging 把整窗搞成不可选（这也是「卡住」的来源之一）。 */
    const onKey = (ev: KeyboardEvent) => {
      if (ev.key !== 'Escape') return
      cancelled = true
      ev.stopPropagation()
      cleanup()
    }
    const up = () => {
      const wasActive = active
      cleanup()
      if (wasActive && !cancelled) {
        setDock(side)
        try { localStorage.setItem('cb-dock-side', side) } catch { /* 隐私模式 */ }
      }
    }
    window.addEventListener('pointermove', move)
    window.addEventListener('pointerup', up)
    window.addEventListener('pointercancel', up)
    window.addEventListener('keydown', onKey, true)
    window.addEventListener('blur', up)
  }, [dock])

  /* 折叠态：▶ 圆钮（向右三角） */
  if (collapsed) {
    return (
      <div
        onContextMenu={(e) => { e.stopPropagation() }}
      onDragOver={(e) => { e.preventDefault(); setDragOver(true) }}
        onDragLeave={() => setDragOver(false)}
        onDrop={(e) => { e.preventDefault(); setDragOver(false); ingestFiles(e.dataTransfer?.files ?? null) }}
        onClick={expand}
        title="展开工作流对话（可拖文件到这里）"
        style={{
          position: 'absolute', bottom: 27, zIndex: 250,
          ...(dock === 'l' ? { left: 18 } : { right: 18 }),
          width: 34, height: 34, borderRadius: 17, cursor: 'pointer',
          background: dragOver ? 'rgba(77,125,253,0.15)' : T.triBg,
          border: `1px solid ${dragOver ? 'rgba(77,125,253,0.6)' : T.triBorder}`, color: T.subText, fontSize: 13,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          boxShadow: T.triShadow,
          transition: 'background 0.15s, border-color 0.15s',
        }}
      >▶</div>
    )
  }

  const showBody = phase === 'grow' || phase === 'open'
  /* 圆角矩形按钮（OO 白色工具栏统一风格） */
  const sqBtn: React.CSSProperties = {
    width: 28, height: 28, fontSize: 12, borderRadius: 8,
    background: T.btnBg, border: `1px solid ${T.btnBorder}`, color: T.btnColor,
    cursor: 'pointer', flexShrink: 0, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    transition: 'border-color 0.12s, color 0.12s',
  }
  const contentOpacity = (phase === 'pill' || contentFade) ? 0 : 1

  /* ★ 2026-11-07 用户反馈「点了复制，Ctrl+V 还是旧内容（旧截图）」：
     navigator.clipboard.writeText 在自定义协议 / 窗口未聚焦时会静默失败（原代码 catch 后什么也没做）。
     优先走主进程 clipboard.writeText（同步、无权限门槛），再退 DOM 选区复制。 */
  const copyPlainText = (t: string) => {
    const api = (window as unknown as { api?: { clipboardWrite?: (s: string) => Promise<boolean> } }).api
    const execCopy = () => {
      try {
        const ta = document.createElement('textarea')
        ta.value = t
        ta.style.cssText = 'position:fixed;left:-9999px;top:0;opacity:0'
        document.body.appendChild(ta)
        ta.select()
        document.execCommand('copy')
        document.body.removeChild(ta)
      } catch { /* 复制失败：用户仍可用「引用到输入框」 */ }
    }
    if (api?.clipboardWrite) { void api.clipboardWrite(t); return }
    if (navigator.clipboard?.writeText) { void navigator.clipboard.writeText(t).catch(execCopy); return }
    execCopy()
  }

  return (
    <>
    <div
      ref={panelRef}
      className="cb-dock-panel"
      style={{
        position: 'absolute', zIndex: 250,
        display: 'flex', flexDirection: 'column',
        background: T.panelBg, backdropFilter: 'blur(10px)',
        border: `1px solid ${dragOver ? 'rgba(77,125,253,0.65)' : T.panelBorder}`,
        boxShadow: T.shadow,
        overflow: 'visible', /* ★ 菜单向上弹出不被裁（圆角由面板自身背景绘制） */
        transition: 'width 0.36s cubic-bezier(0.4,0,0.2,1), max-height 0.38s cubic-bezier(0.4,0,0.2,1), border-radius 0.3s ease, border-color 0.15s, left 0.3s cubic-bezier(0.4,0,0.2,1), right 0.3s cubic-bezier(0.4,0,0.2,1), transform 0.24s cubic-bezier(0.4,0,0.2,1)',
        ...frame,
      }}
      onContextMenu={(e) => {
        /* ★ 2026-11-07 用户「选中右键还是没菜单」：截图里弹的是画布的「全选/删除/复制/置于顶层」，
           说明这次右键冒泡到了 .rf-canvas，画布菜单抢走了它。
           面板内除输入控件（交给系统菜单）外：有选中文字就弹「复制/引用」，
           没有就什么都不弹——总之不让画布菜单出现在对话框上。 */
        const t = e.target as HTMLElement
        if (t.closest('textarea, input, [contenteditable="true"]')) return
        openQuoteMenu(e)
      }}
      onPointerDown={(e) => {
        /* ★ 分级响应（2026-11-07 修「有时没法打字」）：
           按钮/输入框 → 无操作；气泡内 → 允许文字选择；模型菜单内 → 不关不聚焦（否则 pointerdown 移除菜单 DOM，click 丢失）；
           其他空白 → 聚焦输入框 + 可拖动换边 */
        const t = e.target as HTMLElement
        if (t.tagName === 'TEXTAREA') return
        if (t.closest('[data-model-menu]')) return
        const inBtn = !!(t.tagName === 'BUTTON' || t.closest('button'))
        if (modelMenu && !inBtn) setModelMenu(false)
        if (inBtn) return
        if (t.closest('[data-bubble]')) return
        taRef.current?.focus()
        startDrag(e)
      }}
      onDragOver={(e) => { e.preventDefault(); setDragOver(true) }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => { e.preventDefault(); setDragOver(false); ingestFiles(e.dataTransfer?.files ?? null) }}
      /* ★ 2026-10-31 去掉容器点击聚焦——它会打断气泡文字的选择（用户反馈「每次选择都退出」）。
         焦点只在点击输入区时落到 textarea；气泡文字可选中+右键复制/引用 */
      onPaste={(e) => {
        const dt = e.clipboardData
        if (!dt) return
        const text = dt.getData('text/plain')
        const ae = document.activeElement as HTMLElement | null
        const editing = !!ae && (ae.tagName === 'TEXTAREA' || ae.tagName === 'INPUT' || ae.isContentEditable)
        /* ★ 2026-11-07 用户反馈「复制了文字，Ctrl+V 却粘出图片」：
           先看剪贴板有没有文本——有文本就一定是「粘文字」，不再当文件摄入。
           编辑区（输入框）放行给浏览器原生粘贴，光标位置才对；非编辑区追加到输入框。 */
        if (text) {
          if (editing) return
          e.preventDefault()
          setInput((v) => (v ? v + text : text))
          setTimeout(() => taRef.current?.focus(), 0)
          return
        }
        /* 剪贴板里没有文本 → 只可能是图片/文件（截图工具、资源管理器复制），继续走摄入 */
        const fs: File[] = []
        for (let i = 0; i < dt.items.length; i++) {
          const it = dt.items[i]
          if (it.kind === 'file') {
            const f = it.getAsFile()
            if (f) {
              const name = /^image\./.test(f.name) ? ('截图-' + new Date().toTimeString().slice(0, 8).replace(/:/g, '') + '.' + (f.type.split('/')[1] || 'png')) : f.name
              fs.push(new File([f], name, { type: f.type }))
            }
          }
        }
        if (fs.length) { e.preventDefault(); ingestFiles(fs) }
      }}
    >
      {/* 顶部方形按钮排（OO 白色工具栏风格） */}
      {showBody && (
        <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '7px 10px', borderBottom: `1px solid ${T.sep}`, flexShrink: 0, opacity: contentOpacity, transition: 'opacity 0.18s ease', cursor: 'grab' }}>
          <span title="按住拖动 → 切换左右停靠" style={{ color: T.subText, fontSize: 12, letterSpacing: 2, opacity: 0.6 }}>⋮⋮</span>
          <b style={{ color: T.titleText, fontSize: 12 }}>工作流</b>
          <span style={{ flex: 1 }} />
          <button
            title={stretch ? '恢复常规高度' : '对话栏向上拉伸对齐上方边界'}
            onClick={() => setStretch((v) => !v)}
            style={{ ...sqBtn, borderColor: stretch ? '#4d7dfd' : T.btnBorder, color: stretch ? '#4d7dfd' : T.btnColor }}
          >⬆</button>
          <button title="选择工程文件" onClick={() => fileRef2.current?.click()} style={sqBtn}><FolderIcon /></button>
          <input ref={fileRef2} type="file" multiple style={{ display: 'none' }} onChange={(e) => { ingestFiles(e.target.files); e.target.value = '' }} />
          <button title="最小化（折叠对话框）" onClick={collapse} style={sqBtn}>—</button>
        </div>
      )}

      {/* 对话流（向上生长；超高内部滚动） */}
      {showBody && msgs.length > 0 && (
        <div
          ref={flowRef}
          style={{
            flex: 1, minHeight: 60, overflowY: 'auto',
            display: 'flex', flexDirection: 'column', gap: 8,
            padding: '12px 14px 8px',
            opacity: contentOpacity, transition: 'opacity 0.18s ease',
          }}
        >
          {msgs.map((m, i) => (
            <div
              key={i}
              data-bubble="1"
              style={{ alignSelf: m.role === 'user' ? 'flex-end' : 'flex-start', maxWidth: '88%', background: m.role === 'user' ? T.bubbleUser : T.bubbleAi, border: `1px solid ${m.role === 'user' ? T.bubbleUserBorder : T.bubbleAiBorder}`, borderRadius: m.role === 'user' ? '12px 12px 3px 12px' : '12px 12px 12px 3px', padding: '8px 12px', color: T.bodyText, fontSize: 13, whiteSpace: 'pre-wrap', lineHeight: 1.55, userSelect: 'text', cursor: 'text' }}
              onContextMenu={(e) => {
                /* ★ 2026-11-07 用户「沟通的文字选中后能复制，右键也有菜单，且右键能引用」：
                   气泡右键一律由应用内菜单接管（可预测、不和系统菜单打架）；
                   文本优先级：右键瞬间的选区 → 刚选中的那段（快照，须属于本气泡）→ 整条消息。 */
                openQuoteMenu(e, m.text, e.currentTarget)
              }}
            >
              {m.files?.length ? <div style={{ fontSize: 11, color: '#4d7dfd', marginBottom: 4 }}>📎 {m.files.join('、')}</div> : null}
              {/* ★ 思考过程（流式；默认折叠，思考中展开直播） */}
              {m.thought ? (
                <details open={!!m.thinking} style={{ marginBottom: 6, fontSize: 12 }}>
                  <summary style={{ cursor: 'pointer', opacity: 0.65, userSelect: 'none' }}>💭 思考过程{m.thinking ? '（进行中…）' : `（${m.thought.length} 字）`}</summary>
                  <div
                    ref={(el) => {
                      /* 只有正在思考的那条挂 ref：流式更新时 useEffect 把它滚到最底（显示最新思考） */
                      if (m.thinking && el) thoughtBoxRef.current = el
                      else if (thoughtBoxRef.current === el) thoughtBoxRef.current = null
                    }}
                    style={{ marginTop: 4, padding: '6px 9px', borderRadius: 8, background: isLight ? 'rgba(0,0,0,0.045)' : 'rgba(255,255,255,0.06)', color: isLight ? '#6b7480' : '#9aa7b8', fontSize: 11.5, lineHeight: 1.5, maxHeight: 180, overflowY: 'auto', whiteSpace: 'pre-wrap' }}>{m.thought}</div>
                </details>
              ) : null}
              <span data-bubble-text="1">{m.text}</span>
              {/* ★ 联网引用来源：可点开（系统默认浏览器） */}
              {m.sources?.length ? (
                <div style={{ marginTop: 8, paddingTop: 6, borderTop: `1px solid ${isLight ? 'rgba(0,0,0,0.08)' : 'rgba(255,255,255,0.1)'}` }}>
                  <div style={{ fontSize: 11, opacity: 0.65, marginBottom: 4 }}>🔗 引用来源（{m.sources.length}）</div>
                  {m.sources.map((s, si) => (
                    <a
                      key={si}
                      href={s.url}
                      onClick={(e) => { e.preventDefault(); void (window as unknown as { api?: { openUrl?: (u: string) => Promise<{ ok: boolean }> } }).api?.openUrl?.(s.url) }}
                      style={{ display: 'block', fontSize: 11.5, color: isLight ? '#3b6ef5' : '#7aa2ff', textDecoration: 'none', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', marginBottom: 2 }}
                      title={s.title}
                    >[{si + 1}] {s.title}</a>
                  ))}
                </div>
              ) : null}
              {/* 意图卡片：一键创建工作流 */}
              {/* 文档选择菜单（多文档节点分析）：点击分析对应文档 */}
              {m.docChoices && (
                <div style={{ marginTop: 8, display: 'flex', flexDirection: 'column', gap: 6 }}>
                  {m.docChoices.map((c) => (
                    <button
                      key={c.nodeId}
                      onClick={(e) => { e.stopPropagation(); void runDocChoice(c.nodeId, c.mode) }}
                      style={{ textAlign: 'left', padding: '7px 10px', borderRadius: 8, border: `1px solid ${isLight ? '#dfe5ec' : 'rgba(120,140,180,0.3)'}`, background: isLight ? '#f5f7fa' : 'rgba(60,64,74,0.5)', color: isLight ? '#3a4149' : '#dfe7f3', cursor: 'pointer', fontSize: 12 }}
                    >{c.label}</button>
                  ))}
                </div>
              )}
              {m.intent && (
                <div style={{ marginTop: 8, padding: '8px 10px', borderRadius: 10, background: isLight ? '#f5f7fa' : 'rgba(28,30,36,0.6)', border: `1px solid ${isLight ? '#e3e8ee' : 'rgba(120,140,180,0.25)'}` }}>
                  <div style={{ fontSize: 12, fontWeight: 600, color: isLight ? '#3a4149' : '#e6ecf5' }}>🎯 {m.intent.title}</div>
                  <div style={{ fontSize: 11, color: isLight ? '#5a6572' : '#9fb4d8', margin: '3px 0 7px' }}>节点计划：{m.intent.plan}</div>
                  <button
                    onClick={(e) => { e.stopPropagation(); acceptIntent(m.intent!.id) }}
                    disabled={m.intent.accepted}
                    style={{ width: '100%', padding: '6px 0', borderRadius: 8, border: 'none', cursor: m.intent.accepted ? 'default' : 'pointer', fontSize: 12, fontWeight: 600, color: '#fff', background: m.intent.accepted ? 'rgba(80,180,120,0.8)' : 'linear-gradient(135deg,#4d7dfd,#7a5cff)' }}
                  >{m.intent.accepted ? '✓ 已在画布创建' : '一键创建工作流'}</button>
                </div>
              )}
            </div>
          ))}
        </div>
      )}

      {/* 输入栏（📎 左 | 输入中 | ▲ 发送右） */}
      <div style={{ display: 'flex', alignItems: 'flex-end', gap: 7, padding: '9px 10px', flexShrink: 0, marginTop: 'auto' }}>
        {/* ★ 模型选择（2026-11-07 用户「恢复」）：点击 🧠 向上弹菜单，聪明/轻量二选一，localStorage 持久化 */}
        <div style={{ position: 'relative', flexShrink: 0 }}>
          {modelMenu && (
            <div
              data-model-menu="1"
              style={{ position: 'absolute', left: 0, bottom: 42, width: 276, background: T.panelBg, border: `1px solid ${T.panelBorder}`, borderRadius: 12, boxShadow: T.shadow, padding: 5, zIndex: 300 }}
            >
              <div style={{ fontSize: 11, color: T.subText, padding: '4px 10px 6px' }}>选择本地模型</div>
              {LLM_MODEL_OPTIONS.map((m) => (
                <div
                  key={m.file}
                  onClick={() => { setLlmModel(m.file); setModelMenu(false) }}
                  style={{ padding: '7px 10px', borderRadius: 8, cursor: 'pointer', fontSize: 13, background: llmModel === m.file ? 'rgba(122,92,255,0.14)' : 'transparent', fontWeight: llmModel === m.file ? 600 : 400, color: T.bodyText }}
                >
                  <div>{m.icon} {m.label}{llmModel === m.file ? ' ✓' : ''}</div>
                  <div style={{ fontSize: 10.5, opacity: 0.6, marginTop: 1 }}>{m.desc}</div>
                </div>
              ))}
              {/* ★ 思考强度（2026-11-07 用户需求）：off 秒答 / normal 常规 / deep 深度；选完不关菜单方便连调 */}
              <div style={{ fontSize: 11, color: T.subText, padding: '8px 10px 6px', borderTop: `1px solid ${T.panelBorder}`, marginTop: 4 }}>思考强度（影响回答速度与深度）</div>
              {THINKING_LEVELS.map((lv) => (
                <div
                  key={lv.id}
                  onClick={() => setThinkingLv(lv.id)}
                  style={{ padding: '7px 10px', borderRadius: 8, cursor: 'pointer', fontSize: 13, background: thinkingLv === lv.id ? 'rgba(77,125,253,0.12)' : 'transparent', fontWeight: thinkingLv === lv.id ? 600 : 400, color: T.bodyText }}
                >
                  <div>{lv.label}{thinkingLv === lv.id ? ' ✓' : ''}</div>
                  <div style={{ fontSize: 10.5, opacity: 0.6, marginTop: 1 }}>{lv.desc}</div>
                </div>
              ))}
            </div>
          )}
          <button
            title={`本地模型：${llmModelMeta(llmModel).label}（点击切换）`}
            onClick={() => setModelMenu((o) => !o)}
            style={{ width: 34, height: 34, borderRadius: '50%', border: `1px solid ${llmModel === LLM_MODEL_OPTIONS[0].file ? 'rgba(122,92,255,0.6)' : 'rgba(244,166,73,0.65)'}`, background: llmModel === LLM_MODEL_OPTIONS[0].file ? 'rgba(122,92,255,0.18)' : 'rgba(244,166,73,0.16)', color: llmModel === LLM_MODEL_OPTIONS[0].file ? '#7a5cff' : '#e8963a', cursor: 'pointer', fontSize: 14, flexShrink: 0, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
          >{llmModelMeta(llmModel).icon}</button>
        </div>
        <button
          title="添加文件"
          onClick={() => fileRef.current?.click()}
          style={{ width: 34, height: 34, borderRadius: '50%', border: `1px solid ${T.btnBorder}`, background: T.btnBg, color: T.btnColor, cursor: 'pointer', flexShrink: 0, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', opacity: contentOpacity, transition: 'opacity 0.18s ease' }}
        ><PaperclipIcon /></button>
        <input ref={fileRef} type="file" multiple style={{ display: 'none' }} onChange={(e) => { ingestFiles(e.target.files); e.target.value = '' }} />
        {files.map((f, i) => (
          <span key={i} style={{ display: 'inline-flex', alignItems: 'center', gap: 5, padding: '3px 9px', borderRadius: 12, background: T.chipBg, border: `1px solid ${T.chipBorder}`, color: T.chipText, fontSize: 11, whiteSpace: 'nowrap', flexShrink: 0, maxWidth: 110, overflow: 'hidden', textOverflow: 'ellipsis' }}>
            {f}
            <span style={{ cursor: 'pointer', opacity: 0.7 }} onClick={() => { setFiles((fs) => fs.filter((_, j) => j !== i)); setFileObjs((os) => os.filter((_, j) => j !== i)) }}>×</span>
          </span>
        ))}
        <textarea
          ref={taRef}
          value={input}
          rows={1}
          onChange={(e) => { setInput(e.target.value); taGrow(e.target as HTMLTextAreaElement) }}
          onKeyDown={(e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send() } }}
          placeholder="描述工作流，或拖文件到这里…"
          style={{ flex: '1 1 0%', width: '100%', minWidth: 60, maxWidth: '100%', height: 36, background: T.inputBg, border: `1px solid ${T.inputBorder}`, borderRadius: 10, outline: 'none', color: T.inputText, fontSize: 13, padding: '8px 12px', resize: 'none', fontFamily: 'inherit', lineHeight: '20px', overflowY: 'hidden', opacity: contentOpacity, transition: 'opacity 0.18s ease' }}
          onInput={(e) => taGrow(e.target as HTMLTextAreaElement)}
        />
        <button
          title="发送"
          onClick={send}
          style={{ width: 34, height: 34, borderRadius: '50%', border: 'none', cursor: 'pointer', flexShrink: 0, background: dragOver ? 'rgba(110,170,255,0.85)' : 'linear-gradient(135deg,#4d7dfd,#7a5cff)', color: '#fff', fontSize: 13, boxShadow: '0 4px 14px rgba(90,120,255,0.35)', opacity: contentOpacity, transition: 'opacity 0.18s ease' }}
        >▲</button>
      </div>

      {/* 右键菜单：复制选中文字 / 引用到输入框 */}
      {ctxMenu && (
        <>
          <div
            data-dock-menu-mask="1"
            style={{ position: 'fixed', inset: 0, zIndex: 400 }}
            /* ★ 只认左键按下：右键释放后浏览器还会补发 click/auxclick，
               若用 onClick 关闭，菜单刚弹就被自己关掉（实测时有时无的根因）。 */
            onMouseDown={(e) => { if (e.button === 0) setCtxMenu(null) }}
            onContextMenu={(e) => { e.preventDefault(); setCtxMenu(null) }}
          />
          <div
            ref={menuRef}
            data-dock-quote-menu="1"
            onContextMenu={(e) => e.preventDefault()}
            style={{
              position: 'fixed',
              left: menuPos?.left ?? 0,
              top: menuPos?.top ?? 0,
              /* 首帧先隐藏，测量完落位再显示（避免从左上角闪一下） */
              visibility: menuPos ? 'visible' : 'hidden',
              zIndex: 401,
              background: isLight ? '#ffffff' : 'rgba(40,44,52,0.98)',
              border: `1px solid ${isLight ? '#d5dbe3' : 'rgba(120,140,180,0.35)'}`,
              borderRadius: 10, boxShadow: '0 8px 28px rgba(0,0,0,0.22)',
              padding: 4, minWidth: 120,
            }}
          >
            <div
              className="dock-menu-item"
              style={{ padding: '7px 12px', borderRadius: 7, fontSize: 13, color: isLight ? '#333a42' : '#dfe7f3', cursor: 'pointer' }}
              onClick={() => {
                copyPlainText(ctxMenu.text)
                setCtxMenu(null)
              }}
            >{ctxMenu.fromSel ? '复制选中文字' : '复制整条消息'}</div>
            <div
              className="dock-menu-item"
              style={{ padding: '7px 12px', borderRadius: 7, fontSize: 13, color: isLight ? '#333a42' : '#dfe7f3', cursor: 'pointer' }}
              onClick={() => {
                insertQuote(ctxMenu.text)
                setCtxMenu(null)
              }}
            >引用到输入框（作为提示词）</div>
          </div>
        </>
      )}

      {/* ★ 2026-11-07 宽度拖拽手柄（贴左时在右缘 / 贴右时在左缘）：悬停变宽，松手记忆 */}
      {showBody && (
        <div
          onPointerDown={startResize}
          title="拖动调整对话框宽度"
          style={{
            position: 'absolute', top: 12, bottom: 12, width: 7,
            ...(dock === 'l' ? { right: -4 } : { left: -4 }),
            cursor: 'ew-resize', zIndex: 2, borderRadius: 4,
          }}
        />
      )}
    </div>

    {/* ★ 2026-11-07 落点幽灵框（第三次改版，解决「提示没和实际框体对齐」）：
        位置/尺寸由 startDrag 直接从面板 getBoundingClientRect 复制，所以虚线框 = 松手后
        面板会占的那块矩形，1px 都不差（以前是另算 safeTop/offsetHeight，必然错位）。
        常驻挂载、display 由 JS 开关 —— 拖动期间不触发 React 渲染（这是「卡住」的根因）。
        zIndex 260：压在面板之上，但只有 1.5px 虚线边 + 4% 填充，不挡内容。 */}
    <div
      ref={ghostRef}
      data-dock-ghost="1"
      style={{
        position: 'absolute', display: 'none', zIndex: 260, pointerEvents: 'none',
        boxSizing: 'border-box',
        border: '1.5px dashed rgba(77,125,253,0.85)',
        background: 'rgba(77,125,253,0.045)',
        borderRadius: 14,
        boxShadow: '0 0 0 3px rgba(77,125,253,0.10)',
        transition: 'left 0.16s cubic-bezier(0.4,0,0.2,1), top 0.16s cubic-bezier(0.4,0,0.2,1), width 0.16s, height 0.16s',
      }}
    >
      <div
        ref={ghostLabelRef}
        data-dock-ghost-label="1"
        style={{
          position: 'absolute', top: -2, left: '50%', transform: 'translate(-50%,-100%)',
          padding: '5px 14px', borderRadius: 99, whiteSpace: 'nowrap',
          background: 'rgba(77,125,253,0.94)', color: '#fff',
          fontSize: 12, letterSpacing: 0.5, fontWeight: 600,
          boxShadow: '0 6px 18px rgba(77,125,253,0.35)',
        }}
      />
    </div>
    </>
  )
}
