/* ★ 对话上下文管理（2026-11-07 用户「为啥不能联系上下文」+ 四条要求）：
 *  1. 是否符合本软件的主题 → isAppTopic（决定提示语是否按「工作流助手」口径回答）
 *  2. 聊天是否更换话题       → isTopicSwitch（决定是否沿用上文；换话题就丢历史）
 *  3. 当前对话信息是否足够   → needsMoreInfo（不足时让模型先反问，而不是硬编）
 *  4. 好的压缩方式           → compressTurns（近期原文 + 更早截断摘要，按字符预算压缩）
 *
 * 为什么放在前端而不是靠 llama 会话的隐式历史：
 *  · 天气这类答案是**本地工具**生成的，从未进过模型会话，模型当然接不上「穿多少衣服」；
 *  · 隐式历史无上限，聊几轮就把 8192 上下文撑爆，且无法按话题裁剪。
 *  所以每次请求都传一份「压缩后的显式上下文块」，主进程侧重置会话历史（单轮干净）。
 */

export type CtxTurn = { role: 'user' | 'assistant'; text: string }

/** 软件主题关键词：命中即认为与「工作流画布」相关 */
const APP_TOPIC_RE =
  /(节点|画布|工作流|连线|连线|套管|燃气|管道|阀门|PDF|pdf|Word|word|Excel|excel|表格|文档|合并|拆分|瘦身|转换|导出|导入|批量|命名|编号|N\d{2,}|属性|运行|输出|模板|字段|记录|报表)/

/** 追问/指代标记：出现时视为「延续上一轮话题」 */
const FOLLOW_RE =
  /(^|[，。！？、\s])(那|这|这个|那个|它|他们|其|该|上面|刚才|刚刚|还有|然后|接着|继续|再|另外)|为什么|怎么|多少|多久|哪个|哪家|呢[?？]?$|吗[?？]?$|呢$/

/** 过滤占位/系统提示类消息，避免污染上下文 */
const PLACEHOLDER_RE =
  /^(（?(正在|已|搜到|画布上|已调起)[^）]*）?|（思考中…）|（模型|（思考)|^\s*$/

export function isAppTopic(text: string): boolean {
  return APP_TOPIC_RE.test(text)
}

/* ★ 当前时间行（2026-11-08 用户「天气、日期这类都不联网的么」）：
 *  天气走联网 Open-Meteo；日期时间**不需要联网也不需要模型**——本地模型训练截止后根本不知道今天几号，
 *  问它会瞎编。所以把系统时间注入提示词（供模型答相对时间），纯日期问题直接本地直答（见 CanvasChatDock）。 */
export function nowLine(d = new Date()): string {
  const week = '日一二三四五六'[d.getDay()]
  const hh = String(d.getHours()).padStart(2, '0')
  const mm = String(d.getMinutes()).padStart(2, '0')
  return `今天是${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日（星期${week}）${hh}:${mm}`
}

/* 纯日期/时间问句（本地直答，不联网不耗模型）。
 * 只收「今天几号」「星期几」「几点了」这类**不带推算**的问法；
 * 「明天几号」「下周一星期几」「10月1日星期几」留给注入了当前时间的模型算，别在这里拦。
 * 负例（必须放过）：「今天几号发货」「星期三开会」「几点了还不睡」。 */
const DT_QUERY_RE =
  /^(?:今天|今日|现在|当前)?(?:是|的)?(?:几号|多少号|几月几号|星期几|礼拜几|周几|几点了?|几点钟|什么时间|日期是多少|日期)$/

export function isDateTimeQuestion(text: string): boolean {
  return DT_QUERY_RE.test(text.trim().replace(/[\s?？!！。，,]+$/g, ''))
}

/** 是否是对上文的追问（指代词 / 疑问词 / 极短句） */
export function isFollowUp(text: string): boolean {
  const t = text.trim()
  if (!t) return false
  if (FOLLOW_RE.test(t)) return true
  return t.length <= 8 /* 「穿多少衣服」这类极短句多半接着上文说 */
}

/** 中文友好相似度：字符 2-gram Jaccard（去标点与常见虚词） */
export function similarity(a: string, b: string): number {
  const norm = (s: string) => s.replace(/[\s，。！？、,.!?;；:：""''（）()【】\[\]]/g, '').replace(/[的了呢吗吧啊呀哦和与及或是]/g, '')
  const grams = (s: string): Set<string> => {
    const t = norm(s)
    const out = new Set<string>()
    if (t.length === 1) out.add(t)
    for (let i = 0; i + 1 < t.length; i++) out.add(t.slice(i, i + 2))
    return out
  }
  const A = grams(a)
  const B = grams(b)
  if (!A.size || !B.size) return 0
  let hit = 0
  for (const g of A) if (B.has(g)) hit++
  return hit / (A.size + B.size - hit)
}

/** 是否换了话题：无历史 → 不算换；有追问标记 → 没换；与近两轮相似度极低 → 换了 */
export function isTopicSwitch(turns: CtxTurn[], text: string): boolean {
  const users = turns.filter((t) => t.role === 'user' && !PLACEHOLDER_RE.test(t.text)).slice(-2)
  if (!users.length) return false
  if (isFollowUp(text)) return false
  const best = Math.max(...users.map((u) => similarity(u.text, text)))
  return best < 0.08
}

/** 单条截断（保留句读，不硬切半个词） */
function clip(s: string, max: number): string {
  const t = s.replace(/\s+/g, ' ').trim()
  if (t.length <= max) return t
  const cut = t.slice(0, max)
  const p = Math.max(cut.lastIndexOf('。'), cut.lastIndexOf('，'), cut.lastIndexOf('；'))
  return (p > max * 0.5 ? cut.slice(0, p + 1) : cut) + '…'
}

/**
 * 压缩对话：
 *  · 最近 keepRecent 条保留较完整（每条 ≤ recentChars）
 *  · 更早的每条压成一句摘要（≤ earlyChars）
 *  · 总量超预算时从最早开始丢
 */
export function compressTurns(
  turns: CtxTurn[],
  opts: { budgetChars?: number; keepRecent?: number; recentChars?: number; earlyChars?: number } = {},
): { lines: string[]; compressed: boolean; dropped: number } {
  const budget = opts.budgetChars ?? 900
  const keepRecent = opts.keepRecent ?? 6
  const recentChars = opts.recentChars ?? 220
  const earlyChars = opts.earlyChars ?? 60
  const clean = turns.filter((t) => t.text && !PLACEHOLDER_RE.test(t.text))
  if (!clean.length) return { lines: [], compressed: false, dropped: 0 }

  const recent = clean.slice(-keepRecent)
  const early = clean.slice(0, Math.max(0, clean.length - keepRecent))

  const lines: string[] = []
  /* 早期先压成单行摘要（越早越短），从新到旧塞，超预算直接丢更早的 */
  const earlyLines = early.map((t) => `${t.role === 'user' ? '用户' : '助手'}：${clip(t.text, earlyChars)}`)
  const recentLines = recent.map((t) => `${t.role === 'user' ? '用户' : '助手'}：${clip(t.text, recentChars)}`)

  /* 预先占位：最近段落优先保证 */
  const recentTotal = recentLines.reduce((n, l) => n + l.length + 1, 0)
  let remain = Math.max(0, budget - recentTotal)
  const keptEarly: string[] = []
  for (let i = earlyLines.length - 1; i >= 0; i--) {
    const l = earlyLines[i]
    if (l.length + 1 > remain) break
    keptEarly.unshift(l)
    remain -= l.length + 1
  }
  if (keptEarly.length < earlyLines.length) lines.push(`（更早的 ${earlyLines.length - keptEarly.length} 条对话已压缩省略）`)
  lines.push(...keptEarly)
  lines.push(...recentLines)
  return { lines, compressed: keptEarly.length < earlyLines.length, dropped: earlyLines.length - keptEarly.length }
}

/**
 * 生成插进 prompt 的上下文块。
 * switched=true 时只给一句"上一轮话题"提示（避免旧话题污染新问题）。
 */
export function buildContextBlock(
  turns: CtxTurn[],
  text: string,
  opts?: { budgetChars?: number; keepRecent?: number },
): { block: string; switched: boolean; lines: number } {
  const clean = turns.filter((t) => t.text && !PLACEHOLDER_RE.test(t.text))
  if (!clean.length) return { block: '', switched: false, lines: 0 }

  const switched = isTopicSwitch(clean, text)
  if (switched) {
    const lastUser = [...clean].reverse().find((t) => t.role === 'user')
    const hint = lastUser ? `（上一轮聊的是「${clip(lastUser.text, 30)}」，与当前问题无关，可忽略）` : ''
    return { block: hint ? `【对话背景】${hint}\n` : '', switched: true, lines: 0 }
  }

  const { lines } = compressTurns(clean, { budgetChars: opts?.budgetChars ?? 900, keepRecent: opts?.keepRecent ?? 6 })
  if (!lines.length) return { block: '', switched, lines: 0 }
  return {
    block: `【最近对话】（同一话题，必须结合它回答：不要重复索要这里已有的信息；缺少必要信息时先用一句话说明缺什么再给建议）\n${lines.join('\n')}\n`,
    switched: false,
    lines: lines.length,
  }
}

/** 信息是否明显不足（历史为空 + 极短 + 无主题词 → 让模型先反问） */
export function needsMoreInfo(turns: CtxTurn[], text: string): boolean {
  const clean = turns.filter((t) => t.text && !PLACEHOLDER_RE.test(t.text))
  const t = text.trim()
  if (!t) return true
  if (clean.length > 0) return false
  return t.length <= 6 && !isAppTopic(t) && isFollowUp(t)
}
