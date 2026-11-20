import { capsuleKindClass } from './capsule'

/**
 * 行内胶囊 · 混排解析（批次 12）
 * ────────────────────────────────────────────────────────────────
 * 目标：让「一段文字里夹着几个胶囊」成为可存储、可渲染的一等公民。
 * 参照物：WorkBuddy 的引用胶囊（拖参数进正文，胶囊落在文字中间）。
 *
 * 关键设计：胶囊是**文本内的标记**，不改存储结构。
 *   · 纯文字：        "总长 1200"
 *   · 文字+胶囊：     "总长 ⟨ref:nodeId|字段名⟩ 米，其中 ⟨ref:nodeId|字段名⟩ 处需复核"
 * 选这个方案而不是「拆成 token 数组存 store」的原因：
 *   1. 存储零改动 —— 仍是一个字符串，落库/导出/复制粘贴/公式求值全不受影响；
 *   2. 旧数据天然兼容 —— 没有标记的纯文本解析结果就是一个 text 段，不会有行为突变；
 *   3. 用户可手改 —— 在prompt 里敲标记就能插胶囊，不必造UI。
 *
 * 标记语法（选窄括号是为���避开文档正文里常见的尖括号用法）：
 *   ⟨ref:节点ID|字段名⟩        —— 引用某节点的某个参数（字段名为空 = 整节点）
 *   ⟨ref:节点ID|字段名|kind⟩   —— 带节点类型，胶囊按类型分色（省略时回退默认色）
 *   ⟨const:字面量⟩             —— 常量胶囊（不指向节点，值固定）
 *
 * 纯计算层：零 IO 依赖，浏览器可直接跑。
 */

/** 解析出的行片段 */
export type InlineToken =
  | { type: 'text'; value: string }
  | {
      type: 'capsule'
      /** 引用目标：节点 id（node://省略时即nodeId）/ 常量标记 */
      target: string
      /** 参数名；空 = 整节点引用 */
      field: string
      /** 节点类型，决定胶囊色系 */
      kind?: string
      /** 字面量（仅 const胶囊有） */
      literal?: string
      /** 完整 URN（供 title 提示与后续解析用） */
      urn?: string
    }

/** 胶囊左段显示名：节点 id 太长，优先用字段名 */
const REF_RE = /⟨ref:([^|⟩]+)\|([^|⟩]*)(?:\|([^⟩]+))?⟩/g
const CONST_RE = /⟨const:([^⟩]*)⟩/g

/**
 * 把带标记的文本解析成行片段序列。
 * 纯函数，无副作用；未匹配标记时返回单个 text 段。
 */
export function parseInline(raw: unknown): InlineToken[] {
  const s = String(raw ?? '')
  if (!s) return []
  if (!s.includes('⟨')) return [{ type: 'text', value: s }]

  const tokens: InlineToken[] = []
  // 两种标记共用一次扫描：用带捕获组的 alternation 保证原文顺序不乱
  const re = new RegExp(`${REF_RE.source}|${CONST_RE.source}`, 'g')
  let last = 0
  let m: RegExpExecArray | null
  while ((m = re.exec(s)) !== null) {
    if (m.index > last) tokens.push({ type: 'text', value: s.slice(last, m.index) })
    // ① ⟨ref:nodeId|field[|kind]⟩ —— 命中时 m[1] 有值
    if (m[1] !== undefined) {
      const nodeId = m[1]
      const field = m[2] ?? ''
      const kind = m[3] || undefined
      tokens.push({
        type: 'capsule',
        target: nodeId,
        field,
        kind,
        urn: field ? `node://${nodeId}#${field}` : `node://${nodeId}`,
      })
    } else if (m[4] !== undefined) {
      // ② ⟨const:字面量⟩
      tokens.push({ type: 'capsule', target: 'const', field: '', literal: m[4] })
    }
    last = re.lastIndex
  }
  if (last < s.length) tokens.push({ type: 'text', value: s.slice(last) })
  return tokens
}

/** 胶囊 CSS 类（含类型色系与失效态） */
export function inlineCapsuleClass(kind?: string, dead?: boolean): string {
  return ['ref-capsule', 'inline-cap', capsuleKindClass(kind), dead ? 'dead' : '']
    .filter(Boolean)
    .join(' ')
}

/** 胶囊左段文本：常量取字面量；引用优先字段名，空则显示「整体」 */
export function capsuleLabelOf(t: Extract<InlineToken, { type: 'capsule' }>): string {
  if (t.type !== 'capsule') return ''
  if (t.literal !== undefined) return t.literal
  return t.field || '整体'
}

/** 胶囊右段文本：参数名（与左段互为「节点名 / 参数名」两段式） */
export function capsuleFieldOf(t: Extract<InlineToken, { type: 'capsule' }>): string {
  if (t.type !== 'capsule') return ''
  if (t.literal !== undefined) return '常量'
  return t.field || '整节点'
}

/** 胶囊悬停提示：把 URN 原样暴露，便于排查引用问题 */
export function capsuleTitleOf(t: Extract<InlineToken, { type: 'capsule' }>, nodeName?: string): string {
  if (t.type !== 'capsule') return ''
  if (t.literal !== undefined) return `常量：${t.literal}`
  const left = nodeName || t.target
  return t.field ? `${left}.${t.field}\n${t.urn}` : `${left}（整节点）\n${t.urn}`
}

/**
 * 插入点换算：纯文本插入位 → 带标记原文下标。
 *
 * 为什么需要：渲染层与用户心智都是「纯文本」（胶囊按 1 个字符宽排），
 * 但落库是带标记的原文。插入胶囊时拿到纯文本插入位，得映回原文位置。
 *
 * 口径（唯一且自洽的一条）：**`plainOffset` = 已消费的纯文本字符数**，胶囊算 1 个字符。
 * 映射表（"ab⟨ref:n1|f⟩cd"，原文 14，纯文本 "ab"+胶囊+"cd" 共 5 格）：
 *   off=0 → 0     a 之前
 *   off=1 → 1     b 之前
 *   off=2 → 12    b 之后（即胶囊之后；纯文本里第 2 格就是胶囊）
 *   off=3 → 12    c 之前（同属"胶囊之后"，因为胶囊只占 1 格，off=3 已越过它）
 *   off=4 → 13    d 之前
 *   off=5 → 14    末尾
 * 越界（off>5）一律夹到末尾。
 */
export function offsetAfterInsert(src: string, plainOffset: number): number {
  const re = new RegExp(`${REF_RE.source}|${CONST_RE.source}`, 'g')
  let m: RegExpExecArray | null
  let seen = 0 // 已消费的纯文本字符数
  let srcCursor = 0 // 对应的原文下标
  while ((m = re.exec(src)) !== null) {
    const segLen = m.index - srcCursor // 本段纯文本长度
    // ① 插入位严格落在本段纯文本之内
    if (plainOffset < seen + segLen) return srcCursor + (plainOffset - seen)
    // ② 已消费完本段且还没越过胶囊 → 落在胶囊那一格，映到标记之后
    if (plainOffset <= seen + segLen + 1) return m.index + m[0].length
    // ③ 越过胶囊继续往后
    seen += segLen + 1
    srcCursor = m.index + m[0].length
  }
  return Math.min(srcCursor + (plainOffset - seen), src.length)
}

/**
 * 纯文本视图长度：胶囊按 1 个字符计。
 * 与 offsetAfterInsert 的口径配套 —— 拿它当下标传给 offsetAfterInsert 才不错位。
 */
export function plainLength(raw: unknown): number {
  const s = String(raw ?? '')
  if (!s.includes('⟨')) return s.length
  const re = new RegExp(`${REF_RE.source}|${CONST_RE.source}`, 'g')
  // 胶囊各扣掉 (标记长 - 1)：占位是1 格
  let n = s.length
  let m: RegExpExecArray | null
  while ((m = re.exec(s)) !== null) n -= m[0].length - 1
  return Math.max(0, n)
}

/** 生成胶囊标记文本（插入用） */
export function buildRefMark(nodeId: string, field: string, kind?: string): string {
  return `⟨ref:${nodeId}|${field}${kind ? `|${kind}` : ''}⟩`
}

/**
 * 统计一段文本里的胶囊数量（画布节点角标用）。
 * 只做计数，不建对象，开销最小。
 */
export function countCapsules(raw: unknown): number {
  const s = String(raw ?? '')
  if (!s.includes('⟨')) return 0
  const re = new RegExp(`${REF_RE.source}|${CONST_RE.source}`, 'g')
  let n = 0
  while (re.exec(s) !== null) n++
  return n
}