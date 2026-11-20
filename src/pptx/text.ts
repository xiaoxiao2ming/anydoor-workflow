/**
 * pptx 段落级占位符替换 —— 整个 PPT 绑定里最容易翻车的一块。
 *
 * ## 为什么需要它
 *
 * PowerPoint 保存文件时，会按自己的规则把一个 `a:p` 里的文字切成多个 `a:r`
 * （拼写检查边界、语言切换、格式变化、甚至毫无理由）。所以用户明明写的是
 * `{{field:项目名}}`，存进 XML 后可能变成：
 *
 *   <a:r><a:t>{{field:</a:t></a:r><a:r><a:t>项目名}}</a:t></a:r>
 *
 * 直接对单个 `a:t` 做正则匹配，一个都匹配不到 —— 而且不报错、静默失效，
 * 属于最难查的那类 bug。
 *
 * ## 做法
 *
 * 在**段落内**把所有 `a:t` 的原始内容拼成一段 plain text，在拼接结果上找占位符，
 * 再把命中的区间映射回各个 run 做替换。docx 侧踩过的同一个坑，这里一并解决。
 *
 * ## 两个刻意的设计取舍
 *
 * 1. **匹配在「未解码实体」的原始内容上做**。`&amp;` 这类实体会让 plain 的
 *    下标和 XML 原文错位；占位符 `{{...}}` 里不可能出现实体，所以拼原文是安全的，
 *    换来下标一一对应、不用维护映射表。
 * 2. **跨 run 时，值整体放进第一个 run**，其余 run 的对应片段删掉。
 *    格式取第一个 run 的 —— 可预测，且符合直觉（占位符从哪开始就跟哪的格式）。
 */

/** 单个 `a:t` 内容在段落字符串里的区间 [cs, ce) */
interface RunSpan {
  cs: number
  ce: number
}

const A_T_RE = /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g

/** 占位符：`{{field:项目名}}`，kind 决定走哪条绑定路径 */
const PH_RE = /\{\{\s*(field|rich|each|if|image)\s*:\s*([^{}]+?)\s*\}\}/g

export type PlaceholderKind = 'field' | 'rich' | 'each' | 'if' | 'image'

export interface PlaceholderHit {
  kind: PlaceholderKind
  path: string
  /** 原文里的完整占位符，用于定位与计数 */
  raw: string
}

export function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

/**
 * 找出一个 `a:p` 内部的全部占位符（只看文字，不做替换）。
 * 供扫描阶段统计绑定清单用。
 */
export function scanParagraph(block: string): PlaceholderHit[] {
  const plain = plainOf(block)
  const hits: PlaceholderHit[] = []
  for (const m of plain.matchAll(PH_RE)) {
    hits.push({ kind: m[1] as PlaceholderKind, path: m[2], raw: m[0] })
  }
  return hits
}

/**
 * 替换一个 `a:p` 内的占位符。
 *
 * @param resolve 给定占位符文本，返回替换值；返回 null 表示「这个不归我管」
 *                （比如 each 行由表格层处理），原样保留。
 * @returns 新段落字符串 + 实际替换数量
 */
export function replaceInParagraph(
  block: string,
  resolve: (tag: string, hit: PlaceholderHit) => string | null,
): { block: string; count: number } {
  return replaceByPattern(block, PH_RE, (m) =>
    resolve(m[0], { kind: m[1] as PlaceholderKind, path: m[2], raw: m[0] }),
  )
}

/**
 * 跨 run 替换的公用实现：类型化占位符和短占位符走的是同一套逻辑，
 * 只有正则不同 —— 抽出来避免两边各写一份、然后某一侧漏修 bug。
 *
 * resolve 直接拿 match 数组，由调用方按自己正则的分组语义解构；
 * 不在这里猜 m[1] 是名字还是类型。
 */
function replaceByPattern(
  block: string,
  re: RegExp,
  resolve: (m: RegExpMatchArray) => string | null,
): { block: string; count: number } {
  const runs = spansOf(block)
  if (!runs.length) return { block, count: 0 }

  const plain = runs.map((r) => block.slice(r.cs, r.ce)).join('')

  // 收集所有替换动作，稍后从后往前应用（先改尾部，前面的下标才不会失效）
  const edits: { start: number; end: number; text: string }[] = []
  let count = 0

  for (const m of plain.matchAll(re)) {
    const value = resolve(m)
    if (value == null) continue

    const s = m.index!
    const e = s + m[0].length
    const rep = escapeXml(value)
    let first = true

    // 占位符可能横跨多个 run：逐个 run 计算与 [s, e) 的重叠部分
    let cursor = 0
    for (let i = 0; i < runs.length; i++) {
      const rs = cursor
      const rn = cursor + (runs[i].ce - runs[i].cs)
      cursor = rn
      if (rn <= s || rs >= e) continue

      const from = Math.max(s, rs)
      const to = Math.min(e, rn)
      // plain 坐标 → 该 run 内容内的坐标
      const localStart = runs[i].cs + (from - rs)
      const localEnd = runs[i].cs + (to - rs)
      edits.push({ start: localStart, end: localEnd, text: first ? rep : '' })
      first = false
    }
    count++
  }

  if (!edits.length) return { block, count: 0 }

  edits.sort((a, b) => b.start - a.start)
  let out = block
  for (const ed of edits) out = out.slice(0, ed.start) + ed.text + out.slice(ed.end)
  return { block: out, count }
}

/**
 * 短占位符 `{{字段名}}` —— 只在 each 循环行内使用。
 *
 * 负向前瞻排除 `{{field:xxx}}` 这类带类型的写法，否则两者会互相误伤：
 * 类型化的归绑定层，短的归循环层。
 */
const SHORT_RE = /\{\{(?!(?:field|rich|each|if|image)\s*:)\s*([^{}]+?)\s*\}\}/g

/**
 * 替换循环行内的短占位符。与 replaceInParagraph 同一套跨 run 逻辑，
 * 只是匹配的正则不同。
 */
export function replaceShortInParagraph(
  block: string,
  resolve: (name: string) => string | null,
): { block: string; count: number } {
  return replaceByPattern(block, SHORT_RE, (m) => resolve(m[1]))
}

/**
 * 删掉段落里的 each 标记 `{{each:xxx}}`。
 * 标记只用于识别"哪一行是模板行"，本身不该出现在产出里。
 */
export function stripEachMarker(block: string): string {
  return replaceByPattern(block, /\{\{\s*each\s*:\s*[^{}]+?\s*\}\}/g, () => '').block
}

/**
 * `if` 语义：整段删除。调用方在段落外层判断，这里只负责从 slide XML 里摘掉。
 */
export function dropParagraph(xml: string, block: string, at: number): string {
  return xml.slice(0, at) + xml.slice(at + block.length)
}

/* ---------- 内部 ---------- */

function spansOf(block: string): RunSpan[] {
  const spans: RunSpan[] = []
  for (const m of block.matchAll(A_T_RE)) {
    const open = m.index! + m[0].indexOf('>') + 1
    spans.push({ cs: open, ce: open + m[1].length })
  }
  return spans
}

function plainOf(block: string): string {
  return spansOf(block)
    .map((r) => block.slice(r.cs, r.ce))
    .join('')
}

/**
 * 找到一个 slide XML 里的所有 `a:p` 区间。
 *
 * 注意：`a:p` 可以嵌套吗？在 `p:txBody` 里不会（富文本是扁平的 a:p 序列），
 * 但 `a:p` 里可能含 `a:br`/`a:fld`，没有嵌套 a:p，所以非贪婪匹配安全。
 * docx 侧用的是深度扫描配对，那边因为 w:sdt 会嵌套，这边不需要。
 */
export function paragraphBlocks(xml: string): { block: string; at: number }[] {
  const out: { block: string; at: number }[] = []
  const re = /<a:p>([\s\S]*?)<\/a:p>/g
  for (const m of xml.matchAll(re)) out.push({ block: m[0], at: m.index! })
  return out
}
