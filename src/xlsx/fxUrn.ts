/**
 * fx / 格内编辑器的「可读串 ⇄ URN」互转层（批次 17 抽出的纯计算模块）。
 *
 * 为什么单独成文件：这三组函数是**存储层与编辑层之间唯一的翻译层**，口径错一次就是
 * 「点参数看着插进去了、但存的不是引用、联动永远不命中」这种静默故障（批次 17 实测踩过）。
 * 抽出来只为能被 vitest 直接覆盖 —— 它们不碰 DOM / electron / workbook，纯粹字符串→字符串。
 *
 * 语义（用户 2026-09-27 拍板）：
 *   编辑态全程「可读串」（表名·参数名），URN 只活在存储层与提交边界。
 *     选中格 / 插入 / 换参 → urnToReadable；提交 → readableToUrn。
 *     手打 / 粘贴 URN 向下兼容（渲染胶囊 + 提交原样通过 isPureUrn / resolveUrn）。
 */
import { parseUrn } from '../ref/urn'
import { computeNode } from '../core/engine'
import type { TableSchema } from '../core/schema'
import type { Row } from '../core/store'

/** fx 草稿里的 URN token 边界：终止于空白/公式分隔符/运算符/括号/引号 */
export const FX_URN_RE = /node:\/\/[\w一-鿿#.\-]+/g

/** fx 草稿里所有 node:// URN 的 [start,end) 区间 —— 胶囊是锁死 token（引用端不可改，用户 2026-09-26 拍板） */
export const fxUrnRanges = (s: string) => {
  const out: { start: number; end: number; urn: string }[] = []
  for (const m of s.matchAll(FX_URN_RE)) out.push({ start: m.index ?? 0, end: (m.index ?? 0) + m[0].length, urn: m[0] })
  return out
}

export type FxToken = {
  start: number
  end: number
  urn: string
  node: string
  field: string
  kind?: string
  dead: boolean
}

/**
 * 可读 token 的字段来源：**解析后**字段（computeNode 的计算结果），不是 schema 静态 t.fields。
 *
 * ★ 批次 17 修 bug（实测铁证：点「数据排序 (2).数量」插入成功、但 fx 栏胶囊包 data-fx-urn=0，
 *   且提交后单元格存的是 `数据排序 (2)·数量` 原文而非 node://…）：
 *   sort / map / output / docx / xlsx 这类**汇聚型**节点，自身 schema `t.fields` 是空的
 *   （字段全部继承自上游 mergeComputed 的结果）。而 RefPanel / trgSources 用的都是
 *   `computeNode(...).fields`（计算后）—— 三处口径原本不一致：
 *     · 面板能看见「数量」→ 点一下插入可读串；
 *     · readableTokenRanges 遍历 t.fields 枚举不到 → 不画胶囊包（用户看不到这是个引用）；
 *     · readableToUrn 同样枚举不到 → 提交时转不回 URN，isPureUrn 判否 → 当普通公式求值，
 *       单元格存进一串无效文本，联动规则自然永远不命中。
 *   口径统一到 computeNode（与面板/联动面板一致）后，三处同时修复。
 *
 * 形状：`{ [nodeId]: { name, kind?, fields: {id,name}[] } }`；
 * 节点算不出来（缺上游/成环）→ 整节点跳过，不猜（与 RefPanel 同口径）。
 */
export type ResolvedFields = Record<string, { name: string; kind?: string; fields: { id: string; name: string }[] }>

export const resolveNodeFields = (
  tables: Record<string, TableSchema> | null | undefined,
  records: Record<string, Row[]> | null | undefined,
  selfId?: string,
): ResolvedFields => {
  const out: ResolvedFields = {}
  if (!tables) return out
  for (const t of Object.values(tables)) {
    if (t.deleted || (selfId && t.id === selfId)) continue
    let fields: { id: string; name: string }[]
    try {
      fields = computeNode(t.id, tables, records ?? {}).fields.map((f) => ({ id: f.id, name: f.name }))
    } catch {
      continue // 缺上游算不出来 → 不列（不猜）
    }
    out[t.id] = { name: t.name, kind: t.kind, fields }
  }
  return out
}

/** 可读 token 定位：枚举全部「表名·参数名」组合在串里找位置（长名优先防前缀吞并，占位去重叠） */
export const readableTokenRanges = (s: string, tables: ResolvedFields | null | undefined): FxToken[] => {
  const out: FxToken[] = []
  if (!tables || !s) return out
  const toks: { str: string; urn: string; node: string; field: string; kind?: string }[] = []
  for (const [id, t] of Object.entries(tables)) {
    for (const f of t.fields) toks.push({ str: `${t.name}·${f.name}`, urn: `node://${id}#${f.id}`, node: t.name, field: f.name, kind: t.kind })
  }
  toks.sort((a, b) => b.str.length - a.str.length)
  const taken = new Array<boolean>(s.length).fill(false)
  for (const tk of toks) {
    let idx = s.indexOf(tk.str)
    while (idx !== -1) {
      let free = true
      for (let i = idx; i < idx + tk.str.length; i++) if (taken[i]) { free = false; break }
      if (free) {
        for (let i = idx; i < idx + tk.str.length; i++) taken[i] = true
        out.push({ start: idx, end: idx + tk.str.length, urn: tk.urn, node: tk.node, field: tk.field, kind: tk.kind, dead: false })
      }
      idx = s.indexOf(tk.str, idx + 1)
    }
  }
  out.sort((a, b) => a.start - b.start)
  return out
}

/** 编辑串里全部胶囊 token：URN 原文（手打/粘贴向下兼容）+ 可读 token，按位置排序、URN 段占位防重叠 */
export const fxTokenRanges = (s: string, tables: ResolvedFields | null | undefined): FxToken[] => {
  const urnToks: FxToken[] = fxUrnRanges(s).map((r) => {
    const u = parseUrn(r.urn)
    const t = u?.scheme === 'node' ? tables?.[u.target] : null
    const fld = u && u.locator ? t?.fields.find((f) => f.id === u.locator || f.name === u.locator) : null
    return {
      start: r.start, end: r.end, urn: r.urn,
      node: t?.name ?? u?.target ?? '', field: fld?.name ?? u?.locator ?? '',
      kind: t?.kind, dead: !!tables && (!t || (!!u?.locator && !fld)),
    }
  })
  const readToks = readableTokenRanges(s, tables)
  // URN 段优先占位（可读枚举撞上 URN 文本时让位——两种形态不会混指同一 token，防御性去重叠）
  const taken = new Array<boolean>(s.length).fill(false)
  for (const r of urnToks) for (let i = r.start; i < r.end; i++) taken[i] = true
  const merged = [...urnToks]
  for (const r of readToks) {
    let free = true
    for (let i = r.start; i < r.end; i++) if (taken[i]) { free = false; break }
    if (free) {
      for (let i = r.start; i < r.end; i++) taken[i] = true
      merged.push(r)
    }
  }
  return merged.sort((a, b) => a.start - b.start)
}

/** URN 串（公式体/引用）→ 可读串：每个 node://id#fid → 表名·参数名（查不到表/参数的保留原文） */
export const urnToReadable = (s: string, tables: ResolvedFields | null | undefined) =>
  s.replace(FX_URN_RE, (urn) => {
    const u = parseUrn(urn)
    if (u?.scheme !== 'node') return urn
    const t = tables?.[u.target]
    const fld = u.locator ? t?.fields.find((f) => f.id === u.locator || f.name === u.locator) : null
    return t && fld ? `${t.name}·${fld.name}` : urn
  })

/** 可读串 → URN 串（提交边界）：把「表名·参数名」token 替换回 node://id#fid（长名优先，非 token 文本原样）。
 *  备忘 2026-09-27 ⑥-b：单遍扫描替换（旧 split/join 全局替换会把相邻 token 换成相邻 URN —— `#price`+`node://`
 *  粘成 `#pricenode://`，FX_URN_RE 字符类不含 `:` 贪婪吃到 `#pricenode` 停 → dead 红胶囊 + `://param_…#price`
 *  泄漏，截图实锤）。且替换产物紧贴 ASCII 词尾（前次 URN 尾/英文数字）→ 自动补 `+`（用户「点几个节点数据
 *  相加」的预期；紧贴 URN 提交后不可解析，信息已丢失，必须在此处拆开）。 */
export const readableToUrn = (s: string, tables: ResolvedFields | null | undefined) => {
  if (!tables || !s.includes('·')) return s
  const toks: { str: string; urn: string }[] = []
  for (const [id, t] of Object.entries(tables)) {
    for (const f of t.fields) toks.push({ str: `${t.name}·${f.name}`, urn: `node://${id}#${f.id}` })
  }
  if (!toks.length) return s
  toks.sort((a, b) => b.str.length - a.str.length)
  let out = ''
  let i = 0
  while (i < s.length) {
    const hit = toks.find((tk) => s.startsWith(tk.str, i))
    if (!hit) {
      out += s[i]
      i++
      continue
    }
    // 粘连保护：token 替换产物紧贴上一段 ASCII 词尾（URN 尾 / 英文数字）→ 补 + 拆开（中文紧贴无害，FX_URN_RE 从 node:// 起匹配）
    if (/[\w#-]$/.test(out)) out += '+'
    out += hit.urn
    i += hit.str.length
  }
  return out
}
