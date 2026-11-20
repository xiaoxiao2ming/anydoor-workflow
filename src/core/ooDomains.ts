/**
 * OO 域外链接 · 域数据注入（纯计算层）
 *
 * 立档：2026-10-17 · 用户问「改了引用的节点的参数 为什么 oo 节点中的数据没有更新」
 *
 * ★ 为什么需要这个文件（真根因）：
 *   设计上的「活引用」是这样成立的 ——
 *     ① 引用格写公式 `=数据对照表_键`
 *     ② 定义名称 `数据对照表_键` → 隐藏表『域数据』!$A$<row>
 *     ③ 上游值变 ⇒ 改 A<row> ⇒ 公式重算
 *   但**保存走的是提取通道**（读值 → 服务端重建 xlsx），重建产物只有 Sheet1 + 值 + 公式，
 *   隐藏表『域数据』与定义名称**双双丢失**。实测用户节点 ooBytes64（3609B）：
 *     sheets = ['Sheet1']、definedName 数量 = 0，而格子里还留着 `<f>数据对照表_键</f>`
 *   ⇒ 文件里根本没有「域」：刷新域值找不到落点，公式也永远解析不出新值。
 *
 * ★ 本文件的职责：打开文档前，按 ooRefs 的**当前值**把 ②重新造出来（幂等）。
 *   这样「改参数 → 重开 OO 节点」就一定看到新值，且：
 *     · 纯补丁：原文件所有部件原样搬，只动 workbook.xml / rels / Content_Types / 域数据表
 *     · 不动引用格：公式 `=数据对照表_键` 原样保留，样式零丢失
 *     · 每次打开都注入 ⇒ 提取通道把表吃掉也不影响（下次打开再造一次）
 *
 * ★ 与 scripts/oo-core.mjs 的口径必须**逐位一致**（否则定义名对不上格子）：
 *     rawKey = `<节点名>.<字段名>`
 *     name   = rawKey 里 '.'→'_'、非法字符→'_'，首字符是数字则加 'R'
 *     row    = 2 + djb2(rawKey) % 1998
 *   改任一处都要同步另一处，否则「引用格显示空」。
 */

import JSZip from 'jszip'
import { evalRowFormula } from '../xlsx/evalFormula'
import { refColorOf } from './refColor'
import { domainRawKey, domainNameOf, domainRowOf } from './refKey'

/** 一条引用值（准备写进『域数据』的） */
export interface DomainEntry {
  /** 原始键：`节点名.字段名`（与 oo-core 的 rawKey 同源） */
  rawKey: string
  /** 定义名称（已 sanitize） */
  name: string
  /** 『域数据』里的行号（2..1999） */
  row: number
  /** 当前值（转成字符串；空值写空串） */
  value: string
}

export interface InjectResult {
  bytes: Uint8Array
  /** 写进去几条 */
  injected: number
  /** 域的落点描述（诊断用） */
  sites: string[]
  /** 没注入的原因（空 = 正常） */
  why?: string
  /** 是否新建了『域数据』表 */
  createdSheet: boolean
  /** 顺带同步了缓存值的公式格数量（引用格 + 下游公式格） */
  cachePatched?: number
  /** 重算明细（诊断：哪些格改成了什么、哪些求值失败） */
  recalcNotes?: string[]
  /** 上了底色的引用格数量 */
  markedCells?: number
  /** 着色明细（诊断：哪个参数用了什么色、落在哪些格） */
  colorNotes?: string[]
}

/** 隐藏表名（与 oo-core 插件里写死的名字一致，改要一起改） */
export const DOMAIN_SHEET = '域数据'

/** rawKey / 定义名称 / 行号的算法在 core/refKey（零依赖，供引用面板复用同一份） */
export { domainRawKey, domainNameOf, domainRowOf } from './refKey'

/** 值 → 单元格文本（null/undefined → 空串，不写 'undefined'） */
export function fmtDomainValue (v: unknown): string {
  if (v == null) return ''
  if (typeof v === 'number') return Number.isFinite(v) ? String(v) : ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  return String(v)
}

/** 由「节点名/字段名/值」列表算好要注入的条目（同 dedupe 行冲突：后者覆盖前者） */
export function buildDomainEntries (
  items: { nodeName: string; fieldName: string; value: unknown }[],
): DomainEntry[] {
  const byRow = new Map<number, DomainEntry>()
  for (const it of items) {
    const rawKey = domainRawKey(it.nodeName, it.fieldName)
    if (!rawKey) continue
    const name = domainNameOf(rawKey)
    const row = domainRowOf(rawKey)
    byRow.set(row, { rawKey, name, row, value: fmtDomainValue(it.value) })
  }
  return [...byRow.values()].sort((a, b) => a.row - b.row)
}

/* ── XML 小工具 ─────────────────────────────────────────────── */

function esc (s: string): string {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

/** 一个 A 列数据格：数字走 <v>，其它走 inlineStr（不引入 sharedStrings 依赖） */
function cellA (row: number, value: string): string {
  if (value === '') return `<c r="A${row}"/>`
  const num = Number(value)
  if (Number.isFinite(num) && String(num) === value.trim() && value.trim() !== '') {
    return `<c r="A${row}"><v>${value.trim()}</v></c>`
  }
  return `<c r="A${row}" t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`
}

function cellB (row: number, key: string): string {
  return `<c r="B${row}" t="inlineStr"><is><t xml:space="preserve">${esc(key)}</t></is></c>`
}

/** 『域数据』整表内容（这张表是机器产物，每次重建：entries 即全部真相） */
export function buildDomainSheetXml (entries: DomainEntry[]): string {
  const rows = entries
    .map((e) => `<row r="${e.row}">${cellA(e.row, e.value)}${cellB(e.row, e.rawKey)}</row>`)
    .join('')
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + `<sheetData>${rows}</sheetData></worksheet>`
}

/**
 * ★ 同步引用格的**缓存值**（第二道保险）。
 *
 * 为什么需要：只注入定义名称的话，引擎"该不该重算"取决于它的加载策略
 * （实测 x2t 转 csv 仍吐旧缓存 456）。把 `<f>数据对照表_键</f>` 旁边的
 * `<v>456</v>` 一并改成当前值 222，则：
 *   引擎重算 → 用定义名称算出 222（对）
 *   引擎不重算 → 直接显示缓存 222（也对）
 * 两条路都指向同一个正确结果，不赌引擎行为。
 *
 * 只动值、不动公式、不动样式；匹配不上就跳过（不强改用户手写的同名公式）。
 */
export function patchFormulaCaches (
  xml: string,
  entries: DomainEntry[],
): { xml: string; patched: number } {
  const byName = new Map(entries.map((e) => [e.name, e]))
  let patched = 0
  const out = xml.replace(/<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g, (whole, attrs: string, tail: string, inner?: string) => {
    if (!inner) return whole
    const fm = /<f[^>]*>([\s\S]*?)<\/f>/.exec(inner)
    if (!fm) return whole
    const fName = domainNameOf(String(fm[1] ?? '').replace(/^=/, '').trim())
    const e = byName.get(fName)
    if (!e) return whole
    // 值：数字不带 t，文本用 t="str"（公式结果的字符串类型）
    const num = Number(e.value)
    const isNum = e.value.trim() !== '' && Number.isFinite(num)
    const vTag = e.value === '' ? '' : `<v>${esc(isNum ? String(num) : e.value)}</v>`
    let nextInner = inner.replace(/<v>[\s\S]*?<\/v>/g, '')   // 去掉旧缓存
    nextInner = nextInner.replace(/(<\/f>)/, `$1${vTag}`)
    if (!/<f[^>]*\/>/.test(nextInner) && !nextInner.includes('</f>')) {
      // 自闭合 <f/>（无公式文本）不该走到这里；保守放弃
      return whole
    }
    let nextAttrs = attrs.replace(/\s+t="[^"]*"/g, '')
    if (!isNum && e.value !== '') nextAttrs += ' t="str"'
    patched++
    return `<c${nextAttrs}>${nextInner}</c>`
  })
  return { xml: out, patched }
}

/* ── 整表重算：把「依赖引用格的公式」也算出来，写回缓存值 ─────────────────
 *
 * ★ 立档 2026-10-17 · 用户三条反馈（原实现只补「公式恰好等于定义名称」的格，覆盖面太窄）：
 *   ① `=参数`            ⇒ 已覆盖
 *   ② `=参数*2`          ⇒ 漏：公式带运算，patchFormulaCaches 匹配不上
 *   ③ `=A1*2`（A1 是引用格）⇒ 漏：下游格的缓存值没人管
 *   现象就是「单独的引用会变，参与公式的不变，再往下引用更不变」。
 *
 * 做法：把表里所有公式格拿出来，按域值 + 现有缓存值做多轮求值（处理任意深依赖链），
 * 结果写回 `<v>` 缓存；公式本身一个字不动。
 *   引擎重算 → 得到同一个值；引擎不重算 → 直接显示我们算的缓存。不赌引擎行为。
 */

export interface RecalcResult {
  xml: string
  /** 缓存被改写的公式格数量 */
  patched: number
  /** 改写明细（诊断用） */
  changed: { ref: string; formula: string; value: string }[]
  /** 求值失败、保留原缓存的格 */
  failed: { ref: string; formula: string; error: string }[]
}

interface RawCell {
  ref: string
  attrs: string
  inner: string
}

/** XML 反转义（公式体里的 &lt; &amp; 必须还原，否则解析器报错） */
function unescXml (s: string): string {
  return String(s)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/** sharedStrings.xml → 字符串数组（文本单元格的值是索引，不解析会把文字当数字） */
export function parseSharedStrings (xml: string): string[] {
  const out: string[] = []
  const re = /<si\b[^>]*>([\s\S]*?)<\/si>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml))) {
    const parts = [...String(m[1]).matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((x) => unescXml(x[1]))
    out.push(parts.join(''))
  }
  return out
}

/** 拆出表里所有 <c>；ref 缺失（数组公式从属）的忽略 */
function parseCells (xml: string): RawCell[] {
  const out: RawCell[] = []
  const re = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml))) {
    const attrs = m[1]
    const ref = (/\br\s*=\s*"([^"]+)"/.exec(attrs)?.[1] ?? '').toUpperCase()
    if (!ref) continue
    out.push({ ref, attrs, inner: m[2] ?? '' })
  }
  return out
}

/** 单元格公式体（已反转义、不带 `=`）；无公式 / 共享公式从属 ⇒ null */
function formulaOf (c: RawCell): string | null {
  const fm = /<f\b([^>]*)>([\s\S]*?)<\/f>/.exec(c.inner)
  if (!fm) {
    // 共享/数组公式的从属格是 <f t="shared" si="0"/>：本身没有文本，跳过
    return null
  }
  const body = unescXml(fm[2]).trim()
  if (!body) return null
  return body.replace(/^=/, '')
}

/** 单元格缓存值 → JS 值（数字/文本/布尔；错误值原样返回字符串） */
function cellRawValue (c: RawCell, sst: string[]): unknown {
  const t = /\bt\s*=\s*"([^"]*)"/.exec(c.attrs)?.[1] ?? 'n'
  if (t === 'inlineStr') {
    return [...c.inner.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((x) => unescXml(x[1])).join('')
  }
  const vm = /<v>([\s\S]*?)<\/v>/.exec(c.inner)
  if (!vm) return undefined
  const txt = unescXml(vm[1])
  if (t === 's') return sst[Number(txt)] ?? ''
  if (t === 'str') return txt
  if (t === 'b') return txt === '1'
  if (t === 'e') return txt
  const n = Number(txt)
  return Number.isFinite(n) ? n : txt
}

/** 定义名称 → 字面量（数字裸写；文本加引号并按 Excel 规则转义） */
export function inlineDomainNames (body: string, nameVals: Map<string, string>): string {
  let out = body
  const names = [...nameVals.keys()].filter(Boolean).sort((a, b) => b.length - a.length)
  for (const n of names) {
    const raw = nameVals.get(n) ?? ''
    const num = Number(raw)
    const lit = raw.trim() !== '' && Number.isFinite(num)
      ? String(num)
      : '"' + raw.replace(/"/g, '""') + '"'
    // 前后不许贴着标识符字符（否则 `参数` 会命中 `参数2`）；后面跟 ( 说明是函数名
    // ★ 边界规则与 findRefCells 共用 nameBoundary（内部自己做正则转义，这里传原始名）
    const re = nameBoundary(n, 'g')
    out = out.replace(re, lit)
  }
  return out
}

/** 值格式化：返回 [t 属性(可为空), <v> 文本(可为空) ] */
function fmtCellValue (v: unknown): { t: string; text: string } {
  if (v === null || v === undefined || v === '') return { t: '', text: '' }
  if (typeof v === 'boolean') return { t: 'b', text: v ? '1' : '0' }
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) return { t: '', text: '' }
    return { t: '', text: String(v) }
  }
  return { t: 'str', text: esc(String(v)) }
}

function sameVal (a: unknown, b: unknown): boolean {
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 1e-12
  return String(a) === String(b)
}

/**
 * 整表重算并把结果写回缓存。
 *
 * @param xml      sheetN.xml 原文
 * @param entries  本次注入的域值（定义名称 → 值）
 * @param sst      sharedStrings（文本单元格要用）
 * @param maxRounds 依赖链轮数上限（引用格 → 下游 → 更下游……每轮推进一层）
 */
export function recalcSheetCaches (
  xml: string,
  entries: DomainEntry[],
  sst: string[] = [],
  maxRounds = 12,
): RecalcResult {
  const cells = parseCells(xml)
  if (!cells.length) return { xml, patched: 0, changed: [], failed: [] }

  const nameVals = new Map(entries.map((e) => [e.name, e.value]))
  const vals: Record<string, unknown> = {}
  const formulaCells: { ref: string; body: string }[] = []
  const cacheOf = new Map<string, string>()

  for (const c of cells) {
    const body = formulaOf(c)
    if (body != null) {
      formulaCells.push({ ref: c.ref, body })
      const v = cellRawValue(c, sst)
      if (v !== undefined) {
        vals[c.ref] = v
        cacheOf.set(c.ref, fmtCellValue(v).text)
      }
      continue
    }
    const v = cellRawValue(c, sst)
    if (v !== undefined) vals[c.ref] = v
  }
  if (!formulaCells.length) return { xml, patched: 0, changed: [], failed: [] }

  const computed = new Map<string, unknown>()
  const failMap = new Map<string, string>()

  for (let round = 0; round < maxRounds; round++) {
    let progress = false
    failMap.clear()
    for (const fc of formulaCells) {
      const filled = inlineDomainNames(fc.body, nameVals)
      const r = evalRowFormula(filled, vals)
      if (!r.ok) {
        failMap.set(fc.ref, r.error ?? '求值失败')
        continue
      }
      if (!(fc.ref in vals) || !sameVal(vals[fc.ref], r.value)) {
        vals[fc.ref] = r.value
        progress = true
      }
      computed.set(fc.ref, r.value)
    }
    if (!progress) break
  }

  /* 写回：只改 <v> 与 t 属性；与旧缓存相同的格不动（保持字节稳定） */
  const changed: { ref: string; formula: string; value: string }[] = []
  const failed: { ref: string; formula: string; error: string }[] = []
  const bodyOf = new Map(formulaCells.map((f) => [f.ref, f.body]))

  const out = xml.replace(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, (whole, attrs: string, inner?: string) => {
    const ref = (/\br\s*=\s*"([^"]+)"/.exec(attrs)?.[1] ?? '').toUpperCase()
    if (!ref || inner == null) return whole
    if (!bodyOf.has(ref)) return whole
    const formula = bodyOf.get(ref) ?? ''
    if (failMap.has(ref)) {
      if (!failed.some((x) => x.ref === ref)) failed.push({ ref, formula, error: failMap.get(ref) ?? '' })
      return whole
    }
    if (!computed.has(ref)) return whole
    const { t, text } = fmtCellValue(computed.get(ref))
    // 与旧缓存一致 ⇒ 不动
    const oldV = /<v>([\s\S]*?)<\/v>/.exec(inner)?.[1]
    const oldT = (/\bt\s*=\s*"([^"]*)"/.exec(attrs)?.[1]) ?? ''
    if (text !== '' && oldV != null && unescXml(oldV) === unescXml(text) && oldT === t) return whole

    let nextInner = inner.replace(/<v>[\s\S]*?<\/v>/g, '')
    nextInner = t === 'str'
      ? nextInner.replace(/(<\/f>)/, `$1<v>${text}</v>`)
      : nextInner.replace(/(<\/f>)/, `$1${text === '' ? '' : `<v>${text}</v>`}`)
    let nextAttrs = attrs.replace(/\st\s*=\s*"[^"]*"/g, '')
    if (t) nextAttrs += ` t="${t}"`
    changed.push({ ref, formula, value: unescXml(text) })
    return `<c${nextAttrs}>${nextInner}</c>`
  })

  return { xml: out, patched: changed.length, changed, failed }
}

/* ── 引用格着色：让「这格是引用来的」一眼可见 ─────────────────────
 *
 * ★ 立档 2026-10-17 · 用户：「能不能对参数引用的地址加个颜色区分」
 *
 * 为什么做在字节层，而不是打开后调 Api.SetFillColor：
 *   ① 打开后再改单元格会**把文档置脏** ⇒ 脏标记误报「未保存」，与刚修好的链路打架；
 *   ② 保存走提取通道（样式全丢），运行时上的色本来也留不住；
 *   ③ 写进文件样式则由「每次打开都注入」自动重放，行为恒定，不需要额外通道。
 *
 * 颜色规则：一个定义名称（= 一个参数）一个底色，同一参数的多处引用同色；
 *   超过调色板容量则循环复用。只上**底色**，字体 / 边框 / 数字格式原样继承原格。
 */

/** 引用格底色与「定义名称 → 颜色」的算法在 core/refColor（零依赖，供引用面板复用同一份） */
export { REF_PALETTE, refColorOf } from './refColor'

export interface RefCellHit {
  /** 单元格地址（大写，如 D5） */
  ref: string
  /** 命中的定义名称 */
  name: string
}

/** 名称边界正则：前后不许贴着标识符字符（`参数` 不能吃 `参数2`），后面跟 ( 视为函数名 */
function nameBoundary (name: string, flags = ''): RegExp {
  const safe = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return new RegExp(`(?<![A-Za-z0-9_.\\u4e00-\\u9fa5])${safe}(?![A-Za-z0-9_.\\u4e00-\\u9fa5(])`, flags)
}

/** 找出「公式里引用了定义名称」的格（= 需要着色的引用格） */
export function findRefCells (xml: string, entries: DomainEntry[]): RefCellHit[] {
  const names = entries.map((e) => e.name).filter(Boolean)
  if (!names.length) return []
  const out: RefCellHit[] = []
  for (const c of parseCells(xml)) {
    const body = formulaOf(c)
    if (body == null) continue
    const hit = names.find((n) => nameBoundary(n).test(body))
    if (hit) out.push({ ref: c.ref, name: hit })
  }
  return out
}

export interface RefStyleTable {
  /** 新 styles.xml（原文件没有则生成最小件） */
  xml: string
  /** 底色 ARGB → fillId */
  fillIdOf: Map<string, number>
  /** `${baseXfId}|${fillId}` → 新 xfId */
  xfIdOf: Map<string, number>
  /** 是否新建了 styles.xml（调用方要顺手注册 rels / Content_Types） */
  created: boolean
  addedFills: number
  addedXfs: number
}

/** 最小可用 styles.xml（原文件缺样式时用；fills 至少要有 none + gray125 两个） */
function minimalStyles (): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + '<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="等线"/><family val="2"/><charset val="134"/></font></fonts>'
    + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    + '<cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles>'
    + '</styleSheet>'
}

/** 复制一个 xf 并换底色（保留数字格式/字体/边框；applyFill=1 让填充生效） */
function composeFillXf (baseTag: string, fillId: number): string {
  let t = baseTag
  if (/\bfillId\s*=\s*"/.test(t)) t = t.replace(/\bfillId\s*=\s*"[^"]*"/, `fillId="${fillId}"`)
  else t = t.replace(/^<xf\b/, `<xf fillId="${fillId}"`)
  t = t.replace(/\s+applyFill\s*=\s*"[^"]*"/g, '')
  if (/\/>\s*$/.test(t)) return t.replace(/\/>\s*$/, ' applyFill="1"/>')
  // 非自闭合（带 <alignment/> <protection/> 子元素）：属性必须插进开标签，不能追加到 </xf> 后面
  return t.replace(/^(<xf\b[^>]*?)(?=\s*\/?>)/, '$1 applyFill="1"')
}

/** 拆出 <cellXfs> 里的每个 xf 标签 */
function splitXfs (xml: string): { tags: string[]; block: string } | null {
  const m = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)
  if (!m) return null
  const tags = m[1].match(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g) ?? []
  return { tags, block: m[0] }
}

/**
 * 给 styles.xml 补上引用格底色所需的 fill / xf。
 *
 * 幂等：同色 fill、以及「同 base xf + 同 fillId」的 xf 存在就复用，反复注入不会膨胀。
 * @param colors 本次真正用到的 ARGB 列表（由 refColorOf 算出并去重）
 */
export function addRefFills (stylesXml: string | null, colors: string[]): RefStyleTable {
  const uniq = [...new Set(colors.filter(Boolean))]
  const base = stylesXml && /<styleSheet\b/.test(stylesXml) ? stylesXml : minimalStyles()
  const created = !(stylesXml && /<styleSheet\b/.test(stylesXml))

  let xml = base
  const fillIdOf = new Map<string, number>()
  const xfIdOf = new Map<string, number>()
  let addedFills = 0
  let addedXfs = 0

  /* ── ① fills：复用同色，否则追加 ── */
  const fm = /<fills\b([^>]*)>([\s\S]*?)<\/fills>/.exec(xml)
  if (!fm) return { xml: base, fillIdOf, xfIdOf, created, addedFills: 0, addedXfs: 0 }

  const fillTags = fm[2].match(/<fill\b[^>]*\/>|<fill\b[^>]*>[\s\S]*?<\/fill>/g) ?? []
  let fillsCount = fillTags.length
  const existingColorFill = new Map<string, number>()
  fillTags.forEach((t, i) => {
    const rgb = /<fgColor\b[^>]*\brgb\s*=\s*"([^"]+)"/i.exec(t)?.[1]?.toUpperCase()
    if (rgb && !existingColorFill.has(rgb)) existingColorFill.set(rgb, i)
  })

  let nextFills = fillTags.slice()
  for (const color of uniq) {
    const hit = existingColorFill.get(color.toUpperCase())
    if (hit != null) { fillIdOf.set(color, hit); continue }
    fillIdOf.set(color, nextFills.length)
    nextFills.push(`<fill><patternFill patternType="solid"><fgColor rgb="${color}"/><bgColor indexed="64"/></patternFill></fill>`)
    existingColorFill.set(color.toUpperCase(), nextFills.length - 1)
    addedFills++
  }
  const headAttrs = fm[1].replace(/\bcount\s*=\s*"\d*"/, '').trim()
  xml = xml.replace(fm[0], `<fills${headAttrs ? ' ' + headAttrs : ''} count="${nextFills.length}">${nextFills.join('')}</fills>`)

  /* ── ② cellXfs：为「默认 xf(0) + 各底色」预建一个 xf，复用同属性者 ── */
  const cx = splitXfs(xml)
  if (!cx) return { xml: base, fillIdOf, xfIdOf, created, addedFills: 0, addedXfs: 0 }

  let xfTags = cx.tags.slice()
  const cxHead = cx.block.match(/<cellXfs\b([^>]*)>/)?.[1] ?? ''

  const composeXf = composeFillXf

  const ensureXf = (baseXfId: number, fillId: number): number => {
    const key = `${baseXfId}|${fillId}`
    const hit = xfIdOf.get(key)
    if (hit != null) return hit
    const baseTag = xfTags[Math.min(Math.max(baseXfId, 0), xfTags.length - 1)] ?? '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    const tag = composeXf(baseTag, fillId)
    const idx = xfTags.indexOf(tag)
    if (idx >= 0) { xfIdOf.set(key, idx); return idx }
    xfTags.push(tag)
    addedXfs++
    xfIdOf.set(key, xfTags.length - 1)
    return xfTags.length - 1
  }

  for (const color of uniq) {
    const fid = fillIdOf.get(color)
    if (fid != null) ensureXf(0, fid)
  }

  const cxAttrs = cxHead.replace(/\bcount\s*=\s*"\d*"/, '').trim()
  xml = xml.replace(cx.block, `<cellXfs${cxAttrs ? ' ' + cxAttrs : ''} count="${xfTags.length}">${xfTags.join('')}</cellXfs>`)

  return { xml, fillIdOf, xfIdOf, created, addedFills, addedXfs }
}

/** 取（必要时新建）一个「基础样式 + 指定底色」的 xf 索引 */
function xfFor (table: RefStyleTable, baseXfId: number, fillId: number): number | null {
  const key = `${baseXfId}|${fillId}`
  const hit = table.xfIdOf.get(key)
  if (hit != null) return hit
  const cx = splitXfs(table.xml)
  if (!cx) return null
  const tags = cx.tags
  const baseTag = tags[Math.min(Math.max(baseXfId, 0), tags.length - 1)]
  if (!baseTag) return null
  const t = composeFillXf(baseTag, fillId)
  const exist = tags.indexOf(t)
  if (exist >= 0) { table.xfIdOf.set(key, exist); return exist }
  const head = cx.block.match(/<cellXfs\b([^>]*)>/)?.[1] ?? ''
  const attrs = head.replace(/\bcount\s*=\s*"\d*"/, '').trim()
  table.xml = table.xml.replace(cx.block, `<cellXfs${attrs ? ' ' + attrs : ''} count="${tags.length + 1}">${tags.join('')}${t}</cellXfs>`)
  table.xfIdOf.set(key, tags.length)
  table.addedXfs++
  return tags.length
}

export interface MarkResult {
  xml: string
  /** 实际上色的格数 */
  marked: number
  /** 想上色但没找到对应样式的格 */
  skipped: string[]
}

/**
 * 给引用格写入底色（只改 `<c>` 的 s 属性，值 / 公式 / 其它属性一律不动）。
 *
 * @param marks 本表内命中的引用格（由 findRefCells 产出；必须按表分开传，地址是表内相对的）
 */
export function markRefCells (xml: string, marks: RefCellHit[], table: RefStyleTable): MarkResult {
  if (!marks.length) return { xml, marked: 0, skipped: [] }
  const colorOfRef = new Map(marks.map((m) => [m.ref, refColorOf(m.name)]))
  const skipped: string[] = []
  let marked = 0

  const out = xml.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs: string, selfClose: string) => {
    const ref = (/\br\s*=\s*"([^"]+)"/.exec(attrs)?.[1] ?? '').toUpperCase()
    const color = colorOfRef.get(ref)
    if (!ref || !color) return whole
    const fillId = table.fillIdOf.get(color)
    if (fillId == null) { skipped.push(ref); return whole }
    const baseXf = Number(/\bs\s*=\s*"(\d+)"/.exec(attrs)?.[1] ?? 0) || 0
    const xf = xfFor(table, baseXf, fillId)
    if (xf == null) { skipped.push(ref); return whole }
    const nextAttrs = /\bs\s*=\s*"\d*"/.test(attrs)
      ? attrs.replace(/\bs\s*=\s*"\d*"/, `s="${xf}"`)
      : `${attrs} s="${xf}"`
    marked++
    return `<c${nextAttrs}${selfClose}>`
  })

  return { xml: out, marked, skipped }
}

/**
 * 给引用格**去掉**底色（2026-10-18 用户指令：标题栏加「着色开关」，关 = 把已有的引用底色清掉）。
 *
 * 做法：把命中格的 s 换成「同 base xf + fill 0（标准 none 填充）」——
 * markRefCells 当初是 base+色fill 的组合 xf，这里把 fill 拆掉即等价还原原样式；
 * 第二次跑幂等（compose 后与原 tag 相同则跳过）。
 * 注意 fillId 0 依赖标准 styles.xml 的 fills[0]=none 约定（OO/Excel 生成的文件都满足）。
 */
export function unmarkRefCells (xml: string, marks: RefCellHit[], table: RefStyleTable): MarkResult {
  if (!marks.length) return { xml, marked: 0, skipped: [] }
  const hitRefs = new Set(marks.map((m) => m.ref))
  const skipped: string[] = []
  let marked = 0

  const out = xml.replace(/<c\b([^>]*?)(\/?)>/g, (whole, attrs: string, selfClose: string) => {
    const ref = (/\br\s*=\s*"([^"]+)"/.exec(attrs)?.[1] ?? '').toUpperCase()
    if (!ref || !hitRefs.has(ref)) return whole
    const curS = Number(/\bs\s*=\s*"(\d+)"/.exec(attrs)?.[1] ?? 0) || 0
    const xf = xfFor(table, curS, 0)
    if (xf == null || xf === curS) { skipped.push(ref); return whole }
    const nextAttrs = /\bs\s*=\s*"\d*"/.test(attrs)
      ? attrs.replace(/\bs\s*=\s*"\d*"/, `s="${xf}"`)
      : `${attrs} s="${xf}"`
    marked++
    return `<c${nextAttrs}${selfClose}>`
  })

  return { xml: out, marked, skipped }
}

/** 下一个可用的 sheetN.xml 部件名（避开已占用的编号） */
export function nextSheetPart (names: string[]): string {
  let max = 0
  for (const n of names) {
    const m = /^xl\/worksheets\/sheet(\d+)\.xml$/.exec(n)
    if (m) max = Math.max(max, Number(m[1]))
  }
  return `xl/worksheets/sheet${max + 1}.xml`
}

/** 下一个可用的 rId（rels 里已用的数字最大值 +1） */
export function nextRelId (relsXml: string): string {
  let max = 0
  const re = /\bId\s*=\s*"rId(\d+)"/g
  let m: RegExpExecArray | null
  while ((m = re.exec(relsXml))) max = Math.max(max, Number(m[1]))
  return `rId${max + 1}`
}

/* ── 主入口 ─────────────────────────────────────────────────── */

/**
 * 把域值注入 xlsx 字节：新建/更新『域数据』表 + 定义名称（缓存写回见 ⑤；
 * fullCalcOnLoad 已于 2026-10-18 深夜移除——引擎不理会它，带标志反而误标「已修改」）。
 *
 * ★ 幂等：同一份文件反复注入，结果稳定（表被重建、定义名称被 upsert）。
 * ★ 不动引用格：单元格里的 `=<定义名称>` 公式原样保留。
 */
export async function injectDomains (bytes: Uint8Array, entries: DomainEntry[], opts?: { markColors?: boolean }): Promise<InjectResult> {
  if (!entries.length) {
    return { bytes, injected: 0, sites: [], createdSheet: false, why: '没有引用值（ooRefs 为空）' }
  }
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(bytes)
  } catch (e) {
    return { bytes, injected: 0, sites: [], createdSheet: false, why: '不是有效的 xlsx（zip 解析失败）' }
  }
  const wbFile = zip.file('xl/workbook.xml')
  const relFile = zip.file('xl/_rels/workbook.xml.rels')
  const ctFile = zip.file('[Content_Types].xml')
  if (!wbFile || !relFile || !ctFile) {
    return { bytes, injected: 0, sites: [], createdSheet: false, why: '缺 workbook.xml / rels / Content_Types' }
  }
  let wbXml = await wbFile.async('string')
  let relXml = await relFile.async('string')
  let ctXml = await ctFile.async('string')

  /* ① 找到（或新建）『域数据』表 */
  const sheetTags = [...wbXml.matchAll(/<sheet\b[^>]*\/?>/g)].map((m) => m[0])
  const hit = sheetTags.find((t) => new RegExp(`name\\s*=\\s*"${DOMAIN_SHEET}"`).test(t))
  let createdSheet = false
  let part: string

  if (hit) {
    const rid = /\br:id\s*=\s*"([^"]+)"/.exec(hit)?.[1] ?? ''
    const relTag = new RegExp(`<Relationship\\b[^>]*Id\\s*=\\s*"${rid}"[^>]*>`).exec(relXml)?.[0] ?? ''
    const target = /\bTarget\s*=\s*"([^"]+)"/.exec(relTag)?.[1] ?? ''
    part = target ? 'xl/' + target.replace(/^\/?xl\//, '').replace(/^\.\//, '') : ''
    if (!part || !zip.file(part)) {
      // 表在 workbook 里登记了但部件缺失（提取通道吃掉的典型形态）→ 补一个部件
      part = nextSheetPart(Object.keys(zip.files))
      zip.file(part, buildDomainSheetXml(entries))
      createdSheet = true
    }
  } else {
    part = nextSheetPart(Object.keys(zip.files))
    const rid = nextRelId(relXml)
    let maxSheetId = 0
    for (const t of sheetTags) {
      const id = /\bsheetId\s*=\s*"(\d+)"/.exec(t)?.[1]
      if (id) maxSheetId = Math.max(maxSheetId, Number(id))
    }
    // state="hidden"：用户看不见，也不出现在正常视图里
    wbXml = wbXml.replace(
      /<\/sheets>/,
      `<sheet name="${DOMAIN_SHEET}" sheetId="${maxSheetId + 1}" state="hidden" r:id="${rid}"/></sheets>`,
    )
    relXml = relXml.replace(
      /<\/Relationships>/,
      `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${part.replace(/^xl\//, '')}"/></Relationships>`,
    )
    if (!ctXml.includes(`/xl/${part.replace(/^xl\//, '')}`)) {
      ctXml = ctXml.replace(
        /<\/Types>/,
        `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
      )
    }
    zip.file(part, buildDomainSheetXml(entries))
    createdSheet = true
  }

  /* ② 表已存在 ⇒ 重建内容（这张表只由我们维护） */
  if (!createdSheet) zip.file(part, buildDomainSheetXml(entries))

  /* ③ 定义名称 upsert：先删掉所有指向『域数据』的旧名，再写本次的 */
  const dnRe = /<definedName\b[^>]*>[\s\S]*?<\/definedName>/g
  const keptDn = (wbXml.match(dnRe) ?? []).filter((t) => !t.includes(DOMAIN_SHEET))
  const mine = entries.map((e) =>
    `<definedName name="${esc(e.name)}">'${DOMAIN_SHEET}'!$A$${e.row}</definedName>`)
  const dnBlock = `<definedNames>${keptDn.join('')}${mine.join('')}</definedNames>`
  // 先摘掉旧的 <definedNames> 整块（schema 要求它紧跟 sheets 之后、calcPr 之前）
  let wbNext = wbXml.replace(/<definedNames>[\s\S]*?<\/definedNames>/g, '')
  if (!/<sheets\b[^>]*\/>/.test(wbNext) && wbNext.includes('</sheets>')) {
    wbNext = wbNext.replace(/<\/sheets>/, `</sheets>${dnBlock}`)
  }

  /* ④ ★ 不再写 fullCalcOnLoad（2026-10-18 深夜）：本机 OO 引擎不理会它（不重算，
   *   显示走缓存——这正是 recalcSheetCaches 双保险存在的理由），而带这个标志的文件
   *   打开后会被引擎标成「已修改」⇒ 用户什么都没改、关闭却提示保存。
   *   缓存正确性由 ⑤ 的写回保证（唯一事实源）。旧文件若带此标志，顺手摘掉。 */
  wbNext = wbNext.replace(/\s+fullCalcOnLoad="[^"]*"/g, '')

  zip.file('xl/workbook.xml', wbNext)
  zip.file('xl/_rels/workbook.xml.rels', relXml)
  zip.file('[Content_Types].xml', ctXml)
  // 旧 calcChain 会和新定义名称打架 ⇒ 直接丢掉（引擎会按 fullCalcOnLoad 重建）
  if (zip.file('xl/calcChain.xml')) zip.remove('xl/calcChain.xml')

  /* ⑤ 整表重算 + 写回缓存（不赌引擎是否重算，见 recalcSheetCaches 注释）
   *   覆盖三种引用形态：`=参数` / `=参数*2` / `=A1*2`（A1 是引用格）以及更深的依赖链。 */
  const sstFile = zip.file('xl/sharedStrings.xml')
  const sst = sstFile ? parseSharedStrings(await sstFile.async('string')) : []
  let cachePatched = 0
  const recalcNotes: string[] = []
  const sheetAfter: { part: string; xml: string; hits: RefCellHit[] }[] = []
  for (const name of Object.keys(zip.files)) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) continue
    if (name === part) continue
    const f = zip.file(name)
    if (!f) continue
    const xml = await f.async('string')
    const pr = recalcSheetCaches(xml, entries, sst)
    if (pr.patched) {
      cachePatched += pr.patched
      recalcNotes.push(...pr.changed.slice(0, 6).map((c) => `${c.ref} ${c.formula} → ${c.value}`))
      /* ★★★ 2026-10-18 深夜修复「点了更新但表格不变」的真凶 ★★★
       *   着色改版重构时把 `zip.file(name, pr.xml)` 写回弄丢了（原 sheetAfter 只喂着色提取，
       *   着色改纯视觉层后变成 `void sheetAfter`）⇒ 重算只在内存里发生，落盘的 xlsx 缓存
       *   仍是旧值（留档 2026-10-18-20-10-18 实锤：域数据=8/9 正确、A1 缓存还是 222），
       *   而本机 OO 引擎不理会 fullCalcOnLoad（见 recalcSheetCaches 注释「不赌引擎重算」）
       *   ⇒ 用户永远看到旧值。必须写回！ */
      zip.file(name, pr.xml)
    }
    if (pr.failed.length) {
      recalcNotes.push(...pr.failed.slice(0, 3).map((c) => `${c.ref} 求值失败(${c.error})`))
    }
    // 引用格识别在重算后的文本上做（公式未变，但缓存已是最新）
    sheetAfter.push({ part: name, xml: pr.xml, hits: findRefCells(pr.xml, entries) })
  }

  /* ⑥ 引用格着色 —— 2026-10-18 用户指令改版：
   *   旧实现把底色写进 xlsx 的 styles.xml（打开即有，但切换开关必须重载文档才见效，
   *   用户实测「点开关画面不变，要手动重开」完全不可接受）。
   *   新实现：颜色不再写文件，改为**纯视觉覆盖层**（oo-core 的 __ooRefMarks 通道，
   *   与行列高亮同机制）——开关即时生效、不重开、不脏文档。
   *   injectDomains 从此只负责域数据注入；引用格清单由 extractRefMarks 另行提取。 */
  const markedCells = 0
  const colorNotes: string[] = []
  void sheetAfter

  const out = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  return {
    bytes: out,
    injected: entries.length,
    cachePatched,
    recalcNotes,
    markedCells,
    colorNotes,
    sites: entries.map((e) => `${DOMAIN_SHEET}!A${e.row} ← ${e.rawKey}=${e.value || '(空)'}`),
    createdSheet,
  }
}

/** 从 xlsx 里读出『域数据』现有内容（诊断/自检用：不入库、不改文件） */
export async function readDomains (bytes: Uint8Array): Promise<{ row: number; key: string; value: string }[]> {
  let zip: JSZip
  try { zip = await JSZip.loadAsync(bytes) } catch { return [] }
  const wb = zip.file('xl/workbook.xml')
  const rel = zip.file('xl/_rels/workbook.xml.rels')
  if (!wb || !rel) return []
  const wbXml = await wb.async('string')
  const relXml = await rel.async('string')
  const tag = [...wbXml.matchAll(/<sheet\b[^>]*\/?>/g)].map((m) => m[0])
    .find((t) => new RegExp(`name\\s*=\\s*"${DOMAIN_SHEET}"`).test(t))
  if (!tag) return []
  const rid = /\br:id\s*=\s*"([^"]+)"/.exec(tag)?.[1] ?? ''
  const relTag = new RegExp(`<Relationship\\b[^>]*Id\\s*=\\s*"${rid}"[^>]*>`).exec(relXml)?.[0] ?? ''
  const target = /\bTarget\s*=\s*"([^"]+)"/.exec(relTag)?.[1] ?? ''
  const part = 'xl/' + target.replace(/^\/?xl\//, '').replace(/^\.\//, '')
  const f = zip.file(part)
  if (!f) return []
  const xml = await f.async('string')
  const out: { row: number; key: string; value: string }[] = []
  for (const m of xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
    const row = Number(m[1])
    const body = m[2]
    const a = /<c\b[^>]*\br="A\d+"[^>]*>([\s\S]*?)<\/c>/.exec(body)?.[1] ?? ''
    const b = /<c\b[^>]*\br="B\d+"[^>]*>([\s\S]*?)<\/c>/.exec(body)?.[1] ?? ''
    const val = /<v>([\s\S]*?)<\/v>/.exec(a)?.[1] ?? /<t[^>]*>([\s\S]*?)<\/t>/.exec(a)?.[1] ?? ''
    const key = /<t[^>]*>([\s\S]*?)<\/t>/.exec(b)?.[1] ?? ''
    if (key) out.push({ row, key, value: val })
  }
  return out.sort((x, y) => x.row - y.row)
}

/** XML 文本反转义（readDomains 返回的是原文；比对前必须还原，否则 `&amp;` 之类会永远「有变化」） */
function unesc (s: string): string {
  return String(s)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/**
 * 比对「表格里引用格的**显示值**」与「画布当前值」，回答一个业务问题：
 *   **打开后用户看到的数会不会变？**
 *
 * ★ 2026-10-18 深夜改版（用户质问「我没有修改参数引用 还要提示我更新」）：
 *   旧口径比对『域数据』表，而该表会被保存通道剥掉（提取通道重建必丢）⇒
 *   表一缺失就把全部引用当「有变化」⇒ 参数明明没改也每次弹窗。
 *   新口径：直接读**公式格的缓存 <v>**（= 用户眼睛看到的数，唯一显示事实源），
 *   与画布当前值比对：
 *     · 公式格缓存 == 当前值 → 无变化（哪怕域表缺失——重建它不会改变显示值，静默补）
 *     · 表格里根本没有引用该名字的公式格 → 视觉上无从变化 → 不算变化
 *     · 缓存 != 当前值 → 真变化，弹窗列 was → now
 */
export async function diffDomainValues (
  bytes: Uint8Array,
  entries: DomainEntry[],
): Promise<{ hasSheet: boolean; total: number; changed: { name: string; was: string; now: string }[] }> {
  const total = entries.length
  if (!total) return { hasSheet: false, total: 0, changed: [] }
  let zip: JSZip
  try { zip = await JSZip.loadAsync(bytes) } catch {
    return { hasSheet: false, total, changed: [] }
  }
  const wbFile = zip.file('xl/workbook.xml')
  const hasSheet = !!(wbFile && /<sheet\b[^>]*name="域数据"/.test(await wbFile.async('string')))
  /* 每个 name 的「显示值」= 表格里引用它的公式格缓存（取第一处；与 patchFormulaCaches 同源匹配） */
  const nameSet = new Set(entries.map((e) => e.name))
  const displayed = new Map<string, string>()
  for (const part of Object.keys(zip.files)) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(part)) continue
    const xml = await zip.file(part)!.async('string')
    for (const m of xml.matchAll(/<c\b[^>]*>([\s\S]*?)<\/c>/g)) {
      const inner = m[1]
      const fm = /<f[^>]*>([\s\S]*?)<\/f>/.exec(inner)
      if (!fm) continue
      const fName = domainNameOf(String(fm[1] ?? '').replace(/^=/, '').trim())
      if (!fName || displayed.has(fName) || !nameSet.has(fName)) continue
      const v = /<v>([\s\S]*?)<\/v>/.exec(inner)?.[1]
      const t = /<is>\s*<t[^>]*>([\s\S]*?)<\/t>/.exec(inner)?.[1]
      displayed.set(fName, unesc(v ?? t ?? ''))
    }
  }
  const changed: { name: string; was: string; now: string }[] = []
  for (const e of entries) {
    const was = displayed.get(e.name)
    if (was === undefined) continue /* 表格里没有引用格 → 视觉上无从变化 */
    if (was !== e.value) changed.push({ name: e.name, was: was || '(空)', now: e.value || '(空)' })
  }
  return { hasSheet, total, changed }
}

/**
 * 摘掉 fullCalcOnLoad（轻量补丁：只动 workbook.xml，其余部件原样）。
 *
 * 供「不检测 / 检测后提示」等**不注入**的打开路径使用——旧文件里残留的
 * fullCalcOnLoad 会让引擎把刚打开的文档标成「已修改」⇒ 什么都没改、
 * 关闭却提示保存（2026-10-18 用户实测）。
 */
export async function stripFullCalcFlag (bytes: Uint8Array): Promise<Uint8Array> {
  try {
    const zip = await JSZip.loadAsync(bytes)
    const wb = zip.file('xl/workbook.xml')
    if (!wb) return bytes
    const xml = await wb.async('string')
    if (!/fullCalcOnLoad/.test(xml)) return bytes
    const next = xml.replace(/(<calcPr\b[^>]*?)\s+fullCalcOnLoad="[^"]*"/g, '$1')
    if (next === xml) return bytes
    zip.file('xl/workbook.xml', next)
    return await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  } catch {
    return bytes
  }
}

/* ── 引用格覆盖层数据（2026-10-18 着色开关改版）─────────────────────
 * 颜色不再写进 xlsx（旧方案切换开关必须重载文档，用户实测不可接受）。
 * 这里只提取「引用格清单 + 表格几何」，交给 oo-core 在引擎画面上画覆盖层
 * （与行列高亮同机制：纯视觉、不脏文档、开关即时生效）。 */

export interface RefMark {
  /** 工作表名（如 Sheet1） */
  sheet: string
  /** 单元格地址（如 D5） */
  ref: string
  /** 底色 ARGB（同一参数同色） */
  color: string
}

/** 合并单元格区间（0 起，含端点） */
export interface MergedRange { r1: number; c1: number; r2: number; c2: number }

/* ★ 2026-10-18 用户反馈「亮显颜色容易偏移」后的几何模型改版：
 *   旧模型直接存 px（列 round(w*7+5)、行 round(ht*4/3)、headW/headH 写死），
 *   三个致命问题：
 *     ① 隐藏行列（<col hidden>/<row hidden>）被当成正常宽高 ⇒ 其后每一列/行整体位移；
 *     ② 行高 pt→px 用 4/3，而 OO 实测 15pt 渲染 19px（非 20）⇒ 每行累计 1px 漂移；
 *     ③ 公式估算的全局缩放与 OO 实际渲染可能不一致。
 *   新模型只保存 xlsx 原始单位（列=字符数、行=pt）+ 隐藏标记 + 合并区，
 *   由注入层用「活动格实测像素」现场标定比例（kCol/kRow），并以活动格为原点做相对定位
 *   （行/列标头宽度天然消掉），做到与单元格实际位置一致。 */
export interface SheetGeom {
  /** 默认列宽（字符数，xlsx 原始单位） */
  defColCh: number
  /** 默认行高（pt，xlsx 原始单位） */
  defRowPt: number
  /** 自定义列宽：列索引(0 起) → 字符数；隐藏列记 0 */
  colCh: Record<string, number>
  /** 自定义行高：行索引(0 起) → pt；隐藏行记 0 */
  rowPt: Record<string, number>
  /** 合并单元格区间（色块要铺满整块合并区） */
  merges: MergedRange[]
}

/** 'D5' → { row: 4, col: 3 }（0 起）；容忍 $ 绝对引用 */
export function addrToRowCol (ref: string): { row: number; col: number } {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(ref.toUpperCase())
  if (!m) return { row: -1, col: -1 }
  let col = 0
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64)
  return { row: Number(m[2]) - 1, col: col - 1 }
}

/** 列索引 → 字母（0 → A） */
export function colToLetters (col: number): string {
  let s = ''
  let c = col
  while (c >= 0) { s = String.fromCharCode(65 + (c % 26)) + s; c = Math.floor(c / 26) - 1 }
  return s
}

/**
 * 从 xlsx 字节提取「引用格清单 + 工作表几何」（不改文件，只读）。
 * 几何取自**第一个命中引用格的工作表**：默认宽高 + cols 自定义列宽 + row 自定义行高
 * + 隐藏行列 + 合并区。单位保持 xlsx 原样（列=字符、行=pt），换算交给注入层现场标定。
 */
export async function extractRefMarks (bytes: Uint8Array, entries: DomainEntry[]): Promise<{ marks: RefMark[]; geom: SheetGeom | null }> {
  if (!entries.length) return { marks: [], geom: null }
  let zip: JSZip
  try { zip = await JSZip.loadAsync(bytes) } catch { return { marks: [], geom: null } }
  const wb = zip.file('xl/workbook.xml')
  const rel = zip.file('xl/_rels/workbook.xml.rels')
  if (!wb || !rel) return { marks: [], geom: null }
  const wbXml = await wb.async('string')
  const relXml = await rel.async('string')
  /* sheet 名 → part（rId 中转） */
  const relTarget = (rid: string): string => {
    const t = new RegExp(`<Relationship\\b[^>]*Id\\s*=\\s*"${rid}"[^>]*>`).exec(relXml)?.[0] ?? ''
    return 'xl/' + (/\bTarget\s*=\s*"([^"]+)"/.exec(t)?.[1] ?? '').replace(/^\/?xl\//, '').replace(/^\.\//, '')
  }
  const sheets: { name: string; part: string }[] = []
  for (const m of wbXml.matchAll(/<sheet\b[^>]*\/?>/g)) {
    const name = /(?<=\bname\s*=\s*")[^"]*(?=")/.exec(m[0])?.[0] ?? ''
    const rid = /\br:id\s*=\s*"([^"]+)"/.exec(m[0])?.[1] ?? ''
    if (name && rid) sheets.push({ name, part: relTarget(rid) })
  }
  const marks: RefMark[] = []
  let geom: SheetGeom | null = null
  for (const sh of sheets) {
    const f = zip.file(sh.part)
    if (!f) continue
    const xml = await f.async('string')
    const hits = findRefCells(xml, entries)
    if (!hits.length) continue
    for (const h of hits) marks.push({ sheet: sh.name, ref: h.ref, color: refColorOf(h.name) })
    if (!geom) geom = readSheetGeom(xml)
  }
  return { marks, geom }
}

/** 从一个 sheet XML 读出几何（默认宽高 / cols / rows / 隐藏行列 / 合并区）。纯函数，便于单测。 */
export function readSheetGeom (xml: string): SheetGeom {
  const fmt = /<sheetFormatPr\b[^>]*\/?>/.exec(xml)?.[0] ?? ''
  const defRowPt = Number(/defaultRowHeight\s*=\s*"([\d.]+)"/.exec(fmt)?.[1] ?? 15)
  const defColCh = Number(/defaultColWidth\s*=\s*"([\d.]+)"/.exec(fmt)?.[1] ?? 8.43)
  /* 列：<col min max width hidden> —— min/max 是 1 起区间，要展开；hidden 记 0 宽 */
  const colCh: Record<string, number> = {}
  for (const cm of xml.matchAll(/<col\b([^>]*?)\/?>/g)) {
    const a = cm[1]
    const hidden = /\bhidden\s*=\s*"1"/.test(a)
    const min = Number(/\bmin\s*=\s*"(\d+)"/.exec(a)?.[1] ?? 1)
    const max = Number(/\bmax\s*=\s*"(\d+)"/.exec(a)?.[1] ?? min)
    const wRaw = /\bwidth\s*=\s*"([\d.]+)"/.exec(a)?.[1]
    if (!hidden && wRaw == null) continue
    const ch = hidden ? 0 : Number(wRaw)
    for (let i = min; i <= max && i <= 512; i++) colCh[String(i - 1)] = ch
  }
  /* 行：<row r ht hidden> —— 只有显式 ht / hidden 的行需要记 */
  const rowPt: Record<string, number> = {}
  for (const rm of xml.matchAll(/<row\b([^>]*?)\/?>/g)) {
    const a = rm[1]
    const idx = Number(/\br\s*=\s*"(\d+)"/.exec(a)?.[1] ?? 0)
    if (!idx) continue
    const hidden = /\bhidden\s*=\s*"1"/.test(a)
    const htRaw = /\bht\s*=\s*"([\d.]+)"/.exec(a)?.[1]
    if (!hidden && htRaw == null) continue
    rowPt[String(idx - 1)] = hidden ? 0 : Number(htRaw)
  }
  /* 合并区：色块要铺满整块，而不是只盖左上角那一格 */
  const merges: MergedRange[] = []
  for (const mm of xml.matchAll(/<mergeCell\b[^>]*\bref\s*=\s*"([^"]+)"/g)) {
    const parts = mm[1].split(':')
    const a = addrToRowCol(parts[0])
    const b = addrToRowCol(parts[1] ?? parts[0])
    if (a.row < 0 || b.row < 0) continue
    merges.push({
      r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col),
      r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col),
    })
  }
  return { defColCh, defRowPt, colCh, rowPt, merges }
}
