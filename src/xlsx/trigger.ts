/**
 * 参数联动触发规则 —— 纯计算层
 *
 * 口径（用户 2026-10-05 定案）：**不做实时格级联动，改为「规则命中才更新」**。
 *   外部参数改了 → 匹配规则 → 命中才触发写值 + 重算 + 补丁保存。
 *   没命中规则 = 不动文件。这比"每次都重算一遍"更省、更可控、也更符合人的意图。
 *
 * 为什么要独立一层：
 *   规则判定 = 纯逻辑（比较、阈值、开关），值得单测与复用；
 *   WorkbookEditor 只该管交互与落盘，不该长出 if-else 判断森林。
 *
 * 与既有资产的对接（复用，不另起炉灶）：
 *   · 引用寻址沿用 `ref/urn.ts` 的 URN 协议（node:// / param:// / file:// / api:// / const:）
 *   · 写值沿用 `@office-kit/xlsx` 的 ensureCellByCoord + setCellValue
 *   · 重算沿用 `./recompute` 的 recomputeSheet（已支持 node:// 参与求值）
 *   · 落盘沿用 `./patchSave`（补丁不重建，公式 <f>+<v> 双写）
 *
 * 纯计算层铁律：不 import electron / node: / document，浏览器与 Node 同构。
 */

/* ────────────────── 规则模型 ────────────────── */

/** 条件运算符（刻意收窄：只放业务真正常用的，别做表达式引擎） */
export type CondOp =
  | 'always' // 无条件（手动触发/兜底）
  | 'changed' // 值变了
  | 'gt' // 新值 > 阈值
  | 'gte' // 新值 >= 阈值
  | 'lt'
  | 'lte'
  | 'eq'
  | 'ne'
  | 'between' // 闭区间 [lo, hi]
  | 'outside' // 开区间 (lo, hi) 之外
  | 'nonEmpty'
  | 'empty'
  | 'isTrue'
  | 'isFalse'

export interface TriggerCondition {
  op: CondOp
  /** 阈值 / 比较基准 */
  value?: number | string | boolean
  /** between/outside 的上界 */
  hi?: number | string
  /** 容差：数值比较时按此判定相等，防浮点抖动把"没变"判成"变了" */
  epsilon?: number
}

/** 命中后要执行的动作 */
export type TriggerAction =
  | { kind: 'writeCell'; sheet: string; coord: string } // 把来源值写进单元格
  | { kind: 'writeRange'; sheet: string; ref: string } // 写进区域的首格（数组走逐格另算）
  | { kind: 'recompute'; sheet?: string } // 触发重算（sheet 省略 = 全部）
  | { kind: 'commit' } // 重算 + 补丁保存

export interface TriggerRule {
  id: string
  /** 规则名，面板显示 */
  name: string
  /** 开关：关掉的规则不参与判定（保留配置，不丢） */
  enabled: boolean
  /** 监听哪个来源（URN 串，如 param://管径 / node://n1#rows）。'change' = 任何来源变了 */
  watch: string
  condition: TriggerCondition
  action: TriggerAction
  /**
   * 优先级，数字小 = 先跑。用于确定同一次变更里多规则命中时的执行顺序
   * （典型：先写值再重算，所以"重算"规则要排在"写值"规则后面）。
   */
  priority?: number
  /** 去抖窗口（毫秒）：同一来源在这个窗口内的多次改动只触发一次。0 = 不去抖 */
  debounceMs?: number
  /** 备注，仅面板显示 */
  note?: string
}

/* ────────────────── 变更差分 ────────────────── */

export interface SourceChange {
  urn: string
  before: unknown
  after: unknown
}

export interface DiffResult {
  changes: SourceChange[]
  /** 实际发生变化的 URN 集合（before≠after 的那些） */
  changedUrns: Set<string>
}

/** 值是否真的变了。浮点按 epsilon 容差，避免 0.1+0.2 !== 0.3 式的假变化 */
export function valueChanged(before: unknown, after: unknown, epsilon = 0): boolean {
  if (before === after) return false

  // ★ 空归一闸门必须排在类型闸门之前：null / undefined / '' 三者在业务上都是"没填"，
  //   若交给类型闸门判，'' 会因为 toNum('')=null 落到 str 族，被误判成"从空变成了字符串"。
  const be = before == null || before === ''
  const ae = after == null || after === ''
  if (be && ae) return false

  // ★ 类型闸门：不同类型族一律算变化，绝不合并。
  //   toNum 会把 false→0、'1.0'→1，若不先拦类型，`0 → false`、`'1.0' → 1` 会被误判成"没变"，
  //   而这两者在业务上都是实打实的语义变化（布尔开关翻转、文本型数字变成数值型）。
  if (kindOf(before) !== kindOf(after)) return true

  // 两侧都是数值型（含数字串）→ 走容差比较
  const bn = toNum(before)
  const an = toNum(after)
  if (bn != null && an != null) {
    const eps = Math.abs(epsilon)
    return Math.abs(an - bn) > eps
  }

  // 对象 / 数组走 JSON 比对（顺序敏感，符合"数组顺序变了就是变了"的直觉）
  if (typeof before === 'object' || typeof after === 'object') {
    return JSON.stringify(before ?? null) !== JSON.stringify(after ?? null)
  }
  return true
}

/** 值的类型族：用于「同族才可能相等」的闸门 */
function kindOf(v: unknown): string {
  if (v === null) return 'null'
  if (v === undefined) return 'null' // null / undefined 同族（都是"空"）
  if (typeof v === 'boolean') return 'bool'
  if (typeof v === 'number') return 'num'
  if (typeof v === 'string') return toNum(v) != null ? 'num' : 'str' // 数字串与数字同族
  return typeof v // obj / arr
}

function toNum(v: unknown): number | null {
  if (typeof v === 'number' && Number.isFinite(v)) return v
  if (typeof v === 'string' && v.trim() !== '') {
    const n = Number(v)
    if (Number.isFinite(n)) return n
  }
  if (typeof v === 'boolean') return v ? 1 : 0
  return null
}

/** 两份取值快照求差分。快照形状：{ [urn]: value } */
export function diffSnapshots(before: Record<string, unknown>, after: Record<string, unknown>): DiffResult {
  const urns = new Set([...Object.keys(before ?? {}), ...Object.keys(after ?? {})])
  const changes: SourceChange[] = []
  const changedUrns = new Set<string>()
  for (const urn of urns) {
    const b = before?.[urn]
    const a = after?.[urn]
    // 默认 epsilon=0：这里只认"真变了"，容差留给规则条件自己定
    if (valueChanged(b, a, 0)) {
      changes.push({ urn, before: b, after: a })
      changedUrns.add(urn)
    }
  }
  return { changes, changedUrns }
}

/* ────────────────── 条件判定 ────────────────── */

/** 单条规则是否命中本次变更 */
export function evalCondition(
  cond: TriggerCondition,
  change: { before: unknown; after: unknown }
): { hit: boolean; reason?: string } {
  const eps = Math.abs(cond.epsilon ?? 0)
  const a = change.after
  const b = change.before
  const an = toNum(a)
  const bn = toNum(b)

  switch (cond.op) {
    case 'always':
      return { hit: true }
    case 'changed':
      return { hit: valueChanged(b, a, eps), reason: '值需发生变化' }
    case 'nonEmpty':
      return { hit: !(a == null || a === ''), reason: '新值需非空' }
    case 'empty':
      return { hit: a == null || a === '', reason: '新值需为空' }
    case 'isTrue':
      return { hit: a === true || a === 1, reason: '新值需为真' }
    case 'isFalse':
      return { hit: a === false || a === 0, reason: '新值需为假' }

    // 以下都是数值型条件 —— 取不到数就判不命中（不猜、不静默通过）
    case 'gt':
    case 'gte':
    case 'lt':
    case 'lte':
    case 'eq':
    case 'ne': {
      if (an == null || toNum(cond.value) == null) {
        return { hit: false, reason: '新值或阈值不是数字' }
      }
      const t = toNum(cond.value)!
      switch (cond.op) {
        case 'gt':
          return { hit: an > t, reason: `${an} > ${t}` }
        case 'gte':
          return { hit: an >= t - eps, reason: `${an} >= ${t}` }
        case 'lt':
          return { hit: an < t, reason: `${an} < ${t}` }
        case 'lte':
          return { hit: an <= t + eps, reason: `${an} <= ${t}` }
        case 'eq':
          return { hit: Math.abs(an - t) <= eps, reason: `${an} == ${t}（容差 ${eps}）` }
        default:
          return { hit: Math.abs(an - t) > eps, reason: `${an} != ${t}` }
      }
    }
    case 'between':
    case 'outside': {
      if (an == null || toNum(cond.value) == null || toNum(cond.hi) == null) {
        return { hit: false, reason: '新值或上下界不是数字' }
      }
      const lo = toNum(cond.value)!
      const hi = toNum(cond.hi)!
      // 上下界写反时自动纠正（用户填反不该导致规则永远不命中）
      const [a2, b2] = lo <= hi ? [lo, hi] : [hi, lo]
      const inside = an >= a2 - eps && an <= b2 + eps
      return cond.op === 'between'
        ? { hit: inside, reason: `需落在 [${a2}, ${b2}]` }
        : { hit: !inside, reason: `需落在 [${a2}, ${b2}] 之外` }
    }
    default: {
      // 未知算子：保守不命中，并把原因带出去（UI 可显示"规则写错了"）
      const never = (cond as { op: string }).op
      return { hit: false, reason: `未知条件算子：${String(never)}` }
    }
  }
}

/* ────────────────── 规则求值 ────────────────── */

export interface RuleHit {
  rule: TriggerRule
  change: SourceChange
  reason: string
}

export interface RuleSkip {
  rule: TriggerRule
  change: SourceChange
  reason: string
}

export interface EvalResult {
  hits: RuleHit[]
  skips: RuleSkip[]
  /** 命中的动作，按规则优先级 + 声明顺序排好，可直接顺序执行 */
  actions: { rule: TriggerRule; action: TriggerAction }[]
}

/**
 * 拿一次变更去匹配全部规则。
 *
 * watch 匹配口径：
 *   · watch === 'change'  → 任何来源变了都算命中候选
 *   · watch === 某 URN     → 只有该 URN 变了才算
 *   · watch 以 '#' 结尾    → 前缀匹配（node://n1# 能吃到 node://n1#rows / node://n1#cols）
 */
export function evalRules(rules: TriggerRule[], changes: SourceChange[]): EvalResult {
  const hits: RuleHit[] = []
  const skips: RuleSkip[] = []

  for (const rule of rules ?? []) {
    if (!rule || rule.enabled === false) continue
    const myChanges = changes.filter((c) => watchMatch(rule.watch, c.urn))
    if (!myChanges.length) continue

    for (const ch of myChanges) {
      const r = evalCondition(rule.condition ?? { op: 'always' }, ch)
      if (r.hit) hits.push({ rule, change: ch, reason: r.reason ?? '' })
      else skips.push({ rule, change: ch, reason: r.reason ?? '' })
    }
  }

  // 排序：优先级小者先；同优先级按规则声明顺序（稳定）
  const idx = new Map<string, number>()
  ;(rules ?? []).forEach((r, i) => idx.set(r.id, i))
  hits.sort((a, b) => {
    const pa = a.rule.priority ?? 100
    const pb = b.rule.priority ?? 100
    if (pa !== pb) return pa - pb
    return (idx.get(a.rule.id) ?? 0) - (idx.get(b.rule.id) ?? 0)
  })

  // 动作去重：同一 (kind, sheet, coord/ref) 只跑一次（多条规则写同一格时后者是多余的）
  const seen = new Set<string>()
  const actions: { rule: TriggerRule; action: TriggerAction }[] = []
  for (const h of hits) {
    const k = actionKey(h.rule.action)
    if (seen.has(k)) continue
    seen.add(k)
    actions.push({ rule: h.rule, action: h.rule.action })
  }

  return { hits, skips, actions }
}

/**
 * watch 匹配口径：
 *   · watch === 'change' / '*'  → 任何来源变了都算命中候选
 *   · watch 以 '#' 结尾          → 前缀匹配（node://n1# 能吃到 node://n1#rows / node://n1#cols）
 *   · watch 以 '://' 结尾        → 同 scheme 全部匹配（param:// 天然是所有参数的前缀）
 *   · 其余                       → 精确匹配
 */
export function watchMatch(watch: string, urn: string): boolean {
  const w = String(watch ?? '').trim()
  if (!w || w === 'change' || w === '*') return true
  if (w.endsWith('#')) return urn.startsWith(w)
  // scheme 前缀（param:// / node:// / file:// / api://）→ 该 scheme 下所有来源
  if (w.endsWith('://')) return urn.startsWith(w)
  return w === urn
}

function actionKey(a: TriggerAction): string {
  if (a.kind === 'writeCell') return `writeCell|${a.sheet}!${a.coord}`
  if (a.kind === 'writeRange') return `writeRange|${a.sheet}!${a.ref}`
  if (a.kind === 'recompute') return `recompute|${a.sheet ?? '*'}`
  return 'commit'
}

/* ────────────────── 去抖 ────────────────── */

/**
 * 去抖闸门：同一 key 在窗口内的重复请求只放行第一次。
 *
 * 为什么需要：用户连拖三次滑块 → 三次 diff → 若都落盘，xlsx 会被写三遍（每次都全量 zip 重建）。
 * 去抖把「一段时间内的连续改动」合并成一次更新，正是"更新时更新"想要的语义。
 *
 * 纯逻辑（收 nowFn 注入），单测可确定性驱动；浏览器传 Date.now。
 */
export class DebounceGate {
  private lastAt = new Map<string, number>()

  constructor(private nowFn: () => number = () => Date.now()) {}

  /**
   * 请求放行。
   * @returns true = 放行（可执行更新）；false = 被窗口挡住（已合并到上一次）
   */
  allow(key: string, windowMs: number): boolean {
    if (!windowMs || windowMs <= 0) return true
    const now = this.nowFn()
    const last = this.lastAt.get(key)
    if (last != null && now - last < windowMs) return false
    this.lastAt.set(key, now)
    return true
  }

  /** 手动清空（"立即更新"按钮用：先清闸门再执行，保证必过） */
  reset(key?: string): void {
    if (key == null) this.lastAt.clear()
    else this.lastAt.delete(key)
  }
}

/* ────────────────── 计划生成（只算不动） ────────────────── */

export interface WriteOp {
  sheet: string
  coord: string
  value: unknown
  /** 值来源 URN，便于面板回显"这格是哪来的" */
  from: string
}

export interface TriggerPlan {
  writes: WriteOp[]
  recomputeSheets: Set<string>
  /** 是否需要重算全部表（recompute 不带 sheet 的规则命中时为 true） */
  recomputeAll: boolean
  needCommit: boolean
  hits: RuleHit[]
  /** 面板展示用：每条命中的人话说明 */
  log: string[]
}

/**
 * 由「命中的动作 + 本次变更」生成一份执行计划。
 * 纯算不动 —— 面板可以先给用户看"这次会改哪几格"，确认后再落盘。
 */
export function planTrigger(
  rules: TriggerRule[],
  changes: SourceChange[],
  resolveValue: (urn: string) => unknown
): TriggerPlan {
  const { hits, actions } = evalRules(rules, changes)

  const writes: WriteOp[] = []
  const written = new Set<string>()
  const recomputeSheets = new Set<string>()
  let recomputeAll = false
  let needCommit = false
  const log: string[] = []

  // 建 urn → 本次新值 的索引：写值动作要知道该写什么
  const afterByUrn = new Map<string, unknown>()
  for (const c of changes) afterByUrn.set(c.urn, c.after)

  /**
   * 取「这条规则该写的值」。
   * watch='change' 时 watch 本身不是真实 URN（它只是"任何来源"的通配），
   * 键直查必然落空 —— 必须回溯到「本次到底是哪条变更触发的」。
   */
  const valueFor = (rule: TriggerRule): { value: unknown; from: string } => {
    if (rule.watch === 'change' || rule.watch === '*' || !rule.watch) {
      const first = changes[0]
      if (first) return { value: first.after, from: first.urn }
      return { value: undefined, from: rule.watch }
    }
    if (afterByUrn.has(rule.watch)) return { value: afterByUrn.get(rule.watch), from: rule.watch }
    // 前缀 watch（node://n1#）：取第一条匹配的变更
    const matched = changes.find((c) => watchMatch(rule.watch, c.urn))
    if (matched) return { value: matched.after, from: matched.urn }
    // 常驻来源：不在本次变更里，回查 URN
    return { value: resolveValue(rule.watch), from: rule.watch }
  }

  for (const { rule, action } of actions) {
    switch (action.kind) {
      case 'writeCell': {
        const key = `${action.sheet}!${action.coord}`
        if (written.has(key)) break
        const { value: val, from } = valueFor(rule)
        writes.push({ sheet: action.sheet, coord: action.coord, value: val, from })
        written.add(key)
        log.push(`写值 ${key} ← ${from} = ${fmt(val)}`)
        break
      }
      case 'writeRange': {
        // 区域写入：来源是数组时逐格铺，否则只写首格（不猜，行为可预期）
        const { value: src, from } = valueFor(rule)
        const arr = Array.isArray(src) ? src : src == null ? [] : [src]
        const coords = expandCoords(action.ref, arr.length)
        arr.forEach((v, i) => {
          const co = coords[i]
          if (!co) return
          const key = `${action.sheet}!${co}`
          if (written.has(key)) return
          writes.push({ sheet: action.sheet, coord: co, value: v, from })
          written.add(key)
        })
        if (!arr.length) log.push(`区域 ${action.sheet}!${action.ref}：来源为空，跳过`)
        else log.push(`区域 ${action.sheet}!${action.ref} 写入 ${arr.length} 格 ← ${from}`)
        break
      }
      case 'recompute': {
        if (action.sheet) {
          recomputeSheets.add(action.sheet)
          log.push(`重算工作表 ${action.sheet}`)
        } else {
          recomputeAll = true
          log.push('重算全部工作表')
        }
        break
      }
      case 'commit': {
        needCommit = true
        log.push('落盘（重算 + 补丁保存）')
        break
      }
    }
  }

  // 有写值但没显式重算规则 → 自动补一次重算（否则公式缓存值还是旧的，用户会以为没生效）
  if (writes.length && !recomputeAll && !recomputeSheets.size) {
    for (const w of writes) recomputeSheets.add(w.sheet)
    log.push('（自动补：写值后需重算，否则公式缓存仍是旧值）')
  }

  return { writes, recomputeSheets, recomputeAll, needCommit, hits, log }
}

function fmt(v: unknown): string {
  if (v == null) return '（空）'
  if (typeof v === 'object') return JSON.stringify(v)
  return String(v)
}

/** `A1:D10` 之类按数量展开为坐标列表；支持单格 */
function expandCoords(ref: string, need: number): string[] {
  const s = String(ref ?? '').trim()
  if (!s) return []
  if (!s.includes(':')) return [s]
  const [a, b] = s.split(':')
  // 复用 a1 模块的解析口径（这里只做纯字符串展开，避免跨模块循环依赖）
  const pa = /^([A-Z]+)(\d+)$/.exec(a.trim().toUpperCase())
  const pb = /^([A-Z]+)(\d+)$/.exec(b.trim().toUpperCase())
  if (!pa || !pb) return [a]
  const c1 = colNum(pa[1]), r1 = Number(pa[2])
  const c2 = colNum(pb[1]), r2 = Number(pb[2])
  const [cLo, cHi] = c1 <= c2 ? [c1, c2] : [c2, c1]
  const [rLo, rHi] = r1 <= r2 ? [r1, r2] : [r2, r1]
  const out: string[] = []
  for (let r = rLo; r <= rHi && out.length < need; r++) {
    for (let c = cLo; c <= cHi && out.length < need; c++) {
      out.push(colLetters(c) + r)
    }
  }
  return out
}

function colNum(letters: string): number {
  let n = 0
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}

function colLetters(n: number): string {
  let s = ''
  let x = n
  while (x > 0) {
    const r = (x - 1) % 26
    s = String.fromCharCode(65 + r) + s
    x = Math.floor((x - 1) / 26)
  }
  return s
}

/* ────────────────── 预设规则（面板一键加） ────────────────── */

export const PRESET_RULES: Omit<TriggerRule, 'id'>[] = [
  {
    name: '管径变 → 写 A2',
    enabled: true,
    watch: 'param://管径',
    condition: { op: 'changed' },
    action: { kind: 'writeCell', sheet: 'Sheet1', coord: 'A2' },
    priority: 10,
    note: '典型场景：外部参数改了才落盘，没改就完全不碰文件',
  },
  {
    name: '写值后重算',
    enabled: true,
    watch: 'change',
    condition: { op: 'always' },
    action: { kind: 'recompute' },
    priority: 90,
    note: '优先级排在写值之后，保证公式拿到的是新值',
  },
  {
    name: '落盘',
    enabled: true,
    watch: 'change',
    condition: { op: 'always' },
    action: { kind: 'commit' },
    priority: 99,
    debounceMs: 800,
    note: '去抖 800ms：连续拖动滑块只写一次文件',
  },
]
