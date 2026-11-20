/**
 * 节点控件状态 ⇄ 工程文件 的桥接层（2026-10-06）。
 *
 * ★ 这个文件解决的是用户明确提出的第 5 项：
 *   「控件的需求（因为控件对照着节点）」。
 *
 * ── 为什么要桥接，而不是直接把 localStorage 换掉 ──
 *   现有控件状态散落在 7 类 localStorage 键里（见下方 KEYMAP）。它们是**运行时**存储，
 *   不属于工程：换机器全丢、发给同事布局不同、两台机器来回编辑互相覆盖。
 *   但也不能直接删掉 localStorage —— 那是几十处读写点，一次性换掉必然出回归。
 *   故采取**双写 + 工程优先**的过渡口径：
 *     · 保存工程时：从 localStorage 采集进 viewState 写进节点文件
 *     · 打开工程时：把 viewState 回灌 localStorage（让现有控件自然读到）
 *     · 采集不到 / 工程里没有：保持 localStorage 原值（回退，而非清空）
 *   这样新增节点类型时，只要往 KEYMAP 加一行就自动随工程走，不用改控件代码。
 *
 * ── 为什么键位映射要写成一张表（KEYMAP）而不是散在各处 ──
 *   7 类键 × 2 个字段（tab / refPanelOpen / rightPanelOpen / activeSheet / triggerRules）
 *   如果在采集处和恢复处各写一遍 switch，两份逻辑必然漂移：
 *   采集时把 ooref 映射到 refPanelOpen，恢复时却读 refpanel ——
 *   症状是「OO 节点的引用面板关不上」，且只在跨机器打开时才出现。
 *   一张表 + 双向遍历，从结构上排除这种漂移。
 *
 * ── 分层铁律 ──
 *   本文件是**纯计算层**：不 import localStorage、不碰 document、不 import electron。
 *   localStorage 通过 `ViewStateKV` 适配器注入，NodeEditor 等 UI 传真 localStorage，
 *   测试传 Map。这样本文件能被 vitest 直接跑（无 jsdom 依赖）。
 */

import type { NodeViewState } from './projectFormat'
import type { TableSchema } from './schema'

/* ────────────────────────── KV 适配器 ────────────────────────── */

/**
 * localStorage 的最小接口。
 *
 * 只取 get/set/remove 三件事，且全部可能抛（隐私模式、配额满）——
 * 所以桥接层所有调用都包 try，不是「理论上不会抛」。
 */
export interface ViewStateKV {
  getItem(key: string): string | null
  setItem(key: string, value: string): void
  removeItem(key: string): void
}

/** 浏览器实现：真 localStorage。构造时不取 localStorage（可能抛），每次调用现取。 */
export function browserKV(): ViewStateKV {
  const ls = (): Storage => {
    // ★ 每次现取而不是模块加载时缓存：
    //   测试里会替换 globalThis.localStorage，模块加载时缓存就锁死了旧引用。
    const g = globalThis as unknown as { localStorage?: Storage }
    if (!g.localStorage) throw new Error('localStorage 不可用')
    return g.localStorage
  }
  return {
    getItem: (k) => ls().getItem(k),
    setItem: (k, v) => ls().setItem(k, v),
    removeItem: (k) => ls().removeItem(k),
  }
}

/** 永不抛的兜底实现：隐私模式下退化「控件状态不持久化」，但绝不影响主流程 */
export function nullKV(): ViewStateKV {
  return { getItem: () => null, setItem: () => {}, removeItem: () => {} }
}

/* ────────────────────────── 键位映射表（单一真相源）────────────────────────── */

/** 控件状态字段在 NodeViewState 里的落点 */
export type ViewField = 'tab' | 'refPanelOpen' | 'rightPanelOpen' | 'activeSheet' | 'triggerRules'

/**
 * ★ 键位映射表：localStorage 键前缀 ⇄ NodeViewState 字段。
 *
 * 每一项的 `prefix` 后面拼 `:<nodeId>`。**新增控件状态只改这一张表**。
 *
 * 刻意保留多个前缀映射到同一字段（如 ooref / refpanel / wrp-refpanel 都 →refPanelOpen）：
 *   那三个键分属不同编辑器（NodeEditor / WorkbookEditor / WordReportPanel），
 *   各自历史上各存各的。它们语义上都是「这个节点的参数引用面板开合」，
 *   归一到同一字段后，跨机器打开时三处表现一致 —— 这正是「控件随工程走」的意义。
 *   代价是同一节点若被两种编辑器读过，采集时按数组顺序取第一个存在的值（见 collectNodeViewState）。
 */
export interface ViewKeyBinding {
  /** localStorage 键前缀（不含 `:<nodeId>`） */
  prefix: string
  /** 落到NodeViewState 的哪个字段 */
  field: ViewField
  /** 值的编码方式 */
  codec: 'raw' | 'bool' | 'json'
  /**
   * 判定「默认开」还是「默认关」—— 仅用于 bool 编码的**缺省**推断。
   *
   * ★ 为什么要记默认值，而不是只存 true/false：
   *   `ooref` 的语义是 `!== '0'`（缺省开），`rightpanel` 是 `=== '1'`（缺省关）。
   *   若采集时把默认值也写进 viewState，等于把「用户从没动过这个面板」
   *   固化成一次显式设置 —— 以后改默认值时，老工程会一直卡在旧默认上。
   *   所以只在 localStorage 里**确实有值**时才采集，不推断默认值。
   */
  note: string
}

/** ★ 全部已知键位映射。改这里 = 改「什么状态随工程走」的口径。 */
export const VIEW_KEYMAP: readonly ViewKeyBinding[] = [
  { prefix: 'ne-tab', field: 'tab', codec: 'raw', note: 'NodeEditor 当前标签页' },
  { prefix: 'ooref', field: 'refPanelOpen', codec: 'bool', note: 'OO 节点左栏数据引用面板' },
  { prefix: 'refpanel', field: 'refPanelOpen', codec: 'bool', note: 'WorkbookEditor 引用栏' },
  { prefix: 'wrp-refpanel', field: 'refPanelOpen', codec: 'bool', note: 'Word 面板左栏' },
  { prefix: 'rightpanel', field: 'rightPanelOpen', codec: 'bool', note: 'WorkbookEditor 右栏' },
  { prefix: 'wrp-rightpanel', field: 'rightPanelOpen', codec: 'bool', note: 'Word 面板右栏' },
  { prefix: 'wb-sheet', field: 'activeSheet', codec: 'raw', note: '工作簿当前工作表名' },
  { prefix: 'trigrules', field: 'triggerRules', codec: 'json', note: '联动触发规则（xlsx TriggerPanel）' },
]

/** 反查：某字段该写回哪些键（一个字段可能对应多个前缀，全部写） */
function prefixesOf(field: ViewField): string[] {
  return VIEW_KEYMAP.filter((b) => b.field === field).map((b) => b.prefix)
}

/** 取某字段的编解码口径（同一字段口径必须一致，取首个即可） */
function codecOf(field: ViewField): ViewKeyBinding['codec'] {
  return VIEW_KEYMAP.find((b) => b.field === field)?.codec ?? 'raw'
}

/** 节点在画布上的折叠态：键名与其它控件不同源，故单列 */
export const COLLAPSE_PREFIX = 'node-collapsed'

/* ────────────────────────── 采集：localStorage → NodeViewState ────────────────────────── */

/**
 * 采集单个节点的控件状态。
 *
 * ★ 关键约定：**只采集 localStorage 里确实存在的键**，不给的字段一律 undefined。
 *   理由见 ViewKeyBinding.note：把默认值也固化进工程，会让老工程永远卡在旧默认上。
 *   副作用正好是我们要的 —— 工程文件里 viewState 通常只有一两个字段，很干净。
 *
 * @returns 采集不到任何东西时返回 undefined（而不是空对象），
 *   空对象与 undefined 在 buildNodeFile 里的待遇不同（后者才沿用旧值）。
 */
export function collectNodeViewState(nodeId: string, kv: ViewStateKV): NodeViewState | undefined {
  const out: NodeViewState = {}

  for (const b of VIEW_KEYMAP) {
    const raw = safeGet(kv, `${b.prefix}:${nodeId}`)
    if (raw === null) continue
    const v = decode(b.codec, raw)
    if (v === undefined) continue
    // 同字段多前缀：先到先得（VIEW_KEYMAP 数组顺序即优先级）
    if (out[b.field] !== undefined) continue
    out[b.field] = v as never
  }

  const col = safeGet(kv, `${COLLAPSE_PREFIX}:${nodeId}`)
  if (col !== null && (col === '1' || col === '0')) out.collapsed = col === '1'

  return Object.keys(out).length > 0 ? out : undefined
}

/** 采集全量节点。nodeIds 显式传入而不是遍历 tables —— 采集面按调用方的口径走 */
export function collectAllViewState(
  nodeIds: readonly string[],
  kv: ViewStateKV,
): Record<string, NodeViewState> {
  const out: Record<string, NodeViewState> = {}
  for (const id of nodeIds) {
    const vs = collectNodeViewState(id, kv)
    if (vs) out[id] = vs
  }
  return out
}

/**
 * 从 tables 的键自动收集（保存工程的常见入口）。
 *
 * 含回收站节点吗？**含** —— 回收站节点的折叠态与面板状态同样属于工程。
 * 少存一个节点的状态，用户从回收站恢复节点时会发现面板全展开，
 * 而这类「恢复后不一样了」的投诉极难自查。
 */
export function collectViewStateFromTables(
  tables: Record<string, TableSchema>,
  kv: ViewStateKV,
): Record<string, NodeViewState> {
  return collectAllViewState(Object.keys(tables), kv)
}

/* ────────────────────────── 恢复：NodeViewState → localStorage ────────────────────────── */

/**
 * 把工程里的 viewState 回灌 localStorage。
 *
 * ★ 严格「只写不删」：
 *   工程里没有的字段**不动** localStorage。理由：老工程没有 viewState，
 *   若此时把 localStorage 清空，用户在自己机器上存了一年的面板状态会被
 *   「打开一个老工程」这个无害动作抹掉 —— 那是纯粹的破坏。
 *   回退方向永远应该是「保留现状」，不是「以工程为准清空」。
 *
 * @returns 实际写入的键数（供测试与UI 日志用）
 */
export function applyNodeViewState(nodeId: string, vs: NodeViewState | undefined, kv: ViewStateKV): number {
  if (!vs) return 0
  let n = 0

  for (const b of VIEW_KEYMAP) {
    const v = vs[b.field]
    if (v === undefined) continue
    const enc = encode(b.codec, v)
    if (enc === undefined) continue
    if (safeSet(kv, `${b.prefix}:${nodeId}`, enc)) n++
  }

  if (vs.collapsed !== undefined) {
    if (safeSet(kv, `${COLLAPSE_PREFIX}:${nodeId}`, vs.collapsed ? '1' : '0')) n++
  }

  return n
}

/** 全量回灌（打开工程时调一次） */
export function applyAllViewState(
  viewStates: Record<string, NodeViewState | undefined> | undefined,
  kv: ViewStateKV,
): number {
  if (!viewStates) return 0
  let n = 0
  for (const [id, vs] of Object.entries(viewStates)) n += applyNodeViewState(id, vs, kv)
  return n
}

/**
 * 打开工程前的**冲突审计**：哪些节点的本地状态与工程里不一致。
 *
 * ★ 为什么需要它而不是直接覆盖：
 *   同一个节点在两台机器上都编辑过（本地记着「引用栏关着」，工程里记着「开着」）
 *   是常态，不是异常。静默按工程覆盖 = 用户在自己机器上刚调好的布局被工程改回去，
 *   而他完全不知道为什么。
 *   把分歧报出来让上层决定（提示 / 按工程 / 保留本地），这才是能被信任的行为。
 */
export interface ViewStateConflict {
  nodeId: string
  field: ViewField | 'collapsed'
  local: string
  fromProject: string
}

export function auditViewStateConflicts(
  nodeId: string,
  vs: NodeViewState | undefined,
  kv: ViewStateKV,
): ViewStateConflict[] {
  const out: ViewStateConflict[] = []
  const local = collectNodeViewState(nodeId, kv)
  if (!vs || !local) return out

  const cmp = (f: ViewField | 'collapsed', l: unknown, p: unknown) => {
    if (l === undefined || p === undefined) return
    const ls = String(l)
    const ps = String(p)
    if (ls !== ps) out.push({ nodeId, field: f, local: ls, fromProject: ps })
  }

  for (const b of VIEW_KEYMAP) {
    cmp(b.field, local[b.field], vs[b.field])
  }
  cmp('collapsed', local.collapsed, vs.collapsed)
  return out
}

/* ────────────────────────── 工具 ────────────────────────── */

/** localStorage 读：任何异常都当「没存过」 */
function safeGet(kv: ViewStateKV, key: string): string | null {
  try {
    return kv.getItem(key)
  } catch {
    return null
  }
}

/** localStorage 写：任何异常都当「存不了」（配额满/隐私模式），不冒泡阻断保存 */
function safeSet(kv: ViewStateKV, key: string, value: string): boolean {
  try {
    kv.setItem(key, value)
    return true
  } catch {
    return false
  }
}

/**
 * 解码。
 *
 * @returns undefined 表示「这个值不该被采纳」（坏格式 / 空串），
 *   调用方据此跳过该字段 —— 宁可少存一个面板状态，也不要往工程里塞脏值。
 */
function decode(codec: ViewKeyBinding['codec'], raw: string): unknown {
  if (raw === '') return undefined
  switch (codec) {
    case 'raw':
      return raw
    case 'bool':
      // 只认 '1' / '0'。其他值（如旧版本遗留的 'true'）不当成布尔，
      // 宁可跳过让控件走自己的默认值 —— 猜错会让面板开合与用户上次操作相反。
      if (raw === '1') return true
      if (raw === '0') return false
      return undefined
    case 'json': {
      try {
        const v = JSON.parse(raw)
        return v === undefined ? undefined : v
      } catch {
        return undefined
      }
    }
    default:
      return undefined
  }
}

/** 编码；返回 undefined 表示「这个值无法表达成该 codec 的键值对，跳过」 */
function encode(codec: ViewKeyBinding['codec'], value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined
  switch (codec) {
    case 'raw':
      return typeof value === 'string' ? value : String(value)
    case 'bool':
      if (typeof value !== 'boolean') return undefined
      return value ? '1' : '0'
    case 'json':
      try {
        return JSON.stringify(value)
      } catch {
        // 循环引用 / BigInt：跳过这一个字段，不影响其余字段落盘
        return undefined
      }
    default:
      return undefined
  }
}

/** 供 UI 显示「哪些状态会随工程走」的自检（避免文档与实现漂移） */
export function viewStateCoverage(): { field: ViewField | 'collapsed'; prefixes: string[]; note: string }[] {
  const fields = [...new Set(VIEW_KEYMAP.map((b) => b.field)), 'collapsed' as const]
  return fields.map((f) => ({
    field: f,
    prefixes: f === 'collapsed' ? [COLLAPSE_PREFIX] : prefixesOf(f),
    note: f === 'collapsed' ? '画布折叠态' : VIEW_KEYMAP.find((b) => b.field === f)?.note ?? '',
  }))
}
