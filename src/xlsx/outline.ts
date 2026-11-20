/**
 * 章节分级 / 章节号编排 —— 纯计算层
 *
 * 口径来源：`docs/标题分级与目录大纲实现方案.md`（2026-10-04拍板）
 *
 * ── 唯一真相 ────────────────────────────────────────────────
 *   `level` 字段是唯一真相。章节号（no）、左侧目录树、xlsx 标记列
 *   全部是它的派生视图。任何结构变更都必须走commitStructure()，
 *   它是「钳制 → 重算 → 重建树 → 落盘标记」的唯一收口。
 *
 * ── 为什么不叫 Row ───────────────────────────────────────────
 *   `core/store.ts` 已export `type Row = Record<string, unknown>`
 *   （节点引擎的行数据），同名会造成跨模块混淆。本模块一律用
 *   `OutlineRow`，别在 import 处改名绕过。
 *
 * ── 铁律 ───────────────────────────────────────────────────
 *   1. 章节号是**派生数据**，永不手工维护、永不作为源数据存储。
 *   2. 树与表格是同一份 OutlineRow[] 的两个渲染，**严禁两份状态互同步**。
 *      buildTree() 每次从 rows 全量重建，不做增量。
 *   3. 拖动 = 剪贴操作（摘子树 → 插入），子树跟走是免费得到的。
 *   4. 全量重算几千行是毫秒级，**不做增量**。
 *
 * 纯计算层铁律：不 import electron / node: / document，浏览器与 Node 同构。
 */

/* ────────────────── 类型 ────────────────── */

/** 标题层级：0 = 普通数据行，1..5 = 一到五级标题 */
export type RowLevel = 0 | 1 | 2 | 3 | 4 | 5

/** 层级上限（RowLevel 的最大值，用于钳制边界） */
export const MAX_LEVEL = 5

/** 一行 = 标题或普通数据行 */
export interface OutlineRow {
  /** 稳定唯一标识。拖动靠它定位，不用数组下标（下标一挪就全错） */
  id: string
  /** 【唯一真相】标题层级，0 = 普通数据行 */
  level: RowLevel
  /** 标题文字（level > 0 时有效） */
  title: string
  /** 【派生】章节号，如 '1.2.1'。由 renumber() 写入，外部不许赋值 */
  no: string
  /** 该标题是否折叠。折叠状态存在行上（跟着行走），不存全局集合—— */
  /** 存在全局集合会在拖动后错乱（§8 已列此坑） */
  collapsed: boolean
  /** 业务数据列 */
  cells: Record<string, unknown>
}

/** 目录树节点 —— 从OutlineRow[] 派生，不独立存状态 */
export interface TreeNode {
  rowId: string
  level: RowLevel
  no: string
  title: string
  collapsed: boolean
  children: TreeNode[]
  /** 在 OutlineRow[] 中的下标，供点击定位 */
  rowIndex: number
  /** 是否有标题子级（区别于「只挂了普通数据行」） */
  hasChildren: boolean
  /** 本节点下挂的普通数据行数（直属，不含子标题下的） */
  bodyCount: number
}

/** 扁平化后的树节点（虚拟滚动 / 键盘导航用，带深度） */
export interface FlatNode extends TreeNode {
  depth: number
  /** 该节点在 rows 里占的行段长度（含自身与所有后代） */
  span: number
}

/* ────────────────── 计数器：章节号重算 ────────────────── */

/**
 * renumber —— 章节号全量重算（计数器栈）
 *
 * 两处要点，缺一个就出错：
 *   ① 遇到 level=0 普通行**不清计数器** —— 普通行挂在上级标题下，编号延续。
 *   ② `counters.length = r.level` 截断更深层级 —— 回到浅层级时深计数必须作废，
 *      否则会出现 '1.3.1' 之后又冒出 '1.3.1' 的重复。
 *
 * 就地改写 rows[i].no。O(n)，不做增量。
 */
export function renumber(rows: OutlineRow[]): void {
  const counters: number[] = []
  for (const r of rows) {
    if (r.level === 0) {
      r.no = ''
      continue
    }
    counters[r.level - 1] = (counters[r.level - 1] ?? 0) + 1
    // 截断更深层级的计数器 —— 这一行不能省
    counters.length = r.level
    r.no = counters.join('.')
  }
}

/** 只算不写，供预览 / 校验用（不污染传入的 rows） */
export function computeNumbers(rows: readonly OutlineRow[]): string[] {
  const out: string[] = []
  const counters: number[] = []
  for (const r of rows) {
    if (r.level === 0) {
      out.push('')
      continue
    }
    counters[r.level - 1] = (counters[r.level - 1] ?? 0) + 1
    counters.length = r.level
    out.push(counters.join('.'))
  }
  return out
}

/* ────────────────── 跳级钳制 ────────────────── */

/**
 * normalizeLevels —— 钳制非法跳级
 *
 * 规则：上一行是 L，则本行最多 L+1。一级下直接出现三级 → 自动钳为二级。
 * 绝不允许出现 '1.0.0.1' 这类脏编号。
 *
 * 普通行（level=0）不改变 prev —— 它归属于上一个标题层级。
 *
 * @returns 是否发生过修正（true → UI 需给 toast 提示）
 */
export function normalizeLevels(rows: OutlineRow[]): boolean {
  let prev = 0
  let changed = false
  for (const r of rows) {
    if (r.level > 0) {
      const max = prev + 1
      if (r.level > max) {
        r.level = (max <= 0 ? 1 : max > MAX_LEVEL ? MAX_LEVEL : max) as RowLevel
        changed = true
      }
      prev = r.level
    }
    // level === 0：普通行不改变 prev
  }
  return changed
}

/* ────────────────── 标记列解析（xlsx 往返载体）────────────────── */

/**
 * parseLevel —— 解析 xlsx 标记列
 *
 * 主用数字（1/2/3），兼容字符模式（# / ## / ###，纯 Excel 里手工维护用）。
 * 数字优先：'#' 的 Number() 是 NaN，'2' 的 match 也是 null，两者不冲突。
 *
 * 脏输入（全角＃、带空格、越界）一律归0 = 普通数据行，**不抛错**
 * —— 打开用户的文件不能因为一格脏数据就失败。
 */
export function parseLevel(raw: unknown): RowLevel {
  const s = String(raw ?? '').trim()
  if (s === '') return 0
  const num = Number(s)
  if (Number.isInteger(num) && num >= 0 && num <= MAX_LEVEL) return num as RowLevel
  // 兼容字符模式（含全角＃，用户从别处拷来常见）
  const m = s.match(/^[#＃]+/)
  if (m) return Math.min(m[0].length, MAX_LEVEL) as RowLevel
  return 0
}

/** 标记列 → 数字（保存到 xlsx 的 A 列） */
export function formatLevel(level: RowLevel): number {
  return level
}

/* ────────────────── 子树提取与移动 ────────────────── */

/**
 * takeSubtree —— 摘出 rows[fromIdx] 及其全部后代行
 *
 * 「后代行」= 紧随其后、level 严格大于本行 level 的连续行段。
 * level=0 的普通行不是标题，不作为拖动容器的边界依据（见 §8：普通行不可作容器）。
 */
export function takeSubtree(rows: OutlineRow[], fromIdx: number): OutlineRow[] {
  if (fromIdx < 0 || fromIdx >= rows.length) return []
  const end = spanEnd(rows, fromIdx)
  return rows.splice(fromIdx, end - fromIdx)
}

/** 插入行段到目标下标（toIdx 为 rows 移除后的新下标） */
export function insertSubtree(rows: OutlineRow[], seg: OutlineRow[], toIdx: number): void {
  if (seg.length === 0) return
  rows.splice(Math.max(0, Math.min(toIdx, rows.length)), 0, ...seg)
}

/**
 * spanEnd —— 子树行段的结束下标（不含）。takeSubtree / subtreeSpan / 折叠可见性 共用。
 *
 * ★ 这里对方案 §3.3 伪码做了一处**必要修正**（2026-10-04 实测发现，照抄会丢行）：
 *   方案原写法 `while (rows[end].level > L) end++`，但普通数据行（level=0）会**立刻中断**扫描。
 *   实测 level 序列 1,2,0,2,0,0,1,2,1：拖 '1' 时，'1.1' 之后紧跟普通行，
 *   该普通行使条件失败 → span 只截到 2，标题下挂的普通数据行被整段丢掉。
 *
 * 正确语义：普通数据行 level=0 是**标题的正文**，不是边界，遇到就继续。
 *   真正的边界只有：越界，或遇到 level>0 且 <= L 的行（同级标题 / 上级标题）。
 *
 * ⚠️ 对称约束：既然 level=0 会被「跟随」吃掉，它就**不能作为子树起点**——
 *   否则拖一个普通行会连带把它上方标题的后代行全带走。见 assertMovable。
 */
function spanEnd(rows: readonly OutlineRow[], fromIdx: number): number {
  const L = rows[fromIdx].level
  // 普通行（level=0）不是容器，它的"段"一路到下一个标题为止（实际拖动已被 assertMovable 拦住，
  // 这里仍给出确定行为，避免 visibleRowIndexes 之类的只读调用方拿到意外值）
  let end = fromIdx + 1
  while (end < rows.length) {
    const lv = rows[end].level
    if (lv > 0 && lv <= L) break
    end++
  }
  return end
}

/** 子树行段长度（含自身与所有后代），不修改数组 */
export function subtreeSpan(rows: readonly OutlineRow[], fromIdx: number): number {
  if (fromIdx < 0 || fromIdx >= rows.length) return 0
  return spanEnd(rows, fromIdx) - fromIdx
}

/** 移动结果 */
export interface MoveResult {
  ok: boolean
  /** 实际移动的行数 */
  moved: number
  /** 是否发生了跳级修正（→ 需toast） */
  clamped: boolean
  /** 失败原因（ok=false 时有值） */
  reason?: string
}

/**
 * moveRows —— 移动的统一入口
 *
 * 拖表格行、拖树节点、调顺序，**全部归结这一个函数**（方案铁律 3）。
 * 顺序不能变：摘出 → 钳制 → 插入 → 全局钳制 → 重算编号。
 *
 * 目标语义由调用方算好（树上三语义：上 1/3 = 前面、下 1/3 = 后面、中 1/3 = 子级）。
 * 落在普通行（level=0）上 → 拒绝（§8：普通行不可作容器），返回 ok:false。
 */
export function moveRows(
  rows: OutlineRow[],
  fromIdx: number,
  toIdx: number,
  opts?: { asChild?: boolean },
): MoveResult {
  if (fromIdx < 0 || fromIdx >= rows.length) return { ok: false, moved: 0, clamped: false, reason: '源行下标越界' }
  if (toIdx < 0 || toIdx >= rows.length) return { ok: false, moved: 0, clamped: false, reason: '目标下标越界' }
  if (fromIdx === toIdx) return { ok: true, moved: 0, clamped: false }

  // 源行必须是标题：普通行归属于上方标题，拖它会连带带走该标题的其他数据行
  const mv = assertMovable(rows, fromIdx)
  if (!mv.ok) return { ok: false, moved: 0, clamped: false, reason: mv.reason }

  const target = rows[toIdx]
  const asChild = opts?.asChild === true

  // 落点非法：插到普通行之前/之后都不行（普通行不是容器，也不是章节）
  if (target.level === 0 && !asChild) {
    return { ok: false, moved: 0, clamped: false, reason: '不能以普通数据行为落点' }
  }

  // 不能把祖先拖进自己的子树（会形成环）
  if (asChild && isAncestorOf(rows, fromIdx, toIdx)) {
    return { ok: false, moved: 0, clamped: false, reason: '不能拖入自身子树' }
  }

  const segEnd = spanEnd(rows, fromIdx)
  const moved = segEnd - fromIdx

  // 落点若落在被移动的段内 = 无意义操作，提前拒绝（下标还有意义时判）
  if (toIdx >= fromIdx && toIdx < segEnd) {
    return { ok: false, moved: 0, clamped: false, reason: '目标落在被移动的段内' }
  }

  // ★ 顺序铁律：必须**先摘出、再改 level**（2026-10-04 实测踩坑）：
  //   takeSubtree 内部靠 spanEnd 按 level 判边界，若先改了 level，
  //   段边界就按新 level 重算，摘出的段会错（asChild 场景必现）。
  const seg = takeSubtree(rows, fromIdx)
  if (seg.length === 0) return { ok: false, moved: 0, clamped: false, reason: '没有可移动的行' }

  // 变成子级 → 整段降一级（相对新父亲）
  if (asChild) {
    for (const r of seg) r.level = Math.min(MAX_LEVEL, r.level + 1) as RowLevel
  }

  // ★ 落点在**摘出之后**算：摘出后 toIdx 可能前移（fromIdx < toIdx 时要减 moved）。
  const tgt = fromIdx < toIdx ? toIdx - moved : toIdx
  const at = asChild ? spanEnd(rows, tgt) : tgt

  // ⚠️ 这里**故意不对 seg 单独 normalizeLevels**（2026-10-04 实测踩坑）：
  //   asChild 时 seg 刚整体 +1 级，它的新父亲是 rows[tgt]。若单独 normalize，
  //   prev 从 0 起算 → 单个 L2 的行 max=1 被钳回 L1，降级凭空丢失。
  //   段脱离上下文时normalize 本身就是错的判定 —— 合并后统一 normalize 才有意义。
  //   （非 asChild 时 seg 层级不变，合并后 normalize 会覆盖同样结果，故一并省掉。）
  const clampedBefore = asChild ? false : normalizeLevels(seg)

  insertSubtree(rows, seg, at)
  const clampedAll = normalizeLevels(rows)
  renumber(rows)

  return { ok: true, moved, clamped: clampedBefore || clampedAll }
}

/**
 * isAncestorOf —— rows[i] 是否为 rows[j] 的标题祖先
 *
 * 判定：从 j 往回找**最近的 level>0 的行**，看它下标是否 === i。
 * （不能简单比较 level：中间隔了同级标题就不再是祖先，
 *   例如 1 / 1.1 / 2中，'1' 不是 '2' 的祖先，且 '1.1' 也不是。）
 * level=0 的行不参与祖先链——普通行没有子标题。
 */
/**
 * isAncestorOf —— rows[i] 是否为 rows[j] 的祖先（用于防成环）
 *
 * ★ 判定用「段落区间」而非「比较 level」——这是实测逼出来的正确做法：
 *   扁平 level 数组有**信息损失**。level 序列 [1,2,2,3] 里，
 *   下标 1 与 2 都是 level 2（1.1 与 1.2 是兄弟），
 *   下标 3 是 level 3。单看level 无法判断 3 的父是 1 还是 2 —— 两种都说得通。
 *
 *   段落区间法无歧义：i 的段是 [i, spanEnd(i))，j 落在段内即为后代。
 *   兄弟的段不重叠，所以天然判false（探针 11/11 用例全过）。
 */
export function isAncestorOf(rows: readonly OutlineRow[], i: number, j: number): boolean {
  if (i >= j) return false
  if (rows[i].level === 0) return false
  return j < spanEnd(rows, i)
}

/** 行是否可作为拖动起点（普通行不行：它属于上方标题，拖它会连带带走上方数据） */
export function assertMovable(rows: readonly OutlineRow[], idx: number): { ok: boolean; reason?: string } {
  if (idx < 0 || idx >= rows.length) return { ok: false, reason: '行下标越界' }
  if (rows[idx].level === 0) return { ok: false, reason: '普通数据行不能单独拖动（请拖它上方的标题）' }
  return { ok: true }
}

/* ────────────────── 层级调整（树上右键）────────────────── */

/**
 * shiftLevel —— 升级 / 降级
 *
 * 升级 level+1：会连同其子树全部下移一级。
 * 降级 level-1：会挤压子树，降级后越界的子孙由 normalizeLevels 钳制。
 *
 * @returns 同moveRows 的结果（clamped 供 toast 用）
 */
export function shiftLevel(rows: OutlineRow[], rowIdx: number, delta: 1 | -1): Omit<MoveResult, 'ok' | 'reason'> {
  const r = rows[rowIdx]
  if (!r) return { moved: 0, clamped: false }
  if (r.level === 0) {
    // 普通行不是章节，不能升降级（要改层级请在标记列填数字，或先把它变标题）
    return { moved: 0, clamped: false }
  }
  {
    const next = r.level + delta
    if (next < 1 || next > MAX_LEVEL) return { moved: 0, clamped: false }
    // 降级时连带子树：所有更深的后代行一起 -1，防止层级倒挂
    if (delta === -1) {
      const end = rowIdx + subtreeSpan(rows, rowIdx)
      for (let k = rowIdx; k < end; k++) {
        rows[k].level = Math.max(1, rows[k].level - 1) as RowLevel
      }
    } else {
      r.level = next as RowLevel
    }
  }
  const clamped = normalizeLevels(rows)
  renumber(rows)
  return { moved: 1, clamped }
}

/* ────────────────── 目录树派生 ────────────────── */

/**
 * buildTree —— 从 rows 派生目录树（不独立存状态）
 *
 * 普通行（level=0）挂到「最近的标题节点」上，只累加 bodyCount，不进树。
 * 用栈维护当前路径，遇到 level 更浅的节点就往回弹。
 */
export function buildTree(rows: readonly OutlineRow[]): TreeNode[] {
  const roots: TreeNode[] = []
  const stack: TreeNode[] = []
  rows.forEach((r, i) => {
    if (r.level === 0) {
      // 普通行：挂到最近标题节点
      if (stack.length) stack[stack.length - 1].bodyCount++
      return
    }
    const node: TreeNode = {
      rowId: r.id,
      level: r.level,
      no: r.no,
      title: r.title,
      collapsed: r.collapsed,
      children: [],
      rowIndex: i,
      hasChildren: false,
      bodyCount: 0,
    }
    while (stack.length && stack[stack.length - 1].level >= r.level) stack.pop()
    if (stack.length) {
      stack[stack.length - 1].children.push(node)
      stack[stack.length - 1].hasChildren = true
    } else {
      roots.push(node)
    }
    stack.push(node)
  })
  return roots
}

/** 树的标题节点总数 */
export function countTitles(tree: TreeNode[]): number {
  let n = 0
  for (const t of tree) n += 1 + countTitles(t.children)
  return n
}

/** 某节点的父rowIndex，找不到返回 -1 */
export function findParentRowIndex(rows: readonly OutlineRow[], rowIdx: number): number {
  if (rowIdx <= 0) return -1
  const L = rows[rowIdx].level
  for (let i = rowIdx - 1; i >= 0; i--) {
    if (rows[i].level > 0 && rows[i].level < L) return i
  }
  return -1
}

/* ────────────────── 折叠 ────────────────── */

/**
 * visibleRowIndexes —— 折叠后可见的行下标
 *
 * 折叠一个标题 → 隐藏它**之后的整个行段**（子标题 + 普通数据行），
 * **标题行自己保留**（否则用户看不到这个折叠项，也就无从展开）。
 * 方案 §4.3 说的「整段隐藏」指后代行段，不含标题自身。
 *
 * 用「隐藏区间栈」而非「层级栈」：因为扁平 level 数组里，
 * 同级标题的边界靠 spanEnd 才能确定（level 相同即边界），层级栈会漏判。
 */
export function visibleRowIndexes(rows: readonly OutlineRow[]): number[] {
  const out: number[] = []
  // 生效的隐藏区间栈 [from, to)，允许嵌套（外层折叠时内层标记无意义但无害）
  const hides: { from: number; to: number }[] = []
  for (let i = 0; i < rows.length; i++) {
    while (hides.length && hides[hides.length - 1].to <= i) hides.pop()
    if (hides.length > 0) continue // 在某个隐藏区间内
    out.push(i)
    if (rows[i].level > 0 && rows[i].collapsed) {
      const end = spanEnd(rows, i)
      if (end > i + 1) hides.push({ from: i + 1, to: end })
    }
  }
  return out
}

/** 折叠状态取反（只改 collapsed，不需重算编号 —— 方案 §5） */
export function toggleCollapse(rows: OutlineRow[], rowIdx: number): boolean {
  const r = rows[rowIdx]
  if (!r || r.level === 0) return false
  r.collapsed = !r.collapsed
  return r.collapsed
}

/** 折叠全部 / 展开全部（只改 collapsed） */
export function setCollapseAll(rows: OutlineRow[], collapsed: boolean): void {
  for (const r of rows) if (r.level > 0) r.collapsed = collapsed
}

/* ────────────────── 扁平化（渲染用）────────────────── */

/**
 * flattenTree —— 深度优先展开，折叠节点不输出其后代
 *
 * 返回的每项带 depth（缩进用）与 span（行段长度，表格侧定位用）。
 */
export function flattenTree(tree: TreeNode[], rows: readonly OutlineRow[]): FlatNode[] {
  const out: FlatNode[] = []
  const walk = (nodes: TreeNode[], depth: number) => {
    for (const n of nodes) {
      const span = subtreeSpan(rows, n.rowIndex)
      out.push({ ...n, depth, span })
      if (n.collapsed) continue
      walk(n.children, depth + 1)
    }
  }
  walk(tree, 0)
  return out
}

/* ────────────────── 唯一收口 ────────────────── */

export interface CommitHooks {
  /** 树重建回调（UI 层把新树塞进 state） */
  rebuild?: (tree: TreeNode[]) => void
  /** 落盘标记 */
  markDirty?: () => void
  /** toast 提示 */
  toast?: (msg: string) => void
}

export interface CommitResult {
  /** 是否发生过跳级修正（→ 已自动 toast 提示） */
  clamped: boolean
  /** 重建出的树 */
  tree: TreeNode[]
}

/**
 * commitStructure —— 所有结构变更的唯一入口（方案 §5）
 *
 * 钳制 → 重算编号 → 重建树 → 落盘标记。**顺序不可换**：
 * 先钳制再编号，否则脏层级会算出脏编号。
 *
 * 用户动作与调用点：
 *   拖表格行 / 拖树节点 → moveRows 后调本函数（或直接用本函数包住）
 *   树上右键升降级→ shiftLevel 后调
 *   编辑标记列 → parseLevel 后调
 *   打开 xlsx   → 解析 A 列后调
 *   折叠/展开   → **不调**（只改 collapsed，不重算编号）
 */
export function commitStructure(rows: OutlineRow[], hooks: CommitHooks = {}): CommitResult {
  const clamped = normalizeLevels(rows)
  renumber(rows)
  const tree = buildTree(rows)
  hooks.rebuild?.(tree)
  hooks.markDirty?.()
  if (clamped) hooks.toast?.('已自动修正跳级标题')
  return { clamped, tree }
}

/* ────────────────── 工厂 ────────────────── */

let seq = 0
/** 生成一个稳定 id（不用随机数，便于测试复现） */
export function nextRowId(prefix = 'r'): string {
  seq += 1
  return `${prefix}${seq}`
}

/** 造一行 */
export function makeRow(partial: Partial<OutlineRow> = {}): OutlineRow {
  return {
    id: partial.id ?? nextRowId(),
    level: partial.level ?? 0,
    title: partial.title ?? '',
    no: partial.no ?? '',
    collapsed: partial.collapsed ?? false,
    cells: partial.cells ?? {},
  }
}

/**
 * 造一组带层级的行（测试与示例用）
 *
 * @param spec 用**点号深度**表达层级：'1' = 一级，'1.1' = 二级，'0'/null = 普通数据行。
 *   ⚠️ 点号只表达**深度**，不表达编号 —— '1.1' 与 '1.2' 都造出 level 2（兄弟）。
 *   这个写法在「兄弟 vs 父子」需要区分时**有先天歧义**，那种场景请改用 makeRowsByLevel。
 */
export function makeRows(spec: (string | null | undefined)[]): OutlineRow[] {
  return spec.map((s) => {
    if (!s || s === '0') return makeRow({ level: 0 })
    const lv = s.split('.').length
    return makeRow({ level: Math.min(MAX_LEVEL, lv) as RowLevel, title: `标题${s}` })
  })
}

/**
 * 按 level 数组造行 —— **测试里表达层级的首选**
 *
 * 点号串无法区分兄弟与父子（'1.1' 和 '1.2' 都是深度 2），
 * 而 level 数组是 level 字段的真实形状，无歧义。
 *
 * 例：makeRowsByLevel([1, 2, 2, 3]) = 1 / 1.1 / 1.2 / 1.2.1
 */
export function makeRowsByLevel(levels: (number | null)[], titles?: string[]): OutlineRow[] {
  return levels.map((lv, i) => {
    if (lv == null || lv === 0) return makeRow({ level: 0, title: titles?.[i] ?? `数据${i + 1}` })
    const l = Math.max(1, Math.min(MAX_LEVEL, lv)) as RowLevel
    return makeRow({ level: l, title: titles?.[i] ?? `L${l}-${i + 1}` })
  })
}
