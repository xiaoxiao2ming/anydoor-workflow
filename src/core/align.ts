/**
 * 自动对齐 —— 网格归位（2026-10-06 用户：「增加一个自动对齐按钮」）。
 *
 * ★ 与 `graph.layoutLayers`（顶栏「自动排列」）是**两件不同的事，别混**：
 *
 *   | 按钮       | 函数                | 做什么                                | 相对关系 |
 *   |-----------|---------------------|---------------------------------------|---------|
 *   | 自动对齐  | 本文件 alignLayout  | 坐标吸附到网格、消除重叠、拉直同行    | **保持** |
 *   | 自动排列  | graph.layoutLayers  | 按上下游分层，左→右重排                | **改变** |
 *
 *   用户口径（2026-10-06）：「自动对齐 和 自动排列 两个按钮 一个是对齐
 *   一个是彻底重排」。所以本文件**绝不能**按拓扑分层、也**绝不重排行间距** ——
 *   一旦做了，就变成 layoutLayers 的复制品，两个按钮语义塌成一个。
 *
 * 本文件只做三件事，且每一件都有具体的「看起来能跑但会错」的坑：
 *
 * ① **网格吸附用四舍五入，不是 floor**
 *   floor 会把 x=143 拉到 140、x=147 也拉到 140 —— 误差不对称，节点一律向左跳。
 *   round取最近网格点才符合「吸附」的直觉。
 *
 * ② **重叠消解必须逐个单向右推，不能两两互推**
 *   两两互推在三个以上重叠时会震荡（a推b、b推a、a再推b…）直到超出版本上限，
 *   而且结果依赖遍历顺序 ⇒ 同样输入两次调用可能得到不同布局（**不幂等**）。
 *   单向右推是单向的，天然幂等。
 *
 * ③ **分行用 y 区间重叠判定，不是「y 相同」**
 *   节点 y 是用户手拖的，不会恰好相等。用 y 相同分行会把实际同排的节点拆到不同行。
 *
 * ④ **尺寸缺测量值时宁可高估，不可低估**
 *   Canvas 侧节点未渲染时 `measured` 为 undefined。传 0 会让所有节点挤在同一点；
 *   传保守默认（与 layoutLayers 同源）至少保证不重叠成一坨。
 *   ★ 220×90 是**未折叠常规节点**的估值；圆形节点与折叠态实际更小 ——
 *   宁可多留空隙，也不要因低估而重叠。
 *
 * ⑤ **不删除、不新增节点**
 *   返回 map 的键集合必须与传入 id 完全一致 —— 少一个键，React Flow 会把那个节点
 *   「弹回」默认位置，用户看到的是「节点随机跳动」。
 */
import type { NodePos } from './store'

/** 网格步长（px）。吸附的对齐基准就是这个值。 */
export const ALIGN_GRID = 20

/** 同行相邻两节点的最小水平留白（px）。 */
export const ALIGN_GAP_X = 40

/**
 * 未测量到尺寸时的保守默认（与 graph.layoutLayers 的 NODE_W/NODE_H 同源）。
 * 故意高估：高估 → 多留空隙（可接受）；低估 → 节点重叠（不可接受）。
 */
export const ALIGN_FALLBACK_W = 220
export const ALIGN_FALLBACK_H = 90

/** 对齐需要尺寸来算重叠，纯计算层不碰 DOM —— 尺寸由调用方喂进来。 */
export interface AlignSize {
  w: number
  h: number
}

/** 吸附到最近网格点（四舍五入，非 floor —— 见①）。非有限数归0，防NaN 污染全表。 */
export function snapToGrid(v: number, grid: number = ALIGN_GRID): number {
  if (!Number.isFinite(v)) return 0
  if (!(grid > 0)) return v // 网格非法时原样返回，不静默改变语义
  return Math.round(v / grid) * grid
}

/**
 * 把节点按 y 区间重叠分行，返回每行的 id 列表（行按 y 升序，行内按 x 升序）。
 *
 * 分组判据见③：y 区间有交集即视为同一行。
 * 行内按 x 升序 —— 这是「单向右推」能成立的前提。
 */
export function groupIntoRows(
  ids: readonly string[],
  pos: Record<string, NodePos>,
  size: Record<string, AlignSize> = {},
): string[][] {
  interface Row {
    ids: string[]
    minY: number
    maxY: number
  }
  const rows: Row[] = []
  // 按 y 升序遍历：倒序找可挂靠的行（最下面的优先），避免新建多余行
  for (const id of [...ids].sort((a, b) => (pos[a]?.y ?? 0) - (pos[b]?.y ?? 0))) {
    const p = pos[id] ?? { x: 0, y: 0 }
    const h = size[id]?.h ?? ALIGN_FALLBACK_H
    const y0 = p.y
    const y1 = p.y + h
    let hit: Row | undefined
    for (let i = rows.length - 1; i >= 0; i--) {
      // 区间相交（用 < 而非 <=：恰好相接不算同一行）
      if (y0 < rows[i].maxY && y1 > rows[i].minY) {
        hit = rows[i]
        break
      }
    }
    if (hit) {
      hit.ids.push(id)
      hit.minY = Math.min(hit.minY, y0)
      hit.maxY = Math.max(hit.maxY, y1)
    } else {
      rows.push({ ids: [id], minY: y0, maxY: y1 })
    }
  }
  return rows
    .sort((a, b) => a.minY - b.minY)
    .map((r) => r.ids.sort((a, b) => (pos[a]?.x ?? 0) - (pos[b]?.x ?? 0)))
}

/**
 * 自动对齐（网格归位）主函数。纯计算，零 IO。
 *
 * @param ids   参与对齐的节点 id（Canvas 侧传全部未删除节点）
 * @param pos   当前位置表
 * @param size  各节点尺寸；缺失走 ALIGN_FALLBACK_*
 * @returns     新位置表。**键集合与传入 id 完全一致**（见⑤）
 *
 * 三步：① 各自吸附到网格 → ② 按 y 重叠分行 → ③ 逐行从左到右单向右推。
 * ★ 不动行间距、不分层、不改x/y 的相对次序 —— 那是「自动排列」干的事。
 */
export function alignLayout(
  ids: readonly string[],
  pos: Record<string, NodePos>,
  size: Record<string, AlignSize> = {},
): Record<string, NodePos> {
  const out: Record<string, NodePos> = {}
  if (!ids.length) return out

  // ① 先吸附，再用**吸附后**的坐标判定同行（比原始坐标稳定）
  const snapped: Record<string, NodePos> = {}
  for (const id of ids) {
    const p = pos[id] ?? { x: 0, y: 0 }
    snapped[id] = { x: snapToGrid(p.x), y: snapToGrid(p.y) }
  }

  // ③ 逐行：y 统一取行内第一个（x 最小）节点的 y = 上边缘拉齐，
  //    cursor 单向右推保证同行相邻间距≥ GAP_X。
  // ★★ 每一步都必须再吸附一次（见下方「为什么宽高不是网格倍数也不影响」）：
  //   节点实测宽是 182/283/202 这类**非网格倍数**，若只吸附起点，
  //   递推 `x += w + GAP_X` 会把 w mod grid 的余数一路累积下去，
  //   走到第 5~6 个节点就偏出网格（实测 182 宽时第 3 个起就到 262/282）。
  for (const row of groupIntoRows(ids, snapped, size)) {
    const y = snapped[row[0]].y
    let cursor = snapped[row[0]].x
    for (const id of row) {
      // ★ 吸附后再放位：cursor 永远停在网格上（起点已吸附、每步再吸附）
      out[id] = { x: snapToGrid(cursor), y }
      cursor += (size[id]?.w ?? ALIGN_FALLBACK_W) + ALIGN_GAP_X
    }
  }

  return out
}