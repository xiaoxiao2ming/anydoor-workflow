// 图布局与执行顺序计算：根据上下游（数据流）关系做分层排列，并给出执行顺序。
import type { TableSchema } from './schema'
import type { NodePos } from './store'

const NODE_W = 220
const NODE_H = 90
const GAP_X = 130
const GAP_Y = 46

/** 收集某节点的直接上游（来自 linkedInputs + inputs，去重） */
function upstreamOf(t: TableSchema, tables: Record<string, TableSchema>): string[] {
  const ins = [...new Set([...(t.config?.linkedInputs ?? []), ...(t.config?.inputs ?? [])])]
  return ins.filter((id) => id !== t.id && tables[id])
}

/**
 * 分层布局（左 → 右）：x = 层号 * 列宽，y = 同层内的序号。
 * 层号 = 最长上游路径 + 1（源头为 0 层）。可处理环（环上节点降级为 0 层）。
 */
export function layoutLayers(tables: Record<string, TableSchema>): Record<string, NodePos> {
  const ids = Object.keys(tables)
  const up: Record<string, string[]> = {}
  ids.forEach((id) => (up[id] = []))
  for (const t of Object.values(tables)) {
    up[t.id] = upstreamOf(t, tables)
  }

  const layer: Record<string, number> = {}
  const visiting = new Set<string>()
  const getLayer = (id: string): number => {
    if (layer[id] !== undefined) return layer[id]
    if (visiting.has(id)) return 0 // 环保护
    visiting.add(id)
    let lv = 0
    for (const p of up[id] ?? []) lv = Math.max(lv, getLayer(p) + 1)
    visiting.delete(id)
    layer[id] = lv
    return lv
  }
  ids.forEach(getLayer)

  const byLayer: Record<number, string[]> = {}
  for (const id of ids) {
    const l = layer[id] ?? 0
    ;(byLayer[l] ||= []).push(id)
  }

  const pos: Record<string, NodePos> = {}
  Object.keys(byLayer)
    .map(Number)
    .sort((a, b) => a - b)
    .forEach((l) => {
      byLayer[l].forEach((id, i) => {
        pos[id] = { x: 40 + l * (NODE_W + GAP_X), y: 40 + i * (NODE_H + GAP_Y) }
      })
    })
  return pos
}

/** 执行顺序：按层从左到右、同层从上到下。用于「手动执行工作流」的可视化高亮。 */
export function executionOrder(tables: Record<string, TableSchema>): string[] {
  const pos = layoutLayers(tables)
  return Object.keys(pos).sort((a, b) => pos[a].x - pos[b].x || pos[a].y - pos[b].y)
}
