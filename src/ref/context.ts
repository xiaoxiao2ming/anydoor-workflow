/**
 * 引用上下文构建 —— 从「本节点的上游」造出 RefContext + 字段归属反查。
 *
 * 为什么必须共用一份：
 *   上一版是两边各写一份 ——
 *     xlsx/buildXlsxData.ts 里有 ownerOf / columnOf / ctx 构造（走 URN）
 *     docx/buildData.ts 只有 fieldId 直取（没有 URN，因此没有溯源）
 *   结果是一边能答「这个数从哪来」、一边不能。
 *   共用一份才是真收敛：Word 与 Excel 只是同一套引用的两个消费者。
 */

import type { FieldDef, TableSchema } from '../core/schema'
import type { Row } from '../core/store'
import { computeNode } from '../core/engine'
import type { RefContext } from './urn'

export type Agg = 'first' | 'sum' | 'count' | 'max' | 'min' | 'avg'

export interface UpstreamResult {
  id: string
  name: string
  fields: FieldDef[]
  rows: Row[]
}

export interface UpstreamBuild {
  /** 上游节点 id：连线优先，其后是面板手动勾选（去重） */
  ids: string[]
  /** 上游字段并集（保序去重），供「按名字自动匹配」用 */
  fields: FieldDef[]
  /** 每个上游算好的结果 */
  computed: UpstreamResult[]
  /** 引用上下文 —— 喂给 resolveUrn() */
  ctx: RefContext
  /** fieldId → 它属于哪个上游节点（找不到返回 null；宁可缺，也不编造归属） */
  ownerOf(fieldId: string): { id: string; name: string } | null
  /** fieldId → 该字段所属上游的整行集合 */
  rowsOf(fieldId: string): Row[]
  /** fieldId → URN `node://<上游id>#<字段名>`；归属不明时返回 null */
  urnOf(fieldId: string): string | null
}

/** 本节点声明的上游：连线（linkedInputs）优先，其后是面板勾选的 inputs，去重 */
export function upstreamIdsOf(t: TableSchema): string[] {
  const linked = t.config?.linkedInputs ?? []
  const rest = (t.config?.inputs ?? []).filter((id) => !linked.includes(id))
  return [...linked, ...rest]
}

/** 去掉浮点误差尾巴：1286.5000000000002 → 1286.5 */
export function round6(n: number): number {
  return Math.round(n * 1e6) / 1e6
}

/**
 * 聚合一组行里某个字段的值。
 *
 * 注意 `first` 分支把 undefined 归一成 `''` 而不是原样返回：
 * 上游丢下来的 undefined 会被绑定层当成「未提供」而跳过该域，
 * 结果是占位符原样残留进交付稿 —— 那比填个空更难发现。
 */
export function aggregate(rows: Row[], fieldId: string, agg: Agg | undefined): unknown {
  const nums = rows
    .map((r) => r[fieldId])
    .filter((v): v is number => typeof v === 'number' && Number.isFinite(v))
  switch (agg) {
    case 'sum':
      return nums.length ? round6(nums.reduce((a, b) => a + b, 0)) : ''
    case 'count':
      return rows.length
    case 'max':
      return nums.length ? round6(Math.max(...nums)) : ''
    case 'min':
      return nums.length ? round6(Math.min(...nums)) : ''
    case 'avg':
      return nums.length ? round6(nums.reduce((a, b) => a + b, 0) / nums.length) : ''
    case 'first':
    default:
      return rows.length ? (rows[0][fieldId] ?? '') : ''
  }
}

/**
 * 算出本节点的上游上下文。
 *
 * 只收集**真实存在**的上游（tables 里没有的 id 直接忽略）——
 * 画布上删掉一个节点后，配置里可能还留着悬空引用。
 */
export function buildUpstream(
  node: TableSchema,
  tables: Record<string, TableSchema>,
  records: Record<string, Row[]>,
): UpstreamBuild {
  const ids = upstreamIdsOf(node).filter((id) => tables[id])
  const computed: UpstreamResult[] = ids.map((id) => {
    const { fields, rows } = computeNode(id, tables, records)
    return { id, name: tables[id]?.name ?? id, fields, rows }
  })

  const fields: FieldDef[] = []
  const seen = new Set<string>()
  for (const c of computed) {
    for (const f of c.fields) {
      if (!seen.has(f.id)) {
        seen.add(f.id)
        fields.push(f)
      }
    }
  }

  // 引用上下文：每个上游一份「字段名/id → 首行值」+ 一份 rows。
  // 同时给 id 和 name 两个键：用户写 URN 时可能用哪个都不确定，
  // 两个都认比猜一个友好。
  const nodes: Record<string, unknown> = {}
  for (const c of computed) {
    const scalars: Record<string, unknown> = {}
    for (const f of c.fields) {
      const v = aggregate(c.rows, f.id, 'first')
      scalars[f.id] = v
      scalars[f.name] = v
    }
    nodes[c.id] = { ...scalars, rows: c.rows, fields: c.fields }
  }

  const ownerOf = (fieldId: string) => {
    for (const c of computed) {
      const f = c.fields.find((x) => x.id === fieldId)
      if (f) return { id: c.id, name: f.name }
    }
    return null
  }

  const rowsOf = (fieldId: string) => {
    for (const c of computed) {
      if (c.fields.some((f) => f.id === fieldId)) return c.rows
    }
    return []
  }

  const urnOf = (fieldId: string) => {
    const o = ownerOf(fieldId)
    // #<字段名> 而不是 #<字段id>：URN 是给人看、给人写的，
    // 字段名在界面上和模板里都是稳定的（id 是内部生成的流水号）。
    return o ? `node://${o.id}#${o.name}` : null
  }

  return { ids, fields, computed, ctx: { nodes }, ownerOf, rowsOf, urnOf }
}

/**
 * 把一条绑定解析成「用哪个字段 / 哪个 URN」。
 *
 * 优先级：`source`（URN，多源）> `fieldId`（单源）> 名字自动匹配。
 * 返回的 `key` 是要传给 aggregate / rowsOf 的字段 id；拿不到就是 null。
 */
export function resolveBindingTarget(
  binding: { fieldId?: string; source?: string } | undefined,
  slotKey: string,
  ub: UpstreamBuild,
): { fieldId: string | null; urn: string | null; by: 'source' | 'fieldId' | 'auto' | 'none' } {
  // ① 显式 URN：先看它指向本节点的哪个上游字段（好让聚合/整表取值仍能工作）
  if (binding?.source) {
    const m = /^node:\/\/([^#]+)(?:#(.*))?$/.exec(binding.source)
    if (m) {
      const [, nodeId, locator] = m
      const up = ub.computed.find((c) => c.id === nodeId)
      const name = locator || ''
      const f = up?.fields.find((x) => x.name === name || x.id === name)
      return { fieldId: f?.id ?? null, urn: binding.source, by: 'source' }
    }
    // 非 node:// 的 URN（file / api / param / const）：没有本地字段可指，
    // 值只能靠 resolveUrn 去取，交给调用方处理
    return { fieldId: null, urn: binding.source, by: 'source' }
  }

  // ② 单源字段 id
  if (binding?.fieldId) {
    return { fieldId: binding.fieldId, urn: ub.urnOf(binding.fieldId), by: 'fieldId' }
  }

  // ③ 按名字自动匹配（域 key 恰好等于某个上游字段的 id 或名字）
  const f = ub.fields.find((x) => x.id === slotKey || x.name === slotKey)
  if (f) return { fieldId: f.id, urn: ub.urnOf(f.id), by: 'auto' }

  return { fieldId: null, urn: null, by: 'none' }
}
