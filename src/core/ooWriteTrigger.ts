/**
 * OO 域外链接 · 自动写回触发判定
 *
 * 立档：2026-10-06
 *
 * 为什么不用 trigger.ts 的 planTrigger：
 *   它产出的是 Excel 语义的计划（WriteOp / recomputeSheets / needCommit），
 *   而 OO 场景要的是「**哪些 ooRefs 的上游变了**」——判据不同、产物不同。
 *   强行复用会把 Excel 的概念漏进 OO 侧。
 *   ⇒ 单独做，但**值变化判定**复用 trigger.ts 的 valueChanged（同一口径）。
 *
 * 纯计算层：零 IO / 零 electron / 零 document。
 */

import type { OoRef } from './ooRef'
import { valueChanged } from '../xlsx/trigger'

/** 某个节点的上游快照（用于比变更前后） */
export type Snapshot = Record<string, Record<string, unknown>>

/** 一个引用解析出来的当前值 */
export interface RefValue {
  urn: string
  value: unknown
}

/**
 * 解析一个 ooRef 的当前值。
 * @param readNode 取某节点的字段值（由调用方从 computeNode / records 提供）
 */
export function readRefValue (
  ref: OoRef,
  urn: string,
  readNode: (nodeId: string, fieldId: string) => unknown,
): RefValue {
  return { urn, value: readNode(ref.nodeId, ref.fieldId) }
}

/** 一次上游变更：哪个节点的哪个字段变了 */
export interface RefChange {
  urn: string
  before: unknown
  after: unknown
}

/**
 * ★ 核心：比对两次快照，产出「变了哪些引用」。
 *
 * 语义选择（为什么不逐次报事件）：
 *   快照比对 = **状态驱动**，不是事件驱动。
 *   事件驱动的坑：上游连续改 5 次会触发 5 次写回（解包 5 次），
 *   且中途可能取到不完整的值。
 *   状态比对：只关心「现在和上次写回时不一样」⇒ 天然去抖、幂等。
 */
export function diffRefs (
  refs: OoRef[],
  urnOf: (ref: OoRef) => string,
  readNode: (nodeId: string, fieldId: string) => unknown,
  prev: Snapshot | null,
): { changes: RefChange[]; current: Snapshot } {
  const current: Snapshot = {}
  const changes: RefChange[] = []
  for (const ref of refs) {
    const urn = urnOf(ref)
    const value = readNode(ref.nodeId, ref.fieldId)
    current[urn] = { __v: value }
    if (!prev) continue                      // 首次没有基线 ⇒ 不当作变更（见下）
    const before = prev[urn]?.__v
    if (valueChanged(before, value)) {
      changes.push({ urn, before, after: value })
    }
  }
  return { changes, current }
}

/**
 * 首次该不该写回。
 *
 * ★ 这里刻意取「不写」：首次建立基线时把所有引用都当变更，
 *   会导致用户刚点完关联、文件内容瞬间被覆盖 —— 既突兀又危险
 *   （域落点可能还没被发现/对位）。基线建立后，往后的真变更才写。
 *   ⇒ 用户想要「建立关联即刷新一次内容」，由 UI 显式调用 force 写回。
 */
export function shouldWriteOnFirstBind (): boolean { return false }

/** 判定结果 */
export interface WritebackDecision {
  write: boolean
  changes: RefChange[]
  current: Snapshot
  why: string
}

/**
 * 决定要不要写回。
 * @param refs 该 OO 节点的域清单
 * @param urnOf 取 URN（通常 ooRefUrn）
 * @param readNode 取值
 * @param prev 上次写回后的快照（null = 首次）
 */
export function decideWriteback (
  refs: OoRef[],
  urnOf: (ref: OoRef) => string,
  readNode: (nodeId: string, fieldId: string) => unknown,
  prev: Snapshot | null,
): WritebackDecision {
  if (!refs.length) {
    return { write: false, changes: [], current: {}, why: '该节点没有域引用' }
  }
  const { changes, current } = diffRefs(refs, urnOf, readNode, prev)
  if (prev === null) {
    return {
      write: shouldWriteOnFirstBind(),
      changes: [],
      current,
      why: '首次建立基线，不写回（避免刚关联就覆盖文件）',
    }
  }
  if (!changes.length) {
    return { write: false, changes, current, why: '上游值未变化' }
  }
  return {
    write: true,
    changes,
    current,
    why: `${changes.length} 个引用上游已变（${changes.map((c) => c.urn).slice(0, 3).join('、')}…）`,
  }
}
