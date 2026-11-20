/**
 * OO 域外链接 · 点选引用 → 形成关联（纯函数层）
 *
 * 立档：2026-10-06 · 用户指令：
 *   「我先不做线，我做成点击后能引用参数，然后形成关联」
 *
 * 为什么是这个文件：走「连线」那条路会撞落库断点
 * （OO 节点无 fields，被 App.tsx:291 的守卫静默 return）。
 * 走「点击」这条路不需要连线，关联由本文件从 ooRefs 直接推导。
 *
 * ★ 零 IO / 零 electron / 零 document：纯计算，可单测、可在纯计算层复用。
 */

import type { NodeConfig } from './schema'

/** ooRefs 的一项（与 schema.ooRefs 同构） */
export interface OoRef {
  nodeId: string
  fieldId: string
  /** 标准引用地址 node://<nodeId>#<fieldId>；老数据可能缺失 */
  urn?: string
}

/** 拼标准 URN。与 ref/urn.ts 的解析规则严格对称（node://<id>#<locator>） */
export function ooRefUrn (nodeId: string, fieldId: string): string {
  return `node://${nodeId}#${fieldId}`
}

/**
 * 取出权威 URN：优先用已存的 urn，缺失时按协议补齐。
 * ★ 老数据（只有 nodeId/fieldId）走这里 ⇒ 不必做数据迁移。
 */
export function refUrnOf (ref: OoRef): string {
  return ref.urn && ref.urn.trim() ? ref.urn : ooRefUrn(ref.nodeId, ref.fieldId)
}

/**
 * ★ 核心：点选一个参数 ⇒ 更新域清单（去重、补 urn）。
 *
 * @param cur  当前 ooRefs（可能为 undefined）
 * @param nodeId 选中的上游节点
 * @param fieldId 选中的字段
 * @returns 新的 ooRefs（已存在则原样返回，引用对象以便调用方跳过 setState）
 */
export function addOoRef (cur: OoRef[] | undefined, nodeId: string, fieldId: string): OoRef[] {
  const list = cur ?? []
  if (!nodeId || !fieldId) return list
  if (list.some((r) => r.nodeId === nodeId && r.fieldId === fieldId)) return list
  return [...list, { nodeId, fieldId, urn: ooRefUrn(nodeId, fieldId) }]
}

/** 取消引用（面板再次点同一项时） */
export function removeOoRef (cur: OoRef[] | undefined, nodeId: string, fieldId: string): OoRef[] {
  return (cur ?? []).filter((r) => !(r.nodeId === nodeId && r.fieldId === fieldId))
}

/**
 * ★★ 关联推导：把 ooRefs 折成「数据流上游集合」。
 *
 * 这是「形成关联」的关键 —— 数据流边在 projectFormat.edgesOf 里
 * **只从 config.linkedInputs 派生**（单一真相源），
 * 所以只要把引用到的上游写进 linkedInputs，
 * 画布连线、工程文件、触发引擎**三处同时通**（一处改，三处生效）。
 *
 * 语义：一个字段被引用 ⇒ 该节点就是数据流上游。
 * 与「连线」的差别只是**建立方式**：连线是手拖，这里是点选。
 */
export function linkedInputsFromOoRefs (refs: OoRef[] | undefined): string[] {
  const out: string[] = []
  for (const r of refs ?? []) {
    if (r.nodeId && !out.includes(r.nodeId)) out.push(r.nodeId)
  }
  return out
}

/**
 * ★ 一次点选要落的库（合并成一个 patch）。
 *
 * 为什么必须合并成一次：setNodeConfig 是**浅合并**，
 * 分两次写会在中途失败时留下半截状态，且撤销栈会多出一个无意义快照。
 * （与 App.tsx 里「AI 线 + 偏好预填合并成同一个 patch」同一处理。）
 */
export interface OoRefPatch {
  ooRefs: OoRef[]
  linkedInputs: string[]
}

/** 由当前 config 与本次点选，算出要写的 patch */
export function ooRefPatchOf (
  config: NodeConfig | undefined,
  nodeId: string,
  fieldId: string,
): OoRefPatch {
  const cur = (config?.ooRefs as OoRef[] | undefined) ?? []
  const ooRefs = addOoRef(cur, nodeId, fieldId)
  // linkedInputs 要**并集**，不能只写 ooRefs 推出来的 ——
  // 手拖连线写入的上游若不在 ooRefs 里，不能被这次点选抹掉。
  const manual = config?.linkedInputs ?? []
  const derived = linkedInputsFromOoRefs(ooRefs)
  const linkedInputs = [...derived, ...manual.filter((id) => !derived.includes(id))]
  return { ooRefs, linkedInputs }
}

/**
 * 取域的当前值（写回管线用）。
 *
 * 纯函数：只做「按 URN 从表快照里取值」，不碰 IO。
 * 取不到返回 undefined —— 由调用方决定是跳过还是写空。
 */
export function ooRefValue (
  ref: OoRef,
  tables: Record<string, { fields?: { id: string; name: string }[]; rows?: unknown[] }>,
): unknown {
  const t = tables[ref.nodeId]
  if (!t) return undefined
  const row0 = Array.isArray(t.rows) ? (t.rows[0] as Record<string, unknown> | undefined) : undefined
  if (!row0) return undefined
  return row0[ref.fieldId]
}

/** 人类可读引用文本（写进 OO 单元格/正文时用，避免显示 node:// 原文） */
export function ooRefLabel (ref: OoRef, nodeName?: string, fieldName?: string): string {
  const n = nodeName ?? ref.nodeId
  return fieldName ? `${n}.${fieldName}` : n
}
