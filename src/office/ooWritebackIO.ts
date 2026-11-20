/**
 * OO 域外链接 · 写回管线的 IO 层
 *
 * 立档：2026-10-06 · 配套 src/core/ooWriteback.ts（纯计算层）
 *
 * 职责：读模板字节 → 调纯计算层写回 → 存回模板库。
 * ★ 分层理由：把「字节怎么改」留在纯函数里（可单测、浏览器与 Node 同构），
 *   本文件只管「字节从哪来、到哪去」。
 *
 * ⚠ 铁律（见 xlsx/patchSave.ts 注释的实测记录）：
 *   **补丁不重建**。绝不用「load → 自有模型 → 全量写回」，
 *   那会静默丢掉真实文件里的图表/绘图/透视表。
 */

import { writebackOo, type OoRefSite, type OoWritebackResult } from '../core/ooWriteback'
import type { OoRef } from '../core/ooRef'
import { getTemplateStore } from '../docx/templateStore'

export interface WritebackOptions {
  /** OO 节点 id */
  nodeId: string
  /** 绑定的模板 id（config.ooTemplateId） */
  templateId?: string
  /** oosheet / ooword */
  kind: 'oosheet' | 'ooword'
  /** 域落点 */
  sites: OoRefSite[]
  /** 取值 */
  pick: (ref: OoRef) => unknown
  /** Word：可读标签（sdtTag 缺失时用它匹配） */
  labelOf?: (ref: OoRef) => string
}

export interface WritebackOutcome {
  ok: boolean
  /** 没写成功时的原因（缺模板 / 读失败 / 无落点…） */
  why?: string
  /** 实际改了几个域 */
  written?: number
  labels?: string[]
  skipped?: { fieldId: string; why: string }[]
  restored?: string[]
}

/**
 * 写回一个 OO 节点：读模板 → 补丁改 → 存回模板库。
 *
 * ★ 只在**真的有域**时才动作：
 *   没有 ooRefs 的 OO 节点不该被触发（否则每次上游变都白跑一次解包）。
 *   调用方负责先判空；这里再兜一层，避免误用。
 */
export async function writebackNode (opt: WritebackOptions): Promise<WritebackOutcome> {
  const { nodeId, templateId, kind, sites } = opt

  if (!templateId) {
    return { ok: false, why: '节点未绑定模板（config.ooTemplateId 为空）', written: 0 }
  }
  if (!sites.length) {
    return { ok: false, why: '没有域落点（sites 为空）', written: 0 }
  }

  const store = getTemplateStore()
  let bytes: Uint8Array | null
  try {
    bytes = await store.read(templateId)
  } catch (e) {
    return { ok: false, why: '读取模板失败：' + String((e as Error)?.message ?? e), written: 0 }
  }
  if (!bytes || !bytes.length) {
    return { ok: false, why: '模板内容为空（可能被清空或 id 失效）', written: 0 }
  }

  let res: OoWritebackResult
  try {
    res = await writebackOo({
      bytes,
      kind,
      sites,
      pick: opt.pick,
      labelOf: opt.labelOf,
    })
  } catch (e) {
    // ★写回失败**绝不覆盖原模板** —— 宁可这次不更新，也不能存进去一个坏文件。
    return { ok: false, why: '写回失败（已保留原模板）：' + String((e as Error)?.message ?? e), written: 0 }
  }

  // 一个域都没写成 ⇒ 不落库（避免把「无变化」当成新版本存进去）
  if (res.written === 0) {
    return {
      ok: false,
      why: '没有可写的域（落点未命中）',
      written: 0,
      skipped: res.skipped,
    }
  }

  try {
    // ★★ 必须用**原模板名**保存，不能另起名字。
    //   模板库的 save 语义是「按 name 去重、同名覆盖并保留 id」
    //   （idbStore.save 与 electron main.ts tpl:save 两侧一致）。
    //   ⇒ 用新名字会产生一份**副本**，而节点仍绑着旧 ooTemplateId
    //      ⇒ 写回「成功」了但节点读的还是旧字节 —— 静默失效，最难发现。
    //   同名覆盖则 id 不变 ⇒ 节点绑定自动指向新内容。
    const name = await resolveTemplateName(store, templateId, nodeId, kind)
    await store.save(name, res.bytes)
  } catch (e) {
    return {
      ok: false,
      why: '存回模板库失败：' + String((e as Error)?.message ?? e),
      written: res.written,
    }
  }

  return {
    ok: true,
    written: res.written,
    labels: res.labels,
    skipped: res.skipped,
    restored: res.restored,
  }
}

/**
 * 取模板的**原名**（覆盖保存用）。
 * 查不到时回落一个按 nodeId 拼的稳定名字 ——
 * ★ 必须「稳定」：同一个节点每次都落到同一个名字，才谈得上覆盖。
 */
async function resolveTemplateName (
  store: { list(): Promise<{ id: string; name: string }[]> },
  templateId: string,
  nodeId: string,
  kind: 'oosheet' | 'ooword',
): Promise<string> {
  try {
    const list = await store.list()
    const hit = list.find((t) => t.id === templateId)
    if (hit?.name) return hit.name
  } catch { /* 列不出来就走回落 */ }
  return `域外回写-${nodeId}.${kind === 'oosheet' ? 'xlsx' : 'docx'}`
}
