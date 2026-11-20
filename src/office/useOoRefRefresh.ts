/**
 * OO 域外写回 · UI 接线层
 *
 * 立档：2026-10-06 · 把 src/office/ooWritebackIO.ts 与 UI 接起来
 *
 * 职责：把四步串成一次「刷新域值」动作
 *   ① 解析落点（ooRefs × 文件里的域清单 ⇒ a1 / sdtTag）
 *   ② 取当前值（computeNode 首行）
 *   ③ 补丁式写回字节（core/ooWriteback）
 *   ④ 存回模板库
 *
 * ★ 为什么不把这些塞进 NodeEditor.tsx：
 *   那个文件已 1000+ 行且在并行开发中；接线逻辑单独成文件可单测、不易冲突。
 */

import { useCallback, useRef, useState } from 'react'
import type { TableSchema, NodeConfig } from '../core/schema'
import type { Row } from '../core/store'
import { computeNode } from '../core/engine'
import { resolveOoRefSites, describeSites, locatorOf } from '../core/ooRefSites'
import { buildSheetPartMap } from '../core/ooxmlSheetMap'
import { toWordSites } from '../core/docxSdtDiscover'
import type { OoRef } from '../core/ooRef'
import type { OoRefSite } from '../core/ooWriteback'
import { writebackNode } from './ooWritebackIO'

export interface RefreshResult {
  ok: boolean
  /** 写了几处 */
  written: number
  /** 没写成的（域在文件里找不到落点等） */
  skipped: { fieldId: string; why: string }[]
  /** 落点清单（给用户看"写在哪"） */
  sites: string[]
  /** 失败原因 */
  why?: string
}

export interface UseOoRefRefreshOpts {
  nodeId: string
  table: TableSchema | undefined
  tables: Record<string, TableSchema>
  records: Record<string, Row[]>
  /** 读当前模板字节（发现域清单用） */
  readTemplate: (templateId: string) => Promise<Uint8Array | null>
  /** 域清单发现：xlsx 走命名区域；Word 走 sdt。实现方在NodeEditor 里给。 */
  discoverXlsx?: (bytes: Uint8Array) => Promise<{ key: string; ref: string; sheet?: string }[]>
  discoverDocx?: (bytes: Uint8Array) => Promise<{ tag: string; key?: string; part?: string }[]>
  /** 汇报状态（诊断抽屉/提示条） */
  report?: (msg: string) => void
}

/**
 * ★ fieldId → 字段名。
 * 落点解析与可读标签都要用它，两处写两份必然漂移 ⇒ 抽成单一函数。
 * 遍历所有节点找同名字段 id（字段 id 在本项目是全局唯一的，如 'f1'/'qty'）。
 */
function fieldNameOf (tables: Record<string, TableSchema>): (fieldId: string) => string {
  const idx = new Map<string, string>()
  for (const tb of Object.values(tables)) {
    if (!tb || tb.deleted) continue
    for (const f of tb.fields || []) if (!idx.has(f.id)) idx.set(f.id, f.name)
  }
  return (fieldId) => idx.get(fieldId) ?? ''
}

export function useOoRefRefresh (opt: UseOoRefRefreshOpts) {  const [busy, setBusy] = useState(false)
  const [last, setLast] = useState<RefreshResult | null>(null)
  // 快照基线：上次写回后的值，供「自动触发」比对（状态驱动，非事件驱动）
  const snapRef = useRef<Record<string, unknown> | null>(null)

  const kind: 'oosheet' | 'ooword' = opt.table?.kind === 'ooword' ? 'ooword' : 'oosheet'

  /** 解析当前全部引用的落点 */
  const resolveSites = useCallback(async (): Promise<{
    sites: OoRefSite[]
    unmatched: { fieldId: string; why: string }[]
  }> => {
    const cfg = (opt.table?.config ?? {}) as NodeConfig
    const refs = (cfg.ooRefs ?? []) as OoRef[]
    if (!refs.length) return { sites: [], unmatched: [] }
    const tplId = cfg.ooTemplateId
    if (!tplId) return { sites: [], unmatched: [] }
    const bytes = await opt.readTemplate(tplId)
    if (!bytes || !bytes.length) return { sites: [], unmatched: [] }

    if (kind === 'oosheet') {
      if (!opt.discoverXlsx) {
        return {
          sites: [],
          unmatched: refs.map((r) => ({ fieldId: r.fieldId, why: '未接入 xlsx 域发现' })),
        }
      }
      const fields = await opt.discoverXlsx(bytes)
      // ★ nameOf：fieldId → 字段名。XlsxFieldSpec.key 是**命名区域的名字**，
      //   与 fieldId 不是同一套标识 ⇒ 必须经它转换才对得上。
      // ★ part（工作表部件路径）必须来自 **workbook.xml + rels 两跳解析**，
      //   不能按「第 N 张表」猜 —— 实测真实文件里 rId3 指向 sheet1，
      //   rId 与 sheetN 编号毫无关系，猜错就**写到错误的表**。
      //   解析不到时传空对象：resolveOoRefSites 会退回序号兜底，
      //   但主通道已由 writeback 层严格校验 part，不存在双前缀。
      const partMap = await buildSheetPartMap(bytes)
      if (!Object.keys(partMap).length) {
        opt.report?.('[oo] 未能解析工作表映射（缺 workbook.xml 或 rels），落点可能不准')
      }
      return resolveOoRefSites(refs, fields as never, kind, {
        nameOf: fieldNameOf(opt.tables),
        sheetPartMap: partMap,
      })
    }

    // ooword
    if (!opt.discoverDocx) {
      return {
        sites: [],
        unmatched: refs.map((r) => ({ fieldId: r.fieldId, why: '未接入 docx 域发现' })),
      }
    }
    // ★ 从真实 docx 里列出内容控件（2026-10-29 补）——
    //   实测真实 tag 是 field:工程名称 / each:材料表 这种 kind:名称 形态，
    //   toWordSites 会把前缀剥掉作为匹配 key（tag 保持原样用于写回定位）。
    //   opt.discoverDocx 约定返回**已对齐好**的 wordSites（{tag,key?,part?}），
    //   所以这里直接用，不再二次转换。
    const wordSites = await opt.discoverDocx(bytes)
    if (!wordSites.length) {
      opt.report?.('[oo] 文档里没有内容控件（w:sdt）—— Word 侧域需要模板里有占位控件')
    }
    return resolveOoRefSites(refs, [], kind, {
      wordSites, nameOf: fieldNameOf(opt.tables),
    })
  }, [opt, kind])

  /**
   * ★ 值读取：按引用取上游节点该字段的**首行**值。
   * 与 ooRefItems 的取法一致（computeNode 首行），避免两处口径不同。
   */
  const pickValue = useCallback((ref: OoRef): unknown => {
    try {
      const comp = computeNode(ref.nodeId, opt.tables, opt.records)
      const r0 = comp.rows[0] as Record<string, unknown> | undefined
      return r0?.[ref.fieldId]
    } catch {
      return undefined
    }
  }, [opt.tables, opt.records])

  /** 手动「立即刷新域值」 */
  const refresh = useCallback(async (): Promise<RefreshResult> => {
    if (busy) return { ok: false, written: 0, skipped: [], sites: [], why: '正在刷新' }
    setBusy(true)
    try {
      const cfg = (opt.table?.config ?? {}) as NodeConfig
      const r = await resolveSites()
      if (!r.sites.length) {
        const res: RefreshResult = {
          ok: false, written: 0, skipped: r.unmatched,
          sites: [], why: r.unmatched[0]?.why || '没有可写的域落点',
        }
        setLast(res)
        opt.report?.('[oo] 写回未执行：' + res.why)
        return res
      }
      const out = await writebackNode({
        nodeId: opt.nodeId,
        templateId: cfg.ooTemplateId,
        kind,
        sites: r.sites,
        pick: pickValue,
        labelOf: (ref) => {
          const nm = fieldNameOf(opt.tables)(ref.fieldId)
          if (!nm) return locatorOf(ref.urn ?? '')
          // 找该字段所属的节点名，凑成「表甲.数量」
          for (const tb of Object.values(opt.tables)) {
            if (!tb || tb.deleted) continue
            if ((tb.fields || []).some((x) => x.id === ref.fieldId)) return `${tb.name}.${nm}`
          }
          return nm
        },
      })
      const res: RefreshResult = {
        ok: out.ok,
        written: out.written ?? 0,
        skipped: [...r.unmatched, ...(out.skipped ?? [])],
        sites: describeSites(r.sites),
        why: out.why,
      }
      setLast(res)
      // 写回成功后更新基线，供后续自动触发比对
      if (out.ok) {
        const snap: Record<string, unknown> = {}
        for (const s of r.sites) snap[locatorOf(s.ref.urn ?? '')] = pickValue(s.ref)
        snapRef.current = snap
      }
      opt.report?.(
        out.ok
          ? `[oo] 已写回 ${out.written} 处（${describeSites(r.sites).slice(0, 3).join('、')}）`
          : '[oo] 写回未成功：' + (out.why ?? '未知原因'),
      )
      return res
    } catch (e) {
      const res: RefreshResult = {
        ok: false, written: 0, skipped: [], sites: [],
        why: '异常：' + String((e as Error)?.message ?? e),
      }
      setLast(res)
      opt.report?.('[oo] ' + res.why)
      return res
    } finally {
      setBusy(false)
    }
  }, [busy, opt, kind, resolveSites, pickValue])

  /**
   * 自动触发：上游值变了就写回。
   * ★ 状态驱动（非事件驱动）—— 连续改 5 次只解包 1 次，且不会取到中间态。
   * @param force true = 忽略基线强制写一次（「建立关联即刷新」用）
   */
  const autoRefresh = useCallback(async (): Promise<RefreshResult | null> => {
    const cfg = (opt.table?.config ?? {}) as NodeConfig
    const refs = (cfg.ooRefs ?? []) as OoRef[]
    if (!refs.length) return null
    const prev = snapRef.current
    let changed = false
    const now: Record<string, unknown> = {}
    for (const r of refs) {
      const k = locatorOf(r.urn ?? `node://${r.nodeId}#${r.fieldId}`)
      const v = pickValue(r)
      now[k] = v
      if (!prev || Object.is(prev[k], v) === false) changed = true
    }
    if (!prev) {
      // 首次只建基线不写（避免刚关联就覆盖文件）
      snapRef.current = now
      return null
    }
    if (!changed) return null
    return refresh()
  }, [opt, pickValue, refresh])

  return { busy, last, refresh, autoRefresh, resolveSites }
}
