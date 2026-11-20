/**
 * 「点击 → 引用参数 → 形成关联」链路测试
 *
 * 立档：2026-10-06 · 用户：「我先不做线，我做成点击后能引用参数，然后形成关联」
 *
 * 本文件回答一个问题：**走"点选"这条路，关联真的建立起来了吗？**
 * 与 ooLinkKeyPoint.test.ts（走"连线"那条路，断了）互为对照。
 */

import { describe, it, expect } from 'vitest'
import {
  ooRefUrn, refUrnOf, addOoRef, removeOoRef,
  linkedInputsFromOoRefs, ooRefPatchOf, ooRefValue, ooRefLabel,
} from './ooRef'
import { defineNode } from './schema'
import { edgesOf } from './projectFormat'
import { parseUrn } from '../ref/urn'

describe('① 点选即写域清单，且带标准 URN', () => {
  it('新增一项时补上 urn', () => {
    const refs = addOoRef(undefined, 'table_1', 'f1')
    expect(refs).toHaveLength(1)
    expect(refs[0]).toEqual({ nodeId: 'table_1', fieldId: 'f1', urn: 'node://table_1#f1' })
  })

  it('重复点同一项不重复登记（并返回原引用以便调用方跳过 setState）', () => {
    const a = addOoRef(undefined, 't1', 'f1')
    const b = addOoRef(a, 't1', 'f1')
    expect(b).toHaveLength(1)
    expect(b).toBe(a) // ★ 同一引用 ⇒ 调用方可跳过写库，避免多一次历史快照
  })

  it('不同字段各占一项', () => {
    const refs = addOoRef(addOoRef(undefined, 't1', 'f1'), 't1', 'f2')
    expect(refs).toHaveLength(2)
  })

  it('取消引用', () => {
    const refs = removeOoRef(addOoRef(undefined, 't1', 'f1'), 't1', 'f1')
    expect(refs).toHaveLength(0)
  })

  it('urn 与 ref/urn.ts 的解析严格对称（写进去能读出来）', () => {
    const urn = ooRefUrn('t1', 'f1')
    expect(urn).toBe('node://t1#f1')
    const r = parseUrn(urn)
    expect(r).toMatchObject({ scheme: 'node', target: 't1', locator: 'f1' })
  })

  it('老数据（只有 nodeId/fieldId，无 urn）能就地补齐，不需迁移', () => {
    expect(refUrnOf({ nodeId: 't1', fieldId: 'f1' })).toBe('node://t1#f1')
    // 已有 urn 的不覆盖
    expect(refUrnOf({ nodeId: 't1', fieldId: 'f1', urn: 'node://other#x' })).toBe('node://other#x')
  })
})

describe('★★ 核心：点选后关联真的形成（边派生得出来）', () => {
  it('ooRefs → linkedInputs 推导', () => {
    const refs = addOoRef(addOoRef(undefined, 'srcA', 'f1'), 'srcB', 'qty')
    expect(linkedInputsFromOoRefs(refs)).toEqual(['srcA', 'srcB'])
  })

  it('★ 端到端：表甲 --点选 f1--> OO 工作簿 ⇒ edgesOf 派生出 flow 边', () => {
    const src = defineNode('table', '表甲', [
      { id: 'f1', name: '数量', type: 'number' },
    ], {}, 'src1')
    const oo = defineNode('oosheet', 'OO 工作簿', [], {}, 'oo1')

    // 用户操作：面板里点「表甲 · 数量」
    const patch = ooRefPatchOf(oo.config, 'src1', 'f1')
    oo.config = { ...(oo.config ?? {}), ...patch }

    // ★ 之前（走连线）这里是 0 条边；现在应该派生出 1 条
    const edges = edgesOf(oo, { [src.id]: src, [oo.id]: oo })
    const flow = edges.filter((e) => e.kind === 'flow')
    expect(flow).toHaveLength(1)
    expect(flow[0]).toMatchObject({ source: 'src1', target: 'oo1', kind: 'flow' })
  })

  it('★ 与「连线」路径对照：两条路最终产出同构（linkedInputs 是唯一真相）', () => {
    // 连线路径：App.tsx 的 isMergeKind 分支写 linkedInputs
    // 点选路径：ooRefPatchOf 写 linkedInputs
    // ⇒ 两者产出的边完全一样 ⇒ 下游（画布/工程/触发）无需区分来源
    const src = defineNode('table', '表甲', [{ id: 'f1', name: '数量', type: 'number' }], {}, 'src1')
    const byWire = defineNode('oosheet', 'OO', [], {}, 'a')
    byWire.config = { linkedInputs: ['src1'] }
    const byClick = defineNode('oosheet', 'OO', [], {}, 'b')
    byClick.config = ooRefPatchOf(byClick.config, 'src1', 'f1')

    const t1 = { [src.id]: src, [byWire.id]: byWire }
    const t2 = { [src.id]: src, [byClick.id]: byClick }
    // ★ 归一化必须**去掉 target**：两个节点 id 不同（a / b），
    //   target 天然不同 ⇒ 带上它就恒不等（这条红过一次，是我测试写错不是代码错）。
    //   要比的是「边有没有被派生出来」这件事本身。
    const norm = (e: { id: string; source: string; kind: string }) => ({
      kind: e.kind, source: e.source,
    })
    expect(edgesOf(byWire, t1).map(norm)).toEqual(edgesOf(byClick, t2).map(norm))
    // 且两边都真的派生出边（空数组不会「相等」成假通过）
    expect(edgesOf(byWire, t1)).toHaveLength(1)
    expect(edgesOf(byClick, t2)).toHaveLength(1)
  })

  it('★ 一次 patch 同时落两个字段（不拆成两次写）', () => {
    const oo = defineNode('oosheet', 'OO', [], {}, 'oo1')
    const p = ooRefPatchOf(oo.config, 'src1', 'f1')
    expect(p.ooRefs).toHaveLength(1)
    expect(p.linkedInputs).toEqual(['src1'])
  })

  it('★ 不抹掉手拖连线写入的上游（并集，不能只写ooRefs 推出来的）', () => {
    // 用户先用连线连了 manualUpstream，又点了参数 src1
    const oo = defineNode('oosheet', 'OO', [], {}, 'oo1')
    oo.config = { linkedInputs: ['manualUpstream'] }
    const p = ooRefPatchOf(oo.config, 'src1', 'f1')
    expect(p.linkedInputs).toEqual(['src1', 'manualUpstream'])
  })

  it('多次点选累积（不覆盖之前的）', () => {
    let cfg = defineNode('oosheet', 'OO', [], {}, 'oo1').config
    cfg = ooRefPatchOf(cfg, 'srcA', 'f1')
    cfg = ooRefPatchOf(cfg, 'srcB', 'f2')
    expect(cfg.ooRefs).toHaveLength(2)
    expect(cfg.linkedInputs).toEqual(['srcA', 'srcB'])
  })
})

describe('③ 取值与可读标签（写回管线要用）', () => {
  it('按引用取首行值', () => {
    const tables = {
      t1: { fields: [{ id: 'f1', name: '数量' }], rows: [{ f1: 12 }] },
    }
    expect(ooRefValue({ nodeId: 't1', fieldId: 'f1' }, tables)).toBe(12)
  })

  it('节点/字段不存在返回 undefined（由调用方决定跳过还是写空）', () => {
    expect(ooRefValue({ nodeId: 'nope', fieldId: 'f1' }, {})).toBeUndefined()
    expect(ooRefValue({ nodeId: 't1', fieldId: 'nope' }, { t1: { fields: [], rows: [] } })).toBeUndefined()
  })

  it('可读标签不暴露 node://（避免 Excel 显示 #NAME?）', () => {
    // ★ 这条对应批次 17 的教训：把 node:// 写进单元格，Excel 打开是 #NAME?
    const label = ooRefLabel({ nodeId: 't1', fieldId: 'f1', urn: 'node://t1#f1' }, '表甲', '数量')
    expect(label).toBe('表甲.数量')
    expect(label).not.toContain('node://')
  })

  it('只引用整节点时标签是节点名', () => {
    expect(ooRefLabel({ nodeId: 't1', fieldId: '' }, '表甲')).toBe('表甲')
  })
})
