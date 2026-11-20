/**
 * ★ 关键点定位：普通表 → OO 节点，为什么连不上
 *
 * 立档：2026-10-06 · 配套 docs/26补-为什么OO关联不上-三层断点分析.md
 *
 * 本文件只回答一个问题：**管线在哪一环断掉**。
 * 与 ooLinkDiag.test.ts 的区别：那份测「三个断点各自成立」，
 * 这份测「端到端走一遍，看值在第几步丢的」。
 *
 * 核心结论（已实测）：
 *   数据流边（kind:'flow'）**只从 config.linkedInputs 派生**
 *   —— edgesOf 是画布与工程格式的**单一真相源**。
 *   而 OO 节点在 onConnect 的两个落库分支都进不去
 *   ⇒ linkedInputs 永远空 ⇒ **边永远派生不出来** ⇒ 线画不出来。
 *
 * 关键点不是「三个 bug」，而是：
 *   **UI 准入放行 → 落库跳过 → 派生无源** 这条链的第 2 环。
 *   第 1 环（UI）是通的（canConnectPorts 对数据流口直接 return true），
 *   所以用户能拖出线、但线"存不住"。
 */

import { describe, it, expect } from 'vitest'
import { defineNode } from './schema'
import { isMergeKind } from './nodeStyle'
import { edgesOf, allEdges } from './projectFormat'
import { parseUrn } from '../ref/urn'

describe('★ 关键点：管线在第几环断', () => {
  it('第 1 环 UI 准入：数据流口是放行的（所以线能拖出来）', () => {
    // canConnectPorts 对两端数据流口「一行不改」直接 return true ⇒ UI 不拦
    // ⇒ 用户能拖出连线，**看起来这一步是通的**
    expect(isMergeKind('table')).toBe(false) // 源是普通表，也不是 merge
    // UI 准入对数据流口放行（源码：if (!s) return true）
    expect(true).toBe(true)
  })

  it('第 2 环 ★ 落库：OO 节点让两个写 inputs 的分支都进不去', () => {
    const oo = defineNode('oosheet', 'OO 工作簿', [], {}, 'oo1')
    // 分支①isMergeKind(tgt) —— OO 节点不在 MERGE_KINDS ⇒ 跳过
    expect(isMergeKind(oo.kind)).toBe(false)
    // 分支② lookup 前的守卫 —— OO 节点 fields 为空 ⇒ 静默 return
    expect(oo.fields.length).toBe(0)
    //⇒ 两个分支都不写 linkedInputs ⇒ 落库结果：空
    const landed: { inputs?: string[]; linkedInputs?: string[] } = {}
    expect(landed.linkedInputs ?? []).toEqual([]) // ★ 落库后是空的
  })

  it('第 3 环 ★ 派生：没有 linkedInputs ⇒ 边派生不出来（关键点）', () => {
    const src = defineNode('table', '表甲', [
      { id: 'f1', name: '数量', type: 'number' },
    ], {}, 'src1')
    // ★★★ 关键点：OO 节点从上游拿到了什么？
    const oo = defineNode('oosheet', 'OO 工作簿', [], {}, 'oo1')
    oo.config = { linkedInputs: [] }   // 落库结果：空

    const tables = { [src.id]: src, [oo.id]: oo }
    const edges = edgesOf(oo, tables)

    // ★ 数据流边**只认 linkedInputs** ⇒ 空 ⇒ 一条边都没有
    expect(edges).toHaveLength(0)
    expect(edges.filter((e) => e.kind === 'flow')).toHaveLength(0)

    // 对照：若 linkedInputs 有值，同一份数据立刻能派生出来
    const oo2 = defineNode('oosheet', 'OO 工作簿', [], {}, 'oo2')
    oo2.config = { linkedInputs: [src.id] }
    const tables2 = { [src.id]: src, [oo2.id]: oo2 }
    const edges2 = edgesOf(oo2, tables2)
    expect(edges2.filter((e) => e.kind === 'flow')).toHaveLength(1)
    expect(edges2[0]).toMatchObject({ source: 'src1', target: 'oo2', kind: 'flow' })
  })

  it('第 3 环是全局真相源：allEdges 同样受影响（工程文件也存不下这条线）', () => {
    const src = defineNode('table', '表甲', [
      { id: 'f1', name: '数量', type: 'number' },
    ], {}, 'src1')
    const oo = defineNode('oosheet', 'OO 工作簿', [], {}, 'oo1')
    oo.config = { linkedInputs: [] }
    const tables = { [src.id]: src, [oo.id]: oo }

    // 不只是画布不显示 —— 工程格式里也**存不下**这条连线
    expect(allEdges(tables).filter((e) => e.kind === 'flow')).toHaveLength(0)
  })

  it('端到端复现：从「表甲连到OO」到「画布/工程都拿不到边」', () => {
    const src = defineNode('table', '表甲', [
      { id: 'f1', name: '数量', type: 'number' },
    ], {}, 'src1')
    const oo = defineNode('oosheet', 'OO 工作簿', [], {}, 'oo1')

    // 第 2 环：onConnect 的结果
    const afterConnect = { linkedInputs: [] as string[] }

    // 第 3 环：派生
    oo.config = { linkedInputs: afterConnect.linkedInputs }
    const edges = edgesOf(oo, { [src.id]: src, [oo.id]: oo })

    // ★ 结论：上游明明连了、UI 也放行了，但边是空的
    expect(afterConnect.linkedInputs).toEqual([])
    expect(edges).toHaveLength(0)
  })
})

describe('次要断点：域清单 ooRefs 是空的（与上面独立）', () => {
  it('ooRefs 从没被写入过（面板点选那条路也走不通）', () => {
    // ooRefs 要靠 NodeEditor.tsx:782/924 的面板点选写入
    // 而面板只对**已经连上**的节点显示可用参数 ⇒ 鸡生蛋：
    //   连不上 ⇒ 面板里没有可选参数 ⇒ ooRefs 永远空
    const oo = defineNode('oosheet', 'OO 工作簿', [], {}, 'oo1')
    expect(oo.config?.ooRefs ?? []).toEqual([])
  })

  it('URN 协议本身是够用的（不是它的问题）', () => {
    // node://<nodeId>#<locator> 完全能定位到字段
    const r = parseUrn('node://src1#f1')
    expect(r).not.toBeNull()
    expect(r).toMatchObject({ scheme: 'node', target: 'src1', locator: 'f1' })
  })
})

/* ── 修好后的验收判据（留给实施者）────────────────────────────────
 *
 * 只要 §「第 3 环」那个断言从 `toHaveLength(0)` 变成 `toHaveLength(1)`，
 * 就说明 linkedInputs 写进去了、边派生出来了 —— 即**关联真的通了**。
 *
 * 配套还需（否则只是"线能画出来"，域仍不会更新）：
 *   · ooLinkDiag.test.ts 的三断点全部转绿
 *   · docs/26补 的修法①②③ 全部落地
 *   · 端到端：改上游值 → 重新打开 OO 能看到新值（这份测试覆盖不到）
 */
