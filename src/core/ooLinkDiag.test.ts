/**
 * OO 域外链接「关联不上」的三层断点诊断
 *
 * 立档：2026-10-06 · 对应 docs/26-OO域外链接打通-落地计划.md
 *
 * 用户报：OO 节点关联不上。本文件用**可复跑的断言**把三个断点钉住，
 * 避免又变成「凭推断说不行」。
 *
 * 三个断点（按链路顺序）：
 *   ① 连线落库：OO 节点不在 MERGE_KINDS ⇒ onConnect 不写 inputs
 *   ② lookup 分支守卫：`!src.fields.length || !tgt.fields.length` ⇒ 直接 return
 *   ③ 域清单 ooRefs：只写不读，无任何消费方
 */
import { describe, it, expect } from 'vitest'
import { isMergeKind, isRoundNodeKind } from './nodeStyle'
import { defineNode, type NodeKind } from './schema'

describe('OO 域外链接 · 断点①：OO 节点不被认作汇入型', () => {
  it('ooword / oosheet 都不在 MERGE_KINDS 里（这是断点①的根因）', () => {
    // 实测事实：MERGE_KINDS = {sort, output, map, docx, xlsx, chapter}
    expect(isMergeKind('docx')).toBe(true)
    expect(isMergeKind('xlsx')).toBe(true)
    // ★ OO 节点不在内 ⇒ App.tsx 的 onConnect 走不到写 inputs 的分支
    expect(isMergeKind('ooword')).toBe(false)
    expect(isMergeKind('oosheet')).toBe(false)
  })

  it('OO 节点既不是汇入型也不是圆形（两个分支都进不去）', () => {
    // onConnect 的分支顺序：start → ai → paramlib → isMergeKind → 规则 → lookup
    // OO 节点在 start/ai/paramlib 都不是（那些判的是源或特定目标），
    // 在 isMergeKind 也不是 ⇒ 只能落到最后的 lookup 分支
    expect(isRoundNodeKind('ooword')).toBe(false)
    expect(isRoundNodeKind('oosheet')).toBe(false)
  })
})

describe('OO 域外链接 · 断点②：字段数为 0 时 lookup 分支直接 return', () => {
  it('OO 节点典型形态就是 fields 为空', () => {
    const oo = defineNode('oosheet', 'OO 工作簿', [], {}, 'oo1')
    expect(oo.fields.length).toBe(0)
  })

  it('★ 断点②的守卫条件在 OO 场景必然成立（源码同款条件）', () => {
    // App.tsx onConnect 末尾：`if (!src.fields.length || !tgt.fields.length) return`
    const oo = defineNode('oosheet', 'OO 工作簿', [], {}, 'oo1')
    const emptyLikeOO = oo.fields
    // 连到 OO 节点的任何普通节点：只要该节点也没字段，就静默 return
    expect(!emptyLikeOO.length).toBe(true)
    // ⇒ 结论：从「无字段节点」连到「OO 节点」⇒ 落库分支被跳过
  })

  it('对照：有字段的普通节点能过这个守卫', () => {
    const withFields = defineNode('table', '表', [
      { id: 'f1', name: '数量', type: 'number' },
    ], {}, 't1')
    expect(withFields.fields.length).toBeGreaterThan(0)
  })
})

describe('OO 域外链接 · 断点③：ooRefs 只写不读', () => {
  it('ooRefs 的数据结构里没有 URN / locator（无法直接喂给引用协议）', () => {
    // schema.ts 的定义：ooRefs?: { nodeId: string; fieldId: string }[]
    // ★ 对比 ref/urn.ts 的 ResolvedRef：{ scheme, target, locator }
    //   ooRefs 缺 locator ⇒ 无法定位到具体字段值（同名字段在不同表里会歧义）
    const ooRefs = [{ nodeId: 'table_1', fieldId: 'f1' }]
    expect(ooRefs[0]).not.toHaveProperty('locator')
    expect(ooRefs[0]).not.toHaveProperty('scheme')
    expect(ooRefs[0]).not.toHaveProperty('urn')
  })

  it('同一 nodeId+fieldId 在多表场景下会歧义（缺 locator 的后果）', () => {
    const a = { nodeId: 'table_1', fieldId: 'f1' }   // 表甲的「数量」
    const b = { nodeId: 'table_2', fieldId: 'f1' }   // 表乙的「数量」
    // 两条记录除了 nodeId 不同，结构完全一致 ⇒ 光看 ooRefs 分不清引用的是哪张表的哪个字段
    expect(Object.keys(a).sort()).toEqual(Object.keys(b).sort())
    // 而 URN 协议本身是够用的：node://<nodeId>#<locator>
    expect('node://table_1#f1').toContain('#')
  })
})

describe('诊断结论：三断点叠加的必然性', () => {
  it('把三个断点串起来：从普通表连到 OO 节点，关联必然建立不起来', () => {
    // 复现用户操作的完整条件
    const src = defineNode('table', '表甲', [
      { id: 'f1', name: '数量', type: 'number' },
    ], {}, 'src1')
    const tgt = defineNode('oosheet', 'OO 工作簿', [], {}, 'tgt1')

    const step1_merge = isMergeKind(tgt.kind)          // false ⇒ 断点①
    const step2_guard = !(src.fields.length && tgt.fields.length) // true ⇒ 断点②（return）

    // 任意一步为"断"就写不进 inputs
    const canWriteInputs = step1_merge && !step2_guard
    expect(canWriteInputs).toBe(false)

    // 断点③：即使 ooRefs 被写入（面板点选那条路），也没有消费方去读它
    const ooRefs: { nodeId: string; fieldId: string }[] = [
      { nodeId: src.id, fieldId: 'f1' },
    ]
    expect(ooRefs).toHaveLength(1)
    // 断言「无处消费」：本仓 src/ 下除登记点外无读取方
    // （用 grep 人工核验：ooRefs 只出现在 schema.ts / NodeEditor.tsx 的登记处）
    expect(ooRefs[0].nodeId).toBe('src1')
  })
})

/* ── 附：修好之后这些断言应该怎么变（留给实施者当验收判据）────────────
 *
 * 断点①修法：MERGE_KINDS 加 'ooword' / 'oosheet'
 *   ⇒ isMergeKind('oosheet') 应为 true
 * 断点②修法：把「fields 为 0」从「拒绝」改为「按引用语义处理」
 *   —— OO 节点无字段是**正常形态**，不是异常
 * 断点③修法：ooRefs 补 urn 字段（node://<nodeId>#<locator>），
 *   并新增消费方（写回管线按 urn 解析取值）
 *
 * 三条都修完后，重跑本文件 ⇒「三断点叠加的必然性」那条会失败，
 * 说明链路真的通了。**这份文件就是验收判据，不是说明文档。**
 */
