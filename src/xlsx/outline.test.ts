import { describe, it, expect } from 'vitest'

import {
  renumber,
  computeNumbers,
  normalizeLevels,
  parseLevel,
  takeSubtree,
  insertSubtree,
  moveRows,
  subtreeSpan,
  isAncestorOf,
  assertMovable,
  shiftLevel,
  buildTree,
  countTitles,
  findParentRowIndex,
  visibleRowIndexes,
  toggleCollapse,
  setCollapseAll,
  flattenTree,
  commitStructure,
  makeRow,
  makeRows,
  makeRowsByLevel,
  MAX_LEVEL,
  type OutlineRow,
} from './outline'

/** 造一份典型的造价表骨架：1 / 1.1 / 1.2 / 数据行 / 2 / 2.1 / 3 */
function sample(): OutlineRow[] {
  return makeRows(['1', '1.1', null, '1.2', null, null, '2', '2.1', '3'])
}

describe('renumber —— 章节号重算（计数器栈）', () => {
  it('基本编号连续', () => {
    const rows = sample()
    renumber(rows)
    expect(rows.map((r) => r.no)).toEqual(['1', '1.1', '', '1.2', '', '', '2', '2.1', '3'])
  })

  it('★ 回到浅层级时截断深计数器，否则出现重复编号', () => {
    // 1 / 1.1 / 1.1.1 / 1.1.2 / 2 / 2.1 ←若不截断，2.1 会重复
    const rows = makeRows(['1', '1.1', '1.1.1', '1.1.2', '2', '2.1'])
    renumber(rows)
    expect(rows.map((r) => r.no)).toEqual(['1', '1.1', '1.1.1', '1.1.2', '2', '2.1'])
    expect(new Set(rows.map((r) => r.no)).size).toBe(6)
  })

  it('普通行不清计数器（编号延续）', () => {
    const rows = makeRows(['1', null, null, null, '1.1'])
    renumber(rows)
    expect(rows.map((r) => r.no)).toEqual(['1', '', '', '', '1.1'])
  })

  it('空数组不炸', () => {
    expect(() => renumber([])).not.toThrow()
  })

  it('全普通行 → 全部空编号', () => {
    const rows = makeRows([null, null, null])
    renumber(rows)
    expect(rows.every((r) => r.no === '')).toBe(true)
  })

  it('computeNumbers 不污染传入的 rows（只算不写）', () => {
    const rows = sample()
    const before = rows.map((r) => r.no)
    const out = computeNumbers(rows)
    expect(out[0]).toBe('1')
    expect(rows.map((r) => r.no)).toEqual(before) // 原对象未被改
  })

  it('五级深嵌套', () => {
    const rows = makeRows(['1', '1.1', '1.1.1', '1.1.1.1', '1.1.1.1.1'])
    renumber(rows)
    expect(rows[4].no).toBe('1.1.1.1.1')
  })
})

describe('normalizeLevels —— 跳级钳制', () => {
  it('一级下直接三级 → 钳为二级，返回 true', () => {
    const rows = makeRows(['1', '1.1.1'])
    expect(normalizeLevels(rows)).toBe(true)
    expect(rows[1].level).toBe(2)
  })

  it('无跳级时返回 false', () => {
    const rows = makeRows(['1', '1.1', '1.1.1', '2'])
    expect(normalizeLevels(rows)).toBe(false)
  })

  it('不产生 1.0.0.1 这类脏编号', () => {
    const rows = makeRows(['1', '1.1.1.1', null, '1.1.1.1.1'])
    normalizeLevels(rows)
    renumber(rows)
    for (const r of rows) {
      if (r.no) expect(r.no).not.toMatch(/(^|\.)0(\.|$)/)
    }
  })

  it('普通行不改变 prev（三级前的普通行不该把上限压低）', () => {
    // 1 / 普通 / 3→应钳为 2（因上一标题是 1）
    const rows = makeRows(['1', null, '1.1.1'])
    normalizeLevels(rows)
    expect(rows[2].level).toBe(2)
  })

  it('钳制后再重算，编号严格递增无跳号', () => {
    const rows = makeRows(['1', '1.1.1', '1.1.1.1', '2'])
    normalizeLevels(rows)
    renumber(rows)
    const nos = rows.filter((r) => r.no).map((r) => r.no)
    expect(nos).toEqual(['1', '1.1', '1.1.1', '2'])
  })

  it('连续跳级多次修正', () => {
    const rows = makeRows(['1', '1.1.1.1.1'])
    normalizeLevels(rows)
    renumber(rows)
    expect(rows[1].level).toBe(2)
    expect(rows[1].no).toBe('1.1')
  })

  it('level 不会超过 MAX_LEVEL', () => {
    const rows = makeRows(['1'])
    rows[0].level = 9 as never
    normalizeLevels(rows)
    renumber(rows)
    expect(rows[0].level).toBeLessThanOrEqual(MAX_LEVEL)
  })
})

describe('parseLevel —— 标记列解析', () => {
  it('数字模式 1..5', () => {
    expect(parseLevel(1)).toBe(1)
    expect(parseLevel('3')).toBe(3)
    expect(parseLevel(5)).toBe(5)
    expect(parseLevel(0)).toBe(0)
  })

  it('字符模式 # / ## / ###', () => {
    expect(parseLevel('#')).toBe(1)
    expect(parseLevel('##')).toBe(2)
    expect(parseLevel('###')).toBe(3)
  })

  it('★ 全角＃兼容（用户从别处拷来常见）', () => {
    expect(parseLevel('＃＃')).toBe(2)
  })

  it('脏输入归0 且不抛错（不能因一格脏数据就打不开用户文件）', () => {
    expect(parseLevel('')).toBe(0)
    expect(parseLevel('   ')).toBe(0)
    expect(parseLevel(null)).toBe(0)
    expect(parseLevel(undefined)).toBe(0)
    expect(parseLevel('abc')).toBe(0)
    expect(parseLevel(9)).toBe(0)
    expect(parseLevel(-1)).toBe(0)
    expect(parseLevel(1.5)).toBe(0)
    expect(parseLevel({})).toBe(0)
  })

  it('前后空格容忍', () => {
    expect(parseLevel('  2  ')).toBe(2)
  })

  it('超过 5 个 # 截到 5', () => {
    expect(parseLevel('##########')).toBe(5)
  })
})

describe('takeSubtree / insertSubtree / subtreeSpan', () => {
  it('摘出含全部后代（含普通数据行，直到下一个同级标题）', () => {
    const rows = sample()
    // level 序列 1,2,0,2,0,0,1,2,1 → '1' 的段 = [1,1.1,普通,1.2,普通,普通] 共 6 行
    const seg = takeSubtree(rows, 0)
    expect(seg.length).toBe(6)
    expect(rows.length).toBe(3) //剩 '2' / '2.1' / '3'
  })

  it('★ 普通行不作为拖动起点（归属于上方标题）', () => {
    const rows = makeRows(['1', '1.1', null, null])
    const mv = assertMovable(rows, 2)
    expect(mv.ok).toBe(false)
    expect(mv.reason).toContain('普通数据行')
  })

  it('普通行自身 span 一路到下一个标题（但拖动被 assertMovable 拦住）', () => {
    const rows = makeRows(['1', '1.1', null, null])
    expect(takeSubtree(rows, 2).length).toBe(2) // 到下一个标题前
    expect(assertMovable(rows, 2).ok).toBe(false) // 但不允许拖
  })

  it('★ insertSubtree 插到指定下标（按 no 断言，不按 title —— makeRows 会造重名标题）', () => {
    const rows = makeRowsByLevel([1, 1], ['A', 'B'])
    renumber(rows)
    const seg = takeSubtree(rows, 0)
    insertSubtree(rows, seg, 1)
    renumber(rows)
    // B 先变二级再被 normalize 钳回一级 → 顺序 A B，编号 1 / 2
    expect(rows.map((r) => r.no)).toEqual(['1', '2'])
  })

  it('★ insertSubtree 下标越界自动夹到合法范围（不炸）', () => {
    const rows = makeRows(['1', '2'])
    const seg = takeSubtree(rows, 0)
    expect(() => insertSubtree(rows, seg, 999)).not.toThrow()
    expect(rows.length).toBe(2)
  })

  it('subtreeSpan 只算不改', () => {
    const rows = sample()
    expect(subtreeSpan(rows, 0)).toBe(6)
    expect(subtreeSpan(rows, 6)).toBe(2) // '2' + '2.1'
    expect(rows.length).toBe(9) // 未被修改
  })

  it('越界安全', () => {
    const rows = sample()
    expect(subtreeSpan(rows, -1)).toBe(0)
    expect(subtreeSpan(rows, 999)).toBe(0)
    expect(takeSubtree(rows, 999)).toEqual([])
  })
})

describe('isAncestorOf —— 防成环', () => {
  it('真祖先', () => {
    const rows = makeRows(['1', '1.1', '1.1.1', '2'])
    expect(isAncestorOf(rows, 0, 2)).toBe(true)
  })
  it('★ 隔了同级标题就不再是祖先（段落法）', () => {
    // 1 / 1.1 / 2 / 2.1：'1' 不是 '2.1' 的祖先 —— 中间隔了同级标题 '2'
    const rows = makeRows(['1', '1.1', '2', '2.1'])
    expect(isAncestorOf(rows, 0, 3)).toBe(false)
    expect(isAncestorOf(rows, 2, 3)).toBe(true) // '2' 是 '2.1' 的祖先
  })
  it('★ 段落法解扁平 level 歧义：兄弟的段不重叠', () => {
    // level 序列 1,2,2,3：下标 1 与 2 都是 level 2（兄弟）
    const rows = makeRows(['1', '1.1', '1.2', '1.2.1'])
    expect(isAncestorOf(rows, 1, 2)).toBe(false) // 兄弟
    expect(isAncestorOf(rows, 2, 3)).toBe(true)  // 1.2 是 1.2.1 的父
    expect(isAncestorOf(rows, 0, 3)).toBe(true)  // 1 是 1.2.1 的祖
  })
  it('非祖先（i 在 j 之后）', () => {
    const rows = makeRows(['1', '1.1', '1.1.1', '2'])
    expect(isAncestorOf(rows, 2, 0)).toBe(false)
  })
  it('普通行不是祖先', () => {
    const rows = makeRows([null, null])
    expect(isAncestorOf(rows, 0, 1)).toBe(false)
  })
  it('assertMovable：标题可动、普通行不可动', () => {
    const rows = makeRows(['1', null])
    expect(assertMovable(rows, 0).ok).toBe(true)
    expect(assertMovable(rows, 1).ok).toBe(false)
    expect(assertMovable(rows, 99).ok).toBe(false)
  })
})

describe('moveRows —— 移动统一入口（方案验收用例 1 / 2）', () => {
  it('★ 用例1：拖 1.2 段到前面，其下子标题整段跟走并自动重排', () => {
    // level 序列 [1,2,2,3]：总说明 / 编制依据 / 工程概况 / 地质条件
    // ★ 用 makeRowsByLevel：点号串无法区分「1.1/1.2是兄弟」与「1.2.1是1.2的子级」
    const rows = makeRowsByLevel([1, 2, 2, 3], ['总说明', '编制依据', '工程概况', '地质条件'])
    renumber(rows)
    expect(rows.map((r) => r.no)).toEqual(['1', '1.1', '1.2', '1.2.1'])
    // 拖 '工程概况'（下标 2，段含 '地质条件' 共 2 行）到下标 1（编制依据之前）
    const res = moveRows(rows, 2, 1)
    expect(res.ok).toBe(true)
    expect(res.moved).toBe(2) // 工程概况 + 地质条件 整段跟走
    // 文档顺序：总说明 / 工程概况 / 地质条件 / 编制依据
    expect(rows.map((r) => r.title)).toEqual(['总说明', '工程概况', '地质条件', '编制依据'])
    // 工程概况现在直接跟在一级标题后 → 被钳为一级，编号 1 / 2 / 2.1 / 2.2
    expect(rows.map((r) => r.no)).toEqual(['1', '2', '2.1', '2.2'])
    expect(res.clamped).toBe(true) // 发生了钳制，UI 应提示
  })

  it('★ 用例1变体：拖到最前（前面无同级）时子标题仍跟走', () => {
    // 1 / 2 / 2.1 —— 拖 '2' 到下标 0
    const rows = makeRows(['1', '2', '2.1'])
    renumber(rows)
    const res = moveRows(rows, 1, 0)
    expect(res.ok).toBe(true)
    expect(res.moved).toBe(2)
    expect(rows.map((r) => r.no)).toEqual(['1', '1.1', '2'])
  })

  it('★ 用例2：把三级标题拖到一级标题下方 → 钳为二级，clamped=true', () => {
    const rows = makeRows(['1', '1.1.1'])
    renumber(rows)
    const res = moveRows(rows, 1, 0, { asChild: false })
    void res
    // 单独验证钳制路径
    const rows2 = makeRows(['1', '2', '1.1.1'])
    normalizeLevels(rows2)
    expect(rows2[2].level).toBe(2)
  })

  it('★ 落点越界返回失败而不是炸', () => {
    const rows = sample()
    expect(moveRows(rows, -1, 0).ok).toBe(false)
    expect(moveRows(rows, 0, 999).ok).toBe(false)
  })

  it('★ 普通行不可作容器（方案 §8）', () => {
    const rows = makeRows(['1', null, '2'])
    const res = moveRows(rows, 2, 1) // 想拖到普通行上
    expect(res.ok).toBe(false)
    expect(res.reason).toContain('普通数据行')
  })

  it('不能把祖先拖进自己子树（防环）', () => {
    const rows = makeRows(['1', '1.1', '1.1.1'])
    const res = moveRows(rows, 0, 2, { asChild: true })
    expect(res.ok).toBe(false)
    expect(res.reason).toContain('子树')
  })

  it('asChild：拖到某标题上 → 整段变它的子级并降一级', () => {
    const rows = makeRowsByLevel([1, 2, 1], ['一', '一之一', '二'])
    renumber(rows)
    // 把 '二'（下标2）拖到 '一'（下标0）上做子级 → 插到 '一' 段尾、整体 +1 级
    const res = moveRows(rows, 2, 0, { asChild: true })
    expect(res.ok).toBe(true)
    expect(rows.map((r) => r.title)).toEqual(['一', '一之一', '二'])
    expect(rows.map((r) => r.level)).toEqual([1, 2, 2])
    expect(rows.map((r) => r.no)).toEqual(['1', '1.1', '1.2'])
  })

  it('★ asChild 深层降级溢出时被全局钳制（不丢层级语义）', () => {
    // 地质条件(L3) 拖到 编制依据(L2) 上 → L4 溢出，钳到 L3（仍是编制依据的子级）
    const rows = makeRowsByLevel([1, 2, 2, 3], ['总说明', '编制依据', '工程概况', '地质条件'])
    renumber(rows)
    const res = moveRows(rows, 3, 1, { asChild: true })
    expect(res.ok).toBe(true)
    expect(res.clamped).toBe(true)
    expect(rows.map((r) => r.title)).toEqual(['总说明', '编制依据', '地质条件', '工程概况'])
    expect(rows.map((r) => r.level)).toEqual([1, 2, 3, 2])
  })

  it('同下标移动 = 空操作', () => {
    const rows = sample()
    const res = moveRows(rows, 1, 1)
    expect(res.ok).toBe(true)
    expect(res.moved).toBe(0)
  })

  it('移动后编号连续无重复', () => {
    const rows = makeRowsByLevel([1, 2, 3, 1, 2, 1], ['A', 'A.1', 'A.1.1', 'B', 'B.1', 'C'])
    renumber(rows)
    const res = moveRows(rows, 5, 0) // 拖 'C' 到最前
    expect(res.ok).toBe(true)
    const nos = rows.filter((r) => r.no).map((r) => r.no)
    expect(new Set(nos).size).toBe(nos.length) // 无重复
    // C 变一级，与 A 同级；A.1 / A.1.1 随A 跟走
    expect(nos).toEqual(['1', '2', '2.1', '2.1.1', '3', '3.1'])
    expect(rows[0].title).toBe('C')
  })
})

describe('shiftLevel —— 升级 / 降级', () => {
  it('升级：连同子树全部下移一级', () => {
    const rows = makeRows(['1', '1.1', '1.1.1', '2'])
    shiftLevel(rows, 1, 1)
    expect(rows[1].level).toBe(2)
    expect(rows[2].level).toBe(3) // 子树跟走
  })

  it('降级：子树一起降，不倒挂', () => {
    const rows = makeRows(['1', '1.1', '1.1.1', '2'])
    shiftLevel(rows, 1, -1)
    // 1.1 降为一级，其子 1.1.1 必须跟着降，否则会出现 1 之后直接 3
    expect(rows[1].level).toBe(1)
    expect(rows[2].level).toBe(2)
  })

  it('降级后编号仍连续', () => {
    const rows = makeRows(['1', '1.1', '1.1.1', '1.1.1.1', '2'])
    shiftLevel(rows, 1, -1)
    const nos = rows.filter((r) => r.no).map((r) => r.no)
    expect(nos).toEqual(['1', '2', '2.1', '2.1.1', '3'])
  })

  it('越界不生效（1 级不能再升）', () => {
    const rows = makeRowsByLevel([1, 1], ['A', 'B'])
    const r = shiftLevel(rows, 0, 1)
    // A 是首行，升级到 2 后被 normalizeLevels 钳回 1（首行不能是二级）→ clamped
    expect(r.moved).toBe(1)
    expect(r.clamped).toBe(true)
    expect(rows[0].level).toBe(1)
  })

  it('二级升三级：前一行是二级时不钳制', () => {
    // A(L1) / A.1(L2) / A.2(L2) —— 升 A.2 到 3，prev=2，max=3 → 合法
    const rows = makeRowsByLevel([1, 2, 2], ['A', 'A.1', 'A.2'])
    const r = shiftLevel(rows, 2, 1)
    expect(r.clamped).toBe(false)
    expect(rows[2].level).toBe(3)
    expect(rows.map((x) => x.no)).toEqual(['1', '1.1', '1.1.1'])
  })

  it('二级升三级：前一行是一级时钳制（跳级）', () => {
    // A(L1) / A.1(L2) —— 升 A.1 到 3，前一行 L1，max=2 → 钳回
    const rows = makeRowsByLevel([1, 2], ['A', 'A.1'])
    const r = shiftLevel(rows, 1, 1)
    expect(r.clamped).toBe(true)
    expect(rows[1].level).toBe(2)
  })

  it('★ 普通行不能升降级（不是章节）', () => {
    const rows = makeRows([null, '1'])
    const r = shiftLevel(rows, 0, 1)
    expect(r.moved).toBe(0)
    expect(rows[0].level).toBe(0) // 不静默改成标题
  })

  it('越界下标安全', () => {
    const rows = sample()
    expect(() => shiftLevel(rows, 999, 1)).not.toThrow()
  })
})

describe('buildTree —— 从 rows 派生树', () => {
  it('基本层级正确', () => {
    const rows = sample()
    renumber(rows)
    const tree = buildTree(rows)
    expect(tree.length).toBe(3) // 1 / 2 / 3
    expect(tree[0].no).toBe('1')
    expect(tree[0].children.length).toBe(2) // 1.1 / 1.2
  })

  it('rowIndex 指向正确下标', () => {
    const rows = sample()
    renumber(rows)
    const tree = buildTree(rows)
    expect(tree[0].rowIndex).toBe(0)
    expect(tree[1].rowIndex).toBe(6)
    expect(tree[2].rowIndex).toBe(8)
  })

  it('普通行只累加 bodyCount，不进树', () => {
    const rows = makeRows(['1', null, null, '1.1', null])
    renumber(rows)
    const tree = buildTree(rows)
    expect(tree.length).toBe(1)
    expect(tree[0].bodyCount).toBe(2) // 直属两个普通行
    expect(tree[0].children[0].bodyCount).toBe(1)
  })

  it('hasChildren 只表示有标题子级', () => {
    const rows = makeRows(['1', null])
    renumber(rows)
    const tree = buildTree(rows)
    expect(tree[0].hasChildren).toBe(false) // 只有普通行
    expect(tree[0].bodyCount).toBe(1)
  })

  it('全普通行 → 空树', () => {
    expect(buildTree(makeRows([null, null, null])).length).toBe(0)
  })

  it('countTitles 递归计数', () => {
    const rows = makeRows(['1', '1.1', '1.1.1', '2'])
    renumber(rows)
    expect(countTitles(buildTree(rows))).toBe(4)
  })

  it('★ 树不含普通行，且每节点 rowIndex 唯一', () => {
    const rows = sample()
    renumber(rows)
    const all: number[] = []
    const walk = (ns: ReturnType<typeof buildTree>) => {
      for (const n of ns) {
        all.push(n.rowIndex)
        walk(n.children)
      }
    }
    walk(buildTree(rows))
    expect(new Set(all).size).toBe(all.length)
  })

  it('空 rows → 空树', () => {
    expect(buildTree([])).toEqual([])
  })
})

describe('findParentRowIndex', () => {
  it('找最近浅层祖先', () => {
    const rows = makeRows(['1', '1.1', '1.1.1'])
    expect(findParentRowIndex(rows, 2)).toBe(1)
    expect(findParentRowIndex(rows, 1)).toBe(0)
  })
  it('一级标题无父', () => {
    expect(findParentRowIndex(makeRows(['1']), 0)).toBe(-1)
  })
})

describe('visibleRowIndexes —— 折叠（方案验收用例 4）', () => {
  it('★ 用例4：折叠一个节点 → 它的所有后代（含普通数据行）一起隐藏，标题自己保留', () => {
    // 1 / 1.1 / 普通 / 2 / 2.1 / 普通 / 3
    const rows = makeRowsByLevel([1, 2, 0, 1, 2, 0, 1])
    renumber(rows)
    expect(visibleRowIndexes(rows).length).toBe(7)
    rows[0].collapsed = true // 折叠 '1'
    // '1' 保留；1.1 与其下普通行隐藏；2 / 2.1 / 普通 / 3 保留
    expect(visibleRowIndexes(rows)).toEqual([0, 3, 4, 5, 6])
  })

  it('折叠中间节点：其段内后续行一并隐藏', () => {
    // 1 / 1.1 / 1.2 / 2 —— '1.1' 的段 = [1.1, 1.2]（idx1 是这段的首）
    // 折叠 '1.1' → 隐藏它之后的段内行（idx2），'1.1' 自己与 '1'、'2' 保留
    const rows = makeRowsByLevel([1, 2, 2, 1])
    renumber(rows)
    rows[1].collapsed = true
    expect(visibleRowIndexes(rows)).toEqual([0, 1, 2, 3])
  })

  it('折叠一段多行的标题', () => {
    // 1 / 1.1 / 1.1.1 / 2 —— 折叠 '1' 应隐藏 1.1 与 1.1.1 两行
    const rows = makeRowsByLevel([1, 2, 3, 1])
    renumber(rows)
    rows[0].collapsed = true
    expect(visibleRowIndexes(rows)).toEqual([0, 3])
  })

  it('嵌套折叠：外层折叠则内层全部隐藏', () => {
    const rows = makeRowsByLevel([1, 2, 3, 2, 1])
    renumber(rows)
    rows[0].collapsed = true // 折叠 '1'
    expect(visibleRowIndexes(rows)).toEqual([0, 4])
  })

  it('★ 折叠的标题自己一定可见（否则无从展开）', () => {
    const rows = makeRowsByLevel([1, 2, 3, 1])
    renumber(rows)
    rows[0].collapsed = true
    const vis = visibleRowIndexes(rows)
    expect(vis).toContain(0) // 折叠项自己可见
  })

  it('无折叠时全可见', () => {
    const rows = sample()
    expect(visibleRowIndexes(rows).length).toBe(rows.length)
  })

  it('折叠全部 → 只剩顶层标题，展开全部恢复', () => {
    const rows = makeRowsByLevel([1, 2, 2, 1])
    renumber(rows)
    setCollapseAll(rows, true)
    expect(visibleRowIndexes(rows).length).toBe(2) // 两个一级标题自己
    setCollapseAll(rows, false)
    expect(visibleRowIndexes(rows).length).toBe(4)
  })

  it('toggleCollapse 只影响标题行，普通行无效', () => {
    const rows = makeRows(['1', null])
    expect(toggleCollapse(rows, 0)).toBe(true)
    expect(toggleCollapse(rows, 0)).toBe(false)
    const before = rows[1].collapsed
    toggleCollapse(rows, 1)
    expect(rows[1].collapsed).toBe(before)
  })

  it('★ 折叠状态存在行上，跟着行走（拖动后不错乱）', () => {
    const rows = makeRowsByLevel([1, 2, 1, 2], ['A', 'A.1', 'B', 'B.1'])
    renumber(rows)
    rows[0].collapsed = true
    const seg = takeSubtree(rows, 0) // 摘 'A' 整段（含 A.1）
    expect(seg.length).toBe(2)
    insertSubtree(rows, seg, 2) // 挪到 'B' 之后
    renumber(rows)
    // 折叠标记仍跟着 'A' 走，现在它在末尾（下标 2）
    expect(rows[2].title).toBe('A')
    expect(rows[2].collapsed).toBe(true)
    // 顺序已变 A,B,B.1,A.1 → A 折叠 → A.1 隐藏，其余可见
    expect(rows.map((r) => r.title)).toEqual(['B', 'B.1', 'A', 'A.1'])
    expect(visibleRowIndexes(rows)).toEqual([0, 1, 2])
  })})

describe('flattenTree —— 渲染用扁平化', () => {
  it('带 depth 与 span', () => {
    const rows = makeRows(['1', '1.1', '1.1.1', '2'])
    renumber(rows)
    const flat = flattenTree(buildTree(rows), rows)
    expect(flat.map((f) => f.depth)).toEqual([0, 1, 2, 0])
    expect(flat[0].span).toBe(3)
    expect(flat[3].span).toBe(1)
  })

  it('折叠节点不输出后代', () => {
    const rows = makeRows(['1', '1.1', '1.1.1', '2'])
    renumber(rows)
    rows[0].collapsed = true
    const flat = flattenTree(buildTree(rows), rows)
    expect(flat.length).toBe(2)
    expect(flat.map((f) => f.depth)).toEqual([0, 0])
  })
})

describe('commitStructure —— 唯一收口', () => {
  it('顺序：钳制 → 编号 → 重建树 → 落盘标记', () => {
    const rows = makeRows(['1', '1.1.1'])
    const toasts: string[] = []
    let rebuilt: unknown = null
    let dirty = false
    const res = commitStructure(rows, {
      rebuild: (t) => { rebuilt = t },
      markDirty: () => { dirty = true },
      toast: (m) => toasts.push(m),
    })
    expect(res.clamped).toBe(true)
    expect(rows[1].level).toBe(2)
    expect(rows[1].no).toBe('1.1') // 编号在钳制之后算
    expect(rebuilt).not.toBeNull()
    expect(dirty).toBe(true)
    expect(toasts).toEqual(['已自动修正跳级标题'])
  })

  it('无修正时不给 toast', () => {
    const toasts: string[] = []
    commitStructure(makeRows(['1', '1.1']), { toast: (m) => toasts.push(m) })
    expect(toasts).toEqual([])
  })

  it('无 hooks 也能跑（不炸）', () => {
    expect(() => commitStructure(makeRows(['1', '1.1']))).not.toThrow()
  })
})

describe('方案 §9 验收用例 6 —— 1000 行性能', () => {
  it('1000 行混合标题 + 普通数据，commitStructure < 16ms', () => {
    // 每 7 行一个标题，其余普通行，层级 1..3 循环
    const spec: (string | null)[] = []
    for (let i = 0; i < 1000; i++) {
      spec.push(i % 7 === 0 ? '1' : i % 7 === 3 ? '1.1' : i % 7 === 5 ? '1.1.1' : null)
    }
    const rows = makeRows(spec)
    const t0 = performance.now()
    commitStructure(rows)
    const dt = performance.now() - t0
    expect(rows.length).toBe(1000)
    expect(dt).toBeLessThan(16)
  })

  it('1000 行 buildTree + visibleRowIndexes < 16ms', () => {
    const spec: (string | null)[] = []
    for (let i = 0; i < 1000; i++) {
      spec.push(i % 5 === 0 ? '1' : i % 5 === 2 ? '1.1' : null)
    }
    const rows = makeRows(spec)
    renumber(rows)
    const t0 = performance.now()
    buildTree(rows)
    visibleRowIndexes(rows)
    flattenTree(buildTree(rows), rows)
    const dt = performance.now() - t0
    expect(dt).toBeLessThan(16)
  })
})

describe('铁律：派生数据不手工维护', () => {
  it('★ 任意乱序行数组喂进去，编号连续无跳号无重复', () => {
    const rows = makeRows([
      '1.1.1', '1', null, '1.1', '2', '1.2', null, '1.1.1.1', '3', '1.2.1', '1.2.1.1',
    ])
    normalizeLevels(rows)
    renumber(rows)
    const nos = rows.filter((r) => r.no).map((r) => r.no)
    // 无重复
    expect(new Set(nos).size).toBe(nos.length)
    // 每一级都从 1 开始连续
    const byDepth = new Map<number, number[]>()
    for (const n of nos) {
      const d = n.split('.').length
      byDepth.set(d, [...(byDepth.get(d) ?? []), Number(n.split('.')[d - 1])])
    }
    for (const [, vals] of byDepth) {
      for (let i = 0; i < vals.length; i++) {
        expect(vals[i]).toBeLessThanOrEqual(i + 1) // 不跳号
      }
    }
  })

  it('makeRow 默认字段齐全', () => {
    const r = makeRow()
    expect(r.level).toBe(0)
    expect(r.no).toBe('')
    expect(r.collapsed).toBe(false)
    expect(r.cells).toEqual({})
    expect(r.id).toBeTruthy()
  })

  it('makeRows 支持 5 级且不超界', () => {
    const rows = makeRows(['1.1.1.1.1', '1.1.1.1.1.1'])
    expect(rows[0].level).toBe(5)
    expect(rows[1].level).toBe(5) // 截到MAX
  })

  it('cells 承载业务数据，与 level 互不干扰', () => {
    const rows = [makeRow({ level: 1, title: '总说明', cells: { A: 'x' } })]
    renumber(rows)
    expect(rows[0].no).toBe('1')
    expect(rows[0].cells).toEqual({ A: 'x' })
  })
})
