import { describe, it, expect } from 'vitest'
import {
  alignLayout,
  groupIntoRows,
  snapToGrid,
  ALIGN_FALLBACK_H,
  ALIGN_FALLBACK_W,
  ALIGN_GAP_X,
  ALIGN_GRID,
  type AlignSize,
} from './align'
import { layoutLayers } from './graph'
import type { NodePos } from './store'
import type { TableSchema } from './schema'

const S = (w: number, h: number): AlignSize => ({ w, h })
/** n 个等宽节点 */
const sizes = (n: number, w = 200, h = 90): Record<string, AlignSize> =>
  Object.fromEntries([...Array(n)].map((_, i) => [`n${i}`, S(w, h)]))

describe('snapToGrid：吸附到最近网格点', () => {
  it('正好在网格上 → 原样返回', () => {
    expect(snapToGrid(100)).toBe(100)
    expect(snapToGrid(0)).toBe(0)
  })

  it('★ 用四舍五入而非 floor：143 → 140，147 → 140（下靠）但 147 → 140、153 → 160（上靠）', () => {
    // 143 距 140 是 3，距 160 是 17 → 取 140
    expect(snapToGrid(143)).toBe(140)
    // 147 距 140 是 7，距 160 是 13 → 取 140（floor 也是 140，这条不区分）
    expect(snapToGrid(147)).toBe(140)
    // 153 距 160 是 7，距 140 是 13 → 取 160（floor 会错取 140！）
    expect(snapToGrid(153)).toBe(160)
  })

  it('正负误差对称：偏差 3 与偏差 -3 落在同一个网格点', () => {
    // 137 距 140 是 3；143 距 140 也是 3 → 都归140
    expect(snapToGrid(137)).toBe(140)
    expect(snapToGrid(143)).toBe(140)
  })

  it('NaN / Infinity → 0（不让一个坏值污染整张位置表）', () => {
    expect(snapToGrid(NaN)).toBe(0)
    expect(snapToGrid(Infinity)).toBe(0)
    expect(snapToGrid(-Infinity)).toBe(0)
  })

  it('网格非法（0/负数）→ 原样返回，不静默改变语义', () => {
    expect(snapToGrid(137, 0)).toBe(137)
    expect(snapToGrid(137, -10)).toBe(137)
  })

  it('支持自定义网格', () => {
    expect(snapToGrid(12, 10)).toBe(10)
    expect(snapToGrid(16, 10)).toBe(20)
    expect(snapToGrid(16, 5)).toBe(15)
  })
})

describe('groupIntoRows：按 y 区间重叠分行', () => {
  it('y 完全相同 → 同一行', () => {
    const pos: Record<string, NodePos> = {
      a: { x: 0, y: 100 },
      b: { x: 400, y: 100 },
    }
    expect(groupIntoRows(['a', 'b'], pos, { a: S(200, 90), b: S(200, 90) })).toEqual([['a', 'b']])
  })

  it('★ y 不相等但区间相交 → 仍是同一行（用户手拖不会恰好等 y）', () => {
    // a: y100~190，b: y150~240 → 相交 ⇒ 同行
    const pos: Record<string, NodePos> = { a: { x: 0, y: 100 }, b: { x: 400, y: 150 } }
    expect(groupIntoRows(['a', 'b'], pos, { a: S(200, 90), b: S(200, 90) })).toEqual([['a', 'b']])
  })

  it('y 区间不相交 → 分两行', () => {
    // a: 100~190，b: 200~290 → 不相交
    const pos: Record<string, NodePos> = { a: { x: 0, y: 100 }, b: { x: 0, y: 200 } }
    expect(groupIntoRows(['a', 'b'], pos, { a: S(200, 90), b: S(200, 90) })).toEqual([['a'], ['b']])
  })

  it('恰好相接（a 到 190、b 从 190）→ 不同行（用 < 而非 <=）', () => {
    const pos: Record<string, NodePos> = { a: { x: 0, y: 100 }, b: { x: 0, y: 190 } }
    expect(groupIntoRows(['a', 'b'], pos, { a: S(200, 90), b: S(200, 90) })).toEqual([['a'], ['b']])
  })

  it('行内按 x 升序、行间按 y 升序（顺序稳定 = 结果可复现）', () => {
    const pos: Record<string, NodePos> = {
      a: { x: 500, y: 100 },
      b: { x: 100, y: 100 },
      c: { x: 300, y: 400 },
    }
    const sz = { a: S(200, 90), b: S(200, 90), c: S(200, 90) }
    expect(groupIntoRows(['a', 'b', 'c'], pos, sz)).toEqual([['b', 'a'], ['c']])
  })

  it('★ 输入顺序打乱 → 分组结果不变（否则同输入两次调用结果不同）', () => {
    const pos: Record<string, NodePos> = {
      a: { x: 500, y: 100 },
      b: { x: 100, y: 100 },
      c: { x: 300, y: 400 },
    }
    const sz = { a: S(200, 90), b: S(200, 90), c: S(200, 90) }
    const r1 = groupIntoRows(['a', 'b', 'c'], pos, sz)
    const r2 = groupIntoRows(['c', 'b', 'a'], pos, sz)
    const r3 = groupIntoRows(['b', 'a', 'c'], pos, sz)
    expect(r2).toEqual(r1)
    expect(r3).toEqual(r1)
  })

  it('缺尺寸 → 用保守默认高度算区间（不因缺数据把节点错分行）', () => {
    // 默认高 90：a 100~190，b 150~240 → 同行
    const pos: Record<string, NodePos> = { a: { x: 0, y: 100 }, b: { x: 400, y: 150 } }
    expect(groupIntoRows(['a', 'b'], pos, {})).toEqual([['a', 'b']])
    expect(ALIGN_FALLBACK_H).toBeGreaterThan(0)
  })
})

describe('alignLayout：自动对齐（网格归位）', () => {
  it('空输入 → 空输出', () => {
    expect(alignLayout([], {})).toEqual({})
  })

  it('★ 键集合必须与传入 id 完全一致（少一个键 → React Flow 把节点弹回默认位）', () => {
    const ids = ['a', 'b', 'c', 'd']
    const pos: Record<string, NodePos> = {
      a: { x: 13, y: 27 },
      b: { x: 400, y: 27 },
      c: { x: 13, y: 300 },
      d: { x: 400, y: 300 },
    }
    const out = alignLayout(ids, pos, sizes(4))
    expect(Object.keys(out).sort()).toEqual(ids.slice().sort())
    ids.forEach((id) => expect(out[id], `缺少 ${id}`).toBeDefined())
  })

  it('每个节点都落在网格上（这是「对齐」最直接的含义）', () => {
    const ids = ['a', 'b', 'c']
    const pos: Record<string, NodePos> = {
      a: { x: 13, y: 27 },
      b: { x: 137, y: 149 },
      c: { x: 1003, y: 997 },
    }
    const out = alignLayout(ids, pos, sizes(3))
    for (const id of ids) {
      // ★ Math.abs 判倍数：负零问题（JS 里 -320 % 20 === -0，严格 === 0 为 false）
      expect(Math.abs(out[id].x) % ALIGN_GRID, `${id}.x 未对齐`).toBe(0)
      expect(Math.abs(out[id].y) % ALIGN_GRID, `${id}.y 未对齐`).toBe(0)
    }
  })

  it('★ 幂等：对齐两次的结果 === 对齐一次（两两互推不幂等，这是本算法的关键性质）', () => {
    // 三节点完全重叠：最考验互推算法的场景
    const ids = ['a', 'b', 'c']
    const pos: Record<string, NodePos> = { a: { x: 100, y: 100 }, b: { x: 100, y: 100 }, c: { x: 100, y: 100 } }
    const sz = sizes(3)
    const once = alignLayout(ids, pos, sz)
    const twice = alignLayout(ids, once, sz)
    expect(twice).toEqual(once)
    // 三节点不能还在同一个点上
    expect(new Set(ids.map((id) => `${once[id].x},${once[id].y}`)).size).toBe(3)
  })

  it('★ 幂等（普通场景：散乱摆放，含多行）', () => {
    const ids = ['a', 'b', 'c', 'd', 'e']
    const pos: Record<string, NodePos> = {
      a: { x: 7, y: 11 },
      b: { x: 933, y: 47 },
      c: { x: 251, y: 203 },
      d: { x: 611, y: 777 },
      e: { x: 88, y: 512 },
    }
    const sz = sizes(5)
    const once = alignLayout(ids, pos, sz)
    expect(alignLayout(ids, once, sz)).toEqual(once)
  })

  it('同行节点间距 ≥ GAP_X（消除重叠）', () => {
    // 两个节点宽 200、同 y、x 分别为 0 和 10（严重重叠）
    const ids = ['a', 'b']
    const pos: Record<string, NodePos> = { a: { x: 0, y: 100 }, b: { x: 10, y: 100 } }
    const out = alignLayout(ids, pos, { a: S(200, 90), b: S(200, 90) })
    const gap = out.b.x - (out.a.x + 200)
    expect(gap).toBeGreaterThanOrEqual(ALIGN_GAP_X)
    // 区间不重叠
    expect(out.b.x).toBeGreaterThanOrEqual(out.a.x + 200)
  })

  it('★ 同行节点的 y 被拉齐（上边缘对齐 = 用户看得见的「对齐」）', () => {
    const ids = ['a', 'b']
    // a 在 y=100，b 在 y=150（区间相交 ⇒ 同行）
    const pos: Record<string, NodePos> = { a: { x: 0, y: 100 }, b: { x: 500, y: 150 } }
    const out = alignLayout(ids, pos, { a: S(200, 90), b: S(200, 90) })
    expect(out.a.y).toBe(out.b.y)
    expect(out.a.y).toBe(100) // 取 x 最小者的 y
  })

  it('不同行的节点保持各自y（★ 不重排行间距 —— 那是「自动排列」的活）', () => {
    const ids = ['a', 'b']
    // a: 100~190，b: 500~590 → 两行
    const pos: Record<string, NodePos> = { a: { x: 0, y: 100 }, b: { x: 0, y: 500 } }
    const out = alignLayout(ids, pos, { a: S(200, 90), b: S(200, 90) })
    expect(out.a.y).toBe(100)
    expect(out.b.y).toBe(500)
    // 行间距保持 400，而不是被压成固定值
    expect(out.b.y - out.a.y).toBe(400)
  })

  it('★ 不改变 x 的相对次序（左边的仍在左边）—— 保持相对关系才是「对齐」', () => {
    const ids = ['a', 'b', 'c']
    const pos: Record<string, NodePos> = {
      a: { x: 900, y: 100 },
      b: { x: 300, y: 100 },
      c: { x: 600, y: 100 },
    }
    const out = alignLayout(ids, pos, sizes(3))
    const byId = (id: string) => out[id].x
    expect(byId('b')).toBeLessThan(byId('c'))
    expect(byId('c')).toBeLessThan(byId('a'))
  })

  it('缺尺寸 → 仍不重叠（保守默认生效）', () => {
    const ids = ['a', 'b', 'c']
    const pos: Record<string, NodePos> = {
      a: { x: 0, y: 100 },
      b: { x: 0, y: 100 },
      c: { x: 0, y: 100 },
    }
    const out = alignLayout(ids, pos, {}) // 不传尺寸
    expect(new Set(ids.map((id) => out[id].x)).size).toBe(3)
    // 按默认宽 220 算，相邻间距≥ 220 + GAP_X
    expect(out.b.x - out.a.x).toBeGreaterThanOrEqual(ALIGN_FALLBACK_W + ALIGN_GAP_X)
  })

  it('圆形节点（窄）与常规节点（宽）混合同行 → 间距按各自实际宽度算', () => {
    const ids = ['round', 'wide']
    const pos: Record<string, NodePos> = { round: { x: 0, y: 100 }, wide: { x: 5, y: 100 } }
    const out = alignLayout(ids, pos, { round: S(64, 64), wide: S(320, 120) })
    // 第二个的落点 = 第一个的落点 + 64 + GAP_X = 104，再吸附到网格 → 100。
    // ★ 断言「不重叠」而非精确等式：吸附必然让实际间距在 [GAP_X, GAP_X+grid) 之间浮动，
    //   写死 104 是在断言实现细节，下一次改 grid 或 GAP 就会假红。
    expect(out.wide.x).toBeGreaterThanOrEqual(out.round.x + 64)
    expect(out.wide.x).toBeLessThan(out.round.x + 64 + ALIGN_GAP_X + ALIGN_GRID)
    expect(out.wide.x % ALIGN_GRID).toBe(0)
  })

  it('positions 表里有但 ids 没传的节点 → 不出现在结果里（不回退默认位）', () => {
    const pos: Record<string, NodePos> = { a: { x: 0, y: 0 }, ghost: { x: 999, y: 999 } }
    const out = alignLayout(['a'], pos, {})
    expect(Object.keys(out)).toEqual(['a'])
  })

  it('ids 里某个 id 在 pos 里缺失 → 按 (0,0) 处理，不崩、不丢键', () => {
    const out = alignLayout(['ghost'], {}, {})
    expect(Object.keys(out)).toEqual(['ghost'])
    expect(out.ghost).toEqual({ x: 0, y: 0 })
  })

  it('单个节点：只吸附，不位移到别处', () => {
    // x=143 → 吸附到 140（最近网格点）
    const out = alignLayout(['a'], { a: { x: 143, y: 97 } }, { a: S(200, 90) })
    expect(out.a).toEqual({ x: 140, y: 100 })
  })

  /*★ 本组是本轮真bug 的护栏（浏览器验收抓出，纯函数测试曾全绿）：
     节点实测宽是 182 / 283 / 202 这类**非网格倍数**。
     若只吸附行的起点、之后一路 `x += w + GAP_X` 递推，
     w mod grid 的余数会累积，走到第 3~6 个节点就偏出网格
     （实测 182 宽时x 变成 262 / 282 / 484）。
     修复：每放一个节点都再吸附一次。 */
  describe('★ 节点宽/高不是网格倍数时，每一步都必须在网格上', () => {
    /*★★ 这组测试自己踩过一次「假绿」，务必先读：
       第一版把节点放在不同 y（间隔 200）⇒ 每个节点自成一行 ⇒ **递推从未发生**
       （每行cursor 只走一次，起点已吸附，看起来就永远合规）。
       实测：把「每步吸附」注掉，这5 条全绿 —— 0 条转红 = 护栏无效。
       ⇒ 触发递推的唯一条件是**同行的两个以上节点**，y 必须区间相交。 */
    it('宽度 182（非 20 倍数）× 6 个节点**同行** → 全部落在网格上', () => {
      const ids = ['n0', 'n1', 'n2', 'n3', 'n4', 'n5']
      // ★ 高度 57、y 全部 = 0 ⇒ 区间完全相交 ⇒ 同一行 ⇒ 递推必然发生
      const pos: Record<string, NodePos> = Object.fromEntries(
        ids.map((id) => [id, { x: 40, y: 0 }]),
      )
      const sz: Record<string, AlignSize> = Object.fromEntries(ids.map((id) => [id, S(182, 57)]))
      const out = alignLayout(ids, pos, sz)
      // 先自证「确实排成了一行」（否则本条退化成假绿）
      expect(new Set(ids.map((id) => out[id].y)).size).toBe(1)
      expect(new Set(ids.map((id) => out[id].x)).size).toBe(ids.length)
      for (const id of ids) {
        expect(Math.abs(out[id].x) % ALIGN_GRID, `${id} 的 x=${out[id].x} 不在网格上`).toBe(0)
      }
    })

    it('★ 混合真实宽度 182/283/202/68/349 **同行** → 全部落在网格上', () => {
      const ids = ['a', 'b', 'c', 'd', 'e']
      const ws = [182, 283, 202, 68, 349]
      const pos: Record<string, NodePos> = Object.fromEntries(ids.map((id) => [id, { x: 40, y: 0 }]))
      const sz: Record<string, AlignSize> = Object.fromEntries(
        ids.map((id, i) => [id, S(ws[i], 57)]),
      )
      const out = alignLayout(ids, pos, sz)
      expect(new Set(ids.map((id) => out[id].y)).size).toBe(1)
      for (const id of ids) {
        expect(
          Math.abs(out[id].x) % ALIGN_GRID,
          `${id} 的 x=${out[id].x}（宽 ${ws[ids.indexOf(id)]}）不在网格上`,
        ).toBe(0)
      }
    })

    it('★ 吸附后仍不重叠（吸附只会放大间距，不会造成重叠）', () => {
      const ids = ['a', 'b', 'c']
      // ★ y 区间必须相交（高度 57）才是同一行 —— 写成 y 间隔 200 会分成三行、
      //   各自保留原 x，那是正确行为而非缺陷，用跨行输入断言「同行不重叠」
      //   就是断言越界（同 skill 9.42）。
      const pos: Record<string, NodePos> = {
        a: { x: 0, y: 0 },
        b: { x: 0, y: 5 },
        c: { x: 0, y: 10 },
      }
      const out = alignLayout(ids, pos, { a: S(182, 57), b: S(283, 57), c: S(68, 68) })
      expect(out.b.x).toBeGreaterThanOrEqual(out.a.x + 182)
      expect(out.c.x).toBeGreaterThanOrEqual(out.b.x + 283)
      // 三者确实在同一行（y 相同）
      expect(new Set(ids.map((id) => out[id].y)).size).toBe(1)
    })

    it('不同行各自保留原 x（跨行不该互相推 —— 这是「保持相对关系」）', () => {
      const ids = ['a', 'b', 'c']
      // y 间隔 200 > 高度 57 ⇒ 分成三行
      const pos: Record<string, NodePos> = {
        a: { x: 40, y: 0 },
        b: { x: 260, y: 200 },
        c: { x: 140, y: 400 },
      }
      const out = alignLayout(ids, pos, { a: S(182, 57), b: S(283, 57), c: S(68, 68) })
      expect(out.a.x).toBe(40)
      expect(out.b.x).toBe(260)
      expect(out.c.x).toBe(140)
    })

    it('★ 负坐标输入（用户拖到画布左上角外）也在网格上', () => {
      const ids = ['a', 'b', 'c']
      const pos: Record<string, NodePos> = {
        a: { x: -316.663, y: 235.922 },
        b: { x: -164, y: 435.922 },
        c: { x: 145, y: 635.922 },
      }
      const out = alignLayout(ids, pos, { a: S(182, 57), b: S(182, 57), c: S(202, 62) })
      for (const id of ids) {
        // ★ 用 Math.abs 判倍数：JS 里 -320 % 20 === -0（严格 ===0 为 false）
        //   —— 负坐标明明在网格上却报红，是断言写法问题不是实现问题。
        expect(Math.abs(out[id].x) % ALIGN_GRID, `${id} 的 x=${out[id].x}`).toBe(0)
      }
    })
  })
})

describe('★ 与「自动排列」layoutLayers 的语义隔离', () => {
  /** 造一条 3 节点链路：start → material → bill */
  function chain(): Record<string, TableSchema> {
    const mk = (id: string, name: string, inputs?: string[]): TableSchema =>
      ({
        id,
        name,
        kind: 'table',
        fields: [],
        config: inputs ? { inputs } : {},
      }) as unknown as TableSchema
    return {
      s: mk('s', '开始'),
      m: mk('m', '材料', ['s']),
      b: mk('b', '工程量', ['m']),
    }
  }

  it('对齐**不会**把链路分层重排（那是自动排列的语义）', () => {
    // 三节点摆成参差的样子但都在同一行附近
    const ids = ['s', 'm', 'b']
    const pos: Record<string, NodePos> = {
      s: { x: 11, y: 13 },
      m: { x: 15, y: 17 },
      b: { x: 19, y: 23 },
    }
    const out = alignLayout(ids, pos, sizes(3))
    // 三个 y 都在吸附后落进同一区间 ⇒ 只有一行（若被分层，就会出现 3 个不同 y）
    expect(new Set(ids.map((id) => out[id].y)).size).toBe(1)
    // 而 layoutLayers 会把它们分成 3 层（3 个不同 x）
    const laid = layoutLayers(chain())
    expect(new Set(['s', 'm', 'b'].map((id) => laid[id].x)).size).toBe(3)
  })

  it('★ 对齐保留手拖的横向次序，而自动排列按拓扑改次序', () => {
    const ids = ['s', 'm', 'b']
    // 故意把「下游」b 放在最左边
    const pos: Record<string, NodePos> = {
      s: { x: 0, y: 0 },
      m: { x: 400, y: 0 },
      b: { x: 800, y: 0 },
    }
    const aligned = alignLayout(ids, pos, sizes(3))
    // 对齐后 b 仍在最右（次序未变）
    expect(aligned.b.x).toBeGreaterThan(aligned.m.x)
    expect(aligned.m.x).toBeGreaterThan(aligned.s.x)
    // 自动排列：s 在最左（拓扑源头），方向与手拖一致时结论相同
    const laid = layoutLayers(chain())
    expect(laid.s.x).toBeLessThan(laid.m.x)
    expect(laid.m.x).toBeLessThan(laid.b.x)
  })

  it('★ 两个函数对同一输入给出不同结果（证明确实是两个按钮，不是一个）', () => {
    const ids = ['s', 'm', 'b']
    const pos: Record<string, NodePos> = {
      s: { x: 500, y: 0 },
      m: { x: 100, y: 0 },
      b: { x: 300, y: 0 },
    }
    const aligned = alignLayout(ids, pos, sizes(3))
    const laid = layoutLayers(chain())
    // 对齐：m 在最左（尊重手拖次序）
    expect(aligned.m.x).toBeLessThan(aligned.b.x)
    // 自动排列：s 在最左（尊重拓扑）
    expect(laid.s.x).toBeLessThan(laid.m.x)
  })
})