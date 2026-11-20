import { describe, it, expect } from 'vitest'
import { RefGraph } from './refGraph'

/** 搭一张典型的出图流程：计算 → 工程量 → 三份产出 */
function buildGraph(): RefGraph {
  const g = new RefGraph()
  g.addNode({
    id: 'calc',
    title: '水力计算',
    outputs: [{ key: 'Q' }, { key: 'dP_kPa' }],
    run: () => ({ Q: 1200, dP_kPa: 3.24 }),
  })
  g.addNode({
    id: 'boq',
    title: '工程量表',
    slots: [{ nodeId: 'boq', key: '流量' }],
    outputs: [{ key: 'rows' }, { key: '合价' }],
    run: (i) => {
      const rows = [
        { 名称: '钢管 DN200', 数量: Math.round(Number(i['流量']) / 12), 单价: 25.5 },
        { 名称: '阀门 DN50', 数量: 12, 单价: 260 },
      ]
      return { rows, 合价: rows.reduce((s, r) => s + r.数量 * r.单价, 0) }
    },
  })
  g.addNode({ id: 'docx', title: '设计说明.docx', slots: [{ nodeId: 'docx', key: '设计流量' }] })
  g.addNode({ id: 'xlsx', title: '报表.xlsx', slots: [{ nodeId: 'xlsx', key: '设计流量' }] })
  g.addNode({ id: 'chart', slots: [{ nodeId: 'chart', key: 'series' }] })

  g.bind('boq', '流量', 'node://calc#Q')
  g.bind('docx', '设计流量', 'node://calc#Q')
  g.bind('xlsx', '设计流量', 'node://calc#Q')
  g.bind('chart', 'series', 'node://boq#rows')
  return g
}

describe('引用即依赖', () => {
  it('依赖图从引用自动推导，无需额外声明', () => {
    const g = buildGraph()
    const deps = g.deps()
    expect([...deps.get('calc')!]).toEqual([])
    expect([...deps.get('boq')!]).toEqual(['calc'])
    expect([...deps.get('chart')!]).toEqual(['boq'])
  })

  it('拓扑排序把上游排在前面', () => {
    const order = buildGraph().order()
    expect(order.indexOf('calc')).toBeLessThan(order.indexOf('boq'))
    expect(order.indexOf('boq')).toBeLessThan(order.indexOf('chart'))
    expect(order.indexOf('calc')).toBeLessThan(order.indexOf('docx'))
  })

  it('绑到外部 scheme 不产生节点依赖', () => {
    const g = buildGraph()
    g.bind('docx', '审查单位', 'param://审查单位')
    expect([...g.deps().get('docx')!].sort()).toEqual(['calc'])
  })
})

describe('跨节点取值', () => {
  it('下游拿到上游 run 产出的输出', () => {
    const r = buildGraph().resolveAll()
    expect(r.values.get('boq')!['流量']).toBe(1200)
    expect(r.values.get('docx')!['设计流量']).toBe(1200)
    expect(r.values.get('xlsx')!['设计流量']).toBe(1200)
    expect((r.values.get('chart')!['series'] as unknown[]).length).toBe(2)
  })

  it('一个源被多个下游引用，值必然一致', () => {
    const g = buildGraph()
    g.addNode({ id: 'docx2', slots: [{ nodeId: 'docx2', key: 'Q' }] })
    g.bind('docx2', 'Q', 'node://calc#Q')
    const r = g.resolveAll()
    expect(r.values.get('docx')!['设计流量']).toBe(r.values.get('xlsx')!['设计流量'])
    expect(r.values.get('docx2')!['Q']).toBe(1200)
  })

  it('多层链式引用逐级传导', () => {
    const r = buildGraph().resolveAll()
    // calc.Q=1200 → boq 由它算 rows → chart 拿 rows
    expect(r.outputs.get('boq')!['合价']).toBe(100 * 25.5 + 12 * 260)
  })

  it('外部上下文（param / file / api）也能被槽位引用', () => {
    const g = new RefGraph()
    g.addNode({ id: 'doc', slots: [{ nodeId: 'doc', key: '单价' }, { nodeId: 'doc', key: '名称' }] })
    g.bind('doc', '单价', 'file://D:/价目.xlsx#价目表!B2')
    g.bind('doc', '名称', 'param://工程名称')
    const r = g.resolveAll({
      files: { 'D:/价目.xlsx': { 价目表: { B2: 25.5 } } },
      params: { 工程名称: '海南陵水燃气工程' },
    })
    expect(r.values.get('doc')!['单价']).toBe(25.5)
    expect(r.values.get('doc')!['名称']).toBe('海南陵水燃气工程')
  })
})

describe('溯源与错误处理', () => {
  it('每条引用都有出处记录', () => {
    const r = buildGraph().resolveAll()
    const t = r.traces.find((x) => x.nodeId === 'docx' && x.slot === '设计流量')!
    expect(t.urn).toBe('node://calc#Q')
    expect(t.from).toBe('node')
    expect(t.ok).toBe(true)
  })

  it('引用解析不到值时标 ok=false 并给原因，而不是静默留空', () => {
    const g = new RefGraph()
    g.addNode({ id: 'doc', slots: [{ nodeId: 'doc', key: '缺的' }] })
    g.bind('doc', '缺的', 'node://不存在的节点#x')
    const t = g.resolveAll().traces[0]
    expect(t.value).toBeUndefined()
    expect(t.ok).toBe(false)
    expect(t.error).toBeTruthy()
  })
})

describe('环检测', () => {
  it('findCycle 能揪出环路径', () => {
    const g = new RefGraph()
    g.addNode({ id: 'A' }); g.addNode({ id: 'B' }); g.addNode({ id: 'C' })
    g.bind('A', 'x', 'node://B#y')
    g.bind('B', 'y', 'node://C#z')
    g.bind('C', 'z', 'node://A#x')
    const c = g.findCycle()
    expect(c).not.toBeNull()
    expect(c!.length).toBeGreaterThanOrEqual(3)
  })

  it('有环时 order() 直接抛错，不给出错误的执行序', () => {
    const g = new RefGraph()
    g.addNode({ id: 'A' }); g.addNode({ id: 'B' })
    g.bind('A', 'x', 'node://B#y')
    g.bind('B', 'y', 'node://A#x')
    expect(() => g.order()).toThrow(/成环/)
  })

  it('无环图不误报', () => {
    expect(buildGraph().findCycle()).toBeNull()
  })
})

describe('绑定管理', () => {
  it('同一槽位重复绑定会覆盖而不是叠加', () => {
    const g = new RefGraph()
    g.addNode({ id: 'doc', slots: [{ nodeId: 'doc', key: 'v' }] })
    g.bind('doc', 'v', 'param://a')
    g.bind('doc', 'v', 'param://b')
    expect(g.bindingsOf('doc')).toHaveLength(1)
    const r = g.resolveAll({ params: { a: 1, b: 2 } })
    expect(r.values.get('doc')!['v']).toBe(2)
  })

  it('unbind 解绑后不再解析', () => {
    const g = buildGraph()
    g.unbind('docx', '设计流量')
    expect(g.bindingsOf('docx')).toHaveLength(0)
    expect([...g.deps().get('docx')!]).toEqual([])
  })
})
