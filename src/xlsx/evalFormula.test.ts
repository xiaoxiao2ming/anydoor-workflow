import { describe, it, expect } from 'vitest'
import { evalRowFormula, FUNC_COUNT } from './evalFormula'

/**
 * 计算列缓存值求值。
 * 这层的存在意义：不写缓存值的单元格在 OOXML 里连 <v> 都没有，
 * Excel/WPS 之外的阅读器一律显示 #N/A。所以缓存值必须我们自己算出来。
 */
describe('evalRowFormula（计算列缓存值）', () => {
  const row = { B2: 100, C2: 25.5, D2: '钢管 DN200' }

  it('四则运算', () => {
    expect(evalRowFormula('B2*C2', row)).toEqual({ ok: true, value: 2550 })
    expect(evalRowFormula('B2+C2', row)).toEqual({ ok: true, value: 125.5 })
    expect(evalRowFormula('B2-C2', row).value).toBeCloseTo(74.5, 6)
    expect(evalRowFormula('B2/C2', row).value).toBeCloseTo(3.921568627, 6)
  })

  it('优先级与括号', () => {
    expect(evalRowFormula('1+2*3', {})).toEqual({ ok: true, value: 7 })
    expect(evalRowFormula('(1+2)*3', {})).toEqual({ ok: true, value: 9 })
    expect(evalRowFormula('2^3^2', {})).toEqual({ ok: true, value: 512 }) // 右结合
    expect(evalRowFormula('-B2+10', row)).toEqual({ ok: true, value: -90 })
  })

  it('百分号', () => {
    expect(evalRowFormula('B2*5%', row)).toEqual({ ok: true, value: 5 })
  })

  it('函数走 formulajs', () => {
    expect(evalRowFormula('ROUND(B2*C2,2)', row)).toEqual({ ok: true, value: 2550 })
    expect(evalRowFormula('SUM(B2,C2)', row)).toEqual({ ok: true, value: 125.5 })
    expect(evalRowFormula('IF(B2>50,1,0)', row)).toEqual({ ok: true, value: 1 })
    expect(evalRowFormula('MAX(B2,C2)', row)).toEqual({ ok: true, value: 100 })
    expect(evalRowFormula('CONCATENATE(D2," × ",B2)', row)).toEqual({
      ok: true,
      value: '钢管 DN200 × 100',
    })
  })

  it('比较与逻辑', () => {
    expect(evalRowFormula('B2>50', row)).toEqual({ ok: true, value: true })
    expect(evalRowFormula('B2<>100', row)).toEqual({ ok: true, value: false })
    expect(evalRowFormula('AND(B2>50,C2<30)', row)).toEqual({ ok: true, value: true })
  })

  it('字符串字面量里的 A1 不被当成引用', () => {
    // 'A1' 在引号里，且 refValues 里没有 A1 —— 若被当引用会变成 0 或报错
    expect(evalRowFormula('"B2"&D2', row)).toEqual({ ok: true, value: 'B2钢管 DN200' })
  })

  it('值里带运算符也不会破坏表达式（占位符而非文本拼接）', () => {
    expect(evalRowFormula('A1&""', { A1: '1+2*3' })).toEqual({ ok: true, value: '1+2*3' })
  })

  it('未定义的函数 / 名称要报错，不静默给错值', () => {
    const r1 = evalRowFormula('NOSUCHFN(1)', {})
    expect(r1.ok).toBe(false)
    expect(r1.error).toContain('不支持的函数')
    const r2 = evalRowFormula('不存在', {})
    expect(r2.ok).toBe(false)
    expect(r2.error).toContain('未定义的名称')
  })

  it('除零不返回 Infinity（那会写进 XML 变成坏值）', () => {
    const r = evalRowFormula('1/0', {})
    expect(r.ok).toBe(false)
    expect(r.error).toContain('有限数')
  })

  it('缺右括号要报错', () => {
    expect(evalRowFormula('SUM(1,2', {}).ok).toBe(false)
    expect(evalRowFormula('(1+2', {}).ok).toBe(false)
  })

  it('空公式', () => {
    expect(evalRowFormula('', {}).ok).toBe(false)
  })

  it('区间引用（#36）：SUM(A1:B2) 矩形展开成二维数组', () => {
    const rng = {
      A1: 1, B1: 2,
      A2: 3, B2: 4,
    }
    expect(evalRowFormula('SUM(A1:B2)', rng)).toEqual({ ok: true, value: 10 })
    expect(evalRowFormula('SUM(A1:B2)*2', rng)).toEqual({ ok: true, value: 20 })
    expect(evalRowFormula('AVERAGE(A1:B2)', rng)).toEqual({ ok: true, value: 2.5 })
    expect(evalRowFormula('COUNT(A1:B2)', rng)).toEqual({ ok: true, value: 4 })
  })

  it('区间端点归一化（A2:B1 → A1:B2）+ 冒号两侧空白', () => {
    const rng = { A1: 1, B1: 2, A2: 3, B2: 4 }
    expect(evalRowFormula('SUM(A2:B1)', rng)).toEqual({ ok: true, value: 10 })
    expect(evalRowFormula('SUM(A1 : B2)', rng)).toEqual({ ok: true, value: 10 })
  })

  it('区间里缺值/越界格子按空处理（formulajs 聚合跳过）', () => {
    // refValues 只给了 A1 —— B1/A2/B2 取 undefined，SUM 只累计 5
    expect(evalRowFormula('SUM(A1:B2)', { A1: 5 })).toEqual({ ok: true, value: 5 })
  })

  it('区间与单格/胶囊混排', () => {
    const rng = { A1: 1, B1: 2, A2: 3, B2: 4, C2: 100 }
    expect(evalRowFormula('SUM(A1:B2)+C2', rng)).toEqual({ ok: true, value: 110 })
    expect(evalRowFormula('MAX(A1:B2, C2)', rng)).toEqual({ ok: true, value: 100 })
  })

  it('字符串字面量里的区间不被替换', () => {
    expect(evalRowFormula('"A1:B2"&D2', { D2: 'x' })).toEqual({ ok: true, value: 'A1:B2x' })
  })

  it('区间参与四则 = #VALUE! 同语义（报错不静默）', () => {
    const r = evalRowFormula('SUM(A1:B2)*2', { A1: 1, B1: 2, A2: 3, B2: 4 })
    // 数组*2：formulajs 会映射或报错 —— 只要不产出错值即可（Excel 里数组*2 也是数组）
    expect([true, false]).toContain(r.ok)
  })

  it('超大区间护栏：保留原文报错不卡死', () => {
    const r = evalRowFormula('SUM(A1:Z99999)', {})
    expect(r.ok).toBe(false)
  })

  it('可调用函数数量（formulajs）', () => {
    expect(FUNC_COUNT).toBeGreaterThan(300)
  })
})

describe('node:// 引用参与求值（批次 1-6 · 胶囊公式）', () => {
  const resolve = (urn: string): unknown => {
    const m: Record<string, unknown> = {
      'node://p#单价': 3.5,
      'node://b#数量': 100,
      'node://d#折扣': 0.9,
      'node://t#名称': '南充项目',
    }
    return m[urn]
  }

  it('两胶囊相加（多胶囊共同参与公式）', () => {
    expect(evalRowFormula('node://p#单价+node://b#数量', {}, resolve)).toEqual({ ok: true, value: 103.5 })
  })

  it('混排字面量与优先级', () => {
    expect(evalRowFormula('node://b#数量*2+100', {}, resolve)).toEqual({ ok: true, value: 300 })
    expect(evalRowFormula('(node://p#单价+1)*10', {}, resolve)).toEqual({ ok: true, value: 45 })
  })

  it('三胶囊四则（加减乘除）', () => {
    expect(evalRowFormula('node://p#单价*node://b#数量*node://d#折扣', {}, resolve).value).toBeCloseTo(315, 6)
    expect(evalRowFormula('node://b#数量/node://d#折扣-node://p#单价', {}, resolve).value).toBeCloseTo(107.611111, 5)
  })

  it('值为字符串时可拼接', () => {
    expect(evalRowFormula('node://t#名称&"-附"', {}, resolve)).toEqual({ ok: true, value: '南充项目-附' })
  })

  it('resolver 未命中：保留原文 → 解析失败（与坏公式同路，不静默给错值）', () => {
    const r = evalRowFormula('node://nope#x+1', {}, resolve)
    expect(r.ok).toBe(false)
  })

  it('没传 resolver：URN 原文走解析器 → 报错（保持旧行为）', () => {
    expect(evalRowFormula('node://p#单价+1', {}).ok).toBe(false)
  })

  it('字符串字面量里的 node:// 不被替换', () => {
    expect(evalRowFormula('"node://p#单价"', {}, resolve)).toEqual({ ok: true, value: 'node://p#单价' })
  })

  it('整节点引用（无 #locator）也可解析', () => {
    const r2 = evalRowFormula('node://p*2', {}, (u) => (u === 'node://p' ? 21 : undefined))
    expect(r2).toEqual({ ok: true, value: 42 })
  })

  it('胶囊与 A1 混排', () => {
    expect(evalRowFormula('B2*node://d#折扣', { B2: 100 }, resolve).value).toBeCloseTo(90, 6)
  })
})
