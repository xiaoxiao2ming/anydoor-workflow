import { describe, it, expect } from 'vitest'
import { defineField, defineTable } from './schema'
import { createEngine, calcTable, setGlobalParam } from './engine'

describe('公式字段', () => {
  it('[@qty]*[@price] 算出合价', () => {
    const t = defineTable('cost', [
      defineField('qty', 'number'),
      defineField('price', 'number'),
      defineField('total', 'formula', { formula: '=[@qty]*[@price]' }),
    ])
    const out = calcTable(createEngine(), t, [
      { qty: 120, price: 35 },
      { qty: 45, price: 480 },
    ])
    expect(out[0].total).toBe(4200)
    expect(out[1].total).toBe(21600)
  })

  /**
   * 下面这几条是「从 HyperFormula 换成自研求值器」时新增的行为断言。
   * 换引擎最容易踩的坑就是这些：换了之后没人再自动帮你定序，也没人再自动
   * 把字符串参数当命名表达式处理 —— 必须显式测出来。
   */

  it('公式列乱序声明时仍算对（依赖不靠声明顺序）', () => {
    // 注意：total 声明在 qty / price 之前，但它引用了它们。
    // 原 HyperFormula 靠依赖图自动定序；自研求值器靠「多轮迭代」达到同样效果。
    const t = defineTable('cost', [
      defineField('total', 'formula', { formula: '=[@qty]*[@price]' }),
      defineField('qty', 'number'),
      defineField('price', 'number'),
    ])
    const out = calcTable(createEngine(), t, [{ qty: 12, price: 5 }])
    expect(out[0].total).toBe(60)
  })

  it('公式列可以链式引用另一个公式列（含乱序）', () => {
    const t = defineTable('cost', [
      defineField('withTax', 'formula', { formula: '=[@subtotal]*1.13' }),
      defineField('subtotal', 'formula', { formula: '=[@qty]*[@price]' }),
      defineField('qty', 'number'),
      defineField('price', 'number'),
    ])
    const out = calcTable(createEngine(), t, [{ qty: 10, price: 100 }])
    expect(out[0].subtotal).toBe(1000)
    expect(out[0].withTax).toBeCloseTo(1130, 5)
  })

  it('引用未定义的全局参数 → #NAME?（不静默按 0 算）', () => {
    const t = defineTable('cost', [
      defineField('price', 'number'),
      defineField('tax', 'formula', { formula: '=[@price]*[@@NOT_SET]' }),
    ])
    const out = calcTable(createEngine(), t, [{ price: 1000 }])
    // 关键：绝不能是 0。0 是"看起来正常的错数"，比报错更难发现。
    expect(out[0].tax).toBe('#NAME?')
  })

  it('不支持的函数 → #ERROR', () => {
    const t = defineTable('cost', [
      defineField('price', 'number'),
      defineField('x', 'formula', { formula: '=NOSUCHFUNC([@price])' }),
    ])
    const out = calcTable(createEngine(), t, [{ price: 1 }])
    expect(out[0].x).toBe('#ERROR')
  })

  it('字符串全局参数按字面量内联，内部的引号不会破坏公式（注入防护）', () => {
    const t = defineTable('doc', [
      defineField('name', 'text'),
      defineField('title', 'formula', { formula: '=[@name]&[@@SUFFIX]' }),
    ])
    const eng = createEngine()
    // 值里带双引号：若直接字符串拼接，公式会被提前闭合而成语法错误
    setGlobalParam(eng, 'SUFFIX', '"市政"工程')
    const out = calcTable(eng, t, [{ name: 'A' }])
    expect(out[0].title).toBe('A"市政"工程')
  })

  it('引用不存在的字段 → 按 0 参与计算（与上游行为一致）', () => {
    const t = defineTable('cost', [
      defineField('price', 'number'),
      defineField('x', 'formula', { formula: '=[@price]*[@不存在]' }),
    ])
    const out = calcTable(createEngine(), t, [{ price: 7 }])
    expect(out[0].x).toBe(0)
  })
})
