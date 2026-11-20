import { describe, it, expect } from 'vitest'
import { defineField, defineTable } from './schema'
import { createEngine, calcTable, setGlobalParam } from './engine'

describe('全局参数', () => {
  it('[@@TAX_RATE] 命名表达式参与公式', () => {
    const t = defineTable('cost', [
      defineField('price', 'number'),
      defineField('tax', 'formula', { formula: '=[@price]*[@@TAX_RATE]' }),
    ])
    const hf = createEngine()
    setGlobalParam(hf, 'TAX_RATE', 0.1)
    const out = calcTable(hf, t, [{ price: 1000 }, { price: 200 }])
    expect(out[0].tax).toBeCloseTo(100, 5)
    expect(out[1].tax).toBeCloseTo(20, 5)
  })
})
