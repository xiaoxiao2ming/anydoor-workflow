import { describe, it, expect } from 'vitest'
import { defineField, defineTable } from './schema'
import { createEngine, calcTable, resolveLookups } from './engine'

describe('关联 lookup', () => {
  it('从价目表按材料名查单价，再参与公式计算', () => {
    const priceList = defineTable('priceList', [
      defineField('material', 'text'),
      defineField('unitPrice', 'number'),
    ])
    const bill = defineTable('bill', [
      defineField('material', 'text'),
      defineField('qty', 'number'),
      defineField('unitPrice', 'lookup', {
        lookup: { table: 'priceList', sourceField: 'material', matchField: 'material', returnField: 'unitPrice' },
      }),
      defineField('total', 'formula', { formula: '=[@qty]*[@unitPrice]' }),
    ])
    const tables = { priceList, bill }
    const records = {
      priceList: [
        { material: '钢筋', unitPrice: 4200 },
        { material: '水泥', unitPrice: 480 },
      ],
      bill: [
        { material: '钢筋', qty: 8 },
        { material: '水泥', qty: 45 },
      ],
    }

    const resolved = resolveLookups(tables, records)
    expect(resolved.bill[0].unitPrice).toBe(4200)
    expect(resolved.bill[1].unitPrice).toBe(480)

    const out = calcTable(createEngine(), bill, resolved.bill)
    expect(out[0].total).toBe(33600) // 8 * 4200
    expect(out[1].total).toBe(21600) // 45 * 480
  })

  it('查不到匹配时返回 undefined', () => {
    const priceList = defineTable('priceList', [
      defineField('material', 'text'),
      defineField('unitPrice', 'number'),
    ])
    const bill = defineTable('bill', [
      defineField('material', 'text'),
      defineField('unitPrice', 'lookup', {
        lookup: { table: 'priceList', sourceField: 'material', matchField: 'material', returnField: 'unitPrice' },
      }),
    ])
    const resolved = resolveLookups(
      { priceList, bill },
      {
        priceList: [{ material: '钢筋', unitPrice: 4200 }],
        bill: [{ material: '未知材料' }],
      },
    )
    expect(resolved.bill[0].unitPrice).toBeUndefined()
  })
})
