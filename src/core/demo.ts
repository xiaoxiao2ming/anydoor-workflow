// 底座自检：定义一张成本表，跑两行数据，验证公式字段能算出合价
// 运行：npx vite-node src/core/demo.ts   （或后续接 Vitest）
import { defineField, defineTable } from './schema'
import { createEngine, calcTable } from './engine'

export function runDemo() {
  const cost = defineTable('cost', [
    defineField('name', 'text'),
    defineField('qty', 'number'),
    defineField('price', 'number'),
    defineField('total', 'formula', { formula: '=[@qty]*[@price]' }),
  ])

  const rows = [
    { name: '挖沟槽土方', qty: 120, price: 35 },
    { name: 'C30 混凝土', qty: 45, price: 480 },
  ]

  const hf = createEngine()
  const out = calcTable(hf, cost, rows)

  // eslint-disable-next-line no-console
  console.log('底座计算结果：')
  out.forEach((r, i) => {
    // eslint-disable-next-line no-console
    console.log(`  行${i + 1}: ${r.name} × ${r.qty} × ${r.price} = 合价 ${r.total}`)
  })
  return out
}

runDemo()
