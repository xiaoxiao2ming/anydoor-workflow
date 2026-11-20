import { defineField, defineTable } from './schema'
import type { Row } from './store'

// 材料库表（被引用）
export const MATERIALS = defineTable(
  '材料库',
  [
    defineField('name', 'text'),
    defineField('unit', 'text'),
    defineField('unitPrice', 'number', { format: 'currency' }),
  ],
  'materials',
)

// 用量清单表（通过 lookup 引用材料库的单价，total 用公式算）
export const BILL = defineTable(
  '用量清单',
  [
    defineField('name', 'text'),
    defineField('qty', 'number'),
    defineField('unitPrice', 'lookup', {
      lookup: { table: 'materials', sourceField: 'name', matchField: 'name', returnField: 'unitPrice' },
      format: 'currency',
    }),
    defineField('total', 'formula', { formula: '=[@qty]*[@unitPrice]', format: 'currency' }),
  ],
  'bill',
)

export const MATERIALS_SEED: Row[] = [
  { name: '钢筋', unit: 't', unitPrice: 4200 },
  { name: '水泥', unit: 't', unitPrice: 480 },
  { name: '模板', unit: 'm²', unitPrice: 65 },
]

export const BILL_SEED: Row[] = [
  { name: '钢筋', qty: 8 },
  { name: '水泥', qty: 45 },
]

// 开始节点：工作流的触发起点，可设置触发方式（手动/定时/Webhook）
export const START = defineTable(
  '开始',
  [],
  'start',
  'start',
  { trigger: 'manual' },
)
