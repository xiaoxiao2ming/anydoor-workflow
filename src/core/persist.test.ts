import { describe, it, expect } from 'vitest'
import { createRequire } from 'module'
import initSqlJs, { type Database } from 'sql.js'
import { writeState, readState, ensureSchema } from './db'
import { defineField, defineTable } from './schema'

const require = createRequire(import.meta.url)
type SqlJsStatic = Awaited<ReturnType<typeof initSqlJs>>
let SQL: SqlJsStatic | null = null

async function ensureSql(): Promise<SqlJsStatic> {
  if (!SQL) {
    const wasmPath = require.resolve('sql.js/dist/sql-wasm.wasm').replace(/\\/g, '/')
    SQL = await initSqlJs({ locateFile: () => wasmPath })
  }
  return SQL
}

describe('sql.js 持久化', () => {
  it('写状态 → 导出字节 → 重建 → 读回，数据不丢', async () => {
    const cost = defineTable('cost', [
      defineField('name', 'text'),
      defineField('qty', 'number'),
      defineField('price', 'number'),
    ])
    const state = {
      tables: { cost },
      records: {
        cost: [
          { name: '挖沟槽土方', qty: 120, price: 35 },
          { name: 'C30 混凝土', qty: 45, price: 480 },
        ],
      },
      positions: { cost: { x: 10, y: 20 } },
    }

    const Static = await ensureSql()
    const db = new Static.Database()
    ensureSchema(db)
    writeState(db, state)

    // 模拟"落盘 → 重启 → hydrate"
    const bytes = db.export()
    const db2 = new Static.Database(bytes)
    const reloaded = readState(db2)

    expect(reloaded.records.cost.length).toBe(2)
    expect(reloaded.records.cost[0].qty).toBe(120)
    expect(reloaded.tables.cost.fields.length).toBe(3)
    // 位置必须随落盘往返，否则刷新节点回弹
    expect(reloaded.positions.cost).toEqual({ x: 10, y: 20 })
  })

  it('连线(lookup 字段)+位置+记录在 落盘重建后都还在', async () => {
    // 模拟浏览器真实动作：建两张表 → 拖出一条关联线(lookup) → 摆位置 → 填数据
    const A = defineTable('A', [defineField('id', 'text')], 'tbl_a')
    const B = defineTable('B', [defineField('id', 'text')], 'tbl_b')
    // 在 A 上加一个指向 B 的 lookup 字段 = 画布上一条连线
    A.fields.push(
      defineField('B引用', 'lookup', { lookup: { table: 'tbl_b', sourceField: 'id', matchField: 'id', returnField: 'id' } }),
    )

    const state = {
      tables: { tbl_a: A, tbl_b: B },
      records: {
        tbl_a: [{ id: 'a1' }, { id: 'a2' }],
        tbl_b: [{ id: 'b1' }],
      },
      positions: { tbl_a: { x: 120, y: 80 }, tbl_b: { x: 520, y: 300 } },
    }

    const Static = await ensureSql()
    const db = new Static.Database()
    ensureSchema(db)
    writeState(db, state)

    const bytes = db.export()
    const db2 = new Static.Database(bytes)
    const reloaded = readState(db2)

    // 1) 节点数据
    expect(reloaded.tables.tbl_a.fields.some((f) => f.id === 'id')).toBe(true)
    expect(reloaded.records.tbl_a.length).toBe(2)
    // 2) 连线：lookup 字段随表 schema 落盘，重建后连线才能还原
    const reloadedLookup = reloaded.tables.tbl_a.fields.find((f) => f.type === 'lookup')
    expect(reloadedLookup?.lookup?.table).toBe('tbl_b')
    // 3) 位置
    expect(reloaded.positions.tbl_a).toEqual({ x: 120, y: 80 })
    expect(reloaded.positions.tbl_b).toEqual({ x: 520, y: 300 })
  })
})
