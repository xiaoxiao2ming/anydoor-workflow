// sql.js 数据底座：加载 wasm、从持久化层 hydrate、读写结构化表
// 物理结构（真实 .sqlite 文件，可用 DB Browser for SQLite 打开）：
//   __tables__(id, name, schema_json)          —— 表元数据
//   __rows__(id, table_id, row_idx, data_json) —— 行数据（JSON，加字段无需迁移）
//   __positions__(table_id, x, y)              —— 画布节点位置
import initSqlJs, { type Database } from 'sql.js'
import type { TableSchema } from './schema'
import type { Row } from './store'
import { getAdapter } from './persistence'

export interface NodePos {
  x: number
  y: number
}

export interface SaveableState {
  tables: Record<string, TableSchema>
  records: Record<string, Row[]>
  positions: Record<string, NodePos>
}

type SqlJsStatic = Awaited<ReturnType<typeof initSqlJs>>
let sqlStatic: SqlJsStatic | null = null

async function getSqlStatic(): Promise<SqlJsStatic> {
  if (!sqlStatic) {
    // ?url 后缀让 vite 把 wasm 当静态资源返回 URL；动态 import 避免在 node 测试里炸
    const { default: init } = await import('sql.js')
    const { default: wasmUrl } = await import('sql.js/dist/sql-wasm.wasm?url')
    sqlStatic = await init({ locateFile: () => wasmUrl })
  }
  return sqlStatic
}

export function ensureSchema(db: Database) {
  db.run(`
    CREATE TABLE IF NOT EXISTS __tables__ (
      id TEXT PRIMARY KEY,
      name TEXT,
      schema_json TEXT
    );
    CREATE TABLE IF NOT EXISTS __rows__ (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      table_id TEXT,
      row_idx INTEGER,
      data_json TEXT
    );
    CREATE TABLE IF NOT EXISTS __positions__ (
      table_id TEXT PRIMARY KEY,
      x REAL,
      y REAL
    );
  `)
}

export function readState(db: Database): SaveableState {
  ensureSchema(db)
  const tables: Record<string, TableSchema> = {}
  const tRes = db.exec('SELECT id, schema_json FROM __tables__')
  if (tRes[0]) {
    for (const [id, schemaJson] of tRes[0].values as [string, string][]) {
      try {
        tables[id] = JSON.parse(schemaJson)
      } catch {
        /* 跳过损坏行 */
      }
    }
  }
  const records: Record<string, Row[]> = {}
  const rRes = db.exec('SELECT table_id, row_idx, data_json FROM __rows__ ORDER BY table_id, row_idx')
  if (rRes[0]) {
    for (const [tid, idx, dataJson] of rRes[0].values as [string, number, string][]) {
      if (!records[tid]) records[tid] = []
      try {
        records[tid][idx] = JSON.parse(dataJson)
      } catch {
        /* 跳过损坏行 */
      }
    }
  }
  const positions: Record<string, NodePos> = {}
  const pRes = db.exec('SELECT table_id, x, y FROM __positions__')
  if (pRes[0]) {
    for (const [tid, x, y] of pRes[0].values as [string, number, number][]) {
      positions[tid] = { x, y }
    }
  }
  return { tables, records, positions }
}

/** 全量覆盖写（MVP，数据量小；后续可改增量 upsert） */
export function writeState(db: Database, state: SaveableState) {
  ensureSchema(db)
  db.run('BEGIN')
  db.run('DELETE FROM __tables__')
  db.run('DELETE FROM __rows__')
  db.run('DELETE FROM __positions__')
  for (const t of Object.values(state.tables)) {
    db.run('INSERT INTO __tables__(id, name, schema_json) VALUES (?, ?, ?)', [
      t.id,
      t.name,
      JSON.stringify(t),
    ])
  }
  for (const [tid, rows] of Object.entries(state.records)) {
    rows.forEach((r, i) => {
      db.run('INSERT INTO __rows__(table_id, row_idx, data_json) VALUES (?, ?, ?)', [
        tid,
        i,
        JSON.stringify(r),
      ])
    })
  }
  for (const [tid, pos] of Object.entries(state.positions ?? {})) {
    db.run('INSERT INTO __positions__(table_id, x, y) VALUES (?, ?, ?)', [tid, pos.x, pos.y])
  }
  db.run('COMMIT')
}

export async function initDb(): Promise<Database> {
  const SQL = await getSqlStatic()
  const adapter = getAdapter()
  const bytes = await adapter.load()
  const db = bytes ? new SQL.Database(bytes) : new SQL.Database()
  ensureSchema(db)
  return db
}

export async function saveState(db: Database, state: SaveableState): Promise<void> {
  writeState(db, state)
  const bytes = db.export()
  await getAdapter().save(bytes)
}
