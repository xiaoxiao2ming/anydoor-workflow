// 计算内核桥 —— 把 schema + 记录里的公式字段算出来。
//
// 为什么不用 HyperFormula：它是 GPLv3，闭源商用要买授权。
// 我们只需要「算一个值」，不需要它的依赖图 / 增量重算（原实现里每次 createEngine()
// 都重建一个空实例，就是没用上的证据），所以换成
//   src/xlsx/evalFormula.ts = 自研递归下降解析器 + formulajs（MIT，373 个可调用函数）
// 这样计算节点 / Word / Excel 三条线共用同一套函数语义，也不再背 GPL 义务。
import type { FieldDef, TableSchema } from './schema'
import type { Row } from './store'
import { evalRowFormula } from '../xlsx/evalFormula'
import { isMergeKind } from './nodeStyle'

/** 计算引擎上下文：只承载全局参数（原来是 HyperFormula 的命名表达式） */
export interface EngineContext {
  /** 公式里用 [@@NAME] 引用 */
  globals: Record<string, number | string>
}

const LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

/** 0-based 列序号 → A1 列字母（0→A, 25→Z, 26→AA） */
function colLetter(idx: number): string {
  let s = ''
  let n = idx + 1
  while (n > 0) {
    const m = (n - 1) % 26
    s = LETTERS[m] + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

/**
 * 把公式里的占位符翻译成求值器认得的写法：
 *  - `[@@NAME]` → 全局参数**的值**（数字直接内联；字符串加引号并把内部 `"` 转义成 `""`）
 *  - `[@字段名]` → 同行列引用，如 B2
 *
 * 注意：`@@` 必须先替换，否则会被 `[@field]` 的正则吞掉。
 *
 * 返回 `null` = 公式引用了未定义的全局参数。这时整条公式判为失败，
 * 而不是替成 0 继续算 —— 0 会产出一个"看起来正常的错数"，比报错更危险。
 */
function resolveFormula(
  formula: string,
  fields: FieldDef[],
  rowIdx: number,
  globals: Record<string, number | string>,
): string | null {
  let missing: string | null = null
  const named = formula.replace(/\[@@([^\]]+)\]/g, (_, name: string) => {
    const v = globals[name]
    if (v === undefined) {
      missing = name
      return ''
    }
    if (typeof v === 'number') return String(v)
    return '"' + String(v).replace(/"/g, '""') + '"'
  })
  if (missing !== null) return null

  return named.replace(/\[@([^\]]+)\]/g, (_, fieldName: string) => {
    const colIdx = fields.findIndex((f) => f.name === fieldName || f.id === fieldName)
    if (colIdx < 0) return '0'
    return `${colLetter(colIdx)}${rowIdx + 1}`
  })
}

/** 创建一个空的计算引擎（只带一个全局参数表） */
export function createEngine(): EngineContext {
  return { globals: {} }
}

/**
 * 算出一张表的公式字段。
 * - formula 字段：翻译占位符后交给 evalRowFormula 计算（可引用 `[@字段]` 与 `[@@全局参数]`）
 * - 其它字段：原值回填
 * - lookup 字段：先由 resolveLookups 在 JS 层填好值，这里当普通值回填
 *
 * 同一行内的公式字段可以互相引用。原 HyperFormula 靠依赖图自动定序，这里用
 * 「多轮迭代直到没有新进展」达到同样效果 —— 因为公式列的声明顺序未必等于依赖顺序，
 * 单趟顺序求值会在乱序时静默算出错值。
 */
export function calcTable(eng: EngineContext, schema: TableSchema, rows: Row[]): Row[] {
  const fields = schema.fields
  const formulaCols = fields
    .map((f, c) => ({ f, c }))
    .filter((x) => x.f.type === 'formula')

  return rows.map((row, r) => {
    const computed: Row = { ...row }

    // 该行所有单元格的值，键是 A1 坐标 —— evalRowFormula 靠它把引用换成实际值
    const refValues: Record<string, unknown> = {}
    const put = (colIdx: number, v: unknown) => {
      refValues[`${colLetter(colIdx)}${r + 1}`] = v
    }
    fields.forEach((f, c) => {
      if (f.type === 'formula') return
      const v = row[f.id]
      put(c, v === undefined || v === null ? '' : v)
    })

    const settled = new Set<string>()
    for (let pass = 0; pass < formulaCols.length; pass++) {
      let progressed = false
      for (const { f, c } of formulaCols) {
        if (settled.has(f.id)) continue
        const body = resolveFormula(f.formula ?? '', fields, r, eng.globals)
        if (body === null) {
          // 引用了未定义的全局参数：[@@NAME] 的语义就是别处没定义
          computed[f.id] = '#NAME?'
          settled.add(f.id)
          continue
        }
        const out = evalRowFormula(body, refValues)
        if (out.ok) {
          computed[f.id] = out.value
          put(c, out.value)
          settled.add(f.id)
          progressed = true
        }
      }
      if (!progressed) break
    }

    // 依然算不出来的（多半是引用了不存在的字段或用了不支持的函数）：
    // 给明确的错误标记，而不是留一个空值让人以为"这里本来就没数"
    for (const { f } of formulaCols) {
      if (!settled.has(f.id)) computed[f.id] = '#ERROR'
    }
    return computed
  })
}

/**
 * 跨表 lookup 解析（JS 层，在 calcTable 之前跑）。
 * 对每个 lookup 字段：取本行 sourceField 的值，去关联表找 matchField 相等的行，取 returnField 回填。
 * 返回新的 records（不改原对象）。
 */
export function resolveLookups(
  tables: Record<string, TableSchema>,
  records: Record<string, Row[]>,
): Record<string, Row[]> {
  const out: Record<string, Row[]> = {}
  for (const [tid, rows] of Object.entries(records)) {
    if (!Array.isArray(rows)) {
      // 不静默：谁把非数组写进了 records，当场报出 tid 与现场（排查 2026-09-24 白屏崩溃）
      console.error('[resolveLookups] records[' + tid + '] 不是数组：', JSON.stringify(rows), 'tables 有此项 =', tid in tables)
      continue
    }
    const schema = tables[tid]
    out[tid] = rows.map((row) => {
      if (!schema) return row
      const filled: Row = { ...row }
      for (const f of schema.fields) {
        if (f.type === 'lookup' && f.lookup) {
          const src = row[f.lookup.sourceField]
          const target = records[f.lookup.table] ?? []
          const hit = target.find((r) => r[f.lookup!.matchField] === src)
          filled[f.id] = hit ? hit[f.lookup.returnField] : undefined
        }
      }
      return filled
    })
  }
  return out
}

/** 设置全局参数，公式里用 `[@@name]` 引用 */
export function setGlobalParam(eng: EngineContext, name: string, value: number | string): void {
  eng.globals[name] = value
}

/**
 * 数据流合并：把多个上游节点的行汇入本节点。
 * - 列 = 所有上游字段按 id 去重后的并集（保序）
 * - 行 = 依次取每个上游的记录，缺失列填空
 * 返回 { fields, rows }，供排序/输出节点直接当表用。
 */
export function mergeInputs(
  tables: Record<string, TableSchema>,
  records: Record<string, Row[]>,
  inputIds: string[],
): { fields: FieldDef[]; rows: Row[] } {
  const fields: FieldDef[] = []
  const seen = new Set<string>()
  for (const id of inputIds) {
    for (const f of tables[id]?.fields ?? []) {
      if (!seen.has(f.id)) {
        seen.add(f.id)
        fields.push(f)
      }
    }
  }
  const rows: Row[] = []
  for (const id of inputIds) {
    for (const r of records[id] ?? []) {
      const row: Row = {}
      for (const f of fields) row[f.id] = r[f.id]
      rows.push(row)
    }
  }
  return { fields, rows }
}

/** 按字段排序（稳定排序）。sortDir 默认 asc。非数字按字符串比。 */
export function sortRows(
  rows: Row[],
  fields: FieldDef[],
  sortBy: string | undefined,
  sortDir: 'asc' | 'desc' = 'asc',
): Row[] {
  if (!sortBy) return rows
  const col = fields.find((f) => f.id === sortBy)
  const dir = sortDir === 'desc' ? -1 : 1
  const numeric = col?.type === 'number'
  return [...rows].sort((a, b) => {
    const av = a[sortBy]
    const bv = b[sortBy]
    let cmp: number
    if (numeric) {
      cmp = (Number(av) || 0) - (Number(bv) || 0)
    } else {
      cmp = String(av ?? '').localeCompare(String(bv ?? ''), 'zh-CN')
    }
    return cmp * dir
  })
}

/** 合并多个上游节点（按 id 递归计算后的结果）为 { fields, rows } */
function mergeComputed(
  inputIds: string[],
  tables: Record<string, TableSchema>,
  records: Record<string, Row[]>,
  seen: Set<string>,
): { fields: FieldDef[]; rows: Row[] } {
  const fields: FieldDef[] = []
  const seenF = new Set<string>()
  const rows: Row[] = []
  for (const src of inputIds) {
    const up = computeNode(src, tables, records, seen)
    for (const f of up.fields) {
      if (!seenF.has(f.id)) {
        seenF.add(f.id)
        fields.push(f)
      }
    }
    for (const r of up.rows) {
      const row: Row = {}
      for (const f of fields) row[f.id] = r[f.id]
      rows.push(row)
    }
  }
  return { fields, rows }
}

/** 映射列 id 前缀：map_<对照表字段 id>，避免与上游同名字段撞车 */
const MAP_PREFIX = 'map_'

/**
 * 数据对照表（map 节点）核心：把上游的每一行拿去对照表里"查表"，把取回的列追加到行上。
 * - 匹配：上游行[mapKeyField] === 对照表行[mapMatchField]（去空格后精确比，大小写敏感）
 * - 命中：把 mapReturnFields 指定的列追加为 map_<字段id>；多列可一次取回
 * - 未命中：填 mapDefault（默认空串）
 * - 未配置规则（缺 key/match/return 任一）时原样透传上游，不做加工
 * - 追加列与上游列同名时自动加「（对照）」后缀，保证列名唯一
 */
export function applyMapping(
  mapTable: TableSchema,
  mapRows: Row[],
  inputFields: FieldDef[],
  inputRows: Row[],
): { fields: FieldDef[]; rows: Row[] } {
  const cfg = mapTable.config ?? {}
  const keyField = cfg.mapKeyField
  const matchField = cfg.mapMatchField
  const returnIds = cfg.mapReturnFields ?? []
  if (!keyField || !matchField || !returnIds.length) {
    return { fields: inputFields, rows: inputRows }
  }

  // 对照表索引：匹配值 → 命中行（同键多行时取第一条）
  const index = new Map<string, Row>()
  for (const r of mapRows) {
    const raw = r[matchField]
    if (raw === undefined || raw === null || raw === '') continue
    const k = String(raw).trim()
    if (!index.has(k)) index.set(k, r)
  }

  // 追加列定义
  const byId = new Map(mapTable.fields.map((f) => [f.id, f]))
  const used = new Set(inputFields.map((f) => f.name))
  const extra: FieldDef[] = []
  for (const id of returnIds) {
    const f = byId.get(id)
    if (!f) continue
    const name = used.has(f.name) ? `${f.name}（对照）` : f.name
    used.add(name)
    extra.push({ ...f, id: `${MAP_PREFIX}${f.id}`, name })
  }
  if (!extra.length) return { fields: inputFields, rows: inputRows }

  const fallback = cfg.mapDefault ?? ''
  const rows = inputRows.map((row) => {
    const raw = row[keyField]
    const hit = raw === undefined || raw === null || raw === '' ? undefined : index.get(String(raw).trim())
    const out: Row = { ...row }
    for (const ef of extra) {
      const srcId = ef.id.slice(MAP_PREFIX.length)
      out[ef.id] = hit ? (hit[srcId] ?? fallback) : fallback
    }
    return out
  })
  return { fields: [...inputFields, ...extra], rows }
}

/**
 * 计算任意节点的"展示结果"：字段 + 行。
 * - table/param：公式/lookup 算好的原始表
 * - sort：上游汇入后按配置排序
 * - map：上游汇入后按对照表查表补列（无上游时退化为普通对照表）
 * - output：上游汇入的结果
 * seen 用于防环（出现环时退化为原始记录，避免死递归）。
 */
export function computeNode(
  id: string,
  tables: Record<string, TableSchema>,
  records: Record<string, Row[]>,
  seen: Set<string> = new Set(),
): { fields: FieldDef[]; rows: Row[] } {
  const t = tables[id]
  if (!t) return { fields: [], rows: [] }
  if (t.kind === 'start') return { fields: [], rows: [] }
  if (seen.has(id)) return { fields: t.fields, rows: records[id] ?? [] }
  seen.add(id)
  // 汇聚型判定走 nodeStyle 的单一真相源（isMergeKind），勿在本文件另写一份
  if (isMergeKind(t.kind)) {
    // 已连线(linkedInputs)的来源排在前面 = 优先位置；其余（面板手动勾选）排在后面
    const linked = t.config?.linkedInputs ?? []
    const inputs = [...linked, ...((t.config?.inputs ?? []).filter((id) => !linked.includes(id)))]
    if (t.kind === 'sort') {
      const { fields, rows } = mergeComputed(inputs, tables, records, seen)
      return { fields, rows: sortRows(rows, fields, t.config?.sortBy, t.config?.sortDir) }
    }
    if (t.kind === 'map') {
      // 未接上游：当普通对照表用，返回自身对照数据（可直接被别的节点 lookup 引用）
      if (!inputs.length) {
        const resolved = resolveLookups(tables, records)
        const hf = createEngine()
        return { fields: t.fields, rows: calcTable(hf, t, resolved[id] ?? []) }
      }
      const merged = mergeComputed(inputs, tables, records, seen)
      return applyMapping(t, records[id] ?? [], merged.fields, merged.rows)
    }
    return mergeComputed(inputs, tables, records, seen)
  }
  // table / param
  const resolved = resolveLookups(tables, records)
  const hf = createEngine()
  return { fields: t.fields, rows: calcTable(hf, t, resolved[id] ?? []) }
}
