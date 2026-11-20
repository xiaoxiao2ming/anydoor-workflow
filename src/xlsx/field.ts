/**
 * Excel 侧域系统 —— 与 Word 侧（field.ts）共用同一张 ParamTable
 *
 * Excel 的「域」用原生机制承载，和 Word 的 w:sdt 对等：
 *   命名区域 definedName  ←→  单个标量域（相当于 Word 的文本框控件）
 *   Table（ListObject）   ←→  行级循环域（相当于 Word 的 each 行控件）
 *
 * 同样遵守纯计算层铁律：不做文件 IO，Workbook 由调用方传入或给出字节。
 */

import {
  addDefinedName,
  listDefinedNames,
  getDefinedNameTarget,
  getSheet,
  sheetNames,
} from '@office-kit/xlsx/workbook'
import {
  setCellByCoord,
  getCellByCoord,
  listTables,
  addExcelTable,
} from '@office-kit/xlsx/worksheet'
import { makeFormula } from '@office-kit/xlsx/cell'
import { loadWorkbook, workbookToBytes, fromArrayBuffer } from '@office-kit/xlsx/io'
import type { Workbook } from '@office-kit/xlsx/workbook'
import type { Worksheet } from '@office-kit/xlsx/worksheet'
import { ParamTable, type ResolveTrace } from '../ref/paramTable'
import { evalRowFormula } from './evalFormula'
import { colOffset } from './a1'
import { normalizeWorkbookBytes } from './normalize'
import { classifyOpenFailure, sniffContainer, WorkbookOpenError } from './openError'

/*
 * 入口选择（实测结论，别改回 `@office-kit/xlsx/node`）：
 *   · `./io`  → 零 node: 导入、零 Buffer，导出 fromArrayBuffer / fromBlob / loadWorkbook / workbookToBytes
 *   · `./node`→ import 了 node:fs / node:stream，会把 Node 内置模块拖进浏览器包
 * workbookToBytes 的源码注释也自认 browser-safe（内存 Uint8Array sink）。
 * 所以整个 Excel 链路在 Electron 主进程、渲染进程、纯浏览器里跑的是同一份代码。
 */

/* ────────────────── 类型 ────────────────── */

export interface XlsxFieldSpec {
  key: string
  kind: 'definedName' | 'table'
  /** definedName → 'Sheet1!$B$2'；table → 'A1:C4' */
  ref: string
  sheet?: string
  /** table 的列名（可绑定到数组参数的字段） */
  columns?: string[]
  sample?: string
}

export interface XlsxBindResult {
  buffer: Uint8Array
  bound: number
  warnings: string[]
  traces: ResolveTrace[]
  fields: XlsxFieldSpec[]
}

/** 附加配置 */
export interface XlsxBindExtras {
  /**
   * Table 的计算列公式。形如：
   *   { 材料表: { 合价: '={数量}*{单价}' } }
   * 绑定时把 `{列名}` 换成该行的真实单元格引用，写成 formula + cachedValue 双写：
   *   · Excel/WPS 打开 → 能重算（改数量，合价跟着变）
   *   · 不重算也能看 → 缓存值兜底（SP1 实测：不写缓存的单元格 Excel 会显示 #N/A）
   */
  tableFormulas?: Record<string, Record<string, string>>
}

/* ────────────────── 槽位发现 ────────────────── */

/**
 * 遍历工作表。
 * 注意：不能直接用 wb.sheets —— loadWorkbook 读回来的工作簿里该字段结构与内存态不同
 * （实测会拿到 title 为空的对象，导致后续 listTables 静默返回空）。
 * 走 sheetNames + getSheet 才是稳定的。
 */
function allSheets(wb: Workbook): Worksheet[] {
  const out: Worksheet[] = []
  for (const name of sheetNames(wb) ?? []) {
    const ws = getSheet(wb, name)
    if (ws) out.push(ws)
  }
  return out
}

/** 列出工作簿里所有「域」：命名区域 + Table */
export function discoverXlsxFields(wb: Workbook): XlsxFieldSpec[] {
  const fields: XlsxFieldSpec[] = []

  // ① 命名区域（标量域）
  const names = listDefinedNames(wb) ?? []
  for (const n of names as any[]) {
    const tgt = getDefinedNameTarget(wb, n.name, n.scope)
    const first = Array.isArray(tgt) ? tgt[0] : undefined
    fields.push({
      key: n.name,
      kind: 'definedName',
      ref: n.value,
      sheet: first?.sheet,
      sample: '',
    })
  }

  // ② Table（行级循环域）
  for (const ws of allSheets(wb)) {
    for (const t of listTables(ws) ?? []) {
      fields.push({
        key: t.displayName || t.name || '',
        kind: 'table',
        ref: t.ref,
        sheet: ws.title,
        columns: ((t as any).columns ?? []).map((c: any) => (typeof c === 'string' ? c : c.name)),
      })
    }
  }
  return fields
}

/** 便捷入口：直接从字节打开并发现域（面板用；内部用浏览器安全的 io 入口） */
export async function discoverXlsxFieldsFromBytes(
  input: ArrayBuffer | Uint8Array
): Promise<XlsxFieldSpec[]> {
  const { wb } = await openWorkbook(input)
  return discoverXlsxFields(wb)
}

/**
 * 打开工作簿（含归一化）。
 *
 * 真实 Excel/WPS 导出的文件里有 office-kit 解析器不接受的合法构造，
 * 必须先过一遍 normalize，否则会直接抛异常。见 normalize.ts 的说明。
 * 返回 applied 供 UI 如实展示「为了让解析器读懂，改了什么」。
 */
export async function openWorkbook(
  input: ArrayBuffer | Uint8Array
): Promise<{ wb: Workbook; applied: string[] }> {
  const norm = await normalizeWorkbookBytes(input)
  const wb = await loadWorkbook(fromArrayBuffer(norm.bytes))
  return { wb, applied: norm.applied }
}

/* ────────────────── 写入 ────────────────── */

function stripDollars(ref: string): string {
  return String(ref).replace(/\$/g, '')
}

/** 写值：字符串以 '=' 开头视为公式 → 双写 formula + cachedValue */
function writeCell(ws: Worksheet, coord: string, value: unknown): void {
  if (typeof value === 'string' && value.startsWith('=')) {
    const formula = value.slice(1)
    // 缓存值用现有单元格值兜底；没有就不写（调用方可自行补）
    setCellByCoord(ws, coord, makeFormula(formula))
    return
  }
  setCellByCoord(ws, coord, value as any)
}

/** 写公式并同时给缓存值（SP1 实测：不写缓存 Excel 会显示 #N/A） */
export function writeFormula(ws: Worksheet, coord: string, formula: string, cached: unknown): void {
  setCellByCoord(ws, coord, makeFormula(formula, { cachedValue: cached as any }))
}

/* ────────────────── 绑定 ────────────────── */

/**
 * 参数表 → 填充 Excel。
 * 与 Word 侧 bindFromParamTable 对称：同一张 ParamTable 可同时喂两端。
 */
export async function bindXlsxFromParamTable(
  input: ArrayBuffer | Uint8Array,
  table: ParamTable,
  extras: XlsxBindExtras = {}
): Promise<XlsxBindResult> {
  // 真实 Excel/WPS 文件要先归一化，否则 office-kit 解析器会拒收（见 normalize.ts）
  const norm = await normalizeWorkbookBytes(input)
  const wb = await loadWorkbook(fromArrayBuffer(norm.bytes))
  const fields = discoverXlsxFields(wb)
  const data = table.resolveAll()
  const warnings: string[] = []
  let bound = 0

  // 改了用户的文件结构就要如实说 —— 不静默
  for (const a of norm.applied) warnings.push(`为让解析器读懂，已归一化：${a}`)

  for (const f of fields) {
    if (f.kind === 'definedName') {
      const v = data[f.key]
      if (v === undefined) {
        warnings.push(`命名区域「${f.key}」在参数表中无对应值`)
        continue
      }
      if (Array.isArray(v)) {
        warnings.push(`命名区域「${f.key}」只能绑定标量，收到数组（应改用 Table）`)
        continue
      }
      const ws = f.sheet ? getSheet(wb, f.sheet) : undefined
      if (!ws) {
        warnings.push(`命名区域「${f.key}」指向的工作表不存在`)
        continue
      }
      // ref 形如 '工程量表!$B$2' 或 '$B$2'
      const cellRef = stripDollars(String(f.ref).split('!').pop() || '')
      if (!/^[A-Za-z]{1,3}\d{1,7}$/.test(cellRef)) {
        warnings.push(`命名区域「${f.key}」不是单单元格（${f.ref}），暂不支持写入`)
        continue
      }
      writeCell(ws, cellRef, v)
      bound++
      continue
    }

    // ② Table：参数值应为数组
    if (f.kind === 'table') {
      const rows = data[f.key]
      if (!Array.isArray(rows)) {
        warnings.push(`Table「${f.key}」需要数组参数，收到 ${typeof rows}`)
        continue
      }
      const ws = f.sheet ? getSheet(wb, f.sheet) : undefined
      if (!ws || !f.columns?.length) {
        warnings.push(`Table「${f.key}」缺少工作表或列定义`)
        continue
      }
      // ref 形如 'A1:C4'，第一行是表头，从第二行开始写数据
      const m = /^([A-Za-z]+\d+):([A-Za-z]+\d+)$/.exec(f.ref)
      if (!m) {
        warnings.push(`Table「${f.key}」ref 无法解析：${f.ref}`)
        continue
      }
      const startCol = (m[1].match(/^[A-Za-z]+/) || ['A'])[0]
      const startRow = Number((m[1].match(/\d+/) || ['1'])[0])
      const headerRow = startRow
      for (let i = 0; i < rows.length; i++) {
        const r = rows[i] as Record<string, unknown>
        for (let c = 0; c < f.columns.length; c++) {
          const val = r?.[f.columns[c]]
          if (val === undefined) continue
          const colLetter = colOffset(startCol, c)
          writeCell(ws, `${colLetter}${headerRow + 1 + i}`, val)
        }
      }
      // 计算列：写 formula + cachedValue 双写
      //   Excel/WPS 打开 → 能重算（改数量，合价自动跟着变）
      //   不重算也能看   → 缓存值兜底（SP1 实测：不写缓存的单元格 Excel 显示 #N/A）
      const fcols = extras.tableFormulas?.[f.key]
      if (fcols) {
        for (const [colName, tpl] of Object.entries(fcols)) {
          const ci = f.columns.indexOf(colName)
          if (ci < 0) {
            warnings.push(`公式列「${colName}」不在 Table「${f.key}」的列里（现有列：${f.columns.join(',')}）`)
            continue
          }
          const colLetter = colOffset(startCol, ci)
          let failCount = 0
          let failReason = ''
          for (let i = 0; i < rows.length; i++) {
            const rowNum = headerRow + 1 + i
            const rowObj = (rows[i] as Record<string, unknown>) ?? {}
            const fml = rowFormula(tpl, f.columns, startCol, rowNum)

            let cached = rowObj[colName]
            if (cached === undefined) {
              // 上游没提供这一列 —— 自己算缓存值。
              // 不写缓存的话，落盘 XML 里连 <v> 都没有，不执行公式的阅读器一律显示 #N/A。
              const refValues: Record<string, unknown> = {}
              f.columns.forEach((cn, k) => {
                refValues[colOffset(startCol, k) + rowNum] = rowObj[cn] ?? 0
              })
              const ev = evalRowFormula(fml, refValues)
              if (ev.ok) cached = ev.value
              else {
                failCount++
                if (!failReason) failReason = ev.error ?? '未知原因'
              }
            }

            setCellByCoord(
              ws,
              `${colLetter}${rowNum}`,
              cached === undefined ? makeFormula(fml) : makeFormula(fml, { cachedValue: cached as any })
            )
          }
          // 逐行报会刷屏，按列聚合一条
          if (failCount) {
            warnings.push(
              `计算列「${colName}」有 ${failCount} 行没算出缓存值（${failReason}）：已只写公式。` +
                `Excel 打开能正常算，但不执行公式的阅读器会显示 #N/A —— ` +
                `想彻底避免，就在上游节点里也提供一列同名的计算结果。`
            )
          }
        }
      }

      // 扩展 Table 范围，让 Excel/WPS 认为新写入的行也在表内（否则新增行不在 ListObject 里）
      const t = (listTables(ws) ?? []).find((x: any) => (x.displayName || x.name) === f.key)
      if (t && f.columns.length) {
        const endRow = headerRow + rows.length
        ;(t as any).ref = `${startCol}${headerRow}:${colOffset(startCol, f.columns.length - 1)}${endRow}`
      }
      bound++
    }
  }

  const buffer = await workbookToBytes(wb, { mtime: new Date('2026-01-01T00:00:00Z') })
  return { buffer, bound, warnings, traces: table.trace(), fields }
}

/**
 * 行级公式模板 → 该行真实公式。
 *   rowFormula('={数量}*{单价}', ['名称','数量','单价'], 'A', 5)  →  'B5*C5'
 * 返回不带 '='（OOXML 的 <f> 里本来就不带等号）。
 */
function rowFormula(tpl: string, columns: string[], startCol: string, rowNum: number): string {
  const body = tpl.startsWith('=') ? tpl.slice(1) : tpl
  let out = body
  columns.forEach((c, i) => {
    out = out.split('{' + c + '}').join(colOffset(startCol, i) + rowNum)
  })
  return out
}


export { addDefinedName, addExcelTable, getCellByCoord, setCellByCoord, writeCell }
