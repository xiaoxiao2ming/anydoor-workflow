// 数据 IO：导入外部文件 → 表结构+行；导出表 → 各类文件（浏览器下载）
import type { FieldDef } from './schema'
import { defineField } from './schema'
import type { Row } from './store'
import type { ExportFormat, ImportFormat } from './schema'
import { fromArrayBuffer, loadWorkbook, workbookToBytes } from '@office-kit/xlsx/io'
import { createWorkbook, addWorksheet, getSheet, sheetNames } from '@office-kit/xlsx/workbook'
import { getCellByCoord, getMaxRow, getMaxCol, setCellByCoord } from '@office-kit/xlsx/worksheet'
import { richTextToString, type Cell } from '@office-kit/xlsx/cell'
import { indexToCol } from '../xlsx/a1'

/** Cell → 展示文本（parseExcel 用；对齐原 SheetJS sheet_to_json raw 语义的可读子集） */
function cellToText(cell: Cell | undefined): string {
  if (!cell || cell.value == null) return ''
  const v = cell.value
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return String(v)
  if (v instanceof Date) return v.toISOString()
  if (typeof v === 'object' && 'kind' in v) {
    switch (v.kind) {
      case 'duration': return String(v.ms)
      case 'error': return v.code
      case 'rich-text': return richTextToString(v.runs)
      case 'formula': return v.cachedValue != null ? String(v.cachedValue) : ''
    }
  }
  return ''
}

/** 解析上传文件为 { 字段, 行 }。按扩展名判断格式。 */
export async function parseFile(file: File): Promise<{ fields: FieldDef[]; rows: Row[] }> {
  const name = file.name.toLowerCase()
  const text = await file.text()
  if (name.endsWith('.json')) return parseJson(text)
  if (name.endsWith('.xls') || name.endsWith('.xlsx')) return parseExcel(await file.arrayBuffer())
  // 默认当 CSV
  return parseCsv(text)
}

/** 导出为指定格式并触发浏览器下载。 */
export async function exportRows(
  name: string,
  fields: FieldDef[],
  rows: Row[],
  format: ExportFormat,
): Promise<void> {
  const headers = fields.map((f) => f.name)
  const matrix: unknown[][] = rows.map((r) => fields.map((f) => r[f.id] ?? ''))
  let blob: Blob
  let ext = format
  switch (format) {
    case 'csv':
      blob = new Blob([toCsv(headers, matrix)], { type: 'text/csv;charset=utf-8' })
      break
    case 'json':
      blob = new Blob([JSON.stringify(matrixToObjects(headers, matrix), null, 2)], {
        type: 'application/json',
      })
      break
    case 'xlsx':
      blob = await buildXlsxBlob(name, headers, matrix)
      ext = 'xlsx'
      break
    case 'txt':
      blob = new Blob([toTxt(headers, matrix)], { type: 'text/plain;charset=utf-8' })
      break
    case 'md':
      blob = new Blob([toMd(headers, matrix)], { type: 'text/markdown;charset=utf-8' })
      break
    case 'pdf':
      blob = await toPdf(name, headers, matrix)
      ext = 'pdf'
      break
    default:
      blob = new Blob([toCsv(headers, matrix)], { type: 'text/csv;charset=utf-8' })
  }
  downloadBlob(blob, `${name || 'export'}.${ext}`)
}

/* ---------------- 导入 ---------------- */

function inferFields(headers: string[], matrix: string[][]): FieldDef[] {
  const isNum = headers.map((_, i) =>
    matrix.every((r) => r[i] === undefined || r[i] === '' || !isNaN(Number(r[i]))),
  )
  return headers.map((h, i) => defineField(h.trim() || `列${i + 1}`, isNum[i] ? 'number' : 'text', { id: `c${i}` }))
}

function buildRows(fields: FieldDef[], matrix: string[][]): Row[] {
  return matrix.map((r) => {
    const row: Row = {}
    fields.forEach((f, i) => {
      const v = r[i]
      row[f.id] = f.type === 'number' ? (v === '' || v == null ? null : Number(v)) : v ?? ''
    })
    return row
  })
}

/** 极简 CSV 解析：支持引号包裹与转义双引号。 */
function parseCsv(text: string): { fields: FieldDef[]; rows: Row[] } {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let inQuotes = false
  for (let i = 0; i < text.length; i++) {
    const c = text[i]
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          field += '"'
          i++
        } else inQuotes = false
      } else field += c
    } else if (c === '"') inQuotes = true
    else if (c === ',') {
      row.push(field)
      field = ''
    } else if (c === '\n') {
      row.push(field)
      rows.push(row)
      row = []
      field = ''
    } else if (c !== '\r') field += c
  }
  if (field !== '' || row.length) {
    row.push(field)
    rows.push(row)
  }
  const headers = (rows[0] ?? []).map(String)
  const data = rows.slice(1).filter((r) => r.some((v) => (v ?? '').trim() !== ''))
  const fields = inferFields(headers, data)
  return { fields, rows: buildRows(fields, data) }
}

function parseJson(text: string): { fields: FieldDef[]; rows: Row[] } {
  const arr = JSON.parse(text)
  if (!Array.isArray(arr)) throw new Error('JSON 顶层必须是对象数组')
  const keys: string[] = []
  const seen = new Set<string>()
  for (const o of arr) {
    if (o && typeof o === 'object') {
      for (const k of Object.keys(o)) if (!seen.has(k)) {
        seen.add(k)
        keys.push(k)
      }
    }
  }
  const fields = keys.map((k, i) => {
    const allNum = arr.every(
      (o) => o[k] == null || o[k] === '' || typeof o[k] === 'number' || !isNaN(Number(o[k])),
    )
    return defineField(k, allNum ? 'number' : 'text', { id: `c${i}` })
  })
  const rows = arr.map((o) => {
    const row: Row = {}
    fields.forEach((f, i) => {
      const v = o[keys[i]]
      row[f.id] = f.type === 'number' ? (v === '' || v == null ? null : Number(v)) : v ?? ''
    })
    return row
  })
  return { fields, rows }
}

async function parseExcel(buf: ArrayBuffer): Promise<{ fields: FieldDef[]; rows: Row[] }> {
  /* 2026-10-22 去 SheetJS（CVE）：@office-kit/xlsx 同构读（./io 零 node 导入，field.ts 同款入口） */
  const wb = await loadWorkbook(fromArrayBuffer(buf))
  const ws = getSheet(wb, sheetNames(wb)[0] ?? '')
  const matrix: string[][] = []
  if (ws) {
    const maxRow = getMaxRow(ws)
    const maxCol = getMaxCol(ws)
    for (let r = 1; r <= maxRow; r++) {
      const row: string[] = []
      for (let c = 1; c <= maxCol; c++) row.push(cellToText(getCellByCoord(ws, `${indexToCol(c)}${r}`)))
      matrix.push(row)
    }
  }
  const headers = (matrix[0] ?? []).map(String)
  const data = (matrix.slice(1) as string[][]).filter((r) => r.some((v) => v !== '' && v != null))
  const fields = inferFields(headers, data)
  return { fields, rows: buildRows(fields, data.map((r) => r.map(String))) }
}

/* ---------------- 导出 ---------------- */

function matrixToObjects(headers: string[], matrix: unknown[][]): Record<string, unknown>[] {
  return matrix.map((r) => {
    const o: Record<string, unknown> = {}
    headers.forEach((h, i) => (o[h] = r[i] ?? ''))
    return o
  })
}

function toCsv(headers: string[], matrix: unknown[][]): string {
  const esc = (v: unknown) => {
    const s = String(v ?? '')
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s
  }
  const lines = [headers.map(esc).join(',')]
  for (const r of matrix) lines.push(r.map(esc).join(','))
  // 加 BOM 让 Excel 正确识别 UTF-8
  return '﻿' + lines.join('\n')
}

function toTxt(headers: string[], matrix: unknown[][]): string {
  const widths = headers.map((h, i) => Math.max(h.length, ...matrix.map((r) => String(r[i] ?? '').length)))
  const pad = (s: string, w: number) => s.padEnd(w)
  const line = (cells: unknown[]) => cells.map((c, i) => pad(String(c ?? ''), widths[i])).join('  ')
  const out: string[] = []
  out.push(line(headers))
  out.push(widths.map((w) => '─'.repeat(w)).join('  '))
  for (const r of matrix) out.push(line(r))
  return out.join('\n')
}

function toMd(headers: string[], matrix: unknown[][]): string {
  const lines = [`| ${headers.join(' | ')} |`]
  lines.push(`| ${headers.map(() => '---').join(' | ')} |`)
  for (const r of matrix) lines.push(`| ${r.map((c) => String(c ?? '').replace(/\|/g, '\\|')).join(' | ')} |`)
  return lines.join('\n')
}

/** 表数据 → xlsx Blob（导出即下载；独立导出便于 roundtrip 测试） */
export async function buildXlsxBlob(name: string, headers: string[], matrix: unknown[][]): Promise<Blob> {
  /* 2026-10-22 去 SheetJS（CVE）：createWorkbook 从零建簿（NodeEditor createBlankXlsx 同款链路） */
  const wb = createWorkbook()
  const ws = addWorksheet(wb, (name || 'sheet').slice(0, 31))
  headers.forEach((h, c) => setCellByCoord(ws, `${indexToCol(c + 1)}1`, h))
  matrix.forEach((row, r) => {
    row.forEach((v, c) => {
      const s = String(v ?? '')
      if (s !== '') setCellByCoord(ws, `${indexToCol(c + 1)}${r + 2}`, s)
    })
  })
  const out = await workbookToBytes(wb)
  return new Blob([out as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
}

async function toPdf(name: string, headers: string[], matrix: unknown[][]): Promise<Blob> {
  const { jsPDF } = await import('jspdf')
  const autoTable = (await import('jspdf-autotable')).default
  const doc = new jsPDF({ orientation: matrix.length > 20 ? 'landscape' : 'portrait' })
  autoTable(doc, {
    head: [headers],
    body: matrix.map((r) => r.map((c) => String(c ?? ''))),
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [37, 99, 235] },
    margin: { top: 14 },
  })
  doc.text(name || '导出', 14, 10)
  return doc.output('blob')
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

export const IMPORT_FORMATS: ImportFormat[] = ['csv', 'json', 'xlsx']
export const EXPORT_FORMATS: ExportFormat[] = ['csv', 'json', 'xlsx', 'txt', 'md', 'pdf']
