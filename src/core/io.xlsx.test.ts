import { describe, it, expect } from 'vitest'
import { buildXlsxBlob, parseFile } from './io'
import { defineField } from './schema'

/* 2026-10-22 去 SheetJS（CVE）后 io.ts xlsx 链路 roundtrip：
 * buildXlsxBlob（@office-kit createWorkbook 链）→ parseFile（loadWorkbook 链）字段/行一致 */

function fileOf(blob: Blob, name: string): File {
  return new File([blob as BlobPart], name)
}

describe('io xlsx roundtrip（去 SheetJS 后）', () => {
  it('导出→导入：中文字段 + number 推断 + 空值→null', async () => {
    const headers = ['材料名称', '数量', '单价', '备注']
    const matrix: unknown[][] = [
      ['无缝钢管 DN50', '120', '45.5', '米'],
      ['闸阀 Z45T-16', '8', '260', ''],
      ['弯头 90°', '36', '12.8', '焊接'],
    ]
    const blob = await buildXlsxBlob('材料表', headers, matrix)
    expect(blob.size).toBeGreaterThan(0)

    const { fields, rows } = await parseFile(fileOf(blob, '材料表.xlsx'))
    expect(fields.map((f) => f.name)).toEqual(headers)
    expect(fields.map((f) => f.type)).toEqual(['text', 'number', 'number', 'text'])
    expect(rows).toHaveLength(3)
    expect(rows[0]['c0']).toBe('无缝钢管 DN50')
    expect(rows[0]['c1']).toBe(120)
    expect(rows[0]['c2']).toBe(45.5)
    expect(rows[1]['c3']).toBe('')
    expect(rows[2]['c3']).toBe('焊接')
  })

  it('导出→导入：空数据行只有表头', async () => {
    const headers = ['A', 'B']
    const blob = await buildXlsxBlob('空表', headers, [])
    const { fields, rows } = await parseFile(fileOf(blob, '空表.xlsx'))
    expect(fields.map((f) => f.name)).toEqual(headers)
    expect(rows).toHaveLength(0)
  })

  it('字段 id 稳定：c0..cN 顺序与列一致', async () => {
    const headers = ['甲', '乙', '丙']
    const blob = await buildXlsxBlob('t', headers, [['1', '2', '3']])
    const { fields } = await parseFile(fileOf(blob, 't.xlsx'))
    expect(fields.map((f) => f.id)).toEqual(['c0', 'c1', 'c2'])
    expect(fields[0].id).toBe(defineField('x', 'text', { id: 'c0' }).id)
  })
})
