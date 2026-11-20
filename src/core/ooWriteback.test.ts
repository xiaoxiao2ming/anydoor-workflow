/**
 * OO 写回管线测试
 *
 * 立档：2026-10-06 · 配套 src/core/ooWriteback.ts
 *
 * ★ 重点验两件事：
 *   ① 值真的写进去了（且写在了正确的格子/sdt）
 *   ② ★★ 部件零丢失 —— 这是 OO 域的硬指标。
 *      `workbookToBytes` 只写回自己建模的部件，真实文件里的
 *      图表/绘图/透视表会**静默丢失**（见 xlsx/patchSave.ts 的实测记录）。
 */

import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import {
  fmtCellValue, planWrites, splitA1, colToNum,
  patchSheetXml, patchDocxSdt, writebackOo, type OoRefSite,
} from './ooWriteback'

/** 造一个最小 xlsx：含一张表+ 一个图表部件（模拟真实文件） */
async function makeXlsx (cellValue = '旧值'): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="a"/><Default Extension="xml" ContentType="b"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="c"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="d"/>' +
    '<Override PartName="/xl/charts/chart1.xml" ContentType="e"/>' +
    '</Types>')
  zip.folder('_rels')!.file('.rels', '<Relationships/>')
  zip.file('xl/workbook.xml', '<workbook/>')
  zip.file('xl/worksheets/sheet1.xml',
    `<?xml version="1.0"?><worksheet><sheetData>` +
    `<row r="1"><c r="A1" t="str"><v>${cellValue}</v></c></row>` +
    `<row r="3"><c r="B3"><v>100</v></c><c r="C3" t="s"><v>0</v></c></row>` +
    `</sheetData></worksheet>`)
  // ★ 真实文件里这些部件最容易被「重建」丢掉
  zip.file('xl/charts/chart1.xml', '<chart>图表内容</chart>')
  zip.file('xl/drawings/drawing1.xml', '<drawing>绘图</drawing>')
  zip.file('xl/media/image1.png', new Uint8Array([137, 80, 78, 71]))
  return zip.generateAsync({ type: 'uint8array' })
}

describe('① 值格式化：空值不能变成 "undefined"',
  () => {
  it('null/undefined → 空串', () => {
    expect(fmtCellValue(null)).toBe('')
    expect(fmtCellValue(undefined)).toBe('')
  })
  it('数字/布尔/字符串', () => {
    expect(fmtCellValue(12)).toBe('12')
    expect(fmtCellValue(12.5)).toBe('12.5')
    expect(fmtCellValue(true)).toBe('TRUE')
    expect(fmtCellValue('甲')).toBe('甲')
  })
  it('NaN/Infinity → 空串（不写垃圾进文件）', () => {
    expect(fmtCellValue(NaN)).toBe('')
    expect(fmtCellValue(Infinity)).toBe('')
  })
})

describe('② A1 解析', () => {
  it('基本解析', () => {
    expect(splitA1('B3')).toEqual({ col: 'B', row: 3 })
    expect(colToNum('A')).toBe(1)
    expect(colToNum('Z')).toBe(26)
    expect(colToNum('AA')).toBe(27)
  })
  it('非法输入返回 null（不崩）', () => {
    expect(splitA1('')).toBeNull()
    expect(splitA1('3B')).toBeNull()
  })
})

describe('③ 写回计划：落点缺失要报 skipped 而不是静默', () => {
  const sites: OoRefSite[] = [
    { ref: { nodeId: 't1', fieldId: 'f1', urn: 'node://t1#f1' }, a1: 'A1', part: 'xl/worksheets/sheet1.xml' },
    { ref: { nodeId: 't1', fieldId: 'f2' } }, // ★ 无 a1
  ]
  it('产出格子清单 + 记录跳过原因', () => {
    const r = planWrites(sites, () => '新值')
    expect(r.cells).toHaveLength(1)
    expect(r.cells[0]).toEqual({ a1: 'A1', value: '新值' })
    expect(r.skipped).toEqual([{ fieldId: 'f2', why: '无落点（a1 缺失）' }])
  })
  it('同一格被多个引用指向时只写一次', () => {
    const dup: OoRefSite[] = [
      { ref: { nodeId: 't1', fieldId: 'a' }, a1: 'A1', part: 'xl/worksheets/sheet1.xml' },
      { ref: { nodeId: 't2', fieldId: 'b' }, a1: 'A1', part: 'xl/worksheets/sheet1.xml' },
    ]
    expect(planWrites(dup, () => 'x').cells).toHaveLength(1)
  })
})

describe('④ sheetXML 定点替换', () => {
  it('替换目标 cell 的值，保留其它 cell', () => {
    const xml = '<worksheet><sheetData><row r="1"><c r="A1" t="str"><v>旧</v></c>' +
      '<c r="B1"><v>2</v></c></row></sheetData></worksheet>'
    const r = patchSheetXml(xml, 'A1', '新')
    expect(r.hit).toBe(true)
    expect(r.xml).toContain('<v>新</v>')
    expect(r.xml).toContain('<c r="B1"><v>2</v></c>') // ★ 邻居未被动
  })
  it('去掉旧的 t 属性（类型与内容不符会让 Excel 报修复）', () => {
    const xml = '<c r="A1" t="str"><v>旧</v></c>'
    const r = patchSheetXml(xml, 'A1', '新')
    expect(r.xml).not.toContain('t="str"')
  })
  it('找不到该 cell 时不新建（hit=false）', () => {
    const r = patchSheetXml('<c r="A1"><v>1</v></c>', 'Z9', 'x')
    expect(r.hit).toBe(false)
    expect(r.xml).toBe('<c r="A1"><v>1</v></c>') // 原样不动
  })
  it('特殊字符被转义（不破坏 XML）', () => {
    const r = patchSheetXml('<c r="A1"><v>旧</v></c>', 'A1', 'a<b&c"d')
    expect(r.xml).toContain('a&lt;b&amp;c&quot;d')
  })
})

describe('⑤ docx sdt 定点替换', () => {
  const DOC = '<?xml version="1.0"?><w:document><w:body>' +
    '<w:sdt><w:sdtPr><w:tag w:val="表甲.数量"/></w:sdtPr>' +
    '<w:sdtContent><w:p><w:r><w:t>旧值</w:t></w:r></w:p></w:sdtContent></w:sdt>' +
    '<w:sdt><w:sdtPr><w:tag w:val="表乙.单价"/></w:sdtPr>' +
    '<w:sdtContent><w:p><w:r><w:t>保留</w:t></w:r></w:p></w:sdtContent></w:sdt>' +
    '</w:body></w:document>'
  it('只改目标 sdt，另一个不动', () => {
    const r = patchDocxSdt(DOC, '表甲.数量', '新值')
    expect(r.hit).toBe(true)
    expect(r.xml).toContain('<w:t>新值</w:t>')
    expect(r.xml).toContain('<w:t>保留</w:t>')   // ★ 另一个 sdt 未被动
  })
  it('找不到 tag 时不动文件', () => {
    const r = patchDocxSdt(DOC, '不存在', 'x')
    expect(r.hit).toBe(false)
    expect(r.xml).toBe(DOC)
  })
})

describe('★★⑥ 端到端：字节进 → 字节出，且部件零丢失', () => {
  it('oosheet：改到指定单元格，且图表/绘图/媒体全部保留', async () => {
    const src = await makeXlsx('旧值')
    const sites: OoRefSite[] = [
      { ref: { nodeId: 't1', fieldId: 'f1', urn: 'node://t1#f1' }, a1: 'A1', part: 'xl/worksheets/sheet1.xml' },
    ]
    const r = await writebackOo({
      bytes: src, kind: 'oosheet', sites, pick: () => '新值',
    })
    expect(r.written).toBe(1)
    expect(r.labels[0]).toContain('A1=新值')

    // ★★★ 铁律：解包核对部件零丢失
    const out = await JSZip.loadAsync(r.bytes)
    for (const part of [
      '[Content_Types].xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml',
      'xl/charts/chart1.xml', 'xl/drawings/drawing1.xml', 'xl/media/image1.png',
    ]) {
      expect(out.file(part), '部件丢失：' + part).not.toBeNull()
    }
    // 值真的进去了
    const sheet = await out.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('<v>新值</v>')
  })

  it('ooword：改到 sdt，且原部件保留', async () => {
    const zip = new JSZip()
    zip.file('[Content_Types].xml', '<Types/>')
    zip.file('word/document.xml',
      '<w:document><w:body><w:sdt><w:sdtPr><w:tag w:val="表甲.数量"/></w:sdtPr>' +
      '<w:sdtContent><w:p><w:r><w:t>旧</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>')
    zip.file('word/numbering.xml', '<w:numbering/>')   // 编号部件
    zip.file('word/media/image1.png', new Uint8Array([1, 2, 3]))
    const src = await zip.generateAsync({ type: 'uint8array' })

    const r = await writebackOo({
      bytes: src, kind: 'ooword',
      sites: [{ ref: { nodeId: 't1', fieldId: 'f1' }, sdtTag: '表甲.数量' }],
      pick: () => '新值',
    })
    expect(r.written).toBe(1)
    const out = await JSZip.loadAsync(r.bytes)
    expect(out.file('word/numbering.xml')).not.toBeNull()
    expect(out.file('word/media/image1.png')).not.toBeNull()
    const doc = await out.file('word/document.xml')!.async('string')
    expect(doc).toContain('<w:t>新值</w:t>')
  })

  it('没有落点时字节仍可往返（不产出半损坏文件），并如实报 skipped', async () => {
    const src = await makeXlsx()
    const r = await writebackOo({
      bytes: src, kind: 'oosheet',
      sites: [{ ref: { nodeId: 't1', fieldId: 'f1' } }],   // ★ 无 a1 无 part
      pick: () => 'x',
    })
    expect(r.written).toBe(0)
    expect(r.skipped.length).toBeGreaterThan(0)
    const out = await JSZip.loadAsync(r.bytes)
    expect(out.file('xl/charts/chart1.xml')).not.toBeNull() // 原样保留
  })

  it('★ 往返两次字节稳定（幂等，不引入漂移）', async () => {
    const src = await makeXlsx()
    const sites: OoRefSite[] = [
      { ref: { nodeId: 't1', fieldId: 'f1' }, a1: 'A1', part: 'xl/worksheets/sheet1.xml' },
    ]
    const a = await writebackOo({ bytes: src, kind: 'oosheet', sites, pick: () => '稳定值' })
    const b = await writebackOo({ bytes: a.bytes, kind: 'oosheet', sites, pick: () => '稳定值' })
    const sa = await (await JSZip.loadAsync(a.bytes)).file('xl/worksheets/sheet1.xml')!.async('string')
    const sb = await (await JSZip.loadAsync(b.bytes)).file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sa).toBe(sb)   // 同样输入不产生不同结构
  })
})
