import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { createWorkbook, addWorksheet, addDefinedName, getSheet } from '@office-kit/xlsx/workbook'
import { setCellByCoord, addExcelTable, getCellByCoord } from '@office-kit/xlsx/worksheet'
import { isFormulaCell, getFormulaText, getCachedFormulaValue } from '@office-kit/xlsx/cell'
import { workbookToBytes, loadWorkbook } from '@office-kit/xlsx/io'
import { fromBuffer } from '@office-kit/xlsx/node'

import { ParamTable } from '../docx/field'
import { bindFromParamTable } from '../docx/field'
import { sdtifyDocx } from '../docx/sdtify'
import { discoverXlsxFields, bindXlsxFromParamTable } from './field'

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

/** 造一个含 2 个命名区域 + 1 个 Table 的 xlsx */
async function makeXlsx(): Promise<Uint8Array> {
  const wb = createWorkbook()
  const ws = addWorksheet(wb, '工程量表')
  setCellByCoord(ws, 'A1', '工程名称')
  setCellByCoord(ws, 'B1', '待填')
  setCellByCoord(ws, 'A2', '设计流量')
  setCellByCoord(ws, 'B2', 0)
  setCellByCoord(ws, 'A4', '名称')
  setCellByCoord(ws, 'B4', '数量')
  setCellByCoord(ws, 'C4', '单价')
  addDefinedName(wb, { name: '工程名称', value: '工程量表!$B$1' })
  addDefinedName(wb, { name: '设计流量', value: '工程量表!$B$2' })
  addExcelTable(wb, ws, { name: '材料表', ref: 'A4:C5', columns: ['名称', '数量', '单价'] })
  return workbookToBytes(wb, { mtime: new Date('2026-01-01T00:00:00Z') })
}

/** 造一个含绑定标记的 docx */
async function makeDocx(): Promise<Uint8Array> {
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>
<w:p><w:r><w:t>工程名称：</w:t></w:r><w:r><w:t>{{field:工程名称}}</w:t></w:r></w:p>
<w:p><w:r><w:t>设计流量：</w:t></w:r><w:r><w:t>{{field:设计流量}}</w:t></w:r></w:p>
<w:sectPr/></w:body></w:document>`
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`)
  zip.file('word/document.xml', doc)
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

const CTX = {
  nodes: {
    calc1: { Q: 1200 },
    table1: {
      rows: [
        { 名称: '钢管 DN200', 数量: 100, 单价: 25.5 },
        { 名称: '钢管 DN100', 数量: 40, 单价: 18.2 },
        { 名称: '阀门 DN50', 数量: 12, 单价: 260 },
      ],
    },
  },
  params: { 审查单位: '陵水审图' },
}

const SPECS = [
  { key: '工程名称', default: '海南陵水燃气工程' },
  { key: '设计流量', unit: 'Nm³/h', source: 'node://calc1#Q' },
  { key: '材料表', source: 'node://table1#rows' },
  { key: '审查单位', source: 'param://审查单位' },
]

/* ───────────────────────────────────────────── */

describe('Excel 侧槽位发现', () => {
  it('命名区域与 Table 都能发现', async () => {
    const wb = await loadWorkbook(fromBuffer(await makeXlsx()))
    const fields = discoverXlsxFields(wb)
    const keys = fields.map((f) => f.key).sort()
    expect(keys).toEqual(['工程名称', '材料表', '设计流量'])

    const dn = fields.find((f) => f.key === '设计流量')!
    expect(dn.kind).toBe('definedName')
    expect(dn.sheet).toBe('工程量表')
    // definedName 的 value 是原始公式串，带工作表前缀与绝对引用符号
    expect(dn.ref).toBe('工程量表!$B$2')
    expect(dn.ref.replace(/\$/g, '')).toContain('B2')

    const tb = fields.find((f) => f.key === '材料表')!
    expect(tb.kind).toBe('table')
    expect(tb.columns).toEqual(['名称', '数量', '单价'])
  })

  it('读回态的工作表也要能遍历（回归：不能直接用 wb.sheets）', async () => {
    const wb = await loadWorkbook(fromBuffer(await makeXlsx()))
    // 若错误地走 wb.sheets，Table 会静默丢失
    const hasTable = discoverXlsxFields(wb).some((f) => f.kind === 'table')
    expect(hasTable).toBe(true)
  })
})

describe('Excel 侧绑定', () => {
  it('命名区域写入标量', async () => {
    const xlsx = await makeXlsx()
    const t = new ParamTable(SPECS, CTX)
    const r = await bindXlsxFromParamTable(xlsx, t)
    expect(r.warnings).toHaveLength(0)

    const wb = await loadWorkbook(fromBuffer(r.buffer))
    const ws = getSheet(wb, '工程量表')!
    expect(getCellByCoord(ws, 'B1')?.value).toBe('海南陵水燃气工程')
    expect(getCellByCoord(ws, 'B2')?.value).toBe(1200)
  })

  it('Table 写入数组并扩展 Table 范围', async () => {
    const xlsx = await makeXlsx()
    const t = new ParamTable(SPECS, CTX)
    const r = await bindXlsxFromParamTable(xlsx, t)

    const wb = await loadWorkbook(fromBuffer(r.buffer))
    const ws = getSheet(wb, '工程量表')!
    expect(getCellByCoord(ws, 'A5')?.value).toBe('钢管 DN200')
    expect(getCellByCoord(ws, 'B5')?.value).toBe(100)
    expect(getCellByCoord(ws, 'C5')?.value).toBe(25.5)
    expect(getCellByCoord(ws, 'A6')?.value).toBe('钢管 DN100')
    // 第 3 行超出了模板原来的 A4:C5 范围，应被写入并纳入 Table
    expect(getCellByCoord(ws, 'A7')?.value).toBe('阀门 DN50')
  })

  it('参数表里没有的域会给出警告而不是静默失败', async () => {
    const xlsx = await makeXlsx()
    const t = new ParamTable([{ key: '工程名称', default: 'X' }], CTX) // 缺 设计流量 / 材料表
    const r = await bindXlsxFromParamTable(xlsx, t)
    expect(r.warnings.length).toBeGreaterThan(0)
    expect(r.warnings.some((w) => w.includes('设计流量'))).toBe(true)
  })
})

describe('跨文档：Word ← Excel', () => {
  it('Word 的域取 Excel 命名区域的值（外部文件引用）', async () => {
    // 1) 先把 Excel 填好
    const xlsx = await makeXlsx()
    const filled = await bindXlsxFromParamTable(xlsx, new ParamTable(SPECS, CTX))
    const wb = await loadWorkbook(fromBuffer(filled.buffer))
    const ws = getSheet(wb, '工程量表')!

    // 2) 造 Word 模板，域的来源指向那个 Excel 的命名区域
    const docx = await sdtifyDocx(await makeDocx())
    const wordTable = new ParamTable(
      [
        { key: '工程名称', source: 'file://D:/工程量.xlsx#工程名称' },
        { key: '设计流量', source: 'file://D:/工程量.xlsx#设计流量' },
      ],
      {
        // 平台适配层预加载：把 Excel 摊平成 { sheet名: { 单元格: 值 } } 并登记命名区域
        files: {
          'D:/工程量.xlsx': {
            工程量表: { B1: getCellByCoord(ws, 'B1')?.value, B2: getCellByCoord(ws, 'B2')?.value },
            工程名称: getCellByCoord(ws, 'B1')?.value,
            设计流量: getCellByCoord(ws, 'B2')?.value,
          },
        },
      }
    )
    const wr = await bindFromParamTable(docx.buffer, wordTable)

    const z = await JSZip.loadAsync(wr.buffer)
    const xml = await z.file('word/document.xml')!.async('string')
    const texts = [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => m[1])
    expect(texts).toContain('海南陵水燃气工程')
    expect(texts).toContain('1200')
    expect(wr.traces.every((t) => t.from === 'file')).toBe(true)
  })
})

describe('一张参数表同时喂两端', () => {
  it('Word 与 Excel 拿到一致的值', async () => {
    const table = new ParamTable(SPECS, CTX)

    const xlsxOut = await bindXlsxFromParamTable(await makeXlsx(), table)
    const docxOut = await bindFromParamTable((await sdtifyDocx(await makeDocx())).buffer, table)

    const wb = await loadWorkbook(fromBuffer(xlsxOut.buffer))
    const ws = getSheet(wb, '工程量表')!
    const z = await JSZip.loadAsync(docxOut.buffer)
    const xml = await z.file('word/document.xml')!.async('string')
    const texts = [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => m[1])

    const xFlow = getCellByCoord(ws, 'B2')?.value
    const wFlow = texts[texts.indexOf('设计流量：') + 1]
    expect(String(xFlow)).toBe(String(wFlow))
    expect(String(xFlow)).toBe('1200')
  })
})

/* ───────────── Table 计算列：公式双写 ───────────── */

/** 带「合价」计算列的模板 */
async function makeXlsxWithCalcCol(): Promise<Uint8Array> {
  const wb = createWorkbook()
  const ws = addWorksheet(wb, '工程量表')
  setCellByCoord(ws, 'A4', '名称')
  setCellByCoord(ws, 'B4', '数量')
  setCellByCoord(ws, 'C4', '单价')
  setCellByCoord(ws, 'D4', '合价')
  addExcelTable(wb, ws, { name: '材料表', ref: 'A4:D5', columns: ['名称', '数量', '单价', '合价'] })
  return workbookToBytes(wb, { mtime: new Date('2026-01-01T00:00:00Z') })
}

const ROWS = [
  { 名称: '钢管 DN200', 数量: 100, 单价: 25.5, 合价: 2550 },
  { 名称: '钢管 DN100', 数量: 40, 单价: 18.2, 合价: 728 },
  { 名称: '阀门 DN50', 数量: 12, 单价: 260, 合价: 3120 },
]

describe('Table 计算列（公式 + 缓存值双写）', () => {
  it('写成真公式且行号正确平移，同时带缓存值', async () => {
    const t = new ParamTable([{ key: '材料表', source: 'node://table1#rows' }], {
      nodes: { table1: { rows: ROWS } },
    })
    const r = await bindXlsxFromParamTable(await makeXlsxWithCalcCol(), t, {
      tableFormulas: { 材料表: { 合价: '={数量}*{单价}' } },
    })
    expect(r.warnings).toHaveLength(0)

    const ws = getSheet(await loadWorkbook(fromBuffer(r.buffer)), '工程量表')!
    const d5 = getCellByCoord(ws, 'D5')!
    expect(isFormulaCell(d5)).toBe(true)
    expect(getFormulaText(d5)).toBe('B5*C5')          // 第 5 行 → 引用 B5/C5
    expect(getCachedFormulaValue(d5)).toBe(2550)      // 缓存值，离线可看

    const d6 = getCellByCoord(ws, 'D6')!
    expect(getFormulaText(d6)).toBe('B6*C6')          // 行号跟着平移
    expect(getCachedFormulaValue(d6)).toBe(728)

    const d7 = getCellByCoord(ws, 'D7')!
    expect(getFormulaText(d7)).toBe('B7*C7')
    expect(getCachedFormulaValue(d7)).toBe(3120)
  })

  it('没有缓存值时仍写公式（交给 Excel 重算，不静默丢弃）', async () => {
    const rows = [{ 名称: '钢管', 数量: 100, 单价: 25.5 }] // 不带合价
    const t = new ParamTable([{ key: '材料表', source: 'node://table1#rows' }], {
      nodes: { table1: { rows } },
    })
    const r = await bindXlsxFromParamTable(await makeXlsxWithCalcCol(), t, {
      tableFormulas: { 材料表: { 合价: '={数量}*{单价}' } },
    })
    const ws = getSheet(await loadWorkbook(fromBuffer(r.buffer)), '工程量表')!
    const d5 = getCellByCoord(ws, 'D5')!
    expect(isFormulaCell(d5)).toBe(true)
    expect(getFormulaText(d5)).toBe('B5*C5')
  })

  it('公式列名不在 Table 列里时给出警告', async () => {
    const t = new ParamTable([{ key: '材料表', source: 'node://table1#rows' }], {
      nodes: { table1: { rows: ROWS } },
    })
    const r = await bindXlsxFromParamTable(await makeXlsxWithCalcCol(), t, {
      tableFormulas: { 材料表: { 税率: '={单价}*0.13' } },
    })
    expect(r.warnings.some((w) => w.includes('税率'))).toBe(true)
  })
})
