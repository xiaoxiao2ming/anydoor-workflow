/**
 * 造一个"像设计院交付的"Excel 模板，用于端到端验证：
 *   · 汇总表：两个命名区域 设计流量 / 工程名称（标量域）
 *   · 工程量表：一个 Excel Table「材料表」，列为 名称/数量/单价/合价（行级域 + 计算列）
 *
 * 输出：sample-template/测试模板.xlsx
 */
import fs from 'node:fs'
import { createWorkbook, addWorksheet, addDefinedName } from '@office-kit/xlsx/workbook'
import { setCellByCoord, addExcelTable } from '@office-kit/xlsx/worksheet'
import { makeFormula } from '@office-kit/xlsx/cell'
import { workbookToBytes } from '@office-kit/xlsx/io'

const OUT = new URL('../sample-template/', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')

const wb = createWorkbook()

// ── 汇总表：放域 ────────────────────────────────────────────
const ws1 = addWorksheet(wb, '汇总表')
setCellByCoord(ws1, 'A1', '工程名称')
setCellByCoord(ws1, 'B1', '（待填）')
setCellByCoord(ws1, 'A2', '设计流量')
setCellByCoord(ws1, 'B2', 0)
addDefinedName(wb, { name: '工程名称', value: '汇总表!$B$1' })
addDefinedName(wb, { name: '设计流量', value: '汇总表!$B$2' })

// ── 工程量表：Table（行级域）────────────────────────────────
const ws2 = addWorksheet(wb, '工程量表')
const COLS = ['名称', '数量', '单价', '合价']
COLS.forEach((c, i) => setCellByCoord(ws2, String.fromCharCode(65 + i) + '1', c))
// 模板里留一行空数据行，Table 范围 A1:D2（表头 + 1 空行）
addExcelTable(wb, ws2, { name: '材料表', ref: 'A1:D2', columns: COLS })
// 合价列预置公式（生成时会按行重写并带缓存值）
setCellByCoord(ws2, 'D2', makeFormula('B2*C2', { cachedValue: 0 }))

const bytes = await workbookToBytes(wb, { mtime: new Date('2026-01-01T00:00:00Z') })
fs.writeFileSync(OUT + '测试模板.xlsx', bytes)
console.log('已生成 sample-template/测试模板.xlsx  ' + bytes.length + ' 字节')
console.log('  工作表: ' + [ws1.title, ws2.title].join(', '))
console.log('  命名区域: 工程名称(汇总表!$B$1), 设计流量(汇总表!$B$2)')
console.log('  表格: 材料表 A1:D2  列=' + COLS.join('/'))
