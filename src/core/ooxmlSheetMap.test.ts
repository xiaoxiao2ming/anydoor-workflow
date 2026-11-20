/**
 * 工作表名 → 部件路径 映射测试
 *
 * 立档：2026-10-29 · 配套 src/core/ooxmlSheetMap.ts
 *
 * ★ 本文件用**真实 xlsx 样本**做主用例（sample-template/ 下有真实文件），
 *   另加人工构造的乱序 rId 场景 —— 后者才能证明「按序号猜」是错的。
 */

import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { buildSheetPartMap, resolveSheetWrite } from './ooxmlSheetMap'

const SAMPLE = 'sample-template/excel-e2e-out.xlsx'   // 真实双表样本

/** 人工构造：rId 编号与 sheetN **故意错位**（真实文件里完全可能这样） */
async function makeShuffledXlsx (): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<Types/>')
  zip.file('xl/workbook.xml',
    '<workbook><sheets>' +
    '<sheet name="汇总表" sheetId="1" r:id="rId7"/>' +
    '<sheet name="工程量表" sheetId="2" r:id="rId2"/>' +
    '</sheets></workbook>')
  zip.file('xl/_rels/workbook.xml.rels',
    '<Relationships>' +
    // ★ rId2 → sheet2；rId7 → sheet1（完全乱序）
    '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
    '</Relationships>')
  zip.file('xl/worksheets/sheet1.xml', '<worksheet><sheetData/></worksheet>')
  zip.file('xl/worksheets/sheet2.xml', '<worksheet><sheetData/></worksheet>')
  return zip.generateAsync({ type: 'uint8array' })
}

describe('★★① 真实文件：表名 → 部件路径', () => {
  it('从真实 xlsx 解析出映射', async () => {
    const bytes = readFileSync(join(process.cwd(), SAMPLE))
    const m = await buildSheetPartMap(new Uint8Array(bytes))
    // 该样本是「汇总表 / 工程量表」双表
    // ★ 用无序比较：Array.sort() 按码位排，中文顺序与字面量顺序不一定一致
    //   （踩过一次：期望 ['汇总表','工程量表']，实际 sort 后是 ['工程量表','汇总表']，
    //     内容完全对却报红 —— 断言写错，不是实现错。）
    expect(Object.keys(m).sort()).toEqual(['工程量表', '汇总表'])
    expect(m['汇总表']).toBe('xl/worksheets/sheet1.xml')
    expect(m['工程量表']).toBe('xl/worksheets/sheet2.xml')
  })

  it('★ 顺序不能靠猜：映射必须来自 rels，而不是「第 N 张表」', async () => {
    const bytes = await makeShuffledXlsx()
    const m = await buildSheetPartMap(bytes)
    // ★ 汇总表是**第一张** sheet，但 r:id="rId7" ⇒ 它在 sheet1
    //   工程量表是**第二张** sheet，但 r:id="rId2"   ⇒ 它在 sheet2
    // 本例恰好与序号一致 ⇒ 换个数据就能错，故下面再构造真正错位的
    expect(m['汇总表']).toBe('xl/worksheets/sheet1.xml')
    expect(m['工程量表']).toBe('xl/worksheets/sheet2.xml')
  })

  it('★ 真正错位时，按映射走才对（按序号会写错表）', async () => {
    // 手工构造：汇总表在**第二**个部件，工程量表在**第一**个部件
    const zip = new JSZip()
    zip.file('[Content_Types].xml', '<Types/>')
    zip.file('xl/workbook.xml',
      '<workbook><sheets>' +
      '<sheet name="汇总表" sheetId="1" r:id="rId1"/>' +
      '<sheet name="工程量表" sheetId="2" r:id="rId5"/>' +
      '</sheets></workbook>')
    zip.file('xl/_rels/workbook.xml.rels',
      '<Relationships>' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
      '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>')
    zip.file('xl/worksheets/sheet1.xml', '<worksheet/>')
    zip.file('xl/worksheets/sheet2.xml', '<worksheet/>')
    const m = await buildSheetPartMap(await zip.generateAsync({ type: 'uint8array' }))
    // ★ 汇总表⇒ sheet2（不是 sheet1！按序号猜就会写错）
    expect(m['汇总表']).toBe('xl/worksheets/sheet2.xml')
    expect(m['工程量表']).toBe('xl/worksheets/sheet1.xml')
  })
})

describe('② 边界：缺文件/损坏时如实返回空，不猜', () => {
  it('不是 zip ⇒ 空映射', async () => {
    const m = await buildSheetPartMap(new Uint8Array([1, 2, 3, 4]))
    expect(m).toEqual({})
  })
  it('缺 workbook.xml ⇒ 空映射', async () => {
    const zip = new JSZip()
    zip.file('hello.txt', 'x')
    expect(await buildSheetPartMap(await zip.generateAsync({ type: 'uint8array' }))).toEqual({})
  })
  it('缺 rels ⇒ 空映射（单跳不够）', async () => {
    const zip = new JSZip()
    zip.file('xl/workbook.xml', '<workbook><sheets><sheet name="A" r:id="rId1"/></sheets></workbook>')
    expect(await buildSheetPartMap(await zip.generateAsync({ type: 'uint8array' }))).toEqual({})
  })
})

describe('③ 表名含 XML 实体时能正确反转义', () => {
  it('&amp; 等被还原', async () => {
    const zip = new JSZip()
    zip.file('[Content_Types].xml', '<Types/>')
    zip.file('xl/workbook.xml',
      '<workbook><sheets><sheet name="A&amp;B" sheetId="1" r:id="rId1"/></sheets></workbook>')
    zip.file('xl/_rels/workbook.xml.rels',
      '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
    zip.file('xl/worksheets/sheet1.xml', '<worksheet/>')
    const m = await buildSheetPartMap(await zip.generateAsync({ type: 'uint8array' }))
    expect(Object.keys(m)).toEqual(['A&B'])
  })
})

describe('★★ ④ 写前必须确认落点：找不到表就**不写**', () => {
  it('表名在映射里 ⇒ 给出部件', () => {
    const r = resolveSheetWrite({ 汇总表: 'xl/worksheets/sheet1.xml' }, '汇总表')
    expect(r).toEqual({ part: 'xl/worksheets/sheet1.xml' })
  })
  it('★ 表名不在映射里 ⇒ part:null + 说清现有表名（**不猜**）', () => {
    const r = resolveSheetWrite({ 汇总表: 'xl/worksheets/sheet1.xml' }, '不存在的表')
    expect(r.part).toBeNull()
    expect((r as { why: string }).why).toContain('不存在的表')
    expect((r as { why: string }).why).toContain('汇总表')   // 列出现有的帮用户判断
  })
  it('映射为空 ⇒ 如实说解析不到', () => {
    const r = resolveSheetWrite({}, '汇总表')
    expect(r.part).toBeNull()
    expect((r as { why: string }).why).toContain('解析不到')
  })
})
