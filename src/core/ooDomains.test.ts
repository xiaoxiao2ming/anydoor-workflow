import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import {
  domainRawKey, domainNameOf, domainRowOf, buildDomainEntries,
  injectDomains, readDomains, buildDomainSheetXml, nextRelId, nextSheetPart, DOMAIN_SHEET,
  patchFormulaCaches, recalcSheetCaches, inlineDomainNames, parseSharedStrings,
  diffDomainValues,
  refColorOf, findRefCells, addRefFills, markRefCells, unmarkRefCells, REF_PALETTE,
  extractRefMarks, addrToRowCol, colToLetters, readSheetGeom,
} from './ooDomains'

/** 造一份「提取通道产物」形态的 xlsx：单表 Sheet1 + 一个引用公式格（无定义名称） */
async function makeMinimalXlsx (): Promise<Uint8Array> {
  const z = new JSZip()
  z.file('[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    + '</Types>')
  z.file('_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    + '</Relationships>')
  z.file('xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')
  z.file('xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    + '</Relationships>')
  z.file('xl/worksheets/sheet1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
    + '<row r="5"><c r="D5"><f>数据对照表_键</f><v>456</v></c></row>'
    + '</sheetData></worksheet>')
  return z.generateAsync({ type: 'uint8array' })
}

describe('域数据键口径（必须与 oo-core 逐位一致）', () => {
  it('rawKey = 节点名.字段名', () => {
    expect(domainRawKey('数据对照表', '键')).toBe('数据对照表.键')
    expect(domainRawKey('材料库', '')).toBe('材料库')
  })

  it('定义名称把 . 换成 _，中文保留', () => {
    expect(domainNameOf('数据对照表.键')).toBe('数据对照表_键')
    expect(domainNameOf('材料库.unit')).toBe('材料库_unit')
    // 首字符是数字/下划线 ⇒ 前缀 R
    expect(domainNameOf('1号.值')).toBe('R1号_值')
  })

  it('行号是 djb2 落 2..1999，且同一 key 恒定', () => {
    const r1 = domainRowOf('数据对照表.键')
    expect(r1).toBeGreaterThanOrEqual(2)
    expect(r1).toBeLessThanOrEqual(1999)
    expect(domainRowOf('数据对照表.键')).toBe(r1)
    expect(domainRowOf('数据对照表.值')).not.toBe(R1(r1))
  })

  it('buildDomainEntries 去重 + 值格式化', () => {
    const e = buildDomainEntries([
      { nodeName: '数据对照表', fieldName: '键', value: '222' },
      { nodeName: '数据对照表', fieldName: '值', value: 111 },
      { nodeName: '空节点', fieldName: 'x', value: null },
    ])
    expect(e).toHaveLength(3)
    const k = e.find((x) => x.rawKey === '数据对照表.键')!
    expect(k.value).toBe('222')
    const v = e.find((x) => x.rawKey === '数据对照表.值')!
    expect(v.value).toBe('111')
    expect(e.find((x) => x.rawKey === '空节点.x')!.value).toBe('')
  })
})

describe('injectDomains：给没有域的 xlsx 补回「域数据 + 定义名称」', () => {
  it('新建隐藏表 + 定义名称 + fullCalcOnLoad，且不动引用格公式', async () => {
    const src = await makeMinimalXlsx()
    const entries = buildDomainEntries([{ nodeName: '数据对照表', fieldName: '键', value: '222' }])
    const r = await injectDomains(src, entries)
    expect(r.injected).toBe(1)
    expect(r.createdSheet).toBe(true)

    const z = await JSZip.loadAsync(r.bytes)
    const wb = await z.file('xl/workbook.xml')!.async('string')
    // 表存在且隐藏
    expect(wb).toContain(`<sheet name="${DOMAIN_SHEET}"`)
    expect(wb).toContain('state="hidden"')
    // 定义名称指向域数据表
    expect(wb).toContain('<definedName name="数据对照表_键">' + `'${DOMAIN_SHEET}'!$A$${domainRowOf('数据对照表.键')}`)
    // 打开时重算
    // ★ 2026-10-18 深夜改版：不再写 fullCalcOnLoad——本机引擎不理会它，带标志反而让
    //   引擎把刚打开的文档标成「已修改」（没改任何内容、关闭却提示保存）。
    //   缓存正确性由 recalcSheetCaches 写回保证（唯一事实源）。
    expect(wb).not.toContain('fullCalcOnLoad')
    // 引用格公式原样在（补丁不改公式）
    const s1 = await z.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(s1).toContain('<f>数据对照表_键</f>')
    // 新部件登记进 rels + Content_Types
    const rel = await z.file('xl/_rels/workbook.xml.rels')!.async('string')
    expect(rel).toContain('worksheets/sheet2.xml')
    const ct = await z.file('[Content_Types].xml')!.async('string')
    expect(ct).toContain('/xl/worksheets/sheet2.xml')
  })

  it('值是文本时用 inlineStr（不引 sharedStrings）', async () => {
    const src = await makeMinimalXlsx()
    const entries = buildDomainEntries([{ nodeName: '文本', fieldName: '文本', value: '甲乙丙' }])
    const r = await injectDomains(src, entries)
    const dom = await readDomains(r.bytes)
    expect(dom).toEqual([{ row: domainRowOf('文本.文本'), key: '文本.文本', value: '甲乙丙' }])
  })

  it('幂等：连续注入两次不会多出第二张域数据表', async () => {
    const src = await makeMinimalXlsx()
    const e1 = buildDomainEntries([{ nodeName: '数据对照表', fieldName: '键', value: '222' }])
    const a = await injectDomains(src, e1)
    const b = await injectDomains(a.bytes, e1)
    const z = await JSZip.loadAsync(b.bytes)
    const wb = await z.file('xl/workbook.xml')!.async('string')
    expect(wb.match(new RegExp(`name="${DOMAIN_SHEET}"`, 'g'))).toHaveLength(1)
    expect(wb.match(/<definedName /g)).toHaveLength(1)
    expect(b.createdSheet).toBe(false)
    expect(await readDomains(b.bytes)).toHaveLength(1)
  })

  it('值更新：第二次注入把 A 列改成新值', async () => {
    const src = await makeMinimalXlsx()
    const a = await injectDomains(src, buildDomainEntries([{ nodeName: '数据对照表', fieldName: '键', value: '222' }]))
    const b = await injectDomains(a.bytes, buildDomainEntries([{ nodeName: '数据对照表', fieldName: '键', value: '999' }]))
    const dom = await readDomains(b.bytes)
    expect(dom[0].value).toBe('999')
  })

  it('★ 回归（2026-10-18 事故）：引用格公式缓存必须写回 zip——引擎不理会 fullCalcOnLoad，缓存旧=显示旧', async () => {
    /* 事故：着色改版把 zip.file(name, pr.xml) 写回弄丢，重算只在内存里发生，
     * 落盘文件 A1 缓存仍是 456 ⇒ 用户「点了更新表格不变」。留档铁证：
     * 2026-10-18-20-10-18（域数据=8/9 正确、A1 缓存=222 旧值）。 */
    const src = await makeMinimalXlsx() // D5: =数据对照表_键，缓存 456
    const r = await injectDomains(src, buildDomainEntries([{ nodeName: '数据对照表', fieldName: '键', value: '222' }]))
    expect(r.cachePatched).toBeGreaterThan(0)
    const z = await JSZip.loadAsync(r.bytes)
    const s1 = await z.file('xl/worksheets/sheet1.xml')!.async('string')
    const d5 = /<c r="D5"[^>]*>[\s\S]*?<\/c>/.exec(s1)?.[0] ?? ''
    expect(d5).toContain('<v>222</v>')
    expect(d5).not.toContain('456')
  })

  it('没有引用值 ⇒ 原样返回，不建表', async () => {
    const src = await makeMinimalXlsx()
    const r = await injectDomains(src, [])
    expect(r.injected).toBe(0)
    expect(r.bytes).toBe(src)
    expect(r.why).toBeTruthy()
  })

  it('非 zip 字节 ⇒ 如实返回原因，不抛', async () => {
    const r = await injectDomains(new Uint8Array([1, 2, 3]), buildDomainEntries([{ nodeName: 'a', fieldName: 'b', value: 1 }]))
    expect(r.injected).toBe(0)
    expect(r.why).toContain('zip')
  })

  it('域数据表已存在时重建内容（旧 key 被清掉）', async () => {
    const src = await makeMinimalXlsx()
    const a = await injectDomains(src, buildDomainEntries([
      { nodeName: '旧节点', fieldName: 'x', value: 1 },
      { nodeName: '数据对照表', fieldName: '键', value: '222' },
    ]))
    const b = await injectDomains(a.bytes, buildDomainEntries([{ nodeName: '数据对照表', fieldName: '键', value: '222' }]))
    const dom = await readDomains(b.bytes)
    expect(dom.map((d) => d.key)).toEqual(['数据对照表.键'])
  })
})

describe('diffDomainValues：表格显示值 vs 画布当前值（打开时检测的判据）', () => {
  it('★ 域表缺失但公式格缓存=当前值 → 不算变化（修复「没改参数也提示」）', async () => {
    /* 旧口径：域表被保存通道剥掉 → 全部当「有变化」→ 参数没改也弹窗。
     * 新口径：直接比对公式格缓存（用户看到的数）。 */
    const src = await makeMinimalXlsx() // D5: =数据对照表_键，缓存 456
    const z = await JSZip.loadAsync(src)
    await z.file('xl/worksheets/sheet1.xml')!.async('string').then((xml) => {
      // 把 D5 缓存改成 222（= 当前值），模拟「没改参数」的文件
      return z.file('xl/worksheets/sheet1.xml', xml.replace('<v>456</v>', '<v>222</v>'))
    })
    const patched = await z.generateAsync({ type: 'uint8array' })
    const entries = buildDomainEntries([{ nodeName: '数据对照表', fieldName: '键', value: '222' }])
    const d = await diffDomainValues(patched, entries)
    expect(d.hasSheet).toBe(false)
    expect(d.changed).toHaveLength(0)
  })

  it('公式格缓存过期 → 报出名字与新旧值（was=用户看到的，now=画布最新的）', async () => {
    const src = await makeMinimalXlsx() // D5 缓存 456
    const entries = buildDomainEntries([{ nodeName: '数据对照表', fieldName: '键', value: '222' }])
    const d = await diffDomainValues(src, entries)
    expect(d.changed).toEqual([{ name: '数据对照表_键', was: '456', now: '222' }])
    expect(d.total).toBe(1)
  })

  it('值一致 → 无变化；XML 特殊字符经反转义后比对不误报', async () => {
    const src = await makeMinimalXlsx()
    const key = [{ nodeName: '数据对照表', fieldName: '键', value: '222' }]
    const fresh = await injectDomains(src, buildDomainEntries(key))
    const d1 = await diffDomainValues(fresh.bytes, buildDomainEntries(key))
    expect(d1.hasSheet).toBe(true)
    expect(d1.changed).toHaveLength(0)
    const d2 = await diffDomainValues(fresh.bytes, buildDomainEntries([{ nodeName: '数据对照表', fieldName: '键', value: '333' }]))
    expect(d2.changed).toEqual([{ name: '数据对照表_键', was: '222', now: '333' }])
    const spec = await injectDomains(src, buildDomainEntries([{ nodeName: '工程', fieldName: '名称', value: 'A&B<管>' }]))
    const d3 = await diffDomainValues(spec.bytes, buildDomainEntries([{ nodeName: '工程', fieldName: '名称', value: 'A&B<管>' }]))
    expect(d3.changed).toHaveLength(0)
  })

  it('空 entries → 直接无变化（没有引用就没有「更新」可言）', async () => {
    const src = await makeMinimalXlsx()
    const d = await diffDomainValues(src, [])
    expect(d.total).toBe(0)
    expect(d.changed).toHaveLength(0)
  })
})

describe('patchFormulaCaches：引用格缓存值跟着走', () => {
  it('把 <v>456</v> 改成新值，公式与属性保留', () => {
    const xml = '<sheetData><row r="5"><c r="D5" s="3"><f>数据对照表_键</f><v>456</v></c></row></sheetData>'
    const r = patchFormulaCaches(xml, buildDomainEntries([{ nodeName: '数据对照表', fieldName: '键', value: '222' }]))
    expect(r.patched).toBe(1)
    expect(r.xml).toContain('<f>数据对照表_键</f><v>222</v>')
    expect(r.xml).toContain('s="3"')          // 样式索引不动
  })

  it('文本值加 t="str"', () => {
    const xml = '<row r="2"><c r="A2"><f>文本_文本</f><v>旧</v></c></row>'
    const r = patchFormulaCaches(xml, buildDomainEntries([{ nodeName: '文本', fieldName: '文本', value: '甲乙' }]))
    expect(r.patched).toBe(1)
    expect(r.xml).toContain('t="str"')
    expect(r.xml).toContain('<v>甲乙</v>')
  })

  it('非引用格 / 不认识的公式不动', () => {
    const xml = '<row r="1"><c r="A1"><f>SUM(B1:B3)</f><v>9</v></c><c r="A2"><v>5</v></c></row>'
    const r = patchFormulaCaches(xml, buildDomainEntries([{ nodeName: '数据对照表', fieldName: '键', value: '222' }]))
    expect(r.patched).toBe(0)
    expect(r.xml).toBe(xml)
  })
})

describe('工作表/关系编号选择', () => {
  it('nextSheetPart 避开已占用编号', () => {
    expect(nextSheetPart(['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet3.xml'])).toBe('xl/worksheets/sheet4.xml')
    expect(nextSheetPart([])).toBe('xl/worksheets/sheet1.xml')
  })
  it('nextRelId 取最大 rId + 1', () => {
    expect(nextRelId('<Relationship Id="rId1"/><Relationship Id="rId7"/>')).toBe('rId8')
  })
})

/** 断言用：不等于给定值（保持简单的意图表达） */
function R1 (n: number): number { return n }

/** buildDomainSheetXml 本身要能产出两条列（A=值，B=键） */
describe('buildDomainSheetXml', () => {
  it('同时写 A（值）与 B（键）', () => {
    const xml = buildDomainSheetXml([{ rawKey: 'a.b', name: 'a_b', row: 7, value: '3' }])
    expect(xml).toContain('<c r="A7"><v>3</v></c>')
    expect(xml).toContain('<c r="B7" t="inlineStr"><is><t xml:space="preserve">a.b</t></is></c>')
  })
})


/* ── 整表重算：用户三条反馈的验收（2026-10-17） ──────────────────
 * ① `=参数`            → 跟着变（原实现已覆盖）
 * ② `=参数*2`          → 也要变（公式带运算）
 * ③ `=A1*2`（A1 是引用格）→ 下游也要变
 * ④ 再往下 `=A2+1`     → 深链每层都要推进
 */
describe('recalcSheetCaches：引用参与公式/下游引用', () => {
  const entries = buildDomainEntries([{ nodeName: '数据对照表', fieldName: '键', value: '222' }])
  const name = entries[0].name

  it('① 公式就是定义名称：缓存变 222', () => {
    const xml = `<row r="1"><c r="A1"><f>${name}</f><v>456</v></c></row>`
    const r = recalcSheetCaches(xml, entries)
    expect(r.patched).toBe(1)
    expect(r.xml).toContain('<f>数据对照表_键</f><v>222</v>')
  })

  it('② 定义名称参与运算：=参数*2 → 444', () => {
    const xml = `<row r="2"><c r="B2"><f>${name}*2</f><v>912</v></c></row>`
    const r = recalcSheetCaches(xml, entries)
    expect(r.patched).toBe(1)
    expect(r.xml).toContain('<v>444</v>')
  })

  it('③ 下游引用引用格：A1=参数，B1=A1*2 → 444', () => {
    const xml = `<row r="1"><c r="A1"><f>${name}</f><v>456</v></c>`
      + '<c r="B1"><f>A1*2</f><v>912</v></c></row>'
    const r = recalcSheetCaches(xml, entries)
    expect(r.patched).toBe(2)
    expect(r.xml).toContain('<f>A1*2</f><v>444</v>')
  })

  it('④ 深链：A1=参数 → B1=A1*2 → C1=B1+1 = 445', () => {
    const xml = `<row r="1"><c r="A1"><f>${name}</f><v>0</v></c>`
      + '<c r="B1"><f>A1*2</f><v>0</v></c>'
      + '<c r="C1"><f>B1+1</f><v>0</v></c></row>'
    const r = recalcSheetCaches(xml, entries)
    expect(r.xml).toContain('<f>B1+1</f><v>445</v>')
  })

  it('⑤ 区间与函数：SUM(A1:A2)*2', () => {
    const xml = `<row r="1"><c r="A1"><f>${name}</f><v>456</v></c></row>`
      + '<row r="2"><c r="A2"><v>3</v></c></row>'
      + '<row r="3"><c r="A3"><f>SUM(A1:A2)*2</f><v>0</v></c></row>'
    const r = recalcSheetCaches(xml, entries)
    expect(r.xml).toContain('<v>450</v>')   // (222+3)*2
  })

  it('⑥ 文本域参与连接：t="str"', () => {
    const txt = buildDomainEntries([{ nodeName: '参数', fieldName: '名称', value: '甲' }])
    const xml = `<row r="1"><c r="A1"><f>${txt[0].name}&amp;"-乙"</f><v>旧</v></c></row>`
    const r = recalcSheetCaches(xml, txt)
    expect(r.patched).toBe(1)
    expect(r.xml).toContain('t="str"')
    expect(r.xml).toContain('<v>甲-乙</v>')
  })

  it('⑦ 缓存已正确 ⇒ 不动（保持字节稳定）', () => {
    const xml = `<row r="1"><c r="A1"><f>${name}</f><v>222</v></c></row>`
    const r = recalcSheetCaches(xml, entries)
    expect(r.patched).toBe(0)
    expect(r.xml).toBe(xml)
  })

  it('⑧ 不支持/坏公式 ⇒ 保留原缓存并报明细', () => {
    const xml = `<row r="1"><c r="A1"><f>${name}</f><v>456</v></c><c r="B1"><f>VLOOKUP(A1,"x",2,0)+WEIRD(1)</f><v>7</v></c></row>`
    const r = recalcSheetCaches(xml, entries)
    // 坏公式原样保留 + 旧缓存不动（7 还在），只在 failed 里报出来
    expect(r.xml).toContain('<f>VLOOKUP(A1,"x",2,0)+WEIRD(1)</f><v>7</v>')
    expect(r.xml).toContain('<f>数据对照表_键</f><v>222</v>')
    expect(r.failed.length).toBeGreaterThan(0)
  })

  it('⑨ 名称边界：参数 不应命中 参数2 / 函数名', () => {
    const e = buildDomainEntries([{ nodeName: '参数', fieldName: '数量', value: '5' }])
    const nm = e[0].name
    expect(inlineDomainNames(`${nm}+${nm}2`, new Map([[nm, '5']]))).toBe('5+参数_数量2')
    expect(inlineDomainNames(`${nm}*2`, new Map([[nm, '5']]))).toBe('5*2')
  })

  it('⑩ 文本单元格走 sharedStrings（索引不能当数字）', () => {
    const sst = parseSharedStrings('<sst><si><t>甲乙</t></si></sst>')
    expect(sst).toEqual(['甲乙'])
    const e = buildDomainEntries([{ nodeName: '参数', fieldName: '名称', value: '丙' }])
    const xml = `<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><f>A1&amp;${e[0].name}</f><v>旧</v></c></row>`
    const r = recalcSheetCaches(xml, e, sst)
    expect(r.xml).toContain('<v>甲乙丙</v>')
  })
})

/* ── 引用格着色（用户 2026-10-17：「能不能对参数引用的地址加个颜色区分」） ── */
describe('引用格着色：一眼看出「这格是引用来的」', () => {
  const entries = buildDomainEntries([
    { nodeName: '数据对照表', fieldName: '键', value: '222' },
    { nodeName: '参数', fieldName: '数量', value: '5' },
  ])
  const n1 = entries[0].name
  const n2 = entries[1].name

  it('同一名称永远同色（重启/重开会话颜色不变）', () => {
    expect(refColorOf(n1)).toBe(refColorOf(n1))
    expect(REF_PALETTE).toContain(refColorOf(n1))
    expect(refColorOf(n2)).toBe(refColorOf(n2))
    expect(refColorOf(n1)).toMatch(/^FF[0-9A-F]{6}$/)
  })

  it('findRefCells：=参数 / =参数*2 / SUM(参数) 都算引用格', () => {
    const xml = '<row r="1">'
      + `<c r="A1"><f>${n1}</f><v>1</v></c>`
      + `<c r="B1"><f>${n1}*2</f><v>2</v></c>`
      + `<c r="C1"><f>SUM(${n2},A1)</f><v>3</v></c>`
      + '</row>'
    expect(findRefCells(xml, entries).map((h) => h.ref)).toEqual(['A1', 'B1', 'C1'])
  })

  it('findRefCells：名称边界不吃「参数2」；纯下游格不算引用格', () => {
    const xml = '<row r="1">'
      + `<c r="A1"><f>${n1}</f><v>1</v></c>`
      + `<c r="B1"><f>${n1}2</f><v>2</v></c>`
      + '<c r="C1"><f>B1*3</f><v>6</v></c>'
      + '</row>'
    expect(findRefCells(xml, entries).map((h) => h.ref)).toEqual(['A1'])
  })

  it('addRefFills：无 styles.xml 时生成最小件（含 none/gray125 两个必需 fill）', () => {
    const t = addRefFills(null, ['FFFFF2CC'])
    expect(t.created).toBe(true)
    expect(t.xml).toContain('patternType="none"')
    expect(t.xml).toContain('patternType="gray125"')
    expect(t.xml).toContain('FFFFF2CC')
    // none=0、gray125=1，新色从 2 开始
    expect(t.fillIdOf.get('FFFFF2CC')).toBe(2)
  })

  it('addRefFills：同色复用，二次调用零增长（反复打开不膨胀文件）', () => {
    const t1 = addRefFills(null, ['FFFFF2CC', 'FFCCE5FF'])
    const t2 = addRefFills(t1.xml, ['FFFFF2CC', 'FFCCE5FF'])
    expect(t2.addedFills).toBe(0)
    expect(t2.addedXfs).toBe(0)
    expect(t2.fillIdOf.get('FFFFF2CC')).toBe(t1.fillIdOf.get('FFFFF2CC'))
  })

  it('addRefFills：已有样式时只在末尾追加，原 count 被改写而不是叠加', () => {
    const old = '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<fonts count="1"><font><sz val="11"/></font></fonts>'
      + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
      + '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
      + '</styleSheet>'
    const t = addRefFills(old, ['FFFFF2CC'])
    expect(t.created).toBe(false)
    expect(t.xml).toMatch(/<fills[^>]*count="3"/)
    expect(t.xml).toMatch(/<cellXfs[^>]*count="2"/)
    expect(t.xml).toContain('numFmtId="0"')
  })

  it('markRefCells：只改 s，t 属性/公式/值/其它格一个不动', () => {
    const t = addRefFills(null, [refColorOf(n1)])
    const xml = '<row r="1">'
      + `<c r="A1" t="str"><f>${n1}</f><v>旧</v></c>`
      + '<c r="B1" s="0"><v>123</v></c>'
      + '</row>'
    const r = markRefCells(xml, [{ ref: 'A1', name: n1 }], t)
    expect(r.marked).toBe(1)
    expect(r.xml).toContain('t="str"')
    expect(r.xml).toContain(`<f>${n1}</f>`)
    expect(r.xml).toContain('<c r="B1" s="0"><v>123</v></c>')
    expect(r.xml).toMatch(/<c r="A1"[^>]*s="\d+"/)
  })

  it('markRefCells：自闭合空引用格也能加 s', () => {
    const t = addRefFills(null, [refColorOf(n1)])
    const r = markRefCells('<row r="1"><c r="A1"/></row>', [{ ref: 'A1', name: n1 }], t)
    expect(r.marked).toBe(1)
    expect(r.xml).toMatch(/<c r="A1" s="\d+"\/>/)
  })

  it('markRefCells：保留原格数字格式（s 指向「原 xf + 新底色」，不是从零新建）', () => {
    const old = '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
      + '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
      + '<xf numFmtId="10" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/></cellXfs>'
      + '</styleSheet>'
    const t = addRefFills(old, [refColorOf(n1)])
    const r = markRefCells('<row r="1"><c r="A1" s="1"><f>' + n1 + '</f><v>1</v></c></row>', [{ ref: 'A1', name: n1 }], t)
    const s = Number(/<c r="A1" s="(\d+)"/.exec(r.xml)?.[1])
    const xfs = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(t.xml)?.[1] ?? ''
    const tags = xfs.match(/<xf\b[^>]*\/>/g) ?? []
    expect(tags[s]).toContain('numFmtId="10"')   // 百分比格式保留
    expect(tags[s]).toContain('borderId="1"')
    expect(tags[s]).toContain('applyFill="1"')
  })

  it('端到端（2026-10-18 改版）：injectDomains 不再写底色，着色移交覆盖层', async () => {
    const src = await makeMinimalXlsx()   // D5 = 数据对照表_键
    const r = await injectDomains(src, entries)
    expect(r.markedCells).toBe(0)

    const z = await JSZip.loadAsync(r.bytes)
    const styles = await z.file('xl/styles.xml')?.async('string') ?? ''
    // 不把引用色写进 styles.xml（没有 styles.xml 也行——不再需要样式）
    expect(styles).not.toContain(refColorOf('数据对照表_键'))
    const sheet = await z.file('xl/worksheets/sheet1.xml')!.async('string')
    // D5 不带样式索引（引用格与普通格在文件里无差别）
    expect(sheet).not.toMatch(/<c r="D5"[^>]*\bs="/)
  })
})

/* ── 引用格覆盖层数据（2026-10-18 着色改版：extractRefMarks） ── */
describe('extractRefMarks：引用格清单 + 表格几何（纯读取，不改文件）', () => {
  const entries = buildDomainEntries([
    { nodeName: '数据对照表', fieldName: '键', value: '222' },
  ])

  it('addrToRowCol / colToLetters：地址与索引互转', () => {
    expect(addrToRowCol('D5')).toEqual({ row: 4, col: 3 })
    expect(addrToRowCol('$AB$12')).toEqual({ row: 11, col: 27 })
    expect(colToLetters(0)).toBe('A')
    expect(colToLetters(3)).toBe('D')
    expect(colToLetters(27)).toBe('AB')
    expect(addrToRowCol('垃圾')).toEqual({ row: -1, col: -1 })
  })

  it('injectDomains 产物上提取：marks 含命中格与颜色，geom 保留 xlsx 原始单位', async () => {
    const src = await makeMinimalXlsx()
    const r = await injectDomains(src, entries)
    const ex = await extractRefMarks(r.bytes, entries)
    expect(ex.marks.length).toBe(1)
    expect(ex.marks[0].ref).toBe('D5')
    expect(ex.marks[0].color).toBe(refColorOf('数据对照表_键'))
    expect(ex.geom).toBeTruthy()
    /* 列宽按字符、行高按 pt 原样带给注入层，由注入层用活动格实测像素标定比例 */
    expect(ex.geom!.defColCh).toBeGreaterThan(0)
    expect(ex.geom!.defRowPt).toBeGreaterThan(0)
    expect(Array.isArray(ex.geom!.merges)).toBe(true)
  })

  it('隐藏列/行记 0 宽高（否则其后所有列整体位移 → 用户看到的「亮显偏移」）', () => {
    const g = readSheetGeom(
      '<sheetFormatPr defaultRowHeight="14" defaultColWidth="9"/>' +
      '<cols><col min="1" max="1" width="9.5" customWidth="1"/>' +
      '<col min="3" max="3" width="11.6" hidden="1" customWidth="1"/>' +
      '<col min="5" max="7" width="8.69" customWidth="1"/></cols>' +
      '<sheetData><row r="2" ht="30" customHeight="1"><c r="A2"/></row>' +
      '<row r="3" hidden="1"><c r="A3"/></row></sheetData>',
    )
    expect(g.defColCh).toBe(9)
    expect(g.defRowPt).toBe(14)
    expect(g.colCh['0']).toBe(9.5)
    expect(g.colCh['2']).toBe(0)      /* 隐藏列 */
    expect(g.colCh['4']).toBe(8.69)   /* min/max 区间要展开 */
    expect(g.colCh['6']).toBe(8.69)
    expect(g.rowPt['1']).toBe(30)     /* 自定义行高按 pt */
    expect(g.rowPt['2']).toBe(0)      /* 隐藏行 */
  })

  it('合并区解析：色块要铺满整块合并单元格', () => {
    const g = readSheetGeom('<mergeCells count="1"><mergeCell ref="B2:D4"/></mergeCells>')
    expect(g.merges).toEqual([{ r1: 1, c1: 1, r2: 3, c2: 3 }])
  })

  it('原文件直接提取（未注入）：同样能命中', async () => {
    const src = await makeMinimalXlsx()
    const ex = await extractRefMarks(src, entries)
    expect(ex.marks.map((m) => m.ref)).toContain('D5')
  })

  it('非 xlsx 字节：安静返回空', async () => {
    const ex = await extractRefMarks(new Uint8Array([1, 2, 3]), entries)
    expect(ex.marks).toEqual([])
    expect(ex.geom).toBeNull()
  })
})
