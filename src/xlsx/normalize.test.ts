import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { normalizeWorkbookBytes, NORMALIZE_PATCH_IDS } from './normalize'

/**
 * 归一化层测试。
 *
 * 这层的存在是实测逼出来的：@office-kit/xlsx@0.18.0 会拒收两种合法 OOXML 构造，
 * 而 WPS 生成的 190KB 气材料对照表里恰好都有（`<c t="s">` 无 `<v>` 出现 23291 次）。
 * 所以测试的重点不是"能修好"，而是"**只修该修的，绝不误伤**"。
 */

async function makeZip(files: Record<string, string>): Promise<Uint8Array> {
  const z = new JSZip()
  for (const [p, c] of Object.entries(files)) z.file(p, c)
  return z.generateAsync({ type: 'uint8array' })
}

async function readPart(bytes: Uint8Array, path: string): Promise<string> {
  const z = await JSZip.loadAsync(bytes)
  const f = z.file(path)
  if (!f) throw new Error('产物里没有 ' + path)
  return f.async('string')
}

const STYLES_WITH_EMPTY_FILL = `<?xml version="1.0"?><styleSheet>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill/></fills>
</styleSheet>`

const STYLES_OK = `<?xml version="1.0"?><styleSheet>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill></fills>
</styleSheet>`

const SHEET_MIXED = `<?xml version="1.0"?><worksheet><sheetData>
<row r="1"><c r="A1" t="s"><v>3</v></c><c r="B1" t="s" s="2"/><c r="C1" t="n" s="1"/></row>
<row r="2"><c r="A2" t="s" s="4"></c><c r="B2"><v>42</v></c></row>
<row r="3"><c r="A3" t="b" s="1"/><c r="B3" t="str" s="1"/><c r="C3" t="e" s="1"/></row>
</sheetData></worksheet>`

const SHEET_CLEAN = `<?xml version="1.0"?><worksheet><sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>7</v></c></row>
</sheetData></worksheet>`

describe('normalizeWorkbookBytes', () => {
  it('干净的 zip 不该被改动 —— 必须返回同一引用（零副作用）', async () => {
    const src = await makeZip({
      'xl/styles.xml': STYLES_OK,
      'xl/worksheets/sheet1.xml': SHEET_CLEAN,
    })
    const r = await normalizeWorkbookBytes(src)
    expect(r.applied).toEqual([])
    expect(r.bytes).toBe(src) // 同一引用：一次都没重新打包
  })

  it('补 <fill/>（office-kit: fillFromTree has no child element）', async () => {
    const src = await makeZip({
      'xl/styles.xml': STYLES_WITH_EMPTY_FILL,
      'xl/worksheets/sheet1.xml': SHEET_CLEAN,
    })
    const r = await normalizeWorkbookBytes(src)
    expect(r.applied.some((a) => a.startsWith('empty-fill'))).toBe(true)
    const styles = await readPart(r.bytes, 'xl/styles.xml')
    expect(styles).toContain('<fill><patternFill patternType="none"/></fill>')
    // 原有的两个 fill 不能被破坏
    expect((styles.match(/<fill>/g) ?? []).length).toBe(3)
    expect(styles).toContain('patternType="gray125"')
  })

  it('只动 t="s"，其它类型一律不碰（保守原则）', async () => {
    const src = await makeZip({
      'xl/styles.xml': STYLES_OK,
      'xl/worksheets/sheet1.xml': SHEET_MIXED,
    })
    const r = await normalizeWorkbookBytes(src)
    expect(r.applied.some((a) => a.startsWith('cell-type-without-value'))).toBe(true)
    const sheet = await readPart(r.bytes, 'xl/worksheets/sheet1.xml')
    // 自闭合的 → t 被去掉，样式 s 保留
    expect(sheet).toContain('<c r="B1" s="2"/>')
    // 成对标签的 → t 被去掉
    expect(sheet).toContain('<c r="A2" s="4"></c>')
    // t="n" 是默认类型；t="b"/"str"/"e" 有独立语义 —— 都不该被碰
    expect(sheet).toContain('<c r="C1" t="n" s="1"/>')
    expect(sheet).toContain('<c r="A3" t="b" s="1"/>')
    expect(sheet).toContain('<c r="B3" t="str" s="1"/>')
    expect(sheet).toContain('<c r="C3" t="e" s="1"/>')
  })

  it('❗绝不误伤：有 <v> 的单元格必须保持原样（改动=数据损坏）', async () => {
    const src = await makeZip({
      'xl/styles.xml': STYLES_OK,
      'xl/worksheets/sheet1.xml': SHEET_MIXED,
    })
    const r = await normalizeWorkbookBytes(src)
    const sheet = await readPart(r.bytes, 'xl/worksheets/sheet1.xml')
    // A1 是共享字符串且有值，t="s" 必须留着，否则值会被当数字解析
    expect(sheet).toContain('<c r="A1" t="s"><v>3</v></c>')
    // B2 是纯数字
    expect(sheet).toContain('<c r="B2"><v>42</v></c>')
  })

  it('干净文件里的有值单元格也不被动', async () => {
    const src = await makeZip({
      'xl/styles.xml': STYLES_OK,
      'xl/worksheets/sheet1.xml': SHEET_CLEAN,
    })
    const r = await normalizeWorkbookBytes(src)
    expect(r.applied).toEqual([])
  })

  it('非 zip（.xls 老格式）给的是人话，不是堆栈', async () => {
    const notZip = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
    await expect(normalizeWorkbookBytes(notZip)).rejects.toThrow(/另存为 \.xlsx/)
  })

  it('两种补丁同时命中时都报出来', async () => {
    const src = await makeZip({
      'xl/styles.xml': STYLES_WITH_EMPTY_FILL,
      'xl/worksheets/sheet1.xml': SHEET_MIXED,
    })
    const r = await normalizeWorkbookBytes(src)
    expect(r.applied.length).toBe(2)
    expect(NORMALIZE_PATCH_IDS.every((id) => r.applied.some((a) => a.startsWith(id)))).toBe(true)
  })
})
