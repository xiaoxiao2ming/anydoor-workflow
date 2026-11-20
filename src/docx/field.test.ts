import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parseUrn, pickByLocator, ParamTable, discoverFields, bindFromParamTable } from './field'
import { sdtifyDocx } from './sdtify'

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

/** 造一个最小但合法、含绑定标记的 docx */
async function makeTemplate(bodyInner: string): Promise<Uint8Array> {
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>${bodyInner}<w:sectPr/></w:body></w:document>`
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

const textsOf = (xml: string) =>
  [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => m[1])

async function bodyOf(buf: Uint8Array): Promise<string> {
  const z = await JSZip.loadAsync(buf)
  return z.file('word/document.xml')!.async('string')
}

/* ───────────────────────────────────────────── */

describe('统一引用地址 URN', () => {
  it('四种 scheme 都能解析', () => {
    expect(parseUrn('node://calc1#Q')).toEqual({ scheme: 'node', target: 'calc1', locator: 'Q' })
    expect(parseUrn('file://D:/a.xlsx#Sheet1!B2')).toEqual({
      scheme: 'file',
      target: 'D:/a.xlsx',
      locator: 'Sheet1!B2',
    })
    expect(parseUrn('api://price/steel')).toEqual({ scheme: 'api', target: 'price', locator: 'steel' })
    expect(parseUrn('param://审查单位')).toEqual({ scheme: 'param', target: '审查单位', locator: '' })
  })

  it('非法地址返回 null', () => {
    expect(parseUrn('随便写的')).toBeNull()
    expect(parseUrn('')).toBeNull()
  })
})

describe('locator 取值', () => {
  it('单元格 Sheet!A1', () => {
    expect(pickByLocator({ Sheet1: { B2: 25.5 } }, 'Sheet1!B2')).toBe(25.5)
  })
  it('整列 col:列名', () => {
    const rows = [{ 名称: 'a' }, { 名称: 'b' }]
    expect(pickByLocator(rows, 'col:名称')).toEqual(['a', 'b'])
  })
  it('普通路径', () => {
    expect(pickByLocator({ a: { b: 7 } }, 'a.b')).toBe(7)
    expect(pickByLocator({ a: { b: 7 } }, '$.a.b')).toBe(7)
  })
})

describe('参数表优先级与溯源', () => {
  const ctx = {
    nodes: { calc1: { Q: 1200 } },
    files: { 'D:/价目.xlsx': { 价目表: { B2: 25.5 } } },
    params: { 审查单位: '陵水审图' },
  }
  const specs = [
    { key: '设计流量', source: 'node://calc1#Q' },
    { key: '管材单价', source: 'file://D:/价目.xlsx#价目表!B2' },
    { key: '审查单位', source: 'param://审查单位' },
    { key: '工程名称', default: '默认工程名' },
    { key: '缺失项', source: 'node://nope#x' },
  ]

  it('三种来源都能取到值', () => {
    const t = new ParamTable(specs, ctx)
    expect(t.resolve('设计流量')).toBe(1200)
    expect(t.resolve('管材单价')).toBe(25.5)
    expect(t.resolve('审查单位')).toBe('陵水审图')
    expect(t.resolve('工程名称')).toBe('默认工程名')
  })

  it('手填值覆盖引用', () => {
    const t = new ParamTable(specs, ctx)
    t.overrides['设计流量'] = 999
    expect(t.resolve('设计流量')).toBe(999)
  })

  it('溯源能说清每个值的出处，解析不到会报错', () => {
    const t = new ParamTable(specs, ctx)
    const tr = t.trace()
    expect(tr.find((x) => x.key === '设计流量')?.from).toBe('node')
    expect(tr.find((x) => x.key === '管材单价')?.from).toBe('file')
    const bad = tr.find((x) => x.key === '缺失项')
    expect(bad?.ok).toBe(false)
    expect(bad?.error).toBeTruthy()
  })

  it('源一改，值跟着变（域是引用不是副本）', () => {
    const t = new ParamTable(specs, ctx)
    expect(t.resolve('设计流量')).toBe(1200)
    t.setContext({ ...ctx, nodes: { calc1: { Q: 1500 } } })
    expect(t.resolve('设计流量')).toBe(1500)
  })
})

describe('端到端：模板 → 域发现 → 参数表填充', () => {
  it('顶层 field 与行级 each 都能填充（含回归：行内 field: 前缀）', async () => {
    const tpl = await makeTemplate(
      `<w:p><w:r><w:t>工程名称：</w:t></w:r><w:r><w:t>{{field:工程名称}}</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>设计流量：</w:t></w:r><w:r><w:t>{{field:设计流量}}</w:t></w:r></w:p>` +
        `<w:tbl>` +
        `<w:tr><w:tc><w:p><w:r><w:t>名称</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>数量</w:t></w:r></w:p></w:tc></w:tr>` +
        `<w:tr>{{each:材料表}}<w:tc><w:p><w:r><w:t>{{field:名称}}</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:p><w:r><w:t>{{field:数量}}</w:t></w:r></w:p></w:tc></w:tr>` +
        `</w:tbl>`
    )
    const sdtified = (await sdtifyDocx(tpl)).buffer

    const { fields } = await discoverFields(sdtified)
    expect(fields.map((f) => f.key).sort()).toEqual(['工程名称', '材料表', '设计流量'])
    const each = fields.find((f) => f.key === '材料表')!
    expect(each.hasRow).toBe(true)
    expect(each.placeholders).toContain('field:名称')

    const table = new ParamTable(
      [
        { key: '工程名称', default: '海南陵水燃气工程' },
        { key: '设计流量', source: 'node://calc1#Q' },
        { key: '材料表', source: 'node://table1#rows' },
      ],
      {
        nodes: {
          calc1: { Q: 1200 },
          table1: {
            rows: [
              { 名称: '钢管 DN200', 数量: 100 },
              { 名称: '钢管 DN100', 数量: 40 },
            ],
          },
        },
      }
    )

    const out = await bindFromParamTable(sdtified, table)
    expect(out.bound).toBe(3)
    expect(out.warnings).toHaveLength(0)

    const xml = await bodyOf(out.buffer)
    const texts = textsOf(xml)
    expect(texts).toContain('海南陵水燃气工程')
    expect(texts).toContain('1200')
    // 回归点：each 行内占位符带 field: 前缀时，曾全部填成空字符串
    expect(texts).toContain('钢管 DN200')
    expect(texts).toContain('钢管 DN100')
    expect((xml.match(/<w:tr(?:\s[^>]*)?>/g) || []).length).toBe(3) // 表头 1 + 数据 2
  })

  it('run 被拆成多段时填充仍正确', async () => {
    const tpl = await makeTemplate(
      `<w:p><w:r><w:t>设计流量：</w:t></w:r><w:r><w:t>{{field:设计流量}}</w:t></w:r></w:p>`
    )
    const sdtified = (await sdtifyDocx(tpl)).buffer

    // 模拟 Word 把控件内容拆成 3 个 run
    const z = await JSZip.loadAsync(sdtified)
    let xml = await z.file('word/document.xml')!.async('string')
    xml = xml.replace(
      /(<w:tag w:val="field:设计流量"\/>[\s\S]*?<w:sdtContent>)([\s\S]*?)(<\/w:sdtContent>)/,
      (_m, h, _b, t) => h + '<w:r><w:t>{{</w:t></w:r><w:r><w:t>field:设计</w:t></w:r><w:r><w:t>流量}}</w:t></w:r>' + t
    )
    const z2 = new JSZip()
    z2.file('[Content_Types].xml', await z.file('[Content_Types].xml')!.async('string'))
    z2.file('_rels/.rels', await z.file('_rels/.rels')!.async('string'))
    z2.file('word/document.xml', xml)
    const split = await z2.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })

    const table = new ParamTable([{ key: '设计流量', source: 'node://calc1#Q' }], {
      nodes: { calc1: { Q: 1200 } },
    })
    const out = await bindFromParamTable(split, table)
    const texts = textsOf(await bodyOf(out.buffer))
    expect(texts).toContain('1200')
  })
})
