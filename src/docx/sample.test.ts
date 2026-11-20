import { describe, it, expect } from 'vitest'
// 注意：这里用 fs 的同步 API 而不是 fs/promises —— vitest 环境下 fs/promises
// 会被宿主注入的 shim 拦成 callback 版本，await readFile() 会抛 ERR_INVALID_ARG_TYPE。
import { existsSync, readFileSync } from 'fs'
import { join, dirname } from 'path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { sdtifyDocx } from './sdtify'
import { parseTemplate, bind } from './binder'
import { buildData } from './buildData'
import type { FieldDef } from '../core/schema'
import type { Row } from '../core/store'

/**
 * 拿 scripts/make-sample-template.mjs 生成的真实 .docx 走完整链路。
 * 相比手工拼 XML，这份模板带 styles.xml、页眉、真实表格边框，更接近设计院给的文件。
 */

const HERE = dirname(fileURLToPath(import.meta.url))
const TPL = join(HERE, '..', '..', 'sample-template', '市政管道工程量报告-模板.docx')

const FIELDS: FieldDef[] = [
  { id: 'f1', name: '工程名称', type: 'text' },
  { id: 'f2', name: '编制日期', type: 'text' },
  { id: 'f3', name: '编制人', type: 'text' },
  { id: 'f4', name: '管线总长', type: 'number' },
  { id: 'f5', name: '检查井数量', type: 'number' },
  { id: 'f6', name: '总投资', type: 'number' },
]
const ROWS: Row[] = [{ f1: '示例市政管道改造工程', f2: '2026-09-03', f3: '张三', f4: 1286.5, f5: 42, f6: 386.4 }]
const LIST: Row[] = [
  { 序号: 1, 管段编号: 'W1-W2', 管径: 'DN300', 长度: 128.4, 备注: '重力流' },
  { 序号: 2, 管段编号: 'W2-W3', 管径: 'DN400', 长度: 216.7, 备注: '穿越路口' },
  { 序号: 3, 管段编号: 'W3-W4', 管径: 'DN500', 长度: 342.0, 备注: '' },
  { 序号: 4, 管段编号: 'W4-W5', 管径: 'DN600', 长度: 599.4, 备注: '需支护' },
]

const hasSample = existsSync(TPL)

describe.skipIf(!hasSample)('真实模板端到端', () => {
  it('sdtifyDocx 控件化，正文与页眉都覆盖到', async () => {
    const raw = new Uint8Array(readFileSync(TPL))
    const r = await sdtifyDocx(raw)

    expect(r.count).toBe(11) // 正文 10 + 页眉 1
    expect(r.changed).toContain('word/document.xml')
    expect(r.changed).toContain('word/header1.xml')

    // 标记被剔除，不该残留在最终文档里
    const doc = await (await JSZip.loadAsync(r.buffer)).file('word/document.xml')!.async('string')
    expect(doc).not.toContain('{{each:管道清单}}')
    expect(doc).not.toContain('{{if:需风险提示}}')
  })

  it('绑定后产物内容正确、控件剥净、样式完好', async () => {
    const raw = new Uint8Array(readFileSync(TPL))
    const templated = await sdtifyDocx(raw)
    const { bindings } = await parseTemplate(templated.buffer)

    const built = buildData(
      bindings,
      {
        'each:管道清单': { from: 'rows' },
        'rich:结论与建议': {
          from: 'manual',
          value:
            '<p>管线布置<b>满足规范要求</b>，全线采用重力流：</p><ol><li>埋深 3.4 m 段需支护；</li><li>穿越路口处需 <span style="color:#C00000">加强防护</span>。</li></ol><h3>建议</h3><p>施工前完成管线交底。</p>',
        },
        'if:需风险提示': { from: 'manual', value: 'true' },
      },
      FIELDS,
      ROWS,
    )

    const { buffer, bound, warnings } = await bind(
      templated.buffer,
      { ...built.data, 管道清单: LIST },
      { stripControls: true },
    )

    expect(warnings).toEqual([])
    expect(bound).toBe(11)

    const zip = await JSZip.loadAsync(buffer)
    const doc = await zip.file('word/document.xml')!.async('string')
    const hdr = await zip.file('word/header1.xml')!.async('string')

    // 正文
    expect(doc).toContain('示例市政管道改造工程')
    expect(doc).toContain('1286.5')
    expect(doc).toContain('42')
    expect(doc).toContain('386.4')
    expect(doc).toContain('米，共设检查井')
    expect(doc).toContain('座，总投资约')
    // each 展开 4 行
    expect((doc.match(/W\d-W\d/g) || []).length).toBe(4)
    expect(doc).toContain('DN600')
    expect(doc).toContain('599.4')
    // rich：真列表 + 颜色 + 小标题
    expect(doc).toContain('w:numId w:val="101"')
    expect(doc).toContain('w:val="C00000"')
    expect(doc).toContain('加强防护')
    // if 为真保留
    expect(doc).toContain('交叉施工')
    // 页眉也绑上了
    expect(hdr).toContain('示例市政管道改造工程')

    // 交付稿里不剩任何控件和占位符
    expect(doc).not.toContain('<w:sdt>')
    expect(doc).not.toContain('{{')

    // 格式要素保留
    expect(doc).toContain('w:pStyle w:val="Heading1"')
    expect(doc).toContain('w:sz w:val="36"')
    expect(doc).toContain('w:tblBorders')
    expect(doc).toContain('w:shd w:val="clear"')
    // numbering 自动注入了（富文本列表需要）
    expect(zip.file('word/numbering.xml')).toBeTruthy()
  })

  it('if 为假时风险提示整段消失', async () => {
    const raw = new Uint8Array(readFileSync(TPL))
    const templated = await sdtifyDocx(raw)
    const { bindings } = await parseTemplate(templated.buffer)
    const built = buildData(
      bindings,
      {
        'each:管道清单': { from: 'rows' },
        'rich:结论与建议': { from: 'manual', value: '<p>无。</p>' },
        'if:需风险提示': { from: 'manual', value: 'false' },
      },
      FIELDS,
      ROWS,
    )
    const { buffer } = await bind(templated.buffer, { ...built.data, 管道清单: LIST }, { stripControls: true })
    const doc = await (await JSZip.loadAsync(buffer)).file('word/document.xml')!.async('string')
    expect(doc).not.toContain('交叉施工')
    expect(doc).toContain('示例市政管道改造工程')
  })
})
