/**
 * 拿示例模板 + 一份假业务数据，跑一次真实绑定，产出可直接用 Word 打开的 .docx。
 *
 * 这一步产出的文件与 UI 里「导出 Word」拿到的是同一份字节 —— 用来在没有 UI 的情况下
 * 验证全链路，也方便直接拿给设计院看效果。
 *
 * 运行：node_modules/.bin/vite-node scripts/bind-sample.mjs
 */

// 用 fs 的同步 API：fs/promises 在 vite-node / vitest 环境下会被宿主 shim 成
// callback 版本，await 调用会抛 ERR_INVALID_ARG_TYPE。
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'fs'
import { dirname, join } from 'path'
import { fileURLToPath } from 'node:url'
import { sdtifyDocx } from '../src/docx/sdtify.ts'
import { parseTemplate, bind } from '../src/docx/binder.ts'
import { buildData } from '../src/docx/buildData.ts'

const HERE = dirname(fileURLToPath(import.meta.url))
const OUT = join(HERE, '..', 'sample-template')
const TPL = join(OUT, '市政管道工程量报告-模板.docx')

const FIELDS = [
  { id: 'f1', name: '工程名称', type: 'text' },
  { id: 'f2', name: '编制日期', type: 'text' },
  { id: 'f3', name: '编制人', type: 'text' },
  { id: 'f4', name: '管线总长', type: 'number' },
  { id: 'f5', name: '检查井数量', type: 'number' },
  { id: 'f6', name: '总投资', type: 'number' },
]
const ROWS = [{ f1: '示例市政管道改造工程', f2: '2026-09-03', f3: '张三', f4: 1286.5, f5: 42, f6: 386.4 }]

const LIST = [
  { 序号: 1, 管段编号: 'W1-W2', 管径: 'DN300', 长度: 128.4, 备注: '重力流' },
  { 序号: 2, 管段编号: 'W2-W3', 管径: 'DN400', 长度: 216.7, 备注: '穿越路口' },
  { 序号: 3, 管段编号: 'W3-W4', 管径: 'DN500', 长度: 342.0, 备注: '' },
  { 序号: 4, 管段编号: 'W4-W5', 管径: 'DN600', 长度: 599.4, 备注: '需支护' },
]

const RICH =
  '<p>本工程管线布置<b>满足《室外排水设计标准》GB 50014</b>要求，全线采用重力流，主要结论如下：</p>' +
  '<ol><li>管径 DN300～DN600，坡度 0.003～0.005，流速满足自清要求；</li>' +
  '<li>W4-W5 段埋深 <span style="color:#C00000">4.2 m</span>，超过 3.0 m 须做支护；</li>' +
  '<li>穿越示例路口段需与交警部门协调，建议夜间施工。</li></ol>' +
  '<h3>建议</h3>' +
  '<p>施工前完成既有管线探测与交底，重点核查 W2-W3 段与现状给水管的高程冲突。</p>'

const run = async () => {
  if (!existsSync(TPL)) throw new Error(`模板不存在，先跑 node scripts/make-sample-template.mjs：${TPL}`)

  const raw = new Uint8Array(readFileSync(TPL))
  const templated = await sdtifyDocx(raw)
  const { bindings, warnings } = await parseTemplate(templated.buffer)
  if (warnings.length) console.log('解析警告：', warnings)

  const built = buildData(
    bindings,
    {
      'each:管道清单': { from: 'rows' },
      'rich:结论与建议': { from: 'manual', value: RICH },
      'if:需风险提示': { from: 'manual', value: 'true' },
    },
    FIELDS,
    ROWS,
  )

  const data = { ...built.data, 管道清单: LIST }
  const { buffer, bound, warnings: bw } = await bind(templated.buffer, data, { stripControls: true })

  mkdirSync(OUT, { recursive: true })
  const out = join(OUT, '市政管道工程量报告-已填充.docx')
  writeFileSync(out, Buffer.from(buffer))

  console.log('绑定控件数：', bound, '/', bindings.length)
  if (bw.length) console.log('绑定警告：', bw)
  console.log('已生成：', out)
  console.log('大小：', (buffer.byteLength / 1024).toFixed(1), 'KB')
  console.log('取值明细：')
  for (const [tag, desc] of Object.entries(built.resolved)) console.log('  -', tag, '→', desc)
}

run().catch((e) => {
  console.error(e)
  process.exit(1)
})
