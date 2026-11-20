/**
 * Word 侧「取值溯源」浏览器验证
 *
 * 验的是这三件事（都是本轮新加的）：
 *   ① Word 面板的上游走的是共享的 buildUpstream（不再只读 config.inputs）
 *   ② 每个域都能给出 URN 溯源（node://…）
 *   ③ 溯源面板真的渲染出来，且计数正确
 *
 * 用真实浏览器 + 真实模板（sample-template/字段模板.docx，含 3 个 sdt 控件）。
 * 跑之前需要 dev server 在 5173（vite），并预置 sample-template/字段模板.docx。
 */
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { chromium } from 'playwright-core'

const ROOT = path.resolve(fileURLToPath(new URL('..', import.meta.url)))
const DIR = path.join(ROOT, 'sample-template') + path.sep
const URL_ = process.env.APP_URL || 'http://127.0.0.1:5173'

const EDGE = [
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
].find((p) => fs.existsSync(p))

const P = (s) => console.log(s)

if (!EDGE) {
  console.error('找不到 Edge / Chrome')
  process.exit(2)
}
if (!fs.existsSync(DIR + '字段模板.docx')) {
  console.error('缺模板：sample-template/字段模板.docx')
  process.exit(2)
}

const browser = await chromium.launch({ executablePath: EDGE, headless: true })
const page = await browser.newPage({ viewport: { width: 1500, height: 950 } })
const errors = []
page.on('console', (m) => {
  if (m.type() === 'error') errors.push(m.text())
})

await page.goto(URL_, { waitUntil: 'networkidle' })
await page.waitForTimeout(1500)
P('=== 1. 页面加载 ===')
P('  title: ' + (await page.title()))

/* ── 2. 建节点：param 上游 + docx 节点，并连线 ── */
P('')
P('=== 2. 建上游与 docx 节点 ===')
const ids = await page.evaluate(() => {
  const st = window.__store
  if (!st) return { err: 'window.__store 不存在' }
  const s = st.getState()
  // 清掉上次跑留下的测试节点，避免叠加
  for (const t of Object.values(s.tables)) {
    if (/^_trace_/.test(t.id)) s.removeTable(t.id)
  }
  const stamp = Date.now()
  const pid = '_trace_up_' + stamp
  const did = '_trace_docx_' + stamp
  const F = (name, type) => ({ id: name, name, type })

  s.addTable({ id: pid, name: '溯源上游数据', kind: 'param', fields: [], config: {} })
  s.setTableData(
    pid,
    [F('工程名称', 'text'), F('设计流量', 'number'), F('名称', 'text'), F('数量', 'number'), F('单价', 'number')],
    [
      { 工程名称: '海南陵水英州镇燃气管道工程', 设计流量: 1200, 名称: '钢管 DN200', 数量: 100, 单价: 25.5 },
      { 名称: '钢管 DN100', 数量: 40, 单价: 18.2 },
      { 名称: '阀门 DN50', 数量: 12, 单价: 260 },
    ],
  )
  s.setNodePosition(pid, { x: 80, y: 430 })

  s.addTable({ id: did, name: 'Word 溯源测试', kind: 'docx', fields: [], config: { docxMap: {} } })
  s.setNodePosition(did, { x: 640, y: 430 })
  // 同时给 inputs 与 linkedInputs —— 模拟「拖线连上」
  s.setNodeConfig(did, { inputs: [pid], linkedInputs: [pid] })
  return { pid, did }
})
P('  上游 ' + ids.pid + ' / docx ' + ids.did)
await page.waitForTimeout(1200)

/* ── 3. 双击打开放大态编辑器 ── */
P('')
P('=== 3. 双击 docx 节点 ===')
await page.locator(`.react-flow__node[data-id="${ids.did}"] .rf-node--docx`).dblclick()
await page.waitForTimeout(900)
P('  编辑器打开: ' + ((await page.locator('.node-editor-mask').count()) > 0 ? '✅' : '❌'))

const meta = await page.locator('.node-editor .wrp-meta').innerText().catch(() => '(无)')
P('  上游汇入行: ' + meta.replace(/\s+/g, ' '))
P('  含「3 行」: ' + (meta.includes('3 行') ? '✅（连线确实被算进上游）' : '❌'))

/* ── 4. 导入模板（走真实 file chooser）── */
P('')
P('=== 4. 导入 Word 模板 ===')
const [chooser] = await Promise.all([
  page.waitForEvent('filechooser'),
  page.locator('.node-editor').getByRole('button', { name: /导入 Word 模板/ }).first().click(),
])
await chooser.setFiles(DIR + '字段模板.docx')
await page.waitForTimeout(2500)

const items = await page.locator('.node-editor .wrp-item').count()
P('  绑定项数: ' + items + (items === 3 ? ' ✅' : ' ❌（期望 3：工程名称/设计流量/材料表）'))

/* ── 5. 溯源面板 ── */
P('')
P('=== 5. 取值溯源 ===')
const hasTrace = (await page.locator('.node-editor .wrp-trace').count()) > 0
P('  溯源面板出现: ' + (hasTrace ? '✅' : '❌'))
if (hasTrace) {
  await page.locator('.node-editor .wrp-trace > summary').click()
  await page.waitForTimeout(300)
  const rows = await page.locator('.node-editor .wrp-trace-row').count()
  P('  溯源条目数: ' + rows + (rows === items ? ' ✅' : ' ❌'))

  const all = await page.locator('.node-editor .wrp-trace-row').allInnerTexts()
  for (const t of all) P('    ' + t.replace(/\s+/g, ' | '))
  const joined = all.join(' ')
  P('  含 node:// URN: ' + (joined.includes('node://') ? '✅' : '❌'))
  P('  来源标为「画布节点」: ' + (joined.includes('画布节点') ? '✅' : '❌'))
}

await page.screenshot({ path: DIR + 'docx-trace.png' })
P('  截图 → sample-template/docx-trace.png')

/* ── 6. 清理 ── */
await page.evaluate(({ pid, did }) => {
  const st = window.__store
  if (!st) return
  const s = st.getState()
  s.removeTable(did)
  s.removeTable(pid)
}, ids)
await page.waitForTimeout(400)
const left = await page.evaluate(() =>
  Object.keys(window.__store.getState().tables).filter((k) => /^_trace_/.test(k)).length,
)
P('')
P('=== 6. 清理 ===')
P('  残留测试节点: ' + left + (left === 0 ? ' ✅' : ' ❌'))

P('')
P('=== 7. 控制台错误 ===')
if (errors.length === 0) P('  ✅ 无错误')
else {
  P('  ❌ ' + errors.length + ' 条：')
  errors.slice(0, 10).forEach((e) => P('    ' + e.slice(0, 200)))
}

await browser.close()
process.exitCode = errors.length ? 1 : 0
