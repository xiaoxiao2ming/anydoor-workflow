/**
 * 浏览器验证：离线「在软件里打开 Excel」
 *
 * 三个层次：
 *   ① 合成模板（有命名区域 + Table）→ 打开工作簿：网格、工作表页签、域标记
 *   ② 改一个单元格 → 字节真的变了（dirty 计数）
 *   ③ **真实 WPS 文件（190KB）** → 归一化层生效，能打开（改造前这一步是直接抛异常的）
 */
import fs from 'node:fs'
import { chromium } from 'playwright-core'

const BASE = process.env.SMOKE_URL || 'http://127.0.0.1:5173'
const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
const DIR = new URL('../sample-template/', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')
const REAL = '<your-xlsx>/燃气主要材料设备对照表-暂定.xlsx'

const log = []
const P = (s) => {
  console.log(s)
  log.push(s)
}
const errors = []

const browser = await chromium.launch({ executablePath: EDGE, headless: true })
const page = await browser.newPage({ viewport: { width: 1760, height: 1000 } })
page.on('console', (m) => {
  if (m.type() === 'error') errors.push('[console] ' + m.text())
})
page.on('pageerror', (e) => errors.push('[pageerror] ' + String(e)))

try {
  await page.goto(BASE, { waitUntil: 'load', timeout: 40000 })
  await page.waitForTimeout(2600)

  // 清理历史
  await page.evaluate(() => {
    const s = window.__store.getState()
    for (const t of Object.values(s.tables)) {
      if (/^(docx|xlsx|param)_\d+$/.test(t.id)) {
        s.trashTable(t.id)
        s.purgeTable(t.id)
      }
    }
  })
  await page.waitForTimeout(500)

  const ids = await page.evaluate(() => {
    const st = window.__store.getState()
    const t = Date.now()
    const pid = 'param_' + t
    const xid = 'xlsx_' + t
    const F = (name, type) => ({ id: name, name, type })
    st.addTable({
      id: pid,
      name: '上游数据',
      kind: 'param',
      fields: [F('工程名称', 'text'), F('设计流量', 'number'), F('名称', 'text'), F('数量', 'number'), F('单价', 'number')],
      config: {},
    })
    st.setTableData(
      pid,
      [F('工程名称', 'text'), F('设计流量', 'number'), F('名称', 'text'), F('数量', 'number'), F('单价', 'number')],
      [{ 工程名称: '海南陵水英州镇燃气管道工程', 设计流量: 1200, 名称: '钢管 DN200', 数量: 100, 单价: 25.5 }],
    )
    st.setNodePosition(pid, { x: 60, y: 400 })
    st.addTable({ id: xid, name: '测试工作簿', kind: 'xlsx', fields: [], config: { xlsxMap: {} } })
    st.setNodePosition(xid, { x: 600, y: 400 })
    st.setNodeConfig(xid, { inputs: [pid], linkedInputs: [pid] })
    return { pid, xid }
  })
  await page.waitForTimeout(1200)

  // ── ① 打开 xlsx 节点编辑器 ──
  P('=== 1. 双击 xlsx 节点 ===')
  await page.locator(`.react-flow__node[data-id="${ids.xid}"] .rf-node--xlsx`).dblclick()
  await page.waitForTimeout(900)
  const tabs = await page.locator('.ne-tabs .tab').evaluateAll((els) => els.map((e) => e.getAttribute('title')))
  P('  页签: ' + JSON.stringify(tabs))
  P('  含「打开工作簿」: ' + (tabs.includes('打开工作簿') ? '✅' : '❌'))

  // ── ② 先用合成模板 ──
  P('\n=== 2. 导入合成模板（有命名区域 + Table）===')
  await page.locator('.node-editor .xrp input[type="file"][accept=".xlsx"]').setInputFiles(DIR + '测试模板.xlsx')
  await page.waitForTimeout(3500)
  const slots = await page.locator('.node-editor .xrp-item').count()
  P('  发现域: ' + slots + (slots === 3 ? ' ✅' : ' ⚠'))

  P('\n=== 3. 打开工作簿页签 ===')
  await page.locator('.ne-tabs .tab[data-tab="book"]').click()
  await page.waitForTimeout(2600)

  const sheetTabs = await page.locator('.wb-tab').allInnerTexts()
  P('  工作表页签: ' + JSON.stringify(sheetTabs))
  P('  含 汇总表/工程量表: ' + (sheetTabs.includes('汇总表') && sheetTabs.includes('工程量表') ? '✅' : '❌'))

  const stat = await page.locator('.wb-stat').first().innerText().catch(() => '(无)')
  P('  统计栏: ' + stat.replace(/\n/g, ' '))

  const chips = await page.locator('.wb-chip').allInnerTexts()
  P('  域徽章: ' + JSON.stringify(chips))

  const gridRows = await page.locator('.wb-grid .ag-row').count()
  P('  网格可见行数: ' + gridRows + (gridRows > 0 ? ' ✅' : ' ❌'))

  // 域高亮的单元格
  const namedCells = await page.locator('.wb-cell-named').count()
  const tableCells = await page.locator('.wb-cell-table').count()
  P('  命名区域高亮格: ' + namedCells + ' / 表格范围格: ' + tableCells)

  // 网格里应该能看到模板内容（首个工作表是「汇总表」）
  const sheetText = await page.locator('.wb-grid').innerText()
  P('  网格含「工程名称/设计流量」: ' + (sheetText.includes('工程名称') && sheetText.includes('设计流量') ? '✅' : '❌'))

  await page.screenshot({ path: DIR + 'wb-1-grid.png' })
  P('  截图 → sample-template/wb-1-grid.png')

  // ── ③b 切到「工程量表」验表格高亮 ──
  P('\n=== 3b. 切到工程量表（Table 所在表）===')
  await page.locator('.wb-tab', { hasText: '工程量表' }).click()
  await page.waitForTimeout(1200)
  const sheet2Text = await page.locator('.wb-grid').innerText()
  P('  含「名称/数量/单价/合价」: ' +
    (['名称', '数量', '单价', '合价'].every((k) => sheet2Text.includes(k)) ? '✅' : '❌'))
  const tblCells = await page.locator('.wb-cell-table').count()
  P('  表格范围高亮格: ' + tblCells + (tblCells > 0 ? ' ✅' : ' ❌'))
  const chips2 = await page.locator('.wb-chip').allInnerTexts()
  P('  该表域徽章: ' + JSON.stringify(chips2))
  await page.screenshot({ path: DIR + 'wb-1b-table.png' })
  // 切回汇总表
  await page.locator('.wb-tab', { hasText: '汇总表' }).click()
  await page.waitForTimeout(800)

  // ── ④ 改一个单元格 ──
  P('\n=== 4. 改一个单元格 ===')
  // 定位 B2（数量列第 2 行）
  const cell = page.locator('.wb-grid .ag-row[row-index="1"] .ag-cell[col-id="B"]').first()
  await cell.dblclick()
  await page.waitForTimeout(400)
  await page.keyboard.type('777')
  await page.keyboard.press('Enter')
  await page.waitForTimeout(1800)
  const dirty = await page.locator('.wb-dirty').count()
  P('  出现「已改 N 处」: ' + (dirty > 0 ? '✅ ' + (await page.locator('.wb-dirty').innerText()) : '❌'))

  // ── ⑤ 域位置弹层 ──
  P('\n=== 5. 域位置弹层 ===')
  await page.locator('.wb-foot button', { hasText: '域位置' }).click()
  await page.waitForTimeout(700)
  const dlgRows = await page.locator('.wb-domain-row').count()
  const dlgKeys = await page.locator('.wb-domain-key').allInnerTexts()
  const dlgRefs = await page.locator('.wb-domain-ref').allInnerTexts()
  P('  列出域: ' + dlgRows + ' → ' + JSON.stringify(dlgKeys))
  P('  引用显示: ' + JSON.stringify(dlgRefs))
  const dup = dlgRefs.some((r) => /^([^!]+)!\1!/.test(r))
  P('  表名重复（回归项）: ' + (dup ? '❌ 仍重复' : '✅ 已修'))
  await page.screenshot({ path: DIR + 'wb-2-domains.png' })
  await page.locator('.wb-domains-head .ne-close').click()
  await page.waitForTimeout(400)

  // ── ⑥ 真实 WPS 文件 ──
  P('\n=== 6. 换成真实 WPS 文件（190KB，改造前打不开）===')
  P('  文件存在: ' + fs.existsSync(REAL) + '  ' + (fs.existsSync(REAL) ? (fs.statSync(REAL).size / 1024).toFixed(0) + 'KB' : ''))
  // 回到绑定页换模板
  await page.locator('.ne-tabs .tab[data-tab="main"]').click()
  await page.waitForTimeout(500)
  await page.locator('.node-editor .xrp input[type="file"][accept=".xlsx"]').setInputFiles(REAL)
  await page.waitForTimeout(5000)

  const realErr = await page.locator('.node-editor .xrp-err').count()
  if (realErr) P('  ⚠ 绑定面板报错: ' + (await page.locator('.node-editor .xrp-err').first().innerText()).slice(0, 120))
  else P('  绑定面板无报错 ✅')

  await page.locator('.ne-tabs .tab[data-tab="book"]').click()
  await page.waitForTimeout(5000)

  const realSheetTabs = await page.locator('.wb-tab').allInnerTexts()
  P('  工作表: ' + JSON.stringify(realSheetTabs))
  const realErrBox = await page.locator('.wb-err').count()
  if (realErrBox) {
    P('  ❌ 打不开: ' + (await page.locator('.wb-err-msg').first().innerText()).slice(0, 200))
  } else {
    P('  ✅ 打开了')
  }
  const realStat = await page.locator('.wb-stat').first().innerText().catch(() => '(无)')
  P('  统计: ' + realStat.replace(/\n/g, ' '))
  const normNote = await page.locator('.wb-note > summary').count()
  if (normNote) P('  归一化告知: ' + (await page.locator('.wb-note > summary').innerText()))
  else P('  无归一化告知（说明这个文件没命中补丁，或没打开成功）')

  const realRows = await page.locator('.wb-grid .ag-row').count()
  P('  网格可见行: ' + realRows)
  // 看真实数据有没有出来（A1 是「序号」）
  const realText = await page.locator('.wb-grid').innerText().catch(() => '')
  P('  网格含「序号」: ' + (realText.includes('序号') ? '✅' : '❌'))
  await page.screenshot({ path: DIR + 'wb-3-real-wps.png' })
  P('  截图 → sample-template/wb-3-real-wps.png')

  // ── ⑦ 控制台错误 ──
  P('\n=== 7. 控制台错误 ===')
  const real = errors.filter((e) => !/favicon|404/.test(e))
  P(real.length === 0 ? '  ✅ 无（含 office-kit 解析真实文件全程无报错）' : '  ❌ ' + real.length + ' 条：')
  real.slice(0, 10).forEach((e) => P('    ' + e.slice(0, 260)))

  // 清理
  await page.evaluate((o) => {
    const s = window.__store.getState()
    for (const id of [o.pid, o.xid]) s.purgeTable(id)
  }, ids)
  P('\n=== 8. 清理完成 ===')
} catch (e) {
  P('\n!! 脚本抛错: ' + (e?.message ?? String(e)))
  P(String(e?.stack ?? '').split('\n').slice(0, 6).join('\n'))
} finally {
  await browser.close()
  fs.writeFileSync(DIR + 'wb-check.txt', log.join('\n'), 'utf8')
  console.log('\n报告 → sample-template/wb-check.txt')
}
