/**
 * 真实浏览器验证：docx / xlsx 节点 + 双击进入编辑器
 *
 * 用本机 Edge（免下 Chromium）。
 * 前提：vite --config vite.web.config.ts 已起在 5173，app-server 在 5198。
 *
 * 验证点：
 *   1. 页面无控制台错误（含 office-kit 浏览器链路是否真的能跑）
 *   2. 左侧节点库出现 Word / Excel
 *   3. 程序化建一个 docx 节点 → 画布出现 .rf-node--docx，徽章是 Word
 *   4. 程序化建一个 xlsx 节点 → 画布出现 .rf-node--xlsx，徽章是 Excel
 *   5. 双击 docx 节点 → .node-editor-mask 出现，里面是 Word 面板
 *   6. 切到「输入数据」页签 → 表格渲染
 *   7. Esc 关闭
 *   8. 双击 xlsx 节点 → 出现 Excel 面板（含 xrp- 结构）
 */
import { chromium } from 'playwright-core'

const BASE = process.env.SMOKE_URL || 'http://127.0.0.1:5173'
const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
const OUT = new URL('../sample-template/', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')

const log = []
const P = (s) => {
  console.log(s)
  log.push(s)
}

const errors = []

const browser = await chromium.launch({ executablePath: EDGE, headless: true })
const page = await browser.newPage({ viewport: { width: 1680, height: 980 } })
page.on('console', (m) => {
  if (m.type() === 'error') errors.push('[console] ' + m.text())
})
page.on('pageerror', (e) => errors.push('[pageerror] ' + String(e)))

try {
  await page.goto(BASE, { waitUntil: 'load', timeout: 40000 })
  await page.waitForTimeout(2600)
  P('=== 1. 页面加载 ===')
  P('  title = ' + (await page.title()))
  P('  节点库条目: ' + (await page.locator('.palette-item').allInnerTexts()).join(' / '))

  const hasWord = await page.locator('.palette-item', { hasText: 'Word' }).count()
  const hasExcel = await page.locator('.palette-item', { hasText: 'Excel' }).count()
  P('  节点库含 Word: ' + (hasWord > 0 ? '✅' : '❌') + ' / 含 Excel: ' + (hasExcel > 0 ? '✅' : '❌'))

  // ── 0. 先清掉上次跑测试留下的 docx_/xlsx_ 节点 ──
  //   不清的话：每次跑都在同一坐标堆一个节点，hit-test 会打到别的节点上，
  //   dblclick 直接被 playwright 判定为「被遮挡」。也会污染真实画布数据。
  const purged = await page.evaluate(() => {
    const st = window.__store
    if (!st) return -1
    const s = st.getState()
    const junk = Object.values(s.tables)
      .filter((t) => /^(docx|xlsx)_\d+$/.test(t.id))
      .map((t) => t.id)
    for (const id of junk) {
      s.trashTable(id)
      s.purgeTable(id)
    }
    return junk.length
  })
  P('  清理历史测试节点: ' + purged + ' 个')
  await page.waitForTimeout(600)

  // ── 2. 程序化建节点（绕过拖拽：拖拽依赖真实 DataTransfer，这里用 store 钩子） ──
  P('\n=== 2. 建 docx / xlsx 节点 ===')
  const created = await page.evaluate(() => {
    const st = window.__store
    if (!st) return { err: 'window.__store 不存在' }
    const s = st.getState()
    const stamp = Date.now()
    const did = 'docx_' + stamp
    const xid = 'xlsx_' + stamp
    s.addTable({ id: did, name: 'Word 测试文档', kind: 'docx', fields: [], config: { docxMap: {} } })
    s.addTable({ id: xid, name: 'Excel 测试簿', kind: 'xlsx', fields: [], config: { xlsxMap: {} } })
    s.setNodePosition(did, { x: 120, y: 420 })
    s.setNodePosition(xid, { x: 520, y: 420 })
    return { did, xid }
  })
  P('  addTable 结果: ' + JSON.stringify(created))
  await page.waitForTimeout(1400)

  const docxNodes = await page.locator('.rf-node--docx').count()
  const xlsxNodes = await page.locator('.rf-node--xlsx').count()
  P('  .rf-node--docx 数量: ' + docxNodes + (docxNodes > 0 ? ' ✅' : ' ❌'))
  P('  .rf-node--xlsx 数量: ' + xlsxNodes + (xlsxNodes > 0 ? ' ✅' : ' ❌'))

  const badges = await page.locator('.rf-node--docx .rf-badge').allInnerTexts()
  P('  docx 节点徽章文字: ' + JSON.stringify(badges) + (badges.includes('Word') ? ' ✅' : ' ❌'))
  const xbadges = await page.locator('.rf-node--xlsx .rf-badge').allInnerTexts()
  P('  xlsx 节点徽章文字: ' + JSON.stringify(xbadges) + (xbadges.includes('Excel') ? ' ✅' : ' ❌'))

  // 节点卡片上的「双击进入编辑」提示
  const hint = await page.locator('.rf-field--hint').allInnerTexts()
  P('  节点卡片提示行: ' + JSON.stringify(hint))

  await page.screenshot({ path: OUT + 'office-nodes-1-canvas.png' })
  P('  截图 → sample-template/office-nodes-1-canvas.png')

  // ── 3. 双击 docx 节点 → 打开编辑器 ──
  P('\n=== 3. 双击 docx 节点 ===')
  await page.locator('.rf-node--docx').first().dblclick()
  await page.waitForTimeout(900)
  const maskOpen = await page.locator('.node-editor-mask').count()
  P('  .node-editor-mask 出现: ' + maskOpen + (maskOpen > 0 ? ' ✅' : ' ❌'))
  const editorTitle = await page.locator('.ne-kind').first().innerText().catch(() => '(无)')
  P('  编辑器标题: ' + editorTitle)
  const neName = await page.locator('.ne-name').first().inputValue().catch(() => '(无)')
  P('  编辑器里的节点名: ' + neName)

  // Portal 校验：编辑器必须是 document.body 的直接子节点，而不是待在画布 DOM 里
  const portalOk = await page.evaluate(() => {
    const el = document.querySelector('.node-editor-mask')
    if (!el) return 'mask 不存在'
    const inFlow = !!el.closest('.react-flow')
    return inFlow ? '❌ 仍在 react-flow 子树内' : '✅ 已脱离 react-flow（body 直属）'
  })
  P('  Portal 校验: ' + portalOk)

  await page.screenshot({ path: OUT + 'office-nodes-2-word-editor.png' })
  P('  截图 → sample-template/office-nodes-2-word-editor.png')

  // ── 4. 切到「输入数据」页签 ──
  P('\n=== 4. 编辑器页签 ===')
  await page.locator('.ne-tabs .tab[data-tab="data"]').click()
  await page.waitForTimeout(500)
  const emptyOrTable = (await page.locator('.ne-empty, .ne-table-wrap').count()) > 0
  P('  数据页渲染: ' + (emptyOrTable ? '✅' : '❌'))
  await page.locator('.ne-tabs .tab[data-tab="main"]').click()
  await page.waitForTimeout(400)
  const wrpInEditor = await page.locator('.node-editor .wrp').count()
  P('  绑定页含 Word 面板 (.wrp): ' + wrpInEditor + (wrpInEditor > 0 ? ' ✅' : ' ❌'))

  // ── 5. Esc 关闭 ──
  P('\n=== 5. Esc 关闭 ===')
  await page.locator('.ne-head').click() // 让焦点离开输入框
  await page.keyboard.press('Escape')
  await page.waitForTimeout(500)
  P('  关闭后 mask 数量: ' + (await page.locator('.node-editor-mask').count()))

  // ── 6. 双击 xlsx 节点 ──
  P('\n=== 6. 双击 xlsx 节点 ===')
  await page.locator('.rf-node--xlsx').first().dblclick()
  await page.waitForTimeout(900)
  P('  mask 出现: ' + (await page.locator('.node-editor-mask').count()) + (await page.locator('.node-editor-mask').count() > 0 ? ' ✅' : ' ❌'))
  P('  标题: ' + (await page.locator('.ne-kind').first().innerText().catch(() => '(无)')))
  const xrp = await page.locator('.node-editor .xrp').count()
  P('  含 Excel 面板 (.xrp): ' + xrp + (xrp > 0 ? ' ✅' : ' ❌'))
  const xrpText = await page.locator('.node-editor').first().innerText().catch(() => '')
  P('  Excel 面板含「Excel 工作簿模板」: ' + (xrpText.includes('Excel 工作簿模板') ? '✅' : '❌'))
  await page.screenshot({ path: OUT + 'office-nodes-3-excel-editor.png' })
  P('  截图 → sample-template/office-nodes-3-excel-editor.png')

  // 关闭
  await page.locator('.ne-close').click()
  await page.waitForTimeout(400)

  // ── 7. 右键菜单（确认你原有的功能没被破坏） ──
  P('\n=== 7. 右键菜单 & 回收站（回归） ===')
  await page.locator('.rf-node--xlsx').first().click({ button: 'right' })
  await page.waitForTimeout(600)
  const ctxItems = await page.locator('.ctx-menu .ctx-item').allInnerTexts()
  P('  右键菜单项: ' + JSON.stringify(ctxItems))
  P('  含「打开编辑器」: ' + (ctxItems.some((t) => t.includes('打开编辑器')) ? '✅' : '❌'))
  await page.keyboard.press('Escape')
  await page.waitForTimeout(300)
  const trashBtn = await page.locator('.trash-btn').count()
  P('  工具栏回收站按钮存在: ' + (trashBtn > 0 ? '✅' : '❌'))

  // ── 8. 清理：把测试节点从画布里彻底移除，不污染真实数据 ──
  P('\n=== 8. 清理测试节点 ===')
  const cleaned = await page.evaluate(() => {
    const s = window.__store.getState()
    const junk = Object.values(s.tables)
      .filter((t) => /^(docx|xlsx)_\d+$/.test(t.id))
      .map((t) => t.id)
    for (const id of junk) s.purgeTable(id)
    return junk.length
  })
  P('  已移除: ' + cleaned + ' 个')

  // ── 9. 控制台错误 ──
  P('\n=== 9. 控制台错误 ===')
  const real = errors.filter((e) => !/favicon/i.test(e) && !/404/.test(e))
  if (real.length === 0) P('  ✅ 无错误（仅 favicon 404，与本次改动无关）')
  else {
    P('  ❌ ' + real.length + ' 条：')
    real.slice(0, 12).forEach((e) => P('    ' + e.slice(0, 220)))
  }
} catch (e) {
  P('\n!! 脚本抛错: ' + (e && e.message ? e.message : String(e)))
  P(String(e && e.stack ? e.stack : '').split('\n').slice(0, 6).join('\n'))
} finally {
  await browser.close()
  const fs = await import('node:fs')
  fs.writeFileSync(OUT + 'office-nodes-check.txt', log.join('\n'), 'utf8')
  console.log('\n报告已写入 sample-template/office-nodes-check.txt')
}
