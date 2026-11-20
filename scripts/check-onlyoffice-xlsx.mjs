/**
 * 验证 Excel 侧「在线编辑」通道（没有 DS 的情况下能验到哪一步）
 *
 * 能验：
 *   ✅ 按钮存在且可点
 *   ✅ POST /office/upload 成功（app-server 通道通）
 *   ✅ 弹层出现，因为没填 DS 地址 → 显示人话提示而不是堆栈
 *   ✅ 填上假 DS 地址后，报的应该是「Document Server 不可达」而不是白屏
 * 不能验：真实编辑（需要 NAS 上跑着 DS）
 */
import fs from 'node:fs'
import { chromium } from 'playwright-core'

const BASE = 'http://127.0.0.1:5173'
const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
const DIR = new URL('../sample-template/', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')

const log = []
const P = (s) => {
  console.log(s)
  log.push(s)
}

const browser = await chromium.launch({ executablePath: EDGE, headless: true })
const page = await browser.newPage({ viewport: { width: 1760, height: 1000 } })
const uploads = []
page.on('response', (r) => {
  if (r.url().includes('/office/')) uploads.push(r.status() + ' ' + r.request().method() + ' ' + new URL(r.url()).pathname)
})
const errors = []
page.on('pageerror', (e) => errors.push(String(e)))

try {
  await page.goto(BASE, { waitUntil: 'load', timeout: 40000 })
  await page.waitForTimeout(2500)
  // 清掉上次遗留
  await page.evaluate(() => {
    const s = window.__store.getState()
    for (const t of Object.values(s.tables)) if (/^(xlsx|docx|param)_\d+$/.test(t.id)) { s.trashTable(t.id); s.purgeTable(t.id) }
  })
  // 清掉可能残留的 DS 设置，保证测的是「未配置」路径
  await page.evaluate(() => {
    localStorage.removeItem('client-base.office.ds')
    localStorage.removeItem('client-base.office.host')
  })
  await page.waitForTimeout(400)

  const ids = await page.evaluate(() => {
    const st = window.__store.getState()
    const t = Date.now()
    const pid = 'param_' + t
    const xid = 'xlsx_' + t
    const F = (n, ty) => ({ id: n, name: n, type: ty })
    st.addTable({ id: pid, name: '上游', kind: 'param', fields: [F('工程名称', 'text'), F('名称', 'text')], config: {} })
    st.setTableData(pid, [F('工程名称', 'text'), F('名称', 'text')], [{ 工程名称: '测试工程', 名称: '钢管' }])
    st.setNodePosition(pid, { x: 60, y: 420 })
    st.addTable({ id: xid, name: '在线编辑测试', kind: 'xlsx', fields: [], config: { xlsxMap: {} } })
    st.setNodePosition(xid, { x: 620, y: 420 })
    st.setNodeConfig(xid, { inputs: [pid], linkedInputs: [pid] })
    return { pid, xid }
  })
  await page.waitForTimeout(1200)

  P('=== 1. 打开 Excel 节点编辑器并选模板 ===')
  await page.locator(`.react-flow__node[data-id="${ids.xid}"] .rf-node--xlsx`).dblclick()
  await page.waitForTimeout(900)
  await page.locator('.node-editor .xrp input[type="file"][accept=".xlsx"]').setInputFiles(DIR + '测试模板.xlsx')
  await page.waitForTimeout(3500)
  P('  发现域: ' + (await page.locator('.node-editor .xrp-item').count()))

  P('\n=== 2. 点「在线编辑」===')
  const btn = page.locator('.node-editor .xrp-foot button', { hasText: '在线编辑' })
  P('  按钮存在: ' + ((await btn.count()) > 0 ? '✅' : '❌'))
  P('  按钮可用: ' + ((await btn.isEnabled()) ? '✅' : '❌'))
  await btn.click()
  await page.waitForTimeout(2500)

  P('  /office/* 请求: ' + JSON.stringify(uploads))
  const uploadOk = uploads.some((u) => u.startsWith('200') && u.includes('/office/upload'))
  P('  /office/upload 成功: ' + (uploadOk ? '✅' : '❌'))

  const hasModal = await page.locator('.oo-modal').count()
  P('  弹层出现: ' + (hasModal > 0 ? '✅' : '❌'))

  // 未填 DS → 应该给人话
  const hint = await page.locator('.oo-err').count()
  if (hint) {
    const t = await page.locator('.oo-err').first().innerText()
    P('  未配置 DS 时的提示: ' + t.replace(/\n/g, ' ').slice(0, 120))
    P('  是人话（提到部署文档）: ' + (/部署|Document Server 地址/.test(t) ? '✅' : '❌'))
  } else {
    P('  ⚠ 没有 .oo-err（可能已配置过 DS）')
  }
  await page.screenshot({ path: DIR + 'oo-1-nods.png' })

  P('\n=== 3. 填一个不可达的 DS 地址 ===')
  await page.locator('.oo-settings input').first().fill('http://192.0.2.1:8080')
  await page.waitForTimeout(4000)
  const errTxt = await page.locator('.oo-err').first().innerText().catch(() => '(无)')
  P('  报错文本: ' + errTxt.replace(/\n/g, ' ').slice(0, 160))
  P('  提到 Document Server 不可达: ' + (/不可达|dsUrl/.test(errTxt) ? '✅' : '❌'))
  await page.screenshot({ path: DIR + 'oo-2-unreachable.png' })

  P('\n=== 4. 控制台错误 ===')
  const real = errors.filter((e) => !/favicon|404/.test(e))
  P(real.length === 0 ? '  ✅ 无未捕获异常' : '  ❌ ' + real.length + ' 条: ' + real.slice(0, 3).map((e) => e.slice(0, 120)).join(' | '))

  await page.evaluate((o) => {
    const s = window.__store.getState()
    for (const id of [o.pid, o.xid]) s.purgeTable(id)
    localStorage.removeItem('client-base.office.ds')
  }, ids)
  P('\n=== 5. 清理完成 ===')
} catch (e) {
  P('\n!! 抛错: ' + (e?.message ?? String(e)))
} finally {
  await browser.close()
  fs.writeFileSync(DIR + 'oo-check.txt', log.join('\n'), 'utf8')
  console.log('\n报告 → sample-template/oo-check.txt')
}
