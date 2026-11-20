/**
 * 端到端验证：Excel 域绑定全链路（真实浏览器）
 *
 * 链路：
 *   参数节点(3 行数据) ──数据流──▶ xlsx 节点
 *   xlsx 节点 ← 导入 测试模板.xlsx（浏览器里用 office-kit 解析）
 *     ├ 命名区域 工程名称 / 设计流量  → 自动同名匹配
 *     └ Table 材料表              → 行级绑定 + 列映射 + 计算列公式
 *   点「生成 Excel」→ 下载产物 → 读回校验
 *
 * 关键点：office-kit 的 xlsx 读写必须能在浏览器里跑通（不是只有 Node 能跑）。
 */
import fs from 'node:fs'
import { chromium } from 'playwright-core'

const BASE = process.env.SMOKE_URL || 'http://127.0.0.1:5173'
const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
const DIR = new URL('../sample-template/', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')

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
  await page.waitForTimeout(2500)

  // ── 0. 清掉上次遗留 ──
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

  // ── 1. 建参数节点（上游数据）──
  P('=== 1. 建上游参数节点 ===')
  const ids = await page.evaluate(() => {
    const st = window.__store.getState()
    const t = Date.now()
    const pid = 'param_' + t
    const xid = 'xlsx_' + t
    const F = (name, type) => ({ id: name, name, type })

    st.addTable({
      id: pid,
      name: '工程量与参数',
      kind: 'param',
      fields: [
        F('工程名称', 'text'),
        F('设计流量', 'number'),
        F('名称', 'text'),
        F('数量', 'number'),
        F('单价', 'number'),
      ],
      config: {},
    })
    st.setTableData(
      pid,
      [
        F('工程名称', 'text'),
        F('设计流量', 'number'),
        F('名称', 'text'),
        F('数量', 'number'),
        F('单价', 'number'),
      ],
      [
        { 工程名称: '海南陵水英州镇燃气管道工程', 设计流量: 1200, 名称: '钢管 DN200', 数量: 100, 单价: 25.5 },
        { 名称: '钢管 DN100', 数量: 40, 单价: 18.2 },
        { 名称: '阀门 DN50', 数量: 12, 单价: 260 },
      ],
    )
    st.setNodePosition(pid, { x: 80, y: 420 })

    // xlsx 节点 + 连线（数据流）
    st.addTable({ id: xid, name: 'Excel 测试簿', kind: 'xlsx', fields: [], config: { xlsxMap: {} } })
    st.setNodePosition(xid, { x: 620, y: 420 })
    st.setNodeConfig(xid, { inputs: [pid], linkedInputs: [pid] })
    return { pid, xid }
  })
  P('  上游 param: ' + ids.pid + ' / xlsx: ' + ids.xid)
  await page.waitForTimeout(1200)

  // ── 2. 打开 xlsx 编辑器 ──
  P('\n=== 2. 打开 Excel 编辑器 ===')
  await page.locator(`.react-flow__node[data-id="${ids.xid}"] .rf-node--xlsx`).dblclick()
  await page.waitForTimeout(1000)
  P('  编辑器打开: ' + ((await page.locator('.node-editor-mask').count()) > 0 ? '✅' : '❌'))
  const tplText0 = await page.locator('.node-editor').innerText()
  P('  初始状态含「未选择」: ' + (tplText0.includes('未选择') ? '✅' : '❌'))
  // 编辑器内面板与属性栏面板是两个实例（两个入口，设计如此）——这里必须限定在编辑器内
  P('  编辑器内 xrp 面板数: ' + (await page.locator('.node-editor .xrp').count()))
  P('  页面全局 xlsx file input 数: ' + (await page.locator('.xrp input[type="file"][accept=".xlsx"]').count()) + '（编辑器 + 属性栏各一）')

  // ── 3. 导入模板（走真实的 file input，验证浏览器里 office-kit 能解析）──
  P('\n=== 3. 导入 .xlsx 模板（浏览器内解析）===')
  await page.locator('.node-editor .xrp input[type="file"][accept=".xlsx"]').setInputFiles(DIR + '测试模板.xlsx')
  await page.waitForTimeout(3500)

  const slotCount = await page.locator('.node-editor .xrp-item').count()
  P('  发现域数量: ' + slotCount + (slotCount === 3 ? ' ✅ (2 命名区域 + 1 表格)' : ' ⚠ 期望 3'))
  const slotKeys = await page.locator('.node-editor .xrp-key').allInnerTexts()
  P('  域清单: ' + JSON.stringify(slotKeys))
  const kinds = await page.locator('.node-editor .xrp-kind').allInnerTexts()
  P('  域类型: ' + JSON.stringify(kinds))

  const panelText = await page.locator('.node-editor').innerText()
  const errBox = await page.locator('.node-editor .xrp-err').count()
  if (errBox) P('  ❌ 面板报错: ' + (await page.locator('.node-editor .xrp-err').first().innerText()))
  P('  含「发现 3 个域」: ' + (/发现\s*3\s*个域/.test(panelText) ? '✅' : '❌ ' + (panelText.match(/发现[^\n]*/) || [''])[0]))

  // 自动同名匹配是否生效
  const ctlValues = await page.locator('.node-editor .xrp-item').evaluateAll((items) =>
    items.map((it) => {
      const key = it.querySelector('.xrp-key')?.textContent ?? ''
      const from = it.querySelector('.xrp-ctl select')?.value ?? ''
      const field = it.querySelectorAll('.xrp-ctl select')[1]
      const fid = field ? field.value : ''
      const agg = it.querySelectorAll('.xrp-ctl select')[2]
      return { key, from, fid, agg: agg ? agg.value : '' }
    }),
  )
  P('  自动匹配结果: ' + JSON.stringify(ctlValues))
  await page.screenshot({ path: DIR + 'excel-e2e-1-slots.png' })

  // ── 4. 配 Table 的行级绑定（列映射 + 计算列）──
  //    UI 上串联了 3 个下拉 + 逐列 select/input，用 playwright 逐个点代价高且脆弱；
  //    这里直接写 config（等价于用户在面板上选完），再验证面板确实反映出来。
  P('\n=== 4. 配置 Table 行级绑定 ===')
  await page.evaluate(
    (xid) => {
      const s = window.__store.getState()
      const cur = s.tables[xid]?.config?.xlsxMap ?? {}
      s.setNodeConfig(xid, {
        xlsxMap: {
          ...cur,
          材料表: {
            from: 'rows',
            fieldId: '名称',
            columns: { 名称: '名称', 数量: '数量', 单价: '单价' },
            formulas: { 合价: '={数量}*{单价}' },
          },
        },
      })
    },
    ids.xid,
  )
  await page.waitForTimeout(1200)
  const afterText = await page.locator('.node-editor').innerText()
  P('  面板显示行数: ' + ((afterText.match(/(\d+) 行（已映射列）/) || ['', '未显示'])[1]))
  const colInputs = await page.locator('.node-editor .xrp-formula').evaluateAll((els) => els.map((e) => e.value))
  P('  计算列公式输入框: ' + JSON.stringify(colInputs))
  await page.screenshot({ path: DIR + 'excel-e2e-2-bound.png' })

  // ── 5. 生成 → 捕获下载 ──
  P('\n=== 5. 生成 Excel ===')
  const genBtn = page.locator('.node-editor .xrp-foot button', { hasText: '生成 Excel' })
  P('  按钮可用: ' + ((await genBtn.isEnabled()) ? '✅' : '❌'))
  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: 30000 }).catch(() => null),
    genBtn.click(),
  ])
  await page.waitForTimeout(2000)

  let outPath = ''
  if (download) {
    outPath = DIR + 'excel-e2e-out.xlsx'
    await download.saveAs(outPath)
    P('  ✅ 已捕获下载: ' + download.suggestedFilename() + ' → ' + fs.statSync(outPath).size + ' 字节')
  } else {
    P('  ⚠ 未捕获到下载事件')
  }

  const warnText = await page.locator('.node-editor .xrp-warn').count()
  if (warnText) P('  警告: ' + (await page.locator('.node-editor .xrp-warn').innerText()).replace(/\n/g, ' | '))
  const okText = await page.locator('.node-editor .xrp-ok').count()
  if (okText) P('  结果栏: ' + (await page.locator('.node-editor .xrp-ok').innerText()))
  const genErr = await page.locator('.node-editor .xrp-err').count()
  if (genErr) P('  ❌ 生成报错: ' + (await page.locator('.node-editor .xrp-err').first().innerText()))
  await page.screenshot({ path: DIR + 'excel-e2e-3-generated.png' })

  // ── 6. 控制台错误 ──
  P('\n=== 6. 控制台错误 ===')
  const real = errors.filter((e) => !/favicon|404/.test(e))
  P(real.length === 0 ? '  ✅ 无（office-kit 在浏览器里全程无报错）' : '  ❌ ' + real.length + ' 条：')
  real.slice(0, 10).forEach((e) => P('    ' + e.slice(0, 240)))

  // ── 7. 清理 ──
  await page.evaluate(
    (o) => {
      const s = window.__store.getState()
      for (const id of [o.pid, o.xid]) s.purgeTable(id)
    },
    ids,
  )
  P('\n=== 7. 清理完成 ===')

  // 把产物路径写出来给下一步校验
  fs.writeFileSync(DIR + 'excel-e2e-state.json', JSON.stringify({ outPath }, null, 1), 'utf8')
} catch (e) {
  P('\n!! 脚本抛错: ' + (e?.message ?? String(e)))
  P(String(e?.stack ?? '').split('\n').slice(0, 6).join('\n'))
} finally {
  await browser.close()
  fs.writeFileSync(DIR + 'excel-e2e-check.txt', log.join('\n'), 'utf8')
  console.log('\n报告 → sample-template/excel-e2e-check.txt')
}
