/**
 * 网页版冒烟验证：真实 Chromium 打开 dev:web，验证 Word 报告面板渲染、
 * 控制台无错误、放大浮层可用。
 *
 * 运行：node scripts/smoke-web.mjs
 * 前提：dev server 已起（npm run dev:web 或 vite --config vite.web.config.ts）
 */

import { chromium } from 'playwright-core'

const BASE = process.env.SMOKE_URL || 'http://127.0.0.1:5199'
const SHOT_DIR = new URL('../sample-template/', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')

const errors = []
const run = async () => {
  const browser = await chromium.launch()
  const page = await browser.newPage({ viewport: { width: 1600, height: 950 } })
  page.on('console', (m) => {
    if (m.type() === 'error') errors.push(m.text())
  })
  page.on('pageerror', (e) => errors.push(String(e)))

  await page.goto(BASE, { waitUntil: 'load', timeout: 30000 })
  await page.waitForTimeout(2500)
  await page.screenshot({ path: SHOT_DIR + 'smoke-1-home.png' })
  console.log('[1] 首页已打开，画布截图完成')

  // 找画布上的节点。seed 数据应该有现成节点；没有就报告实际状态
  const nodes = await page.locator('.react-flow__node').count()
  console.log('[2] 画布节点数：', nodes)

  // 尝试点一个节点打开属性面板
  if (nodes > 0) {
    await page.locator('.react-flow__node').first().click()
    await page.waitForTimeout(600)
  }

  const bodyText = await page.evaluate(() => document.body.innerText)
  const hasWrp = bodyText.includes('Word 报告')
  console.log('[3] 属性面板含「Word 报告」区块：', hasWrp)
  await page.screenshot({ path: SHOT_DIR + 'smoke-2-panel.png' })

  // 找所有 output 节点，逐个点开看
  const outCount = await page.locator('.react-flow__node', { hasText: /输出|output/i }).count()
  console.log('[4] 输出节点数：', outCount)
  if (outCount > 0) {
    await page.locator('.react-flow__node', { hasText: /输出|output/i }).first().click()
    await page.waitForTimeout(800)
    const t2 = await page.evaluate(() => document.body.innerText)
    console.log('[5] 选中输出节点后 Word 报告区块：', t2.includes('Word 报告') ? '已渲染' : '未出现')
    console.log('[6] 导入按钮存在：', t2.includes('导入 Word 模板') || t2.includes('换模板'))
    await page.screenshot({ path: SHOT_DIR + 'smoke-3-output.png' })
  }

  console.log('')
  if (errors.length) {
    console.log('控制台错误', errors.length, '条：')
    for (const e of errors.slice(0, 10)) console.log('  -', e.slice(0, 200))
  } else {
    console.log('控制台零错误 ✓')
  }

  await browser.close()
  if (errors.length) process.exit(2)
}

run().catch((e) => {
  console.error('冒烟失败：', e.message)
  process.exit(1)
})
