/**
 * OO Word 保存链端到端保真回归工具 v2（单目标：宿主页 + postMessage 桥）
 *
 * 前提：应用已启动、某 ooword 节点已双击打开（编辑器就绪）。
 * 流程（CDP 驱动真实 UI）：
 *   A 前置提取  —— 宿主 styledExtract 桥（callCommand 读文档模型 v2）
 *   B 删除字段  —— __ooTestDeleteAsk 桥清空最后一个非空段落
 *   C 保存      —— 点 💾，等保存链完成
 *   D 重开稳定  —— 轮询后置提取有结果（自动重开+引擎就绪）
 *   E 对比      —— 块数 / 每块文字 / 首run格式（font/size/bold）
 * 退出码：0=PASS 1=FAIL
 * 用法：node scripts/check-oo-e2e.mjs
 */
const CDP_HTTP = 'http://127.0.0.1:9334'
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

async function main () {
  const list = await (await fetch(`${CDP_HTTP}/json`)).json()
  const page = list.find((t) => t.type === 'page' && (t.url || '').includes('10215'))
  if (!page) { console.error('[e2e] 未找到应用页面（应用未启动？）'); process.exit(1) }
  console.log('[e2e] attach:', (page.url || '').slice(0, 60))

  const ws = new WebSocket(page.webSocketDebuggerUrl)
  await new Promise((r, j) => { ws.onopen = r; ws.onerror = j })
  let seq = 0
  const pend = new Map()
  ws.onmessage = (ev) => { const d = JSON.parse(ev.data); if (d.id && pend.has(d.id)) { pend.get(d.id)(d); pend.delete(d.id) } }
  const evalJs = async (expression, awaitPromise = true) => {
    const id = ++seq
    ws.send(JSON.stringify({ id, method: 'Runtime.evaluate', params: { expression, awaitPromise, returnByValue: true } }))
    return new Promise((resolve) => {
      const h = (e) => { const d = JSON.parse(e.data); if (d.id === id) { ws.removeEventListener('message', h); resolve(d.result?.result?.value) } }
      ws.addEventListener('message', h)
    })
  }

  /* 通用提取：宿主 styledExtract 桥（ask→reply） */
  const capture = async () => {
    const r = await evalJs(`(function(){
      return new Promise(function(resolve){
        var done = false
        var h = function(e){
          var d = e.data || {}
          if ('__ooStyledExtractData' in d) {
            window.removeEventListener('message', h); done = true
            resolve(JSON.stringify({ paras: Array.isArray(d.__ooStyledExtractData) ? d.__ooStyledExtractData : null, err: d.__ooStyledExtractErr || '', meta: d.__ooStyledExtractMeta || null }))
          }
        }
        window.addEventListener('message', h)
        window.postMessage({ __ooStyledExtractAsk: 1 }, '*')
        setTimeout(function(){ if (!done) { window.removeEventListener('message', h); resolve(JSON.stringify({ paras: null, err: '10s 超时' })) } }, 10000)
      })
    })()`, true)
    try { return JSON.parse(r) } catch { return { paras: null, err: '解析失败: ' + String(r).slice(0, 80) } }
  }

  const blockText = (b) => (b.runs || []).map((r) => r.text || '').join('')

  // A 前置提取
  console.log('\n[A] 前置提取（当前文档模型）…')
  const before = await capture()
  if (before.err || !before.paras) { console.error('[e2e] 前置提取失败:', before.err || '无 paras', '——先双击打开一个 ooword 节点'); process.exit(1) }
  const beforeText = before.paras.map(blockText)
  console.log(`[e2e] 块数=${before.paras.length} 末段="${(beforeText[beforeText.length - 1] || '').slice(0, 36)}"`)

  // B 删除字段
  console.log('\n[B] 删除字段（__ooTestDeleteAsk → 清空最后一个非空段落）…')
  const delRaw = await evalJs(`(function(){
    return new Promise(function(resolve){
      var done = false
      var h = function(e){ var d = e.data || {}; if ('__ooTestDeleteAck' in d) { window.removeEventListener('message', h); done = true; resolve(JSON.stringify(d.__ooTestDeleteAck || { err: '空' })) } }
      window.addEventListener('message', h)
      var f = document.querySelector('iframe[src*="/poc"]')
      if (!f) { resolve(JSON.stringify({ err: '找不到 /poc iframe' })); return }
      f.contentWindow.postMessage({ __ooTestDeleteAsk: 1 }, '*')
      setTimeout(function(){ if (!done) { window.removeEventListener('message', h); resolve(JSON.stringify({ err: '8s 超时' })) } }, 8000)
    })
  })()`, true)
  let del
  try { del = JSON.parse(delRaw) } catch { del = { err: '解析失败' } }
  if (del.err) { console.error('[e2e] 删除失败:', del.err); process.exit(1) }
  console.log(`[e2e] 已清空第 ${del.target + 1} 块（runs=${del.cleared}）`)

  // C 保存
  console.log('\n[C] 点 💾 保存，等保存链完成…')
  await evalJs(`(function(){ var b = document.querySelector('button[title^="保存："]') || document.querySelector('button[title*="保存"]'); if (b) b.click(); return b ? 'clicked' : 'no-button' })()`, false)
  await sleep(5000)

  // D 重开稳定（自动重开 + 引擎就绪 → 后置提取出结果）
  console.log('\n[D] 等自动重开 + 引擎就绪…')
  await sleep(3000)
  let after = null
  for (let i = 0; i < 6; i++) {
    after = await capture()
    if (after.paras && after.paras.length) break
    await sleep(2000)
  }
  if (!after || !after.paras) { console.error('[e2e] 后置提取失败:', (after && after.err) || '无 paras'); process.exit(1) }
  const afterText = after.paras.map(blockText)

  // E 逐块对比
  console.log('\n[E] 对比 ======')
  let pass = true
  if (before.paras.length !== after.paras.length) {
    console.log(`  ✗ 块数变化: ${before.paras.length} → ${after.paras.length}`)
    pass = false
  } else {
    console.log(`  ✓ 块数一致: ${before.paras.length}`)
  }
  for (let i = 0; i < Math.min(before.paras.length, after.paras.length); i++) {
    const b1 = before.paras[i]; const b2 = after.paras[i]
    const t1 = beforeText[i] || ''; const t2 = afterText[i] || ''
    if (i === del.target) {
      if (t2.replace(/\s/g, '') === '') console.log(`  ✓ 块${i + 1}: 已删除（原 "${t1.slice(0, 24)}…" → 空）`)
      else { console.log(`  ✗ 块${i + 1}: 删除未生效，仍有文本 "${t2.slice(0, 40)}"`); pass = false }
      continue
    }
    if (t1 !== t2) { console.log(`  ✗ 块${i + 1} 文本差异:\n      前="${t1.slice(0, 60)}"\n      后="${t2.slice(0, 60)}"`); pass = false; continue }
    const r1 = (b1.runs || [])[0] || {}; const r2 = (b2.runs || [])[0] || {}
    const fDiff = String(r1.font || '') !== String(r2.font || '')
    const sDiff = r1.size != null && r2.size != null && Number(r1.size) !== Number(r2.size)
    const bDiff = !!r1.bold !== !!r2.bold
    if (fDiff || sDiff || bDiff) {
      console.log(`  ✗ 块${i + 1} 格式差异: font "${r1.font}"→"${r2.font}"｜size ${r1.size}→${r2.size}｜bold ${r1.bold}→${r2.bold}`)
      pass = false
    } else {
      console.log(`  ✓ 块${i + 1}: 文本+格式一致 "${t1.slice(0, 24)}"`)
    }
  }
  console.log('\n=====', pass ? '✅ 保真回归通过' : '❌ 存在保真差异', '=====')
  process.exit(pass ? 0 : 1)
}

main().catch((e) => { console.error('[e2e] 异常:', e); process.exit(1) })
