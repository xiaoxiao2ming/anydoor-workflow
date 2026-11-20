/**
 * OO 加载全链路计时：程序化跑一次完整加载，拆出每个环节秒数
 * 用法：node scripts/oo-timing-probe.mjs [srcFile]
 */
const CDP_HTTP = 'http://127.0.0.1:9334'
const src = process.argv[2] || 'oo-oosheet-muy1l05s.xlsx'
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

async function main () {
  const list = await (await fetch(`${CDP_HTTP}/json`)).json()
  const page = list.find(t => t.type === 'page')
  const ws = new WebSocket(page.webSocketDebuggerUrl)
  await new Promise((r, j) => { ws.onopen = r; ws.onerror = j })

  let seq = 0
  const pending = new Map()
  const netLog = []
  const send = (method, params = {}) => new Promise((resolve) => {
    const id = ++seq
    pending.set(id, resolve)
    ws.send(JSON.stringify({ id, method, params }))
  })
  ws.onmessage = (ev) => {
    const d = JSON.parse(ev.data)
    if (d.id && pending.has(d.id)) { pending.get(d.id)(d); pending.delete(d.id); return }
    if (d.method === 'Network.requestWillBeSent') {
      netLog.push({ ev: 'req', t: d.params.timestamp, u: d.params.request.url })
    } else if (d.method === 'Network.responseReceived') {
      netLog.push({ ev: 'res', s: d.params.response.status, u: d.params.response.url })
    }
  }
  await send('Network.enable')
  await send('Runtime.enable')

  const T0 = Date.now()
  await send('Runtime.evaluate', {
    expression: `(function(){
      const old = document.getElementById('__ooProbeFrame'); if (old) old.remove();
      const f = document.createElement('iframe');
      f.id = '__ooProbeFrame'; f.name = 'frameEditor';
      f.src = '/poc?src=${src}';
      f.style.cssText = 'position:fixed;left:0;bottom:0;width:100%;height:50%;z-index:999999;border:3px solid red';
      document.body.appendChild(f); return 'ok';
    })()`,
    returnByValue: true,
  })
  console.log('[T+0s] poc iframe 已挂载，开始计时…')

  let docsAPImS = -1, engineFoundMs = -1, docReadyMs = -1
  for (let i = 1; i <= 30; i++) {
    await sleep(500)
    const el = Date.now() - T0
    const r = await send('Runtime.evaluate', {
      expression: `(function(){
        const poc = document.querySelector('#__ooProbeFrame');
        if (!poc) return 'gone';
        const w = poc.contentWindow;
        const st = { el: ${el}, docsAPI: !!w.DocsAPI, ready: w.__pocReady || false };
        const eng = w.document.querySelector('iframe');
        if (eng && eng.contentWindow) {
          const ew = eng.contentWindow;
          st.engineSrc = (eng.src||'').slice(-40);
          st.editorObj = !!(ew.editor || (ew.Asc && ew.Asc.editor));
          st.docReady = !!w.__docReadyFlag;
        }
        return JSON.stringify(st);
      })()`,
      returnByValue: true,
    })
    let st = {}
    try { st = JSON.parse(r.result.result.value) } catch (_) { continue }
    if (st.docsAPI && docsAPImS < 0) { docsAPImS = el; console.log(`[T+${(el / 1000).toFixed(1)}s] DocsAPI 创建完成`) }
    if (st.engineSrc && engineFoundMs < 0) { engineFoundMs = el; console.log(`[T+${(el / 1000).toFixed(1)}s] 引擎 iframe 出现: …${st.engineSrc}`) }
    if (st.docReady && docReadyMs < 0) { docReadyMs = el; console.log(`[T+${(el / 1000).toFixed(1)}s] ★ onDocumentReady — 文档完全打开`) }
    if (docReadyMs > 0) break
  }

  // 引擎资源耗时 top
  const r2 = await send('Runtime.evaluate', {
    expression: `(function(){
      const poc = document.querySelector('#__ooProbeFrame');
      if (!poc) return '{}';
      const w = poc.contentWindow;
      const eng = w.document.querySelector('iframe');
      const all = [];
      const collect = (win, tag) => { try {
        (win.performance.getEntriesByType('resource')||[]).forEach(e => all.push({tag, d: Math.round(e.duration), kb: Math.round((e.transferSize||0)/1024), n: e.name.split('/').slice(-3).join('/'), st: e.responseStatus}))
      } catch(_) {} };
      collect(w, 'poc');
      if (eng && eng.contentWindow) collect(eng.contentWindow, 'engine');
      all.sort((a,b) => b.d - a.d);
      return JSON.stringify({total: all.length, slow: all.slice(0, 12), err: all.filter(x => x.st >= 400)});
    })()`,
    returnByValue: true,
  })
  let data = {}
  try { data = JSON.parse(r2.result.result.value) } catch (_) {}
  console.log(`\n=== 资源总计 ${data.total} 项；最慢 12 项 ===`)
  ;(data.slow || []).forEach(x => console.log(`  ${String(x.d).padStart(6)}ms  ${String(x.kb).padStart(6)}KB  [${x.tag}] ${x.st}  ${x.n}`))
  console.log(`\n=== 加载失败 ===`)
  ;(data.err || []).forEach(x => console.log('  ' + x.st + ' ' + x.n)) || (data.err && !data.err.length && console.log('  (无)'))

  // DS 相关请求耗时（从 netLog 算 /doc /cache /work 的 req→res 对）
  console.log('\n=== DS/文档通道请求 ===')
  const opens = netLog.filter(n => n.ev === 'req' && (/\/doc\/|\/cache\/|\/work\/|convert/.test(n.u)))
  opens.slice(0, 15).forEach(n => {
    const res = netLog.find(m => m.ev === 'res' && m.u === n.u)
    console.log('  ' + (res ? res.s : '...') + '  ' + n.u.slice(-80))
  })
  if (!opens.length) console.log('  (无)')

  await send('Runtime.evaluate', { expression: "document.getElementById('__ooProbeFrame') && document.getElementById('__ooProbeFrame').remove()" })
  console.log('[done] 测试 iframe 已清理')
  process.exit(0)
}
main().catch((e) => { console.error('ERR:', e.message); process.exit(1) })
