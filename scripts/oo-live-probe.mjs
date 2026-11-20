/**
 * OO 加载决定性复现：在用户页面内程序化挂 poc iframe（文件已在工作区），
 * 监听网络 + 收集错误，15 秒后抓 frameEditor 内部状态。
 * 用法：node scripts/oo-live-probe.mjs [srcFile]
 */
const CDP_HTTP = 'http://127.0.0.1:9334'
const src = process.argv[2] || 'oo-oosheet-muy1gm4t.xlsx'
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

async function main () {
  const list = await (await fetch(`${CDP_HTTP}/json`)).json()
  const page = list.find(t => t.type === 'page')
  const ws = new WebSocket(page.webSocketDebuggerUrl)
  await new Promise((r, j) => { ws.onopen = r; ws.onerror = j })

  let seq = 0
  const pending = new Map()
  const fails = []
  const send = (method, params = {}) => new Promise((resolve) => {
    const id = ++seq
    pending.set(id, resolve)
    ws.send(JSON.stringify({ id, method, params }))
  })
  ws.onmessage = (ev) => {
    const d = JSON.parse(ev.data)
    if (d.id && pending.has(d.id)) { pending.get(d.id)(d); pending.delete(d.id); return }
    if (d.method === 'Network.responseReceived') {
      const r = d.params.response
      if (r.status >= 400) fails.push(r.status + ' ' + r.url.split('/').slice(-3).join('/'))
    } else if (d.method === 'Network.loadingFailed') {
      fails.push('FAIL ' + (d.params.errorText || '').slice(0, 40))
    }
  }

  await send('Network.enable')
  // 插入测试 iframe（真实 poc 加载链）
  const inject = await send('Runtime.evaluate', {
    expression: `(function(){
      const old = document.getElementById('__ooProbeFrame'); if (old) old.remove();
      const f = document.createElement('iframe');
      f.id = '__ooProbeFrame'; f.name = 'frameEditor';
      f.src = '/poc?src=${src}&theme=dark';
      f.style.cssText = 'position:fixed;left:0;bottom:0;width:100%;height:55%;z-index:999999;border:3px solid red;background:#fff';
      document.body.appendChild(f);
      return 'iframe injected, src=' + f.src.slice(-60);
    })()`,
    returnByValue: true,
  })
  console.log('[live]', inject.result && inject.result.result && inject.result.result.value)

  for (let i = 1; i <= 4; i++) {
    await sleep(4000)
    const r = await send('Runtime.evaluate', {
      expression: `(function(){
        const f = document.querySelector('#__ooProbeFrame')
        if (!f) return JSON.stringify({t:'${i * 4}s', frame:'gone'})
        let w; try { w = f.contentWindow } catch(e) { return JSON.stringify({t:'${i * 4}s', frame:'cross'}) }
        if (!w.document || !w.document.body) return JSON.stringify({t:'${i * 4}s', frame:'no-doc'})
        const out = {t:'${i * 4}s', docsAPI: !!w.DocsAPI, asc: !!w.Asc, ready: w.document.readyState, st: (w.document.querySelector('#st,.st,[data-st]')||{}).textContent || ''}
        try {
          const res = w.performance.getEntriesByType('resource')||[];
          out.resTotal = res.length;
          out.failed = res.filter(e => e.responseStatus >= 400).map(e => e.responseStatus + ' ' + e.name.split('/').slice(-3).join('/')).slice(0, 10);
          out.errs = (w.__ooErrs||[]).slice(0, 5);
        } catch(e) { out.perfErr = e.message }
        return JSON.stringify(out)
      })()`,
      returnByValue: true,
    })
    console.log('[live]', r.result && r.result.result && r.result.result.value)
  }

  console.log('\n=== 期间网络失败 (' + fails.length + ') ===')
  fails.slice(0, 20).forEach(f => console.log(' ', f))

  // 清理测试 iframe
  await send('Runtime.evaluate', { expression: "document.getElementById('__ooProbeFrame') && document.getElementById('__ooProbeFrame').remove()" })
  console.log('[live] 测试 iframe 已移除')
  process.exit(0)
}
main().catch((e) => { console.error('ERR:', e.message); process.exit(1) })
