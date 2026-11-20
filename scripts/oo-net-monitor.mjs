/**
 * OO 加载网络监听（40 秒窗口）：抓 /doc/ 通道与全部失败请求
 * 用法：node scripts/oo-net-monitor.mjs   ← 先跑这个，再重开电子表格节点
 */
const CDP_HTTP = 'http://127.0.0.1:9334'
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

async function main () {
  const list = await (await fetch(`${CDP_HTTP}/json`)).json()
  const page = list.find(t => t.type === 'page')
  const ws = new WebSocket(page.webSocketDebuggerUrl)
  await new Promise((r, j) => { ws.onopen = r; ws.onerror = j })

  let seq = 0
  const pending = new Map()
  const requests = new Map()   // requestId -> {url, status}
  const docEvents = []
  const send = (method, params = {}) => new Promise((resolve) => {
    const id = ++seq
    pending.set(id, resolve)
    ws.send(JSON.stringify({ id, method, params }))
  })
  ws.onmessage = (ev) => {
    const d = JSON.parse(ev.data)
    if (d.id && pending.has(d.id)) { pending.get(d.id)(d); pending.delete(d.id); return }
    if (d.method === 'Network.requestWillBeSent') {
      const u = d.params.request.url
      requests.set(d.params.requestId, { url: u })
      if (u.includes('/doc/') || u.includes('socket')) docEvents.push({ t: new Date().toISOString().slice(11, 23), ev: 'req', u: u.slice(-90) })
    } else if (d.method === 'Network.responseReceived') {
      const r = requests.get(d.params.requestId)
      if (r) {
        r.status = d.params.response.status
        const u = r.url
        if (d.params.response.status >= 400 || u.includes('/doc/')) {
          docEvents.push({ t: new Date().toISOString().slice(11, 23), ev: 'res', s: d.params.response.status, u: u.slice(-90) })
        }
      }
    } else if (d.method === 'Network.loadingFailed') {
      const r = requests.get(d.params.requestId)
      if (r) docEvents.push({ t: new Date().toISOString().slice(11, 23), ev: 'FAIL', u: r.url.slice(-90), e: d.params.errorText })
    }
  }

  await send('Network.enable')
  console.log('[monitor] Network 已启用 — 未来 40 秒请重开电子表格节点并保持打开…')

  for (let i = 1; i <= 18; i++) {
    await sleep(5000)
    const r = await send('Runtime.evaluate', {
      expression: `(function(){
        const f = document.querySelector('iframe[name=frameEditor]')
        if (!f) return JSON.stringify({t:'${i * 5}s', frame:'none'})
        const w = f.contentWindow
        return JSON.stringify({t:'${i * 5}s', docsAPI: !!w.DocsAPI, ready: w.document && w.document.readyState})
      })()`,
      returnByValue: true,
    })
    console.log('[monitor]', r.result && r.result.result && r.result.result.value)
  }

  console.log('\n=== /doc 通道与失败请求 ===')
  docEvents.slice(0, 40).forEach(e => console.log(JSON.stringify(e)))
  if (!docEvents.length) console.log('(无 /doc 相关流量！编辑器根本没尝试连接 DS)')

  const r2 = await send('Runtime.evaluate', { expression: 'JSON.stringify((window.__ooErrs||[]).slice(0,20))', returnByValue: true })
  console.log('\n=== __ooErrs ===')
  try { console.log(JSON.stringify(JSON.parse(r2.result.result.value), null, 1)) } catch(_) { console.log(r2.result.result.value) }
  process.exit(0)
}
main().catch((e) => { console.error('monitor ERR:', e.message); process.exit(1) })
