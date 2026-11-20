/**
 * OO 编辑器真实加载诊断 v2：attach 现有页面（Electron 禁建新 target）
 * 用法：node scripts/oo-cdp-probe.mjs [srcFile]
 */
const CDP_HTTP = 'http://127.0.0.1:9334'
const APP = 'http://127.0.0.1:7940'
const src = process.argv[2] || 'oo-oosheet-muy198by.xlsx'
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

async function main () {
  const list = await (await fetch(`${CDP_HTTP}/json`)).json()
  const page = list.find(t => t.type === 'page')
  if (!page) { console.error('无 page target'); process.exit(1) }
  console.log('[probe] attach:', (page.title || '').slice(0, 30), (page.url || '').slice(0, 50))

  const ws = new WebSocket(page.webSocketDebuggerUrl)
  await new Promise((r, j) => { ws.onopen = r; ws.onerror = j })

  let seq = 0
  const pending = new Map()
  const errors = []
  const send = (method, params = {}) => new Promise((resolve) => {
    const id = ++seq
    pending.set(id, resolve)
    ws.send(JSON.stringify({ id, method, params }))
  })
  ws.onmessage = (ev) => {
    const d = JSON.parse(ev.data)
    if (d.id && pending.has(d.id)) { pending.get(d.id)(d); pending.delete(d.id); return }
    if (d.method === 'Runtime.exceptionThrown') {
      const e = d.params.exceptionDetails
      errors.push({ kind: 'exception', text: ((e.exception && (e.exception.description || e.exception.value)) || e.text || '').slice(0, 400), url: (e.url || '').slice(-70), line: e.lineNumber })
    } else if (d.method === 'Runtime.consoleAPICalled' && ['error', 'warning'].includes(d.params.type)) {
      errors.push({ kind: 'console.' + d.params.type, text: d.params.args.map(a => a.value || a.description || '').join(' ').slice(0, 300) })
    } else if (d.method === 'Log.entryAdded' && d.params.entry.level === 'error') {
      errors.push({ kind: 'log.error', text: (d.params.entry.text || '').slice(0, 300), url: (d.params.entry.url || '').slice(-70) })
    }
  }

  await send('Runtime.enable')
  await send('Log.enable')
  await send('Page.enable')

  await send('Page.navigate', { url: `${APP}/poc?src=${src}` })
  console.log('[probe] navigating /poc?src=' + src + ' — 25s 采样…')

  for (let i = 1; i <= 5; i++) {
    await sleep(5000)
    const r = await send('Runtime.evaluate', {
      expression: `(function(){
        const f = document.querySelector('iframe[name=frameEditor]')
        if (!f) return JSON.stringify({t:'${i * 5}s', frame: 'none'})
        let w; try { w = f.contentWindow } catch(e) { return JSON.stringify({t:'${i * 5}s', frame: 'cross-origin'}) }
        let failed = []
        try {
          failed = (w.performance.getEntriesByType('resource')||[]).filter(e => e.responseStatus >= 400)
            .map(e => e.responseStatus + ' ' + e.name.split('/').slice(-2).join('/')).slice(0, 8)
        } catch(e) {}
        return JSON.stringify({t:'${i * 5}s', docsAPI: !!w.DocsAPI, asc: !!w.Asc, ready: w.document && w.document.readyState, resFailed: failed})
      })()`,
      returnByValue: true,
    })
    console.log('[probe]', r.result && r.result.result && r.result.result.value)
  }

  console.log('\n=== 采到的错误 (' + errors.length + ') ===')
  errors.slice(0, 30).forEach((e, i) => console.log(`[${i}] ${e.kind} ${e.url || ''}:${e.line ?? ''}\n    ${e.text}`))
  if (!errors.length) console.log('(无 console 错误/异常)')

  ws.close()
  process.exit(0)
}
main().catch((e) => { console.error('probe ERR:', e.message); process.exit(1) })
