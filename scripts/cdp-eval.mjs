/* CDP 通用求值工具：node scripts/cdp-eval.mjs '<页面JS表达式>' [--await] [--find 159]
 * 表达式在宿主页（127.0.0.1:10215）执行；--await 时等待 Promise；--find 时打印结果 */
const args = process.argv.slice(2)
const awaitFlag = args.includes('--await')
const expr = args.filter((a) => !a.startsWith('--'))[0]

const list = await (await fetch('http://127.0.0.1:9334/json')).json()
const page = list.find((t) => t.type === 'page' && (t.url || '').includes('10215'))
if (!page) { console.error('NO-PAGE'); process.exit(1) }
const ws = new WebSocket(page.webSocketDebuggerUrl)
await new Promise((r, j) => { ws.onopen = r; ws.onerror = j })
let seq = 0
const pend = new Map()
ws.onmessage = (ev) => {
  const d = JSON.parse(ev.data)
  if (d.id && pend.has(d.id)) { pend.get(d.id)(d); pend.delete(d.id) }
}
const id = ++seq
ws.send(JSON.stringify({ id, method: 'Runtime.evaluate', params: { expression: expr, returnByValue: true, awaitPromise: awaitFlag } }))
const d = await new Promise((res) => {
  const h = (e) => {
    const m = JSON.parse(e.data)
    if (m.id === id) { ws.removeEventListener('message', h); res(m) }
  }
  ws.addEventListener('message', h)
  setTimeout(() => res({ err: 'timeout' }), 150000)
})
ws.close()
if (d.err) { console.error('EVAL-' + d.err); process.exit(2) }
if (d.result?.exceptionDetails) {
  console.error('PAGE-EXC:', JSON.stringify(d.result.exceptionDetails.exception?.description || d.result.exceptionDetails).slice(0, 300))
  process.exit(3)
}
console.log(JSON.stringify(d.result?.result?.value ?? null))
