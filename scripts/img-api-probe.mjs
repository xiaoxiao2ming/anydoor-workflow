import fs from 'node:fs'
/* 枚举引擎插图 API + 自动触发 + 抓引擎 window 错误 */
const list = await (await fetch('http://127.0.0.1:9334/json')).json()
const page = list.find((t) => t.type === 'page' && (t.url || '').includes('10215'))
if (!page) { console.error('NO-PAGE'); process.exit(1) }
const ws = new WebSocket(page.webSocketDebuggerUrl)
await new Promise((r, j) => { ws.onopen = r; ws.onerror = j })
let seq = 0
const pend = new Map()
const errs = []
ws.onmessage = (ev) => {
  const d = JSON.parse(ev.data)
  if (d.id && pend.has(d.id)) { pend.get(d.id)(d); pend.delete(d.id) }
  if (d.method === 'Runtime.exceptionThrown') {
    const e = d.params.exceptionDetails
    errs.push('EXC: ' + ((e.exception && (e.exception.description || e.exception.value)) || e.text).slice(0, 250))
  }
  if (d.method === 'Runtime.consoleAPICalled' && ['error', 'warning'].includes(d.params.type)) {
    errs.push('CONSOLE: ' + (d.params.args.map((a) => a.value ?? a.description ?? '').join(' ')).slice(0, 250))
  }
}
const ev = (expr, awaitFlag = true) => {
  const id = ++seq
  ws.send(JSON.stringify({ id, method: 'Runtime.evaluate', params: { expression: expr, returnByValue: true, awaitPromise: awaitFlag } }))
  return new Promise((res, rej) => {
    const h = (e) => {
      const m = JSON.parse(e.data)
      if (m.id === id) { ws.removeEventListener('message', h); res(m.result?.result?.value) }
    }
    ws.addEventListener('message', h)
    setTimeout(() => rej(new Error('TO')), 25000)
  })
}
ws.send(JSON.stringify({ id: ++seq, method: 'Runtime.enable' }))
await new Promise((r) => setTimeout(r, 500))

/* 1) 枚举 editor 上带 Image 的方法 */
const r1 = await ev(`(function(){
  try {
    var poc = document.querySelector("iframe")
    var fe = poc.contentWindow.document.querySelector("iframe[name=frameEditor]")
    var w2 = fe.contentWindow
    var eng = w2.Asc && w2.Asc.editor
    if (!eng) return "no-engine"
    var names = []
    for (var k in eng) { try { if (typeof eng[k] === "function" && /image|picture/i.test(k)) names.push(k) } catch (e2) {} }
    var wNames = []
    for (var k in w2) { try { if (/image|picture/i.test(k)) wNames.push(k + ":" + typeof w2[k]) } catch (e3) {} }
    return JSON.stringify({ engImageFns: names, winImageFns: wNames.slice(0, 10) })
  } catch (e) { return "err:" + String(e).slice(0, 80) }
})()`, false)
console.log('插图 API:', r1)
await new Promise((r) => setTimeout(r, 500))

/* 2) 自动触发：api.AddImage（data URL，8x8 红 PNG 干净 base64）+ 挂引擎 onerror */
const clean = fs.readFileSync('clean-img-b64.txt', 'utf8').trim()
const r2 = await ev(`(function(){
  try {
    var poc = document.querySelector("iframe")
    var fe = poc.contentWindow.document.querySelector("iframe[name=frameEditor]")
    var w2 = fe.contentWindow
    var eng = w2.Asc && w2.Asc.editor
    if (!eng) return "no-engine"
    var captured = []
    w2.onerror = function(msg, src, line) { captured.push(String(msg).slice(0, 120) + " @" + line); return false }
    if (w2.document && w2.document.addEventListener) { /* noop */ }
    window.__imgErr = captured
    var b64 = "${clean}"
    var fn = eng.AddImage || eng.asc_addImage || eng.put_ImageContent
    if (typeof fn !== "function") return "no-fn (有: " + Object.keys(eng).filter(function(k){ return /image/i.test(k) }).join(",") + ")"
    fn.call(eng, "data:image/png;base64," + b64)
    return "called"
  } catch (e) { return "err:" + String(e).slice(0, 120) }
})()`, false)
console.log('自动插图:', r2)
await new Promise((r) => setTimeout(r, 8000))
const r3 = await ev(`(function(){ var poc = document.querySelector("iframe"); var fe = poc.contentWindow.document.querySelector("iframe[name=frameEditor]"); var cap = fe.contentWindow.__imgErr || []; return JSON.stringify({ captured: cap, dialogUp: !!fe.contentDocument.querySelector(".asc-window, .message, [class*=dialog]") }) })()`, false)
console.log('8s后错误捕获:', r3)
console.log('全局错误流:', errs.length ? errs.slice(-6).join('\n') : '（无）')
process.exit(0)
