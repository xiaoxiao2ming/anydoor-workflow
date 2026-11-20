/* 整合一气呵成：开编辑器→确认 iframe→基线→剪贴板粘贴→轮询 drawing→判定
 * 全程每步自检，任一步环境失效立即报告（避免在关闭状态下测出假失败）
 * 用法：node scripts/clipboard-once.mjs */
const NODE_ID = process.argv[2] || 'ooword_1793336980602'
const PNG_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFklEQVR42mP8z8Dwn4GBgYGRAQIAABQgAsfV1+8AAAAASUVORK5CYII='

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
const ev = (expr, awaitFlag = true, timeoutMs = 40000) => {
  const id = ++seq
  ws.send(JSON.stringify({ id, method: 'Runtime.evaluate', params: { expression: expr, returnByValue: true, awaitPromise: awaitFlag } }))
  return new Promise((res, rej) => {
    const h = (e) => {
      const m = JSON.parse(e.data)
      if (m.id === id) { ws.removeEventListener('message', h); res(m.result?.result?.value) }
    }
    ws.addEventListener('message', h)
    setTimeout(() => rej(new Error('TIMEOUT')), timeoutMs)
  })
}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

/* 通用：编辑器开关（复用 hooks 滑窗定位） */
const openExpr = (id) => `(function(){
  var el = document.querySelector(".react-flow")
  if (!el) return "no-canvas"
  var fk = Object.keys(el).filter(function(k){ return k.indexOf("__reactFiber") === 0 })[0]
  var f = el[fk]
  var app = null
  while (f && !app) {
    if (f.tag === 0 && f.memoizedState && typeof f.memoizedState === "object") {
      var hh = f.memoizedState, g2 = 0
      while (hh && g2 < 300) {
        var a2 = hh, b2 = a2.next, c2 = b2 && b2.next, d2 = c2 && c2.next
        if (a2 && b2 && c2 && d2 && typeof a2.memoizedState === "string" && a2.memoizedState.length < 40 && typeof b2.memoizedState === "string" && (b2.memoizedState === "schema" || b2.memoizedState === "data") && typeof c2.memoizedState === "number" && c2.memoizedState === 320 && typeof d2.memoizedState === "string" && d2.memoizedState.length < 60) { app = f; break }
        hh = hh.next; g2++
      }
    }
    f = f.return
  }
  if (!app) return "no-app"
  var h = app.memoizedState, guard = 0
  while (h && guard < 300) {
    var a = h, b = a.next, c = b && b.next, d = c && c.next
    if (a && b && c && d && typeof a.memoizedState === "string" && a.memoizedState.length < 40 && typeof b.memoizedState === "string" && (b.memoizedState === "schema" || b.memoizedState === "data") && typeof c.memoizedState === "number" && c.memoizedState === 320 && typeof d.memoizedState === "string" && d.memoizedState.length < 60) { d.queue.dispatch("${id}"); return "OK" }
    h = h.next; guard++
  }
  return "nf"
})()`
const closeExpr = openExpr('').replace(`d.queue.dispatch(""); return "OK"`, `d.queue.dispatch(""); return "CLOSED"`)

/* 环境自检（编辑器 iframe + 引擎 + Api） */
const envExpr = `(function(){
  try {
    var poc = (function(){ var all = [...document.querySelectorAll('iframe')]; return all.find(function(f){ return f.offsetWidth > 100 && f.offsetHeight > 100 }) || all[0] })()
    if (!poc) return JSON.stringify({ alive: false, why: "无 poc iframe" })
    var fe = poc.contentWindow.document.querySelector("iframe[name=frameEditor]")
    if (!fe) return JSON.stringify({ alive: false, why: "无 frameEditor" })
    var w2 = fe.contentWindow
    if (!w2 || !w2.Asc || !w2.Asc.editor) return JSON.stringify({ alive: false, why: "无引擎" })
    return JSON.stringify({ alive: true, Api: typeof w2.Api, cc: typeof (w2.Asc.editor.callCommand) })
  } catch (e) { return JSON.stringify({ alive: false, why: String(e).slice(0, 60) }) }
})()`

/* drawing 计数（引擎 window 全局 + callCommand 返回双通道） */
const drawExpr = `(function(){
  var poc = document.querySelector("iframe")
  if (!poc) return JSON.stringify({ alive: false })
  var fe = poc.contentWindow.document.querySelector("iframe[name=frameEditor]")
  if (!fe) return JSON.stringify({ alive: false })
  var w2 = fe.contentWindow
  var eng = w2.Asc && w2.Asc.editor
  if (!eng || !eng.callCommand) return JSON.stringify({ alive: false })
  try { delete w2.__ooDrawCount } catch (e) { w2.__ooDrawCount = undefined }
  eng.callCommand(function(){
    try {
      var dc = Api.GetDocument()
      var n = 0
      for (var i = 0; i < dc.GetElementsCount(); i++) {
        var e = dc.GetElement(i)
        if (!e || e.GetClassType() !== "paragraph") continue
        var rn = e.GetElementsCount ? e.GetElementsCount() : 0
        for (var j = 0; j < rn; j++) {
          var el = e.GetElement(j)
          if (el && el.GetClassType && String(el.GetClassType()).indexOf("Image") >= 0) n++
        }
      }
      window.__ooDrawCount = n
      return n
    } catch (eX) { window.__ooDrawCount = "err:" + String(eX).slice(0, 50); return "err" }
  }, false, false, function() {})
  return JSON.stringify({ alive: true })
})()`
const pollDraw = `(function(){
  try {
    var poc = (function(){ var all = [...document.querySelectorAll('iframe')]; return all.find(function(f){ return f.offsetWidth > 100 && f.offsetHeight > 100 }) || all[0] })()
    var fe = poc.contentWindow.document.querySelector("iframe[name=frameEditor]")
    var v = fe.contentWindow.__ooDrawCount
    return JSON.stringify({ alive: true, v: v === undefined ? "pending" : v })
  } catch (e) { return JSON.stringify({ alive: false }) }
})()`

/* paste 合成 */
const pasteExpr = `(function(){
  try {
    var poc = (function(){ var all = [...document.querySelectorAll('iframe')]; return all.find(function(f){ return f.offsetWidth > 100 && f.offsetHeight > 100 }) || all[0] })()
    var fe = poc.contentWindow.document.querySelector("iframe[name=frameEditor]")
    var w2 = fe.contentWindow
    var bin = atob("${PNG_B64}")
    var u8 = new Uint8Array(bin.length)
    for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
    var dt = new DataTransfer()
    dt.items.add(new File([u8], "clipboard.png", { type: "image/png" }))
    var ev2 = new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true })
    var target = w2.document.getElementById("editor_sdk") || w2.document.activeElement || w2.document.body
    target.dispatchEvent(ev2)
    w2.document.dispatchEvent(ev2)
    return "dispatched"
  } catch (e) { return "err:" + String(e).slice(0, 80) }
})()`

console.log('== 1) 打开编辑器 ==')
console.log(await ev(openExpr(NODE_ID), false))
await sleep(30000)

console.log('== 2) 环境自检 ==')
const e1 = JSON.parse(await ev(envExpr, false))
console.log(JSON.stringify(e1))
if (!e1.alive) { console.log('❌ 环境未就绪：' + e1.why); process.exit(1) }

console.log('== 3) 基线 drawing 数 ==')
await ev(drawExpr, false)
let base = 'pending'
for (let i = 0; i < 15; i++) {
  await sleep(2000)
  const r = JSON.parse(await ev(pollDraw, false))
  if (!r.alive) { console.log('❌ 环境中途失效'); process.exit(1) }
  if (r.v !== 'pending') { base = r.v; break }
}
console.log('基线:', base)
if (base === 'pending') { console.log('❌ callCommand 通道无响应（fn 未执行）'); process.exit(1) }

console.log('== 4) 合成剪贴板粘贴 ==')
console.log(await ev(pasteExpr, false))

console.log('== 5) 轮询 24s ==')
let ok = false
for (let i = 0; i < 12; i++) {
  await sleep(2000)
  const r = JSON.parse(await ev(pollDraw, false))
  if (!r.alive) { console.log('❌ 环境中途失效'); process.exit(1) }
  console.log(`  +${(i + 1) * 2}s drawing=${r.v}`)
  if (r.v !== 'pending' && r.v !== 'err:' && Number(r.v) > Number(base)) { ok = true; break }
}
console.log(ok ? '\n✅ 剪贴板图片粘贴成功' : '\n❌ 粘贴后文档模型无新图片')
