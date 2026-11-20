/* 全程网络监控版：清状态→重开节点→抓 Editor.bin 等关键请求→callCommand 验证 */
const NODE_ID = process.argv[2] || 'ooword_1793336980602'
const list = await (await fetch('http://127.0.0.1:9334/json')).json()
const page = list.find((t) => t.type === 'page' && (t.url || '').includes('10215'))
if (!page) { console.error('NO-PAGE'); process.exit(1) }
const ws = new WebSocket(page.webSocketDebuggerUrl)
await new Promise((r, j) => { ws.onopen = r; ws.onerror = j })
let seq = 0
const pend = new Map()
const reqs = []
const binReqs = []
ws.onmessage = (ev) => {
  const d = JSON.parse(ev.data)
  if (d.id && pend.has(d.id)) { pend.get(d.id)(d); pend.delete(d.id) }
  if (d.method === 'Network.requestWillBeSent') {
    const r = d.params.request
    const line = r.method + ' ' + r.url.slice(0, 130)
    reqs.push(line)
    if (/\.bin|Editor|convert|upload|doc\//i.test(r.url)) binReqs.push(line)
  }
  if (d.method === 'Network.responseReceived') {
    const r = d.params.response
    const line = '    ↳ ' + r.status + ' [' + (r.mimeType || '') + '] ' + (r.headers['content-length'] || '?') + 'B ' + r.url.slice(0, 120)
    reqs.push(line)
    if (/\.bin|Editor|convert|upload|doc\//i.test(r.url)) binReqs.push(line)
  }
  if (d.method === 'Network.loadingFailed') {
    binReqs.push('    ✗✗ 加载失败: ' + (d.params.errorText || '') + ' canceled=' + d.params.canceled + ' blocked=' + d.params.blockedReason)
  }
}
const ev = (expr, awaitFlag = true, timeoutMs = 30000) => {
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
ws.send(JSON.stringify({ id: ++seq, method: 'Network.enable' }))
await new Promise((r) => setTimeout(r, 400))

console.log('== 1) 关+开编辑器（全程抓网络）==')
const closeExpr = `(function(){
  var el = document.querySelector(".react-flow")
  if (!el) return "no-canvas"
  var fk = Object.keys(el).filter(function(k){ return k.indexOf("__reactFiber") === 0 })[0]
  var f = el[fk]
  while (f) {
    if (f.tag === 0 && f.memoizedState && typeof f.memoizedState === "object") {
      var hh = f.memoizedState, g2 = 0
      while (hh && g2 < 300) {
        var a2 = hh, b2 = a2.next, c2 = b2 && b2.next, d2 = c2 && c2.next
        if (a2 && b2 && c2 && d2 && typeof a2.memoizedState === "string" && a2.memoizedState.length < 40 && typeof b2.memoizedState === "string" && (b2.memoizedState === "schema" || b2.memoizedState === "data") && typeof c2.memoizedState === "number" && c2.memoizedState === 320 && typeof d2.memoizedState === "string" && d2.memoizedState.length < 60) { d2.queue.dispatch(""); return "CLOSED" }
        hh = hh.next; g2++
      }
    }
    f = f.return
  }
  return "nf"
})()`
const openExpr = `(function(){
  var el = document.querySelector(".react-flow")
  if (!el) return "no-canvas"
  var fk = Object.keys(el).filter(function(k){ return k.indexOf("__reactFiber") === 0 })[0]
  var f = el[fk]
  while (f) {
    if (f.tag === 0 && f.memoizedState && typeof f.memoizedState === "object") {
      var hh = f.memoizedState, g2 = 0
      while (hh && g2 < 300) {
        var a2 = hh, b2 = a2.next, c2 = b2 && b2.next, d2 = c2 && c2.next
        if (a2 && b2 && c2 && d2 && typeof a2.memoizedState === "string" && a2.memoizedState.length < 40 && typeof b2.memoizedState === "string" && (b2.memoizedState === "schema" || b2.memoizedState === "data") && typeof c2.memoizedState === "number" && c2.memoizedState === 320 && typeof d2.memoizedState === "string" && d2.memoizedState.length < 60) { d2.queue.dispatch("${NODE_ID}"); return "OPENED" }
        hh = hh.next; g2++
      }
    }
    f = f.return
  }
  return "nf"
})()`
console.log(await ev(closeExpr, false))
await new Promise((r) => setTimeout(r, 3000))
console.log(await ev(openExpr, false))
console.log('等 40 秒加载…')
await new Promise((r) => setTimeout(r, 40000))

console.log('== 2) 关键请求（.bin/转换/上传/doc 路径）==')
console.log(binReqs.length ? binReqs.join('\n') : '（无关键请求！）')
console.log('== 3) 全部请求统计 ==')
console.log('总数: ' + reqs.length + '，尾部 6 条:')
console.log(reqs.slice(-6).join('\n'))

console.log('== 4) POC 状态条 ==')
console.log(await ev(`(function(){ var poc = (function(){ var all = [...document.querySelectorAll("iframe")]; return all.find(function(f){ return f.offsetWidth > 100 && f.offsetHeight > 100 }) || all[0] })(); if (!poc) return "无poc"; var st = poc.contentDocument ? poc.contentDocument.getElementById("status") : null; return st ? st.textContent.slice(0, 80) : "无status" })()`, false))

console.log('== 5) callCommand 验证 ==')
await ev(`(function(){ var poc = (function(){ var all = [...document.querySelectorAll("iframe")]; return all.find(function(f){ return f.offsetWidth > 100 && f.offsetHeight > 100 }) || all[0] })(); var fe = poc.contentWindow.document.querySelector("iframe[name=frameEditor]"); var w2 = fe.contentWindow; try { w2.__t = undefined } catch(e){}; w2.Asc.editor.callCommand(function(){ try { window.__t = Api.GetDocument().GetElementsCount() } catch(e){ window.__t = "err:" + String(e).slice(0, 40) }; return 1 }, false, false, function(){}) })()`, false)
await new Promise((r) => setTimeout(r, 5000))
console.log(await ev(`(function(){ var poc = (function(){ var all = [...document.querySelectorAll("iframe")]; return all.find(function(f){ return f.offsetWidth > 100 && f.offsetHeight > 100 }) || all[0] })(); var fe = poc.contentWindow.document.querySelector("iframe[name=frameEditor]"); var v = fe.contentWindow.__t; return v === undefined ? "❌ fn 未执行" : "✅ fn 执行，结果=" + JSON.stringify(v) })()`, false))
process.exit(0)
