/* 剪贴板粘贴图片循环测试：
 * 1) 合成 ClipboardEvent('paste') 带 image/png → dispatch 到编辑器
 * 2) 轮询文档模型数 drawing（图片元素）
 * 3) 循环重试 + 每轮抓 /upload 请求痕迹，直到成功或轮次用尽
 * 用法：node scripts/clipboard-img-test.mjs [最大轮次=6] */
import fs from 'node:fs'

const NODE_ID = process.argv[2] || 'ooword_1793336980602'
const MAX = Number(process.argv[3] || 6)
/* 32×32 红色 PNG（真实可解码） */
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFklEQVR42mP8z8Dwn4GBgYGRAQIAABQgAsfV1%2B8AAAAASUVORK5CYII='
const PNG_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFklEQVR42mP8z8Dwn4GBgYGRAQIAABQgAsfV1+8AAAAASUVORK5CYII='

const list = await (await fetch('http://127.0.0.1:9334/json')).json()
const page = list.find((t) => t.type === 'page' && (t.url || '').includes('10215'))
if (!page) { console.error('NO-PAGE'); process.exit(1) }
const ws = new WebSocket(page.webSocketDebuggerUrl)
await new Promise((r, j) => { ws.onopen = r; ws.onerror = j })
let seq = 0
const pend = new Map()
const netLog = []
ws.onmessage = (ev) => {
  const d = JSON.parse(ev.data)
  if (d.id && pend.has(d.id)) { pend.get(d.id)(d); pend.delete(d.id) }
  if (d.method === 'Network.requestWillBeSent') {
    const u = d.params.request.url
    if (u.includes('/upload') || u.includes('/image')) netLog.push(d.params.request.method + ' ' + u.slice(0, 90))
  }
}
const ev = (expr, awaitFlag = true, timeoutMs = 45000) => {
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

/* 强制刷新会话：关+开编辑器（劣化的 callCommand 队列只有重开能解） */
console.log('刷新编辑器会话…')
await ev('(function(){ var el = document.querySelector(".react-flow"); if (!el) return "no-canvas"; var fk = Object.keys(el).filter(function(k){ return k.indexOf("__reactFiber") === 0 })[0]; var f = el[fk]; var app = null; while (f && !app) { if (f.tag === 0 && f.memoizedState && typeof f.memoizedState === "object") { var hh = f.memoizedState, g2 = 0; while (hh && g2 < 300) { var a2 = hh, b2 = a2.next, c2 = b2 && b2.next, d2 = c2 && c2.next; if (a2 && b2 && c2 && d2 && typeof a2.memoizedState === "string" && a2.memoizedState.length < 40 && typeof b2.memoizedState === "string" && (b2.memoizedState === "schema" || b2.memoizedState === "data") && typeof c2.memoizedState === "number" && c2.memoizedState === 320 && typeof d2.memoizedState === "string" && d2.memoizedState.length < 60) { app = f; break } hh = hh.next; g2++ } } f = f.return } if (!app) return "no-app"; var h = app.memoizedState, guard = 0; while (h && guard < 300) { var a = h, b = a.next, c = b && b.next, d = c && c.next; if (a && b && c && d && typeof a.memoizedState === "string" && a.memoizedState.length < 40 && typeof b.memoizedState === "string" && (b.memoizedState === "schema" || b.memoizedState === "data") && typeof c.memoizedState === "number" && c.memoizedState === 320 && typeof d.memoizedState === "string" && d.memoizedState.length < 60) { d.queue.dispatch(""); return "CLOSED" } h = h.next; guard++ } return "nf" })()', false)
await new Promise(function(r){ setTimeout(r, 4000) })
await ev('(function(){ var el = document.querySelector(".react-flow"); if (!el) return "no-canvas"; var fk = Object.keys(el).filter(function(k){ return k.indexOf("__reactFiber") === 0 })[0]; var f = el[fk]; var app = null; while (f && !app) { if (f.tag === 0 && f.memoizedState && typeof f.memoizedState === "object") { var hh = f.memoizedState, g2 = 0; while (hh && g2 < 300) { var a2 = hh, b2 = a2.next, c2 = b2 && b2.next, d2 = c2 && c2.next; if (a2 && b2 && c2 && d2 && typeof a2.memoizedState === "string" && a2.memoizedState.length < 40 && typeof b2.memoizedState === "string" && (b2.memoizedState === "schema" || b2.memoizedState === "data") && typeof c2.memoizedState === "number" && c2.memoizedState === 320 && typeof d2.memoizedState === "string" && d2.memoizedState.length < 60) { app = f; break } hh = hh.next; g2++ } } f = f.return } if (!app) return "no-app"; var h = app.memoizedState, guard = 0; while (h && guard < 300) { var a = h, b = a.next, c = b && b.next, d = c && c.next; if (a && b && c && d && typeof a.memoizedState === "string" && a.memoizedState.length < 40 && typeof b.memoizedState === "string" && (b.memoizedState === "schema" || b.memoizedState === "data") && typeof c.memoizedState === "number" && c.memoizedState === 320 && typeof d.memoizedState === "string" && d.memoizedState.length < 60) { d.queue.dispatch("' + NODE_ID + '"); return "OPENED" } h = h.next; guard++ } return "nf" })()', false)
console.log('编辑器已重开，等引擎 35s…')
await new Promise(function(r){ setTimeout(r, 35000) })

/* 开网络监听 */
ws.send(JSON.stringify({ id: ++seq, method: 'Network.enable' }))
/* 引擎就绪检查 */
const ready = await ev(`(function(){ try { var poc = document.querySelector("iframe"); var w2 = poc.contentWindow.document.querySelector("iframe[name=frameEditor]").contentWindow; return !!(w2.Asc && w2.Asc.editor) } catch(e) { return "err:" + String(e).slice(0, 60) } })()`, false)
console.log('引擎就绪:', ready)
if (ready !== true) { console.log('引擎未就绪，中止'); process.exit(1) }

/* 当前 drawing 基线 */
const countDrawings = `(function(){
  var r = null
  try {
    var poc = document.querySelector("iframe")
    var w2 = poc.contentWindow.document.querySelector("iframe[name=frameEditor]").contentWindow
    var eng = w2.Asc && w2.Asc.editor
    eng.callCommand(function(){
      try {
        var dc = Api.GetDocument()
        var n = 0
        for (var i = 0; i < dc.GetElementsCount(); i++) {
          var e = dc.GetElement(i)
          if (!e || e.GetClassType !== "function") continue
          var rn = e.GetElementsCount ? e.GetElementsCount() : 0
          for (var j = 0; j < rn; j++) {
            var el = e.GetElement(j)
            if (el && el.GetClassType && String(el.GetClassType()).toLowerCase().indexOf("draw") >= 0) n++
          }
        }
        try { window.__ooDrawCount = n } catch(eg){}
        return n
      } catch (eX) { return "err:" + String(eX).slice(0, 50) }
    }, false, false, function(v){ if (r === null) { r = v } })
    return new Promise(function(res){
      var polls = 0
      var t = setInterval(function(){
        polls++
        try { var poc0 = document.querySelector('iframe'); var e0 = poc0.contentWindow.document.querySelector('iframe[name=frameEditor]').contentWindow; var v = e0.__ooDrawCount } catch(e){ v = undefined }
        if (v !== undefined && v !== null) { clearInterval(t); res(v) }
        if (polls > 16) { clearInterval(t); res("fail") }
      }, 300)
    })
  } catch (eY) { return "err:" + String(eY).slice(0, 60) }
})()`

const base = await ev(countDrawings)
console.log('基线 drawing 数:', base)

let ok = false
for (let round = 1; round <= MAX && !ok; round++) {
  console.log(`\n===== 第 ${round} 轮 =====`)
  netLog.length = 0
  /* 合成 paste 事件（clipboardData 带 image/png） */
  const pasteExpr = `(function(){
    try {
      var poc = document.querySelector("iframe")
      var w2 = poc.contentWindow.document.querySelector("iframe[name=frameEditor]").contentWindow
      var pngB64 = "${PNG_B64}"
      var bin = atob(pngB64)
      var u8 = new Uint8Array(bin.length)
      for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
      var dt = new DataTransfer()
      dt.items.add(new File([u8], "clipboard.png", { type: "image/png" }))
      var ev = new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true })
      var target = w2.document.getElementById("editor_sdk") || w2.document.activeElement || w2.document.body
      target.dispatchEvent(ev)
      w2.document.dispatchEvent(ev)
      return "dispatched"
    } catch (e) { return "err:" + String(e).slice(0, 80) }
  })()`
  const dr = await ev(pasteExpr, false)
  console.log('paste 事件:', dr)
  /* 等待处理 + 轮询 drawing 计数（最多 20s） */
  for (let w = 0; w < 10; w++) {
    await new Promise((r) => setTimeout(r, 2000))
    const cnt = await ev(countDrawings)
    console.log(`  +${(w + 1) * 2}s drawing=${cnt} 网络上传: ${netLog.length ? netLog.join(' | ') : '（无）'}`)
    if (cnt !== 'fail' && base !== 'fail' && Number(cnt) > Number(base)) { ok = true; break }
  }
  if (ok) break
  /* 每轮之间如果网络有 upload 但失败，先看返回 */
}

console.log('\n结果:', ok ? '✅ 剪贴板图片粘贴成功（文档模型出现 drawing）' : '❌ 未成功——网络痕迹: ' + (netLog.length ? netLog.join(' | ') : '无 /upload 请求'))
