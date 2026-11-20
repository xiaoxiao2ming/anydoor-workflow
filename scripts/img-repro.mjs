/* 插图复现：引擎 API 在用户节点插入 1×1 PNG → 看插入是否成功 → 保存链验证图片命运 */
import fs from 'node:fs'

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const NODE_ID = process.argv[2] || 'ooword_1793336980602'
const MODE = process.argv[3] || 'insert' /* insert | save */

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

if (MODE === 'insert') {
  const expr = `(async function(){
    var poc = document.querySelector("iframe")
    var w2 = poc.contentWindow.document.querySelector("iframe[name=frameEditor]").contentWindow
    var eng = w2.Asc && w2.Asc.editor
    if (!eng) return "no-engine"
    try { w2.__ooImgResult = undefined } catch(e){}
    return await new Promise(function(res){
      var done = false
      eng.callCommand(function(){
        try {
          var dc = Api.GetDocument()
          var p = dc.GetElement(1)
          if (!p || (p.GetClassType && p.GetClassType() !== "paragraph")) { return { err: "not-para" } }
          var oDrawing = Api.CreateImage("data:image/png;base64,${PNG}", 1828800, 1828800)
          p.Push(oDrawing)
          try { window.__ooImgResult = { ok: 1 } } catch(eg){}
          return { ok: 1 }
        } catch (eX) { try { window.__ooImgResult = { err: String(eX) } } catch(eg){}; return { err: String(eX) } }
      }, false, false, function(v){ if (!done) { done = true; res("cb:" + JSON.stringify(v)) } })
      var polls = 0
      var t = setInterval(function(){
        polls++
        if (done) { clearInterval(t); return }
        var v = null
        try { v = w2.__ooImgResult } catch(e){}
        if (v !== undefined && v !== null) { done = true; clearInterval(t); res("global:" + JSON.stringify(v)) }
        if (polls > 16) { done = true; clearInterval(t); res("fail") }
      }, 300)
    })
  })()`
  const r = await ev(expr)
  console.log('插入图片:', r)
  /* 置脏 */
  await ev(`(function(){ var poc = document.querySelector("iframe"); var w2 = poc.contentWindow.document.querySelector("iframe[name=frameEditor]").contentWindow; try { w2.Asc.editor.SetDocumentModified(true) } catch(e){}; try { poc.contentWindow.__setOoDirty(true, "img-test") } catch(e2){}; return "dirty" })()`, false)
  console.log('已置脏——现在去应用里点保存（或跑 save 模式）')
}

if (MODE === 'save') {
  const r = await ev(`(function(){
    var head = document.querySelector(".ne-head")
    var saveBtn = [...head.querySelectorAll("button")].find(function(b){ var t = b.title || ""; return (t.indexOf("保存：") === 0) || (t.indexOf("有未保存的修改") === 0) })
    if (!saveBtn) return "no-btn"
    saveBtn.click()
    return "clicked"
  })()`, false)
  console.log('保存点击:', r)
  /* 等落库 */
  const fp = `(function(){ var t = window.__store.getState().tables["${NODE_ID}"]; var b = t && t.config ? (t.config.ooBytes64 || "") : ""; return b.length })()`
  const before = await ev(fp, false)
  for (let i = 0; i < 60; i++) {
    await new Promise((r) => setTimeout(r, 1000))
    try { const now = await ev(fp, false); if (now !== before) { console.log('字节落库 ✓', before, '→', now); break } } catch { /* busy */ }
  }
  /* 验证产物含图片：解包最新留档查 media */
  const r2 = await ev('(function(){ return "done" })()', false)
  console.log('落库等待结束')
}
ws.close()
