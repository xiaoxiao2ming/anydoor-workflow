/* 引擎内探针：GetSpacingLine 返回结构 + 行距提取验证（写入结果文件供主进程读取） */
import fs from 'node:fs'

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
const ev = (expr, awaitFlag = true, timeoutMs = 60000) => {
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

const expr = `(async function(){
  var poc = document.querySelector("iframe")
  var w2 = poc.contentWindow.document.querySelector("iframe[name=frameEditor]").contentWindow
  var eng = w2.Asc && w2.Asc.editor
  if (!eng) return { err: "no-engine" }
  try { w2.__ooCmdResult = undefined } catch(e){}
  return await new Promise(function(res){
    var done = false
    eng.callCommand(function(){
      var probe = {}
      try {
        var doc = Api.GetDocument()
        for (var idx = 3; idx <= 4; idx++) {
          var p = doc.GetElement(idx)
          if (!p || (p.GetClassType && p.GetClassType() !== "paragraph")) continue
          var q = {}
          try { var v = p.GetSpacingLine(); q.line = v; q.lineType = typeof v } catch (e1) { q.lineErr = String(e1).slice(0, 60) }
          try { q.rule = p.GetSpacingLineRule ? String(p.GetSpacingLineRule()) : "none" } catch (e2) { q.ruleErr = String(e2).slice(0, 60) }
          try { q.spB = p.GetSpacingBefore ? p.GetSpacingBefore() : null } catch (e3) {}
          try { q.spA = p.GetSpacingAfter ? p.GetSpacingAfter() : null } catch (e4) {}
          probe["p" + idx] = q
        }
      } catch (eX) { probe.err = String(eX).slice(0, 80) }
      try { window.__ooCmdResult = probe } catch (eG) {}
      return probe
    }, false, false, function(v){ if (!done) { done = true; res({ via: "cb", probe: v }) } })
    var polls = 0
    var t = setInterval(function(){
      polls++
      if (done) { clearInterval(t); return }
      var v = null
      try { v = w2.__ooCmdResult } catch(e){}
      if (v !== undefined && v !== null) { done = true; clearInterval(t); res({ via: "global", probe: v }) }
      if (polls > 20) { done = true; clearInterval(t); res({ via: "fail" }) }
    }, 300)
  })
})()`

try {
  const r = await ev(expr)
  fs.writeFileSync('probe_spacing.json', JSON.stringify(r, null, 2))
  console.log(JSON.stringify(r, null, 2))
} catch (e) {
  console.error('ERR', String(e).slice(0, 150))
  process.exit(1)
}
