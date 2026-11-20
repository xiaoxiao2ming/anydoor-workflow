/* 当前可见编辑器的剪贴板粘贴测试 + 网络抓取（多 iframe 定位正确目标） */
const list = await (await fetch('http://127.0.0.1:9334/json')).json()
const page = list.find((t) => t.type === 'page' && (t.url || '').includes('10215'))
if (!page) { console.error('NO-PAGE'); process.exit(1) }
const ws = new WebSocket(page.webSocketDebuggerUrl)
await new Promise((r, j) => { ws.onopen = r; ws.onerror = j })
let seq = 0
const pend = new Map()
const reqs = []
ws.onmessage = (ev) => {
  const d = JSON.parse(ev.data)
  if (d.id && pend.has(d.id)) { pend.get(d.id)(d); pend.delete(d.id) }
  if (d.method === 'Network.requestWillBeSent') {
    const r = d.params.request
    reqs.push(r.method + ' ' + r.url.slice(0, 120) + (r.postData ? ' BODY[' + String(r.postData.text || '').slice(0, 40) + ']' : ''))
  }
  if (d.method === 'Network.responseReceived') {
    const r = d.params.response
    reqs.push('    ↳ ' + r.status + ' [' + (r.mimeType || '') + '] ' + r.url.slice(0, 110))
  }
}
const ev = (expr) => {
  const id = ++seq
  ws.send(JSON.stringify({ id, method: 'Runtime.evaluate', params: { expression: expr, returnByValue: true } }))
  return new Promise((res, rej) => {
    const h = (e) => {
      const m = JSON.parse(e.data)
      if (m.id === id) { ws.removeEventListener('message', h); res(m.result?.result?.value) }
    }
    ws.addEventListener('message', h)
    setTimeout(() => rej(new Error('TO')), 25000)
  })
}
ws.send(JSON.stringify({ id: ++seq, method: 'Network.enable' }))
await new Promise((r) => setTimeout(r, 500))

const pasteAll = `(function(){
  var pocs = [...document.querySelectorAll("iframe")].filter(function(f){ return f.offsetWidth > 100 && f.offsetHeight > 100 })
  if (!pocs.length) return JSON.stringify({ err: "无可见iframe", total: document.querySelectorAll("iframe").length })
  var results = []
  pocs.forEach(function(poc, idx){
    try {
      var fe = poc.contentWindow.document.querySelector("iframe[name=frameEditor]")
      if (!fe) { results.push({ idx: idx, skip: "无frameEditor" }); return }
      var w2 = fe.contentWindow
      var b64 = "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEUlEQVR4nGO4IyeHFTEMLQkAid1GAXiz9RcAAAAASUVORK5CYIIA"
      var bin = atob(b64)
      var u8 = new Uint8Array(bin.length)
      for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
      var dt = new DataTransfer()
      dt.items.add(new File([u8], "clip.png", { type: "image/png" }))
      var ev2 = new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true })
      var target = w2.document.getElementById("editor_sdk") || w2.document.body
      target.dispatchEvent(ev2)
      results.push({ idx: idx, src: (poc.src || "").slice(-34), dispatched: true })
    } catch (e) { results.push({ idx: idx, err: String(e).slice(0, 60) }) }
  })
  return JSON.stringify(results)
})()`

const c1 = await ev(pasteAll)
console.log('粘贴分发:', c1)
await new Promise((r) => setTimeout(r, 12000))
console.log('=== 网络请求（粘贴期间） ===')
const imgReqs = reqs.filter((r) => r.indexOf('/upload') >= 0 || r.indexOf('/work/') >= 0)
console.log('上传相关:', imgReqs.length ? imgReqs.join('\n') : '（无 /upload 与 /work 请求）')
console.log('--- 全部请求尾部 8 条 ---')
console.log(reqs.slice(-8).join('\n') || '（无）')
process.exit(0)
