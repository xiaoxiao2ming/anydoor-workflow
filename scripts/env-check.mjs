/* 引擎环境快检：Api 存在性 + callCommand 真实状态（不经闭包） */
const list = await (await fetch('http://127.0.0.1:9334/json')).json()
const page = list.find((t) => t.type === 'page' && (t.url || '').includes('10215'))
if (!page) { console.error('NO-PAGE'); process.exit(1) }
const ws = new WebSocket(page.webSocketDebuggerUrl)
await new Promise((r, j) => { ws.onopen = r; ws.onerror = j })
ws.send(JSON.stringify({
  id: 1,
  method: 'Runtime.evaluate',
  params: {
    expression: `(function(){
  try {
    var poc = document.querySelector("iframe")
    var w2 = poc.contentWindow.document.querySelector("iframe[name=frameEditor]").contentWindow
    var eng = w2.Asc && w2.Asc.editor
    return JSON.stringify({
      Api: typeof w2.Api,
      CreateImage: w2.Api && w2.Api.CreateImage ? "has" : "no",
      engCC: eng && eng.callCommand ? "has" : "no",
      isMode: eng && eng.isMode ? "yes" : "no",
      DocEditor: typeof w2.DocEditor,
      title: w2.document.title
    })
  } catch (e) { return "err:" + String(e).slice(0, 80) }
})()`,
    returnByValue: true,
  },
}))
ws.onmessage = (ev) => {
  const d = JSON.parse(ev.data)
  if (d.id === 1) { console.log('引擎环境:', d.result?.result?.value); process.exit(0) }
}
setTimeout(() => { console.log('TO'); process.exit(1) }, 20000)
