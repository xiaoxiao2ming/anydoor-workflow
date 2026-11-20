/* OO 保真循环验证（全后门·分步版）：
 * 每一步都是独立短 CDP 调用，等待逻辑在主进程轮询——避开长 awaitPromise 挂死。
 * 链路：宿主页（CDP）→ 同源 iframe 链 → frameEditor 的 Asc.editor.callCommand 直调引擎。
 * 用法：node scripts/oo-fidelity-loop.mjs [轮数=6]
 * 通过：除被删段外，其余段文字+样式与基线完全一致。 */
const N = Number(process.argv[2] || 6)
const NODE_ID = process.argv[3] || 'ooword_1793288276654'

function cdpConnect() {
  return fetch('http://127.0.0.1:9334/json')
    .then((r) => r.json())
    .then((list) => {
      const page = list.find((t) => t.type === 'page' && (t.url || '').includes('10215'))
      if (!page) throw new Error('NO-PAGE（应用未启动或 CDP 未开）')
      return new Promise((resolve, reject) => {
        const ws = new WebSocket(page.webSocketDebuggerUrl)
        ws.onopen = () => resolve(ws)
        ws.onerror = reject
      })
    })
}

function makeEval(ws) {
  let seq = 0
  const pend = new Map()
  ws.onmessage = (ev) => {
    const d = JSON.parse(ev.data)
    if (d.id && pend.has(d.id)) { pend.get(d.id)(d); pend.delete(d.id) }
  }
  return async function evalPage(expr, awaitFlag = false, timeoutMs = 40000) {
    const id = ++seq
    ws.send(JSON.stringify({ id, method: 'Runtime.evaluate', params: { expression: expr, returnByValue: true, awaitPromise: awaitFlag } }))
    return new Promise((res, rej) => {
      const h = (e) => {
        const m = JSON.parse(e.data)
        if (m.id === id) {
          ws.removeEventListener('message', h)
          if (m.result?.exceptionDetails) rej(new Error('PAGE-EXC: ' + JSON.stringify(m.result.exceptionDetails.exception?.description || '').slice(0, 200)))
          else res(m.result?.result?.value)
        }
      }
      ws.addEventListener('message', h)
      setTimeout(() => rej(new Error('EVAL-TIMEOUT')), timeoutMs)
    })
  }
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

/* ---- 分步表达式（全部短平快） ---- */
const EXPR_EDITOR_OPEN = `(function(){ return !!document.querySelector(".ne-head") })()`
const EXPR_OPEN_EDITOR = `(function(){
  var el = document.querySelector(".react-flow")
  if (!el) return "no-canvas"
  var fk = Object.keys(el).filter(function(k){ return k.indexOf("__reactFiber") === 0 })[0]
  var f = el[fk]
  /* 沿祖先链扫函数组件 hooks（构建压缩名每次构建都变，不能依赖组件名） */
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
  if (!app) return "no-app-fiber"
  var h = app.memoizedState, guard = 0
  while (h && guard < 300) {
    var a = h, b = a.next, c = b && b.next, d = c && c.next
    if (a && b && c && d && typeof a.memoizedState === "string" && a.memoizedState.length < 40 &&
        typeof b.memoizedState === "string" && (b.memoizedState === "schema" || b.memoizedState === "data") &&
        typeof c.memoizedState === "number" && c.memoizedState === 320 &&
        typeof d.memoizedState === "string" && d.memoizedState.length < 60) {
      d.queue.dispatch("${NODE_ID}"); return "OPENED"
    }
    h = h.next; guard++
  }
  return "hooks-not-found"
})()`
const EXPR_ENGINE_READY = `(function(){
  var poc = (function(){ var all = document.querySelectorAll(\'iframe\'); var vis = null; for (var i = 0; i < all.length; i++) { if (all[i].offsetWidth > 100 && all[i].offsetHeight > 100) { vis = all[i]; break } } return vis || all[0] })()
  try { return !!(poc && poc.contentWindow && poc.contentWindow.document.querySelector("iframe[name=frameEditor]")) } catch (e) { return false }
})()`
const EXPR_IFRAME_SRC = `(function(){ var f = (function(){ var all = document.querySelectorAll(\'iframe\'); var vis = null; for (var i = 0; i < all.length; i++) { if (all[i].offsetWidth > 100 && all[i].offsetHeight > 100) { vis = all[i]; break } } return vis || all[0] })(); return f ? f.src : "" })()`

const EXPR_DELETE_DIRTY = `(async function(){
  var poc = (function(){ var all = document.querySelectorAll(\'iframe\'); var vis = null; for (var i = 0; i < all.length; i++) { if (all[i].offsetWidth > 100 && all[i].offsetHeight > 100) { vis = all[i]; break } } return vis || all[0] })()
  var fe = poc.contentWindow.document.querySelector("iframe[name=frameEditor]")
  var w2 = fe && fe.contentWindow
  var eng = null
  try { if (w2.Asc && w2.Asc.editor && w2.Asc.editor.callCommand) eng = w2.Asc.editor } catch (e0) {}
  if (!eng && w2.editor && w2.editor.callCommand) eng = w2.editor
  if (!eng) return { err: "no-engine" }
  var done = false
  var ret = await new Promise(function(res){
    var fn = function(){
      try {
        var dc = Api.GetDocument()
        var n = dc.GetElementsCount()
        var tg = -1
        for (var i = n - 1; i >= 0; i--) {
          var e = dc.GetElement(i)
          if (e && e.GetClassType && e.GetClassType() === "paragraph") {
            var tx = ""
            try { tx = String(e.GetText ? e.GetText() : "") } catch (eT) {}
            if (tx.replace(/\\s/g, "") !== "") { tg = i; break }
          }
        }
        if (tg < 0) return { err: "no-nonempty" }
        var e2 = dc.GetElement(tg)
        var rn = e2.GetElementsCount ? e2.GetElementsCount() : 0
        var deleted = 0
        for (var j = rn - 1; j >= 0; j--) { var r = e2.GetElement(j); if (r && typeof r.Delete === "function") { try { r.Delete(); deleted++ } catch (eD) {} } }
        return { target: tg, deleted: deleted }
      } catch (eX) { return { err: String(eX) } }
    }
    try { eng.callCommand(fn, false, false, function(v){ if (!done) { done = true; res(v || { note: "cb-null" }) } }) } catch (e5) {
      try { eng.callCommand(fn, { isCalculate: false, isRecalc: false }, function(v){ if (!done) { done = true; res(v || { note: "cb-null" }) } }) } catch (e6) { res({ err: String(e6) }) }
    }
    setTimeout(function(){ if (!done) { done = true; res({ note: "no-cb" }) } }, 5000)
  })
  var dirty = {}
  try { eng.SetDocumentModified(true); dirty.eng = eng.isDocumentModified() } catch (eS) { dirty.eng = "ERR" }
  try { poc.contentWindow.__setOoDirty(true, "programmatic-delete"); dirty.poc = poc.contentWindow.__ooDocModified } catch (eP) { dirty.poc = "ERR" }
  return { del: ret, dirty: dirty }
})()`

const EXPR_CLICK_SAVE = `(function(){
  var head = document.querySelector(".ne-head")
  /* ★ 保存按钮 title 动态：无修改=「保存：…」/ 有修改=「有未保存的修改 —— 点击保存…」
   *  注意排除「保存副本」（另存到本地，不写节点！）与「打开历史记录」 */
  var saveBtn = [...head.querySelectorAll("button")].find(function(b){
    var t = b.title || ""
    return (t.indexOf("保存：") === 0) || (t.indexOf("有未保存的修改") === 0)
  })
  if (!saveBtn) return "no-btn"
  if (saveBtn.disabled) return "disabled"
  saveBtn.click()
  return "clicked"
})()`

const EXPR_EXTRACT = `(async function(){
  var poc = (function(){ var all = document.querySelectorAll(\'iframe\'); var vis = null; for (var i = 0; i < all.length; i++) { if (all[i].offsetWidth > 100 && all[i].offsetHeight > 100) { vis = all[i]; break } } return vis || all[0] })()
  if (!poc) return { err: "no-poc-iframe" }
  return new Promise(function(res){
    var onMsg = function(e){
      var d = e.data
      if (d && typeof d === "object" && "__ooStyledExtractData" in d) { window.removeEventListener("message", onMsg); res({ paras: d.__ooStyledExtractData, err: d.__ooStyledExtractErr || null }) }
    }
    window.addEventListener("message", onMsg)
    poc.contentWindow.postMessage({ __ooStyledExtractAsk: 1 }, "*")
    setTimeout(function(){ window.removeEventListener("message", onMsg); res({ paras: null, err: "timeout" }) }, 10000)
  })
})()`

const EXPR_OOBYTES64 = `(function(){
  var t = window.__store.getState().tables["${NODE_ID}"]
  var b = t && t.config ? (t.config.ooBytes64 || "") : ""
  /* 简易内容指纹：长度+抽样字符和（zip 等长不同内容也能识别） */
  var h = 0
  for (var i = 0; i < b.length; i += 97) { h = (h + b.charCodeAt(i) * (i + 1)) % 1000000007 }
  return b.length + ":" + h
})()`

const EXPR_EDITOR_CLOSE = `(function(){
  var el = document.querySelector(".react-flow")
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
    if (a && b && c && d && typeof a.memoizedState === "string" && a.memoizedState.length < 40 &&
        typeof b.memoizedState === "string" && (b.memoizedState === "schema" || b.memoizedState === "data") &&
        typeof c.memoizedState === "number" && c.memoizedState === 320 &&
        typeof d.memoizedState === "string" && d.memoizedState.length < 60) {
      d.queue.dispatch(""); return "CLOSED"
    }
    h = h.next; guard++
  }
  return "hooks-not-found"
})()`

function compare(base, cur) {
  const norm = (p) => {
    if (!p || p.t === 'tbl') return { text: '[表格]', style: 'tbl' }
    return {
      text: (p.runs || []).map((r) => r.text || '').join('').replace(/\s+/g, ' ').trim(),
      style: JSON.stringify([(p.runs || []).map((r) => r.font || ''), (p.runs || []).map((r) => r.size || ''), (p.runs || []).map((r) => !!r.bold), p.align || '', String(p.indLeft ?? ''), String(p.style ?? '')]),
    }
  }
  const diffs = []
  const deleted = []
  const n = Math.max(base.length, cur.length)
  for (let i = 0; i < n; i++) {
    const a = norm(base[i]); const b = norm(cur[i])
    if (!a.text && !b.text) continue
    if (a.text && !b.text) { deleted.push({ i, a: a.text.slice(0, 26) }); continue }
    if (a.text !== b.text || a.style !== b.style) diffs.push({ i, textSame: a.text === b.text, a: a.text.slice(0, 40), b: b.text.slice(0, 40), styleSame: a.style === b.style })
  }
  return { diffs, deleted }
}

/* 等引擎就绪（主进程轮询） */
async function waitEngineReady(ev, maxSec = 240) {
  for (let i = 0; i < maxSec; i++) {
    try {
      if (await ev(EXPR_ENGINE_READY)) {
        /* ★ 2026-10-30 切换文档时引擎对象立即存在（旧引擎未销毁）——必须等 POC 状态条
           documentReady 才算新文档加载完成，否则基线提取打到上一个文档（文档2 实测翻车） */
        let st = ''
        try { st = await ev('(function(){ var poc = (function(){ var all = document.querySelectorAll(\'iframe\'); var vis = null; for (var i = 0; i < all.length; i++) { if (all[i].offsetWidth > 100 && all[i].offsetHeight > 100) { vis = all[i]; break } } return vis || all[0] })(); if (!poc) return ""; var s = poc.contentDocument ? poc.contentDocument.getElementById("status") : null; return s ? s.textContent : "" })()', false) } catch { /* 忙 */ }
        if (String(st).indexOf('documentReady') >= 0) { await sleep(4000); return true }
      }
    } catch { /* 页面忙，继续等 */ }
    await sleep(1000)
  }
  return false
}

;(async () => {
  const ws = await cdpConnect()
  const ev = makeEval(ws)

  /* 打开编辑器（若未开）+ 等引擎 */
  if (!(await ev(EXPR_EDITOR_OPEN))) {
    console.log('编辑器未开，后门打开…', await ev(EXPR_OPEN_EDITOR))
  }
  if (!(await waitEngineReady(ev))) { console.log('引擎未就绪'); process.exit(1) }

  let ex = await ev(EXPR_EXTRACT, true)
  if (!ex.paras) { console.log('基线提取失败:', ex.err); process.exit(1) }
  const base = ex.paras
  const baseNonEmpty = base.filter((p) => p && p.t !== 'tbl' && (p.runs || []).map((r) => r.text || '').join('').replace(/\s/g, '') !== '').length
  let lastNonEmptyCount = baseNonEmpty
  console.log(`基线：${base.length} 段（非空 ${baseNonEmpty}）——第一次打开的样式`)

  for (let round = 1; round <= N; round++) {
    console.log(`\n===== 第 ${round} 轮 =====`)

    /* 每轮确保编辑器开着（PASS 后的强制重开可能有竞态） */
    if (!(await ev(EXPR_EDITOR_OPEN))) {
      console.log('编辑器未开，后门重开…')
      const o = await ev(EXPR_OPEN_EDITOR)
      if (o !== 'OPENED') { console.log('打开失败:', o); break }
      if (!(await waitEngineReady(ev, 60))) { console.log('引擎未就绪'); break }
      await sleep(4000)
    }

    /* 1) 删除+置脏（验证生效，未生效重试最多 3 次） */
    let d1 = null
    let delOk = false
    for (let attempt = 1; attempt <= 3 && !delOk; attempt++) {
      try { d1 = await ev(EXPR_DELETE_DIRTY, true) } catch (e) { d1 = { err: String(e).slice(0, 80) } }
      /* 验证：提取模型，与上轮快照对比确实少了一段非空文字 */
      try {
        const probe = await ev(EXPR_EXTRACT, true)
        if (probe.paras) {
          const nonEmpty = probe.paras.filter((p) => p && p.t !== 'tbl' && (p.runs || []).map((r) => r.text || '').join('').replace(/\s/g, '') !== '').length
          if (typeof lastNonEmptyCount === 'number' ? nonEmpty < lastNonEmptyCount : attempt > 1 ? false : nonEmpty < 33) { delOk = true; lastNonEmptyCount = nonEmpty }
          else lastNonEmptyCount = nonEmpty
        }
      } catch { /* probe fail → retry delete */ }
      if (!delOk) { console.log(`  删除未生效（第 ${attempt} 次尝试），重试…`); await sleep(4000) }
    }
    console.log('删除+置脏:', JSON.stringify(d1), delOk ? '✓已生效' : '✗未验证')

    /* 2) 点保存（立即返回） */
    let ck
    try { ck = await ev(EXPR_CLICK_SAVE) } catch (e) { console.log('点击失败:', String(e).slice(0, 90)); break }
    console.log('保存点击:', ck)
    if (ck !== 'clicked') break

    /* 2.5) ★ 等保存链把新字节写进节点（轮询 ooBytes64 长度变化，60s 上限）
     *     不等就重开会读到旧字节——「假 PASS/内容回退」的元凶 */
    const b64Before = await ev(EXPR_OOBYTES64)
    let b64Written = false
    for (let i = 0; i < 60; i++) {
      await sleep(1000)
      try {
        const b64Now = await ev(EXPR_OOBYTES64)
        if (b64Now !== b64Before) { b64Written = true; break }
      } catch { /* 页面忙 */ }
    }
    console.log('字节落库:', b64Written ? `✓（${b64Before} → ${await ev(EXPR_OOBYTES64)}）` : `⚠ 60s 未变化（${b64Before}）——可能内容无实际变化`)

    /* 3) 主进程轮询重开（src 变化），90s 上限 */
    const before = await ev(EXPR_IFRAME_SRC)
    let reopened = false
    for (let i = 0; i < 90; i++) {
      await sleep(1000)
      try {
        const now = await ev(EXPR_IFRAME_SRC)
        if (now && now !== before) { reopened = true; break }
      } catch { /* context 短暂不可用 */ }
    }
    console.log('自动重开:', reopened ? '是' : '否（将强制重开）')

    /* 4) 保证从「已保存字节」读：未自动重开则后门强制重开 */
    if (!reopened) {
      try { await ev(EXPR_EDITOR_CLOSE) } catch { /* ignore */ }
      await sleep(2000)
      try { await ev(EXPR_OPEN_EDITOR) } catch (e) { console.log('重开失败:', String(e).slice(0, 90)); break }
      if (!(await waitEngineReady(ev, 60))) { console.log('重开后引擎未就绪'); break }
    } else {
      await sleep(8000)
    }

    /* 5) 提取 + 对比（重开后 callCommand 需文档完全 ready——轮询重试） */
    ex = { paras: null }
    for (let t = 0; t < 10; t++) {
      await sleep(5000)
      try { ex = await ev(EXPR_EXTRACT, true) } catch { continue }
      if (ex.paras) break
    }
    if (!ex.paras) { console.log('提取失败（重试10次）:', ex.err); break }
    const cur = ex.paras
    const { diffs, deleted } = compare(base, cur)
    console.log(`已删段 ${deleted.length} 处: ${deleted.map((x) => '[' + x.i + ']').join('') || '无'}`)
    if (diffs.length === 0) {
      console.log(`✅ 第 ${round} 轮 PASS：除被删段外，${cur.length} 段文字+样式与基线完全一致`)
      console.log('\n===== 保真闭环达成 =====')
      process.exit(0)
    }
    console.log(`❌ 第 ${round} 轮差异 ${diffs.length} 处：`)
    diffs.slice(0, 6).forEach((d) => console.log(`  [${d.i}] ${d.textSame ? '(文字同,样式异)' : '文字异'} 基线="${d.a}" 本次="${d.b}"${d.styleSame ? '' : ' [样式也异]'}`))
  }
  console.log('\n===== 未 PASS（差异如上）=====')
  process.exit(1)
})().catch((e) => { console.error('FATAL', String(e).slice(0, 200)); process.exit(2) })
