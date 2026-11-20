/**
 * OO 远程桥（2026-10-17）—— 让网页版也能在线编辑工作簿/文档
 *
 * 背景：桌面端的 OnlyOffice 是 Windows 二进制（vendors/oo-server 的 x2t.exe + 本机 DS:3798），
 * Linux 容器跑不了，网页版因此一直「浏览器版不支持在线编辑」。
 * 现在：群晖上跑官方 Linux 版 DocumentServer（Docker，JWT 关闭），本模块把桌面端那套
 * 「宿主组件 ↔ /poc ↔ 本机 DS」协议原样桥到远程 DS 上——宿主 OnlyOfficeNodeEditor 一行不改。
 *
 * 协议对齐（宿主组件实际用到的子集）：
 *   POST /office-offline/open   {name, dataBase64}      → 工作区落盘，返回 {ok}
 *   GET  /poc?src=<name>                              → 返回远程版 poc 页（内部加载 DS api.js）
 *   GET  /poc-meta?src=<name>                         → {fileUrl,key,callbackUrl,dsPublic,...}
 *   POST /office-offline/save  {name}                 → 取 DS 回执产物（.saved 文件）{ok,dataBase64}
 *   POST /doc/command          {c:'forcesave',key}    → 转发 DS 命令服务（/command）
 *   POST /doc/callback?ws&tok  (DS→本服务)            → status 2/6 下载产物存 .saved
 *   GET  /office-src/:tok/:name (DS→本服务)            → DS 拉源文件（token 门禁，公开路径）
 *
 * 保存链：编辑器 asc_Save → DS 收变更 → 宿主 forcesave(/doc/command) → DS 回调 /doc/callback(6)
 *        → 下载 docx/xlsx 存 .saved → 宿主轮询 /office-offline/save 取走 → 写回节点。
 *
 * 安全边界：
 *   - /poc、/poc-meta、/office-offline/* 在口令门禁后面（不进 auth 白名单）。
 *   - /office-src 与 /doc/callback 是 DS 回访通道，DS 没有登录态，靠 URL 里 32 位随机 token 门禁。
 *   - JWT 本阶段关闭（局域网内网可接受）；上公网必须开 JWT 并配 APP_OO_JWT_SECRET。
 */

import { createHash, randomBytes } from 'node:crypto'
import { readFile, writeFile, mkdir, stat } from 'node:fs/promises'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const BLANK_XLSX_B64 = existsSync(join(__dirname, '_blank_xlsx.b64'))
  ? readFileSync(join(__dirname, '_blank_xlsx.b64'), 'utf8').trim()
  : ''
const BLANK_DOCX_B64 = existsSync(join(__dirname, '_blank_docx.b64'))
  ? readFileSync(join(__dirname, '_blank_docx.b64'), 'utf8').trim()
  : ''

const REMOTE_MIME = {
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
}

export function isOoRemote () {
  const u = process.env.APP_OO_DS_URL
  return !!u && !!String(u).trim()
}

const dsInternal = () => String(process.env.APP_OO_DS_URL || '').replace(/\/$/, '')

const safeName = (n) => {
  const s = String(n || '')
  if (!/^[A-Za-z0-9._-]{1,120}$/.test(s) || s.includes('..')) return ''
  return s
}

const extOf = (n) => {
  const m = /\.(xlsx|docx)$/i.exec(n)
  return m ? m[1].toLowerCase() : ''
}

/** xlsx → cell（表格），docx → word（文档）——DS 的 documentType 术语 */
const docTypeOf = (ext) => (ext === 'xlsx' ? 'cell' : 'word')

const b64 = (buf) => buf.toString('base64')

function wsDir (OFFICE_DIR) {
  return join(OFFICE_DIR, 'remote-ws')
}

function metaPath (dir, name) {
  return join(dir, name + '.json')
}

async function readMeta (dir, name) {
  try {
    return JSON.parse(await readFile(metaPath(dir, name), 'utf8'))
  } catch {
    return null
  }
}

const keyOf = (name, gen) => 'kbs-' + createHash('sha1').update(`${name}:${gen}`).digest('hex').slice(0, 20)

const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]))

function sendJson (res, code, obj) {
  if (res.writableEnded) return
  res.setHeader('Content-Type', 'application/json; charset=utf-8')
  res.setHeader('Cache-Control', 'no-store')
  res.statusCode = code
  res.end(JSON.stringify(obj))
}

/** 读请求体（上限 32MB —— xlsx/docx 工作簿不会更大） */
function readBodyRaw (req, limit = 32 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    const chunks = []
    let n = 0
    req.on('data', (c) => {
      n += c.length
      if (n > limit) { reject(new Error('body too large')); req.destroy(); return }
      chunks.push(c)
    })
    req.on('end', () => resolve(Buffer.concat(chunks)))
    req.on('error', reject)
  })
}

/**
 * 远程版 poc 页。与桌面 /poc 的宿主协议子集对齐：
 *   - window.__setOoDirty(v,why)（宿主 ackSaved 同源直调）
 *   - __ooDirty / __ooDirtyWhy 推送；__ooIsModifiedAsk → {__ooIsModified}
 *   - __ooSaveReq → 深帧 asc_Save（变更推给 DS），产物由宿主 forcesave+轮询取
 *   - window.__docKey（宿主 saveBack 读取，发 forcesave 用）
 *   - __ooStatus / __ooAudit 汇报
 *   - __ooSavedAck → 清脏
 */
function pocPage () {
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="robots" content="noindex,nofollow">
<title>OO Remote</title>
<style>
  html,body{margin:0;padding:0;height:100%;overflow:hidden;background:#e8ecf0}
  #editor{position:absolute;inset:0}
  #tip{position:absolute;left:50%;top:42%;transform:translateX(-50%);font:13px/1.8 -apple-system,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;color:#5a6572;text-align:center;white-space:pre-line}
</style>
</head>
<body>
<div id="editor"></div>
<div id="tip">正在连接文档服务…</div>
<script>
(function () {
  var q = new URLSearchParams(location.search)
  var src = q.get('src') || ''
  var dirty = false
  var editor = null
  function tip (s) { var t = document.getElementById('tip'); if (t) t.textContent = s }
  /* ★ 新功能引导气泡默认关（2026-10-18，与桌面 oo-core.mjs 同款）：
   *   TooltipManager 按 localStorage[tip.name] 判断是否已看过 → 预写全部键。
   *   编辑器与 poc 页同源时生效；远程 DS 跨域时无害（该域浏览器里点一次「知道了」后自持久化）。 */
  try { ['sse-help-tip-cellFormat', 'sse-help-tip-chart-tab', 'sse-help-tip-darkDocument', 'sse-help-tip-solver',
         'de-help-tip-chart-tab', 'de-help-tip-fill-status', 'de-help-tip-header-footer-tab', 'de-help-tip-multipage-view-statusbar', 'de-help-tip-pasteOptions', 'de-help-tip-signature',
         'pe-help-tip-chart-tab', 'pe-help-tip-gif-payback', 'pe-help-tip-pasteOptions', 'pe-help-tip-slideTheme', 'pe-help-tip-slideTransitions',
        ].forEach(function (k) { try { localStorage.setItem(k, '1') } catch (eKT) {} }) } catch (e) {}
  function status (s) { try { parent.postMessage({ __ooStatus: '[remote] ' + s }, '*') } catch (e) {} }
  function audit (kind, action, detail) {
    try { parent.postMessage({ __ooAudit: { kind: kind, action: action, detail: detail || '' } }, '*') } catch (e) {}
  }
  window.__setOoDirty = function (v, why) {
    dirty = !!v
    try { parent.postMessage({ __ooDirty: dirty, __ooDirtyWhy: why || '' }, '*') } catch (e) {}
  }
  /* ── 高亮行列（2026-10-18，与桌面 oo-core.mjs 同款）：asc_getActiveCellCoord 的
   *    稳定公共 getter asc_getX/Y/Width/Height 给出选中格像素矩形，叠两条色带。 ── */
  var __hiRC = { on: false, timer: null, reg: false }
  /* 横竖同色：行带拆成「交叉格左侧 + 右侧」两段，避免与列带叠出更深的交叉点 */
  var __HI_BG = 'rgba(120,190,255,.26)'
  function __hiFrame () {
    var fE = document.querySelector('iframe[name="frameEditor"]') || document.querySelector('iframe')
    if (!fE || !fE.contentDocument) return null
    var cv = fE.contentDocument.getElementById('ws-canvas')
    if (!cv) return null
    return { frame: fE, doc: fE.contentDocument, cv: cv }
  }
  function __hiEnsureEl (t) {
    var doc = t.doc
    var host = doc.getElementById('__cb_rowcol_hi')
    if (host && host.children.length !== 3) {
      try { host.parentNode.removeChild(host) } catch (eHO) { /* 忽略 */ }
      host = null
    }
    if (!host) {
      host = doc.createElement('div')
      host.id = '__cb_rowcol_hi'
      host.style.cssText = 'position:absolute;pointer-events:none;z-index:9997;overflow:hidden;'
      t.cv.parentElement.appendChild(host)
      var rowA = doc.createElement('div'); rowA.id = '__cb_hi_row_a'
      rowA.style.cssText = 'position:absolute;top:0;height:0;background:' + __HI_BG + ';display:none;'
      var rowB = doc.createElement('div'); rowB.id = '__cb_hi_row_b'
      rowB.style.cssText = 'position:absolute;top:0;height:0;background:' + __HI_BG + ';display:none;'
      var col = doc.createElement('div'); col.id = '__cb_hi_col'
      col.style.cssText = 'position:absolute;top:0;height:100%;left:0;width:0;background:' + __HI_BG + ';display:none;'
      host.appendChild(rowA); host.appendChild(rowB); host.appendChild(col)
    }
    host.style.left = t.cv.offsetLeft + 'px'
    host.style.top = t.cv.offsetTop + 'px'
    host.style.width = t.cv.offsetWidth + 'px'
    host.style.height = t.cv.offsetHeight + 'px'
    return host
  }
  function __hiPaint () {
    if (!__hiRC.on) return
    var t = __hiFrame()
    if (!t) return
    var host = __hiEnsureEl(t)
    var w2 = t.frame.contentWindow
    var api = w2.editor || (w2.Asc && w2.Asc.editor)
    if (!api || typeof api.asc_getActiveCellCoord !== 'function') return
    var c = null
    try { c = api.asc_getActiveCellCoord() } catch (eH0) { return }
    if (!c) return
    var gv = function (m) { try { if (c && typeof c[m] === 'function') return c[m]() } catch (eH1) { /* 忽略 */ } return null }
    var x = gv('asc_getX'), y = gv('asc_getY'), w = gv('asc_getWidth'), h = gv('asc_getHeight')
    if (x == null || y == null || w == null || h == null) return
    /* 选区类型（Asc.c_oAscSelectionType 实测：1=RangeCells 2=RangeCol 3=RangeRow）：
     * 点行号/列标选中整行/整列时，只亮对应单条带，不出十字 */
    var st = null
    try {
      if (typeof api.asc_getCellInfo === 'function') {
        var ci = api.asc_getCellInfo()
        if (ci && typeof ci.asc_getSelectionType === 'function') st = ci.asc_getSelectionType()
      }
    } catch (eST) { /* 拿不到就按普通单元格处理 */ }
    var onlyRow = st === 3
    var onlyCol = st === 2
    var hostH = host.clientHeight || t.cv.offsetHeight
    var rowA = host.children[0], rowB = host.children[1], col = host.children[2]
    if (!rowA || !rowB || !col) return
    var hostW = host.clientWidth || t.cv.offsetWidth
    if (onlyRow) {
      rowA.style.left = x + 'px'; rowA.style.width = Math.max(0, hostW - x) + 'px'
      rowA.style.top = y + 'px'; rowA.style.height = h + 'px'; rowA.style.display = 'block'
      rowB.style.display = 'none'; col.style.display = 'none'
      return
    }
    if (onlyCol) {
      col.style.left = x + 'px'; col.style.width = w + 'px'
      col.style.top = y + 'px'; col.style.height = Math.max(0, hostH - y) + 'px'; col.style.display = 'block'
      rowA.style.display = 'none'; rowB.style.display = 'none'
      return
    }
    /* 行带：左段 [0, x) + 右段 [x+w, hostW)，交叉格只由列带上色 → 横竖同色、无叠色 */
    var rightStart = x + w
    var rightW = hostW - rightStart
    rowA.style.left = '0px'; rowA.style.width = Math.max(0, x) + 'px'
    rowA.style.top = y + 'px'; rowA.style.height = h + 'px'; rowA.style.display = 'block'
    rowB.style.left = rightStart + 'px'; rowB.style.width = Math.max(0, rightW) + 'px'
    rowB.style.top = y + 'px'; rowB.style.height = h + 'px'; rowB.style.display = 'block'
    col.style.left = x + 'px'; col.style.width = w + 'px'
    col.style.top = '0px'; col.style.height = '100%'; col.style.display = 'block'
  }
  function __hiTeardown () {
    if (__hiRC.timer) { clearInterval(__hiRC.timer); __hiRC.timer = null }
    var t = __hiFrame()
    var el = t && t.doc.getElementById('__cb_rowcol_hi')
    if (el && el.parentNode) el.parentNode.removeChild(el)
  }
  function __hiSet (on) {
    __hiRC.on = on === true
    if (!__hiRC.on) { __hiTeardown(); return }
    __hiPaint()
    if (!__hiRC.timer) __hiRC.timer = setInterval(__hiPaint, 250)
    if (!__hiRC.reg) {
      __hiRC.reg = true
      try {
        var t = __hiFrame()
        var w2 = t && t.frame.contentWindow
        var api = w2 && (w2.editor || (w2.Asc && w2.Asc.editor))
        if (api && typeof api.asc_registerCallback === 'function') {
          api.asc_registerCallback('asc_onSelectionChanged', function () { try { __hiPaint() } catch (eH2) {} })
        }
      } catch (eH3) { /* 轮询兜底 */ }
    }
  }
  /* ── 引用格覆盖层（2026-10-18 着色开关改版：颜色不写文件，纯视觉层）──
   * 宿主 postMessage {__ooRefMarks: {on, marks:[{sheet,ref,color}], geom}}。
   *
   * ★ 2026-10-18 用户反馈「亮显颜色容易偏移 / 为什么不按单元格实际位置」后的重写。
   *   色块只能浮在上面 —— OO 没有「不改文件的逐格上色」API，写样式就必须重载文档
   *   （用户已明确否决重载）。但**位置不再靠算**，直接问引擎要单元格的真实像素矩形：
   *     ① 名称框跳转是同步的（实测连跳 12 次仅 19ms），跳过去读 asc_getActiveCellCoord
   *        就是该格的精确 x/y/w/h，跳完立刻跳回原选区，整个测量在一个 JS 任务里跑完，
   *        浏览器根本没机会重绘 ⇒ 用户看不到任何闪动，选区/滚动也原样还原（实测 ≤1px）。
   *     ② 量好的矩形随锚点（选区活动格）位移平移：滚动时锚点坐标同量变化，减法即精确跟随。
   *     ③ 用户在表里换选区时重新量一次（下一帧执行，避免重入），始终保持实测。
   *   geom 只用于两件事：③ 估算「还没量到的格」位置做兜底，以及合并单元格尺寸。
   *   几何口径：defColCh/defRowPt（字符 / pt）+ colCh/rowPt（自定义值，隐藏记 0）+ merges；
   *   kCol/kRow 用活动格实测像素现场标定，隐藏列按 0 宽累计，行标头宽/列标头高在相对
   *   定位中天然消掉（旧实现 pt×4/3 每行差 1px、隐藏列照样算宽度 ⇒ 整片位移）。
   */
  var __rm = {
    on: false, marks: [], geom: null, raf: 0, sig: '',
    kCol: 7, kRow: 1.3333,
    meas: null, needMeas: false, measuring: false,
  }
  function __rmArgbCss (argb) {
    var s = String(argb || '')
    if (s.charAt(0) === '#') s = s.slice(1)
    while (s.length < 8) s = 'F' + s
    var r = parseInt(s.slice(2, 4), 16), g = parseInt(s.slice(4, 6), 16), b = parseInt(s.slice(6, 8), 16)
    return 'rgba(' + r + ',' + g + ',' + b + ',.35)'
  }
  /* 列宽（字符）/ 行高（pt）：自定义优先，没有就用表默认值 */
  function __rmColCh (c) {
    var g = __rm.geom || {}
    return (g.colCh && g.colCh[String(c)] != null) ? g.colCh[String(c)] : (g.defColCh != null ? g.defColCh : 8.43)
  }
  function __rmRowPt (r) {
    var g = __rm.geom || {}
    return (g.rowPt && g.rowPt[String(r)] != null) ? g.rowPt[String(r)] : (g.defRowPt != null ? g.defRowPt : 15)
  }
  /* 字符 → 像素：px = kCol × 字符数（线性过原点；OO 实测 9 字符 → 63px、75.6 → 529px） */
  function __rmColPx (c) {
    var ch = __rmColCh(c)
    if (!(ch > 0)) return 0
    return Math.round(__rm.kCol * ch)
  }
  function __rmRowPx (r) {
    var pt = __rmRowPt(r)
    if (!(pt > 0)) return 0
    return Math.round(__rm.kRow * pt)
  }
  function __rmPrefixX (c) { var s = 0; for (var i = 0; i < c; i++) s += __rmColPx(i); return s }
  function __rmPrefixY (r) { var s = 0; for (var i = 0; i < r; i++) s += __rmRowPx(i); return s }
  /* 命中合并区就返回整块，色块铺满合并单元格 */
  function __rmSpan (r, c) {
    var ms = (__rm.geom && __rm.geom.merges) || []
    for (var i = 0; i < ms.length; i++) {
      var m = ms[i]
      if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m
    }
    return { r1: r, c1: c, r2: r, c2: c }
  }
  /* 用活动格实测像素反标定比例：把渲染器真实换算学下来，而不是猜 */
  function __rmCalib (ac, ar, w, h) {
    if (!(w > 4) && !(h > 4)) return
    var sp = __rmSpan(ar, ac)
    var chs = 0
    for (var i = sp.c1; i <= sp.c2; i++) chs += __rmColCh(i)
    var pts = 0
    for (var j = sp.r1; j <= sp.r2; j++) pts += __rmRowPt(j)
    if (w > 4 && chs > 0.5) { var kc = w / chs; if (kc > 2 && kc < 30) __rm.kCol = kc }
    if (h > 4 && pts > 0.5) { var kr = h / pts; if (kr > 0.6 && kr < 4) __rm.kRow = kr }
  }
  /* 名称框跳转（同步生效）+ 读当前格真实矩形 */
  function __rmJump (nb, ref) {
    try {
      nb.focus()
      nb.value = ref
      nb.dispatchEvent(new Event('input', { bubbles: true }))
      nb.dispatchEvent(new Event('change', { bubbles: true }))
      var types = ['keydown', 'keypress', 'keyup']
      for (var i = 0; i < types.length; i++) {
        nb.dispatchEvent(new KeyboardEvent(types[i], {
          key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true, cancelable: true,
        }))
      }
      return true
    } catch (eJ0) { return false }
  }
  function __rmReadRect (c) {
    if (!c) return null
    var g = function (m) { try { return (c && typeof c[m] === 'function') ? c[m]() : null } catch (e) { return null } }
    var x = g('asc_getX'), y = g('asc_getY'), w = g('asc_getWidth'), h = g('asc_getHeight')
    if (x == null || y == null) return null
    return { x: x, y: y, w: w, h: h }
  }
  /* 逐个引用格实测（跳过去 → 读矩形 → 最后跳回原选区），全程同步、屏幕不重绘 */
  function __rmMeasure (t, api, nb, doc) {
    if (__rm.measuring || !nb || !api || typeof api.asc_getActiveCellCoord !== 'function') return
    /* 注：OO 的键盘捕获 textarea（#area_id）永远持有焦点，不能用 activeElement 判编辑态；
     * 而本函数只在「点开关 / 锚点格变化」时触发，这两种操作本身已经把编辑态提交掉了。 */
    __rm.measuring = true
    /* 记住测量前的焦点：下面要用名称框连续跳转，收尾必须还回去（见 finally） */
    var prevAct = null
    try { prevAct = doc.activeElement } catch (ePA0) { prevAct = null }
    /* ★ 测量自身连续跳转 N 次，每次都触发选区回调。旧实现把「测量引发的选区变化」
     *   当成用户换选区 ⇒ 又置 needMeas ⇒ 再量 ⇒ 无限循环（实测静置 4 秒量 14 次，
     *   屏幕上活动格/选区框一直在重绘 = 用户说的「闪来闪去」）。开静默窗，期间一律忽略。 */
    try { __rm.quietUntil = Date.now() + 1500 } catch (eQq0) { __rm.quietUntil = 0 }
    var saved = nb.value ? String(nb.value) : ''
    var curSheet0 = ''
    try {
      var tl0 = doc.querySelector('#status-sheets-bar li.list-item.active')
      curSheet0 = tl0 ? String(tl0.textContent || '').trim() : ''
    } catch (eQ0) { curSheet0 = '' }
    try {
      var rects = {}
      for (var i = 0; i < __rm.marks.length; i++) {
        var mk0 = __rm.marks[i]
        /* 别的表上的引用格：当前表没有它的几何，跳过（否则会拿同地址误量） */
        if (curSheet0 && mk0.sheet && String(mk0.sheet) !== curSheet0) continue
        var ref = String(mk0.ref || '')
        if (!ref) continue
        if (!__rmJump(nb, ref)) continue
        var r = __rmReadRect(api.asc_getActiveCellCoord())
        if (r) rects[ref.replace(/[$]/g, '').toUpperCase()] = r
      }
      /* 名称框为空（刚打开文档还没选过格）时收尾不跳回 ⇒ 选区停在最后一个引用格，
       * 此时快照必须记「实际所在格」，否则 paint 里 useMeas 永远 false、退回模型估位。 */
      if (!saved) { try { saved = nb.value ? String(nb.value) : '' } catch (eSv0) { /* 保持空 */ } }
      if (saved) __rmJump(nb, saved)
      var anchor = __rmReadRect(api.asc_getActiveCellCoord())
      __rm.meas = anchor
        ? { ref: saved, refN: String(saved || '').replace(/[$]/g, '').toUpperCase(),
            ax: anchor.x, ay: anchor.y, aw: anchor.w, ah: anchor.h, rects: rects }
        : null
      /* 记下测量时的画布尺寸：resize 之后快照即失效 */
      __rm.cvW0 = t.cv.offsetWidth
      __rm.cvH0 = t.cv.offsetHeight
      /* 诊断口：编辑器窗口里 __ooRm 可看到实测快照（支持排查用） */
      try { var win0 = t.frame.contentWindow; if (win0) win0.__ooRm = __rm } catch (eD3) { /* 忽略 */ }
    } catch (eM0) {
      __rm.meas = null
    } finally {
      __rm.measuring = false
      /* ★ 把键盘焦点还给表格（#area_id 是 OO 的键盘捕获层）。旧实现跳完把焦点留在名称框，
       *   Ctrl+= / Ctrl+- / Ctrl+0 缩放、方向键、直接打字全进了名称框 —— 用户实测「不能缩放了」。 */
      try {
        if (prevAct && prevAct !== nb && prevAct !== doc.body && typeof prevAct.focus === 'function') prevAct.focus()
        else { var taK = doc.getElementById('area_id'); if (taK && typeof taK.focus === 'function') taK.focus() }
      } catch (ePA1) { /* 元素可能已被移除，忽略 */ }
      /* 收尾：静默窗再延 250ms（引擎的选区回调可能异步派发），并让下一帧按新快照重画 */
      try { __rm.quietUntil = Date.now() + 250 } catch (eQq1) { __rm.quietUntil = 0 }
      __rm.holdUntil = 0
      __rm.sig = ''
    }
  }
  function __rmPaint () {
    if (!__rm.on || !__rm.marks.length) return
    var t = __hiFrame()
    if (!t) return
    var doc = t.doc
    var w2 = t.frame.contentWindow
    var api = w2.editor || (w2.Asc && w2.Asc.editor)
    if (!api || typeof api.asc_getActiveCellCoord !== 'function') return
    var nb = doc.getElementById('ce-cell-name')
    if (!nb) return
    /* 编辑时只藏「正在编辑的那一格」的色块（否则盖住输入内容），其余引用格照常显色。
     * 信号用 api.isCellEdited —— web-apps 主控在 editStart 置 true、editEnd 置 false（权威）。 */
    var inEdit = api.isCellEdited === true
    if (inEdit !== (__rm.editHide === true)) {
      __rm.editHide = inEdit
      __rm.sig = ''          /* 状态翻转强制下轮重画 */
      if (!inEdit) { __rm.needMeas = true; __rm.recheckLeft = 1 }   /* 编辑可能改了行高（自动撑高），退出后重量 */
    }
    /* 编辑中不测量：名称框合成跳转会把没输完的内容提交掉 */
    if (__rm.needMeas && !inEdit) {
      var nowMs = 0
      try { nowMs = Date.now() } catch (eN0) { nowMs = 0 }
      /* 等选区/缩放稳定 120ms 再量（Ctrl+Home 后滚动未落定时量会差 1px）；同类操作 250ms 节流 */
      if (nowMs - (__rm.selAt || 0) >= 180 && nowMs - (__rm.measAt || 0) >= 250 &&
        nowMs >= (__rm.needFrom || 0)) {
        /* 活动格不在视口内时别量：跳转测量的收尾跳回会把用户的滚动位置拽回去 */
        var aNow = __rmReadRect(api.asc_getActiveCellCoord())
        var cw0 = t.cv.offsetWidth, ch0 = t.cv.offsetHeight
        var aVis = !!(aNow && aNow.w > 0 && aNow.h > 0 &&
          aNow.x < cw0 && aNow.y < ch0 && (aNow.x + aNow.w) > 0 && (aNow.y + aNow.h) > 0)
        if (aVis) {
          __rm.needMeas = false
          __rm.measAt = nowMs
          __rmMeasure(t, api, nb, doc)
          /* 引擎的几何有 1~2 帧的落定过程：量完在 300ms 后再补量一次，把残差彻底归零 */
          if (__rm.recheckLeft > 0) {
            __rm.recheckLeft--
            __rm.needMeas = true
            __rm.needFrom = nowMs + 300
          }
        }
      }
    }
    var addr = nb.value ? String(nb.value) : ''
    /* ★ 正则一律用 [$] 和 [0-9]：本文件整体包在 /poc 的 HTML 模板字面量里下发，
     *   模板会把「反斜杠d」吞成 d、「反斜杠美元」吞成 $（1987/2769 行历史惨案同款）。 */
    var am = /([$]?[A-Z]{1,3}[$]?[0-9]+)/.exec(addr.replace(/^[^:]*!/, ''))
    if (!am) return
    var parts = am[1].replace(/[$]/g, '').toUpperCase().match(/^([A-Z]{1,3})([0-9]+)$/)
    if (!parts) return
    var ac0 = 0
    var sL = parts[1]
    for (var i = 0; i < sL.length; i++) ac0 = ac0 * 26 + (sL.charCodeAt(i) - 64)
    ac0 -= 1
    var ar0 = Number(parts[2]) - 1
    var c = null
    try { c = api.asc_getActiveCellCoord() } catch (eR0) { return }
    var anchor = __rmReadRect(c)
    if (!anchor) return
    /* resize / 切页瞬间引擎会返回 0 宽或离谱坐标的瞬时值：这一帧宁可不画，也别画飞 */
    if (!(anchor.w > 0) || !(anchor.h > 0)) return
    if (Math.abs(anchor.x) > 200000 || Math.abs(anchor.y) > 200000) return
    var x = anchor.x, y = anchor.y
    __rmCalib(ac0, ar0, anchor.w, anchor.h)
    var aS = __rmSpan(ar0, ac0)
    var baseX = x - __rmPrefixX(aS.c1)
    var baseY = y - __rmPrefixY(aS.r1)
    /* 实测快照：同一个选区地址才能按锚点位移直接平移 */
    var mKey = addr.replace(/[$]/g, '').toUpperCase()
    /* 正在编辑的格：其色块要藏起来（编辑框底下的浮层会盖住输入文字） */
    var editRef = inEdit ? mKey : ''
    /* 当前工作表名：别的表上的引用格不画（标签可见即可读到） */
    var curSheet = ''
    try {
      var tl1 = doc.querySelector('#status-sheets-bar li.list-item.active')
      curSheet = tl1 ? String(tl1.textContent || '').trim() : ''
    } catch (eQ1) { curSheet = '' }
    var useMeas = !!(__rm.meas && __rm.meas.ref &&
      String(__rm.meas.ref).replace(/[$]/g, '').toUpperCase() === mKey)
    /* 快照过期自检：锚点尺寸变了（Ctrl+滚轮 / 缩放按钮 = 缩放）或画布尺寸变了（窗口 resize、
     * 面板开合）⇒ 按旧快照画必然整体漂，立即安排重量；这一帧先用旧值顶着，下一帧归零。 */
    if (useMeas && __rm.meas.aw > 0 &&
      (Math.abs(anchor.w - __rm.meas.aw) > 0.5 || Math.abs(anchor.h - __rm.meas.ah) > 0.5)) {
      __rm.needMeas = true
      __rm.recheckLeft = 2
    }
    if (__rm.cvW0 && (Math.abs(t.cv.offsetWidth - __rm.cvW0) > 1 || Math.abs(t.cv.offsetHeight - __rm.cvH0) > 1)) {
      __rm.needMeas = true
      __rm.recheckLeft = 2
    }
    var dx = useMeas ? x - __rm.meas.ax : 0
    var dy = useMeas ? y - __rm.meas.ay : 0
    /* 引擎的整数化在滚动边界会来回抖 1px（A1 有时 18 有时 19）：≤1px 视为不抖，避免整层抖 */
    if (Math.abs(dx) <= 1) dx = 0
    if (Math.abs(dy) <= 1) dy = 0
    var cvW = t.cv.offsetWidth, cvH = t.cv.offsetHeight
    var mVer = useMeas ? Object.keys(__rm.meas.rects).length : -1
    var sig = x + ',' + y + ',' + ac0 + ',' + ar0 + ',' + anchor.w + ',' + anchor.h + ',' + __rm.marks.length +
      ',' + Math.round(__rm.kCol * 1000) + ',' + Math.round(__rm.kRow * 1000) + ',' + cvW + ',' + cvH + ',' + mVer + ',' + dx + ',' + dy + ',' + editRef + ',' + curSheet
    if (sig === __rm.sig) return          /* 坐标没变：一帧都不碰 DOM */
    __rm.sig = sig
    var host = doc.getElementById('__cb_ref_marks')
    if (!host) {
      host = doc.createElement('div')
      host.id = '__cb_ref_marks'
      host.style.cssText = 'position:absolute;pointer-events:none;z-index:9996;overflow:hidden;'
      t.cv.parentElement.appendChild(host)
    }
    host.style.left = t.cv.offsetLeft + 'px'
    host.style.top = t.cv.offsetTop + 'px'
    host.style.width = cvW + 'px'
    host.style.height = cvH + 'px'
    var kids = host.children
    while (kids.length > __rm.marks.length) { try { host.removeChild(host.lastChild) } catch (eR2) { break } }
    while (kids.length < __rm.marks.length) { var dv = doc.createElement('div'); dv.style.cssText = 'position:absolute;display:none;'; host.appendChild(dv) }
    for (var k = 0; k < __rm.marks.length; k++) {
      var mk = __rm.marks[k]
      var el = kids[k]
      var key = String(mk.ref || '').replace(/[$]/g, '').toUpperCase()
      if (editRef && key === editRef) {          /* 正在编辑的格：不画（别挡输入） */
        if (el.style.display !== 'none') el.style.display = 'none'
        continue
      }
      if (curSheet && mk.sheet && String(mk.sheet) !== curSheet) {   /* 别的表上的引用格 */
        if (el.style.display !== 'none') el.style.display = 'none'
        continue
      }
      var x1 = null, y1 = null, x2 = null, y2 = null
      var mr = useMeas ? __rm.meas.rects[key] : null
      /* 快照还没到手：宁可不画，也别先画到估的位置再跳回来 */
      if (!mr && __rm.holdUntil && __rm.holdUntil > 0) {
        var nowH = 0
        try { nowH = Date.now() } catch (eHn) { nowH = 0 }
        if (nowH < __rm.holdUntil) { if (el.style.display !== 'none') el.style.display = 'none'; continue }
      }
      if (mr) {
        /* ① 实测矩形 + 锚点位移：与单元格实际位置完全一致 */
        x1 = mr.x + dx
        y1 = mr.y + dy
        x2 = x1 + mr.w
        y2 = y1 + mr.h
      } else {
        var pc = /^([$]?[A-Z]{1,3})([$]?[0-9]+)$/.exec(key)
        if (!pc) { if (el.style.display !== 'none') el.style.display = 'none'; continue }
        var cc = 0
        var sl = pc[1]
        for (var j = 0; j < sl.length; j++) cc = cc * 26 + (sl.charCodeAt(j) - 64)
        cc -= 1
        var rr = Number(pc[2]) - 1
        var sp = __rmSpan(rr, cc)
        /* ② 兜底：按标定后的模型推（还没量到的格） */
        x1 = baseX + __rmPrefixX(sp.c1)
        y1 = baseY + __rmPrefixY(sp.r1)
        x2 = baseX + __rmPrefixX(sp.c2 + 1)
        y2 = baseY + __rmPrefixY(sp.r2 + 1)
      }
      var vw = x2 - x1, vh = y2 - y1
      if (!(vw > 0) || !(vh > 0) || x2 < 0 || y2 < 0 || x1 > cvW || y1 > cvH) {
        if (el.style.display !== 'none') el.style.display = 'none'
        continue
      }
      if (el.style.display !== 'block') el.style.display = 'block'
      el.style.left = x1 + 'px'
      el.style.top = y1 + 'px'
      el.style.width = vw + 'px'
      el.style.height = vh + 'px'
      el.style.background = __rmArgbCss(mk.color)
    }
  }
  function __rmTick () {
    __rm.raf = 0
    if (!__rm.on) return
    try { __rmPaint() } catch (eT0) { /* 单帧失败不打断循环 */ }
    try { __rm.raf = window.requestAnimationFrame(__rmTick) } catch (eT1) { __rm.raf = window.setTimeout(__rmTick, 60) }
  }
  function __rmStop () {
    if (__rm.raf) { try { window.cancelAnimationFrame(__rm.raf) } catch (eT2) { /* 忽略 */ } }
    __rm.raf = 0
    __rm.sig = ''
    __rm.meas = null
    __rm.needMeas = false
    var t0 = __hiFrame()
    var el0 = t0 && t0.doc.getElementById('__cb_ref_marks')
    if (el0 && el0.parentNode) el0.parentNode.removeChild(el0)
  }
  function __rmSet (data) {
    if (!data || typeof data !== 'object') return
    var dOn = data.on === true || data.on === '1'
    var dMarks = Array.isArray(data.marks) ? data.marks : []
    var dGeom = data.geom || null
    var dSig = ''
    try { dSig = (dOn ? '1' : '0') + '|' + JSON.stringify(dMarks) + '|' + JSON.stringify(dGeom) } catch (eRS0) { dSig = '' }
    /* ★ 宿主会重复推送（documentReady 后 0/800ms；refMarks/src 变化后 0/1500/4000ms）。
     *   旧实现每次推送都清快照 ⇒ 先按模型估位画（偏）→ 再测量（正）= 肉眼看到一跳。
     *   内容没变就当心跳：只保证 rAF 在跑，其余状态一律不动。 */
    if (dSig && dSig === __rm.dataSig && dOn === (__rm.on === true) && __rm.on && __rm.marks.length) {
      if (!__rm.raf) {
        try { __rm.raf = window.requestAnimationFrame(__rmTick) } catch (eRS1) { __rm.raf = window.setTimeout(__rmTick, 60) }
      }
      return
    }
    __rm.dataSig = dSig
    /* 地址集合没变（只是开关或颜色变）⇒ 旧快照 rects 依然准确，保留它，切换零跳转 */
    var keepMeas = false
    var prevMeas = __rm.meas
    /* 快照能不能留：除地址集合全覆盖外，**快照锚必须还是当前格** ——
     * 锚都换了的旧快照只能配合旧锚使用，留着必然画出错位（实测锚死在 B3、实际在 A1）。 */
    var curAddr = ''
    try {
      var fCur = __hiFrame()
      var nbCur = fCur && fCur.doc.getElementById('ce-cell-name')
      curAddr = nbCur && nbCur.value ? String(nbCur.value).replace(/[$]/g, '').toUpperCase() : ''
    } catch (eCA0) { curAddr = '' }
    if (prevMeas && prevMeas.refN && prevMeas.rects && dOn && dMarks.length &&
      prevMeas.refN === curAddr) {
      keepMeas = true
      for (var iK = 0; iK < dMarks.length; iK++) {
        var kK = String((dMarks[iK] && dMarks[iK].ref) || '').replace(/[$]/g, '').toUpperCase()
        if (!kK || !prevMeas.rects[kK]) { keepMeas = false; break }
      }
    }
    __rm.on = dOn
    __rm.marks = dMarks
    __rm.geom = dGeom
    __rm.sig = ''
    if (!keepMeas) {
      __rm.meas = null
      __rm.needMeas = true
      __rm.recheckLeft = 1
      __rm.needFrom = 0
      /* 还没测得真值期间先不画：避免先跳到「估的位置」再跳回来（地址真变时才走这里） */
      try { __rm.holdUntil = Date.now() + 1200 } catch (eRS2) { __rm.holdUntil = 0 }
    }
    if (!__rm.on || !__rm.marks.length) { __rmStop(); return }
    __rmPaint()
    if (!__rm.raf) {
      try { __rm.raf = window.requestAnimationFrame(__rmTick) } catch (eT3) { __rm.raf = window.setTimeout(__rmTick, 60) }
    }
    try {
      var t1 = __hiFrame()
      var api1 = t1 && (t1.frame.contentWindow.editor || (t1.frame.contentWindow.Asc && t1.frame.contentWindow.Asc.editor))
      if (!__rm.bound) {
        __rm.bound = true
        /* 滚动/缩放后重锚定：把累积误差归零（纯滚动本可按平移跟随，这里只为消除漂移） */
        try {
          t1.doc.addEventListener('wheel', function (ev) {
            /* ★ 只认 Ctrl+滚轮（= 缩放）。普通滚轮不需要重量：绘制用的是「锚点实时位移」，
             *   滚多远都精确；重量反而每 250ms 跳一次选区，滚动时肉眼可见地抖。 */
            if (ev && ev.ctrlKey && __rm.on) __rm.needMeas = true
          }, true)
        } catch (eW0) { /* 忽略 */ }
        try {
          t1.doc.addEventListener('keydown', function (ev) {
            if (!ev || !ev.ctrlKey) return
            var kk = ev.key
            if (kk === '+' || kk === '-' || kk === '=' || kk === '0') __rm.needMeas = true
          }, true)
        } catch (eK0) { /* 忽略 */ }
      }
      if (api1 && typeof api1.asc_registerCallback === 'function') {
        api1.asc_registerCallback('asc_onSelectionChanged', function () {
          /* ★ 静默窗内（= 正在测量 / 测量刚收尾）的选区变化是测量自己造成的，不能当用户操作 */
          var nowS = 0
          try { nowS = Date.now() } catch (eS7) { nowS = 0 }
          if (!__rm.on || __rm.measuring || nowS < (__rm.quietUntil || 0)) return
          __rm.selAt = nowS
          __rm.recheckLeft = 1
          /* 只有「锚点格换了」才重量（拖选时锚点不变，避免每帧跳）。
           * ★ 换格时在这里**同步**量，不等下一帧的 180ms 稳定窗：否则换格后要先按模型
           *   估画一两百毫秒、再跳到实测位置 —— 每点一下就闪两下（实测点 6 格出现 5 种布局、
           *   8 次跳动，用户看到的就是「闪来闪去」）。同步测量在同一个 JS 任务内完成，
           *   浏览器来不及重绘，只有最终位置可见。 */
          if (!__rm.measuring) {
            try {
              var nb2 = __hiFrame()
              var a2 = nb2 && nb2.doc.getElementById('ce-cell-name')
              var cur = a2 && a2.value ? String(a2.value).replace(/[$]/g, '').toUpperCase() : ''
              var old = __rm.meas && __rm.meas.ref ? String(__rm.meas.ref).replace(/[$]/g, '').toUpperCase() : ''
              if (cur !== old) {
                if (nowS - (__rm.measAt || 0) >= 250) {
                  __rm.measAt = nowS
                  __rm.needMeas = false
                  try { __rmMeasure(nb2, api1, a2, nb2.doc) } catch (eSm0) { __rm.needMeas = true }
                } else {
                  /* 连续换格（拖选/方向键连按）：250ms 内不重复跳，交给下一帧节流后量 */
                  __rm.needMeas = true
                  __rm.needFrom = 0
                }
              }
            } catch (eS9) { __rm.needMeas = true }
          }
          __rm.sig = ''
        })
      }
    } catch (eR4) { /* 轮询兜底 */ }
  }
  window.addEventListener('message', function (ev) {
    var d = ev.data || {}
    if (d && typeof d === 'object' && '__ooRowColHi' in d) {
      try { __hiSet(d.__ooRowColHi === true || d.__ooRowColHi === '1') } catch (eHi) {}
      return
    }
    if (d && typeof d === 'object' && '__ooRefMarks' in d) {
      try { __rmSet(d.__ooRefMarks) } catch (eRm) {}
      return
    }
    if (d.__ooIsModifiedAsk) {
      try { parent.postMessage({ __ooIsModified: dirty }, '*') } catch (e) {}
      return
    }
    if (d.__ooSavedAck) { window.__setOoDirty(false, 'savedAck'); return }
    if (d.__ooSaveReq) {
      try {
        var deep = document.querySelector('iframe[name="frameEditor"]')
        var w = deep && deep.contentWindow
        var api = w && w.editor
        if (api && typeof api.asc_Save === 'function') { api.asc_Save(); status('已触发引擎保存（asc_Save）') }
        else status('⚠ 深帧无 asc_Save，等宿主 forcesave')
      } catch (e) { status('⚠ 保存触发异常: ' + String(e).slice(0, 80)) }
    }
  })
  if (!src) { tip('待命'); status('页面就绪（无 src，等待宿主上传后重开）'); return }
  fetch('/poc-meta?src=' + encodeURIComponent(src), { headers: { Accept: 'application/json' } })
    .then(function (r) { return r.json() })
    .then(function (m) {
      if (!m || !m.ok) throw new Error((m && m.err) || 'meta 失败')
      window.__docKey = m.key
      var s = document.createElement('script')
      s.src = m.dsPublic + '/web-apps/apps/api/documents/api.js'
      s.onload = function () {
        try {
          editor = new DocsAPI.DocEditor('editor', {
            type: 'desktop',
            width: '100%',
            height: '100%',
            documentType: m.documentType,
            document: {
              fileType: m.fileType,
              key: m.key,
              title: m.title,
              url: m.fileUrl,
              permissions: { edit: true, download: true, print: true }
            },
            editorConfig: {
              callbackUrl: m.callbackUrl,
              lang: 'zh-CN',
              user: { id: 'canvas-user', name: '画布用户' },
              customization: { compactHeader: true, hideRightMenus: false, uiTheme: 'theme-classic-light' }
            },
            events: {
              onAppReady: function () {
                document.getElementById('tip').style.display = 'none'
                status('编辑器就绪 key=' + m.key)
                audit('编辑器', '远程打开', '工作区=' + src + ' key=' + m.key + ' DS=' + m.dsPublic)
              },
              onDocumentStateChange: function (e) { window.__setOoDirty(!!(e && e.data), 'onDocumentStateChange') },
              onError: function (e) { status('❌ 编辑器错误 ' + JSON.stringify(e || {}).slice(0, 120)); audit('编辑器', '远程错误', JSON.stringify(e || {}).slice(0, 200)) },
              onWarning: function (e) { status('⚠ ' + JSON.stringify(e || {}).slice(0, 120)) }
            }
          })
        } catch (e) {
          tip('编辑器初始化失败: ' + String(e).slice(0, 160))
          status('❌ DocsAPI 初始化失败: ' + String(e).slice(0, 120))
        }
      }
      s.onerror = function () {
        tip('无法加载文档服务脚本\\n' + m.dsPublic + '\\n请确认群晖上 OnlyOffice 文档服务容器已启动（端口 8080）')
        status('❌ api.js 加载失败 DS=' + m.dsPublic)
        audit('编辑器', '远程失败', 'api.js 加载失败 DS=' + m.dsPublic)
      }
      document.body.appendChild(s)
    })
    .catch(function (e) {
      tip('打开失败: ' + String(e).slice(0, 160))
      status('❌ poc-meta 失败: ' + String(e).slice(0, 120))
    })
})()
</script>
</body>
</html>`
}

/**
 * @param {{req: import('node:http').IncomingMessage, res: import('node:http').ServerResponse,
 *          url: URL, OFFICE_DIR: string, host?: string, log?: (line: string) => void}} ctx
 * @returns {Promise<boolean>} true = 已接管响应
 */
export async function handleOoRemote (ctx) {
  const { req, res, url, OFFICE_DIR, log } = ctx
  if (!isOoRemote()) return false
  const path = url.pathname
  const dir = wsDir(OFFICE_DIR)

  /* ── DS 回访通道（公开路径，token 门禁）────────────────────── */
  if (path.startsWith('/office-src/')) {
    const m = /^\/office-src\/([0-9a-f]{32})\/([A-Za-z0-9._-]+)$/.exec(path)
    const name = m ? safeName(m[2]) : ''
    if (!m || !name) { sendJson(res, 404, { error: 'bad path' }); return true }
    const meta = await readMeta(dir, name)
    if (!meta || meta.token !== m[1]) { sendJson(res, 403, { error: 'bad token' }); return true }
    try {
      const data = await readFile(join(dir, name))
      res.setHeader('Content-Type', REMOTE_MIME['.' + extOf(name)] || 'application/octet-stream')
      res.setHeader('Cache-Control', 'no-store')
      res.end(data)
    } catch (e) {
      sendJson(res, 404, { error: 'workspace missing', detail: String(e && e.message).slice(0, 80) })
    }
    return true
  }

  if (path === '/doc/callback' && req.method === 'POST') {
    const name = safeName(url.searchParams.get('ws'))
    const tok = String(url.searchParams.get('tok') || '')
    if (!name) { sendJson(res, 200, { error: 1 }); return true }
    const meta = await readMeta(dir, name)
    if (!meta || meta.token !== tok) { log?.(`[oo-remote] callback 拒绝 ws=${name}（token 不符）`); sendJson(res, 200, { error: 1 }); return true }
    let body = {}
    try { body = JSON.parse((await readBodyRaw(req, 1024 * 1024)).toString('utf8') || '{}') } catch { body = {} }
    const st = Number(body.status || 0)
    log?.(`[oo-remote] callback ws=${name} status=${st}`)
    if ((st === 2 || st === 6) && body.url) {
      try {
        /* DS 给的 url 可能带它自己视角的 host（容器名等）——统一改写成我们配的内网地址 */
        const u = new URL(String(body.url))
        const dl = new URL(dsInternal() + u.pathname + u.search)
        const r = await fetch(dl, { signal: AbortSignal.timeout(60000) })
        if (!r.ok) throw new Error('download ' + r.status)
        const buf = Buffer.from(await r.arrayBuffer())
        await mkdir(dir, { recursive: true })
        await writeFile(join(dir, name + '.saved'), buf)
        meta.hasSaved = true
        meta.savedAt = Date.now()
        meta.savedSize = buf.length
        await writeFile(metaPath(dir, name), JSON.stringify(meta))
        log?.(`[oo-remote] 产物已落盘 ws=${name} ${buf.length}B status=${st}`)
      } catch (e) {
        log?.(`[oo-remote] 产物下载失败 ws=${name}: ${String(e && e.message).slice(0, 120)}`)
      }
    }
    sendJson(res, 200, { error: 0 })
    return true
  }

  /* ── 以下全部在口令门禁后面 ───────────────────────────────── */

  if (path === '/poc') {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' })
    res.end(pocPage())
    return true
  }

  if (path === '/poc-meta') {
    const name = safeName(url.searchParams.get('src'))
    if (!name) { sendJson(res, 200, { ok: false, err: 'bad src' }); return true }
    let meta = await readMeta(dir, name)
    const wsFile = join(dir, name)
    try { await stat(wsFile) } catch { sendJson(res, 200, { ok: false, err: 'workspace missing' }); return true }
    if (!meta) {
      /* 工作区文件在但 meta 丢（异常重启）→ 补建 */
      meta = { token: randomBytes(16).toString('hex'), gen: 1 }
      await writeFile(metaPath(dir, name), JSON.stringify(meta))
    }
    const ext = extOf(name)
    if (!ext) { sendJson(res, 200, { ok: false, err: '只支持 xlsx/docx' }); return true }
    const origin = String(process.env.APP_OO_ORIGIN || ('http://' + (ctx.host || '127.0.0.1'))).replace(/\/$/, '')
    const dsPublic = String(process.env.APP_OO_DS_PUBLIC || ('http://' + ((ctx.host || '').split(':')[0] || '127.0.0.1') + ':8080')).replace(/\/$/, '')
    sendJson(res, 200, {
      ok: true,
      name,
      title: name,
      fileType: ext,
      documentType: docTypeOf(ext),
      key: keyOf(name, meta.gen || 1),
      dsPublic,
      fileUrl: `${origin}/office-src/${meta.token}/${name}`,
      callbackUrl: `${origin}/doc/callback?ws=${encodeURIComponent(name)}&tok=${meta.token}`,
    })
    return true
  }

  if (path === '/office-offline/open' && req.method === 'POST') {
    let body = {}
    try { body = JSON.parse((await readBodyRaw(req)).toString('utf8') || '{}') } catch { body = {} }
    const name = safeName(body.name)
    if (!name) { sendJson(res, 200, { ok: false, err: 'bad name' }); return true }
    const ext = extOf(name)
    if (!ext) { sendJson(res, 200, { ok: false, err: '远程模式只支持 .xlsx/.docx' }); return true }
    let buf
    if (body.dataBase64) {
      try { buf = Buffer.from(String(body.dataBase64), 'base64') } catch { sendJson(res, 200, { ok: false, err: 'bad base64' }); return true }
    } else {
      buf = Buffer.from(ext === 'xlsx' ? BLANK_XLSX_B64 : BLANK_DOCX_B64, 'base64')
      if (!buf.length) { sendJson(res, 200, { ok: false, err: '空白模板缺失' }); return true }
    }
    await mkdir(dir, { recursive: true })
    await writeFile(join(dir, name), buf)
    const meta = { token: randomBytes(16).toString('hex'), gen: 1, createdAt: Date.now() }
    await writeFile(metaPath(dir, name), JSON.stringify(meta))
    log?.(`[oo-remote] open ws=${name} ${buf.length}B`)
    sendJson(res, 200, { ok: true, bytes: buf.length })
    return true
  }

  if (path === '/office-offline/save' && req.method === 'POST') {
    let body = {}
    try { body = JSON.parse((await readBodyRaw(req, 1024 * 1024)).toString('utf8') || '{}') } catch { body = {} }
    const name = safeName(body.name)
    if (!name) { sendJson(res, 200, { ok: false, err: 'bad name' }); return true }
    const savedPath = join(dir, name + '.saved')
    try {
      const data = await readFile(savedPath)
      sendJson(res, 200, { ok: true, dataBase64: b64(data), digest: null })
    } catch {
      sendJson(res, 200, { ok: false, err: '远程 DS 尚无保存回执（forcesave 未完成或无改动）' })
    }
    return true
  }

  if (path === '/office-offline/saveFromValues') {
    sendJson(res, 200, { ok: false, err: '远程模式暂不支持「从数值新建」（需本机转换引擎）' })
    return true
  }

  if (path === '/doc/command' && req.method === 'POST') {
    let body = {}
    try { body = JSON.parse((await readBodyRaw(req, 1024 * 1024)).toString('utf8') || '{}') } catch { body = {} }
    try {
      const r = await fetch(dsInternal() + '/command', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
        signal: AbortSignal.timeout(15000),
      })
      const txt = await r.text()
      res.setHeader('Content-Type', 'application/json; charset=utf-8')
      res.statusCode = r.status
      res.end(txt)
    } catch (e) {
      sendJson(res, 502, { error: 'command forward failed', detail: String(e && e.message).slice(0, 120) })
    }
    return true
  }

  return false
}
