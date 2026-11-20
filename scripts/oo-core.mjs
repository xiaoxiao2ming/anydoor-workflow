/**
 * OnlyOffice 离线内嵌 POC 服务
 *
 *   /poc（由 poc-server:5199 或 app-server:3799 挂载，同源访问）
 *
 * 架构：
 *   /oo-assets/*        vendors/onlyoffice/ 静态（web-apps/sdkjs/fonts）
 *                       index.html 动态注入 alert 防冻结/消息监控（★ 不再注入 /bridge.js——
 *                       它定义的 AscDesktopEditor 会让 sdkjs 走桌面打开流卡死，2026-09-30 定案）
 *   /work/<file>        POC 工作区（lidu.docx / lidu.editor.bin）
 *   /convert POST       spawn x2t.exe 双向转换（cwd 必须 = converter 目录）
 *   /bridge.js          AscDesktopEditor 桥（保留路由但不再注入，仅供将来壳实验）
 *   /poc                测试页：DocsAPI 标准服务器流挂载（不传 type:'desktop'）
 *
 * 依据 docs/onlyoffice-offline-poc.md 的签名级规格（sdk-all.js 反解）。
 */
import http from 'node:http'
import fs from 'node:fs'
import path from 'node:path'
import { spawn } from 'node:child_process'
import url from 'node:url'
import zlib from 'node:zlib'

const ROOT = path.join(path.dirname(url.fileURLToPath(import.meta.url)), '..')
const ASSETS = path.join(ROOT, 'vendors', 'onlyoffice')
const WORK = path.join(ROOT, 'sample-template', 'oo-poc')
const CONVERTER = path.join(ASSETS, 'converter')
const OO_SAMPLES_DIR = path.join(ROOT, 'server-data', 'office') // ★ 常用样例（标题栏下拉）
const FONTS_DIR = path.join(ASSETS, 'fonts')
const PORT = 5199

/** AllFonts.js 服务策略：真·路径重写（2026-09-30 定案）。
 *  原版是 DesktopEditors 本机绝对路径（C:/Windows/Fonts、C:/Program Files/... 等 13 个根、642 项），
 *  桌面构建 sdkjs 字体引擎会把绝对路径包成 ascdesktop://fonts/<path> 去拉——Electron 的
 *  protocol.handle('ascdesktop') 实测从未被调到（0 命中）→ 字体全空 → onOpenDocument 之后卡
 *  "Loading document"。这里把所有绝对路径改写为 /sysfonts/<basename>（oo-core HTTP 直出，
 *  _fontPathMap 记录 basename→原绝对路径，天然覆盖全部根与子目录），其余变量按索引对齐不动。 */
let _allFontsCache = null
let _fontPathMap = null
function buildAllFonts() {
  if (_allFontsCache) return _allFontsCache
  const raw = fs.readFileSync(path.join(ASSETS, 'sdkjs', 'common', 'AllFonts.js'), 'utf8')
  _fontPathMap = {}
  let js = raw.replace(/"([A-Za-z]:\/[^"]+?\.(?:ttf|ttc|otf))"/gi, (m, p) => {
    const base = p.split('/').pop()
    if (!_fontPathMap[base]) _fontPathMap[base] = p
    return '"/sysfonts/' + encodeURIComponent(base) + '"'
  })
  // ★ g_fonts_selection_bin：base64 内嵌的绝对路径（C:\...）也全部改写为 /sysfonts/<basename>——
  //   引擎按名字查字体走的是这张选择表，不改则查到 C:\ 路径 → 拉不到 → 度量垃圾 → 瓦片 OOM
  //   （"选中切换字体报错"的根因，2026-10-01）。槽位定长、路径 null 结尾：新路径更短，
  //   null 补齐不破坏布局（latin1 往返安全）。Buffer 在本文件顶部已从 node: 引入。
  try {
    const sbRe = /(g_fonts_selection_bin"\] = ")([A-Za-z0-9+/=]+)"/
    const msb = js.match(sbRe)
    if (msb) {
      let bin = Buffer.from(msb[2], 'base64').toString('latin1')
      bin = bin.replace(/[A-Za-z]:\\[^\x00]+?\.(?:ttf|ttc|otf)/gi, (mm) => {
        const base = mm.split('\\').pop().split('/').pop()
        const np = '/sysfonts/' + base
        return np.length <= mm.length ? np + '\x00'.repeat(mm.length - np.length) : mm
      })
      js = js.replace(sbRe, (m0, head) => head + Buffer.from(bin, 'latin1').toString('base64') + '"')
    }
  } catch { /* 选择表改写失败不影响 files 主链路 */ }
  // ★ __font_names：文件索引 → 字体族名（编辑器帧瓦片绘制字体名用）
  try {
    // 变长槽位：[族名, 文件Idx, 样式, 文件Idx, 样式, ...]——中文字体条目常只有常规一档
    const infoRe = /\["([^"]+)"((?:,-?\d+)+)\]/g
    const fNames = []
    let mi
    while ((mi = infoRe.exec(raw))) {
      const nmv = mi[1]
      const nums = mi[2].split(',').filter(function (x) { return x !== '' }).map(Number)
      for (let j = 0; j + 1 < nums.length; j += 2) {
        const fIx = nums[j]
        if (fIx >= 0 && !fNames[fIx]) fNames[fIx] = nmv
      }
    }
    if (fNames.length) js += '\nwindow["__font_names"] = ' + JSON.stringify(fNames) + ';'
    // ★ __font_names_cn：英文名 → 中文名（中文族置顶显示用）
    var cnMap = {
      simsun: '宋体', nsimsun: '新宋体', simhei: '黑体', kaiti: '楷体', kaiti_gb2312: '楷体_gb2312',
      fangsong: '仿宋', fangsong_gb2312: '仿宋_gb2312', 'microsoft yahei': '微软雅黑', dengxian: '等线',
      'simsun-extg': '宋体-ExtB', stcaiyun: '华文彩云', stfangsong: '华文仿宋', stxihei: '华文细黑',
      stkaiti: '华文楷体', stsong: '华文宋体', stzhongsong: '华文中宋', lisu: '隶书', youyuan: '幼圆',
      mingliu: '细明体', pmingliu: '新细明体', 'ms gothic': 'MS 哥特体', 'ms mincho': 'MS 明朝',
      fzshuti: '方正舒体', fzxiaobiaosong: '方正小标宋', fzyaoti: '方正姚体', fzleiti: '方正隶变',
      fzbaosong: '方正报宋', fzheiti: '方正黑体', fzkaib: '方正楷体', fzsong: '方正宋体',
      fzcuyuan: '方正粗圆', fzxiyuan: '方正细圆', fzlanTingHei: '方正兰亭黑'
    }
    var cnOut = {}
    fNames.forEach(function (n) { if (!n) return; var lo = n.toLowerCase(); if (cnMap[lo]) cnOut[lo] = cnMap[lo] })
    js += '\nwindow["__font_names_cn"] = ' + JSON.stringify(cnOut) + ';'
  } catch { /* 名表生成失败不影响主链路 */ }
  _allFontsCache = js
  return _allFontsCache
}

const MIME = {
  '.js': 'application/javascript', '.css': 'text/css', '.html': 'text/html',
  '.json': 'application/json', '.png': 'image/png', '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml', '.woff': 'font/woff', '.woff2': 'font/woff2',
  '.ttf': 'font/ttf', '.otf': 'font/otf', '.bin': 'application/octet-stream',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.exe': 'application/octet-stream', '.wasm': 'application/wasm', '.dll': 'application/octet-stream',
  '.map': 'application/json', '.txt': 'text/plain', '.md': 'text/plain',
}

const send = (res, code, body, type = 'text/plain; charset=utf-8') => {
  try { const dt = Date.now() - (res.__ooT || Date.now()); if (dt > 100) fs.appendFileSync(path.join(ROOT, 'oo-perf.log'), `${dt}ms ${code} ${res.req ? res.req.url : (res.__ooUrl || '?')}\n`) } catch (_) {}
  res.setHeader('Access-Control-Allow-Origin', '*')
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type')
  // ★ 禁缓存（dev 服务器）：浏览器不再用旧缓存的 app.js/AllFonts——补丁改动刷新即见
  res.writeHead(code, {
    'Content-Type': type,
    'Access-Control-Allow-Origin': '*',
    'Cache-Control': 'no-cache, no-store, must-revalidate',
    Pragma: 'no-cache',
    Expires: '0',
  })
  res.end(body)
}

// ★ 2026-10-07 压缩打开时长：
//   sendFile —— 大文件（字体 15-17MB、sdk-all 10MB）改流式 pipe，消除大 Buffer 同步读
//   对事件循环的阻塞（冷首开实测：小 JSON 被大字体同步读排队到 700ms/项，15s 总长主要由
//   78 项资源串行慢速传输构成）。
//   ETag —— vendors 静态库内容不随会话变，mtime ETag + 304 协商缓存：二次打开零下载
//   （此前 no-cache 导致每次开节点 78 项 100MB 全量重传）。
const sendFile = (res, req, code, filePath, type, opts = {}) => {
  const __t = Date.now()
  const st = fs.statSync(filePath)
  const etag = 'W/"' + st.size + '-' + Math.round(st.mtimeMs) + '"'
  res.setHeader('Access-Control-Allow-Origin', '*')
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type')
  res.setHeader('ETag', etag)
  if (opts.negotiate && req && req.headers && req.headers['if-none-match'] === etag) {
    res.writeHead(304, { 'Cache-Control': 'no-cache' })
    res.end()
    return
  }
  res.writeHead(code, {
    'Content-Type': type,
    'Access-Control-Allow-Origin': '*',
    'Cache-Control': opts.cache || 'no-cache, no-store, must-revalidate',
    Pragma: 'no-cache',
    Expires: '0',
  })
  if (st.size > 262144) {
    // >256KB：流式，避免一次性大 Buffer 读入（事件循环不再被 17MB 同步读卡住）
    fs.createReadStream(filePath).on('error', () => { try { res.end() } catch (_) {} }).pipe(res)
    try { const dt = Date.now() - (res.__ooT || __t); if (dt > 100) fs.appendFileSync(path.join(ROOT, 'oo-perf.log'), `${dt}ms(stream) ${code} ${req && req.url ? req.url.slice(-60) : filePath.slice(-40)}\n`) } catch (_) {}
  } else {
    res.end(fs.readFileSync(filePath))
    try { const dt = Date.now() - (res.__ooT || __t); if (dt > 100) fs.appendFileSync(path.join(ROOT, 'oo-perf.log'), `${dt}ms ${code} ${req && req.url ? req.url.slice(-60) : filePath.slice(-40)}\n`) } catch (_) {}
  }
}

/** 注入桥脚本到编辑器 index.html（iframe 加载时桥已在——解决时序）。
 *  ★ 桥必须保留：桌面构建的引导依赖原生消息通道（execCommand webapps:entry/editor:config）。
 *    浏览器端字体走 ascdesktop:// 私有协议（只有 Electron 壳 protocol.handle 能拦截——已实现）。 */
function serveEditorIndex(res, filePath, req) {
  let html = fs.readFileSync(filePath, 'utf8')
  // ★ docKey 与 /poc 路由同款推导（inject 里 window.__docKey 要用）
  let srcFile = 'lidu.docx'
  try {
    var __u = new URL(req.url, 'http://127.0.0.1')
    var __src = __u.searchParams.get('src')
    if (__src) srcFile = String(__src).replace(/[\\/]/g, '_')
  } catch (e) {}
  const docKey = 'oo-' + srcFile.replace(/[^\w.-]+/g, '-')
  /* ★ 不再注入 /bridge.js（2026-09-30 A/B 实测定案）：
   *  桥定义的 window.AscDesktopEditor 会让桌面构建 sdkjs Bic 走「桌面打开流」→
   *  LocalStartOpen() 等假壳喂文档（空壳）→ 永远 "Loading document" 卡死。
   *  剥掉桥后走标准服务器流（jUj）：document.url → DS 转换 → 5s documentReady（无头+Electron 双验证）。
   *  字体 ascdesktop:// 由 Electron 主进程 protocol.handle 网络层接管，不依赖此 JS 桥。 */
  const inject = `<script>fetch("/report",{method:"POST",body:"iframe-load:"+location.search})</script>
<script>(function(){var o=CanvasRenderingContext2D.prototype.createImageData;CanvasRenderingContext2D.prototype.createImageData=function(w,h,a){if(w*h>16777216){var c=document.createElement("canvas");c.width=8;c.height=8;return c.getContext("2d").createImageData(8,8)}return o.call(this,w,h,a)}})()</script><script>window.NATIVE_EDITOR_ENJINE=false;window.__docKey=${JSON.stringify(docKey)};
// ★ 非阻塞化 vendor 层弹窗：同源 iframe 里 window.alert 会冻结整个渲染进程（requirejs 超时弹空框致全应用卡死）
window.alert=function(m){try{fetch("/report",{method:"POST",body:"[alert] "+String(m).slice(0,120)})}catch(e){}};
window.confirm=function(m){try{fetch("/report",{method:"POST",body:"[confirm] "+String(m).slice(0,120)})}catch(e){}return true};
window.__dump=[];
(function(){window.addEventListener('message',function(t){try{var d=typeof t.data==='string'?JSON.parse(t.data):t.data;fetch('/report',{method:'POST',body:'recv:'+(d.command||d.event||'?')})}catch(e){}},false);
var om=window.on_native_message;window.on_native_message=function(c,p){window.__dump.push('nm:'+c+'='+String(p).slice(0,80));om&&om(c,p)};
var _send=window.postMessage.bind(window);window.postMessage=function(m,o){try{var d=typeof m==='string'?JSON.parse(m):m;window.__dump.push('pm:'+(d.event||d.command||'?')+' '+String(d.data||'').slice(0,100))}catch(e){};return _send(m,o)}})();
</script>
<style>
/* 断线提示接管 v2（2026-10-09 用户指令「做成长条状 而且不会影响对文档的任何操作」）：
 * 通栏细长条贴文档最底部，26px 高单行文案；全程 pointer-events:none 零拦截，
 * 4 秒自动淡出（重连提示再现时再亮一轮） */
.asc-synchronizetip.cb-conn-toast{position:fixed!important;left:0!important;right:0!important;top:auto!important;bottom:0!important;transform:none!important;margin:0!important;z-index:99999!important;border-radius:0!important;box-shadow:none!important;background:rgba(17,22,28,.92)!important;pointer-events:none!important;transition:opacity .3s ease,visibility .3s ease!important}
.asc-synchronizetip.cb-conn-toast .tip-arrow{display:none!important}
.asc-synchronizetip.cb-conn-toast>div{display:flex!important;align-items:center!important;justify-content:center!important;padding:0 12px!important;background:transparent!important}
.asc-synchronizetip.cb-conn-toast .tip-text{font-size:12px!important;line-height:26px!important;height:26px!important;color:#c9d1d9!important;max-width:none!important;white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important}
.asc-synchronizetip.cb-conn-toast .tip-text br{display:none!important}
.asc-synchronizetip.cb-conn-toast .tip-header{display:inline!important;font-size:12px!important;font-weight:600!important;color:#e3b341!important;margin:0 8px 0 0!important}
.asc-synchronizetip.cb-conn-toast .btn-close,.asc-synchronizetip.cb-conn-toast [class*="close"]{display:none!important}
.asc-synchronizetip.cb-conn-toast.cb-toast-hide{opacity:0!important;visibility:hidden!important}
</style>
<script>
(function(){
/* ★ 2026-10-10 桌面构建假壳闸门修复（无头浏览器实测铁证）：
 *   frameEditor 里存在残缺 AscDesktopEditor 桩，isSupportPlugins() 返回 undefined
 *   ⇒ sdkjs 插件服务 Mo() 的守卫判定「壳不支持插件」⇒ asc_pluginRun 静默跳过
 *   iframe 创建 ⇒ 插件永远无法加载。强制改写为支持。 */
function __fixDsStub(){
  try{
    if(window.AscDesktopEditor){
      window.AscDesktopEditor.isSupportPlugins=function(){return true}
      return true
    }
  }catch(e){}
  return false
}
if(!__fixDsStub()){
  var __n=0
  var __t=setInterval(function(){ if(__fixDsStub()||++__n>150) clearInterval(__t) },200)
}
})();(function(){
/* ★ 拼写检查默认关（2026-10-09）：DE 启动读 de-settings-spellcheck（缺省 true）→
 *   在 app.js 加载前写入 0，config features.spellcheck.mode=false 的兜底双保险 */
try{localStorage.setItem('de-settings-spellcheck','0')}catch(e){}
/* ★ 新功能引导气泡默认关（2026-10-18）：TooltipManager 用 localStorage.getItem(tip.name)
 *   判断是否显示过（关气泡时才写 1）→ 启动前全部预写，"新建按钮：格式"等气泡不再出现 */
try{['sse-help-tip-cellFormat','sse-help-tip-chart-tab','sse-help-tip-darkDocument','sse-help-tip-solver',
     'de-help-tip-chart-tab','de-help-tip-fill-status','de-help-tip-header-footer-tab','de-help-tip-multipage-view-statusbar','de-help-tip-pasteOptions','de-help-tip-signature',
     'pe-help-tip-chart-tab','pe-help-tip-gif-payback','pe-help-tip-pasteOptions','pe-help-tip-slideTheme','pe-help-tip-slideTransitions'
    ].forEach(function(k){try{localStorage.setItem(k,'1')}catch(eKT){}})}catch(e){}
var KEYS=['失去网络连接','正在尝试连接','连接失败','Connection is lost','Trying to connect']
function report(m){try{fetch('/report',{method:'POST',body:'conn-toast:'+m})}catch(e){}}
function arm(el){clearTimeout(el.__cbT);el.classList.remove('cb-toast-hide');el.__cbT=setTimeout(function(){el.classList.add('cb-toast-hide');report('auto-hidden')},4000)}
function mark(el){el.__cbMarked=true;el.classList.add('cb-conn-toast');arm(el);report('shown')}
function scan(){var tips=document.querySelectorAll('.asc-synchronizetip');for(var i=0;i<tips.length;i++){var el=tips[i];if(el.__cbMarked)continue;var txt=el.textContent||'',hit=false;for(var j=0;j<KEYS.length;j++){if(txt.indexOf(KEYS[j])>-1){hit=true;break}}if(hit)mark(el)}}
var mo=new MutationObserver(function(muts){for(var i=0;i<muts.length;i++){var m=muts[i],t=m.target;if(!t)continue;var e=(t.nodeType===1)?t:t.parentElement,s=e;while(s&&!(s.classList&&s.classList.contains('asc-synchronizetip')))s=s.parentElement;if(!s||!s.__cbMarked)continue;if(m.type==='attributes'&&t===s){if(m.attributeName==='class')continue;try{if(getComputedStyle(s).display==='none')continue}catch(e2){}}if(s.classList.contains('cb-toast-hide'))arm(s)}scan()})
function boot(){if(!document.body){setTimeout(boot,200);return}mo.observe(document.body,{childList:true,subtree:true,attributes:true,attributeFilter:['style','class']});scan()}
boot()
})()
</script>`
  html = html.replace(/<head(\s[^>]*)?>/i, (m) => m + inject)
  send(res, 200, html, 'text/html; charset=utf-8')
}

function readBody(req) {
  return new Promise((ok) => {
    const chunks = []
    req.on('data', (c) => chunks.push(c))
    req.on('end', () => ok(Buffer.concat(chunks)))
  })
}

/** x2t 转换：src → dst（按扩展名推断格式），cwd 必须 = converter 目录 */
function x2t(srcFile, dstFile) {
  return new Promise((ok, fail) => {
    const p = spawn('x2t.exe', [srcFile, dstFile], { cwd: CONVERTER, windowsHide: true })
    let log = ''
    p.stdout.on('data', (c) => (log += c))
    p.stderr.on('data', (c) => (log += c))
    p.on('close', (code) => {
      if (code === 0 && fs.existsSync(dstFile)) ok(dstFile)
      else fail(new Error(`x2t exit=${code} ${log.slice(0, 200)}`))
    })
    p.on('error', fail)
  })
}

const OOLOG = []
const SESS = new Map()

/* ★★ 路径前缀闸门（2026-10-06，修 C 端生产模式「窗口停在 about:blank」）
 *
 * 本函数末尾有个兜底分支 `send(res,404,'nf '+p); return true`，
 * 意思是「只要进了本函数、又没有匹配任何分支，就由我 404 掉」。
 * 在 poc-server(5199) 里这没问题 —— 那里整站都是 OO 的地盘；
 * 但 app-server 把它挂在**静态托管之前**，于是**根路径 `/` 也会被这个兜底 404**，
 * 页面 HTML 永远拿不到 ⇒ Electron 窗口停在 about:blank ⇒ C 端出包后打不开。
 *
 * ★ 这类「兜底即吞掉」的模块只能独占整站，不能被当作中间件挂在有静态托管的服务前面。
 * 这里显式声明边界：只处理 OO 自己的路径前缀，其余**原样放行**（返回 false），
 * 由调用方的静态分支接管。宁可让未匹配的 OO 路径走静态 404，也不能吃掉整站首页。
 */
const OO_PREFIXES = [
  '/doc/', '/cache/', '/oo-assets/', '/work/', '/workfile/', '/office-samples/',
  '/upload/',
  '/fonts/', '/sysfonts/', '/poc', '/office-offline', '/bridge.js',
  '/report', '/convert', '/dumplog',
  /* ★ 2026-10-23：PDF 转换端点（PDF 合并节点的服务端转换通道） */
  '/pdfconv',
  '/docxconv',
  /* ★ 2026-10-17：宿主事件日志通道。原先漏登记，请求在进分支前就被闸门放行到静态托管
   *   （返回 index.html）→ 宿主所有 /office-log 日志从来没写进 oo-events.log。 */
  '/office-log',
  /* ★ 2026-10-17：最近使用文件列表。同样漏登记 + 端点还被嵌进 saveFromValues 里，
   *   双杀 ⇒ 空白对话框的「最近使用」永远空。（openrecent 属 /office-offline 前缀，本来通） */
  '/office-recent',
]
function isOoPath (url) {
  if (!url) return false
  // 去掉查询串再判定
  const pathname = url.split('?')[0].split('#')[0]
  return OO_PREFIXES.some((pre) => pathname === pre || pathname.startsWith(pre))
}

// ★ 共享核心：poc-server(5199) 与 app-server(3799) 共用同一套 OO 请求处理。
//   handleOoRequest 返回 true=已处理。所有 URL 相对化（无 5199 硬编码）。
/* ★★ 2026-10-17：日志会话上下文必须放**模块级** ★★
 *   handleOoRequest 是「每个 HTTP 请求调用一次」的函数，把 __logSid/__logT0 定义在里面
 *   等于每个请求重新初始化 ⇒ 所有日志都是 +0ms，耗时统计完全失效（实测踩到）。
 *   放到模块作用域，跨请求保持，s16 才能算出真实的整链耗时。 */
let __logSid = ''
let __logT0 = 0

export async function handleOoRequest(req, res) {
  /* ★ 统一事件日志（2026-10-12 用户指令）：保存/打开/位置/成败 全记录到 oo-events.log
   * ⚠ 2026-10-17 上移到函数体最顶部：__evt 是 const（存在 TDZ），原先定义在
   *   /office-offline/open 端点**之后**，端点里一调用就抛
   *   "Cannot access '__evt' before initialization"（实测踩过）。必须定义在所有端点之前。 */
  const EVTLOG = path.join(WORK, '..', 'oo-events.log')
  /* ★ 2026-10-17「一步一步查」：日志要能回答四个问题 ——
   *   ① 这一次保存是哪一次？（会话串联）② 每一步花了多久？（性能/卡点）
   *   ③ 为什么走了这条分支？（判定依据）④ 出错的那一步是谁吞了异常？（不再静默）
   *   为此给每条日志加 `[sid +Nms]` 前缀；sid 由宿主在点 💾 时生成，
   *   随 __ooSaveReq 下发到 /poc，再由 /office-log 回传上来，整条链共用同一个 sid。 */
  const __setLogSid = (sid) => {
    if (!sid) return
    if (sid !== __logSid) { __logSid = String(sid); __logT0 = Date.now() }
  }
  const __evt = (type, detail, sidOverride) => {
    try {
      /* ★ 2026-10-17：日志轮转。加上逐步埋点后单次保存会产生 15+ 行，
       *   长期使用会无限增长。超过 512KB 时保留后 2000 行（不影响最近排查）。 */
      try {
        const st = fs.statSync(EVTLOG)
        if (st.size > 512 * 1024) {
          const keep = fs.readFileSync(EVTLOG, 'utf8').split(/\r?\n/).slice(-2000).join('\n')
          fs.writeFileSync(EVTLOG, keep + '\n')
        }
      } catch (eR) { /* 文件不存在=首次写 */ }
      const ts = new Date().toLocaleString('zh-CN', { hour12: false })
      const sid = sidOverride || __logSid
      const head = sid ? '[' + sid + (__logT0 ? ' +' + (Date.now() - __logT0) + 'ms' : '') + '] ' : ''
      fs.appendFileSync(EVTLOG, '[' + ts + '] ' + head + '[' + type + '] ' + detail + '\n')
    } catch (e) { /* 日志失败不影响主流程 */ }
  }
  /* ★ 每次进程启动写一条 SESSION，读取脚本据此切分「本轮」与历史 */
  if (!globalThis.__ooLogBooted) {
    globalThis.__ooLogBooted = 1
    __evt('SESSION', '服务进程启动 pid=' + process.pid + ' work=' + WORK)
  }
  res.__ooT = Date.now()  // ★ 性能打点：send/sendFile 里算服务端处理耗时，>100ms 记 oo-perf.log
  if (!__probeHost && req.headers && req.headers.host) __probeHost = String(req.headers.host)   // ★ 心跳守护启用信号
  // ★ 前缀闸门：不属于 OO 的路径直接放行（详见上方注释）。
  //   CORS 预检（OPTIONS）不受此限 —— 它在下面统一处理。
  if (req.method !== 'OPTIONS' && !isOoPath(req.url)) return false
  // CORS 预检（跨源 fetch 的 OPTIONS 必须先过）
  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*')
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type')
    res.statusCode = 204
    res.end()
    return
  }
  // ★ /doc/<key>/c/ 与 /cache/* 反代到官方 DocService(3798)：auth/permissions/文档下发全由官方代码处理
  //   x-forwarded-host 让 DocService 生成的下载 URL(cache/files) 指回浏览器同源（5173）——
  //   ★ 必须优先用 vite(xfwd) 传入的原始 host：changeOrigin 已把 req.headers.host 改写成 3799，
  //     若用 host 会让 DS 把 cache URL 指到 3799（跨源，编辑器下载会被 CORS 拦）
  if (req.url && (req.url.startsWith('/doc/') || req.url.startsWith('/cache/') || req.url.includes('/doc/'))) {
    const targetPath = req.url.replace('/oo-assets', '')
    const target = http.request(
      { host: '127.0.0.1', port: 3798, path: targetPath, method: req.method, headers: Object.assign({}, req.headers, { host: '127.0.0.1:3798', 'x-forwarded-host': req.headers['x-forwarded-host'] || req.headers.host || '127.0.0.1', 'x-forwarded-proto': 'http' }) },
      (r2) => {
        // ★ 兜底：本回调是异步的，若上游响应到达前 res 已被别处写过（理论上不会，但防重入）则丢弃，绝不再 writeHead
        if (res.headersSent || res.writableEnded) { try { r2.resume() } catch { /* ignore */ } return }
        res.writeHead(r2.statusCode || 200, r2.headers)
        r2.pipe(res)
      }
    )
    // ★ 错误兜底同样要规避「头已发送」：res 已结束/已发头时绝不再写，否则 ERR_HTTP_HEADERS_SENT 会拖垮整个 dev 进程
    target.on('error', () => { try { if (!res.headersSent && !res.writableEnded) { res.statusCode = 502; res.end('docservice down') } } catch { /* ignore */ } })
    req.pipe(target)
    return true
  }
  // ★ cbref 插件静态资源（批次 12 · ③ 数据联动）：走项目自有目录 server-data/oo-plugins，
  //   不塞进 vendors/（那是第三方产物区，只读）。编辑器按 GetInstallPlugins 返回的
  //   /oo-assets/plugins/ 前缀去取，这里把它转到真实目录。
  if (req.url && req.url.startsWith('/oo-assets/plugins/cbref/')) {
    // ★ 2026-10-10：必须剥查询串——插件管理器加载 index.html 会追加 ?lang=...&theme-type=...，
    //   不剥则带参文件名 404 → 插件页空白 → Asc.plugin 永远不出现
    const rel = decodeURIComponent(req.url.split('?')[0].slice('/oo-assets/plugins/cbref/'.length)).replace(/[\\/]+/g, path.sep)
    const base = path.join(ROOT, 'server-data', 'oo-plugins', 'cbref')
    const target = path.join(base, rel)
    // 防目录穿越：必须在 base 之内
    if (target.startsWith(base) && fs.existsSync(target) && fs.statSync(target).isFile()) {
      const mt = target.endsWith('.html') ? 'text/html; charset=utf-8'
        : target.endsWith('.js') ? 'application/javascript; charset=utf-8'
        : target.endsWith('.json') ? 'application/json; charset=utf-8'
        : 'application/octet-stream'
      send(res, 200, fs.readFileSync(target), mt)
    } else {
      send(res, 404, 'nf')
    }
    return
  }

  // OO 节点「打开文件」：POST /office-offline/open {name, dataBase64} → 写工作区
  //   ★ 必须平铺到 WORK/<name>：编辑器 cfg.url=/work/<name>，两者路径必须一致（docId 子目录曾致 404）
  //   ★ mkdir 防御（2026-10-07）：C 端包内首次上传时 WORK 不存在（打包不带工作区临时文件）→ ENOENT 500
  if (req.url === '/office-offline/open' && req.method === 'POST') {
    try {
      const body = JSON.parse((await readBody(req)).toString())
      const name = String(body.name || '').replace(/[\\/]/g, '_')
      const buf = Buffer.from(body.dataBase64 || '', 'base64')
      fs.mkdirSync(WORK, { recursive: true })
      fs.writeFileSync(path.join(WORK, name), buf)
      /* ★ 2026-10-17 用户指令「打开的文件/节点内容也写进日志」：
       *   打开时把「是什么文件、多大、里面有什么」一并落盘 —— 用于一眼分辨
       *   「打开就是空白」到底是文件本身空（上次保存存空了）还是渲染没出来。 */
      const dig = __ooDigest(buf, name)
      __evt('OPEN', '文件=' + name + ' 大小=' + buf.length + 'B ' + dig.text)
      // ★ 2026-10-09 01:30 撤销 x2t 上传旁路（2026-10-08 引入）：实测数据——
      //   x2t.exe 独立进程**每次启动 ~15s**（字体表+引擎 DLL 加载，固有开销，两次实测 15s/16s），
      //   而 lcid 修复后 DS 长驻进程转换同一文件仅 **0.4s**（01:25:46.5 open → 46.9 ok）。
      //   后台预转 15s 才出产物，门禁却因此在关键路径上白等 15s —— 旁路从"救命"变成"拖累"。
      //   DS converter 死亡的唯一根因（lcid ESM）已根治 + 90s 心跳自愈兜底，DS 直转即最优。
      send(res, 200, JSON.stringify({ ok: true, name, size: buf.length, digest: dig }), 'application/json')
    } catch (e) { send(res, 500, JSON.stringify({ ok: false, err: String(e).slice(0, 200) }), 'application/json') }
    return
  }
  /* ── ★ PDF 转换（2026-10-23 用户指令「让 ooword 转换成 pdf 并连接」）──
   *   POST /pdfconv {dataBase64, name} → 写 WORK → DocService(3798) converter
   *   （长驻 x2t，实测 0.4s/份——远快于 /convert 的独立 x2t 冷启 15s）→ PDF 字节 base64。   *   ★ 为什么必须服务端做：DS converter 的请求/产物都在 3798 上，渲染进程跨端口 fetch
   *   会被 CORS 拦；且宿主侧原先的 __ooConvertPdf 桥挂在 /poc iframe 的 window 上，
   *   宿主根本调不到（恒「转换会话未就绪」——PDF 合并节点转换链路断的根因）。
   *   fileuri 用请求 Host 拼（DS 同机回访 /work 静态路由）。 */
  if (req.url === '/pdfconv' && req.method === 'POST') {
    try {
      const body = JSON.parse((await readBody(req)).toString())
      const name = String(body.name || '').replace(/[\\/]/g, '_')
      const buf = Buffer.from(body.dataBase64 || '', 'base64')
      /* ★ 2026-10-23 乱码深挖：记录实际收到的 dataBase64 长度与头尾采样——
         前端发 886KB（base64 1.18M 字符）但 WORK 落盘 22B 碎片的悬案定位
         （服务端收到的到底是多长、头尾长什么样，一看便知前后端责任）。 */
      __evt('PDFCONV-I', 'name=' + name + ' body=' + (req.headers['content-length'] || '?') + 'B dataBase64.len=' + (body.dataBase64 || '').length + ' head=' + JSON.stringify(String(body.dataBase64 || '').slice(0, 24)) + ' tail=' + JSON.stringify(String(body.dataBase64 || '').slice(-24)))
      if (!name || !buf.length) { send(res, 400, JSON.stringify({ ok: false, err: '参数缺失（name/dataBase64）' }), 'application/json'); return }
      fs.mkdirSync(WORK, { recursive: true })
      fs.writeFileSync(path.join(WORK, name), buf)
      const host = String(req.headers.host || '127.0.0.1')
      const fileurl = `http://${host}/work/${encodeURIComponent(name)}`
      const ckey = 'pdfconv-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8)
      /* ★ 字段名 url 而非 fileuri；filetype 必填（二者皆实测踩坑：
       *   本版 DS 源码 converterservice.js 读 params.url（cmd.setUrl(params.url)），
       *   fileuri 静默落空；filetype 缺失同样 CONVERT_PARAMS(-88)→旧码 -7）。 */
      const cr = await fetch('http://127.0.0.1:3798/converter', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ async: false, filename: name, url: fileurl, key: ckey, outputtype: 'pdf', filetype: (name.split('.').pop() || 'docx').toLowerCase() }),
        signal: AbortSignal.timeout(60000),
      }).then((r) => r.json()).catch((e) => ({ error: String(e).slice(0, 200) }))
      if (!cr || cr.error || !cr.fileUrl) {
        __evt('PDFCONV-E', '文件=' + name + ' DS 应答=' + JSON.stringify(cr).slice(0, 240))
        send(res, 500, JSON.stringify({ ok: false, err: 'DS converter 失败: ' + JSON.stringify(cr).slice(0, 240) }), 'application/json')
        return
      }
      // fileUrl 可能是绝对或相对（相对时补 DS origin）
      const furl = String(cr.fileUrl).startsWith('http') ? String(cr.fileUrl) : 'http://127.0.0.1:3798' + String(cr.fileUrl)
      const pdf = await fetch(furl, { signal: AbortSignal.timeout(30000) }).then((r) => r.arrayBuffer()).catch(() => null)
      if (!pdf || !pdf.byteLength) {
        __evt('PDFCONV-E', '产物读取失败 ' + furl)
        send(res, 500, JSON.stringify({ ok: false, err: 'PDF 产物读取失败: ' + furl }), 'application/json')
        return
      }
      __evt('PDFCONV', '文件=' + name + ' → ' + pdf.byteLength + 'B')
      send(res, 200, JSON.stringify({ ok: true, dataBase64: Buffer.from(pdf).toString('base64'), size: pdf.byteLength }), 'application/json')
    } catch (e) { send(res, 500, JSON.stringify({ ok: false, err: String(e).slice(0, 200) }), 'application/json') }
    return
  }
  /* ★ 2026-10-26 套件「PDF转Word · 文字版直转」：POST /docxconv {dataBase64, name}
     与 /pdfconv 同链路（DS 长驻 converter，秒级）——方向反转 outputtype='docx'。
     文字版 PDF → 可编辑 docx；扫描件（图片型）转出来仍是图片页（套件中心有 OCR 套件补位）。 */
  if (req.url === '/docxconv' && req.method === 'POST') {
    try {
      const __t0 = Date.now()
      const body = JSON.parse((await readBody(req)).toString())
      const name = String(body.name || '').replace(/[\\/]/g, '_')
      const buf = Buffer.from(body.dataBase64 || '', 'base64')
      if (!name || !buf.length) { send(res, 400, JSON.stringify({ ok: false, err: '参数缺失（name/dataBase64）' }), 'application/json'); return }
      fs.mkdirSync(WORK, { recursive: true })
      fs.writeFileSync(path.join(WORK, name), buf)
      const host = String(req.headers.host || '127.0.0.1')
      const fileurl = `http://${host}/work/${encodeURIComponent(name)}`
      const ckey = 'docxconv-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8)
      const cr = await fetch('http://127.0.0.1:3798/converter', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ async: false, filename: name, url: fileurl, key: ckey, outputtype: 'docx', filetype: 'pdf' }),
        signal: AbortSignal.timeout(120000),
      }).then((r) => r.json()).catch((e) => ({ error: String(e).slice(0, 200) }))
      if (!cr || cr.error || !cr.fileUrl) {
        __evt('DOCXCONV-E', '文件=' + name + ' DS 应答=' + JSON.stringify(cr).slice(0, 240))
        send(res, 500, JSON.stringify({ ok: false, err: 'DS converter 失败: ' + JSON.stringify(cr).slice(0, 240) }), 'application/json')
        return
      }
      const furl = String(cr.fileUrl).startsWith('http') ? String(cr.fileUrl) : 'http://127.0.0.1:3798' + String(cr.fileUrl)
      const docx = await fetch(furl, { signal: AbortSignal.timeout(30000) }).then((r) => r.arrayBuffer()).catch(() => null)
      if (!docx || !docx.byteLength) {
        __evt('DOCXCONV-E', '产物读取失败 ' + furl)
        send(res, 500, JSON.stringify({ ok: false, err: 'DOCX 产物读取失败: ' + furl }), 'application/json')
        return
      }
      __evt('DOCXCONV', '文件=' + name + ' → ' + docx.byteLength + 'B took=' + (Date.now() - __t0) + 'ms')
      send(res, 200, JSON.stringify({ ok: true, dataBase64: Buffer.from(docx).toString('base64'), size: docx.byteLength, took: Date.now() - __t0 }), 'application/json')
    } catch (e) { send(res, 500, JSON.stringify({ ok: false, err: String(e).slice(0, 200) }), 'application/json') }
    return
  }
  // OO 节点「保存回写」：POST /office-offline/save {name}
  //   name = 工作区文件名 → 派生协作 key → 找 DocService 缓存 Editor.bin → x2t 反转 docx/xlsx → base64
  /* ── ★ 2026-10-11 保存链重构：值数组 → xlsx（零依赖写入器）────────────
     DS 的 Editor.bin 为空壳（协作通道不带编辑内容——今日字节级实锤），
     forcesave 保存的永远是空文档。改为：编辑器引擎 callCommand 直读值数组 →
     此处构建 xlsx → 覆盖工作区 → ooBytes64。DS 彻底退出保存链。 */
  function __crc32(buf) {
    let tbl = __crc32.tbl
    if (!tbl) {
      tbl = __crc32.tbl = new Int32Array(256)
      for (let i = 0; i < 256; i++) { let c = i; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; tbl[i] = c }
    }
    let c = -1
    for (let i = 0; i < buf.length; i++) c = tbl[(c ^ buf[i]) & 0xff] ^ (c >>> 8)
    return (c ^ -1) >>> 0
  }
  function __zipStore(files) {
    const chunks = [], central = []
    let offset = 0
    for (const name of Object.keys(files)) {
      const data = Buffer.from(files[name], "utf8")
      const crc = __crc32(data)
      const nameBuf = Buffer.from(name, "utf8")
      const lh = Buffer.alloc(30)
      lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0, 6)
      lh.writeUInt16LE(0, 8); lh.writeUInt16LE(0, 10); lh.writeUInt16LE(0, 12)
      lh.writeUInt32LE(crc, 14); lh.writeUInt32LE(data.length, 18); lh.writeUInt32LE(data.length, 22)
      lh.writeUInt16LE(nameBuf.length, 26)
      chunks.push(lh, nameBuf, data)
      const ch = Buffer.alloc(46 + nameBuf.length)
      ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6)
      ch.writeUInt16LE(0, 8); ch.writeUInt16LE(0, 10); ch.writeUInt16LE(0, 12)
      ch.writeUInt16LE(0, 14); ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(data.length, 20)
      ch.writeUInt32LE(data.length, 24); ch.writeUInt16LE(nameBuf.length, 28)
      ch.writeUInt32LE(offset, 42)
      nameBuf.copy(ch, 46)
      central.push(ch)
      offset += 30 + nameBuf.length + data.length
    }
    const cdStart = offset
    let cdSize = 0
    for (const c of central) cdSize += c.length
    const eocd = Buffer.alloc(22)
    eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(central.length, 8)
    eocd.writeUInt16LE(central.length, 10); eocd.writeUInt32LE(cdSize, 12)
    eocd.writeUInt32LE(cdStart, 16)
    return Buffer.concat([...chunks, ...central, eocd])
  }
  function __colLetter(c) {
    let s = ""
    c++
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26) }
    return s
  }
  function __buildXlsxFromRows(rows) {
    const esc = (v) => String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    const nR = rows.length, nC = Math.max.apply(null, rows.map(function (r) { return r.length }).concat([1]))
    let sheetRows = ""
    for (let r = 0; r < nR; r++) {
      let cells = ""
        for (let c = 0; c < nC; c++) {
          const cobj = rows[r][c]
          if (cobj == null || cobj === '') continue
          const isObj = typeof cobj === 'object' && !Array.isArray(cobj)
          /* ★ 2026-10-17：空对象 {} 是「这一格没内容」（提取时按 used range 铺满整块，
           *   大量格子只有空对象）。原实现会把它写成 inlineStr 空串 ⇒ 每个文件都拖着
           *   成百上千个空单元格（实测一行写到 Y1），既臃肿又让「非空单元格数」虚高。 */
          if (isObj && !cobj.f && (cobj.v == null || String(cobj.v) === '')) continue
          const ref = __colLetter(c) + (r + 1)
          if (isObj && cobj.f) {
            cells += '<c r="' + ref + '"><f>' + esc(String(cobj.f).replace(/^=/, '')) + '</f><v>' + esc(cobj.v == null ? '' : cobj.v) + '</v></c>'
          } else {
            const vv = isObj ? (cobj.v != null ? cobj.v : (cobj.t != null ? cobj.t : '')) : cobj
            const isNum = typeof vv === 'number' || (vv !== '' && !isNaN(vv) && !isNaN(parseFloat(vv)))
            cells += isNum ? '<c r="' + ref + '"><v>' + vv + '</v></c>' : '<c r="' + ref + '" t="inlineStr"><is><t>' + esc(vv) + '</t></is></c>'
          }
        }
      sheetRows += '<row r="' + (r + 1) + '">' + cells + '</row>'
    }
    const ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
    const rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    const wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
    const wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>'
    const sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + sheetRows + '</sheetData></worksheet>'
    return __zipStore({ '[Content_Types].xml': ct, '_rels/.rels': rels, 'xl/workbook.xml': wb, 'xl/_rels/workbook.xml.rels': wbRels, 'xl/styles.xml': styles, 'xl/worksheets/sheet1.xml': sheet })
  }
  /* ── ★ 2026-10-17 打开内容摘要（用户指令「打开的文件/节点内容写进日志」）─────────────
   * 用途：一眼分辨「打开是空白」到底因为文件本身是空的（上次保存存空了），还是渲染没出来。
   * 纯 Node 实现（不引第三方 zip 库）：最小 zip 遍历 + inflateRaw，取 sheet1/document 的文本。
   * ⚠ 必须定义在 handleOoRequest 的**顶层**——放进某个 if/try 块里会变成块级作用域，
   *   前面（如 /office-offline/open，行号更小）调用时会 ReferenceError → 端点 500（实测踩过）。 */
  function __zipRead (buf) {
    const out = {}
    try {
      let off = 0
      while (off + 30 < buf.length) {
        if (buf.readUInt32LE(off) !== 0x04034b50) break
        const method = buf.readUInt16LE(off + 8)
        const compSize = buf.readUInt32LE(off + 18)
        const nameLen = buf.readUInt16LE(off + 26)
        const extraLen = buf.readUInt16LE(off + 28)
        const nm = buf.slice(off + 30, off + 30 + nameLen).toString('utf8')
        const ds = off + 30 + nameLen + extraLen
        const data = buf.slice(ds, ds + compSize)
        try { out[nm] = (method === 0 ? data : zlib.inflateRawSync(data)).toString('utf8') } catch (eZ) {}
        let next = ds + compSize
        if (next + 16 <= buf.length && buf.readUInt32LE(next) === 0x08074b50) {
          next = ds + compSize + 16 + buf.readUInt32LE(next + 12)   // 有 data descriptor
        }
        off = next
      }
    } catch (eR) { /* 解析失败就返回已有部分 */ }
    return out
  }
  function __ooDigest (buf, name) {
    const d = { parts: 0, kind: '', cells: 0, chars: 0, sample: '', text: '' }
    try {
      const zn = __zipRead(buf)
      d.parts = Object.keys(zn).length
      if (/\.xlsx$/i.test(name) || zn['xl/workbook.xml']) {
        d.kind = 'xlsx'
        const sheets = Object.keys(zn).filter((k) => /^xl\/worksheets\/sheet\d+\.xml$/.test(k))
        d.text = '表数=' + sheets.length
        let total = 0
        const vals = []
        for (const s of sheets.slice(0, 3)) {
          // <t>=文本型，<v>=数值型；两者都算内容，否则纯数字表会被误判为空
          const t = (zn[s].match(/<(t|v)[^>]*>([\s\S]*?)<\/\1>/g) || []).map((m) => m.replace(/<[^>]+>/g, ''))
          const nn = t.filter((x) => String(x).trim() !== '')
          total += nn.length
          if (vals.length < 5) vals.push(...nn.slice(0, 5 - vals.length))
        }
        d.cells = total
        d.sample = vals.slice(0, 5).join(' | ').slice(0, 120)
        d.text += ' 非空单元格=' + total + (d.sample ? ' 前几项=' + d.sample : '（内容为空）')
      } else if (/\.docx$/i.test(name) || zn['word/document.xml']) {
        d.kind = 'docx'
        const t = (zn['word/document.xml'].match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g) || []).map((m) => m.replace(/<[^>]+>/g, '')).join('')
        d.chars = t.length
        d.sample = t.slice(0, 80)
        d.text = '文本长度=' + t.length + (t ? ' 开头=' + d.sample : '（内容为空）')
      } else {
        d.kind = 'other'
        d.text = '非 OOXML（部件 ' + d.parts + ' 个）'
      }
    } catch (e) {
      d.text = '摘要解析失败: ' + String(e).slice(0, 60)
    }
    return d
  }

  // ★ 新端点：值数组 → xlsx → 覆盖工作区 → base64 返回
  /* ★ 2026-10-17 修复：本端点原先被补丁塞在 saveFromValues 的 try 块「里面」，
   *   只有 req.url 命中 saveFromValues 时才可能被执行 —— 等于死代码，
   *   宿主侧所有 /office-log 日志（SAVE-NODE-OK / FAIL / SAVE-BTN…）从来没写进过 oo-events.log。
   *   现移回顶层，与其余端点同级。 */
  if (req.url === '/office-log' && req.method === 'POST') {
    try {
      const body = JSON.parse((await readBody(req)).toString())
      /* ★ 2026-10-17：接收 sid（会话串联）+ 放宽 detail 上限到 700 ——
       *   逐步排查时一行里要同时放下「行数/列数/非空格数/首行/摘要」，300 字符不够用。 */
      if (body.sid) __setLogSid(body.sid)
      __evt(String(body.type || 'RENDER'), String(body.detail || '').slice(0, 700))
      send(res, 200, JSON.stringify({ ok: true }), 'application/json')
    } catch (e) { send(res, 500, JSON.stringify({ ok: false }), "application/json") }
    return
  }
  /* ★ 2026-10-17 修复：这两段原先被补丁塞进 saveFromValues 的 try 块「里面」，
   *   外层 if 只认 /office-offline/saveFromValues ⇒ /office-recent、/openrecent 永远进不来，
   *   请求落静态 index.html ⇒ 「最近使用文件」列表恒空（与 /office-log 同一类坑）。
   *   端点一律放顶层。 */
  if (req.url && req.url.startsWith('/office-recent') && req.method === 'GET') {
    try {
      const u2 = new URL(req.url, 'http://x')
      const ext = u2.searchParams.get('ext') === 'docx' ? '.docx' : '.xlsx'
      const files = []
      for (const f of fs.readdirSync(WORK)) {
        if (!f.toLowerCase().endsWith(ext)) continue
        if (f.includes('-blank-')) continue
        try { const st = fs.statSync(path.join(WORK, f)); files.push({ name: f, mtime: st.mtimeMs, size: st.size }) } catch (eS) {}
      }
      files.sort((a, b) => b.mtime - a.mtime)
      send(res, 200, JSON.stringify({ ok: true, files: files.slice(0, 8) }), 'application/json')
    } catch (e) { send(res, 500, JSON.stringify({ ok: false, err: String(e).slice(0, 120) }), "application/json") }
    return
  }
  if (req.url === '/office-offline/openrecent' && req.method === 'POST') {
    try {
      const body = JSON.parse((await readBody(req)).toString())
      const fname = String(body.file || '').replace(/[\\/]/g, '_')
      if (!fname || fname.includes('..')) { send(res, 400, JSON.stringify({ ok: false, err: '非法文件名' }), 'application/json'); return }
      const fp = path.join(WORK, fname)
      if (!fs.existsSync(fp)) { send(res, 404, JSON.stringify({ ok: false, err: "文件不存在" }), "application/json"); return }
      const buf = fs.readFileSync(fp)
      __evt('OPEN-RECENT', '最近文件=' + fname + ' 大小=' + buf.length + 'B')
      send(res, 200, JSON.stringify({ ok: true, name: fname, dataBase64: buf.toString('base64'), size: buf.length }), 'application/json')
    } catch (e) { send(res, 500, JSON.stringify({ ok: false, err: String(e).slice(0, 120) }), "application/json") }
    return
  }
  if (req.url === '/office-offline/saveFromValues' && req.method === 'POST') {
    try {
      const body = JSON.parse((await readBody(req)).toString())
      const name = String(body.name || '').replace(/[\\/]/g, '_')
      /* ★ 2026-10-22 修 -85（内容与扩展名不匹配）：本端点只产 xlsx——
         docx 请求走这里会把 xlsx 字节存进 .docx（ooword 节点 ooBytes64），
         下次打开必报 -85。docx 一律拒绝（ooword 必须走 OO 主通道原样导出）。 */
      if (/\.docx$/i.test(name)) {
        __evt('SAVE-REJECT', '目标=' + name + ' 是 docx——saveFromValues 只产 xlsx，拒绝（防 -85 内容错配）')
        send(res, 400, JSON.stringify({ ok: false, err: 'docx 不支持值数组重建（会内容错配 -85）——请用 OO 主通道保存' }), 'application/json')
        return
      }
      const rows = Array.isArray(body.rows) ? body.rows : []
      __evt('SAVE-START', '目标=' + name + ' 行数=' + rows.length + ' 模式=' + (body.mode || 'node'))
      const xbuf = __buildXlsxFromRows(rows)
      const outPath = path.join(WORK, name)
      fs.writeFileSync(outPath, xbuf)
      __evt('SAVE-OK', '工作区=' + outPath.replace(/\\/g, '/') + ' 大小=' + xbuf.length + 'B')
      /* ★ 2026-10-17 用户指令「保存路径放到日志里」：回传工作区绝对路径，宿主据此在诊断抽屉里显示 */
      send(res, 200, JSON.stringify({ ok: true, dataBase64: xbuf.toString('base64'), size: xbuf.length, path: outPath.replace(/\\/g, '/') }), 'application/json')
    } catch (e) { __evt('SAVE-FAIL', '原因=' + String(e).slice(0, 120)); send(res, 500, JSON.stringify({ ok: false, err: String(e).slice(0, 200) }), "application/json") }
    return
  }
  /* ★ 2026-10-17 用户指令「换个保存的地方我看看保存的是不是我修改的」：
   *   每次保存额外落一份【带时间戳】的留档文件到 sample-template/保存留档/，
   *   并在日志里写明【内容摘要】（非空单元格数 + 前几项）——用户双击文件即可核对。 */
  if (req.url === '/office-offline/archive' && req.method === 'POST') {
    try {
      const body = JSON.parse((await readBody(req)).toString())
      const nm = String(body.name || 'doc.xlsx').replace(/[\\/]/g, '_')
      const buf = Buffer.from(body.dataBase64 || '', 'base64')
      const dir = path.join(WORK, '..', '保存留档')
      fs.mkdirSync(dir, { recursive: true })
      /* ★ 本地时间（toISOString 是 UTC，会差 8 小时——已实测踩过） */
      const dNow = new Date()
      const p2 = (n) => String(n).padStart(2, '0')
      const ts = dNow.getFullYear() + '-' + p2(dNow.getMonth() + 1) + '-' + p2(dNow.getDate()) + '-' + p2(dNow.getHours()) + '-' + p2(dNow.getMinutes()) + '-' + p2(dNow.getSeconds())
      const out = path.join(dir, ts + '_' + nm)
      fs.writeFileSync(out, buf)
      const dig = __ooDigest(buf, nm)
      __evt('ARCHIVE', '留档=' + out.replace(/\\/g, '/') + ' 大小=' + buf.length + 'B ' + dig.text)
      send(res, 200, JSON.stringify({ ok: true, path: out.replace(/\\/g, '/'), digest: dig }), 'application/json')
    } catch (e) { send(res, 500, JSON.stringify({ ok: false, err: String(e).slice(0, 200) }), 'application/json') }
    return
  }
  if (req.url === '/office-offline/save' && req.method === 'POST') {
    try {
      const body = JSON.parse((await readBody(req)).toString())
      const name = String(body.name || '').replace(/[\\/]/g, '_')
      const key = 'oo-' + name.replace(/[^\w.-]+/g, '-')
      const cacheBin = path.join(ROOT, 'vendors', 'oo-server', 'App_Data', 'data', key, 'Editor.bin')
      if (!fs.existsSync(cacheBin)) {
        send(res, 200, JSON.stringify({ ok: false, err: '该会话还没有构建产物（编辑器未完成打开或未保存）' }), 'application/json')
        return
      }
      const ext = /\.(xlsx)$/i.test(name) ? '.xlsx' : '.docx'
      const outPath = path.join(WORK, 'save-' + key + ext)
      await x2t(cacheBin, outPath)
      const out = fs.readFileSync(outPath)
      /* ★ 2026-10-17 假成功拦截：Editor.bin 是「打开时用 x2t 从源文件转出来的」，
       *   里面从来不含编辑内容（见本文件 959-965 行的构建逻辑）。所以本端点转回来的
       *   产物 == 打开时的原文件；若源是空白新建，产物就是一张空表。
       *   但宿主那边只要「有字节」就当主通道成功（6258B 空表覆盖掉提取通道的 123），
       *   实测 18:59 那次就是这么把用户改的东西丢掉的。
       *   → 这里把内容摘要一并回传，由宿主判断「这份产物到底有没有内容」。 */
      const dig = __ooDigest(out, name)
      send(res, 200, JSON.stringify({ ok: true, dataBase64: out.toString('base64'), size: out.length, digest: dig }), 'application/json')
    } catch (e) { send(res, 500, JSON.stringify({ ok: false, err: String(e).slice(0, 200) }), 'application/json') }
    return
  }
  const u = new URL(req.url, `http://127.0.0.1:${PORT}`)
  const p = decodeURIComponent(u.pathname)


    /* ── 协作 socket 应答（Engine.IO v4 polling + socket.io 最小实现）──
       编辑器文档下发走这条通道（strict 模式也连）：open 事件带
       data['origin.docx'] —— sdkjs 内置 ooxml 解析器直接吃原始 docx。 */
    if (/^(?:\/oo-assets)?\/doc\/[^/]+\/c\/?$/.test(p)) {
      const isPost = req.method === 'POST'
      const sid = u.searchParams.get('sid')
      const transport = u.searchParams.get('transport') || 'polling'
      if (transport !== 'polling') { send(res, 400, 'polling only'); return }
      const sname = decodeURIComponent(p.split('/').filter(Boolean).slice(-2)[0])
      if (isPost) {
        const body = (await readBody(req)).toString()
        OOLOG.push('sock POST[' + sname.slice(0, 8) + '] sid=' + (sid || 'NONE') + ': ' + body.slice(0, 80))
        const dkey = 'doc:' + sname
        const st = SESS.get(dkey) || { q: [] }
        SESS.set(dkey, st)
        // 镜像应答：DocService 对 auth 的应答回含用户/权限/设置——最忠实做法是回显客户端 data + result
          try {
            const parsed = JSON.parse(body.startsWith('40') ? body.slice(1) : body.startsWith('42') ? body.slice(2) : body)
            const d = parsed.data || parsed
            if (d && d.type === 'auth') {
              st.q.push('42["message",' + JSON.stringify(Object.assign({}, d, { result: true })) + ']')
              OOLOG.push('sock auth mirror reply (openCmd=' + (d.openCmd ? 'yes' : 'no') + ')')
            }
          } catch (e) {}
        // 客户端 CONNECT(40) → 应答 CONNECT ack(40{sid}) + open 事件(42)
        if (body.includes('"type":"auth"') || body.startsWith('40')) {
            st.q.push('40{"sid":"' + sid + '"}')
          // 编辑器 auth 后要 editorPermissions + 文档下发
          st.q.push('42["message",{"type":"auth","result":true,"docid":"' + sname + '","licenseType":1,"settings":{},"permissions":{"edit":true,"review":true,"comment":true,"download":true,"print":true,"fillForms":true,"modifyFilter":true,"copy":true},"user":{"id":"uid-poc","name":"poc","indexUser":1},"sessionId":"poc-session","isCloseCoAuthoring":true,"coEditingMode":"strict","lang":"zh","mode":"edit"}]')
          st.q.push('42["message",{"type":"license","result":true,"licenseType":1}]')
          if (!st.opened) { st.opened = true; st.q.push('42["message",{"type":"open","status":"ok","openedAt":"' + new Date().toISOString() + '","data":{"origin.docx":"/work/lidu.docx"}}]') }
          st.q.push(JSON.stringify(['open', {
            status: 'ok',
            openedAt: new Date().toISOString(),
            data: { 'origin.docx': '/work/lidu.docx' }
          }]).replace(/^\[/, '42['))
        }
        send(res, 200, 'ok')
        return
      }
      if (!sid) {
        const nid = 'oo' + Math.random().toString(36).slice(2, 10)
        SESS.set(nid, { q: [] })
        OOLOG.push('sock handshake sid=' + nid)
        send(res, 200, '0' + JSON.stringify({
          sid: nid, upgrades: [], pingInterval: 25000, pingTimeout: 20000, maxPayload: 1000000
        }), 'text/plain; charset=utf-8')
        return
      }
      const st = SESS.get('doc:' + sname) || SESS.get(sid) || { q: [] }
      SESS.set('doc:' + sname, st)
      // 主动推送：队列空时直接下发 auth/license/open（不等客户端 POST auth）
      if (!st.q.length && !st.pushed) {
        st.pushed = true
        OOLOG.push('sock PUSH frames for sid=' + sid)
        st.q.push('40{"sid":"' + sid + '"}')
        st.q.push('42["message",{"type":"auth","result":true,"docid":"' + sname + '","licenseType":1,"settings":{},"permissions":{"edit":true,"review":true,"comment":true,"download":true,"print":true,"copy":true},"user":{"id":"uid-poc","name":"poc","indexUser":1},"sessionId":"poc-session","isCloseCoAuthoring":true,"coEditingMode":"strict","lang":"zh","mode":"edit"}]')
        st.q.push('42["message",{"type":"license","result":true,"licenseType":1}]')
        st.q.push('42["message",{"type":"open","status":"ok","openedAt":"' + new Date().toISOString() + '","data":{"origin.docx":"/work/lidu.docx"}}]')
      }
      // 长轮询：有帧立即返回，否则等 8s 再返回（避免空转）
      if (st.q.length) { const out = st.q.splice(0).join('\u001e'); send(res, 200, out, 'text/plain; charset=utf-8'); return }
      await new Promise((r) => setTimeout(r, 8000))
      send(res, 200, st.q.length ? st.q.splice(0).join('\u001e') : '6', 'text/plain; charset=utf-8')
      return
    }

  if (p === '/dumplog') {
    send(res, 200, JSON.stringify(OOLOG.slice(-40)))
    return
  }
  if (p === '/report' && req.method === 'POST') {
    const b = (await readBody(req)).toString()
    OOLOG.push(b.slice(0, 300))
    send(res, 200, 'ok')
    return
  }
  try {
    /* ── 桥 ── */
    if (p === '/bridge.js') {
      send(res, 200, `
function __rep(m) { try { fetch('/report', { method: 'POST', body: m.slice(0, 280) }) } catch (e) {} }
window.AscDesktopEditor = {
  // ★ IsLocalFile 必须回 false：回 true 时编辑器下载完 Editor.bin 仍走壳的 LocalStartOpen
  //   等壳推本地文件内容（假壳不推 → 永远 Loading）。回 false 则自行消费已下载的二进制。
  IsLocalFile: function () { __rep('IsLocalFile called'); return false },
  GetVersion: function () { return '9.4.0-offline-poc' },
  // convertFile(localPath, 8194, cb) —— cb(binObj)，binObj.get()=Uint8Array / close()
  convertFile: function (srcPath, fmt, cb) { __rep('convertFile: ' + srcPath);
    // srcPath 可能是 http URL——先下载 docx 到工作区，再 x2t 转 bin，再回读
    var dlName = 'bridge-dl-' + Date.now().toString(36) + '.docx'
    fetch(srcPath).then(function (r) { return r.arrayBuffer() }).then(function (buf) {
      var u8 = new Uint8Array(buf)
      var bin64 = ''
      var CH = 0x8000
      for (var i = 0; i < u8.length; i += CH) bin64 += String.fromCharCode.apply(null, u8.subarray(i, i + CH))
      return fetch('/office-offline/open', { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: dlName, dataBase64: btoa(bin64) }) }).then(function (x) { return x.json() })
    }).then(function (j1) {
      if (!j1.ok) throw new Error('open fail')
      return fetch('/office-offline/convert', { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: dlName }) }).then(function (x) { return x.json() })
    }).then(function (j2) {
      if (!j2.ok) throw new Error('convert fail')
      __rep('bridge convert ✓ ' + j2.binName)
      return fetch('/work/' + j2.binName).then(function (r2) { return r2.arrayBuffer() })
    }).then(function (buf) {
      var u8 = new Uint8Array(buf)
      cb({ get: function () { return u8 }, close: function () {} })
    }).catch(function (e) { console.error('[bridge] convert err', e); __rep('❌ convert: ' + e); cb(null) })
  },
  LocalFileGetImageUrl: function (u) { return u },
  GetImageBase: function (u) { return u },
  OpenFilenameDialog: function (kind, isSave, params, cb) { cb && cb(null) },
  GetInstallPlugins: function () {
    // 桌面版返回两组插件（editor/collaborative），sdkjs UpdateSystemPlugins 按下标取
    return JSON.stringify([
      { url: '/oo-assets/plugins/', pluginsData: [] },
      { url: '/oo-assets/plugins/', pluginsData: [] }
    ])
  },
  GetDictionariesPath: function () { return '/oo-assets/dictionaries/' },
  SpellCheck: function () {},
  CryptoMode: function () { return 0 },
  SetAdvancedOptions: function () {},
  SetLocalRestrictions: function () {},
  SetDocumentName: function () {},
  onDocumentModifiedChanged: function () {},
  onDocumentContentReady: function () {},
  CreateEditorApi: function (cfg) {
    __rep('CreateEditorApi')
    // 桌面版：壳返回编辑器 API 代理；浏览器里转发到 sdkjs 实例（window.Asc.editor / editor）
    var e = window.Asc && window.Asc.editor ? window.Asc.editor : (window.editor || null)
    return e
  },
  LocalStartOpen: function () { __rep('LocalStartOpen'); },
  execCommand: function (cmd, param) {
    __rep('execCommand: ' + cmd)
    // 壳→编辑器应答通道：编辑器索要 features / 请求 config 时，壳必须回
    if (cmd === 'webapps:features') {
      setTimeout(function () {
        try { window.on_native_message('window:features', JSON.stringify({ singlewindow: true, features: {} })); __rep('reply window:features') } catch (e) { __rep('reply err ' + e) }
      }, 0)
    }
    if (cmd === 'webapps:entry') {
      // 桌面版：壳把文档配置预置进 native_message_cmd，编辑器启动时重放
      setTimeout(function () {
        try {
          window.on_native_message('editor:config', JSON.stringify({
            documentType: 'word',
            document: { fileType: 'docx', key: 'lidu-poc', title: 'lidu.docx', url: '/work/lidu.docx', permissions: { edit: true } },
            editorConfig: { mode: 'edit', targetApp: 'desktop', coEditing: { mode: 'strict' } }
          }))
          __rep('reply editor:config')
        } catch (e) { __rep('cfg err ' + e) }
      }, 50)
    }
    return ''
  },
  SendCommand: function () {},
  SaveFilePath: function (p) { __rep('SaveFilePath: ' + p); },
  SetFilePath: function (p) { __rep('SetFilePath: ' + p); },
  GetDocumentName: function () { return 'lidu.docx' },
  OpenFile: function (p) { __rep('OpenFile: ' + p); },
  NativeViewerOpen: function () {},
  Print: function () {},
  DownloadFiles: function () {},
  GetEngineVersion: function () { return '9.4.0' },
  getEngineVersion: function () { return '9.4.0' },
  SetDocumentName: function (n) { __rep('SetDocumentName: ' + n) },
  onDocumentModifiedChanged: function () {},
  onDocumentContentReady: function () {},
  isSupportMacroses: function () { return false },
  isSupportPlugins: function () { return false },
  isSupportNetworkFunctionality: function () { return false },
  isBlockchainSupport: function () { return false },
  GetInstallPlugins: function () {
    // 桌面版返回两组插件（editor/collaborative），sdkjs UpdateSystemPlugins 按下标取
    return JSON.stringify([
      { url: '/oo-assets/plugins/', pluginsData: [] },
      { url: '/oo-assets/plugins/', pluginsData: [] }
    ])
  },
  GetSupportedScaleValues: function () { return '[1,1.25,1.5,1.75,2]' },
  LocalFileGetSourcePath: function (u) { return u },
  localSaveToDrawingFormat: function () {},
  LoadFontBase64: function () {},
  getDictionariesPath: function () { return '/oo-assets/dictionaries/' },
  DownloadFiles: function () {},
  isSupportNetworkFunctionality: function () { return false },
  isSupportPlugins: function () { return false },
  emulateCloudPrinting: function () {},
}
`, 'application/javascript')
      return
    }

    /* ── 转换桥 ── */
    if (p === '/convert' && req.method === 'POST') {
      const body = JSON.parse((await readBody(req)).toString())
      try {
        await x2t(body.src, body.dst)
        send(res, 200, JSON.stringify({ ok: true, dst: body.dst }))
      } catch (e) {
        send(res, 500, JSON.stringify({ ok: false, err: String(e).slice(0, 200) }))
      }
      return
    }

    /* ── 图片上传（2026-10-23 修「ooword/ooexcel 插入图片报错」）：
       sdkjs 插入本地图片 POST <origin>/upload/<docId>（octet-stream），
       应答 JSON { "<文件名>": "<可访问 url>" }——引擎再从 url 拉图插入。
       此前无此端点 → 404 → 弹「使用文档时出错…」。落 WORK 复用 /work 静态。 */
    if (p.startsWith('/upload/') && req.method === 'POST') {
      try {
        const body = await readBody(req)
        /* 嗅探图片类型（Content-Type 是 octet-stream 无扩展名可依） */
        let ext = 'png'
        if (body.length > 3 && body[0] === 0xff && body[1] === 0xd8) ext = 'jpg'
        else if (body.length > 6 && body[0] === 0x47 && body[1] === 0x49 && body[2] === 0x46) ext = 'gif'
        else if (body.length > 2 && body[0] === 0x42 && body[1] === 0x4d) ext = 'bmp'
        else if (body.length > 12 && body[4] === 0x66 && body[5] === 0x74 && body[6] === 0x79 && body[7] === 0x70) ext = 'webp'
        const name = 'img-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 6) + '.' + ext
        fs.writeFileSync(path.join(WORK, name), body)
        const url = 'http://' + (req.headers.host || '127.0.0.1') + '/work/' + name
        console.log('[upload] 图片接收：', name, body.length + 'B')
        send(res, 200, JSON.stringify({ [name]: url }), 'application/json')
      } catch (e) { send(res, 500, JSON.stringify({ error: String(e).slice(0, 200) }), 'application/json') }
      return
    }
    /* ── 工作区文件（/workfile/C:/... 或 /work/<name>） ── */
    if (p.startsWith('/workfile/C:/')) {
      const f = p.replace('/workfile/', '')
      if (fs.existsSync(f)) { send(res, 200, fs.readFileSync(f), MIME[path.extname(f)] || 'application/octet-stream') }
      else send(res, 404, 'nf: ' + f)
      return
    }
    if (p.startsWith('/work/')) {
      const f = path.join(WORK, p.replace('/work/', ''))
      if (fs.existsSync(f)) { send(res, 200, fs.readFileSync(f), MIME[path.extname(f)] || 'application/octet-stream') }
      else send(res, 404, 'nf')
      return
    }
    /* ── 常用样例直读（标题栏下拉 2026-10-01）：GET /office-samples/<name> ← server-data/office ── */
    if (p === '/office-samples' && req.method === 'GET') {
      /* ★ 批次17：样例清单（只列名/大小/时间，不吐字节）。
         用途：磁盘上的 .xlsx/.docx 从不自动进前端模板库（IndexedDB）——换浏览器/清缓存后，
         节点 config 里的模板 id 就成了指向不存在的东西，打开只剩「模板已不在模板库中」，
         而文件一直躺在磁盘上。前端拿这份清单做一次性自动索引（seedTemplatesFromDisk）。 */
      try {
        const files = fs
          .readdirSync(OO_SAMPLES_DIR)
          .filter((n) => /\.(xlsx|xlsm|docx|docm)$/i.test(n))
          .map((n) => {
            let st = null
            try {
              st = fs.statSync(path.join(OO_SAMPLES_DIR, n))
            } catch {
              /* 竞态删除：跳过该条，不让整批清单失败 */
            }
            return st && st.isFile()
              ? { name: n, ext: path.extname(n).slice(1).toLowerCase(), size: st.size, updatedAt: st.mtimeMs }
              : null
          })
          .filter(Boolean)
        send(res, 200, JSON.stringify({ dir: OO_SAMPLES_DIR, files }), 'application/json; charset=utf-8')
      } catch (e) {
        send(res, 500, JSON.stringify({ error: String((e && e.message) || e) }), 'application/json; charset=utf-8')
      }
      return
    }
    if (p.startsWith('/office-samples/')) {
      const f = path.normalize(path.join(OO_SAMPLES_DIR, p.replace('/office-samples/', '')))
      if (f.startsWith(OO_SAMPLES_DIR) && /\.(xlsx|docx)$/i.test(f) && fs.existsSync(f)) {
        send(res, 200, fs.readFileSync(f), MIME[path.extname(f)] || 'application/octet-stream')
      } else send(res, 404, 'nf')
      return
    }

    /* ── OO 启动预热（2026-10-09 用户需求：画布有 OO 节点 → 启动时后台焐热，首次打开免 30s 引擎冷加载）──
       干两件事：① ensureDocService（拉起 DS + 等 embedded converter started）；
       ② 顺序预读编辑器引擎大文件（sdk-all.bin / AllFonts.js / api.js）——只求字节过 OS 文件缓存，
       编辑器 iframe 首次加载时全部命中内存缓存。幂等：会话内第二次调用直接返回上次结果。 */
    if (p === '/warmup') {
      try {
        if (!__warmupResult) {
          if (__warmupBusy) { send(res, 202, JSON.stringify({ ok: false, busy: true }), 'application/json'); return }
          __warmupBusy = true
          const t0 = Date.now()
          try { const m = await import('./oo-docservice.mjs'); await m.ensureDocService({ quiet: true }) } catch { /* DS 拉起失败不阻塞预热 */ }
          const files = [
            'sdkjs/word/sdk-all.bin', 'sdkjs/word/sdk-all.js', 'sdkjs/cell/sdk-all.bin', 'sdkjs/cell/sdk-all.js',
            'fonts/AllFonts.js', 'web-apps/apps/api/documents/api.js',
          ]
          const warmed = []
          for (const f of files) {
            try {
              const fp = path.join(ASSETS, f)
              if (!fs.existsSync(fp)) continue
              await fs.promises.readFile(fp)   // 字节过 OS 缓存即弃
              warmed.push(f)
            } catch { /* 单文件失败不阻塞其余 */ }
          }
          __warmupResult = { ok: true, ms: Date.now() - t0, warmed, at: new Date().toISOString() }
          console.log('[oo-core] warmup 完成 ' + __warmupResult.ms + 'ms，预热文件 ' + warmed.length + ' 个')
          __warmupBusy = false
        }
        send(res, 200, JSON.stringify(__warmupResult), 'application/json')
      } catch (e) { __warmupBusy = false; send(res, 500, JSON.stringify({ ok: false, err: String(e).slice(0, 150) }), 'application/json') }
      return
    }

    /* ── POC 测试页（参数化实验：?src=lidu.editor.bin 走 strict 直喂内部格式） ── */
// ★ 2026-10-08 DS 转换器就绪门禁：DS「监听先开、embedded converter 后挂载」——
//   启动窗口期打开文档，转换任务进无人队列永久卡死。/poc 返回前确认 DS 日志出现
//   「embedded converter started」（模块级缓存：判定一次永久就绪，常态零开销；
//   最多等 30s 超时放行不永久卡）。编辑器 iframe init 需 4-5s，等待期被自然覆盖。
let __dsConvReady = false
// ★ /warmup 幂等标记（2026-10-09）：会话内只跑一次，结果缓存
let __warmupResult = null
let __warmupBusy = false
// ★ 2026-10-08 聪明解：/poc 放行门禁从「DS 转换器就绪」升级为「**该文档的
//   Editor.bin 已就位**」——后者是我们用 x2t（OO 官方离线内核，与 DS 底层同一个
//   exe，产物零差异）自己产的文件，完全可控、永不静默死亡。DS 整个转换层
//   （唯一不稳定组件）彻底退出关键路径。
const X2T = path.join(ROOT, 'vendors', 'oo-server', 'FileConverter', 'Bin', 'x2t.exe')
const dsEditorBin = (srcFile) =>
  path.join(ROOT, 'vendors', 'oo-server', 'App_Data', 'data',
    'oo-' + String(srcFile).replace(/[^\w.-]+/g, '-'), 'Editor.bin')

// 字体目录（docx 渲染必需——缺 AllFonts 会让 x2t 挂起不退出）
const FONT_DIR = path.join(ROOT, 'vendors', 'onlyoffice', 'fonts') + path.sep
const ALL_FONTS = path.join(FONT_DIR, 'AllFonts.js')

// 写 x2t XML 参数文件（输入/输出/字体目录）并返回其路径
function x2tParamsXml (srcAbs, binAbs) {
  const norm = (t) => t.split(String.fromCharCode(92)).join('/')
  const xml = `<?xml version="1.0" encoding="utf-8"?>
<Xml>
  <m_sInput>${norm(srcAbs)}</m_sInput>
  <m_sOutput>${norm(binAbs)}</m_sOutput>
  <m_sFontDir>${norm(FONT_DIR)}</m_sFontDir>
  <m_sAllFontsJS>${norm(ALL_FONTS)}</m_sAllFontsJS>
</Xml>`
  const xmlPath = binAbs + '.params.xml'
  fs.writeFileSync(xmlPath, xml)
  return xmlPath
}

// ★ 2026-10-09 01:30 移除 x2tConvertSync / spawnX2tConvert（上传旁路+门禁兜底已撤）：
//   x2t 独立进程启动固有 ~15s，DS 长驻转换 0.4s——关键路径上等 x2t 纯属倒退。
//   x2t 保留职责：/office-offline/save 的 bin→docx 反转、/convert 双向转换 API。

// ★ 2026-10-09 01:30 门禁降级为「缓存快查」：不再轮询等待、不再 x2t 同步补转。
//   数据支撑：DS lcid 修复后转换 0.4s（长驻进程），而 x2t 独立进程启动 15s——
//   旧门禁轮询 15s 超时反而让每次首开都白等；编辑器 open 命令本身会同步等 DS
//   转换完成（documentOpen ok）才渲染，这里等待毫无必要。
//   返回 true=缓存已在（DS 直接下发）；false=无缓存（放行，DS 转换 0.4s 兜底）。
function peekEditorBin (srcFile) {
  try {
    const bin = dsEditorBin(srcFile)
    return fs.existsSync(bin) && fs.statSync(bin).size > 1000
  } catch { return false }
}

async function waitDsConverterReady () {
  if (__dsConvReady) return
  const logFile = path.join(ROOT, 'vendors', 'oo-docservice-log.txt')
  for (let i = 0; i < 60; i++) {
    try {
      if (fs.readFileSync(logFile, 'utf8').includes('embedded converter started')) { __dsConvReady = true; return }
    } catch { /* 日志未生成——DS 还在起，继续等 */ }
    await new Promise((r) => setTimeout(r, 500))
  }
  __dsConvReady = true   // 超时放行：不永久卡，后续有 90s 自救兜底
}

    if (p === '/poc') {
      await waitDsConverterReady()   // DS 层门禁（常态已缓存，直通）
      const srcFile = u.searchParams.get('src') || 'lidu.docx'
      /* ★ 2026-10-23 用户指令「oo 打开 pdf 默认不是编辑状态」：docType=pdfviewer
         → OO PDF 查看模式（documentType 'pdf'、mode view、edit false）——
         轻量快速，OO 顶部自带「编辑 PDF」按钮可随时切换。 */
      const docType = u.searchParams.get('docType') || ''
      // ★ 2026-10-09 门禁 = 缓存快查（非阻塞）：有 Editor.bin 秒过；无缓存立即放行，
      //   DS embedded converter 转换仅 0.4s（lcid 已修复），open 命令自带同步等待。
      try { peekEditorBin(srcFile) } catch { /* ignore */ }
      const ext = path.extname(srcFile).replace('.', '')
      // ★ 主题跟随宿主应用（dark→theme-dark，其余→theme-classic-light）
      // ★★ 2026-10-07 19:3x 修正：这份 vendors 是**桌面构建**，main/index.html 内联脚本取
      //   window.RendererProcessVariable（CEF 壳注入）——纯浏览器壳无此变量，且 themes.json
      //   返回空列表 ⇒ theme-dark 定义不存在。实测宿主 dark → 编辑器引擎初始化直接崩死
      //   （editor=无 永远 loading，19:33 用户现场）。⇒ 强制 classic-light 保可用；
      //   深色编辑器支持需后续注入 RendererProcessVariable polyfill + dark 主题定义。
      const ooTheme = 'theme-classic-light'
      // ★ 标准服务器流：url 必须指原始 docx——DocService 内嵌转换器下载它并 x2t 构建 Editor.bin。
      //   早期「推 bin」模型把 url 重写成 .editor.bin，而 params.xml 模式 x2t 不认 .bin 输入 → exit 88
      const loadFile = srcFile
      // ★ 数据锁定（2026-10-07 用户指令「数据锁定按钮」）：?mode=view → 编辑器只读模式
      const ooMode = u.searchParams.get('mode') === 'view' ? 'view' : 'edit'
      // ★ 协作 key 由 srcFile 确定性派生（保存时按 key 找缓存 Editor.bin 反转回 docx）。
      //   注意：正则必须在这里（真实代码）算好，放进模板字面量时 \w 会被 JS 转义吞掉 → /[^w.-]/ 惨案
      const docKey = 'oo-' + srcFile.replace(/[^\w.-]+/g, '-')
      send(res, 200, `<!DOCTYPE html><html><head><meta charset="utf-8"><title>OO Offline POC</title>
<style>html,body{margin:0;height:100%;background:#0d1117}#mount{width:100%;height:100%}
#status{display:none}
/* ★ 2026-10-09 文档加载层：扁平海洋波浪+加载条（纯 CSS transform 动画，GPU 合成零重排）。
   此前页面背景是 OO 经典天蓝 #87CEEB，编辑器出来前整页裸蓝被用户当成"卡死"。
   documentReady 后加 .hide 淡出露出编辑器。波浪无缝漂移原理：波形宽 200%，平移 -50% 一周期。 */
#boot{position:fixed;inset:0;z-index:9999;background:#0d1117;display:flex;align-items:center;
      justify-content:center;transition:opacity .45s ease}
#boot.hide{opacity:0;pointer-events:none}
#boot .sea{position:absolute;left:0;right:0;bottom:0;height:34%;min-height:150px}
#boot .wave{position:absolute;bottom:0;left:0;width:200%;height:100%;
      animation:bo-drift linear infinite;will-change:transform}
#boot .w1{animation-duration:14s;opacity:.45}
#boot .w2{animation-duration:10s;opacity:.65}
#boot .w3{animation-duration:7s;opacity:.9}
@keyframes bo-drift{to{transform:translateX(-50%)}}
#boot .bar{position:relative;width:240px;height:4px;border-radius:2px;
      background:rgba(255,255,255,.09);overflow:hidden}
#boot .bar i{position:absolute;top:0;left:-30%;width:30%;height:100%;border-radius:2px;
      background:#58a6ff;animation:bo-slide 1.4s ease-in-out infinite}
@keyframes bo-slide{to{left:100%}}
@media (prefers-reduced-motion:reduce){#boot .wave,#boot .bar i{animation:none}}
</style></head>
<body><div id="status">loading editor… src=${srcFile}</div><div id="mount"></div>
<div id="boot">
  <div class="sea">
    <svg class="wave w1" viewBox="0 0 2400 120" preserveAspectRatio="none"><path d="M0,60 C150,110 350,10 600,60 C850,110 1050,10 1200,60 C1350,110 1550,10 1800,60 C2050,110 2250,10 2400,60 L2400,120 L0,120 Z" fill="#1f6feb"/></svg>
    <svg class="wave w2" viewBox="0 0 2400 120" preserveAspectRatio="none"><path d="M0,70 C200,20 400,120 600,70 C800,20 1000,120 1200,70 C1400,20 1600,120 1800,70 C2000,20 2200,120 2400,70 L2400,120 L0,120 Z" fill="#388bfd"/></svg>
    <svg class="wave w3" viewBox="0 0 2400 120" preserveAspectRatio="none"><path d="M0,80 C180,40 320,115 600,80 C880,45 1020,115 1200,80 C1380,40 1520,115 1800,80 C2080,45 2220,115 2400,80 L2400,120 L0,120 Z" fill="#58a6ff"/></svg>
  </div>
  <div class="bar"><i></i></div>
</div>
<script src="/oo-assets/web-apps/apps/api/documents/api.js"></script>
<script>
var S = document.getElementById('status')
/* ★ 2026-10-17 补：postToHost 只被调用、从未定义（全文件搜不到定义）。
 *   所有 try{ postToHost(...) }catch{} 因此一直静默失败 —— 「[提取v4]」这类
 *   保存链诊断日志从来没出现过，排查时看不到任何线索。
 *   与 __cellApiDiag 的 rep() 同一条通道：parent.postMessage(字符串/对象)。 */
function postToHost(payload) { try { parent.postMessage(payload, '*') } catch (e) {} }
/* ★ 2026-10-17 用户指令「每个步骤都监听、写进日志，不要再人肉贴日志」：
 *   直连 /office-log 落盘的统一埋点（同源 fetch，不经过 postMessage —— 
 *   中间任何一环断链都不会让日志凭空消失，这正是之前排查时最痛的点）。
 *   step 用固定编号（O*=打开链, S*=保存链, P*=插件链），读取脚本按编号判定缺环。 */
function __lg(step, detail) {
  try {
    fetch('/office-log', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      /* ★ 2026-10-17：带上会话 ID（宿主点 💾 时生成并随 __ooSaveReq 下发），
       *   服务端据此给整条链的每一行都打同一个 [sid +Nms] 前缀，
       *   多次保存混在一个日志文件里也能精确切开、算出每步耗时。 */
      body: JSON.stringify({ type: String(step), detail: String(detail == null ? '' : detail).slice(0, 700), sid: window.__ooSid || '' })
    })
  } catch (e) {}
}
// ★ 就绪/异常状态不再浮层显示（用户要求收进宿主诊断抽屉）：写入隐藏元素 + postMessage 上报，
//   宿主 OnlyOfficeNodeEditor 监听 __ooStatus 记入诊断日志。
function setSt(t) { S.textContent = t; try { parent.postMessage({ __ooStatus: t }, '*') } catch (e) {}
  __lg('O2', '引擎状态: ' + t)
  // ★ 文档就绪 → 淡出波浪加载层（所有就绪路径统一走这里，含轮询兜底）；淡出后移除 DOM
  if (/documentReady/.test(t)) { var b = document.getElementById('boot'); if (b && !b.classList.contains('hide')) { b.classList.add('hide'); setTimeout(function(){ if (b.parentNode) b.parentNode.removeChild(b) }, 600) } }
  // 自动重试场景：撤掉淡出，重新展示加载层
  if (/未就绪|重试/.test(t)) { var b2 = document.getElementById('boot'); if (b2) b2.classList.remove('hide') }
}
window.onerror = function (m, s, l, c) { setSt('ERR: ' + m + (l ? ' @' + String(s || '').split('/').pop() + ':' + l + ':' + (c || 0) : '')) }
var srcFile = ${JSON.stringify(srcFile)}
var loadFile = ${JSON.stringify(loadFile)}
// ★ 必须是绝对 URL：编辑器会把 document.url 原样放进 open 命令发给 DocService，
//   DS 在服务端下载它（new URL() 解析）——相对路径 /work/... 会 ERR_INVALID_URL →
//   任务失败 -81 → 编辑器弹 "Download failed"（2026-09-30 DS 日志实锤）。
//   用页面自身 origin（location.origin）拼绝对地址：DS 侧 node 下载走本机双栈 vite（::1 也可达）。
window.__docKey = ${JSON.stringify(docKey)}; // ★ 保存链 forcesave 的 key（宿主经此读取）
// ★ kind 以扩展名在本页直接推导（2026-10-10）：单一真相源，不再依赖宿主 postMessage
//   推送（冻结恢复场景推送会丢）；React 侧推送仍保留，仅作双保险覆盖
window.__ooKind = ${JSON.stringify(ext === 'xlsx' ? 'oosheet' : ext === 'pptx' ? 'ooppt' : 'ooword')};
// ★ 插件回传锚点（2026-10-10）：cbref 沿 parent 链爬到带 __cbRefHost 的窗口投递——
//   app 里窗口层级是 app→/poc→frameEditor→插件，投 top（=app 窗口）会丢失回执
window.__cbRefHost = true;
var docUrl = location.origin + '/work/' + loadFile
var cfg = {
  /* ★ 不传 type:'desktop'（2026-09-30 A/B 定案）：type=desktop 会让 web-apps 自带内部
   *  AscDesktopEditor → sdkjs Bic 走桌面打开流 → LocalStartOpen 等壳喂文档 → 永远
   *  "Loading document"。不传 type 走标准服务器流：document.url → DS 转换 → 秒开。 */
  width: '100%', height: '100%',
  /* ★ cbref 插件（批次 12·③ 数据联动）：见 editorConfig.plugins（2026-10-10 迁入，
     顶层位置会被本版 api.js 丢弃 → 插件从未加载）。 */
  document: {
    // ★ key 由 srcFile 确定性派生：保存时按 key 找 DocService 缓存的 Editor.bin 反转回原格式。
    //   动态 key 防了 Memory 旧错误回放，代价是保存要拿得到本 session 的 key —— 用 src 派生即可。
    fileType: ${JSON.stringify(ext)}, key: ${JSON.stringify(docKey)}, __docKey: ${JSON.stringify(docKey)}, title: srcFile,
    url: docUrl, permissions: { edit: ${docType === 'pdfviewer' ? false : true}, download: true, print: true }
  },
  documentType: ${docType === 'pdfviewer' ? JSON.stringify('pdf') : JSON.stringify(ext === 'xlsx' ? 'cell' : ext === 'pptx' ? 'slide' : 'word')},
  editorConfig: {
    mode: ${docType === 'pdfviewer' ? JSON.stringify('view') : JSON.stringify(ooMode)},
    lang: 'zh',
    // ★ 固定本机用户身份 + 关闭匿名署名询问（用户要求：后续画布联动时再接管真实身份）
    user: { id: 'local-user', name: '本机用户' },
    // ★ 不设 targetApp:'desktop'：设了它 web-apps 下载完 Editor.bin 后把打开动作交给壳
    //   （假壳不处理 → 永远 Loading）。去掉后走标准服务器流：sdkjs 自下载自打开。
    coEditing: { mode: 'strict' },
    customization: {
      compactHeader: true, hideRightSettings: true, uiTheme: ${JSON.stringify(ooTheme)},
      anonymous: { request: false, label: '本机用户' },
      /* ★ 隐藏原生 logo（2026-10-18）：OO 官方定制点（headerView branding.logo.visible===false
       *   → addClass('hidden')），但生效依赖 canBranding（license）——主力是注入层样式表
       *   __hideOoLogo()，这里同步声明保证升级/换 license 后语义一致。加回来 = 删掉本行 + __HIDE_OO_LOGO=false */
      logo: { visible: false },
      // ★ 拼写检查默认关（用户指令 2026-10-09「ooword拼写检查 先关闭」）：
      //   mode=false 初值关；change=true 保留高级设置里的手动开关（"先关闭"=可逆）
      features: { spellcheck: { mode: false, change: true } }
    },
    // ★ 插件挂载点 = editorConfig.plugins（2026-10-10 定案：本版 api.js schema 在这里；
    //   之前顶层 plugins 被丢弃 → cbref 插件从未加载 → 域外写入链最后一环断裂）
    plugins: {
      autostart: ['asc.{6B3F1A2C-4D5E-4F60-9A7B-2C3D4E5F6A7B}'],
      pluginsData: ['/oo-assets/plugins/cbref/config.json']
    }
  }
}
    /* ★ Ctrl+S（用户指令 2026-10-11）：/poc 文档级捕获（frameEditor 内的由 documentReady 后补挂） */
    /* ★ callCommand 签名适配（2026-10-12 终极根因）：旧版 4 参 (fn,calc,recalc,cb)，
     *   新版 3 参 (fn,props,cb)——旧签名调新 SDK 时回调被吞（落在 isRecalc 布尔位）。
     *   按函数声明的参数个数自动选择调用形态。 */
    /* ★★ 2026-10-30 callCommand 可靠包装（实测重开后回调**根本不触发**——no-cb 实锤）★★
     *  策略：fn 结果双写（返回值 + 引擎窗口全局 __ooCmdResult），
     *  cb 回调优先；800ms 未回则轮询全局变量（200ms×40），谁先到用谁。 */
    var __callCmd = function (eng, fn, cb) {
      var win = null
      try { var fe0 = document.querySelector('iframe[name="frameEditor"]'); win = fe0 && fe0.contentWindow } catch (eW0) {}
      var done = false
      var finish = function (v) { if (!done) { done = true; try { cb(v) } catch (eC0) {} } }
      /* 包装 fn：原 fn 执行后把返回值挂引擎窗口全局（fn 在引擎上下文执行，window 即引擎窗口） */
      var wrapped = function () {
        var __r = fn()
        try { window.__ooCmdResult = __r } catch (eG) {}
        return __r
      }
      var src = String(wrapped)
      /* 重建可序列化版本：wrapped 的闭包里引用了外部 fn——必须用 eval 方式组装 */
      var fnSrc = 'function(){ var __r = (' + String(fn) + ')(); try { window.__ooCmdResult = __r } catch (eG) {} return __r }'
      try { win && (win.__ooCmdResult = undefined) } catch (eC1) {}
      try {
        if (eng.callCommand.length >= 4) { eng.callCommand(eval('(' + fnSrc + ')'), false, false, finish) }
        else { eng.callCommand(eval('(' + fnSrc + ')'), { isCalculate: false, isRecalc: false }, finish) }
      } catch (e3a) {
        try { eng.callCommand(eval('(' + fnSrc + ')'), false, false, finish) } catch (e3b) { /* 不立即 cb(null)——等轮询兜底 */ }
      }
      /* 回调兜底轮询 */
      var polls = 0
      var timer = setInterval(function () {
        polls++
        if (done) { clearInterval(timer); return }
        var v = undefined
        try { v = win ? win.__ooCmdResult : undefined } catch (eP) {}
        if (v !== undefined) { clearInterval(timer); finish(v); return }
        if (polls > 40) { clearInterval(timer); finish(null) }
      }, 200)
    }
    var __doSaveExtract = function (saveMode) {
      try {
        __lg('S3', '提取开始 mode=' + (saveMode || 'node') + ' srcFile=' + srcFile)
        /* ★ 保存提取 v4（2026-10-12）：直接穿透 frameEditor 引擎——绕过插件转发链。
         *   引擎多候选探测；未提交值经目标帧 Asc.scope 传入，先提交进文档模型再读取。 */
        var fe = document.querySelector('iframe[name="frameEditor"]')
        var fw = fe && fe.contentWindow
        var eng = null, engWhere = ''
        try { if (fw && fw.Asc && fw.Asc.editor && fw.Asc.editor.callCommand) { eng = fw.Asc.editor; engWhere = 'Asc.editor' } } catch (e1) {}
        if (!eng) { try { if (fw && fw.editor && fw.editor.callCommand) { eng = fw.editor; engWhere = 'editor' } } catch (e2) {} }
        if (!eng) { try { if (fw && fw.wb && fw.wb.callCommand) { eng = fw.wb; engWhere = 'wb' } } catch (e3) {} }
        if (!eng) { try { if (window.editor && window.editor.callCommand) { eng = window.editor; engWhere = 'poc.editor' } } catch (e4) {} }
        try { postToHost({ __ooStatus: '[提取v4] 引擎=' + (engWhere || '未找到') }) } catch (eL0) {}
        __lg('S3a', '引擎=' + (engWhere || '未找到') + ' frameEditor=' + !!fe)
        if (!eng) {
          /* ★ 深度枚举：window 一级 + Asc 内一级，找所有 callCommand 持有者，报告到诊断 */
          const cands = []
          try {
            for (const k1 of Object.keys(fw || {})) {
              try { const v1 = fw[k1]; if (v1 && typeof v1.callCommand === "function") cands.push(k1 + ".callCommand") } catch (eC1) {}
            }
            if (fw && fw.Asc) {
              for (const k2 of Object.keys(fw.Asc)) {
                try { const v2 = fw.Asc[k2]; if (v2 && typeof v2.callCommand === "function") cands.push("Asc." + k2 + ".callCommand") } catch (eC2) {}
              }
            }
          } catch (eC0) {}
          try { postToHost({ __ooStatus: '[提取v4] ❌ 引擎未找到。候选：' + (cands.join(', ') || '（无）') }) } catch (eL3) {}
          __lg('S3-E', '❌ 引擎未找到，候选=' + (cands.join(', ') || '（无）'))
        }
        if (!eng) throw new Error('引擎不可达')
        var pendVal = ''
        try { var ai = fe.contentDocument.getElementById('area_id'); if (ai && ai.value) pendVal = String(ai.value) } catch (eA) {}
        try { postToHost({ __ooStatus: '[提取v4] 未提交值=' + (pendVal || '（无）') }) } catch (eL0b) {}
        __lg('S3c', '未提交值=' + (pendVal ? JSON.stringify(pendVal).slice(0, 60) : '（无）'))
        /* ★ 2026-10-17：编辑态值改走 window.__ooPendingVal（文档上下文同源读 parent），
         *   不再依赖 Asc.scope —— 实测 fw.Asc.scope 是 undefined，整条注入链一直是空的。 */
        try { window.__ooPendingVal = pendVal } catch (eWPv) { __lg('S3c-E', '❌ 写入 __ooPendingVal 失败: ' + String(eWPv).slice(0, 120)) }
        try { fw.Asc.scope.__pendVal = pendVal } catch (eScOld) { /* 旧通道，已不依赖 */ }
        /* ★ 2026-10-17 关键修复：srcFile 是闭包变量，进不了文档上下文（踩坑记录第 5 条）。
         *   readFn 里对它做正则 test 一执行就 ReferenceError ⇒ 整个提取回调拿到 null ⇒
         *   rows 恒 0 ⇒ 保存链空转。这就是「点保存没反应」在提取通道上的最终根因。
         *   必须经 Asc.scope 传入。（注意：本文件是模板字符串，注释里不能写反引号） */
        /* ★ 2026-10-17：scope 注入改为「可选兼容」—— 实测 fw.Asc.scope 是 undefined
         *   （Cannot set properties of undefined），而 readFn 现在已改用模板插值 + parent 取值，
         *   不再依赖它。这里保留是为了旧路径兼容，失败只记一条 info，不再当成错误。 */
        try {
          fw.Asc.scope.__srcFile = String(srcFile || '')
          __lg('S3b', 'scope 注入 OK（兼容通道）: __srcFile=' + JSON.stringify(String(srcFile || '').slice(0, 48)))
        } catch (eSc2) {
          __lg('S3b', 'scope 不可用（正常）: ' + String(eSc2).slice(0, 90) + ' → readFn 已改用模板插值，不受影响')
        }
        var readFn = function () {
          /* ★ 2026-10-17 修复（待办 3「编辑态未提交内容丢失」的真根因）：
           *   上面 1156 行把未提交值塞进了 Asc.scope.__pendVal，注释也写着「先提交进文档模型再读取」，
           *   但**函数体里从来没有读过它**（只写不读 = 死代码）⇒ 用户敲完数字不回车直接点 💾，
           *   读到的还是提交前的旧内容。
           *   ⚠ 不能无条件写：__pendVal 来自 #area_id，非编辑态时这个节点可能残留上次输入。
           *   故加两道闸——元素有实际尺寸（可见=正在编辑）+ 值非空。
           *   ⚠ 本函数体在**文档上下文**执行，只能用 Api/Asc.scope/document，取不到外层变量。 */
          /* ★★ 2026-10-17 根治「srcFile is not defined」★★
           *   这个函数体在文档上下文执行，只能用 Api / Asc / document / window。
           *   历史做法是把值塞进 Asc.scope 再从里面取，但实测 fw.Asc.scope 是 undefined
           *   （S3b-E 日志：Cannot set properties of undefined），整条注入链是空的。
           *   现在改成两条更硬的路：
           *   ① srcFile 是服务端已知的确定值 → 直接模板插值成字面量（零依赖）
           *   ② pendVal 是运行时才有的编辑态值 → /poc 页放在 window.__ooPendingVal，
           *      文档上下文经 window.parent 同源读取（读不到就空，不会抛） */
          var __pendV = ''
          try { __pendV = (window.parent && window.parent.__ooPendingVal) || '' } catch (eV0) {}
          if (__pendV) {
            var __areaEl = document.getElementById('area_id')
            var __vis = false
            try {
              if (__areaEl) { var __rc = __areaEl.getBoundingClientRect(); __vis = __rc.width > 1 && __rc.height > 1 }
            } catch (eV1) {}
            if (__vis) {
              try {
                var __selP = Api.GetSelection()
                if (__selP && typeof __selP.SetValue === 'function') __selP.SetValue(__pendV)
              } catch (ePV) {}
            }
          }
          /* ★ 2026-10-12 用户报「关联地址保存后丢失」：GetValue 只取纯值，公式被转静态。
           *   改为按单元格双读 GetFormula()+GetValue()——外部关联地址（公式）与数据都保留。 */
          var __srcFile = ${JSON.stringify(String(srcFile || ''))}
          /* 原来用正则判 docx 后缀，但模板字符串会把反斜杠吃掉，语义已经偏了。
           * 改用 slice 比较，无转义风险。 */
          if (__srcFile.slice(-5).toLowerCase() === '.docx') {
            try {
              var doc = Api.GetDocument()
              var n2 = doc.GetElementsCount()
              var rows2 = []
              for (var i3 = 0; i3 < n2 && i3 < 500; i3++) {
                var t3 = ''
                try { t3 = doc.GetElement(i3).GetText() } catch (eE3) {}
                rows2.push([t3])
              }
              return JSON.stringify({ rows: rows2 })
            } catch (eW) { return JSON.stringify({ rows: [] }) }
          }
          var ws = Api.GetActiveSheet()
          var used = ws.GetUsedRange()
          var rows = []
          if (used) {
            var r0 = used.GetRow(), c0 = used.GetCol()
            var nR = used.GetRowsCount(), nC = used.GetColumnsCount()
            if (nR > 500) nR = 500
            if (nC > 60) nC = 60
            for (var r = 0; r < nR; r++) {
              var line = []
              for (var c = 0; c < nC; c++) {
                var o = {}
                try {
                  var cell = ws.GetCells(r0 + r, c0 + c)
                  if (cell) {
                    /* ★ 2026-10-17 修复 #NAME? 垃圾：GetFormula() 对**纯文本格**也返回格内容，
                     *   原实现无差别记进 o.f ⇒ 宿主重建时把所有文本当公式写 ⇒ 满屏 #NAME?。
                     *   只有 '=' 开头的才是真公式（OO 的 GetFormula 对公式格返回 =...）。
                     *   ★ 同轮追加：老版本诊断会往 Z1 写 =diag.probe / =cbref.probe，
                     *     这些是垃圾不是用户内容，识别出来整格丢弃（否则会被一直带到新文件）。 */
                    var __probe = false
                    try {
                      var f4raw = String(cell.GetFormula() || '').trim()
                      if (/^=(diag|cbref)\.probe$/i.test(f4raw)) __probe = true
                      else if (f4raw.charAt(0) === '=') o.f = f4raw.slice(1)
                    } catch (eF4) {}
                    if (!__probe) {
                      try { var v4 = cell.GetValue(); if (v4 !== null && v4 !== undefined) o.v = (typeof v4 === "object") ? (v4.text !== undefined ? v4.text : String(v4)) : v4 } catch (eV4) {}
                    }
                  }
                } catch (eC4) {}
                line.push(o)
              }
              rows.push(line)
            }
          }
          /* ★★★ 2026-10-17 保存链「永远 rows=0」的真根因（无头端到端实锤）：
           *   callCommand 的回调只能拿到 **字符串/数字/数组**（cbref 注释里就写着这条铁律），
           *   这里却 return 裸对象 {rows} ⇒ 引擎序列化失败 ⇒ 回调收到 null ⇒ rows 恒 0。
           *   连带后果：①「兜底扫 A1:T100」当年判「无效」其实是同样丢在序列化上；
           *            ②宿主拿不到行 → 只能走 saveBack，DS 没产物就整条保存链空转。
           *   实锤路径：写后同引擎回读 used=true/nR=2 正常，同一引擎提取却 rows=0，
           *   两者唯一差异就是返回值形态（string vs object）。 */
          return JSON.stringify({ rows: rows })
        }
        /* ★★ 2026-10-17「一步一步查」最有价值的一步：readFn 静态自检 ★★
         *   这个函数体是在**文档上下文**里 eval 执行的，任何外层闭包变量都会当场
         *   ReferenceError，而报错堆栈只会指向 sdk-all-min.js 第 1549 行（毫无信息）。
         *   历史两次踩坑（srcFile / silent）都是这样静默失败的。
         *   这里在发出前扫一遍函数体源码：出现裸标识符就立刻记日志 + 明确点名，
         *   不用再等「回调恒 null → rows 恒 0」去反推。 */
        try {
          var __rfSrc = String(readFn)
          /* ★ 必须先剥掉注释再扫：这个函数体里有大量说明性注释
           *   （「srcFile 必须取 Asc.scope」之类），直接把整段源码拿来扫会误报。
           *   19:21 实测就误报了 srcFile/pendVal 两个，差点又被带偏。
           *   用 indexOf 手剥而不是正则 —— 本文件是模板字符串，正则会和转义打架。 */
          var __rfCode = __rfSrc
          try {
            var __c1 = __rfCode.indexOf(String.fromCharCode(47) + String.fromCharCode(42)) // /*
            while (__c1 >= 0) {
              var __c2 = __rfCode.indexOf(String.fromCharCode(42) + String.fromCharCode(47), __c1 + 2) // */
              if (__c2 < 0) break
              __rfCode = __rfCode.slice(0, __c1) + __rfCode.slice(__c2 + 2)
              __c1 = __rfCode.indexOf(String.fromCharCode(47) + String.fromCharCode(42))
            }
            /* ⚠ 这里不能直接 split 反斜杠 n：本文件是模板字符串，那个转义会被解析成**真换行**，
             *   生成出来的页面就变成 split('  换行  ') ⇒ 整块 script 语法错、页面全哑
             *   （19:24 实测：S2 之后一条日志都没有）。用字符码构造，绕开转义。 */
            __rfCode = __rfCode.split(String.fromCharCode(10)).map(function (ln) {
              var k = ln.indexOf(String.fromCharCode(47) + String.fromCharCode(47))
              return k >= 0 ? ln.slice(0, k) : ln
            }).join(String.fromCharCode(10))
          } catch (eCm) { __lg('S4-W', '注释剥离失败，按原码扫描（可能误报）: ' + String(eCm).slice(0, 80)) }
          /* ⚠ 列表里的名字必须拆成 'src' + 'File' 再拼：
           *   直接写完整字面量的话，String(readFn) 会把这段自检代码自己也扫进去，
           *   报一个「命中 srcFile」的假阳性（19:25 实测，白查一轮）。 */
          var __forbidden = ['src' + 'File', 'save' + 'Mode', 'pend' + 'Val', 'work' + 'Name', 'node' + 'Id']
          var __hit = []
          var __hitCtx = ''
          /* ★ 这里不能用带 \w 的正则：本文件是模板字符串，先吃一层反斜杠，
           *   JS 字符串字面量再吃一层 ⇒ 最终正则是 [^w.$-]，下划线被当成分隔符，
           *   __srcFile 里的 srcFile 就被误判成裸变量（19:29 实测，白查两轮）。
           *   改成纯字符比较，一个反斜杠都不写，彻底避开这个转义陷阱。
           *   ⚠⚠ 本段在模板字符串内，注释里绝对不能出现反引号（会提前闭合整页，
           *      oo-core 直接语法错，而且报错行在几百行之后，极难定位 —— 已两次踩坑）。 */
          var __isWordCh = function (c) {
            if (!c) return false
            return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c === '_' || c === '$'
          }
          var __findBare = function (code, key) {
            var at = 0
            while ((at = code.indexOf(key, at)) >= 0) {
              var prev = at > 0 ? code.charAt(at - 1) : ''
              var next = code.charAt(at + key.length)
              /* 前一个字符不能是标识符字符，也不能是 . 或 -（排除 fw.srcFile / a-srcFile） */
              var okPrev = !prev || (!__isWordCh(prev) && prev !== '.' && prev !== '-')
              var okNext = !next || !__isWordCh(next)
              if (okPrev && okNext) return at
              at += key.length
            }
            return -1
          }
          for (var __i = 0; __i < __forbidden.length; __i++) {
            var __k = __forbidden[__i]
            var __at = __findBare(__rfCode, __k)
            if (__at >= 0) {
              __hit.push(__k)
              if (!__hitCtx) {
                /* ⚠ 同样不能写 /[\r\n]+/：反斜杠会被模板字符串吃掉。
                 *   逐字符替换，保证日志里是单行。 */
                var __ctx = __rfCode.slice(Math.max(0, __at - 70), __at + 70)
                var __ctx1 = ''
                for (var __j = 0; __j < __ctx.length; __j++) {
                  var __ch = __ctx.charAt(__j)
                  __ctx1 += (__ch === String.fromCharCode(10) || __ch === String.fromCharCode(13)) ? ' / ' : __ch
                }
                __hitCtx = __ctx1
              }
            }
          }
          __lg('S4', 'readFn 自检: 长度=' + __rfSrc.length + ' 有效码=' + __rfCode.length + ' 闭包变量=' + (__hit.length ? '❌ 命中[' + __hit.join(',') + ']（会在文档上下文 ReferenceError）' : '✅ 无') + (__hitCtx ? '｜命中处…' + __hitCtx + '…' : ''))
          if (__hit.length) try { postToHost({ __ooStatus: '[提取v4] ⚠ readFn 含闭包变量: ' + __hit.join(',') }) } catch (eL4x) {}
        } catch (eRfC) { __lg('S4-E', 'readFn 自检异常: ' + String(eRfC).slice(0, 120)) }
        __lg('S5', 'callCommand 已发出（readFn 长度=' + String(readFn).length + ' 字符）')
        __callCmd(eng, readFn, function (res) {
          try {
            __lg('S6', '回调已收到 type=' + (typeof res) + ' 长度=' + (res == null ? 0 : String(typeof res === 'string' ? res : JSON.stringify(res)).length) + ' 内容前120=' + String(typeof res === 'string' ? res : JSON.stringify(res)).slice(0, 120))
            /* 兼容两种形态：新=JSON 字符串；旧（从未真正工作过）=对象 */
            var rows = []
            try {
              var o = (typeof res === 'string') ? JSON.parse(res) : res
              rows = (o && o.rows) || []
              /* ★ 2026-10-17：解析成功也要留痕——回调拿到 null / 空串 / 非预期结构，
               *   是「点了保存没反应」的典型形态，必须能一眼区分是解析失败还是本来就没内容。 */
              __lg('S6a', '解析 OK: 顶层键=' + (o && typeof o === 'object' ? Object.keys(o).join(',') : typeof o) + ' rows=' + rows.length)
            } catch (ePJ) {
              __lg('S6-E', '❌ 回传解析失败 type=' + (typeof res) + ' 内容=' + String(typeof res === 'string' ? res : JSON.stringify(res)).slice(0, 160) + ' err=' + String(ePJ).slice(0, 100))
            }
            try { postToHost({ __ooStatus: '[提取v4] 读取 rows=' + rows.length + ' 首行=' + String(JSON.stringify(rows[0] || [])).slice(0, 60) }) } catch (eL2) {}
            /* ★ 2026-10-17：rows 的「形状」比行数本身更能说明问题——
             *   行数 2 但非空格 0 = 提取到了 used range 却读不到值（API 用错/权限）；
             *   非空格 4 = 真读到内容。列数异常（>50）说明 used range 被撑爆。 */
            try {
              var __nCols = 0, __nVals = 0, __nFormulas = 0
              for (var __r = 0; __r < rows.length; __r++) {
                var __line = rows[__r]
                if (!__line || !__line.length) continue
                if (__line.length > __nCols) __nCols = __line.length
                for (var __c = 0; __c < __line.length; __c++) {
                  var __cell = __line[__c]
                  if (__cell == null || __cell === '') continue
                  if (typeof __cell === 'object') {
                    var __has = false
                    if (__cell.f != null && String(__cell.f) !== '') { __nFormulas++; __has = true }
                    if (__cell.v != null && String(__cell.v) !== '') __has = true
                    if (__cell.t != null && String(__cell.t) !== '') __has = true
                    if (__has) __nVals++
                  } else if (String(__cell) !== '') __nVals++
                }
              }
              __lg('S7', 'rows=' + rows.length + ' 行 × 最大 ' + __nCols + ' 列｜非空格=' + __nVals + ' 公式=' + __nFormulas + '｜首行=' + String(JSON.stringify(rows[0] || [])).slice(0, 140))
              if (rows.length > 0 && __nVals === 0) __lg('S7-W', '⚠ 有行但全空：used range 读到了、值读不到（检查 GetValue/权限）')
            } catch (eS7) {
              __lg('S7-E', '统计异常: ' + String(eS7).slice(0, 120) + ' rows=' + rows.length)
            }
            __lg('S8', '回传宿主 __ooSaveValues name=' + srcFile + ' rows=' + rows.length + ' mode=' + (saveMode || 'node'))
            parent.postMessage({ __ooSaveValues: { name: ${JSON.stringify(srcFile)}, rows: rows, mode: saveMode || 'node' } }, '*')
          } catch (eP) { __lg('S8-E', '❌ 回传失败: ' + String(eP).slice(0, 140) + ' | ' + String(eP && eP.stack ? eP.stack : '').slice(0, 200)) }
        })
      } catch (eR0) { try { parent.postMessage({ __ooSaveReq: 1 }, '*') } catch (eR1) {} }
    }
    /* ★ 2026-10-17 脏标记统一出口（用户指令「做是否修改了的判断 然后提示保存」）：
     *   此前 window.__ooDocModified 只在本地赋值，从不通知宿主 —— 宿主关闭时只能
     *   发 __ooIsModifiedAsk 问、1.5s 超时后返回 false ⇒ 「改了内容，关掉没提示」。
     *   现在改成唯一写入口：状态翻转时主动 postMessage 给宿主，宿主据此亮「未保存」。 */
    window.__setOoDirty = function (v, why) {
      try {
        var nv = !!v
        var nw = String(why || '')
        window.__ooDirtyWhy = nw
        if (window.__ooDocModified === nv) return
        window.__ooDocModified = nv
        try { parent.postMessage({ __ooDirty: nv, __ooDirtyWhy: nw }, '*') } catch (eP0) {}
        try { if (typeof __lg === 'function') __lg('DIRTY', (nv ? '→ 有未保存修改' : '→ 已无未保存修改') + '（来源=' + nw + '）') } catch (eP1) {}
      } catch (eD) {}
    }
    /* ★★ 2026-10-17 用户指令「从现在开始，我对软件的每个点击等操作、键盘输入、软件的反馈，
     *   都放入日志」★★
     *   编辑器本体在 frameEditor 这层内层 iframe 里，**宿主的 document 监听照不到**
     *   （点击发生在 sdkjs 的 canvas 上，事件根本不冒泡到 app 窗口）。
     *   这里同源直取 contentDocument 补挂捕获监听，采集后 postMessage 给宿主，
     *   由宿主转发 audit() 落到 logs/audit-YYYY-MM-DD.log —— 与宿主自己的操作日志同一条时间线。
     *   ⚠ 本文件整体是模板字符串：这段代码里不能出现反引号与美元+花括号。 */
    var __oa = function (kind, action, detail) {
      try { parent.postMessage({ __ooAudit: { kind: String(kind), action: String(action), detail: String(detail == null ? '' : detail).slice(0, 600) } }, '*') } catch (eOA0) {}
    }
    var __oaCls = function (el) {
      try {
        if (typeof el.className === 'string') return el.className
        return (el.getAttribute && el.getAttribute('class')) || ''
      } catch (e) { return '' }
    }
    var __oaDesc = function (el) {
      try {
        if (!el || !el.tagName) return '-'
        var s = String(el.tagName).toLowerCase()
        if (el.id) s += '#' + el.id
        var cl = String(__oaCls(el)).trim().split(/\s+/).filter(Boolean).slice(0, 3).join('.')
        if (cl) s += '.' + cl
        var lb = ''
        var ks = ['data-audit', 'data-testid', 'aria-label', 'title', 'name']
        for (var i = 0; i < ks.length; i++) {
          var v = el.getAttribute && el.getAttribute(ks[i])
          if (v && String(v).trim()) { lb = String(v).trim(); break }
        }
        if (!lb) lb = String(el.innerText || el.textContent || '').replace(/\s+/g, ' ').trim()
        if (lb) s += ' "' + lb.slice(0, 40) + '"'
        return s
      } catch (e) { return '?' }
    }
    var __oaVal = function (el) {
      try {
        if (!el) return ''
        var tag = String(el.tagName || '').toLowerCase()
        if (tag === 'input' || tag === 'textarea') {
          var tp = String(el.type || 'text').toLowerCase()
          if (tp === 'password') return '••••(len=' + String(el.value || '').length + ')'
          return String(el.value == null ? '' : el.value).slice(0, 120)
        }
        if (tag === 'select') {
          var out = []
          for (var i = 0; i < el.selectedOptions.length; i++) {
            var o = el.selectedOptions[i]
            out.push(o.text || o.value)
          }
          return out.join(',')
        }
        if (el.isContentEditable) return String(el.innerText || '').slice(0, 120)
        return ''
      } catch (e) { return '' }
    }
    var __oaBind = function (doc, where) {
      try {
        if (!doc || doc.__ooAuditBound) return
        doc.__ooAuditBound = true
        var opt = { capture: true, passive: true }
        var optA = { capture: true }
        var lpAt = 0, lpTgt = null
        var inTimer = null, inN = 0, inEl = null
        var whTimer = null, whN = 0, whDY = 0, whEl = null
        var wtag = '(' + where + ')'
        doc.addEventListener('pointerdown', function (e) {
          lpAt = Date.now(); lpTgt = e.target
          var btn = ['左键', '中键', '右键', '侧键1', '侧键2'][e.button] || ('button' + e.button)
          __oa('点击', 'pointerdown' + wtag, __oaDesc(e.target) + ' ' + btn + ' 坐标=(' + Math.round(e.clientX) + ',' + Math.round(e.clientY) + ')')
        }, opt)
        /* click 只补记「没有 pointerdown 前导」的（键盘 Enter/Space 激活、脚本触发） */
        doc.addEventListener('click', function (e) {
          if (lpAt && Date.now() - lpAt < 700 && lpTgt === e.target) return
          __oa('点击', 'click' + wtag, __oaDesc(e.target))
        }, opt)
        doc.addEventListener('dblclick', function (e) { __oa('点击', 'dblclick' + wtag, __oaDesc(e.target)) }, opt)
        doc.addEventListener('contextmenu', function (e) { __oa('点击', '右键菜单' + wtag, __oaDesc(e.target)) }, opt)
        /* 键盘：长按收敛（e.repeat 直接丢），输入法组字过程不记 */
        doc.addEventListener('keydown', function (e) {
          if (e.isComposing || e.key === 'Process' || e.repeat) return
          var mods = []
          if (e.ctrlKey) mods.push('Ctrl')
          if (e.altKey) mods.push('Alt')
          if (e.shiftKey) mods.push('Shift')
          if (e.metaKey) mods.push('Meta')
          var k = e.key === ' ' ? 'Space' : e.key
          __oa('键盘', 'keydown' + wtag, (mods.length ? mods.join('+') + '+' : '') + k + ' 焦点=' + __oaDesc(e.target))
        }, optA)
        /* 输入：300ms 合并，只记「改了哪个框、改成什么、改了几次」——逐字符记会把日志撑爆 */
        doc.addEventListener('input', function (e) {
          inN++; inEl = e.target
          if (inTimer) clearTimeout(inTimer)
          inTimer = setTimeout(function () {
            var n = inN, el2 = inEl
            inN = 0; inTimer = null
            var v = __oaVal(el2)
            __oa('输入', 'value' + wtag, __oaDesc(el2) + (n > 1 ? ' ×' + n : '') + (v ? ' 值="' + v + '"' : ''))
          }, 300)
        }, opt)
        /* 滚轮：500ms 聚合（编辑器画布滚动极其高频） */
        doc.addEventListener('wheel', function (e) {
          whN++; whDY += e.deltaY; whEl = e.target
          if (whTimer) clearTimeout(whTimer)
          whTimer = setTimeout(function () {
            var n = whN, dy = whDY, el2 = whEl
            whN = 0; whDY = 0; whTimer = null
            __oa('滚轮', 'wheel' + wtag, __oaDesc(el2) + ' ×' + n + ' ΔY=' + Math.round(dy))
          }, 500)
        }, opt)
        doc.addEventListener('focusin', function (e) {
          if (e.target && e.target !== doc.body) __oa('焦点', 'focus' + wtag, __oaDesc(e.target))
        }, opt)
        __oa('系统', '编辑器审计已挂载', '作用域=' + where + ' 文档=' + String(doc.location && doc.location.pathname || ''))
      } catch (eOB) {
        try { __lg('AUDIT-E', '编辑器审计挂载失败: ' + String(eOB).slice(0, 120)) } catch (eOB2) {}
      }
    }
    __oaBind(document, 'poc')
    /* ★ 隐藏 OO 原生 logo（2026-10-18 用户指令「把原生的 ooexcel 的 logo 隐藏了（我可能会随时加回来）」）：
     *   实测 logo = 工具栏左上角 section.logo > #header-logo（86x20，位于 .box-tabs .extra.left），
     *   http://…/common/main/resources/img/header/header-logo_s.svg。
     *   三层手段：
     *     ① cfg.customization.logo.visible=false —— OO 官方定制点（headerView 消费
     *        branding.logo.visible===false → addClass('hidden')），但依赖 canBranding（license）；
     *     ② 注入样式表压死（不依赖 license，随时生效）；
     *     ③ 幂等 + 定时补挂 —— frameEditor 的 contentDocument 在保存后重开时会被替换，
     *        与 __bindSaveKeys 同一套防替换纪律。
     *   ★ 加回来：把 __HIDE_OO_LOGO 改回 false 即可，无需动其他代码。 */
    var __HIDE_OO_LOGO = true
    var __hideOoLogo = function () {
      try {
        if (!__HIDE_OO_LOGO) return
        var fL = document.querySelector('iframe[name="frameEditor"]')
        var dkL = fL && fL.contentDocument
        if (!dkL || !dkL.head) return
        var stL = dkL.getElementById('ext-hide-logo-style')
        if (!stL) {
          stL = dkL.createElement('style')
          stL.id = 'ext-hide-logo-style'
          dkL.head.appendChild(stL)
        }
        var cssL = 'section.logo{display:none!important}#header-logo{display:none!important}'
        if (stL.textContent !== cssL) stL.textContent = cssL
      } catch (eL) {}
    }
    /* ★ keydown 幂等绑定（见 onDocumentReady 注释）：同一 document 只绑一次 */
    var __bindSaveKeys = function () {
      try {
        var feK = document.querySelector('iframe[name="frameEditor"]')
        var dk = feK && feK.contentDocument
        if (dk && !dk.__ooSaveKeyBound) {
          dk.__ooSaveKeyBound = true
          dk.addEventListener('keydown', __saveKey, true)
        }
        /* ★ 2026-10-17：编辑器内部的点击/键盘审计也在这里补挂（同一份幂等标记语义，
         *   frameEditor 的 contentDocument 在保存后重开时会被替换，必须每次 ready 重新挂） */
        if (dk) __oaBind(dk, 'frameEditor')
      } catch (eBK) {}
    }
    var __saveKey = function (ev) {
      /* ★ 2026-10-19 修「新导入的文件没修改也提示保存」：导航键不算修改——
       *   箭头/Home/End/翻页/Tab/Esc/F1-F12 只是移动光标或调面板，不写内容；
       *   旧逻辑把任何非修饰键都标脏 ⇒ 导入后按方向键看一眼就「未保存」。 */
      try {
        var k0 = (ev.key || '').toLowerCase()
        var __nav = k0 === 'arrowleft' || k0 === 'arrowright' || k0 === 'arrowup' || k0 === 'arrowdown' ||
          k0 === 'home' || k0 === 'end' || k0 === 'pageup' || k0 === 'pagedown' ||
          k0 === 'escape' || k0 === 'tab' || k0 === 'enter' || (k0.length === 2 && k0.charAt(0) === 'f' && !isNaN(parseFloat(k0.charAt(1))))
        if (ev.key && !ev.ctrlKey && !ev.metaKey && !ev.altKey && !__nav) window.__setOoDirty(true, '键盘输入')
      } catch (eM0) {}
      try {
        var k = (ev.key || '').toLowerCase()
        if ((ev.ctrlKey || ev.metaKey) && k === 's') {
          ev.preventDefault(); ev.stopPropagation()
          /* ★ 先提交打开中的单元格编辑（用户输入的最后一段还在格子编辑器里，未进文档模型——
           *   直接 asc_Save 存的是提交前内容 =「123 消失」的根因） */
          try { if (window.editor && typeof window.editor.asc_closeCellEditor === 'function') window.editor.asc_closeCellEditor(false) } catch (eC) {}
          setTimeout(function () { try { __doSaveExtract() } catch (e0) {} }, 800)
        }
      } catch (e1) {}
    }
    document.addEventListener('keydown', __saveKey, true)
var _err = false
var _ready = false
var ed = new DocsAPI.DocEditor('mount', {
  width: '100%', height: '100%',
  document: cfg.document, documentType: cfg.documentType, editorConfig: cfg.editorConfig,
  events: {
    /* ★ 2026-10-12 官方脏跟踪（用户指令搜索全网）：修改时 data:true，提交时 data:false */
    onDocumentStateChange: function (e) {
      if (!window.__ooDocArmed) return // ★ 未武装：初始化期误报，丢弃
      /* ⚠ 只信 data === true，绝不信 false（2026-10-17 实测）：
       *   sdkjs 的 false 语义是「改动已同步到协作端」，而我们走的是自研 socket
       *   （没有真实 DocService），它照样会在 2 秒后自行发 false —— 实测 20:10:42
       *   标脏、20:10:44 就被这条清掉，导致「改了内容关闭时不提示」。
       *   脏标志只能由我们自己确认的保存完成（__ooSavedAck）来清。 */
      if (e && e.data) { try { window.__setOoDirty(true, 'docStateChange') } catch (eS0) {} }
      try { parent.postMessage({ __ooStatus: '[doc] ' + (e && e.data ? '有未保存修改' : '引擎报已同步（不据此清脏）') }, '*') } catch (eX) {}
    },
    onDocumentReady: function () {
      _ready = true; setSt('documentReady ✅');
      /* ★ 2026-10-23 用户指令「oopdf 打开要调整至适应比例」（二版）：检测按 src 的 .pdf
         后缀（fileType=pdf 只在 frameEditor src 里，/poc 页自身没有）；zoomFitToPage
         轮询重试（frameEditor 内 editor 初始化晚于 documentReady，单次 120ms 太早）——
         每 500ms 一次最多 12 次，成功即停；窗口 resize 防抖重适应。 */
      if (/.pdf/i.test(location.search)) {
        var __fitN = 0
        var __fitApi = function () {
          return window.editor || (window.Asc && window.Asc.editor) || (function () {
            var f = document.querySelector('iframe[name="frameEditor"]')
            return f && f.contentWindow ? (f.contentWindow.editor || (f.contentWindow.Asc && f.contentWindow.Asc.editor)) : null
          })()
        }
        var __fitOnce = function () {
          var a = __fitApi()
          if (a && typeof a.zoomFitToPage === 'function') { try { a.zoomFitToPage() } catch (eZ1) {} ; return true }
          return false
        }
        if (!__fitOnce()) {
          var __fitI = setInterval(function () {
            __fitN++
            if (__fitOnce() || __fitN >= 12) clearInterval(__fitI)
          }, 500)
        }
        var __fitT
        /* ★ 2026-10-23 防抖 200→800：缩放重适应=全页重光栅化（OO pdf 内核特性，
           单次即卡）——拖窗口边缘连续触发会卡上加卡，放宽到停止后 0.8s 执行一次 */
        window.addEventListener('resize', function () {
          clearTimeout(__fitT)
          __fitT = setTimeout(function () { __fitOnce() }, 800)
        })
      }
      window.__setOoDirty(false, 'ready') // ★ 基线校准：打开完成≠已修改（用户指令：没修改过不提示）
      window.__ooDocArmed = false // ★ 武装延迟：打开后 2 秒内的状态变更视为初始化误报，丢弃
      /* ★ 2026-10-18 深夜改版：武装从「2s 定时」改为「首次真实交互」。
       *   用户实测「打开什么都没改，关闭却提示保存」——引擎加载后会自行重算/规范化
       *   （旧文件还带 fullCalcOnLoad），2s 一过 stateChange(true) 就把脏标记点亮。
       *   改为：keydown / mousedown / wheel 任一真实输入事件才武装——引擎自导的
       *   状态变更没有输入事件伴随，永远 unarmed；真人编辑必然先有输入 ✓。 */
      var __armOnce = function () {
        try {
          window.__ooDocArmed = true
          document.removeEventListener('keydown', __armOnce, true)
          document.removeEventListener('mousedown', __armOnce, true)
          document.removeEventListener('wheel', __armOnce, true)
          var fA = document.querySelector('iframe[name="frameEditor"]') || document.querySelector('iframe')
          var dA = fA && fA.contentDocument
          if (dA) {
            dA.removeEventListener('keydown', __armOnce, true)
            dA.removeEventListener('mousedown', __armOnce, true)
            dA.removeEventListener('wheel', __armOnce, true)
          }
          try { console.log('[cboo] dirty armed (first user input)') } catch (eAG) {}
        } catch (eA0) {}
      }
      window.__ooDocArm = __armOnce
      document.addEventListener('keydown', __armOnce, true)
      document.addEventListener('mousedown', __armOnce, true)
      document.addEventListener('wheel', __armOnce, true)
      try {
        var fArm = document.querySelector('iframe[name="frameEditor"]') || document.querySelector('iframe')
        var dArm = fArm && fArm.contentDocument
        if (dArm) {
          dArm.addEventListener('keydown', __armOnce, true)
          dArm.addEventListener('mousedown', __armOnce, true)
          dArm.addEventListener('wheel', __armOnce, true)
        }
      } catch (eA1) {}
      /* 打开完成后 3s 做一次兜底基线校准：若期间没有任何真实输入且引擎报未修改，
       * 清一次脏标记（防御个别版本在 ready 之前就误报）。 */
      setTimeout(function () {
        try {
          if (!window.__ooDocArmed && window.__setOoDirty) window.__setOoDirty(false, 'settle-baseline')
        } catch (eA2) {}
      }, 3000)
      /* ★ Ctrl+S 保存（用户指令 2026-10-11）：编辑器内 Ctrl+S → 宿主保存回写节点
       *   ★ 2026-10-17 保存后自动重开文档时 frameEditor 的 contentDocument 会被替换，
       *     只绑一次的监听会甩在旧文档上 —— 之后真人键盘输入就再也标不了脏
       *     （实测：20:46:09 ready 后用户改内容无 DIRTY 记录，手动重开节点后才恢复）。
       *     改成幂等重绑：ready 时 + 1.2s + 3s 各试一次，已在目标 document 上绑过就跳过。 */
      try { __bindSaveKeys() } catch (eS2) {}
      try { setTimeout(function () { try { __bindSaveKeys() } catch (eB1) {} }, 1200) } catch (eB0) {}
      try { setTimeout(function () { try { __bindSaveKeys() } catch (eB3) {} }, 3000) } catch (eB2) {}
      /* ★ 隐藏 OO logo：ready 即挂 + 0.8s/2s 补挂（工具栏 DOM 可能后建 / 文档重开替换） */
      try { __hideOoLogo() } catch (eL0) {}
      try { setTimeout(function () { try { __hideOoLogo() } catch (eL1) {} }, 800) } catch (eL2) {}
      try { setTimeout(function () { try { __hideOoLogo() } catch (eL3) {} }, 2000) } catch (eL4) {}
      // 「刷新域值」按钮（2026-10-06）：写回由宿主执行，这里只负责入口与结果展示。
      // 延后挂载：文档就绪时工具栏 DOM 可能还没建好。
      try { setTimeout(function () { try { __mountRefreshBtn() } catch (e) {} }, 800) } catch (e) {}
      /* ★ 2026-10-18 晚：外部引用更新设置入口挪到**宿主标题栏方形按钮**（用户指令
       *   「引用更新设置 放在标题栏 方形按钮上」）——/poc 顶部按钮随 tab 切换会丢且难找。
       *   这里不再挂工具栏按钮；弹窗（__ooRefAsk）与提示逻辑全部保留。 */
      // ★ 打开完成后补挂最小壳对象：桌面构建多条路径无条件访问 window.AscDesktopEditor.*
      //   （协作改链 zXa→LocalFileSaveChanges、SetDocumentName 等），缺失即抛
      //   "An error occurred during the work with the document"（2026-09-30 19:56 实锤）。
      //   启动阶段绝不能存在（Bic 见它就切桌面流 → LocalStartOpen 死锁），故此刻再注入纯 no-op 壳。
      try {
        var f = document.querySelector('iframe')
        if (f && f.contentWindow && !f.contentWindow.AscDesktopEditor) {
          // ★ Proxy 万能壳：桌面构建散布大量无守卫的 AscDesktopEditor.* 调用
          //   （LocalFileSaveChanges、CheckNeedWheel……逐个补会无限挤牙膏），
          //   未知属性一律返回 no-op 函数；有返回值语义的方法白名单特判。
          var __shimTarget = {}
          f.contentWindow.AscDesktopEditor = new Proxy(__shimTarget, {
            get: function (t, k) {
              if (k in t) return t[k]
              switch (k) {
                case 'IsLocalFile': return function () { return false }
                case 'LocalFileGetSaved': return function () { return false }
                case 'GetVersion': return function () { return '9.4.0-web-poc' }
                // ★ 2026-10-10（无头实测铁证）：Proxy 默认 no-op 让 isSupportPlugins()
                //   返回 undefined ⇒ sdkjs 插件服务 Mo() 判定「壳不支持插件」⇒
                //   asc_pluginRun 静默跳过 iframe 创建 ⇒ cbref 插件永远无法启动。
                //   必须声明支持，插件体系（autostart/asc_pluginRun）才能活。
                case 'isSupportPlugins': return function () { return true }
                case 'GetInstallPlugins': return function () { return JSON.stringify([{ url: '/oo-assets/plugins/', pluginsData: [] }, { url: '/oo-assets/plugins/', pluginsData: [] }]) }
                case 'GetDictionariesPath': return function () { return '/oo-assets/dictionaries/' }
                case 'CryptoMode': return function () { return 0 }
                case 'OpenFilenameDialog': return function (k2, s2, p2, cb) { if (cb) cb(null) }
                default: return function () {}
              }
            },
            set: function (t, k, v) { t[k] = v; return true }
          })
        }
      } catch (e) {}
      // ★ 2026-10-10 确保插件启动（域外引用最后一环）：本桌面构建 autostart 队列
      //   时序不稳（配置晚到/集合晚建都会空转），documentReady 后显式拉起一次。
      //   壳已在上面声明 isSupportPlugins=true，asc_pluginRun 不再被闸门拦截。
      try {
        /* ★ 2026-10-10 自愈轮询版（原一次性 setTimeout(1200) 实测时序脆弱：
         *   引擎没完全就绪时注册/拉起都静默空转，成败全凭运气）。
         *   每秒自检，两件事各自完成才收工，30s 上限：
         *   ① 实时光标跟踪：光标只在面板打开时问一次 → 之后点任何格都写回 A1。
         *      订阅名称框事件——点选即上报，__ooActiveCell 实时刷新 + __ooCursor 推宿主。
         *   ② 确保 cbref 插件已拉起（autostart 队列时序兜底）。 */
        var __cbCursorRegistered = false
        var __cbEnsureTicks = 0
        window.__cbDiag = function () {
          try {
            var fE = document.querySelector('iframe')
            var w2 = fE && fE.contentWindow
            return {
              ticks: __cbEnsureTicks,
              registered: __cbCursorRegistered,
              hasFE: !!fE,
              api2: w2 ? (typeof (w2.editor || (w2.Asc && w2.Asc.editor))) : 'no-frame',
              iframeCount: (function () { try { return fE.contentDocument.querySelectorAll('iframe').length } catch (e) { return 'err' } })()
            }
          } catch (e) { return { err: String(e).slice(0, 80) } }
        }
        var __cbEnsureTimer = setInterval(function () {
          __cbEnsureTicks++
          try {
            var fE = document.querySelector('iframe')
            if (!fE || !fE.contentDocument) return
            var w2 = fE.contentWindow
            var api2 = (w2.editor || (w2.Asc && w2.Asc.editor))
            if (!api2) return
            // ① 光标事件注册（幂等：成功一次即停）
            if (!__cbCursorRegistered && typeof api2.asc_registerCallback === 'function') {
              try {
                api2.asc_registerCallback('asc_onSelectionNameChanged', function (t) {
                  try {
                    var name = (typeof t === 'string') ? t : (t && t.asc_getName ? t.asc_getName() : '')
                    // ★ 2026-10-10 域实锤：本回调被 sdkjs 在【引擎域】执行——
                    //   这里绝不能写 window.__ooActiveCell（会写进 frameEditor 的 window，
                    //   /poc 读不到）。只发消息，缓存由 /poc 域统一落账。
                    try { parent.postMessage({ __ooCursor: name || '(空)' }, '*') } catch (e) {}
                  } catch (e) {
                    try { parent.postMessage({ __ooStatus: '[poc] 名称框事件回调抛错 ' + String(e).slice(0, 60) }, '*') } catch (e2) {}
                  }
                })
                __cbCursorRegistered = true
                try { parent.postMessage({ __ooStatus: '[poc] 光标事件注册成功（asc_onSelectionNameChanged）' }, '*') } catch (e3) {}
              } catch (e) {
                try { parent.postMessage({ __ooStatus: '[poc] 光标事件注册抛错 ' + String(e).slice(0, 60) }, '*') } catch (e5) {}
              }
            }
            // ② 插件拉起（以 iframe 实际出现为完成标志；未出现则每秒重试 asc_pluginRun）
            var has = false
            fE.contentDocument.querySelectorAll('iframe').forEach(function (x) { if ((x.getAttribute('src') || '').indexOf('cbref') >= 0) has = true })
            if (!has && api2 && typeof api2.asc_pluginRun === 'function') {
              api2.asc_pluginRun('asc.{6B3F1A2C-4D5E-4F60-9A7B-2C3D4E5F6A7B}', 0, '')
              if (__cbEnsureTicks === 1) try { parent.postMessage({ __ooStatus: '[poc] 已显式拉起 cbref 插件（asc_pluginRun，未现则重试）' }, '*') } catch (e2) {}
            }
            if ((__cbCursorRegistered && has) || __cbEnsureTicks > 30) clearInterval(__cbEnsureTimer)
          } catch (e1) {}
        }, 1000)
      } catch (e0) {}
    },
    onError: function (e) { _err = true; setSt('onError: ' + JSON.stringify(e).slice(0, 200)) },
    onWarning: function (e) { if (!_err && !_ready) setSt('onWarning: ' + JSON.stringify(e).slice(0, 200)) }
  }
})
  // ★ 工具栏「批注」旁并列注入「参数引用」按钮：克隆批注按钮（样式天然一致），点击向宿主
  //   发 __ooToggleRef（宿主切参数引用面板 = 批注面板同位展开）；宿主回 __ooRefState 做高亮。
  //   定时守卫：工具栏重建后自动补插。
  var __injectRefBtn = function () {
    try {
      var doc = document
      var edoc = doc.querySelector('iframe') && doc.querySelector('iframe').contentDocument
      if (!edoc) return
      if (edoc.getElementById('ext-btn-ooref')) return
      var src = edoc.getElementById('tlbtn-addcomment-0')
      if (!src) return
      var btn = src.cloneNode(true) // 深克隆：连图标/文字结构一起复制，视觉与批注一致
      btn.id = 'ext-btn-ooref'
      btn.removeAttribute('aria-label')
      btn.title = '打开/关闭参数引用'
      var ic = btn.querySelector('.inner-box-icon')
      if (ic) ic.remove() // 去掉批注图标（避免重样）
      var cap = btn.querySelector('.caption')
      if (cap) cap.textContent = '参数引用'
      else btn.textContent = '参数引用'
      btn.style.minWidth = '52px'
      try {
        var vis2 = window.__ooMax === true
        btn.style.display = vis2 ? '' : 'none'
        window.__extApplyMaxVis = function (v) { btn.style.display = v ? '' : 'none' }
      } catch (e) {}
      btn.onclick = function () { try { __refBtnAction() } catch (e) {} }
      src.parentNode.insertBefore(btn, src.nextSibling)
      window.__extRefBtn = btn
      // ★ 左侧图标条：批注图标下方并列注入「引」图标（与批注同款按钮结构）
      var sb = edoc.getElementById('left-btn-comments')
      if (sb && !edoc.getElementById('ext-btn-left-ooref')) {
        var sbtn = sb.cloneNode(true)
        sbtn.id = 'ext-btn-left-ooref'
        sbtn.classList.remove('active', 'notify')
        sbtn.title = '参数引用'
        var ic = sbtn.querySelector('i') || sbtn.querySelector('.icon')
        if (ic) {
          ic.className = 'icon toolbar__icon'
          ic.textContent = '引'
          ic.style.fontSize = '13px'
          ic.style.fontWeight = 'bold'
          ic.style.fontStyle = 'normal'
          ic.style.lineHeight = '20px'
          ic.style.display = 'block'
          ic.style.textAlign = 'center'
        }
        var __applyMaxVis = function () {
          try {
            var vis = window.__ooMax === true
            sbtn.style.display = vis ? '' : 'none'
            if (window.__extRefBtn) window.__extRefBtn.style.display = vis ? '' : 'none'
          } catch (e) {}
        }
        __applyMaxVis()
        sbtn.onclick = function (ev) {
          ev.stopPropagation()
          try { __refBtnAction() } catch (e) {}
        }
        sb.parentNode.insertBefore(sbtn, sb.nextSibling)
        window.__extRefStripBtn = sbtn
        /* ★ 2026-10-19：记住左侧图标条容器——互斥守卫用它判定「点到图标条的其他图标」 */
        try {
          window.__extLeftStrip = sb.parentNode
          console.log('[cboo] left strip hooked: ' + String(sb.parentNode.id || sb.parentNode.className || 'unknown').slice(0, 60))
        } catch (eLS) {}
        // ★ 底部状态条内、「字数统计」左侧：参数引用开关（与状态栏同款内联样式，非悬浮）
        if (!edoc.getElementById('ext-btn-bottomleft')) {
          var bl = edoc.createElement('span')
          bl.id = 'ext-btn-bottomleft'
          bl.textContent = '数据引用'
          bl.style.cssText = 'display:inline-block;width:76px;flex-shrink:0;box-sizing:border-box;text-align:center;margin:0;cursor:pointer;font-size:11px;line-height:18px;color:inherit'
          bl.onclick = function () { try { __refBtnAction() } catch (e) {} }
          window.__extRefBottomBtn = bl
          /* ★ 原生标记按钮（2026-10-18 用户「增加个按钮，以原生的设置进行标记」）：
           * 紧挨「数据引用」，与浮层互斥 —— 填充直接写单元格模型（asc_setCellBackgroundColor），
           * 绝对跟格不漂不闪；代价见 __nf* 段注释。 */
          if (!edoc.getElementById('ext-btn-nativefill')) {
            var bl2 = edoc.createElement('span')
            bl2.id = 'ext-btn-nativefill'
            bl2.textContent = '原生标记'
            bl2.title = '用 OO 原生填充色标记引用格：绝对跟格、编辑不遮挡，新增引用自动补涂。' +
              '注意：文档会变为「未保存」（Ctrl+Z 可逐格撤销）；关闭时按每格原有填充色还原。'
            bl2.style.cssText = 'display:inline-block;width:64px;flex-shrink:0;box-sizing:border-box;text-align:center;margin:0;cursor:pointer;font-size:11px;line-height:18px;color:inherit;position:relative;z-index:9'
            /* ★ 点击异常不再静默：按钮变「原生 ✗」并把栈留在 window.__nfErr，
             *   否则「点了一点反应都没有」既看不到也查不到（旧写法 catch 空转）。 */
            bl2.onclick = function (ev) {
              try {
                window.__nfClicks = (window.__nfClicks || 0) + 1
                if (ev && ev.stopPropagation) ev.stopPropagation()
                /* ★ 2026-10-18：鼠标点击原生标记按钮本身也进日志（含点击时的前态） */
                try { __mkLog('MARK-CLICK', '鼠标点击「原生标记」按钮（第 ' + window.__nfClicks + ' 次）｜点击时状态[' + __nfSnapStr(__nfSnap()) + '] → 即将' + (__rm.native ? '关闭' : '开启')) } catch (eMC) {}
                __nfToggle()
              } catch (eN1) {
                try { window.__nfErr = String((eN1 && eN1.stack) || eN1).slice(0, 600) } catch (eNE) {}
                try { console.log('[cboo] nfToggle ERR ' + String((eN1 && eN1.stack) || eN1).slice(0, 300)) } catch (eLG4) {}
                try { __nfUi(String((eN1 && eN1.message) || eN1)) } catch (eNU) {}
                try { __nfTip('原生标记出错：' + String((eN1 && eN1.message) || eN1).slice(0, 60)) } catch (eNT) {}
              }
            }
            bl2.onmousedown = function (ev) { try { if (ev && ev.preventDefault) ev.preventDefault() } catch (eN0) {} }
            window.__extNativeBtn = bl2
            /* 桌面端诊断：renderer console 会落 .dev-console.log，用户反馈「点了没反应」时靠这几行定位 */
            try { console.log('[cboo] nativefill mounted frame=' + (__nfFrame() ? 'ok' : 'null')) } catch (eLG0) {}
          }
          var wcSpan = null
          edoc.querySelectorAll('span.caption').forEach(function (e) { if (!wcSpan && (e.textContent || '').trim() === '字数统计') wcSpan = e })
          var grp = wcSpan && wcSpan.closest('.status-group')
          // ★ 换位（用户 2026-10-01「左右换个方向」）：按钮放「第x页共y页」组（#label-pages）左侧
          //   ① ID 锚定（模板渲染即存在，挂载时页面文字还是空的，文字匹配不可靠）
          //   ② 文字匹配兜底（双语言） ③ 最后回退字数统计组
          var pgGrp = null
          var pgLbl = edoc.getElementById('label-pages')
          if (pgLbl) { try { pgGrp = pgLbl.closest('.status-group') } catch (e0) {} }
          if (!pgGrp) {
            // ⚠ 模板字面量内正则转义必须双写反斜杠：单写反斜杠d 反斜杠s 会被外层模板吞成
            //   d 和 s；单写反斜杠n 会变成真实换行，直接把注释与正则劈裂成语法错误（两次实锤）。
            var pgRe = /共[0-9]+页|of[0-9]+/i
            var wsRe = new RegExp('[ \\t\\r\\n]+', 'g')
            edoc.querySelectorAll('.status-group').forEach(function (g) {
              if (!pgGrp) {
                var t = (g.textContent || '').replace(wsRe, '')
                if (pgRe.test(t)) pgGrp = g
              }
            })
          }
          // ★ Excel（cell）：独立 status-group 放「< >」组左边（用户 2026-10-01）
          //   ⚠ 不能插进 #status-tabs-scroll 组内部：OO 用该组宽度定位 sheet 容器
          //   （sheets-bar-box.left = padding + tabs-scroll.width [+addtabs.width]），
          //   插组后所有元素右移 ourW 但推导不知道 → Sheet1 压住 ＋≡（23:10 实测）。
          //   修法（23:14）：组的实际宽度补给 tabs-scroll 的 css width → OO 自己的推导
          //   落到 addtabs 右缘 = Sheet1 右移、＋≡ 完整可见（菜单左移生效）；resize 触发重算
          var done = false
          var tabBack = edoc.getElementById('status-btn-tabback')
          if (tabBack) {
            var tabsGrp = tabBack.closest('.status-group')
            if (tabsGrp && tabsGrp.parentElement) {
              var bg = edoc.createElement('div')
              bg.className = 'status-group cnt-flex-center'
              bg.style.flexShrink = '0'
              bg.style.marginLeft = '-2px' // 贴最左（抵消 statusbar padding 0 2px）
              bg.appendChild(bl)
              tabsGrp.parentElement.insertBefore(bg, tabsGrp)
              try {
                // ★ v7（23:59 用户：sheet 贴着菜单右边、不要跳来跳去）：
                //   跳动根源 = 800ms pin 定时器与 OO 内联写入互相拉扯。
                //   改为零定时器：!important 样式表压死 OO 的内联 left（内联永远打不过 !important），
                //   X=8（紧贴）→ tabs宽=64 / addtabs负margin=-8 / sheets.left = addtabs右缘+8（一次算定）
                var X = 0 // Sheet1 完全贴住菜单（00:14 用户：向左移一丝）
                tabsGrp.style.width = (56 + X) + 'px'
                tabsGrp.style.paddingLeft = '4px' // ＜ 更贴数据引用（原 10px 内边距）
                tabsGrp.style.paddingRight = '4px' // ＋ 更贴 ＞（原 12px）
                var addBox = edoc.getElementById('status-addtabs-box')
                if (addBox) addBox.style.marginLeft = (-X) + 'px'
                if (window.__extSheetsPinTimer) { try { clearInterval(window.__extSheetsPinTimer) } catch (e9) {} window.__extSheetsPinTimer = null }
                var applySheetsPin = function () {
                  try {
                    var ab = edoc.getElementById('status-addtabs-box')
                    var sb = edoc.querySelector('.statusbar')
                    if (!ab || !sb) return
                    var abR = ab.getBoundingClientRect()
                    var sbR = sb.getBoundingClientRect()
                    if (abR.width === 0) return
                    var wantLeft = Math.round(abR.right - sbR.left + X)
                    var st = edoc.getElementById('ext-sheets-pin-style')
                    if (!st) {
                      st = edoc.createElement('style')
                      st.id = 'ext-sheets-pin-style'
                      edoc.head.appendChild(st)
                    }
                    /* ★ 2026-10-18「右下角比例点不动」真正根因（实测对照，已定案）：
                     *   原生 OO 里 #status-zoom-box 是 position:absolute，由 JS 定位在右下角
                     *   （1500 宽窗口实测 rect x=1360 w=140），#status-sheets-bar-box 宽度只等于
                     *   标签总宽（单 sheet 时 79px，右缘 269），二者互不干扰 —— 所以缩放本来就是
                     *   原生功能，−／百分比下拉／+ 三件套齐全（点一下 100%→90%→100%，下拉 272px 高）。
                     *   上一版两处写法把它弄坏了：
                     *     ① right:0 !important → sheet 条被拉成 190~1500（1310 宽），整片盖住缩放框；
                     *     ② position:relative !important → zoom 框脱离 absolute，掉回 flex 流第 4 位
                     *        （x=190），直接被 sheet 条压在底下 → 右下角连控件都看不见。
                     *   正确写法：只锁 left（值本身与原生一致），**不写 right**（保留 OO 的标签定宽），
                     *   zoom 框只提 z-index、绝不动 position（absolute 同样吃 z-index）。 */
                    var css = '#status-sheets-bar-box{left:' + wantLeft + 'px !important}' +
                      '#status-zoom-box{z-index:60 !important}'
                    if (st.textContent !== css) st.textContent = css
                  } catch (e8) {}
                }
                requestAnimationFrame(applySheetsPin)
                setTimeout(applySheetsPin, 400)
                if (!window.__extSheetsPinResize) {
                  window.__extSheetsPinResize = true
                  edoc.defaultView.addEventListener('resize', applySheetsPin)
                }
                // ★ v8（00:02 用户：最后一个 sheet 被遮挡 + 右键效果不对）：
                //   根因 = Statusbar 视图缓存的 tabBarDefPosition 仍是原生值，缩放框定位/溢出计算/
                //   右键菜单等一切消费方都拿它算 → 与我的几何脱节。直接 defineProperty 锁定该字段
                //   永远返回正确值（写入被吞），再触发 resize 让整套布局用正确数重排。
                var fixTabBarField = function () {
                  try {
                    var w = edoc.defaultView
                    var view = w.SSE && w.SSE.getController && w.SSE.getController('Statusbar') && w.SSE.getController('Statusbar').getView('Statusbar')
                    var ab = edoc.getElementById('status-addtabs-box')
                    var sb = edoc.querySelector('.statusbar')
                    if (!view || !ab || !sb) return false
                    var want = Math.round(ab.getBoundingClientRect().right - sb.getBoundingClientRect().left + X)
                    if (view.tabBarDefPosition !== want) {
                      try {
                        Object.defineProperty(view, 'tabBarDefPosition', { get: function () { return want }, set: function () {}, configurable: true })
                      } catch (eA) {
                        try { view.tabBarDefPosition = want } catch (eB) {}
                      }
                      try { w.dispatchEvent(new Event('resize')) } catch (eC) {}
                    }
                    return true
                  } catch (eD) { return false }
                }
                setTimeout(fixTabBarField, 400)
                setTimeout(fixTabBarField, 1200)
                setTimeout(fixTabBarField, 3000)
              } catch (e5) {}
              done = true
              // 触发 OO 重算 sheets-bar-box 位置（初始化与布局稳定后各一次）
              try { edoc.defaultView.dispatchEvent(new Event('resize')) } catch (e6) {}
              setTimeout(function () { try { edoc.defaultView.dispatchEvent(new Event('resize')) } catch (e7) {} }, 350)
            }
          }
          if (!done) {
            var anchor = pgGrp || grp
            if (anchor && anchor.parentElement) { anchor.parentElement.insertBefore(bl, anchor) }
            else {
              bl.style.cssText += ';position:absolute;left:8px;bottom:6px;z-index:8100;background:rgba(30,33,38,.85);color:#fff;border:1px solid rgba(255,255,255,.25);border-radius:6px;padding:3px 10px'
              edoc.body.appendChild(bl)
            }
          }
          /* 「原生标记」跟随「数据引用」落位（bl 三个分支都插完了，这里统一补挂） */
          try {
            if (window.__extNativeBtn && bl.parentNode && window.__extNativeBtn.parentNode !== bl.parentNode) {
              bl.parentNode.insertBefore(window.__extNativeBtn, bl.nextSibling)
            }
          } catch (eN2) { /* 忽略 */ }
        }
        // ★ 互斥守卫：挂在编辑器 document 捕获阶段（最先触发，OO 的 stopPropagation 拦不住）——
        //   点其他分类图标 → 关参数引用 + 还原推挤（保持单面板语义）
        if (!edoc.__extRefGuard) {
          edoc.__extRefGuard = true
          edoc.addEventListener('click', function (ev) {
            try {
              var closeIt = false
              var why = ''
              var el = ev.target
              while (el && !(el.classList && el.classList.contains('btn-category'))) el = el.parentElement
              if (el && el.id !== 'ext-btn-left-ooref') { closeIt = true; why = 'btn-category:' + String(el.id || '').slice(0, 30) }
              /* ★ 2026-10-19 用户：左侧图标条的其他图标（批注/搜索/目录…）点击时
               *   也要先关参数引用 —— 否则面板还开着，出现「要关两次」。
               *   容器级判定：图标 id 不一定都叫 left-btn-*，以注入时记下的
               *   __extLeftStrip（left-btn-comments 的父容器）为准；
               *   点在我们自己面板内（__extRefPanel/cont）不算。 */
              if (!closeIt && window.__extLeftStrip) {
                var s = ev.target
                var inPanel = false
                var inStrip = false
                var w = ev.target
                while (w && w !== edoc.body) {
                  if (w === window.__extRefPanel || w === window.__extRefPanel.parentElement) { inPanel = true; break }
                  if (w === window.__extLeftStrip) { inStrip = true; break }
                  w = w.parentElement
                }
                if (inStrip && !inPanel) {
                  var b = ev.target
                  var isOurs = false
                  while (b && b !== window.__extLeftStrip) {
                    if (b.id === 'ext-btn-left-ooref') { isOurs = true; break }
                    b = b.parentElement
                  }
                  if (!isOurs) { closeIt = true; why = 'strip-icon:' + String(ev.target.id || ev.target.className || ev.target.tagName).slice(0, 30) }
                }
              }
              var panelOpen = !!(window.__extRefPanel && window.__extRefPanel.style.display !== 'none')
              /* ★ 用户指令「记录我每一步操作」：图标条区域的每次点击都落日志 */
              var inStripArea = !!(window.__extLeftStrip && (function (t0) { var w0 = t0; while (w0 && w0 !== edoc.body) { if (w0 === window.__extLeftStrip) return true; w0 = w0.parentElement } return false })(ev.target))
              if (inStripArea) {
                var lg = '点击图标条: ' + String(ev.target.id || ev.target.className || ev.target.tagName).slice(0, 40) +
                  '｜命中关闭=' + (closeIt ? why : '否') + '｜参数引用面板开=' + (panelOpen ? '是' : '否')
                try { console.log('[cboo] REF-GUARD ' + lg) } catch (eL1) {}
                try { fetch('/office-log', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ type: 'REF-GUARD', detail: lg }) }) } catch (eL2) {}
              }
              if (closeIt && panelOpen) {
                /* ★ 2026-10-19 修「框还在」：不能只 display='none'——左侧菜单被面板
                 *   撑到 300px 的推挤宽度不会还原，用户看到的就是一条空黑区。
                 *   必须走完整收起流程 __setRefPanelVisible(false)：藏面板 + 还原
                 *   left-menu 宽度 + 触发引擎重排。 */
                try { __setRefPanelVisible(false) } catch (eC1) { window.__extRefPanel.style.display = 'none' }
                if (window.__setRefPush) window.__setRefPush(false)
                if (window.__extRefStripBtn) window.__extRefStripBtn.classList.remove('active')
                try { parent.postMessage({ __ooRefState: false }, '*') } catch (e2) {}
              }
            } catch (e) {}
          }, true)
        }
      }
    } catch (e) {}
  }
  /**
   * 插件（cbref）↔ 宿主 桥（批次 12 · ③）
   * 插件在编辑器自己的 iframe 里跑，parent 就是本页。它用__cbRef* 系列消息与宿主对话。
   * 结果一律经 __ooStatus 回传，进宿主诊断抽屉（用户看界面即可，无需开 Console）。
   */
  window.addEventListener('message', function (e) {
    var d = e.data
    if (!d || typeof d !== 'object') return
    // ★ 2026-10-10 域关键修复：asc_onSelectionNameChanged 回调被 sdkjs 在【引擎域】
    //   （frameEditor）里执行——回调里 window.__ooActiveCell=H163 写进了 frameEditor
    //   的 window，而 /poc 读自己的 → 永远 stale A1（写入全落 A1 的真根因）。
    //   事件回调只发 __ooCursor 消息，光标缓存统一在【本域】落账 + 转发宿主。
    //   （正则用 [0-9]：模板字面量会把 \d 吞成 d —— 2026-10-11 实锤修复，此前守卫恒不匹配）
    if (typeof d.__ooCursor === 'string' && d.__ooCursor && /^[A-Z]+[0-9]+$/.test(d.__ooCursor) && !d.__ooCursorFwd) {
      try { window.__ooActiveCell = d.__ooCursor } catch (e0) {}
      try { parent.postMessage({ __ooCursor: d.__ooCursor, __ooCursorFwd: true }, '*') } catch (e1) {}
    }
    if (d.__cbRefReady) {
      try { parent.postMessage({ __ooStatus: '[cbref] 插件就绪 editorType=' + (d.ready ? 'contentReady' : 'init') + ' ' + JSON.stringify(d.editorType || '') }, '*') } catch (e1) {}
      // ★ 插件就绪即**索要一次光标位置**。
      //   没有这一步，用户「先点单元格、再开参数引用面板」时宿主侧是空的，
      //   插入会回落写 A1 ⇒ 在别的行列上表现为「点了参数引用没反应」。
    }
    if (d.__cbRefCursor) {
      // 插件回传的光标格地址：记下来供插入时用，并同步给宿主（诊断抽屉可见）
      try { __trackCursor(d.__cbRefCursor.a1) } catch (e1c) {}
    }
    if (d.__cbRefRead) {
      // ★ 2026-10-10：插件回读结果（readCell）→ 宿主诊断抽屉（写入链路自证）
      try { parent.postMessage({ __ooStatus: '[cbref-read] ' + (d.__cbRefRead.a1 || '?') + ' 值=' + String(d.__cbRefRead.v == null ? '' : d.__cbRefRead.v).slice(0, 200) + ' 显示=' + String(d.__cbRefRead.t == null ? '' : d.__cbRefRead.t).slice(0, 80) }, '*') } catch (e1d) {}
      /* ★ 保存链重构 v2：读取结果转发宿主（触发 xlsx 构建+ooBytes64 入库） */
      if (window.__saveAwaitRead && d.__cbRefRead && d.__cbRefRead.ok && d.__cbRefRead.values) {
        window.__saveAwaitRead = false
        try { parent.postMessage({ __ooSaveValues: { name: "+JSON.stringify(srcFile)+", rows: d.__cbRefRead.values, mode: saveMode || 'node' } }, '*') } catch (eSV) {}
      }
    }
    if (d.__cbRefShape) {
      // ★ 2026-10-10：ppt 形状插入回调
      try { parent.postMessage({ __ooStatus: '[cbref-ppt] 形状回调 ' + String(d.__cbRefShape.ok || '').slice(0, 40) + ' text=' + String(d.__cbRefShape.text || '').slice(0, 40) }, '*') } catch (e1e) {}
    }
    if (d.__cbRefProbe) {
      // 探针结论：写 Z1 是否成功 + 读回值 —— 直接决定「联动能不能通」
      try { parent.postMessage({ __ooStatus: '[cbref] 探针 wrote=' + d.__cbRefProbe.wrote + ' readBack=' + d.__cbRefProbe.readBack + ' type=' + d.__cbRefProbe.editorType }, '*') } catch (e2) {}
    }
      if (d.__cbRefDone) {
        try { parent.postMessage({ __ooStatus: '[cbref] 写入完成 ' + d.__cbRefDone.wrote + '/' + d.__cbRefDone.total }, '*') } catch (e3) {}
      }
      if (d.__cbRefDefName) {
        // ★ 2026-10-11：定义名注册真实回执（ok/via/readback/recalc/retry/variant/why）
        try { parent.postMessage({ __ooStatus: '[defname] 注册结果 ok=' + d.__cbRefDefName.ok + ' name=' + (d.__cbRefDefName.name || '') + (d.__cbRefDefName.via ? ' via=' + d.__cbRefDefName.via : '') + (d.__cbRefDefName.readback ? ' 读回=' + d.__cbRefDefName.readback : '') + (d.__cbRefDefName.recalc ? ' 重算=' + d.__cbRefDefName.recalc : '') + (d.__cbRefDefName.retry ? ' (重试成功)' : '') + (d.__cbRefDefName.cleared ? ' 已清格重注=' + d.__cbRefDefName.cleared : '') + (d.__cbRefDefName.variant ? ' 变体=' + d.__cbRefDefName.variant : '') + (d.__cbRefDefName.why ? ' why=' + d.__cbRefDefName.why : '') }, '*') } catch (eD3) {}
        /* ★ 变体回执（原名被引擎判 Invalid，变体已注册）→ 改写目标格公式 =变体：
         *   登记表 name 跟随（⟳ 刷新仍按 row）；自愈目标跟随（回车提交后自愈写变体）；
         *   已提交且呈 #NAME?/空的格立即条件改写 */
        if (d.__cbRefDefName.variant && d.__cbRefDefName.a1) {
          var __vv = String(d.__cbRefDefName.variant), __va = String(d.__cbRefDefName.a1), __vo = String(d.__cbRefDefName.name || '')
          try {
            Object.keys(window.__extRefRegistry || {}).forEach(function (k4) {
              var en4 = window.__extRefRegistry[k4]
              if (en4 && en4.a1 === __va && en4.name === __vo) en4.name = __vv
            })
          } catch (eR4) {}
          try { if (window.__editVerify && window.__editVerify.a1 === __va) window.__editVerify.name = __vv } catch (eR5) {}
          try {
            __sendToRefPlugin({ __cbRefCmd: 'script', text: 'try{var ws=Api.GetActiveSheet();var r=ws.GetRange("' + __va + '");var val=String(r.GetValue()==null?"":r.GetValue());if(val===""||val.charAt(0)==="#"){r.SetValue("=' + __vv + '");try{var rc2=ws.GetRange("' + __va + '");rc2.SetBorders("all",Api.CreateStroke(6350,Api.CreateSolidFill(Api.CreateRGBColor(88,166,255))));rc2.SetFillColor(Api.CreateRGBColor(232,244,255));rc2.SetFontColor(Api.CreateRGBColor(9,105,218));rc2.SetBold(true)}catch(eM2){}}}catch(e){}' })
          } catch (eR6) {}
        }
        /* ★ 2026-10-18 引入即标记（用户「引入的数据 引入后 没标记」）：
         *   定义名注册回执 = 引用真正落地的唯一汇合点（两条插入路径都过这里）。
         *   ① 视觉覆盖层（着色开关开）：追加 {sheet,ref,color}——色用与宿主
         *      refColorOf 完全相同的 djb2 六色板（同一参数永远同色），追加后
         *      置 needMeas 触发重测量；② 原生标记（原生标记开）：排队一次选区
         *      校验，__nfCheckSelection 会对当前格（=刚引入的格）就地补涂。 */
        try {
          var __dnM = d.__cbRefDefName
          var __mkA1 = String(__dnM.a1 || '')
          var __mkName = String(__dnM.variant || __dnM.name || '')
          if (__mkA1 && /^[A-Z]+[0-9]+$/.test(__mkA1) && __mkName) {
            try {
              /* ★ 2026-10-18 修「开原生标记时漏涂刚引入的格」：追加**不依赖**覆盖层开关
               *   __rm.on——覆盖层没开时也要把新格写进 marks 清单，否则原生标记涂色
               *   按 marks 走会把新格漏掉（22:42 日志实锤：引入 C1，开灯 30 格无 C1，
               *   2 秒后才靠选区校验补涂）。清单完整 = 两个开关数据一致。 */
              if (__rm.marks) {
                var __has = false
                for (var __mi = 0; __mi < __rm.marks.length; __mi++) { if (String(__rm.marks[__mi].ref) === __mkA1) { __has = true; break } }
                if (!__has) {
                  var __shM = ''
                  try {
                    var __feD = document.querySelector('iframe[name="frameEditor"]') || document.querySelector('iframe')
                    var __liM = __feD && __feD.contentDocument ? __feD.contentDocument.querySelector('#status-sheets-bar li.list-item.active') : null
                    __shM = __liM ? String(__liM.textContent || '').trim() : ''
                  } catch (eS9) {}
                  var __pal = ['FFFFF2CC', 'FFCCE5FF', 'FFD9F2D9', 'FFE8DDFF', 'FFFFE0CC', 'FFFFD6E7']
                  var __hhM = 5381
                  for (var __hi = 0; __hi < __mkName.length; __hi++) { __hhM = (((__hhM << 5) + __hhM) ^ __mkName.charCodeAt(__hi)) >>> 0 }
                  __rm.marks.push({ sheet: __shM, ref: __mkA1, color: __pal[__hhM % __pal.length] })
                  __rm.needMeas = true
                  __rm.recheckLeft = 2
                  __rm.sig = ''
                  if (!__rm.raf) { try { __rm.raf = window.requestAnimationFrame(__rmTick) } catch (eRA0) { __rm.raf = window.setTimeout(__rmTick, 60) } }
                  try { console.log('[cboo] insert mark + ' + __mkA1 + ' color=' + __pal[__hhM % __pal.length]) } catch (eLG9) {}
                }
              }
            } catch (eM9) {}
        try { __mkLog('REF-LAND', '引用落地：' + __mkName + ' → ' + __mkA1 + '（定义名注册回执 ok=' + __dnM.ok + '）') } catch (eLK0) {}
        try {
          if (__rm.native) {
            var __bSnap = __nfSnap()
            /* ★ 2026-10-18 用户指令「不要后面检测 要在插入的时候做标记」：
             *   废弃「排选区检测靠光标还在目标格」的补涂路径——直接对目标格涂色。
             *   延迟 900ms：等插件写公式落地（域数据 0ms → 定义名 100ms → 公式 600ms），
             *   涂色用 __nfPaintKeys（jump→记原色→涂→跳回，不依赖光标位置）。 */
            setTimeout(function () {
              try {
                var fP = __nfFrame()
                var shP = fP ? __nfSheetNow(fP.t.doc) : ''
                var keyP = (shP ? shP + '!' : '') + __mkA1
                var painted = __nfPaintKeys([keyP], true, __rm.nfColor || 'FFF2CC')
                __nfDiffReport('插入即标记 ' + keyP, __bSnap, painted.length ? '' : '涂色失败（编辑器帧未就绪）')
              } catch (eP0) {}
            }, 900)
          } else {
            /* ★ 原生标记没开也要有记录（用户指令）：引入了但没标记，说清楚原因与补救 */
            __mkLog('MARK-SKIP', '原生标记未开启——' + __mkA1 + '（' + __mkName + '）引入后未自动标记；开启原生标记后点该格即可补涂')
          }
        } catch (eN9b) {}
          }
        } catch (eMK0) {}
      }
      // ★ 章节目录结果转发（2026-10-09）：插件 → 本页 → React 宿主侧栏
      if (d.__chaptersData) { try { parent.postMessage({ __chaptersData: d.__chaptersData }, '*') } catch (e4) {} }
      if (d.__chaptersMoved) { try { parent.postMessage({ __chaptersMoved: d.__chaptersMoved }, '*') } catch (e5) {} }
  })
  /* ★ 2026-10-11 版本标记：诊断抽屉首行可见，杜绝「改了没生效」的版本歧义 */
  try {
  /* ★ 版本标记（由 scripts/oo-version.mjs bump 自动改写——勿手改本行） */
  var __ocVer = 'v20261017', __ocHash = '268dfb6', __ocNote = '引用格着色：同一参数同色，写进文件样式（打开即有，不触发脏标记）+ 面板色块图例'
  if (window.__ocVerMarked !== __ocVer) {
    window.__ocVerMarked = __ocVer
    if (window.parent && window.parent !== window) parent.postMessage({ __ooStatus: '[oo-core] ' + __ocVer + ' (' + __ocHash + ') 已加载（' + __ocNote + '）' }, '*')
  }
  } catch (eVer) {}

  /**
   * 把写值指令送进 cbref 插件（批次 12 · ③）。
   *
   * 为什么必须走插件：写单元格要用 Office JS API，它只在插件上下文里可用；
   * 宿主页面直接 api.callCommand 静默失效（已调研确认）。插件由 autostart 自动启动，
   * 跑在自己的 iframe 里，parent 即本页，故用 parent.postMessage 投递。
   *
   * @returns true=已投递到至少一个插件 iframe；false=没找到插件（调用方据此决定是否退回）
   */
  function __sendToRefPlugin(payload) {
    var frames = []
    try {
      var eds = document.querySelectorAll('iframe')
      for (var i = 0; i < eds.length; i++) {
        var src = eds[i].getAttribute('src') || ''
        if (src.indexOf('cbref') >= 0) frames.push(eds[i])
      }
    } catch (e) {}
    /* ★ 2026-10-10 cellDiag 实锤补搜：/poc 全文档只有 frameEditor 一个 iframe，
     *   插件 iframe 活在 frameEditor（web-apps 页）内部 —— 上面本文档搜索永远空。
     *   同源（都走本服务代理）⇒ 进 frameEditor.contentDocument 接着搜。 */
    try {
      var fe2 = document.querySelector('iframe[name="frameEditor"]')
      var feDoc = fe2 && fe2.contentDocument
      if (feDoc) {
        var eds2 = feDoc.querySelectorAll('iframe')
        for (var j = 0; j < eds2.length; j++) {
          var src2 = eds2[j].getAttribute('src') || ''
          if (src2.indexOf('cbref') >= 0) frames.push(eds2[j])
        }
      }
    } catch (e4) {}
    try { parent.postMessage({ __ooStatus: '[cbref] 帧搜索: frames=' + frames.length }, '*') } catch (eF) {}
    if (!frames.length) {
      // ★ 只在「真的要写值且没找到插件」时才报 —— 像 cursor 这种
      //   「顺带问一句」的指令找不到插件属正常，不该刷诊断日志。
      if (payload && payload.cells) {
        try { parent.postMessage({ __ooStatus: '[cbref] 未找到插件 iframe（autostart 未生效？）' }, '*') } catch (e2) {}
      }
      return false
    }
    // ★ 指令名与载荷必须从 payload 透传，不能硬编码 'write'：
    //   原实现固定发 write 且只带 cells，于是 cursor 这类无 cells 的指令
    //   会被降级成「空写」而拿不到任何回传。
    var cmd = (payload && payload.__cbRefCmd) || 'write'
    // ★★ 2026-10-06：同时走**官方 External Listener 信封**。
    //   官方消息格式（sdkjs-plugins/externallistener/test2/test2.js，三处来源一致）：
    //     { frameEditorId, guid, type: 'onExternalPluginMessage', data: {...} }
    //   ★ 我们原先只发自定义字段 __cbRefCmd，而插件侧压根没有 onExternalPluginMessage
    //     处理器 ⇒ 消息全部石沉大海。这才是「点了没反应」在 OO 侧的真正根因。
    //   guid 必须与 config.json 完全一致，否则官方通道不认。
    // ★★★ 2026-10-06 GitHub 官方源码核对后的修正（前一版载荷结构是错的）：
    //   官方 externallistener/scripts/code.js 的 onExternalPluginMessage **只有三种 type**：
    //     case 'close'        → this.executeCommand('close', '')
    //     case 'insertText'   → Asc.scope.text = data.text; this.callCommand(...)
    //     case 'executeCommand' → this.info.recalculate = true;
    //                            this.executeCommand('command', data.text)
    //   ★★ data 是**扁平**的，只有 type 与 text（**text 是脚本文本字符串**）。
    //   我前一版发的是 { type:'writeCells', cells:[...] } ——
    //   writeCells 这个 type 官方 switch 里**根本不存在** ⇒ 全部落空（无 default 分支）。
    //   ⇒ 改为：把写单元格拼成 Office JS 脚本，用官方 executeCommand 通道下发。
    var data = {}
    if (payload) {
      if (cmd === 'write' && payload.cells && payload.cells.length) {
        data.type = 'executeCommand'
        data.text = __buildWriteScript(payload.cells, payload.recalc !== false)
      } else if (cmd === 'cursor') {
        // 取光标：官方无对应 type，仍走自建 getCursor（插件侧已实现）；
        // 同时带上 executeCommand 版本的探测，便于对照哪条路真的通。
        data.type = 'getCursor'
        data.text = 'var s=Api.GetSelection();s&&s.SetValue&&0;'
      } else if (cmd === 'chapter' && payload.chapter) {
        // ★ 章节目录（2026-09）：getChapters/gotoChapter/moveChapter，
        //   插件侧自建 type 走 callCommand（writeCells 先例），结果经 __chaptersData/__chaptersMoved 回传。
        data.type = String(payload.chapter.op || '')
        if (payload.chapter.i !== undefined) data.i = payload.chapter.i
        if (payload.chapter.from !== undefined) data.from = payload.chapter.from
        if (payload.chapter.count !== undefined) data.count = payload.chapter.count
        if (payload.chapter.to !== undefined) data.to = payload.chapter.to
      } else if (cmd === 'read' && payload.a1) {
        // ★ 2026-10-10 新增：读单元格（插件会话内 callCommand 有回调，免费版唯一可靠读取通道）
        data.type = 'readCell'
        data.a1 = payload.a1
      } else if (cmd === 'script' && payload.text) {
        data.type = 'executeCommand'
        data.text = payload.text
      } else if (cmd === 'insertShape' && payload.text !== undefined) {
        // ★ 2026-10-10 新增：ppt 当前页插引用文本形状
        data.type = 'insertShape'
        data.text = payload.text
      } else if (cmd === 'defname' && payload.name && payload.ref) {
        /* ★★ 2026-10-11 修复（用户实测插件自爆根因）：defname 此前无分支 → 落入
         *   兜底 else 的 data.type='close' → 插件 executeCommand('close') 自我关闭
         *   → frames=1→0 全通道死亡（⟳/域数据/写值全停摆，即用户所见"卡住"）。
         *   现走官方信封 type='defname'，由插件 onExternalPluginMessage → applyDefName 处理。
         *   ★ a1 必须透传（11:16 实测漏传 → 插件墓碑清格路径被跳过、宿主变体改写
         *   定位不到格子 → C155 恒 #NAME?） */
        data.type = 'defname'
        data.name = payload.name
        data.ref = payload.ref
        data.a1 = payload.a1 || ''
      } else {
        data.type = 'close'
        data.text = ''
      }
    }
    for (var k = 0; k < frames.length; k++) {
      try {
        // ① 官方信封（postMessage 必须传**字符串**）
        frames[k].contentWindow.postMessage(JSON.stringify({
          frameEditorId: 'frameEditor',
          guid: 'asc.{6B3F1A2C-4D5E-4F60-9A7B-2C3D4E5F6A7B}',
          type: 'onExternalPluginMessage',
          data: data,
        }), '*')
        // ② 旧的自定义字段一并带上（裸 listener 直连插件 iframe 的真实通道）——
        //    ★ 2026-10-10 修复：此前只透传 cells/recalc，chapter/read/insertShape
        //    的载荷全被丢掉；改为全量透传
        var msg = Object.assign({ __cbRefCmd: cmd }, payload || {})
        frames[k].contentWindow.postMessage(msg, '*')
      } catch (e3) {}
    }
    return true
  }

  /**
   * 把「写单元格」编成Office JS 脚本（官方 executeCommand 通道执行的就是它）。
   *
   * ★ 为什么用字符串脚本而不是 Asc.scope 传数组：
   *   官方 executeCommand 分支只做 this.executeCommand('command', data.text)，
   *   data.text 直接当脚本执行 ⇒ 只能走字符串，没有结构化通道。
   *   （Asc.scope 只在 insertText 分支里用到。）
   *
   * ⚠ 值必须安全转义：Office JS 脚本里用双引号包字符串，
   *   值里的引号/反斜杠/换行不处理会破坏脚本 ⇒ 用 JSON.stringify 产出合法字面量。
   */
  function __buildWriteScript (cells, recalc) {
    var lines = []
    lines.push('var s=Api.GetActiveSheet();')
    for (var i = 0; i < cells.length; i++) {
      var c = cells[i]
      if (!c || !c.a1) continue
      // JSON.stringify 把值转成合法 JS 字面量（含引号转义与换行转义）
      var v = JSON.stringify(c.value == null ? '' : String(c.value))
      // ★ '=x.y' 引用文本（2026-10-10）：SetValue 按值写、SetFormula 按公式写 —— 双写保真，
      //   哪个通道生效都能落出公式；SetFormula 不支持时 try/catch 静默吞掉不碍事
      if (c.value != null && String(c.value).charAt(0) === '=') {
        lines.push('var r=s.GetRange(' + JSON.stringify(String(c.a1)) + ');r.SetValue(' + v + ');try{r.SetFormula(' + v + ')}catch(e){};')
      } else {
        lines.push('s.GetRange(' + JSON.stringify(String(c.a1)) + ').SetValue(' + v + ');')
      }
    }
    if (recalc) lines.push('s.GetRange(\"A1\");')   // 触碰一次促使重算（官方靠 info.recalculate=true）
    return lines.join('')
  }

  // ★ 数据引用按钮动作按模式分流：最大化=窗内面板开合；普通=宿主浮卡开合
  function __refBtnAction() {
    // ★ 每次点按钮都重新要一次光标位置 —— 用户很可能在打开面板之后又换了单元格。
    //   只在「文档就绪」时取一次是不够的（面板开合与点单元格都在就绪之后发生）。
    if (window.__ooMax === true) { try { __toggleRefPanel() } catch (e) {} }
    else { try { parent.postMessage({ __ooToggleRef: 1 }, '*') } catch (e) {} }
  }

  /**
   * 「刷新域值」按钮：把画布侧的引用值写回本文件。
   *
   * ★ 为何由宿主执行而不是编辑器内直接写：写回要读模板字节、解析域落点、
   *   做补丁式改写、再存回模板库 —— 这些能力只有宿主（/poc 的parent）有。
   *   编辑器内（/poc）只是个显示壳。
   *
   * ⚠ 本段位于 /poc 的 HTML 模板字符串内，注释里**禁止用反引号**（会提前闭合外层模板）。
   */
  function __mountRefreshBtn() {
    try {
      var edoc = document.getElementById('document-editor') || document
      var src = edoc.getElementById('left-btn-comments')
      if (!src || document.getElementById('ext-btn-refresh')) return
      var btn = src.cloneNode(true)
      btn.id = 'ext-btn-refresh'
      btn.classList.remove('active', 'notify')
      btn.setAttribute('title', '刷新域值：把画布侧引用的当前值写回本文件')
      var icon = btn.querySelector('svg')
      if (icon) { icon.innerHTML = '<path d="M4 12a8 8 0 1 1 2.3 5.7M4 12V7m0 5h5"/>' }
      var lab = btn.querySelector('.btn-text, span')
      if (lab) lab.textContent = '刷新域值'
      btn.onclick = function () {
        // 请求宿主执行写回；结果由宿主回 __ooRefreshMsg 展示
        try { parent.postMessage({ __ooRefreshRef: 1 }, '*') } catch (e) {}
        try { setRefreshMsg('正在写回…') } catch (e2) {}
      }
      src.parentNode.insertBefore(btn, src.nextSibling)
      window.__extRefreshBtn = btn
    } catch (e) { /* 工具栏结构变了就不注入，不影响其他功能 */ }
  }

  /** 显示写回结果（宿主回传） */
  function setRefreshMsg (s) {
    window.__ooRefreshMsgText = String(s == null ? '' : s)
    try {
      var el = document.getElementById('ext-refresh-msg')
      if (!el) {
        el = document.createElement('div')
        el.id = 'ext-refresh-msg'
        el.style.cssText = 'position:absolute;left:8px;bottom:8px;z-index:9999;' +
          'font:12px/1.5 system-ui;background:rgba(0,0,0,.72);color:#fff;' +
          'padding:4px 8px;border-radius:4px;max-width:70%;'
        ;(document.body || document.documentElement).appendChild(el)
      }
      el.textContent = window.__ooRefreshMsgText
      clearTimeout(window.__ooRefreshMsgTimer)
      window.__ooRefreshMsgTimer = setTimeout(function () {
        try { el.style.display = 'none' } catch (e) {}
      }, 6000)
    } catch (e) { /* ignore */ }
  }

  /**
   * 向 cbref 插件索要当前光标单元格地址。
   * 复用 __sendToRefPlugin 的 iframe 查找与投递逻辑，插件回 __cbRefCursor。
   * ⚠ 插入时**不能同步拿到结果**（跨 iframe 往返），所以取回来先缓存进
   *   window.__ooActiveCell，插入时读缓存 —— 这是异步的固有代价。
   */
  function __askCursor() {
    if (String(window.__ooKind || '') !== 'oosheet') return false
    return __sendToRefPlugin({ __cbRefCmd: 'cursor' })
  }

  /**
   * 表格单元格写入能力探针（批次 12 · ③ 排障）
   * ---------------------------------------------------------------
   * 背景：用户报「单元格里输入 = 后点节点参数，参数进不去」。
   * 根因定位：__insertRefFormula 用的是 asc_AddText —— 那是**文档文本流** API（Word 正文用），
   *   对 Excel 单元格无效。宿主直连的 1941 个 asc_* 方法里没有任何「设单元格值」的方法，
   *   真正的入口是 callCommand（OO 插件执行通道，可跑 Office API 代码）。
   *
   * 本探针把结论经 __ooStatus 回传宿主诊断抽屉（用户无需打开 Console）。
   * 注意：本段位于 /poc 的 HTML 模板字符串内，注释里禁止用反引号。
   */
  function __cellApiDiag() {
    function rep(s) { try { parent.postMessage({ __ooStatus: s }, '*') } catch (e) {} }
    //★ 2026-10-07 增强：先把 iframe 树整体报出来。
    //   实测症状（用户诊断日志）：
    //     cellDiag: 窗口大对象=window,self,frames,top,parent,AscCommon,__font_names,Hc
    //     cellDiag: editor=无, DocsAPI=undefined
    //   那串对象名里含 window/self/frames/top/parent ⇒ **枚举的是顶层窗口**，
    //   说明 document.querySelector('iframe[name=frameEditor]') 拿到的是
    //   **编辑器容器 iframe**，而不是文档引擎 iframe。
    //   只看 editor=无 分不清是「没建引擎」还是「找错了窗口」——
    //   ★ 所以必须先把每层 iframe 的 name/src/有无 editor 如实打出来。
    try {
      var all = document.querySelectorAll('iframe');
      var desc = [];
      for (var q = 0; q < all.length && q < 8; q++) {
        var it = all[q];
        var win = null;
        try { win = it.contentWindow } catch (we) {}
        desc.push((it.name || '(无名)') + (it.src ? ' src=' + String(it.src).slice(-28) : ' src=空') +
          ' | editor=' + (win && win.editor ? '有' : '无') +
          ' | DocsAPI=' + (win ? typeof win.DocsAPI : 'n/a'))
      }
      rep('cellDiag: iframe树(' + all.length + ') = ' + (desc.join(' ;; ') || '（一个都没有）'))
    } catch (eTree) {
      rep('cellDiag: iframe 树枚举失败 ' + (eTree && eTree.message))
    }
    var fE = document.querySelector('iframe[name="frameEditor"]')
    if (!fE || !fE.contentWindow) { rep('cellDiag: 无 frameEditor'); return }
    var w = fE.contentWindow
    // ★ 2026-10-07 19:5x 修正判据：editor 实例挂载点因文档类型而异
    //   （word=window.editor；xlsx=Asc.editor 等）——只查 w.editor 会误报「editor=无」
    //   （19:46 实锤：xlsx 文档已打开 documentReady ✅ 却报 editor=无）。
    //   DocsAPI 在 poc 顶层（本 iframe 里恒 undefined），报「顶层/本层」而非误导性的 undefined。
    var api = w.editor || (w.Asc && w.Asc.editor) || null
    var docsapiLoc = (typeof w.DocsAPI !== 'undefined') ? '本层' : (window.DocsAPI ? '顶层(已建)' : '无')
    rep('cellDiag: editor=' + (api ? '有' : '无') + ', Asc=' + (w.Asc ? '有' : '无') + ', DocsAPI=' + docsapiLoc)
    if (!api) {
      // 枚举窗口上像 api 的大对象，便于判断实例挂在哪
      var cands = []
      try {
        for (var k in w) {
          try { if (w[k] && typeof w[k] === 'object' && Object.keys(w[k]).length > 50) cands.push(k) } catch (e1) {}
        }
      } catch (e2) {}
      rep('cellDiag: 窗口大对象=' + cands.slice(0, 8).join(',') )
      return
    }
    var names = {}
    try { for (var n in api) names[n] = 1 } catch (e3) {}
    for (var o2 = Object.getPrototypeOf(api); o2 && o2 !== Object.prototype; o2 = Object.getPrototypeOf(o2)) {
      try { for (var n2 in o2) names[n2] = 1 } catch (e4) {}
    }
    var all = Object.keys(names)
    rep('cellDiag: api 方法数=' + all.length)
    rep('cellDiag: callCommand=' + typeof api.callCommand)
    rep('cellDiag: 单元格相关=' + all.filter(function (n) { return /Cell|Range|SetValue|Formula|Paste|Input/i.test(n) }).slice(0, 12).join(','))
    /* ★ 2026-10-17 用户指令「别再往我文件里塞垃圾」：
     *   原实现在这里往 Z1 写 =diag.probe 来验证通道，而且每 5 秒重试一次、最多 12 轮
     *   ⇒ 每个打开过的文档都被塞进一个 #NAME? 单元格，保存后还会被提取进新文件
     *   （实测留档摘要里的 #NAME? 就是它，不是用户内容）。
     *   通道是否可用根本不需要写格子：callCommand 能回调一个字符串就证明通了。
     *   改为**纯只读**探测（读 GetUsedRange 的维度），Document 完全不被修改。 */
    if (typeof api.callCommand === 'function') {
      try {
        api.callCommand(function () {
          try {
            var sh = Api.GetActiveSheet()
            if (!sh) return 'err:no-sheet'
            var dim = 'empty'
            try {
              var u = sh.GetUsedRange()
              if (u) dim = u.GetRowsCount() + 'x' + u.GetColumnsCount()
            } catch (eU) { dim = 'usedrange-err' }
            return 'ok:' + dim
          } catch (e1) { return 'err:' + String(e1).slice(0, 60) }
        }, false, false, function (res) {
          rep('cellDiag: callCommand 回调 res=' + String(res).slice(0, 60) + '（有回调=直连通道通，全程未改动文档）')
          /* ★ 2026-10-23 通道验证成功即停探测（每轮枚举 2435 方法是持续消耗） */
          try { clearInterval(_diagTimer) } catch (eS) {}
        })
        rep('cellDiag: callCommand 已下发（只读探测）')
      } catch (e5) { rep('cellDiag: callCommand 抛错 ' + String(e5).slice(0, 120)) }
    }
  }
  window.__ooCellDiag = __cellApiDiag
  // ★ 2026-10-07：改为**周期性重试**（原来只跑 1 次、在 4 秒时）。
  //   实测用户日志：4 秒时引擎还没起（editor=无），
  //   而文档实际还在 loading doc…（日志里连续 15 秒都是 loading）。
  //   ⇒ 跑一次只能证明「4 秒时没好」，**证明不了「后来好没好」**——
  //   而这正是之前反复误判的根源（探针在错误的时间点采样）。
  //   改为每 5 秒一次、到 60 秒止；一旦 editor 出现就停（避免刷屏）。
  try {
    var _diagN = 0
    var _diagTimer = setInterval(function () {
      _diagN++
      try {
        __cellApiDiag()
        // 引擎已就绪 ⇒ 不再重复报（★ 2026-10-09 实锤：sheet 引擎在本页 window.editor，
        //   frameEditor.contentWindow.editor 永远是无 → 原停止条件永不触发，刷满 12 轮）
        var f = document.querySelector('iframe[name="frameEditor"]')
        var fApi = f && f.contentWindow ? (f.contentWindow.editor || (f.contentWindow.Asc && f.contentWindow.Asc.editor)) : null
        if (window.editor || fApi) { clearInterval(_diagTimer); return }
      } catch (e2) {}
      if (_diagN >= 4) clearInterval(_diagTimer)   // 48 秒后放弃
    }, 5000)
    setTimeout(function () { __cellApiDiag() }, 4000)   // 保留首次（立即可见）
  } catch (e) {}

  /**
   * 插入引用公式（批次 12 · ③ A 方案）
   * ---------------------------------------------------------------
   * 需求：OO Excel 里点节点参数 → 像 Excel 单元格地址那样参与公式计算。
   *
   * A 方案做法：往光标处写一条文本标记（=节点名.字段名），并登记到 __extRefCapsules
   *   让它在文档里以胶囊形态显示。这样"引用"这件事先跑通，且完全不碰 OO 内核的数值域
   *   （命名区域/公式引擎是另一层，A 方案不引这层依赖）。
   *   实际取值走既有 __ooRefData 回流：宿主算完 → 刷新胶囊显示值。
   *
   * 与双击参数行的差别：双击是「直接关联 = 插值 + 连线动画」，
   *   本函数是「插入可参与公式的引用文本」，语义不同，故分开。
   *
   * ⚠ 本函数位于 /poc 的 HTML 模板字符串内，写注释时禁止用反引号 ——
   *   会提前闭合外层模板字面量，报 "Invalid left-hand side in assignment"。
   */
  /**
   * 取当前光标所在的单元格地址（如 'B3'）。
   * ---------------------------------------------------------------
   * ★ 为什么必须专门加这个：插入引用时要知道写到哪一格，
   *   而宿主下发的 ref 里**没有 a1 字段** ⇒ 原实现退化成 "ref.a1 || 'A1'"，
   *   永远写 A1 —— 表现为「点了参数引用，值进了 A1 而不是光标所在格」，
   *   在别的行/列上就等于「没反应」。
   *
   * ★ 难点：Api 只在插件上下文可用，宿主页面调 callCommand 是**静默失效**的
   *   （前面 cellDiag 已实测过），所以「取值」也得走插件回传，不能同步拿。
   *   ⇒ 这里缓存最近一次已知光标地址（由 __trackCursor 持续更新），
   *     插入时读缓存；拿不到才回落 A1。
   *
   * ⚠ 本段在 /poc 的 HTML 模板字符串内，注释里禁止用反引号。
   */
  function __activeCell() {
    var a = window.__ooActiveCell
    if (a && /^[A-Z]+[0-9]+$/.test(a)) return a
    return 'A1'
  }

  /** 把光标地址回传并记到 window.__ooActiveCell，供插入时读取 */
  function __trackCursor(a1) {
    if (!a1 || !/^[A-Z]+[0-9]+$/.test(a1)) return
    window.__ooActiveCell = a1
    try { parent.postMessage({ __ooCursor: a1 }, '*') } catch (e) {}
  }

  /* ★ 2026-10-11 编辑缓冲引擎通道（修复「点= 点参数 无法引入参数」回归）：
   *   旧路径 execCommand('insertText') 依赖选区落点，失败不报错还谎报成功。
   *   sdkjs 实锤：单元格编辑器单例 = AscCommon.Cv（Goi() 创建，fE 帧内全局），
   *   引擎自带追加方法 XJc(t)：fOk(bbf()+t) 写 textarea + wHf({type:'input'})
   *   走完整引擎同步（diff → XWg → nkj → Wb.w_e），与真人键入同路径。
   *   本段在 /poc 的 HTML 模板字符串内：禁反引号、禁 \d、禁美元花括号。 */
  function __cellEditor() {
    var fE = document.querySelector('iframe[name="frameEditor"]')
    var cands = [window, fE ? fE.contentWindow : null]
    for (var i = 0; i < cands.length; i++) {
      var w = cands[i]
      try {
        var A = w && w.AscCommon
        if (A && A.Cv && typeof A.Cv.XJc === 'function' && typeof A.Cv.bbf === 'function') return A.Cv
      } catch (e) {}
    }
    return null
  }
  function __editBuf() {
    var c = __cellEditor()
    if (c) { try { return String(c.bbf() || '') } catch (e) {} }
    // 引擎不可达时回落 DOM 直读（仅供检测，写入勿走此路——不同步引擎）
    try {
      var fE = document.querySelector('iframe[name="frameEditor"]') || document.querySelector('iframe')
      var el = fE && fE.contentDocument ? fE.contentDocument.getElementById('area_id') : null
      return el ? String(el.value || '') : ''
    } catch (e2) { return '' }
  }
  function __appendEditBuf(t) {
    var c = __cellEditor()
    if (!c) return false
    try {
      var before = String(c.bbf() || '')
      c.XJc(String(t))
      var after = String(c.bbf() || '')
      return after.length === before.length + String(t).length && after.slice(-String(t).length) === String(t)
    } catch (e) { return false }
  }
  function __setEditBuf(v) {
    var c = __cellEditor()
    if (!c) return false
    try {
      window.__sgSup = Date.now() // 程序化写入 → 400ms 内抑制自动提示（防选中后旧下拉重弹）
      c.fOk(String(v))
      c.wHf({ type: 'input', preventDefault: function () {}, stopPropagation: function () {} })
      /* ★ 2026-10-11：程序化 fOk 不触发 input 事件 → 插入后下拉判定不重跑、陈旧下拉残留
       *   （12:09 用户截图实锤）→ 主动刷一次；完整名命中精确匹配抑制即关闭 */
      try { __suggestUpdate() } catch (eSU) {}
      return String(c.bbf() || '') === String(v)
    } catch (e) { return false }
  }
  /* 域外引用登记表：rawKey → {row, a1, name}，供「刷新域值」批量重写与状态展示 */
  function __regExtRef(rawKey, row, a1, name) {
    try {
      window.__extRefRegistry = window.__extRefRegistry || {}
      window.__extRefRegistry[rawKey] = { row: row, a1: a1, name: name, ts: Date.now() }
    } catch (e) {}
  }
  /* rawKey（节点.参数）→ 定义名（与插入路径同一套清洗规则） */
  function __sanitizeNm(rawKey) {
    var nm = String(rawKey || '').replace(/[.]/g, '_').replace(/[^0-9A-Za-z_一-龥_]/g, '_')
    if (/^[0-9_]/.test(nm)) nm = 'R' + nm
    return nm
  }

  /* ★ 2026-10-11 输入自动提示（用户指令 ⑥「输入时自动提示」+ ④「带某符号的参数下拉匹配」）：
   *   公式编辑中每次输入 → 取尾部 token 过滤节点参数 → 公式栏下弹 OO 自动完成风格下拉；
   *   token 以 # 开头 = 参数级强制匹配（与 Office $ 锁定相反：# 是开启提示的钥匙），
   *   选中项 = 替换整个 token 为定义名（走 __setEditBuf 引擎同步）。
   *   ① 边界字符：加号 减号 乘号 除号 幂 圆括号 逗号 冒号 空格；# 属 token 内。
   *   ② 域 = frameEditor 文档（公式栏 #ce-cell-content 同域，坐标精确）。
   *   ③ 程序化写缓冲（选参/替换）后 400ms 抑制（__sgSup），防刚选完又弹。
   *   ④ 注释禁写斜星乘幂连写（2026-10-11 惨案：加乘除幂里的星斜杠提前终止块注释）。 */
  var __sgEl = null
  function __suggestClose() {
    if (__sgEl) { try { __sgEl.remove() } catch (e) {} __sgEl = null }
  }
  function __suggestUpdate() {
    try {
      if (Date.now() - (window.__sgSup || 0) < 400) { __suggestClose(); return }
      var buf = __editBuf()
      if (!buf || buf.indexOf('=') < 0) { __suggestClose(); return }
      var m = buf.match(/[一-龥A-Za-z0-9_#.]+$/)
      var tok = m ? m[0] : ''
      var isHash = tok.charAt(0) === '#'
      var filter = isHash ? tok.slice(1) : tok
      var fl = filter.toLowerCase()
      if (buf.indexOf('=') === 0 && !tok) { __suggestClose(); return } // 裸=不打扰
      var items = []
      ;(__refItems || []).forEach(function (it) {
        (it.fields || []).forEach(function (f) {
          var combined = (it.name || '') + '_' + (f.name || '')
          var cl = combined.toLowerCase()
          var hit = isHash ? (cl.indexOf(fl) >= 0 || String(f.name || '').toLowerCase().indexOf(fl) >= 0)
            : (cl.indexOf(fl) >= 0 || String(it.name || '').toLowerCase().indexOf(fl) >= 0)
          if (hit) items.push({ nm: __sanitizeNm(combined), label: combined, val: f.value })
        })
      })
      if (!items.length) { __suggestClose(); return }
      /* ★ 2026-10-11（用户报「不应该出现两个参数引用」）：双击选参后缓冲已含完整参数名，
       *   尾部 token 与某候选完全一致 = 无可补全 → 不弹（否则选完又弹第二个引用列表） */
      if (!isHash && items.some(function (en) { return en.nm === tok || en.label === tok })) { __suggestClose(); return }
      var fE = document.querySelector('iframe[name="frameEditor"]')
      var edoc = fE && fE.contentDocument ? fE.contentDocument : null
      if (!edoc) { __suggestClose(); return }
      __suggestClose()
      var dd = edoc.createElement('div')
      dd.id = 'ext-suggest'
      dd.__extOwned = true
      var maxShow = Math.min(items.length, 30)
      for (var i = 0; i < maxShow; i++) {
        (function (en) {
          var item = edoc.createElement('div')
          item.style.cssText = 'display:flex;align-items:center;gap:8px;padding:5px 12px;cursor:pointer;color:#e8e8e8;font-size:13px;background:#2b2b2b'
          item.innerHTML = '<span style="display:inline-flex;width:18px;height:18px;align-items:center;justify-content:center;border:1px solid #6e7681;border-radius:4px;font-size:10px;color:#9ab;flex:none">fx</span>' +
            '<span style="font-weight:600;white-space:nowrap">' + en.label + '</span>' +
            (en.val != null && en.val !== '' ? '<span style="color:#8b949e;margin-left:auto;font-size:11px;max-width:110px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + String(en.val).slice(0, 16) + '</span>' : '')
          item.onmouseenter = function () { item.style.background = '#0d419d' }
          item.onmouseleave = function () { item.style.background = '#2b2b2b' }
          var pick = function (eP) {
            try { eP.preventDefault(); eP.stopPropagation() } catch (e0) {}
            if (dd.__picked) return // mousedown+click 双事件只生效一次
            dd.__picked = true
            var nb = buf.slice(0, buf.length - tok.length) + en.nm
            __setEditBuf(nb)
            try { parent.postMessage({ __ooStatus: '[suggest] 已替换 token → ' + en.nm + '（编辑保持，回车确认）' }, '*') } catch (e1) {}
            __suggestClose()
          }
          item.addEventListener('mousedown', pick)
          item.addEventListener('click', pick)
          dd.appendChild(item)
        })(items[i])
      }
      dd.style.cssText = 'position:absolute;z-index:10000;background:#2b2b2b;border:1px solid #454545;border-radius:6px;box-shadow:0 6px 16px rgba(0,0,0,.45);padding:4px 0;min-width:240px;max-height:300px;overflow:auto'
      var fb = edoc.getElementById('ce-cell-content')
      var r = fb ? fb.getBoundingClientRect() : { left: 8, bottom: 120 }
      dd.style.left = Math.max(4, r.left) + 'px'
      dd.style.top = (r.bottom + 3) + 'px'
      dd.addEventListener('mousedown', function (eD) { eD.stopPropagation() })
      ;(edoc.body || edoc.documentElement).appendChild(dd)
      __sgEl = dd
    } catch (eU) {}
  }
  function __wireSuggest() {
    try {
      var fE = document.querySelector('iframe[name="frameEditor"]')
      var edoc = fE && fE.contentDocument ? fE.contentDocument : null
      var el = edoc ? edoc.getElementById('area_id') : null
      if (!el || el.__sgHooked) return
      el.__sgHooked = true
      el.addEventListener('input', function () { __suggestUpdate() })
      el.addEventListener('keydown', function (eK) {
        if (eK.key === 'Enter' || eK.key === 'Escape' || eK.keyCode === 13 || eK.keyCode === 27) __suggestClose()
      })
    } catch (e) {}
  }
  // 调试句柄（无头 e2e / 控制台排查用）：编辑缓冲引擎通道与自动提示探针
  try {
    window.__dbgEditBuf = {
      cellEditor: __cellEditor,
      editBuf: __editBuf,
      append: __appendEditBuf,
      set: __setEditBuf,
      suggest: __suggestUpdate,
      registry: function () { return window.__extRefRegistry || {} },
    }
  } catch (eDbg) {}
  /* ★ 2026-10-11 编辑结束自愈（修「点= 点参数 回车后格子没值」的兜底保险）：
   *   登记后周期巡检：编辑缓冲清空 = 用户回车提交/取消；
   *   等 800ms 让提交落地 → 插件读目标格：值为空 或 仍是文本 '=参数名'
   *   ⇒ 引擎提交链路丢缓冲 ⇒ 条件补写引用公式（含引用格样式）。
   *   引擎正常提交（格子里是公式/值）则零动作零痕迹。 */
  function __editVerifyTick() {
    var v = window.__editVerify
    if (!v) return
    if (__editBuf()) { v.ended = 0; return } // 仍在编辑（或开了新编辑）
    if (!v.ended) v.ended = Date.now()
    if (Date.now() - v.ended < 400) return // 提交落地缓冲期（首联提速：800→400）
    if (Date.now() - v.t > 20000) { window.__editVerify = null; return }
    window.__editVerify = null
    // ⚠ 插件 executeCommand 脚本禁顶层 return（会被判解析错误整段静默不执行——
    //   无头 probe7 实锤连 return 前的落格语句都不跑）。补写走 SetValue('=名')：
    //   带等号字符串会被公式引擎解析；实测 SetFormula 部分版本静默无效（probe4）。
    var sc = 'try{var ws=Api.GetActiveSheet();var r=ws.GetRange("' + v.a1 + '");var val=String(r.GetValue()==null?"":r.GetValue());var isErr=(typeof val==="string"&&val.charAt(0)==="#");if(val===""||val==="=' + v.name + '"||isErr){r.SetValue("=' + v.name + '");try{var rc=ws.GetRange("' + v.a1 + '");rc.SetBorders("all",Api.CreateStroke(6350,Api.CreateSolidFill(Api.CreateRGBColor(88,166,255))));rc.SetFillColor(Api.CreateRGBColor(232,244,255));rc.SetFontColor(Api.CreateRGBColor(9,105,218));rc.SetBold(true)}catch(eM){}}try{var wb2=ws.GetParent?ws.GetParent():null;if(wb2&&wb2.RecalculateAllFormulas)wb2.RecalculateAllFormulas();else if(Api.RecalculateAllFormulas)Api.RecalculateAllFormulas()}catch(eC4){}}catch(e){}'
    __sendToRefPlugin({ __cbRefCmd: 'script', text: sc })
    try { parent.postMessage({ __ooStatus: '[verify] 编辑已结束，自检 ' + v.a1 + '（空/文本形态则补写 =' + v.name + '）' }, '*') } catch (eV2) {}
  }


  function __insertRefFormula(ref) {
    if (!ref || !ref.nodeId) return
    var isSheet = String(window.__ooKind || '') === 'oosheet'
    var isPpt = String(window.__ooKind || '') === 'ooppt'
    var fE = document.querySelector('iframe[name="frameEditor"]')
    /* ★ 引擎窗口实测（2026-10-09 cellDiag 实锤）：sheet 引擎（window.editor，1844 个
     *   asc 方法、callCommand=function）就在本页（/poc）——
     *   frameEditor.contentWindow.editor 反而是无、DocsAPI 也是 undefined。
     *   查找顺序改：本 window.editor 优先，frameEditor 兜底。 */
    var api = window.editor || (fE && fE.contentWindow ? (fE.contentWindow.editor || (fE.contentWindow.Asc && fE.contentWindow.Asc.editor) || null) : null)
    // word 文本流路径（asc_AddText）必须先有引擎；sheet 主路径走 cbref 插件
    // （callCommand 只是兜底），引擎缺失不应阻断插件写入 —— 不再一刀切 return
    if (!api && !isSheet) {
      try { parent.postMessage({ __ooStatus: 'insert: 编辑器引擎未就绪（window.editor / frameEditor.editor 均缺失）' }, '*') } catch (e) {}
      return
    }
    var nodeName = ref.nodeName || ref.nodeId
    var fieldName = ref.fieldName || ''
    // 公式文本：字段名为空 → 引用整节点（Excel 里就是区域引用）
    var formula = fieldName ? '=' + nodeName + '.' + fieldName : '=' + nodeName
    /* ★ 日志链①：点击参数引入（含原生标记前态快照，后续环节全部出「前态→后态」差异） */
    try {
      __mkLog('REF-INSERT', '点击参数引入：' + nodeName + '.' + fieldName +
        ' 值=' + (ref.value == null ? '(空)' : String(ref.value)) +
        ' a1=' + (ref.a1 || '(随光标)') +
        ' 前态[' + __nfSnapStr(__nfSnap()) + ']')
    } catch (eIR) {}
    try { if (fE && fE.contentWindow) fE.contentWindow.focus() } catch (e) {}

    /* ★ 表格与文档走完全不同的 API（这是「参数进不了单元格」的根因）：
       oosheet —— asc_AddText 是文档文本流 API，对单元格无效。写单元格要用 Office JS API
                  （Api.GetActiveSheet().GetCells(i,j).SetValue(v)），而它**只能在插件上下文
                  里执行**：宿主页面直接调 api.callCommand 静默失效（调研已确认）。
                  故表格侧统一转给 cbref 插件执行，见__sendToRefPlugin。
       ooword—— asc_AddText 正确（往正文流里插字）。 */
    /* ★ 写入前实时索取光标（2026-10-10 app 实测二修）：名称框选区事件在 app 环境
     *   未触发（已加埋点追查死因），插件 applyCursor 读的是引擎活体选区、实测可靠
     *   ——无条件问一次，450ms 让回执爬链落地（__cbRefHost 锚点 → /poc 缓存），
     *   再用最新光标写入。 */
    function __doSheetWrite(targetA1, formula, refValue, refName, mode) {
      var isSheetW = mode === 'ref'
      var fE = document.querySelector('iframe[name="frameEditor"]')
      var api = window.editor || (fE && fE.contentWindow ? (fE.contentWindow.editor || (fE.contentWindow.Asc && fE.contentWindow.Asc.editor) || null) : null)
      // 先打=兼容：关编辑态提交缓冲，插件写入随后覆盖为引用公式
      try {
        if (api && typeof api.asc_closeCellEditor === 'function') {
          api.asc_closeCellEditor(false)
          try { parent.postMessage({ __ooStatus: 'insert: 已退出单元格编辑态（缓冲中的 = 由引用公式接管）' }, '*') } catch (e0c) {}
        }
      } catch (e0d) {}
      /* ★ 2026-10-10 定义名称活引用（无头实证 B10 解析 4200）：ref 模式走
       *   ① 隐藏『域数据』表 A<row>=画布现值（行号 djb2(参数名) 确定性映射 2..1999）
       *   ② 定义名称 upsert：参数名 => '域数据'!$A<row>（延迟 600ms 等表落地）
       *   ③ 单元格写公式 =参数名 —— 公式栏显示参数名，格子显示其值；
       *   画布值变 → 重点击/刷新域值 → A<row> 更新 → 全表自动重算（真·活引用）。
       *   值模式（非 ref）：直接写画布现值，所见即所得。 */
      var rawKey = String(refName || '') || ('引用_' + targetA1)
      var name = isSheetW ? rawKey.replace(/[.]/g, '_').replace(/[^0-9A-Za-z_一-龥_]/g, '_') : ''
      if (isSheetW && /^[0-9_]/.test(name)) name = 'R' + name
      var h = 5381
      for (var hi = 0; hi < rawKey.length; hi++) { h = ((h << 5) + h + rawKey.charCodeAt(hi)) & 0x7fffffff }
      var row = 2 + (h % 1998)
      if (isSheetW) __regExtRef(rawKey, row, targetA1, name)
      var val = (refValue != null && String(refValue) !== '') ? String(refValue) : ''
      var numVal = parseFloat(val)
      var cellVal = isNaN(numVal) ? JSON.stringify(val) : String(numVal)
      if (isSheetW) {
        /* ★ 2026-10-11 引用格样式升级（用户指令「颜色样式升级」）：域外链接参数
         *   字体蓝 #0969DA + 加粗（Office 跨表引用同款蓝），保留淡蓝底 + 蓝边框。 */
        __sendToRefPlugin({ __cbRefCmd: 'script', text: 'try{var sheets=Api.GetSheets?Api.GetSheets():[];var dataWs=null;for(var i=0;i<sheets.length;i++){var nm2=sheets[i].GetName?sheets[i].GetName():"";if(nm2==="域数据"){dataWs=sheets[i];break}}if(!dataWs&&Api.AddSheet)dataWs=Api.AddSheet("域数据");if(!dataWs)throw"no data sheet";try{dataWs.SetVisible&&dataWs.SetVisible(false)}catch(eH){}dataWs.GetRange("A' + row + '").SetValue(' + cellVal + ');dataWs.GetRange("B' + row + '").SetValue(' + JSON.stringify(rawKey) + ');try{var rc=Api.GetActiveSheet().GetRange("' + targetA1 + '");rc.SetBorders("all",Api.CreateStroke(6350,Api.CreateSolidFill(Api.CreateRGBColor(88,166,255))));rc.SetFillColor(Api.CreateRGBColor(232,244,255));rc.SetFontColor(Api.CreateRGBColor(9,105,218));rc.SetBold(true)}catch(eM){}}catch(eS){}' })
        setTimeout(function () {
          /* ★ 2026-10-11（#NAME? 排查）：定义名注册改走 cbref 插件通道 ——
           *   applyDefName 用 wb.AddDefName 在【文档上下文】执行（与已实证可用的写值同通道），
           *   并回传真实结果 ok/why（__cbRefDefName）。此前宿主侧 asc_editDefinedNames
           *   日志报"已注册"但引擎求值疑似未生效（=名 落格 #NAME? 的头号嫌疑）。 */
          /* ★ 2026-10-11 首联提速：域数据表已在文档就绪时预建（插件侧注册前还会二次确保），
           *   300ms 等表落地已无意义 → 100ms 让注册与点击几乎同时出发 */
          __sendToRefPlugin({ __cbRefCmd: 'defname', name: name, ref: "'域数据'!$A$" + row, a1: targetA1 })
        }, 100)
      }
      // ★ 写真值优先（2026-10-10 用户问『为什么是 #NAME?』）：v1 伪引用标记会被
      //   Excel 解析成未知名字 => #NAME?。参数点击自带画布现值——直接写真值，
      //   所见即所得；引用语义由画布 ooRefs/边承担（值变化走刷新域值重刷）。
      // ★ 单元格写入：ref 模式写公式 =参数名（公式栏显示参数名，格子显示解析值）；
      //    值模式直接写画布现值，所见即所得。
      var v = isSheetW ? ('=' + name) : ((refValue != null && String(refValue) !== '') ? String(refValue) : formula)
      if (isSheetW) {
        // ★ ref 模式严格时序（2026-10-10 用户实测「时好时坏」根因修复）：
        //   公式必须晚于定义名称注册落格——否则名称不存在时算出 #NAME? 并缓存，
        //   下一个编辑触发重算才解析（用户所见「编辑下一个会变好」）。
        //   时序：域数据值(0ms) → 定义名称(600ms) → 公式写入(900ms) → 面板自愈(1200ms)
        setTimeout(function () {
          try {
            __sendToRefPlugin({ cells: [{ a1: targetA1, value: v, isFormula: true }], recalc: true })
            try { parent.postMessage({ __ooStatus: '[cbref] 已走插件通道写值（活引用） ' + targetA1 + ' = ' + v }, '*') } catch (e0) {}
          } catch (eW) { try { parent.postMessage({ __ooStatus: '[cbref] 公式写入异常 ' + String(eW).slice(0, 60) }, '*') } catch (eW2) {} }
          /* ★ 2026-10-11 用户指令「双击参数会打开参数引用面板，去掉」：
           *   原此处 300ms 后自动 __setRefPanelVisible(true)（所谓"面板自愈"）——
           *   用户不需要插入引用时自动弹面板。删除。面板只由 数据引用按钮/× 手动控制。 */
        }, 600)
        return
      }
      // 值模式（非 ref）：直接写画布现值，所见即所得
      var sent = __sendToRefPlugin({ cells: [{ a1: targetA1, value: v, isFormula: v.charAt(0) === '=' }], recalc: true })
      if (sent) {
        try { parent.postMessage({ __ooStatus: '[cbref] 已走插件通道写值 ' + targetA1 + ' = ' + v }, '*') } catch (e0) {}
        return
      }
      if (api && typeof api.callCommand === 'function') {
        try {
          var wv = (refValue != null && String(refValue) !== '') ? String(refValue) : formula
          var cmdFn = new Function('var a1=' + JSON.stringify(targetA1) + ',fv=' + JSON.stringify(wv) + ';try{var sh=Api.GetActiveSheet();var r=sh.GetRange(a1);r.SetValue(fv);try{r.SetFormula(fv)}catch(e0){}return "ok"}catch(e2){return "err:"+String(e2).slice(0,80)}')
          var __directDone = false
          api.callCommand(cmdFn, false, false, function (res) {
            __directDone = true
            try { parent.postMessage({ __ooStatus: '[direct] callCommand 回调 res=' + String(res).slice(0, 60) }, '*') } catch (e6) {}
          })
          setTimeout(function () {
            if (!__directDone) try { parent.postMessage({ __ooStatus: 'insert: 直连 callCommand 3s 无回调（免费版无插件会话不执行）' }, '*') } catch (e7) {}
          }, 3000)
          try { parent.postMessage({ __ooStatus: '[direct] 兜底下发 ' + targetA1 + ' = ' + formula }, '*') } catch (e5) {}
          return
        } catch (e10) {
          try { parent.postMessage({ __ooStatus: '[direct] callCommand 抛错 ' + String(e10).slice(0, 60) }, '*') } catch (e11) {}
          return
        }
      }
      try { parent.postMessage({ __ooStatus: 'insert: 直连与插件通道均不可用' }, '*') } catch (e4) {}
    }

    if (isSheet) {
      /* ★ 光标三级来源（2026-10-10 终版）：① 名称框 DOM 实况（同步、永不过期、
       *   无事件依赖——web-apps 自己就在实时更新它）；② __ooActiveCell 缓存（事件链）；
       *   ③ 保底 A1。 */
      try { parent.postMessage({ __ooStatus: 'insert: 读取名称框实况确定目标格…' }, '*') } catch (e0b) {}
      setTimeout(function () {
        var live = ''
        try {
          var fE0 = document.querySelector('iframe[name="frameEditor"]')
          var el = fE0 && fE0.contentDocument ? fE0.contentDocument.getElementById('ce-cell-name') : null
          live = el ? String(el.value || '').trim() : ''
        } catch (e) { live = '' }
        // ★ 正则用 [0-9] 不用 \d：本页是模板字面量产物，\d 求值时反斜杠被吞成 d（历史惨案同款）
        try { parent.postMessage({ __ooStatus: 'insert: 名称框读数 live=[' + live + '] active=[' + (window.__ooActiveCell || '') + '] refa1=[' + (ref.a1 || '') + ']' }, '*') } catch (e) {}
        if (!/^[A-Z]+[0-9]+$/.test(live)) live = window.__ooActiveCell || ''
        var targetA1 = /^[A-Z]+[0-9]+$/.test(live) ? live : (ref.a1 || 'A1')
        var rawKey = (ref.nodeName || '') + '.' + (ref.fieldName || '')
        var nm = rawKey.replace(/[.]/g, '_').replace(/[^0-9A-Za-z_一-龥_]/g, '_')
        if (/^[0-9_]/.test(nm)) nm = 'R' + nm
        var rowH = 2 + ((function (s0) { var hh = 5381; for (var q = 0; q < s0.length; q++) { hh = ((hh << 5) + hh + s0.charCodeAt(q)) & 0x7fffffff } return hh })(rawKey) % 1998)
        var valS = (ref.value != null && String(ref.value) !== '') ? String(ref.value) : ''
        var numV = parseFloat(valS)
        var cellVS = isNaN(numV) ? JSON.stringify(valS) : String(numV)
        /* ★ 2026-10-11 编辑态检测 v2（修「点= 点参数 无法引入参数」回归）：
         *   旧路径 execCommand('insertText') 依赖选区，失败也不报错还谎报成功。
         *   新路径走引擎单例 Cv.XJc（fOk+wHf 全引擎同步，与真人键入同路），
         *   每步读回验证；同节点/尾部残 token 点击 = 替换（Office 自动完成语义），
         *   全部失败才回落直写引用公式。 */
        var editBuf = __editBuf()
        if (editBuf.length > 0) {
          __regExtRef(rawKey, rowH, targetA1, nm)
          __sendToRefPlugin({ __cbRefCmd: 'script', text: 'try{var sheets=Api.GetSheets?Api.GetSheets():[];var dataWs=null;for(var i=0;i<sheets.length;i++){var nm2=sheets[i].GetName?sheets[i].GetName():"";if(nm2==="域数据"){dataWs=sheets[i];break}}if(!dataWs&&Api.AddSheet)dataWs=Api.AddSheet("域数据");if(!dataWs)throw"no data sheet";try{dataWs.SetVisible&&dataWs.SetVisible(false)}catch(eH){}dataWs.GetRange("A' + rowH + '").SetValue(' + cellVS + ');dataWs.GetRange("B' + rowH + '").SetValue(' + JSON.stringify(rawKey) + ')}catch(eS){}' })
          setTimeout(function () {
            /* ★ 同上：定义名注册改走插件通道（文档上下文 wb.AddDefName + 真实回执 ok/why） */
            __sendToRefPlugin({ __cbRefCmd: 'defname', name: nm, ref: "'域数据'!$A$" + rowH, a1: targetA1, editing: 1 })
          }, 100)
          // 尾部 token 判定：残 token（含旧引用）= 替换；操作符/空 = 追加；同名 = 跳过
          var mTok = editBuf.match(/[一-龥A-Za-z0-9_.]+$/)
          var tok = mTok ? mTok[0] : ''
          var done = false
          if (tok === nm) {
            try { parent.postMessage({ __ooStatus: '[insert] ' + nm + ' 已在公式中，跳过重复插入' }, '*') } catch (eS0) {}
            done = true
          } else {
            var newBuf = (tok && tok !== '=') ? editBuf.slice(0, editBuf.length - tok.length) + nm : editBuf + nm
            var how = tok ? '替换[' + tok + '→' + nm + ']' : '追加' + nm
            done = __setEditBuf(newBuf)
            if (done) {
              try { parent.postMessage({ __ooStatus: '[insert] 编辑缓冲' + how + '成功（engine 通道，编辑保持，回车确认）' }, '*') } catch (eA2) {}
            } else {
              // 回落1：execCommand（选区路径，读回验证）
              try {
                fE.contentWindow.focus()
                fE.contentDocument.execCommand('insertText', false, nm)
              } catch (eX) {}
              var bufX = __editBuf()
              if (bufX.indexOf(nm) >= 0 && bufX !== editBuf) {
                done = true
                try { parent.postMessage({ __ooStatus: '[insert] 编辑缓冲' + how + '成功（execCommand 通道）' }, '*') } catch (eA3) {}
              }
            }
          }
          if (done) {
            // ★ 自愈登记（2026-10-11）：编辑结束（回车）后自检目标格——若空或仍是文本
            //   形态 =参数名，说明引擎提交链路丢了缓冲（无头对照实锤此风险存在），
            //   由 __editVerifyTick 走插件条件补写引用公式。正常提交则零动作。
            try { window.__editVerify = { a1: targetA1, name: nm, t: Date.now(), ended: 0 } } catch (eV) {}
            return
          }
          // 回落2（最后手段）：引擎/选区均不可达 → 退出编辑直接写引用公式
          try { parent.postMessage({ __ooStatus: 'insert: 编辑态引擎不可达（Cv 缺失且 execCommand 无效），改直写引用公式' }, '*') } catch (eA4) {}
          __doSheetWrite(targetA1, formula, ref.value, rawKey, 'ref')
          return
        }
        __doSheetWrite(targetA1, formula, ref.value, rawKey, ref.mode || 'value')
      }, 0)
      return
    }

    /* ★ ppt（2026-10-10 配齐第三类型）：当前幻灯片插入引用文本形状（直连引擎 callCommand）。
     *   Office JS：Api.GetPresentation().GetCurrentSlide() → CreateShape(rect, 8cm×1cm EMU)
     *   → GetDocContent().GetElement(0) → CreateRun().SetText(引用文本) → AddObject。
     *   sdkjs 本地 bundle 已核实 GetCurrentSlide/CreateShape/"rect" 均存在。 */
    if (isPpt) {
      // ★ 主通道 = cbref 插件 insertShape（插件会话内 callCommand 有回调——sheet 同款实证）
      var sentP = __sendToRefPlugin({ __cbRefCmd: 'insertShape', text: formula })
      if (sentP) {
        try { parent.postMessage({ __ooStatus: '[cbref-ppt] 已走插件通道插形状 = ' + formula }, '*') } catch (e13) {}
        return
      }
      if (api && typeof api.callCommand === 'function') {
        try {
          var pCmdFn = new Function('var fv=' + JSON.stringify(formula) + ';try{var p=Api.GetPresentation();var s=p.GetCurrentSlide()||p.GetSlideByIndex(0);if(!s)return "no-slide";var shp=Api.CreateShape("rect",2880000,360000,Api.CreateFill(Api.CreateRGBColor(235,241,250),0),Api.CreateStroke(0,Api.CreateNoFill()));var para=shp.GetDocContent().GetElement(0);var run=Api.CreateRun();run.SetText(fv);try{run.SetColor(31,111,235)}catch(e0){}para.AddElement(run);s.AddObject(shp);return "ok"}catch(e2){return "err:"+String(e2).slice(0,80)}')
          var __pDone = false
          api.callCommand(pCmdFn, false, false, function (res) {
            __pDone = true
            try { parent.postMessage({ __ooStatus: '[direct-ppt] 回调 res=' + String(res).slice(0, 60) + '（ok=当前页已插引用块）' }, '*') } catch (e6) {}
          })
          setTimeout(function () {
            if (!__pDone) try { parent.postMessage({ __ooStatus: 'insert: ppt 直连 3s 无回调（免费版限制），且插件不可达' }, '*') } catch (e7) {}
          }, 3000)
          try { parent.postMessage({ __ooStatus: '[direct-ppt] 兜底下发 当前页引用块 = ' + formula }, '*') } catch (e5) {}
          return
        } catch (e10) {
          try { parent.postMessage({ __ooStatus: '[direct-ppt] callCommand 抛错 ' + String(e10).slice(0, 60) }, '*') } catch (e11) {}
          return
        }
      }
      try { parent.postMessage({ __ooStatus: 'insert: ppt 无可用引擎（editor|Asc.editor 均缺）' }, '*') } catch (e12) {}
      return
    }

    // ★ ooword 域外关联参数（2026-10-22 用户指令）：插「文本表单」（=sdt 内容控件，
    //   w:tag=key）而不是裸文本 —— 裸文本没有落点，「刷新域值」的 patchDocxSdt
    //   找不到写回目标（Word 侧域写回因此一直没闭环）。表单 key=字段名，
    //   与 binder 约定 field:字段名 同口径（toWordSites 会剥前缀对齐）。
    //   插入后占位符处于选中态 → asc_AddText 写画布现值（所见即所得）；
    //   上游参数变 → 刷新域值 → 按 tag 补丁写回（格式零丢失）。
    var formKey = fieldName || nodeName
    var inserted = false
    try {
      if (typeof api.InsertTextForm === 'function') {
        var sdtForm = api.InsertTextForm({
          key: formKey,
          placeholder: '[' + nodeName + '.' + (fieldName || formKey) + '] 画布参数域',
          keepSelectedTextInForm: false
        })
        inserted = true
        /* ★ 2026-10-22 用户「与文档格式一致 能编辑字体字高段落 不影响显示效果」：
           appearance=hidden —— 内容控件视觉壳（表单底纹/边框）去掉，
           值以普通文本渲染：字体/字高/段落完全继承光标处正文格式，
           选中即可改字体字高段落，显示效果与正文零差异；
           tag 落点保留（刷新域值照常写回，功能不受影响）。 */
        try {
          if (sdtForm && typeof sdtForm.SetAppearance === 'function') {
            sdtForm.SetAppearance('hidden')
          } else if (typeof api.asc_SetGlobalContentControlShowHighlight === 'function') {
            api.asc_SetGlobalContentControlShowHighlight(false)
          }
        } catch (eAp) {
          try { parent.postMessage({ __ooStatus: '[域外参数] appearance 隐藏失败：' + String(eAp).slice(0, 60) }, '*') } catch (eAp2) {}
        }
        __regRefMeta(ref)
        var v1 = (ref.value != null && String(ref.value) !== '') ? String(ref.value) : ''
        setTimeout(function () {
          try { if (v1 !== '' && typeof api.asc_AddText === 'function') api.asc_AddText(v1) } catch (eAT) {}
          try { parent.postMessage({ __ooStatus: '[域外参数] 已插入文本表单 key=' + formKey + '（刷新域值可写回）' }, '*') } catch (eST) {}
        }, 100)
      }
    } catch (eIF) {}
    if (inserted) return
    // 兜底（编辑器不支持表单时）：旧路径裸文本 + 淡色高亮
    try { if (typeof api.SetHighlight === 'function') api.SetHighlight('lightGray') } catch (e5) {}
    try {
      if (typeof api.asc_AddText === 'function') api.asc_AddText(formula)
      else return
    } catch (e3) {}
    try { if (typeof api.SetHighlight === 'function') api.SetHighlight('none') } catch (e4) {}
    // 胶囊登记：与双击路径共用同一批胶囊，按 node#field 去重
    try {
      window.__extRefCapsules = window.__extRefCapsules || []
      var key = ref.nodeId + '#' + (ref.fieldId || '')
      var exists = null
      window.__extRefCapsules.forEach(function (c) { if (c.key === key) exists = c })
      var cw = formula.length * 8 + 14
      if (!exists) {
        exists = {
          key: key, w: cw, h: 22,
          value: ref.value != null ? ref.value : '', fieldName: fieldName,
          isFormula: true, nodeName: nodeName,
        }
        window.__extRefCapsules.push(exists)
      } else {
        // 已存在则刷新宽度与值（重算后同一个胶囊要跟着变宽/变值）
        exists.w = cw
        if (ref.value != null) exists.value = ref.value
      }
      // ★ 定位到光标处：没有 docX/docY 这颗胶囊点不到、也画不出连线
      exists.docX = window.__cursorDoc ? window.__cursorDoc[0] : 40
      exists.docY = window.__cursorDoc ? window.__cursorDoc[1] : 100
      exists.rowKey = key
    } catch (e5) {}
    // 画连线（插入后立即有指向节点的连线）
    try { window.__drawRefLines && window.__drawRefLines() } catch (e6) {}
  }

  /* ── ★ OO→PDF 转换会话（2026-10-22 用户指令「多个文档转pdf并连接」）──
   * 复用当前打开的 OO 编辑器引擎会话：触发「下载为 PDF」，拦截下载 Blob
   * （hook a[download].click，blob: URL fetch 回读）→ base64 回传宿主。
   * ★ 前提：宿主先把目标文档内容加载进当前编辑器（合并节点逐份切换）。
   *   宿主侧调用序：__ooLoadBytes(bytes) → __ooConvertPdf() → 循环下一份。 */
  window.__ooLoadBytes = function (bytes) {
    return new Promise(function (resolve) {
      try {
        var f = new File([bytes], 'convert.docx', { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
        /* 复用宿主打开链：/poc 自己不实现，交回宿主（OnlyOfficeNodeEditor.openFile）
         * ——宿主在调 __ooConvertPdf 前自行完成打开并等 documentReady。 */
        resolve({ ok: true })
      } catch (eL) { resolve({ ok: false, why: String(eL) }) }
    })
  }
  window.__ooConvertPdf = function () {
    return new Promise(function (resolve) {
      try {
        var fE = document.querySelector('iframe[name="frameEditor"]')
        var api = window.editor || (fE && fE.contentWindow ? fE.contentWindow.editor : null)
        if (!api) { resolve({ ok: false, why: '编辑器未打开——请先双击打开任一 OO 文档（转换复用其引擎会话）' }); return }
        var hooked = false
        var origClick = HTMLAnchorElement.prototype.click
        var done = false
        var finish = function (b64) {
          if (done) return
          done = true
          if (hooked) { HTMLAnchorElement.prototype.click = origClick; hooked = false }
          resolve(b64 ? { ok: true, b64: b64 } : { ok: false, why: '未捕获到 PDF 下载流（25s 超时）' })
        }
        var blobToB64 = function (blob, cb) {
          var fr = new FileReader()
          fr.onload = function () { cb(String(fr.result || '').split(',')[1] || '') }
          fr.onerror = function () { cb('') }
          fr.readAsDataURL(blob)
        }
        HTMLAnchorElement.prototype.click = function () {
          try {
            var dn = String(this.getAttribute('download') || '')
            var href = String(this.getAttribute('href') || '')
            if (/\.pdf$/i.test(dn) && href.indexOf('blob:') === 0) {
              fetch(href).then(function (r) { return r.blob() }).then(function (b) { blobToB64(b, finish) }).catch(function () { finish('') })
              return
            }
          } catch (eHk) {}
          return origClick.apply(this, arguments)
        }
        hooked = true
        setTimeout(function () { finish('') }, 25000)
        var w2 = (fE && fE.contentWindow) || window
        /* 触发下载为 PDF：优先 web-apps 控制器（file:saveas 消息 → asc_DownloadAs），
         * 其次引擎直调。两者都不可达时报错（用户按提示手动另存）。 */
        var fired = false
        try {
          if (w2 && typeof w2.postMessage === 'function' && w2 !== window) {
            /* frameEditor 是引擎 canvas 容器不是 web-apps；web-apps 在外层。改向自己发 file:saveas */
          }
        } catch (eP) {}
        try {
          if (typeof api.asc_DownloadAs === 'function') { api.asc_DownloadAs(); fired = true }
        } catch (eD1) {}
        if (!fired) {
          /* web-apps 主控在 window（同页）——直接派发内部消息事件 */
          try {
            var evt = new MessageEvent('message', { data: 'file:saveas', source: window })
            window.dispatchEvent(evt)
            fired = true
          } catch (eD2) {}
        }
        if (!fired) { finish(''); return }
      } catch (eCv) {
        resolve({ ok: false, why: String(eCv).slice(0, 80) })
      }
    })
  }

  /* ── ★ 域外参数悬浮提示（2026-10-22 用户指令「鼠标悬浮显示节点/参数/格式」）──
     * 引擎对 sdt 无 tooltip 渲染 API（SetLabel/SetAlias 是纯元数据）。
     * 方案：插入域时登记元数据（tag→节点/参数/格式/值）；文档页 mousemove
     * 时用 caretRangeFromPoint 找命中文本，扫所在 sdt（WebForms 渲染为带
     * class 的 span），命中即画自研 tooltip（与引用格标记同款覆盖层手法）。 */
  window.__ooRefMeta = window.__ooRefMeta || {}
  function __regRefMeta (ref) {
    try {
      if (!ref || !ref.nodeId) return
      var k = String(ref.fieldName || ref.fieldId || '')
      window.__ooRefMeta[k] = {
        node: String(ref.nodeName || ref.nodeId || ''),
        field: k,
        value: ref.value != null ? String(ref.value) : '',
        fmt: (ref.fmt != null && String(ref.fmt) !== '') ? String(ref.fmt) : '自动',
        t: Date.now()
      }
    } catch (eM) {}
  }
  function __refTooltipEl (doc, create) {
    var el = doc.getElementById('__cb_ref_tip')
    if (!el && create) {
      el = doc.createElement('div')
      el.id = '__cb_ref_tip'
      el.style.cssText = 'position:absolute;z-index:99999;pointer-events:none;display:none;'
        + 'background:rgba(13,17,23,.94);color:#e8edf2;border:1px solid #30363d;'
        + 'border-radius:6px;padding:7px 10px;font:11.5px/1.6 system-ui,sans-serif;'
        + 'max-width:280px;box-shadow:0 4px 14px rgba(0,0,0,.35);white-space:nowrap;'
      el.innerHTML = ''
      doc.body.appendChild(el)
    }
    return el
  }
  function __refTipPaint (doc, x, y, meta) {
    var el = __refTooltipEl(doc, true)
    if (!el) return
    var row = function (label, val, strong) {
      return '<div><span style="color:#6e7681">' + label + '：</span>'
        + '<span style="' + (strong ? 'font-weight:600;color:#58a6ff' : '') + '">' + val + '</span></div>'
    }
    el.innerHTML =
      row('节点', meta.node, true) +
      row('参数', meta.field) +
      row('当前值', meta.value === '' ? '(空)' : meta.value) +
      row('格式', meta.fmt)
    el.style.display = 'block'
    // 跟随光标（右下偏移，贴边翻转）
    var vw = doc.defaultView ? doc.defaultView.innerWidth : 1200
    var w = el.offsetWidth || 200
    var nx = (x + 16 + w > vw - 8) ? Math.max(8, x - w - 12) : x + 16
    el.style.left = nx + 'px'
    el.style.top = (y + 18) + 'px'
  }
  function __refTipScan (ev) {
    try {
      var t = __hiFrame()
      if (!t || !t.doc) return
      var doc = t.doc
      // 只在 word 文档页处理
      if (String(window.__ooKind || '') !== 'ooword') return
      var el = __refTooltipEl(doc, false)
      if (el) el.style.display = 'none'
      // 命中检测：从事件坐标找文本节点，向上找带 key 的 sdt 渲染节点
      var x = ev.clientX, y = ev.clientY
      if (t.frame.getBoundingClientRect) {
        var fr = t.frame.getBoundingClientRect()
        x = ev.clientX - fr.left
        y = ev.clientY - fr.top
      }
      var node = null
      if (doc.caretRangeFromPoint) node = doc.caretRangeFromPoint(x, y)
      else if (doc.caretPositionFromPoint) node = doc.caretPositionFromPoint(x, y)
      var start = node && (node.startContainer || node.offsetNode)
      var el2 = start && start.parentElement ? start.parentElement : (ev.target || null)
      // 沿 DOM 向上找 sdt 渲染元素（OO 的表单 span 带 content 属性/form 标记）
      var hop = 0, hitKey = ''
      while (el2 && el2 !== doc.body && hop < 12) {
        var k = el2.getAttribute && (el2.getAttribute('data-key') || el2.getAttribute('key') || '')
        var cls = String(el2.className || '')
        if (k) { hitKey = k; break }
        // OO 表单渲染特征：class 含 forms/sdt 或 role 等标记时，取其子文本对 meta
        if (/sdt|content-control|oo-form/i.test(cls)) { break }
        el2 = el2.parentElement
        hop++
      }
      if (!hitKey && el2) {
        // 兜底：取命中元素文本，与 meta 键做包含匹配（值本身就是键的可读线索）
        var txt = String(el2.textContent || '').trim()
        var metas = window.__ooRefMeta || {}
        for (var mk in metas) {
          if (txt && (txt.indexOf(mk) >= 0 || mk.indexOf(txt) >= 0) && txt.length > 0) { hitKey = mk; break }
        }
      }
      if (hitKey && window.__ooRefMeta[hitKey]) {
        __refTipPaint(doc, x, y, window.__ooRefMeta[hitKey])
      }
    } catch (eTS) { /* 悬浮提示失败静默 */ }
  }
  var __refTipBound = false
  function __refTipBind () {
    if (__refTipBound) return
    __refTipBound = true
    try { document.addEventListener('mousemove', __refTipScan, true) } catch (eB) {}
    try { window.addEventListener('scroll', function () { var t = __hiFrame(); var el = t && t.doc ? __refTooltipEl(t.doc, false) : null; if (el) el.style.display = 'none' }, true) } catch (eB2) {}
  }
  try { __refTipBind() } catch (eBind) {}

  /**
   * 刷新已登记胶囊的显示值（A 方案的取值回流）。
   * 宿主算完后把最新 __ooRefData 推来，这里按 key 更新 value 并重绘连线/胶囊。
   */
  function __refreshRefValues(items) {
    if (!items || !items.length) return
    var byKey = {}
    items.forEach(function (g) {
      ;(g.fields || []).forEach(function (f) { byKey[g.id + '#' + f.id] = f.value })
    })
    var touched = false
    ;(window.__extRefCapsules || []).forEach(function (c) {
      if (byKey[c.key] === undefined) return
      var v = byKey[c.key]
      if (c.value !== v) { c.value = v; c.w = String(v).length * 8 + 14; touched = true }
    })
    if (touched) { try { window.__drawRefLines && window.__drawRefLines() } catch (e) {} }
  }

  /* ── 高亮行列（2026-10-18 用户指令：标题栏加按钮，选中格所在行/列加底色）──
   * 原理：asc_getActiveCellCoord() 的原型带稳定公共 getter asc_getX/getY/getWidth/getHeight
   *   （数值字段名 RX/via/opa/qgb 是压缩产物，不可依赖），给出选中格在 #ws-canvas 内的像素矩形；
   *   在 canvas 的父容器 #ws-canvas-outer（position:absolute）里叠两条半透明色带
   *   （pointer-events:none，纯视觉层，不改文档一个字节）。
   * 刷新：asc_onSelectionChanged 即时 + 250ms 轮询兜底（纯滚动不改选区时坐标也会变）。
   * 宿主开关：postMessage {__ooRowColHi: true|false}。 */
  var __hiRC = { on: false, timer: null, el: null, reg: false }
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
    /* 兼容旧结构（两段式）：子元素不是 3 个就重建，避免旧 DOM 卡住不刷新 */
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
    /* 每次 重算容器几何（最大化/窗口缩放/文档重开都会变） */
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
    __hiRC.el = host
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
      /* 整行：单条带从活动格左缘（=选区最左列）铺到最右，不出十字 */
      rowA.style.left = x + 'px'; rowA.style.width = Math.max(0, hostW - x) + 'px'
      rowA.style.top = y + 'px'; rowA.style.height = h + 'px'; rowA.style.display = 'block'
      rowB.style.display = 'none'; col.style.display = 'none'
      return
    }
    if (onlyCol) {
      /* 整列：单条带从活动格上缘（=选区最上行）铺到底，不出十字 */
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
    var el = __hiRC.el || (t && t.doc.getElementById('__cb_rowcol_hi'))
    if (el && el.parentNode) el.parentNode.removeChild(el)
    __hiRC.el = null
  }
  function __hiSet (on) {
    __hiRC.on = on === true
    if (!__hiRC.on) { __hiTeardown(); return }
    __hiPaint()
    if (!__hiRC.timer) __hiRC.timer = setInterval(__hiPaint, 250)
    /* 选区事件即时刷新（250ms 轮询已兜底，这里只为跟手）；引擎域回调 → 只碰引擎帧 DOM，安全 */
    if (!__hiRC.reg) {
      __hiRC.reg = true
      try {
        var t = __hiFrame()
        var w2 = t && t.frame.contentWindow
        var api = w2 && (w2.editor || (w2.Asc && w2.Asc.editor))
        if (api && typeof api.asc_registerCallback === 'function') {
          api.asc_registerCallback('asc_onSelectionChanged', function () { try { __hiPaint() } catch (eH2) {} })
        }
      } catch (eH3) { /* 轮询兜底，注册失败不致命 */ }
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
    native: false, nfList: [], nfMap: {}, nfOrig: {},
    nfBusy: false, nfPainting: false, nfEv: false, nfColor: 'FFF2CC',
    nfEditHit: '',
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
    if (__rm.native) return /* 原生标记接管显示：浮层停画，避免叠加偏色 */
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
  /* ── 原生标记（2026-10-18 用户「能不能以原生的设置进行标记」）──
   * asc_setCellBackgroundColor = web-apps 填充按钮选色时的同款调用（app.js 实锤：
   * onColorsBackSelect → api.asc_setCellBackgroundColor(ThemeColor.getRgbColor(e))），
   * 颜色写进单元格模型：零偏移零闪烁，随滚动/缩放/隐藏行列自动正确，编辑时就是
   * 背景色、不遮挡输入 —— 覆盖层四个老大难（偏移/漂/闪/遮挡）在原生路径下不存在。
   * 代价（按钮 title 已告知用户）：① 文档变「未保存」（脏标记实测 true），
   * Ctrl+Z 可逐格撤销；② 清除=设回透明，原格若有填充色会被抹掉（sdk 未暴露读原色）；
   * ③ 不保存就关闭，颜色随文档丢弃（重新点按钮即可恢复）。
   * 与覆盖层互斥：__rm.native 开着时 __rmPaint 直接不画（数据流保留，切回零成本）。 */
  function __nfFrame () {
    var t = __hiFrame()
    if (!t) return null
    var w = t.frame.contentWindow
    var api = w && (w.editor || (w.Asc && w.Asc.editor))
    if (!api || typeof api.asc_setCellBackgroundColor !== 'function') return null
    return { t: t, w: w, api: api, nb: t.doc.getElementById('ce-cell-name') }
  }
  function __nfColor (w, argb) {
    var s = String(argb || 'FFFFF2CC')
    if (s.charAt(0) === '#') s = s.slice(1)
    while (s.length < 8) s = 'F' + s
    s = s.slice(2) /* ARGB → RGB（getRgbColor 吃 6 位 hex） */
    try { return (w.Common && w.Common.Utils && w.Common.Utils.ThemeColor) ? w.Common.Utils.ThemeColor.getRgbColor(s) : null } catch (eC1) { return null }
  }
  /* 点击「原生标记」后的可见提示：浮一个小条 2.2s 淡出。
   * 之前按钮只有 font-weight/color 的弱反馈，用户在「没涂上色」时分不清是没执行还是没效果。 */
  function __nfTip (msg) {
    try {
      var f = __nfFrame()
      var doc = f ? f.t.doc : document
      var el = doc.getElementById('__cb_nf_tip')
      if (!el) {
        el = doc.createElement('div')
        el.id = '__cb_nf_tip'
        el.style.cssText = 'position:fixed;left:50%;bottom:70px;transform:translateX(-50%);z-index:99999;' +
          'pointer-events:none;background:rgba(20,24,30,.92);color:#fff;font:12px/1.7 "Microsoft YaHei",sans-serif;' +
          'padding:6px 14px;border-radius:6px;box-shadow:0 4px 16px rgba(0,0,0,.28);transition:opacity .25s;opacity:0'
        doc.body.appendChild(el)
      }
      el.textContent = msg
      el.style.opacity = '1'
      if (__rm.tipT) { try { clearTimeout(__rm.tipT) } catch (eT9) {} }
      __rm.tipT = setTimeout(function () { try { el.style.opacity = '0' } catch (eT8) {} }, 2400)
    } catch (eT7) { /* 提示失败不影响功能 */ }
  }
  /* 按钮自身状态：开着=橙色加粗 + 已标记格数；异常=「原生 ✗」并把原因写进 title */
  function __nfUi (err) {
    var btn = window.__extNativeBtn || null
    if (!btn) return
    try {
      var on = __rm.native === true
      btn.textContent = err ? '原生 ✗' : (on ? ('原生 ' + (__rm.nfList ? __rm.nfList.length : 0)) : '原生标记')
      btn.style.fontWeight = on ? 'bold' : ''
      btn.style.color = err ? '#d33' : (on ? '#c60' : '')
      btn.style.background = on ? 'rgba(255,190,60,.22)' : ''
      /* ⚠ 不能写反斜杠 n 形式的换行转义：本段会被包进 HTML 的模板字面量，会被吞成真实换行把字符串劈断
       *   （第三次实锤：页面直接 SyntaxError 白屏）。用 String.fromCharCode(10)。 */
      var nl = String.fromCharCode(10)
      btn.title = err
        ? ('原生标记失败：' + String(err).slice(0, 160) + nl + '（点击重试，状态见 window.__nfErr）')
        : ('用 OO 原生填充色标记引用格：绝对跟格、编辑不遮挡，新增引用自动补涂。当前标记 ' +
           (__rm.nfList ? __rm.nfList.length : 0) + ' 格。' + nl +
           '注意：文档会变为「未保存」（Ctrl+Z 可逐格撤销）；关闭时按每格原有填充色还原。')
    } catch (eU1) { /* 忽略 */ }
  }
  function __nfSheetNow (doc) {
    try {
      var li = doc.querySelector('#status-sheets-bar li.list-item.active')
      return li ? String(li.textContent || '').trim() : ''
    } catch (eS1) { return '' }
  }
  function __nfKey (mk) {
    var r = String((mk && mk.ref) || '').replace(/[$]/g, '').toUpperCase()
    if (!r) return ''
    var sh = String((mk && mk.sheet) || '').trim()
    return sh ? sh + '!' + r : r
  }
  /* ★ 读当前格的**原填充色**（原生接口，作用于当前选区）：
   *   asc_getCellInfo().asc_getXfs().asc_getFillColor() → Asc.c_oAscColor
   *   get_auto()=true 表示「自动填充/无填充」，get_r()/get_g()/get_b() 是实际 RGB。
   *   这正是 web-apps 填充按钮回显自身状态用的接口（app.js 的 clrtext_asccolor 分支），
   *   所以读到的值就是用户眼睛看到的填充色 —— 关灯时据此还原，不再一律抹成透明。
   *   局限：主题色读成具体 RGB（写回后与主题脱钩，视觉一致）；图案填充只还原底面纯色。 */
  function __nfReadFill (f) {
    try {
      var xfs = f.api.asc_getCellInfo().asc_getXfs()
      if (!xfs || typeof xfs.asc_getFillColor !== 'function') return null
      var c = xfs.asc_getFillColor()
      if (!c) return { auto: true, r: 255, g: 255, b: 255 }
      return { auto: c.get_auto() === true, r: c.get_r(), g: c.get_g(), b: c.get_b() }
    } catch (eRF) { return null }
  }
  function __nfHex2 (n) {
    var s = Number(n || 0).toString(16).toUpperCase()
    return s.length < 2 ? '0' + s : s
  }
  /* 对一组键（'Sheet1!B3' 格式，名称框可跨表跳）统一设置 / 还原。同步执行：
   * 全部 jump+get/set 在一个 JS 任务里跑完，浏览器不重绘，用户只看到结果。
   * on=true：先把每格原色记进 __rm.nfOrig（只记第一次），再涂标记色；
   * on=false：按 __rm.nfOrig 写回原色（没记到原色的退回透明）。 */
  /* ★ 2026-10-18 用户指令「双击参数→输入引用→原生标记的变化 放在每个日志记录里 记录出差异」：
   *   三通道日志：console（→.dev-console.log）+ __ooStatus（→诊断抽屉 ⓘ）+ /office-log（→oo-events.log）。
   *   每个环节记录「前态 → 后态」差异：原生标记清单新增/移除的引用格、每格填充色变化。 */
  function __mkLog (type, detail, skipDiag) {
    try { console.log('[cboo] ' + type + ' ' + detail) } catch (eL0) {}
    /* ★ 2026-10-18 去重：__ooStatus 会被宿主 addDiag 落一遍 [DIAG]，加上 /office-log
     *   直写的一遍 = 每条两行。高频差异日志（MARK-DIFF/MARK-IDLE）只走文件，
     *   关键节点（REF-INSERT/REF-LAND/MARK-CLICK/MARK-SKIP）保留三通道。 */
    if (!skipDiag) { try { postToHost({ __ooStatus: '[' + type + '] ' + detail }) } catch (eL1) {} }
    try {
      fetch('/office-log', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ type: type, detail: String(detail).slice(0, 600) }) })
    } catch (eL2) {}
  }
  function __nfSnap () {
    return {
      native: __rm.native === true,
      list: (__rm.nfList || []).slice().sort(),
      mapN: Object.keys(__rm.nfMap || {}).length,
      origN: Object.keys(__rm.nfOrig || {}).length,
    }
  }
  function __nfSnapStr (s) {
    return 'native=' + (s.native ? 'on' : 'off') + ' 已标记=' + s.list.length + ' 格 记账=' + s.mapN
  }
  function __nfDiffReport (tag, before, extra) {
    try {
      var after = __nfSnap()
      var added = []
      var removed = []
      var i2 = 0
      for (i2 = 0; i2 < after.list.length; i2++) { if (before.list.indexOf(after.list[i2]) < 0) added.push(after.list[i2]) }
      for (i2 = 0; i2 < before.list.length; i2++) { if (after.list.indexOf(before.list[i2]) < 0) removed.push(before.list[i2]) }
      var txt = tag + '｜前态[' + __nfSnapStr(before) + '] → 后态[' + __nfSnapStr(after) + ']' +
        (added.length ? '｜新增标记=' + added.join(',') : '') +
        (removed.length ? '｜移除标记=' + removed.join(',') : '') +
        (extra ? '｜' + extra : '')
      __mkLog('MARK-DIFF', txt, true) /* 高频：不进诊断抽屉（去重），文件里留档 */
    } catch (eD0) {}
  }

  function __nfPaintKeys (keys, on, colorArgb) {
    var __snap0 = __nfSnap()
    var f = __nfFrame()
    if (!f || !f.nb) { try { window.__nfDbg = 'frame null keys=' + keys.length } catch (eD0) {} return [] }
    /* 涂色期间屏蔽「选区变化 → 逐格校验」回调（paint 自己会跳格，跳格会触发选区回调） */
    __rm.nfPainting = true
    /* ★ 焦点先回表格：按钮点击进入时 focus 在按钮上，OO 处于「控件激活」态，
     *   此时 asc_setCellBackgroundColor(null) 只改工具色不应用（矩阵实验 E2 同调用在
     *   表格焦点下成功）——先把焦点还给 area_id 再动名称框。 */
    try {
      var ar0 = f.t.doc.getElementById('area_id')
      if (ar0 && f.t.doc.activeElement !== ar0) ar0.focus()
    } catch (eF4) { /* 忽略 */ }
    var doneK = []
    var savedAddr = f.nb.value ? String(f.nb.value) : ''
    var savedSheet = __nfSheetNow(f.t.doc)
    var prevAct = null
    try { prevAct = f.t.doc.activeElement } catch (eP1) { prevAct = null }
    var color = on ? __nfColor(f.w, colorArgb) : null
    var dbg = 'paint n=' + keys.length + ' on=' + on
    var cellLog = []
    var markHex = on ? String(colorArgb || __rm.nfColor || 'FFF2CC').slice(-6).toUpperCase() : ''
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i]
      if (!k) continue
      try {
        __rmJump(f.nb, k)
        if (on) {
          if (!__rm.nfOrig[k]) {
            var of = __nfReadFill(f)
            if (of) {
              /* ★ 原色恰好等于本次标记色 = 上一轮标记写进文档后保存下来的残留
               *   （用户「点了没反应」的典型场景：文档里已经是 FFF2CC，再涂同色看不出来）。
               *   这种按「无原色」记账，关闭时还原成透明 —— 关灯才有可见效果。 */
              var cr = -1, cg = -1, cb = -1
              if (color) { try { cr = color.get_r(); cg = color.get_g(); cb = color.get_b() } catch (eCr) {} }
              if (!of.auto && of.r === cr && of.g === cg && of.b === cb) {
                __rm.nfOrig[k] = { auto: true, r: 255, g: 255, b: 255 }
              } else {
                __rm.nfOrig[k] = of
              }
            }
          }
          f.api.asc_setCellBackgroundColor(color)
          try {
            var o0 = __rm.nfOrig[k]
            cellLog.push(k + ': ' + (o0 && !o0.auto ? ('#' + __nfHex2(o0.r) + __nfHex2(o0.g) + __nfHex2(o0.b)) : '无填充') + ' → #' + markHex)
          } catch (eC0) {}
        } else {
          var o = __rm.nfOrig[k]
          var oTxt = o && !o.auto ? ('#' + __nfHex2(o.r) + __nfHex2(o.g) + __nfHex2(o.b)) : '无填充'
          if (o && !o.auto) {
            f.api.asc_setCellBackgroundColor(__nfColor(f.w, 'FF' + __nfHex2(o.r) + __nfHex2(o.g) + __nfHex2(o.b)))
          } else {
            f.api.asc_setCellBackgroundColor(null)
          }
          delete __rm.nfOrig[k]
          cellLog.push(k + ': #' + markHex + ' → ' + oTxt + '（还原原色）')
        }
        doneK.push(k)
      } catch (eP2) { dbg += ' | ' + k + '→ERR ' + String(eP2).slice(0, 60); cellLog.push(k + ': 涂色异常 ' + String(eP2).slice(0, 40)) }
    }
    try { window.__nfDbg = dbg } catch (eD1) {}
    /* ★ 每次涂色/还原都出一份差异报告（三通道），extra 里带每格 前→后 填充色 */
    try {
      var extraC = cellLog.length ? ('逐格：' + cellLog.slice(0, 8).join(' ; ') + (cellLog.length > 8 ? ' …共 ' + cellLog.length + ' 格' : '')) : ''
      __nfDiffReport('原生标记' + (on ? '涂色' : '还原'), __snap0, extraC)
    } catch (eDR0) {}
    /* 跳回原选区（若中途切过表，带上原表名前缀），焦点还给表格 —— 同 __rmMeasure 纪律 */
    try { __rmJump(f.nb, savedAddr ? (savedAddr.indexOf('!') < 0 && savedSheet ? savedSheet + '!' + savedAddr : savedAddr) : 'A1') } catch (eB1) {}
    try { if (prevAct && prevAct.focus) prevAct.focus() } catch (eF2) {}
    try {
      var ar = f.t.doc.getElementById('area_id')
      if (ar && f.t.doc.activeElement !== ar) ar.focus()
    } catch (eF3) {}
    __rm.nfPainting = false
    return doneK
  }
  /* marks 差量同步：新增的涂、消失的还原（宿主幂等推送不会走到这，只有真变化才来） */
  function __nfSync () {
    if (!__rm.native) return
    var f = __nfFrame()
    if (!f) { try { window.__nfSyncDbg = 'frame null' } catch (eSD1) {} return }
    var want = {}
    var wantList = []
    for (var i = 0; i < __rm.marks.length; i++) {
      var k = __nfKey(__rm.marks[i])
      if (k && !want[k]) { want[k] = __rm.marks[i]; wantList.push(k) }
    }
    var toClear = []
    for (var j = 0; j < __rm.nfList.length; j++) { if (!want[__rm.nfList[j]]) toClear.push(__rm.nfList[j]) }
    var toPaint = []
    for (var m = 0; m < wantList.length; m++) { if (!__rm.nfMap[wantList[m]]) toPaint.push(want[wantList[m]]) }
    try { window.__nfSyncDbg = 'nf=' + __rm.nfList.length + ' clear=' + toClear.length + ' paint=' + toPaint.length } catch (eSD2) {}
    if (!toClear.length && !toPaint.length) return
    if (toClear.length) __nfPaintKeys(toClear, false)
    /* 按颜色分组批量涂，减少 jump/焦点往返 */
    var byColor = {}
    var colorOrder = []
    for (var n = 0; n < toPaint.length; n++) {
      var mk = toPaint[n]
      if (!mk) continue
      var ck = String(mk.color || 'FFF2CC')
      if (!byColor[ck]) { byColor[ck] = []; colorOrder.push(ck) }
      var kk = __nfKey(mk)
      if (kk) byColor[ck].push(kk)
    }
    for (var c = 0; c < colorOrder.length; c++) {
      if (byColor[colorOrder[c]].length) __nfPaintKeys(byColor[colorOrder[c]], true, colorOrder[c])
    }
    __rm.nfList = wantList.slice()
    __rm.nfMap = {}
    for (var a = 0; a < wantList.length; a++) __rm.nfMap[wantList[a]] = true
    try { __nfUi() } catch (eU9) {}
  }
  /* ★ 2026-10-18 用户「我新加的都没反应」根因：marks 只在**打开文档那一刻**由宿主
   *   解析 xlsx 算一次（NodeEditor.injectRefDomains → extractRefMarks），此后文档里
   *   新增/改动的引用格永远进不了这个清单 —— 原生标记（和浮层）都只能涂老的那批。
   *   这里让注入层自己认引用格：**公式里出现「定义名称」的格就是引用格**
   *   （与宿主 findRefCells 同一条规则）。用户点/编辑到这种格时就地补涂；
   *   格子里不再有引用时把颜色还原。这样新增引用无需重开文档、无需宿主配合。 */
  var __nfNamesCache = { at: 0, list: [] }
  function __nfNames (f) {
    var now = Date.now()
    if (now - __nfNamesCache.at < 1500) return __nfNamesCache.list
    var list = []
    try {
      var arr = f.api.asc_getDefinedNames(f.w.Asc.c_oAscGetDefinedNamesList.WorksheetWorkbook, true) || []
      for (var i = 0; i < arr.length; i++) {
        try {
          if (typeof arr[i].asc_getIsXlnm === 'function' && arr[i].asc_getIsXlnm()) continue /* 内置名（Print_Area 等）不算 */
          var nm = String(arr[i].asc_getName() || '')
          if (nm) list.push(nm)
        } catch (eN2) { /* 单个坏名跳过 */ }
      }
    } catch (eN1) { list = [] }
    __nfNamesCache.at = now
    __nfNamesCache.list = list
    return list
  }
  /* 名称边界判定（等价宿主的 nameBoundary 前后瞻）：字母/数字/_/./汉字 都不算边界 */
  function __nfBoundOk (ch) {
    if (!ch) return true
    var c = ch.charCodeAt(0)
    if (c >= 48 && c <= 57) return false
    if (c >= 65 && c <= 90) return false
    if (c >= 97 && c <= 122) return false
    if (c === 95 || c === 46) return false
    if (c >= 0x4e00 && c <= 0x9fa5) return false
    return true
  }
  function __nfHasName (text, names) {
    var t = String(text || '')
    if (!t) return ''
    for (var i = 0; i < names.length; i++) {
      var nm = names[i]
      if (!nm) continue
      var at = t.indexOf(nm)
      while (at >= 0) {
        var before = at > 0 ? t.charAt(at - 1) : ''
        var after = t.charAt(at + nm.length)
        if (__nfBoundOk(before) && __nfBoundOk(after) && after !== '(') return nm
        at = t.indexOf(nm, at + 1)
      }
    }
    return ''
  }
  /* 当前格是不是引用格（读的是公式原文，asc_getText 对公式格返回 '=名称'） */
  function __nfCellIsRef (f, names) {
    if (!names || !names.length) return ''
    var txt = ''
    try { txt = String(f.api.asc_getCellInfo().asc_getText() || '') } catch (eC2) { return '' }
    return __nfHasName(txt, names)
  }
  /* '$A$1:$C$5' / '$A$1' → ['Sheet1!A1', ...] */
  function __nfRangeKeys (range, sheet) {
    var out = []
    var m = /^[$]?([A-Z]+)[$]?([0-9]+)(?::[$]?([A-Z]+)[$]?([0-9]+))?$/.exec(String(range || '').toUpperCase())
    if (!m) return out
    var c1 = 0, i
    for (i = 0; i < m[1].length; i++) c1 = c1 * 26 + (m[1].charCodeAt(i) - 64)
    var r1 = Number(m[2])
    var c2 = c1, r2 = r1
    if (m[3]) {
      c2 = 0
      for (i = 0; i < m[3].length; i++) c2 = c2 * 26 + (m[3].charCodeAt(i) - 64)
      r2 = Number(m[4])
    }
    if (c2 < c1) { var tc = c1; c1 = c2; c2 = tc }
    if (r2 < r1) { var tr = r1; r1 = r2; r2 = tr }
    for (var r = r1; r <= r2; r++) {
      for (var cc = c1; cc <= c2; cc++) {
        var s = '', x = cc
        while (x > 0) { var d2 = (x - 1) % 26; s = String.fromCharCode(65 + d2) + s; x = Math.floor((x - 1) / 26) }
        out.push(sheet ? sheet + '!' + s + r : s + r)
      }
    }
    return out
  }
  /* 核心：对一组键逐格校验 —— 是引用格就补涂，格子里已没有引用就还原。
   * 注意：__nfPaintKeys 内部会跳格并跳回，所以第二轮起一律显式 jump，不能假设还在原位。 */
  function __nfCheckKeys (keys) {
    if (!__rm.native || __rm.nfBusy || __rm.nfPainting) return
    try { if (window.__nfOff) return } catch (eOff) { /* 调试开关：探针采样时关掉自动识别 */ }
    var f = __nfFrame()
    if (!f || !keys || !keys.length) return
    var names = __nfNames(f)
    if (!names.length) return
    var cur = ''
    try { cur = String(f.nb.value || '').replace(/[$]/g, '').toUpperCase() } catch (eC0) { cur = '' }
    __rm.nfBusy = true
    try {
      for (var i = 0; i < keys.length; i++) {
        var k = keys[i]
        if (!k) continue
        var short = k.indexOf('!') >= 0 ? k.slice(k.indexOf('!') + 1) : k
        var hit = ''
        try {
          if (cur !== short || i > 0) __rmJump(f.nb, k)
          hit = __nfCellIsRef(f, names)
        } catch (eK2) { hit = '' }
        if (hit && !__rm.nfMap[k]) {
          __rm.nfMap[k] = true
          if (__rm.nfList.indexOf(k) < 0) __rm.nfList.push(k)
          __nfPaintKeys([k], true, __rm.nfColor || 'FFF2CC')
        } else if (!hit && __rm.nfMap[k]) {
          /* 格子里已经没有引用了（用户删了公式）→ 还原原色 */
          delete __rm.nfMap[k]
          var at = __rm.nfList.indexOf(k)
          if (at >= 0) __rm.nfList.splice(at, 1)
          __nfPaintKeys([k], false)
        }
      }
    } finally { __rm.nfBusy = false }
    try { __nfUi() } catch (eU8) {}
  }
  /* 按当前选区校验（用户点/拖选到哪就校验到哪） */
  function __nfCheckSelection () {
    /* ★ 2026-10-18 用户指令「原生标记没开的时候也有记录」：未开时留痕——
     *   当前格是引用格 → 记一条 MARK-IDLE（同格去重，不刷屏），并提示如何补涂。 */
    if (!__rm.native) {
      try {
        var fI = __nfFrame()
        if (fI && fI.nb) {
          var addrI = String(fI.nb.value || '').replace(/[$]/g, '').toUpperCase()
          var shI = __nfSheetNow(fI.t.doc)
          var keyI = (shI ? shI + '!' : '') + addrI
          if (keyI && addrI && __nfCellIsRef(fI, __nfNames(fI))) {
            if (window.__nfIdleLast !== keyI) {
              window.__nfIdleLast = keyI
              __mkLog('MARK-IDLE', '原生标记未开启｜当前格 ' + keyI + ' 是引用格（未标记）——开启原生标记后点该格即可补涂', true)
            }
          } else if (window.__nfIdleLast) {
            window.__nfIdleLast = '' /* 移出引用格：允许下次再记 */
          }
        }
      } catch (eI0) {}
      return
    }
    if (__rm.nfBusy || __rm.nfPainting) return
    try { if (window.__nfOff) return } catch (eOff) { /* 调试开关 */ }
    var f = __nfFrame()
    if (!f) return
    var range = ''
    try { range = String(f.api.asc_getActiveRangeStr(f.w.Asc.referenceType.A, true) || '') } catch (eR2) { range = '' }
    var keys = __nfRangeKeys(range, __nfSheetNow(f.t.doc))
    try {
      window.__nfChk = window.__nfChk || []
      window.__nfChk.push(String(range) + ' → ' + keys.length + '格 ' + keys.slice(0, 4).join(','))
      if (window.__nfChk.length > 30) window.__nfChk.shift()
    } catch (eChk) {}
    if (!keys.length || keys.length > 40) return
    __nfCheckKeys(keys)
  }
  var __nfChkTimer = 0
  /* 去抖：拖选/连点会连续触发选区变化，只跑最后一次，避免逐格校验把主线程占满 */
  function __nfQueueCheck (delay) {
    /* ★ 2026-10-18：原生标记未开也放行——__nfCheckSelection 里会走 MARK-IDLE 留痕分支 */
    if (__nfChkTimer) { try { clearTimeout(__nfChkTimer) } catch (eT0) {} }
    __nfChkTimer = setTimeout(function () {
      __nfChkTimer = 0
      try { __nfCheckSelection() } catch (eT1) { /* 忽略 */ }
    }, delay || 160)
  }
  /* ★ 2026-10-18 用户指令「在公式编辑的时候就判断是否有外部关联引用 同时添加原生标记」：
   *   之前的自动补涂只挂在 editEnd（提交后 200ms）和 selectionChanged —— 编辑输入过程中
   *   没有任何反馈。实测（probe-nfedit4/5/6 三轮对照）：
   *     · #ce-cell-content（TEXTAREA）可读到正在输入的公式原文，但编辑态焦点在 area_id，
   *       打字由引擎程序化写入 → input 事件不触发，用 200ms 轮询比对 value 变化；
   *     · 编辑态下 asc_setCellBackgroundColor **从 /poc 侧调用一律不落模型**（跨 iframe 直调、
   *       f.w.eval 内联都返回成功但无效果；且会产生阻止兜底涂色的幽灵记账）——此路不通；
   *     · 所以最终形态 = 编辑中实时判定（零副作用）+ 状态栏浮条即时反馈 + 提交瞬间立即涂色：
   *       编辑框正好盖住该格，编辑中涂了也看不见；提交瞬间（editEnd 0ms）格色立现。 */
  function __nfEditJudge (f) {
    if (!__rm.native || __rm.nfPainting || __rm.nfBusy) return
    try { if (window.__nfOff) return } catch (eO5) {}
    var ta = f.t.doc.getElementById('ce-cell-content')
    if (!ta) return
    var addr = ''
    try { addr = String(f.nb.value || '').replace(/[$]/g, '').toUpperCase() } catch (eA5) { addr = '' }
    if (!addr) return
    /* ★ 名称框在编辑中可能带表名前缀（实测 'SHEET1!P10'）——去前缀，防止拼出 'Sheet1!SHEET1!P10' 坏键 */
    if (addr.indexOf('!') >= 0) addr = addr.slice(addr.indexOf('!') + 1)
    var sh5 = __nfSheetNow(f.t.doc)
    var k5 = sh5 ? sh5 + '!' + addr : addr
    var hit5 = __nfHasName(ta.value, __nfNames(f))
    /* ★ 2026-10-18 二次定案（probe-nfedit5/6 对照实验）：编辑态下 asc_setCellBackgroundColor
     *   **从 /poc 侧无论跨 iframe 直调还是 f.w.eval 内联调都不落模型**（返回 ok 但无效果），
     *   且先记账后失败会产生「nfMap 幽灵记录」阻止提交后兜底涂色 —— 所以编辑中**只判定不涂色**：
     *     · 命中 → 记 pending（__rm.nfEditHit）+ 状态栏浮条即时反馈（编辑框盖住格子，这
     *       是编辑中唯一可见通道）；
     *     · 提交瞬间 editEnd 处理器**立即**（0ms，不等 200ms）对 pending 格跑校验涂色；
     *     · 编辑中判定本身零副作用，失败态不存在。 */
    if (hit5) {
      if (__rm.nfEditHit !== k5) {
        __rm.nfEditHit = k5
        try { __nfTip('编辑中：外部关联引用「' + hit5 + '」— 提交后自动标记') } catch (eT8) {}
        try { window.__nfEditLog = window.__nfEditLog || []; window.__nfEditLog.push('editJudge hit ' + k5 + ' name=' + hit5) ; if (window.__nfEditLog.length > 20) window.__nfEditLog.shift() } catch (eL5) {}
      }
    } else {
      if (__rm.nfEditHit) {
        __rm.nfEditHit = ''
        try { __nfTip('编辑中：无外部关联引用') } catch (eT9) {}
        try { window.__nfEditLog = window.__nfEditLog || []; window.__nfEditLog.push('editJudge miss ' + k5) ; if (window.__nfEditLog.length > 20) window.__nfEditLog.shift() } catch (eL6) {}
      }
    }
  }
  var __nfEditTypeTimer = 0
  /* 编辑开始（asc_onEditCell st=1/2）后启动轮询监听。
   * ★ 实测（probe-nfedit5 首轮）：编辑态焦点在 area_id（隐藏 textarea），用户打字由引擎
   *   程序化写入 #ce-cell-content —— 挂 input 事件**不会触发**（值是程序设的）。
   *   所以用 200ms 轮询比对 value 变化：变了就判一次；编辑结束（textarea 消失/隐藏）即停。 */
  function __nfEditWatch (f) {
    var doc5 = f.t.doc
    var last5 = ''
    var miss5 = 0
    if (window.__nfEditPoll) { try { clearInterval(window.__nfEditPoll) } catch (eI0) {} }
    window.__nfEditPoll = setInterval(function () {
      try {
        var ta5 = doc5.getElementById('ce-cell-content')
        if (!ta5 || ta5.offsetHeight === 0) {
          miss5++
          if (miss5 > 2) { try { clearInterval(window.__nfEditPoll) } catch (eI1) {} window.__nfEditPoll = null }
          return
        }
        miss5 = 0
        var cur5 = String(ta5.value || '')
        if (cur5 === last5) return
        last5 = cur5
        __nfEditJudge(f)
      } catch (eP5) { /* 忽略 */ }
    }, 200)
    try { window.__nfEditLog = window.__nfEditLog || []; window.__nfEditLog.push('editWatch on') ; if (window.__nfEditLog.length > 20) window.__nfEditLog.shift() } catch (eL7) {}
  }
  function __nfBindEvents () {
    var f = __nfFrame()
    if (!f) return
    /* 编辑器重载后 api 会换新对象（iframe key 变化），按 api 身份判重绑 */
    if (__rm.nfEvApi === f.api) return
    __rm.nfEvApi = f.api
    __rm.nfEv = true
    try {
      f.api.asc_registerCallback('asc_onSelectionChanged', function () {
        if (!__rm.native) return
        __nfQueueCheck(160)
      })
    } catch (eE0) { /* 忽略 */ }
    try {
      /* ★ asc_onEditCell 的参数是**数字状态码**（实测 1=start, 2=inCell, 4=end, 0=idle），
       *   不是字符串 'editEnd' —— 上一版按字符串比导致编辑新增永远不校验。 */
      var stEnd = 4
      try { stEnd = f.w.Asc.c_oAscCellEditorState.editEnd } catch (eSE0) { stEnd = 4 }
      f.api.asc_registerCallback('asc_onEditCell', function (st) {
        if (!__rm.native) return
        /* ★ 编辑开始/进入格子（实测 1=start, 2=inCell）：启动输入轮询 —— 用户指令
         *   2026-10-18「在公式编辑的时候就判断是否有外部关联引用 同时添加原生标记」 */
        if (st === 1 || st === 2) {
          try { window.__nfEditLog = window.__nfEditLog || []; window.__nfEditLog.push('editStart st=' + String(st)) ; if (window.__nfEditLog.length > 20) window.__nfEditLog.shift() } catch (eW1) {}
          try { __nfEditWatch(f) } catch (eW0) { /* 忽略 */ }
          return
        }
        /* ★ 实测（probe-nfedit5）：提交事件的状态码是 **0**（editEnd nb=Q12 st=0），
         *   不是 4 —— c_oAscCellEditorState.editEnd 在本版不代表提交。0 与 stEnd 都接受，
         *   多余的一次校验无害（校验当前格，命中涂/未命中还原幂等）。 */
        if (st !== stEnd && st !== 0 && String(st) !== 'editEnd') return
        /* 此刻 nb.value 仍是刚编辑的格（实测 st=4 时 nb=M9）。先记下地址再延迟校验：
         * 回车后选区会下移到下一格，隔 200ms 再读就变成别的格了。 */
        var addr = ''
        try { addr = String(f.nb.value || '').replace(/[$]/g, '').toUpperCase() } catch (eEA0) { addr = '' }
        try {
          window.__nfEditLog = window.__nfEditLog || []
          window.__nfEditLog.push('editEnd nb=' + addr + ' st=' + String(st))
          if (window.__nfEditLog.length > 20) window.__nfEditLog.shift()
        } catch (eEL1) { /* 忽略 */ }
        if (!addr) return
        /* ★ 名称框此刻可能带表名前缀（实测 'SHEET1!P10'）——去前缀防坏键 */
        if (addr.indexOf('!') >= 0) addr = addr.slice(addr.indexOf('!') + 1)
        /* ★ pending 优先：编辑中判定命中的格，提交瞬间**立即**校验涂色（0ms，不等 200ms）。
         *   pending 的 k5 已含表名，直接用；无 pending 时退回 addr 拼当前表名。 */
        var pend = __rm.nfEditHit || ''
        __rm.nfEditHit = ''
        try {
          __nfCheckKeys([pend || (__nfSheetNow(f.t.doc) ? __nfSheetNow(f.t.doc) + '!' + addr : addr)])
        } catch (eEA2) { /* 忽略 */ }
        setTimeout(function () {
          if (!__rm.native) return
          try {
            var sh = __nfSheetNow(f.t.doc)
            __nfCheckKeys([sh ? sh + '!' + addr : addr])
          } catch (eEA1) { /* 忽略 */ }
        }, 200)
      })
    } catch (eE3) { /* 忽略 */ }
  }
  function __nfToggle () {
    try { console.log('[cboo] nfToggle enter -> native=' + (!__rm.native) + ' marks=' + ((__rm.marks || []).length) + ' on=' + __rm.on + ' frame=' + (__nfFrame() ? 'ok' : 'null')) } catch (eLG2) {}
    var __tSnap = __nfSnap()
    __rm.native = !__rm.native
    if (__rm.native) {
      __rm.nfList = []
      __rm.nfMap = {}
      __rm.nfOrig = {}
      __rm.nfEditHit = ''
      __rmStop() /* 浮层先撤，避免叠色 */
      __rm.nfColor = (__rm.marks && __rm.marks[0] && __rm.marks[0].color) || 'FFF2CC'
      /* 统一走差量路径：按每格自己的 color 涂，同时逐格记原色 */
      __nfSync()
      /* 之后靠编辑器回调自己认新引用格（点/编辑到就补涂），不再依赖宿主推送 */
      try { __nfBindEvents() } catch (eB5) {}
      try { __nfCheckSelection() } catch (eB4) {}
      __nfUi()
      /* 一格都没涂上时必须说话：否则「点了没反应」无从分辨
       * （① 宿主没推 marks；② 文档里的引用是后来加的、还没被点选过） */
      var n = __rm.nfList.length
      if (n) {
        __nfTip('原生标记：已标记 ' + n + ' 格（点选/编辑到新引用格会自动补上）')
      } else if ((__rm.marks || []).length) {
        __nfTip('原生标记：' + __rm.marks.length + ' 个引用格写入失败，详见按钮提示')
      } else {
        __nfTip('原生标记：暂未识别到引用格 —— 点一下带引用的单元格即可标记')
        try { window.__nfNoMarks = true } catch (eNM) {}
      }
    } else {
      var keys2 = __rm.nfList.slice()
      var back = keys2.length ? __nfPaintKeys(keys2, false).length : 0
      __rm.nfList = []
      __rm.nfMap = {}
      __rm.nfOrig = {}
      __nfUi()
      __nfTip(back ? ('原生标记已关闭：还原 ' + back + ' 格原色') : '原生标记已关闭')
      if (__rm.on && __rm.marks.length) {
        /* 切回浮层：重量一次再开画 */
        __rm.needMeas = true
        __rm.recheckLeft = 1
        __rm.sig = ''
        if (!__rm.raf) { try { __rm.raf = window.requestAnimationFrame(__rmTick) } catch (eR1) { __rm.raf = window.setTimeout(__rmTick, 60) } }
      }
    }
    try { console.log('[cboo] nfToggle done native=' + __rm.native + ' list=' + __rm.nfList.length + ' dbg=' + String(window.__nfDbg || '').slice(0, 140)) } catch (eLG3) {}
    try { __nfDiffReport('原生标记开关→' + (__rm.native ? '开' : '关'), __tSnap) } catch (eDR1) {}
  }
  /* ══ 外部关联引用 · 更新策略 UI（2026-10-18 用户指令）════════════════════
   * 用户原话：「打开之后检测是否有外部关联。在上方添加个按钮。点击打开设置
   *   是否自动更新 1、不检测 2、检测后弹窗提示 3、自动更新不提示。
   *   根据这些判断决定是否更新 ooexcel 内的外部的数据引用的更新内容」
   *
   * 分工：检测与写回都在宿主（它才有 xlsx 字节、域表、模板库）；这里只做三件事：
   *   ① 工具栏顶部「引用更新」按钮（排在「参数引用」右边）
   *   ② 点开的设置卡：三选一 + 本次打开检测到的关联数/待更新数
   *   ③ 宿主下发「要不要更新」时的居中弹窗（立即更新 / 暂不更新 / 以后自动更新）
   * 消息：宿主 → __ooRefUpdMode / __ooRefAsk / __ooRefUpdInfo
   *      这里 → __ooRefUpdModeSet / __ooRefApply
   * ⚠ 本段在 /poc 的 HTML 模板字符串内：禁止反斜杠转义、禁止反引号、禁止美元花括号。
   */
  var __ruMode = 'ask'
  var __ruInfo = null
  var __ruShownId = 0
  var __ruLabel = { off: '不检测', ask: '检测后提示', auto: '自动更新' }
  var __ruDesc = {
    off: '打开时不检查外部关联，按保存内容显示（数字可能不是最新）。',
    ask: '打开时检查外部关联，发现数据有变化才弹窗问你。',
    auto: '打开时直接用画布最新值更新，不弹窗不打扰。'
  }
  function __ruEdoc () {
    try {
      var f = document.querySelector('iframe[name="frameEditor"]') || document.querySelector('iframe')
      if (f && f.contentDocument) return f.contentDocument
    } catch (e0) {}
    try { return document } catch (e1) { return null }
  }
  function __ruUi (btn) {
    var b = btn || window.__extRefUpdBtn
    if (!b) return
    try {
      var cap = b.querySelector('.caption')
      if (cap) cap.textContent = '引用更新'
      else if (!b.querySelector('.caption')) b.textContent = '引用更新'
      var nl = String.fromCharCode(10)
      var n = __ruInfo && __ruInfo.total ? __ruInfo.total : 0
      var m = __ruInfo && __ruInfo.changed ? __ruInfo.changed : 0
      b.title = '外部数据引用更新：' + (__ruLabel[__ruMode] || __ruMode) + nl +
        (n ? ('本次打开检测到 ' + n + ' 处引用' + (m ? ('，其中 ' + m + ' 处数据有更新') : '，数据已是最新')) : '打开文档后自动检测（检测结果显示在设置卡里）') +
        nl + '点击设置更新策略'
      b.style.color = __ruMode === 'off' ? '#8a8a8a' : (__ruMode === 'auto' ? '#1f8a4c' : '#b96b00')
      b.style.fontWeight = __ruMode === 'off' ? '' : 'bold'
    } catch (e2) {}
  }
  function __mountRefUpdBtn () {
    try {
      var edoc = __ruEdoc()
      if (!edoc) return
      var refBtn = edoc.getElementById('ext-btn-ooref')
      if (!refBtn || !refBtn.parentNode) return
      var old = edoc.getElementById('ext-btn-refupd')
      if (old) {
        /* 旧实例可能挂在被折叠的 slot 里（实测 slot-comment 在非最大化时 0×0）→ 拆了重挂 */
        var or2 = old.getBoundingClientRect()
        if (or2.width > 4 && or2.height > 4) { __ruUi(old); return }
        try { old.parentNode.removeChild(old) } catch (eX0) {}
      }
      /* ★ 自绘普通按钮，不再克隆 OO 工具栏按钮：克隆的 x-huge/icon-top 结构离开原 slot
       *   就错位（实测文字被顶到标签栏后面，probe-shot2 截图实锤）。自绘样式只对齐
       *   功能区的尺寸/配色，放在第一个可见 slot 右边，任何 tab/模式下都完整可见。 */
      var btn = edoc.createElement('button')
      btn.id = 'ext-btn-refupd'
      btn.type = 'button'
      btn.style.cssText = 'height:24px;padding:0 10px;margin:0;box-sizing:border-box;' +
        'border:1px solid #c9cdd4;border-radius:4px;background:#fff;cursor:pointer;' +
        'font:12px/22px "Microsoft YaHei",system-ui,sans-serif;white-space:nowrap'
      btn.textContent = '引用更新'
      btn.onclick = function (ev) {
        try {
          if (ev && ev.stopPropagation) ev.stopPropagation()
          __ruOpenPop(btn)
        } catch (eB) {}
      }
      /* ★ 挂到**当前激活 tab 的可见面板**里：找第一个可见 panel 的第一个可见 slot，
       *   浅克隆保样式类后插到它右边；用户切 tab 后 slot 随面板销毁，定时器会重挂到新面板。 */
      var slotSrc = null
      var rowEl = null
      try {
        var panels = edoc.querySelectorAll('.box-panels > section.panel')
        for (var pi = 0; pi < panels.length; pi++) {
          var pr = panels[pi].getBoundingClientRect()
          if (!(pr.width > 100 && pr.height > 20)) continue
          var cands = panels[pi].querySelectorAll('.btn-slot')
          for (var si = 0; si < cands.length; si++) {
            var rr = cands[si].getBoundingClientRect()
            if (rr.width > 4 && rr.height > 4) { slotSrc = cands[si]; rowEl = cands[si].parentElement; break }
          }
          if (slotSrc) break
        }
      } catch (eP2) {}
      var slot = slotSrc ? slotSrc.cloneNode(false) : edoc.createElement('span')
      slot.removeAttribute && slot.removeAttribute('id')
      /* ★ 功能区的 slot 行是 flex 且 OO 的 CSS 强制 .btn-slot{width:20px}（probe-shot3 实测
       *   computed width=20px，按钮溢出被后面的字号控件盖住）——内联显式给宽 + 锁不压缩。 */
      slot.style.cssText += ';flex:0 0 auto;display:flex;align-items:center;margin:0 4px;width:78px'
      slot.appendChild(btn)
      if (rowEl && slotSrc) rowEl.insertBefore(slot, slotSrc.nextSibling)
      else refBtn.parentNode.parentNode.insertBefore(slot, refBtn.parentNode.nextSibling)
      window.__extRefUpdBtn = btn
      __ruUi(btn)
      try { console.log('[cboo] refupd btn mounted mode=' + __ruMode + ' slot=' + (slotSrc ? slotSrc.className : 'plain')) } catch (eC) {}
    } catch (e3) { /* 工具栏结构变了就不注入，不影响其它功能 */ }
  }
  function __ruBuildPop (doc) {
    var pop = doc.createElement('div')
    pop.id = '__cb_refupd'
    pop.style.cssText = 'position:fixed;z-index:99997;display:none;width:278px;box-sizing:border-box;' +
      'background:#fff;color:#222;border:1px solid rgba(0,0,0,.14);border-radius:8px;' +
      'box-shadow:0 10px 30px rgba(0,0,0,.24);font:12px/1.6 "Microsoft YaHei",system-ui,sans-serif;padding:10px 12px'
    pop.addEventListener('mousedown', function (ev) { try { ev.stopPropagation() } catch (eP0) {} })
    var hd = doc.createElement('div')
    hd.textContent = '外部数据引用更新'
    hd.style.cssText = 'font-size:13px;font-weight:bold;margin:0 0 2px'
    pop.appendChild(hd)
    var sub = doc.createElement('div')
    sub.id = '__cb_refupd_state'
    sub.style.cssText = 'color:#666;font-size:11px;margin-bottom:8px'
    pop.appendChild(sub)
    var modes = ['off', 'ask', 'auto']
    for (var i = 0; i < modes.length; i++) {
      var mk = modes[i]
      var row = doc.createElement('label')
      row.style.cssText = 'display:block;padding:5px 6px;border-radius:6px;cursor:pointer;margin-bottom:2px'
      var rd = doc.createElement('input')
      rd.type = 'radio'
      rd.name = '__cb_refupd_mode'
      rd.value = mk
      rd.style.cssText = 'vertical-align:-2px;margin:0 6px 0 0'
      rd.addEventListener('change', function () { try { __ruSetMode(mk) } catch (eM0) {} })
      row.appendChild(rd)
      var t = doc.createElement('span')
      t.textContent = __ruLabel[mk]
      t.style.cssText = 'font-weight:bold'
      row.appendChild(t)
      var d = doc.createElement('div')
      d.textContent = __ruDesc[mk]
      d.style.cssText = 'color:#777;font-size:11px;margin:1px 0 0 20px'
      row.appendChild(d)
      pop.appendChild(row)
    }
    var foot = doc.createElement('div')
    foot.style.cssText = 'border-top:1px solid #eee;margin-top:8px;padding-top:6px;color:#999;font-size:11px'
    foot.textContent = '设置立即保存；打开文档与上游改数时按此执行。'
    pop.appendChild(foot)
    try {
      if (!doc.__ruOutside) {
        doc.__ruOutside = true
        doc.addEventListener('mousedown', function (ev) {
          var p = doc.getElementById('__cb_refupd')
          var b = doc.getElementById('ext-btn-refupd')
          if (!p || p.style.display === 'none') return
          if (p.contains(ev.target) || (b && b.contains(ev.target))) return
          p.style.display = 'none'
        }, true)
      }
    } catch (eO0) {}
    return pop
  }
  function __ruFillPop (pop) {
    try {
      var st = pop.querySelector('#__cb_refupd_state')
      if (st) {
        var n = __ruInfo && __ruInfo.total ? __ruInfo.total : 0
        var m = __ruInfo && __ruInfo.changed ? __ruInfo.changed : 0
        st.textContent = n
          ? ('本次打开：' + n + ' 处外部关联' + (m ? ('，' + m + ' 处数据有更新') : '，数据已是最新'))
          : '本次打开：未检测到外部关联（或文档尚未打开）'
      }
      var rds = pop.querySelectorAll('input[name="__cb_refupd_mode"]')
      for (var i = 0; i < rds.length; i++) rds[i].checked = (rds[i].value === __ruMode)
    } catch (eF0) {}
  }
  function __ruOpenPop (btn) {
    try {
      var edoc = __ruEdoc()
      if (!edoc) return
      var pop = edoc.getElementById('__cb_refupd')
      if (pop && pop.style.display === 'block') { pop.style.display = 'none'; return }
      if (!pop) { pop = __ruBuildPop(edoc); edoc.body.appendChild(pop) }
      __ruFillPop(pop)
      var r = btn.getBoundingClientRect()
      var vw = edoc.documentElement.clientWidth || 1000
      var left = Math.round(r.left)
      if (left + 284 > vw - 6) left = Math.max(6, vw - 290)
      pop.style.left = left + 'px'
      pop.style.top = Math.round(r.bottom + 6) + 'px'
      pop.style.display = 'block'
    } catch (eP1) {}
  }
  function __ruSetMode (m) {
    if (m !== 'off' && m !== 'ask' && m !== 'auto') return
    __ruMode = m
    try { localStorage.setItem('ooRefUpdateMode', m) } catch (eS0) {}
    try { postToHost({ __ooRefUpdModeSet: m }) } catch (eS1) {}
    var edoc = __ruEdoc()
    var pop = edoc && edoc.getElementById('__cb_refupd')
    if (pop) __ruFillPop(pop)
    __ruUi()
    __nfTip('外部数据更新：' + __ruLabel[m])
    try { console.log('[cboo] refupd mode=' + m) } catch (eS2) {}
  }
  function __ruCloseAsk () {
    try {
      var edoc = __ruEdoc()
      var el = edoc && edoc.getElementById('__cb_refask')
      if (el && el.parentNode) el.parentNode.removeChild(el)
    } catch (eK0) {}
  }
  function __ruShowAsk (info) {
    try {
      if (!info) return
      var id = Number(info.id || 0)
      if (id && id === __ruShownId) return
      if (id) __ruShownId = id
      var doc = __ruEdoc()
      if (!doc || !doc.body) return
      __ruCloseAsk()
      var wrap = doc.createElement('div')
      wrap.id = '__cb_refask'
      wrap.style.cssText = 'position:fixed;inset:0;z-index:100001;background:rgba(12,16,22,.42);' +
        'display:flex;align-items:center;justify-content:center'
      var card = doc.createElement('div')
      card.style.cssText = 'width:392px;max-width:90%;box-sizing:border-box;background:#fff;color:#222;' +
        'border-radius:10px;box-shadow:0 20px 54px rgba(0,0,0,.34);padding:16px 18px;' +
        'font:13px/1.75 "Microsoft YaHei",system-ui,sans-serif'
      var h = doc.createElement('div')
      h.textContent = '外部数据引用有更新'
      h.style.cssText = 'font-size:14px;font-weight:bold;margin-bottom:6px'
      card.appendChild(h)
      var n = Number(info.count || 0)
      var names = (info.names && info.names.length) ? info.names : []
      var p1 = doc.createElement('div')
      p1.textContent = (info.trigger === 'live' ? '画布里的引用数据已经变化：' : '打开时检测到：') +
        n + ' 处引用与表格里的值不一致。'
      card.appendChild(p1)
      /* ★ 2026-10-18 用户指令「更新数据后要在提示里 提示更新的内容」：弹窗直接列新旧值 */
      var items = (info.items && info.items.length) ? info.items : []
      if (items.length) {
        var list = doc.createElement('div')
        list.style.cssText = 'margin-top:6px;padding:6px 10px;background:#f5f7fa;border:1px solid #e2e6ec;border-radius:6px;font-size:12px'
        for (var li = 0; li < items.length; li++) {
          var it = items[li]
          var row = doc.createElement('div')
          row.style.cssText = 'margin:2px 0'
          var nm = doc.createElement('span')
          nm.textContent = it.name
          nm.style.cssText = 'font-weight:bold'
          row.appendChild(nm)
          var vals = doc.createElement('span')
          vals.textContent = '：' + (it.was || '(空)') + ' → ' + (it.now || '(空)')
          vals.style.cssText = 'color:#555'
          row.appendChild(vals)
          list.appendChild(row)
        }
        card.appendChild(list)
      } else if (names.length) {
        var p2 = doc.createElement('div')
        p2.textContent = '涉及：' + names.join('、') + (n > names.length ? ' 等' : '')
        p2.style.cssText = 'color:#555;margin-top:4px'
        card.appendChild(p2)
      }
      var p3 = doc.createElement('div')
      p3.textContent = '更新会用画布中的最新值重写表格内的引用数据，并重新载入文档（未保存的修改请先保存）。'
      p3.style.cssText = 'color:#777;font-size:12px;margin-top:6px'
      card.appendChild(p3)
      var bar = doc.createElement('div')
      bar.style.cssText = 'display:flex;align-items:center;justify-content:flex-end;gap:8px;margin-top:14px'
      var bno = doc.createElement('button')
      bno.textContent = '暂不更新'
      bno.style.cssText = 'padding:6px 14px;border:1px solid #ccd2da;background:#f6f7f9;color:#333;border-radius:6px;cursor:pointer;font-size:12px'
      bno.onclick = function () { try { __ruCloseAsk() } catch (eN0) {} }
      var byes = doc.createElement('button')
      byes.textContent = '立即更新'
      byes.style.cssText = 'padding:6px 16px;border:1px solid #1a6fd4;background:#1a6fd4;color:#fff;border-radius:6px;cursor:pointer;font-size:12px'
      byes.onclick = function () {
        try {
          postToHost({ __ooRefApply: id || 1 })
          __ruCloseAsk()
          __nfTip('正在用最新值更新外部数据…')
        } catch (eY0) {}
      }
      bar.appendChild(bno)
      bar.appendChild(byes)
      card.appendChild(bar)
      var rem = doc.createElement('label')
      rem.style.cssText = 'display:block;margin-top:10px;color:#888;font-size:11px;cursor:pointer'
      var cb = doc.createElement('input')
      cb.type = 'checkbox'
      cb.style.cssText = 'vertical-align:-2px;margin:0 5px 0 0'
      cb.addEventListener('change', function () { try { if (cb.checked) __ruSetMode('auto') } catch (eC0) {} })
      rem.appendChild(cb)
      var rtxt = doc.createElement('span')
      rtxt.textContent = '以后自动更新，不再询问'
      rem.appendChild(rtxt)
      card.appendChild(rem)
      wrap.appendChild(card)
      doc.body.appendChild(wrap)
      try { console.log('[cboo] refask shown count=' + n + ' trigger=' + info.trigger) } catch (eA0) {}
    } catch (eA1) {}
  }
  function __rmSet (data) {
    if (!data || typeof data !== 'object') return
    var dOn = data.on === true || data.on === '1'
    var dMarks = Array.isArray(data.marks) ? data.marks : []
    try { console.log('[cboo] rmSet on=' + dOn + ' marks=' + dMarks.length) } catch (eLG1) {}
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
    if (!__rm.on || !__rm.marks.length) {
      /* 引用着色总开关关掉 / 引用清空：原生标记颜色也要还原，不能挂在文档上。
       * （临时把 marks 视为空跑一次差量同步，nfList 全部进 toClear → 逐格还原原色） */
      try {
        if (__rm.native) {
          var savedMks = __rm.marks
          __rm.marks = []
          __nfSync()
          __rm.marks = savedMks
        }
      } catch (eNF9) {}
      __rmStop(); return
    }
    __rmPaint()
    /* 原生标记开着：marks 真变了就差量补涂/清色（心跳推送不会走到这） */
    try { if (__rm.native) __nfSync() } catch (eNF0) { try { window.__nfErr = String((eNF0 && eNF0.stack) || eNF0).slice(0, 400) } catch (eNF1) {} }
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


  window.addEventListener('message', function (e) {
    /* ★ 2026-10-17 诊断计数：证明这个监听器是否真的活在当前 /poc 文档上
     *   （排「宿主发了消息但监听器不在」这类问题，只能靠帧内可见标记） */
    try { window.__ooMsgCount = (window.__ooMsgCount || 0) + 1 } catch (eMC) {}
    try {
      var d = e.data
      // ★ 高亮行列开关（宿主标题栏按钮）：开=铺色带+定时刷新；关=拆干净
      if (d && typeof d === 'object' && '__ooRowColHi' in d) {
        try { __hiSet(d.__ooRowColHi === true || d.__ooRowColHi === '1') } catch (eHi) {}
      }
      // ★ 外部数据引用更新策略（2026-10-18 用户指令）：宿主推当前模式 / 检测结果 / 更新提问
      if (d && typeof d === 'object' && '__ooRefUpdMode' in d) {
        try {
          var rum0 = String(d.__ooRefUpdMode || 'ask')
          if (rum0 === 'off' || rum0 === 'ask' || rum0 === 'auto') __ruMode = rum0
          __ruUi()
          var edocRU = __ruEdoc()
          var popRU = edocRU && edocRU.getElementById('__cb_refupd')
          if (popRU) __ruFillPop(popRU)
        } catch (eRU0) {}
      }
      if (d && typeof d === 'object' && '__ooRefUpdInfo' in d) {
        try {
          __ruInfo = d.__ooRefUpdInfo || null
          __ruUi()
          var edocRU2 = __ruEdoc()
          var popRU2 = edocRU2 && edocRU2.getElementById('__cb_refupd')
          if (popRU2) __ruFillPop(popRU2)
        } catch (eRU1) {}
      }
      if (d && typeof d === 'object' && '__ooRefAsk' in d) {
        try { __ruShowAsk(d.__ooRefAsk) } catch (eRU2) {}
      }
      // ★ 引用格覆盖层开关/数据（宿主标题栏「引用格着色」按钮）：即时生效，不碰文档
      if (d && typeof d === 'object' && '__ooRefMarks' in d) {
        try { __rmSet(d.__ooRefMarks) } catch (eRm0) {}
      }
      if (d && typeof d === 'object' && '__ooRefData' in d) {
        // ★ 2026-10-10 用户实测「点击参数就折叠」：宿主每次 ooRefs 落库都会重推
        //   __ooRefData ⇒ 旧实现整表重建，展开状态（挂在旧对象 __open 上）全丢。
        //   ① 内容没变就不重建；② 变了也保留滚动位置（展开状态由 __refOpenSet 持久化）。
        var __newItems = d.__ooRefData || []
        try {
          if (JSON.stringify(__refItems) === JSON.stringify(__newItems)) { try { __refreshRefValues(__refItems) } catch (e9) {} return }
        } catch (e8) { /* 循环结构兜底：照常重建 */ }
        var __scrollTop = __refPanel && __refPanel.querySelector('#ext-ref-body') ? __refPanel.querySelector('#ext-ref-body').scrollTop : 0
        __refItems = __newItems; __wireRefSearch(); __renderRefItems()
        try { var __b2 = __refPanel.querySelector('#ext-ref-body'); if (__b2) __b2.scrollTop = __scrollTop } catch (e7) {}
        try { __refreshRefValues(__refItems) } catch (e) {}
      }
      if (d && typeof d === 'object' && '__ooToggleRef' in d) __toggleRefPanel()
        if (d && typeof d === 'object' && '__ooSaveReq' in d) {
          /* ★ 2026-10-17 诊断：这段原先是一行裸调用，异常被外层 try 吞掉 ⇒
           *   「点保存毫无反应、日志里连 SAVE-BTN 都没有」。现在分支入口/调用异常都可见。 */
          try { window.__ooLastSaveReq = Date.now() } catch (eSR0) {}
          /* ★ 2026-10-17：接管宿主下发的会话 ID，后面每一步日志自动带上它 */
          try { if (d.__ooSid) window.__ooSid = String(d.__ooSid) } catch (eSRs) {}
          __lg('S2', '收到保存请求 mode=' + (d.__ooSaveMode || 'node') + ' sid=' + (window.__ooSid || '-') + ' extract=' + (typeof __doSaveExtract) + ' 文档脏标记=' + (window.__ooDocModified ? 'yes' : 'no'))
          try { postToHost({ __ooStatus: '[saveReq] 收到 mode=' + (d.__ooSaveMode || 'node') + ' extract=' + (typeof __doSaveExtract) }) } catch (eSR1) {}
          try { __doSaveExtract(d.__ooSaveMode || 'node') } catch (eSR2) { __lg('S2-E', '❌ 提取异常: ' + String(eSR2).slice(0, 160) + ' | ' + String(eSR2 && eSR2.stack ? eSR2.stack : '').slice(0, 200)); try { postToHost({ __ooStatus: '[saveReq] ❌ 提取异常: ' + String(eSR2).slice(0, 160) }) } catch (eSR3) {} }
        }
      if (d && typeof d === 'object' && '__ooRefSet' in d) __setRefPanelVisible(d.__ooRefSet === true)
      // ★ 宿主推来的 UI 状态（2026-10-11 最大化/常规内容同步）：搜索词 + 展开集。
      //   防乒乓：只在与当前值不同才应用；应用不回传（宿主侧 fromPoc 同样不回推）。
      if (d && typeof d === 'object' && '__ooRefUi' in d) {
        try {
          var __ui = d.__ooRefUi || {}
          if (typeof __ui.q === 'string') {
            var __qEl = __refPanel && __refPanel.querySelector('#ext-ref-q')
            if (__qEl && __qEl.value !== __ui.q) { __qEl.value = __ui.q; __renderRefItems() }
          }
          if (Array.isArray(__ui.openIds)) {
            var __next = {}
            __ui.openIds.forEach(function (id2) { __next[id2] = true })
            var __cur = window.__refOpenSet || {}
            var __curIds = Object.keys(__cur).filter(function (k4) { return __cur[k4] }).sort().join(',')
            var __nextIds = Object.keys(__next).sort().join(',')
            if (__curIds !== __nextIds) { window.__refOpenSet = __next; __renderRefItems() }
          }
        } catch (eUi) {}
      }
      // ★ 宿主回传「刷新域值」的结果（2026-10-06）——写回在宿主侧做（它才有模板/补丁/存储能力），
      //   完成后把结论推回来显示，让用户知道写了几处、或为什么没写。
      if (d && typeof d === 'object' && '__ooRefreshMsg' in d) {
        try { setRefreshMsg(d.__ooRefreshMsg) } catch (e) {}
      }
      if (d && typeof d === 'object' && '__ooRefreshState' in d) {
        try {
          var rb = document.getElementById('ext-btn-refresh')
          if (rb) rb.style.opacity = d.__ooRefreshState === true ? '0.5' : '1'
        } catch (e2) {}
      }
      // ★ 插入引用（批次 12 · ③ A 方案）：宿主把「点中的参数」交给编辑器，
      //   往当前光标处写一条引用公式文本（=节点名.字段名），像 Excel 的地址引用那样参与计算。
      //   刻意不写 OO 原生 named range（A 方案不碰内核数值域）：先用文本标记把交互跑通，
      //   值由宿主算完经__ooRefData 刷新，B 方案再升级为真单元格引用。
      if (d && typeof d === 'object' && '__ooRefInsert' in d) {
        try { __insertRefFormula(d.__ooRefInsert) } catch (e) {}
      }
      // ★ 章节目录（2026-10-09）：宿主侧栏 → 插件 callCommand。op ∈ getChapters|gotoChapter|moveChapter
      if (d && typeof d === 'object' && '__chapterCmd' in d) {
        var cc = d.__chapterCmd || {}
        __sendToRefPlugin({ __cbRefCmd: 'chapter', chapter: cc })
      }
      // ★ 运行时换主题（不重载文件）：直接调编辑器帧的官方主题 api
      if (d && typeof d === 'object' && '__ooSetTheme' in d) {
        try {
          var thId = d.__ooSetTheme
          var feT = document.querySelector('iframe[name="frameEditor"]') || document.querySelector('iframe')
          var wT = feT && feT.contentWindow
          if (wT && wT.Common && wT.Common.UI && wT.Common.UI.Themes && typeof wT.Common.UI.Themes.setTheme === 'function') wT.Common.UI.Themes.setTheme(thId)
        } catch (e) {}
        /* ★ 2026-10-18：换主题后必须重涂参数引用面板。面板配色是渲染时取的快照，
         *   官方 setTheme 只换编辑器自身主题，不会碰我们注入的 DOM —— 不重绘就一直是旧配色
         *   （用户报「更换主题颜色也不跟着变化」）。 */
        try { __repaintRefPanelSoon() } catch (eRT) {}
      }
      // ★ 关闭确认：回复宿主文档修改状态
      if (d && typeof d === 'object' && '__ooSavedAck' in d) {
        try { window.__setOoDirty(false, 'savedAck') } catch (eAk) {} // ★ 我们自己的保存管线完成 = 内容已入库，清脏标志
      }
      if (d && typeof d === 'object' && '__ooIsModifiedAsk' in d) {
        /* ★ 2026-10-17 改为「取或」：两个判据任一认为改过就算改。
         *   原实现是引擎直查结果**覆盖**本地标志 —— 而 sdkjs 的 isDocumentModified
         *   常年返回 false（引擎直查在嵌入模式下不可靠），于是标志被盖掉 ⇒ 恒 false。
         *   取舍：漏提示 = 丢数据，多提示 = 多点一次「不保存」。宁可多提示。 */
        var modAns = (typeof window.__ooDocModified === 'boolean') ? !!window.__ooDocModified : false
        try {
          var feM = document.querySelector('iframe[name="frameEditor"]') || document.querySelector('iframe')
          /* ★ 2026-10-30 修「程序化修改（callCommand）后保存丢编辑」：直查优先 Asc.editor——
           *   window.editor 是 DocsAPI 外部包装，嵌入模式下 isDocumentModified 常年 false，
           *   而 Asc.editor（sdkjs 内部）才是真状态。SetDocumentModified(true) 只改后者，
           *   旧直查永远看不到 ⇒ dirty 硬门不拦 ⇒ 主通道放行原文件 ⇒ 编辑丢失（实测翻车链）。 */
          var candM = []
          try { if (feM && feM.contentWindow && feM.contentWindow.Asc && feM.contentWindow.Asc.editor) candM.push(feM.contentWindow.Asc.editor) } catch (eA1) {}
          try { if (feM && feM.contentWindow && feM.contentWindow.editor) candM.push(feM.contentWindow.editor) } catch (eA2) {}
          for (var mi = 0; mi < candM.length; mi++) {
            var apiM = candM[mi]
            if (typeof apiM.isDocumentModified === 'function') { if (apiM.isDocumentModified()) { modAns = true; break } }
            else if (typeof apiM.asc_isDocumentModified === 'function') { if (apiM.asc_isDocumentModified()) { modAns = true; break } }
          }
        } catch (e) {}
        try { parent.postMessage({ __ooIsModified: modAns }, '*') } catch (e) {}
      }
      /* ★★ 2026-10-31 原生保存链（保存链 v4 主通道）★★
       *  引擎 asc_DownloadAs(DOCX) 导出「当前修改状态」——OO 自己序列化，格式 100%
       *  （提取+重建方案的架构性替代：提取协议永远追不上引擎的样式维度）。
       *  产物经 hook 下载捕获 → base64 回传宿主 persist。 */
      if (d && typeof d === 'object' && '__ooNativeSaveAsk' in d) {
        ;(function () {
          var reply = function (payload) { try { parent.postMessage(payload, '*') } catch (e) {} }
          try {
            var poc0 = document.querySelector('iframe[name="frameEditor"]')
            var w2 = poc0 && poc0.contentWindow
            var eng = w2 && w2.Asc && w2.Asc.editor
            if (!eng || typeof eng.asc_DownloadAs !== 'function') { reply({ __ooNativeSaveErr: '引擎不可达或无 asc_DownloadAs' }); return }
            if (!w2.Asc.asc_CDownloadOptions) { reply({ __ooNativeSaveErr: '无 asc_CDownloadOptions' }); return }
            var captured = null
            var origClick = HTMLAnchorElement.prototype.click
            HTMLAnchorElement.prototype.click = function () {
              try { if (this.download && this.href && this.href.indexOf('blob:') === 0) captured = this.href } catch (e) {}
              /* 捕获后吞掉，不弹系统下载框 */
            }
            eng.asc_DownloadAs(new w2.Asc.asc_CDownloadOptions(w2.Asc.c_oAscFileType.DOCX))
            var n = 0
            var t = setInterval(function () {
              n++
              if (captured || n > 50) {
                clearInterval(t)
                try { HTMLAnchorElement.prototype.click = origClick } catch (e) {}
                if (captured) {
                  fetch(captured).then(function (r) { return r.arrayBuffer() }).then(function (buf) {
                    var u8 = new Uint8Array(buf)
                    var b64 = ''
                    for (var i = 0; i < u8.length; i += 0x8000) b64 += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000))
                    try { parent.postMessage({ __ooNativeSaveData: btoa(b64) }, '*') } catch (e) {}
                  }).catch(function (e) { reply({ __ooNativeSaveErr: '产物读取失败 ' + String(e).slice(0, 80) }) })
                } else {
                  reply({ __ooNativeSaveErr: '25s 未捕获下载（引擎未产出）' })
                }
              }
            }, 500)
          } catch (e) {
            reply({ __ooNativeSaveErr: String(e).slice(0, 140) })
          }
        })()
      }
      /* ★ 2026-10-29 样式化提取（宿主保存兜底）：callCommand 直读文档模型——
         段落文字 + 字体/字号/加粗/斜体/下划线/颜色/对齐/样式名 + 表格拆行，纯 JSON 回传。
         这是 ooword 保存链唯一能拿到「编辑后内容 + 格式」的通道
         （forcesave 400、Editor.bin 打开时从源转出永不更新——字节级实锤）。 */
      if (d && typeof d === 'object' && '__ooStyledExtractAsk' in d) {
        var __seDone = false
        try {
          var __feS = document.querySelector('iframe[name="frameEditor"]') || document.querySelector('iframe')
          var __winS = __feS && __feS.contentWindow
          var __engS = null
          if (__winS) {
            if (__winS.Asc && __winS.Asc.editor && __winS.Asc.editor.callCommand) __engS = __winS.Asc.editor
            else if (__winS.editor && __winS.editor.callCommand) __engS = __winS.editor
          }
          if (!__engS) {
            try { parent.postMessage({ __ooStyledExtractData: null, __ooStyledExtractErr: 'poc:引擎不可达（深帧无 callCommand）' }, '*') } catch (e) {}
          } else {
            var __seFn = function () {
              /* ★ 2026-10-30 callCommand 的 cb 传值不可靠（实测重开后恒 undefined）——
               *  结果同时挂引擎窗口全局，宿主侧双通道取。 */
              var meta = { out: [] }
              var out = []
              try {
                var doc = Api.GetDocument()
                meta.hasGetJc = !!(doc.GetElement && doc.GetElementsCount)
                var el0 = doc.GetElementsCount ? doc.GetElement(0) : null
                if (el0) {
                  meta.t0 = el0.GetClassType ? el0.GetClassType() : ''
                  meta.fJc = typeof el0.GetJc
                  meta.fIndLeft = typeof el0.GetIndLeft
                  meta.fIndFirstLine = typeof el0.GetIndFirstLine
                  meta.fSpB = typeof el0.GetSpacingBefore
                  meta.fSpA = typeof el0.GetSpacingAfter
                  try { meta.vJc = el0.GetJc ? String(el0.GetJc()) : null } catch (eV1) {}
                  try { meta.vIndLeft = el0.GetIndLeft ? el0.GetIndLeft() : null } catch (eV2) {}
                  try { meta.vIndFirstLine = el0.GetIndFirstLine ? el0.GetIndFirstLine() : null } catch (eV3) {}
                  try { meta.vSpB = el0.GetSpacingBefore ? el0.GetSpacingBefore() : null } catch (eV4) {}
                  try { meta.vSpA = el0.GetSpacingAfter ? el0.GetSpacingAfter() : null } catch (eV5) {}
                }
                var n = doc.GetElementsCount()
                for (var i = 0; i < n; i++) {
                  var el = doc.GetElement(i)
                  var ct = el && el.GetClassType ? el.GetClassType() : ''
                  if (ct === 'paragraph') {
                    var p = { t: 'p', runs: [] }
                    /* ★ 2026-10-29 提取协议 v2：补齐版式三件套——对齐用 GetJc（builder API
                       真名，此前写的 GetAlignment 不存在 → 标题居中丢失）、缩进、段距 */
                    try { var jcv = el.GetJc ? el.GetJc() : null; if (jcv != null) p.align = String(jcv) } catch (eA) {}
                    try { if (p.align == null) { var al2 = el.GetAlignment ? el.GetAlignment() : null; if (al2 != null) p.align = String(al2) } } catch (eA2) {}
                    try { var st = el.GetStyle ? el.GetStyle() : null; if (st && st.GetName) p.style = String(st.GetName()) } catch (eSt) {}
                    try { if (el.GetIndLeft) { var il = el.GetIndLeft(); if (il != null) p.indLeft = il } } catch (eI1) {}
                    try { if (el.GetIndFirstLine) { var ifl = el.GetIndFirstLine(); if (ifl != null) p.indFirstLine = ifl } } catch (eI2) {}
                    try { if (el.GetSpacingBefore) { var sb = el.GetSpacingBefore(); if (sb != null) p.spBefore = sb } } catch (eS1) {}
                    try { if (el.GetSpacingAfter) { var sa = el.GetSpacingAfter(); if (sa != null) p.spAfter = sa } } catch (eS2) {}
                    /* ★ 2026-10-30 v3 行距（版式主差异源）：原版段落直排 line=576 exact——
                     *  16pt 字撑 28.8pt 行格，视觉行间空隙大；不提取则重建变单倍行距挤成 2 页 */
                    try { if (el.GetSpacingLine) { var sl = el.GetSpacingLine(); if (sl != null) p.spLine = sl } } catch (eS3) {}
                    try { if (el.GetSpacingLineRule) { var slr = el.GetSpacingLineRule(); if (slr != null) p.spLineRule = String(slr) } } catch (eS4) {}
                    var rn = el.GetElementsCount ? el.GetElementsCount() : 0
                    for (var j = 0; j < rn; j++) {
                      var r = el.GetElement(j)
                      if (!r || (r.GetClassType && r.GetClassType() !== 'run')) continue
                      var o = { text: '' }
                      try { o.text = String(r.GetText() || '') } catch (eT) {}
                      try {
                        var tp = r.GetTextPr ? r.GetTextPr() : null
                        if (tp) {
                          try { if (tp.GetBold) o.bold = !!tp.GetBold() } catch (e1) {}
                          try { if (tp.GetItalic) o.italic = !!tp.GetItalic() } catch (e2) {}
                          try { if (tp.GetUnderline) { var uu = tp.GetUnderline(); o.underline = !(uu === 'none' || uu === false || uu == null) } } catch (e3) {}
                          try { if (tp.GetFontSize) o.size = tp.GetFontSize() } catch (e4) {}
                          try { if (tp.GetFontFamily) { var ff = tp.GetFontFamily(); o.font = (ff && ff.GetName) ? ff.GetName() : String(ff || '') } } catch (e5) {}
                          try { if (tp.GetColor) { var cc = tp.GetColor(); if (cc && cc.GetHex) o.color = String(cc.GetHex()) } } catch (e6) {}
                        }
                      } catch (eP) {}
                      p.runs.push(o)
                    }
                    /* ★ 自动编号合成：段落级 GetText 含编号（"3."），run 级不含——
                       差值就是编号前缀（限长 12 防错位），合成首 run 保住编号 */
                    try {
                      var fullT = String(el.GetText ? el.GetText() : '')
                      var runT = ''
                      for (var q2 = 0; q2 < p.runs.length; q2++) runT += String(p.runs[q2].text || '')
                      if (fullT.length > runT.length && fullT.slice(fullT.length - runT.length) === runT) {
                        var pre = fullT.slice(0, fullT.length - runT.length)
                        if (pre !== '' && pre.replace(/\s/g, '') !== '' && pre.length <= 12) p.runs.unshift({ text: pre })
                      }
                    } catch (eN) {}
                    out.push(p)
                  } else if (ct === 'table') {
                    var tb = { t: 'tbl', rows: [] }
                    try {
                      var rc = el.GetRowsCount ? el.GetRowsCount() : 0
                      for (var ri = 0; ri < rc; ri++) {
                        var row = el.GetRow(ri)
                        var rowOut = []
                        var cc2 = row && row.GetCellsCount ? row.GetCellsCount() : 0
                        for (var ci = 0; ci < cc2; ci++) {
                          var cellTxts = []
                          var cell = row.GetCell(ci)
                          var cont = cell && cell.GetContent ? cell.GetContent() : null
                          if (cont) {
                            var cn = cont.GetElementsCount ? cont.GetElementsCount() : 0
                            for (var pi = 0; pi < cn; pi++) {
                              var cp = cont.GetElement(pi)
                              if (cp && cp.GetClassType && cp.GetClassType() === 'paragraph') {
                                try { cellTxts.push(String(cp.GetText ? cp.GetText() : '')) } catch (eCT) {}
                              }
                            }
                          }
                          rowOut.push(cellTxts.join('\\n'))
                        }
                        tb.rows.push(rowOut)
                      }
                    } catch (eTB) {}
                    out.push(tb)
                  }
                }
              } catch (eX) { return { err: String(eX), meta: meta } }
              try { window.__ooStyledExtractResult = { meta: meta, out: out } } catch (eG) {}
              return { meta: meta, out: out }
            }
            var __seCb = function (ret) {
              __seDone = true
              /* ★ 2026-10-30 cb 传值不可靠——优先从引擎窗口全局兜底取 */
              if ((!ret || !ret.out) && __winS && __winS.__ooStyledExtractResult) { try { ret = __winS.__ooStyledExtractResult } catch (eG2) {} }
              /* callCommand 内部异常时 fnS 回 {err:...}——透传给宿主诊断（S12a-E） */
              var isErr = ret && !Array.isArray(ret) && typeof ret.err === 'string'
              var payload = ret && ret.out ? ret.out : (isErr ? null : (ret || null))
              try { parent.postMessage({ __ooStyledExtractData: payload, __ooStyledExtractErr: isErr ? String(ret.err) : undefined, __ooStyledExtractMeta: ret && ret.meta ? ret.meta : undefined }, '*') } catch (e) {}
            }
            try {
              if (__engS.callCommand.length >= 4) __engS.callCommand(__seFn, false, false, __seCb)
              else __engS.callCommand(__seFn, { isCalculate: false, isRecalc: false }, __seCb)
            } catch (e3) {
              try { __engS.callCommand(__seFn, false, false, __seCb) } catch (e4) {
                try { parent.postMessage({ __ooStyledExtractData: null, __ooStyledExtractErr: 'poc:callCommand 调用异常 ' + String(e4).slice(0, 120) }, '*') } catch (e) {}
              }
            }
            setTimeout(function () {
              if (!__seDone) {
                /* 超时也先从引擎窗口全局兜底取（cb 传值不可靠的补丁） */
                var retT = null
                try { if (__winS && __winS.__ooStyledExtractResult) retT = __winS.__ooStyledExtractResult } catch (eG3) {}
                if (retT && retT.out) { __seDone = true; try { parent.postMessage({ __ooStyledExtractData: retT.out, __ooStyledExtractMeta: retT.meta }, '*') } catch (e) {} ; return }
                try { parent.postMessage({ __ooStyledExtractData: null, __ooStyledExtractErr: 'poc:4.5s 超时（callCommand 未回调）' }, '*') } catch (e) {}
              }
            }, 4500)
          }
        } catch (eSE) { try { parent.postMessage({ __ooStyledExtractData: null, __ooStyledExtractErr: 'poc:异常 ' + String(eSE).slice(0, 120) }, '*') } catch (e) {} }
      }
      /* ★ 2026-10-29 E2E 回归工具专用：清空最后一个非空段落的文字（模拟「删除部分字段」）。
         仅由 check-oo-e2e 通过宿主→poc postMessage 调用；回 __ooTestDeleteAck。 */
      if (d && typeof d === 'object' && '__ooTestDeleteAsk' in d) {
        var f2 = document.querySelector('iframe[name="frameEditor"]')
        var w2 = f2 && f2.contentWindow
        var eng2 = w2 && ((w2.Asc && w2.Asc.editor && w2.Asc.editor.callCommand) ? w2.Asc.editor : (w2.editor && w2.editor.callCommand ? w2.editor : null))
        if (!eng2) { try { parent.postMessage({ __ooTestDeleteAck: { err: 'no-engine' } }, '*') } catch (e) {} }
        else {
          var fnD = function () {
            try {
              var dc = Api.GetDocument()
              var nD = dc.GetElementsCount()
              var tg = -1
              for (var iD = nD - 1; iD >= 0; iD--) {
                var eE = dc.GetElement(iD)
                if (eE && eE.GetClassType && eE.GetClassType() === 'paragraph' && String(eE.GetText ? eE.GetText() : '').replace(/\s/g, '') !== '') { tg = iD; break }
              }
              if (tg < 0) return { err: '没有可删除的非空段落' }
              var eT2 = dc.GetElement(tg)
              var deleted = 0
              var rN = eT2.GetElementsCount ? eT2.GetElementsCount() : 0
              /* ★ 2026-10-30 改 run.Delete()——SetText 在此版本 ApiRun 上不存在（探针实锤） */
              for (var jD = rN - 1; jD >= 0; jD--) {
                var rD = eT2.GetElement(jD)
                if (rD && typeof rD.Delete === 'function') { try { rD.Delete(); deleted++ } catch (eS2) {} }
              }
              return { target: tg, deleted: deleted }
            } catch (eX2) { return { err: String(eX2) } }
          }
          var cbD = function (retD) {
            /* ★ 2026-10-30 程序化修改后自动置脏（双通道）——否则 dirty 硬门看不到，
             *   保存走主通道放行原文件 ⇒ 删除丢失（实测翻车链 00:01）。 */
            try {
              var feD = document.querySelector('iframe[name="frameEditor"]')
              var engD = feD && feD.contentWindow && feD.contentWindow.Asc && feD.contentWindow.Asc.editor
              if (engD && typeof engD.SetDocumentModified === 'function') engD.SetDocumentModified(true)
            } catch (eD1) {}
            try { if (typeof window.__setOoDirty === 'function') window.__setOoDirty(true, 'programmatic-delete') } catch (eD2) {}
            try { parent.postMessage({ __ooTestDeleteAck: retD || { err: 'null' } }, '*') } catch (e) {}
          }
          try {
            if (eng2.callCommand.length >= 4) eng2.callCommand(fnD, false, false, cbD)
            else eng2.callCommand(fnD, { isCalculate: false, isRecalc: false }, cbD)
          } catch (e5) { try { parent.postMessage({ __ooTestDeleteAck: { err: String(e5) } }, '*') } catch (e) {} }
          setTimeout(function () { try { parent.postMessage({ __ooTestDeleteAck: { err: 'poc:5s 超时' } }, '*') } catch (e) {} }, 5000)
        }
      }
      if (d && typeof d === 'object' && '__ooMaximized' in d) {
        window.__ooMax = d.__ooMaximized === true
        /* ★ 2026-10-18：最大化/还原会触发 hbox 重排，面板此刻的底色采样常是「重排前」的值
         *   （用户报「刚开始/最大化后参数颜色异常」）——重排完再重涂一遍。 */
        try { __repaintRefPanelSoon() } catch (eRM) {}
        // 离开最大化：收起窗内面板（普通模式由宿主浮卡接管，按钮常驻不隐藏）
        if (!window.__ooMax && __refPanel && __refPanel.style.display !== 'none') {
      __refPanel.style.transition = 'none'
      __refPanel.style.display = 'none'
          setTimeout(function () {
            __refPanel.style.display = 'none'
            __setRefPush(false)
            try { parent.postMessage({ __ooRefState: false }, '*') } catch (e) {}
          }, 260)
        }
      }
      if (d && typeof d === 'object' && '__ooRefState' in d) {
        try { fetch('/report', { method: 'POST', body: 'beacon: refState=' + d.__ooRefState }) } catch (e) {}
        if (window.__extRefBtn) {
          window.__extRefBtn.style.background = d.__ooRefState ? '#446995' : ''
          window.__extRefBtn.style.color = d.__ooRefState ? '#fff' : ''
        }
        if (window.__extRefStripBtn) {
          if (d.__ooRefState) window.__extRefStripBtn.classList.add('active')
          else window.__extRefStripBtn.classList.remove('active')
        }

      }
    } catch (err) {}
  })
  // ★ 常驻补插：web-apps 在窗口变化/最大化切换时会重建工具栏 DOM，注入的按钮会丢——
  //   定时器不自停，持续检查补插（各注入函数自带防重复守卫，空转开销可忽略）
  var __refBtnTimer = setInterval(function () {
    if (_ready) {
      __injectRefBtn(); __ensureRefPanel(); __wireSuggest()
      /* ★ 外部引用更新设置入口已挪宿主标题栏（2026-10-18 晚用户指令），不再挂 /poc 工具栏按钮 */
      // ★ 光标锚点（连线终点）：文档内 mouseup = 光标落点（100% 可靠，不依赖 OO 事件）
      //   + 滚动实时重绘
      try {
        var fE = document.querySelector('iframe[name="frameEditor"]') || document.querySelector('iframe')
        var ecc = fE && fE.contentDocument && fE.contentDocument.getElementById('editor-container')
        if (ecc && !ecc.__extCursorHooked) {
          ecc.__extCursorHooked = true
          // ★ OO 自绘滚动（无原生 scroll 事件）：滚轮/键盘累积自算滚动量（连线滚动跟随用）
          ecc.addEventListener('wheel', function (ev) {
            window.__docScrollAcc = (window.__docScrollAcc || 0) + (ev.deltaY || 0)
            try { window.__drawRefLines && window.__drawRefLines() } catch (e) {}
          }, { passive: true })
          fE.contentDocument.addEventListener('keydown', function (ev) {
            var step = { PageDown: Math.round((ecc.clientHeight || 800) * 0.9), PageUp: -Math.round((ecc.clientHeight || 800) * 0.9), ArrowDown: 40, ArrowUp: -40 }[ev.key]
            if (step) { window.__docScrollAcc = (window.__docScrollAcc || 0) + step; try { window.__drawRefLines && window.__drawRefLines() } catch (e) {} }
          })
          ecc.addEventListener('mouseup', function (ev) {
            var ecr = ecc.getBoundingClientRect()
            var stNow = window.__docScrollAcc || 0
            window.__cursorXY = [ev.clientX, ev.clientY]
            window.__cursorDoc = [ev.clientX - ecr.left, ev.clientY - ecr.top + stNow]
            try { window.__drawRefLines && window.__drawRefLines() } catch (e) {}
            // ★ 点击引用文字：命中胶囊 → 面板关着就打开（走模式分流）+ 画/强化关联线
            try {
              var mx0 = ev.clientX - ecr.left, my0 = ev.clientY - ecr.top + stNow
              var hitC = null
              ;(window.__extRefCapsules || []).forEach(function (c) {
                if (mx0 >= c.docX - 4 && mx0 <= c.docX + c.w + 4 && my0 >= c.docY - 25 && my0 <= c.docY + 5) hitC = c
              })
              if (hitC) {
                var pnlHidden = !(window.__extRefPanel && window.__extRefPanel.style.display !== 'none')
                if (pnlHidden) try { __refBtnAction() } catch (e2) {}
                window.__drawRefLines && window.__drawRefLines()
              }
            } catch (e) {}
          })


          ecc.addEventListener('scroll', function (ev) {
            try {
              var tg = ev.target
              if (tg && tg !== edoc2 && tg.scrollTop != null) { window.__docScroller = tg; window.__docScrollTop = tg.scrollTop }
            } catch (e) {}
            try { window.__drawRefLines && window.__drawRefLines() } catch (e) {}
          }, true)
        }
        // 辅助：api 光标事件若可用则叠加更新
        var api = fE && fE.contentWindow && fE.contentWindow.editor
        if (api && typeof api.attachEvent === 'function' && !api.__extCursorHooked) {
          api.__extCursorHooked = true
          // asc_onMouseMove 实证触发且参数带编辑器帧坐标 {YP,ZP}（鼠标出画布归零则忽略）
          api.attachEvent('asc_onMouseMove', function (pt) {
            try {
              if (pt && (pt.YP || pt.ZP)) { window.__cursorXY = [pt.YP, pt.ZP]; window.__drawRefLines && window.__drawRefLines() }
            } catch (e) {}
          })
        }
      } catch (e) {}
    }
  }, 2000)
  setTimeout(function () { clearInterval(__refBtnTimer) }, 60000)
  // ★ 编辑结束自愈巡检（独立常驻：上面 __refBtnTimer 60s 自停，这里不能跟停——
  //   用户任何时刻回车都可能触发自检。200ms 空转只是一次标志位判断，开销可忽略）
  var __evTimer = setInterval(function () { try { __editVerifyTick() } catch (eE2) {} }, 200)
  window.addEventListener('resize', function () {
    try {
      var fe = document.querySelector('iframe[name="frameEditor"]')
      var pnl = fe && fe.contentDocument && fe.contentDocument.getElementById('ext-ref-panel')
      if (pnl && pnl.style.display !== 'none') __alignRefPanel()
    } catch (e) {}
  })
  // ★ 参数引用面板：提取批注面板形态在编辑器文档内重建（同位同款：图标条右、工具栏下、圆角卡）
  var __refPanel = null, __refItems = []
  function __refDark() { return /theme-dark|contrast/.test(document.documentElement.className) }
  // ★ 面板配色实时采样：优先读真实批注面板的 computedStyle，回退 html+body 双源主题判断
  //  （只读 documentElement.className 会漏——实测该类可能为空而主题实际是暗色）
  function __refColors() {
    var fallback = function () {
      var dark = /theme-dark|contrast/.test(((document.documentElement || {}).className || '') + ' ' + ((document.body || {}).className || ''))
      return dark ? { bg: '#404040', fg: 'rgba(255,255,255,0.8)' } : { bg: '#ffffff', fg: '#333333' }
    }
    try {
      var fe = document.querySelector('iframe[name="frameEditor"]')
      var edoc = fe && fe.contentDocument
      var pc = edoc && edoc.querySelector('[id*=panel-comments]')
      if (pc) {
        var cs = getComputedStyle(pc)
        if (cs.backgroundColor && cs.backgroundColor !== 'rgba(0, 0, 0, 0)') return { bg: cs.backgroundColor, fg: cs.color }
      }
    } catch (e) {}
    return fallback()
  }
  function __ensureRefPanel() {
    try {
      var fe = document.querySelector('iframe[name="frameEditor"]') || document.querySelector('iframe')
      var edoc = fe && fe.contentDocument
      if (!edoc || edoc.getElementById('ext-ref-panel')) return
      // ★ 收养进原生容器：.left-panel.side-panel 是所有原生面板的宿主（in-flow，布局自动推挤编辑器）
      var cont = edoc.querySelector('.left-panel.side-panel')
      if (!cont) {
        var nb = edoc.getElementById('left-btn-comments')
        if (nb) nb.click()
        cont = edoc.querySelector('.left-panel.side-panel')
        if (nb) setTimeout(function () { try { nb.click() } catch (e2) {} }, 60)
      }
      if (!cont) return
      var pnl = edoc.createElement('div')
      pnl.id = 'ext-ref-panel'
      pnl.className = 'content-box' // 原生面板同款类
      var c0 = __refColors()
      pnl.style.cssText = 'position:relative;width:100%;height:100%;display:none;flex-direction:column;' +
        'background:' + (window.__refPalette ? window.__refPalette().bg : '#12171e') + ';color:' + (window.__refPalette ? window.__refPalette().text : '#c9d1d9') + ';' +
        'border-radius:0 8px 8px 0;overflow:hidden;font-size:12px'
      pnl.innerHTML =
        '<div style="display:flex;align-items:center;padding:8px 10px 6px;border-bottom:1px solid #21262d">' +
          '<span style="font-weight:600;font-size:12px;color:#e8edf2">参数引用</span><span style="flex:1"></span>' +
          '<span id="ext-ref-refresh" title="按画布现值重写全部域外引用的域数据，并全表重算（设置验证式更新）" style="cursor:pointer;padding:0 6px;color:#8b949e;font-size:13px">⟳</span>' +
          '<span id="ext-ref-close" title="收起" style="cursor:pointer;padding:0 4px;color:#8b949e;font-size:13px">✕</span>' +
        '</div>' +
        '<div style="padding:6px 8px">' +
          '<input id="ext-ref-q" placeholder="搜索节点 / 字段 / 值" style="width:100%;box-sizing:border-box;background:#0d1117;color:#c9d1d9;border:1px solid #30363d;border-radius:6px;padding:5px 8px;font-size:12px;outline:none">' +
        '</div>' +
        '<div id="ext-ref-body" style="flex:1;overflow:auto;padding:6px"></div>'
      cont.insertBefore(pnl, cont.firstChild) // 与原生面板同容器（自动获得槽位与推挤）
      // ★ 2026-10-10 焦点保持（用户指令「点= 点参数 不退出编辑」）：面板内所有
      //   mousedown 阻止默认——点击面板不再抢走单元格编辑器的焦点（blur 会提交编辑），
      //   参数引用条目才能以 execCommand insertText 追加进公式缓冲
      pnl.addEventListener('mousedown', function (eM) { eM.preventDefault() })
      __refPanel = pnl
      window.__extRefPanel = pnl // 互斥守卫从这里读状态
      /* ★ 2026-10-11 主题自适应调色板（用户指令：白色主题黑字）：跟随 OO 编辑器主题，
       *   暗色=GitHub Dark、亮色=应用亮色主题同款（浅底黑字）。__renderRefItems 每次重绘重取。 */
      window.__refPalette = function () {
        var dark = (function () {
          /* ★ 2026-10-18 修正主题检测（用户报「最大化后参数颜色异常、换主题不跟随」）：
           *   采样顺序改为「编辑器 documentElement/body 的主题类 → body 底色亮度 → /poc 页类名」。
           *   只读底色在最大化切换/主题切换的瞬间常常是 rgba(0,0,0,0)（透明，尚未重绘），
           *   亮度判断于是落空，回落到 /poc 页类名——而 /poc 的类名永远不跟编辑器主题，
           *   结果就是亮色主题下面板仍按暗色画。主题类是引擎自己挂的，最可信。 */
          try {
            var fEd = document.querySelector('iframe[name="frameEditor"]')
            var d2 = fEd && fEd.contentDocument
            var hEl = d2 && d2.documentElement
            var bEl = d2 && d2.body
            var cls = String((hEl && hEl.className) || '') + ' ' + String((bEl && bEl.className) || '')
            if (/theme-dark|contrast/.test(cls)) return true
            if (/theme-light|theme-classic-light|theme-classic/.test(cls)) return false
            var bgc = bEl ? (window.getComputedStyle ? getComputedStyle(bEl).backgroundColor : (bEl.style && bEl.style.backgroundColor)) : ''
            var mm = String(bgc).match(/[0-9]+/g)
            if (mm && mm.length >= 3) return (0.299 * mm[0] + 0.587 * mm[1] + 0.114 * mm[2]) < 128
          } catch (eD1) {}
          return __refDark()
        })()
        return dark
          ? { bg: '#12171e', bg0: '#0d1117', bg2: '#161b22', hover: '#1c232c', border: '#30363d', border2: '#21262d', title: '#e8edf2', text: '#c9d1d9', weak: '#6e7681', accent: '#58a6ff' }
          : { bg: '#f6f8fa', bg0: '#ffffff', bg2: '#ffffff', hover: '#eaeef2', border: '#d0d7de', border2: '#d8dee4', title: '#1f2328', text: '#1f2328', weak: '#57606a', accent: '#0969da' }
      }
      // ★ 2026-10-10 面板重建后恢复内容（OO 重建自身 UI 会清掉注入的面板 DOM；
      //   不恢复则节点列表/下拉全消失——用户实测「点参数后面板空了」同类）
      try { __wireRefSearch(); __renderRefItems() } catch (eR) {}
      pnl.querySelector('#ext-ref-close').onclick = function () { try { __setRefPanelVisible(false) } catch (e) {} }
      /* ★ 2026-10-11 刷新域值（用户指令「设置验证的方式 对域外变动点击触发更新」）：
       *   点 ⟳ = 按画布现值批量重写登记表中全部引用的域数据 A 列 + 全表重算。
       *   依赖单元格公式 =定义名 → 自动跟随，无需逐格改写。 */
      pnl.querySelector('#ext-ref-refresh').onclick = function () {
        var reg = window.__extRefRegistry || {}
        var keys = Object.keys(reg)
        if (!keys.length) { try { parent.postMessage({ __ooStatus: '[refresh] 暂无已插入的域外引用（先点参数插入至少一个）' }, '*') } catch (e0) {} return }
        var body = []
        var n = 0, miss = 0
        keys.forEach(function (k) {
          var en = reg[k]
          var val = null
          try {
            (__refItems || []).forEach(function (it) {
              (it.fields || []).forEach(function (f) {
                if (((it.name || '') + '.' + (f.name || '')) === k) val = f.value
              })
            })
          } catch (eX) {}
          if (val == null || val === '') { miss++; return }
          var numV = parseFloat(val)
          var cellVS = isNaN(numV) ? JSON.stringify(String(val)) : String(numV)
          body.push('try{dataWs.GetRange("A' + en.row + '").SetValue(' + cellVS + ')}catch(eR' + n + '){}')
          n++
        })
        if (!n) { try { parent.postMessage({ __ooStatus: '[refresh] 登记的 ' + keys.length + ' 项在画布中找不到现值' }, '*') } catch (e1) {} return }
        // ⚠ executeCommand 脚本禁顶层 return（解析错误整段不执行）
        var sc = 'try{var sheets=Api.GetSheets?Api.GetSheets():[];var dataWs=null;for(var i=0;i<sheets.length;i++){var nm2=sheets[i].GetName?sheets[i].GetName():"";if(nm2==="域数据"){dataWs=sheets[i];break}}if(!dataWs&&Api.AddSheet){dataWs=Api.AddSheet("域数据")}' + (n ? 'if(dataWs){' + body.join('') + 'try{dataWs.GetParent&&dataWs.GetParent().RecalculateAllFormulas&&dataWs.GetParent().RecalculateAllFormulas()}catch(eC){}}' : '') + '}catch(eS){}'
        var sent = __sendToRefPlugin({ __cbRefCmd: 'script', text: sc })
        try { parent.postMessage({ __ooStatus: '[refresh] 域值刷新已下发 ' + n + ' 项' + (miss ? '（' + miss + ' 项画布无现值跳过）' : '') + (sent ? '' : '（插件通道未达！）') }, '*') } catch (e2) {}
        // ★ 同步触发宿主字节级写回（模板落盘，下次打开也有值；与活文档刷新双轨并行）
        try { parent.postMessage({ __ooRefreshRef: Date.now() }, '*') } catch (e3) {}
      }
      __wireRefWheel()
    } catch (e) {}
  }
  /* 参数引用面板自己接管滚轮（2026-10-05 用户报「最大化 ooexcel 时面板滚不动」）。

     症状：面板正常显示、overflow:auto、scrollHeight(4148) 远大于 clientHeight(755)、
     手动设 scrollTop 也能到上限，但真实鼠标滚轮完全无效。

     定位过程（scripts/_dbg_oowheel.mjs 的 [D]~[K] 逐条实测）：
       [D] wheel 事件确实派发到了 #ext-ref-body（cancelable=true），
            但同一个事件在 document 冒泡阶段没有出现 => 有人中途掐断。
       [G][I] defaultPrevented 全程 false => 不是 preventDefault 取消的默认滚动。
       [H] 决定性：把面板整块挪到同一个 document 的 body 根部（position:fixed），
            滚轮立刻恢复正常 => 拦截器在 OO 的祖先链上，面板自己无辜。
     机理：Chromium 里只要祖先链上有 non-passive 的 wheel 监听器，滚动决策就延后到
     它执行完之后，由它决定滚哪个容器。OO 是自绘滚动编辑器（本文件上方注释亦写明
     OO 自绘滚动、无原生 scroll 事件），会接管 wheel；
     面板嵌在它的 .left-panel.side-panel 槽位里，就被一起接管了。

     修法：不再依赖浏览器默认滚动，改为自己滚。
       - 监听器必须 non-passive，否则 preventDefault 无效。
       - 手动改 scrollTop 并 stopPropagation，不让它冒到 OO 去。
       - 到边界就放行（不 preventDefault）：面板滚到头/底时，
         应该让 OO 的文档滚动接手，而不是把滚轮吃掉造成卡在边界。
       - deltaMode 归一：0=像素 / 1=行(乘16) / 2=页(乘 clientHeight)，
         不同设备与浏览器给的单位不一样，不归一会滚得忽快忽慢。
       - 幂等：__ensureRefPanel 由定时器驱动会重复跑，用 __extWheelHooked 守卫，
         否则重复挂载会让滚动量翻倍。 */
  function __wireRefWheel() {
    if (!__refPanel) return
    var body = __refPanel.querySelector('#ext-ref-body')
    if (!body || body.__extWheelHooked) return
    body.__extWheelHooked = true
    var hook = function (e) {
      var max = body.scrollHeight - body.clientHeight
      if (max <= 0) return                                  // 本来就没得滚，放行
      var mul = e.deltaMode === 1 ? 16 : e.deltaMode === 2 ? body.clientHeight : 1
      var dy = (e.deltaY || 0) * mul
      if (!dy) return
      // 到边界放行：让 OO 的自绘滚动接手（面板滚到头就该让文档动）
      if ((dy < 0 && body.scrollTop <= 0) || (dy > 0 && body.scrollTop >= max)) return
      e.preventDefault()
      e.stopPropagation()
      body.scrollTop += dy
    }
    // 留handler 引用：__ensureRefPanel 由定时器驱动、面板 DOM 会被 OO 重建，
    // 出问题时需要能定位到「当前挂的是哪个函数」；诊断探针也靠它做反向验证。
    body.__extWheelHook = hook
    body.addEventListener('wheel', hook, { passive: false })
  }
  /** 只涂壳（标题栏/搜索框/滚动条），不动条目 DOM —— 创建时与主题切换都要用 */
  function __paintRefShell(P) {
    if (!__refPanel) return
    try {
      __refPanel.style.background = P.bg
      __refPanel.style.color = P.text
      var qEl0 = __refPanel.querySelector('#ext-ref-q')
      if (qEl0) { qEl0.style.background = P.bg0; qEl0.style.color = P.text; qEl0.style.borderColor = P.border }
      var r1 = __refPanel.querySelector('#ext-ref-refresh'); if (r1) r1.style.color = P.weak
      var c1 = __refPanel.querySelector('#ext-ref-close'); if (c1) c1.style.color = P.weak
      var hd0 = __refPanel.firstElementChild
      if (hd0) { hd0.style.borderBottomColor = P.border2; var tt0 = hd0.querySelector('span'); if (tt0) tt0.style.color = P.title }
    } catch (eP0) {}
  }
  /**
   * 主题/最大化变化后重涂面板（2026-10-18 用户报「最大化后参数颜色异常，换主题也不跟着变」）
   *
   * 根因：面板颜色是**渲染时**从调色板取的，而宿主切主题只调了编辑器官方 Themes.setTheme，
   * 没有任何人通知面板重绘 ⇒ 面板停在上一套配色。最大化同理（重排后底色采样也可能落空）。
   * 这里在多个时间点重涂（0/120/400/900ms）：主题切换与 hbox 重排都是异步的，
   * 一次重绘经常采到「还没重绘完」的底色。
   */
  function __repaintRefPanelSoon() {
    var delays = [0, 120, 400, 900]
    for (var i = 0; i < delays.length; i++) {
      ;(function (ms) {
        setTimeout(function () {
          try {
            if (!__refPanel) return
            if (__refPanel.style.display === 'none') {
              // 收着也把壳涂对，下次展开不用再等一次重绘
              __paintRefShell((window.__refPalette || function () { return { bg: '#12171e', bg0: '#0d1117', bg2: '#161b22', hover: '#1c232c', border: '#30363d', border2: '#21262d', title: '#e8edf2', text: '#c9d1d9', weak: '#6e7681', accent: '#58a6ff' } })())
              return
            }
            __renderRefItems()
          } catch (eRP) {}
        }, ms)
      })(delays[i])
    }
  }
  function __renderRefItems() {
    /* ★ 2026-10-11 主题自适应（用户指令：白色主题黑字）：每次重绘重取调色板并重涂壳 */
    var P = (window.__refPalette || function () { return { bg: "#12171e", bg0: "#0d1117", bg2: "#161b22", hover: "#1c232c", border: "#30363d", border2: "#21262d", title: "#e8edf2", text: "#c9d1d9", weak: "#6e7681", accent: "#58a6ff" } })()
    __paintRefShell(P)
    if (!__refPanel) return
    var body = __refPanel.querySelector('#ext-ref-body')
    var qEl = __refPanel.querySelector('#ext-ref-q')
    var q = qEl ? qEl.value.toLowerCase() : ''
    // ★ 2026-10-10 用户实测「点击参数就折叠」：展开状态原挂 it.__open，宿主重推
    //   __ooRefData 时对象被整体替换 ⇒ 状态全丢。改挂模块级 __refOpenSet（按 id 持久化）。
    if (!window.__refOpenSet) window.__refOpenSet = {}
    var st0 = body ? body.scrollTop : 0
    body.innerHTML = ''
    if (!__refItems.length) { body.innerHTML = '<div style="padding:8px;color:#888">画布中暂无可引用节点</div>'; return }
    var shown = 0
    __refItems.forEach(function (it) {
      var fields = (it.fields || []).filter(function (f) {
        if (!q) return true
        return it.name.toLowerCase().indexOf(q) >= 0 || f.name.toLowerCase().indexOf(q) >= 0 || (f.value || '').toLowerCase().indexOf(q) >= 0
      })
      if (q && !fields.length && it.name.toLowerCase().indexOf(q) < 0) return
      shown++
      var open = !!window.__refOpenSet[it.id] || !!q
      if (!q) window.__refOpenSet[it.id] = open
      var head = document.createElement('div')
      head.style.cssText = 'display:flex;align-items:center;gap:6px;padding:5px 8px;margin:2px 0;border-radius:4px;cursor:pointer;font-weight:600;font-size:12px;color:' + P.title + ';border:1px solid ' + P.border + ';border-left:3px solid ' + P.accent + ';background:' + P.bg2
      var arrow = document.createElement('span')
      arrow.textContent = open ? '▾' : '▸'
      var nm = document.createElement('span')
      nm.textContent = it.name
      var ct = document.createElement('span')
      ct.style.cssText = 'flex:1;text-align:right;color:' + P.weak + ';font-size:10px;font-weight:normal'
      var regCnt = 0
      try {
        var reg0 = window.__extRefRegistry || {}
        Object.keys(reg0).forEach(function (k0) { if (k0.indexOf((it.name || '') + '.') === 0) regCnt++ })
      } catch (eRC) {}
      ct.setAttribute('data-count-for', it.id)
      ct.textContent = String(regCnt)
      head.appendChild(arrow); head.appendChild(nm); head.appendChild(ct)
      head.onclick = function () {
        // ★ 2026-10-10 编辑态（公式输入中）：点节点 → 弹出数据域下拉（OO 自动完成风格），
        //   条目点击 = 该参数的引用插入公式缓冲；非编辑态 = 展开/折叠（原行为）
        var areaEl = null
        try { areaEl = fE.contentDocument.getElementById('area_id') } catch (eA) {}
        if (areaEl && String(areaEl.value || '').length > 0) { try { parent.postMessage({ __ooStatus: '[poc] 编辑态：弹出数据域下拉' }, '*') } catch (eL) {} __showFieldDropdown(it, head); return }
        try { parent.postMessage({ __ooStatus: '[poc] 非编辑态：展开/折叠' }, '*') } catch (eL2) {}
        window.__refOpenSet[it.id] = !open
        __renderRefItems()
        // ★ 展开集回传宿主（最大化/常规内容同步）
        try {
          var __ids = []
          Object.keys(window.__refOpenSet).forEach(function (k3) { if (window.__refOpenSet[k3]) __ids.push(k3) })
          parent.postMessage({ __ooRefUi: { openIds: __ids } }, '*')
        } catch (eU2) {}
      }
      body.appendChild(head)
      if (!open) return
      fields.forEach(function (f) {
        var row = document.createElement('div')
        var usedKey = it.id + '#' + f.id
        var used = !!(window.__refUsedSet || {})[usedKey]
        row.style.cssText = 'display:flex;gap:6px;padding:3px 6px 3px 24px;cursor:pointer;border-radius:4px;align-items:baseline;border:1px solid ' + (used ? P.border : P.border2) + ';background:' + (used ? P.hover : P.bg0)
        row.setAttribute('data-key', it.id + '#' + f.id)
        row.onmouseenter = function () { row.style.background = P.hover }
        row.onmouseleave = function () { row.style.background = '' }
        /* ★ 2026-10-17 用户指令「对参数引用的地址加个颜色区分」：
         *   小色块 = 该参数在文档里的引用格底色（色值由宿主算好放进 f.color，
         *   这里只负责显示 —— 哈希算法不在这份脚本里复制，避免两处口径漂移）。 */
        if (f.color) {
          var sw = document.createElement('span')
          sw.style.cssText = 'width:9px;height:9px;flex:0 0 auto;align-self:center;border-radius:2px;border:1px solid rgba(128,128,128,.45);background:#' + String(f.color).slice(-6)
          sw.title = '该参数在文档中的引用格底色'
          row.appendChild(sw)
        }
        var n1 = document.createElement('span')
        // ★ 2026-10-10 用户指令：拆分点击区——点节点名=插入引用地址；点数值=插入数值
        n1.textContent = f.name
        n1.style.cssText = 'color:' + P.text + ';cursor:pointer'
        n1.title = '点击插入引用地址（公式栏显示 参数名，格子显示其值）'
        var sp = document.createElement('span')
        sp.style.cssText = 'flex:1'
        var v1 = document.createElement('span')
        v1.style.cssText = 'color:' + P.weak + ';font-size:10.5px;max-width:110px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;cursor:pointer'
        v1.textContent = f.value || ''
        v1.title = '点击插入数值（' + (f.value || '') + '）'
        row.appendChild(n1); row.appendChild(sp); row.appendChild(v1)
        row.title = '点参数名=插入引用地址（公式栏显示参数名）；点数值=插入数值；双击也插入（面板保持打开）'
        // ★ 2026-10-10 用户指令：点击分区——n1(参数名)=引用地址；v1(数值)=数值；双击=数值+收起面板
        var __pick = function (mode) {
          try { parent.postMessage({ __ooRefPick: { nodeId: it.id, fieldId: f.id, fieldName: f.name, value: f.value, mode: mode } }, '*') } catch (e) {}
          /* ★ 2026-10-11 用户指令：点击后留选中痕迹（✓ + 高亮底，会话内持久） */
          try {
            var uk = it.id + '#' + f.id
            window.__refUsedSet = window.__refUsedSet || {}
            window.__refUsedSet[uk] = true
            row.style.background = P.hover
            row.style.borderColor = P.border
          } catch (eM3) {}
          /* ★ 计数即时刷新（用户指令：显示引用数，删除不计入） */
          try {
            var ctEl = __refPanel && __refPanel.querySelector('[data-count-for="' + it.id + '"]')
            if (ctEl) {
              var rc2 = 0
              var reg2 = window.__extRefRegistry || {}
              Object.keys(reg2).forEach(function (k2) { if (k2.indexOf((it.name || '') + '.') === 0) rc2++ })
              ctEl.textContent = String(rc2)
            }
          } catch (eC3) {}
        }
        n1.ondblclick = function (e) { e.stopPropagation(); __pick('ref') }
        v1.ondblclick = function (e) { e.stopPropagation(); __pick('value') }

        row.ondblclick = function () { __pick('value') }
        body.appendChild(row)
      })
    })
    if (!shown && q) body.innerHTML = '<div style="padding:8px;color:#888">无匹配项</div>'
    if (body && st0) body.scrollTop = st0
  }
  /* ★ 2026-10-10 编辑态数据域下拉（用户指令「输入= 点节点 会有下拉 类似公式选择
   *   那种交互 展示节点内的数据域连接提示」）：OO 自动完成风格深色弹层，
   *   列出节点全部参数（图标+参数名+现值预览），条目点击 = 该参数引用追加进公式缓冲。
   *   定位在节点行下方；点外部关闭；面板已带 mousedown preventDefault 焦点保持。 */
  function __showFieldDropdown(it, anchorEl) {
    var old = document.getElementById('ext-fdd')
    if (old) old.remove()
    var dd = document.createElement('div')
    dd.id = 'ext-fdd'
    dd.style.cssText = 'position:absolute;z-index:9600;background:#2b2b2b;border:1px solid #454545;border-radius:6px;box-shadow:0 6px 16px rgba(0,0,0,.45);padding:4px 0;min-width:230px;max-height:320px;overflow:auto'
    var fields = it.fields || []
    fields.forEach(function (f) {
      var item = document.createElement('div')
      item.style.cssText = 'display:flex;align-items:center;gap:8px;padding:6px 12px;cursor:pointer;color:#e8e8e8;font-size:13px'
      item.innerHTML = '<span style="display:inline-flex;width:18px;height:18px;align-items:center;justify-content:center;border:1px solid #6e7681;border-radius:4px;font-size:10px;color:#9ab">ƒx</span>' +
        '<span style="font-weight:600">' + (it.name ? it.name + '_' : '') + f.name + '</span>' +
        (f.value ? '<span style="color:#8b949e;margin-left:auto;font-size:11px;max-width:90px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + String(f.value).slice(0, 14) + '</span>' : '')
      item.onmouseenter = function () { item.style.background = '#0d419d' }
      item.onmouseleave = function () { item.style.background = '' }
      item.onclick = function (e) {
        e.stopPropagation()
        dd.remove()
        try { parent.postMessage({ __ooRefPick: { nodeId: it.id, fieldId: f.id, fieldName: f.name, value: f.value, mode: 'ref' } }, '*') } catch (e2) {}
      }
      item.onmouseenter && dd.addEventListener('mousedown', function (eD) { eD.preventDefault() })
      dd.appendChild(item)
    })
    if (!fields.length) {
      var empty = document.createElement('div')
      empty.style.cssText = 'padding:8px 12px;color:#8b949e;font-size:12px'
      empty.textContent = '（该节点无可引用参数）'
      dd.appendChild(empty)
    }
    var r = anchorEl.getBoundingClientRect()
    /* ★ 2026-10-11 修坐标空间（用户报「点节点参数不显示+下方节点被盖」）：anchor 在
     *   frameEditor 文档内，getBoundingClientRect 是帧内坐标；dd 挂在 /poc 页 body——
     *   帧内坐标直接当页面坐标用会叠到面板顶部盖住下方行。加 iframe 页面偏移换算。 */
    var feOff = { x: 0, y: 0 }
    try {
      var feEl = document.querySelector('iframe[name="frameEditor"]')
      if (feEl) { var fr0 = feEl.getBoundingClientRect(); feOff.x = fr0.left; feOff.y = fr0.top }
    } catch (eF0) {}
    dd.style.left = Math.max(4, feOff.x + r.left) + 'px'
    dd.style.top = (feOff.y + r.bottom + 4) + 'px'
    document.body.appendChild(dd)
    setTimeout(function () {
      document.addEventListener('mousedown', function h(e) {
        if (!dd.contains(e.target)) { dd.remove(); document.removeEventListener('mousedown', h) }
      })
    }, 0)
  }
  function __wireRefSearch() {
    var qEl = __refPanel && __refPanel.querySelector('#ext-ref-q')
    if (qEl && !qEl.__wired) {
      qEl.__wired = true
      qEl.oninput = function () {
        __renderRefItems()
        // ★ 内嵌面板搜索词回传宿主（最大化/常规内容同步；fromPoc 侧不回推，无乒乓）
        try { parent.postMessage({ __ooRefUi: { q: String(qEl.value || '') } }, '*') } catch (e) {}
      }
    }
  }
  // ★ 推挤主界面：#editor-container 左移让位（同批注面板行为）
  function __setRefPush(on) {
    // 推挤已由原生容器布局接管（in-flow）——此函数仅清理历史遗留 margin
    try {
      var fe = document.querySelector('iframe[name="frameEditor"]') || document.querySelector('iframe')
      var ec = fe && fe.contentDocument && fe.contentDocument.getElementById('editor-container')
      if (ec && ec.style.marginLeft) ec.style.marginLeft = ''
    } catch (e) {}
  }
  // ★ 互斥：打开参数引用时，收拢 OO 当前激活的左侧面板（点它的分类按钮即折叠）
  function __collapseOoPanel() {
    try {
      var fe = document.querySelector('iframe[name="frameEditor"]') || document.querySelector('iframe')
      var edoc = fe && fe.contentDocument
      if (!edoc) return
      // ① 走 OO 自己的关闭路径：点激活的分类按钮（同步其内部状态，防重排时“复活”）
      var act = edoc.querySelector('.btn-category.active')
      if (act && act.id !== 'ext-btn-left-ooref') act.click()
      // ② 兜底：60ms 后仍有可见的原生面板就直接藏
      setTimeout(function () {
        try {
          edoc.querySelectorAll('[id*="panel-"]').forEach(function (pc) {
            if (pc.getBoundingClientRect().width > 50) pc.style.display = 'none'
          })
        } catch (e) {}
      }, 60)
    } catch (e) {}
  }
  // ★ 对齐器：套用原生对齐规则（常量来自实测：图标条宽39、工具栏高98、面板宽259）。
  //   不做运行时测量——无头/图标未加载环境里 getBoundingClientRect 会返回 0 污染样式。
  function __alignRefPanel() {
    // 几何已由原生容器布局全权管理（in-flow）——此处只同步主题配色
    try {
      if (!__refPanel) return
      var cc = __refColors()
      __refPanel.style.background = cc.bg
      __refPanel.style.color = cc.fg
    } catch (e) {}
  }
  function __setRefPanelVisible(show) {
    __ensureRefPanel()
    if (!__refPanel) return
    var cont = __refPanel.parentElement
    var edoc2 = __refPanel.ownerDocument
    var lm = edoc2.getElementById('left-menu') // OO 用它的内联宽度(40↔300)控制面板开合推挤
    // ★ 宽度门控（用户 2026-10-01）：可用宽 <900 → 不显示（数据引用整体让位）；跨过阈值 → 自动恢复
    var __wOK = function () { try { return (edoc2.defaultView.innerWidth || 1200) >= 900 } catch (e) { return true } }
    if (show && !__wOK()) { window.__refHiddenByWidth = true; __refPanel.style.display = 'none'; return }
    if (show) window.__refHiddenByWidth = false
    // OO 的 hbox 布局是 JS 计算：宽度变更后必须触发 resize 才会重排编辑器
    var __kickLayout = function () { try { (edoc2.defaultView || (edoc2.defaultView)).dispatchEvent(new Event('resize')) } catch (e) {} }
    if (show) {
      __collapseOoPanel()
      if (cont) cont.style.display = 'block' // 容器可能被 OO 在自身面板全关时隐藏
      if (lm) lm.style.width = '300px' // ★ 撑开左菜单（同原生面板打开时的推挤机制）
      __kickLayout()
      __refPanel.style.transition = 'none'
      __refPanel.style.transform = 'none'
      __refPanel.style.display = 'flex'
      __alignRefPanel(); __wireRefSearch(); __renderRefItems()
      // ★ 压制 OO 窄窗响应式折叠：!important 样式表 > OO 的内联 40px（级联规则），
      //   中等宽度下面板保持停靠（编辑器右移），不再退化为遮挡
      try {
        var stX = edoc2.getElementById('ext-ref-style')
        if (!stX) {
          stX = edoc2.createElement('style')
          stX.id = 'ext-ref-style'
          edoc2.head.appendChild(stX)
        }
        stX.textContent = '#left-menu{width:300px!important;min-width:300px!important;max-width:none!important}.left-panel.side-panel{width:100%!important}'
      } catch (e) {}
    } else {
      var stRm = edoc2.getElementById('ext-ref-style')
      if (stRm) stRm.remove() // 移除压制样式表 → OO 折叠行为自动还原
      window.__extRefLines = [] // 清持久连线
      var lsSvg = edoc2.getElementById('ext-ref-lines-svg')
      if (lsSvg) lsSvg.remove()
      __refPanel.style.transition = 'none'
      __refPanel.style.display = 'none'
      window.__extRefLines = [] // 清持久连线
      var lsSvg = edoc2.getElementById('ext-ref-lines-svg')
      if (lsSvg) lsSvg.remove()
      __refPanel.style.transition = 'none'
      __refPanel.style.display = 'none'
      // ★ 关闭也即时（与批注一致）：面板即隐 + 左菜单即收 + 重排
      try {
        var anyNative = Array.from(cont.children).some(function (c) { return c.id !== 'ext-ref-panel' && c.getBoundingClientRect().width > 50 })
        if (lm && !anyNative) { lm.style.width = '40px'; __kickLayout() }
      } catch (e) {}
    }
    try { parent.postMessage({ __ooRefState: show }, '*') } catch (e) {}
  }
  // ★ 宽度跨阈值：自动隐藏 / 恢复（窗口 resize 时跟随）
  try {
    var feW = edoc2.defaultView
    if (feW && !feW.__refResizeHooked) {
      feW.__refResizeHooked = true
      feW.addEventListener('resize', function () {
        try {
          var okW = (feW.innerWidth || 1200) >= 900
          var vis = __refPanel && __refPanel.style.display !== 'none'
          if (!okW && vis) { __refPanel.style.display = 'none'; window.__refHiddenByWidth = true; try { parent.postMessage({ __ooRefState: false }, '*') } catch (e2) {} }
          if (okW && !vis && window.__refHiddenByWidth) { window.__refHiddenByWidth = false; __setRefPanelVisible(true) }
        } catch (e) {}
      })
    }
  } catch (e) {}
  function __toggleRefPanel() {
    __ensureRefPanel()
    if (!__refPanel) return
    var show = __refPanel.style.display === 'none'
    __setRefPanelVisible(show)
    // ★ 打开且数据为空 → 向宿主拉取（编辑器加载早期推送的消息可能已丢失）
    if (show && !__refItems.length) { try { parent.postMessage({ __ooRefDataRequest: 1 }, '*') } catch (e) {} }
  }
  // ★ 诊断探针：ascdesktop 协议在 Electron 是否可达（结果进 /dumplog）
  setTimeout(function () {
    fetch('ascdesktop://fonts/C:/Windows/Fonts/arial.ttf').then(function (r) {
      return r.arrayBuffer().then(function (b) { fetch('/report', { method: 'POST', body: 'ascdesktop-test: OK status=' + r.status + ' len=' + b.byteLength }) })
    }).catch(function (e) { fetch('/report', { method: 'POST', body: 'ascdesktop-test: FAIL ' + String(e).slice(0, 90) }) })
  }, 2000)
// ★ 就绪轮询（兜底信号）：正文出现 "Page N of M" 且 "Loading document" 已消失才算就绪——
//   只看 "Page N of M" 会在桌面流卡死态误判（状态栏默认就显示 Page 1 of 1）。
//   _ready 由 onDocumentReady 置位；轮询不得覆盖回调写下的状态。
var _poll = setInterval(function () {
  if (_err || _ready) return
  try {
    var f = document.querySelector('iframe')
    var txt = f && f.contentDocument && f.contentDocument.body ? f.contentDocument.body.innerText : ''
    if (txt && /Page\s+\d+\s+of\s+\d+/.test(txt) && !/Loading document/i.test(txt)) { _ready = true; setSt('documentReady \u2705(poll)'); clearInterval(_poll) }
    else setSt('loading doc\u2026; src=' + srcFile)
    // ★ 2026-10-08 自救：open 持续 >90s（DS 启动窗口期/转换器未就绪）→ 自动重载一次；
    //   重载后仍卡（sessionStorage 标记已存在）→ 明示请重启应用，不再默默转圈。
    window.__ooWait = (window.__ooWait || 0) + 1
    if (window.__ooWait === 90) {
      var k = '__oo_retry_' + srcFile
      if (!window.sessionStorage.getItem(k)) {
        window.sessionStorage.setItem(k, '1')
        setSt('\u6587\u6863\u670d\u52a1\u672a\u5c31\u7eea\u2014\u20145 \u79d2\u540e\u81ea\u52a8\u91cd\u8bd5\u4e00\u6b21\u2026')
        setTimeout(function () { location.reload() }, 5000)
      } else {
        setSt('\u6587\u6863\u670d\u52a1\u672a\u5c31\u7eea\u2014\u2014\u8bf7\u70b9\u6807\u9898\u680f\u300c\u91cd\u542f\u300d\u6309\u94ae\u540e\u91cd\u5f00\u8282\u70b9')
        clearInterval(_poll)
      }
    }
  } catch (e) { setSt('poll err: ' + e) }
}, 1000)
setTimeout(function () { clearInterval(_poll) }, 180000)
</script></body></html>`, 'text/html; charset=utf-8')
      return
    }

    /* ── 字体目录（动态 AllFonts.js 的文件源） ── */
    if (p.startsWith('/fonts/')) {
      const rel = decodeURIComponent(p.replace('/fonts/', ''))
      const f = path.join(FONTS_DIR, rel)
      if (!f.startsWith(FONTS_DIR)) { send(res, 403, 'forbidden'); return }
      if (!fs.existsSync(f)) {
        if (rel.endsWith('plugins.json')) { send(res, 200, JSON.stringify({ pluginsData: [] }), 'application/json'); return }
        if (rel.endsWith('themes.json')) { send(res, 200, JSON.stringify({ themes: [] }), 'application/json'); return }
        send(res, 404, 'nf: ' + rel); return
      }
      sendFile(res, req, 200, f, MIME[path.extname(f).toLowerCase()] || 'application/octet-stream', { negotiate: true }); return true /*★ 流式分支必须显式 true（2026-10-10）: 返回 undefined 会被裸调用方当未处理再补刀，响应体被截*/
      return
    }
    // ★ 字体直出：buildAllFonts 把 AllFonts.js 的绝对路径改写为 /sysfonts/<basename>，
    //   这里按 _fontPathMap（basename→原绝对路径，覆盖 13 个根目录含子目录）精确回源；
    //   查不到再回退 C:/Windows/Fonts。路径来自本服务自己生成的映射，无遍历风险。
    if (p.startsWith('/sysfonts/')) {
      const name = decodeURIComponent(p.replace('/sysfonts/', ''))
      let f = _fontPathMap && _fontPathMap[name]
      // ★ 2026-10-10：glue 剥前缀后可能传来全路径（/sysfonts/C:/Windows/Fonts/arial.ttf）
      //   ——取 basename 再查映射/系统字体目录
      if (!f && /[/\\]/.test(name)) {
        const base = name.split(/[\\/]/).pop()
        f = (_fontPathMap && _fontPathMap[base]) || (/^[\w.\- ]+$/.test(base) ? path.join('C:/Windows/Fonts', base) : null)
      }
      if (!f && /^[\w.\- ]+$/.test(name)) f = path.join('C:/Windows/Fonts', name)
      if (!f || !fs.existsSync(f)) { send(res, 404, 'nf: ' + name); return }
      sendFile(res, req, 200, f, MIME[path.extname(f).toLowerCase()] || 'application/octet-stream', { negotiate: true }); return true /*★ 流式分支必须显式 true（2026-10-10）: 返回 undefined 会被裸调用方当未处理再补刀，响应体被截*/
      return
    }

    /* ── 静态资产（documenteditor index.html 动态注桥；AllFonts.js 动态重写） ── */
    if (p.startsWith('/oo-assets/')) {
      const rel = p.replace('/oo-assets/', '')
      const f = path.join(ASSETS, rel)
      if (!f.startsWith(ASSETS)) { send(res, 403, 'forbidden'); return }
      // ★ AllFonts.js 必须重写：原版是 DesktopEditors 绝对本地路径 → 浏览器端字体加载全挂
      if (rel === 'sdkjs/common/AllFonts.js') { send(res, 200, buildAllFonts(), 'application/javascript'); return }
      // ★ service worker 保持 404：12:5x 全链成功时 SW 注册失败（404）；一旦注册成功，
      //   document_editor_service_worker 会拦截 /doc 协作请求 → 卡 LocalStartOpen（实测回归）
      if (rel === 'document_editor_service_worker.js') { send(res, 404, 'sw disabled'); return }
      // ★ sdk-all* 字体加载器剥前缀（2026-09-30）：桌面构建无条件 `XMLHttpRequest("ascdesktop://fonts/"+Va)`
      //   拉 .ttf（sdk-all.js @AscFonts.sje.BSa，无 http 分支），而 Electron 的
      //   protocol.handle('ascdesktop') 实测从不被命中（全天 0 HIT）→ 字体全空 → onOpenDocument
      //   之后卡 "Loading document"。删除该字面量后 XHR 直接用 AllFonts 重写出的
      //   /sysfonts/<basename>（同源 HTTP，_fontPathMap 多根直出），浏览器与 Electron 同路。
      if (/^sdkjs\/(word|cell|slide|visio)\/sdk-all(-min)?\.js$/.test(rel)) {
        const js = fs.readFileSync(f, 'utf8').split('ascdesktop://fonts/').join('')
        send(res, 200, js, 'application/javascript')
        return
      }
      // ★ 引擎 glue isLocal 强制 false：internal_isLocal() 为真时 wasm/字体的全部二进制加载
      //   强制 ascdesktop://（Electron handler 0 命中）→ 字体加载失败 → 度量垃圾 → 瓦片 OOM
      if (/^sdkjs\/common\/(libfont\/engine\/fonts|hash\/hash\/engine|spell\/spell\/spell|zlib\/engine\/zlib)(-ie)?\.js$/.test(rel)) {
        const js = fs.readFileSync(f, 'utf8')
          .replace(/(?<!function )internal_isLocal\(\)/g, 'false').replace(/(?<!function )(?<![a-zA-Z_])isLocal\(\)/g, 'false')
          // ★ 2026-10-10（无头浏览器实测）：glue 的 wasm 侧 fetch('ascdesktop://fonts/'+全路径)
          //   在纯浏览器没有协议处理器 → 字体全挂 → onError 死循环。与 sdk-all 同款剥前缀，
          //   改走同源 /sysfonts/（服务端按 basename 回源）
          .split('ascdesktop://fonts/').join('/sysfonts/')
        send(res, 200, js, 'application/javascript')
        return
      }
      // ★ 字体下拉瓦片防冻结：某字体解析出垃圾尺寸(18.5e9×19.4e9)后 getImage 的像素填充循环
      //   是万亿级迭代 → 主线程冻结（字体下拉点不了）。入口拦截：垃圾尺寸 → 空白画布（该字体显示空预览）
      if (rel === 'web-apps/apps/documenteditor/main/app.js' || rel === 'web-apps/apps/spreadsheeteditor/main/app.js') {
        // ★ getImage 整体替换（2026-10-01 终版）：瓦片=字体名文本（11px 灰字），零 wasm/零像素循环
        //   —— 物理上不可能冻结。旧函数体由 void function 语法吞掉（保持逗号表达式合法）。
        //   恢复原渲染：去掉本 replace 即可。
        const js = fs.readFileSync(f, 'utf8').replace(
          'this.getImage=function(i,n,o){',
          'this.getImage=function(i,n,o){(window.__drawOrder=window.__drawOrder||[]).length<40&&window.__drawOrder.push(i);window.__ff=this;var nm=(window.__font_names&&window.__font_names[Number(i)])||(i==null?"":i);var cnM=window.__font_names_cn||{};var label=cnM[String(nm).toLowerCase()]||nm;n=n||document.createElement("canvas");n.width=520;n.height=44;n.style.width="100%";n.style.height="22px";try{var c=n.getContext("2d");c.clearRect(0,0,520,44);c.font=`28px "${nm}"`;c.fillStyle="#ffffff";c.fillText(label,10,32)}catch(e){}return n},void function(i,n,o){'
        )
          // ★ SSE 参数序不同：(i,o,n)——画布参数是 o。2026-10-02 实测 Excel 字体点击冻结即此失配
          .replace(
            'this.getImage=function(i,o,n){',
            'this.getImage=function(i,o,n){(window.__drawOrder=window.__drawOrder||[]).length<40&&window.__drawOrder.push(i);window.__ff=this;var nm=(window.__font_names&&window.__font_names[Number(i)])||(i==null?"":i);var cnM=window.__font_names_cn||{};var label=cnM[String(nm).toLowerCase()]||nm;o=o||document.createElement("canvas");o.width=520;o.height=44;o.style.width="100%";o.style.height="22px";try{var c=o.getContext("2d");c.clearRect(0,0,520,44);c.font=`28px "${nm}"`;c.fillStyle="#ffffff";c.fillText(label,10,32)}catch(e){}return o},void function(i,o,n){'
          )
          // ★ WPS 式排序：中文字体族排最前，其余按名称排序
          .replace(
            'i.store.set(t.toJSON()),',
            "i.store.set((function(a){var cn=/^(simsun|nsimsun|simhei|kaiti|fangsong|dengxian|microsoft yahei|msyh|fz|\u5b8b\u4f53|\u65b0\u5b8b\u4f53|\u9ed1\u4f53|\u6977\u4f53|\u6977\u4f53_gb2312|\u4eff\u5b8b|\u4eff\u5b8b_gb2312|\u7b49\u7ebf|\u5fae\u8f6f\u96c5\u9ed1|\u534e\u6587|\u65b9\u6b63|\u6c49\u4eea|\u5e7c\u5706|\u96b6\u4e66)/i;window.__storeItem0=a&&a[0];a.sort(function(x,y){var xn=cn.test(String(x.name||'')),yn=cn.test(String(y.name||''));if(xn!==yn)return xn?-1:1;return String(x.name||'').localeCompare(String(y.name||''))});window.__storeSorted0=a&&a[0];window.__storeCnCount=a.filter(function(it){return cn.test(String(it.name||''))}).length;return a})(t.toJSON())),"
          )
        send(res, 200, js, 'application/javascript')
        return
      }
      if (!fs.existsSync(f)) {
        if (rel.endsWith('plugins.json')) { send(res, 200, JSON.stringify({ pluginsData: [] }), 'application/json'); return }
        if (rel.endsWith('themes.json')) { send(res, 200, JSON.stringify({ themes: [] }), 'application/json'); return }
        send(res, 404, 'nf: ' + rel); return
      }
      // ★ 2026-10-10：xlsx/pptx 的 index.html 也走注入层（此前仅 word）——
      //   AscDesktopEditor 假壳闸门修复、断线 toast、拼写检查等注入对三类型一视同仁
      if (/(documenteditor|spreadsheeteditor|presentationeditor)\/main\/index\.html$/.test(f.replace(/\\/g, '/'))) { serveEditorIndex(res, f, req); return }
      // ★ 2026-10-07 压缩打开时长：大文件流式 + ETag 协商缓存（vendors 库内容稳定，二次打开 304 零下载）
      sendFile(res, req, 200, f, MIME[path.extname(f).toLowerCase()] || 'application/octet-stream', { negotiate: true }); return true /*★ 流式分支必须显式 true（2026-10-10）: 返回 undefined 会被裸调用方当未处理再补刀，响应体被截*/
      return
    }

    send(res, 404, 'nf ' + p)
    return true
  } catch (e) {
    send(res, 500, String(e).slice(0, 300))
    return true
  }
}


// ★ WS upgrade 反代隧道：/doc ws 直通 3798（协作/文档下发通道）
export function handleOoUpgrade(req, socket, head) {
  if (req.url && req.url.includes('/doc/')) {
    const target = http.request({
      host: '127.0.0.1', port: 3798, path: req.url.replace('/oo-assets', ''),
      headers: Object.assign({}, req.headers, { host: '127.0.0.1:3798', 'x-forwarded-host': req.headers['x-forwarded-host'] || req.headers.host || '127.0.0.1', 'x-forwarded-proto': 'http' }),
    })
    target.end()
    target.on('upgrade', (res2, upstreamSocket, upstreamHead) => {
      const lines = [`HTTP/1.1 101 Switching Protocols`]
      for (const [k, v] of Object.entries(res2.headers)) lines.push(`${k}: ${v}`)
      socket.write(lines.join('\r\n') + '\r\n\r\n')
      if (upstreamHead && upstreamHead.length) socket.write(upstreamHead)
      upstreamSocket.pipe(socket)
      socket.pipe(upstreamSocket)
      socket.on('error', () => { upstreamSocket.destroy() })
      upstreamSocket.on('error', () => { socket.destroy() })
    })
    target.on('response', (r2) => {
      const head = [`HTTP/1.1 ${r2.statusCode} ${r2.statusMessage || ''}`]
      for (const [k, v] of Object.entries(r2.headers)) head.push(`${k}: ${v}`)
      socket.write(head.join('\r\n') + '\r\n\r\n')
      socket.end()
    })
    target.on('error', () => socket.destroy())
  } else {
    socket.destroy()
  }
}

// hot-reload marker 2026-10-01

// hot-reload marker2

// hot-reload marker 2026-10-01 v2

// ═══ 转换器心跳守护（2026-10-08）═════════════════════════════════════
// 背景：DS 的 embedded converter runner 长时间运行后会**静默死亡**（消费循环
// 断链）：任务被接受（convertRequest start/end）但永不执行、零产物——用户侧
// 表现为所有 OO 文档永久 loading。DS 自身无自愈。
// 守护：每 90s 用 dataURL 喂一个探针转换任务（不依赖端口/host），25s 后查
// App_Data 是否出现该 key 的产物目录；连续 2 轮无产物 → 杀 DS 并重拉。
// 仅在「本会话发生过 OO 请求」后启用（__probeHost 有值），且首轮延迟 3 分钟。
var __probeHost = ''  // var（提升）：handleOoRequest 入口先于本文件尾部引用
let __probeB64 = ''
let __probeFails = 0
let __probeBusy = false
let __dsWasHealthy = false // ★ 2026-10-11 启动容忍：本代 DS 是否曾通过健康检查（从未健康=冷启动期，禁杀）

async function probeHeartbeat () {
  if (__probeBusy) return
  __probeBusy = true
  try {
    /* ★ 2026-10-10 判活重写（假心跳连环坑，无头浏览器实测三连实锤）：
     *   ① 旧版 async:true + 翻找 App_Data 产物 —— 无 callbackUrl 的任务产物不落盘 ⇒ 恒失败；
     *   ② data: URL → DS 直接 error -4（本构建不支持）；
     *   ③ HTTP URL 同步转换 → error -3（本内嵌 DS 的 /converter REST 面整体不可用，
     *      编辑器走 /doc 内部转换所以从未暴露）。
     *   ⇒ 旧心跳判定恒失败 ⇒ 每 ~3 分钟误杀 DS 重启、干掉全部编辑会话
     *   （用户「失去网络连接」弹窗高发、偶发永远 loading 的根源）。
     *   新判活 = 官方 /healthcheck（true 即活），与转换 REST 彻底解耦。 */
    const res = await fetch('http://127.0.0.1:3798/healthcheck', { signal: AbortSignal.timeout(8000) })
    const text = (await res.text().catch(() => '')).trim()
    if (res.ok && text === 'true') {
      __probeFails = 0
      __dsWasHealthy = true
    } else {
      __probeFails++
      console.log('[oo-core] 转换器心跳失败(' + __probeFails + '/2) http=' + res.status + ' body=' + text.slice(0, 40))
      if (__probeFails >= 2) {
        __probeFails = 0
        /* ★ 2026-10-11 启动容忍（12:51/13:01 死循环实锤，247 次启动日志）：DS 冷启动
         *   加载 112M 依赖需数分钟，期间 /healthcheck 必失败——从未健康过的 DS 纯属
         *   启动期，禁杀。否则构成「杀→重启→再杀」死循环（编辑器永远 loading）。 */
        if (!__dsWasHealthy) { console.log('[oo-core] DS 处于启动期（从未健康）——跳过重启，继续观察') } else {
        console.log('[oo-core] 判定 DocService 已死——自动重启 DS')
        try {
          /* ★ 2026-10-11 安全修正：不再按 pid 文件盲杀——Windows 会复用 PID，
           *   盲杀可能误杀无关进程（疑似「突然崩了」来源之一）。
           *   只精确杀当前监听 3798 的进程（下方 netstat 兜底已覆盖该场景）。 */
          // 兜底：杀所有监听 3798 的进程（多实例残骸）
          try {
            const { execSync } = await import('node:child_process')
            // ★ windowsHide: true——否则每次执行都弹可见控制台窗（用户实测「频繁打开黑窗」）
            const out = execSync('netstat -ano | findstr :3798 | findstr LISTENING', { windowsHide: true, stdio: ['ignore', 'pipe', 'ignore'] }).toString()
            const pids = new Set(out.split(/\r?\n/).map((l) => l.trim().split(/\s+/).pop()).filter(Boolean))
            for (const pid of pids) { try { process.kill(Number(pid)) } catch { /* ignore */ } }
          } catch { /* findstr 无结果=无人监听 */ }
        } catch { /* ignore */ }
        // 等 1.5s 让端口释放，再拉新 DS（oo-docservice 的 ensure 幂等：死→启动）
        await new Promise((r) => setTimeout(r, 1500))
        try {
          const m = await import('./oo-docservice.mjs')
          await m.ensureDocService({ quiet: false })
          console.log('[oo-core] DS 已自动重启，转换器恢复')
        } catch (e) { console.log('[oo-core] DS 自动重启失败:', String(e).slice(0, 100)) }
        }
      }
    }
  } catch { /* 网络抖动=本轮失败，连续两轮才重启 */ __probeFails++; if (__probeFails >= 2) { __probeFails = 1 } }
  finally { __probeBusy = false }
}

setInterval(() => {
  if (!__probeHost) return            // 本会话没人用过 OO → 不探
  probeHeartbeat().catch(() => {})
}, 90000)
setTimeout(() => {
  if (__probeHost) probeHeartbeat().catch(() => {})
}, 3 * 60 * 1000)   // 首轮 3 分钟后（给 DS 启动+首个文档留时间）
