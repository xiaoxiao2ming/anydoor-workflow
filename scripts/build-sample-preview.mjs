/**
 * 生成一份单文件预览页，内嵌 JSZip + docx-preview + 模板/产物两份 base64 docx。
 * 用无头 Edge 截图就能验证「预览即终稿」。
 *
 * 运行：node scripts/build-sample-preview.mjs
 */

import { readFileSync, writeFileSync, existsSync } from 'fs'
import { dirname, join } from 'path'
import { fileURLToPath } from 'node:url'

const HERE = dirname(fileURLToPath(import.meta.url))
const ROOT = join(HERE, '..')
const OUT = join(ROOT, 'sample-template')
const FILLED = join(OUT, '市政管道工程量报告-已填充.docx')

if (!existsSync(FILLED)) {
  throw new Error('先跑 scripts/bind-sample.mjs 生成产物')
}

const read = (p) => readFileSync(p, 'utf8')
const jszip = read(join(ROOT, 'node_modules/jszip/dist/jszip.min.js'))
const docxPrev = read(join(ROOT, 'node_modules/docx-preview/dist/docx-preview.min.js'))

const load = (name) => {
  const p = join(OUT, name)
  if (!existsSync(p)) return null
  return { name, b64: readFileSync(p).toString('base64') }
}

const files = [load('市政管道工程量报告-模板.docx'), load('市政管道工程量报告-已填充.docx')].filter(Boolean)

const html = `<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>Word 报告预览验证</title>
<style>
  body{margin:0;background:#0d1117;color:#e6edf3;font:14px/1.6 "Microsoft YaHei",sans-serif}
  .bar{position:sticky;top:0;z-index:9;display:flex;gap:8px;align-items:center;
       padding:10px 16px;background:#161b22;border-bottom:1px solid #30363d}
  .bar button{padding:4px 12px;font-size:13px;color:#c9d1d9;background:#21262d;
       border:1px solid #30363d;border-radius:5px;cursor:pointer}
  .bar button.on{background:#1f6feb;border-color:#1f6feb;color:#fff}
  .bar span{font-size:12px;color:#8b949e}
  #view{padding:20px;min-height:100vh}
  .docx-wrapper{background:#161b22;padding:12px}
  .docx-wrapper>section.docx{background:#fff;color:#000;margin:0 auto 16px;
       box-shadow:0 4px 16px rgba(0,0,0,.5)}
</style></head>
<body>
<div class="bar">
  ${files.map((f, i) => `<button data-i="${i}" class="${i === 1 ? 'on' : ''}">${f.name}</button>`).join('')}
  <span id="stat"></span>
</div>
<div id="view"></div>
<div id="sty"></div>
<script>${jszip}</script>
<script>${docxPrev}</script>
<script>
var FILES = ${JSON.stringify(files.map((f) => ({ name: f.name, b64: f.b64 })))}
function b64ToUint8(b64){var bin=atob(b64);var u=new Uint8Array(bin.length)
  for(var i=0;i<bin.length;i++)u[i]=bin.charCodeAt(i);return u}
function render(i){
  var view=document.getElementById('view'), sty=document.getElementById('sty')
  view.innerHTML=''; sty.innerHTML=''
  document.querySelectorAll('.bar button').forEach(function(b,bi){
    b.className = bi===i ? 'on' : '' })
  var t0=performance.now()
  docx.renderAsync(b64ToUint8(FILES[i].b64), view, sty, {
    inWrapper:true, breakPages:true, renderHeaders:true, renderFooters:true,
    ignoreWidth:false, useBase64URL:true, trimXmlDeclaration:true
  }).then(function(){
    var pages=view.querySelectorAll('section.docx').length
    document.getElementById('stat').textContent =
      FILES[i].name+' · '+pages+' 页 · '+Math.round(performance.now()-t0)+'ms'
  }).catch(function(e){
    document.getElementById('stat').textContent='渲染失败：'+e.message
    console.error(e)
  })
}
document.querySelectorAll('.bar button').forEach(function(b){
  b.onclick=function(){render(Number(b.dataset.i))}})
render(FILES.length-1)
</script>
</body></html>`

const out = join(OUT, 'preview.html')
writeFileSync(out, html, 'utf8')
console.log('已生成：', out)
console.log('大小：', (html.length / 1024).toFixed(0), 'KB')
console.log('内嵌文档：', files.map((f) => f.name).join(' / '))
