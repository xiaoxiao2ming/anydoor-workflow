/**
 * 从 ONLYOFFICE Desktop Editors 安装目录抽取离线编辑器资产
 *
 * 前置：安装 ONLYOFFICE 桌面版（v9.4+，github release 或官网下载）
 *   https://github.com/ONLYOFFICE/DesktopEditors/releases/latest/download/DesktopEditors_x64.exe
 *
 * 用法：
 *   node scripts/extract-onlyoffice-assets.mjs "C:\Program Files\ONLYOFFICE\DesktopEditors"
 *
 * 抽取目标 vendors/onlyoffice/：
 *   editors/            web-apps（编辑器 UI）+ sdkjs（渲染/编辑内核）+ core（x2t 转换 wasm）+ fonts + allfontgen
 *   体积约 400~600MB —— 内嵌打包时再按 word 场景瘦身（web-apps/apps/documenteditor + sdkjs/word + core）
 */
import fs from 'node:fs'
import path from 'node:path'

const src = process.argv[2]
if (!src || !fs.existsSync(src)) {
  console.error('用法: node scripts/extract-onlyoffice-assets.mjs "<桌面版安装目录>"')
  console.error('示例: node scripts/extract-onlyoffice-assets.mjs "C:\\Program Files\\ONLYOFFICE\\DesktopEditors"')
  process.exit(1)
}

const editorsSrc = path.join(src, 'editors')
if (!fs.existsSync(editorsSrc)) {
  console.error(`未找到 ${editorsSrc} —— 请确认安装目录里含 editors/ 子目录`)
  process.exit(1)
}

const dst = path.join(process.cwd(), 'vendors', 'onlyoffice')
fs.mkdirSync(dst, { recursive: true })

/** cp -r（跟随目录递归） */
function cp(from, to) {
  console.log(`  copy ${path.basename(from)} -> ${path.relative(process.cwd(), to)}`)
  fs.cpSync(from, to, { recursive: true })
}

// 桌面版 editors/ 下的关键资产（存在才拷）
const want = ['web-apps', 'sdkjs', 'sdkjs-word', 'core', 'fonts', 'dictionaries', 'file-viewer']
for (const name of want) {
  const p = path.join(editorsSrc, name)
  if (fs.existsSync(p)) cp(p, path.join(dst, name))
}
// 桌面版根下的本地服务/内核（实测 v9.4：converter/x2t.exe、fonts/、dictionaries/）
for (const name of ['converter', 'docservice', 'ServerComponents', 'fonts', 'dictionaries', 'ascdocumentscore.dll']) {
  const p = path.join(src, name)
  if (fs.existsSync(p)) cp(p, path.join(dst, name))
}

const mb = (p) => {
  let n = 0
  const walk = (d) => fs.readdirSync(d, { withFileTypes: true }).forEach((e) => {
    const f = path.join(d, e.name)
    e.isDirectory() ? walk(f) : (n += fs.statSync(f).size)
  })
  walk(p)
  return Math.round(n / 1048576)
}
console.log(`\n完成。vendors/onlyoffice/ 体积 ≈ ${mb(dst)} MB`)
console.log('下一步：node scripts/onlyoffice-poc.mjs （本地起服务加载编辑器打开 docx）')
