/**
 * x2t/DS 转换环境一键配置（2026-10-08 创建，2026-10-09 扩充）
 *
 * 做两件事（均幂等，打包/同步 vendors 后跑一次即可）：
 *
 * 1. x2t 字体环境
 *    x2t 转换 docx→Editor.bin 需要字体（AllFonts.js + fonts 目录），
 *    缺失时 x2t 会挂起不退出（无报错）；xlsx 不需要字体不受影响。
 *
 * 2. DS embedded converter 的 lcid ESM 修复（2026-10-09 根因）
 *    症状：文档永远 loading，DS 日志每次启动必报——
 *      embeddedConverter.start error: ERR_REQUIRE_ESM:
 *      require() of ES Module ...\node_modules\lcid\index.js
 *      from ...\FileConverter\sources\converter.js not supported
 *    根因：vendors 里带进来的 lcid@5.x 是 ESM-only 包（"type":"module"），
 *    而 OO 官方 converter.js 是 CJS，第 45 行 require('lcid') 直接崩——
 *    embedded converter 从进程启动那一刻就是死的，转换任务只进队列没人消费。
 *    修复：降级到 CJS 版 lcid@3.1.1 + invert-kv@3.0.0（converter.js 仅用
 *    lcid.from(number)，v3 API 完全兼容，不改 OO 官方源码）。
 *
 * 用法：node scripts/setup-x2t-fonts.mjs
 */
import { copyFileSync, cpSync, existsSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { execSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const appRoot = join(root, 'out2', 'client-base', 'resources', 'app')
const bin = join(appRoot, 'vendors', 'oo-server', 'FileConverter', 'Bin')
const fonts = join(appRoot, 'vendors', 'onlyoffice', 'fonts')

// ---------- 1. x2t 字体 ----------
if (!existsSync(bin)) { console.error('!! 未找到 Bin 目录: ' + bin); process.exit(1) }
if (!existsSync(fonts)) { console.error('!! 未找到字体目录: ' + fonts); process.exit(1) }
copyFileSync(join(fonts, 'AllFonts.js'), join(bin, 'AllFonts.js'))
mkdirSync(join(bin, 'fonts'), { recursive: true })
cpSync(fonts, join(bin, 'fonts'), { recursive: true })
console.log('✓ x2t 字体环境就绪: ' + join(bin, 'fonts'))

// ---------- 2. lcid ESM → CJS 降级 ----------
const CJS_VERSIONS = { lcid: '3.1.1', 'invert-kv': '3.0.0' }
const nm = join(appRoot, 'vendors', 'oo-server', 'FileConverter', 'node_modules')

function pkgType (name) {
  const pj = join(nm, name, 'package.json')
  if (!existsSync(pj)) return 'missing'
  try {
    const t = JSON.parse(readFileSync(pj, 'utf8'))
    return t.type === 'module' ? 'esm' : 'cjs'
  } catch { return 'unknown' }
}

const bad = Object.keys(CJS_VERSIONS).filter((n) => pkgType(n) === 'esm')
if (bad.length === 0) {
  console.log('✓ lcid/invert-kv 已是 CJS 版，跳过')
} else {
  const tmp = join(tmpdir(), 'cb-lcid-fix-' + Date.now())
  mkdirSync(tmp, { recursive: true })
  const specs = bad.map((n) => n + '@' + CJS_VERSIONS[n]).join(' ')
  console.log('… 正在获取 CJS 版: ' + specs)
  execSync('npm pack --silent ' + specs, { cwd: tmp, stdio: 'ignore' })
  const dirs = {}
  for (const name of bad) {
    const tgz = execSync('ls ' + join(tmp, name + '-' + CJS_VERSIONS[name] + '.tgz')).toString().trim()
    const dst = join(tmp, 'x-' + name)
    execSync('tar -xzf "' + tgz + '" -C "' + tmp + '" && mv "' + join(tmp, 'package') + '" "' + dst + '"')
    dirs[name] = dst
  }
  for (const name of bad) {
    const cur = join(nm, name)
    if (existsSync(cur) && !existsSync(cur + '.esm.bak')) renameSync(cur, cur + '.esm.bak')
    else if (existsSync(cur)) rmSync(cur, { recursive: true, force: true })
    cpSync(dirs[name], cur, { recursive: true })
    console.log('✓ ' + name + ' 降级为 CJS@' + CJS_VERSIONS[name] + '（原 ESM 版备份为 ' + name + '.esm.bak）')
  }
  rmSync(tmp, { recursive: true, force: true })
  console.log('★ lcid 修复完成 —— 需重启 DocService 生效（重启 ClientBase 或等 90s 心跳自愈）')
}

// 修复结果自检
for (const name of Object.keys(CJS_VERSIONS)) {
  const t = pkgType(name)
  if (t !== 'cjs') { console.error('!! ' + name + ' 状态异常: ' + t); process.exitCode = 1 }
}
