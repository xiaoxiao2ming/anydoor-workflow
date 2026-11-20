#!/usr/bin/env node
/**
 * check-cside-sync —— C 端（Electron 生产包）与网页端代码同步门禁
 *
 * ★ 为什么需要这个（2026-10-06 实锤事故）：
 *   C 端 `npm run start` 跑的是 **dist/**（生产构建产物），
 *   网页端 `npm run dev:web` 跑的是 **src/**（源码直出，HMR 即时）。
 *   两者**没有任何机制保证同步**——src 改完，dist 不会自动重建。
 *
 *   实测后果：用户在网页端用得好好的（src 已有 AI 节点 Agent 配置五组面板等
 *   3 个提交的新功能），打开 C 端却"很多内容没有"——
 *   因为 dist 还停在 3 个提交之前的旧包。
 *   而开发期所有检查（tsc / vitest / 验收探针）都跑在 src 上，**全绿**，
 *   没有任何一处会报出这个差异。
 *
 * 检查项（任一不通过 → 退出码 1）
 *   1. dist 是否存在（从未构建过）
 *   2. dist/index.html 的 mtime 是否**晚于**最近一次提交（落后即报红）
 *   3. dist 主 chunk 是否与 index.html 引用的一致（产物没被手工删/换）
 *   4. 抽样若干 src 里的 UI 文案（不会被压缩改名）是否都在主 chunk 里
 *      —— 防「dist 存在但是旧的/不完整的」这种半新不新的状态
 *
 * 用法
 *   node scripts/check-cside-sync.mjs            # 门禁
 *   node scripts/check-cside-sync.mjs --stat     # 只看状态，不判定
 *   node scripts/check-cside-sync.mjs --rebuild  # 落后时自动重建再复检
 */
import { readFileSync, existsSync, statSync, readdirSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { execSync } from 'node:child_process'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const DIST = join(ROOT, 'dist')
const INDEX_HTML = join(DIST, 'index.html')
const args = process.argv.slice(2)
const STAT_ONLY = args.includes('--stat')
const REBUILD = args.includes('--rebuild')

const say = (...a) => console.log(...a)
/** 判定结果统一走这里收集，末尾统一决定退出码——
 *  ★ 绝不能边判边 console.log 后靠return 值串联：
 *    那样"红"与"绿"混在输出里，结论可能与实际红项自相矛盾
 *    （本脚本第一版就犯过：报红却说"可出包"，退出码还错乱）。*/
const FAILS = []
const bad = (msg) => { FAILS.push(msg); say('  [红] ' + msg) }
const good = (msg) => { say('  [绿] ' + msg) }

/** ★ 进入产物的源文件：扩展名白名单 + 排除测试/规格。
 *  **单一真相源** —— 时间判据与内容判据必须用同一份，
 *  两处各写一遍必然漂移（实测漂移过一次：walk 排 test+spec、collectText 只排 test）。
 *  ★ 必须定义在使用点**之前**：const 有 TDZ，放在后面会 ReferenceError。 */
const SRC_FILE_RE = /.(ts|tsx|css|html)$/
const SRC_SKIP_RE = /.(test|spec).tsx?$/
const isSrcFile = (name) => SRC_FILE_RE.test(name) && !SRC_SKIP_RE.test(name)

/* ── 1. dist 是否存在 ── */
say('=== C 端同步检查（dist  vs  src）===')
if (!existsSync(INDEX_HTML)) {
  bad('dist/index.html 不存在 —— C 端从未构建过。')
  say('')
  say('结论：dist 缺失，修复：npm run build')
  if (REBUILD) {
    try { execSync('npm run build', { cwd: ROOT, stdio: 'inherit' }); say('>>> 重建完成，请重跑本脚本复检。') }
    catch (e) { say('>>> 重建失败：' + e.message) }
  }
  process.exit(1)
}

/* ── 2. dist 相对最近提交是否落后 ─-
 * 取max(最近提交时间, 最新源码 mtime)：dist 只要比**源码**新即可。
 *
 * ⚠ 环境坑（本机已实测，必须保留这段降级）：
 *   Node 内 spawnSync 任何可执行文件都返回 **EBUSY**（沙箱限制），
 *   `execSync('git ...')` 会抛 status=null / "spawnSync cmd.exe EBUSY"
 *   ⇒ 取不到提交时间。直读 .git/HEAD 拿 sha 拿不到时间，
 *   故**直接以源码 mtime 为准**（这已经足够：源码比 dist 新才是真问题）。
 *   源码 mtime 不可用时（如只读挂载）才退回 git 通道。
 */
let lastCommitMs = 0
try {
  const out = execSync('git log -1 --format=%ct', { cwd: ROOT, encoding: 'utf8' }).trim()
  lastCommitMs = Number(out) * 1000
} catch {
  say('  [注] git 通道不可用（EBUSY，见脚本头注释），改以源码 mtime 为准')
}

let newestSrcMs = 0
let newestSrcFile = ''
function walk (dir) {
  for (const name of readdirSync(dir)) {
    const p = join(dir, name)
    let st
    try { st = statSync(p) } catch { continue }
    if (st.isDirectory()) walk(p)
    // ★ 只看真正进产物的源文件（见上方 isSrcFile单一真相源）。
    //   测试文件不进产物：改它不该被判成「dist 落后」
    //   （实测只动了 src/xlsx/refPins.test.ts 就报红并拒绝打包）。
    else if (isSrcFile(name)) {
      if (st.mtimeMs > newestSrcMs) { newestSrcMs = st.mtimeMs; newestSrcFile = p }
    }
  }
}
for (const d of ['src', 'electron']) {
  const p = join(ROOT, d)
  if (existsSync(p)) walk(p)
}
const baselineMs = Math.max(lastCommitMs, newestSrcMs)

const distMs = statSync(INDEX_HTML).mtimeMs
const fmt = (ms) => new Date(ms).toLocaleString('zh-CN', { hour12: false })
say(`  最近提交：${lastCommitMs ? fmt(lastCommitMs) : '(取不到，见上方注)'}`)
say(`  最新源码：${newestSrcMs ? fmt(newestSrcMs) + '  ' + newestSrcFile.replace(ROOT + '\\', '') : '(无)'}`)
say(`  dist 构建：${fmt(distMs)}`)

/* ★ 容差说明（重要，别当成 bug 修掉）：
 *   mtime 判据是「启发式」，不是事实。真实事故里它能抓到的前提是
 *   **落后了肉眼可见的时间**（实测那次落后约 1.5 小时 / 3 个提交）。
 *   但 mtime 有两类必然误报：
 *     ① 改完源码立刻检查（`touch`、编辑器保存）⇒ 差几秒，内容其实一致
 *     ② 构建过程中src 被读（部分工具会更新 atime）⇒ 噪声
 *   所以给 60 秒容差：**秒级差不算落后**，避免"改一行就报红"被用户学会忽略。
 *   ★ 真正兜底的是下面第 4 项的**内容校验**（不看时间，只看文案在不在产物里），
 *   它不受 mtime 噪声影响。两层是互补的，不要因为容差就以为时间判据没用了。
 */
const TOLERANCE_MS = 60_000
if (distMs < baselineMs - TOLERANCE_MS) {
  const gapMs = baselineMs - distMs
  const gapText = gapMs < 3_600_000
    ? `${Math.round(gapMs / 60000)} 分钟`
    : `${(gapMs / 3_600_000).toFixed(1)} 小时`
  bad(`dist 比源码旧 ${gapText} ⇒ 用户打开 C 端看到的是旧功能（网页端却是新的）`)
} else {
  good(`dist 不落后于源码（按时间，容差 60 秒）`)
}

/* ── 3. 主 chunk 引用一致性 ── */
let chunkName = ''
try {
  const html = readFileSync(INDEX_HTML, 'utf8')
  const m = /assets\/(index-[A-Za-z0-9_-]+\.js)/.exec(html)
  chunkName = m ? m[1] : ''
} catch { /* 已在第1 项处理 */ }
if (!chunkName) {
  bad('index.html 里找不到主 chunk 引用')
} else if (!existsSync(join(DIST, 'assets', chunkName))) {
  bad(`index.html 引用的 ${chunkName} 在 dist/assets 里不存在`)
} else {
  good('主 chunk 存在：' + chunkName)
}

/* ── 4. 抽样文案是否在 chunk 内 ──
 * ★ 只用**界面文案**做抽样，不能用函数名/标识符：
 *   terser 会重命名局部符号（实测 grep 'aiBuildMessages' 在产物里恒为 0），
 *   而用户可见文案不会被改写。
 * ★ 判据必须是「**src 全目录**里有、dist 里没有」：
 *   早期版本只 grep PropertiesPanel.tsx，于是 `自动对齐`（实际在 CanvasRibbon.tsx）
 *   被算进"src 没有"而侥幸通过；若换个位置放个真缺失的文案就会漏判。
 *   反向（dist 有而 src 没有）不算红——构建产物里可能有源码已删的残留文案。
 */
const SAMPLES = ['自动对齐', '自动排列', '输入变量', '提示词', '技能', '规则跟随']

let srcText = ''
function collectText (dir) {
  for (const name of readdirSync(dir)) {
    const p = join(dir, name)
    let st
    try { st = statSync(p) } catch { continue }
    if (st.isDirectory()) collectText(p)
    // ★ 排除口径必须与上面 walk() 的**同一份正则**，两处各写一遍必然漂移
    //   （本轮就漂移过：walk 排test+spec、collectText 只排 test）。
    else if (/\.(ts|tsx|css|html)$/.test(name) && !/\.(test|spec)\.tsx?$/.test(name)) {
      try { srcText += readFileSync(p, 'utf8') } catch { /* ignore */ }
    }
  }
}
collectText(join(ROOT, 'src'))

if (chunkName) {
  const chunkPath = join(DIST, 'assets', chunkName)
  try {
    const buf = readFileSync(chunkPath, 'utf8')
    const missing = SAMPLES.filter((s) => srcText.includes(s) && !buf.includes(s))
    const checked = SAMPLES.filter((s) => srcText.includes(s))
    if (!checked.length) {
      bad('抽样词在 src 里一个都没找到 —— 清单已漂移，本项失效（需更新 SAMPLES）')
    } else if (missing.length) {
      bad('src 里有、dist 里没有的文案：' + missing.join('、') +
        ' ⇒ dist 内容不完整（时间可能是新的，但内容是旧的）')
    } else {
      good(`抽样文案齐全（src∩产物 ${checked.length}/${SAMPLES.length}：${checked.join('、')}）`)
    }
  } catch (e) {
    bad('读取主 chunk 失败：' + e.message)
  }
}

/* ── 统一结论与退出码（唯一出口，不再边判边退）── */
say('')
if (STAT_ONLY) {
  say(`（--stat）${FAILS.length} 项不通过：${FAILS.join('；') || '无'}`)
  process.exit(0)
}
if (FAILS.length === 0) {
  say('结论：dist 与源码同步，C 端可出包。')
  process.exit(0)
}
say(`结论：★ ${FAILS.length} 项不通过 —— 这类差异会让用户「网页端有、C 端没有」。`)
for (const f of FAILS) say('     · ' + f)
say('     修复：npm run build（或 node scripts/check-cside-sync.mjs --rebuild）')
if (REBUILD) {
  say('')
  say('>>> --rebuild：正在重新构建…')
  try {
    execSync('npm run build', { cwd: ROOT, stdio: 'inherit' })
    say('>>> 重建完成，请重跑本脚本复检。')
  } catch (e) {
    say('>>> 重建失败：' + e.message)
  }
}
process.exit(1)