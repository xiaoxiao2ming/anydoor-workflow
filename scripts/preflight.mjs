#!/usr/bin/env node
/**
 * preflight —— 并行开发开工门禁
 *
 * 场景：同一份代码，多个开发方向（worktree / 分支）同时推进。
 * 每次任务开始前必须先跑本脚本，比对基线漂移与文件重叠，冲突即退出码 1。
 *
 * 用法
 *   node scripts/preflight.mjs# 门禁（漂移 / 重叠 → 退出码 1）
 *   node scripts/preflight.mjs --stat                        只看状态，不判定
 *   node scripts/preflight.mjs --claim "画布=<files>"        登记本分支将改的文件
 *   node scripts/preflight.mjs --claim-auto "画布"           登记=本分支实际改动面
 *   node scripts/preflight.mjs --release                     清除本分支登记
 *
 * ⚠ 环境坑（本脚本必须容忍，不要删这段降级逻辑）
 *   本机Node 进程内 spawn **任何** 可执行文件都返回 **EBUSY**（沙箱限制，非 git 特有），
 *   而 Bash 里调git 正常。故本脚本做**双通道**：
 *     通道 A  spawn git —— 正常环境走这条，信息最全
 *     通道 B  直读 .git 元数据 —— A 失败时自动降级，只读不写
 *
 *   通道 B 能拿到：分支名、各分支 HEAD 的 sha、登记文件（=门禁核心）
 *   通道 B 拿不到：提交图遍历（需解 packfile，本仓 113MB 全在pack 里，代价过高）、
 *                脏文件检测（需读 index）、worktree 列表
 *   → 故降级时**基线漂移只提示人工核对，不报红**。
 *     理由：算不准就报红 = 假警报，比没警报更坏（人会学会忽略它，然后真警报一起被忽略）。
 *
 * 设计约束
 *   1. 纯本地仓库（无 remote），基线以 master 为准，不做 fetch。
 *   2. 登记落 .workbuddy/claims/ —— 元数据目录，不参与构建与测试。
 *      ⚠ 任何「扫描器自身产物」必须在扫描范围外，否则自指死循环（同 oo-scan 基线教训）。
 *   3. 除 --claim/--claim-auto/--release 写登记外，本脚本不碰工作区。
 */

import { execFileSync } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync, rmSync, lstatSync } from 'node:fs'
import { join, dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const CLAIM_DIR = join(ROOT, '.workbuddy', 'claims')
const BASE_BRANCH = 'master'

// ============================================================ 通道定义

let SPAWN_OK = true
const DEGRADED = []

/** 通道 A：spawn git */
function git(args) {
  if (!SPAWN_OK) throw new Error('git 通道已禁用')
  return execFileSync('git', args, { cwd: ROOT, encoding: 'utf8', maxBuffer: 16 * 1024 * 1024 }).trim()
}

/** 通道 A 的安全包装：失败返回 null，不抛 */
function gitTry(args) {
  try {
    return git(args)
  } catch (e) {
    if (e && (e.code === 'EBUSY' || e.errno === -4082)) {
      SPAWN_OK = false
      DEGRADED.push('Node 进程内 spawn 任何可执行文件都返回 EBUSY（实测 node/cmd/where 同样失败，非 git 特有）→ 已降级为 .git 元数据解析')
    }
    return null
  }
}

// ============================================================ 通道 B：直读 .git

/** 解析 gitdir（worktree 下 .git 是文件，指向主库 worktrees/<name>） */
function resolveGitDir(root = ROOT) {
  const p = join(root, '.git')
  if (!existsSync(p)) return null
  if (lstatSync(p).isDirectory()) return p
  const txt = readFileSync(p, 'utf8').trim()
  const m = txt.match(/^gitdir:\s*(.+)$/)
  return m ? resolve(root, m[1]) : null
}

/** 主库 gitdir（worktree 的 gitdir 里commondir 指向主库） */
function mainGitDir(gd) {
  const cd = join(gd, 'commondir')
  if (existsSync(cd)) return resolve(gd, readFileSync(cd, 'utf8').trim())
  return gd
}

/** 读某分支的 ref → commit（沿 reflogs 找最近一次） */
function refToCommit(gitdir, branch) {
  const loose = join(gitdir, 'refs', 'heads', ...branch.split('/'))
  if (existsSync(loose)) return readFileSync(loose, 'utf8').trim()
  const packed = join(gitdir, 'packed-refs')
  if (existsSync(packed)) {
    for (const l of readFileSync(packed, 'utf8').split('\n')) {
      const m = l.match(/^([0-9a-f]{40})\s+refs\/heads\/(.+)$/)
      if (m && m[2] === branch) return m[1]
    }
  }
  return null
}

/** 列出所有本地分支 */
function listBranchesMeta(gitdir) {
  const out = new Set()
  const heads = join(gitdir, 'refs', 'heads')
  const walk = (d, prefix = '') => {
    if (!existsSync(d)) return
    for (const e of readdirSync(d, { withFileTypes: true })) {
      const name = prefix ? `${prefix}/${e.name}` : e.name
      if (e.isDirectory()) walk(join(d, e.name), name)
      else out.add(name)
    }
  }
  walk(heads)
  const packed = join(gitdir, 'packed-refs')
  if (existsSync(packed)) {
    for (const l of readFileSync(packed, 'utf8').split('\n')) {
      const m = l.match(/^([0-9a-f]{40})\s+refs\/heads\/(.+)$/)
      if (m) out.add(m[2])
    }
  }
  return [...out]
}

/**
 * 通道 B 的能力：读当前分支名与 HEAD sha。
 * worktree 的 HEAD 在各自 gitdir 下，commondir 用于找分支 ref。
 */
function metaCurrent(gitdir) {
  const headTxt = readFileSync(join(gitdir, 'HEAD'), 'utf8').trim()
  const m = headTxt.match(/^ref:\s*refs\/heads\/(.+)$/)
  if (m) return { branch: m[1], sha: refToCommit(mainGitDir(gitdir), m[1]) }
  return { branch: headTxt.slice(0, 7), sha: headTxt } // detached
}

/**
 * 通道 B 的提交图遍历 —— **有意不做**。
 *
 * 本仓 .git/objects 只有 3 个松散对象，113MB 全在 pack-*.pack 里。
 * 要读commit 对象得自己实现 pack .idx 二分查找 + zlib 解压 + delta 重建，
 * 约150 行且每次BFS 都要走一遍。为了一个「降级路径下的辅助信息」付这个代价，
 * 且写错了更难发现——故明确放弃：
 *   降级时基线漂移改为提示人工核对，**不报红**（假警报比没警报更坏）。
 *   门禁真正的核心——改动面重叠检测——只读 .workbuddy/claims/，与git 无关，照常生效。
 *
 * ⚠ 若将来仓库转为松散对象（object 散开存），可在此恢复 readCommit + metaAheadBehind。
 *   判断方法：ls .git/objects/<前两位>/ | head
 */
function metaAheadBehind() {
  return { behind: -1, ahead: -1, unknown: true }
}

// ============================================================ 统一取值层
// 有git 通道优先用git；否则用元数据。调用方只认这层。

const GITDIR = resolveGitDir(ROOT)
let branchCache = null

function meBranch() {
  if (branchCache) return branchCache
  const viaGit = gitTry(['rev-parse', '--abbrev-ref', 'HEAD'])
  branchCache = viaGit || (GITDIR ? metaCurrent(GITDIR).branch : '?')
  return branchCache
}

function headShort(branch) {
  const viaGit = gitTry(['log', '-1', '--format=%h %s', branch])
  if (viaGit) return viaGit
  // 降级：拿不到 commit 对象（内容在 packfile 里），但 ref 本身能读，
  // 所以至少有 sha 可展示——比一个孤零零的 "?" 有用得多。
  if (!GITDIR) return '?'
  const c = refToCommit(mainGitDir(GITDIR), branch)
  return c ? c.slice(0, 7) : '?'
}

function allBranches() {
  const viaGit = gitTry(['for-each-ref', '--format=%(refname:short)', 'refs/heads/'])
  if (viaGit) return viaGit.split('\n').filter(Boolean)
  return GITDIR ? listBranchesMeta(mainGitDir(GITDIR)) : []
}

function aheadBehind(branch) {
  const viaGit = gitTry(['rev-list', '--left-right', '--count', `${BASE_BRANCH}...${branch}`])
  if (viaGit) {
    const [behind, ahead] = viaGit.split(/\s+/).map(Number)
    return { behind: behind || 0, ahead: ahead || 0, unknown: false }
  }
  return GITDIR ? metaAheadBehind() : { behind: 0, ahead: 0, unknown: true }
}

/** 已提交差异文件集（仅通道 A可得） */
function committedDiffFiles(branch) {
  const out = gitTry(['diff', '--name-only', `${BASE_BRANCH}...${branch}`])
  if (!out) return null
  return new Set(out.split('\n').filter(Boolean))
}

/** 脏文件集（仅通道 A 可得） */
function dirtyFiles() {
  const out = gitTry(['status', '--porcelain'])
  if (!out) return null
  const set = new Set()
  for (const l of out.split('\n').filter(Boolean)) {
    let p = l.slice(3).trim()
    if (p.includes('->')) p = p.split('->').pop().trim()
    if (p) set.add(p)
  }
  return set
}

function worktreeList() {
  const out = gitTry(['worktree', 'list'])
  return out ? out.split('\n').filter(Boolean) : null
}

// ============================================================ 登记

function claimFile(branch) {
  return join(CLAIM_DIR, `${branch.replace(/[\\\\/]/g, '_')}.json`)
}

function readClaims() {
  if (!existsSync(CLAIM_DIR)) return {}
  const out = {}
  for (const f of readdirSync(CLAIM_DIR)) {
    if (!f.endsWith('.json')) continue
    try {
      const c = JSON.parse(readFileSync(join(CLAIM_DIR, f), 'utf8'))
      // ⚠ key 必须用 JSON 里的 branch 字段，**不能**用文件名。
      //   文件名把 `/` 换成了 `_`（feat/x → feat_x），拿它跟真实分支名比，
      //   带斜杠的分支**永远认不出自己的登记**——表现为「明明刚claim 过，
      //   门禁却说未登记」，且会把自己当成别人的方向去比对。
      //   单分支时看不出来，两个分支时必然暴露。branch 字段是唯一真相。
      const key = c.branch || f.slice(0, -5)
      out[key] = c
    } catch {
      /* 登记损坏按空处理：门禁不该被元数据问题卡死 */
    }
  }
  return out
}

function writeClaim(branch, topic, files) {
  mkdirSync(CLAIM_DIR, { recursive: true })
  writeFileSync(claimFile(branch), JSON.stringify({ topic, files, at: new Date().toISOString(), branch }, null, 2))
}

// ============================================================ 输出

const C = { r: '\x1b[31m', g: '\x1b[32m', y: '\x1b[33m', b: '\x1b[36m', d: '\x1b[2m', x: '\x1b[0m' }
const problems = []
const warns = []
const fail = (m) => problems.push(m)
const warn = (m) => warns.push(m)
const line = (s = '') => console.log(s)

function banner(text, sub) {
  line(`${C.b}▎${text}${C.x}${sub ?`  ${C.d}${sub}${C.x}` : ''}`)
  line()
}

function degradeNotice() {
  if (!DEGRADED.length) return
  line(`  ${C.y}⚠ 降级运行（Node 内 spawn 不了任何子进程，见文件头说明）${C.x}`)
  DEGRADED.forEach((d) => line(`    ${C.d}${d}${C.x}`))
  line(`    ${C.d}仍可用：分支名、各分支 HEAD sha、登记文件（= 门禁核心，比对照样生效）${C.x}`)
  line(`    ${C.d}不可用：提交图遍历（ahead/behind 算不出）、脏文件检测、worktree 列表${C.x}`)
  line()
}

// ============================================================ 参数

const args = process.argv.slice(2)
const has = (f) => args.includes(f)
const val = (f) => {
  const i = args.indexOf(f)
  return i >= 0 ? args[i + 1] : null
}

const me = meBranch()

// ---------------------------------------------------------------- 登记
if (has('--claim') || has('--claim-auto')) {
  const auto = has('--claim-auto')
  const raw = (val('--claim') || val('--claim-auto') || '').trim()
  if (!raw) {
    console.error(`${C.r}✘ 需要参数${C.x}\n  --claim "<topic>=<f1,f2,...>"  显式登记\n  --claim-auto "<topic>"    登记=本分支实际改动面`)
    process.exit(2)
  }
  const eq = raw.indexOf('=')
  const topic = (eq >= 0 ? raw.slice(0, eq) : raw).trim() || me
  let files
  if (auto) {
    const union = new Set()
    const c = committedDiffFiles(me)
    if (c) c.forEach((f) => union.add(f))
    const d = dirtyFiles()
    if (d) d.forEach((f) => union.add(f))
    if (!c && !d) {
      console.error(`${C.r}✘ --claim-auto 需要 git 通道（当前降级中，算不出实际改动面）${C.x}\n  请改用 --claim "<topic>=<files>" 显式登记`)
      process.exit(2)
    }
    files = [...union]
  } else {
    files = (eq >= 0 ? raw.slice(eq + 1) : '').split(',').map((s) => s.trim()).filter(Boolean)
  }
  if (!files.length) {
    console.error(`${C.r}✘ 没有可登记的文件${C.x}`)
    process.exit(2)
  }
  writeClaim(me, topic, files)
  line(`${C.g}✔ 已登记 ${files.length} 个文件${C.x}${C.d}分支 ${me} · 主题 ${topic}${C.x}`)
  files.forEach((f) => line(`  ${C.d}${f}${C.x}`))
  line()
  line(`${C.d}  他人分支跑 npm run preflight 即会检测到与这份登记的重叠。${C.x}`)
  process.exit(0)
}

if (has('--release')) {
  const f = claimFile(me)
  if (existsSync(f)) {
    rmSync(f)
    line(`${C.g}✔ 已清除分支 ${me} 的登记${C.x}`)
  } else {
    line(`${C.d}分支 ${me} 无登记${C.x}`)
  }
  process.exit(0)
}

// ---------------------------------------------------------------- 只看状态
if (has('--stat')) {
  banner('并行开发状态', `基准 ${BASE_BRANCH} · 当前 ${me}`)
  degradeNotice()
  line(`  ${C.d}基准${C.x} ${BASE_BRANCH.padEnd(0)}  ${headShort(BASE_BRANCH)}`)
  line()
  const others = allBranches().filter((b) => b !== BASE_BRANCH)
  if (!others.length) line(`  ${C.d}无其它分支${C.x}`)
  for (const b of others) {
    const ab = aheadBehind(b)
    const tag = ab.unknown
      ? ` ${C.y}?需手动比对${C.x}`
      : ab.behind > 0
        ? ` ${C.y}↓落后${ab.behind}${C.x}`
        : ` ${C.g}↑领先${ab.ahead}${C.x}`
    const mark = b === me ? `${C.g}▶${C.x}` : ' '
    line(`  ${mark} ${C.b}${b.padEnd(32)}${C.x}${headShort(b)}${tag}`)
  }
  line()
  const claims = readClaims()
  const entries = Object.entries(claims)
  if (entries.length) {
    line(`  ${C.d}已登记改动面：${C.x}`)
    for (const [b, c] of entries) {
      line(`  ${C.b}${b}${C.x} ${C.d}(${c.topic}) ${c.files.length} 个文件 · ${String(c.at).slice(0, 16).replace('T', ' ')}${C.x}`)
      c.files.forEach((f) => line(`      ${C.d}${f}${C.x}`))
      line()
    }
  } else {
    line(`  ${C.d}暂无登记。各分支开工前跑 npm run pre:claim登记改动面，重叠才能被发现。${C.x}`)
    line()
  }
  const wt = worktreeList()
  if (wt && wt.length > 1) {
    line(`  ${C.d}worktree：${C.x}`)
    wt.forEach((w) => {
      const m = w.match(/^(.+?)\s+(\S+)\s+\[(.+?)\]/)
      if (m) line(`    ${m[3] === me ? `${C.g}▶${C.x}` : ' '} ${C.d}${m[1]}${C.x} ${C.d}[${m[3]}]${C.x}`)
    })
    line()
  }
  process.exit(0)
}

// ================================================================ 门禁
banner('preflight 门禁', `当前分支 ${me}`)
degradeNotice()

// 1) 基线漂移 ---------------------------------------------------------
line(`  ${C.d}基准${C.x} ${headShort(BASE_BRANCH)}`)
if (me === BASE_BRANCH) {
  line(`  ${C.d}当前${C.x} 在基准分支上`)
} else {
  line(`  ${C.d}当前${C.x} ${headShort(me)}`)
  const ab = aheadBehind(me)
  if (ab.unknown) {
    warn(`无法自动计算与 ${BASE_BRANCH} 的差距（提交图分叉或降级中）。请人工确认：git log --oneline ${BASE_BRANCH}..${me}`)
  } else if (ab.behind > 0) {
    fail(
      `基线漂移：本分支落后 ${BASE_BRANCH} ${ab.behind} 个提交。\n    并行期间基准前进是正常的，但开工前应先 ${C.x}git merge ${BASE_BRANCH}${C.r} 对齐——\n    否则冲突会推迟到写代码时才暴露，那时改动已经铺开了。`
    )
  } else {
    line(`  ${C.g}✔${C.x} 与基准同步${ab.ahead > 0 ? `（领先 ${ab.ahead}）` : ''}`)
  }
}
line()

// 2) 本分支脏工作区 ---------------------------------------------------
const dirty = dirtyFiles()
if (dirty) {
  const n = dirty.size
  if (n > 0) warn(`本分支有 ${n} 项未提交改动（上一轮在制品，不阻断开工，但会计入重叠面）`)
  else line(`  ${C.g}✔${C.x} 工作区干净`)
} else {
  line(`  ${C.d}? 工作区状态未知（降级模式无法读 index）${C.x}`)
}
line()

// 3) 重叠检测（核心）--------------------------------------------------
const claims = readClaims()
const mineClaim = claims[me] || null
let myFiles
let blind = false // 改动面未知 → 禁止报绿

if (mineClaim) {
  myFiles = new Set(mineClaim.files)
  line(`  ${C.g}✔${C.x} 本分支已登记 ${C.d}(${mineClaim.topic})${C.x} ${myFiles.size} 个文件`)
} else {
  const union = new Set()
  const c = committedDiffFiles(me)
  if (c) c.forEach((f) => union.add(f))
  if (dirty) dirty.forEach((f) => union.add(f))
  myFiles = union
  // ⚠ 空集 + 别人有登记 = 必然"无重叠"，这是最危险的假绿。
  //   降级模式下 committedDiffFiles / dirtyFiles 都拿不到，
  //   若此时不拦，会对每条重叠都谎报绿灯。故置 blind，下方一律不报绿。
  blind = !c && !dirty
  if (blind) {
    fail(
      `本分支改动面未知（降级模式读不到 git，且未登记），无法判定重叠。\n` +
        `    必须先显式登记：${C.x}npm run pre:claim "<主题>=<文件1,文件2,...>"${C.r}\n` +
        `    ${C.x}否则本门禁只会谎报"无重叠"——假警报比没警报更坏，故此处直接拦截。`
    )
  } else {
    warn(
      me === BASE_BRANCH
        ? '当前在基准分支上，无并行分支可比对'
        : '本分支未登记改动面（建议 npm run pre:claim-auto "<主题>"），本次按实际改动推定'
    )
  }
}
line()

const others = Object.entries(claims).filter(([b]) => b !== me)
if (!others.length) {
  line(`  ${C.d}暂无其它并行方向登记${C.x}`)
} else {
  for (const [b, c] of others) {
    const ab = aheadBehind(b)
    const their = new Set(c.files)
    const overlap = [...myFiles].filter((f) => their.has(f))
    if (blind) {
      // ⚠ 改动面未知时 myFiles 是空集，overlap 必然为空。
      //   若这里照常报「无重叠」绿灯，等于对着一个未知集合宣称安全——
      //   上一层的 blind拦截拦不住这条路（那条 fail 是针对"没登记"的），
      //   同一状态下换个路径又漏出来。故此处再拦一次，且不重复计 fail。
      line(`  ${C.y}?${C.x} 与 ${C.b}${b}${C.x} ${C.d}(${c.topic})${C.x} ${C.d}无法比对——本分支改动面未知（见上）${C.x}`)
      continue
    }
    if (!overlap.length) {
      line(`  ${C.g}✔${C.x} 与 ${C.b}${b}${C.x} ${C.d}(${c.topic})${C.x} 无重叠${ab.unknown ? '' : ab.behind ? ` ${C.y}(该分支落后基准${ab.behind})${C.x}` : ''}`)
    } else {
      fail(`${C.r}与分支 ${b}（${c.topic}）改动面重叠 ${overlap.length} 个文件${C.x}`)
      overlap.forEach((f) => line(`      ${C.r}✘${C.x} ${f}`))
      line(`    ${C.d}两条路：① 串行——等 ${b} 合入 ${BASE_BRANCH} 再开工；${C.x}`)
      line(`    ${C.d}② 划死界——协商归属，各自 --claim 只登记自己那半，冲突点在合并前就谈清。${C.x}`)
    }
  }
}
line()

// 4) worktree 一览 ----------------------------------------------------
const wt = worktreeList()
if (wt && wt.length > 1) {
  line(`  ${C.d}worktree：${C.x}`)
  wt.forEach((w) => {
    const m = w.match(/^(.+?)\s+(\S+)\s+\[(.+?)\]/)
    if (m) line(`    ${m[3] === me ? `${C.g}▶${C.x}` : ' '} ${C.d}${m[1]}${C.x} ${C.d}[${m[3]}]${C.x}`)
  })
  line()
}

// 5) 汇总 ------------------------------------------------------------
warns.forEach((w) => line(`  ${C.y}!${C.x} ${w}`))
if (warns.length) line()
if (problems.length) {
  // ⚠ 必须把 problems 正文打出来。早先版本只打「N 个问题」，
  //   但 fail() 的文案是唯一说明，只有退出码没有正文——
  //   等于一道只说「有问题」却不说问题的门禁，跟只会说「通过」一样没用。
  problems.forEach((p, i) => {
    line(`  ${C.r}✘ ${i + 1}) ${p}${C.x}`)
    line()
  })
  line(`${C.r}✘ preflight 未通过 —— 先解决上面 ${problems.length} 个问题再开工${C.x}`)
  process.exit(1)
}
line(`${C.g}✔ preflight 通过 —— 可以开工${C.x}`)
