// 并行开发基线比对 —— 任务开始前的强制前置检查
//
// 用户 2026-10-04 口径：
// > 「我打算 2 个及以上的同时开发 同一软件 不同方向
//   每次任务开始前 先对修改的 git 或者更新记录进行比对。」
//
// ★ 为什么必须做成脚本而不是口头约定：
//   两个方向同时改同一仓库时，人脑判不出「这些改动是谁的」。
//   本机实测（2026-10-04）：工作区一度堆着 3 个方向、20+ 文件、4000+ 行未提交改动，
//   其中还混着 451MB 未跟踪的 OnlyOffice 运行时 —— 任何方向的 `git add -A`
//   都会把二进制提交进库。库坏掉，「比对更新记录」本身就无从谈起。
//   所以「先比对」不能是自觉，得是**一条会失败的命令**。
//
// 五个用法：
//   node scripts/dev-baseline.mjs collect              先跑这个：打印 git 采集命令
//   node scripts/dev-baseline.mjs preflight <方向名>   开工前：查是否与其它方向重叠
//   node scripts/dev-baseline.mjs snapshot <方向名>   开工：打基线快照
//   node scripts/dev-baseline.mjs check               收工：比对该方向改了哪些文件
//   node scripts/dev-baseline.mjs report              看全部方向的基线与状态
//
// ★ 为什么每次都要先 collect：本机沙箱**禁止 node 派生子进程**
//   （spawnSync 任何程序都 EBUSY），本脚本无法自己跑 git，
//   只能读外部采集好的探针文件。第一版把这层做成
//   `execSync(...) + catch{return ''}`，结果静默失败 ——
//   工作区有 8173 行改动，它照样打印「一切正常」。

import { readFileSync, writeFileSync, existsSync, mkdirSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { execFileSync } from 'node:child_process'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const BASELINE_DIR = join(ROOT, '.workbuddy', 'dev-baselines')

/**
 * ★ 基线放 .workbuddy/ 而不是 server-data/ 或工程目录内。
 *   理由与 oo-scan.mjs 的基线同源：任何「工具自身的产物」必须放在
 *   工具自身的扫描/比对范围外，否则建完基线立刻 check 就报「新增文件」，
 *   形成自指死循环。
 */
const BASELINE = join(BASELINE_DIR, 'current.json')

/** 基线目录的 git 路径前缀 —— 用于把本工具的产物排除在「方向改动」之外 */
const BASELINE_REL_PREFIX = '.workbuddy/dev-baselines/'

/**
 * 动态记录文件。放 .workbuddy/ 而非 docs/ ——
 *   它是**每次比对追加**的运行时产物，一晚上能长几十条；
 *   放进 docs/ 会污染文档目录，也会让 git 历史全是无意义 diff。
 * ★ 但它是 gitignored（见 .gitignore ④）—— 这是**有意的取舍**：
 *   本地留痕就够了，不需要跨机同步。若将来要长期审计，再移进 docs/。
 */
const JOURNAL = join(BASELINE_DIR, 'journal.md')

/**
 * ★ git 不可用时的降级输入文件。
 *
 *   本机实测（2026-10-04）：**node 派生子进程被沙箱拦截**
 *   （spawnSync 任何可执行文件都报 EBUSY，连 cmd.exe、连 node 自己都起不来），
 *   所以本脚本不能自己跑 git。
 *
 *   这条降级路径的用途：在受限环境里，由外部（Bash 工具 / 用户手动）
 *   把 git 的原始输出落到这个文件，脚本再读它。
 *
 *   ⚠ 落这个文件时**不要用 -z / NUL 分隔**（Windows 下不便解析），
 *     一行一个路径即可。
 */
const PROBE = join(BASELINE_DIR, 'git-probe.txt')

/**
 * ★ 互斥清单：这些文件**同时只能属于一个方向**。
 *
 * 起因：`nodeStyle.ts` 的 KIND_STYLE 是全部节点样式的单一真相源
 * （角色/配色/连线规则/血缘都在这一份）。加节点 kind 的方向必然要改它，
 * 而另一个方向若也在改（比如给同一批 kind 换配色），git 无法自动合并同一处语义。
 *
 * 换句话说：这些文件是**语义枢纽**，不是普通的共享文件。
 * 两个方向同时改 = 必然有一方被覆盖，且 git 未必报冲突（因为改的是不同行）。
 * 这种「静默覆盖」比冲突可怕得多 —— 冲突会拦住你，静默覆盖不会。
 */
const EXCLUSIVE_FILES = [
  'src/core/nodeStyle.ts',
  'src/core/schema.ts',
  'src/index.css',
  'src/core/engine.ts',
  'package.json',
]

/** 允许共享的文件：纯新增、互不覆盖（各方向自己的产物） */
function isSharedOnly(files) {
  return files.every((f) => !EXCLUSIVE_FILES.includes(f))
}

/**
 * ★ 读 git 状态。优先用探针文件，因为本机沙箱禁止 node 派生子进程。
 *
 *   ⚠ 第一版这里写的是 `execSync('git ...')` + `catch { return '' }`，
 *     在本机表现为**全绿但全错**：spawnSync 报 EBUSY，错误被 catch 吞成空串，
 *     于是 report 打印「0 已跟踪 / 0 未跟踪 / 仓库体积 0.0 MB / 无提交」。
 *     工作区明明有 8173 行未提交改动，脚本却报告「一切正常」。
 *
 *     这是本脚本最该拦下的那一类 bug —— 一个只会说「通过」的检查，
 *     等于没有检查。所以现在缺探针文件时**直接退出码 1**，绝不猜。
 *
 * 探针文件是**带标记的纯文本**（不是 JSON）。这不是随手选的格式：
 *   第一版让 bash 拼 JSON 再塞进 `node -e "..."`，转义套了三层
 *   （bash 双引号 → JS 字符串 → 正则），`\\n` 变成字面量两字符，
 *   于是 log 的七行被当成一行、sizeMB 算成 0 —— 而文件本身完全合法，
 *   错得毫无痕迹。纯文本没有转义层，是这里唯一不会骗人的格式。
 *
 * 格式：每段以 `#name` 起头，其后为该段的行。
 *   #head / #branch / #changed / #untracked / #numstat / #size / #log
 */
function loadProbe() {
  if (!existsSync(PROBE)) {
    console.error('')
    console.error('✘ 读不到 git 状态：缺少探针文件')
    console.error('  路径：' + PROBE.replace(ROOT + '\\', ''))
    console.error('')
    console.error('  原因：本机沙箱禁止 node 派生子进程（spawnSync 任何程序都 EBUSY），')
    console.error('        本脚本无法自己执行 git。')
    console.error('')
    console.error('  先在项目根目录执行这条命令生成探针（可粘进 Bash 工具）：')
    console.error('    node scripts/dev-baseline.mjs collect')
    console.error('  它会打印一条现成的采集命令，执行后本脚本即可工作。')
    console.error('')
    console.error('  ⚠ 绝不因为读不到就按「工作区干净」继续 —— 那正是静默失败。')
    process.exit(1)
  }
  return parseProbe(readFileSync(PROBE, 'utf8'))
}

function parseProbe(text) {
  const sec = {}
  let cur = null
  for (const line of text.split(/\r?\n/)) {
    const m = line.match(/^#(\w+)$/)
    if (m) {
      cur = m[1]
      sec[cur] = []
      continue
    }
    if (cur && line.length) sec[cur].push(line)
  }
  // ★ 必须锚定 `size-pack:` —— 第一版写 /([\d.]+)\s*(\w+)/ 在整段文本上匹配，
  //   `count: 0` 抢先命中，sizeMB 恒为 0，于是「仓库超 500MB」这条
  //   并行开发的硬中止条件永远不会触发。又是一个「静默通过」的实例。
  const m = (sec.size ?? []).join(' ').match(/size-pack:\s*([\d.]+)\s*(\w+)/)
  return {
    head: (sec.head ?? [])[0] ?? '(未知)',
    branch: (sec.branch ?? [])[0] ?? '(未知)',
    // ★ 排除本工具自己的产物（探针 / 基线），否则建完基线立刻 check
    //   就报「本方向新增了 current.json + git-probe.txt」——
    //   自指死循环，跟 oo-scan.mjs 当初踩的是同一个坑。
    changed: (sec.changed ?? []).filter((f) => !f.startsWith('.')),
    untracked: (sec.untracked ?? []).filter((f) => !f.startsWith(BASELINE_REL_PREFIX)),
    numstat: (sec.numstat ?? []).map((l) => {
      const [file, add, del] = l.split('\t')
      return { file, add, del }
    }),
    sizeMB: m ? (m[2] === 'GiB' ? +m[1] * 1024 : +m[1]) : 0,
    log: sec.log ?? [],
  }
}

/**
 * 打印采集命令。单独一个命令，而不是让用户自己拼 ——
 * 拼错一条 git 参数就会得到一份错误的状态，而错误的比对结果比没有比对更糟。
 */
function collect() {
  const rel = PROBE.replace(ROOT + '\\', '').replace(/\\/g, '/')
  const cmd = [
    `cd "${ROOT.replace(/\\/g, '/')}"`,
    `{`,
    `  echo "#head";      git rev-parse --short HEAD`,
    `  echo "#branch";    git branch --show-current`,
    `  echo "#changed";   git diff --name-only HEAD`,
    `  echo "#untracked"; git ls-files --others --exclude-standard`,
    `  echo "#numstat";   git diff --numstat HEAD`,
    `  echo "#size";      git count-objects -vH`,
    `  echo "#log";       git log --oneline -8`,
    `} > "${rel}"`,
    `echo "探针已写入 ${rel}"`,
  ].join('\n')
  console.log('='.repeat(68))
  console.log('并行开发 · git 状态采集')
  console.log('='.repeat(68))
  console.log('本机沙箱禁止 node 派生子进程，需由外部执行下面这段来采集 git 状态。')
  console.log('（整段复制粘贴到项目根目录的 Bash 里执行）')
  console.log('')
  console.log(cmd)
  console.log('')
  console.log('执行完再跑 preflight / snapshot / check / report 即可。')
  console.log('='.repeat(68))
  return 0
}

/* ────────────────────────── 动态记录 ────────────────────────── */

/**
 * 追加一条到 journal.md —— 规则 docs/25 的核心产出物。
 *
 * ★ 为什么要机器写而不是人写：
 *   1. 人记不住。一晚上跑三次比对，谁改了哪些文件，靠脑子记必漏。
 *   2. git log 不够。git 只记提交粒度；方向内的**中间态**
 *      （改了没提交的、改到一半叫停的）在 git log 里根本不存在，
 *      而这些恰恰最容易撞车。
 *   3. 撞车后要能追溯「谁先动的、动之前是什么样」。
 *
 * ★ 为什么追加在「人写的决策段」之后而不是文件末尾：
 *   journal 是**倒序**的（最新在最上面），人写的决策段固定在末尾一个分隔线之下。
 *   若直接 append 会把新记录追加到人的决策段中间，把结构冲乱。
 *
 * 只追加，人写的内容一个字都不动。
 */
function appendJournal(entry) {
  mkdirSync(BASELINE_DIR, { recursive: true })
  const isNew = !existsSync(JOURNAL)
  const HEADER = `# 并行开发 · 动态比对记录

> 本文件由 \`npm run dev:diffcheck\` **自动追加**，请勿手工编辑记录区。
> 规则见 \`docs/25-并行协同开发规则.md\`。
> 倒序：最新记录在最上面。人写的「决策与备注」在文件末尾，不参与自动追加。

`
  let text = isNew ? HEADER : readFileSync(JOURNAL, 'utf8')
  // ★ 必须插在**表头之后**（倒序：最新在最上面）。
  //   第一版用 append，导致越新的记录越靠下—— 与规则文档宣称的「倒序」
  //   矛盾。规则与实现不一致时，人会按规则去找记录然后找不到，
  //   那这条记录等于没记。
  //   插入点：表头的空行之后、第一条 '### ' 之前。
  const firstEntry = text.indexOf('### ')
  const block = entry + '\n'
  if (firstEntry >= 0) {
    text = text.slice(0, firstEntry) + block + text.slice(firstEntry)
  } else {
    // 还没有任何记录：插在「## 决策与备注」之前（若有），否则追加到末尾
    const at = text.indexOf('## 决策与备注')
    if (at >= 0) text = text.slice(0, at) + block + '\n' + text.slice(at)
    else text = text.replace(/\s*$/, '\n\n') + block
  }
  writeFileSync(JOURNAL, text, 'utf8')
  return isNew
}

/** 生成本次比对的一行摘要 */
function journalLine(base, st, added, touches) {
  const now = new Date()
  const stamp =
    String(now.getFullYear()) +
    '-' + String(now.getMonth() + 1).padStart(2, '0') +
    '-' + String(now.getDate()).padStart(2, '0') +
    ' ' + String(now.getHours()).padStart(2, '0') +
    ':' + String(now.getMinutes()).padStart(2, '0') +
    ':' + String(now.getSeconds()).padStart(2, '0')
  const addedFiles = added.length ? added.join('、') : '（无新增文件）'
  const ex = touches.length ? touches.join('、') : '无'
  return (
    `### ${stamp} ｜ ${base.direction}\n` +
    `- HEAD：\`${base.head}\` → \`${st.head}\`（${st.branch}）\n` +
    `- 新增/改动文件 **${added.length}** 个：${addedFiles}\n` +
    `- 触及互斥文件：**${ex}**\n` +
    `- 仓库体积 ${(st.sizeMB ?? 0).toFixed(1)} MB\n`
  )
}

/* ────────────────────────── 快照 ────────────────────────── */

/**
 * 打基线快照。
 * 记录：HEAD 提交号 / 工作区已跟踪改动 / 未跟踪文件。
 * 之后 check 时与之比对，即得「这个方向到底动了哪些文件」。
 */
function snapshot(direction) {
  const st = loadProbe()
  mkdirSync(BASELINE_DIR, { recursive: true })
  const data = {
    direction,
    at: new Date().toISOString(),
    head: st.head,
    branch: st.branch,
    changed: st.changed,
    untracked: st.untracked,
  }
  writeFileSync(BASELINE, JSON.stringify(data, null, 2), 'utf8')

  const exclusive = data.changed.filter((f) => EXCLUSIVE_FILES.includes(f))
  console.log('='.repeat(68))
  console.log(`并行开发基线已建立  ·  方向：${direction}`)
  console.log('='.repeat(68))
  console.log(`HEAD    ${data.head}（${data.branch}）`)
  console.log(`基线文件 ${BASELINE.replace(ROOT + '\\', '')}`)
  console.log(`已跟踪改动 ${data.changed.length} · 未跟踪 ${data.untracked.length}`)
  if (exclusive.length) {
    console.log('')
    console.log(`⚠ 该方向触及**互斥文件** ${exclusive.length} 个：`)
    for (const f of exclusive) console.log(`   · ${f}`)
    console.log('  这些是语义枢纽。并行期间若有其它方向也在改，git 可能不报冲突却静默覆盖。')
    console.log('  → 收工时务必用 `check` 复核，并在提交说明里写清改了什么。')
  }
  console.log('='.repeat(68))
  return 0
}

/* ────────────────────────── 比对 ────────────────────────── */

function check() {
  if (!existsSync(BASELINE)) {
    console.log('✘ 没有基线，先跑：node scripts/dev-baseline.mjs snapshot <方向名>')
    return 1
  }
  const base = JSON.parse(readFileSync(BASELINE, 'utf8'))
  const st = loadProbe()
  const baseAll = new Set([...base.changed, ...base.untracked])
  const nowAll = new Set([...st.changed, ...st.untracked])

  const added = [...nowAll].filter((f) => !baseAll.has(f)).sort()
  const removed = [...baseAll].filter((f) => !nowAll.has(f)).sort()

  console.log('='.repeat(68))
  console.log(`方向「${base.direction}」改动比对  ·  基线时间 ${base.at}`)
  console.log('='.repeat(68))
  console.log(`基线 HEAD ${base.head} → 当前 ${st.head}`)

  if (added.length === 0) {
    console.log('本方向未新增任何文件（可能改动都在已跟踪文件的内容里 → 看下方 diff 统计）')
  } else {
    console.log('')
    console.log(`【本方向新增的文件】${added.length} 个`)
    for (const f of added) {
      const tag = EXCLUSIVE_FILES.includes(f) ? '★互斥' : ''
      console.log(`  + ${f}${tag ? '  ' + tag : ''}`)
    }
  }
  if (removed.length) {
    console.log('')
    console.log('【本方向消失的文件】(可能被别的方向提交了 —— 需确认没被误删)')
    for (const f of removed) console.log(`  - ${f}`)
  }

  // 互斥文件的实际改动量：用 numstat 探针字段
  const touchesExclusive = (st.numstat ?? []).some((r) => EXCLUSIVE_FILES.includes(r.file))

  console.log('')
  console.log('─'.repeat(68))
  if (touchesExclusive) {
    console.log('⚠ 本方向触及互斥文件。提交前请逐条确认：')
    for (const r of st.numstat ?? []) {
      if (EXCLUSIVE_FILES.includes(r.file)) {
        console.log(`   ${r.file}: +${r.add} / -${r.del} 行`)
      }
    }
  } else {
    console.log('✔ 本方向未触及互斥文件，可独立提交')
  }
  console.log(`仓库体积 ${(st.sizeMB ?? 0).toFixed(1)} MB`)
  if ((st.sizeMB ?? 0) > 500) {
    console.log('⚠ 仓库已超 500MB。若提交了 OnlyOffice 运行时等二进制，')
    console.log('  改用 docs/24 的做法：记版本号，不提交二进制。')
  }

  // ★ 追加动态记录（规则 docs/25 §4）。只在比对成功走到这一步时才写——
  //   半路失败不该留一条「本方向改了什么」的记录，那会误导追溯。
  const exFiles = (st.numstat ?? []).filter((r) => EXCLUSIVE_FILES.includes(r.file)).map((r) => r.file)
  appendJournal(journalLine(base, st, added, exFiles))
  console.log(`已追加记录到 ${JOURNAL.replace(ROOT + '\\', '')}`)

  console.log('='.repeat(68))
  return 0
}

/* ────────────────────────── 开工前置 ────────────────────────── */

/**
 * 开工前置检查 —— 回答「我现在能不能安全动手」。
 *
 * 三条硬条件，任一不满足就退出码 1：
 *   ① 有未提交的工作区改动且方向不明→ 先把上一轮切干净
 *   ② 库体积异常（多半是二进制被提交进去了）
 *   ③ 已在某个方向的基线里（说明这个方向正在进行中）
 */
function preflight(direction) {
  const problems = []
  const st = loadProbe()

  const changed = st.changed
  const untrackedList = st.untracked
  const dirty = changed.length + untrackedList.length

  if (existsSync(BASELINE)) {
    const base = JSON.parse(readFileSync(BASELINE, 'utf8'))
    if (base.direction === direction) {
      problems.push(`方向「${direction}」已有进行中的基线（${base.at}）。`);
      problems.push('  同一方向不应重复开工 —— 先 check 收工，或换一个方向名。')
    } else {
      console.log(`ℹ 存在另一方向的基线：「${base.direction}」（${base.at}）`)
      console.log('  → 两方向并行时**不要共用工作区**，见下方「并行隔离」提示。')
    }
  }

  if (dirty > 0 && !existsSync(BASELINE)) {
    problems.push(`工作区有 ${dirty} 处未提交改动（${changed.length} 已跟踪 / ${untrackedList.length} 未跟踪）。`)
    problems.push('  并行开工前必须先切干净：否则无法区分改动属于哪个方向。')
    console.log('')
    console.log('未跟踪文件（确认这些该入库还是该 ignore）：')
    for (const f of untrackedList.slice(0, 15)) console.log(`  ? ${f}`)
    if (untrackedList.length > 15) console.log(`  ...另有 ${untrackedList.length - 15} 个`)
  }

  const size = st.sizeMB ?? 0
  if (size > 500) {
    problems.push(`仓库体积 ${size.toFixed(0)}MB 异常，多半有二进制被提交进去。`)
  }

  console.log('='.repeat(68))
  console.log('并行开发 · 开工前置检查')
  console.log('='.repeat(68))
  console.log(`目标方向：${direction ?? '(未指定)'}`)
  console.log(`HEAD    ${st.head}（${st.branch}）`)
  console.log(`工作区：${dirty === 0 ? '干净' : `${dirty} 处未提交改动`}`)
  console.log(`仓库体积：${size.toFixed(1)} MB`)

  // 二进制风险
  const binaryRisk = untrackedList.filter((f) => /\.(dll|exe|wasm|node)$/.test(f) || f.includes('vendors/'))
  if (binaryRisk.length) {
    console.log('')
    console.log('⚠ 未跟踪的二进制/运行时（提交前务必确认是否该入库）：')
    for (const f of binaryRisk.slice(0, 8)) console.log(`  ! ${f}`)
  }

  if (problems.length) {
    console.log('')
    console.log('✘ 暂不能开工：')
    for (const p of problems) console.log('  ' + p)
    console.log('')
    console.log('建议动作：')
    console.log('  1) 先把工作区切干净（按方向分批 commit，见 git log 近期提交风格）')
    console.log('  2) 每个方向用**独立 git worktree**，别共用工作区')
    console.log('     git worktree add ../client-base-<方向> -b feat/<方向>')
    console.log('     这样两个方向物理隔离，git 无法产生静默覆盖')
    console.log('  3) 重新跑本脚本确认通过')
    console.log('='.repeat(68))
    return 1
  }

  console.log('')
  console.log('✔ 前置检查通过')
  if (direction) {
    console.log(`下一步：node scripts/dev-baseline.mjs snapshot ${direction}`)
  }
  console.log('='.repeat(68))
  return 0
}

/* ────────────────────────── 总览 ────────────────────────── */

function report() {
  const st = loadProbe()
  console.log('='.repeat(68))
  console.log('并行开发 · 基线总览')
  console.log('='.repeat(68))
  if (!existsSync(BASELINE)) {
    console.log('暂无基线')
  } else {
    const b = JSON.parse(readFileSync(BASELINE, 'utf8'))
    console.log(`方向    ${b.direction}`)
    console.log(`时间    ${b.at}`)
    console.log(`HEAD    ${b.head}（${b.branch}）`)
    console.log(`基线时  ${b.changed.length} 已跟踪 / ${b.untracked.length} 未跟踪`)
  }
  console.log('')
  console.log(`当前 HEAD  ${st.head}（${st.branch}）`)
  console.log(`当前工作区 ${st.changed.length} 已跟踪 / ${st.untracked.length} 未跟踪`)
  console.log(`仓库体积 ${(st.sizeMB ?? 0).toFixed(1)} MB`)
  console.log('')
  console.log('最近 8 次提交：')
  for (const l of st.log ?? []) console.log('  ' + l)
  console.log('')
  console.log('互斥文件（同时只允许一个方向改）：')
  for (const f of EXCLUSIVE_FILES) console.log(`  · ${f}`)
  console.log('='.repeat(68))
  return 0
}

/** 看动态记录。-n 只看最近 n 条 */
function showJournal(n) {
  if (!existsSync(JOURNAL)) {
    console.log('还没有记录。跑一次 `npm run dev:diffcheck` 就会自动追加。')
    return 0
  }
  const text = readFileSync(JOURNAL, 'utf8')
  const lines = text.split(/\r?\n/)
  if (n) {
    // 只取最近 n 条。记录是**倒序**的（最新在最上面），
    // 所以最近 n 条 = 前 n 个'### ' 块 → 从第 n 个 '### ' 切到文件尾。
    // ★ 第一版写成「从第 (总数-n) 个切」，那是最旧 n 条 —— 与命令名相反。
    const idx = []
    lines.forEach((l, i) => { if (l.startsWith('### ')) idx.push(i) })
    if (idx.length <= n) {
      console.log(lines.join('\n'))
    } else {
      console.log(lines.slice(idx[n - 1]).join('\n'))
    }
  } else {
    console.log(text)
  }
  const count = lines.filter((l) => l.startsWith('### ')).length
  console.log('─'.repeat(68))
  console.log(`共 ${count} 条记录 · ${JOURNAL.replace(ROOT + '\\', '')}`)
  return 0
}

/* ────────────────────────── 入口 ────────────────────────── */

const [cmd, arg] = process.argv.slice(2)
switch (cmd) {
  case 'collect':
    process.exit(collect())
  case 'snapshot':
    process.exit(snapshot(arg ?? '未命名方向'))
  case 'check':
    process.exit(check())
  case 'journal':
    process.exit(showJournal(arg ?Number(arg) : 0))
  case 'preflight':
    process.exit(preflight(arg))
  case 'report':
    process.exit(report())
  default:
    console.log('用法：')
    console.log('  node scripts/dev-baseline.mjs collect              打印 git 采集命令（先跑这个）')
    console.log('  node scripts/dev-baseline.mjs preflight <方向名>   开工前前置检查')
    console.log('  node scripts/dev-baseline.mjs snapshot <方向名>   打基线')
    console.log('  node scripts/dev-baseline.mjs check               收工比对（自动追加 journal）')
    console.log('  node scripts/dev-baseline.mjs journal [n]          看动态记录（默认全量，n=只看最近 n 条）')
    console.log('  node scripts/dev-baseline.mjs report              总览')
    process.exit(0)
}
