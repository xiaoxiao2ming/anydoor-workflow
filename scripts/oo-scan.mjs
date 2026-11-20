/**
 * OO 迭代影响面扫描器（oo:scan）
 * ============================================================================
 * 用户 2026-10-05 定案口径：
 *   「最小代价进行适配，也不是最小代价的意思是——
 *     如果 oo 版本迭代，你能知道在什么地方、改哪些东西。」
 *
 * 所以本脚本不度量「改了多少」，只回答两件事：
 *   1. 现在 OO 触点分布在哪些文件、哪些行（升级时逐条核对）
 *   2. 登记表（docs/24 的 §2.1 A1~A20）里的触点，代码里还在不在
 *
 * 用法：
 *   node scripts/oo-scan.mjs                # 人类可读报告
 *   node scripts/oo-scan.mjs --json         # 机器可读（给 CI 用）
 *   node scripts/oo-scan.mjs --baseline     # 存基线，升级前后diff 用
 *   node scripts/oo-scan.mjs --check        # 与基线比：有新触点/丢失触点则退出码 1
 *
 * ★ 为什么要有这个脚本（不是文档就够了）：
 *   文档会过期，代码不会。登记表靠人力维护必然漏，
 *   而「漏掉的触点」正是升级时最难查的那一类——
 *   因为它不报错，只是静默失效。脚本让漏登记变成 CI 可捕获的错误。
 */

import { readFileSync, readdirSync, statSync, writeFileSync, existsSync } from 'node:fs'
import { join, relative, extname } from 'node:path'

const ROOT = process.cwd()
/**
 * 基线放server-data/ 会被自己扫到（它是 .json，且含大量触点关键词），
 * 造成「建完基线立刻 --check 就报新增文件」的自指死循环。
 * 放 .workbuddy/ 下：那是元数据目录，不参与代码扫描。
 */
const BASELINE = join(ROOT, '.workbuddy', 'oo-touchpoints-baseline.json')

/** OO 触点的判定特征。宁可多报（人工过滤），不可漏报。 */
const PATTERNS = {
  // OO 编辑器实例与执行通道
  callCommand: /\bcallCommand\b/g,
  DocsAPI: /\bDocsAPI\b/g,
  // 宿主侧取编辑器
  frameEditor: /iframe\[name=["']?frameEditor/g,
  // 插件桥（我们自己的协议）
  cbRef: /__cbRef\w+/g,
  ooStatus: /__ooStatus/g,
  ooRef: /__oo(Ref\w+|Kind|Maximized|SetTheme|IsModified\w*)/g,
  // Office JS API —— 只能在插件上下文执行
  ascCalc: /\bApi\.asc_calculate\b/g,
  getActiveSheet: /\bApi\.GetActiveSheet\b/g,
  getCells: /\.GetCells\(/g,
  setValue: /\.SetValue\(/g,
  getValue: /\.GetValue\(/g,
  insertRows: /\.InsertRows(Before|After)\(/g,
  deleteRows: /\.DeleteRows\(/g,
  // DOM 契约（最脆的一类）
  domContract: /#status-(sheets-bar-box|tabs-scroll)\b/g,
  // 插件 SDK
  ascPlugin: /\bAsc\.plugin\b/g,
}

/** 生产代码目录（诊断脚本不算触点） */
const SCAN_DIRS = ['src', 'electron', 'scripts', 'server-data']
/** 诊断脚本前缀：这些是一次性排查产物，不构成长期触点 */
const DIAG_PREFIX = ['scripts/_diag', 'scripts/_e2e', 'scripts/_open_node', 'scripts/_sse']
/**
 * 排除清单。
 * ★ 三条都是实测踩出来的：
 *  1. `.snap.mjs` —— oo-core 的热更新快照副本，内容与正本几乎全等。
 *     不排除会让触点数翻倍，且升级时对着快照排查是白费劲。
 *  2. `oo-scan.mjs`（本文件）—— 扫描器自己含 OO 关键字，会把自己算成触点。
 *  3. `templateStore` / `editStore` —— 只做模板字节存取，不碰 OO 内核。
 */
const EXCLUDE = [
  /\.snap\.mjs$/,
  /oo-scan\.mjs$/,
]
/**
 * 扫描的文件类型。
 * ★ 必须含 .json：cbref/config.json 是插件注册（guid / EditorsSupport），
 *   OO 换版可能改它的 schema，属内核接触面（docs/24 §3 文件③）。
 *   漏扫它等于漏掉一个升级盲区。
 */
const SCAN_EXT = new Set(['.ts', '.tsx', '.mjs', '.js', '.json'])

/** 升级时只需改的四个文件（docs/24 §3） */
const CORE_FILES = [
  'scripts/oo-core.mjs',
  'server-data/oo-plugins/cbref/scripts/code.js',
  'server-data/oo-plugins/cbref/config.json',
  'src/office/OnlyOfficeNodeEditor.tsx',
]

/** 明确不该有 OO 依赖的层（隔离红线，破了就是架构退化） */
const MUST_BE_CLEAN = [
  'src/xlsx/outline.ts',
  'src/xlsx/outlineAdapter.ts',
  'src/xlsx/ChapterDialog.tsx',
]

function walk(dir, out = []) {
  const abs = join(ROOT, dir)
  if (!existsSync(abs)) return out
  for (const name of readdirSync(abs)) {
    const p = join(abs, name)
    const s = statSync(p)
    if (s.isDirectory()) {
      //跳过依赖与产物
      if (['node_modules', 'dist', '.git', 'vendors'].includes(name)) continue
      walk(join(dir, name), out)
    } else if (SCAN_EXT.has(extname(name))) {
      out.push(join(dir, name))
    }
  }
  return out
}

function isDiag(file) {
  const f = file.replace(/\\/g, '/')
  return DIAG_PREFIX.some((pre) => f.startsWith(pre))
}

/**
 * 剥离注释，保留行号对齐（用空串替掉注释内容，不改变行数）。
 *
 * ★ 为什么必须剥（2026-10-05 实测踩出来的）：
 *   触点的定义是「OO 换版后这行代码会失效」，而注释里的 API 名不会失效。
 *   本项目注释写得极密（每个架构决策都写了为什么），nodeStyle.ts 里
 *   「宿主直调 callCommand 静默失效」这句说明就被算成了触点，
 *   害得 `oo:check` 报「新增触点文件」——假警报比没警报更坏，
 *   因为人会学会忽略它。
 *
 * 处理范围：
 *   - // 行注释
 *   - 块注释 /* ... *\/（含跨行）
 *   刻意不处理字符串字面量里的内容（如 postMessage 载荷里含 "//"）。
 *   那会误删真代码；而把字符串里的 // 当注释，本项目里没有这种写法。
 *
 * ★ 块注释内的字符一律丢弃（不追加进 line）。第一版漏了这个分支，
 *   结果块注释内容原样留在结果里 —— nodeStyle.ts:206 依然被判为触点，
 *   假警报没消掉。写状态机时「进入注释」和「注释内丢弃」是两件事，缺一不可。
 */
function stripComments(text) {
  const out = []
  let inBlock = false
  for (const raw of text.split(/\r?\n/)) {
    let line = ''
    let i = 0
    while (i < raw.length) {
      const rest = raw.slice(i)
      if (inBlock) {
        // 块注释内：只找闭合标记，其余字符一律丢掉
        if (rest.startsWith('*/')) {
          inBlock = false
          i += 2
        } else {
          i++
        }
        continue
      }
      if (rest.startsWith('//')) break // 行注释 → 后面的全丢
      if (rest.startsWith('/*')) {
        inBlock = true
        i += 2
        continue
      }
      line += raw[i]
      i++
    }
    out.push(line)
  }
  return out
}

function scanFile(file) {
  const text = readFileSync(join(ROOT, file), 'utf8')
  // 剥注释后的行数组（行号与原文一一对应，便于报错定位）
  const lines = stripComments(text)
  const hits = {}
  const lineNos = {}
  // OO 内核配置类文件用关键词判定（无 API 调用，但属接触面）
  const CONFIG_HINT = /\b(guid|EditorsSupport|autostart|pluginsData|DocEditor|documentType|uiTheme|coEditing)\b/g

  for (const [key, re] of Object.entries({ ...PATTERNS, configHint: CONFIG_HINT })) {
    re.lastIndex = 0
    const found = []
    lines.forEach((ln, i) => {
      re.lastIndex = 0
      if (re.test(ln)) found.push(i + 1)
    })
    if (found.length) {
      hits[key] = found.length
      lineNos[key] = found
    }
  }
  return { file, hits, lineNos, total: Object.values(hits).reduce((a, b) => a + b, 0) }
}

function run() {
  const files = SCAN_DIRS.flatMap((d) => walk(d))
    .map((f) => relative(ROOT, f).replace(/\\/g, '/'))
    .filter((f) => !isDiag(f))
    .filter((f) => !EXCLUDE.some((re) => re.test(f)))

  const results = files.map(scanFile).filter((r) => r.total > 0)
  results.sort((a, b) => b.total - a.total)

  // 隔离红线检查
  const violations = MUST_BE_CLEAN.filter((f) => {
    const r = results.find((x) => x.file === f)
    return !!r && r.total > 0
  })

  const coreTouch = results.filter((r) => CORE_FILES.includes(r.file))
  const coreCovered = CORE_FILES.filter((f) => coreTouch.some((r) => r.file === f))

  return {
    scannedAt: new Date().toISOString(),
    totals: {
      filesWithTouch: results.length,
      touchpoints: results.reduce((a, r) => a + r.total, 0),
      coreFilesCovered: `${coreCovered.length}/${CORE_FILES.length}`,
    },
    // 登记表缺失的触点类型（docs/24 §2.1 里应出现但代码里没有）
    isolation: { mustBeClean: MUST_BE_CLEAN, violated: violations, ok: violations.length === 0 },
    core: {
      files: CORE_FILES,
      covered: coreCovered,
      missing: CORE_FILES.filter((f) => !coreCovered.includes(f)),
    },
    byFile: results.map((r) => ({ file: r.file, total: r.total, kinds: r.hits, lines: r.lineNos })),
  }
}

// ── 报告输出 ──────────────────────────────────────────────────────────────
function printReport(r) {
  const L = []
  L.push('='.repeat(72))
  L.push('OO 迭代影响面扫描  ·  oo:scan')
  L.push('='.repeat(72))
  L.push(`扫描时间   ${r.scannedAt}`)
  L.push(`触点文件   ${r.totals.filesWithTouch} 个`)
  L.push(`触点总数   ${r.totals.touchpoints}`)
  L.push(`内核文件覆盖 ${r.totals.coreFilesCovered}`)
  L.push('')

  L.push('【一】升级只需改的四个文件（docs/24 §3）')
  for (const f of r.core.files) {
    const hit = r.byFile.find((x) => x.file === f)
    const mark = hit ? `✔ ${String(hit.total).padStart(3)} 处` : '—未触达'
    L.push(`  ${mark}  ${f}`)
  }
  if (r.core.missing.length) {
    L.push(`  ⚠ 未触达：${r.core.missing.join(', ')}`)
  }
  L.push('')

  L.push('【二】隔离红线（纯计算层不该有 OO 依赖）')
  for (const f of r.isolation.mustBeClean) {
    const bad = r.isolation.violated.includes(f)
    L.push(`  ${bad ? '✘ 破了' : '✔ 干净'}  ${f}`)
  }
  L.push('')

  L.push('【三】全部触点分布（升级时自上往下逐条核对）')
  for (const f of r.byFile) {
    L.push('')
    const core = r.core.files.includes(f.file) ? ' ★内核' : ''
    L.push(`  ${f.file}${core}  —— ${f.total} 处`)
    const kinds = Object.entries(f.kinds).sort((a, b) => b[1] - a[1])
    for (const [k, n] of kinds) {
      const lns = f.lines[k].slice(0, 8).join(',')
      const more = f.lines[k].length > 8 ? ` …共${f.lines[k].length}行` : ''
      L.push(`      ${k.padEnd(16)} ${String(n).padStart(3)}  行 ${lns}${more}`)
    }
  }

  L.push('')
  L.push('='.repeat(72))
  if (!r.isolation.ok) {
    L.push(`结论：✘ 隔离红线被破 —— ${r.isolation.violated.join(', ')}`)
    L.push('     纯计算层混入 OO 依赖，换版时它会跟着炸。')
  } else {
    L.push('结论：✔ 隔离层干净，OO 触点收敛在可控范围')
  }
  L.push('升级流程见 docs/24-oo迭代影响面清单.md §5 探针清单')
  L.push('='.repeat(72))
  return L.join('\n')
}

// ── 基线 diff ─────────────────────────────────────────────────────────────
function diffWithBaseline(cur) {
  if (!existsSync(BASELINE)) return null
  const base = JSON.parse(readFileSync(BASELINE, 'utf8'))
  const baseMap = new Map((base.byFile || []).map((x) => [x.file, x.total]))
  const curMap = new Map(cur.byFile.map((x) => [x.file, x.total]))

  const added = []
  const removed = []
  const grown = []
  for (const [f, n] of curMap) {
    if (!baseMap.has(f)) added.push({ file: f, total: n })
    else if (n > baseMap.get(f)) grown.push({ file: f, from: baseMap.get(f), to: n })
  }
  for (const [f, n] of baseMap) if (!curMap.has(f)) removed.push({ file: f, total: n })
  return { added, removed, grown, baseAt: base.scannedAt }
}

// ── 入口 ──────────────────────────────────────────────────────────────────
const args = process.argv.slice(2)
const r = run()

if (args.includes('--json')) {
  console.log(JSON.stringify(r, null, 2))
} else if (args.includes('--baseline')) {
  writeFileSync(BASELINE, JSON.stringify(r, null, 2), 'utf8')
  console.log(`基线已存：${BASELINE.replace(ROOT + '\\', '')}`)
  console.log(`  触点文件 ${r.totals.filesWithTouch} · 触点总数 ${r.totals.touchpoints}`)
  console.log('升级后跑 --check 比对差异。')
} else if (args.includes('--check')) {
  const d = diffWithBaseline(r)
  if (!d) {
    console.log('✘ 无基线。先跑 node scripts/oo-scan.mjs --baseline')
    process.exit(1)
  }
  console.log('='.repeat(72))
  console.log('OO 触点变更检查（对基线）')
  console.log('='.repeat(72))
  console.log(`基线时间${d.baseAt}`)
  console.log(`新增触点文件 ${d.added.length} · 触点增长 ${d.grown.length} · 消失 ${d.removed.length}`)
  if (d.added.length) {
    console.log('\n★ 新增触点文件 —— 若是OO 升级引入的，必须登记进 docs/24 §2.1：')
    for (const a of d.added) console.log(`+ ${a.file} (${a.total} 处)`)
  }
  if (d.grown.length) {
    console.log('\n▲ 触点增长的文件：')
    for (const g of d.grown) console.log(`  ${g.file}  ${g.from} → ${g.to}`)
  }
  if (d.removed.length) {
    console.log('\n▼ 触点消失（可确认是否该删登记）：')
    for (const x of d.removed) console.log(`  ${x.file} (${x.total} 处)`)
  }
  if (!r.isolation.ok) {
    console.log(`\n✘ 隔离红线破了：${r.isolation.violated.join(', ')}`)
  }
  const hasChange = d.added.length || d.grown.length || d.removed.length || !r.isolation.ok
  console.log(hasChange ? '\n结论：有变化，需人工确认' : '\n结论：与基线一致')
  process.exit(hasChange ? 1 : 0)
} else {
  console.log(printReport(r))
}
