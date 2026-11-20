#!/usr/bin/env node
/**
 * 节点接入点定位器（2026-11-07）
 *
 * 用处：给出「某个节点类型改哪里」的**精确文件:行**，不靠记忆、不靠手抄表。
 * 真相源仍是 src/core/nodeStyle.ts 的 KIND_STYLE；本脚本只做「按 kind 反查引用点 + 分类」。
 *
 * 用法：
 *   node scripts/node-map.mjs                # 全量表格（markdown）
 *   node scripts/node-map.mjs pdfmerge       # 只看一个 kind（详细到行号）
 *   node scripts/node-map.mjs --json         # 机读
 */
import fs from 'node:fs'
import path from 'node:path'

const ROOT = path.resolve(import.meta.dirname, '..')
const STYLE = path.join(ROOT, 'src/core/nodeStyle.ts')

/** 热区：真正承载「定义 / 渲染 / 编辑器 / 执行 / 属性 / 意图」的目录 */
const HOT_DIRS = ['src', 'electron']
const SKIP = /(\.test\.|node_modules|\.d\.ts$)/

/** 文件 → 接入面分类。顺序即优先级（先命中先算）。 */
const FACES = [
  { key: '定义', test: (p) => p.endsWith('core/nodeStyle.ts') },
  { key: '意图', test: (p) => /core\/chatIntent|CanvasChatDock/.test(p) },
  { key: '画布', test: (p) => /(^|\/)(Canvas|CanvasParts)\.tsx$/.test(p) },
  { key: '编辑器', test: (p) => /(^|\/)(NodeEditor|PopupNodeEditor)\.tsx$/.test(p) },
  { key: '属性', test: (p) => /PropertiesPanel\.tsx$/.test(p) },
  { key: '卡片', test: (p) => /office\/.*(Card|Buttons)\.tsx$/.test(p) },
  { key: '内核', test: (p) => /src\/(core|xlsx|docx|pdf|ref|pptx)\//.test(p) },
  { key: '其他', test: () => true },
]

function walk(dir, out = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name)
    if (e.isDirectory()) {
      if (e.name === 'node_modules' || e.name === 'dist' || e.name === 'dist-electron') continue
      walk(p, out)
    } else if (/\.(ts|tsx|mjs|cjs)$/.test(e.name) && !SKIP.test(p)) out.push(p)
  }
  return out
}

/**
 * 从 nodeStyle.ts 解析 kind → { label, role, line }
 * 逐行扫描（不用正则跨行）：KIND_STYLE 表里的键固定是 2 空格缩进 + `: {`，
 * 块结束是同样 2 空格的 `}`。行扫描对 CRLF / 中间夹注释都免疫。
 */
function readKinds() {
  const lines = fs.readFileSync(STYLE, 'utf8').split(/\r?\n/)
  const kinds = []
  for (let i = 0; i < lines.length; i++) {
    const head = /^ {2}([a-z][a-z0-9_]*): \{$/.exec(lines[i])
    if (!head) continue
    let end = i + 1
    while (end < lines.length && !/^ {2}\},?$/.test(lines[end])) end++
    const body = lines.slice(i + 1, end).join('\n')
    if (!/^\s+label: '/m.test(body)) continue
    kinds.push({
      kind: head[1],
      label: (body.match(/label: '([^']*)'/) || [, '?'])[1],
      role: (body.match(/role: '([^']*)'/) || [, '?'])[1],
      line: i + 1,
    })
    i = end
  }
  return kinds
}

function faceOf(rel) {
  for (const f of FACES) if (f.test(rel)) return f.key
  return '其他'
}

function scan(kinds) {
  const files = HOT_DIRS.flatMap((d) => {
    const abs = path.join(ROOT, d)
    return fs.existsSync(abs) ? walk(abs) : []
  })
  /** kind → face → [{ file, lines }] */
  const map = new Map(kinds.map((k) => [k.kind, new Map()]))
  for (const abs of files) {
    const rel = path.relative(ROOT, abs).replace(/\\/g, '/')
    const text = fs.readFileSync(abs, 'utf8')
    const lines = text.split('\n')
    for (const { kind } of kinds) {
      const hits = []
      const needle = new RegExp(`['"\`]${kind}['"\`]`)
      lines.forEach((ln, i) => {
        if (needle.test(ln)) hits.push(i + 1)
      })
      if (!hits.length) continue
      const face = faceOf(rel)
      const bucket = map.get(kind)
      if (!bucket.has(face)) bucket.set(face, [])
      bucket.get(face).push({ file: rel, lines: hits })
    }
  }
  return map
}

const ORDER = ['定义', '画布', '编辑器', '属性', '卡片', '内核', '意图', '其他']

function main() {
  const kinds = readKinds()
  const map = scan(kinds)
  const only = process.argv[2] && !process.argv[2].startsWith('--') ? process.argv[2] : null

  if (process.argv.includes('--json')) {
    const out = kinds.map((k) => ({
      ...k,
      faces: Object.fromEntries([...(map.get(k.kind) || new Map())].map(([f, arr]) => [f, arr])),
    }))
    process.stdout.write(JSON.stringify(out, null, 2))
    return
  }

  /** 单 kind 详表：逐行号 */
  if (only) {
    const k = kinds.find((x) => x.kind === only)
    if (!k) {
      console.error(`未知 kind: ${only}\n可选：${kinds.map((x) => x.kind).join(', ')}`)
      process.exit(1)
    }
    console.log(`# ${k.kind}（${k.label} · role=${k.role}）定义于 src/core/nodeStyle.ts:${k.line}\n`)
    const bucket = map.get(only) || new Map()
    for (const face of ORDER) {
      const arr = bucket.get(face)
      if (!arr) continue
      console.log(`## ${face}`)
      for (const f of arr) console.log(`- ${f.file}: ${f.lines.join(', ')}`)
      console.log('')
    }
    return
  }

  /** 全量 markdown 表格 */
  console.log('| kind | 名称 | 角色 | 定义 | 画布 | 编辑器 | 属性 | 卡片 | 内核 | 意图 |')
  console.log('|---|---|---|---|---|---|---|---|---|---|')
  for (const k of kinds) {
    const b = map.get(k.kind) || new Map()
    const cell = (face) => {
      const arr = b.get(face)
      if (!arr) return '—'
      const total = arr.reduce((n, f) => n + f.lines.length, 0)
      return `${arr.length} 文件 / ${total} 处`
    }
    console.log(
      `| \`${k.kind}\` | ${k.label} | ${k.role} | nodeStyle:${k.line} | ${cell('画布')} | ${cell('编辑器')} | ${cell('属性')} | ${cell('卡片')} | ${cell('内核')} | ${cell('意图')} |`,
    )
  }
  console.log(`\n共 ${kinds.length} 种 kind。单查：\`node scripts/node-map.mjs <kind>\``)
}

main()
