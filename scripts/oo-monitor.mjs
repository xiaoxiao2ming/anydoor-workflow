#!/usr/bin/env node
/**
 * ★ client-base 画布监控工具（2026-10-11）
 *
 * 用法：
 *   node scripts/oo-monitor.mjs          → 一次性完整状态报告
 *   node scripts/oo-monitor.mjs watch    → 实时监控（3 秒轮询，只打印变化）
 *
 * 监控范围：
 *   ① 全部节点（名称/类型/保存状态：ooBytes64 长度 + 模板绑定）
 *   ② 连线（每个节点的数据来源 inputs + 关联引用 ooRefs 数）
 *   ③ 节点保存文件（pages JSON、模板库文件、工作区文件、保存产物）
 *   ④ 服务健康（文档服务 3798 / 应用服务端口 / 渲染包 / oo-core 版本）
 */
import fs from 'node:fs'
import path from 'node:path'
import { execSync } from 'node:child_process'

const HOME = process.env.USERPROFILE || process.env.HOME
const APPDATA_DIR = path.join(HOME, 'AppData', 'Roaming', 'client-base')
const PAGES_JSON = path.join(APPDATA_DIR, 'pages', 'default.json')
const TPL_DIR = path.join(APPDATA_DIR, 'templates')
const WATCH = process.argv[2] === 'watch'
const STATE_FILE = path.join(process.cwd(), 'scripts', '.monitor-state.json')

const fmtT = (d) => (d ? new Date(d).toLocaleString('zh-CN', { hour12: false }) : '-')
const kb = (n) => (n / 1024).toFixed(1) + 'K'

function findAppPort() {
  try {
    const pids = execSync('tasklist /FI "IMAGENAME eq client-base.exe" /FO CSV 2>nul', { windowsHide: true, stdio: ['ignore', 'pipe', 'ignore'] }).toString()
      .split(/\r?\n/).slice(1).map((l) => l.split('"')[3]).filter(Boolean)
    const uniq = [...new Set(pids)]
    for (const pid of uniq) {
      const out = execSync(`netstat -ano | findstr LISTENING | findstr "${pid}"`, { windowsHide: true, stdio: ['ignore', 'pipe', 'ignore'] }).toString()
      for (const l of out.split(/\r?\n/)) {
        const m = l.match(/127\.0\.0\.1:(\d+)\s/)
        if (m && m[1] !== '3798' && m[1] !== '9334') return m[1]
      }
    }
  } catch { /* 无实例 */ }
  return ''
}

async function health(port) {
  const r = { app: '-', bundle: '-', ds: '-' }
  if (port) {
    try {
      const res = await fetch('http://127.0.0.1:' + port + '/', { signal: AbortSignal.timeout(3000) })
      r.app = 'HTTP ' + res.status
      const h = await res.text()
      r.bundle = (h.match(/assets\/index-[A-Za-z0-9_-]*\.js/) || ['?'])[0]
    } catch (e) { r.app = '✘ ' + String(e).slice(0, 30) }
  }
  try {
    const res = await fetch('http://127.0.0.1:3798/healthcheck', { signal: AbortSignal.timeout(3000) })
    r.ds = (await res.text()).trim() === 'true' ? '健康' : '异常'
  } catch { r.ds = '未运行' }
  return r
}

async function pocVersion(port) {
  try {
    const h = await (await fetch('http://127.0.0.1:' + port + '/poc?src=probe.xlsx', { signal: AbortSignal.timeout(4000) })).text()
    return (h.match(/__ocVer = '([^']*)'/) || [])[1] || '-'
  } catch { return '-' }
}

function collect() {
  const s = { time: Date.now(), nodes: {}, edges: [], files: {} }
  // ① 画布数据
  try {
    const st = fs.statSync(PAGES_JSON)
    s.pagesMtime = st.mtimeMs
    const j = JSON.parse(fs.readFileSync(PAGES_JSON, 'utf8'))
    for (const id in j.tables) {
      const t = j.tables[id]
      const cfg = t.config || {}
      s.nodes[id] = {
        name: (cfg.nodeNo ? "[" + cfg.nodeNo + "] " : "") + t.name, kind: t.kind, kind: t.kind,
        inputs: (cfg.inputs || []).join(','),
        ooRefs: Array.isArray(cfg.ooRefs) ? cfg.ooRefs.length : 0,
        ooBytes64: cfg.ooBytes64 ? String(cfg.ooBytes64).length : 0,
        ooTemplateId: cfg.ooTemplateId || '',
        locked: [cfg.ooDataLocked ? '数据锁' : '', cfg.ooPosLocked ? '位置锁' : ''].filter(Boolean).join('+'),
      }
      if (Array.isArray(cfg.linkedInputs) && cfg.linkedInputs.length) s.edges.push({ from: cfg.linkedInputs.join(','), to: t.name + '(' + id + ')', type: 'linkedInputs' })
      if (Array.isArray(cfg.inputs) && cfg.inputs.length) for (const src of cfg.inputs) s.edges.push({ from: src, to: t.name + '(' + id + ')', type: '数据来源' })
    }
  } catch (e) { s.pagesErr = String(e).slice(0, 60) }
  // ② 模板库
  try {
    for (const f of fs.readdirSync(TPL_DIR)) {
      const st = fs.statSync(path.join(TPL_DIR, f))
      if (st.isFile()) s.files['tpl:' + f] = { size: st.size, mtime: st.mtimeMs }
    }
  } catch { /* 无模板目录 */ }
  // ③ 工作区/保存产物（最新 8 个）
  const workDirs = [
    path.join(HOME, 'WorkBuddy/2026-09-24-19-48-07/migration/client-base-migration/client-base/out2/client-base/resources/app/sample-template/oo-poc'),
  ]
  for (const wd of workDirs) {
    try {
      const all = fs.readdirSync(wd).filter((f) => /^oo-oosheet-|^oo-ooword-|^save-oo-/.test(f))
        .map((f) => { const st = fs.statSync(path.join(wd, f)); return { f, m: st.mtimeMs, size: st.size } })
        .sort((a, b) => b.m - a.m).slice(0, 8)
      for (const it of all) s.files['work:' + it.f] = { size: it.size, mtime: it.m }
    } catch { /* 目录不存在 */ }
  }
  return s
}

function render(s, h, pocV, port) {
  const L = []
  L.push('════════ client-base 画布监控 ════════ ' + new Date(s.time).toLocaleString('zh-CN', { hour12: false }))
  L.push('服务: 应用端口 ' + (port || '未运行') + ' | ' + h.app + ' | 渲染包 ' + h.bundle + ' | 文档服务 ' + h.ds + ' | oo-core ' + pocV)
  if (s.pagesErr) L.push('⚠ 画布数据读取失败: ' + s.pagesErr)
  else L.push('画布: ' + Object.keys(s.nodes).length + ' 节点 | ' + s.edges.length + ' 条连线 | 画布文件保存于 ' + fmtT(s.pagesMtime))
  L.push('')
  L.push('── 节点（保存状态）──')
  for (const id in s.nodes) {
    const n = s.nodes[id]
    const save = n.ooBytes64 ? '💾已保存 ' + kb(n.ooBytes64) + (n.ooTemplateId ? ' tpl:' + n.ooTemplateId : '') : (n.kind.startsWith('oo') ? '⚪未保存' : '')
    L.push('  ' + n.name.padEnd(10) + ' [' + n.kind + '] ' + save + (n.locked ? ' 🔒' + n.locked : '') + '  id:' + id)
  }
  L.push('')
  L.push('── 连线（数据来源/关联）──')
  if (!s.edges.length) L.push('  （无）')
  for (const e of s.edges) L.push('  ' + e.from + ' ─→ ' + e.to + '  (' + e.type + ')')
  L.push('')
  L.push('── 保存文件（最新）──')
  const files = Object.entries(s.files).sort((a, b) => b[1].mtime - a[1].mtime).slice(0, 12)
  for (const [k, v] of files) L.push('  ' + fmtT(v.mtime) + '  ' + kb(v.size).padStart(6) + '  ' + k)
  return L.join('\n')
}

function diff(oldS, newS) {
  const out = []
  for (const id in newS.nodes) {
    const o = oldS.nodes[id], nn = newS.nodes[id]
    if (!o) out.push('[+] 新节点: ' + nn.name + ' (' + nn.kind + ')')
    else if (o.ooBytes64 !== nn.ooBytes64) out.push('[💾] ' + nn.name + ' 保存内容更新: ' + kb(o.ooBytes64) + ' → ' + kb(nn.ooBytes64))
    else if (o.ooTemplateId !== nn.ooTemplateId) out.push('[🔗] ' + nn.name + ' 模板绑定变更: ' + (o.ooTemplateId || '无') + ' → ' + (nn.ooTemplateId || '无'))
  }
  for (const id in oldS.nodes) if (!newS.nodes[id]) out.push('[-] 节点移除: ' + oldS.nodes[id].name)
  if (oldS.edges.length !== newS.edges.length) out.push('[⇄] 连线数: ' + oldS.edges.length + ' → ' + newS.edges.length)
  for (const k in newS.files) if (!oldS.files[k]) out.push('[📄] 新保存文件: ' + k + ' (' + kb(newS.files[k].size) + ')')
  for (const k in oldS.files) if (!newS.files[k]) out.push('[🗑] 文件移除: ' + k)
  return out
}

const snap = collect()
const port = findAppPort()
const h = await health(port)
const pocV = await pocVersion(port)
if (!WATCH) {
  console.log(render(snap, h, pocV, port))
  process.exit(0)
}
console.log(render(snap, h, pocV, port))
console.log('\n▶ 实时监控中（3 秒轮询，Ctrl+C 退出）...\n')
let last = snap
setInterval(async () => {
  const cur = collect()
  const p2 = findAppPort()
  const h2 = await health(p2)
  const d = diff(last, cur)
  if (d.length) { console.log('── 变化 ' + new Date().toLocaleTimeString('zh-CN', { hour12: false }) + ' ──'); d.forEach((l) => console.log('  ' + l)); console.log('') }
  last = cur
}, 3000)
