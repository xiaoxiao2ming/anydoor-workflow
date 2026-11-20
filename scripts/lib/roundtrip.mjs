/**
 * 真实文件往返回归 · 共享核心
 *
 * 为什么这是「门禁」而不是「测试」：
 *   @office-kit/xlsx 是本项目唯一会**改写用户文件**的依赖。
 *   它的往返行为（打开 → 改一格 → 存回）必须在真实文件上验证，
 *   因为所有「解析成自有模型再序列化」的库都有同一个死穴 ——
 *   对象模型不认识的东西会被**静默丢弃**（不报错，就是没了）。
 *
 *   用户改一个数字、存回去发现图表没了 —— 这是交付事故。
 *   所以升级 @office-kit/* 之后必须跑 check-regress.mjs，跑不过不许升。
 *
 * 与 check-roundtrip.mjs 的分工：
 *   · 本文件        共享实现（解压、比对、逐格校验、结果结构）
 *   · check-regress 清单全量 + 汇总 + 退出码（门禁）
 *   · check-roundtrip 单个文件的逐部件详细诊断（排查）
 */
import fs from 'node:fs'
import zlib from 'node:zlib'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

export const ROOT = path.resolve(fileURLToPath(new URL('../..', import.meta.url)))
export const OUT_DIR = path.join(ROOT, 'sample-template')

/* ────────────────── zip ────────────────── */

/**
 * 走**中央目录**解析 zip。
 *
 * 不能按本地文件头解析：很多写入器用 streaming + data descriptor，
 * 本地头里的 compSize 是 0，会把所有部件读成空。
 */
export function unzip(buf) {
  const out = {}
  let i = buf.length - 22
  while (i > 0 && buf.readUInt32LE(i) !== 0x06054b50) i--
  if (i <= 0) throw new Error('找不到 zip 中央目录（不是合法 zip？）')
  const cdOff = buf.readUInt32LE(i + 16)
  const cdCount = buf.readUInt16LE(i + 10)
  let p = cdOff
  for (let k = 0; k < cdCount; k++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) break
    const method = buf.readUInt16LE(p + 10)
    const crc = buf.readUInt32LE(p + 16)
    const csize = buf.readUInt32LE(p + 20)
    const usize = buf.readUInt32LE(p + 24)
    const nameLen = buf.readUInt16LE(p + 28)
    const extraLen = buf.readUInt16LE(p + 30)
    const cmtLen = buf.readUInt16LE(p + 32)
    const lho = buf.readUInt32LE(p + 42)
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen)
    const lnameLen = buf.readUInt16LE(lho + 26)
    const lextraLen = buf.readUInt16LE(lho + 28)
    const dataStart = lho + 30 + lnameLen + lextraLen
    const raw = buf.slice(dataStart, dataStart + csize)
    let data
    if (method === 0) data = raw
    else if (method === 8) data = zlib.inflateRawSync(raw)
    else data = Buffer.alloc(0)
    out[name] = { bytes: data, crc, size: usize }
    p += 46 + nameLen + extraLen + cmtLen
  }
  return out
}

/** zip 里以 "/" 结尾的条目是目录标记，不是文件 —— 不排除就会报成「丢失」的假警报 */
export const isDirEntry = (n) => n.endsWith('/')

/* ────────────────── 部件分类 ────────────────── */

const CATEGORY = [
  [/^xl\/charts\//, '图表'],
  [/^xl\/drawings\//, '绘图'],
  [/^xl\/pivotTables\//, '透视表'],
  [/^xl\/pivotCache\//, '透视缓存'],
  [/^xl\/comments/, '批注'],
  [/^xl\/tables\//, 'Excel 表格'],
  [/^xl\/media\//, '图片'],
  [/^xl\/worksheets\//, '工作表'],
  [/^xl\/styles/, '样式'],
  [/^xl\/sharedStrings/, '共享字符串'],
  [/^xl\/workbook\.xml$/, '工作簿定义'],
  [/^xl\/calcChain/, '计算链'],
  [/^xl\/theme/, '主题'],
  [/^docProps\//, '文档属性'],
  [/^\[Content_Types\]/, '内容类型'],
  [/^_rels|\.rels$/, '关系'],
]
export const categoryOf = (n) => (CATEGORY.find(([re]) => re.test(n)) ?? [/./, '其他'])[1]

/**
 * 「丢了就是事故」的部件。
 * 注意不含 worksheets / styles / sharedStrings —— 那些本来就会被改动。
 */
export const CRITICAL = [
  /^xl\/charts\//,
  /^xl\/drawings\//,
  /^xl\/pivot/,
  /^xl\/comments/,
  /^xl\/tables\//,
  /^xl\/media\//,
]

/* ────────────────── 单元格取值 ────────────────── */

/** 1-based 列号 → 列字母 */
export function colName(c) {
  let s = ''
  let n = c
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

/** 单元格的「规范值」：公式对象归一成 =formula 文本，undefined 归一成 '' */
export function valOf(ws, coord, getCellByCoord) {
  const c = getCellByCoord(ws, coord)
  const v = c && typeof c === 'object' && 'value' in c ? c.value : c
  if (v && typeof v === 'object' && 'formula' in v) return '=' + v.formula
  return v === undefined ? '' : v
}

/* ────────────────── 单文件往返 ────────────────── */

/**
 * 对一个文件跑完整往返并校验。
 *
 * @returns 结构化结果；任何异常都被收敛成 status，不往外抛
 *          （门禁脚本要跑完整张清单，不能因为一个文件挂掉就中断）
 */
export async function roundtripOne(entry, { log = () => {}, keepOut = false } = {}) {
  const file = path.isAbsolute(entry.path) ? entry.path : path.join(ROOT, entry.path)
  const label = entry.label || path.basename(file)
  const origin = entry.origin || 'unknown'
  const base = { file, label, origin }

  if (!fs.existsSync(file)) {
    return { ...base, status: 'missing' }
  }

  let lib, normMod, openErrMod
  try {
    lib = {
      io: await import('@office-kit/xlsx/io'),
      wb: await import('@office-kit/xlsx/workbook'),
      ws: await import('@office-kit/xlsx/worksheet'),
    }
    // 路径相对于本文件（scripts/lib/），所以是 ../../src/
    normMod = await import('../../src/xlsx/normalize.ts')
    openErrMod = await import('../../src/xlsx/openError.ts')
  } catch (e) {
    return { ...base, status: 'error', error: '依赖加载失败：' + String(e.message) }
  }
  const { loadWorkbook, fromArrayBuffer, workbookToBytes } = lib.io
  const { sheetNames, getSheet } = lib.wb
  const { getCellByCoord, setCellByCoord, getValueExtent, getNonEmptyCellCount } = lib.ws

  const srcBuf = fs.readFileSync(file)
  let before
  try {
    before = unzip(srcBuf)
  } catch (e) {
    return { ...base, status: 'error', error: '解压失败：' + String(e.message) }
  }

  // ── 打开（原文不行就过归一化）──
  let openTarget = new Uint8Array(srcBuf)
  let applied = []
  let openedRaw = false
  try {
    await loadWorkbook(fromArrayBuffer(openTarget))
    openedRaw = true
  } catch (firstErr) {
    // 走真实链路：normalize → openWorkbook 的分类逻辑
    const container = openErrMod.sniffContainer(openTarget)
    if (container !== 'zip') {
      const f = openErrMod.classifyOpenFailure(firstErr, openTarget, [])
      return { ...base, status: 'open-failed', openError: { kind: f.kind, title: f.title }, error: f.detail }
    }
    const norm = await normMod.normalizeWorkbookBytes(openTarget)
    applied = norm.applied
    openTarget = norm.bytes
    try {
      await loadWorkbook(fromArrayBuffer(openTarget))
    } catch (e) {
      const f = openErrMod.classifyOpenFailure(e, new Uint8Array(srcBuf), applied)
      return {
        ...base,
        status: 'open-failed',
        openError: { kind: f.kind, title: f.title, advice: f.advice, applied },
        error: f.detail,
      }
    }
  }
  log(`  打开 ✅${openedRaw ? '（原文即可）' : `（归一化 ${applied.length} 处）`}`)
  if (applied.length) applied.forEach((a) => log(`      · ${a}`))

  // ── 打开后的结构 ──
  const wb = await loadWorkbook(fromArrayBuffer(openTarget))
  const names = sheetNames(wb) ?? []
  log(`  工作表 ${names.length} 个: ${names.join(', ')}`)

  let totalNonEmpty = 0
  const extents = []
  for (const n of names) {
    const w = getSheet(wb, n)
    const e = w ? getValueExtent(w) : null
    const ne = w ? getNonEmptyCellCount(w) : 0
    totalNonEmpty += ne
    extents.push({ sheet: n, extent: e, nonEmpty: ne })
  }
  log(`  非空单元格合计 ${totalNonEmpty}`)

  // ── 改一格 → 存回 ──
  const first = names[0]
  const ws = getSheet(wb, first)
  let target = null
  const e0 = ws ? getValueExtent(ws) : null
  if (ws && e0) {
    outer: for (let r = e0.minRow; r <= Math.min(e0.maxRow, e0.minRow + 20); r++) {
      for (let c = e0.minCol; c <= Math.min(e0.maxCol, e0.minCol + 10); c++) {
        const coord = colName(c) + r
        const v = valOf(ws, coord, getCellByCoord)
        if (v !== '' && v !== null && v !== undefined) {
          target = { sheet: first, coord, old: v }
          break outer
        }
      }
    }
  }
  if (ws && target) setCellByCoord(ws, target.coord, '往返测试哨兵值')
  log(`  目标单元格 ${target ? `${target.sheet}!${target.coord}` : '(无)'}`)

  let outBuf
  try {
    // 补丁式保存：office-kit 只写它建模的部件，未建模的（图表/绘图/透视/媒体…）由 patchSave
    // 从原文件原样补回。不用它 —— 真实文件往返会静默丢图表（李渡可研估算表实测）。
    const { patchSaveWorkbook } = await import('../../src/xlsx/patchSave.ts')
    const ps = await patchSaveWorkbook(new Uint8Array(srcBuf), wb, {
      mtime: new Date('2026-01-01T00:00:00Z'),
    })
    outBuf = Buffer.from(ps.bytes)
    if (ps.restored.length) {
      log(`  补丁式保存：从原文件补回 ${ps.restored.length} 个未建模部件`)
      if (log !== (() => {}) && ps.restored.length <= 12) {
        for (const n of ps.restored) log(`      + ${n}`)
      }
    }
  } catch (e) {
    return { ...base, status: 'error', error: '序列化失败：' + String(e.message), openedRaw, applied }
  }
  if (keepOut) fs.writeFileSync(path.join(OUT_DIR, `roundtrip-${path.basename(file)}`), outBuf)

  const after = unzip(outBuf)

  // ── 部件比对 ──
  const srcParts = Object.keys(before).filter((n) => !isDirEntry(n))
  const outParts = Object.keys(after).filter((n) => !isDirEntry(n))
  const lost = []
  const same = []
  const changed = []
  for (const n of srcParts) {
    if (!(n in after)) lost.push(n)
    else if (before[n].crc === after[n].crc) same.push(n)
    else changed.push(n)
  }
  const added = outParts.filter((n) => !(n in before))
  const criticalLost = lost.filter((n) => CRITICAL.some((re) => re.test(n)))
  log(`  部件 ${srcParts.length} → ${outParts.length} · 一致 ${same.length} · 改动 ${changed.length} · 丢失 ${lost.length}${criticalLost.length ? `（关键 ${criticalLost.length}）` : ''}`)

  // ── 逐格数据校验 ──
  // 部件都在 ≠ 数据没坏，必须逐格比。
  let checked = 0
  let diff = 0
  const diffSamples = []
  const nonEmptyDeltas = []
  try {
    const wb2 = await loadWorkbook(fromArrayBuffer(new Uint8Array(outBuf)))
    for (const sheetName of names) {
      const a = getSheet(wb, sheetName)
      const b = getSheet(wb2, sheetName)
      if (!a || !b) {
        diff++
        diffSamples.push({ sheet: sheetName, coord: '(整表)', from: '存在', to: '产物里找不到' })
        continue
      }
      const ea = getValueExtent(a)
      if (!ea) continue
      for (let r = ea.minRow; r <= ea.maxRow; r++) {
        for (let c = ea.minCol; c <= ea.maxCol; c++) {
          const coord = colName(c) + r
          checked++
          if (target && sheetName === target.sheet && coord === target.coord) continue
          const va = String(valOf(a, coord, getCellByCoord))
          const vb = String(valOf(b, coord, getCellByCoord))
          if (va !== vb) {
            if (diffSamples.length < 8) diffSamples.push({ sheet: sheetName, coord, from: va, to: vb })
            diff++
          }
        }
      }
      nonEmptyDeltas.push({
        sheet: sheetName,
        before: getNonEmptyCellCount(a),
        after: getNonEmptyCellCount(b),
      })
    }
  } catch (e) {
    return {
      ...base,
      status: 'error',
      error: '产物无法重新打开（这本身就是严重问题）：' + String(e.message),
      openedRaw,
      applied,
      parts: { src: srcParts.length, out: outParts.length, same: same.length, changed: changed.length, lost, added, criticalLost },
    }
  }
  log(`  逐格 ${checked} 格，除目标格外差异 ${diff}`)

  return {
    ...base,
    status: 'ok',
    openedRaw,
    applied,
    sheets: names,
    nonEmpty: totalNonEmpty,
    extents,
    srcBytes: srcBuf.length,
    outBytes: outBuf.length,
    target,
    parts: {
      src: srcParts.length,
      out: outParts.length,
      same: same.length,
      changed: changed.length,
      lost,
      added,
      criticalLost,
    },
    cells: { checked, diff, diffSamples, nonEmptyDeltas },
  }
}

/* ────────────────── 清单 ────────────────── */

export function loadList(file = path.join(ROOT, 'scripts', 'regress-list.json')) {
  if (!fs.existsSync(file)) return { note: '', files: [] }
  return JSON.parse(fs.readFileSync(file, 'utf8'))
}
