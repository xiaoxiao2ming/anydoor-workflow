import { useCallback, useEffect, useMemo, useRef, useState, forwardRef, useImperativeHandle, type ForwardedRef } from 'react'
import { createPortal } from 'react-dom'
import { AgGridReact } from 'ag-grid-react'
import type { ColDef, CellValueChangedEvent, CellClickedEvent, GridApi, GridReadyEvent } from 'ag-grid-community'
import type { Workbook } from '@office-kit/xlsx/workbook'
import { getSheet, sheetNames } from '@office-kit/xlsx/workbook'
import {
  getCellByCoord,
  setCellByCoord,
  getValueExtent,
  listTables,
  getMergedCells,
  getNonEmptyCellCount,
  setComment,
  listComments,
  removeAllComments,
  setSheetZoom,
  autofitColumns,
  setColumnWidth,
  clearRange,
} from '@office-kit/xlsx/worksheet'
import { makeFormula } from '@office-kit/xlsx/cell'
import { openWorkbook, discoverXlsxFields, type XlsxFieldSpec } from './field'
import { patchSaveWorkbook } from './patchSave'
import { recomputeSheet, collectRefValues } from './recompute'
import { collectUrnSources, scanUrnCells, type UrnCellRef } from './urnScan'
import { evalRowFormula, isPureUrn, FUNC_NAMES, FUNC_COUNT } from './evalFormula'
import {
  insertRowAbove,
  deleteRow as deleteRowAt,
  insertColLeft,
  deleteCol as deleteColAt,
  mergeSel,
  unmergeSel,
  alignSel,
  numberFmtSel,
  fontSel,
  freezeAt,
  pasteTSV,
  insertCellShift,
  addSheetTitled,
  renameSheetTitled,
  removeSheetTitled,
  parseCoord,
  fmtOfFirst,
  applyFmtSel,
  sortRowsByCol,
  parseTSVGrid,
  transposeGrid,
  pasteGrid,
  fmtSnapshotOfRect,
  applyFmtSnapshot,
} from './editOps'
import type { CellRect, FmtState } from './editOps'
import { FmtDialog } from './FmtDialog'
import { cellStyleToCss, getCellNumberFormat, getCellDisplayText } from '@office-kit/xlsx/styles'
import { buildMarks, inAnyTable, displayRef } from './domainMarks'
import { ChapterDialog } from './ChapterDialog'
import { useOutlineSlot, type OutlineSlotConfig } from './useOutlineSlot'
import { asOpenFailure, type OpenFailure } from './openError'
import { OpenFailureNotice } from './OpenFailureNotice'
import { indexToCol, colToIndex } from './a1'
import { translateFormula } from './translate'
import { fillSmartNumber, type NumMatrix } from './fillSeries'
import { Ribbon, type RibbonItem, type RibbonTab } from '../office/Ribbon'
import { RefPanel, RightPanel } from './RefPanel'
import { TriggerPanel } from './TriggerPanel'
import { applyTriggerPlan } from './apply'
import { planTrigger, PRESET_RULES, type TriggerRule, type TriggerPlan } from './trigger'
import { sheetNames as wbSheetNames } from '@office-kit/xlsx/workbook'
import { parseUrn } from '../ref/urn'
import {
  FX_URN_RE, fxUrnRanges, readableTokenRanges, fxTokenRanges, urnToReadable, readableToUrn,
  resolveNodeFields, type ResolvedFields, type FxToken,
} from './fxUrn'
import { computeNode } from '../core/engine'
import { capsuleHtml, capsuleKindClass } from '../ref/capsule'
import type { TableSchema, NodeConfig } from '../core/schema'
import { useBaseStore } from '../core/store'
import type { Row } from '../core/store'

/** 单格坐标判定（#36 防御：公式态拖拽落的是区间引用 'A1:B3'，绝不能进 setSelCoord/parseCoord） */
const isSingleCoord = (s: string | null | undefined): boolean => /^[A-Za-z]{1,3}\d+$/.test(s ?? '')

/** HTML 转义（fx 镜像层非 URN 片段用；胶囊内部转义由 capsuleHtml 自理） */
const escFx = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
/** A1 式坐标（编辑态被引用格淡显/点公式引用亮显）—— WPS 语义（用户 2026-09-27） */
const REF_COORD_RE = /\b[A-Z]{1,3}[1-9]\d*\b/g
/** 转义 + A1 坐标包 data-ref-coord span（镜像层点击命中用；函数调用名 LOG10( / SUM( 等不标） */
const escFxCoords = (text: string): string => {
  if (!text || !/[A-Z]/.test(text)) return escFx(text)
  let out = ''
  let last = 0
  for (const m of text.matchAll(new RegExp(REF_COORD_RE.source, 'g'))) {
    const i = m.index ?? 0
    const tok = m[0]
    if (text[i + tok.length] === '(' || FUNC_NAMES.includes(tok)) continue
    out += escFx(text.slice(last, i)) + `<span data-ref-coord="${tok}">${escFx(tok)}</span>`
    last = i + tok.length
  }
  out += escFx(text.slice(last))
  return out
}
// ★ 批次17：可读串 ⇄ URN 互转层（fxUrnRanges / readableTokenRanges / fxTokenRanges / urnToReadable /
//   readableToUrn / resolveNodeFields / FX_URN_RE）已抽到 ./fxUrn —— 纯计算层、可单测；
//   ★ 口径铁律：字段一律取 computeNode 的**计算后**结果，不是 schema 静态 t.fields（汇聚型节点 t.fields 为空）。
/**
 * token 完整性守卫（胶囊锁死的 onChange 层兜底，用户 2026-09-27：「输入的字符融入胶囊内部」）：
 * onKeyDown 锁死三件套覆盖不了 IME composition / 拖放等绕过键盘的入口 —— 在值入口统一收口。
 * 纯插入且插入点落在旧 token 区间**内部**（非边界）→ 拒绝该次变更。返回 null = 放行；
 * 返回 number = 插入被拒，应把光标推到该位置（token 末尾）。
 */
const tokenInsertGuard = (oldV: string, v: string, ranges: { start: number; end: number }[]): number | null => {
  if (v.length <= oldV.length) return null // 只拦纯插入（替换/删减走键盘层原子吞语义）
  let i = 0
  while (i < oldV.length && i < v.length && oldV[i] === v[i]) i++
  let j = 0
  while (j < oldV.length - i && j < v.length - i && oldV[oldV.length - 1 - j] === v[v.length - 1 - j]) j++
  if (i + j !== oldV.length) return null // 非纯插入
  const hit = ranges.find((r) => i > r.start && i < r.end)
  return hit ? hit.end : null
}
/** 胶囊/坐标 span 的可视 rect：镜像容器 overflow:hidden 会视觉裁掉溢出部分，但 getBoundingClientRect
 *  仍返回完整尺寸 —— 命中几何必须同步裁剪，否则点击被裁区域会误判命中（用户 2026-09-27：点胶囊附近冒出别的字符） */
const clippedRect = (el: HTMLElement, mirror: HTMLElement): DOMRect => {
  const r = el.getBoundingClientRect()
  const m = mirror.getBoundingClientRect()
  const left = Math.max(r.left, m.left)
  const right = Math.min(r.right, m.right)
  return new DOMRect(left, r.top, Math.max(right - left, 0), r.height)
}
/** 点击坐标落在 fx 镜像层哪个胶囊上（胶囊 = 锁死 token，点击弹同节点换参菜单） */
const hitFxMirrorCapsule = (mirror: HTMLElement, cx: number, cy: number) => {
  for (const el of Array.from(mirror.querySelectorAll<HTMLElement>('[data-fx-urn]'))) {
    const r = clippedRect(el, mirror)
    if (r.width > 0 && cx >= r.left && cx <= r.right && cy >= r.top && cy <= r.bottom) {
      return { start: Number(el.dataset.fxStart), end: Number(el.dataset.fxEnd), urn: el.dataset.fxUrn ?? '', left: r.left }
    }
  }
  return null
}
/** 点击坐标落在 fx 镜像层哪个格子引用上（data-ref-coord span）：WPS 语义 —— 点公式里的引用 → 网格对应格亮显。
 *  idx = 命中的是公式里第几个引用实例（与 calcRefInstances 序号一一对应，span 生成顺序一致），供「点引用 = 选中整个实例区间」 */
const hitFxMirrorCoord = (mirror: HTMLElement, cx: number, cy: number) => {
  const els = Array.from(mirror.querySelectorAll<HTMLElement>('[data-ref-coord]'))
  for (let i = 0; i < els.length; i++) {
    const r = clippedRect(els[i], mirror)
    if (r.width > 0 && cx >= r.left && cx <= r.right && cy >= r.top && cy <= r.bottom) {
      return { coord: els[i].dataset.refCoord ?? '', idx: i }
    }
  }
  return null
}
/** 公式串 → A1 引用实例表（含字符区间）：点公式里的引用 = 活跃引用区选中整个实例（打字/点格直接替换，WPS 语义） */
const calcRefInstances = (s: string): { coord: string; s: number; e: number }[] => {
  const out: { coord: string; s: number; e: number }[] = []
  if (!s.startsWith('=')) return out
  for (const m of s.matchAll(new RegExp(REF_COORD_RE.source, 'g'))) {
    const i = m.index ?? 0
    const tok = m[0]
    if (s[i + tok.length] === '(') continue
    if (FUNC_NAMES.includes(tok)) continue
    out.push({ coord: tok, s: i, e: i + tok.length })
  }
  return out
}
/** 点击点之前的可见字符数（几何估算）：TreeWalker 逐字符 Range 比中点 —— caretRangeFromPoint 会跳过
 *  pointer-events:none 的 mirror 穿透到 input（container 不在 mirror 内 → 拿不到 mirror 内偏移），
 *  纯文本区点击/拖选因此拿不到 vis，只能放弃接管 → 原生按透明原文布局定位，穿过胶囊区即错位（实测轨迹 4→5→7→8→10→11）。
 *  公式短（<100 字符）且仅 mousedown/mousemove 拖选时调用，逐字符 createRange 可接受 */
const estimateVisFromX = (mirror: HTMLElement, cx: number): number => {
  const walker = document.createTreeWalker(mirror, NodeFilter.SHOW_TEXT)
  let vis = 0
  let node = walker.nextNode()
  while (node) {
    const t = node.textContent ?? ''
    for (let i = 0; i < t.length; i++) {
      const r = document.createRange()
      try {
        r.setStart(node, i)
        r.setEnd(node, i + 1)
      } catch {
        continue
      }
      const rect = r.getBoundingClientRect()
      if (rect.width === 0 && rect.height === 0) continue
      if (cx < rect.left + rect.width / 2) return vis + i
    }
    vis += t.length
    node = walker.nextNode()
  }
  return vis
}
/**
 * 镜像层点击 → 编辑串光标偏移（备忘 2026-09-27 ②：胶囊态鼠标点不出光标 —— 镜像层吃掉 mousedown，input 不会自行定位）。
 * 可见字符序 → 编辑串偏移：可读 token 的胶囊可见名与 token 原文同长；URN 向下兼容 token 显示名有差，
 * 按 data-fx-start/end 区间长度差修正（点击点之前每个胶囊：editOffset += (end-start) - 可见长）。
 * 备忘 2026-09-27 ⑥-c：caretRangeFromPoint 穿透 mirror（pointer-events:none）→ vis 用 estimateVisFromX 几何估算，
 * 不再返回 null —— 拖选/点击文本区也全量接管（原生拖选按透明原文布局定位，宽差错位实锤）。
 */
const hitFxMirrorOffset = (mirror: HTMLElement, cx: number, cy: number): number | null => {
  const doc = mirror.ownerDocument
  const cr = (doc as Document & { caretRangeFromPoint?: (x: number, y: number) => Range | null }).caretRangeFromPoint?.(cx, cy)
  let vis: number
  if (!cr || !mirror.contains(cr.startContainer)) {
    if (!mirror.textContent) return null
    vis = estimateVisFromX(mirror, cx)
  } else {
    const rng = doc.createRange()
    rng.setStart(mirror, 0)
    try {
      rng.setEnd(cr.startContainer, cr.startOffset)
    } catch {
      return null
    }
    vis = rng.toString().length
  }
  const caps = Array.from(mirror.querySelectorAll<HTMLElement>('[data-fx-start]'))
    .map((cap) => {
      const r2 = doc.createRange()
      r2.setStart(mirror, 0)
      try {
        r2.setEnd(cap, cap.childNodes.length)
      } catch {
        return null
      }
      const s = Number(cap.dataset.fxStart)
      const e = Number(cap.dataset.fxEnd)
      return !Number.isNaN(s) && !Number.isNaN(e) ? { endVis: r2.toString().length, s, e, rect: clippedRect(cap, mirror) } : null
    })
    // 备忘 2026-09-27 ④：width=0（被镜像容器完全裁掉的胶囊）不参与边界钳制 —— 否则容器内任意点击会被误钳到不可见胶囊的 s
    .filter((x): x is { endVis: number; s: number; e: number; rect: DOMRect } => !!x && x.rect.width > 0)
    .sort((a, b) => a.endVis - b.endVis)
  let prevEdit = 0
  let prevVis = 0
  for (const c of caps) {
    // token 边界钳制：镜像层 pointer-events:none，点击穿透到 input 按透明原文布局定位 —— 原文与胶囊宽度不一致
    // 会让光标落进 token 内部（打字即「融入胶囊」）。点击点在胶囊左/右就近钳到 s/e，永入不了 token。
    if (cx < c.rect.left - 0.5) return c.s
    if (cx <= c.rect.right + 0.5) return cx - c.rect.left < c.rect.right - cx ? c.s : c.e
    if (c.endVis <= vis) {
      prevEdit = c.e
      prevVis = c.endVis
    } else break
  }
  return prevEdit + (vis - prevVis)
}
/** 备忘 2026-09-27 ⑥ · 胶囊=一个字符：选区与 token 相交 → 吞入整 token（点选/拖选的原子吸附） */
const tokenAtomicSel = (s: number, e: number, ranges: { start: number; end: number }[]): { s: number; e: number } => {
  let s2 = s
  let e2 = e
  for (const r of ranges) {
    if (s < r.end && e > r.start) { s2 = Math.min(s2, r.start); e2 = Math.max(e2, r.end) }
  }
  return { s: s2, e: e2 }
}
/**
 * 备忘 2026-09-27 ⑥ · 删除守卫（tokenInsertGuard 的删除侧对偶）：token 被「几乎删空但留残段」→ 拒绝。
 * 截图实锤（用户拖选/删除后 fx 泄漏 `://param_…#price`）——URN 段头 `node` 被逐字删掉，残段变普通文本。
 * 残段判定：v 里 token 原文不完整、但「去首 1 字符」或「去尾 1 字符」仍在（渐进连删每步只少 1 字符，必然命中）。
 * 返回 null = 放行；number = 拒绝且光标应推到该 token 末尾（再按一次退格即 keydown 原子删整 token）。
 * 短 token（<5 字符）放行——普通文本误判风险高。
 */
const tokenDeleteGuard = (oldV: string, v: string, ranges: { start: number; end: number }[]): number | null => {
  if (v.length >= oldV.length) return null
  for (const r of ranges) {
    const T = oldV.slice(r.start, r.end)
    if (T.length < 5) continue
    if (v.includes(T)) continue // 完整保留（删除发生在别处）
    if (v.includes(T.slice(1)) || v.includes(T.slice(0, -1))) return r.end
  }
  return null
}
/** 公式串 → A1 式格子引用集合（编辑态被引用格淡显用）：排除函数调用名（LOG10( / SUM( 等）与已知函数名 */
const calcRefCoords = (s: string): Set<string> => {
  const set = new Set<string>()
  if (!s.startsWith('=')) return set
  for (const m of s.matchAll(new RegExp(REF_COORD_RE.source, 'g'))) {
    const tok = m[0]
    const after = s[(m.index ?? 0) + tok.length]
    if (after === '(') continue
    if (FUNC_NAMES.includes(tok)) continue
    set.add(tok)
  }
  return set
}
/** 文本段渲染：切选区高亮（fake-sel）/ 插假光标（fake-caret，严格在段内部）；段边界由 URN 段或首尾处理 */
const renderTextSeg = (
  text: string,
  segStart: number,
  sel: { s: number; e: number } | undefined,
  hasSel: boolean,
  caret: number | null,
): string => {
  if (!text) return ''
  const segEnd = segStart + text.length
  if (hasSel && sel) {
    const a = Math.max(sel.s, segStart)
    const b = Math.min(sel.e, segEnd)
    if (a >= b) return escFxCoords(text)
    return (
      escFxCoords(text.slice(0, a - segStart)) +
      `<span class="fake-sel">${escFxCoords(text.slice(a - segStart, b - segStart))}</span>` +
      escFxCoords(text.slice(b - segStart))
    )
  }
  if (caret !== null && caret > segStart && caret < segEnd) {
    const k = caret - segStart
    return escFxCoords(text.slice(0, k)) + '<span class="fake-caret"></span>' + escFxCoords(text.slice(k))
  }
  return escFxCoords(text)
}

/** 编辑器镜像 HTML 构建（fx 可读化 · 显示态）：token（可读「表名·参数名」或向下兼容的 URN 原文）→ 内联胶囊，
 *  其余文本转义 —— fx 栏与格内编辑器（FCE）共用。**编辑串本身是可读串，URN 不出现在编辑器**（用户 2026-09-27）；
 *  存储层永远是 URN 原文（<f> 公式体），提交边界 readableToUrn 转回。
 *  input 文字与原生 caret 全透明（mirror-on），光标/选区由镜像层绘制：
 *    sel = 输入态 {s,e}（光标/选区偏移）；e>s 时选区相交 token → 整胶囊包 fake-sel、文本段按边界切分；
 *    e===s 时光标画 fake-caret（落 token 内部 → 推到胶囊右缘，与 onSelect 兜底一致）。
 *  点击命中仍走 hitFxMirrorCapsule 几何检测（读 data-fx-*，胶囊锁死 token 唯一修改入口 = 点它换参）。 */
const buildUrnMirrorHtml = (
  s: string,
  tables: ResolvedFields | null | undefined,
  sel?: { s: number; e: number },
) => {
  const hasSel = !!sel && sel.s !== sel.e
  const caret = sel && sel.s === sel.e ? sel.s : null
  const ranges = fxTokenRanges(s, tables)
  if (!ranges.length) {
    // 纯公式态（无胶囊 token）也走镜像：坐标引用包 data-ref-coord span，点它 → 网格对应格亮显（WPS 语义）
    if (!s.startsWith('=')) return ''
    const inner = renderTextSeg(s, 0, sel, hasSel, caret)
    // renderTextSeg 只插段内部 —— 光标在最前/最后（caret===0 / ===s.length）要在这里补桩，否则公式态无光标（用户 2026-09-27）
    if (!hasSel && caret !== null && (caret === 0 || caret === s.length)) {
      const CARET = '<span class="fake-caret"></span>'
      return caret === 0 ? CARET + inner : inner + CARET
    }
    return inner
  }
  const CARET = '<span class="fake-caret"></span>'
  let out = ''
  let last = 0
  let caretDone = false
  // 光标在最前（pos 0 且第一个 token 前没有更早的边界）
  if (caret === 0 && (!ranges[0] || ranges[0].start > 0)) {
    out += CARET
    caretDone = true
  }
  for (const r of ranges) {
    // 光标紧邻本 token 前（尚未插入过）
    if (caret !== null && !caretDone && caret === r.start) {
      out += CARET
      caretDone = true
    }
    // ── 文本段 [last, r.start) ──
    out += renderTextSeg(s.slice(last, r.start), last, sel, hasSel, caret)
    // ── token 段 → 内联胶囊（可读 token 直接用 node/field；URN token 已在 fxTokenRanges parse 好）──
    let cap =
      `<span data-fx-urn="${escFx(r.urn)}" data-fx-start="${r.start}" data-fx-end="${r.end}"` +
      ` style="display:inline-block;vertical-align:middle">` +
      capsuleHtml({
        node: r.node,
        field: r.field,
        kind: r.kind,
        dead: r.dead,
        title: `${r.node}·${r.field}`,
      }) +
      `</span>`
    if (hasSel && sel && sel.s < r.end && sel.e > r.start) cap = `<span class="fake-sel">${cap}</span>`
    if (caret !== null && !hasSel) {
      // 光标在 token 内部/右缘 → 胶囊右缘（onSelect 兜底会把 input 光标推到 r.end，两处语义一致）
      if (caret === r.start && !caretDone) {
        out += CARET
        caretDone = true
      }
      out += cap
      if (caret > r.start && caret <= r.end) {
        out += CARET
        caretDone = true
      }
    } else out += cap
    last = r.end
  }
  // 最后一段：renderTextSeg 只插段内部 —— 光标在整串末尾（caret===s.length 且串尾还有文本）要补桩
  const tail = s.slice(last)
  out += renderTextSeg(tail, last, sel, hasSel, caret)
  if (!hasSel && caret !== null && !caretDone && caret === s.length && tail.length > 0) {
    out += CARET
  }
  return out
}
/** FCE（格内编辑器）是模块顶层组件拿不到 tables —— WorkbookEditor 每渲染同步进来（镜像/换参菜单用）。
 *  ★ 批次17：同步的是 **resolveNodeFields 的结果**（解析后字段），不是 schema —— 否则汇聚型节点
 *  （sort/map/output）的字段枚举不到，格内编辑器画不出胶囊、提交也转不回 URN（同主编辑器口径）。 */
const fceFieldsRef: { current: ResolvedFields | null | undefined } = { current: null }
/** FCE → fx 栏同步桥：格内编辑值变化时同步 fxDraft（fx 栏跟着显示胶囊；fx 编辑态不存在循环） */
let fceChangeHook: ((v: string) => void) | null = null

/**
 * 工作簿编辑器（离线「在软件里打开 Excel」）· v5 —— 完整菜单栏版
 *
 * 定位：**不依赖任何服务端的那条路。**
 *   · 重度、WPS 级编辑 → OnlyOffice（在线，需 Document Server）
 *   · 看内容、改数据、公式编辑/重算、域绑定、行列与格式 → 这里（离线、零服务端、纯前端）
 *
 * 为什么不用 Univer 之类：它的 xlsx 导入导出属付费 Pro 层（2026-09-24 registry 复核：
 * 开源命名空间根本没有 exchange 包，只在 @univerjs-pro）。我们用 office-kit + patchSave
 * 自己做 IO，保真更高。
 *
 * 往返原则：**补丁不重建**。保存走 patchSaveWorkbook —— office-kit 的产物只覆盖它改动过的部件，
 * 图表/绘图/透视/媒体从原文件原样取回（不用它，李渡可研估算表实测会丢图表）。
 *
 * 公式（v5.1）：公式栏**或格子内**输入 = 后 → 立即出常用函数提示、点其他格子插入引用
 * （格内编辑由 FormulaCellEditor 接管，capture 拦截 mousedown 不打断编辑）、373 函数补全、
 * Esc 退出；重算按依赖多轮迭代。**计算权威唯一**：`<v>` 只由我们的 evalFormula 计算（铁律）。
 *
 * 已知限制：插删行列时公式引用**不自动偏移**；shared 公式组写回会碎化（功能等价）；
 * 跨表引用（Sheet!A1）暂不参与重算。
 */

/** 产物时间戳固定 —— 同输入同字节，undo 快照与门禁都需要可复现 */
const FIXED_MTIME = new Date('2026-01-01T00:00:00Z')
/* MAX_CSS 已提升到 index.css（Word 面板的最大化也依赖它，组件内注入会互相不可见） */

/** 键盘直编辑：capture 抢先拦下打印字符（preventDefault+stopPropagation，ag-grid 收不到），经 startEditingCell 注入格内编辑器。
 *  编辑器 mount 前到达的字符会连续到达 —— 累积而非覆盖，否则丢首字符（竞态实测）。 */
let pendingEditChar: string | null = null
/** 属性引用栏 → 格内编辑器的插入桥：FCE 挂载时注册（= 编辑中点左栏参数，插入到光标处） */
let fceInsertHook: ((text: string) => void) | null = null
/** FCE Escape 取消桥：v32.3 customComp 不注入 stopEditing prop（proxy getProps 只有 value/onValueChange/initialValue），
 *  失焦通道 stopEditingWhenCellsLoseFocus 又因 mount 时序被有意关闭 —— FCE 只能经此桥调 gridApi.stopEditing(true) 取消 */
let fceCancelHook: (() => void) | null = null
/** FCE 公式态点格 → 主组件高亮被点格（WPS 同款选中效果）：点格路径调用，编辑态结束由主组件清 */
let fcePickCellHook: ((coord: string) => void) | null = null
/** FCE 存活信号：mount=true / 卸载=false（主组件据此判定「格内编辑中」做被引用格淡显） */
let fceAliveHook: ((alive: boolean) => void) | null = null
/** FCE 点公式里的坐标引用 → 主组件亮显对应格（WPS 语义）；null = 清亮显 */
let fceGlowHook: ((coord: string | null) => void) | null = null
/** 备忘 2026-09-27 ⑤：FCE 悬停 A1 引用 → 主组件 hover 亮显对应格 —— 独立于点击 glowCell（hover 移出不清点击态） */
let fceHoverGlowHook: ((coord: string | null) => void) | null = null
/** #36 FCE 公式态拖拽选区间桥：FCE 内监听经此写主组件 fxDragRangeRef + 刷新高亮；null = 拖拽结束清 */
let fceRangeDragHook: ((rect: { r1: number; c1: number; r2: number; c2: number } | null) => void) | null = null

/** 常用函数（菜单/补全置顶） */
const COMMON_FNS = ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'IF', 'ROUND', 'SUMIF', 'COUNTIF', 'VLOOKUP', 'CONCAT', 'TODAY']

// ── 格内编辑器：= 开头进公式模式（函数补全 + 点其他格子插入引用，Excel 同款交互）──
// 必须定义在模块顶层：cellEditor 引用要稳定，放进组件体 ag-grid 每次渲染都会重建编辑器类型
// v32.3 customComp 机制：ag-grid 不读 ref.getValue()，而是靠 props.onValueChange 回报值！
type FCEHandle = { getValue: () => string; isCancelAfterEnd: () => boolean }
const FormulaCellEditor = forwardRef(function FormulaCellEditor(
  props: { initialValue?: string; urnInit?: string; onValueChange?: (v: string) => void; stopEditing?: (cancel?: boolean) => void },
  ref: ForwardedRef<FCEHandle>,
) {
  // 引用格优先取 urnInit（'=node://…'，cellEditorParams 自定义键）——ag-grid proxy 把
  // initialValue 硬编码成格值（显示值），穿不进来；urnInit 为空走原逻辑（initialValue 可能是数字 → 一律 String）。
  // fx 可读化：编辑态进可读串（表名·参数名），URN 不出现在编辑器（用户 2026-09-27）——fceFieldsRef 父渲染先同步，mount 时已就绪
  const [val, setVal] = useState(() =>
    props.urnInit != null ? urnToReadable(String(props.urnInit), fceFieldsRef.current) : String(props.initialValue == null ? '' : props.initialValue),
  )
  const [compl, setCompl] = useState<{ items: string[]; idx: number; prefix: string } | null>(null)
  const inputRef = useRef<HTMLInputElement>(null)
  /** Escape 取消标记：isCancelAfterEnd 读它丢弃本次编辑（v32.3 customComp 不注入 stopEditing prop 的兜底） */
  const cancelledRef = useRef(false)
  const valRef = useRef(val)
  valRef.current = val
  const formulaMode = val.startsWith('=')
  /** 备忘 · 格内编辑器胶囊镜像层（与 fx 栏同款 —— input 文字透明 + 上层胶囊 + 锁死 token）；fid 供换参菜单高亮当前参数 */
  const [fceMenu, setFceMenu] = useState<{ start: number; end: number; nodeId: string; fid: string; left: number } | null>(null)
  /** WPS 语义 · 活跃引用区：公式态连续点格 = 替换同一地址（不再追加 A1B1C1）；打字/粘贴/光标移动 = 定型解除 */
  const [fcePick, setFcePick] = useState<{ s: number; e: number } | null>(null)
  const fcePickRef = useRef<{ s: number; e: number } | null>(null)
  useEffect(() => {
    if (!fceMenu) return
    const close = (e: MouseEvent) => {
      if ((e.target as HTMLElement | null)?.closest?.('.wb-fx-menu')) return
      setFceMenu(null)
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setFceMenu(null)
    }
    document.addEventListener('mousedown', close, true)
    document.addEventListener('keydown', onKey)
    return () => {
      document.removeEventListener('mousedown', close, true)
      document.removeEventListener('keydown', onKey)
    }
  }, [fceMenu])
  // 草稿一变，菜单区间作废（换参自身触发也无害）
  useEffect(() => {
    setFceMenu(null)
  }, [val])
  // mirror-on 条件 = 含胶囊 token（可读「表名·参数名」或向下兼容的 URN 原文）；公式态（= 开头）也开 —— 坐标引用包 span 供点击亮显（WPS 语义）
  const mirrorOn = val.startsWith('=') || val.includes('node://') || readableTokenRanges(val, fceFieldsRef.current).length > 0
  /** 对照体系 · 输入态光标/选区（selectionchange 跟踪；失焦归 undefined 消假光标 —— 与 fx 栏同款） */
  const [fceCaret, setFceCaret] = useState<{ s: number; e: number } | undefined>(undefined)
  /** 备忘 2026-09-27 ⑤：悬停胶囊 → 画布来源节点亮显（store 瞬态桥）；diff 用 ref + 卸载兜底清 */
  const setFxHoverNode = useBaseStore((s) => s.setFxHoverNode)
  const fceHoverNodeRef = useRef<string | null>(null)
  /** 备忘 2026-09-27 ⑥ · 拖选锚点（同 fx 栏：胶囊=一个字符，拖选全接管防错位） */
  const fceDragRef = useRef<number | null>(null)
  useEffect(() => () => setFxHoverNode(null), [setFxHoverNode])
  useEffect(() => {
    const onSel = () => {
      const el = inputRef.current
      if (!el || document.activeElement !== el) return
      setFceCaret({ s: el.selectionStart ?? 0, e: el.selectionEnd ?? 0 })
    }
    document.addEventListener('selectionchange', onSel)
    return () => document.removeEventListener('selectionchange', onSel)
  }, [])
  /** 对照体系 · URN → 内联胶囊 + 假光标/选区；fcaret null（未聚焦/失焦）→ 不画光标 */
  const mirrorHtml = useMemo(
    () => buildUrnMirrorHtml(val, fceFieldsRef.current, mirrorOn ? fceCaret : undefined),
    [val, mirrorOn, fceCaret],
  )
  /** 同时更新本地 state 与 ag-grid proxy 的值（只 setVal 的话 getValue 永远是旧值）；同步 fx 栏草稿（用户 2026-09-26：格内插引用 fx 要跟着显示）
   *  keepPick=true = 点格插引用路径（活跃引用区由调用方接管，不清）；打字/粘贴/退格路径默认清（引用定型） */
  const setBoth = (v: string, keepPick = false) => {
    // token 完整性守卫：IME/拖放等绕过 onKeyDown 锁死的插入落进胶囊内部 → 拒绝并把光标推出 token（不清 pick）
    const bad = tokenInsertGuard(valRef.current, v, fxTokenRanges(valRef.current, fceFieldsRef.current))
    if (bad != null) {
      requestAnimationFrame(() => inputRef.current?.setSelectionRange(bad, bad))
      return
    }
    // 删除守卫（同 fx 栏）：token 渐进删空留残段 → 拒绝，光标推 token 末尾
    const badDel = tokenDeleteGuard(valRef.current, v, fxTokenRanges(valRef.current, fceFieldsRef.current))
    if (badDel != null) {
      requestAnimationFrame(() => inputRef.current?.setSelectionRange(badDel, badDel))
      return
    }
    setVal(v)
    props.onValueChange?.(v)
    fceChangeHook?.(v)
    if (!keepPick) {
      fcePickRef.current = null
      setFcePick(null)
    }
  }

  useEffect(() => {
    inputRef.current?.focus()
    inputRef.current?.select()
    fceAliveHook?.(true) // 主组件淡显判定需要「格内编辑中」信号（WPS 语义）
    // 补焦：mount 时 React/ag-grid 渲染竞争会把焦点抢回 BODY（实测 activeElement=BODY），键盘事件全丢 ——
    // 两帧后再补一次，只补焦点不动光标（select 语义保留）
    const raf1 = requestAnimationFrame(() => {
      if (document.activeElement !== inputRef.current) inputRef.current?.focus()
      requestAnimationFrame(() => {
        if (document.activeElement !== inputRef.current) inputRef.current?.focus()
      })
    })
    // ① 键盘直编辑：capture 层拦下的首字符在这里注入（覆盖 initialValue）
    if (pendingEditChar) {
      const ch = pendingEditChar
      pendingEditChar = null
      setBoth(ch)
      updateLocalCompl(ch) // 键盘直编辑不走 input onChange —— 补全触发在这里补上（用户 2026-09-26：格内打 = 要出函数清单）
      if (inputRef.current) inputRef.current.value = ch // 同步 DOM，防受控渲染竞态
    }
    // ② 注册插入桥：属性引用栏在「= 编辑中」点参数时，从这里插入到光标处
    fceInsertHook = (text: string) => {
      const el = inputRef.current
      // 选区吸附：选区含半截 token → 吞整 token 再替换（同 fx 栏，防 URN 残段）
      const fs0 = el ? el.selectionStart ?? valRef.current.length : valRef.current.length
      const fe0 = el ? el.selectionEnd ?? fs0 : fs0
      const { s: sAt, e: eDel } = tokenAtomicSel(fs0, fe0, fxTokenRanges(valRef.current, fceFieldsRef.current))
      // 备忘 2026-09-27 ⑥-b：插入点前是 token 尾字符 → 自动补 +（防紧贴提交产生 URN 粘连，同 fx 栏 pickRef）
      const prev = valRef.current[sAt - 1] ?? ''
      const plus = /[\w\u4e00-\u9fff#·-]$/.test(prev) ? '+' : ''
      const next = valRef.current.slice(0, sAt) + plus + text + valRef.current.slice(eDel)
      setBoth(next)
      if (el) {
        el.value = next // 同步 DOM，防受控渲染竞态吞字符
        el.dispatchEvent(new Event('input', { bubbles: true }))
        const pe = sAt + plus.length + text.length
        el.setSelectionRange(pe, pe)
      }
    }
    return () => {
      fceInsertHook = null
      fceAliveHook?.(false) // 卸载 = 格内编辑结束，淡显判定解除
      cancelAnimationFrame(raf1)
    }
  }, [])

  // 补全规则与公式栏一致：字母前缀过滤；=/运算符后直接给常用函数
  const updateLocalCompl = (v: string) => {
    if (!v.startsWith('=')) return setCompl(null)
    const m = /([A-Za-z][A-Za-z0-9.]*)$/.exec(v)
    if (!m) {
      setCompl(/=$|[+\-*/(,]$/.test(v) ? { items: COMMON_FNS.slice(0, 8), idx: 0, prefix: '' } : null)
      return
    }
    const prefix = m[1].toUpperCase()
    const common = COMMON_FNS.filter((f) => f.startsWith(prefix))
    const rest = FUNC_NAMES.filter((f) => f.startsWith(prefix) && !common.includes(f))
    const items = [...common, ...rest].slice(0, 8)
    setCompl(items.length ? { items, idx: 0, prefix: m[1] } : null)
  }

  // 公式态点其他格子 = 插入引用：capture 阶段拦截 mousedown，
  // preventDefault 不让焦点走（编辑不被打断），stopPropagation 不让 ag-grid 结束编辑/改选中。
  // 监听一次挂载、回调实时查 ref（effect 随输入重挂 document 监听会干扰 ag-grid click 判定）
  // #36 拖拽选区间：mousedown 记锚点+高亮（经 fceRangeDragHook 写主组件 fxDragRangeRef），
  // mouseup 才落引用 —— 拖过 = 区间 A1:B3，单击 = 单格（WPS 活跃引用区语义不变）
  const formulaModeRef = useRef(false)
  formulaModeRef.current = formulaMode
  const fceRangeDragRef = useRef<{ r1: number; c1: number; r2: number; c2: number } | null>(null)
  useEffect(() => {
    const hitCell = (e: MouseEvent): { r: number; c: number; cell: HTMLElement } | null => {
      const t = e.target as HTMLElement | null
      const cell = t?.closest?.('.ag-cell') as HTMLElement | null
      if (!cell) return null
      const col = cell.getAttribute('col-id')
      const rowId = cell.closest('.ag-row')?.getAttribute('row-id')
      if (!col || !/^[A-Z]{1,3}$/.test(col) || !rowId || !/^\d+$/.test(rowId)) return null
      return { r: Number(rowId), c: colToIndex(col), cell }
    }
    const onDocMouseDown = (e: MouseEvent) => {
      if (!formulaModeRef.current) return
      const hit = hitCell(e)
      if (!hit) return
      if (inputRef.current && hit.cell.contains(inputRef.current)) return // 点的是编辑中的格子自己
      e.preventDefault()
      e.stopPropagation()
      const rect = { r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c }
      fceRangeDragRef.current = rect
      fceRangeDragHook?.(rect)
    }
    const onDocMouseMove = (e: MouseEvent) => {
      const drag = fceRangeDragRef.current
      if (!drag) return
      const hit = hitCell(e)
      if (!hit) return
      const r1 = Math.min(drag.r1, hit.r)
      const r2 = Math.max(drag.r1, hit.r)
      const c1 = Math.min(drag.c1, hit.c)
      const c2 = Math.max(drag.c1, hit.c)
      if (r1 === drag.r1 && r2 === drag.r2 && c1 === drag.c1 && c2 === drag.c2) return
      const rect = { r1, c1, r2, c2 }
      fceRangeDragRef.current = rect
      fceRangeDragHook?.(rect)
    }
    const onDocMouseUp = () => {
      const drag = fceRangeDragRef.current
      if (!drag) return
      fceRangeDragRef.current = null
      fceRangeDragHook?.(null) // 高亮清
      const cellRef =
        drag.r1 === drag.r2 && drag.c1 === drag.c2
          ? `${indexToCol(drag.c1)}${drag.r1}`
          : `${indexToCol(drag.c1)}${drag.r1}:${indexToCol(drag.c2)}${drag.r2}`
      const el = inputRef.current
      // WPS 语义（用户 2026-09-27）：公式态连续点格 = 替换上一个活跃引用（不再追加 A1B1C1）；首次点格在光标处插入
      const pick = fcePickRef.current
      const s = pick ? pick.s : el ? el.selectionStart ?? valRef.current.length : valRef.current.length
      const eDel = pick ? pick.e : s
      setBoth(valRef.current.slice(0, s) + cellRef + valRef.current.slice(eDel), true)
      const pe = s + cellRef.length
      fcePickRef.current = { s, e: pe }
      setFcePick({ s, e: pe })
      fcePickCellHook?.(cellRef) // 被点格高亮（WPS 选中效果，主组件 state 驱动；区间串匹配不到单格 = 无高亮，无害）
      requestAnimationFrame(() => {
        el?.focus()
        el?.setSelectionRange(pe, pe)
      })
    }
    document.addEventListener('mousedown', onDocMouseDown, true)
    document.addEventListener('mousemove', onDocMouseMove, true)
    document.addEventListener('mouseup', onDocMouseUp, true)
    return () => {
      document.removeEventListener('mousedown', onDocMouseDown, true)
      document.removeEventListener('mousemove', onDocMouseMove, true)
      document.removeEventListener('mouseup', onDocMouseUp, true)
    }
  }, [])

  useImperativeHandle(
    ref,
    () => ({
      getValue: () => valRef.current,
      // v32.3 customComp 官方取消扩展点：proxy getOptionalMethods 会把它搬到自身，
      // stopEditing 后 ag-grid 调它，true = 丢弃本次编辑值（格保持原值）
      isCancelAfterEnd: () => cancelledRef.current,
    }),
    [],
  )

  // Escape 兜底取消：实测 FCE mount 后焦点会被抢回 BODY（时序在别处），input 的 onKeyDown 收不到 Escape
  // —— 挂 document capture 监听，焦点不在 input 时直接走取消桥（stopEditing(true) 丢弃新值）
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      // IME 组合中的 Escape = 取消拼音组合，不是取消编辑（前置判断避免 key 字面量收窄冲突）
      if (e.isComposing || e.key === 'Process' || e.keyCode === 229) return
      if (e.key !== 'Escape') return
      if (document.activeElement === inputRef.current) return // 焦点在 input，onKeyDown 已处理
      e.preventDefault()
      e.stopPropagation()
      cancelledRef.current = true
      fceCancelHook?.()
    }
    // BISECT 解除：监听恢复（段③ Escape 取消闭环已验证）；失焦自动关由显式 stopEditingWhenCellsLoseFocus:false 兜住
    document.addEventListener('keydown', onKey, true)
    return () => document.removeEventListener('keydown', onKey, true)
  }, [])

  return (
    <div style={{ position: 'relative', width: '100%', height: '100%' }} className="wb-celleditor">
      <input
        ref={inputRef}
        spellCheck={false}
        data-mirror={mirrorOn ? 'on' : 'off'}
        value={val}
        onChange={(e) => {
          setBoth(e.target.value)
          updateLocalCompl(e.target.value)
        }}
        onBlur={() => {
          setFceCaret(undefined) // 对照体系：失焦消假光标（否则镜像层稳态常显）
          setTimeout(() => {
            if (!inputRef.current?.matches(':focus')) setCompl(null)
          }, 120)
        }}
        onMouseDown={(e) => {
          // 备忘 2026-09-27 ⑥ · 胶囊=一个字符（同 fx 栏）：点胶囊 = 选中整 token；点坐标引用 = 整实例选中；
          // 其他 = 光标定位。拖选全接管（锚点 fceDragRef + mousemove 镜像几何），换参菜单仍在 onDoubleClick
          const mirror = e.currentTarget.parentElement?.querySelector('.wb-celleditor-mirror')
          if (!mirror || !mirrorOn) return
          const m = mirror as HTMLElement
          const el = e.currentTarget as HTMLInputElement
          // WPS 语义：点公式里的坐标引用 → 网格对应格亮显 + 整实例区间选中；点其他位置 → 清亮显
          const hc = hitFxMirrorCoord(m, e.clientX, e.clientY)
          fceGlowHook?.(hc?.coord ?? null)
          // 引用实例先于 off 计算（同 fx 栏：span 上 caretRangeFromPoint 穿透 → off=null，整实例选中不能依赖 off）
          const inst = hc ? calcRefInstances(valRef.current)[hc.idx] : null
          const cap = hitFxMirrorCapsule(m, e.clientX, e.clientY)
          const off = hitFxMirrorOffset(m, e.clientX, e.clientY)
          if (off == null && !inst && !cap) return
          e.preventDefault()
          el.focus()
          if (inst) {
            // 点中引用实例 → 选中整个实例区间（同 fx 栏 WPS 语义）
            el.setSelectionRange(inst.s, inst.e)
            fcePickRef.current = { s: inst.s, e: inst.e }
            setFcePick({ s: inst.s, e: inst.e })
            fcePickCellHook?.(inst.coord)
            fceDragRef.current = null
            return
          }
          if (cap) {
            // 点胶囊 → 选中整 token；锚点记就近边界（拖走时 tokenAtomicSel 吸附整吞）
            const off2 = off ?? cap.start
            el.setSelectionRange(cap.start, cap.end)
            fceDragRef.current = off2
          } else if (off != null) {
            el.setSelectionRange(off, off)
            fceDragRef.current = off
          }
          document.addEventListener('mouseup', () => { fceDragRef.current = null }, { once: true })
        }}
        onDoubleClick={(e) => {
          // 批次 8-b（单击迁移）：双击胶囊 → 弹同节点换参菜单（格内编辑胶囊锁死，同 fx 栏方案）
          const mirror = e.currentTarget.parentElement?.querySelector('.wb-celleditor-mirror')
          if (!mirror || !mirrorOn) return
          const hit = hitFxMirrorCapsule(mirror as HTMLElement, e.clientX, e.clientY)
          if (!hit) return
          e.preventDefault()
          const u = parseUrn(hit.urn)
          const t = u?.scheme === 'node' ? fceFieldsRef.current?.[u.target] : null
          if (!t || !t.fields.length) return
          const cr = (e.currentTarget as HTMLElement).closest('.wb-celleditor')?.getBoundingClientRect()
          setFceMenu({ start: hit.start, end: hit.end, nodeId: u!.target, fid: u?.locator ?? '', left: cr ? Math.min(Math.max(hit.left - cr.left - 60, 0), Math.max(cr.width - 190, 0)) : 0 })
        }}
        onMouseMove={(e) => {
          const mirror = e.currentTarget.parentElement?.querySelector('.wb-celleditor-mirror')
          if (!mirror || !mirrorOn) return
          const m = mirror as HTMLElement
          const el = e.currentTarget as HTMLInputElement
          // 备忘 2026-09-27 ⑥ · 拖选接管（同 fx 栏）：起点到当前点 tokenAtomicSel，几何全走镜像坐标防错位
          const drag = fceDragRef.current
          if (drag != null) {
            const off = hitFxMirrorOffset(m, e.clientX, e.clientY)
            if (off != null) {
              const a = Math.min(drag, off)
              const b = Math.max(drag, off)
              const { s: s3, e: e3 } = tokenAtomicSel(a, b, fxTokenRanges(val, fceFieldsRef.current))
              if (el.selectionStart !== s3 || el.selectionEnd !== e3) el.setSelectionRange(s3, e3)
            }
          }
          const hit = hitFxMirrorCapsule(m, e.clientX, e.clientY)
          const cur = hit ? 'pointer' : 'text'
          if (el.style.cursor !== cur) el.style.cursor = cur
          // 备忘 2026-09-27 ⑤：悬停 A1 引用 → 对应格 hover 亮显（fceHoverGlowHook 独立桥，不覆盖点击 glow）；
          // 悬停胶囊 → 画布来源节点边框亮显（store 桥）。URM 格扫描仅 fx 栏做（格内编辑器保持轻量）。
          const hc = hit ? null : hitFxMirrorCoord(m, e.clientX, e.clientY)
          fceHoverGlowHook?.(hc?.coord ?? null)
          const u = hit ? parseUrn(hit.urn) : null
          const nid = u?.scheme === 'node' ? u.target : null
          if (fceHoverNodeRef.current !== nid) {
            fceHoverNodeRef.current = nid
            setFxHoverNode(nid)
          }
        }}
        onMouseLeave={() => {
          // 悬停离开 → 清 hover 亮显与画布桥（点击态 fceGlowHook 不受影响）
          fceHoverGlowHook?.(null)
          if (fceHoverNodeRef.current != null) {
            fceHoverNodeRef.current = null
            setFxHoverNode(null)
          }
        }}
        onPaste={(e) => {
          // 粘贴选区触到 token → 相交 token 整体吞入再插入（原子吞语义）；粘贴文本里的 URN 顺手转可读（URN 不进编辑器）
          const el = e.currentTarget as HTMLInputElement
          const fs = el.selectionStart ?? 0
          const fe = el.selectionEnd ?? fs
          const ranges = fxTokenRanges(val, fceFieldsRef.current)
          const txt0 = e.clipboardData.getData('text')
          const txt = urnToReadable(txt0, fceFieldsRef.current)
          const hitTok = ranges.some((r) => fs < r.end && fe > r.start)
          if (!hitTok && txt === txt0) return
          e.preventDefault()
          let s2 = fs
          let e2 = fe
          for (const r of ranges) {
            if (fs < r.end && fe > r.start) { s2 = Math.min(s2, r.start); e2 = Math.max(e2, r.end) }
          }
          setBoth(val.slice(0, s2) + txt + val.slice(e2))
          const p0 = s2 + txt.length
          requestAnimationFrame(() => { el.focus(); el.setSelectionRange(p0, p0) })
        }}
        onSelect={() => {
          // 兜底：无选区时光标落进 token 内部 → 推到胶囊右缘
          const el = inputRef.current
          if (!el) return
          const s0 = el.selectionStart ?? 0
          const e0 = el.selectionEnd ?? 0
          if (s0 === e0) {
            const r = fxTokenRanges(el.value, fceFieldsRef.current).find((q) => s0 > q.start && s0 < q.end)
            if (r) el.setSelectionRange(r.end, r.end)
          } else {
            // 备忘 2026-09-27 ⑥ · 选区吸附（同 fx 栏）：shift 扩选等路径选区相交 token → 吞整 token（幂等无死循环）
            const rr = fxTokenRanges(el.value, fceFieldsRef.current)
            if (rr.some((q) => s0 < q.end && e0 > q.start)) {
              const { s: s3, e: e3 } = tokenAtomicSel(s0, e0, rr)
              if (s3 !== s0 || e3 !== e0) el.setSelectionRange(s3, e3)
            }
          }
          // WPS 语义：光标移出活跃引用区末端 = 引用定型（下次点格重新插入；点格桥会先把光标摆到区末端，不触发清）
          const pk = fcePickRef.current
          if (pk && (s0 !== pk.e || e0 !== pk.e)) {
            fcePickRef.current = null
            setFcePick(null)
          }
        }}
        onKeyDown={(e) => {
          // IME 组合中（拼音/选词/候选导航）放行给输入法：回车=上屏、Escape=取消组合、上下键=候选导航（用户 2026-09-27：中文输公式回车被误当提交）
          if (e.nativeEvent.isComposing || e.key === 'Process' || e.keyCode === 229) return
          // ── 胶囊锁死（与 fx 栏同款：内部打字锁 / 选区相交原子吞 / 退格原子删 / 方向键整跳）──
          const el = e.currentTarget as HTMLInputElement
          const fs = el.selectionStart ?? 0
          const fe = el.selectionEnd ?? fs
          const ranges = fxTokenRanges(val, fceFieldsRef.current)
          const printable = e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey
          if (printable && fe > fs && ranges.some((r) => fs < r.end && fe > r.start)) {
            e.preventDefault()
            let s2 = fs
            let e2 = fe
            for (const r of ranges) {
              if (fs < r.end && fe > r.start) { s2 = Math.min(s2, r.start); e2 = Math.max(e2, r.end) }
            }
            setBoth(val.slice(0, s2) + e.key + val.slice(e2))
            const p0 = s2 + 1
            requestAnimationFrame(() => { el.focus(); el.setSelectionRange(p0, p0) })
            return
          }
          if (printable && fs === fe) {
            const inside = ranges.find((r) => fs > r.start && fs < r.end)
            if (inside) { e.preventDefault(); return }
          }
          // 备忘 2026-09-27 ⑥ · 选区删除原子吞（同 fx 栏）：选区相交 token → 吞整 token 再删（原生删留残段）
          if ((e.key === 'Backspace' || e.key === 'Delete') && fe > fs && ranges.some((r) => fs < r.end && fe > r.start)) {
            e.preventDefault()
            const { s: s2, e: e2 } = tokenAtomicSel(fs, fe, ranges)
            setBoth(val.slice(0, s2) + val.slice(e2))
            requestAnimationFrame(() => { el.focus(); el.setSelectionRange(s2, s2) })
            return
          }
          if ((e.key === 'Backspace' || e.key === 'Delete') && fs === fe) {
            const r = ranges.find((q) => (e.key === 'Backspace' ? fs > q.start && fs <= q.end : fs >= q.start && fs < q.end))
            if (r) {
              e.preventDefault()
              setBoth(val.slice(0, r.start) + val.slice(r.end))
              const p0 = r.start
              requestAnimationFrame(() => { el.focus(); el.setSelectionRange(p0, p0) })
              return
            }
          }
          if ((e.key === 'ArrowLeft' || e.key === 'ArrowRight') && fs === fe) {
            const r = ranges.find((q) => (e.key === 'ArrowLeft' ? fs > q.start && fs <= q.end : fs >= q.start && fs < q.end))
            if (r) {
              e.preventDefault()
              const p0 = e.key === 'ArrowLeft' ? r.start : r.end
              el.setSelectionRange(p0, p0)
              return
            }
          }
          if (fceMenu && e.key === 'Escape') { e.preventDefault(); setFceMenu(null); return }
          if (compl) {
            if (e.key === 'ArrowDown') { e.preventDefault(); setCompl({ ...compl, idx: (compl.idx + 1) % compl.items.length }); return }
            if (e.key === 'ArrowUp') { e.preventDefault(); setCompl({ ...compl, idx: (compl.idx - 1 + compl.items.length) % compl.items.length }); return }
            if (e.key === 'Tab') { e.preventDefault(); acceptLocalCompl(); return }
            if (e.key === 'Escape') { e.preventDefault(); setCompl(null); return }
          }
          if (e.key === 'Enter') { e.preventDefault(); props.stopEditing?.(false) }
          else if (e.key === 'Escape') {
            e.preventDefault()
            cancelledRef.current = true
            // 官方通道优先；v32.3 customComp 无 stopEditing prop → 桥调 gridApi.stopEditing(true)（cancel 语义：丢弃新值）
            if (props.stopEditing) { props.stopEditing(true); return }
            fceCancelHook?.()
          }
        }}
        style={{
          width: '100%', height: '100%', boxSizing: 'border-box', fontSize: 12.5,
          fontFamily: 'Consolas, monospace', color: mirrorOn ? 'transparent' : 'var(--ed-text)', caretColor: mirrorOn ? 'transparent' : 'var(--ed-text)',
          background: 'var(--ed-bg2)', border: '1px solid var(--ed-gold)', outline: 'none', padding: '0 4px', lineHeight: '20px',
        }}
        className="wb-celleditor-input"
      />
      {mirrorOn && (
        <div
          className="wb-celleditor-mirror"
          style={{
            position: 'absolute', inset: 0, padding: '0 5px', pointerEvents: 'none', overflow: 'hidden',
            display: 'flex', alignItems: 'center', whiteSpace: 'pre', fontSize: 12.5, fontFamily: 'Consolas, monospace',
          }}
          dangerouslySetInnerHTML={{ __html: mirrorHtml }}
        />
      )}
      {fceMenu && (() => {
        const t = fceFieldsRef.current?.[fceMenu.nodeId]
        if (!t) return null
        return (
          <div className="ref-chip-menu wb-fx-menu" style={{ top: '100%', bottom: 'auto', left: fceMenu.left, right: 'auto' }}>
            {t.fields.map((f) => (
              <button
                key={f.id}
                className={`ref-chip-item${f.id === fceMenu.fid ? ' on' : ''}`}
                onClick={() => {
                  // fx 可读化：换参 = 把 token 替换成「表名·新参数名」（编辑串全程可读，URN 只在提交边界）
                  const tok = `${t.name}·${f.name}`
                  const nv = val.slice(0, fceMenu.start) + tok + val.slice(fceMenu.end)
                  setFceMenu(null)
                  setBoth(nv)
                  const p = fceMenu.start + tok.length
                  requestAnimationFrame(() => { const el = inputRef.current; el?.focus(); el?.setSelectionRange(p, p) })
                }}
              >
                {f.name}
              </button>
            ))}
          </div>
        )
      })()}
      {compl && (
        <div
          style={{
            position: 'absolute', top: '100%', left: 0, zIndex: 80, minWidth: 190,
            background: 'var(--ed-bg2)', border: '1px solid var(--ed-border2)', borderRadius: 6,
            maxHeight: 200, overflowY: 'auto', boxShadow: '0 10px 28px rgba(0,0,0,.55)',
            // 纯提示层（与公式栏补全一致）：不拦鼠标，点格子直达插引用；键盘 ↑↓+Tab 插入
            pointerEvents: 'none',
          }}
        >
          <div style={{ padding: '3px 10px 4px', fontSize: 10, color: 'var(--ed-text3)', borderBottom: '1px solid var(--ed-border2)' }}>
            ↑↓ 选择 · Tab 插入
          </div>
          {compl.items.map((f, i) => (
            <div
              key={f}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => acceptLocalCompl()}
              onMouseEnter={() => setCompl({ ...compl, idx: i })}
              style={{
                padding: '4px 10px', fontSize: 12, fontFamily: 'Consolas, monospace',
                color: i === compl.idx ? 'var(--ed-bg3)' : 'var(--ed-text)',
                background: i === compl.idx ? 'var(--ed-accent)' : 'transparent', cursor: 'pointer',
              }}
            >
              {f}(…)
            </div>
          ))}
        </div>
      )}
    </div>
  )

  function acceptLocalCompl() {
    if (!compl) return
    const fn = compl.items[compl.idx]
    const el = inputRef.current
    const caret = el ? el.selectionStart ?? valRef.current.length : valRef.current.length
    const before = valRef.current.slice(0, caret - compl.prefix.length)
    const after = valRef.current.slice(caret)
    const next = before + fn + '()' + after
    setBoth(next)
    setCompl(null)
    requestAnimationFrame(() => {
      const p = before.length + fn.length + 1
      el?.focus()
      el?.setSelectionRange(p, p)
    })
  }
})

interface WBState {
  wb: Workbook
  sheets: string[]
  /** 打开时归一化改过什么（如实告知，不静默） */
  applied: string[]
}

export function WorkbookEditor({
  bytes,
  name,
  onChange,
  onSave,
  readOnly,
  theme = 'dark',
  tables,
  records,
  onSetNodeConfig,
  nodeId,
  maximized = false,
  onToggleMax,
  outline,
}: {
  bytes: Uint8Array
  name: string
  /** 每次编辑后回传新字节 */
  onChange: (next: Uint8Array) => void
  /** 另存为 */
  onSave?: (bytes: Uint8Array) => void
  readOnly?: boolean
  /** 属性引用栏所需：画布全部节点 / 数据 / 登记节点关系（登记后画布自动出「数据流」连线） */
  tables?: Record<string, TableSchema>
  records?: Record<string, Row[]>
  onSetNodeConfig?: (tableId: string, patch: Partial<NodeConfig>) => void
  /** 本编辑器对应的节点 id（引用登记用） */
  nodeId?: string
  /** 最大化态（NodeEditor 持有）——属性引用栏的定位随之切换：普通=窗体外圆角卡片；最大化=嵌入窗体左缘 */
  maximized?: boolean
  /** 备忘 · 切换最大化（两栏都退出时自动最大化用） */
  onToggleMax?: () => void
  /** 编辑器主题（亮暗切换时 ag-grid 主题 class 跟着换） */
  theme?: 'dark' | 'light'
  /**
   * ★ 章节编排插槽（2026-10-05 新增，架构关键）
   *
   * 用户要求「基于当前全量功能定制，且差异化内容在 oo 系列更新后能最小代价适配」。
   * 实现方式就是这**一个可选 props**：
   *   · 不传 → 本组件内所有章节相关代码都不执行，行为与改动前逐字节一致（零回归）
   *   · 传了 → 多一个「章节编排」Ribbon 按钮 + 挂ChapterDialog，其余功能完全复用
   *
   * 所有章节逻辑住在 outline.ts / outlineAdapter.ts / ChapterDialog.tsx 三个独立文件，
   * 本文件只留「取行 → 交给窗体 → 收结构变更 → 写回 → commit」这一条细线。
   * oo 系列升级时冲突面仅限这条线，不涉及章节算法本身。
   */
  outline?: boolean | OutlineSlotConfig
}) {
  const [state, setState] = useState<WBState | null>(null)
  /**
   * 打开失败 vs 操作出错，刻意分开：
   *   打开失败 → 整页替代（没东西可显示），要给人话 + 记样本
   *   操作出错 → 顶部一条提示（网格还在，用户不该因为一次改格子出错就丢掉整张表）
   */
  const [openFail, setOpenFail] = useState<OpenFailure | null>(null)
  const [err, setErr] = useState('')
  const [busy, setBusy] = useState('')
  const [active, setActive] = useState('')
  /** 备忘 · 切换工作表即持久化（下次打开恢复现场） */
  useEffect(() => {
    if (!active || !nodeId) return
    try {
      localStorage.setItem(`wb-sheet:${nodeId}`, active)
    } catch {
      /* 存不了就算了 */
    }
  }, [active, nodeId])
  const [dirty, setDirty] = useState(0)
  const [fields, setFields] = useState<XlsxFieldSpec[]>([])
  const [showDomains, setShowDomains] = useState(false)
  /**
   * 联动触发规则（2026-10-05 ·「更新时更新」口径）
   * 按节点持久化到 localStorage：规则是用户配的资产，重开编辑器丢了等于白配。
   * 存 key 沿用 refpanel:<nodeId> 的同款命名法。
   */
  const [triggerOpen, setTriggerOpen] = useState(false)
  const [trgRules, setTrgRules] = useState<TriggerRule[]>(() => {
    try {
      const k = `trigrules:${nodeId}`
      const raw = localStorage.getItem(k)
      if (raw) return JSON.parse(raw) as TriggerRule[]
      return PRESET_RULES.map((r, i) => ({ ...r, id: 'pre' + i }))
    } catch {
      return PRESET_RULES.map((r, i) => ({ ...r, id: 'pre' + i }))
    }
  })
  useEffect(() => {
    try {
      localStorage.setItem(`trigrules:${nodeId}`, JSON.stringify(trgRules))
    } catch {
      /* 配额满/隐私模式：规则只在内存里活，不阻断编辑 */
    }
  }, [trgRules, nodeId])
  /** 选中格（公式栏跟随它）；Shift+点击 = 以先前锚点扩选 */
  const [selCoord, setSelCoord] = useState('')
  const [selAnchor, setSelAnchor] = useState<{ r: number; c: number } | null>(null)
  /** 公式栏草稿 */
  const [fxDraft, setFxDraft] = useState('')
  /** ★ 批次17：可读 token / URN 互转的**唯一**字段来源（与 RefPanel、trgSources 同口径 = computeNode 结果）。
   *  存在意义见 resolveNodeFields 注释：汇聚型节点（sort/map/output/docx/xlsx）自身 t.fields 为空，
   *  旧代码拿 t.fields 枚举 → 点参数「插得进、画不出胶囊、提交也转不回 URN」。
   *  位置在 fceFieldsRef 注入之前 —— FCE 是模块顶层组件，必须在本次渲染内先算出结果才能同步进去。 */
  const resolvedFields = useMemo(
    () => resolveNodeFields(tables, records, nodeId),
    [tables, records, nodeId],
  )
  /** fx 可读化用：effect/callback 里实时取 resolvedFields 而不进依赖（引用变化不重置编辑中草稿） */
  const resolvedFieldsRef = useRef(resolvedFields)
  resolvedFieldsRef.current = resolvedFields
  // 格内编辑器（模块顶层组件）的字段注入 + 值同步桥：格内编辑时 fx 栏跟着显示（含胶囊镜像）
  fceFieldsRef.current = resolvedFields
  /** fx 可读化用：effect/callback 里实时取 tables 而不进依赖（tables 引用变化不重置编辑中草稿） */
  const tablesRef = useRef(tables)
  tablesRef.current = tables

  /** node:// URN → 上游首行字段值（批次 1-6 · 引用参与求值）。
   *  与显示层同语义：records[目标节点] 首行、按字段 id/name 匹配；取不到返回 undefined
   *  （求值侧保留 URN 原文 → 解析报错 #VALUE!，与坏公式同路）。ref 中转：求值调用点
   *  散落在 useCallback/effect 里，避免依赖数组地狱。 */
  const urnResolverRef = useRef<(urn: string) => unknown>(() => undefined)
  urnResolverRef.current = (urn: string): unknown => {
    const u = parseUrn(urn)
    if (u?.scheme === 'const') return u.target
    if (u?.scheme === 'param') {
      // 全局参数：优先取画布上任意节点的同名参数字段（用户心智是"参数名"，不是 nodeId）
      // ★ 批次17：同样走 resolvedFields（计算后字段）—— 汇聚型节点的字段不在 t.fields 里
      for (const [nid, t] of Object.entries(resolvedFieldsRef.current)) {
        const f = t.fields.find((x) => x.name === u.target || x.id === u.target)
        if (f) {
          const v = records?.[nid]?.[0]?.[f.id as keyof Row]
          if (v != null) return v
        }
      }
      return undefined
    }
    if (u?.scheme !== 'node' || !tables || !records) return undefined
    // ★ 批次17：字段与取值都走 **计算后**结果（computeNode），与 RefPanel 值预览 / trgSources 同口径。
    // 旧代码用 tables[target].fields + records[target]：对 sort/map/output/docx/xlsx 这类汇聚型节点，
    // 自身 schema fields 为空、records 里也没有它的行（数据全在 computeNode 的 merge 结果里）
    // → 引用求值一律 undefined → 公式算出 #VALUE!，联动也永远命中不了。
    // 拿不到（缺上游/环）→ 返回 undefined，与旧行为同路（求值侧保留 URN 原文 → #VALUE!）。
    let flds: { id: string; name: string }[] = []
    let rows: Row[] = []
    try {
      const c = computeNode(u.target, tables, records)
      flds = c.fields.map((f) => ({ id: f.id, name: f.name }))
      rows = c.rows
    } catch {
      return undefined
    }
    const fld = u.locator ? flds.find((f) => f.id === u.locator || f.name === u.locator) : null
    const val = rows[0]?.[(fld?.id ?? u.locator) as keyof Row]
    return val == null ? undefined : val
  }

  /**
   * 联动规则可选的来源清单：上游节点的每个字段 → 一个 node:// URN。
   * 与 RefPanel 同款枚举口径（computeNode + 首行值），保证面板里看到的值与实际写入的一致。
   */
  const trgSources = useMemo(() => {
    const out: { urn: string; label: string }[] = []
    if (!tables || !records || !nodeId) return out
    for (const t of Object.values(tables)) {
      if (t.deleted || t.id === nodeId) continue
      let fields: { id: string; name: string }[] = []
      try {
        fields = computeNode(t.id, tables, records).fields.map((f) => ({ id: f.id, name: f.name }))
      } catch {
        continue // 缺上游算不出来 → 不列（不猜）
      }
      for (const f of fields) out.push({ urn: `node://${t.id}#${f.id}`, label: `${t.name}·${f.name}` })
    }
    return out
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tables, records, nodeId])

  useEffect(() => {
    fceChangeHook = (v: string) => setFxDraft(v)
    fcePickCellHook = (coord: string) => {
      setPickCell(coord) // FCE 公式态点格高亮（WPS 选中效果）
      setGlowCell(null) // 点新参数格 = 活跃引用切换，清旧亮显
    }
    fceGlowHook = (coord: string | null) => setGlowCell(coord)
    fceHoverGlowHook = (coord: string | null) => {
      if (hoverGlowRef.current !== coord) {
        hoverGlowRef.current = coord
        gridApiRef.current?.refreshCells({ force: true })
      }
    }
    fceAliveHook = (alive: boolean) => setFceAlive(alive)
    return () => {
      fceChangeHook = null
      fcePickCellHook = null
      fceGlowHook = null
      fceHoverGlowHook = null
      fceAliveHook = null
    }
  }, [])
  /** 公式编辑态：公式栏聚焦且内容以 = 开头 —— 此时点格子 = 插引用 */
  const [fxEditing, setFxEditing] = useState(false)
  /** WPS 语义 · fx 栏活跃引用区：公式态连续点格 = 替换同一地址（不追加）；打字（onFxChange）/光标移动（onSelect）= 定型解除 */
  const [fxPick, setFxPick] = useState<{ s: number; e: number } | null>(null)
  const fxPickRef = useRef<{ s: number; e: number } | null>(null)
  /** WPS 语义 · 被点格高亮（fx 栏与 FCE 公式态共用）：金色选框直到再次点格/编辑态结束（selRect 同款 ref+refreshCells 模式） */
  const [pickCell, setPickCell] = useState<string | null>(null)
  const pickCellRef = useRef<string | null>(null)
  /** WPS 语义 · 点公式里的引用 → 网格对应格亮显（蓝框）；再点别处/编辑态结束清除 */
  const [glowCell, setGlowCell] = useState<string | null>(null)
  const glowCellRef = useRef<string | null>(null)
  /** 备忘 2026-09-27 ⑤ · hover 预览亮显（悬停公式里的引用/胶囊，比点击态淡、无动画）：
   *  hoverGlowRef = 悬停的 A1 引用格；hoverUrnRef = 悬停胶囊时当前 sheet 挂该 URN 的引用格集合；
   *  走 ref + refreshCells（mousemove 高频，不进 state）；mouseleave 全清，不碰点击态 glowCell/pickCell */
  const hoverGlowRef = useRef<string | null>(null)
  /** 备忘 2026-09-27 ⑥ · 拖选锚点（胶囊=一个字符）：mousedown 记起点偏移，mousemove 几何全走镜像坐标算选区——
   *  原生拖选按透明原文布局定位，URN 原文与胶囊宽度差导致高亮与鼠标错位（截图实锤），故 fx 拖选全接管 */
  const fxDragRef = useRef<number | null>(null)
  const hoverUrnRef = useRef<Set<string> | null>(null)
  const hoverUrnKeyRef = useRef<string | null>(null)
  /** 悬停胶囊 → 画布来源节点边框亮显（store 瞬态桥，Canvas TableNode 读 fxHoverNode） */
  const setFxHoverNode = useBaseStore((s) => s.setFxHoverNode)
  const fxHoverNodeRef = useRef<string | null>(null)
  useEffect(() => () => setFxHoverNode(null), [setFxHoverNode]) // 卸载兜底清画布亮显（mouseleave 可能没机会触发）
  /** WPS 语义 · FCE 存活信号（格内编辑中 = 淡显判定的一部分） */
  const [fceAlive, setFceAlive] = useState(false)
  /** WPS 语义 · 编辑态被引用格淡显：fxDraft（fx 栏直接输入 / FCE 经桥同步）里的 A1 引用集合 */
  const refDimRef = useRef<Set<string> | null>(null)
  const [refDimTick, setRefDimTick] = useState(0)
  useEffect(() => {
    const on = (fxEditing || fceAlive) && fxDraft.startsWith('=')
    const next = on ? calcRefCoords(fxDraft) : null
    // 备忘 2026-09-27 ①：外部引用格并入亮显集合 —— fxDraft token（可读名/URN 向下兼容）→ URN →
    // 当前 sheet 中挂该 URN 的引用格（「该单元格已引用的外部单元格」在网格上的锚点）
    if (on && next) {
      try {
        const st = stateRef.current
        const wss = st ? getSheet(st.wb, activeSheetRef.current) : null
        const urns = new Set<string>()
        for (const r of fxTokenRanges(fxDraft, resolvedFields)) if (r.urn) urns.add(r.urn)
        if (wss && urns.size) {
          const ext = getValueExtent(wss)
          if (ext) {
            for (let rr = ext.minRow; rr <= ext.maxRow; rr++) {
              for (let cc = ext.minCol; cc <= ext.maxCol; cc++) {
                const cell = getCellByCoord(wss, indexToCol(cc) + rr)
                const v = cell && typeof cell === 'object' ? (cell as { value?: unknown }).value : null
                if (typeof v === 'string' && urns.has(v)) next.add(indexToCol(cc) + rr)
              }
            }
          }
        }
      } catch {
        // 亮显是锦上添花：扫不动就只保留 A1 集合
      }
    }
    const prev = refDimRef.current
    const same = !!prev === !!next && (!next || [...next].sort().join() === [...prev!].sort().join())
    refDimRef.current = next
    if (!same) setRefDimTick((t) => t + 1)
  }, [fxDraft, fxEditing, fceAlive, tables])
  useEffect(() => {
    pickCellRef.current = pickCell
    gridApiRef.current?.refreshCells({ force: true })
  }, [pickCell])
  useEffect(() => {
    glowCellRef.current = glowCell
    gridApiRef.current?.refreshCells({ force: true })
  }, [glowCell])
  useEffect(() => {
    gridApiRef.current?.refreshCells({ force: true })
  }, [refDimTick])
  /** 编辑态结束 → 清活跃引用区与高亮（Enter/Escape/blur 150ms 三条路径统一走 fxEditing=false） */
  useEffect(() => {
    if (fxEditing) return
    fxPickRef.current = null
    setFxPick(null)
    setPickCell(null)
    setGlowCell(null)
  }, [fxEditing])
  /** 函数自动补全 */
  const [compl, setCompl] = useState<{ items: string[]; idx: number; prefix: string } | null>(null)
  const fxInputRef = useRef<HTMLInputElement>(null)
  /** 备忘 · 批次 8-b：fx 栏胶囊镜像层（input 文字/caret 透明 + 上层镜像画胶囊与假光标；IME 期间降级原文） */
  const [fxComposing, setFxComposing] = useState(false)
  const fxMirrorRef = useRef<HTMLDivElement>(null)
  // mirror-on 条件 = 含胶囊 token（可读「表名·参数名」或向下兼容的 URN 原文）——可读化后编辑串里通常没有 node://
  // 公式态（= 开头）也开镜像：坐标引用包 span 供点击亮显（WPS 语义）；数字/文本稳态不开
  const fxHasTok = fxDraft.includes('node://') || readableTokenRanges(fxDraft, resolvedFields).length > 0
  const fxMirrorOn = (fxHasTok || fxDraft.startsWith('=')) && !fxComposing
  /** 对照体系 · 输入态光标/选区（selectionchange 跟踪，镜像层据此画假光标/选区；失焦归 undefined 消光标） */
  const [fxCaret, setFxCaret] = useState<{ s: number; e: number } | undefined>(undefined)
  useEffect(() => {
    const onSel = () => {
      const el = fxInputRef.current
      if (!el || document.activeElement !== el) return
      setFxCaret({ s: el.selectionStart ?? 0, e: el.selectionEnd ?? 0 })
    }
    document.addEventListener('selectionchange', onSel)
    return () => document.removeEventListener('selectionchange', onSel)
  }, [])
  /** token → 内联胶囊（节点名·参数名）+ 假光标/选区；编辑串是可读串（URN 不进编辑器，用户 2026-09-27）。
   *  fxCaret null（未聚焦/失焦）→ 不画光标（blur 残留修复：不能兜底画在末尾，否则失焦后假光标常显） */
  const fxMirrorHtml = useMemo(
    () => buildUrnMirrorHtml(fxDraft, resolvedFields, fxMirrorOn ? fxCaret : undefined),
    [fxDraft, tables, fxMirrorOn, fxCaret],
  )
  /** 备忘 · 批次 8-b：fx 胶囊换参菜单 —— 胶囊锁死不可文本编辑，唯一修改方式 = 点击选同节点的不同参数（区间随草稿失效） */
  const [fxMenu, setFxMenu] = useState<{ start: number; end: number; nodeId: string; fid: string; left: number } | null>(null)
  useEffect(() => {
    if (!fxMenu) return
    const close = (e: MouseEvent) => {
      if ((e.target as HTMLElement | null)?.closest?.('.wb-fx-menu')) return
      setFxMenu(null)
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setFxMenu(null)
    }
    document.addEventListener('mousedown', close, true)
    document.addEventListener('keydown', onKey)
    return () => {
      document.removeEventListener('mousedown', close, true)
      document.removeEventListener('keydown', onKey)
    }
  }, [fxMenu])
  // 草稿一变，菜单里的区间立即作废（换参替换自身触发也无害）
  useEffect(() => {
    setFxMenu(null)
  }, [fxDraft])
  /** 视图开关（网格/零值/表头/公式显示） */
  const [viewGrid, setViewGrid] = useState(true)
  const [viewZeros, setViewZeros] = useState(true)
  const [viewHeaders, setViewHeaders] = useState(true)
  const [viewFormulas, setViewFormulas] = useState(false)
  /** 阅读模式：高亮选中格所在的行与列（右下角开关） */
  const [hlCross, setHlCross] = useState(false)
  /** 属性引用栏：显示开关（左下角按钮）/ 当前选中格的引用（node://…）/ 引用连线 path / 引用徽章菜单 */
  const [refPanelOpen, setRefPanelOpen] = useState(() => {
    // 备忘 · 记住上次引用栏开/关（按 nodeId）
    try {
      const saved = localStorage.getItem(`refpanel:${nodeId}`)
      if (saved !== null) return saved === '1'
    } catch {
      /* 读不了就算了 */
    }
    return true
  })
  /** 稳定引用（measure effect 依赖里不直接放 state，避免重挂观察器） */
  const refPanelOpenRef = useRef(refPanelOpen)
  refPanelOpenRef.current = refPanelOpen
  /** 备忘 · 引用栏开/关即持久化 */
  useEffect(() => {
    try {
      localStorage.setItem(`refpanel:${nodeId}`, refPanelOpen ? '1' : '0')
    } catch {
      /* 存不了就算了 */
    }
  }, [refPanelOpen, nodeId])

  /** 备忘 · 引用栏一退一进（普通视图）：关闭播退场动画再卸载，打开播入场动画；最大化直接切换 */
  const [panelMounted, setPanelMounted] = useState(refPanelOpen)
  const [panelLeaving, setPanelLeaving] = useState(false)
  const panelLeaveTimerRef = useRef<number | null>(null)
  useEffect(() => {
    if (panelLeaveTimerRef.current != null) {
      window.clearTimeout(panelLeaveTimerRef.current)
      panelLeaveTimerRef.current = null
    }
    if (refPanelOpen) {
      setPanelMounted(true)
      setPanelLeaving(false)
    } else if (maximized) {
      setPanelMounted(false)
      setPanelLeaving(false)
    } else if (panelMounted) {
      setPanelLeaving(true)
      panelLeaveTimerRef.current = window.setTimeout(() => {
        panelLeaveTimerRef.current = null
        setPanelMounted(false)
        setPanelLeaving(false)
      }, 240)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refPanelOpen, maximized])
  useEffect(
    () => () => {
      if (panelLeaveTimerRef.current != null) window.clearTimeout(panelLeaveTimerRef.current)
    },
    [],
  )

  /** 备忘 · 右侧「数据预选追加」面板：开关（持久化）/ 几何 / 进退动画 —— 全部镜像左栏 */
  const [rightOpen, setRightOpen] = useState(() => {
    try {
      return localStorage.getItem(`rightpanel:${nodeId}`) === '1'
    } catch {
      return false
    }
  })
  const rightOpenRef = useRef(rightOpen)
  rightOpenRef.current = rightOpen
  useEffect(() => {
    try {
      localStorage.setItem(`rightpanel:${nodeId}`, rightOpen ? '1' : '0')
    } catch {
      /* 存不了就算了 */
    }
  }, [rightOpen, nodeId])
  const [rightMounted, setRightMounted] = useState(rightOpen)
  const [rightLeaving, setRightLeaving] = useState(false)
  const rightLeaveTimerRef = useRef<number | null>(null)
  useEffect(() => {
    if (rightLeaveTimerRef.current != null) {
      window.clearTimeout(rightLeaveTimerRef.current)
      rightLeaveTimerRef.current = null
    }
    if (rightOpen) {
      setRightMounted(true)
      setRightLeaving(false)
    } else if (maximized) {
      setRightMounted(false)
      setRightLeaving(false)
    } else if (rightMounted) {
      setRightLeaving(true)
      rightLeaveTimerRef.current = window.setTimeout(() => {
        rightLeaveTimerRef.current = null
        setRightMounted(false)
        setRightLeaving(false)
      }, 240)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rightOpen, maximized])
  useEffect(
    () => () => {
      if (rightLeaveTimerRef.current != null) window.clearTimeout(rightLeaveTimerRef.current)
    },
    [],
  )
  /** 备忘 · 两栏都退出 → 自动进入最大化（下降沿触发一次；还原后不再自动循环） */
  const bothClosedRef = useRef(!refPanelOpen && !rightOpen)
  useEffect(() => {
    const both = !refPanelOpen && !rightOpen
    if (both && !bothClosedRef.current && !maximized && onToggleMax) onToggleMax()
    bothClosedRef.current = both
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refPanelOpen, rightOpen, maximized])
  /** 备忘 · 原信息栏 chip 参数菜单开关已随 chip 一并删除（fx 胶囊菜单接棒） */
  /** 备忘 · 右键单元格菜单（WPS 常规项；智能项不加） */
  const [ctxMenu, setCtxMenu] = useState<{ x: number; y: number; kind?: 'cell' | 'row'; row?: number } | null>(null)
  /** #33A/#39 列头右键菜单（独立于单元格菜单：col = 右键命中的列字母） */
  const [hdrMenu, setHdrMenu] = useState<{ x: number; y: number; col: string } | null>(null)
  /** #33A 筛选（v1 单列）：keep = 保留的显示值字符串列表；null = 无筛选 */
  const [filterState, setFilterState] = useState<{ col: string; keep: string[] } | null>(null)
  /** #33A 筛选值清单面板 */
  const [filterDlg, setFilterDlg] = useState<{ x: number; y: number; col: string } | null>(null)
  /** 筛选面板勾选集（打开时按当前筛选初始化，全选 = 语义清除） */
  const [filterPick, setFilterPick] = useState<Set<string>>(new Set())
  /** #33A 选择性粘贴：复制时同帧抓的格式快照（内存级，刷新即失；外部来源无快照 = 贴格式不可用） */
  const clipboardFmtRef = useRef<{ snap: FmtState[][]; tsv: string } | null>(null)
  useEffect(() => {
    if (!ctxMenu && !hdrMenu && !filterDlg) return
    const close = (e: MouseEvent) => {
      if ((e.target as HTMLElement | null)?.closest?.('.wb-ctxmenu')) return
      if ((e.target as HTMLElement | null)?.closest?.('.wb-filterdlg')) return
      setCtxMenu(null)
      setHdrMenu(null)
      setFilterDlg(null)
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setCtxMenu(null)
        setHdrMenu(null)
        setFilterDlg(null)
      }
    }
    document.addEventListener('mousedown', close, true)
    document.addEventListener('keydown', onKey)
    return () => {
      document.removeEventListener('mousedown', close, true)
      document.removeEventListener('keydown', onKey)
    }
  }, [ctxMenu, hdrMenu, filterDlg])
  const [selURN, setSelURN] = useState<{ nodeId: string; fieldId: string } | null>(null)
  /** 连线端点（详情提示条定位用）；线条本体走 DOM 直改 —— 滚动丝滑跟随不走 React 渲染 */
  const [refPts, setRefPts] = useState<{ x1: number; y1: number; x2: number; y2: number } | null>(null)
  const [refHover, setRefHover] = useState(false)
  /** 备忘 · 全部连线三态：0 跟随选中（默认，用户 2026-09-26 晚拍板「点关联的单元格才显示其连线，不要点不点都显示所有」）
   *  / 1 全部常显 / 2 隐藏（悬停按钮临时显全部保留）；点击循环 0→1→2→0。
   *  原后端毛玻璃态（旧 2）已随 wb-lines-back-on 废弃，遮罩渲染段一并删除 */
  const [allMode, setAllMode] = useState<0 | 1 | 2>(0)
  const [allLinesHover, setAllLinesHover] = useState(false)
  const [allLines, setAllLines] = useState<{ key: string; d: string }[]>([])
  const showAllFront = allMode === 1 || (allMode === 2 && allLinesHover)
  useEffect(() => {
    document.body.classList.remove('wb-lines-back-on')
    return () => document.body.classList.remove('wb-lines-back-on')
  }, [])
  const lineSvgRef = useRef<SVGSVGElement | null>(null)
  const linePathRef = useRef<SVGPathElement | null>(null)
  const hitPathRef = useRef<SVGPathElement | null>(null)
  /** 函数库面板 */
  const [openFnLib, setOpenFnLib] = useState(false)
  const [fnSearch, setFnSearch] = useState('')
  /** 批注清单 */
  const [openCmts, setOpenCmts] = useState(false)
  /** 设置单元格格式对话框（#33 第二块）：rect = 打开时的选区快照，init = 首格四轴初值 */
  const [fmtDlg, setFmtDlg] = useState<{ rect: CellRect; init: FmtState } | null>(null)
  /** 撤销栈：字节快照（190KB × 20 ≈ 3.8MB，可接受） */
  const undoStack = useRef<Uint8Array[]>([])
  /** 列选锚点列（Shift+点列头扩选用；不参与渲染） */
  const colAnchorRef = useRef<number | null>(null)
  /** 阅读模式悬停跟踪：鼠标所在格（优先于选中格，移出网格回退到选中格）——纯 ref 不触发渲染 */
  const hoverCellRef = useRef<{ r: number; c: number } | null>(null)
  /** 拖动框选进行中（onCellMouseDown → Over 扩选 → 全局 mouseup 结束） */
  const dragSelRef = useRef(false)
  /** 智能填充：把手位置（相对 .wb-grid）/ 拖拽状态 / 最近预览矩形 */
  const [fillHandle, setFillHandle] = useState<{ left: number; top: number } | null>(null)
  const fillDragRef = useRef<{ origin: { r1: number; c1: number; r2: number; c2: number } } | null>(null)
  const fillLastRef = useRef<{ r1: number; c1: number; r2: number; c2: number } | null>(null)
  const fillPreviewRef = useRef<HTMLDivElement>(null)

  /**
   * ★ 侧边栏 ↔ 弹窗的互斥开关（2026-10-04 用户报：打开数据引用后点批注等侧边菜单，
   *弹窗被侧边栏压住；再点一次只关了弹窗，侧边栏没关）。
   *
   * 根因是两件叠加的事：
   *① z-index 反了 —— 侧边栏写死 8100、弹窗遮罩只有 500，侧边栏浮在遮罩之上
   *  ② 状态各管各的 —— 打开弹窗完全不动 refPanelOpen/triggerOpen/rightOpen
   *
   * 只修 ① = 能点到但仍留下一个侧边栏；只修 ② = 能点到但点到的位置仍被遮。
   * 所以两边一起：z-index 归入--z-* 体系（弹窗用 --z-modal），
   * 而**所有侧边栏/弹窗的开关都从这里过**，保证「开一个必关同位的其他」。
   *
   * 三个侧边栏互斥同理：联动面板与引用栏同占左缘，此前就互相遮挡（见记忆里的踩坑）。
   */
  const closeSides = useCallback(() => {
    setRefPanelOpen(false)
    setTriggerOpen(false)
    setRightOpen(false)
  }, [])
  const openModal = useCallback(() => closeSides(), [closeSides])

  /**
   * 侧边栏互斥 toggle：引用栏/联动规则同占**左**缘，数据预选追加占**右**缘。
   * 关 → 只动自己；开 → 先把同缘的另一个收掉（此前联动面板与引用栏互相遮挡）。
   */
  const toggleSide = useCallback((which: 'ref' | 'trigger' | 'right') => {
    const cur = which === 'ref' ? refPanelOpen : which === 'trigger' ? triggerOpen : rightOpen
    if (cur) {
      if (which === 'ref') setRefPanelOpen(false)
      else if (which === 'trigger') setTriggerOpen(false)
      else setRightOpen(false)
      return
    }
    if (which === 'trigger') setRefPanelOpen(false)
    if (which === 'ref') setTriggerOpen(false)
    if (which === 'ref') setRefPanelOpen(true)
    else if (which === 'trigger') setTriggerOpen(true)
    else setRightOpen(true)
  }, [refPanelOpen, triggerOpen, rightOpen])
  /** 备忘 · 填充柄：悬浮在选区右下角格满 1 秒才出现；离开该格 / 滚动 / 换选区立即消失 */
  const [fillShown, setFillShown] = useState(false)
  const fillTimerRef = useRef<number | null>(null)
  const fillHoverRef = useRef<{ r: number; c: number } | null>(null)
  const clearFillTimer = () => {
    if (fillTimerRef.current != null) {
      window.clearTimeout(fillTimerRef.current)
      fillTimerRef.current = null
    }
  }
  /** ag-grid 的列头/单元格点击事件不带（或不可靠带）鼠标信息 —— 在 document capture 记录最近一次 mousedown 的 shiftKey */
  const lastDownShiftRef = useRef(false)
  useEffect(() => {
    const onDown = (e: MouseEvent) => {
      lastDownShiftRef.current = e.shiftKey
    }
    const onUp = () => {
      dragSelRef.current = false
    }
    document.addEventListener('mousedown', onDown, true)
    document.addEventListener('mouseup', onUp)
    return () => {
      document.removeEventListener('mousedown', onDown, true)
      document.removeEventListener('mouseup', onUp)
    }
  }, [])

  // refs 快照（键盘 handler 用稳定依赖，避免每次编辑重挂）
  const stateRef = useRef(state)
  stateRef.current = state
  const activeSheetRef = useRef(active)
  activeSheetRef.current = active
  const readOnlyRef = useRef(readOnly)
  readOnlyRef.current = readOnly

  // ── 属性引用栏：当前选中格若是引用（node://…）→ 解析出节点/字段（联动左栏高亮与连线）──
  useEffect(() => {
    if (!state || !selCoord || !active) {
      setSelURN(null)
      return
    }
    const wss = getSheet(state.wb, active)
    const cell = wss ? getCellByCoord(wss, selCoord) : null
    const v = cell && typeof cell === 'object' && 'value' in cell ? (cell as { value?: unknown }).value : cell
    const u = typeof v === 'string' ? parseUrn(v) : null
    setSelURN(u?.scheme === 'node' ? { nodeId: u.target, fieldId: u.locator } : null)
  }, [state, selCoord, active, dirty])

  // ── 格内编辑中的实时解析：编辑器里出现 =node://… → 连线与左栏高亮立即跟随（不用等提交）──
  useEffect(() => {
    const onInput = (e: Event) => {
      const t = e.target as HTMLInputElement | null
      if (!t || !t.closest('.wb-celleditor')) return
      const val = String(t.value ?? '')
      if (val.startsWith('=node://')) {
        const u = parseUrn(val.slice(1))
        if (u?.scheme === 'node') setSelURN({ nodeId: u.target, fieldId: u.locator })
      }
    }
    document.addEventListener('input', onInput, true)
    return () => document.removeEventListener('input', onInput, true)
  }, [])

  /** 左栏点字段 → 当前选中格写入引用 + 登记节点关系（画布自动连线） */
  // 属性引用栏定位（viewport 坐标）：
  //   常规视图 → Excel 窗体外左侧，**边缘与窗体完全一致**（同圆角 10px、同边框）、等高上下平齐（不集成在窗体里）
  //   最大化视图 → 嵌入窗体左缘，**top 从功能区底部开始**（不占 Ribbon 那一行）、仍居左
  const [panelPos, setPanelPos] = useState<{ left: number; top: number; height: number; radius: number; border: boolean } | null>(null)
  /** 备忘 · 右栏「数据预选追加」几何（镜像左栏：普通=窗体外右侧卡片 / 最大化=嵌入右缘） */
  const [rightPos, setRightPos] = useState<{ left: number; top: number; height: number; radius: number; border: boolean } | null>(null)
  /** 退场动画期间 panelPos 会被 measure 清空 —— 渲染用最后一次位置，避免面板跳位 */
  const panelPosLastRef = useRef<{ left: number; top: number; height: number; radius: number; border: boolean } | null>(null)
  if (panelPos) panelPosLastRef.current = panelPos
  const posForRender = panelPos ?? panelPosLastRef.current
  const rightPosLastRef = useRef<{ left: number; top: number; height: number; radius: number; border: boolean } | null>(null)
  if (rightPos) rightPosLastRef.current = rightPos
  const rightPosForRender = rightPos ?? rightPosLastRef.current
  useEffect(() => {
    const measure = () => {
      const ne = document.querySelector('.node-editor') as HTMLElement | null
      if (!ne) {
        setPanelPos(null)
        return
      }
      const nr = ne.getBoundingClientRect()
      const isMax = maximized || ne.classList.contains('ne-max')
      // 右栏几何（备忘 · 镜像左栏）：最大化=嵌入右缘；普通=窗体外右侧（屏宽不够则贴屏右）
      if (isMax) {
        const rb = document.querySelector('.rb')
        const rbBottom = rb ? rb.getBoundingClientRect().bottom : nr.top
        setRightPos({ left: nr.right - 230, top: rbBottom, height: nr.bottom - rbBottom, radius: 0, border: false })
      } else if (!rightOpenRef.current) {
        setRightPos(null)
      } else {
        setRightPos({
          left: Math.min(nr.right + 12, window.innerWidth - 230 - 8),
          top: nr.top,
          height: nr.height,
          radius: 10,
          border: true,
        })
      }
      if (isMax) {
        // 最大化：left=窗体左缘；top=功能区（.rb）底部 —— 引用栏不占 Ribbon 那一行
        ne.classList.remove('ne-refroom')
        const rb = document.querySelector('.rb')
        const rbBottom = rb ? rb.getBoundingClientRect().bottom : nr.top
        setPanelPos({ left: nr.left, top: rbBottom, height: nr.bottom - rbBottom, radius: 0, border: false })
      } else if (!refPanelOpenRef.current) {
        // 引用栏关闭：窗体恢复默认布局
        ne.classList.remove('ne-refroom')
        setPanelPos(null)
      } else {
        // 常规：窗体外左侧，边缘与窗体一致（圆角 10px、1px 边框）
        // ★ 判定基于 mask 内容宽的「未让位」几何 —— 不能用当前 rect（加了让位类后 rect 变化会导致条件翻转振荡）
        const w = 230
        const maskR = document.querySelector('.node-editor-mask')?.getBoundingClientRect()
        const pad = 28 // .node-editor-mask padding
        const contentW = (maskR ? maskR.width : window.innerWidth) - pad * 2
        const naturalW = Math.min(1180, contentW)
        const unshiftedLeft = (maskR ? maskR.left + pad : pad) + (contentW - naturalW) / 2
        if (unshiftedLeft - w - 12 < 8) {
          // 空间不够：引用栏保持贴屏左（不叠在 Excel 上），窗体右移让位、宽度收缩
          ne.classList.add('ne-refroom')
          setPanelPos({ left: 8, top: nr.top, height: nr.height, radius: 10, border: true })
        } else {
          ne.classList.remove('ne-refroom')
          setPanelPos({ left: Math.max(8, unshiftedLeft - w - 12), top: nr.top, height: nr.height, radius: 10, border: true })
        }
      }
    }
    measure()
    // Ribbon 收起/展开、窗体尺寸变化都会改变参照 —— ResizeObserver 兜底
    const ro = new ResizeObserver(() => measure())
    const ne = document.querySelector('.node-editor')
    const rb = document.querySelector('.rb')
    if (ne) ro.observe(ne)
    if (rb) ro.observe(rb)
    window.addEventListener('resize', measure)
    const t = setTimeout(measure, 60) // 布局稳定后兜底一次
    return () => {
      ro.disconnect()
      window.removeEventListener('resize', measure)
      clearTimeout(t)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [maximized, refPanelOpen, rightOpen])

  /** 左栏点字段 → 当前选中格写入引用 + 登记节点关系（画布自动连线） */
  // 引用连线：被引用单元格 → 属性引用栏（portal 侧栏）的参数项/节点标题 —— viewport 坐标（fixed SVG）
  // 端点元素查询（选中连线 / 批量连线 / hover 点亮共用）
  const queryRefEndpointEls = (nodeId: string, fieldId: string, coord: string) => {
    let sp: { r: number; c: number }
    try {
      sp = parseCoord(coord)
    } catch {
      return null
    }
    const cellEl = document.querySelector(
      `.ag-center-cols-container .ag-row[row-id="${sp.r}"] .ag-cell[col-id="${indexToCol(sp.c)}"]`,
    )
    // 端点：参数项（展开时）；节点折叠时 → 连到节点标题行（:not(.ref-panel-right) 排除右栏，避免左栏卸载后误连）
    const itemEl =
      document.querySelector(
        `.ref-panel:not(.ref-panel-right) [data-node="${nodeId}"][data-field="${CSS.escape(fieldId)}"]`,
      ) ?? document.querySelector(`.ref-panel:not(.ref-panel-right) .ref-node-row[data-node="${nodeId}"]`)
    if (!cellEl || !itemEl) return null
    return { cellEl, itemEl }
  }

  /** 由两端元素算连线。★ 单元格端钳到网格视口内（备忘）：
   *  竖向在下面被遮挡 → 接到最底滚动条上缘；左右被遮挡 → 接到该行起点（视口左缘）—— 线不出 Excel 界。
   *  备忘 · 单元格端接左缘（cr.left）：引用面板在左，线从面板出来贴到格子左边，不再穿过单元格接右缘 */
  const lineOfEls = (cellEl: Element, itemEl: Element) => {
    const cr = cellEl.getBoundingClientRect()
    const ir = itemEl.getBoundingClientRect()
    const vp = document.querySelector('.ag-body-viewport')?.getBoundingClientRect()
    let x1 = cr.left
    let y1 = cr.top + cr.height / 2
    if (vp) {
      x1 = Math.min(Math.max(x1, vp.left + 2), vp.right - 2)
      y1 = Math.min(Math.max(y1, vp.top + 2), vp.bottom - 2)
    }
    const x2 = ir.right
    const y2 = ir.top + ir.height / 2
    const mx = (x1 + x2) / 2
    return { d: `M ${x1} ${y1} C ${mx} ${y1}, ${mx} ${y2}, ${x2} ${y2}`, pts: { x1, y1, x2, y2 } }
  }

  // 选中格的连线 —— 测量逻辑放进 ref，供滚动/展开收起等实时事件复用（批次1-2 端点跟随）
  const measureRefLineRef = useRef<() => void>(() => {})
  useEffect(() => {
    const measure = () => {
      const svg = lineSvgRef.current
      if (!refPanelOpen || !selURN || !selCoord || !svg) {
        if (svg) svg.style.display = 'none'
        setRefPts(null)
        setRefHover(false)
        return
      }
      const els = queryRefEndpointEls(selURN.nodeId, selURN.fieldId, selCoord)
      if (!els) {
        svg.style.display = 'none'
        setRefPts(null)
        setRefHover(false)
        return
      }
      const { d, pts } = lineOfEls(els.cellEl, els.itemEl)
      svg.style.display = ''
      // ★ DOM 直改 path d —— 滚动丝滑跟随（备忘2），不经 React 渲染
      linePathRef.current?.setAttribute('d', d)
      hitPathRef.current?.setAttribute('d', d)

      setRefPts(pts)
    }
    measureRefLineRef.current = measure
    measure()
    // 备忘 · 面板入场动画是纯 CSS transform（不触发 DOM 变更/MutationObserver）—— 300ms 动画结束后补测一次，
    // 避免端点停留在动画中途的偏移位置（表现为连线没从参数/节点行右缘出线）
    const settle = window.setTimeout(() => measureRefLineRef.current(), 300)
    return () => window.clearTimeout(settle)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selURN, selCoord, refPanelOpen, dirty, state, hlCross, panelPos, maximized])

  // 批量连线（所有引用格 → 引用栏）—— 三态：跟随选中（不显批量，只显选中线）/ 全部常显 / 隐藏+悬停临时显
  const measureAllLinesRef = useRef<() => void>(() => {})
  useEffect(() => {
    const measure = () => {
      if (!showAllFront || !refPanelOpen || !state || !active) {
        setAllLines([])
        return
      }
      try {
        const lines: { key: string; d: string }[] = []
        for (const u of scanUrnCells(state.wb)) {
          if (u.sheet !== active) continue
          const els = queryRefEndpointEls(u.nodeId, u.fieldId, u.coord)
          if (!els) continue
          lines.push({ key: `${u.sheet}!${u.coord}`, d: lineOfEls(els.cellEl, els.itemEl).d })
        }
        setAllLines(lines)
      } catch {
        setAllLines([])
      }
    }
    measureAllLinesRef.current = measure
    measure()
    // 备忘 · 同选中线：面板入场动画结束后补测一次（动画中途的 rect 会让批量线面板端偏移）
    const settle = window.setTimeout(() => measureAllLinesRef.current(), 300)
    return () => window.clearTimeout(settle)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [showAllFront, refPanelOpen, state, active, dirty, panelPos, maximized])

  // 备忘 · 选区外轮廓 = 整块连续矩形 overlay（四角格推边界；虚拟化滚出的边钳到视口内，轮廓不出 Excel 界）
  const selOutlineRef = useRef<HTMLDivElement | null>(null)
  const measureSelOutlineRef = useRef<() => void>(() => {})
  /** 选区外轮廓 portal 挂载点 = ag-grid 滚动容器（.ag-body-viewport）：
   *  absolute + content 坐标只有挂在滚动容器内才能随滚动原生移动（挂 body 下会固定不动） */
  const [gridVpEl, setGridVpEl] = useState<HTMLElement | null>(null)

  // ── 批次1-2 · 端点实时跟随：面板滚动 / 网格滚动 / 节点展开收起 / 折叠按钮 —— 立即重算连线 ──
  useEffect(() => {
    const re = () => {
      measureRefLineRef.current()
      measureAllLinesRef.current()
      measureSelOutlineRef.current()
    }
    const onScroll = () => {
      re()
      // 备忘 · 滚动网格：填充柄立即消失（修「滚动了还停在原地」的残留），需重新悬浮 1 秒
      fillHoverRef.current = null
      clearFillTimer()
      setFillShown(false)
    }
    // capture 捕获所有内部滚动（ag-grid 视口 / 引用栏列表）
    document.addEventListener('scroll', onScroll, true)
    // 点引用栏（展开/收起节点、折叠按钮、搜索）→ 布局变化后重算
    const onClick = (e: MouseEvent) => {
      if ((e.target as HTMLElement)?.closest?.('.ref-panel')) requestAnimationFrame(() => re())
    }
    document.addEventListener('click', onClick, true)
    // 引用栏 DOM 结构变化（展开/收起节点卡）→ 重算
    let mo: MutationObserver | null = null
    const attachMo = () => {
      const list = document.querySelector('.ref-list')
      if (list && !mo) {
        mo = new MutationObserver(() => re())
        mo.observe(list, { childList: true, subtree: true })
      }
    }
    attachMo()
    const t = window.setTimeout(attachMo, 300)
    return () => {
      document.removeEventListener('scroll', onScroll, true)
      document.removeEventListener('click', onClick, true)
      mo?.disconnect()
      window.clearTimeout(t)
    }
  }, [refPanelOpen])

  /** 引线 hover：两端点亮（引用格 + 引用栏参数项/节点标题）—— 与连线共用端点查询 */
  const markRefEndpoints = (on: boolean) => {
    if (!selURN || !selCoord) return
    const els = queryRefEndpointEls(selURN.nodeId, selURN.fieldId, selCoord)
    els?.cellEl.classList.toggle('urn-endpoint-hl', on)
    els?.itemEl.classList.toggle('urn-endpoint-hl', on)
  }

  /** 左栏点字段：**必须先输入 =**（格内编辑或公式栏）才引入 —— 在光标处插入引用 token；非编辑态只提示不写格。
   *  fx 可读化：插入「表名·参数名」可读 token（URN 不进编辑器；提交边界 readableToUrn 转回） */
  const pickRef = (nodeId2: string, nodeName: string, fieldId: string, fieldName: string) => {
    if (readOnlyRef.current) return
    const urn = `node://${nodeId2}#${fieldId}`
    const tok = `${nodeName}·${fieldName}`
    // ① 格内编辑中（FCE 打开且值以 = 开头）→ 插入到格内公式光标处
    const fceInput = document.querySelector('.wb-celleditor input') as HTMLInputElement | null
    if (fceInput && fceInput.value.startsWith('=')) {
      fceInsertHook?.(tok)
      linkNode(nodeId2)
      return
    }
    // ② 公式栏编辑中（已输入 =）→ 插入到公式栏光标处
    if (fxEditing) {
      const el = fxInputRef.current
      // 选区吸附：当前选区含半截 token → 吞整 token 再替换（slice 删半截 = URN 残段，与截图泄漏同源）
      const fs0 = el?.selectionStart ?? fxDraft.length
      const fe0 = el?.selectionEnd ?? fs0
      const { s: sAt, e: eDel } = tokenAtomicSel(fs0, fe0, fxTokenRanges(fxDraft, resolvedFields))
      // 备忘 2026-09-27 ⑥-b：连续点参数自动补 + ——插入点前是 token 尾字符（字母/数字/中文/#/·）= 与上一引用紧贴
      // （用户「几个节点数据=相加」预期 = 点参数即相加项；紧贴提交会产生 URN 粘连 `#pricenode://` 实锤）
      const prev = fxDraft[sAt - 1] ?? ''
      const plus = /[\w\u4e00-\u9fff#·-]$/.test(prev) ? '+' : ''
      setFxDraft(fxDraft.slice(0, sAt) + plus + tok + fxDraft.slice(eDel))
      const pe = sAt + plus.length + tok.length
      requestAnimationFrame(() => {
        el?.focus()
        el?.setSelectionRange(pe, pe)
      })
      linkNode(nodeId2)
      return
    }
    // ③ 非编辑态：不写格，提示先输入 =
    setErr('先在单元格或公式栏输入「=」，再点左侧参数引入')
  }

  /** 引用登记（去重）—— Canvas 的 edges 由 linkedInputs 派生，退出编辑器即见「数据流」连线 */
  const linkNode = (nodeId2: string) => {
    if (!nodeId || !onSetNodeConfig) return
    const cur = tables?.[nodeId]?.config?.linkedInputs ?? []
    if (!cur.includes(nodeId2)) onSetNodeConfig(nodeId, { linkedInputs: [...cur, nodeId2] })
  }
  /** ref 包住最新 linkNode（commitRef 同款）：isPureUrn/粘贴等 useCallback 稳定 handler 里调，防闭包陈旧 */
  const linkNodeRef = useRef(linkNode)
  linkNodeRef.current = linkNode
  /**
   * 引用串/公式体里的 node:// 来源 → linkedInputs 登记 —— 引用格是事实源、画布连线是投影，
   * 但 fx 手打 / 粘贴 / 格内编辑不经过 RefPanel（原本唯一的登记入口），投影永远建不出来 →
   * 重开丢关联、连线消失（_diag_persist P1 实锤）。从写入内容提取来源补登记（含混排公式全部 token）。
   */
  const registerUrnSources = (text: string) => {
    if (!text) return
    if (text.startsWith('node://')) {
      const u = parseUrn(text)
      if (u?.scheme === 'node') linkNodeRef.current(u.target)
      return
    }
    for (const m of text.matchAll(/node:\/\/([^#\s(!*+\-/]+)#/g)) {
      if (m[1]) linkNodeRef.current(m[1])
    }
  }

  // ── 键盘直编辑（Excel 同款）：选中格后直接打字 = 替换输入（capture 抢在 ag-grid 前）；Delete/Backspace = 清除选区 ──
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      // eslint-disable-next-line no-console
      if (readOnlyRef.current) return
      const ae = document.activeElement as HTMLElement | null
      if (ae && (ae.tagName === 'INPUT' || ae.tagName === 'TEXTAREA' || ae.tagName === 'SELECT' || ae.isContentEditable)) return
      if (!gridApiRef.current || !selCoordRef.current) return
      // IME 组合（中文输入法）放行给原生路径
      if (e.isComposing || e.key === 'Process') return
      let sp: { r: number; c: number }
      try {
        sp = parseCoord(selCoordRef.current)
      } catch {
        return
      }
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault()
        e.stopPropagation()
        const rect = selRectRef.current
        const st = stateRef.current
        if (!rect || !st) return
        const wss = getSheet(st.wb, activeSheetRef.current)
        if (!wss) return
        // 清值用 clearRange —— setRangeValues 对 null 跳过不写、setCellByCoord(undefined) 静默忽略（两者实测都清不掉）
        clearRange(wss, `${indexToCol(rect.c1)}${rect.r1}:${indexToCol(rect.c2)}${rect.r2}`)
        void Promise.resolve(commitRef.current?.()).then(() => {
          // dirty 会重算 rowData，但 ag-grid 的 immutable 行更新偶有不重绘（实测 J5 残留旧值）—— 补一次强刷
          gridApiRef.current?.refreshCells({ force: true })
        })
        return
      }
      if (e.key.length !== 1 || e.ctrlKey || e.metaKey || e.altKey) return
      e.preventDefault()
      e.stopPropagation()
      pendingEditChar = (pendingEditChar ?? '') + e.key
      gridApiRef.current.startEditingCell({ rowIndex: Math.max(0, sp.r - 1), colKey: indexToCol(sp.c) })
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [])

  useEffect(() => {
    let dead = false
    ;(async () => {
      try {
        setErr('')
        setBusy('正在解析工作簿…')
        const { wb, applied } = await openWorkbook(bytes)
        if (dead) return
        ;(window as never as { __wb?: Workbook }).__wb = wb // 调试钩子
        const sheets = sheetNames(wb) ?? []
        setState({ wb, sheets, applied })
        // 备忘 · 记住上次活动的工作表（按 nodeId 存；不存在则回退第一个）
        let lastSheet = ''
        try {
          lastSheet = localStorage.getItem(`wb-sheet:${nodeId}`) ?? ''
        } catch {
          /* 读不了就算了 */
        }
        setActive(sheets.includes(lastSheet) ? lastSheet : (sheets[0] ?? ''))
        setFields(discoverXlsxFields(wb))
        undoStack.current = [bytes]
        setDirty(0)
      } catch (e) {
        if (!dead) setErr(e instanceof Error ? e.message : String(e))
      } finally {
        if (!dead) setBusy('')
      }
    })()
    return () => {
      dead = true
    }
    // 只在换文件时重开
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [name])

  /** 工作表清单从 wb 现取 —— 增删改名后 dirty 一变就刷新 */
  const sheets = useMemo(
    () => (state ? sheetNames(state.wb) ?? [] : []),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [state, dirty],
  )

  const marks = useMemo(() => buildMarks(fields), [fields])

  // ── 选区：单击锚定单格，Shift+点击 扩展矩形 ──
  const selRect = useMemo(() => {
    // 防御：selCoord 必须是单格坐标（#36 教训——区间串 'D1:E2' 混进来会让 parseCoord 抛错炸掉整树渲染）
    if (!selCoord || !isSingleCoord(selCoord)) return null
    const end = parseCoord(selCoord)
    const a = selAnchor ?? end
    return {
      r1: Math.min(a.r, end.r),
      r2: Math.max(a.r, end.r),
      c1: Math.min(a.c, end.c),
      c2: Math.max(a.c, end.c),
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selAnchor, selCoord])

  // ── 当前工作表的网格数据 ──
  // 公式格显示**缓存值**（office-kit 的 value 是 {kind:'formula', formula, cachedValue} 嵌套对象，
  // 旧实现 `typeof v === 'object' ? ''` 把它显示成空 —— 缺陷，阶段 1 修复）。
  // dirty 进依赖：重算写回缓存值后网格要跟着刷。
  const grid = useMemo(() => {
    const empty = { rows: [] as Record<string, unknown>[], allRows: [] as Record<string, unknown>[], cols: [] as string[] }
    if (!state || !active) return empty
    const ws = getSheet(state.wb, active)
    if (!ws) return empty
    const ext = getValueExtent(ws)
    const maxRow = Math.max(ext?.maxRow ?? 0, 30)
    const maxCol = Math.max(ext?.maxCol ?? 0, 12)
    const cols = Array.from({ length: maxCol }, (_, i) => indexToCol(i + 1))
    const rows: Record<string, unknown>[] = []
    for (let r = 1; r <= maxRow; r++) {
      const row: Record<string, unknown> = { __row: r }
      for (const c of cols) {
        const cell = getCellByCoord(ws, c + r)
        let disp: unknown = ''
        let urnStr: string | null = null
        if (cell && typeof cell === 'object' && 'value' in cell) {
          const v = (cell as { value?: unknown }).value
          if (v && typeof v === 'object' && (v as { kind?: string }).kind === 'formula') {
            const fv = v as { cachedValue?: unknown }
            disp = fv.cachedValue ?? ''
          } else {
            disp = v ?? ''
          }
        }
        // 引用格（node://…）→ 显示实际值（上游首行），并记 URN 供标记与连线
        if (typeof disp === 'string' && disp.startsWith('node://')) {
          urnStr = disp
          const u = parseUrn(disp)
          // ★ 批次17 补修：字段与取值都走 **计算后**结果（computeNode），与 urnResolverRef / RefPanel 同口径。
          // 旧代码用 tables[target].fields + records[target]：对 sort/map/output/docx/xlsx 这类汇聚型节点，
          // 自身 schema fields 为空、records 里也没有它的行（数据全在 mergeComputed 的结果里）
          // → fld 找不到、val 取不到 → 引用格一律显示「↗（无数据）」，看着像没写进去。
          // 算不出来（缺上游/环）→ 保留 URN 原文显示，不静默吞。
          if (u?.scheme === 'node' && tables) {
            try {
              const c = computeNode(u.target, tables, records ?? {})
              const fld = u.locator ? c.fields.find((f) => f.id === u.locator || f.name === u.locator) : null
              const val = c.rows[0]?.[(fld?.id ?? u.locator) as keyof Row]
              if (val != null) disp = val
              else disp = urnStr
            } catch {
              disp = urnStr
            }
          }
        }
        if (urnStr) row['__urn' + c] = urnStr
        row[c] = disp
      }
      rows.push(row)
    }
    // #33A 筛选（视图层）：不匹配的行隐藏，数据本体与行号不动（行号列显示原 __row）。
    // allRows = 未过滤全量（筛选值清单的来源，与过滤谓词同显示口径）
    let visRows = rows
    if (filterState) {
      const keep = new Set(filterState.keep)
      visRows = rows.filter((row) => keep.has(String(row[filterState.col] ?? '')))
    }
    return { rows: visRows, allRows: rows, cols }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [state, active, dirty, tables, records, filterState])

  // ★ Rules of Hooks：以下两段必须在 if (openFail) / if (!state) early return 之前 ——
  //   放 return 之后会让 loading 态渲染少执行一个 useEffect，工作簿就绪后 hooks 数 +1 → React #310 崩溃（实测）
  /** #33A 筛选值清单：全量行的该列显示值去重计数（与过滤谓词同口径 —— grid.allRows 的 String(row[col])） */
  const filterValueList = (col: string): { key: string; count: number }[] => {
    const counter = new Map<string, number>()
    for (const row of grid.allRows) {
      const key = String(row[col] ?? '')
      counter.set(key, (counter.get(key) ?? 0) + 1)
    }
    // 排序：空白最前、数字组数值序、文本 localeCompare（与排序键分组逻辑同风格）
    return [...counter.entries()]
      .map(([key, count]) => ({ key, count }))
      .sort((a, b) => {
        const ka = a.key === '' ? 2 : /^-?\d+(\.\d+)?$/.test(a.key.trim()) ? 0 : 1
        const kb = b.key === '' ? 2 : /^-?\d+(\.\d+)?$/.test(b.key.trim()) ? 0 : 1
        if (ka !== kb) return ka - kb
        if (ka === 0) return Number(a.key) - Number(b.key)
        return a.key.localeCompare(b.key, 'zh-Hans-CN')
      })
  }
  // 筛选面板打开：勾选集初始化（同列已有筛选 → 保留集；否则全选）
  useEffect(() => {
    if (!filterDlg) return
    setFilterPick(
      filterState?.col === filterDlg.col
        ? new Set(filterState.keep)
        : new Set(filterValueList(filterDlg.col).map((x) => x.key)),
    )
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [filterDlg])

  // colDefs 必须引用稳定：依赖里一旦有 selRect，每次点格子都会 setColumnDefs，
  // ag-grid 会把进行中的格内编辑（自定义编辑器）连根拔掉 —— 实测复现。
  // 选区高亮改走 ref + refreshCells。
  const selRectRef = useRef(selRect)
  /** #37 行高拖拽：行号底边 grip 按下 → 拖拽跟随基线 → 松手写回 rowHeights（18~400 clamp） */
  const rowDragRef = useRef<{ row: number; startY: number; startH: number } | null>(null)
  /** 备忘 · 填充柄悬浮检测用真实 mousemove（ag-grid 同格内移动不再触发 Over）：
   *  监听一次挂载 + ref 回调（capture 监听不随 state 反复挂卸的铁律） */
  const gridMouseMoveRef = useRef<(e: MouseEvent) => void>(() => {})
  gridMouseMoveRef.current = (e) => {
    // #37 行高拖拽中：基线跟随鼠标 y，不进入填充柄检测
    if (rowDragRef.current) {
      const guide = document.querySelector<HTMLElement>('.wb-rowheight-guide')
      if (guide) guide.style.top = e.clientY + 'px'
      return
    }
    const cellEl = (e.target as HTMLElement | null)?.closest?.('.ag-cell') as HTMLElement | null
    const colId = cellEl?.getAttribute('col-id')
    const rowId = cellEl?.closest('.ag-row')?.getAttribute('row-id')
    if (!colId || !rowId || !/^[A-Z]{1,3}$/.test(colId)) {
      fillHoverRef.current = null
      clearFillTimer()
      setFillShown(false)
      return
    }
    const nc = { r: Number(rowId), c: colToIndex(colId) }
    const rect = selRectRef.current
    const isCorner = !!rect && nc.r === rect.r2 && nc.c === rect.c2
    if (isCorner && !readOnlyRef.current) {
      const lastF = fillHoverRef.current
      if (!lastF || lastF.r !== nc.r || lastF.c !== nc.c) {
        fillHoverRef.current = nc
        clearFillTimer()
        fillTimerRef.current = window.setTimeout(() => {
          fillTimerRef.current = null
          setFillShown(true)
        }, 1000)
      }
    } else {
      fillHoverRef.current = null
      clearFillTimer()
      setFillShown(false)
    }
  }
  useEffect(() => {
    const onMove = (e: MouseEvent) => gridMouseMoveRef.current(e)
    document.addEventListener('mousemove', onMove, true)
    return () => document.removeEventListener('mousemove', onMove, true)
  }, [])

  // #37 行高拖拽：grip mousedown 启动（capture 拦截，不进 ag-grid 防误选整行）→ 基线跟随 → 松手写回
  useEffect(() => {
    const onDown = (e: MouseEvent) => {
      const grip = (e.target as HTMLElement | null)?.closest?.('.wb-rownum-grip') as HTMLElement | null
      if (!grip || readOnlyRef.current) return
      const row = Number(grip.getAttribute('data-rowgrip'))
      if (!row || !Number.isFinite(row)) return
      e.preventDefault()
      e.stopPropagation()
      const guide = document.createElement('div')
      guide.className = 'wb-rowheight-guide'
      guide.style.top = e.clientY + 'px'
      document.body.appendChild(guide)
      rowDragRef.current = { row, startY: e.clientY, startH: rowHeightsRef.current[row] ?? 26 }
    }
    const onUp = (e: MouseEvent) => {
      const drag = rowDragRef.current
      if (!drag) return
      rowDragRef.current = null
      document.querySelector('.wb-rowheight-guide')?.remove()
      const h = Math.min(400, Math.max(18, Math.round(drag.startH + (e.clientY - drag.startY))))
      setRowHeights((prev) => (prev[drag.row] === h ? prev : { ...prev, [drag.row]: h }))
    }
    document.addEventListener('mousedown', onDown, true)
    document.addEventListener('mouseup', onUp, true)
    return () => {
      document.removeEventListener('mousedown', onDown, true)
      document.removeEventListener('mouseup', onUp, true)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // #33A/#39 列头右键菜单：document capture 监听 .ag-header-cell（原生 contextmenu 事件拿坐标；
  // ag-grid v33 的 columnHeaderContextMenu 事件对象不带原生 event）。readOnly 不弹。
  const readOnlyRefHdr = useRef(readOnly)
  readOnlyRefHdr.current = readOnly
  useEffect(() => {
    const onCtx = (e: MouseEvent) => {
      if (readOnlyRefHdr.current) return
      const hdr = (e.target as HTMLElement | null)?.closest?.('.ag-header-cell') as HTMLElement | null
      if (!hdr) return
      const colId = hdr.getAttribute('col-id')
      if (!colId || !/^[A-Z]{1,3}$/.test(colId)) return
      e.preventDefault()
      e.stopPropagation()
      setHdrMenu({ x: e.clientX, y: e.clientY, col: colId })
    }
    document.addEventListener('contextmenu', onCtx, true)
    return () => document.removeEventListener('contextmenu', onCtx, true)
  }, [])

  /** 备忘 · 选区一变，填充柄就重置 —— 要重新悬浮右下角格满 1 秒才再出现 */
  useEffect(() => {
    setFillShown(false)
    clearFillTimer()
    fillHoverRef.current = null
  }, [selRect])
  selRectRef.current = selRect
  /** 备忘 · 选区外轮廓 overlay：选区/内容变化即重算（滚动走 realtime re()） */
  useEffect(() => {
  // ── 选区外轮廓：整块连续矩形 overlay —— 挂在 ag-grid 滚动容器内、content 坐标定位（用户 2026-09-26：
  //    「滚动后选择框滞后，要直接放上面」→ 随滚动原生移动，scroll 事件同步整个免掉）──
  const measure = () => {
    const el = selOutlineRef.current
    if (!el) return
    const vp = document.querySelector('.ag-body-viewport')
    if (!vp || !selRectRef.current) {
      el.style.display = 'none'
      return
    }
    const vRect = vp.getBoundingClientRect()
    const q = (r: number, c: number) =>
      document.querySelector(`.ag-center-cols-container .ag-row[row-id="${r}"] .ag-cell[col-id="${indexToCol(c)}"]`)
    // 格 rect → viewport content 坐标（含 scrollLeft/Top 补偿）；格子滚出虚拟化视野 → null，用对角格补
    const box = (r: number, c: number) => {
      const cel = q(r, c)
      if (!cel) return null
      const cr = cel.getBoundingClientRect()
      return { x1: cr.left - vRect.left + vp.scrollLeft, y1: cr.top - vRect.top + vp.scrollTop, x2: cr.right - vRect.left + vp.scrollLeft, y2: cr.bottom - vRect.top + vp.scrollTop }
    }
    const rect = selRectRef.current
    const tl = box(rect.r1, rect.c1) ?? box(rect.r1, rect.c2)
    const br = box(rect.r2, rect.c2) ?? box(rect.r2, rect.c1)
    if (!tl || !br) {
      el.style.display = 'none'
      return
    }
    // 超出 content 范围的部分钳掉（绝对定位溢出会撑大滚动区）
    const left = Math.max(0, Math.min(tl.x1, br.x1))
    const top = Math.max(0, Math.min(tl.y1, br.y1))
    const right = Math.min(Math.max(tl.x2, br.x2), vp.scrollWidth)
    const bottom = Math.min(Math.max(tl.y2, br.y2), vp.scrollHeight)
    el.style.display = ''
    el.style.left = `${left}px`
    el.style.top = `${top}px`
    el.style.width = `${Math.max(0, right - left)}px`
    el.style.height = `${Math.max(0, bottom - top)}px`
  }
    measureSelOutlineRef.current = measure
    measure()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selRect, dirty, active, state, maximized])
  /** 高亮行列：cellClass 走 ref 读取（colDefs 引用稳定，不进依赖） */
  const hlCrossRef = useRef(hlCross)
  hlCrossRef.current = hlCross
  const selCoordRef = useRef(selCoord)
  selCoordRef.current = selCoord
  const gridApiRef = useRef<GridApi | null>(null)
  // FCE Escape 取消桥注册端（fceCancelHook 声明见 fceInsertHook 旁）；stopEditing 无编辑器时 no-op，不清理安全
  fceCancelHook = () => {
    gridApiRef.current?.stopEditing(true)
    setPickCell(null) // 编辑态结束清点格高亮
    setGlowCell(null)
  }
  useEffect(() => {
    gridApiRef.current?.refreshCells({ force: true })
    // 列选高亮画在表头上，单元格刷新带不动它 —— 补一次表头刷新
    gridApiRef.current?.refreshHeader()
  }, [selRect, hlCross])

  // 智能填充柄：贴在选区右下角格的右下顶点（格被虚拟化滚出视野时隐藏）
  // 备忘：只在 fillShown（悬浮右下角格满 1 秒）时定位显示
  useEffect(() => {
    if (!fillShown || readOnly || fxEditing || !selRect || !state) {
      setFillHandle(null)
      return
    }
    const c2L = indexToCol(selRect.c2)
    const cellEl = document.querySelector(
      `.ag-center-cols-container .ag-row[row-id="${selRect.r2}"] .ag-cell[col-id="${c2L}"]`,
    )
    const gridEl = document.querySelector('.wb-grid')
    if (!cellEl || !gridEl) {
      setFillHandle(null)
      return
    }
    const cr = cellEl.getBoundingClientRect()
    const gr = gridEl.getBoundingClientRect()
    setFillHandle({ left: cr.right - gr.left - 4, top: cr.bottom - gr.top - 4 })
  }, [fillShown, selRect, dirty, active, state, fxEditing, readOnly])

  /** 备忘 · 选区一变，填充柄就重置 —— 重置逻辑在 selRect 声明之后（见 selRectRef 定义处） */

  /** 用户拖定的列宽（field → px）：colDefs 重建时带上，否则每次 marks/active 变化列宽都被写死的 108 覆盖（实测回弹） */
  /** #33 行高（行号 → px）：右键菜单「行高…」设置；getRowHeight 回调消费，26 为默认 */
  // 行高列宽按节点持久化（批次 3 #14）：localStorage `xlgeom:<nodeId>`，与 refpanel:/wb-sheet: 同款命名法；
  // 挂载惰性读一次、变更即写（数据量小无需防抖）。nodeId 缺省时静默跳过（不持久化）。
  const geomKey = nodeId ? `xlgeom:${nodeId}` : null
  const loadGeom = (k: 'colWidths' | 'rowHeights'): Record<string | number, number> => {
    if (!geomKey) return {}
    try {
      const raw = localStorage.getItem(geomKey)
      if (!raw) return {}
      const g = JSON.parse(raw) as { colWidths?: Record<string, number>; rowHeights?: Record<number, number> }
      const v = g[k]
      return v && typeof v === 'object' ? v : {}
    } catch {
      return {} /* 损坏数据忽略——几何偏好丢了无伤大雅 */
    }
  }
  const [colWidths, setColWidths] = useState<Record<string, number>>(() => loadGeom('colWidths') as Record<string, number>)
  const [rowHeights, setRowHeights] = useState<Record<number, number>>(() => loadGeom('rowHeights') as Record<number, number>)
  useEffect(() => {
    if (!geomKey) return
    try {
      localStorage.setItem(geomKey, JSON.stringify({ colWidths, rowHeights }))
    } catch {
      /* 写满/隐私模式静默 */
    }
  }, [geomKey, colWidths, rowHeights])
  const onColResized = useCallback((e: { finished?: boolean; column?: { getColId?: () => string; getActualWidth?: () => number } | null }) => {
    if (!e.finished || !e.column) return
    const id = e.column.getColId?.()
    const w = e.column.getActualWidth?.()
    if (!id || !w || !/^[A-Z]{1,3}$/.test(id)) return
    setColWidths((prev) => (prev[id] === w ? prev : { ...prev, [id]: w }))
  }, [])
  // 行高变化 → ag-grid 重算全部行高（getRowHeight 回调里读 ref 最新值，无需逐行 setRowHeight）
  const rowHeightsRef = useRef(rowHeights)
  rowHeightsRef.current = rowHeights
  useEffect(() => {
    gridApiRef.current?.resetRowHeights()
  }, [rowHeights])

  const colDefs = useMemo<ColDef[]>(() => {
    const list: ColDef[] = [
      {
        headerName: '',
        field: '__rownum', // 行号列标识：onCellClicked 靠它做整行选中（值仍走 valueGetter，无真实字段）
        valueGetter: (p) => p.data?.__row,
        width: 48,
        pinned: 'left',
        editable: false,
        cellClass: 'wb-rownum',
        cellDataType: false,
        // #37 行高拖拽：行号格内底部 5px 手柄（.wb-rownum-grip），拖拽逻辑走 gridMouseMoveRef 委托
        // ⚠ 网格是 AgGridReact：cellRenderer 返回 string 会被转义成纯文本、返回原生 DOM 元素会炸 React #31
        //   （两种都实测过）——必须返回 ReactNode（JSX）
        cellRenderer: (p: { value?: unknown }) => (
          <span>
            <span className="wb-rownum-txt">{typeof p.value === 'number' ? p.value : ''}</span>
            <span className="wb-rownum-grip" data-rowgrip={typeof p.value === 'number' ? String(p.value) : ''} />
          </span>
        ),
      },
    ]
    const m = marks.get(active)
    for (const c of grid.cols) {
      const colIdx = colToIndex(c)
      list.push({
        // #33A 筛选：该列有筛选条件 → 列头字母后带漏斗标记（视觉提示）
        headerName: filterState?.col === c ? c + ' ▾' : c,
        field: c,
        editable: !readOnly,
        // 列宽：优先用户拖过的宽度（onColumnResized 持久化到 colWidths），colDefs 重建时不再回弹
        width: colWidths[c] ?? 108,
        // 格内编辑器：= 开头自动进公式模式（补全 + 点格子插引用）
        cellEditor: FormulaCellEditor,
        // 引用格进编辑 = 公式编辑态（用户 2026-09-26：已创建引用，点进编辑应是公式态而非显示值）：
        // 行数据格值已是显示值（上游首行值），URN 原文在 __urn<col> 键 —— FCE 初始值改给 '=node://…'，
        // 镜像胶囊立即生效；提交路径 '= 剥离 → isPureUrn' 原样存回
        cellEditorParams: (p: { data?: Record<string, unknown> | null; value?: unknown }) => {
          const urn = p.data?.['__urn' + c]
          // 不能用 initialValue 键：ag-grid proxy getProps() 在 spread 后硬编码 initialValue=格值，会覆盖；
          // 自定义键 urnInit 穿透（spread 保留），FCE 内优先读取
          if (urn) return { urnInit: '=' + urn }
          // 公式格（<f> 双写）双击 = 编辑公式体而非结果值（用户 2026-09-27：原来是公式生成的结果，进编辑应变成编辑公式，与 fx 栏一致）
          const rowNum = (p.data as Record<string, unknown> | undefined)?.__row
          const wss = stateRef.current ? getSheet(stateRef.current.wb, activeSheetRef.current) : null
          const cell = wss && rowNum ? getCellByCoord(wss, c + rowNum) : null
          const fv = cell && typeof cell === 'object' ? (cell as { formula?: string }).formula : null
          return { urnInit: fv ? '=' + fv : undefined }
        },
        // 引用徽章（批次 8 · 两段式胶囊）：node://… 渲染成「节点名｜参数」胶囊（外部引用样式）
        cellRenderer: (p: { value?: unknown; data?: Record<string, unknown> | null }) => {
          const v = p.value
          if (typeof v === 'string' && v.startsWith('node://')) {
            const u = parseUrn(v)
            if (u?.scheme === 'node') {
              const t = tables?.[u.target]
              const fld = u.locator ? t?.fields.find((f) => f.id === u.locator || f.name === u.locator) : null
              const tip = v.replace(/'/g, '')
              // 节点已删 / 字段不存在 → 红边失效态；无 locator = 整节点引用（右段「整体」）；
              // tables 未加载（暂态）不算失效，走默认金系
              return capsuleHtml({
                node: t?.name ?? u.target,
                field: fld?.name ?? u.locator,
                kind: t?.kind,
                dead: !!tables && (!t || (!!u.locator && !fld)),
                title: `外部引用：${tip}（选中后可在信息栏切换参数）`,
              })
            }
          }
          // 数字格式显示（#33 数字页签配套）：格有非 General 数字格式 → 按 OOXML 格式码渲染
          // 显示值（getCellDisplayText：0.5 + '0%' → '50%'）。只动显示层——行数据/编辑初值/
          // 复制粘贴仍走原值，编辑框里看到的是原始数值（v1 口径，待用户反馈再对齐 WPS）。
          if (v != null) {
            const stNow = stateRef.current
            const rowNum = (p.data as Record<string, number> | undefined)?.__row
            const wss = stNow && rowNum ? getSheet(stNow.wb, active) : null
            const cell = wss ? getCellByCoord(wss, c + rowNum) : null
            if (cell) {
              const fmt = getCellNumberFormat(stNow!.wb, cell)
              if (fmt && fmt !== 'General') {
                try {
                  return getCellDisplayText(stNow!.wb, cell)
                } catch {
                  /* 格式码解析失败回退原值显示 */
                }
              }
            }
          }
          return v == null ? '' : String(v)
        },
        // #33 第二块：OOXML 单元格样式渲染（字体/填充/边框/对齐）——
        // cellStyleToCss 一次出全套 CSS，inline style 盖过主题底色；无样式格返回 null 走默认
        cellStyle: (p: { data?: Record<string, unknown> | null }) => {
          const stNow = stateRef.current
          const rowNum = (p.data as Record<string, number> | undefined)?.__row ?? 0
          const wss = stNow ? getSheet(stNow.wb, active) : null
          const cell = wss ? getCellByCoord(wss, c + rowNum) : null
          if (!cell || !stNow) return null
          const css = cellStyleToCss(stNow.wb, cell)
          return Object.keys(css).length ? css : null
        },
        // 列选高亮：选区起始于第 1 行且覆盖本列 = 列选态（点列头触发）
        headerClass: (p) => {
          const col = (p.colDef as { field?: string }).field
          if (!col) return 'wb-head'
          const rect = selRectRef.current
          const inCol = !!rect && rect.r1 === 1 && colToIndex(col) >= rect.c1 && colToIndex(col) <= rect.c2
          return 'wb-head' + (inCol ? ' wb-head-sel' : '')
        },
        // 真实文件同一列常混类型（数字列里夹文本格/错误格）——关掉 ag-grid 的
        // cellDataType 自动推断，否则它把文本格渲染成 "Invalid Number"（设备材料总表实测）
        cellDataType: false,
        cellClass: (p) => {
          const rowNum = (p.data as Record<string, number> | undefined)?.__row ?? 0
          let cls = ''
          if (m?.named.has(c + rowNum)) cls += 'wb-cell-named'
          else if (inAnyTable(m, rowNum, colIdx)) cls += 'wb-cell-table'
          const rect = selRectRef.current
          if (rect && rowNum >= rect.r1 && rowNum <= rect.r2 && colIdx >= rect.c1 && colIdx <= rect.c2) {
            cls += (cls ? ' ' : '') + 'wb-cell-inrange'
          }
          // 阅读模式：高亮「悬停格」所在的行与列（悬停优先），鼠标不在网格上时回退到选中格
          if (hlCrossRef.current) {
            const hv = hoverCellRef.current
            let eff: { r: number; c: number } | null = hv
            if (!eff && selCoordRef.current) {
              try {
                eff = parseCoord(selCoordRef.current)
              } catch {
                eff = null
              }
            }
            if (eff && (rowNum === eff.r || colIdx === eff.c)) cls += (cls ? ' ' : '') + 'wb-cell-cross'
          }
          // 引用格标记（值是 node://… 的格子：金色虚线框）
          if ((p.data as Record<string, unknown> | undefined)?.['__urn' + p.colDef.field]) {
            cls += (cls ? ' ' : '') + 'wb-cell-refurn'
          }
          // WPS 语义 · 公式态点格高亮（被点格 = 当前活跃引用，金色选框）
          if (pickCellRef.current === c + rowNum) cls += (cls ? ' ' : '') + 'wb-cell-pickref'
          // #36 公式态拖拽选区间高亮（fx 栏与 FCE 共用 fxDragRangeRef）
          const dr = fxDragRangeRef.current
          if (dr && rowNum >= dr.r1 && rowNum <= dr.r2 && colIdx >= dr.c1 && colIdx <= dr.c2) {
            cls += (cls ? ' ' : '') + 'wb-cell-dragrng'
          }
          // WPS 语义 · 编辑态被引用格淡显 / 点公式里的引用亮显
          if (refDimRef.current?.has(c + rowNum)) cls += (cls ? ' ' : '') + 'wb-cell-refdim'
          if (glowCellRef.current === c + rowNum) cls += (cls ? ' ' : '') + 'wb-cell-refglow'
          // 备忘 2026-09-27 ⑤：hover 预览亮显（悬停公式里的 A1 引用 / 胶囊对应的 URN 引用格 —— 比点击态淡、无动画）
          if (hoverGlowRef.current === c + rowNum) cls += (cls ? ' ' : '') + 'wb-cell-hoverglow'
          if (hoverUrnRef.current?.has(c + rowNum)) cls += (cls ? ' ' : '') + 'wb-cell-hoverurn'
          // 备忘 · 选区外轮廓改为整块 overlay 连续矩形（见 measureSelOutline）——
          // 原逐格 wb-eb-* 描边在多格选区时有间断，且曾因未约束矩形内画出整行整列十字
          return cls
        },
      })
    }
    return list
    // eslint-disable-next-line react-hooks/exhaustive-deps
    // dirty 进 deps：格式对话框 applyFmtSel 走 commit（setDirty+1）→ colDefs 重建 →
    // ag-grid 刷新单元格 → cellStyle/cellRenderer 重新执行，新样式/新数字显示立即上屏
  }, [grid.cols, marks, active, readOnly, colWidths, dirty, filterState])

  // ── 选中格信息（公式栏数据源）──
  const selInfo = useMemo(() => {
    if (!state || !active || !selCoord) return null
    const ws = getSheet(state.wb, active)
    if (!ws) return null
    const cell = getCellByCoord(ws, selCoord)
    if (!cell || typeof cell !== 'object') return { coord: selCoord, formula: '', display: '' }
    const v = (cell as { value?: unknown }).value
    if (v && typeof v === 'object' && (v as { kind?: string }).kind === 'formula') {
      const fv = v as { formula?: string; cachedValue?: unknown }
      return { coord: selCoord, formula: fv.formula ?? '', display: String(fv.cachedValue ?? '') }
    }
    // URN 引用格：formula 给 URN 原文（不带 =，与公式惯例一致）→ 公式栏草稿重置为 '=node://…'，
    // 点进即公式编辑态；回车走 isPureUrn 分支存回原文（修复「单击引用格 + 回车清格」）
    if (typeof v === 'string' && v.startsWith('node://')) {
      return { coord: selCoord, formula: v, display: v }
    }
    return { coord: selCoord, formula: '', display: v == null ? '' : String(v) }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [state, active, selCoord, dirty])

  /** 批注清单（必须在所有提前 return 之前调用 hook） */
  const cmtRows = useMemo(() => {
    if (!(openCmts && state && active)) return []
    const w = getSheet(state.wb, active)
    return w ? listComments(w) : []
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [openCmts, state, active, dirty])

  // 选中格变化 → 公式栏草稿重置为该格公式（有公式）或显示值
  useEffect(() => {
    if (!selInfo) {
      setFxDraft('')
      return
    }
    // 换格 = 旧编辑上下文作废：活跃引用区与点格高亮一并清（防 blur 150ms 定时器未到时的残留串位）
    fxPickRef.current = null
    setFxPick(null)
    setPickCell(null)
    setGlowCell(null)
    // fx 可读化：草稿进可读串（表名·参数名），URN 不进编辑器（用户 2026-09-27）；tables 经 ref 取避免依赖抖动
    setFxDraft(selInfo.formula ? '=' + urnToReadable(selInfo.formula, resolvedFieldsRef.current) : selInfo.display)
  }, [selInfo])

  // ── 统一落盘：重算本表公式 → 补丁式保存 → undo 快照 ──
  const commit = useCallback(async () => {
    if (!state) return
    try {
      recomputeSheet(state.wb, active, { resolveUrn: (u) => urnResolverRef.current(u) })
    } catch {
      // 重算失败不阻断保存 —— 保留原缓存值总比整表丢公式好
    }
    const out = (await patchSaveWorkbook(bytes, state.wb, { mtime: FIXED_MTIME })).bytes
    undoStack.current = [...undoStack.current.slice(-19), out]
    setDirty((d) => d + 1)
    onChange(out)
    reconcileRef.current()
  }, [state, active, bytes, onChange])
  /** 键盘直编辑等稳定 handler 用：ref 包住最新 commit（commit 声明在此，TDZ 原因不能提前 useRef(commit)） */
  const commitRef = useRef(commit)
  commitRef.current = commit

  /**
   * ★ 章节编排插槽（2026-10-05）—— 章节能力与内核之间**唯一**的接缝。
   *
   * 架构口径（用户定）：「基于当前全量功能定制，差异化内容在 oo 系列更新后
   * 能用最小代价适配」。落地方式就是这**一个可选 props**：
   *   · outline 不传 → enabled=false，下面所有章节代码都不执行，行为与改动前完全一致
   *   · outline 传了 → 多一个「章节编排」按钮 + 挂ChapterDialog，其余全部复用
   *
   * 章节算法、xlsx 往返、窗体 UI 都在outline.ts / outlineAdapter.ts /
   * ChapterDialog.tsx / useOutlineSlot.ts 四个独立文件里，本文件只负责「调用」。
   */
  const outlineSlot = useOutlineSlot(
    state?.wb ?? null,
    active,
    () => commitRef.current?.(),
    outline,
    setErr,
  )

  /**
   * 引用对账（commit 末尾调用，与 commitRef 同款 TDZ 规避）——
   * 工作簿已无某来源的引用格、且它不是画布手动连线（inputs）→ 从 linkedInputs 摘除，
   * 画布「数据流」连线（由 linkedInputs 派生）随之消失，不留残线。
   * 覆盖所有清引用路径：Delete/Backspace 清区、格内改写、公式栏改写、徽章切换。
   */
  const reconcileRef = useRef<() => void>(() => {})
  reconcileRef.current = () => {
    if (readOnlyRef.current || !nodeId || !onSetNodeConfig) return
    const t = tables?.[nodeId]
    const cur = t?.config?.linkedInputs ?? []
    if (!cur.length) return
    const manual = new Set(t?.config?.inputs ?? [])
    let refs: Set<string>
    try {
      refs = state ? collectUrnSources(state.wb) : new Set()
    } catch {
      return
    }
    const next = Array.from(new Set(cur.filter((id) => manual.has(id) || refs.has(id))))
    if (next.length !== cur.length) onSetNodeConfig(nodeId, { linkedInputs: next })
  }

  /**
   * 孤儿引用格：画布删线/删节点后，引用格还在、但来源已不在 linkedInputs∪inputs。
   * 不静默清数据 —— 提示条告知 + 一键「清空引用格」。
   */
  const orphanRefs = useMemo<UrnCellRef[]>(() => {
    if (!state || !nodeId) return []
    const t = tables?.[nodeId]
    if (!t) return []
    const ok = new Set([...(t.config?.linkedInputs ?? []), ...(t.config?.inputs ?? [])])
    try {
      return scanUrnCells(state.wb).filter((x) => !ok.has(x.nodeId))
    } catch {
      return []
    }
    // dirty 是 wb 原地变更的信号（与既有 effect 同款依赖）
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [state, dirty, tables, nodeId])

  /** 一键清空孤儿引用格（只动引用格，其它数据不碰）→ commit → 对账把残留连线收干净 */
  const clearOrphans = () => {
    if (!state || !orphanRefs.length) return
    try {
      for (const o of orphanRefs) {
        const wss = getSheet(state.wb, o.sheet)
        if (wss) clearRange(wss, `${o.coord}:${o.coord}`)
      }
    } catch (e2) {
      setErr('清空引用格失败：' + (e2 instanceof Error ? e2.message : String(e2)))
      return
    }
    void Promise.resolve(commitRef.current?.()).then(() => {
      gridApiRef.current?.refreshCells({ force: true })
    })
  }

  /** 一键恢复孤儿引用格的连线（历史 linkedInputs 丢失 / 删线后悔）→ linkedInputs 补登 → 画布连线重建 */
  const restoreOrphans = () => {
    if (!orphanRefs.length) return
    for (const o of orphanRefs) linkNodeRef.current(o.nodeId)
  }

  /** 工具栏动作统一入口：改 wb → 落盘；失败进顶部提示条（fn 返回 string = 错误文案；其余返回值忽略） */
  const run = useCallback(
    (fn: () => unknown) => {
      try {
        const msg = fn()
        if (typeof msg === 'string' && msg) setErr(msg)
        void commit()
      } catch (e2) {
        setErr(e2 instanceof Error ? e2.message : String(e2))
      }
    },
    [commit],
  )

  // ── 联动触发：执行一份计划（写值 + 重算已在 apply 内做）→ 统一落盘 ──
  const applyTrgPlan = useCallback(
    async (plan: TriggerPlan) => {
      if (!state) return
      const res = applyTriggerPlan(state.wb, plan, {
        resolveUrn: (u) => urnResolverRef.current(u),
      })
      // 跳过的必须让用户看见 —— 否则"规则命中了但文件没变"会变成无解的疑问
      if (res.skipped.length) {
        setErr(`已写 ${res.written} 格，跳过 ${res.skipped.length} 格：${res.skipped[0].target} ${res.skipped[0].reason}`)
      } else {
        const totalCalc = Object.values(res.recomputed).reduce((a, b) => a + b.updated, 0)
        setErr(`已按规则更新 ${res.written} 格，重算 ${totalCalc} 格`)
      }
      // 走既有 commit：重算本表 + 补丁保存 + undo 快照 + 引用对账，一套口径不另开路径
      await commit()
    },
    [state, commit],
  )

  /** 干跑：只回报"会改哪几格"，不落盘 */
  const dryRunTrgPlan = useCallback(
    (plan: TriggerPlan) => {
      if (!state) return { written: 0, skipped: [] as { target: string; reason: string }[] }
      return applyTriggerPlan(state.wb, plan, {
        resolveUrn: (u) => urnResolverRef.current(u),
        dryRun: true,
      })
    },
    [state],
  )

  // ── 编辑一格 → 就地补丁 → 重算 → 重新序列化 ──
  const onCellChanged = useCallback(
    async (e: CellValueChangedEvent) => {
      // eslint-disable-next-line no-console
      if (!state || !e.colDef.field || e.rowIndex == null) return
      // WPS 语义：FCE 回车提交后选中格切到点选的参数格；没点过格（普通编辑）才清高亮。
      // #36 区间引用（'D1:E2'）不是单格 → 不切（留在编辑格，Excel 同款），否则 parseCoord 炸渲染
      if (pickCellRef.current && isSingleCoord(pickCellRef.current)) setSelCoord(pickCellRef.current)
      else {
        setPickCell(null)
        setGlowCell(null)
      }
      const col = e.colDef.field
      const row = (e.data as Record<string, number> | undefined)?.__row
      if (!row) return
      const ws = getSheet(state.wb, active)
      if (!ws) return
      const coord = col + row

      let v: unknown = e.newValue
      if (typeof v === 'string') {
        // fx 可读化提交边界：编辑串里的「表名·参数名」token → URN（存储层永远是 URN，URN 不进编辑器）
        const s2 = readableToUrn(v, resolvedFields)
        const t = s2.trim()
        if (t === '') v = undefined
        else if (/^-?\d+(\.\d+)?$/.test(t)) v = Number(t)
        else v = s2
      }
      // ★ 格子内直接输入 =公式：识别为公式并双写（缓存值我们自己算 —— 计算权威唯一）
      if (typeof v === 'string' && v.startsWith('=')) {
        const body = v.slice(1)
        // =node://… 纯引用（整格）不是求值公式 —— 存 URN 文本（显示层解析为上游值）；
        // 混排公式（胶囊 token 参与四则）走下方求值分支（批次 1-6）
        if (isPureUrn(body)) {
          setCellByCoord(ws, coord, body)
          registerUrnSources(body)
          await commit()
          return
        }
        const ref = collectRefValues(state.wb, active, { resolveUrn: (u) => urnResolverRef.current(u) })
        const out = evalRowFormula(body, ref, (u) => urnResolverRef.current(u))
        if (out.ok) {
          setCellByCoord(ws, coord, makeFormula(body, { cachedValue: out.value as number | string | boolean }))
        } else {
          setCellByCoord(
            ws,
            coord,
            makeFormula(body, { cachedValue: '#VALUE!', cachedValueType: 'error' as never })
          )
          setErr(`公式「${v}」暂算不出（已写入，Excel 打开会重算）：${out.error}`)
        }
        registerUrnSources(body)
      } else {
        setCellByCoord(ws, coord, (v === '' ? undefined : v) as never)
      }
      try {
        await commit()
      } catch (e2) {
        setErr('保存这一格时出错：' + (e2 instanceof Error ? e2.message : String(e2)))
      }
    },
    [state, active, commit, tables]
  )

  // ── 公式栏提交 ──
  const applyFormulaBar = useCallback(async () => {
    if (!state || !selCoord) return
    const ws = getSheet(state.wb, active)
    if (!ws) return
    // WPS 语义（用户 2026-09-27）：公式态点过参数格 → 回车确认后选中格切到参数格（不是留在编辑格）。
    // #36 区间引用（'D1:E2'）不是单格 → 不切（parseCoord 只认单格，切了会炸 selRect 渲染）
    if (pickCellRef.current && isSingleCoord(pickCellRef.current)) setSelCoord(pickCellRef.current)
    // fx 可读化提交边界：编辑串「表名·参数名」token → URN（存储层永远是 URN；手打 URN 原样通过）
    const raw = readableToUrn(fxDraft.trim(), resolvedFields)
    try {
      if (raw.startsWith('=')) {
        const body = raw.slice(1)
        // =node://… 纯引用（整格）存 URN 文本；混排公式（胶囊 token 参与四则）走求值（批次 1-6）
        if (isPureUrn(body)) {
          setCellByCoord(ws, selCoord, body)
          registerUrnSources(body)
          setFxEditing(false)
          setCompl(null)
          await commit()
          return
        }
        const ref = collectRefValues(state.wb, active, { resolveUrn: (u) => urnResolverRef.current(u) })
        const out = evalRowFormula(body, ref, (u) => urnResolverRef.current(u))
        if (out.ok) {
          setCellByCoord(ws, selCoord, makeFormula(body, { cachedValue: out.value as number | string | boolean }))
        } else {
          setCellByCoord(
            ws,
            selCoord,
            makeFormula(body, { cachedValue: '#VALUE!', cachedValueType: 'error' as never })
          )
          setErr(`公式求值失败（已写入，打开 Excel 时会重算）：${out.error}`)
        }
        registerUrnSources(body)
      } else {
        let v: unknown = raw
        const t = raw.trim()
        if (t === '') v = undefined
        else if (/^-?\d+(\.\d+)?$/.test(t)) v = Number(t)
        setCellByCoord(ws, selCoord, (v === '' ? undefined : v) as never)
      }
      setFxEditing(false)
      setCompl(null)
      await commit()
    } catch (e2) {
      setErr('写入失败：' + (e2 instanceof Error ? e2.message : String(e2)))
    }
  }, [state, active, selCoord, fxDraft, commit, tables])

  // ── 公式编辑态下点格子 = 往公式里插引用（Excel 核心交互）──
  const insertCellRef = useCallback(
    (coord: string) => {
      const el = fxInputRef.current
      // WPS 语义（用户 2026-09-27）：公式态连续点格 = 替换上一个活跃引用（不追加 A1B1C1）；首次点格在光标处插入
      const pick = fxPickRef.current
      // 选区吸附：当前选区（含 WPS 活跃引用区）含半截 token → 吞整 token 再替换（防 URN 残段）
      const fs0 = pick ? 0 : el?.selectionStart ?? fxDraft.length
      const fe0 = pick ? 0 : el?.selectionEnd ?? fs0
      const abs = pick ? { s: pick.s, e: pick.e } : tokenAtomicSel(fs0, fe0, fxTokenRanges(fxDraft, resolvedFields))
      const s = pick ? pick.s : abs.s
      const eDel = pick ? pick.e : abs.e
      // 备忘 2026-09-27 ⑥-b：插入点前是 token 尾字符 → 自动补 +（紧贴提交会产生 URN/坐标粘连，同 pickRef）
      const prev = fxDraft[s - 1] ?? ''
      const plus = /[\w\u4e00-\u9fff#·-]$/.test(prev) ? '+' : ''
      const next = fxDraft.slice(0, s) + plus + coord + fxDraft.slice(eDel)
      setFxDraft(next)
      const pe = s + plus.length + coord.length
      fxPickRef.current = { s, e: pe }
      setFxPick({ s, e: pe })
      setPickCell(coord) // 被点格高亮（WPS 选中效果）
      setGlowCell(null) // 点新参数格 = 活跃引用切换，清旧亮显
      requestAnimationFrame(() => {
        el?.focus()
        el?.setSelectionRange(pe, pe)
      })
    },
    [fxDraft, tables]
  )

  /** 公式态点格子后，onCellClicked 里跳过这一次（引用已由 capture 拦截器插入，防双插） */
  const suppressClickRef = useRef(false)

  // 公式栏公式态：mousedown capture 阶段拦截格子点击 = 插引用 + 关补全。
  // 必须在 capture 层做：补全下拉浮在网格上方，普通 click 会被下拉项吃掉。
  // 监听只挂/卸一次，回调里实时查 ref —— effect 随输入重挂会干扰 ag-grid 的 click 判定
  // ★ 焦点判定：fxEditing 会在 input 失焦后残留 true（无 blur 清理），修复 C 又让稳态草稿常态带 =，
  //   两者叠加 → 残焦态点格全被误拦插引用（实测 URN 被追加 4 个格号）。只有焦点真在 fx input 才算公式编辑态。
  // 备忘 2026-09-27 ③：fx 编辑中不论是否 = 开头，点格一律插入 A1 引用（WPS 同款；纯文本前缀 + 点格 = 拼坐标），
  //   残焦风险仍由下方 activeElement 焦点闸门兜底。
  const fxActiveRef = useRef(false)
  fxActiveRef.current = fxEditing && !readOnly
  /** #36 公式态拖拽选区间：ar/ac = 锚点（mousedown 格），r1/c1/r2/c2 = 归一矩形（cellClass 高亮消费）。
   *  fx 栏直接读写；FCE 经模块级 fceRangeDragHook 桥接写同一个 ref。null = 无拖拽 */
  const fxDragRangeRef = useRef<{ r1: number; c1: number; r2: number; c2: number } | null>(null)
  // #36 FCE 拖拽桥：FCE 内部监听经 fceRangeDragHook 写同一个 fxDragRangeRef（cellClass 统一消费高亮）
  useEffect(() => {
    fceRangeDragHook = (rect) => {
      fxDragRangeRef.current = rect
      gridApiRef.current?.refreshCells({ force: true })
    }
    return () => {
      fceRangeDragHook = null
    }
  }, [])
  useEffect(() => {
    const hitCell = (e: MouseEvent): { r: number; c: number } | null => {
      const t = e.target as HTMLElement | null
      const cell = t?.closest?.('.ag-cell') as HTMLElement | null
      if (!cell) return null
      const col = cell.getAttribute('col-id')
      const rowId = cell.closest('.ag-row')?.getAttribute('row-id')
      if (!col || !/^[A-Z]{1,3}$/.test(col) || !rowId || !/^\d+$/.test(rowId)) return null
      return { r: Number(rowId), c: colToIndex(col) }
    }
    const onDocMouseDown = (e: MouseEvent) => {
      if (!fxActiveRef.current) return
      if (document.activeElement !== fxInputRef.current) return // 残焦态：焦点已离开 fx 栏，放行普通格点击
      const hit = hitCell(e)
      if (!hit) return
      e.preventDefault()
      e.stopPropagation()
      suppressClickRef.current = true
      setCompl(null)
      // #36：mousedown 只记锚点+高亮，mouseup 才落引用 —— 拖过 = 区间 A1:B3，单击 = 单格（原行为）
      fxDragRangeRef.current = { r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c }
      gridApiRef.current?.refreshCells({ force: true })
    }
    const onDocMouseMove = (e: MouseEvent) => {
      const drag = fxDragRangeRef.current
      if (!drag) return
      const hit = hitCell(e)
      if (!hit) return
      const r1 = Math.min(drag.r1, hit.r)
      const r2 = Math.max(drag.r1, hit.r)
      const c1 = Math.min(drag.c1, hit.c)
      const c2 = Math.max(drag.c1, hit.c)
      if (r1 === drag.r1 && r2 === drag.r2 && c1 === drag.c1 && c2 === drag.c2) return
      fxDragRangeRef.current = { r1, c1, r2, c2 }
      gridApiRef.current?.refreshCells({ force: true })
    }
    const onDocMouseUp = () => {
      const drag = fxDragRangeRef.current
      if (!drag) return
      fxDragRangeRef.current = null
      gridApiRef.current?.refreshCells({ force: true })
      const coord =
        drag.r1 === drag.r2 && drag.c1 === drag.c2
          ? `${indexToCol(drag.c1)}${drag.r1}`
          : `${indexToCol(drag.c1)}${drag.r1}:${indexToCol(drag.c2)}${drag.r2}`
      insertCellRef(coord)
    }
    document.addEventListener('mousedown', onDocMouseDown, true)
    document.addEventListener('mousemove', onDocMouseMove, true)
    document.addEventListener('mouseup', onDocMouseUp, true)
    return () => {
      document.removeEventListener('mousedown', onDocMouseDown, true)
      document.removeEventListener('mousemove', onDocMouseMove, true)
      document.removeEventListener('mouseup', onDocMouseUp, true)
    }
  }, [insertCellRef])

  // ── 函数自动补全 ──
  const updateCompl = useCallback((draft: string, caret: number) => {
    if (!draft.startsWith('=')) {
      setCompl(null)
      return
    }
    const before = draft.slice(0, caret)
    const m = /([A-Za-z][A-Za-z0-9.]*)$/.exec(before)
    if (!m) {
      // = 或 +-*/(, 之后光标悬停（还没打字母）→ 直接给常用函数列表（Excel 的 = 行为）
      if (/=$|[+\-*/(,]$/.test(before)) {
        setCompl({ items: COMMON_FNS.slice(0, 8), idx: 0, prefix: '' })
        return
      }
      setCompl(null)
      return
    }
    const prefix = m[1].toUpperCase()
    const common = COMMON_FNS.filter((f) => f.startsWith(prefix))
    const rest = FUNC_NAMES.filter((f) => f.startsWith(prefix) && !common.includes(f))
    const items = [...common, ...rest].slice(0, 8)
    setCompl(items.length ? { items, idx: 0, prefix: m[1] } : null)
  }, [])

  const onFxChange = useCallback(
    (val: string) => {
      // token 完整性守卫：IME/拖放等绕过 onKeyDown 锁死的插入落进胶囊内部 → 拒绝并把光标推出 token（不清 pick）
      const bad = tokenInsertGuard(fxDraft, val, fxTokenRanges(fxDraft, resolvedFields))
      if (bad != null) {
        requestAnimationFrame(() => fxInputRef.current?.setSelectionRange(bad, bad))
        return
      }
      // 删除守卫（插入守卫的删除侧对偶）：token 被渐进删空留残段（如 `://param_…#price`）→ 拒绝，光标推 token 末尾
      const badDel = tokenDeleteGuard(fxDraft, val, fxTokenRanges(fxDraft, resolvedFields))
      if (badDel != null) {
        requestAnimationFrame(() => fxInputRef.current?.setSelectionRange(badDel, badDel))
        return
      }
      setFxDraft(val)
      setFxEditing(true)
      // 打字 = 引用定型（WPS 语义）：活跃引用区解除，下次点格重新插入
      fxPickRef.current = null
      setFxPick(null)
      const caret = fxInputRef.current?.selectionStart ?? val.length
      updateCompl(val, caret)
    },
    [updateCompl, fxDraft, tables]
  )

  const acceptCompl = useCallback(() => {
    if (!compl) return
    const fn = compl.items[compl.idx]
    const el = fxInputRef.current
    const caret = el ? el.selectionStart ?? fxDraft.length : fxDraft.length
    const before = fxDraft.slice(0, caret - compl.prefix.length)
    const after = fxDraft.slice(caret)
    const next = before + fn + '()' + after
    setFxDraft(next)
    setCompl(null)
    requestAnimationFrame(() => {
      const p = before.length + fn.length + 1
      el?.focus()
      el?.setSelectionRange(p, p)
    })
  }, [compl, fxDraft])

  /** 往公式栏插入函数（菜单/函数库入口） */
  const insertFn = useCallback((f: string) => {
    setFxDraft((d) => (d.startsWith('=') ? d + f + '()' : '=' + f + '()'))
    setFxEditing(true)
    fxInputRef.current?.focus()
  }, [])

  // ── 粘贴 TSV（Ctrl+V）──
  useEffect(() => {
    if (!state || readOnly) return
    const onPaste = (e: ClipboardEvent) => {
      if (!selCoord || fxEditing) return
      // 焦点在 input/textarea（fx 栏 / FCE / 搜索框）→ 粘贴交给 input 自己的 onPaste，这里不双写
      const ae = document.activeElement as HTMLElement | null
      if (ae && (ae.tagName === 'INPUT' || ae.tagName === 'TEXTAREA')) return
      const text = e.clipboardData?.getData('text/plain') ?? ''
      // 单格文本也写入（用户 2026-09-26：「excel里面还是不能用 ctrl+C 和ctrl+V」——原来只认含 tab/换行的表格文本）
      if (!text) return
      e.preventDefault()
      try {
        const r = pasteTSV(state.wb, active, selCoord, text)
        registerUrnSources(text)
        setErr('')
        void commit().then(() => setErr(`已粘贴 ${r.rows} 行 × ${r.cols} 列`))
      } catch (e2) {
        setErr('粘贴失败：' + (e2 instanceof Error ? e2.message : String(e2)))
      }
    }
    window.addEventListener('paste', onPaste)
    return () => window.removeEventListener('paste', onPaste)
  }, [state, active, selCoord, readOnly, commit, fxEditing])

  // ── 复制选区（Ctrl+C）：网格态（焦点不在 input/textarea）拦下走 copySelTSV，编辑态放行原生复制 ──
  // 坑：copySelTSV 是组件体后方的 const（TDZ），effect 闭包直接引用实测报 Cannot access before initialization
  // （React 并发渲染下 effect 注册与组件体执行的时序不稳）→ 经 ref 中转，赋值在 copySelTSV 定义之后
  const copySelRef = useRef<() => void>(() => {})
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.key.toLowerCase() !== 'c') return
      const t = e.target as HTMLElement | null
      const ae = document.activeElement as HTMLElement | null
      const inEditor =
        (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable)) ||
        (ae && (ae.tagName === 'INPUT' || ae.tagName === 'TEXTAREA'))
      if (inEditor) return
      if (!selRectRef.current) return
      e.preventDefault()
      copySelRef.current()
    }
    document.addEventListener('keydown', onKey, true)
    return () => document.removeEventListener('keydown', onKey, true)
  }, [])

  // ── 最大化：状态由 NodeEditor 标题栏持有（props.max），面板内不再自持 ──

  const doUndo = () => {
    if (undoStack.current.length <= 1) return
    const stack = [...undoStack.current]
    stack.pop()
    const prev = stack[stack.length - 1]
    undoStack.current = stack
    setDirty((d) => Math.max(0, d - 1))
    onChange(prev)
    // 立即重解析恢复视图 —— 解析 effect 只在换文件（name）时重开，bytes 流回来不触发，
    // 原实现只回滚字节不动 state.wb → 网格显示不变（撤销等于没生效，2026-09-26 _diag_errfloat 实锤）
    void (async () => {
      try {
        const { wb, applied } = await openWorkbook(prev)
        setState({ wb, sheets: sheetNames(wb) ?? [], applied })
        setFields(discoverXlsxFields(wb))
        // rows 重算后 ag-grid immutable 行更新偶有不重绘（同 commit 后强刷先例）
        requestAnimationFrame(() => gridApiRef.current?.refreshCells({ force: true }))
      } catch {
        /* 恢复的字节来自栈内历史快照（解析器自身产物），解析失败保持现状即可 */
      }
    })()
  }

  const reload = async () => {
    setBusy('重新解析…')
    try {
      const { wb } = await openWorkbook(bytes)
      setState({ wb, sheets: sheetNames(wb) ?? [], applied: [] })
      setFields(discoverXlsxFields(wb))
      undoStack.current = [bytes]
      setDirty(0)
      setOpenFail(null)
    } catch (e) {
      setOpenFail(asOpenFailure(e))
    } finally {
      setBusy('')
    }
  }

  if (openFail) {
    return (
      <div className="wb">
        <OpenFailureNotice failure={openFail} name={name} bytes={bytes} onRetry={() => void reload()} />
        {onSave && (
          <div className="wb-foot">
            <span className="wb-sp" />
            <button className="bar-btn" onClick={() => onSave(bytes)} title="解析失败也不影响原始字节">
              仍然另存为…
            </button>
          </div>
        )}
      </div>
    )
  }

  if (!state) return <div className="wb wb-loading">{busy || '正在打开…'}</div>

  const ws = getSheet(state.wb, active)
  const ext = ws ? getValueExtent(ws) : null
  const nonEmpty = ws ? getNonEmptyCellCount(ws) : 0
  const wbTables = ws ? (listTables(ws) ?? []) : []
  const merges = ws ? (getMergedCells(ws) ?? []) : []
  const namedHere = marks.get(active)

  const withRect = (fn: (rect: { r1: number; c1: number; r2: number; c2: number }) => void) => {
    if (!selRect) {
      setErr('先点选一个格子（Shift+点击可扩展选区）')
      return
    }
    run(() => fn(selRect))
  }

  /** 自动求和：对选中格上方同列连续区域求和（v1：从第 1 行到选中行上一行） */
  const autoSum = () => {
    if (!selCoord || !ws) return setErr('先选中求和结果要放的格子')
    const { r, c } = parseCoord(selCoord)
    if (r <= 1) return setErr('请选中求和结果行（其上方需有数据）')
    const col = indexToCol(c)
    const body = `SUM(${col}1:${col}${r - 1})`
    const ref = collectRefValues(state.wb, active, { resolveUrn: (u) => urnResolverRef.current(u) })
    const out = evalRowFormula(body, ref, (u) => urnResolverRef.current(u))
    setCellByCoord(
      ws,
      selCoord,
      makeFormula(body, out.ok ? { cachedValue: out.value as never } : { cachedValue: '#VALUE!', cachedValueType: 'error' as never })
    )
    void commit()
  }

  /** WPS 风格功能区页签（原菜单 + 工具栏合并，全部沿用既有 handler） */
  const ribbonTabs: RibbonTab[] = [
    {
      label: '开始',
      groups: [
        {
          header: '剪贴板',
          items: [
            { kind: 'btn', label: '粘贴', glyph: '📋', title: '从 WPS/Excel 复制后 Ctrl+V，从选中格开始贴入', onClick: () => setErr('复制区域后按 Ctrl+V，会从选中格开始贴入') },
          ],
        },
        {
          header: '行列',
          items: [
            { kind: 'btn', label: '↑插行', title: '在选中行上方插入一行', onClick: () => withRect((rc) => insertRowAbove(state.wb, active, rc.r1)) },
            { kind: 'btn', label: '删行', title: '删除选中行', onClick: () => withRect((rc) => deleteRowAt(state.wb, active, rc.r1)) },
            { kind: 'btn', label: '←插列', title: '在选中列左侧插入一列', onClick: () => withRect((rc) => insertColLeft(state.wb, active, rc.c1)) },
            {
              kind: 'btn',
              label: '删列',
              title: '删除选中的列（可整列选中后删，从右往左删避免索引移位）',
              onClick: () => withRect((rc) => { for (let c = rc.c2; c >= rc.c1; c--) deleteColAt(state.wb, active, c) }),
            },
          ],
        },
        {
          header: '合并',
          items: [
            { kind: 'btn', label: '合并', title: '合并选区（Shift+点击扩展选区）', onClick: () => withRect((rc) => void mergeSel(state.wb, active, rc)) },
            { kind: 'btn', label: '拆分', title: '取消选区内合并', onClick: () => withRect((rc) => void unmergeSel(state.wb, active, rc)) },
          ],
        },
        {
          header: '字体',
          items: [
            { kind: 'btn', label: '加粗', glyph: 'B', title: '加粗（选区）', onClick: () => withRect((rc) => fontSel(state.wb, active, rc, { bold: true })) },
            { kind: 'btn', label: '斜体', glyph: 'I', title: '斜体（选区）', onClick: () => withRect((rc) => fontSel(state.wb, active, rc, { italic: true })) },
            { kind: 'btn', label: '下划线', glyph: 'U', title: '下划线（选区）', onClick: () => withRect((rc) => fontSel(state.wb, active, rc, { underline: true as never })) },
            'sep',
            { kind: 'btn', label: '字号 12', title: '字号 12（选区）', onClick: () => withRect((rc) => fontSel(state.wb, active, rc, { size: 12 })) },
          ],
        },
        {
          header: '对齐',
          items: [
            { kind: 'btn', label: '居中', glyph: '≡', title: '选区居中', onClick: () => withRect((rc) => alignSel(state.wb, active, rc, 'center')) },
            { kind: 'btn', label: '左对齐', title: '选区左对齐', onClick: () => withRect((rc) => alignSel(state.wb, active, rc, 'left')) },
            { kind: 'btn', label: '右对齐', title: '选区右对齐', onClick: () => withRect((rc) => alignSel(state.wb, active, rc, 'right')) },
          ],
        },
        {
          header: '数字',
          items: [
            { kind: 'btn', label: '#,##0', title: '千分位（选区）', onClick: () => withRect((rc) => numberFmtSel(state.wb, active, rc, '#,##0')) },
            { kind: 'btn', label: '0.00', title: '两位小数（选区）', onClick: () => withRect((rc) => numberFmtSel(state.wb, active, rc, '0.00')) },
            { kind: 'btn', label: '%', title: '百分比（选区）', onClick: () => withRect((rc) => numberFmtSel(state.wb, active, rc, '0.00%')) },
          ],
        },
        {
          header: '冻结',
          items: [
            { kind: 'btn', label: '冻结', title: '冻结至选中格上方与左方', onClick: () => run(() => freezeAt(state.wb, active, selCoord || undefined)) },
            { kind: 'btn', label: '解冻', title: '取消冻结', onClick: () => run(() => freezeAt(state.wb, active, undefined)) },
          ],
        },
      ],
    },
    {
      label: '插入',
      groups: [
        {
          header: '函数（' + FUNC_COUNT + ' 个）',
          items: [
            { kind: 'btn', label: '函数库', glyph: 'ƒ', title: '搜索并插入函数', onClick: () => { closeSides(); setFnSearch(''); setOpenFnLib(true) } },
            'sep',
            ...COMMON_FNS.map((f) => ({ kind: 'btn' as const, label: f + '()', title: '把 ' + f + '() 插入公式栏', onClick: () => insertFn(f) })),
          ],
        },
      ],
    },
    {
      label: '页面',
      groups: [
        {
          header: '列宽',
          items: [
            { kind: 'btn', label: '自适应', title: '自适应列宽（全表）', onClick: () => run(() => { if (ws) autofitColumns(ws) }) },
            { kind: 'btn', label: '宽 ×1.5', title: '选中列宽 ×1.5', onClick: () => withRect((rc) => { if (ws) for (let c = rc.c1; c <= rc.c2; c++) setColumnWidth(ws, c, 108 * 1.5) }) },
            { kind: 'btn', label: '宽 ×0.7', title: '选中列宽 ×0.7', onClick: () => withRect((rc) => { if (ws) for (let c = rc.c1; c <= rc.c2; c++) setColumnWidth(ws, c, 108 * 0.7) }) },
          ],
        },
        {
          header: '缩放',
          items: [
            { kind: 'btn', label: '150%', onClick: () => run(() => { if (ws) setSheetZoom(ws, 150) }) },
            { kind: 'btn', label: '100%', onClick: () => run(() => { if (ws) setSheetZoom(ws, 100) }) },
            { kind: 'btn', label: '75%', onClick: () => run(() => { if (ws) setSheetZoom(ws, 75) }) },
          ],
        },
      ],
    },
    {
      label: '公式',
      groups: [
        {
          header: '计算',
          items: [
            { kind: 'btn', label: 'Σ 自动求和', glyph: 'Σ', title: '上方同列区域求和（结果放选中格）', onClick: () => run(() => autoSum()) },
            { kind: 'btn', label: '函数库', title: '搜索并插入函数', onClick: () => { closeSides(); setFnSearch(''); setOpenFnLib(true) } },
            'sep',
            ...COMMON_FNS.slice(0, 6).map((f) => ({ kind: 'btn' as const, label: f + '()', onClick: () => insertFn(f) })),
            'sep',
            { kind: 'btn', label: '重算工作表', title: '重算本表全部公式', onClick: () => run(() => { recomputeSheet(state.wb, active, { resolveUrn: (u) => urnResolverRef.current(u) }) }) },
          ],
        },
      ],
    },
    {
      label: '数据',
      groups: [
        {
          header: '域与粘贴',
          items: [
            { kind: 'btn', label: refPanelOpen ? '✓ 属性引用栏' : '属性引用栏', title: '左栏：选择外部关联的参数', onClick: () => toggleSide('ref') },
            {
              kind: 'btn',
              label: triggerOpen ? '✓ 联动规则' : '联动规则',
              title: '左栏：参数改了按什么规则更新进表格',
              onClick: () => toggleSide('trigger'),
            },
            { kind: 'btn', label: '域位置清单', onClick: () => { closeSides(); setShowDomains(true) } },
            { kind: 'btn', label: '粘贴', title: '从 WPS/Excel 复制后 Ctrl+V', onClick: () => setErr('复制区域后按 Ctrl+V，会从选中格开始贴入') },
          ],
        },
      ],
    },
    {
      label: '审阅',
      groups: [
        {
          header: '批注',
          items: [
            {
              kind: 'btn',
              label: '新建批注',
              title: '在选中格新建批注',
              onClick: () => {
                const text = prompt('批注内容：')
                if (!text) return
                run(() => { if (ws) { setComment(ws, { ref: selCoord, author: '我', text }); return '批注已加到 ' + selCoord } })
              },
            },
            { kind: 'btn', label: '清空批注', title: '清空本表全部批注', onClick: () => run(() => { if (ws) { const n = removeAllComments(ws); return '已清空 ' + n + ' 条批注' } }) },
            { kind: 'btn', label: '批注清单', onClick: () => { closeSides(); setOpenCmts(true) } },
            // ★ 章节编排入口（仅 outline 插槽启用时出现，不传 outline 则这一项不存在）
            //   位置紧挨批注清单：两者都是「清单型辅助窗体」，交互心智一致。
            //   openModal 走内核既有的侧边栏互斥通道（与批注清单同一套，不新开互斥逻辑）。
            ...(outlineSlot.enabled
              ? ([
                  {
                    kind: 'btn' as const,
                    label: '章节编排',
                    title: '目录大纲：折叠、升降级、定位章节行',
                    onClick: () => {
                      openModal()
                      outlineSlot.setOpen(true)
                    },
                  },
                ] as RibbonItem[])
              : []),
          ],
        },
      ],
    },
    {
      label: '视图',
      groups: [
        {
          header: '窗口',
          items: [
            { kind: 'btn', label: '域位置', disabled: !fields.length, title: '域位置清单', onClick: () => { closeSides(); setShowDomains(true) } },
          ],
        },
      ],
    },
    {
      label: '工具',
      groups: [
        {
          header: '编辑',
          items: [
            { kind: 'btn', label: '撤销', onClick: () => doUndo() },
            { kind: 'btn', label: '重新载入', title: '放弃未保存改动，重新解析', onClick: () => void reload() },
            ...(onSave ? [{ kind: 'btn' as const, label: '另存为副本', onClick: () => onSave(bytes) }] : []),
          ],
        },
      ],
    },
  ]

  /** 函数库过滤结果（普通计算 —— 不能放提前 return 之后的 useMemo） */
  const fnQ = fnSearch.trim().toUpperCase()
  const common = COMMON_FNS.filter((f) => !fnQ || f.includes(fnQ))
  const rest = FUNC_NAMES.filter((f) => (!fnQ || f.includes(fnQ)) && !common.includes(f))
  const fnLibItems = [...common, ...rest]

  // ── 备忘 · 右键单元格菜单（WPS 常规项）──
  type CtxItem = { label: string; run?: () => void; sep?: boolean; title?: string; children?: CtxItem[] }
  const copySelTSV = () => {
    const rect = selRectRef.current
    const st = stateRef.current
    if (!rect || !st) return
    const wss = getSheet(st.wb, activeSheetRef.current)
    if (!wss) return
    // 取显示值（公式取缓存值；URN 引用格解析上游首行值，解析失败保留原文）——TSV 与裁尾共用
    const cellTextAt = (r: number, c: number): string => {
      const cell = getCellByCoord(wss, indexToCol(c) + r)
      let v = cell && typeof cell === 'object' && 'value' in cell ? (cell as { value?: unknown }).value : cell
      if (v && typeof v === 'object' && (v as { kind?: string }).kind === 'formula') {
        v = (v as { cachedValue?: unknown }).cachedValue
      }
      // 引用格复制显示值（同网格显示层）：node://… → 上游首行值；解析失败保留 URN 原文不静默丢
      // ★ 批次17 补修：同显示层，走 computeNode 而非 tables[target].fields（汇聚型节点恒取不到）
      if (typeof v === 'string' && v.startsWith('node://') && tables) {
        const u = parseUrn(v)
        if (u?.scheme === 'node') {
          try {
            const c0 = computeNode(u.target, tables, records ?? {})
            const fld = c0.fields.find((f) => f.id === u.locator || f.name === u.locator)
            const rv = c0.rows[0]?.[(fld?.id ?? u.locator) as keyof Row]
            if (rv != null) v = rv
          } catch {
            /* 算不出来 → 保留 URN 原文 */
          }
        }
      }
      return v == null ? '' : String(v)
    }
    // #12/#15 Excel 同款裁尾：整行选区（点行号/Shift 行扩选，选区=全列宽）复制只带到数据边界、
    // 整列选区（点列头）同理——空尾巴不进剪贴板（拖选的普通选区保留原形状不裁）
    let { r1, c1, r2, c2 } = rect
    if (c1 === 1 && c2 >= grid.cols.length) {
      // 行内全扫记最大非空列（不能遇非空就 break——那会把首个非空列右边的数据全裁掉）
      let maxUsed = 0
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          if (cellTextAt(r, c) !== '' && c > maxUsed) maxUsed = c
        }
      }
      c2 = Math.max(c1, maxUsed)
    }
    if (r1 === 1 && r2 >= grid.rows.length) {
      let maxUsedR = 0
      for (let c = c1; c <= c2; c++) {
        for (let r = r1; r <= r2; r++) {
          if (cellTextAt(r, c) !== '' && r > maxUsedR) maxUsedR = r
        }
      }
      r2 = Math.max(r1, maxUsedR)
    }
    const lines: string[] = []
    for (let r = r1; r <= r2; r++) {
      const row: string[] = []
      for (let c = c1; c <= c2; c++) row.push(cellTextAt(r, c))
      lines.push(row.join('\t'))
    }
    const text = lines.join('\n')
    // #15 text/html 双格式（<table>）：粘到 WPS/Excel 保表格；text/plain(TSV) 粘到记事本不受影响。
    // ClipboardItem.write 不可用/被拒时降级 writeText（纯 TSV，老链路）
    const escHtml = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    const html = `<table>${lines
      .map((ln) => `<tr>${ln.split('\t').map((v) => `<td>${escHtml(v)}</td>`).join('')}</tr>`)
      .join('')}</table>`
    // #33A 选择性粘贴：复制时同帧抓格式快照（内存级）——「贴格式/值和格式」的数据源。
    // 外部（WPS/Excel）复制的内容没有快照，选择性粘贴里格式项不可用（报提示）。
    try {
      clipboardFmtRef.current = { snap: fmtSnapshotOfRect(st.wb, activeSheetRef.current, rect), tsv: text }
    } catch {
      clipboardFmtRef.current = null
    }
    const writeBoth = async () => {
      try {
        if (typeof ClipboardItem !== 'undefined' && navigator.clipboard?.write) {
          await navigator.clipboard.write([
            new ClipboardItem({
              'text/plain': new Blob([text], { type: 'text/plain' }),
              'text/html': new Blob([html], { type: 'text/html' }),
            }),
          ])
          return
        }
      } catch {
        /* 写双格式失败 → 降级纯文本 */
      }
      try {
        await navigator.clipboard.writeText(text)
      } catch {
        setErr('复制失败：剪贴板访问被拒绝')
      }
    }
    void writeBoth()
  }
  copySelRef.current = copySelTSV // Ctrl+C 快捷键经此 ref 调用（TDZ 规避，见监听 effect 注释）

  /** #33A/#39 列头右键菜单：排序 / 筛选 / 插删列 / 列宽 / 数据格式 / 清空 */
  const buildHdrItems = (col: string): CtxItem[] => {
    const st = stateRef.current
    if (!st) return []
    const sheet = activeSheetRef.current
    const c = colToIndex(col)
    const maxRow = Math.max(1, grid.rows.length)
    const filtered = filterState?.col === col
    const items: CtxItem[] = [
      { label: `${col} 列 · 升序`, title: '整表按本列排序（行数据整体搬移，公式引用不偏移）', run: () => run(() => sortRowsByCol(st.wb, sheet, c, 'asc')) },
      { label: `${col} 列 · 降序`, run: () => run(() => sortRowsByCol(st.wb, sheet, c, 'desc')) },
      {
        label: filtered ? '重新筛选…' : '筛选…',
        run: () => setFilterDlg({ x: hdrMenu!.x, y: hdrMenu!.y, col }),
      },
    ]
    if (filtered) items.push({ label: '清除筛选', run: () => setFilterState(null) })
    items.push(
      { label: '', sep: true },
      { label: '左侧插入列', run: () => run(() => insertColLeft(st.wb, sheet, c)) },
      { label: `删除 ${col} 列`, run: () => run(() => deleteColAt(st.wb, sheet, c)) },
      {
        label: '清空列内容',
        run: () =>
          run(() => {
            const wss = getSheet(st.wb, sheet)
            if (wss) clearRange(wss, `${col}1:${col}${maxRow}`)
          }),
      },
      { label: '', sep: true },
      {
        label: '列宽…',
        run: () => {
          const v = prompt(`${col} 列列宽（像素，40~600）：`, String(colWidths[col] ?? 108))
          if (v === null) return
          const w = Math.round(Number(v.trim()))
          if (!Number.isFinite(w) || w < 40 || w > 600) {
            setErr('列宽需在 40~600 像素之间')
            return
          }
          setColWidths((prev) => ({ ...prev, [col]: w }))
        },
      },
      {
        // #39 数据格式入口：整列矩形打开格式对话框（fmtOfFirst 读首格初值，四页签复用）
        label: '设置数据格式…',
        run: () => {
          const rect: CellRect = { r1: 1, c1: c, r2: maxRow, c2: c }
          setFmtDlg({ rect, init: fmtOfFirst(st.wb, sheet, rect) })
        },
      },
    )
    return items
  }
  const buildCtxItems = (): CtxItem[] => {
    const rect = selRectRef.current
    const st = stateRef.current
    if (!st || !rect) return [{ label: '先点选一个单元格', sep: false }]
    const sheet = activeSheetRef.current
    const nR = rect.r2 - rect.r1 + 1
    const nC = rect.c2 - rect.c1 + 1
    // #39 行号右键菜单：针对右键命中的那一行（ctxMenu.row）做行级操作
    if (ctxMenu?.kind === 'row' && ctxMenu.row) {
      const r = ctxMenu.row
      return [
        { label: '在上方插入行', run: () => run(() => insertRowAbove(st.wb, sheet, r)) },
        { label: '在下方插入行', run: () => run(() => insertRowAbove(st.wb, sheet, r + 1)) },
        { label: `删除第 ${r} 行`, run: () => run(() => deleteRowAt(st.wb, sheet, r)) },
        {
          label: '清空行内容',
          run: () =>
            run(() => {
              const wss = getSheet(st.wb, sheet)
              const ext = wss ? getValueExtent(wss) : null
              const maxC = Math.max(ext?.maxCol ?? 0, 12)
              if (wss) clearRange(wss, `A${r}:${indexToCol(maxC)}${r}`)
            }),
        },
        { label: '', sep: true },
        {
          label: '行高…',
          run: () => {
            const v = prompt(`第 ${r} 行行高（像素，18~400）：`, String(rowHeightsRef.current[r] ?? 26))
            if (v === null) return
            const h = Math.round(Number(v.trim()))
            if (!Number.isFinite(h) || h < 18 || h > 400) {
              setErr('行高需在 18~400 像素之间')
              return
            }
            setRowHeights((prev) => ({ ...prev, [r]: h }))
          },
        },
      ]
    }
    return [
      { label: '复制', run: copySelTSV },
      {
        // #33A 选择性粘贴（Excel 同款二级）：值 = 原粘贴；转置/格式 = 内存快照链路
        label: '选择性粘贴',
        children: [
          {
            label: '粘贴（数值）',
            run: () => {
              void navigator.clipboard
                ?.readText()
                .then((t) => {
                  if (!t) return
                  run(() => {
                    pasteTSV(st.wb, sheet, `${indexToCol(rect.c1)}${rect.r1}`, t)
                    registerUrnSources(t)
                  })
                })
                .catch(() => setErr('粘贴失败：剪贴板访问被拒绝'))
            },
          },
          {
            label: '转置粘贴',
            title: '行列互换后贴入（从本软件复制的内容才可用）',
            run: () => {
              const clip = clipboardFmtRef.current
              if (!clip) return setErr('转置粘贴只支持从本表格复制的内容（外部来源无内存快照）')
              run(() =>
                pasteGrid(st.wb, sheet, `${indexToCol(rect.c1)}${rect.r1}`, transposeGrid(parseTSVGrid(clip.tsv))),
              )
            },
          },
          {
            label: '值和格式',
            title: '贴入复制时的值与四轴格式（字体/填充/边框/数字）',
            run: () => {
              const clip = clipboardFmtRef.current
              if (!clip) return setErr('「值和格式」只支持从本表格复制的内容（外部来源无内存快照）')
              run(() => {
                pasteGrid(st.wb, sheet, `${indexToCol(rect.c1)}${rect.r1}`, parseTSVGrid(clip.tsv))
                applyFmtSnapshot(st.wb, sheet, `${indexToCol(rect.c1)}${rect.r1}`, clip.snap)
              })
            },
          },
          {
            label: '仅格式',
            title: '只贴字体/填充/边框/数字格式，不动值',
            run: () => {
              const clip = clipboardFmtRef.current
              if (!clip) return setErr('「仅格式」只支持从本表格复制的内容（外部来源无内存快照）')
              run(() => applyFmtSnapshot(st.wb, sheet, `${indexToCol(rect.c1)}${rect.r1}`, clip.snap))
            },
          },
        ],
      },
      { label: '', sep: true },
      {
        // #33A 排序（Excel「数据→排序」同语义）：按选区起始列整表排序，行数据整体搬移写回本体
        label: '排序',
        children: [
          {
            label: `按 ${indexToCol(rect.c1)} 列 · 升序`,
            run: () => run(() => sortRowsByCol(st.wb, sheet, rect.c1, 'asc')),
          },
          {
            label: `按 ${indexToCol(rect.c1)} 列 · 降序`,
            run: () => run(() => sortRowsByCol(st.wb, sheet, rect.c1, 'desc')),
          },
        ],
      },
      { label: `在上方插入行（${nR}）`, run: () => run(() => { for (let i = 0; i < nR; i++) insertRowAbove(st.wb, sheet, rect.r1) }) },
      { label: `在左侧插入列（${nC}）`, run: () => run(() => { for (let i = 0; i < nC; i++) insertColLeft(st.wb, sheet, rect.c1) }) },
      {
        // #33 插入单元格（WPS 同款二级）：腾出选区大小空白，原内容按方向推移（尾部溢出丢弃）
        label: '插入单元格',
        children: [
          { label: '下方单元格下移', run: () => run(() => insertCellShift(st.wb, sheet, rect, 'down')) },
          { label: '右侧单元格右移', run: () => run(() => insertCellShift(st.wb, sheet, rect, 'right')) },
        ],
      },
      { label: `删除所选行（${nR}）`, run: () => run(() => { for (let r = rect.r2; r >= rect.r1; r--) deleteRowAt(st.wb, sheet, r) }) },
      { label: `删除所选列（${nC}）`, run: () => run(() => { for (let c = rect.c2; c >= rect.c1; c--) deleteColAt(st.wb, sheet, c) }) },
      {
        // #33 行高：选区各行统一设值；ag-grid 社区版无原生行高 UI，经 getRowHeight 回调消费
        label: '行高…',
        run: () => {
          const v = prompt(`行高（像素，作用于选区 ${nR} 行，18~400）：`, String(rowHeightsRef.current[rect.r1] ?? 26))
          if (v === null) return
          const h = Math.round(Number(v.trim()))
          if (!Number.isFinite(h) || h < 18 || h > 400) {
            setErr('行高需在 18~400 像素之间')
            return
          }
          setRowHeights((prev) => {
            const next = { ...prev }
            for (let r = rect.r1; r <= rect.r2; r++) next[r] = h
            return next
          })
        },
      },
      {
        // #33 列宽：选区各列统一设值（与拖拽列宽同走 colWidths，colDefs 重建不回弹）
        label: '列宽…',
        run: () => {
          const v = prompt(`列宽（像素，作用于选区 ${nC} 列，40~600）：`, String(colWidths[indexToCol(rect.c1)] ?? 108))
          if (v === null) return
          const w = Math.round(Number(v.trim()))
          if (!Number.isFinite(w) || w < 40 || w > 600) {
            setErr('列宽需在 40~600 像素之间')
            return
          }
          setColWidths((prev) => {
            const next = { ...prev }
            for (let c = rect.c1; c <= rect.c2; c++) next[indexToCol(c)] = w
            return next
          })
        },
      },
      {
        label: '清除内容',
        run: () =>
          run(() => {
            const wss = getSheet(st.wb, sheet)
            if (wss) clearRange(wss, `${indexToCol(rect.c1)}${rect.r1}:${indexToCol(rect.c2)}${rect.r2}`)
          }),
      },
      {
        // #33 第二块：设置单元格格式对话框（字体/边框/填充/数字四页签）
        label: '设置单元格格式…',
        run: () => {
          const init = fmtOfFirst(st.wb, sheet, rect)
          setFmtDlg({ rect: { ...rect }, init })
        },
      },
      { label: '', sep: true },
      { label: '合并单元格', run: () => run(() => mergeSel(st.wb, sheet, rect)) },
      { label: '取消合并', run: () => run(() => unmergeSel(st.wb, sheet, rect)) },
      { label: '', sep: true },
      { label: '对齐：左', run: () => run(() => alignSel(st.wb, sheet, rect, 'left')) },
      { label: '对齐：居中', run: () => run(() => alignSel(st.wb, sheet, rect, 'center')) },
      { label: '对齐：右', run: () => run(() => alignSel(st.wb, sheet, rect, 'right')) },
      { label: '', sep: true },
      { label: '数字：千分位', run: () => run(() => numberFmtSel(st.wb, sheet, rect, '#,##0.00')) },
      { label: '数字：百分比', run: () => run(() => numberFmtSel(st.wb, sheet, rect, '0.00%')) },
      { label: '数字：货币', run: () => run(() => numberFmtSel(st.wb, sheet, rect, '¥#,##0.00')) },
      { label: '数字：常规', run: () => run(() => numberFmtSel(st.wb, sheet, rect, 'General')) },
      { label: '', sep: true },
      { label: '字体加粗', title: 'v1 提示：会按 OOXML 规则整体替换字体样式', run: () => run(() => fontSel(st.wb, sheet, rect, { bold: true })) },
      { label: '取消加粗', run: () => run(() => fontSel(st.wb, sheet, rect, { bold: false })) },
      { label: '斜体', run: () => run(() => fontSel(st.wb, sheet, rect, { italic: true })) },
      { label: '下划线', run: () => run(() => fontSel(st.wb, sheet, rect, { underline: true })) },
    ]
  }

  return (
    <div className="wb">
      {/* ── WPS 风格功能区（与 Word 面板共用 Ribbon 壳，样式一致）── */}
      <Ribbon
        tabs={ribbonTabs}
        /* 行1 快捷区方形撤销按钮（WPS 快速访问栏同位；无改动可撤时置灰） */
        quick={
          <button
            className="rb-quickbtn"
            title="撤销上一步改动"
            disabled={undoStack.current.length <= 1}
            onClick={doUndo}
          >
            ↩
          </button>
        }
      />

      {/* 主体：左 = 属性引用栏（portal 到 body，贴屏幕左缘）· 右 = 工作表区；引用连线为 fixed 全屏 SVG */}
      <div className="wb-main">
      <div
        className="wb-col"
        style={{
          marginLeft: maximized && refPanelOpen && panelPos ? 230 : 0,
          marginRight: maximized && rightOpen && rightPos ? 230 : 0,
        }}
      >
        {/* 孤儿引用提示条已改悬浮在网格底部滚动条上方（wb-grid 内 absolute，同 wb-err-inline），不再占文档流推挤布局 */}
      <div className="wb-tabs">
        {sheets.map((s) => (
          <button
            key={s}
            className={`wb-tab${s === active ? ' on' : ''}`}
            onClick={() => setActive(s)}
            onDoubleClick={() =>
              run(() => {
                const nn = prompt('重命名工作表：', s)
                if (nn && nn.trim()) {
                  renameSheetTitled(state.wb, s, nn.trim())
                  setActive(nn.trim())
                }
              })
            }
            title="单击切换 · 双击重命名"
          >
            {s}
          </button>
        ))}
        {!readOnly && (
          <button
            className="bar-btn"
            title="新增工作表"
            onClick={() =>
              run(() => {
                const nn = prompt('新工作表名称：', 'Sheet' + (sheets.length + 1))
                if (nn && nn.trim()) {
                  addSheetTitled(state.wb, nn.trim())
                  setActive(nn.trim())
                }
              })
            }
          >
            ＋
          </button>
        )}
        {!readOnly && sheets.length > 1 && (
          <button
            className="bar-btn"
            title="删除当前工作表"
            onClick={() => {
              if (!confirm(`删除工作表「${active}」？`)) return
              run(() => {
                removeSheetTitled(state.wb, active)
                setActive(sheets.find((s) => s !== active) ?? '')
              })
            }}
          >
            删表
          </button>
        )}
        <span className="wb-sp" />
        {(namedHere?.named.size ?? 0) > 0 && (
          <span className="wb-chip wb-chip--named">命名区域 {namedHere?.named.size}</span>
        )}
        {wbTables.length > 0 && <span className="wb-chip wb-chip--table">表格 {wbTables.length}</span>}
        <span className="wb-stat">
          {ext ? `${ext.maxRow} 行 × ${ext.maxCol} 列` : '空表'} · 非空 {nonEmpty}
          {merges.length ? ` · 合并 ${merges.length} 处` : ''}
        </span>
      </div>

      {/* 错误/操作提示已改为悬浮在网格底部滚动条上方（wb-grid 内 absolute），不再占文档流推挤布局 —— 用户 2026-09-26：
          「这个提示放在excel下方滚动条上面悬浮，不要在原来地方，因为提示让excel的界面位置发生变化」*/}

      {state.applied.length > 0 && (
        <details className="wb-note" open>
          <summary>为让解析器读懂，打开时做了 {state.applied.length} 处归一化</summary>
          <div className="wb-note-body">
            {state.applied.map((a, i) => (
              <div key={i}>· {a}</div>
            ))}
            <div className="wb-note-foot">
              只针对 Excel/WPS 写出的、schema 允许但解析器不接受的构造替换，语义不变。
            </div>
          </div>
        </details>
      )}

      {/* 公式栏：编辑态点格子插引用；输入 = 触发函数补全（功能区收起时公式栏常显，WPS 同款） */}
      <div
        className="wb-formula"
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: 6,
          padding: '4px 8px',
          borderBottom: '1px solid var(--ed-border2)',
          background: 'var(--ed-bg1)',
          position: 'relative',
        }}
      >
        <span
          style={{
            minWidth: 64, textAlign: 'center', fontSize: 12, color: 'var(--ed-text3)',
            border: '1px solid var(--ed-border2)', borderRadius: 4, padding: '3px 6px', background: 'var(--ed-bg2)',
          }}
        >
          {selInfo?.coord || '—'}
        </span>
        <span style={{ fontStyle: 'italic', fontSize: 12, color: 'var(--ed-gold)', fontWeight: 600 }}>fx</span>
        {/* 备忘 · 对照体系：input 含 node:// 时文字/caret 透明，上层镜像画胶囊+假光标/假选区；
            data-mirror 供 CSS ::selection 透明选择（选中态原生 HighlightText 会强制显形文字叠在胶囊下——用户 2026-09-27 截图实锤「上下两层」）；
            spellCheck 关掉拼写红线（URN 永远拼不对）；点击/键盘全在 input 上（mirror pointer-events:none） */}
        <div style={{ position: 'relative', flex: 1, minWidth: 0 }} data-fx-wrap>
          <input
            ref={fxInputRef}
            spellCheck={false}
            data-mirror={fxMirrorOn ? 'on' : 'off'}
            style={{
              display: 'block', boxSizing: 'border-box', width: '100%', fontSize: 12.5, fontFamily: 'Consolas, monospace', lineHeight: '18px',
              color: fxMirrorOn ? 'transparent' : 'var(--ed-text)', caretColor: fxMirrorOn ? 'transparent' : 'var(--ed-text)',
              background: 'var(--ed-bg2)', border: '1px solid var(--ed-border2)',
              borderRadius: 4, padding: '3px 8px',
            }}
            value={fxDraft}
            readOnly={readOnly}
            placeholder="=开头出公式提示（格内输入=同样生效）· 点其他格子插入引用 · 回车写入"
            onChange={(e) => onFxChange(e.target.value)}
            onCompositionStart={() => setFxComposing(true)}
            onCompositionEnd={() => setFxComposing(false)}
            onMouseDown={(e) => {
              // 备忘 2026-09-27 ⑥ · 胶囊=一个字符：点胶囊 = 选中整个 token（可整选后打字替换/退格整删）；
              // 点坐标引用 = 整实例选中；其他 = 光标定位。拖选全接管（锚点记在 fxDragRef，mousemove 几何走镜像坐标）——
              // 原生拖选按透明原文布局定位，宽差致高亮与鼠标错位（截图实锤）。换参菜单仍在 onDoubleClick。
              const mirror = fxMirrorRef.current
              if (!mirror || !fxMirrorOn) return
              const el = fxInputRef.current
              // WPS 语义：点公式里的坐标引用 → 网格对应格亮显 + 整实例区间选中（打字/点格直接替换）；点其他位置 → 清亮显
              const hc = hitFxMirrorCoord(mirror, e.clientX, e.clientY)
              setGlowCell(hc?.coord ?? null)
              // 引用实例先于 off 计算：点在引用 span 上时 caretRangeFromPoint 穿透 pointer-events:none 命中 input → off=null，
              // 整实例选中不能依赖 off（纯几何命中就够）——否则漏 preventDefault，原生定位接管（实测 ss=2 光标落引用中间）
              const inst = hc ? calcRefInstances(fxDraft)[hc.idx] : null
              const cap = hitFxMirrorCapsule(mirror, e.clientX, e.clientY)
              const off = hitFxMirrorOffset(mirror, e.clientX, e.clientY)
              if (off == null && !inst && !cap) return
              e.preventDefault()
              el?.focus()
              // 点中引用实例 → 选中整个实例区间（WPS：点公式里的引用 = 选中该参数，打字/点格直接替换）
              if (inst) {
                el?.setSelectionRange(inst.s, inst.e)
                fxPickRef.current = { s: inst.s, e: inst.e }
                setFxPick({ s: inst.s, e: inst.e })
                setPickCell(inst.coord)
                fxDragRef.current = null
                return
              }
              if (cap) {
                // 点胶囊 → 选中整 token（一个字符语义）；锚点记就近边界（拖走时从边界扩展，tokenAtomicSel 吸附整吞）
                const off2 = off ?? cap.start
                el?.setSelectionRange(cap.start, cap.end)
                fxDragRef.current = off2
              } else if (off != null) {
                el?.setSelectionRange(off, off)
                fxDragRef.current = off
              }
              // 拖选结束（任意处松开）→ 清锚点（拖出 fx 栏松开也要清，挂 document once）
              document.addEventListener('mouseup', () => { fxDragRef.current = null }, { once: true })
            }}
            onDoubleClick={(e) => {
              // 备忘 · 批次 8-b（单击迁移）：双击胶囊 → 弹同节点换参菜单（胶囊锁死不可文本编辑，单击让位给光标定位 —— 2026-09-27 ④）
              const mirror = fxMirrorRef.current
              if (!mirror || !fxMirrorOn) return
              const hit = hitFxMirrorCapsule(mirror, e.clientX, e.clientY)
              if (!hit) return
              e.preventDefault()
              const u = parseUrn(hit.urn)
              const t = u?.scheme === 'node' ? tables?.[u.target] : null
              if (!t || !t.fields.length) return
              const fr = (e.currentTarget as HTMLElement).closest('.wb-formula')?.getBoundingClientRect()
              const left = fr ? Math.min(Math.max(hit.left - fr.left, 0), fr.width - 190) : 0
              setFxMenu({ start: hit.start, end: hit.end, nodeId: t.id, fid: u?.locator ?? '', left })
            }}
            onMouseMove={(e) => {
              // 悬停胶囊 → pointer + 锁定提示（mirror 是 pointer-events:none，光标只能挂在 input 上动态切）
              const mirror = fxMirrorRef.current
              if (!mirror || !fxMirrorOn) return
              const el = e.currentTarget as HTMLInputElement
              // 备忘 2026-09-27 ⑥ · 拖选接管：按住拖动 = 起点到当前点做 tokenAtomicSel（相交 token 整吞），
              // 几何全走镜像坐标——与鼠标可视位置一致，根治原生拖选的宽差错位
              const drag = fxDragRef.current
              if (drag != null) {
                const off = hitFxMirrorOffset(mirror, e.clientX, e.clientY)
                if (off != null) {
                  const a = Math.min(drag, off)
                  const b = Math.max(drag, off)
                  const { s: s3, e: e3 } = tokenAtomicSel(a, b, fxTokenRanges(fxDraft, resolvedFields))
                  if (el.selectionStart !== s3 || el.selectionEnd !== e3) el.setSelectionRange(s3, e3)
                }
              }
              const hit = hitFxMirrorCapsule(mirror, e.clientX, e.clientY)
              const cur = hit ? 'pointer' : 'text'
              if (el.style.cursor !== cur) el.style.cursor = cur
              const tip = hit ? '引用胶囊已锁定 · 双击选择同节点的其他参数' : ''
              if (el.title !== tip) el.title = tip
              // 备忘 2026-09-27 ⑤：悬停胶囊 → 画布来源节点边框亮显（store 桥）+ 表格里挂该 URN 的引用格亮显；
              // 悬停 A1 引用 → 对应格亮显；移出全清（onMouseLeave）。diff 变化才 refreshCells。
              const u = hit ? parseUrn(hit.urn) : null
              const nid = u?.scheme === 'node' ? u.target : null
              if (fxHoverNodeRef.current !== nid) {
                fxHoverNodeRef.current = nid
                setFxHoverNode(nid)
              }
              const hc = hit ? null : hitFxMirrorCoord(mirror, e.clientX, e.clientY)
              const coord = hc?.coord ?? null
              if (hoverGlowRef.current !== coord) {
                hoverGlowRef.current = coord
                gridApiRef.current?.refreshCells({ force: true })
              }
              const urn = hit?.urn ?? null
              if (hoverUrnKeyRef.current !== urn) {
                hoverUrnKeyRef.current = urn
                // URN 变化 → 扫当前 sheet 挂该 URN 的引用格（与 refdim 外部并入同款扫描，仅在胶囊间切换时执行一次）
                let cells: Set<string> | null = null
                if (urn) {
                  try {
                    const st = stateRef.current
                    const wss = st ? getSheet(st.wb, activeSheetRef.current) : null
                    if (wss) {
                      const ext = getValueExtent(wss)
                      if (ext) {
                        cells = new Set()
                        for (let rr = ext.minRow; rr <= ext.maxRow; rr++) {
                          for (let cc = ext.minCol; cc <= ext.maxCol; cc++) {
                            const cell = getCellByCoord(wss, indexToCol(cc) + rr)
                            const v = cell && typeof cell === 'object' ? (cell as { value?: unknown }).value : null
                            if (v === urn) cells.add(indexToCol(cc) + rr)
                          }
                        }
                      }
                    }
                  } catch {
                    cells = null
                  }
                }
                hoverUrnRef.current = cells
                gridApiRef.current?.refreshCells({ force: true })
              }
            }}
            onMouseLeave={() => {
              // 悬停离开 → 清 hover 亮显与画布桥（点击态 glowCell/pickCell 是独立 ref/state，不受影响）
              if (hoverGlowRef.current != null) {
                hoverGlowRef.current = null
                gridApiRef.current?.refreshCells({ force: true })
              }
              if (hoverUrnRef.current != null || hoverUrnKeyRef.current != null) {
                hoverUrnRef.current = null
                hoverUrnKeyRef.current = null
                gridApiRef.current?.refreshCells({ force: true })
              }
              if (fxHoverNodeRef.current != null) {
                fxHoverNodeRef.current = null
                setFxHoverNode(null)
              }
            }}
            onPaste={(e) => {
              // 粘贴选区触到 token → 相交 token 整体吞入选区再插入（原子吞语义）；粘贴文本里的 URN 顺手转可读（URN 不进编辑器）
              const el = e.currentTarget as HTMLInputElement
              const fs = el.selectionStart ?? 0
              const fe = el.selectionEnd ?? fs
              const ranges = fxTokenRanges(fxDraft, resolvedFields)
              const txt0 = e.clipboardData.getData('text')
              const txt = urnToReadable(txt0, resolvedFields)
              const hitTok = ranges.some((r) => fs < r.end && fe > r.start)
              if (!hitTok && txt === txt0) return
              e.preventDefault()
              let s2 = fs
              let e2 = fe
              for (const r of ranges) {
                if (fs < r.end && fe > r.start) { s2 = Math.min(s2, r.start); e2 = Math.max(e2, r.end) }
              }
              onFxChange(fxDraft.slice(0, s2) + txt + fxDraft.slice(e2))
              const p0 = s2 + txt.length
              requestAnimationFrame(() => { el.focus(); el.setSelectionRange(p0, p0) })
            }}
            onScroll={() => { if (fxMirrorRef.current) fxMirrorRef.current.scrollLeft = fxInputRef.current?.scrollLeft ?? 0 }}
            onSelect={() => {
              // 镜像层滚动同步 + 补全跟随（原 onSelect 逻辑）
              if (fxMirrorRef.current) fxMirrorRef.current.scrollLeft = fxInputRef.current?.scrollLeft ?? 0
              const t = fxInputRef.current
              if (!t) return
              // 兜底：无选区时光标落进 token 内部（点在透明文本上）→ 推到胶囊右缘（命中胶囊本体的点击已在 onMouseDown 拦下）
              const s0 = t.selectionStart ?? 0
              const e0 = t.selectionEnd ?? 0
              if (s0 === e0) {
                const r = fxTokenRanges(t.value, resolvedFields).find((q) => s0 > q.start && s0 < q.end)
                if (r) t.setSelectionRange(r.end, r.end)
              } else {
                // 备忘 2026-09-27 ⑥ · 选区吸附：键盘 shift 扩选等未接管路径，选区与 token 相交 → 吞入整 token
                // （幂等：吸附后边界恰为 token 端点，再算不变，无死循环）
                const rr = fxTokenRanges(t.value, resolvedFields)
                if (rr.some((q) => s0 < q.end && e0 > q.start)) {
                  const { s: s3, e: e3 } = tokenAtomicSel(s0, e0, rr)
                  if (s3 !== s0 || e3 !== e0) t.setSelectionRange(s3, e3)
                }
              }
              // WPS 语义：光标移出活跃引用区末端 = 引用定型（点格桥会先把光标摆到区末端，不触发清）
              const pk = fxPickRef.current
              if (pk && (s0 !== pk.e || e0 !== pk.e)) {
                fxPickRef.current = null
                setFxPick(null)
              }
              updateCompl(t.value, t.selectionStart ?? 0)
            }}
            onFocus={() => {
              setFxEditing(true)
              // 对照体系：聚焦即刻同步光标（programmatic focus + selection 未变时 selectionchange 不触发，缺这步首帧无光标）
              const el = fxInputRef.current
              setFxCaret({ s: el?.selectionStart ?? 0, e: el?.selectionEnd ?? 0 })
            }}
            onBlur={() => {
              setFxCaret(undefined) // 对照体系：失焦消假光标（镜像层不再画 caret）
              setTimeout(() => {
                if (!fxInputRef.current?.matches(':focus')) {
                  setFxEditing(false)
                  setCompl(null)
                }
              }, 150)
            }}
            onKeyDown={(e) => {
              // IME 组合中（拼音/选词/候选导航）放行给输入法：回车=上屏、Escape=取消组合、上下键=候选导航（用户 2026-09-27：中文输公式回车被误当提交）
              if (e.nativeEvent.isComposing || e.key === 'Process' || e.keyCode === 229) return
              // ── 胶囊锁死（锁死 token 三件套：打字/粘贴原子吞 / 退格原子删 / 方向键整跳）──
              // 语义：胶囊不可切半、不可原地改字符；选区相交的打字 = 相交 token 整体吞入选区再插字符（全选重写一步直达）
              const fxEl = e.currentTarget as HTMLInputElement
              const fs = fxEl.selectionStart ?? 0
              const fe = fxEl.selectionEnd ?? fs
              const ranges = fxTokenRanges(fxDraft, resolvedFields)
              const printable = e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey
              if (printable && fe > fs && ranges.some((r) => fs < r.end && fe > r.start)) {
                e.preventDefault()
                let s2 = fs
                let e2 = fe
                for (const r of ranges) {
                  if (fs < r.end && fe > r.start) { s2 = Math.min(s2, r.start); e2 = Math.max(e2, r.end) }
                }
                onFxChange(fxDraft.slice(0, s2) + e.key + fxDraft.slice(e2))
                const p0 = s2 + 1
                requestAnimationFrame(() => { const el = fxInputRef.current; el?.focus(); el?.setSelectionRange(p0, p0) })
                return
              }
              // 光标落在 token 内部打字（无选区）→ 锁死不动（想动胶囊：点它换参，或退格整体删）
              if (printable && fs === fe) {
                const inside = ranges.find((r) => fs > r.start && fs < r.end)
                if (inside) { e.preventDefault(); return }
              }
              // 备忘 2026-09-27 ⑥ · 选区删除原子吞：选区（拖选/shift 扩选/Ctrl+A）相交 token → 吞整 token 再删。
              // 原生删会拆 token 留残段（截图实锤 `://param_…#price` 泄漏）——根治入口之一
              if ((e.key === 'Backspace' || e.key === 'Delete') && fe > fs && ranges.some((r) => fs < r.end && fe > r.start)) {
                e.preventDefault()
                const { s: s2, e: e2 } = tokenAtomicSel(fs, fe, ranges)
                onFxChange(fxDraft.slice(0, s2) + fxDraft.slice(e2))
                requestAnimationFrame(() => { const el2 = fxInputRef.current; el2?.focus(); el2?.setSelectionRange(s2, s2) })
                return
              }
              if ((e.key === 'Backspace' || e.key === 'Delete') && fs === fe) {
                const r = ranges.find((q) => (e.key === 'Backspace' ? fs > q.start && fs <= q.end : fs >= q.start && fs < q.end))
                if (r) {
                  e.preventDefault()
                  onFxChange(fxDraft.slice(0, r.start) + fxDraft.slice(r.end))
                  const p = r.start
                  requestAnimationFrame(() => { const el = fxInputRef.current; el?.focus(); el?.setSelectionRange(p, p) })
                  return
                }
              }
              if ((e.key === 'ArrowLeft' || e.key === 'ArrowRight') && fs === fe) {
                // 方向键把胶囊当整 token 跳过（左跳到 token 左缘 / 右跳到右缘，不进入内部）
                const r = ranges.find((q) => (e.key === 'ArrowLeft' ? fs > q.start && fs <= q.end : fs >= q.start && fs < q.end))
                if (r) {
                  e.preventDefault()
                  const p = e.key === 'ArrowLeft' ? r.start : r.end
                  fxEl.setSelectionRange(p, p)
                  return
                }
              }
              if (fxMenu && e.key === 'Escape') { e.preventDefault(); setFxMenu(null); return }
              if (compl) {
                if (e.key === 'ArrowDown') { e.preventDefault(); setCompl({ ...compl, idx: (compl.idx + 1) % compl.items.length }); return }
                if (e.key === 'ArrowUp') { e.preventDefault(); setCompl({ ...compl, idx: (compl.idx - 1 + compl.items.length) % compl.items.length }); return }
                if (e.key === 'Tab') { e.preventDefault(); acceptCompl(); return }
                if (e.key === 'Escape') { e.preventDefault(); setCompl(null); return }
              }
              if (e.key === 'Enter' && !readOnly) void applyFormulaBar()
              if (e.key === 'Escape') {
                setFxEditing(false)
                setFxDraft(selInfo?.formula ? '=' + urnToReadable(selInfo.formula, resolvedFieldsRef.current) : (selInfo?.display ?? ''))
                fxInputRef.current?.blur()
              }
            }}
          />
          {fxMirrorOn && (
            <div
              ref={fxMirrorRef}
              className="wb-fx-mirror"
              style={{
                position: 'absolute', inset: 0, padding: '3px 9px', borderRadius: 4,
                pointerEvents: 'none', overflow: 'hidden', whiteSpace: 'pre',
              }}
              dangerouslySetInnerHTML={{ __html: fxMirrorHtml }}
            />
          )}
        </div>
        {compl && (
          <div
            className="wb-fn-drop"
            style={{
              position: 'absolute', top: '100%', left: 78, width: 280, zIndex: 70,
              background: 'var(--ed-bg2)', border: '1px solid var(--ed-border2)', borderRadius: 6,
              maxHeight: 240, overflow: 'auto',
              // 纯提示层（Excel 同款）：不拦截鼠标 —— 点格子永远直达（插引用），
              // 函数用 ↑↓ + Tab 插入；鼠标选函数走「插入/函数库」
              pointerEvents: 'none',
            }}
          >
            <div style={{ padding: '3px 12px 4px', fontSize: 10, color: 'var(--ed-text3)', borderBottom: '1px solid var(--ed-border2)' }}>
              ↑↓ 选择 · Tab 插入（鼠标选函数去「插入」菜单）
            </div>
            {compl.items.map((f, i) => (
              <div
                key={f}
                style={{
                  padding: '5px 12px', fontSize: 12, fontFamily: 'Consolas, monospace',
                  color: i === compl.idx ? 'var(--ed-bg3)' : 'var(--ed-text)',
                  background: i === compl.idx ? 'var(--ed-accent)' : 'transparent', cursor: 'pointer',
                }}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => acceptCompl()}
                onMouseEnter={() => setCompl({ ...compl, idx: i })}
              >
                {f}(…)
              </div>
            ))}
          </div>
        )}
        {/* fx 胶囊换参菜单：胶囊是锁死 token（引用端不可改），唯一修改 = 选同节点的不同参数（用户 2026-09-26 拍板） */}
        {fxMenu && (() => {
          // ★ 批次17：用 resolvedFields（计算后字段）而非 tables（schema 静态字段）——
          // 汇聚型节点（sort/map/output）自身 t.fields 为空，旧口径换参菜单是空的
          const t = resolvedFields?.[fxMenu.nodeId]
          if (!t) return null
          return (
            <div className="ref-chip-menu wb-fx-menu" style={{ left: fxMenu.left, right: 'auto' }}>
              {t.fields.map((f) => (
                <button
                  key={f.id}
                  className={`ref-chip-item${f.id === fxMenu.fid ? ' on' : ''}`}
                  onClick={() => {
                    // fx 可读化：换参 = token 替换成「表名·新参数名」（编辑串全程可读）
                    const tok = `${t.name}·${f.name}`
                    const nv = fxDraft.slice(0, fxMenu.start) + tok + fxDraft.slice(fxMenu.end)
                    setFxMenu(null)
                    if (fxEditing) {
                      // 编辑中：只替换草稿（回车确认），光标落回新 token 右缘
                      onFxChange(nv)
                      const p = fxMenu.start + tok.length
                      requestAnimationFrame(() => { const el = fxInputRef.current; el?.focus(); el?.setSelectionRange(p, p) })
                    } else {
                      // 稳态（选中格未编辑）：原位换参直接生效（接棒已删的信息栏 chip 行为）
                      // 落格走提交边界：可读 token → URN（存储层永远是 URN）
                      const cellVal = readableToUrn(nv.startsWith('=') ? nv.slice(1) : nv, resolvedFields)
                      const st = stateRef.current
                      const wss = st ? getSheet(st.wb, activeSheetRef.current) : null
                      if (wss) setCellByCoord(wss, selCoord, cellVal)
                      setFxDraft(nv)
                      void Promise.resolve(commitRef.current?.()).then(() => gridApiRef.current?.refreshCells({ force: true }))
                    }
                  }}
                >
                  {f.name}
                  <span className="ref-chip-sample">{String(records?.[fxMenu.nodeId]?.[0]?.[f.id as keyof Row] ?? '')}</span>
                </button>
              ))}
            </div>
          )
        })()}
        {/* 公式栏行尾原「外部引用 chip」已删（用户 2026-09-26：fx 栏里已有胶囊，不要后面那个重复的）——
            引用显示走 fx 镜像胶囊，稳态换参 = 点 fx 胶囊弹菜单选参数直接生效（fxMenu onClick 稳态分支） */}
      </div>

      {/* 工具栏已并入 Ribbon「开始」页签 */}

      <div
        className={`wb-grid ${theme === 'light' ? 'ag-theme-quartz' : 'ag-theme-quartz-dark'}`}
        onMouseLeave={() => {
          // 备忘 · 真正移出网格才回退阅读模式十字到选中格（格间切换不清，防闪回）
          hoverCellRef.current = null
          if (hlCrossRef.current) gridApiRef.current?.refreshCells({ force: true })
        }}
      >
        <AgGridReact
          rowData={grid.rows}
          columnDefs={colDefs}
          defaultColDef={{ resizable: true, sortable: false, filter: false }}
          // 失焦自动关编辑器必须显式关死：FCE mount 瞬间 focus 时序不稳（焦点会短暂落回 BODY），
          // 若此开关开着 → 编辑器刚开就被判失焦自动 stopEditing → 单击+Enter 进编辑又闪退、引用格被幽灵提交破坏
          stopEditingWhenCellsLoseFocus={false}
          onGridReady={(p: GridReadyEvent) => {
            gridApiRef.current = p.api
            // 捕获滚动容器作为选区外轮廓的 portal 挂载点（content 坐标随滚动原生移动）
            setGridVpEl(document.querySelector<HTMLElement>('.ag-body-viewport'))
          }}
          onColumnResized={onColResized}
          onCellMouseDown={(e) => {
            // 拖动框选起点：按住左键划过即扩展（onCellMouseOver 里续，全局 mouseup 收）
            if (readOnly) return
            if ((e.event as MouseEvent | undefined)?.button !== 0) return
            const col = e.colDef.field
            const row = (e.data as Record<string, number> | undefined)?.__row
            if (!col || !row || !/^[A-Z]{1,3}$/.test(col)) return
            dragSelRef.current = true
            // Shift 按下 = 扩展模式：保留既有锚点（否则 mousedown 先于 click 把锚重置成本格，
            // click 里的「保留锚点」保留的是被污染的锚 → Shift 扩选永远退化成单格，实测 R5）
            if (!(lastDownShiftRef.current && selAnchor)) {
              setSelAnchor({ r: row, c: colToIndex(col) })
            }
            setSelCoord(col + row)
          }}
          onCellMouseOver={(e) => {
            const col = e.colDef.field
            const row = (e.data as Record<string, number> | undefined)?.__row
            if (!col || !row || !/^[A-Z]{1,3}$/.test(col)) return
            // 拖动框选：按住左键划过的格子并入选区
            if (dragSelRef.current && !readOnly) {
              setSelCoord(col + row)
            }
            // 备忘 · 填充柄的悬浮检测在 gridMouseMoveRef（真实 mousemove，同格移动也能触发）
            // 阅读模式：悬停格的行列十字跟随（ref 驱动 + 只刷可见格，不触发 React 重渲染）
            if (!hlCrossRef.current || readOnly) return
            const nc = { r: row, c: colToIndex(col) }
            const last = hoverCellRef.current
            if (last && last.r === nc.r && last.c === nc.c) return
            hoverCellRef.current = nc
            gridApiRef.current?.refreshCells({ force: true })
          }}
          // 填充柄清理也在 gridMouseMoveRef；Out 不做处理（mousemove 已覆盖格间切换与离开网格）
          onCellContextMenu={(e) => {
            // 备忘 · WPS 同款：右键弹自定义单元格菜单。
            // 右键格在现有选区内 → 保留多选（用户 2026-09-26：右键不能把多选打回单格，否则复制只拿到一格）；
            // 在选区外 → 该格成为新选区
            const ev = e.event as MouseEvent | undefined
            const col = e.colDef.field
            const row = (e.data as Record<string, number> | undefined)?.__row
            if (col && row && /^[A-Z]{1,3}$/.test(col) && !readOnly) {
              const rect = selRectRef.current
              const cIdx = colToIndex(col)
              const inSel = !!rect && row >= rect.r1 && row <= rect.r2 && cIdx >= rect.c1 && cIdx <= rect.c2
              if (!inSel) {
                setSelAnchor({ r: row, c: cIdx })
                setSelCoord(col + row)
              }
            }
            if (ev) {
              ev.preventDefault()
              // #39 行号右键 = 行级操作菜单（插入/删除/清空/行高），与单元格菜单分流
              if (col === '__rownum' && row) setCtxMenu({ x: ev.clientX, y: ev.clientY, kind: 'row', row })
              else setCtxMenu({ x: ev.clientX, y: ev.clientY })
            }
          }}
          // 列头右键菜单走 document capture 监听（下方 effect）：本版 ag-grid 的
          // ColumnHeaderContextMenuEvent 不带原生 event（v33 移除），拿不到弹窗坐标
          onColumnHeaderClicked={(e) => {
            // 点列头 = 选中整列（锚=列顶、尾=列底，选区表达为矩形 → 删列/列宽/格式全链路复用）
            // Shift+点列头 = 从锚点列扩到当前列（header 事件不带鼠标信息，shiftKey 取自 document capture）
            if (readOnly) return
            const colApi = e.column as { getColId?: () => string }
            const colId = colApi.getColId?.()
            if (!colId || !/^[A-Z]{1,3}$/.test(colId)) return
            const c = colToIndex(colId)
            const r2 = Math.max(1, grid.rows.length)
            const shift = lastDownShiftRef.current
            const anchorC = shift && colAnchorRef.current != null ? colAnchorRef.current : c
            colAnchorRef.current = anchorC
            setSelAnchor({ r: 1, c: anchorC })
            setSelCoord(colId + r2)
          }}
          // #33 行高：默认 26，「行高…」菜单设置过的行用自定义值（state 变化经 resetRowHeights effect 重算）
          getRowHeight={(p: { data?: Record<string, unknown> | null }) => rowHeightsRef.current[(p.data?.__row as number) ?? -1] ?? 26}
          headerHeight={28}
          onCellValueChanged={(e) => void onCellChanged(e)}
          onCellClicked={(e: CellClickedEvent) => {
            // 公式态的格子点击已被 capture 拦截器处理（插引用），这里只消费防双插标记
            if (suppressClickRef.current) {
              suppressClickRef.current = false
              return
            }
            const col = e.colDef.field
            const row = (e.data as Record<string, number> | undefined)?.__row
            if (!col || !row || col === '__row') return
            // 行号列点击 = 选中整行（用户 2026-09-26：行要能整体选中）；锚点行首 + 终点行尾 → selRect 派生整行矩形
            // Shift+点行号 = 从锚点行扩到当前行（多行整块，Excel 同款）：保留锚点只推 selCoord，矩形自动成块
            if (col === '__rownum') {
              if (readOnly) return
              const maxC = grid.cols.length
              if (maxC < 1) return
              // shiftKey 走 document capture 记录（e.event 不可靠，同列头）
              const shift = lastDownShiftRef.current
              if (!(shift && selAnchor)) setSelAnchor({ r: row, c: 1 })
              setSelCoord(indexToCol(maxC) + row)
              return
            }
            const nc = { r: row, c: colToIndex(col) }
            // 公式编辑态：点格子 = 插入引用（不切换选中）—— Excel 的核心交互
            // 焦点判定同 fxActiveRef：fxEditing 残焦时不误插（见 onDocMouseDown 处注释）
            if (fxEditing && fxDraft.startsWith('=') && !readOnly && document.activeElement === fxInputRef.current) {
              insertCellRef(col + row)
              return
            }
            // shiftKey 走 document capture 记录：ag-grid cellClicked 的 e.event 不带鼠标信息（同列头，实测 R5 假红揪出）
            const shift = lastDownShiftRef.current
            if (shift && selAnchor) {
              // Shift+点击：保留锚点，扩展矩形
            } else {
              setSelAnchor(nc)
            }
            setSelCoord(col + row)
          }}
          getRowId={(p) => String((p.data as Record<string, unknown>).__row)}
          suppressMovableColumns
        />
        {/* 错误/操作提示：悬浮在网格底部滚动条上方（不占文档流，Excel 界面不再被推挤） */}
        {err && (
          <div className="wb-err-inline">
            <span>{err}</span>
            <button className="bar-btn" onClick={() => setErr('')}>
              知道了
            </button>
          </div>
        )}
        {/* 孤儿引用提示：同款悬浮（用户 2026-09-26 拍板与错误提示一致，不推挤布局） */}
        {orphanRefs.length > 0 && (
          <div className="wb-orphanbar">
            <span>
              有 {orphanRefs.length} 个引用格的连线已在画布删除（如 {orphanRefs[0].sheet}!{orphanRefs[0].coord}），值已保留。
            </span>
            <button className="bar-btn" onClick={restoreOrphans} title="把这些引用格的来源重新登记为数据流连线（画布连线恢复）">
              恢复连线
            </button>
            <button className="bar-btn" onClick={clearOrphans} title="清空这些引用格，画布不再有残留">
              清空引用格
            </button>
          </div>
        )}
        {/* 智能填充预览框 + 填充柄（选区右下角小方块，按住拖动 = 复制扩展） */}
        <div ref={fillPreviewRef} className="wb-fill-preview" style={{ display: 'none' }} />
        {fillHandle && !readOnly && !fxEditing && (
          <div
            className="wb-fill-handle"
            style={{ left: fillHandle.left, top: fillHandle.top }}
            title="拖动填充：等差数字序列自动延续（1,2,3→4,5,6），公式相对引用偏移（=B2→=B3），其余复制扩展"
            onMouseDown={(e) => {
              e.preventDefault()
              e.stopPropagation()
              if (!selRect || !state) return
              const preview = fillPreviewRef.current
              if (!preview) return
              const origin = { r1: selRect.r1, c1: selRect.c1, r2: selRect.r2, c2: selRect.c2 }
              fillDragRef.current = { origin }
              fillLastRef.current = null
              const gridEl = document.querySelector('.wb-grid')
              const onMove = (ev: MouseEvent) => {
                const el = document.elementFromPoint(ev.clientX, ev.clientY) as HTMLElement | null
                const cell = el?.closest('.ag-cell') as HTMLElement | null
                if (!cell || !gridEl) return
                const col = cell.getAttribute('col-id')
                const rowId = cell.closest('.ag-row')?.getAttribute('row-id')
                if (!col || !rowId || !/^[A-Z]{1,3}$/.test(col) || !/^\d+$/.test(rowId)) return
                const rect = {
                  r1: origin.r1,
                  c1: origin.c1,
                  r2: Math.max(origin.r2, Number(rowId)),
                  c2: Math.max(origin.c2, colToIndex(col)),
                }
                fillLastRef.current = rect
                const a = document.querySelector(
                  `.ag-center-cols-container .ag-row[row-id="${rect.r1}"] .ag-cell[col-id="${indexToCol(rect.c1)}"]`,
                )
                const b = document.querySelector(
                  `.ag-center-cols-container .ag-row[row-id="${rect.r2}"] .ag-cell[col-id="${indexToCol(rect.c2)}"]`,
                )
                if (a && b) {
                  const gr = gridEl.getBoundingClientRect()
                  const ar = a.getBoundingClientRect()
                  const br = b.getBoundingClientRect()
                  preview.style.display = 'block'
                  preview.style.left = ar.left - gr.left + 'px'
                  preview.style.top = ar.top - gr.top + 'px'
                  preview.style.width = br.right - ar.left + 'px'
                  preview.style.height = br.bottom - ar.top + 'px'
                }
              }
              const onUp = () => {
                document.removeEventListener('mousemove', onMove)
                document.removeEventListener('mouseup', onUp)
                fillDragRef.current = null
                preview.style.display = 'none'
                const rect = fillLastRef.current
                fillLastRef.current = null
                if (!rect) return
                const ws = getSheet(state.wb, active)
                if (!ws) return
                const oh = origin.r2 - origin.r1 + 1
                const ow = origin.c2 - origin.c1 + 1
                // #18 智能序列：origin 块数字矩阵（只认 typeof number；公式/文本/空 = null）
                const nums: NumMatrix = []
                for (let r = origin.r1; r <= origin.r2; r++) {
                  const row: (number | null)[] = []
                  for (let c = origin.c1; c <= origin.c2; c++) {
                    const src = getCellByCoord(ws, indexToCol(c) + r)
                    const sv = src && typeof src === 'object' && 'value' in src ? (src as { value?: unknown }).value : src
                    row.push(typeof sv === 'number' && Number.isFinite(sv) ? sv : null)
                  }
                  nums.push(row)
                }
                let changed = 0
                for (let r = rect.r1; r <= rect.r2; r++) {
                  for (let c = rect.c1; c <= rect.c2; c++) {
                    if (r >= origin.r1 && r <= origin.r2 && c >= origin.c1 && c <= origin.c2) continue
                    const sr = origin.r1 + ((r - origin.r1) % oh)
                    const sc = origin.c1 + ((c - origin.c1) % ow)
                    const src = getCellByCoord(ws, indexToCol(sc) + sr)
                    let v: unknown = undefined
                    if (src && typeof src === 'object' && 'value' in src) {
                      const sv = (src as { value?: unknown }).value
                      if (sv && typeof sv === 'object' && (sv as { kind?: string }).kind === 'formula') {
                        const fv = sv as { formula?: string; cachedValue?: unknown }
                        // #18 公式引用偏移：按 (r-sr, c-sc) 平移相对引用（$ 绝对/字符串字面量由 translateFormula 保护）；越界降级原样
                        // ⚠ fv.formula 经 makeFormula 剥过前导 '='（存的是 'C1*2'），translateFormula 只认 '=' 开头——先补回再平移
                        const rawF = fv.formula ?? ''
                        const withEq = rawF.startsWith('=') ? rawF : '=' + rawF
                        let f = withEq
                        try {
                          f = translateFormula(withEq, indexToCol(sc) + sr, { rowDelta: r - sr, colDelta: c - sc })
                        } catch {
                          /* 平移越界（相对引用推过 A1/行0）→ 原样复制 */
                        }
                        setCellByCoord(
                          ws,
                          indexToCol(c) + r,
                          makeFormula(f, { cachedValue: (fv.cachedValue ?? '') as never }),
                        )
                        changed++
                        continue
                      }
                      v = sv
                    }
                    // #18 数字序列：纵向/横向扩展带严格等差外推（fillSmartNumber）；角块/非序列 → 平铺
                    const smart = fillSmartNumber(nums, origin, r, c)
                    if (smart != null) {
                      setCellByCoord(ws, indexToCol(c) + r, smart as never)
                      changed++
                      continue
                    }
                    setCellByCoord(ws, indexToCol(c) + r, v as never)
                    changed++
                  }
                }
                if (changed) void commit()
              }
              document.addEventListener('mousemove', onMove)
              document.addEventListener('mouseup', onUp)
            }}
          />
        )}
      </div>

      <div className="wb-foot">
        {/* 备忘 · 节点数据引用（原属性引用）显示/折叠 开关（底栏最左端 · 用户指定两侧放置） */}
        <button
          className="bar-btn"
          title={refPanelOpen ? '折叠节点数据引用' : '显示节点数据引用'}
          style={
            refPanelOpen
              ? { background: 'var(--ed-bg3)', color: 'var(--ed-accent)', borderColor: 'var(--ed-accent)', fontWeight: 600 }
              : undefined
          }
          onClick={() => toggleSide('ref')}
        >
          ☰ 节点数据引用
        </button>
        {/* 备忘 · 全部连线两态：点击循环 前端（默认）→ 隐藏 → 前端；隐藏态悬停临时显。
            2026-09-26 用户拍板：窗体和引用面板一律不透明，原「后端透明化露线」态废弃（wb-lines-back-on 不再触发） */}
        <button
          className={`bar-btn${allMode !== 0 ? ' wb-alllock' : ''}`}
          title="引用连线显示：跟随选中（点关联单元格显示其连线）→ 全部常显 → 隐藏（悬停可临时显全部），点击循环"
          onMouseEnter={() => setAllLinesHover(true)}
          onMouseLeave={() => setAllLinesHover(false)}
          onClick={() => setAllMode(((allMode + 1) % 3) as 0 | 1 | 2)}
        >
          ⌗ 连线：{allMode === 0 ? '跟随选中' : allMode === 1 ? '全部' : '隐藏'}
        </button>
        <span className="wb-file" title={name}>
          {name}
        </span>
        {dirty > 0 && <span className="wb-dirty">已改 {dirty} 处</span>}
        {readOnly && <span className="wb-ro">只读</span>}
        <span className="wb-sp" />
        {busy && <span className="wb-busy">{busy}</span>}
        <button className="bar-btn" disabled={undoStack.current.length <= 1} onClick={doUndo}>
          撤销改动
        </button>
        <button className="bar-btn" onClick={() => void reload()} title="丢弃未保存改动，重新解析">
          重新载入
        </button>
        <button className="bar-btn" onClick={() => { closeSides(); setShowDomains(true) }} disabled={!fields.length}>
          域位置（{fields.length}）
        </button>
        <button
          className="bar-btn"
          title="阅读模式：高亮选中格所在的行与列"
          style={
            hlCross
              ? { background: 'var(--ed-bg3)', color: 'var(--ed-accent)', borderColor: 'var(--ed-accent)', fontWeight: 600 }
              : undefined
          }
          onClick={() => setHlCross((v) => !v)}
        >
          ☰ 高亮行列
        </button>
        {/* 备忘 · 数据预选追加 折叠开关（底栏最右端 · 用户指定两侧放置；进退动画同左栏） */}
        <button
          className="bar-btn"
          title={rightOpen ? '折叠数据预选追加' : '显示数据预选追加'}
          style={
            rightOpen
              ? { background: 'var(--ed-bg3)', color: 'var(--ed-accent)', borderColor: 'var(--ed-accent)', fontWeight: 600 }
              : undefined
          }
          onClick={() => toggleSide('right')}
        >
          ☰ 数据预选
        </button>
      </div>
      </div>

      {/* 属性引用栏（portal · 普通视图=窗体外圆角卡片 / 最大化=嵌入窗体左缘）+ 引用连线（fixed 全屏 SVG） */}
      {/* ★ 批次17 修 bug：records 是可选 prop（类型 `records?: Record<string, Row[]>`），
          它只用于引用行的「值预览」（RefPanel 内取值走 records?.… 可选链，取不到显示空即可）。
          原先把它当挂载必需条件 → 节点未绑定输入数据（records 为 undefined）时，
          整个属性引用栏**静默不挂载**：按钮显示 ✓（状态是开的）但面板不存在，
          用户看到的是「点参数毫无反应」，且没有任何报错。
          正确口径：records 缺失只影响值预览，不影响「选参数 → 插入引用」这条主链路。 */}
      {panelMounted && tables && onSetNodeConfig && nodeId && posForRender && (
        <RefPanel
          tables={tables}
          records={records ?? {}}
          selfId={nodeId}
          activeRef={selURN}
          onPick={pickRef}
          onClose={() => setRefPanelOpen(false)}
          animClass={maximized ? undefined : panelLeaving ? 'ref-panel-anim-out' : 'ref-panel-anim-in'}
          panelStyle={{
            position: 'fixed',
            left: posForRender.left,
            top: posForRender.top,
            height: posForRender.height,
            width: 230,
            borderRadius: posForRender.radius,
            border: posForRender.border ? '1px solid var(--ed-border)' : undefined,
            // 底栏（.ref-tip）据此收圆：常规态 radius 10→ 变量 9；最大化态 radius 0 → 变量 0（直角）
            ['--ref-radius' as string]: `${Math.max(0, posForRender.radius - 1)}px`,
          }}
        />
      )}

      {/* 联动触发规则面板（portal · 与属性引用栏同位：普通=窗体外左卡片 / 最大化=嵌入左缘） */}
      {triggerOpen && state && tables && posForRender && (
        <div
          className="ref-panel"
          style={{
            position: 'fixed',
            left: posForRender.left,
            top: posForRender.top,
            height: posForRender.height,
            width: 300,
            borderRadius: posForRender.radius,
            border: posForRender.border ? '1px solid var(--ed-border)' : undefined,
            ['--ref-radius' as string]: `${Math.max(0, posForRender.radius - 1)}px`,
          }}
        >
          <div className="ref-panel-head">
            <span>联动规则</span>
            <button className="ref-panel-x" onClick={() => setTriggerOpen(false)} title="关闭">
              ×
            </button>
          </div>
          <TriggerPanel
            sheetNames={wbSheetNames(state.wb)}
            sources={trgSources}
            rules={trgRules}
            onRulesChange={setTrgRules}
            onApply={applyTrgPlan}
            onDryRun={dryRunTrgPlan}
            readSource={(u) => urnResolverRef.current(u)}
          />
        </div>
      )}
      {/* 备忘 · 右侧「数据预选追加」面板（镜像左栏：普通=窗体外右侧卡片 / 最大化=嵌入右缘） */}
      {rightMounted && tables && records && onSetNodeConfig && nodeId && rightPosForRender && (
        <RightPanel
          tables={tables}
          records={records}
          selfId={nodeId}
          inputs={tables[nodeId]?.config?.inputs ?? []}
          onToggle={(srcId, on) => {
            const cur = tables[nodeId]?.config?.inputs ?? []
            const next = on ? (cur.includes(srcId) ? cur : [...cur, srcId]) : cur.filter((x) => x !== srcId)
            onSetNodeConfig(nodeId, { inputs: next })
          }}
          onClose={() => setRightOpen(false)}
          animClass={maximized ? undefined : rightLeaving ? 'ref-panel-anim-out-r' : 'ref-panel-anim-in-r'}
          panelStyle={{
            position: 'fixed',
            left: rightPosForRender.left,
            top: rightPosForRender.top,
            height: rightPosForRender.height,
            width: 230,
            borderRadius: rightPosForRender.radius,
            border: rightPosForRender.border ? '1px solid var(--ed-border)' : undefined,
            ['--ref-radius' as string]: `${Math.max(0, rightPosForRender.radius - 1)}px`,
          }}
        />
      )}
      {/* 备忘 · 右键单元格菜单（WPS 常规项） */}
      {ctxMenu &&
        createPortal(
          <div className="wb-ctxmenu" style={{ left: ctxMenu.x, top: ctxMenu.y }}>
            {buildCtxItems().map((it, i) =>
              it.sep ? (
                <div key={i} className="wb-ctxmenu-sep" />
              ) : it.children ? (
                /* #33 二级子菜单容器：hover 浮出右侧（纯 CSS，无额外 state） */
                <div key={i} className="wb-ctxmenu-itemwrap">
                  <button className="wb-ctxmenu-item" title={it.title}>
                    {it.label}
                    <span className="wb-ctxmenu-arrow">▸</span>
                  </button>
                  <div className="wb-ctxmenu sub">
                    {it.children.map((c, j) =>
                      c.sep ? (
                        <div key={j} className="wb-ctxmenu-sep" />
                      ) : (
                        <button
                          key={j}
                          className="wb-ctxmenu-item"
                          title={c.title}
                          onClick={() => {
                            setCtxMenu(null)
                            c.run?.()
                          }}
                        >
                          {c.label}
                        </button>
                      ),
                    )}
                  </div>
                </div>
              ) : (
                <button
                  key={i}
                  className="wb-ctxmenu-item"
                  title={it.title}
                  onClick={() => {
                    setCtxMenu(null)
                    it.run?.()
                  }}
                >
                  {it.label}
                </button>
              ),
            )}
          </div>,
          document.body,
        )}
      {/* #33A/#39 列头右键菜单（排序/筛选/插删列/列宽/数据格式） */}
      {hdrMenu &&
        createPortal(
          <div className="wb-ctxmenu" style={{ left: hdrMenu.x, top: hdrMenu.y }}>
            {buildHdrItems(hdrMenu.col).map((it, i) =>
              it.sep ? (
                <div key={i} className="wb-ctxmenu-sep" />
              ) : (
                <button
                  key={i}
                  className="wb-ctxmenu-item"
                  title={it.title}
                  onClick={() => {
                    setHdrMenu(null)
                    it.run?.()
                  }}
                >
                  {it.label}
                </button>
              ),
            )}
          </div>,
          document.body,
        )}
      {/* #33A 筛选值清单面板（列头右键「筛选…」弹出；确定 = 写 filterState，全选 = 清除） */}
      {filterDlg &&
        createPortal(
          <div className="wb-filterdlg" style={{ left: filterDlg.x, top: filterDlg.y }} data-filterdlg="1">
            <div className="wb-filterdlg-head">
              筛选列 <span className="wb-filterdlg-col">{filterDlg.col}</span>
              <span style={{ flex: 1 }} />
              <button
                className="bar-btn"
                onClick={() => setFilterPick(new Set(filterValueList(filterDlg.col).map((x) => x.key)))}
              >
                全选
              </button>
            </div>
            <div className="wb-filterdlg-list">
              {filterValueList(filterDlg.col).map(({ key, count }) => (
                <label key={key || '(空)'} className="wb-filterdlg-item">
                  <input
                    type="checkbox"
                    checked={filterPick.has(key)}
                    onChange={(e) => {
                      const next = new Set(filterPick)
                      if (e.target.checked) next.add(key)
                      else next.delete(key)
                      setFilterPick(next)
                    }}
                  />
                  <span title={key === '' ? '(空白)' : key}>{key === '' ? '(空白)' : key}</span>
                  <em style={{ marginLeft: 'auto', fontStyle: 'normal', color: 'var(--ed-text4)', fontSize: 10 }}>{count}</em>
                </label>
              ))}
            </div>
            <div className="wb-filterdlg-foot">
              <button
                className="bar-btn"
                onClick={() => {
                  setFilterState(null)
                  setFilterDlg(null)
                }}
              >
                清除筛选
              </button>
              <span className="spacer" />
              <button className="bar-btn" onClick={() => setFilterDlg(null)}>
                取消
              </button>
              <button
                className="bar-btn"
                data-filterdlg-ok="1"
                onClick={() => {
                  // 全选 = 语义清除；部分勾选 = 保留勾选值
                  const all = filterValueList(filterDlg.col)
                  const isAll = all.length > 0 && all.every(({ key }) => filterPick.has(key))
                  setFilterState(isAll ? null : { col: filterDlg.col, keep: [...filterPick] })
                  setFilterDlg(null)
                }}
              >
                确定
              </button>
            </div>
          </div>,
          document.body,
        )}
      {/* 备忘 · 选区外轮廓：整块连续矩形（替代逐格描边，多格选区不断线） */}
      {state && gridVpEl && createPortal(<div ref={selOutlineRef} className="wb-seloutline" style={{ display: 'none' }} />, gridVpEl)}
      {refPanelOpen && !!selURN && createPortal(
        <svg ref={lineSvgRef} className="wb-refline">
          {/* 全部连线 · 常显模式（细灰实线，弱于选中线）—— allMode=1 全部 / 2 隐藏+悬停临时显 */}
          {showAllFront &&
            allLines.map((l) => <path key={l.key} className="wb-refline-all" d={l.d} />)}
          <g className={refHover ? 'urn-hover' : undefined}>
            {/* 宽透明命中条：父级 svg pointer-events:none，只有这条线附近 14px 可 hover —— 不挡格子点击 */}
            <path
              ref={hitPathRef}
              className="wb-refline-hit"
              d=""
              onMouseEnter={() => {
                setRefHover(true)
                markRefEndpoints(true)
              }}
              onMouseLeave={() => {
                setRefHover(false)
                markRefEndpoints(false)
              }}
            />
            {/* 选中线：金色实线；d 由 measure 直写 DOM（滚动丝滑跟随） */}
            <path ref={linePathRef} className="wb-refline-line" d="" />
          </g>
        </svg>,
        document.body,
      )}
      {/* 后端毛玻璃渲染段已删（wb-lines-back-on 同批废弃，2026-09-26 晚拍板） */}
      {refHover && refPts && selURN && refPanelOpen && createPortal(
        (() => {
          const t = tables?.[selURN.nodeId]
          const fld = t?.fields.find((f) => f.id === selURN.fieldId || f.name === selURN.fieldId)
          const sample = records?.[selURN.nodeId]?.[0]?.[selURN.fieldId as keyof Row]
          return (
            <div
              className="wb-refline-tip"
              style={{ left: (refPts.x1 + refPts.x2) / 2, top: Math.min(refPts.y1, refPts.y2) }}
            >
              <b>{t?.name ?? selURN.nodeId}.{fld?.name ?? selURN.fieldId}</b>
              <span className="tip-go"> → {selCoord}</span>
              <div className="tip-v">当前值：{String(sample ?? '（空）')}</div>
            </div>
          )
        })(),
        document.body,
      )}
      </div>

      {showDomains && (
        <DomainsDialog
          fields={fields}
          onClose={() => setShowDomains(false)}
          onGo={(sheet) => {
            if (sheet) setActive(sheet)
            setShowDomains(false)
          }}
        />
      )}

      {/* 函数库面板（373 函数搜索 + 插入） */}
      {openFnLib && (
        <div
          style={{ position: 'fixed', inset: 0, zIndex: 'var(--z-modal)', background: 'rgba(1,4,9,.6)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
          onPointerDown={(e) => e.target === e.currentTarget && setOpenFnLib(false)}
        >
          <div
            className="wb-fnlib"
            style={{
              width: 460, maxHeight: 640, display: 'flex', flexDirection: 'column',
              background: 'var(--ed-bg0)', border: '1px solid var(--ed-border)', borderRadius: 10, overflow: 'hidden',
            }}
          >
            <div style={{ padding: '12px 16px', borderBottom: '1px solid var(--ed-border2)' }}>
              <div style={{ fontSize: 14, color: 'var(--ed-text)', marginBottom: 8 }}>
                函数库 · 共 {FUNC_COUNT} 个（formulajs · MIT）
              </div>
              <input
                autoFocus
                placeholder="搜索函数，如 SUM / IF / VLOOKUP…"
                value={fnSearch}
                onChange={(e) => setFnSearch(e.target.value)}
                style={{
                  width: '100%', boxSizing: 'border-box', fontSize: 12.5, color: 'var(--ed-text)',
                  background: 'var(--ed-bg2)', border: '1px solid var(--ed-border2)', borderRadius: 6, padding: '6px 10px',
                }}
              />
            </div>
            <div style={{ flex: 1, overflowY: 'auto' }}>
              {fnLibItems.map((f) => (
                <div
                  key={f}
                  className="wb-fnlib-item"
                  style={{ padding: '6px 16px', fontSize: 12.5, fontFamily: 'Consolas, monospace', color: 'var(--ed-text)', cursor: 'pointer' }}
                  title="点击：把函数插入公式栏"
                  onClick={() => {
                    insertFn(f)
                    setOpenFnLib(false)
                  }}
                >
                  {f}()
                </div>
              ))}
            </div>
            <div style={{ padding: '8px 16px', borderTop: '1px solid var(--ed-border2)', fontSize: 11, color: 'var(--ed-text3)' }}>
              双击插入到公式栏 · 选中共 {COMMON_FNS.length} 个常用函数置顶
            </div>
          </div>
        </div>
      )}

      {/* 批注清单 */}
      {openCmts && (
        <div
          style={{ position: 'fixed', inset: 0, zIndex: 'var(--z-modal)', background: 'rgba(1,4,9,.6)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
          onPointerDown={(e) => e.target === e.currentTarget && setOpenCmts(false)}
        >
          <div
            className="wb-cmtlist"
            style={{
              width: 520, maxHeight: 560, display: 'flex', flexDirection: 'column',
              background: 'var(--ed-bg0)', border: '1px solid var(--ed-border)', borderRadius: 10, overflow: 'hidden',
            }}
          >
            <div style={{ padding: '12px 16px', borderBottom: '1px solid var(--ed-border2)', fontSize: 14, color: 'var(--ed-text)' }}>
              批注清单（{cmtRows.length}）
            </div>
            <div style={{ flex: 1, overflowY: 'auto' }}>
              {cmtRows.length === 0 && (
                <div style={{ padding: 20, fontSize: 12, color: 'var(--ed-text3)' }}>本工作表还没有批注。「审阅 → 新建批注」给选中格加一条。</div>
              )}
              {cmtRows.map((c, i) => (
                <div key={i} className="wb-cmt-row" style={{ padding: '8px 16px', borderBottom: '1px solid var(--ed-bg3)' }}>
                  <div style={{ fontSize: 11, color: 'var(--ed-text3)', marginBottom: 3 }}>
                    {c.ref} · {c.author}
                  </div>
                  <div style={{ fontSize: 12.5, color: 'var(--ed-text)', whiteSpace: 'pre-wrap' }}>{c.text}</div>
                  <button
                    className="bar-btn"
                    style={{ marginTop: 6 }}
                    onClick={() => {
                      setSelCoord(c.ref)
                      setOpenCmts(false)
                    }}
                  >
                    定位
                  </button>
                </div>
              ))}
            </div>
            <div style={{ padding: '8px 16px', borderTop: '1px solid var(--ed-border2)' }}>
              <button className="bar-btn" onClick={() => setOpenCmts(false)}>关闭</button>
            </div>
          </div>
        </div>
      )}

      {/* 章节编排窗体（章节工作簿节点）——
          形态与上方「批注清单」逐项对齐：同一个 fixed+inset:0 遮罩、同一个 --z-modal、
          同一套卡片底色/圆角/三段式布局。因此**最大化态与普通视口表现完全一致**，
          不需要写第二套分支（用户 2026-10-05 明确要求）。 */}
      {outlineSlot.enabled && outlineSlot.open && (
        <ChapterDialog
          rows={outlineSlot.rows}
          onCommit={outlineSlot.commit}
          onClose={() => outlineSlot.setOpen(false)}
          // 点章节 → 选中对应行并关窗体（与批注清单的「定位」同一套口径：
          // 只改选中格，靠 ag-grid 自行滚到可视区，不另写滚动逻辑）
          onGoto={(rowIndex) => {
            setSelCoord('A' + (rowIndex + 1))
            setPickCell(null)
            setGlowCell(null)
            outlineSlot.setOpen(false)
          }}
          readOnly={outlineSlot.readOnly || readOnly === true}
        />
      )}

      {/* 设置单元格格式对话框（#33 第二块）—— 与章节窗体同一套 fixed 遮罩 + --z-modal；
          确定走 run()（改 wb → commit 落盘 + undo 快照），与右键菜单其他项同一口径 */}
      {fmtDlg && state && (
        <FmtDialog
          init={fmtDlg.init}
          readOnly={readOnly === true}
          onCancel={() => setFmtDlg(null)}
          onOk={(next) => {
            run(() => applyFmtSel(state.wb, active, fmtDlg.rect, next))
            setFmtDlg(null)
          }}
        />
      )}
    </div>
  )
}

/** 域清单：列出所有域并支持跳到所在工作表 */
function DomainsDialog({
  fields,
  onClose,
  onGo,
}: {
  fields: XlsxFieldSpec[]
  onClose: () => void
  onGo: (sheet?: string) => void
}) {
  return (
    <div className="wb-domains" onPointerDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="wb-domains-box">
        <div className="wb-domains-head">
          <span>这份工作簿里的域（{fields.length}）</span>
          <span className="wb-sp" />
          <button className="ne-close" onClick={onClose} title="关闭">
            ×
          </button>
        </div>
        <div className="wb-domains-body">
          {fields.length === 0 && (
            <div className="wb-empty">
              没有发现域。在 Excel/WPS 里选中单元格 → 公式 → 定义名称，或把数据区做成「表格」（Ctrl+T）。
            </div>
          )}
          {fields.map((f) => (
            <div key={f.key} className="wb-domain-row">
              <span className={`xrp-kind ${f.kind === 'table' ? 'xrp-kind--table' : 'xrp-kind--name'}`}>
                {f.kind === 'table' ? '表格' : '命名区域'}
              </span>
              <span className="wb-domain-key">{f.key}</span>
              <span className="wb-domain-ref">{displayRef(f.sheet, f.ref)}</span>
              <button className="bar-btn wb-domain-go" onClick={() => onGo(f.sheet)} disabled={!f.sheet}>
                定位
              </button>
            </div>
          ))}
          <div className="wb-domains-foot">
            网格里：<b className="wb-hl-named">浅蓝格</b> = 命名区域指向的格子；
            <b className="wb-hl-table">左侧绿条</b> = 表格（行级域）范围。
          </div>
        </div>
      </div>
    </div>
  )
}
