/**
 * pptx 模板绑定：扫描占位符 + 填数据。
 *
 * ## 与 docx 侧的关键差异（决定了实现不能照抄）
 *
 * | | docx | pptx |
 * |---|---|---|
 * | 结构标记 | 有内容控件 `w:sdt`，控件外壳常驻 | **没有内容控件**，只能靠纯文本占位符 |
 * | 可否反复重绑 | 可以（外壳在，可增量更新） | **不可以** —— 填完占位符就没了 |
 * | 元素定位 | 控件 tag | 形状名（真实底稿里是「文本框 160」，没法用） |
 *
 * 所以 pptx 走的是 **一次性填充**：模板 + 数据 → 产物，产物不能再当模板用。
 * 这是格式限制，不是实现偷懒 —— 用法上必须「从模板重新生成」，和 docx 的
 * 「主路径永远是模板+数据→产物」是同一条原则。
 *
 * ## 支持的占位符
 *
 *   {{field:项目名}}   文字替换（可跨 run）
 *   {{if:显示附录}}    段落级，假则整段删除
 *   {{each:清单}}      放在表格模板行的任意单元格里，该行按数据行数克隆；
 *                      行内用短占位符 {{字段名}} 取每行的值
 *   {{rich:说明}}      暂不支持（多段富文本需重建 txBody，见文末 TODO）
 *   {{image:照片}}     暂不支持（需 media + rels + Content_Types 三处同步）
 */

import JSZip from 'jszip'
import {
  paragraphBlocks,
  replaceInParagraph,
  replaceShortInParagraph,
  scanParagraph,
  stripEachMarker,
  type PlaceholderKind,
} from './text'

export interface PptxBinding {
  kind: PlaceholderKind
  /** 占位符里写的名字，如「项目名」 */
  path: string
  /** 完整占位符文本 {{field:项目名}} */
  tag: string
  /** 所在部件，如 ppt/slides/slide1.xml */
  part: string
  /** 页码，从 1 开始 */
  slide: number
  /** each 专属：模板行内的短占位符名列表，供列映射 UI */
  placeholders: string[]
  /** 模板里该控件附近的原文，UI 上给用户看上下文 */
  sample: string
}

export interface PptxParseResult {
  bindings: PptxBinding[]
  warnings: string[]
}

export interface PptxBindResult {
  buffer: Uint8Array
  bound: number
  warnings: string[]
}

/** slide 部件按页码排序：slide2 必须排在 slide10 前面，字典序会排反 */
function slideParts(zip: JSZip): string[] {
  const names = Object.keys(zip.files).filter(
    (n) => !zip.files[n].dir && /^ppt\/slides\/slide\d+\.xml$/.test(n),
  )
  return names.sort((a, b) => num(a) - num(b))
}

const num = (s: string) => Number(/(\d+)/.exec(s)?.[1] ?? 0)

const EACH_MARKER_RE = /\{\{\s*each\s*:\s*([^{}]+?)\s*\}\}/
const ROW_RE = /<a:tr(?:\s[^>]*)?>[\s\S]*?<\/a:tr>/g
const TABLE_RE = /<a:tbl>[\s\S]*?<\/a:tbl>/g
const PARA_RE = /<a:p>[\s\S]*?<\/a:p>/g

/** 取一段 XML 里所有 a:t 的可见文字，用于给 UI 显示上下文 */
function sampleOf(block: string): string {
  return [...block.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)]
    .map((m) => m[1])
    .join('')
    .trim()
    .slice(0, 80)
}

/* ------------------------------------------------------------------ *
 * 扫描
 * ------------------------------------------------------------------ */

export async function parsePptxTemplate(
  input: ArrayBuffer | Uint8Array,
): Promise<PptxParseResult> {
  const zip = await JSZip.loadAsync(input)
  const bindings: PptxBinding[] = []
  const warnings: string[] = []
  const seen = new Set<string>()

  for (const part of slideParts(zip)) {
    const f = zip.file(part)
    if (!f) continue
    const xml = await f.async('string')
    const slide = num(part)

    for (const { block } of paragraphBlocks(xml)) {
      for (const hit of scanParagraph(block)) {
        const key = `${part}::${hit.raw}`
        // 同一个占位符可能在一页里出现多次（比如页眉页脚），只报一次
        if (seen.has(key)) continue
        seen.add(key)

        const b: PptxBinding = {
          kind: hit.kind,
          path: hit.path,
          tag: hit.raw,
          part,
          slide,
          placeholders: [],
          sample: sampleOf(block),
        }
        if (hit.kind === 'each') b.placeholders = shortPlaceholdersOfRow(xml, hit.raw)
        bindings.push(b)
      }
    }
  }

  if (!bindings.length) warnings.push('模板里没有找到任何占位符（写法：{{field:项目名}}）')
  const unsupported = [...new Set(bindings.filter((b) => b.kind === 'rich' || b.kind === 'image').map((b) => b.kind))]
  for (const k of unsupported) {
    warnings.push(`${k}: 暂不支持（见 src/pptx/binder.ts 顶部说明），将原样保留`)
  }
  return { bindings, warnings }
}

/** each 标记所在表格行里的短占位符名 —— 就是这行能填哪些列 */
function shortPlaceholdersOfRow(xml: string, marker: string): string[] {
  for (const t of xml.matchAll(TABLE_RE)) {
    for (const r of t[0].matchAll(ROW_RE)) {
      if (!EACH_MARKER_RE.test(r[0]) || !r[0].includes(marker)) continue
      const names = new Set<string>()
      for (const p of r[0].matchAll(PARA_RE)) {
        for (const m of p[0].matchAll(
          /\{\{(?!(?:field|rich|each|if|image)\s*:)\s*([^{}]+?)\s*\}\}/g,
        )) {
          names.add(m[1])
        }
      }
      return [...names]
    }
  }
  return []
}

/* ------------------------------------------------------------------ *
 * 绑定
 * ------------------------------------------------------------------ */

export interface PptxBindOptions {
  /** each 数据为空时，是否保留一行空行（默认删掉整行） */
  keepEmptyRow?: boolean
}

export async function bindPptx(
  input: ArrayBuffer | Uint8Array,
  data: Record<string, unknown>,
  opts: PptxBindOptions = {},
): Promise<PptxBindResult> {
  const zip = await JSZip.loadAsync(input)
  const warnings: string[] = []
  let bound = 0

  for (const part of slideParts(zip)) {
    const f = zip.file(part)
    if (!f) continue
    let xml = await f.async('string')

    // 顺序不能反：先展开表格（会改变长度、新增行），再处理段落。
    // 段落阶段会重新计算偏移量，所以表格改动不影响它。
    const t = expandTables(xml, data, warnings, opts)
    xml = t.xml
    bound += t.bound

    const p = fillParagraphs(xml, data, warnings)
    xml = p.xml
    bound += p.bound

    zip.file(part, xml)
  }

  const buffer = await zip.generateAsync({ type: 'uint8array' })
  return { buffer, bound, warnings }
}

/* ---------- 表格行循环 ---------- */

function expandTables(
  xml: string,
  data: Record<string, unknown>,
  warnings: string[],
  opts: PptxBindOptions,
): { xml: string; bound: number } {
  const tables = [...xml.matchAll(TABLE_RE)]
  if (!tables.length) return { xml, bound: 0 }
  let bound = 0
  let out = xml

  // 从后往前改，前面的偏移量才不会失效
  for (let i = tables.length - 1; i >= 0; i--) {
    const at = tables[i].index!
    const tbl = tables[i][0]
    const rows = [...tbl.matchAll(ROW_RE)]
    let newTbl = tbl

    for (let j = rows.length - 1; j >= 0; j--) {
      const rm = rows[j]
      const rowXml = rm[0]
      const marker = EACH_MARKER_RE.exec(rowXml)
      if (!marker) continue
      const name = marker[1].trim()
      const list = data[name]

      if (!Array.isArray(list)) {
        warnings.push(`each「${name}」的数据不是数组，该行已原样保留（第 ${i + 1} 个表格）`)
        continue
      }
      if (!list.length) {
        if (opts.keepEmptyRow) {
          newTbl =
            newTbl.slice(0, rm.index!) +
            stripEachMarker(rowXml) +
            newTbl.slice(rm.index! + rowXml.length)
        } else {
          newTbl = newTbl.slice(0, rm.index!) + newTbl.slice(rm.index! + rowXml.length)
        }
        continue
      }

      const clones = list.map((row) => {
        const rec = (row ?? {}) as Record<string, unknown>
        return fillRow(stripEachMarker(rowXml), rec)
      })
      bound += list.length
      newTbl =
        newTbl.slice(0, rm.index!) + clones.join('') + newTbl.slice(rm.index! + rowXml.length)
    }

    out = out.slice(0, at) + newTbl + out.slice(at + tbl.length)
  }
  return { xml: out, bound }
}

/** 填一行里的短占位符 {{字段名}} */
function fillRow(rowXml: string, row: Record<string, unknown>): string {
  let out = rowXml
  const paras = [...rowXml.matchAll(PARA_RE)]
  for (let i = paras.length - 1; i >= 0; i--) {
    const p = paras[i]
    const { block } = replaceShortInParagraph(p[0], (name) => {
      const v = row[name]
      return v == null ? null : String(v)
    })
    out = out.slice(0, p.index!) + block + out.slice(p.index! + p[0].length)
  }
  return out
}

/* ---------- 段落：field / if ---------- */

function fillParagraphs(
  xml: string,
  data: Record<string, unknown>,
  warnings: string[],
): { xml: string; bound: number } {
  let out = xml
  let bound = 0
  const paras = paragraphBlocks(out)

  for (let i = paras.length - 1; i >= 0; i--) {
    const { block, at } = paras[i]
    const hits = scanParagraph(block)

    // if 为假：整段删除。判断用原始占位符，不受替换顺序影响
    const ifHit = hits.find((h) => h.kind === 'if')
    if (ifHit && !truthy(data[ifHit.path])) {
      out = out.slice(0, at) + out.slice(at + block.length)
      continue
    }

    const { block: next, count } = replaceInParagraph(block, (_raw, hit) => {
      // if 标记本身不显示：条件为真时它已完成使命，从正文里去掉
      if (hit.kind === 'if') return ''
      if (hit.kind !== 'field') return null // rich / image / each 不在这里处理
      const v = data[hit.path]
      if (v == null) {
        warnings.push(`field「${hit.path}」没有数据，占位符原样保留`)
        return null
      }
      return String(v)
    })
    bound += count
    if (next !== block) out = out.slice(0, at) + next + out.slice(at + block.length)
  }
  return { xml: out, bound }
}

function truthy(v: unknown): boolean {
  if (v === undefined || v === null || v === '' || v === false) return false
  if (typeof v === 'string') return v !== 'false' && v !== '0'
  return Boolean(v)
}

/*
 * TODO（按性价比排序）
 *  1. rich：把 HTML 转成多个 a:p 塞进 p:txBody，复用 docx 的 html2ooxml 思路，
 *     但发射的是 a:p/a:r/a:t 而不是 w:p/w:r/w:t。
 *  2. image：与 docx 同样的三处注入（media + rels + Content_Types），
 *     pptx 的 rels 在 ppt/slides/_rels/ 下，Content_Types 同一份。
 *  3. notesSlides：备注页目前不参与绑定。
 */
