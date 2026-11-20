/**
 * OO 域外链接 · 写回管线（值变 → 写回 OO 文件）
 *
 * 立档：2026-10-06 · 用户：「实现它」（接上一轮做完的"点选形成关联"）
 *
 * 链路：
 *   ooRefs（关联清单）
 *     → 取当前值（computeNode 首行）
 *     → 按域位置算出要写的格子 / 文本
 *     → 补丁式改 OO 字节（保格式：图表/绘图等部件零丢失）
 *     → 存回模板库（ooTemplateId）
 *
 * ★ 分层：本文件是**纯计算层**（零 IO / 零 electron / 零 document）——
 *   字节进、字节出。IO（读模板、存模板）在 src/office/ooWriteback.ts。
 *   这样浏览器与 Node 同构，且可单测。
 *
 * ★ 铁律：补丁不重建。`workbookToBytes` 只写回自己建模的部件，
 *   真实文件里的图表/绘图/透视表会**静默丢失**（见 xlsx/patchSave.ts 注释实测）。
 */

import JSZip from 'jszip'
import type { OoRef } from './ooRef'

/** 域在 OO 文件里的落点（由 discover 阶段给出） */
export interface OoRefSite {
  /** 域清单项（用 nodeId/fieldId 对齐） */
  ref: OoRef
  /** oosheet：单元格地址如 'B3'；ooword：正文锚点标识 */
  a1?: string
  /** Word：内容控件（sdt）的 tag */
  sdtTag?: string
  /** Word：章节/段落定位用的序号 */
  part?: string
}

/** 写回结果 */
export interface OoWritebackResult {
  bytes: Uint8Array
  /** 实际改了几个域 */
  written: number
  /** 命中的域名（用于诊断） */
  labels: string[]
  /** 从原文件补回的部件名（格式保真的证据） */
  restored: string[]
  /** 没写成功的（如域找不到落点） */
  skipped: { fieldId: string; why: string }[]
}

/* ── ① 值→ 单元格文本 ─────────────────────────────────────────── */

/** 值格式化：空值给空串（不是 'undefined'/'null'） */
export function fmtCellValue (v: unknown): string {
  if (v == null) return ''
  if (typeof v === 'number') return Number.isFinite(v) ? String(v) : ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  const s = String(v)
  // Excel 单元格不认前导 = （会被当公式），需要时用 CONCAT 之类，这是已知限制
  return s
}

/**
 * 算出一个 OO 节点要写的全部值。
 * 纯函数：从 sites + 取值函数 → { a1 → 值 } 映射。
 */
export function planWrites (
  sites: OoRefSite[],
  pick: (ref: OoRef) => unknown,
): { cells: { a1: string; value: string }[]; skipped: { fieldId: string; why: string }[] } {
  const cells: { a1: string; value: string }[] = []
  const skipped: { fieldId: string; why: string }[] = []
  const seen = new Set<string>()
  for (const s of sites) {
    if (!s.a1) { skipped.push({ fieldId: s.ref.fieldId, why: '无落点（a1 缺失）' }); continue }
    // 一个格子只写一次（多个引用指向同一格时取第一个）
    if (seen.has(s.a1)) continue
    seen.add(s.a1)
    cells.push({ a1: s.a1, value: fmtCellValue(pick(s.ref)) })
  }
  return { cells, skipped }
}

/* ── ② xlsx 补丁：改 sharedStrings / sheet XML 的单元格值 ──────── */

/** 从 A1 解析出列字母 + 行号（1-based） */
export function splitA1 (a1: string): { col: string; row: number } | null {
  const m = /^([A-Z]+)(\d+)$/.exec(String(a1 || '').trim().toUpperCase())
  if (!m) return null
  return { col: m[1], row: Number(m[2]) }
}

/** 列字母 → 1-based 列号（用于排序/定位） */
export function colToNum (col: string): number {
  let n = 0
  for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}

/** 在 sheetN.xml 里找到目标 cell 元素并替换其内容。
 *  策略：正则定位 <c r="B3"[^>]*>...</c>，保留属性、只换内部值。
 *  找不到该 cell 时不新建（保持文件最小改动）—— 由调用方决定是否放弃。
 */
export function patchSheetXml (xml: string, a1: string, value: string): { xml: string; hit: boolean } {
  const parsed = splitA1(a1)
  if (!parsed) return { xml, hit: false }
  const re = new RegExp(`(<c r="${a1}"[^>]*?)(/>|>([\\s\\S]*?)</c>)`)
  const m = re.exec(xml)
  if (!m) return { xml, hit: false }
  // 数值型不带 t 属性（t="str"/"inlineStr" 是字符串），这里统一写成内联字符串
  const inner = `<v>${escapeXml(value)}</v>`
  const attrs = m[1]
  // 去掉可能存在的 t 属性，避免类型与内容不符
  const cleaned = attrs.replace(/\s+t="[^"]*"/g, '')
  const next = `${cleaned}>${inner}</c>`
  return { xml: xml.replace(m[0], next), hit: true }
}

function escapeXml (s: string): string {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

/* ── ③ docx 补丁：改 sdt 里的文本 ──────────────────────────────── */

/** 改 document.xml 里指定 tag 的 sdt 内容（w:t 文本）。
 *  找不到该 sdt 时 hit=false（不新建，避免破坏文档结构）。
 */
export function patchDocxSdt (xml: string, sdtTag: string, value: string): { xml: string; hit: boolean } {
  const tagEsc = escapeXml(sdtTag)
  // sdt 可能带命名空间 w:tag，属性顺序不定 ⇒ 用宽松匹配
  const re = new RegExp(
    `(<w:sdt>(?:(?!</w:sdt>)[\\s\\S])*?<w:tag w:val="${tagEsc}"\\s*/>(?:(?!</w:sdt>)[\\s\\S])*?<w:sdtContent>)` +
    `([\\s\\S]*?)(</w:sdtContent>)`,
  )
  const m = re.exec(xml)
  if (!m) return { xml, hit: false }
  // 只替换 sdtContent 里的第一个 w:t，其余节点原样保留（不重建结构）
  const inner = m[2]
  const wt = /(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/.exec(inner)
  if (!wt) return { xml, hit: false }
  const newInner = inner.replace(wt[0], `${wt[1]}${escapeXml(value)}${wt[3]}`)
  const next = xml.replace(m[0], `${m[1]}${newInner}${m[3]}`)
  return { xml: next, hit: true }
}

/* ── ③b customXml 数据源（层②外部数据关联）───────────────────── */

/** sites → customXml part 的 XML（命名空间固定，xpath=/ooRefs/entry@k） */
export function buildCustomXmlData (
  sites: OoRefSite[],
  input: { pick: (ref: OoRef) => unknown; labelOf?: (ref: OoRef) => string },
): string {
  const rows: string[] = []
  const seen = new Set<string>()
  for (const s of sites) {
    const tag = s.sdtTag ?? (input.labelOf ? input.labelOf(s.ref) : '')
    if (!tag || seen.has(tag)) continue
    seen.add(tag)
    const v = fmtCellValue(input.pick(s.ref))
    rows.push('    <entry k="' + escapeXml(tag) + '">' + escapeXml(v) + '</entry>')
  }
  const NL = String.fromCharCode(10)
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + NL +
    '<ooRefs xmlns="urn:client-base:ooRefs">' + NL +
    rows.join(NL) + (rows.length ? NL : '') +
    '</ooRefs>' + NL
}

/* ── ④ 主入口：字节进 → 字节出 ────────────────────────────────── */

export interface WritebackInput {
  /** 当前 OO 文件字节（从模板库读出） */
  bytes: Uint8Array
  /** oosheet 或 ooword */
  kind: 'oosheet' | 'ooword'
  /** 域落点（discover 阶段给出） */
  sites: OoRefSite[]
  /** 取值：从表快照按引用取值 */
  pick: (ref: OoRef) => unknown
  /** Word：可读标签（当 sdtTag 缺失时用它匹配） */
  labelOf?: (ref: OoRef) => string
}

/**
 * ★ 补丁式写回：只替换改动项，其余部件原样保留。
 * 绝不用「load → 自有模型 → 全量写回」—— 那会静默丢图表/绘图/透视表。
 */
export async function writebackOo (input: WritebackInput): Promise<OoWritebackResult> {
  const { bytes, kind, sites, pick } = input
  const { cells, skipped } = kind === 'oosheet'
    ? planWrites(sites, pick)
    : { cells: [], skipped: [] }

  const srcZip = await JSZip.loadAsync(bytes)
  const outZip = new JSZip()

  // 先把**原文件所有部件**原样搬过去（补丁的底）
  const srcNames: string[] = []
  for (const name of Object.keys(srcZip.files)) {
    if (srcZip.files[name].dir) continue
    srcNames.push(name)
    outZip.file(name, await srcZip.files[name].async('uint8array'))
  }

  const labels: string[] = []
  let written = 0

  if (kind === 'oosheet') {
    // 按 sheet 分组，避免同一 sheet 反复 load/save
    const bySheet = new Map<string, { a1: string; value: string }[]>()
    for (const c of cells) {
      const sn = sheetNameOf(c.a1, sites)
      if (!sn) { skipped.push({ fieldId: c.a1, why: '无法确定所属工作表' }); continue }
      if (!bySheet.has(sn)) bySheet.set(sn, [])
      bySheet.get(sn)!.push(c)
    }
    for (const [sheetFile, list] of bySheet) {
      // ★ sheetFile 已是**完整部件路径**（xl/worksheets/sheet1.xml，来自
      //   ooxmlSheetMap 的两跳解析）⇒ 直接用，**不能再拼** `xl/worksheets/`。
      //   否则会得到 xl/worksheets/xl/worksheets/sheet1.xml（双前缀）⇒ 找不到文件。
      //   同时兼容只有 'sheet1.xml' 的旧数据（按 in 结尾判断）。
      const path = sheetFile.startsWith('xl/') ? sheetFile : `xl/worksheets/${sheetFile}`
      if (!srcZip.file(path)) {
        for (const c of list) skipped.push({ fieldId: c.a1, why: `找不到 ${path}` })
        continue
      }
      let xml = await srcZip.file(path)!.async('string')
      for (const c of list) {
        const r = patchSheetXml(xml, c.a1, c.value)
        if (r.hit) { xml = r.xml; written++; labels.push(`${sheetFile}!${c.a1}=${c.value}`) }
        else skipped.push({ fieldId: c.a1, why: `cell ${c.a1} 不在 ${sheetFile} 里` })
      }
      outZip.file(path, xml)
    }
  } else {
    const path = 'word/document.xml'
    if (!srcZip.file(path)) {
      skipped.push({ fieldId: 'document', why: `找不到 ${path}` })
    } else {
      let xml = await srcZip.file(path)!.async('string')
      for (const s of sites) {
        const tag = s.sdtTag ?? (input.labelOf ? input.labelOf(s.ref) : '')
        if (!tag) { skipped.push({ fieldId: s.ref.fieldId, why: '无 sdtTag' }); continue }
        const r = patchDocxSdt(xml, tag, fmtCellValue(pick(s.ref)))
        if (r.hit) { xml = r.xml; written++; labels.push(`${tag}=…`) }
        else skipped.push({ fieldId: s.ref.fieldId, why: `找不到 sdt tag=${tag}` })
      }
      outZip.file(path, xml)
      /* ★ 2026-10-22 层② customXml 数据源注入（用户指令「帮我实现」）：
         把 sites 的 (tag→值) 序列化成 customXml part 写进 docx——
         Word/OnlyOffice 里 SetDataBinding 绑定的控件按 xpath 读到最新值，
         「外部数据关联」持久化在文件本体（不依赖画布在场）。幂等（同名覆盖）。 */
      if (written > 0) {
        try {
          outZip.file('customXml/ooRefs.xml', buildCustomXmlData(sites, input))
          const relsPath = 'word/_rels/document.xml.rels'
          const relsXml = srcZip.file(relsPath)
            ? await srcZip.file(relsPath)!.async('string')
            : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
          if (!relsXml.includes('customXml/ooRefs.xml')) {
            const addRel = '<Relationship Id="ooRefData1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/ooRefs.xml"/>'
            outZip.file(relsPath, relsXml.replace('</Relationships>', addRel + '</Relationships>'))
          }
          const ctPath = '[Content_Types].xml'
          const ctXml = srcZip.file(ctPath) ? await srcZip.file(ctPath)!.async('string') : ''
          if (ctXml && !ctXml.includes('customXml/ooRefs.xml')) {
            const addCt = '<Override PartName="/customXml/ooRefs.xml" ContentType="application/xml"/>'
            outZip.file(ctPath, ctXml.replace('</Types>', addCt + '</Types>'))
          }
          labels.push('customXml 数据源已注入')
        } catch { /* 注入失败不阻断值写回 */ }
      }
    }
  }

  const out = await outZip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  // restored恒为空：本实现是「全量搬原文件 + 定点替换」，
  //     天然零丢失。保留字段是为了将来若改用 office-kit 产物路径时有证据。
  return { bytes: out, written, labels, restored: [], skipped }
}

/**
 * ★ 取 site 的工作表部件路径。
 *
 * 口径（2026-10-29 修正）：`OoRefSite.part` 存的是**部件相对路径**
 * （`xl/worksheets/sheet1.xml`），由 core/ooxmlSheetMap 从
 * workbook.xml + rels 两跳解析而来。
 * ⇒ **不按序号猜、不在这里补 .xml** —— 猜错会写到别的表（用户看到"值写进错误的表"）。
 * 上游没给 part 时返回 null，由调用方如实报 skipped。
 */
function sheetNameOf (a1: string, sites: OoRefSite[]): string | null {
  const s = sites.find((x) => x.a1 === a1)
  if (!s || !s.part) return null
  return s.part
}
