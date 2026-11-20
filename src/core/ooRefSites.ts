/**
 * OO 域落点解析 ——把「域清单 ooRefs」与「文件里的实际位置 a1/sdtTag」对上
 *
 * 立档：2026-10-06 · 配套 src/core/ooWriteback.ts（写回需要 OoRefSite）
 *
 * ★ 为什么必须有这个：写回要知道「写到哪个格子」，而 ooRefs 只说「引用了哪个字段」。
 *   两边对不上就没法写 —— 这是写回管线的硬前提。
 *
 * ★ Excel 侧本来就有落点信息（XlsxFieldSpec.key + ref），本文件就是把它接过来：
 *   ooRefs.urn 里的 key（域名）⇄ XlsxFieldSpec.key
 *   XlsxFieldSpec.ref            ⇒ a1（单格）或区域（table）
 *
 * 纯计算层：零 IO / 零 electron / 零 document。
 */

import type { XlsxFieldSpec } from '../xlsx/field'
import type { OoRef } from './ooRef'
import type { OoRefSite } from './ooWriteback'

/** 从 URN 取 locator（= 域名，对应 XlsxFieldSpec.key） */
export function locatorOf (urn: string): string {
  const i = urn.indexOf('#')
  return i >= 0 ? urn.slice(i + 1) : urn
}

/**
 * ★ 把 ref 规整成「能直接写」的落点。
 *
 * 坑（XlsxFieldSpec 的 ref 有两种形态）：
 *   definedName → 'Sheet1!$B$2'  （**带表名**）
 *   table       → 'A1:C4'        （**只有范围**）
 *   无脑拼 `${sheet}!${ref}` 会得到 'Sheet1!Sheet1!$B$2' —— 重复表名。
 *   ⇒ ref 已经带 '!' 就不再补前缀（与 xlsx/domainMarks.displayRef 同一坑）。
 */
export function normalizeRef (sheet: string | undefined, ref: string): { part: string; a1: string } | null {
  const r = String(ref ?? '').trim()
  if (!r) return null
  const bang = r.lastIndexOf('!')
  if (bang >= 0) {
    const part = r.slice(0, bang).replace(/^'|'$/g, '') || (sheet || '')
    const a1 = r.slice(bang + 1).replace(/\$/g, '')      // 去绝对引号
    // 区域（A1:C4）不能当单格写 ⇒ 取左上角
    return { part, a1: a1.split(':')[0] }
  }
  // 纯 A1 或区域 ⇒ 用 sheet（或默认第一张）
  const a1 = r.replace(/\$/g, '').split(':')[0]
  return { part: sheet || '', a1 }
}

/** 工作表名 → OOXML 里的文件名（sheet1.xml）。
 *  ★ 表名与文件名**不是一回事**（「汇总表」对应 sheetN.xml，N 由顺序决定），
 *  拿不到映射时留空，让调用方按「只有一张表」兜底。 */
export function sheetPartOf (sheet: string | undefined, index: number): string {
  // 保守：只有当 sheet 名本身就是 sheetN 形式时才直接用
  const m = /^sheet(\d+)$/i.exec(String(sheet || ''))
  if (m) return 'sheet' + m[1] + '.xml'
  return index >= 0 ? 'sheet' + (index + 1) + '.xml' : ''
}

/**
 * ★核心：ooRefs × 文件落点清单 ⇒ OoRefSite[]（写回管线要的）。
 *
 * @param refs     域清单（关联了哪些上游字段）
 * @param fields   文件里发现的落点（XlsxFieldSpec[] 或 Word 的 sdt 列表）
 * @param kind     oosheet / ooword
 * @param opts.wordSites Word 的内容控件落点：[{ tag, refKey? }]
 * @returns 解析结果 + 未匹配上的引用（如实报告，不静默丢）
 */
export function resolveOoRefSites (
  refs: OoRef[],
  fields: XlsxFieldSpec[],
  kind: 'oosheet' | 'ooword',
  opts?: {
    /** ★ fieldId → 字段名（对齐 XlsxFieldSpec.key 的桥梁，见下） */
    nameOf?: (fieldId: string) => string
    /** Word：sdt 落点（tag → 对应的域 key） */
    wordSites?: { tag: string; key?: string; part?: string }[]
    /** 工作表名→文件名 的映射（能拿到时优先用） */
    sheetPartMap?: Record<string, string>
  },
): { sites: OoRefSite[]; unmatched: { fieldId: string; why: string }[] } {
  const sites: OoRefSite[] = []
  const unmatched: { fieldId: string; why: string }[] = []
  // key → 落点（同key 多处时取第一处，与 domainMarks 口径一致）
  const byKey = new Map<string, XlsxFieldSpec>()
  for (const f of fields) if (f && f.key && !byKey.has(f.key)) byKey.set(f.key, f)

  for (const ref of refs) {
    const locator = locatorOf(ref.urn ?? `node://${ref.nodeId}#${ref.fieldId}`)
    // ★★ 候选键：域名(urn locator) → 字段名 → fieldId
    //   `XlsxFieldSpec.key` 是**命名区域的名字**（discoverXlsxFields 里 `key: n.name`，
    //   形如「数量」），而 `OoRef.fieldId` 是**节点字段 id**（形如 f1）——
    //   两者不是同一套标识，只按 fieldId 去对 key **永远匹配不上**。
    //   既有绑定 bindXlsxFromParamTable 也是按 `data[f.key]`（字段名）对齐的。
    const fieldName = opts?.nameOf?.(ref.fieldId) ?? ''
    const candidates = [locator, fieldName, ref.fieldId].filter(Boolean) as string[]

    if (kind === 'ooword') {
      const ws = opts?.wordSites?.find((w) => w.key != null && candidates.includes(w.key))
        ?? opts?.wordSites?.find((w) => candidates.includes(w.tag))
      if (!ws) {
        unmatched.push({ fieldId: ref.fieldId, why: `文件里找不到 key=${candidates[0] ?? ref.fieldId} 的内容控件` })
        continue
      }
      sites.push({ ref, sdtTag: ws.tag, part: ws.part })
      continue
    }

    // oosheet
    let f: XlsxFieldSpec | undefined
    for (const c of candidates) {
      const got = byKey.get(c)
      if (got) { f = got; break }
    }
    if (!f) {
      unmatched.push({
        fieldId: ref.fieldId,
        why: `文件里找不到域（试过：${candidates.join(' / ') || '无'}）`,
      })
      continue
    }
    const n = normalizeRef(f.sheet, f.ref)
    if (!n || !n.a1) { unmatched.push({ fieldId: ref.fieldId, why: `域的 ref 无法解析（${f.ref}）` }); continue }
    const part = opts?.sheetPartMap?.[f.sheet ?? ''] ?? sheetPartOf(f.sheet, fields.indexOf(f))
    sites.push({ ref, a1: n.a1, part })
  }
  return { sites, unmatched }
}

/**
 * ★ 供 UI 显示：把落点解析结果转成人类可读清单。
 * 用途：诊断抽屉 /「查看域清单」面板，让用户知道每个引用写在哪。
 */
export function describeSites (sites: OoRefSite[]): string[] {
  return sites.map((s) => {
    if (s.sdtTag) return 'Word·内容控件 ' + s.sdtTag
    if (s.part && s.a1) return s.part.replace(/\.xml$/, '') + '!' + s.a1
    if (s.a1) return s.a1
    return '（无落点）'
  })
}
