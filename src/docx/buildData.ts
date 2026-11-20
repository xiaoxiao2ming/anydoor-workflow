/**
 * 把引擎算好的数据按绑定映射翻译成 bind() 需要的 data 对象。
 *
 * 这是「画布计算」与「Word 渲染」之间唯一的翻译层：
 *   画布侧只认 fieldId 和 Row[]，Word 侧只认模板 tag 名。
 *
 * 引用模型与 Excel 侧**共用** src/ref/context.ts：
 *   给了上游上下文（opts.ub）→ 每条绑定都能算出 URN，
 *   于是能答「这个数从哪来」，也能直接引用外部文件 / API / 全局参数。
 *   不给 → 退回纯 fieldId 模式（老调用点与单测走这条）。
 */

import type { DocxBinding, FieldDef, TableSchema } from '../core/schema'
import type { Row } from '../core/store'
import type { BindingInfo } from './binder'
import type { RefContext } from '../ref/urn'
import { fmtValue, parseUrn, resolveUrn } from '../ref/urn'
import type { ResolveTrace } from '../ref/paramTable'
import {
  aggregate,
  buildUpstream,
  resolveBindingTarget,
  type Agg,
  type UpstreamBuild,
} from '../ref/context'

export interface BuildDataResult {
  data: Record<string, unknown>
  /** 每个 tag 的取值说明，供 UI 展示「这个控件取到了什么」 */
  resolved: Record<string, string>
  missing: string[]
  /**
   * 逐域溯源：值来自哪个 URN、解析成功与否。
   * 送审被问「这个数哪来的」时要能一键答出来 —— 这是它存在的唯一理由。
   */
  traces: ResolveTrace[]
}

export interface BuildDataOptions {
  /** 上游上下文。有它才走 URN；没有则退回 fieldId 模式（老签名兼容）。 */
  ub?: UpstreamBuild
}

/** 按名字或 id 找字段；找不到返回 null */
function resolveField(fields: FieldDef[], key: string): FieldDef | null {
  return fields.find((f) => f.id === key || f.name === key) ?? null
}

/**
 * 富文本里的 {{字段名}} 动态引用：结论段里写「本工程全长 {{管线总长}} 米」，
 * 生成时替换为上游算好的值。找不到字段时保留原文并提示。
 */
function interpolateRich(html: string, fields: FieldDef[], rows: Row[], missing: string[]): string {
  return html.replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (whole, key: string) => {
    const f = resolveField(fields, key)
    if (!f) {
      missing.push(`rich 内引用的字段「${key}」在上游不存在`)
      return whole
    }
    return escLike(aggregate(rows, f.id, 'first'))
  })
}

/** 富文本是 HTML，值要按 HTML 转义，防止数字/文本里的 < & 破坏结构 */
function escLike(v: unknown): string {
  return String(v ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** 一条溯源记录 */
function trace(key: string, urn: string | null, by: string, value: unknown, ok: boolean, error?: string): ResolveTrace {
  const u = urn ? parseUrn(urn) : null
  return {
    key,
    urn: urn ?? (by === 'none' ? '(未绑定)' : by === 'auto' ? '(按名自动匹配)' : '(手工值)'),
    from: u ? u.scheme : 'default',
    value,
    ok,
    error,
  }
}

/**
 * 解析一个域指向哪个字段。
 * 有上游上下文走共享的 resolveBindingTarget（支持 source / fieldId / 自动匹配），
 * 没有就退回「在传入的 fields 里按名字或 id 找」。
 */
function pickTarget(
  spec: DocxBinding | undefined,
  key: string,
  fields: FieldDef[],
  ub: UpstreamBuild | undefined,
): { fieldId: string | null; urn: string | null; by: 'source' | 'fieldId' | 'auto' | 'none' } {
  if (ub) return resolveBindingTarget(spec, key, ub)
  const fid = spec?.fieldId ?? resolveField(fields, key)?.id ?? null
  return { fieldId: fid, urn: null, by: spec?.fieldId ? 'fieldId' : fid ? 'auto' : 'none' }
}

/** 整表取值时用哪一批行：URN 指向某个上游就用那个上游的，否则用传入的 */
function rowsFor(urn: string | null, ub: UpstreamBuild | undefined, fallback: Row[]): Row[] {
  if (!ub || !urn) return fallback
  const m = /^node:\/\/([^#]+)/.exec(urn)
  if (!m) return fallback
  return ub.computed.find((c) => c.id === m[1])?.rows ?? fallback
}

/**
 * 按绑定映射把上游数据算成 data。
 *
 * @param bindings parseTemplate 解析出来的控件清单
 * @param map      用户配置的 tag → 取值来源；未配置的按「tag 名 == 字段名」自动匹配
 * @param fields   上游字段定义（computeNode 返回的）
 * @param rows     上游数据行（computeNode 返回的）
 * @param opts     上游上下文（可选；给了才能走 URN + 溯源）
 */
export function buildData(
  bindings: BindingInfo[],
  map: Record<string, DocxBinding>,
  fields: FieldDef[],
  rows: Row[],
  opts: BuildDataOptions = {},
): BuildDataResult {
  const ub = opts.ub
  const ctx: RefContext = ub?.ctx ?? {}
  const data: Record<string, unknown> = {}
  const resolved: Record<string, string> = {}
  const missing: string[] = []
  const traces: ResolveTrace[] = []

  const fieldById = new Map(fields.map((f) => [f.id, f]))

  for (const b of bindings) {
    const spec = map[b.tag]
    const name = b.path

    // ── each：整张表塞进去 ────────────────────────────────────
    if (b.kind === 'each') {
      let tgt = pickTarget(spec, name, fields, ub)
      // each 不配绑定也能工作（默认吃「全部上游」）。这时不能标成「未绑定」——
      // 明明取了数据，标成未绑定是自相矛盾的显示。
      const allUpstream = tgt.by === 'none' && !!ub && ub.computed.length > 0
      if (allUpstream && ub!.computed.length === 1) {
        tgt = { fieldId: null, urn: `node://${ub!.computed[0].id}#rows`, by: 'auto' }
      }
      const src = rowsFor(tgt.urn, ub, rows)
      const cols = spec?.columns
      const mapped =
        cols && Object.keys(cols).length
          ? // 列映射：模板占位符名 → 上游字段 id。把每行的 key 换成模板里的名字
            src.map((r) => Object.fromEntries(Object.entries(cols).map(([tplKey, fid]) => [tplKey, r[fid] ?? ''])))
          : src
      data[name] = mapped
      const withCols = !!(cols && Object.keys(cols).length)
      const tail = withCols ? '（已映射列）' : ''
      const from =
        !allUpstream || ub!.computed.length === 1
          ? tgt.urn
            ? ` ← ${tgt.urn}`
            : ''
          : ` ← 全部上游合并（${ub!.computed.length} 个）`
      resolved[b.tag] = `${mapped.length} 行${tail}${from}`
      traces.push(trace(b.tag, tgt.urn, tgt.by, mapped, true))
      continue
    }

    // ── 手工值 ────────────────────────────────────────────────
    if (spec?.from === 'manual') {
      const v = spec.value ?? ''
      if (b.kind === 'if') {
        // 手工开关：UI 存的是 'true' / 'false' 字符串
        data[name] = v === 'true' || v === '1' || v === '是' || v === 'on'
      } else if (b.kind === 'rich') {
        // 富文本：支持 {{字段名}} 动态引用，结论段里能引用上游算好的数字
        data[name] = interpolateRich(v, fields, rows, missing)
      } else {
        data[name] = v
      }
      resolved[b.tag] = b.kind === 'rich' ? '手工富文本' : `手工值：${String(v).slice(0, 12) || '（空）'}`
      traces.push(trace(b.tag, null, 'manual', data[name], true))
      continue
    }

    // ── 解析取值目标 ──────────────────────────────────────────
    const tgt = pickTarget(spec, name, fields, ub)

    // 非 node:// 的 URN（外部文件 / API / 全局参数 / 常量）：
    // 本地没有对应字段，值只能靠 resolveUrn 取。
    if (tgt.by === 'source' && tgt.fieldId === null && tgt.urn && !tgt.urn.startsWith('node://')) {
      const v = resolveUrn(tgt.urn, ctx)
      if (v === undefined) {
        missing.push(b.tag)
        resolved[b.tag] = `引用未解析：${tgt.urn}`
        traces.push(trace(b.tag, tgt.urn, tgt.by, undefined, false, '引用未解析到值（外部数据是否已加载？）'))
        continue
      }
      data[name] = b.kind === 'rich' ? escLike(v) : v
      resolved[b.tag] = `${fmtValue(v)} ← ${tgt.urn}`
      traces.push(trace(b.tag, tgt.urn, tgt.by, v, true))
      continue
    }

    if (!tgt.fieldId) {
      missing.push(b.tag)
      resolved[b.tag] = spec ? '字段已失效' : '未绑定'
      traces.push(trace(b.tag, tgt.urn, tgt.by, undefined, false, spec ? '绑定的字段在上游不存在' : '未绑定'))
      continue
    }

    const fid = tgt.fieldId
    const f = fieldById.get(fid) ?? (ub ? ub.fields.find((x) => x.id === fid) : undefined)
    if (!f) {
      missing.push(b.tag)
      resolved[b.tag] = '字段已失效'
      traces.push(trace(b.tag, tgt.urn, tgt.by, undefined, false, `字段 ${fid} 在上游不存在`))
      continue
    }

    const src = rowsFor(tgt.urn, ub, rows)

    if (b.kind === 'if') {
      // 条件段：取首行的布尔语义（非空 / 非 0 / 非 "false" 即真）
      const raw = src.length ? src[0][f.id] : ''
      const truthy = !(raw === '' || raw == null || raw === 0 || raw === false || raw === 'false' || raw === '否')
      data[name] = truthy
      resolved[b.tag] = `${f.name} → ${truthy ? '真' : '假'}${tgt.urn ? ` ← ${tgt.urn}` : ''}`
      traces.push(trace(b.tag, tgt.urn, tgt.by, truthy, true))
      continue
    }

    const v = aggregate(src, f.id, spec?.agg as Agg | undefined)
    data[name] = v
    const aggSuffix = src.length > 1 && spec?.agg && spec.agg !== 'first' ? `（${spec.agg}）` : ''
    const autoSuffix = tgt.by === 'auto' ? '（自动匹配）' : ''
    resolved[b.tag] = tgt.urn
      ? `${f.name}${aggSuffix}${autoSuffix} ← ${tgt.urn}`
      : `${f.name}${aggSuffix}${autoSuffix}`
    traces.push(trace(b.tag, tgt.urn, tgt.by, v, true))
  }

  return { data, resolved, missing, traces }
}

/** 便捷入口：直接由画布节点构造上游上下文再构建（WordReportPanel 用） */
export function buildDataFromNode(
  bindings: BindingInfo[],
  map: Record<string, DocxBinding>,
  node: TableSchema,
  tables: Record<string, TableSchema>,
  records: Record<string, Row[]>,
): BuildDataResult & { ub: UpstreamBuild } {
  const ub = buildUpstream(node, tables, records)
  // fields/rows 用上游并集 —— 与 computeNode 的行为保持一致（多上游时取并集）
  const fields: FieldDef[] = ub.fields
  const rows: Row[] = ub.computed.length === 1 ? ub.computed[0].rows : ub.computed.flatMap((c) => c.rows)
  return { ...buildData(bindings, map, fields, rows, { ub }), ub }
}
