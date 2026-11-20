/**
 * Word 侧域适配器
 *
 * 引用协议本身在 src/ref/（与节点类型无关）；这里只做 Word 特有的两件事：
 *   ① 槽位发现 discoverFields —— 从 w:sdt 的 tag 读出域清单
 *   ② 填充 bindFromParamTable —— 把参数表的值写进 sdtContent
 *
 * 为保持向后兼容，本文件把引用协议层的内容再导出一遍，
 * 老的 `import { ParamTable } from './field'` 之类不受影响。
 */

import { parseTemplate, bind } from './binder'
import { ParamTable, type ResolveTrace } from '../ref/paramTable'

/* ── 引用协议层再导出（真实实现在 src/ref/） ── */
export { parseUrn, resolveUrn, pickByLocator, getPath, fmtValue } from '../ref/urn'
export type { RefContext, ResolvedRef, RefScheme } from '../ref/urn'
export { ParamTable } from '../ref/paramTable'
export type { ParamSpec } from '../ref/paramTable'
export { RefGraph } from '../ref/refGraph'
export type {
  PortSpec,
  SlotSpec,
  Binding,
  NodeDef,
  GraphTrace,
  GraphResult,
} from '../ref/refGraph'
export { parseTag } from './binder'

/* ── Word 特有 ── */

export type BindKind = 'field' | 'rich' | 'each' | 'if' | 'image'

/** 文档里的槽位 */
export interface FieldSpec {
  key: string
  kind: BindKind
  part: string
  hasRow: boolean
  sample: string
  placeholders: string[]
}

/** 从 docx 模板发现所有域 */
export async function discoverFields(input: ArrayBuffer | Uint8Array): Promise<{
  fields: FieldSpec[]
  bindings: unknown[]
}> {
  const tpl = await parseTemplate(input)
  const fields: FieldSpec[] = (tpl.bindings || []).map((b: any) => ({
    key: b.path,
    kind: b.kind,
    part: b.part,
    hasRow: !!b.hasRow,
    sample: b.sample || '',
    placeholders: b.placeholders || [],
  }))
  return { fields, bindings: tpl.bindings }
}

/** 参数表 → 填充 Word 模板 */
export async function bindFromParamTable(
  input: ArrayBuffer | Uint8Array,
  table: ParamTable,
  opts: { stripControls?: boolean } = {}
): Promise<{ buffer: Uint8Array; bound: number; warnings: string[]; traces: ResolveTrace[] }> {
  const data = table.resolveAll()
  const out = await bind(input, data as Record<string, unknown>, opts)
  return { buffer: out.buffer, bound: out.bound, warnings: out.warnings, traces: table.trace() }
}
