/**
 * 引用协议层 · 参数表（单一真相源）
 *
 * 纯计算层，零依赖。
 * 参数表是整张画布「唯一说得出每个值从哪来」的地方，因此溯源是一等能力。
 */

import { parseUrn, resolveUrn, fmtValue, type RefContext } from './urn'

/** 参数规格：参数表的一行 */
export interface ParamSpec {
  key: string
  label?: string
  unit?: string
  /** 数字格式码（Excel numFmt 同款），如 '0.00'、'#,##0.00' */
  format?: string
  /** 引用地址；缺省则用 default */
  source?: string
  default?: unknown
}

export interface ResolveTrace {
  key: string
  urn: string
  from: 'override' | 'default' | 'node' | 'file' | 'api' | 'param' | 'const' | 'unresolved'
  value: unknown
  ok: boolean
  error?: string
}

export class ParamTable {
  specs: ParamSpec[] = []
  /** 手填值，优先级最高（盖过 source） */
  overrides: Record<string, unknown> = {}
  ctx: RefContext = {}

  constructor(specs: ParamSpec[] = [], ctx: RefContext = {}) {
    this.specs = specs
    this.ctx = ctx
  }

  setContext(ctx: RefContext): void {
    this.ctx = ctx
  }

  /** 手填 > source 引用 > default */
  resolve(key: string): unknown {
    if (key in this.overrides) return this.overrides[key]
    const spec = this.specs.find((s) => s.key === key)
    if (!spec) return undefined
    if (spec.source) {
      const v = resolveUrn(spec.source, this.ctx)
      if (v !== undefined) return v
    }
    return spec.default
  }

  /** 解析全部，产出可直接喂给下游节点的数据对象 */
  resolveAll(): Record<string, unknown> {
    const out: Record<string, unknown> = {}
    for (const s of this.specs) out[s.key] = this.resolve(s.key)
    for (const k of Object.keys(this.overrides)) if (!(k in out)) out[k] = this.overrides[k]
    return out
  }

  /** 溯源：每个参数的值从哪来 */
  trace(): ResolveTrace[] {
    return this.specs.map((s) => {
      if (s.key in this.overrides) {
        return { key: s.key, urn: '(手填)', from: 'override', value: this.overrides[s.key], ok: true }
      }
      if (s.source) {
        const u = parseUrn(s.source)
        const v = resolveUrn(s.source, this.ctx)
        if (v !== undefined) {
          return { key: s.key, urn: s.source, from: u ? u.scheme : 'unresolved', value: v, ok: true }
        }
        return {
          key: s.key,
          urn: s.source,
          from: u ? u.scheme : 'unresolved',
          value: s.default,
          ok: s.default !== undefined,
          error:
            s.default === undefined
              ? '引用未解析到值，且无默认值'
              : '引用未命中，已回退到默认值',
        }
      }
      return {
        key: s.key,
        urn: '(默认)',
        from: s.default !== undefined ? 'default' : 'unresolved',
        value: s.default,
        ok: s.default !== undefined,
      }
    })
  }

  /** 按 format / unit 渲染显示值 */
  format(key: string): string {
    const v = this.resolve(key)
    if (v == null) return ''
    const spec = this.specs.find((s) => s.key === key)
    let s = fmtValue(v)
    if (typeof v === 'number' && spec?.format) {
      const decimals = (spec.format.split('.')[1] || '').replace(/[^0#]/g, '').length
      s = v.toFixed(decimals)
    }
    if (spec?.unit) s += ' ' + spec.unit
    return s
  }
}
