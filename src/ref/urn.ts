/**
 * 引用协议层 · 统一引用地址（URN）
 *
 * 纯计算层，零依赖 —— 不 import 任何 IO 库，浏览器可直接跑。
 * 引用与「节点是什么类型」无关：Word、Excel、计算、图表……都只是节点。
 */

export type RefScheme = 'node' | 'file' | 'api' | 'param' | 'const'

/** 已解析好的外部数据（由平台适配层预加载后塞进来） */
export interface RefContext {
  /** 画布节点输出：key = nodeId */
  nodes?: Record<string, unknown>
  /** 全局参数：key = 参数名 */
  params?: Record<string, unknown>
  /** 外部文件：key = 绝对路径，value = 该文件解析出的结构化数据 */
  files?: Record<string, unknown>
  /** 网络 API：key = sourceId */
  apis?: Record<string, unknown>
}

export interface ResolvedRef {
  scheme: RefScheme
  target: string
  locator: string
}

/**
 * 统一引用地址：
 *   node://<nodeId>#<locator>      画布节点
 *   file://<abs path>#<locator>    外部本地文件
 *   api://<sourceId>/<path>        网络 API
 *   param://<name>                 全局参数
 *   const:<value>                  常量
 */
export function parseUrn(urn: string): ResolvedRef | null {
  const s = String(urn ?? '').trim()
  let m = /^node:\/\/([^#]+)(?:#(.*))?$/.exec(s)
  if (m) return { scheme: 'node', target: m[1], locator: m[2] || '' }
  m = /^file:\/\/([^#]+)(?:#(.*))?$/.exec(s)
  if (m) return { scheme: 'file', target: m[1], locator: m[2] || '' }
  m = /^api:\/\/([^/]+)(?:\/(.*))?$/.exec(s)
  if (m) return { scheme: 'api', target: m[1], locator: m[2] || '' }
  m = /^param:\/\/([^#]+)$/.exec(s)
  if (m) return { scheme: 'param', target: m[1], locator: '' }
  m = /^const:([\s\S]*)$/.exec(s)
  if (m) return { scheme: 'const', target: m[1], locator: '' }
  return null
}

/** 点号路径取值，支持 a.b.c 与 $.a.b */
export function getPath(data: unknown, path: string): unknown {
  let p = String(path)
  if (p.startsWith('$.')) p = p.slice(2)
  return p.split('.').reduce<unknown>((o, k) => {
    if (o == null) return undefined
    if (Array.isArray(o)) return o[Number(k)]
    return (o as Record<string, unknown>)[k]
  }, data)
}

/** 在已加载的数据里按 locator 取值 */
export function pickByLocator(data: unknown, locator: string): unknown {
  if (!locator) return data

  // ① 电子表格单元格：Sheet1!B2
  const cell = /^([^!]+)!([A-Z]+\d+)$/.exec(locator)
  if (cell) {
    const [, sheet, ref] = cell
    const d = data as Record<string, unknown> | null
    const sh =
      d && typeof d === 'object'
        ? ((d[sheet] as unknown) ?? ((d as Record<string, unknown>).sheets as Record<string, unknown> | undefined)?.[sheet])
        : undefined
    if (sh && typeof sh === 'object') {
      const v = (sh as Record<string, unknown>)[ref]
      if (v !== undefined) {
        return typeof v === 'object' && v !== null && 'value' in v
          ? (v as { value: unknown }).value
          : v
      }
    }
    return undefined
  }

  // ② 整列：col:列名
  const col = /^col:(.+)$/.exec(locator)
  if (col) {
    const rows = Array.isArray(data)
      ? data
      : ((data as Record<string, unknown> | null)?.rows as unknown[])
    if (!Array.isArray(rows)) return undefined
    return rows.map((r) =>
      r && typeof r === 'object' ? (r as Record<string, unknown>)[col[1]] : undefined
    )
  }

  // ③ 其余按路径
  return getPath(data, locator)
}

/** 解析一个 URN（同步；外部数据必须已预加载到 ctx） */
export function resolveUrn(urn: string, ctx: RefContext): unknown {
  const u = parseUrn(urn)
  if (!u) return undefined
  switch (u.scheme) {
    case 'node':
      return pickByLocator(ctx.nodes?.[u.target], u.locator)
    case 'file':
      return pickByLocator(ctx.files?.[u.target], u.locator)
    case 'api':
      return pickByLocator(ctx.apis?.[u.target], u.locator)
    case 'param':
      return ctx.params?.[u.target]
    case 'const':
      return u.target
    default:
      return undefined
  }
}

/** 值的显示化（与 Word/Excel 无关的通用实现） */
export function fmtValue(v: unknown): string {
  if (v == null) return ''
  if (v instanceof Date) return v.toISOString().slice(0, 10)
  if (typeof v === 'boolean') return v ? '是' : '否'
  if (typeof v === 'object') return JSON.stringify(v)
  return String(v)
}
