/**
 * 参数引用标识（单一真源 · 零依赖）
 *
 * 立档：2026-10-17 · 从 ooDomains.ts 拆出。
 * 原因：引用面板（RefPanel）只需要算「这个参数叫什么名字 / 会是哪个颜色」，
 *   不该为此把 jszip 拖进面板的 chunk。这里是纯字符串运算，谁都能引。
 *
 * ★ 与 scripts/oo-core.mjs 的口径必须**逐位一致**（否则定义名对不上格子，引用格显示空）：
 *     rawKey = `<节点名>.<字段名>`
 *     name   = rawKey 里 '.'→'_'、非法字符→'_'，首字符是数字/下划线则加 'R'
 *     row    = 2 + djb2(rawKey) % 1998
 *   改任一处都要同步另一处。
 */

/** rawKey：节点名.字段名 */
export function domainRawKey (nodeName: string, fieldName: string): string {
  const n = String(nodeName ?? '').trim()
  const f = String(fieldName ?? '').trim()
  if (!n && !f) return ''
  return f ? `${n}.${f}` : n
}

/**
 * 定义名称：与 oo-core `__doSheetWrite` 内联的正则完全等价
 *   rawKey.replace(/[.]/g,'_').replace(/[^0-9A-Za-z_一-龥_]/g,'_')，首字符数字加 'R'
 * ★ 这里用字面量字符类而不是 \w —— oo-core 是模板字符串产物，\w 的反斜杠会被吞（历史惨案）。
 */
export function domainNameOf (rawKey: string): string {
  let name = String(rawKey ?? '').replace(/\./g, '_').replace(/[^0-9A-Za-z_\u4e00-\u9fa5_]/g, '_')
  if (/^[0-9_]/.test(name)) name = 'R' + name
  return name
}

/** 行号：djb2 后落 2..1999（与 oo-core 同一算法、同一取模） */
export function domainRowOf (rawKey: string): number {
  let h = 5381
  for (let i = 0; i < rawKey.length; i++) h = ((h << 5) + h + rawKey.charCodeAt(i)) & 0x7fffffff
  return 2 + (h % 1998)
}
