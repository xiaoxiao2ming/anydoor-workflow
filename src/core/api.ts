// 集中存储 API 客户端：所有画布页面读写都走后端（同源 /api），
// 多浏览器/多设备访问同一份数据，解决「不同浏览器布局不一致」问题。

export interface PageMeta {
  id: string
  name: string
}

export interface PageData {
  id: string
  name: string
  tables: Record<string, unknown>
  records: Record<string, unknown>
  positions: Record<string, unknown>
}

async function req(path: string, method: string, body?: unknown): Promise<Response> {
  const r = await fetch(path, {
    method,
    headers: body !== undefined ? { 'Content-Type': 'application/json' } : undefined,
    body: body !== undefined ? JSON.stringify(body) : undefined,
  })
  /* ★ 2026-10-17 口令过期/被改（网页端 APP_PASSWORD）：接口回 401 时把用户送回登录页，
   *   而不是让画布停在「加载中」或抛一个看不懂的 listPages failed: 401。
   *   只在浏览器里跳（Node/vitest 没有 location）。 */
  if (r.status === 401 && typeof window !== 'undefined' && window.location) {
    const next = window.location.pathname + window.location.search
    window.location.href = '/login?next=' + encodeURIComponent(next)
  }
  return r
}

export async function listPages(): Promise<PageMeta[]> {
  const r = await req('/api/pages', 'GET')
  if (!r.ok) throw new Error('listPages failed: ' + r.status)
  return (await r.json()) as PageMeta[]
}

export async function getPage(id: string): Promise<PageData> {
  const r = await req(`/api/pages/${encodeURIComponent(id)}`, 'GET')
  if (!r.ok) throw new Error('getPage failed: ' + r.status)
  return (await r.json()) as PageData
}

export async function savePage(
  id: string,
  data: { name?: string; tables: unknown; records: unknown; positions: unknown },
): Promise<void> {
  const r = await req(`/api/pages/${encodeURIComponent(id)}`, 'PUT', data)
  if (!r.ok) throw new Error('savePage failed: ' + r.status)
}

export async function createPage(name?: string): Promise<PageMeta> {
  const r = await req('/api/pages', 'POST', { name })
  if (!r.ok) throw new Error('createPage failed: ' + r.status)
  return (await r.json()) as PageMeta
}

export async function deletePage(id: string): Promise<void> {
  const r = await req(`/api/pages/${encodeURIComponent(id)}`, 'DELETE')
  if (!r.ok) throw new Error('deletePage failed: ' + r.status)
}
