/**
 * Word 模板库存储层
 *
 * 与 persistence.ts 同一套双适配范式：
 * - Electron（npm run dev / 打包后）：经 preload IPC 落到 userData/templates/
 * - 浏览器（npm run dev:web）：落到 IndexedDB，刷新不丢
 *
 * 渲染层只认这四个方法，不关心底层在哪。
 */

import type { OpenFailure, OpenFailureSample } from '../xlsx/openError'

export interface TemplateMeta {
  id: string
  name: string
  size: number
  updatedAt: number
}

export interface TemplateStore {
  list(): Promise<TemplateMeta[]>
  save(name: string, data: Uint8Array): Promise<TemplateMeta>
  read(id: string): Promise<Uint8Array | null>
  remove(id: string): Promise<void>
}

/* ---------- 浏览器：IndexedDB ---------- */

const IDB_NAME = 'client-base'
const IDB_STORE = 'tpl'
const IDB_INDEX = 'tpl-index'

function openIdb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    // 版本必须与 editStore 一致（v3）：editStore 一旦把库升到 v3，
    // 任何更低版本的 open 都会抛 VersionError；更高版本的 upgrade 又会被
    // 这里未关闭的旧连接 blocked —— 两个模块谁先打开都必须能幂等补建全部 store。
    const req = indexedDB.open(IDB_NAME, 3)
    req.onupgradeneeded = () => {
      const db = req.result
      if (!db.objectStoreNames.contains(IDB_STORE)) db.createObjectStore(IDB_STORE)
      if (!db.objectStoreNames.contains(IDB_INDEX)) db.createObjectStore(IDB_INDEX)
      if (!db.objectStoreNames.contains('docx-edit')) db.createObjectStore('docx-edit')
    }
    req.onsuccess = () => resolve(req.result)
    req.onerror = () => reject(req.error)
  })
}

function idbGet<T>(store: string, key: string): Promise<T | undefined> {
  return openIdb().then(
    (db) =>
      new Promise<T | undefined>((resolve, reject) => {
        const tx = db.transaction(store, 'readonly')
        const req = tx.objectStore(store).get(key)
        req.onsuccess = () => resolve(req.result as T | undefined)
        req.onerror = () => reject(req.error)
      }),
  )
}

function idbPut(store: string, key: string, value: unknown): Promise<void> {
  return openIdb().then(
    (db) =>
      new Promise<void>((resolve, reject) => {
        const tx = db.transaction(store, 'readwrite')
        tx.objectStore(store).put(value, key)
        tx.oncomplete = () => resolve()
        tx.onerror = () => reject(tx.error)
      }),
  )
}

function idbDel(store: string, key: string): Promise<void> {
  return openIdb().then(
    (db) =>
      new Promise<void>((resolve, reject) => {
        const tx = db.transaction(store, 'readwrite')
        tx.objectStore(store).delete(key)
        tx.oncomplete = () => resolve()
        tx.onerror = () => reject(tx.error)
      }),
  )
}

const newId = () => `t${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`

const idbStore: TemplateStore = {
  async list() {
    return (await idbGet<TemplateMeta[]>(IDB_INDEX, 'list')) ?? []
  },
  async save(name, data) {
    const list = (await idbGet<TemplateMeta[]>(IDB_INDEX, 'list')) ?? []
    const hit = list.find((t) => t.name === name)
    const id = hit?.id ?? newId()
    const meta: TemplateMeta = { id, name, size: data.byteLength, updatedAt: Date.now() }
    await idbPut(IDB_STORE, id, data)
    const next = hit ? list.map((t) => (t.id === id ? meta : t)) : [...list, meta]
    await idbPut(IDB_INDEX, 'list', next)
    return meta
  },
  async read(id) {
    return (await idbGet<Uint8Array>(IDB_STORE, id)) ?? null
  },
  async remove(id) {
    const list = (await idbGet<TemplateMeta[]>(IDB_INDEX, 'list')) ?? []
    await idbDel(IDB_STORE, id)
    await idbPut(IDB_INDEX, 'list', list.filter((t) => t.id !== id))
  },
}

/* ---------- Electron：preload IPC ---------- */

interface ElectronTplApi {
  tplList?(): Promise<TemplateMeta[]>
  tplSave?(name: string, data: Uint8Array): Promise<TemplateMeta>
  tplRead?(id: string): Promise<Uint8Array>
  tplDelete?(id: string): Promise<void>
  tplPick?(kind?: OfficeKind): Promise<{ name: string; data: Uint8Array } | null>
  docSaveAs?(defaultName: string, data: Uint8Array, kind?: OfficeKind): Promise<string | null>
  docPdf?(html: string, defaultName: string): Promise<string | null>
  failSave?(payload: unknown, bytes: Uint8Array | null): Promise<string>
  failList?(): Promise<unknown[]>
}

/** 支持的 Office 文件种类（决定对话框过滤器与 MIME） */
export type OfficeKind = 'docx' | 'xlsx'

function electronApi(): ElectronTplApi | undefined {
  return (window as unknown as { api?: ElectronTplApi }).api
}

const electronStore: TemplateStore = {
  async list() {
    return (await electronApi()?.tplList?.()) ?? []
  },
  async save(name, data) {
    return electronApi()!.tplSave!(name, data)
  },
  async read(id) {
    return (await electronApi()?.tplRead?.(id)) ?? null
  },
  async remove(id) {
    await electronApi()?.tplDelete?.(id)
  },
}

export function getTemplateStore(): TemplateStore {
  const api = electronApi()
  if (api && typeof api.tplList === 'function' && typeof api.tplSave === 'function') {
    return electronStore
  }
  return idbStore
}

/**
 * 各 Office 文件种类的浏览器侧参数。
 * Electron 侧由主进程的 FILE_KINDS 决定，这里只服务 `dev:web` 的降级路径。
 */
const WEB_KINDS: Record<OfficeKind, { accept: string; mime: string }> = {
  docx: {
    accept: '.docx',
    mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  },
  xlsx: {
    accept: '.xlsx',
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  },
}

/** 打开系统文件对话框挑模板；浏览器环境降级为 <input type=file> */
export function pickTemplateFile(kind: OfficeKind = 'docx'): Promise<{ name: string; data: Uint8Array } | null> {
  const api = electronApi()
  if (api?.tplPick) return api.tplPick(kind)

  const spec = WEB_KINDS[kind]
  return new Promise((resolve) => {
    const el = document.createElement('input')
    el.type = 'file'
    el.accept = spec.accept
    el.onchange = async () => {
      const f = el.files?.[0]
      if (!f) return resolve(null)
      resolve({ name: f.name, data: new Uint8Array(await f.arrayBuffer()) })
    }
    // 用户直接关掉对话框时，change 不触发；用 window 焦点回退判定不可靠，
    // 这里保持 Pending 由调用方取消即可（React 卸载不影响 resolve 幂等）
    el.click()
  })
}

/** 产物另存为；浏览器环境降级为 <a download> 直接下载 */
export async function saveDocAs(
  defaultName: string,
  data: Uint8Array,
  kind: OfficeKind = 'docx',
): Promise<string | null> {
  const api = electronApi()
  if (api?.docSaveAs) {
    /* ★ 2026-10-12 用户指令：另存为记住上次目录——下次对话框默认定位到那里。
     *   技巧：defaultPath 本身接受完整路径，把「记住的目录+文件名」拼进 defaultName 即可，无需改主进程。 */
    let lastDir = ''
    try { lastDir = localStorage.getItem('oo.lastSaveDir') || '' } catch { /* 无 localStorage 则忽略 */ }
    const withDir = lastDir ? lastDir.replace(/[\\/]+$/, '') + '\\' + defaultName : defaultName
    const saved = await api.docSaveAs(withDir, data, kind)
    if (saved) {
      try { localStorage.setItem('oo.lastSaveDir', saved.replace(/[\/][^\/]*$/, '')) } catch { /* 忽略 */ }
    }
    return saved
  }

  const blob = new Blob([data as BlobPart], { type: WEB_KINDS[kind].mime })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = defaultName
  a.click()
  setTimeout(() => URL.revokeObjectURL(url), 10_000)
  return defaultName
}

/* ---------- ★ 批次17：磁盘模板自动索引（工作目录 → 模板库）---------- */

/**
 * 为什么要这层：
 *   节点绑定的是模板库 id，而模板库（浏览器侧 IndexedDB / Electron 侧工作目录）两套存储各活各的。
 *   磁盘 `server-data/office/` 里的 .xlsx/.docx 从不自动进索引 —— 换浏览器、清缓存、换机器之后，
 *   节点 config 里的 id 就成了「指向不存在的东西」，打开只剩一句
 *   「模板已不在模板库中」，而文件其实一直躺在磁盘上。这是批次 17 实测抓到的断点。
 *
 * 端点选择（踩过的坑，别改）：
 *   dev 模式下 **必须用 `/office-samples`**（scripts/oo-core.mjs 提供）。
 *   `/office/files` 看着更"正规"，但在 dev 下永远到不了 app-server ——
 *   oo-core 的静态分发会先吞掉未匹配路径并 return，表现为返回 index.html 的开头两字符（"nf"）。
 *   `/office/files` 只在独立 app-server 部署（electron 生产 / vite.web.config）下有效，故保留。
 *
 * 口径（刻意保守，宁可不建也不重建）：
 *   ① 只补**缺失**的条目 —— 索引里已有的同名项一律跳过，绝不覆盖用户在编辑器里改过的模板；
 *   ② 按 **name** 去重，重复调用幂等，不产生重复条目；
 *   ③ 服务端不可用 / 纯静态托管 → 静默返回 0，绝不阻塞正常打开流程；
 *   ④ 逐文件独立 try —— 一个坏文件不让整批种子失败。
 */

/** 磁盘清单（与 oo-core.mjs 的 GET /office-samples 契约对齐） */
export interface DiskTemplateFile {
  name: string
  ext: string
  size: number
  updatedAt: number
}

/** 拉磁盘清单；任何失败都当作「没有磁盘模板」返回空数组 */
export async function listDiskTemplates(base = ''): Promise<DiskTemplateFile[]> {
  try {
    const r = await fetch(base + '/office-samples', { headers: { accept: 'application/json' } })
    if (!r.ok) return []
    const j = (await r.json()) as { files?: DiskTemplateFile[] }
    return Array.isArray(j.files) ? j.files : []
  } catch {
    return []
  }
}

/** 下拉磁盘模板字节（dev 与生产两套端点，前者为主，后者兜底） */
async function fetchDiskBytes(base: string, name: string): Promise<Uint8Array | null> {
  const enc = encodeURIComponent(name)
  for (const path of ['/office-samples/', '/office/files/']) {
    try {
      const r = await fetch(base + path + enc)
      if (!r.ok) continue
      const b = new Uint8Array(await r.arrayBuffer())
      if (b.byteLength) return b
    } catch {
      /* 换下一个端点 */
    }
  }
  return null
}

/**
 * 一次性把磁盘模板补进模板库，返回新补进去的条数。
 * 幂等：按 name 去重，重复调用不会重复入库。
 */
export async function seedTemplatesFromDisk(base = ''): Promise<number> {
  const disk = await listDiskTemplates(base)
  if (!disk.length) return 0
  const store = getTemplateStore()
  const have = new Set((await store.list().catch(() => [])).map((t) => t.name))
  let added = 0
  for (const f of disk) {
    if (have.has(f.name)) continue
    const bytes = await fetchDiskBytes(base, f.name)
    if (!bytes || !bytes.byteLength) continue
    try {
      await store.save(f.name, bytes)
      added++
    } catch {
      /* 单个失败不影响整批 */
    }
  }
  return added
}

/**
 * 导出 PDF。
 *
 * Electron：主进程用隐藏窗口 printToPDF，静默落盘，返回真实路径。
 * 浏览器：没有原生 PDF 生成器，降级为隐藏 iframe + 系统打印对话框，
 * 用户在打印面板里选「另存为 PDF」。排版与桌面端完全一致（同一份 HTML）。
 */
export async function savePdfAs(defaultName: string, html: string): Promise<string | null> {
  const api = electronApi()
  if (api?.docPdf) return api.docPdf(html, defaultName)

  const ifr = document.createElement('iframe')
  ifr.setAttribute('aria-hidden', 'true')
  ifr.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden;'
  ifr.srcdoc = html
  document.body.appendChild(ifr)

  // srcdoc 的 load 事件在部分浏览器不触发，加超时兜底
  await new Promise<void>((resolve) => {
    let done = false
    const go = () => {
      if (done) return
      done = true
      resolve()
    }
    ifr.onload = go
    setTimeout(go, 1500)
  })

  ifr.contentWindow?.focus()
  ifr.contentWindow?.print()
  // 打印对话框是同步阻塞的，但用户可能直接关掉；留足时间再清理
  setTimeout(() => ifr.remove(), 60_000)
  return defaultName
}

/* ============================================================
   打开失败的样本记录
   ------------------------------------------------------------
   为什么值得做：归一化层（src/xlsx/normalize.ts）是对第三方库严格性的补偿，
   而「还需要补什么」只有真实失败样本能回答。用户当场只看到一句提示，
   但文件副本 + 错误原文留在本地，下次就能照着扩充补丁。

   为什么放在这一层：src/xlsx/ 是纯计算层，不许碰 IO。
   落盘是平台适配的活，所以归这里。
   ============================================================ */

const FAIL_KEY = 'client-base.openFailures'
/** 浏览器降级路径只留最近 N 条（localStorage 有 5MB 上限，且这里不存字节） */
const FAIL_KEEP = 20

export interface OpenFailureRecord extends OpenFailureSample {
  /** 原始字节落在哪（浏览器降级路径为 null —— 那里不留字节） */
  bytesPath: string | null
}

/** 文件名的安全化：去掉路径分隔符与 Windows 非法字符，避免写到目录外 */
function safeFileName(name: string): string {
  const base = name.replace(/[\\/]/g, '_').replace(/[<>:"|?*\u0000-\u001f]/g, '_').trim()
  return (base || 'unnamed').slice(0, 120)
}

/**
 * 字节指纹。优先 WebCrypto 的 SHA-256；不可用时（非安全上下文）降级为 FNV-1a。
 * 这个值只用于「同一份文件是否重复打不开」的去重标识，不用于安全用途，
 * 所以降级是可接受的 —— 但要在结果里标出来，别让人误以为是真 SHA。
 */
export async function hashBytes(bytes: Uint8Array): Promise<string> {
  const subtle = (globalThis.crypto as Crypto | undefined)?.subtle
  if (subtle) {
    try {
      const buf = await subtle.digest('SHA-256', bytes as unknown as ArrayBuffer)
      return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('')
    } catch {
      /* 落到下面的降级 */
    }
  }
  let h = 0x811c9dc5
  for (let i = 0; i < bytes.length; i++) {
    h ^= bytes[i]
    h = Math.imul(h, 0x01000193) >>> 0
  }
  return `fnv1a-${h.toString(16).padStart(8, '0')}-len${bytes.length}`
}

/**
 * 记录一次「打不开」。返回一句人话，直接显示给用户（告诉样本去哪了）。
 * 任何内部错误都被吞掉 —— 记录样本失败绝不能再抛错盖住真正的问题。
 */
export async function reportOpenFailure(args: {
  name: string
  bytes: Uint8Array | null
  failure: OpenFailure
}): Promise<{ hint: string; record: OpenFailureRecord }> {
  const { name, bytes, failure } = args
  const sha256 = bytes && bytes.length ? await hashBytes(bytes) : '(无字节)'
  const base: OpenFailureSample = {
    at: Date.now(),
    name: safeFileName(name),
    size: bytes?.byteLength ?? 0,
    sha256,
    kind: failure.kind,
    title: failure.title,
    detail: failure.detail,
    applied: failure.applied,
  }

  const api = electronApi()
  if (api?.failSave) {
    try {
      const where = await api.failSave(base, bytes)
      return { hint: `已记录失败样本：${where}`, record: { ...base, bytesPath: where } }
    } catch {
      /* 落到 localStorage 兜底 */
    }
  }

  try {
    const list = JSON.parse(localStorage.getItem(FAIL_KEY) ?? '[]') as OpenFailureRecord[]
    const record: OpenFailureRecord = { ...base, bytesPath: null }
    list.unshift(record)
    localStorage.setItem(FAIL_KEY, JSON.stringify(list.slice(0, FAIL_KEEP)))
    return { hint: `已记录失败样本（本地，最近 ${FAIL_KEEP} 条）`, record }
  } catch {
    return { hint: '样本记录失败（不影响使用）', record: { ...base, bytesPath: null } }
  }
}

/** 列出已记录的失败样本（诊断面板 / 交给开发排查用） */
export async function listOpenFailures(): Promise<OpenFailureRecord[]> {
  const api = electronApi()
  if (api?.failList) {
    try {
      return (await api.failList()) as OpenFailureRecord[]
    } catch {
      return []
    }
  }
  try {
    return JSON.parse(localStorage.getItem(FAIL_KEY) ?? '[]') as OpenFailureRecord[]
  } catch {
    return []
  }
}
