/**
 * 画布大数据 blob 存储（2026-10-24 性能根治）：
 * PDF 源文件/OO 文档的 base64 从节点 config 移出——config 瘦身到 KB 级，
 * 拖停保存的全量序列化不再阻塞主线程。
 * 旧数据兼容：读侧 config.b64 优先回退（渐进迁移，无需一次性搬迁）。
 */
const DB_NAME = 'clientbase-blobs'
const STORE = 'blobs'
let dbp: Promise<IDBDatabase> | null = null

function openDb(): Promise<IDBDatabase> {
  if (!dbp) {
    dbp = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, 1)
      req.onupgradeneeded = () => {
        if (!req.result.objectStoreNames.contains(STORE)) req.result.createObjectStore(STORE)
      }
      req.onsuccess = () => resolve(req.result)
      req.onerror = () => reject(req.error)
    })
  }
  return dbp
}

export async function putBlob(id: string, bytes: Uint8Array): Promise<void> {
  const db = await openDb()
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite')
    tx.objectStore(STORE).put(bytes, id)
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error)
  })
}

export async function getBlob(id: string): Promise<Uint8Array | null> {
  const db = await openDb()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readonly')
    const req = tx.objectStore(STORE).get(id)
    req.onsuccess = () => resolve((req.result as Uint8Array) ?? null)
    req.onerror = () => reject(req.error)
  })
}

/* ★ 2026-10-24 照片内容寻址（用户指令「同样的文件不改就做一个 不要重复保存」）：
   SHA-1 内容哈希作 key——同一文件（无论拖进多少个节点）算出同一指纹，
   天然只存一份；文件没改 → 指纹不变 → 不产生重复记录。 */
export async function hashBytes (bytes: Uint8Array | ArrayBuffer): Promise<string> {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const digest = await crypto.subtle.digest('SHA-1', u8.slice().buffer as ArrayBuffer)
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('')
}

/** 内容寻址写入：同内容幂等（key 相同，覆盖同数据无害）。返回存储 key。 */
export async function putContentAddressed (prefix: string, bytes: Uint8Array | ArrayBuffer): Promise<string> {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const hash = await hashBytes(u8)
  const key = `${prefix}-${hash}`
  await putBlob(key, u8)
  return key
}

/* ★ 2026-10-24 缩略图分级：原图（4K）纹理在画布缩放时 GPU 重采样是拖动卡主因。
   导入时下采样 256px 缩略图入库——网格显示用缩略图（小纹理零压力），预览/合并用原图。 */
export async function makeThumbnail (buf: Uint8Array | ArrayBuffer): Promise<Blob> {
  const blob = new Blob([buf instanceof Uint8Array ? buf.slice().buffer as ArrayBuffer : buf], { type: 'image/*' })
  const bmp = await createImageBitmap(blob)
  const scale = Math.min(1, 256 / Math.max(bmp.width, bmp.height))
  const w = Math.max(1, Math.round(bmp.width * scale))
  const h = Math.max(1, Math.round(bmp.height * scale))
  const cv = document.createElement('canvas')
  cv.width = w; cv.height = h
  const ctx = cv.getContext('2d')!
  ctx.drawImage(bmp, 0, 0, w, h)
  bmp.close()
  return await new Promise<Blob>((res) => cv.toBlob((b) => res(b ?? blob), 'image/jpeg', 0.8))
}

/** 是否已存在同指纹资产（导入时查重用——命中即跳过写入，只记录引用） */
export async function hasBlob (key: string): Promise<boolean> {
  return (await getBlob(key)) !== null
}

/** 遍历库内全部条目（导出资产查重用）——返回 { name(=key+原始大小), bytesB64 } */
export async function listAllBlobs (): Promise<{ name: string; bytesB64: string }[]> {
  const db = await openDb()
  return new Promise((resolve, reject) => {
    const out: { name: string; bytesB64: string }[] = []
    const tx = db.transaction(STORE, 'readonly')
    const req = tx.objectStore(STORE).openCursor()
    req.onsuccess = () => {
      const c = req.result
      if (!c) { db.close(); return resolve(out) }
      const val = c.value as Uint8Array
      const kb = (val.length / 1024).toFixed(1) + 'KB'
      out.push({ name: c.key as string + ' [' + kb + '].bytes', bytesB64: Buffer_polyfill(val) })
      c.continue()
    }
    req.onerror = () => reject(req.error)
  })
}
function Buffer_polyfill (u8: Uint8Array): string {
  let bin = ''
  const CH = 0x8000
  for (let i = 0; i < u8.length; i += CH) bin += String.fromCharCode(...u8.subarray(i, i + CH))
  return btoa(bin)
}
