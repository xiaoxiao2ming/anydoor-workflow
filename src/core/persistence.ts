// 持久化适配层：sql.js 导出的 .sqlite 字节存到哪里
// - Electron（npm run dev）：经 preload IPC 写到磁盘 userData/client-base.sqlite
// - 浏览器（npm run dev:web）：落到 IndexedDB，刷新不丢（等价于"磁盘"）
export interface PersistenceAdapter {
  load(): Promise<Uint8Array | null>
  save(data: Uint8Array): Promise<void>
}

const IDB_NAME = 'client-base'
const IDB_STORE = 'kv'
const IDB_KEY = 'sqlite'

function openIdb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, 1)
    req.onupgradeneeded = () => req.result.createObjectStore(IDB_STORE)
    req.onsuccess = () => resolve(req.result)
    req.onerror = () => reject(req.error)
  })
}

const indexedDbAdapter: PersistenceAdapter = {
  async load() {
    const idb = await openIdb()
    return new Promise((resolve, reject) => {
      const tx = idb.transaction(IDB_STORE, 'readonly')
      const req = tx.objectStore(IDB_STORE).get(IDB_KEY)
      req.onsuccess = () => resolve((req.result as Uint8Array | undefined) ?? null)
      req.onerror = () => reject(req.error)
    })
  },
  async save(data) {
    const idb = await openIdb()
    await new Promise<void>((resolve, reject) => {
      const tx = idb.transaction(IDB_STORE, 'readwrite')
      tx.objectStore(IDB_STORE).put(data, IDB_KEY)
      tx.oncomplete = () => resolve()
      tx.onerror = () => reject(tx.error)
    })
  },
}

interface ElectronApi {
  readDb(): Promise<Uint8Array | null>
  writeDb(data: Uint8Array): Promise<void>
}

function electronApi(): ElectronApi | undefined {
  return (window as unknown as { api?: ElectronApi }).api
}

const electronAdapter: PersistenceAdapter = {
  async load() {
    return (await electronApi()?.readDb?.()) ?? null
  },
  async save(data) {
    await electronApi()?.writeDb?.(data)
  },
}

/** 自动选：有 window.api（Electron）就走磁盘文件，否则走 IndexedDB */
export function getAdapter(): PersistenceAdapter {
  const api = electronApi()
  if (api && typeof api.readDb === 'function' && typeof api.writeDb === 'function') {
    return electronAdapter
  }
  return indexedDbAdapter
}
