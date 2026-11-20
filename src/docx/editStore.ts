/**
 * Word 就地编辑稿的持久化（按节点 id）。
 *
 * 关键决策：编辑稿 ≠ 模板。
 *   就地编辑的对象是「绑定产物」（bind stripControls 后的字节），里面已经没有
 *   sdt 控件——如果把它覆盖回模板库，parseTemplate 会找不到任何绑定控件，
 *   数据连接就断了。所以编辑稿单独存放，模板字节永不被编辑覆盖。
 *
 * 与 templateStore 同库（client-base）不同 store（docx-edit）。
 * 注意版本号必须与 templateStore 严格一致（v3）—— 一边高一边低，
 * 低版本的 open 抛 VersionError、高版本的 upgrade 被对方未关的连接 blocked。
 * 两边 upgrade 都幂等补建对方/自己的 store，保证谁先打开都能正常 upgrade。
 *
 * 生命周期（WordReportPanel 持有）：
 *   打开编辑器 → 有稿则恢复（提示「编辑稿模式」）
 *   数据/模板重新绑定 → 稿作废清除（以数据为准，与 OnlyOffice「在线编辑」同语义）
 *   编辑中 → 防抖写入；关闭编辑器 → flush
 */

const DB_NAME = 'client-base'
const DB_VERSION = 3
const STORE = 'docx-edit'

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION)
    req.onupgradeneeded = () => {
      const db = req.result
      // 幂等补建：templateStore(v2) 的两个 store 可能先建也可能后建
      if (!db.objectStoreNames.contains('tpl')) db.createObjectStore('tpl')
      if (!db.objectStoreNames.contains('tpl-index')) db.createObjectStore('tpl-index')
      if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE)
    }
    req.onsuccess = () => resolve(req.result)
    req.onerror = () => reject(req.error)
  })
}

function withStore<T>(mode: IDBTransactionMode, fn: (os: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  return openDb().then(
    (db) =>
      new Promise<T>((resolve, reject) => {
        const tx = db.transaction(STORE, mode)
        const req = fn(tx.objectStore(STORE))
        req.onsuccess = () => resolve(req.result)
        req.onerror = () => reject(req.error)
      }),
  )
}

/** 稿的落库形态：带 tplId —— 换模板后旧稿自动失效，不会张冠李戴恢复到新模板上 */
export interface EditedDocx {
  tplId: string
  bytes: Uint8Array
}

export async function loadEditedDocx(nodeId: string): Promise<EditedDocx | null> {
  try {
    return (await withStore<EditedDocx | undefined>('readonly', (os) => os.get(nodeId))) ?? null
  } catch {
    return null
  }
}

export async function saveEditedDocx(nodeId: string, tplId: string, bytes: Uint8Array): Promise<void> {
  try {
    await withStore('readwrite', (os) => os.put({ tplId, bytes }, nodeId))
  } catch {
    /* 存不了就算了：编辑稿本就是尽力持久化，失败不影响当前会话 */
  }
}

export async function clearEditedDocx(nodeId: string): Promise<void> {
  try {
    await withStore('readwrite', (os) => os.delete(nodeId))
  } catch {
    /* 同上 */
  }
}
