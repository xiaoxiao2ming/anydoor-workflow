/**
 * .wfc 工程单文件保存（2026-10-06 用户：「点击保存能保存当前画布所有内容为 wfc」）。
 *
 * ── .wfc 是什么 ──
 * .wfc（WorkFlow Canvas）= 已拍板的 .cbproj 目录式工程（docs/18，Q1=B 目录式）
 * 的**无损 zip 单文件**，与 .cbproj.zip 仅扩展名不同 —— zip 内部目录树逐字节一致
 * （projectPack.packProject 产物：project.json + edges.json + nodes/ + assets/ + …）。
 * ★ 刻意不造第二套格式：目录式是单一真相源，.wfc 只是「整个目录带走」的交付外壳；
 *   将来的「打开 .wfc」= unpackProject 还原目录 → openProjectFull，同一条链路，
 *   不存在「wfc 一套、cbproj 一套」的两头维护。
 *
 * ── 「当前画布所有内容」的口径 ──
 * tables + records + positions（全部画布页、含回收站节点——deleted 标记随 schema 走）
 * + 控件状态（localStorage → viewState）+ 节点资产字节（模板/就地编辑稿
 * → assets/）。撤销历史不在口径内（阶段二，工程格式 history/ 目录已预留）。
 *
 * ── 分层位置 ──
 * 本文件在 src/ 根而非 core/ —— 它碰 document/URL（触发浏览器下载），
 * 属「UI 侧胶水」，与 CanvasRibbon 同层；core 零依赖它的任何东西。
 *
 * ── 落盘路径 ──
 * ① saveProjectFull：运行时快照 → IDB 暂存工程区（固定 rootDir，每次整体覆盖）
 * ② packProject：IDB 工程目录 → zip 字节
 * ③ Blob + a[download]：触发下载（Electron 的 webview 同样走下载通道弹保存框，
 *    浏览器/Electron 行为一致，无需平台分支）
 */
import { useBaseStore } from './core/store'
import { saveProjectFull } from './core/projectIO'
import { createIdbProjectFs, type ProjectRef } from './core/projectStore'
import { packProject } from './core/projectPack'
import { browserKV } from './core/viewState'
import type { AssetRuntime } from './core/nodeAssets'
import { getTemplateStore } from './docx/templateStore'
import { loadEditedDocx, saveEditedDocx } from './docx/editStore'

/** .wfc 单文件固定扩展名（内部 = .cbproj 目录树的 zip，见文件头说明） */
export const WFC_EXT = '.wfc'

/** IDB 里的导出暂存工程根（每次保存整体覆盖写，不累积历史） */
export const WFC_WORKSPACE_ROOT = 'wfc-export'

/**
 * templateStore + editStore → AssetRuntime 适配。
 *
 * 为什么需要这层：nodeAssets.AssetRuntime 是接口（core 不许 import 浏览器存储），
 * 真实运行时在 docx/templateStore（模板字节，IDB）与 docx/editStore（就地编辑稿）。
 * 之前 saveProjectFull 只有测试在调（mock runtime）——UI 从未接过真存储，本函数补上。
 */
export function makeBrowserAssetRuntime(): AssetRuntime {
  const ts = () => getTemplateStore()
  return {
    readTemplate: (id) => ts().read(id),
    templateMeta: async (id) => (await ts().list()).find((m) => m.id === id) ?? null,
    readEdited: async (nodeId) => {
      const e = await loadEditedDocx(nodeId)
      return e ? { tplId: e.tplId, bytes: e.bytes } : null
    },
    ensureTemplate: (name, bytes) => ts().save(name, bytes).then((m) => m.id),
    writeEdited: (nodeId, tplId, bytes) => saveEditedDocx(nodeId, tplId, bytes),
    /* ★ 2026-10-25 attachment 回灌：词条/产物字节按原 key 灌回 IDB pdfBlobStore */
    restoreBlob: async (key, bytes) => {
      const { putBlob } = await import('./core/pdfBlobStore')
      await putBlob(key, bytes)
    },
  }
}

/** 保存结果：给 UI 弹提示用（含资产问题的完整清单） */
export interface WfcSaveResult {
  fileName: string
  /** 资产内容去重：重复字节跳写数量（同文件多节点引用只存一份） */
  dedupInfo?: { count: number }
  /** .wfc 文件字节数（展示用 KB 换算由调用方做） */
  bytes: number
  /** 资产层面问题（模板丢失/路径非法等）——不阻断保存，但必须让用户看到 */
  problems: string[]
}

/**
 * 保存当前画布全部内容为 .wfc 单文件（触发浏览器下载）。
 *
 * @param pageName 当前画布页名（文件名主体；非法文件名字符替换为 _）
 * @throws 存储失败/写资产失败时原样抛出（调用方弹错误提示）
 */
export type WfcSaveStep = { label: string; pct: number }
/** 可选进度回调：预查编辑稿逐节点 + 各阶段（2026-10-24 用户指令「保存进度条 每一步展示」） */
export async function saveCanvasAsWfc(
    pageName: string,
    onStep?: (st: WfcSaveStep) => void,
    opts?: { includeHistory?: boolean | Set<string>; writeTo?: string; excludeNodeIds?: string[] },
  ): Promise<WfcSaveResult> {
  /* ★ 2026-10-25 保存前强制 b64 剥离（与 saveNow 同款兜底）——手动导出路径也要走：
     任何残留 b64（pdfItems/pdfHistory/ooBytes64/photoItems）先入 IndexedDB 再剥离，
     保证 .wfc 产物轻量。幂等（已剥离项跳过）。 */
  onStep?.({ label: '清理历史残留数据…', pct: 1 })
  try {
    const { putBlob, putContentAddressed } = await import('./core/pdfBlobStore')
    const b64ToU8 = (b64: string) => {
      const bin = atob(b64)
      const u8 = new Uint8Array(bin.length)
      for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
      return u8
    }
    const cur0 = useBaseStore.getState()
    for (const t of Object.values(cur0.tables)) {
      const cfg0 = (t?.config ?? {}) as {
        pdfItems?: { id: string; name?: string; b64?: string; inDb?: boolean }[]
        pdfHistory?: { id?: string; at?: number; b64?: string; inDb?: boolean }[]
        ooBytes64?: string
        ooBytesDbId?: string
        photoItems?: { id: string; name: string; type: string; b64?: string; photoDbId?: string }[]
      }
      let touched = false
      const patchCfg: Record<string, unknown> = {}
      if (cfg0.pdfItems?.some((it) => it.b64 && !it.inDb)) {
        patchCfg.pdfItems = await Promise.all(cfg0.pdfItems.map(async (it) => {
          if (!it.b64 || it.inDb) return it
          await putBlob('pdfitem-' + it.id, b64ToU8(it.b64))
          touched = true
          return { ...it, b64: undefined, inDb: true }
        }))
      }
      if (cfg0.pdfHistory?.some((h) => h.b64 && !h.inDb)) {
        patchCfg.pdfHistory = await Promise.all(cfg0.pdfHistory.map(async (h) => {
          if (!h.b64 || h.inDb) return h
          await putBlob('pdfhist-' + (h.id ?? 'x') + '-' + (h.at ?? 0), b64ToU8(h.b64))
          touched = true
          return { ...h, b64: undefined, inDb: true }
        }))
      }
      if (cfg0.ooBytes64 && !cfg0.ooBytesDbId) {
        await putBlob('oobytes-' + t.id, b64ToU8(cfg0.ooBytes64))
        patchCfg.ooBytesDbId = 'oobytes-' + t.id
        patchCfg.ooBytes64 = undefined
        touched = true
      }
      if (cfg0.photoItems?.some((p) => p.b64 && !p.photoDbId)) {
        patchCfg.photoItems = await Promise.all(cfg0.photoItems.map(async (p) => {
          if (!p.b64 || p.photoDbId) return p
          p.photoDbId = await putContentAddressed('photo', b64ToU8(p.b64))
          touched = true
          return { ...p, b64: undefined }
        }))
      }
      if (touched) {
        const curT = useBaseStore.getState().tables[t.id]
        useBaseStore.setState({ ...useBaseStore.getState(), tables: { ...useBaseStore.getState().tables, [t.id]: { ...curT, config: { ...(curT?.config ?? {}), ...patchCfg } } } })
        console.log('[导出前迁移]', t.id, '已剥离 b64 → IndexedDB')
      }
    }
  } catch (eMig) {
    console.error('[导出前迁移] 失败（保留原数据继续导出）:', eMig)
  }
  onStep?.({ label: '残留清理完成', pct: 2 })
  const sLive0 = useBaseStore.getState()
  /* ★ 2026-10-25 用户令「孤立节点可以选择不保存」：本次保存剔除勾掉的孤立节点
     （只影响 .wfc 内容，画布不受影响） */
  const excl = new Set(opts?.excludeNodeIds ?? [])
  const sLive = excl.size
    ? {
        ...sLive0,
        tables: Object.fromEntries(Object.entries(sLive0.tables).filter(([id]) => !excl.has(id))),
        records: Object.fromEntries(Object.entries(sLive0.records).filter(([id]) => !excl.has(id))),
        positions: Object.fromEntries(Object.entries(sLive0.positions).filter(([id]) => !excl.has(id))),
      } as typeof sLive0
    : sLive0
  /* ★ 2026-10-24 用户指令「历史记录文件 保存时勾选是否存入」：默认不勾——
     保存副本剥离 pdfHistory（字节留在 IndexedDB，画布不受影响），勾选才随 .wfc 打包 */
  const s = opts?.includeHistory === true || opts?.includeHistory instanceof Set
    ? sLive
    : {
        ...sLive,
        tables: Object.fromEntries(Object.entries(sLive.tables).map(([id, t]) => {
          const cfg = t?.config as { pdfHistory?: unknown[] } | undefined
          return cfg?.pdfHistory?.length
            ? [id, { ...t, config: { ...cfg, pdfHistory: [] } }]
            : [id, t]
        })),
      } as typeof sLive
  const fs = createIdbProjectFs()
  const ref: ProjectRef = { rootDir: WFC_WORKSPACE_ROOT }

  // ⓪ 预查询「哪些节点有就地编辑稿」：planAllNodeAssets 的 hasEdited 是**同步**签名
  //   （不传视为 false → 稿会静默漏出工程），而判定只能查 editStore（异步 IDB）
  //   —— 只能先查好存集合，保存时同步判定。miss 的 get 很快，逐个查可接受。
  const editedIds = new Set<string>()
  const allIds = Object.keys(s.tables)
  let editedI = 0
  onStep?.({ label: `预查编辑稿（0/${allIds.length}）`, pct: 3 })
  for (const id of allIds) {
    if (await loadEditedDocx(id)) editedIds.add(id)
    editedI++
    if (editedI % 3 === 0 || editedI === allIds.length) onStep?.({ label: `预查编辑稿（${editedI}/${allIds.length}）`, pct: 3 + Math.round((editedI / Math.max(1, allIds.length)) * 17) })
  }
  onStep?.({ label: '生成运行时快照…', pct: 22 })

  // ① 运行时快照 → SaveProjectInput（name/tables/records/positions 为必填最小集；
  //    viewState/assets 由 saveProjectFull 内部采集后并入，不经这里）
  // ② 全量保存到 IDB 暂存区（先资产后清单的崩溃一致性由 saveProject 保证）
  onStep?.({ label: '写入工程数据与资产…', pct: 30 })
  const tSave0 = Date.now()
  const out = await saveProjectFull(
    fs,
    ref,
    { name: pageName, tables: s.tables, records: s.records, positions: s.positions },
    {
      runtime: makeBrowserAssetRuntime(),
      kv: browserKV(),
      includeHistory: opts?.includeHistory === true,
      // 编辑稿字节由 rt.readEdited → editStore（IDB 持久稿）提供；
      // 打开中编辑器的内存稿是否已落 editStore 取决于编辑器自身——那是它的职责
      hasEdited: (nodeId) => editedIds.has(nodeId),
    },
  )

  // ②.5 资产写入完成——展示去重结果（2026-10-24 用户指令「进度写上重复文件数 压缩 保存」）
  const dd = out.dedupInfo?.count ?? 0
  onStep?.({ label: dd > 0 ? `写入完成（重复文件 ${dd} 个已合并）…` : '写入工程数据与资产…', pct: 55 })

  // ③ 压缩打包为单文件字节（zip 外壳，内部 = .cbproj 目录树，结构由 packProject 保证）
  onStep?.({ label: '压缩打包 .wfc…', pct: 65 })
  const bytes = await packProject(fs, WFC_WORKSPACE_ROOT)
  onStep?.({ label: '压缩完成（' + (bytes.byteLength / 1024 / 1024).toFixed(1) + ' MB）', pct: 92 })

  // ④ 交付：writeTo 存在 = 字节直写选定路径（Electron 桥——用户先选的位置）；
  //    否则触发浏览器下载（web 端）。文件名主体 = 画布页名，Windows 非法字符替换为 _。
  const safeName = (pageName || '未命名画布').replace(/[\\/:*?"<>|]/g, '_')
  const fileName = `${safeName}${WFC_EXT}`
  if (opts?.writeTo) {
    onStep?.({ label: '写入选定位置…', pct: 96 })
    const api = (globalThis as { api?: { wfcWriteFile?: (p: string, b: Uint8Array) => Promise<{ ok: boolean }> } }).api
    if (!api?.wfcWriteFile) throw new Error('桌面桥不可用——无法写入选定位置')
    await api.wfcWriteFile(opts.writeTo, bytes)
  } else {
    const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
    const blob = new Blob([ab], { type: 'application/octet-stream' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = fileName
    document.body.appendChild(a)
    a.click()
    a.remove()
    setTimeout(() => URL.revokeObjectURL(url), 10_000)
  }

  onStep?.({ label: '触发保存下载…', pct: 96 })
  onStep?.({ label: '完成', pct: 100 })
  return {
    fileName,
    bytes: bytes.byteLength,
    problems: out.assetProblems.map((p) => p.message),
    dedupInfo: out.dedupInfo,
  }
}

/* ★ 2026-10-24 保存前文件清单（用户指令「保存之前把所有文件列出来 每个大小」）：
   按节点分组列出全部资产（含历史产物），重复合并引用只计一次体积。 */
export interface AssetManifestEntry { node: string; nodeId: string; name: string; sizeMB: number; dup: boolean }
export interface AssetManifest { groups: { node: string; nodeId: string; entries: AssetManifestEntry[]; totalMB: number }[]; totalMB: number; totalN: number }

export async function collectAssetManifest (): Promise<AssetManifest> {
  const s = useBaseStore.getState()
  const { getBlob } = await import('./core/pdfBlobStore')
  const seen = new Set<string>()
  const groups: AssetManifest['groups'] = []
  let totalMB = 0
  let totalN = 0
  const fmt = (bytes: number) => Math.round((bytes / 1024 / 1024) * 100) / 100
  for (const t of Object.values(s.tables)) {
    const cfg = (t?.config ?? {}) as {
      pdfItems?: { id: string; name?: string; inDb?: boolean }[]
      photoItems?: { id: string; name?: string; type?: string; thumbDbId?: string; photoDbId?: string }[]
      ooBytesDbId?: string
      ooTemplateName?: string
      pdfHistory?: { id: string; name?: string; at?: number; inDb?: boolean }[]
    }
    const entries: AssetManifestEntry[] = []
    const nodeName = t?.name ?? t?.id ?? '节点'
    // PDF 文件组
    for (const it of cfg.pdfItems ?? []) {
      const key = 'pdfitem-' + it.id
      const bytes = await getBlob(key)
      const mb = bytes ? fmt(bytes.length) : 0
      const dup = seen.has(key)
      if (!dup) { seen.add(key); totalMB += mb; totalN++ }
      entries.push({ node: nodeName, nodeId: t.id, name: it.name ?? key, sizeMB: mb, dup })
    }
    // 照片（缩略图/原图分级——清单按原图计，标注缩略）
    for (const it of cfg.photoItems ?? []) {
      const key = it.photoDbId ?? ''
      if (!key) continue
      const bytes = await getBlob(key)
      const mb = bytes ? fmt(bytes.length) : 0
      const dup = seen.has(key)
      if (!dup) { seen.add(key); totalMB += mb; totalN++ }
      entries.push({ node: nodeName, nodeId: t.id, name: (it.name ?? key) + (it.thumbDbId ? ' [已生成缩略图]' : ''), sizeMB: mb, dup })
    }
    // OO 文档
    if (cfg.ooBytesDbId) {
      const bytes = await getBlob(cfg.ooBytesDbId)
      const mb = bytes ? fmt(bytes.length) : 0
      const dup = seen.has(cfg.ooBytesDbId)
      if (!dup) { seen.add(cfg.ooBytesDbId); totalMB += mb; totalN++ }
      entries.push({ node: nodeName, nodeId: t.id, name: 'OO 文档（' + (cfg.ooTemplateName ?? cfg.ooBytesDbId) + '）', sizeMB: mb, dup })
    }
    // 历史产物（合并生成）
    for (const h of cfg.pdfHistory ?? []) {
      const key = 'pdfhist-' + h.id + '-' + h.at
      const bytes = await getBlob(key)
      const mb = bytes ? fmt(bytes.length) : 0
      const dup = seen.has(key)
      if (!dup) { seen.add(key); totalMB += mb; totalN++ }
      entries.push({ node: nodeName + '（历史产物）', nodeId: t.id, name: (h.name ?? key) + ' · ' + new Date(h.at ?? 0).toLocaleString('zh-CN', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }), sizeMB: mb, dup })
    }
    if (entries.length) groups.push({ node: nodeName, nodeId: t.id, entries, totalMB: Math.round(entries.reduce((a2, b2) => a2 + b2.sizeMB, 0) * 100) / 100 })
  }
  return { groups, totalMB, totalN }
}
