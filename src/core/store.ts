// 单一数据源 Store —— 表 schema 与记录都存在这里，视图层（画布/表单）共享这一份。
// 持久化改为「集中存储后端」：所有画布页面读写都走 /api/pages，
// 多浏览器/多设备访问同一份数据（解决不同浏览器布局不一致问题）。
// 支持多画布页面（pages）：切换/新建/改名/删除，每页独立保存 {tables,records,positions}。
import { create } from 'zustand'
import type { TableSchema, NodeConfig, FieldDef, NodeKind, NodeRuleSet } from './schema'
import { MATERIALS, BILL, START, MATERIALS_SEED, BILL_SEED } from './seed'
import { isRuleNodeKind } from './nodeStyle'
// 身份槽占用判定走 ruleSync.identityOwnerOf（单一真相源，不在store 里另写一份）。
// 无循环依赖：schema / nodeStyle 都是叶子模块，ruleSync 不反向依赖 store。
import { identityOwnerOf } from './ruleSync'
// 运行动画三态的类型（判定与迁移本体在 roundRun，这里只存瞬态值）
import type { RoundRunState } from './roundRun'
import * as api from './api'
// ★ 2026-10-17 操作审计：setCell（所有单元格改动的唯一收口）在那里记一条
import { audit } from './audit'

/** ★ 节点唯一编号（2026-10-12 用户指令）：现有 nodeNo 最大值 +1，格式 N001…（创建/复制时分配） */
function nextNoOf(tables: Record<string, TableSchema>): string {
  let max = 0
  for (const t of Object.values(tables)) {
    const m = String((t.config as { nodeNo?: string } | undefined)?.nodeNo || '').match(/^N(\d+)$/)
    if (m) max = Math.max(max, parseInt(m[1], 10))
  }
  return 'N' + String(max + 1).padStart(3, '0')
}

export type Row = Record<string, unknown>

export interface NodePos {
  x: number
  y: number
}

export interface SaveableState {
  tables: Record<string, TableSchema>
  records: Record<string, Row[]>
  positions: Record<string, NodePos>
}

// 撤销/重做快照：记录可写状态的一个时刻
type Snapshot = SaveableState

interface PageMeta {
  id: string
  name: string
}

interface BaseState extends SaveableState {
  ready: boolean
  pages: PageMeta[]
  activePageId: string | null
  init: () => Promise<void>
  // 页面管理
  switchPage: (id: string) => Promise<void>
  newPage: (name?: string) => Promise<string>
  /** 导入一个画布页（打开 .wfc 用，2026-10-07）：服务端建页 + 写内容 + 追加页签并切换 */
  importPage: (
    name: string,
    tables: Record<string, TableSchema>,
    records: Record<string, Row[]>,
    positions: Record<string, NodePos>,
  ) => Promise<string>
  renamePage: (id: string, name: string) => Promise<void>
  deletePage: (id: string) => Promise<void>
  // 表 / 字段 写操作
  addTable: (schema: TableSchema) => void
  addField: (tableId: string, field: FieldDef) => void
  updateField: (tableId: string, fieldId: string, patch: Partial<FieldDef>) => void
  removeField: (tableId: string, fieldId: string) => void
  renameTable: (tableId: string, name: string) => void
  duplicateTable: (tableId: string) => string
  removeTable: (tableId: string) => void
  // 回收站：软删除（移入回收站，可恢复）/ 恢复 / 彻底删除 / 清空
  trashTable: (tableId: string) => void
  trashTables: (tableIds: string[]) => void
  restoreTable: (tableId: string) => void
  restoreLastTrash: () => number
  purgeTable: (tableId: string) => void
  emptyTrash: () => void
  trashStack: string[][]
  setNodePosition: (tableId: string, pos: NodePos) => void
  setNodePositions: (updates: { id: string; pos: NodePos }[]) => void
  setNodeConfig: (tableId: string, patch: Partial<NodeConfig>) => void
  /* ── 规则锁定（圆形规则节点，2026-10-05）─────────────────────────────
     只写规则节点一份，辅节点零写入 —— 同步在读取时完成（ruleSync.ts）。
     故这几个动作的入参都**不含被同步的辅节点列表**：没有列表，就没有循环写。 */
  /** 改某 kind 的规则；全部被锁辅节点立即生效 */
  setRule: (ruleNodeId: string, kind: string, patch: Partial<NodeRuleSet>) => void
  /** 清掉某 kind 的规则，辅节点回落自身 config */
  clearRule: (ruleNodeId: string, kind: string) => void
  /** 从规则节点断开：目标节点脱离锁定，回落自身 config（属性面板「解除锁定」按钮走这里） */
  unlinkFromRuleNode: (ruleNodeId: string, targetId: string) => void
  /** 设主节点（用户：第一个默认是可修改的主节点；双击面板可改主次） */
  setRulePrimary: (ruleNodeId: string, targetId: string) => void
  /** 切换跟随 / 不跟随（用户：双击能改「跟随不跟随」） */
  setRuleFollow: (ruleNodeId: string, targetId: string, follow: boolean) => void
  /* 锁定（连上）刻意**不提供** store 动作：连线统一由 App.handleConnect 处理，
     判定在 core/ruleSync.ruleConnectPatch，写入走 setNodeConfig。
     曾有个对称的 linkToRuleNode，实测 UI 一处都没调 ——
     留着就是第二套改 linkedInputs 的入口，哪天被误用就出现绕过锁定的写入路径。 */
  /** OO 节点类型切换（按导入文件扩展名：docx→ooword / xlsx→oosheet） */
  setTableKind: (tableId: string, kind: NodeKind) => void
  setTableData: (tableId: string, fields: FieldDef[], rows: Row[]) => void
  addRow: (tableId: string, row?: Row) => void
  removeRow: (tableId: string, rowIdx: number) => void
  setCell: (tableId: string, rowIdx: number, fieldId: string, value: unknown) => void
  getRows: (tableId: string) => Row[]
  // 撤销 / 重做历史栈
  past: Snapshot[]
  future: Snapshot[]
  pushHistory: () => void
  undo: () => void
  redo: () => void
  // 切换页面前，把当前页立即落盘
  saveNow: () => Promise<void>
  /**
   * 瞬态（不持久化）· fx/FCE 悬停胶囊 → 画布对应来源节点边框亮显。
   * 备忘 2026-09-27 ⑤：编辑器窗体与画布跨组件通信的唯一通道（zustand 同 root），
   * null = 无悬停。仅供 Canvas TableNode 读，不进任何快照/落盘。
   */
  fxHoverNode: string | null
  setFxHoverNode: (id: string | null) => void
  /**
   * 瞬态（不持久化）· 当前被选中的规则节点 id（2026-10-05 用户：
   * 「点击规则跟随 与该节点跟随 节点连接的节点 的锚点 在点击选中该节点的时候没显示」）。
   * Canvas 的 onSelectionChange 写，TableNode 读：成员节点据此给自己的规则口
   * （上1/上2/下1/下2）挂 .rf-rule-ports-show 显形。null = 没有规则节点被选中。
   * ★ 为什么走 store 而不是 setNodes 塞 data/className：
   *   选中变化若驱动 nodes 重建，位置/测量字段的保留逻辑全要过一遍手，
   *   且会连带边表重算 —— fxHoverNode 已趟平的「store 瞬态桥」是零副作用通道。
   */
  selRuleNode: string | null
  setSelRuleNode: (id: string | null) => void
  /**
   * 瞬态（不持久化）· 常规节点全局折叠态（2026-10-06 用户：
   * 「做一种节点样式。每个常规节点都折叠成正方形。点击菜单的一个按钮 都变成折叠状态」）。
   * CanvasRibbon 的按钮写，TableNode 读：常规节点挂 .rf-node--folded 收成 64px 正方形。
   * false = 全部展开（默认）。规则节点不受影响（它已是圆形特殊形态）。
   * ★ 为什么走 store 而不是 Canvas state：TableNode 要读它，
   *   nodeTypes 是模块级常量拿不到 prop —— fxHoverNode/selRuleNode 同一条先例。
   */
  nodesFolded: boolean
  setNodesFolded: (v: boolean) => void
  /** ★ 连接关系快照（Canvas edges 同步）——孤立节点计算用 */
  edgeLinks: string[]
  setEdgeLinks: (ids: string[]) => void
  /** ★ PDF 合并进度（nodeId → pct 0-100；undefined=无进度） */
  pdfmergeProgress: Record<string, number>
  setPdfmergeProgress: (nodeId: string, pct: number | null) => void

  /**
   * 瞬态（不持久化）· 圆形节点（AI/参数库）运行动画状态（2026-10-06 用户：
   * 1「对选中的节点点击动态特效才会有效果」
   *  2「增加特效的结束的退出效果 测试的时候3秒结束」）。
   *
   * 三态而非布尔：'off' | 'running' | 'exiting'。
   * ★ 为什么必须有 exiting —— class 一摘，::before/::after 元素立刻被移除，
   *   动画无从播完，用户看到的是「啪」一下消失。
   *   要有退出效果（淡出+收缩），就得让元素在退出期间**留在 DOM 上**。
   * 判定与迁移本体在 core/roundRun（纯函数，有单测钉死），
   * 这里只存瞬态值 —— 与 fxHoverNode / selRuleNode / nodesFolded 同一套桥。
   */
  roundRunState: RoundRunState
  setRoundRunState: (s: RoundRunState) => void
}

let saveTimer: ReturnType<typeof setTimeout> | null = null

const MAX_HISTORY = 50

// 深拷贝快照（structuredClone 优先，兼容不支持的环境）
function clone<T>(v: T): T {
  try {
    return structuredClone(v)
  } catch {
    return JSON.parse(JSON.stringify(v))
  }
}

// 软删除单个表：捕获其它表指向它的引用并剥离（避免连线悬空），标记 deleted + 记录 trashMeta
function softDeleteTable(tables: Record<string, TableSchema>, tableId: string): Record<string, TableSchema> {
  const target = tables[tableId]
  if (!target || target.deleted) return tables
  const refs: Record<string, { inputs: string[]; linkedInputs: string[]; lookups: FieldDef[] }> = {}
  const next = { ...tables }
  for (const t of Object.values(next)) {
    if (t.id === tableId) continue
    const ins = t.config?.inputs ?? []
    const lk = t.config?.linkedInputs ?? []
    const lkFields = t.fields.filter((f) => f.type === 'lookup' && f.lookup?.table === tableId)
    if (ins.includes(tableId) || lk.includes(tableId) || lkFields.length) {
      refs[t.id] = {
        inputs: ins.includes(tableId) ? [tableId] : [],
        linkedInputs: lk.includes(tableId) ? [tableId] : [],
        lookups: lkFields,
      }
    }
  }
  for (const t of Object.values(next)) {
    if (t.id === tableId || !refs[t.id]) continue
    const cfg = { ...(t.config ?? {}) }
    if (cfg.inputs) cfg.inputs = cfg.inputs.filter((x) => x !== tableId)
    if (cfg.linkedInputs) cfg.linkedInputs = cfg.linkedInputs.filter((x) => x !== tableId)
    const fields = t.fields.filter((f) => !(f.type === 'lookup' && f.lookup?.table === tableId))
    next[t.id] = { ...t, config: cfg, fields }
  }
  next[tableId] = {
    ...target,
    deleted: true,
    trashMeta: { deletedAt: Date.now(), refs },
  }
  return next
}

// 恢复单个已软删除的表：把记录的引用加回对应表，解除 deleted
function restoreDeletedTable(tables: Record<string, TableSchema>, tableId: string): Record<string, TableSchema> {
  const target = tables[tableId]
  if (!target || !target.trashMeta) return tables
  const refs = target.trashMeta.refs
  const next = { ...tables }
  for (const [tid, r] of Object.entries(refs)) {
    const t = next[tid]
    if (!t) continue
    const cfg = { ...(t.config ?? {}) }
    cfg.inputs = Array.from(new Set([...(cfg.inputs ?? []), ...r.inputs]))
    cfg.linkedInputs = Array.from(new Set([...(cfg.linkedInputs ?? []), ...r.linkedInputs]))
    const fields = [...t.fields]
    for (const f of r.lookups) if (!fields.some((x) => x.id === f.id)) fields.push(f)
    next[tid] = { ...t, config: cfg, fields }
  }
  const tgt = { ...next[tableId] }
  delete tgt.trashMeta
  tgt.deleted = false
  next[tableId] = tgt
  return next
}

// 从删除批次栈里移除某个 id（单条恢复/彻底删除时调用）；空批次一并移除
function removeFromTrashStack(stack: string[][], id: string): string[][] {
  return stack
    .map((batch) => batch.filter((x) => x !== id))
    .filter((batch) => batch.length > 0)
}

/** ★ 2026-11-07 白屏修复（用户「画布软件 打开 不显示」）：
 *  读盘数据规范化——每个节点的 fields / records 必须是数组。
 *  根因：聊天「添加节点」曾以 { id, name, kind } 直接 addTable（fields=undefined），
 *  落盘后下次启动渲染层遍历 t.fields 抛 TypeError: e.fields is not iterable，
 *  React 树整体崩溃 → 窗口在、内容全白。这里对历史坏数据就地补齐（幂等）。 */
function normalizeTables(raw: unknown): Record<string, TableSchema> {
  const src = (raw ?? {}) as Record<string, TableSchema>
  const out: Record<string, TableSchema> = {}
  for (const [id, t] of Object.entries(src)) {
    if (!t || typeof t !== 'object') continue
    const fields = Array.isArray(t.fields) ? t.fields.filter((f) => !!f && typeof f === 'object') : []
    out[id] = { ...t, fields }
  }
  return out
}

function normalizeRecords(raw: unknown, tables: Record<string, TableSchema>): Record<string, Row[]> {
  const src = (raw ?? {}) as Record<string, Row[]>
  const out: Record<string, Row[]> = {}
  for (const id of Object.keys(tables)) out[id] = Array.isArray(src[id]) ? src[id] : []
  for (const [id, rows] of Object.entries(src)) if (!(id in out) && Array.isArray(rows)) out[id] = rows
  return out
}

function readActiveId(pages: PageMeta[]): string {
  try {
    const saved = localStorage.getItem('activePageId')
    if (saved && pages.some((p) => p.id === saved)) return saved
  } catch {
    /* ignore */
  }
  return pages[0]?.id ?? ''
}

function writeActiveId(id: string) {
  try {
    localStorage.setItem('activePageId', id)
  } catch {
    /* ignore */
  }
}

// 默认画布种子：开始节点 + 材料库 + 用量清单（演示 lookup 关联与数据流）
function seedDefaultData(): SaveableState {
  const tables: Record<string, TableSchema> = {
    [MATERIALS.id]: MATERIALS,
    [BILL.id]: BILL,
    [START.id]: START,
  }
  const records: Record<string, Row[]> = {
    [MATERIALS.id]: MATERIALS_SEED as Row[],
    [BILL.id]: BILL_SEED as Row[],
    [START.id]: [],
  }
  const positions: Record<string, NodePos> = {
    [START.id]: { x: 40, y: 160 },
    [MATERIALS.id]: { x: 360, y: 40 },
    [BILL.id]: { x: 360, y: 320 },
  }
  return { tables, records, positions }
}

function scheduleSave(get: () => BaseState) {
  if (saveTimer) clearTimeout(saveTimer)
  /* 2026-10-24: idle save - big JSON.stringify deferred to main-thread idle */
  saveTimer = setTimeout(() => {
    const s = get()
    if (!s.activePageId) return
    const doSave = () => {
      const s2 = get()
      if (!s2.activePageId) return
      const data = { tables: s2.tables, records: s2.records, positions: s2.positions }
      api
        .savePage(s2.activePageId, data)
        .catch((e) => console.error('[save] fail', e))
    }
    const w = window as unknown as { requestIdleCallback?: (cb: () => void, o?: { timeout: number }) => number }
    if (typeof w.requestIdleCallback === 'function') w.requestIdleCallback(doSave, { timeout: 15000 })
    else setTimeout(doSave, 2000)
  }, 400)
}

export const useBaseStore = create<BaseState>((setRaw, get) => {
  // ── 调试守卫（2026-09-24 白屏排查）：谁把非数组写进 records，当场打出调用栈 ──
  const set: typeof setRaw = ((patch: unknown, replace?: unknown) => {
    try {
      const s = get()
      const next =
        typeof patch === 'function'
          ? (patch as (s: BaseState) => Partial<BaseState>)(s)
          : (patch as Partial<BaseState>)
      if (next && typeof next === 'object' && 'records' in next) {
        for (const [k, v] of Object.entries((next as { records?: Record<string, unknown> }).records ?? {})) {
          if (v !== undefined && !Array.isArray(v)) {
            console.error(
              '[records 守卫] 非法写入 tid=' + k + ' value=' + JSON.stringify(v) +
                '\n' + new Error().stack
            )
          }
        }
      }
    } catch {
      /* 守卫自身绝不干扰正常写入 */
    }
    return setRaw(patch as never, replace as never)
  }) as typeof setRaw
  return {
  tables: {},
  records: {},
  positions: {},
  ready: false,
  pages: [],
  activePageId: null,
  past: [],
  future: [],
  trashStack: [],
  fxHoverNode: null,
  setFxHoverNode: (id) => set({ fxHoverNode: id }),
  selRuleNode: null,
  setSelRuleNode: (id) => set({ selRuleNode: id }),
  nodesFolded: false,
  setNodesFolded: (v) => set({ nodesFolded: v }),
  edgeLinks: [],
  setEdgeLinks: (ids) => set({ edgeLinks: ids }),
  pdfmergeProgress: {},
  setPdfmergeProgress: (nodeId, pct) => set((s0) => { const m = { ...s0.pdfmergeProgress }; if (pct === null) delete m[nodeId]; else m[nodeId] = pct; return { pdfmergeProgress: m } }),
  roundRunState: 'off',
  setRoundRunState: (s) => set({ roundRunState: s }),
  init: async () => {
    try {
      let pages = await api.listPages()
      if (!pages.length) {
        const id = 'default'
        const data = seedDefaultData()
        await api.savePage(id, { name: '画布1', ...data })
        pages = [{ id, name: '画布1' }]
      }
      const activeId = readActiveId(pages)
      const full = await api.getPage(activeId)
      const loadedTables = normalizeTables(full.tables)
      set({
        tables: loadedTables,
        records: normalizeRecords(full.records, loadedTables),
        positions: full.positions as Record<string, NodePos>,
        pages,
        activePageId: activeId,
        ready: true,
        past: [],
        future: [],
        trashStack: [],
      })
      /* ★ 2026-10-24 一次性迁移：pdfItems 里内嵌的 b64 → IndexedDB，config 瘦身。
         失败静默（旧数据仍在 config，功能不受影响，只是没瘦身）。
         ★ 提取为函数 + 3s 补跑：页数据为异步二次恢复（竞态会覆盖迁移成果），
         补跑幂等（已迁移项被条件跳过）。 */
      const migrateB64Payloads = async () => {
        try {
          const { putBlob } = await import('./pdfBlobStore')
          /* ★ 2026-10-25 原生解码：fetch data URL 走浏览器 C++ 实现，
             比 atob 逐字符循环快 20 倍以上（MB 级 b64 从秒级到毫秒级） */
          const b64ToU8 = async (b64: string) => {
            const blob = await (await fetch('data:application/octet-stream;base64,' + b64)).blob()
            return new Uint8Array(await blob.arrayBuffer())
          }
          const cur = get()
          for (const t of Object.values(cur.tables)) {
            const items = (t?.config as { pdfItems?: { id: string; name?: string; b64?: string; inDb?: boolean }[] } | undefined)?.pdfItems
            if (!items?.length) continue
            let changed = false
            const next: { id: string; name?: string; b64?: string; inDb?: boolean }[] = []
            for (const it of items) {
              if (it.b64 && it.id) {
                try { await putBlob('pdfitem-' + it.id, await b64ToU8(it.b64)); next.push({ ...it, b64: undefined, inDb: true }); changed = true }
                catch { next.push(it) }
              } else next.push(it)
            }
            /* photoItems 迁移：b64 → 内容寻址 IndexedDB（photoDbId，SHA-1 去重） */
            const phItems = (t?.config as { photoItems?: { id: string; name: string; type: string; b64?: string; photoDbId?: string }[] } | undefined)?.photoItems
            if (phItems?.some((p) => p.b64 && !p.photoDbId)) {
              try {
                const { putContentAddressed: putPhoto } = await import('./pdfBlobStore')
                for (const p of phItems) {
                  if (!p.b64 || p.photoDbId) continue
                  const bin = atob(p.b64)
                  p.photoDbId = await putPhoto('photo', new Uint8Array(bin.length).map((_, i2) => bin.charCodeAt(i2)))
                  p.b64 = undefined
                }
                const curP = get().tables[t.id]
                const nextP = { ...curP, config: { ...(curP?.config ?? {}), photoItems: phItems } }
                set({ ...get(), tables: { ...get().tables, [t.id]: nextP } } as Partial<BaseState>)
              } catch { /* 迁移失败保留原 b64 */ }
            }
                        /* ooBytes64 迁移：b64 → IndexedDB（ooBytesDbId），config 剥离 */
            const ooB = (t?.config as { ooBytes64?: string; ooBytesDbId?: string } | undefined)
            if (ooB?.ooBytes64 && !ooB.ooBytesDbId) {
              try {
                const { putBlob: putOo } = await import('./pdfBlobStore')
                const bin = atob(ooB.ooBytes64)
                await putOo('oobytes-' + t.id, new Uint8Array(bin.length).map((_, i2) => bin.charCodeAt(i2)))
                const curT3 = get().tables[t.id]
                const nextTables3 = { ...get().tables, [t.id]: { ...curT3, config: { ...(curT3?.config ?? {}), ooBytesDbId: 'oobytes-' + t.id, ooBytes64: undefined } } }
                set(nextTables3 as Partial<BaseState>)
              } catch { /* 迁移失败保留原 b64 */ }
            }
                        const hist = (t?.config as { pdfHistory?: { id: string; name: string; at: number; b64?: string; inDb?: boolean }[] } | undefined)?.pdfHistory
            if (hist?.length) {
              let hChanged = false
              const nextH = hist.map((h) => {
                if (h.b64) { hChanged = true; return { ...h, b64: undefined, inDb: true } }
                return h
              })
              if (hChanged) {
                const curT2 = get().tables[t.id]
                const nextTables2 = { ...get().tables, [t.id]: { ...curT2, config: { ...(curT2?.config ?? {}), pdfHistory: nextH } } }
                set(nextTables2 as Partial<BaseState>)
              }
            }
            if (changed) {
              const curT = get().tables[t.id]
              const nextTables = { ...get().tables, [t.id]: { ...curT, config: { ...(curT?.config ?? {}), pdfItems: next } } }
              set(nextTables as Partial<BaseState>)
            }
          }
        } catch (eM) { console.error('[迁移] b64→IDB 失败（保留原数据）:', eM) }
      }
      void migrateB64Payloads().catch((eO) => console.error('[迁移] 外层异常:', eO))
      /* ★ 页数据为异步二次恢复（App 层可能覆盖迁移成果），3s 后补跑一次（幂等） */
      setTimeout(() => { void migrateB64Payloads() }, 3000)
    } catch (e) {
      console.error('[底座] 初始化失败，使用空画布', e)
      set({ ready: true, pages: [], activePageId: null, tables: {}, records: {}, positions: {}, past: [], future: [], trashStack: [] })
    }
  },

  saveNow: async () => {
    if (saveTimer) {
      clearTimeout(saveTimer)
      saveTimer = null
    }
    const s = get()
    if (!s.activePageId) return
    /* ★ 2026-10-25 终极兜底：保存前强制 b64 剥离（幂等）——
       无论 init 迁移是否被时序竞态覆盖，保存产物必然轻量。
       剥离的字节全部先写入 IndexedDB（数据不丢），config 只留索引。 */
    try {
      const { putBlob, putContentAddressed } = await import('./pdfBlobStore')
      /* ★ 2026-10-25 原生解码（同上） */
      const b64ToU8 = async (b64: string) => {
        const blob = await (await fetch('data:application/octet-stream;base64,' + b64)).blob()
        return new Uint8Array(await blob.arrayBuffer())
      }
      const cur0 = get()
      for (const t of Object.values(cur0.tables)) {
        const cfg0 = (t?.config ?? {}) as {
          pdfItems?: { id: string; name?: string; b64?: string; inDb?: boolean }[]
          pdfHistory?: { b64?: string; inDb?: boolean }[]
          ooBytes64?: string
          ooBytesDbId?: string
          photoItems?: { id: string; name: string; type: string; b64?: string; photoDbId?: string }[]
        }
        let touched = false
        const patchCfg: Record<string, unknown> = {}
        if (cfg0.pdfItems?.some((it) => it.b64 && !it.inDb)) {
          patchCfg.pdfItems = await Promise.all(cfg0.pdfItems.map(async (it) => {
            if (!it.b64 || it.inDb) return it
            await putBlob('pdfitem-' + it.id, await b64ToU8(it.b64))
            touched = true
            return { ...it, b64: undefined, inDb: true }
          }))
        }
        if (cfg0.pdfHistory?.some((h) => h.b64 && !h.inDb)) {
          patchCfg.pdfHistory = await Promise.all(cfg0.pdfHistory.map(async (h) => {
            const hh = h as { id?: string; at?: number; b64?: string; inDb?: boolean }
            if (!hh.b64 || hh.inDb) return h
            await putBlob('pdfhist-' + (hh.id ?? 'x') + '-' + (hh.at ?? 0), await b64ToU8(hh.b64))
            touched = true
            return { ...h, b64: undefined, inDb: true }
          }))
        }
        if (cfg0.ooBytes64 && !cfg0.ooBytesDbId) {
          await putBlob('oobytes-' + t.id, await b64ToU8(cfg0.ooBytes64))
          patchCfg.ooBytesDbId = 'oobytes-' + t.id
          patchCfg.ooBytes64 = undefined
          touched = true
        }
        if (cfg0.photoItems?.some((p) => p.b64 && !p.photoDbId)) {
          patchCfg.photoItems = await Promise.all(cfg0.photoItems.map(async (p) => {
            if (!p.b64 || p.photoDbId) return p
            p.photoDbId = await putContentAddressed('photo', await b64ToU8(p.b64))
            touched = true
            return { ...p, b64: undefined }
          }))
        }
        if (touched) {
          const curT = get().tables[t.id]
          set({ ...get(), tables: { ...get().tables, [t.id]: { ...curT, config: { ...(curT?.config ?? {}), ...patchCfg } } } } as Partial<BaseState>)
          console.log('[保存前迁移]', t.id, '已剥离 b64 → IndexedDB')
        }
      }
    } catch (eMig) {
      console.error('[保存前迁移] 失败（保留原数据继续保存）:', eMig)
    }
    // ★ 编号回填（2026-10-12）：缺 nodeNo 的节点在落盘前补齐——节点与存储的对照键
    let tables = s.tables
    let needNo = false
    for (const t of Object.values(tables)) {
      if (!(t.config as { nodeNo?: string } | undefined)?.nodeNo) { needNo = true; break }
    }
    if (needNo) {
      let next = 1
      const existing = new Set<string>()
      for (const t of Object.values(tables)) {
        const no = (t.config as { nodeNo?: string } | undefined)?.nodeNo
        if (no) existing.add(no)
      }
      const out: typeof tables = {}
      for (const id of Object.keys(tables).sort()) {
        const t = tables[id]
        const cfg = (t.config as { nodeNo?: string } | undefined)?.nodeNo
        if (cfg) { out[id] = t; continue }
        while (existing.has('N' + String(next).padStart(3, '0'))) next++
        const no = 'N' + String(next).padStart(3, '0')
        existing.add(no)
        out[id] = { ...t, config: { ...(t.config || {}), nodeNo: no } }
        next++
      }
      tables = out
      set({ tables })
    }
    /* ★ 2026-10-12 类型-名称一致性回填（仅限无内容的 OO 节点）：
     *   名字含「工作簿」→ oosheet；含「文档」→ ooword。有内容的节点不动（避免格式错配）。 */
    for (const id of Object.keys(tables)) {
      const t = tables[id]
      if (t.kind !== 'oosheet' && t.kind !== 'ooword') continue
      if ((t.config as { ooBytes64?: string } | undefined)?.ooBytes64) continue
      const wantsSheet = /工作簿/.test(t.name)
      const wantsWord = /文档/.test(t.name)
      if (wantsSheet && t.kind !== 'oosheet') tables = { ...tables, [id]: { ...t, kind: 'oosheet' } }
      else if (wantsWord && t.kind !== 'ooword') tables = { ...tables, [id]: { ...t, kind: 'ooword' } }
    }
    const data = { tables: s.tables, records: s.records, positions: s.positions }
    await api.savePage(s.activePageId, data).catch((e) => console.error('[底座] 落盘失败', e))
  },

  switchPage: async (id) => {
    const s = get()
    if (id === s.activePageId) return
    // 先落盘当前页
    if (s.activePageId) {
      await api
        .savePage(s.activePageId, { tables: s.tables, records: s.records, positions: s.positions })
        .catch((e) => console.error('[底座] 切换前落盘失败', e))
    }
    try {
      const full = await api.getPage(id)
      writeActiveId(id)
      const loadedTables = normalizeTables(full.tables)
      set({
        activePageId: id,
        tables: loadedTables,
        records: normalizeRecords(full.records, loadedTables),
        positions: full.positions as Record<string, NodePos>,
        past: [],
        future: [],
        trashStack: [],
      })
    } catch (e) {
      console.error('[底座] 加载页面失败', e)
    }
  },

  newPage: async (name) => {
    const meta = await api.createPage(name)
    set((s) => ({ pages: [...s.pages, { id: meta.id, name: meta.name }] }))
    await get().switchPage(meta.id)
    return meta.id
  },

  importPage: async (name, tables, records, positions) => {
    // ★ 服务端 createPage 建空壳，savePage 覆写内容 —— 与 newPage+手动编辑落盘同一条链路
    const meta = await api.createPage(name)
    await api.savePage(meta.id, { name: meta.name, tables, records, positions })
    set((s) => ({ pages: [...s.pages, { id: meta.id, name: meta.name }] }))
    await get().switchPage(meta.id)
    return meta.id
  },

  renamePage: async (id, name) => {
    const s = get()
    if (id === s.activePageId) {
      await api
        .savePage(id, { name, tables: s.tables, records: s.records, positions: s.positions })
        .catch((e) => console.error('[底座] 改名落盘失败', e))
    } else {
      try {
        const full = await api.getPage(id)
        await api
          .savePage(id, {
            name,
            tables: full.tables as Record<string, TableSchema>,
            records: full.records as Record<string, Row[]>,
            positions: full.positions as Record<string, NodePos>,
          })
          .catch((e) => console.error('[底座] 改名落盘失败', e))
      } catch (e) {
        console.error('[底座] 改名读取失败', e)
      }
    }
    set((st) => ({ pages: st.pages.map((p) => (p.id === id ? { ...p, name } : p)) }))
  },

  deletePage: async (id) => {
    await api.deletePage(id).catch((e) => console.error('[底座] 删除页失败', e))
    const remaining = get().pages.filter((p) => p.id !== id)
    set({ pages: remaining })
    if (get().activePageId === id) {
      if (remaining.length) await get().switchPage(remaining[0].id)
      else {
        // 删光了：新建一个空白页
        await get().newPage('画布1')
      }
    }
  },

  addTable: (schema) => {
    /* ★ 2026-11-07 fields 兜底：调用方漏传 fields 会让节点缺字段表，
       渲染层 t.fields.map 直接抛异常崩整屏（详见 normalizeTables 注释） */
    const safeSchema: TableSchema = { ...schema, fields: Array.isArray(schema.fields) ? schema.fields : [] }
    set((s) => ({
      tables: { ...s.tables, [safeSchema.id]: { ...safeSchema, config: { ...(safeSchema.config || {}), nodeNo: nextNoOf(s.tables) } } },
      records: { ...s.records, [safeSchema.id]: [] },
    }))
    scheduleSave(get)
  },

  addField: (tableId, field) => {
    set((s) => {
      const t = s.tables[tableId]
      if (!t) return s
      return {
        tables: { ...s.tables, [tableId]: { ...t, fields: [...t.fields, field] } },
      }
    })
    scheduleSave(get)
  },

  removeField: (tableId, fieldId) => {
    set((s) => {
      const t = s.tables[tableId]
      if (!t) return s
      return {
        tables: { ...s.tables, [tableId]: { ...t, fields: t.fields.filter((f) => f.id !== fieldId) } },
      }
    })
    scheduleSave(get)
  },

  updateField: (tableId, fieldId, patch) => {
    set((s) => {
      const t = s.tables[tableId]
      if (!t) return s
      return {
        tables: {
          ...s.tables,
          [tableId]: { ...t, fields: t.fields.map((f) => (f.id === fieldId ? { ...f, ...patch, id: f.id } : f)) },
        },
      }
    })
    scheduleSave(get)
  },

  renameTable: (tableId, name) => {
    set((s) => {
      const t = s.tables[tableId]
      if (!t) return s
      return { tables: { ...s.tables, [tableId]: { ...t, name } } }
    })
    scheduleSave(get)
  },

  duplicateTable: (tableId) => {
    const newId = `${tableId}_copy_${Date.now()}`
    set((s) => {
      const t = s.tables[tableId]
      if (!t) return s
      return {
        tables: { ...s.tables, [newId]: { ...t, id: newId, config: { ...(t.config || {}), nodeNo: nextNoOf(s.tables) }, name: `${t.name} 副本` } },
        records: { ...s.records, [newId]: [] },
      }
    })
    scheduleSave(get)
    return newId
  },

  removeTable: (tableId) => {
    set((s) => {
      const tables = { ...s.tables }
      delete tables[tableId]
      const records = { ...s.records }
      delete records[tableId]
      const positions = { ...s.positions }
      delete positions[tableId]
      // 清理其它表里指向该表的 lookup 字段（否则留下孤儿连线）
      for (const t of Object.values(tables)) {
        if (t.fields.some((f) => f.type === 'lookup' && f.lookup?.table === tableId)) {
          tables[t.id] = {
            ...t,
            fields: t.fields.filter((f) => !(f.type === 'lookup' && f.lookup?.table === tableId)),
          }
        }
      }
      return { tables, records, positions }
    })
    scheduleSave(get)
  },

  // 软删除：把节点移入回收站（画布不再渲染，但数据保留可恢复）。
  // 删除时记录其它节点指向本表的引用（数据流来源 inputs/linkedInputs、lookup 字段），并从那些表里移除，
  // 这样连线不会悬空；恢复时再原样加回去，连线自动复现。
  // 单条删除归为一个 [id] 批次；批量删除用 trashTables 归为同一个批次，便于「逐步恢复上一次删除」。
  trashTable: (tableId) => {
    get().trashTables([tableId])
  },

  trashTables: (tableIds) => {
    const ids = tableIds.filter((id) => {
      const t = get().tables[id]
      return t && !t.deleted
    })
    if (!ids.length) return
    set((s) => {
      let tables = s.tables
      for (const id of ids) tables = softDeleteTable(tables, id)
      const stack = [...(s.trashStack ?? []), ids]
      return { tables, trashStack: stack }
    })
    scheduleSave(get)
  },

  // 恢复单个：把节点从回收站放回画布，并从删除批次栈中移除该 id（空批次一并移除）
  restoreTable: (tableId) => {
    set((s) => {
      const tables = restoreDeletedTable(s.tables, tableId)
      const stack = removeFromTrashStack(s.trashStack ?? [], tableId)
      return { tables, trashStack: stack }
    })
    scheduleSave(get)
  },

  // 恢复上一次删除的内容（整个批次，按删除的逆序逐步找回；常用于批量删除后逐步还原）
  restoreLastTrash: () => {
    let restored = 0
    set((s) => {
      const stack = [...(s.trashStack ?? [])]
      if (!stack.length) return s
      const batch = stack.pop() as string[]
      let tables = s.tables
      for (const id of batch) {
        if (tables[id] && tables[id].trashMeta) {
          tables = restoreDeletedTable(tables, id)
          restored++
        }
      }
      return { tables, trashStack: stack }
    })
    scheduleSave(get)
    return restored
  },

  // 彻底删除：从回收站永久移除（真正删库），并从批次栈移除
  purgeTable: (tableId) => {
    get().removeTable(tableId)
    set((s) => ({ trashStack: removeFromTrashStack(s.trashStack ?? [], tableId) }))
  },

  // 清空回收站：把所有已软删除的节点永久删除，并清空批次栈
  emptyTrash: () => {
    const ids = Object.values(get().tables)
      .filter((t) => t.deleted)
      .map((t) => t.id)
    ids.forEach((id) => get().removeTable(id))
    set({ trashStack: [] })
  },

  setNodePosition: (tableId, pos) => {
    set((s) => ({ positions: { ...s.positions, [tableId]: pos } }))
    scheduleSave(get)
  },

  setNodePositions: (updates) => {
    if (!updates.length) return
    set((s) => {
      const positions = { ...s.positions }
      for (const u of updates) positions[u.id] = u.pos
      return { positions }
    })
    scheduleSave(get)
  },

  setNodeConfig: (tableId, patch) => {
    set((s) => {
      const t = s.tables[tableId]
      if (!t) return s
      return {
        tables: {
          ...s.tables,
          [tableId]: { ...t, config: { ...t.config, ...patch } },
        },
      }
    })
    scheduleSave(get)
  },
  setTableKind: (tableId, kind) => {
    set((s) => {
      const t = s.tables[tableId]
      if (!t || t.kind === kind) return s
      /* ★ 2026-10-12 用户指令「工作簿就打开工作簿，文档就打开文档」：kind 切换时同步名称关键词，
       *   防止「名字叫工作簿、打开是 Word」的类型错位 */
      let name = t.name
      if (kind === 'ooword') name = name.replace(/工作簿/g, '文档')
      else if (kind === 'oosheet') name = name.replace(/文档/g, '工作簿')
      return {
        tables: {
          ...s.tables,
          [tableId]: { ...t, kind, name },
        },
      }
    })
    scheduleSave(get)
  },

  setTableData: (tableId, fields, rows) => {
    set((s) => {
      const t = s.tables[tableId]
      if (!t) return s
      return {
        tables: { ...s.tables, [tableId]: { ...t, fields } },
        records: { ...s.records, [tableId]: rows },
      }
    })
    scheduleSave(get)
  },

  /* ── 规则锁定（2026-10-05：圆形「规则节点」给连上的节点发规则）────────
     用户原话：「一旦匹配了，默认不能每个改规则，只能改主流程的节点进行修改。
                修改后所有相同规则的节点规则同步改。」

     ★★★ 本段唯一且最重要的实现约束 ★★★
     下面的 setRule **只写规则节点自己一份**，绝不给辅节点写副本。
     「同步」靠effectiveRule() 读时跳回源完成 —— 见 ruleSync.ts 顶部说明。
     若将来有人在此加「同步写进辅节点」的循环，症状是规则副本漂移：
     画面正常，只是某处不跟着变，极难发现。故此处刻意不写循环，也不加注释式TODO。

     ★★ 撤销栈：以下四个动作**不自己调 pushHistory**。
        全项目写操作统一由 App 层的 push() 包一层（同 setNodeConfig 的既有做法）。
        若这里再自压一次，一次操作会压两个相同快照 —— 症状是「撤销一次没反应」，
        要按两下才回退，极易被误判成撤销坏了。
        （这行注释不是啰嗦：初版在此自压了 4 处，实测才发现。）  */

  /** 改某kind 的规则（只写规则节点自身；全部辅节点立刻生效） */
  setRule: (ruleNodeId: string, kind: string, patch: Partial<NodeRuleSet>) => {
    set((s) => {
      const t = s.tables[ruleNodeId]
      if (!t) return s
      const prevKind = t.config?.rules?.[kind] ?? {}
      return {
        tables: {
          ...s.tables,
          [ruleNodeId]: {
            ...t,
            config: { ...t.config, rules: { ...t.config?.rules, [kind]: { ...prevKind, ...patch } } },
          },
        },
      }
    })
    scheduleSave(get)
  },

  /** 清掉某 kind 的规则（该 kind 的辅节点回落到自身 config） */
  clearRule: (ruleNodeId: string, kind: string) => {
    set((s) => {
      const t = s.tables[ruleNodeId]
      if (!t || !t.config?.rules) return s
      const next = { ...t.config.rules }
      delete next[kind]
      return {
        tables: {
          ...s.tables,
          [ruleNodeId]: { ...t, config: { ...t.config, rules: next } },
        },
      }
    })
    scheduleSave(get)
  },

  /** 从规则节点断开：目标节点脱离锁定，回落自身 config */
  unlinkFromRuleNode: (ruleNodeId: string, targetId: string) => {
    set((s) => {
      const target = s.tables[targetId]
      if (!target) return s
      const ups = target.config?.linkedInputs ?? []
      if (!ups.includes(ruleNodeId)) return s
      // ★ 连线断了必须一并清掉 ruleRef，否则留下指向已解除锁定的悬空引用。
      //   当前判定不会因此出错（bindingsOfRuleNode 要求 linkedInputs 里有，
      //   modeOfNode 也会校验 ruleNodeId 匹配），但数据留着不干净：
      //   以后若有人改成「按 ruleRef 判锁定」就会踩坑，且很难查。
      const { ruleRef, ...rest } = target.config ?? {}
      /* ★ 显式标注 Record<string, TableSchema>：不标注时 TS 会把
         `{...s.tables, [targetId]: {...}}` 推成「NodeConfig | 展开对象」的联合类型，
         后面读 t.config.ruleRef 就报TS2339（ruleRef 不在联合的某一项上）。
         显式标注让全表保持单一类型，规则读写不再报错。*/
      const tables: Record<string, TableSchema> = {
        ...s.tables,
        [targetId]: {
          ...target,
          config: { ...rest, linkedInputs: ups.filter((x) => x !== ruleNodeId) },
        },
      }
      /* ★ 移出的若是主节点，必须把主位让给组内下一个成员 —— 与 setRuleFollow
         里的同一道防护对称。
           漏了这步的症状：移出主节点后组里一个 primary 都没有，面板显示
           「暂无主节点」，而圆形节点上的规则编辑区仍可点，点了规则不生效 ——
           因为「以主节点为锚」这个语义断了。两处必须一起改，漏一处就漏一个入口。*/
      if (ruleRef?.ruleNodeId === ruleNodeId && ruleRef.mode === 'primary') {
        const heir = Object.values(tables).find((t) => {
          if (t.deleted || isRuleNodeKind(t.kind) || t.id === targetId) return false
          // 局部取出 config：config 是可选的，写 t.config?.a 判过之后
          // 再读 t.config.b，TS 不会记得前面那层判空（两个独立表达式）
          const cfg = t.config
          if (!cfg) return false
          if (!(cfg.linkedInputs ?? []).includes(ruleNodeId)) return false
          // 候选不能是「明确不跟随」的成员
          return !(cfg.ruleRef?.ruleNodeId === ruleNodeId && cfg.ruleRef.mode === 'ignore')
        })
        if (heir) {
          tables[heir.id] = { ...heir, config: { ...heir.config, ruleRef: { ruleNodeId, mode: 'primary' } } }
        }
        // 找不到继承人时不指定新的 —— 组变成「无主」。这与 setRuleFollow
        // 「一个成员都没有 → 拒绝」不同：这里是成员确实被移走了，
        // 拦着不移除反而违背用户本意。让 primaryOf 的兜底去显示这个状态。
      }
      return { tables }
    })
    scheduleSave(get)
  },

  /** 设主节点（用户：第一个默认是可修改的主节点；双击面板可改主次） */
  setRulePrimary: (ruleNodeId: string, targetId: string) => {
    set((s) => {
      const rule = s.tables[ruleNodeId]
      const target = s.tables[targetId]
      if (!rule || !target) return s
      if ((target.config?.linkedInputs ?? []).indexOf(ruleNodeId) < 0) return s
      /* ★ 身份槽被别的规则组占用 → 拒绝。
         ruleRef 是单槽，写上去就把另一个组的身份抹掉了。症状是 A 组的主节点
         被人在 B 组里改掉，A 组面板随之显示别人是主节点 —— 跨组数据损坏。
         与其默默改错，不如不改；UI 侧会把这个按钮置灰并说明原因。*/
      const owner = identityOwnerOf(targetId, s.tables, ruleNodeId)
      if (owner) return s
      // 已被锁定的成员全部重算：目标转primary，其余非 ignore 的一律降为 follow。
      // ★ 原来那位主节点不能变成「无主」—— 组内必须恒有且仅有一个 primary，
      //   否则「改规则以主节点为锚」的语义就断了。
      const next: Record<string, TableSchema> = {}
      for (const t of Object.values(s.tables)) {
        if (t.deleted || isRuleNodeKind(t.kind)) continue
        if (!(t.config?.linkedInputs ?? []).includes(ruleNodeId)) continue
        if (t.id === targetId) {
          next[t.id] = { ...t, config: { ...t.config, ruleRef: { ruleNodeId, mode: 'primary' } } }
        } else if (t.config?.ruleRef?.ruleNodeId === ruleNodeId && t.config.ruleRef.mode === 'primary') {
          next[t.id] = { ...t, config: { ...t.config, ruleRef: { ruleNodeId, mode: 'follow' } } }
        } else {
          // 保持原身份（含 ignore）；但要补上可能缺失的 ruleRef
          next[t.id] = t.config?.ruleRef?.ruleNodeId === ruleNodeId
            ? t
            : { ...t, config: { ...t.config, ruleRef: { ruleNodeId, mode: 'follow' } } }
        }
      }
      return { tables: { ...s.tables, ...next } }
    })
    scheduleSave(get)
  },

  /** 切换跟随 / 不跟随（用户：双击能改「跟随不跟随」） */
  setRuleFollow: (ruleNodeId: string, targetId: string, follow: boolean) => {
    set((s) => {
      const target = s.tables[targetId]
      if (!target) return s
      if ((target.config?.linkedInputs ?? []).indexOf(ruleNodeId) < 0) return s
      // ★ 身份槽被别的规则组占用 → 拒绝（理由同 setRulePrimary）
      if (identityOwnerOf(targetId, s.tables, ruleNodeId)) return s
      const cur = target.config?.ruleRef?.ruleNodeId === ruleNodeId ? target.config.ruleRef.mode ?? 'follow' : 'follow'
      /* 「跟随」= true 时不能原样保留 cur。
         ★ 这里曾写成 `follow ? cur : 'ignore'` —— 当 cur 恰是 'ignore' 时，
           点「跟随」会把 mode 又写成 'ignore'，即**按钮点了没反应**。
           症状：面板上「不跟随」这一行怎么点都切不回来，只能重连连线。
           正确语义：跟随=true 且原来是 ignore → 必须真的变成 follow；
           只有 primary 才保留 primary（见下面主节点让位那段）。*/
      let mode: 'primary' | 'follow' | 'ignore' =
        follow ? (cur === 'primary' ? 'primary' : 'follow') : 'ignore'
      // 主节点不能直接设成 ignore —— 组内必须留一个能改规则的节点。
      // 症状：全组都不跟随后，圆形节点上的规则改了没人执行，
      // 而 UI 上「规则」按钮仍可点，点了就是「改了没反应」。
      // 故主节点要退跟随时，把它降为 follow 并把主位让给组内下一个成员。
      if (mode === 'ignore' && cur === 'primary') {
        const others = Object.values(s.tables).filter(
          (t) =>
            !t.deleted &&
            !isRuleNodeKind(t.kind) &&
            t.id !== targetId &&
            (t.config?.linkedInputs ?? []).includes(ruleNodeId),
        )
        // 一个成员都没有 → 拒绝，不让组变成「无主」
        if (!others.length) return s
        mode = 'follow'
        const tables = { ...s.tables }
        for (const o of others) {
          const om =
            o.config?.ruleRef?.ruleNodeId === ruleNodeId ? o.config.ruleRef.mode ?? 'follow' : 'follow'
          tables[o.id] = {
            ...o,
            config: { ...o.config, ruleRef: { ruleNodeId, mode: om === 'ignore' ? 'ignore' : 'primary' } },
          }
        }
        return {
          tables: {
            ...tables,
            [targetId]: { ...target, config: { ...target.config, ruleRef: { ruleNodeId, mode } } },
          },
        }
      }
      return {
        tables: {
          ...s.tables,
          [targetId]: { ...target, config: { ...target.config, ruleRef: { ruleNodeId, mode } } },
        },
      }
    })
    scheduleSave(get)
  },

  addRow: (tableId, row = {}) => {
    set((s) => ({
      records: { ...s.records, [tableId]: [...(s.records[tableId] ?? []), row] },
    }))
    scheduleSave(get)
  },

  removeRow: (tableId, rowIdx) => {
    set((s) => {
      const rows = s.records[tableId] ?? []
      if (rowIdx < 0 || rowIdx >= rows.length) return s
      return { records: { ...s.records, [tableId]: rows.filter((_, i) => i !== rowIdx) } }
    })
    scheduleSave(get)
  },

  setCell: (tableId, rowIdx, fieldId, value) => {
    /* ★ 2026-10-17 用户指令「软件里的每个操作都进日志」：
     *   这里是**所有单元格改动的唯一收口**（表格、表单、AG Grid、脚本都走它），
     *   在这里埋一条即可覆盖「用户改了哪个节点的哪个字段」，
     *   又不会因为视图层多路调用而重复记录。 */
    const before = (get().records[tableId] || [])[rowIdx]?.[fieldId]
    set((s) => {
      const rows = s.records[tableId] ?? []
      const next = rows.slice()
      next[rowIdx] = { ...next[rowIdx], [fieldId]: value }
      return { records: { ...s.records, [tableId]: next } }
    })
    try {
      const t = get().tables[tableId]
      const f = t?.fields?.find((x) => x.id === fieldId)
      audit(
        '数据',
        '改单元格',
        `节点=${t ? t.name + '(' + t.id + ')' : tableId} 行=${rowIdx + 1} 字段=${f ? f.name : fieldId}` +
          ` 旧值=${JSON.stringify(before) ?? '空'} 新值=${JSON.stringify(value) ?? '空'}`,
      )
    } catch { /* 审计失败不影响写入 */ }
    scheduleSave(get)
  },

  getRows: (tableId) => get().records[tableId] ?? [],

  // 写操作前调用：把当前状态压入历史栈（清除 redo 分支），供 undo 回退
  pushHistory: () => {
    const { tables, records, positions, past } = get()
    const snap: Snapshot = {
      tables: clone(tables),
      records: clone(records),
      positions: clone(positions),
    }
    const next = [...past, snap]
    if (next.length > MAX_HISTORY) next.shift()
    set({ past: next, future: [] })
  },

  undo: () => {
    const { past, future, tables, records, positions } = get()
    if (!past.length) return
    const prev = past[past.length - 1]
    const cur: Snapshot = { tables: clone(tables), records: clone(records), positions: clone(positions) }
    set({
      tables: prev.tables,
      records: prev.records,
      positions: prev.positions,
      past: past.slice(0, -1),
      future: [...future, cur].slice(-MAX_HISTORY),
    })
    scheduleSave(get)
  },

  redo: () => {
    const { past, future, tables, records, positions } = get()
    if (!future.length) return
    const nxt = future[future.length - 1]
    const cur: Snapshot = { tables: clone(tables), records: clone(records), positions: clone(positions) }
    set({
      tables: nxt.tables,
      records: nxt.records,
      positions: nxt.positions,
      past: [...past, cur],
      future: future.slice(0, -1),
    })
    scheduleSave(get)
  }
  }
})
