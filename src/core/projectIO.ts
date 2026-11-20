/**
 * 工程资产与控件状态的 IO 编排（2026-10-06）。
 *
 * ★ 分层位置：
 *   projectFormat.ts / projectStore.ts / nodeAssets.ts / viewState.ts  → 纯计算 + 抽象 IO
 *   **本文件**→ 编排：把上面几层的产物真正落到工程目录 / 灌回运行时。
 *   它仍然不碰 electron / node:（磁盘路径由 ProjectFs 适配器提供），
 *   放core 下是因为它只是「多步流程的胶水」，不含平台判断。
 *
 * ── 为什么单独一层，而不是把 saveProject 写胖 ──
 *   saveProject 的职责是「把内存状态写成工程文件」——它对每个节点只写一个 JSON。
 *   资产是**字节**，还有「从哪读」「写哪」「回灌到哪」三个问题，
 *   混进 saveProject 会让那个函数的参数膨胀到十几项，且把 IO 关注点
 *   和格式关注点搅在一起。分开之后各自都能独立测试。
 */

import {
  saveProject,
  loadProject,
  assetPath,
  PROJECT_DIRS,
  type ProjectFs,
  type ProjectRef,
  type SaveProjectInput,
  type LoadedProject,
} from './projectStore'
import { planAllNodeAssets, restoreNodeAssets, rewireTemplateIds, type AssetRuntime } from './nodeAssets'
import {
  collectViewStateFromTables,
  applyAllViewState,
  auditViewStateConflicts,
  type ViewStateKV,
} from './viewState'
import type { NodeAssetRef, ProjectManifest, ProjectNodeFile, NodeViewState } from './projectFormat'
import type { TableSchema } from './schema'
import type { NodePos, Row } from './store'

/* ────────────────────────── 保存 ────────────────────────── */

/** 保存工程时的外部依赖 */
export interface SaveProjectExtras {
  /** 运行时资产字节来源（templateStore / editStore 的适配器） */
  runtime: AssetRuntime
  /** localStorage 适配器 */
  kv: ViewStateKV
  /**
   * 哪些节点当前有就地编辑稿。
   * ★ 由调用方告知而非本层去查：编辑稿的存在性往往只存在于
   *   打开着的编辑器组件的内存里（编辑器的 useState），
   *   本层去猜会得到「全都���有稿」——于是稿永远不进工程。
   */
  hasEdited?: (nodeId: string) => boolean

  /** 历史记录产物是否随 .wfc 打包（2026-10-24 勾选） */
  includeHistory?: boolean | Set<string>
}

/** 保存结果：把「工程写成了什么」与「过程中发现的问题」一起带回给 UI */
export interface SaveProjectOutcome {
  manifest: ProjectManifest
  /** 资产层面的问题（模板丢失等）——不阻断保存，但必须让用户知道 */
  assetProblems: { nodeId: string; message: string }[]
  /** 资产内容去重：重复字节跳写数量（同文件多节点引用只存一份） */
  dedupInfo?: { count: number }
  /** 实际写入的资产文件数 */
  assetCount: number
}

/**
 * 完整保存：坐标 + 设置 + 连线 + 行数据 + 控件状态 + 资产字节。
 *
 * 写序仍是「先节点、后 project.json」（见 saveProject 的崩溃一致性注释），
 * 资产字节在写project.json **之前**落盘——
 *   顺序很重要：若先写索引再写资产，中途崩溃会留下一个
 *   「清单说有 3 个资产，实际只有 1 个」的工程，
 *   而 loadProject 只报 warn，用户以为存好了。
 */
export async function saveProjectFull(
  fs: ProjectFs,
  ref: ProjectRef,
  input: SaveProjectInput,
  extras: SaveProjectExtras,
  prev?: LoadedProject | null,
): Promise<SaveProjectOutcome> {
  const root = ref.rootDir

  // ① 采集控件状态（localStorage → viewState）
  const viewState = collectViewStateFromTables(input.tables, extras.kv)

  // ② 规划资产（不写盘，只算该写什么）
  const prevRefs: Record<string, NodeAssetRef[]> = {}
  for (const [id, f] of Object.entries(prev?.nodeFiles ?? {})) {
    if (f.assets?.length) prevRefs[id] = f.assets
  }
  const { byNode, problems } = await planAllNodeAssets(input.tables, extras.runtime, {
    hasEdited: extras.hasEdited,
    prevRefs,
    includeHistory: extras.includeHistory,
  })

  // ③ 落资产字节 + 汇总登记
  const assets: Record<string, NodeAssetRef[]> = {}
  let assetCount = 0
  const nodeAssetDir = input.assetDir ?? PROJECT_DIRS.assets
  /* ★ 2026-10-24 资产内容去重（用户实测保存 100+MB：同一文件被多个节点引用时
     逐节点各写一份全尺寸字节）。按内容指纹只写一份，重复项登记指向首写路径——
      读回按 relPath 定位，格式兼容。 */
  const { hashBytes } = await import('./pdfBlobStore')
  /* ★ 2026-10-24 用户指令「文件统一放在一个地方 相同的文件都引用」：
     统一资产仓 assets/content/{指纹}.{ext} —— 所有节点引用同一份字节，
     位置可预测、天然去重、引用共享。旧工程 assets/{nodeId}/ 结构读回兼容
     （按登记 relPath 定位）。 */
  const writtenByFp = new Map<string, string>()
  let dedupCount = 0
  let dedupInfo: { count: number } | undefined
  for (const [id, plan] of Object.entries(byNode)) {
    if (plan.refs.length === 0) continue
    for (const [rel, bytes] of plan.files) {
      if (!fs.writeBytes) {
        // 没有二进制写入能力 = 无法保证资产不被静默损坏。
        // 此时**不静默降级**：直接抛出去，让上层明确知道这次保存不完整。
        throw new Error('当前存储不支持二进制写入，无法保存节点资产（文档/音频/视频）')
      }
      const fp = await hashBytes(bytes)
      const ext = rel.includes('.') ? rel.slice(rel.lastIndexOf('.') + 1) : 'bin'
      const contentRel = `content/${fp}.${ext}`
      // 统一仓路径（fp 为自身生成的 hex、ext 来自已过安全校验的 rel——路径可控）
      const p = `${root}/${nodeAssetDir}/${contentRel}`
      if (!writtenByFp.has(fp)) {
        await fs.writeBytes(p, bytes)
        writtenByFp.set(fp, contentRel)
        assetCount++
      } else {
        dedupCount++
      }
      // 登记条目指向统一仓路径
      const refItem = plan.refs.find((r) => r.relPath === rel || r.relPath.endsWith('/' + rel))
      if (refItem) refItem.relPath = contentRel
    }
    assets[id] = plan.refs
  }
  if (dedupCount > 0) {
    // 去重信息随返回值带给调用方（日志可见）
    dedupInfo = { count: dedupCount }
  }

  // ④ 写工程文件（含 viewState / assets 登记）
  const manifest = await saveProject(
    fs,
    ref,
    {
      ...input,
      viewState,
      assets,
      assetDir: input.assetDir,
      prev: input.prev ?? prev?.manifest,
      prevNodeFiles: input.prevNodeFiles ?? prev?.nodeFiles,
    },
  )

  return { manifest, assetProblems: problems, assetCount, dedupInfo }
}

/* ────────────────────────── 打开 ────────────────────────── */

/** 打开工程的结果：在 LoadedProject 之外补上「资产已入库」的映射与冲突清单 */
export interface OpenProjectOutcome extends LoadedProject {
  /** nodeId → { 旧 tplId: 新 tplId }；调用方须据此改写 schema.config */
  tplIdMap: Record<string, Record<string, string>>
  /** 资产恢复过程中的问题 */
  assetProblems: string[]
  /** 控件状态与本地不一致之处（由上层决定是否提示） */
  viewConflicts: { nodeId: string; field: string; local: string; fromProject: string }[]
}

/**
 * 完整打开：读工程 + 资产字节灌回运行时 + 控件状态回灌 localStorage。
 *
 * ★ 顺序：先恢复资产再回灌控件状态。
 *   反过来的话，OO 节点在「模板还没入库」的时刻就可能渲染编辑器，
 *   触发一次「模板已不在模板库中」的错误提示 —— 而模板其实马上就回来了。
 *   用户会看到一个本不该出现的报错。
 */
export async function openProjectFull(
  fs: ProjectFs,
  ref: ProjectRef,
  extras: { runtime: AssetRuntime; kv: ViewStateKV },
): Promise<OpenProjectOutcome> {
  const loaded = await loadProject(fs, ref)
  const root = ref.rootDir

  // ① 资产字节 → 运行时存储
  let tplIdMap: Record<string, Record<string, string>> = {}
  const assetProblems: string[] = []
  if (fs.readBytes) {
    const out = await restoreNodeAssets(
      loaded.nodeFiles,
      async (nodeId, ref2) => {
        const file = loaded.nodeFiles[nodeId]
        /* ★ 2026-10-24 统一资产仓：content/ 开头的 relPath 无节点目录前缀，
           直接拼 assets 根；旧格式（{nodeId}/{rel}）走 assetPath（含校验） */
        const p = ref2.relPath.startsWith('content/')
          ? `${root}/${file?.assetDir ?? PROJECT_DIRS.assets}/${ref2.relPath}`
          : assetPath(root, nodeId, ref2, file?.assetDir)
        if (!p) return null
        return fs.readBytes!(p)
      },
      extras.runtime,
    )
    tplIdMap = out.tplIdMap
    assetProblems.push(...out.problems)
  }

  // ② 控件状态 → localStorage（只写不删，见 viewState.applyNodeViewState 注释）
  const vsAll: Record<string, NodeViewState | undefined> = {}
  for (const [id, f] of Object.entries(loaded.nodeFiles)) vsAll[id] = f.viewState
  applyAllViewState(vsAll, extras.kv)

  // ③ 冲突审计（只报不改）
  const viewConflicts: OpenProjectOutcome['viewConflicts'] = []
  for (const [id, f] of Object.entries(loaded.nodeFiles)) {
    viewConflicts.push(...auditViewStateConflicts(id, f.viewState, extras.kv))
  }

  return { ...loaded, tplIdMap, assetProblems, viewConflicts }
}

/* ────────────────────────── 供 UI 直接用的组装辅助 ────────────────────────── */

/**
 * 把 openProjectFull 的结果套用到 tables 上（改写模板绑定 id）。
 *
 * 返回**新对象**，不原地改 —— store 的重渲染靠引用变化，
 * 原地改会被浅比较吃掉，症状是「工程打开了但绑定还显示旧模板名」。
 */
export function applyTplIdMap(
  tables: Record<string, TableSchema>,
  tplIdMap: Record<string, Record<string, string>>,
): Record<string, TableSchema> {
  const map = tplIdMap
  if (Object.keys(map).length === 0) return tables
  const out: Record<string, TableSchema> = {}
  let changed = false
  for (const [id, t] of Object.entries(tables)) {
    const m = map[id]
    if (!m) {
      out[id] = t
      continue
    }
    const next = rewireTemplateIds(t, m)
    if (next !== t) changed = true
    out[id] = next
  }
  return changed ? out : tables
}

/** 便捷重导出：让 UI 只从一个文件 import 常用件 */
export { collectViewStateFromTables, applyAllViewState, saveProject, loadProject }
export type { ProjectFs, ProjectRef, SaveProjectInput, LoadedProject, NodePos, Row }
