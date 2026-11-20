// 工程包IO 层 —— 读写目录式工程（Q1已拍板 = B 目录式）
//
// 分层铁律（与项目既有约束一致）：
//   projectFormat.ts（纯计算：形状 + 派生 + 判定）→ 本文件（IO：只管搬字节）
//   本文件**不碰electron / node:** 的具体 API，一切走 ProjectFs 适配器接口，
//   Electron 与 浏览器各一套实现 —— 与 persistence.ts 的 PersistenceAdapter 同一套路。
//
// ★ 阶段一的范围（docs/18 Q4 已划定，不要越界）：
//   做：目录结构落盘/读取 + project.json 清单 + 节点文件 + 引用索引
//   不做：history/ 的实际写入（阶段二）、data/ 的 embed 打包（等 Q7/Q8）
//        data/ 在本阶段只是**读清单**，不拷贝文件。
//
// ★ 为什么目录式要"一节点一文件"（Q1 拍板理由）：
//   数百节点的工程，git diff 能看出「哪个节点的什么参数改了」。
//   单文件 json 做不到这点 —— 任何改动都是整文件 diff，等于没有diff。

import type { TableSchema } from './schema'
import type { NodePos, Row } from './store'
import {
  PROJECT_FORMAT_VERSION,
  type ProjectManifest,
  type ProjectNodeFile,
  type ProjectEdgeFile,
  type ProjectFormatProblem,
  type NodeViewState,
  type NodeAssetRef,
  validateProject,
  validateNodeFile,
  buildProject,
  buildNodeFile,
  fallbackPosition,
  isReadable,
  formatProblems,
  allEdges,
} from './projectFormat'

/* ────────────────────────── 目录常量 ────────────────────────── */

export const PROJECT_DIRS = {
  nodes: 'nodes',
  history: 'history',
  data: 'data',
  /**
   * 节点资产字节的落点（模板 / 编辑稿 / 导入文件 / 附件）。
   *
   * ★ 为什么要单独一棵树而不是塞进 data/：
   *   data/ 装的是「外部数据源」（价目表、图纸等按pluginId 归类），
   *   语义是**输入**；assets/ 装的是**节点自己挂着的文件**，
   *   语义是「这个节点的一部分」。两者生命周期不同（data/ 阶段一甚至不拷贝字节），
   *   混在一起会导致「删外部数据源时把节点自己的文档也删了」这类误删。
   */
  assets: 'assets',
} as const

export const PROJECT_FILES = {
  manifest: 'project.json',
  edges: 'edges.json',
} as const

/* ────────────────────────── 文件系统适配器 ────────────────────────── */

/**
 * 工程目录的文件系统抽象。
 *
 * 只暴露「列目录 / 读文件 / 写文件 / 建目录」四件事 —— 够用，且不绑任何平台。
 * 为什么不做「删文件」：工程保存是**覆盖式**的，不做增量删除。
 *   真要删，得加「本次工程有哪些文件」的差集计算，而那会引入
 *   「删错文件 = 丢数据」的风险 —— 阶段一不值得。孤儿文件留着无害。
 */
export interface ProjectFs {
  /**
   * 列目录下的**文件名**（不含路径，不含子目录名）。不存在时返回空数组。
   *
   * ★ 返回值里**可能包含子目录名**（实现方无法可靠区分时宁可多给，
   *   由调用方判断）—— 打包时需要靠它递归。
   *   反过来，实现方**不要**把「递归列出所有路径」塞进来：
   *   那会让每个调用方的边界判断失效。
   */
  list(dir: string): Promise<string[]>
  /** 读 UTF-8 文本；文件不存在返回 null */
  readText(path: string): Promise<string | null>
  /**
   * 读二进制字节；文件不存在返回 null。
   *
   * ★ 为什么必须有：节点资产（文档/音频/视频/附件）是二进制的，
   *   readText 会把它们解成 UTF-8 —— 中文文件名下的 docx 会被解出乱码，
   *   且**长度就变了**，属静默损坏。资产必须按字节读写。
   */
  readBytes?(path: string): Promise<Uint8Array | null>
  /**
   * 写二进制字节（自动建父目录）。
   * 与 readBytes 成对：只有读没有写等于只能看不能存。
   */
  writeBytes?(path: string, bytes: Uint8Array): Promise<void>
  /** 写 UTF-8 文本（自动建父目录） */
  writeText(path: string, text: string): Promise<void>
  /** 建目录（幂等） */
  mkdir(dir: string): Promise<void>
  /**
   * 文件是否存在。
   *
   * ★ 刻意设成**可选**：老适配器（已写好的 IndexedDB 实现）没这个方法。
   *   有了它才能校验「资产字节是否跟着工程走」，没有就跳过校验——
   *   不能因为适配器不齐就让整个加载报错，那属于用「更糟」换「更好」。
   */
  exists?(path: string): Promise<boolean>
}

/* ────────────────────────── 工程定位 ────────────────────────── */

/**
 * 工程句柄 —— 一条工程在文件系统里的位置。
 *
 * ★ 阶段一：工程根目录的 id 就是「当前激活页 id」，不引入工程注册表。
 *   将来若要「多工程 / 最近打开列表」，那是 projectStore 的另一层职责，不污染这里。
 */
export interface ProjectRef {
  /** 工程根目录名（在适配器的存储空间里唯一定位） */
  rootDir: string
}

/** 节点文件路径。nodeId 里可能有路径分隔符等危险字符，必须转义 */
export function nodePath(rootDir: string, nodeId: string): string {
  return `${rootDir}/${PROJECT_DIRS.nodes}/${encodeURIComponent(nodeId)}.json`
}

/**
 * 相对路径是否安全（不含越界段）。
 *
 * ★ 这是安全项，不是洁癖：
 *   资产relPath 来自工程文件，而工程文件是可以从别人那儿收来的。
 *   `../../../Windows/System32/...` 这类路径若直接拼上去，
 *   「读资产」就变成了任意文件读取，「写资产」变成任意文件覆盖。
 *
 *   判定规则（逐段看，不做字符串拼接后再检查 —— 拼接后再查会被
 *   `a/../../b` 这类中间带 .. 的形态绕过）：
 *   - 空段（`a//b`）不合法
 *   - `.` 与 `..` 段一律不合法
 *   - 含反斜杠不合法（Windows 上 `\` 也是分隔符，统一按非法处理）
 *   - 含绝对路径前缀（`/` 或 `C:`）不合法
 *
 * @returns 合法返回 true；非法返回 false（调用方应跳过并报 warn）
 */
export function relPathIsSafe(rel: string): boolean {
  if (!rel || rel.length === 0) return false
  if (rel.includes('\\')) return false
  if (rel.startsWith('/')) return false
  if (/^[A-Za-z]:/.test(rel)) return false
  const segs = rel.split('/')
  for (const s of segs) {
    if (s === '' || s === '.' || s === '..') return false
  }
  return true
}

/** 资产路径（工程根内）。relPath 非法时返回 null，由调用方决定怎么报 */
export function assetPath(
  rootDir: string,
  nodeId: string,
  asset: { relPath: string },
  assetDir?: string,
): string | null {
  const dir = assetDir || PROJECT_DIRS.assets
  if (!relPathIsSafe(asset.relPath)) return null
  return `${rootDir}/${dir}/${encodeURIComponent(nodeId)}/${asset.relPath}`
}

/* ────────────────────────── 保存 ────────────────────────── */

/** 保存入参（从 store 的 SaveableState 投影而来） */
export interface SaveProjectInput {
  name: string
  tables: Record<string, TableSchema>
  records: Record<string, Row[]>
  positions: Record<string, NodePos>
  dataSources?: ProjectManifest['dataSources']
  /**
   * 逐节点的控件状态（2026-10-06 新增）。
   *
   * 不给时 buildNodeFile 会沿用旧文件里的值，故老调用点不必改。
   */
  viewState?: Record<string, NodeViewState>
  /**
   * 逐节点的资产登记（2026-10-06 新增）。
   * 不给时沿用旧文件里的值，理由同viewState。
   */
  assets?: Record<string, NodeAssetRef[]>
  /** 复用旧清单以保留 createdAt 与已有指纹 */
  prev?: ProjectManifest
  /**
   * 资产字节的存放目录（相对工程根）。
   *
   * ★ 不给时沿用旧文件里的值，再不给则用 PROJECT_DIRS.assets。
   */
  assetDir?: string
  /**
   * ★ 旧工程的**完整原始节点文件**（2026-10-06 新增）。
   *
   * 直接把 loadProject 返回的 nodeFiles 整个传回来即可。
   * 它承载的是本程序看不懂的部分（未知顶层键 / 未知 schema 字段 / extensions）——
   * 少传这一项，「不丢数据」就只在单进程内成立，
   * 关掉程序再打开另存时那些字段就没了。这也是最容易被漏掉的一环：
   * 同一个会话里连存两次看不出问题，重开程序才发现丢。
   */
  prevNodeFiles?: Record<string, ProjectNodeFile>
}

/**
 * 保存工程到目录。
 *
 * 写序（有意如此）：**先写节点文件，最后写 project.json**。
 *   project.json 是索引 —— 若中途失败，它仍指向旧版本，
 *   下次打开读到的是「上次的完整工程 + 这次的半截节点」，而不是「索引说有N 个节点但只落了2 个」。
 *   反过来写（先写索引）会造出一个宣称有 N 个节点、实际只有 2 个的工程。
 *
 * ★ 这是「崩溃一致性」的最低要求：不追求原子，但要保证**不会宣称成功**。
 */
export async function saveProject(fs: ProjectFs, ref: ProjectRef, input: SaveProjectInput): Promise<ProjectManifest> {
  const root = ref.rootDir

  const manifest = buildProject(
    {
      name: input.name,
      tables: input.tables,
      records: input.records,
      positions: input.positions,
      dataSources: input.dataSources,
    },
    input.prev?.createdAt, // 保留原始创建时间
    input.prev, // ★ 传旧清单：否则工程级未知字段每次保存都被抹掉一次
  )

  // ★ assetDir 优先取调用方给的；没给就用默认值。
  //   buildNodeFile 内部会再沿用旧文件里的 assetDir（它比这里更细，
  //   因为逐节点可能不同），这里只用来决定建哪个目录。
  const assetDir = input.assetDir

  await fs.mkdir(`${root}/${PROJECT_DIRS.nodes}`)
  await fs.mkdir(`${root}/${PROJECT_DIRS.history}`) // 阶段二用，先建好避免以后改结构
  await fs.mkdir(`${root}/${PROJECT_DIRS.data}`) // C5 的落点，阶段一只有清单
  // 资产目录：自定义时建自定义的，否则建默认的
  await fs.mkdir(`${root}/${assetDir || PROJECT_DIRS.assets}`)

  // 逐节点落盘。**保留旧指纹** —— 保存不是重算，指纹该由阶段二的执行器更新。
  //
  // ★ 旧文件的来源有两条，优先用调用方给的 prevNodeFiles：
  //   ① prevNodeFiles：loadProject 的原始回传（含内存里已被改过的内容）
  //   ② 回读磁盘：没传 prevNodeFiles 时的兜底，保证旧的调用点不传也能工作
  //   两条都没有 → 当新节点写。
  for (const id of manifest.nodeIds) {
    let old: Record<string, unknown> | null = null
    if (input.prevNodeFiles && input.prevNodeFiles[id]) {
      old = input.prevNodeFiles[id] as unknown as Record<string, unknown>
    } else {
      const oldText = await fs.readText(nodePath(root, id))
      // ★ 旧文件按 Record 读，不按 ProjectNodeFile 读：
      //   类型上「未知的字段」是不存在的成员，用 ProjectNodeFile 接等于在类型层面
      //   抹掉它们（运行时对象里其实还在，但那就依赖「类型转换不真的拷贝」
      //   这个脆弱前提）。显式按 Record 走，让 buildNodeFile 的未知键收集
      //   在类型层面也成立 —— 「不丢数据」必须硬，不能靠隐式行为。
      old = oldText ? (safeParse(oldText) as Record<string, unknown> | null) : null
    }
    const file = buildNodeFile(
      id,
      {
        name: input.name,
        tables: input.tables,
        records: input.records,
        positions: input.positions,
        // ★ 控件状态与资产登记要显式传进去。漏传的表现很隐蔽：
        //   字段能写出去（buildNodeFile 会沿用旧值），但**用户新拖进来的附件
        //   永远存不进工程** —— 而这种缺陷在开发机上完全看不出来，
        //   因为开发机上旧值恰好是对的。
        viewState: input.viewState,
        assets: input.assets,
        assetDir,
      },
      (old ?? undefined) as ProjectNodeFile | undefined,
    )
    // ★ buildNodeFile 返回 null = 程序和旧文件都没有这个节点（见其注释）。
    //   此时的正确动作是**什么都不做**，让磁盘上的旧文件原样留着。
    //   写成空文件覆盖 = 真的把用户数据抹了，且不可追溯 —— 这是最坏的一种。
    //   之所以能这样，是因为 manifest.nodeIds 来自 tables 的键，
    //   走到这里说明 tables 里有它但 schema 取不到，属于不该发生的中间态；
    //   保留旧文件是失败时损失最小的选择。
    if (!file) continue
    await fs.writeText(nodePath(root, id), JSON.stringify(file, null, 2))
  }

  // 连线表（派生结果，落一份便于「打开工程就知道有哪些连线」）
  const edges: ProjectEdgeFile = { edges: allEdges(input.tables) }
  await fs.writeText(`${root}/${PROJECT_FILES.edges}`, JSON.stringify(edges, null, 2))

  // 索引最后写
  await fs.writeText(`${root}/${PROJECT_FILES.manifest}`, JSON.stringify(manifest, null, 2))

  return manifest
}

/* ────────────────────────── 读取 ────────────────────────── */

/** 读回来的工程内容 */
export interface LoadedProject {
  manifest: ProjectManifest
  tables: Record<string, TableSchema>
  records: Record<string, Row[]>
  positions: Record<string, NodePos>
  /** 逐节点的上游指纹（阶段二用） */
  upstream: Record<string, ProjectNodeFile['upstream']>
  /**
   * ★逐节点的**完整原始文件**（2026-10-06 新增）。
   *
   * 为什么必须回传整份原始对象，而不只是 schema/rows/position：
   *   本程序看不懂的那部分（未知顶层键、未知 schema 字段、extensions）
   *   都不会进 tables/records/positions。若 load 阶段就只挑已知字段出来，
   *   上层拿不到旧文件，save 时也就无从继承 —— 「不丢数据」在跨进程
   *   （关掉程序再打开另存）时必然失效。
   *
   *   契约：把这份对象原样带回 saveProject 的 prev位，
   *   由 buildNodeFile 负责合并。上层**不要**去理解它。
   */
  nodeFiles: Record<string, ProjectNodeFile>
  /** 连线表内容；为 null 表示 edges.json 缺失（老工程） */
  edges: ProjectEdgeFile | null
  /** 校验问题：error 级会阻止打开，warn 级仅提示 */
  problems: ProjectFormatProblem[]
}

/**
 * 读取工程。
 *
 * ★ 遇到损坏节点**不整体失败**，而是记 error 并跳过 —— 但只要有 error 就`ok=false`，
 *   由调用方决定是否拒绝打开。
 *   为什么不全失败：单个节点文件手滑改坏，不该导致几百节点的工程全打不开，
 *   用户连"改回来"的机会都没有。
 */
export async function loadProject(fs: ProjectFs, ref: ProjectRef): Promise<LoadedProject> {
  const root = ref.rootDir
  const problems: ProjectFormatProblem[] = []

  const manifestText = await fs.readText(`${root}/${PROJECT_FILES.manifest}`)
  if (!manifestText) {
    return {
      manifest: emptyManifest(),
      tables: {},
      records: {},
      positions: {},
      upstream: {},
      nodeFiles: {},
      edges: null,
      problems: [{ level: 'error', message: '工程目录里没有 project.json，可能不是有效的工程' }],
    }
  }

  const manifest = safeParse(manifestText) as ProjectManifest | null
  const mProblems = validateProject(manifest)
  problems.push(...mProblems)
  const m = (isReadable(mProblems) ? manifest : emptyManifest()) as ProjectManifest

  const tables: Record<string, TableSchema> = {}
  const records: Record<string, Row[]> = {}
  const positions: Record<string, NodePos> = {}
  const upstream: Record<string, ProjectNodeFile['upstream']> = {}
  const nodeFiles: Record<string, ProjectNodeFile> = {}

  // 索引里的 nodeIds 是权威清单 —— 不扫目录，避免把孤儿文件当成节点读进来
  for (const id of m.nodeIds ?? []) {
    const text = await fs.readText(nodePath(root, id))
    if (!text) {
      problems.push({ level: 'error', message: '清单列出的节点文件缺失', nodeId: id })
      continue
    }
    const file = safeParse(text) as ProjectNodeFile | null
    const np = validateNodeFile(file, id)
    problems.push(...np)
    if (!file || !file.schema) continue
    tables[id] = file.schema
    records[id] = file.rows ?? []
    // ★ 缺坐标必须兜底，不能整项跳过：
    //   跳过 = 该节点在画布上没坐标 → React Flow 给它随便放一个 → 多个节点叠一起，
    //   症状是「节点丢了」。兜底位让它们散开，至少看得到一片节点。
    positions[id] = file.position ?? fallbackPosition(id)
    if (file.upstream) upstream[id] = file.upstream
    // ★ 整份原始文件留着，供再保存时继承未知字段（见 LoadedProject.nodeFiles 注释）
    nodeFiles[id] = file
    // ★ 资产字节是否真的在工程目录里（2026-10-06）。
    //   为什么必须在这里查：ProjectFs 只给 readText，且资产是二进制。
    //   症状链是「工程在、节点在、连线在，但双击打开是空白」——
    //   用户看到的是「我节点里的东西没了」，而真正丢的是字节。
    //   在打开时就报出来，比用户自己发现强得多。
    if (file.assets && Array.isArray(file.assets)) {
      for (const a of file.assets) {
        if (!a || typeof a.relPath !== 'string') continue
        const path = assetPath(root, id, a, file.assetDir)
        if (!path) {
          problems.push({
            level: 'warn',
            message: `资产路径非法或越界，已跳过：${String(a.relPath)}`,
            nodeId: id,
          })
          continue
        }
        const exists = await fs.exists?.(path)
        // 有 exists 时才判缺失 —— 老适配器没实现这个方法时不能瞎报
        if (exists === false) {
          problems.push({
            level: 'warn',
            message: `资产字节缺失（${a.name}）：${path}`,
            nodeId: id,
          })
        }
      }
    }
  }

  const edgesText = await fs.readText(`${root}/${PROJECT_FILES.edges}`)
  const edges = edgesText ? (safeParse(edgesText) as ProjectEdgeFile | null) : null
  if (!edges) {
    // 老工程没这个文件不算错 —— edges 是派生数据，缺失时可从 tables 重算
    problems.push({ level: 'warn', message: '缺 edges.json，连线将在打开后按派生规则重算' })
  }

  return { manifest: m, tables, records, positions, upstream, nodeFiles, edges, problems }
}

/**
 * 读工程的**简版**：只要能不能打开。
 * UI 启动时用它做探测，不做全量解析。
 */
export async function probeProject(fs: ProjectFs, ref: ProjectRef): Promise<ProjectFormatProblem[]> {
  const text = await fs.readText(`${ref.rootDir}/${PROJECT_FILES.manifest}`)
  if (!text) return [{ level: 'error', message: '工程不存在或缺少 project.json' }]
  return validateProject(safeParse(text))
}

/* ────────────────────────── 工具 ────────────────────────── */

function emptyManifest(): ProjectManifest {
  const now = new Date().toISOString()
  return {
    formatVersion: PROJECT_FORMAT_VERSION,
    formatKind: 'cbproj-dir',
    name: '未命名工程',
    createdAt: now,
    updatedAt: now,
    nodeIds: [],
    refs: {},
    dataSources: [],
  }
}

/**
 * JSON.parse 包一层：坏文件不该让整个工程打不开。
 * 返回 null 交给上层的校验去报「error + 跳过」。
 */
function safeParse(text: string): unknown {
  try {
    return JSON.parse(text)
  } catch {
    return null
  }
}

/** 把校验问题转成一句可展示的话 */
export function describeProblems(problems: ProjectFormatProblem[]): string {
  if (problems.length === 0) return ''
  return formatProblems(problems)
}

/* ────────────────────────── 浏览器实现（IndexedDB）────────────────────────── */

/**
 * IndexedDB 版工程存储 —— 浏览器（npm run dev:web）下用。
 *
 * ★ 与 Electron 版并存而非二选一：阶段一先把格式与逻辑跑通，
 *   磁盘版在阶段二做打包/交付时再接（docs/18 Q1 派生要求）。
 *
 * 键设计：path 作为 IndexedDB 的key —— 一条记录就是一个文件。
 *   这样 list(dir) 就是 range 查询，无需维护目录树。
 */
export function createIdbProjectFs(dbName = 'client-base-projects'): ProjectFs {
  const STORE = 'files'
  let dbPromise: Promise<IDBDatabase> | null = null

  function open(): Promise<IDBDatabase> {
    if (dbPromise) return dbPromise
    dbPromise = new Promise((resolve, reject) => {
      const req = indexedDB.open(dbName, 1)
      req.onupgradeneeded = () => {
        if (!req.result.objectStoreNames.contains(STORE)) req.result.createObjectStore(STORE)
      }
      req.onsuccess = () => resolve(req.result)
      req.onerror = () => reject(req.error)
    })
    return dbPromise
  }

  async function tx(mode: IDBTransactionMode): Promise<IDBObjectStore> {
    const db = await open()
    return db.transaction(STORE, mode).objectStore(STORE)
  }

  return {
    async list(dir) {
      const store = await tx('readonly')
      const prefix = dir.endsWith('/') ? dir : `${dir}/`
      return new Promise((resolve, reject) => {
        const out: string[] = []
        // ★ 必须同时返回「本层文件名」与「本层子目录名」。
        //   只返回文件名的话，打包时递归进不去子目录（assets/n1/…），
        //   症状是「打包成功、解包后文档全没了」且无任何报错 ——
        //   成功提示 + 数据丢失，是最坏的组合。
        //   IndexedDB 是扁平 key 存储，列不出目录实体，
        //   所以从 key 的剩余段里判断「还有斜杠 = 下一层」。
        const seen = new Set<string>()
        const req = store.openCursor(IDBKeyRange.bound(prefix, `${prefix}￿`))
        req.onsuccess = () => {
          const c = req.result
          if (!c) return resolve(out)
          const key = String(c.key)
          const rest = key.slice(prefix.length)
          if (rest) {
            const slash = rest.indexOf('/')
            if (slash === -1) {
              const name = decodeURIComponent(rest)
              if (!seen.has(name)) {
                seen.add(name)
                out.push(name)
              }
            } else {
              // 子目录名：去掉尾部斜杠（目录 key 惯例上可能带）
              const sub = decodeURIComponent(rest.slice(0, slash))
              if (sub && !seen.has(sub)) {
                seen.add(sub)
                out.push(sub)
              }
            }
          }
          c.continue()
        }
        req.onerror = () => reject(req.error)
      })
    },
    async readText(path) {
      const store = await tx('readonly')
      return new Promise((resolve, reject) => {
        const req = store.get(path)
        req.onsuccess = () => resolve((req.result as string | undefined) ?? null)
        req.onerror = () => reject(req.error)
      })
    },
    // ★ 资产是二进制，必须按字节存。落库形态是 {b: Uint8Array} 而不是裸 Uint8Array：
    //   裸 Uint8Array 在部分浏览器上会被 structuredClone 正常处理，
    //   但包一层对象后可以在将来加校验字段（如 hash）而不改存储格式。
    async readBytes(path) {
      const store = await tx('readonly')
      return new Promise((resolve, reject) => {
        const req = store.get(path)
        req.onsuccess = () => {
          const v = req.result as { b?: Uint8Array } | undefined
          resolve(v?.b ?? null)
        }
        req.onerror = () => reject(req.error)
      })
    },
    async writeBytes(path, bytes) {
      const store = await tx('readwrite')
      await new Promise<void>((resolve, reject) => {
        const req = store.put({ b: bytes }, path)
        req.onsuccess = () => resolve()
        req.onerror = () => reject(req.error)
      })
    },
    async exists(path) {
      const store = await tx('readonly')
      return new Promise((resolve, reject) => {
        const req = store.getKey(path)
        // ★ 必须用 getKey 而不是 count：count 会对整条range 做计数，
        //   在几十万文件的工程上是全表扫。getKey 只查一个键。
        req.onsuccess = () => resolve(req.result !== undefined)
        req.onerror = () => reject(req.error)
      })
    },
    async writeText(path, text) {
      const store = await tx('readwrite')
      await new Promise<void>((resolve, reject) => {
        const req = store.put(text, path)
        req.onsuccess = () => resolve()
        req.onerror = () => reject(req.error)
      })
    },
    async mkdir() {
      // IndexedDB 无目录概念 —— key 本身就是路径，目录「存在」= 有文件在其下
      return
    },
  }
}
