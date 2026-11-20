// 工程文件格式（.cbproj 目录式）—— 纯计算层，零 IO /零 electron / 零 document
//
// 口径来源：docs/18-域外连接与工程文件格式-需求备忘.md
//   Q1 已拍板（2026-10-04）= B 目录式：工程根/ + nodes/ + edges.json + data/ + history/
//   Q4 已拍板 = 先做阶段一（工程格式），历史记录留阶段二
//
// ★ 本文件只管「数据长什么样 + 怎么互相转换」，不碰磁盘。
//   IO 在 projectStore.ts，Electron/浏览器差异在 persistence.ts 的适配器模式里解决。
//
// ★★ 单一真相源：连线派生（edgesOf）必须与 Canvas.tsx:444 的 useEffect 逐字对齐。
//   画布上的边是**派生**出来的（不单独存），工程格式里也存派生结果。
//   两边各写一份必然漂移 —— 一旦漂移，症状是「保存后重开连线全没了」，
//   而这类问题只在跨进程重开工程时才暴露，本地开发根本看不见。
//   故这里导出唯一的 edgesOf，Canvas.tsx 改为引用它。

import type { TableSchema, FieldDef } from './schema'
import type { NodePos, Row } from './store'

/* ────────────────────────── 格式版本 ────────────────────────── */

/**
 * 工程格式版本号。
 *
 * ★ 为什么必须有：目录式工程会演进（新增 history/、新增 data/ 打包…）。
 *   没有版本号，老工程用新代码打开时只能靠 try-catch 猜 —— 猜错就是静默丢数据。
 *   用户报「我那三个老工程打不开」时，没有版本号就无从定位。
 *
 * 规则：
 *   - 结构**新增**可选字段 → minor+1，老工程能读
 *   - 结构**删除/改名/语义变** → major+1，老工程需迁移（且必须写迁移器）
 *
 * 2026-10-06：1 → 2。新增 nodeVersion / extensions 两个**可选**字段，
 *   且老数据不带也能正常读写，故按minor 处理，老工程无须迁移。
 */
export const PROJECT_FORMAT_VERSION = 2

/** 当前版本的格式名（留作将来 B 目录式 → C zip 式切换时的判据） */
export const PROJECT_FORMAT_KIND = 'cbproj-dir' as const

/* ────────────────────────── 落盘结构 ────────────────────────── */

/**
 * 工程清单 `project.json`。
 *
 * 刻意**不含**节点数据 —— 节点各占一个文件，便于 git diff 看出「哪个节点的什么参数改了」。
 * 这正是 Q1 选 B（目录式）而非 C（zip 单文件）的核心理由。
 */
export interface ProjectManifest {
  formatVersion: number
  formatKind: typeof PROJECT_FORMAT_KIND
  /** 工程显示名 */
  name: string
  /** 创建时间 / 最后修改时间（ISO 字符串，跨平台可读，不受时区二次解释） */
  createdAt: string
  updatedAt: string
  /** 节点 id清单（索引，避免读工程时必须遍历目录） */
  nodeIds: string[]
  /**
   * 引用索引：节点 id → 它引用的上游节点 id 集合。
   * ★ 这是「节点直接引用关系」的正式落点（C3 九要素之一）。
   *   为什么要冗余存一份而不是打开时重算：
   *   ① 引用图是排查「这个值哪来的」的第一入口，重算要遍历全部节点；
   *   ② 阶段二做上游指纹（Q14已拍板 A）时要快速取「我的上游列表」，不能每次全扫。
   *   冗余的代价是可能与 nodes/ 不同步 —— 故collectReferences 与 buildProject 必须同源产出。
   */
  refs: Record<string, string[]>
  /**
   * data/ 清单（C5「一并存储」的占位）。
   * ★ 阶段一只做**读取与约定**，不做实际打包 —— 打包粒度等 Q7（存原文件还是解析后数据）。
   *   每个条目对应 data/<pluginId>/ 下的一个外部数据源。
   */
  dataSources: ProjectDataSource[]
  /**
   * 当前程序不认识的工程级字段，原样保留（与节点文件同思路，2026-10-06 新增）。
   *
   * 例：将来的画布分页信息、协作锁、只读标记 —— 都可能先落在工程上。
   * 与节点 extensions 一样：只保证数据不丢，不保证本程序会用它。
   */
  extensions?: Record<string, unknown>
}

/** 一个外部数据源的登记条目（阶段一只有登记，没有实际拷贝） */
export interface ProjectDataSource {
  /** 数据输入型插件 id（Q11 已拍板 C：先内置 web/api/file 三种，接口按插件设计） */
  pluginId: string
  /** 原始地址：本地绝对路径 / URL / 接口地址 */
  address: string
  /**
   * 是否要一并存储（C5双策略）。
   * ★ 阶段一**恒为 false** —— 打包动作等 Q7/Q8 拍板后才有意义。
   *   字段先占位，是为了让阶段一的格式不再改（格式一改，老工程全废）。
   */
  embed: boolean
  /** 指向 data/ 下的相对路径；embed=false 时为 undefined（只有地址） */
  relPath?: string
  /** 上次成功读取的时间（ISO）—— C2 快照的时间戳，便于判断数据新鲜度 */
  fetchedAt?: string
}

/**
 * 单节点文件 `nodes/<id>.json`。
 *
 * 为什么要把坐标（position）与schema 放一起：
 *   C3 九要素里「节点坐标」原盘点为「React Flow 内部持有，未进数据模型」——
 *   实测**已修正**：store.ts 的 SaveableState.positions 一直在存坐标。
 *   所以坐标不需要"抽取"，只是要落进工程文件而已。
 */
export interface ProjectNodeFile {
  /** 节点 id（与文件名一致；显式存一份便于校验，防文件名被手工改坏） */
  id: string
  /** = tables[id] 原样落盘。schema/config 一字不改，避免「工程里的表和画布上的表不是一回事」 */
  schema: TableSchema
  /** 节点坐标。缺失时给默认位（新节点从 0,0 起会重叠，故给个按 id 散开的兜底） */
  position: NodePos
  /** 本节点的全部行数据（C3「节点参数数据」：现在与 schema 分离，纳入工程文件版本管理） */
  rows: Row[]
  /**
   * 上游指纹（Q14 已拍板 A）。
   * ★ 阶段一只定义结构与判定纯函数，**不写入** —— 写入要等阶段二的执行器。
   *   阶段一先把字段占好：格式字段后加 = 老工程又要迁移一次。
   */
  upstream?: UpstreamFingerprint[]

  /* ── 向前兼容承载位（2026-10-06新增，见下方长注释）───────────── */

  /**
   * **本节点自己的**格式版本号。
   *
   * ★ 为什么工程级 formatVersion 不够用：
   *   formatVersion 管的是「整个工程文件的结构」（目录长什么样、清单有哪些字段）。
   *   而真正会持续演进的是**节点内部**—— oosheet 的绑定、章节工作簿的层级、
   *   将来新节点类型自己的 config。这些演进不该拖着全工程文件一起升版本，
   *   否则每加一个节点参数就得让所有老工程迁移一次。
   *   所以每个节点文件自带版本号，节点内部演进只动这一层。
   *
   *   当前所有节点的 nodeVersion 都是 undefined（等价于 1），字段先占位。
   */
  nodeVersion?: number

  /**
   * 当前程序**不认识**的字段，原样存这里。
   *
   * ★ 这是「节点还没做完」这条需求的核心答案：
   *   用户明确说了节点类型和节点内部结构都还在演进。那么必然会有一段时间，
   *   工程文件里存在本程序看不懂的东西 —— 可能是未来版本的节点，可能是
   *   别的分支写的字段，也可能是用户手工加的备注。
   *   **这类数据绝不能被静默丢弃**。静默丢弃是这类软件最恶心的bug：
   *   用户不会立刻发现，等到需要那份数据时才发现三个月前保存时就没了，
   *   而且无法追溯。
   *
   *   落地方式（buildNodeFile 里做，不用改调用方）：
   *   重新保存时，把旧文件里所有**本程序不认识的顶层键**搬进extensions。
   *   于是：不认识的东西至少能原样存回来，往返一次不丢。
   *
   *   注意这不等于「能正确渲染」—— 本程序读不懂就是读不懂，
   *   extensions 只保证数据不丢，不保证语义可执行。
   */
  extensions?: Record<string, unknown>

  /* ── 控件状态（用户 2026-10-06 第 5 项：「控件的需求」）─────────────
     ★ 为什么必须在工程文件里，而不能继续放在 localStorage：
       本程序此前的做法是每个控件各存一个 localStorage 键：
         ne-tab:<nodeId>      NodeEditor 当前标签页
         ooref:<nodeId>       参数引用面板开合
         wb-sheet:<nodeId>    工作簿当前工作表
         refpanel:<nodeId>    右侧参数面板开合
         rightpanel:<nodeId>  右面板开合
         wrp-refpanel:<id>    Word 面板开合
         trigrules:<nodeId>   触发规则
       那样做的后果是**这套状态不属于工程**：
         · 换台机器 / 换个浏览器 → 全部控件回到初始状态
         · 工程发给同事 → 对方看到的是另一套面板布局
         · 同一份工程在两台机器上反复来回 → 面板状态互相覆盖
       控件状态是「这个节点长什么样、当前停在哪」的组成部分，
       它和坐标、字段一样属于节点的身份，不属于「这个用户」。故随工程走。

     刻意**不**把它叫 UI/prefs：那类字段（主题、窗口大小）才是用户级的。
       这里是节点级的，故叫 viewState 且挂在节点文件上。
     ────────────────────────────────────────────────────────────────── */

  /**
   * 节点内的控件状态。
   *
   * ★ 全是可选的：一个键都不给也能正常打开（回落到各自的默认值）。
   *   这一点是必须的 —— 老工程没有这个字段，不能因此打不开。
   */
  viewState?: NodeViewState

  /* ── 节点资产（用户 2026-10-06 第 4 项：「节点文档、音频、视频、文件」）──
     ★ 为什么资产必须进工程文件（这是本轮最大的隐患修复）：
       节点上会挂实际文件 —— ooword/oosheet 的模板与编辑稿、导入的表格、
       将来节点上的音频/视频/附件。本程序此前把字节全放在 IndexedDB
       （templateStore / editStore），工程文件里**一个字都没记**。
       后果很直接：工程文件发给同事，对方拿到的是「一堆空节点」——
       节点在、字段在、连线在，双击打开却是空白，因为字节没跟过来。
       而且这个故障在开发机上永远看不见（自己导出的工程自己打得开）。
     ────────────────────────────────────────────────────────────────── */

  /**
   * 本节点挂着的资产登记（**字节本体落在工程目录的 assets/ 下**）。
   *
   * ★ 只登记「指向哪」，不在 JSON 里塞字节 —— 文档几百 KB 起，
   *   base64 进 JSON 会让节点文件变成几十MB，git diff 完全没法看，
   *   恰好毁掉 Q1 选目录式的全部理由。
   */
  assets?: NodeAssetRef[]

  /**
   * 资产字节的存放目录（相对工程根）。
   *
   * 默认 `assets`，单独留字段是为了将来支持「一个工程带多份文档副本」
   * 时不必改结构。存**相对**路径而非绝对路径：绝对路径换台机器就失效，
   * 那样等于白存。
   */
  assetDir?: string
}

/**
 * 一份资产的登记条目。
 *
 * 刻意分「登记」与「字节」两层（见 ProjectNodeFile.assets 注释）：
 * 这里只说「有这个文件、多大、什么类型、叫什么」，字节落在工程目录里。
 */
export interface NodeAssetRef {
  /**
   * 资产 id。
   *
   * ★ 为什么不用文件名当 id：文件名会改（用户重命名）、会冲突
   *   （两个节点都叫「方案.docx」）。用稳定 id，字节存
   *   `<assetDir>/<nodeId>/<assetId>.<ext>`，改名字不影响字节定位。
   */
  id: string
  /** 字节在工程目录里的相对路径（含扩展名） */
  relPath: string
  /** 原始文件名（给人看、给「另存为」用） */
  name: string
  /**
   * 资产用途（决定它怎么参与节点行为）。
   *
   * - template    绑定模板（ooword/oosheet 的底稿）
   * - edited      就地编辑稿（有它则打开时提示「编辑稿模式」）
   * - imported    导入的外部数据（param 节点的导入文件）
   * - attachment  节点附件（音频/视频/图片/文档等）
   */
  role: 'template' | 'edited' | 'imported' | 'attachment'
  /**
   * MIME 类型。
   *
   * ★ 为什么要显式记而不是靠扩展名猜：
   *   .xlsx 到底是 OOXML 还是老的 BIFF，MIME 才准；
   *   .doc 也可能是 Word 97 或 RTF。猜错会让打开路径走错分支，
   *   症状是「双击打开提示格式不对」，用户以为是文件坏了。
   */
  mime: string
  /** 字节数；用于容量显示与「文件是否被截断」的自检 */
  size: number
  /** 最后修改时间（ISO） */
  updatedAt?: string
  /**
   * ★ 字节指纹（见 fingerprintOf）。
   *
   * 用途：打包 / 复制后校验字节有没有被改坏。
   * 资产是二进制，JSON 校验对它无能为力 —— 唯一能验的就是哈希。
   * 没有它，一个损坏的工程只能靠用户打开时逐个发现。
   */
  hash?: string
}

/**
 * 节点内的控件状态（随工程走，见 ProjectNodeFile.viewState 的长注释）。
 *
 * ★ 只收录「**换台机器就该跟着走**」的控件状态。
 *   不收主题、窗口大小、最近打开列表 —— 那些是用户偏好，属于人而不属于工程，
 *   混进工程文件会让工程在不同人的机器上互相覆盖偏好。
 */
export interface NodeViewState {
  /** NodeEditor 当前打开的标签页 */
  tab?: string
  /** 参数引用面板是否展开 */
  refPanelOpen?: boolean
  /** 右侧面板是否展开 */
  rightPanelOpen?: boolean
  /** 工作簿当前工作表名（xlsx/oosheet） */
  activeSheet?: string
  /** 触发规则面板的规则（xlsx TriggerPanel） */
  triggerRules?: unknown
  /**
   * 画布上的折叠态。
   *
   * ★ 为什么要把它从 localStorage 搬进节点文件：
   *   折叠态是「这个节点在画布上是什么样」的一部分。用户整理好一套
   *   「主流程收起来、细节展开」的布局发给同事，对方全展开，
   *   几百节点立刻变成一张看不懂的网 —— 这不是小事。
   */
  collapsed?: boolean
}

/**
 * 上游指纹：记下「我上次重算时，各上游的值/版本是什么」。
 *
 * Q14 已拍板选 A（上游指纹）而非 B（版本号），核心理由：
 *   要能点开看「**哪个**上游变了、什么时候变的」，版本号只能回答「脏了」。
 */
export interface UpstreamFingerprint {
  /** 上游节点 id */
  nodeId: string
  /** 上游当时的值指纹（哈希或原值，见 projectFormat.fingerprintOf 的口径说明） */
  valueHash: string
  /** 本节点上次重算完成的时间（ISO）—— 用于「什么时候算的」 */
  computedAt: string
}

/* ────────────────────────── 连线表 ────────────────────────── */

/**
 * 连线 `edges.json`。
 *
 * ★ 连线在画布上是**派生**的（Canvas.tsx:444），不单独存 —— 但工程格式里要落一份。
 *   理由：落一份才能做到「打开工程不看代码也知道这个工程有哪些连线」，
 *   且Q14 的引用索引、阶段二的执行范围判断都要用。
 *   一致性由 edgesOf() 单一真相源保证。
 */
export interface ProjectEdgeFile {
  edges: ProjectEdge[]
}

export interface ProjectEdge {
  /** 边 id。两种格式，与 Canvas 保持一致：`flow:<src>-><tgt>` / `<src>.<field>-><tgt>` */
  id: string
  source: string
  target: string
  /** 'flow' = 数据流（linkedInputs）；'assoc' = lookup 字段关联 */
  kind: 'flow' | 'assoc'
  /** assoc 边的字段 id（lookup 关联用），flow 边为 undefined */
  fieldId?: string
}

/* ────────────────────────── 引用派生（单一真相源）────────────── */

/**
 * 派生一个节点的连线 —— **与 Canvas.tsx:444 的 useEffect 逐字对齐**。
 *
 * ★★ 改这里必须同步改 Canvas.tsx（它改为引用本函数）。
 *   不对齐的症状：工程保存后重开，连线全没了，或 lookup 边凭空多出来。
 *
 * @param tables 全量表（需要互相查表，故必须全量传）
 * @returns 该节点发出的边。注意是「以本节点为 source/target 之一」的边，
 *          lookup 边方向是 字段所在节点 → 被查表，与 flow 边方向一致。
 */
export function edgesOf(t: TableSchema, tables: Record<string, TableSchema>): ProjectEdge[] {
  const out: ProjectEdge[] = []
  // 软删除（回收站）的节点不产生边 —— 与画布一致，避免落盘出悬空连线
  if (t.deleted) return out

  // ① lookup 字段关联边：t 的某个 lookup 字段 → 被关联的表
  for (const f of t.fields) {
    if (f.type !== 'lookup' || !f.lookup) continue
    const tgt = tables[f.lookup.table]
    // 目标已软删除 → 跳过，与画布一致
    if (!tgt || tgt.deleted) continue
    out.push({
      id: `${t.id}.${f.id}->${f.lookup.table}`,
      source: t.id,
      target: f.lookup.table,
      kind: 'assoc',
      fieldId: f.id,
    })
  }

  // ② 数据流边：linkedInputs 里的每个上游 → 本节点
  for (const src of t.config?.linkedInputs ?? []) {
    const srcT = tables[src]
    if (!srcT || srcT.deleted) continue
    out.push({
      id: `flow:${src}->${t.id}`,
      source: src,
      target: t.id,
      kind: 'flow',
    })
  }

  return out
}

/** 全工程所有连线（按节点 id 字典序排，保证落盘稳定可比对） */
export function allEdges(tables: Record<string, TableSchema>): ProjectEdge[] {
  const out: ProjectEdge[] = []
  for (const id of Object.keys(tables).sort()) {
    out.push(...edgesOf(tables[id], tables))
  }
  return out
}

/**
 * 收集一个节点的直接引用（关联到的其它节点 id 集合）—— refs 索引的来源。
 *
 * ★ 语义澄清（这里最容易搞混）：本函数返回的是「**双向关联**」，不是「上游」。
 *   - flow 边：src→tgt。对 tgt 来说 src 是上游；对 src 来说 tgt 是**下游**。
 *   - assoc 边：字段所在表→被查表。对字段所在表来说被查表是它去查的表（下游）；
 *     对被查表来说字段所在表是**引用它的**（上游）。
 *   两个方向都收，因为「节点直接引用关系」（C3 九要素之一）关心的是**有关系**，
 *   方向判定留给上层（要看上游就按边的 kind 与朝向过滤）。
 *
 * ★ 为什么必须用 allEdges 而不能只看本节点派生的边：
 *   assoc 边是由**字段所在表**派生的，被查表自己派不出这条边——
 *   只看自己就永远查不到「谁在查我」。
 *   实测踩过：refs[price] 恒为空，症状是「价目表节点从来看不到谁引用了它」。
 *
 * ★ 范围：只收Canvas 能画成连线的那两类（linkedInputs 数据流 + lookup 关联）。
 *   `config.inputs` **不收** —— 它是「面板手动勾选的数据来源，无连线但仍汇入」
 *   （见 store.ts NodeConfig.inputs 注释）。把它算进引用会让引用图虚高，
 *   而 Q14 的上游指纹是按引用图算的，虚高会导致大量假「脏」标记。
 */
export function collectReferences(
  t: TableSchema,
  tables: Record<string, TableSchema>,
): string[] {
  if (t.deleted) return []
  const set = new Set<string>()
  for (const e of allEdges(tables)) {
    if (e.source === t.id) set.add(e.target)
    if (e.target === t.id) set.add(e.source)
  }
  return Array.from(set).sort()
}

/** 全工程 refs 索引 */
export function buildRefsIndex(tables: Record<string, TableSchema>): Record<string, string[]> {
  // ★ allEdges 只算一遍然后复用，不在循环里重复调 collectReferences：
  //   那样每节点都重扫全表 → O(n²)，数百节点的工程上会明显卡。
  const edges = allEdges(tables)
  const refs: Record<string, string[]> = {}
  for (const id of Object.keys(tables).sort()) {
    if (tables[id].deleted) {
      refs[id] = []
      continue
    }
    const set = new Set<string>()
    for (const e of edges) {
      if (e.source === id) set.add(e.target)
      if (e.target === id) set.add(e.source)
    }
    refs[id] = Array.from(set).sort()
  }
  return refs
}

/* ────────────────────────── 指纹（Q14 的判定口径）─────────────── */

/**
 * 算一个值的指纹。
 *
 * Q14 已拍板选「上游指纹」，但**存原值还是存哈希**未定（docs/18 已记录此细节）。
 * 本阶段先按**哈希**实现，理由：
 *   - 原值可能是几 MB 的行数据，进指纹就等于把数据复制一遍
 *   - 哈希能塞进 project.json 的索引里，Q14 要的「快速取上游列表」才成立
 * 代价：点开看「哪个上游变了」时，值本身要回上游节点查（只多一步反查，可接受）。
 *
 * ★ 用 JSON.stringify + FNV-1a：必须**键排序**，否则 {a:1,b:2} 与 {b:2,a:1}
 *   会算出不同哈希 —— 那种假「脏」标记比不标记更糟。
 */
export function fingerprintOf(value: unknown): string {
  return fnv1a(stableStringify(value))
}

/** 稳定序列化：对象键排序，保证同内容同哈希 */
export function stableStringify(v: unknown): string {
  if (v === null || typeof v !== 'object') return JSON.stringify(v) ?? 'null'
  if (Array.isArray(v)) return '[' + v.map(stableStringify).join(',') + ']'
  const keys = Object.keys(v as Record<string, unknown>).sort()
  return (
    '{' +
    keys
      .map((k) => JSON.stringify(k) + ':' + stableStringify((v as Record<string, unknown>)[k]))
      .join(',') +
    '}'
  )
}

/** FNV-1a 32 位哈希（够用：只判「变没变」，不抗攻击） */
function fnv1a(s: string): string {
  let h = 0x811c9dc5
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i)
    // h *= 16777619，用位运算避免浮点精度问题
    h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0
  }
  return h.toString(16).padStart(8, '0')
}

/* ────────────────────────── 上游指纹判定（Q14/C10）────────────── */

/**
 * 判断一个节点的下游是否「过期」（C10 的标记判定）。
 *
 * 触发条件（docs/18 C10）：某节点已更新，而它的**直接下游**在本次执行中没有被更新
 * → 下游出现标记。
 *
 * @param currentNodeId 正在被判定的节点（= 下游）。**必传**：下面的「本次已重算」检查的是**它**
 *   自己，不是它的上游 —— 上游重算本来就是这里要检测的事件，拿它去短路会漏判。
 * @param fingerprints 该节点记录的上游指纹
 * @param upstreamNow  上游节点**当前**的指纹（按 nodeId 索引）
 * @param updatedThisRun 本次执行中实际更新过的节点 id 集合
 * @returns 有过期的上游 id 列表；空数组 = 不该标记
 *
 * ★ 为什么要传 updatedThisRun：光比对指纹只能发现「上游变了」，
 *   但如果本节点自己在本次执行里已经重算过了，它就是最新的 —— 不该再标记。
 *   这两个条件缺一不可。
 *
 * ★ 这条曾经写错过：第一版判`updatedThisRun.has(f.nodeId)`（查上游），
 *   症状是「本节点刚重算完却仍被标脏」—— 一个自相矛盾的标记，比没标记更误导人。
 */
export function staleUpstreamsOf(
  currentNodeId: string,
  fingerprints: UpstreamFingerprint[] | undefined,
  upstreamNow: Record<string, string>,
  updatedThisRun: ReadonlySet<string>,
): string[] {
  // 本节点本次已重算 → 它已是最新，不该有标记
  if (updatedThisRun.has(currentNodeId)) return []
  if (!fingerprints || fingerprints.length === 0) return []
  const stale: string[] = []
  for (const f of fingerprints) {
    const now = upstreamNow[f.nodeId]
    // 上游已不存在（节点被删）：也算过期，且必须报出来 —— 这是最容易被忽略的一种断链
    if (now === undefined || now !== f.valueHash) stale.push(f.nodeId)
  }
  return stale.sort()
}

/** 节点是否该打「输入已过期」标记 */
export function isStaleNode(
  node: ProjectNodeFile,
  upstreamNow: Record<string, string>,
  updatedThisRun: ReadonlySet<string>,
): boolean {
  return staleUpstreamsOf(node.id, node.upstream, upstreamNow, updatedThisRun).length > 0
}

/* ────────────────────────── 组装 / 校验 ────────────────────────── */

/** 工程状态（从 store 的 SaveableState 投影而来） */
export interface ProjectSource {
  name: string
  tables: Record<string, TableSchema>
  records: Record<string, Row[]>
  positions: Record<string, NodePos>
  /** data/ 清单；阶段一从 store 的约定位置取（暂由调用方给） */
  dataSources?: ProjectDataSource[]
  /**
   * 逐节点的控件状态（2026-10-06 新增）。
   * 不给时 buildNodeFile 会沿用旧文件里的值，故老调用点不必改。
   */
  viewState?: Record<string, NodeViewState>
  /**
   * 逐节点的资产登记（2026-10-06 新增）。
   * 不给时沿用旧文件里的值，理由同viewState。
   */
  assets?: Record<string, NodeAssetRef[]>
  /** 资产目录（相对工程根） */
  assetDir?: string
}

/** 格式校验结果 */
export interface ProjectFormatProblem {
  level: 'error' | 'warn'
  message: string
  nodeId?: string
}

/**
 * 校验一份工程是否可读。
 *
 * ★ 分级而不是一票否决：老工程缺 history/（阶段二才有）该是 warn 不是 error，
 *   不该因为「字段还没长出来」就打不开。error 只给真正会丢数据的。
 */
export function validateProject(manifest: unknown): ProjectFormatProblem[] {
  const problems: ProjectFormatProblem[] = []
  if (!manifest || typeof manifest !== 'object') {
    return [{ level: 'error', message: 'project.json 不是对象，文件可能损坏' }]
  }
  const m = manifest as Partial<ProjectManifest>

  if (typeof m.formatVersion !== 'number') {
    problems.push({ level: 'error', message: '缺 formatVersion，无法判断工程版本' })
  } else if (m.formatVersion > PROJECT_FORMAT_VERSION) {
    problems.push({
      level: 'error',
      message: `工程版本 ${m.formatVersion} 高于当前程序支持的 ${PROJECT_FORMAT_VERSION}，请升级软件`,
    })
  }

  if (m.formatKind !== PROJECT_FORMAT_KIND) {
    problems.push({
      level: 'warn',
      message: `格式类型为 ${String(m.formatKind)}，当前按 ${PROJECT_FORMAT_KIND} 解析`,
    })
  }
  if (!Array.isArray(m.nodeIds)) {
    problems.push({ level: 'error', message: '缺 nodeIds，无法定位节点文件' })
  }
  if (!m.refs || typeof m.refs !== 'object') {
    problems.push({ level: 'warn', message: '缺 refs 引用索引，打开后可重建' })
  }
  return problems
}

/**
 * 组装工程内容（内存态，不落盘）。
 *
 * 单一真相源：nodes/ 与 refs 由同一份 tables 派生，
 * 保证「索引说的引用」与「节点文件实际连的线」永远一致。
 *
 * @param createdAt 复用旧清单的创建时间
 * @param prev 旧清单；用于把本程序不认识的工程级字段搬进 extensions。
 *   ★ 必须传 —— 不传的话旧工程里的未知字段每次保存都被抹掉一次，
 *   而「保存」通常是用户认为安全的操作，抹数据在这里最不可接受。
 */
export function buildProject(
  src: ProjectSource,
  createdAt?: string,
  prev?: ProjectManifest,
): ProjectManifest {
  const now = new Date().toISOString()
  const nodeIds = Object.keys(src.tables).sort()
  const manifest: ProjectManifest = {
    formatVersion: PROJECT_FORMAT_VERSION,
    formatKind: PROJECT_FORMAT_KIND,
    name: src.name,
    createdAt: createdAt ?? prev?.createdAt ?? now,
    updatedAt: now,
    nodeIds,
    refs: buildRefsIndex(src.tables),
    dataSources: src.dataSources ?? prev?.dataSources ?? [],
  }
  const ext = collectUnknownKeys(prev as Record<string, unknown> | undefined, KNOWN_MANIFEST_KEYS)
  if (ext) manifest.extensions = ext
  return manifest
}

/**
 * 节点文件里「本程序认识的」顶层键。
 *
 * ★ 这份名单是向前兼容的唯一判据 —— 改它等于改「本程序认识什么」。
 *   buildNodeFile 用它把旧文件里不在名单上的键搬进 extensions。
 *   名单漏了某个已知字段 → 那个字段会被搬进 extensions（数据不丢，但位置变了，
 *   下次读的时候要能从 extensions 里捞回来，故下面的搬运逻辑要双向处理）。
 *   名单多写了某个未知字段 → 那个字段会被静默丢弃，这是最危险的错法。
 *
 * 维护规则：**新增顶层字段必须同步加进这份名单**（有测试守着，见 projectFormat.test.ts）。
 */
const KNOWN_NODE_KEYS: ReadonlySet<string> = new Set([
  'id',
  'schema',
  'position',
  'rows',
  'upstream',
  'nodeVersion',
  'extensions',
  'viewState',
  'assets',
  'assetDir',
])

/** 同样适用于工程清单 —— 未知键也要能透传 */
const KNOWN_MANIFEST_KEYS: ReadonlySet<string> = new Set([
  'formatVersion',
  'formatKind',
  'name',
  'createdAt',
  'updatedAt',
  'nodeIds',
  'refs',
  'dataSources',
  'extensions',
])

/**
 * 把「旧文件里本程序不认识的顶层键」收集成一个 extensions 对象。
 *
 * ★ 为什么单独一个函数而不是内联：这段逻辑有两个来源必须合并 ——
 *   ① 旧文件的**顶层**未知键（比如未来版本加的 layoutHint / lockedBy）
 *   ② 旧文件里**已有的** extensions（可能被上层塞过东西）
 *   漏掉任何一个都是丢数据。两处都要搬，且**旧的 extensions 优先级更低**
 *   —— 因为 extensions 是「已经归过一次档」的内容，显式顶层键更可能是本程序
 *   将来会正式支持的字段，让它覆盖归档值更安全。
 */
function collectUnknownKeys(
  old: Record<string, unknown> | undefined,
  known: ReadonlySet<string>,
): Record<string, unknown> | undefined {
  if (!old) return undefined
  const out: Record<string, unknown> = {}
  // 先铺旧 extensions（低优先级）
  const prevExt = old.extensions
  if (prevExt && typeof prevExt === 'object' && !Array.isArray(prevExt)) {
    for (const [k, v] of Object.entries(prevExt as Record<string, unknown>)) out[k] = v
  }
  // 再铺顶层未知键（高优先级，覆盖上面的同名字段）
  for (const [k, v] of Object.entries(old)) {
    if (known.has(k)) continue
    if (k === 'extensions') continue
    out[k] = v
  }
  return Object.keys(out).length > 0 ? out : undefined
}

/**
 * ★ schema 内的未知字段保留。
 *
 * 为什么单独处理：schema 是 `src.tables[id]` 的直接引用，程序不认识的新节点类型
 * 会带着一堆新字段整个进来，**原样落盘时不会丢**。
 * 但危险在于另一头：`loadProject` 把 schema 塞进 tables，
 * 而 store 在跑过一轮计算/规范化后可能把不认识的字段抹掉，
 * 下一次 saveProject 再存 —— 丢在这里。
 *
 * ⚠⚠ 关键约束：**本程序"不认识某个键" ≠ "该键应该被保留"**。
 *   这两件事完全不同，混淆它们会造成比丢数据更坏的后果：
 *   程序**主动删掉**的字段（比如迁移到 v2 后废弃的 legacyV1）
 *   也会被当成「不认识」而复活 —— 于是用户改过的配置会一次次自己长回来，
 *   且症状极其隐蔽（配置删不掉），比丢数据更难排查。
 *
 *   所以判定必须是「**程序本轮产出的 schema 是不是一个完整认识的状态**」：
 *   - 程序认知里**这个节点存在**（tables 里有它）→ 程序是权威，按它落盘，
 *     只补「程序完全没有产生过的键」。
 *   - 程序认知里**这个节点不存在**（tables 里没有它）→ 程序根本不认识它，
 *     此时旧 schema 就是唯一真相，整份沿用。
 *
 *   分不清这两种情况就会出上面那个「删不掉的字段」问题，故这里显式区分。
 *
 * 刻意**不**递归深层：`fields` 数组是本程序一定会遍历的结构，
 * 往里塞不认识的字段反而可能干扰判断。保真优先，但不越权改语义。
 */
function mergeUnknownSchemaKeys(
  next: TableSchema | undefined,
  old: TableSchema | undefined,
): TableSchema | undefined {
  // 程序不认识这个节点 → 旧 schema 是唯一真相，整份沿用
  if (!next) return old
  // 程序认识它 → 程序本轮产出的是权威结果
  if (!old || typeof old !== 'object') return next
  // ★ 经 unknown 中转再转 Record：TableSchema 没有索引签名，
  //   直接 as Record 会报错（TS2352）。这里刻意走unknown，
  //   因为本函数的**全部目的**就是绕过类型系统去看运行时多出来的键 ——
  //   类型系统此刻正是我们要无视的东西。
  const nextRec = next as unknown as Record<string, unknown>
  const oldRec = old as unknown as Record<string, unknown>
  // ★ 关键：以**程序已知的表结构键集合**为基准，而不是「next 里出现过的键」。
  //   next 这次可能恰好没带某个已知字段（比如 store 还没填 defaultValue），
  //   若用 nextSet 做基准就会把旧值当未知字段塞回去 —— 又是「删不掉的字段」。
  const known = KNOWN_SCHEMA_KEYS
  const merged: Record<string, unknown> = { ...nextRec }
  for (const [k, v] of Object.entries(oldRec)) {
    // 程序已知的键：一律以程序本轮的值为准，不看旧的
    if (known.has(k)) continue
    // 程序本轮已经产出的键：同样以本轮为准
    if (k in nextRec) continue
    merged[k] = v
  }
  return merged as unknown as TableSchema
}

/**
 * schema 里「本程序认识的」顶层键。
 *
 * ★ 与 KNOWN_NODE_KEYS 同源的维护要求：新增 schema 顶层字段必须同步加进来。
 *   漏加的后果：程序新增的字段会被当成「未知键」，于是**旧文件里的旧值**
 *   会因为「next 本轮恰好没带它」而覆盖新值 —— 新字段永远更新不了。
 *   有测试守着（projectFormat.test.ts 会核对）。
 */
const KNOWN_SCHEMA_KEYS: ReadonlySet<string> = new Set([
  'id',
  'name',
  'kind',
  'fields',
  'config',
  'deleted',
  // ★ trashMeta 必须在内：它是「回收站」的元数据。
  //   漏了它 → 程序把节点从回收站彻底删除后，旧文件里的 trashMeta 会被当
  //   「未知键」塞回去 → 节点复活成「仍在回收站」，用户删了又出现，
  //   且反复删除都删不掉。这与legacyV1 是同一类病，故列在这里。
  'trashMeta',
])

/**
 * 把一个节点组装成落盘文件。
 *
 * ★ 本函数是「保存时唯一会丢数据的地方」，因此对**四类来源**逐一分派，
 *   每一类都要明确「谁的优先级更高」，不能靠一串 `??` 了事：
 *
 *   schema:   程序本轮的产出（权威）→ 旧文件里程序不认识的键（保底）
 *   position: store 坐标 → 旧文件坐标 → 按 id 散开的兜底位
 *   rows:     store 行数据 → 旧文件行数据 → 空
 *   未知键:   一律搬进 extensions（数据不丢，但本程序不用它）
 *
 *   ⚠ 最容易写错的是把「程序不认识」当成「该保留」。
 *     程序主动删掉的字段（已迁移的 legacy 字段、已清空的 trashMeta）
 *     也会被认成「不认识」而复活 —— 症状是用户删了删不掉的配置，
 *     比丢数据更难排查。判定依据见 mergeUnknownSchemaKeys 的注释。
 *
 * @returns 节点文件；程序与旧文件**都没有**该节点时返回 null（见下）
 */
export function buildNodeFile(
  id: string,
  src: ProjectSource,
  existing?: ProjectNodeFile,
): ProjectNodeFile | null {
  const fresh = src.tables[id]
  const schema = mergeUnknownSchemaKeys(fresh, existing?.schema)
  const pos = src.positions[id]

  // ★ 程序和旧文件都没有这个节点 → 无法组装出合法节点文件。
  //   返回 null 而不是硬造一个空 schema：造出来的文件能通过校验（看着完全正常），
  //   但重新打开时是一个没有名字、没有字段的空节点 —— 那种「静默劣化」比报错难查得多，
  //   而且不可追溯（用户只会觉得「我的节点怎么内容空了」）。
  //   调用方（saveProject）遇到 null 应当**跳过该节点、保留磁盘上的旧文件**，
  //   绝不能拿空文件覆盖 —— 覆盖才是真正丢数据。
  if (!schema) return null

  const file: ProjectNodeFile = {
    id,
    schema,
    position: pos ?? existing?.position ?? fallbackPosition(id),
    rows: src.records[id] ?? existing?.rows ?? [],
    // 已有指纹原样保留：组装不等于重算，重算该由阶段二的执行器做
    upstream: existing?.upstream,
  }
  if (existing?.nodeVersion !== undefined) file.nodeVersion = existing.nodeVersion
  // ★ viewState：程序本轮没产出就沿用旧的（控件状态本程序当前还没接store，
  //   但字段先占好 —— 与 upstream 同样的理由：字段后加 = 老工程又要迁移一次）。
  //   一旦上层开始产出 viewState，本行的优先级会由src 传入（见下方注释）。
  if (src.viewState?.[id]) file.viewState = src.viewState[id]
  else if (existing?.viewState) file.viewState = existing.viewState
  // ★ assets 同理：登记信息必须跟着工程走，否则节点在、字节不在。
  if (src.assets?.[id]) file.assets = src.assets[id]
  else if (existing?.assets) file.assets = existing.assets
  if (src.assetDir) file.assetDir = src.assetDir
  else if (existing?.assetDir) file.assetDir = existing.assetDir

  const ext = collectUnknownKeys(
    existing as Record<string, unknown> | undefined,
    KNOWN_NODE_KEYS,
  )
  if (ext) file.extensions = ext
  return file
}

/**
 * 节点无坐标时的兜底位。
 *
 * ★ 为什么要有：store 里 positions 理论上都有值，但老工程/手工改过的文件可能缺。
 *   全给 {0,0} 会让所有节点叠在一起，看着像「节点丢了」——
 *   叠一起比报错更难排查。按 id 散开至少能看到一片节点。
 */
export function fallbackPosition(id: string): NodePos {
  let h = 0
  for (let i = 0; i < id.length; i++) h = (h * 31 + id.charCodeAt(i)) >>> 0
  const cols = 6
  const gapX = 280
  const gapY = 200
  return { x: (h % cols) * gapX, y: ((h >>> 8) % cols) * gapY }
}

/** 校验单个节点文件 */
export function validateNodeFile(file: unknown, expectId: string): ProjectFormatProblem[] {
  const problems: ProjectFormatProblem[] = []
  if (!file || typeof file !== 'object') {
    return [{ level: 'error', message: '节点文件不是对象', nodeId: expectId }]
  }
  const f = file as Partial<ProjectNodeFile>
  if (f.id !== expectId) {
    problems.push({
      level: 'error',
      message: `节点文件内 id（${String(f.id)}）与文件名（${expectId}）不一致，可能被手工改过`,
      nodeId: expectId,
    })
  }
  if (!f.schema) {
    problems.push({ level: 'error', message: '节点文件缺 schema', nodeId: expectId })
  }
  if (!f.position || typeof f.position.x !== 'number' || typeof f.position.y !== 'number') {
    problems.push({ level: 'warn', message: '节点缺坐标，已用兜底位', nodeId: expectId })
  }
  if (f.rows !== undefined && !Array.isArray(f.rows)) {
    problems.push({ level: 'error', message: 'rows 不是数组', nodeId: expectId })
  }
  // ★ 未知 kind 只 warn，绝不 error —— 这是「节点还没做完」的必然状态。
  //   用户说节点类型还在演进，那么用新版本存的工程被旧程序打开时，
  //   kind 必然在白名单外。此时若报 error，程序会因为「不认识的节点」
  //   拒绝打开整个工程 —— 几百节点的工程为一个新节点全打不开，
  //   且用户无法自救（改 JSON 去掉节点更糟）。正确姿势是：能打开、
  //   数据原样保留、画布上按未知节点降级渲染、提示用户升级程序。
  if (f.schema && typeof f.schema === 'object') {
    const kind = (f.schema as { kind?: unknown }).kind
    if (typeof kind === 'string' && !KNOWN_NODE_KINDS.has(kind)) {
      problems.push({
        level: 'warn',
        message: `节点类型 ${kind} 当前程序不认识，数据会原样保留，建议升级软件`,
        nodeId: expectId,
      })
    }
  }
  if (f.nodeVersion !== undefined && typeof f.nodeVersion !== 'number') {
    problems.push({ level: 'warn', message: 'nodeVersion 不是数字，已忽略', nodeId: expectId })
  }
  if (f.extensions !== undefined && (typeof f.extensions !== 'object' || f.extensions === null)) {
    problems.push({ level: 'warn', message: 'extensions 不是对象，已忽略', nodeId: expectId })
  }
  // ★ viewState 只做类型兜底，不做内容校验：
  //   它全是可选的 UI 状态，坏掉的顶多显示成默认值，不该影响打开。
  if (f.viewState !== undefined && (typeof f.viewState !== 'object' || f.viewState === null)) {
    problems.push({ level: 'warn', message: 'viewState 不是对象，已忽略', nodeId: expectId })
  }
  // ★ assets 的校验**分级**：结构问题只 warn（丢的是登记信息），
  //   但字段缺失是 error —— 因为那会让打开资产的路径直接崩掉。
  //   字节本体是否存在由 loadProject 用 fs 单独校验（这里看不到文件系统）。
  if (f.assets !== undefined) {
    if (!Array.isArray(f.assets)) {
      problems.push({ level: 'error', message: 'assets 不是数组', nodeId: expectId })
    } else {
      for (const [i, a] of f.assets.entries()) {
        if (!a || typeof a !== 'object') {
          problems.push({
            level: 'warn',
            message: `assets[${i}] 不是对象，已忽略`,
            nodeId: expectId,
          })
          continue
        }
        const ref = a as Partial<NodeAssetRef>
        const missing = (['id', 'relPath', 'name', 'role', 'mime'] as const).filter(
          (k) => typeof ref[k] !== 'string' || ref[k] === '',
        )
        if (missing.length > 0) {
          problems.push({
            level: 'error',
            message: `assets[${i}] 缺字段 ${missing.join('、')}，该资产无法定位`,
            nodeId: expectId,
          })
        }
      }
    }
  }
  return problems
}

/**
 * 当前程序认识的节点类型（用于「不认识的 kind 只 warn」判定）。
 *
 * ★ 这里必须与 schema.ts 的 NodeKind 联合类型保持一致，改一处必须改两处。
 *   有测试守着（projectFormat.test.ts 会拿 NodeKind 全量交叉核对），漏改会被测出来。
 *   之所以要冗余抄一份而不用 `satisfies` 派生：NodeKind 是**编译期**类型，
 *   拿不到运行时数组；想要运行时数组只能手抄。手抄 + 测试守是可控的。
 */
const KNOWN_NODE_KINDS: ReadonlySet<string> = new Set([
  'table',
  'param',
  'sort',
  'output',
  'start',
  'map',
  'docx',
  'xlsx',
  'text',
  'rule',
  'ai',
  'paramlib',
  'group',
  'note',
  'pdfmerge',
  'photo',
  'pdf',
  'folder',
  'webmon',
  'xlsx2dxf',
  'pdfreplace',
  'pdfslim',
  'pdf2word',
  'fileout',
  'ooword',
  'oosheet',
  'chapter',
  'test_table',
])

/** 工程是否可读：无 error 即可 */
export function isReadable(problems: ProjectFormatProblem[]): boolean {
  return !problems.some((p) => p.level === 'error')
}

/** 汇总校验问题为一行文本（报错时展示） */
export function formatProblems(problems: ProjectFormatProblem[]): string {
  return problems.map((p) => `[${p.level}] ${p.nodeId ? `${p.nodeId}: ` : ''}${p.message}`).join('\n')
}
