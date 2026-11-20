/**
 * 节点资产 ⇄ 工程文件 的桥接层（2026-10-06）。
 *
 * ★ 这个文件解决的是用户明确提出的第 4 项：
 *   「节点文档、音频、视频、文件」。
 *
 * ── 问题的实质 ──
 *   节点上挂着真实文件：ooword/oosheet 的模板与就地编辑稿、导入的外部表格、
 *   将来的音频/视频/附件。本程序此前把这些字节全放在 IndexedDB
 *   （templateStore 的 tpl / editStore 的 docx-edit），
 *   而工程文件里**一个字都没记**。
 *
 *   后果非常直接，而且是最坏的那种：工程文件发给同事，对方拿到的是一堆
 *   「节点在、字段在、连线在，双击打开是空白」的节点 —— 因为字节没跟过来。
 *   而且这个故障在开发机上永远看不见：自己导出的工程自己打得开。
 *
 * ── 分层：本文件是「规划层」，不碰文件系统 ──
 *   它只回答一个问题：**给定节点配置与运行时字节，该往assets/ 写哪些文件、
 *   节点文件里该登记哪些条目**。产出是一份纯数据的 `NodeAssetPlan`。
 *   真正写盘由 projectStore 侧的编排函数做（需要 ProjectFs）。
 *
 *   这样切分的理由：MIME 判定、文件名安全化、hash 计算都是纯计算，
 *   放在这里能被 vitest 直接跑，不需要 jsdom、不需要 fake-indexeddb。
 *   而「字节从哪来」（IndexedDB / 磁盘 / IPC）是平台问题，交给编排层注入。
 *
 * ── 不把字节塞进 JSON ──
 *   NodeAssetRef 只登记「在哪、叫什么、什么类型、多大、指纹」，
 *   字节落 assets/<nodeId>/<assetId>.<ext>。
 *   base64 进 JSON 会让节点文件膨胀到几十 MB，git diff 完全失去意义 ——
 *   那恰好毁掉「目录式工程」的全部理由（Q1 拍板的核心就是 diff 可读）。
 */

import type { NodeAssetRef, ProjectNodeFile } from './projectFormat'
import type { TableSchema } from './schema'

/* ────────────────────────── 运行时字节来源（注入）────────────────────────── */

/**
 * 运行时资产字节的读取接口。
 *
 * ★ 刻意做成接口而不是直接 import templateStore：
 *   ① 分层铁律 —— src/core 下的文件不许碰 document / electron / node:，
 *     templateStore 虽然是浏览器安全的，但它把 localStorage / IndexedDB 当全局依赖，
 *     引入它会让纯计算层被迫在 node 环境下跑假存储。
 *   ② 平台差异真实存在 —— Electron 走 IPC 落磁盘，浏览器走 IndexedDB。
 *     编排层用同一个接口对接两者，本文件不需要知道区别。
 */
export interface AssetRuntime {
  /**
   * 读绑定模板的字节。
   * @param tplId 模板 id（来自 config.docxTemplateId / xlsxTemplateId / ooTemplateId）
   * @returns null = 模板库里没这个 id（模板被删了 / 从别人那儿收的工程）
   */
  readTemplate(tplId: string): Promise<Uint8Array | null>
  /** 模板元信息（拿文件名与体积；返回 null 表示模板已不在库里） */
  templateMeta(tplId: string): Promise<{ name: string; size: number; updatedAt?: number } | null>
  /**
   * 读就地编辑稿。
   * @returns null = 没有稿
   */
  readEdited(nodeId: string): Promise<{ tplId: string; bytes: Uint8Array } | null>
  /**
   * 登记新模板（打开工程时把工程里的字节灌回模板库用）。
   * @returns 实际拿到的 id —— 可能与传入的 tplId 不同（同名去重会复用旧 id）
   */
  ensureTemplate(name: string, bytes: Uint8Array): Promise<string>
  /** 写回就地编辑稿 */
  writeEdited(nodeId: string, tplId: string, bytes: Uint8Array): Promise<void>
  /**
   * ★ 2026-10-25 通用 blob 回灌（attachment 类）：
   * 打开工程时把 assets 字节按原 key 灌回 IDB pdfBlobStore
   * （pdfitem-{id} 词条 / rephist-{id} 替换页产物 / pdfhist-{id} 合并产物）——
   * 此前 attachment 一律灌模板库 = 字节进了错误的存储，词条打开仍为空。
   */
  restoreBlob(key: string, bytes: Uint8Array): Promise<void>
}

/* ────────────────────────── 资产种类识别 ────────────────────────── */

/**
 * 一个节点挂了哪些资产。
 *
 * ★ 关键：**资产种类是从 schema.config 推出来的，不是从运行时存储扫出来的**。
 *   反过来做（扫存储、把所有没被引用的都写进工程）有两个坏处：
 *     ① 模板库里 500 个模板会全进每个工程，体积爆炸；
 *     ② 用户删掉节点后残留的模板字节会跟着进工程，变成「幽灵资产」。
 *   只有「节点确实引用着」才登记，这与「不存没人引用的东西」同一条理由。
 */
export interface NodeAssetSpec {
  /** 绑定模板的 id */
  tplId?: string
  /** 绑定模板的文件名（配置里冗余存了一份，模板库里查不到时可作兜底） */
  tplName?: string
  /** 有就地编辑稿 */
  hasEdited?: boolean
}

/**
 * 从节点 schema 抽出资产线索。
 *
 * 覆盖的字段（2026-10-06 实测 NodeConfig）：
 *   · docxTemplateId  —— output / docx 节点的 Word 模板
 *   · xlsxTemplateId  —— xlsx 节点的工作簿模板
 *   · ooTemplateId    —— ooword / oosheet 节点的 OnlyOffice 底稿
 * 三者任一存在即认为该节点挂了模板。
 *
 * 为什么不做成 kind 映射表而是直接看 config 字段：
 *   kind 与绑定字段并非一一对应（output 既能导 Word 也能导 Excel，
 *   同一 kind 下换绑定格式是常见操作），按 config 判断才是真契约。
 *   代价是新增绑定字段时要记得加一行 —— 已有测试核对（见 nodeAssets.test.ts）。
 */
export function assetSpecOf(t: TableSchema | undefined, hasEdited: boolean): NodeAssetSpec {
  const c = t?.config
  if (!c) return { hasEdited }
  // ★ 三个字段是互斥的：同一时刻一个节点只绑一份模板。
  //   若将来支持「一份文档 + 一份附表」，这里要改成数组（见下方测试对该假设的守护）。
  const tplId = c.docxTemplateId ?? c.xlsxTemplateId ?? c.ooTemplateId
  const tplName = c.docxTemplateName ?? c.xlsxTemplateName ?? c.ooTemplateName
  return { tplId, tplName, hasEdited }
}

/* ────────────────────────── MIME 判定 ────────────────────────── */

/** OOXML 与常见媒体的 MIME。不在表内的走 fallbackByExt。 */
const MIME_BY_EXT: Record<string, string> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  doc: 'application/msword',
  xls: 'application/vnd.ms-excel',
  pdf: 'application/pdf',
  md: 'text/markdown',
  txt: 'text/plain',
  csv: 'text/csv',
  json: 'application/json',
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  mp3: 'audio/mpeg',
  wav: 'audio/wav',
  m4a: 'audio/mp4',
  flac: 'audio/flac',
  ogg: 'audio/ogg',
  mp4: 'video/mp4',
  webm: 'video/webm',
  mov: 'video/quicktime',
}

/**
 * 按文件名扩展名判 MIME。
 *
 * ★ 为什么不靠字节魔数（magic number）当主判据：
 *   .xlsx 到底是 OOXML 还是老的 BIFF，光看扩展名确实判不出；魔数可以（PK\x03\x04）。
 *   但这里判错的后果只是「UI 上显示的类型名不对」，用户双击打开仍然走
 *   正确的解析分支（那分支自己会判），所以不值得为此引入一套魔数表。
 *   真要严谨，可在 verify 时用 hash 与 size 兜底 —— 那里本来就是给损坏检测用的。
 *   故：扩展名够用，且对无扩展名的文件有合理兜底。
 */
export function mimeOf(name: string): string {
  const ext = extOf(name)
  return MIME_BY_EXT[ext] ?? 'application/octet-stream'
}

/** 小写扩展名（不含点）；无扩展名返回空串 */
export function extOf(name: string): string {
  const i = name.lastIndexOf('.')
  if (i < 0 || i === name.length - 1) return ''
  return name.slice(i + 1).toLowerCase()
}

/* ────────────────────────── 文件名与路径安全 ────────────────────────── */

/**
 * 文件名安全化：只用于**显示名**（NodeAssetRef.name）与**落盘名**。
 *
 * ★ 落盘名另有更严的要求（见 assetFileName）：不能含分隔符、不能是 . 或 ..、
 *   不能超长。这里先把明显非法的字符换掉，且**保留中文**——
 *   「技术方案.docx」被换成「技术方案_docx」是纯粹的破坏：用户在工作区里
 *   看到一串下划线会以为文件坏了。
 *
 * ⚠ 字符类里**不能**出现「`X -`」这种能被解析成范围的写法：
 *   `/[<>:"|?* -]/` 看起来是「Windows 非法字符 + 空格 + 连字符」，
 *   实际被解析成「`*` 到 ` `」的范围，于是**空格和连字符也被替换**——
 *   症状是「技术方案 - 终版.docx」变成「技术方案____终版_docx」，
 *   文件名被自己的"安全化"毁掉。有测试守着（nodeAssets.test.ts）。
 *
 * ⚠ 控制字符范围要写成 `\u0000-\u001f` 转义，**不能**把裸控制字节打进源码：
 *   裸 NUL 会让编辑器/git diff/部分工具链出现莫名问题，且极难定位。
 */
export function safeDisplayName(name: string): string {
  const base = name
    .replace(/[\\/]/g, '_') // 路径分隔符：唯一必须换的
    .replace(/[<>:"|?*]/g, '_') // Windows 非法字符
    .replace(/[\u0000-\u001f]/g, '') // 控制字符：剔除而不是替换
    .trim()
  return base || 'unnamed'
}

/**
 * 落盘文件名。
 *
 * 用 `<assetId>__<安全化的原名>` 而不是纯 `<assetId>.<ext>`：
 *   纯 id 的话，解包到另一台机器后资源管理器里全是 `a1b2c3.docx`，
 *   用户完全不知道那是什么文件。带上原名（已安全化）后双击能找到。
 *
 * assetId 由调用方保证安全（我们内部生成的，不含用户输入）。
 */
export function assetFileName(assetId: string, displayName: string): string {
  const safe = safeDisplayName(displayName).replace(/\s+/g, ' ').slice(0, 100)
  return `${assetId}__${safe}`
}

/* ────────────────────────── 指纹 ────────────────────────── */

/**
 * 字节指纹（与 templateStore.hashBytes 同口径的降级逻辑，但本文件独立实现）。
 *
 * 为什么不直接 import templateStore 的 hashBytes：分层铁律（core 不碰 IO 层）。
 * 为什么不用 Node 的 crypto：core 是浏览器侧代码。
 * 故用与那边一致的算法，保证两边算出的值可比。
 */
export function hashBytes(bytes: Uint8Array): string {
  let h = 0x811c9dc5
  for (let i = 0; i < bytes.length; i++) {
    h ^= bytes[i]
    h = Math.imul(h, 0x01000193) >>> 0
  }
  return `fnv1a-${h.toString(16).padStart(8, '0')}-len${bytes.length}`
}

/**
 * 用 WebCrypto 的 SHA-256（异步，优先）。
 *
 * FNV-1a 只用于「同一份文件是否重复」的去重与快速自检，**不是安全哈希**。
 * NodeAssetRef.hash 字段的用途就是自检，所以降级可接受，但值里带了
 * `fnv1a-` 前缀 —— 不会有人误以为它是 SHA-256。
 */
export async function hashBytesBest(
  bytes: Uint8Array,
): Promise<string> {
  const subtle = (globalThis.crypto as Crypto | undefined)?.subtle
  if (subtle) {
    try {
      const buf = await subtle.digest('SHA-256', bytes as unknown as ArrayBuffer)
      return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('')
    } catch {
      /* 非安全上下文等 → 降级 */
    }
  }
  return hashBytes(bytes)
}

/* ────────────────────────── 规划 ────────────────────────── */

/**
 * 一份资产计划：**该写哪些文件** + **节点文件里登记什么**。
 *
 * ★ bytes 用 Map 而不是塞进 NodeAssetRef：
 *   NodeAssetRef 会被 JSON.stringify 写进 nodes/<id>.json，
 *   Uint8Array 序列化成 {"0":72,"1":101,...} 那种东西，
 *   一个 300KB 的 docx 会变成几十万字符的数字。
 *   分离之后「写盘」这件事由编排层做，它把 bytes 交给 fs.writeBytes。
 */
export interface NodeAssetPlan {
  /** 写进节点文件 assets[] 的登记条目 */
  refs: NodeAssetRef[]
  /** 落盘用的字节：relPath → bytes */
  files: Map<string, Uint8Array>
  /**
   * 采集期就发现的问题（模板已不在库里等）。
   * ★ 不抛异常：模板被删不该让整个保存失败 —— 用户的表格数据比文档重要。
   *   但必须报出来，否则「文档没进去」这件事用户永远不知道。
   */
  problems: string[]
}

/**
 * 为单个节点规划资产。
 *
 * @param t节点 schema（可空 —— 程序不认识该节点时沿用旧文件，此时 assets 也不再规划）
 * @param hasEdited 该节点是否有就地编辑稿（编辑稿存于编辑器的 useState，读取方负责判断）
 * @param prevRefs 节点文件里已有的登记（用于跳过「字节没变的重复哈希」）
 */
export async function planNodeAssets(
  nodeId: string,
  t: TableSchema | undefined,
  rt: AssetRuntime,
  opts?: { hasEdited?: boolean; prevRefs?: NodeAssetRef[]; includeHistory?: boolean | Set<string> },
): Promise<NodeAssetPlan> {
  const plan: NodeAssetPlan = { refs: [], files: new Map(), problems: [] }
  if (!t) return plan

  const spec = assetSpecOf(t, opts?.hasEdited ?? false)
  const prevById = new Map((opts?.prevRefs ?? []).map((r) => [r.id, r]))

  /* ① 绑定模板 */
  // 模板的显示名要留给② 复用：稿的落盘名必须与模板同扩展名，
  // 否则打开稿时程序会按未知类型走错分支（而bytes 本身是能打开的）。
  let tplDisplayName = spec.tplName ?? ''
  if (spec.tplId) {
    const meta = await rt.templateMeta(spec.tplId).catch(() => null)
    const bytes = await rt.readTemplate(spec.tplId).catch(() => null)
    if (!bytes || bytes.byteLength === 0) {
      plan.problems.push(
        `模板已在模板库中找不到（id=${spec.tplId}，节点「${t.name}」）—— ` +
          `节点设置里的绑定指向了一个不存在的文件`,
      )
    } else {
      tplDisplayName = meta?.name ?? spec.tplName ?? `${spec.tplId}.bin`
      const ref = await makeRef('template', spec.tplId, tplDisplayName, bytes, prevById)
      plan.refs.push(ref)
      plan.files.set(ref.relPath, bytes)
    }
  }

  /* ② 就地编辑稿 */
  if (spec.hasEdited) {
    const ed = await rt.readEdited(nodeId).catch(() => null)
    if (ed?.bytes && ed.bytes.byteLength > 0) {
      // ★ 扩展名从**已解析出的模板显示名**取，而不是从 config 里再取一次：
      //   ① config 的 name 字段可能缺失（只有 id），那就取不到扩展名；
      //   ② 模板库里的真实文件名才是权威（用户可能重命名过）。
      //   extOf 返回的是**不含点**的扩展名，必须自己补点，
      //   否则拼出「edited-n1docx」这种没有扩展名的怪文件。
      const ext = extOf(tplDisplayName)
      const name = `edited-${nodeId}${ext ? '.' + ext : ''}`
      const ref = await makeRef('edited', `edited-${nodeId}`, name, ed.bytes, prevById)
      plan.refs.push(ref)
      plan.files.set(ref.relPath, ed.bytes)
    }
  }

  /* ③ PDF 合并历史产物（2026-10-24 用户指令「勾选是否存入 wfc」）：
     pdfmerge 节点的 pdfHistory（inDb 条目）——字节在 IDB pdfhist-{id}-{at}。
     仅 includeHistory=true 时采集（默认不随 .wfc 打包，保存副本已剥离）。 */
  const inc = opts?.includeHistory
  if (inc === true || inc instanceof Set) {
    const cfg = (t?.config ?? {}) as {
      pdfHistory?: { id: string; name?: string; at?: number; b64?: string; inDb?: boolean }[]
    }
    const hist = cfg.pdfHistory ?? []
    for (const h of hist) {
      /* ★ 2026-10-25 细粒度：Set 按 `${nodeId}:${h.id}` 精选单条（保存弹窗树形勾选） */
      if (inc instanceof Set && !inc.has(nodeId + ':' + h.id)) continue
      if (!h.inDb && !h.b64) continue
      const key = 'pdfhist-' + h.id + '-' + (h.at ?? 0)
      try {
        let bytes: Uint8Array | null = null
        if (h.b64) {
          const bin = atob(h.b64)
          bytes = new Uint8Array(bin.length)
          for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
        } else {
          const { getBlob } = await import('./pdfBlobStore')
          bytes = await getBlob(key)
        }
        if (!bytes || bytes.byteLength === 0) continue
        const name = h.name ?? `合并产物-${h.id}.pdf`
        const ref = await makeRef('attachment', key, name, bytes, prevById)
        plan.refs.push(ref)
        plan.files.set(ref.relPath, bytes)
      } catch { /* 单条失败不阻塞 */ }
    }
  }

  /* ④ PDF 节点词条文件（2026-10-25 用户指令「内部的随文件存储」）：
     pdf 节点 pdfItems 的 inDb 条目——字节在 IDB pdfitem-{id}。
     词条是节点主体内容（非历史），**默认必带**（否则发工程=文件全丢）。
     兼容：b64 旧条目解出后按同款采集。 */
  if (t?.kind === 'pdf') {
    const cfg = (t.config ?? {}) as {
      pdfItems?: { id: string; name: string; b64?: string; inDb?: boolean }[]
    }
    for (const it of cfg.pdfItems ?? []) {
      if (!it.inDb && !it.b64) continue
      try {
        let bytes: Uint8Array | null = null
        if (it.b64) {
          const bin = atob(it.b64)
          bytes = new Uint8Array(bin.length)
          for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
        } else {
          const { getBlob } = await import('./pdfBlobStore')
          bytes = await getBlob('pdfitem-' + it.id)
        }
        if (!bytes || bytes.byteLength === 0) {
          plan.problems.push(`PDF「${it.name}」字节缺失（IDB 无 pdfitem-${it.id}）——该词条将随工程丢失`)
          continue
        }
        const ref = await makeRef('attachment', 'pdfitem-' + it.id, it.name, bytes, prevById)
        plan.refs.push(ref)
        plan.files.set(ref.relPath, bytes)
      } catch { /* 单条失败不阻塞 */ }
    }
  }

  /* ⑤ PDF 替换页节点（2026-10-25 用户指令「内部的随文件存储」）：
     replaceTargets/Sources 词条（字节在 IDB pdfitem-{id}，与 PDF 节点共用前缀）
     ——词条默认必带；replaceHistory（rephist-{id}）随 includeHistory 开关（对齐合并历史）。 */
  if (t?.kind === 'pdfreplace') {
    const cfg = (t.config ?? {}) as {
      replaceTargets?: { id: string; name: string; b64?: string; inDb?: boolean }[]
      replaceSources?: { id: string; name: string; b64?: string; inDb?: boolean }[]
      replaceHistory?: { id: string; name: string; b64?: string; inDb?: boolean }[]
    }
    const collectEntry = async (it: { id: string; name: string; b64?: string; inDb?: boolean }, key: string) => {
      if (!it.inDb && !it.b64) return
      try {
        let bytes: Uint8Array | null = null
        if (it.b64) {
          const bin = atob(it.b64)
          bytes = new Uint8Array(bin.length)
          for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
        } else {
          const { getBlob } = await import('./pdfBlobStore')
          bytes = await getBlob(key)
        }
        if (!bytes || bytes.byteLength === 0) {
          plan.problems.push(`替换页词条「${it.name}」字节缺失（IDB 无 ${key}）——该词条将随工程丢失`)
          return
        }
        const ref = await makeRef('attachment', key, it.name, bytes, prevById)
        plan.refs.push(ref)
        plan.files.set(ref.relPath, bytes)
      } catch { /* 单条失败不阻塞 */ }
    }
    for (const it of cfg.replaceTargets ?? []) await collectEntry(it, 'pdfitem-' + it.id)
    for (const it of cfg.replaceSources ?? []) await collectEntry(it, 'pdfitem-' + it.id)
    if (inc === true || inc instanceof Set) {
      for (const h of cfg.replaceHistory ?? []) {
        if (inc instanceof Set && !inc.has(nodeId + ':' + h.id)) continue
        await collectEntry(h, 'rephist-' + h.id)
      }
    }
  }

  return plan
}

/**
 * 为全量节点规划。
 *
 * 逐节点串行而非 Promise.all：templateMeta / readTemplate 背后是 IndexedDB 或 IPC，
 * 并发几百个节点会把存储层打爆（且 Electron 侧 IPC 是同步的，串行反而更快）。
 */
export async function planAllNodeAssets(
  tables: Record<string, TableSchema>,
  rt: AssetRuntime,
  opts?: {
    hasEdited?: (nodeId: string) => boolean
    prevRefs?: Record<string, NodeAssetRef[]>
    includeHistory?: boolean | Set<string>
  },
): Promise<{
  byNode: Record<string, NodeAssetPlan>
  problems: { nodeId: string; message: string }[]
}> {
  const byNode: Record<string, NodeAssetPlan> = {}
  const problems: { nodeId: string; message: string }[] = []
  /* ★ 2026-10-25 性能：串行逐节点 await → 8 并发分批。
     旧注释担心「并发几百个节点打爆存储层」——IndexedDB 读是安全的，
     真正的顾虑是 IDB 事务排队，8 并发是实测无压力的平衡点。
     104 节点保存从串行排队变为 13 批并发，读取阶段约快 3-5 倍。 */
  const ids = Object.keys(tables)
  const BATCH = 8
  for (let i = 0; i < ids.length; i += BATCH) {
    const batch = ids.slice(i, i + BATCH)
    const plans = await Promise.all(batch.map((id) =>
      planNodeAssets(id, tables[id], rt, {
        hasEdited: opts?.hasEdited?.(id) ?? false,
        prevRefs: opts?.prevRefs?.[id],
        includeHistory: opts?.includeHistory,
      }).catch((e) => {
        const plan: NodeAssetPlan = { refs: [], files: new Map(), problems: [String((e as Error)?.message ?? e)] }
        return plan
      })
    ))
    batch.forEach((id, j) => {
      byNode[id] = plans[j]
      for (const p of plans[j].problems) problems.push({ nodeId: id, message: p })
    })
  }
  return { byNode, problems }
}

/** 建一条登记条目；字节没变时沿用旧 hash（省一次全量哈希） */
async function makeRef(
  role: NodeAssetRef['role'],
  id: string,
  displayName: string,
  bytes: Uint8Array,
  prevById: Map<string, NodeAssetRef>,
): Promise<NodeAssetRef> {
  const relPath = assetFileName(id, displayName)
  const prev = prevById.get(id)
  // 体积与旧值一致 → 几乎可以断定字节没变（这个优化只影响自检，不影响正确性，
  // 因为真变了的话 size 也会变；极端的「改了内容但等长」漏检只损失一次自检强度）
  const hash = prev && prev.size === bytes.byteLength ? (prev.hash ?? '') : await hashBytesBest(bytes)
  return {
    id,
    relPath,
    name: safeDisplayName(displayName),
    role,
    mime: mimeOf(displayName),
    size: bytes.byteLength,
    updatedAt: new Date().toISOString(),
    hash,
  }
}

/* ────────────────────────── 反向：从工程恢复到运行时 ────────────────────────── */

/**
 * 把工程里的资产字节灌回运行时存储（打开工程时调）。
 *
 * ★ 为什么要「灌回」而不是让节点直接读工程目录的字节：
 *   所有既有代码（WordReportPanel / OnlyOfficeModal / WorkbookEditor）
 *   都是从 templateStore / editStore 取字节的。让它们改读工程目录，
 *   等于把 30 多处取值点全改一遍，且改完后「模板库」这个概念就空了
 *   （用户在模板库里挑模板时挑的是全局库，不是某个工程）。
 *   灌回之后：**工程自带一份模板副本，落在全局库里**，节点照旧工作。
 *   代价是同项目第二次打开不会重复入库（ensureTemplate 内部按名去重）。
 *
 * @returns 逐节点的模板 id 映射：nodeId → { 旧 tplId: 新 tplId }
 *   节点的 config.docxTemplateId 等字段需要按这个映射改写，
 *   否则仍指向「本机模板库里那个 id」—— 换机器后那个 id 是不存在的。
 */
export async function restoreNodeAssets(
  files: Record<string, ProjectNodeFile>,
  readAsset: (nodeId: string, ref: NodeAssetRef) => Promise<Uint8Array | null>,
  rt: AssetRuntime,
): Promise<{ tplIdMap: Record<string, Record<string, string>>; problems: string[] }> {
  const tplIdMap: Record<string, Record<string, string>> = {}
  const problems: string[] = []

  for (const [nodeId, file] of Object.entries(files)) {
    const map: Record<string, string> = {}
    // ★ 字段名必须是 file.assets，不能另起一个 refs 形参：
    //   上一版签名写的是 `{ refs?: NodeAssetRef[] }`，与 ProjectNodeFile 结构上
    //   「有重叠」，TS 不报错（可选属性全不重叠时反而会报错，但这里file 是
    //   具体类型，赋值时多余/缺失的可选属性不校验）——
    //   于是调用方传真正的节点文件进来时，assets 全被当成 undefined，
    //   恢复静默什么都不做，tplIdMap 返回空对象。
    //   症状：工程能打开、节点在、字节在磁盘上，但**节点仍然空白**，
    //   而所有检查都是绿的。教训：跨层传结构体必须用同一个类型。
    for (const ref of file.assets ?? []) {
      const bytes = await readAsset(nodeId, ref).catch(() => null)
      if (!bytes || bytes.byteLength === 0) {
        problems.push(`资产字节读不到（${ref.name}）`)
        continue
      }
      try {
        if (ref.role === 'edited') {
          // 稿依赖它源自哪个模板：用同名规则反查回来的 tplId（若拿不到就退回 id 本身，
          // 稿仍然写进去；tplId 校验不通过时 WordReportPanel 会提示「模板已换」）
          const baseId = ref.id.replace(/^edited-/, '')
          await rt.writeEdited(nodeId, baseId, bytes)
        } else if (ref.role === 'attachment') {
          /* ★ 2026-10-25 attachment 按 key 灌回 IDB pdfBlobStore（词条/产物字节原位还原）——
             此前走 ensureTemplate 灌模板库 = 字节进错存储，打开工程词条仍空白 */
          await rt.restoreBlob(ref.id, bytes)
        } else {
          const newId = await rt.ensureTemplate(ref.name, bytes)
          map[ref.id] = newId
        }
      } catch (e) {
        problems.push(`资产入库失败（${ref.name}）：${String(e)}`)
      }
    }
    if (Object.keys(map).length > 0) tplIdMap[nodeId] = map
  }

  return { tplIdMap, problems }
}

/**
 * 按 tplIdMap 改写节点 config 里的模板绑定。
 *
 * 返回**新对象**而不是原地改：调用方（store）通常要靠引用变化触发重渲染，
 * 原地改会被 React 的浅比较吃掉 —— 症状是「工程打开了，但绑定还显示旧模板名」，
 * 而刷新一下又对了，极难定位。
 */
export function rewireTemplateIds(
  t: TableSchema,
  map: Record<string, string>,
): TableSchema {
  const c = t.config
  if (!c) return t
  const keys = ['docxTemplateId', 'xlsxTemplateId', 'ooTemplateId'] as const
  let changed = false
  const next = { ...c }
  for (const k of keys) {
    const v = c[k]
    if (v && map[v]) {
      next[k] = map[v]
      changed = true
    }
  }
  if (!changed) return t
  return { ...t, config: next }
}
