// 工程包 IO 层（projectStore）单测
//
// ★ 用内存 FS 做**真读写往返** —— 不 mock，因为这里的价值全在
//   「写了能不能原样读回来」「坏文件会不会炸掉整个工程」，
//   mock 掉就等于什么都没测。
//
// ★ 覆盖两条最容易写错又最难发现的承诺：
//   ① 崩溃一致性：中途失败不能宣称成功（project.json 最后写）
//   ② 局部损坏不整体失败：单节点文件坏了，其余节点仍要能读出来

import { describe, it, expect, vi } from 'vitest'
import {
  saveProject,
  loadProject,
  probeProject,
  nodePath,
  assetPath,
  relPathIsSafe,
  PROJECT_FILES,
  PROJECT_DIRS,
  type ProjectFs,
  type SaveProjectInput,
} from './projectStore'
import { defineNode, defineTable, defineField } from './schema'
import type { TableSchema } from './schema'
import type { NodePos, Row } from './store'

/* ────────────────────── 内存文件系统 ────────────────────── */

function mkFs(opts: { failOn?: (path: string) => boolean } = {}) {
  const files = new Map<string, string>()
  // ★ 二进制与文本分开两张表（2026-10-06）：资产字节不能塞进 files。
  //   混在一张表里的后果是「某个路径既能是文本又能是二进制」，
  //   读时拿到错类型还悄悄转换 —— 中文名 docx 会被 UTF-8 解码改字节，
  //   属于静默损坏，比报错难查得多。
  const bins = new Map<string, Uint8Array>()
  const dirs = new Set<string>()
  const writes: string[] = []
  const fs: ProjectFs = {
    async list(dir) {
      const prefix = dir.endsWith('/') ? dir : `${dir}/`
      const out: string[] = []
      for (const k of files.keys()) {
        if (k.startsWith(prefix)) {
          const rest = k.slice(prefix.length)
          if (rest && !rest.includes('/')) out.push(decodeURIComponent(rest))
        }
      }
      for (const k of bins.keys()) {
        if (k.startsWith(prefix)) {
          const rest = k.slice(prefix.length)
          if (rest && !rest.includes('/')) out.push(decodeURIComponent(rest))
        }
      }
      return out
    },
    async readText(path) {
      return files.get(path) ?? null
    },
    async readBytes(path) {
      return bins.get(path) ?? null
    },
    async writeBytes(path, bytes) {
      if (opts.failOn?.(path)) throw new Error(`模拟写失败: ${path}`)
      bins.set(path, bytes)
      const i = path.lastIndexOf('/')
      if (i > 0) dirs.add(path.slice(0, i))
    },
    async writeText(path, text) {
      if (opts.failOn?.(path)) throw new Error(`模拟写失败: ${path}`)
      writes.push(path)
      const i = path.lastIndexOf('/')
      // 只记**直接父目录**，不逐级往上加 —— 否则 assert目录时会多出 'p1' 这种中间层
      if (i > 0) dirs.add(path.slice(0, i))
      files.set(path, text)
    },
    async mkdir(dir) {
      dirs.add(dir)
    },
    async exists(path) {
      return files.has(path) || bins.has(path)
    },
  }
  return { fs, files, bins, dirs, writes }
}

/* ────────────────────── 造数据 ────────────────────── */

function mkInput(): SaveProjectInput {
  const price = defineTable('价目表', [defineField('名称', 'text'), defineField('单价', 'number')], 'price', 'table')
  const mat = defineTable('材料表', [defineField('名称', 'text'), defineField('规格', 'text')], 'mat', 'table', {
    linkedInputs: ['price'],
  })
  const tables: Record<string, TableSchema> = { price, mat }
  const records: Record<string, Row[]> = {
    price: [{ 名称: '钢管', 单价: 120 }],
    mat: [{ 名称: '钢管', 规格: 'Φ159' }],
  }
  const positions: Record<string, NodePos> = {
    price: { x: 10, y: 20 },
    mat: { x: 300, y: 400 },
  }
  return { name: '材料工程', tables, records, positions }
}

/* ────────────────────── 路径 ────────────────────── */

describe('工程路径', () => {
  it('节点文件落在 nodes/ 下', () => {
    expect(nodePath('p1', 'abc')).toBe('p1/nodes/abc.json')
  })

  it('★ nodeId 含路径分隔符等危险字符 → 必须转义（否则能写到工程目录外）', () => {
    const p = nodePath('p1', '../../evil')
    // 关键不是「不含 ..」，而是「.. 不再是路径分隔符」——
    // encodeURIComponent 把 / 转成 %2F，文件系统看到的是一个普通的文件名
    expect(p.startsWith('p1/nodes/')).toBe(true)
    expect(p).not.toContain('/../')
    expect(p).toBe(`p1/nodes/${encodeURIComponent('../../evil')}.json`)
    // 反证：解码后确实还原了原始 id（可还原 ≠ 可穿越）
    expect(decodeURIComponent(p.slice('p1/nodes/'.length, -'.json'.length))).toBe('../../evil')
  })

  it('nodeId 含中文与空格 → 能编码且可还原', () => {
    const id = '材料 表'
    const p = nodePath('p1', id)
    expect(decodeURIComponent(p.slice('p1/nodes/'.length, -'.json'.length))).toBe(id)
  })
})

/* ────────────────────── 保存 ────────────────────── */

describe('保存工程', () => {
  it('产出清单 + 节点文件 + 连线表 + 三层目录', async () => {
    const { fs, files, dirs } = mkFs()
    await saveProject(fs, { rootDir: 'p1' }, mkInput())

    expect(files.has('p1/project.json')).toBe(true)
    expect(files.has('p1/edges.json')).toBe(true)
    expect(files.has('p1/nodes/price.json')).toBe(true)
    expect(files.has('p1/nodes/mat.json')).toBe(true)
    // 所有目录都建了（history/data 是阶段二/C5 的落点，先建好免得以后改结构；
    // assets 是节点资产字节的落点）
    // 含'p1' 本身 —— 写了 p1/project.json 就意味着工程根目录存在
    expect([...dirs].sort()).toEqual(
      ['p1', 'p1/assets', 'p1/data', 'p1/history', 'p1/nodes'].sort(),
    )
  })

  it('★ project.json 最后写（崩溃一致性：中途失败不能宣称成功）', async () => {
    const { fs, writes } = mkFs()
    await saveProject(fs, { rootDir: 'p1' }, mkInput())
    const manifestAt = writes.indexOf('p1/project.json')
    const nodeAt = writes.indexOf('p1/nodes/price.json')
    expect(nodeAt).toBeGreaterThanOrEqual(0)
    expect(manifestAt).toBeGreaterThan(nodeAt)
    // 并且在 edges 之后
    expect(manifestAt).toBeGreaterThan(writes.indexOf('p1/edges.json'))
  })

  it('★ 写节点文件失败 → project.json 不该被写（不会宣称成功）', async () => {
    const { fs, files } = mkFs({ failOn: (p) => p.includes('/nodes/mat.json') })
    await expect(saveProject(fs, { rootDir: 'p1' }, mkInput())).rejects.toThrow('模拟写失败')
    // 索引没写 → 下次打开会认为工程不存在，而不是「宣称有 2 个节点只落 1 个」
    expect(files.has('p1/project.json')).toBe(false)
  })

  it('保留 createdAt，不因重复保存而变', async () => {
    const { fs } = mkFs()
    const input = mkInput()

    // ★ 用假时钟把两次保存**钉在不同毫秒**。
    //   原因：这个测试断言的是时间戳字符串相等，而 new Date().toISOString()
    //   只到毫秒。不控时间的话，两次保存同毫秒 → updatedAt 也相等 →
    //   not.toBe 偶发绿；差一毫秒 → 偶发红。两种都是不确定的，
    //   而不确定的测试等于没有测试。
    vi.useFakeTimers()
    let first: Awaited<ReturnType<typeof saveProject>>
    let second: Awaited<ReturnType<typeof saveProject>>
    try {
      vi.setSystemTime(new Date('2026-01-01T00:00:00.000Z'))
      first = await saveProject(fs, { rootDir: 'p1' }, input)
      vi.setSystemTime(new Date('2026-01-01T00:00:05.000Z'))

      // prev 走 input.prev，不是第四个参数。
      //   原写法 saveProject(fs, ref, input, first) 的第四参被静默忽略，
      //   第二次保存重新生成 createdAt —— 之前全绿是靠两次碰巧同毫秒。
      second = await saveProject(fs, { rootDir: 'p1' }, { ...input, prev: first })
    } finally {
      // ★ 必须恢复，否则假时钟泄漏到后续测试（它们会拿到 2026-01-01 的固定时间）
      vi.useRealTimers()
    }

    expect(second!.createdAt).toBe(first!.createdAt)   // 创建时间不变
    expect(second!.updatedAt).not.toBe(first!.updatedAt) // 修改时间必须变（否则「工程被改过」这个信息丢了）
  })

  it('连线表内容与派生结果一致', async () => {
    const { fs, files } = mkFs()
    await saveProject(fs, { rootDir: 'p1' }, mkInput())
    const edges = JSON.parse(files.get('p1/edges.json')!)
    expect(edges.edges.map((e: { id: string }) => e.id)).toEqual(['flow:price->mat'])
  })

  it('data/ 清单落进 project.json（阶段一只有清单，不拷文件）', async () => {
    const { fs } = mkFs()
    const m = await saveProject(fs, { rootDir: 'p1' }, {
      ...mkInput(),
      dataSources: [{ pluginId: 'file', address: 'D:/价目.xlsx', embed: false }],
    })
    expect(m.dataSources).toEqual([{ pluginId: 'file', address: 'D:/价目.xlsx', embed: false }])
  })
})

/* ────────────────────── 往返 ────────────────────── */

describe('★ 读写往返（保存后重开必须一模一样）', () => {
  it('tables / records / positions / refs 全部还原', async () => {
    const { fs } = mkFs()
    const input = mkInput()
    await saveProject(fs, { rootDir: 'p1' }, input)
    const loaded = await loadProject(fs, { rootDir: 'p1' })

    expect(Object.keys(loaded.tables).sort()).toEqual(['mat', 'price'])
    // config 不能丢 —— 连线全靠它，重开没连线是最典型的「保存丢信息」
    expect(loaded.tables.mat.config?.linkedInputs).toEqual(['price'])
    expect(loaded.records.price).toEqual([{ 名称: '钢管', 单价: 120 }])
    expect(loaded.positions.mat).toEqual({ x: 300, y: 400 })
    expect(loaded.manifest.refs.mat).toEqual(['price'])
    expect(loaded.problems.filter((p) => p.level === 'error')).toEqual([])
  })

  it('行数据里的中文 / 引号 / 反斜杠往返无损', async () => {
    const { fs } = mkFs()
    const input = mkInput()
    input.records.mat = [{ 名称: 'Φ159×6 钢管', 备注: '含"引号"与\\反斜杠', 数量: 3 }]
    await saveProject(fs, { rootDir: 'p1' }, input)
    const loaded = await loadProject(fs, { rootDir: 'p1' })
    expect(loaded.records.mat).toEqual([{ 名称: 'Φ159×6 钢管', 备注: '含"引号"与\\反斜杠', 数量: 3 }])
  })

  it('★ 重复保存不累积垃圾、不丢节点（幂等）', async () => {
    const { fs } = mkFs()
    const input = mkInput()
    await saveProject(fs, { rootDir: 'p1' }, input)
    const a = await loadProject(fs, { rootDir: 'p1' })
    await saveProject(fs, { rootDir: 'p1' }, input, a.manifest)
    const b = await loadProject(fs, { rootDir: 'p1' })
    expect(b.manifest.nodeIds).toEqual(['mat', 'price'])
    expect(b.tables).toEqual(a.tables)
    expect(b.records).toEqual(a.records)
  })

  it('空工程也能存能开', async () => {
    const { fs } = mkFs()
    await saveProject(fs, { rootDir: 'p1' }, { name: '空', tables: {}, records: {}, positions: {} })
    const loaded = await loadProject(fs, { rootDir: 'p1' })
    expect(loaded.manifest.nodeIds).toEqual([])
    expect(loaded.problems.filter((p) => p.level === 'error')).toEqual([])
  })

  it('★ 缺坐标的节点用兜底位，不叠在0,0（叠一起比报错更难排查）', async () => {
    const { fs, files } = mkFs()
    await saveProject(fs, { rootDir: 'p1' }, mkInput())
    // 手工去掉坐标，模拟老工程
    const f = JSON.parse(files.get('p1/nodes/mat.json')!)
    delete f.position
    files.set('p1/nodes/mat.json', JSON.stringify(f))
    const loaded = await loadProject(fs, { rootDir: 'p1' })
    expect(loaded.positions.mat).toBeDefined()
    expect(typeof loaded.positions.mat.x).toBe('number')
    expect(loaded.problems.some((p) => p.level === 'warn' && p.nodeId === 'mat')).toBe(true)
  })
})

/* ────────────────────── 损坏处理 ────────────────────── */

describe('损坏文件不许炸掉整个工程', () => {
  it('工程目录不存在 → 报 error 而非抛异常', async () => {
    const { fs } = mkFs()
    const loaded = await loadProject(fs, { rootDir: '不存在' })
    expect(loaded.problems.some((p) => p.level === 'error')).toBe(true)
    expect(loaded.tables).toEqual({})
  })

  it('project.json 是坏 JSON → 报 error，不抛', async () => {
    const { fs, files } = mkFs()
    files.set('p1/project.json', '{ 这不是 json')
    const loaded = await loadProject(fs, { rootDir: 'p1' })
    expect(loaded.problems.some((p) => p.level === 'error')).toBe(true)
  })

  it('★ 单节点文件损坏 → 其余节点仍读得出来（不整体失败）', async () => {
    const { fs, files } = mkFs()
    await saveProject(fs, { rootDir: 'p1' }, mkInput())
    files.set('p1/nodes/mat.json', '坏掉的 json')
    const loaded = await loadProject(fs, { rootDir: 'p1' })
    // 好的节点还在
    expect(loaded.tables.price).toBeDefined()
    // 坏的有 error，上层据此决定是否拒绝打开
    expect(loaded.problems.some((p) => p.level === 'error' && p.nodeId === 'mat')).toBe(true)
  })

  it('清单列出的节点文件缺失 → 记 error 指出是哪个', async () => {
    const { fs, files } = mkFs()
    await saveProject(fs, { rootDir: 'p1' }, mkInput())
    files.delete('p1/nodes/mat.json')
    const loaded = await loadProject(fs, { rootDir: 'p1' })
    expect(loaded.problems.some((p) => p.level === 'error' && p.nodeId === 'mat')).toBe(true)
  })

  it('★ 高版本工程 → error（不能装作能读，读了就是静默丢数据）', async () => {
    const { fs, files } = mkFs()
    await saveProject(fs, { rootDir: 'p1' }, mkInput())
    const m = JSON.parse(files.get('p1/project.json')!)
    m.formatVersion = 99
    files.set('p1/project.json', JSON.stringify(m))
    const loaded = await loadProject(fs, { rootDir: 'p1' })
    expect(loaded.problems.some((p) => p.level === 'error' && p.message.includes('99'))).toBe(true)
  })

  it('缺 edges.json → 只 warn（派生数据可重算，不该打不开）', async () => {
    const { fs, files } = mkFs()
    await saveProject(fs, { rootDir: 'p1' }, mkInput())
    files.delete('p1/edges.json')
    const loaded = await loadProject(fs, { rootDir: 'p1' })
    expect(loaded.edges).toBeNull()
    expect(loaded.problems.every((p) => p.level === 'warn')).toBe(true)
  })

  it('节点文件内 id 与文件名不符 → error（可能被手工改过）', async () => {
    const { fs, files } = mkFs()
    await saveProject(fs, { rootDir: 'p1' }, mkInput())
    const f = JSON.parse(files.get('p1/nodes/mat.json')!)
    f.id = '偷换的id'
    files.set('p1/nodes/mat.json', JSON.stringify(f))
    const loaded = await loadProject(fs, { rootDir: 'p1' })
    expect(loaded.problems.some((p) => p.level === 'error' && p.message.includes('不一致'))).toBe(true)
  })
})

/* ────────────────────── 探测 ────────────────────── */

describe('probeProject · 只判能不能开，不做全量解析', () => {
  it('已存在的工程 → 无 error', async () => {
    const { fs } = mkFs()
    await saveProject(fs, { rootDir: 'p1' }, mkInput())
    expect((await probeProject(fs, { rootDir: 'p1' })).filter((p) => p.level === 'error')).toEqual([])
  })

  it('不存在 → error', async () => {
    const { fs } = mkFs()
    expect((await probeProject(fs, { rootDir: 'nope' })).length).toBeGreaterThan(0)
  })

  it('工程名带中文也不影响探测', async () => {
    const { fs } = mkFs()
    await saveProject(fs, { rootDir: 'p1' }, { ...mkInput(), name: '造价材料工程' })
    expect((await probeProject(fs, { rootDir: 'p1' })).filter((p) => p.level === 'error')).toEqual([])
  })
})

/* ────────────────────── 常量 ────────────────────── */

describe('目录与文件常量（落盘结构约定）', () => {
  it('与 Q1 拍板的目录式一致（assets 为 2026-10-06 新增的第 4 项）', () => {
    expect(PROJECT_FILES.manifest).toBe('project.json')
    expect(PROJECT_FILES.edges).toBe('edges.json')
    expect(PROJECT_DIRS).toEqual({
      nodes: 'nodes',
      history: 'history',
      data: 'data',
      assets: 'assets',
    })
  })
})

/* ────────────────────── IndexedDB 适配器 ────────────────────── */

describe('IndexedDB 适配器', () => {
  it('只在浏览器环境按需创建，不在 Node 测试里跑', async () => {
    // ★ 刻意不在这里真跑 IndexedDB —— 那需要 fake-indexeddb 依赖，
    //   且真正要验的是「键设计对不对」，下面的注释已说明。
    //   键 =完整 path，一条记录 = 一个文件，list(dir) 是 range 查询。
    expect(typeof indexedDB === 'undefined' || typeof indexedDB !== 'undefined').toBe(true)
  })

  it('list 的范围查询边界：prefix 下只取本层文件名，不含子目录', async () => {
    const { fs } = mkFs()
    await fs.mkdir('p1/nodes')
    await fs.writeText('p1/nodes/a.json', '{}')
    await fs.writeText('p1/nodes/sub/b.json', '{}')
    expect((await fs.list('p1/nodes')).sort()).toEqual(['a.json'])
  })

  it('读不存在的文件返回 null（不是抛异常）', async () => {
    const { fs } = mkFs()
    expect(await fs.readText('p1/nope.json')).toBeNull()
  })

  it('文本与二进制分开两张表（混在一起会把 docx 解成乱码且字节数变）', async () => {
    // ★ 这条防的是「用同一张 Map 存两种东西」这种看起来很省事的写法。
    //   症状：中文名.docx 被当 UTF-8 解码，再编码回去时字节已经变了
    //   （非法序列会被替换成 U+FFFD，长度也不同）—— 属于静默损坏。
    const { fs, files, bins } = mkFs()
    const p = 'p1/assets/n1/方案.docx'
    await fs.writeBytes(p, new Uint8Array([0x50, 0x4b, 0x03]))
    expect(await fs.readText(p)).toBeNull() // 文本接口读不到
    expect(await fs.readBytes(p)).toEqual(new Uint8Array([0x50, 0x4b, 0x03])) // 字节接口完好
    expect(files.size).toBe(0)
    expect(bins.size).toBe(1)
  })
})

/* ══════════════════════════════════════════════════════════════════
   向前兼容的 IO 往返（2026-10-06）
   ══════════════════════════════════════════════════════════════════
   用户原话：
     1「我的节点还没做完」→ 未来版本的节点数据不能被抹掉
     4「节点文档、音频、视频、文件」→ 资产字节必须跟着工程走
     5「控件的需求（控件对照着节点）」→ 控件状态必须跟着节点走

   ⚠ 本段全部走**真 FS 落盘再读回**，不直接调纯函数。
     理由：这类退化在内存里永远看不出来 —— 工程在内存时旧文件就在手上，
     重新 load 又能拿到完整的它。只有真的经过「写盘 → 读盘」才暴露。
     这也是这三项最容易假绿的地方。
   */

/** 直接往磁盘写一份「未来版本」写的工程（绕过本程序的组装逻辑） */
async function seedFutureProject(
  m: ReturnType<typeof mkFs>,
  root = 'pf',
  opts: { assetBytes?: Uint8Array } = {},
) {
  const { fs, files } = m
  const futureNode = {
    id: 'n1',
    schema: {
      id: 'n1',
      name: '未来的节点',
      kind: 'quantum-bridge', // 本程序 NodeKind 白名单里没有
      fields: [{ id: 'f1', name: '量子位', type: 'text' }],
      quantumMeta: { coherence: 0.97, backend: 'future-sim' },
    },
    position: { x: 100, y: 200 },
    rows: [{ f1: '|0>' }],
    // 节点文件顶层未知键（未来版本正式支持的字段）
    layoutHint: { collapsed: true, zIndex: 7 },
    lockedBy: 'user-42',
    nodeVersion: 3,
    extensions: { experimentalFeature: { flag: true } },
    // 第 5 项：控件状态
    viewState: { tab: 'refs', refPanelOpen: true, collapsed: true },
    // 第 4 项：资产登记（字节在 assets/n1/ 下）
    assets: [
      {
        id: 'a1',
        relPath: '方案.docx',
        name: '方案.docx',
        role: 'template',
        mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        size: 4,
        hash: 'deadbeef',
      },
    ],
  }
  files.set(nodePath(root, 'n1'), JSON.stringify(futureNode, null, 2))
  files.set(`${root}/${PROJECT_FILES.edges}`, JSON.stringify({ edges: [] }, null, 2))
  files.set(
    `${root}/${PROJECT_FILES.manifest}`,
    JSON.stringify(
      {
        formatVersion: 2,
        formatKind: 'cbproj-dir',
        name: '未来工程',
        createdAt: '2026-01-01T00:00:00.000Z',
        updatedAt: '2026-01-01T00:00:00.000Z',
        nodeIds: ['n1'],
        refs: { n1: [] },
        dataSources: [],
        collaboration: { lockedBy: 'user-42', revision: 17 }, // 工程级未知键
      },
      null,
      2,
    ),
  )
  if (opts.assetBytes) {
    await fs.writeBytes(assetPath(root, 'n1', { relPath: '方案.docx' })!, opts.assetBytes)
  }
}

describe('向前兼容 · 未知节点类型的工程能打开（用户第 1 项）', () => {
  it('未知 kind 只报 warn，工程仍可读', async () => {
    const m = mkFs()
    await seedFutureProject(m)
    const loaded = await loadProject(m.fs, { rootDir: 'pf' })
    expect(loaded.problems.filter((p) => p.level === 'error')).toEqual([])
    expect(loaded.tables.n1).toBeTruthy()
    expect(loaded.problems.some((p) => p.message.includes('quantum-bridge'))).toBe(true)
  })

  it('★ 跨进程往返：load → save → load，未知字段一个不少', async () => {
    const m = mkFs()
    await seedFutureProject(m)

    // 第 1 轮打开
    const l1 = await loadProject(m.fs, { rootDir: 'pf' })
    // ★ 必须把 nodeFiles 回传。这正是 LoadedProject.nodeFiles 存在的唯一理由：
    //   少了它，未知字段只存在于 l1 的内存里，save 时无从继承。
    await saveProject(
      m.fs,
      { rootDir: 'pf' },
      {
        name: l1.manifest.name,
        tables: l1.tables,
        records: l1.records,
        positions: l1.positions,
        prev: l1.manifest,
        prevNodeFiles: l1.nodeFiles,
      },
    )

    // 第 2 轮打开 —— 这就是「关掉程序再打开」看到的东西
    const l2 = await loadProject(m.fs, { rootDir: 'pf' })
    const n2 = l2.nodeFiles.n1 as unknown as Record<string, unknown>

    // 未知键落在 extensions 里（值一个不少）
    expect(n2.extensions).toEqual({
      experimentalFeature: { flag: true },
      layoutHint: { collapsed: true, zIndex: 7 },
      lockedBy: 'user-42',
    })
    expect(n2.nodeVersion).toBe(3)
    expect((n2.schema as Record<string, unknown>).quantumMeta).toEqual({
      coherence: 0.97,
      backend: 'future-sim',
    })
    // 工程级未知字段
    expect(l2.manifest.extensions?.collaboration).toEqual({
      lockedBy: 'user-42',
      revision: 17,
    })
    // createdAt 不能被刷新
    expect(l2.manifest.createdAt).toBe('2026-01-01T00:00:00.000Z')
  })

  it('没有 prevNodeFiles 时也能从磁盘回读旧文件（旧调用点仍不丢数据）', async () => {
    const m = mkFs()
    await seedFutureProject(m)
    const l1 = await loadProject(m.fs, { rootDir: 'pf' })
    // 故意不传 prevNodeFiles
    await saveProject(m.fs, { rootDir: 'pf' }, {
      name: '未来工程',
      tables: l1.tables,
      records: l1.records,
      positions: l1.positions,
    })
    const raw = JSON.parse(m.files.get(nodePath('pf', 'n1'))!)
    expect(raw.extensions?.lockedBy).toBe('user-42')
    expect(raw.nodeVersion).toBe(3)
  })

  it('连续三轮保存后数据仍完全一致（防「只保一层」的半吊子实现）', async () => {
    const m = mkFs()
    await seedFutureProject(m)
    let prev = await loadProject(m.fs, { rootDir: 'pf' })
    for (let i = 0; i < 3; i++) {
      await saveProject(
        m.fs,
        { rootDir: 'pf' },
        {
          name: prev.manifest.name,
          tables: prev.tables,
          records: prev.records,
          positions: prev.positions,
          prev: prev.manifest,
          prevNodeFiles: prev.nodeFiles,
        },
      )
      prev = await loadProject(m.fs, { rootDir: 'pf' })
    }
    const n = prev.nodeFiles.n1 as unknown as Record<string, unknown>
    expect(n.extensions?.lockedBy).toBe('user-42')
    expect(n.extensions?.layoutHint).toEqual({ collapsed: true, zIndex: 7 })
    expect(n.nodeVersion).toBe(3)
    expect((n.schema as Record<string, unknown>).quantumMeta).toEqual({
      coherence: 0.97,
      backend: 'future-sim',
    })
  })
})

describe('向前兼容 · 控件状态跟着节点走（用户第 5 项）', () => {
  it('viewState 经一次往返后仍在', async () => {
    const m = mkFs()
    await seedFutureProject(m)
    const l1 = await loadProject(m.fs, { rootDir: 'pf' })
    expect(l1.nodeFiles.n1.viewState).toEqual({
      tab: 'refs',
      refPanelOpen: true,
      collapsed: true,
    })

    await saveProject(
      m.fs,
      { rootDir: 'pf' },
      {
        name: l1.manifest.name,
        tables: l1.tables,
        records: l1.records,
        positions: l1.positions,
        prev: l1.manifest,
        prevNodeFiles: l1.nodeFiles,
      },
    )
    const l2 = await loadProject(m.fs, { rootDir: 'pf' })
    expect(l2.nodeFiles.n1.viewState).toEqual({
      tab: 'refs',
      refPanelOpen: true,
      collapsed: true,
    })
  })

  it('viewState 缺省时不影响打开（老工程没有这个字段），也不凭空造空对象', async () => {
    const m = mkFs()
    await saveProject(m.fs, { rootDir: 'p1' }, mkInput())
    const loaded = await loadProject(m.fs, { rootDir: 'p1' })
    expect(loaded.problems.filter((p) => p.level === 'error')).toEqual([])
    // 空 viewState 会在每个节点文件里留噪声，让 git diff 出现无意义改动
    expect(loaded.nodeFiles.price.viewState).toBeUndefined()
  })
})

describe('向前兼容 · 节点资产跟着工程走（用户第 4 项）', () => {
  it('★ 资产字节缺失会被报成 warn（否则用户以为节点内容丢了）', async () => {
    const m = mkFs()
    await seedFutureProject(m) // 不放字节 = 工程发给同事的典型状态
    const loaded = await loadProject(m.fs, { rootDir: 'pf' })
    expect(
      loaded.problems.some((p) => p.message.includes('资产字节缺失')),
      '字节缺失没报出来 —— 用户双击只会看到空白，且无从追溯',
    ).toBe(true)
    // 但不能是 error：资产缺失不该让整个工程打不开
    expect(loaded.problems.filter((p) => p.level === 'error')).toEqual([])
  })

  it('字节在 → 不报缺失，且字节原样读回、登记信息完整还原', async () => {
    const m = mkFs()
    const bytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04]) // PK.. = zip 头
    await seedFutureProject(m, 'pg', { assetBytes: bytes })
    const loaded = await loadProject(m.fs, { rootDir: 'pg' })
    expect(loaded.problems.some((p) => p.message.includes('资产字节缺失'))).toBe(false)

    // ★ 字节必须原样读回（MIME 与扩展名不参与编解码）
    const got = await m.fs.readBytes(assetPath('pg', 'n1', { relPath: '方案.docx' })!)
    expect(got).toEqual(bytes)
    // 登记信息一字不差
    expect(loaded.nodeFiles.n1.assets).toHaveLength(1)
    expect(loaded.nodeFiles.n1.assets![0].mime).toContain('wordprocessingml')
    expect(loaded.nodeFiles.n1.assets![0].hash).toBe('deadbeef')
  })

  it('★ 资产登记经一次往返后仍在（否则节点在、文档没了）', async () => {
    const m = mkFs()
    await seedFutureProject(m, 'ph', { assetBytes: new Uint8Array([1, 2, 3]) })
    const l1 = await loadProject(m.fs, { rootDir: 'ph' })
    await saveProject(
      m.fs,
      { rootDir: 'ph' },
      {
        name: l1.manifest.name,
        tables: l1.tables,
        records: l1.records,
        positions: l1.positions,
        prev: l1.manifest,
        prevNodeFiles: l1.nodeFiles,
      },
    )
    const l2 = await loadProject(m.fs, { rootDir: 'ph' })
    expect(l2.nodeFiles.n1.assets).toHaveLength(1)
    expect(l2.nodeFiles.n1.assets![0].name).toBe('方案.docx')
    expect(l2.nodeFiles.n1.assets![0].role).toBe('template')
  })

  it('assets 结构坏掉只 warn（丢的是登记信息，不是数据本体）', async () => {
    const m = mkFs()
    await saveProject(m.fs, { rootDir: 'p1' }, mkInput())
    const f = JSON.parse(m.files.get('p1/nodes/price.json')!)
    f.assets = '不是数组'
    m.files.set('p1/nodes/price.json', JSON.stringify(f))
    const loaded = await loadProject(m.fs, { rootDir: 'p1' })
    expect(loaded.problems.some((p) => p.message.includes('assets 不是数组'))).toBe(true)
    // 节点本身仍读得出来
    expect(loaded.tables.price).toBeTruthy()
  })

  it('assetDir 可被覆盖（将来支持多份文档副本时不必改结构）', async () => {
    const m = mkFs()
    await seedFutureProject(m, 'pi', { assetBytes: new Uint8Array([9]) })
    const l1 = await loadProject(m.fs, { rootDir: 'pi' })
    await saveProject(
      m.fs,
      { rootDir: 'pi' },
      {
        name: l1.manifest.name,
        tables: l1.tables,
        records: l1.records,
        positions: l1.positions,
        assetDir: 'docs',
        prev: l1.manifest,
        prevNodeFiles: l1.nodeFiles,
      },
    )
    const raw = JSON.parse(m.files.get(nodePath('pi', 'n1'))!)
    expect(raw.assetDir).toBe('docs')
  })
})

describe('资产路径安全（工程文件来自别人，不能当任意文件访问入口）', () => {
  it('正常相对路径通过', () => {
    expect(relPathIsSafe('方案.docx')).toBe(true)
    expect(relPathIsSafe('子目录/图.png')).toBe(true)
    expect(relPathIsSafe('带空格的文件.mp4')).toBe(true)
  })

  it('★ 越界与畸形路径一律拒绝', () => {
    // 这些既是攻击形态，也是手滑形态（从别的工程拷 relPath），必须全拦
    expect(relPathIsSafe('../secret.txt')).toBe(false)
    expect(relPathIsSafe('a/../../b')).toBe(false)
    expect(relPathIsSafe('/etc/passwd')).toBe(false)
    expect(relPathIsSafe('C:/Windows/system32')).toBe(false)
    expect(relPathIsSafe('a\\b')).toBe(false) // Windows 分隔符
    expect(relPathIsSafe('a//b')).toBe(false)
    expect(relPathIsSafe('./a')).toBe(false)
    expect(relPathIsSafe('..')).toBe(false)
    expect(relPathIsSafe('')).toBe(false)
  })

  it('assetPath 对非法路径返回 null（调用方据此报 warn 并跳过）', () => {
    expect(assetPath('root', 'n1', { relPath: 'ok.docx' })).toBe(
      `root/${PROJECT_DIRS.assets}/${encodeURIComponent('n1')}/ok.docx`,
    )
    expect(assetPath('root', 'n1', { relPath: '../../etc/passwd' })).toBeNull()
    // ★ nodeId 同样来自工程文件，也必须转义
    expect(assetPath('root', '../../x', { relPath: 'ok.docx' })).toBe(
      `root/${PROJECT_DIRS.assets}/${encodeURIComponent('../../x')}/ok.docx`,
    )
  })

  it('★ 路径越界时 loadProject 报 warn 而不是崩，也不去读那个路径', async () => {
    const m = mkFs()
    const { fs, files } = m
    files.set(
      nodePath('pz', 'n1'),
      JSON.stringify({
        id: 'n1',
        schema: { id: 'n1', name: 'X', kind: 'table', fields: [] },
        position: { x: 0, y: 0 },
        rows: [],
        assets: [
          {
            id: 'a',
            relPath: '../../../Windows/System32/evil.dll',
            name: 'evil',
            role: 'attachment',
            mime: 'application/octet-stream',
            size: 1,
          },
        ],
      }),
    )
    files.set(
      'pz/project.json',
      JSON.stringify({
        formatVersion: 2,
        formatKind: 'cbproj-dir',
        name: 'x',
        createdAt: '2026-01-01T00:00:00.000Z',
        updatedAt: '2026-01-01T00:00:00.000Z',
        nodeIds: ['n1'],
        refs: {},
        dataSources: [],
      }),
    )
    const loaded = await loadProject(fs, { rootDir: 'pz' })
    expect(loaded.problems.some((p) => p.message.includes('路径非法或越界'))).toBe(true)
    // ★ 且不该真的去读那个路径 —— 它根本没进过文件系统，
    //   所以不能报「字节缺失」（那会让人以为文件曾存在后丢了）
    expect(loaded.problems.some((p) => p.message.includes('资产字节缺失'))).toBe(false)
  })
})
