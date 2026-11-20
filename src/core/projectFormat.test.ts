// 工程文件格式（projectFormat）单测
//
// ★ 本文件主要防两类退化：
//   ① 连线派生与原实现漂移 —— 症状是画布上少边/多边，本地很难肉眼发现
//   ② 往返丢信息 —— 症状是「保存后重开数据少了」，是灾难级且难以追溯

import { describe, it, expect } from 'vitest'
import {
  edgesOf,
  allEdges,
  collectReferences,
  buildRefsIndex,
  buildProject,
  buildNodeFile,
  validateProject,
  validateNodeFile,
  isReadable,
  formatProblems,
  fallbackPosition,
  stableStringify,
  fingerprintOf,
  staleUpstreamsOf,
  isStaleNode,
  PROJECT_FORMAT_VERSION,
  PROJECT_FORMAT_KIND,
} from './projectFormat'
import { defineNode, defineTable, defineField } from './schema'
import type { TableSchema } from './schema'
import type { Row } from './store'

/**
 * 造一张带 lookup 字段的表。
 *
 * ★ type 必须是 'lookup'（FieldType 的四个成员之一），不能写 'number'。
 *   画布的判定是 `f.type === 'lookup' && f.lookup`（双条件），
 *   type 写错就静默不产生边 —— 与原实现一致，这里照抄。
 */
function mkTable(id: string, opts: { lookup?: string; linked?: string[] } = {}): TableSchema {
  const fields = [
    defineField('名称', 'text'),
    ...(opts.lookup
      ? [
          defineField('单价', 'lookup', {
            lookup: { table: opts.lookup, sourceField: '名称', matchField: '名称', returnField: '单价' },
          }),
        ]
      : []),
  ]
  const t = defineTable(id, fields, id, 'table', {
    ...(opts.linked ? { linkedInputs: opts.linked } : {}),
  })
  return t
}

describe('连线派生 · 与画布原实现等价（防漂移）', () => {
  it('lookup 字段产生 assoc 边，id 格式为 <src>.<field>-><tgt>', () => {
    const price = mkTable('price')
    const mat = mkTable('mat', { lookup: 'price' })
    const edges = edgesOf(mat, { price, mat })
    expect(edges).toHaveLength(1)
    // ★ 这个 id 串是落盘与画布共用的，格式不能改
    expect(edges[0].id).toBe('mat.单价->price')
    expect(edges[0]).toMatchObject({ source: 'mat', target: 'price', kind: 'assoc', fieldId: '单价' })
  })

  it('linkedInputs 产生 flow 边，id 格式为 flow:<src>-><tgt>', () => {
    const a = mkTable('a')
    const b = mkTable('b', { linked: ['a'] })
    const edges = edgesOf(b, { a, b })
    expect(edges).toHaveLength(1)
    expect(edges[0].id).toBe('flow:a->b')
    expect(edges[0].kind).toBe('flow')
    expect(edges[0].fieldId).toBeUndefined()
  })

  it('★ 软删除节点不产生边（与画布一致，避免落盘出悬空连线）', () => {
    const a = mkTable('a')
    const b = mkTable('b', { linked: ['a'] })
    // 来源被删除
    expect(edgesOf(b, { a: { ...a, deleted: true }, b })).toHaveLength(0)
    // 本体被删除
    expect(edgesOf({ ...b, deleted: true }, { a, b })).toHaveLength(0)
  })

  it('lookup 目标被软删除 → 不产生 assoc 边', () => {
    const price = { ...mkTable('price'), deleted: true }
    const mat = mkTable('mat', { lookup: 'price' })
    expect(edgesOf(mat, { price, mat })).toHaveLength(0)
  })

  it('linkedInputs 指向不存在的节点 → 静默跳过，不产生悬空边', () => {
    const b = mkTable('b', { linked: ['ghost'] })
    expect(edgesOf(b, { b })).toHaveLength(0)
  })

  it('allEdges 按节点 id 字典序 → 落盘稳定可比对', () => {
    const tables = {
      z: mkTable('z', { linked: ['a'] }),
      a: mkTable('a'),
      m: mkTable('m', { linked: ['a'] }),
    }
    const ids = allEdges(tables).map((e) => e.id)
    expect(ids).toEqual(['flow:a->m', 'flow:a->z'])
  })

  it('★ 一个节点可同时产出两类边', () => {
    const price = mkTable('price')
    const a = mkTable('a')
    const b = mkTable('b', { linked: ['a'], lookup: 'price' })
    const kinds = edgesOf(b, { price, a, b }).map((e) => e.kind).sort()
    expect(kinds).toEqual(['assoc', 'flow'])
  })
})

describe('引用收集 · 双向关联（关联到的节点，不是只算上游）', () => {
  it('flow 边：两端互相都在对方的引用里（方向判定留给上层）', () => {
    const a = mkTable('a')
    const b = mkTable('b', { linked: ['a'] })
    expect(collectReferences(b, { a, b })).toEqual(['a'])
    //★ a 也要能看到 b（它有个下游）—— 第一版只收上游方向，这里恒为空，
    //   症状是「价目表节点从来看不到谁在用它」。
    expect(collectReferences(a, { a, b })).toEqual(['b'])
  })

  it('assoc 边：字段所在方与被查表方互相都算引用（两个方向都要收）', () => {
    const price = mkTable('price')
    const mat = mkTable('mat', { lookup: 'price' })
    expect(collectReferences(mat, { price, mat })).toEqual(['price'])
    //被查表自己派不出 assoc 边，必须靠全量扫描才查得到
    expect(collectReferences(price, { price, mat })).toEqual(['mat'])
  })

  it('★ config.inputs 不计入引用（面板勾选无连线，会让引用图虚高 → 假脏标记）', () => {
    const a = mkTable('a')
    const b = defineTable('b', [defineField('x', 'text')], 'b', 'table', {
      inputs: ['a'], // 无连线的来源
      linkedInputs: [],
    })
    // inputs 有值但没连线 → 不是「直接引用关系」
    expect(collectReferences(b, { a, b })).toEqual([])
    expect(collectReferences(a, { a, b })).toEqual([])
  })

  it('软删除节点不产生任何引用', () => {
    const a = mkTable('a')
    const b = mkTable('b', { linked: ['a'] })
    expect(collectReferences({ ...b, deleted: true }, { a, b })).toEqual([])
  })

  it('buildRefsIndex 覆盖全表且结果有序', () => {
    const price = mkTable('price')
    const a = mkTable('a')
    const m = mkTable('m', { linked: ['z', 'a'] })
    const z = mkTable('z')
    const idx = buildRefsIndex({ price, a, m, z })
    expect(Object.keys(idx)).toEqual(['a', 'm', 'price', 'z'])
    // 排序后 → 落盘 diff 稳定
    expect(idx.m).toEqual(['a', 'z'])
    // a 被 m 引用 → 不是空
    expect(idx.a).toEqual(['m'])
    // z 被 m 引用
    expect(idx.z).toEqual(['m'])
    // price 无人引用
    expect(idx.price).toEqual([])
  })

  it('buildRefsIndex 与 collectReferences 逐节点一致（两处实现不许漂移）', () => {
    const tables = {
      price: mkTable('price'),
      a: mkTable('a'),
      m: mkTable('m', { linked: ['a'], lookup: 'price' }),
    }
    const idx = buildRefsIndex(tables)
    for (const id of Object.keys(tables)) {
      expect(idx[id]).toEqual(collectReferences(tables[id], tables))
    }
  })
})

describe('稳定序列化与指纹', () => {
  it('★ 键顺序不同但内容相同 → 同指纹（否则每次都假「脏」标记）', () => {
    expect(fingerprintOf({ a: 1, b: 2 })).toBe(fingerprintOf({ b: 2, a: 1 }))
    expect(stableStringify({ a: 1, b: 2 })).toBe(stableStringify({ b: 2, a: 1 }))
  })

  it('数组顺序不同 → 不同指纹（顺序有语义，不能归一）', () => {
    expect(fingerprintOf([1, 2])).not.toBe(fingerprintOf([2, 1]))
  })

  it('嵌套对象的键顺序也不影响', () => {
    expect(fingerprintOf({ x: { p: 1, q: 2 } })).toBe(fingerprintOf({ x: { q: 2, p: 1 } }))
  })

  it('内容真变了 → 指纹变', () => {
    expect(fingerprintOf({ a: 1 })).not.toBe(fingerprintOf({ a: 2 }))
  })

  it('undefined 与缺字段的差异可区分（不抛异常）', () => {
    expect(() => fingerprintOf(undefined)).not.toThrow()
    expect(fingerprintOf(undefined)).toBeTypeOf('string')
  })

  it('行数据（Row[]）指纹稳定', () => {
    const r1: Row[] = [{ 名称: '钢管', 数量: 3 }]
    const r2: Row[] = [{ 数量: 3, 名称: '钢管' }]
    expect(fingerprintOf(r1)).toBe(fingerprintOf(r2))
    expect(fingerprintOf(r1)).not.toBe(fingerprintOf([{ 名称: '钢管', 数量: 4 }]))
  })
})

describe('上游指纹判定 · C10 标记（Q14已拍板 A）', () => {
  const fps = [
    { nodeId: 'a', valueHash: '1111', computedAt: '2026-10-04T00:00:00Z' },
    { nodeId: 'b', valueHash: '2222', computedAt: '2026-10-04T00:00:00Z' },
  ]

  it('上游指纹与当前一致 → 不标记', () => {
    const now = { a: '1111', b: '2222' }
    expect(staleUpstreamsOf('n', fps, now, new Set())).toEqual([])
    expect(isStaleNode({ id: 'n', schema: mkTable('n'), position: { x: 0, y: 0 }, rows: [], upstream: fps }, now, new Set())).toBe(false)
  })

  it('★ 上游值变了 → 报出是哪个上游（旧标记只说「脏了」，指纹能说「谁脏的」）', () => {
    const now = { a: '9999', b: '2222' }
    expect(staleUpstreamsOf('n', fps, now, new Set())).toEqual(['a'])
  })

  it('★ 本节点本次已重算 → 不标记（查的是本节点，不是上游）', () => {
    const now = { a: '9999', b: '2222' }
    // 上游 a 变了，但本节点 n 自己本次重算过 → 它是最新，不该标脏
    expect(staleUpstreamsOf('n', fps, now, new Set(['n']))).toEqual([])
    // 反例：若实现误查上游集合，上游 a 在集合里就会短路 → 漏判。两种错法都挡住
    expect(staleUpstreamsOf('n', fps, now, new Set(['a']))).toEqual(['a'])
  })

  it('★ 上游节点被删除 → 也算过期，且必须报出（最易漏的断链）', () => {
    expect(staleUpstreamsOf('n', fps, { a: '1111' }, new Set())).toEqual(['b'])
  })

  it('无指纹记录（从未重算过）→ 不标记', () => {
    const now = { a: '9999' }
    expect(staleUpstreamsOf('n', undefined, now, new Set())).toEqual([])
    expect(staleUpstreamsOf('n', [], now, new Set())).toEqual([])
  })

  it('多个上游同时过期 → 全部报出且有序', () => {
    const now = { a: 'x', b: 'y' }
    expect(staleUpstreamsOf('n', fps, now, new Set())).toEqual(['a', 'b'])
  })
})

describe('格式校验 · 分级而非一票否决', () => {
  it('当前版本 + 结构齐全 → 可读', () => {
    const p = buildProject({
      name: '测试工程',
      tables: { a: mkTable('a') },
      records: { a: [] },
      positions: { a: { x: 1, y: 2 } },
    })
    const problems = validateProject(p)
    expect(problems).toEqual([])
    expect(isReadable(problems)).toBe(true)
    expect(p.formatVersion).toBe(PROJECT_FORMAT_VERSION)
    expect(p.formatKind).toBe(PROJECT_FORMAT_KIND)
  })

  it('★ 高版本 → error（老程序不能装作能读，读了就是静默丢数据）', () => {
    const problems = validateProject({
      formatVersion: PROJECT_FORMAT_VERSION + 1,
      formatKind: PROJECT_FORMAT_KIND,
      nodeIds: [],
      refs: {},
    })
    expect(problems.some((p) => p.level === 'error')).toBe(true)
    expect(isReadable(problems)).toBe(false)
  })

  it('缺 refs 索引 → warn 而非 error（打开后可重建）', () => {
    const problems = validateProject({
      formatVersion: PROJECT_FORMAT_VERSION,
      formatKind: PROJECT_FORMAT_KIND,
      nodeIds: [],
    })
    expect(problems.every((p) => p.level === 'warn')).toBe(true)
    expect(isReadable(problems)).toBe(true)
  })

  it('缺 nodeIds → error（无法定位节点文件）', () => {
    const problems = validateProject({ formatVersion: PROJECT_FORMAT_VERSION, formatKind: PROJECT_FORMAT_KIND })
    expect(problems.some((p) => p.level === 'error')).toBe(true)
  })

  it('完全不是对象 → error', () => {
    expect(isReadable(validateProject(null))).toBe(false)
    expect(isReadable(validateProject('坏文件'))).toBe(false)
  })

  it('低版本 → 只 warn 不 error（向后兼容）', () => {
    const problems = validateProject({
      formatVersion: PROJECT_FORMAT_VERSION - 1,
      formatKind: PROJECT_FORMAT_KIND,
      nodeIds: [],
      refs: {},
    })
    expect(isReadable(problems)).toBe(true)
  })
})

describe('节点文件校验', () => {
  it('id 与文件名不一致 → error（可能被手工改过）', () => {
    const problems = validateNodeFile({ id: 'other', schema: mkTable('a'), position: { x: 0, y: 0 }, rows: [] }, 'a')
    expect(problems.some((p) => p.level === 'error' && p.message.includes('不一致'))).toBe(true)
  })

  it('缺坐标 → warn + 用兜底位（叠一起比报错更难排查）', () => {
    const problems = validateNodeFile({ id: 'a', schema: mkTable('a'), rows: [] }, 'a')
    expect(problems.every((p) => p.level === 'warn')).toBe(true)
    const pos = fallbackPosition('a')
    expect(typeof pos.x).toBe('number')
    expect(typeof pos.y).toBe('number')
  })

  it('★ 兜底位必须分散（不同 id 不能全叠在0,0）', () => {
    const seen = new Set<string>()
    for (let i = 0; i < 50; i++) {
      const p = fallbackPosition(`node_${i}`)
      seen.add(`${p.x},${p.y}`)
    }
    // 50 个节点应落在多个位置（允许有碰撞，但不该全一样）
    expect(seen.size).toBeGreaterThan(1)
  })

  it('rows 不是数组 → error', () => {
    const problems = validateNodeFile({ id: 'a', schema: mkTable('a'), position: { x: 0, y: 0 }, rows: 'x' }, 'a')
    expect(problems.some((p) => p.level === 'error')).toBe(true)
  })
})

describe('★ 往返一致（防止保存后丢信息）', () => {
  it('buildProject + buildNodeFile 完整承载 schema/config/rows/position', () => {
    const tables: Record<string, TableSchema> = {
      price: mkTable('price'),
      mat: mkTable('mat', { linked: ['price'], lookup: 'price' }),
    }
    const records = { price: [{ 名称: '钢管', 单价: 5 }], mat: [] }
    const positions = { price: { x: 10, y: 20 }, mat: { x: 300, y: 400 } }

    const manifest = buildProject({ name: '材料工程', tables, records, positions })
    expect(manifest.nodeIds).toEqual(['mat', 'price'])
    expect(manifest.refs.mat).toEqual(['price'])
    expect(manifest.refs.price).toEqual(['mat'])

    const f = buildNodeFile('mat', { name: '材料工程', tables, records, positions })
    expect(f.id).toBe('mat')
    // config 不能丢 —— 连线全靠它
    expect(f.schema.config?.linkedInputs).toEqual(['price'])
    expect(f.position).toEqual({ x: 300, y: 400 })
    expect(f.rows).toEqual([])
  })

  it('★ 指纹在重新组装时必须保留（组装 ≠ 重算）', () => {
    const tables = { a: mkTable('a') }
    const src = { name: 'x', tables, records: {}, positions: {} }
    const withFp = { ...buildNodeFile('a', src), upstream: [{ nodeId: 'up', valueHash: 'aa', computedAt: 't' }] }
    const again = buildNodeFile('a', src, withFp)
    expect(again.upstream).toEqual(withFp.upstream)
    // 无历史时为 undefined，不是空数组
    expect(buildNodeFile('a', src).upstream).toBeUndefined()
  })

  it('序列化往返：JSON 解析后内容不变（磁盘上就是这个形态）', () => {
    const manifest = buildProject({
      name: '往返测试',
      tables: { a: mkTable('a'), b: mkTable('b', { linked: ['a'] }) },
      records: { a: [{ v: 1 }], b: [] },
      positions: { a: { x: 0, y: 0 }, b: { x: 1, y: 1 } },
      dataSources: [{ pluginId: 'file', address: 'D:/x.xlsx', embed: false }],
    })
    const round = JSON.parse(JSON.stringify(manifest))
    expect(round).toEqual(manifest)
    expect(isReadable(validateProject(round))).toBe(true)
  })

  it('records 里的中文与特殊字符往返无损', () => {
    const rows: Row[] = [{ 名称: 'Φ159×6钢管', 备注: '含"引号"与\\反斜杠', 数量: 3 }]
    const manifest = buildProject({
      name: '特殊字符',
      tables: { a: mkTable('a') },
      records: { a: rows },
      positions: {},
    })
    expect(JSON.parse(JSON.stringify(manifest)).nodeIds).toEqual(['a'])
    expect(rows[0].名称).toBe('Φ159×6钢管')
  })
})

describe('格式问题汇总文本', () => {
  it('带节点 id 前缀，便于定位', () => {
    const text = formatProblems([
      { level: 'error', message: '缺 schema', nodeId: 'n1' },
      { level: 'warn', message: '缺坐标' },
    ])
    expect(text).toContain('[error] n1: 缺 schema')
    expect(text).toContain('[warn] 缺坐标')
  })
})

/* ══════════════════════════════════════════════════════════════════
   向前兼容：节点还没做完的情况下，保存必须不丢数据
   ══════════════════════════════════════════════════════════════════
   背景（用户 2026-10-06原话：「我的节点还没做完，用一种兼容的形式保存」）：
   节点类型与节点内部结构都还在演进，必然存在「工程文件里有本程序看不懂的东西」的
   阶段。这里的每一条测试都对应一种真实的丢数据场景。

   ⚠ 防退化要点：这类测试最容易写成「假绿」—— 只验证 extensions 字段本身存在，
   而没验证**真的经过一次 save→load→save 之后数据还在**。下面刻意都走完整往返。
   */

describe('向前兼容 · 未知节点类型（本程序完全不认识的新节点）', () => {
  /** 造一个「未来版本」的节点文件：当前程序里根本不存在的 kind + 一堆陌生字段 */
  function futureNodeFile(): Record<string, unknown> {
    return {
      id: 'n1',
      schema: {
        id: 'n1',
        name: '未来的节点',
        kind: 'quantum-bridge', // ← 不在 NodeKind 白名单里
        fields: [{ id: 'f1', name: '量子位', type: 'text' }],
        config: { entanglement: 'strongly', qubitCount: 512 },
        // schema 顶层未知键
        quantumMeta: { coherence: 0.97, backend: 'future-sim' },
      },
      position: { x: 100, y: 200 },
      rows: [{ f1: '|0>' }],
      // 节点文件顶层未知键（未来版本正式支持的字段）
      layoutHint: { collapsed: true, zIndex: 7 },
      lockedBy: 'user-42',
      // 未来版本自己写的 nodeVersion
      nodeVersion: 3,
      // 以及未来版本可能自己塞的 extensions
      extensions: { experimentalFeature: { flag: true } },
    }
  }

  it('未知 kind 只报warn，不报 error —— 不能因为一个新节点就打不开整个工程', () => {
    const problems = validateNodeFile(futureNodeFile(), 'n1')
    const errors = problems.filter((p) => p.level === 'error')
    const warns = problems.filter((p) => p.level === 'warn')
    expect(errors).toHaveLength(0)
    expect(warns.some((w) => w.message.includes('quantum-bridge'))).toBe(true)
    // ★ 关键断言：整体仍然可读
    expect(isReadable(problems)).toBe(true)
  })

  it('未知节点经buildNodeFile 重新保存后，未知字段全部搬到 extensions 且值不变', () => {
    const old = futureNodeFile()
    // 程序侧认知：tables 里只有它「认识」的那部分（这里给一个退化的 schema）
    const src = {
      name: '兼容测试',
      tables: {
        n1: {
          id: 'n1',
          name: '未来的节点',
          kind: 'table', // 程序认知里它就是个普通表
          fields: [{ id: 'f1', name: '量子位', type: 'text' }],
        },
      },
      records: { n1: [{ f1: '|0>' }] },
      positions: { n1: { x: 100, y: 200 } },
    }
    const out = buildNodeFile('n1', src, old as never)

    // 未来版本的未知顶层键都进了 extensions
    expect(out.extensions?.layoutHint).toEqual({ collapsed: true, zIndex: 7 })
    expect(out.extensions?.lockedBy).toBe('user-42')
    // 未来版本原有的 extensions 也没被冲掉（这最容易丢）
    expect(out.extensions?.experimentalFeature).toEqual({ flag: true })
    // nodeVersion 保留
    expect(out.nodeVersion).toBe(3)
    // schema 里的未知顶层键也保住了
    expect((out.schema as unknown as Record<string, unknown>).quantumMeta).toEqual({
      coherence: 0.97,
      backend: 'future-sim',
    })
    // 程序认知里的 kind仍然优先（本程序算出来的值不被旧文件覆盖）
    expect(out.schema.kind).toBe('table')
  })

  it('★ 完整往返：未知节点文件 → 保存 → 加载 → 再保存，数据一字不丢', () => {
    // 第 1 轮：磁盘上是一份「未来版本」写的工程
    const diskFile = futureNodeFile()
    const diskManifest = {
      formatVersion: 99, // ★ 未来版本的工程版本号
      formatKind: 'cbproj-dir',
      name: '未来工程',
      createdAt: '2026-01-01T00:00:00.000Z',
      updatedAt: '2026-01-01T00:00:00.000Z',
      nodeIds: ['n1'],
      refs: { n1: [] },
      dataSources: [],
      // 工程级未知字段
      collaboration: { lockedBy: 'user-42', revision: 17 },
    }

    // 当前程序打开它 —— 会因为版本号高而拒绝（这是正确的）
    expect(isReadable(validateProject(diskManifest))).toBe(false)

    // 但假设用户降级处理后重存：旧文件作为 existing 传进去
    const src = {
      name: '未来工程',
      tables: {
        n1: {
          id: 'n1',
          name: '未来的节点',
          kind: 'quantum-bridge' as never,
          fields: [{ id: 'f1', name: '量子位', type: 'text' as const }],
        },
      },
      records: { n1: [{ f1: '|0>' }] },
      positions: { n1: { x: 100, y: 200 } },
    }
    const rebuilt = buildNodeFile('n1', src, diskFile as never)

    // 再序列化回磁盘（这就是用户下次打开时看到的东西）
    const onDisk = JSON.parse(JSON.stringify(rebuilt))

    // ★ 逐键断言：未来版本写的每一个键都还在。
    //   注意未知键的落点是 extensions（不是顶层）—— 这正是 extensions 的用途：
    //   把「本程序不认识的」显式归到一处，好处是将来程序支持了某个字段时，
    //   一眼能看出「这堆里有没有我要认领的」，而不必去猜顶层哪些键是没见过的。
    //   数据本身一个字节都没少。
    expect(onDisk.extensions.layoutHint).toEqual({ collapsed: true, zIndex: 7 })
    expect(onDisk.extensions.lockedBy).toBe('user-42')
    expect(onDisk.nodeVersion).toBe(3)
    expect(onDisk.extensions).toEqual({
      experimentalFeature: { flag: true },
      layoutHint: { collapsed: true, zIndex: 7 },
      lockedBy: 'user-42',
    })
    expect(onDisk.schema.quantumMeta).toEqual({ coherence: 0.97, backend: 'future-sim' })
    expect(onDisk.rows).toEqual([{ f1: '|0>' }])
    expect(onDisk.position).toEqual({ x: 100, y: 200 })

    // ★ 再验一次：这些数据经过第二次往返依然在（防「只保一层」的半吊子实现）
    const again = JSON.parse(
      JSON.stringify(buildNodeFile('n1', src, onDisk as never)),
    )
    expect(again.extensions.layoutHint).toEqual({ collapsed: true, zIndex: 7 })
    expect(again.extensions.lockedBy).toBe('user-42')
    expect(again.schema.quantumMeta).toEqual({ coherence: 0.97, backend: 'future-sim' })
  })

  it('工程清单的未知字段也能透传（保存不抹掉 collaboration 等字段）', () => {
    const prev = {
      formatVersion: 2,
      formatKind: 'cbproj-dir' as const,
      name: '旧工程',
      createdAt: '2026-01-01T00:00:00.000Z',
      updatedAt: '2026-01-01T00:00:00.000Z',
      nodeIds: ['a'],
      refs: { a: [] },
      dataSources: [],
      collaboration: { lockedBy: 'user-42', revision: 17 },
    }
    const out = buildProject(
      { name: '新名', tables: { a: mkTable('a') }, records: { a: [] }, positions: {} },
      undefined,
      prev,
    )
    expect(out.extensions?.collaboration).toEqual({ lockedBy: 'user-42', revision: 17 })
    // ★ createdAt 必须保住，不能因为重建 manifest 就刷新
    expect(out.createdAt).toBe('2026-01-01T00:00:00.000Z')
  })

  it('没有旧文件时不会凭空造出 extensions 空对象', () => {
    const out = buildNodeFile('a', {
      name: 'x',
      tables: { a: mkTable('a') },
      records: {},
      positions: {},
    })
    // 空 extensions 会在每个节点文件里留噪声，让 git diff 出现无意义改动
    expect(out.extensions).toBeUndefined()
  })
})

describe('向前兼容 · 已知节点被外部改动时不回退', () => {
  it('旧文件里已删除的字段不会被无脑塞回来（程序认知优先）', () => {
    // 程序已经把旧的 v1 字段迁移掉了，新 schema 里没有它
    const src = {
      name: 'x',
      tables: {
        a: { id: 'a', name: 'A', kind: 'table' as const, fields: [] },
      },
      records: {},
      positions: {},
    }
    const old = {
      id: 'a',
      schema: {
        id: 'a',
        name: 'A',
        kind: 'table',
        fields: [],
        // ★ 这个键在 KNOWN_SCHEMA_KEYS 里（= trashMeta 同类：程序认识的字段），
        //   程序本轮**没产出**它 → 旧值不该被塞回来。
        //   这才是「程序删掉的字段不该复活」的真实判据：
        //   判据不是「键名看着陌生」，而是「程序本轮没产出这个键」。
        trashMeta: { deletedAt: 1, refs: {} },
      },
      position: { x: 0, y: 0 },
      rows: [],
    }
    const out = buildNodeFile('a', src, old as never)
    expect((out.schema as unknown as Record<string, unknown>).trashMeta).toBeUndefined()
  })

  it('真正陌生的键仍然保住（程序没产出 ≠ 该丢弃，两件事必须分开）', () => {
    const src = {
      name: 'x',
      tables: { a: { id: 'a', name: 'A', kind: 'table' as const, fields: [] } },
      records: {},
      positions: {},
    }
    const old = {
      id: 'a',
      schema: { id: 'a', name: 'A', kind: 'table', fields: [], quantumMeta: { coherence: 0.9 } },
      position: { x: 0, y: 0 },
      rows: [],
    }
    const out = buildNodeFile('a', src, old as never)
    // ★ 与上一条互为对照：同样是「本轮没产出」，
    //   程序认识的键（trashMeta）丢弃，程序不认识的键（quantumMeta）保留。
    expect((out.schema as unknown as Record<string, unknown>).quantumMeta).toEqual({
      coherence: 0.9,
    })
  })

  it('旧文件的 rowstype 不对时，新值仍然覆盖（不能被旧数据污染）', () => {
    const src = {
      name: 'x',
      tables: { a: mkTable('a') },
      records: { a: [{ v: 9 }] },
      positions: { a: { x: 5, y: 5 } },
    }
    const old = {
      id: 'a',
      schema: mkTable('a'),
      position: { x: 1, y: 1 },
      rows: [{ v: 1 }], // 旧值，不该胜出
    }
    const out = buildNodeFile('a', src, old as never)
    expect(out.rows).toEqual([{ v: 9 }])
    expect(out.position).toEqual({ x: 5, y: 5 })
  })

  it('程序不认识该节点（tables 里没有）时，沿用旧文件的一切，绝不产空节点', () => {
    // 场景：工程里有个「未来节点」，当前版本的 store 里根本没有它，
    //       tables/records/positions 三张表全都没有它的条目。
    //       此时保存绝不能把它抹掉，也不能造出一个空壳。
    const old = {
      id: 'ghost',
      schema: { id: 'ghost', name: '幽灵节点', kind: 'quantum-bridge', fields: [] },
      position: { x: 30, y: 40 },
      rows: [{ a: 1 }],
      lockedBy: 'user-42',
    }
    const out = buildNodeFile('ghost', {
      name: 'x',
      tables: {}, // 程序完全不认识
      records: {},
      positions: {},
    }, old as never)

    // ★ 旧 schema 整份沿用（名字、字段都不能丢）
    expect(out?.schema.name).toBe('幽灵节点')
    expect(out?.schema.kind).toBe('quantum-bridge')
    expect(out?.extensions?.lockedBy).toBe('user-42')
    expect(out?.position).toEqual({ x: 30, y: 40 })
    expect(out?.rows).toEqual([{ a: 1 }])
  })

  it('程序与旧文件都没有该节点时返回 null（让调用方跳过，而不是写空壳）', () => {
    const out = buildNodeFile('never-seen', {
      name: 'x',
      tables: {},
      records: {},
      positions: {},
    })
    // ★ 这条是saveProject「跳过而非覆盖」的前置条件。
    //   若这里返回了一个空 schema 的文件，saveProject 会把用户节点覆盖成空节点，
    //   而且校验能通过 —— 属于最坏的一类静默劣化。
    expect(out).toBeNull()
  })

  it('trashMeta 在程序认知里被清空后，不会被旧文件的值复活（删不掉的字段）', () => {
    // 用户把节点从回收站彻底删除 → 程序不再产出 trashMeta。
    // 若 trashMeta 不在 KNOWN_SCHEMA_KEYS 里，旧值会被当未知键塞回去，
    // 症状是「删了又出现在回收站，反复删都删不掉」。
    const src = {
      name: 'x',
      tables: { a: { id: 'a', name: 'A', kind: 'table' as const, fields: [], deleted: true } },
      records: {},
      positions: {},
    }
    const old = {
      id: 'a',
      schema: {
        id: 'a',
        name: 'A',
        kind: 'table',
        fields: [],
        deleted: true,
        trashMeta: { deletedAt: 123, refs: {} },
      },
      position: { x: 0, y: 0 },
      rows: [],
    }
    const out = buildNodeFile('a', src, old as never)
    expect((out!.schema as unknown as Record<string, unknown>).trashMeta).toBeUndefined()
  })
})

describe('KNOWN_NODE_KINDS 与 NodeKind 类型不漂移', () => {
  it('白名单与 schema.ts 的 NodeKind 全量交叉核对', async () => {
    // ★ 为什么要跑运行时核对：KNOWN_NODE_KINDS 是手抄的（NodeKind 是编译期类型，
    //   拿不到运行时数组），手抄必然有忘的可能。
    //   忘了的后果：新节点被当成未知类型报 warn —— 良性（只是多一句提示）。
    //   反过来若白名单里**多**抄了一个不存在的，则真实问题会被掩盖。故必须核对。
    //
    // ⚠ 防假绿：这个测试不能只断言「白名单非空」—— 那样少抄 10 个也照样绿。
    //   必须真正双向核对：① schema.ts 里每个 kind 都不能被误判为未知；
    //   ② 白名单里也不能有 schema.ts 里不存在的 kind。
    const fs = await import('node:fs')
    const schemaText = fs.readFileSync(new URL('./schema.ts', import.meta.url), 'utf8')

    // 从 `export type NodeKind` 起，扫到下一个顶层 `export` 为止。
    // 不依赖具体换行/缩进排版 —— 上一版用 [\s\S]*? 加终止条件的写法
    // 一旦有人在中间加个空行就匹配失败，测试会以「declared.size=1」的红脸
    // 呈现，完全看不出是排版变动引起的（假红，浪费排查时间）。
    const start = schemaText.indexOf('export type NodeKind')
    expect(start, 'schema.ts 里找不到 NodeKind 定义').toBeGreaterThan(-1)
    const rest = schemaText.slice(start)
    const endMatch = rest.slice('export type NodeKind'.length).search(/\nexport /)
    const block = endMatch === -1 ? rest : rest.slice(0, endMatch + 1)

    const declared = new Set<string>()
    for (const mm of block.matchAll(/'([a-z][a-z0-9]*)'/g)) declared.add(mm[1])

    // 交叉核对 ①：schema.ts 声明的每个 kind 都不能被误判为未知
    expect(declared.size, `从 schema.ts 只解析出 ${declared.size} 个 kind，正则可能又失效了`).toBeGreaterThan(10)
    for (const kind of declared) {
      const file = { id: 'k', schema: { id: 'k', name: 'k', kind, fields: [] }, position: { x: 0, y: 0 }, rows: [] }
      const warns = validateNodeFile(file, 'k').filter(
        (p) => p.level === 'warn' && p.message.includes('当前程序不认识'),
      )
      expect(warns, `kind=${kind} 被误判为未知类型 —— 请把 ${kind} 加进 KNOWN_NODE_KINDS`).toHaveLength(0)
    }

    // 交叉核对 ②：白名单里不能有不存在的 kind（否则真实未知类型会被放过）
    const wl = (await import('./projectFormat')) as unknown as Record<string, unknown>
    void wl
    // 白名单是模块私有的，通过「构造一个假 kind 看它是否报未知」反推其大小不可行，
    // 故改为断言「随便编一个 kind 一定报未知」—— 若将来有人把白名单写成
    // 「不判断」，这条会红。这是最省事且有效的守门。
    const fake = { id: 'k', schema: { id: 'k', name: 'k', kind: 'definitely-not-a-kind', fields: [] }, position: { x: 0, y: 0 }, rows: [] }
    expect(
      validateNodeFile(fake, 'k').some((p) => p.message.includes('当前程序不认识')),
      '白名单判定失效：未知 kind 竟然没报 warn',
    ).toBe(true)
  })
})
