/**
 * 工程 IO 编排层测试（saveProjectFull / openProjectFull）。
 *
 * ★ 这是整条资产链的**唯一真验收点**。前面的 nodeAssets.test.ts 验的是
 *   「规划得对不对」，本文件验的是「存进去、读回来、逐字节相同」。
 *   两者的失败症状完全不同：
 *     规划错 → 登记条目不对（可能还能用）
 *     IO 错   → 工程看着完好，双击打开是空白（数据静默丢失）
 *   所以本文件刻意不 mock 掉 ProjectFs，而是用一个真实落盘的内存 FS
 *   跑完整链路 —— mock 掉 IO 就等于把要验的东西也mock 掉了。
 */

import { describe, it, expect } from 'vitest'
import { saveProjectFull, openProjectFull, applyTplIdMap } from './projectIO'
import { PROJECT_DIRS, PROJECT_FILES, type ProjectFs } from './projectStore'
import { packProject, unpackProject } from './projectPack'
import type { AssetRuntime } from './nodeAssets'
import type { ViewStateKV } from './viewState'
import type { TableSchema } from './schema'

/* ---------- 真实落盘的内存 ProjectFs（不 mock IO） ---------- */

function memFs() {
  const files = new Map<string, string>()
  const bins = new Map<string, Uint8Array>()
  const dirs = new Set<string>()

  const ensureDir = (dir: string) => {
    // 记录所有祖先目录，模拟「建目录」语义
    const parts = dir.split('/')
    for (let i = 1; i <= parts.length; i++) dirs.add(parts.slice(0, i).join('/'))
  }

  const fs: ProjectFs = {
    async list(dir) {
      const prefix = dir.endsWith('/') ? dir : `${dir}/`
      const out = new Set<string>()
      for (const k of [...files.keys(), ...bins.keys()]) {
        if (!k.startsWith(prefix)) continue
        const rest = k.slice(prefix.length)
        const slash = rest.indexOf('/')
        out.add(slash === -1 ? rest : rest.slice(0, slash))
      }
      return [...out]
    },
    async readText(p) {
      return files.get(p) ?? null
    },
    async readBytes(p) {
      return bins.get(p) ?? null
    },
    async writeText(p, t) {
      ensureDir(p.slice(0, p.lastIndexOf('/')))
      files.set(p, t)
    },
    async writeBytes(p, b) {
      ensureDir(p.slice(0, p.lastIndexOf('/')))
      bins.set(p, b)
    },
    async mkdir(d) {
      ensureDir(d)
    },
    async exists(p) {
      return files.has(p) || bins.has(p)
    },
  }
  return { fs, files, bins, dirs }
}

/* ---------- 运行时资产存储（模拟 templateStore + editStore） ---------- */

function memRuntime() {
  const templates = new Map<string, { name: string; bytes: Uint8Array }>()
  const edited = new Map<string, { tplId: string; bytes: Uint8Array }>()
  let seq = 0
  const rt: AssetRuntime = {
    async readTemplate(id) {
      return templates.get(id)?.bytes ?? null
    },
    async templateMeta(id) {
      const t = templates.get(id)
      return t ? { name: t.name, size: t.bytes.byteLength } : null
    },
    async readEdited(id) {
      return edited.get(id) ?? null
    },
    async ensureTemplate(name, bytes) {
      for (const [id, t] of templates) if (t.name === name) return id
      const id = `tpl${++seq}`
      templates.set(id, { name, bytes })
      return id
    },
    async writeEdited(id, tplId, bytes) {
      edited.set(id, { tplId, bytes })
    },
  }
  return { rt, templates, edited }
}

function memKV(seed?: Record<string, string>) {
  const m = new Map(Object.entries(seed ?? {}))
  const kv: ViewStateKV = {
    getItem: (k) => (m.has(k) ? (m.get(k) as string) : null),
    setItem: (k, v) => {
      m.set(k, v)
    },
    removeItem: (k) => {
      m.delete(k)
    },
  }
  return { m, kv }
}

const B = (s: string) => new TextEncoder().encode(s)
const S = (id: string, name: string, config?: Record<string, unknown>): TableSchema =>
  ({ id, name, kind: 'ooword', fields: [], config }) as unknown as TableSchema

/** 有真实 OOXML 头部特征的字节（不是纯 ASCII，才测得出二进制通路） */
const DOCX = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00, 0xff, 0xfe, 0x80, 0x81])

/* ================= 保存 ================= */

describe('saveProjectFull', () => {
  it('★ 资产字节真的落进 assets/ 且逐字节相同', async () => {
    const m = memFs()
    const { rt, templates } = memRuntime()
    templates.set('t1', { name: '技术方案.docx', bytes: DOCX })

    const out = await saveProjectFull(
      m.fs,
      { rootDir: 'p' },
      {
        name: '工程',
        tables: { n1: S('n1', '报告', { ooTemplateId: 't1', ooTemplateName: '技术方案.docx' }) },
        records: { n1: [] },
        positions: { n1: { x: 1, y: 2 } },
      },
      { runtime: rt, kv: memKV().kv },
    )

    expect(out.assetProblems).toEqual([])
    expect(out.assetCount).toBe(1)
    const ref = JSON.parse(m.files.get(`p/${PROJECT_DIRS.nodes}/n1.json`) as string).assets[0]
    /* ★ 2026-10-24 统一资产仓：字节落 assets/content/{指纹}.{ext}（refs.relPath 指向它） */
    const path = `p/${PROJECT_DIRS.assets}/${ref.relPath}`
    expect(path).toContain('/content/')
    expect(m.bins.has(path)).toBe(true)
    expect(Array.from(m.bins.get(path)!)).toEqual(Array.from(DOCX))
  })

  it('控件状态从 localStorage 进了节点文件', async () => {
    const m = memFs()
    const { rt } = memRuntime()
    const { kv } = memKV({ 'ne-tab:n1': 'book', 'refpanel:n1': '0', 'wb-sheet:n1': 'S1' })

    await saveProjectFull(
      m.fs,
      { rootDir: 'p' },
      {
        name: '工程',
        tables: { n1: S('n1', 'a') },
        records: { n1: [] },
        positions: { n1: { x: 0, y: 0 } },
      },
      { runtime: rt, kv },
    )

    const f = JSON.parse(m.files.get(`p/${PROJECT_DIRS.nodes}/n1.json`) as string)
    expect(f.viewState).toEqual({ tab: 'book', refPanelOpen: false, activeSheet: 'S1' })
  })

  it('★ 没有 writeBytes 的适配器 → 明确抛错，不静默丢资产', async () => {
    // 静默降级是这类软件最坏的失败方式：用户看到「保存成功」，
    // 而工程里只有节点没有文档。换机器打开就是一堆空节点。
    const m = memFs()
    const { rt, templates } = memRuntime()
    templates.set('t1', { name: 'a.docx', bytes: DOCX })
    const noBin: ProjectFs = { ...m.fs, writeBytes: undefined }

    await expect(
      saveProjectFull(
        noBin,
        { rootDir: 'p' },
        {
          name: '工程',
          tables: { n1: S('n1', 'a', { ooTemplateId: 't1' }) },
          records: { n1: [] },
          positions: { n1: { x: 0, y: 0 } },
        },
        { runtime: rt, kv: memKV().kv },
      ),
    ).rejects.toThrow(/二进制/)
  })

  it('模板缺失时保存仍成功，只报问题', async () => {
    const m = memFs()
    const { rt } = memRuntime()
    const out = await saveProjectFull(
      m.fs,
      { rootDir: 'p' },
      {
        name: '工程',
        tables: { n1: S('n1', 'a', { ooTemplateId: 'gone' }) },
        records: { n1: [] },
        positions: { n1: { x: 0, y: 0 } },
      },
      { runtime: rt, kv: memKV().kv },
    )
    // 工程文件必须还在（用户的数据比文档重要）
    expect(m.files.has(`p/${PROJECT_FILES.manifest}`)).toBe(true)
    expect(out.assetProblems).toHaveLength(1)
  })

  it('有编辑稿时两个资产都落盘', async () => {
    const m = memFs()
    const { rt, templates, edited } = memRuntime()
    templates.set('t1', { name: 'a.docx', bytes: DOCX })
    edited.set('n1', { tplId: 't1', bytes: B('EDITED-BYTES') })

    const out = await saveProjectFull(
      m.fs,
      { rootDir: 'p' },
      {
        name: '工程',
        tables: { n1: S('n1', 'a', { ooTemplateId: 't1' }) },
        records: { n1: [] },
        positions: { n1: { x: 0, y: 0 } },
      },
      { runtime: rt, kv: memKV().kv, hasEdited: (id) => id === 'n1' },
    )
    expect(out.assetCount).toBe(2)
  })
})

/* ================= 打开 ================= */

describe('openProjectFull', () => {
  it('★ 完整闭环：存 → 读 → 资产字节逐字节一致', async () => {
    // 这是整个需求的核心断言。前面所有测试都是为这一条服务的。
    const m = memFs()
    const { rt, templates } = memRuntime()
    templates.set('t1', { name: '技术方案.docx', bytes: DOCX })
    const kvSave = memKV({ 'ne-tab:n1': 'book', 'wb-sheet:n1': '材料表' })

    await saveProjectFull(
      m.fs,
      { rootDir: 'p' },
      {
        name: '工程',
        tables: { n1: S('n1', '报告', { ooTemplateId: 't1', ooTemplateName: '技术方案.docx' }) },
        records: { n1: [{ f1: '值1' }] },
        positions: { n1: { x: 30, y: 40 } },
      },
      { runtime: rt, kv: kvSave.kv },
    )

    // 换一套全新的运行时与 localStorage（模拟「同事的机器」）
    const rt2 = memRuntime().rt
    const kv2 = memKV()
    const opened = await openProjectFull(m.fs, { rootDir: 'p' }, { runtime: rt2, kv: kv2.kv })

    // ① 结构
    expect(opened.tables.n1.name).toBe('报告')
    expect(opened.records.n1).toEqual([{ f1: '值1' }])
    expect(opened.positions.n1).toEqual({ x: 30, y: 40 })
    // ② 资产字节
    const newId = opened.tplIdMap.n1.t1
    expect(newId).toBeTruthy()
    const got = await rt2.readTemplate(newId)
    expect(Array.from(got!)).toEqual(Array.from(DOCX))
    // ③ 控件状态
    expect(kv2.m.get('ne-tab:n1')).toBe('book')
    expect(kv2.m.get('wb-sheet:n1')).toBe('材料表')
  })

  it('★ tplIdMap 能把 config 改到本机真实 id 上', async () => {
    const m = memFs()
    const { rt, templates } = memRuntime()
    templates.set('t1', { name: 'a.docx', bytes: DOCX })
    await saveProjectFull(
      m.fs,
      { rootDir: 'p' },
      {
        name: '工程',
        tables: { n1: S('n1', 'a', { ooTemplateId: 't1', ooTemplateName: 'a.docx' }) },
        records: { n1: [] },
        positions: { n1: { x: 0, y: 0 } },
      },
      { runtime: rt, kv: memKV().kv },
    )

    const rt2 = memRuntime().rt
    const opened = await openProjectFull(m.fs, { rootDir: 'p' }, { runtime: rt2, kv: memKV().kv })
    const fixed = applyTplIdMap(opened.tables, opened.tplIdMap)
    expect(fixed.n1.config?.ooTemplateId).toBe(opened.tplIdMap.n1.t1)
    // 原表不被改（原地改会被 React 浅比较吃掉）
    expect(opened.tables.n1.config?.ooTemplateId).toBe('t1')
  })

  it('applyTplIdMap 无映射时返回原对象（不触发无谓重渲染）', () => {
    const t = { n1: S('n1', 'a') }
    expect(applyTplIdMap(t, {})).toBe(t)
  })

  it('★ 打开老工程（无 viewState）不清空本地控件状态', async () => {
    // 最关键的一条：老工程没有 viewState，
    // 若实现是「以工程为准清空」，用户存了一年的面板布局会被
    // 「打开一个老工程」这个无害动作抹掉。
    const m = memFs()
    // 手工造一个 v1 工程（无 viewState / 无 assets）
    await m.fs.writeText(`p/${PROJECT_FILES.manifest}`, JSON.stringify({
      formatVersion: 1,
      formatKind: 'cbproj-dir',
      name: '老工程',
      createdAt: '2026-01-01T00:00:00.000Z',
      updatedAt: '2026-01-01T00:00:00.000Z',
      nodeIds: ['n1'],
      refs: {},
      dataSources: [],
    }))
    await m.fs.writeText(`p/${PROJECT_DIRS.nodes}/n1.json`, JSON.stringify({
      id: 'n1',
      schema: { id: 'n1', name: '老节点', kind: 'table', fields: [] },
      position: { x: 1, y: 2 },
      rows: [],
    }))

    const kv = memKV({ 'ne-tab:n1': 'book', 'refpanel:n1': '0' })
    const opened = await openProjectFull(m.fs, { rootDir: 'p' }, {
      runtime: memRuntime().rt,
      kv: kv.kv,
    })
    expect(opened.problems.filter((p) => p.level === 'error')).toEqual([])
    expect(kv.m.get('ne-tab:n1')).toBe('book')
    expect(kv.m.get('refpanel:n1')).toBe('0')
  })

  it('资产字节缺失时报 warn 而非 error（工程仍能打开）', async () => {
    const m = memFs()
    const { rt, templates } = memRuntime()
    templates.set('t1', { name: 'a.docx', bytes: DOCX })
    await saveProjectFull(
      m.fs,
      { rootDir: 'p' },
      {
        name: '工程',
        tables: { n1: S('n1', 'a', { ooTemplateId: 't1' }) },
        records: { n1: [] },
        positions: { n1: { x: 0, y: 0 } },
      },
      { runtime: rt, kv: memKV().kv },
    )
    // 模拟用户手工删了 assets 下的字节
    for (const k of [...m.bins.keys()]) m.bins.delete(k)

    const opened = await openProjectFull(m.fs, { rootDir: 'p' }, {
      runtime: memRuntime().rt,
      kv: memKV().kv,
    })
    expect(opened.problems.some((p) => p.message.includes('资产字节缺失'))).toBe(true)
    expect(opened.problems.filter((p) => p.level === 'error')).toEqual([])
  })

  it('未知节点类型的工程仍能打开（降级显示，不整体拒绝）', async () => {
    const m = memFs()
    await m.fs.writeText(`p/${PROJECT_FILES.manifest}`, JSON.stringify({
      formatVersion: 2,
      formatKind: 'cbproj-dir',
      name: '未来工程',
      createdAt: '2026-10-01T00:00:00.000Z',
      updatedAt: '2026-10-01T00:00:00.000Z',
      nodeIds: ['n1', 'n2'],
      refs: {},
      dataSources: [],
    }))
    await m.fs.writeText(`p/${PROJECT_DIRS.nodes}/n1.json`, JSON.stringify({
      id: 'n1',
      schema: { id: 'n1', name: '认识', kind: 'table', fields: [] },
      position: { x: 0, y: 0 },
      rows: [],
    }))
    await m.fs.writeText(`p/${PROJECT_DIRS.nodes}/n2.json`, JSON.stringify({
      id: 'n2',
      schema: { id: 'n2', name: '未来节点', kind: 'quantumlink', quantumMeta: { x: 1 } },
      position: { x: 9, y: 9 },
      rows: [],
    }))

    const opened = await openProjectFull(m.fs, { rootDir: 'p' }, {
      runtime: memRuntime().rt,
      kv: memKV().kv,
    })
    expect(opened.problems.filter((p) => p.level === 'error')).toEqual([])
    expect(opened.tables.n2.name).toBe('未来节点')
    // 未知字段被原样带出，重新保存时才能继承
    expect((opened.tables.n2 as unknown as Record<string, unknown>).quantumMeta).toEqual({ x: 1 })
  })

  it('viewConflicts 报出本地与工程的分歧', async () => {
    const m = memFs()
    const { rt } = memRuntime()
    await saveProjectFull(
      m.fs,
      { rootDir: 'p' },
      { name: '工程', tables: { n1: S('n1', 'a') }, records: { n1: [] }, positions: { n1: { x: 0, y: 0 } } },
      { runtime: rt, kv: memKV({ 'ne-tab:n1': 'book' }).kv },
    )
    // 本地已改成 data
    const kv = memKV({ 'ne-tab:n1': 'data' })
    const opened = await openProjectFull(m.fs, { rootDir: 'p' }, { runtime: memRuntime().rt, kv: kv.kv })
    // 恢复后本地应被工程覆盖为 book，故最终无分歧
    expect(kv.m.get('ne-tab:n1')).toBe('book')
    expect(opened.viewConflicts).toEqual([])
  })
})

/* ================= 跨进程：保存后再开再存 ================= */

describe('跨进程往返', () => {
  it('★ 连续两次「打开 → 另存」后资产与控件状态不丢', async () => {
    // 最容易被漏掉的一环：单进程内连存两次看不出问题，
    // 关掉程序再打开另存才发现丢了。所以必须真的重建运行时存储。
    const m = memFs()
    const { rt, templates } = memRuntime()
    templates.set('t1', { name: 'a.docx', bytes: DOCX })
    const kv1 = memKV({ 'ne-tab:n1': 'book' })

    await saveProjectFull(
      m.fs,
      { rootDir: 'p' },
      { name: '工程', tables: { n1: S('n1', 'a', { ooTemplateId: 't1' }) }, records: { n1: [{ f: 1 }] }, positions: { n1: { x: 5, y: 6 } } },
      { runtime: rt, kv: kv1.kv },
    )

    // ── 第二次会话：全新运行时、全新 localStorage ──
    const rt2 = memRuntime().rt
    const kv2 = memKV()
    const opened = await openProjectFull(m.fs, { rootDir: 'p' }, { runtime: rt2, kv: kv2.kv })
    const tables2 = applyTplIdMap(opened.tables, opened.tplIdMap)
    const newTplId = tables2.n1.config?.ooTemplateId as string

    // 另存（用本机的新 tplId + 本机的控件状态）
    await saveProjectFull(
      m.fs,
      { rootDir: 'p2' },
      { name: '工程', tables: tables2, records: opened.records, positions: opened.positions },
      { runtime: rt2, kv: kv2.kv, prev: opened },
    )

    // ── 第三次会话验证 ──
    const rt3 = memRuntime().rt
    const opened3 = await openProjectFull(m.fs, { rootDir: 'p2' }, { runtime: rt3, kv: memKV().kv })
    expect(opened3.records.n1).toEqual([{ f: 1 }])
    expect(opened3.positions.n1).toEqual({ x: 5, y: 6 })
    const id3 = opened3.tplIdMap.n1[newTplId]
    expect(id3).toBeTruthy()
    const bytes3 = await rt3.readTemplate(id3)
    expect(Array.from(bytes3!)).toEqual(Array.from(DOCX))
  })
})

/* ================= 与 zip 打包串起来 ================= */

describe('端到端 · 工程 → zip → 新目录 → 打开', () => {
  it('★ 打包解包后资产字节与控件状态全在', async () => {
    const m = memFs()
    const { rt, templates } = memRuntime()
    templates.set('t1', { name: '技术方案.docx', bytes: DOCX })
    await saveProjectFull(
      m.fs,
      { rootDir: 'p' },
      {
        name: '工程',
        tables: { n1: S('n1', '报告', { ooTemplateId: 't1' }) },
        records: { n1: [{ f: 'v' }] },
        positions: { n1: { x: 7, y: 8 } },
      },
      { runtime: rt, kv: memKV({ 'ne-tab:n1': 'book' }).kv },
    )

    const zip = await packProject(m.fs, 'p')
    const dst = memFs()
    const unp = await unpackProject(dst.fs, 'out', zip)
    expect(unp.problems).toEqual([])

    const rt2 = memRuntime().rt
    const kv2 = memKV()
    const opened = await openProjectFull(dst.fs, { rootDir: 'out' }, { runtime: rt2, kv: kv2.kv })

    expect(opened.records.n1).toEqual([{ f: 'v' }])
    expect(opened.positions.n1).toEqual({ x: 7, y: 8 })
    const id = opened.tplIdMap.n1.t1
    expect(Array.from((await rt2.readTemplate(id))!)).toEqual(Array.from(DOCX))
    expect(kv2.m.get('ne-tab:n1')).toBe('book')
  })

  it('同一工程打包两次字节相同（进 git 才有意义）', async () => {
    const m = memFs()
    const { rt, templates } = memRuntime()
    templates.set('t1', { name: 'a.docx', bytes: DOCX })
    const input = {
      name: '工程',
      tables: { n1: S('n1', 'a', { ooTemplateId: 't1' }) },
      records: { n1: [] },
      positions: { n1: { x: 0, y: 0 } },
    }
    await saveProjectFull(m.fs, { rootDir: 'p' }, input, { runtime: rt, kv: memKV().kv })
    const z1 = await packProject(m.fs, 'p')
    const z2 = await packProject(m.fs, 'p')
    expect(z1.length).toBe(z2.length)
    expect(z1.every((b, i) => b === z2[i])).toBe(true)
  })
})
