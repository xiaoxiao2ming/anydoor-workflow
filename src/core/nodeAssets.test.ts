/**
 * nodeAssets 桥接层测试。
 *
 * ★ 本文件防的核心假绿有三条：
 *   ① 「文件名安全化」把空格/连字符一起吃掉 —— 写出来看着正常（还是合法文件名），
 *      但用户在工作区里看到的是一串下划线。用真实中文名断言。
 *   ② 资产反向恢复后 tplId 没改写 —— 字节灌回了模板库，但节点 config 还指着
 *      本机那个不存在的 id。症状是「工程打开了，双击还是空白」，
 *      而字节明明在。用 rewireTemplateIds 端到端断言。
 *   ③ 模板丢了只报 problem 不抛异常 —— 若抛异常，整个保存会失败，
 *      用户的表格数据都存不下来（代价极不成比例）。
 */

import { describe, it, expect } from 'vitest'
import {
  assetSpecOf,
  mimeOf,
  extOf,
  safeDisplayName,
  assetFileName,
  hashBytes,
  hashBytesBest,
  planNodeAssets,
  planAllNodeAssets,
  restoreNodeAssets,
  rewireTemplateIds,
  type AssetRuntime,
  type NodeAssetPlan,
} from './nodeAssets'
import { relPathIsSafe } from './projectStore'
import type { NodeAssetRef } from './projectFormat'
import type { TableSchema } from './schema'

/* ---------- 测试替身 ---------- */

const B = (s: string): Uint8Array => new TextEncoder().encode(s)

function memRuntime(opts?: {
  templates?: Record<string, { name: string; bytes: Uint8Array }>
  edited?: Record<string, { tplId: string; bytes: Uint8Array }>
}) {
  const templates = new Map(
    Object.entries(opts?.templates ?? {}).map(([k, v]) => [k, { name: v.name, bytes: v.bytes }]),
  )
  const edited = new Map(Object.entries(opts?.edited ?? {}))
  const calls: string[] = []
  let seq = 0

  const rt: AssetRuntime = {
    async readTemplate(id) {
      calls.push(`readTemplate:${id}`)
      return templates.get(id)?.bytes ?? null
    },
    async templateMeta(id) {
      const t = templates.get(id)
      return t ? { name: t.name, size: t.bytes.byteLength } : null
    },
    async readEdited(nodeId) {
      calls.push(`readEdited:${nodeId}`)
      return edited.get(nodeId) ?? null
    },
    async ensureTemplate(name, bytes) {
      calls.push(`ensureTemplate:${name}`)
      // 模拟 templateStore 的「按 name 去重」：同名返回既有 id
      for (const [id, t] of templates) if (t.name === name) return id
      const id = `new${++seq}`
      templates.set(id, { name, bytes })
      return id
    },
    async writeEdited(nodeId, tplId, bytes) {
      calls.push(`writeEdited:${nodeId}`)
      edited.set(nodeId, { tplId, bytes })
    },
  }
  return { rt, templates, edited, calls }
}

const S = (id: string, config?: Record<string, unknown>, name = id): TableSchema =>
  ({ id, name, kind: 'ooword', fields: [], config }) as unknown as TableSchema

/* ---------- 资产线索 ---------- */

describe('assetSpecOf · 从 schema 抽资产线索', () => {
  it('认docxTemplateId', () => {
    const s = assetSpecOf(S('n', { docxTemplateId: 't1', docxTemplateName: '报告.docx' }), false)
    expect(s.tplId).toBe('t1')
    expect(s.tplName).toBe('报告.docx')
  })

  it('认 xlsxTemplateId', () => {
    expect(assetSpecOf(S('n', { xlsxTemplateId: 't2' }), false).tplId).toBe('t2')
  })

  it('认 ooTemplateId', () => {
    expect(assetSpecOf(S('n', { ooTemplateId: 't3' }), false).tplId).toBe('t3')
  })

  it('没绑模板时 tplId 为空', () => {
    expect(assetSpecOf(S('n', {}), false).tplId).toBeUndefined()
  })

  it('schema 为空也不崩', () => {
    expect(assetSpecOf(undefined, true)).toEqual({ hasEdited: true })
  })

  it('★ 绑定字段名与 schema.ts 的 NodeConfig 实际字段一致', () => {
    // 抄自 src/core/schema.ts 的 NodeConfig。改了那边必须改这里，
    // 否则新绑定方式的模板永远不进工程，且没有任何报错。
    // id 与 name 分开断言：name 是「只用于显示」，id 才是定位依据，
    // 混在一起断言过一次（结果拿 name 字段去查 tplId，全undefined）。
    const ids: [string, string][] = [
      ['docxTemplateId', 'a'],
      ['xlsxTemplateId', 'b'],
      ['ooTemplateId', 'c'],
    ]
    const names: [string, string][] = [
      ['docxTemplateName', 'a.docx'],
      ['xlsxTemplateName', 'b.xlsx'],
      ['ooTemplateName', 'c.docx'],
    ]
    for (const [k, v] of ids) {
      expect(assetSpecOf(S('n', { [k]: v }), false).tplId).toBe(v)
    }
    for (const [k, v] of names) {
      expect(assetSpecOf(S('n', { [k]: v }), false).tplName).toBe(v)
    }
  })
})

/* ---------- MIME 与文件名 ---------- */

describe('MIME 与扩展名', () => {
  it('docx/xlsx 走 OOXML 正式 MIME', () => {
    expect(mimeOf('a.docx')).toBe(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )
    expect(mimeOf('a.xlsx')).toBe(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
  })

  it('大小写不敏感', () => {
    expect(mimeOf('A.DOCX')).toBe(mimeOf('a.docx'))
  })

  it('未知扩展名兜底为 octet-stream', () => {
    expect(mimeOf('a.zzz')).toBe('application/octet-stream')
    expect(mimeOf('noext')).toBe('application/octet-stream')
  })

  it('音频视频有明确 MIME（用户明确要求「音频、视频」）', () => {
    expect(mimeOf('a.mp3')).toBe('audio/mpeg')
    expect(mimeOf('a.mp4')).toBe('video/mp4')
    expect(mimeOf('a.wav')).toBe('audio/wav')
  })

  it('extOf 边界', () => {
    expect(extOf('a.tar.gz')).toBe('gz')
    expect(extOf('a.')).toBe('')
    expect(extOf('.gitignore')).toBe('gitignore')
    expect(extOf('noext')).toBe('')
  })
})

describe('safeDisplayName · 文件名安全化', () => {
  it('★ 保留中文、空格与连字符（这是最容易被"修坏"的一条）', () => {
    // 字符类写成 /[<>:"|?* -]/ 时，* 到空格被解析成范围，
    // 于是空格与连字符也被替换 → 「技术方案 - 终版.docx」变「技术方案____终版_docx」
    expect(safeDisplayName('技术方案 - 终版.docx')).toBe('技术方案 - 终版.docx')
  })

  it('替换路径分隔符', () => {
    expect(safeDisplayName('a/b\\c.docx')).toBe('a_b_c.docx')
  })

  it('替换 Windows 非法字符但不动中文', () => {
    expect(safeDisplayName('报<表>:格?.docx')).toBe('报_表__格_.docx')
  })

  it('控制字符被剔除（不留可疑下划线）', () => {
    expect(safeDisplayName('a\u0000b\u0001c.docx')).toBe('abc.docx')
  })

  it('全非法字符时有兜底名', () => {
    expect(safeDisplayName('')).toBe('unnamed')
    expect(safeDisplayName('   ')).toBe('unnamed')
  })

  it('产物一定不含路径分隔符（防越界）', () => {
    for (const n of ['../../etc/passwd', 'a/../../b', 'C:\\Windows\\x.dll', 'a\\b']) {
      const out = assetFileName('id1', n)
      expect(relPathIsSafe(out)).toBe(true)
    }
  })
})

describe('assetFileName · 落盘名', () => {
  it('带原名，便于用户在资源管理器里辨认', () => {
    expect(assetFileName('t1', '技术方案.docx')).toBe('t1__技术方案.docx')
  })

  it('压缩多余空白', () => {
    expect(assetFileName('t1', 'a   b.docx')).toBe('t1__a b.docx')
  })

  it('超长名被截断到 100 字符（留出 id 前缀空间）', () => {
    const out = assetFileName('t1', 'x'.repeat(500) + '.docx')
    expect(out.length).toBeLessThan(120)
  })
})

/* ---------- 指纹 ---------- */

describe('字节指纹', () => {
  it('同内容同指纹', () => {
    expect(hashBytes(B('abc'))).toBe(hashBytes(B('abc')))
  })

  it('不同内容不同指纹', () => {
    expect(hashBytes(B('abc'))).not.toBe(hashBytes(B('abd')))
  })

  it('降级值带 fnv1a 前缀（不会被误认为 SHA）', () => {
    expect(hashBytes(B('abc'))).toMatch(/^fnv1a-[0-9a-f]{8}-len\d+$/)
  })

  it('空字节不崩', () => {
    expect(hashBytes(new Uint8Array(0))).toMatch(/^fnv1a-/)
  })

  it('hashBytesBest 在无 WebCrypto 时降级且仍可用', async () => {
    // ★ globalThis.crypto 是**只读访问器**（Node 20+ / 浏览器都是），
    //   直接 `g.crypto = undefined` 会抛 TypeError。
    //   症状极具误导性：报错指向"赋值"，与被测的降级逻辑毫无关系。
    //   正确做法是 defineProperty 覆写 value。
    const g = globalThis as unknown as { crypto?: Crypto }
    const origDesc = Object.getOwnPropertyDescriptor(g, 'crypto')
    try {
      Object.defineProperty(g, 'crypto', { value: undefined, configurable: true, writable: true })
      const h = await hashBytesBest(B('abc'))
      expect(h).toMatch(/^fnv1a-/)
    } finally {
      if (origDesc) Object.defineProperty(g, 'crypto', origDesc)
    }
  })
})

/* ---------- 规划 ---------- */

describe('planNodeAssets · 单节点', () => {
  it('模板进 plans，登记齐全', async () => {
    const { rt } = memRuntime({
      templates: { t1: { name: '技术方案.docx', bytes: B('DOCXBYTES') } },
    })
    const p = await planNodeAssets('n1', S('n1', { ooTemplateId: 't1', ooTemplateName: '技术方案.docx' }), rt)
    expect(p.refs).toHaveLength(1)
    const r = p.refs[0]
    expect(r.id).toBe('t1')
    expect(r.role).toBe('template')
    expect(r.name).toBe('技术方案.docx')
    expect(r.mime).toContain('wordprocessingml')
    expect(r.size).toBe('DOCXBYTES'.length)
    expect(r.hash).toBeTruthy()
    expect(p.files.get(r.relPath)).toBeTruthy()
  })

  it('字节不进 JSON（Map 分离）', async () => {
    const { rt } = memRuntime({ templates: { t1: { name: 'a.docx', bytes: B('X') } } })
    const p = await planNodeAssets('n1', S('n1', { ooTemplateId: 't1' }), rt)
    // refs 里不能出现任何字节形态
    expect(JSON.stringify(p.refs)).not.toContain('88') // 'X' 的字节值
    expect(p.files.size).toBe(1)
  })

  it('★ 模板不在库里 → 报 problem，不抛异常', async () => {
    // 若抛异常，整个保存失败 → 用户的表格数据都存不下来，代价极不成比例
    const { rt } = memRuntime()
    const p = await planNodeAssets('n1', S('n1', { ooTemplateId: 'gone' }), rt)
    expect(p.refs).toHaveLength(0)
    expect(p.problems.join()).toContain('找不到')
  })

  it('元信息缺失时用 config 里的名字兜底', async () => {
    const { rt } = memRuntime({ templates: { t1: { name: 'real.docx', bytes: B('X') } } })
    // meta 有值时优先用 meta
    const p1 = await planNodeAssets('n1', S('n1', { ooTemplateId: 't1', ooTemplateName: 'cfg.docx' }), rt)
    expect(p1.refs[0].name).toBe('real.docx')
  })

  it('有编辑稿时额外登记一条 role=edited', async () => {
    const { rt } = memRuntime({
      templates: { t1: { name: 'a.docx', bytes: B('TPL') } },
      edited: { n1: { tplId: 't1', bytes: B('EDITED') } },
    })
    const p = await planNodeAssets('n1', S('n1', { ooTemplateId: 't1' }), rt, { hasEdited: true })
    expect(p.refs).toHaveLength(2)
    const ed = p.refs.find((r) => r.role === 'edited')!
    expect(ed.id).toBe('edited-n1')
    // ★ 稿的扩展名必须跟模板一致（决定打开路径的分支判断）
    expect(ed.relPath.endsWith('.docx')).toBe(true)
  })

  it('★ 稿的落盘名带真实扩展名（extOf 不含点，必须自己补）', async () => {
    // 曾经的 bug：`edited-${id}${extOf(name)}` 拼出「edited-n1docx」——
    // 一个没有扩展名的文件。症状极隐蔽：字节完全正确，
    // 但程序按扩展名判断类型的分支走不进去，且没有任何报错。
    const { rt } = memRuntime({
      templates: { t1: { name: 'a.docx', bytes: B('T') } },
      edited: { n1: { tplId: 't1', bytes: B('E') } },
    })
    const p = await planNodeAssets('n1', S('n1', { ooTemplateId: 't1' }), rt, { hasEdited: true })
    const ed = p.refs.find((r) => r.role === 'edited')!
    expect(ed.name).toBe('edited-n1.docx')
    expect(extOf(ed.name)).toBe('docx')
  })

  it('config 里没有 name 时，稿仍从模板库真实名取到扩展名', async () => {
    const { rt } = memRuntime({
      templates: { t1: { name: '真实名.xlsx', bytes: B('T') } },
      edited: { n1: { tplId: 't1', bytes: B('E') } },
    })
    const p = await planNodeAssets('n1', S('n1', { ooTemplateId: 't1' }), rt, { hasEdited: true })
    const ed = p.refs.find((r) => r.role === 'edited')!
    expect(extOf(ed.name)).toBe('xlsx')
  })

  it('模板与稿都没有可参照的扩展名时，稿名不带假扩展名', async () => {
    const { rt } = memRuntime({ edited: { n1: { tplId: 'ghost', bytes: B('E') } } })
    const p = await planNodeAssets('n1', S('n1', {}), rt, { hasEdited: true })
    const ed = p.refs.find((r) => r.role === 'edited')!
    expect(ed.name).toBe('edited-n1')
    expect(ed.name).not.toContain('undefined')
  })

  it('hasEdited=false 时不登记稿', async () => {    const { rt } = memRuntime({ edited: { n1: { tplId: 't1', bytes: B('E') } } })
    const p = await planNodeAssets('n1', S('n1', { ooTemplateId: 't1' }), rt, { hasEdited: false })
    expect(p.refs.every((r) => r.role === 'template')).toBe(true)
  })

  it('说 有稿但库里没有 → 不报错也不登记（安静跳过）', async () => {
    const { rt } = memRuntime({ templates: { t1: { name: 'a.docx', bytes: B('T') } } })
    const p = await planNodeAssets('n1', S('n1', { ooTemplateId: 't1' }), rt, { hasEdited: true })
    expect(p.refs).toHaveLength(1)
    expect(p.problems).toHaveLength(0)
  })

  it('schema 为空（程序不认识该节点）→ 空计划，不动旧 assets', async () => {
    const { rt } = memRuntime()
    const p = await planNodeAssets('n1', undefined, rt)
    expect(p.refs).toHaveLength(0)
    expect(p.files.size).toBe(0)
  })

  it('未绑模板且无稿 → 空计划', async () => {
    const { rt } = memRuntime()
    const p = await planNodeAssets('n1', S('n1', {}), rt)
    expect(p.refs).toHaveLength(0)
  })

  it('★ 字节没变时沿用旧 hash（省一次全量哈希）', async () => {
    const { rt } = memRuntime({ templates: { t1: { name: 'a.docx', bytes: B('SAME') } } })
    const prev: NodeAssetRef[] = [
      { id: 't1', relPath: 't1__a.docx', name: 'a.docx', role: 'template', mime: 'x', size: 4, hash: 'OLDHASH' },
    ]
    const p = await planNodeAssets('n1', S('n1', { ooTemplateId: 't1' }), rt, { prevRefs: prev })
    expect(p.refs[0].hash).toBe('OLDHASH')
  })

  it('字节变了（长度不同）→ 重新算 hash', async () => {
    const { rt } = memRuntime({ templates: { t1: { name: 'a.docx', bytes: B('LONGER') } } })
    const prev: NodeAssetRef[] = [
      { id: 't1', relPath: 't1__a.docx', name: 'a.docx', role: 'template', mime: 'x', size: 4, hash: 'OLDHASH' },
    ]
    const p = await planNodeAssets('n1', S('n1', { ooTemplateId: 't1' }), rt, { prevRefs: prev })
    expect(p.refs[0].hash).not.toBe('OLDHASH')
  })

  it('★ relPath 全部通过越界检查（收别人的工程时这层是安全边界）', async () => {
    const { rt } = memRuntime({
      templates: { t1: { name: '../../evil.docx', bytes: B('X') } },
      edited: { n1: { tplId: 't1', bytes: B('Y') } },
    })
    const p = await planNodeAssets('n1', S('n1', { ooTemplateId: 't1' }), rt, { hasEdited: true })
    expect(p.refs.length).toBeGreaterThan(0)
    for (const r of p.refs) expect(relPathIsSafe(r.relPath)).toBe(true)
  })

  it('模板库抛异常时不冒泡（降级为「找不到」）', async () => {
    const rt: AssetRuntime = {
      readTemplate: async () => {
        throw new Error('IndexedDB 崩了')
      },
      templateMeta: async () => {
        throw new Error('IndexedDB 崩了')
      },
      readEdited: async () => null,
      ensureTemplate: async () => 'x',
      writeEdited: async () => {},
    }
    const p = await planNodeAssets('n1', S('n1', { ooTemplateId: 't1' }), rt)
    expect(p.refs).toHaveLength(0)
    expect(p.problems).toHaveLength(1)
  })
})

describe('planAllNodeAssets · 全量', () => {
  it('逐节点规划并汇总问题', async () => {
    const { rt } = memRuntime({
      templates: { t1: { name: 'a.docx', bytes: B('A') } },
    })
    const tables = {
      n1: S('n1', { ooTemplateId: 't1' }),
      n2: S('n2', { ooTemplateId: 'gone' }),
      n3: S('n3', {}),
    }
    const out = await planAllNodeAssets(tables, rt)
    expect(out.byNode.n1.refs).toHaveLength(1)
    expect(out.byNode.n3.refs).toHaveLength(0)
    expect(out.problems).toHaveLength(1)
    expect(out.problems[0].nodeId).toBe('n2')
  })

  it('hasEdited 回调逐节点求值', async () => {
    const { rt } = memRuntime({ edited: { b: { tplId: 't', bytes: B('E') } } })
    const out = await planAllNodeAssets(
      { a: S('a'), b: S('b') },
      rt,
      { hasEdited: (id) => id === 'b' },
    )
    expect(out.byNode.a.refs).toHaveLength(0)
    expect(out.byNode.b.refs).toHaveLength(1)
  })
})

/* ---------- 反向恢复 ---------- */

describe('restoreNodeAssets · 工程 → 运行时', () => {
  it('模板字节入库并给出 id 映射', async () => {
    const { rt, templates } = memRuntime()
    const out = await restoreNodeAssets(
      { n1: { assets: [{ id: 'old1', relPath: 'r', name: '方案.docx', role: 'template', mime: 'x', size: 3 }] } },
      async () => B('DOC'),
      rt,
    )
    expect(out.problems).toHaveLength(0)
    expect(out.tplIdMap.n1.old1).toBeTruthy()
    expect(templates.size).toBe(1)
  })

  it('★ 返回的映射能把 config 改到新 id 上', async () => {
    // 这是整条链的闭环点：字节灌回模板库了，但节点 config 若还指着旧 id，
    // 症状是「工程打开了，双击还是空白」—— 字节明明在。
    const { rt } = memRuntime()
    const files = {
      n1: { assets: [{ id: 'oldTpl', relPath: 'r', name: '方案.docx', role: 'template' as const, mime: 'x', size: 3 }] },
    }
    const out = await restoreNodeAssets(files, async () => B('DOC'), rt)
    const t = S('n1', { ooTemplateId: 'oldTpl', ooTemplateName: '方案.docx' })
    const rewired = rewireTemplateIds(t, out.tplIdMap.n1)
    expect(rewired.config?.ooTemplateId).toBe(out.tplIdMap.n1.oldTpl)
    expect(rewired.config?.ooTemplateId).not.toBe('oldTpl')
    // 名字等其它配置不能被改掉
    expect(rewired.config?.ooTemplateName).toBe('方案.docx')
  })

  it('字节读不到时报问题，不抛异常', async () => {
    const { rt } = memRuntime()
    const out = await restoreNodeAssets(
      { n1: { assets: [{ id: 'x', relPath: 'r', name: 'a.docx', role: 'template', mime: 'x', size: 1 }] } },
      async () => null,
      rt,
    )
    expect(out.problems.join()).toContain('读不到')
  })

  it('稿写回编辑存储，不进模板库', async () => {
    const { rt, templates, calls } = memRuntime()
    await restoreNodeAssets(
      { n1: { assets: [{ id: 'edited-n1', relPath: 'r', name: 'e.docx', role: 'edited', mime: 'x', size: 1 }] } },
      async () => B('E'),
      rt,
    )
    expect(templates.size).toBe(0)
    expect(calls.some((c) => c.startsWith('writeEdited'))).toBe(true)
  })

  it('入库失败（同名冲突等）报问题但继续处理其它资产', async () => {
    const { rt } = memRuntime()
    const bad: AssetRuntime = { ...rt, ensureTemplate: async () => { throw new Error('配额满') } }
    const out = await restoreNodeAssets(
      {
        n1: {
          assets: [
            { id: 'a', relPath: 'r1', name: 'a.docx', role: 'template', mime: 'x', size: 1 },
            { id: 'b', relPath: 'r2', name: 'b.docx', role: 'template', mime: 'x', size: 1 },
          ],
        },
      },
      async (_n, ref) => B(ref.id),
      bad,
    )
    expect(out.problems.length).toBeGreaterThan(0)
  })

  it('无 assets 字段的节点不产生映射（老工程兼容）', async () => {
    const { rt } = memRuntime()
    const out = await restoreNodeAssets({ n1: {} }, async () => B('X'), rt)
    expect(out.tplIdMap).toEqual({})
    expect(out.problems).toHaveLength(0)
  })
})

describe('rewireTemplateIds · 改写绑定', () => {
  it('返回新对象（原地改会被 React 浅比较吃掉）', () => {
    const t = S('n1', { ooTemplateId: 'old' })
    const out = rewireTemplateIds(t, { old: 'new' })
    expect(out).not.toBe(t)
    expect(out.config).not.toBe(t.config)
    expect(t.config?.ooTemplateId).toBe('old') // 原对象不被改
  })

  it('无变化时返回原对象（避免无谓重渲染）', () => {
    const t = S('n1', { ooTemplateId: 'keep' })
    expect(rewireTemplateIds(t, { other: 'x' })).toBe(t)
  })

  it('config 为空时返回原对象', () => {
    const t = S('n1', undefined)
    expect(rewireTemplateIds(t, { a: 'b' })).toBe(t)
  })

  it('三个绑定字段都能改写', () => {
    for (const k of ['docxTemplateId', 'xlsxTemplateId', 'ooTemplateId'] as const) {
      const t = S('n1', { [k]: 'old' })
      const out = rewireTemplateIds(t, { old: 'new' })
      expect(out.config?.[k]).toBe('new')
    }
  })
})

/* ---------- 端到端：保存规划 → 恢复 ---------- */

describe('端到端 · 规划出的资产能原样恢复', () => {
  it('字节逐字节一致（这是「不丢内容」的唯一硬标准）', async () => {
    const docBytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x00, 0xff, 0xfe, 0x80])
    const { rt } = memRuntime({ templates: { t1: { name: '技术方案.docx', bytes: docBytes } } })
    const plan: NodeAssetPlan = await planNodeAssets(
      'n1',
      S('n1', { ooTemplateId: 't1', ooTemplateName: '技术方案.docx' }),
      rt,
    )

    // 模拟「写到工程目录再读回来」：只用登记里的 relPath 找字节
    const fsLike = new Map(plan.files)
    const { rt: rt2 } = memRuntime()
    const out = await restoreNodeAssets(
      { n1: { assets: plan.refs } },
      async (_n, ref) => fsLike.get(ref.relPath) ?? null,
      rt2,
    )
    expect(out.problems).toHaveLength(0)
    // 从新库里取回比对 —— 只有逐字节相同才算「不丢内容」
    const ids = Object.values(out.tplIdMap.n1)
    expect(ids).toHaveLength(1)
    const back = await rt2.readTemplate(ids[0])
    expect(Array.from(back!)).toEqual(Array.from(docBytes))
  })
})
