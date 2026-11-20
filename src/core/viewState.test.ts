/**
 * viewState 桥接层测试。
 *
 * ★ 本文件重点防两类假绿：
 *   ① 键位映射漂移 —— 采集与恢复各写一份 switch 时必然发生，
 *      症状是「同字段多前缀」在采集和恢复下取到不同键。用 roundtrip 测试锁死。
 *   ② 「以工程为准」的越界实现 —— 用老工程（无 viewState）覆盖本地状态，
 *      症状是用户存了一年的面板状态被一次无害的打开动作抹掉。用「只写不删」测试锁死。
 */

import { describe, it, expect } from 'vitest'
import {
  VIEW_KEYMAP,
  COLLAPSE_PREFIX,
  browserKV,
  collectNodeViewState,
  collectAllViewState,
  collectViewStateFromTables,
  applyNodeViewState,
  applyAllViewState,
  auditViewStateConflicts,
  viewStateCoverage,
  type ViewStateKV,
} from './viewState'
import type { TableSchema } from './schema'

/** 内存 KV：给 Map 包一层，同时统计写了哪些键（用于「只写不删」断言） */
function memKV(seed?: Record<string, string>) {
  const m = new Map<string, string>(Object.entries(seed ?? {}))
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

const T = (id: string, kind = 'xlsx'): TableSchema =>
  ({ id, name: id, kind, fields: [], config: {} }) as unknown as TableSchema

describe('viewState · 采集', () => {
  it('采得到 tab / activeSheet / triggerRules', () => {
    const { kv } = memKV({
      'ne-tab:n1': 'book',
      'wb-sheet:n1': '管道清单',
      [`trigrules:n1`]: JSON.stringify([{ id: 'r1', on: 'change' }]),
    })
    const vs = collectNodeViewState('n1', kv)
    expect(vs?.tab).toBe('book')
    expect(vs?.activeSheet).toBe('管道清单')
    expect(vs?.triggerRules).toEqual([{ id: 'r1', on: 'change' }])
  })

  it('采得到折叠态', () => {
    const { kv } = memKV({ [`${COLLAPSE_PREFIX}:n1`]: '1' })
    expect(collectNodeViewState('n1', kv)?.collapsed).toBe(true)
    const { kv: kv2 } = memKV({ [`${COLLAPSE_PREFIX}:n1`]: '0' })
    expect(collectNodeViewState('n1', kv2)?.collapsed).toBe(false)
  })

  it('★ 一个键都没有时返回 undefined，而不是空对象', () => {
    // 为什么重要：buildNodeFile 里 undefined = 沿用旧值，空对象 = 覆盖成空。
    // 返回空对象会把老工程里已有的 viewState 抹成空 —— 控件状态全丢。
    const { kv } = memKV()
    expect(collectNodeViewState('n1', kv)).toBeUndefined()
  })

  it('★ 不把默认值固化进 viewState', () => {
    // 用户从没动过引用栏 → localStorage 无此键 → 不该写 refPanelOpen: true。
    // 若固化，将来改默认（改成一上来就关），老工程会一直卡在旧的 true 上。
    const { kv } = memKV()
    const vs = collectNodeViewState('n1', kv)
    expect(vs?.refPanelOpen).toBeUndefined()
    expect(vs?.rightPanelOpen).toBeUndefined()
  })

  it('坏值不采纳，退回「无此状态」而不是猜', () => {
    // 'true' 不是本程序约定的 bool 编码。猜成 true 会让面板开合与用户上次操作相反。
    const { kv } = memKV({ 'refpanel:n1': 'true' })
    const vs = collectNodeViewState('n1', kv)
    expect(vs).toBeUndefined()
  })

  it('trigrules 是坏 JSON 时只丢该字段，不影响同节点其它字段', () => {
    const { kv } = memKV({ 'trigrules:n1': '{oops', 'ne-tab:n1': 'data' })
    const vs = collectNodeViewState('n1', kv)
    expect(vs?.tab).toBe('data')
    expect(vs?.triggerRules).toBeUndefined()
  })

  it('空串不采纳', () => {
    const { kv } = memKV({ 'ne-tab:n1': '' })
    expect(collectNodeViewState('n1', kv)).toBeUndefined()
  })

  it('★ 同字段多前缀：按 VIEW_KEYMAP 顺序取第一个存在的', () => {
    // ooref 与 refpanel 都 → refPanelOpen。两者都存在时必须稳定取 ooref，
    // 否则同一份数据两次采集可能得到不同结果（若实现里用了 Object.keys 遍历）。
    const { kv } = memKV({ 'ooref:n1': '0', 'refpanel:n1': '1' })
    const vs = collectNodeViewState('n1', kv)
    expect(vs?.refPanelOpen).toBe(false)
  })

  it('前一个前缀不存在时，取后一个（不能因为前者缺失就整个字段为空）', () => {
    const { kv } = memKV({ 'refpanel:n1': '0' })
    expect(collectNodeViewState('n1', kv)?.refPanelOpen).toBe(false)
    const { kv: k2 } = memKV({ 'wrp-refpanel:n1': '1' })
    expect(collectNodeViewState('n1', k2)?.refPanelOpen).toBe(true)
  })

  it('KV 抛异常时整体退化为「无状态」，不冒泡阻断保存', () => {
    const boom: ViewStateKV = {
      getItem: () => {
        throw new Error('隐私模式')
      },
      setItem: () => {
        throw new Error('配额满')
      },
      removeItem: () => {},
    }
    expect(collectNodeViewState('n1', boom)).toBeUndefined()
    expect(applyNodeViewState('n1', { tab: 'book' }, boom)).toBe(0)
  })
})

describe('viewState · 采集（批量）', () => {
  it('collectAllViewState 跳过无状态的节点', () => {
    const { kv } = memKV({ 'ne-tab:a': 'book', 'wb-sheet:b': 'S2' })
    const all = collectAllViewState(['a', 'b', 'c'], kv)
    expect(Object.keys(all).sort()).toEqual(['a', 'b'])
    expect(all.b.activeSheet).toBe('S2')
  })

  it('collectViewStateFromTables 含回收站节点', () => {
    // 少存一个节点的状态 → 恢复节点后面板全展开 → 用户投诉且难自查
    const { kv } = memKV({ 'ne-tab:trashed': 'data' })
    const all = collectViewStateFromTables({ live: T('live'), trashed: T('trashed') }, kv)
    expect(all.trashed?.tab).toBe('data')
  })
})

describe('viewState · 恢复（回灌 localStorage）', () => {
  it('一个字段写回它对应的全部前缀', () => {
    const { m, kv } = memKV()
    applyNodeViewState('n1', { refPanelOpen: false }, kv)
    expect(m.get('ooref:n1')).toBe('0')
    expect(m.get('refpanel:n1')).toBe('0')
    expect(m.get('wrp-refpanel:n1')).toBe('0')
  })

  it('★ 严格「只写不删」：工程没给的字段不动本地', () => {
    // 这是最关键的一条。老工程没有 viewState，
    // 若实现是「以工程为准清空本地」，用户存了一年的面板状态会被
    // 「打开一个老工程」这个无害动作抹掉。
    const { m, kv } = memKV({ 'ne-tab:n1': 'book', 'refpanel:n1': '1' })
    applyNodeViewState('n1', { rightPanelOpen: true }, kv)
    // 本地原值必须还在
    expect(m.get('ne-tab:n1')).toBe('book')
    expect(m.get('refpanel:n1')).toBe('1')
    // 新值已写入
    expect(m.get('rightpanel:n1')).toBe('1')
    expect(m.get('wrp-rightpanel:n1')).toBe('1')
  })

  it('viewState 为 undefined 时一个键都不动', () => {
    const { m, kv } = memKV({ 'ne-tab:n1': 'book' })
    expect(applyNodeViewState('n1', undefined, kv)).toBe(0)
    expect(m.get('ne-tab:n1')).toBe('book')
  })

  it('applyAllViewState 返回总写入键数', () => {
    const { kv } = memKV()
    const n = applyAllViewState(
      { a: { tab: 'book' }, b: { collapsed: true }, c: undefined },
      kv,
    )
    // a: tab 1个；b: collapsed 1个；c: 0
    expect(n).toBe(2)
  })

  it('applyAllViewState(undefined) 不崩', () => {
    expect(applyAllViewState(undefined, memKV().kv)).toBe(0)
  })
})

describe('viewState · 采集 → 恢复 往返（锁死映射漂移）', () => {
  it('★ 全字段往返后逐字一致', () => {
    // 若采集与恢复各写一份 switch，这一条必然失败 ——
    // 两边的字段映射只要有一处不一致，往返就丢字段或多写键。
    const src = memKV({
      'ne-tab:n1': 'book',
      'ooref:n1': '0',
      'refpanel:n1': '1', // 与 ooref 冲突：采集取 ooref='0'
      'rightpanel:n1': '1',
      'wrp-rightpanel:n1': '0', // 与 rightpanel 冲突：采集取 rightpanel='1'
      'wb-sheet:n1': '材料表',
      'trigrules:n1': JSON.stringify([{ id: 'r1' }]),
      [`${COLLAPSE_PREFIX}:n1`]: '1',
    })
    const vs = collectNodeViewState('n1', src.kv)
    const dst = memKV()
    applyNodeViewState('n1', vs, dst.kv)
    const vs2 = collectNodeViewState('n1', dst.kv)

    expect(vs2).toEqual(vs)
    // 恢复后各前缀一致（同字段写到全部前缀）
    expect(dst.m.get('ooref:n1')).toBe('0')
    expect(dst.m.get('refpanel:n1')).toBe('0')
    expect(dst.m.get('rightpanel:n1')).toBe('1')
    expect(dst.m.get('wrp-rightpanel:n1')).toBe('1')
  })

  it('往返两次仍稳定（幂等）', () => {
    let kv = memKV({ 'ne-tab:n1': 'book', 'refpanel:n1': '0' }).kv
    for (let i = 0; i < 3; i++) {
      const vs = collectNodeViewState('n1', kv)
      const dst = memKV()
      applyNodeViewState('n1', vs, dst.kv)
      kv = dst.kv
    }
    expect(collectNodeViewState('n1', kv)).toEqual({
      tab: 'book',
      refPanelOpen: false,
    })
  })

  it('triggerRules 往返保结构（含嵌套对象）', () => {
    const rules = [{ id: 'r1', when: { field: 'f1' }, then: ['a', 'b'] }]
    const src = memKV({ 'trigrules:n1': JSON.stringify(rules) })
    const vs = collectNodeViewState('n1', src.kv)
    const dst = memKV()
    applyNodeViewState('n1', vs, dst.kv)
    expect(JSON.parse(dst.m.get('trigrules:n1') as string)).toEqual(rules)
  })
})

describe('viewState · 冲突审计', () => {
  it('本地与工程不一致时列出分歧', () => {
    const { kv } = memKV({ 'ne-tab:n1': 'book' })
    const out = auditViewStateConflicts('n1', { tab: 'main' }, kv)
    expect(out).toEqual([{ nodeId: 'n1', field: 'tab', local: 'book', fromProject: 'main' }])
  })

  it('一致时无分歧', () => {
    const { kv } = memKV({ 'ne-tab:n1': 'book' })
    expect(auditViewStateConflicts('n1', { tab: 'book' }, kv)).toEqual([])
  })

  it('任一侧缺该字段则不算分歧', () => {
    const { kv } = memKV({ 'ne-tab:n1': 'book' })
    // 工程没给 tab → 无从冲突
    expect(auditViewStateConflicts('n1', { rightPanelOpen: true }, kv)).toEqual([])
  })

  it('任一侧整体缺失时返回空', () => {
    const { kv } = memKV({ 'ne-tab:n1': 'book' })
    expect(auditViewStateConflicts('n1', undefined, kv)).toEqual([])
    expect(auditViewStateConflicts('n1', { tab: 'main' }, memKV().kv)).toEqual([])
  })
})

describe('viewState · 映射表自检', () => {
  it('★ 键位映射与控件实测用法一致（防前缀写错）', () => {
    // 这几个前缀是从 NodeEditor / WorkbookEditor / WordReportPanel 的实际代码里抄的。
    // 任何人改了那些文件里的键名，这里必须跟着改，否则工程会静默不带该状态。
    const prefixes = new Set(VIEW_KEYMAP.map((b) => b.prefix))
    expect(prefixes).toContain('ne-tab')
    expect(prefixes).toContain('ooref')
    expect(prefixes).toContain('wb-sheet')
    expect(prefixes).toContain('refpanel')
    expect(prefixes).toContain('rightpanel')
    expect(prefixes).toContain('wrp-refpanel')
    expect(prefixes).toContain('wrp-rightpanel')
    expect(prefixes).toContain('trigrules')
  })

  it('前缀不重复（重复会导致后写的静默覆盖前一个）', () => {
    const seen = new Set<string>()
    for (const b of VIEW_KEYMAP) {
      expect(seen.has(b.prefix)).toBe(false)
      seen.add(b.prefix)
    }
  })

  it('每项都指向合法字段', () => {
    const legal = new Set(['tab', 'refPanelOpen', 'rightPanelOpen', 'activeSheet', 'triggerRules'])
    for (const b of VIEW_KEYMAP) expect(legal.has(b.field)).toBe(true)
  })

  it('★ 同字段的 codec 必须一致（不一致时往返必然坏）', () => {
    // 例：把 refPanelOpen 的某一项标成 json，恢复时写出 '[object Object]'。
    // 单一真相源在这里给出保证：按字段分组检查。
    const byField = new Map<string, Set<string>>()
    for (const b of VIEW_KEYMAP) {
      const s = byField.get(b.field) ?? new Set<string>()
      s.add(b.codec)
      byField.set(b.field, s)
    }
    for (const [f, s] of byField) expect(s.size, `字段 ${f} 的 codec 不唯一`).toBe(1)
  })

  it('viewStateCoverage 覆盖全部字段', () => {
    const cov = viewStateCoverage()
    expect(cov.map((c) => c.field).sort()).toEqual(
      ['activeSheet', 'collapsed', 'refPanelOpen', 'rightPanelOpen', 'tab', 'triggerRules'].sort(),
    )
  })

  it('browserKV 适配器存在且不缓存 localStorage 引用', () => {
    // 只验接口形状：替换 globalThis.localStorage 后必须立刻生效
    const g = globalThis as unknown as { localStorage?: Storage }
    const orig = g.localStorage
    try {
      g.localStorage = {
        getItem: (k: string) => (k === 'ne-tab:x' ? 'book' : null),
        setItem: () => {},
        removeItem: () => {},
        length: 0,
        clear: () => {},
        key: () => null,
      } as unknown as Storage
      expect(collectNodeViewState('x', browserKV())?.tab).toBe('book')
    } finally {
      g.localStorage = orig
    }
  })
})
