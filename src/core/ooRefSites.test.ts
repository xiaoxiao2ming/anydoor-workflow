/**
 * 域落点解析测试
 *
 * 立档：2026-10-06 · 配套 src/core/ooRefSites.ts
 */

import { describe, it, expect } from 'vitest'
import { locatorOf, normalizeRef, sheetPartOf, resolveOoRefSites, describeSites } from './ooRefSites'
import { ooRefUrn, type OoRef } from './ooRef'
import type { XlsxFieldSpec } from '../xlsx/field'

const ref = (nodeId: string, fieldId: string): OoRef => ({
  nodeId, fieldId, urn: ooRefUrn(nodeId, fieldId),
})

describe('① URN → 域名', () => {
  it('取 # 之后的部分', () => {
    expect(locatorOf('node://t1#f1')).toBe('f1')
    expect(locatorOf('node://t1#')).toBe('')
  })
  it('没有 # 时返回整体', () => {
    expect(locatorOf('const:abc')).toBe('const:abc')
  })
})

describe('★★② ref 规整：重复表名那个坑', () => {
  it('definedName（带表名）→ 不重复补前缀', () => {
    const r = normalizeRef(undefined, '汇总表!$B$2')
    expect(r).toEqual({ part: '汇总表', a1: 'B2' })
    // ★ 关键断言：不能变成 '汇总表!汇总表!$B$2'
    expect(r!.a1).not.toContain('!')
  })

  it('纯 A1 + 有 sheet 名 → 补表名', () => {
    expect(normalizeRef('Sheet1', '$C$3')).toEqual({ part: 'Sheet1', a1: 'C3' })
  })

  it('区域取左上角（不能当单格写）', () => {
    expect(normalizeRef('Sheet1', 'A1:C4')).toEqual({ part: 'Sheet1', a1: 'A1' })
    expect(normalizeRef('Sheet1', '汇总表!$A$1:$D$9')!.a1).toBe('A1')
  })

  it('空/无效输入返回 null（不崩）', () => {
    expect(normalizeRef('S', '')).toBeNull()
    expect(normalizeRef('S', '   ')).toBeNull()
  })

  it('单引号包裹的表名要去掉', () => {
    expect(normalizeRef(undefined, "'我的表'!$A$1")!.part).toBe('我的表')
  })
})

describe('③ 工作表名 → 文件名', () => {
  it('表名本身是 sheetN 形式 ⇒ 直接用', () => {
    expect(sheetPartOf('sheet2', -1)).toBe('sheet2.xml')
  })
  it('中文表名按序号兜底', () => {
    expect(sheetPartOf('汇总表', 0)).toBe('sheet1.xml')
    expect(sheetPartOf('汇总表', 2)).toBe('sheet3.xml')
  })
})

describe('★★ ④ 核心：ooRefs × 落点清单 ⇒ sites', () => {
  // ★★ 关键口径：XlsxFieldSpec.key 是**命名区域的名字**（discoverXlsxFields: key = n.name），
  //   而 OoRef.fieldId 是节点字段 id —— **两者不是同一套标识**。
  //   所以必须提供 nameOf（fieldId → 字段名）才能对上。
  //   （第一版测试没给 nameOf，直接拿 f1 去对「数量」⇒ 永远匹配不上，
  //     3 条红暴露了我实现里的同一处错误，已一并修掉。）
  const fields: XlsxFieldSpec[] = [
    { key: '数量', kind: 'definedName', ref: '汇总表!$B$2', sheet: '汇总表' },
    { key: '单价', kind: 'definedName', ref: '汇总表!$B$3', sheet: '汇总表' },
  ]
  /** 字段名表：f1 → 数量，f2 → 单价 */
  const nameOf = (id: string) => ({ f1: '数量', f2: '单价' }[id] ?? '')

  it('字段名对上 ⇒ 解析出 a1', () => {
    const r = resolveOoRefSites([ref('t1', 'f1')], fields, 'oosheet', {
      nameOf, sheetPartMap: { 汇总表: 'sheet1.xml' },
    })
    expect(r.unmatched).toHaveLength(0)
    expect(r.sites).toHaveLength(1)
    expect(r.sites[0].a1).toBe('B2')
    expect(r.sites[0].part).toBe('sheet1.xml')
  })

  it('★ 没有 nameOf 时，如实报「试过哪些键」（不静默丢）', () => {
    const r = resolveOoRefSites([ref('t1', 'f1')], fields, 'oosheet')
    expect(r.sites).toHaveLength(0)
    expect(r.unmatched[0].why).toContain('f1')   // 把试过的键写出来，便于定位
  })

  it('字段名对不上 ⇒ 如实报 unmatched', () => {
    const r = resolveOoRefSites([ref('t1', '不存在')], fields, 'oosheet', { nameOf })
    expect(r.sites).toHaveLength(0)
    expect(r.unmatched).toHaveLength(1)
  })

  it('无 urn 的老数据也能解析（兜底 nodeId#fieldId + 字段名）', () => {
    const old: OoRef[] = [{ nodeId: 't1', fieldId: 'f1' }]  // ★ 无 urn
    const r = resolveOoRefSites(old, fields, 'oosheet', { nameOf })
    expect(r.sites).toHaveLength(1)
    expect(r.sites[0].a1).toBe('B2')
  })

  it('★ urn 的 locator 若是域名，可直接对上（无需 nameOf）', () => {
    const byName: OoRef[] = [{ nodeId: 't1', fieldId: 'f1', urn: 'node://t1#数量' }]
    const r = resolveOoRefSites(byName, fields, 'oosheet')
    expect(r.sites).toHaveLength(1)
    expect(r.sites[0].a1).toBe('B2')
  })

  it('多引用各拿各的落点', () => {
    const r = resolveOoRefSites(
      [ref('t1', 'f1'), ref('t1', 'f2')], fields, 'oosheet', { nameOf },
    )
    expect(r.sites.map((s) => s.a1)).toEqual(['B2', 'B3'])
  })

  it('同名key 多处落点时取第一处（与 domainMarks 口径一致）', () => {
    const dup: XlsxFieldSpec[] = [
      { key: 'K', kind: 'definedName', ref: 'S1!$A$1' },
      { key: 'K', kind: 'definedName', ref: 'S2!$B$9' },
    ]
    const r = resolveOoRefSites([{ nodeId: 't1', fieldId: 'x', urn: 'node://t1#K' }], dup, 'oosheet')
    expect(r.sites[0].a1).toBe('A1')
  })
})

describe('⑤ Word侧：按 key 匹配 sdt', () => {
  it('找到 ⇒ 带sdtTag', () => {
    const r = resolveOoRefSites(
      [ref('t1', 'f1')], [], 'ooword',
      { wordSites: [{ tag: '表甲.数量', key: 'f1' }] },
    )
    expect(r.sites).toHaveLength(1)
    expect(r.sites[0].sdtTag).toBe('表甲.数量')
  })
  it('找不到 ⇒ unmatched 说清是内容控件', () => {
    const r = resolveOoRefSites([ref('t1', 'f1')], [], 'ooword', { wordSites: [] })
    expect(r.unmatched[0].why).toContain('内容控件')
  })
  it('tag 直接等于 key 时也能匹配', () => {
    const r = resolveOoRefSites([ref('t1', 'f1')], [], 'ooword', { wordSites: [{ tag: 'f1' }] })
    expect(r.sites).toHaveLength(1)
  })
})

describe('⑥ 描述清单（给用户看落点）', () => {
  it('xlsx 格式化成 sheet1!B2', () => {
    expect(describeSites([{ ref: ref('t1','f1'), a1: 'B2', part: 'sheet1.xml' }])).toEqual(['sheet1!B2'])
  })
  it('word 格式化成内容控件 tag', () => {
    expect(describeSites([{ ref: ref('t1','f1'), sdtTag: '表甲.数量' }])).toEqual(['Word·内容控件 表甲.数量'])
  })
  it('无落点如实说', () => {
    expect(describeSites([{ ref: ref('t1','f1') }])).toEqual(['（无落点）'])
  })
})
