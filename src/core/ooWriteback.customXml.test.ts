import { describe, it, expect } from 'vitest'
import { buildCustomXmlData } from './ooWriteback'
import type { OoRef } from './ooRef'
import type { OoRefSite } from './ooWriteback'

/** 构造一个 word site */
function site (tag: string, nodeId = 'n1', fieldId = 'f1'): OoRefSite {
  const ref: OoRef = { nodeId, fieldId }
  return { ref, sdtTag: tag }
}

describe('buildCustomXmlData（层② customXml 数据源）', () => {
  it('sites → XML：tag/值转义、命名空间固定', () => {
    const xml = buildCustomXmlData(
      [site('field:项目名称'), site('field:金额&备注', 'n2', 'f2')],
      {
        pick: (ref) => (ref.fieldId === 'f1' ? '燃气工程<一期>' : ''),
        labelOf: (ref) => 'L_' + ref.fieldId,
      },
    )
    expect(xml).toContain('<ooRefs xmlns="urn:client-base:ooRefs">')
    expect(xml).toContain('<entry k="field:项目名称">燃气工程&lt;一期&gt;</entry>')
    expect(xml).toContain('<entry k="field:金额&amp;备注"></entry>')
    expect(xml.startsWith('<?xml')).toBe(true)
  })

  it('同 tag 去重；无 sdtTag 时用 labelOf 兜底', () => {
    const xml = buildCustomXmlData(
      [site('t1'), site('t1'), { ref: { nodeId: 'n2', fieldId: 'f2' } }],
      { pick: () => 42, labelOf: () => '兜底标签' },
    )
    expect(xml.match(/<entry /g)?.length).toBe(2)
    expect(xml).toContain('k="兜底标签"')
  })

  it('空 sites / 全无标签 → 空壳（仍合法 XML）', () => {
    const xml = buildCustomXmlData([], { pick: () => 1 })
    expect(xml).toContain('<ooRefs xmlns="urn:client-base:ooRefs">')
    expect(xml).not.toContain('<entry')
  })
})
