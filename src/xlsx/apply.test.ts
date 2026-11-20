import { describe, it, expect } from 'vitest'
import { createWorkbook, addWorksheet, getSheet, type Workbook } from '@office-kit/xlsx/workbook'
import { setCellByCoord, getCellByCoord } from '@office-kit/xlsx/worksheet'
import { setFormula } from '@office-kit/xlsx/cell'
import { applyTriggerPlan, normalizeForCell, formulaAt } from './apply'
import { planTrigger, type TriggerRule } from './trigger'

/** 照 field.test.ts 的建簿口径：createWorkbook + addWorksheet（一步到位） */
function wbWith(sheets: string[]): Workbook {
  const wb = createWorkbook()
  for (const s of sheets) addWorksheet(wb, s)
  return wb
}

/** 取工作表（不存在直接抛 —— 测试里静默才是 bug） */
function wsOf(wb: Workbook, name: string) {
  const s = getSheet(wb, name)
  if (!s) throw new Error('no sheet ' + name)
  return s
}

/** 读格值：公式格取缓存值，规避 value.kind 的类型细节 */
function cellValue(wb: Workbook, sheet: string, coord: string): unknown {
  const c = getCellByCoord(wsOf(wb, sheet), coord) as { value?: unknown } | undefined
  const v = c?.value
  if (v && typeof v === 'object' && (v as { kind?: string }).kind === 'formula') {
    return (v as { cachedValue?: unknown }).cachedValue
  }
  return v
}

const R = (p: Partial<TriggerRule>): TriggerRule => ({
  id: 'r1',
  name: 'r',
  enabled: true,
  watch: 'param://a',
  condition: { op: 'changed' },
  action: { kind: 'writeCell', sheet: 'Sheet1', coord: 'A1' },
  ...p,
})

/* ── normalizeForCell ── */

describe('normalizeForCell', () => {
  it('标量原样返回', () => {
    expect(normalizeForCell(5)).toBe(5)
    expect(normalizeForCell('x')).toBe('x')
    expect(normalizeForCell(true)).toBe(true)
  })

  it('null / undefined → 空串（OO 里 cell 必须有值）', () => {
    expect(normalizeForCell(null)).toBe('')
    expect(normalizeForCell(undefined)).toBe('')
  })

  it('非有限数 → 空串（NaN/Infinity 会产出非法 OOXML）', () => {
    expect(normalizeForCell(NaN)).toBe('')
    expect(normalizeForCell(Infinity)).toBe('')
  })

  it('对象 / 数组 → JSON 文本（直接塞对象会产出非法 cell）', () => {
    expect(normalizeForCell({ a: 1 })).toBe('{"a":1}')
    expect(normalizeForCell([1, 2])).toBe('[1,2]')
  })

  it('Date → 日期串', () => {
    expect(normalizeForCell(new Date('2026-10-05T00:00:00Z'))).toBe('2026-10-05')
  })

  it('循环引用对象不炸（退回 String）', () => {
    const a: Record<string, unknown> = {}
    a.self = a
    expect(() => normalizeForCell(a)).not.toThrow()
  })
})

/* ── 写值 ── */

describe('applyTriggerPlan · 写值', () => {
  const ch = [{ urn: 'param://a', before: 1, after: 2 }]

  it('真写进单元格', () => {
    const wb = wbWith(['Sheet1'])
    const p = planTrigger([R({})], ch, () => 0)
    const r = applyTriggerPlan(wb, p)
    expect(r.written).toBe(1)
    expect(cellValue(wb, 'Sheet1', 'A1')).toBe(2)
  })

  it('dryRun 不真写，但如实回报会写几格', () => {
    const wb = wbWith(['Sheet1'])
    const p = planTrigger([R({})], ch, () => 0)
    const r = applyTriggerPlan(wb, p, { dryRun: true })
    expect(r.written).toBe(1)
    expect(cellValue(wb, 'Sheet1', 'A1')).toBeUndefined() // 真的没写
  })

  it('工作表不存在 → 跳过并带原因（不静默吞）', () => {
    const wb = wbWith(['Sheet1'])
    const p = planTrigger(
      [R({ action: { kind: 'writeCell', sheet: '没有这张表', coord: 'A1' } })],
      ch,
      () => 0
    )
    const r = applyTriggerPlan(wb, p)
    expect(r.written).toBe(0)
    expect(r.skipped[0].reason).toContain('不存在')
  })

  it('非法 A1 坐标 → 跳过并带原因', () => {
    const wb = wbWith(['Sheet1'])
    const p = planTrigger(
      [R({ action: { kind: 'writeCell', sheet: 'Sheet1', coord: '不是坐标!' } })],
      ch,
      () => 0
    )
    const r = applyTriggerPlan(wb, p)
    expect(r.written).toBe(0)
    expect(r.skipped[0].reason).toContain('合法 A1')
  })

  it('★ 公式格不被静默覆盖（安全护栏）', () => {
    const wb = wbWith(['Sheet1'])
    setFormula(setCellByCoord(wsOf(wb, 'Sheet1'), 'A1')!, 'SUM(B1:B9)', { cachedValue: 3 })
    const p = planTrigger([R({})], ch, () => 0)
    const r = applyTriggerPlan(wb, p)
    expect(r.written).toBe(0)
    expect(r.skipped[0].reason).toContain('公式')
    // 公式还在，没被毁
    expect(formulaAt(wb, 'Sheet1', 'A1')).toBe('SUM(B1:B9)')
  })

  it('区域写入：数组铺多格', () => {
    const wb = wbWith(['Sheet1'])
    const p = planTrigger(
      [R({ action: { kind: 'writeRange', sheet: 'Sheet1', ref: 'A1:A3' } })],
      [{ urn: 'param://a', before: [1, 2], after: [7, 8, 9] }],
      () => 0
    )
    const r = applyTriggerPlan(wb, p)
    expect(r.written).toBe(3)
    expect(cellValue(wb, 'Sheet1', 'A1')).toBe(7)
    expect(cellValue(wb, 'Sheet1', 'A3')).toBe(9)
  })

  it('空计划 → 什么都不做也不炸', () => {
    const wb = wbWith(['Sheet1'])
    const r = applyTriggerPlan(wb, {
      writes: [],
      recomputeSheets: new Set(),
      recomputeAll: false,
      needCommit: false,
      hits: [],
      log: [],
    })
    expect(r.written).toBe(0)
    expect(r.skipped).toHaveLength(0)
  })
})

/* ── 重算 ── */

describe('applyTriggerPlan · 重算', () => {
  it('重算指定工作表并回报结果', () => {
    const wb = wbWith(['Sheet1'])
    setCellByCoord(wsOf(wb, 'Sheet1'), 'A1', 4)
    setFormula(setCellByCoord(wsOf(wb, 'Sheet1'), 'B1')!, 'A1*2')

    const p = planTrigger(
      [
        R({
          id: 'rc',
          watch: 'change',
          condition: { op: 'always' },
          action: { kind: 'recompute', sheet: 'Sheet1' },
        }),
      ],
      [{ urn: 'param://a', before: 1, after: 2 }],
      () => 0
    )
    const r = applyTriggerPlan(wb, p)
    expect(r.recomputed['Sheet1']).toBeDefined()
    expect(r.recomputed['Sheet1'].updated).toBeGreaterThan(0)
  })

  it('recomputeAll → 所有工作表都算', () => {
    const wb = wbWith(['S1', 'S2', 'S3'])
    const p = planTrigger(
      [R({ watch: 'change', condition: { op: 'always' }, action: { kind: 'recompute' } })],
      [{ urn: 'param://a', before: 1, after: 2 }],
      () => 0
    )
    const r = applyTriggerPlan(wb, p)
    expect(Object.keys(r.recomputed).sort()).toEqual(['S1', 'S2', 'S3'])
  })

  it('重算不存在的工作表 → 跳过并带原因', () => {
    const wb = wbWith(['S1'])
    const r = applyTriggerPlan(wb, {
      writes: [],
      recomputeSheets: new Set(['缺失表']),
      recomputeAll: false,
      needCommit: false,
      hits: [],
      log: [],
    })
    expect(r.skipped[0].reason).toContain('不存在')
  })

  it('★ 端到端：参数变 → 写 A1 → 自动重算 → B1 缓存变 42', () => {
    const wb = wbWith(['Sheet1'])
    setFormula(setCellByCoord(wsOf(wb, 'Sheet1'), 'B1')!, 'A1*2', { cachedValue: 0 })

    const p = planTrigger(
      [R({ action: { kind: 'writeCell', sheet: 'Sheet1', coord: 'A1' } })],
      [{ urn: 'param://a', before: 1, after: 21 }],
      () => 0
    )
    expect(p.recomputeSheets.has('Sheet1')).toBe(true) // 自动补上了
    const r = applyTriggerPlan(wb, p)
    expect(r.written).toBe(1)
    expect(cellValue(wb, 'Sheet1', 'B1')).toBe(42)
  })

  it('resolveUrn 注入：node:// 引用参与求值', () => {
    const wb = wbWith(['Sheet1'])
    setCellByCoord(wsOf(wb, 'Sheet1'), 'A1', 'node://n1#rows')
    setFormula(setCellByCoord(wsOf(wb, 'Sheet1'), 'B1')!, 'A1*2', { cachedValue: 0 })

    const p = planTrigger(
      [
        R({
          id: 'rc2',
          watch: 'change',
          condition: { op: 'always' },
          action: { kind: 'recompute', sheet: 'Sheet1' },
        }),
      ],
      [{ urn: 'node://n1#rows', before: 1, after: 3 }],
      () => 0
    )
    applyTriggerPlan(wb, p, { resolveUrn: () => 3 })
    expect(cellValue(wb, 'Sheet1', 'B1')).toBe(6)
  })
})

/* ── formulaAt ── */

describe('formulaAt', () => {
  it('公式格返回公式文本', () => {
    const wb = wbWith(['S'])
    setFormula(setCellByCoord(wsOf(wb, 'S'), 'A1')!, 'SUM(B:B)')
    expect(formulaAt(wb, 'S', 'A1')).toBe('SUM(B:B)')
  })

  it('普通格 / 空格 / 不存在的表 → null', () => {
    const wb = wbWith(['S'])
    setCellByCoord(wsOf(wb, 'S'), 'A1', 1)
    expect(formulaAt(wb, 'S', 'A1')).toBeNull()
    expect(formulaAt(wb, 'S', 'Z9')).toBeNull()
    expect(formulaAt(wb, '没有', 'A1')).toBeNull()
  })
})
