import { describe, expect, it } from 'vitest'
import { fillSmartNumber, seriesDelta } from './fillSeries'

describe('seriesDelta 严格等差识别', () => {
  it('1,2,3 → delta 1', () => {
    expect(seriesDelta([1, 2, 3])).toBe(1)
  })
  it('1,3,5,7 → delta 2', () => {
    expect(seriesDelta([1, 3, 5, 7])).toBe(2)
  })
  it('等差可为负（5,3,1 → -2）', () => {
    expect(seriesDelta([5, 3, 1])).toBe(-2)
  })
  it('差不等（1,2,4）→ null（平铺，同 Excel 默认）', () => {
    expect(seriesDelta([1, 2, 4])).toBeNull()
  })
  it('单值 → null（1 平铺 1,1,1）', () => {
    expect(seriesDelta([1])).toBeNull()
  })
  it('含非数字（null）→ null', () => {
    expect(seriesDelta([1, null, 3])).toBeNull()
  })
  it('空片段 → null', () => {
    expect(seriesDelta([])).toBeNull()
  })
  it('小数等差（0.5,1,1.5 → 0.5）', () => {
    expect(seriesDelta([0.5, 1, 1.5])).toBe(0.5)
  })
})

describe('fillSmartNumber 目标格取值', () => {
  // 源块 A1:B2（r1=1,c1=1,r2=2,c2=2），nums 第一列 1,2,3 语义见下
  const o = { r1: 1, c1: 1, r2: 3, c2: 1 } // 单列 A1:A3
  const nums = [
    [1],
    [2],
    [3],
  ]

  it('纵向带：源 1,2,3 下拖 → 4,5,6（k=r-r1）', () => {
    expect(fillSmartNumber(nums, o, 4, 1)).toBe(4)
    expect(fillSmartNumber(nums, o, 5, 1)).toBe(5)
    expect(fillSmartNumber(nums, o, 6, 1)).toBe(6)
  })

  it('origin 内的格 → null（UI 循环本就 skip origin，不发序列值）', () => {
    expect(fillSmartNumber(nums, o, 2, 1)).toBeNull()
    expect(fillSmartNumber(nums, o, 1, 1)).toBeNull()
  })

  it('非序列列 → null（平铺）', () => {
    const nums2 = [
      [1],
      [2],
      [4],
    ]
    expect(fillSmartNumber(nums2, o, 4, 1)).toBeNull()
  })

  it('角块（r、c 同时超出 origin）→ null 平铺', () => {
    const o2 = { r1: 1, c1: 1, r2: 3, c2: 2 }
    const nums2 = [
      [1, 10],
      [2, 20],
      [3, 30],
    ]
    expect(fillSmartNumber(nums2, o2, 4, 3)).toBeNull()
  })

  it('横向带：行切片 1,2,3 右拖 → 4,5,6', () => {
    const o2 = { r1: 1, c1: 1, c2: 3, r2: 1 }
    const nums2 = [[1, 2, 3]]
    expect(fillSmartNumber(nums2, o2, 1, 4)).toBe(4)
    expect(fillSmartNumber(nums2, o2, 1, 5)).toBe(5)
  })

  it('横向带混文本行 → null', () => {
    const o2 = { r1: 1, c1: 1, c2: 3, r2: 1 }
    const nums2 = [[1, 2, null]]
    expect(fillSmartNumber(nums2, o2, 1, 4)).toBeNull()
  })

  it('多列块各列独立：A 列等差 B 列混文本，B 列平铺 A 列外推', () => {
    const o2 = { r1: 1, c1: 1, r2: 2, c2: 2 }
    const nums2 = [
      [1, 10],
      [2, null],
    ]
    expect(fillSmartNumber(nums2, o2, 3, 1)).toBe(3) // A 列等差
    expect(fillSmartNumber(nums2, o2, 3, 2)).toBeNull() // B 列混文本 → 平铺
  })
})
