import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  dropPin,
  hasPin,
  loadPins,
  PRIO_KEY,
  savePins,
  STAR_KEY,
  subscribePins,
  upsertPin,
  type RefPin,
} from './refPins'

/** node 环境无 localStorage —— 内存 stub，按 key 隔离 */
function stubStorage(): Map<string, string> {
  const m = new Map<string, string>()
  ;(globalThis as Record<string, unknown>).localStorage = {
    getItem: (k: string) => (m.has(k) ? m.get(k)! : null),
    setItem: (k: string, v: string) => void m.set(k, v),
  }
  return m
}

describe('refPins · 优先/星标参数全局存储', () => {
  let backing: Map<string, string>
  beforeEach(() => {
    backing = stubStorage()
  })
  afterEach(() => {
    delete (globalThis as Record<string, unknown>).localStorage
  })

  it('空存储 → loadPins 返回空数组（两个 key 都如此）', () => {
    expect(loadPins(PRIO_KEY)).toEqual([])
    expect(loadPins(STAR_KEY)).toEqual([])
  })

  it('savePins 落盘 → loadPins 原样读回', () => {
    const pins: RefPin[] = [
      { nodeId: 'n1', fieldId: 'f1', ts: 1 },
      { nodeId: 'n2', fieldId: 'f2', ts: 2 },
    ]
    savePins(PRIO_KEY, pins)
    expect(loadPins(PRIO_KEY)).toEqual(pins)
    // 另一个 key 互不干扰
    expect(loadPins(STAR_KEY)).toEqual([])
  })

  it('损坏 JSON / 非数组 / 非法项逐项过滤，绝不抛错', () => {
    backing.set(PRIO_KEY, '{oops')
    expect(loadPins(PRIO_KEY)).toEqual([])
    backing.set(PRIO_KEY, JSON.stringify({ a: 1 }))
    expect(loadPins(PRIO_KEY)).toEqual([])
    backing.set(PRIO_KEY, JSON.stringify([{ nodeId: 'ok', fieldId: 'ok' }, { nodeId: '', fieldId: 'x' }, null, 42, { nodeId: 'n', fieldId: 3 }]))
    expect(loadPins(PRIO_KEY)).toEqual([{ nodeId: 'ok', fieldId: 'ok', ts: 0 }])
  })

  it('upsertPin：新项插头部；已存在不重复不重排', () => {
    let pins = upsertPin([], 'a', 'f1', 10)
    expect(pins).toEqual([{ nodeId: 'a', fieldId: 'f1', ts: 10 }])
    pins = upsertPin(pins, 'b', 'f2', 20)
    expect(pins.map((p) => p.nodeId)).toEqual(['b', 'a'])
    // 重复收录 → 原样返回（同一引用）
    const again = upsertPin(pins, 'a', 'f1', 99)
    expect(again).toBe(pins)
    expect(again.map((p) => p.ts)).toEqual([20, 10])
  })

  it('dropPin：只删命中项，原数组不被修改', () => {
    const pins: RefPin[] = [
      { nodeId: 'a', fieldId: 'f1', ts: 1 },
      { nodeId: 'b', fieldId: 'f2', ts: 2 },
    ]
    const next = dropPin(pins, 'a', 'f1')
    expect(next).toEqual([{ nodeId: 'b', fieldId: 'f2', ts: 2 }])
    expect(pins).toHaveLength(2)
    // 未命中 → 内容不变（新数组）
    expect(dropPin(pins, 'zz', 'zz')).toEqual(pins)
  })

  it('hasPin 精确匹配 nodeId+fieldId', () => {
    const pins = upsertPin([], 'a', 'f1', 1)
    expect(hasPin(pins, 'a', 'f1')).toBe(true)
    expect(hasPin(pins, 'a', 'f2')).toBe(false)
    expect(hasPin(pins, 'b', 'f1')).toBe(false)
  })

  it('savePins 触发订阅回调；退订后不再触发', () => {
    let n = 0
    const off = subscribePins(() => void n++)
    savePins(PRIO_KEY, [])
    savePins(STAR_KEY, [])
    expect(n).toBe(2)
    off()
    savePins(PRIO_KEY, [])
    expect(n).toBe(2)
  })

  it('无 localStorage 环境：load 空数组、save 不抛错仍广播', () => {
    delete (globalThis as Record<string, unknown>).localStorage
    expect(loadPins(PRIO_KEY)).toEqual([])
    let hit = 0
    const off = subscribePins(() => void hit++)
    savePins(PRIO_KEY, [{ nodeId: 'x', fieldId: 'y', ts: 0 }])
    expect(hit).toBe(1)
    off()
  })
})
