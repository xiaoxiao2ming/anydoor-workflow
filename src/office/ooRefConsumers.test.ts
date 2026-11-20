/**
 * 域外写回消费者注册表测试
 *
 * 立档：2026-10-29 · 配套 src/office/ooRefConsumers.ts
 *
 * 重点验两件易错的事：
 *  ① 注销后必须把空Set 从 map 里删掉（否则 map 无限增长 = 内存泄漏）
 *  ② 同一节点被多个编辑器实例引用时，一个注销不影响另一个
 */

import { describe, it, expect, beforeEach } from 'vitest'
import { ooRefConsumers, type OoRefConsumer } from './ooRefConsumers'

beforeEach(() => { ooRefConsumers._reset() })

describe('① 登记与查询', () => {
  it('登记后可查到（编辑方靠它决定要不要触发）', () => {
    expect(ooRefConsumers.get('t1')).toBeNull()
    const un = ooRefConsumers.register('t1', () => {})
    expect(ooRefConsumers.get('t1')!.size).toBe(1)
    un()
  })

  it('无人引用时返回 null（供调用方零成本跳过）', () => {
    expect(ooRefConsumers.get('t1')).toBeNull()
    expect(ooRefConsumers.get(undefined)).toBeNull()
  })

  it('isReferenced 反映真实引用状态', () => {
    expect(ooRefConsumers.isReferenced('t1')).toBe(false)
    const un = ooRefConsumers.register('t1', () => {})
    expect(ooRefConsumers.isReferenced('t1')).toBe(true)
    un()
    expect(ooRefConsumers.isReferenced('t1')).toBe(false)
  })
})

describe('★★ ② 注销必须清掉空 Set（防内存泄漏）', () => {
  it('注销后 map 里不再有这个 key', () => {
    const un = ooRefConsumers.register('t1', () => {})
    expect(ooRefConsumers._size()).toBe(1)
    un()
    // ★ 关键：不是 set.size===0，而是**整个 key 消失**
    expect(ooRefConsumers._size()).toBe(0)
    expect(ooRefConsumers.get('t1')).toBeNull()
  })

  it('注销函数可重复调用而不报错（React 严格模式会双调用）', () => {
    const un = ooRefConsumers.register('t1', () => {})
    un()
    expect(() => un()).not.toThrow()
    expect(ooRefConsumers._size()).toBe(0)
  })
})

describe('③ 多实例：同一节点被多个 OO 节点引用', () => {
  it('一个注销不影响另一个', () => {
    let n = 0
    const a: OoRefConsumer = () => { n++ }
    const b: OoRefConsumer = () => { n++ }
    const unA = ooRefConsumers.register('t1', a)
    ooRefConsumers.register('t1', b)
    expect(ooRefConsumers.get('t1')!.size).toBe(2)

    unA()                       // A 走了（编辑器关闭）
    expect(ooRefConsumers.get('t1')!.size).toBe(1)   // B 还在
    ooRefConsumers.get('t1')!.forEach((f) => void f())
    expect(n).toBe(1)            // 只触发 B
  })

  it('同一回调重复登记只算一次（Set 语义）', () => {
    const fn: OoRefConsumer = () => {}
    const un1 = ooRefConsumers.register('t1', fn)
    const un2 = ooRefConsumers.register('t1', fn)
    expect(ooRefConsumers.get('t1')!.size).toBe(1)
    un1(); un2()
    expect(ooRefConsumers._size()).toBe(0)
  })
})

describe('④ 触发语义', () => {
  it('get 返回的集合可安全遍历（只读语义，不影响注册表）', () => {
    const un = ooRefConsumers.register('t1', () => {})
    const s = ooRefConsumers.get('t1')!
    s.forEach(() => {})
    expect(ooRefConsumers.get('t1')!.size).toBe(1)   // 遍历后仍在
    un()
  })

  it('异步消费者被触发时不阻塞（返回 Promise 也安全）', async () => {
    const un = ooRefConsumers.register('t1', async () => { await Promise.resolve() })
    const s = ooRefConsumers.get('t1')!
    s.forEach((f) => { void f() })   // 调用方用 void 忽略 Promise
    expect(ooRefConsumers.get('t1')!.size).toBe(1)
    un()
  })
})
