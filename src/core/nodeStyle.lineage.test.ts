/**
 * 血缘（衍生关系）一致性测试
 * ============================================================================
 * 用户 2026-10-05：「在节点系统上增加一个衍生的关联标记
 * 　哪个是哪个节点衍生的，做了哪些适配。都有记录。」
 *
 * 为什么这些断言值得写死：
 * 血缘声明一旦写歪（父kind 不存在、代际名重复、adaptations 写「优化一下」这种
 * 没有执行力的废话），**没人会发现** —— 它只在 UI 上显示一行字，不影响运行。
 * 所以要用测试把它钉死，让「随手改坏」变成「CI 报错」。
 */

import { describe, it, expect } from 'vitest'
import {
  KIND_STYLE,
  derivesFromOf,
  isDerivedKind,
  isOOLineage,
  lineageChainOf,
  derivedKinds,
} from './nodeStyle'

describe('血缘声明 · 基本性质', () => {
  it('原始节点没有血缘（derivesFromOf 返回 null，不是 undefined）', () => {
    expect(derivesFromOf('table')).toBeNull()
    expect(derivesFromOf('param')).toBeNull()
    expect(derivesFromOf('xlsx')).toBeNull()
    expect(derivesFromOf('oosheet')).toBeNull()
  })

  it('未登记的 kind 返回 null，不抛（脏配置不该让画布白屏）', () => {
    expect(derivesFromOf('根本不存在的kind')).toBeNull()
    expect(derivesFromOf('')).toBeNull()
    expect(derivesFromOf(undefined)).toBeNull()
  })

  it('chapter 声明派生自 oosheet（用户 2026-10-05 定案：ooec 从 ooexcel 改的）', () => {
    const d = derivesFromOf('chapter')
    expect(d).not.toBeNull()
    expect(d!.kind).toBe('oosheet')
    expect(d!.generation).toBe('ooec')
  })

  it('isDerivedKind 能区分衍生节点与原始节点', () => {
    expect(isDerivedKind('chapter')).toBe(true)
    expect(isDerivedKind('oosheet')).toBe(false)
    expect(isDerivedKind('table')).toBe(false)
  })
})

describe('血缘声明 · 结构性约束（防写歪）', () => {
  it('★ 父 kind 必须真实存在于 KIND_STYLE', () => {
    // 父 kind 拼错 → UI 上会渲染出一条指向空气的血缘，且无人察觉
    for (const k of derivedKinds()) {
      const d = KIND_STYLE[k].derivesFrom!
      expect(Object.keys(KIND_STYLE)).toContain(d.kind)
      expect(KIND_STYLE[d.kind]).toBeDefined()
    }
  })

  it('★ 不得自我派生（a→a 会让血缘链成环）', () => {
    for (const k of derivedKinds()) {
      expect(KIND_STYLE[k].derivesFrom!.kind).not.toBe(k)
    }
  })

  it('★ 代际名全局唯一（两节点不能用同一个 generation）', () => {
    const gens = derivedKinds().map((k) => KIND_STYLE[k].derivesFrom!.generation)
    expect(new Set(gens).size).toBe(gens.length)
  })

  it('★ 每条 adaptations 必须写清「改哪里」——空 area 或空 note 无执行力', () => {
    for (const k of derivedKinds()) {
      const ad = KIND_STYLE[k].derivesFrom!.adaptations
      expect(ad.length).toBeGreaterThan(0)
      for (const a of ad) {
        expect(a.area.trim().length).toBeGreaterThan(0)
        expect(a.note.trim().length).toBeGreaterThan(0)
        // 过短的说明等于没写（如 note: '优化' / 'fix'）
        expect(a.note.trim().length).toBeGreaterThan(2)
        expect(a.area.trim().length).toBeGreaterThan(1)
      }
    }
  })

  it('★ 血缘链不得成环（lineageChainOf 能正常终止）', () => {
    for (const k of derivedKinds()) {
      const chain = lineageChainOf(k)
      expect(chain.length).toBeLessThanOrEqual(6)// 防御上限 5 步 + 自身
      expect(new Set(chain).size).toBe(chain.length) // 无重复 = 无环
    }
  })
})

describe('血缘链 · 遍历', () => {
  it('chapter 链为 chapter → oosheet', () => {
    expect(lineageChainOf('chapter')).toEqual(['chapter', 'oosheet'])
  })

  it('原始节点链只有自己', () => {
    expect(lineageChainOf('table')).toEqual(['table'])
    expect(lineageChainOf('oosheet')).toEqual(['oosheet'])
  })

  it('空/未登记 kind 返回空链', () => {
    expect(lineageChainOf('')).toEqual([])
    expect(lineageChainOf(undefined)).toEqual([])
  })
})

describe('OO 影响面判定 · 扫描器与 UI 同源', () => {
  it('OO 节点与其衍生节点都算 OO 血缘', () => {
    expect(isOOLineage('oosheet')).toBe(true)
    expect(isOOLineage('ooword')).toBe(true)
    expect(isOOLineage('chapter')).toBe(true) // 派生自 oosheet
  })

  it('非 OO 节点不算', () => {
    expect(isOOLineage('table')).toBe(false)
    expect(isOOLineage('xlsx')).toBe(false) // 自研内核，与 OO 无关
    expect(isOOLineage(undefined)).toBe(false)
  })
})

describe('定案口径 · 防止回退到错误配置', () => {
  it('★ chapter 是数据处理型（role=process，可有下游）', () => {
    // 用户 2026-10-05：「ooec 是数据处理型节点」
    // 我曾错建 output + terminal:true + allowsDownstream:false
    const s = KIND_STYLE.chapter
    expect(s.role).toBe('process')
    expect(s.terminal).toBe(false)
    expect(s.allowsDownstream).toBe(true)
  })

  it('★ chapter 的徽章文案是「OO 材料」', () => {
    expect(KIND_STYLE.chapter.label).toBe('OO 材料')
  })

  it('★ 衍生节点都带 adaptations 记录（用户：都有记录）', () => {
    const d = derivesFromOf('chapter')!
    expect(d.adaptations.length).toBeGreaterThanOrEqual(5)
    // 每条都要指向可核对的位置
    const areas = d.adaptations.map((a) => a.area)
    expect(areas).toContain('插件通道')
    expect(areas).toContain('章节分级')
  })
})
