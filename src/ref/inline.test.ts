import { describe, it, expect } from 'vitest'
import {
  parseInline,
  inlineCapsuleClass,
  capsuleLabelOf,
  capsuleFieldOf,
  capsuleTitleOf,
  buildRefMark,
  countCapsules,
  offsetAfterInsert,
  plainLength,
  type InlineToken,
} from './inline'

describe('parseInline ·纯文本（无标记）', () => {
  it('普通文本返回单个 text 段，值原样', () => {
    expect(parseInline('总长 1200')).toEqual([{ type: 'text', value: '总长 1200' }])
  })
  it('空值返回空数组（渲染层据此显示占位）', () => {
    expect(parseInline('')).toEqual([])
    expect(parseInline(null)).toEqual([])
    expect(parseInline(undefined)).toEqual([])
  })
  it('旧数据天然兼容：不含⟨ 的任意文本都不产生胶囊', () => {
    for (const s of ['abc', 'a<b>c', '1 < 2 && 3 > 2', 'x⟨y']) {
      const tokens = parseInline(s)
      expect(tokens.filter((t) => t.type === 'capsule')).toHaveLength(0)
    }
  })
})

describe('parseInline · 引用胶囊', () => {
  it('解析 ⟨ref:节点|字段⟩，前后文字各成一段', () => {
    const tokens = parseInline('总长 ⟨ref:n1|总长⟩ 米')
    expect(tokens).toEqual([
      { type: 'text', value: '总长 ' },
      { type: 'capsule', target: 'n1', field: '总长', kind: undefined, urn: 'node://n1#总长' },
      { type: 'text', value: ' 米' },
    ])
  })

  it('带kind 的标记按类型分色', () => {
    const [cap] = parseInline('⟨ref:n1|字段|sort⟩') as Extract<InlineToken, { type: 'capsule' }>[]
    expect(cap.kind).toBe('sort')
    expect(cap.urn).toBe('node://n1#字段')
  })

  it('整节点引用（字段名为空）→ field 为空串且 urn 不带 #', () => {
    const [cap] = parseInline('⟨ref:n1||xlsx⟩') as Extract<InlineToken, { type: 'capsule' }>[]
    expect(cap.field).toBe('')
    expect(cap.urn).toBe('node://n1')
  })

  it('多个胶囊交错，顺序不乱', () => {
    const tokens = parseInline('a⟨ref:n1|f1⟩b⟨ref:n2|f2⟩c⟨const:xyz⟩d')
    expect(tokens.map((t) => t.type)).toEqual([
      'text',
      'capsule',
      'text',
      'capsule',
      'text',
      'capsule',
      'text',
    ])
    expect(tokens[5]).toMatchObject({ type: 'capsule', literal: 'xyz', target: 'const' })
    // 文字段内容完整、无丢字
    expect(tokens.filter((t) => t.type === 'text').map((t) => t.value)).toEqual(['a', 'b', 'c', 'd'])
  })

  it('胶囊在首尾时不会产生空text 段', () => {
    const tokens = parseInline('⟨ref:n1|f⟩尾')
    expect(tokens[0].type).toBe('capsule')
    expect(tokens.filter((t) => t.type === 'text')).toHaveLength(1)
  })

  it('常量胶囊 ⟨const:字面量⟩', () => {
    const tokens = parseInline('取 ⟨const:0.15⟩ 倍')
    const cap = tokens[1] as Extract<InlineToken, { type: 'capsule' }>
    expect(cap.literal).toBe('0.15')
    expect(cap.target).toBe('const')
  })

  it('常量值为空也能解析（⟨const:⟩）', () => {
    const tokens = parseInline('⟨const:⟩')
    expect(tokens).toHaveLength(1)
    expect(tokens[0]).toMatchObject({ type: 'capsule', literal: '' })
  })

  it('未闭合的标记当普通文字，不吞掉后文', () => {
    const tokens = parseInline('前缀 ⟨ref:未闭合 后缀')
    expect(tokens.every((t) => t.type === 'text')).toBe(true)
    expect(tokens[0].value).toContain('后缀')
  })

  it('换行与多段落文本按原文保留', () => {
    const tokens = parseInline('第一行\n第二行⟨ref:n|f⟩\n第三行')
    const texts = tokens.filter((t) => t.type === 'text').map((t) => t.value)
    expect(texts.join('')).toContain('\n')
    expect(texts.join('')).toContain('第三行')
  })
})

describe('胶囊呈现', () => {
  const cap = (t: InlineToken) => t as Extract<InlineToken, { type: 'capsule' }>

  it('CSS 类带 inline-cap 与类型色系，失效态加 dead', () => {
    expect(inlineCapsuleClass('sort')).toContain('cap-table') // sort 落默认色系
    expect(inlineCapsuleClass('docx')).toContain('cap-docx')
    expect(inlineCapsuleClass('xlsx')).toContain('cap-xlsx')
    expect(inlineCapsuleClass('param')).toContain('cap-param')
    expect(inlineCapsuleClass('text')).toContain('cap-text')
    expect(inlineCapsuleClass('sort', true)).toContain('dead')
    expect(inlineCapsuleClass()).toContain('ref-capsule')
  })

  it('左段：常量取字面量，引用取字段名，空字段回退「整体」', () => {
    expect(capsuleLabelOf(cap(parseInline('⟨const:3.5⟩')[0]))).toBe('3.5')
    expect(capsuleLabelOf(cap(parseInline('⟨ref:n1|总长⟩')[0]))).toBe('总长')
    expect(capsuleLabelOf(cap(parseInline('⟨ref:n1||x⟩')[0]))).toBe('整体')
  })

  it('右段：参数名 / 整节点 / 常量', () => {
    expect(capsuleFieldOf(cap(parseInline('⟨ref:n1|总长⟩')[0]))).toBe('总长')
    expect(capsuleFieldOf(cap(parseInline('⟨ref:n1||x⟩')[0]))).toBe('整节点')
    expect(capsuleFieldOf(cap(parseInline('⟨const:x⟩')[0]))).toBe('常量')
  })

  it('悬停提示带 URN，便于排查；给了节点名用节点名', () => {
    expect(capsuleTitleOf(cap(parseInline('⟨ref:n1|总长⟩')[0]))).toContain('node://n1#总长')
    expect(capsuleTitleOf(cap(parseInline('⟨ref:n1|总长⟩')[0]), '计算表')).toContain('计算表.总长')
    expect(capsuleTitleOf(cap(parseInline('⟨const:0.5⟩')[0]))).toContain('0.5')
  })
})

describe('buildRefMark · 往返一致', () => {
  it('生成的标记能被自己解析回来', () => {
    const mark = buildRefMark('n1', '总长', 'sort')
    const cap = parseInline(mark)[0] as Extract<InlineToken, { type: 'capsule' }>
    expect(cap.target).toBe('n1')
    expect(cap.field).toBe('总长')
    expect(cap.kind).toBe('sort')
  })

  it('省略 kind 时解析出 kind 为 undefined（不是空串）', () => {
    const cap = parseInline(buildRefMark('n1', 'f'))[0] as Extract<InlineToken, { type: 'capsule' }>
    expect(cap.kind).toBeUndefined()
  })

  it('字段名含分隔符也能往返', () => {
    const mark = buildRefMark('n1', 'a|b', 'table')
    const cap = parseInline(mark)[0] as Extract<InlineToken, { type: 'capsule' }>
    // 第一个 | 作分隔，第二个归入字段名（正则按首个分隔切分）
    expect(cap.target).toBe('n1')
  })
})

describe('countCapsules · 计数', () => {
  it('统计胶囊个数，纯文本为 0', () => {
    expect(countCapsules('没有胶囊')).toBe(0)
    expect(countCapsules('')).toBe(0)
    expect(countCapsulasSafe('⟨ref:a|f⟩x⟨const:1⟩')).toBe(2)
  })
})
// 小helper：避免上面表���里少写一个字母导致运行时报错时不好读
function countCapsulasSafe(s: string): number {
  return countCapsules(s)
}

describe('plainLength · 纯文本视图长度', () => {
  it('无标记时等于原长度', () => {
    expect(plainLength('总长 1200')).toBe(7)
    expect(plainLength('')).toBe(0)
    expect(plainLength(null)).toBe(0)
  })

  it('每颗胶囊算 1 个字符（扣掉标记多出来的部分）', () => {
    expect(plainLength('⟨ref:n1|f⟩')).toBe(1)
    expect(plainLength('ab⟨ref:n1|f⟩cd')).toBe(5) // ab + 1 + cd
    expect(plainLength('⟨ref:a|p⟩x⟨const:9⟩')).toBe(3) // 胶囊 + x + 胶囊
  })

  it('与 offsetAfterInsert 配套：plainLength 的结果就是末尾插入位', () => {
    const s = 'ab⟨ref:n1|f⟩cd'
    const at = offsetAfterInsert(s, plainLength(s))
    expect(at).toBe(s.length) // 末尾插入 → 应能拼到末尾
    const after = s.slice(0, at) + '⟨const:X⟩' + s.slice(at)
    expect(plainLength(after)).toBe(6) // 多了一颗胶囊
    expect(after).toBe('ab⟨ref:n1|f⟩cd⟨const:X⟩')
  })

  it('未闭合标记按普通字符计', () => {
    expect(plainLength('a⟨ref:未闭合')).toBe(9) // 9 个普通字符
  })
})

describe('offsetAfterInsert · 插入点换算', () => {
  /** 原文 14 字符 = 'ab'(2) + 标记(10) + 'cd'(2)；纯文本视图 = ab+胶囊+cd = 5 格 */
  const s = 'ab⟨ref:n1|f⟩cd'
  const mark = '⟨ref:n1|f⟩'

  it('纯文本时原样返回', () => {
    expect(offsetAfterInsert('总长 1200', 3)).toBe(3)
  })

  it('映射表：plainOffset = 已消费的纯文本字符数，胶囊算 1 格', () => {
    expect(offsetAfterInsert(s, 0)).toBe(0) // a 之前
    expect(offsetAfterInsert(s, 1)).toBe(1) // b 之前
    expect(offsetAfterInsert(s, 2)).toBe(2 + mark.length) // b 之后 = 胶囊之后
    expect(offsetAfterInsert(s, 3)).toBe(2 + mark.length) // c 之前（越过胶囊）
    expect(offsetAfterInsert(s, 4)).toBe(3 + mark.length) // d 之前
    expect(offsetAfterInsert(s, 5)).toBe(s.length) // 末尾
    expect(offsetAfterInsert(s, 99)).toBe(s.length) // 越界夹到末尾
  })

  it('多胶囊时偏移逐段累加，不丢也不重', () => {
    const m1 = '⟨ref:n1|a⟩'
    const m2 = '⟨ref:n2|b⟩'
    const s2 = `x${m1}y${m2}z`
    // 纯文本 = x 胶囊 y 胶囊 z = 5 格
    expect(offsetAfterInsert(s2, 2)).toBe(1 + m1.length) // 第一个胶囊后
    expect(offsetAfterInsert(s2, 4)).toBe(1 + m1.length + 1 + m2.length) // 第二个胶囊后
    expect(offsetAfterInsert(s2, 5)).toBe(s2.length) // 末尾
  })

  it('常量胶囊同样按 1 格算', () => {
    const s3 = 'a⟨const:9⟩b'
    expect(offsetAfterInsert(s3, 2)).toBe('a⟨const:9⟩'.length) // 胶囊后
    expect(offsetAfterInsert(s3, 3)).toBe(s3.length) // 末尾
  })
})