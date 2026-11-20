import { describe, it, expect } from 'vitest'
import { buildContextBlock, compressTurns, isAppTopic, isDateTimeQuestion, isFollowUp, isTopicSwitch, needsMoreInfo, nowLine, similarity } from './chatContext'
import type { CtxTurn } from './chatContext'

describe('nowLine', () => {
  it('输出含日期+星期+时间，供提示词注入与本地直答', () => {
    const line = nowLine(new Date(2026, 10, 8, 14, 5)) // 2026-11-08 是星期日
    expect(line).toContain('2026年11月8日')
    expect(line).toContain('星期日')
    expect(line).toContain('14:05')
  })
})

describe('isDateTimeQuestion（纯日期问题本地直答，不耗模型）', () => {
  it('纯日期/时间问法命中', () => {
    expect(isDateTimeQuestion('今天几号')).toBe(true)
    expect(isDateTimeQuestion('今天几号？')).toBe(true)
    expect(isDateTimeQuestion('今天星期几')).toBe(true)
    expect(isDateTimeQuestion('现在几点')).toBe(true)
    expect(isDateTimeQuestion('几点了')).toBe(true)
    expect(isDateTimeQuestion('今天是几月几号')).toBe(true)
    expect(isDateTimeQuestion('现在日期')).toBe(true)  // 现在 + 日期? → 见下
  })
  it('带相对推算/带后续动作的不拦（交给注入了当前时间的模型）', () => {
    expect(isDateTimeQuestion('明天几号')).toBe(false)
    expect(isDateTimeQuestion('下周一星期几')).toBe(false)
    expect(isDateTimeQuestion('今天几号发货')).toBe(false)
    expect(isDateTimeQuestion('星期三开会')).toBe(false)
    expect(isDateTimeQuestion('几点了还不睡')).toBe(false)
  })
})

const weatherTurns: CtxTurn[] = [
  { role: 'user', text: '今天的天气' },
  { role: 'assistant', text: '上海当前：大毛雨，15°C，风速 11 km/h。今日 15~19°C，降水概率 20%。' },
]

describe('isAppTopic', () => {
  it('节点/PDF/表格 属于本软件主题', () => {
    expect(isAppTopic('N071节点是什么')).toBe(true)
    expect(isAppTopic('把这几份PDF合并')).toBe(true)
    expect(isAppTopic('表格导出成 excel')).toBe(true)
  })
  it('穿衣/天气闲聊不算软件主题', () => {
    expect(isAppTopic('穿多少衣服')).toBe(false)
    expect(isAppTopic('今天心情不错')).toBe(false)
  })
})

describe('isFollowUp', () => {
  it('指代词与疑问词算追问', () => {
    expect(isFollowUp('穿多少衣服')).toBe(true) // 「多少」+ 极短句
    expect(isFollowUp('那明天呢')).toBe(true)
    expect(isFollowUp('它为什么报错')).toBe(true)
  })
  it('完整的新指令不算追问', () => {
    expect(isFollowUp('帮我把这三份扫描件合成一个PDF文件')).toBe(false)
  })
})

describe('similarity', () => {
  it('同话题高分、跨话题 0', () => {
    expect(similarity('N071节点', 'N071节点跑一次')).toBeGreaterThan(0.3)
    expect(similarity('今天的天气', '把PDF合并')).toBe(0)
  })
})

describe('isTopicSwitch', () => {
  it('「穿多少衣服」接着天气聊 → 不算换话题', () => {
    expect(isTopicSwitch(weatherTurns, '穿多少衣服')).toBe(false)
  })
  it('从天气跳到 PDF 合并 → 换话题', () => {
    expect(isTopicSwitch(weatherTurns, '帮我把这三份扫描件合成一个PDF文件')).toBe(true)
  })
  it('完全同题（N071 追问）→ 没换', () => {
    const t: CtxTurn[] = [{ role: 'user', text: 'N071节点是什么' }, { role: 'assistant', text: 'MERGED' }]
    expect(isTopicSwitch(t, 'N071执行一次')).toBe(false)
  })
  it('没有历史 → 不算换话题', () => {
    expect(isTopicSwitch([], '随便说点什么')).toBe(false)
  })
})

describe('buildContextBlock', () => {
  it('同话题时把上一轮天气带进上下文（这是截图里的病根）', () => {
    const r = buildContextBlock(weatherTurns, '穿多少衣服')
    expect(r.switched).toBe(false)
    expect(r.block).toContain('上海当前：大毛雨')
    expect(r.block).toContain('必须结合它回答')
  })
  it('换话题时只留一句背景提示，不塞旧内容', () => {
    const r = buildContextBlock(weatherTurns, '帮我把这三份扫描件合成一个PDF文件')
    expect(r.switched).toBe(true)
    expect(r.block).not.toContain('大毛雨')
  })
  it('没有历史 → 空块', () => {
    expect(buildContextBlock([], '你好').block).toBe('')
  })
  it('占位消息（正在搜…/思考中）不进上下文', () => {
    const r = buildContextBlock([{ role: 'assistant', text: '（思考中…）' }, ...weatherTurns], '明天呢')
    expect(r.block).not.toContain('思考中')
  })
})

describe('compressTurns', () => {
  it('超预算时压缩更早的对话并给出省略提示', () => {
    const many: CtxTurn[] = []
    for (let i = 0; i < 20; i++) {
      many.push({ role: 'user', text: `第${i}个问题：` + '这是一个很长的问题描述'.repeat(4) })
      many.push({ role: 'assistant', text: `第${i}个回答：` + '这是一个很长的回答内容'.repeat(4) })
    }
    const r = compressTurns(many, { budgetChars: 600, keepRecent: 4 })
    expect(r.compressed).toBe(true)
    expect(r.lines.join('\n')).toContain('已压缩省略')
    /* 最近的对话必须保留 */
    expect(r.lines.join('\n')).toContain('第19个回答')
  })
  it('短对话不压缩', () => {
    const r = compressTurns(weatherTurns, { budgetChars: 900 })
    expect(r.compressed).toBe(false)
    expect(r.lines).toHaveLength(2)
  })
  it('单条超长会截断并加省略号', () => {
    const r = compressTurns([{ role: 'assistant', text: '啊'.repeat(500) }], { budgetChars: 900, recentChars: 100 })
    expect(r.lines[0].length).toBeLessThanOrEqual(120)
    expect(r.lines[0]).toContain('…')
  })
})

describe('needsMoreInfo', () => {
  it('没历史 + 极短追问 → 信息不足', () => {
    expect(needsMoreInfo([], '穿多少')).toBe(true)
  })
  it('有历史就不再判不足（交给模型结合上文）', () => {
    expect(needsMoreInfo(weatherTurns, '穿多少')).toBe(false)
  })
  it('明确指令不算不足', () => {
    expect(needsMoreInfo([], '把 PDF 合并')).toBe(false)
  })
})
