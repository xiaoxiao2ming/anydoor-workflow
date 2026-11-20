import { describe, expect, it } from 'vitest'
import { matchIntent, extractWebQuery, isUsableQuery, fallbackQueryFromTurns, stripQueryPhrases } from './chatIntent'

describe('web-search 意图', () => {
  it('「搜一下 量子计算最新进展」→ web-search', () => {
    expect(matchIntent('搜一下 量子计算最新进展', [])?.id).toBe('web-search')
  })
  it('「帮我查一下 Rust 和 Go 的区别」→ web-search', () => {
    expect(matchIntent('帮我查一下 Rust 和 Go 的区别', [])?.id).toBe('web-search')
  })
  it('「今天有什么新闻」→ web-search', () => {
    expect(matchIntent('今天有什么新闻', [])?.id).toBe('web-search')
  })
  it('「查天气」仍然归 query-weather（不被 web-search 抢走）', () => {
    expect(matchIntent('查一下天气', [])?.id).toBe('query-weather')
    expect(matchIntent('北京天气怎么样', [])?.id).toBe('query-weather')
  })
  it('工作流指令不误伤：「把图片转成word」→ 非 web-search', () => {
    expect(matchIntent('把图片转成word', [])?.id).not.toBe('web-search')
  })
})

describe('extractWebQuery', () => {
  it('去掉指令前缀', () => {
    expect(extractWebQuery('帮我查一下 Rust 和 Go 的区别')).toBe('Rust 和 Go 的区别')
  })
  it('引号内容优先', () => {
    expect(extractWebQuery('搜一下「量子计算 最新进展」')).toBe('量子计算 最新进展')
  })
  it('去尾部标点', () => {
    expect(extractWebQuery('搜索 node-llama-cpp 是什么？')).toBe('node-llama-cpp 是什么')
  })
  it('没有指令词时原样返回', () => {
    expect(extractWebQuery('量子计算')).toBe('量子计算')
  })
})

/* ★ 2026-11-08 用户截图：上文聊「根据黄历」，用户接「你联网查询」——
 *  旧版剥不掉人称主语，拿着「你联网查询」整句去 Bing 搜，返回「你」字的词典解释。
 *  修复：人称/指令词剥净 → 剥空返回 ''，由调用方结合最近对话补全搜索词。 */
describe('extractWebQuery — 人称主语剥净（2026-11-08）', () => {
  it('「你联网查询」剥空返回空串（不瞎搜）', () => {
    expect(extractWebQuery('你联网查询')).toBe('')
  })

  it('「你联网查询黄历」→ 黄历', () => {
    expect(extractWebQuery('你联网查询黄历')).toBe('黄历')
  })

  it('「你帮我联网查一下黄历」→ 黄历', () => {
    expect(extractWebQuery('你帮我联网查一下黄历')).toBe('黄历')
  })

  it('「麻烦你搜一下成都天气」→ 成都天气', () => {
    expect(extractWebQuery('麻烦你搜一下成都天气')).toBe('成都天气')
  })

  it('「我查查最新新闻」→ 最新新闻', () => {
    expect(extractWebQuery('我查查最新新闻')).toBe('最新新闻')
  })
})

describe('isUsableQuery / stripQueryPhrases / fallbackQueryFromTurns', () => {
  it('isUsableQuery：纯代词/语气词无效，实义词有效', () => {
    expect(isUsableQuery('你')).toBe(false)
    expect(isUsableQuery('什么')).toBe(false)
    expect(isUsableQuery('吗')).toBe(false)
    expect(isUsableQuery('黄历')).toBe(true)
    expect(isUsableQuery('成都天气')).toBe(true)
  })

  it('stripQueryPhrases：多层指令词循环剥净', () => {
    expect(stripQueryPhrases('你联网查询')).toBe('')
    expect(stripQueryPhrases('帮我搜索一下')).toBe('')
    expect(stripQueryPhrases('你帮我查一下黄历')).toBe('黄历')
  })

  it('fallbackQueryFromTurns：跳过纯指令句，取上文主题', () => {
    const turns = [
      { role: 'user', text: '今天适合理发么' },
      { role: 'assistant', text: '理发没有天气上的硬性禁忌…' },
      { role: 'user', text: '根据黄历' },
      { role: 'assistant', text: '我无法直接查询黄历信息…' },
      { role: 'user', text: '你联网查询' },
    ]
    expect(fallbackQueryFromTurns(turns)).toBe('黄历')
  })

  it('fallbackQueryFromTurns：没有可用主题返回空串', () => {
    expect(fallbackQueryFromTurns([
      { role: 'user', text: '你联网查询' },
      { role: 'assistant', text: '好的' },
    ])).toBe('')
  })
})
