import { describe, expect, it } from 'vitest'
import { webmonToCsv, type WebmonHistEntry } from './webmon'

/* ★ 2026-11-09 网址监控节点：历史导出 CSV 的转义与 BOM（Excel 中文兼容） */
describe('webmonToCsv', () => {
  it('BOM 头 + 表头 + 行序', () => {
    const hist: WebmonHistEntry[] = [
      { at: new Date('2026-11-09T10:00:00').getTime(), targetName: '首页', value: '¥129', changed: true },
    ]
    const csv = webmonToCsv(hist)
    expect(csv.startsWith('\ufeff')).toBe(true)
    expect(csv).toContain('时间,目标,值,变化,错误')
    expect(csv).toContain('"首页","¥129",是,')
  })

  it('值内引号转义为双引号（CSV 规范）', () => {
    const hist: WebmonHistEntry[] = [
      { at: Date.now(), targetName: 'T', value: '他说"你好"', changed: false },
    ]
    const csv = webmonToCsv(hist)
    expect(csv).toContain('"他说""你好"""')
  })

  it('错误条目进错误列', () => {
    const hist: WebmonHistEntry[] = [
      { at: Date.now(), targetName: 'T', value: '', changed: false, error: 'HTTP 404' },
    ]
    const csv = webmonToCsv(hist)
    expect(csv).toContain('"HTTP 404"')
  })
})
