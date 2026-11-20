import { describe, expect, it } from 'vitest'
import { matchNodeAlias } from './chatIntent'

/* ★ 2026-11-07 用户指令「要接一个 pdf 合并 而不是 PDF 节点，PDF 节点是承载 PDF 文件的」：
 *  「输出为PDF / 生成PDF」必须落到 pdfmerge（显示名 PDF生成）；pdf kind 只承接承载型话术。 */
describe('matchNodeAlias — PDF 类节点映射', () => {
  it('「输出为PDF」→ pdfmerge（PDF生成），不是 pdf 容器节点', () => {
    const a = matchNodeAlias('后面增加个节点 输出为PDF')
    expect(a?.kind).toBe('pdfmerge')
    expect(a?.name).toBe('PDF生成')
  })

  it('「生成PDF / 出个PDF / 导出PDF」→ pdfmerge', () => {
    for (const t of ['生成PDF', '出个PDF', '导出 PDF']) {
      expect(matchNodeAlias(t)?.kind, t).toBe('pdfmerge')
    }
  })

  it('「PDF合并」仍 → pdfmerge，但默认名已改「PDF生成」', () => {
    const a = matchNodeAlias('加一个 PDF 合并 节点')
    expect(a?.kind).toBe('pdfmerge')
    expect(a?.name).toBe('PDF生成')
  })

  it('「PDF文件 / 导入PDF」→ pdf 承载节点', () => {
    expect(matchNodeAlias('加个PDF文件节点')?.kind).toBe('pdf')
    expect(matchNodeAlias('导入PDF')?.kind).toBe('pdf')
  })

  it('其他类型不受影响', () => {
    expect(matchNodeAlias('输出为Word')?.kind).toBe('docx')
    expect(matchNodeAlias('PDF瘦身')?.kind).toBe('pdfslim')
    expect(matchNodeAlias('PDF转Word')?.kind).toBe('pdf2word')
  })
})
