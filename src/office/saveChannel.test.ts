import { describe, expect, it } from 'vitest'
import { decideMainChannel, digestIsEmpty, digestSame, rowsHaveContent, type SaveCandidate } from './saveChannel'

/**
 * 回归背景（2026-10-17 用户「还是不行」）：
 *   oo-events.log 18:59:52 —— 提取通道刚读到 rows=1 首行=[{"v":123}]，
 *   主通道随即返回一份 6258B、非空单元格=0 的空表，宿主把它当成功写进节点，
 *   用户重开就是空白。这组用例锁死「空产物 / 原文件产物 不得覆盖提取内容」。
 */

const bytes = (n = 10) => new Uint8Array(n).fill(7)

const cand = (over: Partial<SaveCandidate> = {}): SaveCandidate => ({
  bytes: bytes(),
  digest: { kind: 'xlsx', cells: 4, text: '表数=1 非空单元格=4' },
  sameAsSource: false,
  ...over,
})

describe('rowsHaveContent', () => {
  it('空数组 / undefined → false', () => {
    expect(rowsHaveContent([])).toBe(false)
    expect(rowsHaveContent(undefined)).toBe(false)
    expect(rowsHaveContent(null)).toBe(false)
  })

  it('全空对象（used range 铺满产物）→ false', () => {
    expect(rowsHaveContent([[{}, {}], [{}, {}]])).toBe(false)
  })

  it('有值 / 有公式 → true', () => {
    expect(rowsHaveContent([[{ v: 123 }]])).toBe(true)
    expect(rowsHaveContent([[{}, { f: 'A1*2', v: 4 }]])).toBe(true)
    expect(rowsHaveContent([['文本']])).toBe(true)
    expect(rowsHaveContent([[0]])).toBe(true) // 数字 0 也算内容
  })
})

describe('digestIsEmpty', () => {
  it('xlsx cells=0 → 空；docx chars=0 → 空', () => {
    expect(digestIsEmpty({ kind: 'xlsx', cells: 0 })).toBe(true)
    expect(digestIsEmpty({ kind: 'docx', chars: 0 })).toBe(true)
    expect(digestIsEmpty({ kind: 'xlsx', cells: 4 })).toBe(false)
  })

  it('没有摘要时不判空（不误杀）', () => {
    expect(digestIsEmpty(null)).toBe(false)
    expect(digestIsEmpty(undefined)).toBe(false)
  })
})

describe('decideMainChannel', () => {
  it('主通道没出产物 → 不用（交给提取通道）', () => {
    expect(decideMainChannel(null, [[{ v: 123 }]], true).use).toBe(false)
    expect(decideMainChannel({ bytes: new Uint8Array(0), digest: null, sameAsSource: false }, [[{ v: 1 }]], true).use).toBe(false)
  })

  /* ★★ 2026-10-29 dirty 硬门：forcesave 在本部署架构性 400（无会话），
   *   文档有未保存编辑时产物数学上不可能含编辑——符号规范化会造出「摘要恰好不同」
   *   的假阳性（21:38 docx 实测 804≠789），必须先于一切内容比对放行拒绝。 */
  it('★ dirty 硬门：有未保存编辑 → 一律拒绝（交样式化重建/沿用）', () => {
    const d = decideMainChannel(cand(), [[{ v: 123 }]], true)
    expect(d.use).toBe(false)
    expect(d.reject).toContain('未保存编辑')
    expect(d.reject).toContain('forcesave')
  })

  it('★ 空表产物 + 提取有内容 → 拒绝（18:59 真凶）', () => {
    const d = decideMainChannel(cand({ digest: { kind: 'xlsx', cells: 0, text: '表数=1 非空单元格=0（内容为空）' } }), [[{ v: 123 }]], false)
    expect(d.use).toBe(false)
    expect(d.reject).toContain('空表')
  })

  /* 规则②（字节相同+dirty → 拒绝）在硬门引入后为防御纵深：dirty=true 时硬门先拦，
     该分支经公开路径不可达，不再单列用例（保留实现兜未来入口）。 */

  it('产物与打开时字节相同 + 没改过 → 用主通道（保住原样式）', () => {
    expect(decideMainChannel(cand({ sameAsSource: true }), [[{ v: 123 }]], false).use).toBe(true)
  })

  it('提取为空时 → 用主通道（用户打开有内容的文件、未编辑即保存）', () => {
    expect(decideMainChannel(cand(), [], false).use).toBe(true)
    expect(decideMainChannel(cand(), [[{}, {}]], false).use).toBe(true)
  })

  it('产物有内容且与原文件不同 + 未编辑 → 用主通道（正常路径）', () => {
    expect(decideMainChannel(cand(), [[{ v: 123 }]], false).use).toBe(true)
  })

  /* ★★ 2026-10-17 用户「第一次可以 第二次第三次不行」★★
   *   实测日志（out2/oo-events.log 19:46:29 Scbzidm）：
   *     打开时源 = 3003B（A1=1），产物 = 6285B，digest 也「非空单元格=1 前几项=1」
   *     ⇒ x2t 把极简 xlsx 重新排版，字节变了、内容没变。
   *   旧判据只比字节（sameAsSource=no）⇒ 认为「有新内容」⇒ 采用 ⇒ 丢 A2=2。
   *   必须按内容维度判：产物内容 == 打开时内容 ⇒ 这次编辑根本没进产物。 */
  it('★ 产物内容与打开时完全相同（只是字节被 x2t 重排）→ 拒绝', () => {
    const d = decideMainChannel(
      cand({
        digest: { kind: 'xlsx', cells: 1, sample: '1', text: '表数=1 非空单元格=1 前几项=1' },
        srcDigest: { kind: 'xlsx', cells: 1, sample: '1', text: '表数=1 非空单元格=1 前几项=1' },
        sameAsSource: false,
      }),
      [[{ v: 1 }], [{ v: 2 }]],
      false,
    )
    expect(d.use).toBe(false)
    expect(d.reject).toContain('内容与打开时完全相同')
  })

  it('★ 产物格数少于提取到的内容格数 → 拒绝（产物丢格）', () => {
    const d = decideMainChannel(
      cand({
        digest: { kind: 'xlsx', cells: 1, sample: '1', text: '非空单元格=1' },
        srcDigest: { kind: 'xlsx', cells: 0, sample: '', text: '非空单元格=0' },
      }),
      [[{ v: 1 }], [{ v: 2 }]],
      false,
    )
    expect(d.use).toBe(false)
    expect(d.reject).toContain('产物丢格')
  })

  it('产物内容确实变多（真编辑进了产物）→ 用主通道', () => {
    const d = decideMainChannel(
      cand({
        digest: { kind: 'xlsx', cells: 3, sample: '1 | 2 | 3' },
        srcDigest: { kind: 'xlsx', cells: 1, sample: '1' },
      }),
      [[{ v: 1 }], [{ v: 2 }], [{ v: 3 }]],
      false,
    )
    expect(d.use).toBe(true)
  })
})

describe('digestSame', () => {
  it('xlsx 按 cells + sample 比（承载「内容一致但字节不同」）', () => {
    expect(digestSame({ kind: 'xlsx', cells: 1, sample: '1' }, { kind: 'xlsx', cells: 1, sample: '1' })).toBe(true)
    expect(digestSame({ kind: 'xlsx', cells: 1, sample: '1' }, { kind: 'xlsx', cells: 2, sample: '1 | 2' })).toBe(false)
    expect(digestSame({ kind: 'xlsx', cells: 1, sample: '1' }, { kind: 'xlsx', cells: 1, sample: '2' })).toBe(false)
  })

  it('docx 按 chars + sample 比', () => {
    expect(digestSame({ kind: 'docx', chars: 5, sample: 'hello' }, { kind: 'docx', chars: 5, sample: 'hello' })).toBe(true)
    expect(digestSame({ kind: 'docx', chars: 5, sample: 'hello' }, { kind: 'docx', chars: 6, sample: 'hello!' })).toBe(false)
  })

  it('缺摘要 / 类型不同 → 不认为是同一份（不误杀）', () => {
    expect(digestSame(null, { kind: 'xlsx', cells: 0 })).toBe(false)
    expect(digestSame({ kind: 'xlsx', cells: 0 }, null)).toBe(false)
    expect(digestSame({ kind: 'xlsx', cells: 1 }, { kind: 'docx', chars: 1 })).toBe(false)
  })
})
