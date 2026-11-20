import { describe, it, expect } from 'vitest'
import {
  classifyOpenFailure,
  formatFailureForReport,
  sniffContainer,
  WorkbookOpenError,
  type OpenFailureSample,
} from './openError'

/* ── 字节构造 ── */

const OLE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]
const ZIP = [0x50, 0x4b, 0x03, 0x04]
const ZIP_EMPTY = [0x50, 0x4b, 0x05, 0x06]

function bytes(...parts: (number[] | Uint8Array)[]): Uint8Array {
  const all = parts.flatMap((p) => [...p])
  return new Uint8Array(all)
}

/** 把字符串编成 UTF-16LE（OLE 目录项存流名的方式） */
function utf16(s: string): number[] {
  const out: number[] = []
  for (const ch of s) {
    const c = ch.charCodeAt(0)
    out.push(c & 0xff, (c >> 8) & 0xff)
  }
  return out
}

/** 一段像 OLE 复合文档的内容：魔数 + 一堆填充 + 若干 UTF-16LE 流名 */
function oleWith(streamNames: string[]): Uint8Array {
  return bytes(OLE, new Array(512).fill(0), ...streamNames.map(utf16))
}

describe('容器嗅探', () => {
  it('OLE 魔数 → ole', () => {
    expect(sniffContainer(oleWith(['Workbook']))).toBe('ole')
  })

  it('OLE + EncryptionInfo 流 → ole-encrypted（加密的 OOXML 也是 OLE 容器）', () => {
    expect(sniffContainer(oleWith(['EncryptionInfo', 'EncryptedPackage']))).toBe('ole-encrypted')
  })

  it('zip 魔数 → zip', () => {
    expect(sniffContainer(bytes(ZIP, [1, 2, 3]))).toBe('zip')
  })

  it('空 zip（只有中央目录）→ empty-zip', () => {
    expect(sniffContainer(bytes(ZIP_EMPTY, new Array(18).fill(0)))).toBe('empty-zip')
  })

  it('既不是 OLE 也不是 zip → unknown', () => {
    expect(sniffContainer(new Uint8Array([0x3c, 0x3f, 0x78, 0x6d]))).toBe('unknown')
  })

  it('字节太短不崩', () => {
    expect(sniffContainer(new Uint8Array([1]))).toBe('unknown')
    expect(sniffContainer(new Uint8Array(0))).toBe('unknown')
  })
})

describe('失败分类：以容器为准（比错误文本可靠）', () => {
  it('老 .xls → legacy-xls，并明确告诉用户「另存为 .xlsx」', () => {
    const f = classifyOpenFailure(new Error('whatever'), oleWith(['Workbook']), [])
    expect(f.kind).toBe('legacy-xls')
    expect(f.advice).toContain('.xlsx')
    expect(f.applied).toEqual([])
  })

  it('加密的 xlsx → encrypted', () => {
    const f = classifyOpenFailure(new Error('whatever'), oleWith(['EncryptionInfo']), [])
    expect(f.kind).toBe('encrypted')
    expect(f.advice).toContain('密码')
  })

  it('空壳 zip → corrupt', () => {
    const f = classifyOpenFailure(new Error('whatever'), bytes(ZIP_EMPTY, new Array(18).fill(0)), [])
    expect(f.kind).toBe('corrupt')
  })

  it('不是 zip 也不是 OLE → not-zip', () => {
    const f = classifyOpenFailure(new Error('whatever'), new Uint8Array([1, 2, 3, 4]), [])
    expect(f.kind).toBe('not-zip')
  })
})

describe('失败分类：无字节时退回错误文本', () => {
  it('OpenXmlSchemaError → schema-strict，且说明「文件没问题，是解析器读不了」', () => {
    const f = classifyOpenFailure(
      new Error('OpenXmlSchemaError: fillFromTree: <fill> has no child element'),
      null,
      ['empty-fill（1 处）'],
    )
    expect(f.kind).toBe('schema-strict')
    // 这条措辞很关键：不能让用户以为自己的文件坏了
    expect(f.title).toContain('文件本身没问题')
    expect(f.advice).toContain('另存为')
    // 已应用的补丁要带出来 —— 这是判断「补丁没覆盖」还是「另有毛病」的线索
    expect(f.applied).toEqual(['empty-fill（1 处）'])
  })

  it('missing <v> 也归 schema-strict', () => {
    const f = classifyOpenFailure(new Error('worksheet: <c t="s"> missing <v>'), null, ['a', 'b'])
    expect(f.kind).toBe('schema-strict')
    expect(f.advice).toContain('2 条归一化补丁')
  })

  it('zip 中央目录损坏 → corrupt', () => {
    const f = classifyOpenFailure(new Error("Can't find end of central directory"), null, [])
    expect(f.kind).toBe('corrupt')
  })

  it('提到 password → encrypted', () => {
    const f = classifyOpenFailure(new Error('File is password protected'), null, [])
    expect(f.kind).toBe('encrypted')
  })

  it('认不出的错误 → unknown，但仍给人话与建议', () => {
    const f = classifyOpenFailure(new Error('ECONNRESET'), null, [])
    expect(f.kind).toBe('unknown')
    expect(f.title).toBeTruthy()
    expect(f.advice).toBeTruthy()
    // 原始错误必须留底，否则没法排查
    expect(f.detail).toContain('ECONNRESET')
  })

  it('非 Error 抛出物也能处理', () => {
    const f = classifyOpenFailure('字符串异常', null, [])
    expect(f.kind).toBe('unknown')
    expect(f.detail).toBe('字符串异常')
  })
})

describe('WorkbookOpenError', () => {
  it('携带结构化 failure，message 就是给用户看的那句话', () => {
    const f = classifyOpenFailure(new Error('x'), oleWith(['Workbook']), [])
    const err = new WorkbookOpenError(f)
    expect(err.name).toBe('WorkbookOpenError')
    expect(err.failure.kind).toBe('legacy-xls')
    expect(err.message).toBe(f.title)
    expect(err).toBeInstanceOf(Error)
  })
})

describe('诊断文本', () => {
  it('把样本拼成可复制给开发的一段话', () => {
    const s: OpenFailureSample = {
      at: new Date('2026-09-20T10:00:00+08:00').getTime(),
      name: '材料表.xlsx',
      size: 2048,
      sha256: 'abc123',
      kind: 'schema-strict',
      title: '文件本身没问题，是我们的解析器读不了它。',
      detail: 'OpenXmlSchemaError: fillFromTree',
      applied: ['empty-fill（1 处）'],
    }
    const t = formatFailureForReport(s)
    expect(t).toContain('材料表.xlsx')
    expect(t).toContain('abc123')
    expect(t).toContain('schema-strict')
    expect(t).toContain('empty-fill（1 处）')
    expect(t).toContain('OpenXmlSchemaError')
  })

  it('没有命中补丁时如实写「无命中」', () => {
    const s: OpenFailureSample = {
      at: Date.now(),
      name: 'a.xlsx',
      size: 1,
      sha256: 'x',
      kind: 'unknown',
      title: 't',
      detail: 'd',
      applied: [],
    }
    expect(formatFailureForReport(s)).toContain('（无命中）')
  })
})
