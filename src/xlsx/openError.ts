/**
 * 打开工作簿失败时，把第三方库的异常翻译成「人话 + 该怎么办」。
 *
 * 为什么需要这一层：
 *   office-kit 抛出的原始异常长这样 ——
 *     OpenXmlSchemaError: fillFromTree: <fill> has no child element
 *     OpenXmlSchemaError: worksheet: <c t="s"> missing <v>
 *   这两句对用户毫无意义，而且会让人以为「文件坏了」。
 *   实际上文件完全合法，是解析器太严格（normalize.ts 会补掉一部分）。
 *
 *   「打不开」这件事本身不可避免（客户的文件千奇百怪），
 *   但至少要做到：① 不甩堆栈 ② 说清是哪一类问题 ③ 给出可执行的下一步。
 *
 * 纯计算层：只做分类，不做 IO。失败样本的落盘由平台适配层负责
 * （见 templateStore.ts 的 reportOpenFailure）。
 */

export type OpenFailureKind =
  /** 老版 .xls（OLE 复合二进制），不是 zip */
  | 'legacy-xls'
  /** 加密：要么整本加密码，要么工作表/工作簿结构被保护 */
  | 'encrypted'
  /** 不是 zip，也不像 OLE —— 多半是别的东西改了后缀，或文件被截断 */
  | 'not-zip'
  /** 解析器的严格性：合法 OOXML 但 office-kit 拒收（归一化层尽力了但没覆盖） */
  | 'schema-strict'
  /** zip 结构损坏：中央目录不完整、CRC 对不上、部件缺失 */
  | 'corrupt'
  | 'unknown'

export interface OpenFailure {
  kind: OpenFailureKind
  /** 一句话说清是什么问题（给用户看） */
  title: string
  /** 该怎么办（给用户看，要可执行） */
  advice: string
  /** 原始异常文本，保留给排查，不直接怼到界面上 */
  detail: string
  /** 打开前已应用的归一化补丁（排查「是不是补丁没覆盖」的关键信息） */
  applied: string[]
}

/** 打开失败时抛出；带结构化信息，UI 只需读 .failure */
export class WorkbookOpenError extends Error {
  readonly failure: OpenFailure
  constructor(failure: OpenFailure) {
    super(failure.title)
    this.name = 'WorkbookOpenError'
    this.failure = failure
  }
}

/* ────────────────── 容器嗅探 ────────────────── */

const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]
/** 空 zip（只有中央目录，没有条目）—— 能过 zip 解析但必然没有 workbook.xml */
const ZIP_EMPTY = [0x50, 0x4b, 0x05, 0x06]

export type ContainerKind = 'zip' | 'ole' | 'ole-encrypted' | 'empty-zip' | 'unknown'

/**
 * 看文件头判断容器类型。
 *
 * 加密的 OOXML（ECMA-376 Agile Encryption）**也是** OLE 复合文档 ——
 * 所以不能只看魔数，还要找 `EncryptionInfo` 流名。
 * 该流名以 UTF-16LE 存在 OLE 目录项里，直接用字节搜索即可。
 */
export function sniffContainer(bytes: Uint8Array): ContainerKind {
  if (bytes.length >= 8 && OLE_MAGIC.every((b, i) => bytes[i] === b)) {
    return hasUtf16(bytes, 'EncryptionInfo') ? 'ole-encrypted' : 'ole'
  }
  if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b) {
    if (ZIP_EMPTY.every((b, i) => bytes[i] === b)) return 'empty-zip'
    return 'zip'
  }
  return 'unknown'
}

/** 在字节流里找某个字符串的 UTF-16LE 编码（OLE 目录项用 UTF-16LE 存流名） */
function hasUtf16(bytes: Uint8Array, text: string): boolean {
  const pat: number[] = []
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i)
    pat.push(c & 0xff, (c >> 8) & 0xff)
  }
  const n = bytes.length - pat.length
  outer: for (let i = 0; i <= n; i++) {
    for (let j = 0; j < pat.length; j++) if (bytes[i + j] !== pat[j]) continue outer
    return true
  }
  return false
}

/* ────────────────── 分类 ────────────────── */

const RX_SCHEMA = /OpenXmlSchemaError|fillFromTree|has no child element|missing <v>|unexpected element|invalid child/i
const RX_ZIP = /end of central directory|corrupted zip|invalid zip|end of data|CRC32|unexpected end|central directory/i
const RX_MISSING_PART = /workbook\.xml|xl\/worksheets\/|Cannot read properties of undefined/i
const RX_ENCRYPT = /password|encrypt|密码|受保护|protected/i

/**
 * 把任意异常分类成 OpenFailure。
 *
 * @param err     捕获到的异常
 * @param bytes   原始字节（若已读到；用于容器嗅探）
 * @param applied 归一化层已命中的补丁 id
 */
export function classifyOpenFailure(
  err: unknown,
  bytes: Uint8Array | null,
  applied: string[] = [],
): OpenFailure {
  const detail = err instanceof Error ? `${err.name}: ${err.message}` : String(err)

  // ① 先看容器 —— 这比看错误文本可靠得多
  if (bytes && bytes.length) {
    const c = sniffContainer(bytes)
    if (c === 'ole-encrypted') {
      return {
        kind: 'encrypted',
        title: '这个文件有密码保护，打不开。',
        advice: '在 Excel/WPS 里打开它、输入密码，然后「另存为」一份不带密码的 .xlsx 再导入。',
        detail,
        applied,
      }
    }
    if (c === 'ole') {
      return {
        kind: 'legacy-xls',
        title: '这是旧版 .xls 格式（Excel 97-2003），不是 .xlsx。',
        advice: '在 Excel/WPS 里打开它，「另存为」时格式选「Excel 工作簿 (*.xlsx)」，再导入新的那份。',
        detail,
        applied,
      }
    }
    if (c === 'empty-zip') {
      return {
        kind: 'corrupt',
        title: '这个 .xlsx 是个空壳（zip 里没有任何内容）。',
        advice: '多半是导出/下载时被中断了。回源头重新导出一份，并确认文件大小不是 0。',
        detail,
        applied,
      }
    }
    if (c === 'unknown') {
      return {
        kind: 'not-zip',
        title: '这个文件不是 .xlsx 格式。',
        advice: '确认后缀是 .xlsx（不是 .xls / .csv / .et）。若是从邮件或网盘拿的，先确认下载完整。',
        detail,
        applied,
      }
    }
  }

  // ② 再看错误文本
  if (RX_ENCRYPT.test(detail)) {
    return {
      kind: 'encrypted',
      title: '这个文件被加密或保护，打不开。',
      advice: '去掉密码或撤销工作表保护后，另存为一份新的 .xlsx 再导入。',
      detail,
      applied,
    }
  }
  if (RX_SCHEMA.test(detail) || RX_MISSING_PART.test(detail)) {
    const patched = applied.length ? `已尝试 ${applied.length} 条归一化补丁，仍未覆盖到。` : '归一化层没有命中任何已知问题。'
    return {
      kind: 'schema-strict',
      title: '文件本身没问题，是我们的解析器读不了它。',
      advice: `已自动记录此样本，供后续扩充解析兼容性。临时办法：在 Excel/WPS 里「另存为」一份新的 .xlsx 通常就能打开。${patched}`,
      detail,
      applied,
    }
  }
  if (RX_ZIP.test(detail)) {
    return {
      kind: 'corrupt',
      title: '文件内容不完整，可能下载或复制时损坏了。',
      advice: '试试重新拿一份；或在 Excel/WPS 里打开后另存为新文件，让它重建内部结构。',
      detail,
      applied,
    }
  }

  return {
    kind: 'unknown',
    title: '打不开这个文件，原因未知。',
    advice: '已自动记录此样本。可以先用 Excel/WPS 另存为一份新的 .xlsx 试一次。',
    detail,
    applied,
  }
}

/**
 * 把任意异常收敛成 OpenFailure。
 * WorkbookOpenError 直接取它的 failure；其它异常（比如存储层读文件失败）
 * 走一次无字节分类，至少保证 UI 上永远拿到 title/advice/detail 三件套。
 */
export function asOpenFailure(err: unknown): OpenFailure {
  if (err instanceof WorkbookOpenError) return err.failure
  return classifyOpenFailure(err, null, [])
}

/** 失败样本：落盘后可用于「哪个客户文件又打不开了」的复盘 */
export interface OpenFailureSample {
  at: number
  name: string
  size: number
  sha256: string
  kind: OpenFailureKind
  title: string
  detail: string
  applied: string[]
}

/** 给人看的一行摘要（日志 / 诊断面板 / 剪贴板） */
export function formatFailureForReport(s: OpenFailureSample): string {
  return [
    `时间：${new Date(s.at).toLocaleString('zh-CN')}`,
    `文件：${s.name}（${(s.size / 1024).toFixed(1)} KB）`,
    `sha256：${s.sha256}`,
    `分类：${s.kind}`,
    `现象：${s.title}`,
    `归一化：${s.applied.length ? s.applied.join('；') : '（无命中）'}`,
    `原始错误：${s.detail}`,
  ].join('\n')
}
