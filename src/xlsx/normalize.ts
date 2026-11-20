/**
 * OOXML 归一化层 —— 让「解析器」能读懂真实世界写出来的合法文件。
 *
 * 为什么必须有这一层（本机实测，2026-09-20）：
 *   @office-kit/xlsx@0.18.0 的解析器对 schema 里「可选」的构造过于严格，
 *   遇到真实 Excel/WPS 导出的文件会直接抛异常拒收：
 *
 *     ①  `<fill/>`            → OpenXmlSchemaError: fillFromTree: <fill> has no child element
 *     ②  `<c t="s">` 无 `<v>` → OpenXmlSchemaError: worksheet: <c t="s"> missing <v>
 *
 *   两份文件都是 WPS 生成的合法 OOXML，Excel 能正常打开，但 office-kit 打不开。
 *   实测那份 190KB 的气材料对照表：`<fill/>` 出现 1 次、`<c t="s">` 无 `<v>` 出现 **23291 次**。
 *
 * 策略：不做「猜」，只做「登记」。
 *   每条补丁都是窄口径、可审计的字符串替换，且只在必要时才重新打包 zip
 *   （没命中任何补丁 → 原字节原样返回，绝无副作用）。
 *
 * ⚠️ 维护提醒：这是对第三方库严格性的补偿层，不是业务逻辑。
 *    库升级后应重跑 `scripts/check-normalize.mjs` + `scripts/check-roundtrip.mjs`，
 *    若某条补丁已不再命中，就删掉它。
 */

import JSZip from 'jszip'

export interface NormalizeResult {
  /** 归一化后的字节；无补丁命中时 === 入参（同一引用） */
  bytes: Uint8Array
  /** 命中的补丁说明，供 UI 如实展示（改了什么，就要说清楚） */
  applied: string[]
}

interface Patch {
  id: string
  /** 为什么需要它 */
  why: string
  /** 只在这些部件上应用 */
  files: RegExp[]
  re: RegExp
  to: string
}

const PATCHES: Patch[] = [
  {
    id: 'empty-fill',
    why: 'schema 允许 <fill> 无子元素（表示无填充），office-kit 要求必须有 → 补成 patternFill none',
    files: [/^xl\/styles\.xml$/],
    // <fill/>  /  <fill></fill>  /  <fill ></fill>
    re: /<fill(\s*)\/>|<fill(\s*)><\/fill>/g,
    to: '<fill><patternFill patternType="none"/></fill>',
  },
  {
    id: 'cell-type-without-value',
    why: 'WPS 会写出「声明为共享字符串但没值」的空单元格，office-kit 要求有 <v> → 去掉残留的 t="s"',
    files: [/^xl\/worksheets\/sheet\d+\.xml$/],
    /*
     * 只收窄到 t="s"。
     *
     * office-kit 抛的错原文是 `worksheet: <c t="s"> missing <v>` —— 它只对共享字符串类型不满。
     * 早先版本写成 t="[^"]*" 会把 t="n" / t="b" / t="e" 也一并去掉：虽然 t="n" 本来就是默认类型、
     * 去掉语义等价，但「只动必须动的」是这层的底线 —— 多余的改写就是多余的损坏风险。
     * （这条是 normalize.test.ts 里那条断言抓出来的。）
     *
     * 只匹配「标签一闭合就结束」的单元格：
     *   <c r="A1" t="s" s="3"/>     →  <c r="A1" s="3"/>
     *   <c r="A1" t="s" s="3"></c>  →  <c r="A1" s="3"></c>
     * 有 <v> 的形如 <c r="A1" t="s"><v>5</v></c>：因 [^>]*? 不匹配 '>'，正则到不了 '></c>'，
     * 天然不会被误伤 —— 这一点由单测的「绝不误伤」用例守着。
     */
    re: /<(c\b[^>]*?)\st="s"([^>]*?)(\/>|><\/c>)/g,
    to: '<$1$2$3',
  },
]

/**
 * 归一化一份 xlsx 字节，使其能被 office-kit 解析。
 * 无补丁命中时返回原引用（字节完全不变，零副作用）。
 */
export async function normalizeWorkbookBytes(input: ArrayBuffer | Uint8Array): Promise<NormalizeResult> {
  const src = input instanceof Uint8Array ? input : new Uint8Array(input)

  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(src)
  } catch (e) {
    // 不是合法 zip（可能是 .xls 老二进制格式）——交给调用方给出人话提示
    throw new Error(
      '这个文件不是 xlsx（zip）格式，可能是旧版 .xls。请在 Excel/WPS 里「另存为 .xlsx」后再用。' +
        `（细节：${e instanceof Error ? e.message : String(e)}）`
    )
  }

  const applied: string[] = []
  let touchedAny = false

  for (const p of PATCHES) {
    let touched = 0
    for (const path of Object.keys(zip.files)) {
      if (zip.files[path].dir) continue
      if (!p.files.some((re) => re.test(path))) continue
      const f = zip.file(path)
      if (!f) continue
      const old = await f.async('string')
      // 用函数式替换，避免值里出现 $ 时被当成反向引用
      const next = old.replace(p.re, (...args: unknown[]) => {
        touched++
        return p.to.replace(/\$(\d)/g, (_, d: string) => String(args[Number(d)] ?? ''))
      })
      if (touched) zip.file(path, next)
    }
    if (touched) {
      touchedAny = true
      applied.push(`${p.id}（${touched} 处）：${p.why}`)
    }
  }

  // 关键：没命中就原样返回，一次都不重新打包 —— 避免无谓地改写用户的文件
  if (!touchedAny) return { bytes: src, applied }

  const out = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  return { bytes: out, applied }
}

/** 当前登记的补丁清单（UI / 文档展示用） */
export const NORMALIZE_PATCH_IDS = PATCHES.map((p) => p.id)
