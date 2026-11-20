import { PDFDocument } from 'pdf-lib'

/**
 * PDF 合并（2026-10-22 用户指令：「多个文档转pdf并连接。提供排序，连接pdf，保存文件」）
 *
 * 纯计算层：Uint8Array[] 进 → 合并后的 Uint8Array 出。零 IO / 零 electron。
 * 排序由调用方（UI）决定——这里只按传入顺序拼接。
 */

export interface MergePdfResult {
  ok: boolean
  bytes?: Uint8Array
  /** 合了几份 */
  count: number
  /** 每份的页数（诊断/显示用） */
  pages: number[]
  why?: string
}

/** 按顺序拼接多份 PDF 字节为一个 PDF */
export async function mergePdfs (list: Uint8Array[]): Promise<MergePdfResult> {
  const valid = list.filter((b) => b && b.length > 4)
  if (!valid.length) return { ok: false, count: 0, pages: [], why: '没有可合并的 PDF（上游转换失败或为空）' }
  try {
    const out = await PDFDocument.create()
    const pages: number[] = []
    for (const bytes of valid) {
      const src = await PDFDocument.load(bytes, { ignoreEncryption: true })
      const added = await out.copyPages(src, src.getPageIndices())
      for (const pg of added) out.addPage(pg)
      pages.push(added.length)
    }
    const merged = await out.save()
    return { ok: true, bytes: merged, count: valid.length, pages }
  } catch (e) {
    return { ok: false, count: 0, pages: [], why: '合并失败：' + String((e as Error)?.message ?? e) }
  }
}
