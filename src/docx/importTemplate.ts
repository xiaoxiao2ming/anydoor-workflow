import { sdtifyDocx } from './sdtify'
import { getTemplateStore, type TemplateMeta } from './templateStore'

/**
 * 导入 .docx 模板（共享逻辑）：
 * 设计院给的是普通文档 → 把 ｛｛field:x｝｝ 等标记原地转成 Word 内容控件后入库；
 * 没有标记也照常入库（原样载入，可预览/导出）—— noMarks=true 让调用方自己决定怎么提示。
 *
 * 两个入口共用（不许各写一份）：WordReportPanel 的「导入模板」按钮、NodeEditor 标题栏 📂。
 */
export async function importDocxTemplate(picked: {
  name: string
  data: Uint8Array
}): Promise<{ meta: TemplateMeta; noMarks: boolean; tags: string[] }> {
  const { buffer, count, tags } = await sdtifyDocx(picked.data)
  if (!count) {
    const meta = await getTemplateStore().save(picked.name, picked.data)
    return { meta, noMarks: true, tags }
  }
  const meta = await getTemplateStore().save(picked.name, buffer)
  return { meta, noMarks: false, tags }
}
