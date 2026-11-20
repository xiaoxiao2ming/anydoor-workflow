import JSZip from 'jszip'
import { workbookToBytes } from '@office-kit/xlsx/io'
import type { Workbook } from '@office-kit/xlsx/workbook'

/**
 * 补丁式保存 —— 把「补丁不重建」从口号落成代码
 *
 * 为什么必须有这一层：
 *   `workbookToBytes` 的 SaveOptions **只有 mtime 与 compressionLevel**，
 *   没有任何「保留未建模部件」的开关（已核对 save.d.ts）。它只写回自己建模的部件，
 *   于是含图表 / 绘图 / 透视表的真实文件，往返后会**静默丢失**这些部件。
 *
 *   本机实测（李渡可研估算表，47 表 / 375,877 非空 / 424,981 格）：
 *     丢失 xl/charts/_rels/chart2.xml.rels、xl/drawings/_rels/drawing7.xml.rels、xl/drawings/drawing7.xml
 *     —— 图表 ×1、绘图 ×2，且**不报错**。用户改一个数字、存回去发现图表没了，是交付事故。
 *
 * 做法：
 *   原文件是唯一真相。office-kit 的产物只用来**覆盖它改动过的部件**，
 *   产物里没有的部件（图表 / 绘图 / 透视 / 媒体 / 主题 / 关系…）一律从原文件原样取回。
 *
 *   ❌ 重建：load → 自有模型 → 丢弃未建模内容 → 写新文件
 *   ✅ 补丁：open → 定位目标 → 只替换改动项 → 其余部件原样保留
 *
 * 纯计算层：不碰 electron / node: / document，浏览器与 Node 同构。
 */

export interface PatchSaveResult {
  bytes: Uint8Array
  /** 从原文件补回的部件名（office-kit 产物里缺失的） */
  restored: string[]
}

export async function patchSaveWorkbook(
  original: Uint8Array,
  wb: Workbook,
  opts?: { mtime?: Date }
): Promise<PatchSaveResult> {
  const out = await workbookToBytes(wb, opts)
  const srcZip = await JSZip.loadAsync(original)
  const outZip = await JSZip.loadAsync(out)

  const outNames = new Set(Object.keys(outZip.files).filter((n) => !outZip.files[n].dir))

  const restored: string[] = []
  for (const name of Object.keys(srcZip.files)) {
    if (srcZip.files[name].dir) continue
    // 产物里已有 → 用产物的，那才是改动后的版本
    if (outNames.has(name)) continue
    restored.push(name)
    outZip.file(name, await srcZip.files[name].async('uint8array'), { date: opts?.mtime })
  }

  const bytes = await outZip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
  })
  return { bytes, restored }
}
