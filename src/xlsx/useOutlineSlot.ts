/**
 * 章节编排插槽 —— WorkbookEditor 与章节能力之间的**唯一细线**。
 *
 * 为什么要单独一个文件（而不是把代码塞进 WorkbookEditor）：
 *   WorkbookEditor 有 4266 行，是 OO 内核的主战场。
 *   章节能力是「外挂在全量功能之上」的差异化层，若与内核同居一个文件，
 *   每次 oo 系列升级都要在同4266 行里做冲突消解 —— 正是用户要避免的。
 *
 *   本文件把「取行 → 给窗体 → 收变更 → 写回 → 落盘」这条链收在一处，
 *   WorkbookEditor 里只剩三行调用（见下方 useOutlineSlot 的用法）。
 *   将来章节交互大改，改这个文件即可，不动内核。
 *
 * 纯计算层铁律：本文件是 UI 接线层（会用 React），算法仍在 outline.ts / outlineAdapter.ts。
 */

import { useCallback, useEffect, useMemo, useState } from 'react'
import { getSheet } from '@office-kit/xlsx/workbook'
import type { Workbook } from '@office-kit/xlsx/workbook'
import { readOutline, commitToSheet, type OutlineCols } from './outlineAdapter'
import type { OutlineRow } from './outline'

/** 插槽配置。传true 等同于{} （全默认）。 */
export interface OutlineSlotConfig {
  /** 列布局，默认 A=level / B=no / C=title */
  cols?: OutlineCols
  /** 只读：窗体内所有结构操作禁用 */
  readOnly?: boolean
}

/** 插槽对外暴露的句柄（WorkbookEditor 内部只用到这四样） */
export interface OutlineSlot {
  /** 是否启用（false → 内核一行章节代码都不跑） */
  enabled: boolean
  /** 章节编排窗体开合 */
  open: boolean
  setOpen: (v: boolean) => void
  /** 当前工作表的章节行（未启用时为空数组） */
  rows: OutlineRow[]
  /**
   * 结构变更收口：钳制 → 重算编号 → 写回 A/B/C 列 → 落盘。
   * @param msg 给用户的提示（跳级修正等），空串 = 不提示
   */
  commit: (msg?: string) => void
  /** 列布局（给窗体做只读判定与说明展示） */
  cols: OutlineCols
  /** 只读态 */
  readOnly: boolean
}

/**
 * useOutlineSlot —— 在 WorkbookEditor 内挂载章节能力。
 *
 * @param wbRef当前工作簿的可变引用（内核的 state.wb 每次 commit 都是新对象，故用 ref 兜住最新值）
 * @param sheetName 当前工作表名
 * @param commit 内核的落盘函数（重算 + 补丁保存 + undo 快照 + 引用对账，一套口径不另开路径）
 * @param cfg    插槽配置（undefined → 完全不启用）
 * @param onMsg  顶部提示条（复用内核的 err 通道，不另造提示系统）
 */
export function useOutlineSlot(
  wb: Workbook | null,
  sheetName: string,
  commit: () => Promise<void> | void,
  cfg: boolean | OutlineSlotConfig | undefined,
  onMsg: (msg: string) => void,
): OutlineSlot {
  const enabled = cfg !== undefined && cfg !== false
  const opts: OutlineSlotConfig = typeof cfg === 'object' && cfg !== null ? cfg : {}
  const cols = opts.cols
  const readOnly = opts.readOnly === true

  const [open, setOpen] = useState(false)
  /**
   * 行的唯一真相在**工作表**里（level 存A 列，方案 §6），
   * 内存里这份只是给窗体看的派生副本 —— 所以 rows 由 wb + sheetName 派生，不另存state。
   * 这样用户在网格里直接改 A 列，窗体下次打开立刻反映，不存在两份状态互同步的问题。
   *
   * wb 每次 commit 都是新对象 → 依赖它天然触发重算，无需手动 dirty。
   */
  const rows = useMemo<OutlineRow[]>(() => {
    if (!enabled) return []
    if (!wb || !sheetName) return []
    try {
      const ws = getSheet(wb, sheetName)
      const r = readOutline(ws, cols)
      // 读的时候钳制过跳级 —— 如实回报，不静默（用户需要知道自己的表被改过层级）
      if (r.clamped) onMsg('已自动修正跳级标题（编号以 A 列标记为准）')
      return r.rows
    } catch (e) {
      onMsg('章节解析失败：' + (e instanceof Error ? e.message : String(e)))
      return []
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [enabled, wb, sheetName, cols])

  /**
   * 结构变更落盘。
   *
   * 顺序不可换：先钳制+重算（commitToSheet 内做）→ 再 commit 走内核统一落盘。
   * 不另开写盘路径 —— 补丁保存 / undo 快照 / 引用对账都在内核 commit 里，绕开就会丢。
   */
  const commitOutline = useCallback(
    (msg?: string) => {
      if (!wb || !sheetName) return
      try {
        const ws = getSheet(wb, sheetName)
        const res = commitToSheet(ws, rows, cols)
        if (res.clamped || msg) onMsg(msg || '已自动修正跳级标题')
        void Promise.resolve(commit()).then(() => setOpen(false))
      } catch (e) {
        onMsg('章节写回失败：' + (e instanceof Error ? e.message : String(e)))
      }
    },
    [rows, wb, sheetName, cols, commit, onMsg],
  )

  // 切工作表时收起窗体，避免拿着上一张表的行去改当前表
  useEffect(() => {
    setOpen(false)
  }, [sheetName])

  return { enabled, open, setOpen, rows, commit: commitOutline, cols: cols ?? {}, readOnly }
}
