/**
 * 工作簿内 node:// 引用扫描 —— 「引线创建/删除交互」的纯计算层（批次1-1）。
 *
 * 事实源模型：**引用格（node:// URN 文本）是引用的事实源**，画布「数据流」连线
 * （config.linkedInputs）只是它的投影。本模块给两个方向的对账提供数据：
 *   · 清引用格后  → collectUrnSources 找「已无引用格」的来源节点（linkedInputs 摘除）
 *   · 画布删线后 → scanUrnCells 找「来源已不在连线里」的孤儿引用格（编辑器内提示）
 *
 * 对账保护：画布手动拖线（App.handleConnect）会同时写 inputs 与 linkedInputs，
 * inputs 里的来源视为「手动连线」——即使没有引用格也不摘除。
 *
 * 纯计算层铁律：不碰 electron / node: / document（与 recompute.ts 同级约束）。
 */
import { getSheet, sheetNames } from '@office-kit/xlsx/workbook'
import { getCellByCoord, getValueExtent } from '@office-kit/xlsx/worksheet'
import { indexToCol } from './a1'
import { parseUrn } from '../ref/urn'
import type { Workbook } from '@office-kit/xlsx/workbook'

/** 一个 node:// 引用格的落点与指向 */
export interface UrnCellRef {
  /** 所在工作表名 */
  sheet: string
  /** A1 坐标（如 B2） */
  coord: string
  /** 来源节点 id */
  nodeId: string
  /** 来源字段 id */
  fieldId: string
}

/** 扫描全簿：所有值为 node://… 文本的格子（跳过解析失败的） */
export function scanUrnCells(wb: Workbook): UrnCellRef[] {
  const out: UrnCellRef[] = []
  for (const sn of sheetNames(wb)) {
    const ws = getSheet(wb, sn)
    if (!ws) continue
    const e = getValueExtent(ws)
    if (!e) continue
    for (let r = e.minRow; r <= e.maxRow; r++) {
      for (let c = e.minCol; c <= e.maxCol; c++) {
        const coord = indexToCol(c) + r
        const cell = getCellByCoord(ws, coord)
        if (!cell || typeof cell !== 'object') continue
        const v = (cell as { value?: unknown }).value
        if (typeof v !== 'string' || !v.startsWith('node://')) continue
        const u = parseUrn(v)
        if (u?.scheme === 'node') out.push({ sheet: sn, coord, nodeId: u.target, fieldId: u.locator })
      }
    }
  }
  return out
}

/**
 * 全簿引用到的来源节点 id 集合（含公式体里内嵌的 node:// token —— 待办「引用参与求值」
 * 落地后公式会混排引用，对账不能把它们当「无引用」误摘连线）。
 */
export function collectUrnSources(wb: Workbook): Set<string> {
  const set = new Set<string>()
  for (const sn of sheetNames(wb)) {
    const ws = getSheet(wb, sn)
    if (!ws) continue
    const e = getValueExtent(ws)
    if (!e) continue
    for (let r = e.minRow; r <= e.maxRow; r++) {
      for (let c = e.minCol; c <= e.maxCol; c++) {
        const cell = getCellByCoord(ws, indexToCol(c) + r)
        if (!cell || typeof cell !== 'object') continue
        const v = (cell as { value?: unknown }).value
        if (typeof v === 'string') {
          if (v.startsWith('node://')) {
            const u = parseUrn(v)
            if (u?.scheme === 'node') set.add(u.target)
          }
        } else if (v && typeof v === 'object' && (v as { kind?: string }).kind === 'formula') {
          // 公式体混排引用 token：node://<id>#… 的 id 一律计入（保守，宁可不摘不错杀）
          const f = (v as { formula?: string }).formula ?? ''
          for (const m of f.matchAll(/node:\/\/([^#\s(!*+\-/]+)#/g)) {
            if (m[1]) set.add(m[1])
          }
        }
      }
    }
  }
  return set
}
