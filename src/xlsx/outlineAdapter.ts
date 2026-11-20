/**
 * 章节层适配器 —— 在「工作簿」与「章节行数组」之间双向搬运数据。
 *
 * ════════════════════════════════════════════════════════════════════
 * 这个文件是章节能力与 OO 内核之间**唯一的接缝**（用户 2026-10-04 定的架构）：
 *   章节节点 = 现有 Excel 全量功能 + 外挂一层章节编排。
 *   内核（@office-kit/xlsx + WorkbookEditor）一行不改，
 *   oo 系列将来更新时，本文件吸收全部差异即可，UI 层无感。
 * ════════════════════════════════════════════════════════════════════
 *
 * 列约定（方案 §6「标记列格式」）：
 *   A 列= 标记列  存 level（**唯一真相**，1/2/3 数字，也兼容 #/##/###）
 *   B 列  = 章节号  存 no（**派生静态值**，如 1.2.1；只作展示对照，可被重算覆盖）
 *   C 列  = 标题    存 title（level>0 时有效）
 *   D 列起= 业务数据，原样保留
 *
 * 打开策略（方案 §6 明确要求）：**以 A 列为准重新解析并重算**，
 * B 列的历史脏值一律不认——否则用户手改过 B 列就会把编号带歪。
 *
 * 纯计算层铁律：不 import electron / node: / 不碰 document → 浏览器与 Node 同构，可直接 Vitest。
 *本文件只 import @office-kit/xlsx（其dist 已确认无 node: 引用）与 ./outline、./a1。
 */

import { getCellByCoord, getValueExtent, setCellByCoord } from '@office-kit/xlsx/worksheet'
import type { Worksheet } from '@office-kit/xlsx/worksheet'
import {
  MAX_LEVEL,
  commitStructure,
  computeNumbers,
  makeRow,
  parseLevel,
  renumber,
  type OutlineRow,
  type RowLevel,
} from './outline'
import { indexToCol } from './a1'

/** 列布局。做成可配置而非写死，是为「不同模板把标记列放在不同列」留口。 */
export interface OutlineCols {
  /** 标记列（level，唯一真相）。默认 'A' */
  level?: string
  /** 章节号列（no，派生）。默认 'B' */
  no?: string
  /** 标题列（title）。默认 'C' */
  title?: string
}

/** 归一化后的列布局（内部用，全部是大写单字母） */
interface ResolvedCols {
  level: string
  no: string
  title: string
}

const DEFAULT_COLS: ResolvedCols = { level: 'A', no: 'B', title: 'C' }

/** 列字母容错：空/非法 → 回落默认（不抛，脏配置不该让整张表打不开） */
function resolveCols(cols?: OutlineCols): ResolvedCols {
  const one = (v: string | undefined, fb: string) => {
    const s = String(v ?? '').trim().toUpperCase()
    return /^[A-Z]{1,3}$/.test(s) ? s : fb
  }
  return {
    level: one(cols?.level, DEFAULT_COLS.level),
    no: one(cols?.no, DEFAULT_COLS.no),
    title: one(cols?.title, DEFAULT_COLS.title),
  }
}

/**
 * 取格子的原始值 → 变成能parseLevel / 显示的标量。
 *
 * office-kit 的 cell.value 有三种形态，缺一不可：
 *   普通值  123 / 'abc'
 *   公式值  { kind:'formula', formula, cachedValue }  ← 必须取 cachedValue，
 *            否则公式格会被当成对象、parseLevel 归0、标题凭空消失
 *   空      undefined
 */
function rawValue(ws: Worksheet, coord: string): unknown {
  const cell = getCellByCoord(ws, coord)
  if (!cell || typeof cell !== 'object' || !('value' in cell)) return undefined
  const v = (cell as { value?: unknown }).value
  if (v && typeof v === 'object' && (v as { kind?: string }).kind === 'formula') {
    return (v as { cachedValue?: unknown }).cachedValue
  }
  return v
}

/** 任意值 → 显示用字符串（null/undefined → 空串） */
function str(v: unknown): string {
  if (v === undefined || v === null) return ''
  return String(v)
}

/**
 * 读工作表 → 章节行数组。
 *
 * 口径（与方案 §6 一致）：
 *   · A 列 parseLevel → level（脏输入归 0，即普通数据行，不抛）
 *   · level 有跳级时**就地钳制**并回报 clamped（调用方可toast），不静默
 *   · C 列 → title
 *   · 编号一律 renumber() 重算，**不读 B 列**（避免历史脏值污染编号）
 *   · 其余列原样进 cells（键用列字母，不含 A/B/C 三列，避免与真相字段打架）
 *
 * @returns 行数组 + 是否发生过跳级修正
 */
export function readOutline(
  ws: Worksheet | null | undefined,
  cols?: OutlineCols,
): { rows: OutlineRow[]; clamped: boolean } {
  if (!ws) return { rows: [], clamped: false }
  const c = resolveCols(cols)
  const ext = getValueExtent(ws)
  const maxRow = ext?.maxRow ?? 0
  // ★ 性能铁律：getValueExtent 是全表扫描，**必须提到行循环外**。
  //   放进循环里就是每行扫一次全表 → 1000 行实测 39ms（O(n²)），提到外面降到 3ms。
  const maxCol = ext?.maxCol ?? 0
  const rows: OutlineRow[] = []
  for (let r = 1; r <= maxRow; r++) {
    const level = parseLevel(rawValue(ws, c.level + r)) as RowLevel
    const title = str(rawValue(ws, c.title + r))
    const cells: Record<string, unknown> = {}
    // 业务列：从 D 列（索引 4）起到本行实际最后一列
    for (let ci = 4; ci <= maxCol; ci++) {
      const L = indexToCol(ci)
      const v = rawValue(ws, L + r)
      if (v !== undefined && v !== '') cells[L] = v
    }
    rows.push(makeRow({ level, title, cells }))
  }
  // 钳制跳级 + 重算编号（唯一入口，顺序由 commitStructure 保证）
  const res = commitStructure(rows)
  return { rows, clamped: res.clamped }
}

/**
 * 写回工作表 —— 把章节行数组的结构部分落盘。
 *
 * 只写三列（A level / B no / C title），**业务列一律不碰**：
 * 章节编排是结构调整，不是数据编辑，越权写业务列会吞掉用户的表格数据。
 *
 * · level>0 → A 列写数字 level、C 列写标题
 * · level=0 → A/B/C 三列清空（普通行不该留残影）
 * · no 写静态值：导出到 Excel 能直接看见编号（方案 §6 明确要求）
 *
 * ★ no 在这里**现算**（computeNumbers）而不是读 row.no：
 *   no 是派生数据，调用方手上的 row.no 可能是改层级前的旧值或空串
 *   （例如直接用 makeRowsByLevel 造的测试数据就没跑过 renumber）。
 *   写盘是最后一道关口，必须自己保证写出去的三列自洽。
 */
export function writeOutline(
  ws: Worksheet | null | undefined,
  rows: readonly OutlineRow[],
  cols?: OutlineCols,
): void {
  if (!ws) return
  const c = resolveCols(cols)
  // 纯函数算编号，不改入参（rows 是 readonly，调用方不该被写盘副作用影响）
  const nos = computeNumbers(rows)
  for (let i = 0; i < rows.length; i++) {
    const r = i + 1
    const row = rows[i]
    if (row.level > 0) {
      setCellByCoord(ws, c.level + r, row.level as never)
      setCellByCoord(ws, c.no + r, nos[i] as never)
      setCellByCoord(ws, c.title + r, (row.title || '') as never)
    } else {
      // 清空用 clearRange 同款语义：setCellByCoord(undefined) 会被静默忽略（实测），
      // 所以这里显式写空串，保证「从标题降回普通行」在 Excel 里也能看出痕迹被清掉。
      setCellByCoord(ws, c.level + r, '' as never)
      setCellByCoord(ws, c.no + r, '' as never)
      setCellByCoord(ws, c.title + r, '' as never)
    }
  }
}

/**
 * 结构变更落盘的一步到位入口 —— 树面板/ 表格侧任何结构操作都走这里。
 *
 * 顺序铁律：钳制 → 重算编号 → 写盘。三步缺一就会出现「编号与层级对不上」。
 *
 * @param ws   目标工作表
 * @param rows 行数组（**原地**被规范化，调用方不要另行持有旧引用）
 * @returns 实际写盘行数 + 是否发生过跳级修正（供 toast）
 */
export function commitToSheet(
  ws: Worksheet | null | undefined,
  rows: OutlineRow[],
  cols?: OutlineCols,
): { written: number; clamped: boolean } {
  if (!ws) return { written: 0, clamped: false }
  const clamped = commitStructure(rows).clamped
  writeOutline(ws, rows, cols)
  return { written: rows.length, clamped }
}

/**
 * 只重算编号、不写盘 —— 用于「折叠/展开」这类不碰结构的操作后保持编号一致。
 * 单独暴露是因为折叠不该触发一次全表写盘（无谓的 undo 快照 + 补丁保存）。
 */
export function renumberOnly(rows: OutlineRow[]): void {
  renumber(rows)
}

/** 该行是否可作拖动起点（level>0）—— UI 层禁用拖拽手柄时用，避免用户点了才报错 */
export function isTitleRow(row: OutlineRow | undefined): boolean {
  return !!row && row.level > 0 && row.level <= MAX_LEVEL
}

/** 行 → 标记列单元格应写的值（供表格侧就地渲染，逻辑与 writeOutline 共用一份口径） */
export function levelCellValue(row: OutlineRow): string {
  return row.level > 0 ? String(row.level) : ''
}

/** 行 → 章节号单元格应写的值 */
export function noCellValue(row: OutlineRow): string {
  return row.level > 0 ? row.no : ''
}
