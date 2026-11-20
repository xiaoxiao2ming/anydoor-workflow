/**
 * 参数库侧板（2026-10-23 用户指令「增加一个展开参数库的侧边 和展开节点库并列；
 * 第一个参数节点 燃气材料节点库」）。
 *
 * 列出内置参数库条目——点击「添加」→ 画布新建 param 节点并灌入内置种子数据
 * （燃气材料 = 燃气主要材料设备对照表「规格表」1312 条 × 10 列，见 core/paramSeeds.ts）。
 * 后续条目只需在 PARAMLIB_ITEMS 加一行（数据源可为内置种子或未来远程源）。
 */
import { useState } from 'react'
import { defineField, type FieldDef } from './core/schema'
import type { Row } from './core/store'
import { GAS_MATERIAL_FIELDS, GAS_MATERIAL_ROWS } from './core/paramSeeds'

export interface ParamlibEntry {
  /** 节点名（uniqueName 去重前的基础名） */
  name: string
  desc: string
  /** 字段列（text；列头来自数据源） */
  fields: string[]
  /** 数据行（与 fields 等宽） */
  rows: string[][]
}

/** 内置参数库清单：第一个 = 燃气参数库（规格表全量）。
 *  ★ 2026-10-23 用户指令「燃气材料节点库 换成 燃气参数库 而且能拖动进画布的圆形节点」：
 *  条目可拖拽 → 画布建圆形参数库（paramlib）节点；「＋ 添加」仍为含全量数据的数据参数表。 */
export const PARAMLIB_ITEMS: ParamlibEntry[] = [
  {
    name: '燃气参数库',
    desc: `规格表 · ${GAS_MATERIAL_ROWS.length} 条 × ${GAS_MATERIAL_FIELDS.length} 列 · 可拖出圆形节点`,
    fields: GAS_MATERIAL_FIELDS,
    rows: GAS_MATERIAL_ROWS,
  },
]

export function ParamlibPanel ({
  onAdd,
  onClose,
}: {
  /** App 侧：建 param 节点 + 灌数据（uniqueName/选中等由宿主处理） */
  onAdd: (name: string, fields: FieldDef[], rows: Row[]) => void
  onClose: () => void
}) {
  const [busyName, setBusyName] = useState<string | null>(null)
  const add = (it: ParamlibEntry) => {
    if (busyName) return
    setBusyName(it.name)
    // 让 loading 有一帧渲染窗口（1312 行写入 ~几十 ms）
    setTimeout(() => {
      try {
        const fields = it.fields.map((f) => defineField(f, 'text'))
        const rows: Row[] = it.rows.map((r) => {
          const o: Row = {}
          it.fields.forEach((f, i) => { o[f] = r[i] ?? '' })
          return o
        })
        onAdd(it.name, fields, rows)
      } finally {
        setBusyName(null)
      }
    }, 30)
  }
  return (
    <aside className="paramlib-panel" style={{ position: 'relative' }}>
      <button className="side-strip-btn paramlib-fold-inline" onClick={onClose} title="折叠参数库">
        <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true">
          <path d="M15 6l-6 6 6 6" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
        </svg>
      </button>
      <div className="sidebar-section">
        <p className="sidebar-label">参数库</p>
        {PARAMLIB_ITEMS.map((it) => (
          <div
            key={it.name}
            className="paramlib-item"
            title={it.name + '：拖到画布 = 圆形参数库节点；＋ 添加 = 数据参数表（全量数据）'}
            draggable
            onDragStart={(e) => {
              /* 与调色板同格式：画布 onDrop 按 'application/reactflow' 建 paramlib 圆节点；
                 额外带库名供建后 rename（否则落在默认名「参数库」） */
              e.dataTransfer.setData('application/reactflow', 'paramlib')
              e.dataTransfer.setData('application/x-paramlib-name', it.name)
              e.dataTransfer.effectAllowed = 'move'
            }}
          >
            <div className="paramlib-item-main">
              <span className="paramlib-item-name">{it.name}</span>
              <span className="paramlib-item-desc">{it.desc}</span>
            </div>
            <button
              className="paramlib-add"
              disabled={busyName === it.name}
              title={'添加「' + it.name + '」为数据参数表节点（含全量数据）'}
              onClick={() => add(it)}
            >
              {busyName === it.name ? '导入中…' : '＋ 数据表'}
            </button>
          </div>
        ))}
        <p className="paramlib-note">参数节点可连入计算/AI 节点作为数据源</p>
      </div>
    </aside>
  )
}
