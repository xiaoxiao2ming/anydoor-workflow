import { useMemo, useState } from 'react'
import { createPortal } from 'react-dom'
import { computeNode } from './core/engine'
import { TableForm } from './TableForm'
import { IcX } from './ui/icons'
import { badgeClassOf, styleOf } from './core/nodeStyle'
import type { TableSchema, FieldDef, FieldType } from './core/schema'
import type { Row } from './core/store'

/**
 * ★ 弹窗形态节点编辑器（2026-10-19 用户指令）：
 *   「很多节点的属性在右侧（oo系列 excel系列暂时不动）的 现在重新设计。
 *     同时增加一种 弹窗的类似于oo系列的形态。」
 *   —— 常规数据节点（表/参数表/数据表/对照/排序/文本）双击打开**居中弹窗**编辑
 *   （模仿 OO 系列的弹窗交互），替代"属性都在右侧面板"的老路径。
 *   OO 系列（ooword/oosheet/chapter）与 Excel 系列（xlsx）暂不动，仍走原入口。
 */

export const POPUP_FORM_KINDS = ['table', 'param', 'test_table', 'map', 'sort', 'text']

export function hasPopupForm(kind: string | undefined): boolean {
  return !!kind && POPUP_FORM_KINDS.includes(kind)
}

export function PopupNodeEditor({
  nodeId,
  tables,
  records,
  onClose,
  onRename,
  onAddField,
  onUpdateField,
  onRemoveField,
  onAddRow,
  onSetCell,
  onRemoveRow,
  onPushHistory,
}: {
  nodeId: string
  tables: Record<string, TableSchema | undefined>
  records: Record<string, Row[]>
  onClose: () => void
  onRename: (id: string, name: string) => void
  onAddField: (id: string, field: FieldDef) => void
  onUpdateField: (id: string, fieldId: string, patch: Partial<FieldDef>) => void
  onRemoveField: (id: string, fieldId: string) => void
  onAddRow: (id: string) => void
  onSetCell: (id: string, rowIdx: number, fieldId: string, v: unknown) => void
  onRemoveRow: (id: string, rowIdx: number) => void
  onPushHistory: () => void
}) {
  const t = tables[nodeId]
  const [draftName, setDraftName] = useState(t?.name ?? '')
  const [fieldEdit, setFieldEdit] = useState(false)
  const computed = useMemo(
    () => (t ? computeNode(nodeId, tables as Record<string, TableSchema>, records) : { fields: [], rows: [] }),
    [nodeId, tables, records, t],
  )
  if (!t) return null
  const viewSchema: TableSchema = { ...t, fields: computed.fields }
  const commitName = () => {
    const n = draftName.trim()
    if (n && n !== t.name) onRename(nodeId, n)
    else setDraftName(t.name)
  }
  const kindsLabel = styleOf(t.kind).label

  return createPortal(
    <div className="node-editor-mask" onClick={onClose}>
      <div
        className="node-editor popup-node-editor"
        role="dialog"
        aria-label={`${kindsLabel} 编辑器`}
        onClick={(e) => e.stopPropagation()}
      >
        <header className="ne-head">
          <span className={`rf-badge ${badgeClassOf(t.kind)}`} style={{ flexShrink: 0 }}>{kindsLabel}</span>
          <input
            className="app-docname"
            value={draftName}
            title="节点名 · 点击重命名"
            onChange={(e) => setDraftName(e.target.value)}
            onBlur={commitName}
            onKeyDown={(e) => { if (e.key === 'Enter') commitName() }}
          />
          <div className="ne-fileops">
            <button
              className={`bar-btn${fieldEdit ? ' primary' : ''}`}
              title="切换字段编辑：显示字段列表（参数名 / 数据格式，可增删改）"
              onClick={() => setFieldEdit((v) => !v)}
            >
              {fieldEdit ? '完成字段编辑' : '编辑字段'}
            </button>
            {fieldEdit && (
              <button className="bar-btn" onClick={() => { onPushHistory(); onAddField(nodeId, { id: '字段' + (t.fields.length + 1) + '_' + Date.now(), name: '字段' + (t.fields.length + 1), type: 'text' }) }}>
                ＋ 添加字段
              </button>
            )}
            {!fieldEdit && (
              <button className="bar-btn" onClick={() => { onPushHistory(); onAddRow(nodeId) }}>
                ＋ 添加行
              </button>
            )}
          </div>
          <button className="ne-close" title="关闭" onClick={onClose}>
            <IcX size={15} />
          </button>
        </header>
        <div className="ne-body ne-body--flush" style={{ display: 'flex', flexDirection: 'column' }}>
          {fieldEdit ? (
            <div style={{ flex: 1, overflow: 'auto', padding: 12, display: 'flex', flexDirection: 'column', gap: 6 }}>
              {t.fields.map((f) => (
                <div key={f.id} style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                  <input
                    className="popup-field-input"
                    value={f.name}
                    onChange={(e) => { onPushHistory(); onUpdateField(nodeId, f.id, { name: e.target.value }) }}
                  />
                  <select
                    className="popup-field-input"
                    style={{ flex: '0 0 110px' }}
                    value={f.type === 'number' ? 'number' : 'text'}
                    onChange={(e) => { onPushHistory(); onUpdateField(nodeId, f.id, { type: e.target.value as FieldType }) }}
                  >
                    <option value="text">文本</option>
                    <option value="number">数字</option>
                  </select>
                  <button
                    className="bar-btn"
                    style={{ flexShrink: 0 }}
                    title="删除该字段"
                    onClick={() => { onPushHistory(); onRemoveField(nodeId, f.id) }}
                  >
                    ✕
                  </button>
                </div>
              ))}
              {t.fields.length === 0 && (
                <div style={{ color: 'var(--ed-text3, var(--c-text3))', fontSize: 12 }}>还没有字段 —— 点上方「＋ 添加字段」</div>
              )}
            </div>
          ) : (
            <TableForm
              schema={viewSchema}
              rows={computed.rows}
              onCellEdit={(idx, fid, v) => { onPushHistory(); onSetCell(nodeId, idx, fid, v) }}
            />
          )}
        </div>
      </div>
    </div>,
    document.body,
  )
}
