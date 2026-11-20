import { useEffect, useMemo, useState } from 'react'
import { AgGridReact } from 'ag-grid-react'
import type { ColDef } from 'ag-grid-community'
import type { TableSchema } from './core/schema'
import type { Row } from './core/store'

const yuan = (v: unknown) =>
  '¥' + (Number(v) || 0).toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })

/** schema 驱动的表单：number/text 可编辑，formula/lookup 只读（由引擎算好传入） */
export function TableForm({
  schema,
  rows,
  onCellEdit,
}: {
  schema: TableSchema
  rows: Row[]
  onCellEdit: (rowIdx: number, fieldId: string, value: unknown) => void
}) {
  const colDefs = useMemo<ColDef<Row>[]>(() => {
    const cols: ColDef<Row>[] = [
      {
        headerName: '#',
        valueGetter: (p) => (p.node?.rowIndex ?? 0) + 1,
        editable: false,
        width: 52,
        pinned: 'left',
        cellClass: 'row-no',
      },
    ]
    for (const f of schema.fields) {
      const computed = f.type === 'formula' || f.type === 'lookup'
      cols.push({
        headerName: f.name,
        field: f.id,
        editable: !computed,
        width: f.type === 'text' ? 160 : 120,
        type: f.type === 'number' || computed ? 'numericColumn' : undefined,
        cellClass: computed ? 'col-computed' : undefined,
        valueFormatter: (p) => {
          if (p.value === undefined || p.value === null || p.value === '') return ''
          if (f.format === 'currency') return yuan(p.value)
          return String(p.value)
        },
      })
    }
    return cols
  }, [schema])

  const defaultColDef = useMemo<ColDef>(() => ({ sortable: true, filter: true, resizable: true }), [])


  /* ★ 主题跟随（2026-10-20）：根 data-theme 切换时表格换 quartz/quartz-dark */
  const [dark, setDark] = useState(() => document.documentElement.dataset.theme !== 'light')
  useEffect(() => {
    const h = (e: Event) => setDark((e as CustomEvent).detail !== 'light')
    window.addEventListener('app-theme-change', h)
    return () => window.removeEventListener('app-theme-change', h)
  }, [])

  return (
    <div className={dark ? "ag-theme-quartz-dark" : "ag-theme-quartz"} style={{ flex: 1, minHeight: 0 }}>
      <AgGridReact<Row>
        rowData={rows}
        columnDefs={colDefs}
        defaultColDef={defaultColDef}
        stopEditingWhenCellsLoseFocus
        rowHeight={30}
        headerHeight={34}
        onCellValueChanged={(e) => {
          if (e.colDef.field && e.rowIndex != null && e.newValue !== e.oldValue) {
            const fid = e.colDef.field
            const f = schema.fields.find((x) => x.id === fid)
            const v = f?.type === 'number' ? Number(e.newValue) : e.newValue
            onCellEdit(e.rowIndex, fid, v)
          }
        }}
      />
    </div>
  )
}
