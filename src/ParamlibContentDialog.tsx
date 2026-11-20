/**
 * 参数库内容窗体（2026-10-23 用户指令「双击参数库 打开窗体 选择 参数库内容」）：
 * 双击圆形参数库节点 → 居中大窗体（ooexcel 窗体语言）——
 *   头部：库切换下拉（PARAMLIB_ITEMS）+ 已选内容提示
 *   内容：AG Grid 全量数据（虚拟滚动，1312 行无压力），单击行 = 选中
 *   选中写 config.libSelectedRow（随节点持久）——下游连线语义后续接。
 * 取代旧的小选择卡（选库 + 浏内容合一）。
 */
import { useEffect, useMemo, useState } from 'react'
import { createPortal } from 'react-dom'
import { AgGridReact } from 'ag-grid-react'
import type { ColDef } from 'ag-grid-community'
import { useBaseStore } from './core/store'
import { pushLog } from './ui/logBus'
import { PARAMLIB_ITEMS } from './ParamlibPanel'

export function ParamlibContentDialog ({ nodeId, onClose }: { nodeId: string; onClose: () => void }) {
  const tables = useBaseStore((s) => s.tables)
  const cfg = (tables[nodeId]?.config ?? {}) as {
    libName?: string
    libSelectedRow?: { idx: number; lib: string; cells: string[] }
  }
  const [curLib, setCurLib] = useState(cfg.libName ?? PARAMLIB_ITEMS[0]?.name ?? '')
  const [selData, setSelData] = useState<Record<string, string> | null>(null)
  const st = () => useBaseStore.getState()
  /* ★ 主题跟随（同 TableForm）：根 data-theme 切换时换 quartz/quartz-dark */
  const [dark, setDark] = useState(() => document.documentElement.dataset.theme !== 'light')
  useEffect(() => {
    const ob = new MutationObserver(() => setDark(document.documentElement.dataset.theme !== 'light'))
    ob.observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] })
    return () => ob.disconnect()
  }, [])

  const entry = PARAMLIB_ITEMS.find((it) => it.name === curLib) ?? PARAMLIB_ITEMS[0]
  const rowData = useMemo(() => (entry ? entry.rows.map((r) => Object.fromEntries(entry.fields.map((f, i) => [f, r[i] ?? '']))) : []), [entry])
  const colDefs = useMemo<ColDef[]>(() => (entry ? entry.fields.map((f, i) => ({
    field: f,
    headerName: f,
    width: i === 0 ? 70 : undefined,
    flex: i === 0 ? 0 : 1,
    minWidth: 90,
    filter: true,
    sortable: true,
    resizable: true,
  })) : []), [entry])

  const pick = (idx: number | undefined, data: Record<string, string> | undefined) => {
    if (!entry || idx == null || !data) return
    const cells = entry.rows[idx] ?? []
    setSelData(data)
    st().pushHistory()
    st().setNodeConfig(nodeId, {
      libName: entry.name,
      libSelectedRow: { idx, lib: entry.name, cells },
    })
    const summary = cells.map((c, i) => (c ? `${entry.fields[i]}:${c}` : '')).filter(Boolean).slice(0, 3).join(' · ')
    pushLog(`参数库：已选「${entry.name}」第 ${idx + 1} 行（${summary}）`, 'ok')
  }

  const selSummary = cfg.libSelectedRow
    ? cfg.libSelectedRow.cells.map((c, i) => (c ? c : '')).filter(Boolean).slice(0, 3).join(' · ')
    : ''

  return createPortal(
    <div className="pm-pdfpreview nodrag nopan" onClick={onClose}>
      <div className="pn-preview-card pm-pdfpreview-card libpick-big" onClick={(e) => e.stopPropagation()}>
        <div className="pn-preview-bar">
          <span className="pn-preview-nm">📚 参数库内容</span>
          <select
            className="libpick-select"
            value={curLib}
            title="切换参数库"
            onChange={(e) => { setCurLib(e.target.value); setSelData(null) }}
          >
            {PARAMLIB_ITEMS.map((it) => <option key={it.name} value={it.name}>{it.name}</option>)}
          </select>
          <span className="libpick-sel-hint" title={selSummary}>
            {cfg.libSelectedRow && cfg.libSelectedRow.lib === curLib ? `已选：${selSummary}` : '单击行选择参数内容'}
          </span>
          <button className="pn-preview-close" title="关闭" onClick={onClose}>×</button>
        </div>
        <div className={(dark ? 'ag-theme-quartz-dark' : 'ag-theme-quartz') + ' libpick-grid'}>
          <AgGridReact
            rowData={rowData}
            columnDefs={colDefs}
            rowHeight={30}
            headerHeight={34}
            animateRows={false}
            onRowClicked={(e) => pick(e.rowIndex ?? undefined, e.data as Record<string, string> | undefined)}
            getRowStyle={(p) => ({ background: p.data != null && p.data === selData ? 'var(--c-blue2)' : 'transparent' })}
          />
        </div>
      </div>
    </div>,
    document.body,
  )
}
