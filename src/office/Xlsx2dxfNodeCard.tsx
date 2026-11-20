/* ─────────────────────────────────────────────────────────────
 * Excel转DXF 节点卡片（2026-11-10 用户指令「增加一个 excel转dxf 节点」
 *   +「外面的文件拖拽在节点上就载入显示 然后生成一次 下方就增加一个 参照 PDF生成」
 *   +「仔细看看 PDF生成节点 怎么承载文件的 怎么输出文件 显示样式的」
 *   +「样式不对 你没有按照规范」→ 按 节点规范.md 全面对齐：
 *     装载条 NodeBtnBar 唯一（第五章）· 历史 🕘钮模式禁文字标题行（7.1）·
 *     节点上拖入 dragDepth 防抖+高亮+松手加入（7.2）· 条目行 pn-row（第四章）·
 *     图标一律 icons.tsx 禁 emoji（第三章）
 *
 * 数据源三通道（优先级）：自身 records（拖放/导入落在这里）> 连线上游 > 本地导入 state
 * 产物：xlsx2dxfHistory 索引（新在前 cap 10）+ DXF 全文存 pdfBlobStore（'x2d-'+id）
 * ───────────────────────────────────────────────────────────── */
import { useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { useBaseStore } from '../core/store'
import { parseFile, exportRows } from '../core/io'
import { putBlob, getBlob } from '../core/pdfBlobStore'
import { rowsToDxf, tableToDxf, parseDxfEntities, type DxfMapSpec, type DxfOptions } from '../core/dxf'
import { NodeBtnBar, NodeSqBtn } from '../ui/NodeBtnBar'
import { IcDownload, IcGrid, IcHistory, IcMaximize, IcMinimize, IcX } from '../ui/icons'
import { pushLog } from '../ui/logBus'

type Row = Record<string, unknown>
interface Src { fields: { id: string; name: string; type?: string }[]; rows: Row[]; from: string }
interface HistEntry { id: string; name: string; at: number; count: number; skipped: number; inDb?: boolean }

const OPT_DEFAULT: DxfOptions = { drawPoint: true, drawLabel: true, drawLine: false, textHeight: 3.5, labelOffset: 1.5, flipY: false }

export function Xlsx2dxfNodeCard({ nodeId }: { nodeId: string }) {
  const cfg = useBaseStore((s) => s.tables[nodeId]?.config) as {
    xlsx2dxf?: Partial<DxfMapSpec & DxfOptions & { fileName?: string }>
    xlsx2dxfHistory?: HistEntry[]
    linkedInputs?: string[]
    inputs?: string[]
  } | undefined
  const tables = useBaseStore((s) => s.tables)
  const allRecords = useBaseStore((s) => s.records)

  const [src, setSrc] = useState<Src | null>(null)
  const [busy, setBusy] = useState(false)
  const [summary, setSummary] = useState<string>('')
  const [dragOver, setDragOver] = useState(false)
  const [histOpen, setHistOpen] = useState(false)
  const [menu, setMenu] = useState<{ x: number; y: number; entry: HistEntry } | null>(null)
  const [preview, setPreview] = useState<{ name: string; text: string } | null>(null)
  const [previewMax, setPreviewMax] = useState(false)
  const fileRef = useRef<HTMLInputElement>(null)
  const dragDepth = useRef(0)

  const hist = cfg?.xlsx2dxfHistory ?? []
  const histCount = hist.length

  const readEntryBytes = async (h: HistEntry): Promise<Uint8Array | null> =>
    getBlob('x2d-' + h.id).catch(() => null)

  /* ★ 数据源三通道（优先级）：自身 records（拖放/导入落在这里）> 连线上游 > 本地导入 state */
  const self = useMemo((): Src | null => {
    const t = tables[nodeId]
    if (!t || t.deleted || t.kind !== 'xlsx2dxf') return null
    const fields = (t.fields ?? []).filter((f) => !!f?.name && f.name !== 'text')
    const rows = allRecords[nodeId] ?? []
    if (!fields.length || !rows.length) return null
    const named = rows.map((r) => {
      const o: Row = {}
      for (const f of fields) o[f.name] = r[f.id]
      return o
    })
    return { fields: fields.map((f) => ({ id: f.name, name: f.name })), rows: named, from: t.name }
  }, [tables, allRecords, nodeId])

  const upstream = useMemo((): Src | null => {
    const ups = cfg?.linkedInputs ?? cfg?.inputs ?? []
    for (const uid of ups) {
      const t = tables[uid]
      if (!t || t.deleted) continue
      const fields = (t.fields ?? []).filter((f) => !!f?.name)
      const rows = allRecords[t.id] ?? []
      if (fields.length && rows.length) {
        const named = rows.map((r) => {
          const o: Row = {}
          for (const f of fields) o[f.name] = r[f.id]
          return o
        })
        return { fields: fields.map((f) => ({ id: f.name, name: f.name })), rows: named, from: t.name }
      }
    }
    return null
  }, [cfg, tables, allRecords])

  const active: Src | null = self ?? upstream ?? src
  const fieldNames = active?.fields.map((f) => f.name) ?? []
  const m: DxfMapSpec = {
    nameCol: cfg?.xlsx2dxf?.nameCol ?? fieldNames[0] ?? '',
    xCol: cfg?.xlsx2dxf?.xCol ?? fieldNames[1] ?? '',
    yCol: cfg?.xlsx2dxf?.yCol ?? fieldNames[2] ?? '',
    zCol: cfg?.xlsx2dxf?.zCol ?? '',
  }
  const o: DxfOptions = {
    drawPoint: cfg?.xlsx2dxf?.drawPoint ?? OPT_DEFAULT.drawPoint,
    drawLabel: cfg?.xlsx2dxf?.drawLabel ?? OPT_DEFAULT.drawLabel,
    drawLine: cfg?.xlsx2dxf?.drawLine ?? OPT_DEFAULT.drawLine,
    textHeight: cfg?.xlsx2dxf?.textHeight ?? 2,
    labelOffset: cfg?.xlsx2dxf?.labelOffset ?? 1.5,
    flipY: cfg?.xlsx2dxf?.flipY ?? false,
  }
  const patch = (p: Record<string, unknown>) => useBaseStore.getState().setNodeConfig(nodeId, p)
  const setMap = (p: Partial<DxfMapSpec>) => patch({ xlsx2dxf: { ...m, ...cfg?.xlsx2dxf, ...p } })
  const setOpt = (p: Partial<DxfOptions>) => patch({ xlsx2dxf: { ...m, ...cfg?.xlsx2dxf, ...o, ...p } })

  /** 载入数据文件（拖放到节点 / 导入钮共用）：parseFile → 自身 records（持久） */
  const loadDataFile = async (f: File): Promise<void> => {
    try {
      const r = await parseFile(f)
      useBaseStore.getState().pushHistory()
      useBaseStore.getState().setTableData(nodeId, r.fields, r.rows)
      useBaseStore.getState().renameTable(nodeId, f.name.replace(/\.(xlsx|xls|csv)$/i, ''))
      setSummary(`已载入「${f.name}」：${r.fields.length} 列 / ${r.rows.length} 行`)
    } catch (e) {
      setSummary(`载入失败：${String((e as Error)?.message ?? e).slice(0, 100)}`)
    }
  }

  const generate = async (mode: 'points' | 'table'): Promise<void> => {
    if (!active || !active.rows.length) { setSummary('没有数据——拖 Excel 到本节点、导入，或连线表格节点'); return }
    setBusy(true)
    try {
      if (mode === 'points') {
        if (!m.xCol || !m.yCol) { setSummary('先选择 X / Y 列'); return }
        const r = rowsToDxf(active.rows, m, o)
        if (!r.count) {
          const numValid = (col?: string): number => active.rows.filter((row) => Number.isFinite(Number(row[col ?? '']))).length
          setSummary(`0 个可用点——X 列「${m.xCol}」有效 ${numValid(m.xCol)}/${active.rows.length} 行，Y 列「${m.yCol}」有效 ${numValid(m.yCol)} 行。可用列：${fieldNames.join('、')}`)
          return
        }
        const id = 'h' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6)
        const name = `${active.from}-${r.count}点.dxf`
        await putBlob('x2d-' + id, new TextEncoder().encode(r.dxf))
        patch({ xlsx2dxfHistory: [{ id, name, at: Date.now(), count: r.count, skipped: r.skipped.length, inDb: true }, ...hist].slice(0, 10), xlsx2dxfResult: { at: Date.now(), count: r.count, skipped: r.skipped.length, fileName: name } })
        setHistOpen(true)
        setSummary(`✓ 已生成 ${r.count} 个点（跳过 ${r.skipped.length} 行），已加入下方历史`)
      } else {
        /* 非空边界整表：网格线 + 全部单元格文字（列宽自适应） */
        const r = tableToDxf(fieldNames, active.rows, { textHeight: o.textHeight })
        const id = 'h' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6)
        const name = `${active.from}-表格${r.usedRows}x${r.usedCols}.dxf`
        await putBlob('x2d-' + id, new TextEncoder().encode(r.dxf))
        patch({ xlsx2dxfHistory: [{ id, name, at: Date.now(), count: r.usedRows * r.usedCols, skipped: 0, inDb: true }, ...hist].slice(0, 10) })
        setHistOpen(true)
        setSummary(`✓ 已按非空边界整表生成（${r.usedRows} 行 × ${r.usedCols} 列），已加入下方历史`)
      }
    } finally { setBusy(false) }
  }

  const saveAsDownload = async (h: HistEntry): Promise<void> => {
    const bytes = await readEntryBytes(h)
    if (!bytes) { setSummary('产物已不在本机'); return }
    const blob = new Blob([bytes.slice().buffer as ArrayBuffer], { type: 'application/dxf' })
    const a = document.createElement('a')
    a.href = URL.createObjectURL(blob)
    a.download = h.name || '展点.dxf'
    document.body.appendChild(a); a.click(); a.remove()
    setTimeout(() => URL.revokeObjectURL(a.href), 5000)
  }

  const openPreview = async (h: HistEntry): Promise<void> => {
    const bytes = await readEntryBytes(h)
    if (!bytes) { setSummary('产物已不在本机：' + h.name); return }
    setPreview({ name: h.name, text: new TextDecoder().decode(bytes) })
  }

  const clearSelf = (): void => {
    useBaseStore.getState().pushHistory()
    useBaseStore.getState().setTableData(nodeId, [], [])
    useBaseStore.getState().renameTable(nodeId, 'Excel转DXF')
    setSummary('已清除数据')
  }

  const exportSelfCsv = async (): Promise<void> => {
    if (!self) return
    const fields = self.fields.map((f) => ({ id: f.name, name: f.name, type: 'text' as const }))
    await exportRows(`${self.from}.csv`, fields, self.rows as never, 'csv')
  }

  const sel = (value: string, onChange: (v: string) => void, placeholder: string, allowEmpty: boolean) => (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      style={{ flex: 1, minWidth: 0, fontSize: 11, padding: '2px 4px', border: '1px solid var(--c-border)', borderRadius: 4, background: 'var(--c-bg1)', color: 'var(--c-text1)' }}
    >
      {(allowEmpty ? [{ n: '', label: '（无）' }] : []).concat(fieldNames.map((n) => ({ n, label: n }))).map((x) => (
        <option key={x.n || '_'} value={x.n}>{x.label || placeholder}</option>
      ))}
      {!fieldNames.length && <option value="">{placeholder}</option>}
    </select>
  )

  return (
    <div
      className="pn-root"
      style={{ gap: 5, outline: dragOver ? '2px dashed var(--c-blue, #3b82f6)' : undefined, borderRadius: 8, transition: 'outline-color .15s' }}
      onDragEnter={(e) => { e.preventDefault(); dragDepth.current++; setDragOver(true) }}
      onDragOver={(e) => { e.preventDefault(); try { e.dataTransfer.dropEffect = 'copy' } catch { /* ignore */ } }}
      onDragLeave={() => { dragDepth.current--; if (dragDepth.current <= 0) { dragDepth.current = 0; setDragOver(false) } }}
      onDrop={async (e) => {
        e.preventDefault(); e.stopPropagation()
        dragDepth.current = 0; setDragOver(false)
        if (e.dataTransfer.files?.length) {
          const f = e.dataTransfer.files[0]
          await loadDataFile(f)
          pushLog(`拖入文件已载入：${f.name}`)
        }
      }}
    >
      {/* ★ 数据文件条目（pn-row 对齐 PDF 节点条目）：pm-fmt X 方块 + 名 + 行数 + 拖出窗体=导出 CSV + × 清除 */}
      {self && (
        <div
          className="pn-row nodrag nopan"
          draggable
          title="已载入的数据文件（拖出窗体 = 导出 CSV；× 清除）"
          onDragStart={(e) => {
            e.dataTransfer.setData('application/x-cb-xlsx-src', JSON.stringify({ nodeId, name: self.from }))
            e.dataTransfer.effectAllowed = 'copyMove'
          }}
          onDragEnd={(e) => {
            const out = e.clientX < 0 || e.clientY < 0 || e.clientX > window.innerWidth || e.clientY > window.innerHeight
            if (!out) return
            void exportSelfCsv()
            setSummary('已拖出窗体：数据导出 CSV')
          }}
        >
          <span className="pm-fmt" style={{ background: '#217346', flexShrink: 0 }}>X</span>
          <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{self.from} · {self.rows.length} 行</span>
          <button
            className="pm-del"
            title="清除已载入数据"
            draggable={false}
            onClick={(e) => { e.stopPropagation(); clearSelf() }}
          >×</button>
        </div>
      )}

      {/* ★ 空态导入长条钮（规范第100行 .pn-empty-btn：全宽虚线 + hover 蓝；纯文字空态已废） */}
      {!self && !upstream && !src && (
        <button
          className="pn-empty-btn nodrag nopan"
          title="点击导入 Excel（可拖入；或连线表格节点）"
          onClick={() => fileRef.current?.click()}
        >＋ 点击导入 Excel（也可拖入）</button>
      )}
      {active && self && <div className="pn-row" style={{ color: 'var(--c-text3)', fontSize: 10.5 }}>列映射基于：{active.from}（{active.rows.length} 行 × {active.fields.length} 列）</div>}

      {/* 列映射（pn-row 容器） */}
      <div className="pn-row" style={{ flexDirection: 'column', alignItems: 'stretch' }}>
        <div style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
          <span style={{ width: 44, flexShrink: 0, color: 'var(--c-text3)' }}>点号列</span>
          {sel(m.nameCol ?? '', (v) => setMap({ nameCol: v || undefined }), '自动 P 序号', true)}
        </div>
        <div style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
          <span style={{ width: 44, flexShrink: 0, color: 'var(--c-red, #ef4444)' }}>X 列 *</span>
          {sel(m.xCol, (v) => setMap({ xCol: v }), '选 X 列', false)}
        </div>
        <div style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
          <span style={{ width: 44, flexShrink: 0, color: 'var(--c-green, #22c55e)' }}>Y 列 *</span>
          {sel(m.yCol, (v) => setMap({ yCol: v }), '选 Y 列', false)}
        </div>
        <div style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
          <span style={{ width: 44, flexShrink: 0, color: 'var(--c-text3)' }}>高程列</span>
          {sel(m.zCol ?? '', (v) => setMap({ zCol: v || undefined }), '（可选）', true)}
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', color: 'var(--c-text2)' }}>
          {([['drawPoint', '生成点'], ['drawLabel', '点号标注'], ['drawLine', '按行连线'], ['flipY', '翻转Y']] as const).map(([k, label]) => (
            <label key={k} style={{ display: 'flex', gap: 3, alignItems: 'center', cursor: 'pointer' }}>
              <input type="checkbox" checked={o[k]} onChange={(e) => setOpt({ [k]: e.target.checked })} />{label}
            </label>
          ))}
        </div>
      </div>

      {/* ★ 装载条（规范第五章唯一）：导入(无数据时) → 历史(有历史时) → 生成 → 边界整图 */}
      <NodeBtnBar className="nb-bar--flat">
        {histCount > 0 && (
          <NodeSqBtn icon={<IcHistory size={19} />} title="生成历史（双击条目预览 · 拖出窗体另存 · 右键菜单）" active={histOpen} onClick={() => setHistOpen((v) => !v)} />
        )}
        <NodeSqBtn
          icon={<IcDownload size={19} />}
          title="生成 DXF（按列映射展点：POINT + 点号标注 + 可选连线）"
          disabled={!active}
          busy={busy}
          onClick={() => { void generate('points') }}
        />
        <NodeSqBtn
          icon={<IcGrid size={19} />}
          title="非空边界整图转DXF：确定表格非空区域，整表画成网格+文字（任何 Excel 表通用）"
          disabled={!active}
          busy={busy}
          onClick={() => { void generate('table') }}
        />
      </NodeBtnBar>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) void loadDataFile(f); e.target.value = '' }} />

      {/* 执行摘要（pn-row） */}
      {summary && <div className="pn-row" style={{ color: 'var(--c-text2)' }}>{summary}</div>}

      {/* ★ 生成历史列表（🕘 钮展开；pn-row 条目：序号+内/外+名+时间+操作；双击预览/拖出另存/右键菜单） */}
      {histCount > 0 && histOpen && (
        <div className="pm-hist nodrag nopan" style={{ padding: '2px 0' }}>
          {hist.map((h, i) => (
            <div
              key={h.id + ':' + i}
              className="pn-row"
              draggable
              title={h.name + `（${h.count} 格${h.skipped ? '，跳过 ' + h.skipped : ''}；双击预览；拖出窗体=另存；右键菜单）`}
              onDragStart={(e) => { e.dataTransfer.effectAllowed = 'copy' }}
              onDragEnd={(e) => {
                const out = e.clientX < 0 || e.clientY < 0 || e.clientX > window.innerWidth || e.clientY > window.innerHeight
                if (!out) return
                void saveAsDownload(h)
                setSummary(`已拖出窗体：另存 ${h.name}`)
              }}
              onDoubleClick={() => { void openPreview(h) }}
              onContextMenu={(e) => { e.preventDefault(); e.stopPropagation(); setMenu({ x: e.clientX, y: e.clientY, entry: h }) }}
              onPointerDown={(e) => e.stopPropagation()}
              onPointerUp={(e) => e.stopPropagation()}
            >
              <span className="pm-idx">{i + 1}</span>
              <span className={'pm-hist-src ' + (h.inDb ? 'pm-hist-src--canvas' : 'pm-hist-src--ext')}>{h.inDb ? '内' : '外'}</span>
              <span className="pm-nm" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{h.name}</span>
              <span className="pm-hist-at">{new Date(h.at).toLocaleString('zh-CN', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })}</span>
              <button
                className="pm-del"
                title="删除这条历史记录"
                draggable={false}
                onClick={(e) => {
                  e.stopPropagation()
                  useBaseStore.getState().pushHistory()
                  patch({ xlsx2dxfHistory: hist.filter((hh) => hh.id !== h.id || hh.at !== h.at) })
                }}
              >×</button>
            </div>
          ))}
        </div>
      )}

      {/* 居中预览窗（DXF 文本；对齐 pm-pdfpreview 结构） */}
      {preview && createPortal(
        <div className="pm-pdfpreview" onClick={() => setPreview(null)}>
          <div className={'pn-preview-card pm-pdfpreview-card' + (previewMax ? ' pm-card-max' : '')} onClick={(e) => e.stopPropagation()}>
            <div className="pn-preview-bar">
              <span className="pn-preview-nm">🕘 {preview.name}</span>
              <button className="pn-preview-close" title={previewMax ? '还原' : '最大化'} onClick={() => setPreviewMax((v) => !v)}>
                {previewMax ? <IcMinimize size={19} /> : <IcMaximize size={19} />}
              </button>
              <button className="pn-preview-close" title="关闭预览" onClick={() => setPreview(null)}><IcX size={19} /></button>
            </div>
            <div className="pm-preview-body" style={{ overflow: 'auto', padding: '10px 14px', display: 'flex', justifyContent: 'center' }}>
              <DxfSvgPreview dxf={preview.text} />
            </div>
          </div>
        </div>,
        document.body,
      )}

      {/* 右键菜单：打开 / 另存为 / 删除记录（对齐 PDF 历史三件套） */}
      {menu && createPortal(
        <div
          className="libpick-mask nodrag nopan"
          style={{ background: 'transparent' }}
          onClick={() => setMenu(null)}
          onContextMenu={(e) => { e.preventDefault(); setMenu(null) }}
        >
          <div
            className="pm-ctx-menu"
            style={{ left: Math.min(menu.x, window.innerWidth - 170), top: Math.min(menu.y, window.innerHeight - 90) }}
            onClick={(e) => e.stopPropagation()}
          >
            <button className="pm-ctx-item" onClick={() => { const ent = menu.entry; setMenu(null); void openPreview(ent) }}>打开</button>
            <button className="pm-ctx-item" onClick={() => { const ent = menu.entry; setMenu(null); void saveAsDownload(ent) }}>另存为…</button>
            <button
              className="pm-ctx-item"
              onClick={() => {
                const ent = menu.entry
                setMenu(null)
                useBaseStore.getState().pushHistory()
                patch({ xlsx2dxfHistory: hist.filter((hh) => hh.id !== ent.id || hh.at !== ent.at) })
              }}
            >删除记录</button>
          </div>
        </div>,
        document.body,
      )}
    </div>
  )
}


/* ── DXF 预览渲染（2026-11-10 用户「右键打开也显示不对」：原始组码文本→图形可视化）──
 *  解析 POINT/TEXT/LINE → SVG（Y 翻转、按边界自适应、按层配色、暗色主题跟随） */
function DxfSvgPreview({ dxf }: { dxf: string }) {
  const ent = useMemo(() => parseDxfEntities(dxf), [dxf])
  const total = ent.points.length + ent.texts.length + ent.lines.length
  if (!total) return <div className="pn-row" style={{ color: 'var(--c-text2)' }}>无可预览的图形实体</div>
  const xs: number[] = []
  const ys: number[] = []
  for (const p of ent.points) { xs.push(p.x); ys.push(p.y) }
  for (const t of ent.texts) { xs.push(t.x); ys.push(t.y) }
  for (const l of ent.lines) { xs.push(l.x1, l.x2); ys.push(l.y1, l.y2) }
  const minX = Math.min(...xs); const maxX = Math.max(...xs)
  const minY = Math.min(...ys); const maxY = Math.max(...ys)
  const pad = Math.max((maxX - minX) * 0.06, (maxY - minY) * 0.06, 1)
  const svgY = (y: number): number => maxY + pad - y /* DXF Y 向上 → SVG Y 向下 */
  const colorOf = (layer: string): string =>
    layer.includes('TABLE') ? '#22c55e' : layer.includes('POINTS') ? '#ef4444' : layer.includes('LINES') ? '#3b82f6' : '#eab308'
  return (
    <svg
      viewBox={`${minX - pad} ${-pad} ${maxX - minX + pad * 2} ${maxY - minY + pad * 2}`}
      style={{ maxWidth: '100%', maxHeight: '62vh', color: 'var(--c-text1)' }}
    >
      {ent.lines.map((l, i) => (
        <line key={'l' + i} x1={l.x1} y1={svgY(l.y1)} x2={l.x2} y2={svgY(l.y2)} stroke={colorOf('CB_LINES')} strokeWidth={Math.max((maxY - minY) / 400, 0.4)} />
      ))}
      {ent.points.map((p, i) => (
        <circle key={'p' + i} cx={p.x} cy={svgY(p.y)} r={Math.max((maxY - minY) / 150, 0.8)} fill="#ef4444" />
      ))}
      {ent.texts.map((t, i) => (
        <text key={'t' + i} x={t.x} y={svgY(t.y)} fontSize={t.h} fill="#eab308" style={{ fontFamily: 'Consolas, monospace' }}>{t.text}</text>
      ))}
    </svg>
  )
}
