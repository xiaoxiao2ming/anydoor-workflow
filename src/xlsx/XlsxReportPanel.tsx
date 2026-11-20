import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { NodeConfig, TableSchema, XlsxBinding, FieldDef } from '../core/schema'
import type { Row } from '../core/store'
import { computeNode } from '../core/engine'
import { bindXlsxFromParamTable, discoverXlsxFieldsFromBytes, type XlsxFieldSpec } from './field'
import { asOpenFailure, WorkbookOpenError, type OpenFailure } from './openError'
import { OpenFailureNotice } from './OpenFailureNotice'
import { buildXlsxParamTable } from './buildXlsxData'
import { displayRef } from './domainMarks'
import { getTemplateStore, saveDocAs } from '../docx/templateStore'
import { OnlyOfficeModal } from '../office/OnlyOfficeModal'

/**
 * xlsx 节点的「Excel 报表」区块 —— 与 WordReportPanel 对称的另一半。
 *
 * Excel 的「域」用原生机制承载（不是我们发明的标记）：
 *   命名区域（definedName）  ←→  标量域
 *   Table（ListObject）      ←→  行级循环域
 * 好处是用户自己在 Excel 里也能定义域，我们只负责识别与刷新。
 */

const AGGS: { v: NonNullable<XlsxBinding['agg']>; t: string }[] = [
  { v: 'first', t: '首行' },
  { v: 'sum', t: '合计' },
  { v: 'count', t: '计数' },
  { v: 'max', t: '最大' },
  { v: 'min', t: '最小' },
  { v: 'avg', t: '平均' },
]

export function XlsxReportPanel({
  table,
  tables,
  records,
  onSetNodeConfig,
  refreshKey,
}: {
  table: TableSchema
  tables: Record<string, TableSchema>
  records: Record<string, Row[]>
  onSetNodeConfig: (tableId: string, patch: Partial<NodeConfig>) => void
  /**
   * 外部改了模板内容（比如在「打开工作簿」里改了单元格 / 加了命名区域）时递增。
   * 光靠 templateId 判断会漏 —— id 没变但内容变了，域清单必须重新发现。
   */
  refreshKey?: number
}) {
  const config = table.config ?? {}
  const templateId = config.xlsxTemplateId
  const map = config.xlsxMap ?? {}

  const [tplBytes, setTplBytes] = useState<Uint8Array | null>(null)
  const [slots, setSlots] = useState<XlsxFieldSpec[]>([])
  const [warn, setWarn] = useState<string[]>([])
  const [busy, setBusy] = useState('')
  const [err, setErr] = useState('')
  /**
   * 「模板打不开」单独一类：这是客户文件最常见的支持问题，
   * 值得给完整的人话提示 + 自动留样本，而不是一句话糊过去。
   */
  const [openFail, setOpenFail] = useState<OpenFailure | null>(null)
  /** 手动重试用：模板 id 没变时把载入 effect 顶起来 */
  const [retryTick, setRetryTick] = useState(0)
  const [lastOut, setLastOut] = useState('')
  /** 生成后每个域「取到了什么、从哪来」——溯源自证，送审被问时能直接答 */
  const [trace, setTrace] = useState<{ key: string; from: string; value: string; ok: boolean }[]>([])
  /** OnlyOffice 在线编辑：当前在工作区里的文件名 */
  const [editing, setEditing] = useState<string | null>(null)
  /** DS 保存回来的字节，供「保存到本机」 */
  const [editedBytes, setEditedBytes] = useState<Uint8Array | null>(null)
  const loadedId = useRef<string | null>(null)

  // 上游字段：供绑定选择器用
  const upstream = useMemo(() => {
    const linked = table.config?.linkedInputs ?? []
    const ins = [...linked, ...((table.config?.inputs ?? []).filter((x) => !linked.includes(x)))]
    const fields: FieldDef[] = []
    const seen = new Set<string>()
    for (const id of ins) {
      if (!tables[id]) continue
      for (const f of computeNode(id, tables, records).fields) {
        if (!seen.has(f.id)) {
          seen.add(f.id)
          fields.push(f)
        }
      }
    }
    return { fields, ids: ins }
  }, [table, tables, records])

  const setBinding = (key: string, next: XlsxBinding | null) => {
    const m = { ...map }
    if (next) m[key] = next
    else delete m[key]
    onSetNodeConfig(table.id, { xlsxMap: m })
  }

  // 载入当前绑定的模板，并把里面的「域」扫出来
  useEffect(() => {
    let cancelled = false
    if (!templateId) {
      setTplBytes(null)
      setSlots([])
      loadedId.current = null
      return
    }
    // 注意：判等键要带 refreshKey —— 用户可能在「打开工作簿」里改了模板内容，
    // 此时 templateId 没变，只按 id 判等会让域清单一直是旧的。
    const loadKey = `${templateId}@${refreshKey ?? 0}`
    if (loadedId.current === loadKey) return
    loadedId.current = loadKey
    ;(async () => {
      try {
        const bytes = await getTemplateStore().read(templateId)
        if (!bytes) {
          if (!cancelled) {
            setErr('模板文件已不在模板库中，请重新导入')
            setTplBytes(null)
            setSlots([])
          }
          return
        }
        if (cancelled) return
        setTplBytes(bytes)
        const found = await discoverXlsxFieldsFromBytes(bytes)
        if (cancelled) return
        setSlots(found)
        setErr('')
        setOpenFail(null)
        // 首次发现域时，为每个域补一条默认绑定（按同名上游字段自动匹配）
        if (found.length) {
          const patch: Record<string, XlsxBinding> = { ...map }
          for (const s of found) {
            if (patch[s.key]) continue
            const f = upstream.fields.find((x) => x.id === s.key || x.name === s.key)
            if (!f) continue
            patch[s.key] = s.kind === 'table' ? { from: 'rows', fieldId: f.id } : { from: 'field', fieldId: f.id }
          }
          onSetNodeConfig(table.id, { xlsxMap: patch })
        }
      } catch (e) {
        if (cancelled) return
        if (e instanceof WorkbookOpenError) {
          // 模板本身读不了：这是「文件/兼容性」问题，不是「操作」问题，
          // 用结构化提示说清是哪一类，并把样本留下来。
          setOpenFail(asOpenFailure(e))
          setSlots([])
          setErr('')
        } else {
          setErr('读取模板失败：' + (e as Error).message)
        }
      }
    })()
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [templateId, refreshKey])

  /** 由某个已选中的文件导入（浏览器 input 路径）—— 文件级入口已上移标题栏，此文件不再使用 */

  const generate = async () => {
    if (!tplBytes) return
    setBusy('生成中…')
    setErr('')
    setWarn([])
    try {
      const res = await buildBound()
      const name = `${table.name}.xlsx`
      const path = await saveDocAs(name, res.buffer, 'xlsx')
      setLastOut(path ? `已导出：${path}` : '已生成（导出被取消）')
    } catch (e) {
      setErr('生成失败：' + (e as Error).message)
    } finally {
      setBusy('')
    }
  }

  /** 绑定并产出字节（导出与在线编辑共用同一条路径，避免两边算出不同结果） */
  const buildBound = useCallback(async () => {
    if (!tplBytes) throw new Error('还没选模板')
    const built = buildXlsxParamTable(slots, map, table, tables, records)
    const res = await bindXlsxFromParamTable(tplBytes, built.table, { tableFormulas: built.formulas })
    setWarn([...built.missing.map((k) => `未取到值：${k}`), ...res.warnings])
    // 溯源：把「值从哪来」落成一行行可读记录（URN 能答"这个数哪来的"）
    setTrace(
      res.traces.map((t) => ({
        key: t.key,
        from:
          t.from === 'override'
            ? '手填'
            : t.from === 'default'
              ? '引擎计算值'
              : t.from === 'unresolved'
                ? '未解析'
                : t.urn,
        value: fmtTrace(t.value),
        ok: t.ok,
      })),
    )
    return res
  }, [tplBytes, slots, map, table, tables, records])

  /**
   * 在线编辑：绑定产物推到 app-server 工作目录 → DS（docType 由扩展名判定为 cell）
   * 打开编辑器。DS 保存后写回工作目录，再由「保存到本机」取回。
   *
   * 与 Word 侧的差别：Excel 没有内置的 WPS 级编辑器（Univer 的 xlsx IO 属付费 Pro 层），
   * 所以重排格式、图表这类操作**只能**走这条路；离线时用「打开工作簿」看内容改数据。
   */
  const openInEditor = async () => {
    setErr('')
    setBusy('正在上传到编辑工作区…')
    try {
      const res = await buildBound()
      let bin = ''
      const CH = 0x8000
      for (let i = 0; i < res.buffer.length; i += CH) {
        bin += String.fromCharCode(...res.buffer.subarray(i, i + CH))
      }
      const name = `${table.name || '工作簿'}.xlsx`
      const r = await fetch('/office/upload', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, dataBase64: btoa(bin) }),
      })
      if (!r.ok) throw new Error(`上传失败（${r.status}）——app-server 是否在运行？`)
      setEditing(name)
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy('')
    }
  }

  /** DS 保存回调完成（文件已写回工作目录）→ 拉回来供下载 */
  const onOfficeSaved = useCallback(async () => {
    if (!editing) return
    try {
      const r = await fetch(`/office/files/${encodeURIComponent(editing)}`)
      if (!r.ok) return
      setEditedBytes(new Uint8Array(await r.arrayBuffer()))
    } catch {
      /* 服务端已确认写回，拉取失败不打断 */
    }
  }, [editing])

  return (
    <div className="xrp">
      {/* 文件级操作（导入/历史/保存/生成）已上移到编辑器标题栏 —— 这里只管「域绑定与产出」 */}

      {busy && <div className="xrp-busy">{busy}</div>}
      {openFail && (
        <OpenFailureNotice
          failure={openFail}
          name={config.xlsxTemplateName ?? '模板.xlsx'}
          bytes={tplBytes}
          onRetry={() => {
            loadedId.current = null
            setOpenFail(null)
            // 触发一次重新载入：把 key 置空后由 effect 的依赖变化驱动
            onSetNodeConfig(table.id, { xlsxMap: { ...config.xlsxMap } })
          }}
        />
      )}
      {err && <div className="xrp-err">{err}</div>}

      {!slots.length ? null : (
        <>
          <div className="xrp-meta">
            发现 {slots.length} 个域 · 已绑 {Object.keys(map).length} 个 · 上游 {upstream.fields.length} 个字段
          </div>
          <div className="xrp-list">
            {slots.map((s) => (
              <SlotRow
                key={s.key}
                slot={s}
                spec={map[s.key]}
                fields={upstream.fields}
                onChange={(next) => setBinding(s.key, next)}
              />
            ))}
          </div>
        </>
      )}

      <div className="xrp-foot">
        <button className="bar-btn primary export-btn" disabled={!tplBytes || !!busy} onClick={() => void generate()}>
          生成 Excel
        </button>
        <button
          className="bar-btn"
          disabled={!tplBytes || !!busy}
          onClick={() => void openInEditor()}
          title="把绑定产物推到编辑工作区，用 OnlyOffice 像 Excel 一样编辑（重排格式、加图表只能走这条）"
        >
          在线编辑
        </button>
        {editedBytes && (
          <button
            className="bar-btn"
            onClick={() => void saveDocAs(`${table.name}-已编辑.xlsx`, editedBytes, 'xlsx')}
            title="把 DS 里保存回来的结果存到本机"
          >
            取回编辑稿…
          </button>
        )}
        {lastOut && <span className="xrp-ok">{lastOut}</span>}
      </div>

      {editing && (
        <OnlyOfficeModal
          file={editing}
          onSaved={() => void onOfficeSaved()}
          onClose={() => setEditing(null)}
          hint={
            <>
              这里编辑的是<b>绑定产物</b>（DS 按扩展名判定为 Excel 表格）。保存后点「取回编辑稿」存到本机。
              <br />
              只想看内容 / 改数据 / 认域位置，用「打开工作簿」页签——那条路不依赖 Document Server。
            </>
          }
        />
      )}

      {warn.length > 0 && (
        <div className="xrp-warn">
          {warn.map((w, i) => (
            <div key={i}>· {w}</div>
          ))}
        </div>
      )}

      {trace.length > 0 && (
        <details className="xrp-trace" open>
          <summary>取值溯源（{trace.length} 个域）</summary>
          <div className="xrp-trace-body">
            {trace.map((t, i) => (
              <div key={i} className={`xrp-trace-row${t.ok ? '' : ' xrp-trace-row--bad'}`}>
                <span className="xrp-trace-key">{t.key}</span>
                <span className="xrp-trace-from" title={t.from}>
                  {t.from}
                </span>
                <span className="xrp-trace-val" title={t.value}>
                  {t.value}
                </span>
              </div>
            ))}
          </div>
        </details>
      )}
    </div>
  )
}

/** 溯源值的一行显示：数组只说行数，字符串截断 */
function fmtTrace(v: unknown): string {
  if (v === undefined) return '（空）'
  if (v === null) return 'null'
  if (Array.isArray(v)) return `数组 ${v.length} 行`
  if (typeof v === 'object') return JSON.stringify(v).slice(0, 40)
  const s = String(v)
  return s.length > 40 ? s.slice(0, 40) + '…' : s
}

/* ────────────── 单个域 ────────────── */

function SlotRow({
  slot,
  spec,
  fields,
  onChange,
}: {
  slot: XlsxFieldSpec
  spec: XlsxBinding | undefined
  fields: FieldDef[]
  onChange: (next: XlsxBinding | null) => void
}) {
  const isTable = slot.kind === 'table'
  const from: 'auto' | 'field' | 'rows' | 'manual' = !spec ? 'auto' : spec.from

  return (
    <div className={`xrp-item${isTable ? ' xrp-item--table' : ''}`}>
      <div className="xrp-item-head">
        <span className={`xrp-kind ${isTable ? 'xrp-kind--table' : 'xrp-kind--name'}`}>
          {isTable ? '表格' : '命名区域'}
        </span>
        <span className="xrp-key" title={`${slot.sheet ?? ''} ${slot.ref}`}>
          {slot.key}
        </span>
        <span className="xrp-ref">{displayRef(slot.sheet, slot.ref)}</span>
      </div>

      <div className="xrp-ctl">
        <select
          value={from}
          onChange={(e) => {
            const v = e.target.value
            if (v === 'auto') onChange(null)
            else onChange({ from: v as XlsxBinding['from'] })
          }}
        >
          <option value="auto">按同名上游字段</option>
          <option value="field">指定字段{isTable ? '（首行）' : ''}</option>
          {isTable && <option value="rows">整表（行级）</option>}
          <option value="manual">{isTable ? '手工 JSON' : '手工值'}</option>
        </select>

        {(from === 'field' || from === 'rows') && (
          <select
            value={spec?.fieldId ?? ''}
            onChange={(e) => onChange({ ...(spec ?? { from }), from, fieldId: e.target.value })}
          >
            <option value="">（选字段）</option>
            {fields.map((f) => (
              <option key={f.id} value={f.id}>
                {f.name}
              </option>
            ))}
          </select>
        )}

        {from === 'field' && (
          <select
            value={spec?.agg ?? 'first'}
            onChange={(e) => onChange({ ...(spec ?? { from }), from, agg: e.target.value as XlsxBinding['agg'] })}
            title="多行时如何取一个值"
          >
            {AGGS.map((a) => (
              <option key={a.v} value={a.v}>
                {a.t}
              </option>
            ))}
          </select>
        )}
      </div>

      {from === 'manual' && (
        <textarea
          className="xrp-manual"
          rows={isTable ? 4 : 1}
          placeholder={
            isTable
              ? '[{"名称":"钢管","数量":100},{"名称":"阀门","数量":12}]'
              : '手工填写，生成时原样写入这个域'
          }
          value={spec?.value ?? ''}
          onChange={(e) => onChange({ from: 'manual', value: e.target.value })}
        />
      )}

      {from === 'rows' && isTable && (
        <ColumnMapper
          columns={slot.columns ?? []}
          fields={fields}
          map={spec?.columns ?? {}}
          onChange={(cols) => onChange({ ...(spec ?? { from: 'rows' }), from: 'rows', columns: cols })}
          formulas={spec?.formulas ?? {}}
          onFormulas={(fm) => onChange({ ...(spec ?? { from: 'rows' }), from: 'rows', formulas: fm })}
        />
      )}
    </div>
  )
}

/** 列映射 + 计算列。计算列走 formula + cachedValue 双写，Excel 里能重算。 */
function ColumnMapper({
  columns,
  fields,
  map,
  onChange,
  formulas,
  onFormulas,
}: {
  columns: string[]
  fields: FieldDef[]
  map: Record<string, string>
  onChange: (next: Record<string, string>) => void
  formulas: Record<string, string>
  onFormulas: (next: Record<string, string>) => void
}) {
  return (
    <div className="xrp-cols">
      {columns.map((c) => (
        <div key={c} className="xrp-col">
          <span className="xrp-col-name" title={c}>
            {c}
          </span>
          <select
            value={map[c] ?? ''}
            onChange={(e) => {
              const next = { ...map }
              if (e.target.value) next[c] = e.target.value
              else delete next[c]
              onChange(next)
            }}
          >
            <option value="">（不填）</option>
            {fields.map((f) => (
              <option key={f.id} value={f.id}>
                {f.name}
              </option>
            ))}
          </select>
          <input
            className="xrp-formula"
            placeholder="计算列公式，如 ={数量}*{单价}"
            value={formulas[c] ?? ''}
            onChange={(e) => {
              const next = { ...formulas }
              if (e.target.value) next[c] = e.target.value
              else delete next[c]
              onFormulas(next)
            }}
            title="行级模板：{列名} 会换成该行的真实引用，写盘时 formula + 缓存值双写"
          />
        </div>
      ))}
    </div>
  )
}
