/* ─────────────────────────────────────────────────────────────
 * 网址监控内核（2026-11-09 用户指令「增加一个网址监控节点…提取、记录、或者输出数据」）
 *
 * 能力集参照 changedetection.io（35.3k stars）：URL 列表、CSS/JSON/关键词提取、
 * 变化检测标记、历史快照、CSV 导出。
 *
 * 分层：本文件是渲染层调度与数据管理——网络与提取走主进程 IPC（webmon:fetch，
 * 绕渲染层 CORS，cheerio/JSON 解析在主进程）。轮询用渲染层计时器（应用开着才监控）。
 * ───────────────────────────────────────────────────────────── */
import { useEffect } from 'react'
import { useBaseStore } from './store'
import { pushLog } from '../ui/logBus'
import { setNodeRunStatus, setNodeRunStatusAuto } from './nodeRunStatus'

export interface WebmonTarget {
  id: string
  url: string
  name?: string
  mode: 'text' | 'json' | 'keyword'
  selector?: string
  jsonPath?: string
  keyword?: string
}
export interface WebmonHistEntry {
  at: number
  targetName: string
  value: string
  changed: boolean
  error?: string
}
export interface WebmonConfig {
  webmonTargets?: WebmonTarget[]
  webmonIntervalSec?: number
  webmonHistory?: WebmonHistEntry[]
  webmonLastValues?: Record<string, string>
  webmonLastRun?: number
}

const HISTORY_CAP = 100
/** 值快照截断：历史里只存前 300 字（全量值在 lastValues 里用于下次比对，也截断防 config 膨胀） */
const VALUE_CAP = 300

type WebmonApi = {
  webmonFetch?: (req: { url: string; mode: 'text' | 'json' | 'keyword'; selector?: string; keyword?: string; jsonPath?: string }) => Promise<{ ok: boolean; value?: string; error?: string }>
}
const webmonApi = (): WebmonApi | undefined =>
  (globalThis as { api?: WebmonApi }).api

/** 运行中防重入（轮询 5s tick + 手动按钮可能撞车） */
const runningChecks = new Set<string>()

interface CheckResult { label: string; value: string; changed: boolean; error?: string }

/** ★ 检查完成后推送到连线下游（2026-11-09 用户「提取信息怎么写入下一个节点」）：
 *  找出把本节点列为上游（linkedInputs/inputs）的节点，按 kind 分发：
 *  - text：单值写入（第一条成功结果「目标：值」）
 *  - table/param/test_table：追加一行，按下游表现有字段名模糊匹配填值（时间/目标/值） */
function pushToDownstream(nodeId: string, results: CheckResult[]): void {
  const st = useBaseStore.getState()
  const okResults = results.filter((r) => !r.error)
  const first = okResults[0] ?? results[0]
  for (const t of Object.values(st.tables)) {
    if (!t || t.deleted) continue
    const cfg = (t.config ?? {}) as { linkedInputs?: string[]; inputs?: string[] }
    const ups = cfg.linkedInputs ?? cfg.inputs ?? []
    if (!ups.includes(nodeId)) continue
    try {
      if (t.kind === 'text') {
        st.setTableData(t.id, [{ id: 'text', name: 'text', type: 'text' }], [{ text: first ? `${first.label}：${first.value}` : '' }])
        pushLog(`[网址监控] 已写入下游文本节点「${t.name}」`, 'ok')
      } else if (t.kind === 'ai') {
        /* ★ 2026-11-09「下游的节点能介入AI分析」：检测到变化才自动运行（无变化不浪费算力）；
           runAiNode 在 core/aiRun（面板按钮与自动触发同一条链路），延迟 1s 等数据落定 */
        if (results.some((r) => r.changed)) {
          const aiId = t.id
          window.setTimeout(() => {
            void import('./aiRun').then((m) => m.runAiNode(aiId, useBaseStore.getState())).then((r) => {
              pushLog(`[网址监控] 自动触发 AI 分析「${t.name}」：${r.ok ? '完成' : r.msg.slice(0, 80)}`, r.ok ? 'ok' : 'err')
            })
          }, 1000)
          pushLog(`[网址监控] 检测到变化，自动触发下游 AI「${t.name}」分析…`, 'ok')
        }
      } else if (t.kind === 'table' || t.kind === 'param' || t.kind === 'test_table') {
        const fields = t.fields ?? []
        if (!fields.length) continue
        const newRow: Record<string, unknown> = {}
        for (const f of fields) {
          const n = f.name
          if (/时间|日期|at$|time/i.test(n)) newRow[n] = new Date().toLocaleString('zh-CN')
          else if (/目标|来源|站点|url|site/i.test(n)) newRow[n] = first?.label ?? ''
          else if (/值|内容|value|数据|提取/i.test(n)) newRow[n] = first?.value ?? ''
        }
        const recs = st.records[t.id] ?? []
        st.setTableData(t.id, fields, [...recs, newRow])
        pushLog(`[网址监控] 已追加到下游表「${t.name}」（第 ${recs.length + 1} 行）`, 'ok')
      }
    } catch (e) {
      pushLog(`[网址监控] 写入下游「${t.name}」失败：${String((e as Error)?.message ?? e).slice(0, 100)}`, 'err')
    }
  }
}

/** 立即检查一个监控节点的全部目标：并发 fetch → 与上次值比对 → 标变化 → 写历史 */
export async function checkWebmonNow(nodeId: string): Promise<{ ok: boolean; changed: number; errors: number; msg: string }> {
  if (runningChecks.has(nodeId)) return { ok: false, changed: 0, errors: 0, msg: '检查进行中' }
  const st = useBaseStore.getState()
  const t = st.tables[nodeId]
  if (!t || t.deleted) return { ok: false, changed: 0, errors: 0, msg: '节点不存在（可能已被删除）' }
  const cfg = (t.config ?? {}) as WebmonConfig
  const targets = cfg.webmonTargets ?? []
  if (!targets.length) return { ok: false, changed: 0, errors: 0, msg: '没有监控目标——先添加 URL' }
  const api = webmonApi()
  if (!api?.webmonFetch) return { ok: false, changed: 0, errors: 0, msg: '缺少网址监控模块（请重启应用获取桥）' }

  runningChecks.add(nodeId)
  setNodeRunStatus(nodeId, 'running', `检查 ${targets.length} 个目标…`)
  try {
    const lastValues: Record<string, string> = { ...(cfg.webmonLastValues ?? {}) }
    const hist: WebmonHistEntry[] = [...(cfg.webmonHistory ?? [])]
    let changed = 0
    let errors = 0
    const results: CheckResult[] = []

    await Promise.all(
      targets.map(async (tg) => {
        const label = tg.name || tg.url
        try {
          const r = await api.webmonFetch!({ url: tg.url, mode: tg.mode, selector: tg.selector, keyword: tg.keyword, jsonPath: tg.jsonPath })
          if (!r.ok) {
            errors++
            results.push({ label, value: '', changed: false, error: r.error ?? '未知错误' })
            hist.unshift({ at: Date.now(), targetName: label, value: '', changed: false, error: r.error ?? '未知错误' })
            return
          }
          const value = (r.value ?? '').slice(0, VALUE_CAP)
          const prev = lastValues[tg.id]
          /* 首次检查不算变化（changedetection.io 同语义：第一次快照只是基线） */
          const isChanged = prev !== undefined && prev !== value
          if (isChanged) changed++
          lastValues[tg.id] = value
          results.push({ label, value, changed: isChanged })
          hist.unshift({ at: Date.now(), targetName: label, value, changed: isChanged })
        } catch (e) {
          errors++
          const msg = String((e as Error)?.message ?? e).slice(0, 120)
          results.push({ label, value: '', changed: false, error: msg })
          hist.unshift({ at: Date.now(), targetName: label, value: '', changed: false, error: msg })
        }
      }),
    )

    const cappedHist = hist.slice(0, HISTORY_CAP)
    useBaseStore.getState().setNodeConfig(nodeId, {
      webmonHistory: cappedHist,
      webmonLastValues: lastValues,
      webmonLastRun: Date.now(),
    })
    /* ★ 自带 records（2026-11-09「打通整个交互」）：历史映射为 [时间/目标/值/变化] 行——
       下游 sort/map/output/引用体系走通用引擎自动可读，双击节点就是表格视图 */
    try {
      useBaseStore.getState().setTableData(
        nodeId,
        [
          { id: 'wm-time', name: '时间', type: 'text' },
          { id: 'wm-target', name: '目标', type: 'text' },
          { id: 'wm-value', name: '值', type: 'text' },
          { id: 'wm-changed', name: '变化', type: 'text' },
        ],
        cappedHist.map((h) => ({
          时间: new Date(h.at).toLocaleString('zh-CN'),
          目标: h.targetName,
          值: h.error ? `✕ ${h.error}` : h.value,
          变化: h.changed ? '是' : '否',
        })),
      )
    } catch { /* records 同步失败不影响主流程 */ }
    /* ★ 检查完成 → 推送连线下游（text 写值 / 表格追加行）；检测到变化时自动触发下游 AI 分析 */
    pushToDownstream(nodeId, results)
    const msg = `${targets.length} 目标 · ${changed} 变化 · ${errors} 失败`
    if (errors === targets.length) setNodeRunStatus(nodeId, 'error', '全部目标检查失败')
    else if (changed > 0) setNodeRunStatusAuto(nodeId, 'ok', `检测到 ${changed} 处变化`)
    else setNodeRunStatusAuto(nodeId, 'ok', `${targets.length} 个目标无变化`)
    pushLog(`[网址监控] ${t.name}：${msg}`, errors === targets.length ? 'err' : 'ok')
    return { ok: true, changed, errors, msg }
  } finally {
    runningChecks.delete(nodeId)
  }
}

/** 历史 → CSV（BOM 头让 Excel 正确识别 UTF-8；值内引号/换行转义） */
export function webmonToCsv(hist: WebmonHistEntry[]): string {
  const esc = (s: string) => `"${s.replace(/"/g, '""')}"`
  const rows = [['时间', '目标', '值', '变化', '错误'].join(',')]
  for (const h of hist) {
    rows.push([new Date(h.at).toLocaleString('zh-CN'), esc(h.targetName), esc(h.value), h.changed ? '是' : '否', esc(h.error ?? '')].join(','))
  }
  return '\ufeff' + rows.join('\r\n')
}

/** 触发浏览器下载 CSV */
export function downloadWebmonCsv(hist: WebmonHistEntry[], nodeName: string): void {
  const blob = new Blob([webmonToCsv(hist)], { type: 'text/csv;charset=utf-8' })
  const a = document.createElement('a')
  a.href = URL.createObjectURL(blob)
  a.download = `${nodeName || '网址监控'}-历史.csv`
  a.click()
  setTimeout(() => URL.revokeObjectURL(a.href), 2000)
}

/** 轮询器（挂画布级）：5s tick 扫描全部 webmon 节点，到期（intervalSec≥10 且距上次检查足够久）自动检查 */
export function useWebmonPoller(): void {
  useEffect(() => {
    const timer = window.setInterval(() => {
      const st = useBaseStore.getState()
      const now = Date.now()
      for (const t of Object.values(st.tables)) {
        if (!t || t.deleted || t.kind !== 'webmon') continue
        const cfg = (t.config ?? {}) as WebmonConfig
        const iv = cfg.webmonIntervalSec ?? 0
        if (iv < 10) continue
        const last = cfg.webmonLastRun ?? 0
        if (now - last < iv * 1000) continue
        if (runningChecks.has(t.id)) continue
        void checkWebmonNow(t.id)
      }
    }, 5000)
    return () => window.clearInterval(timer)
  }, [])
}
