/* ─────────────────────────────────────────────────────────────
 * 网址监控节点卡片（2026-11-09 用户指令「增加一个网址监控节点…提取、记录、或者输出数据」）
 *
 * 卡片承载全部交互（folder 模式）：
 *  - 目标列表 CRUD（url + 提取模式 text/json/keyword + 规则）
 *  - 「立即检查」（全部目标并发，变化打标、error 打标）
 *  - 历史快照表（最近 12 条：时间/目标/值/变化标记）+ 「导出 CSV」
 *  - 轮询间隔在属性面板（webmonIntervalSec，≥10s 生效）
 * ───────────────────────────────────────────────────────────── */
import { useState } from 'react'
import { useBaseStore } from '../core/store'
import { checkWebmonNow, downloadWebmonCsv, type WebmonConfig, type WebmonTarget } from '../core/webmon'

const MODE_LABEL: Record<WebmonTarget['mode'], string> = { text: '网页文本', json: 'JSON 接口', keyword: '关键词' }

export function WebmonNodeCard({ nodeId }: { nodeId: string }) {
  const cfg = useBaseStore((s) => s.tables[nodeId]?.config) as WebmonConfig | undefined
  const name = useBaseStore((s) => s.tables[nodeId]?.name) ?? '网址监控'
  const targets = cfg?.webmonTargets ?? []
  const hist = cfg?.webmonHistory ?? []
  const lastRun = cfg?.webmonLastRun ?? 0
  const [busy, setBusy] = useState(false)
  const [showHist, setShowHist] = useState(false)

  const patch = (p: Partial<WebmonConfig>) => useBaseStore.getState().setNodeConfig(nodeId, p)
  const setTargets = (list: WebmonTarget[]) => patch({ webmonTargets: list })

  const addTarget = () => {
    setTargets([...targets, { id: 'wm-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5), url: 'https://', mode: 'text', selector: '' }])
  }
  const patchTarget = (id: string, p: Partial<WebmonTarget>) => setTargets(targets.map((t) => (t.id === id ? { ...t, ...p } : t)))
  const removeTarget = (id: string) => setTargets(targets.filter((t) => t.id !== id))

  const checkNow = async () => {
    if (!targets.length) { alert('先添加监控目标（URL）'); return }
    setBusy(true)
    try { await checkWebmonNow(nodeId) } finally { setBusy(false) }
  }

  const changedRecent = hist.filter((h) => h.changed).length

  return (
    <div className="pn-root" style={{ gap: 5 }}>
      {/* 摘要行：目标数 / 上次检查 / 变化数 */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '0 6px', fontSize: 11.5, color: 'var(--c-text2)' }}>
        <span>{targets.length} 个目标</span>
        <span>{lastRun ? '上次检查 ' + new Date(lastRun).toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' }) : '未检查'}</span>
        {changedRecent > 0 && <span style={{ color: 'var(--c-green, #22c55e)' }}>累计 {changedRecent} 变化</span>}
      </div>

      {/* 操作行：立即检查 / 导出 CSV / 添加目标 */}
      <div style={{ display: 'flex', gap: 6, padding: '0 6px' }}>
        <button
          className="pm-btn"
          style={{ flex: 1 }}
          disabled={busy || !targets.length}
          onClick={() => { void checkNow() }}
          title="立即检查全部目标（提取 → 与上次值比对 → 记录历史）"
        >
          {busy ? '检查中…' : '⟳ 立即检查'}
        </button>
        <button className="pm-btn" disabled={!hist.length} onClick={() => downloadWebmonCsv(hist, name)} title="历史导出 CSV（Excel 可直接打开）">导出</button>
        <button className="pm-btn" onClick={addTarget} title="添加监控目标">＋目标</button>
      </div>

      {/* 目标列表（可编辑） */}
      {targets.map((tg, i) => (
        <div key={tg.id} style={{ display: 'flex', flexDirection: 'column', gap: 3, padding: '5px 6px', background: 'var(--c-bg2, #8881)', borderRadius: 6 }}>
          <div style={{ display: 'flex', gap: 4 }}>
            <input
              value={tg.url}
              placeholder="https://…"
              onChange={(e) => patchTarget(tg.id, { url: e.target.value })}
              style={{ flex: 1, minWidth: 0, fontSize: 11.5, padding: '3px 6px', border: '1px solid var(--c-border)', borderRadius: 4, background: 'var(--c-bg1)', color: 'var(--c-text1)' }}
              title={tg.url}
            />
            <button className="pm-btn" style={{ padding: '2px 7px', color: 'var(--c-red, #ef4444)' }} onClick={() => removeTarget(tg.id)} title="删除此目标">✕</button>
          </div>
          <div style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
            <select
              value={tg.mode}
              onChange={(e) => patchTarget(tg.id, { mode: e.target.value as WebmonTarget['mode'] })}
              style={{ fontSize: 11, padding: '2px 4px', border: '1px solid var(--c-border)', borderRadius: 4, background: 'var(--c-bg1)', color: 'var(--c-text1)' }}
              title="提取模式"
            >
              <option value="text">网页文本</option>
              <option value="json">JSON 接口</option>
              <option value="keyword">关键词</option>
            </select>
            {tg.mode === 'text' && (
              <input
                value={tg.selector ?? ''}
                placeholder="CSS 选择器（如 .price、#title，空=全正文）"
                onChange={(e) => patchTarget(tg.id, { selector: e.target.value })}
                style={{ flex: 1, minWidth: 0, fontSize: 11, padding: '2px 6px', border: '1px solid var(--c-border)', borderRadius: 4, background: 'var(--c-bg1)', color: 'var(--c-text1)' }}
              />
            )}
            {tg.mode === 'json' && (
              <input
                value={tg.jsonPath ?? ''}
                placeholder="点路径（如 data.0.price）"
                onChange={(e) => patchTarget(tg.id, { jsonPath: e.target.value })}
                style={{ flex: 1, minWidth: 0, fontSize: 11, padding: '2px 6px', border: '1px solid var(--c-border)', borderRadius: 4, background: 'var(--c-bg1)', color: 'var(--c-text1)' }}
              />
            )}
            {tg.mode === 'keyword' && (
              <input
                value={tg.keyword ?? ''}
                placeholder="关键词（出现/未出现）"
                onChange={(e) => patchTarget(tg.id, { keyword: e.target.value })}
                style={{ flex: 1, minWidth: 0, fontSize: 11, padding: '2px 6px', border: '1px solid var(--c-border)', borderRadius: 4, background: 'var(--c-bg1)', color: 'var(--c-text1)' }}
              />
            )}
            <span style={{ fontSize: 10.5, color: 'var(--c-text3)', flexShrink: 0 }}>#{i + 1}</span>
          </div>
        </div>
      ))}

      {/* 历史表 */}
      {hist.length > 0 && (
        <div style={{ padding: '0 6px 4px' }}>
          <button className="pm-btn" style={{ width: '100%', fontSize: 11 }} onClick={() => setShowHist((v) => !v)}>
            {showHist ? '▾' : '▸'} 历史（{hist.length} 条）
          </button>
          {showHist && (
            <div style={{ marginTop: 4, maxHeight: 180, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 3 }}>
              {hist.slice(0, 12).map((h, i) => (
                <div key={i} style={{ display: 'flex', gap: 6, fontSize: 11, alignItems: 'baseline', color: h.error ? 'var(--c-red, #ef4444)' : h.changed ? 'var(--c-green, #22c55e)' : 'var(--c-text2)' }}>
                  <span style={{ flexShrink: 0, color: 'var(--c-text3)' }}>{new Date(h.at).toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' })}</span>
                  <span style={{ flexShrink: 0, maxWidth: 90, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={h.targetName}>{h.targetName}</span>
                  <span style={{ flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={h.error ?? h.value}>
                    {h.error ? '✕ ' + h.error : (h.changed ? '↻ ' : '') + h.value}
                  </span>
                </div>
              ))}
              {hist.length > 12 && <div style={{ fontSize: 10.5, color: 'var(--c-text3)' }}>…共 {hist.length} 条（导出看全部）</div>}
            </div>
          )}
        </div>
      )}
    </div>
  )
}
