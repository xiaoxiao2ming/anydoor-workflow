/**
 * 触发规则面板 —— 规则编辑 / 命中预览 / 手动触发
 *
 * 设计原则：
 *   · **先看后写**。参数变了先出「命中预览 + 会改哪几格」，用户点「立即更新」才真落盘。
 *     这就是用户要的"更新时更新"——不是每次都写，也不是要用户猜会不会写。
 *   · 跳过项必须显示。写不进去的格子（表不存在 / 是公式格）要如实报出来，
 *     否则用户会以为"规则命中了但文件没变"，进而怀疑整套机制。
 *   · 复用 trigger/apply 的纯逻辑，本组件只管收集输入与渲染，不含判定规则。
 */

import { useMemo, useState, useCallback, useRef } from 'react'
import {
  diffSnapshots,
  evalRules,
  planTrigger,
  DebounceGate,
  PRESET_RULES,
  type TriggerRule,
  type SourceChange,
  type CondOp,
} from './trigger'

const OP_LABEL: Record<CondOp, string> = {
  always: '无条件',
  changed: '值变化时',
  gt: '大于',
  gte: '大于等于',
  lt: '小于',
  lte: '小于等于',
  eq: '等于',
  ne: '不等于',
  between: '介于',
  outside: '区间外',
  nonEmpty: '非空时',
  empty: '为空时',
  isTrue: '为真时',
  isFalse: '为假时',
}

const ACTION_LABEL: Record<string, string> = {
  writeCell: '写单元格',
  writeRange: '写区域',
  recompute: '重算',
  commit: '落盘',
}

export interface TriggerPanelProps {
  /** 当前工作簿里有哪些表（给用户挑写入目标） */
  sheetNames: string[]
  /** 取来源当前值：URN → 值 */
  readSource: (urn: string) => unknown
  /** 供用户挑的候选来源（画布节点字段 / 全局参数） */
  sources: { urn: string; label: string }[]
  /** 规则集（受控） */
  rules: TriggerRule[]
  onRulesChange: (next: TriggerRule[]) => void
  /** 执行一份计划。返回给用户看的日志。 */
  onApply: (plan: ReturnType<typeof planTrigger>) => Promise<void> | void
  /** 干跑预览（不算真实副作用），用于「会改哪几格」 */
  onDryRun?: (plan: ReturnType<typeof planTrigger>) => { written: number; skipped: { target: string; reason: string }[] } | void
}

let seq = 0
const newId = () => 'tr' + Date.now().toString(36) + (seq++).toString(36)

export function TriggerPanel({
  sheetNames,
  readSource,
  sources,
  rules,
  onRulesChange,
  onApply,
  onDryRun,
}: TriggerPanelProps) {
  const [snapshot, setSnapshot] = useState<Record<string, unknown>>({})
  const [lastPlan, setLastPlan] = useState<ReturnType<typeof planTrigger> | null>(null)
  const [lastSkipped, setLastSkipped] = useState<{ target: string; reason: string }[]>([])
  const [busy, setBusy] = useState(false)
  const gate = useRef(new DebounceGate())

  /* ── 拉一次全部来源的当前值（规则 watch 的来源可能不止一个） ── */
  const readAll = useCallback((): Record<string, unknown> => {
    const out: Record<string, unknown> = {}
    for (const s of sources) {
      try {
        out[s.urn] = readSource(s.urn)
      } catch {
        out[s.urn] = undefined
      }
    }
    return out
  }, [sources, readSource])

  /** 快照 vs 上次快照的差分 = 本次变更 */
  const changes: SourceChange[] = useMemo(() => {
    const { changes: cs } = diffSnapshots(snapshot, readAll())
    return cs
  }, [snapshot, readAll])

  const plan = useMemo(
    () => planTrigger(rules, changes, readSource),
    [rules, changes, readSource]
  )

  const evalStats = useMemo(() => evalRules(rules, changes), [rules, changes])

  /* ── 改规则 ── */
  const patch = (id: string, p: Partial<TriggerRule>) =>
    onRulesChange(rules.map((r) => (r.id === id ? { ...r, ...p } : r)))

  const addRule = (preset?: Omit<TriggerRule, 'id'>) => {
    onRulesChange([...rules, { ...(preset ?? PRESET_RULES[0]), id: newId() }])
  }

  const removeRule = (id: string) => onRulesChange(rules.filter((r) => r.id !== id))

  /* ── 预览：只算不写 ── */
  const preview = () => {
    const p = planTrigger(rules, changes, readSource)
    setLastPlan(p)
    setLastSkipped(onDryRun?.(p)?.skipped ?? [])
  }

  /** 立即更新：先清去抖闸门，保证必过 */
  const runNow = async () => {
    const p = planTrigger(rules, changes, readSource)
    if (!p.hits.length) {
      setLastPlan(p)
      return
    }
    // 去抖：命中 commit 类规则且配了窗口时才拦
    const commitRule = p.hits.find((h) => h.rule.action.kind === 'commit')?.rule
    if (commitRule?.debounceMs && !gate.current.allow('commit', commitRule.debounceMs)) {
      setLastPlan(p)
      return
    }
    setBusy(true)
    try {
      await onApply(p)
      setLastPlan(p)
      setLastSkipped([])
      // 落盘后把当前值设为新基线，避免同一批变更被反复触发
      setSnapshot(readAll())
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="trg-panel">
      <div className="trg-head">
        <b>联动触发规则</b>
        <span className="trg-sub">参数改了 → 命中规则才更新</span>
      </div>

      {/* 本次变更 */}
      <div className="trg-sec">
        <div className="trg-sec-t">当前变更</div>
        {changes.length === 0 ? (
          <div className="trg-empty">外部参数无变化 —— 不会触发任何更新</div>
        ) : (
          <ul className="trg-changes">
            {changes.map((c) => (
              <li key={c.urn}>
                <code>{c.urn}</code>: {fmt(c.before)} → <b>{fmt(c.after)}</b>
              </li>
            ))}
          </ul>
        )}
      </div>

      {/* 规则表 */}
      <div className="trg-sec">
        <div className="trg-sec-t">
          规则（{rules.length}）
          <button className="trg-mini" onClick={() => addRule()}>
            + 规则
          </button>
          <button className="trg-mini" onClick={() => addRule(PRESET_RULES[0])}>
            + 预设
          </button>
        </div>
        {rules.length === 0 ? (
          <div className="trg-empty">还没有规则。参数改了也不会写进表格 —— 点「+ 预设」快速起一条。</div>
        ) : (
          <ul className="trg-rules">
            {rules.map((r) => (
              <RuleRow
                key={r.id}
                rule={r}
                sheetNames={sheetNames}
                sources={sources}
                hits={evalStats.hits.filter((h) => h.rule.id === r.id).length}
                onPatch={(p) => patch(r.id, p)}
                onRemove={() => removeRule(r.id)}
              />
            ))}
          </ul>
        )}
      </div>

      {/* 操作 */}
      <div className="trg-acts">
        <button className="trg-btn" onClick={preview} disabled={!changes.length}>
          预览会改哪几格
        </button>
        <button className="trg-btn trg-btn-primary" onClick={runNow} disabled={busy || !plan.hits.length}>
          {busy ? '更新中…' : `立即更新（命中 ${plan.hits.length} 条）`}
        </button>
        <button
          className="trg-mini"
          onClick={() => {
            setSnapshot(readAll())
            setLastPlan(null)
            setLastSkipped([])
          }}
        >
          重设基线
        </button>
      </div>

      {/* 预览 / 执行结果 */}
      {lastPlan && (
        <div className="trg-sec">
          <div className="trg-sec-t">本次会改（预览）</div>
          {lastPlan.log.length === 0 ? (
            <div className="trg-empty">没有规则命中</div>
          ) : (
            <ul className="trg-log">
              {lastPlan.log.map((l, i) => (
                <li key={i}>{l}</li>
              ))}
            </ul>
          )}
          {lastSkipped.length > 0 && (
            <>
              <div className="trg-sec-t trg-warn-t">跳过</div>
              <ul className="trg-log trg-warn">
                {lastSkipped.map((s, i) => (
                  <li key={i}>
                    {s.target}：{s.reason}
                  </li>
                ))}
              </ul>
            </>
          )}
        </div>
      )}

      {evalStats.skips.length > 0 && (
        <div className="trg-sec">
          <div className="trg-sec-t">未命中</div>
          <ul className="trg-log trg-dim">
            {evalStats.skips.map((s, i) => (
              <li key={i}>
                {s.rule.name}：{s.reason}
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  )
}

/* ── 单条规则行 ── */

function RuleRow({
  rule,
  sheetNames,
  sources,
  hits,
  onPatch,
  onRemove,
}: {
  rule: TriggerRule
  sheetNames: string[]
  sources: { urn: string; label: string }[]
  hits: number
  onPatch: (p: Partial<TriggerRule>) => void
  onRemove: () => void
}) {
  const a = rule.action
  const needSheet = a.kind === 'writeCell' || a.kind === 'writeRange' || a.kind === 'recompute'
  const needRange = a.kind === 'writeRange'

  return (
    <li className={rule.enabled ? 'trg-rule' : 'trg-rule trg-off'}>
      <div className="trg-rule-top">
        <input
          type="checkbox"
          checked={rule.enabled}
          onChange={(e) => onPatch({ enabled: e.target.checked })}
          title={rule.enabled ? '停用这条规则' : '启用这条规则'}
        />
        <input
          className="trg-name"
          value={rule.name}
          onChange={(e) => onPatch({ name: e.target.value })}
          placeholder="规则名"
        />
        {hits > 0 && <span className="trg-badge">命中 {hits}</span>}
        <button className="trg-x" onClick={onRemove} title="删除规则">
          ×
        </button>
      </div>

      <div className="trg-rule-cond">
        <span className="trg-k">当</span>
        <select value={rule.watch} onChange={(e) => onPatch({ watch: e.target.value })}>
          <option value="change">任何来源变化</option>
          {sources.map((s) => (
            <option key={s.urn} value={s.urn}>
              {s.label}（{s.urn}）
            </option>
          ))}
        </select>
        <select
          value={rule.condition.op}
          onChange={(e) => onPatch({ condition: { ...rule.condition, op: e.target.value as CondOp } })}
        >
          {Object.entries(OP_LABEL).map(([k, v]) => (
            <option key={k} value={k}>
              {v}
            </option>
          ))}
        </select>
        {needsThreshold(rule.condition.op) && (
          <>
            <input
              className="trg-num"
              value={String(rule.condition.value ?? '')}
              onChange={(e) => onPatch({ condition: { ...rule.condition, value: num(e.target.value) } })}
              placeholder="值"
            />
            {rule.condition.op === 'between' || rule.condition.op === 'outside' ? (
              <>
                <span className="trg-k">~</span>
                <input
                  className="trg-num"
                  value={String(rule.condition.hi ?? '')}
                  onChange={(e) => onPatch({ condition: { ...rule.condition, hi: num(e.target.value) } })}
                  placeholder="上界"
                />
              </>
            ) : null}
          </>
        )}
      </div>

      <div className="trg-rule-act">
        <span className="trg-k">则</span>
        <select
          value={a.kind}
          onChange={(e) => {
            const kind = e.target.value as typeof a.kind
            onPatch({
              action:
                kind === 'writeCell'
                  ? { kind, sheet: sheetNames[0] ?? 'Sheet1', coord: 'A1' }
                  : kind === 'writeRange'
                    ? { kind, sheet: sheetNames[0] ?? 'Sheet1', ref: 'A1:A1' }
                    : kind === 'recompute'
                      ? { kind }
                      : { kind: 'commit' },
            })
          }}
        >
          {Object.entries(ACTION_LABEL).map(([k, v]) => (
            <option key={k} value={k}>
              {v}
            </option>
          ))}
        </select>

        {needSheet && 'sheet' in a && (
          <select
            value={a.sheet ?? ''}
            onChange={(e) => onPatch({ action: { ...a, sheet: e.target.value } as TriggerRule['action'] })}
          >
            {sheetNames.map((s) => (
              <option key={s} value={s}>
                {s}
              </option>
            ))}
          </select>
        )}

        {a.kind === 'writeCell' && (
          <input
            className="trg-coord"
            value={a.coord}
            onChange={(e) => onPatch({ action: { ...a, coord: e.target.value } })}
            placeholder="A1"
          />
        )}
        {needRange && 'ref' in a && (
          <input
            className="trg-coord"
            value={a.ref}
            onChange={(e) => onPatch({ action: { ...a, ref: e.target.value } as TriggerRule['action'] })}
            placeholder="A1:C10"
          />
        )}

        <input
          className="trg-num"
          value={String(rule.priority ?? 100)}
          onChange={(e) => onPatch({ priority: num(e.target.value) ?? 100 })}
          title="优先级，小的先执行"
          placeholder="优先级"
        />
        <input
          className="trg-num"
          value={String(rule.debounceMs ?? 0)}
          onChange={(e) => onPatch({ debounceMs: num(e.target.value) ?? 0 })}
          title="去抖窗口（毫秒）"
          placeholder="去抖ms"
        />
      </div>

      {rule.note && <div className="trg-note">{rule.note}</div>}
    </li>
  )
}

function needsThreshold(op: CondOp): boolean {
  return ['gt', 'gte', 'lt', 'lte', 'eq', 'ne', 'between', 'outside'].includes(op)
}

function num(s: string): number | undefined {
  const t = s.trim()
  if (t === '') return undefined
  const n = Number(t)
  return Number.isFinite(n) ? n : undefined
}

function fmt(v: unknown): string {
  if (v == null) return '（空）'
  if (typeof v === 'object') {
    try {
      const s = JSON.stringify(v)
      return s.length > 30 ? s.slice(0, 30) + '…' : s
    } catch {
      return '（对象）'
    }
  }
  return String(v)
}
