import { useMemo } from 'react'
import type { TableSchema, NodeRuleSet } from './core/schema'
import { KIND_STYLE } from './core/nodeStyle'
import {
  bindingsOfRuleNode,
  ruleFieldsOf,
  ruleNodesOf,
  validateRule,
  RULE_FIELDS_OF_KIND,
  COMMON_RULE_FIELDS,
} from './core/ruleSync'

/**
 * 规则编辑区（2026-10-05，圆形「规则节点」）
 *
 * 用户口径：
 *   「一旦匹配了，默认不能每个改规则，只能改主流程的节点进行修改。
 *     修改后所有相同规则的节点规则同步改。」
 *   「圆形节点是对连线的节点进行规则锁定（统一）」
 *
 * ── 三态设计（同一个组件承担，靠 isRule / locked 分支）──
 *   ① 规则节点本身 → 可编辑，按被锁节点的 kind 分桶
 *   ② 被锁的辅节点 → **只读**，明确写「由规则节点统一管理」，不给任何输入框
 *   ③ 普通节点     → 若已被锁，只显示一条提示（不出编辑区）
 *
 * ★ 为什么辅节点完全不给输入框（而灰掉但可改）：
 *   灰掉的可编辑控件一旦被绕过（脚本、快捷键），就出现「改了没生效」的诡异状态。
 *   直接不渲染 = 绕过也没地方改，与「规则只存一份」的结构一致。
 */
export function RuleEditor({
  table,
  tables,
  onSetRule,
  onClearRule,
  onUnlink,
}: {
  table: TableSchema
  tables: Record<string, TableSchema>
  onSetRule: (ruleNodeId: string, kind: string, patch: Partial<NodeRuleSet>) => void
  onClearRule: (ruleNodeId: string, kind: string) => void
  onUnlink: (ruleNodeId: string, targetId: string) => void
}) {
  const isRule = table.kind === 'rule'
  const bindings = useMemo(
    () => (isRule ? bindingsOfRuleNode(table.id, tables) : []),
    [isRule, table.id, tables],
  )
  const lockedBy = useMemo(
    () => (isRule ? [] : ruleNodesOf(table.id, tables)),
    [isRule, table.id, tables],
  )

  /*★★ hooks 必须全部在所有条件返回之前声明（Rules of Hooks）。

     症状（2026-10-05 用户报「规则跟随节点点击一下就空了」）：
     选中规则节点 → 右侧属性栏挂上本组件 → 若它此前渲染的是
     「② 辅节点只读」或「①但buckets 为空」这两条早退路径，
     切到完整路径时才会执行下面那个 useMemo —— hook 数从 2 变 3，
     React 抛「Rendered more hooks than during the previous render」，
     **整棵画布树卸载，规则节点连同计数一起消失**（用户看到的「空了」）。

     ★ 为什么之前没炸：这些分支要靠「选中不同类型的节点」才切换得到，
       单测（只调内核纯函数）与常规浏览路径都不一定走到，
       597 个用例全绿也照样漏。*/
  const buckets = useMemo(() => {
    const m = new Map<string, string[]>() // kind → 被锁节点名
    for (const b of bindings) {
      const arr = m.get(b.targetKind) ?? []
      arr.push(tables[b.targetId]?.name ?? b.targetId)
      m.set(b.targetKind, arr)
    }
    return [...m.entries()]
  }, [bindings, tables])

  /* ── ② 辅节点：只读提示 ── */
  if (!isRule) {
    if (!lockedBy.length) return null
    return (
      <div className="props-field rule-panel rule-panel--locked">
        <label>规则由圆形节点统一管理</label>
        {lockedBy.map((rid) => (
          <div key={rid} className="rule-locked-row">
            <span className="rule-locked-name">{tables[rid]?.name ?? rid}</span>
            <button
              type="button"
              className="rule-btn rule-btn--danger"
              title="解除锁定：断开连线后本节点回用自己的设置"
              onClick={() => onUnlink(rid, table.id)}
            >
              解除锁定
            </button>
          </div>
        ))}
        <p className="hint">
          要改规则，请到<b>圆形规则节点</b>上改 —— 所有被它锁定的同类节点会同步更新。
          这样规则只有一份，不会出现「改了这边忘了那边」。
        </p>
      </div>
    )
  }

  /* ── ① 规则节点：可编辑 ── */
  if (!buckets.length) {
    return (
      <div className="props-field rule-panel">
        <label>规则锁定</label>
        <p className="hint">
          还没锁定任何节点。<b>把别的节点连到本节点上</b>，就能对它下发规则。
          连线即锁定范围 —— 连着的节点会跟着变，没连的不受影响。
        </p>
      </div>
    )
  }

  return (
    <div className="props-field rule-panel">
      <label>规则锁定（{bindings.length} 个节点被统一管理）</label>
      <p className="hint" style={{ marginBottom: 8 }}>
        在这里改，下面列出的<b>全部同类节点立即同步</b>。
        规则只存在本节点上，被锁节点不存副本，所以不存在「改了一处忘了另一处」。
      </p>
      {buckets.map(([kind, names]) => (
        <RuleBucketEditor
          key={kind}
          ruleNodeId={table.id}
          kind={kind}
          targetNames={names}
          tables={tables}
          onSetRule={onSetRule}
          onClearRule={onClearRule}
        />
      ))}
    </div>
  )
}

/** 单个 kind 的规则桶 —— 只渲染该 kind 适配的规则项 */
function RuleBucketEditor({
  ruleNodeId,
  kind,
  targetNames,
  tables,
  onSetRule,
  onClearRule,
}: {
  ruleNodeId: string
  kind: string
  targetNames: string[]
  tables: Record<string, TableSchema>
  onSetRule: (ruleNodeId: string, kind: string, patch: Partial<NodeRuleSet>) => void
  onClearRule: (ruleNodeId: string, kind: string) => void
}) {
  const ruleNode = tables[ruleNodeId]
  const rule: NodeRuleSet = ruleNode?.config?.rules?.[kind] ?? {}
  const problems = validateRule(rule, kind)
  const label = KIND_STYLE[kind]?.label ?? kind
  const fields = ruleFieldsOf(kind)
  const behaviors = RULE_FIELDS_OF_KIND[kind] ?? []
  const has = (f: keyof NodeRuleSet) => rule[f] !== undefined && rule[f] !== ''

  return (
    <div className="rule-bucket">
      <div className="rule-bucket-head">
        <span className={`rf-badge ${KIND_STYLE[kind]?.family ? `rf-badge--${KIND_STYLE[kind].family}` : ''}`}>{label}</span>
        <span className="rule-bucket-names">{targetNames.join('、')}</span>
        {Object.keys(rule).length > 0 && (
          <button type="button" className="rule-btn" onClick={() => onClearRule(ruleNodeId, kind)} title="清空该类型的规则，被锁节点回用自己的设置">
            清空
          </button>
        )}
      </div>

      {/* 视觉规则：所有 kind 通用 */}
      <div className="rule-grid">
        {COMMON_RULE_FIELDS.includes('family') && (
          <label className="rule-item">
            <span>统一配色族</span>
            <select value={rule.family ?? ''} onChange={(e) => onSetRule(ruleNodeId, kind, { family: e.target.value || undefined })}>
              <option value="">（各用各的）</option>
              {Object.keys(KIND_STYLE)
                .filter((k) => KIND_STYLE[k].family && k !== 'rule')
                .map((k) => (
                  <option key={k} value={KIND_STYLE[k].family}>
                    {KIND_STYLE[k].family}
                  </option>
                ))}
            </select>
          </label>
        )}
        {COMMON_RULE_FIELDS.includes('scale') && (
          <label className="rule-item">
            <span>缩放倍率</span>
            <input
              type="number"
              step={0.1}
              min={0.5}
              max={2}
              value={rule.scale ?? 1}
              onChange={(e) => onSetRule(ruleNodeId, kind, { scale: Number(e.target.value) })}
            />
          </label>
        )}
      </div>

      {/* 行为规则：按 kind 适配（用户：「针对不同节点的内容适配不同的规则样式体系」）*/}
      {behaviors.includes('sortBy') && (
        <label className="rule-item">
          <span>统一排序列</span>
          <input
            value={rule.sortBy ?? ''}
            placeholder="字段 id"
            onChange={(e) => onSetRule(ruleNodeId, kind, { sortBy: e.target.value || undefined })}
          />
        </label>
      )}
      {behaviors.includes('sortDir') && (
        <label className="rule-item">
          <span>统一升降序</span>
          <select value={rule.sortDir ?? ''} onChange={(e) => onSetRule(ruleNodeId, kind, { sortDir: (e.target.value || undefined) as 'asc' | 'desc' | undefined })}>
            <option value="">（各自为准）</option>
            <option value="asc">升序</option>
            <option value="desc">降序</option>
          </select>
        </label>
      )}
      {behaviors.includes('exportFormat') && (
        <label className="rule-item">
          <span>统一导出格式</span>
          <select
            value={rule.exportFormat ?? ''}
            onChange={(e) => onSetRule(ruleNodeId, kind, { exportFormat: (e.target.value || undefined) as NodeRuleSet['exportFormat'] })}
          >
            <option value="">（各自为准）</option>
            <option value="csv">CSV</option>
            <option value="json">JSON</option>
            <option value="xlsx">Excel</option>
            <option value="txt">TXT</option>
            <option value="md">Markdown</option>
            <option value="pdf">PDF</option>
          </select>
        </label>
      )}
      {behaviors.includes('mapKeyField') && (
        <label className="rule-item">
          <span>统一用上游哪个字段去查</span>
          <input
            value={rule.mapKeyField ?? ''}
            placeholder="字段 id"
            onChange={(e) => onSetRule(ruleNodeId, kind, { mapKeyField: e.target.value || undefined })}
          />
        </label>
      )}
      {behaviors.includes('mapMatchField') && (
        <label className="rule-item">
          <span>统一用本表哪个字段匹配</span>
          <input
            value={rule.mapMatchField ?? ''}
            placeholder="字段 id"
            onChange={(e) => onSetRule(ruleNodeId, kind, { mapMatchField: e.target.value || undefined })}
          />
        </label>
      )}

      {/* 校验提示：缺项/越界标黄而不是静默接受 */}
      {problems.length > 0 && (
        <div className="rule-problems">
          {problems.map((p) => (
            <div key={p.field} className="rule-problem">
              ⚠ {p.msg}
            </div>
          ))}
        </div>
      )}
      {fields.length === 0 && <p className="hint">该类型暂无可适配的规则项。</p>}
    </div>
  )
}