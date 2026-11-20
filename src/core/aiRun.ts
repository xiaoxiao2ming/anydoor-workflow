/**
 * AI 节点执行管线 · 纯逻辑层（2026-10-06 用户：「先做一个输入文字输出文字的样例」）。
 * 只做「组装」：上游内容 → 变量表 → 插值 → chat messages。
 * 真正的 fetch 在 PropertiesPanel 组件里（async UI 动作），本文件保持纯同步可测。
 * 依赖方向：aiRun → engine/schema/store(类型)，无环。
 */
import { computeNode } from './engine'
import type { NodeConfig, TableSchema } from './schema'
import type { Row } from './store'

/**
 * 组装某个上游节点的内容文本。
 * - text 节点 = 原文直出（多行合并）——「输入文字输出文字」样例的主链路；
 * - 其余节点 = computeNode 后 JSON 兜底（通用可跑，格式化交给后续按类型精修）。
 */
export function aiUpstreamContentOf(
  nodeId: string,
  tables: Record<string, TableSchema>,
  records: Record<string, Row[]>,
): string {
  const t = tables[nodeId]
  if (!t) return ''
  if (t.kind === 'text') {
    return (records[nodeId] ?? []).map((r) => String(r.text ?? '')).join('\n')
  }
  /* ★ 网址监控上游（2026-11-09）：变量值 = 各目标最新提取值——监控变化喂给 AI 分析 */
  if (t.kind === 'webmon') {
    const cfg = (t.config ?? {}) as { webmonTargets?: { id: string; url: string; name?: string }[]; webmonLastValues?: Record<string, string>; webmonHistory?: { changed: boolean }[] }
    const targets = cfg.webmonTargets ?? []
    if (!targets.length) return ''
    const lines = targets.map((tg) => `${tg.name || tg.url}：${cfg.webmonLastValues?.[tg.id] ?? '（未检查）'}`)
    const changedNote = (cfg.webmonHistory ?? []).some((h) => h.changed) ? '\n（本节点近期检测到内容变化）' : ''
    return lines.join('\n') + changedNote
  }
  const { rows } = computeNode(nodeId, tables, records)
  if (!rows.length) return ''
  return JSON.stringify(rows, null, 2)
}

/**
 * {{变量名}} 插值：命中替换，未命中原样保留——
 * 用户可能故意在提示词里写花括号字面量，不能吞掉。
 * 变量名允许两侧空白：`{{ 内容 }}` 与 `{{内容}}` 等价。
 */
export function aiInterpolate(tpl: string, vars: Record<string, string>): string {
  return tpl.replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (m, name: string) => (name in vars ? vars[name] : m))
}

/**
 * 变量名 → 上游内容。命名口径与 PropertiesPanel.varNameOf 同源：
 * aiInputVars 配置名优先，未起名用上游节点名兜底。
 */
export function aiVarMapOf(
  config: NodeConfig,
  tables: Record<string, TableSchema>,
  records: Record<string, Row[]>,
): Record<string, string> {
  const map: Record<string, string> = {}
  for (const nid of config.linkedInputs ?? []) {
    if (!tables[nid]) continue
    const named = (config.aiInputVars ?? []).find((x) => x.from === nid)
    const name = named?.name ?? tables[nid].name
    map[name] = aiUpstreamContentOf(nid, tables, records)
  }
  return map
}

export type AiMessage = { role: 'system' | 'user'; content: string }

/** ★ AI 节点执行（2026-11-09 从 PropertiesPanel 抽到 core）：面板按钮与 webmon 变化自动触发共用。
 *  拉上游 → 组装 messages → fetch OpenAI 兼容 chat/completions → aiLastOutput 直写 store。
 *  「执行不是编辑」，不走 pushHistory（与面板原实现同语义）。 */
export async function runAiNode(
  nodeId: string,
  st: {
    tables: Record<string, TableSchema>
    records: Record<string, Row[]>
    setNodeConfig: (id: string, patch: Partial<NodeConfig>) => void
  },
): Promise<{ ok: boolean; msg: string }> {
  const t = st.tables[nodeId]
  if (!t || t.deleted) return { ok: false, msg: '节点不存在（可能已被删除）' }
  const config = (t.config ?? {}) as NodeConfig
  const base = (config.aiApiBase ?? '').trim().replace(/\/+$/, '')
  const key = (config.aiApiKey ?? '').trim()
  const model = (config.aiModel ?? '').trim()
  const messages = aiBuildMessages(config, st.tables, st.records)

  /* ★ 2026-11-09 本地模式（用户「全部跑通」）：API 地址留空 = 走本地大模型（node-llama-cpp，
     9B/3B 会话池）——零配置即可分析，数据不出电脑。外部 API 照旧填 base+key+model。 */
  if (!base) {
    const api = (globalThis as { api?: { llmChat?: (p: string, o?: { noThink?: boolean; maxTokens?: number; temperature?: number }) => Promise<{ ok: boolean; text?: string; reason?: string }> } }).api
    if (!api?.llmChat) { st.setNodeConfig(nodeId, { aiLastError: '本地模型运行时不可用（缺少推理模块），或在属性面板填外部 API 设置' }); return { ok: false, msg: '本地模型不可用' } }
    const prompt = messages.map((m) => `${m.role === 'system' ? '[指令]' : '[数据]'}${m.content}`).join('\n\n')
    st.setNodeConfig(nodeId, { aiLastError: undefined })
    const r = await api.llmChat(prompt, { noThink: true, maxTokens: 1024, temperature: 0.3 })
    if (!r.ok) { st.setNodeConfig(nodeId, { aiLastError: r.reason ?? '本地推理失败' }); return { ok: false, msg: r.reason ?? '本地推理失败' } }
    const text = (r.text ?? '').trim()
    st.setNodeConfig(nodeId, { aiLastOutput: text || '（本地模型没有返回内容）' })
    return { ok: true, msg: text.slice(0, 120) || '（空）' }
  }

  if (!key || !model) {
    st.setNodeConfig(nodeId, { aiLastError: '填了接口地址后还需 Key 与模型名（或清空接口地址改用本地模型）' })
    return { ok: false, msg: '缺 Key / 模型名' }
  }
  try {
    const resp = await fetch(`${base}/chat/completions`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${key}` },
      body: JSON.stringify({
        model,
        messages,
        temperature: config.aiTemperature ?? 0.7,
        max_tokens: config.aiMaxTokens ?? 2048,
      }),
    })
    if (!resp.ok) {
      const body = await resp.text().catch(() => '')
      throw new Error(`HTTP ${resp.status}${body ? ' ' + body.slice(0, 200) : ''}`)
    }
    const data = (await resp.json()) as { choices?: { message?: { content?: unknown } }[] }
    const out = data?.choices?.[0]?.message?.content
    if (typeof out !== 'string') throw new Error('响应缺 choices[0].message.content：' + JSON.stringify(data).slice(0, 200))
    st.setNodeConfig(nodeId, { aiLastOutput: out, aiLastError: undefined })
    return { ok: true, msg: out.slice(0, 120) }
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e)
    st.setNodeConfig(nodeId, { aiLastError: msg })
    return { ok: false, msg }
  }
}

/**
 * 组装 chat messages（OpenAI 兼容）：
 * - system = 系统提示词 + 输出说明（aiOutputNote 拼成「输出要求：…」）；
 * - user   = 用户提示词插值 {{变量}}；模板为空时用全部上游内容兜底
 *   （「输入文字输出文字」零配置闭环的关键——连完线不填提示词也能跑）。
 */
export function aiBuildMessages(
  config: NodeConfig,
  tables: Record<string, TableSchema>,
  records: Record<string, Row[]>,
): AiMessage[] {
  const msgs: AiMessage[] = []
  const sys = [
    config.aiSystemPrompt ?? '',
    config.aiOutputNote ? `输出要求：${config.aiOutputNote}` : '',
  ]
    .filter((s) => s.trim() !== '')
    .join('\n')
  if (sys) msgs.push({ role: 'system', content: sys })
  const vars = aiVarMapOf(config, tables, records)
  const tpl = config.aiUserPrompt ?? ''
  const user = tpl.trim() !== '' ? aiInterpolate(tpl, vars) : Object.values(vars).join('\n')
  msgs.push({ role: 'user', content: user })
  return msgs
}
