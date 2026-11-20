import { useEffect, useMemo, useState } from 'react'
import { defineField, AI_SKILLS } from './core/schema'
import { aiBuildMessages, runAiNode } from './core/aiRun'
import type { TableSchema, FieldDef, FieldType, LookupSpec, NodeKind, NodeConfig, NodeRuleSet } from './core/schema'
import { isMergeKind, isWorkbookNodeKind, derivesFromOf, lineageChainOf, KIND_STYLE } from './core/nodeStyle'
import { RuleEditor } from './RuleEditor'
import { computeNode } from './core/engine'
import { useBaseStore } from './core/store'
import { EXPORT_FORMATS } from './core/io'
import { buildRefMark, offsetAfterInsert, plainLength } from './ref/inline'
import { WordReportPanel } from './docx/WordReportPanel'
import { XlsxReportPanel } from './xlsx/XlsxReportPanel'
import type { Row } from './core/store'

const TYPES: FieldType[] = ['text', 'number', 'formula', 'lookup']

const KIND_LABEL: Record<NodeKind, string> = {
  table: '表（可手写数据）',
  param: '数据参数表（导入外部数据）',
  test_table: '数据表节点（手动输入参数+参数数据；双击节点编辑），供 OO Excel 引用调试',
  sort: '排序（重排上游数据）',
  output: '输出（导出为文件）',
  start: '开始（工作流触发起点）',
  map: '数据对照表（上游查表补列）',
  docx: 'Word 文档（模板域绑定）',
  xlsx: 'Excel 工作簿（命名区域 / 表格绑定）',
  text: '文本（一个值：数字或一串文字）',
ooword: 'OO 文档',
webmon: '网址监控（轮询网页/接口，提取数据记录变化）',
xlsx2dxf: 'Excel转DXF（坐标表展点成图，生成 DXF 文件）',
        oosheet: 'OO 工作簿',
  chapter: 'OO 材料（数据处理 · 章节分级 + 目录大纲）',
  rule: '规则跟随（给连上的同类节点统一下发样式/行为）',
  ai: 'AI 节点（圆形 · 普通节点下1出线连它，接收数据做 AI 处理）',
  paramlib: '参数库（圆形 · 出线连普通节点上1口，为它供参数）',
'group': '节点组',
'note': '便利贴',
  folder: '文件夹（读取位置/保存目录）',
'pdfmerge': 'PDF生成（上游文档/图片按序生成一份 PDF）',
'photo': '照片（承载照片组，连入 PDF 合并——每图一页）',
  'pdf': 'PDF（承载 PDF 文件组，连入 PDF 合并——按行序拼接）',
  'pdfreplace': 'PDF替换页（上面=被替换/下面=替换页，执行后产物入历史）',
  'pdfslim': 'PDF瘦身（承载 PDF 文件组，压缩体积产物在节点上操作）',
  'fileout': '文件输出（连线上游产物，一键存储到固定目录）',
  'pdf2word': 'PDF转word（百度 PaddleOCR 引擎还原可编辑 docx，产物在节点历史）',
}

export function PropertiesPanel({
  table,
  tables,
  records,
  rows,
  onRename,
  onAddField,
  onUpdateField,
  onRemoveField,
  onSetNodeConfig,
  onSetRule,
  onClearRule,
  onUnlinkRuleNode,
  onImport,
  onExport,
  onAddRow,
  onSetCell,
  onRemoveRow,
}: {
  table: TableSchema | undefined
  tables: Record<string, TableSchema>
  records: Record<string, Row[]>
  /** 当前节点自身的记录行（对照表用：展示/编辑对照规则） */
  rows: Row[]
  onRename: (tableId: string, name: string) => void
  onAddField: (tableId: string, field: FieldDef) => void
  onUpdateField: (tableId: string, fieldId: string, patch: Partial<FieldDef>) => void
  onRemoveField: (tableId: string, fieldId: string) => void
  onSetNodeConfig: (tableId: string, patch: Partial<NodeConfig>) => void
  /** 规则节点：改某 kind 的规则（全部被锁节点同步生效） */
  onSetRule: (ruleNodeId: string, kind: string, patch: Partial<NodeRuleSet>) => void
  /** 规则节点：清掉某 kind 的规则 */
  onClearRule: (ruleNodeId: string, kind: string) => void
  /** 解除锁定：断开规则节点与目标节点的连线 */
  onUnlinkRuleNode: (ruleNodeId: string, targetId: string) => void
  onImport: (file: File) => void
  onExport: () => void
  onAddRow: (tableId: string) => void
  onSetCell: (tableId: string, rowIdx: number, fieldId: string, value: unknown) => void
  onRemoveRow: (tableId: string, rowIdx: number) => void
}) {
  // 上游汇入的字段并集（供排序/输出节点选择排序列、展示来源）
  // 注意：Hook 必须无条件调用。原先 `if (!table) return` 写在 useMemo 之前，
  // 会在「空态（0 个 Hook）↔ 属性态（1 个 Hook）」切换时破坏 Hook 顺序，
  // 让 React 反复报错/重建整棵子树，是属性栏抖动的加重因素。
  const inputFields = useMemo(() => {
    const fields: FieldDef[] = []
    if (!table) return fields
    const seen = new Set<string>()
    for (const id of table.config?.inputs ?? []) {
      const { fields: f } = computeNode(id, tables, records)
      for (const ff of f) if (!seen.has(ff.id)) {
        seen.add(ff.id)
        fields.push(ff)
      }
    }
    return fields
  }, [table, tables, records])

  if (!table) {
    return <div className="props-empty">选中一个节点以编辑其配置</div>
  }

  const kind = table.kind
  /** 血缘声明（衍生自哪个节点 + 做了哪些适配）；原始节点为 null */
  const lineage = derivesFromOf(kind)
  const config = table.config ?? {}

  return (
    <div className="props-body">
      <div className="props-field">
        <label>节点名</label>
        <input value={table.name} onChange={(e) => onRename(table.id, e.target.value)} />
      </div>
      <div className="props-field">
        <label>类型（不可改）</label>
        <input value={KIND_LABEL[kind]} disabled />
      </div>
      <div className="props-divider" />

      {/* ══ 文件族节点标准设置（2026-10-26 用户指令「看 PDF节点设置形式 先做标准版本」）══
          统一格式：统计区（只读清单+内外标）+ 设置区（仅 config 可序列化项）。
          编辑动作（增删/排序/预览/导入）留在节点卡片——单一真相源，面板不重复编辑入口。
          容差等数值设置面板与卡片双入口同源（同写 config）。 */}
      {(kind === 'pdf' || kind === 'pdfslim') && (() => {
        const items = ((config as { pdfItems?: { id: string; name: string; inDb?: boolean; b64?: string }[] }).pdfItems ?? [])
        const nIn = items.filter((it) => it.inDb || it.b64).length
        return (
          <>
            <div className="props-field">
              <label>文件组（{items.length} 份 · 内 {nIn} / 外 {items.length - nIn}）</label>
              <div className="props-filelist">
                {items.map((it) => (
                  <div key={it.id} className="props-file-row" title={it.name}>
                    <span className={'pn-loc ' + ((it.inDb || it.b64) ? 'pn-loc--in' : 'pn-loc--ext')}>{(it.inDb || it.b64) ? '内' : '外'}</span>
                    <span className="props-file-nm">{it.name}</span>
                  </div>
                ))}
                {!items.length && <span className="hint">空——在节点卡片点「导入」或把 PDF 拖进节点</span>}
              </div>
              <p className="hint">{kind === 'pdfslim'
                ? '瘦身节点：文件增删在节点卡片操作，压缩输出即将上线。'
                : '文件增删/排序/预览在节点卡片操作（本面板只读）。'}</p>
            </div>
          </>
        )
      })()}

      {kind === 'pdfmerge' && (() => {
        const inputs = (config as { inputs?: string[] }).inputs ?? []
        return (
          <div className="props-field">
            <label>上游节点（{inputs.length} 个，由连线决定）</label>
            <div className="props-filelist">
              {inputs.map((uid) => {
                const up = tables[uid]
                return (
                  <div key={uid} className="props-file-row" title={uid}>
                    <span className="props-file-nm">{up ? up.name : uid + '（已删除）'}</span>
                  </div>
                )
              })}
              {!inputs.length && <span className="hint">空——把上游节点（OO 文档/PDF/照片）连到本节点左口</span>}
            </div>
            <p className="hint">合并顺序 = 连线顺序（左侧输入口从上到下）。增删上游在画布连线操作。</p>
          </div>
        )
      })()}

      {kind === 'pdfreplace' && (() => {
        const cfgR = config as { replaceTargets?: unknown[]; replaceSources?: unknown[]; replaceHistory?: unknown[]; replaceTolerance?: number }
        const tol = cfgR.replaceTolerance ?? 60
        return (
          <>
            <div className="props-field">
              <label>组统计</label>
              <p className="hint">被替换 {(cfgR.replaceTargets ?? []).length} 份 · 替换页 {(cfgR.replaceSources ?? []).length} 份 · 历史 {(cfgR.replaceHistory ?? []).length} 条</p>
            </div>
            <div className="props-field">
              <label>匹配容差（{tol}）</label>
              <input
                type="range" min={0} max={100} step={5} value={tol}
                onChange={(e) => onSetNodeConfig(table.id, { replaceTolerance: Number(e.target.value) })}
              />
              <p className="hint">0 = 严格逐字比对，100 = 宽松。与节点卡片上的容差设置同源（同写 config）。</p>
            </div>
          </>
        )
      })()}

      {kind === 'photo' && (() => {
        const n = ((config as { photoItems?: unknown[] }).photoItems ?? []).length
        return (
          <div className="props-field">
            <label>照片组（{n} 张）</label>
            <p className="hint">图片增删在节点卡片操作（本面板只读）。可连入 PDF 合并（每图一页 A4 适配）。</p>
          </div>
        )
      })()}

      {kind === 'fileout' && (() => {
        const cfgO = config as { outDir?: string; lastSaved?: { name: string; at: number; size: number }[] }
        const saved = cfgO.lastSaved ?? []
        return (
          <>
            <div className="props-field">
              <label>输出目录</label>
              <input value={cfgO.outDir ?? ''} disabled placeholder="未设置——在节点卡片点文件夹钮选择" />
              <p className="hint">连线上游产物（PDF/合并/替换页）→ 点节点上「保存」一键写入该目录。</p>
            </div>
            <div className="props-field">
              <label>最近保存（{saved.length} 条）</label>
              <div className="props-filelist">
                {saved.map((r) => (
                  <div key={r.name + r.at} className="props-file-row" title={r.name}>
                    <span className="props-file-nm">{r.name}</span>
                  </div>
                ))}
                {!saved.length && <span className="hint">尚未保存过。</span>}
              </div>
            </div>
          </>
        )
      })()}

      {kind === 'pdf2word' && (() => {
        const cfgP = config as { p2wItems?: { id: string; name: string; inDb?: boolean; b64?: string }[]; p2wHistory?: { id: string; name: string; at: number }[]; p2wLang?: 'ch' | 'en'; p2wProfile?: 'fast' | 'std' | 'pro'; p2wEngine?: 'x2t' | 'ocr' }
        const items = cfgP.p2wItems ?? []
        const hist = cfgP.p2wHistory ?? []
        const isOcr = cfgP.p2wEngine === 'ocr'
        return (
          <>
            <div className="props-field">
              <label>转换语言（引擎模型随语言切换）</label>
              <select
                value={cfgP.p2wLang ?? 'ch'}
                onChange={(e) => onSetNodeConfig(table.id, { p2wLang: (e.target.value === 'en' ? 'en' : 'ch') as 'ch' | 'en' })}
              >
                <option value="ch">中文（缺省）</option>
                <option value="en">英文</option>
              </select>
              <p className="hint">百度 PaddleOCR 引擎装在本机（首次转换前先装：安装脚本见安装说明）。词条导入/上游连线/执行转换/历史管理都在节点卡片操作。</p>
            </div>
            {isOcr && (
              <div className="props-field">
                <label>速度档位（精度 ↔ 耗时取舍）</label>
                <select
                  value={cfgP.p2wProfile ?? 'pro'}
                  onChange={(e) => onSetNodeConfig(table.id, { p2wProfile: (['fast', 'std', 'pro'].includes(e.target.value) ? e.target.value : 'pro') as 'fast' | 'std' | 'pro' })}
                >
                  <option value="pro">精准（缺省）—— 表格还原成 Word 表格，约 13 秒/页</option>
                  <option value="std">均衡 —— 表格退化为纯文本行，省 4 个模型明显更快</option>
                  <option value="fast">极速 —— 轻量识别模型 + 不还原表格，长文档首选</option>
                </select>
                <p className="hint">档位只在引擎=百度 OCR 时生效；切换档位后引擎会重建一次（多等约 10-90 秒加载）。</p>
              </div>
            )}
            <div className="props-field">
              <label>待转词条（{items.length} 份）</label>
              <div className="props-filelist">
                {items.map((it) => (
                  <div key={it.id} className="props-file-row" title={it.name}>
                    <span className={'pn-loc ' + ((it.inDb || it.b64) ? 'pn-loc--in' : 'pn-loc--ext')}>{(it.inDb || it.b64) ? '内' : '外'}</span>
                    <span className="props-file-nm">{it.name}</span>
                  </div>
                ))}
                {!items.length && <span className="hint">空——在节点卡片导入 PDF 或把上游连到左口。</span>}
              </div>
            </div>
            <div className="props-field">
              <label>最近转换（{hist.length} 条）</label>
              <div className="props-filelist">
                {hist.map((r) => (
                  <div key={r.id} className="props-file-row" title={r.name}>
                    <span className="props-file-nm">{r.name}</span>
                  </div>
                ))}
                {!hist.length && <span className="hint">尚未转换过。</span>}
              </div>
            </div>
          </>
        )
      })()}

      {kind === 'folder' && (() => {
        const cfgF = config as { folderPath?: string; folderMode?: 'read' | 'write' }
        return (
          <div className="props-field">
            <label>目录（{cfgF.folderMode === 'write' ? '写入模式' : '读取模式'}）</label>
            <input value={cfgF.folderPath ?? ''} disabled placeholder="未设置——在节点卡片选择目录" />
            <p className="hint">目录选择/模式切换在节点卡片操作。</p>
          </div>
        )
      })()}

      {kind === 'webmon' && (() => {
        const cfgW = config as { webmonIntervalSec?: number; webmonTargets?: unknown[] }
        return (
          <div className="props-field">
            <label>自动轮询间隔（秒）</label>
            <input
              type="number"
              min={0}
              step={10}
              value={Number(cfgW.webmonIntervalSec) || 0}
              onChange={(e) => onSetNodeConfig(table.id, { webmonIntervalSec: Math.max(0, Number(e.target.value) || 0) })}
            />
            <p className="hint">≥10 秒生效；0 = 仅手动检查（应用打开期间自动轮询）。目标与历史在节点卡片上管理。</p>
          </div>
        )
      })()}

      {kind === 'param' && (
        <>
          <div className="props-field">
            <label>导入外部数据</label>
            <input
              type="file"
              accept=".csv,.json,.xlsx,.xls"
              onChange={(e) => {
                const f = e.target.files?.[0]
                if (f) onImport(f)
                e.target.value = ''
              }}
            />
            <p className="hint">支持 CSV / JSON / Excel。导入后会覆盖本节点的字段与数据。</p>
          </div>
          {/* 备忘 · 待办38：属性页内嵌数据表格 —— 导入/手建的参数行在此直接增删改，
              number 字段输入自动转数值，字段格式（金额¥/数字）在列头标注、数据页按格式渲染 */}
          <ParamDataEditor table={table} rows={rows} onAddRow={onAddRow} onSetCell={onSetCell} onRemoveRow={onRemoveRow} />
        </>
      )}

      {kind === 'start' && (
        <div className="props-field">
          <label>触发方式</label>
          <select
            value={config.trigger ?? 'manual'}
            onChange={(e) => onSetNodeConfig(table.id, { trigger: e.target.value as 'manual' | 'schedule' | 'webhook' })}
          >
            <option value="manual">手动执行</option>
            <option value="schedule">定时（Cron）</option>
            <option value="webhook">Webhook</option>
          </select>
          {config.trigger === 'schedule' && (
            <input
              placeholder="如 0 9 * * *（每天 9 点）"
              value={config.triggerCron ?? ''}
              onChange={(e) => onSetNodeConfig(table.id, { triggerCron: e.target.value })}
            />
          )}
          {config.trigger === 'webhook' && (
            <input
              placeholder="https://... 回调地址"
              value={config.triggerWebhook ?? ''}
              onChange={(e) => onSetNodeConfig(table.id, { triggerWebhook: e.target.value })}
            />
          )}
          <p className="hint">开始节点是工作流的触发起点。点击画布「运行」按钮，会沿数据流连线高亮执行顺序（可视化，不真正处理数据）。</p>
        </div>
      )}

      {/* ── 血缘标记（2026-10-05 用户要求「在节点系统上增加衍生的关联标记」）──
          记录「本节点由哪个节点衍生 + 做了哪些适配」。
          判据不是「记一下」，而是**升级父内核时能照着这张表查出该改哪**。
          声明在 core/nodeStyle 的 KIND_STYLE.derivesFrom（单一真相源），
          这里只读不写。 */}
      {lineage && (
        <div className="props-field props-lineage">
          <label>节点血缘（衍生自{KIND_STYLE[lineage.kind]?.label ?? lineage.kind}）</label>
          <div className="props-hint" style={{ marginBottom: 6 }}>
            代际名 <b>{lineage.generation}</b> · 派生链路{' '}
            {lineageChainOf(kind)
              .slice(1)
              .map((k) => KIND_STYLE[k]?.label ?? k)
              .join(' → ')}
            。<b>OO 内核升级时按下面这张表逐条核对</b>（完整清单见 docs/24-oo迭代影响面清单.md）。
          </div>
          <div className="props-hint" style={{ marginBottom: 6 }}>
            共 {lineage.adaptations.length} 项适配：
          </div>
          {lineage.adaptations.map((a, i) => (
            <div className="lineage-item" key={`${a.area}-${i}`}>
              <span className="lineage-area">{a.area}</span>
              <span className="lineage-note">{a.note}</span>
            </div>
          ))}
        </div>
      )}

      {/* ── 规则锁定（2026-10-05 圆形规则节点）──
          规则节点可编辑并显示锁定清单；被锁的辅节点只读并标「由规则节点统一管理」。
          组件在 src/RuleEditor.tsx，此处只挂载。 */}
      <RuleEditor
        table={table}
        tables={tables}
        onSetRule={onSetRule}
        onClearRule={onClearRule}
        onUnlink={onUnlinkRuleNode}
      />

      {/* 汇聚型节点：数据来源勾选（判定走 nodeStyle 单一真相源，与 engine/App 同源） */}
      {isMergeKind(kind) && (
        <div className="props-field">
          <label>数据来源（可多选，免连线也能引入）</label>
          <div className="src-picker">
            {Object.values(tables)
              .filter((t) => t.id !== table.id)
              .map((t) => {
                const checked = (config.inputs ?? []).includes(t.id)
                const linked = (config.linkedInputs ?? []).includes(t.id)
                return (
                  <label key={t.id} className={`src-opt${checked ? ' src-opt-on' : ''}`}>
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={(e) => {
                        const inputs = config.inputs ?? []
                        const lk = config.linkedInputs ?? []
                        if (e.target.checked) {
                          // 手动勾选：加入 inputs（汇入数据），不画连线；若已连线则保留
                          if (!inputs.includes(t.id)) {
                            onSetNodeConfig(table.id, { inputs: [...inputs, t.id] })
                          }
                        } else {
                          // 取消：从 inputs 与 linkedInputs 都移除
                          onSetNodeConfig(table.id, {
                            inputs: inputs.filter((id) => id !== t.id),
                            linkedInputs: lk.filter((id) => id !== t.id),
                          })
                        }
                      }}
                    />
                    <span>{t.name}</span>
                    {linked && <span className="src-badge">已连线·优先</span>}
                  </label>
                )
              })}
          </div>
          <p className="hint">
            勾选即可引入该节点数据（无需连线）。从上游节点右侧蓝点拖线连到本节点则标记为「优先来源」，其字段排在合并结果前列。
          </p>
        </div>
      )}

      {/* ── ai 节点：Agent 配置（2026-10-06 用户：「模仿常规agent和提示词。
          补充ai节点的，api设置，提示词，skill，输入，输出」）——
          五组配置在子组件，此处只挂载（同 ParamDataEditor 惯例）。 */}
      {/* key=节点id：切换节点时整组件重挂载，内部草稿态归零（否则 A 节点的草稿串到 B） */}
      {kind === 'ai' && (
        <AiAgentConfig key={table.id} table={table} tables={tables} config={config} onSetNodeConfig={onSetNodeConfig} />
      )}

      {kind === 'map' && (
        <>
          <div className="props-field">
            <label>匹配键（用上游哪个字段去查）</label>
            <select
              value={config.mapKeyField ?? ''}
              onChange={(e) => onSetNodeConfig(table.id, { mapKeyField: e.target.value || undefined })}
            >
              <option value="">（未设置）</option>
              {inputFields.map((f) => (
                <option key={f.id} value={f.id}>
                  {f.name}
                </option>
              ))}
            </select>
            {!inputFields.length && (
              <p className="hint">先在上方「数据来源」勾选上游节点，或从上游节点右侧蓝点拖线连到本节点。</p>
            )}
          </div>
          <div className="props-field">
            <label>对照表匹配列</label>
            <select
              value={config.mapMatchField ?? ''}
              onChange={(e) => onSetNodeConfig(table.id, { mapMatchField: e.target.value || undefined })}
            >
              <option value="">（未设置）</option>
              {table.fields.map((f) => (
                <option key={f.id} value={f.id}>
                  {f.name}
                </option>
              ))}
            </select>
          </div>
          <div className="props-field">
            <label>取回列（命中后追加到上游行，可多选）</label>
            <div className="src-picker">
              {table.fields.map((f) => {
                const checked = (config.mapReturnFields ?? []).includes(f.id)
                return (
                  <label key={f.id} className={`src-opt${checked ? ' src-opt-on' : ''}`}>
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={(e) => {
                        const ret = config.mapReturnFields ?? []
                        onSetNodeConfig(table.id, {
                          mapReturnFields: e.target.checked ? [...ret, f.id] : ret.filter((x) => x !== f.id),
                        })
                      }}
                    />
                    <span>{f.name}</span>
                  </label>
                )
              })}
            </div>
          </div>
          <div className="props-field">
            <label>未命中时填</label>
            <input
              value={config.mapDefault ?? ''}
              placeholder="留空则填空字符串"
              onChange={(e) => onSetNodeConfig(table.id, { mapDefault: e.target.value })}
            />
          </div>
          <MapRulesEditor
            table={table}
            rows={rows}
            onAddRow={onAddRow}
            onSetCell={onSetCell}
            onRemoveRow={onRemoveRow}
          />
          <p className="hint">
            对照表本体是一张普通数据表，可自由加列（键 / 值 / 单位 …）。上游数据从本节点流过时，每行按「匹配键」在本表查到对应行，把「取回列」追加到该行；没查到就填默认值。
            不接上游时，它就是一张普通对照表，可被其它节点用 lookup 引用。结果在右侧「数据」页预览。
          </p>
        </>
      )}

      {kind === 'sort' && (
        <>
          <div className="props-field">
            <label>排序列</label>
            <select
              value={config.sortBy ?? ''}
              onChange={(e) => onSetNodeConfig(table.id, { sortBy: e.target.value || undefined })}
            >
              <option value="">（不排序）</option>
              {inputFields.map((f) => (
                <option key={f.id} value={f.id}>
                  {f.name}
                </option>
              ))}
            </select>
          </div>
          <div className="props-field">
            <label>方向</label>
            <select
              value={config.sortDir ?? 'asc'}
              onChange={(e) => onSetNodeConfig(table.id, { sortDir: e.target.value as 'asc' | 'desc' })}
            >
              <option value="asc">升序</option>
              <option value="desc">降序</option>
            </select>
          </div>
        </>
      )}

      {kind === 'output' && (
        <>
          <div className="props-field">
            <label>导出格式</label>
            <select
              value={config.exportFormat ?? 'csv'}
              onChange={(e) => onSetNodeConfig(table.id, { exportFormat: e.target.value as NodeConfig['exportFormat'] })}
            >
              {EXPORT_FORMATS.map((fmt) => (
                <option key={fmt} value={fmt}>
                  {fmt.toUpperCase()}
                </option>
              ))}
            </select>
          </div>
          <button className="bar-btn export-btn" onClick={onExport}>
            导出文件
          </button>
          <p className="hint">导出上游汇入后的全部数据。</p>

          <WordReportPanel
            table={table}
            tables={tables}
            records={records}
            onSetNodeConfig={onSetNodeConfig}
            compact
          />
        </>
      )}

      {/* 文本节点：直接编辑内容（数字或一串文字） */}
      {kind === 'text' && <TextEditor table={table} records={records} onSetNodeConfig={onSetNodeConfig} />}

      {/* Word 文档节点：属性栏里给一份同样的面板，双击则是放大态（同一套组件，两处入口） */}
      {kind === 'docx' && (
        <>
          <p className="hint">
            在模板里放好「内容控件」当作域，再把域绑到上游字段。上游一变，产物跟着变。
            <br />
            想要更大的编辑空间，<b>双击画布上的这个节点</b>。
          </p>
          <WordReportPanel
            table={table}
            tables={tables}
            records={records}
            onSetNodeConfig={onSetNodeConfig}
            compact
          />
        </>
      )}

      {/* 工作簿类节点（Excel / 章节工作簿）—— 绑定面板同款，差异只在章节编排窗体 */}
      {isWorkbookNodeKind(kind) && (
        <>
          <p className="hint">
            工作簿里的「命名区域」与「表格」会自动列成可绑定的域。
            <br />
            {kind === 'chapter' ? (
              <>
                <b>标题分级</b>：在 A 列填 1/2/3 把该行变成章节标题，
                <b>审阅 → 章节编排</b> 可折叠、升降级、生成目录大纲。
                <br />
              </>
            ) : null}
            想要更大的编辑空间，<b>双击画布上的这个节点</b>。
          </p>
          <XlsxReportPanel
            table={table}
            tables={tables}
            records={records}
            onSetNodeConfig={onSetNodeConfig}
          />
        </>
      )}

      {(kind === 'table' || kind === 'param' || kind === 'map') && (
        <div className="props-field">
          <label>字段（{table.fields.length}）</label>
          {table.fields.map((f) => (
            <div key={f.id} className="field-row">
              <input
                className="fr-name"
                value={f.name}
                onChange={(e) => onUpdateField(table.id, f.id, { name: e.target.value })}
                placeholder="字段名"
              />
              <select
                className="fr-type"
                value={f.type}
                onChange={(e) => onUpdateField(table.id, f.id, { type: e.target.value as FieldType })}
              >
                {TYPES.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </select>
              {/* 备忘 · 待办38：数据格式（FieldDef.format）编辑入口 —— currency=金额¥ / number=数字 / text=文本，
                  消费方：数据页 TableForm valueFormatter + 参数表数据表格列头标注 */}
              <select
                className="fr-fmt"
                value={f.format ?? ''}
                title="数据格式（金额带 ¥、数字、文本）"
                onChange={(e) => onUpdateField(table.id, f.id, { format: (e.target.value || undefined) as FieldDef['format'] })}
              >
                <option value="">格式</option>
                <option value="currency">¥ 金额</option>
                <option value="number">数字</option>
                <option value="text">文本</option>
              </select>
              <button className="fr-del" onClick={() => onRemoveField(table.id, f.id)} title="删除字段">
                ×
              </button>
              {f.type === 'formula' && (
                <input
                  className="fr-formula"
                  value={f.formula ?? ''}
                  onChange={(e) => onUpdateField(table.id, f.id, { formula: e.target.value })}
                  placeholder="=[@qty]*[@price]"
                />
              )}
              {f.type === 'lookup' && (
                <LookupEditor field={f} sourceTable={table} tables={tables} onChange={(patch) => onUpdateField(table.id, f.id, patch)} />
              )}
            </div>
          ))}
          <button
            className="bar-btn"
            onClick={() =>
              onAddField(table.id, defineField('新字段', 'text', { id: 'f_' + Math.random().toString(36).slice(2, 10) }))
            }
          >
            ＋ 加字段
          </button>
        </div>
      )}
    </div>
  )
}

/**
 * 对照规则表编辑器：直接在属性面板里维护对照表的行（键/值/…）。
 * 放在属性页而不是「数据」页，是因为数据页显示的是「上游 + 映射结果」的预览，
 * 那些行并不对应本节点自身的记录，在这里改才是对照表本体。
 */
function MapRulesEditor({
  table,
  rows,
  onAddRow,
  onSetCell,
  onRemoveRow,
}: {
  table: TableSchema
  rows: Row[]
  onAddRow: (tableId: string) => void
  onSetCell: (tableId: string, rowIdx: number, fieldId: string, value: unknown) => void
  onRemoveRow: (tableId: string, rowIdx: number) => void
}) {
  return (
    <div className="props-field">
      <label>对照规则（{rows.length} 行）</label>
      <div className="map-rules">
        <div
          className="map-row map-row-head"
          style={{ gridTemplateColumns: `repeat(${table.fields.length}, minmax(0, 1fr)) 24px` }}
        >
          {table.fields.map((f) => (
            <span key={f.id}>{f.name}</span>
          ))}
          <span />
        </div>
        {rows.map((r, i) => (
          <div
            className="map-row"
            key={i}
            style={{ gridTemplateColumns: `repeat(${table.fields.length}, minmax(0, 1fr)) 24px` }}
          >
            {table.fields.map((f) => (
              <input
                key={f.id}
                value={r[f.id] === undefined || r[f.id] === null ? '' : String(r[f.id])}
                placeholder={f.name}
                onChange={(e) => onSetCell(table.id, i, f.id, e.target.value)}
              />
            ))}
            <button className="map-row-del" onClick={() => onRemoveRow(table.id, i)} title="删除该行">
              ×
            </button>
          </div>
        ))}
        {!rows.length && <p className="hint">还没有对照条目，点下面加一行开始建对照关系。</p>}
      </div>
      <button className="bar-btn" onClick={() => onAddRow(table.id)}>
        ＋ 加对照行
      </button>
      {!table.fields.length && <p className="hint">先在下面给对照表加字段（如 键 / 值）。</p>}
    </div>
  )
}

/**
 * 备忘 · 待办38 · 参数表数据表格（属性页内嵌行编辑器，2026-09-27）：
 *   · 导入外来数据 / 手建字段后，参数行直接在此增删改（此前只能去「数据」页，属性页看不到行）
 *   · number 字段输入可解析的数字 → 自动转 Number 存库（下游公式/求值吃数值，不吃字符串）
 *   · 字段数据格式（format）在列头标注（currency 列头带 ¥）；格式化渲染在数据页 TableForm 生效
 * 独立于 MapRulesEditor：对照规则编辑器全 string 直写无转换，二者语义不同不复用（补丁不重建）。
 */
function ParamDataEditor({
  table,
  rows,
  onAddRow,
  onSetCell,
  onRemoveRow,
}: {
  table: TableSchema
  rows: Row[]
  onAddRow: (tableId: string) => void
  onSetCell: (tableId: string, rowIdx: number, fieldId: string, value: unknown) => void
  onRemoveRow: (tableId: string, rowIdx: number) => void
}) {
  const fmtBadge = (f: FieldDef) => (f.format === 'currency' ? ' ¥' : f.format === 'number' ? ' #' : '')
  return (
    <div className="props-field">
      <label>数据表格（{rows.length} 行）</label>
      <div className="map-rules">
        <div
          className="map-row map-row-head"
          style={{ gridTemplateColumns: `repeat(${table.fields.length}, minmax(0, 1fr)) 24px` }}
        >
          {table.fields.map((f) => (
            <span key={f.id} title={f.format === 'currency' ? '金额格式' : f.format === 'number' ? '数字格式' : undefined}>
              {f.name}
              {fmtBadge(f)}
            </span>
          ))}
          <span />
        </div>
        {rows.map((r, i) => (
          <div
            className="map-row"
            key={i}
            style={{ gridTemplateColumns: `repeat(${table.fields.length}, minmax(0, 1fr)) 24px` }}
          >
            {table.fields.map((f) => {
              const raw = r[f.id]
              return (
                <input
                  key={f.id}
                  value={raw === undefined || raw === null ? '' : String(raw)}
                  placeholder={f.name}
                  title={f.type === 'number' ? '数字列：输入数字自动转数值' : undefined}
                  onChange={(e) => {
                    const v = e.target.value
                    // number 字段：可解析非空输入 → Number（空串/非法串原样存，不吞用户的中间态输入）
                    const num = f.type === 'number' && v.trim() !== '' && Number.isFinite(Number(v)) ? Number(v) : v
                    onSetCell(table.id, i, f.id, num)
                  }}
                />
              )
            })}
            <button className="map-row-del" onClick={() => onRemoveRow(table.id, i)} title="删除该行">
              ×
            </button>
          </div>
        ))}
        {!rows.length && <p className="hint">还没有数据行。点「加行」手动建，或从上面导入文件。</p>}
        {!table.fields.length && <p className="hint">先在下方字段区加字段（如 名称 / 单价）。</p>}
      </div>
      <button className="bar-btn" onClick={() => onAddRow(table.id)}>
        ＋ 加行
      </button>
    </div>
  )
}

function LookupEditor({
  field,
  sourceTable,
  tables,
  onChange,
}: {
  field: FieldDef
  sourceTable: TableSchema
  tables: Record<string, TableSchema>
  onChange: (patch: Partial<FieldDef>) => void
}) {
  const lk: LookupSpec = field.lookup ?? { table: '', sourceField: '', matchField: '', returnField: '' }
  const tgt = tables[lk.table]
  const set = (patch: Partial<LookupSpec>) => onChange({ lookup: { ...lk, ...patch } })
  return (
    <div className="lookup-edit">
      <select value={lk.sourceField} onChange={(e) => set({ sourceField: e.target.value })}>
        <option value="">源字段(本表)</option>
        {sourceTable.fields.map((f) => (
          <option key={f.id} value={f.id}>
            {f.name}
          </option>
        ))}
      </select>
      <span className="lk-arrow">→</span>
      <select
        value={lk.table}
        onChange={(e) => set({ table: e.target.value, matchField: '', returnField: '' })}
      >
        <option value="">目标表</option>
        {Object.values(tables).map((t) => (
          <option key={t.id} value={t.id}>
            {t.name}
          </option>
        ))}
      </select>
      <select value={lk.matchField} onChange={(e) => set({ matchField: e.target.value })}>
        <option value="">匹配字段</option>
        {tgt?.fields.map((f) => (
          <option key={f.id} value={f.id}>
            {f.name}
          </option>
        ))}
      </select>
      <select value={lk.returnField} onChange={(e) => set({ returnField: e.target.value })}>
        <option value="">取回字段</option>
        {tgt?.fields.map((f) => (
          <option key={f.id} value={f.id}>
            {f.name}
          </option>
        ))}
      </select>
    </div>
  )
}

/**
 * 备忘 · 文本节点编辑器（重做）：
 *   · 多条目 —— 每条 = 一行记录（引用栏取「文本」字段，首行值即第一条）
 *   · 文本格式 —— WPS 风格常用格式（粗斜下/字号/颜色/对齐），存 config.textStyle，画布节点渲染应用
 */
function TextEditor({
  table,
  records,
  onSetNodeConfig,
}: {
  table: TableSchema
  records: Record<string, Row[]>
  onSetNodeConfig: (tableId: string, patch: Partial<NodeConfig>) => void
}) {
  const setTableData = useBaseStore((s) => s.setTableData)
  const pushHistory = useBaseStore((s) => s.pushHistory)
  const tables = useBaseStore((s) => s.tables)
  const rows = records[table.id] ?? []
  /* 可插入的上游字段（批次 12·①）：数据来源里勾选/连线的节点，其字段作为胶囊候选 */
  const sources = useMemo(() => {
    const ids = table.config?.inputs ?? []
    const out: { nodeId: string; nodeName: string; kind: string; fields: FieldDef[] }[] = []
    for (const id of ids) {
      const t = tables[id]
      if (!t) continue
      const { fields } = computeNode(id, tables, records)
      if (!fields.length) continue
      out.push({ nodeId: id, nodeName: t.name, kind: t.kind, fields })
    }
    return out
  }, [table.config?.inputs, tables, records])
  const style = table.config?.textStyle ?? {}
  const setStyle = (patch: Partial<NonNullable<NodeConfig['textStyle']>>) =>
    onSetNodeConfig(table.id, { textStyle: { ...style, ...patch } })
  const [drafts, setDrafts] = useState<string[]>(rows.map((r) => String((r as Record<string, unknown>).text ?? '')))
  useEffect(() => setDrafts(rows.map((r) => String((r as Record<string, unknown>).text ?? ''))), [records[table.id]])
  const commitRows = (texts: string[]) => {
    pushHistory()
    setTableData(table.id, [{ id: 'text', name: '文本', type: 'text' }], texts.map((t) => ({ text: t })) as Row[])
  }
  const saveAt = (i: number, v: string) => {
    const next = [...drafts]
    next[i] = v
    setDrafts(next)
    if (v !== String((rows[i] as Record<string, unknown> | undefined)?.text ?? '')) commitRows(next)
  }
  /* ── 插入引用胶囊（批次 12 · ①）─────────────────────────────
     在第 i 行正文末尾追加一颗引用胶囊。点「参数」按钮即插入 ——
     拖拽也支持，但那是同一份载荷（见下方cap-drag），点击是零学习成本的入口。 */
  const insertCapsule = (i: number, nodeId: string, fieldId: string, kind: string) => {
    const src = drafts[i] ?? ''
    const at = offsetAfterInsert(src, plainLength(src))
    const mark = buildRefMark(nodeId, fieldId, kind)
    const next = [...drafts]
    next[i] = src.slice(0, at) + mark + src.slice(at)
    setDrafts(next)
    commitRows(next)
  }
  const addEntry = () => {
    commitRows([...drafts, ''])
    setDrafts((d) => [...d, ''])
  }
  const removeAt = (i: number) => {
    commitRows(drafts.filter((_, j) => j !== i))
  }
  const B = ({ on, label, title, toggle }: { on: boolean; label: string; title: string; toggle: () => void }) => (
    <button
      className={`bar-btn${on ? ' wb-alllock' : ''}`}
      style={{ minWidth: 30, fontWeight: on ? 700 : 400 }}
      title={title}
      onClick={toggle}
    >
      {label}
    </button>
  )
  return (
    <>
      <p className="hint">
        文本条目（每条 = 一行记录，其它节点可按条目引用）+ 显示格式。双击画布节点可快速改第一条。
      </p>
      {/* 格式工具行（WPS 风格） */}
      <div className="pp-textfmt" style={{ display: 'flex', flexWrap: 'wrap', gap: 4, alignItems: 'center' }}>
        <B on={!!style.bold} label="B" title="加粗" toggle={() => setStyle({ bold: !style.bold })} />
        <B on={!!style.italic} label="I" title="斜体" toggle={() => setStyle({ italic: !style.italic })} />
        <B on={!!style.underline} label="U" title="下划线" toggle={() => setStyle({ underline: !style.underline })} />
        <select
          className="bar-btn"
          style={{ minWidth: 56 }}
          value={style.fontSize ?? ''}
          onChange={(e) => setStyle({ fontSize: e.target.value ? Number(e.target.value) : undefined })}
          title="字号"
        >
          <option value="">字号</option>
          {[12, 14, 16, 18, 20, 24, 30].map((s) => (
            <option key={s} value={s}>
              {s}
            </option>
          ))}
        </select>
        <input
          type="color"
          value={style.color || '#e8edf2'}
          onChange={(e) => setStyle({ color: e.target.value })}
          title="文字颜色"
          style={{ width: 30, height: 26, padding: 0, border: '1px solid var(--ed-border, #ccc)', borderRadius: 4, background: 'transparent', cursor: 'pointer' }}
        />
        {(['left', 'center', 'right'] as const).map((a) => (
          <B
            key={a}
            on={style.align === a}
            label={a === 'left' ? '⯇' : a === 'center' ? '☰' : '⯈'}
            title={`对齐：${a === 'left' ? '左' : a === 'center' ? '中' : '右'}`}
            toggle={() => setStyle({ align: style.align === a ? undefined : a })}
          />
        ))}
      </div>
      {/* 条目列表 */}
      {drafts.map((d, i) => (
        <div key={i} style={{ display: 'flex', gap: 4, alignItems: 'flex-start', marginTop: 6 }}>
          <span style={{ fontSize: 11, color: 'var(--pp-dim, var(--c-text3))', lineHeight: '30px', flex: 'none' }}>{i + 1}.</span>
          <textarea
            className="pp-textarea"
            value={d}
            rows={2}
            style={{ flex: 1 }}
            placeholder={i === 0 ? '第一条（引用栏取的就是它）' : `第 ${i + 1} 条`}
            onChange={(e) => {
              const next = [...drafts]
              next[i] = e.target.value
              setDrafts(next)
            }}
            onBlur={(e) => saveAt(i, e.target.value)}
          />
          {/* ①插入引用胶囊：把上游参数作为胶囊塞进这行文字末尾。
              载荷与拖拽一致（cap-drag），点击是零学习成本的入口。 */}
          {sources.length > 0 && (
            <div className="pp-cap-insert">
              <span className="pp-cap-insert-label" title="点一下，把该参数作为胶囊插入这行文字末尾">
                插入引用
              </span>
              <div className="pp-cap-insert-list">
                {sources.map((s) => (
                  <div key={s.nodeId} className="pp-cap-insert-grp">
                    <span className="pp-cap-insert-src" title={`来源节点：${s.nodeName}`}>
                      {s.nodeName}
                    </span>
                    {s.fields.map((f) => (
                      <button
                        key={f.id}
                        className={`ref-capsule cap-cap-kind-${s.kind} pp-cap-chip`}
                        title={`插入引用：${s.nodeName}.${f.name}（点击插到本行末尾，或拖到画布正文任意位置）`}
                        draggable
                        onDragStart={(e) => {
                          //拖拽载荷：把要插入的胶囊标记原文带出去，落点由 Canvas 按指针位置定
                          e.dataTransfer.setData(
                            'application/capsule-mark',
                            buildRefMark(s.nodeId, f.id, s.kind),
                          )
                          e.dataTransfer.effectAllowed = 'copy'
                        }}
                        onClick={() => insertCapsule(i, s.nodeId, f.id, s.kind)}
                      >
                        <i className="cap-dot" aria-hidden="true" />
                        <span className="cap-field">{f.name}</span>
                      </button>
                    ))}
                  </div>
                ))}
              </div>
            </div>
          )}
          <button
            className="bar-btn"
            title="删除该条目"
            style={{ flex: 'none', color: '#f0883e' }}
            onClick={() => removeAt(i)}
          >
            ×
          </button>
        </div>
      ))}
      <button className="bar-btn" style={{ marginTop: 6 }} onClick={addEntry}>
        ＋ 添加条目
      </button>
    </>
  )
}

/* ═════════════════════════════════════════════════════════════════════
   AI 节点 Agent 配置（2026-10-06 用户：「模仿常规agent和提示词。
   补充ai节点的，api设置，提示词，skill，输入，输出」）
   五组：① API 设置 ② 技能 ③ 提示词 ④ 输入变量 ⑤ 输出。
   ★ 提交纪律：pushSetNodeConfig 每调用压一条历史快照（App.tsx，栈上限 50）
     —— 文本/数字一律「本地草稿 + onBlur 才落库」，打字不清空撤销栈；
     select 一次一击，直连无碍。
   ★ 技能单一真相源是 schema.AI_SKILLS：下拉、骨架提示词、推荐输出模式
     全从注册表取——往数组加一项即出现在面板，零组件改动。
   ★ 输入变量以 linkedInputs 为准：断线行自动消失（schema 注释同款口径）。
   ═════════════════════════════════════════════════════════════════════ */

const AI_MODE_LABEL: Record<'text' | 'json' | 'markdown', string> = {
  text: '纯文本',
  json: 'JSON',
  markdown: 'Markdown',
}

function AiAgentConfig({
  table,
  tables,
  config,
  onSetNodeConfig,
}: {
  table: TableSchema
  tables: Record<string, TableSchema>
  config: NodeConfig
  onSetNodeConfig: (tableId: string, patch: Partial<NodeConfig>) => void
}) {
  const id = table.id

  /* 长文本/数字草稿：受控编辑、失焦才落库（保护撤销栈） */
  const [draft, setDraft] = useState({
    apiBase: config.aiApiBase ?? '',
    apiKey: config.aiApiKey ?? '',
    model: config.aiModel ?? '',
    temperature: config.aiTemperature === undefined ? '' : String(config.aiTemperature),
    maxTokens: config.aiMaxTokens === undefined ? '' : String(config.aiMaxTokens),
    systemPrompt: config.aiSystemPrompt ?? '',
    userPrompt: config.aiUserPrompt ?? '',
    outputNote: config.aiOutputNote ?? '',
  })
  /* 撤销 / 技能骨架预填等外部 config 变更 → 草稿跟随；
     打字只动本地草稿不改 config，不会触发本 effect（不扰输入） */
  useEffect(() => {
    setDraft({
      apiBase: config.aiApiBase ?? '',
      apiKey: config.aiApiKey ?? '',
      model: config.aiModel ?? '',
      temperature: config.aiTemperature === undefined ? '' : String(config.aiTemperature),
      maxTokens: config.aiMaxTokens === undefined ? '' : String(config.aiMaxTokens),
      systemPrompt: config.aiSystemPrompt ?? '',
      userPrompt: config.aiUserPrompt ?? '',
      outputNote: config.aiOutputNote ?? '',
    })
  }, [config])

  const [showKey, setShowKey] = useState(false)
  /* 变量名草稿（按上游节点 id 存）：blur 提交后清掉回受控显示 */
  const [varDrafts, setVarDrafts] = useState<Record<string, string>>({})

  const setD = (k: keyof typeof draft, v: string) => setDraft((d) => ({ ...d, [k]: v }))

  /** 文本失焦提交：空串 = 清除配置（undefined → 面板回占位/默认），否则原样保存 */
  const commitText = (
    key: 'aiApiBase' | 'aiApiKey' | 'aiModel' | 'aiSystemPrompt' | 'aiUserPrompt' | 'aiOutputNote',
    raw: string,
  ) => {
    onSetNodeConfig(id, { [key]: raw.trim() === '' ? undefined : raw } as Partial<NodeConfig>)
  }
  /** 数字失焦提交：空/非法 = 清除（执行端用缺省 0.7 / 2048）；温度 clamp 0~2，token 取正整 */
  const commitNum = (key: 'aiTemperature' | 'aiMaxTokens', raw: string) => {
    const v = raw.trim()
    if (v === '' || Number.isNaN(Number(v))) {
      onSetNodeConfig(id, { [key]: undefined } as Partial<NodeConfig>)
      return
    }
    let n = Number(v)
    if (key === 'aiTemperature') n = Math.min(2, Math.max(0, n))
    else n = Math.max(1, Math.round(n))
    onSetNodeConfig(id, { [key]: n } as Partial<NodeConfig>)
  }

  /** 选技能 = 套模板：骨架系统提示词 + 推荐输出模式一并预填（custom/清空不覆盖已填内容） */
  const pickSkill = (skillId: string) => {
    const patch: Partial<NodeConfig> = { aiSkillId: skillId || undefined }
    const skill = AI_SKILLS.find((s) => s.id === skillId)
    if (skill) {
      patch.aiSystemPrompt = skill.systemPrompt
      patch.aiOutputMode = skill.outputMode
    }
    onSetNodeConfig(id, patch)
  }

  /** 上游节点（仍连着的）：渲染变量行的依据，断线即消失 */
  const upstream = useMemo(
    () => (config.linkedInputs ?? []).map((nid) => tables[nid]).filter((t): t is TableSchema => !!t),
    [config.linkedInputs, tables],
  )
  /** 变量显示名：起过名用变量名，否则用上游节点名兜底（执行端同款口径） */
  const varNameOf = (nodeId: string) => {
    const named = (config.aiInputVars ?? []).find((x) => x.from === nodeId)
    return named?.name ?? tables[nodeId]?.name ?? nodeId
  }
  /** 改变量名只动 aiInputVars，不碰连线；清空 = 删除该条目（兜底回节点名） */
  const commitVarName = (from: string, raw: string) => {
    const vars = config.aiInputVars ?? []
    const v = raw.trim()
    const next = v
      ? vars.some((x) => x.from === from)
        ? vars.map((x) => (x.from === from ? { ...x, name: v } : x))
        : [...vars, { name: v, from }]
      : vars.filter((x) => x.from !== from)
    onSetNodeConfig(id, { aiInputVars: next.length ? next : undefined })
  }

  /** 运行中标记（瞬态，不进 config） */
  const [running, setRunning] = useState(false)
  /**
   * 运行本节点：拉上游 → 插值 → fetch OpenAI 兼容 chat/completions。
   * 结果/错误走 store 直写 setNodeConfig（不经过 onSetNodeConfig 的 push 包装）
   * —— 运行是「执行」不是「编辑」，不该吃历史快照；persist 照常保存
   * （store.setNodeConfig 本身写 state，持久化层跟 state 走）。
   */
  const runAi = async () => {
    /* ★ 2026-11-09 执行逻辑抽到 core/aiRun.runAiNode（webmon 变化自动触发共用同一条链路） */
    setRunning(true)
    try { await runAiNode(id, useBaseStore.getState()) } finally { setRunning(false) }
  }


  const skill = AI_SKILLS.find((s) => s.id === config.aiSkillId)

  return (
    <>
      {/* ① API 设置 */}
      <div className="props-divider" />
      <div className="props-field">
        <label>API 设置（OpenAI 兼容）</label>
        <input
          placeholder="接口地址，如 https://api.deepseek.com/v1"
          value={draft.apiBase}
          onChange={(e) => setD('apiBase', e.target.value)}
          onBlur={(e) => commitText('aiApiBase', e.target.value)}
        />
        <p className="hint">★ 留空 = 使用本地大模型（零配置，数据不出电脑）；填外部 OpenAI 兼容接口则需 Key 与模型名。</p>
        <div style={{ display: 'flex', gap: 4 }}>
          <input
            type={showKey ? 'text' : 'password'}
            placeholder="API Key（sk-...）"
            value={draft.apiKey}
            onChange={(e) => setD('apiKey', e.target.value)}
            onBlur={(e) => commitText('aiApiKey', e.target.value)}
            style={{ flex: 1 }}
          />
          <button className="bar-btn" title={showKey ? '隐藏 Key' : '显示 Key'} onClick={() => setShowKey(!showKey)}>
            {showKey ? '隐藏' : '显示'}
          </button>
        </div>
        <input
          placeholder="模型名，如 deepseek-chat"
          list="ai-model-list"
          value={draft.model}
          onChange={(e) => setD('model', e.target.value)}
          onBlur={(e) => commitText('aiModel', e.target.value)}
        />
        <datalist id="ai-model-list">
          {['deepseek-chat', 'deepseek-reasoner', 'gpt-4o-mini', 'gpt-4o', 'qwen-plus', 'glm-4-flash'].map((m) => (
            <option key={m} value={m} />
          ))}
        </datalist>
        <div style={{ display: 'flex', gap: 6 }}>
          <input
            type="number"
            min={0}
            max={2}
            step={0.1}
            placeholder="温度（默认 0.7）"
            title="温度 0~2：越高越发散，默认 0.7"
            value={draft.temperature}
            onChange={(e) => setD('temperature', e.target.value)}
            onBlur={(e) => commitNum('aiTemperature', e.target.value)}
            style={{ flex: 1 }}
          />
          <input
            type="number"
            min={1}
            step={100}
            placeholder="最大输出（默认 2048）"
            title="最大输出 token 数"
            value={draft.maxTokens}
            onChange={(e) => setD('maxTokens', e.target.value)}
            onBlur={(e) => commitNum('aiMaxTokens', e.target.value)}
            style={{ flex: 1 }}
          />
        </div>
        <p className="hint">
          ★ API Key 明文随工程（.wfc）一起保存——单机自用换机即用；把工程文件发给别人前，请先清掉这一格。
        </p>
      </div>

      {/* ② 技能 */}
      <div className="props-field">
        <label>技能</label>
        <select value={config.aiSkillId ?? ''} onChange={(e) => pickSkill(e.target.value)}>
          <option value="">自定义（不套模板）</option>
          {AI_SKILLS.map((s) => (
            <option key={s.id} value={s.id}>
              {s.name}
            </option>
          ))}
        </select>
        {skill && (
          <p className="hint">
            {skill.desc}。已按技能预填骨架提示词与推荐输出（{AI_MODE_LABEL[skill.outputMode]}），下方可再手改。
          </p>
        )}
      </div>

      {/* ③ 提示词 */}
      <div className="props-field">
        <label>提示词</label>
        <textarea
          rows={5}
          placeholder={'系统提示词：给 AI 定角色与规则（选技能会预填骨架）。'}
          value={draft.systemPrompt}
          onChange={(e) => setD('systemPrompt', e.target.value)}
          onBlur={(e) => commitText('aiSystemPrompt', e.target.value)}
        />
        <textarea
          rows={4}
          placeholder={'用户提示词：本次任务内容，可用 {{变量名}} 引用输入。\n如：请总结 {{客户资料}} 中的欠费风险点'}
          value={draft.userPrompt}
          onChange={(e) => setD('userPrompt', e.target.value)}
          onBlur={(e) => commitText('aiUserPrompt', e.target.value)}
        />
        {upstream.length > 0 && (
          <p className="hint">可用变量：{upstream.map((u) => `{{${varNameOf(u.id)}}}`).join('、')}</p>
        )}
      </div>

      {/* ④ 输入变量（以画布连线为准） */}
      <div className="props-field">
        <label>输入变量（{upstream.length}）</label>
        {upstream.length === 0 ? (
          <p className="hint">
            还没有上游：从普通节点的下1口拖线连到本节点，再在这里给每路上游起变量名，提示词里用 {'{{变量名}}'} 引用。
          </p>
        ) : (
          upstream.map((up) => (
            <div key={up.id} style={{ display: 'flex', gap: 4, alignItems: 'center', marginBottom: 4 }}>
              <input
                style={{ flex: 1 }}
                placeholder={up.name}
                title="变量名（提示词里用 {{}} 引用）"
                value={varDrafts[up.id] ?? varNameOf(up.id)}
                onChange={(e) => setVarDrafts((m) => ({ ...m, [up.id]: e.target.value }))}
                onBlur={(e) => {
                  commitVarName(up.id, e.target.value)
                  setVarDrafts((m) => {
                    const n = { ...m }
                    delete n[up.id]
                    return n
                  })
                }}
              />
              <span style={{ flex: 'none', fontSize: 11, color: 'var(--pp-dim, var(--c-text3))' }}>← {up.name}</span>
            </div>
          ))
        )}
        {upstream.length > 0 && (
          <p className="hint">
            以画布连线为准：断开连线后对应行自动消失；变量名清空则用「← 上游节点名」兜底。改名只影响提示词引用，不动连线。
          </p>
        )}
      </div>

      {/* ⑤ 输出 */}
      <div className="props-field">
        <label>输出</label>
        <select
          value={config.aiOutputMode ?? 'text'}
          onChange={(e) => onSetNodeConfig(id, { aiOutputMode: e.target.value as 'text' | 'json' | 'markdown' })}
        >
          <option value="text">纯文本</option>
          <option value="json">JSON（下游可按字段解析）</option>
          <option value="markdown">Markdown</option>
        </select>
        <textarea
          rows={3}
          placeholder={
            (config.aiOutputMode ?? 'text') === 'json'
              ? '字段清单（给模型的格式约束），如：{"名称": string, "金额": number, "结论": string}'
              : '格式要求（可留空），如：按条列出，每条不超过 40 字'
          }
          value={draft.outputNote}
          onChange={(e) => setD('outputNote', e.target.value)}
          onBlur={(e) => commitText('aiOutputNote', e.target.value)}
        />
        <p className="hint">输出说明会拼进系统消息：json 模式写字段清单（下游节点将来按字段解析），text/markdown 写格式要求。</p>
      </div>

      {/* ⑥ 运行（2026-10-06 用户：「先做一个输入文字输出文字的样例」）——
          拉上游 → 插值 → 调 OpenAI 兼容 chat/completions → 结果随工程保存。
          结果写 config 走 store 直写（不进撤销栈——运行是执行不是编辑）。 */}
      <div className="props-field">
        <label>运行</label>
        <div style={{ display: 'flex', gap: 4 }}>
          <button
            className="bar-btn"
            style={{ flex: 1 }}
            disabled={running || upstream.length === 0}
            title={upstream.length === 0 ? '先连上游再运行' : '调 API 运行本节点'}
            onClick={() => void runAi()}
          >
            {running ? '运行中…' : '▶ 运行（调 API）'}
          </button>
          {!!config.aiLastOutput && (
            <button className="bar-btn" title="复制运行结果" onClick={() => void navigator.clipboard.writeText(config.aiLastOutput ?? '')}>
              复制
            </button>
          )}
        </div>
        {config.aiLastError && (
          <p className="hint" style={{ color: '#f0883e' }}>✗ {config.aiLastError}</p>
        )}
        {config.aiLastOutput != null && config.aiLastOutput !== '' && (
          <textarea rows={6} readOnly value={config.aiLastOutput} title="最近一次运行结果" />
        )}
        <p className="hint">
          拉取上游内容 → 替换 {'{{变量}}'} → 调 API（OpenAI 兼容 chat/completions）→ 结果显示在这里并随工程保存（不进撤销栈）。用户提示词为空时，上游原文直接作为本次消息。
        </p>
      </div>
    </>
  )
}
