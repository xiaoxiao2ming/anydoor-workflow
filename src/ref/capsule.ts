import type { NodeKind } from '../core/schema'

/**
 * 批次 8 · 统一引用胶囊（两段式，纯表现层 —— 存储零改动，值始终是 node://… 字符串）
 *
 * 结构约定（面板参数行 / 格内徽章 / 信息栏 chip 共用一套 CSS 类）：
 *   <span class="ref-capsule cap-{type}" title="…">
 *     <i class="cap-dot" aria-hidden="true" />
 *     <span class="cap-node">节点名</span>      ← 深底段
 *     <span class="cap-field">参数名</span>      ← 类型色浅底段（整节点引用无 locator → 显示「整体」）
 *   </span>
 * 左缘类型色点 + 中间细缝由 CSS 完成；按节点类型分色（text/xlsx/docx/param 各一色系，其余默认金系），
 * 暗亮双主题各一套（:root[data-theme] 变量，portal 面板同样跟随）。
 * 节点/字段失效（节点被删、字段不存在）→ 追加 .dead 红边失效态。
 */

/** kind → 类型色类。设计定稿拍板：text/xlsx/docx/param 各一色系；其余（table/lookup/sort/output/start/map）落默认金系 */
export function capsuleKindClass(kind?: NodeKind | string): string {
  switch (kind) {
    case 'text':
      return 'cap-text'
    case 'xlsx':
      return 'cap-xlsx'
    case 'docx':
      return 'cap-docx'
    case 'param':
      return 'cap-param'
    default:
      return 'cap-table'
  }
}

const esc = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

export type CapsuleHtmlInput = {
  /** 左段：节点名（节点已失效时传 URN 里的 target 原样展示） */
  node: string
  /** 右段：参数名；空 = 整节点引用 → 显示「整体」 */
  field?: string
  /** 节点类型（决定色系） */
  kind?: NodeKind | string
  /** 悬停全名提示 */
  title?: string
  /** 失效态（红边） */
  dead?: boolean
}

/** 输出两段式胶囊 HTML 字符串（ag-grid cellRenderer 等字符串渲染处使用；JSX 处直接按结构约定手写） */
export function capsuleHtml(o: CapsuleHtmlInput): string {
  const cls = ['ref-capsule', capsuleKindClass(o.kind), o.dead ? 'dead' : ''].filter(Boolean).join(' ')
  const dot = '<i class="cap-dot" aria-hidden="true"></i>'
  const node = `<span class="cap-node">${esc(o.node)}</span>`
  const field = `<span class="cap-field">${esc(o.field || '整体')}</span>`
  const title = o.title ? ` title="${esc(o.title)}"` : ''
  return `<span class="${cls}"${title}>${dot}${node}${field}</span>`
}
