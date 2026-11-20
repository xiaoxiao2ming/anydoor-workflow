import { Fragment } from 'react'
import {
  parseInline,
  inlineCapsuleClass,
  capsuleLabelOf,
  capsuleFieldOf,
  capsuleTitleOf,
  type InlineToken,
} from './inline'

/**
 * 行内胶囊渲染（批次 12· React 侧）
 *
 * 与 capsule.ts 的区别：那个函数产出 HTML 字符串（给 ag-grid cellRenderer 等字符串渲染处用），
 * 这个是 React 组件（给画布节点正文这类富排版处用）。两者结构约定完全一致，
 * 同一个标记在任何地方长一个样。
 *
 * 用法：
 *   <InlineCapsules value={text} nodeNameOf={(id) => tables[id]?.name} deadOf={(id) => !tables[id]} />
 */
export function InlineCapsules({
  value,
  nodeNameOf,
  deadOf,
  placeholder = '（空）',
}: {
  value: unknown
  /** nodeId → 节点显示名（悬停提示用，让胶囊显示人话而不是 id） */
  nodeNameOf?: (nodeId: string) => string | undefined
  /** nodeId → 该引用是否已失效（节点被删） */
  deadOf?: (nodeId: string) => boolean | undefined
  placeholder?: string
}) {
  const tokens = parseInline(value)
  if (!tokens.length) return <>{placeholder}</>

  return (
    <>
      {tokens.map((t, i) =>
        t.type === 'text' ? (
          <Fragment key={i}>{t.value}</Fragment>
        ) : (
          <CapsuleSpan key={i} token={t} nodeNameOf={nodeNameOf} deadOf={deadOf} />
        ),
      )}
    </>
  )
}

/** 单个胶囊。两段式：左段（深底）= 节点名/常量值，右段（浅底）= 参数名。 */
function CapsuleSpan({
  token,
  nodeNameOf,
  deadOf,
}: {
  token: Extract<InlineToken, { type: 'capsule' }>
  nodeNameOf?: (nodeId: string) => string | undefined
  deadOf?: (nodeId: string) => boolean | undefined
}) {
  // 常量胶囊不查节点；引用胶囊按目标节点判失效
  const isConst = token.literal !== undefined
  const dead = !isConst && token.target !== 'const' ? deadOf?.(token.target) : false

  return (
    <span className={inlineCapsuleClass(token.kind, dead)} title={capsuleTitleOf(token, nodeNameOf?.(token.target))}>
      <i className="cap-dot" aria-hidden="true" />
      <span className="cap-node">{capsuleLabelOf(token)}</span>
      <span className="cap-field">{capsuleFieldOf(token)}</span>
    </span>
  )
}