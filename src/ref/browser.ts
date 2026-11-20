/**
 * 引用协议层的浏览器入口
 *
 * src/ref/ 是纯计算层（零 IO 依赖），所以可以直接打进浏览器。
 * 这个入口用于 Web 端 / 演示页 —— 注意它不含 xlsx、docx 适配器（那些依赖 Node 侧库）。
 */

export { RefGraph } from './refGraph'
export type {
  PortSpec,
  SlotSpec,
  Binding,
  NodeDef,
  GraphTrace,
  GraphResult,
} from './refGraph'

export { ParamTable } from './paramTable'
export type { ParamSpec, ResolveTrace } from './paramTable'

export { parseUrn, resolveUrn, pickByLocator, getPath, fmtValue } from './urn'
export type { RefContext, ResolvedRef, RefScheme } from './urn'
