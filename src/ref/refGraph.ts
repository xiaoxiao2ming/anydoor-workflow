/**
 * 节点间数据引用协议（与节点类型无关）
 *
 * 核心认知：Word、Excel 只是两种「节点」。数据引用是画布层面的通用机制，
 * 引用层不该知道、也不关心两端是什么节点。
 *
 *   · 每个节点声明 输出端口（outputs）与 输入槽位（slots）
 *   · 任意槽位可绑定一个引用地址（URN）
 *   · 引用即依赖 —— 依赖图直接从 `node://X#port` 这种引用里推导，不需要单独声明
 *   · 按拓扑序解析：上游先执行产出，下游再取值；成环直接报错
 *
 * 于是「所有节点之间的引用方式」收敛成一件事：
 *   槽位 ← URN → 值
 */

import { resolveUrn, parseUrn, type RefContext } from './urn'

/* ────────────────── 端口 / 槽位 / 绑定 ────────────────── */

export interface PortSpec {
  key: string
  label?: string
  /** 语义类型提示，如 'number' | 'rows' | 'text'，仅用于 UI 与校验 */
  type?: string
}

export interface SlotSpec {
  nodeId: string
  key: string
  label?: string
  /** 由节点适配器定义：'sdt' | 'definedName' | 'table' | 'param' | 'series' | … */
  kind?: string
  meta?: Record<string, unknown>
}

export interface Binding {
  nodeId: string
  slot: string
  urn: string
}

export interface NodeDef {
  id: string
  /** 展示名，如「水力计算」「工程量表」「设计说明」 */
  title?: string
  /** 节点类型标签，仅用于展示 */
  kind?: string
  outputs?: PortSpec[]
  slots?: SlotSpec[]
  /** 执行器：吃槽位值，吐输出端口值；纯函数，不碰 IO */
  run?: (inputs: Record<string, unknown>) => Record<string, unknown>
}

export interface GraphTrace {
  nodeId: string
  slot: string
  urn: string
  from: string
  value: unknown
  ok: boolean
  error?: string
}

export interface GraphResult {
  /** nodeId → { 槽位: 解析到的值 } */
  values: Map<string, Record<string, unknown>>
  /** nodeId → 执行器产出的输出端口值 */
  outputs: Map<string, Record<string, unknown>>
  traces: GraphTrace[]
  order: string[]
}

/* ────────────────── 引用图 ────────────────── */

export class RefGraph {
  nodes = new Map<string, NodeDef>()
  bindings: Binding[] = []

  addNode(def: NodeDef): this {
    this.nodes.set(def.id, def)
    return this
  }

  /** 绑定：某节点的某槽位 ← 引用地址（同一槽位重复绑定会覆盖） */
  bind(nodeId: string, slot: string, urn: string): this {
    const i = this.bindings.findIndex((b) => b.nodeId === nodeId && b.slot === slot)
    if (i >= 0) this.bindings[i] = { nodeId, slot, urn }
    else this.bindings.push({ nodeId, slot, urn })
    return this
  }

  unbind(nodeId: string, slot: string): this {
    this.bindings = this.bindings.filter((b) => !(b.nodeId === nodeId && b.slot === slot))
    return this
  }

  bindingsOf(nodeId: string): Binding[] {
    return this.bindings.filter((b) => b.nodeId === nodeId)
  }

  /**
   * 依赖图：从引用里推导。
   * 只要 A 的某槽位引用了 node://B#…，A 就依赖 B —— 引用即依赖，不用额外声明。
   */
  deps(): Map<string, Set<string>> {
    const m = new Map<string, Set<string>>()
    for (const id of this.nodes.keys()) m.set(id, new Set())
    for (const b of this.bindings) {
      const u = parseUrn(b.urn)
      if (u?.scheme === 'node' && this.nodes.has(u.target) && u.target !== b.nodeId) {
        m.get(b.nodeId)?.add(u.target)
      }
    }
    return m
  }

  /** 拓扑排序；成环抛错并给出环路径 */
  order(): string[] {
    const deps = this.deps()
    const out: string[] = []
    const state = new Map<string, 0 | 1 | 2>() // 0 未访问 / 1 访问中 / 2 已完成
    const visit = (id: string, stack: string[]): void => {
      const st = state.get(id) ?? 0
      if (st === 2) return
      if (st === 1) throw new Error('节点引用成环：' + [...stack, id].join(' → '))
      state.set(id, 1)
      for (const d of deps.get(id) ?? []) visit(d, [...stack, id])
      state.set(id, 2)
      out.push(id)
    }
    for (const id of this.nodes.keys()) visit(id, [])
    return out
  }

  /** 找环（返回一条环路径，无环返回 null） */
  findCycle(): string[] | null {
    const deps = this.deps()
    const state = new Map<string, 0 | 1 | 2>()
    let cycle: string[] | null = null
    const visit = (id: string, stack: string[]): void => {
      if (cycle) return
      const st = state.get(id) ?? 0
      if (st === 2) return
      if (st === 1) {
        const i = stack.indexOf(id)
        cycle = [...stack.slice(i), id]
        return
      }
      state.set(id, 1)
      for (const d of deps.get(id) ?? []) visit(d, [...stack, id])
      state.set(id, 2)
    }
    for (const id of this.nodes.keys()) visit(id, [])
    return cycle
  }

  /**
   * 按拓扑序解析全部槽位；有 run 的节点顺带执行并把输出注册给下游。
   * base 里放外部上下文（file / api / param）。
   */
  resolveAll(base: RefContext = {}): GraphResult {
    const order = this.order()
    const values = new Map<string, Record<string, unknown>>()
    const outputs = new Map<string, Record<string, unknown>>()
    const traces: GraphTrace[] = []
    const ctx: RefContext = { ...base, nodes: { ...(base.nodes ?? {}) } }

    for (const nodeId of order) {
      const node = this.nodes.get(nodeId)!
      const vals: Record<string, unknown> = {}

      for (const b of this.bindingsOf(nodeId)) {
        const v = resolveUrn(b.urn, ctx)
        vals[b.slot] = v
        const u = parseUrn(b.urn)
        traces.push({
          nodeId,
          slot: b.slot,
          urn: b.urn,
          from: u ? u.scheme : 'unresolved',
          value: v,
          ok: v !== undefined,
          error: v === undefined ? '引用未解析到值' : undefined,
        })
      }
      values.set(nodeId, vals)

      if (node.run) {
        const out = node.run(vals) ?? {}
        outputs.set(nodeId, out)
        // 注册到上下文，下游用 node://<nodeId>#<port> 就能取到
        ctx.nodes![nodeId] = out
      }
    }
    return { values, outputs, traces, order }
  }
}
