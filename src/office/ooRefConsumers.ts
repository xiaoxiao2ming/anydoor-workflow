/**
 * OO 域外写回 · 消费者注册表
 *
 * 立档：2026-10-29 · 解决「改一个单元格 ⇒ 该不该写回、写到哪几个 OO」
 *
 * 为什么需要：store 没有「数据变化订阅」机制，
 * 而在每个编辑入口都手动触发写回会漏（将来新增编辑点必然忘记挂）。
 * ⇒ 反过来做：**引用方（OO 节点）自己注册**，
 *   编辑方只按「谁引用了我」查表触发 —— 新增编辑点无需改动。
 *
 * ★ 模块级单例的理由：注册表是**跨组件的全局路由**，
 *   放 store 里会给纯数据层塞进 IO 语义；放 Context 会让 NodeEditor 之外
 *   的地方（如命令面板、宏）无法登记。
 */

export type OoRefConsumer = () => void | Promise<void>

/** nodeId → 该 OO 节点的 autoRefresh 回调集合（可能有多个编辑器实例） */
const registry = new Map<string, Set<OoRefConsumer>>()

/** 供 App.tsx 在编辑时零成本查询：「谁引用了我」 */
export const ooRefConsumers = {
  /** 某个节点被几个 OO 节点引用（用来快速跳过无关编辑） */
  get (nodeId?: string): Set<OoRefConsumer> | null {
    if (!nodeId) return null
    const s = registry.get(nodeId)
    return s && s.size ? s : null
  },
  /** 引用方（OO 节点）登记自己；返回注销函数 */
  register (nodeId: string, fn: OoRefConsumer): () => void {
    let set = registry.get(nodeId)
    if (!set) { set = new Set(); registry.set(nodeId, set) }
    set.add(fn)
    return () => {
      const cur = registry.get(nodeId)
      if (!cur) return
      cur.delete(fn)
      if (!cur.size) registry.delete(nodeId)   // 空的删掉，避免 map 无限增长
    }
  },
  /** 该OO 节点是否被别处引用（登记时自检，避免自我循环） */
  isReferenced (nodeId: string): boolean {
    const s = registry.get(nodeId)
    return !!(s && s.size)
  },
  /** 测试用：清空全部登记 */
  _reset (): void { registry.clear() },
  /** 测试用：当前登记总数 */
  _size (): number { return registry.size },
}
