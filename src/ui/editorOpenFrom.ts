/* ★ 2026-10-23 用户指令「双击节点打开加动效，像打开 app 一样，苹果那种丝滑展开」：
 *   Canvas 双击时把节点屏幕 rect 写进这里；NodeEditor 挂载时消费一次作展开起点
 *   （计算窗体与节点中心的位移/缩放，macOS app-open 式 keyframes 展开）。
 *   模块级暂存（非事件——编辑器是双击之后才挂载，事件已错过）。 */
export interface OpenFromRect { x: number; y: number; w: number; h: number }

export const editorOpenFrom: { rect: OpenFromRect | null } = { rect: null }
