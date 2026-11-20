import { memo, type ReactNode, type MouseEvent } from 'react'

/* ★ 节点底部按钮装载条（2026-10-26 用户指令「创建一套长方形空间 专门在上面装载方形按钮
   增加一个就排列上去一个」——节点规范.md 第五章配套唯一实现）：
   - 长方形条状空间：全宽、flex 自动排列、溢出自动换行——新按钮追加进 children 即排上
   - ★ 位置规则（2026-10-26 用户截图指正「排列有点松散」修订）：**紧跟内容行之后**
     （pn-root gap 3px 纵向节奏，卡片高度随内容）——禁弹性占位/min-height 钉底
   - 顺序语言（锁→存/导入→历史→执行）由调用方 children 顺序决定
   - 与 pm-sq-btn（31×31 圆角 6）同语言；行内图标尺寸与宿主行一致（19）；图标一律取自 src/ui/icons.tsx
   - nodrag nopan：条上起手不得拖动画布 */

/** 按钮装载条：children 顺序 = 排列顺序，加一个排一个，超宽自动换行 */
export const NodeBtnBar = memo(function NodeBtnBar ({ children, className = '' }: {
  children: ReactNode
  className?: string
}) {
  return <div className={'nb-bar nodrag nopan' + (className ? ' ' + className : '')}>{children}</div>
})

/** 标准方形钮（pm-sq-btn 封装）：active=蓝框 on 态；busy=禁点；ghost 为默认素色 */
export function NodeSqBtn ({ icon, title, onClick, active, disabled, busy, className = '' }: {
  icon: ReactNode
  title: string
  onClick?: (e: MouseEvent<HTMLButtonElement>) => void
  active?: boolean
  disabled?: boolean
  busy?: boolean
  className?: string
}) {
  const cls = [
    'pm-sq-btn',
    'pm-sq-btn--ghost',
    active ? 'pm-sq-btn--on' : '',
    className,
  ].filter(Boolean).join(' ')
  return (
    <button
      type="button"
      className={cls}
      title={title}
      aria-label={title}
      disabled={disabled || busy}
      onClick={onClick}
      draggable={false}
    >
      {icon}
    </button>
  )
}
