import { useState, type ReactNode } from 'react'

/**
 * WPS 风格功能区（Ribbon）—— Excel / Word 编辑器共用外壳
 *
 * 结构（对齐 WPS）：
 *   行1：[快捷区(文件/保存…)] [页签：开始 插入 页面…（点当前页签=收起/展开功能行）] [弹性] [右侧(最大化…)]
 *   行2：功能组（每组底部有组名，组间竖分隔线）
 *
 * 一致性铁律：两个面板都用本组件，不允许各自再手搓菜单栏。
 */

export type RibbonItem =
  | {
      kind: 'btn'
      label: string
      /** 大图标（显示在 label 上方）；不传则渲染纯文字小按钮 */
      glyph?: string
      /** 图标内联样式（WPS 的 B/I/U/S 需要各自字形：粗体/斜体/下划线/删除线） */
      glyphStyle?: React.CSSProperties
      /** disabled 时可省略 */
      onClick?: () => void
      disabled?: boolean
      title?: string
      /** 当前处于激活态（如选区已加粗）—— 渲染 WPS 那种按下灰底 */
      active?: boolean
    }
  | {
      kind: 'select'
      value: string
      options: { value: string; label: string }[]
      /** disabled 时可省略 */
      onChange?: (v: string) => void
      disabled?: boolean
      title?: string
      width?: number
    }
  | 'sep'

export type RibbonGroup = { header: string; items: RibbonItem[] }
export type RibbonTab = { label: string; groups: RibbonGroup[] }

export function Ribbon(props: {
  tabs: RibbonTab[]
  /** 行1 左侧快捷区（如「文件」按钮） */
  quick?: ReactNode
  /** 行1 右侧（如最大化） */
  right?: ReactNode
  initialTab?: string
  /** 功能行默认收起（表格等矮面板用） */
  collapsedByDefault?: boolean
}) {
  const [active, setActive] = useState(props.initialTab ?? props.tabs[0]?.label ?? '')
  const [collapsed, setCollapsed] = useState(!!props.collapsedByDefault)

  const toggleTab = (label: string) => {
    if (active === label) setCollapsed((c) => !c)
    else {
      setActive(label)
      setCollapsed(false)
    }
  }

  const tab = props.tabs.find((t) => t.label === active)

  return (
    <div className="rb">
      {/* 行1：快捷区 + 页签 + 右侧 */}
      <div className="rb-row1">
        {props.quick}
        <div className="rb-tabs">
          {props.tabs.map((t) => (
            <button
              key={t.label}
              className={`rb-tab${t.label === active ? ' on' : ''}`}
              onClick={() => toggleTab(t.label)}
            >
              {t.label}
            </button>
          ))}
        </div>
        <span style={{ flex: 1 }} />
        {props.right}
        <button className="rb-collapse" title={collapsed ? '展开功能区' : '收起功能区'} onClick={() => setCollapsed((c) => !c)}>
          {collapsed ? '⌃' : '⌄'}
        </button>
      </div>

      {/* 行2：功能组 */}
      {!collapsed && tab && (
        <div className="rb-body">
          {tab.groups.map((g, gi) => (
            <div key={g.header + gi} className="rb-group">
              <div className="rb-group-items">
                {g.items.map((it, i) => {
                  if (it === 'sep') return <div key={i} className="rb-vsep" />
                  if (it.kind === 'select')
                    return (
                      <select
                        key={i}
                        className="rb-sel"
                        value={it.value}
                        disabled={it.disabled}
                        title={it.title}
                        style={it.width ? { width: it.width } : undefined}
                        onChange={(e) => it.onChange?.(e.target.value)}
                      >
                        {it.options.map((o) => (
                          <option key={o.value} value={o.value}>
                            {o.label}
                          </option>
                        ))}
                      </select>
                    )
                  return (
                    <button
                      key={i}
                      className={`rb-btn${it.glyph ? ' has-glyph' : ''}${it.disabled ? ' disabled' : ''}${it.active ? ' on' : ''}`}
                      disabled={it.disabled}
                      title={it.title}
                      onClick={() => it.onClick?.()}
                    >
                      {it.glyph && (
                        <span className="rb-glyph" style={it.glyphStyle}>
                          {it.glyph}
                        </span>
                      )}
                      {it.label ? <span className="rb-label">{it.label}</span> : null}
                    </button>
                  )
                })}
              </div>
              <div className="rb-group-header">{g.header}</div>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}
