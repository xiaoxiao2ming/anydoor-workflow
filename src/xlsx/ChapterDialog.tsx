/**
 * 章节编排窗体 —— 目录大纲的辅助操作面板
 *
 * 形态口径（用户 2026-10-04 定）：**与「批注清单」窗体完全一致**。
 *   批注窗体用的是`position:fixed; inset:0` 全屏遮罩 + 居中卡片，
 *   天然不受宿主容器尺寸影响 → **最大化态与普通视口表现同一套**，不写两套分支。
 *   本组件照抄那一套（同样的 z-index var(--z-modal)、同样的头/体/脚三段、同样圆角 10）。
 *
 * 数据层：树与表格是同一份 OutlineRow[] 的两个渲染（方案铁律 2）。
 *   本组件**不持有**结构真相 —— rows 由宿主（useOutlineSlot）从工作表派生，
 *   所有变更走 outline.ts 的纯函数，再由宿主统一 commitStructure + 写盘 + 落盘。
 *
 * 纯计算层铁律：outline.ts 不 import electron / node: / document（UI 层不受此限）。
 */

import { useMemo, useState } from 'react'
import {
  buildTree,
  commitStructure,
  countTitles,
  flattenTree,
  setCollapseAll,
  shiftLevel,
  toggleCollapse,
  visibleRowIndexes,
  type FlatNode,
  type OutlineRow,
  type RowLevel,
} from './outline'

export interface ChapterDialogProps {
  rows: OutlineRow[]
  /**
   * 结构变更收口：宿主在这里 commitStructure + 写回 A/B/C 列 + 落盘。
   * @param msg 给用户的提示（空串/不传 = 不提示）。钳制信息由宿主自行判断后回传。
   */
  onCommit: (msg?: string) => void
  /** 关闭窗体 */
  onClose: () => void
  /** 点击某章节 → 定位到表格对应行（宿主负责滚动 + 闪烁高亮） */
  onGoto?: (rowIndex: number, row: OutlineRow) => void
  /** 当前编辑器的行数组引用是否可变（false 时全部操作只读） */
  readOnly?: boolean
}

/** 层级 → 左侧缩进 px（每级 14px，与 Word 大纲观感接近） */
const INDENT_PX = 14

export function ChapterDialog({ rows, onCommit, onClose, onGoto, readOnly }: ChapterDialogProps) {
  // 树的展开/折叠状态存在 row.collapsed 上（跟着行走），此处只做派生
  const [, forceRender] = useState(0)
  const [filter, setFilter] = useState('')

  const tree = useMemo(() => buildTree(rows), [rows])
  const flat = useMemo(() => flattenTree(tree, rows), [tree, rows])
  const totalTitles = useMemo(() => countTitles(tree), [tree])
  const visibleCount = useMemo(() => visibleRowIndexes(rows).length, [rows])

  /**
   * 搜索过滤：命中标题时**保留整条祖先路径**（否则命中项会因父级被折叠而看不见）。
   *
   * 祖先链用「逐个找父」的方式向上走（父 = 往前第一个 depth 更小的节点，
   * 这是扁平数组的天然性质），不改动任何原始对象。
   */
  const shown = useMemo(() => {
    const q = filter.trim().toLowerCase()
    if (!q) return flat
    const keep = new Set<string>()
    const parentOf = (i: number): number => {
      const d = flat[i].depth
      for (let k = i - 1; k >= 0; k--) if (flat[k].depth < d) return k
      return -1
    }
    for (let i = 0; i < flat.length; i++) {
      if (!`${flat[i].no} ${flat[i].title}`.toLowerCase().includes(q)) continue
      keep.add(flat[i].rowId)
      for (let k = parentOf(i); k >= 0; k = parentOf(k)) keep.add(flat[k].rowId)
    }
    return flat.filter((n) => keep.has(n.rowId))
  }, [flat, filter])

  const doToggle = (rowIndex: number) => {
    if (readOnly) return
    toggleCollapse(rows, rowIndex)
    forceRender((v) => v + 1)
  }

  const doShift = (rowIndex: number, delta: 1 | -1) => {
    if (readOnly) return
    const r = shiftLevel(rows, rowIndex, delta)
    if (r.moved === 0) return
    onCommit(r.clamped ? '已自动修正跳级标题' : undefined)
    forceRender((v) => v + 1)
  }

  const doCollapseAll = (collapsed: boolean) => {
    if (readOnly) return
    setCollapseAll(rows, collapsed)
    forceRender((v) => v + 1)
  }

  return (
    <div
      style={{
        position: 'fixed',
        inset: 0,
        zIndex: 'var(--z-modal)',
        background: 'rgba(1,4,9,.6)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
      }}
      onPointerDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        className="wb-cmtlist wb-chapters"
        style={{
          width: 560,
          maxHeight: 600,
          display: 'flex',
          flexDirection: 'column',
          background: 'var(--ed-bg0)',
          border: '1px solid var(--ed-border)',
          borderRadius: 10,
          overflow: 'hidden',
        }}
      >
        {/* ── 头：标题 + 搜索 + 折叠全部 ── */}
        <div
          style={{
            padding: '12px 16px',
            borderBottom: '1px solid #2a323d',
            fontSize: 14,
            color: 'var(--ed-text)',
            display: 'flex',
            alignItems: 'center',
            gap: 10,
          }}
        >
          <span>章节编排（{totalTitles} 章 / 可见 {visibleCount} 行）</span>
          <span style={{ flex: 1 }} />
          <button className="bar-btn" onClick={() => doCollapseAll(true)} disabled={readOnly}>
            折叠全部
          </button>
          <button className="bar-btn" onClick={() => doCollapseAll(false)} disabled={readOnly}>
            展开全部
          </button>
          <button className="ne-close" onClick={onClose} title="关闭">
            ×
          </button>
        </div>

        {/* 搜索条 */}
        <div style={{ padding: '8px 16px', borderBottom: '1px solid var(--ed-bg3)' }}>
          <input
            value={filter}
            onChange={(e) => setFilter(e.target.value)}
            placeholder="按章节号或标题筛选…"
            style={{
              width: '100%',
              boxSizing: 'border-box',
              fontSize: 12.5,
              color: 'var(--ed-text)',
              background: 'var(--ed-bg2)',
              border: '1px solid var(--ed-border2)',
              borderRadius: 6,
              padding: '6px 10px',
            }}
          />
        </div>

        {/* ── 体：树 ── */}
        <div style={{ flex: 1, overflowY: 'auto', minHeight: 120 }}>
          {totalTitles === 0 && (
            <div style={{ padding: 20, fontSize: 12, color: '#9aa7b4' }}>
              这张表还没有章节标题。在标记列（最左侧）填 1 / 2 / 3 即可把该行变成一级 / 二级 / 三级标题，
              章节号会自动重排。
            </div>
          )}
          {totalTitles > 0 && shown.length === 0 && (
            <div style={{ padding: 20, fontSize: 12, color: '#9aa7b4' }}>没有匹配的章节。</div>
          )}
          {shown.map((n) => {
            const r = rows[n.rowIndex]
            if (!r) return null
            const hasBody = n.bodyCount > 0
            return (
              <div
                key={n.rowId}
                className="wb-cmt-row wb-chapter-row"
                style={{
                  padding: '6px 16px',
                  borderBottom: '1px solid var(--ed-bg3)',
                  display: 'flex',
                  alignItems: 'center',
                  gap: 8,
                  paddingLeft: 16 + n.depth * INDENT_PX,
                }}
              >
                {/* 折叠钮：可展开才显示，无子级时留空位保持缩进对齐 */}
                {n.hasChildren ? (
                  <button
                    className="ne-iconbtn"
                    style={{ width: 18, height: 18, fontSize: 10 }}
                    title={n.collapsed ? '展开' : '折叠'}
                    onClick={() => doToggle(n.rowIndex)}
                  >
                    {n.collapsed ? '▸' : '▾'}
                  </button>
                ) : (
                  <span style={{ width: 18, flex: '0 0 18px' }} />
                )}

                {/* 章节号 */}
                <span
                  style={{
                    fontSize: 11.5,
                    color: 'var(--ed-gold)',
                    fontFamily: 'Consolas, monospace',
                    minWidth: 44,
                  }}
                >
                  {n.no}
                </span>

                {/* 标题 */}
                <span
                  style={{
                    fontSize: 12.5,
                    color: 'var(--ed-text)',
                    fontWeight: n.level <= 1 ? 600 : 400,
                    flex: 1,
                    cursor: onGoto ? 'pointer' : 'default',
                    overflow: 'hidden',
                    textOverflow: 'ellipsis',
                    whiteSpace: 'nowrap',
                  }}
                  onClick={() => onGoto?.(n.rowIndex, r)}
                  title={onGoto ? '点击定位到表格对应行' : undefined}
                >
                  {n.title || <em style={{ color: 'var(--ed-text4)' }}>（无标题）</em>}
                </span>

                {/* 直属普通数据行数 */}
                {hasBody && (
                  <span style={{ fontSize: 10.5, color: 'var(--ed-text4)' }} title="本章节下的数据行数">
                    {n.bodyCount} 行
                  </span>
                )}

                {/* 升降级 */}
                <button
                  className="ne-iconbtn"
                  style={{ width: 20, height: 20, fontSize: 11 }}
                  title="升级（降一级）"
                  disabled={readOnly || r.level >= 5}
                  onClick={() => doShift(n.rowIndex, 1)}
                >
                  ←
                </button>
                <button
                  className="ne-iconbtn"
                  style={{ width: 20, height: 20, fontSize: 11 }}
                  title="降级（升一级）"
                  disabled={readOnly || r.level <= 1}
                  onClick={() => doShift(n.rowIndex, -1)}
                >
                  →
                </button>
              </div>
            )
          })}
        </div>

        {/* ── 脚：说明 + 关闭 ── */}
        <div style={{ padding: '8px 16px', borderTop: '1px solid #2a323d' }}>
          <div style={{ fontSize: 11, color: '#9aa7b4', marginBottom: 6 }}>
            ← 升级（更浅） · → 降级（更深） · 章节号是派生数据，改层级后自动重排，勿手工填写
          </div>
          <button className="bar-btn" onClick={onClose}>
            关闭
          </button>
        </div>
      </div>
    </div>
  )
}
