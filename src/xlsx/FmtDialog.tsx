/**
 * 设置单元格格式对话框（#33 右键菜单 v2 第二块）—— WPS/Excel 同款四页签：字体 / 边框 / 填充 / 数字。
 *
 * 形态口径：与「章节编排/批注清单」同一套 fixed+inset:0 遮罩 + 居中卡片 + --z-modal
 * （最大化态与普通视口表现一致，不写第二套分支）。
 * 数据流：纯受控组件——初值由调用方 fmtOfFirst 读首格，本组件只改本地副本；
 * 确定时回传完整 FmtState，由调用方 applyFmtSel 四轴一次写入选区（补丁 xf 合并）。
 * 读首格语义（WPS 同款）：混合选区显示首格格式；对话框未建模的属性（vertAlign/scheme/
 * 渐变填充/斜边框）确定后会丢——已在 editOps.applyFmtSel 注释里声明。
 */
import { useEffect, useState } from 'react'
import type { FmtState, FmtSideState } from './editOps'

/** 边框线型全集（kit 的 SIDE_STYLES 未从 styles 主入口导出，这里本地同款清单） */
const SIDE_STYLES = [
  'thin', 'medium', 'thick', 'double', 'hair', 'dotted', 'dashed',
  'dashDot', 'dashDotDot', 'mediumDashed', 'mediumDashDot', 'mediumDashDotDot', 'slantDashDot',
] as const

export interface FmtDialogProps {
  init: FmtState
  readOnly?: boolean
  /** 确定按钮：回传编辑后的完整四轴状态 */
  onOk: (st: FmtState) => void
  onCancel: () => void
}

type Tab = 'font' | 'border' | 'fill' | 'num'

const TABS: { id: Tab; label: string }[] = [
  { id: 'font', label: '字体' },
  { id: 'border', label: '边框' },
  { id: 'fill', label: '填充' },
  { id: 'num', label: '数字' },
]

/** 边框线型中文标签（SIDE_STYLES 顺序展示用） */
const SIDE_LABEL: Record<string, string> = {
  thin: '细线',
  medium: '中线',
  thick: '粗线',
  double: '双线',
  hair: '发丝线',
  dotted: '点线',
  dashed: '虚线',
  dashDot: '点划线',
  dashDotDot: '双点划线',
  mediumDashed: '中虚线',
  mediumDashDot: '中点划线',
  mediumDashDotDot: '中双点划线',
  slantDashDot: '斜点划线',
}

/** 填充色板（工程表格常用色，探针点 data-swatch 驱动） */
const SWATCHES = [
  '#FFFF00', '#FF0000', '#00B050', '#0070C0', '#FFC000', '#7030A0',
  '#FFFFFF', '#D9D9D9', '#F2F2F2', '#FFF2CC', '#E2EFDA', '#DDEBF7',
]

/** 数字格式预设（覆盖工程量/造价场景高频格式） */
const NUM_PRESETS: { code: string; label: string }[] = [
  { code: 'General', label: '常规' },
  { code: '0', label: '数值' },
  { code: '0.00', label: '数值（两位小数）' },
  { code: '#,##0', label: '千分位' },
  { code: '#,##0.00', label: '千分位（两位小数）' },
  { code: '0%', label: '百分比' },
  { code: '0.00%', label: '百分比（两位小数）' },
  { code: '¥#,##0.00', label: '货币' },
  { code: 'yyyy-mm-dd', label: '日期' },
  { code: '@', label: '文本' },
]

const SIDE_KEYS = ['top', 'bottom', 'left', 'right'] as const
const SIDE_NAME: Record<(typeof SIDE_KEYS)[number], string> = {
  top: '上边框',
  bottom: '下边框',
  left: '左边框',
  right: '右边框',
}

export function FmtDialog({ init, readOnly, onOk, onCancel }: FmtDialogProps) {
  const [st, setSt] = useState<FmtState>(() => JSON.parse(JSON.stringify(init)) as FmtState)
  const [tab, setTab] = useState<Tab>('font')

  // Escape=取消 / Enter=确定（对话框通用键位；与批注清单/章节窗体一致）
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onCancel()
      else if (e.key === 'Enter') onOk(st)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [st, onOk, onCancel])

  const setFont = (patch: Partial<FmtState['font']>) => setSt((s) => ({ ...s, font: { ...s.font, ...patch } }))
  const setSide = (k: (typeof SIDE_KEYS)[number], patch: Partial<FmtSideState>) =>
    setSt((s) => ({ ...s, border: { ...s.border, [k]: { ...s.border[k], ...patch } } }))

  return (
    <div className="wb-fmtdlg" data-fmt-dlg="1" onPointerDown={(e) => e.target === e.currentTarget && onCancel()}>
      <div className="wb-fmtdlg-box">
        <div className="wb-fmtdlg-head">
          <span>设置单元格格式</span>
          <button className="wb-fmtdlg-x" title="关闭" onClick={onCancel}>
            ×
          </button>
        </div>
        <div className="wb-fmtdlg-tabs">
          {TABS.map((t) => (
            <button key={t.id} data-fmt-tab={t.id} className={tab === t.id ? 'on' : ''} onClick={() => setTab(t.id)}>
              {t.label}
            </button>
          ))}
        </div>

        <div className="wb-fmtdlg-body">
          {tab === 'font' && (
            <>
              <div className="wb-fmtdlg-row">
                <label>字体名</label>
                <input data-fmt="fontname" value={st.font.name} readOnly={readOnly}
                  onChange={(e) => setFont({ name: e.target.value })} />
              </div>
              <div className="wb-fmtdlg-row">
                <label>字号</label>
                <input data-fmt="fontsize" type="number" min={6} max={72} value={st.font.size} readOnly={readOnly}
                  onChange={(e) => setFont({ size: Math.max(6, Math.min(72, Number(e.target.value) || 11)) })} />
                <label className="wb-fmtdlg-gap">字体颜色</label>
                <input data-fmt="fontcolor" type="color" value={st.font.color ?? '#000000'} disabled={readOnly}
                  onChange={(e) => setFont({ color: e.target.value })} />
                <button data-fmt="fontcolor-auto" onClick={() => setFont({ color: null })}>自动</button>
              </div>
              <div className="wb-fmtdlg-row">
                <label><input data-fmt="bold" type="checkbox" checked={st.font.bold} disabled={readOnly}
                  onChange={(e) => setFont({ bold: e.target.checked })} /> 加粗</label>
                <label><input data-fmt="italic" type="checkbox" checked={st.font.italic} disabled={readOnly}
                  onChange={(e) => setFont({ italic: e.target.checked })} /> 斜体</label>
                <label><input data-fmt="underline" type="checkbox" checked={st.font.underline} disabled={readOnly}
                  onChange={(e) => setFont({ underline: e.target.checked })} /> 下划线</label>
                <label><input data-fmt="strike" type="checkbox" checked={st.font.strike} disabled={readOnly}
                  onChange={(e) => setFont({ strike: e.target.checked })} /> 删除线</label>
              </div>
            </>
          )}

          {tab === 'border' && (
            <>
              {SIDE_KEYS.map((k) => (
                <div className="wb-fmtdlg-row" key={k}>
                  <label>{SIDE_NAME[k]}</label>
                  <select data-fmt={'bd-' + k} value={st.border[k].style} disabled={readOnly}
                    onChange={(e) => setSide(k, { style: e.target.value as FmtSideState['style'] })}>
                    <option value="none">无</option>
                    {SIDE_STYLES.map((s) => (
                      <option key={s} value={s}>{SIDE_LABEL[s] ?? s}</option>
                    ))}
                  </select>
                  <input data-fmt={'bdc-' + k} type="color" value={st.border[k].color ?? '#000000'}
                    disabled={readOnly || st.border[k].style === 'none'}
                    onChange={(e) => setSide(k, { color: e.target.value })} />
                </div>
              ))}
              <div className="wb-fmtdlg-row">
                <button data-fmt="bd-all" disabled={readOnly}
                  onClick={() => setSt((s) => {
                    const c = s.border.top.color ?? '#000000'
                    const all = { style: 'thin' as const, color: c }
                    return { ...s, border: { top: all, bottom: all, left: all, right: all } }
                  })}>四边细线</button>
                <button data-fmt="bd-clear" disabled={readOnly}
                  onClick={() => setSt((s) => {
                    const none = { style: 'none' as const, color: null }
                    return { ...s, border: { top: none, bottom: none, left: none, right: none } }
                  })}>清除边框</button>
              </div>
            </>
          )}

          {tab === 'fill' && (
            <>
              <div className="wb-fmtdlg-row">
                <label>填充颜色</label>
                <input data-fmt="fill" type="color" value={st.fill ?? '#FFFFFF'} disabled={readOnly}
                  onChange={(e) => setSt((s) => ({ ...s, fill: e.target.value }))} />
                <button data-fmt="fill-clear" onClick={() => setSt((s) => ({ ...s, fill: null }))}>无填充</button>
              </div>
              <div className="wb-fmtdlg-swatches">
                {SWATCHES.map((c) => (
                  <button key={c} data-swatch={c} title={c} style={{ background: c }} disabled={readOnly}
                    onClick={() => setSt((s) => ({ ...s, fill: c }))} />
                ))}
              </div>
            </>
          )}

          {tab === 'num' && (
            <>
              {NUM_PRESETS.map((p) => (
                <label className="wb-fmtdlg-radio" key={p.code}>
                  <input data-numpreset={p.code} type="radio" name="wb-fmtnum" checked={st.numFmt === p.code}
                    disabled={readOnly} onChange={() => setSt((s) => ({ ...s, numFmt: p.code }))} />
                  {p.label}（{p.code}）
                </label>
              ))}
              <div className="wb-fmtdlg-row">
                <label>自定义</label>
                <input data-fmt="numfmt" value={st.numFmt} readOnly={readOnly}
                  onChange={(e) => setSt((s) => ({ ...s, numFmt: e.target.value }))} />
              </div>
            </>
          )}
        </div>

        <div className="wb-fmtdlg-foot">
          <button data-fmt="ok" className="wb-fmtdlg-ok" disabled={readOnly} onClick={() => onOk(st)}>确定</button>
          <button data-fmt="cancel" onClick={onCancel}>取消</button>
        </div>
      </div>
    </div>
  )
}
