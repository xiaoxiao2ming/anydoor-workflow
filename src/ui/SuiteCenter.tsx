/* ─────────────────────────────────────────────────────────────
 * 套件中心（2026-10-26 用户指令「帮我做个套件中心」）
 *
 * 弹窗：套件卡片（状态徽章/能力/体积/操作按钮）+ 存储目录。
 * 入口：功能区右侧拼图按钮（CanvasRibbon）/ 节点缺套件引导行。
 * 开关走 core/suites.ts 的 suiteCenterUI 模块 store（跨组件零接线）。
 * ───────────────────────────────────────────────────────────── */
import { useState, useSyncExternalStore } from 'react'
import { createPortal } from 'react-dom'
import { SUITES, suiteCenterUI, suiteStore, isSuiteOn, setSuiteOn, refreshOcrStatus } from '../core/suites'
import { pushLog } from './logBus'
import { IcX } from './icons'

const api = (globalThis as {
  api?: {
    suiteOpenDir?: () => Promise<string>
    suiteOcrInstall?: () => Promise<{ ok: boolean; msg: string }>
    suiteOcrUninstall?: () => Promise<{ ok: boolean; msg: string }>
  }
}).api

/** OCR 套件状态徽章（0=未知 1=检测中 2=可用 3=未安装） */
type OcrSt = 0 | 1 | 2 | 3

export function SuiteCenter () {
  const open = useSyncExternalStore(suiteCenterUI.sub, suiteCenterUI.get)
  const st = useSyncExternalStore(suiteStore.sub, suiteStore.get)
  const [ocrSt, setOcrSt] = useState<OcrSt>(0)
  const [flash, setFlash] = useState('')
  const flashMsg = (m: string) => { setFlash(m); setTimeout(() => setFlash(''), 3500) }
  if (!open) return null
  const dirHint = 'userData/suites（点底部按钮打开）'

  const detectOcr = async () => {
    setOcrSt(1)
    const ok = await refreshOcrStatus()
    setOcrSt(ok ? 2 : 3)
  }
  const installOcr = async () => {
    if (!api?.suiteOcrInstall) { flashMsg('桌面桥缺失（重启应用）'); return }
    const r = await api.suiteOcrInstall()
    flashMsg(r.msg)
    if (r.ok) pushLog('[套件中心] OCR 引擎安装窗口已打开', 'ok')
  }
  const uninstallOcr = async () => {
    if (!window.confirm('卸载百度 OCR 套件？将删除套件目录（userData/suites/pdf2docx-ocr，含引擎 venv）；模型缓存（~/.paddlex）保留，重装后可复用。')) return
    if (!api?.suiteOcrUninstall) { flashMsg('桌面桥缺失（重启应用）'); return }
    const r = await api.suiteOcrUninstall()
    flashMsg(r.msg)
    setOcrSt(3)
    void refreshOcrStatus()
  }

  const badgeOf = (id: string): { cls: string; text: string } => {
    if (id === 'pdf2docx-ocr') {
      if (ocrSt === 1) return { cls: 'suite-badge--warn', text: '检测中…' }
      if (ocrSt === 2) return { cls: 'suite-badge--ok', text: '已安装 · 可用' }
      if (ocrSt === 3) return { cls: 'suite-badge--warn', text: '未安装' }
      return { cls: 'suite-badge--idle', text: st.ocrOk === true ? '已安装 · 可用' : st.ocrOk === false ? '未安装' : '待检测' }
    }
    return isSuiteOn(id) ? { cls: 'suite-badge--ok', text: '已内置 · 启用中' } : { cls: 'suite-badge--off', text: '已停用' }
  }

  return createPortal(
    <div
      className="suite-mask nodrag nopan"
      onClick={() => suiteCenterUI.close()}
      onContextMenu={(e) => { e.preventDefault(); suiteCenterUI.close() }}
    >
      <div className="suite-dialog" onClick={(e) => e.stopPropagation()}>
        <header className="suite-head">
          <b>套件中心</b>
          <span className="suite-head-sub">为节点提供能力的可插拔组件——安装/停用即改即生效，节点无需重启</span>
          <button className="suite-close" title="关闭" onClick={() => suiteCenterUI.close()}><IcX size={16} /></button>
        </header>
        <div className="suite-body">
          {SUITES.map((s) => {
            const on = isSuiteOn(s.id)
            const badge = badgeOf(s.id)
            return (
              <div key={s.id} className={'suite-card' + (on ? '' : ' suite-card--off')}>
                <div className="suite-card-top">
                  <span className={'suite-badge ' + badge.cls}>{badge.text}</span>
                  <b className="suite-name">{s.name}</b>
                  <span className="suite-kind">{s.kind === 'builtin' ? '内置' : '动态'}</span>
                </div>
                <p className="suite-desc">{s.desc}</p>
                <div className="suite-meta">
                  <span>能力：<code>{s.caps.join(' / ')}</code></span>
                  <span>体积：{s.sizeHint}</span>
                </div>
                <div className="suite-ops">
                  {s.id === 'pdf2docx-x2t' && (
                    <button
                      className={'suite-btn' + (on ? ' suite-btn--warn' : ' suite-btn--on')}
                      onClick={() => { void setSuiteOn(s.id, !on); pushLog(`[套件中心] ${s.name} 已${on ? '停用' : '启用'}`, 'ok') }}
                    >{on ? '停用' : '启用'}</button>
                  )}
                  {s.id === 'pdf2docx-ocr' && (
                    <>
                      <button className="suite-btn" onClick={() => void detectOcr()}>检测</button>
                      <button className="suite-btn suite-btn--on" title="引擎装入套件目录（userData/suites/pdf2docx-ocr/venv，不碰系统 Python 环境）" onClick={() => void installOcr()}>装入套件目录</button>
                      {(ocrSt === 2 || st.ocrOk === true) && (
                        <button className="suite-btn suite-btn--warn" onClick={() => void uninstallOcr()}>卸载</button>
                      )}
                    </>
                  )}
                </div>
              </div>
            )
          })}
        </div>
        <footer className="suite-foot">
          <span className="suite-foot-dir">套件存储目录：{dirHint}</span>
          <button
            className="suite-btn"
            onClick={() => {
              if (!api?.suiteOpenDir) { flashMsg('桌面桥缺失（重启应用）'); return }
              void api.suiteOpenDir()
            }}
          >打开目录</button>
          {flash && <span className="suite-flash">{flash}</span>}
        </footer>
      </div>
    </div>,
    document.body,
  )
}
