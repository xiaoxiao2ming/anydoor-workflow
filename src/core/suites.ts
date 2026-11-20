/* ─────────────────────────────────────────────────────────────
 * 套件中心 · 注册表与状态内核（2026-10-26）
 *
 * 规范（全文见 docs/30-套件中心与节点接入规范.md）：
 *   - 套件 = 为节点提供某种「能力」的可插拔组件（如 convert:pdf2docx）
 *   - 存储：userData/suites/（动态套件安装根目录）+ state.json（停用状态）
 *   - 两类套件：
 *     · builtin —— 随软件分发（引擎字节在 vendors，删不得）→ 「停用/启用」= 卸载语义
 *     · dynamic —— 装到 userData/suites/<id>/，卸载 = 删目录（预留扩展点）
 *   - 节点接入：节点声明所需能力（caps），**任一套件提供该能力即可执行**；
 *     全部缺失 → 缺套件态（执行钮置灰 + 卡片引导进套件中心，见规范文档）
 *
 * 单一真相源：本文件 SUITES 表。UI（SuiteCenter）与消费方（节点引擎）都只读这里。
 * ───────────────────────────────────────────────────────────── */
import { useSyncExternalStore } from 'react'

export interface SuiteDef {
  id: string
  name: string
  desc: string
  /** 提供的能力（节点按能力找套件，不按 id——同一能力可多引擎竞争） */
  caps: string[]
  /** builtin=随软件（停用/启用）；dynamic=装到 userData/suites（安装/卸载） */
  kind: 'builtin' | 'dynamic'
  sizeHint: string
}

/** 套件清单（单一真相源——新套件在此注册一行 + 在下表能力映射里登记） */
export const SUITES: SuiteDef[] = [
  {
    id: 'smart-llm-qwen',
    name: '对话助手 · 本地聪明模型（Qwen2.5-3B）',
    desc: 'Qwen2.5-3B 本地量化模型（2GB，完全离线）：对话框「聪明模式」的推理引擎——自由对话、工作流意图判定、画布节点内容分析总结。内置随软件分发；停用后对话框回落「本地快速」（bge 语义 + 规则）。',
    caps: ['llm:chat'],
    kind: 'builtin',
    sizeHint: '模型 2GB（随软件内置）',
  },
  {
    id: 'pdf2docx-x2t',
    name: 'PDF转Word · 文字版直转',
    desc: '复用软件内置 OnlyOffice 转换引擎（x2t 长驻，约 0.4s/份）：文字版 PDF 秒转可编辑 docx，版面/文本/图片搬运。扫描件（图片型 PDF）不支持——需另装 OCR 套件。',
    caps: ['convert:pdf2docx'],
    kind: 'builtin',
    sizeHint: '0 增量（复用软件内置转换引擎）',
  },
  {
    id: 'pdf2docx-ocr',
    name: 'PDF转Word · 扫描件 OCR（百度 PaddleOCR）',
    desc: '百度开源 PaddleOCR（飞桨）PP-Structure 文档解析：版面分析 + OCR + 表格识别 + 版面恢复，扫描件/图片型 PDF 也能转出可编辑 docx。★ 引擎装入套件目录（userData/suites/pdf2docx-ocr/，独立环境不碰系统 Python）——本机没有 Python 时自动下载内嵌 Python 3.12（11MB），卸载即删目录。首次转换自动下载模型（1-2GB，仅一次）。',
    caps: ['convert:pdf2docx-ocr'],
    kind: 'dynamic',
    sizeHint: '引擎约 600MB（装入套件目录）+ 模型 1-2GB（一次性）｜无需系统 Python',
  },
]

/* ── 状态存储（disabled 停用表 + OCR 引擎可用性缓存）────────────── */
let state: { disabled: string[]; ocrOk: boolean | null } = { disabled: [], ocrOk: null }
const subs = new Set<() => void>()
const emit = () => subs.forEach((f) => f())

const api = (globalThis as {
  api?: {
    suiteRead?: () => Promise<{ dir: string; disabled: string[] }>
    suiteWrite?: (patch: { disabled?: string[] }) => Promise<{ ok: boolean }>
  }
}).api

/** 初始化（App 挂载时调一次；失败不阻塞——全启用兜底） */
export async function initSuites (): Promise<void> {
  if (api?.suiteRead) {
    try {
      const r = await api.suiteRead()
      state = { ...state, disabled: r.disabled }
      emit()
    } catch { /* 桌面桥缺失 → 全启用兜底 */ }
  }
}

export const suiteStore = {
  get: () => state,
  sub: (f: () => void) => { subs.add(f); return () => { subs.delete(f) } },
}

/** 套件是否启用（builtin 停用表；dynamic 恒 true——卸载即从 SUITES 移除） */
export function isSuiteOn (id: string): boolean {
  return !state.disabled.includes(id)
}

/** 停用/启用（builtin 的卸载语义） */
export async function setSuiteOn (id: string, on: boolean): Promise<void> {
  const next = on ? state.disabled.filter((x) => x !== id) : [...new Set([...state.disabled, id])]
  state = { ...state, disabled: next }
  emit()
  if (api?.suiteWrite) { try { await api.suiteWrite({ disabled: next }) } catch { /* 状态留内存 */ } }
}

/* OCR 引擎可用性（异步检测→缓存；套件中心「检测」与节点执行共用） */
export async function refreshOcrStatus (): Promise<boolean> {
  const { detectP2wEngine } = await import('../office/pdf2wordEngine')
  const r = await detectP2wEngine()
  state = { ...state, ocrOk: r.ok }
  emit()
  return r.ok
}

/** OCR 套件是否可用（未知时视为可用——执行时还有真检测兜底，不误伤） */
export function ocrAvailable (): boolean {
  return state.ocrOk !== false
}

/* ── 能力查询（节点接入的唯一入口）────────────────────────────── */
/** 指定能力是否至少有一个启用套件提供——节点执行闸门就用它 */
export function hasCapability (cap: string): boolean {
  for (const s of SUITES) {
    if (s.caps.includes(cap) && isSuiteOn(s.id)) {
      if (s.id === 'pdf2docx-ocr' && state.ocrOk === false) continue
      return true
    }
  }
  return false
}

/* ── 套件中心弹窗开关（模块级 store——Ribbon 按钮 / 节点引导共用） */
let scOpen = false
const scSubs = new Set<() => void>()
export const suiteCenterUI = {
  get: () => scOpen,
  open: () => { scOpen = true; scSubs.forEach((f) => f()) },
  close: () => { scOpen = false; scSubs.forEach((f) => f()) },
  toggle: () => { scOpen = !scOpen; scSubs.forEach((f) => f()) },
  sub: (f: () => void) => { scSubs.add(f); return () => { scSubs.delete(f) } },
}

/** React 订阅便捷钩子 */
export function useSuiteState () {
  return useSyncExternalStore(suiteStore.sub, suiteStore.get)
}
