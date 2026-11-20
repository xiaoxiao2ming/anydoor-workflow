/**
 * .wfc 工程单文件打开（2026-10-07 用户：「双击*.wfc文件 打开软件并加载该文件内容」）
 *
 * ── 与 wfcSave.ts 严格对称 ──
 *   保存：快照 → saveProjectFull（IDB 工程区）→ packProject（zip）→ 下载
 *   打开：**读文件 → unpackProject（zip → IDB 工程区）→ openProjectFull → 灌 store**
 *   逆过程复用同一条 .cbproj 链路，不存在第二套格式。
 *
 * ── 打开语义 ──
 *   .wfc = 单画布工程（保存口径是「当前画布」）。
 *   打开 = 作为**新画布页**导入当前工程（store.importPage），不覆盖已有页 ——
 *   双击文件进来的人，多半不想毁掉他正开着的其他画布。
 *
 * ── 分层位置 ──
 *   与 wfcSave.ts 同层（src/ 根）：它只碰浏览器 File / 弹窗，不进 core。
 */

import { useBaseStore } from './core/store'
import { openProjectFull, applyTplIdMap } from './core/projectIO'
import { createIdbProjectFs, type ProjectRef } from './core/projectStore'
import { unpackProject } from './core/projectPack'
import { browserKV } from './core/viewState'
import { makeBrowserAssetRuntime, WFC_WORKSPACE_ROOT } from './wfcSave'

/** IDB 工程区根目录（打开用独立目录，不与导出暂存区互相覆盖） */
export const WFC_OPEN_ROOT = 'wfc-open'

export interface WfcOpenResult {
  /** 工程显示名（manifest.name，也是新画布页名） */
  name: string
  /** 导入的节点数 */
  nodes: number
  /** 新画布页 id */
  pageId: string
  /** 资产层面问题（模板缺失等）——不阻断打开，但必须让用户看到 */
  problems: string[]
}

/**
 * 从 .wfc 文件字节打开工程：解包 → 灌资产 → 导入为新画布页并切换。
 *
 * @param bytes  .wfc 文件字节（zip：project.json + nodes/ + assets/ + …）
 * @param fallbackName 文件名主体（manifest.name 缺失时的页名兜底）
 * @throws zip 解析失败/缺 project.json 时原样抛出（调用方弹错误提示）
 */
export async function openCanvasFromWfcBytes(
  bytes: Uint8Array,
  fallbackName = '导入的画布',
  onStep?: (st: { label: string; pct: number }) => void,
): Promise<WfcOpenResult> {
  const fs = createIdbProjectFs()
  const ref: ProjectRef = { rootDir: WFC_OPEN_ROOT }

  // ① zip → IDB 工程目录（unpackProject 自带路径越界拦截与 project.json 检查）
  onStep?.({ label: '写入工程目录…', pct: 50 })
  const up = await unpackProject(fs, WFC_OPEN_ROOT, bytes)
  if (!up.written.length && up.problems.length) {
    throw new Error(up.problems.join('\n'))
  }

  // ② 读工程 + 资产灌回模板库/编辑稿 + 控件状态回灌 localStorage
  onStep?.({ label: '恢复节点与资产…', pct: 75 })
  const out = await openProjectFull(fs, ref, {
    runtime: makeBrowserAssetRuntime(),
    kv: browserKV(),
  })

  // ③ 模板 id 改写（工程里的 tplId 与本机模板库的 id 不同 ⇒ rewire 保持绑定有效）
  const tables = applyTplIdMap(out.tables, out.tplIdMap)

  // ④ 导入为新画布页并切换（不覆盖用户已开的画布）
  const s = useBaseStore.getState()
  const pageName = out.manifest.name || fallbackName
  const pageId = await s.importPage(pageName, tables, out.records, out.positions)

  return {
    name: pageName,
    nodes: Object.keys(tables).filter((id) => !tables[id].deleted).length,
    pageId,
    problems: out.assetProblems,
  }
}

/**
 * 从浏览器 File 对象打开（「📂 打开」按钮 / input[type=file] 路径用）。
 * 独立出来是因为命令行路径（Electron）与 File 对象（弹窗）两个入口共用后续逻辑。
 */
export async function openCanvasFromWfcFile(file: File, onStep?: (st: { label: string; pct: number }) => void): Promise<WfcOpenResult> {
  onStep?.({ label: '读取文件…', pct: 10 })
  const buf = await file.arrayBuffer()
  const bytes = new Uint8Array(buf)
  const fallback = file.name.replace(/\.wfc$/i, '')
  onStep?.({ label: '解压工程数据…', pct: 30 })
  const r = await openCanvasFromWfcBytes(bytes, fallback, onStep)
  onStep?.({ label: '载入完成', pct: 100 })
  return r
}
