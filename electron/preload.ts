import { contextBridge, ipcRenderer } from 'electron'

export interface TemplateMeta {
  id: string
  name: string
  size: number
  updatedAt: number
}

// 暴露给渲染层的持久化通道：读写 userData/client-base.sqlite
// 渲染层 sql.js 导出字节 → writeDb 落盘；启动时 readDb 取回字节 hydrate
contextBridge.exposeInMainWorld('api', {
  readDb: (): Promise<Uint8Array | null> => ipcRenderer.invoke('db:read'),
  writeDb: (data: Uint8Array): Promise<void> => ipcRenderer.invoke('db:write', data),

  // ── 本地聪明模型（GGUF + node-llama-cpp，主进程懒加载；opts.model 可指定模型文件） ──────
  llmChat: (prompt: string, opts?: { model?: string; noThink?: boolean; maxTokens?: number; temperature?: number; tier?: 'light' | 'heavy' }): Promise<{ ok: boolean; text?: string; reason?: string }> =>
    ipcRenderer.invoke('llm:chat', prompt, opts),
  // ★ 网址监控（2026-11-09）：主进程 fetch 绕 CORS + cheerio/JSON 提取 ──与 preload.cjs 同步 ──
  webmonFetch: (req: { url: string; mode: 'text' | 'json' | 'keyword'; selector?: string; keyword?: string; jsonPath?: string }): Promise<{ ok: boolean; value?: string; error?: string }> =>
    ipcRenderer.invoke('webmon:fetch', req),
  // ★ 流式版：onChunk 边生成边回调（kind: 'text'=正文 | 'thought'=思考段；llm:chunk 事件按 streamId 过滤），resolve 时给完整正文
  llmChatStream: (
    prompt: string,
    opts: { model?: string; noThink?: boolean; maxTokens?: number; temperature?: number; tier?: 'light' | 'heavy' },
    onChunk: (part: { kind: 'text' | 'thought'; text: string }) => void,
  ): Promise<{ ok: boolean; text?: string; reason?: string }> => {
    const streamId = 's' + Math.random().toString(36).slice(2)
    const listener = (_ev: unknown, p: { streamId: string; kind: 'text' | 'thought'; text: string }): void => {
      if (p.streamId === streamId) onChunk({ kind: p.kind, text: p.text })
    }
    ipcRenderer.on('llm:chunk', listener)
    return ipcRenderer
      .invoke('llm:chat-stream', prompt, { ...opts, streamId })
      .finally(() => ipcRenderer.removeListener('llm:chunk', listener))
  },

  // ── 联网搜索（Bing 抓取，返回 [{title,url,snippet}]）与打开引用链接 ──────
  webSearch: (query: string): Promise<{ ok: boolean; items?: Array<{ title: string; url: string; snippet: string }>; reason?: string }> =>
    ipcRenderer.invoke('web:search', query),
  openUrl: (url: string): Promise<{ ok: boolean }> => ipcRenderer.invoke('shell:openUrl', url),

  // ── Word 模板库（userData/templates） ───────────────────────────
  tplList: (): Promise<TemplateMeta[]> => ipcRenderer.invoke('tpl:list'),
  tplSave: (name: string, data: Uint8Array): Promise<TemplateMeta> =>
    ipcRenderer.invoke('tpl:save', name, data),
  tplRead: (id: string): Promise<Uint8Array> => ipcRenderer.invoke('tpl:read', id),
  tplDelete: (id: string): Promise<void> => ipcRenderer.invoke('tpl:delete', id),
  /** 打开系统文件对话框挑模板（kind: 'docx' | 'xlsx'，默认 docx），取消返回 null */
  tplPick: (kind?: 'docx' | 'xlsx'): Promise<{ name: string; data: Uint8Array } | null> =>
    ipcRenderer.invoke('tpl:pick', kind ?? 'docx'),
  /** 产物另存为（kind 决定对话框标题与扩展名）；取消返回 null */
  gpuGetAccel: (): Promise<boolean> => ipcRenderer.invoke('gpu:getAccel'),
  showRemoteDialog: (): Promise<boolean> => ipcRenderer.invoke('remote:showDialog'),
  pdfDiffView: (a: { name: string; bytes: Uint8Array }, b: { name: string; bytes: Uint8Array }): Promise<{ ok: boolean; reason?: string }> => ipcRenderer.invoke('pdf:diffView', a, b),
  folderPickDir: (): Promise<string | null> => ipcRenderer.invoke('folder:pickDir'),
  folderList: (dir: string): Promise<{ ok: boolean; files: { name: string; size: number; isDir: boolean; mtime: number }[]; reason?: string }> => ipcRenderer.invoke('folder:list', dir),
  folderReadFile: (dir: string, name: string): Promise<{ ok: boolean; bytes: Uint8Array; reason?: string }> => ipcRenderer.invoke('folder:readFile', dir, name),
  projectNativePack: (files: { path: string; bytes: Uint8Array }[], dirs: string[], savePath?: string): Promise<{ bytes: Uint8Array | null; savedTo: string | null }> => ipcRenderer.invoke('project:nativePack', files, dirs, savePath),
  projectNativeUnpack: (zipBytes: Uint8Array): Promise<{ ok: boolean; files: Record<string, Uint8Array>; reason?: string }> => ipcRenderer.invoke('project:nativeUnpack', zipBytes),
  wfcPickSavePath: (defaultName: string): Promise<string | null> => ipcRenderer.invoke('wfc:pickSavePath', defaultName),
  wfcWriteFile: (filePath: string, bytes: Uint8Array): Promise<{ ok: boolean; size: number }> => ipcRenderer.invoke('wfc:writeFile', filePath, bytes),
  /** PDF转Word：本机 PaddleOCR 引擎通道（2026-10-26）——工作目录 + python 检测/转换 */
  p2wTmpdir: (): Promise<{ dir: string }> => ipcRenderer.invoke('p2w:tmpdir'),
  p2wExec: (payload: { mode: 'detect' | 'convert'; input?: string; output?: string; lang?: string }): Promise<{ ok: boolean; code: number; stdout: string; stderr: string; exe: string; bytes?: Uint8Array }> =>
    ipcRenderer.invoke('p2w:exec', payload),
  /** 套件中心（2026-10-26）：状态读写 / 打开存储目录 / OCR 套件动态安装卸载 */
  suiteRead: (): Promise<{ dir: string; disabled: string[] }> => ipcRenderer.invoke('suite:read'),
  suiteWrite: (patch: { disabled?: string[] }): Promise<{ ok: boolean }> => ipcRenderer.invoke('suite:write', patch),
  suiteOpenDir: (): Promise<string> => ipcRenderer.invoke('suite:openDir'),
  suiteOcrInstall: (): Promise<{ ok: boolean; msg: string }> => ipcRenderer.invoke('suite:ocrInstall'),
  suiteOcrUninstall: (): Promise<{ ok: boolean; msg: string }> => ipcRenderer.invoke('suite:ocrUninstall'),
  assetsExport: (items: { name: string; bytesB64: string }[]): Promise<{ ok: boolean; dir?: string; count?: number; reason?: string }> => ipcRenderer.invoke('assets:export', items),
  gpuSetAccel: (on: boolean): Promise<boolean> => ipcRenderer.invoke('gpu:setAccel', on),
  openExternalFile: (name: string, data: Uint8Array): Promise<string> =>
    ipcRenderer.invoke('shell:openExternalFile', name, data),
  /** 用本机浏览器打开（Edge/Chrome 自带 PDF 查看器；2026-10-26） */
  openInBrowser: (name: string, data: Uint8Array): Promise<string> =>
    ipcRenderer.invoke('shell:openInBrowser', name, data),
  docSaveAs: (defaultName: string, data: Uint8Array, kind?: 'docx' | 'xlsx'): Promise<string | null> =>
    ipcRenderer.invoke('doc:saveAs', defaultName, data, kind ?? 'docx'),
  /** 完整 HTML → PDF（隐藏窗口打印），返回保存路径；取消返回 null */
  docPdf: (html: string, defaultName: string): Promise<string | null> =>
    ipcRenderer.invoke('doc:pdf', html, defaultName),
  /**
   * 记录一次「打不开的文件」样本（元数据 + 原始字节），返回落盘目录。
   * 文档解析失败时调用 —— 归一化层要补什么，只能靠真实样本回答。
   */
  failSave: (payload: unknown, bytes: Uint8Array | null): Promise<string> =>
    ipcRenderer.invoke('fail:save', payload, bytes),
  /** 列出已记录的失败样本（诊断用） */
  failList: (): Promise<unknown[]> => ipcRenderer.invoke('fail:list'),

  // ── 操作审计日志（2026-10-17 用户指令：点击/键盘/反馈全记录）───────
  /** 批量追加已格式化的日志行 → logs/audit-YYYY-MM-DD.log */
  auditWrite: (lines: string[]): Promise<{ ok: boolean; file: string; err: string }> =>
    ipcRenderer.invoke('audit:write', lines),
  /** 当前日志文件绝对路径 */
  auditPath: (): Promise<{ dir: string; file: string }> => ipcRenderer.invoke('audit:path'),

  // ── 系统剪贴板（2026-11-07）───────────────────────────────────
  /** 把纯文本写入系统剪贴板（气泡右键「复制」；比 navigator.clipboard 可靠） */
  clipboardWrite: (text: string): Promise<boolean> => ipcRenderer.invoke('clipboard:write', text),
  /** 渲染层已自绘右键菜单 → 通知主进程本次不弹系统菜单 */
  ctxHandled: (): void => { ipcRenderer.send('ctx:handled') },
  /** 订阅「右键菜单 → 引用到输入框」：回调收到选中原文，返回退订函数 */
  onQuoteText: (cb: (text: string) => void): (() => void) => {
    const h = (_e: unknown, text: string): void => cb(String(text ?? ''))
    ipcRenderer.on('chat:quote-text', h)
    return () => ipcRenderer.removeListener('chat:quote-text', h)
  },

  // ── 应用生命周期 ───────────────────────────────────────────────
  /** 应用内重启：主进程 relaunch 新实例 + 立即退出当前（内嵌服务随进程重建） */
  relaunch: (): Promise<void> => ipcRenderer.invoke('app:relaunch'),
  /** 彻底退出：进程与内嵌服务全部终止（区别于点 X 仅隐藏窗口） */
  quit: (): Promise<void> => ipcRenderer.invoke('app:quit'),
})
