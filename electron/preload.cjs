/**
 * Electron preload —— CommonJS 手写版。
 *
 * 为什么不走 vite 构建：vite-plugin-electron 0.28.8 按 package.json "type":"module"
 * 硬编码输出 ESM（配置里的 lib.formats:['cjs'] 覆盖被忽略），而 Electron sandbox
 * 以 CJS 加载 preload → "Cannot use import statement outside a module" → window.api
 * 整个消失。.cjs 后缀在任何 package type 下都是 CJS，绕开整类问题。
 *
 * 与 electron/preload.ts 保持 API 同步（preload.ts 保留作为类型来源，不再参与构建）。
 */
const { contextBridge, ipcRenderer } = require('electron')

/* ★ UI 偏好快照（2026-10-20）：主进程 settings:get-sync 同步取一次，
   随 contextBridge 注入 window.api.prefs —— 渲染层可同步读（端口随机导致
   localStorage 每次启动换源，界面主题/背景等改存文件持久化）。 */
let __prefs = {}
try { __prefs = ipcRenderer.sendSync('settings:get-sync') || {} } catch (e) { __prefs = {} }

contextBridge.exposeInMainWorld('api', {
  /** ★ UI 偏好快照（主题/背景等；同步注入，启动即可读） */
  prefs: __prefs,
  /** UI 偏好增量保存（合并进 userData/ui-settings.json） */
  settingsSet: (patch) => ipcRenderer.invoke('settings:set', patch),
  readDb: () => ipcRenderer.invoke('db:read'),
  writeDb: (data) => ipcRenderer.invoke('db:write', data),

  // ── 本地聪明模型（GGUF + node-llama-cpp，主进程懒加载；opts.model 可指定模型文件）──与 preload.ts 同步 ──
  // ★ 网址监控（2026-11-09）──与 preload.ts 同步 ──
  webmonFetch: (req) => ipcRenderer.invoke('webmon:fetch', req),
  llmChat: (prompt, opts) => ipcRenderer.invoke('llm:chat', prompt, opts),
  // ★ 流式版：onChunk 边生成边回调（kind: 'text'=正文 | 'thought'=思考段；llm:chunk 事件按 streamId 过滤）──与 preload.ts 同步 ──
  llmChatStream: (prompt, opts, onChunk) => {
    const streamId = 's' + Math.random().toString(36).slice(2)
    const listener = (_ev, p) => { if (p.streamId === streamId) onChunk({ kind: p.kind, text: p.text }) }
    ipcRenderer.on('llm:chunk', listener)
    return ipcRenderer
      .invoke('llm:chat-stream', prompt, { ...opts, streamId })
      .finally(() => ipcRenderer.removeListener('llm:chunk', listener))
  },

  // ── 联网搜索（Bing 抓取）与打开引用链接 ──
  webSearch: (query) => ipcRenderer.invoke('web:search', query),
  openUrl: (url) => ipcRenderer.invoke('shell:openUrl', url),

  // ── Word 模板库（userData/templates） ───────────────────────────
  tplList: () => ipcRenderer.invoke('tpl:list'),
  tplSave: (name, data) => ipcRenderer.invoke('tpl:save', name, data),
  tplRead: (id) => ipcRenderer.invoke('tpl:read', id),
  tplDelete: (id) => ipcRenderer.invoke('tpl:delete', id),
  /** 打开系统文件对话框挑模板（kind: 'docx' | 'xlsx'，默认 docx），取消返回 null */
  tplPick: (kind) => ipcRenderer.invoke('tpl:pick', kind || 'docx'),
  /** 产物另存为（kind 决定对话框标题与扩展名）；取消返回 null */
  docSaveAs: (defaultName, data, kind) => ipcRenderer.invoke('doc:saveAs', defaultName, data, kind || 'docx'),
  openExternalFile: (name, data) => ipcRenderer.invoke('shell:openExternalFile', name, data),
  /** 用本机浏览器打开（Edge/Chrome 自带 PDF 查看器；2026-10-26） */
  openInBrowser: (name, data) => ipcRenderer.invoke('shell:openInBrowser', name, data),

  /* ★★ 2026-10-26 桥同步大补（此前 cjs 落后 preload.ts 一大批——打包版桌面桥缺失根因）。
     以后 preload.ts 加通道必须同步这里（cjs 是打包版唯一生效的桥）！ */
  gpuGetAccel: () => ipcRenderer.invoke('gpu:getAccel'),
  gpuSetAccel: (on) => ipcRenderer.invoke('gpu:setAccel', on),
  showRemoteDialog: () => ipcRenderer.invoke('remote:showDialog'),
  pdfDiffView: (a, b) => ipcRenderer.invoke('pdf:diffView', a, b),
  folderPickDir: () => ipcRenderer.invoke('folder:pickDir'),
  folderList: (dir) => ipcRenderer.invoke('folder:list', dir),
  folderReadFile: (dir, name) => ipcRenderer.invoke('folder:readFile', dir, name),
  projectNativePack: (files, dirs, savePath) => ipcRenderer.invoke('project:nativePack', files, dirs, savePath),
  projectNativeUnpack: (zipBytes) => ipcRenderer.invoke('project:nativeUnpack', zipBytes),
  wfcPickSavePath: (defaultName) => ipcRenderer.invoke('wfc:pickSavePath', defaultName),
  wfcWriteFile: (filePath, bytes) => ipcRenderer.invoke('wfc:writeFile', filePath, bytes),
  assetsExport: (items) => ipcRenderer.invoke('assets:export', items),

  // ── PDF转Word：本机 PaddleOCR 引擎通道（2026-10-26）──────────────
  p2wTmpdir: () => ipcRenderer.invoke('p2w:tmpdir'),
  p2wExec: (payload) => ipcRenderer.invoke('p2w:exec', payload),

  // ── 系统剪贴板（2026-11-07）：把纯文本写进系统剪贴板 ─────────────
  /** 气泡右键「复制」用：主进程 clipboard.writeText，比 navigator.clipboard 可靠
      （后者在自定义协议 / 无焦点时会静默失败，导致 Ctrl+V 粘出上一次的旧截图） */
  clipboardWrite: (text) => ipcRenderer.invoke('clipboard:write', text),

  /** ★ 2026-11-07 气泡右键：渲染层自己弹了菜单，通知主进程这次别弹系统菜单（避免两层叠加） */
  ctxHandled: () => ipcRenderer.send('ctx:handled'),

  /** ★ 2026-11-07 用户「右键能引用作为提示词工程的一部分」：
      右键菜单（主进程）点「引用到输入框」→ chat:quote-text → 这里回调渲染层，
      把选中的原文插进对话输入框。返回退订函数。 */
  onQuoteText: (cb) => {
    const h = (_e, text) => { try { cb(String(text ?? '')) } catch { /* ignore */ } }
    ipcRenderer.on('chat:quote-text', h)
    return () => ipcRenderer.removeListener('chat:quote-text', h)
  },

  // ── 套件中心（2026-10-26）：状态读写 / 打开存储目录 / OCR 动态装卸 ──
  suiteRead: () => ipcRenderer.invoke('suite:read'),
  suiteWrite: (patch) => ipcRenderer.invoke('suite:write', patch),
  suiteOpenDir: () => ipcRenderer.invoke('suite:openDir'),
  suiteOcrInstall: () => ipcRenderer.invoke('suite:ocrInstall'),
  suiteOcrUninstall: () => ipcRenderer.invoke('suite:ocrUninstall'),

  /** 完整 HTML → PDF（隐藏窗口打印），返回保存路径；取消返回 null */
  docPdf: (html, defaultName) => ipcRenderer.invoke('doc:pdf', html, defaultName),
  /** 记录一次「打不开的文件」样本（元数据 + 原始字节），返回落盘目录 */
  failSave: (payload, bytes) => ipcRenderer.invoke('fail:save', payload, bytes),
  /** 列出已记录的失败样本（诊断用） */
  failList: () => ipcRenderer.invoke('fail:list'),

  // ── 操作审计日志（2026-10-17 用户指令：点击/键盘/反馈全记录）───────
  /** 批量追加已格式化的日志行 → logs/audit-YYYY-MM-DD.log；返回 { ok, file, err } */
  auditWrite: (lines) => ipcRenderer.invoke('audit:write', lines),
  /** 当前日志文件绝对路径（界面/排查时告诉用户日志在哪） */
  auditPath: () => ipcRenderer.invoke('audit:path'),

  // ── 应用生命周期 ───────────────────────────────────────────────
  /** 应用内重启：主进程 relaunch 新实例 + 立即退出当前（内嵌服务随进程重建） */
  relaunch: () => ipcRenderer.invoke('app:relaunch'),
  /** 彻底退出：进程与内嵌服务全部终止（区别于点 X 仅隐藏窗口） */
  quit: () => ipcRenderer.invoke('app:quit'),

  // ── .wfc 文件关联（2026-10-07：双击 .wfc → 打开软件并加载）────
  /** 取启动参数里的 .wfc 路径（首启带文件双击场景；取走即清，避免下次误触发） */
  getLaunchWfc: () => ipcRenderer.invoke('wfc:get-launch'),
  /** 订阅「运行中又双击了另一个 .wfc」（second-instance 转发）；返回退订函数 */
  onOpenWfc: (cb) => {
    const h = (_e, path) => cb(path)
    ipcRenderer.on('wfc:open', h)
    return () => ipcRenderer.removeListener('wfc:open', h)
  },
  /** 📂 系统对话框选 .wfc（C 端原生文件对话框，浏览器端走 input[type=file]） */
  pickWfc: () => ipcRenderer.invoke('wfc:pick'),
  /** 读 .wfc 字节（文件在主进程侧读——渲染层无磁盘权限） */
  readWfc: (path) => ipcRenderer.invoke('wfc:read', path),
})
