// 核心数据模型 —— 底座的地基
// 表与字段都是"元数据"，由代码定义，运行时可增可改。UI 只是它的投影。

export type FieldType = 'number' | 'text' | 'formula' | 'lookup'

/** 节点角色：决定它在流水线里的功能 */
export type NodeKind = 'table' | 'param' | 'sort' | 'output' | 'start' | 'map' | 'docx' | 'xlsx' | 'text' | 'test_table'
  /** ★ 节点组（2026-10-21 用户指令）：视觉容器——框住多个节点整体拖动/管理 */
  | 'group'
  /** ★ 便利贴/标注（2026-10-22 用户指令）：纯视觉标注，双击改文字 */
  | 'note'
  /** ★ PDF 合并（2026-10-22 用户指令）：上游 OO 文档按顺序转 PDF 并拼接 */
  | 'pdfmerge'
  /** ★ 照片/照片组（2026-10-23 用户指令「增加一个承载照片或者照片组的节点，后面测试合并PDF」）：
   *  承载多张照片（base64 随画布走），可连入 PDF 合并节点（每图一页 A4 适配） */
  | 'photo'
  /** ★ PDF 文件节点（2026-10-23 用户指令「增加一个承载PDF的节点」）：
   *  承载多份 PDF（base64 随画布走），连入 PDF 合并节点按行序直接拼接（零转换） */
  | 'pdf'
  /** ★ Excel转DXF（2026-11-10 用户指令「增加一个 excel转dxf 节点」）：
   *  表格坐标列 → 展点图（POINT/TEXT/LINE），生成 DXF R12 供 CAD 打开 */
  | 'xlsx2dxf'
  /** ★ 网址监控（2026-11-09 用户指令「增加一个网址监控节点…提取、记录、或者输出数据」）：
   *  轮询 URL 列表提取内容（CSS 选择器/JSON 路径/关键词），记录历史快照、标记变化 */
  | 'webmon'
  /** ★ 文件夹节点（2026-10-25 用户指令「作为读取位置和保存位置」）：
   *  指向本机一个目录——读模式=列出/读取目录内文件；写模式=作为保存输出目录 */
  | 'folder'
  /** ★ PDF 替换页节点（2026-10-25 用户指令「增加一个PDF替换页的节点」）：
   *  上面=被替换 PDF（含页码），下面=替换页来源 PDF（含页码），
   *  执行后产物入历史（replaceHistory）——图纸修订场景 */
  | 'pdfreplace'
  /** ★ PDF 瘦身节点（2026-10-26 用户指令「拖PDF进空白区域弹选择」）：
   *  承载多份 PDF（config 同 pdf 结构：pdfItems/pdfFileOrder），压缩瘦身在节点上操作 */
  | 'pdfslim'
  /** ★ PDF转Word 节点（2026-10-26 用户指令「利用百度的开源 增加个节点 pdf2word」）：
   *  百度 PaddleOCR PP-Structure 引擎（本机外挂），PDF → 可编辑 docx，产物入历史 */
  | 'pdf2word'
  /** ★ 文件输出节点（2026-10-26 用户指令「输出节点 可以存储文件到固定位置」）：
   *  连线上游产物（PDF/合并/替换页）→ 一键存储到固定目录（outDir 随画布走） */
  | 'fileout'
/**
 * 规则节点（圆形）：给连上的节点发规则（样式锁定 / 统一）。
 * 用户 2026-10-05：「画布内容中增加个圆形的节点。类似于功能节点。为规则适配节点。」
 * 自身不做数据处理、不进计算引擎 —— 纯控制节点。
 */
  | 'rule'
/**
 * AI 节点（圆形，2026-10-06）：模仿规则跟随节点样式。
 * 用户：「模仿规则跟随节点样式 增加一个AI节点。AI节点的连线放在普通节点的
 * 下1锚点位置。」→ 连线方向 = 普通节点下1 出线 → AI（AI 是数据消费者，
 * 上游存在自己 linkedInputs 里，与数据流同构落库）。
 */
  | 'ai'
/**
 * 参数库节点（圆形，2026-10-06）：模仿规则跟随节点样式。
 * 用户：「模仿规则跟随节点样式 增加一个参数库节点。参数库节点的连线放在
 * 普通节点的 上2锚点位置。」→ 连线方向 = 参数库出线 → 普通节点上2
 * （参数库是数据提供者，下游普通节点把它记进 linkedInputs）。
 */
  | 'paramlib'
/** OnlyOffice 节点：离线内嵌 OnlyOffice 内核（Word/Excel），替代原 docx/xlsx 的新一代形态 */
  | 'ooword' | 'oosheet'
/**
 * 章节工作簿（2026-10-05）：Excel 全量功能 + 外挂章节编排层
 *
 * 刻意单开kind 而非给xlsx 加配置项—— 差异化代码全部隔离在 outline* 与 ChapterDialog，
 * WorkbookEditor 只留一个可选outline 插槽。oo 系列升级时冲突面最小。
 */
  | 'chapter'

/** 开始节点触发器类型：手动 / 定时 / Webhook */
export type TriggerType = 'manual' | 'schedule' | 'webhook'

/** text 节点文本格式（备忘 · WPS 风格常用格式，画布节点渲染应用） */
export interface TextStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontSize?: number
  color?: string
  align?: 'left' | 'center' | 'right'
}

/** 导入格式（数据参数表从外部文件读入） */
export type ImportFormat = 'csv' | 'json' | 'xlsx'
/** 导出格式（数据输出节点存为文件） */
export type ExportFormat = 'csv' | 'json' | 'xlsx' | 'txt' | 'md' | 'pdf'

/** 节点配置：不同 kind 用不同字段 */
export interface NodeConfig {
  /** ★ 节点唯一编号（2026-10-12 用户指令）：N001 起，创建分配、复制重排——画布节点与存储文件的对照键 */
  nodeNo?: string
  /** text 节点：画布显示的文本格式（备忘 · WPS 风格常用格式） */
  textStyle?: TextStyle
  /** param 节点：导入来源格式 */
  importFormat?: ImportFormat
  /** sort 节点：按哪个字段排（field id） */
  sortBy?: string
  /** sort 节点：方向 */
  sortDir?: 'asc' | 'desc'
  /** output 节点：导出格式 */
  exportFormat?: ExportFormat
  /** sort/output 节点：上游来源节点 id 列表（全部引入的数据，含连线与面板手动勾选） */
  inputs?: string[]
  /** sort/output 节点：真正画了数据流连线的来源（连线 = 优先来源；面板勾选则无连线但仍汇入） */
  linkedInputs?: string[]
  /** start 节点：触发方式 */
  trigger?: TriggerType
  /** trigger=schedule 时的 cron / 简易表达式（如 "0 9 * * *"） */
  triggerCron?: string
  /** trigger=webhook 时的回调地址 */
  triggerWebhook?: string
  /** map 节点：用上游行的哪个字段的值去查对照表（字段 id，取上游字段） */
  mapKeyField?: string
  /** map 节点：对照表里用哪个字段做匹配键（字段 id，取本节点字段） */
  mapMatchField?: string
  /** map 节点：命中后取回对照表的哪些字段（字段 id 数组，可多列） */
  mapReturnFields?: string[]
  /** map 节点：未命中对照表时的默认填充值 */
  mapDefault?: string

  /* ── output 节点：Word 报告（docx 动态绑定） ────────────────────── */

  /** 绑定的模板 id（对应 userData/templates/<id>.docx） */
  docxTemplateId?: string
  /** 模板文件名，仅用于 UI 展示 */
  docxTemplateName?: string
  /**
   * 绑定映射：模板里的内容控件 tag → 取值来源
   * 例：{ 'field:管线总长': { from: 'field', fieldId: 'f_len' ,
   *      'each:管道清单':  { from: 'rows' },
   *      'rich:结论与建议': { from: 'manual', value: '<p>…</p>' } }
   */
  docxMap?: Record<string, DocxBinding>
  /** ★ 2026-10-30 版式档案（行距/分页/页面设置，首份原件固化）——重建与降级的版式兜底 */
  ooLineProfile?: import("./ooLayoutProfile").OoLayoutProfile
  /** 导出交付稿时是否剥掉内容控件外壳（默认 true：交付稿干净） */
  docxStripControls?: boolean

  /* ── xlsx 节点：Excel 报表（命名区域 / Table 动态绑定） ─────────── */

  /** 绑定的工作簿模板 id（对应 userData/templates/<id>.xlsx） */
  xlsxTemplateId?: string
  /** 模板文件名，仅用于 UI 展示 */
  xlsxTemplateName?: string
  /**
   * 绑定映射：工作簿里的「域」→ 取值来源。
   * Excel 的域用原生机制承载，与 Word 的 w:sdt 对等：
   *   命名区域（definedName） ←→ 标量域（相当于 Word 的内容控件）
   *   Table（ListObject）    ←→ 行级循环域（相当于 Word 的 each 行控件）
   * 例：{ '设计流量': { from: 'field', fieldId: 'f_q' },
   *      '材料表':   { from: 'rows', columns: { 名称: 'f_name', 数量: 'f_qty' } } }
   */
  xlsxMap?: Record<string, XlsxBinding>

  /* ── chapter 节点：章节工作簿（章节层级 + 目录大纲）─────────────── */

  /**
   * 章节标记列布局（列字母）。
   * 缺省走 A=level / B=no / C=title（方案 §6）。留这个配置是因为
   * 不同模板可能把标记列放在别的列 —— 改配置即可，不必改代码。
   */
  chapterCols?: { level?: string; no?: string; title?: string }
  /** 折叠状态按节点持久化（key 形如 `chapter-collapsed:<nodeId>`） */
  chapterCollapsed?: string[]

  /* ── OO 节点（ooword/oosheet）：OnlyOffice 离线编辑 ─────────────── */

  /** 保存回写后绑定的模板 id（对应 userData/templates/<id>.docx|xlsx）—— 重开节点自动恢复 */
  ooTemplateId?: string
  /** 模板文件名，仅用于 UI 展示 */
  ooTemplateName?: string
  /** ★ OO 内容 base64（2026-10-11）：随画布节点配置持久化——加载优先级高于 ooTemplateId（换机/清缓存不丢） */
  ooBytes64?: string
  /** ★ 2026-10-24：OO 内容 IndexedDB 索引（b64 出 config 的性能根治）——加载优先级同 ooBytes64 */
  ooBytesDbId?: string
  /** OO 文档数据锁定：true = 编辑器以只读模式打开（防误编辑，重开保持） */
  ooDataLocked?: boolean
  /** OO 节点位置锁定：true = 画布上该节点不可拖动（防误拖，重开保持） */
  ooPosLocked?: boolean
  /** ★ 统一数据锁（2026-10-26 全节点推广）：true = 节点内容防变动（OO 只读打开 /
   *  文件族禁导入/删除/排序/执行）——全 kind 生效；ooDataLocked 为 OO 时代遗留，读取兼容 */
  dataLocked?: boolean
  /** ★ 统一位置锁（2026-10-26 全节点推广）：true = 画布不可拖动——全 kind 生效；
   *  ooPosLocked 为 OO 时代遗留，读取兼容 */
  posLocked?: boolean
  /** OO 工作簿最后选中单元格（如 'D5'）——保存时随节点记录，打开时恢复选中并滚动可见
   *  （2026-10-22 用户指令「ooexcel 中也保存最后选中单元格的位置等信息」） */
  ooLastSel?: string
  /** PDF 合并节点（2026-10-23 行式列表重做）：默认保存文件名 / 上游文档行排序 */
  pdfName?: string
  pdfOrder?: string[]
  /** 合并历史（2026-10-23）：新在前 cap 10。来源区分——b64 存在 = 保存于画布
   *  （产物内嵌随画布走）；无 b64 = 保存到外部（磁盘），预览走模板库（IndexedDB）回退。 */
  /** ★ 文件夹节点（2026-10-25）：指向目录路径 + 读写模式 */
  folderPath?: string
  folderMode?: 'read' | 'write'
  /** ★ Excel转DXF（2026-11-10）：列映射与生成选项 + 最近一次产物元数据（DXF 文本存 pdfBlobStore） */
  xlsx2dxf?: {
    nameCol?: string
    xCol?: string
    yCol?: string
    zCol?: string
    drawPoint?: boolean
    drawLabel?: boolean
    drawLine?: boolean
    textHeight?: number
    labelOffset?: number
    flipY?: boolean
    fileName?: string
  }
  xlsx2dxfResult?: { at: number; count: number; skipped: number; fileName: string }
  /** ★ Excel转DXF 生成历史（2026-11-10 用户「生成一次 下方就增加一个 参照 PDF生成」）：
   *  新在前 cap 10；DXF 全文存 pdfBlobStore（'x2d-'+hid） */
  xlsx2dxfHistory?: { id: string; at: number; count: number; skipped: number; fileName: string }[]
  /** ★ 网址监控节点（2026-11-09）：监控目标列表 + 轮询间隔 + 历史快照（新在前 cap 100）。
   *  提取模式：text=CSS 选择器取文本（空=全正文）、json=点路径取字段、keyword=关键词出现检测 */
  webmonTargets?: { id: string; url: string; name?: string; mode: 'text' | 'json' | 'keyword'; selector?: string; jsonPath?: string; keyword?: string }[]
  webmonIntervalSec?: number
  webmonHistory?: { at: number; targetName: string; value: string; changed: boolean; error?: string }[]
  webmonLastValues?: Record<string, string>
  webmonLastRun?: number
  pdfHistory?: { id: string; name: string; at: number; b64?: string; color?: string }[]
  /** 照片节点（2026-10-23）：承载的照片组（id/name/类型/base64，随画布走）；
   *  photoOrder = 行序（id 序，拖拽排序持久）= 合并 PDF 的页序 */
  photoItems?: { id: string; name: string; type: string; b64?: string; photoDbId?: string; thumbDbId?: string }[]
  photoOrder?: string[]
  /** 圆形参数库节点（2026-10-23）：已关联的库名（双击节点从参数库面板选择） */
  libName?: string
  /** 主节点（OO 等）数据引用面板选中的参数库（2026-10-23）——优先展示连线 paramlib 的库 */
  paramLibSel?: string
  /** 双击内容窗体中选中的行（库内 idx + 全行 cells）——随节点持久 */
  libSelectedRow?: { idx: number; lib: string; cells: string[] }
  /** PDF 文件节点（2026-10-23）：承载的 PDF 文件组（name/base64，随画布走）。
   *  行序字段用 pdfFileOrder——pdfOrder 已被 PDF 合并节点（上游行排序）占用。 */
  pdfItems?: { id: string; name: string; b64?: string; inDb?: boolean }[]
  /** ★ PDF转Word 节点 config（2026-10-26）：词条/行序/历史/引擎语言——字节一律不进 config
   *  （词条=p2witem-{id}，产物=p2whist-{id}，见《文件保存规则.md》2f 节） */
  p2wItems?: { id: string; name: string; b64?: string; inDb?: boolean }[]
  p2wOrder?: string[]
  p2wHistory?: { id: string; name: string; at: number; inDb?: boolean; source?: string }[]
  /** 引擎识别语言：ch 中文（缺省）/ en 英文——PP-Structure 版面与 OCR 模型随语言切换 */
  p2wLang?: 'ch' | 'en'
  /** ★ 转换引擎（2026-10-26 用户指令「节点上的方形按钮 点击切换引擎」）：
   *  'x2t'=文字版直转（缺省，内置套件）/ 'ocr'=百度 PaddleOCR（扫描件，套件目录引擎）——
   *  节点按钮行方形钮点击切换，执行严格按选定引擎（不再自动探测路由） */
  p2wEngine?: 'x2t' | 'ocr'
  /** ★ OCR 速度档位（2026-10-26 用户问「有没有容差设置」）：
   *  'pro'=精准（缺省，server OCR+表格还原 9 模型 ~13s/页）
   *  'std'=均衡（server OCR+关表格还原，省 4 模型）
   *  'fast'=极速（mobile OCR+关表格还原，长文档首选）——仅 ocr 引擎生效 */
  p2wProfile?: 'fast' | 'std' | 'pro'
  /** ★ PDF 替换页节点 config（2026-10-25）：页码 1 基 */
  replaceTargets?: { id: string; name: string; b64?: string; inDb?: boolean; page: number }[]
  replaceSources?: { id: string; name: string; b64?: string; inDb?: boolean; page: number }[]
  replaceHistory?: { id: string; name: string; at: number; b64?: string; inDb?: boolean; pages?: number[] }[]
  /** 匹配容差（0-100）：相似度低于该值的配对不进确认列表（缺省 60，UI 兜底） */
  replaceTolerance?: number
  /** 文件输出节点（fileout）：固定输出目录 + 最近保存记录（轻量索引，≤10 条） */
  outDir?: string
  lastSaved?: { name: string; at: number; size: number }[]
  pdfFileOrder?: string[]
  /** PDF 文件节点「从上游生成」历史（2026-10-29）：与 PDF 合并节点 pdfHistory 同口径——
   *  旧 pdfItems 逐份归档（字节入库 pdfgenhist-<id>，config 只留索引），新在前 cap 10 */
  pdfGenHistory?: { id: string; name: string; at: number; inDb?: boolean }[]
  /** OO 文档登记的数据引用（域清单）。
   *  ★ urn 是**权威地址**（node://<nodeId>#<fieldId>），用 ref/urn.ts 的协议存，
   *    写回管线直接喂给 refGraph / trigger，不再自己拼字符串。
   *    nodeId/fieldId 保留是为了向后兼容老数据与快速比较；
   *    缺失 urn 的老数据由 ooRefUrn() 就地补齐（见 core/ooRef.ts）。 */
  ooRefs?: { nodeId: string; fieldId: string; urn?: string }[]
  /** ★ 外部数据引用更新策略（2026-10-18 用户指令「设置好 保存最后的状态」）：
   *  'off' 不检测 / 'ask' 检测后弹窗提示（默认）/ 'auto' 自动更新不提示。
   *  随节点配置走画布文件（换机/清缓存不丢）；localStorage 只作旧数据回退。 */
  ooRefUpdMode?: 'off' | 'ask' | 'auto'
  /** ★ 打开历史（2026-10-19 用户指令「打开文件左侧加个打开历史记录的按钮」）：
   *  最近导入的文件（字节随存，点历史条目直接重开）。上限 3 条、单条 >4MB 不入史
   *  （防配置膨胀）；同名去重（重开同名文件移到最前）。 */
  ooHistory?: { name: string; ts: number; b64: string }[]
  /** ★ 节点组（2026-10-21）：kind='group' 专用——成员节点 id 清单（随组拖动）、
   *  边框尺寸（右下角拖拽调整）。存在组节点自己的 config 上。 */
  groupW?: number
  groupH?: number
  /** 便利贴尺寸（NodeResizer 落库） */
  noteW?: number
  noteH?: number
  members?: string[]
  /** ★ 节点组折叠：true = 组内成员收成折叠方块（成员渲染侧扫描此标记） */
  groupFolded?: boolean
  /** ★ 节点组最大化：放大前的位置/尺寸（还原用） */
  groupPrev?: { x: number; y: number; w: number; h: number }
  /** ★ 锁定组（2026-10-19 用户指令）：true = 本节点与连着它的圆形节点（AI/参数库/
   *  规则跟随）相对位置锁定——拖任何一个，整组一起移动。存在主节点（常规节点）上。 */
  grpLocked?: boolean

  /* ── ai 节点：Agent 配置（2026-10-06 用户：「模仿常规agent和提示词。
     补充ai节点的，api设置，提示词，skill，输入，输出」）──────────────
     字段全部可选：undefined = 用 PropertiesPanel 占位提示/默认值，
     不给配置也能保存工程（与 start/sort 等节点的既有惯例一致）。
     ★ 真正调 LLM 的运行时不在本层 —— 这里只是「agent 的设置表」；
       执行管线（拉上游数据 → 插值 {{变量}} → 调 API → 产出）是后续任务。 */

  /** API 地址（OpenAI 兼容 base URL，如 https://api.deepseek.com/v1） */
  aiApiBase?: string
  /**
   * API Key。★ 随工程保存（.wfc 单文件带走）——单机 C 端工具的取舍：
   * 密钥跟工程走换机即用；代价是工程文件含敏感信息，分享时须自行清除。
   * PropertiesPanel 的提示里写明。
   */
  aiApiKey?: string
  /** 模型名（如 deepseek-chat / gpt-4o-mini / qwen-plus） */
  aiModel?: string
  /** 温度 0~2（缺省 0.7） */
  aiTemperature?: number
  /** 最大输出 token（缺省 2048） */
  aiMaxTokens?: number
  /** 系统提示词（角色/规则；选技能时会预填骨架，可改） */
  aiSystemPrompt?: string
  /** 用户提示词模板，{{变量名}} 引用输入变量 */
  aiUserPrompt?: string
  /** 技能 id（AI_SKILLS 注册表；'custom' = 自由填写不绑定骨架） */
  aiSkillId?: string
  /**
   * 输入变量映射：变量名 ← 上游节点（linkedInputs 里连了线的上游）。
   * ★ 以 linkedInputs 为准：断线后对应变量行在面板自动消失（映射留着无害，
   *   渲染时只展示仍连着的）；改名只动本表，不碰连线。
   */
  aiInputVars?: { name: string; from: string }[]
  /** 输出模式：纯文本 / JSON（下游可按字段解析）/ Markdown */
  aiOutputMode?: 'text' | 'json' | 'markdown'
  /** 输出说明：json 模式写字段清单（给模型的 schema hint），text/markdown 写格式要求 */
  aiOutputNote?: string
  /**
   * 最近一次运行结果——随工程保存，重开可见。
   * 运行是「执行」不是「编辑」：写它走 store 直写（不进撤销栈），
   * 否则每跑一次吃一条历史快照，撤销语义就被执行污染了。
   */
  aiLastOutput?: string
  /** 最近一次运行的失败原因（如实显示在面板运行区） */
  aiLastError?: string

  /* ── 规则锁定（2026-10-05：圆形节点为规则适配节点）───────────────
     用户原话：
     「圆形节点是对连线的节点进行规则锁定（统一）」
     「一旦匹配了，默认不能每个改规则，只能改主流程的节点进行修改。
       修改后所有相同规则的节点规则同步改。」

     ★ 铁律：规则**只存一份**，就在圆形节点身上。
       辅节点一律不存规则副本，只存 ruleRef 引用。
       若辅节点各自存一份，改主节点时必须逐个同步写，
       漏一个就产生「规则副本漂移」—— 且这种漂移极难发现
       （画面看着正常，只是某处没跟着变）。
       引用式存储让「同步」这件事**在结构上不可能漏**。
     ────────────────────────────────────────────────────────────────── */

  /**
   * rule 节点（圆形）持有的规则集。
   * 按被锁节点的 kind 分桶 —— 不同节点类型适配不同的规则项
   * （用户：「针对不同节点的内容 适配不同的规则样式体系」）。
   *
   * 例：{ sort: {...排序规则}, map: {...匹配规则}, text: {...文本样式} }
   */
  rules?: Record<string, NodeRuleSet>

  /**
   * 辅节点身上的**唯一**规则信息：被哪个规则节点锁定了、主从与跟随状态。
   * 不含任何规则内容 —— 读规则一律跳到 ruleRef 指向的 rule 节点取。
   *
   * ★ 为什么主从/跟随状态存在**辅节点身上**而不是规则节点的列表里：
   *   主从是「这个节点在这个组里的身份」，天然属于节点自己。
   *   若把成员表放规则节点上，就得额外维护一份「谁在组内」，
   *   而那份表与连线（linkedInputs）一有出入就出现两份真相——
   *   与「不另存边表」同一条理由：连线就是成员关系，不另存。
   */
  ruleRef?: {
    /** 规则节点 id */
    ruleNodeId: string
    /**
     * 组内身份（用户：「规则节点第一个默认是可修改的主节点。
     * 连接剩下的是从节点」）。
     *
     * · primary  主节点 —— 组内唯一，排在第一个。改规则时以它为锚。
     *· follow   从节点 —— 跟随主节点的规则。
     * · ignore   不跟随 —— 保留连线但不受该规则影响（用户：
     *   「双击能修改节点的主次、跟随不跟随」里的「不跟随」）。
     *
     * ⚠ 向后兼容：早期版本没有这个字段。缺失时**按 follow 处理**
     *   （即照旧跟随规则）—— 不能缺失就当 ignore，否则升级后
     *   已有连线会静默失去规则，表现为「规则一夜之间失效了」。
     */
    mode?: 'primary' | 'follow' | 'ignore'
  }
}

/**
 * AI 节点内置技能注册表（模仿常规 agent 平台的 skill 插槽；单一真相源：
 * PropertiesPanel 的下拉、提示词骨架、推荐输出模式都从这取，不另写一份）。
 * ★ 扩展方式：往数组加一项即可在面板出现对应技能，零组件改动。
 */
export const AI_SKILLS: {
  id: string
  name: string
  desc: string
  systemPrompt: string
  outputMode: 'text' | 'json' | 'markdown'
}[] = [
  {
    id: 'text-io',
    name: '文本处理',
    desc: '输入文字 → 输出文字（改写/润色/翻译等，按用户要求）',
    outputMode: 'text',
    systemPrompt:
      '你是文本处理助手。对用户提供的文字进行加工（改写、润色、翻译、格式化等，以用户消息里的要求为准），' +
      '直接输出处理后的文字本身——不要解释、不要前缀、不要包引号。',
  },
  {
    id: 'summarize',
    name: '内容总结',
    desc: '把上游数据归纳成要点摘要',
    outputMode: 'text',
    systemPrompt:
      '你是一名严谨的数据分析助手。请基于用户提供的资料进行总结：\n' +
      '1. 先用一句话概括整体结论；\n2. 分条列出关键要点（每条不超过 40 字）；\n' +
      '3. 如数据存在缺失或矛盾，明确指出，不要编造。',
  },
  {
    id: 'extract',
    name: '字段提取',
    desc: '从上游数据抽取结构化字段（JSON）',
    outputMode: 'json',
    systemPrompt:
      '你是一个信息抽取引擎。从用户提供的资料中抽取指定字段，只输出一个 JSON 对象，' +
      '不要输出任何解释文字。字段清单与类型见用户的输出要求；抽取不到的字段填 null。',
  },
  {
    id: 'translate',
    name: '中英互译',
    desc: '上游内容中英互译，保留术语',
    outputMode: 'text',
    systemPrompt:
      '你是一名专业翻译。将用户提供的内容在中英文之间互译（自动判断方向）：\n' +
      '1. 技术/工程术语按行业标准译法，首次出现附原文；\n2. 数字、单位、编号原样保留；\n3. 不增删信息。',
  },
  {
    id: 'cost-qa',
    name: '造价问答',
    desc: '工程造价领域问答（贴行业口径）',
    outputMode: 'markdown',
    systemPrompt:
      '你是工程造价领域的资深工程师，熟悉清单计价、定额、燃气/市政等专业口径。' +
      '基于用户提供的资料回答问题：\n1. 结论先行；\n2. 引用资料中的具体数字时注明出处字段；\n' +
      '3. 资料不足以支撑结论时明确说明，并给出需要补充什么数据。',
  },
  {
    id: 'custom',
    name: '自定义',
    desc: '自由填写提示词，不绑定骨架',
    outputMode: 'text',
    systemPrompt: '',
  },
]

/**
 * 按上游节点类型匹配 AI 偏好（2026-10-06 用户：「针对不同类型的节点。
 * ai接入后匹配不同的偏好。先做一个输入文字输出文字的样例」）。
 *
 * 结构刻意薄：只存 skillId + 可选的 userPrompt/outputNote 覆盖——
 * 系统提示词骨架与推荐输出模式一律从 AI_SKILLS 按 skillId 取
 * （那已是单一真相源，这里再存一份必然漂移）。
 * ★ 扩展：给新类型节点接 AI 时往这里加一项即可，handleConnect 零改动。
 */
export const AI_KIND_PREFS: Partial<
  Record<NodeKind, { skillId: string; userPrompt?: string; outputNote?: string }>
> = {
  /* text = 第一个样例：文本节点接入 → 自动配「文字进文字出」。
   * userPrompt 留 undefined = 运行时上游原文直接当用户消息 —— 零配置闭环：
   * 连完线点运行就能跑（执行引擎见 core/aiRun.ts）。 */
  text: { skillId: 'text-io' },
}

/**
 * AI 节点接入上游时算「偏好 patch」：未配置过技能才匹配（aiSkillId 非空 =
 * 用户已表达意图，含选了 custom —— 一律不覆盖）；多上游取第一个命中类型的偏好。
 * 返回 null = 不需要预填，调用方只写 linkedInputs。
 */
export function aiPrefsPatchOf(
  config: NodeConfig,
  upstreamKinds: string[],
): Partial<NodeConfig> | null {
  if (config.aiSkillId) return null
  for (const k of upstreamKinds) {
    const pref = AI_KIND_PREFS[k as NodeKind]
    if (!pref) continue
    const skill = AI_SKILLS.find((s) => s.id === pref.skillId)
    if (!skill) continue
    const patch: Partial<NodeConfig> = {
      aiSkillId: skill.id,
      aiSystemPrompt: skill.systemPrompt,
      aiOutputMode: skill.outputMode,
    }
    if (pref.userPrompt !== undefined) patch.aiUserPrompt = pref.userPrompt
    if (pref.outputNote !== undefined) patch.aiOutputNote = pref.outputNote
    return patch
  }
  return null
}

/**
 * 一份规则 —— 不同 kind 适配不同的字段，故用可选字段而非联合类型。
 *
 * 用户：「针对不同节点的内容 适配不同的规则样式体系」。
 * 用可选字段而非联合类型，是为了让「按 kind 校验哪些字段必填」
 * 集中在一处（RULE_FIELDS_OF_KIND），而不是散在各个 if 分支里。
 */
export interface NodeRuleSet {
  /** 视觉：族色（见 nodeStyle.ts 的 family） */
  family?: string
  /** 视觉：形状（见 nodeStyle.ts 的 shape） */
  shape?: 'flow-in' | 'flow' | 'flow-out' | 'trigger' | 'gate' | 'store' | 'page'
  /** 视觉：尺寸倍率，1 = 默认。圆形节点用不着，但方节点可以统一缩放 */
  scale?: number

  /* sort 节点 */
  sortBy?: string
  sortDir?: 'asc' | 'desc'

  /* map 节点 */
  mapKeyField?: string
  mapMatchField?: string
  mapReturnFields?: string[]

  /* text 节点 */
  textStyle?: TextStyle

  /* output / docx / xlsx 节点 */
  exportFormat?: ExportFormat
}

/**
 * 「一个域取什么值」—— Word / Excel 共用的基类型。
 *
 * 为什么要抽出来：DocxBinding 与 XlsxBinding 原先各自写了一遍同样的四个字段，
 * 两边改一处忘一处是迟早的事。现在只有一处定义，两种产物共享同一套语义，
 * 同一张参数表可以同时驱动 Word 和 Excel。
 *
 * 取值的两种表达方式（二选一，`source` 优先）：
 *   · `fieldId` —— 单源快捷方式：引用本节点的某个上游字段
 *   · `source`  —— 统一引用地址（URN）：node:// file:// api:// param:// const:
 *                  多源引用只能用它。见 src/ref/urn.ts
 */
export interface FieldBinding {
  from: 'field' | 'rows' | 'manual'
  /** 取哪个上游字段。与 source 二选一；给了 source 就以 source 为准。 */
  fieldId?: string
  /**
   * 统一引用地址（URN）。跨节点 / 跨文件 / 网络 API / 全局参数 / 常量都走它。
   * 例：'node://calc1#设计流量'、'file://D:/价目.xlsx#价目表!B2'、'param://审查单位'
   */
  source?: string
  /** from=field 且上游多行时的聚合方式；缺省取首行 */
  agg?: 'first' | 'sum' | 'count' | 'max' | 'min' | 'avg'
  /** from=manual 时的值；Word 的 rich 类型存 HTML，其它存纯文本 */
  value?: string
  /**
   * from=rows（行级循环域）时的列映射：模板里那一行的占位符名 → 上游字段 id。
   * 不配则按名字自动匹配；上游字段名与模板列名不一致时必须在这里指定。
   */
  columns?: Record<string, string>
}

/**
 * Word 文档里一个「域」（w:sdt 内容控件）的取值来源。
 * 字段全在 FieldBinding 里，这里只是取个语义化名字。
 */
export type DocxBinding = FieldBinding

/**
 * Excel 里一个「域」的取值来源。
 *
 * 与 Word 侧的唯一差别：Table 可以带「计算列」公式 ——
 * Excel 有公式引擎，Word 没有。行级模板 `{列名}` 会在写盘时换成该行的真实引用，
 * 并做 formula + cachedValue 双写（只写公式不写缓存，Excel 可能显示 #N/A）。
 */
export interface XlsxBinding extends FieldBinding {
  /**
   * 计算列，仅 from=rows 有效。形状与 bindXlsxFromParamTable 的 tableFormulas 一致：
   *   列名 → 行级公式模板（`{列名}` 引用同一行的其它列）
   * 例：{ 合价: '={数量}*{单价}' } → 写盘后成为 '=B5*C5'
   *
   * 之所以不写成一个 expr + 猜目标列：猜不出来。列名必须由用户明确指定。
   */
  formulas?: Record<string, string>
}

export interface LookupSpec {
  /** 关联表 id */
  table: string
  /** 用本表哪个字段去匹配 */
  sourceField: string
  /** 关联表里用哪个字段做匹配键 */
  matchField: string
  /** 命中后取关联表的哪个字段回来 */
  returnField: string
}

export interface FieldDef {
  id: string
  name: string
  type: FieldType
  /** formula 字段：表达式，可用 [@字段名] 引用同行其他字段，如 '=[@qty]*[@price]' */
  formula?: string
  /** lookup 字段：跨表取值的配置 */
  lookup?: LookupSpec
  /** 显示格式 */
  format?: 'currency' | 'number' | 'text'
}

export interface TableSchema {
  id: string
  name: string
  kind: NodeKind
  fields: FieldDef[]
  config?: NodeConfig
  /** 软删除标记：被移入回收站的节点，画布不渲染，可恢复；默认未删除 */
  deleted?: boolean
  /** 回收站元数据：删除时记录其它节点指向本表的关联引用，用于恢复后还原连线 */
  trashMeta?: {
    deletedAt: number
    refs: Record<string, { inputs: string[]; linkedInputs: string[]; lookups: FieldDef[] }>
  }
}

/** 定义一个字段。id 默认取 name，要求表内唯一。 */
export function defineField(
  name: string,
  type: FieldType,
  opts: Partial<Omit<FieldDef, 'id' | 'name' | 'type'>> & { id?: string } = {},
): FieldDef {
  return {
    id: opts.id ?? name,
    name,
    type,
    formula: opts.formula,
    lookup: opts.lookup,
    format: opts.format,
  }
}

/** 定义一张表/节点。id 默认取 name，要求全局唯一。 */
export function defineTable(
  name: string,
  fields: FieldDef[],
  id?: string,
  kind: NodeKind = 'table',
  config?: NodeConfig,
): TableSchema {
  return { id: id ?? name, name, kind, fields, config }
}

/** 便捷的节点构造器（默认带唯一 id），kind 必填。 */
export function defineNode(
  kind: NodeKind,
  name: string,
  fields: FieldDef[] = [],
  config: NodeConfig = {},
  id?: string,
): TableSchema {
  return defineTable(name, fields, id ?? `${kind}_${Date.now()}`, kind, config)
}
