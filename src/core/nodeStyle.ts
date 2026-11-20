// ════════════════════════════════════════════════════════════════════════
// 节点样式规则表 —— 单一真相源
// ════════════════════════════════════════════════════════════════════════
// 目的：新增一种节点类型时，只改本文件一处，UI 全线自动一致。
//
// 为什么要规则化（2026-10-02）：
//   改前散点式——11 个 kind 里只有 5 个有 --nk-* token（text/xlsx/docx/param/table），
//   sort/map/output/start/ooword/oosheet 的颜色是各处 CSS 写死的十六进制，
//   形状（palette-*-shape）另有 6 套手写规则。加一种新 kind 要动 4~6 个文件，
//   且极易漏（典型症状：ooword 徽章没色，回退成默认灰）。
//
// 铁律三条：
//   1. 本文件是唯一真相源。颜色 / 形状 / 徽章文案 / 连线规则，都不在别处写死。
//   2. CSS 侧一律 `var(--nk-<kind>-*)` 引用，不许出现裸 hex。
//   3. 加 kind = 在 KIND_STYLE 加一行 + 在 CSS 的 --nk-* 区块加 5 个变量；
//      形状靠 CSS 类组合自动生效，不另写新规则。
//
// 维度设计（关键）：
//   role —— 七类之一，单值，决定节点在流水线里的**位置与连线约束**：
//     input 输入型 / process 数据处理型 / output 输出型 / action 动作型
//     loop 循环型 / library 资料库型 / template 模板型
//   form —— 形态标签，多值，决定**编辑器面板渲染哪套 UI**：
//     table 表格 / text 文本 / file 文件 / doc 文档
//   同一种 kind 的 role 是唯一确定的，但 form 可以组合（如 oosheet = output +
//   {table, file, doc}），这样"输出型"和"模板型"不再互斥——分类信息不丢失。
// ════════════════════════════════════════════════════════════════════════

/** 七类角色 */
export type NodeRole = 'input' | 'process' | 'output' | 'action' | 'loop' | 'library' | 'template'

/** 形态标签 */
export type NodeForm = 'table' | 'text' | 'file' | 'doc'

/** 一个 kind 的完整样式与行为定义 */
export interface KindStyle {
  /** 徽章文案（画布节点头 / 节点库面板共用） */
  label: string
  /** 七类角色，决定连线约束与执行语义 */
  role: NodeRole
  /** 形态标签，决定编辑器面板渲染哪套 UI */
  forms: NodeForm[]
  /**
   * 色相族名，对应 CSS 的 --nk-<family>-* 变量组。
   * 同族可复用（同族不同明度），不同族必须自成一族（否则无法一眼区分）。
   * 留空 = 不参与配色（终端节点无徽章）。
   */
  family?: string
  /**
   * 调色板形状后缀，对应 CSS 的 .palette-<shape>-shape。
   * 规则：同一 role 用同一形状（形状表达"角色"，颜色表达"族"），
   * 这��让调色板成为可扫读的图例 —— 形状看角色，颜色看具体 kind。
   */
  shape: 'flow-in' | 'flow' | 'flow-out' | 'trigger' | 'gate' | 'store' | 'page' | 'group'
  /** 是否可作为流水线终点（output 终止，input 起始） */
  terminal?: boolean
  /** 能否有下游（output不可有下游——它是终点） */
  allowsDownstream?: boolean
  /**
   * 血缘声明（2026-10-05，用户要求「在节点系统上增加衍生的关联标记」）：
   * 本kind 是从哪个 kind 派生的、做了哪些适配。
   *
   * ★ 为什么放这里（单一真相源铁律）：
   *   「衍生关系」与「角色 / 配色 / 连线规则」是同一类判定，
   *   已经在 engine / App / Canvas / PropertiesPanel 各写一遍就漏过。
   *   故与 KIND_STYLE 同处，UI 只读不写。
   *
   * 不声明 = 原始节点（无衍生关系）。
   */
  derivesFrom?: {
    /** 父kind 名，必须存在于 KIND_STYLE */
    kind: string
    /** 这个 kind 叫什么（衍生代际名，如 oosheet → ooec） */
    generation: string
    /**
     * 做了哪些适配 —— 升级父内核时逐条核对这张表（docs/24 §2.1 的语境）。
     * 每条须写清「改哪里」，否则这张表没有执行力。
     */
    adaptations: { area: string; note: string }[]
  }
}

/**
 * kind → 样式规则。全表。
 *
 * 族名沿用既有 v1 色谱（text 紫 / xlsx 浅绿 / docx 浅蓝 / param 深绿 / table 蓝），
 * 新增 kind 复用或新开一族，规则一致。
 */
export const KIND_STYLE: Record<string, KindStyle> = {
  // ── 输入型 ──────────────────────────────────────────────
  param: {
    label: '参数表',
    role: 'input',
    forms: ['table'],
    family: 'param',
    shape: 'flow-in',
    allowsDownstream: true,
  },
  /* ★ 2026-10-18 用户指令：新增「数据表节点（测试）」——手动输入参数+参数数据，
   *   专供 OO Excel 引用调试（引用面板点选 / 引用更新策略 / 原生标记全链验证）。
   *   行为与参数表完全同款（table 表单：可编辑字段 + 数据行），配色复用 param 族。 */
  test_table: {
    label: '数据表节点',
    role: 'input',
    forms: ['table'],
    family: 'param',
    shape: 'flow-in',
    allowsDownstream: true,
  },
  start: {
    label: '开始',
    role: 'action',
    forms: [],
    family: 'start',
    shape: 'trigger',
    terminal: true,
    allowsDownstream: true,
  },

  // ── 数据处理型 ──────────────────────────────────────────
  sort: {
    label: '数据排序',
    role: 'process',
    forms: ['table'],
    family: 'sort',
    shape: 'flow',
    allowsDownstream: true,
  },
  map: {
    label: '数据对照表',
    role: 'process',
    forms: ['table'],
    family: 'map',
    shape: 'flow',
    allowsDownstream: true,
  },
  text: {
    label: '文本',
    role: 'process',
    forms: ['text'],
    family: 'text',
    shape: 'flow',
    allowsDownstream: true,
  },

  /**
   * 便利贴/标注（2026-10-22 用户指令）：纯视觉标注节点。
   * 亮黄便利贴外观，双击改文字；无连线语义（allowsDownstream=false）。
   */
  note: {
    label: '便利贴',
    role: 'output',
    forms: ['text'],
    family: 'note',
    shape: 'flow',
    allowsDownstream: false,
    terminal: true,
  },

  /**
   * PDF 合并（2026-10-22 用户指令）：上游 OO 文档（ooword/oosheet）按输入口顺序
   * 转 PDF 并拼接成一个文件。合并/导出在节点卡片上操作。
   */
  pdfmerge: {
    /* ★ 2026-11-07 用户指令「PDF合并节点 名称改成PDF生成」 */
    label: 'PDF生成',
    role: 'output',
    forms: ['file'],
    family: 'pdf',
    shape: 'flow',
    allowsDownstream: false,
    terminal: true,
  },

  /**
   * 照片/照片组（2026-10-23 用户指令）：承载多张照片，连入 PDF 合并节点
   * （每图一页 A4 适配）。导入/管理在节点卡片上操作。
   */
  pdf: {
    label: 'PDF',
    role: 'library',
    forms: ['file'],
    family: 'pdf',
    shape: 'flow',
  },
  photo: {
    label: '照片',
    role: 'library',
    forms: ['file'],
    family: 'photo',
    shape: 'flow',
  },

  /**
   * 文件夹节点（2026-10-25 用户指令「作为读取位置和保存位置」）：
   * 指向本机目录——读模式列文件，写模式作保存输出目录。
   */
  folder: {
    label: '文件夹',
    role: 'library',
    forms: ['file'],
    family: 'folder',
    shape: 'flow',
  },

  /**
   * Excel转DXF（2026-11-10 用户指令「增加一个 excel转dxf 节点」）：
   * 表格坐标列 → 展点图（DXF R12：POINT/TEXT/LINE），生成文件下载或落盘。
   */
  xlsx2dxf: {
    label: 'Excel转DXF',
    role: 'action',
    forms: ['file'],
    family: 'xlsx',
    shape: 'flow',
  },

  /**
   * 网址监控（2026-11-09 用户指令「增加一个网址监控节点…提取、记录、或者输出数据」）：
   * 轮询 URL 提取内容记录变化（参照 changedetection.io 能力集）。
   */
  webmon: {
    label: '网址监控',
    role: 'library',
    forms: ['file'],
    family: 'folder',
    shape: 'flow',
  },

  /**
   * PDF 替换页节点（2026-10-25 用户指令「增加一个PDF替换页的节点」）：
   * 上组=被替换 PDF、下组=替换页来源——执行按行配对替换页，产物入历史。
   */
  pdfreplace: {
    label: 'PDF替换页',
    role: 'library',
    forms: ['file'],
    family: 'pdf',
    shape: 'flow',
  },

  /**
   * PDF 瘦身（2026-10-26 用户指令「拖PDF进空白区域弹选择」）：承载多份 PDF——
   * 压缩瘦身产物在节点上操作（config 同 pdf：pdfItems/pdfFileOrder，字节走 pdfitem-* IDB）
   */
  pdfslim: {
    label: 'PDF瘦身',
    role: 'library',
    forms: ['file'],
    family: 'pdf',
    shape: 'flow',
  },

  /**
   * 文件输出（2026-10-26 用户指令「输出节点 可以存储文件到固定位置」）：
   * 汇聚型终点——连线上游产物，一键保存到固定目录（outDir）。族色复用 output（同为输出语义）。
   */
  fileout: {
    label: '文件输出',
    role: 'output',
    forms: ['file'],
    family: 'output',
    shape: 'flow',
    allowsDownstream: false,
    terminal: true,
  },

  /**
   * PDF转Word（2026-10-26 用户指令「利用百度的开源 增加个节点 pdf2word」）：
   * 汇聚型终点——本机 PaddleOCR（百度飞桨开源）引擎把 PDF 还原成可编辑 docx，
   * 产物入节点历史。引擎外挂（不进安装包），缺失时节点上给安装指引。
   */
  pdf2word: {
    label: 'PDF转word',
    role: 'output',
    forms: ['file'],
    family: 'pdf',
    shape: 'flow',
    allowsDownstream: false,
    terminal: true,
  },

  // ── 输出型 ──────────────────────────────────────────────
  output: {
    label: '数据输出',
    role: 'output',
    forms: ['file'],
    family: 'output',
    shape: 'flow-out',
    terminal: true,
    allowsDownstream: false,
  },

  // ── 资料库型 ────────────────────────────────────────────
  table: {
    label: '表',
    role: 'library',
    forms: ['table'],
    family: 'table',
    shape: 'store',
    allowsDownstream: true,
  },

  // ── 模板型（role=output，form 里有 file/doc）─────────────
  docx: {
    label: 'Word',
    role: 'output',
    forms: ['doc', 'file'],
    family: 'docx',
    shape: 'page',
    terminal: true,
    allowsDownstream: false,
  },
  xlsx: {
    label: 'Excel',
    role: 'output',
    forms: ['table', 'file', 'doc'],
    family: 'xlsx',
    shape: 'page',
    terminal: true,
    allowsDownstream: false,
  },
  ooword: {
    label: 'OO 文档',
    role: 'output',
    forms: ['doc', 'file'],
    family: 'ooword',
    shape: 'page',
    terminal: true,
    allowsDownstream: false,
  },
  oosheet: {
    label: 'OO 工作簿',
    role: 'output',
    forms: ['table', 'file', 'doc'],
    family: 'oosheet',
    shape: 'page',
    terminal: true,
    allowsDownstream: false,
  },
  /**
   * 规则节点（2026-10-05，用户要求「画布内容中增加个圆形的节点。类似于功能节点。
   * 为规则适配节点。」）
   *
   * ★ 它不干活，职责是**给连上的节点发规则**（样式锁定 / 统一）：
   *   自己不做数据处理、不进计算引擎、不产出记录 —— 纯控制节点。
   *   组内主辅关系：圆形节点 = 主（规则持有者，可编辑），
   *   被它连上的同类节点 = 辅（只读跟随）。
   *   用户原话：「一旦匹配了，默认不能每个改规则，只能改主流程的节点进行修改。
   *   修改后所有相同规则的节点规则同步改。」
   *
   * shape 复用 'trigger'（圆形）—— 形状在本体系里表达**角色**，
   * 圆形 = 触发器语义（它「触发」一次规则下发），不新造形状。
   */
  rule: {
    label: '规则跟随',
    role: 'action',
    forms: [],
    family: 'rule',
    shape: 'trigger',
    allowsDownstream: true,
  },

  /**
   * AI 节点（2026-10-06，用户：「模仿规则跟随节点样式 增加一个AI节点。
   * AI节点的连线放在普通节点的 下1锚点位置」）
   *
   * ★ 与 rule 同构的圆形控制节点：吃数据（普通节点下1 出线 → AI），
   *   自身无字段、不进计算引擎。连线落点/方向见 ruleSync.triggerLineSemantics。
   * shape 'trigger'（圆形）—— 与 rule 一致，「模仿规则跟随节点样式」。
   */
  ai: {
    label: 'AI',
    role: 'action',
    forms: [],
    family: 'ai',
    shape: 'trigger',
    allowsDownstream: true,
  },

  /**
   * 参数库节点（2026-10-06，用户：「模仿规则跟随节点样式 增加一个参数库节点。
   * 参数库节点的连线放在普通节点的 上2锚点位置」）
   *
   * ★ 圆形供数节点：出线 → 普通节点上2（参数库是上游，下游普通节点把它
   *   记进 linkedInputs —— 与数据流同构）。role 取 'library'（资料库型，
   *   存数据供引用），形状仍用 'trigger' —— 用户显式要求圆形样式，
   *   「形状表达角色」的图例纯度在此为用户口径让路（hint 已解释语义）。
   */
  paramlib: {
    label: '参数库',
    role: 'library',
    forms: [],
    family: 'paramlib',
    shape: 'trigger',
    allowsDownstream: true,
  },

  /**
   * 节点组（2026-10-21 用户指令）：视觉容器——框住多个节点（Ctrl+G 建组 /
   * 从节点库拖入），整组拖动、边框可拖拽调整大小。无字段、不进计算引擎。
   * shape 'group'（新）：调色板里是矩形框轮廓。
   */
  group: {
    label: '节点组',
    role: 'action',
    forms: [],
    family: 'group',
    shape: 'group',
    allowsDownstream: false,
  },

  /**
   * 章节工作簿 / ooec（2026-10-05 新增，10-05 晚经用户两次纠正后定案）
   *
   * ★ 定案口径（用户原话，勿再改）：
   *   「ooec 是数据处理型节点」
   *   「ooec 从 ooexcel 上改的。」
   *   「衍生节点都要这样做」→血缘登记在下方 derivesFrom
   *
   * 我先前建错两处，此处已改正：
   *   ① role 错建 'output' + terminal + 不许下游 → 改 'process'（数据处理型）
   *   ② 内核错挂在 xlsx（@office-kit/xlsx 自研）→ 应派生自 oosheet（OnlyOffice 内核）
   *
   * ★ 派生自 oosheet 的架构含义（这是「升级时能定位」的根据）：
   *   oosheet 内核在 iframe 里，单元格读写**必须穿插件 postMessage 桥**
   *   （宿主直调 callCommand 静默失效，已调研确认）。故本章节的差异化层
   *   只允许落在「插件通道 + 纯计算层」，绝不能写进 OO 内核接触面。
   *
   * 隔离红线（scripts/oo-scan.mjs 有自动检查，破了立刻报红）：
   *   outline.ts / outlineAdapter.ts / ChapterDialog.tsx 必须零 OO 依赖。
   *   倒序数第三层往下的全部与 OO 版本无关 —— 换版时一行都不用动。
   */
  chapter: {
    label: 'OO 材料',
    role: 'process',
    forms: ['table', 'file', 'doc'],
    family: 'chapter',
    shape: 'page',
    terminal: false,
    allowsDownstream: true,
    derivesFrom: {
      kind: 'oosheet',
      generation: 'ooec',
      adaptations: [
        { area: '内核切换', note: '底座由 @office-kit/xlsx 改为 OnlyOffice（iframe 内核）' },
        { area: '插件通道', note: '读值 A16 / 插行 A17 / 删行 A18 —— 宿主直调无效，必走 cbref 插件' },
        { area: '章节分级', note: '第一列字符标记（#/##）为唯一真相，纯计算层 outline.ts' },
        { area: '章节编号', note: '第二列自动编号，renumber() 派生，永不手工维护' },
        { area: '列类型层', note: '第三列起右键选类型（名称/材料库/参数引用/单位/数量/系数/备注）' },
        { area: '写盘边界', note: '章节层只写前两列 + name 列，其余类型列一律不碰' },
        { area: '联动', note: 'param列存纯 URN，否则 isPureUrn 判否导致联动永不命中' },
      ],
    },
  },
}

/** 七类的中文名与顺序（节点库面板按此分组） */
export const ROLE_ORDER: NodeRole[] = ['action', 'input', 'process', 'library', 'template', 'loop', 'output']

export const ROLE_LABEL: Record<NodeRole, string> = {
  action: '动作型',
  input: '输入型',
  process: '数据处理型',
  library: '资料库型',
  template: '模板型',
  loop: '循环型',
  output: '输出型',
}

/** 形状 → 调色板 CSS 类后缀（形状表达角色） */
export const SHAPE_CLASS: Record<KindStyle['shape'], string> = {
  'flow-in': 'flow-in',
  flow: 'flow',
  'flow-out': 'flow-out',
  trigger: 'trigger',
  gate: 'gate',
  store: 'store',
  page: 'page',
  group: 'group',
}

/** 取某个 kind 的样式规则（未知 kind 兜底为 library 族，避免回退成默认灰） */
export function styleOf(kind: string): KindStyle {
  return (
    KIND_STYLE[kind] ?? {
      label: kind,
      role: 'library',
      forms: ['table'],
      family: 'table',
      shape: 'store',
      allowsDownstream: true,
    }
  )
}

/** 徽章 CSS 类后缀：无色族返回空串（终端节点不画徽章） */
export function badgeClassOf(kind: string): string {
  const f = styleOf(kind).family
  return f ? `rf-badge--${f}` : ''
}

/**
 * 汇聚型节点 —— 自身不带数据，靠连线吃上游。
 *
 * ★ 为什么要这个函数（2026-10-05）：
 *   判定原本在 core/engine.ts 与 App.tsx 各写一遍长串`||`。
 *   加 chapter（章节工作簿）时漏改一处 → 节点在计算引擎里有数据、
 *   在画布连线判定里却没有，症状是「线连上了但数据不来」，极难查。
 *   单一真相源铁律同样适用于**行为判定**，不只适用于颜色。
 *
 *   新增汇聚型 kind 时只改这一处。
 */
/* ★ 2026-10-23 修「前面的节点连不上 PDF 节点」：pdfmerge 是汇聚型
   （多上游 OO 文档按输入口顺序汇入转 PDF）——不在集合里时 handleConnect
   走不进汇聚分支，config.inputs 永不写入 → PDF 节点恒「无上游」。 */
const MERGE_KINDS = new Set(['sort', 'output', 'map', 'docx', 'xlsx', 'chapter', 'pdfmerge', 'fileout', 'pdf2word'])

/** 是否汇聚型节点（连线 = 上游数据流，面板可勾选来源） */
export function isMergeKind(kind: string | undefined): boolean {
  return !!kind && MERGE_KINDS.has(kind)
}

/**
 * 工作簿类节点 —— 走 WorkbookEditor 那条路径（xlsx 与 chapter）。
 *两者共用模板库绑定（xlsxTemplateId/Name），差别只在章节编排插槽。
 */
const WORKBOOK_KINDS = new Set(['xlsx', 'chapter'])

/** 是否工作簿类节点 */
export function isWorkbookNodeKind(kind: string | undefined): boolean {
  return !!kind && WORKBOOK_KINDS.has(kind)
}

/* ────────────────────────── 规则节点（样式锁定）──────────────────────────
   用户 2026-10-05：「圆形节点是对连线的节点进行规则锁定（统一）」。
   规则数据与同步逻辑都在 ruleSync.ts，本处只放 kind 级行为判定，
   避免「判定写两处→ 漏改一处」的老问题（同 isMergeKind 的由来）。
   ────────────────────────────────────────────────────────────────── */

/** 规则节点 kind（圆形，规则持有者） */
export function isRuleNodeKind(kind: string | undefined): boolean {
  return kind === 'rule'
}

/* ────────────────────────── 圆形节点家族（rule / ai / paramlib）────────────
   用户 2026-10-06：「模仿规则跟随节点样式 增加一个AI节点 / 参数库节点」。
   三种圆形节点共用一套画布渲染（圆形、无字段、单口、不折叠），
   差异只在连线端口与族色。判定集中在此，避免「谁是圆形」写散各处。
   ────────────────────────────────────────────────────────────────── */

/** 圆形节点家族：规则跟随 / AI / 参数库（画布上共用圆形形态） */
export function isRoundNodeKind(kind: string | undefined): boolean {
  return kind === 'rule' || kind === 'ai' || kind === 'paramlib'
}

/** 节点组（视觉容器）：框住多个节点整体拖动/管理 */
export function isGroupNodeKind(kind: string | undefined): boolean {
  return kind === 'group'
}

/** 节点组适配几何（2026-10-21 用户「适应大小的按钮和拖拽的扩展 不在同一位置」）：
 *  所有「组框包住某矩形」的路径（适应大小钮 / 拖入扩展 / 成员拖停扩展）
 *  统一走 groupFitRect——同一内边距约定，终点几何完全一致。 */
export const GROUP_PAD = 24
export const GROUP_HEAD = 34
export function groupFitRect(rect: { x: number; y: number; w: number; h: number }) {
  /* ★ 2026-10-21 用户「右边多了 下方少了」：顶部要给标题栏留 GROUP_HEAD+PAD/2，
     但底部只留 GROUP_PAD——旧公式把 HEAD 又加进了高度（底部只剩 +2px）。 */
  const top = GROUP_HEAD + GROUP_PAD / 2
  return {
    x: Math.round(rect.x - GROUP_PAD),
    y: Math.round(rect.y - top),
    w: Math.round(rect.w + GROUP_PAD * 2),
    h: Math.round(rect.h + top + GROUP_PAD),
  }
}
export function unionRect(
  a: { x: number; y: number; w: number; h: number },
  b: { x: number; y: number; w: number; h: number },
) {
  const x = Math.min(a.x, b.x)
  const y = Math.min(a.y, b.y)
  return { x: Math.round(x), y: Math.round(y), w: Math.round(Math.max(a.x + a.w, b.x + b.w) - x), h: Math.round(Math.max(a.y + a.h, b.y + b.h) - y) }
}

/** 我是否处于「已折叠」的节点组内（组头 groupFolded=true → 成员折叠成方块） */
export function isMemberOfFoldedGroup(
  nodeId: string,
  tables: Record<string, { kind?: string; config?: { groupFolded?: boolean; members?: string[] } } | undefined>,
): boolean {
  for (const t of Object.values(tables)) {
    if (!t || t.kind !== 'group' || t.config?.groupFolded !== true) continue
    if (((t.config?.members ?? []) as string[]).includes(nodeId)) return true
  }
  return false
}

/**
 * 规则节点的锁定范围 = **沿连线可达的普通节点**。
 *
 * ★ 为什么只锁「沿连线可达」而不锁「同 kind 全部」：
 *   用户要的是「对连线的节点进行规则锁定」，连线是锁的作用域。
 *   若锁全画布同 kind，一个节点被锁会让远处无关节点跟着变，且无从撤销。
 *   连线可见 = 影响范围可见，这条边界必须由数据推导而非由 UI 假装。
 */

/* ────────────────────────── 血缘（衍生关系）──────────────────────────
   用户 2026-10-05：「在节点系统上增加一个衍生的关联标记，
   哪个是哪个节点衍生的，做了哪些适配。都有记录。」
   —— 目���不是「记一下」，而是**在 UI 上看得见**，
      且升级父内核时能按图索骥查出该改哪。
   ────────────────────────────────────────────────────────────────── */

/** 血缘声明（KindStyle.derivesFrom 的具名形态，便于函数签名引用） */
export type DerivesFrom = NonNullable<KindStyle['derivesFrom']>

/** 取某kind 的血缘声明；原始节点返回 null（不是 undefined —— 便于 `== null` 判断） */
export function derivesFromOf(kind: string | undefined): DerivesFrom | null {
  if (!kind) return null
  return KIND_STYLE[kind]?.derivesFrom ?? null
}

/** 是否衍生节点（有血缘声明） */
export function isDerivedKind(kind: string | undefined): boolean {
  return !!derivesFromOf(kind)
}

/** 是否参与 OO 迭代影响面（血缘链上有 OO 节点，或自身就是 OO 节点）。
 *  scripts/oo-scan.mjs 用同一口径，两边不许各写一份。 */
const OO_KIND_TOKENS = ['oo']

export function isOOLineage(kind: string | undefined): boolean {
  if (!kind) return false
  if (OO_KIND_TOKENS.some((t) => kind.includes(t))) return true
  const d = derivesFromOf(kind)
  return !!d && OO_KIND_TOKENS.some((t) => d.kind.includes(t))
}

/**
 * 沿血缘链上溯，返回从自身到根的 kind 序列（含自身）。
 * 例：chapter → ['chapter', 'oosheet']。无血缘则只返回自身。
 */
export function lineageChainOf(kind: string | undefined): string[] {
  if (!kind) return []
  const chain: string[] = [kind]
  const seen = new Set<string>([kind])
  let cur: DerivesFrom | null = derivesFromOf(kind)
  // 防御性上限：血缘若成环（a→b→a），5 步足够暴露问题，不无限循环
  for (let i = 0; i < 5 && cur; i++) {
    if (seen.has(cur.kind)) break
    chain.push(cur.kind)
    seen.add(cur.kind)
    cur = derivesFromOf(cur.kind)
  }
  return chain
}

/** 全部衍生节点清单（UI 血缘面板与扫描器共用） */
export function derivedKinds(): string[] {
  return Object.keys(KIND_STYLE).filter((k) => !!KIND_STYLE[k].derivesFrom)
}

/** 节点族CSS 类后缀（节点头/边框取色用） */
export function familyClassOf(kind: string): string {
  const f = styleOf(kind).family
  return f ? `nk--${f}` : ''
}

/** 调色板形状类（完整类名）：形状表达角色 */
export function shapeClassOf(kind: string): string {
  return `palette-${SHAPE_CLASS[styleOf(kind).shape]}-shape`
}

/**
 * 调色板条目的完整类名 = 形状类 + 族类。
 *
 * 为什么两个都要：形状类只区分「角色」，若不带族色，同角色的多个 kind 会共用
 * 一份中性灰底灰框（实测 sort/map/text 三个都渲染成灰块，看不出差别）。
 * 族类（.nk--<族>）负责取族色，形状类负责轮廓 —— 两个维度正交，缺一不可。
 */
export function paletteItemClassOf(kind: string): string {
  return `${shapeClassOf(kind)} ${familyClassOf(kind)}`
}

/** 按 role 分组，供节点库面板渲染七类折叠 */
export function groupedByRole(kinds: string[]): { role: NodeRole; label: string; kinds: string[] }[] {
  const out = new Map<NodeRole, string[]>()
  for (const k of kinds) {
    const r = styleOf(k).role
    if (!out.has(r)) out.set(r, [])
    out.get(r)!.push(k)
  }
  return ROLE_ORDER.filter((r) => out.has(r)).map((role) => ({
    role,
    label: ROLE_LABEL[role],
    kinds: out.get(role)!,
  }))
}

/* ── 节点库面板数据 ────────────────────────────────────────
   顺序 = ROLE_ORDER（动作→输入→处理→资料库→模板→循环→输出），按流水线心智排：
   先「什么时候跑」，再「数据从哪来」，然后「怎么算」，最后「出什么」。 */

/** 每类的一句话说明（调色板分组标题 + 拖拽提示） */
const ROLE_HINT: Record<NodeRole, string> = {
  action: '决定流程什么时候执行：手动 / 定时 / Webhook',
  input: '数据从外部进来：文件导入或参数录入',
  process: '对数据做加工：排序、查表补列、文本',
  library: '存数据供引用，本身不产生产物',
  template: '套模板出交付稿：域绑定后灌数据生成文档',
  loop: '设置执行条件，直到满足要求为止',
  output: '流程终点：导出文件，不接受下游',
}

/**
 * kind 级提示覆盖（2026-10-05）
 *
 * 为什么需要：默认 hint 是 `${label} —— ${ROLE_HINT[role]}`，即「条目名 + 所属角色的组提示」。
 * 大多数 kind 确实能被组提示准确描述，但**规则节点不行**：
 * 它的 role 是 'action'（放在动作型分组里），而组提示是「决定流程什么时候执行：手动/定时/Webhook」——
 * 用户拖出来看到的是一句彻底误导的话，会以为规则节点管执行时机。
 *
 * 症状隐蔽：节点库一切正常、拖得出来、连得上，只是**说明文字在说另一件事**，
 * 不实际点开属性面板发现不了。
 *
 * 判据：kind 的实际语义与其 role 的组提示**不是同一件事**时，在此覆盖。
 */
const KIND_HINT_OVERRIDE: Record<string, string> = {
  group: '节点组：框住多个节点整体管理。选中节点按 Ctrl+G 建组；拖入画布后拖标题栏移动整组、右下角拖拽调大小',
  rule: '规则跟随：连上几个同类节点，在本节点上改一次，它们一起变',
  ai: 'AI：从普通节点的下1口出线连到本节点，接收它的数据做 AI 处理',
  paramlib: '参数库：出线连到普通节点的上1口，为它提供参数',
  test_table: '数据表节点：手动输入参数与参数数据（双击节点打开编辑弹窗），供 OO Excel 引用调试',
}

/** 节点库面板用的完整分组数据 */
export interface PaletteGroup {
  role: NodeRole
  label: string
  hint: string
  kinds: { kind: string; label: string; hint: string }[]
}

export const PALETTE_GROUPS: PaletteGroup[] = (() => {
  /* ★ 2026-10-19 用户指令「把AI节点、规则跟随节点、参数库节点 放在一类」：
   *  三家原本按 role 散在「动作型」（rule/ai）和「资料库型」（paramlib）两组，
   *  抽出来合成一个「圆形节点」组（它们行为同族：拖到普通节点上直接连线）。 */
  const ROUND_KINDS = ['rule', 'ai', 'paramlib']
  const groups = groupedByRole(Object.keys(KIND_STYLE)).map((g) => ({
    role: g.role,
    label: g.label,
    hint: ROLE_HINT[g.role],
    kinds: g.kinds
      .filter((k) => !ROUND_KINDS.includes(k))
      .map((k) => ({
        kind: k,
        label: KIND_STYLE[k].label,
        hint: KIND_HINT_OVERRIDE[k] ?? `${KIND_STYLE[k].label} —— ${ROLE_HINT[KIND_STYLE[k].role]}`,
      })),
  }))
  const roundKinds = ROUND_KINDS.filter((k) => KIND_STYLE[k])
  if (roundKinds.length) {
    groups.push({
      role: 'action',
      label: '圆形节点（规则 / AI / 参数库）',
      hint: '拖到普通节点上直接连线；也可从锚点拉线',
      kinds: roundKinds.map((k) => ({
        kind: k,
        label: KIND_STYLE[k].label,
        hint: KIND_HINT_OVERRIDE[k] ?? `${KIND_STYLE[k].label} —— ${ROLE_HINT[KIND_STYLE[k].role]}`,
      })),
    })
  }
  return groups.filter((g) => g.kinds.length > 0)
})()