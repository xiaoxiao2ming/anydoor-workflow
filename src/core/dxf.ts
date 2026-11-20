/* ─────────────────────────────────────────────────────────────
 * Excel转DXF 内核（2026-11-10 用户指令「增加一个 excel转dxf 节点」）
 *
 * 手写 DXF R12 ASCII（零依赖）：全 CAD 兼容（AutoCAD/中望/浩辰）。
 * 典型场景：测量坐标表（点号/X/Y/高程）→ 展点（POINT + 点号 TEXT + 可选连线）。
 * 纯函数：pickPoints（行→点，含容错）与 pointsToDxf（点→DXF 文本）分开，方便单测。
 * ───────────────────────────────────────────────────────────── */

export interface DxfPoint {
  name: string
  x: number
  y: number
  z?: number
}

export interface DxfMapSpec {
  nameCol?: string
  xCol: string
  yCol: string
  zCol?: string
}

export interface DxfOptions {
  drawPoint: boolean
  drawLabel: boolean
  drawLine: boolean
  /** 标注文字高度（图面单位，默认 2.0） */
  textHeight: number
  /** 标注相对点的偏移（图面单位，默认 1.5） */
  labelOffset: number
  /** Y 轴取反（数据为屏幕坐标时用；测量坐标 Y=北 向上无需翻转） */
  flipY: boolean
}

export const DXF_DEFAULT_OPTIONS: DxfOptions = {
  drawPoint: true,
  drawLabel: true,
  drawLine: false,
  /** ★ 字高 3.5（2026-11-10 用户指定，勘察制图标准注释字高；配 dytxt+hztxt 大字体、宽度 0.8） */
  textHeight: 3.5,
  labelOffset: 1.5,
  flipY: false,
}

const LAYERS = [
  { name: 'CB_POINTS', aci: 1 },
  { name: 'CB_TEXTS', aci: 2 },
  { name: 'CB_LINES', aci: 5 },
]

/* 表格整图图层 */
export const TABLE_LAYERS = [
  { name: 'CB_TABLE_LINES', aci: 7 },
  { name: 'CB_TABLE_TEXT', aci: 3 },
]

const fmt = (n: number): string => {
  const s = n.toFixed(3)
  return s.replace(/\.?0+$/, '') || '0'
}

/** 行数组 → 点位数组。x/y 解析失败的行跳过并记原因（不中断）。 */
export function pickPoints(
  rows: Array<Record<string, unknown>>,
  map: DxfMapSpec,
): { points: DxfPoint[]; skipped: string[] } {
  const points: DxfPoint[] = []
  const skipped: string[] = []
  rows.forEach((row, i) => {
    const no = i + 2 /* 表头占第 1 行，数据从 2 起 */
    const num = (col?: string): number | null => {
      if (!col) return null
      const v = row[col]
      if (v == null || v === '') return null
      const n = Number(v)
      return Number.isFinite(n) ? n : null
    }
    const x = num(map.xCol)
    const y = num(map.yCol)
    if (x == null || y == null) {
      const why = x == null && y == null ? `X/Y 列（${map.xCol}/${map.yCol}）都非数值` : x == null ? `X 列（${map.xCol}）非数值` : `Y 列（${map.yCol}）非数值`
      skipped.push(`第 ${no} 行：${why}`)
      return
    }
    const z = num(map.zCol)
    const rawName = map.nameCol ? row[map.nameCol] : ''
    /* 空点名按有效点序补号（跳过坏行不跳号，避免 P 序列出现空洞） */
    const name = String(rawName ?? '').trim() || `P${points.length + 1}`
    points.push({ name, x, y, z: z ?? undefined })
  })
  return { points, skipped }
}

/** 点位数组 → DXF R12 ASCII 全文（HEADERS/TABLES 图层/ENTITIES 实体/EOF）。 */
export function pointsToDxf(points: DxfPoint[], opts: DxfOptions): string {
  const yOf = (y: number): number => (opts.flipY ? -y : y)
  const L: string[] = []
  const pair = (code: number, value: string | number): void => {
    L.push(String(code).padStart(3, ' '))
    L.push(String(value))
  }

  pair(0, 'SECTION')
  pair(2, 'HEADER')
  pair(9, '$ACADVER')
  pair(1, 'AC1009') /* R12 */
  pair(0, 'ENDSEC')

  pair(0, 'SECTION')
  pair(2, 'TABLES')
  pair(0, 'TABLE')
  pair(2, 'LAYER')
  pair(70, LAYERS.length)
  for (const ly of LAYERS) {
    pair(0, 'LAYER')
    pair(2, ly.name)
    pair(70, 0)
    pair(62, ly.aci)
    pair(6, 'CONTINUOUS')
  }
  pair(0, 'ENDTAB')
  pair(0, 'TABLE')
  pair(2, 'STYLE')
  pair(70, 1)
  pair(0, 'STYLE')
  pair(2, 'CB_CN')
  pair(70, 0)
  pair(40, 0.0)
  pair(41, 0.8)
  pair(50, 0.0)
  pair(71, 0)
  pair(42, 3.5)
  pair(3, 'dytxt.shx')
  pair(4, 'hztxt.shx')
  pair(0, 'ENDTAB')
  pair(0, 'ENDSEC')

  pair(0, 'SECTION')
  pair(2, 'ENTITIES')
  for (const p of points) {
    const y = yOf(p.y)
    if (opts.drawPoint) {
      pair(0, 'POINT')
      pair(8, 'CB_POINTS')
      pair(10, fmt(p.x))
      pair(20, fmt(y))
      if (p.z != null) pair(30, fmt(p.z))
    }
    if (opts.drawLabel) {
      pair(0, 'TEXT')
      pair(8, 'CB_TEXTS')
      pair(10, fmt(p.x + opts.labelOffset))
      pair(20, fmt(y + opts.labelOffset))
      pair(40, fmt(opts.textHeight))
      pair(7, 'CB_CN')
      pair(1, p.name)
    }
  }
  if (opts.drawLine) {
    for (let i = 0; i + 1 < points.length; i++) {
      const a = points[i]
      const b = points[i + 1]
      pair(0, 'LINE')
      pair(8, 'CB_LINES')
      pair(10, fmt(a.x))
      pair(20, fmt(yOf(a.y)))
      pair(11, fmt(b.x))
      pair(21, fmt(yOf(b.y)))
    }
  }

  pair(0, 'ENDSEC')
  pair(0, 'EOF')
  return L.join('\r\n') + '\r\n'
}

/* ── 非空边界整表转 DXF（2026-11-10 用户澄清「边界确定是非空边界」）──────────
 * 把表格的非空区域（used range：过滤全空行/全空列）整块画成 DXF 表格图：
 * 单元格文字（TEXT）+ 网格线（LINE），列宽按内容自适应。任何 Excel 表通用（不限坐标表）。 */

const cellText = (v: unknown): string => String(v ?? '').replace(/\s+/g, ' ').trim()

/** 非空边界裁剪：返回去掉全空行/全空列后的（列名数组, 行数组） */
export function trimToUsedRange(
  fields: string[],
  rows: Array<Record<string, unknown>>,
): { cols: string[]; rows: Array<Record<string, unknown>> } {
  const colUsed = fields.map((_, ci) => rows.some((r) => cellText(r[fields[ci]]) !== ''))
  const cols = fields.filter((_, ci) => colUsed[ci])
  const out = rows
    .map((r) => {
      const o: Record<string, unknown> = {}
      for (const c of cols) o[c] = r[c]
      return o
    })
    .filter((r) => cols.some((c) => cellText(r[c]) !== ''))
  return { cols, rows: out }
}

export interface TableDxfOptions {
  /** 单元格文字高度（默认 2.5） */
  textHeight?: number
}

/** 非空边界整表 → DXF 表格图（网格线 + 全部单元格文字；表头在最上） */
export function tableToDxf(
  fields: string[],
  rows: Array<Record<string, unknown>>,
  opts?: TableDxfOptions,
): { dxf: string; usedRows: number; usedCols: number } {
  const th = opts?.textHeight ?? 2.5
  const { cols, rows: used } = trimToUsedRange(fields, rows)
  /* 列宽：列内最长文本 × 近似字符宽（中文≈1.4×th）+ 左右各 0.5×th 内边距 */
  const widthOf = (text: string): number => {
    let w = 0
    for (const ch of text) w += ch.charCodeAt(0) > 255 ? 1.4 : 0.62
    return w * th
  }
  const colW = cols.map((c) => {
    const cells = [c, ...used.map((r) => cellText(r[c]))]
    return Math.max(th * 2, ...cells.map((t) => widthOf(t) + th))
  })
  const rh = th * 2
  const totalW = colW.reduce((a, b) => a + b, 0)
  const totalH = rh * (used.length + 1) /* 表头 + 数据行 */

  const L: string[] = []
  const pair = (code: number, value: string | number): void => {
    L.push(String(code).padStart(3, ' '))
    L.push(String(value))
  }
  const text = (x: number, y: number, content: string, layer: string): void => {
    pair(0, 'TEXT')
    pair(8, layer)
    pair(10, fmt(x))
    pair(20, fmt(y))
    pair(40, fmt(th))
    pair(7, 'CB_CN')
    pair(1, content)
  }
  const line = (ax: number, ay: number, bx: number, by: number): void => {
    pair(0, 'LINE')
    pair(8, 'CB_TABLE_LINES')
    pair(10, fmt(ax))
    pair(20, fmt(ay))
    pair(11, fmt(bx))
    pair(21, fmt(by))
  }

  pair(0, 'SECTION')
  pair(2, 'HEADER')
  pair(9, '$ACADVER')
  pair(1, 'AC1009')
  pair(0, 'ENDSEC')
  pair(0, 'SECTION')
  pair(2, 'TABLES')
  pair(0, 'TABLE')
  pair(2, 'LAYER')
  pair(70, 2)
  for (const ly of TABLE_LAYERS) {
    pair(0, 'LAYER')
    pair(2, ly.name)
    pair(70, 0)
    pair(62, ly.aci)
    pair(6, 'CONTINUOUS')
  }
  pair(0, 'ENDTAB')
  pair(0, 'TABLE')
  pair(2, 'STYLE')
  pair(70, 1)
  pair(0, 'STYLE')
  pair(2, 'CB_CN')
  pair(70, 0)
  pair(40, 0.0)
  pair(41, 0.8)
  pair(50, 0.0)
  pair(71, 0)
  pair(42, 3.5)
  pair(3, 'dytxt.shx')
  pair(4, 'hztxt.shx')
  pair(0, 'ENDTAB')
  pair(0, 'ENDSEC')
  pair(0, 'SECTION')
  pair(2, 'ENTITIES')

  /* 网格：横线 = 数据行+2（表头顶/表头分隔/行分隔/底），纵线 = 列+1 */
  const xEdge: number[] = [0]
  for (const w of colW) xEdge.push(xEdge[xEdge.length - 1] + w)
  const yEdge: number[] = [0]
  for (let i = 0; i < used.length + 1; i++) yEdge.push(yEdge[yEdge.length - 1] + rh)
  for (const y of yEdge) line(0, y, totalW, y)
  for (const x of xEdge) line(x, 0, x, totalH)

  /* 表头（最上行） */
  cols.forEach((c, j) => text(xEdge[j] + th * 0.5, yEdge[yEdge.length - 2] + th * 0.5, c, 'CB_TABLE_TEXT'))
  /* 数据行（自上而下） */
  used.forEach((r, i) => {
    const yBase = yEdge[yEdge.length - 3 - i]
    cols.forEach((c, j) => {
      const t = cellText(r[c])
      if (t) text(xEdge[j] + th * 0.5, yBase + th * 0.5, t, 'CB_TABLE_TEXT')
    })
  })

  pair(0, 'ENDSEC')
  pair(0, 'EOF')
  return { dxf: L.join('\r\n') + '\r\n', usedRows: used.length, usedCols: cols.length }
}

/** 预览解析（2026-11-10 用户「右键打开显示不对」）：读回 DXF 实体供卡片 SVG 渲染。
 *  只需解析本生成器输出的实体子集（POINT/TEXT/LINE），组码行 trim 兼容右对齐。 */
export interface DxfPreviewEnt {
  points: { x: number; y: number }[]
  texts: { x: number; y: number; h: number; text: string }[]
  lines: { x1: number; y1: number; x2: number; y2: number }[]
}
export function parseDxfEntities(dxf: string): DxfPreviewEnt {
  const rows = dxf.split(/\r?\n/)
  const out: DxfPreviewEnt = { points: [], texts: [], lines: [] }
  let cur: 'POINT' | 'TEXT' | 'LINE' | null = null
  let curEnt: Record<string, number | string> = {}
  const flush = (): void => {
    if (!cur) return
    const n = (k: string): number => Number(curEnt[k] ?? 0)
    if (cur === 'POINT') out.points.push({ x: n('10'), y: n('20') })
    else if (cur === 'TEXT') out.texts.push({ x: n('10'), y: n('20'), h: n('40') || 2.5, text: String(curEnt['1'] ?? '') })
    else if (cur === 'LINE') out.lines.push({ x1: n('10'), y1: n('20'), x2: n('11'), y2: n('21') })
    cur = null
    curEnt = {}
  }
  for (let i = 0; i + 1 < rows.length; i += 2) {
    const code = Number(rows[i].trim())
    const value = rows[i + 1]
    if (!Number.isFinite(code)) { i -= 1; continue }
    if (code === 0) {
      flush()
      if (value === 'POINT' || value === 'TEXT' || value === 'LINE') cur = value
      continue
    }
    if (!cur) continue
    if (code === 1 || code === 8) { if (code === 1) curEnt['1'] = value; continue }
    if ([10, 11, 20, 21, 30, 40].includes(code)) curEnt[String(code)] = Number(value)
  }
  flush()
  return out
}

/** 组合入口：行 → （DXF 全文, 点数, 跳过原因列表） */
export function rowsToDxf(
  rows: Array<Record<string, unknown>>,
  map: DxfMapSpec,
  opts: DxfOptions,
): { dxf: string; count: number; skipped: string[] } {
  const { points, skipped } = pickPoints(rows, map)
  return { dxf: pointsToDxf(points, opts), count: points.length, skipped }
}
