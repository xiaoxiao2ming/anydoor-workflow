/**
 * 计算列公式求值 —— 只为算出「缓存值」。
 *
 * 为什么必须有它：
 *   OOXML 的单元格可以只写 `<f>` 不写 `<v>`，Excel/WPS 打开时会自己算，
 *   但任何不执行公式的阅读器（预览、报表提取、部分 WPS 场景）会显示 #N/A。
 *   SP1 实测确认过：光写公式不写缓存值，落盘 XML 里连 `<v>` 元素都没有。
 *   所以生成 Excel 时必须 formula + cachedValue 双写，而缓存值得我们自己算出来。
 *
 * 实现：零依赖递归下降解析器 + formulajs（MIT，实测 373 个可调用函数）做函数库。
 * hyperformula 不用 —— 它是 GPLv3，闭源商用要买授权，且这里只需要"算一个值"，
 * 不需要它的依赖图/增量重算。
 */

import * as FX from '@formulajs/formulajs'

export interface EvalOutcome {
  ok: boolean
  value?: unknown
  error?: string
}

/* ────────────────── 词法 ────────────────── */

type Tok =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'val'; v: unknown }
  | { t: 'id'; v: string }
  | { t: 'op'; v: string }
  | { t: 'lp' }
  | { t: 'rp' }
  | { t: 'comma' }

/** 单元格值在公式文本里的占位符：\u0000<下标>\u0001。
 *  用占位符而不是直接把值贴进文本 —— 否则值里带的运算符/引号会破坏表达式（注入）。 */
const SENTINEL = /\u0000(\d+)\u0001/g

function tokenize(src: string, values: unknown[]): Tok[] {
  const out: Tok[] = []
  let i = 0
  // 先把占位符切成 val token
  const chunks: ({ kind: 'text'; s: string } | { kind: 'val'; v: unknown })[] = []
  let last = 0
  SENTINEL.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = SENTINEL.exec(src))) {
    if (m.index > last) chunks.push({ kind: 'text', s: src.slice(last, m.index) })
    chunks.push({ kind: 'val', v: values[Number(m[1])] })
    last = m.index + m[0].length
  }
  if (last < src.length) chunks.push({ kind: 'text', s: src.slice(last) })

  for (const ch of chunks) {
    if (ch.kind === 'val') {
      out.push({ t: 'val', v: ch.v })
      continue
    }
    const s = ch.s
    i = 0
    while (i < s.length) {
      const c = s[i]
      if (c === ' ' || c === '\t' || c === '\n' || c === '\r') {
        i++
        continue
      }
      if (c === '"') {
        // Excel 字符串：内部双引号转义为 ""
        let j = i + 1
        let str = ''
        while (j < s.length) {
          if (s[j] === '"') {
            if (s[j + 1] === '"') {
              str += '"'
              j += 2
              continue
            }
            break
          }
          str += s[j]
          j++
        }
        if (j >= s.length) throw new Error('字符串字面量未闭合')
        out.push({ t: 'str', v: str })
        i = j + 1
        continue
      }
      if (/[0-9]/.test(c) || (c === '.' && /[0-9]/.test(s[i + 1] ?? ''))) {
        const num = /^[0-9]*\.?[0-9]+([eE][+-]?[0-9]+)?/.exec(s.slice(i))
        if (!num) throw new Error('数字解析失败：' + s.slice(i, i + 12))
        out.push({ t: 'num', v: Number(num[0]) })
        i += num[0].length
        continue
      }
      if (/[A-Za-z_\u4e00-\u9fa5]/.test(c)) {
        const id = /^[A-Za-z_\u4e00-\u9fa5][A-Za-z0-9_.\u4e00-\u9fa5]*/.exec(s.slice(i))
        if (!id) throw new Error('标识符解析失败')
        out.push({ t: 'id', v: id[0] })
        i += id[0].length
        continue
      }
      if (c === '(') {
        out.push({ t: 'lp' })
        i++
        continue
      }
      if (c === ')') {
        out.push({ t: 'rp' })
        i++
        continue
      }
      if (c === ',') {
        out.push({ t: 'comma' })
        i++
        continue
      }
      const two = s.slice(i, i + 2)
      if (two === '<=' || two === '>=' || two === '<>') {
        out.push({ t: 'op', v: two })
        i += 2
        continue
      }
      if ('+-*/^%&=<>'.includes(c)) {
        out.push({ t: 'op', v: c })
        i++
        continue
      }
      throw new Error(`不认识的字符「${c}」`)
    }
  }
  return out
}

/* ────────────────── 语法 ────────────────── */

const FUNCS = FX as unknown as Record<string, (...a: unknown[]) => unknown>

/** Excel 的算术语义：能转数字就转，转不了抛错（而不是返回 NaN 悄悄污染结果） */
function toNum(v: unknown): number {
  if (typeof v === 'number') return v
  if (typeof v === 'boolean') return v ? 1 : 0
  if (v === null || v === undefined || v === '') return 0
  if (typeof v === 'string') {
    const t = v.trim()
    if (t === '') return 0
    const n = Number(t)
    if (!Number.isNaN(n)) return n
  }
  throw new Error(`「${String(v)}」不是数字`)
}

function toBool(v: unknown): boolean {
  if (typeof v === 'boolean') return v
  if (typeof v === 'number') return v !== 0
  const s = String(v ?? '').trim().toLowerCase()
  return !(s === '' || s === 'false' || s === '0' || s === '否')
}

class Parser {
  private p = 0
  constructor(private toks: Tok[]) {}

  private peek(): Tok | undefined {
    return this.toks[this.p]
  }
  private eat(): Tok {
    const t = this.toks[this.p++]
    if (!t) throw new Error('公式意外结束')
    return t
  }

  parse(): unknown {
    const v = this.expr()
    if (this.p < this.toks.length) throw new Error(`公式尾部有多余内容（第 ${this.p + 1} 个记号）`)
    return v
  }

  /** 比较（最低优先级） */
  private expr(): unknown {
    let l = this.concat()
    for (;;) {
      const t = this.peek()
      if (t && t.t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(t.v)) {
        this.eat()
        const r = this.concat()
        switch (t.v) {
          case '=':
            l = toNum(l) === toNum(r)
            break
          case '<>':
            l = toNum(l) !== toNum(r)
            break
          case '<':
            l = toNum(l) < toNum(r)
            break
          case '>':
            l = toNum(l) > toNum(r)
            break
          case '<=':
            l = toNum(l) <= toNum(r)
            break
          default:
            l = toNum(l) >= toNum(r)
        }
        continue
      }
      return l
    }
  }

  /** 字符串连接 & */
  private concat(): unknown {
    let l = this.add()
    for (;;) {
      const t = this.peek()
      if (t && t.t === 'op' && t.v === '&') {
        this.eat()
        const r = this.add()
        l = String(l ?? '') + String(r ?? '')
        continue
      }
      return l
    }
  }

  private add(): unknown {
    let l = this.mul()
    for (;;) {
      const t = this.peek()
      if (t && t.t === 'op' && (t.v === '+' || t.v === '-')) {
        this.eat()
        const r = this.mul()
        l = t.v === '+' ? toNum(l) + toNum(r) : toNum(l) - toNum(r)
        continue
      }
      return l
    }
  }

  private mul(): unknown {
    let l = this.pow()
    for (;;) {
      const t = this.peek()
      if (t && t.t === 'op' && (t.v === '*' || t.v === '/')) {
        this.eat()
        const r = this.pow()
        l = t.v === '*' ? toNum(l) * toNum(r) : toNum(l) / toNum(r)
        continue
      }
      return l
    }
  }

  private pow(): unknown {
    const l = this.unary()
    const t = this.peek()
    if (t && t.t === 'op' && t.v === '^') {
      this.eat()
      // 右结合：2^3^2 = 2^(3^2)
      return Math.pow(toNum(l), toNum(this.pow()))
    }
    return l
  }

  private unary(): unknown {
    const t = this.peek()
    if (t && t.t === 'op' && (t.v === '-' || t.v === '+')) {
      this.eat()
      const v = this.unary()
      return t.v === '-' ? -toNum(v) : toNum(v)
    }
    return this.primary()
  }

  private primary(): unknown {
    let v: unknown
    const t = this.eat()
    if (t.t === 'num') v = t.v
    else if (t.t === 'str') v = t.v
    else if (t.t === 'val') v = t.v
    else if (t.t === 'lp') {
      v = this.expr()
      const c = this.eat()
      if (c.t !== 'rp') throw new Error('缺少右括号')
    } else if (t.t === 'id') {
      const upper = t.v.toUpperCase()
      const nt = this.peek()
      if (nt && nt.t === 'lp') {
        this.eat()
        const args: unknown[] = []
        if (this.peek()?.t !== 'rp') {
          args.push(this.expr())
          while (this.peek()?.t === 'comma') {
            this.eat()
            args.push(this.expr())
          }
        }
        const c = this.eat()
        if (c.t !== 'rp') throw new Error(`函数 ${upper} 缺少右括号`)
        const fn = FUNCS[upper]
        if (typeof fn !== 'function') throw new Error(`不支持的函数：${upper}`)
        v = fn(...(args as never[]))
      } else if (upper === 'TRUE') v = true
      else if (upper === 'FALSE') v = false
      else throw new Error(`未定义的名称：${t.v}`)
    } else {
      throw new Error('表达式语法错误：意外的记号 ' + JSON.stringify(t))
    }

    // 后缀百分号：50% → 0.5
    while (this.peek()?.t === 'op' && (this.peek() as { v: string }).v === '%') {
      this.eat()
      v = toNum(v) / 100
    }
    return v
  }
}

/* ────────────────── 对外入口 ────────────────── */

const CELL_RE = /(?<![A-Za-z0-9_.])(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})(?![\d(])/g
const STR_LIT = /"([^"]|"")*"/g

/** 区间引用（#36）：A1:B3（冒号两侧允许空白）。断言同 CELL_RE，额外不许右邻冒号（防止 A1:B3 被拆成单格先吃）。 */
const RANGE_RE =
  /(?<![A-Za-z0-9_.:])(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})\s*:\s*(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})(?![\d:])/g
/** 区间展开护栏：超过这个格子数保留原文（解析器报错），防手滑 A1:Z99999 卡死求值 */
const RANGE_MAX_CELLS = 10000

/** 'B' → 2（1-based 列号）；'ABC' → 731 —— 区间展开用，保持本文件零项目内依赖（不走 ./a1） */
const colToIndexNum = (col: string): number => {
  let n = 0
  for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}
/** 2 → 'B'（1-based 列号反解） */
const colNumToIndex = (n: number): string => {
  let s = ''
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

/**
 * node:// 引用 token（批次 1-6 · 引用参与求值）：
 *   node://<nodeId>#<字段名>  —— target 不含 # 与空白；locator 可中文，终止于运算符/空白/标点。
 * `#` 可省（整节点引用）。字符串字面量不受影响 —— 替换只发生在 code 段。
 */
const URN_RE = /node:\/\/[^#\s+\-*/^&()<>=!,%]+(?:#[^\s+\-*/^&()<>=!,%]*)?/g

/**
 * 纯 URN 判定：整个公式体就是一个 node:// token（允许尾随空白）→ 整格引用，存 URN 文本；
 * 混排（token 参与四则/函数）→ 求值分支（批次 1-6）。
 */
export function isPureUrn(body: string): boolean {
  return new RegExp(`^${URN_RE.source}$`).test(body.trim())
}

/**
 * 求一行计算列的值。
 *
 * @param formulaBody 已经平移到目标行的公式体（不带 `=`），如 'B2*C2'
 * @param refValues   该行各单元格的值，键是 A1 坐标（忽略 `$`），如 { B2: 100, C2: 25.5 }
 * @param resolveUrn  node:// 引用取值（批次 1-6）。公式体混排 `node://id#字段` 时逐 token 调用；
 *                    返回 null/undefined 则保留原文（解析器自然报错，与坏公式同路）。
 *                    值经 SENTINEL 占位符注入（不贴进文本），防注入语义与 A1 引用一致。
 */
export function evalRowFormula(
  formulaBody: string,
  refValues: Record<string, unknown>,
  resolveUrn?: (urn: string) => unknown
): EvalOutcome {
  const body = formulaBody.startsWith('=') ? formulaBody.slice(1) : formulaBody
  try {
    // ① 保护字符串字面量，避免 "A1" 这种文本被当成引用
    const parts: { t: 'code' | 'str'; s: string }[] = []
    STR_LIT.lastIndex = 0
    let last = 0
    let m: RegExpExecArray | null
    while ((m = STR_LIT.exec(body))) {
      if (m.index > last) parts.push({ t: 'code', s: body.slice(last, m.index) })
      parts.push({ t: 'str', s: m[0] })
      last = m.index + m[0].length
    }
    if (last < body.length) parts.push({ t: 'code', s: body.slice(last) })

    // ② 引用 → 占位符（node:// token 先于 A1 替换 —— URN 体内片段不允许被当坐标）
    const values: unknown[] = []
    const withSentinels = parts
      .map((p) => {
        if (p.t === 'str') return p.s
        let s = p.s
        if (resolveUrn) {
          s = s.replace(URN_RE, (urn) => {
            const v = resolveUrn(urn)
            if (v == null) return urn // 取不到值：保留原文 → 解析器自然报错（与坏公式同路）
            const idx = values.length
            values.push(v)
            return `\u0000${idx}\u0001`
          })
        }
        return s.replace(RANGE_RE, (whole, _d1: string, c1s: string, _d2: string, r1s: string, _d3: string, c2s: string, _d4: string, r2s: string) => {
          // #36 区间：矩形展开成二维数组（行优先）占位注入 —— formulajs 的 SUM/AVERAGE/COUNT
          // 等接受嵌套数组参数；区间参与四则（如 A1:B3*2）toNum(数组) 抛错 = Excel #VALUE! 同语义。
          // 归一化端点（A3:B1 → A1:B3）；越出 refValues 的格子取 undefined（formulajs 聚合函数跳过空值）。
          const c1 = colToIndexNum(c1s.toUpperCase())
          const c2 = colToIndexNum(c2s.toUpperCase())
          const r1 = Number(r1s)
          const r2 = Number(r2s)
          const loC = Math.min(c1, c2)
          const hiC = Math.max(c1, c2)
          const loR = Math.min(r1, r2)
          const hiR = Math.max(r1, r2)
          if ((hiC - loC + 1) * (hiR - loR + 1) > RANGE_MAX_CELLS) return whole
          const grid: unknown[][] = []
          for (let r = loR; r <= hiR; r++) {
            const row: unknown[] = []
            for (let c = loC; c <= hiC; c++) {
              const key = `${colNumToIndex(c)}${r}`
              row.push(key in refValues ? refValues[key] : undefined)
            }
            grid.push(row)
          }
          const idx = values.length
          values.push(grid)
          return `\u0000${idx}\u0001`
        }).replace(CELL_RE, (whole, _d1: string, col: string, _d2: string, row: string) => {
          const key = `${col.toUpperCase()}${row}`
          if (!(key in refValues)) return whole
          const idx = values.length
          values.push(refValues[key])
          return `\u0000${idx}\u0001`
        })
      })
      .join('')

    // ③ 解析求值
    const toks = tokenize(withSentinels, values)
    if (!toks.length) return { ok: false, error: '空公式' }
    const out = new Parser(toks).parse()

    if (typeof out === 'number' && !Number.isFinite(out)) {
      return { ok: false, error: '结果不是有限数（可能除零或溢出）' }
    }
    return { ok: true, value: out }
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : String(e) }
  }
}

/** formulajs 可调用函数数量（UI 里展示"支持多少函数"用） */
export const FUNC_COUNT = Object.keys(FUNCS).filter((k) => typeof FUNCS[k] === 'function').length

/** 全部可用函数名（公式自动补全的数据源），字母序 */
export const FUNC_NAMES = Object.keys(FUNCS)
  .filter((k) => typeof FUNCS[k] === 'function')
  .sort()
