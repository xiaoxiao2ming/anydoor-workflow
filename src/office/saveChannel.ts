/**
 * 保存通道裁决（纯函数，2026-10-17）
 *
 * 背景 —— 用户报「还是不行」的真凶（oo-events.log 18:59:52 实锤）：
 *   两条通道同时存在，主通道（OO 自身 / DS forcesave）先返回，宿主只要看到「有字节」
 *   就认定成功，于是把提取通道刚读到的正确内容整个覆盖。
 *   那份字节是 6258B、非空单元格 = 0 的空表 —— 用户重开当然是空白。
 *
 * 为什么主通道必然是空壳：
 *   Editor.bin 是本机 x2t 在「打开时」从源文件转出来的（oo-core.mjs 的 dsEditorBin 门禁），
 *   它从诞生那一刻起就不包含任何编辑内容。forcesave 把它转回来，产物只能等于「打开时的原文件」；
 *   源是空白新建时，就是一张空表。
 *
 * 因此主通道只有在「产物确实带了内容、且不等于打开时的原文件」时才可信。
 * 这里把这个判据抽成纯函数，便于单测回归（历史上这条链反复回退）。
 */

/** 服务端 __ooDigest 回传的内容摘要 */
export interface OoDigest {
  kind?: string
  cells?: number
  chars?: number
  parts?: number
  sample?: string
  text?: string
}

/** saveBack 的返回：产物字节 + 内容摘要 + 是否与打开时的原始字节完全相同 */
export interface SaveCandidate {
  bytes: Uint8Array
  digest: OoDigest | null
  sameAsSource: boolean
  /** 打开时源文件的内容摘要（字节指纹不可靠，内容摘要才是真判据） */
  srcDigest?: OoDigest | null
}

/** 两份内容摘要是否描述同一份内容。
 *  ★ 2026-10-17 第二次保存丢数据的真因：x2t 往返会把极简 xlsx 重新排版，
 *    3003B → 6285B，字节指纹 `sameAsSource` 因此永远为 no；
 *    但内容其实一模一样（都只有 A1=1）⇒ 主通道被判为「有新内容」而采用，
 *    把提取到的 A2=2 整个覆盖。必须改用内容维度比对。 */
export function digestSame(a: OoDigest | null | undefined, b: OoDigest | null | undefined): boolean {
  if (!a || !b) return false
  if (!a.kind || a.kind !== b.kind) return false
  if (a.kind === 'xlsx') return (a.cells || 0) === (b.cells || 0) && String(a.sample || '') === String(b.sample || '')
  if (a.kind === 'docx') return (a.chars || 0) === (b.chars || 0) && String(a.sample || '') === String(b.sample || '')
  return (a.parts || 0) === (b.parts || 0)
}

/** 提取到的值数组里到底有没有真内容。
 *  形态是 [[{v:123},{v:'abc',f:'=A1*2'},{}]]：空对象 = 该格没内容（按 used range 铺满产生）。 */
export function rowsHaveContent(rows: unknown[][] | undefined | null): boolean {
  if (!rows || !rows.length) return false
  for (const r of rows) {
    if (!Array.isArray(r)) continue
    for (const c of r) {
      if (c == null || c === '') continue
      if (typeof c === 'object') {
        const o = c as { v?: unknown; t?: unknown; f?: unknown }
        if (o.f != null && String(o.f) !== '') return true
        if (o.v != null && String(o.v) !== '') return true
        if (o.t != null && String(o.t) !== '') return true
        continue
      }
      if (String(c) !== '') return true
    }
  }
  return false
}

/** 产物本身是否为空（按 OOXML 摘要判断；拿不到摘要时按「不空」放行） */
export function digestIsEmpty(d: OoDigest | null | undefined): boolean {
  if (!d) return false
  if (d.kind === 'xlsx') return (d.cells || 0) === 0
  if (d.kind === 'docx') return (d.chars || 0) === 0
  return false
}

/** 提取内容形状统计——比「行数」更能说明问题：
 *  行数 2 但非空格 0 = 读到了 used range 却取不到值；公式数 > 0 说明关联地址在。 */
export function rowStats(rows: unknown[][] | undefined | null): { rows: number; cols: number; vals: number; formulas: number; firstRow: string } {
  const st = { rows: 0, cols: 0, vals: 0, formulas: 0, firstRow: '' }
  if (!rows || !rows.length) return st
  st.rows = rows.length
  for (const r of rows) {
    if (!Array.isArray(r)) continue
    if (r.length > st.cols) st.cols = r.length
    for (const c of r) {
      if (c == null || c === '') continue
      if (typeof c === 'object') {
        const o = c as { v?: unknown; t?: unknown; f?: unknown }
        if (o.f != null && String(o.f) !== '') st.formulas++
        if ((o.f != null && String(o.f) !== '') || (o.v != null && String(o.v) !== '') || (o.t != null && String(o.t) !== '')) st.vals++
      } else if (String(c) !== '') st.vals++
    }
  }
  try { st.firstRow = JSON.stringify(rows[0] || []).slice(0, 120) } catch { st.firstRow = '(不可序列化)' }
  return st
}

export interface ChannelDecision {
  /** 是否采用主通道产物 */
  use: boolean
  /** 拒绝原因（空串 = 通过）；用于日志与诊断抽屉 */
  reject: string
}

/**
 * 裁决主通道产物能不能用。
 * @param fb      saveBack 结果（null = 主通道没出产物）
 * @param rows    提取通道拿到的值数组
 * @param dirty   ★ 2026-10-29 升级为**硬门**：文档有未保存编辑时，forcesave 在本部署
 *                （自研 socket 无 DS 会话）架构性 400 ⇒ 产物数学上不可能含这次编辑——
 *                无论内容摘要怎么变（符号字体会造出 804≠789 的假阳性，21:38 实测）。
 *                dirty=true 一律不放行，交给样式化重建/沿用守卫。
 */
export function decideMainChannel(
  fb: SaveCandidate | null | undefined,
  rows: unknown[][] | undefined | null,
  dirty: boolean,
): ChannelDecision {
  if (!fb || !fb.bytes || !fb.bytes.length) return { use: false, reject: '' }
  const extractReal = rowsHaveContent(rows)
  if (!extractReal) return { use: true, reject: '' }
  /* ⓪ ★ 2026-10-29 dirty 硬门（必须排在一切内容比对之前——符号规范化会让
   *    摘要「恰好不同」骗过 ③，把不含编辑的产物当成功采用，自动重开就把编辑闪没了） */
  if (dirty) {
    return { use: false, reject: '文档有未保存编辑，forcesave 无会话（架构性 400）⇒ 产物不可能含这次编辑' }
  }
  /* ① 产物是空表：源是空白新建 → 永远命中这条 */
  if (digestIsEmpty(fb.digest)) {
    return { use: false, reject: '产物是空表（' + (fb.digest && fb.digest.text ? fb.digest.text : '无摘要') + '）' }
  }
  /* ② 产物与打开时的原始字节逐字节相同 ⇒ 这次编辑根本没进产物。
   *   只在「确实改过」时才拒绝：用户没改就点保存时，用主通道能保住原样式。 */
  if (fb.sameAsSource && dirty) {
    return { use: false, reject: '产物与打开时字节完全相同（没吃到这次编辑）' }
  }
  /* ③ ★ 2026-10-17 第二次/第三次保存丢数据的真凶：
   *   产物「内容」与打开时源文件完全一致 ⇒ 它只是打开时那份的重新排版，
   *   这次编辑一个字都没进去。必须按内容比，不能按字节比（x2t 往返会改字节）。 */
  if (fb.srcDigest && digestSame(fb.digest, fb.srcDigest)) {
    return {
      use: false,
      reject: '产物内容与打开时完全相同（' + (fb.srcDigest.text || '摘要一致') + '）⇒ 这次编辑没进产物',
    }
  }
  /* ④ 产物里的非空格数少于提取通道读到的 ⇒ 产物丢格子。
   *   第二次保存实测：提取 2 格（1、2），产物只有 1 格（1）—— 采用就是丢 A2。 */
  if (fb.digest && fb.digest.kind === 'xlsx') {
    const st = rowStats(rows)
    if (st.vals > (fb.digest.cells || 0)) {
      return {
        use: false,
        reject: '产物只有 ' + (fb.digest.cells || 0) + ' 格，提取到 ' + st.vals + ' 格内容（产物丢格）',
      }
    }
  }
  return { use: true, reject: '' }
}
