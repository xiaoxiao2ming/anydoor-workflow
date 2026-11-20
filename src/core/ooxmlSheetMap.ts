/**
 * OOXML 工作表名 → 部件文件名 映射
 *
 * 立档：2026-10-29· 修「写错表」的硬错误
 *
 * ★ 为什么必须查 rels、不能按序号猜（2026-10-29 实测两个真实样本）：
 *   样本 A（单表）  规格表      sheetId=1  r:id="rId3"  → xl/worksheets/sheet1.xml
 *   样本 B（双表）  汇总表 rId1 → sheet1.xml
 *                   工程量表 rId2 → sheet2.xml
 *
 *   ★ 样本 A 证明：**rId 与 sheetN 的编号毫无关系**（rId3 对应 sheet1）。
 *     样本 B 只是**恰好**顺序一致，容易让人误以为「按序号排就行」。
 *   ⇒ 真实解析链：workbook.xml 的 <sheet name r:id> → workbook.xml.rels 的
 *     <Relationship Id="rIdN" Target="worksheets/sheetN.xml">
 *     两跳都必须走；任何一跳猜错都会**写到错误的表**（用户看到「值写进别的表了」）。
 *
 * 纯计算层：字节进 → 映射出，零 IO。
 */

import JSZip from 'jszip'

/** 表名 → 部件相对路径（如 'xl/worksheets/sheet1.xml'） */
export type SheetPartMap = Record<string, string>

/**
 * ★ 从 xlsx 字节解析「表名 → 部件路径」映射。
 *
 * @param bytes xlsx 字节
 * @returns 映射表（表名 → 'xl/worksheets/sheetN.xml'）；
 *          解析不到时返回空对象（**不猜、不兜底成按序号**）——
 *          上层据此如实报错，好过默默写错表。
 */
export async function buildSheetPartMap (bytes: Uint8Array): Promise<SheetPartMap> {
  const out: SheetPartMap = {}
  let zip
  try {
    zip = await JSZip.loadAsync(bytes)
  } catch {
    return out // 不是 zip（或损坏）⇒ 交上层报错
  }
  const wbFile = zip.file('xl/workbook.xml')
  const relFile = zip.file('xl/_rels/workbook.xml.rels')
  //★ 两跳都必需：只有 workbook.xml 拿不到文件名，只有 rels 拿不到表名
  if (!wbFile || !relFile) return out

  let wbXml = ''
  let relXml = ''
  try {
    wbXml = await wbFile.async('string')
    relXml = await relFile.async('string')
  } catch {
    return out
  }

  // ① rels：rId → Target（worksheets/xxx.xml）
  const byRid: Record<string, string> = {}
  const relRe = /<Relationship\b[^>]*>/g
  let m: RegExpExecArray | null
  while ((m = relRe.exec(relXml))) {
    const tag = m[0]
    const id = /\bId\s*=\s*"([^"]+)"/.exec(tag)?.[1]
    const target = /\bTarget\s*=\s*"([^"]+)"/.exec(tag)?.[1]
    // 只要 worksheet 关系（工作簿还挂在 styles/theme 等其它关系上）
    if (!id || !target) continue
    if (!/relationships\/worksheet/.test(tag)) continue
    // Target 可能是 'worksheets/sheet1.xml' 或绝对路径 '/xl/worksheets/sheet1.xml'
    const norm = target.replace(/^\/?xl\//, '').replace(/^\//, '')
    byRid[id] = 'xl/' + norm.replace(/^\.\//, '')
  }

  // ② workbook.xml：表名 → rId
  const sheetRe = /<sheet\b[^>]*>/g
  while ((m = sheetRe.exec(wbXml))) {
    const tag = m[0]
    const name = /\bname\s*=\s*"([^"]*)"/.exec(tag)?.[1]
    //r:id 带命名空间前缀，宽松匹配 r:id
    const rid = /\br:id\s*=\s*"([^"]+)"/.exec(tag)?.[1]
    if (!name || !rid) continue
    const part = byRid[rid]
    if (part) out[unescapeXmlAttr(name)] = part
  }
  return out
}

/** XML 属性里的实体反转义（表名可能含 &amp; 等） */
function unescapeXmlAttr (s: string): string {
  return String(s)
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/**
 * ★ 按映射写单元格（严格：找不到映射就**不写**，不猜表）。
 *
 * @param sheetPartMap 表名 → 部件路径
 * @param sheetName    目标表名
 * @param a1           单元格地址
 * @param value        值（字符串）
 * @param readPart     从 zip 读某部件内容（由调用方提供，便于测试）
 * @param writePart    写回某部件内容
 * @returns { hit, why? }
 */
export function resolveSheetWrite (
  sheetPartMap: SheetPartMap,
  sheetName: string,
): { part: string } | { part: null; why: string } {
  if (!sheetName) return { part: null, why: '落点未给工作表名' }
  const part = sheetPartMap[sheetName]
  if (!part) {
    const known = Object.keys(sheetPartMap).join('、') || '（解析不到）'
    return { part: null, why: `文件里没有工作表「${sheetName}」（现有：${known}）` }
  }
  return { part }
}
