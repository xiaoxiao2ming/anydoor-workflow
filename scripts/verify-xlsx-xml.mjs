/** 直接看产物 XML 里 D2:D4（合价列）的原文 —— 验证 formula + cachedValue 双写 */
import fs from 'node:fs'
import zlib from 'node:zlib'

const DIR = new URL('../sample-template/', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')
const buf = fs.readFileSync(DIR + 'excel-e2e-out.xlsx')

function unzip(b) {
  const out = {}
  let i = b.length - 22
  while (i > 0 && b.readUInt32LE(i) !== 0x06054b50) i--
  const cdOff = b.readUInt32LE(i + 16)
  const cdCount = b.readUInt16LE(i + 10)
  let p = cdOff
  for (let k = 0; k < cdCount; k++) {
    if (b.readUInt32LE(p) !== 0x02014b50) break
    const method = b.readUInt16LE(p + 10)
    const csize = b.readUInt32LE(p + 20)
    const nameLen = b.readUInt16LE(p + 28)
    const extraLen = b.readUInt16LE(p + 30)
    const cmtLen = b.readUInt16LE(p + 32)
    const lho = b.readUInt32LE(p + 42)
    const name = b.toString('utf8', p + 46, p + 46 + nameLen)
    const lnameLen = b.readUInt16LE(lho + 26)
    const lextraLen = b.readUInt16LE(lho + 28)
    const dataStart = lho + 30 + lnameLen + lextraLen
    const raw = b.slice(dataStart, dataStart + csize)
    out[name] = method === 0 ? raw : zlib.inflateRawSync(raw)
    p += 46 + nameLen + extraLen + cmtLen
  }
  return out
}

const z = unzip(buf)
const s2 = z['xl/worksheets/sheet2.xml'].toString('utf8')

console.log('=== sheet2.xml（工程量表）D 列单元格原文 ===')
// 提取每个 <c r="Dn" ...>...</c>
for (const m of s2.matchAll(/<c r="(D\d+)"[^>]*>[\s\S]*?<\/c>/g)) {
  console.log('  ' + m[0])
}

console.log('\n=== 汇总表 sheet1.xml B1/B2 ===')
const s1 = z['xl/worksheets/sheet1.xml'].toString('utf8')
for (const m of s1.matchAll(/<c r="(B[12])"[^>]*>[\s\S]*?<\/c>/g)) {
  console.log('  ' + m[0])
}

console.log('\n=== 共享字符串表（前 12 条）===')
const ss = z['xl/sharedStrings.xml']?.toString('utf8') ?? '（无）'
const items = [...ss.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => m[1].replace(/<[^>]+>/g, ''))
items.slice(0, 12).forEach((t, i) => console.log(`  [${i}] ${t}`))

console.log('\n=== table 定义 ===')
const tParts = Object.keys(z).filter((n) => /tables\//.test(n))
for (const n of tParts) {
  console.log('  ' + n)
  if (n.endsWith('.xml')) console.log('    ' + z[n].toString('utf8').slice(0, 400))
}

// 断言
const d = {}
for (const m of s2.matchAll(/<c r="(D\d+)"[^>]*>([\s\S]*?)<\/c>/g)) {
  const inner = m[2]
  const f = /<f>([\s\S]*?)<\/f>/.exec(inner)?.[1] ?? null
  const v = /<v>([\s\S]*?)<\/v>/.exec(inner)?.[1] ?? null
  d[m[1]] = { f, v }
}
console.log('\n=== 断言 ===')
const expect = { D2: '2550', D3: '728', D4: '3120' }
let ok = true
for (const [k, e] of Object.entries(expect)) {
  const c = d[k] ?? { f: null, v: null }
  const good = !!c.f && c.v === e
  if (!good) ok = false
  console.log(`  ${k}: <f>=${JSON.stringify(c.f)}  <v>=${JSON.stringify(c.v)}  期望 ${e}  ${good ? '✅' : '❌'}`)
}
console.log('\n' + (ok ? '✅ 公式 + 缓存值双写全部正确（Excel 能重算，离线也能看）' : '❌ 双写有问题'))
