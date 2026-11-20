// 工程打包/解包（projectPack）测试
//
// ★ 本文件防的核心退化：**打包不是无损的**。
//   症状链很典型：用户把工程发给同事 → 对方打开发现文档没了/ 连线变了 /
//   面板状态不对 → 但因为「打包成功」提示过，没人会怀疑是打包环节丢的。
//   这个故障在开发机上永远看不见（自己打包自己解，必然成功）。
//
// ★ 三条判定标准，全部在这里验证：
//   ① 无损：pack → unpack 后每个文件的字节**逐字节相等**
//   ② 无关内容不变：pack 不修改源目录
//   ③ 确定性：同样内容两次 pack 得到**同样字节**（否则进 git 就是噪声）

import { describe, it, expect } from 'vitest'
import { packProject, unpackProject, PROJECT_PACK_EXT } from './projectPack'
import { saveProject, nodePath, PROJECT_DIRS, type ProjectFs } from './projectStore'
import { defineTable, defineField } from './schema'
import type { TableSchema } from './schema'

/* ────────────────────── 内存 FS ────────────────────── */

function memFs() {
  const files = new Map<string, string>()
  const bins = new Map<string, Uint8Array>()
  const dirs = new Set<string>()

  const putDirs = (path: string) => {
    const segs = path.split('/')
    for (let i = 1; i < segs.length; i++) dirs.add(segs.slice(0, i).join('/'))
  }

  const fs: ProjectFs = {
    async list(dir) {
      const prefix = dir.endsWith('/') ? dir : `${dir}/`
      const out: string[] = []
      const seen = new Set<string>()
      const push = (name: string) => {
        if (name && !seen.has(name)) {
          seen.add(name)
          out.push(name)
        }
      }
      // ★ 必须同时给文件名与子目录名（对齐 IndexedDB 实现的新契约）。
      //   只给文件名的话打包递归进不去 assets/n1/，
      //   测试会表现为「解包后资产全没了」，但因为整个测试文件都是绿的，
      //   很容易误判成「测试写得对，实现有问题」而找错方向。
      for (const k of [...files.keys(), ...bins.keys()]) {
        if (!k.startsWith(prefix)) continue
        const rest = k.slice(prefix.length)
        if (!rest) continue
        const slash = rest.indexOf('/')
        if (slash === -1) push(decodeURIComponent(rest))
        else push(decodeURIComponent(rest.slice(0, slash)))
      }
      return out
    },
    async readText(path) {
      return files.get(path) ?? null
    },
    async readBytes(path) {
      return bins.get(path) ?? null
    },
    async writeText(path, text) {
      files.set(path, text)
      putDirs(path)
    },
    async writeBytes(path, bytes) {
      bins.set(path, bytes)
      putDirs(path)
    },
    async mkdir(dir) {
      dirs.add(dir)
    },
    async exists(path) {
      return files.has(path) || bins.has(path)
    },
  }
  return { fs, files, bins, dirs }
}

/** 造一个带资产、控件状态、未知字段的工程（覆盖本轮全部新增能力） */
async function seedRich(fs: ProjectFs, root = 'pk') {
  const a: TableSchema = defineTable(
    'a',
    [defineField('名称', 'text'), defineField('单价', 'number')],
    'a',
    'table',
  )
  const b: TableSchema = defineTable(
    'b',
    [defineField('名称', 'text')],
    'b',
    'sort',
    { linkedInputs: ['a'] },
  )
  await saveProject(
    fs,
    { rootDir: root },
    {
      name: '打包测试工程',
      tables: { a, b },
      records: { a: [{ 名称: 'Φ159×6 钢管', 单价: 120 }], b: [] },
      positions: { a: { x: 1, y: 2 }, b: { x: 300, y: 400 } },
      viewState: {
        a: { tab: 'refs', refPanelOpen: true, collapsed: true },
        b: { activeSheet: 'Sheet2' },
      },
      assets: {
        a: [
          {
            id: 'a1',
            relPath: '技术方案.docx',
            name: '技术方案.docx',
            role: 'template',
            mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            size: 5,
            hash: 'abc123',
          },
        ],
      },
    },
  )
  // 放一份二进制资产字节（含 0x00 与高位字节，确保不被当成文本）
  const docBytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x00, 0xff, 0xfe])
  await fs.writeBytes(
    `${root}/${PROJECT_DIRS.assets}/${encodeURIComponent('a')}/技术方案.docx`,
    docBytes,
  )
  return { docBytes }
}

/* ────────────────────── 手工 zip 构造（安全测试用） ────────────────────── */

/**
 * 手写一个 zip 字节，里面放一个指定路径的条目 + 一个合法 project.json。
 *
 * ★ 为什么要手写而不用 JSZip 生成：
 *   JSZip 在 generate 时会把 `../pwned.txt` 规范化成 `pwned.txt`
 *   （实测：entries 变成 ['project.json', '/', 'pwned.txt']），
 *   于是样本里根本没有越界条目，安全测试会「通过」但什么都没验到 ——
 *   这是最典型的假绿：测试绿了，产品有洞。
 *   防护代码不能依赖上游库永远帮你规范化，那天它改了行为就全线失守。
 *
 * 实现：store 模式（method=0，不压缩），无CRC 校验值也能被JSZip 读出内容，
 * 但为了真实，这里**正确填 CRC32**。
 */
function handCraftZipWithPath(path: string, content: string): Uint8Array {
  const enc = new TextEncoder()
  const files = [
    { name: 'project.json', data: enc.encode('{}') },
    { name: path, data: enc.encode(content) },
  ]
  const parts: Uint8Array[] = []
  const central: Uint8Array[] = []
  let offset = 0
  for (const f of files) {
    const nameBytes = enc.encode(f.name)
    const crc = crc32(f.data)
    const local = new Uint8Array(30 + nameBytes.length)
    const dv = new DataView(local.buffer)
    dv.setUint32(0, 0x04034b50, true) // local header
    dv.setUint16(4, 20, true) // version needed
    dv.setUint16(6, 0, true) // flags
    dv.setUint16(8, 0, true) // method = store
    dv.setUint16(10, 0, true) // time（固定）
    dv.setUint16(12, 0x0021, true) // date = 1980-01-01
    dv.setUint32(14, crc, true)
    dv.setUint32(18, f.data.length, true) // compressed size
    dv.setUint32(22, f.data.length, true) // uncompressed size
    dv.setUint16(26, nameBytes.length, true)
    dv.setUint16(28, 0, true) // extra len
    local.set(nameBytes, 30)

    const cd = new Uint8Array(46 + nameBytes.length)
    const cdv = new DataView(cd.buffer)
    cdv.setUint32(0, 0x02014b50, true) // central header
    cdv.setUint16(4, 20, true)
    cdv.setUint16(6, 20, true)
    cdv.setUint16(8, 0, true)
    cdv.setUint16(10, 0, true) // method
    cdv.setUint16(12, 0, true)
    cdv.setUint16(14, 0x0021, true)
    cdv.setUint32(16, crc, true)
    cdv.setUint32(20, f.data.length, true)
    cdv.setUint32(24, f.data.length, true)
    cdv.setUint16(28, nameBytes.length, true)
    cdv.setUint32(42, offset, true) // 本地头偏移
    cd.set(nameBytes, 46)

    parts.push(local, f.data)
    central.push(cd)
    offset += local.length + f.data.length
  }
  const cdSize = central.reduce((n, c) => n + c.length, 0)
  const end = new Uint8Array(22)
  const edv = new DataView(end.buffer)
  edv.setUint32(0, 0x06054b50, true) // EOCD
  edv.setUint16(8, files.length, true)
  edv.setUint16(10, files.length, true)
  edv.setUint32(12, cdSize, true)
  edv.setUint32(16, offset, true)

  const all = [...parts, ...central, end]
  const total = all.reduce((n, c) => n + c.length, 0)
  const out = new Uint8Array(total)
  let p = 0
  for (const c of all) {
    out.set(c, p)
    p += c.length
  }
  return out
}

/** CRC32（zip 规范用） */
function crc32(data: Uint8Array): number {
  let table = crc32.table
  if (!table) {
    table = new Uint32Array(256)
    for (let n = 0; n < 256; n++) {
      let c = n
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
      table[n] = c >>> 0
    }
    crc32.table = table
  }
  let crc = 0xffffffff
  for (let i = 0; i < data.length; i++) {
    crc = table[(crc ^ data[i]) & 0xff] ^ (crc >>> 8)
  }
  return (crc ^ 0xffffffff) >>> 0
}
crc32.table = undefined as Uint32Array | undefined

/* ────────────────────── 打包 ────────────────────── */

describe('packProject · 打包', () => {
  it('★ 无损：pack → unpack 后每个文件字节逐字节相等', async () => {
    const src = memFs()
    const { docBytes } = await seedRich(src.fs)
    const zip = await packProject(src.fs, 'pk')

    const dst = memFs()
    const r = await unpackProject(dst.fs, 'out', zip)
    expect(r.problems).toEqual([])

    // ★ 解包按「是不是合法 UTF-8 文本」分流：文本走 writeText、其余走 writeBytes。
    //   两边都必须逐字节/逐字符相等 ——
    //   只比字符串会让「解包是否破坏二进制」这个关键问题测不出来
    //   （UTF-8 对合法文本本来就无损，损坏只在二进制上暴露）。
    for (const [path, text] of src.files) {
      const rel = path.replace(/^pk\//, '')
      expect(dst.files.get(`out/${rel}`), `文本文件 ${rel} 内容变了`).toBe(text)
    }
    for (const [path, bytes] of src.bins) {
      const rel = path.replace(/^pk\//, '')
      expect(dst.bins.get(`out/${rel}`), `二进制文件 ${rel} 没写出来或变了`).toEqual(bytes)
    }
    // ★ 尤其：二进制资产字节没变（这一条最容易被静默降级破坏）
    expect(
      dst.bins.get(`out/${PROJECT_DIRS.assets}/${encodeURIComponent('a')}/技术方案.docx`),
    ).toEqual(docBytes)
    // ★ 且文本没有误进字节表（否则 loadProject 的 readText 读不到）
    expect(dst.files.has('out/project.json')).toBe(true)
  })

  it('★ 确定性：同一份工程打包两次得到完全相同的字节', async () => {
    // ★ 这是「工程能进 git」的前提。若两次打包字节不同，
    //   每次保存都会凭空产生一条二进制变更，git 完全失去意义。
    //
    // ⚠ 前置条件必须用**同一个 FS** 打两次，不能「造两份内容相同的工程」。
    //   造两份的话，两次 saveProject 的 updatedAt 差了几毫秒，
    //   内容本就不同，字节必然不同 —— 那样测的是 saveProject 的时间戳，
    //   不是 packProject 的确定性。两种失败原因会互相掩盖，排查时极难分辨。
    const m = memFs()
    await seedRich(m.fs)
    const a = await packProject(m.fs, 'pk')
    const b = await packProject(m.fs, 'pk')
    expect(a.length).toBe(b.length)
    expect(Array.from(a)).toEqual(Array.from(b))
  })

  it('pack 不修改源目录（打包是只读操作）', async () => {
    const m = memFs()
    await seedRich(m.fs)
    const before = {
      files: new Map(m.files),
      bins: new Map(m.bins),
      dirs: new Set(m.dirs),
    }
    await packProject(m.fs, 'pk')
    expect(m.files.size).toBe(before.files.size)
    expect(m.bins.size).toBe(before.bins.size)
    expect([...m.dirs].sort()).toEqual([...before.dirs].sort())
    for (const [k, v] of before.files) expect(m.files.get(k)).toBe(v)
  })

  it('中文文件名不乱码（走 UTF-8 标记，不是 GBK）', async () => {
    const m = memFs()
    await seedRich(m.fs)
    const zip = await packProject(m.fs, 'pk')
    const dst = memFs()
    await unpackProject(dst.fs, 'out', zip)
    // 文件名原样还原（decodeURIComponent 在 list 里做过一次）
    expect([...dst.bins.keys()].some((k) => k.includes('技术方案.docx'))).toBe(true)
  })

  it('空目录也在包里（history/ 空目录解包后仍存在）', async () => {
    const m = memFs()
    await saveProject(m.fs, { rootDir: 'pk' }, {
      name: '空',
      tables: {},
      records: {},
      positions: {},
    })
    const zip = await packProject(m.fs, 'pk')
    const dst = memFs()
    await unpackProject(dst.fs, 'out', zip)
    // ★ 四个目录都必须建出来，否则打开工程报「缺目录」而工程本身是好的
    for (const d of Object.values(PROJECT_DIRS)) {
      expect(dst.dirs.has(`out/${d}`), `解包后缺目录 ${d}`).toBe(true)
    }
  })

  it('清单之外的文件也会被带走（孤儿节点文件 / 用户手放的数据）', async () => {
    const m = memFs()
    await seedRich(m.fs)
    // 手工塞一个不在清单里的文件
    await m.fs.writeText(`${'pk'}/${PROJECT_DIRS.data}/外部数据/价目.json`, '{"x":1}')
    const zip = await packProject(m.fs, 'pk')
    const dst = memFs()
    await unpackProject(dst.fs, 'out', zip)
    // ★ 漏掉它等于丢数据 —— 打包的语义是「把工程目录整个带走」
    expect(dst.files.get(`out/${PROJECT_DIRS.data}/外部数据/价目.json`)).toBe('{"x":1}')
  })

  it('打包后解包，loadProject 能正常打开且连线还在', async () => {
    const src = memFs()
    await seedRich(src.fs)
    const zip = await packProject(src.fs, 'pk')
    const dst = memFs()
    await unpackProject(dst.fs, 'out', zip)

    const { loadProject } = await import('./projectStore')
    const loaded = await loadProject(dst.fs, { rootDir: 'out' })
    expect(loaded.problems.filter((p) => p.level === 'error')).toEqual([])
    expect(Object.keys(loaded.tables).sort()).toEqual(['a', 'b'])
    // ★ 连线在：flow:a->b。这是「打包没毁掉派生数据」的端到端证据
    expect(loaded.edges?.edges.map((e) => e.id)).toEqual(['flow:a->b'])
    // 控件状态与资产登记
    expect(loaded.nodeFiles.a.viewState).toEqual({
      tab: 'refs',
      refPanelOpen: true,
      collapsed: true,
    })
    expect(loaded.nodeFiles.a.assets?.[0].name).toBe('技术方案.docx')
  })

  it('扩展名常量带点（拼接文件名时不会拼出 "工程.wfcwfc"）', () => {
    // 2026-10-06 用户定名 .wfc（内部仍是 .cbproj 目录的 zip，见 PROJECT_PACK_EXT 注释）
    expect(PROJECT_PACK_EXT).toBe('.wfc')
    expect(`工程${PROJECT_PACK_EXT}`).toBe('工程.wfc')
  })
})

/* ────────────────────── 解包 ────────────────────── */

describe('unpackProject · 解包', () => {
  it('不是工程包（无 project.json）→ 明确报出来', async () => {
    const src = memFs()
    await src.fs.writeText('random/a.txt', 'hello')
    const zip = await packProject(src.fs, 'random')
    const dst = memFs()
    const r = await unpackProject(dst.fs, 'out', zip)
    // ★ 必须明确说「这不是工程包」，不能静默开出一个空工程 ——
    //   用户会以为是自己的工程丢了。
    expect(r.problems.some((p) => p.includes('不是一个工程包'))).toBe(true)
  })

  it('坏 zip（不是 zip 字节）→ 报解析失败，不抛异常', async () => {
    const dst = memFs()
    const r = await unpackProject(dst.fs, 'out', new Uint8Array([1, 2, 3, 4, 5]))
    expect(r.problems.some((p) => p.includes('解析失败'))).toBe(true)
    expect(r.written).toEqual([])
  })

  it('★ 越界路径不会写到工程目录外（双层防护）', async () => {
    // ★ 这是安全项：工程 zip 常来自「同事发来的文件」，
    //   含越界条目的 zip 若原样写就能覆盖工程目录外的文件。
    //
    // ⚠ 造样本必须**手写 zip 字节**，不能用 JSZip 生成：
    //   JSZip 在 generate 时就会把 `../pwned.txt` 规范化成 `pwned.txt`，
    //   样本里根本没有越界条目，测试会「通过」但什么都没验到 —— 典型假绿。
    //
    // ⚠ 实测确认：**JSZip 在 load 时也会规范化 `../`**：
    //   手写的 `../pwned.txt` 与 `a/../../b.txt` 读回来都变成了
    //   `pwned.txt` / `b.txt`。所以 `../` 这类由上游库挡，
    //   我们的 isSafeZipPath 根本收不到 —— 这是**正常的双层防护**，
    //   不是防护失效。
    //   但绝不能因此删掉 isSafeZipPath：
    //   上游库哪天改了行为（或换成别的实现），那一层就没了。
    //   本层要挡住的是它**不挡**的形态：绝对路径与盘符路径。
    for (const path of ['../pwned.txt', 'a/../../b.txt', '/etc/passwd', 'C:/Windows/x.dll']) {
      const bytes = handCraftZipWithPath(path, 'gotcha')
      const dst = memFs()
      const r = await unpackProject(dst.fs, 'out', bytes)
      const all = [...dst.files.keys(), ...dst.bins.keys()]
      // ★ 唯一的安全属性：**不能写到 out/ 之外**。
      //   注意不要断言「内容没写出来」——
      //   `../pwned.txt` 被 JSZip 规范化成 `pwned.txt` 后落在 out/ 内，
      //   这完全安全（只是文件名被上游改了），断言它没写出来是**过严且错误**的：
      //   会把一个正确的防护判成失败，逼着后人去掉断言或改防护，
      //   反而降低安全性。这类「为了测试而削弱防护」是最坏的一种修正。
      const outside = all.filter((k) => !k.startsWith('out/'))
      expect(outside, `路径 ${path} 写到了工程目录外：${outside.join(', ')}`).toEqual([])
      // ★ 只对**我们这层拒掉的**形态要求报告。
      //   `../` 那种被 JSZip 规范化后安全落在目录内的，不该报问题 ——
      //   报「跳过了 pwned.txt」会让用户以为自己的文件丢了，而它明明在包里。
      //   这类「过度报告」和静默失败一样有害：都是让用户对提示失去信任。
      if (path.startsWith('/') || /^[A-Za-z]:/.test(path)) {
        expect(
          r.problems.some((p) => p.includes('非法路径')),
          `${path} 被本层拒掉却没有报告`,
        ).toBe(true)
      }
    }
  })

  it('★ 绝对路径与盘符路径由我们这层挡住（JSZip 不挡这两种）', async () => {
    // 这两种 JSZip 会原样透传，所以能真正验到 isSafeZipPath
    for (const path of ['/etc/passwd', 'C:/Windows/system32/x.dll']) {
      const dst = memFs()
      const r = await unpackProject(dst.fs, 'out', handCraftZipWithPath(path, 'x'))
      expect(r.problems.some((p) => p.includes('非法路径')), `${path} 没被本层拦住`).toBe(true)
    }
  })

  it('Windows 反斜杠路径被归一（不造出嵌套目录名）', async () => {
    const src = memFs()
    await seedRich(src.fs)
    const JSZip = (await import('jszip')).default
    const z = await JSZip.loadAsync(await packProject(src.fs, 'pk'))
    z.file('nodes\\子目录\\a.json', '{}')
    const zip = await z.generateAsync({ type: 'uint8array' })

    const dst = memFs()
    await unpackProject(dst.fs, 'out', zip)
    // ★ 归一后应是 out/nodes/子目录/a.json，
    //   而不是 out/nodes\子目录\a.json 这种带反斜杠的单段名。
    expect(dst.files.has('out/nodes/子目录/a.json')).toBe(true)
    expect([...dst.files.keys()].some((k) => k.includes('\\'))).toBe(false)
  })

  it('存储不支持二进制写入时拒绝解包（绝不静默降级损坏文件）', async () => {
    const src = memFs()
    await seedRich(src.fs)
    const zip = await packProject(src.fs, 'pk')

    // 造一个只有文本接口的 FS
    const texts = new Map<string, string>()
    const dirs = new Set<string>()
    const weakFs: ProjectFs = {
      async list() {
        return []
      },
      async readText(p) {
        return texts.get(p) ?? null
      },
      async writeText(p, t) {
        texts.set(p, t)
        const segs = p.split('/')
        for (let i = 1; i < segs.length; i++) dirs.add(segs.slice(0, i).join('/'))
      },
      async mkdir(d) {
        dirs.add(d)
      },
    }
    const r = await unpackProject(weakFs, 'out', zip)
    // ★ 必须明确报「需要 writeBytes」，而不是写一通损坏的文本文件
    expect(r.problems.some((p) => p.includes('writeBytes'))).toBe(true)
    expect(texts.size).toBe(0)
  })

  it('空工程也能打包能解包', async () => {
    const src = memFs()
    await saveProject(src.fs, { rootDir: 'pk' }, {
      name: '空',
      tables: {},
      records: {},
      positions: {},
    })
    const zip = await packProject(src.fs, 'pk')
    const dst = memFs()
    const r = await unpackProject(dst.fs, 'out', zip)
    expect(r.written.length).toBeGreaterThan(0) // 至少有 project.json
    const { loadProject } = await import('./projectStore')
    const loaded = await loadProject(dst.fs, { rootDir: 'out' })
    expect(loaded.manifest.nodeIds).toEqual([])
  })

  it('★ 完整链路：存 → 打包 → 解包 → 再打包，字节稳定（工程可反复交付）', async () => {
    const src = memFs()
    await seedRich(src.fs)
    const zip1 = await packProject(src.fs, 'pk')

    const mid = memFs()
    await unpackProject(mid.fs, 'pk', zip1)
    const zip2 = await packProject(mid.fs, 'pk')

    // ★ 解包再打包必须得到同样字节：说明中间过程真的无损，
    //   且不会因为「读过一遍」而改变内容（编码、排序、时间戳都没被污染）
    expect(Array.from(zip2)).toEqual(Array.from(zip1))
  })
})