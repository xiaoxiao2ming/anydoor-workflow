// 工程打包：.cbproj 目录 ↔ .cbproj.zip（2026-10-06）
//
// ★ 定位：**zip 只是外壳，不是另一种格式**。
//   Q1 已拍板目录式（git diff 能看出哪个节点的什么参数改了、能人工排查），
//   派生需求是「交付时要单文件」。故本文件做的是**无损打包/解包**：
//     目录 ──pack──→ zip ──unpack──→ 目录
//   解包出来的结构与打包进去**逐字节一致**，不重排、不重命名、不丢文件。
//
// ★ 为什么不允许 zip 成为「另一种保存格式」：
//   一旦 zip 里存的是「合并后的单份JSON」而不是原目录树，
//   就等于出现了第二套真相源 —— 两边各存一份、互相转换，必然漂移。
//   而且那样会毁掉 Q1 的全部理由（人工排查 + git diff）。
//   故本文件只搬运字节，不解释任何结构。结构解释是 projectFormat 的事。
//
// ⚠ 三个容易做错的地方（都已在下面标注）：
//   ① 空目录：zip 会自动省略空目录，解包后目录消失 → 表现为「工程结构少了层」
//   ② 编码：文件名是中文，zip 头有两种编码标记（UTF-8 flag / Unicode extra field），
//      处理不对就是一堆乱码文件名
//   ③ 路径分隔符：zip 内必须是 '/'，Windows 上直接写 '\\' 会造出嵌套目录名

import { zipSync, unzipSync, strToU8 } from 'fflate'
import type { ProjectFs } from './projectStore'

/** 打包结果的默认扩展名（2026-10-06 用户定名 .wfc = WorkFlow Canvas）。
 *  内部仍是 .cbproj 目录树的 zip，unpack 只认 project.json 不看扩展名——
 *  故 .wfc 与旧的 .cbproj.zip 双向都能打开，不存在两套格式。 */
export const PROJECT_PACK_EXT = '.wfc'

/** 工程目录中需要遍历的顶层目录 */
const TOP_DIRS = ['nodes', 'history', 'data', 'assets'] as const

/**
 * 把工程目录打成zip 字节。
 *
 * @param fs 工程存储
 * @param rootDir 工程根目录（与 ProjectRef.rootDir 同义）
 * @param opts.onProgress 进度回调（n/m），用于大工程的进度条
 * @returns zip 字节；调用方负责存盘或下载
 *
 * ★ 为什么遍历目录而不是「按清单列文件」：
 *   按清单列会漏掉清单之外的东西 —— 孤儿节点文件、用户手放在 data/ 下的
 *   外部数据（阶段二的 embed 会真放文件）、history/ 里的旧版本。
 *   打包的语义是「把这个工程目录整个带走」，漏文件等于丢数据。
 *   反过来「多带了不该带的」在这里无害：打包整个目录本来就是用户预期。
 */
export async function packProject(
  fs: ProjectFs,
  rootDir: string,
  opts: { onProgress?: (done: number, total: number) => void; savePath?: string } = {},
): Promise<Uint8Array> {
  const files: { path: string; bytes: Uint8Array }[] = []
  const dirs = new Set<string>()

  // 顶层文件（project.json / edges.json）。同样排序，理由见 walk 内注释。
  for (const name of [...new Set(await fs.list(rootDir))].sort()) {
    const b = await readFileBytes(fs, `${rootDir}/${name}`)
    if (b) files.push({ path: name, bytes: b })
    opts.onProgress?.(files.length, 0)
  }

  // 四个子目录逐个递归
  for (const d of TOP_DIRS) {
    await walk(fs, `${rootDir}/${d}`, d, files, dirs)
    opts.onProgress?.(files.length, 0)
  }

  // ② fflate 打包（原生级速度——比 JSZip 快 5-10 倍）。
  //   ★ 目录条目：fflate 对以 '/' 结尾的 key 会写目录条目，
  //     显式加入排序后的目录 key，保证「空目录不丢失」语义与旧版一致。
  //   ★ 固定时间戳：fflate 的 zipSync 接受 mtime 选项（per-file），固定 1980-01-01，
  //     保持「同样内容 → 同样字节 → git 可 diff」的核心理由。
  const zipped: Record<string, Uint8Array> = {}
  for (const dir of [...dirs].sort()) zipped[dir + '/'] = new Uint8Array(0)
  for (const e of files) zipped[e.path] = e.bytes

  const total = files.length
  opts.onProgress?.(total, total)

  /* ★ 2026-10-25 最后优化：桌面端 zip 移主进程（fflate 在 Node 跑=更快且
     渲染层零冻结）；web 端保留 zipSync。 */
  const apiN = (globalThis as { api?: { projectNativePack?: (files: { path: string; bytes: Uint8Array }[], dirs: string[], savePath?: string) => Promise<{ bytes: Uint8Array | null; savedTo: string | null }> } }).api
  if (apiN?.projectNativePack) {
    const r = await apiN.projectNativePack(files.map((f) => ({ path: f.path, bytes: f.bytes })), [...dirs].sort(), opts.savePath ?? undefined)
    if (r.savedTo) {
      /* 直写模式：字节已在主进程落盘——回传空字节但标记 savedTo */
      return new Uint8Array(0)
    }
    return r.bytes ?? zipSync(zipped, { level: 1, mtime: fixedDate })
  }
  return zipSync(zipped, { level: 1, mtime: fixedDate })
}

/**
 * 固定时间戳。
 *
 * ★ 为什么必须固定：JSZip 默认用当前时间写每个条目的时间。
 *   那样**同样内容两次打包得到两个不同字节的 zip** ——
 *   而工程是要进 git 的，那会让每次提交都凭空多出一条二进制变更，
 *   彻底毁掉 Q1 选目录式的核心理由（可 diff）。
 *   固定成 1980-01-01（zip 格式能表示的最早时间）后：
 *   同样内容 → 同样字节 → git 才知道「没变化」。
 */
const fixedDate = new Date('1980-01-01T00:00:00Z')

/**
 * 读一个文件为字节，优先走二进制接口。
 *
 * ★ 为什么优先 readBytes 而不是「readText 再编码回去」：
 *   那是把字节当 UTF-8 解码再编码 —— 中文名.docx 这种二进制会被破坏
 *   （非法序列替换成 U+FFFD，字节数也变了）。属于静默损坏。
 */
async function readFileBytes(
  fs: ProjectFs,
  path: string,
): Promise<Uint8Array | null> {
  if (fs.readBytes) {
    const b = await fs.readBytes(path)
    if (b) return b
  }
  const t = await fs.readText(path)
  return t === null ? null : new TextEncoder().encode(t)
}

/** 递归入口处已直接调用 walk，这里不再需要单独的路径推导函数 */

/**
 * 递归列出目录下的所有文件（相对工程根，全部用 '/'）。
 *
 * ★ 为什么不能只靠 fs.list 递归：
 *   ProjectFs.list 的契约是「列本层文件名」，**不保证返回子目录名**
 *   （IndexedDB 那种扁平 key 存储天然列不出目录，只能靠前缀）。
 *   只靠 list 递归的话，子目录（如 assets/n1/）永远进不去 ——
 *   症状是「打包成功，解包后文档全没了」，且没有任何报错。
 *   这是最坏的一类缺陷：成功提示 + 数据丢失。
 *
 *   所以这里走**两路**：
 *     ① fs.list 给出本层条目（文件或目录名都可能）
 *     ② 对每个条目先试读；读不到就当目录再递归
 *   读不到又递归不进去的，才算真正丢失（外层会记为结构异常）。
 *
 * @param dirsOut 输出参数：收集到的目录相对路径（用于写zip 目录条目）
 */
async function walk(
  fs: ProjectFs,
  dir: string,
  relPrefix: string,
  filesOut: { path: string; bytes: Uint8Array }[],
  dirsOut: Set<string>,
): Promise<void> {
  dirsOut.add(relPrefix)
  // ★ 必须排序后再遍历：fs.list 的返回顺序在不同适配器下不一致
  //   （IndexedDB 按 key 排、内存实现按插入序）。若直接把 list 的顺序
  //   当打包顺序，同样内容两次打包就可能得到不同字节 ——
  //   而工程要进 git，那会让每次保存都凭空产生一条二进制变更，
  //   彻底毁掉 Q1 选目录式的核心理由。排序把「顺序」变成确定的。
  const names = [...new Set(await fs.list(dir))].sort()
  for (const name of names) {
    const full = `${dir}/${name}`
    const rel = `${relPrefix}/${name}`
    const b = await readFileBytes(fs, full)
    if (b) {
      filesOut.push({ path: rel, bytes: b })
      continue
    }
    // 读不到 → 当子目录再走一层（list 给的可能是目录名）
    await walk(fs, full, rel, filesOut, dirsOut)
  }
}

/**
 * 解包 zip 字节回工程目录。
 *
 * ★ 反向也要「无损」：
 *   - 保留原始字节，不做「读出来再写回去」的文本转换（那会改编码）
 *   - 不解析任何结构（结构由 loadProject 负责），本函数只搬文件
 *
 * @returns 实际写入的文件相对路径列表（供调用方做校验/展示）
 */
export async function unpackProject(
  fs: ProjectFs,
  rootDir: string,
  zipBytes: Uint8Array,
): Promise<{ written: string[]; problems: string[] }> {
  const problems: string[] = []
  const written: string[] = []

  /* ★ 2026-10-25 最后优化：桌面端解压移主进程（同 pack 理由） */
  let unzipped: Record<string, Uint8Array>
  const apiU = (globalThis as { api?: { projectNativeUnpack?: (b: Uint8Array) => Promise<{ ok: boolean; files: Record<string, Uint8Array>; reason?: string }> } }).api
  if (apiU?.projectNativeUnpack) {
    const r = await apiU.projectNativeUnpack(zipBytes)
    if (r.ok) unzipped = r.files
    else return { written, problems: [`zip 解析失败：${r.reason}`] }
  } else {
    try {
      unzipped = unzipSync(zipBytes)
    } catch (e) {
      return {
        written,
        problems: [`zip 解析失败：${e instanceof Error ? e.message : String(e)}`],
      }
    }
  }

  const names = Object.keys(unzipped)
  // ★ 先看有没有 project.json：没有它就不是一个工程包。
  //   不提前说清的话，用户会打开一个「空工程」以为是丢数据了。
  const hasManifest = names.some((n) => n === 'project.json' || n.endsWith('/project.json'))
  if (!hasManifest) {
    problems.push('压缩包里没有 project.json —— 这不是一个工程包')
  }

  // ★ 排序遍历：与 pack 侧同样理由（written 顺序确定，便于测试与展示）
  for (const name of Object.keys(unzipped).sort()) {
    const bytes2 = unzipped[name]
    if (name.endsWith('/') || (bytes2.length === 0 && name.endsWith('/'))) {
      // 目录条目：mkdir 即可（mkdir 幂等）
      await fs.mkdir(joinRel(rootDir, name))
      continue
    }
    // ★ 路径越界必须拦：zip 是从别人那儿收来的，
    //   含 `../../` 的条目若原样写入就能写到工程目录之外。
    const rel = name.replace(/\\/g, '/') // ③ Windows 分隔符归一
    const safe = isSafeZipPath(rel)
    if (!safe) {
      problems.push(`跳过非法路径条目：${name}`)
      continue
    }
    const bytes = bytes2
    const full = joinRel(rootDir, rel)
    await fs.mkdir(joinRel(rootDir, rel.slice(0, rel.lastIndexOf('/'))))

    // ★ 文本类走 writeText，其余走 writeBytes —— 不是「一律 writeBytes」。
    //   原因：ProjectFs 的 readText 只读文本表。若解包时把 project.json
    //   写进字节表，loadProject 的 readText 就读不到，
    //   症状是「解包成功，打开工程报『没有 project.json』」——
    //   而工程包是好的，纯属自伤，且用户完全无从下手。
    //   判据用「是不是合法 UTF-8 文本」而不是「扩展名是不是 .json」：
    //   data/ 下的外部数据也可能是文本，而中文文件名不能靠扩展名猜。
    if (looksLikeUtf8Text(bytes)) {
      await fs.writeText(full, new TextDecoder().decode(bytes))
    } else if (fs.writeBytes) {
      await fs.writeBytes(full, bytes)
    } else {
      // ★ 绝不静默降级成 writeText。
      //   那条路是「UTF-8 解码 → 存字符串」：docx/xlsx/png 一进去就损坏，
      //   而工程表面上是打开的 —— 用户几天后双击文档才发现坏了，且完全无法追溯
      //   （因为解包当时报了「成功」）。这属于最坏的一类静默失败。
      problems.push(
        `当前存储不支持二进制写入，无法解包（需要 ProjectFs.writeBytes）：${rel}`,
      )
      return { written, problems }
    }
    written.push(rel)
  }

  // ★ 解包后必建的目录：即使 zip 里没有对应目录条目，也先建出来。
  //   否则打开工程报「缺 nodes 目录」这类结构错误，而 zip 本身是好的。
  for (const d of TOP_DIRS) {
    await fs.mkdir(`${rootDir}/${d}`)
  }

  return { written, problems }
}

/**
 * 判断一段字节是否是「可存成文本的 UTF-8」。
 *
 * ★ 判据是**能否无损往返**（decode 后再 encode 字节不变），而不是扩展名。
 *   理由：
 *   ① 扩展名会骗人：.json 可能是 UTF-16，.md 可能是 GBK
 *   ② 反过来，data/ 下的外部数据（CSV 价目表、文本图纸说明）也常是文本，
 *      它们不该被当二进制
 *   ③ 中文文件名不能靠扩展名推断编码
 *
 * 判据：解码后重新编码，字节完全一致 → 纯文本；否则按二进制处理。
 * 这个判定对「恰好是合法 UTF-8 的二进制」会误判成文本（极少见，
 * 且那种文件本来也没有可识别的文本语义）。
 */
function looksLikeUtf8Text(bytes: Uint8Array): boolean {
  try {
    const text = new TextDecoder('utf-8', { fatal: true }).decode(bytes)
    const again = new TextEncoder().encode(text)
    if (again.length !== bytes.length) return false
    for (let i = 0; i < again.length; i++) {
      if (again[i] !== bytes[i]) return false
    }
    return true
  } catch {
    return false
  }
}

/** 拼工程内路径 */
function joinRel(rootDir: string, rel: string): string {
  const r = rel.replace(/^\/+/, '')
  return r ? `${rootDir}/${r}` : rootDir
}

/**
 * zip 条目路径是否安全。
 *
 * ★ 安全项：拒绝 `..` 段与绝对路径。
 *   zip slip 漏洞是真实存在的攻击面（解压时写到任意位置），
 *   而工程 zip 是「同事发来的文件」这一典型场景。
 */
function isSafeZipPath(rel: string): boolean {
  if (rel.startsWith('/')) return false
  if (/^[A-Za-z]:/.test(rel)) return false
  for (const seg of rel.split('/')) {
    if (seg === '..') return false
  }
  return true
}