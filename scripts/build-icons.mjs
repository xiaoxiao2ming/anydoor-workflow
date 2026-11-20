/**
 * client-base 图标体系构建（2026-10-07）
 *
 * 依据：docs/icons/ICONS.md 交付规格书（水墨定稿 v11，三件套）
 *   主 Logo   logo-FINAL-v11-transparent.png → client-base-app.ico（垫宣纸白砖）
 *   快捷方式  icon-shortcut-clean.png        → client-base-shortcut.ico
 *   文档类型  icon-document-clean.png        → client-base-doc.ico（.wfc 关联用）
 *
 * ★ 零依赖：PNG 解码（含 5 种滤波逆变换）与面积平均缩放全部手写——
 *   node 内置 zlib 提供解压；不引任何图像库（与 make-wfc-icon 同一技术路线）。
 *
 * 规格书硬约束（不得违反）：
 *   · 墨标记不得重绘/改色/加文字；只允许整体缩放与垫底
 *   · 主 Logo 透明底缩到小尺寸前**必须垫宣纸白砖**（#F5F3EE 圆角），否则黑墨失去可读支撑
 *   · 小尺寸(16/24/32)允许去掉残影墨点——但源图已定稿，此处只做缩放不做修饰
 *
 * 运行：node scripts/build-icons.mjs
 * 产物：build/client-base-{app,shortcut,doc}.ico
 */
import { inflateSync, deflateSync } from 'node:zlib'
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const ASSETS = String.raw`<logo-assets-dir>\client-base-icons_assets`

/* ── PNG 解码 ────────────────────────────────────────────────── */

/** 解码 PNG（8bit，colortype 0/2/3/4/6，非隔行）→ {width,height,rgba} */
export function decodePng (buf) {
  const sig = buf.slice(0, 8)
  if (sig.toString('hex') !== '89504e470d0a1a0a') throw new Error('非 PNG 签名')
  let pos = 8
  let width = 0, height = 0, bitDepth = 0, colorType = 0, interlace = 0
  const idat = []
  let plte = null, trns = null
  while (pos < buf.length) {
    const len = buf.readUInt32BE(pos)
    const type = buf.slice(pos + 4, pos + 8).toString('ascii')
    const data = buf.slice(pos + 8, pos + 8 + len)
    if (type === 'IHDR') {
      width = data.readUInt32BE(0)
      height = data.readUInt32BE(4)
      bitDepth = data[8]
      colorType = data[9]
      interlace = data[12]
    } else if (type === 'PLTE') plte = Buffer.from(data)
    else if (type === 'tRNS') trns = Buffer.from(data)
    else if (type === 'IDAT') idat.push(data)
    else if (type === 'IEND') break
    pos += 12 + len
  }
  if (bitDepth !== 8) throw new Error('仅支持 8bit 位深（源图 ' + bitDepth + '）')
  if (interlace !== 0) throw new Error('不支持隔行 PNG')
  const channels = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[colorType]
  if (!channels) throw new Error('未知 colorType ' + colorType)

  const raw = inflateSync(Buffer.concat(idat))
  const stride = width * channels
  // 逆滤波（PNG filter 0~4）
  const out = Buffer.alloc(stride * height)
  const paeth = (a, b, c) => {
    const p = a + b - c
    const pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c)
    return pa <= pb && pa <= pc ? a : pb <= pc ? b : c
  }
  for (let y = 0; y < height; y++) {
    const f = raw[y * (stride + 1)]
    const rowIn = y * (stride + 1) + 1
    const rowOut = y * stride
    for (let x = 0; x < stride; x++) {
      const rb = raw[rowIn + x]
      const left = x >= channels ? out[rowOut + x - channels] : 0
      const up = y > 0 ? out[(y - 1) * stride + x] : 0
      const ul = y > 0 && x >= channels ? out[(y - 1) * stride + x - channels] : 0
      let v
      switch (f) {
        case 0: v = rb; break
        case 1: v = rb + left; break
        case 2: v = rb + up; break
        case 3: v = rb + ((left + up) >> 1); break
        case 4: v = rb + paeth(left, up, ul); break
        default: throw new Error('未知滤波 ' + f)
      }
      out[rowOut + x] = v & 0xff
    }
  }

  // 归一到 RGBA
  const rgba = Buffer.alloc(width * height * 4)
  if (colorType === 6) {
    // ★ out 布局已是 RGBA 顺序，整块直拷（毫秒级）。
    //   ★★ 原写法 `rgba.copy(rgba, i*4, out, i*4, i*4+4)` 把 Buffer 当 sourceStart 传
    //   （第三参应为数字）——Node 内部强转 NaN→0，每次循环从 0 拷到 i*4 ⇒ O(n²) 字节拷贝，
    //   1536² 约 11 万亿字节永远跑不完 ⇒ 进程被 SIGTERM。这正是 app.ico（logo-FINAL
    //   为 colorType6）此前「从未构建成功、以为是沙箱随机杀」的确定性根因（2026-10-07 实锤：
    //   shortcut/doc 是 colorType2 走 case 2 正常路径所以总能成）。
    out.copy(rgba)
  } else {
    for (let i = 0; i < width * height; i++) {
      switch (colorType) {
        case 2: rgba[i*4]=out[i*3]; rgba[i*4+1]=out[i*3+1]; rgba[i*4+2]=out[i*3+2]; rgba[i*4+3]=255; break
        case 0: rgba[i*4]=rgba[i*4+1]=rgba[i*4+2]=out[i]; rgba[i*4+3]=255; break
        case 4: rgba[i*4]=rgba[i*4+1]=rgba[i*4+2]=out[i*2]; rgba[i*4+3]=out[i*2+1]; break
        case 3: {
          const idx = out[i]
          rgba[i*4]=plte[idx*3]; rgba[i*4+1]=plte[idx*3+1]; rgba[i*4+2]=plte[idx*3+2]
          rgba[i*4+3]=trns && idx < trns.length ? trns[idx] : 255
          break
        }
      }
    }
  }
  return { width, height, rgba }
}

/* ── 缩放（box 面积平均：大比例下采样质量最好）────────────────── */
export function resize (src, sw, sh, dw, dh) {
  const out = Buffer.alloc(dw * dh * 4)
  const xr = sw / dw, yr = sh / dh
  for (let y = 0; y < dh; y++) {
    const sy0 = y * yr, sy1 = Math.min(sh, (y + 1) * yr)
    for (let x = 0; x < dw; x++) {
      const sx0 = x * xr, sx1 = Math.min(sw, (x + 1) * xr)
      let r = 0, g = 0, b = 0, a = 0, wsum = 0
      for (let sy = Math.floor(sy0); sy < sy1; sy++) {
        const hy = Math.min(sy + 1, sy1) - Math.max(sy, sy0)
        for (let sx = Math.floor(sx0); sx < sx1; sx++) {
          const hx = Math.min(sx + 1, sx1) - Math.max(sx, sx0)
          const wgt = hx * hy
          const si = (sy * sw + sx) * 4
          const sa = src[si + 3] / 255
          // 预乘 alpha 混合，避免边缘出现黑晕
          r += src[si] * sa * wgt; g += src[si+1] * sa * wgt; b += src[si+2] * sa * wgt
          a += sa * wgt; wsum += wgt
        }
      }
      const di = (y * dw + x) * 4
      if (a <= 0.0001) { out[di+3] = 0; continue }
      out[di]   = Math.round(r / a)
      out[di+1] = Math.round(g / a)
      out[di+2] = Math.round(b / a)
      out[di+3] = Math.round((a / wsum) * 255)
    }
  }
  return out
}

/* ── 垫宣纸白圆角砖（透明 logo → 有底支撑）───────────────────── */
export function onPaperTile (rgba, w, h, brick = [0xf5, 0xf3, 0xee]) {
  const out = Buffer.from(rgba)
  const pad = Math.round(Math.min(w, h) * 0.06)      // 外边距 6%
  const rx = Math.round(Math.min(w, h) * 0.15)       // 圆角 15%
  const inner = (x, y) => {
    const ix = Math.max(pad + rx - x, x - (w - pad - rx), 0)
    const iy = Math.max(pad + rx - y, y - (h - pad - ry(h, pad, rx)), 0)
    return Math.sqrt(ix*ix + iy*iy) - rx < 0
  }
  const ry = (hh, p, r) => r   // 方形砖：横向纵向同参
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = (y * w + x) * 4
      if (inner(x, y)) {
        // 砖内：logo 不透明处保留 logo；透明处垫砖色
        const a = out[i + 3] / 255
        out[i]   = Math.round(brick[0] * (1 - a) + out[i] * a)
        out[i+1] = Math.round(brick[1] * (1 - a) + out[i+1] * a)
        out[i+2] = Math.round(brick[2] * (1 - a) + out[i+2] * a)
        out[i+3] = 255
      } else {
        out[i + 3] = 0    // 砖外透明
      }
    }
  }
  return out
}

/* ── PNG 编码 + 多尺寸 ICO ───────────────────────────────────── */
function crc32 (buf) {
  // 与 node:zlib.crc32 等价（此处自带，兼容旧 node）
  let table = crc32.table
  if (!table) {
    table = crc32.table = new Int32Array(256)
    for (let n = 0; n < 256; n++) {
      let c = n
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
      table[n] = c
    }
  }
  let crc = -1
  for (let i = 0; i < buf.length; i++) crc = (crc >>> 8) ^ table[(crc ^ buf[i]) & 0xff]
  return (crc ^ -1) >>> 0
}
function chunk (type, data) {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length)
  const t = Buffer.from(type, 'ascii')
  const c = Buffer.alloc(4); c.writeUInt32BE(crc32(Buffer.concat([t, data])))
  return Buffer.concat([len, t, data, c])
}
function encodePng (rgba, w, h) {
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4)
  ihdr[8] = 8; ihdr[9] = 6
  const raw = Buffer.alloc((w * 4 + 1) * h)
  for (let y = 0; y < h; y++) rgba.copy(raw, y * (w * 4 + 1) + 1, y * w * 4, (y + 1) * w * 4)
  return Buffer.concat([
    Buffer.from([0x89,0x50,0x4e,0x47,0x0d,0x0a,0x1a,0x0a]),
    chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw, {level:9})), chunk('IEND', Buffer.alloc(0)),
  ])
}
function buildIco (sizes, pngBySize) {
  const head = Buffer.alloc(6)
  head.writeUInt16LE(0, 0); head.writeUInt16LE(1, 2); head.writeUInt16LE(sizes.length, 4)
  // ★ dir 是**独立缓冲**（不含 6B head）⇒ 条目偏移必须从 0 起算。
  //   写成 `6 + i*16`（把 head 也算进去）会让最后一个条目越界 16 字节 → ERR_OUT_OF_RANGE。
  const dir = Buffer.alloc(16 * sizes.length)
  const blobs = []
  let offset = 6 + 16 * sizes.length
  sizes.forEach((sz, i) => {
    const png = pngBySize[sz]
    const e = i * 16
    dir[e] = sz >= 256 ? 0 : sz
    dir[e+1] = sz >= 256 ? 0 : sz
    dir[e+2] = 0; dir[e+3] = 0
    dir.writeUInt16LE(1, e+4); dir.writeUInt16LE(32, e+6)
    dir.writeUInt32LE(png.length, e+8)
    dir.writeUInt32LE(offset, e+12)
    offset += png.length
    blobs.push(png)
  })
  return Buffer.concat([head, dir, ...blobs])
}

/* ── 主流程 ──────────────────────────────────────────────────── */
const SIZES = [256, 128, 64, 48, 32, 24, 16]
const OUT = join(root, 'build')
mkdirSync(OUT, { recursive: true })

/** 垫砖在中间尺寸做（2026-10-07）：全尺寸(1536²)逐像素垫砖在本机沙箱 SIGTERM
 *  （app.ico 因此一直没出）；512² 计算量降 9 倍，再 box 缩到各尺寸圆角边缘仍平滑。 */
const TILE = 512

/** 源图 → 多尺寸 ico；padBrick=true 时先缩到 TILE 再垫宣纸白圆角砖（app/doc 用） */
function makeIco (srcFile, outFile, { padBrick = false } = {}) {
  const srcBuf = readFileSync(srcFile)
  const { width, height, rgba } = decodePng(srcBuf)
  const small = resize(rgba, width, height, TILE, TILE)
  const base = padBrick ? onPaperTile(small, TILE, TILE) : small
  const pngBySize = {}
  for (const sz of SIZES) {
    const scaled = resize(base, TILE, TILE, sz, sz)
    pngBySize[sz] = encodePng(scaled, sz, sz)
  }
  const ico = buildIco(SIZES, pngBySize)
  writeFileSync(outFile, ico)
  console.log('  ✓ ' + outFile.replace(root, '') + '  ' + ico.length + 'B  (' + SIZES.join('/') + ')')
  // 顺带导出 256 png 供人工目检
  writeFileSync(outFile.replace(/\.ico$/, '-preview.png'), pngBySize[256])
}

/* ── wfc 文档图标（2026-10-07 用户三连指令 + 22:1x 对照 Excel 重做）──
   用户对照 Excel 图标后定调：**干净纸面 + 角上徽章**，不要墨圈泼满整纸。
   ⇒ 白纸（极浅蓝调呼应上一轮「浅蓝」）+ 右上折角 + 主 logo 垫宣纸白砖
     （圆角徽章）嵌纸面中央。程序化绘制，不依赖旧源图。 */
// ★ alpha 侵蚀（min-filter）：边缘 radius 内取最小 alpha —— 消除缩放混色带回的白色 fringe
function erodeAlpha (rgba, w, h, radius = 1) {
  const out = Buffer.from(rgba)
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      let mn = 255
      for (let yy = y - radius; yy <= y + radius; yy++) {
        for (let xx = x - radius; xx <= x + radius; xx++) {
          if (xx < 0 || yy < 0 || xx >= w || yy >= h) { mn = 0; break }
          const a = rgba[(yy * w + xx) * 4 + 3]
          if (a < mn) mn = a
        }
        if (mn === 0) break
      }
      out[(y * w + x) * 4 + 3] = mn
    }
  }
  return out
}

function drawDocPaper (tile, logoRgba, logoW, logoH) {
  const out = Buffer.alloc(tile * tile * 4)
  // ★ 22:2x 按用户反馈重调：① 纸占比回白砖版（pad 6%，纸占 88%——「图标的大小之前
  //   是对的 别改太小」）② logo 回 70% 纸宽（「logo 尺寸没错 改回去」）
  //   ③ 折角几何修正（「折叠的方向错了」）：斜线 dx=dy 右上**全部透明**（切口露背景），
  //     翻折面=斜线左下紧贴带（直角朝左下、贴斜线）——此前误把切口填成了翻折面色。
  // ★ 23:1x 终版（用户：「和 excel 文件一样 长条形 其他全透明」）：
  //   竖长白纸（宽 72% × 高 88% tile，水平居中）+ 右上折角 + 白砖徽章嵌纸面。
  //   纸外全透明（本函数即程序化生成，天然无白底）。
  const w = Math.round(tile * 0.72)
  const h = Math.round(tile * 0.88)
  const x0 = Math.round((tile - w) / 2), y0 = Math.round((tile - h) / 2)
  const x1 = x0 + w, y1 = y0 + h
  const r = Math.round(tile * 0.047)          // 纸圆角
  const fold = Math.round(w * 0.20)           // 折角直角边长（常规比例）
  const paper = [0xff, 0xff, 0xff]            // 白纸面
  const foldCol = [0xd7, 0xe6, 0xf2]          // 翻折面（淡蓝灰，含蓄）
  const edge = [0xc9, 0xd4, 0xde]             // 纸边缘线（浅灰）
  for (let y = 0; y < tile; y++) {
    for (let x = 0; x < tile; x++) {
      const i = (y * tile + x) * 4
      const inRect = x >= x0 && x < x1 && y >= y0 && y < y1
      if (!inRect) continue
      const dx = x - (x1 - fold), dy = y - y0
      if (dx >= 0 && dy >= 0 && dx - dy >= 0) continue   // 切口：斜线右上全部透明
      if ((x < x0 + r || x >= x1 - r) && (y < y0 + r || y >= y1 - r)) {
        const cx = x < x0 + r ? x0 + r : x1 - r
        const cy = y < y0 + r ? y0 + r : y1 - r
        if ((x - cx) ** 2 + (y - cy) ** 2 > r * r) continue
      }
      const isFoldFace = dx >= 0 && dy >= 0 && dy - dx <= (fold >> 1) && dy <= fold
      const col = isFoldFace ? foldCol : paper
      out[i] = col[0]; out[i + 1] = col[1]; out[i + 2] = col[2]; out[i + 3] = 255
      if (Math.abs(dy - dx - (fold >> 1)) <= 1 && isFoldFace) {
        out[i] = 0xbc; out[i + 1] = 0xce; out[i + 2] = 0xde   // 折面斜边描边
      } else if (x <= x0 + 1 || x >= x1 - 2 || y <= y0 + 1 || y >= y1 - 2) {
        out[i] = edge[0]; out[i + 1] = edge[1]; out[i + 2] = edge[2]
      }
    }
  }
  // 主 logo 缩放叠加：宽 = 纸宽 70%（「logo 尺寸没错」），居中略偏上
  const lw = Math.round(w * 0.70)
  const lh = Math.round(logoH * (lw / logoW))
  const lg = resize(logoRgba, logoW, logoH, lw, lh)
  const lx = x0 + Math.round((w - lw) / 2), ly = y0 + Math.round((h - lh) / 2) - Math.round(tile * 0.02)
  for (let y = 0; y < lh; y++) {
    for (let x = 0; x < lw; x++) {
      const si = (y * lw + x) * 4
      const sa = lg[si + 3] / 255
      if (sa <= 0.01) continue
      const di = ((ly + y) * tile + (lx + x)) * 4
      const da = out[di + 3] / 255
      const a = sa + da * (1 - sa)
      out[di] = Math.round(lg[si] * sa + out[di] * da * (1 - sa))
      out[di + 1] = Math.round(lg[si + 1] * sa + out[di + 1] * da * (1 - sa))
      out[di + 2] = Math.round(lg[si + 2] * sa + out[di + 2] * da * (1 - sa))
      out[di + 3] = Math.round(a * 255)
    }
  }
  return out
}

const A = (f) => join(ASSETS, f)
console.log('=== 图标体系构建（水墨定稿 v11 · 圆角矩形版）===')
// --app-only / --doc-only：单张重跑开关（沙箱对长进程随机 SIGTERM，分张跑提高成活率）
const APP_ONLY = process.argv.includes('--app-only')
const DOC_ONLY = process.argv.includes('--doc-only')
// ★ 2026-10-07 用户：「wfc 文件的样式用在软件快捷方式和任务栏等地方」——
//   shortcut 产物与 wfc 文档图标**同款**（主 logo 垫宣纸白圆角砖，不再用旧角标源图）。
const LOGO = A('logo-FINAL-v11-transparent.png')
if (!APP_ONLY && !DOC_ONLY) {
  // ★ 22:4x 用户：「快捷方式的图片改成最开始的圆角矩形 改成蓝底的」——
  //   回到圆角砖形（onPaperTile），砖底改天蓝（与编辑器背景同色系），墨圈 logo 保留。
  const ld = decodePng(readFileSync(LOGO))
  const logoSmall = resize(ld.rgba, ld.width, ld.height, TILE, TILE)
  const blue = onPaperTile(logoSmall, TILE, TILE, [0x87, 0xce, 0xeb])
  const p2png = {}
  for (const sz of SIZES) p2png[sz] = encodePng(resize(blue, TILE, TILE, sz, sz), sz, sz)
  writeFileSync(join(OUT, 'client-base-shortcut.ico'), buildIco(SIZES, p2png))
  writeFileSync(join(OUT, 'client-base-shortcut-preview.png'), p2png[256])
  console.log('  ✓ \\build\\client-base-shortcut.ico  (圆角矩形 · 蓝底)')
}
if (!APP_ONLY) {
  // ★ 23:1x 终版（用户：「和 excel 文件一样 长条形 其他全透明」）—— drawDocPaper
  //   程序化长条白纸 + 折角 + 白砖徽章（源图 flood-fill 方案弃用：源图近方形且白底）。
  const ld0 = decodePng(readFileSync(LOGO))
  const paper0 = drawDocPaper(TILE, ld0.rgba, ld0.width, ld0.height)
  const p0png = {}
  for (const sz of SIZES) p0png[sz] = encodePng(resize(paper0, TILE, TILE, sz, sz), sz, sz)
  writeFileSync(join(OUT, 'client-base-doc.ico'), buildIco(SIZES, p0png))
  writeFileSync(join(OUT, 'client-base-doc-preview.png'), p0png[256])
  console.log('  ✓ \\build\\client-base-doc.ico  (长条白纸+折角+徽章 · 透明底)')
}
console.log('完成。')
