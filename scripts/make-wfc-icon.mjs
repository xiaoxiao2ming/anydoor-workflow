/**
 * 生成 .wfc 文件关联图标（wfc.ico）+ 资源管理器注册脚本
 *
 * 立档：2026-10-07 · 用户：「1、给*.wfc文件加个图例。不要默认是白板 2、双击*.wfc文件打开软件」
 *
 * 图标风格照抄页面 favicon（index.html 里的 data-uri svg）：
 *   暗底圆角矩形 #0d1117 + 斜连线 #30363d + 蓝/绿双节点 #58a6ff/#7ee787
 *   —— 用户在画布上天天见的视觉，认得出是「这个软件的画布」。
 *
 * ★ 技术选型：PNG-in-ICO（Vista+ 标准特性），node 内置 zlib 即可编码，
 *   不需要任何图像库。ICO 头 6B + 目录项 16B + PNG 数据。
 *
 * 运行：node scripts/make-wfc-icon.mjs
 */
import { deflateSync, crc32 } from 'node:zlib'
import { writeFileSync, mkdirSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')

/* ── PNG 编码（RGBA8，无 filter 优化）───────────────────────── */
function chunk (type, data) {
  const len = Buffer.alloc(4)
  len.writeUInt32BE(data.length)
  const typeBuf = Buffer.from(type, 'ascii')
  const crcBuf = Buffer.alloc(4)
  crcBuf.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])) >>> 0)
  return Buffer.concat([len, typeBuf, data, crcBuf])
}

function encodePng (rgba, w, h) {
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(w, 0)
  ihdr.writeUInt32BE(h, 4)
  ihdr[8] = 8    // bit depth
  ihdr[9] = 6    // color type: RGBA
  // 10-12: compression/filter/interlace 全 0
  // 每行前置 filter byte 0
  const raw = Buffer.alloc((w * 4 + 1) * h)
  for (let y = 0; y < h; y++) {
    raw[y * (w * 4 + 1)] = 0
    rgba.copy(raw, y * (w * 4 + 1) + 1, y * w * 4, (y + 1) * w * 4)
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), // PNG 签名
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ])
}

/* ── 绘制：256×256 RGBA ──────────────────────────────────────── */
const S = 256
const px = Buffer.alloc(S * S * 4)

// 颜色（与 favicon svg 一致）
const BG = [0x0d, 0x11, 0x17]        // 暗底
const LINE = [0x30, 0x36, 0x3d]      // 斜连线
const BLUE = [0x58, 0xa6, 0xff]      // 上游节点
const GREEN = [0x7e, 0xe7, 0x87]     // 下游节点

/** 有符号距离平方：点到线段 AB */
function distToSeg2 (px_, py, ax, ay, bx, by) {
  const dx = bx - ax, dy = by - ay
  const t = Math.max(0, Math.min(1, ((px_ - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)))
  const cx = ax + t * dx, cy = ay + t * dy
  return (px_ - cx) ** 2 + (py - cy) ** 2
}

// 几何参数（favicon 32 → ×8 = 256）
const k = 8
const rx = 7 * k                    // 圆角半径
const lineW = 2.4 * k               // 连线粗
const r1 = 3.6 * k                  // 节点半径
const A = [10 * k, 11 * k]
const B = [22 * k, 21 * k]
const pad = k                       // 外边距
const inner = S - pad * 2           // 圆角矩形内区

// 命中判定（带 2px 抗锯齿过渡带）
const cov = (d2, r) => {
  const d = Math.sqrt(d2)
  return Math.max(0, Math.min(1, (r + 1 - d) / 2))
}

for (let y = 0; y < S; y++) {
  for (let x = 0; x < S; x++) {
    const i = (y * S + x) * 4
    // ① 圆角矩形背景（SDF：到内缩矩形的距离）
    const ix = Math.max(pad + rx - x, x - (S - pad - rx), 0)
    const iy = Math.max(pad + rx - y, y - (S - pad - rx), 0)
    const dRect = Math.sqrt(ix * ix + iy * iy) - rx
    const aBg = Math.max(0, Math.min(1, (1 - dRect) / 2))
    if (aBg <= 0) { px[i + 3] = 0; continue }
    // ② 连线（先画，节点盖在上面）
    const dLine = distToSeg2(x + .5, y + .5, A[0], A[1], B[0], B[1])
    const aLine = cov(dLine, lineW / 2)
    // ③ 两个节点圆
    const dA = (x + .5 - A[0]) ** 2 + (y + .5 - A[1]) ** 2
    const dB = (x + .5 - B[0]) ** 2 + (y + .5 - B[1]) ** 2
    const aA = cov(dA, r1)
    const aB = cov(dB, r1)

    // 合成：bg → line → blue → green（后者覆盖前者）
    let r = BG[0], g = BG[1], b = BG[2], a = aBg
    const over = (c, ca) => {
      const na = ca + a * (1 - ca)
      if (na <= 0) return
      r = (c[0] * ca + r * a * (1 - ca)) / na
      g = (c[1] * ca + g * a * (1 - ca)) / na
      b = (c[2] * ca + b * a * (1 - ca)) / na
      a = na
    }
    over(LINE, aLine)
    over(BLUE, aA)
    over(GREEN, aB)

    px[i] = Math.round(r)
    px[i + 1] = Math.round(g)
    px[i + 2] = Math.round(b)
    px[i + 3] = Math.round(a * 255)
  }
}

/* ── PNG → ICO（256px 单尺寸，PNG-in-ICO）────────────────────── */
const png = encodePng(px, S, S)
const icon = Buffer.alloc(6 + 16)
icon.writeUInt16LE(0, 0)        // reserved
icon.writeUInt16LE(1, 2)        // type: icon
icon.writeUInt16LE(1, 4)        // count: 1
// 目录项
icon[6] = 0                     // width 256 → 0
icon[7] = 0                     // height 256 → 0
icon[8] = 0                     // colors
icon[9] = 0                     // reserved
icon.writeUInt16LE(1, 10)       // planes
icon.writeUInt16LE(32, 12)      // bit count
icon.writeUInt32BE(png.length, 14)
icon.writeUInt32BE(22, 18)      // offset

const out = join(root, 'build', 'wfc.ico')
mkdirSync(dirname(out), { recursive: true })
writeFileSync(out, Buffer.concat([icon, png]))
console.log('已生成', out, (6 + 16 + png.length), '字节')
