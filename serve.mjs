// 极简静态服务器：给 dist-web 提供 http 访问，含正确的 .wasm MIME（file:// 会卡）
import { createServer } from 'node:http'
import { readFile, stat } from 'node:fs/promises'
import { join, extname, normalize } from 'node:path'

const ROOT = process.argv[2] ?? 'dist-web'
const PORT = Number(process.argv[3] ?? 5174)
const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.wasm': 'application/wasm',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.map': 'application/json; charset=utf-8',
  '.ico': 'image/x-icon',
}

const server = createServer(async (req, res) => {
  try {
    let p = normalize(join(ROOT, '.' + new URL(req.url, 'http://x').pathname))
    let s
    try {
      s = await stat(p)
    } catch {
      // 单页应用兜底：找不到就回 index.html
      p = join(ROOT, 'index.html')
      s = await stat(p)
    }
    if (s.isDirectory()) {
      p = join(p, 'index.html')
    }
    const data = await readFile(p)
    res.setHeader('Content-Type', MIME[extname(p).toLowerCase()] ?? 'application/octet-stream')
    res.setHeader('Cache-Control', 'no-cache')
    res.end(data)
  } catch (e) {
    res.statusCode = 404
    res.end('not found: ' + (e instanceof Error ? e.message : String(e)))
  }
})

server.listen(PORT, '127.0.0.1', () => {
  console.log(`serving ${ROOT} at http://127.0.0.1:${PORT}/`)
})
