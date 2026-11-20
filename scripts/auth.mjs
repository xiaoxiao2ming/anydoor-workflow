/**
 * 网页端访问口令（2026-10-17）
 *
 * 起因：路由器把 5678–5681 映射到群晖，画布从此可以从公网访问。
 * 服务原本零鉴权（谁拿到 IP:端口 都能读画布、改数据），必须先上一道门。
 *
 * 设计取舍：
 *   - 单用户口令，不用账号体系：`APP_PASSWORD` 未设 = 关闭验证（本地开发零打扰）。
 *   - 无状态会话：cookie 值 = HMAC-SHA256(口令, 固定串)。改口令 → 所有旧 cookie 立刻失效，
 *     不需要服务端 session 表，容器重启不掉登录态。
 *   - HttpOnly + SameSite=Lax：JS 拿不到，跨站表单提交也带不上。
 *   - 只有密码没有用户名：这是「门锁」不是「账号系统」，够用且少一个泄露面。
 *   - 失败限流按 IP 计数：公网暴露下必须挡暴力破解（5 次 / 10 分钟）。
 *
 * 界面仿 DSM 登录页（用户指令：仿照群晖那个登录界面）：
 *   深蓝渐变背景 + 居中白卡片 + 单密码框 + 蓝色登录按钮。
 */

import { createHash, createHmac, timingSafeEqual } from 'node:crypto'

export const COOKIE_NAME = 'cb_auth'
const SESSION_SALT = 'client-base-session-v1'
const MAX_AGE_REMEMBER = 30 * 24 * 3600 // 30 天

/** 是否开启口令验证（未设 APP_PASSWORD 即关闭） */
export function authEnabled () {
  return !!String(process.env.APP_PASSWORD || '')
}

/** 由口令推导会话票据；口令变则票据变 */
export function sessionToken (password = String(process.env.APP_PASSWORD || '')) {
  if (!password) return ''
  return createHmac('sha256', password).update(SESSION_SALT).digest('hex')
}

const sha256 = (s) => createHash('sha256').update(String(s)).digest()

/** 定长比较，避免时序侧信道 */
export function passwordMatches (input) {
  const pw = String(process.env.APP_PASSWORD || '')
  if (!pw) return true
  return timingSafeEqual(sha256(input || ''), sha256(pw))
}

/** 解析 Cookie 头 */
export function parseCookies (req) {
  const out = {}
  const raw = req.headers.cookie || ''
  for (const part of raw.split(';')) {
    const i = part.indexOf('=')
    if (i < 0) continue
    const k = part.slice(0, i).trim()
    if (!k) continue
    try {
      out[k] = decodeURIComponent(part.slice(i + 1).trim())
    } catch {
      out[k] = part.slice(i + 1).trim()
    }
  }
  return out
}

/** 本次请求是否已通过口令 */
export function isAuthed (req) {
  if (!authEnabled()) return true
  const got = parseCookies(req)[COOKIE_NAME]
  if (!got) return false
  const want = sessionToken()
  if (got.length !== want.length) return false
  try {
    return timingSafeEqual(Buffer.from(got, 'utf8'), Buffer.from(want, 'utf8'))
  } catch {
    return false
  }
}

export function setSessionCookie (res, remember) {
  const attrs = [
    `${COOKIE_NAME}=${sessionToken()}`,
    'Path=/',
    'HttpOnly',
    'SameSite=Lax',
    remember ? `Max-Age=${MAX_AGE_REMEMBER}` : '',
  ].filter(Boolean)
  res.setHeader('Set-Cookie', attrs.join('; '))
}

export function clearSessionCookie (res) {
  res.setHeader('Set-Cookie', `${COOKIE_NAME}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0`)
}

/** next 只允许站内路径，挡掉 //evil.com、http://... 这类开放跳转 */
export function safeNext (next) {
  const s = String(next || '')
  if (!s.startsWith('/') || s.startsWith('//')) return '/'
  if (s.startsWith('/login')) return '/'
  return s
}

/* ── 失败限流：5 次 / 10 分钟，按 IP ─────────────────────────── */
const FAIL_WINDOW_MS = 10 * 60 * 1000
const FAIL_MAX = 5
const failures = new Map() // ip -> { count, first, until }

export function clientIp (req) {
  const xf = String(req.headers['x-forwarded-for'] || '').split(',')[0].trim()
  return xf || req.socket.remoteAddress || 'unknown'
}

export function checkLimit (ip) {
  const rec = failures.get(ip)
  if (!rec) return { ok: true }
  if (rec.until && rec.until > Date.now()) {
    return { ok: false, retrySec: Math.ceil((rec.until - Date.now()) / 1000) }
  }
  return { ok: true }
}

export function noteFailure (ip) {
  const now = Date.now()
  let rec = failures.get(ip)
  if (!rec || now - rec.first > FAIL_WINDOW_MS) rec = { count: 0, first: now, until: 0 }
  rec.count += 1
  if (rec.count > FAIL_MAX) rec.until = now + FAIL_WINDOW_MS
  failures.set(ip, rec)
  return { count: rec.count, until: rec.until }
}

export function noteSuccess (ip) {
  failures.delete(ip)
}

/* ── 请求体：登录表单是 urlencoded，不是 JSON ─────────────────── */
export function readFormBody (req, limit = 8192) {
  return new Promise((resolve) => {
    let buf = ''
    let done = false
    const finish = () => {
      if (done) return
      done = true
      resolve(buf)
    }
    req.on('data', (c) => {
      if (done) return
      buf += c
      if (buf.length > limit) {
        done = true
        resolve('')
        req.destroy()
      }
    })
    req.on('end', finish)
    req.on('error', finish)
    req.setTimeout(5000, () => {
      finish()
      req.destroy()
    })
  })
}

const esc = (s) =>
  String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]))

/**
 * 登录页 HTML（自包含：无外链、无字体文件，断网也能开）
 * @param {{error?: string, next?: string, version?: string, host?: string, notice?: string}} o
 */
export function loginPage (o = {}) {
  const err = o.error ? `<div class="msg">${esc(o.error)}</div>` : ''
  const notice = o.notice ? `<div class="msg info">${esc(o.notice)}</div>` : ''
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>登录 · 客户端底座</title>
<style>
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
  html,body{height:100%}
  body{
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;
    background:#071019;color:#1c1f23;display:flex;align-items:center;justify-content:center;
    padding:24px;overflow:hidden;position:relative;
  }
  /* 背景：深蓝底 + 两团青/蓝柔光，接近 DSM 登录页的观感 */
  .bg{position:fixed;inset:0;z-index:0;
    background:
      radial-gradient(80% 55% at 18% 8%, rgba(58,140,170,.55) 0%, rgba(58,140,170,0) 60%),
      radial-gradient(70% 60% at 88% 92%, rgba(46,105,196,.45) 0%, rgba(46,105,196,0) 62%),
      linear-gradient(160deg,#123047 0%,#0c1c2a 45%,#060d15 100%);
  }
  .bg::after{content:"";position:absolute;inset:0;
    background-image:radial-gradient(rgba(255,255,255,.045) 1px, transparent 1px);
    background-size:3px 3px;opacity:.5}
  .card{position:relative;z-index:1;width:100%;max-width:392px;background:#fff;border-radius:10px;
    padding:38px 36px 26px;box-shadow:0 24px 70px rgba(0,0,0,.45),0 2px 8px rgba(0,0,0,.25);
    animation:in .32s ease-out both}
  @keyframes in{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
  .logo{display:flex;justify-content:center;margin-bottom:16px}
  .logo img{width:64px;height:64px;border-radius:14px;box-shadow:0 6px 18px rgba(0,0,0,.30),0 1px 3px rgba(0,0,0,.22);display:block}
  h1{font-size:21px;font-weight:600;text-align:center;letter-spacing:.5px}
  .sub{font-size:12.5px;color:#7a838f;text-align:center;margin-top:7px;line-height:1.6}
  form{margin-top:24px}
  .field{position:relative}
  input[type=password],input[type=text]{
    width:100%;height:46px;padding:0 44px 0 14px;font-size:15px;color:#1c1f23;background:#f7f9fb;
    border:1px solid #d8dee6;border-radius:5px;outline:none;transition:border-color .15s,box-shadow .15s,background .15s}
  input::placeholder{color:#9aa4b2}
  input:focus{background:#fff;border-color:#1a82c4;box-shadow:0 0 0 3px rgba(26,130,196,.14)}
  .eye{position:absolute;right:6px;top:6px;width:34px;height:34px;border:0;background:transparent;
    cursor:pointer;font-size:15px;opacity:.45;border-radius:50%}
  .eye:hover{opacity:.9;background:#eef3f8}
  .msg{margin-top:12px;padding:9px 12px;border-radius:5px;font-size:12.5px;line-height:1.5;
    background:#fdecec;color:#bf3b34;border:1px solid #f6d3d0}
  .msg.info{background:#eef6fd;color:#21618c;border-color:#cfe4f5}
  .remember{display:flex;align-items:center;gap:8px;margin:14px 0 18px;font-size:12.5px;color:#5d6773;
    cursor:pointer;user-select:none}
  .remember input{width:15px;height:15px;accent-color:#1a82c4;cursor:pointer}
  .submit{width:100%;height:46px;border:0;border-radius:5px;color:#fff;font-size:15px;font-weight:500;
    letter-spacing:1px;cursor:pointer;background:linear-gradient(180deg,#1f8ccd,#1673ad);
    box-shadow:0 2px 6px rgba(22,115,173,.32);transition:filter .15s,transform .06s}
  .submit:hover{filter:brightness(1.06)}
  .submit:active{transform:translateY(1px)}
  .foot{margin-top:22px;padding-top:14px;border-top:1px solid #eef1f5;display:flex;
    justify-content:space-between;font-size:11px;color:#a3abb6;letter-spacing:.3px}
  @media (max-width:420px){.card{padding:30px 22px 20px}}
  @media (prefers-color-scheme:dark){
    /* 卡片保持浅色：与 DSM 登录页一致，不随系统主题跳变 */
  }
</style>
</head>
<body>
<div class="bg"></div>
<main class="card">
  <div class="logo"><img src="/logo-tile.png" alt="客户端底座" width="64" height="64" onerror="this.style.display='none'"></div>
  <h1>客户端底座</h1>
  <p class="sub">画布工作台 · 请输入访问密码</p>
  <form method="post" action="/login" autocomplete="off">
    <input type="hidden" name="next" value="${esc(o.next || '/')}">
    <div class="field">
      <input id="pw" type="password" name="password" placeholder="密码" autocomplete="current-password"
             ${o.error ? 'autofocus' : 'autofocus'} required>
      <button class="eye" id="eye" type="button" title="显示密码" aria-label="显示密码">👁</button>
    </div>
    ${err}${notice}
    <label class="remember"><input type="checkbox" name="remember" value="1" checked> 记住这台设备 30 天</label>
    <button class="submit" type="submit">登 录</button>
  </form>
  <div class="foot"><span>client-base ${esc(o.version || '')}</span><span>${esc(o.host || '')}</span></div>
</main>
<script>
  var pw = document.getElementById('pw'), eye = document.getElementById('eye');
  eye.addEventListener('click', function () {
    var show = pw.type === 'password';
    pw.type = show ? 'text' : 'password';
    eye.title = show ? '隐藏密码' : '显示密码';
    pw.focus();
  });
</script>
</body>
</html>`
}

/**
 * 门禁：返回 true 表示已经接管响应（调用方直接 return）
 * @param {{req: import('node:http').IncomingMessage, res: import('node:http').ServerResponse,
 *          url: URL, version?: string, log?: (line: string) => void}} ctx
 */
export async function handleAuth (ctx) {
  const { req, res, url } = ctx
  const path = url.pathname

  /* 登录页 */
  if (path === '/login') {
    if (req.method === 'POST') {
      const ip = clientIp(req)
      const lim = checkLimit(ip)
      if (!lim.ok) {
        res.writeHead(429, { 'Content-Type': 'text/html; charset=utf-8', 'Retry-After': String(lim.retrySec) })
        res.end(loginPage({
          error: `尝试次数过多，请 ${Math.ceil(lim.retrySec / 60)} 分钟后再试`,
          version: ctx.version,
          host: ctx.host,
        }))
        return true
      }
      const form = new URLSearchParams(await readFormBody(req))
      const pw = form.get('password') || ''
      const remember = form.get('remember') === '1'
      const next = safeNext(form.get('next'))
      if (passwordMatches(pw)) {
        noteSuccess(ip)
        setSessionCookie(res, remember)
        ctx.log?.(`[auth] 登录成功 ip=${ip} remember=${remember ? 1 : 0}`)
        res.writeHead(302, { Location: next, 'Cache-Control': 'no-store' })
        res.end()
      } else {
        const rec = noteFailure(ip)
        ctx.log?.(`[auth] 登录失败 ip=${ip} 第${rec.count}次${rec.until ? ' → 已临时封禁' : ''}`)
        res.writeHead(401, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' })
        res.end(loginPage({
          error: '密码不正确',
          next,
          version: ctx.version,
          host: ctx.host,
        }))
      }
      return true
    }
    // 已登录再访问 /login → 直接回首页
    if (isAuthed(req)) {
      res.writeHead(302, { Location: '/' })
      res.end()
      return true
    }
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' })
    res.end(loginPage({ next: safeNext(url.searchParams.get('next')), version: ctx.version, host: ctx.host }))
    return true
  }

  /* 退出 */
  if (path === '/logout') {
    clearSessionCookie(res)
    res.writeHead(302, { Location: '/login', 'Cache-Control': 'no-store' })
    res.end()
    return true
  }

  /* 品牌 logo 放行：登录页要用（必须能在鉴权前加载）。
   * 文件在 dist-web 根下，是静态资源不是数据接口，泄露面为零。 */
  if (path === '/logo-tile.png' || path === '/favicon.ico') return false

  /* OO 远程 DS 回访通道（2026-10-17）：DocumentServer 没有登录态，
   * 拉源文件/回传产物走 URL 里的 32 位随机 token 门禁（scripts/oo-remote.mjs 校验）。 */
  if (path.startsWith('/office-src/') || path === '/doc/callback') return false

  if (isAuthed(req)) return false

  /* 未登录：文档导航 → 跳登录页；fetch/XHR → 401 JSON（前端好识别）。
   * 判定优先用 sec-fetch-dest（现代浏览器都发）；
   * 缺失时（curl / 老浏览器 / 部分内嵌 WebView）按「GET 且没标 x-requested-with」当导航处理。 */
  const dest = String(req.headers['sec-fetch-dest'] || '')
  const wantsHtml = dest === 'document' ||
    (dest === '' && (req.method === 'GET' || req.method === 'HEAD') && !req.headers['x-requested-with']) ||
    (String(req.headers.accept || '').includes('text/html') && !path.startsWith('/api/'))
  if (wantsHtml) {
    const to = path + (url.search || '')
    res.writeHead(302, { Location: '/login?next=' + encodeURIComponent(to), 'Cache-Control': 'no-store' })
    res.end()
  } else {
    res.writeHead(401, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' })
    res.end(JSON.stringify({ error: 'auth_required', login: '/login' }))
  }
  return true
}
