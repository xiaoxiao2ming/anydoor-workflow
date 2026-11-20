/**
 * DocService(3798) 幂等启动器 —— OO 编辑器协作链的必跑依赖
 *
 * ★ 背景（2026-10-07 排障实锤）：
 *   oo-core.mjs 把 /doc/ 与 /cache/ 反代到官方 DocService(127.0.0.1:3798)，
 *   编辑器「标准服务器流」（auth → open → Editor.bin 下发）全走这条 socket.io 通道。
 *   3798 不在跑 → 同源 /doc 502 "docservice down" → 编辑器永远 loading，
 *   cellDiag 诊断打「editor=无, DocsAPI=undefined」死循环（资源层全 200 也白搭）。
 *   9-30 它是手动起的，没进任何启动链，重启机器后就丢了 —— 本文件就是补这个洞。
 *
 * 设计：
 *   - 幂等：先探 3798，活着就直接返回（多入口重复调用安全）
 *   - detached：子进程独立于调用方生命周期（vite 重启 / Electron 退出不连带杀 DS）
 *   - 日志：vendors/oo-docservice-log.txt（append，与历史日志同文件）
 *   - 启动参数：NODE_ENV=development-windows + NODE_CONFIG_DIR=Common/config
 *     （缺 NODE_ENV 时 logger.js 读 log.filePath 为空直接崩 —— 2026-10-07 实测）
 *
 * 用法：
 *   node scripts/oo-docservice.mjs            # 探测+按需启动，探针输出
 *   npm run oo:ds                             # 同上
 *   import { ensureDocService } from ...      # 供 poc-server / main.ts 调用
 */
import { spawn } from 'node:child_process'
import net from 'node:net'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

/** 探测 3798 是否可连（150ms 超时，DocService 本机回环） */
export function isDocServiceAlive (port = 3798, timeoutMs = 150) {
  return new Promise((resolve) => {
    const sock = net.connect({ host: '127.0.0.1', port })
    const done = (ok) => { try { sock.destroy() } catch { /* ignore */ } resolve(ok) }
    sock.setTimeout(timeoutMs, () => done(false))
    sock.on('connect', () => done(true))
    sock.on('error', () => done(false))
  })
}

/** 解析 oo-server 根（调用方可能从任意 CWD / 打包布局触发） */
export function resolveOoServerRoot (startDir = __dirname) {
  // repo 布局：scripts/oo-docservice.mjs → ../vendors/oo-server
  // 打包布局：dist-electron/main.js 内联了本逻辑 → 向上找 resources/app/vendors/oo-server
  let dir = path.resolve(startDir)
  for (let i = 0; i < 8; i++) {
    const cand = path.join(dir, 'vendors', 'oo-server')
    if (fs.existsSync(path.join(cand, 'DocService', 'sources', 'server.js'))) return cand
    const parent = path.dirname(dir)
    if (parent === dir) break
    dir = parent
  }
  return null
}

/**
 * 确保 DocService 在跑：活→返回 {started:false}；死→detached 启动并探活。
 * 返回 { started, alive, root, log } 供调用方打日志。
 */
export async function ensureDocService (opts = {}) {
  const quiet = opts.quiet === true
  const log = (...a) => { if (!quiet) console.log('[oo-docservice]', ...a) }
  if (await isDocServiceAlive()) {
    log('DocService(3798) 已在跑，跳过启动')
    return { started: false, alive: true }
  }
  const root = resolveOoServerRoot()
  if (!root) {
    log('!! 未找到 vendors/oo-server/DocService（从', __dirname, '向上 8 级探测失败）')
    return { started: false, alive: false, root: null }
  }
  const serverJs = path.join(root, 'DocService', 'sources', 'server.js')
  const commonCfg = path.join(root, 'Common', 'config')
  /* ★ 2026-10-23 x2t 崩溃自愈（0xC0000005 v8::ToLocalChecked）：
     FileConverter/editors 目录迁移缺失，DoctRenderer.config 的 ../editors/* 指向
     全落空 → x2t 启动即崩。启动时检测旧路径并重定向到 vendors/onlyoffice
     既有资产（git 追踪，换机不丢）。config 本身是运行态（gitignore 目录）——
     修复代码化在这里，每次启动幂等检测。 */
  try {
    const cfgPath = path.join(root, 'FileConverter', 'Bin', 'DoctRenderer.config')
    if (fs.existsSync(cfgPath)) {
      const cfgTxt = fs.readFileSync(cfgPath, 'utf8')
      if (cfgTxt.includes('../editors/')) {
        const onlyofficeOk = ['sdkjs/common/Native/native.js', 'sdkjs/common/AllFonts.js', 'web-apps/vendor/xregexp/xregexp-all-min.js']
          .every((p) => fs.existsSync(path.join(root, '..', 'onlyoffice', p)))
        if (onlyofficeOk) {
          fs.writeFileSync(cfgPath, [
            '<Settings>',
            '<file>../../../onlyoffice/sdkjs/common/Native/native.js</file>',
            '<file>../../../onlyoffice/sdkjs/common/Native/jquery_native.js</file>',
            '<allfonts>../../../onlyoffice/sdkjs/common/AllFonts.js</allfonts>',
            '<file>../../../onlyoffice/web-apps/vendor/xregexp/xregexp-all-min.js</file>',
            '<sdkjs>../../../onlyoffice/sdkjs</sdkjs>',
            '<dictionaries>../../dictionaries</dictionaries>',
            '</Settings>',
            '',
          ].join('\n'))
          log('DoctRenderer.config 已重定向 onlyoffice 资产（x2t 崩溃自愈）')
        } else {
          log('!! vendors/onlyoffice 资产不全，DoctRenderer.config 保持原样（x2t 转换可能崩溃）')
        }
      }
    }
  } catch (eCfg) { /* 检测失败不阻断 DS 启动 */ }
  /* ★ 2026-10-23 中文字体自愈（PDF 转换乱码根因）：AllFonts.js 声明 642 个字体，
     其中 415 个在 /sysfonts/（宋体/雅黑/黑体/楷体等 CJK）——Bin/sysfonts 目录
     迁移丢失（运行态资产不进 git，老机器物理存在）。启动时对照 AllFonts.js 清单
     从 C:/Windows/Fonts 幂等补齐（缺即复制；系统也没有的——WPS 特殊符号类——跳过）。 */
  try {
    const binDir = path.join(root, 'FileConverter', 'Bin')
    const allFontsPath = path.join(binDir, 'AllFonts.js')
    if (fs.existsSync(allFontsPath)) {
      const filesMatch = fs.readFileSync(allFontsPath, 'utf8').match(/window\["__fonts_files"\] = \[([\s\S]*?)\];/)
      const sysDir = path.join(binDir, 'sysfonts')
      if (filesMatch) {
        const files = JSON.parse('[' + filesMatch[1].replace(/,\s*$/, '') + ']')
        const need = files.filter((f) => typeof f === 'string' && f.startsWith('/sysfonts/')).map((f) => f.slice('/sysfonts/'.length))
        let copied = 0
        if (need.length) {
          fs.mkdirSync(sysDir, { recursive: true })
          for (const name of need) {
            const dst = path.join(sysDir, name)
            if (fs.existsSync(dst)) continue
            const src = path.join('C:', 'Windows', 'Fonts', name)
            if (fs.existsSync(src)) {
              try { fs.copyFileSync(src, dst); copied++ } catch { /* 单个失败跳过 */ }
            }
          }
        }
        const have = fs.existsSync(sysDir) ? fs.readdirSync(sysDir).length : 0
        if (copied > 0) log(`sysfonts 中文字体已从系统补齐 ${copied} 个（共 ${have}，清单 ${need.length}）`)
        else if (have === 0 && need.length > 0) log('!! sysfonts 为空且系统字体不可读——PDF 转换中文将乱码')
      }
    }
  } catch (eFont) { /* 字体自愈失败不阻断 DS 启动 */ }
  const logFile = path.join(root, '..', 'oo-docservice-log.txt')
  const out = fs.openSync(logFile, 'a')
  const child = spawn(process.execPath, [serverJs], {
    cwd: path.join(root, 'DocService'),
    env: {
      ...process.env,
      // ★ Electron 里 process.execPath=electron.exe，必须 ELECTRON_RUN_AS_NODE=1
      //   才以纯 node 模式跑 server.js（否则 spawn 出来又是一个 Electron GUI）。
      //   纯 node 场景此变量被忽略，无副作用。
      ELECTRON_RUN_AS_NODE: '1',
      NODE_ENV: 'development-windows',
      NODE_CONFIG_DIR: commonCfg,
      SUPPRESS_NO_CONFIG_WARNING: '1',
    },
    stdio: ['ignore', out, out],
    detached: true,
    windowsHide: true,
  })
  child.unref()
  // ★ 2026-10-08「彻底关闭」配套：DS pid 落盘——main.ts app:quit 时读取并终止 DS
  //   （detached 进程不随 Electron 退出；此前点「关闭」后 DS 一直在后台占着 3798）。
  try { fs.writeFileSync(path.join(root, '..', 'oo-docservice.pid'), String(child.pid)) } catch { /* ignore */ }
  log('spawn pid=' + child.pid + ' NODE_ENV=development-windows')
  // 最多等 8s 探活（首次启动要加载 112M node_modules）
  for (let i = 0; i < 16; i++) {
    await new Promise((r) => setTimeout(r, 500))
    if (await isDocServiceAlive()) {
      log('DocService(3798) 端口已通')
      // ★ 2026-10-08 根治「启动窗口期卡死」：DS 是**监听先开、embedded converter 后挂载**——
      //   端口通≠能转文档。此前用户启动后立刻点文档，转换任务进了未挂 runner 的队列，
      //   永远没人消费 → 永远 loading。这里继续等日志出现 'embedded converter started'
      //   才算真正就绪（最多 +12s；老进程跳过启动分支不受影响）。
      let convReady = false
      for (let j = 0; j < 24; j++) {
        await new Promise((r) => setTimeout(r, 500))
        try {
          const tail = fs.readFileSync(logFile, 'utf8').slice(-4096)
          if (tail.includes('embedded converter started')) { convReady = true; break }
        } catch { /* ignore */ }
      }
      if (convReady) log('DocService(3798) 转换器就绪（embedded converter started）')
      else log('!! 12s 内未见 embedded converter started——转换可能滞后，首个文档可能需要稍候')
      return { started: true, alive: true, root, pid: child.pid, convReady }
    }
  }
  log('!! DocService 启动后 8s 内未探活，看日志：' + logFile)
  return { started: true, alive: false, root, pid: child.pid }
}

// ── CLI 直跑 ──
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  ensureDocService().then((r) => {
    process.exit(r.alive ? 0 : 1)
  })
}
