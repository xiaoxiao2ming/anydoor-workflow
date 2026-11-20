#!/usr/bin/env node
/**
 * 版本发布脚本 —— client-base 版本约定（用户 2026-09-26 拍板：版本从 0.00.01 开始）
 *
 * 每完成一批交付（功能/修复验收全绿后）跑一次：
 *   node scripts/release.mjs          # patch 位 +1（0.00.01 → 0.00.02）
 *   node scripts/release.mjs 0.01.00  # 显式指定版本号
 *
 * 动作：读 VERSION → 计算新版本 → 写回 VERSION + 同步 package.json.version
 *      → git commit「版本 X.Y.Z」→ git tag vX.Y.Z
 * 前提：工作区代码改动已先行提交（本脚本只提交版本文件本身）。
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { execSync } from 'node:child_process'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const vfile = join(root, 'VERSION')
const cur = readFileSync(vfile, 'utf8').trim()

let next = process.argv[2]
if (!next) {
  const parts = cur.split('.').map(Number)
  if (parts.length !== 3 || parts.some((n) => !Number.isFinite(n))) {
    console.error(`VERSION 内容不合法：${JSON.stringify(cur)}（应为 X.Y.Z）`)
    process.exit(1)
  }
  // 前导零约定（0.00.01）：次版本/修订号补两位，主版本号不补位（0.00.02 → 0.00.03，不是 00.00.03）
  const pad2 = (n) => String(n).padStart(2, '0')
  next = `${parts[0]}.${pad2(parts[1])}.${pad2(parts[2] + 1)}`
}
if (!/^\d+\.\d+\.\d+$/.test(next)) {
  console.error(`版本号不合法：${next}（应为 X.Y.Z）`)
  process.exit(1)
}

writeFileSync(vfile, next + '\n')
const pkgPath = join(root, 'package.json')
const pkg = JSON.parse(readFileSync(pkgPath, 'utf8'))
pkg.version = next
writeFileSync(pkgPath, JSON.stringify(pkg, null, 2) + '\n')

execSync(`git add VERSION package.json && git commit -m "版本 ${next}" && git tag v${next}`, {
  cwd: root,
  stdio: 'inherit',
})
console.log(`✅ 已发布 v${next}（上一版 v${cur}）`)
