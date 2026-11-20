/**
 * 版本号单一出口 —— 根目录 VERSION 是唯一真相源。
 *
 * 为什么单独一个文件：桌面端（vite.config.ts）、网页端（vite.web.config.ts）、
 * 构建脚本（build-web.mjs）、服务端（app-server.mjs 的 /version）都要报同一个版本，
 * 任何一处硬写都会「两边对不上」。全部从这里读。
 *
 * 约定（见 release.mjs / CHANGELOG.md）：三位 X.Y.Z，次版本与修订号补两位（0.00.15）。
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')

/** 读根目录 VERSION；读不到返回 '0.0.0' 而不是抛错（构建不能因版本文件缺失而挂） */
export function readVersion() {
  try {
    return readFileSync(join(ROOT, 'VERSION'), 'utf8').trim() || '0.0.0'
  } catch {
    return '0.0.0'
  }
}

/** 构建时间戳（ISO，秒级）——网页端「这份产物是什么时候出的」 */
export function buildStamp() {
  return new Date().toISOString().replace(/\.\d+Z$/, 'Z')
}

/** 给 Vite 的 define 用：版本 + 渠道（desktop / web） */
export function versionDefine(channel) {
  return {
    __APP_VERSION__: JSON.stringify(readVersion()),
    __APP_CHANNEL__: JSON.stringify(channel),
    __APP_BUILD__: JSON.stringify(buildStamp()),
  }
}
