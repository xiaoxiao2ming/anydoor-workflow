/// <reference types="vite/client" />

/* 构建期由 Vite define 注入（见 scripts/app-version.mjs 的 versionDefine）
 *   __APP_VERSION__ 根目录 VERSION 的内容，桌面端与网页端同源
 *   __APP_CHANNEL__ 'desktop' | 'web'
 *   __APP_BUILD__   构建时间戳（ISO） */
declare const __APP_VERSION__: string
declare const __APP_CHANNEL__: 'desktop' | 'web'
declare const __APP_BUILD__: string
