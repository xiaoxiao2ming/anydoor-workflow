/** scripts/app-version.mjs 的类型声明（vite.config.ts 是 TS，导入 .mjs 需要声明） */
export declare function readVersion(): string
export declare function buildStamp(): string
export declare function versionDefine(channel: 'desktop' | 'web'): Record<string, string>
