// oo-docservice.mjs 的类型声明（供 electron/main.ts 以 TS 方式 import）
export interface DocServiceEnsureResult {
  /** 本次是否实际 spawn 了新进程（false = 已在跑或启动失败） */
  started: boolean
  /** 返回时 3798 是否可连 */
  alive: boolean
  /** 解析到的 vendors/oo-server 根目录（未找到为 null） */
  root: string | null
  /** spawn 出的子进程 pid（仅 started=true 时可能存在） */
  pid?: number
}

export declare function ensureDocService(opts?: { quiet?: boolean }): Promise<DocServiceEnsureResult>
export declare function isDocServiceAlive(port?: number, timeoutMs?: number): Promise<boolean>
export declare function resolveOoServerRoot(startDir?: string): string | null
