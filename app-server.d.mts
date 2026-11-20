// app-server.mjs 的类型声明（供 electron/main.ts 与 vite.config.ts 以 TS 方式 import）
export interface AppServerOptions {
  /** 静态资源根目录；null 表示只提供 /api（静态请求返回 404） */
  root?: string | null
  /** 监听端口；0 表示由系统分配空闲端口 */
  port?: number
  /** 监听地址；默认 127.0.0.1（本机），跨设备接入传 0.0.0.0 */
  host?: string
  /** 页面数据目录，默认项目内 server-data/pages */
  dataDir?: string
  /** 日志前缀 */
  label?: string
}

export interface AppServerHandle {
  /** 实际监听端口（port 传 0 时为系统分配值） */
  port: number
  /** 当前生效的数据目录 */
  dataDir: string
  /** 关闭服务 */
  close: () => Promise<void>
}

export declare function startAppServer(opts?: AppServerOptions): Promise<AppServerHandle>
