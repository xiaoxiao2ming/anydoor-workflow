import type { ReactNode } from 'react'

/**
 * ★ 2026-10-19 用户「重新做按钮的logo / 还是有点丑」：
 *   手绘 16 格 SVG 换成 Lucide（ISC 许可）24 格标准网格路径，内联嵌入零依赖。
 *   统一 stroke=2 / 圆头圆角，与 lucide 官方渲染一致。
 *   size 显式传（标题栏 19 / 节点卡 19），不依赖 CSS 推断。
 */
export function LIcon({ size = 19, children }: { size?: number; children: ReactNode }) {
  return (
    <svg
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={2}
      strokeLinecap="round"
      strokeLinejoin="round"
      style={{ width: size, height: size, display: 'block' }}
      aria-hidden="true"
    >
      {children}
    </svg>
  )
}

/* 打开文件（folder-open） */
export const IcFolderOpen = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="m6 14 1.5-2.9A2 2 0 0 1 9.24 10H20a2 2 0 0 1 1.94 2.5l-1.54 6a2 2 0 0 1-1.95 1.5H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h3.9a2 2 0 0 1 1.69.9l.81 1.2a2 2 0 0 0 1.67.9H18a2 2 0 0 1 2 2v2" /></LIcon>
)

/* 保存（database 圆柱） */
export const IcDatabase = ({ size }: { size?: number }) => (
  <LIcon size={size}><ellipse cx="12" cy="5" rx="9" ry="3" /><path d="M3 5V19A9 3 0 0 0 21 19V5" /><path d="M3 12A9 3 0 0 0 21 12" /></LIcon>
)

/* 导出（download） */
export const IcDownload = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" x2="12" y1="15" y2="3" /></LIcon>
)

/* 打开历史（clock） */
export const IcHistory = ({ size }: { size?: number }) => (
  <LIcon size={size}><circle cx="12" cy="12" r="10" /><polyline points="12 6 12 12 16 14" /></LIcon>
)

/* 左栏 / 右栏（panel-left / panel-right） */
export const IcPanelLeft = ({ size }: { size?: number }) => (
  <LIcon size={size}><rect width="18" height="18" x="3" y="3" rx="2" /><path d="M9 3v18" /></LIcon>
)
export const IcPanelRight = ({ size }: { size?: number }) => (
  <LIcon size={size}><rect width="18" height="18" x="3" y="3" rx="2" /><path d="M15 3v18" /></LIcon>
)

/* 诊断（info） */
export const IcInfo = ({ size }: { size?: number }) => (
  <LIcon size={size}><circle cx="12" cy="12" r="10" /><path d="M12 16v-4" /><path d="M12 8h.01" /></LIcon>
)

/* 引用格着色（square + 中心点；off 叠斜杠） */
export const IcMark = ({ size, off }: { size?: number; off?: boolean }) => (
  <LIcon size={size}>
    <rect x="3" y="3" width="18" height="18" rx="3" />
    <rect x="9" y="9" width="6" height="6" rx="1" fill="currentColor" stroke="none" />
    {off && <path d="M5 19 19 5" />}
  </LIcon>
)

/* 高亮行列（grid-3x3） */
export const IcGrid = ({ size }: { size?: number }) => (
  <LIcon size={size}><rect x="3" y="3" width="18" height="18" rx="2" /><path d="M3 9h18M3 15h18M9 3v18M15 3v18" /></LIcon>
)

/* 引用更新（refresh-cw） */
export const IcRefresh = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8" /><path d="M21 3v5h-5" /><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16" /><path d="M8 16H3v5" /></LIcon>
)

/* 不检测（ban） */
export const IcBan = ({ size }: { size?: number }) => (
  <LIcon size={size}><circle cx="12" cy="12" r="10" /><path d="m4.9 4.9 14.2 14.2" /></LIcon>
)

/* 视觉对比（git-compare——PDF 替换页历史 diff 按钮，2026-10-25） */
export const IcGitCompare = ({ size }: { size?: number }) => (
  <LIcon size={size}><circle cx="18" cy="18" r="3" /><circle cx="6" cy="6" r="3" /><path d="M13 6h3a2 2 0 0 1 2 2v7" /><path d="M11 18H8a2 2 0 0 1-2-2V9" /></LIcon>
)

/* 漏斗过滤（filter——容差/筛选类设置，2026-10-25） */
export const IcFilter = ({ size }: { size?: number }) => (
  <LIcon size={size}><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3" /></LIcon>
)

/* 主题切换（sun / moon） */
export const IcSun = ({ size }: { size?: number }) => (
  <LIcon size={size}><circle cx="12" cy="12" r="4" /><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41" /></LIcon>
)
export const IcMoon = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z" /></LIcon>
)

/* 最大化 / 还原（maximize-2 / minimize-2） */
export const IcMaximize = ({ size }: { size?: number }) => (
  <LIcon size={size}><polyline points="15 3 21 3 21 9" /><polyline points="9 21 3 21 3 15" /><line x1="21" x2="14" y1="3" y2="10" /><line x1="3" x2="10" y1="21" y2="14" /></LIcon>
)
export const IcMinimize = ({ size }: { size?: number }) => (
  <LIcon size={size}><polyline points="4 14 10 14 10 20" /><polyline points="20 10 14 10 14 4" /><line x1="14" x2="21" y1="10" y2="3" /><line x1="3" x2="10" y1="21" y2="14" /></LIcon>
)

/* 关闭（x） */
export const IcX = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M18 6 6 18" /><path d="m6 6 12 12" /></LIcon>
)

/* 预加载（zap） */
export const IcZap = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M13 2 3 14h9l-1 8 10-12h-9l1-8z" /></LIcon>
)

/* 数据锁（lock / unlock） */
export const IcLock = ({ size }: { size?: number }) => (
  <LIcon size={size}><rect width="18" height="11" x="3" y="11" rx="2" ry="2" /><path d="M7 11V7a5 5 0 0 1 10 0v4" /></LIcon>
)
export const IcUnlock = ({ size }: { size?: number }) => (
  <LIcon size={size}><rect width="18" height="11" x="3" y="11" rx="2" ry="2" /><path d="M7 11V7a5 5 0 0 1 9.9-1" /></LIcon>
)

/* 位置锁（map-pin） */
export const IcMapPin = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0Z" /><circle cx="12" cy="10" r="3" /></LIcon>
)

/* 重启（power） */
export const IcPower = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M12 2v10" /><path d="M18.4 6.6a9 9 0 1 1-12.77.04" /></LIcon>
)

/* 套件中心（puzzle） */
export const IcPuzzle = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M19.439 7.85c-.049.322.059.648.289.878l1.568 1.568c.47.47.706 1.087.706 1.704s-.235 1.233-.706 1.704l-1.611 1.611a.98.98 0 0 1-.837.276c-.47-.07-.802-.48-.968-.925a2.501 2.501 0 1 0-3.214 3.214c.446.166.855.497.925.968a.979.979 0 0 1-.276.837l-1.61 1.61a2.404 2.404 0 0 1-1.705.707 2.402 2.402 0 0 1-1.704-.706l-1.568-1.568a1.026 1.026 0 0 0-.877-.29c-.493.074-.84.504-1.02.968a2.5 2.5 0 1 1-3.237-3.237c.464-.18.894-.527.967-1.02a1.026 1.026 0 0 0-.289-.877l-1.568-1.568A2.402 2.402 0 0 1 1.998 12c0-.617.236-1.234.706-1.704L4.23 8.77c.24-.24.581-.353.917-.303.515.077.877.528 1.073 1.01a2.5 2.5 0 1 0 3.259-3.259c-.482-.196-.933-.558-1.01-1.073a1.026 1.026 0 0 1 .303-.917l1.525-1.525A2.402 2.402 0 0 1 12 1.998c.617 0 1.234.236 1.704.706l1.568 1.568c.23.23.556.338.877.29.493-.074.84-.504 1.02-.968a2.5 2.5 0 1 1 3.237 3.237c-.464.18-.894.527-.967 1.02Z" /></LIcon>
)

/* 重启（power） */

/* 重启备选（rotate-ccw 逆时针回转箭头） */
export const IcRotateCcw = ({ size }: { size?: number }) => (
  <LIcon size={size}><polyline points="1 4 1 10 7 10" /><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10" /></LIcon>
)

/* 背景（image） */
export const IcImage = ({ size }: { size?: number }) => (
  <LIcon size={size}><rect width="18" height="18" x="3" y="3" rx="2" ry="2" /><circle cx="9" cy="9" r="2" /><path d="m21 15-3.086-3.086a2 2 0 0 0-2.828 0L6 21" /></LIcon>
)

/* 回收站（trash2） */
export const IcTrash = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M3 6h18" /><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6" /><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" /><line x1="10" x2="10" y1="11" y2="17" /><line x1="14" x2="14" y1="11" y2="17" /></LIcon>
)

/* 收拢（shrink） */
export const IcShrink = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M15 15h6V9" /><path d="M9 9H3v6" /><path d="m21 21-6-6" /><path d="m3 3 6 6" /></LIcon>
)

/* 折叠成员（fold-vertical 简化：双箭头对折 + 中线） */
export const IcFold = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M2 12h20" /><polyline points="16 4 12 8 8 4" /><polyline points="8 20 12 16 16 20" /></LIcon>
)

/* 磁吸（magnet） */
export const IcMagnet = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="m6 15-4-4 6.75-6.77a7.79 7.79 0 0 1 11 11L13 22l-4-4 6.39-6.36a2.14 2.14 0 0 0-3-3L6 15" /><path d="m5 8 4 4" /><path d="m12 15 4 4" /></LIcon>
)

/* 文件（file） */
export const IcFile = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z" /><path d="M14 2v5h5" /></LIcon>
)

/* 另存为（save-as） */
export const IcSaveAs = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M15.2 3a2 2 0 0 1 1.4.6l3.8 3.8a2 2 0 0 1 .6 1.4V19a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z" /><path d="M17 2v4h4" /><path d="M12 12v3" /><path d="m8 17 4-4 4 4" /></LIcon>
)
/* 清空（eraser） */
export const IcClear = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="m7 21-4.3-4.3c-1-1-1-2.5 0-3.4l9.6-9.6c1-1 2.5-1 3.4 0l5.6 5.6c1 1 1 2.5 0 3.4L13 21" /><path d="M22 21H7" /><path d="m5 11 9 9" /></LIcon>
)

/* 相机（照片导入，2026-10-23 用户指令「OO 系列按钮 logo 换 SVG 线条」） */
export const IcCamera = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M14.5 4h-5L7 7H4a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-3l-2.5-3z" /><circle cx="12" cy="13" r="3" /></LIcon>
)
/* 地球/远程访问（2026-10-24 用户指令「右上角远程访问设置」）——lucide globe 同款 */
export const IcGlobe = ({ size }: { size?: number }) => (
  <LIcon size={size}><circle cx="12" cy="12" r="10" /><path d="M12 2a14.5 14.5 0 0 0 0 20 14.5 14.5 0 0 0 0-20" /><path d="M2 12h20" /></LIcon>
)

/* 移动/拖动显示（2026-10-24 用户指令「右上角按钮 点击选择」）——lucide move 同款 */
export const IcMove = ({ size }: { size?: number }) => (
  <LIcon size={size}><polyline points="5 9 2 12 5 15" /><polyline points="9 5 12 2 15 5" /><polyline points="15 19 12 22 9 19" /><polyline points="19 9 22 12 19 15" /><line x1="2" x2="22" y1="12" y2="12" /><line x1="12" x2="12" y1="2" y2="22" /></LIcon>
)

/* 芯片/硬件加速（2026-10-24 用户指令「硬件加速开关」）——lucide cpu 同款 */
export const IcChip = ({ size }: { size?: number }) => (
  <LIcon size={size}><rect x="4" y="4" width="16" height="16" rx="2" /><rect x="9" y="9" width="6" height="6" /><path d="M15 2v2" /><path d="M15 20v2" /><path d="M2 15h2" /><path d="M2 9h2" /><path d="M20 15h2" /><path d="M20 9h2" /><path d="M9 2v2" /><path d="M9 20v2" /></LIcon>
)

/* 文档（PDF/Word 导入） */
export const IcDoc = ({ size }: { size?: number }) => (
  <LIcon size={size}><path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z" /><path d="M14 2v4a2 2 0 0 0 2 2h4" /></LIcon>
)
