import { prefGet, prefSet } from './prefs'

/**
 * ★ 画布背景（2026-10-19 用户指令「增加背景图片设置 找一批默认图片 我点开设置可以选择」）：
 *   默认图 8 张（public/bg/*.svg，暗色渐变+微光斑+点阵，与画布 UI 同语言），
 *   选择经 prefs 持久化（桌面=ui-settings.json，跨端口/会话不丢），应用到 .canvas-wrap
 *   （45% 暗罩保证节点对比度）。
 */
export const DEFAULT_BGS: { id: string; name: string; url: string; cat: string }[] = [
  // ── 暗色（原 8 张，与画布 UI 同语言）──
  { id: 'bg1', name: '深空蓝', url: '/bg/bg1.svg', cat: '暗色' },
  { id: 'bg2', name: '暗夜紫', url: '/bg/bg2.svg', cat: '暗色' },
  { id: 'bg3', name: '墨绿', url: '/bg/bg3.svg', cat: '暗色' },
  { id: 'bg4', name: '酒红', url: '/bg/bg4.svg', cat: '暗色' },
  { id: 'bg5', name: '石墨', url: '/bg/bg5.svg', cat: '暗色' },
  { id: 'bg6', name: '深海青', url: '/bg/bg6.svg', cat: '暗色' },
  { id: 'bg7', name: '暖沙', url: '/bg/bg7.svg', cat: '暗色' },
  { id: 'bg8', name: '极光', url: '/bg/bg8.svg', cat: '暗色' },
  // ── 卡通 ──
  { id: 'cartoon1', name: '卡通·雪人', url: '/bg/cartoon1.svg', cat: '卡通' },
  { id: 'cartoon2', name: '卡通·冰淇淋', url: '/bg/cartoon2.svg', cat: '卡通' },
  // ── 商务 ──
  { id: 'business1', name: '商务·折线', url: '/bg/business1.svg', cat: '商务' },
  { id: 'business2', name: '商务·暗纹', url: '/bg/business2.svg', cat: '商务' },
  // ── 清新 ──
  { id: 'fresh1', name: '薄荷', url: '/bg/fresh1.svg', cat: '清新' },
  { id: 'fresh2', name: '云青', url: '/bg/fresh2.svg', cat: '清新' },
  // ── 活力 ──
  { id: 'warm', name: '暖阳', url: '/bg/warm.svg', cat: '活力' },
  { id: 'pastel', name: '活力·日落', url: '/bg/pastel.svg', cat: '活力' },
  // ── 夜空/纸感 ──
  { id: 'starry', name: '星空', url: '/bg/starry.svg', cat: '夜空' },
  { id: 'paper', name: '夜空·月夜山影', url: '/bg/paper.svg', cat: '夜空' },
]

export function readCanvasBg(): string | null {
  const v = prefGet('canvasBg')
  return v && DEFAULT_BGS.some((b) => b.url === v) ? v : null
}

export function applyCanvasBg(url: string | null) {
  /* ★ 2026-10-20 用户「不要因为属性影响背景图」：背景打在 **body**（全窗口固定），
     画布容器 .canvas-wrap 在有背景图时置透明——属性面板显隐/拖宽不再拉伸背景。 */
  const wrap = document.querySelector('.canvas-wrap') as HTMLElement | null
  if (!url) {
    document.body.style.background = ''
    if (wrap) wrap.style.background = ''
    return
  }
  /* ★ 2026-10-22 用户「白色的主题 颜色有点不协调 尤其是选中后」：罩跟主题走——
     暗色主题维持深色氛围罩（原行为，背景图与深卡片同语言）；
     亮色主题换浅雾罩，深色背景图退成淡雅底纹，与白卡片同调不再割裂。
     主题切换后由 Canvas 监听 app-theme-change 重应用本函数。 */
  const light = document.documentElement.dataset.theme === 'light'
  const veil = light ? 'rgba(248, 250, 253, 0.72)' : 'rgba(10, 14, 20, 0.45)'
  document.body.style.background = `linear-gradient(${veil}, ${veil}), url('${url}') center / cover no-repeat fixed`
  if (wrap) wrap.style.background = 'transparent'
}

export function saveCanvasBg(url: string | null) {
  prefSet('canvasBg', url)
  applyCanvasBg(url)
}
