import { defineConfig } from 'vitest/config'

export default defineConfig({
  // SSR 转换缓存固定落在项目内，别让它去 %TEMP% 建随机子目录
  cacheDir: 'node_modules/.vite-cache-test',
  test: {
    environment: 'node',
    include: ['src/**/*.test.ts'],
    /* ★ 必须单进程跑（2026-10-05 实测）
     *
     * 现象：多 worker 并发时 Vite 写临时目录偶发被拦：
     *   EPERM: operation not permitted, open '%TEMP%\<随机串>\ssr\<hash>'
     *   → vitest 当 unhandled error 抛出并**中途终止整轮**
     *   → 表现为「29/30 文件、463/544 用例」，exit code 非 0。
     *
     * 为什么这个症状值得专门写一段：
     *   它长得极像「有用例挂了」，会把人引去查错方向（改代码/改用例），
     *   而实际是工具链中途崩了。实测 4 次里 3 次踩中，且**每次缺的文件不同**，
     *   更像随机失败。唯一可靠的判据是「Test Files 少于 include 匹配到的文件数」。
     *
     * 根因：沙箱的 fs 代理在并发写同一批临时路径时有竞态，不是 vitest 缺陷。
     * 单进程（3/3 稳定 30 文件零 EPERM）代价是慢几秒，换来结果可信，值得。
     * 换到不受限的环境后可以摘掉。 */
    pool: 'forks',
    poolOptions: { forks: { singleFork: true } },
  },
})
