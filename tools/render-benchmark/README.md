# 渲染兼容性基准集（目标：自研渲染尽可能接近 WPS）

## 用法（三步）

1. **放文档**：把典型业务 docx/xlsx 放进 `docs/`
2. **做标准答案**：每份用**桌面 WPS** 打开 → 另存为 PDF → 放进 `golden/`（**同名**，如 `报告.docx` → `报告.pdf`）
3. **放自研结果**：用我们渲染导出的 PDF/PNG 放进 `mine/`（同名）

## 对比

```bash
cd tools/render-benchmark
npm i pdfjs-dist pixelmatch pngjs   # 首次
node compare.mjs                    # 全量对比
node compare.mjs 报告                # 只对比文件名含"报告"的
```

输出：每份文档的差异百分比 + 差异热力图（`reports/xxx.diff.png`，红处=与 WPS 不一致）

## 判读标准

| 差异 | 含义 |
|---|---|
| < 2% | ✅ 优秀（肉眼几乎一致） |
| 2%~5% | 🟡 可用（细节有差） |
| > 5% | 🔴 需要修渲染规则 |

## 铁律

**改渲染代码前跑一次、改完再跑一次**——修好了 A、B 的差异却变大 = 修偏了，禁止合入。
这就是「不修坏」的机械化保障。
