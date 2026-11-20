# 午马未央 · Canvas Workflow

画布式 AI 工作流桌面应用——节点拖拽连线组成流水线，本地大模型离线推理，覆盖网页监控、数据提取、文档生成、PDF 处理与 AI 分析全链路。

![screenshot](docs/screenshot.png)

## 特性

- **27 种节点**：数据表 / 参数表 / 网址监控 / Excel转DXF / PDF 生成·瘦身·替换·转Word / Word·Excel 模板域回写 / OnlyOffice 内嵌编辑 / AI 分析等
- **对话式搭建**：对话框一句话（如「搭一个网址监控工作流」）自动创建节点链并连线、预填配置
- **本地大模型**：node-llama-cpp + GGUF 离线推理，会话池多模型常驻，任务分层路由（重任务主力模型 / 轻任务秒回）；也支持外部 OpenAI 兼容接口
- **网址监控**：轮询抓取（CSS 选择器 / JSON 路径 / 关键词）→ 变化检测 → 数据自动写入下游 → 变化触发 AI 分析
- **文档能力**：OnlyOffice 桌面内核内嵌、模板域（w:sdt）回写生成 Word、Excel 命名区域绑定
- **引用体系**：`{{field}}` 内联引用贯通表格、模板与画布

## 技术栈

Electron 44 · React 18 · React Flow v12 · zustand · sql.js（IndexedDB）· node-llama-cpp · cheerio · AG Grid · TipTap · docx · mupdf · pdf-lib · OnlyOffice（ascdesktop）

详见 [docs/技术栈总结.md](docs/技术栈总结.md)。

## 快速开始

```bash
# 环境要求：Node 22+ · Windows 10/11
npm install

# 开发
npm run dev

# 构建 + 同步打包目录
npm run build && node scripts/sync-to-out.mjs
```

### 本地大模型

模型文件（GGUF）放置于 `models/` 目录，例如 Qwen3.5-9B：

```bash
python -m pip install modelscope
python -c "from modelscope import snapshot_download; snapshot_download('Qwen/Qwen3.5-9B-GGUF', allow_file_pattern=['*q4_k_m*.gguf'], local_dir='models/')"
```

| 档位 | 建议显存 | 说明 |
|---|---|---|
| 27B-UD-Q3_K_XL | 16GB | 最聪明（heavy 首选） |
| 9B-Q4_K_M | 8GB+ | 平衡 |
| 3B-Q4_K_M | 4GB+ | 快（light 首选） |

不放置模型也可使用——AI 节点支持外部 OpenAI 兼容接口（接口地址留空则走本地模型）。

## 目录结构

```
├── electron/        # 主进程（模型推理 / 文件 IO / 定时任务 / 网页抓取）
├── src/core/        # 领域层：schema · store · engine · 意图引擎 · AI · 网址监控
├── src/office/      # 节点卡片（PDF / OnlyOffice / 监控 / DXF …）
├── src/ui/          # 通用 UI（按钮装载条 / 图标库 / 运行状态）
├── docs/            # 技术文档
└── models/          # 本地 GGUF 模型（不入库，见上方下载说明）
```

## License

[MIT](LICENSE)
