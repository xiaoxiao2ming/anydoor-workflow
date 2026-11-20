# Anydoor

Anydoor 是面向普通办公人群的桌面 C 端工作流应用，专为非技术从业人员打造。

> 工程基线：`client-base`，快照版本：2026-10-22-21-44-20
> 项目归属 WorkBuddy 工作台体系

## ✨ 项目简介

本项目目前上线首个基线版本，仍在持续迭代开发中。

我们希望跳出传统在线 SaaS 的限制，以**离线优先**为核心设计思路，打造一套自主可控、面向未来的个人工作平台。适合需要自建工作流程、沉淀复用知识模板的办公用户，同时支持控件库共建，让能力可以共享、扩展。

## 🎯 核心能力

- **离线优先办公，数据本地自主管理**：数据保存在本机，不依赖云端；断网也能工作。
- **可视化工作流编排**：节点拖拽连线组成流水线，无需编程即可搭建数据处理流程。
- **知识模板自定义、快速复用**：表格、参数、模板域沉淀为可复用资产，一次搭建反复使用。
- **可共建的控件库，扩展能力边界**：节点/控件可持续扩展，能力共享、边界不断生长。

## 🧱 工程说明

目录结构说明（本地开发快照）：

```
├── electron/        # 主进程：文件 IO · 定时任务 · 网页抓取 · 本地大模型推理
├── src/
│   ├── core/        # 领域层：schema · store · 计算引擎 · 意图引擎 · AI · 网址监控
│   ├── office/      # 文档类节点卡片：PDF / OnlyOffice / 网址监控 / Excel转DXF 等
│   ├── ui/          # 通用 UI：按钮装载条 · 图标库 · 运行状态条
│   ├── xlsx/        # Excel 处理与 Excel转DXF 逻辑
│   └── docx/ pptx/  # Word / PPT 文档处理
├── scripts/         # 构建、同步、回归检查等工程脚本
├── docs/            # 技术文档（含技术栈总结）
├── models/          # 本地 GGUF 模型目录（不入库，按需下载）
├── vendors/         # OnlyOffice SDK 等第三方依赖
└── public/          # 静态资源与运行时模型文件
```

### 快速开始

```bash
# 环境要求：Node 22+ · Windows 10/11
npm install

# 开发
npm run dev

# 构建 + 同步打包目录
npm run build && node scripts/sync-to-out.mjs
```

### 技术栈

Electron 44 · React 18 · React Flow v12 · zustand · sql.js（IndexedDB）· node-llama-cpp · cheerio · AG Grid · TipTap · docx · mupdf · pdf-lib · OnlyOffice（ascdesktop）

详见 [docs/技术栈总结.md](docs/技术栈总结.md)。

### 本地大模型（可选）

模型文件（GGUF）放置于 `models/` 目录。不放置模型也可使用——AI 节点支持外部 OpenAI 兼容接口（接口地址留空则走本地模型）。

| 档位 | 建议显存 | 说明 |
|---|---|---|
| 27B-UD-Q3_K_XL | 16GB | 最聪明（heavy 首选） |
| 9B-Q4_K_M | 8GB+ | 平衡 |
| 3B-Q4_K_M | 4GB+ | 快（light 首选） |

## License

[MIT](LICENSE)
