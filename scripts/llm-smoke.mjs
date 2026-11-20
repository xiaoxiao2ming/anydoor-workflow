#!/usr/bin/env node
/**
 * 模型加载冒烟测试（2026-11-07 升级 Qwen3.5-9B 前的兼容性验证）
 *
 * 用法：node scripts/llm-smoke.mjs [模型文件名]
 *   node scripts/llm-smoke.mjs models/Qwen3.5-9B-Q4_K_M.gguf
 *   node scripts/llm-smoke.mjs models/qwen2.5-3b-instruct-q4_k_m.gguf
 *
 * 验证三件事：
 *   1. 架构能不能被 node-llama-cpp 内置的 llama.cpp 认出来（Qwen3.5 是新架构 qwen35）
 *   2. 一次真实 prompt 能不能出字（对话模板 / tokenizer 是否工作）
 *   3. 速度多少 tok/s（决定要不要试 Vulkan）
 *
 * 为什么要先测：node-llama-cpp 3.22.3 打包的 llama.cpp 如果早于 Qwen3.5（2026-03）发布，
 * loadModel 会直接抛「unknown architecture」——那 6GB 就白下了。
 */
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import fs from 'node:fs'

const arg = process.argv[2] || 'models/Qwen3.5-9B-Q4_K_M.gguf'
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const modelPath = path.isAbsolute(arg) ? arg : path.join(ROOT, arg)

if (!fs.existsSync(modelPath)) {
  console.error(`模型文件不存在：${modelPath}`)
  process.exit(1)
}
console.log(`模型：${modelPath}（${(fs.statSync(modelPath).size / 1e9).toFixed(2)} GB）`)

const { getLlama, LlamaChatSession } = await import('node-llama-cpp')

const t0 = Date.now()
const llama = await getLlama()
console.log(`llama 运行时就绪（${((Date.now() - t0) / 1000).toFixed(1)}s）`)

const t1 = Date.now()
const model = await llama.loadModel({ modelPath })
console.log(`模型加载完成（${((Date.now() - t1) / 1000).toFixed(1)}s）`)

const t2 = Date.now()
const context = await model.createContext()
const session = new LlamaChatSession({ contextSequence: context.getSequence() })
console.log(`上下文就绪（${((Date.now() - t2) / 1000).toFixed(1)}s），开始对话测试…`)

const t3 = Date.now()
const reply = await session.prompt('用一句话回答：你是什么模型？', { maxTokens: 96 })
const dt = (Date.now() - t3) / 1000
console.log(`回复（${dt.toFixed(1)}s）：${reply}`)

/* 结束要显式 dispose，不然进程挂着不退 */
await model.dispose()
console.log('✅ 冒烟测试通过')
process.exit(0)
