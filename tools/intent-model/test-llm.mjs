// Qwen2.5-0.5B 本地推理验证（node-llama-cpp）
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { getLlama, LlamaChatSession } from 'node-llama-cpp'
const dir = path.dirname(fileURLToPath(import.meta.url))

console.log('[1/3] 初始化 llama（CPU）…')
const t0 = Date.now()
const llama = await getLlama()
const model = await llama.loadModel({ modelPath: path.join(dir, '../../models/qwen2.5-0.5b-instruct-q4_k_m.gguf') })
const context = await model.createContext()
const session = new LlamaChatSession({ contextSequence: context.getSequence() })
console.log('      就绪，耗时', ((Date.now() - t0) / 1000).toFixed(1), 's')

console.log('[2/3] 对话测试…')
const a1 = await session.prompt('你是什么模型？用一句话回答。')
console.log('  Q: 你是什么模型？\n  A:', a1.slice(0, 100))

console.log('[3/3] 工作流意图理解测试…')
const a2 = await session.prompt('用户想把一张截图变成可编辑的 Word 文档。请只输出 JSON：{"intent":"image-to-word","needNodes":["图片导入","Word输出"]}，不要其他内容。')
console.log('  Q: 截图→可编辑文档\n  A:', a2.slice(0, 160))
process.exit(0)
