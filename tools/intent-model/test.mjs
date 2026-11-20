// 本地语义匹配验证脚本（bge-small-zh-v1.5 量化版 24MB，本地离线加载）
// 用法：node tools/intent-model/test.mjs
import path from 'node:path'
import { fileURLToPath } from 'node:url'
const MODEL_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), 'model')

const { pipeline, cos_sim } = await import('@huggingface/transformers')

console.log('[1/3] 加载本地模型（tools/intent-model/model，离线）…')
const t0 = Date.now()
const extractor = await pipeline('feature-extraction', MODEL_DIR, { dtype: 'q8' })
console.log('      模型就绪，耗时', ((Date.now() - t0) / 1000).toFixed(1), 's')

console.log('[2/3] 构建意图向量…')
const intents = [
  { id: 'image-to-word', desc: '把图片转成 word 文档，图片文字识别转 docx' },
  { id: 'pdf-merge', desc: '把多份 PDF 合并成一份文件' },
  { id: 'xlsx-merge', desc: '把多个 Excel 表格合并汇总到一个表' },
]
const iOut = await extractor(intents.map((x) => x.desc), { pooling: 'mean', normalize: true })
const iMat = iOut.tolist()

console.log('[3/3] 语义匹配测试（用户换了说法也能命中）…')
const tests = [
  '帮我把这个截图弄成能编辑的文档',   // 无关键词，纯语义 → 应命中 image-to-word
  '几张扫描件合成一个文件',           // → pdf-merge
  '打开画布',                          // → 低分（无意图）
]
const uOut = await extractor(tests, { pooling: 'mean', normalize: true })
const uMat = uOut.tolist()
for (let i = 0; i < tests.length; i++) {
  const scores = intents.map((it, j) => ({ id: it.id, s: cos_sim(uMat[i], iMat[j]) })).sort((a, b) => b.s - a.s)
  const top = scores[0]
  const verdict = top.s > 0.55 ? '✅' : '⬇️ 低置信（走规则兜底）'
  console.log(`  「${tests[i]}」→ ${verdict} ${top.id} (${top.s.toFixed(3)})`)
}
console.log('完成 ✔')
process.exit(0)
