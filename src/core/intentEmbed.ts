/* ★ 本地语义意图匹配（2026-11-01 部署的 bge-small-zh-v1.5 量化版 24MB）
 * 渲染进程懒加载：首次调用才加载模型（约 24MB，本地静态资源，完全离线）
 * 用法：const intent = await matchByEmbedding('帮我把截图弄成能编辑的文档')
 * 命中置信 <0.55 返回 null（调用方走规则兜底/引导） */
import type { ChatIntent } from './chatIntent'

let extractor: ((texts: string[], opts: Record<string, unknown>) => Promise<{ tolist: () => number[][] }>) | null = null
let loading: Promise<void> | null = null
let intentVecs: number[][] | null = null

/* 意图描述（与 CanvasChatDock INTENTS 同步维护——id 对应） */
const INTENT_DESCS: Array<{ id: string; desc: string }> = [
  { id: 'analyze-node', desc: '分析画布上的节点内容和文档，总结一下画布里有什么' },
  { id: 'image-to-word', desc: '把图片转成 word 文档，图片文字识别转 docx，截图变可编辑文件' },
  { id: 'word-to-pdf', desc: '把 word 文档转成 pdf，docx 导出为 PDF 文件' },
  { id: 'pdf-merge', desc: '把多份 PDF 合并成一份文件，扫描件合成' },
  { id: 'xlsx-merge', desc: '把多个 Excel 表格合并汇总到一个表' },
]
/** CanvasChatDock INTENTS 的 id → 描述 id 映射后回传用（返回 id 原样） */

function cosSim(a: number[], b: number[]): number {
  let dot = 0
  const n = Math.min(a.length, b.length)
  for (let i = 0; i < n; i++) dot += a[i] * b[i]
  return dot /* 已归一化向量，点积=余弦 */
}

async function ensureModel(): Promise<void> {
  if (extractor) return
  if (loading) return loading
  loading = (async () => {
    /* 动态 import：不用意图功能就完全不加载（bundle 懒切分） */
    const { pipeline, env } = await import('@huggingface/transformers')
    env.allowLocalModels = true
    env.allowRemoteModels = false /* 离线：模型与 wasm 全走本地静态资源 */
    env.backends.onnx!.wasm!.wasmPaths = '/models/ort/' /* ★ wasm 本地化（默认 jsdelivr CDN 离线拉不到，实测 ERR_CONNECTION_RESET） */
    extractor = (await pipeline('feature-extraction', '/models/intent-bge', { dtype: 'q8' })) as never
  })()
  return loading
}

/** 意图描述向量（首次调用时构建） */
async function ensureIntentVecs(): Promise<number[][]> {
  if (intentVecs) return intentVecs
  await ensureModel()
  const out = await extractor!(INTENT_DESCS.map((x) => x.desc), { pooling: 'mean', normalize: true })
  intentVecs = out.tolist()
  return intentVecs
}

/** 语义匹配：始终返回最高分（id=命中意图，ok=是否达到自动执行阈值）
 *  ok=false 且 score 较高时调用方可给出「你是不是想…」猜测提示 */
export async function matchByEmbedding(text: string): Promise<{ id: string; score: number; ok: boolean }> {
  try {
    const vecs = await ensureIntentVecs()
    const out = await extractor!([text], { pooling: 'mean', normalize: true })
    const u = out.tolist()[0]
    let best = { id: INTENT_DESCS[0].id, s: -1 }
    INTENT_DESCS.forEach((it, j) => {
      const s = cosSim(u, vecs[j])
      if (s > best.s) best = { id: it.id, s }
    })
    return { id: best.id, score: best.s, ok: best.s >= 0.5 }
  } catch (e) {
    console.warn('[意图语义] 匹配失败，走兜底：', e)
    return { id: '', score: 0, ok: false }
  }
}
