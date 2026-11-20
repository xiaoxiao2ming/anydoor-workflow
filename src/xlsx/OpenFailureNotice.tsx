import { useEffect, useRef, useState } from 'react'
import { formatFailureForReport, type OpenFailure, type OpenFailureSample } from './openError'
import { reportOpenFailure } from '../docx/templateStore'

/**
 * 「打不开这个文件」的统一提示。
 *
 * 三条原则：
 *   ① 不甩堆栈 —— 用户看到的是一句人话 + 一个可执行的下一步
 *   ② 说清是哪一类 —— 用户能自己判断「是文件的问题还是软件的问题」
 *   ③ 留证据 —— 自动落一份失败样本（含原始字节），技术细节可一键复制。
 *      归一化层要补什么，只能靠真实样本回答；没有样本就只能等下次再踩一遍。
 *
 * 样本记录在挂载时自动做一次，用 ref 兜住重复渲染（React 严格模式会跑两遍 effect）。
 */

const KIND_LABEL: Record<OpenFailure['kind'], string> = {
  'legacy-xls': '格式不符',
  encrypted: '已加密',
  'not-zip': '格式不符',
  'schema-strict': '解析器兼容性',
  corrupt: '文件损坏',
  unknown: '未知原因',
}

export function OpenFailureNotice({
  failure,
  name,
  bytes,
  onRetry,
}: {
  failure: OpenFailure
  /** 原始文件名，写进样本里方便回溯 */
  name: string
  /** 原始字节。有就一并留档（离线复盘的关键）；没有也能记录。 */
  bytes: Uint8Array | null
  onRetry?: () => void
}) {
  const [hint, setHint] = useState('')
  const [sample, setSample] = useState<OpenFailureSample | null>(null)
  const [copied, setCopied] = useState(false)
  const reported = useRef('')

  useEffect(() => {
    // 同一个失败只报一次：React 严格模式下 effect 会跑两遍，
    // 而「多留一份样本」比「漏留」更烦人（目录里全是重复副本）
    const key = `${failure.kind}|${failure.detail}|${name}`
    if (reported.current === key) return
    reported.current = key

    let dead = false
    void (async () => {
      const r = await reportOpenFailure({ name, bytes, failure })
      if (dead) return
      setHint(r.hint)
      setSample(r.record)
    })()
    return () => {
      dead = true
    }
  }, [failure, name, bytes])

  const copy = async () => {
    if (!sample) return
    const text = formatFailureForReport(sample)
    try {
      await navigator.clipboard.writeText(text)
      setCopied(true)
      setTimeout(() => setCopied(false), 2000)
    } catch {
      // 剪贴板不可用（权限/非安全上下文）：退化为选中文本让用户手动复制
      const el = document.querySelector('.open-fail-pre')
      if (el) {
        const r = document.createRange()
        r.selectNodeContents(el)
        const sel = window.getSelection()
        sel?.removeAllRanges()
        sel?.addRange(r)
      }
    }
  }

  return (
    <div className="open-fail">
      <div className="open-fail-head">
        <span className="open-fail-kind">{KIND_LABEL[failure.kind]}</span>
        <b>{failure.title}</b>
      </div>

      <div className="open-fail-advice">{failure.advice}</div>

      {failure.applied.length > 0 && (
        <div className="open-fail-applied">
          打开前已应用 {failure.applied.length} 处归一化：
          {failure.applied.map((a, i) => (
            <div key={i}>· {a}</div>
          ))}
        </div>
      )}

      {hint && <div className="open-fail-hint">{hint}</div>}

      <details className="open-fail-detail">
        <summary>技术细节（可复制给开发排查）</summary>
        <pre className="open-fail-pre">{sample ? formatFailureForReport(sample) : failure.detail}</pre>
      </details>

      <div className="open-fail-actions">
        <button className="bar-btn" onClick={() => void copy()}>
          {copied ? '已复制 ✓' : '复制诊断信息'}
        </button>
        {onRetry && (
          <button className="bar-btn" onClick={onRetry}>
            重试
          </button>
        )}
      </div>
    </div>
  )
}
