import { useEffect, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'

/**
 * 所见即所得预览：用 docx-preview 在浏览器里真实分页渲染。
 *
 * 关键点：传进来的必须是「已绑定且已剥壳」的产物字节 —— 与导出给用户的是同一份，
 * 这才是「预览即终稿」。行级控件如果不剥壳，docx-preview 会把数据行整段吞掉。
 */
export function DocxPreview({ data, className }: { data: Uint8Array | null; className?: string }) {
  const bodyRef = useRef<HTMLDivElement>(null)
  const styleRef = useRef<HTMLDivElement>(null)
  const [err, setErr] = useState<string>('')

  useEffect(() => {
    const body = bodyRef.current
    const style = styleRef.current
    if (!body || !style) return
    body.innerHTML = ''
    style.innerHTML = ''
    if (!data || !data.byteLength) return

    let alive = true
    setErr('')
    renderAsync(data as unknown as Blob, body, style, {
      className: 'docx-render',
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreFonts: false,
      breakPages: true,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      useBase64URL: true,
      trimXmlDeclaration: true,
      debug: false,
      experimental: false,
    }).catch((e: unknown) => {
      if (alive) setErr(e instanceof Error ? e.message : String(e))
    })

    return () => {
      alive = false
    }
  }, [data])

  return (
    <div className={className}>
      <div ref={styleRef} className="docx-style-host" />
      {err && <div className="docx-preview-err">预览失败：{err}</div>}
      {!data && !err && <div className="docx-preview-empty">选择模板并绑定数据后，这里显示最终排版效果</div>}
      <div ref={bodyRef} className="docx-preview-body" />
    </div>
  )
}
