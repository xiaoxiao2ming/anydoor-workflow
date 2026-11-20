import { useEffect, useRef } from 'react'

/**
 * 极简富文本编辑器（零依赖，contenteditable + execCommand）
 *
 * 选型理由：只需要 粗/斜/下划线/删除线/标题/列表/缩进/颜色/清除格式 这几种，
 * 全部能一对一映射到 OOXML（见 html2ooxml.ts）。引入 Tiptap/Quill 会多几百 KB，
 * 且它们产出的 HTML 里带大量 class 和嵌套 span，反而增加转换器的解析负担。
 *
 * 输出就是 innerHTML —— 与 htmlToOoxml 的输入格式直接对齐。
 */

type Cmd = 'bold' | 'italic' | 'underline' | 'strikeThrough' | 'insertUnorderedList' | 'insertOrderedList'

const BTNS: { cmd: Cmd; label: string; title: string }[] = [
  { cmd: 'bold', label: 'B', title: '加粗' },
  { cmd: 'italic', label: 'I', title: '斜体' },
  { cmd: 'underline', label: 'U', title: '下划线' },
  { cmd: 'strikeThrough', label: 'S', title: '删除线' },
  { cmd: 'insertUnorderedList', label: '•', title: '无序列表' },
  { cmd: 'insertOrderedList', label: '1.', title: '有序列表' },
]

const COLORS = ['', 'C00000', '185FA5', '0F6E56', 'BA7517', '7F7F7F']

export function RichEditor({
  value,
  onChange,
  placeholder,
}: {
  value: string
  onChange: (html: string) => void
  placeholder?: string
}) {
  const ref = useRef<HTMLDivElement>(null)

  // 只在「外部值 ≠ 编辑器内容」且编辑器没被聚焦时同步，避免打字时光标跳到开头
  useEffect(() => {
    const el = ref.current
    if (!el) return
    if (document.activeElement === el) return
    if (el.innerHTML !== value) el.innerHTML = value || ''
  }, [value])

  const exec = (cmd: Cmd | 'removeFormat' | 'formatBlock' | 'foreColor', arg?: string) => {
    ref.current?.focus()
    document.execCommand(cmd, false, arg)
    const html = ref.current?.innerHTML ?? ''
    onChange(html === '<br>' ? '' : html)
  }

  const pickColor = (c: string) => {
    if (!c) return exec('removeFormat')
    exec('foreColor', c)
  }

  return (
    <div className="richedit">
      <div className="richedit-bar">
        {BTNS.map((b) => (
          <button
            key={b.cmd}
            title={b.title}
            className="re-btn"
            // 保持选区不丢：默认的 mousedown 会让编辑器失焦，execCommand 就找不到选区
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => exec(b.cmd)}
          >
            {b.label}
          </button>
        ))}
        <span className="re-sep" />
        <select
          className="re-h"
          defaultValue=""
          onMouseDown={(e) => e.preventDefault()}
          onChange={(e) => {
            exec('formatBlock', e.target.value ? `<${e.target.value}>` : '<p>')
            e.target.value = ''
          }}
        >
          <option value="">正文</option>
          <option value="h2">标题</option>
          <option value="h3">小标题</option>
          <option value="blockquote">引用</option>
        </select>
        <span className="re-sep" />
        {COLORS.map((c) => (
          <button
            key={c || 'auto'}
            title={c ? `颜色 #${c}` : '清除格式'}
            className="re-color"
            style={c ? { background: `#${c}` } : undefined}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => pickColor(c)}
          >
            {c ? '' : '×'}
          </button>
        ))}
      </div>
      <div
        ref={ref}
        className="richedit-body"
        contentEditable
        suppressContentEditableWarning
        data-placeholder={placeholder ?? '输入正文…'}
        onInput={(e) => onChange((e.target as HTMLDivElement).innerHTML)}
        onBlur={(e) => {
          const html = e.currentTarget.innerHTML
          onChange(html === '<br>' ? '' : html)
        }}
      />
    </div>
  )
}
