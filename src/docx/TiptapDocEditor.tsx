/**
 * Tiptap 富文本编辑器内核（ProseMirror）
 *
 * 定位：替代 RichEditor.tsx（contenteditable + execCommand，MDN 已废弃方案）。
 * 为什么换：
 *   - execCommand 的格式能力残缺、浏览器行为不一致、选区/光标有怪癖，撤销栈不可靠
 *   - Tiptap 是文档模型（schema）驱动：格式操作走 transaction，状态可预测、undo 可靠
 *   - 格式能力开箱即用：字号精确 pt / 字体族 / 任意色高亮 / 行距 / 对齐 / 多级列表 / 表格
 *
 * 与既有 OOXML 转换层的关系（重要）：
 *   本组件只负责「HTML ↔ 编辑体验」，不做 docx 读写。
 *   进：外部 HTML（来自 ooxml2html.docxBodyToHtml）→ getHTML() 输出规整 HTML
 *   出：onUpdate(html) → 交给 html2ooxml 写回 docx
 *   因此 html2ooxml 需能吃 Tiptap 的 HTML 形态（<mark> 高亮、style="text-align" 等）。
 *
 * 样式保真：Tiptap 的 TextStyle/Color/Highlight 等扩展产出内联 style，
 * 尽量避免 class —— 转换器按 style 解析，class 会在导出时丢失语义。
 */

import { useEffect, useRef } from 'react'
import { useEditor, EditorContent, type Editor } from '@tiptap/react'
import StarterKit from '@tiptap/starter-kit'
import TextAlign from '@tiptap/extension-text-align'
import TextStyle from '@tiptap/extension-text-style'
import Color from '@tiptap/extension-color'
import Highlight from '@tiptap/extension-highlight'
import Image from '@tiptap/extension-image'
import Subscript from '@tiptap/extension-subscript'
import Superscript from '@tiptap/extension-superscript'
import Underline from '@tiptap/extension-underline'
import FontFamily from '@tiptap/extension-font-family'
import Table from '@tiptap/extension-table'
import TableRow from '@tiptap/extension-table-row'
import TableCell from '@tiptap/extension-table-cell'
import TableHeader from '@tiptap/extension-table-header'

/* ---------- 字号：Tiptap 官方无内置，用 TextStyle 的 fontSize 属性扩展 ---------- */

/** 字号（pt）作为 TextStyle 的一个属性，产出 <span style="font-size:12pt"> */
const FontSize = TextStyle.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      fontSize: {
        default: null,
        parseHTML: (el) => (el as HTMLElement).style.fontSize || null,
        renderHTML: (attrs) => (attrs.fontSize ? { style: `font-size:${attrs.fontSize}` } : {}),
      },
    }
  },
  addCommands() {
    return {
      ...this.parent?.(),
      setFontSize:
        (size: string) =>
        ({ chain }: { chain: () => CommandChain }) =>
          chain().setMark('textStyle', { fontSize: size }).run(),
      unsetFontSize:
        () =>
        ({ chain }: { chain: () => CommandChain }) =>
          chain().setMark('textStyle', { fontSize: null }).removeEmptyTextStyle().run(),
    } as never
  },
})

/** chain() 的最小形状（避免依赖 Tiptap 内部类型导出） */
interface CommandChain {
  setMark: (name: string, attrs: Record<string, unknown>) => CommandChain
  removeEmptyTextStyle: () => CommandChain
  run: () => boolean
}

/** 行距：段落级属性，产出 <p style="line-height:1.5"> */
const LineHeight = TextStyle.extend({
  name: 'lineHeight',
  addAttributes() {
    return {
      lineHeight: {
        default: null,
        parseHTML: (el) => (el as HTMLElement).style.lineHeight || null,
        renderHTML: (attrs) => (attrs.lineHeight ? { style: `line-height:${attrs.lineHeight}` } : {}),
      },
    }
  },
})

/* ---------- 图片：保留 data-wrap / data-hside（环绕语义），宽度显式写死 ---------- */

/**
 * Image 扩展增强：把环绕属性作为 node attrs 保存。
 * 默认 Image 只认 src/alt/title，data-wrap 会在 HTML 序列化时丢失 →
 * 导出时判不出环绕类型（这正是「环绕设置不生效」的根因）。
 */
const ImageWrap = Image.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      width: {
        default: null,
        parseHTML: (el) => (el as HTMLElement).getAttribute('width') || null,
        renderHTML: (attrs) => (attrs.width ? { width: attrs.width } : {}),
      },
      height: {
        default: null,
        parseHTML: (el) => (el as HTMLElement).getAttribute('height') || null,
        renderHTML: (attrs) => (attrs.height ? { height: attrs.height } : {}),
      },
      /** inline | square | topBottom */
      dataWrap: {
        default: 'inline',
        parseHTML: (el) => (el as HTMLElement).getAttribute('data-wrap') || 'inline',
        renderHTML: (attrs) => (attrs.dataWrap && attrs.dataWrap !== 'inline' ? { 'data-wrap': attrs.dataWrap } : {}),
      },
      /** square 时有效：left | right */
      dataHside: {
        default: null,
        parseHTML: (el) => (el as HTMLElement).getAttribute('data-hside') || null,
        renderHTML: (attrs) => (attrs.dataHside ? { 'data-hside': attrs.dataHside } : {}),
      },
    }
  },
})

/**
 * 取编辑内容 HTML，剔除 ProseMirror 的内部占位元素。
 *
 * ProseMirror 会在 DOM 里插 `.ProseMirror-separator`（空 img）和
 * `.ProseMirror-trailingBreak`（尾随 br）等纯布局用节点，它们不属于文档内容。
 * 直接 getHTML() 偶尔会把 separator 的 img 带出来 → html2ooxml 当成真图片注入
 * media（实测多出一份 media，且打开 Word 是空白图）。这里在序列化后用 DOM 清洗。
 */
function cleanHtml(html: string): string {
  if (typeof document === 'undefined') return html
  const tmp = document.createElement('div')
  tmp.innerHTML = html
  tmp.querySelectorAll('.ProseMirror-separator, .ProseMirror-trailingBreak').forEach((el) => el.remove())
  // 去掉仅用于布局的空 img（无 src）
  tmp.querySelectorAll('img:not([src]), img[src=""]').forEach((el) => el.remove())
  return tmp.innerHTML
}

/* ---------- 扩展集合 ---------- */

const extensions = [
  StarterKit.configure({
    // 标题级别与 OOXML 的 heading1-6 对齐
    heading: { levels: [1, 2, 3, 4, 5, 6] },
    // StarterKit 自带 bold/italic/strike/code/blockquote/bulletList/orderedList/history 等
    codeBlock: false, // 造价报告用不到代码块，且 OOXML 侧无映射
  }),
  Underline,
  Subscript,
  Superscript,
  TextStyle,
  FontSize,
  LineHeight,
  Color,
  Highlight.configure({
    multicolor: true, // 允许任意色（不只 16 命名色），配合 w:shd 导出
  }),
  FontFamily,
  TextAlign.configure({ types: ['heading', 'paragraph'] }),
  ImageWrap.configure({ inline: false, allowBase64: true }),
  Table.configure({ resizable: false, lastColumnResizable: false }),
  TableRow,
  TableHeader,
  TableCell,
]

/* ---------- 组件 ---------- */

export { cleanHtml }

export interface TiptapDocEditorProps {
  /** 初始 HTML（来自 ooxml2html；只在挂载时读一次） */
  value: string
  /** 内容变化回调（防抖由调用方负责） */
  onChange: (html: string) => void
  /** 失焦回写 */
  onBlur?: () => void
  className?: string
  /** 供外部调用命令（Ribbon 按钮接线用） */
  editorRef?: { current: Editor | null }
  /**
   * 编辑器实例就绪回调。
   * 存在的理由：`editorRef.current` 在子组件 effect 里赋值，
   * 父组件没法保证自己的 effect 在它之后跑 —— 只靠 ref 订阅 selectionUpdate 会拿到 null
   * （2026-09-28 实测：选区格式同步失效，S3 hasOn=false）。
   * 用 state 存实例并回调通知，父组件据此重渲染并重新订阅。
   */
  onReady?: (editor: Editor | null) => void
}

export function TiptapDocEditor({ value, onChange, onBlur, className, editorRef, onReady }: TiptapDocEditorProps) {  // onChange 用 ref 转发：避免每次父组件重渲染都重建编辑器
  const onChangeRef = useRef(onChange)
  onChangeRef.current = onChange
  const onBlurRef = useRef(onBlur)
  onBlurRef.current = onBlur
  const onReadyRef = useRef(onReady)
  onReadyRef.current = onReady

  const editor = useEditor({
    extensions,
    content: value,
    editable: true,
    editorProps: {
      attributes: {
        class: 'wrp-edit-inner',
        spellcheck: 'false',
      },
    },
    onUpdate: ({ editor: e }) => onChangeRef.current(cleanHtml(e.getHTML())),
    onBlur: () => onBlurRef.current?.(),
  })

  // 暴露 editor 实例给父组件（Ribbon 命令层用）+ 通知父组件实例就绪（用于订阅选区变化）
  useEffect(() => {
    if (editorRef) editorRef.current = editor
    onReadyRef.current?.(editor)
    return () => {
      if (editorRef) editorRef.current = null
      onReadyRef.current?.(null)
    }
  }, [editor, editorRef])

  // 外部 value 变化且与当前内容不同时同步（仅当编辑器未聚焦，避免打断输入）
  useEffect(() => {
    if (!editor) return
    if (editor.isFocused) return
    const cur = editor.getHTML()
    if (cur !== value) editor.commands.setContent(value || '<p></p>', false)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value, editor])

  if (!editor) return null
  return <EditorContent editor={editor} className={className} />
}

export type { Editor }
