/**
 * 操作审计日志（src/core/audit.ts）单测
 *
 * 测试环境是 node（vitest.config.ts: environment: 'node'），没有 window/document。
 * 这正是最关键的一条行为：**在非浏览器环境里 audit() 必须彻底空转** ——
 * store.setCell 里埋了审计，若它真的入队并排 250ms 定时器，
 * 现有 1000+ 用例会在收尾阶段被 "work after teardown" 拖挂。
 *
 * 其余用例覆盖纯函数：日志片段清洗、按键记法、空元素描述。
 */
import { describe, it, expect, beforeEach } from 'vitest'
import {
  audit,
  auditFeedback,
  auditEditor,
  auditRecent,
  auditStatus,
  cleanAuditText,
  describeElement,
  keyLabel,
  valueOf,
  setAuditEnabled,
} from './audit'

describe('audit — 非浏览器环境（node / vitest）', () => {
  beforeEach(() => {
    // 每个用例前确认开关是开的，避免用例间污染
    setAuditEnabled(true)
  })

  it('audit() 不抛异常', () => {
    expect(() => audit('点击', 'pointerdown', 'button#x')).not.toThrow()
  })

  it('不入队（buffered 恒为 0，不产生定时器）', () => {
    audit('键盘', 'keydown', 'Ctrl+S')
    auditFeedback('S9', '保存完成')
    auditEditor('pointerdown', 'canvas')
    expect(auditStatus().buffered).toBe(0)
  })

  it('内存 ring 也为空（没有环境就别假装记了）', () => {
    audit('点击', 'click', 'x')
    expect(auditRecent(10)).toEqual([])
  })

  it('会话 id 形如 s + base36，且每次进程内唯一', () => {
    const s = auditStatus().session
    expect(s).toMatch(/^s[0-9a-z]+$/)
    expect(auditStatus().session).toBe(s)
  })
})

describe('cleanAuditText — 保证「一行一条」', () => {
  it('换行折叠成 ⏎（否则一条记录会被拆成多行，grep 全乱）', () => {
    expect(cleanAuditText('a\nb')).toBe('a⏎b')
    expect(cleanAuditText('a\r\nb')).toBe('a⏎b')
    expect(cleanAuditText('a\tb')).toBe('a b')
  })

  it('超长截断并标注多出来的长度', () => {
    const long = 'x'.repeat(1200)
    const out = cleanAuditText(long)
    expect(out.length).toBeLessThan(long.length)
    expect(out).toContain('…(+')
  })

  it('对象序列化成 JSON（不落 [object Object]）', () => {
    expect(cleanAuditText({ a: 1 })).toBe('{"a":1}')
  })

  it('null / undefined → 空串', () => {
    expect(cleanAuditText(null)).toBe('')
    expect(cleanAuditText(undefined)).toBe('')
  })

  it('数字与布尔照常转字符串', () => {
    expect(cleanAuditText(0)).toBe('0')
    expect(cleanAuditText(false)).toBe('false')
  })
})

describe('keyLabel — 按键记法', () => {
  const k = (o: Partial<KeyboardEvent>) => o as KeyboardEvent

  it('单个字符键', () => {
    expect(keyLabel(k({ key: 'a' }))).toBe('a')
  })

  it('修饰键顺序固定 Ctrl → Alt → Shift → Meta', () => {
    expect(keyLabel(k({ key: 's', ctrlKey: true }))).toBe('Ctrl+s')
    expect(keyLabel(k({ key: 'S', ctrlKey: true, shiftKey: true }))).toBe('Ctrl+Shift+S')
    expect(keyLabel(k({ key: 'z', ctrlKey: true, altKey: true, shiftKey: true, metaKey: true }))).toBe('Ctrl+Alt+Shift+Meta+z')
  })

  it('空格记作 Space（否则日志里出现裸双空格，无法辨认）', () => {
    expect(keyLabel(k({ key: ' ' }))).toBe('Space')
  })

  it('功能键原样保留', () => {
    expect(keyLabel(k({ key: 'Enter' }))).toBe('Enter')
    expect(keyLabel(k({ key: 'F2' }))).toBe('F2')
  })
})

describe('describeElement / valueOf — 空输入的兜底', () => {
  it('null / undefined 不抛，返回占位符', () => {
    expect(describeElement(null)).toBe('-')
    expect(describeElement(undefined)).toBe('-')
    expect(valueOf(null)).toBe('')
  })
})
