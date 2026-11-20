import { useEffect, useMemo } from 'react'
import type { TableSchema } from './core/schema'
import { KIND_STYLE } from './core/nodeStyle'
import { identityOwnerOf, membersOf, primaryOf } from './core/ruleSync'

/**
 * 规则跟随节点的「成员管理」浮层（2026-10-05）
 *
 * 用户口径：
 *   「规则节点第一个默认是可修改的主节点。连接剩下的是从节点。」
 *   「规则节点上显示几个节点。且双击能修改节点的主次、跟随不跟随。」
 *   （双击面板要能做的操作 →）「切换跟随/不跟随, 改主次顺序, 移出该组」
 *
 * ── 面板只做三件事，一件不多 ──
 *   ① 改主次：把某个成员设成主节点（原主自动降为从）
 *   ② 跟随/ 不跟随：连线保留，只切是否受本规则影响
 *   ③ 移出该组：断开连线（= unlinkFromRuleNode）
 *
 * ★ 刻意**不**在这里放「编辑规则内容」：
 *   规则只存在规则节点上，要改规则去右侧属性栏的 RuleEditor（那里按 kind 分桶）。
 *   两处都放编辑框会出现两份规则输入路径，早晚会有一份没同步上。
 *
 * ★ 为什么状态判定全走 core/ruleSync.membersOf：
 *   「谁是主、谁跟随、谁是主」若在 UI 里另数一遍，规则一改两处就漂移
 *   （症状：面板显示 3 个成员，圆心上写2 个）。单一真相源铁律。
 */
export function RuleMembersPanel({
  ruleNodeId,
  tables,
  onClose,
  onSetPrimary,
  onSetFollow,
  onUnlink,
}: {
  ruleNodeId: string
  tables: Record<string, TableSchema>
  onClose: () => void
  /** 设为主节点；原主节点自动降为从节点（store 内保证组内恒有且仅有一个 primary） */
  onSetPrimary: (ruleNodeId: string, targetId: string) => void
  /** 切换跟随 / 不跟随 */
  onSetFollow: (ruleNodeId: string, targetId: string, follow: boolean) => void
  /** 移出该组（断开连线） */
  onUnlink: (ruleNodeId: string, targetId: string) => void
}) {
  const ruleNode = tables[ruleNodeId]
  // membersOf 已按 主 → 从 → 不跟随 排序，主节点天然在第一行
  const members = useMemo(() => membersOf(ruleNodeId, tables), [ruleNodeId, tables])
  const primary = useMemo(() => primaryOf(ruleNodeId, tables), [ruleNodeId, tables])

  // Esc 关闭 + 点击遮罩关闭。浮层不是画布节点，走全局监听最省事。
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  // 规则节点被删 / 进回收站 → 自动关掉浮层，否则留下指向不存在节点的浮层
  useEffect(() => {
    if (!ruleNode || ruleNode.deleted) onClose()
  }, [ruleNode, onClose])

  if (!ruleNode) return null

  return (
    <div
      className="rule-members-mask"
      // 遮罩按下即关闭；面板本体 stopPropagation，避免点面板内按钮时误关
      onPointerDown={onClose}
    >
      <div className="rule-members" onPointerDown={(e) => e.stopPropagation()}>
        <div className="rule-members-head">
          <span className="rule-members-title">{ruleNode.name}</span>
          <span className="rule-members-sub">
            规则跟随 · 共 {members.length} 个节点
            {primary ? ` · 主：${tables[primary.targetId]?.name ?? primary.targetId}` : ' · 暂无主节点'}
          </span>
          <button type="button" className="rule-members-x" title="关闭（Esc）" onClick={onClose}>
            ×
          </button>
        </div>

        {/* ★ 内容区必须单独一层：.rule-members 改成 flex 列布局后，
           若把列表直接挂在面板下，max-height + overflow 就不成立
           （列表会顶着 footer 一起溢出，底部按钮被裁掉）。
           flex:1 + min-height:0 是滚动容器嵌在 flex 列里的必要条件。*/}
        <div className="rule-members-body">
          {members.length === 0 ? (
            <p className="rule-members-empty">
              还没有节点连到本规则节点。
              <br />
              把其它节点的<b>上方连接口</b>拖到圆形的<b>下方口</b>，就能纳入这一组。
              <br />
              同一组内的节点必须是<b>同一种类型</b>（第一个成员的类型由它定下）。
            </p>
          ) : (
            <div className="rule-members-list">
              {members.map((m, i) => {
                const name = tables[m.targetId]?.name ?? m.targetId
                const kindLabel = KIND_STYLE[m.targetKind]?.label ?? m.targetKind
                const isPrimary = m.mode === 'primary'
                const following = m.mode !== 'ignore'
                /* 该成员的身份槽是不是被**别的**规则组占着。
                   ruleRef 是单槽，占用时本组改它就会抹掉另一个组的身份（store 已挡），
                   故这里如实置灰并说明，而不是让用户点了没反应。*/
                const takenBy = identityOwnerOf(m.targetId, tables, ruleNodeId)
                return (
                  <div key={m.targetId} className={`rule-members-row rule-members-row--${m.mode}`}>
                    <span className="rule-members-idx">{i + 1}</span>
                    <div className="rule-members-info">
                      {/* 身份徽标（2026-10-05 用户第 4 条）。
                          原来只有下面那行10px 小字写着「主节点/从节点」，
                          字号太小、颜色太灰，用户反馈「颜色不一致」——
                          现在把身份提到**名字前面**做成实心/描边徽标，
                          与画布上规则线「实线=主、虚线=跟随」同一套语义，
                          扫一眼列表就知道谁说了算，不必逐行读小字。*/}
                      <span className="rule-members-name" title={name}>
                        <span
                          className={`rule-members-badge rule-members-badge--${isPrimary ? 'primary' : following ? 'follow' : 'ignore'}`}
                        >
                          {isPrimary ? '主' : following ? '从' : '不跟'}
                        </span>
                        {name}
                      </span>
                      <span className="rule-members-kind">
                        {isPrimary ? '规则以它为准' : following ? '跟随主节点' : '不跟随（用回自己的设置）'}
                        {' · '}
                        {kindLabel}
                      </span>
                      {takenBy && (
                        <span className="rule-members-taken" title={`它的主从/跟随身份归「${tables[takenBy]?.name ?? takenBy}」所有，本组改不了`}>
                          身份归「{tables[takenBy]?.name ?? takenBy}」，本组只跟随
                        </span>
                      )}
                    </div>
                    <div className="rule-members-ops">
                      {!isPrimary && (
                        <button
                          type="button"
                          className="rule-btn"
                          title={
                            takenBy
                              ? `它的身份归「${tables[takenBy]?.name ?? takenBy}」，改主次会破坏那个组`
                              : '设为主节点：规则以它为准；原主节点自动降为从节点'
                          }
                          disabled={!!takenBy}
                          onClick={() => onSetPrimary(ruleNodeId, m.targetId)}
                        >
                          设为主
                        </button>
                      )}
                      {isPrimary ? (
                        /* 主节点不能直接「不跟随」—— 组内必须留一个能改规则的节点。
                           store 里 setRuleFollow 会把它降为从并把主位让给下一个成员；
                           一个成员都没有时直接拒绝。这里如实告知用户会怎样。*/
                        <button
                          type="button"
                          className="rule-btn"
                          title={
                            members.length > 1
                              ? '取消主节点跟随：会自动降为从节点，主位让给下一个成员'
                              : '本组只有它一个成员，不能取消跟随（否则没有人执行规则）'
                          }
                          disabled={members.length <= 1 || !!takenBy}
                          onClick={() => onSetFollow(ruleNodeId, m.targetId, false)}
                        >
                          取消跟随
                        </button>
                      ) : (
                        <button
                          type="button"
                          className={`rule-btn${following ? '' : ' rule-btn--primary'}`}
                          title={
                            takenBy
                              ? `它的身份归「${tables[takenBy]?.name ?? takenBy}」，本组改不了`
                              : following
                                ? '改为不跟随：连线保留，但不受本规则影响'
                                : '改为跟随主节点的规则'
                          }
                          disabled={!!takenBy}
                          onClick={() => onSetFollow(ruleNodeId, m.targetId, !following)}
                        >
                          {following ? '不跟随' : '跟随'}
                        </button>
                      )}
                      <button
                        type="button"
                        className="rule-btn rule-btn--danger"
                        title="移出该组：断开连线，本节点回用自己的设置"
                        onClick={() => onUnlink(ruleNodeId, m.targetId)}
                      >
                        移出
                      </button>
                    </div>
                  </div>
                )
              })}
            </div>
          )}
        </div>

        <p className="hint rule-members-foot">
          规则只存在<b>规则跟随节点</b>上，被跟随的节点不存副本 ——
          所以不存在「改了一处忘了另一处」。要改规则内容，去右侧属性栏。
        </p>
        {/* 全组身份都被别的规则组占着时 membersOf 的兜底也不指定主节点
            （挑不出「可写」的成员）。这里如实说明，不让用户对着
            「主节点：」那行空格猜。 */}
        {members.length > 0 && !primary && (
          <p className="hint rule-members-warn">
            ⚠ 本组成员的<b>主从身份都被其他规则组占用</b>，本组暂时没有主节点。
            这说明同一个节点被连到了两个规则跟随节点 —— 一个节点的身份只能归一组管。
            把它们分到不同的规则节点，或先在另一个组里「移出」即可。
          </p>
        )}
      </div>
    </div>
  )
}
