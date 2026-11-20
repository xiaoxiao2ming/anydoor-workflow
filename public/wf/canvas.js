// ══════════════════════════════════════════════════
//  工作流画布引擎  canvas.js
// ══════════════════════════════════════════════════

;(function () {
  'use strict'

  // ── 全局状态 ──────────────────────────────────────
  const state = {
    nodes: [],
    edges: [],
    selectedIds: new Set(),
    selectedEdgeIds: new Set(),
    tool: 'select',         // 'select' | 'connect'
    view: { x: 0, y: 0, scale: 1 },
    drag: null,             // 拖拽状态
    connect: null,          // 连线状态
    selectBox: null,        // 框选状态
    history: [],
    historyIndex: -1,
    fileName: '未命名.wf',
    idCounter: 0,
    editingEdge: null,       // 正在编辑的连线ID
    edgeHandleDrag: null     // { edgeId, side, otherPt, otherNodeId, otherPort }
  }

  // ── DOM 引用 ──────────────────────────────────────
  const svg       = document.getElementById('mainSvg')
  const nodeLayer = document.getElementById('nodeLayer')
  const edgeLayer = document.getElementById('edgeLayer')
  const tempLayer = document.getElementById('tempLayer')
  const container = document.getElementById('canvasContainer')
  const propsBody = document.getElementById('propsBody')
  const ctxMenu   = document.getElementById('ctxMenu')
  const fileNameEl= document.getElementById('fileName')
  const zoomLabel = document.getElementById('zoomLabel')
  const zoomSlider= document.getElementById('zoomSlider')
  const minimapCanvas = document.getElementById('minimapCanvas')
  const mCtx = minimapCanvas.getContext('2d')

  // ── 节点类型定义 ──────────────────────────────────
  const NODE_TYPES = {
    start:      { label: '开始',   w: 100, h: 40,  color: '#639922', textColor: '#fff', badge: 'badge-start' },
    end:        { label: '结束',   w: 100, h: 40,  color: '#a32d2d', textColor: '#fff', badge: 'badge-end' },
    task:       { label: '任务',   w: 140, h: 56,  color: '#378ADD', textColor: '#0c447c', badge: 'badge-task' },
    decision:   { label: '判断',   w: 100, h: 70,  color: '#BA7517', textColor: '#633806', badge: 'badge-decision' },
    subprocess: { label: '子流程', w: 140, h: 56,  color: '#534AB7', textColor: '#3c3489', badge: 'badge-subprocess' }
  }

  // ── 工具函数 ──────────────────────────────────────
  function uid() { return 'n' + (++state.idCounter) + '_' + Date.now().toString(36) }

  function svgPoint(clientX, clientY) {
    const rect = svg.getBoundingClientRect()
    return {
      x: (clientX - rect.left - state.view.x) / state.view.scale,
      y: (clientY - rect.top  - state.view.y) / state.view.scale
    }
  }

  function applyTransform() {
    const g = document.getElementById('gridLayer')
    const edgeG = edgeLayer
    const nodeG = nodeLayer
    const t = `translate(${state.view.x},${state.view.y}) scale(${state.view.scale})`
    document.querySelectorAll('#gridLayer,#edgeLayer,#nodeLayer,#tempLayer').forEach(el => {
      el.setAttribute('transform', t)
    })
    zoomLabel.textContent = Math.round(state.view.scale * 100) + '%'
    zoomSlider.value = Math.round(state.view.scale * 100)
    updateMinimap()
  }

  // ── 历史 / 撤销 ──────────────────────────────────
  function saveHistory() {
    const snap = JSON.stringify({ nodes: state.nodes, edges: state.edges, idCounter: state.idCounter })
    state.history = state.history.slice(0, state.historyIndex + 1)
    state.history.push(snap)
    if (state.history.length > 80) state.history.shift()
    state.historyIndex = state.history.length - 1
  }

  function undo() {
    if (state.historyIndex <= 0) return
    state.historyIndex--
    restoreSnap(state.history[state.historyIndex])
  }

  function redo() {
    if (state.historyIndex >= state.history.length - 1) return
    state.historyIndex++
    restoreSnap(state.history[state.historyIndex])
  }

  function restoreSnap(snap) {
    const data = JSON.parse(snap)
    state.nodes = data.nodes
    state.edges = data.edges
    state.idCounter = data.idCounter
    state.selectedIds.clear()
    state.selectedEdgeIds.clear()
    renderAll()
    updatePropsPanel()
  }

  // ── 渲染 ──────────────────────────────────────────
  function renderAll() {
    renderEdges()
    renderNodes()
    updateMinimap()
  }

  // ── 节点渲染 ──────────────────────────────────────
  function renderNodes() {
    nodeLayer.innerHTML = ''
    state.nodes.forEach(n => renderNode(n))
  }

  function renderNode(n) {
    const def = NODE_TYPES[n.type]
    const g = document.createElementNS('http://www.w3.org/2000/svg', 'g')
    g.setAttribute('class', 'node-group' + (state.selectedIds.has(n.id) ? ' selected' : ''))
    g.setAttribute('data-id', n.id)

    const hw = def.w / 2, hh = def.h / 2

    // ── 节点主体形状 ──
    let body
    if (n.type === 'start' || n.type === 'end') {
      body = document.createElementNS('http://www.w3.org/2000/svg', 'rect')
      body.setAttribute('x', n.x - hw)
      body.setAttribute('y', n.y - hh)
      body.setAttribute('width', def.w)
      body.setAttribute('height', def.h)
      body.setAttribute('rx', def.h / 2)
      body.setAttribute('fill', def.color)
      body.setAttribute('stroke', state.selectedIds.has(n.id) ? '#639922' : 'none')
      body.setAttribute('stroke-width', '2')
    } else if (n.type === 'decision') {
      body = document.createElementNS('http://www.w3.org/2000/svg', 'polygon')
      const dw = def.w / 2, dh = def.h / 2
      body.setAttribute('points',
        `${n.x},${n.y - dh} ${n.x + dw},${n.y} ${n.x},${n.y + dh} ${n.x - dw},${n.y}`)
      body.setAttribute('fill', '#fff')
      body.setAttribute('stroke', state.selectedIds.has(n.id) ? '#639922' : def.color)
      body.setAttribute('stroke-width', state.selectedIds.has(n.id) ? '2' : '1.5')
    } else {
      body = document.createElementNS('http://www.w3.org/2000/svg', 'rect')
      body.setAttribute('x', n.x - hw)
      body.setAttribute('y', n.y - hh)
      body.setAttribute('width', def.w)
      body.setAttribute('height', def.h)
      body.setAttribute('rx', '6')
      body.setAttribute('fill', '#fff')
      body.setAttribute('stroke', state.selectedIds.has(n.id) ? '#639922' : def.color)
      body.setAttribute('stroke-width', state.selectedIds.has(n.id) ? '2' : '1.5')
    }
    body.setAttribute('class', 'node-body')
    g.appendChild(body)

    // ── 子流程下边框线 ──
    if (n.type === 'subprocess') {
      const line = document.createElementNS('http://www.w3.org/2000/svg', 'line')
      line.setAttribute('x1', n.x - hw + 8)
      line.setAttribute('x2', n.x + hw - 8)
      line.setAttribute('y1', n.y + hh - 9)
      line.setAttribute('y2', n.y + hh - 9)
      line.setAttribute('stroke', def.color)
      line.setAttribute('stroke-width', '1.5')
      g.appendChild(line)
    }

    // ── 节点文字 ──
    const label = document.createElementNS('http://www.w3.org/2000/svg', 'text')
    label.setAttribute('x', n.x)
    label.setAttribute('y', n.y + (n.type === 'task' || n.type === 'subprocess' ? -6 : 0))
    label.setAttribute('class', 'node-label')
    label.setAttribute('fill', (n.type === 'start' || n.type === 'end') ? def.textColor : def.textColor)
    label.textContent = n.label || def.label
    g.appendChild(label)

    if (n.type === 'task' || n.type === 'subprocess') {
      const sub = document.createElementNS('http://www.w3.org/2000/svg', 'text')
      sub.setAttribute('x', n.x)
      sub.setAttribute('y', n.y + 9)
      sub.setAttribute('class', 'node-type-label')
      sub.textContent = n.desc ? (n.desc.length > 16 ? n.desc.slice(0, 14) + '…' : n.desc) : (n.type === 'task' ? '点击编辑描述' : '子工作流')
      g.appendChild(sub)
    }

    // ── 连接端口 ──
    const ports = getPortPositions(n)
    ports.forEach(p => {
      const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle')
      circle.setAttribute('cx', p.x)
      circle.setAttribute('cy', p.y)
      circle.setAttribute('r', '5')
      circle.setAttribute('class', 'port')
      circle.setAttribute('data-port', p.dir)
      circle.setAttribute('data-node', n.id)
      g.appendChild(circle)
    })

    g.addEventListener('mousedown', onNodeMouseDown)
    g.addEventListener('contextmenu', onNodeCtxMenu)

    nodeLayer.appendChild(g)
  }

  function getPortPositions(n) {
    const def = NODE_TYPES[n.type]
    const hw = def.w / 2, hh = def.h / 2
    if (n.type === 'decision') {
      const dw = def.w / 2, dh = def.h / 2
      return [
        { dir: 'top',    x: n.x,       y: n.y - dh },
        { dir: 'right',  x: n.x + dw,  y: n.y },
        { dir: 'bottom', x: n.x,       y: n.y + dh },
        { dir: 'left',   x: n.x - dw,  y: n.y }
      ]
    }
    return [
      { dir: 'top',    x: n.x,       y: n.y - hh },
      { dir: 'right',  x: n.x + hw,  y: n.y },
      { dir: 'bottom', x: n.x,       y: n.y + hh },
      { dir: 'left',   x: n.x - hw,  y: n.y }
    ]
  }

  function getNodeById(id) { return state.nodes.find(n => n.id === id) }

  // ── 连线流动动画（JS 驱动，比 CSS 更可靠） ─────────
  const flowAnims = {}  // { edgeId: { path, offset, raf } }

  function startEdgeFlow(pathEl) {
    const eid = pathEl.getAttribute('data-edge-id')
    if (flowAnims[eid]) cancelAnimationFrame(flowAnims[eid].raf)

    let offset = 0
    function tick() {
      offset -= 1.5
      pathEl.setAttribute('stroke-dashoffset', offset)
      flowAnims[eid] = { path: pathEl, offset, raf: requestAnimationFrame(tick) }
    }
    tick()
  }

  function stopEdgeFlow(eid) {
    if (flowAnims[eid]) {
      cancelAnimationFrame(flowAnims[eid].raf)
      delete flowAnims[eid]
    }
  }

  function stopAllEdgeFlows() {
    Object.keys(flowAnims).forEach(id => stopEdgeFlow(id))
  }

  // ── 右键拆离时连线高亮 ─────────────────────────
  function highlightEdge(eid) {
    const els = edgeLayer.querySelectorAll(`[data-edge-id="${eid}"]`)
    els.forEach(el => {
      if (el.classList.contains('edge-path')) {
        el.classList.add('detach-hover')
      }
    })
  }

  function unhighlightEdge(eid) {
    const els = edgeLayer.querySelectorAll(`[data-edge-id="${eid}"]`)
    els.forEach(el => {
      el.classList.remove('detach-hover')
    })
  }

  // ── 边/连线渲染 ──────────────────────────────────
  function renderEdges() {
    edgeLayer.innerHTML = ''
    state.edges.forEach(e => renderEdge(e))
  }

  function renderEdge(e) {
    const src = getNodeById(e.source)
    const dst = getNodeById(e.target)
    if (!src || !dst) return

    const srcPt = getPortPoint(src, e.sourcePort || 'right')
    const dstPt = getPortPoint(dst, e.targetPort || 'left')
    const path = makeBezierPath(srcPt, dstPt)
    const isEditing = state.editingEdge === e.id

    // 宽hit区
    const hitPath = document.createElementNS('http://www.w3.org/2000/svg', 'path')
    hitPath.setAttribute('d', path)
    hitPath.setAttribute('fill', 'none')
    hitPath.setAttribute('stroke', 'transparent')
    hitPath.setAttribute('stroke-width', '12')
    hitPath.setAttribute('data-edge-id', e.id)
    hitPath.style.cursor = 'pointer'
    hitPath.addEventListener('mousedown', onEdgeMouseDown)
    hitPath.addEventListener('contextmenu', onEdgeCtxMenu)
    hitPath.addEventListener('dblclick', onEdgeDblClick)
    edgeLayer.appendChild(hitPath)

    // 可见路径
    const visPath = document.createElementNS('http://www.w3.org/2000/svg', 'path')
    visPath.setAttribute('d', path)
    const cls = ['edge-path']
    if (state.selectedEdgeIds.has(e.id)) cls.push('selected')
    if (isEditing) cls.push('editing')
    visPath.setAttribute('class', cls.join(' '))
    visPath.setAttribute('data-edge-id', e.id)
    visPath.setAttribute('pointer-events', 'none')
    if (isEditing) {
      visPath.setAttribute('stroke-dasharray', '8 4')
      visPath.setAttribute('stroke', '#378ADD')
      visPath.setAttribute('stroke-width', '2')
      startEdgeFlow(visPath)
    }
    edgeLayer.appendChild(visPath)

    // 编辑模式：端点拖拽手柄
    if (isEditing) {
      renderEdgeHandle(srcPt, e, 'source', src)
      renderEdgeHandle(dstPt, e, 'target', dst)
    }

    // 边标签
    if (e.label) {
      const mid = getMidPoint(srcPt, dstPt)
      const bg = document.createElementNS('http://www.w3.org/2000/svg', 'rect')
      bg.setAttribute('x', mid.x - 20)
      bg.setAttribute('y', mid.y - 9)
      bg.setAttribute('width', 40)
      bg.setAttribute('height', 16)
      bg.setAttribute('class', 'edge-label-bg')
      bg.setAttribute('rx', '4')
      edgeLayer.appendChild(bg)

      const txt = document.createElementNS('http://www.w3.org/2000/svg', 'text')
      txt.setAttribute('x', mid.x)
      txt.setAttribute('y', mid.y)
      txt.setAttribute('class', 'edge-label')
      txt.textContent = e.label
      edgeLayer.appendChild(txt)
    }
  }

  function renderEdgeHandle(pt, edge, side, node) {
    const g = document.createElementNS('http://www.w3.org/2000/svg', 'g')
    g.setAttribute('class', 'edge-handle')
    g.setAttribute('data-edge-id', edge.id)
    g.setAttribute('data-side', side)

    // 外圈（脉动）
    const outer = document.createElementNS('http://www.w3.org/2000/svg', 'circle')
    outer.setAttribute('cx', pt.x)
    outer.setAttribute('cy', pt.y)
    outer.setAttribute('r', '12')
    outer.setAttribute('fill', 'none')
    outer.setAttribute('stroke', '#378ADD')
    outer.setAttribute('stroke-width', '1')
    outer.setAttribute('opacity', '0.3')
    outer.setAttribute('class', 'edge-handle-pulse')
    g.appendChild(outer)

    // 拖拽手柄
    const handle = document.createElementNS('http://www.w3.org/2000/svg', 'circle')
    handle.setAttribute('cx', pt.x)
    handle.setAttribute('cy', pt.y)
    handle.setAttribute('r', '7')
    handle.setAttribute('fill', '#fff')
    handle.setAttribute('stroke', '#378ADD')
    handle.setAttribute('stroke-width', '2')
    handle.setAttribute('class', 'edge-handle-grip')
    handle.style.cursor = 'grab'
    handle.addEventListener('mousedown', (e) => onEdgeHandleMouseDown(e, edge, side, node))
    g.appendChild(handle)

    edgeLayer.appendChild(g)
  }

  function getPortPoint(n, dir) {
    const def = NODE_TYPES[n.type]
    const hw = def.w / 2, hh = def.h / 2
    if (n.type === 'decision') {
      const dw = def.w / 2, dh = def.h / 2
      if (dir === 'top')    return { x: n.x,       y: n.y - dh }
      if (dir === 'right')  return { x: n.x + dw,  y: n.y }
      if (dir === 'bottom') return { x: n.x,       y: n.y + dh }
      if (dir === 'left')   return { x: n.x - dw,  y: n.y }
      return { x: n.x + dw, y: n.y }
    }
    if (dir === 'top')    return { x: n.x,      y: n.y - hh }
    if (dir === 'right')  return { x: n.x + hw, y: n.y }
    if (dir === 'bottom') return { x: n.x,      y: n.y + hh }
    if (dir === 'left')   return { x: n.x - hw, y: n.y }
    return { x: n.x + hw, y: n.y }
  }

  function makeBezierPath(a, b) {
    const dx = Math.abs(b.x - a.x) * 0.5 + 30
    return `M ${a.x} ${a.y} C ${a.x + dx} ${a.y}, ${b.x - dx} ${b.y}, ${b.x} ${b.y}`
  }

  function getMidPoint(a, b) {
    return { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 }
  }

  // ── 鼠标事件：画布 ────────────────────────────────
  let panStart = null
  let spaceDown = false

  document.addEventListener('keydown', e => {
    if (e.key === ' ') { spaceDown = true; container.style.cursor = 'grab' }
    if (e.key === 'v' || e.key === 'V') setTool('select')
    if (e.key === 'c' || e.key === 'C') setTool('connect')
    if (e.key === 'Escape') {
      if (state.editingEdge) { exitEdgeEdit(); return }
      clearSelection()
      state.connect = null
      tempLayer.innerHTML = ''
    }
    if ((e.key === 'Delete' || e.key === 'Backspace') && !['INPUT','TEXTAREA'].includes(document.activeElement.tagName)) {
      if (state.editingEdge) {
        state.edges = state.edges.filter(ed => ed.id !== state.editingEdge)
        state.editingEdge = null
        state.selectedEdgeIds.clear()
        saveHistory()
        renderAll()
        updatePropsPanel()
        return
      }
      deleteSelected()
    }
    if (e.ctrlKey && e.key === 'z') { e.preventDefault(); undo() }
    if (e.ctrlKey && e.key === 'y') { e.preventDefault(); redo() }
    if (e.ctrlKey && e.key === 's') { e.preventDefault(); saveFile() }
    if (e.ctrlKey && e.key === 'a') { e.preventDefault(); selectAll() }
  })
  document.addEventListener('keyup', e => {
    if (e.key === ' ') { spaceDown = false; container.style.cursor = '' }
  })

  svg.addEventListener('mousedown', onSvgMouseDown)
  svg.addEventListener('mousemove', onSvgMouseMove)
  svg.addEventListener('mouseup',   onSvgMouseUp)
  svg.addEventListener('wheel', onWheel, { passive: false })
  svg.addEventListener('contextmenu', e => { e.preventDefault(); hideCtxMenu() })
  document.addEventListener('click', hideCtxMenu)

  function onSvgMouseDown(e) {
    if (e.button !== 0 && e.button !== 1) return
    // 点击空白处退出编辑模式
    if (state.editingEdge && (e.target === svg || e.target.closest('#gridLayer'))) {
      exitEdgeEdit()
      return
    }
    if (e.target !== svg && !e.target.closest('#gridLayer')) {
      // 点击非画布区域由节点/边事件处理
      if (!e.target.closest('.node-group') && !e.target.hasAttribute('data-edge-id')) {
        clearSelection()
        updatePropsPanel()
      }
      return
    }

    const pt = svgPoint(e.clientX, e.clientY)

    if (spaceDown || e.button === 1) {
      panStart = { mx: e.clientX, my: e.clientY, vx: state.view.x, vy: state.view.y }
      container.style.cursor = 'grabbing'
      e.preventDefault()
      return
    }

    if (state.tool === 'select') {
      clearSelection()
      updatePropsPanel()
      // 开始框选
      state.selectBox = { sx: pt.x, sy: pt.y, ex: pt.x, ey: pt.y }
    }
  }

  function onSvgMouseMove(e) {
    if (panStart) {
      state.view.x = panStart.vx + (e.clientX - panStart.mx)
      state.view.y = panStart.vy + (e.clientY - panStart.my)
      applyTransform()
      return
    }

    if (state.drag) {
      const pt = svgPoint(e.clientX, e.clientY)
      const dx = pt.x - state.drag.startPt.x
      const dy = pt.y - state.drag.startPt.y
      state.drag.ids.forEach(id => {
        const n = getNodeById(id)
        if (!n) return
        const orig = state.drag.origPositions[id]
        n.x = Math.round((orig.x + dx) / 10) * 10
        n.y = Math.round((orig.y + dy) / 10) * 10
      })
      renderAll()
      return
    }

    if (state.connect) {
      const pt = svgPoint(e.clientX, e.clientY)
      const srcPt = state.connect.srcPt
      tempLayer.innerHTML = ''
      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path')
      path.setAttribute('d', makeBezierPath(srcPt, pt))
      path.setAttribute('class', 'temp-edge')
      tempLayer.appendChild(path)
      return
    }

    if (state.selectBox) {
      const pt = svgPoint(e.clientX, e.clientY)
      state.selectBox.ex = pt.x
      state.selectBox.ey = pt.y
      drawSelectBox()
    }
  }

  function onSvgMouseUp(e) {
    if (panStart) {
      panStart = null
      container.style.cursor = spaceDown ? 'grab' : ''
      return
    }

    if (state.drag) {
      state.drag = null
      saveHistory()
      return
    }

    if (state.selectBox) {
      finishSelectBox()
      state.selectBox = null
      tempLayer.innerHTML = ''
    }
  }

  function onWheel(e) {
    e.preventDefault()
    const delta = e.deltaY > 0 ? 0.9 : 1.1
    const rect = svg.getBoundingClientRect()
    const mx = e.clientX - rect.left
    const my = e.clientY - rect.top
    const newScale = Math.max(0.2, Math.min(3, state.view.scale * delta))
    state.view.x = mx - (mx - state.view.x) * (newScale / state.view.scale)
    state.view.y = my - (my - state.view.y) * (newScale / state.view.scale)
    state.view.scale = newScale
    applyTransform()
  }

  // ── 节点事件 ──────────────────────────────────────
  function onNodeMouseDown(e) {
    e.stopPropagation()
    if (e.button !== 0) return

    const group = e.currentTarget
    const id = group.getAttribute('data-id')

    // 连线模式：点击端口
    if (state.tool === 'connect' || e.target.classList.contains('port')) {
      const portEl = e.target.classList.contains('port') ? e.target : null
      if (portEl) {
        const srcNode = getNodeById(portEl.getAttribute('data-node'))
        const dir = portEl.getAttribute('data-port')
        const srcPt = getPortPoint(srcNode, dir)
        state.connect = { srcId: srcNode.id, srcPort: dir, srcPt }
        e.preventDefault()
        return
      }
    }

    // 选择模式：选中并开始拖拽
    if (!state.selectedIds.has(id)) {
      if (!e.shiftKey) clearSelection()
      state.selectedIds.add(id)
      state.selectedEdgeIds.clear()
      renderNodes()
    }

    const pt = svgPoint(e.clientX, e.clientY)
    const origPositions = {}
    state.selectedIds.forEach(sid => {
      const n = getNodeById(sid)
      if (n) origPositions[sid] = { x: n.x, y: n.y }
    })
    state.drag = { ids: Array.from(state.selectedIds), startPt: pt, origPositions }
    updatePropsPanel()
  }

  function onNodeCtxMenu(e) {
    e.preventDefault()
    e.stopPropagation()
    const id = e.currentTarget.getAttribute('data-id')
    if (!state.selectedIds.has(id)) {
      clearSelection()
      state.selectedIds.add(id)
      renderNodes()
      updatePropsPanel()
    }
    showCtxMenu(e.clientX, e.clientY, { type: 'node', nodeId: id })
  }

  // ── 连线鼠标释放（落在端口上） ───────────────────
  svg.addEventListener('mouseup', e => {
    if (!state.connect) return
    const portEl = e.target.classList.contains('port') ? e.target : null
    if (portEl) {
      const tgtId = portEl.getAttribute('data-node')
      const tgtDir = portEl.getAttribute('data-port')
      if (tgtId !== state.connect.srcId) {
        // 防重复连线
        const dup = state.edges.find(ed => ed.source === state.connect.srcId && ed.target === tgtId)
        if (!dup) {
          const edge = {
            id: uid(),
            source: state.connect.srcId,
            sourcePort: state.connect.srcPort,
            target: tgtId,
            targetPort: tgtDir,
            label: ''
          }
          state.edges.push(edge)
          saveHistory()
          renderAll()
        }
      }
    }
    state.connect = null
    tempLayer.innerHTML = ''
  })

  // ── 边事件 ────────────────────────────────────────
  function onEdgeMouseDown(e) {
    e.stopPropagation()
    if (e.button !== 0) return
    const eid = e.currentTarget.getAttribute('data-edge-id')
    if (state.editingEdge === eid) return  // 双击同一连线，不退出
    if (state.editingEdge) { exitEdgeEdit() }
    state.selectedEdgeIds.clear()
    state.selectedEdgeIds.add(eid)
    renderEdges()
    updatePropsPanel()
  }

  function onEdgeDblClick(e) {
    e.stopPropagation()
    e.preventDefault()
    if (state.tool !== 'select') return
    const eid = e.currentTarget.getAttribute('data-edge-id')
    enterEdgeEdit(eid)
  }

  function onEdgeCtxMenu(e) {
    e.preventDefault()
    e.stopPropagation()
    if (state.editingEdge) return
    const eid = e.currentTarget.getAttribute('data-edge-id')
    state.selectedEdgeIds.add(eid)
    renderEdges()
    showCtxMenu(e.clientX, e.clientY, { type: 'edge' })
  }

  // ── 编辑模式 ─────────────────────────────────────
  function enterEdgeEdit(eid) {
    if (state.editingEdge && state.editingEdge !== eid) stopEdgeFlow(state.editingEdge)
    state.editingEdge = eid
    state.connect = null
    state.drag = null
    renderAll()
    updatePropsPanel()
  }

  function exitEdgeEdit() {
    if (state.editingEdge) stopEdgeFlow(state.editingEdge)
    state.editingEdge = null
    state.edgeHandleDrag = null
    tempLayer.innerHTML = ''
    renderAll()
    updatePropsPanel()
  }

  // ── 端点拖拽 ─────────────────────────────────────
  function onEdgeHandleMouseDown(e, edge, side, node) {
    e.stopPropagation()
    e.preventDefault()
    const edgeObj = state.edges.find(ed => ed.id === edge.id)
    if (!edgeObj) return

    const otherNodeId = side === 'source' ? edgeObj.target : edgeObj.source
    const otherPort   = side === 'source' ? edgeObj.targetPort : edgeObj.sourcePort
    const otherNode   = getNodeById(otherNodeId)
    if (!otherNode) return
    const otherPt = getPortPoint(otherNode, otherPort || (side === 'source' ? 'left' : 'right'))

    state.edgeHandleDrag = {
      edgeId: edge.id,
      side,
      otherPt,
      otherNodeId,
      otherPort,
      oldSource: edgeObj.source,
      oldTarget: edgeObj.target,
      oldSourcePort: edgeObj.sourcePort,
      oldTargetPort: edgeObj.targetPort
    }

    // 初步移除该连线（从 edges 数组暂存），连到新目标后再加回去
    state.edges = state.edges.filter(ed => ed.id !== edge.id)
    state.selectedEdgeIds.delete(edge.id)
    renderAll()
  }

  // ── 画布鼠标移动时处理端点拖拽 ─────────────────────
  // 在 onSvgMouseMove 中追加
  const origSvgMouseMove = onSvgMouseMove
  onSvgMouseMove = function(e) {
    if (state.edgeHandleDrag) {
      const pt = svgPoint(e.clientX, e.clientY)
      tempLayer.innerHTML = ''
      // 从拖拽端点到另一端的临时连线
      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path')
      path.setAttribute('d', makeBezierPath(pt, state.edgeHandleDrag.otherPt))
      path.setAttribute('class', 'temp-edge')
      path.setAttribute('stroke', state.edgeHandleDrag.side === 'source' ? '#378ADD' : '#378ADD')
      tempLayer.appendChild(path)

      // 拖拽端圆点
      const dot = document.createElementNS('http://www.w3.org/2000/svg', 'circle')
      dot.setAttribute('cx', pt.x)
      dot.setAttribute('cy', pt.y)
      dot.setAttribute('r', '7')
      dot.setAttribute('fill', '#378ADD')
      dot.setAttribute('stroke', '#fff')
      dot.setAttribute('stroke-width', '2')
      tempLayer.appendChild(dot)
      return
    }
    origSvgMouseMove(e)
  }

  // ── 画布鼠标释放时处理端点拖拽放落 ───────────────
  const origSvgMouseUp = onSvgMouseUp
  onSvgMouseUp = function(e) {
    if (state.edgeHandleDrag) {
      const pt = svgPoint(e.clientX, e.clientY)
      const portEl = e.target.classList.contains('port') ? e.target : null

      if (portEl) {
        const tgtNodeId = portEl.getAttribute('data-node')
        const tgtPort   = portEl.getAttribute('data-port')
        const hd = state.edgeHandleDrag

        // 重建连线
        const sourceId = hd.side === 'source' ? tgtNodeId : hd.otherNodeId
        const targetId = hd.side === 'source' ? hd.otherNodeId : tgtNodeId
        const sourcePort = hd.side === 'source' ? tgtPort : hd.otherPort
        const targetPort = hd.side === 'source' ? hd.otherPort : tgtPort

        // 不允许自环连线
        if (sourceId !== targetId) {
          const newEdge = {
            id: uid(),
            source: sourceId,
            target: targetId,
            sourcePort: sourcePort || (hd.side === 'source' ? tgtPort : (hd.oldSourcePort || 'right')),
            targetPort: targetPort || (hd.side === 'source' ? (hd.oldTargetPort || 'left') : tgtPort),
            label: ''
          }
          state.edges.push(newEdge)
        }
        saveHistory()
      } else {
        // 鼠标没有落在端口上：保留在断开状态（不重建）
        saveHistory()
      }

      state.edgeHandleDrag = null
      const oldEditId = state.editingEdge
      state.editingEdge = null
      if (oldEditId) stopEdgeFlow(oldEditId)
      tempLayer.innerHTML = ''
      renderAll()
      return
    }
    origSvgMouseUp(e)
  }

  // ── 框选 ─────────────────────────────────────────
  function drawSelectBox() {
    tempLayer.innerHTML = ''
    const { sx, sy, ex, ey } = state.selectBox
    const r = document.createElementNS('http://www.w3.org/2000/svg', 'rect')
    r.setAttribute('x', Math.min(sx, ex))
    r.setAttribute('y', Math.min(sy, ey))
    r.setAttribute('width', Math.abs(ex - sx))
    r.setAttribute('height', Math.abs(ey - sy))
    r.setAttribute('class', 'select-rect')
    tempLayer.appendChild(r)
  }

  function finishSelectBox() {
    const { sx, sy, ex, ey } = state.selectBox
    const minX = Math.min(sx, ex), maxX = Math.max(sx, ex)
    const minY = Math.min(sy, ey), maxY = Math.max(sy, ey)
    state.nodes.forEach(n => {
      if (n.x >= minX && n.x <= maxX && n.y >= minY && n.y <= maxY) {
        state.selectedIds.add(n.id)
      }
    })
    renderNodes()
    updatePropsPanel()
  }

  // ── 工具切换 ─────────────────────────────────────
  function setTool(t) {
    state.tool = t
    document.getElementById('toolSelect').classList.toggle('active', t === 'select')
    document.getElementById('toolConnect').classList.toggle('active', t === 'connect')
    svg.style.cursor = t === 'connect' ? 'crosshair' : ''
  }

  document.getElementById('toolSelect').addEventListener('click', () => setTool('select'))
  document.getElementById('toolConnect').addEventListener('click', () => setTool('connect'))

  // ── 选择管理 ─────────────────────────────────────
  function clearSelection() {
    state.selectedIds.clear()
    state.selectedEdgeIds.clear()
    renderNodes()
    renderEdges()
  }

  function selectAll() {
    state.nodes.forEach(n => state.selectedIds.add(n.id))
    renderNodes()
    updatePropsPanel()
  }

  // ── 删除选中 ─────────────────────────────────────
  function deleteSelected() {
    if (state.selectedIds.size === 0 && state.selectedEdgeIds.size === 0) return
    state.selectedIds.forEach(id => {
      state.nodes = state.nodes.filter(n => n.id !== id)
      state.edges = state.edges.filter(e => e.source !== id && e.target !== id)
    })
    state.selectedEdgeIds.forEach(id => {
      state.edges = state.edges.filter(e => e.id !== id)
    })
    clearSelection()
    saveHistory()
    renderAll()
    updatePropsPanel()
  }

  // ── 右键菜单 ─────────────────────────────────────
  const ctxDetachArea = document.getElementById('ctxDetachArea')

  function showCtxMenu(x, y, ctx) {
    ctxMenu.style.left = x + 'px'
    ctxMenu.style.top  = y + 'px'
    ctxMenu._ctx = ctx

    // 清空旧的拆离选项
    ctxDetachArea.innerHTML = ''

    // 仅节点右键时，生成拆离列表
    if (ctx.type === 'node' && ctx.nodeId) {
      const nid = ctx.nodeId
      const nodeEdges = state.edges.filter(e => e.source === nid || e.target === nid)

      if (nodeEdges.length > 0) {
        const sep = document.createElement('div')
        sep.className = 'ctx-sep'
        ctxDetachArea.appendChild(sep)

        nodeEdges.forEach((edge, i) => {
          const src = getNodeById(edge.source)
          const dst = getNodeById(edge.target)
          const srcLabel = src ? src.label : '?'
          const dstLabel = dst ? dst.label : '?'
          const dir = edge.source === nid ? '→ ' + dstLabel : srcLabel + ' →'
          const item = document.createElement('div')
          item.className = 'ctx-item ctx-detach-item'
          item.textContent = `拆离 ${i + 1}：${dir}`
          item.title = `${srcLabel} → ${dstLabel}${edge.label ? '（' + edge.label + '）' : ''}`

          // 悬浮高亮对应连线
          item.addEventListener('mouseenter', () => {
            highlightEdge(edge.id)
          })
          item.addEventListener('mouseleave', () => {
            unhighlightEdge(edge.id)
          })

          item.addEventListener('click', () => {
            state.edges = state.edges.filter(e => e.id !== edge.id)
            state.selectedEdgeIds.delete(edge.id)
            saveHistory()
            renderAll()
            hideCtxMenu()
          })
          ctxDetachArea.appendChild(item)
        })

        if (nodeEdges.length > 1) {
          const sep2 = document.createElement('div')
          sep2.className = 'ctx-sep'
          ctxDetachArea.appendChild(sep2)

          const allItem = document.createElement('div')
          allItem.className = 'ctx-item ctx-detach-all'
          allItem.textContent = `全部拆除（${nodeEdges.length} 条）`

          allItem.addEventListener('mouseenter', () => {
            nodeEdges.forEach(e => highlightEdge(e.id))
          })
          allItem.addEventListener('mouseleave', () => {
            nodeEdges.forEach(e => unhighlightEdge(e.id))
          })

          allItem.addEventListener('click', () => {
            const ids = new Set(nodeEdges.map(e => e.id))
            state.edges = state.edges.filter(e => !ids.has(e.id))
            ids.forEach(id => state.selectedEdgeIds.delete(id))
            saveHistory()
            renderAll()
            hideCtxMenu()
          })
          ctxDetachArea.appendChild(allItem)
        }
      }
    }

    ctxMenu.classList.remove('hidden')

    // 防止菜单超出屏幕底部
    requestAnimationFrame(() => {
      const rect = ctxMenu.getBoundingClientRect()
      const wh = window.innerHeight
      if (rect.bottom > wh - 8) {
        ctxMenu.style.top = (y - rect.height) + 'px'
      }
    })
  }

  function hideCtxMenu() {
    // 清除拆离高亮
    edgeLayer.querySelectorAll('.edge-path.detach-hover').forEach(el => {
      el.classList.remove('detach-hover')
    })
    ctxMenu.classList.add('hidden')
  }

  document.getElementById('ctxDelete').addEventListener('click', () => { deleteSelected(); hideCtxMenu() })

  document.getElementById('ctxDuplicate').addEventListener('click', () => {
    const newIds = new Set()
    state.selectedIds.forEach(id => {
      const n = getNodeById(id)
      if (!n) return
      const copy = { ...n, id: uid(), x: n.x + 24, y: n.y + 24 }
      state.nodes.push(copy)
      newIds.add(copy.id)
    })
    clearSelection()
    newIds.forEach(id => state.selectedIds.add(id))
    saveHistory()
    renderAll()
    updatePropsPanel()
    hideCtxMenu()
  })

  document.getElementById('ctxBringFront').addEventListener('click', () => {
    state.selectedIds.forEach(id => {
      const idx = state.nodes.findIndex(n => n.id === id)
      if (idx >= 0) state.nodes.push(state.nodes.splice(idx, 1)[0])
    })
    renderAll()
    hideCtxMenu()
  })

  document.getElementById('ctxSendBack').addEventListener('click', () => {
    state.selectedIds.forEach(id => {
      const idx = state.nodes.findIndex(n => n.id === id)
      if (idx >= 0) state.nodes.unshift(state.nodes.splice(idx, 1)[0])
    })
    renderAll()
    hideCtxMenu()
  })

  // ── 拖入节点（从左侧面板） ─────────────────────────
  document.querySelectorAll('.palette-item').forEach(item => {
    item.addEventListener('dragstart', e => {
      e.dataTransfer.setData('node-type', item.getAttribute('data-type'))
    })
  })

  container.addEventListener('dragover', e => e.preventDefault())
  container.addEventListener('drop', e => {
    e.preventDefault()
    const type = e.dataTransfer.getData('node-type')
    if (!type) return
    const pt = svgPoint(e.clientX, e.clientY)
    const def = NODE_TYPES[type]
    const node = {
      id: uid(),
      type,
      x: Math.round(pt.x / 10) * 10,
      y: Math.round(pt.y / 10) * 10,
      label: def.label,
      desc: ''
    }
    state.nodes.push(node)
    saveHistory()
    clearSelection()
    state.selectedIds.add(node.id)
    renderAll()
    updatePropsPanel()
  })

  // ── 适应屏幕 ─────────────────────────────────────
  function fitView() {
    if (state.nodes.length === 0) { resetView(); return }
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity
    state.nodes.forEach(n => {
      const def = NODE_TYPES[n.type]
      minX = Math.min(minX, n.x - def.w / 2)
      minY = Math.min(minY, n.y - def.h / 2)
      maxX = Math.max(maxX, n.x + def.w / 2)
      maxY = Math.max(maxY, n.y + def.h / 2)
    })
    const rect = container.getBoundingClientRect()
    const pad = 60
    const scaleX = (rect.width  - pad * 2) / (maxX - minX)
    const scaleY = (rect.height - pad * 2) / (maxY - minY)
    const scale = Math.min(scaleX, scaleY, 1.5)
    state.view.scale = scale
    state.view.x = rect.width  / 2 - (minX + maxX) / 2 * scale
    state.view.y = rect.height / 2 - (minY + maxY) / 2 * scale
    applyTransform()
  }

  function resetView() {
    state.view = { x: 0, y: 0, scale: 1 }
    applyTransform()
  }

  // ── 小地图 ────────────────────────────────────────
  function updateMinimap() {
    const W = 160, H = 100
    mCtx.clearRect(0, 0, W, H)
    mCtx.fillStyle = '#f5f5f0'
    mCtx.fillRect(0, 0, W, H)

    if (state.nodes.length === 0) return

    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity
    state.nodes.forEach(n => {
      minX = Math.min(minX, n.x - 80)
      minY = Math.min(minY, n.y - 50)
      maxX = Math.max(maxX, n.x + 80)
      maxY = Math.max(maxY, n.y + 50)
    })
    const pad = 10
    const sx = (W - pad * 2) / (maxX - minX || 1)
    const sy = (H - pad * 2) / (maxY - minY || 1)
    const s = Math.min(sx, sy)

    const toMX = x => (x - minX) * s + pad
    const toMY = y => (y - minY) * s + pad

    // 画边
    mCtx.strokeStyle = '#b4b2a9'
    mCtx.lineWidth = 0.5
    state.edges.forEach(e => {
      const src = getNodeById(e.source)
      const dst = getNodeById(e.target)
      if (!src || !dst) return
      mCtx.beginPath()
      mCtx.moveTo(toMX(src.x), toMY(src.y))
      mCtx.lineTo(toMX(dst.x), toMY(dst.y))
      mCtx.stroke()
    })

    // 画节点
    state.nodes.forEach(n => {
      const def = NODE_TYPES[n.type]
      const colors = { start: '#639922', end: '#a32d2d', task: '#378ADD', decision: '#BA7517', subprocess: '#534AB7' }
      mCtx.fillStyle = colors[n.type] || '#888'
      mCtx.beginPath()
      mCtx.arc(toMX(n.x), toMY(n.y), 3, 0, Math.PI * 2)
      mCtx.fill()
    })

    // 视口框
    const rect = container.getBoundingClientRect()
    const vpX = (-state.view.x) / state.view.scale
    const vpY = (-state.view.y) / state.view.scale
    const vpW = rect.width  / state.view.scale
    const vpH = rect.height / state.view.scale
    mCtx.strokeStyle = '#639922'
    mCtx.lineWidth = 1
    mCtx.strokeRect(toMX(vpX), toMY(vpY), vpW * s, vpH * s)
  }

  // ── 属性面板 ─────────────────────────────────────
  function updatePropsPanel() {
    if (state.selectedEdgeIds.size === 1) {
      const eid = Array.from(state.selectedEdgeIds)[0]
      renderEdgeProps(eid)
      return
    }
    if (state.selectedIds.size === 1) {
      const id = Array.from(state.selectedIds)[0]
      renderNodeProps(id)
      return
    }
    if (state.selectedIds.size > 1) {
      propsBody.innerHTML = `<div class="props-empty">已选中 ${state.selectedIds.size} 个节点</div>`
      return
    }
    propsBody.innerHTML = '<div class="props-empty">选中节点或连线以查看属性</div>'
  }

  function renderNodeProps(id) {
    const n = getNodeById(id)
    if (!n) return
    const def = NODE_TYPES[n.type]
    propsBody.innerHTML = `
      <span class="props-badge badge-${n.type}">${def.label}</span>
      <div class="props-field">
        <label>节点名称</label>
        <input type="text" id="pf-label" value="${esc(n.label)}" placeholder="${def.label}">
      </div>
      <div class="props-field">
        <label>描述备注</label>
        <textarea id="pf-desc" placeholder="可选描述">${esc(n.desc || '')}</textarea>
      </div>
      <div class="props-divider"></div>
      <div class="props-field">
        <label>X 坐标</label>
        <input type="number" id="pf-x" value="${n.x}" step="10">
      </div>
      <div class="props-field">
        <label>Y 坐标</label>
        <input type="number" id="pf-y" value="${n.y}" step="10">
      </div>
      <div class="props-divider"></div>
      <div class="props-field">
        <label>节点 ID</label>
        <input type="text" value="${n.id}" readonly style="color:#b4b2a9">
      </div>
    `
    const bind = (sel, key, isNum) => {
      const el = propsBody.querySelector(sel)
      if (!el) return
      el.addEventListener('input', () => {
        n[key] = isNum ? (parseFloat(el.value) || 0) : el.value
        renderAll()
        if (key === 'x' || key === 'y') saveHistory()
      })
      el.addEventListener('change', () => saveHistory())
    }
    bind('#pf-label', 'label')
    bind('#pf-desc',  'desc')
    bind('#pf-x', 'x', true)
    bind('#pf-y', 'y', true)
  }

  function renderEdgeProps(eid) {
    const e = state.edges.find(ed => ed.id === eid)
    if (!e) return
    const src = getNodeById(e.source)
    const dst = getNodeById(e.target)
    propsBody.innerHTML = `
      <span class="props-badge badge-edge">连线</span>
      <div class="props-field">
        <label>连线标签</label>
        <input type="text" id="pf-edge-label" value="${esc(e.label || '')}" placeholder="如：是 / 否">
      </div>
      <div class="props-divider"></div>
      <div class="props-field">
        <label>来源节点</label>
        <input type="text" value="${src ? esc(src.label) : '?'}" readonly style="color:#b4b2a9">
      </div>
      <div class="props-field">
        <label>目标节点</label>
        <input type="text" value="${dst ? esc(dst.label) : '?'}" readonly style="color:#b4b2a9">
      </div>
    `
    propsBody.querySelector('#pf-edge-label').addEventListener('input', el => {
      e.label = el.target.value
      renderEdges()
    })
    propsBody.querySelector('#pf-edge-label').addEventListener('change', () => saveHistory())
  }

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  }

  // ── 文件操作 ─────────────────────────────────────
  async function saveFile() {
    if (!window.electronAPI) return
    const data = {
      version: '1.0',
      nodes: state.nodes,
      edges: state.edges,
      idCounter: state.idCounter,
      view: state.view
    }
    const res = await window.electronAPI.saveFile(data)
    if (res.success) {
      state.fileName = res.path.split(/[\\/]/).pop()
      fileNameEl.textContent = state.fileName
    }
  }

  async function openFile() {
    if (!window.electronAPI) return
    const res = await window.electronAPI.openFile()
    if (res.success) {
      const d = res.data
      state.nodes = d.nodes || []
      state.edges = d.edges || []
      state.idCounter = d.idCounter || 0
      if (d.view) state.view = d.view
      saveHistory()
      renderAll()
      applyTransform()
      clearSelection()
      updatePropsPanel()
    }
  }

  // ── 工具栏按钮 ────────────────────────────────────
  document.getElementById('btnFit').addEventListener('click', fitView)
  document.getElementById('btnClear').addEventListener('click', () => {
    if (!confirm('确认清空画布？此操作可撤销')) return
    state.nodes = []
    state.edges = []
    clearSelection()
    saveHistory()
    renderAll()
    updatePropsPanel()
  })
  document.getElementById('btnSave').addEventListener('click', saveFile)
  document.getElementById('btnOpen').addEventListener('click', openFile)
  document.getElementById('btnFit').addEventListener('click', fitView)
  document.getElementById('btnDeleteSelected').addEventListener('click', deleteSelected)
  document.getElementById('btnUndo').addEventListener('click', undo)
  document.getElementById('btnRedo').addEventListener('click', redo)

  zoomSlider.addEventListener('input', () => {
    const scale = parseInt(zoomSlider.value) / 100
    const rect = container.getBoundingClientRect()
    const cx = rect.width / 2, cy = rect.height / 2
    state.view.x = cx - (cx - state.view.x) * (scale / state.view.scale)
    state.view.y = cy - (cy - state.view.y) * (scale / state.view.scale)
    state.view.scale = scale
    applyTransform()
  })

  // ── 标题栏控件 ────────────────────────────────────
  if (window.electronAPI) {
    document.getElementById('btnMin').addEventListener('click', () => window.electronAPI.windowMinimize())
    document.getElementById('btnMax').addEventListener('click', () => window.electronAPI.windowMaximize())
    document.getElementById('btnClose').addEventListener('click', () => window.electronAPI.windowClose())
    window.electronAPI.onMenuNew(() => {
      if (confirm('新建将清空当前画布，继续？')) {
        state.nodes = []; state.edges = []
        saveHistory(); renderAll(); clearSelection(); updatePropsPanel()
      }
    })
    window.electronAPI.onMenuSave(() => saveFile())
    window.electronAPI.onMenuOpen(() => openFile())
    window.electronAPI.onMenuUndo(() => undo())
    window.electronAPI.onMenuRedo(() => redo())
    window.electronAPI.onMenuFit(() => fitView())
    window.electronAPI.onMenuDelete(() => deleteSelected())
    window.electronAPI.onMenuSelectAll(() => selectAll())
    window.electronAPI.onMenuResetZoom(() => resetView())
  }

  // ── 初始化示例数据 ─────────────────────────────────
  function addSampleFlow() {
    const nodes = [
      { id: uid(), type: 'start',      x: 200, y: 300, label: '开始', desc: '' },
      { id: uid(), type: 'task',       x: 380, y: 220, label: '数据采集', desc: '从数据源获取原始数据' },
      { id: uid(), type: 'task',       x: 380, y: 380, label: '权限验证', desc: '验证用户访问权限' },
      { id: uid(), type: 'decision',   x: 580, y: 300, label: '数据有效？', desc: '' },
      { id: uid(), type: 'task',       x: 760, y: 200, label: '数据处理', desc: '清洗并转换数据格式' },
      { id: uid(), type: 'subprocess', x: 760, y: 380, label: '错误处理', desc: '记录日志并告警' },
      { id: uid(), type: 'end',        x: 960, y: 300, label: '结束', desc: '' }
    ]
    state.nodes = nodes
    const [s, t1, t2, d, t3, t4, e] = nodes
    state.edges = [
      { id: uid(), source: s.id,  sourcePort: 'right',  target: t1.id, targetPort: 'left',   label: '' },
      { id: uid(), source: s.id,  sourcePort: 'right',  target: t2.id, targetPort: 'left',   label: '' },
      { id: uid(), source: t1.id, sourcePort: 'right',  target: d.id,  targetPort: 'left',   label: '' },
      { id: uid(), source: t2.id, sourcePort: 'right',  target: d.id,  targetPort: 'left',   label: '' },
      { id: uid(), source: d.id,  sourcePort: 'top',    target: t3.id, targetPort: 'left',   label: '是' },
      { id: uid(), source: d.id,  sourcePort: 'bottom', target: t4.id, targetPort: 'left',   label: '否' },
      { id: uid(), source: t3.id, sourcePort: 'right',  target: e.id,  targetPort: 'left',   label: '' },
      { id: uid(), source: t4.id, sourcePort: 'right',  target: e.id,  targetPort: 'left',   label: '' }
    ]
    saveHistory()
    renderAll()
    setTimeout(fitView, 100)
  }

  // ── 启动 ─────────────────────────────────────────
  applyTransform()
  addSampleFlow()
  updatePropsPanel()
})()
