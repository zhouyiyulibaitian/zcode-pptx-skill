/* ===========================================================================
 * pptx-studio · 编辑态前端 (assets/editor.js)
 *
 * 由 scripts/lib/html_render.py 内联进 mode=edit 的页面，所以：
 *   - 不是 ES 模块，没有 import/export，不引任何外部资源，没有构建步骤；
 *   - 页面里的 .slide / .el / .tx / .shp / .imw / .tbl 全部由 Python 渲染器
 *     生成，这里只做增量修改，绝不重新实现排版（视觉真相永远在服务端）；
 *   - 单位：几何一律英寸；DOM 里的像素是 96dpi 的 CSS 像素（1in = 96px），
 *     字号 pt -> px 是 4/3；
 *   - 所有修改先在本地排队（queue），保存时一次性 POST /api/patch。
 *
 * 唯一的全局：window.__editor
 * =========================================================================== */
(function () {
  'use strict';

  /* ---------------------------------------------------------------- 常量 */

  var PPI = 96;                      // 1 英寸 = 96 CSS 像素
  var PT2PX = PPI / 72;              // 字号：pt -> px = 4/3
  var R3 = 1000;                     // 英寸保留 3 位小数（和 op 走线格式一致）
  var MIN_BOX = 8;                   // 缩放时元素最小边长（px）
  var BLEED = 0.02;                  // 允许出血的容差（英寸）
  var HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
  var LABEL_SAVE = '保存';
  var LABEL_DIRTY = '保存 *';

  /* ---------------------------------------------------------------- 状态 */

  var queue = [];        // [{op, undo, key}]：待保存的操作；op 就是 /api/patch 的数组元素
  var sel = null;        // 当前选中的元素 id
  var drag = null;       // 拖拽 / 缩放会话
  var curIdx = 0;        // 当前页：最后点击的页，或最靠近视口中心的页
  var saving = false;
  var meta = null;       // GET /api/meta 的结果
  var rawDeck = null;    // GET /api/deck 的原始 deck（只读，仅用于图表数据回填）
  var themeWas = '';     // 主题切换前的值（撤销用）
  var dom = {};
  var handleBox = null, dlg = null;
  var toastTimer = 0, scrollTick = 0;

  /* ---------------------------------------------------------------- 小工具 */

  function byId(id) { return document.getElementById(id); }

  function qsa(sel, root) {
    return Array.prototype.slice.call((root || document).querySelectorAll(sel));
  }

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function clamp(v, lo, hi) { return v < lo ? lo : (v > hi ? hi : v); }
  function r3(v) { return Math.round(v * R3) / R3; }
  function pxNum(inches) { return Math.round(inches * PPI * 100) / 100; }   // 英寸 -> 像素（与渲染器同为 2 位）
  function px2s(inches) { return pxNum(inches) + 'px'; }
  function num(v) { var n = parseFloat(v); return isFinite(n) ? n : 0; }
  function inchOfPx(v) { return r3(num(v) / PPI); }
  function cs(el, prop) { return getComputedStyle(el).getPropertyValue(prop); }
  function trim(s) { return String(s == null ? '' : s).replace(/^\s+|\s+$/g, ''); }
  function nonEmpty(s) { return s !== ''; }

  function closestOf(node, sel) {
    if (!node || node.nodeType !== 1 || !node.closest) return null;
    return node.closest(sel);
  }

  function isTyping(node) {
    if (!node || node.nodeType !== 1) return false;
    if (node.isContentEditable) return true;
    var t = node.tagName;
    return t === 'INPUT' || t === 'TEXTAREA' || t === 'SELECT';
  }

  function slides() { return qsa('.slide'); }

  function elBy(id) {
    if (!id) return null;
    try { return document.querySelector('[data-elid="' + String(id).replace(/["\\]/g, '') + '"]'); }
    catch (e) { return null; }
  }

  function slideOf(node) { return node ? closestOf(node, '.slide') : null; }
  function slideId(node) { var s = slideOf(node); return s ? (s.dataset.id || '') : ''; }
  function slideById(sid) {
    var list = slides();
    for (var i = 0; i < list.length; i++) if (list[i].dataset.id === sid) return list[i];
    return null;
  }
  function slideIndexById(sid) {
    var list = slides();
    for (var i = 0; i < list.length; i++) if (list[i].dataset.id === sid) return i;
    return -1;
  }

  function stageSize() {
    var w = parseFloat(window.__STAGE_W), h = parseFloat(window.__STAGE_H);
    if (!isFinite(w) || w <= 0 || !isFinite(h) || h <= 0) {
      var css = getComputedStyle(document.documentElement);
      w = parseFloat(css.getPropertyValue('--sw')) || 1280;
      h = parseFloat(css.getPropertyValue('--sh')) || 720;
    }
    if (!isFinite(h) || h <= 0) h = 720;
    return { w: w, h: h };
  }

  /* 屏幕像素 -> 版面像素的换算系数。
     版面被外面那套 fit 逻辑用 CSS transform 缩放过，所以必须实测：
     .slide 的实际宽度 / __STAGE_W 就是当前缩放比。拖拽时不做这一步换算，
     窗口小的时候位移就会偏大（这正是 __STAGE_W 存在的原因）。 */
  function screenScale() {
    var first = document.querySelector('.slide');
    if (!first) return 1;
    var w = first.getBoundingClientRect().width;
    var sw = stageSize().w;
    return (w > 0 && sw > 0) ? (w / sw) : 1;
  }

  function geomOf(el) {
    return {
      x: inchOfPx(el.style.left), y: inchOfPx(el.style.top),
      w: inchOfPx(el.style.width), h: inchOfPx(el.style.height)
    };
  }

  /* 把几何写回内联样式（op 的走线格式）。内层盒子（图片/占位/形状/表格/SVG）
     的尺寸是渲染器按元素尺寸算死的，缩放时必须一起改，否则只有外框在动。 */
  function applyGeom(el, g) {
    el.style.left = px2s(g.x);
    el.style.top = px2s(g.y);
    el.style.width = px2s(g.w);
    el.style.height = px2s(g.h);
    var box = el.querySelector('.imw,.ph,.shp,.tbl');
    if (box && box !== el) { box.style.width = px2s(g.w); box.style.height = px2s(g.h); }
    var lsvg = el.querySelector('svg.shp');            // line / arrow：SVG 视口就是元素本身
    if (lsvg) {
      var w = pxNum(g.w), h = pxNum(g.h);
      lsvg.setAttribute('width', w);
      lsvg.setAttribute('height', h);
      lsvg.setAttribute('viewBox', '0 0 ' + w + ' ' + h);
    }
    var csvg = el.querySelector('.el-chart svg');
    if (csvg) { csvg.setAttribute('width', pxNum(g.w)); csvg.setAttribute('height', pxNum(g.h)); }
  }

  /* 夹到画布内：允许 0.02in 出血，宽度不超过画布 + 出血 */
  function clampGeom(g) {
    var s = stageSize(), W = s.w / PPI, H = s.h / PPI;
    g.w = Math.min(Math.max(r3(g.w), MIN_BOX / PPI), W + 2 * BLEED);
    g.h = Math.min(Math.max(r3(g.h), MIN_BOX / PPI), H + 2 * BLEED);
    g.x = clamp(r3(g.x), -BLEED, Math.max(-BLEED, W + BLEED - g.w));
    g.y = clamp(r3(g.y), -BLEED, Math.max(-BLEED, H + BLEED - g.h));
    return g;
  }

  /* ---------------------------------------------------------------- 提示条 */

  function toast(msg, ms) {
    if (!dom.toast) return;
    dom.toast.textContent = String(msg == null ? '' : msg);
    dom.toast.classList.add('on');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { dom.toast.classList.remove('on'); }, ms || 1800);
  }

  /* 长提示（导出路径）：8s，路径部分可选中复制 */
  function toastPath(prefix, path) {
    if (!dom.toast) return;
    dom.toast.textContent = '';
    dom.toast.appendChild(document.createTextNode(prefix));
    var sp = document.createElement('span');
    sp.className = 'pick';
    sp.textContent = String(path || '');
    dom.toast.appendChild(sp);
    dom.toast.title = String(path || '');
    dom.toast.classList.add('on');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { dom.toast.classList.remove('on'); dom.toast.title = ''; }, 8000);
  }

  /* ---------------------------------------------------------------- 网络 */
  /* 统一封装：永不 reject，失败也返回 {ok:false,error}，调用方必须处理，
     不允许出现“看起来保存成功”的情况。 */
  function api(url, body) {
    var init = { method: body === undefined ? 'GET' : 'POST' };
    if (body !== undefined) {
      init.headers = { 'Content-Type': 'application/json' };
      init.body = JSON.stringify(body);
    }
    return fetch(url, init).then(function (r) {
      return r.text().then(function (txt) {
        var j = null;
        try { j = JSON.parse(txt); } catch (e) { j = null; }
        if (!j || typeof j !== 'object') {
          return { ok: false, error: r.ok ? '服务端返回的不是 JSON' : ('HTTP ' + r.status) };
        }
        return j;
      });
    }).catch(function (err) {
      return { ok: false, error: '网络错误：' + ((err && err.message) || err) };
    });
  }

  /* ---------------------------------------------------------------- 队列 */

  function mark() {
    var dirty = queue.length > 0;
    if (dom.save) dom.save.textContent = dirty ? LABEL_DIRTY : LABEL_SAVE;
    var flag = byId('ed-dirty');
    if (flag) flag.className = dirty ? 'dirty on' : 'dirty';
  }

  function mergeOp(into, from) {
    if (from.op === 'geometry') { into.x = from.x; into.y = from.y; into.w = from.w; into.h = from.h; }
    else if (from.op === 'text') { into.paragraphs = from.paragraphs; }
    else if (from.op === 'style') { if (from.set) for (var k in from.set) into.set[k] = from.set[k]; }
    else if (from.op === 'notes') { into.text = from.text; }
    else if (from.op === 'image') { into.src = from.src; }
  }

  /* 入队。同 key 且间隔很短的相邻两条合并成一条（连按方向键、连续拖色盘、
     连续输入讲稿），合并时保留最早那条的 undo —— 它记录的是这串修改之前的状态。
     加时间窗是为了撤销粒度：拖完隔一会儿再微调，应该算两次操作。 */
  var COALESCE_MS = 1200;
  function pushOp(op, undoFn, key) {
    var last = queue.length ? queue[queue.length - 1] : null;
    var now = Date.now();
    if (key && last && last.key === key && (now - last.ts) < COALESCE_MS) {
      mergeOp(last.op, op);
      last.ts = now;                       // 连续手势不断续期
      return;
    }
    queue.push({ op: op, undo: undoFn || null, key: key || '', ts: now });
    mark();
  }

  function opLabel(op) {
    if (!op) return '操作';
    if (op.op === 'geometry') return '移动/缩放';
    if (op.op === 'text') return '文字';
    if (op.op === 'style') return op.set && op.set.chart ? '图表数据' : '样式';
    if (op.op === 'notes') return '讲稿';
    if (op.op === 'image') return '图片';
    if (op.op === 'deck') return '主题';
    if (op.op === 'slide') {
      return ({ add: '新增页', delete: '删除页', duplicate: '复制页', move: '移动页', layout: '版式' })[op.action] || '页面';
    }
    if (op.op === 'element') {
      return ({ add: '新增元素', delete: '删除元素', duplicate: '复制元素' })[op.action] || '元素';
    }
    return op.op || '操作';
  }

  /* 撤销：只对本次未保存的队列有效，能把 DOM 恢复到入队前的样子 */
  function undoLast() {
    if (!queue.length) { toast('没有可撤销的修改'); return; }
    var rec = queue.pop();
    if (rec.undo) {
      try { rec.undo(); } catch (err) { console.warn('[editor] 撤销时出错', err); }
    }
    if (sel && !elBy(sel)) sel = null;
    mark();
    refreshUI();
    toast('已撤销：' + opLabel(rec.op));
  }

  /* ---------------------------------------------------------------- 选中 */

  function clearHandles() {
    if (handleBox && handleBox.parentNode) handleBox.parentNode.removeChild(handleBox);
  }

  function positionHandles(el) {
    if (!handleBox || !el) return;
    var slide = slideOf(el);
    if (!slide) { clearHandles(); return; }
    if (handleBox.parentNode !== slide) slide.appendChild(handleBox);
    var g = geomOf(el);
    handleBox.style.left = px2s(g.x);
    handleBox.style.top = px2s(g.y);
    handleBox.style.width = px2s(g.w);
    handleBox.style.height = px2s(g.h);
    handleBox.style.transform = el.style.transform || 'none';   // 元素带旋转时手柄一起转
    var s = screenScale();
    handleBox.style.setProperty('--k', String(s > 0 ? 1 / s : 1)); // 手柄按屏幕尺寸显示，不随版面缩放
  }

  function buildHandles(el) {
    if (!handleBox) {
      handleBox = document.createElement('div');
      handleBox.id = '__handles';
    }
    if (!handleBox.childNodes.length) {
      handleBox.innerHTML = HANDLES.map(function (h) {
        return '<div class="h h-' + h + '" data-h="' + h + '" title="缩放"></div>';
      }).join('');
    }
    positionHandles(el);
  }

  function select(id) {
    if (sel && sel !== id) {
      var prev = elBy(sel);
      if (prev) prev.classList.remove('sel');
    }
    sel = null;
    clearHandles();
    var el = id ? elBy(id) : null;
    if (el && el.classList.contains('locked')) el = null;   // 背景装饰不可选
    if (el) {
      el.classList.add('sel');
      sel = el.dataset.elid;
      buildHandles(el);
      var i = slideIndexById(slideId(el));
      if (i >= 0 && i !== curIdx) { curIdx = i; updateHud(); syncLayoutSel(); }
    }
    renderPanel();
  }

  function refreshUI() {
    if (sel) select(sel);
    else { clearHandles(); renderPanel(); }
  }

  /* ---------------------------------------------------------------- 拖拽 */

  function beginDrag(e, el, handle) {
    var g0 = geomOf(el);
    drag = {
      el: el, id: el.dataset.elid, slide: slideId(el), handle: handle || '',
      sx: e.clientX, sy: e.clientY, sc: screenScale(), g0: g0, g: g0, moved: false
    };
    try { if (el.setPointerCapture) el.setPointerCapture(e.pointerId); } catch (err) { /* 忽略 */ }
    document.body.classList.add('ed-drag');
    clearOvf();
    e.preventDefault();     // 别让浏览器顺手开始选文字
  }

  function moveGeomTo(d, dx, dy) {
    return clampGeom({ x: d.g0.x + dx, y: d.g0.y + dy, w: d.g0.w, h: d.g0.h });
  }

  function resizeGeomTo(d, dx, dy) {
    var h = d.handle, minW = MIN_BOX / PPI, minH = MIN_BOX / PPI;
    var g = { x: d.g0.x, y: d.g0.y, w: d.g0.w, h: d.g0.h };
    if (h.indexOf('e') >= 0) g.w = d.g0.w + dx;
    if (h.indexOf('s') >= 0) g.h = d.g0.h + dy;
    if (h.indexOf('w') >= 0) { g.w = d.g0.w - dx; g.x = d.g0.x + dx; }
    if (h.indexOf('n') >= 0) { g.h = d.g0.h - dy; g.y = d.g0.y + dy; }
    if (g.w < minW) { if (h.indexOf('w') >= 0) g.x = d.g0.x + d.g0.w - minW; g.w = minW; }
    if (g.h < minH) { if (h.indexOf('n') >= 0) g.y = d.g0.y + d.g0.h - minH; g.h = minH; }
    return clampGeom(g);
  }

  function onPointerDown(e) {
    if (e.button !== undefined && e.button !== 0) return;      // 只处理左键
    if (e.target && e.target.isContentEditable) return;        // 编辑文字时不打断

    var hEl = closestOf(e.target, '#__handles .h');
    if (hEl && sel) {
      var hel = elBy(sel);
      if (hel) { beginDrag(e, hel, hEl.dataset.h || ''); return; }   // 手柄拖拽 = 缩放，不当作移动
    }
    var el = closestOf(e.target, '.el');
    if (!el) {
      var sl = closestOf(e.target, '.slide');
      if (sl) {                                                // 点幻灯片空白 -> 当前页 + 取消选中
        var i = slideIndexById(sl.dataset.id);
        if (i >= 0) setCurrent(i, false);
        select(null);
      }
      return;
    }
    if (el.classList.contains('locked')) { select(null); return; }
    select(el.dataset.elid);
    beginDrag(e, el, '');
  }

  function onPointerMove(e) {
    if (!drag) return;
    if (!drag.moved && (Math.abs(e.clientX - drag.sx) > 2 || Math.abs(e.clientY - drag.sy) > 2)) drag.moved = true;
    if (!drag.moved) return;
    var k = drag.sc || 1;
    var dx = (e.clientX - drag.sx) / k / PPI;      // 屏幕像素 -> 版面像素 -> 英寸
    var dy = (e.clientY - drag.sy) / k / PPI;
    var g = drag.handle ? resizeGeomTo(drag, dx, dy) : moveGeomTo(drag, dx, dy);
    if (g.x === drag.g.x && g.y === drag.g.y && g.w === drag.g.w && g.h === drag.g.h) return;
    drag.g = g;
    applyGeom(drag.el, g);
    positionHandles(drag.el);
    syncGeomInputs(g);
    e.preventDefault();
  }

  function onPointerUp(e) {
    if (!drag) return;
    var d = drag;
    drag = null;
    document.body.classList.remove('ed-drag');
    try { if (d.el.releasePointerCapture) d.el.releasePointerCapture(e.pointerId); } catch (err) { /* 忽略 */ }
    var g = d.g;
    if (d.moved && !(g.x === d.g0.x && g.y === d.g0.y && g.w === d.g0.w && g.h === d.g0.h)) {
      pushOp(
        { op: 'geometry', slide: d.slide, el: d.id, x: g.x, y: g.y, w: g.w, h: g.h },
        function () { applyGeom(d.el, d.g0); positionHandles(d.el); },
        'geom:' + d.slide + ':' + d.id
      );
    }
    positionHandles(d.el);
    syncGeomInputs(geomOf(d.el));
  }

  /* 键盘微调复用同一条 geometry op */
  function nudge(dx, dy) {
    var el = sel ? elBy(sel) : null;
    if (!el) return;
    var g0 = geomOf(el);
    var g = clampGeom({ x: g0.x + dx, y: g0.y + dy, w: g0.w, h: g0.h });
    if (g.x === g0.x && g.y === g0.y) return;
    applyGeom(el, g);
    pushOp(
      { op: 'geometry', slide: slideId(el), el: el.dataset.elid, x: g.x, y: g.y, w: g.w, h: g.h },
      function () { applyGeom(el, g0); },
      'geom:' + slideId(el) + ':' + el.dataset.elid
    );
    positionHandles(el);
    syncGeomInputs(g);
  }

  function syncGeomInputs(g) {
    ['x', 'y', 'w', 'h'].forEach(function (k) {
      var inp = byId('ed-' + k);
      if (inp && document.activeElement !== inp) inp.value = g[k];
    });
  }

  function commitGeom(el, g) {
    var before = geomOf(el);
    if (before.x === g.x && before.y === g.y && before.w === g.w && before.h === g.h) return;
    applyGeom(el, g);
    pushOp(
      { op: 'geometry', slide: slideId(el), el: el.dataset.elid, x: g.x, y: g.y, w: g.w, h: g.h },
      function () { applyGeom(el, before); },
      'geom:' + slideId(el) + ':' + el.dataset.elid
    );
    positionHandles(el);
  }

  /* ---------------------------------------------------------------- 文字编辑 */

  function startTextEdit(tx) {
    if (tx.isContentEditable) return;
    var el = closestOf(tx, '.el');
    if (el && el.classList.contains('locked')) { toast('这是锁定的背景元素，不能编辑'); return; }
    if (el) select(el.dataset.elid);
    try { document.execCommand('defaultParagraphSeparator', false, 'p'); } catch (e) { /* 老内核忽略 */ }
    tx.__edOrig = tx.innerHTML;
    tx.contentEditable = 'true';
    tx.spellcheck = false;
    tx.focus();
    var r = document.createRange();
    r.selectNodeContents(tx);
    var s = window.getSelection();
    if (s) { s.removeAllRanges(); s.addRange(r); }
  }

  function collectParagraphs(root) {
    var norm = function (s) {
      return String(s == null ? '' : s).replace(/\u00a0/g, ' ').replace(/\r\n?/g, '\n').replace(/\s+$/, '');
    };
    var ps = qsa('p', root);
    if (!ps.length) ps = qsa('div', root);          // Chromium 有时插入 div 而不是 p
    if (!ps.length) {
      var t = norm(root.innerText);
      return [{ runs: [{ text: t }] }];
    }
    return ps.map(function (p) { return { runs: [{ text: norm(p.innerText) }] }; });
  }

  function endTextEdit(tx) {
    if (!tx || !tx.isContentEditable) return;
    tx.contentEditable = 'false';
    var orig = tx.__edOrig;
    tx.__edOrig = null;
    if (orig == null || orig === tx.innerHTML) return;    // 没真改就不发 op
    var el = closestOf(tx, '.el');
    if (!el) return;
    pushOp(
      { op: 'text', slide: slideId(el), el: el.dataset.elid, paragraphs: collectParagraphs(tx) },
      function () { tx.innerHTML = orig; },
      ''
    );
    toast('文字已改，保存后生效');
  }

  /* ---------------------------------------------------------------- 文字样式 */

  /* 段落 p 上的内联样式压不过 run（span）上的内联样式，所以两边一起改，
     否则面板里改了颜色/字号会“没反应”。真正的落库由服务端按 op 重算。 */
  function setTextProp(el, prop, val) {
    var tx = el.querySelector('.tx');
    if (!tx) return;
    qsa('p', tx).forEach(function (p) {
      var nodes = [p].concat(qsa('span', p));
      nodes.forEach(function (n) {
        if (prop === 'size') n.style.fontSize = (val * PT2PX) + 'px';
        else if (prop === 'color') n.style.color = val;
        else if (prop === 'bold') n.style.fontWeight = val ? '700' : '400';
      });
      if (prop === 'align') p.style.textAlign = val;
      if (prop === 'lineSpacing') p.style.lineHeight = String(val);
    });
  }

  function pushStyle(el, key, set) {
    var tx = el.querySelector('.tx');
    var snap = tx ? tx.innerHTML : null;      // 快照必须在改 DOM 之前取
    pushOp(
      { op: 'style', slide: slideId(el), el: el.dataset.elid, set: set },
      snap === null ? null : function () { if (tx) tx.innerHTML = snap; },
      'style:' + slideId(el) + ':' + el.dataset.elid + ':' + key
    );
  }

  function toHex(color) {
    var m = String(color || '').match(/\d+(\.\d+)?/g);
    if (!m || m.length < 3) return '#111111';
    var out = '#';
    for (var i = 0; i < 3; i++) {
      var v = clamp(Math.round(parseFloat(m[i])), 0, 255).toString(16);
      out += (v.length < 2 ? '0' : '') + v;
    }
    return out;
  }

  function normAlign(v) {
    var a = trim(v);
    if (a === 'start' || a === 'justify-all') return a === 'start' ? 'left' : 'justify';
    if (a === 'end') return 'right';
    if (a === 'center' || a === 'right' || a === 'justify' || a === 'left') return a;
    return 'left';
  }

  function lineHeightOf(p) {
    var v = trim(cs(p, 'line-height'));
    if (!v || v === 'normal') return 1.35;
    var n = parseFloat(v);
    if (!isFinite(n)) return 1.35;
    if (v.indexOf('px') >= 0) {
      var fs = parseFloat(cs(p, 'font-size')) || 16;
      n = fs > 0 ? n / fs : 1.35;
    }
    return Math.round(n * 100) / 100;
  }

  /* ---------------------------------------------------------------- 图片 */

  function setImageSrc(el, url) {
    var img = el.querySelector('img');
    if (img) { img.src = url; return; }
    var imw = el.querySelector('.imw');
    if (!imw) {
      imw = document.createElement('div');
      imw.className = 'imw';
      imw.style.cssText = 'width:' + px2s(geomOf(el).w) + ';height:' + px2s(geomOf(el).h);
      var ph = el.querySelector('.ph');
      if (ph && ph.parentNode) ph.parentNode.replaceChild(imw, ph);
      else el.appendChild(imw);
    }
    imw.innerHTML = '';
    var ni = document.createElement('img');
    ni.src = url;
    ni.style.cssText = 'object-fit:cover;width:100%;height:100%;display:block';
    imw.appendChild(ni);
  }

  function uploadImage(input) {
    var f = input.files && input.files[0];
    input.value = '';
    if (!f) return;
    var el = sel ? elBy(sel) : null;
    if (!el || el.dataset.type !== 'image') { toast('请先选中一个图片元素'); return; }
    var fd = new FormData();
    fd.append('file', f);
    toast('图片上传中…', 1200);
    fetch('/api/upload', { method: 'POST', body: fd })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (!j || !j.ok) { toast('上传失败：' + ((j && j.error) || '未知错误'), 5000); return; }
        var snap = el.innerHTML;
        pushOp(
          { op: 'image', slide: slideId(el), el: el.dataset.elid, src: j.path },
          function () { el.innerHTML = snap; },
          ''
        );
        setImageSrc(el, j.url || j.path);
        toast('图片已替换，保存后生效');
      })
      .catch(function (err) { toast('上传失败：' + ((err && err.message) || err), 5000); });
  }

  /* ---------------------------------------------------------------- 图表 */

  function rawSlide(sid) {
    if (!rawDeck || !rawDeck.slides) return null;
    for (var i = 0; i < rawDeck.slides.length; i++) {
      if (rawDeck.slides[i] && rawDeck.slides[i].id === sid) return rawDeck.slides[i];
    }
    return null;
  }

  /* 图表数据在 deck 里有两种落点：显式元素的 categories/series，
     或版式页的 content.chart.categories/series。 */
  function chartSeed(el) {
    var sid = slideId(el), eid = el.dataset.elid;
    var raw = rawSlide(sid);
    if (raw) {
      var els = raw.elements || [];
      for (var i = 0; i < els.length; i++) {
        if (els[i] && els[i].id === eid && (els[i].categories || els[i].series)) {
          return { categories: els[i].categories || [], series: els[i].series || [] };
        }
      }
      var c = (raw.content && raw.content.chart) || null;
      if (c && (c.categories || c.series)) return { categories: c.categories || [], series: c.series || [] };
      if (raw.categories || raw.series) return { categories: raw.categories || [], series: raw.series || [] };
    }
    return { categories: [], series: [{ name: '系列 1', values: [] }] };
  }

  function ensureDialog() {
    if (dlg) return dlg;
    var wrap = document.createElement('div');
    wrap.id = '__dlg';
    wrap.innerHTML =
      '<div class="box"><h4>编辑图表数据</h4>' +
      '<label class="f col"><span>类别</span><input id="__cat" type="text" placeholder="一季度, 二季度, 三季度"></label>' +
      '<p class="hint">用逗号分隔；类别个数必须和每个系列的数值个数一致。</p>' +
      '<label class="f col"><span>系列</span><textarea id="__ser" rows="6" spellcheck="false"></textarea></label>' +
      '<p class="hint">每行一个系列，格式：系列名: 1, 2, 3</p>' +
      '<div class="row"><button class="btn pri" id="__apply">应用</button>' +
      '<button class="btn" id="__cancel">取消</button></div></div>';
    document.body.appendChild(wrap);
    wrap.addEventListener('click', function (e) {
      if (e.target === wrap) closeDialog();
      else if (e.target.id === '__apply') applyDialog();
      else if (e.target.id === '__cancel') closeDialog();
    });
    dlg = wrap;
    return dlg;
  }

  function openChartDialog(el) {
    var d = ensureDialog();
    var seed = chartSeed(el);
    d.dataset.slide = slideId(el);
    d.dataset.el = el.dataset.elid;
    byId('__cat').value = (seed.categories || []).join(', ');
    byId('__ser').value = (seed.series || []).map(function (s) {
      return (s.name || '系列') + ': ' + (s.values || []).join(', ');
    }).join('\n');
    d.classList.add('on');
    byId('__ser').focus();
  }

  function closeDialog() { if (dlg) dlg.classList.remove('on'); }

  function parseChartText(catText, serText) {
    var cats = String(catText || '').split(/[，,;；\t]/).map(trim).filter(nonEmpty);
    if (!cats.length) return { error: '请先填写类别' };
    var lines = String(serText || '').split(/\r?\n/).map(trim).filter(nonEmpty);
    if (!lines.length) return { error: '请至少填写一个系列' };
    var series = [];
    for (var i = 0; i < lines.length; i++) {
      var ln = lines[i], cut = ln.search(/[:：]/);
      var name = cut >= 0 ? trim(ln.slice(0, cut)) : ('系列 ' + (i + 1));
      var toks = (cut >= 0 ? ln.slice(cut + 1) : ln).split(/[，,;；\s]+/).map(trim).filter(nonEmpty);
      if (!name) name = '系列 ' + (i + 1);
      var vals = [];
      for (var k = 0; k < toks.length; k++) {
        var v = Number(toks[k]);
        if (!isFinite(v)) return { error: '无法识别的数值：' + toks[k] };
        vals.push(v);
      }
      if (vals.length !== cats.length) {
        return { error: '系列“' + name + '”有 ' + vals.length + ' 个数值，类别有 ' + cats.length + ' 个' };
      }
      series.push({ name: name, values: vals });
    }
    return { categories: cats, series: series };
  }

  function applyDialog() {
    if (!dlg) return;
    var parsed = parseChartText(byId('__cat').value, byId('__ser').value);
    if (parsed.error) { toast(parsed.error, 5000); return; }
    pushOp({
      op: 'style', slide: dlg.dataset.slide, el: dlg.dataset.el,
      set: { chart: { categories: parsed.categories, series: parsed.series } }
    }, null, '');
    toast('图表数据已排队，保存后重新生成');
    closeDialog();
  }

  /* ---------------------------------------------------------------- 元素操作 */

  function deleteElement(el) {
    var parent = el.parentNode, next = el.nextSibling;
    var slide = slideId(el), id = el.dataset.elid;
    pushOp({ op: 'element', action: 'delete', slide: slide, el: id }, function () {
      if (parent) parent.insertBefore(el, next);          // 撤销 = 放回原位
    }, '');
    if (sel === id) { sel = null; clearHandles(); }
    el.remove();
    renderPanel();
    toast('已删除，保存后生效');
  }

  function deleteSel() {
    var el = sel ? elBy(sel) : null;
    if (!el) return;
    deleteElement(el);
  }

  function duplicateSel() {
    var el = sel ? elBy(sel) : null;
    if (!el) return;
    pushOp({ op: 'element', action: 'duplicate', slide: slideId(el), el: el.dataset.elid }, null, '');
    toast('已复制，保存后生效');
  }

  function addTextBox() {
    var s = slides()[curIdx];
    if (!s) return;
    var S = stageSize(), W = S.w / PPI, H = S.h / PPI;
    pushOp({
      op: 'element', action: 'add', slide: s.dataset.id,
      element: {
        type: 'text',
        x: clamp(1.2, 0, Math.max(0, W - 4)),
        y: clamp(1.2, 0, Math.max(0, H - 0.8)),
        w: 4, h: 0.8,
        paragraphs: [{ runs: [{ text: '新文本框' }], size: 18, color: '#111111' }]
      }
    }, null, '');
    toast('已添加文本框，保存后生效');
  }

  /* ---------------------------------------------------------------- 页面操作 */

  function addSlide() {
    pushOp({ op: 'slide', action: 'add', index: curIdx }, null, '');
    toast('已在本页后新增一页，保存后生效');
  }

  function duplicateSlide() {
    pushOp({ op: 'slide', action: 'duplicate', index: curIdx }, null, '');
    toast('已复制本页，保存后生效');
  }

  function deleteSlide() {
    if (slides().length <= 1) { toast('至少要保留一页'); return; }
    if (!window.confirm('删除第 ' + (curIdx + 1) + ' 页？保存后生效（可用 Ctrl+Z 撤销队列里的操作）')) return;
    pushOp({ op: 'slide', action: 'delete', index: curIdx }, null, '');
    toast('已删除本页，保存后生效');
  }

  function setNotes(v) {
    var s = slides()[curIdx];
    if (!s) return;
    var sid = s.dataset.id, before = s.dataset.notes || '';
    if (before === v) return;
    s.dataset.notes = v;                        // 让面板重绘/撤销都有一致的来源
    pushOp({ op: 'notes', slide: sid, text: v }, function () {
      var t = slideById(sid);
      if (t) t.dataset.notes = before;
      syncNotesBox();
    }, 'notes:' + sid);
  }

  function syncNotesBox() {
    var box = byId('ed-notes');
    if (!box || document.activeElement === box) return;
    var s = slides()[curIdx];
    box.value = s ? (s.dataset.notes || '') : '';
  }

  function nearestIndex() {
    var list = slides();
    if (!list.length) return 0;
    var cy = (window.innerHeight || 800) / 2, best = 0, bd = Infinity;
    for (var i = 0; i < list.length; i++) {
      var r = list[i].getBoundingClientRect();
      var d = Math.abs((r.top + r.height / 2) - cy);
      if (d < bd) { bd = d; best = i; }
    }
    return best;
  }

  function updateHud() {
    if (!dom.pageinfo) return;
    var n = slides().length || parseInt(window.__PAGE_COUNT, 10) || 0;
    dom.pageinfo.textContent = (curIdx + 1) + ' / ' + n;
  }

  function setCurrent(i, fromScroll) {
    var n = slides().length;
    if (!n) return;
    i = clamp(i | 0, 0, n - 1);
    var changed = i !== curIdx;
    curIdx = i;
    updateHud();
    if (!changed) return;
    syncLayoutSel();
    /* 滚动时若正在输入（讲稿框），不要重绘面板打断输入 */
    if (!sel && !(fromScroll && isTyping(document.activeElement))) renderPanel();
  }

  function onScroll() {
    if (scrollTick) return;
    scrollTick = window.requestAnimationFrame(function () {
      scrollTick = 0;
      setCurrent(nearestIndex(), true);
    });
  }

  /* ---------------------------------------------------------------- 版式 / 主题 */

  function shortName(desc, id) {
    var s = trim(String(desc || '').split(/[：:]/)[0]);
    return s || id;
  }

  function populateLayouts() {
    var lo = dom.layout;
    if (!lo) return;
    var layouts = (meta && meta.layouts) || {};
    var html = ['<option value="">（保持当前）</option>'];
    Object.keys(layouts).forEach(function (k) {
      html.push('<option value="' + esc(k) + '" title="' + esc(layouts[k]) + '">' +
        esc(k + ' · ' + shortName(layouts[k], k)) + '</option>');
    });
    lo.innerHTML = html.join('');
    syncLayoutSel();
  }

  function syncLayoutSel() {
    var lo = dom.layout;
    if (!lo) return;
    var s = slides()[curIdx];
    var want = (s && s.dataset.layout) || '';
    var has = false;
    for (var i = 0; i < lo.options.length; i++) if (lo.options[i].value === want && want) { has = true; break; }
    lo.value = has ? want : '';
  }

  function onLayoutChange() {
    var lo = dom.layout;
    if (!lo) return;
    var v = lo.value;
    var s = slides()[curIdx];
    if (!v || !s) { syncLayoutSel(); return; }
    var raw = rawSlide(s.dataset.id);
    var before = s.dataset.layout || '';
    var op = { op: 'slide', action: 'layout', index: curIdx, layout: v };
    if (raw && raw.content) op.content = raw.content;    // 版式页：带上原内容，换版式不丢内容
    pushOp(op, function () { s.dataset.layout = before; syncLayoutSel(); }, '');
    s.dataset.layout = v;
    if (raw && raw.elements && raw.elements.length) {
      // 这一页已经是独立元素页（服务端 detach 过），套版式会按新布局重建元素列表
      toast('本页已是独立元素页：换版式会重建这一页的元素，本页的改动会丢失', 5000);
    } else {
      toast('版式已排队，保存后生效');
    }
  }

  function onThemeChange() {
    var th = dom.theme;
    if (!th) return;
    var v = th.value, before = themeWas;
    themeWas = v;
    pushOp({ op: 'deck', set: { theme: v } }, function () { th.value = before; themeWas = before; }, '');
    toast('主题已排队，保存后生效');
  }

  function populateThemes() {
    var th = dom.theme;
    if (!th || th.options.length || !meta || !meta.themes) return;
    var html = meta.themes.map(function (t) {
      return '<option value="' + esc(t.id) + '" title="' + esc(t.desc) + '">' + esc(t.name || t.id) + '</option>';
    });
    th.innerHTML = html.join('');
    themeWas = th.value;
  }

  /* ---------------------------------------------------------------- 属性面板 */

  function typeName(el) {
    var t = el.dataset.type;
    return ({ text: '文本框', shape: '形状', image: '图片', table: '表格', chart: '图表' })[t] || t || '元素';
  }

  function field(label, inner) {
    return '<label class="f"><span>' + label + '</span>' + inner + '</label>';
  }

  function geomFieldsHTML(g) {
    return '<div class="grid4">' +
      field('X', '<input id="ed-x" type="number" step="0.01" value="' + g.x + '">') +
      field('Y', '<input id="ed-y" type="number" step="0.01" value="' + g.y + '">') +
      field('宽', '<input id="ed-w" type="number" step="0.01" value="' + g.w + '">') +
      field('高', '<input id="ed-h" type="number" step="0.01" value="' + g.h + '">') +
      '</div>';
  }

  function textFieldsHTML(el) {
    var tx = el.querySelector('.tx');
    var p = tx ? tx.querySelector('p') : null;
    var size = p ? Math.round(parseFloat(cs(p, 'font-size')) * 0.75 * 10) / 10 : 18;
    var color = p ? toHex(cs(p, 'color')) : '#111111';
    var align = p ? normAlign(cs(p, 'text-align')) : 'left';
    var bold = p ? (parseInt(cs(p, 'font-weight'), 10) >= 600) : false;
    var lh = p ? lineHeightOf(p) : 1.35;
    if (!isFinite(size) || size <= 0) size = 18;
    return '<div class="sect"><h4>文字</h4>' +
      field('字号 pt', '<input id="ed-size" type="number" step="0.5" min="1" value="' + size + '">') +
      field('颜色', '<input id="ed-color" type="color" value="' + esc(color) + '">') +
      field('对齐', '<select id="ed-align">' +
        ['left:左', 'center:居中', 'right:右', 'justify:两端'].map(function (o) {
          var kv = o.split(':');
          return '<option value="' + kv[0] + '"' + (align === kv[0] ? ' selected' : '') + '>' + kv[1] + '</option>';
        }).join('') + '</select>') +
      field('加粗', '<input id="ed-bold" type="checkbox"' + (bold ? ' checked' : '') + '>') +
      field('行距', '<input id="ed-lh" type="number" step="0.05" min="0.5" value="' + lh + '">') +
      '<p class="hint">双击文字可直接编辑。编辑某一段会用纯文本替换该段的行内样式（run），段落级样式由服务端保留。</p>' +
      '</div>';
  }

  function elementSectionHTML(el) {
    var h = [];
    h.push('<div class="hd"><b>' + esc(typeName(el)) + '</b><code>' + esc(el.dataset.elid) + '</code>' +
      '<span id="ed-dirty" class="dirty">未保存</span></div>');
    h.push('<div class="sect"><h4>位置与尺寸（英寸）</h4>' + geomFieldsHTML(geomOf(el)) +
      '<p class="hint">拖动移动，八个小方块缩放；方向键 0.02in，Shift+方向键 0.1in。</p></div>');
    if (el.dataset.type === 'text' || el.querySelector('.tx')) h.push(textFieldsHTML(el));
    if (el.dataset.type === 'image') {
      h.push('<div class="sect"><h4>图片</h4>' +
        '<button class="btn" id="ed-imgpick">替换图片…</button>' +
        '<input id="ed-imgfile" type="file" accept="image/*" hidden>' +
        '<p class="hint">上传后立即替换预览，保存后写入 deck。</p></div>');
    }
    if (el.dataset.type === 'chart') {
      h.push('<div class="sect"><h4>图表数据</h4>' +
        '<button class="btn" id="ed-chartedit">编辑数据…</button>' +
        '<p class="hint">每行一个系列（系列名: 1, 2, 3），类别与数值个数要一致。</p></div>');
    }
    h.push('<div class="sect"><h4>元素</h4><div class="row">' +
      '<button class="btn" id="ed-dup">复制元素</button>' +
      '<button class="btn danger" id="ed-del">删除元素</button></div></div>');
    return h.join('');
  }

  function slideSectionHTML() {
    var list = slides(), s = list[curIdx] || list[0];
    var h = [];
    h.push('<div class="hd"><b>第 ' + (curIdx + 1) + ' / ' + list.length + ' 页</b>' +
      '<code>' + esc(s ? (s.dataset.id || '') : '') + '</code>' +
      '<span id="ed-dirty" class="dirty">未保存</span></div>');
    h.push('<div class="sect"><h4>本页</h4><div class="row">' +
      '<button class="btn" id="ed-addslide">新增页</button>' +
      '<button class="btn" id="ed-dupslide">复制本页</button>' +
      '<button class="btn danger" id="ed-delslide">删除本页</button></div>' +
      '<p class="hint">新增页插在本页之后。选中元素后这里会变成元素属性。</p></div>');
    h.push('<div class="sect"><h4>讲稿</h4><textarea id="ed-notes" rows="5" placeholder="这一页的讲稿…">' +
      esc(s ? (s.dataset.notes || '') : '') + '</textarea>' +
      '<p class="hint">只写入 deck 的 notes 字段，不会出现在幻灯片画面上。</p></div>');
    return h.join('');
  }

  function helpHTML() {
    return '<div class="sect help"><h4>快捷键</h4>' +
      '<p class="hint">双击文字 编辑 · 拖动 移动 · 八个手柄 缩放</p>' +
      '<p class="hint">方向键 0.02in · Shift+方向键 0.1in</p>' +
      '<p class="hint">Delete 删除选中 · Esc 取消选中</p>' +
      '<p class="hint">Ctrl/Cmd+S 保存 · Ctrl/Cmd+Z 撤销未保存的修改</p></div>';
  }

  function renderPanel() {
    if (!dom.props) return;
    var el = sel ? elBy(sel) : null;
    if (sel && !el) sel = null;
    dom.props.innerHTML = (el ? elementSectionHTML(el) : slideSectionHTML()) + helpHTML();
    mark();
  }

  function onPanelClick(e) {
    var t = e.target;
    if (!t || t.nodeType !== 1 || !t.id) return;
    switch (t.id) {
      case 'ed-dup': duplicateSel(); break;
      case 'ed-del': deleteSel(); break;
      case 'ed-imgpick': {
        var fi = byId('ed-imgfile');
        if (fi) fi.click();
        break;
      }
      case 'ed-chartedit': {
        var el = sel ? elBy(sel) : null;
        if (el) openChartDialog(el);
        break;
      }
      case 'ed-addslide': addSlide(); break;
      case 'ed-dupslide': duplicateSlide(); break;
      case 'ed-delslide': deleteSlide(); break;
    }
  }

  function onPanelChange(e) {
    var t = e.target;
    if (!t || t.nodeType !== 1) return;
    var id = t.id;
    if (id === 'ed-imgfile') { uploadImage(t); return; }
    var el = sel ? elBy(sel) : null;
    if (!el) return;

    if (id === 'ed-x' || id === 'ed-y' || id === 'ed-w' || id === 'ed-h') {
      var g = geomOf(el);
      ['x', 'y', 'w', 'h'].forEach(function (k) {
        var inp = byId('ed-' + k);
        var v = inp ? parseFloat(inp.value) : NaN;
        if (isFinite(v)) g[k] = v;
      });
      g = clampGeom(g);
      commitGeom(el, g);
      syncGeomInputs(geomOf(el));
      return;
    }
    if (id === 'ed-size') {
      var size = parseFloat(t.value);
      if (!isFinite(size) || size <= 0) return;
      pushStyle(el, 'size', { size: size });
      setTextProp(el, 'size', size);
      return;
    }
    if (id === 'ed-align') {
      pushStyle(el, 'align', { align: t.value });
      setTextProp(el, 'align', t.value);
      return;
    }
    if (id === 'ed-lh') {
      var lh = parseFloat(t.value);
      if (!isFinite(lh) || lh <= 0) return;
      pushStyle(el, 'lineSpacing', { lineSpacing: lh });
      setTextProp(el, 'lineSpacing', lh);
      return;
    }
    if (id === 'ed-bold') {
      var b = !!t.checked;
      pushStyle(el, 'bold', { bold: b });
      setTextProp(el, 'bold', b);
    }
  }

  function onPanelInput(e) {
    var t = e.target;
    if (!t || t.nodeType !== 1) return;
    if (t.id === 'ed-color') {
      var el = sel ? elBy(sel) : null;
      if (!el) return;
      pushStyle(el, 'color', { color: t.value });
      setTextProp(el, 'color', t.value);
      return;
    }
    if (t.id === 'ed-notes') setNotes(t.value);
  }

  /* ---------------------------------------------------------------- 溢出检查 */

  function clearOvf() {
    qsa('.ovf').forEach(function (n) { n.classList.remove('ovf'); });
  }

  function scanOverflow(silent) {
    clearOvf();
    var hits = [];
    qsa('.slide .tx').forEach(function (tx) {
      if (tx.scrollHeight > tx.clientHeight + 2 || tx.scrollWidth > tx.clientWidth + 2) hits.push(tx);
    });
    hits.forEach(function (tx) {
      tx.classList.add('ovf');
      var host = closestOf(tx, '.el');
      if (host) host.classList.add('ovf');
    });
    if (!hits.length) {
      console.info('[editor] 溢出检查：未发现文字溢出');
      if (!silent) toast('未发现溢出');
      return 0;
    }
    var where = hits.map(function (tx) {
      var host = closestOf(tx, '.el');
      return host ? (slideId(host) + '/' + host.dataset.elid) : '?';
    });
    console.info('[editor] 溢出检查：' + hits.length + ' 处', where);
    if (!silent) {
      toast(hits.length + ' 处文字可能溢出', 4000);
      try { hits[0].scrollIntoView({ block: 'center', behavior: 'smooth' }); }
      catch (e) { hits[0].scrollIntoView(); }
    }
    return hits.length;
  }

  /* ---------------------------------------------------------------- 保存 / 导出 */

  /* 保存协议：把队列里的 op 一次性 POST /api/patch。
     - ok:false -> 保留队列（绝不假装成功），把服务端错误 toast 出来；
     - ok:true + reload:true -> 服务端已经改写 deck，页面上的 DOM 过期了，
       等提示条露个脸再 location.reload()；
     - then(done) 给导出用：done() 之后才允许刷新。 */
  function save(then) {
    if (saving) { toast('正在保存，请稍候'); return; }
    if (!queue.length) {
      if (then) then(function () { });        // 导出：没有待保存内容就直接继续
      else toast('没有需要保存的修改');
      return;
    }
    saving = true;
    if (dom.save) { dom.save.disabled = true; dom.save.textContent = '保存中…'; }
    var sent = queue.slice();                 // 请求期间新入队的操作要保留
    var payload = sent.map(function (r) { return r.op; });
    api('/api/patch', { ops: payload }).then(function (j) {
      saving = false;
      if (dom.save) dom.save.disabled = false;
      if (!j || !j.ok) {
        mark();
        console.error('[editor] 保存失败', j, payload);
        toast('保存失败：' + ((j && j.error) || '未知错误') + '（修改仍在队列里，可用 Ctrl+Z 撤销）', 6000);
        return;
      }
      queue = queue.filter(function (r) { return sent.indexOf(r) < 0; });
      mark();
      /* 服务端用 warnings 逐条报告 op 的问题（找不到元素/页面…）。ok:true 只说明
         请求成功，不代表每条 op 都落地了，所以警告也必须让用户看见。 */
      var warns = j.warnings || [];
      if (warns.length) {
        console.warn('[editor] 服务端警告', warns);
        toast('已保存，但服务端有 ' + warns.length + ' 条警告：' + warns.slice(0, 3).join('；'), 6000);
      } else {
        toast('已保存');
      }
      var done = function () {
        // 结构性改动（增删页/元素、换版式/主题）会让页面上的 DOM 过期，必须重载；
        // 有警告时多等一会儿，别把提示吃掉
        if (j.reload) setTimeout(function () { location.reload(); }, warns.length ? 2600 : 650);
      };
      if (then) then(done); else done();
    });
  }

  function doExport(btn, url, label, busyText) {
    if (!btn || btn.disabled) return;
    save(function (done) {
      var old = btn.textContent;
      btn.disabled = true;
      btn.textContent = busyText;
      api(url, {}).then(function (j) {
        btn.disabled = false;
        btn.textContent = old || label;
        if (j && j.ok) toastPath('已导出：', j.path || '');
        else toast('导出失败：' + ((j && j.error) || '未知错误'), 6000);
        done();
      });
    });
  }

  /* ---------------------------------------------------------------- 布局 */

  function fit() {
    if (!dom.stage) return;
    var S = stageSize();
    var avail = Math.max(160, dom.stage.clientWidth - 32);
    var s = Math.min(1, avail / S.w);
    slides().forEach(function (sl) {
      sl.style.transform = 'scale(' + s + ')';
      // transform 不改变盒子高度，用负外边距吃掉多出来的空白，页与页之间不留大片空隙
      sl.style.marginBottom = Math.round(26 - (1 - s) * S.h) + 'px';
    });
    if (sel) {
      var el = elBy(sel);
      if (el) positionHandles(el);
    }
  }

  function onResize() { fit(); }

  /* ---------------------------------------------------------------- 键盘 */

  function onKeyDown(e) {
    if (e.key === 'Escape') {
      if (dlg && dlg.classList.contains('on')) { closeDialog(); e.preventDefault(); return; }
      if (sel) { select(null); e.preventDefault(); }
      return;
    }
    var mod = e.ctrlKey || e.metaKey;
    if (mod && (e.key === 's' || e.key === 'S')) { e.preventDefault(); save(); return; }
    if (isTyping(e.target)) return;             // 正在输入时绝不抢键
    if (mod && (e.key === 'z' || e.key === 'Z') && !e.shiftKey) { e.preventDefault(); undoLast(); return; }
    if (!sel) return;
    if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); deleteSel(); return; }
    var step = e.shiftKey ? 0.1 : 0.02;
    if (e.key === 'ArrowLeft') { e.preventDefault(); nudge(-step, 0); }
    else if (e.key === 'ArrowRight') { e.preventDefault(); nudge(step, 0); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); nudge(0, -step); }
    else if (e.key === 'ArrowDown') { e.preventDefault(); nudge(0, step); }
  }

  function onBeforeUnload(e) {
    if (!queue.length) return;
    e.preventDefault();
    e.returnValue = '';
    return '';
  }

  /* ---------------------------------------------------------------- 启动 */

  function boot() {
    dom.topbar = byId('topbar');
    dom.stage = byId('stage');
    dom.props = byId('props');
    dom.toast = byId('toast');
    dom.hud = byId('hud');
    dom.pageinfo = byId('pageinfo');
    dom.save = byId('savebtn');
    dom.theme = byId('themebtn');
    dom.layout = byId('layoutsel');

    if (dom.save) dom.save.addEventListener('click', function () { save(); });
    if (dom.theme) { themeWas = dom.theme.value; dom.theme.addEventListener('change', onThemeChange); }
    if (dom.layout) dom.layout.addEventListener('change', onLayoutChange);

    var add = byId('addelem');
    if (add) add.addEventListener('click', addTextBox);
    var ovf = byId('ovfbtn');
    if (ovf) ovf.addEventListener('click', function () { scanOverflow(false); });
    var pres = byId('presentbtn');
    if (pres) pres.addEventListener('click', function () { window.open('/preview'); });
    var exp = byId('exportbtn');
    if (exp) exp.addEventListener('click', function () { doExport(exp, '/api/export/pptx', '导出 PPTX', '导出中…'); });
    var pdf = byId('pdfbtn');
    if (pdf) pdf.addEventListener('click', function () { doExport(pdf, '/api/export/pdf', '导出 PDF', '生成中…'); });

    if (dom.props) {
      dom.props.addEventListener('click', onPanelClick);
      dom.props.addEventListener('change', onPanelChange);
      dom.props.addEventListener('input', onPanelInput);
    }

    document.addEventListener('pointerdown', onPointerDown, true);
    document.addEventListener('pointermove', onPointerMove);
    document.addEventListener('pointerup', onPointerUp);
    document.addEventListener('pointercancel', onPointerUp);
    document.addEventListener('dblclick', function (e) {
      var tx = closestOf(e.target, '.tx');
      if (tx) startTextEdit(tx);
    });
    document.addEventListener('focusout', function (e) {
      var tx = e.target;
      if (tx && tx.classList && tx.classList.contains('tx')) endTextEdit(tx);
    });
    document.addEventListener('keydown', onKeyDown, true);
    window.addEventListener('resize', onResize);
    window.addEventListener('scroll', onScroll, { passive: true });
    window.addEventListener('beforeunload', onBeforeUnload);

    fit();
    setCurrent(nearestIndex(), false);
    renderPanel();
    scanOverflow(true);          // 载入时静默检查一次，只在控制台报告

    api('/api/meta', undefined).then(function (j) {
      if (!j || !j.ok) { toast('版式列表加载失败：' + ((j && j.error) || '未知错误'), 5000); return; }
      meta = j;
      populateThemes();
      populateLayouts();
    });
    api('/api/deck', undefined).then(function (j) {
      if (j && j.ok && j.deck) rawDeck = j.deck;
    });

    window.__editor = {
      save: save,
      undo: undoLast,
      select: select,
      scanOverflow: scanOverflow,
      ops: function () { return queue.map(function (r) { return r.op; }); },
      isDirty: function () { return queue.length > 0; },
      version: '1.0'
    };
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
