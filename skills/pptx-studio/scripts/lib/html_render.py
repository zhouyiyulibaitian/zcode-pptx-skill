"""Deck -> standalone HTML preview (and the DOM contract the editor edits).

One HTML file, no external assets, no CDN.  Geometry is emitted as absolute
pixels at exactly 96 dpi against the deck's own stage size, so what the browser
shows is the same geometry the PPTX exporter writes -- the preview is not an
approximation produced by measuring the DOM.
"""

from __future__ import annotations

import html
import json
import math
import os

from . import model as M

STAGE_DPR = 96.0


def _px(inches: float) -> float:
    return round(float(inches) * STAGE_DPR, 2)


def _pt2px(pt: float) -> float:
    return round(float(pt) * M.PT_TO_PX, 2)


def _rgba(color: str, alpha: float | None) -> str:
    if color is None:
        return "transparent"
    if not color.startswith("#"):
        return color
    r, g, b = M.hex_to_rgb(color)
    if alpha is None or float(alpha) >= 1:
        return color
    return f"rgba({r},{g},{b},{round(float(alpha), 3)})"


def _esc(s) -> str:
    return html.escape(str(s if s is not None else ""), quote=True)


# --------------------------------------------------------------------------- #
# element renderers
# --------------------------------------------------------------------------- #

def _text_css(el, theme) -> str:
    pad = M._padding(el)
    valign = {"top": "flex-start", "middle": "center", "bottom": "flex-end"}.get(
        el.get("valign", "top"), "flex-start")
    bits = [f"padding:{_px(pad[1])}px {_px(pad[2])}px {_px(pad[3])}px {_px(pad[0])}px",
            f"display:flex", f"flex-direction:column", f"justify-content:{valign}"]
    if el.get("fill"):
        f = el["fill"]
        if isinstance(f, dict) and f.get("type", "solid") == "solid":
            bits.append(f"background:{_rgba(f.get('color', '#fff'), f.get('alpha'))}")
    if el.get("line"):
        ln = el["line"]
        if isinstance(ln, dict) and ln.get("color"):
            bits.append(f"border:{_pt2px(ln.get('width', 1))}px solid {ln['color']}")
    if el.get("radius") and _is_round_safe(el):
        bits.append(f"border-radius:{_px(el['radius'])}px")
    return ";".join(bits)


def _is_round_safe(el) -> bool:
    return True


def _run_html(run, para, theme) -> str:
    style = []
    size = run.get("size") or para.get("size")
    if size:
        style.append(f"font-size:{_pt2px(size)}px")
    color = run.get("color") or para.get("color")
    if color:
        style.append(f"color:{color}")
    bold = run.get("bold", para.get("bold"))
    if bold:
        style.append("font-weight:700")
    italic = run.get("italic", para.get("italic"))
    if italic:
        style.append("font-style:italic")
    if run.get("underline"):
        style.append("text-decoration:underline")
    if run.get("spacing"):
        style.append(f"letter-spacing:{_pt2px(run['spacing'])}px")
    font = run.get("font") or para.get("font")
    if font:
        style.append(f"font-family:{_font_stack(font, theme)}")
    return f'<span style="{";".join(style)}">{_esc(run.get("text", ""))}</span>'


def _font_stack(font, theme) -> str:
    if isinstance(font, dict):
        f = font
    else:
        f = {"latin": font, "ea": font}
    fams = []
    for key in ("latin", "ea"):
        v = f.get(key)
        if v and v not in fams:
            fams.append(v)
    # keep the theme's own stack as the tail so a missing font degrades gracefully
    tail = ((theme.get("fonts") or {}).get("body") or {}).get("html")
    for extra in ("Microsoft YaHei", "微软雅黑", "sans-serif"):
        fams.append(extra)
    css = ",".join(f"'{x}'" if " " in x or not x.isascii() else x for x in fams)
    if tail:
        css += "," + tail
    return css


def _paragraph_html(p, theme, first=False) -> str:
    style = []
    align = p.get("align")
    if align:
        style.append(f"text-align:{align}")
    size = p.get("size")
    if p.get("lineSpacing"):
        style.append(f"line-height:{p['lineSpacing']}")
    elif size:
        style.append("line-height:1.35")
    color = p.get("color")
    if color:
        style.append(f"color:{color}")
    if size:
        style.append(f"font-size:{_pt2px(size)}px")
    if p.get("spaceBefore") and not first:
        style.append(f"margin-top:{_pt2px(p['spaceBefore'])}px")
    if p.get("spaceAfter"):
        style.append(f"margin-bottom:{_pt2px(p['spaceAfter'])}px")
    if p.get("bullet"):
        style.append("padding-left:0.95em;text-indent:-0.95em")
    if p.get("indent"):
        style.append(f"padding-left:{_px(p['indent'])}px")
    if p.get("font") and not any(r.get("font") for r in p.get("runs", [])):
        style.append(f"font-family:{_font_stack(p['font'], theme)}")
    runs = p.get("runs") or [{"text": ""}]
    inner = "".join(_run_html(r, p, theme) for r in runs)
    if p.get("bullet"):
        inner = '<span style="display:inline-block;width:0.95em">•</span>' + inner
    return f'<p style="{";".join(style)}">{inner}</p>'


def _shape_svg(el, theme) -> str:
    """CSS/SVG rendering for the shape vocabulary shared with the PPTX exporter."""
    shape = el.get("shape", "rect")
    w, h = _px(el.get("w", 1)), _px(el.get("h", 1))
    fill = el.get("fill") or {}
    line = el.get("line") or {}
    fc = fill.get("color") if isinstance(fill, dict) else None
    fa = fill.get("alpha") if isinstance(fill, dict) else None
    lc = line.get("color") if isinstance(line, dict) else None
    lw = float(line.get("width", 1)) if isinstance(line, dict) else 1
    dash = line.get("dash") if isinstance(line, dict) else None
    dash_attr = f' stroke-dasharray="{_pt2px(lw) * 3:.1f},{_pt2px(lw) * 2.5:.1f}"' if dash in ("dash", "dot") else ""
    if shape in ("line", "arrow"):
        y = h / 2
        marker = ""
        if shape == "arrow":
            marker = (" marker-end=\"url(#arw)\"" if lc else "")
        return (f'<svg class="shp" width="{w}" height="{h}" viewBox="0 0 {w} {h}" '
                f'preserveAspectRatio="none"><defs><marker id="arw" markerWidth="6" markerHeight="6" '
                f'refX="5" refY="3" orient="auto"><path d="M0,0 L6,3 L0,6 z" fill="{lc or "#000"}"/></marker>'
                f'</defs><line x1="0" y1="{y}" x2="{w}" y2="{y}" stroke="{lc or "#000"}" '
                f'stroke-width="{max(_pt2px(lw), 1)}"{dash_attr}{marker}/></svg>')
    clip = {
        "triangle": "polygon(50% 0,100% 100%,0 100%)",
        "diamond": "polygon(50% 0,100% 50%,50% 100%,0 50%)",
        "chevron": "polygon(0 0,75% 0,100% 50%,75% 100%,0 100%,25% 50%)",
        "pentagon": "polygon(50% 0,100% 38%,82% 100%,18% 100%,0 38%)",
        "parallelogram": "polygon(18% 0,100% 0,82% 100%,0 100%)",
        "star5": "polygon(50% 0,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%)",
        "plus": "polygon(35% 0,65% 0,65% 35%,100% 35%,100% 65%,65% 65%,65% 100%,35% 100%,35% 65%,0 65%,0 35%,35% 35%)",
        "arc": "polygon(0 100%,100% 100%,100% 0)",
        "ring": None,
    }.get(shape)
    radius = None
    if shape == "roundRect":
        radius = _px(el.get("radius", 0.1))
    elif shape == "ellipse":
        radius = "50%"
    elif shape == "ring":
        radius = "50%"
    style = [f"width:{w}px", f"height:{h}px",
             f"background:{_rgba(fc, fa)}" if fc else "background:transparent"]
    if clip:
        style.append(f"clip-path:{clip}")
    if radius:
        style.append(f"border-radius:{radius}")
    if lc:
        style.append(f"border:{max(_pt2px(lw), 1)}px {'dashed' if dash else 'solid'} {lc}")
    if shape == "ring" and fc:
        style.append(f"box-shadow:inset 0 0 0 {max(_px(max(el.get('w', 1), 1) * 0.18), 4)}px {_rgba(fc, fa)}")
    if shape == "arc":
        style = [s for s in style if not s.startswith("border")]
    style.append("box-sizing:border-box")
    return f'<div class="shp" style="{";".join(style)}"></div>'


def _image_html(el, theme, href_prefix: str) -> str:
    src = el.get("src") or ""
    w, h = _px(el.get("w", 1)), _px(el.get("h", 1))
    radius = _px(el.get("radius", 0)) if el.get("radius") else 0
    if not src:
        return (f'<div class="ph" style="width:{w}px;height:{h}px;border-radius:{radius}px">'
                f'<span>{_esc(el.get("placeholder") or "图片占位")}</span></div>')
    href = src if os.path.isabs(src) or src.startswith("http") else href_prefix + src
    fit = "cover" if el.get("fit", "cover") == "cover" else "contain"
    op = el.get("opacity")
    op_css = f";opacity:{op}" if op is not None and float(op) < 1 else ""
    return (f'<div class="imw" style="width:{w}px;height:{h}px;border-radius:{radius}px{op_css}">'
            f'<img src="{_esc(href)}" alt="{_esc(el.get("alt", ""))}" '
            f'style="object-fit:{fit};width:100%;height:100%;display:block" loading="lazy"></div>')


def _table_html(el, theme) -> str:
    cols = el.get("cols") or []
    rows = el.get("rows") or []
    header = bool(el.get("header", True))
    ncol = len(cols) or (len(rows[0]) if rows else 1)
    total_w = _px(el.get("w", 6))
    if cols:
        colgroup = "".join(f'<col style="width:{_px(c) / total_w * 100:.4f}%">' for c in cols)
    else:
        colgroup = "".join(f'<col style="width:{100 / max(ncol, 1):.4f}%">' for _ in range(ncol))
    row_h = _px(el.get("rowHeight", 0.45))
    body = []
    for ri, row in enumerate(rows):
        is_head = header and ri == 0
        cells = []
        for cell in row:
            c = cell if isinstance(cell, dict) else {"text": str(cell)}
            style = [f"height:{row_h}px"]
            if c.get("fill"):
                style.append(f"background:{c['fill']}")
            elif is_head:
                style.append(f"background:{_rgba(theme['colors'].get('primary'), 1)}")
            # Every cell needs an explicit colour: the editor chrome sets a light
            # text colour on <body>, and without this it would be inherited onto
            # white cells and the table would be unreadable.
            if c.get("color"):
                style.append(f"color:{c['color']}")
            elif is_head:
                style.append(f"color:{theme['colors'].get('onPrimary', '#fff')}")
            else:
                style.append(f"color:{theme['colors'].get('text', '#111111')}")
            size = c.get("size") or (theme["geometry"].get("bodySize", 14) * 0.92)
            style.append(f"font-size:{_pt2px(size)}px")
            if c.get("bold") or is_head:
                style.append("font-weight:700")
            style.append(f"text-align:{c.get('align', 'center' if is_head else 'left')}")
            style.append("padding:0.18em 0.5em")
            if c.get("align") is None:
                pass
            cells.append(f'<td style="{";".join(style)}">{_esc(c.get("text", ""))}</td>')
        body.append("<tr>" + "".join(cells) + "</tr>")
    return (f'<table class="tbl" style="width:{total_w}px;height:{_px(el.get("h", 1))}px">'
            f"<colgroup>{colgroup}</colgroup><tbody>{''.join(body)}</tbody></table>")


def _chart_html(el, theme) -> str:
    """Charts are rendered as inline SVG by svgchart so the HTML stays offline-safe.

    A missing or failing chart module degrades to a visible placeholder: a broken
    chart must never take down the whole preview.
    """
    try:
        from . import svgchart
    except ImportError:
        return '<div class="ph">图表模块 svgchart.py 缺失</div>'
    try:
        return svgchart.chart_svg(dict(el), theme, _px(el.get("w", 6)), _px(el.get("h", 3.5)))
    except Exception as exc:
        return f'<div class="ph">图表渲染失败: {_esc(exc)}</div>'


def render_element(el: dict, theme: dict, href_prefix: str = "") -> str:
    x, y = _px(el.get("x", 0)), _px(el.get("y", 0))
    w, h = _px(el.get("w", 1)), _px(el.get("h", 1))
    rot = float(el.get("rotation") or 0)
    style = [f"left:{x}px", f"top:{y}px", f"width:{w}px", f"height:{h}px",
             f"z-index:{int(el.get('z', 0)) + 50}"]
    if rot:
        style.append(f"transform:rotate({rot}deg)")
    if el.get("opacity") is not None and float(el["opacity"]) < 1:
        style.append(f"opacity:{el['opacity']}")
    cls = f"el el-{el.get('type', 'text')}"
    if el.get("locked"):
        cls += " locked"
    if el.get("role"):
        cls += f" role-{el['role']}"
    etype = el.get("type")
    if etype == "text":
        inner = "".join(_paragraph_html(p, theme, first=(i == 0))
                        for i, p in enumerate(el.get("paragraphs") or [{"runs": [{"text": ""}]}]))
        inner = f'<div class="tx" style="{_text_css(el, theme)}">{inner}</div>'
    elif etype == "shape":
        inner = _shape_svg(el, theme)
        if el.get("paragraphs"):
            txt = "".join(_paragraph_html(p, theme, first=(i == 0))
                          for i, p in enumerate(el["paragraphs"]))
            v = {"top": "flex-start", "middle": "center", "bottom": "flex-end"}.get(el.get("valign", "middle"), "center")
            inner += (f'<div class="tx shp-tx" style="position:absolute;inset:0;display:flex;'
                      f'flex-direction:column;justify-content:{v};padding:0.2em 0.35em">{txt}</div>')
    elif etype == "image":
        inner = _image_html(el, theme, href_prefix)
    elif etype == "table":
        inner = _table_html(el, theme)
    elif etype == "chart":
        inner = _chart_html(el, theme)
    else:
        inner = f'<div class="unknown">{_esc(etype)}</div>'
    return (f'<div class="{cls}" data-elid="{_esc(el.get("id"))}" data-type="{_esc(etype)}" '
            f'style="{";".join(style)}">{inner}</div>')


def _background_css(bg: dict | None, theme: dict) -> str:
    if not bg:
        return f"background:{theme['colors'].get('bg', '#fff')}"
    t = bg.get("type", "solid")
    if t == "gradient":
        ang = float(bg.get("angle", 90)) + 90.0
        return (f"background:linear-gradient({ang:.0f}deg,{bg.get('from', '#fff')},"
                f"{bg.get('to', '#000')})")
    if t == "image" and bg.get("src"):
        return (f"background-image:url('{_esc(bg['src'])}');background-size:cover;"
                f"background-position:center")
    return f"background:{_esc(bg.get('color', theme['colors'].get('bg', '#fff')))}"


# --------------------------------------------------------------------------- #
# document
# --------------------------------------------------------------------------- #

VIEWER_JS = r"""
(function(){
  var slides=[].slice.call(document.querySelectorAll('.slide'));
  var idx=0, stageW=window.__STAGE_W, label=document.getElementById('pageinfo');
  function fit(){
    if(document.body.classList.contains('present')){
      var s=Math.min(window.innerWidth/stageW,(window.innerHeight-0)/window.__STAGE_H);
      slides.forEach(function(sl){sl.style.transform='scale('+s+')';});
    }else{
      var avail=Math.min(window.innerWidth-32, stageW);
      var s=avail/stageW;
      slides.forEach(function(sl){sl.style.transform='scale('+s+')';});
    }
  }
  function show(i,scroll){
    idx=Math.max(0,Math.min(slides.length-1,i));
    if(label) label.textContent=(idx+1)+' / '+slides.length;
    if(document.body.classList.contains('present')){
      slides.forEach(function(sl,k){sl.style.visibility=(k===idx?'visible':'hidden');});
    }else if(scroll!==false){ slides[idx].scrollIntoView({behavior:'smooth',block:'center'}); }
  }
  function togglePresent(on){
    document.body.classList.toggle('present',on===undefined?!document.body.classList.contains('present'):on);
    fit();show(idx,false);
    if(document.body.classList.contains('present')&&document.documentElement.requestFullscreen){document.documentElement.requestFullscreen().catch(function(){});}
  }
  window.addEventListener('resize',fit);
  document.addEventListener('keydown',function(e){
    if(e.target&&e.target.isContentEditable) return;
    if(e.key==='ArrowRight'||e.key==='ArrowDown'||e.key===' '||e.key==='PageDown'){e.preventDefault();show(idx+1);}
    else if(e.key==='ArrowLeft'||e.key==='ArrowUp'||e.key==='PageUp'){e.preventDefault();show(idx-1);}
    else if(e.key==='Home'){show(0);} else if(e.key==='End'){show(slides.length-1);}
    else if(e.key==='p'||e.key==='P'){window.print();}
    else if(e.key==='f'||e.key==='F'){togglePresent();}
    else if(e.key==='Escape'&&document.body.classList.contains('present')){togglePresent(false);}
  });
  document.body.addEventListener('click',function(e){
    if(e.target.closest('.el'))return;
    if(document.body.classList.contains('present')){ show(idx+1); }
    else togglePresent(true);
  });
  var io=new IntersectionObserver(function(es){es.forEach(function(en){if(en.isIntersecting)show(+en.target.dataset.index,false);});},{threshold:0.55});
  slides.forEach(function(s){io.observe(s);});
  fit();show(0,false);
  window.__go=show;
})();
"""

EDIT_JS = r"""
/* Edit mode: mutate the server-rendered DOM in place, then send patches.
   The visual truth stays the Python renderer -- this never re-implements it. */
(function(){
  var PPI=96, ops=[], sel=null, drag=null, dirty=false;
  var deckPath=window.__DECK_PATH, stageW=window.__STAGE_W, stageH=window.__STAGE_H;
  var panel=document.getElementById('props'), toastEl=document.getElementById('toast');
  function toast(t,ms){toastEl.textContent=t;toastEl.classList.add('on');clearTimeout(toastEl._t);toastEl._t=setTimeout(function(){toastEl.classList.remove('on')},ms||1800);}
  function px(v){return parseFloat(v)||0;}
  function inch(v){return Math.round(v/PPI*1000)/1000;}
  function cs(el,prop){return getComputedStyle(el).getPropertyValue(prop);}
  function push(op){ops.push(op);dirty=true;mark();}
  function mark(){var b=document.getElementById('savebtn');if(b)b.textContent=dirty?'保存 *':'保存';}
  function elBy(id){return document.querySelector('[data-elid="'+id+'"]');}
  function slideOf(node){var s=node.closest('.slide');return s?s.dataset.id:null;}

  /* ---- selection ---- */
  function select(id){
    if(sel){var p=elBy(sel);if(p)p.classList.remove('sel');}
    sel=id;
    var el=elBy(id);
    if(el){el.classList.add('sel');}
    buildPanel(el);
  }
  document.addEventListener('click',function(e){
    var t=e.target.closest('.el');
    if(!t){select(null);return;}
    if(e.target.isContentEditable)return;
    e.stopPropagation();select(t.dataset.elid);
  },true);

  /* ---- inline text editing ---- */
  document.addEventListener('dblclick',function(e){
    var tx=e.target.closest('.tx');
    if(!tx)return;
    tx.contentEditable='true';tx.focus();
    document.execCommand&&document.execCommand('selectAll',false,null);
    tx._orig=tx.innerHTML;
  });
  document.addEventListener('blur',function(e){
    var tx=e.target;
    if(!tx||!tx.classList||!tx.classList.contains('tx')||!tx.isContentEditable)return;
    tx.contentEditable='false';
    if(tx._orig===tx.innerHTML)return;
    var id=tx.closest('.el').dataset.elid, slide=slideOf(tx);
    var paras=[].slice.call(tx.querySelectorAll('p')).map(function(p){
      return {runs:[{text:p.innerText.replace(/\u00a0/g,' ')}]};
    });
    push({op:'text',slide:slide,el:id,paragraphs:paras});
    toast('文字已改，记得保存');
  },true);

  /* ---- drag + resize ---- */
  function screenScale(){ var s=document.querySelector('.slide'); if(!s)return 1; return s.getBoundingClientRect().width/(stageW); }
  document.addEventListener('pointerdown',function(e){
    var el=e.target.closest('.el'); if(!el)return;
    if(e.target.isContentEditable)return;
    if(el.classList.contains('locked'))return;
    var handle=e.target.classList.contains('h')?e.target.dataset.h:null;
    var box=el.getBoundingClientRect(), sc=screenScale();
    drag={el:el,handle:handle,sx:e.clientX,sy:e.clientY,
          x0:px(el.style.left),y0:px(el.style.top),w0:px(el.style.width),h0:px(el.style.height),sc:sc};
    el.setPointerCapture&&el.setPointerCapture(e.pointerId);
    e.preventDefault();
  });
  document.addEventListener('pointermove',function(e){
    if(!drag)return;
    var dx=(e.clientX-drag.sx)/drag.sc, dy=(e.clientY-drag.sy)/drag.sc, el=drag.el, h=drag.handle;
    if(!h||h==='move'){
      el.style.left=(drag.x0+dx)+'px'; el.style.top=(drag.y0+dy)+'px';
    }else{
      if(h.indexOf('e')>=0) el.style.width=Math.max(12,drag.w0+dx)+'px';
      if(h.indexOf('s')>=0) el.style.height=Math.max(12,drag.h0+dy)+'px';
      if(h.indexOf('w')>=0){el.style.width=Math.max(12,drag.w0-dx)+'px'; el.style.left=(drag.x0+dx)+'px';}
      if(h.indexOf('n')>=0){el.style.height=Math.max(12,drag.h0-dy)+'px'; el.style.top=(drag.y0+dy)+'px';}
      var inner=el.querySelector('.tx'); if(inner){inner.style.height='100%';}
      var svg=el.querySelector('svg'); if(svg){svg.setAttribute('width',px(el.style.width));svg.setAttribute('height',px(el.style.height));svg.setAttribute('viewBox','0 0 '+px(el.style.width)+' '+px(el.style.height));}
    }
  });
  document.addEventListener('pointerup',function(){
    if(!drag)return;
    var el=drag.el,id=el.dataset.elid;
    push({op:'geometry',slide:slideOf(el),el:id,
          x:inch(px(el.style.left)),y:inch(px(el.style.top)),
          w:inch(px(el.style.width)),h:inch(px(el.style.height))});
    drag=null;mark();
  });

  /* ---- property panel ---- */
  function field(label,html){return '<label class="f"><span>'+label+'</span>'+html+'</label>';}
  function buildPanel(el){
    if(!el){panel.innerHTML='<div class="ph">选中一个元素后在这里改属性。</div>'+
      '<div class="sect"><h4>页面</h4><div id="slideprops"></div></div>';buildSlidePanel();return;}
    var id=el.dataset.elid, type=el.dataset.type, slide=slideOf(el);
    var tx=el.querySelector('.tx'), p=(tx&&tx.querySelector('p'))||null;
    var h=[];
    h.push('<div class="hd">'+type+' · '+id+'</div>');
    h.push('<div class="sect"><h4>位置尺寸 (英寸)</h4><div class="grid4">'+
      '<input id="gx" value="'+inch(px(el.style.left))+'" title="x"><input id="gy" value="'+inch(px(el.style.top))+'" title="y">'+
      '<input id="gw" value="'+inch(px(el.style.width))+'" title="w"><input id="gh" value="'+inch(px(el.style.height))+'" title="h"></div></div>');
    if(type==='text'||el.querySelector('.tx')){
      var size=p?parseFloat(cs(p,'font-size'))*0.75:14;
      var color=p?rgb2hex(cs(p,'color')):'#111111';
      h.push('<div class="sect"><h4>文字</h4>'+
        field('字号 pt','<input id="tsize" type="number" step="0.5" value="'+Math.round(size*10)/10+'">')+
        field('颜色','<input id="tcolor" type="color" value="'+color+'">')+
        field('对齐','<select id="talign"><option value="left">左</option><option value="center">中</option><option value="right">右</option></select>')+
        field('加粗','<input id="tbold" type="checkbox"'+(p&&cs(p,'font-weight')>=600?' checked':'')+'>')+
        field('行距','<input id="tlh" type="number" step="0.05" value="'+(p?cs(p,'line-height'):'1.4')+'">')+
        '</div>');
    }
    if(type==='chart'){
      h.push('<div class="sect"><h4>图表数据</h4><p class="hint">在下方表格中用逗号或换行分隔数值，改完点应用。</p><button id="chartedit" class="btn">编辑数据…</button></div>');
    }
    if(type==='image'){
      h.push('<div class="sect"><h4>图片</h4><button id="imgpick" class="btn">替换图片…</button><input id="imgfile" type="file" accept="image/*" style="display:none"></div>');
    }
    h.push('<div class="sect"><h4>操作</h4><button id="dup" class="btn">复制元素</button><button id="del" class="btn danger">删除元素</button></div>');
    panel.innerHTML=h.join('');
    var g=function(i){return document.getElementById(i)};
    ['gx','gy','gw','gh'].forEach(function(k){
      g(k).addEventListener('change',function(){
        push({op:'geometry',slide:slide,el:id,x:+g('gx').value,y:+g('gy').value,w:+g('gw').value,h:+g('gh').value});
        el.style.left=_px(g('gx').value);el.style.top=_px(g('gy').value);el.style.width=_px(g('gw').value);el.style.height=_px(g('gh').value);
      });
    });
    if(g('tsize'))g('tsize').addEventListener('change',function(){
      push({op:'style',slide:slide,el:id,set:{size:+this.value}});
      if(tx)[].slice.call(tx.querySelectorAll('p')).forEach(function(p){p.style.fontSize=_px2(+g('tsize').value)});
    });
    if(g('tcolor'))g('tcolor').addEventListener('input',function(){
      if(tx)[].slice.call(tx.querySelectorAll('p')).forEach(function(p){p.style.color=this.value}.bind(this));
      push({op:'style',slide:slide,el:id,set:{color:g('tcolor').value}});
    });
    if(g('talign'))g('talign').addEventListener('change',function(){
      if(tx)[].slice.call(tx.querySelectorAll('p')).forEach(function(p){p.style.textAlign=this.value}.bind(this));
      push({op:'style',slide:slide,el:id,set:{align:g('talign').value}});
    });
    if(g('tbold'))g('tbold').addEventListener('change',function(){
      var b=this.checked;
      if(tx)[].slice.call(tx.querySelectorAll('p')).forEach(function(p){p.style.fontWeight=b?'700':'400'});
      push({op:'style',slide:slide,el:id,set:{bold:b}});
    });
    if(g('tlh'))g('tlh').addEventListener('change',function(){
      if(tx)[].slice.call(tx.querySelectorAll('p')).forEach(function(p){p.style.lineHeight=this.value}.bind(this));
      push({op:'style',slide:slide,el:id,set:{lineSpacing:+g('tlh').value}});
    });
    if(g('del'))g('del').addEventListener('click',function(){
      push({op:'element',action:'delete',slide:slide,el:id});el.remove();select(null);toast('已删除，保存后生效');
    });
    if(g('dup'))g('dup').addEventListener('click',function(){
      push({op:'element',action:'duplicate',slide:slide,el:id});toast('已复制，保存后生效');
    });
    if(g('imgpick'))g('imgpick').addEventListener('click',function(){g('imgfile').click()});
    if(g('imgfile'))g('imgfile').addEventListener('change',function(){
      var f=this.files[0]; if(!f)return;
      var fd=new FormData(); fd.append('file',f);
      fetch('/api/upload',{method:'POST',body:fd}).then(function(r){return r.json()}).then(function(j){
        if(!j.ok){toast('上传失败');return;}
        push({op:'image',slide:slide,el:id,src:j.path});
        var img=el.querySelector('img'); if(!img){el.querySelector('.ph').outerHTML='<img style="width:100%;height:100%;object-fit:cover">';img=el.querySelector('img');}
        img.src=j.url||j.path;
        toast('图片已替换');
      });
    });
  }
  function _px(v){return (v*96)+'px'}
  function _px2(v){return (v*96/72)+'px'}
  function rgb2hex(c){var m=c.match(/\d+/g);if(!m)return '#111111';return '#'+m.slice(0,3).map(function(x){return ('0'+parseInt(x).toString(16)).slice(-2)}).join('');}

  /* ---- slide operations ---- */
  function buildSlidePanel(){
    var sp=document.getElementById('slideprops'); if(!sp)return;
    var cur=null; var slideEl=null;
    var sl=document.querySelector('.slide:hover')||document.querySelector('.slide');
    var slides=[].slice.call(document.querySelectorAll('.slide'));
    var i=slides.findIndex(function(s){return s.dataset.id===(sel?slideOf(elBy(sel)):null)});
    h=[];
    h.push('<div class="row"><button id="addslide" class="btn">在本页后新增页</button></div>');
    if(i>=0){
      h.push('<div class="row"><button id="delslide" class="btn danger">删除本页</button><button id="dupslide" class="btn">复制本页</button></div>');
    }
    h.push('</div>');
    sp.innerHTML=h.join('');
    sp.querySelector('#addslide').onclick=function(){push({op:'slide',action:'add',index:(i<0?slides.length-1:i)});toast('已添加，保存后生效');};
    if(sp.querySelector('#delslide'))sp.querySelector('#delslide').onclick=function(){push({op:'slide',action:'delete',index:i});toast('已删除，保存后生效');};
    if(sp.querySelector('#dupslide'))sp.querySelector('#dupslide').onclick=function(){push({op:'slide',action:'duplicate',index:i});toast('已复制，保存后生效');};
  }

  /* ---- notes ---- */
  document.addEventListener('dblclick',function(){});
  var notesBox=document.getElementById('notes');
  if(notesBox){
    notesBox.addEventListener('input',function(){
      var s=document.querySelector('.sel')||document.querySelector('.slide');
      push({op:'notes',slide:(document.querySelector('.sel')?slideOf(document.querySelector('.sel')):notesBox.dataset.slide),text:notesBox.value});
    });
  }
  window.__pickNotes=function(slideId,text){notesBox.dataset.slide=slideId;notesBox.value=text||'';};

  /* ---- save / export ---- */
  function save(then){
    if(!ops.length){then&&then();return;}
    fetch('/api/patch',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({ops:ops})})
      .then(function(r){return r.json()}).then(function(j){
        if(!j.ok){toast('保存失败: '+(j.error||''));return;}
        ops=[];dirty=false;mark();toast('已保存');
        if(j.reload){setTimeout(function(){location.reload()},250);} else if(then){then();}
      }).catch(function(e){toast('保存失败: '+e)});
  }
  document.getElementById('savebtn').onclick=function(){save()};
  var exp=document.getElementById('exportbtn');
  if(exp)exp.onclick=function(){
    save(function(){
      exp.textContent='导出中…';exp.disabled=true;
      fetch('/api/export/pptx',{method:'POST'}).then(function(r){return r.json()}).then(function(j){
        exp.textContent='导出 PPTX';exp.disabled=false;
        toast(j.ok?('已导出: '+j.path):('导出失败: '+(j.error||'')),6000);
        if(j.ok)window.__lastExport=j.path;
      }).catch(function(e){exp.textContent='导出 PPTX';exp.disabled=false;toast('导出失败: '+e)});
    });
  };
  var pdf=document.getElementById('pdfbtn');
  if(pdf)pdf.onclick=function(){save(function(){
    pdf.disabled=true;pdf.textContent='生成中…';
    fetch('/api/export/pdf',{method:'POST'}).then(function(r){return r.json()}).then(function(j){
      pdf.disabled=false;pdf.textContent='导出 PDF';toast(j.ok?('已导出: '+j.path):('失败: '+(j.error||'')),6000);
    }).catch(function(e){pdf.disabled=false;pdf.textContent='导出 PDF';toast('失败: '+e)});
  })};
  var add=document.getElementById('addelem');
  if(add)add.onclick=function(){
    var m=el; var s=document.querySelector('.sel')||document.querySelector('.slide');
    var sid=s.dataset.id;
    push({op:'element',action:'add',slide:sid,element:{type:'text',x:1.2,y:1.2,w:4,h:0.7,
      paragraphs:[{runs:[{text:'新文本框'}],size:18,color:'#111111'}]}});
    toast('已添加文本框，保存后生效');
  };
  var th=document.getElementById('themebtn');
  if(th)th.onchange=function(){push({op:'deck',set:{theme:th.value}});toast('主题已改，保存后生效')};
  var lo=document.getElementById('layoutsel');
  if(lo)lo.onchange=function(){push({op:'slide',action:'layout',index:+lo.dataset.index,layout:lo.value});toast('版式已换，保存后生效')};

  window.addEventListener('beforeunload',function(e){if(dirty){e.preventDefault();e.returnValue='';}});
  document.addEventListener('keydown',function(e){
    if((e.ctrlKey||e.metaKey)&&e.key==='s'){e.preventDefault();save();}
    if(e.key==='Delete'&&sel&&!(e.target.isContentEditable)){var b=document.getElementById('del');b&&b.click();}
  });
  buildPanel(null);
  var first=document.querySelector('.slide');if(first)first.scrollIntoView();
})();
"""

PAGE_CSS = """
*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{background:#20242B;font-family:'Microsoft YaHei','Segoe UI',system-ui,sans-serif;color:#E6E9EF}
.slide{position:relative;width:var(--sw);height:var(--sh);margin:18px auto;transform-origin:top center;
       overflow:hidden;box-shadow:0 8px 30px rgba(0,0,0,.45);color-scheme:light;flex:none}
.slide>.el{position:absolute}
.slide .el{overflow:hidden}
.tx{width:100%;height:100%;overflow:hidden}
.tx p{margin:0;white-space:pre-wrap;word-break:break-word}
.imw{overflow:hidden;background:#0000}
.imw img{width:100%;height:100%}
.ph{display:flex;align-items:center;justify-content:center;background:#E9ECF2;color:#8A93A3;
    border:2px dashed #B9C2D0;font-size:13px;border-radius:4px}
.shp{display:block}
.tbl{border-collapse:collapse;table-layout:fixed;font-family:inherit}
.tbl td{border:1px solid rgba(0,0,0,.10);overflow:hidden;word-break:break-word}
.unknown{padding:6px;background:#fee;color:#900;font-size:12px}
body.present{background:#000;overflow:hidden}
body.present .slide{margin:0;position:fixed;left:0;top:0}
#hud{position:fixed;right:14px;bottom:12px;z-index:99999;background:rgba(0,0,0,.55);color:#fff;
     padding:6px 12px;border-radius:20px;font-size:12px;letter-spacing:.04em;display:flex;gap:10px;align-items:center}
body.present #hud{opacity:.35}
#hud button{background:none;border:0;color:#fff;cursor:pointer;font-size:12px;opacity:.85}
#hud button:hover{opacity:1}
.editwrap{display:flex;gap:0;align-items:flex-start}
#stage{flex:1;min-width:0;padding:8px 0 60px}
#props{width:288px;position:sticky;top:0;height:100vh;overflow:auto;background:#181C22;
       border-left:1px solid #2C333D;padding:14px;font-size:12.5px;color:#C9D1DC;flex:none}
#props h4{margin:0 0 8px;font-size:11px;letter-spacing:.12em;color:#7C8899;text-transform:uppercase;font-weight:600}
#props .hd{font-weight:600;color:#fff;margin-bottom:10px;font-size:13px}
#props .sect{border-top:1px solid #2C333D;padding:12px 0}
#props .f{display:flex;align-items:center;justify-content:space-between;gap:8px;margin:6px 0}
#props .f>span{color:#8E9AAA;font-size:12px}
#props input,#props select{background:#232A33;border:1px solid #333C48;color:#E6E9EF;border-radius:6px;
       padding:5px 7px;font-size:12px;width:104px;font-family:inherit}
#props input[type=color]{padding:0;height:26px;width:44px}
#props .grid4{display:grid;grid-template-columns:1fr 1fr;gap:6px}
#props .grid4 input{width:100%}
#props .btn{background:#2C3542;border:1px solid #3A4553;color:#DCE3EC;border-radius:6px;padding:6px 10px;
       cursor:pointer;font-size:12px;margin:3px 4px 3px 0;font-family:inherit}
#props .btn:hover{background:#3A4553}
#props .btn.danger{color:#FF9A8B;border-color:#5A3A36}
#props .hint{color:#7C8899;font-size:11.5px;line-height:1.5;margin:0 0 8px}
#props .row{display:flex;gap:6px;flex-wrap:wrap}
#props .ph{background:none;border:0;color:#7C8899;padding:0 0 10px;display:block;height:auto}
.sel{outline:2px solid #4C8DFF;outline-offset:0}
.el:hover{outline:1px dashed rgba(76,141,255,.65)}
.el.locked{cursor:default}
.el:not(.locked){cursor:move}
.mode-edit .locked{outline:none!important}
.tx:focus{outline:none;background:rgba(76,141,255,.06)}
#toast{position:fixed;left:50%;transform:translateX(-50%);bottom:54px;background:rgba(20,24,30,.94);
   color:#fff;padding:9px 16px;border-radius:8px;font-size:13px;opacity:0;transition:.2s;pointer-events:none;z-index:99999}
#toast.on{opacity:1}
#topbar{position:fixed;top:0;left:0;right:0;height:44px;background:#12161B;border-bottom:1px solid #2C333D;
   display:flex;align-items:center;gap:10px;padding:0 14px;z-index:9999;font-size:12.5px}
#topbar .t{color:#fff;font-weight:600;margin-right:auto;font-size:13px}
#topbar button,#topbar select{background:#232A33;border:1px solid #333C48;color:#DCE3EC;border-radius:6px;
   padding:5px 10px;cursor:pointer;font-size:12px;font-family:inherit}
#topbar button.pri{background:#2F6FE4;border-color:#2F6FE4;color:#fff}
#topbar .sep{width:1px;height:20px;background:#2C333D}
.mode-edit #stage{padding-top:56px}
.mode-edit .slide{margin-bottom:26px}
.tag{position:absolute;left:-52px;top:0;color:#67707E;font-size:11px;text-align:right;width:46px}
@media print{
  html,body{background:#fff;margin:0;padding:0}
  #hud,#topbar,#props,#toast,.tag,.notes{display:none!important}
  .editwrap{display:block!important}
  #stage{padding:0!important;margin:0!important;display:block!important}
  .slide{transform:none!important;margin:0!important;box-shadow:none!important;
         break-after:page;page-break-after:always;break-inside:avoid}
  .slide:last-child{break-after:auto;page-break-after:auto}
}
"""


def render_html(deck: dict, theme: dict, deck_path: str, mode: str = "view",
                title: str | None = None, pdf: bool = False) -> str:
    resolved = M.resolve_deck(deck, theme)
    W, H = M.slide_size(resolved)
    sw, sh = _px(W), _px(H)
    href_prefix = ""
    slides_html = []
    for i, slide in enumerate(resolved.get("slides", [])):
        els = sorted(slide.get("elements") or [], key=lambda e: float(e.get("z", 0)))
        body = "".join(render_element(e, theme, href_prefix) for e in els)
        bg = _background_css(slide.get("background"), theme)
        notes = slide.get("notes") or ""
        slides_html.append(
            f'<section class="slide" data-index="{i}" data-id="{_esc(slide.get("id"))}" '
            f'data-layout="{_esc(slide.get("layout") or "")}" '
            f'data-notes="{_esc(notes)}" style="{bg};color:{theme["colors"].get("text", "#111111")}">'
            f'{body}<div class="notes" hidden>{_esc(notes)}</div></section>')
    deck_title = title or deck["deck"].get("title") or "演示文稿"
    scripts = _asset("editor.js", EDIT_JS) if mode == "edit" else VIEWER_JS
    extra_css = _asset("editor.css", "") if mode == "edit" else ""
    topbar = ""
    if mode == "edit":
        topbar = (
            '<div id="topbar"><span class="t">' + _esc(deck_title) + '</span>'
            '<button id="addelem">+ 文本框</button>'
            '<span class="sep"></span>'
            '<select id="themebtn" title="整体主题">' + _theme_options(theme) + '</select>'
            '<select id="layoutsel" title="本页版式"></select>'
            '<span class="sep"></span>'
            '<button id="savebtn">保存</button>'
            '<button id="exportbtn" class="pri">导出 PPTX</button>'
            '<button id="pdfbtn">导出 PDF</button>'
            '</div>')
    body_class = "mode-edit" if mode == "edit" else ""
    hud = ("<div id=\"hud\"><span id=\"pageinfo\">1 / " + str(len(slides_html)) +
           "</span><button id=\"ovfbtn\" title=\"检查文字溢出\">溢出检查</button>"
           "<button onclick=\"window.print()\">打印/PDF</button>"
           "<button id=\"presentbtn\">放映</button></div>")
    return f"""<!doctype html>
<html lang="{_esc(deck['deck'].get('language', 'zh'))}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{_esc(deck_title)}</title>
<style>
{PAGE_CSS}
{extra_css}
/* @page cannot take a var(), so the stage size is written in literally here */
@media print {{ @page {{ size: {W}in {H}in; margin: 0; }} }}
:root{{--sw:{sw}px;--sh:{sh}px;--pw:{W}in;--ph:{H}in}}
</style>
</head>
<body class="{body_class}">
{topbar}
<div class="editwrap">
<div id="stage">
{''.join(slides_html)}
</div>
{'<aside id="props"></aside>' if mode == 'edit' else ''}
</div>
{hud}
<div id="toast"></div>
<script>
window.__STAGE_W={sw};window.__STAGE_H={sh};window.__DECK_PATH={json.dumps(os.path.abspath(deck_path))};
window.__PAGE_COUNT={len(slides_html)};
</script>
<script>
{scripts}
</script>
</body>
</html>
"""


def _asset(name: str, fallback: str) -> str:
    """Editor assets live in <skill>/assets/ so they are editable without touching Python."""
    root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    path = os.path.join(root, "assets", name)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as fh:
                return fh.read()
        except Exception:
            return fallback
    return fallback


def _theme_options(current_theme: dict) -> str:
    root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    with open(os.path.join(root, "assets", "themes.json"), encoding="utf-8") as fh:
        themes = json.load(fh)
    out = []
    for tid, t in themes.items():
        sel = " selected" if tid == current_theme.get("id") else ""
        out.append(f'<option value="{_esc(tid)}"{sel}>{_esc(t.get("name", tid))}</option>')
    return "".join(out)


def write_html(deck: dict, theme: dict, deck_path: str, out_path: str | None = None,
               mode: str = "view") -> str:
    out_path = out_path or os.path.splitext(deck_path)[0] + ".html"
    html_doc = render_html(deck, theme, deck_path, mode=mode)
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(html_doc)
    return out_path
