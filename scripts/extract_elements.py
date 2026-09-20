# -*- coding: utf-8 -*-
"""逐页提取 HTML 渲染后的元素矩形 / 文字 runs / 素材信息，输出 JSON 供 pptx 生成用。

要点：
- slide 必须加 .visible/.active（否则 visibility:hidden 全黑）
- 注入 CSS 禁用 transition/animation（否则 reveal 元素 opacity=0 被漏掉）
- 文字元素带 runs（DOM 顺序的行内混排段），行内子元素标记 absorbed 不再单独输出
- flex/grid 容器的子元素不吸收（它们由 flex 定位，各自独立输出）
"""
import json, os, urllib.parse
from playwright.sync_api import sync_playwright

# 用法: python extract_elements.py <html目录(含pages/与assets/)> [输出目录]
import sys as _sys
BASE = _sys.argv[1] if len(_sys.argv) > 1 else '.'
OUT = _sys.argv[2] if len(_sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(BASE)), 'analysis')
os.makedirs(OUT, exist_ok=True)

JS = r"""
() => {
  const slide = document.querySelector('.slide');
  slide.classList.add('visible','active');
  const absorbed = new Set();
  const results = [];

  const styleOf = (el) => getComputedStyle(el);

  // 收集 el 内的行内文字段；只下钻：文本节点、<br>、父级非 flex/grid 下的 inline 子元素
  const collectRuns = (el, runs) => {
    const cs = styleOf(el);
    const isFlex = cs.display === 'flex' || cs.display === 'grid' || cs.display.startsWith('inline-flex') || cs.display.startsWith('inline-grid');
    // 有块级子元素 -> 子元素全部自行输出，只保留 el 自身文字
    let hasBlockChild = false;
    for (const c of Array.from(el.children)) {
      const ccs = styleOf(c);
      if (ccs.position === 'absolute' || ccs.position === 'fixed' || ccs.display === 'none') continue;
      if (!ccs.display.startsWith('inline')) { hasBlockChild = true; break; }
    }
    for (const node of Array.from(el.childNodes)) {
      if (node.nodeType === 3) {
        const t = node.textContent;
        if (t.trim()) runs.push({ t: t, color: cs.color, weight: cs.fontWeight, size: cs.fontSize,
                                  style: cs.fontStyle, family: cs.fontFamily.split(',')[0],
                                  ls: cs.letterSpacing, br: false });
      } else if (node.nodeType === 1) {
        const tag = node.tagName.toLowerCase();
        const ncs = styleOf(node);
        const pos = ncs.position;
        if (tag === 'br') { runs.push({ br: true }); continue; }
        if (pos === 'absolute' || pos === 'fixed') continue;
        if (['img','video','svg','canvas','iframe'].includes(tag)) continue;
        if (ncs.display === 'none' || parseFloat(ncs.opacity) === 0) continue;
        const inline = ncs.display.startsWith('inline');
        if (inline && !isFlex && !hasBlockChild) {
          absorbed.add(node);
          if (tag === 'br') { runs.push({ br: true }); continue; }
          collectRuns(node, runs);
        }
        // block 子元素 / flex item：不吸收，由外层作为独立元素输出
      }
    }
  };

  const walk = (el) => {
    if (el.classList.contains('deck-controls') || el.classList.contains('hint')) return;
    const cs = getComputedStyle(el);
    if (cs.display === 'none' || cs.visibility === 'hidden' || parseFloat(cs.opacity) === 0) return;
    const r = el.getBoundingClientRect();
    const tag = el.tagName.toLowerCase();
    const isImg = tag === 'img', isVideo = tag === 'video', isSvg = tag === 'svg';
    const hasBg = cs.backgroundImage !== 'none' || (cs.backgroundColor !== 'rgba(0, 0, 0, 0)' && cs.backgroundColor !== 'transparent');
    const hasBorder = ['Top','Right','Bottom','Left'].some(s => parseFloat(cs['border'+s+'Width']) > 0);
    const runs = [];
    if (!isImg && !isVideo && !isSvg) collectRuns(el, runs);
    const hasRuns = runs.some(s => !s.br && s.t && s.t.trim());
    const meaningful = isImg || isVideo || isSvg || hasRuns || hasBg || hasBorder;
    if (r.width > 0 && r.height > 0 && meaningful && !absorbed.has(el)) {
      const item = {
        tag, cls: el.className && el.className.baseVal !== undefined ? el.className.baseVal : (el.className || ''),
        id: el.id || '',
        x: Math.round(r.x*10)/10, y: Math.round(r.y*10)/10,
        w: Math.round(r.width*10)/10, h: Math.round(r.height*10)/10, z: cs.zIndex,
      };
      if (isImg) { item.src = el.getAttribute('src'); item.fit = cs.objectFit; }
      if (isVideo) { item.src = el.getAttribute('src'); }
      if (hasRuns) { item.runs = runs.filter(s => !s.br || runs.length > 1); item.align = cs.textAlign;
                     item.pad = [cs.paddingLeft, cs.paddingTop, cs.paddingRight, cs.paddingBottom];
                     // 用回退字体实测文字总宽（pptx 里 Consolas/雅黑 比 JetBrains Mono/Noto 宽）
                     try {
                       const cvx = document.createElement('canvas').getContext('2d');
                       let tw = 0;
                       for (const r of item.runs) {
                         if (!r.t) continue;
                         const fam = /JetBrains|Consolas|Mono/i.test(r.family) ? 'Consolas' : 'Microsoft YaHei';
                         cvx.font = (r.weight === '400' ? '' : r.weight + ' ') + r.size + ' ' + fam;
                         tw += cvx.measureText(r.t).width + (parseFloat(r.ls) || 0) * r.t.length;
                       }
                       item.textW = Math.round(tw * 10) / 10;
                     } catch (e) {} }
      if (hasBg) { item.bg = cs.backgroundColor !== 'rgba(0, 0, 0, 0)' ? cs.backgroundColor : cs.backgroundImage.slice(0, 300); }
      if (hasBorder) {
        item.borderSides = {};
        for (const s of ['Top','Right','Bottom','Left']) {
          const w = parseFloat(cs['border'+s+'Width']);
          if (w > 0) item.borderSides[s.toLowerCase()] = { w: w, color: cs['border'+s+'Color'] };
        }
        item.radius = cs.borderRadius;
      }
      if (isSvg) item.svgBox = true;
      results.push(item);
    }
    if (isSvg) {
      Array.from(el.querySelectorAll('circle,line,rect,path,polygon,text,polyline,ellipse')).forEach(ch => {
        const cr = ch.getBoundingClientRect();
        results.push({ tag: 'svg:' + ch.tagName.toLowerCase(), cls: '', id: '',
          x: Math.round(cr.x*10)/10, y: Math.round(cr.y*10)/10, w: Math.round(cr.width*10)/10, h: Math.round(cr.height*10)/10, z: '1',
          svgAttrs: ch.getAttribute ? Array.from(ch.attributes).reduce((o,a)=>(o[a.name]=a.value,o),{}) : {},
          svgText: ch.tagName.toLowerCase() === 'text' ? ch.textContent : undefined });
      });
      return;
    }
    Array.from(el.children).forEach(walk);
  };
  Array.from(slide.children).forEach(walk);

  // 伪元素：有文字内容的 / 1px 装饰线的
  const pseudos = [];
  Array.from(slide.querySelectorAll('*')).forEach(el => {
    if (absorbed.has(el)) return;
    for (const p of ['::before', '::after']) {
      const cs = getComputedStyle(el, p);
      if (!cs.content || cs.content === 'none' || cs.content === 'normal') continue;
      const r = el.getBoundingClientRect();
      if (r.width === 0 && r.height === 0) continue;
      const er = el.getBoundingClientRect();
      let box = null;
      // 用 Range 取伪元素实际盒
      try {
        const s = window.getSelection(); s.removeAllRanges();
        // 伪元素无法直接 Range，退回估算：全宽伪元素取文字左侧；这里记录宿主矩形+样式
      } catch(e) {}
      pseudos.push({ el: el.tagName.toLowerCase() + '.' + (el.className || ''), pseudo: p,
        content: cs.content === '""' ? '' : cs.content.slice(0, 80),
        host: [Math.round(er.x), Math.round(er.y), Math.round(er.width), Math.round(er.height)],
        bg: cs.background.slice(0, 140), color: cs.color, size: cs.width + 'x' + cs.height,
        fontSize: cs.fontSize, fontWeight: cs.fontWeight, position: cs.position,
        borderLeft: cs.borderLeftWidth + ' ' + cs.borderLeftStyle + ' ' + cs.borderLeftColor,
        display: cs.display, alignSelf: cs.alignSelf });
    }
  });
  return { slide: document.title, elements: results, pseudos };
}
"""

pages = ['p%d.html' % i for i in range(1, 17)]
with sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page(viewport={'width': 1920, 'height': 1080})
    for p in pages:
        url = 'file:///' + urllib.parse.quote((BASE + '\\pages\\' + p).replace('\\', '/'))
        page.goto(url)
        page.add_style_tag(content="*,*::before,*::after{transition:none!important;animation:none!important}")
        page.wait_for_load_state('networkidle')
        page.wait_for_timeout(500)
        data = page.evaluate(JS)
        out = os.path.join(OUT, p.replace('.html', '.json'))
        with open(out, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        print(p, '->', len(data['elements']), 'elements,', len(data['pseudos']), 'pseudos')
    browser.close()
print('done')
