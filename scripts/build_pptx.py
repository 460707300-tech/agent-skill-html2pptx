# -*- coding: utf-8 -*-
"""
HTML PPT → PPTX 渲染器。
数据源：analysis/pN.json（playwright 实测）+ analysis/bg_*.png（烘焙背景）+ 原html/assets 素材。
用法：python build_pptx.py [页码...]   不带参数 = 全部 16 页
"""
import json, os, re, sys
from PIL import Image
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn

# 用法: python build_pptx.py [项目根目录] [页码...]
# 项目根布局: <root>/原html/{pages,assets} + <root>/analysis/ + 输出到 <root>/pptx输出/
ROOT = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].isdigit() else '.'
_args = [a for a in sys.argv[1:] if a.isdigit()]
A = os.path.join(ROOT, 'analysis')
ASSETS = os.path.join(ROOT, '原html', 'assets')
OUTDIR = os.path.join(ROOT, 'pptx输出')
os.makedirs(OUTDIR, exist_ok=True)
TMP = os.path.join(A, '_fit')
os.makedirs(TMP, exist_ok=True)

EMU_PER_PX = 6350          # 1920px = 13.333in = 144px/inch
PX2PT = 0.5                # 72pt/inch ÷ 144px/inch（不是 0.75！0.75 会让所有文字大 1.5 倍）

# 字体映射（本机实测：无 JetBrains Mono / Noto Sans SC）
FONT_MONO = 'Consolas'
FONT_CN = 'Microsoft YaHei'
FONT_LAT = 'Arial'

def map_family(f):
    f = (f or '').strip('"')
    if 'JetBrains' in f or 'Consolas' in f or 'Mono' in f: return FONT_MONO
    if 'Archivo' in f: return FONT_LAT
    return FONT_CN

def parse_color(s):
    """'rgb(a,b,c)' / 'rgba(a,b,c,a)' / '#xxx' -> (RGBColor, alpha|None)"""
    if not s: return None, None
    s = s.strip()
    m = re.match(r'rgba?\(([^)]+)\)', s)
    if m:
        parts = [p.strip() for p in m.group(1).split(',')]
        r, g, b = int(parts[0]), int(parts[1]), int(parts[2])
        a = float(parts[3]) if len(parts) > 3 else None
        return RGBColor(r, g, b), a
    m = re.match(r'#([0-9a-fA-F]{6})$', s)
    if m:
        v = m.group(1)
        return RGBColor(int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)), None
    return None, None

def set_fill_alpha(fill_obj, alpha):
    """在 solidFill 的 srgbClr 上注入 <a:alpha>"""
    srgb = fill_obj.fore_color._xFill.find(qn('a:solidFill') + '/' + qn('a:srgbClr'))
    if srgb is None:
        el = fill_obj.fore_color._xFill
        srgb = el.find(qn('a:srgbClr'))
    if srgb is None:
        return
    a = srgb.find(qn('a:alpha'))
    if a is None:
        a = srgb.makeelement(qn('a:alpha'), {})
        srgb.append(a)
    a.set('val', str(int(alpha * 100000)))

def set_ea_font(run, name):
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', name)

def set_char_spacing(run, css_ls):
    """css letter-spacing px -> spc (1/100 pt)"""
    m = re.match(r'([\d.]+)px', css_ls or '')
    if not m: return
    spc = int(round(float(m.group(1)) * PX2PT * 100))
    if spc:
        run._r.get_or_add_rPr().set('spc', str(spc))

def set_run_outline(run, color, width_px):
    """文字描边（空心数字用）：rPr 注入 a:ln"""
    rPr = run._r.get_or_add_rPr()
    ln = rPr.makeelement(qn('a:ln'), {'w': str(int(width_px * PX2PT * 12700))})
    sf = rPr.makeelement(qn('a:solidFill'), {})
    clr = rPr.makeelement(qn('a:srgbClr'), {'val': '%02X%02X%02X' % color})
    sf.append(clr); ln.append(sf)
    rPr.insert(0, ln)

def px(v): return Emu(int(round(v * EMU_PER_PX)))

def fit_contain(img_path, box_w, box_h):
    """object-fit:contain 的实际显示矩形"""
    try:
        iw, ih = Image.open(img_path).size
    except Exception:
        return 0, 0, box_w, box_h
    s = min(box_w / iw, box_h / ih)
    w, h = iw * s, ih * s
    return (box_w - w) / 2, (box_h - h) / 2, w, h

def fit_cover(img_path, box_w, box_h, key):
    """object-fit:cover -> 裁剪出与盒子同比例的临时图"""
    out = os.path.join(TMP, key + '_cover.jpg')
    if os.path.exists(out): return out
    im = Image.open(img_path).convert('RGB')
    iw, ih = im.size
    s = max(box_w / iw, box_h / ih)
    w, h = iw * s, ih * s
    x0, y0 = (w - box_w) / 2 / s, (h - box_h) / 2 / s
    im = im.crop((int(x0), int(y0), int(x0 + box_w / s), int(y0 + box_h / s)))
    im.save(out)
    return out

# clip-path 多边形（minipy / tier / pyw 通用）
POLY = {
    's3': [(50, 0), (100, 100), (0, 100)],
    't3': [(50, 0), (100, 100), (0, 100)],
    's2': [(15, 0), (85, 0), (100, 100), (0, 100)],
    't2': [(15, 0), (85, 0), (100, 100), (0, 100)],
    's1': [(11, 0), (89, 0), (100, 100), (0, 100)],
    't1': [(11, 0), (89, 0), (100, 100), (0, 100)],
}

def add_polygon(slide, el, shapes):
    toks = set((el['cls'] or '').split())
    pts = None
    for k, v in POLY.items():
        if k in toks:
            pts = v; break
    if pts is None: return False
    x, y, w, h = el['x'], el['y'], el['w'], el['h']
    xs = [x + w * p[0] / 100 for p in pts]
    ys = [y + h * p[1] / 100 for p in pts]
    fb = shapes.build_freeform(px(xs[0]), px(ys[0]), scale=1.0)
    fb.add_line_segments([(px(xs[i]), px(ys[i])) for i in range(1, len(pts))], close=True)
    shp = fb.convert_to_shape()
    shp.shadow.inherit = False
    bg = el.get('bg', '')
    if bg.startswith('linear-gradient'):
        cols = re.findall(r'rgba?\(([^)]+)\)', bg)
        c1, a1 = parse_color('rgba(%s)' % cols[0]) if cols else (None, None)
        c2, a2 = parse_color('rgba(%s)' % cols[1]) if len(cols) > 1 else (None, None)
        f = shp.fill; f.gradient()
        f.gradient_stops[0].color.rgb = c1; f.gradient_stops[0].position = 0.0
        f.gradient_stops[1].color.rgb = c2; f.gradient_stops[1].position = 1.0
        if a1 is not None: set_fill_alpha_gs(f.gradient_stops[0], a1)
        if a2 is not None: set_fill_alpha_gs(f.gradient_stops[1], a2)
        try: f.gradient_angle = 90  # 自上而下
        except Exception: pass
    else:
        c, al = parse_color(bg)
        if c:
            shp.fill.solid(); shp.fill.fore_color.rgb = c
            if al is not None: set_fill_alpha(shp.fill, al)
    shp.line.fill.background()
    return True

def set_fill_alpha_gs(stop, alpha):
    srgb = stop._gs.find(qn('a:srgbClr'))
    if srgb is None: return
    a = srgb.makeelement(qn('a:alpha'), {'val': str(int(alpha * 100000))})
    srgb.append(a)

def add_text(slide, el, shapes, dx=0.0):
    x, y, w, h = el['x'] + dx, el['y'], el['w'], el['h']
    def _pxv(s):
        m = re.match(r'([\d.]+)px', s or '')
        return float(m.group(1)) if m else 0
    bg = el.get('bg', '')
    sides = el.get('borderSides') or {}
    full = all(k in sides for k in ('top', 'right', 'bottom', 'left')) and len(sides)
    needs_box = bool(bg) or full
    m = re.match(r'([\d.]+)px', el.get('radius', '0px') or '')
    rad = float(m.group(1)) if m else 0
    is_oval = (el.get('radius', '').strip() in ('50%',) or el.get('radius', '').startswith('999'))
    if needs_box:
        if is_oval:
            shape_type = MSO_SHAPE.OVAL
        elif rad >= 4:
            shape_type = MSO_SHAPE.ROUNDED_RECTANGLE
        else:
            shape_type = MSO_SHAPE.RECTANGLE
        # 回退字体偏宽：贴合色块按实测文字宽度加宽（保持 x 不变向右扩展）
        max_px0 = max([float(re.match(r'([\d.]+)', r['size']).group(1)) for r in (el.get('runs') or [{'size': '16px'}]) if r.get('size')] or [16])
        if not is_oval and el['h'] <= max_px0 * 2.6 and el.get('textW'):
            pad_lr = _pxv(el.get('pad', ['0px'])[0]) + _pxv(el.get('pad', ['0px', '0px', '0px'])[2] if len(el.get('pad', [])) > 2 else 0)
            need_w = el['textW'] * 1.15 + pad_lr  # 15% 余量：回退字体略宽（PX2PT 已修为 0.5，文字恢复正常大小）
            if need_w > w:
                w = need_w
        tb = slide.shapes.add_shape(shape_type, px(x), px(y), px(max(w, 4)), px(max(h, 4)))
        tb.shadow.inherit = False
        if not is_oval and rad >= 4:
            try: tb.adjustments[0] = min(0.5, rad / min(w, h))
            except Exception: pass
        # 背景填充
        if bg and not bg.startswith(('linear', 'radial')):
            c, al = parse_color(bg)
            if c:
                tb.fill.solid(); tb.fill.fore_color.rgb = c
                if al is not None: set_fill_alpha(tb.fill, al)
        elif bg.startswith('linear-gradient'):
            cols = re.findall(r'rgba?\(([^)]+)\)', bg)
            if cols:
                c1, _ = parse_color('rgba(%s)' % cols[0])
                c2, _ = parse_color('rgba(%s)' % cols[1]) if len(cols) > 1 else (None, None)
                f = tb.fill; f.gradient()
                f.gradient_stops[0].color.rgb = c1; f.gradient_stops[0].position = 0.0
                if c2: f.gradient_stops[1].color.rgb = c2; f.gradient_stops[1].position = 1.0
                try: f.gradient_angle = 90
                except Exception: pass
        else:
            tb.fill.background()
        if full:
            v = list(sides.values())[0]
            c, al = parse_color(v['color'])
            tb.line.color.rgb = c
            tb.line.width = Pt(v['w'] * PX2PT)
            if al is not None: set_fill_alpha_ln(tb.line, al)
        else:
            tb.line.fill.background()
    else:
        tb = slide.shapes.add_textbox(px(x), px(y), px(max(w, 4)), px(max(h, 4)))
        if len(sides) == 1:  # 单侧边框 -> 细矩形
            k, v = list(sides.items())[0]
            c, al = parse_color(v['color'])
            if k == 'top':      lb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x), px(y), px(w), px(v['w']))
            elif k == 'bottom': lb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x), px(y + h - v['w']), px(w), px(v['w']))
            elif k == 'left':   lb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x), px(y), px(v['w']), px(h))
            else:               lb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x + w - v['w']), px(y), px(v['w']), px(h))
            lb.fill.solid(); lb.fill.fore_color.rgb = c
            if al is not None: set_fill_alpha(lb.fill, al)
            lb.line.fill.background(); lb.shadow.inherit = False
    tf = tb.text_frame
    # 单行文本（盒高 < 2.6×最大字号）不换行，避免字体回退变宽导致折行
    max_px = max([float(re.match(r'([\d.]+)', r['size']).group(1)) for r in (el.get('runs') or [{'size': '16px'}]) if r.get('size')] or [16])
    tf.word_wrap = el['h'] > max_px * 2.6
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_ANCHOR.TOP
    pad = el.get('pad') or ['0px'] * 4
    tf.margin_left = Emu(int(_pxv(pad[0]) * EMU_PER_PX))
    tf.margin_right = Emu(int(_pxv(pad[2]) * EMU_PER_PX))
    tf.margin_top = Emu(int(_pxv(pad[1]) * EMU_PER_PX))
    tf.margin_bottom = Emu(int(_pxv(pad[3]) * EMU_PER_PX))
    # 文本
    runs = el.get('runs') or []
    p = tf.paragraphs[0]
    first = True
    for r in runs:
        if r.get('br') and not first:
            p = tf.add_paragraph()
        if not r.get('t'): continue
        run = p.add_run(); run.text = r['t']
        f = run.font
        f.size = Pt(float(re.match(r'([\d.]+)', r['size']).group(1)) * PX2PT)
        f.bold = (r.get('weight') in ('700', '800', '900', 'bold'))
        f.italic = (r.get('style') == 'italic')
        c, al = parse_color(r.get('color'))
        if c:
            f.color.rgb = c
            if al is not None:
                srgb = f.color._xFill.find(qn('a:srgbClr'))
                if srgb is not None:
                    srgb.append(srgb.makeelement(qn('a:alpha'), {'val': str(int(al * 100000))}))
        fam = map_family(r.get('family'))
        f.name = fam
        set_ea_font(run, fam)
        set_char_spacing(run, r.get('ls'))
        if el.get('_outline'):
            set_run_outline(run, (27, 124, 136), 2)
            # 文字不填充（空心）：rPr 里把 solidFill 换成 noFill
            rPr = run._r.get_or_add_rPr()
            sf = rPr.find(qn('a:solidFill'))
            if sf is not None:
                rPr.remove(sf)
            nf = rPr.makeelement(qn('a:noFill'), {})
            rPr.insert(list(rPr).index(rPr.find(qn('a:ln'))) + 1 if rPr.find(qn('a:ln')) is not None else 0, nf)
        first = False
    align = el.get('align')
    if align == 'center':
        for pp in tf.paragraphs: pp.alignment = PP_ALIGN.CENTER
    elif align == 'right':
        for pp in tf.paragraphs: pp.alignment = PP_ALIGN.RIGHT
    elif align == 'left' or needs_box:  # 形状默认居中，必须显式压回左对齐
        for pp in tf.paragraphs: pp.alignment = PP_ALIGN.LEFT
    if is_oval:  # 圆点内文字居中
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for pp in tf.paragraphs: pp.alignment = PP_ALIGN.CENTER
    return tb

def set_fill_alpha_ln(line, alpha):
    srgb = line.color._xFill.find(qn('a:srgbClr'))
    if srgb is not None:
        srgb.append(srgb.makeelement(qn('a:alpha'), {'val': str(int(alpha * 100000))}))

def add_shape(slide, el, shapes):
    x, y, w, h = el['x'], el['y'], el['w'], el['h']
    bg = el.get('bg', '')
    if bg.startswith('radial-gradient'):
        return  # 光晕：放弃
    r = el.get('radius', '0px')
    m = re.match(r'([\d.]+)px', r or '')
    rad = float(m.group(1)) if m else 0
    if rad >= 4:
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(x), px(y), px(w), px(h))
        try: shp.adjustments[0] = min(0.5, rad / min(w, h))
        except Exception: pass
    else:
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x), px(y), px(w), px(h))
    shp.shadow.inherit = False
    if bg.startswith('linear-gradient'):
        cols = re.findall(r'rgba?\(([^)]+)\)', bg)
        c1, a1 = parse_color('rgba(%s)' % cols[0]) if cols else (None, None)
        c2, a2 = parse_color('rgba(%s)' % cols[1]) if len(cols) > 1 else (None, None)
        f = shp.fill; f.gradient()
        if c1: f.gradient_stops[0].color.rgb = c1; f.gradient_stops[0].position = 0.0
        if c2: f.gradient_stops[1].color.rgb = c2; f.gradient_stops[1].position = 1.0
        if a1 is not None: set_fill_alpha_gs(f.gradient_stops[0], a1)
        if a2 is not None: set_fill_alpha_gs(f.gradient_stops[1], a2)
        try: f.gradient_angle = 90
        except Exception: pass
    else:
        c, al = parse_color(bg)
        if c:
            shp.fill.solid(); shp.fill.fore_color.rgb = c
            if al is not None: set_fill_alpha(shp.fill, al)
        else:
            shp.fill.background()
    sides = el.get('borderSides') or {}
    full = all(k in sides for k in ('top', 'right', 'bottom', 'left'))
    if full and len(sides):
        v = sides['top']
        c, al = parse_color(v['color'])
        shp.line.color.rgb = c
        shp.line.width = Pt(v['w'] * PX2PT)
        if al is not None: set_fill_alpha_ln(shp.line, al)
    else:
        shp.line.fill.background()
        if len(sides) == 1:  # 单侧边框 -> 细矩形
            k, v = list(sides.items())[0]
            c, al = parse_color(v['color'])
            if k == 'top':      lb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x), px(y), px(w), px(v['w']))
            elif k == 'bottom': lb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x), px(y + h - v['w']), px(w), px(v['w']))
            elif k == 'left':   lb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x), px(y), px(v['w']), px(h))
            else:               lb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x + w - v['w']), px(y), px(v['w']), px(h))
            lb.fill.solid(); lb.fill.fore_color.rgb = c
            if al is not None: set_fill_alpha(lb.fill, al)
            lb.line.fill.background(); lb.shadow.inherit = False
    return shp

def add_svg_child(slide, el, shapes):
    a = el['svgAttrs']; t = el['tag'][4:]
    if t == 'circle':
        r = float(a.get('r', 0)); cx = float(a.get('cx', 0)) + el['x'] - r; cy = float(a.get('cy', 0)) + el['y'] - r
        shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, px(cx), px(cy), px(2 * r), px(2 * r))
        shp.shadow.inherit = False
        shp.fill.background()
        st = a.get('stroke'); sw = float(a.get('stroke-width', 1))
        if st:
            c, _ = parse_color(st); shp.line.color.rgb = c; shp.line.width = Pt(sw * PX2PT)
    elif t == 'line':
        x1 = float(a.get('x1', 0)) + el['x']; y1 = float(a.get('y1', 0)) + el['y']
        x2 = float(a.get('x2', 0)) + el['x']; y2 = float(a.get('y2', 0)) + el['y']
        ln = slide.shapes.add_connector(1, px(x1), px(y1), px(x2), px(y2))
        ln.shadow.inherit = False
        st = a.get('stroke')
        if st:
            c, _ = parse_color(st); ln.line.color.rgb = c
        ln.line.width = Pt(float(a.get('stroke-width', 1)) * PX2PT)
    elif t == 'text':
        tb = slide.shapes.add_textbox(px(el['x'] - 10), px(el['y']), px(el['w'] + 20), px(el['h'] + 4))
        tf = tb.text_frame; tf.word_wrap = False; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.auto_size = MSO_AUTO_SIZE.NONE
        run = tf.paragraphs[0].add_run(); run.text = el.get('svgText') or ''
        f = run.font
        f.size = Pt(float(a.get('font-size', 12)) * PX2PT)
        f.bold = a.get('font-weight') in ('700', '900')
        c, _ = parse_color(a.get('fill', '#fff'))
        if c: f.color.rgb = c
        fam = map_family(a.get('font-family', ''))
        f.name = fam; set_ea_font(run, fam)
        if a.get('text-anchor') == 'middle':
            tb.left = px(el['x'] - 30); tb.width = px(el['w'] + 60)
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    return

SKIP_CLS = {'grid-layer', 'frame', 'glow', 'bgimg', 'deck-controls', 'hint'}

def src_cert_poster():
    return make_poster(os.path.join(ASSETS, 'video', 'carousel_p2_certs.mp4'))

def src_slide01_poster():
    return make_poster(os.path.join(ASSETS, 'video', 'carousel_p14_slides.mp4'))

def _vdim(v):
    """视频尺寸（宽,高），避免重复探测失败时按 16:9"""
    try:
        with Image.open(make_poster(v)) as im:
            if im.width > 64:
                return float(im.width), float(im.height)
    except Exception:
        pass
    return 1280.0, 720.0

def build_page(prs, n):
    data = json.load(open(os.path.join(A, 'p%d.json' % n), encoding='utf-8'))
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    bg = os.path.join(A, 'bg_p1.png' if n == 1 else 'bg_common.png')
    slide.shapes.add_picture(bg, 0, 0, px(1920), px(1080))
    for el in data['elements']:
        tag, cls = el['tag'], el['cls']
        if cls in SKIP_CLS or any(cls.startswith(s + ' ') or cls == s for s in SKIP_CLS): continue
        # P2/P14 轮播图 -> 动态轮播视频（截图变会动的画面）；视频是 16:9，按 contain 居中放进原框
        if n == 2 and tag == 'img' and '个人工作照' in el['src']:
            v = os.path.join(ASSETS, 'video', 'carousel_p2_certs.mp4')
            ox, oy, w, h = fit_contain(_vdim(v), el['w'], el['h'])
            slide.shapes.add_movie(v, px(el['x'] + ox), px(el['y'] + oy), px(w), px(h),
                                   poster_frame_image=src_cert_poster(), mime_type='video/mp4')
            continue
        if n == 2 and el.get('cls') == 'cap' and el.get('runs') and 'BIM 工作场景' in el['runs'][0].get('t', ''):
            continue  # 字幕条已烘焙进轮播视频，静态字幕删除避免图文不符
        if n == 14 and tag == 'img' and 'slide-01' in el['src']:
            v = os.path.join(ASSETS, 'video', 'carousel_p14_slides.mp4')
            ox, oy, w, h = fit_contain(_vdim(v), el['w'], el['h'])
            slide.shapes.add_movie(v, px(el['x'] + ox), px(el['y'] + oy), px(w), px(h),
                                   poster_frame_image=src_slide01_poster(), mime_type='video/mp4')
            continue
        if tag == 'img':
            src = os.path.join(ASSETS, '..', el['src'].replace('../', ''))
            src = os.path.normpath(src)
            if not os.path.exists(src):
                print('  !! missing img', src); continue
            fit = el.get('fit', 'fill')
            if fit == 'contain':
                ox, oy, w, h = fit_contain(src, el['w'], el['h'])
                slide.shapes.add_picture(src, px(el['x'] + ox), px(el['y'] + oy), px(w), px(h))
            elif fit == 'cover':
                key = 'p%d_%s' % (n, re.sub(r'\W+', '_', os.path.basename(src)))
                slide.shapes.add_picture(fit_cover(src, el['w'], el['h'], key), px(el['x']), px(el['y']), px(el['w']), px(el['h']))
            else:
                slide.shapes.add_picture(src, px(el['x']), px(el['y']), px(el['w']), px(el['h']))
        elif tag == 'video':
            vsrc = os.path.normpath(os.path.join(ASSETS, '..', el['src'].replace('../', '')))
            poster = make_poster(vsrc)
            mv = slide.shapes.add_movie(vsrc, px(el['x']), px(el['y']), px(el['w']), px(el['h']),
                                        poster_frame_image=poster, mime_type='video/mp4')
        elif tag == 'svg':
            continue  # 子形状单独处理
        elif tag.startswith('svg:'):
            add_svg_child(slide, el, slide.shapes)
        elif el['cls'] and ({'t3', 't2', 't1'} & set(el['cls'].split()) or el['cls'].startswith('seg ')):
            if not add_polygon(slide, el, slide.shapes):
                add_shape(slide, el, slide.shapes)
        elif el.get('runs'):
            if 'bignum' in el['cls'].split():
                el['_outline'] = True
                el = dict(el)
                el['y'] = el['y'] - 90  # PPT 行框比 css line-height:.9 低 ~90px，上移补偿
            dx = 68 if (n == 1 and el['cls'] == 'eyebrow') else 0
            add_text(slide, el, slide.shapes, dx=dx)
        else:
            add_shape(slide, el, slide.shapes)
    # 伪元素装饰
    for ps in data['pseudos']:
        if n == 10 and 'bgap' in ps['el']:
            continue  # bgap 由下方特例绘制（虚线+标签），避免重复
        hx, hy, hw, hh = ps['host']
        content = ps['content']
        if content:  # 有文字的伪元素
            tb = slide.shapes.add_textbox(px(hx), px(hy), px(hw), px(hh))
            tf = tb.text_frame; tf.word_wrap = False; tf.auto_size = MSO_AUTO_SIZE.NONE
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            run = tf.paragraphs[0].add_run()
            run.text = content.strip('"')
            f = run.font
            f.size = Pt(float(re.match(r'([\d.]+)', ps['fontSize']).group(1)) * PX2PT)
            f.bold = ps['fontWeight'] in ('700', '900')
            c, _ = parse_color(ps['color'])
            if c: f.color.rgb = c
            # 符号字形（✦✕●等）用 Segoe UI Symbol，避免雅黑缺字形
            txt = content.strip('"')
            fam = 'Segoe UI Symbol' if any(ord(ch) > 0x2600 for ch in txt if ord(ch) < 0x4E00) else map_family('')
            f.name = fam; set_ea_font(run, fam)
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER if ps['position'] == 'absolute' else PP_ALIGN.LEFT
        elif n == 1 and ps['el'] == 'div.eyebrow' and ps['pseudo'] == '::before':
            pass  # 已由 eyebrow dx=68 处理（短线在原位由形状补）
    # P1 eyebrow 前置短线
    if n == 1:
        ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(84), px(412), px(52), px(1))
        ln.fill.solid(); ln.fill.fore_color.rgb = RGBColor(0x2F, 0xD8, 0xE8); ln.line.fill.background(); ln.shadow.inherit = False
    # P10 bgap（✕ 无法送达）
    if n == 10:
        ln = slide.shapes.add_connector(1, px(1626), px(383), px(1626), px(483))
        ln.line.color.rgb = RGBColor(0xFF, 0x8B, 0x6A); ln.line.width = Pt(1.5)
        ln.line._get_or_add_ln().append(ln.line._get_or_add_ln().makeelement(qn('a:prstDash'), {'val': 'dash'}))
        ln.shadow.inherit = False
        tb = slide.shapes.add_textbox(px(1526), px(413), px(200), px(40))
        tf = tb.text_frame; tf.word_wrap = False; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        run = tf.paragraphs[0].add_run(); run.text = '✕ 无法送达'
        f = run.font; f.size = Pt(14); f.bold = True; f.color.rgb = RGBColor(0xFF, 0x8B, 0x6A)
        f.name = FONT_CN; set_ea_font(run, FONT_CN)
        tb.fill.solid(); tb.fill.fore_color.rgb = RGBColor(0x0A, 0x0E, 0x12)
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    return slide

def make_poster(video_path):
    """用 headless chromium 抽视频首帧做封面图"""
    key = re.sub(r'\W+', '_', os.path.basename(video_path)) + '.png'
    out = os.path.join(TMP, key)
    if os.path.exists(out): return out
    try:
        from playwright.sync_api import sync_playwright
        html = ('<body style="margin:0;background:#0D1319"><video id=v style="display:block" src="file:///' +
                video_path.replace('\\', '/') + '" muted></video></body>')
        tmp = os.path.join(TMP, '_v.html')
        open(tmp, 'w').write(html)
        with sync_playwright() as pw:
            b = pw.chromium.launch(args=['--allow-file-access-from-files', '--autoplay-policy=no-user-gesture-required'])
            pg = b.new_page(viewport={'width': 1280, 'height': 720})
            pg.goto('file:///' + tmp.replace('\\', '/'))
            vw, vh = pg.evaluate("() => { const v = document.getElementById('v'); return [v.videoWidth, v.videoHeight]; }")
            pg.evaluate("() => { const v = document.getElementById('v'); v.style.width = v.videoWidth + 'px'; v.style.height = v.videoHeight + 'px'; }")
            pg.set_viewport_size({'width': max(int(vw), 320), 'height': max(int(vh), 180)})
            pg.evaluate("document.getElementById('v').currentTime = 0.5")
            pg.wait_for_timeout(800)
            pg.locator('video').screenshot(path=out)
            b.close()
        return out
    except Exception as e:
        print('  poster fail', video_path, e)
        return None

def main():
    pages = [int(a) for a in _args] or list(range(1, 17))
    prs = Presentation()
    prs.slide_width = px(1920)
    prs.slide_height = px(1080)
    for n in pages:
        build_page(prs, n)
        print('P%d built' % n)
    out = os.path.join(OUTDIR, '复赛PPT.pptx')
    prs.save(out)
    print('saved:', out)

if __name__ == '__main__':
    main()
