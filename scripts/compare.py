# -*- coding: utf-8 -*-
"""生成 原页截图 vs pptx导出 的并排对比图，供逐页自检。"""
import os, sys
from playwright.sync_api import sync_playwright
from PIL import Image

# 用法: python compare.py <html目录> <analysis目录> [页码...]
import sys as _sys
BASE = _sys.argv[1] if len(_sys.argv) > 1 else '.'
A = _sys.argv[2] if len(_sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(BASE)), 'analysis')
P = os.path.join(A, 'pdf_png')

def shot_html(n):
    out = os.path.join(A, 'p%d_shot.png' % n)
    if os.path.exists(out): return out
    with sync_playwright() as pw:
        b = pw.chromium.launch()
        pg = b.new_page(viewport={'width': 1920, 'height': 1080})
        pg.goto('file:///' + (BASE + '\\pages\\p%d.html' % n).replace('\\', '/'))
        pg.add_style_tag(content="*,*::before,*::after{transition:none!important;animation:none!important}")
        pg.evaluate("document.querySelector('.slide').classList.add('visible','active')")
        pg.wait_for_load_state('networkidle')
        pg.wait_for_timeout(800)
        pg.screenshot(path=out)
        b.close()
    return out

_pages = [int(a) for a in _sys.argv[3:]] or list(range(1, 17))
for n in _pages:
    left = Image.open(shot_html(n)).convert('RGB').resize((1280, 720))
    right = Image.open(os.path.join(A, 'pdf_png', 'p%d.png' % n)).convert('RGB').resize((1280, 720))
    comp = Image.new('RGB', (1280, 1452), (40, 40, 40))
    comp.paste(left, (0, 0)); comp.paste(right, (0, 732))
    comp.save(os.path.join(A, 'cmp_p%d.png' % n))
print('done')
