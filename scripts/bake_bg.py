# -*- coding: utf-8 -*-
"""烘焙每页共用背景：蓝图网格底纹 + 图框（P1 另含 28% 透明背景照片 + 渐变遮罩）。"""
import os, urllib.parse
from playwright.sync_api import sync_playwright

# 用法: python bake_bg.py <html目录(含assets/)> [输出目录]
import sys as _sys
BASE = _sys.argv[1] if len(_sys.argv) > 1 else '.'
OUT = _sys.argv[2] if len(_sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(BASE)), 'analysis')
os.makedirs(OUT, exist_ok=True)

HTML_COMMON = """<!DOCTYPE html><html><head><meta charset="UTF-8">
<link rel="stylesheet" href="./assets/fonts/fonts.css">
<link rel="stylesheet" href="./assets/css/deck-base.css">
<style>html,body{background:#0A0E12}</style></head>
<body><main class="deck-stage" style="width:1920px;height:1080px;position:absolute;transform:none">
<section class="slide bp active visible"><div class="grid-layer"></div><div class="frame"></div></section>
</main></body></html>"""

HTML_P1 = HTML_COMMON.replace(
    '<div class="grid-layer"></div><div class="frame"></div>',
    '<div class="bgimg" style="position:absolute;inset:0;background:url(./assets/img/cover-bg.jpg) center/cover no-repeat;opacity:.28"></div>'
    '<div style="position:absolute;inset:0;background:linear-gradient(180deg,rgba(10,14,18,.5),rgba(10,14,18,.2) 45%,rgba(10,14,18,.65))"></div>'
    '<div class="grid-layer"></div><div class="frame"></div>')

tmp1 = os.path.join(BASE, '_bake_common.html')
tmpp = os.path.join(BASE, '_bake_p1.html')
open(tmp1, 'w', encoding='utf-8').write(HTML_COMMON)
open(tmpp, 'w', encoding='utf-8').write(HTML_P1)

with sync_playwright() as pw:
    b = pw.chromium.launch()
    pg = b.new_page(viewport={'width': 1920, 'height': 1080})
    for src, out in [(tmp1, 'bg_common.png'), (tmpp, 'bg_p1.png')]:
        pg.goto('file:///' + urllib.parse.quote(src.replace('\\', '/')))
        pg.wait_for_load_state('networkidle')
        pg.wait_for_timeout(400)
        pg.screenshot(path=os.path.join(OUT, out))
        print('baked', out)
    b.close()
os.remove(tmp1); os.remove(tmpp)
print('done')
