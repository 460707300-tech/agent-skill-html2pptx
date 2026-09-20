# -*- coding: utf-8 -*-
"""pptx -> PDF -> 逐页 PNG（PyMuPDF 渲染，矢量精确，不受系统 DPI 影响）。
用法: python export_pdf.py <pptx路径> [输出目录] [pdf中间文件路径]"""
import os, sys
import fitz
import win32com.client

pptx = sys.argv[1] if len(sys.argv) > 1 else os.path.join('.', 'pptx输出', '复赛PPT.pptx')
outdir = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(pptx)), '..', 'analysis', 'pdf_png')
pdf = sys.argv[3] if len(sys.argv) > 3 else os.path.join(os.path.dirname(os.path.abspath(pptx)), '..', 'analysis', '复赛PPT.pdf')
os.makedirs(outdir, exist_ok=True)

app = win32com.client.Dispatch("PowerPoint.Application")
pres = app.Presentations.Open(os.path.abspath(pptx), ReadOnly=True, WithWindow=False)
pres.SaveAs(os.path.abspath(pdf), 32)
pres.Close(); app.Quit()

doc = fitz.open(pdf)
zoom = 1920 / 960  # PDF 页 960x540pt -> 1920x1080
for i, pg in enumerate(doc, 1):
    pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    pix.save(os.path.join(outdir, 'p%d.png' % i))
print('rendered', len(doc), 'pages ->', outdir)
