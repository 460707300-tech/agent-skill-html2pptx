# -*- coding: utf-8 -*-
"""轮播视频 v2：PowerPoint CreateVideo 只出 16:9，所以临时演示文稿直接用 16:9（1280x720, 1in=96px），
深色底 + 图片 contain 满幅，避免白边。P2 字幕条仍烘焙进视频。"""
import os, time
import win32com.client
from PIL import Image

# 用法: python make_carousel_mp4.py <项目根目录>（含 原html/assets，视频输出到 原html/assets/video/）
import sys as _sys
ROOT = _sys.argv[1] if len(_sys.argv) > 1 else '.'
ASSETS = os.path.join(ROOT, '原html', 'assets')
OUT = os.path.join(ASSETS, 'video')

def contain(iw, ih, bw, bh):
    s = min(bw / iw, bh / ih)
    return (bw - iw * s) / 2, (bh - ih * s) / 2, iw * s, ih * s

def make_video(name, imgs, advance, caps=None):
    app = win32com.client.Dispatch('PowerPoint.Application')
    pres = app.Presentations.Add(WithWindow=False)
    pres.PageSetup.SlideWidth = 720.0   # 10in * 72pt，16:9
    pres.PageSetup.SlideHeight = 405.0  # 5.625in * 72pt
    W, H = 1280.0, 720.0                # 逻辑 px（1in=96px）
    for i, p in enumerate(imgs):
        slide = pres.Slides.Add(i + 1, 12)
        slide.FollowMasterBackground = False  # 不关这条，背景修改无效（保持母版白底）
        bg = slide.Background.Fill
        bg.Solid()  # 先定类型再给色，顺序反了会保持白色背景
        bg.ForeColor.RGB = 13 * 65536 + 19 * 256 + 25
        iw, ih = Image.open(p).size
        ox, oy, w, h = contain(iw, ih, W, H)
        slide.Shapes.AddPicture(p, False, True, ox / W * 720, oy / H * 405, w / W * 720, h / H * 405)
        if caps and caps[i]:
            cap_h = 40
            bar = slide.Shapes.AddShape(1, 0, (H - cap_h) / H * 405, 720, cap_h / H * 405)
            bar.Fill.ForeColor.RGB = 7 * 65536 + 10 * 256 + 13
            bar.Fill.Transparency = 0.18
            bar.Line.Visible = False
            tf = bar.TextFrame
            tf.MarginLeft = 16 / W * 720; tf.MarginRight = 0
            tf.MarginTop = 0; tf.MarginBottom = 0
            tf.VerticalAnchor = 3
            r = tf.TextRange
            r.Text = caps[i]
            r.Font.Size = 10  # ~13.5px@96dpi
            r.Font.Name = 'Consolas'
            r.Font.Color.RGB = 232 * 65536 + 244 * 256 + 246
        st = slide.SlideShowTransition
        st.EntryEffect = 1793
        st.AdvanceOnTime = True
        st.AdvanceTime = advance[i]
        st.Duration = 0.5
    mp4 = os.path.join(OUT, name + '.mp4')
    pres.CreateVideo(mp4, True, 3, 720, 30, 85)
    while pres.CreateVideoStatus in (1, 2):
        time.sleep(2)
    pres.Close()
    app.Quit()
    print(name, '-> %.1fMB' % (os.path.getsize(mp4) / 1048576))

if __name__ == '__main__':
    # ↓↓↓ 示例配置（本项目实际用法），换你自己的图片清单/节奏/字幕 ↓↓↓
    certs = ['个人工作照.jpg', '25-陕西省技术能手证书.jpg', '24-陕西省青年岗位能手证书.jpg']
    make_video('carousel_p2_certs',
               [os.path.join(ASSETS, 'img', '证书', c) for c in certs], [3, 3, 3],
               caps=['BIM 工作场景 · 数智科创中心', '陕西省技术能手 · 2023', '陕西省优秀青年岗位能手 · 2022'])
    slides19 = ['slide-%02d.jpg' % i for i in range(1, 20)]
    make_video('carousel_p14_slides',
               [os.path.join(ASSETS, 'img', '普及课件', '翻页', s) for s in slides19], [3] + [1] * 18)
