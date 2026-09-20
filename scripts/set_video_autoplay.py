# -*- coding: utf-8 -*-
"""COM 后处理：所有视频对象设为 放映时自动播放 + 循环播放（复刻 HTML 的自动循环行为）。"""
import os, sys
import win32com.client

# 用法: python set_video_autoplay.py <pptx路径>
import sys as _sys
pptx = _sys.argv[1] if len(_sys.argv) > 1 else os.path.join('.', 'pptx输出', '复赛PPT.pptx')
app = win32com.client.Dispatch('PowerPoint.Application')
pres = app.Presentations.Open(pptx, WithWindow=False)
count = 0
for slide in pres.Slides:
    for shp in slide.Shapes:
        if shp.Type == 16:  # msoMedia
            ps = shp.AnimationSettings.PlaySettings
            ps.PlayOnEntry = True          # 放映到该页即自动播放
            ps.LoopUntilStopped = True     # 循环直到翻页
            ps.PauseAnimation = False
            count += 1
            print('slide %d: %s -> autoplay+loop' % (slide.SlideIndex, shp.Name))
pres.Save()
pres.Close()
app.Quit()
print('done,', count, 'videos set')
