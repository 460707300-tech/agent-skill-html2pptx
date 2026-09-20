# html2pptx — Claude Code Skill：HTML 幻灯片 → 可编辑 PPTX 忠实转化

> HTML slides → pixel-accurate, fully editable .pptx — measured, not guessed.

把 1920×1080 风格的 HTML 演示文稿（网页版 PPT）转化为 **位置精准、文字可编辑、视频可播放** 的 PowerPoint 文件。

## 为什么不是"整页截图贴图"

市面工具通常是两种思路：整页截图（快但全不可编辑）或规则解析 HTML（可编辑但 flex/grid 位置跑偏）。
本 skill 走第三条路：

1. **实测代替解析** — playwright 让页面在 1920×1080 真渲染，逐元素抓 `getBoundingClientRect`，布局引擎的复杂度零手工计算
2. **分层转化** — 网格底纹/渐变遮罩烘焙为背景图保视觉；文字/图片/视频/形状转原生对象保可编辑
3. **渲染-比对闭环** — 每页生成后 PowerPoint→PDF→PNG 与原页并排比对，过关才进下一页
4. **动态素材** — 演示视频内嵌且放映自动播放+循环；JS 图片轮播用 PowerPoint 编码成自动循环视频嵌回原位

## 能转化什么

| HTML 元素 | PPTX 呈现 |
|---|---|
| 文字（含行内双色混排、字距、字重） | 可编辑文本框，多 run 保色 |
| 图片（object-fit contain/cover） | 嵌入图片，比例不变形 |
| mp4 视频 | 内嵌视频对象，**放映自动播放+循环** |
| JS 图片轮播 | 转成自动循环的轮播视频（含字幕条） |
| CSS 渐变色块/卡片/圆角 | 形状 + 渐变填充 |
| clip-path 金字塔/梯形 | 自由多边形 |
| 文字描边空心大字 | 文字描边（XML 注入） |
| 网格底纹/渐变遮罩 | 烘焙为整页背景图 |
| 入场动画/光晕/阴影 | 放弃（pptx 无法等价呈现） |

## 安装

```bash
# 依赖
pip install playwright python-pptx pymupdf pywin32 pillow
playwright install chromium
# 本机需装有 Microsoft PowerPoint（COM 导出比对用）
```

安装 skill：把本仓库克隆到 `~/.claude/skills/html2pptx/`（Windows 为 `C:\Users\<你>\.claude\skills\html2pptx\`）。

## 项目布局（转化时的工作目录）

```
<项目根>/
├── 原html/            ← html 副本（pages/*.html + assets/），转化期间不再改动
├── analysis/          ← 实测 JSON、背景图、比对图（中间产物）
└── pptx输出/          ← 最终 pptx
```

## 使用（配合 Claude Code）

对 Claude 说「把这份 html 幻灯片转成 pptx」即可，skill 会按 7 步流程走：
实测提取 → 背景烘焙 → 生成 → 逐页比对 → 视频自动播放 → 轮播动态化 → 交付。

也可以手动逐步执行（在项目根目录）：

```bash
python scripts/extract_elements.py 原html analysis      # 1. 逐元素实测 → analysis/pN.json
python scripts/bake_bg.py 原html analysis               # 2. 烘焙背景
python scripts/build_pptx.py .                          # 3. 生成（可带页码增量重建：. 3 7）
python scripts/export_pdf.py "pptx输出/复赛PPT.pptx"    # 4. 渲染出图
python scripts/compare.py 原html analysis               # 5. 上下并排比对图 cmp_pN.png
python scripts/set_video_autoplay.py "pptx输出/复赛PPT.pptx"  # 6. 视频自动播放+循环
python scripts/make_carousel_mp4.py .                   # 7.（可选）轮播图转动态视频
```

`SKILL.md` 里有完整的 15 条踩坑清单（字号换算 144px/in、flex 子项 blockify、COM 背景三步、防截字规则……），写渲染器前必读。

## 字体说明

HTML 的字体（woff2）随网页打包，pptx 只能用系统字体回退（JetBrains Mono→Consolas、Noto Sans SC→微软雅黑、Archivo→Arial，可在 `build_pptx.py` 顶部改映射）。全 Windows 机器播放不会乱码；要 100% 一致可安装原字体并用 PowerPoint「嵌入字体」保存。

## License

MIT
