---
name: html2pptx
description: 把 HTML 版 PPT/幻灯片（1920×1080 风格的网页演示）忠实转化为可编辑的 .pptx。位置零误差（无头浏览器实测）、文字可编辑、图片嵌入、视频内嵌自动播放、轮播图转动态视频、特殊图案（金字塔/描边字/虚线）形状绘制、逐页渲染比对自检。当用户要求「html转pptx」「网页幻灯片转PowerPoint」「把html演示转成ppt」时使用。
---

# HTML 幊灯片 → PPTX 忠实转化

## 核心理念（三条，别走偏）

1. **位置靠实测不靠解析**：不用规则解析 HTML/CSS（flex/grid 算不准），用 playwright 让页面在 1920×1080 真渲染，逐元素抓 `getBoundingClientRect`。这是整条流水线的地基。
2. **背景烘焙、内容转化**：网格底纹/渐变遮罩/图框等装饰层截成 1 张背景 PNG；文字/图片/视频/形状按实测矩形逐个转化，文字保持可编辑。
3. **渲染-比对闭环**：每页生成后必须走 PowerPoint COM→PDF→PyMuPDF 出图，与原页并排比对，过关才进下一页。没有闭环就没有质量。

## 工作流程（按序执行）

### 第 0 步：准备
- 确认依赖：`python -c "import playwright, pptx, fitz, win32com.client, PIL"`；缺啥装啥（playwright 还需 `playwright install chromium`；win32com 需 `pip install pywin32`）
- 确认本机装有 PowerPoint（COM 导出用）；检查字体：`ls C:/Windows/Fonts`，决定字体映射（见 build_pptx.py 顶部 FONT_MONO/FONT_CN/FONT_LAT）
- 在目标项目建工作目录结构：
  ```
  <项目>/原html/     ← html 副本（pages/*.html + assets/），之后不再改动
  <项目>/analysis/   ← 实测 JSON、背景图、比对图（中间产物）
  <项目>/pptx输出/   ← 最终 pptx
  ```
- 写一份 `工作日志.md`（总体任务/方法/进度/逐页登记/踩坑），每完成一页更新一次——这是上下文断了能续命的唯一保障

### 第 1 步：逐页实测提取 → `analysis/pN.json`
在项目根目录运行：
```bash
python scripts/extract_elements.py 原html analysis
```
产出的 JSON 里每个元素有：tag/类/矩形[x,y,w,h]/文字 runs（含行内混排分段与颜色字号）/图片视频 src/边框分侧/内边距/实测文字宽度 textW/伪元素。

### 第 2 步：烘焙背景 → `analysis/bg_common.png`（+ `bg_p1.png` 若首页有背景照片）
运行 `python scripts/bake_bg.py 原html analysis`。原理：起一个只含 grid-layer+frame（+首页背景照片）的最小页面截图。

### 第 3 步：生成 → `build_pptx.py`
脚本按约定布局自动定位（原html/analysis/pptx输出），字体映射在文件头改。运行：
```bash
python scripts/build_pptx.py .          # 全部页
python scripts/build_pptx.py . 3 7      # 只重建指定页
```
通用渲染规则已在脚本内实现（文本 runs/图片 contain-cover/视频 add_movie/圆角与单侧边框/clip-path 多边形/SVG 子形状/单行防折行/色块按 textW 加宽）。每页若有 HTML 特例（描边空心字、✕虚线、轴网标记等），在 `build_page` 里按页号加特例分支。

### 第 4 步：逐页比对自检 → `export_pdf.py` + `compare.py`
```bash
taskkill //F //IM POWERPNT.EXE ; python export_pdf.py ; python compare.py [页码...]
```
COM 先存 PDF（**不要用 SaveAs PNG 导出，文字会虚胖**），PyMuPDF 渲染成 1920px PNG，compare.py 生成「上=原HTML截图 下=PPTX渲染」的并排图 `analysis/cmp_pN.png`。逐页 Read 检查：元素齐全/位置贴合/图片视频区分正确/无截字无重叠。有问题改 build_pptx.py 后重跑该页。

### 第 5 步：视频自动播放后处理 → `set_video_autoplay.py`
运行 `python scripts/set_video_autoplay.py "pptx输出/复赛PPT.pptx"`：所有视频对象设为放映时自动播放+循环（复刻 HTML 的自动播放行为）。

### 第 6 步：轮播图转动态视频（可选，页面里有 JS 轮播时）
复制 `scripts/make_carousel_mp4.py`，按目标轮播的图片清单/节奏/字幕改 `__main__` 参数。原理：用 PowerPoint COM 建临时演示（16:9 深色底、图片 contain、自动换页），`CreateVideo` 编码成 mp4，嵌回主文件原位置（16:9 形状在原框内居中）。编码后重跑 build + autoplay。

### 第 7 步：交付
整体检查（页数/视频数/文件大小）→ 更新工作日志 → 告知路径（>30MB 发不了飞书，给本机路径）→ 附逐页对比图供用户验收。

## ⚠️ 踩坑清单（每条都真实踩过，写码前重读）

1. **字号换算：1px = 1/144 英寸，px→pt 是 ×0.5 不是 ×0.75**（1920px→13.333in = 144px/in）。用 0.75 全部文字大 1.5 倍，是本项目最大的 bug。
2. 分页 html 直接截图全黑：slide 默认 `visibility:hidden`，必须先 `classList.add('visible','active')`。
3. 提取前注入 `*{transition:none!important;animation:none!important}`，否则 reveal 元素 opacity=0 被漏掉。
4. flex/grid 子项被 CSS **blockify**（display 计算成 block）：行内吸收判断不能只看 display；规则=「元素有块级子元素就不吸收、只保留自身文字节点」。
5. 行内混排文字（如 `<h1>从「用AI」到<span>「造工具」</span></h1>`）必须用 runs 合成一个文本框，否则文字重叠/丢失。
6. PowerPoint **形状默认文字居中**，必须显式设左对齐；textbox 无此问题。
7. object-fit:contain 的图片要按图片真实宽高比重算显示矩形（fit_contain），否则拉伸变形。
8. 圆角 50% → 椭圆（文字要居中）；单侧边框 → 细矩形，位置公式别写串。
9. 类名匹配用 token 集合（`'bignum' in cls.split()`），类名常是 "bignum reveal" 多段。
10. 单行文本（盒高 < 2.6×字号）设 `word_wrap=False` 防折行；贴合色块按 canvas 实测文字宽 ×1.15 加宽，防截字。
11. COM 导出 PDF（ppSaveAsPNG=18 出的 PNG 在高 DPI 机器文字虚胖，弃用；SaveAs 32=PDF + PyMuPDF 渲染才准）。
12. COM 报错先 `taskkill //F //IM POWERPNT.EXE`（残留进程导致 Open 失败/保存被拒）。
13. lark-cli 发图发文件必须用相对路径；>30MB 会被飞书拒收。
14. 终端 GBK 乱码是显示问题，文件是 UTF-8，别被终端骗了去"修"文件。
15. CreateVideo 固定输出 16:9；临时演示设背景必须 `FollowMasterBackground=False` → `Fill.Solid()` → `ForeColor.RGB`（顺序错=白底）；封面帧用 playwright 按 videoWidth/Height 设定元素尺寸再截图。

## 字体说明（写进交付说明）
HTML 的字体随文件打包（woff2），pptx 只能用系统字体回退：JetBrains Mono→Consolas、Noto Sans SC→微软雅黑、Archivo→Arial（按本机 Fonts 目录实测调整）。全 Windows 机器（含导播系统）都不会乱码。想要 100% 一致：装原字体 + PowerPoint「嵌入字体」保存。
