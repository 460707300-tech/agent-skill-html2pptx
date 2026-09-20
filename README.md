# html2pptx — Agent Skill: HTML Slides → Pixel-Accurate, Fully Editable PPTX

> Convert 1920×1080-style HTML slide decks into PowerPoint files that are **position-exact, fully editable, and keep videos playing** — measured from a real browser, not guessed from CSS.

**Works with any coding agent** — Claude Code, Codex CLI, Cursor, Cline, Windsurf, or any assistant that can read a markdown playbook and run shell commands. The skill is just a markdown playbook plus standalone Python scripts; there is no runtime lock-in.

一个把 1920×1080 风格的 HTML 幻灯片转化为 **位置精准、文字可编辑、视频可播放** 的 PowerPoint 的 Agent Skill。（中文说明在下方 / Chinese guide below.）

---

## English

### Why not "screenshot every page"

Typical html→pptx tools take one of two approaches: full-page screenshots (fast but nothing is editable) or rule-based HTML parsing (editable but flex/grid positions drift). This skill takes a third path:

1. **Measure, don't parse** — Playwright renders the deck in a real Chromium at 1920×1080 and captures every element's `getBoundingClientRect`. The layout engine does the math.
2. **Convert in layers** — decorative layers (grid patterns, gradient overlays, frames) are baked into one background image; text/images/videos/shapes become native PPTX objects, so text stays editable.
3. **Render-compare loop** — every generated slide is exported (PowerPoint→PDF→PNG) and compared side-by-side with the original page. Nothing ships until it matches.
4. **Keep media alive** — embedded mp4s auto-play and loop in presentation mode; JS image carousels are re-encoded into looping videos with PowerPoint's own encoder.

### What gets converted

| HTML element | PPTX result |
|---|---|
| Text (inline mixed-color runs, letter-spacing, weights) | Editable text boxes, multi-run colors preserved |
| Images (object-fit contain/cover) | Embedded pictures, aspect-correct |
| mp4 videos | Embedded video objects, **auto-play + loop on slideshow** |
| JS image carousels | Re-encoded as auto-looping carousel videos (captions baked in) |
| Gradient cards / rounded boxes | Shapes with gradient fills |
| clip-path pyramids / trapezoids | Freeform polygons |
| Outlined (stroked) display type | Text outline via raw XML injection |
| Grid patterns / gradient overlays | Baked full-page background image |
| Entrance animations / glows / shadows | Dropped (no faithful pptx equivalent) |

### Install

```bash
pip install playwright python-pptx pymupdf pywin32 pillow
playwright install chromium
# Microsoft PowerPoint (desktop) is required for the compare/export loop
```

Install the skill: clone this repo to your agent's skill directory, e.g. `~/.claude/skills/html2pptx/` for Claude Code, or simply point your agent at `SKILL.md` — it is a plain markdown playbook.

### Usage

Working directory layout:

```
<project>/
├── 原html/            ← copy of the HTML deck (pages/*.html + assets/), frozen afterwards
├── analysis/          ← measured JSON, baked backgrounds, comparison images
└── pptx输出/          ← final pptx
```

Ask your agent to "convert this HTML deck to pptx" and it follows the 7-step playbook in `SKILL.md`:

```bash
python scripts/extract_elements.py 原html analysis      # 1. measure every element -> analysis/pN.json
python scripts/bake_bg.py 原html analysis               # 2. bake decorative background
python scripts/build_pptx.py .                          # 3. build (incremental: . 3 7)
python scripts/export_pdf.py "pptx输出/复赛PPT.pptx"    # 4. render slides to PNG (via PDF)
python scripts/compare.py 原html analysis               # 5. side-by-side comparison cmp_pN.png
python scripts/set_video_autoplay.py "pptx输出/复赛PPT.pptx"  # 6. autoplay + loop all videos
python scripts/make_carousel_mp4.py .                   # 7. (optional) carousels -> looping videos
```

`SKILL.md` also carries a 15-item pitfall list (the 144px/in font-scale constant, flex item blockification, COM background three-step, anti-clipping rules…), learned the hard way — read it before touching the renderer.

### Fonts

HTML decks ship their fonts as woff2; PPTX can only reference installed system fonts, so the renderer maps JetBrains Mono→Consolas, Noto Sans SC→Microsoft YaHei, Archivo→Arial (editable at the top of `build_pptx.py`). Any Windows machine renders it correctly; for 100% fidelity install the original fonts and save with PowerPoint's embedded-fonts option.

---

## 中文说明

把 1920×1080 风格的 HTML 幻灯片转化为**位置精准、文字可编辑、视频可播放**的 PowerPoint。

**与常见工具的区别**：不做整页截图贴图（不可编辑），也不靠规则解析 CSS（flex 布局算不准）——用无头浏览器真渲染一遍，逐元素实测坐标，再分层转化：装饰层烘焙成背景图，文字/图片/视频转成原生 PPTX 对象；每页生成后走 PowerPoint→PDF→PNG 与原页并排比对，过关才算完成。

**适用于任何能读 markdown、能跑命令的编程 Agent**（Claude Code / Codex CLI / Cursor / Cline / Windsurf 等）。

**转化能力**：文字混排→可编辑文本框｜图片→按比例嵌入｜mp4→内嵌且放映自动播放循环｜JS 轮播→转自动循环视频｜渐变卡片/金字塔/描边大字→形状绘制｜底纹遮罩→背景烘焙｜入场动画/光晕→按约定放弃。

**安装**：`pip install playwright python-pptx pymupdf pywin32 pillow && playwright install chromium`（本机需装桌面版 PowerPoint 用于比对闭环），本仓库放入 agent 的 skills 目录即可。

**用法**：目录布局见上方英文部分；对 Agent 说「把这份 HTML 幻灯片转成 pptx」，或手动逐步执行上方 7 条命令。`SKILL.md` 含 15 条实战踩坑清单，改渲染器前必读。

**字体**：HTML 字体随网页打包（woff2），pptx 只能引用系统字体，默认回退映射为 Consolas/微软雅黑/Arial（可改）。任意 Windows 机器（含导播系统）不会乱码；要 100% 一致可安装原字体并用 PowerPoint 嵌入字体保存。

## License

MIT
