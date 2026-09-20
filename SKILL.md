---
name: html2pptx
description: Convert HTML slide decks (1920×1080-style web presentations) into faithful, fully editable .pptx files. Zero-drift geometry (measured in a real browser via Playwright), editable text runs, embedded images, autoplay+loop videos, JS carousels converted to looping videos, special graphics (pyramids/stroked display type/dashed connectors) drawn as native shapes, and a per-page render-compare self-check loop. Works with any coding agent (Claude Code, Codex CLI, Cursor, Cline, …). Use when asked to "convert HTML slides to pptx" or "turn this web presentation into PowerPoint".
---

# HTML Slide Deck → Faithful PPTX

## Core principles (do not deviate)

1. **Measure, don't parse.** Never derive positions by reading HTML/CSS (flex/grid cannot be hand-computed reliably). Render the deck in Chromium at 1920×1080 via Playwright and capture every element's `getBoundingClientRect`. This is the foundation of the whole pipeline.
2. **Bake backgrounds, convert content.** Decorative layers (grid patterns, gradient overlays, frames) become one baked background PNG. Text/images/videos/shapes are converted individually at their measured rects; text stays editable.
3. **Render-compare loop.** Every generated page must be exported (PowerPoint COM → PDF → PyMuPDF PNG) and compared side-by-side with the original page. A page ships only when it matches. No loop, no quality.

## Workflow (in order)

### Step 0: Prepare
- Check dependencies: `python -c "import playwright, pptx, fitz, win32com.client, PIL"`; install what's missing (`playwright install chromium`; `pip install pywin32` for COM).
- Microsoft PowerPoint (desktop) must be installed — the compare/export loop drives it via COM.
- Check installed fonts (`ls C:/Windows/Fonts`) and set the font mapping at the top of `build_pptx.py` (`FONT_MONO/FONT_CN/FONT_LAT`).
- Create the working layout:
  ```
  <project>/html-deck/   ← copy of the deck (pages/*.html + assets/), frozen afterwards
  <project>/analysis/    ← measured JSON, baked backgrounds, comparisons
  <project>/pptx-out/    ← final pptx
  ```
- Write a `WORKLOG.md` (task / method / progress / per-page log / pitfalls) and update it after every page — it is the only lifeline if the context window dies mid-job.

### Step 1: Measure every page → `analysis/pN.json`
```bash
python scripts/extract_elements.py html-deck analysis
```
The JSON records per element: tag/class/rect[x,y,w,h]/text runs (inline mixed-color segments with size, weight, color, letter-spacing)/image & video src/border sides/padding/measured text width (textW)/pseudo-elements.

### Step 2: Bake backgrounds → `analysis/bg_common.png` (+ `bg_p1.png` if page 1 has a photo background)
```bash
python scripts/bake_bg.py html-deck analysis
```
Principle: render a minimal page containing only the shared decorative layers (grid + frame [+ cover photo]) and screenshot it.

### Step 3: Build → `build_pptx.py`
```bash
python scripts/build_pptx.py .          # all pages
python scripts/build_pptx.py . 3 7      # rebuild only pages 3 and 7
```
The generic renderer already handles: text runs / images (contain & cover math) / videos via add_movie / rounded corners & one-side borders / clip-path polygons / SVG children / single-line no-wrap / snug-box widening by measured textW. Per-page HTML specials (stroked hollow display type, dashed ✕ connector, axis-cross SVG mark, …) go in as page-specific branches in `build_page`.

### Step 4: Per-page compare loop → `export_pdf.py` + `compare.py`
```bash
python scripts/export_pdf.py "pptx-out/deck.pptx"
python scripts/compare.py html-deck analysis [pages...]
```
COM saves PDF first (**never export PNG via SaveAs — text renders fat on high-DPI machines**), PyMuPDF rasterizes at 1920px, `compare.py` writes `analysis/cmp_pN.png` (top = original HTML screenshot, bottom = PPTX render). Read each comparison and check: completeness / geometry / image-vs-video distinction / no clipped or overlapping text. Fix `build_pptx.py`, rebuild the page, re-compare.

### Step 5: Video autoplay post-process → `set_video_autoplay.py`
```bash
python scripts/set_video_autoplay.py "pptx-out/deck.pptx"
```
Sets every video object to **auto-play + loop on slideshow**, matching the HTML deck's behavior.

### Step 6: Carousels → looping videos (only if the deck has JS image carousels)
Adapt `scripts/make_carousel_mp4.py` (`__main__` section: image list, per-slide dwell time, captions). Principle: build a temp presentation via COM (16:9, dark background, images contained, timed transitions) and encode with `Presentation.CreateVideo`; then embed the mp4 back at the original rect (16:9 shape centered in the original box). Re-run steps 3–5 afterwards.

### Step 7: Deliver
Final audit (page count / video count / file size) → update `WORKLOG.md` → report the file path with the per-page comparison images for review.

## ⚠️ Pitfall list (all real; re-read before writing any code)

1. **Font scale: 1px = 1/144 inch, so px→pt is ×0.5, NOT ×0.75** (1920px→13.333in = 144px/in). Using 0.75 makes every text 1.5× too big — the single worst bug of the original job.
2. A raw page screenshot is all black: slides are `visibility:hidden` until you `classList.add('visible','active')`.
3. Inject `*{transition:none!important;animation:none!important}` before measuring, or `.reveal` elements (opacity:0) get dropped.
4. Flex/grid children are **blockified** (computed display = block): inline-absorption cannot rely on display alone. Rule: if an element has any block-level child, absorb nothing and keep only its own text nodes.
5. Inline mixed-color headings (`<h1>text<span>colored</span></h1>`) must become ONE textbox with multiple runs — otherwise text overlaps or vanishes.
6. PowerPoint **autoshapes default to centered text**; set left alignment explicitly. Plain textboxes are unaffected.
7. object-fit:contain images need the display rect recomputed from the image's natural size (fit_contain), otherwise they stretch.
8. border-radius 50% → oval (with centered text); one-side borders → thin rectangles; don't mix up the side formulas.
9. Match class names by token set (`'bignum' in cls.split()`); class attributes are often multi-token ("bignum reveal").
10. Single-line text (box height < 2.6× font size) gets `word_wrap=False` to prevent re-wrapping under fallback fonts; snug filled boxes widen by measured textW ×1.15 to prevent clipping.
11. Export via PDF (SaveAs PNG renders fat text on high-DPI machines — use SaveAs 32 = PDF + PyMuPDF rasterize).
12. Any COM error: `taskkill //F //IM POWERPNT.EXE` first (stale processes break Open/Save).
13. `Presentation.CreateVideo` always outputs 16:9; setting a slide background via COM requires `slide.FollowMasterBackground = False` → `Fill.Solid()` → `ForeColor.RGB` (order matters or it stays white); poster frames: size the `<video>` element by videoWidth/videoHeight before screenshotting or you get white bars.
14. Terminal GBK mojibake is display-only; files are UTF-8. Don't "fix" files because the terminal looks wrong.
15. Report before fixing: when the client reports a bug, analyze root cause and list the fix plan first; fixes you discover yourself get reported after the fact.

## Fonts (put this in the delivery note)

HTML decks ship fonts as bundled woff2; PPTX can only reference installed system fonts. Default mapping: JetBrains Mono→Consolas, Noto Sans SC→Microsoft YaHei, Archivo→Arial (adjust to the target machine). No Windows machine (including broadcast/director systems) will show tofu. For 100% fidelity: install the original fonts and save with PowerPoint's embedded-fonts option.
