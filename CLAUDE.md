# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project purpose

Slide generator for Burkholder/Grout/Palisca *A History of Western Music* (10th ed.), 39 chapters, bilingual (繁中 + English), 16:9. Each chapter lives in `chXX_name.js`, builds `ChXX_Name.pptx` via [pptxgenjs](https://github.com/gitbrent/PptxGenJS); LibreOffice converts to `ChXX_Name.pdf`. Only PDFs (and README.md) are committed — the repo is published on GitHub for students to download.

## Build pipeline

```bash
# 1. Generate .pptx
node ch26_romantic_classical.js

# 2. Convert to PDF (requires LibreOffice installed as `soffice`)
soffice --headless --convert-to pdf Ch26_Romantic_Classical.pptx

# 3a. Render specific page(s) for visual verification
pdftoppm -r 70 -png -f 4 -l 4 Ch26_Romantic_Classical.pdf /tmp/verify/p
# 3b. Render every page (required before claiming a chapter done)
pdftoppm -r 70 -png Ch26_Romantic_Classical.pdf /tmp/verify/p
# Then Read each PNG to check layout
```

Standard 7-step per-chapter workflow (memory `project_workflow.md`): JS → YouTube links → PPTX → PDF → push to GitHub → update README.md → clean tmp files.

## Reference format: Ch26 (definitive template)

`ch26_romantic_classical.js` is the format standard. All 39 chapters conform to it. Required structural elements:

1. **Cover** — chapter title page (not textbook intro)
2. **Outline** — per-chapter TOC; Ch26 uses two slides (14 entries × 2); Ch33–39 use one two-column slide (~13 entries)
3. **Overview** — chapter concept summary (row-table format on light background)
4. **Content slides** — two-column panel layout (varies per chapter length)
5. **NAWM slides** — one per Norton Anthology piece, each with a `youtu.be/...` listening link
6. **Timeline** — dated events (light background, two-column)
7. **Key Terms** — glossary grid (light background, two-column)

**NAWM sequence by chapter:** Ch31=169a/b, Ch32=170–172, Ch33=173–175 (Jazz/Pop only; Ch33 has no NAWM in some editions), Ch34=176–179, Ch35=180–182, Ch36=183–185, Ch37=186–188, Ch38=189–191, Ch39=192–194.

## Two-column panel layout (critical coordinates)

These must be followed exactly. Title ends at y=1.78 and content starts at y=1.70 — the 0.08" nominal overlap is deliberate (the 14pt title only fills the top of its 0.4" box). Deviation causes visible overlap or bottom truncation.

```javascript
// Panel background
s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.30, w: 4.6, h: 4.1, fill: { color: C.panel } });
// Panel title
s.addText("■ 小節標題", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
// Panel content (paraSpaceAfter MUST be 0)
s.addText("• bullet\n• bullet", { x: 0.5, y: 1.70, w: 4.35, h: 3.65, fontSize: 14, color: C.ivory, fontFace: "Calibri", valign: "top", paraSpaceAfter: 0 });
```

Right panel mirrors at x: 5.1 (background) / 5.25 (title) / 5.3 (content).

**Special-case layouts** — two recurring exceptions that must NOT be converted to standard two-column:
- **Reduced-panel + bottom bar** (e.g. Sonata Types slides): panel `h: 3.10`, content `h: 2.55`, bottom bar at `y: 4.52`. Keeps a summary bar visible below the panels.
- **3-card format** (e.g. Concerto Types slides): three horizontally-arranged cards spanning the full slide width. Keep as-is.

**C-phase row→column recipe** — when converting a 4-bullet horizontal-row slide to two-column: split bullets 2+2 into left/right panels, name each panel by its combined theme (e.g. 「神學起點 / 政治與印刷」, 「定義與來源 / 合集與影響」). Hand-wrap content with `\n` to keep each panel at ~12-14 visible lines (6-7 per bullet group) at 14pt within the 4.35w × 3.65h container.

**Known bug pattern** — earlier chapters used `y: 1.58` + `paraSpaceAfter: 2`, which produced panel-title/content overlap. Sed fix:

```bash
sed -i '' 's/y: 1.58, w: 4.35, h: 3.65/y: 1.70, w: 4.35, h: 3.65/g' chXX_*.js
sed -i '' 's/paraSpaceAfter: 2/paraSpaceAfter: 0/g' chXX_*.js
```

Before applying blindly: chapters whose content was sized for the old coordinates will overflow at the bottom (e.g. Ch06 p20). Always rebuild and render-verify after.

Other recurring overflow cause: blank-line spacers (`\n\n`) consume the same vertical space as a content line — main culprit at 17+ effective lines. Remove spacers before shrinking text.

**`\n\nyoutu.be/` anti-pattern** — a blank line before the YouTube link in a panel creates a visible gap between content and link, making the link look orphaned. Always use `\nyoutu.be/...` (single newline) at the end of panel content.

## Layout verification rule (non-negotiable)

From memory `feedback_layout_first.md`: never ship slides with overlapping or truncated text. Spot-checking 4-5 pages is insufficient — the Ch09 p16 overflow was missed that way. **Render every page and visually verify before claiming a chapter is done.** Every row-based or stacked-panel layout has its own overflow risk even if the standard two-column template renders clean.

## Font size floor

Minimum 14pt anywhere on a slide. Don't shrink text to fix overflow — trim content instead (combine bullets, shorten phrases). Textbook-level detail is fine to cut; the slide is a talking point, not a transcript.

## Unicode glyph caveat

U+30FB (Katakana middle dot ・) does not render in some fonts via LibreOffice → boxes appear. Use U+00B7 (·) instead.

## Chapter format status

All 39 chapters (Ch01–Ch39) are complete and verified page-by-page as of 2026-04-19. All conform to the Ch26 two-column panel format.

## Git conventions

`.gitignore` excludes `*.pptx`, `*.js` (via `ch*.js`), `node_modules/`, `package*.json`, and `*.jpg` except the textbook cover. **Only commit PDFs and README.md.** Commit message pattern: `ChNN: <short description>` (see `git log`).

## Color palettes

Each chapter defines its own `C` object at the top of its JS file, matching the period (e.g. Ch26 = forest green + gold for early-Romantic orchestral). Helpers `darkSlide()`, `lightSlide()`, `topBar()`, `bottomBar()`, `header()` are defined per-file — intentionally duplicated rather than shared, so each chapter can tweak visuals independently.
