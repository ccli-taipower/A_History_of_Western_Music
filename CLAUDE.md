# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project purpose

Slide generator for Burkholder/Grout/Palisca *A History of Western Music* (10th ed.), 39 chapters, bilingual (繁中 + English), 16:9. Each chapter lives in `chXX_name.js`, builds `ChXX_Name.pptx` via [pptxgenjs](https://github.com/gitbrent/PptxGenJS); LibreOffice converts to `ChXX_Name.pdf`. Only PDFs (and README.md) are committed — the repo is published on GitHub for students to download.

## Prerequisites

`package.json` / `node_modules/` are gitignored, so a fresh clone needs:

```bash
npm install pptxgenjs       # only runtime dep (pptxgenjs ^4.0.1)
# LibreOffice must be installed system-wide and reachable as `soffice`
# macOS: brew install --cask libreoffice
```

## Build pipeline

```bash
# 0. (once) mkdir -p /tmp/verify   — pdftoppm writes to the prefix but does NOT create the dir
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

**NAWM sequence by chapter (verified from textbook PDF NAWM Guide pp.14–22):** Ch01=1–2 (Seikilos 1, Euripides Orestes 2), Ch02 references NAWM 3 for notation examples, Ch03=3–7 (Christmas Mass 3a–3k, Office 4a–4b, Tropes 5a–5b, Sequences 6a–6b, Hildegard Ordo virtutum 7), Ch04=8–13 (Bernart 8, Comtessa de Dia 9, Adam 10, Walther 11, Cantiga 12, Estampie 13), Ch05=14–23 (Musica enchiriadis organum 14a–c, Aquitanian 15–16, Léonin 17, clausula 18, Pérotin 19, motets 20–21, conductus 22, Sumer 23), Ch06=24–31 (Vitry 24, Machaut Messe 25a–b, Machaut secular 26–27, Caserta 28, Jacopo 29, Landini 30–31), Ch07=no assigned NAWM (survey chapter; listening examples borrowed from Ch08–11), Ch08=32–37 (carol A newë work 32, Dunstable 33, Binchois 34, Du Fay 35–37a/b), Ch09=38–45 (Busnoys 38, Ockeghem 39, Isaac 40–41, Josquin 42–45), Ch10=46–57 (Encina 46, Arcadelt 47, Rore 48, Marenzio 49, Gesualdo 50, Sermisy 51, Janequin 52, Lassus 53, Le Jeune 54, Morley/Weelkes/Gastoldi 55–56, Dowland 57), Ch11=58–65 (Luther/Walter 58a–d, Geneva Psalter 59a–b, Tallis 60, Byrd Sing joyfully 61, Gombert 62, Palestrina Pope Marcellus 63a–b, Victoria O magnum 64a–d, Lassus 65), Ch12=66–70 (Susato 66a–c, Holborne 67a–b, Narvaez 68a–b, Byrd Variations 69, G.Gabrieli Canzon septimi toni 70), Ch13=71 (Monteverdi Cruda Amarilli 71), Ch14=72–76 (Caccini Vedrò 72, Peri Euridice 73a–b, L'Orfeo 74a–e, Poppea 75a–d, Cavalli Artemisia 76a–c), Ch15=77–84 (Strozzi Lagrime mie 77, G.Gabrieli In ecclesiis 78, Grandi 79, Carissimi Jephte 80a–b, Schütz Saul 81, Frescobaldi Toccata 82, Frescobaldi Ricercare 83, Marini Sonata IV 84), Ch16=85–92 (Lully Armide 85a–c, Lully Te Deum 86, Charpentier 87a–d, Gaultier 88, Jacquet de la Guerre Suite 89a–h, Purcell Dido 90a–c, Torrejón La púrpura 91a–b, Padilla 92), Ch17=93–97 (Sartorio Giulio Cesare 93a–e, Scarlatti Clori vezzosa 94a–b, Scarlatti La Griselda 95, Corelli Trio Sonata 96a–d, Buxtehude Praeludium 97), Ch18=98–100 (Vivaldi Op.3/6 98a–c, Couperin 25e ordre 99a–c, Rameau Hippolyte 100a–b), Ch19=101–108 (Telemann Paris Quartet 101a–c, Bach BWV 543 102a–b, Bach Durch Adams Fall 103, Bach WTC I No.8 104a–b, Bach Cantata BWV 62 105a–f, Bach St Matthew Passion 106a–e, Handel Giulio Cesare 107a–c, Handel Saul 108a–c), Ch20=116 (Galuppi Sonata D Major Op.2/1, esp. 116c pp.465–469). Ch31=169a/b, Ch32=170–179 (Debussy 170, Schoenberg 171, Stravinsky 172, Ravel 173, Falla 174, Holst 175, Rachmaninoff 176, Scriabin 177, Sibelius 178, Satie 179), Ch33=180–189 (Schoenberg 180–181, Berg 182, Webern 183, Stravinsky 184–185, Bartók 186–187, Ives 188–189), Ch34=190–193 (Gershwin 190, Smith 191, Armstrong 192, Ellington 193), Ch35=194–204 (Milhaud 194, Weill 195, Hindemith 196, Prokofiev 197, Shostakovich 198, Villa-Lobos 199, Varèse 200, Cowell 201, Crawford 202, Copland 203, Still 204), Ch36=205–208 (Parker 205, Davis 206, Coltrane 207, Bernstein 208), Ch37=209–216 (Britten 209, Messiaen 210, Boulez 211, Cage 212–213, Varèse elec. 214, Babbitt 215, Penderecki 216), Ch38=217–223 (Bright Sheng 217, Reich 218, Adams 219, Ligeti 220, Gubaidulina 221, Schnittke 222, Pärt 223), Ch39=224–226 (Adès 224, Muhly 225, Widmann 226).

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

**Outline slide step rule** — single-column outlines (Ch01–Ch05 style) use `startY=1.08, rowH=0.28`; this is safe for up to 16 items (last item y=5.28; the 13pt text sits in the upper portion of the 0.28" box and stays clear of the bottom bar at y=5.5). Two-column outlines (Ch26 style, right column): `y: 1.25 + i * step`; step=0.3 is safe for ≤14 items; with 16 items step must be ≤0.26 (last item y=5.41). Ch38 outline (16 right-column items) was fixed this way.

## Layout verification rule (non-negotiable)

From memory `feedback_layout_first.md`: never ship slides with overlapping or truncated text. Spot-checking 4-5 pages is insufficient — the Ch09 p16 overflow was missed that way. **Render every page and visually verify before claiming a chapter is done.** Every row-based or stacked-panel layout has its own overflow risk even if the standard two-column template renders clean.

## Content audit rules

When verifying slides against the textbook, use three layers: (1) structure — compare textbook section headings against slide outline; (2) score/musical analysis — identify textbook Examples not reflected in slides; (3) facts — spot-check dates, names, works against textbook text.

**Content rule:** Slides may extend textbook content (YouTube links, cross-references to other chapters, bilingual labels) but must not fabricate content not in the textbook. Common violations found:
- Attributing influence to 20c composers for medieval styles (e.g. "影響：Messiaen · Ligeti" for Ars Subtilior)
- Naming specific modern ensembles (e.g. "Gothic Voices") when textbook only describes the general debate
- Placeholder YouTube URLs (e.g. `youtu.be/search`) — must be replaced with verified links

**Known format deviation — Key Terms slides in some chapters (Ch07, Ch09):** Use `y:1.0` for panel background and `y:1.05` for panel title instead of the standard `y:1.30` / `y:1.38`. This pushes the panel higher to fit more terms. When editing these chapters, change to standard coordinates and check that content still fits at `h:3.65` before pushing.

**Cover page range accuracy:** Always verify the chapter's printed page range from CLAUDE.md's chapter page table before writing `Textbook pp. XX–YY` on the cover slide. Wrong ranges found: Ch10 said 202–227 (correct: 205–228).

## Font size floor

Minimum 14pt anywhere on a slide. Don't shrink text to fix overflow — trim content instead (combine bullets, shorten phrases). Textbook-level detail is fine to cut; the slide is a talking point, not a transcript.

## Unicode glyph caveat

U+30FB (Katakana middle dot ・) does not render in some fonts via LibreOffice → boxes appear. Use U+00B7 (·) instead.

## Chapter format status

All 39 chapters (Ch01–Ch39) are complete and verified page-by-page. All conform to the Ch26 two-column panel format. Total: **933 slides** (as of 2026-04-28).

**Ch01–Ch20 content audit (2026-04-28):** Systematic comparison of every chapter against textbook pages (using official NAWM Guide and PDF extraction). All NAWM numbers verified, fabricated content removed, missing NAWM slides added. Key fixes per group:

- **Ch01–Ch05 (2026-04-27):** Added missing NAWM analysis slides, score-analysis content, conclusion slides. Ch03: Ordo virtutum corrected to "earliest NON-LITURGICAL music drama" (not simply earliest drama — Quem queritis liturgical dramas predate it).
- **Ch06–Ch10 (2026-04-27/28):** Fixed fabricated content (Gothic Voices, Messiaen/Ligeti for Ars Subtilior). Ch10: added Castiglione *Book of the Courtier* (1528) reference to amateur music market slide.
- **Ch11–Ch15 (2026-04-28):** NAWM numbers were systematically wrong (off by 1–5). Fixed all: Gombert NAWM 62, Holborne NAWM 67, Strozzi NAWM 77, G.Gabrieli In ecclesiis NAWM 78, Frescobaldi Ricercare NAWM 83, etc. Removed fabricated NAWMs (Franco Salve regina, Zefiro torna as NAWM, Brandenburg No.5 as NAWM 106). Fixed all cover page ranges.
- **Ch16–Ch20 (2026-04-28):** Ch16: added 5 missing NAWM slides (86–88, 91–92). Ch17: NAWM 97 corrected Torelli→Buxtehude. Ch19: NAWM 104/105 swapped, NAWM 106 corrected to St. Matthew Passion. Ch20: completely rebuilt to align with textbook pp.454–470 only (removed Ch21/Ch22 content; added NAWM 116 Galuppi as the chapter's primary textbook example, textbook Example 20.1 pp.465–469).

**Outline slide pairing rule (critical — from 2026-04-28 fixes):** For chapters with N slides, the two-column TOC array must have exactly N items interleaved as `[left_item], [right_item], ...`. For odd N: left column = slides 1–⌈N/2⌉, right column = slides ⌈N/2⌉+1–N, last pair left-only. Violation causes "15 appearing in bottom-right" or numbered-but-empty entries. Always verify the rendered outline page after editing slide count.

**Earlier history:** Ch33 added 2026-04-21; Ch34–Ch38 renumbered and corrected 2026-04-21 (NAWM numbers, Ch36 content rewritten, Ch35 recreated); Ch32 expanded 28→36 slides 2026-04-21.

**YouTube links status** — All NAWM YouTube links are filled; no `youtu.be/PENDING` placeholders remain anywhere.

## Textbook reference PDF

`A HISTORY of WESTERN TENTН MUSIC.pdf` in the project root is the complete 1,121-page electronic copy of Burkholder 10e (gitignored by filename). Use it whenever slide content needs exact page citations. Extract the TOC with:

```bash
pdftotext -layout -f 6 -l 22 "A HISTORY of WESTERN TENTН MUSIC.pdf" - | grep -E "^\s+[0-9]{1,2}\s+[A-Z]"
```

**PDF page offset:** printed page N = PDF page (N + 37). Front matter occupies PDF pages 1–37; Chapter 1 starts at PDF page 41 (printed page 4). To extract a specific printed page range, use `-f (N+37) -l (M+37)`.

Chapter page ranges (printed-book pagination, derived from the TOC) used by `qa_100.js`: Ch01 4-19 · Ch02 20-41 · Ch03 42-62 · Ch04 63-79 · Ch05 80-105 · Ch06 106-132 · Ch07 136-158 · Ch08 159-179 · Ch09 180-204 · Ch10 205-228 · Ch11 229-253 · Ch12 254-277 · Ch13 278-296 · Ch14 297-316 · Ch15 317-338 · Ch16 339-370 · Ch17 371-401 · Ch18 402-423 · Ch19 424-453 · Ch20 454-470 · Ch21 471-493 · Ch22 494-513 · Ch23 514-553 · Ch24 554-579 · Ch25 580-617 · Ch26 618-645 · Ch27 646-670 · Ch28 671-710 · Ch29 711-730 · Ch30 731-755 · Ch31 756-769 · Ch32 770-803 · Ch33 804-847 · Ch34 848-868 · Ch35 869-897 · Ch36 898-918 · Ch37 919-953 · Ch38 954-989 · Ch39 990-1020.

## Condensed Review (`condensed_review.js`)

100-slide end-of-semester review covering all 39 chapters in 9 historical periods. Only `Condensed_Review.pdf` is committed — the source JS stays local.

**Critical layout note:** `condensed_review.js` uses `pres.layout = 'LAYOUT_16x9'` (10" × 5.625"). Do NOT change it to `LAYOUT_WIDE` — that was a bug that caused panels to occupy only the upper-left 73% of the slide. The coordinates (panel `x:0.3`, `w:4.6`; right panel `x:5.1`; bottom bar `y:5.50`) are calibrated for LAYOUT_16x9.

Structure: 9 periods × 11 slides (Period 7 has 12). Each period: cover → 時代背景 → 核心作曲家 → 代表體裁 → 風格特徵 → NAWM×2 → 術語 → 時間軸 → 考試重點 → 跨期比較/NAWM3.

Build: same pipeline as chapters — `node condensed_review.js` → `soffice --headless --convert-to pdf Condensed_Review.pptx`.

## A3 Cheat Sheet (`cheat_sheet.js`)

One-page landscape poster (420×297 mm, A3) that synthesizes the whole book: 6 color-coded composer tracks on a 900–2025 timeline, 9 era bands, genre evolution strip, and 9 period summary cards. Built the same way as chapters but with a custom layout:

```javascript
pres.defineLayout({ name:'A3_LAND', width:16.54, height:11.69 });
pres.layout = 'A3_LAND';
```

Per-country sub-row allocation uses greedy interval scheduling. Dense tracks (German, Italy, France) have 3–4 sub-rows; at font sizes 6–6.5pt, A3 print is still legible.

**Gitignore note:** `cheat_sheet.js` is ignored by `.gitignore` (all `*.js` excluded) — only `Cheat_Sheet.pdf` is committed.

## A3 Top 100 Q&A (`qa_100.js`)

One-page A3 landscape Q&A sheet (`QA_100.pdf`). Same custom A3 layout as Cheat Sheet:

```javascript
pres.defineLayout({ name:'A3_LAND', width:16.54, height:11.69 });
pres.layout = 'A3_LAND';
```

Structure: 3×3 grid of period blocks (matching the 9-period Condensed Review division, Ch01-39), each block has a colored header band + 11-12 Q&A split into two sub-columns. Q&A items carry mixed type tags `[申]` essay / `[名]` term / `[短]` short-answer, bilingual (Chinese answer with English technical terms preserved), and exact Burkholder 10e page references (e.g. `[pp. 192-203, 234-243]`) extracted from the textbook reference PDF above.

Typography: 6.5pt question + 6pt answer + 5.5pt italic page-ref, all Calibri, tight line-spacing. A `totalQA !== 100` sanity check aborts the build if counts drift.

Build: same pipeline as chapters — `node qa_100.js` → `soffice --headless --convert-to pdf QA_100.pptx`.

## A3 Lineage Map (`lineage_map.js`)

One-page A3 landscape poster (`Lineage_Map.pdf`) visualizing ~105 composers on a **1400–2025 timeline × 6 nationality tracks** with teacher-student, influence, rivalry, and romantic relationship lines.

**Coordinate system:**
- X axis: `X0=0.65, X_END=16.3, YEAR0=1400, YEAR_END=2025`; `function xp(year)` maps year → x-inch
- Y axis: 6 tracks (`TRACK_Y=[0.82, 2.38, 3.94, 5.50, 7.06, 8.62]`) for France/Italy/German-Austrian/England/Russia/America-Other; sub-rows at `SUB=0.30"` offsets; `function yp(track, sub)`
- `dy` field on a composer entry adds a fine y-offset applied during pre-compute: `c.cy = yp(c.track, c.sub) + (c.dy || 0)`

**Node tiers:** `TIER_R={1:0.095, 2:0.065, 3:0.045}` (core / NAWM-important / general)

**Label positioning** — mutually exclusive fields on each composer entry:
- `lp:'ul'|'ur'|'ll'|'lr'` — diagonal corner; `labelRight:true` / `labelLeft:true` — horizontal beside circle; `labelAbove:true/false` — above/below centered (default: sub<0→above)
- `leaderLine:true` — draws a thin grey line from label center to circle center before rendering the node, making ambiguous labels in dense clusters unambiguous

**Relationship line types:** `teacher` (gold solid →), `influence` (blue solid →), `cross` (blue dashed →), `rival` (red dashed ↔), `couple` (pink solid), `adore` (pink dashed). Drawn via `dline(x1,y1,x2,y2,lineOpts)` helper (function declaration, hoisted — safe to call before its textual position).

**Build:**
```bash
node lineage_map.js
soffice --headless --convert-to pdf Lineage_Map.pptx
pdftoppm -r 250 -png Lineage_Map.pdf /tmp/verify/p   # visual check
```

Only `Lineage_Map.pdf` is committed; `lineage_map.js` is gitignored.

## Git conventions

`.gitignore` excludes `*.pptx`, `*.js` (all generators: chapter files, `cheat_sheet.js`, `condensed_review.js`, `qa_100.js`, etc.), `node_modules/`, `package*.json`, and `*.jpg` except the textbook cover. **Only PDFs and README.md are committed — source JS stays local.** Commit message prefixes: `ChNN:` (single chapter edit), `ChNN-MM:` (multi-chapter batch), `Add:` (new artifacts), `Condensed_Review:`, `Lineage_Map:`, `README:`.

## Color palettes

Each chapter defines its own `C` object at the top of its JS file, matching the period (e.g. Ch26 = forest green + gold for early-Romantic orchestral). Helpers `darkSlide()`, `lightSlide()`, `topBar()`, `bottomBar()`, `header()` are defined per-file — intentionally duplicated rather than shared, so each chapter can tweak visuals independently.
