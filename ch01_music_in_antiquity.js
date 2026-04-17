const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 1: Music in Antiquity";
pres.author = "A History of Western Music, 10th ed.";

// ── Color palette (ancient/parchment theme) ──────────────────────────────────
const C = {
  darkBg:   "2C1810",  // deep espresso brown (dark slides bg)
  gold:     "C8A020",  // amber gold (accent)
  cream:    "FBF5E6",  // warm parchment (light slide bg)
  wine:     "7A2830",  // deep wine red (section banners)
  rust:     "A84030",  // terracotta rust (highlights)
  darkText: "2C1810",  // body text on light bg
  lightText:"FBF5E6",  // text on dark bg
  midBrown: "5C3A28",  // subheadings on light bg
  sand:     "E8D8A8",  // soft gold for dividers
  slate:    "4A3828",  // dark text secondary
};

// ── Master background helper ──────────────────────────────────────────────────
function darkSlide(pres) {
  const s = pres.addSlide();
  s.background = { color: C.darkBg };
  return s;
}
function lightSlide(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };
  return s;
}

// ── Reusable: gold accent bar across top ─────────────────────────────────────
function topBar(s, color) {
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.12, fill: { color: color || C.gold } });
}
function bottomBar(s, color) {
  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.5, w: "100%", h: 0.125, fill: { color: color || C.gold } });
}

// ── SLIDE 1 · Title ──────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  // Gold top/bottom bars
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0,    w: "100%", h: 0.15, fill: { color: C.gold } });
  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.47, w: "100%", h: 0.155, fill: { color: C.gold } });

  // Chapter label
  s.addText("A HISTORY OF WESTERN MUSIC · TENTH EDITION", {
    x: 0.5, y: 0.45, w: 9, h: 0.35,
    fontSize: 14, color: C.sand, charSpacing: 3, align: "center", fontFace: "Georgia",
  });

  // Chapter number
  s.addText("CHAPTER 1", {
    x: 0.5, y: 0.9, w: 9, h: 0.55,
    fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });

  // Main title
  s.addText("MUSIC IN ANTIQUITY", {
    x: 0.5, y: 1.5, w: 9, h: 1.3,
    fontSize: 52, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
  });

  // Decorative rule
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 2.9, w: 5, h: 0.04, fill: { color: C.gold } });

  // Subtitle / era
  s.addText("Prehistoric Civilizations · Mesopotamia · Ancient Greece · Rome", {
    x: 0.5, y: 3.05, w: 9, h: 0.4,
    fontSize: 16, color: C.sand, italic: true, align: "center", fontFace: "Georgia",
  });

  // Page range
  s.addText("Textbook pp. 4–19", {
    x: 0.5, y: 4.8, w: 9, h: 0.3,
    fontSize: 14, color: C.gold, align: "center", fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("本章概覽 Chapter Overview", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6,
    fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const sections = [
    ["■", "最早的音樂 The Earliest Music", "史前時期 40,000 BCE 起的骨笛與壁畫"],
    ["■", "古美索不達米亞 Ancient Mesopotamia", "蘇美文明、第一位作曲家、最早的記譜法"],
    ["■", "古希臘 Ancient Greece", "樂器、哲學、倫理學說、音樂理論"],
    ["■", "古羅馬 Ancient Rome", "從希臘繼承的音樂傳統與羅馬的發展"],
    ["■", "希臘遺產 The Greek Heritage", "西方音樂思想的奠基石"],
  ];

  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.9;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y: y, w: 0.6, h: 0.65, fill: { color: C.wine }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.55, fontSize: 22, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y: y, w: 8.4, h: 0.33, fontSize: 15, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.31, w: 8.4, h: 0.28, fontSize: 14, color: C.midBrown, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 3 · How Do We Know? (4 types of evidence) ─────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("如何研究古代音樂？How We Study Music of the Past", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.6,
    fontSize: 22, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addText("音樂是聲音，聲音本質上是短暫的。我們依靠四種歷史痕跡 (historical traces) 重建過去的音樂。", {
    x: 0.4, y: 0.78, w: 9.2, h: 0.45,
    fontSize: 14, color: C.slate, fontFace: "Calibri", italic: true, valign: "top",
  });

  const cards = [
    ["①", "實物遺存\nPhysical Remains", "樂器、演奏場所\n(e.g. 骨笛、豎琴、劇場)"],
    ["②", "視覺圖像\nVisual Images", "陶器、壁畫上的\n演奏場景描繪"],
    ["③", "文字記述\nWritten Sources", "關於音樂與音樂家\n的哲學與理論著作"],
    ["④", "音樂本身\nActual Music", "以記譜法保存或口傳，\n1870年代後有錄音"],
  ];

  cards.forEach(([num, title, desc], i) => {
    const x = 0.4 + i * 2.35;
    s.addShape(pres.ShapeType.rect, { x, y: 1.3, w: 2.15, h: 3.8, fill: { color: C.wine }, rounding: true });
    s.addText(num, { x, y: 1.4, w: 2.15, h: 0.7, fontSize: 30, bold: true, color: C.gold, align: "center", fontFace: "Georgia" });
    s.addText(title, { x: x + 0.1, y: 2.15, w: 1.95, h: 0.85, fontSize: 14, bold: true, color: C.cream, align: "center", fontFace: "Georgia" });
    s.addShape(pres.ShapeType.rect, { x: x + 0.3, y: 3.05, w: 1.55, h: 0.03, fill: { color: C.gold } });
    s.addText(desc, { x: x + 0.1, y: 3.15, w: 1.95, h: 1.7, fontSize: 14, color: C.sand, align: "center", fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 4 · The Earliest Music (Prehistoric) ───────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("最早的音樂 The Earliest Music", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.8, w: 9.2, h: 0.03, fill: { color: C.sand } });

  s.addText("史前時代 Prehistoric Evidence", {
    x: 0.4, y: 0.9, w: 4.5, h: 0.4, fontSize: 16, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });

  const bullets = [
    "石器時代於動物骨骼、猛獁象牙鑽孔製笛；最古老骨笛約 40,000 BCE（德國 Hohle Fels）\n Stone Age bone/ivory flutes; oldest ~40,000 BCE at Hohle Fels, Germany",
    "洞穴壁畫疑似描繪樂器演奏；土耳其新石器壁畫（BC 6000）顯示鼓手為舞者伴奏\n Cave paintings depict instruments; Neolithic Turkey murals show drummers",
    "青銅時代（BC 4000 起）：金屬樂器（鐘、鈸、角）和撥弦樂器出現\n Bronze Age (4th mill. BCE): metal instruments and plucked strings emerge",
    "沒有文字記錄 → 史前音樂幾乎完全沉默\n No written record → prehistoric music is nearly silent",
  ];

  s.addText(bullets.map((b, i) => ({
    text: b, options: { bullet: true, breakLine: i < bullets.length - 1, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 4 }, valign: "top",
  })), { x: 0.4, y: 1.35, w: 5.7, h: 4.0 });

  // Right column — key image caption box
  s.addShape(pres.ShapeType.rect, { x: 6.3, y: 0.9, w: 3.3, h: 4.55, fill: { color: C.wine }, rounding: true });
  s.addText("■ Hohle Fels 骨笛", { x: 6.4, y: 1.0, w: 3.1, h: 0.45, fontSize: 15, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText([
    { text: "• 距今約 40,000–42,000 年\n", options: { bold: true } },
    { text: "• 格里芬禿鷲橈骨製成\n" },
    { text: "• 五個指孔，目前出土最完整的早期笛子\n\n" },
    { text: "40,000–42,000 years old\nMade from griffon vulture wing bone\n5 finger holes — most complete early flute found", options: { italic: true } },
  ], { x: 6.4, y: 1.55, w: 3.1, h: 2.5, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" });

  s.addText("■ 關鍵觀念 Key Insight", { x: 6.4, y: 3.85, w: 3.1, h: 0.3, fontSize: 14, bold: true, color: C.gold, align: "center" });
  s.addText("文字的發明標誌史前時代的結束，\n也是音樂歷史的真正起點。\nWriting marks the true beginning\nof music history.", {
    x: 6.4, y: 4.18, w: 3.1, h: 1.05, fontSize: 14, color: C.sand, fontFace: "Calibri", align: "center", valign: "top",
  });
}

// ── SLIDE 5 · Mesopotamia Overview ───────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("古美索不達米亞的音樂", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("Music in Ancient Mesopotamia", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 18, color: C.sand, fontFace: "Georgia", align: "center", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const facts = [
    ["地理位置", "底格里斯河與幼發拉底河之間（今伊拉克、敘利亞）\nBetween Tigris & Euphrates Rivers (modern Iraq & Syria)"],
    ["文明", "蘇美人（Sumerians）建立最早城市；楔形文字（Cuneiform）\nSumerians built first true cities; developed cuneiform writing"],
    ["音樂用途", "婚禮歌、哀歌、軍樂、工作歌、酒館音樂、神聖儀式\nWedding songs, laments, military, work, tavern, sacred ceremonies"],
    ["文獻", "ca. 2500 BCE 起的詞匯表：包含樂器、音律、演奏技術\nWord lists from ~2500 BCE: instruments, tuning, performing techniques"],
    ["理論", "巴比倫人使用七聲全音階；七種調式（對應鋼琴白鍵）\nBabylonians used 7-note diatonic scales; 7 modes (piano white keys)"],
  ];

  facts.forEach(([label, content], i) => {
    const y = 1.2 + i * 0.85;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.75, fill: { color: "3A2015" }, rounding: true });
    s.addText(label, { x: 0.45, y: y + 0.07, w: 2.0, h: 0.55, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(content, { x: 2.4, y: y + 0.05, w: 7.1, h: 0.62, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 6 · Mesopotamian Instruments ───────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("美索不達米亞的樂器", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addText("Mesopotamian Instruments · ca. 2500 BCE, Royal Tombs at Ur", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 14, color: C.midBrown, fontFace: "Calibri", italic: true, margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.02, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const instrs = [
    ["豎琴 Lyre", "弦線與音板平行；弦繞在可調音的棒子上\nStrings run parallel to soundboard; tunable crossbar"],
    ["竪琴 Harp", "弦線垂直於音板；琴頸直接連接共鳴箱\nStrings perpendicular to soundboard; neck attached to soundbox"],
    ["公牛里拉 Bull Lyre", "蘇美獨特風格；琴箱飾以公牛頭（宗教意涵）\nDistinctively Sumerian; soundbox features a bull's head (religious)"],
    ["魯特琴 Lute", "長頸撥弦樂器\nLong-necked plucked string instrument"],
    ["管樂 Pipes & Drums", "雙管笛、鼓、鈸、拍板、搖鈴\nDouble pipes, drums, cymbals, clappers, rattles, bells"],
  ];

  // Two-column top: Lyre & Harp on left, Lute & Pipes on right
  const topLeft  = instrs.slice(0, 2);   // Lyre, Harp
  const topRight = instrs.slice(3);      // Lute, Pipes & Drums
  const featured = instrs[2];            // Bull Lyre — full-width feature row

  topLeft.forEach(([name, desc], i) => {
    const y = 1.1 + i * 1.35;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 4.5, h: 1.2, fill: { color: C.wine }, rounding: true });
    s.addText(name, { x: 0.55, y: y + 0.08, w: 4.2, h: 0.4, fontSize: 14.5, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(desc, { x: 0.55, y: y + 0.5, w: 4.2, h: 0.65, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0 });
  });
  topRight.forEach(([name, desc], i) => {
    const y = 1.1 + i * 1.35;
    s.addShape(pres.ShapeType.rect, { x: 5.1, y, w: 4.5, h: 1.2, fill: { color: C.rust }, rounding: true });
    s.addText(name, { x: 5.25, y: y + 0.08, w: 4.2, h: 0.4, fontSize: 14.5, bold: true, color: C.cream, fontFace: "Georgia", margin: 0 });
    s.addText(desc, { x: 5.25, y: y + 0.5, w: 4.2, h: 0.65, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0 });
  });

  // Bull Lyre — featured full-width card at the bottom
  {
    const [bullName, bullDesc] = featured;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y: 3.78, w: 9.2, h: 1.08, fill: { color: C.midBrown }, rounding: true });
    s.addText(bullName, { x: 0.6, y: 3.86, w: 3.2, h: 0.38, fontSize: 14.5, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(bullDesc, { x: 0.6, y: 4.26, w: 9.0, h: 0.52, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0 });
  }

  s.addText("烏爾出土的豎琴（ca. 2500 BCE）是現存最早的撥弦樂器實物之一 · The lyres and harps from Ur are among the earliest surviving plucked string instruments", {
    x: 0.4, y: 5.0, w: 9.2, h: 0.48, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 7 · First Named Composer: Enheduanna ───────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("第一位有名字的作曲家", {
    x: 0.5, y: 0.2, w: 9, h: 0.52, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("The First Composer Known by Name", {
    x: 0.5, y: 0.72, w: 9, h: 0.35, fontSize: 16, color: C.sand, italic: true, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.1, w: 7, h: 0.04, fill: { color: C.gold } });

  // Big name
  s.addText("Enheduanna", {
    x: 0.5, y: 1.25, w: 9, h: 0.85, fontSize: 48, bold: true, color: C.lightText, fontFace: "Georgia", align: "center",
  });
  s.addText("恩赫杜安納", {
    x: 0.5, y: 2.1, w: 9, h: 0.5, fontSize: 22, color: C.gold, fontFace: "Georgia", align: "center",
  });

  const details = [
    ["⏱ 年代", "活躍於 ca. 2300 BCE"],
    ["■‍■ 身份", "阿卡德族（Akkadian）月神南納（Nanna）的女大祭司，任職於烏爾"],
    ["■ 作品", "為月神南納與女神伊南娜（Inanna）創作讚美詩（hymns）"],
    ["■ 現存", "歌詞刻於楔形文字泥板上保存——但音樂本身未傳下來"],
    ["■ 意義", "比古埃及、希臘的已知作曲家早了一千多年"],
  ];

  details.forEach(([label, text], i) => {
    const y = 2.7 + i * 0.55;
    s.addText(label + "  ", { x: 0.6, y, w: 1.4, h: 0.45, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", align: "right", margin: 0 });
    s.addText(text, { x: 2.1, y, w: 7.5, h: 0.45, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 8 · Oldest Notation: Hurrian Hymn ──────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("最古老的音樂記譜", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 26, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addText("The Oldest Known Musical Notation", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.32, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.02, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Left: tablet info
  s.addText("■ 胡利安頌歌 Hurrian Hymn (H.6)", {
    x: 0.4, y: 1.1, w: 5.4, h: 0.45, fontSize: 16, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  const hLines = [
    "• 年代 Date: ca. 1400–1250 BCE",
    "• 出土地 Found: Ugarit（敘利亞沿岸商業城邦）",
    "• 載體 Medium: 楔形文字泥板（Clay tablet）",
    "• 內容 Content: 獻給月神妻子尼卡爾（Nikkal）的讚美詩",
    "• 現存最古老、近乎完整的記譜實例",
    "  Oldest nearly complete piece with musical notation",
    "• 記譜法至今仍難以完全破解",
    "  Notation too poorly understood to read with certainty",
    "• 當時音樂家可能不靠記譜演奏，而用它如食譜重建旋律",
    "  Used as recipe to reconstruct melody, not for live reading",
  ];
  s.addText(hLines.map((t, i) => ({
    text: t, options: { breakLine: i < hLines.length - 1, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2 }, valign: "top",
  })), { x: 0.4, y: 1.6, w: 5.5, h: 3.6 });

  // Right: Seikilos info
  s.addShape(pres.ShapeType.rect, { x: 6.0, y: 1.1, w: 3.6, h: 4.2, fill: { color: C.wine }, rounding: true });
  s.addText("■ Seikilos 墓誌銘\nEpitaph of Seikilos", {
    x: 6.1, y: 1.2, w: 3.4, h: 0.75, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText([
    { text: "• 年代：ca. 1st century CE\n", options: { bold: true } },
    { text: "• 現存最古老、完整的音樂作品\n" },
    { text: "  Oldest complete surviving composition\n\n" },
    { text: "歌詞（原文）\n", options: { bold: true, color: "C8A020" } },
    { text: "Ὅσον ζῇς, φαίνου…\n\n", options: { italic: true } },
    { text: "中文大意：\n", options: { bold: true } },
    { text: "「趁著活著，閃耀吧；別讓任何事令你憂愁」\n\nWhile you live, shine; let nothing grieve you." },
  ], { x: 6.1, y: 1.98, w: 3.4, h: 3.2, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top" });
}

// ── SLIDE 9 · Timeline ────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("歷史時間軸 Timeline", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });

  const events = [
    ["ca. 3500–3000 BCE", "蘇美城市興起 Rise of Sumerian cities"],
    ["ca. 3100 BCE",      "楔形文字建立 Cuneiform writing established"],
    ["ca. 2500 BCE",      "烏爾皇家陵墓出土豎琴 Royal Tombs at Ur — lyres & harps"],
    ["ca. 2300 BCE",      "恩赫杜安納作曲 Enheduanna composes hymns"],
    ["ca. 1800 BCE",      "巴比倫音樂理論著作 Babylonian writings about music"],
    ["ca. 1400–1250 BCE", "最古老近乎完整的樂譜 Oldest nearly complete notation (Ugarit)"],
    ["ca. 800 BCE",       "希臘城邦興起 Rise of Greek city-states; Homer"],
    ["586 BCE",           "薩卡達斯贏得皮提亞競賽 Sakadas wins Pythic Games"],
    ["ca. 408 BCE",       "歐里庇得斯《奧瑞斯提斯》附有音樂 Euripides' Orestes (with music)"],
    ["ca. 330 BCE",       "亞里斯多塞諾斯音樂理論 Aristoxenus, Harmonic Elements"],
    ["146 BCE",           "希臘成為羅馬行省 Greece becomes province of Rome"],
    ["ca. 1st cent. CE",  "Seikilos 墓誌銘 Epitaph of Seikilos (oldest complete composition)"],
  ];

  // Timeline line
  s.addShape(pres.ShapeType.rect, { x: 2.55, y: 0.88, w: 0.05, h: 4.5, fill: { color: C.gold } });

  events.forEach(([date, event], i) => {
    const y = 0.88 + i * 0.37;
    // Dot
    s.addShape(pres.ShapeType.ellipse, { x: 2.42, y: y + 0.04, w: 0.28, h: 0.28, fill: { color: C.gold } });
    // Date (left)
    s.addText(date, { x: 0.1, y, w: 2.25, h: 0.33, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", align: "right", margin: 0 });
    // Event (right)
    s.addText(event, { x: 2.9, y, w: 6.8, h: 0.33, fontSize: 14, color: C.lightText, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 10 · Music in Ancient Greece — Overview ────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("古希臘的音樂 Music in Ancient Greece", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.58, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 0.8, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addText("古希臘是現存資料最豐富的古代音樂文明，保存了：\n40 餘件音樂作品或片段 · 大量文字理論著作 · 數百件陶器圖像 · 部分實物樂器", {
    x: 0.4, y: 0.88, w: 9.2, h: 0.7, fontSize: 14, color: C.sand, italic: true, fontFace: "Calibri", align: "center", valign: "top",
  });

  const cols = [
    {
      title: "音樂與宗教\nMusic & Religion",
      items: [
        "起源被神話化：阿波羅、俄耳甫斯",
        "Origins mythologized (Apollo, Orpheus)",
        "遍及希臘生活：軍事、學校、宗教、劇場",
        "Pervaded all of Greek life",
      ],
    },
    {
      title: "音樂與詩歌\nMusic & Poetry",
      items: [
        "「melos」= 旋律＋歌詞＋舞蹈",
        "Melos: melody + text + dance",
        "主要是單音音樂 monophonic",
        "「lyric poetry」= 配里拉琴的詩",
        "Lyric poetry = sung to the lyre",
      ],
    },
    {
      title: "音樂與數字\nMusic & Numbers",
      items: [
        "Pythagoras：8度=2:1, 5度=3:2",
        "Octave 2:1, Fifth 3:2, Fourth 4:3",
        "harmonia = 秩序、數學比例、音階",
        "Mathematical order of universe",
        "「音樂之球」天文學類比",
        "Music of the Spheres (Plato)",
      ],
    },
  ];

  cols.forEach(({ title, items }, i) => {
    const x = 0.3 + i * 3.25;
    s.addShape(pres.ShapeType.rect, { x, y: 1.65, w: 3.05, h: 3.65, fill: { color: "3A2015" }, rounding: true });
    s.addText(title, { x: x + 0.1, y: 1.75, w: 2.85, h: 0.55, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
    s.addShape(pres.ShapeType.rect, { x: x + 0.2, y: 2.32, w: 2.65, h: 0.03, fill: { color: C.gold } });
    s.addText(items.map((t, j) => ({
      text: t, options: { bullet: true, breakLine: j < items.length - 1, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 3 }, valign: "top",
    })), { x: x + 0.1, y: 2.38, w: 2.85, h: 2.85 });
  });
}

// ── SLIDE 11 · Greek Instruments ─────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("希臘樂器 Greek Musical Instruments", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.72, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const instruments = [
    {
      name: "Aulos (αὐλός)",
      zh: "雙管笛",
      god: "▶ 酒神戴奧尼索斯 Dionysus",
      desc: [
        "雙管、各有指孔與蘆葦管（reed）",
        "Double pipe with finger holes and reed mouthpiece",
        "常見於飲酒聚會（symposium）和戲劇",
        "Used at symposia, Dionysian festivals, theater",
        "推測二管演奏同音或相距特定音程",
        "Two pipes likely played in unison or intervals",
      ],
    },
    {
      name: "Lyre (λύρα)",
      zh: "里拉琴",
      god: "▶ 太陽神阿波羅 Apollo",
      desc: [
        "七弦，龜殼共鳴箱，用撥片彈奏",
        "7 strings, tortoise-shell soundbox, played with plectrum",
        "雅典教育的核心科目",
        "Core element of Athenian education",
        "伴奏舞蹈、歌唱、史詩朗誦",
        "Accompanied dancing, singing, epic recitation",
      ],
    },
    {
      name: "Kithara (κιθάρα)",
      zh: "基薩拉",
      god: "▶ 用於公眾儀式與比賽",
      desc: [
        "大型里拉琴，通常站立演奏",
        "Larger lyre, played standing",
        "用於遊行、儀式和劇場",
        "Processions, ceremonies, theater",
        "kitharode = 邊彈邊唱的歌手",
        "Singer accompanying himself",
      ],
    },
  ];

  instruments.forEach(({ name, zh, god, desc }, i) => {
    const x = 0.35 + i * 3.2;
    s.addShape(pres.ShapeType.rect, { x, y: 0.82, w: 3.0, h: 4.4, fill: { color: i === 1 ? C.wine : C.rust }, rounding: true });
    s.addText(name, { x: x + 0.1, y: 0.92, w: 2.8, h: 0.45, fontSize: 15, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
    s.addText(zh, { x: x + 0.1, y: 1.37, w: 2.8, h: 0.35, fontSize: 14, color: C.cream, fontFace: "Georgia", align: "center" });
    s.addText(god, { x: x + 0.1, y: 1.72, w: 2.8, h: 0.32, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", align: "center", italic: true });
    s.addShape(pres.ShapeType.rect, { x: x + 0.25, y: 2.07, w: 2.5, h: 0.03, fill: { color: C.gold } });
    s.addText(desc.map((t, j) => ({
      text: t, options: { bullet: true, breakLine: j < desc.length - 1, fontSize: 14, color: C.cream, fontFace: "Calibri", paraSpaceAfter: 2 }, valign: "top",
    })), { x: x + 0.1, y: 2.13, w: 2.8, h: 3.05 });
  });

  s.addText("希臘人主要靠耳朵學習音樂，而非讀譜 · Greeks learned by ear, not notation", {
    x: 0.4, y: 5.25, w: 9.2, h: 0.22, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Calibri", margin: 0, align: "center", valign: "top",
  });
}

// ── SLIDE 12 · IN PERFORMANCE: Competitions ──────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  // "In Performance" label
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.2, w: 3.2, h: 0.45, fill: { color: C.gold }, rounding: true });
  s.addText("IN PERFORMANCE", { x: 0.4, y: 0.2, w: 3.2, h: 0.45, fontSize: 14, bold: true, color: C.darkBg, align: "center", fontFace: "Georgia" });

  s.addText("音樂競賽與職業音樂家\nCompetitions and Professional Musicians", {
    x: 0.4, y: 0.75, w: 9.2, h: 0.75, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.5, w: 9.2, h: 0.04, fill: { color: C.gold } });

  const items = [
    ["■ 競賽起源", "西元前 6 世紀起，阿夫洛斯和基薩拉成為獨奏樂器，開始舉辦競賽\nFrom 6th cent. BCE, aulos & kithara played as solos; competitions held"],
    ["■ Sakadas of Argos", "586、582、578 BCE 三度贏得皮提亞競賽（Pythian Games）阿夫洛斯獨奏首獎\nWon 3 times; performed the Pythic Nomos (Apollo vs. Python)"],
    ["■ 名演奏家", "巡迴演出積累財富；觀眾可達數千人；部分女性演奏家亦成名\nFamous virtuosos toured and attracted thousands; some women gained fame"],
    ["■ 社會地位", "競賽外的職業演奏家社會地位低，許多是奴隸\nOutside competitions, most professional performers were of low status, often slaves"],
    ["■ 獎品", "圖 1.9 的陶罐（裝酒或油）即為競賽獎品\nThe amphora in Fig. 1.9 was awarded as a competition prize"],
  ];

  items.forEach(([label, text], i) => {
    const y = 1.6 + i * 0.77;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.67, fill: { color: "3A2015" }, rounding: true });
    s.addText(label, { x: 0.45, y: y + 0.07, w: 2.2, h: 0.5, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(text, { x: 2.7, y: y + 0.05, w: 6.9, h: 0.55, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 13 · Music and Ethos ────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("音樂與倫理 Music and Ethos", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 26, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addText("希臘哲學家相信音樂能直接影響人的性格與道德\nGreek philosophers believed music could affect ethical character", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.52, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Calibri", margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.24, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Two philosopher boxes
  const philos = [
    {
      name: "柏拉圖 Plato",
      years: "ca. 429–347 BCE",
      book: "《理想國》Republic",
      color: C.wine,
      points: [
        "音樂塑造靈魂；只有特定音樂適合教育",
        "Music molds the soul; only certain music suits education",
        "贊成 Dorian 和 Phrygian 調式（促進節制與勇氣）",
        "Endorsed Dorian & Phrygian harmoniai (temperance & courage)",
        "禁止複雜或混合類型的音樂",
        "Banned complex or mixed musical genres",
        "■ 「音樂法律不可改變」→ 藝術無法度導致社會無法紀",
        "Musical conventions must not change: lawlessness in art → anarchy",
      ],
    },
    {
      name: "亞里斯多德 Aristotle",
      years: "384–322 BCE",
      book: "《政治學》Politics",
      color: C.rust,
      points: [
        "音樂模仿情感，激發聽者的 ethos",
        "Music imitates and arouses ethos",
        "比柏拉圖更開放：音樂亦可用於娛樂",
        "More open: music can serve enjoyment",
        "負面情緒可藉音樂宣洩（catharsis）",
        "Negative emotions purged — catharsis",
        "Mixolydian 悲傷、Dorian 平靜、Phrygian 激昂",
        "Mode effects shape character and emotion",
      ],
    },
  ];

  philos.forEach(({ name, years, book, color, points }, i) => {
    const x = 0.35 + i * 4.8;
    s.addShape(pres.ShapeType.rect, { x, y: 1.32, w: 4.5, h: 4.05, fill: { color }, rounding: true });
    s.addText(name, { x: x + 0.1, y: 1.42, w: 4.3, h: 0.45, fontSize: 18, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
    s.addText(`${years} · ${book}`, { x: x + 0.1, y: 1.87, w: 4.3, h: 0.32, fontSize: 14, color: C.sand, align: "center", italic: true, fontFace: "Calibri", valign: "top" });
    s.addShape(pres.ShapeType.rect, { x: x + 0.3, y: 2.22, w: 3.9, h: 0.03, fill: { color: C.gold } });
    s.addText(points.map((t, j) => ({
      text: t, options: { bullet: true, breakLine: j < points.length - 1, fontSize: 14, color: C.cream, fontFace: "Calibri", paraSpaceAfter: 2 }, valign: "top",
    })), { x: x + 0.1, y: 2.28, w: 4.3, h: 3.02 });
  });
}

// ── SLIDE 14 · Greek Music Theory ─────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("希臘音樂理論 Greek Music Theory", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.72, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const concepts = [
    {
      term: "Tetrachord 四音列",
      def: "四個音，跨越完全四度；是希臘音樂體系的基本單位",
      eng: "4 notes spanning a perfect fourth — building block of Greek scales",
      extra: "三種類型：全音階（diatonic）、半音階（chromatic）、異名同音（enharmonic）",
    },
    {
      term: "Greater Perfect System 大完整體系",
      def: "四個四音列組合成兩個八度的音域體系",
      eng: "4 tetrachords combined to span two octaves",
      extra: "各音和四音列均有專用名稱（如 mese、hypaton、diezeugmenon）",
    },
    {
      term: "Species of Octave 八度音種",
      def: "七種全音/半音排列方式 → 後來影響中世紀調式",
      eng: "7 arrangements of T & S within an octave → influenced medieval modes",
      extra: "Mixolydian, Lydian, Phrygian, Dorian, Hypolydian, Hypophrygian, Hypodorian",
    },
    {
      term: "Tonoi 調性",
      def: "Aristoxenus 等人以移調方式定義的 15 種音高範圍",
      eng: "15 scale positions defined by transposition (Aristoxenus et al.)",
      extra: "較高的 tonoi 被認為充滿活力，較低的則平靜",
    },
  ];

  concepts.forEach(({ term, def, eng, extra }, i) => {
    const y = 0.82 + i * 1.17;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 9.2, h: 1.08, fill: { color: i % 2 === 0 ? "EDE0C4" : "F5ECD7" }, rounding: true });
    s.addText(term, { x: 0.6, y: y + 0.07, w: 3.5, h: 0.38, fontSize: 14.5, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
    s.addText(def, { x: 0.6, y: y + 0.45, w: 4.4, h: 0.32, fontSize: 14, color: C.darkText, fontFace: "Calibri", valign: "top", margin: 0 });
    s.addShape(pres.ShapeType.rect, { x: 5.1, y: y + 0.08, w: 0.03, h: 0.88, fill: { color: C.sand } });
    s.addText(eng, { x: 5.2, y: y + 0.07, w: 4.2, h: 0.38, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Calibri", valign: "top", margin: 0 });
    s.addText(extra, { x: 5.2, y: y + 0.5, w: 4.2, h: 0.47, fontSize: 14, color: C.slate, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 15 · Surviving Greek Music: Seikilos ────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("現存的古希臘音樂 Surviving Ancient Greek Music", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 0.73, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addText([
    { text: "現存約 45 件古希臘音樂作品或片段", options: { bold: true } },
    { text: "，時間跨度從西元前 5 世紀到西元 4 世紀。\n最早的兩件來自歐里庇得斯（Euripides）劇作中的合唱。" },
  ], { x: 0.4, y: 0.82, w: 9.2, h: 0.55, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" });

  // Seikilos feature
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.42, w: 9.4, h: 3.8, fill: { color: "3A2015" }, rounding: true });
  s.addText("Seikilos 墓誌銘 · Epitaph of Seikilos", {
    x: 0.5, y: 1.55, w: 9, h: 0.5, fontSize: 20, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("現存最古老完整的音樂作品 · ca. 1st century CE · 出土於土耳其特拉勒斯（Tralles）", {
    x: 0.5, y: 2.05, w: 9, h: 0.35, fontSize: 14, color: C.sand, align: "center", italic: true, fontFace: "Calibri", valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 2.42, w: 7, h: 0.03, fill: { color: C.gold } });

  // Greek text + translation
  s.addText([
    { text: "「Ὅσον ζῇς, φαίνου·\n", options: { italic: true, fontSize: 16, color: C.gold } },
    { text: "μηδὲν ὅλως σύ λυποῦ·\n", options: { italic: true, fontSize: 16, color: C.gold } },
    { text: "πρὸς ὀλίγον ἐστὶ τὸ ζῆν,\n", options: { italic: true, fontSize: 16, color: C.gold } },
    { text: "τὸ τέλος ὁ χρόνος ἀπαιτεῖ.」\n\n", options: { italic: true, fontSize: 16, color: C.gold } },
    { text: "趁著活著，閃耀吧；莫讓任何事令你憂愁。\n生命短暫，時間終會把它帶走。\n\n", options: { fontSize: 14, color: C.cream } },
    { text: "While you live, shine; let nothing grieve you.\nLife is short, and time claims its toll.", options: { fontSize: 14, color: C.sand, italic: true } },
  ], { x: 1.0, y: 2.5, w: 8, h: 2.6, fontFace: "Georgia", align: "center" });
}

// ── SLIDE 16 · Music in Ancient Rome ─────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("古羅馬的音樂 Music in Ancient Rome", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 26, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.72, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Two columns
  const leftItems = [
    ["■ 功能 Function", "婚禮、葬禮、公共競技場（鬥劍士）、宗教儀式、公共演出\nWeddings, funerals, arena (gladiatorial games), religion, public performance"],
    ["■ 起源 Origin", "146 BCE 希臘成為羅馬行省後，大量吸收希臘音樂文化\nAfter 146 BCE Greek conquest → massive Greek musical influence on Rome"],
    ["■ Hydraulis 水風琴", "在競技場大受歡迎；用水壓控制的管風琴\nPopular in arenas; early organ operated by water pressure"],
    ["■ Tuba", "長約 1.3 公尺的直管銅號；軍隊與公共儀式\nLong straight bronze trumpet ~1.3m; military & ceremonies"],
  ];

  leftItems.forEach(([label, content], i) => {
    const y = 0.82 + i * 1.17;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 5.6, h: 1.08, fill: { color: i % 2 === 0 ? "EDE0C4" : "F5ECD7" }, rounding: true });
    s.addText(label, { x: 0.55, y: y + 0.07, w: 5.3, h: 0.35, fontSize: 14, bold: true, color: C.rust, fontFace: "Georgia", margin: 0 });
    s.addText(content, { x: 0.55, y: y + 0.43, w: 5.3, h: 0.58, fontSize: 14, color: C.darkText, fontFace: "Calibri", valign: "top", margin: 0 });
  });

  // Right: instrument list
  s.addShape(pres.ShapeType.rect, { x: 6.2, y: 0.82, w: 3.4, h: 4.4, fill: { color: C.rust }, rounding: true });
  s.addText("羅馬樂器\nRoman Instruments", { x: 6.3, y: 0.92, w: 3.2, h: 0.6, fontSize: 15, bold: true, color: C.cream, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 6.5, y: 1.55, w: 2.8, h: 0.03, fill: { color: C.gold } });
  const romInstr = [
    "Tibia（希臘 Aulos 的羅馬版）",
    "Lyra / Cithara（弦樂）",
    "Tuba（直管銅號）",
    "Cornu（C 形號角）",
    "Hydraulis（水風琴）",
    "Sistrum（叉鈴，宗教用）",
    "鼓、鈸、鈴 Drums, Cymbals, Bells",
  ];
  s.addText(romInstr.map((t, i) => ({
    text: t, options: { bullet: true, breakLine: i < romInstr.length - 1, fontSize: 14, color: C.cream, fontFace: "Calibri", paraSpaceAfter: 3 }, valign: "top",
  })), { x: 6.3, y: 1.6, w: 3.2, h: 3.55 });
}

// ── SLIDE 17 · The Greek Heritage ─────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("希臘的遺產 The Greek Heritage", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("為西方音樂奠定的基石 — Foundations of Western Music Thought", {
    x: 0.4, y: 0.73, w: 9.2, h: 0.35, fontSize: 15, color: C.sand, italic: true, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.1, w: 7, h: 0.04, fill: { color: C.gold } });

  const legacies = [
    ["音階與調式", "七聲音階與調式體系，直接影響中世紀教會調式和現代大小調\n7-note scales and modes → medieval church modes → modern major/minor"],
    ["音樂理論", "音程、音階、節奏的定義；仍是今日音樂理論的基礎\nInterval, scale, rhythm definitions — still foundational today"],
    ["音樂倫理學", "音樂影響性格的觀念，從柏拉圖到現代流行音樂爭議都有回響\nMusic affects character — echoed in every era's debates (jazz, rock, rap...)"],
    ["數學與音響學", "Pythagoras 的諧音比例（2:1, 3:2, 4:3）是聲學和調音的基礎\nPythagorean ratios underlie acoustics and tuning theory"],
    ["理論傳承", "希臘理論著作傳入中世紀歐洲，成為大學音樂課程的核心\nGreek treatises passed to medieval Europe; formed core of university curriculum"],
  ];

  legacies.forEach(([label, text], i) => {
    const y = 1.2 + i * 0.83;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.73, fill: { color: "3A2015" }, rounding: true });
    s.addText(label, { x: 0.45, y: y + 0.08, w: 2.1, h: 0.55, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(text, { x: 2.6, y: y + 0.05, w: 7.0, h: 0.62, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 18 · Summary ─────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.15, fill: { color: C.gold } });
  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.47, w: "100%", h: 0.155, fill: { color: C.gold } });

  s.addText("本章重點回顧 Chapter Summary", {
    x: 0.5, y: 0.25, w: 9, h: 0.52, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 0.8, w: 7, h: 0.04, fill: { color: C.gold } });

  const summaryPoints = [
    "研究古代音樂需靠四種歷史痕跡：實物、圖像、文字、音樂本身\nStudy of ancient music relies on 4 types of evidence: physical remains, images, writings, music",
    "史前時代已有樂器（骨笛 ~40,000 BCE）；文字發明後音樂歷史才真正開始\nPrehistoric instruments existed (40,000 BCE); writing marks the true beginning of music history",
    "美索不達米亞：七聲全音階、最早作曲家（恩赫杜安納）、最古老記譜（胡利安頌歌）\nMesopotamia: diatonic scales, first composer (Enheduanna), oldest notation (Hurrian Hymn)",
    "希臘：音樂與宗教、詩歌、教育高度結合；柏拉圖與亞里斯多德論倫理學說\nGreece: music intertwined with religion, poetry, education; Plato & Aristotle on ethos",
    "希臘音樂理論（四音列、調式、音階）奠定了後世西方音樂理論的基礎\nGreek theory (tetrachord, modes, scales) founded all subsequent Western music theory",
    "古羅馬大量吸收希臘音樂文化；古典遺產透過文字傳入中世紀歐洲\nRome absorbed Greek music; classical heritage transmitted to medieval Europe through texts",
  ];

  s.addText(summaryPoints.map((t, i) => ({
    text: t, options: { bullet: true, breakLine: i < summaryPoints.length - 1, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 5 }, valign: "top",
  })), { x: 0.5, y: 0.92, w: 9, h: 4.4 });
}

// ── SLIDE 19 · 延伸閱讀 Further Reading ──────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("延伸閱讀與補充教材", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 24, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addText("Further Reading & Supplementary Resources", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.02, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const resources = [
    {
      cat: "■ 聆聽 YouTube",
      items: [
        "Seikilos Epitaph（最古老完整歌曲，里拉琴版）: youtube.com/watch?v=8Vkcolt-nmU",
        "Hurrian Hymn H.6（胡利安頌歌重建版，Peter Pringle）: youtube.com/watch?v=w8tfBLvlN98",
        "古希臘音樂重建 · Euripides Orestes合唱（牛津大學）: youtube.com/watch?v=4hOK7bU0S1Y",
        "古羅馬音樂重建 Synaulia（弦樂卷）: youtube.com/watch?v=HlFpiNAOdUo",
        "Michael Levy 古希臘里拉琴演奏系列: youtube.com/channel/UCJ1X6F7lGMEadnNETSzTv8A",
      ],
    },
    {
      cat: "■ 閱讀 Read",
      items: [
        "維基百科 Music of ancient Greece: en.wikipedia.org/wiki/Music_of_ancient_Greece",
        "Stanford Encyclopedia: History of Western Philosophy of Music: plato.stanford.edu/entries/hist-westphilmusic-to-1800",
        "World History Encyclopedia — Greek Music: worldhistory.org/Greek_Music",
        "Music of Mesopotamia (Wikipedia): en.wikipedia.org/wiki/Music_of_Mesopotamia",
      ],
    },
    {
      cat: "■ 學術 Academic",
      items: [
        "牛津大學：古希臘奧瑞斯提斯音樂重建計畫 (Oxford): classics.ox.ac.uk/recreating-music-ancient-greek-chorus-euripides-orestes",
        "胡利安頌歌與楔形文字樂譜研究 (Open Culture): openculture.com/2025/04/hear-the-worlds-oldest-known-song",
        "Oldest Songs in the World — Oldest.org: oldest.org/music/songs",
      ],
    },
  ];

  resources.forEach(({ cat, items }, i) => {
    const y = 1.1 + i * 1.4;
    s.addText(cat, { x: 0.4, y, w: 9.2, h: 0.32, fontSize: 14.5, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
    s.addText(items.map((t, j) => ({
      text: t, options: { bullet: true, breakLine: j < items.length - 1, fontSize: 14, color: C.slate, fontFace: "Calibri", paraSpaceAfter: 1 }, valign: "top",
    })), { x: 0.6, y: y + 0.32, w: 9.0, h: 1.05 });
  });
}

// ── Write ─────────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "/Users/ccli/Downloads/music_history_claude/Ch01_Music_in_Antiquity.pptx" })
  .then(() => console.log("■ Ch01_Music_in_Antiquity.pptx 已產生"))
  .catch(e => console.error("■", e));
