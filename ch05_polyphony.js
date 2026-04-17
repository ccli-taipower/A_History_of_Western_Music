const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 5: Polyphony through the Thirteenth Century";
pres.author = "A History of Western Music, 10th ed.";

// ── Color palette (same parchment/ancient theme as Ch01–Ch04) ────────────────
const C = {
  darkBg:   "2C1810",
  gold:     "C8A020",
  cream:    "FBF5E6",
  wine:     "7A2830",
  rust:     "A84030",
  darkText: "2C1810",
  lightText:"FBF5E6",
  midBrown: "5C3A28",
  sand:     "E8D8A8",
  slate:    "4A3828",
};

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
function topBar(s, color) {
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.12, fill: { color: color || C.gold } });
}
function bottomBar(s, color) {
  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.5, w: "100%", h: 0.125, fill: { color: color || C.gold } });
}

// ── SLIDE 1 · Title ──────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0,    w: "100%", h: 0.15, fill: { color: C.gold } });
  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.47, w: "100%", h: 0.155, fill: { color: C.gold } });

  s.addText("A HISTORY OF WESTERN MUSIC · TENTH EDITION", {
    x: 0.5, y: 0.45, w: 9, h: 0.35,
    fontSize: 14, color: C.sand, charSpacing: 3, align: "center", fontFace: "Georgia",
  });

  s.addText("CHAPTER 5", {
    x: 0.5, y: 0.9, w: 9, h: 0.55,
    fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });

  s.addText("POLYPHONY THROUGH\nTHE THIRTEENTH CENTURY", {
    x: 0.3, y: 1.4, w: 9.4, h: 2.0,
    fontSize: 38, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
    paraSpaceAfter: 0,
  });

  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.55, w: 5, h: 0.04, fill: { color: C.gold } });

  s.addText("Organum · Aquitanian · Notre Dame · Motet · Franconian Notation", {
    x: 0.4, y: 3.7, w: 9.2, h: 0.4,
    fontSize: 14, color: C.sand, italic: true, align: "center", fontFace: "Georgia",
  });

  s.addText("Textbook pp. 80–104", {
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
    ["■", "西方音樂的四項特徵 Four Traits of Western Music", "對位、和聲、記譜的核心地位、作曲與演奏的分離"],
    ["■", "早期奧爾加農 Early Organum", "Musica enchiriadis、Winchester Troper、Guido of Arezzo、Ad organum faciendum"],
    ["■", "亞奎丹複音 Aquitanian Polyphony", "聖馬夏修院與聖地牙哥朝聖路線；discant vs florid organum"],
    ["■", "巴黎聖母院學派 Notre Dame School", "Léonin · Pérotin · Magnus liber organi · 6 個節奏模式"],
    ["■", "經文歌 Motet", "Clausula → Latin motet → French/secular motet；Franconian 記譜"],
    ["■", "英格蘭複音 English Polyphony", "喜用三度與六度；Rondellus · Rota · Sumer is icumen in"],
  ];

  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.wine }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 14, color: C.midBrown, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 3 · Four Distinguishing Concepts ───────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("西方音樂的四項特徵", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("Four Concepts That Distinguish Western Music from Most Other Traditions", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const concepts = [
    ["①", "對位 Counterpoint", "多條獨立旋律線同時結合——旋律間的對話\nTwo or more melodic lines combined so each retains its own shape"],
    ["②", "和聲 Harmony", "垂直聲響的組合、進行與張力解決\nThe vertical dimension: chord progressions, tension, resolution"],
    ["③", "記譜的核心地位 Centrality of Notation", "不僅記錄音樂，還是作曲工具與保存方式\nNotation is both a tool for composition and a means of preservation"],
    ["④", "作曲與演奏分離 Composition vs. Performance", "作曲家成為獨立角色——作品先於演出而存在\nThe composer emerges as a distinct role; the work exists prior to performance"],
  ];

  concepts.forEach(([num, name, desc], i) => {
    const y = 1.3 + i * 0.95;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.85, fill: { color: "3A2015" }, rounding: true });
    s.addText(num, { x: 0.4, y: y + 0.15, w: 0.7, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center", margin: 0 });
    s.addShape(pres.ShapeType.rect, { x: 1.1, y: y + 0.12, w: 0.03, h: 0.6, fill: { color: C.gold } });
    s.addText(name, { x: 1.2, y: y + 0.08, w: 8.4, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(desc, { x: 1.2, y: y + 0.4, w: 8.4, h: 0.45, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 });
  });

  s.addText("這些特徵在 9–13 世紀的複音音樂中成形，成為西方音樂傳統的根基", {
    x: 0.4, y: 5.2, w: 9.2, h: 0.25, fontSize: 14, color: C.gold, fontFace: "Calibri", align: "center", valign: "top",
  });
}

// ── SLIDE 4 · Early Organum ──────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("早期奧爾加農 Early Organum", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 24, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addText("9th–11th Centuries · The First Written Polyphony of the West", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 14, color: C.midBrown, fontFace: "Calibri", margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.04, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Parallel organum
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.2, w: 4.6, h: 2.0, fill: { color: C.wine }, rounding: true });
  s.addText("① 平行奧爾加農 Parallel Organum", {
    x: 0.45, y: 1.26, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("Musica enchiriadis（ca. 850–890 抄本）\n• 主要聲部唱聖詠\n• 附加聲部（organal voice）平行五度或四度\n• 四度平行因三全音（tritone）問題而較複雜\n  → 產生混合平行與斜行的修正版本", {
    x: 0.5, y: 1.76, w: 4.2, h: 1.39, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top",
  });

  // Mixed/oblique
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.2, w: 4.6, h: 2.0, fill: { color: C.midBrown }, rounding: true });
  s.addText("② 混合平行與斜行 Mixed Parallel/Oblique", {
    x: 5.25, y: 1.26, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("Guido of Arezzo — Micrologus (ca. 1025–28)\n• 為避免三全音，聲部有時保持同音（斜行）\n• 兩聲部得以同音開始、結束\n• Winchester Troper（英格蘭，ca. 992–996）\n  收錄 174 首由 Wulfstan 創作的奧爾加農", {
    x: 5.3, y: 1.76, w: 4.2, h: 1.39, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top",
  });

  // Note-against-note
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.3, w: 4.6, h: 2.0, fill: { color: C.rust }, rounding: true });
  s.addText("③ 音對音奧爾加農 Note-Against-Note", {
    x: 0.45, y: 3.36, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("• 11 世紀後期；附加聲部置於聖詠之上\n• 使用逆向（contrary）與斜行（oblique）\n• 聲部常交叉——更自由的橫向線條\n• 已使用三度作為協和音程\n• 由 organal voice 變為 upper voice", {
    x: 0.5, y: 3.82, w: 4.2, h: 1.44, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0,
  });

  // Ad organum faciendum
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 3.3, w: 4.6, h: 2.0, fill: { color: C.slate }, rounding: true });
  s.addText("■ Ad organum faciendum", {
    x: 5.25, y: 3.36, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("「如何作奧爾加農」· ca. 1100\n• 最重要的早期複音實用教本\n• 提供音對音複音具體範例\n• 複音仍被視為對聖詠的即興裝飾\n• 記譜僅供教學；實際多為即興口傳", {
    x: 5.3, y: 3.82, w: 4.2, h: 1.44, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0,
  });
}

// ── SLIDE 5 · Aquitanian Polyphony ───────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("亞奎丹複音 Aquitanian Polyphony", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("12th Century · St. Martial at Limoges · Santiago de Compostela", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Context
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 1.25, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ 來源與背景 Sources & Context", {
    x: 0.45, y: 1.36, w: 9.1, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("• 南法 St. Martial 修道院（Limoges）· 12 世紀手稿保存複音 versus 與聖詠裝飾\n• Codex Calixtinus（ca. 1173）· 聖地牙哥德孔波斯特拉朝聖聖殿手稿 · 含 21 首複音\n• 作品多為兩聲部 · 兩聲部皆使用拉丁文 · 多數為節慶禮儀用 · 或拉丁讚美詩 versus", {
    x: 0.5, y: 1.86, w: 9.0, h: 0.69, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top",
  });

  // Discant style
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 2.7, w: 4.6, h: 2.55, fill: { color: C.wine }, rounding: true });
  s.addText("■ Discant 對唱風格", {
    x: 0.45, y: 2.78, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.55, y: 3.13, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("• 兩聲部以「音對音」方式行進\n  Note-against-note counterpoint\n• 速度相近——兩聲部有明確節奏\n• 常用逆向與斜行\n• 常見於以音節式聖詠為基底之處\n• 英文 \"discant\" 源自拉丁 discantus\n  = 「分開地唱」", {
    x: 0.5, y: 3.28, w: 4.2, h: 1.89, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top",
  });

  // Florid style
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 2.7, w: 4.6, h: 2.55, fill: { color: C.rust }, rounding: true });
  s.addText("■ Florid 華麗奧爾加農", {
    x: 5.25, y: 2.78, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 5.35, y: 3.13, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("• 下方聲部（tenor）以持續長音唱聖詠\n  拉丁 tenere = 「持續、держать」→ tenor\n• 上方聲部在長音之上唱華麗裝飾花唱\n  Upper voice sings melismas over sustained tones\n• 節奏自由——無固定拍子\n• 聲音效果像「在聖詠上懸浮」\n• 開創了巴黎聖母院時期的奧爾加農風格", {
    x: 5.3, y: 3.28, w: 4.2, h: 1.89, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 6 · Notre Dame Cathedral & Polyphony ───────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("巴黎聖母院學派", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addText("Notre Dame Polyphony · Paris, Late 12th to Early 13th Century", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 14, color: C.midBrown, fontFace: "Calibri", margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.04, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Cathedral / context
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.18, w: 9.4, h: 2.0, fill: { color: C.wine }, rounding: true });
  s.addText("■ 聖母院與大學 Cathedral & University", {
    x: 0.45, y: 1.24, w: 9.1, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("• 建築：奠基 ca. 1160、祭壇 1182、中殿 ca. 1200、正立面 ca. 1250\n  Choir 1182, nave ca. 1200, façade ca. 1250\n• 與巴黎大學（ca. 1150）密不可分——哲學與神學的中心\n  Linked to the University of Paris — center of Scholastic thought\n• 音樂亦長、繁、結構化——仿教堂建築\n  Music paralleled the cathedral: long, intricate, structured", {
    x: 0.5, y: 1.70, w: 9.0, h: 1.45, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0,
  });

  // Transmission
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.25, w: 9.4, h: 2.0, fill: { color: C.rust }, rounding: true });
  s.addText("■ 傳播與口頭傳承 Transmission & Orality", {
    x: 0.45, y: 3.32, w: 9.1, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("• 曲目在西歐傳唱逾一世紀——從西班牙到蘇格蘭\n  Sung across Europe for over a century\n• 現存最早手稿來自 1240 年代，晚於曲目數十年\n  Earliest manuscripts from the 1240s\n• 同曲在不同手稿中文本差異——顯示口傳為主要媒介\n  Manuscript variations show oral transmission", {
    x: 0.5, y: 3.78, w: 9.0, h: 1.45, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0,
  });
}

// ── SLIDE 7 · The Six Rhythmic Modes ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("六種節奏模式 The Rhythmic Modes", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("Notre Dame's Great Innovation · First Notation for Duration Since Ancient Greece", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Context
  s.addText("Johannes de Garlandia (ca. 1260) 在 De mensurabili musica 歸納為六種「模式」，以連音符組合表示", {
    x: 0.4, y: 1.22, w: 9.2, h: 0.35, fontSize: 14, color: C.sand, fontFace: "Calibri", italic: true, align: "center", valign: "top", margin: 0,
  });

  // Modes table header
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.65, w: 9.2, h: 0.45, fill: { color: C.gold } });
  s.addText("模式 Mode", { x: 0.5, y: 1.68, w: 1.8, h: 0.4, fontSize: 14, bold: true, color: C.darkBg, fontFace: "Georgia", align: "center" });
  s.addText("模式 Pattern", { x: 2.4, y: 1.68, w: 3.2, h: 0.4, fontSize: 14, bold: true, color: C.darkBg, fontFace: "Georgia", align: "center" });
  s.addText("現代節奏等值 Modern Equivalent", { x: 5.7, y: 1.68, w: 3.8, h: 0.4, fontSize: 14, bold: true, color: C.darkBg, fontFace: "Georgia", align: "center" });

  const modes = [
    ["I",   "L B",    "■ ■   (長-短)",          "最古老最常用 · trochaic"],
    ["II",  "B L",    "■ ■   (短-長)",          "iambic 反向"],
    ["III", "L B B",  "■. ■ ■   (長-短-短)",    "dactylic"],
    ["IV",  "B B L",  "■ ■ ■.   (短-短-長)",    "anapestic · 極少用"],
    ["V",   "L L",    "■. ■.   (連續長音)",     "最古老 · tenor 常用"],
    ["VI",  "B B B",  "■ ■ ■   (連續短音)",     "tribrachic · 裝飾性"],
  ];

  modes.forEach(([num, pat, eq, note], i) => {
    const y = 2.15 + i * 0.48;
    const bg = i % 2 === 0 ? "3A2015" : "4A2820";
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 9.2, h: 0.46, fill: { color: bg } });
    s.addText(num, { x: 0.5, y: y + 0.05, w: 1.8, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
    s.addText(pat, { x: 2.4, y: y + 0.05, w: 1.5, h: 0.35, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", align: "center" });
    s.addText(eq, { x: 3.9, y: y + 0.05, w: 2.7, h: 0.35, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top" });
    s.addText(note, { x: 6.6, y: y + 0.05, w: 3.0, h: 0.35, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" });
  });

  s.addText("基本時間單位 tempus 永遠以三為一組——體現三位一體的完美數字「3」", {
    x: 0.4, y: 5.12, w: 9.2, h: 0.3, fontSize: 14, color: C.gold, fontFace: "Calibri", align: "center", valign: "top",
  });
}

// ── SLIDE 8 · Léonin & the Magnus liber organi ───────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("萊奧寧與《大奧爾加農書》", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 24, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addText("Léonin & the Magnus liber organi (ca. 1160s–ca. 1200)", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 14, color: C.midBrown, fontFace: "Calibri", margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.04, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Léonin
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.wine }, rounding: true });
  s.addText("■ Léonin (fl. 1150s–ca. 1201)", {
    x: 0.45, y: 1.28, w: 4.3, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.55, y: 1.66, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("• 1150 年代起任巴黎主教座堂神職\n• Notre Dame 參事、聖職者；與 St. Victor 修道院有關\n• 同時為詩人——曾以詩句改寫創世紀前八卷\n• Anonymous IV（ca. 1285 論文）稱其為\n  「excellent organista」（傑出的奧爾加農歌者/作曲家）\n• 據記載編纂了 Magnus liber organi（「大奧爾加農書」）\n\n■ 但並非獨力完成——如同聖母院的建築，這是\n  數代歌者的集體創作", {
    x: 0.5, y: 1.78, w: 4.2, h: 3.4, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top",
  });

  // Magnus liber
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.midBrown }, rounding: true });
  s.addText("■ Magnus liber organi", {
    x: 5.25, y: 1.28, w: 4.3, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 5.35, y: 1.66, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("■ 內容\n  教會年度主要節慶之\n  Graduals、Alleluias、Office Responsories\n  的獨唱段落的兩聲部複音設定\n\n■ 特點\n  • 在 Leoninus 的 organum duplum 中\n  • tenor 以聖詠為基、行緩慢長音\n  • upper voice（duplum）以華麗花唱盤旋\n\n■ 現存形式\n  • 「大書」本身已失傳\n  • 後世手稿（兩部於 Wolfenbüttel，\n    一部於 Florence）保存了曲目", {
    x: 5.3, y: 1.78, w: 4.2, h: 3.4, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 9 · Pérotin, Clausula, Conductus ───────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("佩羅丁、克勞蘇拉與導引歌", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("Pérotin · Substitute Clausula · Conductus", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Pérotin
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 1.55, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ Pérotin (fl. late 12th – early 13th c.)", {
    x: 0.45, y: 1.36, w: 9.1, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("• Anonymous IV 稱其為「Perotinus the Great」· 編訂 Magnus liber\n• 首創三聲部 organum triplum 與四聲部 organum quadruplum\n• 代表作：Viderunt omnes（1198）· Sederunt principes · Alleluia Nativitas（NAWM 19）\n• 所有上聲部使用節奏模式 · 常以 voice exchange 結構長段落", {
    x: 0.5, y: 1.86, w: 9.0, h: 0.95, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top",
  });

  // Clausula
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 2.95, w: 4.6, h: 2.3, fill: { color: C.wine }, rounding: true });
  s.addText("■ Substitute Clausula", {
    x: 0.45, y: 3.02, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("clausula = 「句子、子句」\n• 奧爾加農中的自足段落\n• 「替代式 clausula」可取代原曲相同段落\n  讓聖詠有多種處理選擇\n• 多為 discant 風格（兩聲部皆有節奏模式）\n• Florence 手稿收錄 10 種 Dominus 替代版本", {
    x: 0.5, y: 3.52, w: 4.2, h: 1.65, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top",
  });

  // Conductus
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 2.95, w: 4.6, h: 2.3, fill: { color: C.rust }, rounding: true });
  s.addText("■ Polyphonic Conductus", {
    x: 5.25, y: 3.02, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("Notre Dame 複音作曲家也創作 conductus：\n• tenor 為新作旋律——不借自聖詠\n• 各聲部同時歌唱相同歌詞\n• 歌詞多為音節式設定\n• 句首句尾常出現花唱尾聲 caudae\n• 主題為拉丁宗教或嚴肅題材詩歌", {
    x: 5.3, y: 3.52, w: 4.2, h: 1.65, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 10 · The Motet: Origins ────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("經文歌的誕生 The Birth of the Motet", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 24, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addText("Early 13th Century · From Clausula to New Genre", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 14, color: C.midBrown, fontFace: "Calibri", margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.04, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Origin
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.18, w: 9.4, h: 1.75, fill: { color: C.wine }, rounding: true });
  s.addText("■ 起源 Origin", {
    x: 0.45, y: 1.26, w: 9.1, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("• 13 世紀初，巴黎樂手為 discant clausula 上聲部加上拉丁新詞（如同 tropes 之於花唱）\n• 新體裁稱為「motet」：拉丁 motetus < 法文 mot 字\n• 音樂裝飾反過來被文字裝飾：decoration gets decorated in turn\n• 常以各聲部首字組成複合標題：Factum est salutare / Dominus", {
    x: 0.5, y: 1.76, w: 9.0, h: 1.16, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top",
  });

  // Development
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.05, w: 9.4, h: 2.15, fill: { color: C.rust }, rounding: true });
  s.addText("■ 演變 Development Over the 13th Century", {
    x: 0.45, y: 3.13, w: 9.1, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("• 初期：拉丁宗教文本以 clausula 為骨架，可在彌撒中演唱\n• 中期：加入三、四聲部 → double / triple motet（多歌詞）\n• 後期：法文、世俗主題 → 脫離教會，成為菁英的娛樂\n• tenor 由聖詠段落轉為純音樂骨幹 → Hieronymus de Moravia ca. 1270 稱「cantus firmus」\n• 13 世紀中葉 motet 取代 organum 與 conductus，成為主導體裁", {
    x: 0.5, y: 3.63, w: 9.0, h: 1.55, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 11 · Franconian Notation & Late 13th c. Motet ──────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("Franconian 記譜與後期經文歌", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("Franco of Cologne (ca. 1280) · A New Way of Writing Rhythm", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Franconian notation
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ Ars cantus mensurabilis", {
    x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("Franco of Cologne · ca. 1280\n「可度量音樂的藝術」\n\n■ 革命性突破\n用音符形狀本身表示相對時值\n\n四種單音符號：\n• Double long（倍長）· Long（長）\n• Breve（短）· Semibreve（半短）\n\n■ 基本單位 tempus 以三為一組\n  三個 tempora = 一個 perfection\n  （完美的三位一體結構）\n\n■ 此體系沿用數世紀——現代音符的祖先", {
    x: 0.5, y: 1.88, w: 4.2, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top",
  });

  // Late motet types
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ 後期經文歌風格", {
    x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("① Franconian Motet\n• 以 Franco 命名的新式經文歌\n• 各上聲部獨立節奏，不遵循節奏模式\n• 例：Adam de la Halle\n\n② Petronian Motet\n• Petrus de Cruce / Pierre de la Croix\n• 每 tempus 可容納多達 7 個 semibreves\n• 三聲部速度層級：tenor 慢+duplum 中+triplum 快\n• 例：Aucun ont trouvé (NAWM 22)\n\n③ Choirbook Format\n• 三聲部分寫於同頁或對頁\n• 1280 起至 16 世紀沿用", {
    x: 5.3, y: 1.88, w: 4.2, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 12 · English Polyphony ─────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("英格蘭複音 English Polyphony", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 24, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addText("A Distinctive Musical Dialect in 13th-Century England", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 14, color: C.midBrown, fontFace: "Calibri", margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.04, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Characteristics
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.18, w: 4.7, h: 4.15, fill: { color: C.wine }, rounding: true });
  s.addText("■ 英格蘭風格特徵", {
    x: 0.45, y: 1.26, w: 4.4, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.55, y: 1.62, w: 4.2, h: 0.02, fill: { color: C.gold } });
  s.addText("• 1066 諾曼征服後，英格蘭文化音樂與法國緊密連結\n  寫作 Notre Dame 各體裁\n• 偏好 conductus 的「同節奏」與規律樂句\n• 喜用不完美協和音程（三度、六度）\n  常以平行方式出現——受本土民間複音影響\n\n• Gerald of Wales（ca. 1200）記載\n  威爾斯與北不列顛的即興部分歌唱\n• Hymn to St. Magnus（12 世紀）\n  奧克尼群島守護聖人頌歌——連續平行三度\n\n■ 這些特色在 15 世紀影響歐陸作曲家\n  孕育出「國際風格 international style」", {
    x: 0.5, y: 1.76, w: 4.3, h: 3.43, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top",
  });

  // Rondellus / Rota / Sumer
  s.addShape(pres.ShapeType.rect, { x: 5.2, y: 1.18, w: 4.5, h: 4.15, fill: { color: C.midBrown }, rounding: true });
  s.addText("■ 聲部互換與循環", {
    x: 5.35, y: 1.26, w: 4.2, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 5.45, y: 1.62, w: 4.0, h: 0.02, fill: { color: C.gold } });
  s.addText("■ Rondellus 輪唱複音\n  三個樂句 a b c 由三聲部輪流交換：\n   V1: a b c · V2: c a b · V3: b c a\n  同一段三聲部樂句重複三次，角色互換\n\n■ Rota 卡農輪唱\n  同度永恆卡農（perpetual round）\n\n■ Sumer is icumen in（ca. 1250）\n  「夏日已來到」——最著名的 rota\n  • 四聲部同度卡農\n  • 下方兩聲部唱 pes（持續低音）\n    彼此形成兩聲部 rondellus\n  • 產生 F-A-C-F 與 G-B■-D 交替\n  • 英文歌詞讚美夏日與杜鵑鳥歌唱", {
    x: 5.4, y: 1.76, w: 4.2, h: 3.43, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 13 · Key Figures ───────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("關鍵人物 Key Figures", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 0.78, w: 7, h: 0.04, fill: { color: C.gold } });

  const figures = [
    ["■", "Guido of Arezzo", "ca. 991 – post-1033", "義大利本篤會修士 · Micrologus（ca. 1025–28）描述混合平行與斜行式奧爾加農 · 亦創造 solmization 與四線譜系統"],
    ["■", "Wulfstan of Winchester", "fl. 992–996", "Winchester Troper 中 174 首奧爾加農的作者——現存最早的大型英格蘭複音曲集"],
    ["■", "Léonin (Leoninus)", "fl. 1150s – ca. 1201", "巴黎聖母院神職與詩人 · Anonymous IV 稱其為「excellent organista」· 與《Magnus liber organi》的編纂密不可分"],
    ["■", "Pérotin (Perotinus)", "fl. late 12th – early 13th c.", "「Perotinus the Great」· 擴充 Magnus liber · 首創三聲部與四聲部 organum · 代表作 Viderunt omnes（1198）"],
    ["■", "Johannes de Garlandia", "ca. 1260", "De mensurabili musica 論文整理六種節奏模式——把即興傳統系統化為可書寫的節奏體系"],
    ["■", "Franco of Cologne", "ca. 1280", "Ars cantus mensurabilis —— 確立用音符形狀本身表示時值的記譜法，現代記譜法的直接祖先"],
    ["■", "Adam de la Halle", "ca. 1240–1288", "同時是 trouvère 與經文歌作者 · De ma dame vient / Dieus / Omnes 體現 Franconian 經文歌風格"],
    ["■", "Petrus de Cruce", "mid-13th – early 14th c.", "Pierre de la Croix · 進一步細分 semibreves · 開啟 Ars nova 的節奏複雜性之先聲"],
  ];

  figures.forEach(([icon, name, date, desc], i) => {
    const y = 0.9 + i * 0.56;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.5, fill: { color: "3A2015" }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.08, w: 0.5, h: 0.35, fontSize: 16, align: "center", margin: 0 });
    s.addText(name, { x: 0.95, y: y + 0.03, w: 3.0, h: 0.24, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(date, { x: 0.95, y: y + 0.25, w: 3.0, h: 0.22, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 });
    s.addShape(pres.ShapeType.rect, { x: 4.0, y: y + 0.08, w: 0.025, h: 0.35, fill: { color: C.gold } });
    s.addText(desc, { x: 4.1, y: y + 0.03, w: 5.5, h: 0.45, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 14 · Timeline ──────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("歷史時間軸 Timeline", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const events = [
    ["ca. 850–890",   "Musica enchiriadis · 最早描述平行奧爾加農"],
    ["ca. 992–996",   "Wulfstan of Winchester · Winchester Troper 中 174 首奧爾加農"],
    ["11th c.",       "羅馬式教堂與修院時期 · 音對音奧爾加農"],
    ["ca. 1025–28",   "Guido of Arezzo · Micrologus"],
    ["ca. 1100",      "Ad organum faciendum · 奧爾加農實用教本"],
    ["1109",          "St. Anselm 卒 · 經院哲學之父"],
    ["12th c.",       "Aquitanian polyphony（St. Martial · Limoges）"],
    ["ca. 1140",      "最早的哥德式建築出現"],
    ["mid-12th c.",   "巴黎大學成立"],
    ["1150s–ca.1201", "Léonin 於巴黎主教座堂活動"],
    ["ca. 1160–1250", "巴黎聖母院建造"],
    ["1173",          "Codex Calixtinus（聖地牙哥德孔波斯特拉）"],
    ["1198",          "Pérotin · Viderunt omnes（四聲部）"],
    ["1209",          "St. Francis of Assisi 創立方濟會"],
    ["ca. 1250",      "Sumer is icumen in（現存最早的英語 rota）"],
    ["ca. 1260",      "Johannes de Garlandia · De mensurabili musica"],
    ["1267–73",       "St. Thomas Aquinas · Summa Theologica"],
    ["ca. 1280",      "Franco of Cologne · Ars cantus mensurabilis"],
    ["ca. 1285",      "Anonymous IV 論文（記載 Léonin、Pérotin）"],
  ];

  s.addShape(pres.ShapeType.rect, { x: 2.8, y: 0.9, w: 0.05, h: 4.4, fill: { color: C.wine } });

  events.forEach(([date, event], i) => {
    const y = 0.9 + i * 0.24;
    s.addShape(pres.ShapeType.ellipse, { x: 2.67, y: y + 0.04, w: 0.26, h: 0.26, fill: { color: C.wine } });
    s.addText(date, { x: 0.2, y, w: 2.45, h: 0.26, fontSize: 14, color: C.midBrown, fontFace: "Calibri", valign: "top", align: "right", margin: 0 });
    s.addText(event, { x: 3.05, y, w: 6.6, h: 0.26, fontSize: 14, color: C.darkText, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 15 · A Polyphonic Tradition (Chapter Summary) ──────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("複音音樂的遺產 A Polyphonic Tradition", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("Chapter Summary · Why This Music Mattered", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const points = [
    ["■", "11–13 世紀複音興起：對位、和聲、垂直音響成為西方傳統核心\nPolyphony 11th–13th c.: counterpoint and harmony enter Western tradition"],
    ["■", "記譜法引入兩大特徵：聲部垂直對齊、音符形狀表時值\nNotation: vertical alignment of parts; note shapes indicating duration"],
    ["■", "聖母院學派的音樂如大教堂——長度、結構、多聲部協調前所未見\nNotre Dame music matched the cathedral: length, structure, grandeur"],
    ["■", "經文歌由 clausula 填詞演變為世俗娛樂，主導中世紀晚期\nThe motet evolved from clausula trope to elite entertainment"],
    ["■", "大多中世紀複音一兩代後即被視為粗糙，20 世紀才重新發掘\nMost medieval polyphony faded quickly, revived only in the 20th c."],
    ["■", "中世紀聲響回聲於 Debussy 的平行和弦、極簡主義、嘻哈分層文本\nMedieval sounds echo in Debussy, minimalism, and hip hop"],
  ];

  points.forEach(([icon, text], i) => {
    const y = 1.25 + i * 0.68;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.6, fill: { color: "3A2015" }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.08, w: 0.55, h: 0.45, fontSize: 18, align: "center", margin: 0 });
    s.addText(text, { x: 1.05, y: y + 0.05, w: 8.5, h: 0.55, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 16 · Further Reading & Key Terms ───────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("延伸閱讀與補充教材\nFurther Reading & Supplementary Resources", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.72,
    fontSize: 22, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.95, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Listening (full-width)
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.05, w: 9.2, h: 2.4, fill: { color: C.wine }, rounding: true });
  s.addText("■ 聆聽 Listen", {
    x: 0.55, y: 1.1, w: 8.9, h: 0.3, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", valign: "top", margin: 0,
  });
  const listening = [
    "Musica enchiriadis — parallel organum (NAWM 14)  youtu.be/kGnCQKRwMfc",
    "Aquitanian: Jubilemus, exultemus (NAWM 15)  youtu.be/22srNVUmF3Q",
    "Léonin — Viderunt omnes, duplum (NAWM 17)  youtu.be/_p9WQlyVPrA",
    "Pérotin — Viderunt omnes, quadruplum (NAWM 19)  youtu.be/fPZrMpFnygw",
    "Factum est salutare / Dominus (NAWM 20)  youtu.be/iLKKwuirSoo",
    "Adam — De ma dame vient (NAWM 21)  youtu.be/zGV5APEq7bc",
    "Sumer is icumen in (NAWM 23)  youtu.be/sMCA9nYnLWo",
  ];
  s.addText(listening.map((l, i) => ({
    text: l, options: { bullet: true, breakLine: i < listening.length - 1, fontSize: 14, color: C.cream, fontFace: "Calibri" },
  })), { x: 0.55, y: 1.4, w: 8.9, h: 2.0, valign: "top", margin: 0 });

  // Reading (2-col)
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 3.55, w: 9.2, h: 1.2, fill: { color: C.rust }, rounding: true });
  s.addText("■ 閱讀 Read", {
    x: 0.55, y: 3.59, w: 8.9, h: 0.28, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", valign: "top", margin: 0,
  });
  const readingLeft = [
    "Hoppin — Medieval Music",
    "Roesner — Who 'Made' the Magnus Liber?",
    "Wright — Notre Dame of Paris",
  ];
  const readingRight = [
    "Fuller — Early Polyphony (NOHM)",
    "Wikipedia: Organum · Motet",
    "Grout HWM Online · Ch 5",
  ];
  s.addText(readingLeft.map((r, i) => ({
    text: r, options: { bullet: true, breakLine: i < readingLeft.length - 1, fontSize: 14, color: C.cream, fontFace: "Calibri" },
  })), { x: 0.55, y: 3.89, w: 4.4, h: 0.82, valign: "top", margin: 0 });
  s.addText(readingRight.map((r, i) => ({
    text: r, options: { bullet: true, breakLine: i < readingRight.length - 1, fontSize: 14, color: C.cream, fontFace: "Calibri" },
  })), { x: 5.05, y: 3.89, w: 4.4, h: 0.82, valign: "top", margin: 0 });

  // Key terms (compact)
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 4.85, w: 9.2, h: 0.6, fill: { color: C.midBrown }, rounding: true });
  s.addText("■ Key Terms", {
    x: 0.55, y: 4.88, w: 1.6, h: 0.25, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", valign: "top", margin: 0,
  });
  const terms = "Organum · Polyphony · Musica enchiriadis · Magnus liber organi · Léonin / Pérotin · Clausula · Conductus · Motet · Rhythmic modes · Sumer is icumen in";
  s.addText(terms, {
    x: 2.2, y: 4.88, w: 7.25, h: 0.55, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0,
  });
}

// ── Generate file ─────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "Ch05_Polyphony.pptx" })
  .then(() => console.log("■ Ch05_Polyphony.pptx created successfully"))
  .catch(err => console.error("■ Error:", err));
