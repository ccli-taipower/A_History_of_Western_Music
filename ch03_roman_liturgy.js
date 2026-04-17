const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 3: Roman Liturgy and Chant";
pres.author = "A History of Western Music, 10th ed.";

// ── Color palette (same parchment/ancient theme as Ch01/Ch02) ────────────────
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

  s.addText("CHAPTER 3", {
    x: 0.5, y: 0.9, w: 9, h: 0.55,
    fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });

  s.addText("ROMAN LITURGY\nAND CHANT", {
    x: 0.3, y: 1.4, w: 9.4, h: 2.0,
    fontSize: 44, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
    paraSpaceAfter: 0,
  });

  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.5, w: 5, h: 0.04, fill: { color: C.gold } });

  s.addText("The Mass & Office · Genres of Chant · Tropes · Sequences · Hildegard", {
    x: 0.4, y: 3.65, w: 9.2, h: 0.4,
    fontSize: 14, color: C.sand, italic: true, align: "center", fontFace: "Georgia",
  });

  s.addText("Textbook pp. 42–62", {
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
    ["■", "羅馬禮儀 The Roman Liturgy", "教會日曆、彌撒（Mass）與聖務日課（Office）的架構"],
    ["■", "聖詠的特徵 Characteristics of Chant", "演唱方式、歌詞設定風格、旋律與唱詞的關係"],
    ["■", "聖詠的類型 Genres of Chant", "吟誦公式、對唱歌、讚美詩、Gradual、Alleluia、Tract、Offertory"],
    ["■", "新添加的聖詠 Additions to Chant", "Trope（增飾）、Sequence（繼抒詠）、禮儀劇（Liturgical Drama）"],
    ["■", "賓根的希爾德加 Hildegard of Bingen", "中世紀女修道院長、作曲家、神秘主義者（1098–1179）"],
    ["■", "聖詠的延續 The Continuing Presence", "從中世紀到當代：影響、改革、流行文化中的再生"],
  ];

  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.wine }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 14, color: C.midBrown, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 3 · The Roman Liturgy (Purpose & Calendar) ─────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("羅馬禮儀", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("The Roman Liturgy · Purpose & Church Calendar", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 16, color: C.sand, fontFace: "Georgia", align: "center", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Purpose box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.25, w: 9.4, h: 1.9, fill: { color: "3A2015" }, rounding: true });
  s.addText("◆ 禮儀的目的 Purpose of the Liturgy", {
    x: 0.45, y: 1.32, w: 9.1, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0, valign: "top",
  });
  s.addText("禮儀（liturgy）是教會中誦念或歌唱的文本與所進行的儀式；其雙重目的：\nThe liturgy = texts spoken/sung and rituals performed during services. Dual aim:\n① 對信徒：強化教義並引導走向救恩 · Reinforce doctrine and clarify the path to salvation\n② 對上帝：如天使般不斷獻上頌讚 · Addressing God as the primary audience of unceasing praise", {
    x: 0.5, y: 1.72, w: 9.0, h: 1.4, fontSize: 14, color: C.sand, fontFace: "Calibri", margin: 0, valign: "top",
  });

  // Calendar
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.25, w: 9.4, h: 1.5, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ 教會日曆 Church Calendar", {
    x: 0.45, y: 3.32, w: 9.1, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0, valign: "top",
  });

  const feasts = [
    ["■ 聖誕節 Christmas", "12/25 · 紀念耶穌誕生", "待降節 Advent：聖誕前四主日"],
    ["■ 復活節 Easter", "春分滿月後主日 · 耶穌復活", "大齋期 Lent：復活節前 46 天"],
  ];
  feasts.forEach(([title, date, prep], i) => {
    const y = 3.72 + i * 0.48;
    s.addText(title, { x: 0.55, y, w: 2.6, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0, valign: "top" });
    s.addText(date, { x: 3.2, y, w: 3.2, h: 0.4, fontSize: 14, color: C.cream, fontFace: "Calibri", margin: 0, valign: "top" });
    s.addText(prep, { x: 6.45, y, w: 3.15, h: 0.4, fontSize: 14, color: C.sand, italic: true, fontFace: "Calibri", margin: 0, valign: "top" });
  });

  s.addText("每年教會以「節日」(feast day) 紀念聖經事件與聖徒；大部分禮儀相同，部分隨季節改變\nThe church commemorates events and saints with feast days; most liturgy is fixed, parts change seasonally", {
    x: 0.5, y: 4.85, w: 9, h: 0.55, fontSize: 14, color: C.sand, italic: true, fontFace: "Calibri", align: "center", margin: 0, valign: "top",
  });
}

// ── SLIDE 4 · The Mass: Proper vs Ordinary ──────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("彌撒 The Mass", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 28, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addText("羅馬教會最重要的儀式 · 源自最後晚餐的象徵性重演 · Central ritual: symbolic reenactment of the Last Supper", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.35, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Calibri", margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.06, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Proper box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.wine }, rounding: true });
  s.addText("■ Proper 變化部分", {
    x: 0.45, y: 1.28, w: 4.3, h: 0.4, fontSize: 16, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("歌詞隨教會日曆變化；以「功能」命名", {
    x: 0.45, y: 1.7, w: 4.3, h: 0.3, fontSize: 14, color: C.cream, italic: true, fontFace: "Calibri", valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.55, y: 2.05, w: 4.1, h: 0.02, fill: { color: C.gold } });

  const properItems = [
    "1. Introit 進堂詠（選自詩篇）",
    "4. Collect 集禱經",
    "5. Epistle 書信誦讀",
    "6. Gradual 升階經 ■",
    "7. Alleluia 或 Tract ■",
    "8. Sequence 繼抒詠（大節慶）",
    "9. Gospel 福音",
    "12. Offertory 奉獻詠 ■",
    "20. Communion 領主曲",
  ];
  s.addText(properItems.map((t, i) => ({
    text: t, options: { bullet: false, breakLine: i < properItems.length - 1, fontSize: 14, color: C.cream, fontFace: "Calibri", paraSpaceAfter: 2 }, valign: "top",
  })), { x: 0.5, y: 2.1, w: 4.2, h: 3.0 });

  // Ordinary box
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.midBrown }, rounding: true });
  s.addText("■ Ordinary 固定部分", {
    x: 5.25, y: 1.28, w: 4.3, h: 0.4, fontSize: 16, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("歌詞永不變（雖旋律可不同）；以首句命名", {
    x: 5.25, y: 1.7, w: 4.3, h: 0.3, fontSize: 14, color: C.sand, italic: true, fontFace: "Calibri", valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 5.35, y: 2.05, w: 4.1, h: 0.02, fill: { color: C.gold } });

  const ordinaryItems = [
    "2. Kyrie 求主垂憐（希臘文）",
    "3. Gloria 榮耀頌",
    "11. Credo 信經（1014 最後加入）",
    "16. Sanctus 聖哉經",
    "19. Agnus Dei 羔羊讚",
    "",
    "■ 14 世紀以後大多數「彌撒曲」",
    "    專指這五段的音樂設定：",
    "    Kyrie · Gloria · Credo",
    "    Sanctus · Agnus Dei",
  ];
  s.addText(ordinaryItems.map((t, i) => ({
    text: t, options: { bullet: false, breakLine: i < ordinaryItems.length - 1, fontSize: 14, color: C.cream, fontFace: "Calibri", paraSpaceAfter: 2 }, valign: "top",
  })), { x: 5.3, y: 2.1, w: 4.3, h: 3.0 });

  s.addText("■ = 彌撒中最華麗的音樂高潮（Gradual, Alleluia, Offertory）—— 由獨唱者與唱詩班以應答式演唱", {
    x: 0.4, y: 5.26, w: 9.2, h: 0.22, fontSize: 14, color: C.rust, italic: true, fontFace: "Calibri", align: "center", valign: "top",
  });
}

// ── SLIDE 5 · The Office: Eight Daily Hours ─────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("聖務日課", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.5, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("The Divine Office · Eight Daily Services", {
    x: 0.4, y: 0.68, w: 9.2, h: 0.35, fontSize: 15, color: C.sand, italic: true, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.06, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addText("依《聖本篤會規》（ca. 530）編纂 · 在修道院中每日八次按時舉行 · 每週唱完全部 150 首詩篇\nCodified in the Rule of St. Benedict (ca. 530); eight services daily in monasteries; all 150 psalms sung weekly", {
    x: 0.4, y: 1.18, w: 9.2, h: 0.6, fontSize: 14, color: C.sand, italic: true, fontFace: "Calibri", align: "center", valign: "top",
  });

  // 8 Hours timeline — 4 columns × 2 rows
  const hours = [
    ["■", "Matins\n晨禱",     "午夜之後",       "midnight–dawn",      "*最重要（音樂上）"],
    ["■", "Lauds\n讚美禱",    "黎明時分",       "at sunrise",          "*最重要（音樂上）"],
    ["■",  "Prime\n初時課",    "6 AM",           "first hour",          "Little Hours"],
    ["■", "Terce\n三時課",    "9 AM",           "third hour",          "Little Hours"],
    ["■", "Sext\n六時課",     "正午",           "midday",              "Little Hours"],
    ["■", "None\n九時課",     "3 PM",           "ninth hour",          "Little Hours"],
    ["■", "Vespers\n晚禱",    "日落時分",       "at sunset",           "*最重要（音樂上）"],
    ["■", "Compline\n安息禱", "就寢前",         "before bed",          "結束一日"],
  ];

  hours.forEach(([icon, name, zh, en, note], i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    const x = 0.3 + col * 2.4;
    const y = 1.95 + row * 1.62;
    const isMajor = note.startsWith("*");
    s.addShape(pres.ShapeType.rect, { x, y, w: 2.25, h: 1.5, fill: { color: isMajor ? C.wine : "3A2015" }, rounding: true });
    s.addText(icon, { x, y: y + 0.02, w: 2.25, h: 0.3, fontSize: 18, align: "center" });
    s.addText(name, { x: x + 0.05, y: y + 0.3, w: 2.15, h: 0.48, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", align: "center", margin: 0 });
    s.addText(`${zh}\n${en}`, { x: x + 0.05, y: y + 0.82, w: 2.15, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", align: "center", margin: 0 });
    s.addText(note.replace("*", ""), { x: x + 0.05, y: y + 1.22, w: 2.15, h: 0.25, fontSize: 14, color: isMajor ? C.gold : C.sand, italic: true, fontFace: "Calibri", valign: "top", align: "center", margin: 0 });
  });

  s.addText("■ Matins, Lauds, Vespers = 音樂上最重要的三個時辰 · The three musically most important Hours", {
    x: 0.4, y: 5.22, w: 9.2, h: 0.22, fontSize: 14, color: C.gold, italic: true, fontFace: "Calibri", align: "center", valign: "top",
  });
}

// ── SLIDE 6 · Characteristics of Chant ──────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("聖詠的特徵 Characteristics of Chant", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Two columns: Performance manner + Text setting
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 0.92, w: 4.65, h: 3.95, fill: { color: C.wine }, rounding: true });
  s.addText("◆ 演唱方式 Manner of Performance", {
    x: 0.45, y: 1.0, w: 4.4, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.55, y: 1.38, w: 4.2, h: 0.02, fill: { color: C.gold } });

  const manners = [
    ["① Responsorial", "獨唱者與唱詩班/會眾交替應答\nSoloist alternates with choir or congregation"],
    ["② Antiphonal", "唱詩班分成兩組輪流應答\nTwo halves of choir alternate"],
    ["③ Direct", "不分組，連續演唱\nNo alternation; sung straight through"],
  ];
  manners.forEach(([title, desc], i) => {
    const y = 1.5 + i * 1.1;
    s.addText(title, { x: 0.55, y, w: 4.2, h: 0.3, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0, valign: "top" });
    s.addText(desc, { x: 0.55, y: y + 0.32, w: 4.2, h: 0.75, fontSize: 14, color: C.cream, fontFace: "Calibri", margin: 0, valign: "top" });
  });

  // Right column
  s.addShape(pres.ShapeType.rect, { x: 5.05, y: 0.92, w: 4.65, h: 3.95, fill: { color: C.rust }, rounding: true });
  s.addText("◆ 歌詞設定 Text Setting", {
    x: 5.2, y: 1.0, w: 4.4, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 5.3, y: 1.38, w: 4.2, h: 0.02, fill: { color: C.gold } });

  const settings = [
    ["Syllabic 音節式", "幾乎每個音節對應一個音\nAlmost every syllable gets a single note"],
    ["Neumatic 紐姆式", "每個音節 1–6 個音（一個紐姆）\nOne to six notes per syllable"],
    ["Melismatic 花腔式", "單一音節上的長串旋律\nLong melodic passages on a single syllable"],
  ];
  settings.forEach(([title, desc], i) => {
    const y = 1.5 + i * 1.1;
    s.addText(title, { x: 5.3, y, w: 4.2, h: 0.3, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0, valign: "top" });
    s.addText(desc, { x: 5.3, y: y + 0.32, w: 4.2, h: 0.75, fontSize: 14, color: C.cream, fontFace: "Calibri", margin: 0, valign: "top" });
  });

  s.addText("旋律反映拉丁語朗誦：多數樂句呈「弧形」（起於低音，升至高點，下行結束）；重音節常用高音\nMelodies reflect Latin declamation: phrases arc (low → peak → descend); accented syllables get higher notes", {
    x: 0.4, y: 4.95, w: 9.2, h: 0.5, fontSize: 14, color: C.wine, italic: true, fontFace: "Calibri", align: "center", margin: 0, valign: "top",
  });
}

// ── SLIDE 7 · Recitation Formulas & Psalm Tones ─────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("吟誦公式與詩篇調式", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("Recitation Formulas & Psalm Tones · The Simplest Chants", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.35, fontSize: 15, color: C.sand, italic: true, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.06, w: 7, h: 0.04, fill: { color: C.gold } });

  // Left: Recitation Formulas
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.2, w: 4.6, h: 4.05, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ Recitation Formulas 吟誦公式", {
    x: 0.45, y: 1.28, w: 4.3, h: 0.38, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.55, y: 1.68, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("最簡單的聖詠——用於誦讀祈禱與聖經（Collect、Epistle、Gospel）\nThe simplest chants, for prayers and Bible readings (Collect, Epistle, Gospel)\n\n• 幾乎完全音節式（syllabic）\n• 在「誦念音」（reciting note, 通常 A 或 C）上唸讀\n• 短動機標示句末\n• 比調式系統更古老——不屬任何調式\n• 由神父或助祭演唱", {
    x: 0.5, y: 1.75, w: 4.2, h: 3.4, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top",
  });

  // Right: Psalm Tones
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.2, w: 4.6, h: 4.05, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ Psalm Tones 詩篇調式", {
    x: 5.25, y: 1.28, w: 4.3, h: 0.38, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 5.35, y: 1.68, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("較吟誦公式稍複雜；聖務日課中誦唱詩篇用 · 八種調式各有其專屬詩篇調\nSlightly more complex; for Office psalms. One tone for each of the 8 modes.\n\n詩篇調的五部分：\nFive parts of a psalm tone:\n① Intonation 起唱（只用於第一節）\n② Recitation 誦念音（誦唱大部分歌詞）\n③ Mediant 中止（詩節中間的半終止）\n④ Recitation（繼續誦念）\n⑤ Termination 終止（詩節結尾）", {
    x: 5.3, y: 1.75, w: 4.2, h: 3.4, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 8 · Office Antiphons & Hymns ──────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("日課中的對唱歌與讚美詩", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.5, fontSize: 24, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addText("Office Antiphons & Hymns", {
    x: 0.4, y: 0.68, w: 9.2, h: 0.35, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.04, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Antiphon box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.18, w: 9.4, h: 2.0, fill: { color: C.wine }, rounding: true });
  s.addText("◆ Antiphon 對唱歌", {
    x: 0.45, y: 1.24, w: 9.1, h: 0.32, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0, valign: "top",
  });
  s.addText("在詩篇（或讚歌 canticle）前後演唱的短聖詠；歌詞隨教會日曆變化\nA short chant framing a psalm (before and after); text varies with the calendar\n• 對唱歌的調式決定搭配的詩篇調（mode 1 → mode 1 psalm tone）\n• 每個詩篇調有多種終止式，以流暢銜接不同對唱歌開頭\n• 簡單、以音節式為主（每日要唱 30 多個）· Mostly syllabic; 30+ sung daily", {
    x: 0.5, y: 1.6, w: 9.0, h: 1.55, fontSize: 14, color: C.cream, fontFace: "Calibri", margin: 0, valign: "top",
  });

  // Hymn box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.25, w: 9.4, h: 2.2, fill: { color: C.rust }, rounding: true });
  s.addText("◆ Hymn 讚美詩", {
    x: 0.45, y: 3.31, w: 9.1, h: 0.32, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0, valign: "top",
  });
  s.addText("最熟悉的宗教歌曲類型；每個聖務日課都要唱一首\nThe most familiar type of sacred song, sung in every Office service\n• 分節歌式 (strophic)：數段共用同一旋律\n• 每段 4–7 行，有時押韻；多數音節式，旋律常重複若干樂句\n範例 Example：Christe Redemptor omnium (NAWM 4b)", {
    x: 0.5, y: 3.65, w: 9.0, h: 1.7, fontSize: 14, color: C.cream, fontFace: "Calibri", margin: 0, valign: "top",
  });
}

// ── SLIDE 9 · Mass Chants Overview ──────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("彌撒中的聖詠類型", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.5, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("Chants of the Mass · From Simple to Elaborate", {
    x: 0.4, y: 0.68, w: 9.2, h: 0.35, fontSize: 15, color: C.sand, italic: true, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.06, w: 7, h: 0.04, fill: { color: C.gold } });

  const chants = [
    ["Introit 進堂詠", "Antiphonal", "Neumatic", "開堂遊行時由唱詩班唱 · 對唱歌 + 一詩節 + 小聖三頌\nEntrance chant — antiphon + psalm verse + Lesser Doxology"],
    ["Communion 領主曲", "Antiphonal", "Neumatic", "領聖體時演唱 · 後來縮短為僅對唱歌\nSung during communion; eventually shortened to antiphon only"],
    ["Gradual 升階經 ■", "Responsorial", "Melismatic", "書信誦讀之後 · 以階梯（gradus）得名 · 獨唱者 + 唱詩班\nAfter Epistle; named from the stair (gradus); soloist + choir"],
    ["Alleluia 哈利路亞 ■", "Responsorial", "Melismatic", "「alleluia」的最後音節延伸為「jubilus」長花腔——「言語無法表達的喜悅」\nExtended jubilus melisma on final syllable: \"joy beyond words\""],
    ["Tract 牽引經", "Direct solo", "Very florid", "大齋期中取代 Alleluia · 最長的聖詠，僅限 mode 2 或 8\nReplaces Alleluia in Lent; longest chants, only in modes 2 or 8"],
    ["Offertory 奉獻詠 ■", "Responsorial", "Melismatic", "神父預備餅酒時演唱 · 如 Gradual 般華麗\nDuring the offering of bread and wine; as melismatic as Graduals"],
  ];

  chants.forEach(([name, manner, style, desc], i) => {
    const y = 1.2 + i * 0.67;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.6, fill: { color: "3A2015" }, rounding: true });
    s.addText(name, { x: 0.45, y: y + 0.05, w: 2.1, h: 0.5, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(`${manner}\n${style}`, { x: 2.55, y: y + 0.05, w: 1.9, h: 0.52, fontSize: 14, color: C.sand, italic: true, fontFace: "Calibri", valign: "top", margin: 0 });
    s.addShape(pres.ShapeType.rect, { x: 4.5, y: y + 0.1, w: 0.025, h: 0.42, fill: { color: C.gold } });
    s.addText(desc, { x: 4.6, y: y + 0.04, w: 5.05, h: 0.55, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 10 · Tropes ───────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("增飾 Tropes", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.5, fontSize: 26, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addText("對已有聖詠的擴充：新增文字、音樂或兩者 · Expansions of existing chants: new words, music, or both", {
    x: 0.4, y: 0.66, w: 9.2, h: 0.35, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Calibri", margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.02, w: 9.2, h: 0.03, fill: { color: C.sand } });

  s.addText("Trope 的功能：在「授權聖詠」邊緣提供創作出口，類似中世紀學者在聖經邊註評論。文字增飾提供「釋義」（gloss），把聖詠文本與當日場合更緊密連結。\nTropes gave musicians creative outlets in the margins of the authorized repertory. Added words served as glosses, linking chant texts more closely to the occasion.", {
    x: 0.4, y: 1.15, w: 9.2, h: 0.65, fontSize: 14, color: C.slate, italic: true, fontFace: "Calibri", valign: "top",
  });

  const tropes = [
    ["① 新文字 + 新音樂", "Text + Music", "在聖詠之前（introductory trope）\n或每個樂句之前（intercalated）\nBefore the chant or before each phrase",
     "範例：Quem queritis in presepe（聖誕節 Introit 前的對話型 trope，late 10th c.）"],
    ["② 僅音樂", "Music only", "延伸既有 melisma 或新增 melisma\nExtending existing melismas or adding new ones",
     "Introit 結尾加入無字花腔 textless melisma"],
    ["③ 僅文字", "Text only — prosula", "為既有 melisma 填上文字，使之成為音節式\nAdding text to existing melismas",
     "新詞的母音常模仿被取代音節的母音——產生迴響效果"],
  ];

  tropes.forEach(([title, subtitle, mech, ex], i) => {
    const y = 1.88 + i * 1.12;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 1.0, fill: { color: C.wine }, rounding: true });
    s.addText(title, { x: 0.45, y: y + 0.08, w: 2.6, h: 0.3, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(subtitle, { x: 0.45, y: y + 0.38, w: 2.6, h: 0.26, fontSize: 14, color: C.sand, italic: true, fontFace: "Calibri", valign: "top", margin: 0 });
    s.addShape(pres.ShapeType.rect, { x: 3.2, y: y + 0.12, w: 0.03, h: 0.76, fill: { color: C.gold } });
    s.addText(mech, { x: 3.3, y: y + 0.08, w: 3.3, h: 0.85, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0 });
    s.addText(ex, { x: 6.65, y: y + 0.08, w: 3.0, h: 0.85, fontSize: 14, color: C.sand, italic: true, fontFace: "Calibri", valign: "top", margin: 0 });
  });

  s.addText("興盛於 9-11 世紀修道院；12 世紀後式微；1562-63 特倫多會議全面禁止", {
    x: 0.4, y: 5.2, w: 9.2, h: 0.25, fontSize: 14, color: C.rust, italic: true, fontFace: "Calibri", align: "center", valign: "top",
  });
}

// ── SLIDE 11 · Sequences ────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("繼抒詠 Sequences", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.5, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("Freestanding Chants Sung after the Alleluia · 9th – 13th Centuries", {
    x: 0.4, y: 0.68, w: 9.2, h: 0.35, fontSize: 14, color: C.sand, italic: true, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.06, w: 7, h: 0.04, fill: { color: C.gold } });

  // Form box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.2, w: 9.4, h: 1.35, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ 形式 Form", {
    x: 0.45, y: 1.28, w: 9.1, h: 0.3, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("以音節式設定的詩句，多數成對（couplet）· 一對之內兩句音節數相同、共用同一旋律 · 每對新音樂、新歌詞\n結構：A BB CC DD … N（首末單獨詩句 + 中間成對詩句）· 與 Alleluia 之後的 sequentia 長花腔有歷史關聯", {
    x: 0.5, y: 1.6, w: 9.0, h: 0.85, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top",
  });

  // Three famous sequences
  const seqs = [
    ["■ Notker Balbulus\n諾特克·巴爾布魯斯", "ca. 840–912", "聖加侖修道院的法蘭克修士；最著名的早期繼抒詠詞作者；884 完成《Liber hymnorum》\nFrankish monk at St. Gall; most famous early sequence writer; completed Liber hymnorum 884"],
    ["■ Wipo\n維波", "ca. 995–1050", "神聖羅馬皇帝的宮廷牧師；創作 Victimae paschali laudes（復活節繼抒詠）——現存四首仍被保留的古繼抒詠之一\nImperial chaplain; composed Victimae paschali laudes (Easter)"],
    ["■ Thomas of Celano\n切拉諾的托馬斯", "ca. 1190–1260", "聖方濟各的傳記作者；創作 Dies irae（末日經）——中世紀最著名詩作之一；形式 AA BB CC / AA BB CC / AA BB C / D E\nBiographer of St. Francis; composed Dies irae — one of the most famous poems of the Middle Ages"],
  ];

  seqs.forEach(([name, date, desc], i) => {
    const y = 2.68 + i * 0.88;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.8, fill: { color: C.wine }, rounding: true });
    s.addText(name, { x: 0.45, y: y + 0.06, w: 2.6, h: 0.65, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(date, { x: 3.1, y: y + 0.1, w: 1.3, h: 0.3, fontSize: 14, color: C.sand, italic: true, fontFace: "Calibri", valign: "top", margin: 0 });
    s.addShape(pres.ShapeType.rect, { x: 4.45, y: y + 0.12, w: 0.03, h: 0.55, fill: { color: C.gold } });
    s.addText(desc, { x: 4.55, y: y + 0.06, w: 5.1, h: 0.7, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0 });
  });

  s.addText("■ 特倫多會議（1562–63）禁用大多數繼抒詠，僅保留四首（含 Victimae paschali laudes 與 Dies irae）", {
    x: 0.4, y: 5.26, w: 9.2, h: 0.22, fontSize: 14, color: C.gold, italic: true, fontFace: "Calibri", align: "center", valign: "top",
  });
}

// ── SLIDE 12 · Liturgical Drama ─────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("禮儀劇 Liturgical Drama", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addText("從對話型 trope 發展出的宗教戲劇 · Dramatic dialogues added to the liturgy", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.35, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Calibri", margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.06, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Quem queritis in sepulchro
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.2, w: 9.4, h: 2.15, fill: { color: C.wine }, rounding: true });
  s.addText("◆ Quem queritis in sepulchro（10 世紀）", {
    x: 0.45, y: 1.26, w: 9.1, h: 0.32, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0, valign: "top",
  });
  s.addText("The earliest known liturgical drama · 在復活節 Introit 前演唱的對話型 trope", {
    x: 0.45, y: 1.58, w: 9.1, h: 0.28, fontSize: 14, color: C.sand, italic: true, fontFace: "Calibri", margin: 0, valign: "top",
  });
  s.addText("對話內容（取自馬可福音 16:5–7）：\n天使（Angel）：「Quem queritis in sepulchro? (你們在墓中尋找誰？)」\n三位瑪利亞：「Jesum Nazarenum. (尋找拿撒勒人耶穌)」\n天使：「Non est hic, surrexit sicut predixerat... (祂不在這裡，祂已如所預言復活)」", {
    x: 0.5, y: 1.9, w: 9.0, h: 1.4, fontSize: 14, color: C.cream, fontFace: "Calibri", margin: 0, valign: "top",
  });

  // Features box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.45, w: 9.4, h: 2.0, fill: { color: C.rust }, rounding: true });
  s.addText("◆ 禮儀劇的特徵 Features", {
    x: 0.45, y: 3.52, w: 9.1, h: 0.32, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0, valign: "top",
  });

  const features = [
    "• 以應答方式演唱，並伴以戲劇性動作 · Sung responsively with dramatic action",
    "• 記錄於禮儀書中，在教堂內演出 · Recorded in liturgical books, performed in church",
    "• 最常見：復活節與聖誕節的對話劇 · Easter and Christmas dialogues most common",
    "• 12 世紀後更多劇本：弗勒里 Fleury 10 齣、Beauvais《但以理劇》(ca. 1210s)",
    "• 大多角色由男性神職人員演唱；極少地方允許修女參與",
  ];
  s.addText(features.map((t, i) => ({
    text: t, options: { bullet: false, breakLine: i < features.length - 1, fontSize: 14, color: C.cream, fontFace: "Calibri", paraSpaceAfter: 3 }, valign: "top",
  })), { x: 0.5, y: 3.85, w: 9.0, h: 1.55, margin: 0, valign: "top" });
}

// ── SLIDE 13 · Hildegard of Bingen ──────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("賓根的希爾德加", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 30, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("Hildegard of Bingen (1098–1179)", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.35, fontSize: 16, italic: true, color: C.sand, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.1, w: 7, h: 0.04, fill: { color: C.gold } });

  // Biography box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.2, w: 4.6, h: 4.05, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ 生平 Biography", {
    x: 0.45, y: 1.28, w: 4.3, h: 0.38, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.55, y: 1.68, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("• 德國萊茵地區貴族家庭出身\n  Born to noble family in the Rhineland\n• 8 歲被父母獻給教會\n  Consecrated to church at age 8\n• 本篤會修道院德西博登貝格\n  Disibodenberg Benedictine monastery\n• 1136：當選修女院長（magistra）\n• ca. 1150：創立魯柏斯貝格修道院，任院長（abbess）\n• 以預言與異象聞名，與教宗、皇帝、主教通信\n  Famous prophetess; corresponded with popes, emperors\n• 著作：Scivias（異象錄，1141–51）、Physica（自然學）、Causae et Curae（醫學）", {
    x: 0.5, y: 1.78, w: 4.2, h: 3.4, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top",
  });

  // Music box
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.2, w: 4.6, h: 4.05, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ 音樂作品 Music", {
    x: 5.25, y: 1.28, w: 4.3, h: 0.38, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 5.35, y: 1.68, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("• 是中世紀留下最多聖詠的作曲家\n  More surviving chants than any other medieval composer\n• 詞曲皆自創 · Both words and music by her\n• 主要作品：43 首對唱歌、18 首應答歌、7 首繼抒詠、4 首讚美詩\n• 旋律特色：音域常超過八度四或五度；反覆運用少量旋律動機\n  Melodies often exceed an octave by a 4th or 5th\n• Ordo virtutum（德行劇，ca. 1151）：\n  — 最早的非禮儀性宗教音樂劇\n  — 82 首歌曲 · 82 songs\n  — 角色：先知、德行、快樂/悲傷/懺悔的靈魂\n  — 魔鬼只能「說」不能「唱」——象徵與神的分離\n    Devil can only speak, not sing — symbolizing separation from God", {
    x: 5.3, y: 1.78, w: 4.2, h: 3.4, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 14 · Timeline ─────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("歷史時間軸 Timeline", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });

  const events = [
    ["4th–7th c. CE",    "Introit 進堂詠加入彌撒 · Introit added to Mass"],
    ["ca. 530 CE",       "《聖本篤會規》編纂日課禮儀 · Rule of St. Benedict codifies Office liturgy"],
    ["884 CE",           "Notker Balbulus 完成《讚美詩集》· Notker completes Liber hymnorum"],
    ["9th–11th c.",      "Trope 創作在修道院中興盛 · Trope composition flourishes"],
    ["late 10th c.",     "Quem queritis in presepe（聖誕禮儀劇）· Christmas liturgical drama"],
    ["1014 CE",          "Credo 信經最後加入彌撒 · Credo added to Mass"],
    ["ca. 1020–50",      "Wipo: Victimae paschali laudes（復活節繼抒詠）"],
    ["ca. 1025",         "聖誕彌撒 trope 抄入里摩日附近修道院手稿"],
    ["1054 CE",          "東西教會正式分裂 · Final split of Roman and Byzantine churches"],
    ["1066 CE",          "諾曼征服英格蘭 · Battle of Hastings"],
    ["1095–99 CE",       "第一次十字軍東征 · First Crusade"],
    ["1098–1179",        "Hildegard of Bingen 生卒"],
    ["1098–1146",        "Adam of St. Victor 活躍於巴黎 · Adam of St. Victor active in Paris"],
    ["ca. 1151",         "Hildegard, Ordo virtutum"],
    ["ca. 1210–50",      "Thomas of Celano, Dies irae"],
    ["1562–63",          "特倫多會議禁用大多數 trope 與 sequence · Council of Trent bans tropes, most sequences"],
  ];

  s.addShape(pres.ShapeType.rect, { x: 2.6, y: 0.85, w: 0.05, h: 4.55, fill: { color: C.gold } });

  events.forEach(([date, event], i) => {
    const y = 0.85 + i * 0.285;
    s.addShape(pres.ShapeType.ellipse, { x: 2.47, y: y + 0.04, w: 0.26, h: 0.26, fill: { color: C.gold } });
    s.addText(date, { x: 0.1, y, w: 2.28, h: 0.28, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", align: "right", margin: 0 });
    s.addText(event, { x: 2.92, y, w: 6.8, h: 0.28, fontSize: 14, color: C.lightText, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 15 · Continuing Presence ──────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("聖詠的延續", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.5, fontSize: 28, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addText("The Continuing Presence of Chant · From Middle Ages to Today", {
    x: 0.4, y: 0.68, w: 9.2, h: 0.35, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.05, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const eras = [
    ["■", "中世紀—16 世紀", "Middle Ages – 16th c.", "Léonin、Du Fay、Ockeghem、Josquin、Palestrina 等作曲家主要工作是演唱和指揮聖詠；多聲部音樂以聖詠為基礎發展\nMajor composers spent most of their time singing and directing chant; polyphony built on chant foundation"],
    ["■", "宗教改革—19 世紀", "Reformation – 19th c.", "聖詠旋律改編為新教眾讚歌（chorale）與聖公會聖歌；在天主教地區持續使用\nChant melodies adapted for Protestant chorales and Anglican hymns; still used in Catholic areas"],
    ["■", "梵二會議 1962–65", "Vatican II", "允許以地方語言舉行彌撒；拉丁聖詠不再是必須——從定期禮拜中幾近消失\nPermitted vernacular Mass; Latin chant no longer required, virtually disappeared from regular services"],
    ["*", "當代復興", "Contemporary revival", "1993 Silos 修士《Chant》專輯歐洲暢銷冠軍；2008 Heiligenkreuz《Chant》全球熱銷；電玩《Halo》等使用聖詠風格\n1993 Chant album bestseller; 2008 Heiligenkreuz platinum; Halo uses chant style"],
  ];

  eras.forEach(([icon, title, subtitle, desc], i) => {
    const y = 1.18 + i * 1.03;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.92, fill: { color: C.wine }, rounding: true });
    s.addShape(pres.ShapeType.rect, { x: 0.4, y: y + 0.08, w: 0.75, h: 0.76, fill: { color: C.gold }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.14, w: 0.75, h: 0.65, fontSize: 22, align: "center", margin: 0 });
    s.addText(title, { x: 1.3, y: y + 0.07, w: 3.5, h: 0.3, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(subtitle, { x: 1.3, y: y + 0.34, w: 3.5, h: 0.25, fontSize: 14, color: C.sand, italic: true, fontFace: "Calibri", valign: "top", margin: 0 });
    s.addShape(pres.ShapeType.rect, { x: 4.85, y: y + 0.12, w: 0.03, h: 0.68, fill: { color: C.gold } });
    s.addText(desc, { x: 4.95, y: y + 0.08, w: 4.7, h: 0.8, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 16 · Chapter Summary ──────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("本章重點回顧 Chapter Summary", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 0.76, w: 7, h: 0.04, fill: { color: C.gold } });

  const points = [
    ["■", "羅馬禮儀的雙重目的：對上帝獻頌讚，對信徒傳教義；彌撒為中心，日課為日常\nRoman liturgy's dual aim: praise to God + instruction for worshippers; Mass central, Office daily"],
    ["■", "彌撒分為 Proper（隨日變動）與 Ordinary（固定不變）；音樂高潮為 Gradual、Alleluia、Offertory\nMass = Proper (varies) + Ordinary (fixed); musical peaks: Gradual, Alleluia, Offertory"],
    ["■", "聖詠依演唱方式分 responsorial/antiphonal/direct，依歌詞設定分 syllabic/neumatic/melismatic\nChants classified by performance (responsorial/antiphonal/direct) and text setting (syllabic/neumatic/melismatic)"],
    ["■", "Trope（增飾）與 Sequence（繼抒詠）是 9–12 世紀創作的出口，在授權聖詠的「邊緣」發揮創意\nTropes and sequences gave 9–12 c. musicians creative outlets in the margins of the authorized repertory"],
    ["■", "禮儀劇由對話型 trope 發展而來（如 Quem queritis in sepulchro）；由神職人員在教堂內演出\nLiturgical drama grew from dialogue tropes (Quem queritis in sepulchro); performed in church by clergy"],
    ["■", "Hildegard of Bingen 是中世紀留下最多聖詠的作曲家，也是第一位留下完整音樂劇（Ordo virtutum）的女性\nHildegard left more chants than any medieval composer; her Ordo virtutum is the earliest surviving non-liturgical music drama"],
  ];

  points.forEach(([icon, text], i) => {
    const y = 0.9 + i * 0.77;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.66, fill: { color: "3A2015" }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.08, w: 0.55, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(text, { x: 1.05, y: y + 0.05, w: 8.5, h: 0.58, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 17 · Further Reading ──────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("延伸閱讀與補充教材\nFurther Reading & Supplementary Resources", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.72,
    fontSize: 22, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.95, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Listening (full-width, 6 items one per line)
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.05, w: 9.2, h: 2.05, fill: { color: C.wine }, rounding: true });
  s.addText("■ 聆聽 YouTube", {
    x: 0.55, y: 1.1, w: 8.9, h: 0.3, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", valign: "top", margin: 0,
  });
  const listening = [
    "Introit: Puer natus est (NAWM 3a)  youtu.be/f4iR0qxeiLg",
    "Gradual: Viderunt omnes (NAWM 3d)  youtu.be/uvC0NGmpuFY",
    "Alleluia: Dies sanctificatus (NAWM 3e)  youtu.be/pv8lFyLgjN4",
    "Wipo — Victimae paschali laudes  youtu.be/v9HraDioS_Q",
    "Hildegard — O virga ac diadema (Sequentia)  youtu.be/CO7IqAZ4BtI",
    "Dies irae (Thomas of Celano)  youtu.be/Vj7R9LeiJPE",
  ];
  s.addText(listening.map((l, i) => ({
    text: l, options: { bullet: true, breakLine: i < listening.length - 1, fontSize: 14, color: C.cream, fontFace: "Calibri" },
  })), { x: 0.55, y: 1.4, w: 8.9, h: 1.65, valign: "top", margin: 0 });

  // Reading (2-column)
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 3.2, w: 9.2, h: 1.2, fill: { color: C.rust }, rounding: true });
  s.addText("■ 閱讀 Read", {
    x: 0.55, y: 3.24, w: 8.9, h: 0.28, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", valign: "top", margin: 0,
  });
  const readingLeft = [
    "Liber usualis（Solesmes 修院通用聖詠集）",
    "Rule of St. Benedict — CCEL 線上全文",
    "Wikipedia: Mass / Hildegard of Bingen",
  ];
  const readingRight = [
    "Crocker — The Early Medieval Sequence",
    "Scivias — Oliver Davies 英譯本",
    "NAWM 3–7（Grout 樂譜集對應樂曲）",
  ];
  s.addText(readingLeft.map((r, i) => ({
    text: r, options: { bullet: true, breakLine: i < readingLeft.length - 1, fontSize: 14, color: C.cream, fontFace: "Calibri" },
  })), { x: 0.55, y: 3.54, w: 4.4, h: 0.82, valign: "top", margin: 0 });
  s.addText(readingRight.map((r, i) => ({
    text: r, options: { bullet: true, breakLine: i < readingRight.length - 1, fontSize: 14, color: C.cream, fontFace: "Calibri" },
  })), { x: 5.05, y: 3.54, w: 4.4, h: 0.82, valign: "top", margin: 0 });

  // Key terms
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 4.5, w: 9.2, h: 0.95, fill: { color: C.midBrown }, rounding: true });
  s.addText("■ 本章關鍵術語 Key Terms", {
    x: 0.55, y: 4.54, w: 8.9, h: 0.26, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", valign: "top", margin: 0,
  });
  const terms = "Liturgy 禮儀 · Mass 彌撒 · Office 聖務日課 · Proper / Ordinary · Introit · Kyrie · Gloria · Gradual · Alleluia · Tract · Offertory · Communion · Credo · Sanctus · Agnus Dei · Antiphon · Psalm tone · Jubilus · Trope · Sequence · Dies irae";
  s.addText(terms, {
    x: 0.55, y: 4.82, w: 8.9, h: 0.6, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0,
  });
}

// ── Generate file ─────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "Ch03_Roman_Liturgy.pptx" })
  .then(() => console.log("■ Ch03_Roman_Liturgy.pptx created successfully"))
  .catch(err => console.error("■ Error:", err));
