const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 4: Song and Dance Music to 1300";
pres.author = "A History of Western Music, 10th ed.";

// ── Color palette (same parchment/ancient theme as Ch01/Ch02/Ch03) ───────────
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
    fontSize: 11, color: C.sand, charSpacing: 3, align: "center", fontFace: "Georgia",
  });

  s.addText("CHAPTER 4", {
    x: 0.5, y: 0.9, w: 9, h: 0.55,
    fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });

  s.addText("SONG AND DANCE\nMUSIC TO 1300", {
    x: 0.3, y: 1.4, w: 9.4, h: 2.0,
    fontSize: 42, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
    paraSpaceAfter: 0,
  });

  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.55, w: 5, h: 0.04, fill: { color: C.gold } });

  s.addText("Troubadours · Trouvères · Minnesinger · Cantigas · Estampie", {
    x: 0.4, y: 3.7, w: 9.2, h: 0.4,
    fontSize: 14, color: C.sand, italic: true, align: "center", fontFace: "Georgia",
  });

  s.addText("Textbook pp. 63–79", {
    x: 0.5, y: 4.8, w: 9, h: 0.3,
    fontSize: 11, color: C.gold, align: "center", fontFace: "Calibri",
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
    ["🏰", "歐洲社會 800–1300 European Society", "查理曼加冕、封建領主、城市興起、大學出現、識字率提升"],
    ["📜", "拉丁與方言歌曲 Latin & Vernacular Song", "Versus、Conductus、Goliard 學者歌；方言史詩 Chanson de geste"],
    ["🎭", "遊唱詩人 Jongleurs & Minstrels", "從被視為賤民到組織行會（巴黎 1321）——音樂職業化的起點"],
    ["💘", "Troubadours & Trouvères", "南法 Occitan 與北法古法語的抒情詩人；「fine amour」典範"],
    ["🌍", "他地之歌 Songs in Other Lands", "Minnesinger（德）、Lauda（義）、Cantigas de Santa María（西）"],
    ["💃", "中世紀樂器與舞曲 Instruments & Dance", "Vielle、Psaltery、風笛、管風琴；Carole 圓舞、Estampie"],
  ];

  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.wine }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 11, color: C.midBrown, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 3 · European Society 800–1300 ──────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("歐洲社會 800–1300", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("European Society · The Context for Medieval Song", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 15, color: C.sand, fontFace: "Georgia", align: "center", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Three empires box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 1.55, fill: { color: "3A2015" }, rounding: true });
  s.addText("🌍 三個繼承羅馬的帝國 Three Successors to Rome", {
    x: 0.45, y: 1.38, w: 9.1, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("① Byzantine Empire 拜占庭：保存希臘羅馬科學與文化；多數古希臘著作因拜占庭抄寫員而倖存\n② Arab world 阿拉伯世界：最強盛的一支；延伸希臘哲學與科學；在醫學、化學、數學上貢獻卓著\n③ Western Europe 西歐：最弱最窮；800 年查理曼加冕為羅馬皇帝，宣告與拜占庭獨立\n    Western Europe was weakest; Charlemagne's coronation in 800 asserted independence from Byzantium", {
    x: 0.5, y: 1.72, w: 9.0, h: 1.1, fontSize: 10.5, color: C.sand, fontFace: "Calibri",
  });

  // Social developments box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 2.95, w: 9.4, h: 2.3, fill: { color: "3A2015" }, rounding: true });
  s.addText("📊 社會與文化變遷 Social & Cultural Developments", {
    x: 0.45, y: 3.03, w: 9.1, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  const dev = [
    "• 人口 1000–1300 年間增長三倍；農業技術進步、水車與風車普及",
    "  Population tripled 1000–1300; agricultural advances, water mills and windmills",
    "• 三階級社會：神職人員（祈禱）、貴族騎士（征戰）、農民商人（勞動）",
    "  Three estates: clergy who prayed · nobility who fought · commoners who worked",
    "• 城市興起：巴黎約 20 萬人、倫敦 7 萬、威尼斯/米蘭/佛羅倫斯各約 10 萬",
    "• 行會（guild）保護工匠；學徒制度；音樂家也開始組織行會",
    "• 大學成立：波隆那、巴黎、牛津（12 世紀起）· 識字率大幅提升",
    "• 從修道院到主教座堂學校；教會重心由鄉村移至城市",
  ];
  s.addText(dev.map((t, i) => ({
    text: t, options: { bullet: false, breakLine: i < dev.length - 1, fontSize: 10, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1 },
  })), { x: 0.55, y: 3.4, w: 9.0, h: 1.85 });
}

// ── SLIDE 4 · Latin and Vernacular Song ──────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("拉丁與方言歌曲 Latin & Vernacular Song", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 24, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addText("教會之外、修道院外的中世紀歌唱傳統", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 12, color: C.midBrown, italic: true, fontFace: "Calibri", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.04, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Latin song box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.wine }, rounding: true });
  s.addText("📜 拉丁歌曲 Latin Song", {
    x: 0.45, y: 1.28, w: 4.3, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.55, y: 1.66, w: 4.1, h: 0.02, fill: { color: C.gold } });

  s.addText("Versus（詩句）:\n• 11–13 世紀亞奎丹地區 · 多半神聖 · 押韻、規律重音\n• 主題：耶穌誕生、聖母、道成肉身\n• 影響了 troubadour 與亞奎丹複音\n\nConductus（導引歌）:\n• 12 世紀北法 · 嚴肅的拉丁歌\n• 神聖或世俗皆可 · 新創旋律不以聖詠為基\n\nGoliard 歌曲:\n• 10–13 世紀 · 流浪學者與教士\n• 主題：愛情、春天、飲酒、諷刺\n• 作者常為受人尊敬的教師與宮廷人", {
    x: 0.5, y: 1.78, w: 4.2, h: 3.4, fontSize: 10, color: C.cream, fontFace: "Calibri",
  });

  // Vernacular song box
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.midBrown }, rounding: true });
  s.addText("🗣 方言歌曲 Vernacular Song", {
    x: 5.25, y: 1.28, w: 4.3, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 5.35, y: 1.66, w: 4.1, h: 0.02, fill: { color: C.gold } });

  s.addText("大多失傳：因中世紀 90% 人口為不識字的鄉民\nMost are lost — 90% of the population was nonliterate\n\n倖存的主要類型：\n\n• Epic 史詩：以簡單旋律公式吟唱\n  Chanson de geste（「功業之歌」）\n  最著名：《羅蘭之歌》(ca. 1100)\n  其他：Beowulf（古英文）、The Nibelungs（德語）、Norse Eddas\n\n• 市井街頭叫賣、民歌、搖籃曲——僅因被複音音樂引用而留存（「如琥珀中的蒼蠅」）\n  Street cries preserved like flies in amber", {
    x: 5.3, y: 1.78, w: 4.2, h: 3.4, fontSize: 10, color: C.cream, fontFace: "Calibri",
  });
}

// ── SLIDE 5 · Minstrels & Professional Musicians ─────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("職業音樂家", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("Bards · Jongleurs · Minstrels · The Rise of Musicians' Guilds", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.35, fontSize: 14, color: C.sand, italic: true, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.08, w: 7, h: 0.04, fill: { color: C.gold } });

  const cats = [
    ["🎤 Bards 遊吟詩人", "Celtic lands", "在宴會上演唱史詩，以豎琴、費德爾或其他弦樂器自伴\nSang epics at banquets, accompanying themselves on harp or fiddle"],
    ["🤹 Jongleurs 雜耍藝人", "Itinerant · Lower-class", "來自拉丁文「juggler」· 流浪賣藝：雜耍、說書、演奏樂器\nItinerant lower-class performers — tricks, stories, instruments"],
    ["🎺 Minstrels 宮廷樂師", "From 13th c. · Specialized", "來自拉丁 minister（「服侍者」）· 受雇於宮廷或城市 · 出身背景多元\nFrom Latin minister; hired by courts or cities; varied backgrounds"],
    ["⚖ Guilds 音樂家行會", "From 12th c.", "提供法律保護、界定城市內的演出專屬權、規範行為準則\nProtected legal rights; defined performance exclusivity; regulated conduct"],
  ];

  cats.forEach(([name, tag, desc], i) => {
    const y = 1.25 + i * 0.85;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.75, fill: { color: "3A2015" }, rounding: true });
    s.addText(name, { x: 0.45, y: y + 0.08, w: 2.8, h: 0.35, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(tag, { x: 0.45, y: y + 0.42, w: 2.8, h: 0.3, fontSize: 9.5, color: C.sand, italic: true, fontFace: "Calibri", margin: 0 });
    s.addShape(pres.ShapeType.rect, { x: 3.3, y: y + 0.1, w: 0.03, h: 0.55, fill: { color: C.gold } });
    s.addText(desc, { x: 3.4, y: y + 0.08, w: 6.2, h: 0.65, fontSize: 10, color: C.cream, fontFace: "Calibri", margin: 0 });
  });

  s.addText("🔑 里程碑：1321 年巴黎的 Confrérie de St.-Julien des Menestriers 成立，37 位男女樂師共同簽署章程", {
    x: 0.4, y: 4.85, w: 9.2, h: 0.3, fontSize: 10.5, color: C.gold, italic: true, fontFace: "Calibri", align: "center",
  });
  s.addText("Milestone: In 1321 Paris musicians formed the Confrérie de St.-Julien des Menestriers — 37 men and women signatories", {
    x: 0.4, y: 5.1, w: 9.2, h: 0.3, fontSize: 10, color: C.sand, italic: true, fontFace: "Calibri", align: "center",
  });
}

// ── SLIDE 6 · Troubadours and Trouvères Overview ─────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("遊唱詩人 Troubadours & Trouvères", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 24, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addText("The Fountainhead of Western Vernacular Poetry", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 12, color: C.midBrown, italic: true, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.04, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Troubadour box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.wine }, rounding: true });
  s.addText("🌅 Troubadours", {
    x: 0.45, y: 1.28, w: 4.3, h: 0.35, fontSize: 15, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("南法 · 12 世紀起 · Occitan（langue d'oc）", {
    x: 0.45, y: 1.65, w: 4.3, h: 0.3, fontSize: 10, color: C.sand, italic: true, fontFace: "Calibri",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.55, y: 2.0, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("• Trobador（陰性 trobairitz）\n• Trobar = 「創作歌曲」→「找到」「發明」\n• 存 ca. 2,600 首詩 · 約 1/10 有旋律\n\n重要人物:\n◆ Guillaume IX（亞奎丹公爵，1071–1126）\n   現存最早留名的 troubadour\n◆ Bernart de Ventadorn (ca. 1130–1200)\n   最具影響力之一 · 城堡僕役之子\n◆ Comtessa de Dia（12 世紀末 / 13 世紀初）\n   唯一有音樂留存的女性遊唱詩人 trobairitz", {
    x: 0.5, y: 2.08, w: 4.2, h: 3.1, fontSize: 9.5, color: C.cream, fontFace: "Calibri",
  });

  // Trouvère box
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.rust }, rounding: true });
  s.addText("🌇 Trouvères", {
    x: 5.25, y: 1.28, w: 4.3, h: 0.35, fontSize: 15, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("北法 · 12 世紀末起 · 古法語（langue d'oïl）", {
    x: 5.25, y: 1.65, w: 4.3, h: 0.3, fontSize: 10, color: C.sand, italic: true, fontFace: "Calibri",
  });
  s.addShape(pres.ShapeType.rect, { x: 5.35, y: 2.0, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("• Trover 也意為「創作歌曲」\n• 存 ca. 2,100 首詩 · 2/3 有旋律（遠高於 troubadour）\n• 由南法北傳 · Bernart de Ventadorn 功不可沒\n\n重要人物:\n◆ Eleanor of Aquitaine (1122–1204)\n   亞奎丹女公爵、法國與英國王后 · 贊助 Bernart\n◆ Richard I（獅心王，1157–1199）\n   Eleanor 之子，本人就是 trouvère，以法文作詞\n◆ Adam de la Halle (ca. 1240–1288)\n   第一位完整作品被集結手稿的方言詩人\n   代表作：Jeu de Robin et de Marion", {
    x: 5.3, y: 2.08, w: 4.2, h: 3.1, fontSize: 9.5, color: C.cream, fontFace: "Calibri",
  });

  s.addText("歌曲收錄於手稿集「Chansonnier」（歌曲集）· 現存大多為 13 世紀中後期抄本", {
    x: 0.4, y: 5.25, w: 9.2, h: 0.22, fontSize: 10, color: C.wine, italic: true, fontFace: "Calibri", align: "center",
  });
}

// ── SLIDE 7 · Fin' Amors — The Poetry ────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("詩歌主題 · Fin' Amors", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("\"Refined Love\" — The Central Theme of Troubadour Poetry", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, italic: true, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // What is fin' amors
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 1.75, fill: { color: "3A2015" }, rounding: true });
  s.addText("💘 何謂 Fin' amors（精鍊之愛）", {
    x: 0.45, y: 1.38, w: 9.1, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("• 非雙方對等的愛，而是形式化、理想化的愛——情人藉之自我精鍊\n  Not mutual; formal, idealized love through which the lover is refined\n• 對象通常是另一貴族之妻；從遠方、帶著謹慎、尊敬、謙卑地愛慕\n  Object usually another man's wife; adored from a distance with discretion, respect, humility\n• 愛慕語言常與對聖母的崇拜相近——女士被塑造成高不可攀、無法回應的理想\n  The language borders on devotional songs to the Virgin Mary\n• 「courtly love」一詞為 19 世紀後創；不是事實而是藝術形式，向男性同儕展現教養", {
    x: 0.5, y: 1.72, w: 9.0, h: 1.3, fontSize: 10.5, color: C.sand, fontFace: "Calibri",
  });

  // Genres box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.17, w: 9.4, h: 2.05, fill: { color: "3A2015" }, rounding: true });
  s.addText("📚 詩歌類型 Poetic Genres", {
    x: 0.45, y: 3.25, w: 9.1, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });

  const genres = [
    ["Canso 情歌", "Love song · 最主要的類型"],
    ["Alba 黎明歌", "Dawn song · 情人於拂曉被迫分別"],
    ["Balada 舞歌", "Dance song · 有迴旋副歌（refrain）"],
    ["Planh 輓歌", "Lament · 悼念逝者"],
    ["Tenso 辯論歌", "Debate song · 兩人就某一主題對話"],
    ["Pastourelle 牧歌", "牧羊女與追求她的騎士"],
  ];
  genres.forEach(([name, desc], i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + col * 4.6;
    const y = 3.66 + row * 0.48;
    s.addText(name, { x, y, w: 1.8, h: 0.35, fontSize: 11, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(desc, { x: x + 1.85, y, w: 2.6, h: 0.35, fontSize: 9.5, color: C.sand, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 8 · Melodies & AAB Form ────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("旋律與 AAB 形式", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 26, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addText("Melodies & the AAB Form · Forms at a Glance", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 12, color: C.midBrown, italic: true, fontFace: "Calibri", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.04, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Melody characteristics
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.wine }, rounding: true });
  s.addText("🎵 旋律特徵 Melody", {
    x: 0.45, y: 1.28, w: 4.3, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.55, y: 1.66, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("• 分節歌式（strophic）\n  每段歌詞共用同一旋律\n\n• 多半音節式（mostly syllabic）\n  偶有兩三音的 neume\n\n• 音域窄，少超過九度\n  Range seldom over a ninth\n\n• 主要級進，偶有三度跳進\n\n• 樂句呈「弧形」：升至高點後下行至終止\n  Phrases arch-shaped\n\n• 多符合教會調式——mode 1 與 mode 7 最常見", {
    x: 0.5, y: 1.78, w: 4.2, h: 3.4, fontSize: 10, color: C.cream, fontFace: "Calibri",
  });

  // AAB form
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.midBrown }, rounding: true });
  s.addText("📐 AAB 形式", {
    x: 5.25, y: 1.28, w: 4.3, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 5.35, y: 1.66, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("• 一段樂句群唱兩次（第二次填新詞）+ 對比段\n  Section sung twice (new text) + contrasting section\n\n• A 段與 B 段常以相同樂句結束\n  → 產生「音樂韻腳」(musical rhyme)\n\n• 德國 Minnesinger 稱為 Bar Form\n  A 段 = Stollen（下唇）\n  B 段 = Abgesang（末唱）\n\n• 四首 NAWM 範例:\n  A chantar（Comtessa de Dia）\n  Palästinalied（Walther）\n  Robins m'aime（Adam de la Halle）\n  Non sofre Santa María（Cantiga 159）", {
    x: 5.3, y: 1.78, w: 4.2, h: 3.4, fontSize: 9.5, color: C.cream, fontFace: "Calibri",
  });

  s.addText("🔑 節奏記譜：以聖詠音高符號寫成，不指示節奏——演出時可自由（愛情歌）或按韻律節拍（舞歌）", {
    x: 0.4, y: 5.26, w: 9.2, h: 0.22, fontSize: 10, color: C.rust, italic: true, fontFace: "Calibri", align: "center",
  });
}

// ── SLIDE 9 · Songs in Other Lands ───────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("他地之歌 Songs in Other Lands", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("The Troubadour Tradition Spreads Across Europe", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.35, fontSize: 14, color: C.sand, italic: true, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.08, w: 7, h: 0.04, fill: { color: C.gold } });

  const lands = [
    ["🏴󠁧󠁢󠁥󠁮󠁧󠁿 English Song", "England", "1066 諾曼征服後 · 法語為王室語言 · Eleanor 之子 Richard I 獅心王本人為 trouvère，以法文創作\nFew Middle English melodies survive"],
    ["🦅 Minnesinger", "Germany, 12–14 c.", "以中古高地德語創作 · Minne = 愛 · 更重靈性、忠誠與義務 · AAB 形式稱 Bar Form\nWalther von der Vogelweide: Palästinalied（十字軍歌）"],
    ["✝ Lauda 讚美歌", "Italy, 13 c.+", "義大利神聖單旋律歌 · 在城市（非宮廷）創作 · 由宗教苦行遊行與兄弟會演唱\nBlends chant, hymn, and troubadour influences"],
    ["🎼 Cantigas de Santa María", "Castile, ca. 1270–90", "400+ 首 Galician-Portuguese 歌頌聖母的歌 · 卡斯提爾王 Alfonso el Sabio（智者）監督編纂 · 每第 10 首為讚美聖母\nAlmost all have refrains; suitable for dancing"],
  ];

  lands.forEach(([name, tag, desc], i) => {
    const y = 1.22 + i * 0.95;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.85, fill: { color: "3A2015" }, rounding: true });
    s.addText(name, { x: 0.45, y: y + 0.08, w: 3.1, h: 0.35, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(tag, { x: 0.45, y: y + 0.42, w: 3.1, h: 0.35, fontSize: 9.5, color: C.sand, italic: true, fontFace: "Calibri", margin: 0 });
    s.addShape(pres.ShapeType.rect, { x: 3.6, y: y + 0.1, w: 0.03, h: 0.65, fill: { color: C.gold } });
    s.addText(desc, { x: 3.7, y: y + 0.08, w: 5.95, h: 0.75, fontSize: 9.5, color: C.cream, fontFace: "Calibri", margin: 0 });
  });

  s.addText("1208 起教皇 Innocent III 發動反 Albigensian 十字軍——導致南法貴族崩潰，troubadours 流散至義大利等地", {
    x: 0.4, y: 5.2, w: 9.2, h: 0.25, fontSize: 10, color: C.gold, fontFace: "Calibri", align: "center",
  });
}

// ── SLIDE 10 · Medieval Instruments ──────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("中世紀樂器 Medieval Instruments", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 24, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addText("弦、木管、銅管、打擊、管風琴——中世紀樂師已有豐富的音色調色盤", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 11, color: C.midBrown, italic: true, fontFace: "Calibri", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.0, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const groups = [
    ["🎻 弦樂器 Strings", C.wine, [
      "Vielle 費德爾：五弦、四度與五度調音，可由空弦作持續音——中世紀主要拉弦樂器，小提琴的前身",
      "Hurdy-gurdy 搖弦琴：三弦 vielle，旋轉輪擦弦發聲；另一弦作持續音；類似風笛音響",
      "Harp 豎琴：此類型可能源自不列顛群島",
      "Psaltery 詩琴：撥弦樂器，彈撥木框上的弦——大鍵琴與鋼琴的遠祖",
      "Lute 魯特琴（後來由阿拉伯傳入歐洲）",
    ]],
    ["🌬 管樂 · 打擊 Winds & Percussion", C.rust, [
      "Transverse flute 橫笛：木製或象牙，無按鍵",
      "Shawm 蕭姆管：雙簧樂器，雙簧管前身",
      "Trumpet 號角：直管無活塞，僅能吹奏泛音列",
      "Pipe and tabor：左手持高音哨笛，右手敲小鼓——單人可演奏舞曲",
      "Bagpipe 風笛：遍布歐洲的民間樂器",
      "Bells 鐘 · Organs 管風琴（portative 可攜式 / positive 小型桌上型）",
    ]],
  ];

  groups.forEach(([title, color, items], i) => {
    const x = 0.3 + i * 4.8;
    s.addShape(pres.ShapeType.rect, { x, y: 1.15, w: 4.6, h: 4.05, fill: { color }, rounding: true });
    s.addText(title, { x: x + 0.15, y: 1.23, w: 4.3, h: 0.35, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
    s.addShape(pres.ShapeType.rect, { x: x + 0.25, y: 1.62, w: 4.1, h: 0.02, fill: { color: C.gold } });
    s.addText(items.map((t, j) => ({
      text: t, options: { bullet: true, breakLine: j < items.length - 1, fontSize: 9.5, color: C.cream, fontFace: "Calibri", paraSpaceAfter: 3 },
    })), { x: x + 0.2, y: 1.72, w: 4.3, h: 3.4 });
  });
}

// ── SLIDE 11 · Dance Music ───────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("舞曲 Dance Music", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("The Earliest Notated Instrumental Music", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, italic: true, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Carole
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 1.7, fill: { color: "3A2015" }, rounding: true });
  s.addText("💃 Carole 圓舞", {
    x: 0.45, y: 1.38, w: 9.1, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("12–14 世紀法國最流行的社交舞 · 圓圈舞蹈 · 通常由其中一位或多位舞者歌唱伴奏\nThe most popular social dance in France, 12th–14th c.; circle dance usually led by singing dancers\n• 可加入樂師伴奏（如《玫瑰傳奇》中所描述）\n• 儘管流行，現存旋律僅約 24 首——多靠口傳心授，極少寫下", {
    x: 0.5, y: 1.72, w: 9.0, h: 1.25, fontSize: 10.5, color: C.sand, fontFace: "Calibri",
  });

  // Estampie
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.1, w: 9.4, h: 2.15, fill: { color: "3A2015" }, rounding: true });
  s.addText("🎺 Estampie 埃斯坦皮舞曲", {
    x: 0.45, y: 3.18, w: 9.1, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("現存最常見的中世紀器樂舞曲形式 · 13–14 世紀 · 約 50 首現存（多為單音，少數有鍵盤複音設定）\nThe most common surviving medieval dance; ca. 50 extant tunes (mostly monophonic)\n\n• 結構：數個樂段，每段演奏兩次，使用兩種不同結尾：\n  ① Ouvert 開放結尾（半終止）\n  ② Clos 封閉結尾（完全終止）\n  Each section played twice: open (ouvert) + closed (clos) endings\n• 法國 estampie 全部為三拍子 · 義大利 istampita 用二拍或複合拍，段落較長\n• 範例：Le manuscrit du roi（國王手稿）中的八首「皇家 estampies」（NAWM 13 為第四首）", {
    x: 0.5, y: 3.52, w: 9.0, h: 1.7, fontSize: 10, color: C.sand, fontFace: "Calibri",
  });
}

// ── SLIDE 12 · Key Figures ───────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("關鍵人物 Key Figures", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 26, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const figures = [
    ["👑", "Guillaume IX", "1071–1126", "亞奎丹公爵、普瓦提埃伯爵；現存最早留名的 troubadour。貴族出身顯示這一藝術深植於上流社會。"],
    ["🎵", "Bernart de Ventadorn", "?ca. 1130–?ca. 1200", "最具影響力的 troubadour 之一；Ventadorn 城堡麵包師/僕役之子；後為 Eleanor of Aquitaine 服務，把 troubadour 傳統帶入北法，催生 trouvère 運動。代表作《Can vei la lauzeta mover》"],
    ["👸", "Comtessa de Dia", "fl. late 12th – early 13th c.", "極少數留下音樂作品的 trobairitz（女性遊唱詩人）。《A chantar》是唯一有旋律保存下來的 trobairitz 歌曲（AAB 形式、mode 1）"],
    ["🛡", "Walther von der Vogelweide", "?ca. 1170–?ca. 1230", "最著名的 Minnesinger；以宮廷詩歌與政治詩聞名；代表作《Palästinalied》（十字軍歌）仍被當代「中世紀搖滾」樂團改編演奏"],
    ["🎭", "Adam de la Halle", "ca. 1240–?1288", "Arras 的 trouvère；第一位完整作品被集結為手稿的方言詩人；《Jeu de Robin et de Marion》（ca. 1284）是最早的世俗音樂劇之一；《Robins m'aime》為 rondeau 形式"],
    ["📚", "Alfonso X el Sabio", "King of Castile 1252–1284", "卡斯提爾與萊昂國王「智者阿方索」；監督編纂《Cantigas de Santa María》——400+ 首讚美聖母的歌，收錄於四部華麗的插圖手稿中"],
  ];

  figures.forEach(([icon, name, date, desc], i) => {
    const y = 0.9 + i * 0.76;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.65, fill: { color: C.wine }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.1, w: 0.55, h: 0.45, fontSize: 18, align: "center", margin: 0 });
    s.addText(name, { x: 1.0, y: y + 0.05, w: 3.1, h: 0.3, fontSize: 11, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(date, { x: 1.0, y: y + 0.32, w: 3.1, h: 0.28, fontSize: 8.5, color: C.sand, italic: true, fontFace: "Calibri", margin: 0 });
    s.addShape(pres.ShapeType.rect, { x: 4.15, y: y + 0.1, w: 0.025, h: 0.45, fill: { color: C.gold } });
    s.addText(desc, { x: 4.25, y: y + 0.04, w: 5.4, h: 0.58, fontSize: 8.5, color: C.cream, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 13 · Timeline ──────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("歷史時間軸 Timeline", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });

  const events = [
    ["8th c.",          "Beowulf（古英文史詩）· Beowulf (Old English epic)"],
    ["800 CE",          "查理曼加冕為羅馬皇帝 · Charlemagne crowned emperor in Rome"],
    ["834 CE",          "路易一世逝世，查理曼帝國分裂 · Louis the Pious dies, empire divided"],
    ["1000–1300",       "歐洲人口成長三倍 · European population triples"],
    ["ca. 1050",        "風車與水車普及 · Windmills and water mills spread"],
    ["1066 CE",         "諾曼征服英格蘭 · Normans conquer England"],
    ["1071–1126",       "Guillaume IX（最早留名 troubadour）"],
    ["1095–99",         "第一次十字軍東征 · First Crusade"],
    ["ca. 1100",        "《羅蘭之歌》· Song of Roland"],
    ["12th c.",         "聖母崇拜興盛；波隆那、巴黎、牛津大學成立"],
    ["ca. 1170s",       "Bernart de Ventadorn, Can vei la lauzeta mover"],
    ["ca. 1200",        "Comtessa de Dia, A chantar"],
    ["1208 CE",         "教皇發動反 Albigensian 十字軍 · Crusade against Albigensians"],
    ["1215 CE",         "《大憲章》簽署 · Magna Carta signed"],
    ["ca. 1228",        "Walther von der Vogelweide, Palästinalied"],
    ["ca. 1270–90",     "Cantigas de Santa María（阿方索十世監督編纂）"],
    ["ca. 1284",        "Adam de la Halle, Jeu de Robin et de Marion"],
    ["1321 CE",         "巴黎音樂家行會成立 · Paris minstrels' guild founded"],
  ];

  s.addShape(pres.ShapeType.rect, { x: 2.6, y: 0.85, w: 0.05, h: 4.55, fill: { color: C.gold } });

  events.forEach(([date, event], i) => {
    const y = 0.85 + i * 0.253;
    s.addShape(pres.ShapeType.ellipse, { x: 2.47, y: y + 0.04, w: 0.26, h: 0.26, fill: { color: C.gold } });
    s.addText(date, { x: 0.1, y, w: 2.28, h: 0.28, fontSize: 8.5, color: C.sand, fontFace: "Calibri", align: "right", margin: 0 });
    s.addText(event, { x: 2.92, y, w: 6.8, h: 0.28, fontSize: 8.5, color: C.lightText, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 14 · The Lover's Complaint (Continuing Presence) ───────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("情人的哀嘆 · 延續至今", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 24, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addText("The Lover's Complaint · From Troubadours to Today", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 12, color: C.midBrown, italic: true, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.05, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Common traits
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.18, w: 9.4, h: 1.9, fill: { color: C.wine }, rounding: true });
  s.addText("🎼 中世紀歌曲的共同特徵（至今仍是「典型西方歌曲」）", {
    x: 0.45, y: 1.26, w: 9.1, h: 0.35, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("• 分節歌式（strophic）· 自然音階（diatonic）· 主要音節式（syllabic）\n  Strophic · diatonic · primarily syllabic\n• 主要級進，音域約一個八度 · 清楚的音高中心\n  Mostly stepwise, range ~an octave, clear pitch center\n• 樂句短且大致等長，弧形起伏，回落至終止\n  Short musical phrases, roughly equal in size, arch-shaped\n• 常用副歌（refrain）· 主題多為「未能如願的純潔之愛」\n  Refrains common · theme of pure, unattainable love", {
    x: 0.5, y: 1.62, w: 9.0, h: 1.42, fontSize: 10.5, color: C.cream, fontFace: "Calibri",
  });

  // Revival box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.18, w: 9.4, h: 2.0, fill: { color: C.rust }, rounding: true });
  s.addText("🔄 後世的重新發現 Rediscovery", {
    x: 0.45, y: 3.26, w: 9.1, h: 0.35, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("• 18–19 世紀：民族文化復興 → 對中世紀的興趣重燃 · 收集整理 trouvère 歌曲配鋼琴伴奏出版\n  18th–19th c.: nationalism revives interest in the Middle Ages\n• 1872：Adam de la Halle 音樂作品被轉寫並出版\n  1872: Adam de la Halle's works transcribed and published\n• 20 世紀：更多中世紀曲目版本陸續問世\n• 近幾十年：中世紀世俗歌曲與器樂舞曲復興——音樂會與錄音\n  Recent decades: revival in concerts and recordings\n• 當代「中世紀搖滾」改編如 Walther von der Vogelweide《Palästinalied》（被多支歐洲樂團錄製）\n  Contemporary \"medieval rock\" arrangements", {
    x: 0.5, y: 3.62, w: 9.0, h: 1.55, fontSize: 10, color: C.cream, fontFace: "Calibri",
  });
}

// ── SLIDE 15 · Chapter Summary ───────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("本章重點回顧 Chapter Summary", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 0.76, w: 7, h: 0.04, fill: { color: C.gold } });

  const points = [
    ["🏰", "歐洲 800–1300 經歷人口倍增、城市崛起、大學成立；教會重心從修道院移至城市主教座堂\nEurope 800–1300: population growth, urban rise, universities; center shifts to urban cathedrals"],
    ["📜", "拉丁歌曲（versus、conductus、goliard）與方言史詩並存；多數世俗歌曲失傳因其口傳性質\nLatin songs coexist with vernacular epics; most secular songs are lost due to oral transmission"],
    ["🎤", "樂師從邊緣人物走向職業化：jongleurs → minstrels → guilds（1321 巴黎行會）\nMusicians professionalized: jongleurs → minstrels → guilds (Paris 1321)"],
    ["💘", "Troubadours（南法 Occitan）與 trouvères（北法古法語）以 fin' amors「精鍊之愛」為主題；AAB 形式盛行\nTroubadours and trouvères sing of fin' amors; AAB form predominates"],
    ["🌍", "傳統傳播：英國（法語為王室語言）、德國 Minnesinger（Bar Form）、義大利 lauda、西班牙 Cantigas\nTradition spreads: England, German Minnesinger, Italian lauda, Spanish Cantigas"],
    ["💃", "中世紀樂器類別齊備（弦木管銅管打擊管風琴）；Estampie 為現存最早的西方記譜器樂舞曲\nFull medieval instrument palette; Estampie = earliest notated Western instrumental dance"],
  ];

  points.forEach(([icon, text], i) => {
    const y = 0.9 + i * 0.77;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.66, fill: { color: "3A2015" }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.08, w: 0.55, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(text, { x: 1.05, y: y + 0.05, w: 8.5, h: 0.58, fontSize: 11, color: C.sand, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 16 · Further Reading ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("延伸閱讀與補充教材\nFurther Reading & Supplementary Resources", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.72,
    fontSize: 22, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.95, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Listening
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.05, w: 4.3, h: 2.5, fill: { color: C.wine }, rounding: true });
  s.addText("🎧 聆聽 YouTube", {
    x: 0.55, y: 1.12, w: 4.0, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  const listening = [
    "Bernart — Can vei la lauzeta mover (NAWM 8)  youtu.be/wjwDTer7R5Y",
    "Comtessa de Dia — A chantar (NAWM 9)  youtu.be/0aZcf5S9HGk",
    "Adam de la Halle — Robins m'aime (NAWM 10)  youtu.be/zNNm-wnfZ-U",
    "Walther — Palästinalied (NAWM 11)  youtu.be/zXoXcAToCBE",
    "Cantiga 159 — Non sofre Sta. María (NAWM 12)  youtu.be/RGK9bqWctKE",
    "La quarte estampie royal (NAWM 13)  youtu.be/rXqgEnrgtSg",
  ];
  s.addText(listening.map((l, i) => ({
    text: l, options: { bullet: true, breakLine: i < listening.length - 1, fontSize: 10, color: C.cream, fontFace: "Calibri", paraSpaceAfter: 3 },
  })), { x: 0.55, y: 1.56, w: 4.0, h: 1.85 });

  // Reading
  s.addShape(pres.ShapeType.rect, { x: 4.9, y: 1.05, w: 4.7, h: 2.5, fill: { color: C.rust }, rounding: true });
  s.addText("📖 閱讀 Read", {
    x: 5.05, y: 1.12, w: 4.4, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  const reading = [
    "Paden & Paden — Troubadour Poems from the South of France",
    "Samuel Rosenberg — Songs of the Troubadours and Trouvères",
    "John Stevens — Words and Music in the Middle Ages",
    "Christopher Page — The Owl and the Nightingale",
    "Wikipedia: Troubadour / Minnesinger / Cantigas de Santa María",
    "NAWM 8–13（Grout 樂譜集對應樂曲）",
  ];
  s.addText(reading.map((r, i) => ({
    text: r, options: { bullet: true, breakLine: i < reading.length - 1, fontSize: 10, color: C.cream, fontFace: "Calibri", paraSpaceAfter: 3 },
  })), { x: 5.05, y: 1.56, w: 4.4, h: 1.85 });

  // Key terms
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 3.65, w: 9.2, h: 1.7, fill: { color: C.midBrown }, rounding: true });
  s.addText("🔑 本章關鍵術語 Key Terms", {
    x: 0.6, y: 3.72, w: 8.8, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia",
  });
  const terms = "Versus · Conductus · Goliard · Chanson de geste · Jongleur · Minstrel · Guild · Troubadour / Trobairitz · Trouvère · Langue d'oc / Langue d'oïl · Chansonnier · Contrafactum · Fin' amors · Canso · Alba · Balada · Planh · Tenso · Pastourelle · Vida · AAB form · Bar Form · Stollen / Abgesang · Minnesinger · Minnelied · Lauda · Cantiga · Vielle · Hurdy-gurdy · Psaltery · Shawm · Pipe and tabor · Portative / Positive organ · Carole · Estampie · Ouvert / Clos · Istampita · Rondeau";
  s.addText(terms, {
    x: 0.6, y: 4.18, w: 8.8, h: 1.02, fontSize: 9.5, color: C.cream, fontFace: "Calibri",
  });
}

// ── Generate file ─────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "Ch04_Song_and_Dance.pptx" })
  .then(() => console.log("✅ Ch04_Song_and_Dance.pptx created successfully"))
  .catch(err => console.error("❌ Error:", err));
