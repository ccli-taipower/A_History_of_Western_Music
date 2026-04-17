const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 2: The Christian Church in the First Millennium";
pres.author = "A History of Western Music, 10th ed.";

// ── Color palette (same parchment/ancient theme as Ch01) ─────────────────────
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

// ── Master background helpers ─────────────────────────────────────────────────
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

// ── Reusable: accent bars ─────────────────────────────────────────────────────
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

  s.addText("CHAPTER 2", {
    x: 0.5, y: 0.9, w: 9, h: 0.55,
    fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });

  s.addText("THE CHRISTIAN CHURCH\nIN THE FIRST MILLENNIUM", {
    x: 0.3, y: 1.35, w: 9.4, h: 2.0,
    fontSize: 40, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
    paraSpaceAfter: 0,
  });

  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.45, w: 5, h: 0.04, fill: { color: C.gold } });

  s.addText("The Diffusion of Christianity · The Judaic Heritage · Chant · Notation", {
    x: 0.5, y: 3.6, w: 9, h: 0.4,
    fontSize: 15, color: C.sand, italic: true, align: "center", fontFace: "Georgia",
  });

  s.addText("Textbook pp. 20–41", {
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
    ["■", "基督教的傳播 The Diffusion of Christianity", "從耶穌到君士坦丁，基督教如何在羅馬帝國擴散"],
    ["■", "猶太傳統 The Judaic Heritage", "聖殿詩篇、吟誦、對早期基督教音樂的影響"],
    ["■", "早期教會音樂 Music in the Early Church", "感恩詩篇、教父思想、單聲部演唱傳統"],
    ["■", "教會分裂與聖詠方言 Divisions & Dialects of Chant", "拜占庭、格里高利、安布羅西等聖詠傳統"],
    ["■", "記譜法的發展 The Development of Notation", "紐姆符號 → 有線譜 → 阿雷佐的桂多"],
    ["■", "音樂理論與實踐 Music Theory and Practice", "波愛修斯、調式體系、唱名法（ut re mi fa sol la）"],
  ];

  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.wine }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 14, color: C.midBrown, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 3 · The Diffusion of Christianity ──────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("基督教的傳播", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("The Diffusion of Christianity", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 18, color: C.sand, fontFace: "Georgia", align: "center", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const facts = [
    ["■ 起源 Origin",        "拿撒勒耶穌（Jesus of Nazareth）—猶太人，羅馬帝國臣民；其教導催生基督教運動\nJesus of Nazareth, a Jew and subject of Rome; his teachings sparked Christianity"],
    ["◆ 傳播 Spread",       "聖彼得、聖保羅等使徒將信仰帶至近東、希臘、義大利各地\nApostles (St. Peter, St. Paul) spread the faith throughout the Near East, Greece, and Italy"],
    ["§ 迫害到合法 Legal",   "313 CE：君士坦丁一世頒布《米蘭勒令》，基督教合法化\nConstantine I issued the Edict of Milan (313 CE); Christianity legalized"],
    ["■ 國教 Official",     "392 CE：狄奧多西一世定基督教為羅馬帝國唯一官方宗教\nEmperor Theodosius I (392 CE) made Christianity the sole official religion"],
    ["■ 版圖 Reach",        "600 CE 前，前羅馬帝國幾乎全境皆已基督教化\nBy 600 CE, virtually all of the former Roman Empire was Christian"],
  ];

  facts.forEach(([label, content], i) => {
    const y = 1.22 + i * 0.83;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.73, fill: { color: "3A2015" }, rounding: true });
    s.addText(label, { x: 0.45, y: y + 0.07, w: 2.1, h: 0.55, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(content, { x: 2.55, y: y + 0.05, w: 6.95, h: 0.62, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 4 · The Judaic Heritage ────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("猶太傳統 The Judaic Heritage", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Left column - content
  s.addText("聖殿儀式 Temple Rites", {
    x: 0.4, y: 0.88, w: 5.5, h: 0.4, fontSize: 16, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  const bullets = [
    "第二聖殿（ca. 516 BCE – 70 CE）：利未人唱詩班演唱詩篇，豎琴與聖詠相伴\nSecond Temple (ca. 516 BCE – 70 CE): Levite choirs sang psalms with harp and psaltery",
    "節日與安息日舉行犧牲獻祭，配以吹號與銅鈸\nSacrifices on festivals and Sabbath accompanied by trumpets and cymbals",
    "70 CE 羅馬摧毀聖殿 → 儀式轉移至會堂（synagogue）\nRomans destroyed the Temple (70 CE) → worship moved to synagogues",
    "會堂以誦讀聖經為主，採用吟誦（cantillation）公式\nSynagogue centered on public reading of Scripture using cantillation (chanting formulas)",
  ];
  s.addText(bullets.map((b, i) => ({
    text: b, options: { bullet: true, breakLine: i < bullets.length - 1, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 5 }, valign: "top",
  })), { x: 0.4, y: 1.32, w: 5.6, h: 3.5 });

  // Right column - key terms box
  s.addShape(pres.ShapeType.rect, { x: 6.2, y: 0.88, w: 3.4, h: 4.2, fill: { color: C.wine }, rounding: true });
  s.addText("■ 關鍵詞 Key Terms", { x: 6.3, y: 0.98, w: 3.2, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });

  const terms = [
    ["Psalms 詩篇", "希伯來詩歌聖詩；基督教與猶太教皆用"],
    ["Cantillation 吟誦", "以旋律公式演唱聖典，依句法分段"],
    ["Levites 利未人", "聖殿中負責音樂的祭司階層"],
    ["Synagogue 會堂", "猶太教聚會場所，聖殿崩毀後的禮拜中心"],
  ];
  terms.forEach(([term, def], i) => {
    const y = 1.48 + i * 0.92;
    s.addShape(pres.ShapeType.rect, { x: 6.35, y, w: 3.1, h: 0.82, fill: { color: "5A2020" }, rounding: true });
    s.addText(term, { x: 6.45, y: y + 0.04, w: 2.9, h: 0.32, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(def, { x: 6.45, y: y + 0.38, w: 2.9, h: 0.38, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 5 · Music in the Early Church ──────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("早期教會音樂 Music in the Early Church", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.03, fill: { color: C.sand } });

  s.addText("最早有文字記載的基督教音樂活動是演唱讚美詩與詩篇（馬太 26:30；馬可 14:26；以弗所書 5:19）\nEarliest recorded musical activity: singing hymns and psalms (Matthew 26:30; Ephesians 5:19)", {
    x: 0.4, y: 0.84, w: 9.2, h: 0.55, fontSize: 14, color: C.slate, italic: true, fontFace: "Calibri", valign: "top",
  });

  const cards = [
    ["■\n會堂", "Basilica", "四世紀起公開禮拜在大型教堂舉行\n吟誦有助聲音傳播", "Large basilica buildings\nchanting carries clearly"],
    ["■\n修道院", "Monasteries", "修士每日數次吟誦詩篇\n靈性操練的核心", "Psalm singing\nas spiritual discipline"],
    ["■\n教父", "Church Fathers", "聖巴西流、聖奧古斯丁以希臘哲學詮釋音樂", "Music shapes soul\nbut pleasure suspect"],
    ["■\n樂器", "Instruments", "早期教會禁用器樂\n音樂只服侍語言", "Unaccompanied singing\nfor a millennium"],
  ];

  cards.forEach(([icon, name, zh, en], i) => {
    const x = 0.3 + i * 2.35;
    s.addShape(pres.ShapeType.rect, { x, y: 1.48, w: 2.15, h: 3.75, fill: { color: C.wine }, rounding: true });
    s.addText(icon, { x, y: 1.55, w: 2.15, h: 0.7, fontSize: 22, align: "center" });
    s.addText(name, { x: x + 0.1, y: 2.28, w: 1.95, h: 0.38, fontSize: 14, bold: true, color: C.gold, align: "center", fontFace: "Georgia" });
    s.addShape(pres.ShapeType.rect, { x: x + 0.3, y: 2.68, w: 1.55, h: 0.03, fill: { color: C.gold } });
    s.addText(zh, { x: x + 0.08, y: 2.74, w: 1.99, h: 1.3, fontSize: 14, color: C.cream, align: "center", fontFace: "Calibri", valign: "top" });
    s.addText(en, { x: x + 0.08, y: 4.08, w: 1.99, h: 1.0, fontSize: 14, color: C.sand, align: "center", italic: true, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 6 · Church Fathers on Music ────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("教父論音樂 Church Fathers on Music", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("Source Readings 原典摘錄", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.35, fontSize: 14, color: C.sand, italic: true, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.1, w: 7, h: 0.04, fill: { color: C.gold } });

  // St. Basil - left card
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.2, w: 4.5, h: 4.05, fill: { color: "3A2015" }, rounding: true });
  s.addText("聖巴西流 St. Basil\nca. 330–379 CE", {
    x: 0.45, y: 1.28, w: 4.2, h: 0.65, fontSize: 15, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.8, y: 1.95, w: 3.7, h: 0.03, fill: { color: C.gold } });
  s.addText("「聖靈見人心傾向罪惡，便將教義的甘甜與旋律融合，以使人在不知不覺中受到教化。」\n\nThe Holy Spirit blended the delight of melody with doctrine, so that through the pleasantness of what is heard, we might unawares receive what is useful.\n\n讚美詩篇演唱：培育靈魂、凝聚社群、以音樂驅逐怒氣\nPsalm-singing: educates the soul, builds community, calms anger", {
    x: 0.45, y: 2.02, w: 4.2, h: 3.1, fontSize: 14, color: C.sand, fontFace: "Calibri", italic: false, valign: "top",
  });

  // St. Augustine - right card
  s.addShape(pres.ShapeType.rect, { x: 5.2, y: 1.2, w: 4.5, h: 4.05, fill: { color: "3A2015" }, rounding: true });
  s.addText("聖奧古斯丁 St. Augustine\n354–430 CE", {
    x: 5.35, y: 1.28, w: 4.2, h: 0.65, fontSize: 15, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 5.7, y: 1.95, w: 3.7, h: 0.03, fill: { color: C.gold } });
  s.addText("「我在歌聲中所獲的感動……與歌詞的含意俱來；我承認這種做法是有益的。」\n\nI acknowledge the great benefit of singing in church, while wavering between the peril of pleasure and the benefit of my experience.\n\n態度：音樂有益於靈性，但若旋律比歌詞更動人，則視為罪\nTension: music aids devotion, but if the song moves more than the text, it is sinful", {
    x: 5.35, y: 2.02, w: 4.2, h: 3.1, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 7 · Divisions in the Church & Dialects of Chant ───────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("教會分裂與聖詠方言", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addText("Divisions in the Church and Dialects of Chant", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Calibri", margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.02, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Two main blocks side by side
  const traditions = [
    {
      title: "拜占庭聖詠 Byzantine Chant",
      color: C.rust,
      items: [
        "東羅馬帝國（Constantinople）教會傳統",
        "Eastern Roman Empire (Constantinople)",
        "以希臘語演唱聖詠",
        "Chanted in Greek",
        "八種調式（echoi）——後成西方八種 modes 基礎",
        "Eight echoi (modes) → basis of Western 8 modes",
        "讚美詩（hymns）高度發展，現仍在希臘東正教使用",
        "Hymns highly developed, still sung in Greek Orthodox services",
      ],
    },
    {
      title: "西方聖詠方言 Western Chant Dialects",
      color: C.wine,
      items: [
        "各地域發展獨立傳統（使用拉丁語）",
        "Regional traditions developed (in Latin)",
        "• 高盧聖詠 Gallican (Gaul / France)",
        "• 塞爾特聖詠 Celtic (British Isles)",
        "• 莫扎拉伯聖詠 Mozarabic (Spain)",
        "• 安布羅西聖詠 Ambrosian (Milan)",
        "8–11世紀：羅馬教廷與法蘭克王國逐步統一為格里高利聖詠",
        "8th–11th c.: standardized as Gregorian chant under Rome & Franks",
      ],
    },
  ];

  traditions.forEach(({ title, color, items }, col) => {
    const x = col === 0 ? 0.3 : 5.15;
    s.addShape(pres.ShapeType.rect, { x, y: 1.1, w: 4.55, h: 3.85, fill: { color }, rounding: true });
    s.addText(title, { x: x + 0.15, y: 1.18, w: 4.25, h: 0.5, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
    s.addShape(pres.ShapeType.rect, { x: x + 0.35, y: 1.7, w: 3.85, h: 0.03, fill: { color: C.gold } });
    s.addText(items.map((t, i) => ({
      text: t, options: { bullet: !t.startsWith("•"), breakLine: i < items.length - 1, fontSize: 14, color: C.cream, fontFace: "Calibri", paraSpaceAfter: 3 }, valign: "top",
    })), { x: x + 0.15, y: 1.77, w: 4.25, h: 3.1 });
  });

  // Key fact: 1054 split
  s.addText("1054 CE：東西教會正式分裂 → 羅馬天主教 vs 東正教（拜占庭）\n1054 CE: Final split — Roman Catholic Church vs. Eastern Orthodox (Byzantine) Church", {
    x: 0.3, y: 4.98, w: 9.4, h: 0.42, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Calibri", align: "center", valign: "top",
  });
}

// ── SLIDE 8 · The Creation of Gregorian Chant ────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("格里高利聖詠的形成", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("The Creation of Gregorian Chant", {
    x: 0.4, y: 0.72, w: 9.2, h: 0.35, fontSize: 16, color: C.sand, italic: true, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.1, w: 7, h: 0.04, fill: { color: C.gold } });

  const steps = [
    ["■ 教宗額我略一世\nPope Gregory I (r. 590–604)", "建立詠唱學校（Schola Cantorum）；聖詠以「格里高利聖詠」之名流傳（實為後人追授）\nFounded Schola Cantorum; chant bears his name though he likely did not compose it"],
    ["■ 矮子丕平 Pippin the Short\n(r. 751–768)", "法蘭克王授命全國採用羅馬禮儀與聖詠，以鞏固王權\nFrankish king ordered Roman liturgy and chant throughout his lands to consolidate authority"],
    ["■ 查理曼 Charlemagne\n(r. 768–814)", "繼續推廣政策；派遣羅馬歌手北上教授聖詠；800 CE 由教宗加冕為皇帝\nContinued promotion; sent Roman singers north; crowned Emperor by the pope in 800 CE"],
    ["■ 西歐標準化 Standardization\n(9th–11th centuries)", "格里高利聖詠成為西歐共同音樂語言；記譜法的發明使旋律固定化\nGregorian chant became the common musical language of Western Europe; notation fixed the melodies"],
  ];

  steps.forEach(([label, text], i) => {
    const y = 1.2 + i * 1.0;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.88, fill: { color: "3A2015" }, rounding: true });
    s.addText(label, { x: 0.45, y: y + 0.06, w: 2.6, h: 0.75, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addShape(pres.ShapeType.rect, { x: 3.15, y: y + 0.12, w: 0.04, h: 0.65, fill: { color: C.gold } });
    s.addText(text, { x: 3.28, y: y + 0.08, w: 6.3, h: 0.72, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 });
  });
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
    ["ca. 530s–516 BCE", "耶路撒冷第二聖殿建立 Second Temple built in Jerusalem"],
    ["33 CE",            "耶穌被釘十字架 Jesus's crucifixion"],
    ["70 CE",            "羅馬摧毀聖殿，驅逐猶太人 Romans destroy Temple, expel the Jews"],
    ["313 CE",           "君士坦丁《米蘭勒令》合法化基督教 Edict of Milan legalizes Christianity"],
    ["392 CE",           "基督教成為羅馬官方宗教 Christianity becomes official Roman religion"],
    ["ca. 500–10 CE",    "波愛修斯《音樂基礎》De institutione musica (Boethius)"],
    ["590–604 CE",       "教宗額我略一世 Reign of Pope Gregory I (the Great)"],
    ["ca. 754 CE",       "丕平令全境採用羅馬禮儀 Pippin orders Roman liturgy and chant"],
    ["800 CE",           "查理曼加冕為皇帝 Charlemagne crowned emperor by the pope"],
    ["ca. 850–900 CE",   "《音樂手冊》Musica enchiriadis and Scelica enchiriadis"],
    ["1025–28 CE",       "阿雷佐的桂多《小論文》Guido of Arezzo, Micrologus"],
    ["1054 CE",          "東西教會正式分裂 Final split of Roman and Byzantine churches"],
  ];

  // Timeline line
  s.addShape(pres.ShapeType.rect, { x: 2.6, y: 0.85, w: 0.05, h: 4.5, fill: { color: C.gold } });

  events.forEach(([date, event], i) => {
    const y = 0.85 + i * 0.375;
    s.addShape(pres.ShapeType.ellipse, { x: 2.46, y: y + 0.04, w: 0.28, h: 0.28, fill: { color: C.gold } });
    s.addText(date, { x: 0.1, y, w: 2.28, h: 0.33, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", align: "right", margin: 0 });
    s.addText(event, { x: 2.95, y, w: 6.8, h: 0.33, fontSize: 14, color: C.lightText, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 10 · The Development of Notation (Neumes) ─────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust);
  bottomBar(s, C.rust);

  s.addText("記譜法的發展 The Development of Notation", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.rust, fontFace: "Georgia", margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.76, w: 9.2, h: 0.03, fill: { color: C.sand } });

  s.addText("「除非被記住，否則聲音就消逝了；它們無法寫下。」—— 塞維利亞的伊西多爾 (ca. 560–636)\nUnless sounds are remembered by man, they perish, for they cannot be written down. — Isidore of Seville", {
    x: 0.4, y: 0.84, w: 9.2, h: 0.5, fontSize: 14, color: C.slate, italic: true, fontFace: "Calibri", valign: "top",
  });

  const stages = [
    ["① 口耳相傳\nOral Transmission", "記譜法出現前，聖詠靠口傳記憶；歌手學習數以百計的旋律，全靠記憶\nBefore notation, chants were memorized through oral transmission; singers memorized hundreds of melodies"],
    ["② 紐姆符號\nNeumes (9th c.)", "寫在文字上方的記號（gesture/shape），提示旋律輪廓——但不指定確切音高\nSigns placed above text to indicate melodic contour—direction, shape—but not precise pitches or intervals"],
    ["③ 有高度紐姆\nHeighted Neumes (10–11th c.)", "紐姆符號寫在不同高度以指示音高；加一條橫線輔助辨識音高（音高相對但仍不確定）\nNeumes placed at varying heights above text to indicate relative pitch; a scratched line helped identify one note"],
    ["④ 線譜與音符\nLines, Clef & Staff", "阿雷佐的桂多（ca. 991–1033）：四線譜，F、C 譜號，紐姆符號改良為精確音符\nGuido of Arezzo: four-line staff with F and C clefs; neumes reshaped to indicate pitch exactly; music could now be read at sight"],
  ];

  stages.forEach(([title, desc], i) => {
    const y = 1.42 + i * 0.98;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.86, fill: { color: C.wine }, rounding: true });
    s.addText(title, { x: 0.45, y: y + 0.06, w: 2.4, h: 0.72, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addShape(pres.ShapeType.rect, { x: 2.95, y: y + 0.12, w: 0.04, h: 0.6, fill: { color: C.gold } });
    s.addText(desc, { x: 3.08, y: y + 0.07, w: 6.5, h: 0.72, fontSize: 14, color: C.cream, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 11 · Boethius & Three Types of Music ───────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("波愛修斯與音樂理論", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("Boethius (ca. 480–524) · De institutione musica", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.38, fontSize: 16, italic: true, color: C.sand, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.1, w: 7, h: 0.04, fill: { color: C.gold } });

  // Context
  s.addText("波愛修斯生於羅馬貴族，後受哥德統治者狄奧多里克任命，依據希臘文獻（畢達哥拉斯、托勒密）撰寫《音樂基礎》。此書在中世紀被奉為最高音樂理論權威超過一千年。\nBoethius, born into Roman nobility, compiled De institutione musica from Greek sources (Pythagoras, Ptolemy). It was the authoritative text on music theory for over a thousand years.", {
    x: 0.4, y: 1.18, w: 9.2, h: 0.75, fontSize: 14, color: C.sand, fontFace: "Calibri", italic: false, valign: "top",
  });

  // Three types of music
  const types = [
    {
      name: "Musica Mundana\n宇宙音樂",
      desc: "天體運行、四季更替的數學比例\nNumerical ratios of the heavens and seasons",
      color: "1A3A5C",
      icon: "■",
    },
    {
      name: "Musica Humana\n人體音樂",
      desc: "人體與靈魂之間的和諧\nHarmony between body and soul",
      color: "2A5C1A",
      icon: "■",
    },
    {
      name: "Musica Instrumentalis\n樂器音樂",
      desc: "人聲或樂器產生的可聽音樂\nAudible music from instruments or voices",
      color: "5C2A1A",
      icon: "■",
    },
  ];

  types.forEach(({ name, desc, color, icon }, i) => {
    const x = 0.3 + i * 3.25;
    s.addShape(pres.ShapeType.rect, { x, y: 2.05, w: 3.05, h: 2.75, fill: { color }, rounding: true });
    s.addText(icon, { x, y: 2.1, w: 3.05, h: 0.45, fontSize: 22, align: "center" });
    s.addText(name, { x: x + 0.1, y: 2.55, w: 2.85, h: 0.6, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
    s.addShape(pres.ShapeType.rect, { x: x + 0.3, y: 3.18, w: 2.45, h: 0.03, fill: { color: C.gold } });
    s.addText(desc, { x: x + 0.1, y: 3.3, w: 2.85, h: 1.4, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", align: "center" });
  });

  s.addText("對波愛修斯而言，音樂首先是知識的對象；真正的音樂家是「理解」音樂的人，而非演奏者\nFor Boethius, music is knowledge, not practice; the true musician understands, not plays", {
    x: 0.4, y: 4.92, w: 9.2, h: 0.5, fontSize: 14, color: C.sand, italic: true, fontFace: "Calibri", align: "center", valign: "top",
  });
}

// ── SLIDE 12 · The Church Modes ──────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine);
  bottomBar(s, C.wine);

  s.addText("教會調式體系 The Church Modes", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", margin: 0,
  });
  s.addText("源自拜占庭的八種 echoi，是組織與記憶格里高利聖詠的核心工具 · Adapted from Byzantine echoi; central tool for classifying and memorizing Gregorian chant", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.35, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Calibri", margin: 0, valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.06, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Table header
  const headers = ["調式 Mode", "最終音 Final", "類型 Type", "音程特徵 Intervals", "調式名 Greek Name"];
  const hWidths = [2.15, 1.1, 1.45, 2.3, 2.2];
  let hx = 0.35;
  headers.forEach((h, i) => {
    s.addShape(pres.ShapeType.rect, { x: hx, y: 1.12, w: hWidths[i] - 0.05, h: 0.4, fill: { color: C.wine }, rounding: false });
    s.addText(h, { x: hx + 0.05, y: 1.14, w: hWidths[i] - 0.1, h: 0.36, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    hx += hWidths[i];
  });

  // Mode rows
  const modes = [
    ["1 Dorian",      "D", "正格 Authentic", "全 全 半 | 全 全 半 全", "Dorian 多利安"],
    ["2 Hypodorian",  "D", "變格 Plagal",    "全 半 全 | 全 半 全 全", "Hypodorian"],
    ["3 Phrygian",    "E", "正格 Authentic", "半 全 全 | 半 全 全 全", "Phrygian 弗里幾安"],
    ["4 Hypophrygian","E", "變格 Plagal",    "半 全 全 | 全 全 全 半", "Hypophrygian"],
    ["5 Lydian",      "F", "正格 Authentic", "全 全 全 | 半 全 全 半", "Lydian 利底安"],
    ["6 Hypolydian",  "F", "變格 Plagal",    "全 全 半 | 全 全 全 半", "Hypolydian"],
    ["7 Mixolydian",  "G", "正格 Authentic", "全 全 半 | 全 全 半 全", "Mixolydian 混利底安"],
    ["8 Hypomixolydian","G", "變格 Plagal",  "半 全 全 | 半 全 全 半", "Hypomixolydian"],
  ];

  modes.forEach(([mode, final, type, intervals, greek], i) => {
    const y = 1.55 + i * 0.42;
    const bg = i % 2 === 0 ? "F0E8D8" : C.cream;
    const row = [mode, final, type, intervals, greek];
    let rx = 0.35;
    row.forEach((cell, j) => {
      s.addShape(pres.ShapeType.rect, { x: rx, y, w: hWidths[j] - 0.05, h: 0.4, fill: { color: bg } });
      s.addText(cell, {
        x: rx + 0.05, y: y + 0.02, w: hWidths[j] - 0.1, h: 0.36,
        fontSize: 14,
        bold: j === 0,
        color: j === 0 ? C.wine : C.darkText,
        fontFace: j === 0 ? "Georgia" : "Calibri",
        margin: 0,
      });
      rx += hWidths[j];
    });
  });

  s.addText("正格 (authentic)：音域從 final 上行一個八度；變格 (plagal)：音域跨越 final 上下\nAuthentic: range extends up an octave from final; Plagal: range spans both sides of final", {
    x: 0.35, y: 4.95, w: 9.3, h: 0.5, fontSize: 14, color: C.midBrown, italic: true, fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 13 · Solmization & Guido of Arezzo ─────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("唱名法與阿雷佐的桂多", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addText("Solmization & Guido of Arezzo (ca. 991 – after 1033)", {
    x: 0.4, y: 0.7, w: 9.2, h: 0.35, fontSize: 15, italic: true, color: C.sand, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.08, w: 7, h: 0.04, fill: { color: C.gold } });

  // Left: Solmization
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.18, w: 5.1, h: 4.05, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ 唱名法 Solmization", {
    x: 0.45, y: 1.26, w: 4.8, h: 0.45, fontSize: 16, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("桂多從讚美詩《Ut queant laxis》每句開頭音節命名六個音\nGuido named six syllables from the hymn Ut queant laxis", {
    x: 0.45, y: 1.74, w: 4.8, h: 0.55, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top",
  });

  // Syllable table
  const syllables = [
    ["ut*", "C"],
    ["re", "D"],
    ["mi", "E"],
    ["fa", "F"],
    ["sol", "G"],
    ["la", "A"],
  ];
  syllables.forEach(([syl, note], i) => {
    const sx = 0.55 + i * 0.8;
    s.addShape(pres.ShapeType.rect, { x: sx, y: 2.35, w: 0.68, h: 0.9, fill: { color: C.wine }, rounding: true });
    s.addText(syl, { x: sx, y: 2.4, w: 0.68, h: 0.42, fontSize: 18, bold: true, color: C.gold, align: "center", fontFace: "Georgia" });
    s.addText(note, { x: sx, y: 2.82, w: 0.68, h: 0.38, fontSize: 14, color: C.cream, align: "center", fontFace: "Calibri", valign: "top" });
  });

  s.addText("* ut 後改為 do (later changed to do)\n唱名法的功能：幫助歌手辨識音程；mi→fa 是半音，其他相鄰音為全音\nSolmization helped singers identify intervals: mi–fa is a semitone, all others are whole tones", {
    x: 0.45, y: 3.35, w: 4.8, h: 1.75, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top",
  });

  // Right: Guido achievements
  s.addShape(pres.ShapeType.rect, { x: 5.6, y: 1.18, w: 4.1, h: 4.05, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ 桂多的貢獻\nGuido's Achievements", {
    x: 5.75, y: 1.26, w: 3.8, h: 0.58, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 5.95, y: 1.86, w: 3.4, h: 0.03, fill: { color: C.gold } });

  const guido = [
    ["■ 四線譜 Staff", "發明現代五線譜前身：四線有色譜（F 線紅、C 線黃）\nFour-line staff with colored F and C clef lines"],
    ["■ 字母譜號 Clefs", "制訂 A–G 音符字母系統（沿用至今）\nDeveloped A–G letter names for notes"],
    ["■ 桂多手勢 Hand", "以左手關節代表音符，視覺化教學工具\nGuidonian hand: mnemonic using finger joints"],
  ];
  guido.forEach(([label, desc], i) => {
    const y = 2.0 + i * 1.05;
    s.addText(label, { x: 5.75, y, w: 3.8, h: 0.32, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(desc, { x: 5.75, y: y + 0.34, w: 3.8, h: 0.65, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 14 · Chapter Summary / Echoes of History ──────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold);
  bottomBar(s, C.gold);

  s.addText("本章重點回顧 Chapter Summary", {
    x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 0.76, w: 7, h: 0.04, fill: { color: C.gold } });

  const points = [
    ["■", "基督教脫胎於猶太傳統（詩篇、吟誦），在羅馬帝國中傳播，313 CE 合法化，392 CE 成為官方宗教\nChristianity grew from Jewish roots (psalms, cantillation), spread through Rome, legalized 313, official religion 392"],
    ["■", "早期教會以單聲部演唱為主；樂器被排除在禮拜之外超過一千年\nEarly church favored monophonic singing; instruments excluded from worship for over a millennium"],
    ["■", "東西方教會在儀式、語言（希臘文 vs. 拉丁文）、聖詠方言上逐漸分歧；1054 正式分裂\nEastern and Western churches diverged in rite, language, and chant dialects; final split 1054"],
    ["■", "格里高利聖詠在法蘭克王國支持下統一西歐禮拜音樂，成為西方音樂最重要的基礎\nGregorian chant unified Western church music with Frankish support; became the foundation of Western music"],
    ["■", "記譜法從口傳→紐姆→有高度紐姆→四線譜；桂多的發明使音樂可以「視唱」\nNotation evolved from oral→neumes→heighted neumes→Guido's 4-line staff; sight-reading became possible"],
    ["■", "波愛修斯的三種音樂分類與教會八種調式奠定了中世紀音樂理論的基礎，影響延續到文藝復興\nBoethius's three types of music and the 8 church modes formed the basis of medieval theory into the Renaissance"],
  ];

  points.forEach(([icon, text], i) => {
    const y = 0.88 + i * 0.74;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.64, fill: { color: "3A2015" }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.07, w: 0.55, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(text, { x: 1.05, y: y + 0.04, w: 8.5, h: 0.56, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 15 · Further Reading ────────────────────────────────────────────────
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
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.05, w: 4.3, h: 2.85, fill: { color: C.wine }, rounding: true });
  s.addText("■ 聆聽 YouTube", {
    x: 0.55, y: 1.15, w: 4.0, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  s.addText("Gregorian — Viderunt omnes\nyoutu.be/uvC0NGmpuFY\nHildegard — O Virtus Sapientiae\nyoutu.be/77M2pkkdH10\nByzantine — Cherubic Hymn\nyoutu.be/PqYUu7uGXjs\nAmbrosian — Lucernario\nGuido — Ut queant laxis", {
    x: 0.55, y: 1.65, w: 4.0, h: 2.15, fontSize: 14, color: C.cream, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top", margin: 0,
  });

  // Reading
  s.addShape(pres.ShapeType.rect, { x: 4.9, y: 1.05, w: 4.7, h: 2.85, fill: { color: C.rust }, rounding: true });
  s.addText("■ 閱讀 Read", {
    x: 5.05, y: 1.15, w: 4.4, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  const reading = [
    "Wikipedia: Gregorian chant",
    "Stanford: Medieval Music Theory",
    "CPDL: chant scores",
    "Boethius, De institutione musica",
    "Guido of Arezzo, Micrologus",
  ];
  s.addText(reading.map((r, i) => ({
    text: r, options: { bullet: true, breakLine: i < reading.length - 1, fontSize: 14, color: C.cream, fontFace: "Calibri", paraSpaceAfter: 6 }, valign: "top",
  })), { x: 5.05, y: 1.65, w: 4.4, h: 2.15, margin: 0 });

  // Academic / key concepts
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 4.0, w: 9.2, h: 1.4, fill: { color: C.midBrown }, rounding: true });
  s.addText("■ 本章關鍵術語 Key Terms", {
    x: 0.6, y: 4.05, w: 8.8, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0,
  });
  const terms = "Psalm 詩篇 · Cantillation 吟誦 · Gregorian Chant 格里高利聖詠 · Plainchant 平唱 · Rite 儀式 · Liturgy 禮拜儀式 · Neumes 紐姆符號 · Mode 調式 · Final 最終音 · Authentic 正格 · Plagal 變格 · Solmization 唱名法";
  s.addText(terms, {
    x: 0.6, y: 4.42, w: 8.8, h: 0.95, fontSize: 14, color: C.cream, fontFace: "Calibri", margin: 0, valign: "top",
  });
}

// ── Generate file ─────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "Ch02_Christian_Church.pptx" })
  .then(() => console.log("■ Ch02_Christian_Church.pptx created successfully"))
  .catch(err => console.error("■ Error:", err));
