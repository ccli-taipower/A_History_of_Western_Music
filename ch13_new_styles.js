const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 13: New Styles in the Seventeenth Century";
pres.author = "A History of Western Music, 10th ed.";

// Baroque palette — dramatic, theatrical
const C = {
  darkBg: "1A1A2E", gold: "C8A030", cream: "F5F0E0",
  navy: "16213E", royal: "0F3460", darkText: "1A1A2E",
  lightText: "F5F0E0", sand: "E8D8A8", slate: "1A2238",
  teal: "1A5276", steel: "2C3E50"
};

function darkSlide(pres) { const s = pres.addSlide(); s.background = { color: C.darkBg }; return s; }
function lightSlide(pres) { const s = pres.addSlide(); s.background = { color: C.cream }; return s; }
function topBar(s, color) { s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.12, fill: { color: color || C.gold } }); }
function bottomBar(s, color) { s.addShape(pres.ShapeType.rect, { x: 0, y: 5.5, w: "100%", h: 0.125, fill: { color: color || C.gold } }); }

// ── SLIDE 1 · Title ──────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.15, fill: { color: C.gold } });
  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.47, w: "100%", h: 0.155, fill: { color: C.gold } });

  s.addText("A HISTORY OF WESTERN MUSIC · TENTH EDITION", {
    x: 0.5, y: 0.45, w: 9, h: 0.35, fontSize: 18, color: C.sand, charSpacing: 3, align: "center", fontFace: "Georgia",
  });
  s.addText("CHAPTER 13", {
    x: 0.5, y: 0.95, w: 9, h: 0.55, fontSize: 24, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });
  s.addText("NEW STYLES IN THE\nSEVENTEENTH CENTURY\n十七世紀的新風格", {
    x: 0.3, y: 1.6, w: 9.4, h: 2.2, fontSize: 36, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.9, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Monteverdi · Caccini · Basso Continuo · Baroque", {
    x: 0.4, y: 4.05, w: 9.2, h: 0.4, fontSize: 18, color: C.sand, align: "center", fontFace: "Georgia",
  });
  s.addText("Textbook pp. 278–301", {
    x: 0.5, y: 4.75, w: 9, h: 0.3, fontSize: 18, color: C.gold, align: "center", fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 2 · Chapter Overview ──────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);

  s.addText("本章概覽 Chapter Overview", { x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.teal, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const sections = [
    ["十七世紀歐洲", "Europe in the 17th Century: science, war, patronage"],
    ["從文藝復興到巴洛克", "From Renaissance to Baroque: drama & the affections"],
    ["第一實踐 vs 第二實踐", "Prima Pratica vs. Seconda Pratica: Artusi vs. Monteverdi"],
    ["單聲歌曲與數字低音", "Monody, Camerata, Basso Continuo & Le nuove musiche"],
    ["巴洛克音樂特徵", "General Traits: texture, rhythm, tonality, styles"],
  ];
  sections.forEach(([zh, en], i) => {
    const y = 1.05 + i * 0.85;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 9.2, h: 0.72, fill: { color: C.teal }, rounding: true });
    s.addText(zh, { x: 0.6, y: y + 0.02, w: 8.8, h: 0.35, fontSize: 20, bold: true, color: C.lightText, fontFace: "Georgia" });
    s.addText(en, { x: 0.6, y: y + 0.36, w: 8.8, h: 0.32, fontSize: 18, color: C.sand, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 3 · Europe in the 17th Century: Science & War ─────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("十七世紀的歐洲", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Europe in the 17th Century: Science, War, & Exploration", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const bullets = [
    "科學革命 Scientific Revolution: Kepler (1609), Galileo, Bacon, Descartes, Newton",
    "三十年戰爭 Thirty Years' War (1618-48) 摧毀日耳曼地區",
    "英國內戰 English Civil War (1642-49); 王權復辟 1660",
    "法國太陽王 Louis XIV (r. 1643-1715) 專制王權的高峰",
    "殖民擴張 Colonies in Americas, Caribbean, Africa, Asia",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.lightText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 4 · Patronage & National Styles ───────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);

  s.addText("贊助制度與民族風格", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addText("Patronage & National Styles", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.steel, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.teal } });

  const bullets = [
    "資本主義 Capitalism: 股份公司資助歌劇院 (Hamburg, London)",
    "義大利 Italy: 貿易富裕，佛羅倫斯 / 羅馬 / 威尼斯的貴族贊助音樂創新",
    "法國 France: Louis XIV 集權控制藝術，法國風格廣受模仿",
    "日耳曼 Germany: 三十年戰爭後，宮廷與自由城市融合意法風格",
    "公共音樂場所興起 Public opera (Venice, 1637) & concerts (London, 1672)",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.darkText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 5 · The Term "Baroque" ────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("「巴洛克」一詞的由來", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("The Baroque as Term and Period (ca. 1600–1750)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const bullets = [
    "源自葡萄牙語 barroco（不規則珍珠），原為貶義",
    "18 世紀批評家用以形容華麗、不協和、大膽的音樂與藝術",
    "19 世紀 → 正面意涵：裝飾性、戲劇性、情感表現力",
    "1950 年代確立為音樂史分期：約 1600–1750",
    "非單一風格，而是共享戲劇性、情感表現理念的時代",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.lightText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 6 · The Dramatic Baroque ──────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.steel); bottomBar(s, C.steel);

  s.addText("戲劇性的巴洛克", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.steel, fontFace: "Georgia", align: "center" });
  s.addText("The Dramatic Baroque: Art, Architecture, & Music", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.steel } });

  const bullets = [
    "戲劇是巴洛克核心衝動 Drama is central to Baroque art & music",
    "Bernini《大衛》(1623) vs. Michelangelo《大衛》— 靜態 → 動態情感",
    "Bernini《聖乙蕾莎的狂喜》(1645-52) — 劇場化的宗教體驗",
    "音樂同樣追求 dramatic effect：旋律動態、和聲對比、節奏張力",
    "觀眾概念出現 — 表演者成為專業人士，聽眾成為被動接收者",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.darkText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 7 · The Affections ────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("情感論 The Affections", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Moving the Passions through Music", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const bullets = [
    "Affections 情感 = 理性化的情緒 (sadness, joy, anger, love, fear)",
    "亞里斯多德觀：情感由外在刺激引發靈魂狀態",
    "笛卡兒《靈魂的激情》(1649)：情感是客觀的、可分類的",
    "作曲家用特定手法喚起對應情感 — 非個人表達，而是「再現」",
    "每首樂曲或樂章通常只呈現一種主要情感",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.lightText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 8 · Prima Pratica vs Seconda Pratica ──────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);

  s.addText("第一實踐 vs. 第二實踐", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addText("Prima Pratica vs. Seconda Pratica", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 22, color: C.steel, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.teal } });

  // Two-column comparison
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.5, h: 3.9, fill: { color: C.navy }, rounding: true });
  s.addText("Prima Pratica 第一實踐", { x: 0.5, y: 1.4, w: 4.1, h: 0.45, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 由 Zarlino 法則主導的對位法\n• 音樂有其自身規則\n• 和聲支配文字\n  Harmony is mistress of the words", {
    x: 0.5, y: 1.95, w: 4.1, h: 3.0, fontSize: 18, color: C.lightText, fontFace: "Calibri", paraSpaceAfter: 6, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.2, y: 1.3, w: 4.5, h: 3.9, fill: { color: C.royal }, rounding: true });
  s.addText("Seconda Pratica 第二實踐", { x: 5.4, y: 1.4, w: 4.1, h: 0.45, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Monteverdi 兄弟提出 (1605/07)\n• 文字是音樂的主宰\n• 為表現歌詞情感可打破對位規則\n  Words are mistress of the harmony", {
    x: 5.4, y: 1.95, w: 4.1, h: 3.0, fontSize: 18, color: C.lightText, fontFace: "Calibri", paraSpaceAfter: 6, valign: "top",
  });
}

// ── SLIDE 9 · The Artusi–Monteverdi Debate ──────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("阿圖西—蒙台威爾第之爭", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("The Artusi–Monteverdi Debate (1600–1607)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const bullets = [
    "Artusi《L'Artusi》(1600) 抨擊 Monteverdi 牧歌中的不協和音",
    "指控 Cruda Amarilli 違反對位法 — 未預備的不協和、禁止的經過音",
    "Monteverdi 於第五冊牧歌序言 (1605) 回應：存在「第二實踐」",
    "弟弟 Giulio Cesare 詳述 (1607)：文字主宰和聲，非和聲主宰文字",
    "此論戰確立了「新 vs. 舊」風格的理論基礎",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.lightText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 10 · NAWM 75 Cruda Amarilli ───────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.steel); bottomBar(s, C.steel);

  s.addText("NAWM 75: Cruda Amarilli", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.steel, fontFace: "Georgia", align: "center" });
  s.addText("Monteverdi · Fifth Book of Madrigals (1605)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.steel } });

  const bullets = [
    "五聲部牧歌 — 詩文取自 Guarini 牧劇《忠實的牧羊人》",
    "\"Cruda\"（殘忍）與 \"ahi lasso\"（嗚呼）處使用大膽不協和音",
    "未預備的不協和音作為修辭手段 — 文字畫 word painting",
    "Artusi 批評的正是此曲：故意違反規則以表現文字痛苦",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.9;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.8, fontSize: 18, color: C.darkText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });

  s.addText("https://www.youtube.com/watch?v=8elXHf0gXTM", {
    x: 0.6, y: 5.0, w: 8.8, h: 0.35, fontSize: 18, color: C.teal, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=8elXHf0gXTM" },
  });
}

// ── SLIDE 11 · Monody & The Florentine Camerata ─────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("單聲歌曲與佛羅倫斯卡梅拉塔", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Monody & The Florentine Camerata", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const bullets = [
    "Camerata: 1570s Count Bardi 沙龍，討論古希臘音樂的力量",
    "Vincenzo Galilei《古今音樂對話》(1581) 攻擊複音、倡導獨唱旋律",
    "Monody 單聲歌曲 = 獨唱 + 伴奏（非 monophony 無伴奏獨唱）",
    "Girolamo Mei 研究結論：古希臘音樂為單旋律，因此最具情感力量",
    "這些理念最終催生了歌劇 Opera 的誕生 (ca. 1598)",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.lightText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 12 · Caccini & Le nuove musiche ───────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);

  s.addText("卡契尼與《新音樂》", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addText("Giulio Caccini · Le nuove musiche (1602)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.steel, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.teal } });

  const bullets = [
    "Caccini (ca. 1550–1618): Camerata 成員，歌手兼作曲家",
    "Le nuove musiche (新音樂, 1602): 獨唱歌曲集 (arias + solo madrigals)",
    "以持續低音伴奏 — 每行詩句獨立成樂句，結束於終止式",
    "旋律配合文字自然聲調 — 裝飾音增強情感，非炫技",
    "序言詳述聲樂裝飾技法 — 成為演唱實踐的重要文獻",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.darkText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 13 · NAWM 74 Vedro 'l mio sol ─────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("NAWM 74: Vedro 'l mio sol", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Caccini · from Le nuove musiche (1602)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const bullets = [
    "Solo madrigal（獨唱牧歌）— 曾在 Camerata 聚會中獲掌聲",
    "Figured bass 數字低音的早期範例 — 數字標示和弦",
    "旋律自然跟隨詩句 — 每行結束於一個 cadence 終止式",
    "裝飾音寫入譜中 — 增強文字意涵而非單純展示技巧",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.9;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.8, fontSize: 18, color: C.lightText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });

  s.addText("https://www.youtube.com/watch?v=urqSFnKGjIQ", {
    x: 0.6, y: 5.0, w: 8.8, h: 0.35, fontSize: 18, color: C.gold, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=urqSFnKGjIQ" },
  });
}

// ── SLIDE 14 · Basso Continuo ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.steel); bottomBar(s, C.steel);

  s.addText("數字低音 Basso Continuo", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.steel, fontFace: "Georgia", align: "center" });
  s.addText("The Foundation of Baroque Texture", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.steel } });

  const bullets = [
    "高低聲部的兩極化 Treble-bass polarity 取代文藝復興聲部平等",
    "作曲家寫旋律 + 低音線；內聲部由演奏者即興填充",
    "Figured bass 數字低音 = 低音線上方數字標示和弦",
    "伴奏樂器: 大鍵琴 harpsichord、管風琴 organ、魯特琴 lute、長頸琴 theorbo",
    "Realization 即興實現 — 演奏者決定和弦排列與裝飾",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.darkText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 15 · Monteverdi: Life & Career ────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("蒙台威爾第：生平與成就", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Claudio Monteverdi (1567–1643)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 22, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const bullets = [
    "出生於 Cremona，師承 Marc'Antonio Ingegneri",
    "1590 年起任職 Mantua 公爵宮廷 — 牧歌集 Books 1–5",
    "1607《奧菲歐》L'Orfeo — 早期最偉大歌劇之一",
    "1613 任威尼斯聖馬可大教堂樂長至逝世 — 最崇高職位",
    "九冊牧歌集橫跨從文藝復興到巴洛克的轉型",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.lightText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 16 · Monteverdi's Madrigal Books ──────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);

  s.addText("蒙台威爾第的牧歌集", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addText("Monteverdi's Nine Books of Madrigals (1587–1651)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.steel, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.teal } });

  const bullets = [
    "Books 1–4 (1587–1603): 五聲部 a cappella 牧歌，延續文藝復興傳統",
    "Book 5 (1605): 加入 basso continuo — 新舊風格交界",
    "Books 6–7: 獨唱牧歌、二重唱、連作牧歌組曲",
    "Book 8《戰爭與愛情牧歌》(1638): stile concitato 激動風格",
    "從對等複音 → 獨唱 + 低音，見證音樂史最重大的風格轉變",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.darkText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 17 · NAWM 76 Zefiro torna ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("NAWM 76: Zefiro torna", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Monteverdi · from Scherzi musicali (1632)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const bullets = [
    "二重唱 + basso continuo — 詩人 Ottavio Rinuccini 的十四行詩",
    "Ciaccona 夏康舞曲低音反覆 — 持續不斷的循環 bass pattern",
    "歡樂段落: 兩位男高音以華麗花腔歌唱自然之美",
    "尾聲突變: \"solo io\" — 低音停止循環，進入悲傷宣敘",
    "極致的情感對比 — 歡愉與孤寂在同一曲中並置",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.72;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.65, fontSize: 18, color: C.lightText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });

  s.addText("https://www.youtube.com/watch?v=Q3T9pinGCrM", {
    x: 0.6, y: 5.05, w: 8.8, h: 0.35, fontSize: 16, color: C.gold, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=Q3T9pinGCrM" },
  });
}

// ── SLIDE 18 · Baroque Traits: Texture & Harmony ────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.steel); bottomBar(s, C.steel);

  s.addText("巴洛克音樂特徵 (一)", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.steel, fontFace: "Georgia", align: "center" });
  s.addText("Texture, Harmony, & the Concertato Medium", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.steel } });

  const bullets = [
    "高低兩極化 Treble-bass polarity 取代文藝復興的聲部平等織度",
    "和弦思維 — 不協和音由和弦而非音程定義",
    "半音主義 Chromaticism 用於表現強烈情感",
    "Concertato medium 協奏媒介: 人聲 + 器樂對比組合",
    "Concerted madrigal、Sacred concerto 等新體裁出現",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.darkText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 19 · Baroque Traits: Rhythm, Idiom, Performance ───────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("巴洛克音樂特徵 (二)", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Rhythm, Idiomatic Writing, & Performance", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const bullets = [
    "節奏兩極: 自由的宣敘調 vs. 嚴格的舞曲節拍 (measures 小節線出現)",
    "器樂慣用法 Idiomatic writing: 小提琴、魯特琴、鍵盤各有專屬技法",
    "演奏者為中心 — 即興裝飾、cadenza、ornaments (trills, mordents)",
    "樂譜是「劇本」而非定稿 — 演奏者有權改編",
    "聽眾角色轉變: 從參與者變成被動接收情感的觀眾",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.lightText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 20 · Three Styles ─────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);

  s.addText("三種風格分類", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addText("Church, Chamber, & Theater Styles", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 22, color: C.steel, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.teal } });

  const items = [
    ["教堂風格 Church Style (Stile Ecclesiastico)", "莊嚴的對位法，管風琴伴奏，拉丁經文，sacred concerto"],
    ["室內風格 Chamber Style (Stile Cubiculare)", "世俗題材，精緻裝飾，獨唱或小團體，宮廷貴族私人場合"],
    ["劇場風格 Theater Style (Stile Theatrale)", "歌劇、宣敘調、詠嘆調，強烈戲劇性，最具創新的領域"],
  ];
  items.forEach(([title, desc], i) => {
    const y = 1.35 + i * 1.3;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 9.2, h: 1.1, fill: { color: C.navy }, rounding: true });
    s.addText(title, { x: 0.6, y: y + 0.08, w: 8.8, h: 0.4, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
    s.addText(desc, { x: 0.6, y: y + 0.52, w: 8.8, h: 0.45, fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 21 · From Modal to Tonal ──────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("從調式到調性", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("From Modal to Tonal Music", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 22, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const bullets = [
    "17 世紀初作曲家仍使用教會調式或 12 調式系統",
    "世紀末 Corelli、Lully 等人已完全在大小調體系中寫作",
    "數字低音促進和弦思維 → 和聲進行逐漸標準化",
    "Rameau《和聲論》(1722) 首次完整理論化調性系統",
    "調性 (tonality) 是漸進演化，非一夕取代調式",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.lightText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 22 · Historically Informed Performance ────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.steel); bottomBar(s, C.steel);

  s.addText("歷史知情演奏 HIP", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.steel, fontFace: "Georgia", align: "center" });
  s.addText("Historically Informed Performance & Its Controversies", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 20, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.steel } });

  const bullets = [
    "使用時代樂器 Period instruments 重現原始音色",
    "研究文獻以重建裝飾、速度、律制 (mean-tone temperament)",
    "爭議: Richard Taruskin 認為追求「原音」是不可能且誤導的目標",
    "Harry Haskell: 可知的歷史與不可知的聲音之間有巨大鴻溝",
    "正面觀點: 學習歷史演奏法可拓展現代演奏者的創造力",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.35 + i * 0.82;
    s.addText(txt, { x: 0.6, y, w: 8.8, h: 0.72, fontSize: 18, color: C.darkText, fontFace: "Calibri", bullet: true, valign: "middle" });
  });
}

// ── SLIDE 23 · Enduring Innovations ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("持久的創新遺產", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Enduring Innovations of the 17th Century", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 22, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Two columns: lasting vs. faded
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.5, h: 3.9, fill: { color: C.slate }, rounding: true });
  s.addText("延續至今 Still With Us", { x: 0.5, y: 1.4, w: 4.1, h: 0.45, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 戲劇性與情感表現力\n• 規則打破作為修辭手段\n• 高低聲部兩極化\n• 和弦式和聲\n• 半音主義\n• 調性系統\n• 器樂慣用寫法", {
    x: 0.5, y: 1.95, w: 4.1, h: 3.0, fontSize: 18, color: C.lightText, fontFace: "Calibri", paraSpaceAfter: 4, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.2, y: 1.3, w: 4.5, h: 3.9, fill: { color: C.slate }, rounding: true });
  s.addText("逐漸消失 Faded Away", { x: 5.4, y: 1.4, w: 4.1, h: 0.45, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 數字低音\n• 演奏者即興填充的傳統\n• 演奏者作為「共同作曲家」\n• 樂譜作為演出「劇本」而非定稿", {
    x: 5.4, y: 1.95, w: 4.1, h: 3.0, fontSize: 18, color: C.lightText, fontFace: "Calibri", paraSpaceAfter: 4, valign: "top",
  });
}

// ── SLIDE 24 · Timeline ─────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);

  s.addText("時間線 Timeline", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 0.75, w: 7, h: 0.04, fill: { color: C.teal } });

  const events = [
    ["1598", "Henri IV Edict of Nantes 南特敕令"],
    ["1600", "Artusi 攻擊 Monteverdi 牧歌"],
    ["1602", "Caccini Le nuove musiche 出版"],
    ["1605", "Monteverdi Fifth Book of Madrigals (含 Cruda Amarilli)"],
    ["1609", "Kepler 天文定律"],
    ["1618-48", "Thirty Years' War 三十年戰爭"],
    ["1637", "Venice 首座公共歌劇院 Teatro San Cassiano"],
    ["1643-1715", "Louis XIV 統治法國"],
    ["1649", "Descartes《靈魂的激情》"],
    ["1687", "Newton Principia mathematica"],
  ];

  // Two columns for timeline
  events.forEach(([ year, desc], i) => {
    const col = i < 5 ? 0 : 1;
    const row = i < 5 ? i : i - 5;
    const x = col === 0 ? 0.3 : 5.1;
    const y = 0.95 + row * 0.88;
    s.addShape(pres.ShapeType.rect, { x, y, w: 1.1, h: 0.75, fill: { color: C.teal }, rounding: true });
    s.addText(year, { x, y: y + 0.02, w: 1.1, h: 0.7, fontSize: 18, bold: true, color: C.lightText, fontFace: "Georgia", align: "center", valign: "middle" });
    s.addText(desc, { x: x + 1.2, y, w: 3.5, h: 0.75, fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "middle" });
  });
}

// ── SLIDE 25 · Key Terms & Listening ────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("關鍵術語與聆聽", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Key Terms & Listening Guide", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 22, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Key terms in two columns
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.5, h: 2.4, fill: { color: C.slate }, rounding: true });
  s.addText("核心術語 Key Terms", { x: 0.5, y: 1.35, w: 4.1, h: 0.4, fontSize: 20, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Baroque 巴洛克 | Affections 情感\nPrima/Seconda pratica\nMonody 單聲歌曲\nBasso continuo 數字低音\nConcertato medium 協奏媒介", {
    x: 0.5, y: 1.8, w: 4.1, h: 1.8, fontSize: 18, color: C.lightText, fontFace: "Calibri", paraSpaceAfter: 4, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.2, y: 1.3, w: 4.5, h: 2.4, fill: { color: C.slate }, rounding: true });
  s.addText("更多術語 More Terms", { x: 5.4, y: 1.35, w: 4.1, h: 0.4, fontSize: 20, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Figured bass 數字低音記譜\nOrnaments 裝飾音\nCadenza 華彩段\nTonality 調性\nIdiomatic writing 器樂慣用法", {
    x: 5.4, y: 1.8, w: 4.1, h: 1.8, fontSize: 18, color: C.lightText, fontFace: "Calibri", paraSpaceAfter: 4, valign: "top",
  });

  // Listening list
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.9, w: 9.4, h: 1.45, fill: { color: C.navy }, rounding: true });
  s.addText("NAWM 聆聽清單", { x: 0.5, y: 3.95, w: 9.0, h: 0.38, fontSize: 20, bold: true, color: C.gold, fontFace: "Georgia" });

  s.addText("NAWM 74  Caccini, Vedro 'l mio sol", {
    x: 0.5, y: 4.35, w: 4.5, h: 0.3, fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=urqSFnKGjIQ" },
  });
  s.addText("NAWM 75  Monteverdi, Cruda Amarilli", {
    x: 0.5, y: 4.65, w: 4.5, h: 0.3, fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=8elXHf0gXTM" },
  });
  s.addText("NAWM 76  Monteverdi, Zefiro torna", {
    x: 0.5, y: 4.95, w: 4.5, h: 0.3, fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=Q3T9pinGCrM" },
  });
}

// ── Save ────────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "Ch13_New_Styles.pptx" })
  .then(() => console.log("Created → Ch13_New_Styles.pptx"))
  .catch(err => console.error(err));
