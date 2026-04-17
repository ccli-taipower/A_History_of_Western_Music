const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 17: Italy and Germany in the Late Seventeenth Century";
pres.author = "A History of Western Music, 10th ed.";

// Warm Baroque palette — sienna / gold / cream
const C = {
  darkBg:   "1E1510",
  gold:     "C8A030",
  cream:    "F5F0E0",
  sienna:   "8B4513",
  copper:   "B87333",
  darkText: "1E1510",
  lightText:"F5F0E0",
  sand:     "E8D8A8",
  slate:    "2A1E14",
  umber:    "3A2510",
  rust:     "A0522D",
};

function darkSlide(pres) { const s = pres.addSlide(); s.background = { color: C.darkBg }; return s; }
function lightSlide(pres) { const s = pres.addSlide(); s.background = { color: C.cream }; return s; }
function accentSlide(pres) { const s = pres.addSlide(); s.background = { color: C.umber }; return s; }
function topBar(s, color) { s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.12, fill: { color: color || C.gold } }); }
function bottomBar(s, color) { s.addShape(pres.ShapeType.rect, { x: 0, y: 5.5, w: "100%", h: 0.125, fill: { color: color || C.gold } }); }

// ── SLIDE 1 · Title ─────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.15, fill: { color: C.gold } });
  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.47, w: "100%", h: 0.155, fill: { color: C.gold } });

  s.addText("A HISTORY OF WESTERN MUSIC  ·  TENTH EDITION", {
    x: 0.5, y: 0.45, w: 9, h: 0.4, fontSize: 18, color: C.sand, charSpacing: 3, align: "center", fontFace: "Georgia",
  });
  s.addText("CHAPTER 17", {
    x: 0.5, y: 1.0, w: 9, h: 0.55, fontSize: 24, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });
  s.addText("ITALY AND GERMANY IN THE\nLATE SEVENTEENTH CENTURY", {
    x: 0.3, y: 1.6, w: 9.4, h: 1.8, fontSize: 36, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.55, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("十七世紀晚期的義大利與德國", {
    x: 0.4, y: 3.7, w: 9.2, h: 0.5, fontSize: 24, color: C.sand, align: "center", fontFace: "Georgia",
  });
  s.addText("Scarlatti · Corelli · Torelli · Buxtehude · Stradivari", {
    x: 0.4, y: 4.3, w: 9.2, h: 0.4, fontSize: 18, color: C.copper, align: "center", fontFace: "Georgia",
  });
  s.addText("Textbook pp. 378–401", {
    x: 0.5, y: 5.0, w: 9, h: 0.35, fontSize: 18, color: C.gold, align: "center", fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 2 · Chapter Overview ──────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.sienna); bottomBar(s, C.sienna);

  s.addText("本章概覽 Chapter Overview", {
    x: 0.4, y: 0.25, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.88, w: 9.2, h: 0.035, fill: { color: C.sand } });

  const items = [
    "義大利歌劇：情感論、詠嘆調發展、Da Capo Aria",
    "Italian Opera: affections, aria development, da capo form",
    "聲樂室內樂：清唱劇 Cantata — Scarlatti, Strozzi",
    "器樂：奏鳴曲 Sonata (chiesa vs. camera) — Corelli",
    "協奏曲：concerto grosso · solo concerto — Torelli",
    "德國音樂生活：宮廷、教會、Collegium musicum",
    "德國聖樂與管風琴音樂：Buxtehude · Toccata · Fugue",
  ];
  items.forEach((t, i) => {
    s.addText(t, {
      x: 0.6, y: 1.1 + i * 0.58, w: 8.8, h: 0.52,
      fontSize: 20, color: i % 2 === 0 ? C.darkText : C.rust, fontFace: "Calibri", valign: "top",
      bullet: i % 2 === 0,
    });
  });
}

// ── SLIDE 3 · Italian Opera — Affections & Arias ────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("義大利歌劇：情感論與詠嘆調", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Italian Opera: The Affections & Aria Development", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.28, w: 7, h: 0.035, fill: { color: C.gold } });

  const bullets = [
    "1660s 受 Descartes 影響 — 情感 (passions) 是穩定、客觀的\n情緒狀態，可透過音樂激發聽者的特定情感反應",
    "詠嘆調取代宣敘調成為歌劇的情感核心\n觀眾愛聽優美獨唱旋律，明星歌手吸引觀眾",
    "詠嘆調數量從 1640 年約 24 首增至 1670s 約 60 首\n常見形式：分節歌、固定低音、ABB'/ABA/ABACA",
    "到世紀末 Da Capo Aria (ABA) 成為主導形式\nB 段對比調性，A 段重複時歌手加入裝飾",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 1.0, w: 9, h: 0.9,
      fontSize: 19, color: C.lightText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 4 · Sartorio & Giulio Cesare (NAWM 93) ────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust); bottomBar(s, C.rust);

  s.addText("Sartorio《乘著凱撒》NAWM 93", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  s.addText("Antonio Sartorio: Giulio Cesare in Egitto (1676)", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.rust, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.28, w: 9.2, h: 0.035, fill: { color: C.sand } });

  const bullets = [
    "威尼斯 Teatro San Salvatore 上演\n以歌劇作為情感衝擊的劇場手段",
    "首位在詠嘆調中使用小號的作曲家\n65 首詠嘆調中含 4 首小號詠嘆調",
    "宣敘調：功能性的，和弦反覆、頻繁離調\nNAWM 93a — 推動劇情的「乾宣敘調」風格",
    "多首 Da Capo Aria (NAWM 93b) 及分節形式\n結合 ABA 與 strophic，展現形式靈活性",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.98, w: 9, h: 0.9,
      fontSize: 19, color: C.darkText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });

  s.addText("https://www.youtube.com/watch?v=iNXNdI-usmA", {
    x: 0.5, y: 5.1, w: 9, h: 0.35, fontSize: 18, color: C.copper, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=iNXNdI-usmA" },
  });
}

// ── SLIDE 5 · Da Capo Aria Structure ────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.copper); bottomBar(s, C.copper);

  s.addText("Da Capo Aria 返始詠嘆調結構", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Forms at a Glance (p. 376)", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.4, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.22, w: 7, h: 0.035, fill: { color: C.gold } });

  // A section box
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.4, w: 4.3, h: 2.8, fill: { color: C.slate }, rounding: true });
  s.addText("A Section 段", { x: 0.6, y: 1.5, w: 4, h: 0.4, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Ritornello (器樂前奏) → I 調\nA1 聲樂陳述 → 離調\nRitornello → 新調\nA2 聲樂陳述 → 回歸主調\nRitornello → I 調", {
    x: 0.6, y: 2.0, w: 4, h: 2.0, fontSize: 19, color: C.lightText, fontFace: "Calibri", lineSpacingMultiple: 1.15, valign: "top",
  });

  // B section box
  s.addShape(pres.ShapeType.rect, { x: 5.3, y: 1.4, w: 4.3, h: 1.3, fill: { color: C.slate }, rounding: true });
  s.addText("B Section 段", { x: 5.5, y: 1.5, w: 4, h: 0.4, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("對比調性 · 第二節詩\n通常無 Ritornello · 較短", {
    x: 5.5, y: 2.0, w: 4, h: 0.6, fontSize: 19, color: C.lightText, fontFace: "Calibri", lineSpacingMultiple: 1.15, valign: "top",
  });

  // A Repeats box
  s.addShape(pres.ShapeType.rect, { x: 5.3, y: 2.9, w: 4.3, h: 1.3, fill: { color: C.slate }, rounding: true });
  s.addText("A Repeats 返始", { x: 5.5, y: 3.0, w: 4, h: 0.4, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("重複 A 段 · 歌手加入裝飾\n展現聲樂技巧與即興能力", {
    x: 5.5, y: 3.5, w: 4, h: 0.6, fontSize: 19, color: C.lightText, fontFace: "Calibri", lineSpacingMultiple: 1.15, valign: "top",
  });

  s.addText("整體形式：A — B — A（Da Capo = 從頭再來）", {
    x: 0.4, y: 4.7, w: 9.2, h: 0.5, fontSize: 22, bold: true, color: C.copper, fontFace: "Georgia", align: "center",
  });
}

// ── SLIDE 6 · Vocal Chamber Music: Cantata ──────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.sienna); bottomBar(s, C.sienna);

  s.addText("義大利清唱劇 Italian Cantata", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  s.addText("Vocal Chamber Music — 聲樂室內樂的精緻藝術", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.rust, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.28, w: 9.2, h: 0.035, fill: { color: C.sand } });

  const bullets = [
    "室內清唱劇：為小型知識菁英聚會而作\n無舞台、無佈景 — 講究精緻、機智與暗語",
    "1690s 標準格式：2-3 對宣敘調+詠嘆調交替\n獨唱 + 通奏低音，8-15 分鐘",
    "文本：田園愛情詩 — 牧羊人的苦戀與甜蜜",
    "代表作曲家：Barbara Strozzi (NAWM 77)\nAlessandro Scarlatti — 超過 600 首清唱劇",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.95, w: 9, h: 0.85,
      fontSize: 19, color: C.darkText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 7 · Scarlatti Cantata: NAWM 94 & 95 ───────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Scarlatti 的清唱劇與歌劇", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("NAWM 94 Clori vezzosa · NAWM 95 La Griselda", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.28, w: 7, h: 0.035, fill: { color: C.gold } });

  // Left column - NAWM 94
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.45, w: 4.5, h: 3.0, fill: { color: C.slate }, rounding: true });
  s.addText("NAWM 94 · Clori vezzosa, e bella", {
    x: 0.5, y: 1.55, w: 4.2, h: 0.4, fontSize: 20, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("• 獨唱清唱劇 ca. 1690-1710\n• 牧羊人向仙女表白熱戀之情\n• 宣敘調：寬廣音域、半音進行\n  減七和弦表達「痛苦的折磨」\n• 詠嘆調 Si, si ben mio：Da Capo\n  吉格舞曲節奏的 Ritornello", {
    x: 0.5, y: 2.05, w: 4.2, h: 2.2, fontSize: 19, color: C.lightText, fontFace: "Calibri", lineSpacingMultiple: 1.1, valign: "top",
  });

  // Right column - NAWM 95
  s.addShape(pres.ShapeType.rect, { x: 5.2, y: 1.45, w: 4.5, h: 3.0, fill: { color: C.slate }, rounding: true });
  s.addText("NAWM 95 · La Griselda (1720-21)", {
    x: 5.4, y: 1.55, w: 4.2, h: 0.4, fontSize: 20, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("• Scarlatti 最後一部歌劇\n• In voler ciò che tu brami\n• A 段展現 Griselda 的尊嚴與\n  溫柔 — 兩段聲樂陳述\n• B 段：堅定的愛——永不停止\n• 更豐富的調性對比與管弦色彩", {
    x: 5.4, y: 2.05, w: 4.2, h: 2.2, fontSize: 19, color: C.lightText, fontFace: "Calibri", lineSpacingMultiple: 1.1, valign: "top",
  });

  s.addText("NAWM 95: https://www.youtube.com/watch?v=PG4Rgg3ujjo", {
    x: 0.5, y: 4.9, w: 9, h: 0.35, fontSize: 18, color: C.copper, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=PG4Rgg3ujjo" },
  });
}

// ── SLIDE 8 · Church Music & Oratorio ───────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust); bottomBar(s, C.rust);

  s.addText("義大利教會音樂與神劇", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  s.addText("Church Music & Oratorio in the Late 17th Century", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.rust, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.28, w: 9.2, h: 0.035, fill: { color: C.sand } });

  const bullets = [
    "延續 Palestrina 式對位法，同時加入通奏低音新風格\nCazzati 在 Bologna San Petronio 教堂 (1657-71)",
    "同一作品可混合舊式對位與新式協奏風格\nMessa a cappella (1670) vs. 華麗的 Magnificat a 4",
    "San Petronio 也是器樂重鎮\nCazzati 出版含小號的奏鳴曲集",
    "神劇：四旬期替代歌劇 · 義大利文 · 分兩部分\n在王公貴族宮殿與學院上演",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.95, w: 9, h: 0.85,
      fontSize: 19, color: C.darkText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 9 · Instrumental Music: Sonata Types ──────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.copper); bottomBar(s, C.copper);

  s.addText("奏鳴曲類型 Sonata Types", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Sonata da chiesa vs. Sonata da camera (ca. 1660)", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.28, w: 7, h: 0.035, fill: { color: C.gold } });

  // Left - Chiesa
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.45, w: 4.5, h: 3.2, fill: { color: C.slate }, rounding: true });
  s.addText("Sonata da chiesa 教堂奏鳴曲", {
    x: 0.5, y: 1.55, w: 4.2, h: 0.4, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("• 抽象樂章 · 無舞曲標題\n• 常含舞曲節奏或二段式\n• 可在教堂禮儀中使用\n  (替代 Mass Proper 或晚禱)\n• 也可在私人音樂會演出", {
    x: 0.5, y: 2.1, w: 4.2, h: 2.3, fontSize: 20, color: C.lightText, fontFace: "Calibri", lineSpacingMultiple: 1.15, valign: "top",
  });

  // Right - Camera
  s.addShape(pres.ShapeType.rect, { x: 5.2, y: 1.45, w: 4.5, h: 3.2, fill: { color: C.slate }, rounding: true });
  s.addText("Sonata da camera 室內奏鳴曲", {
    x: 5.4, y: 1.55, w: 4.2, h: 0.4, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("• 一系列風格化舞曲樂章\n• 通常以前奏曲開始\n• 舞曲標題：allemande,\n  courante, sarabande, gigue\n• 娛樂用途，私人場合", {
    x: 5.4, y: 2.1, w: 4.2, h: 2.3, fontSize: 20, color: C.lightText, fontFace: "Calibri", lineSpacingMultiple: 1.15, valign: "top",
  });

  s.addText("1670 後最常見編制：兩把高音樂器 + 通奏低音 = Trio Sonata 三重奏鳴曲", {
    x: 0.4, y: 4.9, w: 9.2, h: 0.45, fontSize: 20, bold: true, color: C.copper, fontFace: "Georgia", align: "center",
  });
}

// ── SLIDE 10 · Arcangelo Corelli — Life ─────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.sienna); bottomBar(s, C.sienna);

  s.addText("Arcangelo Corelli (1653-1713)", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  s.addText("柯瑞里 — 巴洛克器樂之父", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 24, color: C.rust, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.28, w: 9.2, h: 0.035, fill: { color: C.sand } });

  const bullets = [
    "生於 Fusignano · Bologna 學琴 · 1675 定居羅馬\n小提琴家、教師、樂團指揮，深受貴族贊助",
    "不寫聲樂——幾乎全是器樂：三重奏鳴曲、\n獨奏奏鳴曲、大協奏曲 · 6 部作品集 (Op.1-6)",
    "影響深遠：Handel · Bach · Telemann · Couperin\n被稱為「首位聲譽完全建立在器樂上的大作曲家」",
    "確立了形式、風格與和聲語言的標準\n奠定 18 世紀初器樂音樂的國際語言",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.98, w: 9, h: 0.9,
      fontSize: 19, color: C.darkText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 11 · Corelli's Trio Sonatas ───────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Corelli 三重奏鳴曲", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Trio Sonatas: Op. 1-4 · 48 Works", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.28, w: 7, h: 0.035, fill: { color: C.gold } });

  const bullets = [
    "教堂奏鳴曲 (Op.1, 3)：四樂章 慢-快-慢-快\n抒情對話勝過炫技 · 兩把小提琴常交叉與交換",
    "Walking bass 行進低音 · 掛留音鏈 · 自由模仿\n兩把小提琴如對話般向上攀升",
    "室內奏鳴曲 (Op.2, 4)：前奏曲 + 舞曲\n首兩樂章常似教堂奏鳴曲 · 舞曲為二段式",
    "每個樂章基於一個主題展開，環環相扣\n各樂章以同一動機的變體開始 (Op.3 No.2)",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.98, w: 9, h: 0.9,
      fontSize: 19, color: C.lightText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 12 · NAWM 96 Corelli Trio Sonata ──────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust); bottomBar(s, C.rust);

  s.addText("NAWM 96 Corelli Trio Sonata", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  s.addText("Op. 3, No. 2 in D Major — 教堂奏鳴曲分析", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.rust, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.28, w: 9.2, h: 0.035, fill: { color: C.sand } });

  const mvts = [
    "I. Grave：對位織體 · 莊嚴肅穆\n掛留音鏈 · 行進低音 · 小提琴自由模仿",
    "II. Allegro：快板 · 模仿主題 · 低音全程參與\n保留 canzona 元素 · 教堂奏鳴曲的重心",
    "III. 慢板：如歌式 · 三拍子 · 關係小調\n弗里吉亞終止 · 適合即興裝飾",
    "IV. Allegro 終曲：模仿式吉格 · 二段式\n主題反轉 · 對位手法如第二樂章",
  ];
  mvts.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.95, w: 9, h: 0.85,
      fontSize: 19, color: C.darkText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });

  s.addText("https://www.youtube.com/watch?v=wH64J5f-DHY", {
    x: 0.5, y: 5.1, w: 9, h: 0.35, fontSize: 18, color: C.copper, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=wH64J5f-DHY" },
  });
}

// ── SLIDE 13 · Corelli's Solo Sonatas ───────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.copper); bottomBar(s, C.copper);

  s.addText("Corelli 獨奏奏鳴曲 Solo Sonatas", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Op. 5 (1700) — 12 Violin Sonatas", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.28, w: 7, h: 0.035, fill: { color: C.gold } });

  const bullets = [
    "1700 年後獨奏奏鳴曲人氣超越三重奏鳴曲\n允許更多炫技：雙音、三和弦、快速跑句",
    "快樂章中模擬三重奏鳴曲的三聲部織體\n賦格手法、永動式樂句、華彩段",
    "慢樂章記譜簡潔，實際演奏大量裝飾\n1710 年 Roger 出版裝飾版本 (Fig. 17.6)",
    "義大利式裝飾風格 — 不同於法式 agrements\n根源在 16 世紀義大利即興傳統",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.98, w: 9, h: 0.9,
      fontSize: 19, color: C.lightText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 14 · Corelli's Tonal Organization & Influence ─────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.sienna); bottomBar(s, C.sienna);

  s.addText("Corelli 的調性語言與影響", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  s.addText("Tonal Organization & Legacy", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.rust, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.28, w: 9.2, h: 0.035, fill: { color: C.sand } });

  const bullets = [
    "Rameau 以 Corelli 的音樂作為描述調性系統的基礎\n五度圈進行 · 掛留音鏈推動和聲向前",
    "幾乎全為自然音階 · 偶用減七或拿坡里六和弦\n轉調至屬調或關係調 — 邏輯而清晰",
    "每個樂章統一於單一調性\n或同調、或大調奏鳴曲第二慢板用關係小調",
    "Vivaldi · Handel · Bach · Telemann 延續其技法\n與 Frescobaldi 並列首批純器樂經典的創造者",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.98, w: 9, h: 0.9,
      fontSize: 19, color: C.darkText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 15 · The Concerto — Types ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("協奏曲的誕生 The Concerto", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Three Types by ca. 1700 · 三種協奏曲類型", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.28, w: 7, h: 0.035, fill: { color: C.gold } });

  // Three boxes
  const types = [
    { title: "Orchestral Concerto\n管弦協奏曲", desc: "強調第一小提琴與低音\n的對比 · 區別於奏鳴曲\n的對位織體" },
    { title: "Concerto Grosso\n大協奏曲", desc: "Concertino (小組獨奏)\nvs. Ripieno (大樂團)\n如擴充版三重奏鳴曲" },
    { title: "Solo Concerto\n獨奏協奏曲", desc: "一位獨奏者 vs. 弦樂團\n最常見、最重要的類型\nTutti / Ripieno = 全奏" },
  ];
  types.forEach((t, i) => {
    const x = 0.3 + i * 3.2;
    s.addShape(pres.ShapeType.rect, { x, y: 1.5, w: 3.0, h: 3.2, fill: { color: C.slate }, rounding: true });
    s.addText(t.title, { x: x + 0.15, y: 1.6, w: 2.7, h: 0.85, fontSize: 20, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
    s.addShape(pres.ShapeType.rect, { x: x + 0.3, y: 2.55, w: 2.4, h: 0.03, fill: { color: C.copper } });
    s.addText(t.desc, { x: x + 0.15, y: 2.7, w: 2.7, h: 1.8, fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top", align: "center", lineSpacingMultiple: 1.15 });
  });
}

// ── SLIDE 16 · Corelli's Concerti Grossi ────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust); bottomBar(s, C.rust);

  s.addText("Corelli 大協奏曲 Op. 6", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  s.addText("Concerti Grossi — 1680s 作曲，1714 出版", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.rust, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.28, w: 9.2, h: 0.035, fill: { color: C.sand } });

  const bullets = [
    "羅馬樂團以 concertino (小組) vs. ripieno (大組) 為基礎\nCorelli 經常指揮 40 人以上的「臨時」大樂團",
    "本質上是三重奏鳴曲的擴充版\n大組呼應小組 · 強化終止 · 加厚音響",
    "Georg Muffat 描述如何將奏鳴曲改為協奏曲\n3 人演奏 → 4-5 人 → 加 ripieno 成完整協奏曲",
    "Corelli 的手法被 Italy · England · Germany 廣泛模仿",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.95, w: 9, h: 0.85,
      fontSize: 19, color: C.darkText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 17 · Torelli & Ritornello Form (NAWM 97) ──────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Torelli 與 Ritornello 形式", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Giuseppe Torelli (1658-1709) · NAWM 97", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.28, w: 7, h: 0.035, fill: { color: C.gold } });

  const bullets = [
    "Bologna 學派領袖 · 作三種類型協奏曲\nOp. 5 (1692) 首批出版協奏曲 · Op. 8 (1709)",
    "三樂章格式：快-慢-快（源自義大利歌劇序曲）\nAlbinoni Op. 2 (1700) 確立此標準",
    "快樂章常用 Ritornello Form 反覆段形式\n結構類似 Da Capo Aria 的 A 段",
    "Ritornello (全奏) 框住兩段獨奏段落\n獨奏展現全新材料 · 轉調至近關係調\nRitornello 回歸帶來穩定感與終止感",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 1.0, w: 9, h: 0.9,
      fontSize: 19, color: C.lightText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });

  s.addText("https://www.youtube.com/watch?v=URt-9qZZxPU", {
    x: 0.5, y: 5.1, w: 9, h: 0.35, fontSize: 18, color: C.copper, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=URt-9qZZxPU" },
  });
}

// ── SLIDE 18 · The Italian Style — Summary ──────────────────────────────────
{
  const s = accentSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("義大利風格總結 The Italian Style", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("17 世紀後三分之一 · 共同特徵", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.28, w: 7, h: 0.035, fill: { color: C.gold } });

  const bullets = [
    "音樂悅耳、情感豐富、展現演奏者魅力\n從抒情歌唱到炫技華彩的多樣旋律風格",
    "獨奏者的突出地位 — 從世紀初延續至今\n詠嘆調、獨奏奏鳴曲、協奏曲皆強調個人",
    "調性 (tonality) 成為強大的組織力量\n建立主調→離調→回歸主調的基本模式",
    "Reprise (再現) 原則 — 回歸開頭材料\n成為此後兩個世紀的基本形式原則",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.98, w: 9, h: 0.9,
      fontSize: 19, color: C.lightText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 19 · Germany & Austria — Musical Life ─────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.sienna); bottomBar(s, C.sienna);

  s.addText("德國與奧地利的音樂生活", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  s.addText("Germany & Austria: Courts, Cities, Churches", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.rust, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.28, w: 9.2, h: 0.035, fill: { color: C.sand } });

  const bullets = [
    "1648 三十年戰爭結束 — 德國分裂為約 300 個政治單位\n小邦林立 · 城市較小 · 但音樂卻蓬勃發展",
    "宮廷：仿效 Louis XIV 以藝術彰顯權威\n聘請歌手、器樂家、作曲家",
    "城市：Stadtpfeifer (市鎮吹奏手) 負責公共音樂\n教會音樂家由市議會任命 · 路德教會贊助音樂會",
    "業餘音樂：Collegium musicum 中產階級音樂社團\n學校與大學也組織 · 18 世紀發展為公開音樂會",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.98, w: 9, h: 0.9,
      fontSize: 19, color: C.darkText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 20 · German Opera & Song ──────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.copper); bottomBar(s, C.copper);

  s.addText("德國歌劇與歌曲", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("German Opera, Song & Cantata", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.28, w: 7, h: 0.035, fill: { color: C.gold } });

  const bullets = [
    "義大利歌劇主導宮廷 — Pallavicino (Dresden)\nSteffani (Munich, Hanover) 等義大利人活躍",
    "1678 Hamburg 首家德語公共歌劇院開幕\n商業運營 · 面向中產 · 早期取材聖經題材",
    "德語歌劇折衷混合：Da Capo Aria + 法式風格\n下層角色用分節歌 · Reinhard Keiser 寫 60 部",
    "Adam Krieger (1634-1666)：最重要歌曲作曲家\n分節式旋律配管弦 ritornello · 清唱劇亦然",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.98, w: 9, h: 0.9,
      fontSize: 19, color: C.lightText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 21 · German Sacred Music ──────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust); bottomBar(s, C.rust);

  s.addText("德國教會音樂 Sacred Music", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  s.addText("Catholic South · Lutheran North", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.rust, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.28, w: 9.2, h: 0.035, fill: { color: C.sand } });

  const bullets = [
    "天主教南部 (Munich · Salzburg · Vienna)\n皇帝支持 · 複合唱 · 管弦前奏 · Biber Missa (1682)",
    "路德教會：正統派 vs. 敬虔派的衝突\n正統派擁護豐富禮儀音樂 · 敬虔派偏好簡樸",
    "路德宗三大元素：協奏聲樂(聖經) + 獨唱詠嘆調\n(義大利風格) + 聖詠 Chorale (德國傳統)",
    "Buxtehude 的 Wachet auf：聖詠變奏曲範例\n每段旋律重新改編 · 各段風格各異",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.95, w: 9, h: 0.85,
      fontSize: 19, color: C.darkText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 22 · Buxtehude ────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Dieterich Buxtehude (ca. 1637-1707)", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("布克斯特胡德 — 北德管風琴大師", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 20, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.28, w: 7, h: 0.035, fill: { color: C.gold } });

  const bullets = [
    "Lubeck St. Mary's Church 管風琴師\n以管風琴曲與聖樂著稱 — 約 120 首聖樂作品",
    "Abendmusiken 晚間音樂會：聖誕前五個週日\n免費公開 · 吸引全德音樂家 · 巴赫步行百哩聆聽",
    "管風琴作品：約 40 首聖詠編曲 · 22 首前奏/觸技曲\n19 首大鍵琴組曲 · 20 首室內奏鳴曲",
    "對年輕 J.S. Bach (1705) 產生深遠影響\n1705 年巴赫從 Arnstadt 步行至 Lubeck 聆聽",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.98, w: 9, h: 0.9,
      fontSize: 19, color: C.lightText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 23 · Lutheran Organ Music: Toccata & Fugue ────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.sienna); bottomBar(s, C.sienna);

  s.addText("路德教會管風琴音樂", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  s.addText("Toccata, Fugue & Chorale Settings", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.rust, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.28, w: 9.2, h: 0.035, fill: { color: C.sand } });

  const bullets = [
    "1650-1750 管風琴音樂的黃金時代\n北方：Buxtehude · Bohm · 中部：J.C. Bach · Pachelbel",
    "Toccata 觸技曲：自由段落 vs. 賦格段落交替\nBuxtehude Praeludium in E 含 5 個自由段 + 4 個賦格段",
    "Fugue 賦格：獨立作品或觸技曲的一部分\n主題 (subject) → 答題 (answer) → 插句 (episode)",
    "管風琴建造藝術：Schnitger · Silbermann\nHauptwerk · Ruckpositiv · Brustwerk · 腳鍵盤",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.95, w: 9, h: 0.85,
      fontSize: 19, color: C.darkText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 24 · Chorale Settings ─────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.copper); bottomBar(s, C.copper);

  s.addText("聖詠編曲 Chorale Settings", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Organ Chorales · Chorale Preludes · Chorale Variations", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.28, w: 7, h: 0.035, fill: { color: C.gold } });

  const bullets = [
    "Organ Chorale：以和聲與對位豐富聖詠旋律\nChorale Variations (partita) / Chorale Fantasia",
    "Chorale Prelude 聖詠前奏曲：完整旋律出現一次\n四種手法：各段模仿 / 長音定旋律 / 裝飾旋律 / 動機伴奏",
    "Buxtehude: Nun komm, der Heiden Heiland\n旋律加裝飾音 · 短顫音 · 鄰音 · 義大利式華彩",
    "德國作曲家融合法/義/德三國風格\n為 18 世紀 Bach 的輝煌管風琴音樂奠基",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.98, w: 9, h: 0.9,
      fontSize: 19, color: C.lightText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 25 · Other German Instrumental Music ──────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust); bottomBar(s, C.rust);

  s.addText("德國其他器樂 Other Instrumental", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  s.addText("Violin Sonatas · Keyboard Sonatas · Suites", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 22, color: C.rust, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.28, w: 9.2, h: 0.035, fill: { color: C.sand } });

  const bullets = [
    "小提琴奏鳴曲：Biber Mystery Sonatas (ca.1675)\n特殊定弦 (scordatura) · Walther Scherzi (1676)",
    "鍵盤奏鳴曲：Kuhnau Frische Clavier Fruchte (1696)\n首批鍵盤奏鳴曲 · Biblical Sonatas (1700) 標題音樂",
    "大鍵琴組曲：Froberger 傳入法式風格\nAllemande-Courante-Sarabande-Gigue 成標準",
    "管弦組曲 (Orchestral Suite)：1690-1740 德國時尚\nMuffat Florilegium (1695, 1698) 引入 Lully 風格",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.95, w: 9, h: 0.85,
      fontSize: 19, color: C.darkText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 26 · Stradivari & Violin Making ───────────────────────────────────
{
  const s = accentSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Stradivari 與製琴藝術", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Music in Context: The Stradivarius Violin Workshop (p.378)", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.45, fontSize: 20, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.28, w: 7, h: 0.035, fill: { color: C.gold } });

  const bullets = [
    "Cremona 製琴家族：Amati · Stradivari · Guarneri\n17-18 世紀義大利小提琴製作達到巔峰",
    "Antonio Stradivari (ca.1644-1737)：超過 1100 件樂器\n含豎琴、吉他、中提琴、大提琴 · 約半數存世",
    "選材精良 · 弧度精準 · F孔優雅 · 橙棕色亮漆\n漆的秘密至今未解 — 科學家無法測出特殊成分",
    "現代小提琴已改裝：指板加長、琴橋升高、琴弦改鋼\n極少數 Strad 保留原始 17 世紀形態",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.45 + i * 0.98, w: 9, h: 0.9,
      fontSize: 19, color: C.lightText, fontFace: "Calibri", bullet: true, lineSpacingMultiple: 1.05, valign: "top",
    });
  });
}

// ── SLIDE 27 · Timeline ─────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("大事年表 Timeline", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 0.82, w: 7, h: 0.035, fill: { color: C.copper } });

  const eventsL = [
    "1647  Cruger, Praxis pietatis melica",
    "1648  三十年戰爭結束",
    "1668  Buxtehude 任職 Lubeck",
    "1675  Biber, Mystery Sonatas",
    "1676  Sartorio, Giulio Cesare",
    "1678  Hamburg 歌劇院開幕",
    "1681  Corelli Op.1 出版",
  ];
  const eventsR = [
    "1682  Biber, Missa salisburgensis",
    "1689  Corelli Op.3 出版",
    "1692  Torelli 首批協奏曲",
    "1696  Kuhnau 鍵盤奏鳴曲",
    "1700  Corelli Op.5 (12 Violin Sonatas)",
    "1705  Bach 聆聽 Buxtehude",
    "1720  Scarlatti, La Griselda",
  ];

  eventsL.forEach((t, i) => {
    s.addText(t, {
      x: 0.3, y: 1.0 + i * 0.58, w: 4.7, h: 0.5,
      fontSize: 18, color: i % 2 === 0 ? C.gold : C.sand, fontFace: "Calibri", valign: "top",
    });
  });
  eventsR.forEach((t, i) => {
    s.addText(t, {
      x: 5.2, y: 1.0 + i * 0.58, w: 4.7, h: 0.5,
      fontSize: 18, color: i % 2 === 0 ? C.gold : C.sand, fontFace: "Calibri", valign: "top",
    });
  });
}

// ── SLIDE 28 · Key Terms & Listening Guide ──────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.sienna); bottomBar(s, C.sienna);

  s.addText("關鍵術語與聆聽指南", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  s.addText("Key Terms & Listening Guide", {
    x: 0.4, y: 0.8, w: 9.2, h: 0.4, fontSize: 24, color: C.rust, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.22, w: 9.2, h: 0.035, fill: { color: C.sand } });

  // Left - Terms
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.4, w: 4.5, h: 3.6, fill: { color: C.sand }, rounding: true });
  s.addText("Key Terms 關鍵術語", {
    x: 0.5, y: 1.5, w: 4.2, h: 0.4, fontSize: 22, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  s.addText("Da Capo Aria 返始詠嘆調\nSonata da chiesa / da camera\nTrio Sonata 三重奏鳴曲\nConcerto Grosso 大協奏曲\nRitornello Form 反覆段形式\nCollegium musicum 音樂社團\nToccata 觸技曲 / Fugue 賦格\nChorale Prelude 聖詠前奏曲\nWalking Bass 行進低音", {
    x: 0.5, y: 2.0, w: 4.2, h: 2.8, fontSize: 18, color: C.darkText, fontFace: "Calibri", lineSpacingMultiple: 1.2, valign: "top",
  });

  // Right - Listening
  s.addShape(pres.ShapeType.rect, { x: 5.2, y: 1.4, w: 4.5, h: 3.6, fill: { color: C.sand }, rounding: true });
  s.addText("Listening 聆聽作業", {
    x: 5.4, y: 1.5, w: 4.2, h: 0.4, fontSize: 22, bold: true, color: C.sienna, fontFace: "Georgia",
  });
  const ch17links = [
    ["NAWM 93 Sartorio, Giulio Cesare", ""],
    ["NAWM 94 Scarlatti, Clori vezzosa", ""],
    ["NAWM 95 Scarlatti, La Griselda", "https://www.youtube.com/watch?v=_M62Dr5LmmY"],
    ["NAWM 96 Corelli, Trio Sonata Op.3/2", ""],
    ["NAWM 97 Torelli, Concerto Op.8/8", ""],
  ];
  ch17links.forEach(([label, url], i) => {
    const opts = { x: 5.4, y: 2.0 + i * 0.5, w: 4.2, h: 0.45, fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "top" };
    if (url) { opts.hyperlink = { url }; opts.color = C.sienna; }
    s.addText(label, opts);
  });
}

// ── Save ────────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "Ch17_Italy_Germany.pptx" })
  .then(() => console.log("Ch17_Italy_Germany.pptx created successfully!"))
  .catch(err => console.error(err));
