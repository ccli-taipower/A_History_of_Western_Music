const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 17: Italy and Germany in the Late Seventeenth Century";
pres.author = "A History of Western Music, 10th ed.";

const C = {
  darkBg:   "1E1510",
  gold:     "C8A030",
  cream:    "F5F0E0",
  sienna:   "6A3A1A",
  copper:   "B87333",
  darkText: "1E1510",
  lightText:"F5F0E0",
  sand:     "E8D8A8",
  slate:    "2A2018",
  amber:    "D4A030",
  brown:    "4A2A1A",
};

function darkSlide(p) { const s = p.addSlide(); s.background = { color: C.darkBg }; return s; }
function lightSlide(p) { const s = p.addSlide(); s.background = { color: C.cream }; return s; }
function topBar(s, c) { s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.12, fill: { color: c || C.gold } }); }
function bottomBar(s, c) { s.addShape(pres.ShapeType.rect, { x: 0, y: 5.5, w: "100%", h: 0.125, fill: { color: c || C.gold } }); }

// ── SLIDE 1 · Title ──────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.15, fill: { color: C.gold } });
  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.47, w: "100%", h: 0.155, fill: { color: C.gold } });
  s.addText("A HISTORY OF WESTERN MUSIC · TENTH EDITION", { x: 0.5, y: 0.45, w: 9, h: 0.35, fontSize: 11, color: C.sand, charSpacing: 3, align: "center", fontFace: "Georgia" });
  s.addText("CHAPTER 17", { x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6 });
  s.addText("ITALY AND GERMANY IN\nTHE LATE SEVENTEENTH CENTURY", { x: 0.3, y: 1.5, w: 9.4, h: 2.0, fontSize: 30, color: C.lightText, bold: true, align: "center", fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.65, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Da Capo Aria · Scarlatti · Corelli · Concerto · Buxtehude", { x: 0.4, y: 3.8, w: 9.2, h: 0.4, fontSize: 13, color: C.sand, align: "center", fontFace: "Georgia" });
  s.addText("Textbook pp. 371–399", { x: 0.5, y: 4.8, w: 9, h: 0.3, fontSize: 11, color: C.gold, align: "center", fontFace: "Calibri" });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.sienna); bottomBar(s, C.sienna);
  s.addText("本章概覽 Chapter Overview", { x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.sienna, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });
  const sections = [
    ["🎭", "Italian Opera 義大利歌劇", "Da capo aria · Sartorio · Giulio Cesare in Egitto (NAWM 93)"],
    ["🎤", "Vocal Chamber Music 聲樂室內樂", "Cantata · Scarlatti · Clori vezzosa (NAWM 94) · serenata"],
    ["🎻", "Sonata 奏鳴曲", "Trio sonata · solo sonata · Corelli Op. 3 (NAWM 96) · 調性建立"],
    ["🎺", "Concerto 協奏曲", "Concerto grosso · orchestral concerto · Torelli · ritornello form"],
    ["⛪", "Germany: Sacred Music 德國宗教音樂", "Buxtehude · chorale concerto · 正統 vs 敬虔主義"],
    ["🎹", "Germany: Organ & Keyboard 德國管風琴", "Toccata · fugue · chorale prelude · Pachelbel · Biber"],
  ];
  sections.forEach(([icon, title, desc], i) => {
    const y = 1.0 + i * 0.72;
    s.addShape(pres.ShapeType.rect, { x: 0.45, y, w: 9.1, h: 0.62, fill: { color: i % 2 === 0 ? "E8E0D0" : "DED6C6" }, rounding: true });
    s.addText(icon, { x: 0.55, y, w: 0.5, h: 0.62, fontSize: 20, align: "center" });
    s.addText(title, { x: 1.1, y, w: 2.6, h: 0.62, fontSize: 13, bold: true, color: C.sienna, fontFace: "Georgia", valign: "middle" });
    s.addText(desc, { x: 3.8, y, w: 5.65, h: 0.62, fontSize: 10.5, color: C.darkText, fontFace: "Calibri", valign: "middle" });
  });
}

// ── SLIDE 3 · Italian Opera & Da Capo Aria ──────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("🎭 義大利歌劇與返始詠嘆調 Italian Opera & Da Capo Aria", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 20, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.copper } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Opera in the Late 1600s", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("情感觀的轉變 New View of the Affections：\n• Descartes 影響：情感 = 客觀、穩定的心理狀態\n• 非個人化 → 可被音樂精確捕捉\n• 宣敘調 / 詠嘆調分工更清晰：\n  — 宣敘調推動劇情（多為功能性 secco）\n  — 詠嘆調表達情感（抒情高潮）\n\nAria 的發展：\n• 1640 ca. 24首 → 1670s ca. 60首 / 歌劇\n• 常見形式：strophic · ground bass · ABA\n  ABB' · ABACA · rondo\n• Da Capo Aria 返始詠嘆調成為主流：\n  — ABA 形式：B 段末標 \"Da capo\"\n  — 返回 A 段時加入新裝飾 → 展現歌藝\n  — 18世紀歌劇與清唱劇的標準形式\n\nDa Capo 的結構特徵：\n  A: Rit → A1 (I→V) → Rit → A2 (V→I) → Rit\n  B: 對比調性 / 情感\n  A: 重複（加裝飾）", {
    x: 0.45, y: 1.45, w: 4.3, h: 3.6, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Sartorio: Giulio Cesare (1676 · NAWM 93)", { x: 5.2, y: 1.08, w: 4.4, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Antonio Sartorio (1630–1680)：\n• 威尼斯 Teatro San Salvatore\n• 歌劇 = 情感震撼的戲劇娛樂\n• 首位在詠嘆調中使用小號的作曲家\n\nGiulio Cesare in Egitto Act II, Sc. 3–4：\n• Cleopatra 偽裝為 Lidia 向 Cesare 求愛\n\n(93a) Recitative — 功能性：和弦式、\n  主要用重複音與常見和弦推動劇情\n\n(93b) Se qualcuna mi bramasse — Nireno\n  — Da capo aria (ABB'A)\n\n(93c) Son prigioniero — Cesare\n  — ABA 形式 + 詩節內 ABA 嵌套\n\n(93d) Alla carcere d'un crine — Cleopatra\n  — 台下歌唱 → Cesare 被音樂迷惑\n  — 三段式：每段前後都有宣敘調\n  — 營造「偷聽」的戲劇效果\n\n(93e) Alla carcere — Cesare 回應版", {
    x: 5.25, y: 1.45, w: 4.3, h: 3.6, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 4 · Scarlatti & Cantata ───────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("🎤 A. Scarlatti 與清唱劇 Scarlatti & the Cantata", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.copper } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Cantata 清唱劇", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("羅馬 = 清唱劇中心：\n• 貴族/外交官的 academy 聚會\n• 為小型知識分子觀眾而寫\n• 追求精緻、機智、微妙\n\nCantata 約 1690s 標準形式：\n• 2–3 組 recitative + aria 交替\n• 獨唱 + 通奏低音（偶用小型樂團）\n• 8–15 分鐘\n• 田園愛情詩 → bittersweet 風格\n\nAlessandro Scarlatti (1660–1725)：\n• 600+ 清唱劇 — 這一體裁的巔峰\n• 活躍於羅馬與那不勒斯\n\nClori vezzosa, e bella (ca. 1690–1710)：\n• 牧歌風：牧羊人傾訴愛慕之苦\n\n(NAWM 94a) 第二段宣敘調：\n• 半音進行、七和弦、遠關係轉調\n• \"affanni miei\" → F 小調三和弦（哀苦）\n• 減七和弦 on \"il martire\" → 一語雙關", {
    x: 0.45, y: 1.45, w: 4.3, h: 3.6, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Da Capo Aria & La Griselda", { x: 5.2, y: 1.08, w: 4.4, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("(NAWM 94b) Sì, sì ben mio — Da Capo Aria：\n• 吉格節奏的 ritornello\n• A 段諷刺式：\"更多折磨給我的心\"\n• B 段轉大調 → 更樂觀的情感\n• 缺少 A 段末 ritornello\n\nScarlatti 晚期 da capo aria 特徵：\n• A 段兩次歌唱陳述 (A1, A2)\n  — A1: I → V (轉調)\n  — A2: V → I (回歸)\n• 每次陳述前後有器樂 ritornello\n• B 段：1–2 次歌唱，通常無 ritornello\n• 整體更宏大、對比更豐富\n\n(NAWM 95) La Griselda: \"In voler ciò che\ntu brami\" (1720–21)：\n• Scarlatti 最後一部歌劇\n• A/B 段用不同素材\n  — 唱 ≠ ritornello 旋律\n• 順從妻子的尊嚴 vs 不屈的愛\n\nSerenata 小夜曲：\n• 介於清唱劇與歌劇之間\n• Stradella (1639–1682) 先驅", {
    x: 5.25, y: 1.45, w: 4.3, h: 3.6, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 5 · Corelli & the Sonata ──────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("🎻 柯瑞里與奏鳴曲 Corelli & the Sonata", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.copper } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Sonata Types 奏鳴曲分類", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("約 1660 年確立兩大類型：\n\nSonata da chiesa 教堂奏鳴曲：\n• 多為抽象樂章（無舞曲標題）\n• 常含一首以上使用舞曲節奏的樂章\n• 可用於教堂禮拜\n\nSonata da camera 室內奏鳴曲：\n• 序列舞曲樂章，常有前奏曲\n• 供私人娛樂\n\nTrio Sonata 三重奏鳴曲：\n• 三聲部織體：2 高音旋律 + 通奏低音\n• 實際需要 4 人：2 小提琴 + 大提琴 + 鍵盤\n\nSolo Sonata 獨奏奏鳴曲：\n• 獨奏樂器 + 通奏低音\n• 1700 年後漸增；技巧要求更高\n\nCremona 製琴黃金時代：\n• Amati (1596–1684)\n• Stradivari (1644–1737) → 1,100+ 樂器\n• Guarneri (1698–1744)\n→ 小提琴成為器樂之王", {
    x: 0.45, y: 1.45, w: 4.3, h: 3.6, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Arcangelo Corelli (1653–1713)", { x: 5.2, y: 1.08, w: 4.4, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Bologna 學琴 → Rome 定居\n• 小提琴家 / 教師 / 指揮 / 作曲家\n• 第一位幾乎純以器樂立名的大作曲家\n• 6 套出版作品 (Opp. 1–6)：\n  Op.1 (1681): 12 trio da chiesa\n  Op.2 (1685): 12 trio da camera\n  Op.3 (1689): 12 trio da chiesa\n  Op.4 (1695): 12 trio da camera\n  Op.5 (1700): 12 solo violin sonatas\n  Op.6 (1714): 12 concerti grossi\n\nTrio Sonata Op. 3, No. 2 in D (NAWM 96)：\n教堂奏鳴曲典型四樂章 Slow–Fast–Slow–Fast\n\n(96a) Grave — 對位織體，walking bass\n  — 掛留音鏈 → 強力前進動力\n(96b) Allegro — 模仿式，如 canzona\n(96c) Adagio — 歌唱性，關係小調\n(96d) Allegro — 二段式 gigue\n\n調性特徵：功能和聲、五度圈進行\n→ Rameau 以 Corelli 為調性理論基礎", {
    x: 5.25, y: 1.45, w: 4.3, h: 3.6, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 6 · The Concerto ──────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("🎺 協奏曲的誕生 The Birth of the Concerto", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.copper } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Three Types of Concerto 三種協奏曲", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("1680–90s 義大利作曲家創造管弦樂新類型：\n\n① Orchestral Concerto 管弦樂協奏曲：\n  — 強調第一小提琴 + 低音\n  — 區別於奏鳴曲的對位織體\n\n② Concerto Grosso 大協奏曲：\n  — Concertino（獨奏群，通常 = trio sonata）\n    vs. Ripieno / Tutti（全體樂團）\n  — 本質上 = 合奏奏鳴曲 + 段落加倍\n  — Corelli Op. 6 (1714)：12 首經典\n\n③ Solo Concerto 獨奏協奏曲：\n  — 1+ 獨奏樂器 vs. 弦樂團\n  — 最常見、影響最深遠的類型\n\n前身：\n• Lully 歌劇中 solo vs. tutti 對比\n• Stradella 清唱劇中 concertino 段落\n• Bologna 教會奏鳴曲加入小號\n• 自然小號 → triads + 音階 + 重複音", {
    x: 0.45, y: 1.45, w: 4.3, h: 3.6, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Torelli & Ritornello Form", { x: 5.2, y: 1.08, w: 4.4, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Georg Muffat (1653–1704)：\n• 德奧作曲家，早期到 Rome 聽 Corelli\n• 在 Salzburg 出版 Florilegium (1695, 1698)\n  — 法國管弦樂 suite 引入德國\n• 1701 出版 concerti grossi 並附序言\n  — 解說如何將奏鳴曲改編為協奏曲\n  — 重要的第一手文獻 (Source Reading)\n\nGiuseppe Torelli (1658–1709)：\n• Bologna San Petronio\n• 出版最早的協奏曲集 (Op. 5, 1692)\n• Op. 6 (1698)：可能最早的小提琴獨奏協奏曲\n• Op. 8 (1709)：6 concerti grossi + 6 violin concertos\n\nRitornello Form 利托內洛形式：\n• 仿 da capo aria A 段結構\n• Ritornello (tutti)：主題，首尾在主調\n  — 中間在不同調出現（簡化版）\n• Solo episodes：新素材，展現技巧\n  — 轉調，製造對比\n• Torelli → Vivaldi 發展成熟\n  （見 Ch. 18）\n\nTomaso Albinoni (1671–1750)：\n• Op. 2 (1700)：確立 Fast–Slow–Fast 三樂章", {
    x: 5.25, y: 1.45, w: 4.3, h: 3.6, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 7 · Germany: Opera, Song & Sacred Music ───────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("🇩🇪 德奧：歌劇、歌曲與教會音樂", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.copper } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("German Musical Life", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("三十年戰爭後 (1648)：\n• 神聖羅馬帝國 = ~300 獨立政治單元\n• 宮廷、城市、教會各自雇用音樂家\n\n三類音樂家：\n① Court musicians 宮廷\n② Stadtpfeifer 城鎮樂手\n   — 專屬權利 + 師徒制\n   — Bach 家族由此傳統培育\n③ Church musicians 教會\n   — Lutheran 教會直接雇用\n\nCollegium musicum 音樂社團：\n• 中產階級 + 大學生業餘合奏\n• 18世紀轉型為公共音樂會\n\nOpera in German 德語歌劇：\n• Hamburg (1678) 首家德語歌劇院\n• 商業取向，吸引中產階級\n• Reinhard Keiser (1674–1739)：\n  — ~60部歌劇 → 最多產的早期德語歌劇家\n  — 折衷風格：da capo + French airs + 德國民歌\n\nSong: Adam Krieger (1634–1666)\n  — 簡潔旋律 + 短管弦序奏", {
    x: 0.45, y: 1.45, w: 4.3, h: 3.6, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Sacred Music 宗教音樂", { x: 5.2, y: 1.08, w: 4.4, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Catholic South 天主教南方：\n• Munich · Salzburg · Vienna\n• 四位皇帝 (1637–1740) 支持且創作音樂\n• 融合 stile antico + 新 concertato\n  — 管弦樂前奏 + ritornello + 獨唱 + 合唱\n• Biber, Missa salisburgensis (1682)\n  — 53聲部！16 歌手 + 37 器樂手\n  — Salzburg 大教堂四個合唱閣樓\n\nLutheran North 路德宗北方：\n\n正統 vs 敬虔主義的衝突：\n• Orthodox 正統派：\n  — 堅持公共崇拜 + 合唱/器樂音樂\n• Pietists 敬虔派：\n  — 重私人靈修 → 偏好簡單歌曲\n\nChorale 聖詠傳統延續：\n• Crüger, Praxis pietatis melica (1647)\n  — 最具影響力的聖詩集\n  — 40+ 版次\n\nDieterich Buxtehude (ca. 1637–1707)：\n• Wachet auf 聖詠變奏 — 管弦 + 合唱\n• Abendmusiken 晚間音樂會（免費公開）\n  — Bach 1705 年步行 200 英里去聆聽", {
    x: 5.25, y: 1.45, w: 4.3, h: 3.6, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1,
  });
}

// ── SLIDE 8 · German Organ Music ────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("🎹 德國管風琴音樂 German Organ Music", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.copper } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Toccata, Prelude & Fugue", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Lutheran Organ Music 路德宗管風琴：\n• 1650–1750 黃金時代\n• 功能：聖詠/經文的前奏\n\nBuxtehude 的 Toccata/Praeludium：\n• 自由段落 vs 嚴格賦格段 交替\n  — 承 Frescobaldi + Froberger 傳統\n• 自由段：\n  — 不規則節奏、16分音符驅動\n  — 突然變化方向/力度/調性\n  — 華麗的踏板獨奏 (pedaliter)\n• 賦格段：\n  — 每段使用不同主題\n  — 主題間存在 \"family resemblance\"\n\nPraeludium in E Major BuxWV 141 (NAWM 97)：\n• 五個自由段落 + 四個賦格段\n• 四個賦格主題互為變體\n• 17世紀：toccata/prelude/praeludium\n  ≈ 同義詞（含賦格段落）\n• 18世紀：賦格段獨立 → Prelude + Fugue 分開", {
    x: 0.45, y: 1.45, w: 4.3, h: 3.6, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Chorale Settings & Other Genres", { x: 5.2, y: 1.08, w: 4.4, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Fugue 賦格：\n• 17世紀末取代 ricercar / fantasia / capriccio\n• 主題 (subject) 更個性化、節奏鮮明\n• Exposition 呈示部：\n  — 主題 (tonic) → 答句 answer (dominant)\n  — 各聲部輪流加入\n• Episodes 插段：自由對位\n\nChorale Settings 聖詠設定：\n• Organ chorale 管風琴聖詠\n• Chorale variations / chorale partita\n  — 一系列變奏\n• Chorale prelude 聖詠前奏曲：\n  四種類型：\n  ① 每句為模仿點 (point of imitation)\n  ② 高聲部長音旋律 + 低聲部前仿\n  ③ 高聲部裝飾旋律 + 自由伴奏\n  ④ 旋律 + 無關動機伴奏\n\nOther Instrumental Music：\n• Harpsichord suites (Froberger → A-C-S-G)\n• Orchestral suites (仿 Lully 風格)\n  — Muffat Florilegium (1695, 1698)\n• Violin sonatas：Biber Mystery Sonatas\n  — 15首 + Passacaglia · scordatura\n• Keyboard sonatas：Kuhnau (1696)", {
    x: 5.25, y: 1.45, w: 4.3, h: 3.6, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1,
  });
}

// ── SLIDE 9 · Key Terms & Listening ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("📝 關鍵詞彙 · 延伸閱讀 · 聆聽 Key Terms & Listening", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 20, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.copper } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Key Terms 關鍵詞彙", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• da capo aria · ABA form · ritornello\n• sonata da chiesa · sonata da camera\n• trio sonata · solo sonata\n• walking bass · circle of fifths\n• concerto · concerto grosso\n• concertino · ripieno · tutti\n• orchestral concerto · solo concerto\n• ritornello form\n• cantata · serenata\n• toccata · praeludium · prelude\n• fugue · exposition · subject · answer · episode\n• chorale prelude · chorale variations\n• chorale partita · organ chorale\n• Stadtpfeifer · collegium musicum\n• Orthodox vs Pietist\n• Abendmusiken\n• scordatura · Mystery Sonatas\n• Frische Clavier Früchte\n• Stradivarius · Cremona", {
    x: 0.5, y: 1.42, w: 4.35, h: 3.6, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("📚 Further Reading & 🎧 Listening", { x: 5.25, y: 1.08, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Grout & Williams. A Short History of Opera (2003)\n• Allsop. The Italian Trio Sonata (1992)\n• Talbot. The Vivaldi Compendium (2011)\n• Snyder. Dieterich Buxtehude (2007)\n\n🎧 NAWM 精選聆聽 (YouTube)\n• 93 · Sartorio · Giulio Cesare in Egitto  youtu.be/mvCMFqgN8JU\n• 94 · A. Scarlatti · Clori vezzosa, e bella  youtu.be/8EVdPg5TeKE\n• 95 · A. Scarlatti · La Griselda  youtu.be/h0oSd1uNQII\n• 96 · Corelli · Trio Sonata Op. 3 No. 2  youtu.be/ozEfJugPMe4\n• 97 · Buxtehude · Praeludium in E BuxWV 141  youtu.be/yg7_GPKCTWI", {
    x: 5.3, y: 1.42, w: 4.35, h: 3.6, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

pres.writeFile({ fileName: "Ch17_Italy_Germany.pptx" })
  .then(fn => console.log(`✅ ${fn} created successfully`))
  .catch(err => console.error("❌ Error:", err));
