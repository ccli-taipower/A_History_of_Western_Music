const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 10: Madrigal and Secular Song in the Sixteenth Century";
pres.author = "A History of Western Music, 10th ed.";

const C = {
  darkBg:   "2A1A2E",
  gold:     "C8A020",
  cream:    "FBF5E6",
  wine:     "7A1A3A",
  rose:     "A03050",
  darkText: "2A1A2E",
  lightText:"FBF5E6",
  blush:    "E8A0B0",
  sand:     "E8D8A8",
  plum:     "5A2040",
  slate:    "3A2840",
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
    x: 0.5, y: 0.45, w: 9, h: 0.35, fontSize: 14, color: C.sand, charSpacing: 3, align: "center", fontFace: "Georgia",
  });
  s.addText("CHAPTER 10", {
    x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });
  s.addText("MADRIGAL AND SECULAR SONG\nIN THE SIXTEENTH CENTURY", {
    x: 0.3, y: 1.5, w: 9.4, h: 2.0, fontSize: 30, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.65, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Villancico · Frottola · Madrigal · Chanson · Lied · Lute Song", {
    x: 0.4, y: 3.8, w: 9.2, h: 0.4, fontSize: 14, color: C.sand, align: "center", fontFace: "Georgia",
  });
  s.addText("Textbook pp. 205–228", {
    x: 0.5, y: 4.8, w: 9, h: 0.3, fontSize: 14, color: C.gold, align: "center", fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("本章概覽 Chapter Overview", { x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const sections = [
    ["■", "First Market for Music 音樂的市場", "Petrucci 印刷術催生業餘用樂市場 · 民族語言興起"],
    ["ES", "Spain 西班牙", "Villancico · Juan del Encina · Oy comamos y bebamos"],
    ["IT", "Italy — Frottola & Madrigal", "Frottola · Verdelot · Arcadelt · Willaert · Rore · Gesualdo"],
    ["FR", "France — New Chansons", "Sermisy · Janequin · Lassus · Musique mesurée"],
    ["DE", "Germany — Lied & Meistersinger", "Hans Sachs · Senfl · Lassus 德語歌曲"],
    ["GB", "England — Madrigal & Lute Song", "Morley · Weelkes · Dowland · Triumphes of Oriana"],
  ];
  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.wine }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 14, color: C.rose, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 3 · First Market for Music ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("音樂的第一個市場", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("The First Market for Music · Printing, Amateurs & National Styles", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "3A2840" }, rounding: true });
  s.addText("■ 音樂成為商品", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 1501 Petrucci《Odhecaton》首部複音印刷集\n• 印刷降低成本——樂譜首次大量流通\n• 音樂從「服務」變成「商品」\n• 業餘者市場迅速成長\n\n■ 社交音樂\n• Castiglione《廷臣論》(1528) 要求廷臣讀譜演唱\n• 城市中產與貴族視讀譜為社交教養\n• partbooks 於社交聚會中傳閱", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "3A2840" }, rounding: true });
  s.addText("■ 民族風格興起", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("業餘者偏好母語演唱——促成 16 世紀各國民族樂種\n\n• Spain → villancico\n• Italy → frottola / madrigal\n• France → chanson\n• Germany → Lied / Meistersinger\n• England → consort song / madrigal / ayre\n\n■ 共同特徵\n• 以文字意義與情感驅動音樂\n• Text depiction & expression 成為主流美學\n• 承繼 Josquin——humanism 深化", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 4 · Spain & Italy Frottola ─────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rose); bottomBar(s, C.rose);

  s.addText("西班牙 Villancico · 義大利 Frottola", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("Courtly Songs with Popular Flavor", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 14, color: C.plum, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  // Villancico
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("ES Villancico", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 源自 villano（農夫）——「鄉村風歌曲」\n• 短、分節、音節式、同節奏\n• 結構 AAB：estribillo + coplas (mudanza → vuelta)\n• 上聲部主導；餘聲部可唱或奏\n• 為貴族寫的「模擬民謠」\n\n■ Juan del Encina (1468–1529)\n• 西班牙首位劇作家 · villancico 大師\n• 田園戲劇中穿插 villancico\n• Oy comamos y bebamos (NAWM 46)\n  封齋前夕狂歡歌 · 頻繁 hemiola", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Frottola
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("IT Frottola", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 四聲部、分節、音節式、同節奏\n• 簡單自然音階——幾乎全 root-position 三和弦\n• 旋律在上聲部，下聲部提供和聲\n• 由即興詩歌演唱傳統衍生\n\n■ 宮廷風情\n• 盛行於 Mantua、Ferrara、Urbino 宮廷\n• Cara · Tromboncino（Mantua）\n• 贊助者 Isabella d'Este\n\n■ Petrucci 1504–1514 出版 13 冊\n• 印刷時代首個「暢銷音樂商品」\n• 常以 solo voice + lute 改編流通", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 5 · The Italian Madrigal — Definition ──────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("義大利牧歌 The Italian Madrigal", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Renaissance 末期最重要的世俗樂種 · pave the way for opera", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Definition
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "3A2840" }, rounding: true });
  s.addText("■ 定義與特徵", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 1530 後指義大利詩歌的音樂配樂\n• 單節詩——7/11 音節、無副歌、自由形式\n• 與 14 世紀 madrigal 同名異物\n• 通篇作曲（through-composed）\n• 早期 4 聲部 → 中期 5 聲部 → 晚期 6+\n• 每聲部一人；常加樂器 double\n\n■ 詩材\n• Petrarca · Ariosto · Tasso · Guarini · Marino\n• 主題：愛情、田園、感傷", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Aesthetics
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "3A2840" }, rounding: true });
  s.addText("■ 美學：文字驅動音樂", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("以音樂再現詩的聲音、意象與情感\n\n• Text expression（情感表達）· 核心\n• Word painting（文字繪畫）\n• Declamation（依重音配長短）\n\n■ 社交場合\n• 學院晚宴後演唱 · 男女混編業餘圈\n• 1570 後出現職業歌手\n\n■ 產量\n• 1530–1600 出版 2,000+ 部\n• 為 opera 鋪路——義大利成歐洲音樂領袖", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 6 · Early Madrigalists — Verdelot & Arcadelt ───────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("早期牧歌作曲家", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("Early Madrigalists · Verdelot · Arcadelt · Il bianco e dolce cigno", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 14, color: C.plum, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  // Verdelot
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("■ Philippe Verdelot (ca. 1480/85–?1530)", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 法國出身——活躍於羅馬與佛羅倫斯\n• 1520s Florence madrigal 成形期核心人物\n• 4 聲部：同節奏、cadence 分句（承襲 frottola）\n• 5/6 聲部：近 motet，頻繁模仿\n\n■ Jacques Arcadelt (ca. 1507–1568)\n• Franco-Flemish · 1530s 起 Florence/Rome 近 30 年\n• 1551 返法任 Sainte-Chapelle 神職\n• 風格平衡、優雅、易唱\n• 1539 First Book of Madrigals", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Il bianco e dolce cigno
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("■ Il bianco e dolce cigno (NAWM 47)", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("Arcadelt 1538 · 最著名的早期 madrigal\n\n「潔白甜美的天鵝哀唱著死亡」\n• 天鵝之死 vs.「死於至樂」的愛情\n• 隱喻「小死」（la petite mort）\n\n■ 音樂處理\n• 大部分同聲織度——清晰、平衡\n• 天鵝之死：上下半音「哀歎音型」\n• 結尾「mille mort' il dì」多聲模仿\n• 簡潔巧妙——成為教科書典範\n\n■ 16 世紀前半最流傳的 madrigal", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 7 · Willaert, Rore & the Petrarchan Movement ─────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Willaert · Rore · Petrarchan 運動", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Venetian Madrigal · Bembo 美學 · 半音的誕生", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Willaert
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "3A2840" }, rounding: true });
  s.addText("■ Adrian Willaert & Bembo 美學", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Adrian Willaert (ca. 1490–1562)\n• 法蘭德斯 · 35 年任 St Mark's maestro\n• 培養 Rore、Vicentino、Gabrieli、Zarlino\n\n■ Bembo 美學\n• Bembo 編訂 Petrarch《Canzoniere》(1501)\n• 兩極美學：\n  — piacevolezza（悅耳）：柔滑\n  — gravità（莊嚴）：斷促\n• Willaert《Aspro core e selvaggio》\n  嚴肅／溫柔句以不同音程對比\n• Zarlino (1558) 將其理論化", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Rore & Vicentino
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "3A2840" }, rounding: true });
  s.addText("■ Cipriano de Rore & Nicola Vicentino", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Cipriano de Rore (1516–1565)\n• Willaert 學生 · Ferrara/Parma/威尼斯\n• Da le belle contrade d'oriente (NAWM 48)\n  — 1566 死後遺集\n  — 每字情感都被音樂抓住\n  — 「Te'n vai, haimè」半音、小三度刻畫悲嘆\n  — 一小段用盡 12 個半音\n• 此前不允許 B■→B■ 直接半音\n• Rore 讓直接半音成 grief 象徵\n\nNicola Vicentino (1511–ca. 1576)\n• 《L'antica musica》(1555)\n• 復興古希臘 chromatic / enharmonic\n• chromaticism 成 16 世紀後期共同語言", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 8 · Late Madrigalists & Gesualdo ───────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("晚期牧歌與 Gesualdo", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("Marenzio · Gesualdo · Women Composers & Performers", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 14, color: C.plum, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  // Marenzio & Casulana
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("■ Marenzio & Casulana", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("Luca Marenzio (1553–1599)\n• 16 世紀末最重要義大利 madrigalist\n• Solo e pensoso (NAWM 49) · Petrarch\n• 半音上行逾八度——「孤獨沉思」\n• 「逃」「躲」用快速緊密模仿\n\n■ Maddalena Casulana (ca. 1544–90s)\n• 首位出版樂曲的女性作曲家\n• First Book of Madrigals (1568)\n\n■ Concerto delle donne (Ferrara, 1580)\n• Alfonso II d'Este 召募\n• 將 madrigal 從業餘轉為職業", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Gesualdo
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("■ Carlo Gesualdo, Prince of Venosa", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("1566–1613 · 音樂史最驚心動魄者之一\n\n• 貴族兼職業作曲家（罕見）\n• 殺死通姦妻子與情人\n\n■ 極端對比美學\n• 偏好強烈意象的現代詩\n• diatonic ↔ chromatic 劇烈切換\n• 不協和／協和並置 · 常切斷詩句\n\n■ \"Io parto\" (NAWM 50) · 1611\n• 慢速、半音、不協和描寫悲嘆\n• 「復活」(vivo son) 轉快速 diatonic\n• 將文字戲劇化推向極致", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 9 · Lighter Italian Genres & Legacy ────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("輕型義大利樂種與牧歌的遺產", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Villanella · Canzonetta · Balletto · Legacy of the Madrigal", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Light genres
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "3A2840" }, rounding: true });
  s.addText("■ Light Genres", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Villanella\n• 1540s 於 Naples 興起 · 三聲部、活潑、同節奏\n• 寫平行五度模仿鄉村樸拙，常戲擬 madrigal\n\nCanzonetta（小歌）\n• Orazio Vecchi · 1580–1597 六冊\n• 融合 madrigal 與 villanella\n\nBalletto（小舞曲）\n• Gastoldi · 1591/1594 · AABB 分節曲式\n• 「fa-la-la」副歌\n• Morley · Sing we and chant it (NAWM 55)", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Legacy
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "3A2840" }, rounding: true });
  s.addText("■ The Legacy of the Madrigal", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• madrigal 使音樂成為文字戲劇載體\n• 業餘社交 → 私人音樂會 → 舞台\n\n■ 音樂理想\n• 旋律與節奏依循自然語言\n• 詩意透過具象音型 (madrigalism)\n\n■ 承先啟後\n• 人文主義在音樂的最高表現\n• 孕育 1600 年誕生的 opera\n• 義大利成歐洲音樂領袖逾 200 年", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 10 · French Chanson ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rose); bottomBar(s, C.rose);

  s.addText("新式法國香頌", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("The New Parisian Chanson · Sermisy · Janequin · Lassus", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 14, color: C.plum, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  const chansons = [
    { title: "Lyric Chanson 抒情香頌", ex: "Sermisy · Tant que vivray (NAWM 51)", desc: "第一人稱抒情——愛情主題 · 音節、同節奏、頂部主旋律 · Marot 詩 · 短句樂段明確" },
    { title: "Narrative Chanson 敘事香頌", ex: "Janequin · Martin menoit son pourceau (NAWM 52)", desc: "低俗幽默故事 · 點狀模仿開句、同聲終止 · 模仿喜歌劇般的口白節奏" },
    { title: "Descriptive Chanson 描繪香頌", ex: "Janequin · La guerre · Le chant des oiseaux", desc: "模擬鳥鳴、戰號、街市叫賣 · 長度較長 · 織度變化多端" },
    { title: "Musique Mesurée 量音香頌", ex: "Le Jeune · Revecy venir (NAWM 54)", desc: "Académie de Poésie et de Musique (1570) · Baïf 的 vers mesurés · 長音配長音節" },
  ];

  chansons.forEach((c, i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 0.3 + col * 4.8;
    const y = 1.35 + row * 1.95;
    s.addShape(pres.ShapeType.rect, { x, y, w: 4.6, h: 1.85, fill: { color: "F3DDE5" }, rounding: true });
    s.addText(c.title, { x: x + 0.15, y: y + 0.08, w: 4.35, h: 0.32, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia" });
    s.addText(c.ex, { x: x + 0.15, y: y + 0.42, w: 4.35, h: 0.28, fontSize: 14, color: C.plum, fontFace: "Georgia" });
    s.addText(c.desc, { x: x + 0.15, y: y + 0.72, w: 4.35, h: 1.1, fontSize: 14, color: C.darkText, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 11 · Lassus & Sermisy / Janequin Context ──────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Attaingnant · Sermisy · Janequin · Lassus", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Publishers & Masters of the French Chanson", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Left — Attaingnant & composers
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "3A2840" }, rounding: true });
  s.addText("■ Pierre Attaingnant & Composers", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Pierre Attaingnant (ca. 1494–ca. 1552)\n• 法國首位音樂印刷商 · 1528 起\n• 1528–1552 出版 50+ 本香頌集\n• 約 1,500 首作品——使香頌普及家庭\n\nClaudin de Sermisy (ca. 1490–1562)\n• Paris Sainte-Chapelle\n• 12 彌撒、100 經文歌，以 175 首香頌最著名\n\nClément Janequin (ca. 1485–1558)\n• Bordeaux · Angers · 晚年 Paris\n• 250 首香頌 · 「描繪香頌之王」", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Right — Lassus
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "3A2840" }, rounding: true });
  s.addText("■ Orlande de Lassus (ca. 1532–1594)", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("歐洲最國際化的作曲家\n• 生於 Mons（比利時 Hainaut）\n• 青年服侍 Mantua、Milan、Naples、Rome\n• 1556 進 Bavaria Albrecht V 宮廷\n• 1563–94 任 Munich ducal chapel maestro\n• 首位獲皇帝授予版權的作曲家\n\n■ 巨量產量\n• 57 彌撒 · 700+ 經文歌 · 101 Magnificat\n• 150 chansons · 200 madrigals · 90 Lieder\n• 兒子編《Magnum opus musicum》(1604)\n\n• La nuict froide et sombre (NAWM 53)\n  — du Bellay 詩（Pléiade）\n  — 融合 madrigal 與 chanson", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 12 · Germany · Lied & Meistersinger ───────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("德國：Lied 與 Meistersinger", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("German Secular Song · Meistersinger · Polyphonic Lied", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 14, color: C.plum, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  // Meistersinger
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("■ Meistersinger 工匠名歌手", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 德語城市工匠的音樂行會\n• 承 Minnesinger 單聲歌曲傳統\n• 嚴格規則創作 · 公開比賽\n• 14 世紀始 → 16 世紀顛峰 → 19 世紀消亡\n\n■ Töne（旋律模式）\n• 每首詩依既有 Ton 模板填詞\n• 皆為 bar form（AAB）\n• 許多源自 Minnelied 舊譜\n\n■ Hans Sachs (1494–1576)\n• 紐倫堡製鞋匠\n• 作詞數千首 · 新創 13 首 Toöne\n• Wagner《紐倫堡的名歌手》主角", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Polyphonic Lied
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("■ Polyphonic Lied 德語複音歌曲", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 延續 tenor 主旋律 + 對位聲部的傳統\n• 主旋律可在 tenor 或 cantus\n• 紐倫堡為 16 世紀上半葉文化中心\n\n■ Ludwig Senfl (ca. 1486–1542/3)\n• Isaac 的學生 · 宮廷禮拜堂作曲家\n• Polyphonic Lied 的代表人物\n\n■ 下半葉義大利化\n• 1550 後偏好義大利 madrigal 與 villanella\n• Lied 吸收 madrigal——模仿與同節奏交替\n\n• Lassus 寫 7 冊德文 Lied——\n  實為使用德語的 madrigal", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 13 · England · Consort Song · Madrigal · Lute Song ────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("英格蘭：Consort Song · Madrigal · Lute Song", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Morley · Weelkes · Wilbye · Dowland · Campion", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Consort song & English madrigal
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "3A2840" }, rounding: true });
  s.addText("■ Consort Song & English Madrigal", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Consort Song\n• 英格蘭獨有——獨唱 + viol consort\n• Byrd《Psalmes, Sonets and Songs》(1588)\n\nItalian Influence\n• 1588 Yonge《Musica transalpina》\n  英譯義大利 madrigal——引爆風潮\n\nThomas Morley (1557/8–1602)\n• 最早、最多產的英國 madrigalist\n• 仿 Gastoldi 寫 Sing we (NAWM 55)\n\n■ 《Triumphes of Oriana》(1601)\n• 23 作曲家 · 獻 Elizabeth I", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Weelkes & Dowland
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "3A2840" }, rounding: true });
  s.addText("■ Weelkes · Dowland · Campion", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Thomas Weelkes (ca. 1575–1623)\n• As Vesta was (NAWM 56) · 極盡 word-painting\n  — ascending/descending 音階對應\n  — 結尾「Long live fair Oriana」近 50 次\n\nLute Song (Ayre) · 1600s 初\n• 獨唱 + lute——更個人、更文學\n\nJohn Dowland (1563–1626)\n• Flow, my tears (NAWM 57) · Pavane · 陰鬱\n\nCampion (1567–1620) · 詩人兼音樂家", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 14 · Timeline ─────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("時間軸 · Timeline", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 0.82, w: 5, h: 0.04, fill: { color: C.wine } });

  const events = [
    ["ca. 1495", "Juan del Encina · Oy comamos y bebamos"],
    ["1504–14", "Petrucci 出版 13 冊 frottola 集"],
    ["1515–47", "法王 Francis I 在位"],
    ["1528", "Castiglione《廷臣論》 · Attaingnant 首部香頌集"],
    ["1539", "Arcadelt First Book of Madrigals"],
    ["1555", "Vicentino《L'antica musica》"],
    ["1556–94", "Lassus 任 Bavaria 宮廷"],
    ["1558", "Zarlino《Le istitutioni harmoniche》"],
    ["1558–1603", "英國 Elizabeth I 在位"],
    ["1566", "Rore Fifth Book of Madrigals"],
    ["1568", "Maddalena Casulana First Book of Madrigals"],
    ["1580", "Concerto delle donne 於 Ferrara 成立"],
    ["1588", "Yonge《Musica transalpina》"],
    ["1595", "Morley《First Book of Balletts》"],
    ["1599", "Marenzio 最後 madrigal 集"],
    ["1600", "Dowland Second Book of Songs or Ayres"],
    ["1601", "《The Triumphes of Oriana》"],
    ["1611", "Gesualdo 最後 madrigal 集"],
  ];
  events.forEach(([date, desc], i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 0.3 + col * 4.8;
    const y = 1.0 + row * 0.47;
    s.addShape(pres.ShapeType.rect, { x, y, w: 1.1, h: 0.38, fill: { color: C.wine } });
    s.addText(date, { x: x + 0.05, y: y + 0.04, w: 1.0, h: 0.3, fontSize: 14, bold: true, color: C.lightText, align: "center", fontFace: "Georgia" });
    s.addText(desc, { x: x + 1.2, y, w: 3.55, h: 0.38, fontSize: 14, color: C.darkText, fontFace: "Calibri", valign: "middle" });
  });
}

// ── SLIDE 15 · Key Terms & Further Reading ──────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("關鍵詞彙 · 延伸閱讀", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Key Terms & Further Reading", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "3A2840" }, rounding: true });
  s.addText("■ Key Terms", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• villancico · estribillo · coplas · mudanza · vuelta\n• frottola · Isabella d'Este · Petrucci\n• madrigal (16c) · through-composed\n• text declamation / depiction / expression\n• madrigalism · word painting\n• piacevolezza · gravità · Bembo\n• chromatic genus · direct chromatic motion\n• villanella · canzonetta · balletto\n• chanson (Parisian) · lyric / narrative / descriptive\n• musique mesurée · vers mesurés à l'antique\n• Meistersinger · Ton · bar form · Tenorlied\n• consort song · English madrigal · ballett\n• lute song / ayre · pavane\n• concerto delle donne · partbooks", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "3A2840" }, rounding: true });
  s.addText("■ Further Reading & ■ Listening", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Fenlon & Haar. The Italian Madrigal (Early 16c)\n• Einstein. The Italian Madrigal (1949)\n• Kerman. The Elizabethan Madrigal (1962)\n• Freedman. Music in the Renaissance (Norton)\n\n■ NAWM 精選聆聽\n47 Arcadelt · Il bianco e dolce cigno\nyoutu.be/wMImLeCewio\n56 Weelkes · As Vesta Was Descending\nyoutu.be/95DJ7oqTWK8\n57 Dowland · Flow my tears\nyoutu.be/y3REIVlo2Ss", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

pres.writeFile({ fileName: "Ch10_Madrigal.pptx" })
  .then(fn => console.log(`■ ${fn} created successfully`))
  .catch(err => console.error("■ Error:", err));
