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
    x: 0.5, y: 0.45, w: 9, h: 0.35, fontSize: 11, color: C.sand, charSpacing: 3, align: "center", fontFace: "Georgia",
  });
  s.addText("CHAPTER 10", {
    x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });
  s.addText("MADRIGAL AND SECULAR SONG\nIN THE SIXTEENTH CENTURY", {
    x: 0.3, y: 1.5, w: 9.4, h: 2.0, fontSize: 30, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.65, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Villancico · Frottola · Madrigal · Chanson · Lied · Lute Song", {
    x: 0.4, y: 3.8, w: 9.2, h: 0.4, fontSize: 13, color: C.sand, align: "center", fontFace: "Georgia",
  });
  s.addText("Textbook pp. 205–228", {
    x: 0.5, y: 4.8, w: 9, h: 0.3, fontSize: 11, color: C.gold, align: "center", fontFace: "Calibri",
  });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("本章概覽 Chapter Overview", { x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const sections = [
    ["📘", "First Market for Music 音樂的市場", "Petrucci 印刷術催生業餘用樂市場 · 民族語言興起"],
    ["🇪🇸", "Spain 西班牙", "Villancico · Juan del Encina · Oy comamos y bebamos"],
    ["🇮🇹", "Italy — Frottola & Madrigal", "Frottola · Verdelot · Arcadelt · Willaert · Rore · Gesualdo"],
    ["🇫🇷", "France — New Chansons", "Sermisy · Janequin · Lassus · Musique mesurée"],
    ["🇩🇪", "Germany — Lied & Meistersinger", "Hans Sachs · Senfl · Lassus 德語歌曲"],
    ["🇬🇧", "England — Madrigal & Lute Song", "Morley · Weelkes · Dowland · Triumphes of Oriana"],
  ];
  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.wine }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 11, color: C.rose, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 3 · First Market for Music ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("音樂的第一個市場", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("The First Market for Music · Printing, Amateurs & National Styles", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2840" }, rounding: true });
  s.addText("💰 音樂成為商品", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 1501 Petrucci 首部複音印刷集《Odhecaton》\n• 印刷降低成本——第一次使樂譜得以大量流通\n• 音樂從「服務」變成可販售的「商品」\n• 業餘者市場迅速成長\n• 出版商依需求出版——從精英到通俗的各種風格\n\n🎨 社交音樂的興起\n• Castiglione《廷臣論》(1528) 要求廷臣能讀譜演唱\n• 城市中產與貴族皆將讀譜演奏視為社交教養\n• 合唱 partbooks 於社交聚會中傳閱", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 9, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 3,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2840" }, rounding: true });
  s.addText("🌍 民族風格興起", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("業餘者偏好自己的母語演唱——\n促使各國在 16 世紀發展出獨特的民族樂種\n\n• Spain（西班牙）→ villancico\n• Italy（義大利）→ frottola / madrigal\n• France（法國）→ chanson（新型三式）\n• Germany（德國）→ Lied / Meistersinger 歌曲\n• England（英格蘭）→ consort song / madrigal / ayre\n\n✨ 共同特徵\n• 以文字意義與情感驅動音樂\n• Text depiction & expression 成為主流美學\n• 承繼 Josquin 一代的文字主義——humanism 深化", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 9, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 4 · Spain & Italy Frottola ─────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rose); bottomBar(s, C.rose);

  s.addText("西班牙 Villancico · 義大利 Frottola", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("Courtly Songs with Popular Flavor", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 13, color: C.plum, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  // Villancico
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("🇪🇸 Villancico", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 源自 villano（農夫）——「鄉村風歌曲」\n• 短、分節、音節式、主要同節奏\n• 結構：estribillo（副歌）+ coplas（詩節）\n  — mudanza（變化段）→ vuelta（回到副歌）\n  — 整體 AAB 模式\n• 最上聲部主導旋律；其餘聲部可唱或奏\n• 為西班牙貴族寫作的「模擬民謠」\n\n🌟 Juan del Encina (1468–1529)\n• 第一位西班牙劇作家 · villancico 大師\n• 田園戲劇中穿插 villancico\n• Oy comamos y bebamos (NAWM 46)\n  封齋前夕的狂歡歌 · 嬉鬧 · 頻繁 hemiola", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  // Frottola
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("🇮🇹 Frottola", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 四聲部、分節、音節式、同節奏\n• 簡單自然音階和聲——幾乎全用 root-position 三和弦\n• 旋律在最上聲部——下聲部提供和聲支撐\n• 由即興詩歌演唱傳統衍生而來\n\n🎭 宮廷風情\n• 1500 年代前後盛行於 Mantua、Ferrara、Urbino 宮廷\n• 幾乎全由義大利作曲家創作\n• Marchetto Cara · Bartolomeo Tromboncino（Mantua）\n• 贊助者 Isabella d'Este——文藝復興最著名的女性贊助人\n\n📚 Petrucci 於 1504–1514 出版 13 冊 frottola 集\n• 成為印刷時代第一個「暢銷音樂商品」\n• 常以 solo voice + lute 改編版本流通\n（Bossinensis 1509 出版）", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 5 · The Italian Madrigal — Definition ──────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("義大利牧歌 The Italian Madrigal", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Renaissance 末期最重要的世俗樂種 · pave the way for opera", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Definition
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2840" }, rounding: true });
  s.addText("📖 定義與特徵", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 1530 年後用「madrigal」指義大利詩歌的音樂配樂\n• 單節詩——7 或 11 音節詩行、無副歌、不受固定形式拘束\n• 與 14 世紀 madrigal「同名異物」\n• 通篇作曲（through-composed）——每行詩配新音樂\n• 早期四聲部（1520–1540）\n• 中期起五聲部為主（1540–）\n• 晚期六聲部或更多\n• 每聲部一人——為室內聲樂\n• 常加入樂器 double 或替代聲部\n\n📝 詩材來源\n• Petrarca · Ariosto · Tasso · Guarini · Marino\n• 主題以愛情、田園、感傷為主", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  // Aesthetics
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2840" }, rounding: true });
  s.addText("✨ 美學：文字驅動音樂", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("作曲家以音樂再現詩的聲音、意象與情感\n\n• Text expression（情感表達）· 最核心\n• Text depiction / word painting（文字繪畫）\n• Declamation（言語宣告）· 依重音配長短\n\n🌹 社交場合\n• 學院（academies）晚宴後演唱\n• 男女混編的業餘圈子\n• 1570 後開始出現職業歌手為觀眾演唱\n\n📊 產量驚人\n• 1530–1600 間出版 2,000 餘部 madrigal 集\n• 一直盛行到 17 世紀初\n\n🎭 歷史意義\n• 為 opera 的戲劇表達鋪路\n• 義大利因 madrigal 成為歐洲音樂領袖", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 6 · Early Madrigalists — Verdelot & Arcadelt ───────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("早期牧歌作曲家", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("Early Madrigalists · Verdelot · Arcadelt · Il bianco e dolce cigno", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 13, color: C.plum, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  // Verdelot
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("🌸 Philippe Verdelot (ca. 1480/85–?1530)", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 11, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 法國出身——活躍於羅馬與佛羅倫斯\n• madrigal 1520 年代在 Florence 成形時的核心人物\n• 四聲部 madrigal：多為同節奏、樂句以 cadence 標示——承襲 frottola 傳統\n• 五／六聲部 madrigal：更接近 motet——頻繁模仿、聲部交疊\n\n🌟 Jacques Arcadelt (ca. 1507–1568)\n• Franco-Flemish——1530 年代起在 Florence 與 Rome 工作近 30 年\n• 1551 返回法國——任 Sainte-Chapelle 神職\n• 風格：同聲織度與偶發模仿交織——平衡、優雅、易唱\n• 1539 First Book of Madrigals", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  // Il bianco e dolce cigno
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("🦢 Il bianco e dolce cigno (NAWM 47)", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 11, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("Arcadelt 1538 · 最著名的早期 madrigal\n\n「潔白甜美的天鵝哀唱著死亡」\n• 詩意：天鵝之死對比說話者「死於至樂」的愛情\n• 可能暗指新柏拉圖主義——愛情如死而復生\n• 亦可能隱喻「小死」（la petite mort · 情慾高潮）\n\n🎵 音樂處理\n• 大部分同聲織度——清晰、平衡\n• 天鵝的死以上下半音刻畫——哀歎音型\n• 結尾「mille mort' il dì」（千死一日）\n  以多聲模仿進入——形成鮮明對比\n• 簡單卻巧妙——成為 madrigal 教科書典範\n\n✨ 歷史地位\n• 16 世紀前半最流傳的 madrigal", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 7 · Willaert, Rore & the Petrarchan Movement ─────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Willaert · Rore · Petrarchan 運動", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Venetian Madrigal · Bembo 美學 · 半音的誕生", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Willaert
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2840" }, rounding: true });
  s.addText("📜 Adrian Willaert & Bembo 美學", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 11, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Adrian Willaert (ca. 1490–1562)\n• 法蘭德斯出身 · 35 年任威尼斯 St Mark's maestro di cappella\n• 培養 Cipriano de Rore、Vicentino、A. Gabrieli、Zarlino 等\n\n💎 Bembo 美學\n• Cardinal Pietro Bembo 編訂 Petrarch《Canzoniere》(1501)\n• 提出兩極美學：\n  — piacevolezza（悅耳）：大調六度／小三度／柔滑\n  — gravità（莊嚴）：小調六度／大三度／斷促\n• Willaert 設 Petrarch《Aspro core e selvaggio》\n  — 嚴肅句：大三度旋律、尖銳和聲\n  — 溫柔句：半音、小三度與小六度\n\n• Zarlino《Le istitutioni harmoniche》(1558)\n  將 Willaert 的做法理論化為教科書原則", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  // Rore & Vicentino
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2840" }, rounding: true });
  s.addText("🎭 Cipriano de Rore & Nicola Vicentino", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 11, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Cipriano de Rore (1516–1565)\n• Willaert 學生 · Ferrara、Parma、威尼斯工作\n• Da le belle contrade d'oriente (NAWM 48)\n  — 1566 出版於死後遺集\n  — 每個字的情感都被音樂抓住\n  — 「Te'n vai, haimè, adio」以上下半音、小三度刻畫悲嘆\n  — 一小段內用盡 12 個半音——直接 chromatic 進行\n• 此前對位法不允許 B♭→B♮ 的直接半音\n• Rore 讓直接半音成為 grief 的強力象徵\n\nNicola Vicentino (1511–ca. 1576)\n• 《L'antica musica ridotta alla moderna prattica》(1555)\n• 主張復興古希臘的 chromatic 與 enharmonic 音系\n• L'aura che 'l verde lauro——用希臘 chromatic 四音列做模仿動機\n• 最終 chromaticism 成為 16 世紀後期的共同語言", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 8 · Late Madrigalists & Gesualdo ───────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("晚期牧歌與 Gesualdo", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("Marenzio · Gesualdo · Women Composers & Performers", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 13, color: C.plum, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  // Marenzio & Casulana
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("🌟 Marenzio & Casulana", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 11, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("Luca Marenzio (1553–1599)\n• 16 世紀末最重要的義大利 madrigalist\n• Solo e pensoso (NAWM 49)——Petrarch 十四行詩\n• 開頭：以每小節半音上行超過八度\n  描繪「孤獨沉思、緩慢步伐」\n• 「逃」「躲」則用快速緊密模仿\n\n👩 Maddalena Casulana (ca. 1544–ca. 1590s)\n• 首位出版樂曲、首位自稱「職業作曲家」的女性\n• First Book of Madrigals (1568)\n• 獻辭反擊男性壟斷：\n  「揭露以為女性無法具有高智者之謬誤」\n\n🎤 Concerto delle donne (Ferrara, 1580)\n• Alfonso II d'Este 召募訓練\n• Peverara · Guarini · d'Arco 三位女歌手\n• 引發 Mantua / Florence 仿效\n• 將 madrigal 從業餘社交轉為職業音樂", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  // Gesualdo
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("🗡️ Carlo Gesualdo, Prince of Venosa", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 11, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("1566–1613 · 音樂史最驚心動魄的人物之一\n\n• 貴族兼職業作曲家——極為罕見\n• 發現妻子與情人在床——殺死兩人\n• 1593 再娶 Leonora d'Este（Alfonso II 姪女）\n\n🎨 極端的對比美學\n• 偏好充滿強烈意象的現代詩\n• 在 diatonic 與 chromatic 之間劇烈切換\n• 不協和與協和、chordal 與 imitative 並置\n• 慢速與快速節奏對比\n• 常切斷詩句以突顯單字\n\n🎵 \"Io parto\" e non più dissi (NAWM 50)\n• 1611 最後一本 madrigal 集\n• 「我將走，不再多言」\n  以慢速、半音、不協和描寫女方的悲嘆\n• 情郎「復活」(vivo son) 則轉為快速 diatonic 模仿\n• 將 madrigal 的文字戲劇化推向極致", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 9 · Lighter Italian Genres & Legacy ────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("輕型義大利樂種與牧歌的遺產", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Villanella · Canzonetta · Balletto · Legacy of the Madrigal", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Light genres
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2840" }, rounding: true });
  s.addText("💃 Light Genres", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Villanella\n• 1540s 於 Naples 興起\n• 三聲部、活潑、同節奏\n• 故意寫平行五度 → 模仿鄉村樸拙\n• 常戲擬嚴肅的 madrigal\n\nCanzonetta（小歌）\n• Orazio Vecchi 首先使用 · 1580–1597 六冊\n• 融合 madrigal 與 villanella 元素\n\nBalletto（小舞曲）\n• Gastoldi · 1591 五聲部集、1594 三聲部集\n• 舞曲節奏 · AABB 分節曲式\n• 著名「fa-la-la」副歌\n• 影響英德作曲家——Morley《Sing we and chant it》仿 Gastoldi《A lieta vita》(NAWM 55)", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  // Legacy
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2840" }, rounding: true });
  s.addText("🏛️ The Legacy of the Madrigal", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• madrigal 使音樂成為文字戲劇的載體\n• 業餘社交 → 私人音樂會 → 舞台演出\n• 從「演唱給自己聽」走向「演唱給觀眾聽」\n\n✨ 共同的音樂理想\n• 旋律與節奏依循自然語言\n• 詩意透過具象音型呈現 (madrigalism)\n• 音樂元素整體對應詩的情感\n\n🎭 承先啟後\n• 人文主義在音樂中的最高表現\n• 直接孕育 1600 年誕生的 opera\n• 所有後世歌劇、藝術歌曲、電影配樂\n  的情感表達語言皆源於此\n\n• Wills the Italians「第一次」成為歐洲音樂領袖——一領導就超過 200 年", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 10 · French Chanson ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rose); bottomBar(s, C.rose);

  s.addText("新式法國香頌", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("The New Parisian Chanson · Sermisy · Janequin · Lassus", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 13, color: C.plum, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  const chansons = [
    { title: "Lyric Chanson 抒情香頌", ex: "Sermisy · Tant que vivray (NAWM 51)", desc: "第一人稱抒情——愛情主題 · 音節、同節奏、頂部主旋律 · Marot 詩 · 短句樂段明確" },
    { title: "Narrative Chanson 敘事香頌", ex: "Janequin · Martin menoit son pourceau (NAWM 52)", desc: "低俗幽默故事 · 點狀模仿開句、同聲終止 · 模仿喜歌劇般的口白節奏" },
    { title: "Descriptive Chanson 描繪香頌", ex: "Janequin · La guerre · Le chant des oiseaux", desc: "模擬鳥鳴、戰號、街市叫賣 · 長度較長 · 織度變化多端" },
    { title: "Musique Mesurée 量音香頌", ex: "Le Jeune · Revecy venir du printans (NAWM 54)", desc: "Académie de Poésie et de Musique (1570) · Baïf 的 vers mesurés · 長音節配長音 · 二三拍自由交替" },
  ];

  chansons.forEach((c, i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 0.3 + col * 4.8;
    const y = 1.35 + row * 1.95;
    s.addShape(pres.ShapeType.rect, { x, y, w: 4.6, h: 1.85, fill: { color: "F3DDE5" }, rounding: true });
    s.addText(c.title, { x: x + 0.15, y: y + 0.08, w: 4.35, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
    s.addText(c.ex, { x: x + 0.15, y: y + 0.42, w: 4.35, h: 0.28, fontSize: 9.5, color: C.plum, fontFace: "Georgia" });
    s.addText(c.desc, { x: x + 0.15, y: y + 0.72, w: 4.35, h: 1.1, fontSize: 8.5, color: C.darkText, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 11 · Lassus & Sermisy / Janequin Context ──────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Attaingnant · Sermisy · Janequin · Lassus", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Publishers & Masters of the French Chanson", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Left — Attaingnant & composers
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2840" }, rounding: true });
  s.addText("📚 Pierre Attaingnant & Composers", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Pierre Attaingnant (ca. 1494–ca. 1552)\n• 法國第一位音樂印刷商 · 1528 開始\n• 1528–1552 出版 50+ 本香頌集\n• 約 1,500 首作品——使香頌成為家家戶戶的娛樂\n\nClaudin de Sermisy (ca. 1490–1562)\n• 法國王室禮拜堂 · Paris Sainte-Chapelle\n• 作 12 部彌撒、100 經文歌，但以 175 首香頌最著名\n\nClément Janequin (ca. 1485–1558)\n• Bordeaux · Angers · 晚年 Paris\n• 250 首香頌 · 「compositeur ordinaire du roi」\n• 描繪香頌之王", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  // Right — Lassus
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2840" }, rounding: true });
  s.addText("🌍 Orlande de Lassus (ca. 1532–1594)", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 11, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("歐洲最國際化的作曲家\n• 生於 Mons（比利時 Hainaut 區）\n• 青年時期服侍 Mantua、Sicily、Milan、Naples、Rome\n• 24 歲已出版 madrigal、chanson、motet 集\n• 1556 進 Bavaria Albrecht V 宮廷\n• 1563 起任 Munich ducal chapel maestro · 至 1594 辭世\n• 獲法王、神聖羅馬皇帝授予音樂版權——首位掌握版權的作曲家\n\n📊 巨大的產量\n• 57 部彌撒 · 700+ 經文歌 · 101 首 Magnificat\n• 150 法文 chansons · 200 義大利 madrigals · 90 德文 Lieder\n• 死後由兒子編《Magnum opus musicum》(1604) 出版經文歌集\n\n• La nuict froide et sombre (NAWM 53)\n  — du Bellay 詩（Pléiade 詩派）\n  — 融合 madrigal 與 chanson 傳統", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 12 · Germany · Lied & Meistersinger ───────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("德國：Lied 與 Meistersinger", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("German Secular Song · Meistersinger · Polyphonic Lied", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 13, color: C.plum, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  // Meistersinger
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("🔨 Meistersinger 工匠名歌手", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 都市商人與工匠組織的歌唱行會\n• 以中世紀 Minnesinger 單聲歌曲為基礎\n• 依嚴格規則創作 · 公開比賽演唱\n• 14 世紀始於 → 16 世紀達顛峰 → 19 世紀末葉消亡\n\n📜 Toöne（旋律與格律模式）\n• 每首詩依既有 Ton 模板填詞\n• 皆為 bar form（AAB）\n• 許多 Toöne 源自 Minnelied 舊譜\n\n🌟 Hans Sachs (1494–1576)\n• 紐倫堡的製鞋匠\n• 作詞數千首 · 新創 13 首 Toöne\n• 最知名的 Meistersinger\n• 後為 Wagner 歌劇《紐倫堡的名歌手》主角", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  // Polyphonic Lied
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "F3DDE5" }, rounding: true });
  s.addText("🎼 Polyphonic Lied 德語複音歌曲", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 德國作曲家延續 tenor 主旋律、其餘聲部對位的傳統\n• 主旋律可在 tenor 或 cantus\n• 紐倫堡為 16 世紀上半葉德國文化中心——大量出版 Lied 集\n\n🌟 Ludwig Senfl (ca. 1486–1542/3)\n• Isaac 的學生\n• 宮廷禮拜堂作曲家\n• Polyphonic Lied 的代表人物\n\n📈 下半葉轉向義大利化\n• 1550 後德國聽眾偏好義大利 madrigal 與 villanella\n• Lied 吸收 madrigal 特徵——模仿與同節奏交替、密切文字配合\n\n• Lassus 寫 7 冊德文 Lied——\n  實質上是使用德語的 madrigal\n  聲部平等、文字敏感度高", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 13 · England · Consort Song · Madrigal · Lute Song ────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("英格蘭：Consort Song · Madrigal · Lute Song", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Morley · Weelkes · Wilbye · Dowland · Campion", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Consort song & English madrigal
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2840" }, rounding: true });
  s.addText("🎻 Consort Song & English Madrigal", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 11, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Consort Song\n• 英格蘭獨有——獨唱 + 提琴重奏 (viol consort)\n• William Byrd《Psalmes, Sonets and Songs》(1588)\n  將 consort song 提升為藝術體裁\n\nItalian Influence\n• 1560s 義大利 madrigal 開始在英倫流傳\n• 1588 Nicholas Yonge《Musica transalpina》\n  把義大利 madrigal 譯成英文出版——引爆英國 madrigal 風潮\n\nThomas Morley (1557/8–1602)\n• 最早、最多產的英國 madrigalist\n• 仿 Gastoldi《A lieta vita》寫《Sing we and chant it》(NAWM 55)\n• 《A Plaine and Easie Introduction to Practicall Musicke》(1597)\n  面向業餘者的音樂教本\n\n👑 《The Triumphes of Oriana》(1601)\n• 23 位作曲家、25 首 madrigal 合集\n• 每首以「Long live fair Oriana」收尾——獻給 Elizabeth I", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  // Weelkes & Dowland
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2840" }, rounding: true });
  s.addText("🎵 Weelkes · Dowland · Campion", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 11, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Thomas Weelkes (ca. 1575–1623)\n• As Vesta was (NAWM 56)——收於《Triumphes of Oriana》\n• 自撰詩 + 音樂——極盡 word-painting\n  — 上行音階「ascending」/ 下行「descending」\n  — 一、二、三聲部表示「alone」「two by two」「three by three」\n  — 結尾「Long live fair Oriana」動機進入近 50 次\n  + augmentation 4、8 倍——象徵長壽的機智\n\nLute Song (Ayre) · 1600s 初\n• 獨唱 + lute 伴奏——更個人、更文學\n• 整體情感刻畫、word-painting 減少\n\nJohn Dowland (1563–1626)\n• Flow, my tears (NAWM 57)——1600 第二本 Songs or Ayres\n• Pavane 形式 · aabbCC · 優美陰鬱\n• 歐洲知名——哥本哈根、德累斯頓皆有任職\n\nThomas Campion (1567–1620)\n• 詩人兼音樂家——兼及文學理論\n\n📖 1620 年代後英國風潮消退——\n為 17 世紀 solo song 鋪路", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
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
    s.addText(date, { x: x + 0.05, y: y + 0.04, w: 1.0, h: 0.3, fontSize: 9, bold: true, color: C.lightText, align: "center", fontFace: "Georgia" });
    s.addText(desc, { x: x + 1.2, y, w: 3.55, h: 0.38, fontSize: 8, color: C.darkText, fontFace: "Calibri", valign: "middle" });
  });
}

// ── SLIDE 15 · Key Terms & Further Reading ──────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("關鍵詞彙 · 延伸閱讀", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Key Terms & Further Reading", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2840" }, rounding: true });
  s.addText("🔑 Key Terms", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• villancico · estribillo · coplas · mudanza · vuelta\n• frottola · Isabella d'Este · Petrucci\n• madrigal (16c) · through-composed\n• text declamation / depiction / expression\n• madrigalism · word painting\n• piacevolezza · gravità · Bembo\n• chromatic genus · direct chromatic motion\n• villanella · canzonetta · balletto\n• chanson (Parisian) · lyric / narrative / descriptive\n• musique mesurée · vers mesurés à l'antique\n• Meistersinger · Ton · bar form · Tenorlied\n• consort song · English madrigal · ballett\n• lute song / ayre · pavane\n• concerto delle donne · partbooks", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "3A2840" }, rounding: true });
  s.addText("📚 Further Reading & 🎧 Listening", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Fenlon & Haar. The Italian Madrigal in the Early 16c\n• Einstein. The Italian Madrigal (1949)\n• Kerman. The Elizabethan Madrigal (1962)\n• Freedman. Music in the Renaissance (Norton)\n\n🎧 NAWM 精選聆聽 (YouTube)\n• 46 · Encina · Oy comamos  youtu.be/oeW8z3-RNVI\n• 47 · Arcadelt · Il bianco e dolce cigno  youtu.be/ium1MZN-58g\n• 48 · Rore · Da le belle contrade  youtu.be/PgoT6Klkf0o\n• 49 · Marenzio · Solo e pensoso  youtu.be/ZJlj1uy8cSA\n• 50 · Gesualdo · Io parto  youtu.be/TBC-45-FfVQ\n• 51 · Sermisy · Tant que vivray  youtu.be/yD7qRFELl8w\n• 52 · Janequin · Martin menoit  youtu.be/_5VDzWU7vlc\n• 53 · Lassus · La nuict froide  youtu.be/89gNkOjZ8Dg\n• 54 · Le Jeune · Revecy venir  youtu.be/ieUrg8d3z70\n• 55a · Gastoldi · A lieta vita  youtu.be/xRQJ8DW8Cnc\n• 55b · Morley · Sing we & chant it  youtu.be/AcA6QdMAvO8\n• 56 · Weelkes · As Vesta was  youtu.be/9LLDwTNj6f4\n• 57 · Dowland · Flow my tears  youtu.be/Y9HKl8H0PWg", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1,
  });
}

pres.writeFile({ fileName: "Ch10_Madrigal.pptx" })
  .then(fn => console.log(`✅ ${fn} created successfully`))
  .catch(err => console.error("❌ Error:", err));
