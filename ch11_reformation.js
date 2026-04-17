const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 11: Sacred Music in the Era of the Reformation";
pres.author = "A History of Western Music, 10th ed.";

// Reformation palette — sober, ecclesiastical; Lutheran indigo / sacramental gold / altar cream
const C = {
  darkBg:    "1A2240",  // deep indigo
  gold:      "C89440",  // altar gold
  cream:     "F7F2E2",
  wine:      "5A1F2E",  // Catholic purple-red
  rose:      "8A3C50",
  darkText:  "1A2240",
  lightText: "F7F2E2",
  sand:      "E8D8A8",
  slate:     "2A334A",
  lilac:     "B8A8C8",
  parchment: "EFE3C4",
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
  s.addText("CHAPTER 11", {
    x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });
  s.addText("SACRED MUSIC IN THE ERA\nOF THE REFORMATION", {
    x: 0.3, y: 1.5, w: 9.4, h: 2.0, fontSize: 32, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.65, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Luther · Calvin · Tallis · Byrd · Palestrina · Victoria · Lassus", {
    x: 0.4, y: 3.8, w: 9.2, h: 0.4, fontSize: 14, color: C.sand, align: "center", fontFace: "Georgia",
  });
  s.addText("Textbook pp. 229–253", {
    x: 0.5, y: 4.8, w: 9, h: 0.3, fontSize: 14, color: C.gold, align: "center", fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("本章概覽 Chapter Overview", { x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const sections = [
    ["■", "The Reformation 宗教改革", "Luther · Calvin · Henry VIII 撕裂西歐教會統一"],
    ["■", "Lutheran Music", "Chorale · Kontrafaktur · Johann Walter · 四聲部合唱本"],
    ["■", "Calvinist & English", "Geneva Psalter · Anglican Service · Anthem"],
    ["■", "Catholic Reform", "Council of Trent · Palestrina · Pope Marcellus Mass"],
    ["■", "Spain & New World", "Victoria · Morales · Hernando Franco (Mexico)"],
    ["■", "Germany & Jewish Music", "Lassus · Handl · Ashkenazic / Sephardic 聖歌傳統"],
  ];
  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.wine }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 14, color: C.rose, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 3 · The Reformation — Historical Context ──────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("宗教改革的歷史背景", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("The Reformation · A Church Divided (1517– )", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.slate }, rounding: true });
  s.addText("■ 一場撕裂歐洲的運動", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 1517 路德《九十五條論綱》釘於 Wittenberg\n• 反對贖罪券、教廷腐敗、教宗權威\n• 原本只求改革——卻導致分裂\n\n三條主要路線\n• Lutheran · 保留大量禮儀與音樂\n• Calvinist · 簡化禮拜、詩篇歌唱\n• Anglican · 亨利八世脫離羅馬 (1534)\n\n結果\n• 西歐基督宗教統一瓦解\n• 各地以民族語言敬拜\n• 音樂依教派神學發展不同風格", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.slate }, rounding: true });
  s.addText("■ 天主教回應 · Counter-Reformation", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 天主教會啟動自我更新\n• Council of Trent (1545–1563)\n  — 界定教義 · 改革神職 · 規範禮儀\n  — 音樂決議：文字清晰、驅逐世俗旋律\n\n音樂影響\n• Palestrina 式「清晰對位」成典範\n• 拉丁文仍為禮儀標準\n• 耶穌會將音樂帶到新世界\n\n■ 整體圖像\n• 1500–1600 各教派用音樂定義信仰", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 4 · Luther & Music ────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rose); bottomBar(s, C.rose);

  s.addText("路德與音樂 Luther & Music", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("\"Next to the Word of God, music deserves the highest praise.\"", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 14, color: C.rose, italic: true, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.parchment }, rounding: true });
  s.addText("■ 路德的音樂觀", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("• 受過完整音樂訓練——能唱、彈魯特琴\n• 喜愛 Josquin——稱「音符的主人」\n• 音樂是神所賜 · 次於 Word 的贈禮\n\n■ 保留與改革並行\n• 保留拉丁禮儀於學校與大城市\n• 創新德語禮儀 (Deutsche Messe, 1526)\n• 鼓勵會眾參與唱詩\n\n■ 音樂教育\n• 強調學校必須教音樂\n• 奠定三百年日耳曼音樂傳統\n• 影響深及 Schütz · J. S. Bach", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.parchment }, rounding: true });
  s.addText("■ Chorale 會眾聖詠", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("• 德語單聲部會眾歌曲——齊唱\n• 單節式 · 音節式 · 易記憶\n• 文字皆德語——會眾真正理解\n\n■ 來源\n1. 新創——路德作詞、部份作曲\n2. 拉丁聖歌改編\n3. 前宗教改革德語歌曲修訂\n4. 世俗歌曲改填聖詞 (Kontrafaktur)\n\n■ 代表作\n• Ein feste Burg · 路德 ca. 1529 · 改革「國歌」\n• Aus tiefer Not\n• Vom Himmel hoch\n\n■ Achtliederbuch (1524) · 首部官方聖詠本", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 5 · Chorale & Polyphonic Settings ─────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("聖詠的多聲部設置", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Chorale · Johann Walter · Ein feste Burg (NAWM 58)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.slate }, rounding: true });
  s.addText("■ Johann Walter (1496–1570)", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 路德最親近的音樂家\n• Geystliche gesangk Buchleyn (1524)\n  — 首部路德宗多聲部聖詠集\n  — 38 德語 + 5 拉丁聖歌\n  — 四或五聲部 · 旋律在 tenor\n\n■ 功能\n• 非為會眾——為學校合唱團練習\n• 家庭與學校的教學場景\n\n■ 發展\n• 1550 後 chorale 旋律移至最上聲部\n• 1586 Osiander 首部 cantus-on-top 和聲集\n• 為日後 Bach 聖詠和聲法奠基", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.slate }, rounding: true });
  s.addText("■ Ein feste Burg ist unser Gott", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 詞曲：Martin Luther, ca. 1529\n• 取材 Psalm 46——「神是我們的避難所」\n• 四節詩 · bar form (AAB)\n• 被稱「宗教改革的戰歌」\n\n■ Walter 四聲部設置 (NAWM 58)\n• Tenor 持聖詠旋律\n• 其他聲部以對位裝飾\n• 典型路德宗 tenor-cantus firmus 織度\n\n■ 歷史迴響\n• 德國「新教國歌」\n• Bach BWV 80 同名清唱劇\n• Mendelssohn、Wagner、Mahler 皆引用", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 6 · Calvin · Geneva Psalter ───────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("加爾文與日內瓦詩篇集", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("John Calvin · Psalms Only · Geneva Psalter (1562)", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 14, color: C.rose, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.parchment }, rounding: true });
  s.addText("■ 加爾文的嚴格路線", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("• 法裔改革神學家 · 活躍於日內瓦\n• 音樂強大但危險——必須謹慎\n• 敬拜只能唱聖經原文 (Psalmody)\n• 反對多聲、管風琴\n\n原則\n1. 單聲部齊唱\n2. 法語押韻詩篇\n3. 素樸旋律——會眾易學\n4. 禮拜僅用人聲\n\n■ 多聲部僅限家中或學校\n• Goudimel 四聲部日內瓦詩篇 (1564)", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.parchment }, rounding: true });
  s.addText("■ Geneva Psalter (1562)", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("• 150 篇押韻法文詩篇集\n• 詞：Marot · de Bèze\n• 曲：Loys Bourgeois (ca. 1510–60)\n\n特徵\n• 每篇獨立旋律——共 125 首曲調\n• 自然音階 · 旋律優雅上口\n\n■ Psalm 134 (NAWM 59)\n• 後改填為 Doxology\n• 英語「Old 100th」旋律源自此\n\n■ 傳播\n• 翻譯 20+ 種語言 · 法語圈、荷、蘇格蘭\n• 加爾文派禮拜基石 · 至今仍用", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 7 · Church of England ─────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("英格蘭教會音樂", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Church of England · Tallis · Byrd · Anglican Service & Anthem", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.slate }, rounding: true });
  s.addText("■ 政治脈絡 1534–1603", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 1534 亨利八世脫離羅馬\n• 1549 Book of Common Prayer\n• Mary I (1553–58) 短暫復辟天主教\n• Elizabeth I (1558–1603) 折衷路線\n  — 保留主教制與禮儀之美\n  — 改用英語但詞曲精緻\n\n■ Anglican 音樂特徵\n• 英語為主 · 一音一字\n• 保留大教堂合唱\n\n■ Service\n• Morning / Evening Prayer · Communion", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.slate }, rounding: true });
  s.addText("■ Anthem 英文聖歌", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 英語相當於拉丁 motet\n• Full anthem · 全合唱\n• Verse anthem · 獨唱/合唱交替 + 管風琴\n\n■ Thomas Tallis (ca. 1505–1585)\n• 跨四朝 Chapel Royal 音樂家\n• 早期拉丁 motet · 後適應英文\n• If ye love me (NAWM 60) · 完美範例\n• Spem in alium · 40 聲部傑作\n\n■ William Byrd (ca. 1540–1623)\n• Tallis 學生 · Chapel Royal\n• 私下天主教徒——在新教國度寫拉丁禮儀\n• Sing joyfully unto God (NAWM 61) · 6 聲部", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 8 · Catholic Reform & Trent ───────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("天主教改革 · 特倫特公會議", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("Council of Trent (1545–1563) · Reform of Liturgical Music", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 14, color: C.rose, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.parchment }, rounding: true });
  s.addText("■ Trent 對音樂的討論", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("• 1545–1563 間三次開會\n• 討論：彌撒音樂是否過度世俗化？\n\n改革派訴求\n• 文字必須清晰可辨\n• 驅逐世俗 cantus firmus 彌撒\n• 激進派主張驅逐所有複音\n\n最終決議 (1562)\n• 允許複音 · 文字必須聽懂\n• 禁止「猥褻」旋律\n• 未指定特定風格\n\n■ 地方執行各異\n• Borromeo 在米蘭積極推行", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.parchment }, rounding: true });
  s.addText("■ Palestrina 傳說", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("流傳的故事\n• Palestrina《Missa Papae Marcelli》\n  拯救複音免被禁\n• 以清晰文字說服樞機團\n\n事實\n• 此彌撒 1562–3 年創作\n• 清晰風格已是改革共識\n\n■ 象徵意義\n• 19 世紀視為「天主教音樂的救主」\n• Pfitzner 1917 歌劇《Palestrina》\n\n■ 後續\n• Palestrina 對位成標準教材\n• Fux《Gradus ad Parnassum》(1725)", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 9 · Palestrina · Life & Style ─────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Giovanni Pierluigi da Palestrina", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("ca. 1525–1594 · 文藝復興天主教音樂的典範", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.slate }, rounding: true });
  s.addText("■ 生平與職位", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 出生於羅馬近郊 Palestrina 鎮\n• 羅馬各大聖堂任 maestro di cappella\n\n■ 主要任職\n• Cappella Giulia · 1551–\n• Sistine Chapel · 1554（已婚被迫離開）\n• St. John Lateran · 1555–1560\n• Santa Maria Maggiore · 1561–1566\n• 重返 Cappella Giulia · 1571 至離世\n\n■ 作品：104 Masses · 375+ motets", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.slate }, rounding: true });
  s.addText("■ The Palestrina Style", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("標準化文藝復興對位——後世學理典範\n\n旋律\n• 音階級進為主 · 跳進後必級進回\n• 跳進上限五度或八度 · 節奏平滑\n\n和聲與不協和\n• 不協和皆為 passing/neighbor/suspension\n• Suspension 嚴格準備與解決\n\n織度與美學\n• 4–6 聲部 · 模仿進入\n• 不誇張、內斂崇高 · a cappella 代名詞", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 10 · Pope Marcellus Mass (NAWM 62) ───────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("教宗馬切魯斯彌撒 (NAWM 62)", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("Missa Papae Marcelli · Credo", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 14, color: C.rose, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.parchment }, rounding: true });
  s.addText("■ 背景", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("• 1562–3 創作 · 1567 收入第二部 Mass 集\n• 獻給 Marcellus II（在位 3 週, 1555）\n• 6 聲部 SATTBB\n• paraphrase 技法（非固定 cantus firmus）\n• 無世俗旋律——符合 Trent 精神\n\n■ Credo 特徵\n• 長文字需清晰\n• 大部分同節奏 (homorhythm)\n• 聲部幾乎同時換字\n• et incarnatus / crucifixus 靜思感\n• 模仿段落較少——優先清晰度\n\n■ 織度\n• 每段獨立分節\n• SA · TB · 全體 · 六聲部輪流", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.parchment }, rounding: true });
  s.addText("■ 為何成為經典？", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("• 完美回應 Trent 會議關切\n• 平衡「清晰文字」與「多聲藝術」\n• 宣告詞意同時提供神聖美感\n\n■ 歷史影響\n• 1725 Fux《Gradus ad Parnassum》典範\n• 19 世紀 Cecilian Movement 視為正統\n• 20 世紀對位法教材仍取自 Palestrina\n\n■ 文化符碼\n• 「純粹教會音樂」代名詞\n• Pfitzner 1917《Palestrina》歌劇\n• Stravinsky · Pärt 曾研究其風格", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 11 · Spain & New World ────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("西班牙與新世界", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Victoria · Morales · New World Polyphony · Hernando Franco", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.slate }, rounding: true });
  s.addText("■ Victoria & Morales", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("■ Tomás Luis de Victoria (1548–1611)\n• Avila 出生——至羅馬 German College\n• 可能師事 Palestrina\n• 返西班牙任王室禮拜堂樂長\n• 僅寫宗教音樂——近 180 首\n• 比 Palestrina 更熾烈、神秘\n\n■ O magnum mysterium (NAWM 63) · 1572\n• 四聲部聖誕 motet\n• fifths/octaves 表達「偉大奧祕」\n• 最著名的文藝復興 motet 之一\n\n■ Morales (ca. 1500–1553)\n• 西班牙首位國際知名作曲家", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.slate }, rounding: true });
  s.addText("■ New World Polyphony", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 西班牙征服使複音首度抵達美洲\n• 墨西哥 (1521) · 秘魯 Cuzco\n• 原住民子弟入修道院學拉丁與五線譜\n\n■ 教學與作曲\n• 方濟會、道明會傳教士教導\n• 聖詠本、譜冊普及各聖堂\n• 原住民與西班牙樂長並肩工作\n\n■ Hernando Franco (1532–1585)\n• 西班牙 Galicia 出生 · 1554 抵墨西哥\n• 1575–85 任 Mexico City 樂長\n\n■ Salve regina (NAWM 64)\n• 織度近似 Palestrina——未「美洲化」", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 12 · Lassus & Germany · Jewish Music ──────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("Lassus · 德國 · 猶太音樂", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("Orlande de Lassus · Jacob Handl · Jewish Liturgical Music", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 14, color: C.rose, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.parchment }, rounding: true });
  s.addText("■ Orlande de Lassus (1532–1594)", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("• Franco-Flemish · Munich 宮廷 1556–94\n• 拉、義、法、德四語創作\n• 2,000+ 作品跨各體裁\n• 天主教徒但作品新舊兩教共享\n\n■ Cum essem parvulus (NAWM 65)\n• 六聲部拉丁 motet\n• 《林前 13:11》「我小時候...」\n• 高聲部代表童聲 · 低聲部代表成人\n• Lassus 文字敏感度的巔峰\n\n■ 影響\n• Jesuit 教育體系推廣全歐\n• 承 Palestrina 清晰原則 + 表達力\n• 16 世紀後半最重要國際音樂家", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.parchment }, rounding: true });
  s.addText("■ 中東歐與猶太音樂", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("■ Jacob Handl / Gallus (1550–1591)\n• 斯洛維尼亞 · 活躍於 Prague、Olomouc\n• Opus musicum (1586–91) · 全年 motet 集\n• 融合 Palestrina 清晰與威尼斯雙合唱\n\n■ 波蘭、波希米亞、匈牙利\n• 新舊教並存 · 與義大利往來密切\n• Gomółka（波蘭）亦寫詩篇集\n\n■ Jewish Music\n• 會堂禮拜口傳吟唱——無記譜\n• Salamone Rossi (1570–1630)\n  — 曼圖阿宮廷 · 首位記譜猶太複音", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 13 · Byrd Mass for 4 Voices + Legacy ──────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Byrd 的四聲部彌撒 · 宗教音樂的遺產", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Byrd · Mass for Four Voices (NAWM 66) · The Legacy of 16c Sacred Music", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.slate }, rounding: true });
  s.addText("■ Byrd Mass for Four Voices", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 1592–93 出版（無版權頁）\n• 為英國天主教家庭秘密彌撒而作\n• 3、4、5 聲部三冊中最簡潔一本\n\n■ 風格\n• 模仿進入 · 典型文藝復興織度\n• 避免冗長——適合私人場合\n• Credo 精簡緊湊\n• 模仿 Taverner 舊英國傳統\n\n■ Elizabeth I 時代天主教徒\n• 私下彌撒違法——可罰款入獄\n• Byrd 受 Elizabeth 寵愛得以繼續\n• 僅在鄉村莊園小禮拜堂使用", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.1, fill: { color: C.slate }, rounding: true });
  s.addText("■ The Legacy of 16c Sacred Music", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 16 世紀：西歐四種神學 · 四種音樂\n\n■ 長遠影響\n• Lutheran chorale → Bach 聖詠\n• Calvinist psalmody → 英語聖詩\n• Anglican anthem → Purcell · Handel\n• Palestrina 對位 → 神學院必修\n\n■ 音樂作為宗教認同\n• 每教派以獨特聲音界定自身\n• 音樂風格帶神學意涵\n\n■ 全球化起點\n• 歐洲多聲部抵達美洲\n• 傳教士帶拉丁禮儀至中、日、菲", {
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
    ["1517", "Luther《九十五條論綱》· Wittenberg"],
    ["1524", "Achtliederbuch · Walter Geystliche gesangk Buchleyn"],
    ["ca. 1529", "Luther · Ein feste Burg ist unser Gott"],
    ["1534", "Act of Supremacy · Henry VIII 脫離羅馬"],
    ["1539", "Calvin 首版 Strasbourg 詩篇集"],
    ["1545–63", "Council of Trent"],
    ["1549", "Book of Common Prayer (Cranmer)"],
    ["1553–58", "Mary I 短暫復辟天主教"],
    ["1554–85", "Palestrina 活躍於羅馬"],
    ["1558–1603", "Elizabeth I 在位"],
    ["1562", "Geneva Psalter 完成"],
    ["1567", "Palestrina · Missa Papae Marcelli 出版"],
    ["1572", "Victoria · O magnum mysterium"],
    ["1575", "Tallis & Byrd · 英國音樂出版專利"],
    ["1575", "Franco 任 Mexico City 樂長"],
    ["1586–91", "Handl · Opus musicum"],
    ["1592–93", "Byrd · Mass for Four Voices"],
    ["1594", "Palestrina · Lassus 同年辭世"],
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

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.25, fill: { color: C.slate }, rounding: true });
  s.addText("■ Key Terms", { x: 0.45, y: 1.36, w: 4.3, h: 0.3, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", valign: "top", margin: 0 });
  s.addText("• Reformation · Counter-Reformation\n• Lutheran · Calvinist · Anglican · Catholic\n• chorale · Kontrafaktur · bar form\n• Achtliederbuch · Geystliche gesangk Buchleyn\n• metrical psalm · Geneva Psalter · Old 100th\n• Book of Common Prayer · Anglican Service\n• full anthem · verse anthem\n• Council of Trent · Canon on music\n• Missa Papae Marcelli · paraphrase mass\n• Palestrina style · stile antico\n• O magnum mysterium · Hernando Franco\n• New World polyphony · Chapel Royal\n• Salamone Rossi · Cecilian movement", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.75, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.25, fill: { color: C.slate }, rounding: true });
  s.addText("■ Further Reading & ■ Listening", { x: 5.25, y: 1.36, w: 4.3, h: 0.3, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", valign: "top", margin: 0 });
  s.addText("• Leaver. Luther's Liturgical Music\n• Rodin. Josquin's Rome\n• Monson. Disembodied Voices\n• Snyder. The Chorale\n• Fenlon (ed.). The Renaissance\n\n■ NAWM 精選聆聽\n60 Tallis · If ye love me  youtu.be/yHe2FDlHHa8\n61 Byrd · Sing joyfully  youtu.be/9uK9nVVbGHw\n62 Palestrina · Pope Marcellus  youtu.be/oeLIgzAe5sI\n63 Victoria · O magnum mysterium  youtu.be/RqkHy_Os5h8", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.75, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0,
  });
}

pres.writeFile({ fileName: "Ch11_Reformation.pptx" })
  .then(fn => console.log(`■ ${fn} created successfully`))
  .catch(err => console.error("■ Error:", err));
