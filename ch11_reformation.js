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
    x: 0.5, y: 0.45, w: 9, h: 0.35, fontSize: 11, color: C.sand, charSpacing: 3, align: "center", fontFace: "Georgia",
  });
  s.addText("CHAPTER 11", {
    x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });
  s.addText("SACRED MUSIC IN THE ERA\nOF THE REFORMATION", {
    x: 0.3, y: 1.5, w: 9.4, h: 2.0, fontSize: 32, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.65, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Luther · Calvin · Tallis · Byrd · Palestrina · Victoria · Lassus", {
    x: 0.4, y: 3.8, w: 9.2, h: 0.4, fontSize: 13, color: C.sand, align: "center", fontFace: "Georgia",
  });
  s.addText("Textbook pp. 229–253", {
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
    ["⛪", "The Reformation 宗教改革", "Luther · Calvin · Henry VIII 撕裂西歐教會統一"],
    ["🎵", "Lutheran Music", "Chorale · Kontrafaktur · Johann Walter · 四聲部合唱本"],
    ["📖", "Calvinist & English", "Geneva Psalter · Anglican Service · Anthem"],
    ["🇮🇹", "Catholic Reform", "Council of Trent · Palestrina · Pope Marcellus Mass"],
    ["🌍", "Spain & New World", "Victoria · Morales · Hernando Franco (Mexico)"],
    ["🇩🇪", "Germany & Jewish Music", "Lassus · Handl · Ashkenazic / Sephardic 聖歌傳統"],
  ];
  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.wine }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 11, color: C.rose, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 3 · The Reformation — Historical Context ──────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("宗教改革的歷史背景", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("The Reformation · A Church Divided (1517– )", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("📜 一場撕裂歐洲的運動", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 1517 馬丁·路德《九十五條論綱》釘於 Wittenberg 城堡教堂\n• 反對贖罪券、教廷腐敗、教宗權威\n• 原本只求改革——卻導致分裂\n\n三條主要路線\n• Lutheran 路德宗 · 保留大量禮儀與音樂\n• Calvinist 加爾文宗 · 簡化禮拜、詩篇歌唱\n• Anglican 英國國教 · 亨利八世脫離羅馬 (1534)\n\n結果\n• 西歐基督宗教統一瓦解\n• 各地以民族語言敬拜\n• 音樂因應各教派神學發展不同風格", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 9, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 3,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("✝ 天主教回應 · Counter-Reformation", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 天主教會並非坐視——啟動自我更新\n• Council of Trent (1545–1563)\n  — 界定教義 · 改革神職 · 規範禮儀\n  — 對音樂的決議：文字必須清晰、驅逐世俗旋律\n\n音樂的影響\n• Palestrina 式「清晰的對位」成為典範\n• 拉丁文仍為禮儀標準——不讓步給民族語言\n• 耶穌會傳教將 Counter-Reformation 音樂帶到新世界\n\n📈 整體圖像\n• 1500–1600 的一百年——\n  每一種教派都在用「音樂」定義自己的信仰\n• 教會音樂從此多元並存\n• 為 17 世紀巴洛克風格的擴展提供沃土", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 4 · Luther & Music ────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rose); bottomBar(s, C.rose);

  s.addText("路德與音樂 Luther & Music", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("\"Next to the Word of God, music deserves the highest praise.\"", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 12, color: C.rose, italic: true, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.parchment }, rounding: true });
  s.addText("🎵 路德的音樂觀", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 受過完整音樂訓練——能唱、能彈魯特琴\n• 喜愛 Josquin——稱其為「音符的主人」\n• 認為音樂是神所賜 · 次於 Word 的最佳贈禮\n• 音樂應服務於教育與敬拜\n\n📖 保留與改革並行\n• 保留拉丁禮儀音樂於學校及大城市\n• 同時創造新的德語禮儀 (Deutsche Messe, 1526)\n• 鼓勵會眾參與唱詩——這是路德宗的決定性創新\n\n🎓 音樂教育\n• 強調學校必須教音樂\n• 宗教音樂家由學校培養——奠定三百年日耳曼音樂傳統\n• 影響深及 Schütz · J. S. Bach", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.parchment }, rounding: true });
  s.addText("🎼 Chorale 會眾聖詠", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 德語單聲部會眾歌曲——每人一節, 齊唱\n• 單節式 · 音節式 · 易於記憶\n• 文字皆為德語——讓會眾真正理解內容\n\n📚 來源\n1. 新創——路德本人作詞、部份作曲\n2. 拉丁聖歌改編——Contrafactum（世俗→聖）\n3. 前宗教改革德語歌曲修訂\n4. 由世俗歌曲改填聖詞（Kontrafaktur）\n\n🌟 代表作\n• Ein feste Burg ist unser Gott (上主是我堅固保障)\n  路德 ca. 1529 作詞作曲 · 宗教改革的「國歌」\n• Aus tiefer Not (從深處向你呼求)\n• Vom Himmel hoch (從天上而降)\n\n📖 首部官方聖詠本 Achtliederbuch (1524)\n• 8 首歌曲 · 路德與 Paul Speratus 共著", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 5 · Chorale & Polyphonic Settings ─────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("聖詠的多聲部設置", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Chorale · Johann Walter · Ein feste Burg (NAWM 58)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎼 Johann Walter (1496–1570)", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 路德最親近的音樂家——協助路德編配禮儀\n• Geystliche gesangk Buchleyn (1524)\n  — 首部路德宗多聲部聖詠集\n  — 38 首德語 + 5 首拉丁聖歌\n  — 四或五聲部 · 旋律在 tenor\n\n💡 功能\n• 非為會眾唱——為學校合唱團練習\n• 供聖詠學習與記憶\n• 走向家庭與學校的教學場景\n\n📊 發展\n• 1550 後 chorale 旋律逐漸移到最上聲部\n• 1586 Lucas Osiander《Fünfftzig geistliche Lieder》\n  — 首部 cantus-on-top 和聲化聖詠集\n  — 為日後 Bach 聖詠和聲法奠基", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🏰 Ein feste Burg ist unser Gott", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 詞曲：Martin Luther, ca. 1529\n• 取材 Psalm 46——「神是我們的避難所」\n• 四節詩 · bar form (AAB)\n• 旋律剛勁如軍樂——被稱「宗教改革的戰歌」\n\n🎵 Walter 的四聲部設置 (NAWM 58)\n• Tenor 持聖詠旋律\n• 其他聲部以對位裝飾\n• 音節式 · 節奏穩定\n• 典型路德宗「tenor-cantus firmus」織度\n\n📜 歷史迴響\n• 德國「新教國歌」\n• 17 世紀後成為 Bach 清唱劇素材\n  (BWV 80《Ein feste Burg》同名清唱劇)\n• 19 世紀德國浪漫主義民族精神象徵\n• Mendelssohn、Wagner、Mahler 皆引用", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 6 · Calvin · Geneva Psalter ───────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("加爾文與日內瓦詩篇集", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("John Calvin · Psalms Only · Geneva Psalter (1562)", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 13, color: C.rose, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.parchment }, rounding: true });
  s.addText("⛪ 加爾文的嚴格路線", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 法裔改革神學家 · 活躍於日內瓦\n• 相信音樂強大但危險——必須謹慎使用\n• 敬拜中只能唱聖經原文 (Psalmody)\n• 反對多聲、管風琴、非 biblical 詩歌\n\n原則\n1. 單聲部齊唱\n2. 母語 (法語) 押韻詩篇\n3. 素樸旋律——會眾易學\n4. 禮拜中僅用人聲\n\n📜 多聲部設置\n• 可以有——但只能在家中或學校歌唱\n• 不能進入禮拜\n• Claude Goudimel (1505–1572)\n  — 四聲部和聲化日內瓦詩篇 (1564)\n  — 旋律多在 cantus——方便家庭歌唱", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.parchment }, rounding: true });
  s.addText("📖 Geneva Psalter (1562)", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 150 篇全部完成的押韻法文詩篇集\n• 詞：Clément Marot · Théodore de Bèze\n• 曲：Loys Bourgeois (ca. 1510–ca. 1560)\n\n特徵\n• 每篇配獨立旋律——共 125 首曲調\n• 多為自然音階 · 旋律優雅上口\n• 各詩篇格律不同——旋律須適應文字\n\n🌟 Psalm 134: Or sus, serviteurs du Seigneur (NAWM 59)\n• 後來被改填為 Doxology\n  (Praise God from whom all blessings flow)\n• 英語「Old 100th」旋律源自此\n\n📊 傳播\n• 1562 首版——法語圈、荷蘭、瑞士、蘇格蘭、英國\n• 翻譯成 20 多種語言\n• 成為加爾文派與改革宗教會的禮拜基石\n• 至今仍在北美 Presbyterian / Reformed 教會使用", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 7 · Church of England ─────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("英格蘭教會音樂", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Church of England · Tallis · Byrd · Anglican Service & Anthem", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("👑 政治脈絡 1534–1603", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 1534 亨利八世《至尊法案》——脫離羅馬\n• 1549 Book of Common Prayer (Cranmer)\n  — 英語禮儀首次確立\n• Mary I (1553–58) 短暫復辟天主教\n• Elizabeth I (1558–1603) 建立「折衷路線」\n  — 保留主教制與禮儀之美\n  — 改用英語——但詞曲仍可精緻\n\n🎵 Anglican 音樂特徵\n• 文字以英語為主——少數拉丁經文仍容許\n• 音樂清晰——一音一字為原則\n• 保留合唱傳統——大教堂與皇家禮拜堂\n\n📖 Service\n• Morning Prayer · Evening Prayer · Communion\n• 完整的 Service 設置包含全禮儀所需唱段\n• Great Service (大型) vs. Short Service (簡易)", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎼 Anthem 英文聖歌", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 英語相當於拉丁 motet 的曲種\n• Full anthem · 全合唱\n• Verse anthem · solo 與合唱交替 + 管風琴伴奏\n\n🌟 Thomas Tallis (ca. 1505–1585)\n• 跨四朝皇帝的 Chapel Royal 音樂家\n• 早期為拉丁 motet · 後適應英文\n• If ye love me (NAWM 60)\n  — 英文聖歌完美範例\n  — 四聲部 · 清晰的英語朗誦\n  — 簡潔而感人的聖禮用曲\n• Spem in alium——40 聲部合唱傑作\n\n🌟 William Byrd (ca. 1540–1623)\n• Tallis 的學生——Chapel Royal 成員\n• 私下是天主教徒——在新教國度下寫拉丁禮儀\n• Sing joyfully unto God (NAWM 61)\n  — 六聲部 full anthem\n  — 活潑、模仿、歡慶的經典英國合唱\n• Byrd · Tallis 兩人獨佔英國音樂出版特權 (1575)", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 8 · Catholic Reform & Trent ───────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("天主教改革 · 特倫特公會議", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("Council of Trent (1545–1563) · Reform of Liturgical Music", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 13, color: C.rose, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.parchment }, rounding: true });
  s.addText("⚖ Trent 對音樂的討論", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 1545–1563 間三次開會\n• 討論重點：彌撒中的音樂是否過度世俗化？\n\n改革派訴求\n• 文字必須清晰可辨\n• 驅逐以世俗歌曲 (chanson/madrigal) 為 cantus firmus 的彌撒\n• 拒絕不敬虔的表演\n• 驅逐所有複音？（激進派主張）\n\n最終決議（1562 Canon）\n• 允許複音音樂\n• 但文字必須可被聽懂\n• 禁止「猥褻」與「不潔」旋律\n• 並未指定特定風格\n\n💡 主教與地方執行\n• Cardinal Carlo Borromeo 推動改革——\n  在米蘭積極執行「清晰」原則\n• 帝國教堂繼續多樣風格", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.parchment }, rounding: true });
  s.addText("📖 Palestrina 傳說", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("流傳已久的故事\n• Palestrina 為拯救複音音樂免被禁——\n  寫作《Missa Papae Marcelli》（教宗馬切魯斯彌撒）\n• 據說此曲以清晰文字說服樞機團允許複音\n\n事實上\n• 此彌撒實為 1562–3 年創作\n• 確實展現「清晰文字」的典範風格\n• 但並非單一轉折事件——\n  清晰文字風格已是當時改革共識\n\n🎭 象徵意義\n• 19 世紀以來被視為「天主教音樂的救主」\n• Pfitzner 1917 歌劇《Palestrina》將此傳說戲劇化\n• 影響所謂 Palestrina style 的學院化地位\n\n📚 後續\n• 此後數百年 Palestrina 對位成為神學院標準教材\n• Fux《Gradus ad Parnassum》(1725) 以其為範本", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 9 · Palestrina · Life & Style ─────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Giovanni Pierluigi da Palestrina", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("ca. 1525–1594 · 文藝復興天主教音樂的典範", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("⛪ 生平與職位", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 出生於羅馬近郊 Palestrina 鎮\n• 青年期在羅馬各大聖堂任 maestro di cappella\n\n🏛 主要任職\n• Cappella Giulia (St. Peter's) · 1551–\n• Sistine Chapel · 1554（因已婚被迫離開）\n• St. John Lateran · 1555–1560\n• Santa Maria Maggiore · 1561–1566\n• 重返 Cappella Giulia · 1571 直至離世\n\n📊 作品產量\n• 104 首 Masses（現存西方最多）\n• 約 375 首 motets\n• Lamentations · Magnificats · Offertories\n• 少量世俗 madrigals（後來本人對此表示後悔）", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎼 The Palestrina Style", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("標準化的文藝復興對位——後世學理典範\n\n旋律\n• 以音階級進為主、跳進後必級進回\n• 跳進上限為完全五度或八度\n• 節奏平滑、避免突兀\n\n和聲與不協和\n• 所有不協和皆為 passing · neighbor · suspension\n• Suspension 嚴格準備與解決\n• 終止式以 authentic cadence 為主\n\n織度\n• 4–6 聲部 · 每聲部平等\n• 模仿進入——各聲部模仿主題\n• 同節奏段落強調重要字句\n• 文字清晰度極高\n\n美學\n• 不誇張 · 不戲劇 · 內斂崇高\n• 教會聲樂的「純粹」理想\n• 成為「a cappella 風格」的代名詞", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 10 · Pope Marcellus Mass (NAWM 62) ───────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("教宗馬切魯斯彌撒 (NAWM 62)", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("Missa Papae Marcelli · Credo", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 13, color: C.rose, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.parchment }, rounding: true });
  s.addText("📜 背景", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 1562–3 年創作 · 1567 收入第二部 Mass 集出版\n• 獻給短命的 Marcellus II (僅在位 3 週, 1555)\n• 六聲部 · SATTBB\n• 使用 \"paraphrase\" 技法而非固定 cantus firmus\n• 非以任何世俗旋律為素材——符合 Trent 改革精神\n\n🎵 Credo 的特徵\n• 長文字需盡可能清晰\n• 大部分使用同節奏 (homorhythm)\n• 文字逐字宣告——聲部幾乎同時換字\n• 重要教義片段 (et incarnatus, crucifixus) 有顯著靜思感\n• 模仿段落較少——優先清晰度\n\n📊 織度分析\n• 每段獨立分節——與各行 Credo 文本對應\n• 聲部組合常變化——提供對比\n• SA · TB · 全體 · 六聲部輪流", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.parchment }, rounding: true });
  s.addText("💡 為何成為經典？", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 完美回應 Trent 會議的關切\n• 在「清晰文字」與「多聲藝術」間取得平衡\n• 不僅宣告詞意——同時提供神聖的美感\n\n📚 歷史影響\n• 1725 Fux 以 Palestrina 為對位法典範\n  (Gradus ad Parnassum)\n• 羅馬公會成為其永久保守者\n• 19 世紀 Cecilian Movement 視其為唯一正統\n• 20 世紀對位法教材仍多取自 Palestrina 風格\n\n🎭 文化符碼\n• Palestrina = 「純粹教會音樂」的代名詞\n• Pfitzner 1917《Palestrina》歌劇將其戲劇化\n• Prokofiev · Stravinsky · Pärt 都曾研究其風格\n\n🌿 對 20 世紀極簡主義與新教會音樂的啟發", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 11 · Spain & New World ────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("西班牙與新世界", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Victoria · Morales · New World Polyphony · Hernando Franco", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🇪🇸 Victoria & Morales", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("🌟 Tomás Luis de Victoria (1548–1611)\n• 出生 Avila——年輕時至羅馬學習\n• 在羅馬 German College 任職——可能師事 Palestrina\n• 後返西班牙任王室禮拜堂樂長\n• 僅寫宗教音樂——近 180 首作品\n• 風格：更熾烈、更神秘——比 Palestrina 多情感深度\n\n🌟 O magnum mysterium (NAWM 63)\n• 四聲部聖誕 motet · 1572\n• 開頭 fifths · octaves 表達「偉大奧祕」\n• 文字與音樂完美對應\n• 最著名的文藝復興 motet 之一\n\n🎵 Cristóbal de Morales (ca. 1500–1553)\n• 西班牙最早的國際知名作曲家\n• 亦曾任羅馬教廷合唱團\n• 25 首 Masses · 110 首 motets\n• 為 Victoria 鋪路", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🌎 New World Polyphony", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 16 世紀西班牙征服使複音音樂首度抵達美洲\n• 墨西哥城 (1521) · 秘魯 Cuzco · 其他殖民據點\n• 原住民子弟入修道院——快速掌握拉丁文與五線譜\n\n📖 教學與作曲\n• 方濟會與道明會傳教士為教導\n• 編寫聖詠本、手寫譜冊普及各大聖堂\n• 原住民作曲家與西班牙樂長並肩工作\n\n🌟 Hernando Franco (1532–1585)\n• 生於西班牙 Galicia · 1554 抵達墨西哥\n• 1575–1585 任 Mexico City 主教座堂樂長\n• 以其 Magnificat 與 Salve Regina 聞名\n\n🎵 Salve regina (NAWM 64)\n• 拉丁 Marian 安提芬的複音設置\n• 展現歐洲風格如何在美洲落地生根\n• 織度近似 Palestrina——音樂並未「美洲化」\n• 是歐洲音樂首次成為全球語言的印證", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 12 · Lassus & Germany · Jewish Music ──────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("Lassus · 德國 · 猶太音樂", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("Orlande de Lassus · Jacob Handl · Jewish Liturgical Music", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 13, color: C.rose, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.parchment }, rounding: true });
  s.addText("🌟 Orlande de Lassus (1532–1594)", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• Franco-Flemish——任職 Munich Bavarian 宮廷 1556–94\n• 國際主義者：拉丁、義、法、德四語創作\n• 2,000+ 首作品（Masses, motets, chansons, madrigals, Lieder）\n• 天主教徒但作品為新舊兩教共享\n\n🎵 Cum essem parvulus (NAWM 65)\n• 六聲部拉丁 motet\n• 取自《林前 13:11》——「我小時候...」\n• 高聲部代表童聲 · 低聲部代表成人\n• 織度對比具戲劇感——體現文字意義\n• 代表 Lassus 文字敏感度的巔峰\n\n📊 影響\n• Jesuit 教育體系推廣其作品至歐洲各地\n• 承接 Palestrina 清晰原則——同時保持表達力\n• 是 16 世紀後半最重要的國際音樂家", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.parchment }, rounding: true });
  s.addText("🕍 中東歐與猶太音樂", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("🎼 Jacob Handl / Gallus (1550–1591)\n• 斯洛維尼亞出身——活躍於 Prague、Olomouc\n• Opus musicum (1586–91) · 4 冊全年度 motet 集\n• 同時為拉丁及斯拉夫天主教禮儀提供音樂\n• 風格融合 Palestrina 清晰與威尼斯雙合唱\n\n🌟 波蘭、波希米亞、匈牙利\n• 新教與天主教並存——音樂與義大利往來密切\n• Mikołaj Gomółka（波蘭）亦寫詩篇集\n\n🕎 Jewish Music\n• 猶太會堂禮拜持續口傳吟唱——無記譜傳統\n• Ashkenazic (德語區) · Sephardic (伊比利半島) 兩大支流\n• Salamone Rossi (1570–1630) · 曼圖阿宮廷猶太音樂家\n  — 將文藝復興多聲部用於希伯來詩篇（Hashirim asher lishlomo, 1623）\n  — 首位有記譜的猶太複音作曲家\n• 象徵文藝復興時期音樂與宗教身分的多樣性", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 13 · Byrd Mass for 4 Voices + Legacy ──────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Byrd 的四聲部彌撒 · 宗教音樂的遺產", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Byrd · Mass for Four Voices (NAWM 66) · The Legacy of 16c Sacred Music", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 12, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎵 Byrd Mass for Four Voices", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 1592–1593 出版——ostensibly 無版權頁\n• 私下為英國仍信天主教的家庭秘密舉行彌撒而作\n• 三冊 Masses（3、4、5 聲部）之中最簡潔一本\n\n🎼 風格\n• 開頭模仿進入 · 典型文藝復興織度\n• 避免冗長——適合小型私人場合\n• Credo 採精簡、緊湊寫法\n  — 部分段落同節奏、部分段落模仿\n  — 模仿 Taverner 舊式英國傳統\n\n📖 Elizabeth I 時代的天主教徒\n• 私下舉行彌撒為違法——可被罰款或下獄\n• Byrd 因受 Elizabeth 寵愛得以繼續工作\n• 其 Masses 僅在鄉村莊園的小禮拜堂使用\n\n🌿 同時 Byrd 亦創作英文 anthem\n— 兩教身份共存於一位作曲家之中", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("📜 The Legacy of 16c Sacred Music", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 16 世紀是西方「宗教音樂」最多樣的一個世紀\n• 一個西歐四種神學 · 四種敬拜 · 四種音樂\n\n📊 長遠影響\n• Lutheran chorale → Bach 聖詠和聲與清唱劇\n• Calvinist psalmody → 北美與英語聖詩傳統\n• Anglican anthem → Purcell · Handel · 至 20 世紀\n• Palestrina 對位 → 神學院與音樂學院必修\n\n🎭 音樂作為宗教認同的標記\n• 每一教派以「獨特的聲音」界定自身\n• 音樂風格帶有神學意涵——不只是美學選擇\n\n🌐 全球化的起點\n• 西班牙征服使歐洲多聲部抵達美洲\n• 傳教士將拉丁禮儀帶到中國、日本、菲律賓\n• 文藝復興音樂首次成為「全球語言」\n\n💡 為 17 世紀巴洛克教會音樂鋪路\n• Schütz · Frescobaldi · Purcell · Bach 皆承襲此基礎", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
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

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🔑 Key Terms", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Reformation · Counter-Reformation\n• Lutheran · Calvinist · Anglican · Catholic\n• chorale · Kontrafaktur · bar form\n• Achtliederbuch · Geystliche gesangk Buchleyn\n• metrical psalm · Geneva Psalter\n• Old 100th · psalmody\n• Book of Common Prayer · Anglican Service\n• Great Service · Short Service\n• full anthem · verse anthem\n• Council of Trent · Canon on music\n• Missa Papae Marcelli · paraphrase mass\n• Palestrina style · stile antico\n• O magnum mysterium · Hernando Franco\n• New World polyphony · Mission music\n• Chapel Royal · maestro di cappella\n• Salamone Rossi · Hashirim asher lishlomo\n• Cecilian movement · Gradus ad Parnassum", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("📚 Further Reading & 🎧 Listening", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Robin Leaver. Luther's Liturgical Music (2007)\n• Jesse Rodin. Josquin's Rome (2012)\n• Craig Monson. Disembodied Voices (1995)\n• Kerala J. Snyder. The Chorale (2008)\n• Iain Fenlon (ed.). The Renaissance\n\n🎧 NAWM 精選聆聽 (YouTube)\n• 58 · Luther/Walter · Ein feste Burg  youtu.be/KPEhnb3utSM\n• 59 · Bourgeois · Psalm 134  youtu.be/qgMLHdUrMHQ\n• 60 · Tallis · If ye love me (Tenebrae)  youtu.be/HI5Y9l2NHIo\n• 61 · Byrd · Sing joyfully (Tallis Scholars)  youtu.be/9uK9nVVbGHw\n• 62 · Palestrina · Missa Papae Marcelli (Tallis Scholars)  youtu.be/6RfiPXiXneY\n• 63 · Victoria · O magnum mysterium  youtu.be/YCaHfnRqboU\n• 64 · Franco · Salve regina (Tallis Scholars)  youtu.be/1GD6lxHMnyA\n• 65 · Lassus · Cum essem parvulus (Cappella Amsterdam)  youtu.be/JWLhbJPQZLM\n• 66 · Byrd · Mass for 4 voices Credo (Tallis Scholars)  youtu.be/rVG67CWoslQ", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1,
  });
}

pres.writeFile({ fileName: "Ch11_Reformation.pptx" })
  .then(fn => console.log(`✅ ${fn} created successfully`))
  .catch(err => console.error("❌ Error:", err));
