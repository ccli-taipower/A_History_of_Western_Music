const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 19: German Composers of the Late Baroque";
pres.author = "A History of Western Music, 10th ed.";

// Deep navy & amber — German Baroque gravitas
const C = {
  darkBg:   "0F1A2E",
  gold:     "D4A030",
  cream:    "F5F0E0",
  navy:     "1A2A4A",
  steel:    "3A4A6A",
  darkText: "1A1A2E",
  lightText:"F5F0E0",
  sand:     "E8D8B8",
  slate:    "2A3A5A",
  amber:    "D4A030",
  copper:   "B87333",
};

function darkSlide(p) { const s = p.addSlide(); s.background = { color: C.darkBg }; return s; }
function lightSlide(p) { const s = p.addSlide(); s.background = { color: C.cream }; return s; }
function topBar(s, c) { s.addShape(pres.ShapeType.rect, { x:0, y:0, w:"100%", h:0.12, fill:{color:c||C.gold} }); }
function bottomBar(s, c) { s.addShape(pres.ShapeType.rect, { x:0, y:5.5, w:"100%", h:0.125, fill:{color:c||C.gold} }); }

// ── SLIDE 1 · Title ─────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  s.addShape(pres.ShapeType.rect, { x:0, y:0, w:"100%", h:0.15, fill:{color:C.gold} });
  s.addShape(pres.ShapeType.rect, { x:0, y:5.47, w:"100%", h:0.155, fill:{color:C.gold} });

  s.addText("A HISTORY OF WESTERN MUSIC · TENTH EDITION", {
    x:0.5, y:0.45, w:9, h:0.35, fontSize:18, color:C.sand, charSpacing:3, align:"center", fontFace:"Georgia",
  });
  s.addText("CHAPTER 19", {
    x:0.5, y:1.0, w:9, h:0.55, fontSize:24, color:C.gold, bold:true, align:"center", fontFace:"Georgia", charSpacing:6,
  });
  s.addText("GERMAN COMPOSERS\nOF THE LATE BAROQUE", {
    x:0.3, y:1.7, w:9.4, h:1.6, fontSize:38, color:C.lightText, bold:true, align:"center", fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:2.5, y:3.5, w:5, h:0.04, fill:{color:C.gold} });
  s.addText("Telemann · J. S. Bach · Handel · Mixed Taste · Synthesis of Styles", {
    x:0.4, y:3.7, w:9.2, h:0.4, fontSize:16, color:C.sand, align:"center", fontFace:"Georgia",
  });
  s.addText("Textbook pp. 424–453", {
    x:0.5, y:4.8, w:9, h:0.3, fontSize:16, color:C.gold, align:"center", fontFace:"Calibri",
  });
}

// ── SLIDE 2 · Chapter Overview ──────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.navy); bottomBar(s, C.navy);

  s.addText("本章概覽 Chapter Overview", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.navy, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:9.2, h:0.03, fill:{color:C.sand} });

  const items = [
    { emoji:"■", title:"音樂背景 Contexts for Music", desc:"日耳曼諸邦、路德教會傳統、Collegium musicum" },
    { emoji:"■", title:"泰雷曼 Georg Philipp Telemann", desc:"最多產的巴洛克作曲家，融合多國風格的「混合品味」" },
    { emoji:"■", title:"巴赫：生平與管風琴 J. S. Bach: Life & Organ", desc:"威瑪、科騰、萊比錫三大時期；聖詠前奏曲、托卡塔與賦格" },
    { emoji:"■", title:"巴赫：鍵盤與聲樂 Bach: Keyboard & Vocal", desc:"平均律、清唱劇、受難曲、B小調彌撒" },
    { emoji:"■", title:"韓德爾：歌劇與神劇 Handel: Opera & Oratorio", desc:"義大利歌劇在倫敦、英語神劇《彌賽亞》《掃羅》" },
    { emoji:"■", title:"持久遺產 An Enduring Legacy", desc:"巴赫與韓德爾的歷史地位與後世影響" },
  ];
  items.forEach((item, i) => {
    const y = 1.05 + i * 0.73;
    s.addShape(pres.ShapeType.rect, { x:0.4, y, w:0.55, h:0.55, fill:{color:C.navy}, rounding:true });
    s.addText(item.emoji, { x:0.4, y, w:0.55, h:0.55, fontSize:22, align:"center", valign:"middle" });
    s.addText(item.title, { x:1.1, y, w:8.4, h:0.3, fontSize:16, bold:true, color:C.darkText, fontFace:"Georgia" });
    s.addText(item.desc, { x:1.1, y:y+0.28, w:8.4, h:0.27, fontSize:14, color:C.navy, fontFace:"Calibri" });
  });
}

// ── SLIDE 3 · Contexts for Music ────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("音樂背景", { x:0.4, y:0.18, w:9.2, h:0.5, fontSize:26, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText("Contexts for Music in German-Speaking Lands", { x:0.4, y:0.65, w:9.2, h:0.3, fontSize:14, color:C.sand, fontFace:"Georgia", align:"center" });

  const items = [
    { emoji:"■", label:"諸邦 States", zh:"德語區分為數百政治實體：奧地利、薩克森、布蘭登堡-普魯士及小邦", en:"Hundreds of entities: Austria, Saxony, Brandenburg-Prussia, principalities" },
    { emoji:"■", label:"教會 Church", zh:"路德教會城鎮由市議會聘音樂總監；教會是音樂演出核心場域", en:"Lutheran councils hired music directors; church was the central venue" },
    { emoji:"■", label:"貴族 Patrons", zh:"腓特烈大帝（長笛家）、貴族業餘音樂家眾多，效仿路易十四", en:"Frederick the Great (flutist); aristocrats emulated Louis XIV" },
    { emoji:"■", label:"英國 England", zh:"宮廷薪酬低，音樂家需兼職；倫敦公共音樂會蓬勃發展", en:"Low court salaries; London public concerts flourished" },
    { emoji:"■", label:"出版 Publishing", zh:"作曲家靠出版補充收入，但版稅概念尚未形成", en:"Extra income from publishing, but no royalty system existed" },
  ];
  items.forEach((item, i) => {
    const y = 1.1 + i * 0.85;
    s.addShape(pres.ShapeType.rect, { x:0.3, y, w:9.4, h:0.75, fill:{color:C.slate}, rounding:true });
    s.addText(item.emoji + " " + item.label, { x:0.5, y, w:2.2, h:0.75, fontSize:16, bold:true, color:C.gold, fontFace:"Georgia", valign:"middle" });
    s.addText(item.zh, { x:2.7, y, w:6.8, h:0.38, fontSize:14, color:C.lightText, fontFace:"Calibri", valign:"bottom" });
    s.addText(item.en, { x:2.7, y:y+0.38, w:6.8, h:0.37, fontSize:14, color:C.sand, fontFace:"Calibri", italic:true, valign:"top" });
  });
}

// ── SLIDE 4 · Telemann ──────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.navy); bottomBar(s, C.navy);

  s.addText("泰雷曼 Georg Philipp Telemann (1681–1767)", {
    x:0.4, y:0.18, w:9.2, h:0.55, fontSize:24, bold:true, color:C.navy, fontFace:"Georgia", align:"center",
  });
  s.addShape(pres.ShapeType.rect, { x:2.5, y:0.78, w:5, h:0.03, fill:{color:C.copper} });

  // Left: biography
  const bio = [
    "• 同代人推崇的最佳作曲家之一，作品超過三千首",
    "• 融合波蘭、法國、義大利風格，確立「混合品味」",
    "• 30 部歌劇、46 部受難曲、千首清唱劇",
    "• 在漢堡自行出版，創辦首份音樂期刊",
    "• 生前比巴赫更受歡迎，偏好相對簡潔的風格",
  ];
  s.addText(bio.join("\n"), {
    x:0.4, y:0.95, w:5.0, h:2.8, fontSize:14, color:C.darkText, fontFace:"Calibri", paraSpaceAfter:8,
  });

  // Right: Paris Quartets
  s.addShape(pres.ShapeType.rect, { x:5.6, y:0.95, w:4.1, h:4.2, fill:{color:C.navy}, rounding:true });
  s.addText("■ Paris Quartets", { x:5.8, y:1.0, w:3.7, h:0.35, fontSize:16, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 12 首四重奏（1730 & 1738）\n  長笛、小提琴、古中提琴 + 數字低音\n\n• 中提琴獨立於數字低音，\n  與獨奏平等\n\n• 第一集混合義大利協奏曲、\n  德國奏鳴曲、法國組曲各二首\n\n• 各體裁內亦融合跨國元素", {
    x:5.8, y:1.4, w:3.7, h:3.6, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:3,
  });
}

// ── SLIDE 5 · NAWM 101 Telemann ─────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("NAWM 101 泰雷曼《巴黎四重奏》第一號", { x:0.4, y:0.18, w:9.2, h:0.45, fontSize:22, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText("Telemann: Paris Quartet No. 1 — Concerto primo", { x:0.4, y:0.6, w:9.2, h:0.3, fontSize:14, color:C.sand, fontFace:"Georgia", align:"center" });

  // Left: Structure
  s.addShape(pres.ShapeType.rect, { x:0.3, y:1.0, w:4.6, h:4.3, fill:{color:C.slate}, rounding:true });
  s.addText("樂曲結構 Structure", { x:0.5, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• Presto 急板：利都奈羅形式，\n  義大利協奏曲風\n• 三件旋律樂器平等交替\n• 慢樂章：短小，連接急板與快板\n• Allegro 快板：吉格舞曲節奏，\n  法式迴旋曲形式\n• 可逆對位法處理三條旋律線\n• Presto = ritornello form\n  Allegro = French rondeau\n  → 義、德、法三國風格融於一曲", {
    x:0.5, y:1.4, w:4.2, h:3.6, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:3,
  });

  // Right: Style Features
  s.addShape(pres.ShapeType.rect, { x:5.1, y:1.0, w:4.6, h:4.3, fill:{color:C.slate}, rounding:true });
  s.addText("風格特色 Style Features", { x:5.3, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 對比織體：利都奈羅段 vs. 獨奏段交替\n\n• 五度循環進行、清晰調性結構\n\n• 小提琴風格的音型，\n  展現韋瓦第影響\n\n• 德國對位法為骨幹，\n  融入法義裝飾", {
    x:5.3, y:1.4, w:4.2, h:3.6, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:4,
  });
}

// ── SLIDE 6 · Bach: Life ────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.navy); bottomBar(s, C.navy);

  s.addText("巴赫生平 Johann Sebastian Bach (1685–1750)", {
    x:0.4, y:0.18, w:9.2, h:0.55, fontSize:24, bold:true, color:C.navy, fontFace:"Georgia", align:"center",
  });
  s.addShape(pres.ShapeType.rect, { x:2.5, y:0.78, w:5, h:0.03, fill:{color:C.copper} });

  const periods = [
    { label:"■ 早年\n1685–1708", items:["出生於艾森納赫音樂世家","呂訥堡求學，接觸法國風格"] },
    { label:"■ 威瑪\n1708–1717", items:["宮廷管風琴師→樂長","改編韋瓦第協奏曲"] },
    { label:"■ 科騰\n1717–1723", items:["宮廷樂長，專注器樂","布蘭登堡、無伴奏、平均律"] },
    { label:"■ 萊比錫\n1723–1750", items:["聖湯瑪斯領唱，管四座教堂","清唱劇、受難曲、B小調彌撒"] },
  ];
  periods.forEach((p, i) => {
    const x = 0.3 + i * 2.4;
    s.addShape(pres.ShapeType.rect, { x, y:0.95, w:2.25, h:4.3, fill:{color: i%2===0 ? C.navy : C.steel}, rounding:true });
    s.addText(p.label, { x:x+0.1, y:1.0, w:2.05, h:0.7, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
    s.addText(p.items.map(t => "• " + t).join("\n"), {
      x:x+0.1, y:1.8, w:2.05, h:3.3, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:6,
    });
  });
}

// ── SLIDE 7 · Bach: Organ Works ─────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("巴赫：管風琴作品", { x:0.4, y:0.18, w:9.2, h:0.45, fontSize:24, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText("Preludes, Fugues & Chorale Settings", { x:0.4, y:0.6, w:9.2, h:0.3, fontSize:14, color:C.sand, fontFace:"Georgia", align:"center" });

  // Left: Preludes & Fugues
  s.addShape(pres.ShapeType.rect, { x:0.3, y:1.0, w:4.6, h:4.3, fill:{color:C.slate}, rounding:true });
  s.addText("前奏曲與賦格 Preludes & Fugues", { x:0.5, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 布克斯特胡德開創自由幻想曲 + 賦格交替\n• 巴赫確立「前奏曲＋賦格」標準組合\n• A小調 BWV 543（NAWM 102）\n  前奏曲：托卡塔段與小提琴風格音型\n  賦格：類協奏曲快板，利都奈羅式結構\n• 韋瓦第影響：簡潔主題、和聲進行、五度圈", {
    x:0.5, y:1.4, w:4.2, h:3.5, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:5,
  });

  // Right: Chorale Settings
  s.addShape(pres.ShapeType.rect, { x:5.1, y:1.0, w:4.6, h:4.3, fill:{color:C.slate}, rounding:true });
  s.addText("聖詠前奏曲 Chorale Settings", { x:5.3, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 200+ 聖詠前奏曲，各種類型變化\n• 《管風琴小品集》Orgelbüchlein（威瑪）\n  45 首短聖詠前奏曲，兼具教學目的\n• Durch Adams Fall BWV 637（NAWM 103）\n  上聲部：聖詠旋律\n  低音：大跳＝亞當墮落\n  中音部：半音線＝蛇的誘惑\n  次中音：下滑＝罪惡哀傷\n  → Word-painting in every voice!\n• 晚期作品更宏大，概括表達取代細節描繪", {
    x:5.3, y:1.4, w:4.2, h:3.5, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:3,
  });
}

// ── SLIDE 8 · Bach: Keyboard Works ──────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.navy); bottomBar(s, C.navy);

  s.addText("巴赫：鍵盤音樂 Bach: Keyboard Works", {
    x:0.4, y:0.18, w:9.2, h:0.55, fontSize:24, bold:true, color:C.navy, fontFace:"Georgia", align:"center",
  });
  s.addShape(pres.ShapeType.rect, { x:2.5, y:0.78, w:5, h:0.03, fill:{color:C.copper} });

  // Left: Suites & Variations
  s.addText("組曲與變奏 Suites & Variations", { x:0.4, y:0.95, w:4.5, h:0.35, fontSize:16, bold:true, color:C.navy, fontFace:"Georgia" });
  s.addText("• 英國組曲、法國組曲及 Partitas\n  名稱非指國籍，皆融合法義德風格\n• 標準四舞曲：阿勒曼德、庫朗特、\n  薩拉班德、吉格\n• Goldberg Variations (1741)\n  30 段變奏，每三段含一首卡農\n• 晚期：《音樂獻禮》(1747)\n  ＆《賦格的藝術》(1740s)", {
    x:0.4, y:1.35, w:4.8, h:3.4, fontSize:14, color:C.darkText, fontFace:"Calibri", paraSpaceAfter:4,
  });

  // Right: Well-Tempered Clavier
  s.addShape(pres.ShapeType.rect, { x:5.4, y:0.95, w:4.3, h:4.3, fill:{color:C.navy}, rounding:true });
  s.addText("■ 平均律鍵盤曲集 WTC", { x:5.6, y:1.0, w:3.9, h:0.35, fontSize:16, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 兩冊（1722 & ca.1740），\n  各含 24 首前奏曲與賦格\n  涵蓋所有大小調（C 到 B）\n\n• 展示近平均律調音系統的可能性\n\n• 前奏曲：各含一項技巧或風格練習\n\n• 賦格：2–5 聲部，\n  從古老 ricercare 到韋瓦第風格\n\n• NAWM 105：升 D 小調賦格\n  展示倒影、增值、密接", {
    x:5.6, y:1.4, w:3.9, h:3.7, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:2,
  });
}

// ── SLIDE 9 · Bach: Vocal Works ─────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("巴赫：聲樂作品", { x:0.4, y:0.18, w:9.2, h:0.45, fontSize:24, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText("Cantatas, Passions & Mass in B Minor", { x:0.4, y:0.6, w:9.2, h:0.3, fontSize:14, color:C.sand, fontFace:"Georgia", align:"center" });

  // Left column: Cantatas + Passions
  s.addShape(pres.ShapeType.rect, { x:0.3, y:1.0, w:4.6, h:4.3, fill:{color:C.slate}, rounding:true });
  s.addText("清唱劇與受難曲", { x:0.5, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 約 200 首教會 + 20 首世俗清唱劇\n• 1723–1729 完成至少三個完整年度循環\n  （各 ~60 首）\n• 結合聖詠、宣敘調/詠嘆調、合唱\n\n受難曲 Passions：\n• 《約翰受難曲》(1724)\n  與《馬太受難曲》(1727)\n• 男高音演唱福音書敘事\n• 聖詠穿插反映會眾情感", {
    x:0.5, y:1.4, w:4.2, h:3.6, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:3,
  });

  // Right column: Mass + Synthesis
  s.addShape(pres.ShapeType.rect, { x:5.1, y:1.0, w:4.6, h:4.3, fill:{color:C.slate}, rounding:true });
  s.addText("B 小調彌撒 ＋ 藝術綜合", { x:5.3, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 唯一完整天主教彌撒常規曲\n  1747–49 編纂\n• 大量改編自先前清唱劇樂章\n• 1733 先呈 Kyrie & Gloria\n  給薩克森選侯\n• 新作混用 stile antico 與現代風格\n• 生前從未完整演出\n\n巴赫的藝術綜合：\n• 融合和聲與對位的衝突\n• 將所有時代風格吸收、發展至極致", {
    x:5.3, y:1.4, w:4.2, h:3.6, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:3,
  });
}

// ── SLIDE 10 · NAWM Bach Analysis ───────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.navy); bottomBar(s, C.navy);

  s.addText("NAWM 選曲分析 Bach Works Analysis", {
    x:0.4, y:0.18, w:9.2, h:0.55, fontSize:24, bold:true, color:C.navy, fontFace:"Georgia", align:"center",
  });
  s.addShape(pres.ShapeType.rect, { x:2.5, y:0.78, w:5, h:0.03, fill:{color:C.copper} });

  const works = [
    { nawm:"NAWM 102", title:"前奏曲與賦格 A 小調 BWV 543", desc:"前奏曲：托卡塔段與協奏曲獨奏\n賦格主題如小提琴，利都奈羅結構" },
    { nawm:"NAWM 103", title:"聖詠前奏曲《因亞當墮落》BWV 637", desc:"四聲部各以音型象徵歌詞意象\n低音大跳＝墮落，半音線＝蛇" },
    { nawm:"NAWM 104", title:"清唱劇 BWV 62 第一樂章", desc:"聖詠定旋律 + 協奏曲利都奈羅\n混合古老 cantus firmus 與現代風格" },
    { nawm:"NAWM 105", title:"平均律 I：升 D 小調賦格", desc:"展示倒影、增值、密接等賦格技法" },
    { nawm:"NAWM 106", title:"布蘭登堡協奏曲第五號", desc:"大鍵琴華彩段開創鍵盤協奏曲先河" },
  ];
  works.forEach((w, i) => {
    const y = 0.95 + i * 0.88;
    s.addShape(pres.ShapeType.rect, { x:0.4, y, w:1.3, h:0.78, fill:{color:C.navy} });
    s.addText(w.nawm, { x:0.45, y, w:1.2, h:0.78, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia", align:"center", valign:"middle" });
    s.addText(w.title, { x:1.85, y, w:3.5, h:0.78, fontSize:14, bold:true, color:C.darkText, fontFace:"Georgia", valign:"middle" });
    s.addText(w.desc, { x:5.5, y, w:4.2, h:0.78, fontSize:14, color:C.darkText, fontFace:"Calibri", valign:"middle" });
  });
}

// ── SLIDE 11 · Bach: Orchestral & Chamber ───────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("巴赫：管弦樂與室內樂", { x:0.4, y:0.18, w:9.2, h:0.45, fontSize:24, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText("Brandenburg Concertos & Chamber Music", { x:0.4, y:0.6, w:9.2, h:0.3, fontSize:14, color:C.sand, fontFace:"Georgia", align:"center" });

  // Left: Brandenburg
  s.addShape(pres.ShapeType.rect, { x:0.3, y:1.0, w:4.6, h:4.3, fill:{color:C.slate}, rounding:true });
  s.addText("布蘭登堡協奏曲 Brandenburg", { x:0.5, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 1721 年獻給布蘭登堡侯爵，共六首\n• 採義大利三樂章形式：快—慢—快\n• 第 3、6 號：多聲部弦樂 + 數字低音\n• 其餘以不同獨奏樂器組合對抗樂隊\n• 第 5 號：大鍵琴驚人的長華彩段\n  → 首開鍵盤協奏曲之先河\n• 利都奈羅素材擴展至樂章各段落", {
    x:0.5, y:1.4, w:4.2, h:3.5, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:5,
  });

  // Right: Chamber Music
  s.addShape(pres.ShapeType.rect, { x:5.1, y:1.0, w:4.6, h:4.3, fill:{color:C.slate}, rounding:true });
  s.addText("■ 室內樂與管弦樂", { x:5.3, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 6 首獨奏小提琴奏鳴曲與組曲\n  6 首大提琴無伴奏組曲\n• 無伴奏作品以和聲與對位\n  在單一樂器上製造幻象\n• 15 首大鍵琴奏鳴曲（小提琴/長笛）\n  多為教堂奏鳴曲四樂章\n• 4 首管弦組曲：法式序曲 + 義大利風\n• Collegium musicum（1729 起）\n  指揮大學生樂團演出協奏曲", {
    x:5.3, y:1.4, w:4.2, h:3.5, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:3,
  });
}

// ── SLIDE 12 · Handel: Life ─────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.navy); bottomBar(s, C.navy);

  s.addText("韓德爾生平 George Frideric Handel (1685–1759)", {
    x:0.4, y:0.18, w:9.2, h:0.55, fontSize:24, bold:true, color:C.navy, fontFace:"Georgia", align:"center",
  });
  s.addShape(pres.ShapeType.rect, { x:2.5, y:0.78, w:5, h:0.03, fill:{color:C.copper} });

  const periods = [
    { label:"DE 德國\n1685–1706", items:["哈勒出生，隨扎霍學習管風琴與作曲","1705：首部歌劇《阿爾米拉》漢堡首演"] },
    { label:"IT 義大利\n1706–1710", items:["結識梅迪奇王子、魯斯波利侯爵","義大利清唱劇、歌劇、教會音樂"] },
    { label:"GB 英國\n1710–1759", items:["1711《里納爾多》轟動倫敦","皇室資助：安妮女王→喬治一世→二世","1727 入英國籍"] },
  ];
  periods.forEach((p, i) => {
    const x = 0.3 + i * 3.15;
    const w = 2.95;
    s.addShape(pres.ShapeType.rect, { x, y:0.95, w, h:4.3, fill:{color: i%2===0 ? C.navy : C.steel}, rounding:true });
    s.addText(p.label, { x:x+0.1, y:1.0, w:w-0.2, h:0.7, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
    s.addText(p.items.map(t => "• " + t).join("\n"), {
      x:x+0.1, y:1.8, w:w-0.2, h:3.3, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:6,
    });
  });
}

// ── SLIDE 13 · Handel: Opera ────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("韓德爾：歌劇", { x:0.4, y:0.18, w:9.2, h:0.45, fontSize:24, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText("Handel: Italian Opera in London", { x:0.4, y:0.6, w:9.2, h:0.3, fontSize:14, color:C.sand, fontFace:"Georgia", align:"center" });

  // Left: Royal Academy
  s.addShape(pres.ShapeType.rect, { x:0.3, y:1.0, w:4.6, h:4.3, fill:{color:C.slate}, rounding:true });
  s.addText("皇家音樂學院 Royal Academy", { x:0.5, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 1718–19：60 位富紳合資成立\n  義大利歌劇團\n• 韓德爾任音樂總監，赴德國招募歌手\n• 閹伶塞內西諾為最大牌明星\n• 代表作：Radamisto (1720),\n  Giulio Cesare (1724), Rodelinda (1725)\n• 題材：羅馬英雄、十字軍、奇幻冒險\n• 1729 學院因財務壓力解散\n• 韓德爾自組新團，但競爭對手出現", {
    x:0.5, y:1.4, w:4.2, h:3.6, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:3,
  });

  // Right: Operatic Style
  s.addShape(pres.ShapeType.rect, { x:5.1, y:1.0, w:4.6, h:4.3, fill:{color:C.slate}, rounding:true });
  s.addText("歌劇風格 Operatic Style", { x:5.3, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 國際風格融合：法式序曲 +\n  義大利詠嘆調 + 德國對位\n• Secco 乾宣敘調：數字低音伴奏\n  Accompagnato：管弦樂伴奏\n• Da capo 詠嘆調種類豐富：\n  花腔、悲歌、田園、進行曲風\n• 場景複合體打破靜態詠嘆調框架\n• NAWM 107: Giulio Cesare Act II\n  克麗奧佩脫拉 V'adoro pupille\n  法國薩拉班德 + 義大利 da capo", {
    x:5.3, y:1.4, w:4.2, h:3.6, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:3,
  });
}

// ── SLIDE 14 · Handel: Oratorio & Instrumental ──────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.navy); bottomBar(s, C.navy);

  s.addText("韓德爾：神劇與器樂", { x:0.4, y:0.18, w:9.2, h:0.5, fontSize:24, bold:true, color:C.navy, fontFace:"Georgia", align:"center" });
  s.addText("Handel: Oratorio & Instrumental", { x:0.4, y:0.65, w:9.2, h:0.3, fontSize:14, color:C.steel, fontFace:"Georgia", align:"center" });

  // Left: Oratorio
  s.addText("英語神劇 English Oratorio", { x:0.4, y:1.0, w:4.5, h:0.35, fontSize:16, bold:true, color:C.navy, fontFace:"Georgia" });
  s.addText("• 1730s 韓德爾轉向英語神劇\n• 首部 Esther (1718/1732)\n• Saul (1739)：融合歌劇宣敘調、\n  詠嘆調與合唱（NAWM 108）\n• Messiah (1741)：最著名作品\n  非敘事而是沉思默想\n• 合唱為核心創新：參與、敘事、\n  評論如希臘悲劇\n• 不需舞台布景，吸引中產階級", {
    x:0.4, y:1.4, w:4.5, h:3.5, fontSize:14, color:C.darkText, fontFace:"Calibri", paraSpaceAfter:4,
  });

  // Right: Instrumental
  s.addShape(pres.ShapeType.rect, { x:5.4, y:1.0, w:4.3, h:4.2, fill:{color:C.navy}, rounding:true });
  s.addText("■ 器樂作品 Instrumental", { x:5.6, y:1.05, w:3.9, h:0.35, fontSize:16, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 鍵盤組曲兩集\n• ~20 獨奏奏鳴曲 + ~20 三重奏鳴曲\n  柯雷利影響，快樂章更華麗\n• Water Music (1717)\n  泰晤士河上為國王演奏\n• Music for Royal Fireworks (1749)\n• 12 首大協奏曲 Op. 6 (1739)\n  採柯雷利教堂奏鳴曲模式\n• 首創管風琴協奏曲（神劇中場）", {
    x:5.6, y:1.45, w:3.9, h:3.6, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:3,
  });
}

// ── SLIDE 15 · Key Terms & NAWM Listening ───────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("重要術語與 NAWM 聆聽", { x:0.4, y:0.18, w:9.2, h:0.48, fontSize:24, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText("Key Terms & NAWM Listening Guide", { x:0.4, y:0.62, w:9.2, h:0.3, fontSize:14, color:C.sand, fontFace:"Georgia", align:"center" });

  // Left: Key Terms
  s.addShape(pres.ShapeType.rect, { x:0.3, y:1.0, w:4.6, h:4.3, fill:{color:C.slate}, rounding:true });
  s.addText("■ 重要術語 Key Terms", { x:0.5, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• Mixed taste 混合品味\n• Collegium musicum 大學音樂社團\n• Chorale prelude 聖詠前奏曲\n• Orgelbüchlein 管風琴小品集\n• Well-Tempered Clavier 平均律\n• Cantata 清唱劇（教會/世俗）\n• Passion 受難曲（約翰/馬太）\n• Fore-imitation 先行模仿\n• Da capo aria 返始詠嘆調\n• Scene complex 場景複合體\n• Recitativo secco / accompagnato\n• Stile antico 古風格\n• Borrowing 借用（改編與仿作）", {
    x:0.5, y:1.4, w:4.2, h:3.8, fontSize:14, color:C.sand, fontFace:"Calibri", paraSpaceAfter:1,
  });

  // Right: NAWM Listening
  s.addShape(pres.ShapeType.rect, { x:5.1, y:1.0, w:4.6, h:4.3, fill:{color:C.slate}, rounding:true });
  s.addText("■ NAWM 聆聽 Listening", { x:5.3, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("101 Telemann Paris Quartet No. 1\nyoutube.com/watch?v=Dh2vPuqXw84\n102 Bach Prelude & Fugue A min\nyoutube.com/watch?v=_xhThihIIC4\n103 Bach Durch Adams Fall\nyoutube.com/watch?v=Z8Dpe0gjesg\n104 Bach Cantata BWV 62\nyoutube.com/watch?v=7oj63klgeEg\n106 Bach Brandenburg No. 5\nyoutube.com/watch?v=LHjbRMIIhuM\n107 Handel Giulio Cesare\nyoutube.com/watch?v=xRImsDQbaYY\n108 Handel Saul\nyoutube.com/watch?v=jQ9lz1fDkug", {
    x:5.3, y:1.4, w:4.2, h:3.8, fontSize:14, color:C.sand, fontFace:"Calibri", paraSpaceAfter:0,
  });
}

pres.writeFile({ fileName: "Ch19_German_Composers.pptx" })
  .then(() => console.log("Ch19_German_Composers.pptx created"))
  .catch(err => console.error("Error:", err));
