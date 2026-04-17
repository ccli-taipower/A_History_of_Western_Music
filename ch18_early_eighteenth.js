const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 18: The Early Eighteenth Century in Italy and France";
pres.author = "A History of Western Music, 10th ed.";

// Warm burgundy & gold — late Baroque elegance
const C = {
  darkBg:   "1E1528",
  gold:     "C8A030",
  cream:    "F5F0E0",
  burgundy: "6D2E46",
  rose:     "A26769",
  darkText: "2A1F30",
  lightText:"F5F0E0",
  sand:     "E8D8B8",
  violet:   "3A2050",
  amber:    "D4A030",
  plum:     "4A2040",
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
  s.addText("CHAPTER 18", {
    x:0.5, y:1.0, w:9, h:0.55, fontSize:24, color:C.gold, bold:true, align:"center", fontFace:"Georgia", charSpacing:6,
  });
  s.addText("THE EARLY EIGHTEENTH\nCENTURY IN ITALY\nAND FRANCE", {
    x:0.3, y:1.6, w:9.4, h:2.0, fontSize:36, color:C.lightText, bold:true, align:"center", fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:2.5, y:3.75, w:5, h:0.04, fill:{color:C.gold} });
  s.addText("Vivaldi · Couperin · Rameau · Concerto · Goûts réunis · Harmonic Theory", {
    x:0.4, y:3.9, w:9.2, h:0.4, fontSize:16, color:C.sand, align:"center", fontFace:"Georgia",
  });
  s.addText("Textbook pp. 402–423", {
    x:0.5, y:4.8, w:9, h:0.3, fontSize:16, color:C.gold, align:"center", fontFace:"Calibri",
  });
}

// ── SLIDE 2 · Chapter Overview ──────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.burgundy); bottomBar(s, C.burgundy);

  s.addText("本章概覽 Chapter Overview", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.burgundy, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:9.2, h:0.03, fill:{color:C.sand} });

  const items = [
    { emoji:"■", title:"變革中的歐洲 Europe in a Century of Change", desc:"理性時代的政治、經濟、社會背景 (pp. 402–404)" },
    { emoji:"■", title:"義大利的音樂 Music in Italy", desc:"歌劇正劇、音樂學院、閹人歌手、拿坡里與羅馬 (pp. 405–406)" },
    { emoji:"■", title:"維瓦第 Antonio Vivaldi", desc:"紅髮祭司、皮耶塔、協奏曲大師 (pp. 407–414)" },
    { emoji:"FR", title:"法國音樂 Music in France", desc:"風格融合、協調法義兩國品味 (pp. 415–416)" },
    { emoji:"■", title:"庫普蘭 François Couperin", desc:"大鍵琴組曲、性格小品、goûts réunis (pp. 416–418)" },
    { emoji:"■", title:"拉莫 Jean-Philippe Rameau", desc:"和聲理論革命家與歌劇作曲家 (pp. 418–423)" },
  ];
  items.forEach((item, i) => {
    const y = 1.05 + i * 0.73;
    s.addShape(pres.ShapeType.rect, { x:0.4, y, w:0.55, h:0.55, fill:{color:C.burgundy}, rounding:true });
    s.addText(item.emoji, { x:0.4, y, w:0.55, h:0.55, fontSize:22, align:"center", valign:"middle" });
    s.addText(item.title, { x:1.1, y, w:8.4, h:0.3, fontSize:16, bold:true, color:C.darkText, fontFace:"Georgia" });
    s.addText(item.desc, { x:1.1, y:y+0.28, w:8.4, h:0.27, fontSize:14, color:C.burgundy, fontFace:"Calibri" });
  });
}

// ── SLIDE 3 · Europe in a Century of Change ─────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("變革中的歐洲", { x:0.4, y:0.18, w:9.2, h:0.5, fontSize:26, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText("Europe in a Century of Change", { x:0.4, y:0.65, w:9.2, h:0.3, fontSize:14, color:C.sand, fontFace:"Georgia", align:"center" });

  const items = [
    { emoji:"■", label:"權力平衡", zh:"十八世紀初歐洲形成多國均勢：法國最大軍隊、英國最強海軍、奧地利崛起、普魯士建國", en:"Balance of power: France (army), Britain (navy), Austria rising, Prussia emerging" },
    { emoji:"■", label:"經濟成長", zh:"人口膨脹、農業改良、貿易擴張；中產階級壯大，城市化加速", en:"Population growth, improved agriculture, expanding trade; urban middle class grew" },
    { emoji:"■", label:"理性時代", zh:"啟蒙思想興起：伏爾泰、牛頓、笛卡兒；以理性解釋自然與社會", en:"Age of Reason: Voltaire, Newton, Descartes; reason and science" },
    { emoji:"■", label:"音樂需求", zh:"公眾對新音樂的需求不斷增長；音樂家成為教師、表演者、出版者", en:"Public demand for new music grew; musicians as teachers, performers, publishers" },
    { emoji:"■", label:"風格變遷", zh:"巴洛克晚期與古典早期風格重疊；「新」vs.「舊」的品味之爭貫穿整個世紀", en:"Late Baroque and early Classic styles overlapped; taste debates all century" },
  ];
  items.forEach((item, i) => {
    const y = 1.1 + i * 0.85;
    s.addShape(pres.ShapeType.rect, { x:0.3, y, w:9.4, h:0.75, fill:{color:C.violet}, rounding:true });
    s.addText(item.emoji + " " + item.label, { x:0.5, y, w:2.0, h:0.75, fontSize:16, bold:true, color:C.gold, fontFace:"Georgia", valign:"middle" });
    s.addText(item.zh, { x:2.5, y, w:7.0, h:0.38, fontSize:14, color:C.lightText, fontFace:"Calibri", valign:"bottom" });
    s.addText(item.en, { x:2.5, y:y+0.38, w:7.0, h:0.37, fontSize:14, color:C.sand, fontFace:"Calibri", italic:true, valign:"top" });
  });
}

// ── SLIDE 4 · Music in Italy ────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.burgundy); bottomBar(s, C.burgundy);

  s.addText("義大利的音樂 Music in Italy", { x:0.4, y:0.18, w:9.2, h:0.55, fontSize:26, bold:true, color:C.burgundy, fontFace:"Georgia", align:"center" });
  s.addShape(pres.ShapeType.rect, { x:2.5, y:0.78, w:5, h:0.03, fill:{color:C.rose} });

  // Left: Three Centers
  s.addText("三大音樂中心 Three Centers", { x:0.4, y:0.95, w:4.5, h:0.35, fontSize:16, bold:true, color:C.burgundy, fontFace:"Georgia" });
  const centers = [
    "■ 拿坡里 Naples：四所音樂學院訓練孤兒；閹人歌手盛行",
    "■ 羅馬 Rome：教廷限制歌劇；贊助者舉辦清唱劇與奏鳴曲",
    "■ 威尼斯 Venice：六家以上歌劇院，年均十部新歌劇",
    "■ A. Scarlatti 為拿坡里最重要的歌劇作曲家",
  ];
  s.addText(centers.join("\n"), {
    x:0.4, y:1.35, w:4.8, h:2.4, fontSize:14, color:C.darkText, fontFace:"Calibri", paraSpaceAfter:6,
  });

  // Right: Key Concepts
  s.addText("■ 關鍵概念 Key Concepts", { x:5.4, y:0.95, w:4.3, h:0.35, fontSize:16, bold:true, color:C.burgundy, fontFace:"Georgia" });
  const concepts = [
    { term:"Opera seria 正歌劇", desc:"梅塔斯塔齊奧編寫的嚴肅歌劇形式\n宣敘調與返始詠嘆調交替" },
    { term:"Castrato 閹人歌手", desc:"為保童聲而受閹割的男歌手\n法里內利最為著名" },
    { term:"Conservatory 音樂院", desc:"原為孤兒收容所，後發展為\n專業音樂教育機構" },
    { term:"Da capo aria 返始詠嘆調", desc:"ABA 形式；歌手在再現部\n自由添加裝飾" },
  ];
  concepts.forEach((c, i) => {
    const y = 1.35 + i * 0.95;
    s.addShape(pres.ShapeType.rect, { x:5.4, y, w:4.3, h:0.85, fill:{color:C.burgundy}, rounding:true });
    s.addText(c.term, { x:5.55, y, w:4.0, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
    s.addText(c.desc, { x:5.55, y:y+0.28, w:4.0, h:0.55, fontSize:14, color:C.lightText, fontFace:"Calibri" });
  });
}

// ── SLIDE 5 · Antonio Vivaldi ───────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("安東尼奧·維瓦第", { x:0.4, y:0.18, w:9.2, h:0.5, fontSize:26, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText("Antonio Vivaldi (1678–1741)", { x:0.4, y:0.65, w:9.2, h:0.3, fontSize:14, color:C.sand, fontFace:"Georgia", align:"center" });

  const items = [
    { emoji:"■", label:"紅髮祭司", text:"威尼斯出生，父為聖馬可教堂小提琴手；因紅髮得名 il prete rosso" },
    { emoji:"■", label:"皮耶塔", text:"1703–1740 任皮耶塔慈善院小提琴教師、樂長，為女學生創作大量音樂" },
    { emoji:"■", label:"產量驚人", text:"約 500 首協奏曲（含《四季》）、49 部歌劇、60 首宗教聲樂作品" },
    { emoji:"■", label:"歐洲影響", text:"作品經印刷傳遍歐洲；J. S. Bach 改編至少九首韋瓦第協奏曲為鍵盤曲" },
    { emoji:"■", label:"身後遺忘", text:"1741 年客死維也納，行乞者葬禮；1920s 重新發現手稿後聲譽復興" },
  ];
  items.forEach((item, i) => {
    const y = 1.1 + i * 0.85;
    s.addShape(pres.ShapeType.rect, { x:0.3, y, w:9.4, h:0.75, fill:{color:C.violet}, rounding:true });
    s.addText(item.emoji, { x:0.4, y, w:0.6, h:0.75, fontSize:22, align:"center", valign:"middle" });
    s.addText(item.label, { x:1.0, y, w:1.6, h:0.75, fontSize:16, bold:true, color:C.gold, fontFace:"Georgia", valign:"middle" });
    s.addText(item.text, { x:2.6, y, w:6.9, h:0.75, fontSize:14, color:C.lightText, fontFace:"Calibri", valign:"middle" });
  });
}

// ── SLIDE 6 · Concerto & Ritornello Form ────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.burgundy); bottomBar(s, C.burgundy);

  s.addText("協奏曲與利都奈羅形式", { x:0.4, y:0.18, w:9.2, h:0.5, fontSize:26, bold:true, color:C.burgundy, fontFace:"Georgia", align:"center" });
  s.addText("The Concerto & Ritornello Form", { x:0.4, y:0.65, w:9.2, h:0.3, fontSize:14, color:C.rose, fontFace:"Georgia", align:"center" });

  // Left column: Three-Movement Plan
  s.addText("三樂章結構", { x:0.4, y:1.0, w:4.5, h:0.35, fontSize:16, bold:true, color:C.burgundy, fontFace:"Georgia" });
  const mvts = [
    "第一樂章：快板（ritornello form）",
    "第二樂章：慢板（同調或近關係調）",
    "第三樂章：快板（較短，回到主調）",
    "此架構由 Albinoni 開創，Vivaldi 固定並推廣",
  ];
  s.addText(mvts.join("\n"), {
    x:0.4, y:1.4, w:4.5, h:1.8, fontSize:14, color:C.darkText, fontFace:"Calibri", paraSpaceAfter:8,
  });

  // Right column: Ritornello Form details
  s.addText("■ Ritornello Form 利都奈羅", { x:5.3, y:1.0, w:4.4, h:0.35, fontSize:16, bold:true, color:C.burgundy, fontFace:"Georgia" });
  const details = [
    "全奏利都奈羅與獨奏段落交替",
    "開頭利都奈羅由多個短小動機 (A, B, C...) 組成",
    "後續利都奈羅只用部分動機，且轉調",
    "獨奏段落以炫技樂句、音階、琶音為主",
    "首末利都奈羅在主調；中間在屬調或近關係調",
    "韋瓦第的利都奈羅形式成為此後協奏曲的標準",
  ];
  s.addText(details.map(d => "• " + d).join("\n"), {
    x:5.3, y:1.4, w:4.4, h:3.8, fontSize:14, color:C.darkText, fontFace:"Calibri", paraSpaceAfter:6,
  });
}

// ── SLIDE 7 · NAWM 98 Vivaldi Concerto ──────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("NAWM 98 作品分析", { x:0.4, y:0.18, w:9.2, h:0.45, fontSize:24, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText("Vivaldi: Concerto in A Minor, Op. 3, No. 6", { x:0.4, y:0.6, w:9.2, h:0.3, fontSize:14, color:C.sand, fontFace:"Georgia", align:"center" });

  // Left: First Movement
  s.addShape(pres.ShapeType.rect, { x:0.3, y:1.0, w:4.6, h:4.3, fill:{color:C.violet}, rounding:true });
  s.addText("第一樂章 First Movement (Allegro)", { x:0.5, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 開頭利都奈羅含 A, B, C, C' 四個動機\n• 後續利都奈羅僅用部分動機，轉調出現\n• 獨奏段落利用空弦技巧，展現炫技音型\n• 全奏與獨奏交替——尾聲時甚至交織", {
    x:0.5, y:1.4, w:4.2, h:2.0, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:6,
  });

  // Right: Third Movement
  s.addShape(pres.ShapeType.rect, { x:5.1, y:1.0, w:4.6, h:4.3, fill:{color:C.violet}, rounding:true });
  s.addText("第三樂章 Third Movement (Presto)", { x:5.3, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 利都奈羅動機為 ABABCDEF，結構更複雜\n• 利都奈羅本身轉調——打破常規慣例\n• 管弦樂與獨奏者交替呈示利都奈羅片段\n• 慢板樂章：低音沉默，獨奏以弦樂伴奏", {
    x:5.3, y:1.4, w:4.2, h:2.0, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:6,
  });
}

// ── SLIDE 8 · French Baroque Style ──────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.burgundy); bottomBar(s, C.burgundy);

  s.addText("法國巴洛克風格", { x:0.4, y:0.18, w:9.2, h:0.5, fontSize:26, bold:true, color:C.burgundy, fontFace:"Georgia", align:"center" });
  s.addText("French Baroque Style", { x:0.4, y:0.65, w:9.2, h:0.3, fontSize:14, color:C.rose, fontFace:"Georgia", align:"center" });

  // Left: France vs. Italy
  s.addText("法國 vs. 義大利風格之爭", { x:0.4, y:1.0, w:4.5, h:0.35, fontSize:16, bold:true, color:C.burgundy, fontFace:"Georgia" });
  const left = [
    "• 巴黎是法國唯一的文化中心",
    "• 皇家音樂學院壟斷歌劇首演權",
    "• 路易十五時代：贊助分散至沙龍與公共音樂會",
    "• 義大利音樂在法國既被歡迎也被抵制",
  ];
  s.addText(left.join("\n"), {
    x:0.4, y:1.4, w:4.5, h:2.2, fontSize:14, color:C.darkText, fontFace:"Calibri", paraSpaceAfter:8,
  });

  // Right: Goûts réunis
  s.addShape(pres.ShapeType.rect, { x:5.2, y:1.0, w:4.5, h:4.2, fill:{color:C.burgundy}, rounding:true });
  s.addText("■ Goûts réunis 風格融合", { x:5.4, y:1.05, w:4.1, h:0.35, fontSize:16, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 法國作曲家融合義大利體裁\n  （奏鳴曲、清唱劇、協奏曲）\n• Clérambault：混用呂利式宣敘調\n  與義大利風詠嘆調\n• Leclair：柯瑞里的純淨\n  與法國的優雅融為小提琴奏鳴曲\n• 庫普蘭：「完美的音樂將是\n  法義兩國風格的聯姻」", {
    x:5.4, y:1.45, w:4.1, h:3.6, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:4,
  });
}

// ── SLIDE 9 · François Couperin ─────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("弗朗索瓦·庫普蘭", { x:0.4, y:0.18, w:9.2, h:0.5, fontSize:26, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText("François Couperin (1668–1733)", { x:0.4, y:0.65, w:9.2, h:0.3, fontSize:14, color:C.sand, fontFace:"Georgia", align:"center" });

  const items = [
    { emoji:"■", label:"宮廷生涯", text:"國王御用管風琴師（St. Gervais）；出版教材 L'art de toucher le clavecin (1716)" },
    { emoji:"■", label:"風格融合", text:"宣稱完美音樂應為法義聯姻；組曲 Parnassus (1724) 讓呂利與柯瑞里會面" },
    { emoji:"■", label:"三重奏鳴曲", text:"法國最早、最重要的三重奏鳴曲作曲家；Les nations (1726) 含四組曲" },
    { emoji:"■", label:"大鍵琴組曲", text:"27 套 ordres（1713–1730 四冊）；以性格小品 (pièces de caractère) 著稱" },
    { emoji:"■", label:"Concerts", text:"12 首 concerts + Les goûts-réunis (1724)，標誌法義風格融合" },
  ];
  items.forEach((item, i) => {
    const y = 1.1 + i * 0.85;
    s.addShape(pres.ShapeType.rect, { x:0.3, y, w:9.4, h:0.75, fill:{color:C.violet}, rounding:true });
    s.addText(item.emoji, { x:0.4, y, w:0.6, h:0.75, fontSize:22, align:"center", valign:"middle" });
    s.addText(item.label, { x:1.0, y, w:1.6, h:0.75, fontSize:16, bold:true, color:C.gold, fontFace:"Georgia", valign:"middle" });
    s.addText(item.text, { x:2.6, y, w:6.9, h:0.75, fontSize:14, color:C.lightText, fontFace:"Calibri", valign:"middle" });
  });
}

// ── SLIDE 10 · NAWM 99 Couperin ─────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.burgundy); bottomBar(s, C.burgundy);

  s.addText("NAWM 99 作品分析", { x:0.4, y:0.18, w:9.2, h:0.45, fontSize:24, bold:true, color:C.burgundy, fontFace:"Georgia", align:"center" });
  s.addText("Couperin: Vingt-cinquième ordre (1730)", { x:0.4, y:0.6, w:9.2, h:0.3, fontSize:14, color:C.rose, fontFace:"Georgia", align:"center" });

  const pieces = [
    { title:"99a La visionnaire 幻視者", color:C.burgundy,
      desc:"法國序曲風格；附點節奏 + tirades（快速音階）\n暗示神秘的先知或靈視者" },
    { title:"99b La muse victorieuse 凱旋的繆思", color:C.plum,
      desc:"三拍子舞曲；快速跳躍音型、音域轉換\n描繪繆思克服技巧挑戰的勝利" },
    { title:"99c Les ombres errantes 徘徊的幽魂", color:C.violet,
      desc:"緩慢抒情；下行旋律、嘆息音型、半音不協和\n捕捉幽魂飄蕩的意象" },
  ];
  pieces.forEach((p, i) => {
    const y = 1.05 + i * 1.15;
    s.addShape(pres.ShapeType.rect, { x:0.4, y, w:9.2, h:1.0, fill:{color:p.color}, rounding:true });
    s.addText(p.title, { x:0.6, y, w:8.8, h:0.35, fontSize:16, bold:true, color:C.gold, fontFace:"Georgia" });
    s.addText(p.desc, { x:0.6, y:y+0.32, w:8.8, h:0.65, fontSize:14, color:C.lightText, fontFace:"Calibri" });
  });

  // Shared traits
  s.addShape(pres.ShapeType.rect, { x:0.4, y:4.55, w:9.2, h:0.8, fill:{color:C.darkBg}, rounding:true });
  s.addText("共通特色：混合法國裝飾音 (agréments)、呂利的序列手法、柯瑞里的五度圈和聲 — 體現 goûts réunis 理想", {
    x:0.6, y:4.6, w:8.8, h:0.7, fontSize:14, color:C.sand, fontFace:"Calibri", valign:"middle",
  });
}

// ── SLIDE 11 · Jean-Philippe Rameau ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("尚-菲利普·拉莫", { x:0.4, y:0.18, w:9.2, h:0.5, fontSize:26, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText("Jean-Philippe Rameau (1683–1764)", { x:0.4, y:0.65, w:9.2, h:0.3, fontSize:14, color:C.sand, fontFace:"Georgia", align:"center" });

  const items = [
    { emoji:"■", label:"早年", text:"生於第戎，管風琴師之子；在外省任管風琴師二十年後定居巴黎 (1722)" },
    { emoji:"■", label:"理論家", text:"1722 年出版《和聲論》(Traité de l'harmonie)，奠定現代和聲理論基礎" },
    { emoji:"■", label:"歌劇家", text:"首部歌劇 Hippolyte et Aricie (1733) 一鳴驚人；被攻擊為激進派" },
    { emoji:"■", label:"贊助者", text:"稅務官 La Pouplinière 提供沙龍舞台；1745 年國王授予年金" },
    { emoji:"■", label:"代表作", text:"5 齣悲劇歌劇、6 齣其他歌劇、7 齣芭蕾、大鍵琴曲集、三重奏鳴曲" },
  ];
  items.forEach((item, i) => {
    const y = 1.1 + i * 0.85;
    s.addShape(pres.ShapeType.rect, { x:0.3, y, w:9.4, h:0.75, fill:{color:C.violet}, rounding:true });
    s.addText(item.emoji, { x:0.4, y, w:0.6, h:0.75, fontSize:22, align:"center", valign:"middle" });
    s.addText(item.label, { x:1.0, y, w:1.6, h:0.75, fontSize:16, bold:true, color:C.gold, fontFace:"Georgia", valign:"middle" });
    s.addText(item.text, { x:2.6, y, w:6.9, h:0.75, fontSize:14, color:C.lightText, fontFace:"Calibri", valign:"middle" });
  });
}

// ── SLIDE 12 · Rameau's Traité de l'harmonie ────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.burgundy); bottomBar(s, C.burgundy);

  s.addText("拉莫的和聲理論革命", { x:0.4, y:0.18, w:9.2, h:0.5, fontSize:24, bold:true, color:C.burgundy, fontFace:"Georgia", align:"center" });
  s.addText("Rameau's Traité de l'harmonie", { x:0.4, y:0.62, w:9.2, h:0.3, fontSize:14, color:C.rose, fontFace:"Georgia", align:"center" });

  // Left: Core Concepts
  s.addText("核心概念 Core Concepts", { x:0.4, y:1.0, w:4.5, h:0.35, fontSize:16, bold:true, color:C.burgundy, fontFace:"Georgia" });
  const core = [
    "• 受笛卡兒與牛頓啟發，以理性原則解釋音樂",
    "• 三和弦與七和弦為和聲基本元素，源自自然泛音",
    "• 基礎低音：和弦無論轉位都保持同一根音身份",
    "• 創造 tonic、dominant、subdominant 等術語",
    "• 最強進行：屬七和弦 → 主和弦（V7 → I）",
    "• 轉調：承認調性可以改變，但一曲只有一個主調",
  ];
  s.addText(core.join("\n"), {
    x:0.4, y:1.4, w:4.8, h:3.2, fontSize:14, color:C.darkText, fontFace:"Calibri", paraSpaceAfter:6,
  });

  // Right: Impact
  s.addShape(pres.ShapeType.rect, { x:5.4, y:1.0, w:4.3, h:4.2, fill:{color:C.burgundy}, rounding:true });
  s.addText("■ 影響 Impact", { x:5.6, y:1.05, w:3.9, h:0.35, fontSize:16, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 拉莫理論成為十八世紀後期\n  教學音樂的主要範式\n\n• 今日每位音樂學生學習的和聲概念\n  絕大部分源自拉莫\n\n• 在理性時代的潮流下，音樂首次\n  被系統地視為一門「科學」", {
    x:5.6, y:1.5, w:3.9, h:3.5, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:4,
  });
}

// ── SLIDE 13 · Rameau's Operas & NAWM 100 ───────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("拉莫的歌劇 Rameau's Operas", { x:0.4, y:0.18, w:9.2, h:0.45, fontSize:24, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText("NAWM 100: Hippolyte et Aricie, Act IV finale (1733)", { x:0.4, y:0.6, w:9.2, h:0.3, fontSize:14, color:C.sand, fontFace:"Georgia", align:"center" });

  // Left: Opera Style
  s.addShape(pres.ShapeType.rect, { x:0.3, y:1.0, w:4.6, h:4.3, fill:{color:C.violet}, rounding:true });
  s.addText("■ 歌劇風格 Opera Style", { x:0.5, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 繼承呂利傳統：寫實唸誦、精確節奏記譜\n• 旋律植根於和聲——三和弦式旋律清晰\n• 管弦樂寫作獨具創意：描繪性配器\n  （雷聲、海浪、地震）\n• 合唱在法國歌劇中保持重要地位\n• NAWM 100：第四幕結尾\n  獵人嬉遊曲突轉為海怪場景", {
    x:0.5, y:1.4, w:4.2, h:3.5, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:5,
  });

  // Right: Lullistes vs. Ramistes
  s.addShape(pres.ShapeType.rect, { x:5.1, y:1.0, w:4.6, h:4.3, fill:{color:C.violet}, rounding:true });
  s.addText("■ Lullistes vs. Ramistes", { x:5.3, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• 呂利派攻擊拉莫音樂：\n  太難、太吵、太機械、不自然\n\n• 拉莫回應：「我並非模仿呂利，\n  而是效法自然本身」\n\n• 1750s 拉莫反成法國最偉大的\n  在世作曲家——反對者自相矛盾地\n  稱他為「法國音樂的冠軍」", {
    x:5.3, y:1.4, w:4.2, h:3.5, fontSize:14, color:C.lightText, fontFace:"Calibri", paraSpaceAfter:4,
  });
}

// ── SLIDE 14 · Timeline ─────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.burgundy); bottomBar(s, C.burgundy);

  s.addText("大事年表 Timeline", { x:0.4, y:0.18, w:9.2, h:0.55, fontSize:26, bold:true, color:C.burgundy, fontFace:"Georgia", align:"center" });
  s.addShape(pres.ShapeType.rect, { x:2.5, y:0.78, w:5, h:0.03, fill:{color:C.rose} });

  const events = [
    ["1701","普魯士成為王國"],
    ["1703","維瓦第任皮耶塔小提琴教師"],
    ["1711","維瓦第 L'estro armonico Op. 3 出版"],
    ["1716","庫普蘭《大鍵琴演奏法》"],
    ["1722","拉莫《和聲論》"],
    ["1724","庫普蘭 Les goûts-réunis 出版"],
    ["1725","維瓦第《四季》/ Concert Spirituel 成立"],
    ["1730","庫普蘭第四冊大鍵琴曲集"],
    ["1733","拉莫 Hippolyte et Aricie 首演"],
    ["1741","維瓦第客死維也納"],
    ["1745","路易十五授予拉莫年金"],
  ];
  events.forEach(([date, desc], i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 0.3 + col * 4.85;
    const y = 0.95 + row * 0.72;
    s.addShape(pres.ShapeType.rect, { x, y, w:0.95, h:0.58, fill:{color:C.burgundy} });
    s.addText(date, { x:x+0.05, y:y+0.06, w:0.85, h:0.46, fontSize:14, bold:true, color:C.lightText, align:"center", fontFace:"Georgia", valign:"middle" });
    s.addText(desc, { x:x+1.05, y, w:3.7, h:0.58, fontSize:14, color:C.darkText, fontFace:"Calibri", valign:"middle" });
  });
}

// ── SLIDE 15 · Key Terms & Listening ────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("關鍵術語與聆聽", { x:0.4, y:0.18, w:9.2, h:0.48, fontSize:24, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText("Key Terms & NAWM Listening", { x:0.4, y:0.62, w:9.2, h:0.3, fontSize:14, color:C.sand, fontFace:"Georgia", align:"center" });

  // Left: Key Terms
  s.addShape(pres.ShapeType.rect, { x:0.3, y:1.0, w:4.6, h:4.3, fill:{color:C.violet}, rounding:true });
  s.addText("■ 本章關鍵術語 Key Terms", { x:0.5, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("• Ritornello form 利都奈羅形式\n• Concerto 協奏曲\n• Opera seria 正歌劇\n• Castrato 閹人歌手\n• Conservatory 音樂學院\n• Da capo aria 返始詠嘆調\n• Goûts réunis 風格融合\n• Ordre 組曲\n• Pièce de caractère 性格小品\n• Agrément 裝飾音\n• Fundamental bass 基礎低音\n• Tonic · Dominant · Subdominant\n• Modulation 轉調\n• Tirade 快速音階", {
    x:0.5, y:1.4, w:4.2, h:3.8, fontSize:14, color:C.sand, fontFace:"Calibri", paraSpaceAfter:1,
  });

  // Right: NAWM Listening
  s.addShape(pres.ShapeType.rect, { x:5.1, y:1.0, w:4.6, h:4.3, fill:{color:C.violet}, rounding:true });
  s.addText("■ NAWM 聆聽 YouTube", { x:5.3, y:1.05, w:4.2, h:0.3, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia" });
  s.addText("NAWM 98  Vivaldi, Concerto Op. 3/6\nyoutu.be/BL-KzcwHDbY\n\nNAWM 99a  Couperin, La Visionnaire\nNAWM 99b  La Muse victorieuse\nNAWM 99c  Les ombres errantes\nyoutu.be/Gn5z0Hb-bb8\n\nNAWM 100  Rameau, Hippolyte\nyoutu.be/86rY74qyVSA", {
    x:5.3, y:1.4, w:4.2, h:3.8, fontSize:14, color:C.sand, fontFace:"Calibri", paraSpaceAfter:1,
  });
}

pres.writeFile({ fileName: "Ch18_Early_Eighteenth.pptx" })
  .then(() => console.log("Ch18_Early_Eighteenth.pptx created"))
  .catch(err => console.error("Error:", err));
