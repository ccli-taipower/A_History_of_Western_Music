const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 12: The Rise of Instrumental Music";
pres.author = "A History of Western Music, 10th ed.";

// Forest/earth palette — organic, instrumental warmth
const C = {
  darkBg:   "1E2D1E",
  gold:     "C8A020",
  cream:    "F5F0E0",
  forest:   "2A5A2A",
  olive:    "5A7A3A",
  darkText: "1E2D1E",
  lightText:"F5F0E0",
  sand:     "E8D8A8",
  slate:    "2A3A2A",
  amber:    "D4A020",
  brown:    "5A4030",
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
  s.addText("CHAPTER 12", {
    x:0.5, y:1.0, w:9, h:0.55, fontSize:24, color:C.gold, bold:true, align:"center", fontFace:"Georgia", charSpacing:6,
  });
  s.addText("THE RISE OF\nINSTRUMENTAL MUSIC", {
    x:0.3, y:1.6, w:9.4, h:2.0, fontSize:38, color:C.lightText, bold:true, align:"center", fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:2.5, y:3.75, w:5, h:0.04, fill:{color:C.gold} });
  s.addText("Susato · Narvaez · Byrd · Merulo · Andrea & Giovanni Gabrieli", {
    x:0.4, y:3.9, w:9.2, h:0.4, fontSize:18, color:C.sand, align:"center", fontFace:"Georgia",
  });
  s.addText("Textbook pp. 254-277", {
    x:0.5, y:4.8, w:9, h:0.3, fontSize:18, color:C.gold, align:"center", fontFace:"Calibri",
  });
}

// ── SLIDE 2 · Chapter Overview ──────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);

  s.addText("本章概覽 Chapter Overview", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.forest, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:9.2, h:0.03, fill:{color:C.sand} });

  const items = [
    "器樂的興起：1480年後器樂逐漸獲得獨立地位\nThe Rise: Instrumental music gains status after 1480",
    "樂器與合奏：家族樂器、consort、高音/低音之分\nInstruments & Ensembles: families, consorts, haut/bas",
    "舞曲音樂：pavane-galliard 配對、Susato Danserye\nDance Music: paired dances, Susato's Danserye (1551)",
    "聲樂改編與變奏：intabulation、Narvaez、Byrd\nArrangements & Variations: intabulation, Narvaez, Byrd",
    "抽象器樂：ricercar、fantasia、canzona、toccata\nAbstract Works: ricercar, fantasia, canzona, toccata",
    "威尼斯與聖馬可：Gabrieli 家族的輝煌音樂\nVenice & St. Mark's: the splendor of the Gabrielis",
  ];
  items.forEach((txt, i) => {
    const y = 1.05 + i * 0.73;
    s.addShape(pres.ShapeType.rect, { x:0.4, y, w:0.08, h:0.55, fill:{color:C.forest}, rounding:true });
    s.addText(txt, { x:0.65, y, w:8.9, h:0.65, fontSize:18, color:C.darkText, fontFace:"Calibri", valign:"middle" });
  });
}

// ── SLIDE 3 · Why Instrumental Music Rose ───────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("器樂為何興起？ Why Did Instrumental Music Rise?", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.gold, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.amber} });

  const bullets = [
    "1480年後教會、宮廷、城市、業餘愛好者日益培養器樂\nAfter 1480, churches, courts, cities & amateurs cultivated instruments",
    "器樂被記譜保存 — 顯示其地位提升，值得書寫傳播\nMore music written down — shows it was deemed worthy of preservation",
    "演奏者識譜能力提高，樂器教學書籍大量出版\nPerformers more literate; instruction books published in growing numbers",
    "新體裁不再依賴舞蹈或歌唱：variations、ricercar、toccata 等\nNew genres independent of dance/song: variations, ricercar, toccata",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.05, w:8.8, h:0.95, fontSize:19, color:C.lightText, fontFace:"Calibri",
      bullet:{ code:"25C6" }, lineSpacing:24,
    });
  });
}

// ── SLIDE 4 · Instruments & Ensembles ───────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);

  s.addText("樂器與合奏 Instruments & Ensembles", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.forest, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.sand} });

  const bullets = [
    "樂器教學書：Virdung《Musica getutscht》(1511)、Praetorius《Syntagma musicum》\nInstruction books: Virdung (1511), Praetorius Syntagma musicum (1618-20)",
    "家族樂器 (instrumental families)：同音色從高到低音，統一的音色\nInstrumental families: uniform timbre from soprano to bass range",
    "Consort（同族合奏）= 3-7件同族樂器；晚15世紀起流行\nConsort = 3-7 instruments of same family; popular from late 15th c.",
    "高音/低音之分：haut (alta) = 管樂舞會；bas = 室內弦樂/鍵盤\nHaut vs. bas: loud winds for dances; soft strings/keyboards for chambers",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.05, w:8.8, h:0.95, fontSize:18, color:C.darkText, fontFace:"Calibri",
      bullet:true, lineSpacing:24,
    });
  });
}

// ── SLIDE 5 · Wind, String & Keyboard ───────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("主要樂器類別 Major Instrument Categories", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.gold, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.amber} });

  // Three boxes
  const cats = [
    { title:"管樂 Winds", items:"直笛 recorder、橫笛 transverse flute\n蕭姆管 shawm、小號 cornett\nSackbut（長號前身）、Crumhorn" },
    { title:"弦樂 Strings", items:"魯特琴 lute（最流行家用樂器）\nVihuela（西班牙吉他形魯特琴）\nViol / viola da gamba（腿上弓弦）\n小提琴 violin（16世紀初出現）" },
    { title:"鍵盤 Keyboards", items:"管風琴 organ（教堂/正面風琴）\nClavichord（擊弦、可漸強）\nHarpsichord（撥弦，較響亮）\n= virginal / clavecin / clavicembalo" },
  ];
  cats.forEach((cat, i) => {
    const x = 0.3 + i * 3.2;
    s.addShape(pres.ShapeType.rect, { x, y:1.1, w:3.0, h:0.5, fill:{color:C.forest}, rounding:true });
    s.addText(cat.title, { x, y:1.1, w:3.0, h:0.5, fontSize:20, bold:true, color:C.lightText, align:"center", fontFace:"Georgia" });
    s.addText(cat.items, { x, y:1.7, w:3.0, h:3.5, fontSize:18, color:C.lightText, fontFace:"Calibri", lineSpacing:26 });
  });
}

// ── SLIDE 6 · Embellishment & Diminutions ───────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.brown); bottomBar(s, C.brown);

  s.addText("裝飾法與教學書 Embellishment & Treatises", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.brown, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.sand} });

  const bullets = [
    "16世紀演奏者被期待即興裝飾旋律（diminutions/divisions）\nPerformers expected to embellish melodies with diminutions/divisions",
    "Ganassi《Opera intitulata Fontegara》(1535)：直笛裝飾法經典\nGanassi's Fontegara (1535): classic recorder ornamentation manual",
    "Ortiz《Tratado de glosas》(1553)：古提琴裝飾的重要論著\nOrtiz's Tratado de glosas (1553): key treatise on viol diminutions",
    "裝飾保留原始旋律輪廓，避免平行五/八度等禁止進行\nOrnaments preserve melodic outline; avoid parallel 5ths/octaves",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.05, w:8.8, h:0.95, fontSize:18, color:C.darkText, fontFace:"Calibri",
      bullet:true, lineSpacing:24,
    });
  });
}

// ── SLIDE 7 · Five Categories of Instrumental Music ─────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("文藝復興器樂的五大類別 Five Categories", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.gold, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.amber} });

  const cats = [
    ["1", "舞曲音樂 Dance Music", "pavane、galliard、allemande、basse danse"],
    ["2", "聲樂多聲部改編 Arrangements of Vocal Music", "合奏或獨奏器樂版的經文歌/尚頌"],
    ["3", "既有旋律的設定 Settings of Existing Melodies", "聖歌管風琴曲、In Nomine 傳統"],
    ["4", "變奏曲 Variations", "在主題上的一系列裝飾變化"],
    ["5", "抽象器樂 Abstract Instrumental Works", "ricercar、fantasia、canzona、toccata、sonata"],
  ];
  cats.forEach(([num, title, sub], i) => {
    const y = 1.05 + i * 0.85;
    s.addShape(pres.ShapeType.rect, { x:0.4, y, w:0.55, h:0.55, fill:{color:C.forest}, rounding:true });
    s.addText(num, { x:0.4, y, w:0.55, h:0.55, fontSize:22, bold:true, color:C.lightText, align:"center", fontFace:"Georgia", valign:"middle" });
    s.addText(title, { x:1.1, y, w:8.5, h:0.3, fontSize:20, bold:true, color:C.gold, fontFace:"Georgia" });
    s.addText(sub, { x:1.1, y:y+0.32, w:8.5, h:0.28, fontSize:18, color:C.sand, fontFace:"Calibri" });
  });
}

// ── SLIDE 8 · Dance Music ───────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);

  s.addText("舞曲音樂 Dance Music", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.forest, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.sand} });

  const bullets = [
    "文藝復興社交舞極為重要 — 展現優雅、社交能力、健康\nSocial dance was vital — displayed grace, social skill, fitness",
    "舞曲常成對出現：慢二拍 + 快三拍（同旋律的變奏）\nDances often paired: slow duple + fast triple (varied from same tune)",
    "Basse danse（低舞）= 15-16世紀初最受歡迎的宮廷舞\nBasse danse = most popular courtly dance, 15th-early 16th c.",
    "Pavane（莊嚴二拍）+ Galliard（活潑三拍，含跳躍踢腿）\nPavane (stately duple) + Galliard (lively triple with hops & kicks)",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.05, w:8.8, h:0.95, fontSize:18, color:C.darkText, fontFace:"Calibri",
      bullet:true, lineSpacing:24,
    });
  });
}

// ── SLIDE 9 · Renaissance Dances Table ──────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("四種文藝復興舞曲 Four Renaissance Dances", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.gold, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.amber} });

  const rows = [
    [{ text:"舞曲 Dance", options:{bold:true,fontSize:19,color:C.darkBg,fontFace:"Georgia"} },
     { text:"拍子 Meter", options:{bold:true,fontSize:19,color:C.darkBg,fontFace:"Georgia"} },
     { text:"曲式 Form", options:{bold:true,fontSize:19,color:C.darkBg,fontFace:"Georgia"} },
     { text:"特徵 Character", options:{bold:true,fontSize:19,color:C.darkBg,fontFace:"Georgia"} }],
    [{ text:"Basse danse", options:{fontSize:18,color:C.lightText} },
     { text:"二/三拍", options:{fontSize:18,color:C.lightText} },
     { text:"反覆樂句", options:{fontSize:18,color:C.lightText} },
     { text:"莊嚴優雅 Stately, graceful", options:{fontSize:18,color:C.lightText} }],
    [{ text:"Pavane", options:{fontSize:18,color:C.lightText} },
     { text:"二拍 Duple", options:{fontSize:18,color:C.lightText} },
     { text:"AABBCC", options:{fontSize:18,color:C.lightText} },
     { text:"滑步莊嚴 Stately, gliding", options:{fontSize:18,color:C.lightText} }],
    [{ text:"Galliard", options:{fontSize:18,color:C.lightText} },
     { text:"三拍 Triple", options:{fontSize:18,color:C.lightText} },
     { text:"AABBCC", options:{fontSize:18,color:C.lightText} },
     { text:"活潑跳躍 Lively, hops & kicks", options:{fontSize:18,color:C.lightText} }],
    [{ text:"Allemande", options:{fontSize:18,color:C.lightText} },
     { text:"二拍 Duple", options:{fontSize:18,color:C.lightText} },
     { text:"2-3段反覆", options:{fontSize:18,color:C.lightText} },
     { text:"中庸簡潔 Moderate, upbeat start", options:{fontSize:18,color:C.lightText} }],
  ];
  const colW = [2.0, 1.8, 2.0, 3.8];
  const border = { pt:1, color:C.olive };
  s.addTable(rows, {
    x:0.4, y:1.1, w:9.6, colW,
    border, rowH:[0.55, 0.55, 0.55, 0.55, 0.55],
    autoPage:false,
  });
  // Header row fill
  s.addShape(pres.ShapeType.rect, { x:0.4, y:1.1, w:9.6, h:0.55, fill:{color:C.gold} });
  // Re-add table on top
  s.addTable(rows, {
    x:0.4, y:1.1, w:9.6, colW,
    border, rowH:[0.55, 0.55, 0.55, 0.55, 0.55],
    autoPage:false,
  });
}

// ── SLIDE 10 · NAWM 67 — Susato Danserye ────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);

  s.addText("NAWM 67 · Susato《Danserye》(1551)", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:26, bold:true, color:C.forest, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.sand} });

  const bullets = [
    "Tielman Susato 約1551年在安特衛普出版《Danserye》\nTielman Susato published Danserye in Antwerp, 1551",
    "收錄多對 pavane-galliard 舞曲，含 La dona 等配對\nContains pavane-galliard pairs including La dona (NAWM 66a-b)",
    "consort 合奏的理想曲目：可用 viols、recorders 等演奏\nIdeal consort music: playable on viols, recorders, or other instruments",
    "旋律在最高聲部，織體主要為和聲式（主調音樂）\nMelody in top voice; texture mostly homophonic",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.0, w:8.8, h:0.9, fontSize:18, color:C.darkText, fontFace:"Calibri",
      bullet:true, lineSpacing:24,
    });
  });

  s.addShape(pres.ShapeType.rect, { x:0.4, y:5.05, w:9.2, h:0.4, fill:{color:C.forest}, rounding:true });
  s.addText("YouTube: https://www.youtube.com/watch?v=Ln7ea-5dsoo", {
    x:0.6, y:5.05, w:8.8, h:0.4, fontSize:18, color:C.lightText, fontFace:"Calibri",
  });
}

// ── SLIDE 11 · Intabulation — Narvaez ───────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("移植改編 Intabulation", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.gold, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.amber} });

  const bullets = [
    "Intabulation = 將聲樂作品改編為魯特琴/鍵盤的 tablature\nIntabulation = arranging vocal pieces into lute/keyboard tablature",
    "最早的 tablature 記錄：Robertsbridge Codex（約1360年）\nEarliest surviving tablature: Robertsbridge Codex (ca. 1360)",
    "由於撥弦樂器無法持續音，改編者需用裝飾音填補\nPlucked instruments can't sustain; arrangers fill with ornamental runs",
    "16世紀大量 intabulation 出版，顯示其極高人氣\nGreat numbers published in 16th c., testifying to their popularity",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.05, w:8.8, h:0.95, fontSize:18, color:C.lightText, fontFace:"Calibri",
      bullet:{ code:"25C6" }, lineSpacing:24,
    });
  });
}

// ── SLIDE 12 · NAWM 68 — Narvaez ───────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.brown); bottomBar(s, C.brown);

  s.addText("NAWM 68 · Narvaez: Cancion Mille regretz", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:24, bold:true, color:C.brown, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.sand} });

  const bullets = [
    "Luys de Narvaez (fl. 1526-49)：西班牙 vihuela 大師\nLuys de Narvaez (fl. 1526-49): Spanish vihuela master",
    "1538年出版《Los seys libros del Delphin》：含最早的變奏曲集\nPublished Los seys libros del Delphin (1538): first published variation sets",
    "改編 Josquin 的四聲部《Mille regretz》為 vihuela 版\nIntabulation of Josquin's 4-voice Mille regretz for vihuela",
    "保留原作織體，加入 runs、turns、diminutions 增添趣味\nPreserves original texture; adds runs, turns & diminutions",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.0, w:8.8, h:0.9, fontSize:18, color:C.darkText, fontFace:"Calibri",
      bullet:true, lineSpacing:24,
    });
  });

  s.addShape(pres.ShapeType.rect, { x:0.4, y:5.05, w:9.2, h:0.4, fill:{color:C.brown}, rounding:true });
  s.addText("YouTube: https://www.youtube.com/watch?v=xKAeB5nDCmA", {
    x:0.6, y:5.05, w:8.8, h:0.4, fontSize:18, color:C.lightText, fontFace:"Calibri",
  });
}

// ── SLIDE 13 · Variations ───────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("變奏曲 Variations", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.gold, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.amber} });

  const bullets = [
    "變奏曲 = 16世紀的發明：在主題上呈現一系列裝飾變化\nVariation form = 16th-c. invention: series of embellished variants on a theme",
    "主題可為既有旋律、低音線、和聲進行或新創曲調\nTheme can be an existing tune, bass line, harmonic plan, or new melody",
    "最早的書面變奏：Dalza 的 pavane 變奏 (1508, 魯特琴)\nEarliest written variations: Dalza's pavane variations (1508, lute tablature)",
    "固定低音 (ostinato) 變奏：passamezzo antico/moderno、romanesca\nOstinato variations: passamezzo antico/moderno, romanesca, Ruggiero",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.05, w:8.8, h:0.95, fontSize:18, color:C.lightText, fontFace:"Calibri",
      bullet:{ code:"25C6" }, lineSpacing:24,
    });
  });
}

// ── SLIDE 14 · English Virginalists & Byrd ──────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);

  s.addText("英國維乃琴樂派 English Virginalists", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.forest, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.sand} });

  const bullets = [
    "Virginal = 英國小型撥弦鍵盤樂器（harpsichord 家族）\nVirginal = small English harpsichord-family keyboard instrument",
    "三大作曲家：William Byrd、John Bull、Orlando Gibbons\nThree masters: William Byrd, John Bull, Orlando Gibbons",
    "《Parthenia》(1613)：最早出版的 virginal 作品集\nParthenia (1613): first published collection of virginal music",
    "英國人偏好旋律變奏（vs. 西班牙/義大利的低音變奏）\nEnglish preferred melodic variation (vs. Spanish/Italian bass patterns)",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.0, w:8.8, h:0.9, fontSize:18, color:C.darkText, fontFace:"Calibri",
      bullet:true, lineSpacing:24,
    });
  });
}

// ── SLIDE 15 · NAWM 69 — Byrd ──────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("NAWM 69 · Byrd: John come kiss me now", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:26, bold:true, color:C.gold, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.amber} });

  const bullets = [
    "以流行歌曲為主題的 virginal 變奏曲\nVirginal variations on a popular song of the time",
    "旋律貫穿每一變奏，多在高音部或中音部\nSong melody present in every variation, mostly in treble or middle voice",
    "每段變奏引入新的動機或節奏型態，漸次加速\nEach variation introduces a new motivic/rhythmic figure, gradually quickening",
    "從四分音符→八分音符→十六分音符，形成節奏高潮\nPace accelerates: quarter → eighth → sixteenth notes, building rhythmic climax",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.05, w:8.8, h:0.95, fontSize:18, color:C.lightText, fontFace:"Calibri",
      bullet:{ code:"25C6" }, lineSpacing:24,
    });
  });

  s.addShape(pres.ShapeType.rect, { x:0.4, y:5.05, w:9.2, h:0.4, fill:{color:C.forest}, rounding:true });
  s.addText("YouTube: https://www.youtube.com/watch?v=DD7luwIuM40", {
    x:0.6, y:5.05, w:8.8, h:0.4, fontSize:18, color:C.lightText, fontFace:"Calibri",
  });
}

// ── SLIDE 16 · Abstract Works: Prelude, Fantasia, Ricercar ──────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.brown); bottomBar(s, C.brown);

  s.addText("前奏曲·幻想曲·利乃爾卡 Prelude, Fantasia, Ricercar", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:24, bold:true, color:C.brown, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.sand} });

  const bullets = [
    "Prelude / Fantasia / Ricercar：不依賴既有旋律的自由即興風格\nPrelude / Fantasia / Ricercar: free improvisatory style, no borrowed tune",
    "最早的 ricercar 為魯特琴即興小品；轉入鍵盤後加入模仿\nEarliest ricercari: brief lute improvisations; gained imitation on keyboard",
    "1540年後 ricercar 發展為連續模仿主題的「無詞經文歌」\nBy 1540, ricercar = successive imitative themes — a textless motet",
    "Luis Milan《El Maestro》(1536)：vihuela fantasia 的代表作\nLuis Milan's El Maestro (1536): exemplary vihuela fantasias",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.05, w:8.8, h:0.95, fontSize:18, color:C.darkText, fontFace:"Calibri",
      bullet:true, lineSpacing:24,
    });
  });
}

// ── SLIDE 17 · NAWM 70 — Andrea Gabrieli Ricercar ───────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("NAWM 70 · A. Gabrieli: Ricercar del 12 tuono", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:24, bold:true, color:C.gold, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.amber} });

  const bullets = [
    "Andrea Gabrieli (ca. 1532-1585)：威尼斯聖馬可教堂管風琴師\nAndrea Gabrieli (ca. 1532-1585): organist at St. Mark's, Venice",
    "此 ricercar 為連續模仿作品，結構如同無詞經文歌\nThis ricercar unfolds as successive imitation — a textless motet",
    "多個主題依次發展，各段以終止式劃分\nMultiple subjects developed in turn, sections defined by cadences",
    "展現鍵盤上對位法的精湛技巧\nDemonstrates mastery of contrapuntal writing for keyboard",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.05, w:8.8, h:0.95, fontSize:18, color:C.lightText, fontFace:"Calibri",
      bullet:{ code:"25C6" }, lineSpacing:24,
    });
  });

  s.addShape(pres.ShapeType.rect, { x:0.4, y:5.05, w:9.2, h:0.4, fill:{color:C.forest}, rounding:true });
  s.addText("YouTube: https://www.youtube.com/watch?v=DRYBJDFxQo8", {
    x:0.6, y:5.05, w:8.8, h:0.4, fontSize:18, color:C.lightText, fontFace:"Calibri",
  });
}

// ── SLIDE 18 · Canzona & Toccata ────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);

  s.addText("坎佐那與觸技曲 Canzona & Toccata", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.forest, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.sand} });

  // Two-column comparison
  s.addShape(pres.ShapeType.rect, { x:0.4, y:1.1, w:4.3, h:0.5, fill:{color:C.forest}, rounding:true });
  s.addText("Canzona 坎佐那", { x:0.4, y:1.1, w:4.3, h:0.5, fontSize:20, bold:true, color:C.lightText, align:"center", fontFace:"Georgia" });
  s.addText(
    "源自法國 chanson 的器樂改編\nOriginated from French chanson\n\n輕快、節奏鮮明、對位織體\nLight, rhythmic, contrapuntal\n\n典型開頭：長-短-短節奏型\nTypical opening: long-short-short\n\n多段對比樂段依次展開\nMultiple contrasting sections",
    { x:0.4, y:1.75, w:4.3, h:3.5, fontSize:18, color:C.darkText, fontFace:"Calibri", lineSpacing:24 }
  );

  s.addShape(pres.ShapeType.rect, { x:5.3, y:1.1, w:4.3, h:0.5, fill:{color:C.brown}, rounding:true });
  s.addText("Toccata 觸技曲", { x:5.3, y:1.1, w:4.3, h:0.5, fontSize:20, bold:true, color:C.lightText, align:"center", fontFace:"Georgia" });
  s.addText(
    "名稱源自 toccare（「觸摸」）\nName from Italian toccare (\"to touch\")\n\n即興風格的鍵盤獨奏曲\nImprovisatory-style keyboard solo\n\n展現演奏者的技巧與想像力\nShowcases performer's skill & imagination\n\n自由段落與模仿段落交替\nFree passages alternate with imitative",
    { x:5.3, y:1.75, w:4.3, h:3.5, fontSize:18, color:C.darkText, fontFace:"Calibri", lineSpacing:24 }
  );
}

// ── SLIDE 19 · NAWM 71 — Merulo Toccata ─────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("NAWM 71 · Merulo: Toccata IV (6th Mode)", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:24, bold:true, color:C.gold, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.amber} });

  const bullets = [
    "Claudio Merulo (1533-1604)：威尼斯聖馬可教堂管風琴師\nClaudio Merulo (1533-1604): organist at St. Mark's, Venice",
    "Toccata IV 收錄於1604年第二冊觸技曲集\nToccata IV from his second book of toccatas (1604)",
    "結構：自由和弦開頭 → 模仿中段 → 華麗自由尾聲\nStructure: free chordal opening → imitative middle → brilliant free closing",
    "利用管風琴持續音的能力：延留音、不協和音、音階跑動\nExploits organ's sustaining power: suspensions, dissonances, scale runs",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.05, w:8.8, h:0.95, fontSize:18, color:C.lightText, fontFace:"Calibri",
      bullet:{ code:"25C6" }, lineSpacing:24,
    });
  });

  s.addShape(pres.ShapeType.rect, { x:0.4, y:5.05, w:9.2, h:0.4, fill:{color:C.forest}, rounding:true });
  s.addText("YouTube: search \"Merulo Toccata quarta\"", {
    x:0.6, y:5.05, w:8.8, h:0.4, fontSize:18, color:C.lightText, fontFace:"Calibri",
  });
}

// ── SLIDE 20 · Venice & St. Mark's ──────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);

  s.addText("威尼斯與聖馬可教堂 Venice & St. Mark's", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.forest, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.sand} });

  const bullets = [
    "威尼斯 = 僅次於羅馬的義大利第二大城，貿易帝國\nVenice = second most important Italian city after Rome; trading empire",
    "政府大量投資公共音樂活動作為文化宣傳\nGovernment invested lavishly in public music as cultural propaganda",
    "聖馬可大教堂：拜占庭風格，義大利最令人嚮往的音樂職位\nSt. Mark's: Byzantine-style basilica, most coveted musical post in Italy",
    "1568年起建立常設器樂團：cornett、sackbut、violin、dulcian\nPermanent instrumental ensemble from 1568: cornetts, sackbuts, violins",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.05, w:8.8, h:0.95, fontSize:18, color:C.darkText, fontFace:"Calibri",
      bullet:true, lineSpacing:24,
    });
  });
}

// ── SLIDE 21 · The Gabrielis ────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("Gabrieli 家族 The Gabrielis at St. Mark's", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.gold, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.amber} });

  // Two columns
  s.addShape(pres.ShapeType.rect, { x:0.4, y:1.1, w:4.3, h:0.65, fill:{color:C.forest}, rounding:true });
  s.addText("Andrea Gabrieli\n(ca. 1532–1585)", { x:0.4, y:1.1, w:4.3, h:0.65, fontSize:18, bold:true, color:C.lightText, align:"center", fontFace:"Georgia", valign:"middle", lineSpacingMultiple:0.9 });
  s.addText(
    "聖馬可教堂管風琴師\nOrganist at St. Mark's\n\n大量多合唱團作品\nNumerous polychoral works\n\nRicercar、canzona、toccata\nKey organ genres",
    { x:0.4, y:1.9, w:4.3, h:2.85, fontSize:19, color:C.lightText, fontFace:"Calibri", lineSpacing:28 }
  );

  s.addShape(pres.ShapeType.rect, { x:5.3, y:1.1, w:4.3, h:0.55, fill:{color:C.amber}, rounding:true });
  s.addText("G. Gabrieli (1555–1612)", { x:5.3, y:1.1, w:4.3, h:0.55, fontSize:18, bold:true, color:C.darkBg, align:"center", fontFace:"Georgia", valign:"middle" });
  s.addText(
    "Andrea 的姪子，接任管風琴師\nAndrea's nephew; succeeded him\n\n~100 經文歌、37 合奏坎佐那、7 奏鳴曲\n~100 motets, 37 canzonas, 7 sonatas\n\n首創指定樂器、標示力度記號\nFirst to specify instruments & dynamics",
    { x:5.3, y:1.95, w:4.3, h:2.8, fontSize:18, color:C.lightText, fontFace:"Calibri", lineSpacing:26, margin:0 }
  );
}

// ── SLIDE 22 · Polychoral Music & Cori Spezzati ─────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);

  s.addText("多合唱團音樂 Polychoral Music (Cori spezzati)", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:24, bold:true, color:C.forest, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.sand} });

  const bullets = [
    "Cori spezzati = 分離合唱團：2-5組合唱團分置教堂不同位置\nCori spezzati = divided choirs: 2-5 groups placed in different locations",
    "Willaert 1550年出版雙合唱團詩篇集，引領風潮\nWillaert published double-choir psalms in 1550, sparking the fashion",
    "Andrea Gabrieli 將其發展為大型慶典的常規形式\nAndrea Gabrieli developed it as standard for grand ceremonies",
    "Giovanni Gabrieli 混合高低聲部與不同音色的樂器群\nGiovanni mixed high/low voices with diverse instrumental timbres",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.05, w:8.8, h:0.95, fontSize:18, color:C.darkText, fontFace:"Calibri",
      bullet:true, lineSpacing:24,
    });
  });
}

// ── SLIDE 23 · NAWM 72 — Sonata pian' e forte ───────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("NAWM 72 · G. Gabrieli: Sonata pian' e forte", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:24, bold:true, color:C.gold, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.amber} });

  const bullets = [
    "收錄於《Sacrae symphoniae》(1597)：音樂史上的里程碑\nFrom Sacrae symphoniae (1597): a landmark in music history",
    "最早指定樂器的合奏曲之一：第一組 cornett + 3 sackbut，第二組 violin + 3 sackbut\nAmong first to specify instruments: Choir 1 cornett + 3 sackbuts; Choir 2 violin + 3 sackbuts",
    "最早標示力度 pian (弱) 與 forte (強) 的作品之一\nAmong first to mark dynamics: pian (soft) & forte (loud)",
    "兩組對話、對比、合奏 — 純器樂作品的深度與多樣性\nTwo groups dialogue, contrast & join — depth & variety in pure instrumental music",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.05, w:8.8, h:0.95, fontSize:18, color:C.lightText, fontFace:"Calibri",
      bullet:{ code:"25C6" }, lineSpacing:22,
    });
  });

  s.addShape(pres.ShapeType.rect, { x:0.4, y:5.05, w:9.2, h:0.4, fill:{color:C.forest}, rounding:true });
  s.addText("YouTube: https://www.youtube.com/watch?v=QXRITlQBitc", {
    x:0.6, y:5.05, w:8.8, h:0.4, fontSize:18, color:C.lightText, fontFace:"Calibri",
  });
}

// ── SLIDE 24 · NAWM 73 — In ecclesiis ───────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);

  s.addText("NAWM 73 · G. Gabrieli: In ecclesiis", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:26, bold:true, color:C.forest, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.sand} });

  const bullets = [
    "壯麗的多合唱團經文歌（聲樂 + 器樂 + 管風琴）\nGrand polychoral motet: voices + instruments + organ",
    "獨唱、合唱、器樂群交替與合奏，展現威尼斯音響的極致\nSoloists, choir & instruments alternate and combine — Venetian splendor",
    "Alleluia 副歌貫穿全曲，形成統一結構\nAlleluia refrain runs throughout, creating structural unity",
    "代表文藝復興晚期到巴洛克早期的過渡風格\nRepresents the transitional style from late Renaissance to early Baroque",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.0, w:8.8, h:0.9, fontSize:18, color:C.darkText, fontFace:"Calibri",
      bullet:true, lineSpacing:24,
    });
  });

  s.addShape(pres.ShapeType.rect, { x:0.4, y:5.05, w:9.2, h:0.4, fill:{color:C.forest}, rounding:true });
  s.addText("YouTube: https://www.youtube.com/watch?v=Xf8oWn1Hj_8", {
    x:0.6, y:5.05, w:8.8, h:0.4, fontSize:18, color:C.lightText, fontFace:"Calibri",
  });
}

// ── SLIDE 25 · Timeline ─────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("年表 Timeline", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.gold, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.amber} });

  const events = [
    ["ca.1480s", "Martini《La Martinella》— 最早廣泛流傳的器樂作品"],
    ["1501", "Petrucci 出版《Odhecaton》— 首部活字印刷音樂集"],
    ["1507", "Petrucci 首部魯特琴 tablature 出版"],
    ["1511", "Virdung《Musica getutscht》— 首部樂器教學書"],
    ["1535", "Ganassi《Fontegara》— 直笛裝飾法"],
    ["1538", "Narvaez《Los seys libros del Delphin》— 含首部變奏曲集"],
    ["1551", "Susato《Danserye》— 安特衛普舞曲集"],
    ["1585-1612", "Giovanni Gabrieli 任職聖馬可教堂"],
    ["1597", "Gabrieli《Sacrae symphoniae》含 Sonata pian' e forte"],
    ["1604", "Merulo 觸技曲第二冊出版"],
    ["1613", "《Parthenia》— 首部維乃琴作品集出版"],
  ];

  events.forEach(([yr, desc], i) => {
    const y = 1.05 + i * 0.39;
    s.addText(yr, { x:0.4, y, w:1.5, h:0.36, fontSize:18, bold:true, color:C.gold, fontFace:"Georgia", align:"right" });
    s.addText(desc, { x:2.1, y, w:7.5, h:0.36, fontSize:18, color:C.lightText, fontFace:"Calibri" });
  });
}

// ── SLIDE 26 · Key Terms ────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.brown); bottomBar(s, C.brown);

  s.addText("重要術語 Key Terms", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.brown, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.sand} });

  const terms = [
    ["Consort 同族合奏", "同一樂器家族的合奏組合"],
    ["Diminutions 裝飾音", "將長音符分割為短裝飾音群"],
    ["Intabulation 移植改編", "聲樂作品改編為樂器 tablature"],
    ["Variations 變奏曲", "在主題上的一系列裝飾變化"],
    ["Ricercar 利乃爾卡", "連續模仿主題的器樂曲（無詞經文歌）"],
    ["Canzona 坎佐那", "源自 chanson 的對位器樂曲"],
    ["Toccata 觸技曲", "即興風格的鍵盤獨奏炫技曲"],
    ["Sonata 奏鳴曲", "「被奏出的」— canzona 的近親"],
    ["Cori spezzati", "分離合唱團（威尼斯多合唱團風格）"],
    ["Ostinato 固定低音", "不斷反覆的低音型態"],
  ];

  terms.forEach(([term, def], i) => {
    const y = 1.0 + i * 0.43;
    s.addText(term, { x:0.4, y, w:3.4, h:0.4, fontSize:18, bold:true, color:C.brown, fontFace:"Georgia" });
    s.addText(def, { x:3.9, y, w:5.7, h:0.4, fontSize:18, color:C.darkText, fontFace:"Calibri" });
  });
}

// ── SLIDE 27 · NAWM Listening Guide ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);

  s.addText("聆聽指南 NAWM Listening Guide", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:28, bold:true, color:C.gold, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.amber} });

  const nawms = [
    ["NAWM 67", "Susato: Pavane & Galliard (Danserye)", "youtube.com/watch?v=Ln7ea-5dsoo"],
    ["NAWM 68", "Narvaez: Cancion Mille regretz (vihuela)", "youtube.com/watch?v=xKAeB5nDCmA"],
    ["NAWM 69", "Byrd: John come kiss me now (virginal)", "youtube.com/watch?v=DD7luwIuM40"],
    ["NAWM 70", "A. Gabrieli: Ricercar del 12 tuono", "youtube.com/watch?v=DRYBJDFxQo8"],
    ["NAWM 71", "Merulo: Toccata IV (6th Mode)", "search: Merulo Toccata quarta"],
    ["NAWM 72", "G. Gabrieli: Sonata pian' e forte", "youtube.com/watch?v=QXRITlQBitc"],
    ["NAWM 73", "G. Gabrieli: In ecclesiis", "youtube.com/watch?v=Xf8oWn1Hj_8"],
  ];

  nawms.forEach(([num, title, url], i) => {
    const y = 1.05 + i * 0.6;
    s.addShape(pres.ShapeType.rect, { x:0.4, y, w:1.3, h:0.5, fill:{color:C.forest}, rounding:true });
    s.addText(num, { x:0.4, y, w:1.3, h:0.5, fontSize:18, bold:true, color:C.lightText, align:"center", fontFace:"Georgia", valign:"middle" });
    s.addText(title, { x:1.85, y, w:5.0, h:0.28, fontSize:18, bold:true, color:C.gold, fontFace:"Calibri" });
    s.addText(url, { x:1.85, y:y+0.26, w:7.7, h:0.24, fontSize:18, color:C.sand, fontFace:"Calibri" });
  });
}

// ── SLIDE 28 · Instrumental Music Gains Independence ────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);

  s.addText("器樂的獨立 Instrumental Music Gains Independence", {
    x:0.4, y:0.25, w:9.2, h:0.6, fontSize:24, bold:true, color:C.forest, fontFace:"Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x:0.4, y:0.88, w:4, h:0.03, fill:{color:C.sand} });

  const bullets = [
    "16世紀器樂逐漸脫離舞蹈與聲樂，成為獨立藝術\nInstrumental music gradually freed from dance/voice, becoming an art in itself",
    "作曲家開始指定樂器與力度 — 書寫音樂的精確度提升\nComposers began specifying instruments & dynamics — precision in written music",
    "Gabrieli 的作品證明純器樂可與聲樂媲美深度與表現力\nGabrieli proved purely instrumental works could rival vocal music in depth",
    "為巴洛克時期奏鳴曲、協奏曲、交響曲奠定基礎\nLaid foundations for Baroque sonatas, concertos & symphonies",
  ];
  bullets.forEach((txt, i) => {
    s.addText(txt, {
      x:0.6, y:1.1 + i * 1.05, w:8.8, h:0.95, fontSize:18, color:C.darkText, fontFace:"Calibri",
      bullet:true, lineSpacing:24,
    });
  });
}

// ── Generate ────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "Ch12_Instrumental.pptx" })
  .then(() => console.log("Ch12_Instrumental.pptx created"))
  .catch(err => console.error(err));
