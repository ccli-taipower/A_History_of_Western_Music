const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 7: Music and the Renaissance";
pres.author = "A History of Western Music, 10th ed.";

const C = {
  darkBg:   "1C2A3A",
  gold:     "C8A020",
  cream:    "FBF5E6",
  wine:     "7A2830",
  teal:     "2E5A5A",
  darkText: "1C2A3A",
  lightText:"FBF5E6",
  midBlue:  "3A5070",
  sand:     "E8D8A8",
  slate:    "38495C",
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
  s.addText("CHAPTER 7", {
    x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });
  s.addText("MUSIC AND\nTHE RENAISSANCE", {
    x: 0.3, y: 1.5, w: 9.4, h: 2.0, fontSize: 40, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.65, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Humanism · Patronage · New Counterpoint · Printing · Europe 1400–1600", {
    x: 0.4, y: 3.8, w: 9.2, h: 0.4, fontSize: 14, color: C.sand, italic: true, align: "center", fontFace: "Georgia",
  });
  s.addText("Textbook pp. 134–158", {
    x: 0.5, y: 4.8, w: 9, h: 0.3, fontSize: 14, color: C.gold, align: "center", fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);

  s.addText("本章概覽 Chapter Overview", { x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.teal, fontFace: "Georgia", margin: 0 });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const sections = [
    ["■", "Europe 1400–1600", "經濟復甦 · 大航海 · 印刷術 · 宗教改革——現代世界的雛形"],
    ["■", "Humanism 人文主義", "古希臘羅馬復興 · 修辭學 · studia humanitatis · 音樂教育"],
    ["■", "Renaissance in Art", "Donatello · Masaccio · 透視法 · 明暗法 · 自然主義與個人"],
    ["■", "Patronage & Training", "宮廷三部門 · 教堂唱詩學校 · benefices · 國際職涯流動"],
    ["■", "The New Counterpoint", "三六度協和 · 嚴格處理不協和 · 四聲部 · 模仿對位 · 和聲思維"],
    ["■", "Printing & New Ideas", "Petrucci 1501 · 古希臘理論復興 · 調律 · 詞曲關係 · 表情"],
  ];
  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.teal }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 14, color: C.midBlue, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 3 · Europe 1400–1600 ───────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("歐洲 1400–1600：新世界的誕生", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Europe from 1400 to 1600 · Political, Economic, and Technological Change", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Left: Political/Religious
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.1, fill: { color: "2A3C50" }, rounding: true });
  s.addText("■ 政治與宗教 Political & Religious", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 1417 大分裂結束——教宗回歸羅馬\n• 1453 百年戰爭結束 · 君士坦丁堡陷落——拜占庭帝國滅亡\n• 1517 Martin Luther 宗教改革開始\n• 新教三大宗派：路德派、喀爾文派、聖公會\n• Catholic Reformation 反宗教改革——Palestrina 成為 16 世紀對位法典範\n• 鄂圖曼土耳其繼續征服巴爾幹、匈牙利\n• 百年以上的宗教戰爭塑造歐洲版圖", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Right: Economic/Technological
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.1, fill: { color: "2A3C50" }, rounding: true });
  s.addText("■ 經濟與技術 Economic & Technological", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• ca. 1400 歐洲經濟復甦 · 區域專業化與長距離貿易\n• 中產階級（商人、工匠、律師）興起\n• 1450 Gutenberg 活字印刷術\n• 1487 葡萄牙繞過非洲南端\n• 1492 Columbus 抵達西印度群島\n• 歐洲殖民擴張至美洲、非洲、亞洲\n• 科學、造船、火炮、航海技術突破\n• 義大利城邦的宮廷贊助成為藝術發展的沃土", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 4 · Humanism ───────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("人文主義 Humanism", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("studia humanitatis · The Revival of Ancient Learning", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.12, w: 5, h: 0.04, fill: { color: C.wine } });

  // Key points
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.1, fill: { color: "F3E5C5" }, rounding: true });
  s.addText("■ 核心理念 Core Ideas", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("• studia humanitatis：文法、修辭、詩學、歷史、道德哲學\n• 反對經院哲學的邏輯與權威——轉而重視個人道德、公民生活\n• 古典希臘與羅馬文獻的廣泛取得（拜占庭學者西遷）\n• 對人類尊嚴、理性、感官經驗的信心\n• 與基督教融合——而非取代\n• 重塑大學課程，影響藝術家、音樂家、政治家", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Effects on music
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.1, fill: { color: "F3E5C5" }, rounding: true });
  s.addText("■ 對音樂的影響 Effects on Music", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("• 古代音樂論著的再發現（Aristides Quintilianus、Ptolemy、Cleonides）\n• Aristotle、Plato、Quintilian 視音樂為公民教育必備\n• 修辭學思維進入作曲——結構、說服、清晰度\n• 對詞曲關係的新重視——自然音節與意義\n• 音樂被視為「有表情的語言」\n• 音樂作為紳士仕女社交禮儀的一部分（彈琵琶而非吹笛）\n• Franchino Gaffurio（1451–1522）融合古代理論與當代實踐", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 5 · Renaissance in Art ─────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Renaissance 藝術與音樂的平行", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Sculpture, Painting, Architecture · Parallels with Music", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const artRows = [
    ["Donatello 《David》(1440s)", "羅馬時代以來第一尊獨立裸體雕像——古典復興傳達宗教主題", "Composers 使用對比音色、力度支持文字情感"],
    ["Masaccio 《Holy Trinity》(1425–28)", "早期透視法運用——單一消失點、深度錯覺、個人捐贈者像", "樂曲調式清晰——所有樂句終止於主要音（final、reciting tone）"],
    ["Perspective 透視法", "平行線匯聚於單一消失點——二維平面創造三維幻覺", "音樂的「方向感」——圍繞單一參考點的和聲結構"],
    ["Chiaroscuro 明暗法", "自然的光影處理——物體立體感與空間感", "高低音域、厚薄織度的對比——光與影的音樂版本"],
    ["Clarity & Classical Models", "柱式、拱形、對稱——古代建築典範", "清晰終止式、規則對位——古典秩序與美感"],
    ["Interest in Individuals", "肖像畫興起——捐贈者、藝術家、個人性格", "作曲家成為名人——個人風格、作品集、被研究與模仿"],
  ];

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 0.35, fill: { color: C.gold } });
  s.addText("藝術元素", { x: 0.4, y: 1.32, w: 2.9, h: 0.4, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", align: "center" });
  s.addText("藝術中的表現 Artistic Expression", { x: 3.3, y: 1.32, w: 3.4, h: 0.4, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", align: "center" });
  s.addText("音樂中的對應 Musical Parallel", { x: 6.7, y: 1.32, w: 3.0, h: 0.4, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", align: "center" });

  artRows.forEach(([k, a, m], i) => {
    const y = 1.68 + i * 0.63;
    const bgColor = i % 2 === 0 ? "2A3C50" : "35485E";
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.6, fill: { color: bgColor } });
    s.addText(k, { x: 0.4, y: y + 0.04, w: 2.9, h: 0.52, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(a, { x: 3.3, y: y + 0.04, w: 3.4, h: 0.52, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" });
    s.addText(m, { x: 6.7, y: y + 0.04, w: 3.0, h: 0.52, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 6 · Patronage and Training ─────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);

  s.addText("贊助與音樂家訓練", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addText("Patronage and the Training of Musicians", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.12, w: 5, h: 0.04, fill: { color: C.teal } });

  // Three court divisions
  const divs = [
    ["■", "Chapel 禮拜堂", "神職人員與歌手\n負責宗教儀式音樂\n展現君主虔誠與榮耀"],
    ["■", "Chamber 宮內", "僕從與樂手服侍家族\nbas 軟聲樂器（琵琶、笛）\n展現文化修養"],
    ["■", "Public Court 公務", "公共儀式與節慶\nhaut 硬聲樂器（小號、shawm）\n展現政治軍事權力"],
  ];
  divs.forEach(([icon, title, desc], i) => {
    const x = 0.3 + i * 3.2;
    s.addShape(pres.ShapeType.rect, { x, y: 1.28, w: 3.0, h: 1.8, fill: { color: "E0EEEC" }, line: { color: C.teal, width: 1.5 }, rounding: true });
    s.addText(icon, { x: x + 0.1, y: 1.32, w: 0.6, h: 0.4, fontSize: 18 });
    s.addText(title, { x: x + 0.7, y: 1.35, w: 2.25, h: 0.3, fontSize: 14, bold: true, color: C.teal, fontFace: "Georgia", margin: 0 });
    s.addText(desc, { x: x + 0.12, y: 1.75, w: 2.78, h: 1.3, fontSize: 14, color: C.darkText, fontFace: "Calibri", valign: "top", paraSpaceAfter: 1 });
  });

  // Training
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.2, w: 9.4, h: 2.1, fill: { color: "E0EEEC" }, rounding: true });
  s.addText("■ 音樂家的養成 Training and Mobility", { x: 0.5, y: 3.3, w: 9.0, h: 0.35, fontSize: 14, bold: true, color: C.teal, fontFace: "Georgia", margin: 0 });
  s.addText("• 歌手：唱詩學校訓練唱歌、樂理、對位、作曲；Cambrai、Bruges、Antwerp 為重鎮\n• 器樂手：家族傳承與工會師徒制；多數不識譜\n• Benefices（教會俸祿）：教宗授予神職人員——作曲家多為神職\n• 國際流動：法蘭德斯人遍布義大利；Medici、Este、Sforza 競相延攬\n• Ciconia（ca. 1370–1412）最早赴義大利的北方作曲家", {
    x: 0.5, y: 3.65, w: 9.0, h: 1.6, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 7 · The New Counterpoint ───────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("The New Counterpoint 新對位法", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("The Core of the International Style · ca. 1420 onward", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // 4 key principles
  const principles = [
    ["■", "偏好協和音 Consonance Preferred", "三度、六度與完全五度、八度皆為協和——三六度不再是不協和"],
    ["■", "嚴格處理不協和 Strict Dissonance Control", "只允許經過音、鄰音（弱拍）與留音（終止式）——平行五度八度禁止"],
    ["■", "四聲部織度 Four-Voice Texture", "15 世紀下半取代三聲部——低音線加入 tenor 之下成為現代織度基礎"],
    ["⇄", "聲部平等 Equality of Voices", "所有聲部同時構思（和聲思維）——不再是層層疊加 cantus/tenor 框架"],
  ];
  principles.forEach(([icon, title, desc], i) => {
    const y = 1.32 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.68, fill: { color: "2A3C50" }, rounding: true });
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 0.1, h: 0.68, fill: { color: C.gold } });
    s.addText(icon, { x: 0.5, y: y + 0.12, w: 0.55, h: 0.5, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
    s.addText(title, { x: 1.1, y: y + 0.05, w: 8.5, h: 0.32, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(desc, { x: 1.1, y: y + 0.35, w: 8.5, h: 0.3, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" });
  });

  // Tinctoris quote
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 4.45, w: 9.4, h: 0.78, fill: { color: "3E2030" }, rounding: true });
  s.addText("Johannes Tinctoris (ca. 1435–1511) · Liber de arte contrapuncti (1477)", { x: 0.45, y: 4.5, w: 9.1, h: 0.3, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("「40 年前的作品沒有一首值得演奏……只有 Dunstable、Binchois、Du Fay 與 Ockeghem 這一輩的音樂才值得模仿。」", { x: 0.45, y: 4.78, w: 9.1, h: 0.42, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" });
}

// ── SLIDE 8 · New Textures ───────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("兩種新織度 Two New Textures", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("Imitative Counterpoint & Homophony · Dominant in the 16th Century", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.12, w: 5, h: 0.04, fill: { color: C.wine } });

  // Imitation
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.1, fill: { color: "F3E5C5" }, rounding: true });
  s.addText("■ Imitative Counterpoint 模仿對位", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("定義：一聲部提出動機或樂句，其他聲部在不同音高（通常五度、四度或八度）回應模仿", {
    x: 0.5, y: 1.88, w: 4.35, h: 0.54, fontSize: 14, color: C.darkText, fontFace: "Calibri", valign: "top",
  });
  s.addText("• 所有聲部平等參與——無 cantus/tenor 主從之分\n• 每個樂段通常有一個音樂動機（motto）\n• 文字宣告清晰——所有聲部同節奏唱出相同字句\n• 容許結構複雜而仍然清楚\n• 成為 16 世紀經文歌、彌撒、器樂作品的核心技法\n• 集大成者：Josquin、Lassus、Palestrina", {
    x: 0.5, y: 2.5, w: 4.35, h: 2.7, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Homophony
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.1, fill: { color: "F3E5C5" }, rounding: true });
  s.addText("■ Homophony 主調織度", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("定義：所有聲部以相同（或近乎相同）節奏進行——低聲部以協和音伴奏 cantus", {
    x: 5.3, y: 1.88, w: 4.35, h: 0.54, fontSize: 14, color: C.darkText, fontFace: "Calibri", valign: "top",
  });
  s.addText("• 結構簡單直接——容易演唱、聆聽\n• 文字宣告極為清晰——每字每音一目瞭然\n• 16 世紀通俗歌曲（frottola、villancico、chanson）的主要風格\n• 與模仿對位常交替使用於同一作品中——創造對比\n• 展現 Renaissance 對詞曲關係的新重視\n• 後來發展為主旋律伴奏（monody）——通往 Baroque", {
    x: 5.3, y: 2.5, w: 4.35, h: 2.7, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 9 · Tuning and Temperament ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("調律與律制 Tuning and Temperament", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("The Challenge of Thirds and Sixths", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const tunings = [
    ["Pythagorean Intonation", "中世紀傳統", "ratios 2:1, 3:2, 4:3 · 純四五八度 · 三度（81:64）粗糙不協和 · 適用於僅需完全音程的中世紀音樂"],
    ["Just Intonation 純律", "1482 Ramis de Pareia", "Walter Odington 早在 1300 即觀察 5:4、6:5 為協和三度 · 純三六度 · 缺點：部分三五度走音，G■ 與 A■ 不同音"],
    ["Mean-Tone Temperament 中庸律", "16 世紀主流鍵盤律制", "五度略縮——多數大三度純或近純 · 16 世紀至 19 世紀中葉鍵盤樂器的標準律制"],
    ["Equal Temperament 十二平均律", "16 世紀晚期首見於理論", "所有半音相等 · 任何調皆可用 · 僅八度純 · 三六度差 1/7 半音 · 19 世紀後半才廣泛採用"],
  ];
  tunings.forEach(([name, when, desc], i) => {
    const y = 1.3 + i * 0.9;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.82, fill: { color: i % 2 === 0 ? "2A3C50" : "35485E" }, rounding: true });
    s.addText(name, { x: 0.5, y: y + 0.05, w: 4.0, h: 0.32, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(when, { x: 4.5, y: y + 0.08, w: 5.0, h: 0.28, fontSize: 14, color: C.sand, italic: true, fontFace: "Georgia" });
    s.addText(desc, { x: 0.5, y: y + 0.38, w: 9.0, h: 0.42, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 10 · Words and Music ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);

  s.addText("Words and Music 詞曲關係", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addText("Text Declamation, Emotion, and the Influence of Ancient Writings", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.12, w: 5, h: 0.04, fill: { color: C.teal } });

  // Four topics grid
  const topics = [
    ["■ Text Declamation 文字宣告", "formes fixes 逐漸消失 · 依文字重音與節奏作曲 · 15 世紀末期自然宣告成為準則 · 模仿對位與主調織度讓所有聲部同時宣告"],
    ["■ Emotional Expression 情感表達", "中世紀音樂少見表情意識 · 15–16 世紀起作曲家使用特定音程、和聲、旋律線條傳達文字情緒 · Cicero、Quintilian 修辭學支持"],
    ["■ Expressive Power of Modes", "Plato、Aristotle 認為各 harmoniai 具不同 ethos · 作曲家據古代權威選擇調式傳達情感 · 後來延伸至大小調、各調性格的觀念"],
    ["■ Chromaticism 半音主義", "古希臘 chromatic genus 啟發 · 16 世紀中葉作曲家開始使用直接半音進行（如 B→B■）作為表情手法 · Gregorian 以來首見"],
  ];
  topics.forEach(([title, desc], i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.3 + col * 4.75;
    const y = 1.3 + row * 2.0;
    s.addShape(pres.ShapeType.rect, { x, y, w: 4.6, h: 1.85, fill: { color: "E0EEEC" }, line: { color: C.teal, width: 1 }, rounding: true });
    s.addText(title, { x: x + 0.15, y: y + 0.1, w: 4.35, h: 0.35, fontSize: 14, bold: true, color: C.teal, fontFace: "Georgia", margin: 0 });
    s.addText(desc, { x: x + 0.15, y: y + 0.5, w: 4.35, h: 1.35, fontSize: 14, color: C.darkText, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 11 · Music Printing ────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("音樂印刷術 Music Printing", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Petrucci · Attaingnant · The Technological Revolution", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Timeline of printing developments
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 1.85, fill: { color: "2A3C50" }, rounding: true });
  s.addText("■ 印刷技術的發展 Development of Printing Technology", { x: 0.45, y: 1.38, w: 9.1, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• ca. 1450：Johann Gutenberg 在歐洲完善活字印刷術\n• 1470s：活字首次用於音樂——禮拜用書的 chant 記譜\n• 1501：Ottaviano Petrucci（1466–1539）在威尼斯出版 Harmonice musices odhecaton A——第一本完全用活字印刷的複音樂譜（96 首作品）\n  三次壓印：譜線、文字、音符與花體字母分三次印刷——費時但精美\n• ca. 1520：John Rastell 在倫敦首創「單次壓印」\n• 1528：Pierre Attaingnant（ca. 1494–1551/52）在巴黎大規模使用單次壓印——譜線斷續但便宜實用——成為後續標準", {
    x: 0.5, y: 1.88, w: 9.0, h: 1.24, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Impact
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.25, w: 9.4, h: 2.0, fill: { color: "2A3C50" }, rounding: true });
  s.addText("■ 印刷術的衝擊 Impact of Music Printing", { x: 0.45, y: 3.33, w: 9.1, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 樂譜價格大幅降低——業餘愛好者可負擔\n• 音樂識讀能力普及——中產階級家庭成為音樂場域\n• 作曲家名聲跨越地域——音樂「作品」概念成形\n• 新類型蓬勃：madrigal、chanson、villancico、lute song\n• 器樂譜首次大量問世：toccata、ricercare、canzona\n• 印刷中心：Venice、Paris、Rome、Nuremberg、London\n• 宗教改革借助印刷快速傳播新音樂類型", {
    x: 0.5, y: 3.83, w: 9.0, h: 1.37, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 12 · Key Theorists & Figures ───────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("重要理論家與人物 Key Theorists & Figures", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 0.82, w: 5, h: 0.03, fill: { color: C.wine } });

  const figures = [
    ["Johannes Tinctoris", "ca. 1435–1511", "Liber de arte contrapuncti (1477) · 12 部論著 · 新對位法典範"],
    ["Franchino Gaffurio", "1451–1522", "Theorica musice (1492) · 融合古希臘理論與當代實踐"],
    ["Pietro Aaron", "ca. 1480–ca. 1550", "Toscanello in musica (1523) · 首位以義大利文寫作的理論家"],
    ["Heinrich Glareanus", "1488–1563", "Dodecachordon (1547) · 新增 Aeolian、Ionian 四個調式"],
    ["Gioseffo Zarlino", "1517–1590", "Le istitutioni harmoniche (1558) · 對位法與和聲理論的集大成者"],
    ["Johann Gutenberg", "ca. 1400–1468", "活字印刷術的完善者——改變知識傳播"],
    ["Ottaviano Petrucci", "1466–1539", "1501 威尼斯首本複音樂印刷集——Harmonice musices odhecaton"],
    ["Pierre Attaingnant", "ca. 1494–1551/52", "單次壓印法巴黎推廣者——使樂譜便宜普及"],
  ];
  figures.forEach(([name, dates, desc], i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.3 + col * 4.75;
    const y = 1.05 + row * 1.07;
    s.addShape(pres.ShapeType.rect, { x, y, w: 4.6, h: 0.97, fill: { color: "F3E5C5" }, line: { color: C.wine, width: 1 }, rounding: true });
    s.addText(name, { x: x + 0.15, y: y + 0.05, w: 3.4, h: 0.3, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
    s.addText(dates, { x: x + 0.15, y: y + 0.33, w: 4.3, h: 0.22, fontSize: 14, color: C.slate, fontFace: "Georgia" });
    s.addText(desc, { x: x + 0.15, y: y + 0.53, w: 4.3, h: 0.42, fontSize: 14, color: C.darkText, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 13 · Timeline ──────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("年代大事記 Timeline", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Europe and Music 1400–1600", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const events = [
    ["ca. 1400", "歐洲經濟開始復甦"],
    ["1417", "大分裂結束——教宗回歸羅馬"],
    ["ca. 1425–28", "Masaccio《聖三位一體》"],
    ["ca. 1440s", "Donatello《大衛》青銅像"],
    ["ca. 1450", "Gutenberg 完善活字印刷"],
    ["1453", "百年戰爭結束 · 君士坦丁堡陷落"],
    ["1477", "Tinctoris · Liber de arte contrapuncti"],
    ["1482", "Ramis de Pareia 提出純律"],
    ["1487", "葡萄牙繞過好望角"],
    ["1492", "Columbus 抵達西印度群島 · Gaffurio Theorica"],
    ["1495", "Leonardo da Vinci《最後的晚餐》"],
    ["1501", "Petrucci 出版 Odhecaton"],
    ["1517", "Luther 宗教改革開始"],
    ["1523", "Pietro Aaron · Toscanello in musica"],
    ["1528", "Attaingnant 單次壓印法"],
    ["1532", "Machiavelli《君王論》出版"],
    ["1547", "Glareanus · Dodecachordon"],
    ["1558", "Zarlino · Le istitutioni harmoniche"],
    ["1594", "Shakespeare《羅密歐與茱麗葉》"],
  ];
  events.forEach(([date, ev], i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.3 + col * 4.75;
    const y = 1.3 + row * 0.4;
    s.addText(date, { x: x + 0.1, y, w: 1.5, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(ev, { x: x + 1.6, y, w: 3.1, h: 0.35, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 14 · Legacy / Summary ──────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);

  s.addText("Renaissance 的遺產 The Legacy", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addText("Chapter Summary", { x: 0.4, y: 0.75, w: 9.2, h: 0.35, fontSize: 14, italic: true, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.12, w: 5, h: 0.04, fill: { color: C.teal } });

  const summary = [
    ["■", "一個統一的時代", "15–16 世紀以新對位法為基石，構成音樂史上的「Renaissance」時代（ca. 1400–1600）"],
    ["■", "人文主義的力量", "古典學術的復興影響音樂美學——詞曲關係、修辭、表情與清晰度成為新準則"],
    ["■", "新對位法的建立", "三六度協和、嚴格不協和處理、聲部平等——奠定此後數百年和聲與對位的基礎"],
    ["■", "印刷術的革命", "音樂從精英走向業餘——新類型、國家風格、器樂音樂與宗教改革音樂因此蓬勃"],
    ["■", "新舊律制並存", "純律、中庸律、平均律的發展反映音樂實踐對理論的主導地位"],
    ["■", "音樂成為「作品」", "固定記譜 · 作曲家個人身份 · 被收藏、研究、模仿——現代音樂觀念的起點"],
  ];
  summary.forEach(([icon, title, desc], i) => {
    const y = 1.3 + i * 0.67;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.55, h: 0.55, fill: { color: C.teal }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.55, h: 0.5, fontSize: 18, align: "center" });
    s.addText(title, { x: 1.1, y, w: 8.5, h: 0.28, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(desc, { x: 1.1, y: y + 0.26, w: 8.5, h: 0.35, fontSize: 14, color: C.slate, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 15 · Key Terms & Further Reading ───────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("重要術語與延伸閱讀", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Key Terms · Further Reading", { x: 0.4, y: 0.75, w: 9.2, h: 0.35, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.12, w: 5, h: 0.04, fill: { color: C.gold } });

  // Key terms
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 2.3, fill: { color: "2A3C50" }, rounding: true });
  s.addText("■ 重要術語 Key Terms", { x: 0.45, y: 1.36, w: 9.1, h: 0.3, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });

  const termsLeft = [
    "Renaissance · studia humanitatis",
    "Humanism · Rhetoric",
    "International style",
    "New counterpoint",
    "Imitative counterpoint · Homophony",
    "Cantus-tenor framework",
  ];
  const termsRight = [
    "Pythagorean · Just · Mean-tone · Equal temp.",
    "Chromaticism (chromatic genus)",
    "Benefices · Choir schools",
    "Partbooks · Single/triple impression",
    "Chapel · Chamber · Public court",
    "Haut and bas instruments",
  ];
  termsLeft.forEach((t, i) => {
    s.addText("• " + t, { x: 0.5, y: 1.7 + i * 0.3, w: 4.1, h: 0.28, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" });
  });
  termsRight.forEach((t, i) => {
    s.addText("• " + t, { x: 4.7, y: 1.7 + i * 0.3, w: 4.95, h: 0.28, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" });
  });

  // Listening (new)
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.7, w: 9.4, h: 1.0, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ 聆聽 Listen (文藝復興代表曲)", { x: 0.45, y: 3.76, w: 9.1, h: 0.3, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText(
    "Tinctoris — Missa L'homme armé: Kyrie  youtu.be/VwbnG2XD540\n" +
    "Josquin — Ave Maria...virgo serena (Tallis Scholars)  youtu.be/scQ5YBRpwNg\n" +
    "Palestrina — Sicut cervus (Cambridge Singers)  youtu.be/0yd5EE0hAB8",
    { x: 0.5, y: 4.04, w: 9.1, h: 0.65, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" }
  );

  // Further reading
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 4.75, w: 9.4, h: 0.85, fill: { color: "3E2030" }, rounding: true });
  s.addText("■ 延伸閱讀 Further Reading", { x: 0.45, y: 4.78, w: 9.1, h: 0.25, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("Leeman Perkins, Music in the Age of the Renaissance (1999) · Allan Atlas, Renaissance Music (1998)\nReinhard Strohm, The Rise of European Music 1380–1500 (1993) · Tinctoris, The Art of Counterpoint (trans. Seay)", {
    x: 0.5, y: 5.03, w: 9.1, h: 0.55, fontSize: 14, color: C.sand, fontFace: "Calibri", italic: true, valign: "top",
  });
}

pres.writeFile({ fileName: "Ch07_Renaissance.pptx" }).then(() => {
  console.log("■ Ch07_Renaissance.pptx created successfully");
});
