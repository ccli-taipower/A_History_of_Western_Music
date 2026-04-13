const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 6: New Developments in the Fourteenth Century";
pres.author = "A History of Western Music, 10th ed.";

const C = {
  darkBg:   "2C1810",
  gold:     "C8A020",
  cream:    "FBF5E6",
  wine:     "7A2830",
  rust:     "A84030",
  darkText: "2C1810",
  lightText:"FBF5E6",
  midBrown: "5C3A28",
  sand:     "E8D8A8",
  slate:    "4A3828",
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
  s.addText("CHAPTER 6", {
    x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });
  s.addText("NEW DEVELOPMENTS\nIN THE FOURTEENTH CENTURY", {
    x: 0.3, y: 1.4, w: 9.4, h: 2.0, fontSize: 34, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.55, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Ars Nova · Machaut · Ars Subtilior · Trecento · Landini", {
    x: 0.4, y: 3.7, w: 9.2, h: 0.4, fontSize: 13, color: C.sand, italic: true, align: "center", fontFace: "Georgia",
  });
  s.addText("Textbook pp. 106–133", {
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
    ["⚔", "14 世紀的歐洲 Europe in the 14th Century", "黑死病、百年戰爭、亞維農教廷、教會大分裂——動盪中的創造力"],
    ["📜", "新藝術 The Ars Nova in France", "Philippe de Vitry、二等分拍、minim、記譜革命"],
    ["👑", "Guillaume de Machaut", "詩人兼作曲家 · 140 作品 · Messe de Nostre Dame · formes fixes"],
    ["🎭", "Ars Subtilior", "巴比倫式的節奏與記譜複雜——亞維農與義大利北部的極致風格"],
    ["🇮🇹", "義大利 Trecento 音樂", "Madrigal · Caccia · Ballata · Francesco Landini · Squarcialupi Codex"],
    ["🎼", "演出實務與遺產 Performance & Legacy", "Musica ficta · 高低樂器 · Ars Nova 記譜是現代記譜的直系祖先"],
  ];
  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.wine }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 11, color: C.midBrown, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 3 · Europe in the 14th Century ─────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("14 世紀的歐洲", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Europe in the Fourteenth Century · Disruption and Creativity", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Crisis box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 1.9, fill: { color: "3A2015" }, rounding: true });
  s.addText("💀 災難的世紀 A Century of Calamity", { x: 0.45, y: 1.36, w: 9.1, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 1315–22 年大饑荒：西北歐洪水造成 1/10 人口喪生\n• 1347–50 年黑死病（腺鼠疫+肺鼠疫）橫掃歐洲——三分之一人口死亡\n• 1337–1453 英法百年戰爭；各地農民與城市暴動\n• 教會危機：1309–77 教宗移駐亞維農（「巴比倫之囚」）· 1378–1417 教會大分裂\n• William of Ockham（ca. 1285–1349）主張知識應基於感官經驗——奠定現代科學方法", {
    x: 0.5, y: 1.7, w: 9.0, h: 1.45, fontSize: 10, color: C.sand, fontFace: "Calibri",
  });

  // Culture box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.3, w: 9.4, h: 1.95, fill: { color: "3A2015" }, rounding: true });
  s.addText("🎨 文化與技術的躍進 Cultural & Technological Leaps", { x: 0.45, y: 3.36, w: 9.1, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Giotto（ca. 1266–1337）畫作呈現自然主義與空間深度——繪畫的 Renaissance 先聲\n• Dante《神曲》(1307) · Boccaccio《十日談》(1348–53) · Chaucer《坎特伯雷故事集》(ca. 1387–1400)\n• 新技術：眼鏡、磁羅盤、機械鐘——時間從教堂鐘聲轉為普世精確度量\n• Roman de Fauvel (ca. 1317)：諷刺教會政治的寓言詩，手稿含 169 首音樂作品\n  Allegorical poem satirizing church/politics; manuscript contains 169 musical pieces", {
    x: 0.5, y: 3.7, w: 9.0, h: 1.5, fontSize: 10, color: C.sand, fontFace: "Calibri",
  });
}

// ── SLIDE 4 · The Ars Nova in France ─────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust); bottomBar(s, C.rust);

  s.addText("新藝術 The Ars Nova in France", { x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 24, bold: true, color: C.rust, fontFace: "Georgia" });
  s.addText("Philippe de Vitry (1291–1361) · A Revolution in Rhythm", { x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 12, color: C.midBrown, fontFace: "Calibri" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.04, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Vitry / treatise
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.wine }, rounding: true });
  s.addText("👤 Philippe de Vitry", { x: 0.45, y: 1.28, w: 4.3, h: 0.35, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.55, y: 1.66, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("• 1291–1361 · 法國作曲家、詩人、教會參事\n  隨後成為 Meaux 主教\n\n• 被稱為「新藝術的發明者」\n• 約 1320 年的 Ars nova 論文——\n  以其教學為基礎\n\n• 代表作：\n  Cum statua / Hugo, Hugo / Magister\n  invidie（NAWM 24）· 三聲部\n\n• Ars Nova = 新藝術/新方法\n  1310 年代始 · 持續到 1370 年代\n\n• 反對者：Jacobus de Ispania\n  Speculum musicae (ca. 1330)\n  擁護舊的「ars antiqua」", {
    x: 0.5, y: 1.78, w: 4.2, h: 3.4, fontSize: 10, color: C.cream, fontFace: "Calibri",
  });

  // Notation innovations
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.midBrown }, rounding: true });
  s.addText("📐 記譜革命 Notation Revolution", { x: 5.25, y: 1.28, w: 4.3, h: 0.35, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 5.35, y: 1.66, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("① 二等分拍子（imperfect）與三等分拍子\n   equally valid——首次可寫下 duple meter\n\n② 新的小音符 minim（「最小」）\n   將 semibreve 再細分\n\n③ 時值結構三層：\n   mode（long 分拍）· time（breve 分拍）\n   · prolation（semibreve 分拍）\n\n④ Mensuration signs（ca. 1340 Jehan\n   des Murs）——現代拍號的祖先\n\n⑤ 首次可寫下 syncopation 切分音\n\n❖ Jehan des Murs：\n  「whatever can be sung can be written」\n  只要唱得出來就寫得出來\n\n❖ Ars Nova 記譜是現代記譜的直系祖先", {
    x: 5.3, y: 1.78, w: 4.2, h: 3.4, fontSize: 9, color: C.cream, fontFace: "Calibri",
  });
}

// ── SLIDE 5 · Isorhythm ──────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("等節奏 Isorhythm", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("The Structural Device of the Ars Nova Motet", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Talea
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 2.0, fill: { color: "3A2015" }, rounding: true });
  s.addText("🎵 Talea（節奏單元）", { x: 0.45, y: 1.36, w: 4.3, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("「talea」= 拉丁「cutting」\n• tenor 聲部中反覆出現的節奏模式\n• 較 13 世紀 clausula 更長更複雜\n• 形成整首樂曲的節奏骨架", {
    x: 0.5, y: 1.7, w: 4.2, h: 1.5, fontSize: 10, color: C.sand, fontFace: "Calibri",
  });

  // Color
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 2.0, fill: { color: "3A2015" }, rounding: true });
  s.addText("🎨 Color（旋律單元）", { x: 5.25, y: 1.36, w: 4.3, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("「color」= 反覆出現的旋律段\n• tenor 的旋律循環\n• 與 talea 長度可以不同\n• 常見：color 延伸跨越 2、3 個 taleae\n• 使兩者的結尾錯開——層層結構", {
    x: 5.3, y: 1.7, w: 4.2, h: 1.5, fontSize: 10, color: C.sand, fontFace: "Calibri",
  });

  // Hocket + characteristics
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.4, w: 9.4, h: 1.85, fill: { color: "3A2015" }, rounding: true });
  s.addText("⚡ Hocket（打嗝效果）與 Ars Nova 特徵", { x: 0.45, y: 3.46, w: 9.1, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Hocket 來自法文 hoquet「打嗝」——兩聲部快速交替，一聲部休息時另一聲部發聲\n  Two voices alternate rapidly, one resting while the other sings\n• 常見於 14 世紀 isorhythmic 作品中 · 有時整首為 hocket · 可為人聲或器樂\n• Ars Nova 的和聲：三度與六度（imperfect consonances）使用更頻繁——聽起來更甜美\n• 但仍可見平行五度八度——與 15–16 世紀對位有別\n• 這些 motets 寫給具文學與音樂素養的聽眾——詞樂結構交織的樂趣", {
    x: 0.5, y: 3.8, w: 9.0, h: 1.4, fontSize: 9.5, color: C.sand, fontFace: "Calibri",
  });
}

// ── SLIDE 6 · Guillaume de Machaut ───────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("Guillaume de Machaut (ca. 1300–1377)", { x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 22, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("The Most Important Composer and Poet of the Ars Nova", { x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 12, color: C.midBrown, fontFace: "Calibri" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.04, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Biography
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.wine }, rounding: true });
  s.addText("📖 生平 Biography", { x: 0.45, y: 1.28, w: 4.3, h: 0.35, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.55, y: 1.66, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("• 出生於法國香檳省（Champagne）\n• 受教育為教士\n• ca. 1323 任波希米亞國王 John of\n  Luxembourg 的秘書 · 隨其東征西伐\n• 1340 起為 Reims 主教座堂參事——\n  有充分時間創作\n• 贊助人：Bonne、Navarre 王、法國王、\n  Berry 與 Burgundy 公爵\n\n❖ 劃時代意義：\n• 是第一位主動編纂自己全集的作曲家\n• 生前監督多部華麗插圖手稿的製作\n• 與愛人 Peronne 通信討論創作方法\n  Le livre du voir dit (1363–65)\n• 展現作曲家的自我意識——\n  這在 19 世紀前極為罕見", {
    x: 0.5, y: 1.78, w: 4.2, h: 3.4, fontSize: 9, color: C.cream, fontFace: "Calibri",
  });

  // Works
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.midBrown }, rounding: true });
  s.addText("🎼 主要作品 Major Works", { x: 5.25, y: 1.28, w: 4.3, h: 0.35, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 5.35, y: 1.66, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("❖ 宗教音樂\n• La Messe de Nostre Dame\n  單一作曲家完成的首部完整\n  複音常規彌撒（4 聲部）\n• 23 motets（20 isorhythmic）\n• Hoquetus David\n\n❖ 世俗音樂（formes fixes）\n• 42 ballades（1 首單聲部）\n• 22 rondeaux\n• 33 virelais（25 首單聲部）\n• 19 lais（15 首單聲部）\n• 1 complainte · 1 chanson royale\n\n❖ 詩作\n• Remede de Fortune\n• Le livre du voir dit\n• 其他 280+ 首詩\n\n❖ 影響：Chaucer 受其啟發；\n  美國說唱組 Panda Bear 採樣其 rondeau", {
    x: 5.3, y: 1.78, w: 4.2, h: 3.4, fontSize: 9, color: C.cream, fontFace: "Calibri",
  });
}

// ── SLIDE 7 · Machaut's Messe de Nostre Dame + Formes Fixes ──────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Messe de Nostre Dame 與定型歌", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("The First Polyphonic Mass Cycle · The Three Formes Fixes", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Mass
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 1.55, fill: { color: "3A2015" }, rounding: true });
  s.addText("⛪ La Messe de Nostre Dame (ca. 1360s)", { x: 0.45, y: 1.36, w: 9.1, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 首部由單一作曲家設想為整體的常規彌撒（Mass Ordinary）複音曲\n• 4 聲部 · 加入 contratenor 與 tenor 同音域（上下互換）· 為 Reims 主教座堂聖母禮拜所作\n• Kyrie / Sanctus / Agnus Dei / Ite missa est — isorhythmic · 帶 cantus firmus\n• Gloria / Credo — discant / conductus 風格（syllabic、同節奏聲部）· Gloria 的 \"Jesu Christe\" 以持續和弦突顯\n• 前三樂章以 D 為調心，後三樂章以 F 為調心 —— 風格統一", {
    x: 0.5, y: 1.7, w: 9.0, h: 1.15, fontSize: 9.5, color: C.sand, fontFace: "Calibri",
  });

  // Formes fixes
  const forms = [
    ["🎶 Ballade", "aabC", "最嚴肅 · 三段詩節 · 每段結以相同副歌 C", "哲學、歷史、愛情主題"],
    ["🔄 Rondeau", "ABaAabAB", "愛情主題 · 單一詩節 · 副歌分切穿插", "Machaut: Rose, liz, printemps, verdure"],
    ["💃 Virelai", "AbbaA...", "自然景物 + 愛情 · chanson baladée 「跳舞的歌」", "Machaut: Douce dame jolie"],
  ];
  forms.forEach(([name, structure, desc, ex], i) => {
    const y = 3.0 + i * 0.77;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.68, fill: { color: "3A2015" }, rounding: true });
    s.addText(name, { x: 0.4, y: y + 0.05, w: 2.2, h: 0.3, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(structure, { x: 0.4, y: y + 0.33, w: 2.2, h: 0.3, fontSize: 11, color: C.sand, italic: true, fontFace: "Calibri", margin: 0 });
    s.addShape(pres.ShapeType.rect, { x: 2.65, y: y + 0.1, w: 0.025, h: 0.48, fill: { color: C.gold } });
    s.addText(desc, { x: 2.75, y: y + 0.05, w: 4.0, h: 0.58, fontSize: 9.5, color: C.cream, fontFace: "Calibri", margin: 0 });
    s.addText(ex, { x: 6.85, y: y + 0.05, w: 2.75, h: 0.58, fontSize: 9, color: C.sand, fontFace: "Calibri", italic: true, margin: 0 });
  });
}

// ── SLIDE 8 · The Ars Subtilior ──────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust); bottomBar(s, C.rust);

  s.addText("極致藝術 The Ars Subtilior", { x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 24, bold: true, color: C.rust, fontFace: "Georgia" });
  s.addText("\"The More Subtle Manner\" · Late 14th Century · Avignon & N. Italy", { x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 12, color: C.midBrown, fontFace: "Calibri" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.04, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Context
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.18, w: 9.4, h: 1.85, fill: { color: C.wine }, rounding: true });
  s.addText("🎭 背景 Context", { x: 0.45, y: 1.26, w: 9.1, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 名稱由音樂學家 Ursula Günther 所創——源自 Philipoctus de Caserta 論文中的\n  「artem magis subtiliter」（更細緻的方式）\n• 矛盾地，亞維農教皇宮廷成為世俗音樂的主要贊助中心\n• 主要體裁：以 ballade 為主的 formes fixes chansons\n• 對象：貴族、教士、博學的鑑賞家——極端精緻的藝術品", {
    x: 0.5, y: 1.6, w: 9.0, h: 1.4, fontSize: 10, color: C.cream, fontFace: "Calibri",
  });

  // Features
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.15, w: 9.4, h: 2.1, fill: { color: C.rust }, rounding: true });
  s.addText("🔬 特徵 Features", { x: 0.45, y: 3.23, w: 9.1, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 節奏複雜度達頂峰——直到 20 世紀才再見\n  Rhythmic complexity unseen again until the 20th century\n• 多重拍號同時進行、hemiolas、鏈式切分音、刻意模糊的和聲\n• 記譜特技：紅色與黑色音符交錯、以心形寫成情歌、以圓形寫成卡農\n• Philipoctus de Caserta — En remirant vo douce pourtraiture（NAWM 28）\n• 同時期北法樂師行會發展出較簡單的世俗複音——描繪市井、狩獵等日常場景\n• 此風格僅流行一代人——20 世紀作曲家如 Messiaen、Ligeti 受其啟發", {
    x: 0.5, y: 3.56, w: 9.0, h: 1.65, fontSize: 9.5, color: C.cream, fontFace: "Calibri",
  });
}

// ── SLIDE 9 · Italian Trecento Music ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("義大利 Trecento 音樂", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Italian Music of the Fourteenth Century · The Trecento", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Context
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 1.1, fill: { color: "3A2015" }, rounding: true });
  s.addText("🇮🇹 背景：義大利是城邦聯合體，Trecento = \"mille trecento\" = 1300 年代", { x: 0.45, y: 1.36, w: 9.1, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 主要中心：Bologna、Padua、Milan、Perugia、Naples——尤其是佛羅倫斯\n• 義大利的記譜系統與法國不同：breve 可分為 2/3/4/6/8/9/12 semibreves；以點代替小節線\n• 教會複音多為即興 · 現存主要是世俗歌曲——為菁英觀眾創作", {
    x: 0.5, y: 1.68, w: 9.0, h: 0.72, fontSize: 9.5, color: C.sand, fontFace: "Calibri",
  });

  // Three genres
  const genres = [
    ["🎵 Madrigal", "14 世紀的牧歌", "兩到三聲部 · 全部聲部唱相同歌詞\n• 詩節結構：aa b（ritornello 尾韻）\n• 田園、諷刺、或情詩主題\n• Jacopo da Bologna: Non al suo amante（NAWM 29）"],
    ["🐎 Caccia", "狩獵歌", "字面意為「狩獵」· 上兩聲部嚴格卡農\n• 下方 untexted tenor 緩慢支持\n• 主題：狩獵、市集、戰鬥——\n  生動的對話與回聲 · 常用 hocket\n• Landini: Così pensoso（NAWM 30）"],
    ["💃 Ballata", "舞曲歌", "「ballare」= 跳舞 · AbbaA 形式\n• 類似法國 virelai 的單一詩節\n• 1365 年後多為 2–3 聲部複音\n• 高聲部主導（受法國 chanson 影響）\n• Landini: Non avrà ma' pietà（NAWM 31）"],
  ];
  genres.forEach(([name, tag, desc], i) => {
    const x = 0.3 + i * 3.18;
    const w = 3.05;
    s.addShape(pres.ShapeType.rect, { x, y: 2.55, w, h: 2.7, fill: { color: "3A2015" }, rounding: true });
    s.addText(name, { x: x + 0.12, y: 2.62, w: w - 0.2, h: 0.3, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
    s.addText(tag, { x: x + 0.12, y: 2.9, w: w - 0.2, h: 0.25, fontSize: 9, color: C.sand, italic: true, fontFace: "Calibri" });
    s.addShape(pres.ShapeType.rect, { x: x + 0.2, y: 3.22, w: w - 0.4, h: 0.02, fill: { color: C.gold } });
    s.addText(desc, { x: x + 0.15, y: 3.3, w: w - 0.25, h: 1.9, fontSize: 8.5, color: C.cream, fontFace: "Calibri" });
  });
}

// ── SLIDE 10 · Francesco Landini ─────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("Francesco Landini (ca. 1325–1397)", { x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 22, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("The Foremost Italian Musician of the Trecento", { x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 12, color: C.midBrown, fontFace: "Calibri" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.04, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Bio
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.wine }, rounding: true });
  s.addText("📖 生平", { x: 0.45, y: 1.28, w: 4.3, h: 0.35, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.55, y: 1.66, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("• 出生於佛羅倫斯（或鄰近 Fiesole）\n• 畫家之子 · 幼年染天花失明\n• 轉向音樂——成為卓越的演奏家、作曲家、詩人\n• 擅長多種樂器，尤以 organetto\n  （小型可攜式管風琴）聞名\n• 1361–65 任 Santa Trinità 管風琴師\n• 1365–97 任 San Lorenzo 教堂 chaplain\n\n❖ 軼事（Giovanni da Prato 記載）：\n  據說當 Landini 彈奏 organetto 時，\n  樹上的鳥群會停止歌唱而傾聽，\n  其中一隻夜鶯棲在他頭上\n  的樹枝繼續鳴叫", {
    x: 0.5, y: 1.78, w: 4.2, h: 3.4, fontSize: 9, color: C.cream, fontFace: "Calibri",
  });

  // Music characteristics
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.2, w: 4.6, h: 4.0, fill: { color: C.midBrown }, rounding: true });
  s.addText("🎼 音樂特徵 Musical Style", { x: 5.25, y: 1.28, w: 4.3, h: 0.35, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 5.35, y: 1.66, w: 4.1, h: 0.02, fill: { color: C.gold } });
  s.addText("❖ 作品：140 ballate · 12 madrigals\n  · 1 caccia · 1 virelai —— 無宗教音樂\n\n❖ 旋律優雅如拱——級進為主、弧形線條\n  · 比 Machaut 更平滑流暢\n\n❖ 三度與六度的和聲頻繁出現——甜美\n\n❖ 每行詩的首末音節加花唱，\n  中間部分清楚的音節式宣讀\n\n❖ Landini 終止式（Landini cadence）：\n  終止前上聲部先下行到下鄰音，\n  再上跳三度到主音——\n  從他開始成為風格標誌\n  在 14 世紀末到 15 世紀初法義音樂中普及\n\n❖ 埋葬於 San Lorenzo\n  墓碑上刻著他彈奏 organetto 的形象", {
    x: 5.3, y: 1.78, w: 4.2, h: 3.4, fontSize: 9, color: C.cream, fontFace: "Calibri",
  });
}

// ── SLIDE 11 · Performance Practice ──────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("演出實務 Performance Practice", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Voices, Instruments, and Musica Ficta", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Haut / Bas + voices or instruments
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 1.85, fill: { color: "3A2015" }, rounding: true });
  s.addText("🎺 Haut 與 Bas 樂器 · Voices or Instruments?", { x: 0.45, y: 1.36, w: 9.1, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 14–16 世紀以音量區分樂器：haut「高」(響亮) · bas「低」(柔和)\n  Haut: shawms, trumpets, cornetts · Bas: harps, vielles, lutes, psalteries, recorders\n• 長期爭論：聲部由人聲唱還是樂器奏？\n  Christopher Page · David Fallows 等 1970–80 年代學者論證：14 世紀複音通常\n  為每聲部一位歌者，無器樂——\n• 但宮廷世俗場合可能有器樂參與 · Gothic Voices 等團體的錄音證實此觀點", {
    x: 0.5, y: 1.7, w: 9.0, h: 1.45, fontSize: 9.5, color: C.sand, fontFace: "Calibri",
  });

  // Musica ficta
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.25, w: 9.4, h: 2.0, fill: { color: "3A2015" }, rounding: true });
  s.addText("🎵 Musica Ficta（「虛構的音樂」）", { x: 0.45, y: 3.31, w: 9.1, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 演出時對某些音作半音調整——避免 F–B 三全音、平滑旋律、終止甜美化\n• 稱為 ficta（feigned）因為這些音落在 Guido 手的 musica recta 範圍外\n• 演唱者受訓練於演出時即時判斷何時應升/降半音——作曲家只在必要處寫出\n• Double leading-tone cadence（雙導音終止）——上下兩聲部皆半音上行至完全協和\n  成為 14–15 世紀音樂的招牌聲響\n• Prosdocimo de' Beldomandi（卒 1428）的 Contrapunctus (1412) 詳細討論此原則\n  「較完美協和音程越近，聲響越甜美」", {
    x: 0.5, y: 3.65, w: 9.0, h: 1.6, fontSize: 9.5, color: C.sand, fontFace: "Calibri",
  });
}

// ── SLIDE 12 · Key Figures ───────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.rust); bottomBar(s, C.rust);

  s.addText("關鍵人物 Key Figures", { x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 26, bold: true, color: C.rust, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const figures = [
    ["🎵", "Philippe de Vitry", "1291–1361", "Ars Nova 的奠基者——教會參事、詩人、Meaux 主教 · Cum statua / Hugo 為最早的 isorhythmic motet 之一"],
    ["👑", "Guillaume de Machaut", "ca. 1300–1377", "14 世紀最重要作曲家兼詩人 · Messe de Nostre Dame · 首位主動編纂全集的作曲家 · 展現作者意識"],
    ["📚", "Jehan des Murs", "ca. 1290–ca. 1355", "數學家、天文學家兼樂論家 · Ars nova 相關論文 · 1340 描述 mensuration signs"],
    ["🔻", "Jacobus de Ispania", "ca. 1260–after 1330", "Speculum musicae (ca. 1330)——最長的中世紀音樂論文 · 反對 Ars Nova · 擁護 ars antiqua"],
    ["🎨", "Philipoctus de Caserta", "fl. 1370s", "亞維農教廷作曲家兼樂論家 · \"Ars Subtilior\" 一詞源自其論文 · En remirant vo douce pourtraiture"],
    ["🎹", "Francesco Landini", "ca. 1325–1397", "Trecento 最重要作曲家 · 失明的 organetto 大師 · 140 ballate · Landini cadence 得名於他"],
    ["🎭", "Jacopo da Bologna", "fl. 1340s–60s", "Trecento 早期重要 madrigal 作曲家 · Non al suo amante 為 Petrarch 詩譜曲"],
    ["📜", "Prosdocimo de' Beldomandi", "d. 1428", "Padua 大學教授 · Contrapunctus (1412) 最清楚地闡述 musica ficta 規則"],
  ];
  figures.forEach(([icon, name, date, desc], i) => {
    const y = 0.9 + i * 0.56;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.5, fill: { color: C.wine }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.08, w: 0.5, h: 0.35, fontSize: 16, align: "center", margin: 0 });
    s.addText(name, { x: 0.95, y: y + 0.03, w: 3.0, h: 0.24, fontSize: 10, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(date, { x: 0.95, y: y + 0.25, w: 3.0, h: 0.22, fontSize: 8, color: C.sand, fontFace: "Calibri", margin: 0 });
    s.addShape(pres.ShapeType.rect, { x: 4.0, y: y + 0.08, w: 0.025, h: 0.35, fill: { color: C.gold } });
    s.addText(desc, { x: 4.1, y: y + 0.03, w: 5.5, h: 0.45, fontSize: 8.5, color: C.cream, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 13 · Timeline ──────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("歷史時間軸 Timeline", { x: 0.4, y: 0.18, w: 9.2, h: 0.52, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });

  const events = [
    ["ca. 1300",    "機械時鐘問世 · Mechanical clocks invented"],
    ["ca. 1305",    "Giotto · Scrovegni Chapel 壁畫"],
    ["1307",        "Dante · Divine Comedy 以托斯卡納方言寫成"],
    ["1309",        "教皇 Clement V 移駐亞維農「巴比倫之囚」"],
    ["ca. 1317",    "Roman de Fauvel 手稿（含 169 首音樂作品）"],
    ["ca. 1320",    "Philippe de Vitry · Cum statua/Hugo · Ars nova 論文"],
    ["ca. 1323",    "Machaut 入 John of Luxembourg 之服"],
    ["ca. 1330",    "Jacobus de Ispania · Speculum musicae"],
    ["1337–1453",   "英法百年戰爭 · Hundred Years' War"],
    ["ca. 1340",    "Jehan des Murs 描述 mensuration signs"],
    ["1340",        "Machaut 成為 Reims 主教座堂參事"],
    ["1347–50",     "黑死病 · Black Death 殺死歐洲 1/3 人口"],
    ["1348–53",     "Boccaccio · Decameron"],
    ["ca. 1350s",   "Jacopo da Bologna · Non al suo amante"],
    ["ca. 1360s",   "Machaut · Messe de Nostre Dame"],
    ["1365–97",     "Francesco Landini 任 San Lorenzo 教堂 chaplain"],
    ["1378–1417",   "教會大分裂 · Great Schism of the Papacy"],
    ["ca. 1370s",   "Philipoctus de Caserta · En remirant vo douce pourtraiture"],
    ["ca. 1387–1400","Chaucer · Canterbury Tales"],
    ["ca. 1410–15", "Squarcialupi Codex 抄本完成"],
  ];
  s.addShape(pres.ShapeType.rect, { x: 2.6, y: 0.82, w: 0.05, h: 4.58, fill: { color: C.gold } });
  events.forEach(([date, event], i) => {
    const y = 0.82 + i * 0.228;
    s.addShape(pres.ShapeType.ellipse, { x: 2.47, y: y + 0.04, w: 0.26, h: 0.26, fill: { color: C.gold } });
    s.addText(date, { x: 0.1, y, w: 2.28, h: 0.26, fontSize: 8, color: C.sand, fontFace: "Calibri", align: "right", margin: 0 });
    s.addText(event, { x: 2.92, y, w: 6.8, h: 0.26, fontSize: 8, color: C.lightText, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 14 · Echoes of the New Art (Summary) ───────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("新藝術的迴響 Echoes of the New Art", { x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontSize: 24, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("Chapter Summary · Why This Music Mattered", { x: 0.4, y: 0.7, w: 9.2, h: 0.3, fontSize: 12, color: C.midBrown, fontFace: "Calibri" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.04, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const points = [
    ["📐", "Ars Nova 記譜是現代記譜的直系祖先——首次能精確記下任何節奏，包括切分音\nArs Nova notation is the direct ancestor of modern notation, first able to capture any rhythm"],
    ["✍", "精確記譜使作品能獨立流通——催生作曲家的「作者意識」，Machaut 是第一個主動編全集的例子\nAccurate notation enabled composers to take pride in authorship, first seen in Machaut"],
    ["⚖", "動盪的世紀產出最精緻的藝術——結構（isorhythm、formes fixes）與感官愉悅並重\nA turbulent century produced the most refined art: structure and sensory pleasure in balance"],
    ["🎵", "三度與六度的甜美與平行五八度的古老並存——14 世紀和聲的獨特聲響\nThirds and sixths coexist with parallel fifths and octaves—a distinctive 14th-century sound"],
    ["🇫🇷🇮🇹", "法國理性結構與義大利旋律流暢——為 15 世紀國際風格鋪路\nFrench structure + Italian lyricism foreshadow the 15th-century international style"],
    ["🎧", "Messiaen、Ligeti 受 Ars Subtilior 影響；Panda Bear 取樣 Machaut；Landini 進入 Judy Collins 的專輯\nModern echoes: Messiaen, Ligeti, Panda Bear's \"I'm Not\", Judy Collins's Wildflowers"],
  ];
  points.forEach(([icon, text], i) => {
    const y = 1.15 + i * 0.72;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.64, fill: { color: C.wine }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.08, w: 0.65, h: 0.48, fontSize: 16, align: "center", margin: 0 });
    s.addText(text, { x: 1.1, y: y + 0.05, w: 8.4, h: 0.58, fontSize: 10, color: C.cream, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 15 · Further Reading & Key Terms ───────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("延伸閱讀與補充教材\nFurther Reading & Supplementary Resources", { x: 0.4, y: 0.2, w: 9.2, h: 0.72, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.98, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Listening
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.1, w: 4.3, h: 2.4, fill: { color: "3A2015" }, rounding: true });
  s.addText("🎧 聆聽 Listen", { x: 0.55, y: 1.17, w: 4.0, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  const listening = [
    "Vitry — Cum statua / Hugo (NAWM 24)  youtu.be/cS9bdjm0hN0",
    "Machaut — Messe: Kyrie (NAWM 25a)  youtu.be/JtFMfmG5VlY",
    "Machaut — Messe: Gloria (NAWM 25b)  youtu.be/RCyvt86_2Ko",
    "Machaut — Douce dame jolie (NAWM 26)  youtu.be/pSjXxAOkSM8",
    "Machaut — Rose, liz, printemps (NAWM 27)  youtu.be/VYY1WO6FimA",
    "Caserta — En remirant (NAWM 28)  youtu.be/_LziNn1jpf0",
    "Jacopo — Non al suo amante (NAWM 29)  youtu.be/SqDSGNZmtUw",
    "Landini — Così pensoso (NAWM 30)  youtu.be/eB0QXbTsB4U",
    "Landini — Non avrà ma' pietà (NAWM 31)  youtu.be/TXpgPLW_6IQ",
  ];
  s.addText(listening.map((l, i) => ({ text: l, options: { bullet: true, breakLine: i < listening.length - 1, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2 } })), { x: 0.55, y: 1.58, w: 4.0, h: 1.9 });

  // Reading
  s.addShape(pres.ShapeType.rect, { x: 4.9, y: 1.1, w: 4.7, h: 2.4, fill: { color: "3A2015" }, rounding: true });
  s.addText("📖 閱讀 Read", { x: 5.05, y: 1.17, w: 4.4, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  const reading = [
    "Daniel Leech-Wilkinson — The Modern Invention of Medieval Music",
    "Anne Walters Robertson — Guillaume de Machaut and Reims",
    "Elizabeth Eva Leach — Guillaume de Machaut: Secretary, Poet, Musician",
    "Michael Long — \"Musical Tastes in Fourteenth-Century Italy\"",
    "Virginia Newes — \"Writing, Reading and Memorizing\" (Early Music 18)",
    "Wikipedia: Ars Nova · Ars Subtilior · Trecento · Formes fixes",
  ];
  s.addText(reading.map((r, i) => ({ text: r, options: { bullet: true, breakLine: i < reading.length - 1, fontSize: 9, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2 } })), { x: 5.05, y: 1.58, w: 4.4, h: 1.9 });

  // Key terms
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 3.62, w: 9.2, h: 1.72, fill: { color: "3A2015" }, rounding: true });
  s.addText("🔑 本章關鍵術語 Key Terms", { x: 0.6, y: 3.69, w: 8.8, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia" });
  const terms = "Ars Nova · Ars Antiqua · Philippe de Vitry · Jehan des Murs · Minim · Mode / Time / Prolation · Mensuration sign · Perfect / Imperfect · Syncopation · Isorhythm · Talea · Color · Hocket · Ars Subtilior · Musica ficta · Musica recta · Double leading-tone cadence · Landini cadence · Formes fixes · Ballade · Rondeau · Virelai · Chanson baladée · Chant royal · Lai · Complainte · Treble-dominated style · Cantus / Tenor / Contratenor · Machaut · Messe de Nostre Dame · Roman de Fauvel · Trecento · Madrigal (14th c.) · Ritornello · Caccia · Chace · Ballata · Ripresa / Piedi / Volta · Francesco Landini · Organetto · Squarcialupi Codex · Haut / Bas instruments · Portative / Positive organ";
  s.addText(terms, { x: 0.6, y: 4.13, w: 8.8, h: 1.15, fontSize: 8.5, color: C.sand, fontFace: "Calibri" });
}

pres.writeFile({ fileName: "Ch06_Fourteenth_Century.pptx" })
  .then(() => console.log("✅ Ch06_Fourteenth_Century.pptx created successfully"))
  .catch(err => console.error("❌ Error:", err));
