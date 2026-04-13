const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 9: Franco-Flemish Composers, 1450-1520";
pres.author = "A History of Western Music, 10th ed.";

const C = {
  darkBg:   "1F1B2E",
  gold:     "C8A020",
  cream:    "FBF5E6",
  wine:     "7A2830",
  violet:   "4A2C5A",
  darkText: "1F1B2E",
  lightText:"FBF5E6",
  midViolet:"6B4480",
  sand:     "E8D8A8",
  slate:    "3A3048",
  plum:     "5E2A4A",
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
  s.addText("CHAPTER 9", {
    x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });
  s.addText("FRANCO-FLEMISH COMPOSERS\n1450–1520", {
    x: 0.3, y: 1.5, w: 9.4, h: 2.0, fontSize: 34, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.65, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Ockeghem · Busnoys · Obrecht · Isaac · Josquin Desprez", {
    x: 0.4, y: 3.8, w: 9.2, h: 0.4, fontSize: 13, color: C.sand, align: "center", fontFace: "Georgia",
  });
  s.addText("Textbook pp. 180–204", {
    x: 0.5, y: 4.8, w: 9, h: 0.3, fontSize: 11, color: C.gold, align: "center", fontFace: "Calibri",
  });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.violet); bottomBar(s, C.violet);

  s.addText("本章概覽 Chapter Overview", { x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.violet, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const sections = [
    ["🌍", "Political Change 政治變局", "Spain 統一 1479 · Hapsburg 聯姻 · Charles V · 義大利戰爭 1494"],
    ["👑", "Ockeghem 奧克岡", "ca. 1420–1497 · 法國王室禮拜堂 · Missa prolationum · 卡農大師"],
    ["🎵", "Busnoys 布斯瓦", "ca. 1430–1492 · Je ne puis vivre · 勃艮第宮廷 · V–I 終止式先驅"],
    ["📜", "Generation 1480–1520", "Obrecht · Isaac · 模仿對位普及 · 國際化樂風"],
    ["⭐", "Josquin Desprez", "ca. 1450–1521 · 文藝復興最偉大 · text expression · 18+ 彌撒"],
    ["⛪", "Masses on Borrowed Material", "Cantus-firmus · Paraphrase · Imitation · 四種借材彌撒"],
  ];
  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.violet }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 11, color: C.midViolet, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 3 · Political Change ───────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("政治變局：歐洲版圖重劃", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Political Change & the Rise of National Monarchies", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Left — Kingdoms
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "2E2740" }, rounding: true });
  s.addText("🏰 新興王國 Emerging Kingdoms", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 1453 百年戰爭結束——法國收復大部分領土\n• 1453 君士坦丁堡陷落——東羅馬帝國滅亡\n• 1477 勃艮第 Charles the Bold 戰死\n  → 勃艮第公國分裂、大部分歸入 Hapsburg\n• 1479 卡斯提爾 Isabella + 亞拉岡 Ferdinand 統一 Spain\n• 1492 Spain 征服 Granada——收復失地運動完成\n• 1492 Columbus 抵達美洲", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 9.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 3,
  });

  // Right — Hapsburg & Italy
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "2E2740" }, rounding: true });
  s.addText("⚜ Hapsburg & 義大利戰爭", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Hapsburg 的婚姻帝國\n• 1478 Maximilian I 娶 Mary of Burgundy\n• 其孫 Charles V (1519–1556 神聖羅馬皇帝)\n  同時統治 Spain、尼德蘭、奧地利、美洲\n\n1494 義大利戰爭 Italian Wars\n• 法王 Charles VIII 入侵義大利\n• 法、西、神聖羅馬帝國長期爭奪義大利半島\n• 義大利城邦文化外溢——隨軍音樂家往返南北\n• 刺激印刷、出版業興盛\n\n文化效應\n• 北方（法蘭德斯）作曲家南下——義大利禮拜堂任職\n• 音樂家成為國際商品——跨國職涯", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 4 · Ockeghem Biography ─────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.violet); bottomBar(s, C.violet);

  s.addText("Johannes Ockeghem", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 32, bold: true, color: C.violet, fontFace: "Georgia", align: "center" });
  s.addText("ca. 1420–1497 · 法國王室禮拜堂首席作曲家", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 13, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.violet } });

  // Left — Career
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "EADFEC" }, rounding: true });
  s.addText("📖 生平 Career", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.violet, fontFace: "Georgia" });
  s.addText("• 生於今比利時法蘭德斯區（可能在 Saint-Ghislain）\n• 1443 在 Antwerp 大教堂擔任歌手\n• 1446–1448 服侍 Bourbon 公爵 Charles I\n• 1451 起進入法國王室禮拜堂\n  服侍 Charles VII、Louis XI、Charles VIII 三位國王\n• 1453 升任禮拜堂首席（premier chapelain）\n• 1459 獲授 Saint-Martin of Tours 司庫——顯赫俸祿\n• 名聲遠播——Josquin 寫 motet《Nymphes des bois》悼念其逝世", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 9, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 3,
  });

  // Right — Works Overview
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "EADFEC" }, rounding: true });
  s.addText("🎼 作品總覽 Works", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.violet, fontFace: "Georgia" });
  s.addText("• 13 首完整彌撒曲——是 Du Fay 之後最多者\n• 1 首 Requiem——現存最早的複音安魂曲\n• 10 首經文歌（motets）\n• 20 餘首世俗香頌（chansons）\n\n✨ 風格特徵\n• 聲部線條綿長、呼吸錯落——避免齊句\n• 喜好低音域——四部常延伸到男低音\n• 各聲部節奏獨立、對位高度流動\n• 神秘、沉思的氣質——與當時北方神秘主義共鳴\n\n著名彌撒\n• Missa prolationum（mensuration canon 傑作）\n• Missa cuiusvis toni（可按任一調式演唱）\n• Missa De plus en plus（以 Binchois 香頌為 cantus firmus）", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 5 · Ockeghem's Masses & Canon Art ──────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Ockeghem 的對位藝術", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Canon, Mensuration & Modal Ingenuity", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Missa prolationum
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "2E2740" }, rounding: true });
  s.addText("🎵 Missa prolationum", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("「節拍比例彌撒」——中世紀以來最精巧的技術之作\n\n• 四個聲部僅以兩行記譜——每行同時以兩種 mensuration 讀出\n• 每樂章為一組 double mensuration canon\n• 兩個卡農同時進行——兩對聲部以不同速度與時值展開\n• 卡農音程逐樂章擴大：從 unison → 二度 → 三度…… → 八度\n\n意義\n• 展現作曲家對 ars perfecta 的終極掌握\n• 同時保持流暢的對位與虔誠的宗教感\n• 以技巧表達神學完美——聲音神學 theologia sonans", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 9, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 3,
  });

  // Cuiusvis toni & other innovations
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "2E2740" }, rounding: true });
  s.addText("🎼 Missa cuiusvis toni & 其他", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Missa cuiusvis toni「可用任何調式之彌撒」\n• 同一份樂譜可用 Dorian / Phrygian / Lydian / Mixolydian 任意演唱\n• 只需更換 clef——旋律關係隨調式改變\n• 同一作品產生四種不同的聲響世界\n\nRequiem（Missa pro defunctis）\n• 現存最早的複音安魂曲（可能早於 Du Fay 失傳版）\n• 保留大量葛利果 chant——以 paraphrase 華麗化\n\n風格總評\n• 後世視 Ockeghem 為「對位魔法師」\n• Tinctoris 列入最具才華的現代作曲家\n• 影響 Josquin、Obrecht、Pierre de la Rue 一整代", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 6 · Busnoys & Je ne puis vivre ─────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("Antoine Busnoys", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 32, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("ca. 1430–1492 · 勃艮第末代香頌大師", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 13, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  // Left — Bio
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "F3E0E0" }, rounding: true });
  s.addText("📖 生平", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("• 生於法蘭德斯\n• 1461 起服侍 Burgundian 宮廷\n• 服侍 Philip the Good、Charles the Bold 兩位公爵\n• 1477 Charles the Bold 戰死後轉投 Hapsburg 宮廷\n• 卒於 Bruges\n\n創作總覽\n• 2 部完整彌撒（含 Missa O crux lignum）\n• 8 首經文歌\n• 60+ 首世俗香頌——大多數為 rondeau 形式\n• 風格介於 Ockeghem 流動對位 與 更現代的清晰旋律之間", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 9, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 3,
  });

  // Right — Je ne puis vivre
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "F3E0E0" }, rounding: true });
  s.addText("🎵 Je ne puis vivre (NAWM 38)", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.wine, fontFace: "Georgia" });
  s.addText("「若非如此我不能活」——三聲部 virelai\n\n• 結構：bergerette（單節 virelai）\n• 上聲部主導旋律——下兩聲部伴奏性質\n• 規律節奏、清晰樂句——與 Ockeghem 迥異\n\n🌟 歷史意義\n• 低音在終止式使用跳進四／五度\n• 產生近代意義上的 V–I 終止感\n• 成為後世 authentic cadence 的遠祖\n• 可視為從中世紀終止式走向調性終止的關鍵一步\n\n• 文本可能為 Busnoys 自撰——字首藏詩暗示 Jacqueline d'Hacqueville，一位宮廷女性——作曲家的戀慕對象", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 7 · Generation of 1480-1520 ────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("1480–1520 世代", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("The Generation of 1480–1520 · Imitation Becomes the Norm", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const composers = [
    { name: "Jacob Obrecht", dates: "1457/8–1505", region: "法蘭德斯 → 義大利 (Ferrara)", focus: "~30 彌撒 · Missa Fortuna desperata · 數學式架構" },
    { name: "Heinrich Isaac", dates: "ca. 1450–1517", region: "法蘭德斯 → Florence → Maximilian 宮廷", focus: "Choralis Constantinus · Innsbruck ich muss dich lassen" },
    { name: "Josquin Desprez", dates: "ca. 1450–1521", region: "法蘭德斯 → 米蘭 / 羅馬 / Ferrara", focus: "18+ 彌撒 · Ave Maria · 文字表達革命者" },
    { name: "Pierre de la Rue", dates: "ca. 1452–1518", region: "Hapsburg-Burgundian Chapel", focus: "低音偏好 · 30+ 彌撒 · 卡農技巧" },
    { name: "Alexander Agricola", dates: "ca. 1446–1506", region: "Milan · Florence · Spanish Court", focus: "綿長旋律線 · 接近 Ockeghem 晚期風格" },
    { name: "Loyset Compère", dates: "ca. 1445–1518", region: "Milan → 法國王室", focus: "新式 motet · motetti missales" },
  ];

  composers.forEach((c, i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 0.3 + col * 4.8;
    const y = 1.35 + row * 1.3;
    s.addShape(pres.ShapeType.rect, { x, y, w: 4.6, h: 1.2, fill: { color: "2E2740" }, rounding: true });
    s.addText(c.name, { x: x + 0.15, y: y + 0.08, w: 3, h: 0.28, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
    s.addText(c.dates, { x: x + 3.1, y: y + 0.08, w: 1.4, h: 0.28, fontSize: 9, color: C.sand, fontFace: "Calibri", align: "right" });
    s.addText(c.region, { x: x + 0.15, y: y + 0.38, w: 4.35, h: 0.26, fontSize: 9, italic: false, color: "C0A0D8", fontFace: "Calibri" });
    s.addText(c.focus, { x: x + 0.15, y: y + 0.62, w: 4.35, h: 0.55, fontSize: 8.5, color: C.sand, fontFace: "Calibri" });
  });
}

// ── SLIDE 8 · Jacob Obrecht & Missa Fortuna desperata ────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.violet); bottomBar(s, C.violet);

  s.addText("Jacob Obrecht", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 32, bold: true, color: C.violet, fontFace: "Georgia", align: "center" });
  s.addText("1457/8–1505 · 結構宏大的彌撒建築師", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 13, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.violet } });

  // Bio
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "EADFEC" }, rounding: true });
  s.addText("📖 生平", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.violet, fontFace: "Georgia" });
  s.addText("• 生於 Ghent（根特）——父親為城市吹奏手\n• Utrecht、Cambrai、Bruges、Antwerp 主教座堂任職\n• 1487 及 1504 兩度應邀至 Ferrara 宮廷\n• 1505 於 Ferrara 因瘟疫過世\n\n作品總覽\n• 約 30 部彌撒——僅次於 Palestrina 的文藝復興產量\n• 30+ 經文歌\n• 30+ 世俗歌曲（法文、荷蘭文、義大利文）\n• 為 Gregorian chant、俗歌、器樂曲均曾取材作 cantus firmus", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 9, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 3,
  });

  // Missa Fortuna desperata
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "EADFEC" }, rounding: true });
  s.addText("🎼 Missa Fortuna desperata", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.violet, fontFace: "Georgia" });
  s.addText("取自 Busnoys 同名義大利文香頌為 cantus firmus\n\n✨ 技術特徵\n• 四聲部——tenor 持 cantus firmus\n• 每樂章重排 c.f. 的節奏與時值（拆解重組）\n• 在其他三聲部中啟動「點狀模仿」(points of imitation)\n  — 短動機於各聲部依次陳述\n• 段落分明、架構如建築般對稱\n\n🌟 Obrecht 的貢獻\n• 把模仿對位從偶然裝飾提升為整體結構原則\n• 將 cantus-firmus 技術推向數學極限\n• 為 Josquin 的 imitation mass 鋪路", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 9 · Heinrich Isaac ─────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Heinrich Isaac", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 30, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("ca. 1450–1517 · 國際化的北方作曲家", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.12, w: 5, h: 0.04, fill: { color: C.gold } });

  // Left — Career
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "2E2740" }, rounding: true });
  s.addText("🌍 Career Across Europe", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 生於法蘭德斯\n• 1485–1496 Florence Medici 宮廷——與 Lorenzo il Magnifico 同時代\n  為卡內瓦爾歌曲、儀式樂配樂\n• 1497 進入 Maximilian I 神聖羅馬宮廷——任 Hofkomponist\n• 晚年 1514 退居 Florence\n• 作品跨越拉丁、德文、義大利文、法文\n• 國際名聲：Paul Hofhaimer 作詩云\n  「Harmoniae chorales cantorum principi」\n  （合唱和聲歌手之君）", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 9, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 3,
  });

  // Right — Works
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "2E2740" }, rounding: true });
  s.addText("🎼 代表作品", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Choralis Constantinus\n• 三冊巨型 Proper（專用經）合唱集\n• 原為 Constance 主教座堂委託\n• 涵蓋教會年的 Introit、Alleluia、Sequence、Communion\n• 死後出版——對後世德語地區影響深遠\n\nPuer natus est (NAWM 40)\n• 聖誕日 Introit——四聲部 · paraphrase chant in tenor\n\nInnsbruck, ich muss dich lassen (NAWM 41)\n• 德語 Tenorlied · 四聲部\n• 流行旋律——後成 Lutheran chorale 《O Welt, ich muss dich lassen》底本\n• 被 Bach 引用於 St Matthew Passion", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 10 · Josquin Biography ─────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("Josquin Desprez", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 34, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("ca. 1450–1521 · 文藝復興最偉大的作曲家", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 13, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  // Career timeline
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 3.95, fill: { color: "F3E0E0" }, rounding: true });
  s.addText("📖 生平與職涯 Career", { x: 0.5, y: 1.38, w: 9, h: 0.32, fontSize: 13, bold: true, color: C.wine, fontFace: "Georgia" });

  const events = [
    ["ca. 1450", "生於 Condé-sur-l'Escaut 一帶（今法國北境）"],
    ["1475–1478", "服侍 René d'Anjou 於普羅旺斯"],
    ["1484–1489", "Milan Sforza 宮廷——與 Ascanio Sforza 樞機主教合作"],
    ["1489–1495", "羅馬教宗禮拜堂 Sistine Chapel 歌手"],
    ["1501–1503", "可能服務法國王室 Louis XII"],
    ["1503–1504", "Ferrara Ercole d'Este 宮廷——薪酬最高的作曲家之一"],
    ["1504–1521", "退居故鄉 Condé，任 Collegiate Church 議事司鐸"],
    ["1521.8.27", "逝世於 Condé——生前已是整個歐洲的音樂偶像"],
  ];
  events.forEach(([date, desc], i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 0.45 + col * 4.7;
    const y = 1.78 + row * 0.83;
    s.addText(date, { x, y, w: 1.2, h: 0.32, fontSize: 10, bold: true, color: C.wine, fontFace: "Georgia" });
    s.addText(desc, { x: x + 1.25, y, w: 3.35, h: 0.75, fontSize: 9, color: C.darkText, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 11 · Josquin's Motets & Text Expression ───────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Josquin 的經文歌與文字表達革命", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Motets · Text Declamation · Text Depiction · Text Expression", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Left — Ave Maria
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "2E2740" }, rounding: true });
  s.addText("🌹 Ave Maria...virgo serena (NAWM 44)", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 11, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("四聲部 motet——16 世紀初最著名的聖母頌\n\n• 開頭：四聲部依次以 point of imitation 進入——S→A→T→B\n• 每段經文對應不同織度：\n  — 同節奏 homorhythm 表示莊嚴祈求\n  — 成對二重唱 (duet pairs) 交替\n  — 卡農式走音展現謙卑\n• 結尾轉為同節奏 O Mater Dei, memento mei——\n  以最簡單的和聲結束最複雜的經文\n\n✨ 意義\n• 被 Petrucci 置於 1502 年 Motetti A 首篇——\n  印刷時代的第一件經文歌作品", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  // Right — 3 Techniques
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "2E2740" }, rounding: true });
  s.addText("📝 三種文字處理", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("1️⃣ Text Declamation 文字宣告\n• 節奏對應自然言語重音\n• 每個音節清晰可辨\n\n2️⃣ Text Depiction 文字描繪\n• 音樂形象呼應字詞具象意義\n• 「上升」用上行音階；「下降」用下行音階\n• 「嘆息」用短促休止；「哭泣」用半音\n\n3️⃣ Text Expression 文字表達\n• 音樂整體情緒對應文本情感\n• 不是模仿字面——而是捕捉心境\n• Josquin 首次系統性運用——被 Glareanus 稱為「音樂中的 Virgil」\n\n其他名作\n• Faulte d'argent (NAWM 42) · 五聲部卡農香頌\n• Mille regretz (NAWM 43) · 查理五世之愛歌", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 12 · Josquin's Masses ──────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.violet); bottomBar(s, C.violet);

  s.addText("Josquin 的彌撒曲", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.violet, fontFace: "Georgia", align: "center" });
  s.addText("18+ Masses · Four Great Examples", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 13, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.violet } });

  const masses = [
    { title: "Missa Pange lingua (NAWM 45)", type: "Paraphrase Mass", desc: "以 Pange lingua 聖體讚美詩為素材——旋律分解進入各聲部——成為 paraphrase mass 的典範" },
    { title: "Missa L'homme armé super voces musicales", type: "Cantus-firmus Mass", desc: "將 L'homme armé 曲調依序置於六個六音音階位置（ut–la）——技巧高超的系統性建構" },
    { title: "Missa Hercules dux Ferrariae", type: "Soggetto cavato", desc: "由 Ercole d'Este 大公名字的母音抽出「re-ut-re-ut-re-fa-mi-re」八音動機作主題——為 Ferrara 獻禮" },
    { title: "Missa Fortuna desperata", type: "Imitation / Parody 前身", desc: "取用 Busnoys 同名香頌全部聲部——預示了日後的 imitation mass" },
  ];

  masses.forEach((m, i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 0.3 + col * 4.8;
    const y = 1.35 + row * 1.95;
    s.addShape(pres.ShapeType.rect, { x, y, w: 4.6, h: 1.85, fill: { color: "EADFEC" }, rounding: true });
    s.addText(m.title, { x: x + 0.15, y: y + 0.08, w: 4.35, h: 0.32, fontSize: 12, bold: true, color: C.violet, fontFace: "Georgia" });
    s.addText(m.type, { x: x + 0.15, y: y + 0.42, w: 4.35, h: 0.26, fontSize: 9.5, color: C.wine, fontFace: "Georgia" });
    s.addText(m.desc, { x: x + 0.15, y: y + 0.7, w: 4.35, h: 1.1, fontSize: 9, color: C.darkText, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 13 · Masses on Borrowed Material (Comparison) ─────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("以借材為基礎的四種彌撒", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Four Types of Masses on Borrowed Material (Figure 9.5)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const rows = [
    ["類型 Type", "借材 Borrowing", "技術 Technique", "代表 Example"],
    ["Cantus-firmus mass", "chant 或俗歌的單一聲部", "置於 tenor，常以長音值出現", "Du Fay Missa Se la face"],
    ["Cantus-firmus / imitation", "單一聲部 + 模仿擴散", "tenor 持 c.f.，同時動機貫穿各部", "Obrecht Missa Fortuna desperata"],
    ["Paraphrase mass", "chant 旋律（單聲部）", "旋律華麗化、散布於所有聲部", "Josquin Missa Pange lingua"],
    ["Imitation / Parody", "整首複音作品（全部聲部）", "重用原作織度與動機全體", "Févin Missa Ave Maria（首例）"],
  ];
  const colW = [2.0, 2.8, 2.8, 2.2];
  const xStart = 0.3;
  rows.forEach((row, ri) => {
    let x = xStart;
    row.forEach((cell, ci) => {
      const isHeader = ri === 0;
      s.addShape(pres.ShapeType.rect, {
        x, y: 1.3 + ri * 0.78, w: colW[ci], h: 0.78,
        fill: { color: isHeader ? C.violet : "2E2740" },
        line: { color: C.gold, width: 0.5 },
      });
      s.addText(cell, {
        x: x + 0.1, y: 1.3 + ri * 0.78, w: colW[ci] - 0.2, h: 0.78,
        fontSize: isHeader ? 11 : 9, bold: isHeader,
        color: isHeader ? C.lightText : C.sand,
        fontFace: isHeader ? "Georgia" : "Calibri",
        valign: "middle", align: "center",
      });
      x += colW[ci];
    });
  });
}

// ── SLIDE 14 · Timeline & Legacy ─────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.violet); bottomBar(s, C.violet);

  s.addText("時間軸 · Timeline & Legacy", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.violet, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 0.82, w: 5, h: 0.04, fill: { color: C.violet } });

  const events = [
    ["1453", "百年戰爭結束 · 君士坦丁堡陷落"],
    ["1477", "Charles the Bold 戰死 · 勃艮第瓦解"],
    ["1479", "Ferdinand & Isabella 統一 Spain"],
    ["ca. 1450", "Ockeghem 任法國王室禮拜堂歌手"],
    ["1485", "Isaac 抵達 Florence Medici 宮廷"],
    ["1489", "Josquin 進入 Sistine Chapel"],
    ["1494", "義大利戰爭爆發"],
    ["1497", "Ockeghem 逝世 · Josquin 作輓歌"],
    ["1501", "Petrucci 出版 Odhecaton——第一部印刷複音集"],
    ["1502", "Petrucci Motetti A 以 Ave Maria 開卷"],
    ["1505", "Obrecht 於 Ferrara 瘟疫身亡"],
    ["1517", "Isaac 逝世 · Luther 發表九十五條論綱"],
    ["1521", "Josquin 逝世——文藝復興一代宗師"],
  ];
  events.forEach(([date, desc], i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 0.3 + col * 4.8;
    const y = 1.0 + row * 0.63;
    s.addShape(pres.ShapeType.rect, { x, y, w: 0.95, h: 0.52, fill: { color: C.violet } });
    s.addText(date, { x: x + 0.05, y: y + 0.06, w: 0.85, h: 0.4, fontSize: 11, bold: true, color: C.lightText, align: "center", fontFace: "Georgia" });
    s.addText(desc, { x: x + 1.05, y, w: 3.65, h: 0.52, fontSize: 9, color: C.darkText, fontFace: "Calibri", valign: "middle" });
  });
}

// ── SLIDE 15 · Key Terms & Further Reading ──────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("關鍵詞彙 · 延伸閱讀", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Key Terms & Further Reading", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "2E2740" }, rounding: true });
  s.addText("🔑 Key Terms", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• mensuration canon · 節拍比例卡農\n• soggetto cavato dalle vocali · 母音抽字主題\n• paraphrase mass · 華飾彌撒\n• imitation mass (parody mass) · 模仿／戲擬彌撒\n• cantus-firmus–imitation mass\n• points of imitation · 模仿點\n• text declamation / depiction / expression\n• Requiem (Missa pro defunctis)\n• bergerette · 單節 virelai\n• authentic cadence · 正格終止（前身）\n• Proper cycle · 專用經套曲 (Choralis Constantinus)\n• Tenorlied · 德文 tenor 歌曲", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 9, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 3,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "2E2740" }, rounding: true });
  s.addText("📚 Further Reading", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Fallows, David. Josquin (2009)\n• Sherr, ed. The Josquin Companion (2000)\n• Wegman. Born for the Muses: Obrecht (1994)\n• Picker. Henricus Isaac: A Guide (1991)\n• Planchart. Guillaume Du Fay (2018)\n\n🎧 NAWM 精選聆聽 (YouTube)\n• 38 · Busnoys · Je ne puis vivre  youtu.be/-hcvz1qvJwc\n• 39 · Ockeghem · Missa prolationum Kyrie  youtu.be/ZWLsLAujZzI\n• 40 · Isaac · Puer natus est  youtu.be/Y2k88BgrFBY\n• 41 · Isaac · Innsbruck ich muss...  youtu.be/Dk84ddXJ8tY\n• 42 · Josquin · Faulte d'argent  youtu.be/WLkbxj85bKU\n• 43 · Josquin · Mille regretz  youtu.be/1fSZ7sTYNTM\n• 44 · Josquin · Ave Maria  youtu.be/scQ5YBRpwNg\n• 45 · Josquin · Pange lingua Kyrie  youtu.be/_zxnFVWZVcE", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1,
  });
}

pres.writeFile({ fileName: "Ch09_Franco_Flemish.pptx" })
  .then(fn => console.log(`✅ ${fn} created successfully`))
  .catch(err => console.error("❌ Error:", err));
