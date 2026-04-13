const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 14: The Invention of Opera";
pres.author = "A History of Western Music, 10th ed.";

const C = {
  darkBg:   "2A0A1A",
  gold:     "D4A830",
  cream:    "FBF5E6",
  crimson:  "8A1030",
  rose:     "B03050",
  darkText: "2A0A1A",
  lightText:"FBF5E6",
  sand:     "E8D8A8",
  slate:    "3A1A2A",
  blush:    "E8A0B0",
};

function darkSlide(p) { const s = p.addSlide(); s.background = { color: C.darkBg }; return s; }
function lightSlide(p) { const s = p.addSlide(); s.background = { color: C.cream }; return s; }
function topBar(s, c) { s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.12, fill: { color: c || C.gold } }); }
function bottomBar(s, c) { s.addShape(pres.ShapeType.rect, { x: 0, y: 5.5, w: "100%", h: 0.125, fill: { color: c || C.gold } }); }

// ── SLIDE 1 · Title ──────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.15, fill: { color: C.gold } });
  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.47, w: "100%", h: 0.155, fill: { color: C.gold } });
  s.addText("A HISTORY OF WESTERN MUSIC · TENTH EDITION", { x: 0.5, y: 0.45, w: 9, h: 0.35, fontSize: 11, color: C.sand, charSpacing: 3, align: "center", fontFace: "Georgia" });
  s.addText("CHAPTER 14", { x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6 });
  s.addText("THE INVENTION\nOF OPERA", { x: 0.3, y: 1.5, w: 9.4, h: 2.0, fontSize: 36, color: C.lightText, bold: true, align: "center", fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.65, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Camerata · Peri · Monteverdi · L'Orfeo · Poppea · Venice", { x: 0.4, y: 3.8, w: 9.2, h: 0.4, fontSize: 13, color: C.sand, align: "center", fontFace: "Georgia" });
  s.addText("Textbook pp. 297–316", { x: 0.5, y: 4.8, w: 9, h: 0.3, fontSize: 11, color: C.gold, align: "center", fontFace: "Calibri" });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.crimson); bottomBar(s, C.crimson);
  s.addText("本章概覽 Chapter Overview", { x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.crimson, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });
  const sections = [
    ["🏛", "Forerunners 歌劇先驅", "Intermedio · Pastoral drama · Camerata · Greek tragedy revival"],
    ["🎤", "First Operas 首批歌劇", "Peri Euridice (NAWM 77) · Recitative 的誕生"],
    ["🌟", "Monteverdi L'Orfeo", "1607 首部偉大歌劇 · 戲劇力量 (NAWM 78)"],
    ["🏢", "Public Opera in Venice", "1637 第一座公共歌劇院 · 商業模式 · Impresario"],
    ["👑", "L'incoronazione di Poppea", "Monteverdi 最後傑作 · Pur ti miro (NAWM 79)"],
    ["🎭", "Opera as Drama", "Aria · Recitative · Cesti Orontea (NAWM 80)"],
  ];
  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.crimson }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 11, color: C.rose, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 3 · Forerunners ────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("歌劇的先驅", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Intermedio · Pastoral · Camerata · Galilei", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎭 Intermedio & Pastoral", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Intermedio（幕間劇）\n• 插入戲劇各幕之間的音樂/舞台表演\n• 1539 佛羅倫斯婚禮慶典——最早大型實例\n• 1589 Medici 婚禮 intermedi = 歷史上最奢華\n  — 6 段 · 歌唱、舞蹈、佈景\n  — 參與者包括 Caccini、Peri\n\nPastoral Drama（田園劇）\n• 以 Arcadia 為背景的舞台劇\n• Tasso《Aminta》(1573)\n• Guarini《Il pastor fido》(1590)\n• 結合音樂的田園場景\n• 提供歌劇題材——牧羊人、仙女、自然\n\n⚠ 問題：如何讓整部戲劇用歌唱？\n• 需要一種新的歌唱方式——\n  像說話一樣自然，但有音樂的情感力量", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🏛 Camerata & Greek Revival", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Florentine Camerata (ca. 1573–87)\n• Count Giovanni de' Bardi 主持的文化沙龍\n• 知識分子、詩人、音樂家聚會\n• 討論如何恢復古希臘悲劇的音樂力量\n\n🌟 Vincenzo Galilei (ca. 1520–1591)\n• Galileo 之父 · 魯特琴家\n• Dialogo della musica antica et moderna (1581)\n• 批評：複音使文字不可理解\n• 主張：只有單聲才能表達情感——\n  因為古希臘正是用單旋律吟唱\n\n📖 Girolamo Mei (1519–1594)\n• 研究古希臘音樂文獻的學者\n• 結論：希臘音樂為單聲、以旋律表達情感\n• 間接影響 Camerata 的方向\n\n💡 Camerata 的成果\n• 催生 monody 與 recitative 的實驗\n• 催生歌劇——全劇歌唱的戲劇", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 4 · First Operas ───────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.crimson); bottomBar(s, C.crimson);
  s.addText("首批歌劇 · Peri Euridice", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.crimson, fontFace: "Georgia", align: "center" });
  s.addText("Dafne · Euridice · Recitative Style (NAWM 77)", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 13, color: C.rose, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.crimson } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "F8E0E8" }, rounding: true });
  s.addText("🌿 Dafne & Euridice", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.crimson, fontFace: "Georgia" });
  s.addText("Dafne (1598)\n• 第一部歌劇——音樂已遺失\n• 詩：Rinuccini · 曲：Peri (可能含 Corsi)\n• 私人上演於 Corsi 宮邸\n\nEuridice (1600)\n• 首部完整保存的歌劇\n• 為 Maria de' Medici 與 Henri IV 婚禮而作\n• 兩版：Peri 版（1601 出版）· Caccini 版（1601 搶先出版）\n\n🌟 Jacopo Peri (1561–1633)\n• 佛羅倫斯歌手/作曲家\n• 發明「stile recitativo」（吟唱風格）\n• Euridice 序言描述其方法：\n  — 介於歌唱與說話之間\n  — Bass 在和聲變化點才移動\n  — 聲線模仿自然語調", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "F8E0E8" }, rounding: true });
  s.addText("🎵 NAWM 77: Euridice 片段", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.crimson, fontFace: "Georgia" });
  s.addText("Recitative 的特徵\n• Voice 自由跟隨語言節奏\n• Bass 在語意轉折處才移動\n• 不協和音自然出現——表現情感\n• 「heightened speech」——提升的說話\n\n三種歌唱層次（Peri 的區分）\n1. Narrative recitative · 敘事\n   — 低張力 · 快速 · bass 簡單\n2. Expressive recitative · 情感高潮\n   — 不協和增多 · 更多裝飾\n3. Song-like passages · 近似歌曲\n   — 有旋律輪廓 · 但仍非正式 aria\n\n💡 為何重要？\n• Recitative 成為歌劇的基石——\n  推進情節、交代對話\n• 後來分化為 secco / accompagnato\n• 至今歌劇仍使用 recitative + aria 架構", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 5 · Monteverdi L'Orfeo ─────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("蒙特威爾第 · 奧菲歐", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("L'Orfeo (1607) · The First Great Opera (NAWM 78)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("📖 L'Orfeo 概述", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 1607 於 Mantua Gonzaga 宮廷首演\n• 劇本：Alessandro Striggio\n• Orpheus 與 Eurydice 的神話\n• 五幕 + Prologue（「音樂女神」開場）\n\n📊 音樂資源\n• 大型管弦樂團（約 40 人）——有具體編制指定\n• Toccata（銅管開場）→ 歌劇樂隊的先聲\n• 多種 recitative 風格交織\n• Strophic song · Dance · Chorus · Ritornello\n• 比 Euridice 豐富十倍的音樂語彙\n\n🌟 戲劇高潮\n• 第 2 幕：信使報告 Eurydice 之死\n  — Orpheus「Tu se' morta」——震撼的 recitative\n• 第 3 幕：Possente spirto（NAWM 78 核心）\n  — Orpheus 以歌聲感動冥府守門人\n  — 精美裝飾唱段 + 管弦伴奏", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎵 NAWM 78 分析", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Possente spirto / Tu se' morta\n\n「Possente spirto」\n• Orpheus 對冥河擺渡人 Caronte 的哀歌\n• Strophic variation 形式\n  — 每節相同的 bass pattern\n  — 旋律逐節裝飾加重\n• 兩版本並印：plain + ornamented\n  — 展示裝飾唱法的「教學」功能\n• 器樂伴奏逐節變化（小提琴→長笛→號角→低音提琴）\n\n「Tu se' morta」\n• 得知 Eurydice 死訊後的獨白\n• 簡潔 recitative——幾乎無裝飾\n• 半音下行 · 不協和 · 沉重\n• 音樂史最動人的 recitative 之一\n\n💡 L'Orfeo 的歷史意義\n• 首部融合所有現有技法的歌劇\n• 證明音樂戲劇可以達到感人至深\n• 至今仍是常演劇目", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 6 · Public Opera in Venice ─────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.crimson); bottomBar(s, C.crimson);
  s.addText("威尼斯公共歌劇", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.crimson, fontFace: "Georgia", align: "center" });
  s.addText("Public Opera · Impresario · Diva · Poppea (NAWM 79)", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 13, color: C.rose, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.crimson } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "F8E0E8" }, rounding: true });
  s.addText("🏢 1637：商業歌劇誕生", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.crimson, fontFace: "Georgia" });
  s.addText("• 1637 威尼斯 Teatro San Cassiano\n  — 第一座售票的公共歌劇院\n• 17 世紀末威尼斯已有 9 座歌劇院\n• 由 impresario（經理人）經營——自負盈虧\n\n📊 商業模式\n• 貴族投資場地——impresario 負責製作\n• 觀眾買票入場——非宮廷獨享\n• 歌手是最大賣點——明星制度誕生\n• Castrato · Soprano 成為偶像\n\n🎭 歌劇的變化\n• 題材從神話 → 歷史 · 宮廷陰謀\n• 減少合唱 · 降低成本\n• 增加獨唱：aria 越來越重要\n• Recitative + Aria 分化日益明顯\n• 佈景機關驚人——飛天、變景、特效", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "F8E0E8" }, rounding: true });
  s.addText("👑 Poppea · Pur ti miro (NAWM 79)", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.crimson, fontFace: "Georgia" });
  s.addText("L'incoronazione di Poppea (1643)\n• Monteverdi 最後一部歌劇\n• 第一部以歷史（非神話）為題材的歌劇\n• Nero 與 Poppea 的愛情/陰謀\n\n📊 特色\n• 道德模糊——邪惡戰勝正義、色慾得償\n• 角色心理刻畫——比之前任何歌劇都深入\n• Recitative 極富表情 · aria 旋律優美\n\n🎵 「Pur ti miro」終場二重唱\n• Nero 與 Poppea 的愛情二重唱\n• 簡單的 ground bass pattern\n• 兩聲交纏 · 旋律模仿 · 合三唱\n• 西方歌劇最美的一首二重唱\n• 可能非 Monteverdi 親筆——學術有爭議\n\n💡 Poppea 的意義\n• 歌劇從宮廷→市場——內容也隨之世俗化\n• 為後世歌劇（Handel · Mozart · Verdi）鋪路", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 7 · Cesti & Opera Conventions ──────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("歌劇慣例的成形", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Cesti · Orontea (NAWM 80) · Aria & Recitative 分化", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🌟 Antonio Cesti (1623–1669)", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 方濟會修士——卻成為最受歡迎的歌劇作曲家\n• 活躍於威尼斯、Innsbruck、維也納\n• 超過 100 部歌劇（多數已失）\n\n📖 Orontea (1656)\n• 17 世紀中期最受歡迎的歌劇之一\n• 女王 Orontea 的愛情喜劇\n\n🎵 NAWM 80: Intorno all'idol mio\n• Orontea 的獨唱 Aria\n  — 她在夢中對睡著的情人傾訴愛意\n• 結構：ABA（da capo 的前身）\n• 旋律優美 · 抒情 · 感傷\n• 已接近後世 bel canto 的美學\n• Bass 更有規律——走向 tonal harmony", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("📊 Aria vs. Recitative 分化", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("17 世紀中期的趨勢\n\nRecitative（宣敘調）\n• 推進情節 · 對話 · 敘事\n• 語音式 · 自由節拍\n• 簡單的 continuo 伴奏\n• 「drama 的引擎」\n\nAria（詠嘆調）\n• 表達情感 · 反思 · 獨白\n• 有固定旋律 · 節拍規則\n• 旋律美 · 展現歌手技巧\n• 逐漸成為觀眾最愛\n\nArioso\n• 介於 recitative 與 aria 之間\n• 比 recitative 更旋律化\n• 但未成為獨立曲段\n\n💡 這個分化到 1660s 基本完成\n• 之後歌劇 = recitative (敘事) + aria (情感)\n• Da capo aria (ABA) 到 1680s 成為標準\n• 此架構延續到 Mozart 甚至 Verdi", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 8 · Timeline ───────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.crimson); bottomBar(s, C.crimson);
  s.addText("時間軸 · Timeline", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.crimson, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 0.82, w: 5, h: 0.04, fill: { color: C.crimson } });
  const events = [
    ["ca. 1573", "Camerata meetings at Bardi's"],
    ["1581", "Galilei · Dialogo"],
    ["1589", "Medici 婚禮 intermedi"],
    ["1598", "Peri/Corsi · Dafne (music lost)"],
    ["1600", "Peri · Euridice (première)"],
    ["1601", "Caccini · Euridice (published first)"],
    ["1602", "Caccini · Le nuove musiche"],
    ["1607", "Monteverdi · L'Orfeo"],
    ["1608", "Monteverdi · L'Arianna (lament survives)"],
    ["1637", "Teatro San Cassiano — public opera"],
    ["1640", "Monteverdi · Il ritorno d'Ulisse"],
    ["1643", "Monteverdi · L'incoronazione di Poppea"],
    ["1649", "Cavalli · Giasone"],
    ["1656", "Cesti · Orontea"],
  ];
  events.forEach(([date, desc], i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 0.3 + col * 4.8;
    const y = 1.0 + row * 0.55;
    s.addShape(pres.ShapeType.rect, { x, y, w: 1.1, h: 0.44, fill: { color: C.crimson } });
    s.addText(date, { x: x + 0.05, y: y + 0.06, w: 1.0, h: 0.32, fontSize: 9, bold: true, color: C.lightText, align: "center", fontFace: "Georgia" });
    s.addText(desc, { x: x + 1.2, y, w: 3.55, h: 0.44, fontSize: 8, color: C.darkText, fontFace: "Calibri", valign: "middle" });
  });
}

// ── SLIDE 9 · Key Terms & Listening ──────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("關鍵詞彙 · 延伸閱讀", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Key Terms & Further Reading", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🔑 Key Terms", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• opera · libretto · librettist\n• intermedio · pastoral drama\n• Florentine Camerata · Count Bardi\n• stile recitativo · recitative\n• monody · basso continuo\n• aria · arioso · da capo aria\n• strophic variation\n• secco recitative · accompagnato\n• castrato · prima donna · diva\n• impresario · public opera\n• basso ostinato · ground bass\n• ritornello · sinfonia\n• L'Orfeo · L'incoronazione di Poppea\n• Euridice · Orontea", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("📚 Further Reading & 🎧 Listening", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Carter & Goldthwaite. Orpheus in the Marketplace (2017)\n• Whenham, John (ed.). Claudio Monteverdi: Orfeo (1986)\n• Rosand, Ellen. Opera in Seventeenth-Century Venice (1991)\n• Heller, Wendy. Emblems of Eloquence (2003)\n\n🎧 NAWM 精選聆聽 (YouTube)\n• 77 · Peri · Euridice: Nel puro ardor  youtu.be/1hBM_keRlRo\n• 78 · Monteverdi · L'Orfeo: Possente spirto  youtu.be/ngeurQnM4qM\n• 79 · Monteverdi · Poppea: Pur ti miro  youtu.be/AjlIwv0ljX8\n• 80 · Cesti · Orontea: Intorno all'idol mio  youtu.be/u6k11o1JEkM", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

pres.writeFile({ fileName: "Ch14_Opera.pptx" })
  .then(fn => console.log(`✅ ${fn} created successfully`))
  .catch(err => console.error("❌ Error:", err));
