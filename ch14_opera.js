const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 14: The Invention of Opera";
pres.author = "A History of Western Music, 10th ed.";

// Opera palette — theatrical plum / gold / cream
const C = {
  darkBg:   "2D1B2E",
  gold:     "C8A030",
  cream:    "F5F0E0",
  plum:     "5B2C6F",
  mauve:    "7D3C98",
  darkText: "2D1B2E",
  lightText:"F5F0E0",
  sand:     "E8D8A8",
  slate:    "3A1F3D",
  wine:     "6C2040",
  rose:     "884060",
};

function darkSlide(pres) { const s = pres.addSlide(); s.background = { color: C.darkBg }; return s; }
function lightSlide(pres) { const s = pres.addSlide(); s.background = { color: C.cream }; return s; }
function topBar(s, color) { s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.12, fill: { color: color || C.gold } }); }
function bottomBar(s, color) { s.addShape(pres.ShapeType.rect, { x: 0, y: 5.5, w: "100%", h: 0.125, fill: { color: color || C.gold } }); }

// ── SLIDE 1 · Title ─────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.15, fill: { color: C.gold } });
  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.47, w: "100%", h: 0.155, fill: { color: C.gold } });

  s.addText("A HISTORY OF WESTERN MUSIC · TENTH EDITION", {
    x: 0.5, y: 0.45, w: 9, h: 0.4, fontSize: 18, color: C.sand, charSpacing: 3, align: "center", fontFace: "Georgia",
  });
  s.addText("CHAPTER 14", {
    x: 0.5, y: 1.0, w: 9, h: 0.6, fontSize: 24, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });
  s.addText("THE INVENTION OF OPERA\n歌劇的誕生", {
    x: 0.3, y: 1.7, w: 9.4, h: 1.8, fontSize: 38, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.65, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Peri · Caccini · Monteverdi · Cavalli · Cesti", {
    x: 0.4, y: 3.85, w: 9.2, h: 0.45, fontSize: 20, color: C.sand, align: "center", fontFace: "Georgia",
  });
  s.addText("Textbook pp. 302–325", {
    x: 0.5, y: 4.8, w: 9, h: 0.35, fontSize: 18, color: C.gold, align: "center", fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 2 · Chapter Overview ──────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.plum); bottomBar(s, C.plum);

  s.addText("本章概覽 Chapter Overview", {
    x: 0.4, y: 0.25, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.plum, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.9, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "1.  歌劇前身：Intermedio 與田園劇 Pastoral Drama\n" +
    "2.  Florentine Camerata 與希臘復興理念\n" +
    "3.  Peri《Euridice》— 最早存世的完整歌劇\n" +
    "4.  Monteverdi《L'Orfeo》— 歌劇的突破\n" +
    "5.  威尼斯公共歌劇院 (1637) 與商業化\n" +
    "6.  Aria vs Recitative 的發展與定型",
    { x: 0.5, y: 1.1, w: 9.0, h: 4.2, fontSize: 22, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );
}

// ── SLIDE 3 · Precursors: Intermedio & Pastoral Drama ───────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("歌劇的前身 Precursors to Opera", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Intermedio · Pastoral Drama · Monody", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 20, color: C.sand, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "Intermedio 幕間劇\n" +
    "  • 戲劇幕間的音樂娛樂，含歌唱、舞蹈、奇觀佈景\n" +
    "  • 1589 年佛羅倫斯婚禮盛大 intermedi 最具影響力\n\n" +
    "Pastoral Drama 田園劇\n" +
    "  • 以神話牧歌為題材的舞台劇\n" +
    "  • 含合唱、獨唱段落 → 歌劇的直接先驅",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.8, fontSize: 20, color: C.lightText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.35 }
  );
}

// ── SLIDE 4 · Florentine Camerata & Greek Revival ───────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("佛羅倫斯同好會 Florentine Camerata", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.plum, fontFace: "Georgia",
  });
  s.addText("復興古希臘音樂的理想 Reviving Greek Music", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 22, color: C.wine, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• Giovanni de' Bardi 伯爵府上的知識分子聚會\n" +
    "• Vincenzo Galilei《古今音樂對話》(1581) 批判對位法\n" +
    "• 主張：古希臘人以單聲旋律 (monody) 感動聽者\n" +
    "• 目標：用歌唱模仿「說話」— 旋律應追隨語言\n" +
    "• 這個理念直接催生了 recitative 宣敘調",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.8, fontSize: 22, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );
}

// ── SLIDE 5 · Peri's Dafne & Euridice ───────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("首批歌劇 The First Operas", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Jacopo Peri (1561–1633)", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 22, color: C.sand, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "《Dafne》(1598) — 最早的歌劇（樂譜已佚）\n" +
    "  • Rinuccini 作詞、Peri (& Corsi) 作曲\n\n" +
    "《L'Euridice》(1600) — 最早完整存世的歌劇\n" +
    "  • 為法國國王亨利四世婚禮慶典演出\n" +
    "  • Caccini 也為同劇本譜曲 → 兩版並存\n" +
    "  • Peri 版更適合戲劇表達",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.8, fontSize: 21, color: C.lightText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.4 }
  );
}

// ── SLIDE 6 · Peri's Recitative Style ───────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.plum); bottomBar(s, C.plum);

  s.addText("Peri 的宣敘調風格 Recitative Style", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.plum, fontFace: "Georgia",
  });
  s.addText("\"介於歌唱與說話之間\" — halfway between speech and song", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 20, color: C.wine, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• 援引古希臘「diastematic motion」理論\n" +
    "• 持續低音固定和聲 → 歌聲自由地穿越協和與不協和\n" +
    "• 重音音節落在協和音上、經過音可不協和\n" +
    "• 兩種模式：敘事性宣敘調 vs 抒情性宣敘調\n" +
    "• 融合了牧歌、詠嘆調、田園劇的傳統",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.8, fontSize: 22, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );
}

// ── SLIDE 7 · NAWM 77: Peri Euridice ────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.mauve); bottomBar(s, C.mauve);

  s.addText("NAWM 77 — Peri《Euridice》選段", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Narrative & Expressive Recitative", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 22, color: C.sand, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "Tirsi 的詠嘆調 (strophic aria)\n" +
    "  • 以 sinfonia 引入 → ritornello 反覆\n" +
    "  • 節奏活潑如 canzonetta 舞曲風格\n\n" +
    "Dafne 報告死訊 — 敘事性宣敘調\n" +
    "  • 無固定節奏、無旋律型態 → 純語言節奏\n\n" +
    "Orfeo 哀歌 — 抒情性宣敘調\n" +
    "  • 休止符表達驚愕、不協和表現悲痛",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.4, fontSize: 19, color: C.lightText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.2 }
  );

  s.addText("https://www.youtube.com/watch?v=bt8KaCIGBEk", {
    x: 0.5, y: 5.15, w: 9.0, h: 0.3, fontSize: 16, color: C.gold, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=bt8KaCIGBEk" },
  });
}

// ── SLIDE 8 · Monteverdi Biography ──────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("蒙台威爾第 Claudio Monteverdi", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.plum, fontFace: "Georgia",
  });
  s.addText("ca. 1567–1643 · 歌劇藝術的奠基者", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 22, color: C.wine, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• 生於 Cremona · 少年即出版聲樂作品\n" +
    "• 1590 入 Mantua 宮廷 → 1601 任宮廷教堂主管\n" +
    "• 1613 任威尼斯聖馬可教堂樂長（至去世）\n" +
    "• 出版九冊牧歌集：從文藝復興到巴洛克的轉型\n" +
    "• 三部存世歌劇：L'Orfeo, Il ritorno d'Ulisse,\n" +
    "  L'incoronazione di Poppea",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.8, fontSize: 22, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );
}

// ── SLIDE 9 · L'Orfeo Overview ──────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("蒙台威爾第《奧菲歐》L'Orfeo (1607)", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Mantua · Libretto: Alessandro Striggio", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 22, color: C.sand, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• 五幕結構 · 每幕以 Orfeo 歌聲為中心\n" +
    "• 希臘悲劇式合唱團 · 開場與結尾呼應\n" +
    "• 豐富的器樂編制：木管銅管弦樂管風琴豎琴\n" +
    "• Strophic variation 分節變奏手法\n" +
    "• 比 Peri 更多樣的風格：從歌唱到痛苦的宣敘調",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.8, fontSize: 22, color: C.lightText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );
}

// ── SLIDE 10 · L'Orfeo Structure ────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.plum); bottomBar(s, C.plum);

  s.addText("《L'Orfeo》五幕結構", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.plum, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.9, w: 9.2, h: 0.03, fill: { color: C.gold } });

  // Left column
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.1, w: 4.55, h: 4.2, fill: { color: C.sand }, rectRadius: 0.1 });
  s.addText(
    "Prologue\n" +
    "  Music 女神以分節變奏詠唱\n\n" +
    "Act I — 婚禮慶典\n" +
    "  靜態拱形結構 · 合唱+牧歌\n\n" +
    "Act II — 喜轉悲\n" +
    "  使者報噩耗 · Tu se' morta 哀歌",
    { x: 0.45, y: 1.2, w: 4.25, h: 3.95, fontSize: 20, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.25 }
  );

  // Right column
  s.addShape(pres.ShapeType.rect, { x: 5.15, y: 1.1, w: 4.55, h: 4.2, fill: { color: C.sand }, rectRadius: 0.1 });
  s.addText(
    "Act III — 冥界入口\n" +
    "  Possente spirto 炫技唱段\n\n" +
    "Act IV — 回望失敗\n" +
    "  Orfeo 回頭 · 再失 Euridice\n\n" +
    "Act V — 升天\n" +
    "  Apollo 攜子升天 · 合唱結尾",
    { x: 5.3, y: 1.2, w: 4.25, h: 3.95, fontSize: 20, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.25 }
  );
}

// ── SLIDE 11 · NAWM 78a: Possente spirto ────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.mauve); bottomBar(s, C.mauve);

  s.addText("NAWM 78a — Possente spirto", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("L'Orfeo Act III · Orfeo 懇求冥河船夫 Caronte", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 20, color: C.sand, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• Strophic variation：每一節旋律不同但和聲相近\n" +
    "• 前四節寫出華麗裝飾奏 (ornamental version)\n" +
    "• 展示 Orfeo 超凡歌藝 → 說服冥界\n" +
    "• 器樂 obbligato 對話：小提琴、短號、豎琴\n" +
    "• 最後轉為樸素的宣敘調 → 真情流露",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.5, fontSize: 22, color: C.lightText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );

  s.addText("https://www.youtube.com/watch?v=LUvJmIcp7z0", {
    x: 0.5, y: 5.05, w: 9.0, h: 0.35, fontSize: 18, color: C.gold, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=LUvJmIcp7z0" },
  });
}

// ── SLIDE 12 · NAWM 78b: Tu se' morta ───────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("NAWM 78b — Tu se' morta", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.plum, fontFace: "Georgia",
  });
  s.addText("L'Orfeo Act II · Orfeo 的哀歌 Lament", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 22, color: C.wine, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• 抒情性宣敘調的巔峰之作\n" +
    "• 每一句音樂層層遞進 · 情感逐步升高\n" +
    "• 不協和音對持續和弦：表現 Orfeo 苦澀的感受\n" +
    "• E 大調→ G 小調：他仍活著而她已死去的諷刺\n" +
    "• 遠超 Peri 的單聲實驗 → 真正的戲劇抒情",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.8, fontSize: 22, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );

  s.addText("https://www.youtube.com/watch?v=MIadFIEB1cg", {
    x: 0.5, y: 5.05, w: 9.0, h: 0.35, fontSize: 18, color: C.wine, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=MIadFIEB1cg" },
  });
}

// ── SLIDE 13 · Opera from Florence to Rome ──────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("歌劇的傳播 Florence → Rome", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Gagliano · Francesca Caccini · Roman Opera", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 22, color: C.sand, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• Gagliano 新版《Dafne》(1608) 廣受好評\n" +
    "• Francesca Caccini《La liberazione di Ruggiero》\n" +
    "  (1625) — 女性作曲家的重要歌劇\n" +
    "• 1620s 羅馬：Barberini 家族贊助歌劇\n" +
    "• 羅馬歌劇發展出 arioso（半詠嘆調）\n" +
    "• 女性禁演 → castrati 閹伶登場",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.8, fontSize: 21, color: C.lightText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.4 }
  );
}

// ── SLIDE 14 · Public Opera in Venice ───────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.plum); bottomBar(s, C.plum);

  s.addText("威尼斯公共歌劇 Public Opera in Venice", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.plum, fontFace: "Georgia",
  });
  s.addText("Teatro San Cassiano (1637) — 歷史性的轉變", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 22, color: C.wine, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• 1637 第一座公共歌劇院開幕 · 售票入場\n" +
    "• Carnival 嘉年華季節演出 → 吸引各階層觀眾\n" +
    "• 至 1678 年威尼斯已有九座歌劇院競爭\n" +
    "• Impresario 經紀人制度：雇用作曲家與歌手\n" +
    "• 歌劇從貴族私人娛樂 → 商業化公眾藝術",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.8, fontSize: 22, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );
}

// ── SLIDE 15 · Venetian Opera Conventions ───────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("威尼斯歌劇的慣例 Venetian Conventions", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.9, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• 三幕結構（取代早期五幕）+ 序幕\n" +
    "• 題材：神話、史詩、歷史 — 含戲劇衝突與奇觀\n" +
    "• Prima donna / Primo uomo 明星制度\n" +
    "• 詠嘆調數量增至 50-60 首 · 合唱大幅減少\n" +
    "• Recitative: versi sciolti（自由詩）\n" +
    "  Aria: 規律韻律的押韻詩",
    { x: 0.5, y: 1.1, w: 9.0, h: 4.2, fontSize: 22, color: C.lightText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );
}

// ── SLIDE 16 · The Impresario & Diva ────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("經紀人與歌劇天后 Impresario & Diva", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.plum, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.9, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• Impresario 相當於現代製作人\n" +
    "  — 管理劇院、僱用全部人員、承擔盈虧\n" +
    "• Diva 現象始於 Anna Renzi (1640s)\n" +
    "  — 歌手比作曲家更有票房號召力\n" +
    "• 歌手薪酬可達作曲家的 2-6 倍\n" +
    "• Prima donna 影響劇本與角色創作",
    { x: 0.5, y: 1.1, w: 9.0, h: 4.2, fontSize: 22, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );
}

// ── SLIDE 17 · Monteverdi L'incoronazione di Poppea ─────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("《波佩雅的加冕》L'incoronazione di Poppea", {
    x: 0.3, y: 0.2, w: 9.4, h: 0.65, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Monteverdi · Venice 1643 · 常被視為其最高傑作", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 18, color: C.sand, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• 歷史題材：羅馬暴君 Nero 與情婦 Poppea\n" +
    "• 為商業劇院而寫 → 編制精簡\n" +
    "  （大鍵琴+低音提琴+兩把小提琴）\n" +
    "• 風格多變：宣敘調、詠嘆調、arioso 交替\n" +
    "• 人物刻畫與情感深度超越《Orfeo》",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.5, fontSize: 22, color: C.lightText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );
}

// ── SLIDE 18 · NAWM 79: Pur ti miro ────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.mauve); bottomBar(s, C.mauve);

  s.addText("NAWM 79 — Pur ti miro", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.plum, fontFace: "Georgia",
  });
  s.addText("L'incoronazione di Poppea · 終曲二重唱", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 22, color: C.wine, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• Nero 與 Poppea 的愛情二重唱\n" +
    "• 抒情三拍子 · 旋律優美而甜蜜\n" +
    "• 聲部交織 → 象徵兩人結合\n" +
    "• 歌劇史上最著名的終曲之一\n" +
    "• 學者爭議：可能非 Monteverdi 親筆",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.5, fontSize: 22, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );

  s.addText("https://www.youtube.com/watch?v=v_Se1XVkOiU", {
    x: 0.5, y: 5.05, w: 9.0, h: 0.35, fontSize: 18, color: C.plum, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=v_Se1XVkOiU" },
  });
}

// ── SLIDE 19 · Francesco Cavalli ────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("卡瓦利 Francesco Cavalli (1602–1676)", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("威尼斯歌劇的第一人 Leading Venetian Opera Composer", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 20, color: C.sand, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• Monteverdi 的學生 · 聖馬可管風琴師→樂長\n" +
    "• 1639–1673 作近 30 部歌劇 · 最成功的歌劇作曲家\n" +
    "• 與劇作家 Faustini 合作確立威尼斯歌劇慣例\n" +
    "• 音樂特色：宣敘調捕捉語言情感\n" +
    "  詠嘆調以三拍子優雅旋律見長",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.5, fontSize: 22, color: C.lightText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );
}

// ── SLIDE 20 · Cesti & Orontea / NAWM 80 ────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.plum); bottomBar(s, C.plum);

  s.addText("切斯提與《Orontea》Cesti & NAWM 80", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.plum, fontFace: "Georgia",
  });
  s.addText("Antonio Cesti (1623–1669) · Innsbruck Opera", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 22, color: C.wine, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• Innsbruck Tyrol 宮廷 · 義大利歌劇走向國際\n" +
    "• 《Orontea》(1656) 是 17 世紀最常上演的歌劇之一\n" +
    "• 義大利歌劇傳播至維也納、巴黎、英國\n" +
    "• NAWM 80: Intorno all'idol mio\n" +
    "  — 優美的搖籃曲式詠嘆調 · 旋律流暢",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.5, fontSize: 22, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );

  s.addText("https://www.youtube.com/watch?v=WjrNblgYAhw", {
    x: 0.5, y: 5.05, w: 9.0, h: 0.35, fontSize: 18, color: C.plum, fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=WjrNblgYAhw" },
  });
}

// ── SLIDE 21 · Aria vs Recitative Development ───────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("詠嘆調 vs 宣敘調的演化", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Aria vs Recitative — How they differentiated", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 22, color: C.sand, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  // Left column — Recitative
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.5, w: 4.55, h: 3.8, fill: { color: C.slate }, rectRadius: 0.1 });
  s.addText("Recitative 宣敘調", {
    x: 0.45, y: 1.6, w: 4.25, h: 0.45, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText(
    "• 推動劇情、對話\n" +
    "• 自由節奏 · 追隨語言\n" +
    "• Versi sciolti 自由詩\n" +
    "• 和聲簡單 · 音節式",
    { x: 0.5, y: 2.15, w: 4.2, h: 2.8, fontSize: 20, color: C.lightText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );

  // Right column — Aria
  s.addShape(pres.ShapeType.rect, { x: 5.15, y: 1.5, w: 4.55, h: 3.8, fill: { color: C.slate }, rectRadius: 0.1 });
  s.addText("Aria 詠嘆調", {
    x: 5.3, y: 1.6, w: 4.25, h: 0.45, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText(
    "• 抒發情感、反思\n" +
    "• 規律節拍 · 常為三拍\n" +
    "• 押韻定型詩行\n" +
    "• 旋律性強 · 流暢優美",
    { x: 5.35, y: 2.15, w: 4.2, h: 2.8, fontSize: 20, color: C.lightText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );
}

// ── SLIDE 22 · Combattimento & Stile Concitato ──────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("戰鬥與激情風格 Stile Concitato", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.plum, fontFace: "Georgia",
  });
  s.addText("Combattimento di Tancredi e Clorinda (1624)", {
    x: 0.4, y: 0.85, w: 9.2, h: 0.4, fontSize: 22, color: C.wine, fontFace: "Georgia", italic: true,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.3, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText(
    "• Tasso《耶路撒冷的解放》中的戰鬥場面\n" +
    "• 結合敘事者（男高音）、角色演唱與啞劇\n" +
    "• Stile concitato 激情風格：快速反覆同音\n" +
    "  → 弦樂 tremolo 技法的先驅\n" +
    "• Monteverdi 最具創新性的戲劇實驗之一",
    { x: 0.5, y: 1.5, w: 9.0, h: 3.8, fontSize: 22, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.5 }
  );
}

// ── SLIDE 23 · Timeline ─────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("歌劇誕生年表 Timeline", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.9, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const events = [
    ["1573", "Florentine Camerata 成立"],
    ["1598", "Peri《Dafne》首演（樂譜已佚）"],
    ["1600", "Peri & Caccini《Euridice》"],
    ["1607", "Monteverdi《L'Orfeo》· Mantua"],
    ["1624", "Monteverdi《Combattimento》"],
    ["1625", "F. Caccini《La liberazione di Ruggiero》"],
    ["1637", "威尼斯第一座公共歌劇院"],
    ["1643", "Monteverdi《L'incoronazione di Poppea》"],
    ["1656", "Cesti《Orontea》· Innsbruck"],
  ];

  events.forEach(([year, desc], i) => {
    const y = 1.05 + i * 0.48;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 1.2, h: 0.4, fill: { color: C.mauve }, rectRadius: 0.05 });
    s.addText(year, { x: 0.4, y, w: 1.2, h: 0.4, fontSize: 18, bold: true, color: C.lightText, fontFace: "Georgia", align: "center" });
    s.addText(desc, { x: 1.8, y, w: 7.8, h: 0.4, fontSize: 18, color: C.sand, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 24 · Key Terms ────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.plum); bottomBar(s, C.plum);

  s.addText("核心術語 Key Terms", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.plum, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.9, w: 9.2, h: 0.03, fill: { color: C.gold } });

  // Left column
  s.addText(
    "Recitative 宣敘調\n" +
    "  模仿語言的歌唱風格\n\n" +
    "Aria 詠嘆調\n" +
    "  抒情歌唱段落\n\n" +
    "Arioso 半詠嘆調\n" +
    "  介於宣敘調與詠嘆調之間\n\n" +
    "Strophic variation\n" +
    "  分節變奏",
    { x: 0.4, y: 1.05, w: 4.5, h: 4.3, fontSize: 19, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.2 }
  );

  // Right column
  s.addText(
    "Ritornello 反覆奏\n" +
    "  器樂反覆的間奏段\n\n" +
    "Sinfonia 序曲\n" +
    "  器樂前奏\n\n" +
    "Castrato 閹伶\n" +
    "  為保高音而閹割的歌手\n\n" +
    "Impresario 經紀人\n" +
    "  歌劇製作人",
    { x: 5.2, y: 1.05, w: 4.5, h: 4.3, fontSize: 19, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.2 }
  );
}

// ── SLIDE 25 · NAWM Listening Guide ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.mauve); bottomBar(s, C.mauve);

  s.addText("聆聽指南 NAWM Listening Guide", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.65, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.9, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const pieces = [
    ["NAWM 77", "Peri · Euridice 選段", "bt8KaCIGBEk"],
    ["NAWM 78a", "Monteverdi · Possente spirto", "LUvJmIcp7z0"],
    ["NAWM 78b", "Monteverdi · Tu se' morta", "MIadFIEB1cg"],
    ["NAWM 79", "Monteverdi · Pur ti miro", "v_Se1XVkOiU"],
    ["NAWM 80", "Cesti · Orontea (Intorno all'idol mio)", "WjrNblgYAhw"],
  ];

  pieces.forEach(([nawm, desc, vid], i) => {
    const y = 1.1 + i * 0.85;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 9.2, h: 0.75, fill: { color: C.slate }, rectRadius: 0.08 });
    s.addText(nawm, { x: 0.55, y: y + 0.05, w: 2.0, h: 0.35, fontSize: 20, bold: true, color: C.gold, fontFace: "Georgia" });
    s.addText(desc, { x: 0.55, y: y + 0.38, w: 5.5, h: 0.3, fontSize: 18, color: C.sand, fontFace: "Calibri", valign: "top" });
    s.addText("YouTube", {
      x: 7.5, y: y + 0.15, w: 1.8, h: 0.4, fontSize: 18, bold: true, color: C.gold, fontFace: "Calibri", align: "center", valign: "top",
      hyperlink: { url: `https://www.youtube.com/watch?v=${vid}` },
    });
  });
}

// ── Generate ────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "Ch14_Opera.pptx" })
  .then(() => console.log("■ Ch14_Opera.pptx created — 25 slides"))
  .catch(err => console.error(err));
