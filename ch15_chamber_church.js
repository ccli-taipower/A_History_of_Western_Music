const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 15: Music for Chamber and Church in the Early Seventeenth Century";
pres.author = "A History of Western Music, 10th ed.";

const C = {
  darkBg:   "1A2030",
  gold:     "C89440",
  cream:    "F5F0E0",
  navy:     "1A3050",
  teal:     "2A6A6A",
  darkText: "1A2030",
  lightText:"F5F0E0",
  sand:     "E8D8A8",
  slate:    "1A2838",
  steel:    "5A7A8A",
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
  s.addText("CHAPTER 15", { x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6 });
  s.addText("MUSIC FOR CHAMBER AND\nCHURCH IN THE EARLY\nSEVENTEENTH CENTURY", { x: 0.3, y: 1.4, w: 9.4, h: 2.2, fontSize: 28, color: C.lightText, bold: true, align: "center", fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.75, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Concertato Madrigal · Oratorio · Schütz · Frescobaldi · Violin Sonata", { x: 0.4, y: 3.9, w: 9.2, h: 0.4, fontSize: 12, color: C.sand, align: "center", fontFace: "Georgia" });
  s.addText("Textbook pp. 317–338", { x: 0.5, y: 4.8, w: 9, h: 0.3, fontSize: 11, color: C.gold, align: "center", fontFace: "Calibri" });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);
  s.addText("本章概覽 Chapter Overview", { x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.navy, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });
  const sections = [
    ["🎤", "Italian Vocal Chamber", "Concertato madrigal · Monteverdi (NAWM 82) · Cantata 興起"],
    ["⛪", "Catholic Sacred Music", "Sacred concerto · Grandi (NAWM 83) · Carissimi Jephte (NAWM 84)"],
    ["✝", "Lutheran Church Music", "Schütz · Saul (NAWM 85) · Kleine geistliche Concerte (NAWM 86)"],
    ["🕍", "Jewish Music", "Salamone Rossi · 會堂音樂與記譜"],
    ["⌨", "Keyboard: Frescobaldi", "Toccata No. 3 (NAWM 87) · Organ Mass · Stylus fantasticus"],
    ["🎻", "Violin: Marini", "Sonata per il violino (NAWM 88) · 小提琴的獨立"],
  ];
  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.teal }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 11, color: C.teal, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 3 · Italian Vocal Chamber Music ────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("義大利世俗聲樂", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Concertato Madrigal · Basso Ostinato · Cantata", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎤 Concertato Madrigal", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 從 a cappella 牧歌 → 有 basso continuo 伴奏\n• Monteverdi 第 5–8 部牧歌集的演變\n• 少數聲部 + 器樂 ritornello + continuo\n• 兼具室內樂親密感與劇場表現力\n\n🎵 NAWM 82: Monteverdi · Ohimè ch'io cado\n• 第七部牧歌集 (Concerto, 1619)\n• 二聲部 + basso continuo\n• 表現跌倒、痛苦——半音下行\n• 示範 concertato madrigal 的新語彙\n\n📊 Basso Ostinato（固定低音）\n• 重複 bass pattern 上方旋律自由變化\n• Ciaccona（恰空）/ Passacaglia（帕薩卡利亞）\n• Romanesca · Ruggiero 等固定 bass 模式\n• → Purcell · Pachelbel · Bach 的重要技法", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("📜 Cantata 清唱曲的興起", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Cantata = 「被唱的」(cantare)\n  對比 Sonata = 「被奏的」(sonare)\n• 1620 年代興起——義大利室內聲樂的主流\n\n📊 早期形式\n• Strophic variation（分節變奏）\n• 或 recitative + aria 交替\n• 通常獨唱 + continuo\n• 長度 5–15 分鐘\n• 為貴族或學院私人場合演唱\n\n🌟 重要作曲家\n• Luigi Rossi (1597–1653)\n• Giacomo Carissimi (1605–1674)\n• Barbara Strozzi (1619–1677)\n  — 最多產的 17 世紀女性作曲家\n  — 8 冊出版物——幾乎全為世俗聲樂\n  — Cantata · Arietta · Duet\n\n💡 Cantata 到世紀末成為：\n• 義大利最精緻的聲樂形式\n• 作曲家的能力檢驗標準\n• → A. Scarlatti · Handel · Bach 繼承", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 4 · Catholic Sacred Music ──────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);
  s.addText("天主教宗教音樂", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.navy, fontFace: "Georgia", align: "center" });
  s.addText("Sacred Concerto · Grandi (NAWM 83) · Carissimi Jephte (NAWM 84)", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 12, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.teal } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "E0ECE8" }, rounding: true });
  s.addText("⛪ Sacred Concerto 宗教協奏曲", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.navy, fontFace: "Georgia" });
  s.addText("• 宗教歌詞 + concertato 技法\n• 聲樂 + 器樂 + continuo\n• 從 G. Gabrieli In ecclesiis 延伸\n• 1–多聲部皆可 · 規模差異大\n\n🌟 Alessandro Grandi (ca. 1586–1630)\n• 威尼斯 St. Mark's 副樂長（Monteverdi 之下）\n• O quam tu pulchra es (NAWM 83)\n  — 獨唱 + continuo\n  — 雅歌文字 · 旋律溫柔\n  — Recitative + aria 段落交替\n  — 小型 sacred concerto 典範\n\n📊 Stile antico vs. Stile moderno\n• 教會同時保留兩種風格\n• Palestrina 風格繼續用於正式禮儀\n• 新 concertato 風格用於節慶/特殊場合\n• 兩者共存——作曲家需掌握兩套技術", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "E0ECE8" }, rounding: true });
  s.addText("🎭 Oratorio 神劇的興起", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.navy, fontFace: "Georgia" });
  s.addText("• 宗教題材 · 有戲劇情節 · 但不上舞台\n• 源於 Filippo Neri 的 Oratory（祈禱會）\n• Narration + dialogue + chorus\n• 不用佈景、服裝、動作\n\n🌟 Giacomo Carissimi (1605–1674)\n• 羅馬 Collegio Germanico 樂長\n• 拉丁神劇大師\n\n🎵 Jephte (NAWM 84)\n• 舊約士師記：Jephtha 戰勝後須獻祭女兒\n• Historicus（敘事者）+ 角色 + 合唱\n• 「Plorate, colles」——女兒的哀歌\n  — 女聲獨唱 · 悲痛的下行半音\n  — 合唱呼應「Plorate」——催淚\n• 不協和 + 半音 = 巴洛克情感力量\n\n💡 Oratorio → Handel Messiah · Bach Passion\n• 成為巴洛克大型宗教音樂的核心類型", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 5 · Schütz & Lutheran Music ────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("許茨與路德宗音樂", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Heinrich Schütz · Saul (NAWM 85) · Kleine geistliche Concerte (NAWM 86)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 11, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🌟 Heinrich Schütz (1585–1672)", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 17 世紀最重要的德國作曲家\n• 橫跨義大利新風格與德國路德宗傳統的橋樑\n\n📖 生平\n• 1609–12/1628 兩度赴威尼斯學習\n  — 師事 G. Gabrieli → 後師事 Monteverdi\n• 1617 起任 Dresden 宮廷樂長——直至逝世\n• 三十年戰爭期間極力維持音樂生活\n\n📊 主要作品集\n• Psalmen Davids (1619) · 威尼斯 polychoral 風格\n• Cantiones sacrae (1625) · 拉丁經文歌\n• Symphoniae sacrae I–III (1629/47/50)\n• Kleine geistliche Concerte I–II (1636/39)\n  — 戰時縮編：1–5 聲 + continuo\n• Musikalische Exequien (1636) · 德文安魂曲\n• 3 部 Passion（晚年無伴奏）", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎵 NAWM 85 & 86", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("NAWM 85 · Saul, Saul, was verfolgst du mich?\n• Symphoniae sacrae III (1650)\n• 使徒行傳 9:4 ——「掃羅，你為何逼迫我？」\n• 大型 concertato：6 聲部 + 2 小提琴 + continuo\n  + 補充合唱 (cappella)\n• 耶穌呼喊「Saul」——echo 效果\n• 力度從 fortissimo → pianissimo\n  — 聲音漸遠 = 神聖聲音的消逝\n• 融合威尼斯 polychoral + 劇場戲劇性\n\nNAWM 86 · O lieber Herre Gott\n• Kleine geistliche Concerte (1636)\n• 戰爭時期——樂團縮編\n• 獨唱 + continuo 的小型 sacred concerto\n• 以路德聖歌旋律為基礎\n• 簡單但深情——戰時的信仰表達\n\n💡 Schütz 為 Bach 鋪路\n• 義大利技法 + 德國虔誠 = 路德宗音樂的典範", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 6 · Frescobaldi & Keyboard ─────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);
  s.addText("弗雷斯科巴爾迪與鍵盤音樂", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.navy, fontFace: "Georgia", align: "center" });
  s.addText("Frescobaldi · Toccata (NAWM 87) · Organ Mass · Stylus fantasticus", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 12, color: C.teal, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.teal } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "E0ECE8" }, rounding: true });
  s.addText("🌟 Girolamo Frescobaldi (1583–1643)", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 11, bold: true, color: C.navy, fontFace: "Georgia" });
  s.addText("• 巴洛克初期最偉大的鍵盤作曲家\n• 1608–1628/1634–43 任 St. Peter's (Roma) 管風琴師\n• 傳聞首場演出吸引 3 萬名聽眾（誇張但象徵其名望）\n\n📊 主要作品\n• Toccate e partite d'intavolatura, libro I (1615)\n  — 12 首 Toccata · 序言說明彈性速度\n• Ricercar · Canzona · Capriccio\n• Fiori musicali (1635)\n  — 三套管風琴彌撒\n  — Bach 親手抄寫此譜\n\n🎵 NAWM 87: Toccata No. 3\n• 多段式 · 各段不同 tempo 與織度\n• 自由段 vs. 模仿段交替\n• Frescobaldi 自己說：\n  「開頭要慢 · 依表情調整速度」\n• 展現 stylus fantasticus（幻想風格）", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "E0ECE8" }, rounding: true });
  s.addText("🎻 Marini & Violin Sonata", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.navy, fontFace: "Georgia" });
  s.addText("小提琴：17 世紀獨奏之王\n\n📊 從合奏樂器到獨奏明星\n• 16 世紀：小提琴主要用於舞蹈伴奏\n• 1600 後：進入教堂與室內——地位躍升\n• 義大利製琴師：Amati → Stradivari → Guarneri\n\n🌟 Biagio Marini (1594–1663)\n• 小提琴家 · 作曲家\n• 曾任 St. Mark's 樂手（Monteverdi 的樂團）\n\n🎵 NAWM 88: Sonata IV per il violino (Op. 8)\n• Violin + continuo\n• 多段式 · 各段不同速度與風格\n• 使用 double stops（雙音）\n  — 為小提琴首批技巧性寫作\n• 有 affetti（情感效果段落）\n• 展示小提琴獨有的音色語彙\n\n💡 意義\n• 小提琴 sonata 成為 17–18 世紀核心類型\n• → Corelli · Vivaldi · Tartini · Bach · Mozart", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 7 · Timeline ───────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.teal); bottomBar(s, C.teal);
  s.addText("時間軸 · Timeline", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.navy, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 0.82, w: 5, h: 0.04, fill: { color: C.teal } });
  const events = [
    ["1607", "Monteverdi · L'Orfeo · Schütz 赴 Venice"],
    ["1615", "Frescobaldi · Toccate libro I"],
    ["1617", "Schütz 任 Dresden 樂長"],
    ["1618–48", "三十年戰爭"],
    ["1619", "Monteverdi · Seventh Book (Concerto)"],
    ["1623", "Rossi · Hashirim asher lishlomo"],
    ["1626", "Schütz · Cantiones sacrae"],
    ["1629", "Schütz · Symphoniae sacrae I"],
    ["1635", "Frescobaldi · Fiori musicali"],
    ["1636", "Schütz · Kleine geistliche Concerte I"],
    ["ca. 1640", "Carissimi · Jephte"],
    ["1643", "Frescobaldi · Monteverdi 逝世"],
    ["1650", "Schütz · Symphoniae sacrae III"],
    ["1655", "Marini · Op. 8 Violin Sonatas"],
  ];
  events.forEach(([date, desc], i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 0.3 + col * 4.8;
    const y = 1.0 + row * 0.55;
    s.addShape(pres.ShapeType.rect, { x, y, w: 1.1, h: 0.44, fill: { color: C.teal } });
    s.addText(date, { x: x + 0.05, y: y + 0.06, w: 1.0, h: 0.32, fontSize: 9, bold: true, color: C.lightText, align: "center", fontFace: "Georgia" });
    s.addText(desc, { x: x + 1.2, y, w: 3.55, h: 0.44, fontSize: 8, color: C.darkText, fontFace: "Calibri", valign: "middle" });
  });
}

// ── SLIDE 8 · Key Terms & Listening ──────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("關鍵詞彙 · 延伸閱讀", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Key Terms & Further Reading", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🔑 Key Terms", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• concertato madrigal · concerted\n• basso ostinato · ciaccona · passacaglia\n• romanesca · ruggiero · folia\n• cantata · strophic variation\n• sacred concerto · grand concerto\n• oratorio · historicus · chorus\n• Kleine geistliche Concerte\n• Symphoniae sacrae\n• toccata · ricercar · canzona · capriccio\n• stylus fantasticus · tempo rubato\n• Fiori musicali · organ Mass\n• sonata · violin sonata\n• double stops · affetti\n• Barbara Strozzi · Salamone Rossi", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("📚 Further Reading & 🎧 Listening", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Silbiger. Frescobaldi (2004)\n• Moser/Pfatteicher. Heinrich Schütz (1959)\n• Smither. A History of the Oratorio (1977)\n• Allsop. The Italian Trio Sonata (1992)\n\n🎧 NAWM 精選聆聽 (YouTube)\n• 81 · Caccini · Sfogava con le stelle  youtu.be/gVcc4ZNBLPw\n• 82 · Monteverdi · Ohimè ch'io cado  youtu.be/OOpUglExpoA\n• 83 · Grandi · O quam tu pulchra es  youtu.be/G-u-bieKf24\n• 84 · Carissimi · Jephte: Plorate  youtu.be/aEk9vLzCPBw\n• 85 · Schütz · Saul, Saul  youtu.be/vTiMOsMsv2I\n• 86 · Schütz · O lieber Herre Gott  youtu.be/HX9JhVVC-l8\n• 87 · Frescobaldi · Toccata No. 3  youtu.be/PV4VRPIwbYw\n• 88 · Marini · Sonata IV  youtu.be/DsaYKQI6elQ", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

pres.writeFile({ fileName: "Ch15_Chamber_Church.pptx" })
  .then(fn => console.log(`✅ ${fn} created successfully`))
  .catch(err => console.error("❌ Error:", err));
