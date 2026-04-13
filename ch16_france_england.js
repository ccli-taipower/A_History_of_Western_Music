const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 16: France, England, Spain, the New World, and Russia";
pres.author = "A History of Western Music, 10th ed.";

const C = {
  darkBg:   "1A1A2E",
  gold:     "C8A840",
  cream:    "F5F0E0",
  royal:    "2A2A6A",
  wine:     "6A2040",
  darkText: "1A1A2E",
  lightText:"F5F0E0",
  sand:     "E8D8A8",
  slate:    "2A2A3A",
  amber:    "D4A030",
  plum:     "4A1A3A",
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
  s.addText("CHAPTER 16", { x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6 });
  s.addText("FRANCE, ENGLAND, SPAIN,\nTHE NEW WORLD, AND RUSSIA", { x: 0.3, y: 1.5, w: 9.4, h: 2.0, fontSize: 30, color: C.lightText, bold: true, align: "center", fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.65, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Louis XIV · Lully · Purcell · Versailles · National Styles", { x: 0.4, y: 3.8, w: 9.2, h: 0.4, fontSize: 13, color: C.sand, align: "center", fontFace: "Georgia" });
  s.addText("Textbook pp. 339–370", { x: 0.5, y: 4.8, w: 9, h: 0.3, fontSize: 11, color: C.gold, align: "center", fontFace: "Calibri" });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.royal); bottomBar(s, C.royal);
  s.addText("本章概覽 Chapter Overview", { x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.royal, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });
  const sections = [
    ["🏰", "France 法國", "Louis XIV · 宮廷芭蕾 · Versailles · Lully 的音樂獨裁"],
    ["🎭", "French Opera 法國歌劇", "Tragédie en musique · 法式序曲 · divertissement · 宣敘調"],
    ["🎹", "Lute & Keyboard 魯特與鍵盤", "Gaultier · Jacquet de la Guerre · suite · agréments"],
    ["🇬🇧", "England 英格蘭", "Masque · Purcell · Dido and Aeneas · 公共音樂會之誕生"],
    ["🌎", "Spain & New World 西班牙與新世界", "Zarzuela · villancico · Torrejón · Padilla"],
    ["🇷🇺", "Russia 俄羅斯", "Eastern Orthodox · znamenny chant · kontsert · kanty"],
  ];
  sections.forEach(([icon, title, desc], i) => {
    const y = 1.0 + i * 0.72;
    s.addShape(pres.ShapeType.rect, { x: 0.45, y, w: 9.1, h: 0.62, fill: { color: i % 2 === 0 ? "E8E0D0" : "DED6C6" }, rounding: true });
    s.addText(icon, { x: 0.55, y, w: 0.5, h: 0.62, fontSize: 20, align: "center" });
    s.addText(title, { x: 1.1, y, w: 2.6, h: 0.62, fontSize: 13, bold: true, color: C.royal, fontFace: "Georgia", valign: "middle" });
    s.addText(desc, { x: 3.8, y, w: 5.65, h: 0.62, fontSize: 10.5, color: C.darkText, fontFace: "Calibri", valign: "middle" });
  });
}

// ── SLIDE 3 · France: Louis XIV & Versailles ────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("🏰 法國：路易十四與凡爾賽 France: Louis XIV & Versailles", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("The Sun King 太陽王 (r. 1643–1715)", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 自比太陽神 Apollo；藝術即權力工具\n• 建立皇家學院：舞蹈 (1661)、文學 (1663)、科學 (1669)、歌劇 (1669)、建築 (1671)\n• 凡爾賽宮 Versailles：政治控制 + 藝術展演\n  — 貴族長住宮中 → 受王權掌控\n  — 花園幾何對稱 = 中央集權秩序\n• 宮廷舞蹈文化：danses à deux / 宮廷芭蕾\n  — Pierre Beauchamp：五個基本腳位\n  — Feuillet (1700)：首套完整舞譜系統\n\nDance reinforced order, refinement, and\nsubordination — the very image of the\ncentralized monarchy.", {
    x: 0.45, y: 1.45, w: 4.3, h: 3.6, fontSize: 9, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 3,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Music at Court 宮廷音樂", { x: 5.2, y: 1.08, w: 4.4, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("三大音樂機構 (150–200 musicians):\n\n① The Chapel 皇家禮拜堂\n   — 歌手、管風琴師、宗教儀式器樂手\n\n② The Chamber 室內\n   — 獨唱、弦樂、魯特琴、大鍵琴、長笛\n   — 室內娛樂用途\n\n③ The Great Stable 大馬廄\n   — 管樂與銅管 → 軍事 / 戶外典禮\n   — 催生現代雙簧管 (oboe)\n   — Jean Hotteterre 家族：木管革新\n\n弦樂團 String Orchestra:\n• Vingt-quatre Violons du Roi (24 把)\n• Petits Violons (18 把，1648)\n→ \"orchestra\" 一詞 1670s 開始使用", {
    x: 5.25, y: 1.45, w: 4.3, h: 3.6, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 4 · Lully & French Opera ──────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("🎭 呂利與法國歌劇 Lully & French Opera", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Jean-Baptiste Lully (1632–1687)", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 義大利佛羅倫斯出生 → 14 歲赴巴黎\n• 以舞藝引起路易十四注意\n• 1661 任器樂總監，掌管 Petits Violons\n• 建立統一弓法、禁止自由裝飾 → 現代管弦樂團紀律\n• 1672 取得皇家特許：法國唯一歌劇演出權\n  → 創立 Académie Royale de Musique\n• 與劇作家 Philippe Quinault 合作\n\nTragédie en musique 音樂悲劇：\n• 五幕劇；古典神話 / 騎士傳奇題材\n• 結合戲劇、音樂、芭蕾\n• Quinault 劇本：暗頌路易 + 道德寓言\n• Lully 死於指揮 Te Deum 時杖擊足部\n  → 壞疽感染身亡", {
    x: 0.45, y: 1.45, w: 4.3, h: 3.6, fontSize: 9, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Armide (1686) — NAWM 85", { x: 5.2, y: 1.08, w: 4.4, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("French Overture 法式序曲 (NAWM 85a):\n• 兩段式：慢板（附點節奏 + 齊奏）→ 快板（模仿式）\n• 快板末尾常回歸慢板速度\n→ 此形式流傳超過一世紀\n\nDivertissement 穿插段 (NAWM 85b):\n• Act II 內：仙女、牧人歌舞\n• Solo airs + 合唱 + 器樂舞曲\n• Airs：旋律優美、節拍規整、少花腔\n  — 與義大利 aria 截然不同\n\nArmide 獨白 (Act II, Sc. 5 · NAWM 85c):\n• 女巫 Armide 手持匕首面對沉睡的 Renaud\n• 宣敘調自由轉換 2/4 和 3/4 拍\n• récitatif → air \"Venez, venez\" 三拍\n• Tirades：附點音型 + 快速音階\n→ 音樂服務戲劇，不為歌唱炫技", {
    x: 5.25, y: 1.45, w: 4.3, h: 3.6, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 5 · French Performance Practice & Church Music ────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("🎶 法國演奏風格與教會音樂 French Style & Church Music", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 21, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Performance Practice 演奏慣例", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Notes inégales 不等分音符：\n• 記譜為等值八分音符 → 演奏為長短交替\n  （類似搖擺三連音效果）\n• 被視為 expression 與 elegance 的表現\n\nOverdotting 加附點：\n• 附點音符比記譜更長 → 短音更短\n• 強化節奏尖銳感\n\nAgréments 裝飾音：\n• 固定裝飾：trills, appoggiatura, mordent\n• D'Anglebert (1689) 裝飾記號表\n  — 最完整的巴洛克裝飾音系統\n• 強調 cadence 與重拍\n• 法式 vs 義式裝飾：簡短固定 vs 即興華麗\n\n\"Nature and expression are the two\nprimary qualities of music\" — Lecerf", {
    x: 0.45, y: 1.45, w: 4.3, h: 3.6, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Church Music & Song", { x: 5.2, y: 1.08, w: 4.4, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Motets 經文歌：\n• Petit motet：少數歌手 + 通奏低音\n• Grand motet：獨唱 + 大/小合唱 + 管弦樂\n  — 多段、多速度、多風格\n• Lully Te Deum (1677 · NAWM 86)\n  — 150人演出；獨唱 + 合唱 + 弦樂 + 銅管\n\nMarc-Antoine Charpentier (1643–1704):\n• 赴羅馬學習 Carissimi 風格\n• 融合義大利不協和 + 法國高雅\n• Le reniement de Saint Pierre (NAWM 87)\n  — 彼得三次否認主：經文敘事體\n  — 結尾合唱：不協和掛留音 → 催人淚下\n\nMichel-Richard de Lalande (1657–1726):\n• 路易十四晚年寵臣\n• 70+ grand motets：風格多變\n\nSong: air sérieux / air à boire\n• Michel Lambert (ca. 1610–1696)：\n  Lully 岳父，air 先驅", {
    x: 5.25, y: 1.45, w: 4.3, h: 3.6, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 6 · Lute, Keyboard & Dance Suites ─────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("🎹 魯特琴、鍵盤與舞曲組曲 Lute, Keyboard & Suite", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 21, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Lute & Harpsichord 魯特琴與大鍵琴", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Denis Gaultier (1603–1672)：\n• La rhétorique des dieux 魯特琴集\n• La coquette virtuose (NAWM 88) — courante\n• Style luthé / style brisé 碎和弦風格：\n  — 分散和弦替代齊奏 → 暗示多聲部\n  — 深刻影響大鍵琴音樂\n\nClavecin 大鍵琴 (= harpsichord)：\n• 17世紀取代魯特琴成為主奏獨奏樂器\n• 重要 clavecinists 鍵盤家：\n  — Chambonnières (1601–1672)\n  — D'Anglebert (1629–1691)\n  — Jacquet de la Guerre (1665–1729)\n  — François Couperin (1668–1733)\n\nAgréments 在鍵盤的角色：\n• 撥弦 → 快速衰減 → 裝飾音延長共鳴\n• 模擬持續音效果\n• 法國鍵盤風格的精髓", {
    x: 0.45, y: 1.45, w: 4.3, h: 3.6, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Suite & Dance Forms 組曲與舞曲", { x: 5.2, y: 1.08, w: 4.4, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Jacquet de la Guerre Suite No. 3 in Am (NAWM 89):\n\n標準四舞曲 + 附加舞曲：\n┌──────────────────────────────────┐\n│ 舞曲       拍子    性格     來源      │\n├──────────────────────────────────┤\n│ Allemande  4/4    莊重     德國      │\n│ Courante   3/2    機智     法國      │\n│ Sarabande  3/4    高貴     西/新世界  │\n│ Gigue      6/8    活潑     英/愛     │\n└──────────────────────────────────┘\n\n(NAWM 89a) Unmeasured prelude 非量化前奏\n(89b) Allemande — style luthé 碎和弦\n(89c) Courante — 2/3拍交替\n(89d) Sarabande — 強調第二拍\n(89e) Gigue — 模仿式開頭\n(89f) Chaconne — rondeau 形式\n(89g) Gavotte — 半小節弱起\n(89h) Minuet — 優雅三拍\n\nBinary form 二段式：I→V || V→I", {
    x: 5.25, y: 1.45, w: 4.3, h: 3.6, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1,
  });
}

// ── SLIDE 7 · England: Purcell & Dido ───────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("🇬🇧 英格蘭：Purcell 與 Dido and Aeneas", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("England: Politics & Music", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("政治動盪與音樂：\n• Charles I → 內戰 (1642–49) → 清教政權\n• Cromwell 禁止公開戲劇 (但非私人音樂)\n• 1660 王政復辟 → Charles II（流亡法國，\n  帶回法國音樂品味）\n• Glorious Revolution (1689) → 君主立憲\n• 王室財力遠不及法國 → 公共音樂發展\n\nMasque 假面劇：\n• Henry VIII 以來的宮廷傳統\n• 音樂 + 舞蹈 + 佈景 + 台詞\n• 非單一作曲家的統一歌劇\n• Cupid and Death (1653)：唯一存世完整\n  假面劇，Locke + Gibbons 音樂\n\n兩部全唱英語歌劇：\n• John Blow, Venus and Adonis (ca. 1683)\n• Henry Purcell, Dido and Aeneas (ca. 1688)", {
    x: 0.45, y: 1.45, w: 4.3, h: 3.6, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Henry Purcell (1659–1695)", { x: 5.2, y: 1.08, w: 4.4, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("\"The British Orpheus\" 英國的奧菲歐：\n• Chapel Royal 唱詩班出身\n• Westminster Abbey 管風琴師 (1679)\n• Chapel Royal 管風琴師 (1682)\n• 36歲英年早逝，葬於 Westminster Abbey\n\nDido and Aeneas (ca. 1688 · NAWM 90)：\n• 女子寄宿學校私人演出\n• 微型歌劇傑作：4角色，約1小時\n• 融合義法英三風格：\n  — French overture + homophonic choruses\n  — Italian ground-bass arias (3首)\n  — English tunefulness + word-painting\n\n(90a) Thy hand, Belinda — 宣敘調\n  — 半音下行 → Dido 瀕死的哀傷\n(90b) When I am laid in earth — Dido's Lament\n  — 下行四度 chromatic ground bass × 11 次\n  — 歌劇史上最動人的詠嘆調之一\n(90c) With drooping wings — 結尾合唱\n  — 仿 Blow Venus and Adonis 的終曲", {
    x: 5.25, y: 1.45, w: 4.3, h: 3.6, fontSize: 7.8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 8 · English Musical Life ──────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("🎵 英國音樂生活 English Musical Life", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Vocal & Church Music", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Purcell 的戲劇音樂 (1690s)：\n• Semi-opera / dramatic opera：\n  — 口語劇 + 四個以上 masque 段落\n  — The Fairy Queen (1692)：改編莎翁\n  — 配樂 (incidental music) 近50齣戲\n\n聲樂音樂 Vocal Music：\n• Ode for St. Cecilia's Day (1692)\n  — 受法式 grand motet 影響\n  — 直接啟發 Handel 清唱劇\n• Catch：輪唱式幽默歌曲（常粗俗）\n  — 全男性社交娛樂\n\nAnglican Church Music 聖公會：\n• 清教期間廢除 → 復辟後恢復\n• Charles II 偏好法式風格 → solo + 管弦伴奏\n• Anthem + Service 為主要體裁\n• Coronation music 加冕音樂尤其精緻", {
    x: 0.45, y: 1.45, w: 4.3, h: 3.6, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Instrumental Music & Public Concert", { x: 5.2, y: 1.08, w: 4.4, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("Consort Music 合奏音樂：\n• Viol consort：In Nomines + fantasias\n  — Jenkins、Locke、Purcell (ca. 1680)\n  — 英格蘭此類體裁的最後重要作品\n• Purcell 亦寫大鍵琴舞曲、trio sonatas\n  — 受義大利影響\n\n社交舞蹈：\n• John Playford, The English Dancing\n  Master (1651)：鄉村舞曲集\n  — 最早印刷民間曲調集之一\n  — 暢銷至1728年\n\n公共音樂會的誕生 The Public Concert：\n• 英國獨創發明！1672年12月倫敦首創\n  — John Banister：\"George Tavern\"\n  — 每週收費公演\n• 很快傳播：\n  — 1725 巴黎 Concert Spirituel\n  — 1730s 北美殖民地\n→ 現代音樂會制度的起源", {
    x: 5.25, y: 1.45, w: 4.3, h: 3.6, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 9 · Spain, New World & Russia ─────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("🌎 西班牙、新世界與俄羅斯", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Spain & the New World 西班牙與新世界", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("西班牙黃金時代：Cervantes · Lope de Vega · Velázquez\n\nOpera & Zarzuela 歌劇與查蘇艾拉：\n• Juan Hidalgo (1614–1685) + Calderón de la Barca\n  — 首部 zarzuela：輕鬆神話劇\n  — 歌唱 + 對白交替（非全唱）\n• Torrejón y Velasco, La púrpura de la rosa\n  (1701 · NAWM 91)\n  — 新世界首部歌劇！（利馬，秘魯）\n  — 對話非宣敘調 → 使用 strophic song\n\nChurch Music 教會音樂：\n• Villancico 聖誕/復活節合唱曲\n  — vernacular（方言）非拉丁文\n  — estribillo 疊句 + coplas 詩節\n• Padilla, Albricias pastores (NAWM 92)\n  — Puebla, Mexico 大教堂唱詩班長\n  — 雙合唱 antiphonal 對唱\n\nInstrumental: organ tiento + guitar", {
    x: 0.45, y: 1.45, w: 4.3, h: 3.6, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Russia 俄羅斯", { x: 5.2, y: 1.08, w: 4.4, h: 0.32, fontSize: 13, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("俄羅斯正教會 Russian Orthodox Church：\n• 拜占庭傳統 → 禁止樂器\n• Znamenny chant 標記聖歌：\n  — 單聲部、專用記譜法\n• 16世紀：三聲部 polyphony 加入\n• 1600年：沙皇與教會支持合唱團\n  → 偉大的無伴奏合唱傳統之始\n\n西方影響 (1650s–)：\n• Patriarch Nikon (1652)：開放西方風格\n• 五線譜 + 西歐和聲觀念傳入\n• Nikolay Diletsky (ca. 1630–1680+)\n  — Idea grammatikii musikiyskoy (1679)\n  — 最早的完整五度圈描述之一\n\nKontsert 聲樂協奏曲：\n• 改自 sacred concerto，純聲樂（無樂器）\n• Vasiliy Titov (ca. 1650–1715)：\n  — Psaltir' rifmovannaya (1686)\n  — Beznevestnaya Devo 複合唱\n\nKanty 三聲部頌歌：簡單、平行三度\n\nPeter the Great (r. 1682–1725)：\n• 強力西化 → 彼得堡 (1703)", {
    x: 5.25, y: 1.45, w: 4.3, h: 3.6, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1,
  });
}

// ── SLIDE 10 · Key Terms & Listening ────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("📝 關鍵詞彙 · 延伸閱讀 · 聆聽 Key Terms & Listening", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 20, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.025, fill: { color: C.wine } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.0, w: 4.55, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("Key Terms 關鍵詞彙", { x: 0.4, y: 1.08, w: 4.35, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• court ballet / ballet de cour\n• tragédie en musique / tragédie lyrique\n• French overture · ouverture\n• divertissement · entrée · air\n• récitatif simple / récitatif mesuré\n• tirade · notes inégales · overdotting\n• agréments · tremblement · pincé\n• style luthé / style brisé\n• clavecin · clavecinists\n• suite · binary form\n• allemande · courante · sarabande · gigue\n• gavotte · minuet · chaconne · rondeau\n• unmeasured prelude\n• petit motet · grand motet\n• masque · semi-opera · catch\n• zarzuela · villancico · tiento\n• kontsert · kanty · znamenny chant\n• Vingt-quatre Violons du Roi\n• The English Dancing Master", {
    x: 0.5, y: 1.42, w: 4.35, h: 3.6, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.0, w: 4.6, h: 4.3, fill: { color: C.slate }, rounding: true });
  s.addText("📚 Further Reading & 🎧 Listening", { x: 5.25, y: 1.08, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Isherwood. Music in the Service of the King (1973)\n• Anthony, James. French Baroque Music (1997)\n• Holman. Henry Purcell (1994)\n• Price, Curtis. Henry Purcell and the London Stage (1984)\n\n🎧 NAWM 精選聆聽 (YouTube)\n• 85 · Lully · Armide (overture+scene)  youtu.be/HWIaanFmhYY\n• 86 · Lully · Te Deum  youtu.be/QARAbmTbArU\n• 87 · Charpentier · Le reniement de St Pierre  youtu.be/S3diWOcgSOI\n• 88 · Gaultier · La coquette virtuose  youtu.be/rB2iH0N1Np4\n• 89 · Jacquet de la Guerre · Suite No. 3  youtu.be/CD9DYeq4_FQ\n• 90 · Purcell · Dido: When I am laid in earth  youtu.be/uGQq3HcOB0Y\n• 91 · Torrejón · La púrpura de la rosa  youtu.be/ZJ_OxrK177A\n• 92 · Padilla · Albricias pastores  youtu.be/CNvMJxwt4FQ", {
    x: 5.3, y: 1.42, w: 4.35, h: 3.6, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

pres.writeFile({ fileName: "Ch16_France_England.pptx" })
  .then(fn => console.log(`✅ ${fn} created successfully`))
  .catch(err => console.error("❌ Error:", err));
