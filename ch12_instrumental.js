const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 12: The Rise of Instrumental Music";
pres.author = "A History of Western Music, 10th ed.";

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
function topBar(s, c) { s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.12, fill: { color: c || C.gold } }); }
function bottomBar(s, c) { s.addShape(pres.ShapeType.rect, { x: 0, y: 5.5, w: "100%", h: 0.125, fill: { color: c || C.gold } }); }

// ── SLIDE 1 · Title ──────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.15, fill: { color: C.gold } });
  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.47, w: "100%", h: 0.155, fill: { color: C.gold } });
  s.addText("A HISTORY OF WESTERN MUSIC · TENTH EDITION", { x: 0.5, y: 0.45, w: 9, h: 0.35, fontSize: 11, color: C.sand, charSpacing: 3, align: "center", fontFace: "Georgia" });
  s.addText("CHAPTER 12", { x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6 });
  s.addText("THE RISE OF\nINSTRUMENTAL MUSIC", { x: 0.3, y: 1.5, w: 9.4, h: 2.0, fontSize: 34, color: C.lightText, bold: true, align: "center", fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.65, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Dance · Intabulation · Variation · Canzona · Toccata · Venice", { x: 0.4, y: 3.8, w: 9.2, h: 0.4, fontSize: 13, color: C.sand, align: "center", fontFace: "Georgia" });
  s.addText("Textbook pp. 254–277", { x: 0.5, y: 4.8, w: 9, h: 0.3, fontSize: 11, color: C.gold, align: "center", fontFace: "Calibri" });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);
  s.addText("本章概覽 Chapter Overview", { x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.forest, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });
  const sections = [
    ["🎻", "Instruments & Ensembles", "家族式樂器 · consort · haut/bas · 記譜傳統興起"],
    ["💃", "Dance Music 舞曲", "Pavane · Galliard · Susato Danserye (NAWM 67)"],
    ["🎸", "Intabulation 鍵盤/魯特改編", "Narváez vihuela · Byrd keyboard variations (NAWM 68–69)"],
    ["🏛", "Abstract Genres 抽象樂種", "Ricercar · Canzona · Toccata · Prelude · Fantasia"],
    ["⛪", "Music in Venice 威尼斯", "Cori spezzati · Gabrieli 叔姪 · Sonata pian' e forte (NAWM 70–73)"],
    ["🎼", "Independence 獨立地位", "器樂從附屬→獨立——為巴洛克器樂鋪路"],
  ];
  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.forest }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 11, color: C.olive, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 3 · Instruments & Ensembles ────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("樂器與合奏", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Instruments, Ensembles & Embellishment in the 16th Century", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎻 Instrumental Families 樂器家族", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 同族樂器 soprano→bass——統一音色\n• Consort（英文）= 同族三到七件樂器合奏\n• Whole consort = 同族 · Broken consort = 混合\n\n🎵 常見家族\n• Recorder（直笛）家族\n• Viol（維奧爾琴）家族——有品\n• Violin（小提琴）家族——1530s 出現\n• Cornett + Sackbut（木管號角+長號）\n• Shawm（蕭姆管）家族\n• Keyboard：Organ · Harpsichord · Clavichord\n• Lute/Vihuela——16 世紀最流行的獨奏樂器\n\n📚 重要文獻\n• Virdung, Musica getutscht (1511)\n• Praetorius, Syntagma musicum (1618–20)\n• Agricola, Musica instrumentalis (1529/45)", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎨 Embellishment 裝飾法", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 16 世紀演奏者有義務即興裝飾旋律\n• Passaggi（快速走句）= 分割長音符為短音群\n• Graces（裝飾音）= 顫音、附倚音\n\n📖 裝飾法教本\n• Silvestro Ganassi, Opera intitulata Fontegara (1535)\n  ——直笛裝飾手冊\n• Diego Ortiz, Tratado de glosas (1553)\n  ——viol 即興法\n• Girolamo Dalla Casa, Il vero modo (1584)\n\n💡 為何重要？\n• 器樂如同聲樂——以裝飾展現技藝\n• 寫定 vs. 即興的分界模糊\n• HIP（歷史知情演奏）中的核心議題\n\n🏛 社交舞蹈\n• 宮廷舞蹈為社交要求 · 出版大量舞譜\n• Arbeau, Orchésographie (1589)\n  ——舞步 + 節奏對照圖", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 4 · Dance Music ────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);
  s.addText("舞曲 Dance Music", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.forest, fontFace: "Georgia", align: "center" });
  s.addText("Pavane & Galliard · Susato · Danserye (NAWM 67)", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 13, color: C.olive, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.forest } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "E8E0D0" }, rounding: true });
  s.addText("💃 舞曲配對", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.forest, fontFace: "Georgia" });
  s.addText("• 16 世紀慣例：慢舞 + 快舞成對\n• Pavane（莊重 · 二拍）+ Galliard（活潑 · 三拍）\n• Passamezzo + Saltarello（義大利）\n• Allemande + Courante（法國）\n• 兩首通常共享旋律素材\n\n📊 Pavane 特徵\n• 緩慢 · 莊嚴 · 通常 duple meter\n• 三段式 AABB'CC'\n• 適合遊行或開場\n\n📊 Galliard 特徵\n• 活潑跳躍 · triple meter\n• 常以 pavane 旋律作三拍改編\n• hemiola 效果頻繁（3+3 vs 2+2+2）", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "E8E0D0" }, rounding: true });
  s.addText("🎵 Susato · Danserye (1551)", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.forest, fontFace: "Georgia" });
  s.addText("🌟 Tielman Susato (ca. 1510/15–after 1570)\n• 安特衛普的樂器演奏家、印刷商\n• 出版樂譜數十種——包括 chanson 與器樂曲集\n\n📖 Het derde musyck boexken (1551)\n• Danserye = 六冊舞曲集中的第三冊\n• 為四聲部合奏寫成\n• 包含 Pavane「La dona」+ Galliard (NAWM 67)\n\n🎵 NAWM 67 分析\n• Pavane: 三段 · 同節奏 · 旋律在 superius\n• Galliard: 改用三拍 · 旋律為 pavane 變體\n• 各段重複（||: A :||: B :||: C :||）\n• 清晰、短小、適合業餘合奏\n\n📈 影響\n• 舞曲集成為 16–17 世紀印刷暢銷品\n• 為 Baroque suite（組曲）的前身", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 5 · Intabulation & Variations ─────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("改編與變奏", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Intabulation · Narváez (NAWM 68) · Byrd Variations (NAWM 69)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎸 Intabulation 改編譜", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 將聲樂作品改編為鍵盤或弦撥樂器獨奏譜\n• 使用 tablature（圖譜）記譜法\n  — lute tab = 指法 + 弦位\n  — organ tab = 字母或數字代音\n• 16 世紀最大宗的器樂出版品\n\n🎸 Luis de Narváez (fl. 1526–49)\n• 西班牙 vihuela（維韋拉琴）大師\n• Los seys libros del Delphin (1538)\n  — 六冊維韋拉曲集\n• 包含 Josquin Mille regretz 改編 (NAWM 68)\n\n🎵 NAWM 68 分析\n• 忠於原作四聲部架構\n• 添加大量 passaggi 裝飾——\n  展現 vihuela 的觸弦特色\n• 重要：展示 16 世紀「改編＝再創作」的觀念", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎹 Variations 變奏曲", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 變奏曲是器樂首批真正的「獨立類型」\n• 以既有旋律為基礎 · 反覆但每次變化\n\n三種類型\n1. Cantus firmus variations — 旋律不變 · 對位變化\n2. Melodic paraphrase — 旋律裝飾性改變\n3. Ground bass / Harmonic — 和聲骨架不變 · 上方自由\n\n🌟 William Byrd (ca. 1540–1623)\n• 英國 virginal（小型鍵盤）音樂大師\n• John come kiss me now (NAWM 69)\n  — 以流行歌曲旋律為基礎\n  — 16 段變奏 · 漸進式複雜化\n  — 技巧漸增 · 節奏越來越快\n  — 鍵盤特有的 figuration 展現\n\n📚 The Fitzwilliam Virginal Book\n• 英國最大鍵盤曲手稿集\n• 300 首——Byrd、Bull、Gibbons 等\n• 變奏曲占大比例", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 6 · Abstract Genres ────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);
  s.addText("抽象器樂曲種", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.forest, fontFace: "Georgia", align: "center" });
  s.addText("Ricercar · Fantasia · Canzona · Toccata · Prelude", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 13, color: C.olive, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.forest } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "E8E0D0" }, rounding: true });
  s.addText("📜 模仿式樂種", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.forest, fontFace: "Georgia" });
  s.addText("🎵 Ricercar（利切卡爾）\n• 源自「尋找 (ricercare)」——最初即興性質\n• 16 世紀後轉為嚴謹的模仿對位曲\n• 一個或多個主題 · 各段以新主題進入\n• 功能如 motet 的器樂版\n• 代表：Andrea Gabrieli 12° tuono ricercar (NAWM 71)\n\n🎵 Fantasia\n• 與 ricercar 近似但更自由\n• 速度與織度變化更大\n• 英國：Fancy (viol consort 代表曲種)\n\n🎵 Canzona（坎佐那）\n• 源自 chanson 的器樂改編\n• 特徵：「長-短-短」開頭節奏\n• 多段式 · 各段對比 · 接近後來的 sonata\n• Giovanni Gabrieli · Canzon septimi toni (NAWM 70)", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "E8E0D0" }, rounding: true });
  s.addText("⌨ 即興式樂種", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.forest, fontFace: "Georgia" });
  s.addText("🎵 Toccata（觸技曲）\n• 源自 toccare = 觸碰——鍵盤樂器專屬\n• 即興風格 · 快速音群 · 自由和聲\n• 展現演奏者技藝 · 常有踏板音\n• 早期：Andrea Gabrieli · Claudio Merulo\n• 後期：Frescobaldi 將 toccata 提升至藝術頂峰\n\n🎵 Prelude / Praeambulum\n• 禮拜或演奏前的引奏\n• 建立調性 · 暖手 · 即興性質\n• 為後世 Prelude & Fugue 配對的前身\n\n🎵 Intonazione\n• 威尼斯管風琴家為合唱定音的短曲\n• Andrea & Giovanni Gabrieli 有許多範例\n\n💡 共通趨勢\n• 16 世紀器樂從「模仿聲樂」走向「發展獨特語彙」\n• 到世紀末：canzona → sonata · ricercar → fugue", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 7 · Venice & the Gabrielis ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("威尼斯與加布里埃利", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("St. Mark's · Cori Spezzati · Andrea & Giovanni Gabrieli", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🏛 St. Mark's 聖馬可大教堂", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 威尼斯共和國的國家教堂\n• 建築特殊：對置的雙管風琴閣樓\n  → 啟發「對唱合唱」(cori spezzati)\n\n⛪ Cori spezzati 分置合唱\n• 兩組或多組合唱（+器樂）分置教堂兩端\n• 輪流唱、齊唱、呼應——空間音效\n• Adrian Willaert（1527–62 任樂長）開創傳統\n\n🌟 Andrea Gabrieli (ca. 1532–1585)\n• St. Mark's 管風琴師\n• 器樂：ricercar · intonazione · canzona\n• 聲樂：大型慶典音樂\n• Ricercar del 12° tuono (NAWM 71)\n  — 單主題嚴謹模仿——接近後來的 fugue\n  — 四聲部 · 管風琴", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🌟 Giovanni Gabrieli (ca. 1554/7–1612)", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Andrea 的姪子——繼任 St. Mark's 管風琴師\n• 16 世紀末最重要的威尼斯作曲家\n• 吸引全歐學生（包括 Schütz）\n\n📖 Sacrae symphoniae (1597)\n• 宗教聲樂 + 器樂合集\n• 包含 Canzon septimi toni à 8 (NAWM 70)\n  — 八聲部分為兩組 SATB 合奏\n  — 模仿、呼應、齊奏交替\n  — chanson 開頭節奏（長-短-短）\n\n🎵 Sonata pian' e forte (NAWM 73)\n• 音樂史上首次在譜面標明力度對比\n• 八聲部分兩組——cornett/violin + sackbut\n• piano (p) 與 forte (f) 交替\n• 重要：器樂「音色」首次作為結構元素\n\n🎵 In ecclesiis (NAWM 72)\n• Grand concerto——聲樂+合唱+器樂+管風琴\n• 多層次 concertato 織度——巴洛克的預兆", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 8 · Instrumental Independence ──────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);
  s.addText("器樂的獨立", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.forest, fontFace: "Georgia", align: "center" });
  s.addText("Instrumental Music Gains Independence", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 13, color: C.olive, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.forest } });

  const points = [
    ["📊", "從附屬到獨立", "16 世紀前器樂多為舞蹈伴奏或聲樂改編——1600 年後器樂開始追求自身美學\nBefore 1600: functional / derivative → After 1600: autonomous art music"],
    ["🎻", "樂器特性語彙", "作曲家開始為特定樂器寫作——鍵盤 figuration、弦樂弓法、管樂呼吸——不再照搬聲樂\nIdiomatic writing: composers exploited each instrument's unique capabilities"],
    ["📖", "記譜法完善", "Tablature → Staff notation · 出版樂譜市場成形——器樂首次大量保存\nMore instrumental music was written down, published, and preserved"],
    ["🏛", "威尼斯典範", "Gabrieli 的空間化大型合奏 → 為 Baroque concertato 與 concerto 鋪路\nVenice's polychoral grand concerto foreshadowed Baroque orchestral music"],
    ["🔗", "類型的演化", "canzona → sonata · ricercar → fugue · dance pairs → suite · toccata → prelude\nGenre evolution set the agenda for the next 200 years of instrumental music"],
  ];
  points.forEach(([icon, title, desc], i) => {
    const y = 1.2 + i * 0.85;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.55, h: 0.55, fill: { color: C.forest }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.06, w: 0.55, h: 0.44, fontSize: 18, align: "center" });
    s.addText(title, { x: 1.1, y, w: 8.5, h: 0.28, fontSize: 13, bold: true, color: C.darkText, fontFace: "Georgia" });
    s.addText(desc, { x: 1.1, y: y + 0.28, w: 8.5, h: 0.46, fontSize: 9, color: C.brown, fontFace: "Calibri" });
  });
}

// ── SLIDE 9 · Timeline ───────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);
  s.addText("時間軸 · Timeline", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.forest, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 0.82, w: 5, h: 0.04, fill: { color: C.forest } });
  const events = [
    ["1511", "Virdung · Musica getutscht"],
    ["1529/45", "Agricola · Musica instrumentalis"],
    ["1535", "Ganassi · Fontegara (recorder)"],
    ["1538", "Narváez · Los seys libros del Delphin"],
    ["1549", "Willaert → di Rore 繼任 St. Mark's"],
    ["1551", "Susato · Danserye"],
    ["1553", "Ortiz · Tratado de glosas"],
    ["1584", "Merulo · Toccate d'intavolatura"],
    ["1585", "Andrea Gabrieli 逝世"],
    ["1589", "Arbeau · Orchésographie"],
    ["1593", "Fitzwilliam Virginal Book 開始抄寫"],
    ["1597", "G. Gabrieli · Sacrae symphoniae (Sonata pian' e forte)"],
    ["1612", "G. Gabrieli 逝世"],
    ["1615", "G. Gabrieli · Canzoni et sonate (posthumous)"],
    ["1618–20", "Praetorius · Syntagma musicum"],
  ];
  events.forEach(([date, desc], i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 0.3 + col * 4.8;
    const y = 1.0 + row * 0.52;
    s.addShape(pres.ShapeType.rect, { x, y, w: 1.1, h: 0.42, fill: { color: C.forest } });
    s.addText(date, { x: x + 0.05, y: y + 0.06, w: 1.0, h: 0.3, fontSize: 9, bold: true, color: C.lightText, align: "center", fontFace: "Georgia" });
    s.addText(desc, { x: x + 1.2, y, w: 3.55, h: 0.42, fontSize: 8, color: C.darkText, fontFace: "Calibri", valign: "middle" });
  });
}

// ── SLIDE 10 · Key Terms & Further Reading ───────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("關鍵詞彙 · 延伸閱讀", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Key Terms & Further Reading", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🔑 Key Terms", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• consort · whole consort · broken consort\n• haut / bas instruments\n• tablature · intabulation\n• pavane · galliard · passamezzo · saltarello\n• variation · cantus firmus var. · ground bass\n• passaggi · divisions · diminution\n• ricercar · fantasia · fancy\n• canzona · long-short-short rhythm\n• toccata · prelude · intonazione\n• sonata (earliest usage)\n• cori spezzati · polychoral\n• concertato · grand concerto\n• Sacrae symphoniae\n• Sonata pian' e forte\n• Fitzwilliam Virginal Book\n• Danserye · Orchésographie", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("📚 Further Reading & 🎧 Listening", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Selfridge-Field. Venetian Instrumental Music from Gabrieli to Vivaldi (1975/94)\n• Silbiger, ed. Keyboard Music Before 1700 (2004)\n• Brown, Howard M. Instrumental Music Printed Before 1600\n• Holman. Four and Twenty Fiddlers (1993)\n\n🎧 NAWM 精選聆聽 (YouTube)\n• 67 · Susato · Pavane + Galliard  youtu.be/Ln7ea-5dsoo\n• 68 · Narváez · Mille regretz (vihuela)  youtu.be/INGuCsQtefA\n• 69 · Byrd · John come kiss me now  youtu.be/DD7luwIuM40\n• 70 · G. Gabrieli · Canzon septimi toni  youtu.be/5DgJdrZcYoA\n• 71 · A. Gabrieli · Ricercar 12° tuono  youtu.be/7Nhe7m8Q_9s\n• 72 · G. Gabrieli · In ecclesiis  youtu.be/1kZ_Ld9-wH4\n• 73 · G. Gabrieli · Sonata pian' e forte  youtu.be/QXRITlQBitc", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

pres.writeFile({ fileName: "Ch12_Instrumental.pptx" })
  .then(fn => console.log(`✅ ${fn} created successfully`))
  .catch(err => console.error("❌ Error:", err));
