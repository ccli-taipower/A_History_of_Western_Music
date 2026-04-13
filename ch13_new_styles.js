const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 13: New Styles in the Seventeenth Century";
pres.author = "A History of Western Music, 10th ed.";

const C = {
  darkBg:   "1A1A30",
  gold:     "C89440",
  cream:    "F5F0E0",
  indigo:   "2A2A5A",
  royal:    "4A3A7A",
  darkText: "1A1A30",
  lightText:"F5F0E0",
  sand:     "E8D8A8",
  slate:    "2A2A40",
  mauve:    "7A5A8A",
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
  s.addText("CHAPTER 13", { x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6 });
  s.addText("NEW STYLES IN THE\nSEVENTEENTH CENTURY", { x: 0.3, y: 1.5, w: 9.4, h: 2.0, fontSize: 34, color: C.lightText, bold: true, align: "center", fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.65, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Baroque · Basso Continuo · Monody · Seconda Pratica · Affections", { x: 0.4, y: 3.8, w: 9.2, h: 0.4, fontSize: 13, color: C.sand, align: "center", fontFace: "Georgia" });
  s.addText("Textbook pp. 278–296", { x: 0.5, y: 4.8, w: 9, h: 0.3, fontSize: 11, color: C.gold, align: "center", fontFace: "Calibri" });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.royal); bottomBar(s, C.royal);
  s.addText("本章概覽 Chapter Overview", { x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.indigo, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });
  const sections = [
    ["🌍", "Europe in the 17th Century", "科學革命 · 30 年戰爭 · 絕對王權 · 殖民擴張"],
    ["🔄", "From Renaissance to Baroque", "第一實踐→第二實踐 · Monteverdi · Artusi 論戰"],
    ["🎤", "Monody & Basso Continuo", "新單聲歌曲 · Caccini《Le nuove musiche》(NAWM 74)"],
    ["🎭", "Seconda Pratica", "Monteverdi Cruda Amarilli (NAWM 75) · 文字主導和聲"],
    ["🎵", "General Baroque Traits", "情感論 · 調性 · 對比 · 裝飾 · 即興 · 記譜"],
    ["📜", "Enduring Innovations", "basso continuo · 調性和聲 · 獨奏家 · HIP 論爭"],
  ];
  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.royal }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 11, color: C.mauve, fontFace: "Calibri", margin: 0 });
  });
}

// ── SLIDE 3 · Historical Context ─────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("十七世紀的歐洲", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Europe in the Seventeenth Century", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🔬 科學革命", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Kepler 行星運動三定律 (1609)\n• Galileo 望遠鏡 · 運動定律 (1610s)\n• Bacon 經驗主義 · Descartes 理性主義\n• Newton 萬有引力 (1660s)\n• 世界從「權威說了算」→「觀察實驗說了算」\n\n⚔ 三十年戰爭 (1618–1648)\n• 席捲中歐 · 德意志人口損失 1/3\n• Westphalia 和約——主權國家體制確立\n• 日耳曼音樂在戰爭中重創\n\n👑 絕對王權\n• 法國 Louis XIV (1643–1715) 為典型\n• 凡爾賽宮：中央集權的象徵\n• 音樂 = 權力展示工具", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎭 音樂家的處境", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("🇮🇹 義大利\n• 城邦與教廷——威尼斯、羅馬、佛羅倫斯\n• 歌劇誕生 (1600) → 公共劇院 (1637 Venice)\n• 義大利音樂家遍布歐洲——國際語言\n\n🇫🇷 法國\n• 高度集權 · 音樂由王室品味主導\n• Lully 壟斷法國歌劇\n• Ballet de cour → Tragédie en musique\n\n🇬🇧 英格蘭\n• 内戰 (1640s) · Commonwealth (1649–60)\n• 清教徒禁止教堂音樂、關閉劇院\n• 1660 復辟後音樂復興 · Purcell\n\n🇩🇪 日耳曼\n• 三十年戰爭後百廢待舉\n• 各小邦仿效 Versailles——建立宮廷樂團\n• Schütz 為關鍵橋樑人物", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 4 · From Renaissance to Baroque ────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.royal); bottomBar(s, C.royal);
  s.addText("從文藝復興到巴洛克", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.indigo, fontFace: "Georgia", align: "center" });
  s.addText("Prima vs. Seconda Pratica · Artusi vs. Monteverdi", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 13, color: C.mauve, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.royal } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "E8E0F0" }, rounding: true });
  s.addText("📜 兩種實踐", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.indigo, fontFace: "Georgia" });
  s.addText("Prima Pratica（第一實踐）\n• Palestrina 式對位法的延續\n• 音樂的規則高於文字\n• 不協和音嚴格準備與解決\n• 1600 後稱 stile antico（古風格）\n• 持續用於教會音樂\n\nSeconda Pratica（第二實踐）\n• 文字（詩的情感）主導音樂\n• 為了表達詩意可以打破對位規則\n• 不協和音用於 dramatic effect\n• Monteverdi 的第五部牧歌集宣言 (1605)\n• 成為 17 世紀新風格的理論基礎\n\n💡 兩者並非取代——而是共存\n• stile antico (教會) · stile moderno (劇場)\n• 作曲家必須掌握兩種風格", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "E8E0F0" }, rounding: true });
  s.addText("⚔ Artusi vs. Monteverdi", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.indigo, fontFace: "Georgia" });
  s.addText("Giovanni Maria Artusi (1540–1613)\n• 保守派理論家 · 1600 發表批評\n• 批評某些未出版的牧歌（即 Monteverdi 的）\n• 指其不協和處理違反 Zarlino 規則\n\n🌟 Monteverdi 的回應\n• 1605 第五牧歌集序言\n• 宣告「第二實踐」——以文字為女主人\n• 承諾撰寫理論著作（未完成）\n• 弟弟 Giulio Cesare 1607 加以闡述\n\n🎵 Cruda Amarilli (NAWM 75)\n• 五聲部牧歌 · Guarini《Pastor fido》詩\n• 被 Artusi 批評的曲目之一\n• 不協和音未準備——用於表達「殘酷」\n• 七度、四度直接出現——衝擊 16 世紀規則\n• 音樂史上「新」vs.「舊」論戰的里程碑", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 5 · Monody & Basso Continuo ────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("單歌曲與數字低音", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Monody · Basso Continuo · Caccini · Le nuove musiche (NAWM 74)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 12, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎤 Monody 單歌曲", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 單獨歌聲 + basso continuo 伴奏\n• 約 1600 年誕生於佛羅倫斯\n• 目的：恢復古希臘音樂的情感力量\n• 文字清晰 · 歌手主導表達\n\n📚 Florentine Camerata\n• 1570–80s Bardi 伯爵的文化沙龍\n• Vincenzo Galilei (Galileo 之父) 批評複音\n  — 主張：只有單聲才能傳達文字情感\n• Caccini、Peri 等嘗試新的歌唱風格\n\n🌟 Giulio Caccini (1551–1618)\n• Le nuove musiche (1602) ——「新音樂」\n• 序言闡述新歌唱美學\n• Vedro 'l mio sol (NAWM 74)\n  — 以 strophic variation 形式\n  — 裝飾性歌唱 · 表現歌手技巧\n  — 典型 monody 範例", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🎹 Basso Continuo 數字低音", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• 巴洛克音樂最重要的技術創新\n• 最低聲部（bass line）持續進行\n  + 數字標記和聲（figured bass）\n• 由和聲樂器即興填充\n  — harpsichord / organ / lute / theorbo\n  — + cello / bassoon / violone 加倍低音\n\n📊 功能\n1. 支撐和聲架構\n2. 讓獨唱/獨奏者自由表達\n3. 統一合奏——任何編制皆可使用\n4. 簡化記譜——不用寫出全部聲部\n\n💡 觀念轉變\n• 從「聲部等權」→「兩極織度」\n  — 最上聲部（旋律）+ 最低聲部（bass）\n  — 中間聲部由 continuo 填充\n• 這就是 Baroque 的基本聲音\n\n📈 1600–1750 幾乎所有西方音樂使用 continuo\n• 被稱為 thoroughbass era", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 6 · Monteverdi Innovations ─────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.royal); bottomBar(s, C.royal);
  s.addText("蒙特威爾第的創新", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.indigo, fontFace: "Georgia", align: "center" });
  s.addText("Monteverdi · Zefiro torna (NAWM 76) · Stile concitato", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 13, color: C.mauve, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.royal } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "E8E0F0" }, rounding: true });
  s.addText("🌟 Monteverdi (1567–1643)", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.indigo, fontFace: "Georgia" });
  s.addText("跨越文藝復興與巴洛克的關鍵人物\n\n前半生 · Mantua (1590–1612)\n• Gonzaga 公爵的宮廷音樂家\n• 5 部牧歌集 (第 3–5 冊成為新風格宣言)\n• L'Orfeo (1607) · 首部偉大歌劇\n\n後半生 · Venice (1613–1643)\n• St. Mark's 樂長——最崇高的音樂職位\n• 第 6–8 部牧歌集\n• L'incoronazione di Poppea (1643)\n\n📊 八部牧歌集的演化\n• I–IV (1587–1603): 五聲部 a cappella\n• V (1605): 加 basso continuo\n• VI (1614): 1–6 聲部 + continuo\n• VII (1619): Concerto——2 聲部 + continuo\n• VIII (1638): Madrigali guerrieri et amorosi\n  — stile concitato（激動風格）\n  — 快速重複音 = 戰爭、憤怒", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "E8E0F0" }, rounding: true });
  s.addText("🎵 Zefiro torna (NAWM 76)", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.indigo, fontFace: "Georgia" });
  s.addText("• 取自《Scherzi musicali》(1632)\n• Ottavio Rinuccini 的十四行詩\n• 兩位男高音 + basso continuo\n\n📊 結構：Ciaccona（恰空）\n• 重複低音型（basso ostinato）\n  — 4 小節低音固定 · 全曲反覆\n• 上方兩聲部以各種方式變化\n• 描繪春風、花、鳥——word painting\n\n🎭 戲劇性轉折\n• 進入到「我獨自嘆息」時——\n  突然打斷 ciaccona 節奏\n• 改為 recitative 風格——自由拍\n• 表達孤獨與痛苦——音樂「停頓」\n• 然後又回到 ciaccona\n\n💡 為何重要？\n• 完美示範 basso ostinato 與情感表達的結合\n• ground bass 技法——影響 Purcell · Bach\n• 巴洛克「對比原則」的縮影", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 7 · General Traits of Baroque Music ────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("巴洛克音樂的共通特徵", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("General Traits of Baroque Music (ca. 1600–1750)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const traits = [
    ["🎭", "情感論 Doctrine of Affections", "音樂的目的 = 激發特定情感 · 每段音樂表現單一 Affekt\n每種情感有對應的音樂手段（調式、節奏、音型）"],
    ["📊", "兩極織度", "旋律 + bass 為兩極 · 中間聲部由 continuo 填充\n從 Renaissance 的聲部平等 → Baroque 的上下極端"],
    ["🎵", "節奏的兩極", "穩定的舞曲節奏 vs. 自由的 recitative\n速度對比成為結構原則 · 快-慢-快"],
    ["🏛", "功能和聲的確立", "調性（major/minor）取代 mode\nI–IV–V–I 進行成為基礎 · 離調為色彩"],
    ["🎤", "獨奏家崇拜", "virtuoso 歌手/器樂家地位上升 · castrato · 炫技\n即興裝飾成為表演核心要素"],
    ["📜", "三種風格共存", "Church · Chamber · Theater 三種場域三種風格\n+ stile antico / moderno / concitato 的區分"],
  ];
  traits.forEach(([icon, title, desc], i) => {
    const y = 1.2 + i * 0.7;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.62, fill: { color: C.slate }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.08, w: 0.5, h: 0.44, fontSize: 16, align: "center" });
    s.addText(title, { x: 0.95, y: y + 0.02, w: 3.2, h: 0.58, fontSize: 10, bold: true, color: C.gold, fontFace: "Georgia" });
    s.addText(desc, { x: 4.2, y: y + 0.02, w: 5.4, h: 0.58, fontSize: 8, color: C.sand, fontFace: "Calibri" });
  });
}

// ── SLIDE 8 · HIP & Enduring Innovations ─────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.royal); bottomBar(s, C.royal);
  s.addText("持久的創新 · 歷史知情演奏", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 24, bold: true, color: C.indigo, fontFace: "Georgia", align: "center" });
  s.addText("Enduring Innovations · Historically Informed Performance (HIP)", { x: 0.4, y: 0.76, w: 9.2, h: 0.35, fontSize: 12, color: C.mauve, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.royal } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: "E8E0F0" }, rounding: true });
  s.addText("💡 持久的創新", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.indigo, fontFace: "Georgia" });
  s.addText("1600 年前後引入的概念——至今仍是西方音樂的基石\n\n✓ Basso continuo → 後世和聲伴奏\n✓ Major-minor tonality → 共同語言至 20 世紀\n✓ 獨奏 + 伴奏的二元結構 → concerto, sonata\n✓ Opera / recitative / aria → 歌劇傳統\n✓ 情感作為音樂首要目的 → Romantic 繼承\n✓ 作曲家 vs. 演奏者的分工 → 現代音樂觀\n\n📊 時期劃分\n• Early Baroque (1600–1650): 義大利實驗期\n• Middle Baroque (1650–1700): 國際擴散\n• Late Baroque (1700–1750): Vivaldi · Bach · Handel\n\n⚠ 但「Baroque」一詞有爭議\n• 原意「不規則的珍珠」——曾含貶義\n• 20 世紀才廣泛用於音樂\n• 部分學者偏好「Thoroughbass era」", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: "E8E0F0" }, rounding: true });
  s.addText("🎻 HIP 歷史知情演奏", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.indigo, fontFace: "Georgia" });
  s.addText("Historically Informed Performance\n\n📖 核心理念\n• 使用原始樂器（period instruments）\n• 參考 17–18 世紀文獻的演奏指示\n• 嘗試重建「原本的聲音」\n\n🎵 實踐\n• Baroque violin（無 chinrest, gut strings）\n• 古鍵盤（harpsichord, fortepiano）\n• 較低音高（A=415 Hz vs. modern A=440）\n• 更少 vibrato · 更多 articulation\n• 即興裝飾——演奏者的責任\n\n🌟 代表團體\n• Jordi Savall · Hespèrion XXI\n• Nikolaus Harnoncourt · Concentus Musicus Wien\n• John Eliot Gardiner · English Baroque Soloists\n• Les Arts Florissants · William Christie\n• Ton Koopman · Amsterdam Baroque Orchestra\n\n⚠ 爭議\n• 「忠實原作」是否可能或必要？\n• HIP 本身已成「傳統」——年輕團體再反思", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 7.5, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

// ── SLIDE 9 · Timeline ───────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.royal); bottomBar(s, C.royal);
  s.addText("時間軸 · Timeline", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.indigo, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 0.82, w: 5, h: 0.04, fill: { color: C.royal } });
  const events = [
    ["1581", "V. Galilei · Dialogo della musica antica et della moderna"],
    ["ca. 1587", "Florentine Camerata at Bardi's house"],
    ["1600", "Artusi 批評 Monteverdi 牧歌"],
    ["1601", "Caccini · Euridice (published)"],
    ["1602", "Caccini · Le nuove musiche"],
    ["1605", "Monteverdi · Fifth Book (序言)"],
    ["1607", "Monteverdi · L'Orfeo"],
    ["1609", "Kepler · Astronomia nova"],
    ["1618–48", "三十年戰爭"],
    ["1619", "Monteverdi · Seventh Book (Concerto)"],
    ["1632", "Monteverdi · Scherzi musicali (Zefiro torna)"],
    ["1637", "Venice 首座公共歌劇院"],
    ["1638", "Monteverdi · Eighth Book (stile concitato)"],
    ["1643", "Monteverdi 逝世 · Poppea 上演"],
  ];
  events.forEach(([date, desc], i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 0.3 + col * 4.8;
    const y = 1.0 + row * 0.55;
    s.addShape(pres.ShapeType.rect, { x, y, w: 1.1, h: 0.44, fill: { color: C.royal } });
    s.addText(date, { x: x + 0.05, y: y + 0.06, w: 1.0, h: 0.32, fontSize: 9, bold: true, color: C.lightText, align: "center", fontFace: "Georgia" });
    s.addText(desc, { x: x + 1.2, y, w: 3.55, h: 0.44, fontSize: 8, color: C.darkText, fontFace: "Calibri", valign: "middle" });
  });
}

// ── SLIDE 10 · Key Terms & Listening ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);
  s.addText("關鍵詞彙 · 延伸閱讀", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Key Terms & Further Reading", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 13, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("🔑 Key Terms", { x: 0.45, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Baroque · thoroughbass era\n• prima pratica · seconda pratica\n• stile antico · stile moderno · stile concitato\n• monody · recitative · aria · arioso\n• basso continuo · figured bass · thoroughbass\n• concertato · concerto\n• Florentine Camerata · Count Bardi\n• Le nuove musiche · sprezzatura\n• doctrine of affections (Affektenlehre)\n• basso ostinato · ground bass · ciaccona\n• major/minor tonality\n• HIP · period instruments · A=415", {
    x: 0.5, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 3.95, fill: { color: C.slate }, rounding: true });
  s.addText("📚 Further Reading & 🎧 Listening", { x: 5.25, y: 1.38, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C.gold, fontFace: "Georgia" });
  s.addText("• Palisca, Claude. Baroque Music (1991)\n• Carter, Tim. Music in Late Renaissance and Early Baroque Italy (1992)\n• Fabbri, Paolo. Monteverdi (1994)\n• Tomlinson. Monteverdi and the End of the Renaissance (1987)\n\n🎧 NAWM 精選聆聽 (YouTube)\n• 74 · Caccini · Vedro 'l mio sol  youtu.be/s-DaH6zpLjs\n• 75 · Monteverdi · Cruda Amarilli  youtu.be/bKTQQ28sSNo\n• 76 · Monteverdi · Zefiro torna  youtu.be/85tCzdRt6UE", {
    x: 5.3, y: 1.72, w: 4.35, h: 3.45, fontSize: 8.5, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 2,
  });
}

pres.writeFile({ fileName: "Ch13_New_Styles.pptx" })
  .then(fn => console.log(`✅ ${fn} created successfully`))
  .catch(err => console.error("❌ Error:", err));
