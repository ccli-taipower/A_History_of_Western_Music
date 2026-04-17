const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 16: France, England, Spain, the New World, and Russia in the Seventeenth Century";
pres.author = "A History of Western Music, 10th ed.";

// Versailles palette — regal French Baroque: deep midnight / Bourbon gold / ivory cream
const C = {
  darkBg:   "1A1428",   // deep purple-midnight
  gold:     "C8A030",   // Bourbon gold
  cream:    "F5F0E0",   // parchment ivory
  royal:    "1E3A5F",   // royal blue
  azure:    "2E5A88",   // lighter blue accent
  darkText: "1A1428",
  lightText:"F5F0E0",
  sand:     "E8D8A8",
  slate:    "1A2440",
  crimson:  "8B1A1A",   // deep red
  navy:     "1B2838",
};

function darkSlide(pres) { const s = pres.addSlide(); s.background = { color: C.darkBg }; return s; }
function lightSlide(pres) { const s = pres.addSlide(); s.background = { color: C.cream }; return s; }
function accentSlide(pres) { const s = pres.addSlide(); s.background = { color: C.royal }; return s; }
function topBar(s, color) { s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.12, fill: { color: color || C.gold } }); }
function bottomBar(s, color) { s.addShape(pres.ShapeType.rect, { x: 0, y: 5.5, w: "100%", h: 0.125, fill: { color: color || C.gold } }); }

// ── SLIDE 1 · Title ──────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.15, fill: { color: C.gold } });
  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.47, w: "100%", h: 0.155, fill: { color: C.gold } });

  s.addText("A HISTORY OF WESTERN MUSIC  ·  TENTH EDITION", {
    x: 0.5, y: 0.4, w: 9, h: 0.4, fontSize: 18, color: C.sand, charSpacing: 3, align: "center", fontFace: "Georgia",
  });
  s.addText("CHAPTER 16", {
    x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 24, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });
  s.addText("FRANCE, ENGLAND, SPAIN,\nTHE NEW WORLD, AND RUSSIA\nIN THE SEVENTEENTH CENTURY", {
    x: 0.3, y: 1.55, w: 9.4, h: 2.2, fontSize: 36, color: C.lightText, bold: true, align: "center", fontFace: "Georgia", lineSpacingMultiple: 1.1,
  });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.85, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Louis XIV · Lully · Purcell · Jacquet de la Guerre · Hidalgo", {
    x: 0.4, y: 4.0, w: 9.2, h: 0.45, fontSize: 18, color: C.sand, align: "center", fontFace: "Georgia",
  });
  s.addText("Textbook pp. 339–370", {
    x: 0.5, y: 4.9, w: 9, h: 0.35, fontSize: 18, color: C.gold, align: "center", fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.royal); bottomBar(s, C.royal);

  s.addText("本章概覽 Chapter Overview", { x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.royal, fontFace: "Georgia" });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const sections = [
    ["France 法國", "Louis XIV · 宮廷舞蹈文化 · 音樂機構 · Lully 的悲劇歌劇"],
    ["French Style 法式風格", "Notes inegales · agrements · grand motet · 魯特琴與大鍵琴音樂"],
    ["England 英格蘭", "內戰與復辟 · 假面劇 · Purcell 與 Dido and Aeneas"],
    ["Spain & New World", "Zarzuela · Hidalgo · 新世界第一部歌劇 · villancico"],
    ["Russia 俄羅斯", "東正教傳統 · kontsert · kanty · 彼得大帝西化"],
  ];
  sections.forEach(([title, sub], i) => {
    const y = 1.05 + i * 0.88;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.08, h: 0.7, fill: { color: C.royal }, rectRadius: 0.04 });
    s.addText(title, { x: 0.65, y, w: 8.9, h: 0.36, fontSize: 22, bold: true, color: C.darkText, fontFace: "Georgia" });
    s.addText(sub, { x: 0.65, y: y + 0.36, w: 8.9, h: 0.32, fontSize: 18, color: C.azure, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 3 · Louis XIV & the Sun King ───────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("太陽王路易十四 Louis XIV, the Sun King", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("r. 1643–1715 · 七十年統治改變了歐洲文化", {
    x: 0.4, y: 0.78, w: 9.2, h: 0.4, fontSize: 20, color: C.sand, fontFace: "Calibri", valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.22, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const bullets = [
    "自稱「太陽王」，以希臘太陽神 Apollo 自居\nStyled himself the Sun King; identified with Apollo, god of music & light",
    "集中藝術權力：建立皇家學院——舞蹈(1661)、文學(1663)、科學(1669)、歌劇(1669)、建築(1671)\nCentralized arts through royal academies",
    "重建羅浮宮，興建凡爾賽宮——貴族被「留」在宮中\nRebuilt the Louvre; built Versailles to control the nobility",
    "法國藝術強調秩序、優雅、克制——與義大利炫技形成對比\nFrench art emphasized order, elegance, and restraint vs. Italian virtuosity",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.4 + i * 1.0, w: 9, h: 0.9,
      fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.05,
    });
  });
}

// ── SLIDE 4 · Dance at Court ─────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.royal); bottomBar(s, C.royal);

  s.addText("宮廷舞蹈文化 Dance at Court", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.royal, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const bullets = [
    "舞蹈是法國文化核心——展現自制、優雅、高貴\nDance was central to French culture: self-control, grace, grandeur",
    "Court ballet (ballet de cour)：結合歌曲、合唱、器樂舞蹈的戲劇作品\nLouis XIII & XIV regularly performed in ballets",
    "Beauchamp 建立芭蕾五個基本腳位（至今沿用）\nEstablished 5 basic foot positions still used in ballet today",
    "Feuillet《Choregraphie》(1700)：第一套完整舞蹈記譜法\nFirst complete method for recording dance steps and gestures",
    "舞蹈作為政治控制工具——貴族忙於練舞，遠離權力鬥爭\nDance as political control: kept nobility busy at Versailles",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.0 + i * 0.88, w: 9, h: 0.8,
      fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.0,
    });
  });
}

// ── SLIDE 5 · Court Music Institutions ───────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("宮廷音樂機構 Music at Court", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.gold } });

  s.addText("宮廷音樂分為三大部門，共 150–200 位音樂家\nThree divisions with 150–200 musicians, organized like the state itself", {
    x: 0.5, y: 0.95, w: 9, h: 0.55, fontSize: 18, color: C.sand, fontFace: "Calibri", valign: "top",
  });

  const depts = [
    ["Chapel 教堂", "歌手、管風琴家、器樂手——負責宗教禮拜\nSingers, organists, instrumentalists for religious services"],
    ["Chamber 室內", "獨唱、弦樂、魯特琴、大鍵琴、長笛——室內娛樂\nSolo singers, strings, lute, harpsichord, flute for indoor entertainment"],
    ["Great Stable 大馬廄", "管樂、銅管、定音鼓——軍事與戶外儀式\nWind, brass, timpani for military and outdoor ceremonies"],
  ];
  depts.forEach(([title, desc], i) => {
    const y = 1.6 + i * 1.1;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 9.2, h: 0.95, fill: { color: C.slate }, rectRadius: 0.08 });
    s.addText(title, { x: 0.6, y, w: 3.5, h: 0.95, fontSize: 22, bold: true, color: C.gold, fontFace: "Georgia", valign: "middle" });
    s.addText(desc, { x: 4.0, y, w: 5.4, h: 0.95, fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "middle", lineSpacingMultiple: 1.05 });
  });

  s.addText("Vingt-quatre Violons du Roi (24 Violins) — 首個大型弦樂合奏 = 現代管弦樂團原型\nPetits Violons (18 strings) — Louis XIV 的私人樂團", {
    x: 0.5, y: 4.95, w: 9, h: 0.52, fontSize: 14, color: C.gold, fontFace: "Calibri", lineSpacingMultiple: 1.1, valign: "top",
  });
}

// ── SLIDE 6 · Jean-Baptiste Lully — Life ─────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.crimson); bottomBar(s, C.crimson);

  s.addText("呂利 Jean-Baptiste Lully (1632–1687)", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.crimson, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const bullets = [
    "生於佛羅倫斯，14 歲來巴黎——雖為義大利人，卻塑造了法國音樂\nBorn in Florence; came to Paris at 14 as Italian tutor to a cousin of Louis XIV",
    "出色的舞者——在《Ballet de la nuit》(1653) 中與國王共舞\nBrilliant dancer; Louis appointed him court composer & director of Petits Violons",
    "1661 成為國王室內樂總監，掌管所有宮廷樂團\nSuperintendent of Music for the King's Chamber",
    "1672 購得皇家特許——獨佔法國歌劇製作權\nPurchased royal privilege: exclusive right to produce sung drama in France",
    "1687 指揮《Te Deum》時以指揮杖擊中腳趾，傷口感染壞疽而亡\nDied after hitting his foot with conducting staff; gangrene followed",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.0 + i * 0.88, w: 9, h: 0.8,
      fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.0,
    });
  });
}

// ── SLIDE 7 · Tragedie en musique ────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("悲劇歌劇 Tragedie en musique", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Lully 與劇作家 Philippe Quinault 合作創造的法國歌劇形式", {
    x: 0.4, y: 0.78, w: 9.2, h: 0.4, fontSize: 20, color: C.sand, fontFace: "Calibri", valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.22, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const bullets = [
    "五幕結構，取材古代神話或騎士傳奇\nFive-act dramas from ancient mythology or chivalric tales",
    "序幕(prologue)歌頌國王——劇情暗喻克制激情的美德\nPrologue with allegorical praise of the king",
    "穿插大量 divertissements（舞蹈、合唱、器樂插曲）\nFrequent divertissements: dances, choruses, instrumental interludes",
    "French overture：慢-快兩段式序曲（莊嚴附點節奏 + 賦格快板）\nFrench overture: slow dotted section + fast imitative section",
    "歌唱「airs」比義大利詠嘆調更簡潔——強調詩意與旋律\nAirs: simpler than Italian arias; syllabic, tuneful, no virtuosic display",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.4 + i * 0.82, w: 9, h: 0.75,
      fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.0,
    });
  });
}

// ── SLIDE 8 · NAWM 85 — Lully Armide ─────────────────────────────────────────
{
  const s = accentSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("NAWM 85b · Lully, Armide (1686)", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Act II 的 divertissement — 女巫 Armide 命魔鬼引誘戰士 Renaud", {
    x: 0.4, y: 0.78, w: 9.2, h: 0.4, fontSize: 20, color: C.sand, fontFace: "Calibri", valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.22, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const bullets = [
    "牧羊人/牧羊女/仙女歌頌愛的歡愉，穿插器樂舞蹈\nShepherds & nymphs sing of love; instrumental dances interspersed",
    "NAWM 85c：Armide 的獨白——宣敘調與詠嘆調交融\nArmide's monologue: recitative flows into air without a break",
    "Lully 的宣敘調模仿法語語調——不規則拍號交替\nRecitative follows French speech: irregular meters (duple/triple alternate)",
    "附點節奏的管弦序奏 = tirades，象徵超自然力量\nOrchestral prelude with dotted rhythms (tirades) = supernatural power",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.4 + i * 0.95, w: 9, h: 0.85,
      fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.05,
    });
  });

  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 5.0, w: 7, h: 0.4, fill: { color: C.navy }, rectRadius: 0.08 });
  s.addText("https://www.youtube.com/watch?v=P1N_FYfcqMg", {
    x: 1.5, y: 5.0, w: 7, h: 0.4, fontSize: 18, color: C.gold, align: "center", fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=P1N_FYfcqMg" },
  });
}

// ── SLIDE 9 · French Performance Practice ────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.royal); bottomBar(s, C.royal);

  s.addText("法式演奏慣例 French Performance Practice", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.royal, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const items = [
    ["Notes inegales\n不等音符", "記譜為均等的音符，演奏時長短交替（如搖擺節奏）\nEvenly notated notes performed with alternating long-short durations"],
    ["Overdotting\n過度附點", "附點音符被延長得比記譜更長——使節奏更銳利\nDotted notes held longer than written; sharpens rhythmic profile"],
    ["Agrements\n裝飾音", "顫音、倚音等小裝飾——強調重要音、增添表情\nTrills, appoggiaturas: emphasize cadences, enhance expression"],
  ];
  items.forEach(([title, desc], i) => {
    const y = 1.05 + i * 1.4;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 2.6, h: 1.2, fill: { color: C.royal }, rectRadius: 0.08 });
    s.addText(title, { x: 0.5, y, w: 2.4, h: 1.2, fontSize: 20, bold: true, color: C.lightText, fontFace: "Georgia", valign: "middle", align: "center" });
    s.addText(desc, { x: 3.2, y, w: 6.4, h: 1.2, fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "middle", lineSpacingMultiple: 1.05 });
  });

  s.addText("這些慣例是法國品味的象徵——義大利人的炫技裝飾被視為粗俗\nThese practices were signs of refined French taste; Italian embellishments were seen as vulgar", {
    x: 0.5, y: 5.0, w: 9, h: 0.45, fontSize: 18, italic: true, color: C.azure, fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 10 · French Church Music ───────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("法國教會音樂 French Church Music", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const bullets = [
    "Petit motet（小經文歌）：少數聲部 + 數字低音\nSmall motet for few voices with continuo",
    "Grand motet（大經文歌）：獨唱、大小合唱團、管弦樂團\nLarge motet for soloists, double chorus, and orchestra — Lully's Te Deum (NAWM 86)",
    "Marc-Antoine Charpentier (1643–1704)：引入義大利不協和與半音對位\nIntroduced Italian dissonance & chromaticism; studied with Carissimi in Rome",
    "Charpentier, Le reniement de Saint Pierre (NAWM 87)：音樂化彼得三次不認主\nDramatizes Peter's threefold denial; closing chorus of ravishing dissonant suspensions",
    "Michel-Richard de Lalande (1657–1726)：Louis XIV 晚年最愛的教會作曲家\n70+ grand motets with masterly command of the genre",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.0 + i * 0.88, w: 9, h: 0.8,
      fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.0,
    });
  });
}

// ── SLIDE 11 · Air de cour & Song ────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.royal); bottomBar(s, C.royal);

  s.addText("法國歌曲 Air de cour & Song", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.royal, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const bullets = [
    "Air（歌謠）是 17 世紀法國最重要的聲樂室內曲類型\nThe air was the leading genre of vocal chamber music in France",
    "Air de cour（宮廷歌）逐漸退流行，被 air serieux（嚴肅歌）取代\nAir de cour gradually replaced by air serieux (serious air)",
    "Air a boire（飲酒歌）：輕鬆、幽默、政治或愛情主題\nDrinking songs on love, pastoral, political, or frivolous topics",
    "Michel Lambert (ca.1610–1696)：Lully 的岳父，最具影響力的歌曲作曲家\nLully's father-in-law; published first collection of airs with basso continuo (1660)",
    "通常為分節歌形式，1–3 聲部 + 魯特琴或數字低音伴奏\nStrophic; scored for 1–3 voices with lute or continuo accompaniment",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.0 + i * 0.88, w: 9, h: 0.8,
      fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.0,
    });
  });
}

// ── SLIDE 12 · Lute & Harpsichord Music ──────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("魯特琴與大鍵琴音樂 Lute & Keyboard Music", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const bullets = [
    "Denis Gaultier (1603–1672)：最重要的魯特琴作曲家\nLeading lute composer; La rhetorique des dieux (NAWM 88)",
    "大鍵琴(clavecin)在 17 世紀中取代魯特琴成為主要獨奏樂器\nHarpsichord displaced the lute as the main solo instrument",
    "Style luthe / Style brise（琉特風格 / 碎奏風格）：和弦琶音化而非同時彈奏\nBroken-chord texture imitating lute style; chords arpeggiated, not blocked",
    "重要大鍵琴家：Chambonnieres · D'Anglebert · Jacquet de la Guerre · F. Couperin\nClavecinists served Louis XIV and published collections for amateurs",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.1 + i * 1.05, w: 9, h: 0.95,
      fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.05,
    });
  });
}

// ── SLIDE 13 · Suite & Ordre ─────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.royal); bottomBar(s, C.royal);

  s.addText("組曲與套曲 Suite & Ordre", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.royal, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  s.addText("法國作曲家將一系列舞曲組合為 suite（組曲）\nFrench composers grouped dances into suites, each dance with unique character", {
    x: 0.5, y: 0.95, w: 9, h: 0.5, fontSize: 18, color: C.azure, fontFace: "Calibri", valign: "top",
  });

  const dances = [
    ["Allemande", "中慢速、4/4、嚴肅", "German origin; serious character"],
    ["Courante", "中速、3/2 或 6/4、機智", "French; witty, metric ambiguity"],
    ["Sarabande", "慢速、3/4、莊嚴", "Spanish/New World; stress on beat 2"],
    ["Gigue", "快速、6/8 或 12/8、活潑", "English/Irish; lively, imitative"],
  ];
  dances.forEach(([name, zh, en], i) => {
    const y = 1.6 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 2.2, h: 0.65, fill: { color: C.royal }, rectRadius: 0.06 });
    s.addText(name, { x: 0.4, y, w: 2.2, h: 0.65, fontSize: 20, bold: true, color: C.lightText, fontFace: "Georgia", align: "center", valign: "middle" });
    s.addText(zh, { x: 2.8, y, w: 3.3, h: 0.65, fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "middle" });
    s.addText(en, { x: 6.1, y, w: 3.5, h: 0.65, fontSize: 18, color: C.azure, fontFace: "Calibri", valign: "middle" });
  });

  s.addText("多數舞曲為 binary form（二段體）：||: A :||: B :|| — 各段調性走向不同\nMost dances in binary form; optional dances: gavotte, minuet, chaconne, etc.", {
    x: 0.5, y: 4.7, w: 9, h: 0.65, fontSize: 18, color: C.darkText, fontFace: "Calibri", lineSpacingMultiple: 1.1, valign: "top",
  });
}

// ── SLIDE 14 · NAWM 89 — Jacquet de la Guerre ────────────────────────────────
{
  const s = accentSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("NAWM 89 · Jacquet de la Guerre", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Pieces de clavecin, Suite No. 3 in A Minor (1687)", {
    x: 0.4, y: 0.78, w: 9.2, h: 0.4, fontSize: 20, color: C.sand, fontFace: "Calibri", valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.22, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const bullets = [
    "Elisabeth-Claude Jacquet de la Guerre (1665–1729)：五歲即為太陽王演奏\nChild prodigy who sang and played harpsichord for Louis XIV from age 5",
    "89a Prelude：unmeasured prelude（無拍號前奏曲）——法國獨特體裁\nFree rhythmic notation; sounds improvisatory, like a toccata",
    "89b Allemande → 89c Courante → 89d Sarabande → 89e Gigue\n展示組曲四大標準舞曲，全為二段體",
    "89f Chaconne：rondeau 形式 + 89g Gavotte + 89h Minuet\nRondeau form with alternating couplets",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.4 + i * 0.88, w: 9, h: 0.82,
      fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.05,
    });
  });

  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 5.05, w: 7, h: 0.38, fill: { color: C.navy }, rectRadius: 0.08 });
  s.addText("https://www.youtube.com/watch?v=xPgHQpp_tBM", {
    x: 1.5, y: 5.05, w: 7, h: 0.38, fontSize: 16, color: C.gold, align: "center", fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=xPgHQpp_tBM" },
  });
}

// ── SLIDE 15 · Emulation of French Style ─────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("法國風格的歐洲影響 Emulation of French Style", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const bullets = [
    "三十年戰爭後法國成為歐洲領導強權（1648–）\nFrance was the leading power in Europe after the Thirty Years' War",
    "法國品味、禮儀、藝術被廣泛模仿——從英國到俄國\nFrench tastes, manners, and arts widely imitated across Europe",
    "1660s 起歐洲存在兩大音樂風格：義大利 vs. 法國\nTwo dominant national styles: Italian (opera, sonata, toccata) vs. French (dance, suite, overture)",
    "法式序曲影響 Bach、Handel 的序曲與組曲\nFrench overture used across Europe for ballets, operas, oratorios, instrumental works",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.1 + i * 1.05, w: 9, h: 0.95,
      fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.05,
    });
  });
}

// ── SLIDE 16 · England: Historical Context ───────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.crimson); bottomBar(s, C.crimson);

  s.addText("英格蘭歷史背景 England: Civil War to Restoration", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.crimson, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const bullets = [
    "英國為有限君主制——國王與國會共享權力\nLimited monarchy: king shared rule with Parliament",
    "Charles I (r.1625–49) 擴張王權，引發內戰(1642–49)——被處決\nCivil War; king executed 1649; Commonwealth under Cromwell (1649–58)",
    "1660 Restoration 復辟：Charles II 回歸——宮廷音樂復興\nMonarchy restored; Charles II returned from exile in France",
    "1689 光榮革命 & 權利法案：君主立憲確立\nGlorious Revolution; Bill of Rights → constitutional monarchy",
    "王室財力不及法國——公眾付費音樂會因此誕生(1672)\nRoyal house had less money → public concerts invented in England",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.0 + i * 0.88, w: 9, h: 0.8,
      fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.0,
    });
  });
}

// ── SLIDE 17 · English Masque & Theater ──────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.crimson); bottomBar(s, C.crimson);

  s.addText("假面劇與劇場 English Masque & Theater", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const bullets = [
    "Masque（假面劇）：自亨利八世以來的宮廷娛樂——結合舞蹈、歌唱、佈景\nFavorite court entertainment since Henry VIII; shared elements with opera",
    "Cromwell 清教徒時期禁止舞台戲劇——但允許私人音樂\nPuritans banned stage plays but not private musical entertainments",
    "僅存完整假面劇：Cupid and Death (1653) — Locke & Gibbons\nOnly complete masque whose music survives",
    "復辟後劇場加入音樂插曲——但英國觀眾不接受全唱歌劇\nAfter 1660, plays included musical episodes; fully sung opera didn't catch on",
    "僅兩部全唱歌劇：Blow's Venus and Adonis (ca.1683) & Purcell's Dido and Aeneas\nOnly two through-sung dramas for private audiences",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.0 + i * 0.88, w: 9, h: 0.8,
      fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.0,
    });
  });
}

// ── SLIDE 18 · Henry Purcell — Life ──────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.crimson); bottomBar(s, C.crimson);

  s.addText("乍浦爾 Henry Purcell (1659–1695)", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.crimson, fontFace: "Georgia",
  });
  s.addText("\"The British Orpheus\" — 英國最偉大的巴洛克作曲家", {
    x: 0.4, y: 0.78, w: 9.2, h: 0.4, fontSize: 20, color: C.azure, fontFace: "Calibri", valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.22, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const bullets = [
    "父親為皇家教堂成員；Purcell 自幼加入教堂唱詩班\nFather was Chapel Royal member; Purcell joined as choirboy",
    "八歲出版第一首歌曲——公認的神童作曲家\nPublished first song at age 8; recognized as a gifted prodigy",
    "1677 繼 Matthew Locke 為宮廷小提琴作曲家；1679 西敏寺管風琴師\nComposer-in-ordinary for the violins; organist of Westminster Abbey",
    "創作涵蓋幾乎所有體裁：歌劇、附隨音樂、頌歌、歌曲、室內樂、鍵盤\nWrote in almost every genre: opera, incidental music, anthems, songs, chamber, keyboard",
    "36 歲英年早逝，葬於西敏寺——被譽為「不列顛的奧菲斯」\nDied at 36; buried in Westminster Abbey",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.4 + i * 0.8, w: 9, h: 0.72,
      fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.0,
    });
  });
}

// ── SLIDE 19 · Dido and Aeneas ───────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.crimson); bottomBar(s, C.crimson);

  s.addText("Dido and Aeneas (ca. 1687–88)", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("迷你歌劇傑作——三幕僅約一小時，四位主角", {
    x: 0.4, y: 0.78, w: 9.2, h: 0.4, fontSize: 20, color: C.sand, fontFace: "Calibri", valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.22, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const bullets = [
    "首演於女子寄宿學校——可能原為宮廷創作\nFirst known performance at a girls' boarding school; may have been intended for court",
    "以 Blow 的 Venus and Adonis 為範本——融合英、法、義三國風格\nModeled on Blow's Venus and Adonis; combines English, French, Italian elements",
    "French overture + Lully 式齊唱合唱 + 義大利式詠嘆調（建立在固定低音上）\nFrench overture; homophonic choruses like Lully; Italian arias over ground bass",
    "英式宣敘調：靈活配合英語的重音、節奏與情感\nEnglish recitatives mold flexibly to accents and emotions of English text",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.4 + i * 1.0, w: 9, h: 0.9,
      fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.05,
    });
  });
}

// ── SLIDE 20 · NAWM 90 — Purcell Dido's Lament ──────────────────────────────
{
  const s = accentSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("NAWM 90b · Purcell, \"When I am laid in earth\"", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 24, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Dido's Lament — 歌劇史上最動人的詠嘆調之一", {
    x: 0.4, y: 0.78, w: 9.2, h: 0.4, fontSize: 20, color: C.sand, fontFace: "Calibri", valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.22, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const bullets = [
    "建立在下行半音固定低音(descending chromatic tetrachord)之上\nDescending chromatic ground bass — Italian tradition of lament",
    "Purcell 在低音線加入更多半音，增強悲劇色彩\nAdds chromaticism to bass line; sighing figures in voice and violins",
    "懸留音(suspensions)落在強拍，強化不協和的痛苦感\nSuspended notes rearticulated on strong beats intensify dissonance",
    "NAWM 90a：Dido 的臨終宣敘調 \"Thy hand, Belinda\"\nSlow stepwise chromatic descent; portrays dying Dido",
    "NAWM 90c：結尾合唱 \"With drooping wings\" — 模仿 Blow\nClosing chorus with descending figures; profound sorrow",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.35 + i * 0.72, w: 9, h: 0.65,
      fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.0,
    });
  });

  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 5.05, w: 7, h: 0.38, fill: { color: C.navy }, rectRadius: 0.08 });
  s.addText("https://www.youtube.com/watch?v=GjfbMBSzzXg", {
    x: 1.5, y: 5.05, w: 7, h: 0.38, fontSize: 16, color: C.gold, align: "center", fontFace: "Calibri", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=GjfbMBSzzXg" },
  });
}

// ── SLIDE 21 · English Vocal & Instrumental Music ────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.crimson); bottomBar(s, C.crimson);

  s.addText("英國聲樂與器樂 English Vocal & Instrumental Music", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.crimson, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const bullets = [
    "Semi-opera（半歌劇）：spoken play + overture + masques，如 The Fairy Queen (1692)\nDramatic opera: spoken play with 4+ masques or musical episodes",
    "Catch（輪唱曲）：幽默、粗俗的卡農——紳士社交娛樂\nRound or canon with humorous, often ribald text; gentlemen's entertainment",
    "Anglican church music：復辟後恢復 anthem 與 Service 傳統\nAfter Restoration, anthems and Services revived; Charles II favored French grand motet style",
    "Viol consort 音樂：In Nomines & fantasias — Jenkins, Locke, Purcell\nAmateur music-making; last important examples ca. 1680",
    "Public concerts 公開音樂會：1672 年 London 首創——付費入場\nFirst public concert series advertised in London Gazette, December 1672",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.0 + i * 0.88, w: 9, h: 0.8,
      fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.0,
    });
  });
}

// ── SLIDE 22 · Spain and the New World ───────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.crimson); bottomBar(s, C.crimson);

  s.addText("西班牙與新世界 Spain and the New World", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const bullets = [
    "1600 年西班牙帝國為全球最富——白銀來自新世界殖民地\nBy 1600, Spain was the richest country: silver from New World colonies",
    "Juan Hidalgo (1614–1685)：西班牙音樂劇場之父\nFounder of Spanish musical theater; created the zarzuela with Calderon",
    "Zarzuela：交替唱說的輕歌劇——西班牙的 national genre\nLight mythological play alternating sung and spoken dialogue",
    "La purpura de la rosa (NAWM 91)：1701 Lima——新世界第一部歌劇\nFirst opera produced in the New World; by Tomas de Torrejon y Velasco (1644–1728)",
    "教會音樂：villancico（聖誕歌）——Padilla (NAWM 92) 在 Puebla, Mexico\nChristmas villancico Albricias pastores by Juan Gutierrez de Padilla (ca.1590–1664)",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.0 + i * 0.88, w: 9, h: 0.8,
      fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.0,
    });
  });
}

// ── SLIDE 23 · Spanish Instrumental & Church Music ───────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.crimson); bottomBar(s, C.crimson);

  s.addText("西班牙器樂與教會音樂 Spanish Instrumental & Sacred Music", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.crimson, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const bullets = [
    "管風琴：tiento（提恩托）——即興式模仿作品，類似幻想曲\nTiento: improvisatory imitative genre akin to the fantasia",
    "Juan Bautista Jose Cabanilles (1644–1712)：西班牙管風琴領袖\nLeading Spanish organist; Tiento de batalla imitates trumpet calls",
    "豎琴與吉他：主要室內樂器——圍繞舞曲與固定低音變奏\nHarp and guitar: main chamber instruments; sarabande, chacona, passacalle",
    "殖民地教會音樂將歐洲技法與當地文化融合\nColonial church music blended European polyphony with local traditions",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.1 + i * 1.05, w: 9, h: 0.95,
      fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.05,
    });
  });
}

// ── SLIDE 24 · Russia ────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("俄羅斯 Russia", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("中世紀以來文化孤立——17 世紀開始吸收西歐影響", {
    x: 0.4, y: 0.78, w: 9.2, h: 0.4, fontSize: 20, color: C.sand, fontFace: "Calibri", valign: "top",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.22, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const bullets = [
    "俄羅斯東正教會：禁止樂器——znamenny chant（知名聖詠）為單聲部\nRussian Orthodox Church banned instruments; znamenny chant was monophonic",
    "Nikolay Diletsky (ca.1630–after 1680)：引入西方和聲與對位\nIdea grammatikii musikiyskoy (1679): earliest known description of the circle of fifths",
    "Kontsert（聖樂協奏曲）：多聲部合唱，Vasiliy Titov (ca.1650–1715) 為代表\nSacred concerto for voices alone; Titov set all 150 psalms as kanty",
    "Kanty（三聲部歌曲）：上二聲部平行三度——簡單、流行\nThree-voice song; top voices in parallel thirds; sacred or secular",
    "Peter the Great (r.1682–1725)：西化改革——建立聖彼得堡，引入歐洲音樂\nWesternization: founded St. Petersburg (1703); brought European music and theater",
  ];
  bullets.forEach((t, i) => {
    s.addText(t, {
      x: 0.5, y: 1.4 + i * 0.82, w: 9, h: 0.75,
      fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: { indent: 18 }, lineSpacingMultiple: 1.0,
    });
  });
}

// ── SLIDE 25 · Timeline ──────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.royal); bottomBar(s, C.royal);

  s.addText("年表 Timeline", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.royal, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const eventsL = [
    ["1621–65", "Philip IV of Spain 西班牙"],
    ["1625–49", "Charles I (England) 英國內戰"],
    ["1643–1715", "Louis XIV 法國太陽王"],
    ["1652", "First London coffeehouse"],
    ["1653", "Ballet de la nuit; Cupid & Death"],
    ["1657", "Hidalgo: first zarzuela"],
    ["1660", "Restoration of English monarchy"],
  ];
  const eventsR = [
    ["1661", "Louis XIV takes power; Versailles"],
    ["1672", "Lully begins operas; first public concert (London)"],
    ["1679", "Diletsky, Idea of Musical Grammar"],
    ["1686", "Lully, Armide"],
    ["1687", "Jacquet de la Guerre, Pieces de clavecin"],
    ["ca.1688", "Purcell, Dido and Aeneas"],
    ["1701", "Torrejon, La purpura de la rosa (Lima)"],
  ];

  eventsL.forEach(([yr, ev], i) => {
    const y = 0.95 + i * 0.6;
    s.addText(yr, { x: 0.3, y, w: 1.1, h: 0.5, fontSize: 18, bold: true, color: C.royal, fontFace: "Georgia" });
    s.addText(ev, { x: 1.4, y, w: 3.6, h: 0.5, fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "top" });
  });
  eventsR.forEach(([yr, ev], i) => {
    const y = 0.95 + i * 0.6;
    s.addText(yr, { x: 5.2, y, w: 1.2, h: 0.5, fontSize: 18, bold: true, color: C.royal, fontFace: "Georgia" });
    s.addText(ev, { x: 6.4, y, w: 3.3, h: 0.5, fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 26 · Key Terms ─────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("關鍵術語 Key Terms", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.gold } });

  const termsL = [
    "Court ballet / ballet de cour 宮廷芭蕾",
    "Tragedie en musique 悲劇歌劇",
    "French overture 法式序曲",
    "Divertissement 嬉遊曲插段",
    "Air / air serieux 歌謠",
    "Notes inegales 不等音符",
    "Overdotting 過度附點",
    "Agrements 裝飾音",
  ];
  const termsR = [
    "Style luthe / brise 琉特/碎奏風格",
    "Suite / ordre 組曲",
    "Grand motet 大經文歌",
    "Masque 假面劇",
    "Semi-opera 半歌劇",
    "Catch 輪唱曲",
    "Zarzuela 薩蘇拉歌劇",
    "Kontsert / kanty 俄式聖樂/歌曲",
  ];

  termsL.forEach((t, i) => {
    s.addText(t, { x: 0.5, y: 1.0 + i * 0.52, w: 4.5, h: 0.48, fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top", bullet: { indent: 14 } });
  });
  termsR.forEach((t, i) => {
    s.addText(t, { x: 5.2, y: 1.0 + i * 0.52, w: 4.5, h: 0.48, fontSize: 18, color: C.lightText, fontFace: "Calibri", valign: "top", bullet: { indent: 14 } });
  });
}

// ── SLIDE 27 · Listening Guide ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.royal); bottomBar(s, C.royal);

  s.addText("聆聽指南 Listening Guide", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.royal, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const works = [
    ["NAWM 85", "Lully, Armide (overture, divertissement, monologue)", "https://www.youtube.com/watch?v=P1N_FYfcqMg"],
    ["NAWM 86", "Lully, Te Deum (conclusion)", "https://www.youtube.com/watch?v=r798LTlyZTU"],
    ["NAWM 87", "Charpentier, Le reniement de Saint Pierre", ""],
    ["NAWM 88", "Denis Gaultier, La coquette virtuose", ""],
    ["NAWM 89", "Jacquet de la Guerre, Suite No. 3 in A Minor", "https://www.youtube.com/watch?v=xPgHQpp_tBM"],
    ["NAWM 90", "Purcell, Dido and Aeneas (recit, lament, chorus)", "https://www.youtube.com/watch?v=GjfbMBSzzXg"],
    ["NAWM 91", "Torrejon y Velasco, La purpura de la rosa", ""],
    ["NAWM 92", "Padilla, Albricias pastores (villancico)", ""],
  ];

  works.forEach(([num, title, url], i) => {
    const y = 1.0 + i * 0.55;
    s.addText(num, { x: 0.4, y, w: 1.5, h: 0.48, fontSize: 18, bold: true, color: C.royal, fontFace: "Georgia" });
    const opts = { x: 1.9, y, w: 7.7, h: 0.48, fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "top" };
    if (url) { opts.hyperlink = { url }; opts.color = C.azure; }
    s.addText(title, opts);
  });
}

// ── SLIDE 28 · End ───────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.15, fill: { color: C.gold } });
  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.47, w: "100%", h: 0.155, fill: { color: C.gold } });

  s.addText("CHAPTER 16", {
    x: 0.5, y: 1.0, w: 9, h: 0.6, fontSize: 24, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });
  s.addText("France, England, Spain,\nthe New World, and Russia\nin the Seventeenth Century", {
    x: 0.5, y: 1.7, w: 9, h: 1.8, fontSize: 36, color: C.lightText, bold: true, align: "center", fontFace: "Georgia", lineSpacingMultiple: 1.1,
  });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.65, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("End of Chapter 16 · 第十六章結束", {
    x: 0.5, y: 3.85, w: 9, h: 0.5, fontSize: 22, color: C.sand, align: "center", fontFace: "Georgia",
  });
  s.addText("Next: Chapter 17 — Italy and Germany in the Late Seventeenth Century", {
    x: 0.5, y: 4.5, w: 9, h: 0.4, fontSize: 18, color: C.gold, align: "center", fontFace: "Calibri", valign: "top",
  });
}

// ── Generate ─────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "Ch16_France_England.pptx" })
  .then(() => console.log("Created Ch16_France_England.pptx"))
  .catch(e => console.error(e));
