const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 8: England and Burgundy in the Fifteenth Century";
pres.author = "A History of Western Music, 10th ed.";

const C = {
  darkBg:   "1B2B1F",
  gold:     "C8A020",
  cream:    "FBF5E6",
  wine:     "7A2830",
  forest:   "2E5A3A",
  darkText: "1B2B1F",
  lightText:"FBF5E6",
  midGreen: "3E6B4C",
  sand:     "E8D8A8",
  slate:    "3A4C38",
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
  s.addText("CHAPTER 8", {
    x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });
  s.addText("ENGLAND AND BURGUNDY\nIN THE FIFTEENTH CENTURY", {
    x: 0.3, y: 1.5, w: 9.4, h: 2.0, fontSize: 34, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.65, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("Contenance angloise · Dunstable · Binchois · Du Fay · The Polyphonic Mass", {
    x: 0.4, y: 3.8, w: 9.2, h: 0.4, fontSize: 14, color: C.sand, italic: true, align: "center", fontFace: "Georgia",
  });
  s.addText("Textbook pp. 159–179", {
    x: 0.5, y: 4.8, w: 9, h: 0.3, fontSize: 14, color: C.gold, align: "center", fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);

  s.addText("本章概覽 Chapter Overview", { x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.forest, fontFace: "Georgia", margin: 0 });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.82, w: 9.2, h: 0.03, fill: { color: C.sand } });

  const sections = [
    ["■", "English Music 英格蘭音樂", "Sarum rite · Faburden · Carol · 三六度協和 · Dunstable"],
    ["■", "Contenance Angloise 英式風韻", "Martin Le Franc 1440 詩篇 · 大陸作曲家學習英格蘭風格"],
    ["■", "Burgundian Court 勃艮第宮廷", "Philip the Good · 歐洲最盛的禮拜堂樂團 · 國際樂風"],
    ["■", "Binchois 賓切瓦", "ca. 1400–1460 · 勃艮第 chanson 大師 · rondeau De plus en plus"],
    ["■", "Guillaume Du Fay", "ca. 1397–1474 · 集大成者 · Se la face ay pale · Nuper rosarum flores"],
    ["■", "The Polyphonic Mass 複音彌撒", "Cantus-firmus mass · 四聲部 · L'homme armé 傳統 · 最尊貴的類型"],
  ];
  sections.forEach(([icon, title, sub], i) => {
    const y = 1.0 + i * 0.75;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.6, h: 0.58, fill: { color: C.forest }, rounding: true });
    s.addText(icon, { x: 0.4, y: y + 0.05, w: 0.6, h: 0.5, fontSize: 20, align: "center", margin: 0 });
    s.addText(title, { x: 1.15, y, w: 8.4, h: 0.3, fontSize: 14, bold: true, color: C.darkText, fontFace: "Georgia", margin: 0 });
    s.addText(sub, { x: 1.15, y: y + 0.28, w: 8.4, h: 0.26, fontSize: 14, color: C.midGreen, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── SLIDE 3 · English Music & Faburden ───────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("英格蘭音樂的特色", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("English Music · Sarum Rite · Faburden · Carol", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "2A3E2E" }, rounding: true });
  s.addText("■ English Sound 英式音響", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 頻繁的和聲三度、六度——常以平行進行\n• 徹底的協和——極少不協和\n• 簡單旋律、規則樂句、音節對音\n• 同節奏（homorhythmic）織度\n• Sarum rite：英格蘭特有的聖歌方言（中世紀至宗教改革）\n• 三聲部織度常以中聲部持 chant、上聲部平行四度、下聲部平行三度", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "2A3E2E" }, rounding: true });
  s.addText("■ Faburden & Carol", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("Faburden（ca. 1430 首次見諸文獻）\n  • 即興複音技法——中聲部唱 chant、上方四度、下方三度\n  • 規則化系統——不識譜的修士也能唱出正確複音\n  • 結果是連續不斷的 6/3 和弦（三和六度）\n\nCarol（英格蘭獨有體裁）\n  • 源自中世紀 carole（舞歌）\n  • 兩到三聲部 · 英文／拉丁文或混合\n  • 結構：stanzas（詩節）+ burden（副歌）\n  • 多為宗教題材：聖誕節、聖母馬利亞\n  • 例：Alleluia: A newë work (NAWM 32)", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 4 · John Dunstable ─────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("John Dunstable", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 32, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("ca. 1390–1453 · 15 世紀最傑出的英格蘭作曲家", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 14, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  // Bio
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "F3E5C5" }, rounding: true });
  s.addText("■ 生平 Biography", { x: 0.45, y: 1.4, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("• 同時為數學家、天文學家——回應中世紀四藝（quadrivium）傳統\n• 非神職人員——服侍多位貴族贊助人\n• ca. 1422–1437 服侍 John, Duke of Bedford（攝政法國）\n• 之後服侍 Queen Joan、Humphrey Duke of Gloucester\n• 1437 繼承 Bedford 在法國的部分土地——可能長年居於法國\n• 作品大部分保存在歐陸手稿中——證明其國際影響力", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.4, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Works
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "F3E5C5" }, rounding: true });
  s.addText("■ 主要作品 Major Works", { x: 5.25, y: 1.4, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("涵蓋當時所有複音類型（約 60 部）：\n• 3 部 polyphonic mass cycles\n• 2 組 Gloria-Credo 配對\n• 15 部其他 Mass Ordinary 樂章\n• 12 部 isorhythmic motets\n• 6 部 plainchant settings\n• 20 部其他拉丁聖樂 · 5 首世俗歌曲\n\n代表作：\n• Quam pulchra es (NAWM 33)\n  自由創作、文字清晰、自然節奏\n• Regina caeli laetare — cantus paraphrase", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.4, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 5 · Contenance Angloise ────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Contenance Angloise 英式風韻", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Martin Le Franc · Le champion des dames (1440–42)", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Quote box
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 1.7, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ Le Franc 原詩意譯（1440–42）", { x: 0.45, y: 1.38, w: 9.1, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("「Du Fay 與 Binchois 有一種新的實踐方式——在高音與低音中\n創造活潑的協和。他們採取了英格蘭人的風韻（contenance angloise），\n並追隨 Dunstable——這賦予他們的音樂奇妙的愉悅感。」", {
    x: 0.5, y: 1.75, w: 9.0, h: 1.25, fontSize: 14, color: C.sand, italic: true, fontFace: "Georgia", align: "center",
  });

  // What it means
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.1, w: 9.4, h: 2.15, fill: { color: "2A3E2E" }, rounding: true });
  s.addText("■ 英式風韻的音樂特徵 Musical Features", { x: 0.45, y: 3.2, w: 9.1, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 頻繁使用和聲三度與六度——常以平行進行呈現\n• 徹底的協和——極少不協和（僅允許經過音與留音）\n• 避免平行五度與八度（中世紀複音常見）\n• 同節奏織度（homorhythmic）——各聲部節奏相似\n• 簡單、優美、可歌的旋律線條 · 清晰音節宣告\n• 成為 15 世紀國際風格的核心 — foundation of mid-15c international style", {
    x: 0.5, y: 3.68, w: 9.0, h: 1.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 6 · Burgundian Court ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);

  s.addText("勃艮第宮廷 The Burgundian Court", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.forest, fontFace: "Georgia", align: "center" });
  s.addText("Europe's Foremost Center of Music · ca. 1400–1477", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.12, w: 5, h: 0.04, fill: { color: C.forest } });

  // Duchy info
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "E0EEDA" }, rounding: true });
  s.addText("■ The Duchy 公國領地", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.forest, fontFace: "Georgia", margin: 0 });
  s.addText("• 名義上為法王封臣，實際幾與法王平起平坐\n• 封地：Burgundy、荷蘭、比利時、東北法國、盧森堡、洛林\n• 重心城市：Lille、Bruges、Ghent、Brussels\n• 1419–35 與英格蘭結盟對抗法王\n• 1477 Charles the Bold 戰死 → 法國吸收勃艮第\n\n四位主要公爵：\n• Philip the Bold (1363–1404) 建立禮拜堂\n• John the Fearless (1404–19)\n• Philip the Good (1419–67) 全盛期\n• Charles the Bold (1467–77) 業餘作曲家", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Chapel info
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "E0EEDA" }, rounding: true });
  s.addText("■ The Chapel 禮拜堂樂團", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.forest, fontFace: "Georgia", margin: 0 });
  s.addText("• 1384 Philip the Bold 建立\n• 1445 達 23 位歌手——歐洲僅次於英王\n• 音樂家主要來自法蘭德斯與低地諸省\n• 另維持樂手隊：trumpet、drums、vielle、lute、harp、organ\n\n四種主要作品類型\n• Chansons · Motets · Magnificats · Mass Ordinary\n\n• 多數為三聲部，音域略廣於 14 世紀\n• cantus 主旋律 · tenor 對位 · contratenor 填和聲", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 7 · Gilles Binchois ────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Gilles Binchois", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 32, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("ca. 1400–1460 · 勃艮第宮廷 chanson 大師", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.gold } });

  // Bio
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "2A3E2E" }, rounding: true });
  s.addText("■ 生平 Biography", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("• 生於 Mons — 受訓為唱詩班男童\n• 曾服侍英王軍隊 William Pole\n   → 直接接觸英格蘭音樂\n• 1427 加入 Philip the Good 禮拜堂\n  服侍三十餘年\n• 1453 以豐厚年金退休\n• 與 Dunstable、Du Fay 並列三大\n\n主要作品：\n• 28 彌撒樂章 · 6 Magnificats\n• 29 motets · 51 rondeaux · 7 ballades", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // De plus en plus
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "2A3E2E" }, rounding: true });
  s.addText("■ De plus en plus (NAWM 34, ca. 1425)", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
  s.addText("形式：rondeau (ABaAabAB)\n15 世紀最常見的 chanson 形式\n\n風格混合\n• 6/8 拍 · hemiola 交叉節奏\n• cantus 宣告清晰 · 音節對音為主\n• 主旋律上行三和弦後級進下行\n• tenor 節奏較慢 · 旋律優美\n• contratenor 填補和聲\n• 大量三度與六度——幾乎完全協和\n\n終止式\n• 傳統 major 6th → octave\n• 新式 contratenor 八度躍升→像屬→主", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.sand, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 8 · Guillaume Du Fay Biography ─────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("Guillaume Du Fay", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 32, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("ca. 1397–1474 · 國際風格的集大成者 · 時代最富盛名的作曲家", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 14, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.wine } });

  // Career timeline
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 2.35, fill: { color: "F3E5C5" }, rounding: true });
  s.addText("■ 生涯軌跡 Career Timeline", { x: 0.45, y: 1.4, w: 9.1, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("• 生於布魯塞爾附近 · 1409 進入 Cambrai 唱詩班\n• 1414–18 出席康士坦斯公會議\n• 1420 服侍 Rimini 的 Malatesta 家族\n• 1426–28 服侍 Bologna 樞機主教\n• 1428–33、1435–37 兩度服務教宗禮拜堂\n• 1433–39 擔任 Savoy 宮廷禮拜堂長\n• 晚年終老於 Cambrai — 作品保存於約 100 份手稿", {
    x: 0.5, y: 1.85, w: 9.0, h: 1.8, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Major works
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.75, w: 9.4, h: 1.6, fill: { color: "F3E5C5" }, rounding: true });
  s.addText("■ 主要作品 Major Works", { x: 0.45, y: 3.83, w: 9.1, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("• 至少 6 部 masses · 35 部其他彌撒樂章 · 4 部 Magnificats\n• 60 首聖詩聖歌設定 · 24 部 motets · 34 首 plainchant 旋律\n• 60 首 rondeaux · 8 首 ballades · 13 首其他世俗歌曲\n• 代表作：Resvellies vous (NAWM 35)、Christe redemptor (NAWM 36)、\n  Missa Se la face ay pale (NAWM 37)、Nuper rosarum flores", {
    x: 0.5, y: 4.28, w: 9.0, h: 1.05, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 9 · Du Fay Works ───────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("Du Fay 的代表作品", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("Chansons · Motets · Fauxbourdon · Nuper rosarum flores", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const works = [
    ["Resvellies vous (1423, NAWM 35)", "Ballade 形式 · 寫於 Rimini 婚禮 · 法義元素、少見英格蘭特徵 · Ars Subtilior 式快速音型"],
    ["Christe redemptor omnium (NAWM 36)", "Office hymn · 使用 fauxbourdon · cantus 中 paraphrase chant · 偶數詩節為複音、奇數為齊唱"],
    ["Se la face ay pale (ca. 1430s, NAWM 37a)", "Ballade · 自由創作而非固定形式 · 強烈英格蘭影響——成熟國際風格"],
    ["Nuper rosarum flores (1436)", "Isorhythmic motet · 寫於佛羅倫斯大教堂 Brunelleschi 穹頂落成典禮 · 兩條 isorhythmic tenor 呼應雙層穹頂結構"],
    ["Supremum est mortalibus bonum (1433)", "混合 isorhythm、fauxbourdon、自由對位 · 教宗 Eugene 與神聖羅馬皇帝 Sigismund 會晤"],
    ["Missa Se la face ay pale (1450s)", "Cantus-firmus mass · 首部以世俗歌曲為 cantus firmus 的完整彌撒（見下頁）"],
  ];
  works.forEach(([title, desc], i) => {
    const y = 1.32 + i * 0.66;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.58, fill: { color: "2A3E2E" }, rounding: true });
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 0.1, h: 0.58, fill: { color: C.gold } });
    s.addText(title, { x: 0.55, y: y + 0.05, w: 9.1, h: 0.28, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(desc, { x: 0.55, y: y + 0.3, w: 9.1, h: 0.28, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 10 · Fauxbourdon ───────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);

  s.addText("Fauxbourdon 假低音", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 32, bold: true, color: C.forest, fontFace: "Georgia", align: "center" });
  s.addText("Continental Adaptation of English Faburden", { x: 0.4, y: 0.78, w: 9.2, h: 0.35, fontSize: 14, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.15, w: 5, h: 0.04, fill: { color: C.forest } });

  // Explanation
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 1.7, fill: { color: "E0EEDA" }, rounding: true });
  s.addText("■ 技法說明 The Technique", { x: 0.45, y: 1.38, w: 9.1, h: 0.4, fontSize: 14, bold: true, color: C.forest, fontFace: "Georgia", margin: 0 });
  s.addText("• 只寫出兩個聲部：cantus 與 tenor——主要以平行六度進行\n• 第三聲部不寫出——由歌手自動唱在 cantus 下方精確的平行四度\n• 結果形成連續的 6/3 和弦（類似英格蘭 faburden）\n• 每樂句末終止於開放五度與八度\n• 受英格蘭 faburden 啟發——但技法不同", {
    x: 0.5, y: 1.88, w: 9.0, h: 1.09, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Usage & comparison
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.1, w: 4.6, h: 2.15, fill: { color: "E0EEDA" }, rounding: true });
  s.addText("■ 用途 Usage", { x: 0.45, y: 3.18, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.forest, fontFace: "Georgia", margin: 0 });
  s.addText("• Du Fay 有 24 部作品使用\n• 其他作曲家超過 100 部\n• 主要用於較簡單的日課聖歌：聖詩、對唱曲、詩篇、頌歌\n• cantus 中通常是 chant 的 paraphrase", {
    x: 0.5, y: 3.68, w: 4.35, h: 1.54, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 3.1, w: 4.6, h: 2.15, fill: { color: "E0EEDA" }, rounding: true });
  s.addText("🆚 Faburden vs. Fauxbourdon", { x: 5.25, y: 3.18, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.forest, fontFace: "Georgia", margin: 0 });
  s.addText("Faburden（英格蘭）\n• 中聲部 = chant · 上下三四度\n• 即興、規則式\n\nFauxbourdon（歐陸）\n• cantus = paraphrased chant\n• 中聲部四度以下（未寫出）\n• 寫譜作品", {
    x: 5.3, y: 3.68, w: 4.35, h: 1.54, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 11 · The Polyphonic Mass ───────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("The Polyphonic Mass 複音彌撒", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("The Most Prestigious Musical Genre · 15th–16th Century", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  // Five unification methods
  const methods = [
    ["1", "Stylistic Coherence 風格一致", "僅靠整體風格連結——最鬆散 · 無旋律關聯"],
    ["2", "Plainsong Mass 聖歌彌撒", "各樂章用各自對應的 chant（如 Machaut 彌撒）· 僅禮儀相關"],
    ["3", "Motto Mass 動機彌撒", "各樂章以相同 head motive（開頭音型）開始"],
    ["4", "Cantus-firmus / Tenor Mass", "同一 cantus firmus 置於所有樂章的 tenor · 15 世紀下半主流 · 通常四聲部"],
    ["5", "Cantus-firmus/Imitation Mass", "不只借用 tenor，也借用原曲其他聲部的片段（Du Fay Missa Se la face ay pale）"],
  ];
  methods.forEach(([n, title, desc], i) => {
    const y = 1.32 + i * 0.78;
    s.addShape(pres.ShapeType.rect, { x: 0.3, y, w: 9.4, h: 0.7, fill: { color: "2A3E2E" }, rounding: true });
    s.addShape(pres.ShapeType.ellipse, { x: 0.45, y: y + 0.1, w: 0.55, h: 0.5, fill: { color: C.gold } });
    s.addText(n, { x: 0.45, y: y + 0.1, w: 0.55, h: 0.5, fontSize: 18, bold: true, color: C.darkText, fontFace: "Georgia", align: "center" });
    s.addText(title, { x: 1.15, y: y + 0.08, w: 8.5, h: 0.3, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(desc, { x: 1.15, y: y + 0.36, w: 8.5, h: 0.3, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 12 · Cantus-firmus Mass details ────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.wine); bottomBar(s, C.wine);

  s.addText("Cantus-firmus Mass 的發展", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.wine, fontFace: "Georgia", align: "center" });
  s.addText("Four Voices · Missa Caput · L'homme armé Tradition", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.12, w: 5, h: 0.04, fill: { color: C.wine } });

  // Key innovations
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 4.6, h: 4.15, fill: { color: "F3E5C5" }, rounding: true });
  s.addText("■ 關鍵創新 Key Innovations", { x: 0.45, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("• 英格蘭作曲家最早寫作——Dunstable、Leonel Power（d. 1445）\n• 早期為三聲部——cantus firmus 置於 tenor\n• 問題：tenor 若為最低聲部，則無法作為和聲基礎\n• 解決：1440s 匿名 Missa Caput 首創——在 tenor 下加入第四聲部\n• 四聲部命名：\n  - superius / cantus / discantus（英文 soprano 來源）\n  - contratenor altus → altus → alto\n  - tenor（持 cantus firmus）\n  - contratenor bassus → bassus → bass\n• 四聲部織度至此成為標準——持續至今", {
    x: 0.5, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });

  // Du Fay Missa Se la face
  s.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.3, w: 4.6, h: 4.15, fill: { color: "F3E5C5" }, rounding: true });
  s.addText("■ Missa Se la face ay pale (NAWM 37)", { x: 5.25, y: 1.38, w: 4.3, h: 0.4, fontSize: 14, bold: true, color: C.wine, fontFace: "Georgia", margin: 0 });
  s.addText("Du Fay 1450s 寫於 Savoy\n• 首部以世俗歌曲為 cantus firmus 的完整彌撒\n• 以自己的 ballade tenor 為 c.f.\n• 可能紀念 1453 都靈裹屍布——歌詞「蒼白面容」喻基督受難\n\n節奏變形（augmentation）\n• Kyrie、Sanctus、Agnus：音值加倍\n• Gloria、Credo：c.f. 出現三次\n  三倍 → 雙倍 → 原速（旋律最清楚可辨）\n• Amen 借用 ballade 其他聲部（c.f./imitation）\n• 各樂章共用 head motive——結構統一", {
    x: 5.3, y: 1.88, w: 4.35, h: 3.55, fontSize: 14, color: C.darkText, fontFace: "Calibri", paraSpaceAfter: 1, valign: "top",
  });
}

// ── SLIDE 13 · Timeline ──────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("年代大事記 Timeline", { x: 0.4, y: 0.18, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia", align: "center" });
  s.addText("England, Burgundy, and Music · 1400–1477", { x: 0.4, y: 0.72, w: 9.2, h: 0.38, fontSize: 14, color: C.sand, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.12, w: 7, h: 0.04, fill: { color: C.gold } });

  const events = [
    ["1414–18", "Council of Constance 結束大分裂"],
    ["1415", "英王 Henry V 於 Agincourt 大勝"],
    ["1419–67", "Philip the Good 統治勃艮第"],
    ["1419–35", "勃艮第與英格蘭結盟"],
    ["1422", "Henry VI 即位 · Bedford 攝政法國"],
    ["1423", "Du Fay 在 Rimini 寫 Resvellies vous"],
    ["ca. 1425", "Binchois · De plus en plus"],
    ["1427", "Binchois 加入 Philip the Good 禮拜堂"],
    ["1428–39", "Du Fay 往返教宗禮拜堂與 Savoy"],
    ["ca. 1430s", "首批 cantus-firmus mass cycles"],
    ["1436", "Du Fay · Nuper rosarum flores（佛羅倫斯穹頂）"],
    ["1440–42", "Le Franc · Le champion des dames"],
    ["1440s", "匿名 Missa Caput——四聲部開端"],
    ["1450s", "Du Fay · Missa Se la face ay pale"],
    ["1453", "百年戰爭結束 · Dunstable 卒"],
    ["1460", "Binchois 卒"],
    ["1474", "Du Fay 卒 於 Cambrai"],
    ["1477", "Charles the Bold 戰死 · 法國併吞勃艮第"],
  ];
  events.forEach(([date, ev], i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.3 + col * 4.75;
    const y = 1.3 + row * 0.42;
    s.addText(date, { x: x + 0.1, y, w: 1.5, h: 0.35, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", margin: 0 });
    s.addText(ev, { x: x + 1.6, y, w: 3.1, h: 0.35, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top" });
  });
}

// ── SLIDE 14 · Legacy / Summary ──────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.forest); bottomBar(s, C.forest);

  s.addText("An Enduring Musical Language", { x: 0.4, y: 0.2, w: 9.2, h: 0.55, fontSize: 28, bold: true, color: C.forest, fontFace: "Georgia", align: "center" });
  s.addText("Chapter Summary · 持久的音樂語言", { x: 0.4, y: 0.75, w: 9.2, h: 0.35, fontSize: 14, italic: true, color: C.slate, fontFace: "Georgia", align: "center" });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 1.12, w: 5, h: 0.04, fill: { color: C.forest } });

  const summary = [
    ["■󠁧󠁢󠁥󠁮󠁧󠁿", "英格蘭的貢獻", "三六度協和、頻繁平行音程、同節奏織度——15 世紀大陸風格的靈感來源"],
    ["■", "國際風格的誕生", "Du Fay、Binchois 融合法、義、英三種傳統——mid-15c international style"],
    ["■", "複音彌撒的確立", "從配對樂章到完整 cycles——cantus-firmus mass 成為最尊貴的類型"],
    ["■", "四聲部織度", "匿名 Missa Caput 首創（1440s）——SATB 架構自此成為標準"],
    ["■", "新的音響美學", "徹底協和、控制不協和、避免平行五八度——15–16 世紀的共同語言"],
    ["■", "持久的影響", "Du Fay 時代的聲響對現代人仍聽來「熟悉」——中世紀與文藝復興的分水嶺"],
  ];
  summary.forEach(([icon, title, desc], i) => {
    const y = 1.3 + i * 0.67;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 0.55, h: 0.55, fill: { color: C.forest }, rounding: true });
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

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 1.3, w: 9.4, h: 1.95, fill: { color: "2A3E2E" }, rounding: true });
  s.addText("■ 重要術語 Key Terms", { x: 0.45, y: 1.34, w: 9.1, h: 0.28, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", valign: "top", margin: 0 });

  const termsLeft = [
    "Contenance angloise",
    "Sarum rite · Faburden",
    "Carol · Burden",
    "Cantilena · Motet",
    "Fauxbourdon",
    "Head motive · Motto mass",
  ];
  const termsRight = [
    "Polyphonic mass cycle",
    "Cantus-firmus / Tenor mass",
    "Cantus-firmus/Imitation mass",
    "Plainsong mass",
    "Superius · Altus · Tenor · Bassus",
    "L'homme armé · Missa Caput",
  ];
  termsLeft.forEach((t, i) => {
    s.addText("• " + t, { x: 0.5, y: 1.65 + i * 0.25, w: 4.5, h: 0.24, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 });
  });
  termsRight.forEach((t, i) => {
    s.addText("• " + t, { x: 5.1, y: 1.65 + i * 0.25, w: 4.5, h: 0.24, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 });
  });

  // Listening (new)
  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 3.30, w: 9.4, h: 1.75, fill: { color: "3A2015" }, rounding: true });
  s.addText("■ 聆聽 Listen (NAWM 32–37)", { x: 0.45, y: 3.34, w: 9.1, h: 0.28, fontSize: 14, bold: true, color: C.gold, fontFace: "Georgia", valign: "top", margin: 0 });
  s.addText(
    "• A newë work (English carol, NAWM 32)  youtu.be/odCOfgR0HkI\n" +
    "• Dunstable — Quam pulchra es (NAWM 33)  youtu.be/lMoyfuCbnjs\n" +
    "• Binchois — De plus en plus (NAWM 34)  youtu.be/dzjB7HaFKFg\n" +
    "• Du Fay — Resvellies vous (NAWM 35)  youtu.be/iObHTuqHxRo\n" +
    "• Du Fay — Conditor alme siderum (NAWM 36)  youtu.be/UqYgWMdkhSg\n" +
    "• Du Fay — Se la face ay pale (NAWM 37)  youtu.be/_EMbGN2jeno",
    { x: 0.5, y: 3.62, w: 9.1, h: 1.40, fontSize: 14, color: C.sand, fontFace: "Calibri", valign: "top", margin: 0 }
  );

  s.addShape(pres.ShapeType.rect, { x: 0.3, y: 5.10, w: 9.4, h: 0.45, fill: { color: "3E2030" }, rounding: true });
  s.addText("Fallows, Dufay · Bent, Dunstaple · Kirkman, Polyphonic Mass · Strohm, Rise of European Music", {
    x: 0.5, y: 5.15, w: 9.1, h: 0.35, fontSize: 14, color: C.sand, fontFace: "Calibri", italic: true, valign: "top", margin: 0,
  });
}

pres.writeFile({ fileName: "Ch08_England_Burgundy.pptx" }).then(() => {
  console.log("■ Ch08_England_Burgundy.pptx created successfully");
});
