const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Chapter 15: Music for Chamber and Church in the Early Seventeenth Century";
pres.author = "A History of Western Music, 10th ed.";

// Early 17C palette — warm, olive-gold, old master tonality
const C = {
  darkBg:   "1C1C14",
  gold:     "C8A030",
  cream:    "F5F0E0",
  olive:    "556B2F",
  sage:     "6B8E4E",
  darkText: "1C1C14",
  lightText:"F5F0E0",
  sand:     "E8D8A8",
  slate:    "2A2A1E",
  bronze:   "8C7853",
  moss:     "4A5A2A",
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
  s.addText("CHAPTER 15", {
    x: 0.5, y: 0.9, w: 9, h: 0.55, fontSize: 20, color: C.gold, bold: true, align: "center", fontFace: "Georgia", charSpacing: 6,
  });
  s.addText("MUSIC FOR CHAMBER AND CHURCH\nIN THE EARLY SEVENTEENTH CENTURY", {
    x: 0.3, y: 1.55, w: 9.4, h: 1.8, fontSize: 32, color: C.lightText, bold: true, align: "center", fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 2.5, y: 3.5, w: 5, h: 0.04, fill: { color: C.gold } });
  s.addText("十七世紀初的室內樂與教會音樂", {
    x: 0.4, y: 3.65, w: 9.2, h: 0.5, fontSize: 24, color: C.sand, align: "center", fontFace: "Georgia",
  });
  s.addText("Monteverdi · Strozzi · Carissimi · Schütz · Frescobaldi · Marini", {
    x: 0.4, y: 4.25, w: 9.2, h: 0.4, fontSize: 18, color: C.bronze, align: "center", fontFace: "Georgia",
  });
  s.addText("Textbook pp. 326–353", {
    x: 0.5, y: 4.9, w: 9, h: 0.35, fontSize: 18, color: C.gold, align: "center", fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 2 · Chapter Overview ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.olive); bottomBar(s, C.olive);

  s.addText("本章概覽 Chapter Overview", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.6, fontSize: 28, bold: true, color: C.olive, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.85, w: 9.2, h: 0.04, fill: { color: C.sand } });

  const bullets = [
    "Italian Vocal Chamber Music 義大利聲樂室內樂\n  — Concertato madrigal, basso ostinato, cantata",
    "Catholic Sacred Music 天主教聖樂\n  — Sacred concerto, oratorio (Grandi · Carissimi)",
    "Lutheran Church Music 路德宗教會音樂\n  — Heinrich Schütz: Italian style meets German text",
    "Jewish Liturgical Music 猶太禮儀音樂\n  — Salamone Rossi: polyphony in the synagogue",
    "Instrumental Music 器樂\n  — Frescobaldi (toccata) · Marini (violin sonata)",
  ];
  bullets.forEach((txt, i) => {
    const y = 1.05 + i * 0.88;
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 9.2, h: 0.78, fill: { color: i % 2 === 0 ? C.sand : "EDE5CC" }, rectRadius: 0.08 });
    s.addText(txt, { x: 0.6, y: y + 0.05, w: 8.8, h: 0.7, fontSize: 19, color: C.darkText, fontFace: "Calibri", valign: "middle" });
  });
}

// ── SLIDE 3 · Italian Vocal Chamber Music: Concertato Madrigal ───────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("義大利聲樂室內樂：協奏曲風牧歌", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Italian Vocal Chamber Music: The Concertato Madrigal", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.45, w: 7, h: 0.04, fill: { color: C.gold } });

  const pts = [
    "Concertato medium 協奏媒介: solo voices + basso continuo\n  + optional instruments — texture contrast is key",
    "Monteverdi's Books 5–8 (1605–1638) trace the evolution:\n  from polyphonic madrigal → concerted madrigal with continuo",
    "Book 7 Concerto (1619): strophic variations, canzonettas\n  Book 8 Madrigali guerrieri et amorosi (1638): stile concitato",
    "Forms include madrigals, canzonettas, arias, recitatives,\n  dialogues — all with ritornellos & basso continuo",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.67 + i * 0.92, w: 9, h: 0.82, fontSize: 19, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: { type: "number", numberStartAt: i + 1, color: C.gold },
    });
  });
}

// ── SLIDE 4 · NAWM 82: Monteverdi Ohimè dov'è il mio ben ───────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.olive); bottomBar(s, C.olive);

  s.addText("NAWM 82: Monteverdi — Ohimè dov'è il mio ben", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 26, bold: true, color: C.olive, fontFace: "Georgia",
  });
  s.addText("蒙特威爾第《噢，我的愛人在何方》", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.bronze, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.45, w: 9.2, h: 0.04, fill: { color: C.sand } });

  const pts = [
    "Romanesca bass pattern: repeating harmonic progression\nover which two soprano voices weave expressive duet",
    "From Book 7 Concerto (1619): strophic variations on a\nground bass — each stanza varies melody over same bass",
    "Two sopranos + basso continuo: intimate chamber texture;\nritornello frames each section (instrumental interlude)",
    "Demonstrates shift from polyphonic madrigal → concertato\nchamber song with basso continuo as structural foundation",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.62 + i * 0.82, w: 9, h: 0.78, fontSize: 19, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.olive,
    });
  });

  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 5.0, w: 9.2, h: 0.4, fill: { color: C.olive }, rectRadius: 0.06 });
  s.addText("https://www.youtube.com/watch?v=bJfHoxZlMYw", {
    x: 0.5, y: 5.0, w: 9, h: 0.4, fontSize: 18, color: C.cream, fontFace: "Calibri", align: "center", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=bJfHoxZlMYw" },
  });
}

// ── SLIDE 5 · Basso Ostinato ─────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("固定低音 Basso Ostinato / Ground Bass", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Repeating Bass Patterns as Structural Foundation", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.45, w: 7, h: 0.04, fill: { color: C.gold } });

  const pts = [
    "Basso ostinato 固定低音: a bass pattern that repeats while\nmelody changes above — usually 2, 4, or 8 bars in triple meter",
    "Romanesca, Ruggiero: well-known bass patterns from\nSpain/Italy used for songs & instrumental variations",
    "Descending tetrachord 下行四度: stepwise descent spanning\na 4th — associated with lament; Monteverdi's Lamento della ninfa",
    "Chacona / Ciaccona 夏乾舞曲: from the Americas via Spain;\nI–V–vi–V bass progression; Monteverdi's Zefiro torna (1632)",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.67 + i * 0.92, w: 9, h: 0.82, fontSize: 19, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.gold,
    });
  });
}

// ── SLIDE 6 · Cantata Development ────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.olive); bottomBar(s, C.olive);

  s.addText("清唱套曲的發展 The Cantata", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 28, bold: true, color: C.olive, fontFace: "Georgia",
  });
  s.addText("From Strophic Song to Multi-Section Drama", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.bronze, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.45, w: 9.2, h: 0.04, fill: { color: C.sand } });

  const pts = [
    "Cantata = \"a piece to be sung\" (It. cantare);\nbefore 1620 meant a published collection of strophic arias",
    "By mid-century: secular composition for solo voice + continuo,\nmultiple sections of recitative + aria on a dramatic text",
    "Composed for private aristocratic performances;\npreserved mostly in manuscripts, not printed",
    "Leading cantata composers: Luigi Rossi, Antonio Cesti,\nGiacomo Carissimi, and Barbara Strozzi",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.62 + i * 0.82, w: 9, h: 0.78, fontSize: 19, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.olive,
    });
  });
}

// ── SLIDE 7 · Barbara Strozzi ────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("芭芭拉·史特羅乞 Barbara Strozzi (1619–1677)", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("The Most Prolific Composer of Secular Vocal Music of Her Century", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 20, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.45, w: 7, h: 0.04, fill: { color: C.gold } });

  const pts = [
    "Born in Venice; adopted (perhaps biological) daughter of\npoet Giulio Strozzi; studied with Cavalli (Monteverdi's student)",
    "Published 8 collections (1644–1664): over 100 madrigals,\narias, cantatas, motets — more cantatas than any peer",
    "Lagrime mie (NAWM 77): solo cantata alternating recitative,\narioso, & aria; focus on unrequited love; dissonant expression",
    "Did not perform publicly due to social class;\npublished as a way to be \"heard\" beyond private gatherings",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.67 + i * 0.92, w: 9, h: 0.82, fontSize: 19, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.gold,
    });
  });
}

// ── SLIDE 8 · Catholic Sacred Music: Sacred Concerto ─────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.olive); bottomBar(s, C.olive);

  s.addText("天主教聖樂：宗教協奏曲", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 28, bold: true, color: C.olive, fontFace: "Georgia",
  });
  s.addText("Catholic Sacred Music: The Sacred Concerto", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.bronze, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.45, w: 9.2, h: 0.04, fill: { color: C.sand } });

  const pts = [
    "Stile antico 古風 vs. stile moderno 現代風 coexist;\nPalestrina's style preserved alongside basso continuo textures",
    "Large sacred concerto 大型宗教協奏曲: cori spezzati,\nmultiple choirs + soloists + instruments (G. Gabrieli tradition)",
    "Small sacred concerto 小型宗教協奏曲: 1–4 solo voices +\ncontinuo (± 1–2 violins); practical for smaller churches",
    "Viadana's Cento concerti ecclesiastici (1602):\nfirst printed sacred vocal music with basso continuo",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.62 + i * 0.82, w: 9, h: 0.78, fontSize: 19, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.olive,
    });
  });
}

// ── SLIDE 9 · NAWM 83: Grandi O quam tu pulchra es ──────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("NAWM 83: Grandi — O quam tu pulchra es", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("格蘭第《妳多麼美麗》(1625)", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.45, w: 7, h: 0.04, fill: { color: C.gold } });

  const pts = [
    "Alessandro Grandi (1586–1630): Monteverdi's deputy at\nSt. Mark's, Venice; pioneered the solo motet with continuo",
    "Text from Song of Songs 雅歌: dialogue of lovers read as\nmetaphor for God's love for the Church",
    "Blends recitative style (\"O quam tu pulchra es\") with\nlyric aria in triple meter (\"Surge, propera, sponsa mea\")",
    "Contrasting styles within a single piece mirror the new\noperatic vocabulary applied to sacred expression",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.67 + i * 0.92, w: 9, h: 0.82, fontSize: 19, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.gold,
    });
  });

  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 5.0, w: 9.2, h: 0.4, fill: { color: C.olive }, rectRadius: 0.06 });
  s.addText("https://www.youtube.com/watch?v=KkDj_rJCk6E", {
    x: 0.5, y: 5.0, w: 9, h: 0.4, fontSize: 18, color: C.cream, fontFace: "Calibri", align: "center", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=KkDj_rJCk6E" },
  });
}

// ── SLIDE 10 · Music in Convents ─────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.olive); bottomBar(s, C.olive);

  s.addText("修道院中的音樂 Music in Convents", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 28, bold: true, color: C.olive, fontFace: "Georgia",
  });
  s.addText("Hidden Musical Cultures Behind Convent Walls", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.bronze, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.45, w: 9.2, h: 0.04, fill: { color: C.sand } });

  const pts = [
    "Church officials restricted music in convents, barring\nmale teachers; yet nuns developed thriving musical cultures",
    "Lucrezia Vizzana (1590–1662): entered Santa Cristina in\nBologna as a child; Componimenti musicali (1623) — motets",
    "Chiara Margarita Cozzolani (1602–ca. 1677): four collections\nof sacred concertos; polychoral Vespers with modern arias",
    "These women published despite restrictions, showing\nconvent music rivalled standards outside convent walls",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.62 + i * 0.82, w: 9, h: 0.78, fontSize: 19, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.olive,
    });
  });
}

// ── SLIDE 11 · Oratorio & Carissimi ──────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("神劇 Oratorio", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 38, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Sacred Drama Without Staging", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 24, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.45, w: 7, h: 0.04, fill: { color: C.gold } });

  const pts = [
    "Oratorio = sacred dramatic work combining narrative,\ndialogue, & commentary; NOT staged, with a narrator (testo)",
    "Italian oratorio 義大利語神劇: vernacular, shorter,\nfor spreading faith to commoners; performed during Lent",
    "Latin oratorio 拉丁語神劇: for aristocratic courts;\nlonger, more elaborate — like opera without costumes",
    "Named after the oratory (prayer hall) where devotional\nmeetings with music were held in 17th-century Rome",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.67 + i * 0.92, w: 9, h: 0.82, fontSize: 19, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.gold,
    });
  });
}

// ── SLIDE 12 · NAWM 84: Carissimi Jephte ────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.olive); bottomBar(s, C.olive);

  s.addText("NAWM 84: Carissimi — Historia di Jephte", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 30, bold: true, color: C.olive, fontFace: "Georgia",
  });
  s.addText("卡里西米《耶弗他》(ca. 1648)", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.bronze, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.45, w: 9.2, h: 0.04, fill: { color: C.sand } });

  const pts = [
    "Giacomo Carissimi (1605–1674): leading Latin oratorio\ncomposer; based on Judges 11:29–40 (Jephtha's vow)",
    "Narrator in recitative; Jephtha's victory recounted by\n6-voice ensemble with stile concitato battle effects",
    "Final scene: daughter laments with descending tetrachord\nbass; two sopranos echo her cadential flourishes",
    "Six-voice chorus responds with polychoral & madrigalistic\neffects — blending operatic drama with sacred purpose",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.62 + i * 0.82, w: 9, h: 0.78, fontSize: 19, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.olive,
    });
  });

  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 5.0, w: 9.2, h: 0.4, fill: { color: C.olive }, rectRadius: 0.06 });
  s.addText("https://www.youtube.com/watch?v=2s1Gf3b2wRs", {
    x: 0.5, y: 5.0, w: 9, h: 0.4, fontSize: 18, color: C.cream, fontFace: "Calibri", align: "center", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=2s1Gf3b2wRs" },
  });
}

// ── SLIDE 13 · Lutheran Church Music ─────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("路德宗教會音樂", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 38, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Lutheran Church Music: Italian Influence in Germany", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.45, w: 7, h: 0.04, fill: { color: C.gold } });

  const pts = [
    "German composers absorbed Italian monodic & concertato\ntechniques while maintaining polyphonic chorale traditions",
    "Sacred concerto in Germany: Hassler, Praetorius adopted\nthe large-scale Venetian model; small concerto even more common",
    "J. H. Schein: Opella nova (1618, 1626) — German sacred\nconcertos blending Lutheran chorale with Italian concertato",
    "Thirty Years' War (1618–48): devastated resources;\nforced smaller-scale works suited to reduced church forces",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.67 + i * 0.92, w: 9, h: 0.82, fontSize: 19, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.gold,
    });
  });
}

// ── SLIDE 14 · Heinrich Schütz: Life ─────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.olive); bottomBar(s, C.olive);

  s.addText("海因里希·許茨 Heinrich Schütz (1585–1672)", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.7, fontSize: 28, bold: true, color: C.olive, fontFace: "Georgia",
  });
  s.addText("Master of Conveying the Meaning of Words Through Music", {
    x: 0.4, y: 0.88, w: 9.2, h: 0.35, fontSize: 18, color: C.bronze, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.45, w: 9.2, h: 0.04, fill: { color: C.sand } });

  const pts = [
    "Innkeeper's son; singing talent noticed by Landgrave of\nHesse at age 12; sent to study law, then music in Venice",
    "Studied with Giovanni Gabrieli in Venice (1609–1612);\nvisited again during Monteverdi's era — absorbed both styles",
    "Chapelmaster at the Saxon court in Dresden (1615–1672);\nfirst German opera (Dafne, 1627, music lost)",
    "Thirty Years' War forced him to publish Kleine geistliche\nKonzerte (1636, 1639) for reduced forces — widely popular",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.62 + i * 0.82, w: 9, h: 0.78, fontSize: 19, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.olive,
    });
  });
}

// ── SLIDE 15 · Schütz: Major Works ───────────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("乖茨的主要作品 Schütz: Major Works", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Sacred Collections Spanning Five Decades", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.45, w: 7, h: 0.04, fill: { color: C.gold } });

  const pts = [
    "Psalmen Davids (1619): German polychoral psalms —\nVenetian grandeur with sensitive German text treatment",
    "Cantiones sacrae (1625): Latin motets with harmonic\nnovelties; Symphoniae sacrae I (1629): Italian concertato style",
    "Kleine geistliche Konzerte (1636, 1639): 1–5 voices +\ncontinuo only — composed for war-depleted church resources",
    "Symphoniae sacrae III (1650): large-scale concertos;\nHistoria (Seven Last Words), Christmas History, 3 Passions",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.67 + i * 0.92, w: 9, h: 0.82, fontSize: 19, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.gold,
    });
  });
}

// ── SLIDE 16 · NAWM 85: Schütz — Saul, Saul, was verfolgst du mich ─────────
{
  const s = lightSlide(pres);
  topBar(s, C.olive); bottomBar(s, C.olive);

  s.addText("NAWM 85: Schütz — Saul, Saul, was verfolgst du mich", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 28, bold: true, color: C.olive, fontFace: "Georgia",
  });
  s.addText("乖茨《掃羅，掃羅，你為何迫害我？》(1650)", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 20, color: C.bronze, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.45, w: 9.2, h: 0.04, fill: { color: C.sand } });

  const pts = [
    "From Symphoniae sacrae III (1650): large sacred concerto\nfor 2 choirs + 6 soloists + 2 violins + continuo",
    "Based on Acts 9 & 26: Christ's blinding voice to Saul\non the road to Damascus — a dramatic conversion scene",
    "Musical figures: cadentiae duriusculae (harsh dissonances)\nat \"Saul\"; saltus duriusculus (harsh leap) in solo section",
    "Combines Gabrieli's polychoral grandeur with Monteverdi's\ndissonant rhetoric — Italian form with German sacred purpose",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.62 + i * 0.82, w: 9, h: 0.78, fontSize: 19, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.olive,
    });
  });

  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 5.0, w: 9.2, h: 0.4, fill: { color: C.olive }, rectRadius: 0.06 });
  s.addText("https://www.youtube.com/watch?v=RCDI_Jy6bEk", {
    x: 0.5, y: 5.0, w: 9, h: 0.4, fontSize: 18, color: C.cream, fontFace: "Calibri", align: "center", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=RCDI_Jy6bEk" },
  });
}

// ── SLIDE 17 · NAWM 86: Schütz — O lieber Herre Gott ────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("NAWM 86: Schütz — O lieber Herre Gott", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("乖茨《親愛的主，上帝》— Kleine geistliche Konzerte", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 20, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.45, w: 7, h: 0.04, fill: { color: C.gold } });

  const pts = [
    "From Kleine geistliche Konzerte (1636): small sacred\nconcerto for solo voice + basso continuo only",
    "Wartime austerity: no instruments beyond continuo;\nyet Schütz achieves profound expression with minimal forces",
    "Musical figures capture every nuance of the German text;\nword-painting conveys meaning through melody & harmony",
    "Popular throughout Germany — suitable for small churches\nwith limited resources during the Thirty Years' War",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.67 + i * 0.92, w: 9, h: 0.82, fontSize: 19, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.gold,
    });
  });

  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 5.0, w: 9.2, h: 0.4, fill: { color: C.gold }, rectRadius: 0.06 });
  s.addText("https://www.youtube.com/watch?v=eZZMrP9bET4", {
    x: 0.5, y: 5.0, w: 9, h: 0.4, fontSize: 18, color: C.darkBg, fontFace: "Calibri", align: "center", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=eZZMrP9bET4" },
  });
}

// ── SLIDE 18 · Schütz Legacy ─────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.olive); bottomBar(s, C.olive);

  s.addText("乖茨的遺產 Schütz's Legacy", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 28, bold: true, color: C.olive, fontFace: "Georgia",
  });
  s.addText("Bridging Italian Innovation and German Sacred Tradition", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.bronze, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.45, w: 9.2, h: 0.04, fill: { color: C.sand } });

  const pts = [
    "Historia tradition: The Seven Last Words of Christ (1650s?)\nand Christmas History (1664) — narrative + concertato scenes",
    "Three Passions (1666): returned to the older German\ntradition of plainsong narrative + polyphonic motet choruses",
    "During his lifetime known mainly in Lutheran areas;\nrediscovered in the 19th–20th centuries as a towering figure",
    "His synthesis of German & Italian elements laid the\nfoundation from Schein to Bach to Brahms",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.62 + i * 0.82, w: 9, h: 0.78, fontSize: 19, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.olive,
    });
  });
}

// ── SLIDE 19 · Jewish Liturgical Music ───────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("猶太禮儀音樂 Jewish Liturgical Music", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Salamone Rossi & Polyphony in the Synagogue", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.45, w: 7, h: 0.04, fill: { color: C.gold } });

  const pts = [
    "Cantillation remained primary Jewish musical form;\nnew techniques emerged in early 1600s despite rabbinical resistance",
    "Leon Modena (1571–1648): rabbi & humanist; promoted\npolyphony at the Venice synagogue from 1607 onward",
    "Salamone Rossi (ca. 1570–ca. 1630): Mantuan composer;\nHashirim asher lish'lomo (Songs of Solomon, 1622–23)",
    "33 Hebrew psalms, hymns, synagogue songs in Italian\npolyphonic style — first published Jewish liturgical polyphony",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.67 + i * 0.92, w: 9, h: 0.82, fontSize: 19, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.gold,
    });
  });
}

// ── SLIDE 20 · Instrumental Music: Types & Categories ────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.olive); bottomBar(s, C.olive);

  s.addText("器樂類型 Types of Instrumental Music", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 28, bold: true, color: C.olive, fontFace: "Georgia",
  });
  s.addText("Four Ways to Categorize Baroque Instrumental Works", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.bronze, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.45, w: 9.2, h: 0.04, fill: { color: C.sand } });

  // Two-column comparison
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.6, w: 4.4, h: 3.7, fill: { color: C.sand }, rectRadius: 0.08 });
  s.addText("Improvisatory Style 即興風格", { x: 0.5, y: 1.65, w: 4.2, h: 0.35, fontSize: 20, bold: true, color: C.olive, fontFace: "Georgia" });
  s.addText("Toccata, fantasia, prelude\n觸技曲、幻想曲、前奏曲\n\nImitative Counterpoint 模仿對位\nRicercare, fugue, canzona\n利切卡爾、賦格、坎佐那\n\nVariations 變奏\nPartita, chaconne, passacaglia\n帕蒂塔、夏康、帕薩卡利亞", {
    x: 0.55, y: 1.85, w: 4.15, h: 3.3, fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "top",
  });

  s.addShape(pres.ShapeType.rect, { x: 5.2, y: 1.6, w: 4.4, h: 3.7, fill: { color: C.sand }, rectRadius: 0.08 });
  s.addText("Settings / Dance 舞曲", { x: 5.3, y: 1.65, w: 4.2, h: 0.35, fontSize: 20, bold: true, color: C.olive, fontFace: "Georgia" });
  s.addText("Organ verse, chorale prelude\n風琴聖詠前奏曲\n\nSonata 奏鳴曲\nContrasting sections for 1–2\nmelody instruments + continuo\n\nSuite 組曲\nPadouana, gagliarda, courante,\nallemande — linked dances", {
    x: 5.35, y: 1.85, w: 4.15, h: 3.3, fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "top",
  });
}

// ── SLIDE 21 · Frescobaldi: Keyboard Master ──────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("乎雷斯科巴第 Girolamo Frescobaldi (1583–1643)", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 26, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Organist at St. Peter's, Rome — Keyboard Genius", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.45, w: 7, h: 0.04, fill: { color: C.gold } });

  const pts = [
    "Born in Ferrara; organist at St. Peter's, Rome from 1608;\nalso served the Grand Duke of Tuscany & Barberini family",
    "Toccatas: succession of brief sections, each with a distinct\nfigure; tempo is free, not tied to a regular beat (affetti)",
    "Fiori musicali (Musical Flowers, 1635): 3 organ masses\nwith toccatas, ricercares, canzonas for liturgical use",
    "His keyboard works were a model for J. S. & C. P. E. Bach;\nFroberger (his student) spread the style across Europe",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.67 + i * 0.92, w: 9, h: 0.82, fontSize: 19, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.gold,
    });
  });
}

// ── SLIDE 22 · NAWM 87: Frescobaldi Toccata No. 3 ───────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.olive); bottomBar(s, C.olive);

  s.addText("NAWM 87: Frescobaldi — Toccata No. 3 (1615)", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 30, bold: true, color: C.olive, fontFace: "Georgia",
  });
  s.addText("乎雷斯科巴第《觸技曲第三號》", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.bronze, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.45, w: 9.2, h: 0.04, fill: { color: C.sand } });

  const pts = [
    "From Primo Libro di Toccate (First Book of Toccatas, 1615):\nfor harpsichord — succession of brief contrasting sections",
    "Each section focuses on a particular figure: virtuoso\npassagework, imitative exchanges, cadential weakening",
    "Sections can be played separately; player may end at any\ncadence — written music as a platform for performance",
    "Tempo varies with mood: \"not subject to a regular beat\nbut modified according to the affections\" (Frescobaldi)",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.62 + i * 0.82, w: 9, h: 0.78, fontSize: 19, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.olive,
    });
  });

  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 5.0, w: 9.2, h: 0.4, fill: { color: C.olive }, rectRadius: 0.06 });
  s.addText("https://www.youtube.com/watch?v=6dJmqSi_qlM", {
    x: 0.5, y: 5.0, w: 9, h: 0.4, fontSize: 18, color: C.cream, fontFace: "Calibri", align: "center", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=6dJmqSi_qlM" },
  });
}

// ── SLIDE 23 · Ricercare, Canzona, Sonata ────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("利切卡爾·坎佐那·奏鳴曲", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Ricercare, Canzona, and Sonata", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 24, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.45, w: 7, h: 0.04, fill: { color: C.gold } });

  const pts = [
    "Ricercare: serious imitative work for organ/harpsichord;\none subject developed continuously — ancestor of the fugue",
    "Canzona: imitative piece in contrasting sections;\nmarkedly rhythmic themes; for keyboard or ensemble",
    "Sonata 奏鳴曲: originally any instrumental piece; by\nmid-century = contrasting sections for 1–2 instruments + continuo",
    "By c. 1650 canzona and sonata had merged;\n\"sonata\" became the standard term for both",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.67 + i * 0.92, w: 9, h: 0.82, fontSize: 19, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.gold,
    });
  });
}

// ── SLIDE 24 · NAWM 88: Marini Sonata IV ─────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.olive); bottomBar(s, C.olive);

  s.addText("NAWM 88: Marini — Sonata IV, Op. 8 (1629)", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 30, bold: true, color: C.olive, fontFace: "Georgia",
  });
  s.addText("馬里尼《第四奏鳴曲》— 小提琴 + 數字低音", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 20, color: C.bronze, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.45, w: 9.2, h: 0.04, fill: { color: C.sand } });

  const pts = [
    "Biagio Marini (1594–1663): violinist at St. Mark's under\nMonteverdi; published 22 collections of vocal & instrumental music",
    "Sonata IV per il violino per sonar con due corde:\n\"instrumental monody\" — idiomatic violin writing with continuo",
    "Contrasting sections: expressive opening like a Caccini\nmadrigal → virtuosic figuration, double stops, large leaps",
    "Rhapsodic & metrical sections alternate; recitative and\naria styles adapted from vocal music to the violin",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.62 + i * 0.82, w: 9, h: 0.78, fontSize: 19, color: C.darkText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.olive,
    });
  });

  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 5.0, w: 9.2, h: 0.4, fill: { color: C.olive }, rectRadius: 0.06 });
  s.addText("https://www.youtube.com/watch?v=1OyyJb3WNGI", {
    x: 0.5, y: 5.0, w: 9, h: 0.4, fontSize: 18, color: C.cream, fontFace: "Calibri", align: "center", valign: "top",
    hyperlink: { url: "https://www.youtube.com/watch?v=1OyyJb3WNGI" },
  });
}

// ── SLIDE 25 · Variations, Chaconne & Suite ──────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("變奏·夏康·組曲", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 38, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("Variations, Chaconne / Passacaglia, and Suite", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.45, w: 7, h: 0.04, fill: { color: C.gold } });

  const pts = [
    "Variations 變奏: on borrowed or newly composed themes;\ncantus-firmus type, melodic embellishment, or harmonic bass",
    "Chaconne 夏康 & Passacaglia 帕薩卡利亞: variations over\na ground bass; triple meter, 4 bars; Frescobaldi's Partite (1627)",
    "Suite 組曲: linked dances sharing a key; J. H. Schein's\nBanchetto musicale (1617) — padouana, gagliarda, courante, etc.",
    "By 1700 chaconne and passacaglia merged as terms;\nsuites became a major genre of the late Baroque",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.67 + i * 0.92, w: 9, h: 0.82, fontSize: 19, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.gold,
    });
  });
}

// ── SLIDE 26 · Timeline ──────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.olive); bottomBar(s, C.olive);

  s.addText("年表 Timeline", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 28, bold: true, color: C.olive, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.78, w: 9.2, h: 0.04, fill: { color: C.sand } });

  const events = [
    ["1600", "Peri, L'Euridice — first opera whose music survives"],
    ["1602", "Viadana, Cento concerti ecclesiastici (first printed sacred + bc)"],
    ["1605", "Monteverdi, Fifth Book of Madrigals (basso continuo)"],
    ["1615", "Frescobaldi, First Book of Toccatas"],
    ["1618–48", "Thirty Years' War devastates Germany"],
    ["1619", "Schütz, Psalmen Davids; Monteverdi, Book 7"],
    ["1622–23", "Salamone Rossi, Hashirim asher lish'lomo"],
    ["1625", "Grandi, O quam tu pulchra es"],
    ["1629", "Marini, Sonate, Op. 8; Schütz, Symphoniae sacrae I"],
    ["1635", "Frescobaldi, Fiori musicali (organ masses)"],
    ["1636–39", "Schütz, Kleine geistliche Konzerte (I & II)"],
    ["ca. 1648", "Carissimi, Historia di Jephte"],
    ["1650", "Schütz, Symphoniae sacrae III (Saul)"],
    ["1659", "Strozzi, Diporti di Euterpe (Lagrime mie)"],
  ];

  events.forEach(([yr, desc], i) => {
    const y = 0.92 + i * 0.32;
    const bgColor = i % 2 === 0 ? C.sand : "EDE5CC";
    s.addShape(pres.ShapeType.rect, { x: 0.4, y, w: 9.2, h: 0.3, fill: { color: bgColor } });
    s.addText(yr, { x: 0.45, y, w: 1.3, h: 0.3, fontSize: 18, bold: true, color: C.olive, fontFace: "Georgia", valign: "middle" });
    s.addText(desc, { x: 1.8, y, w: 7.7, h: 0.3, fontSize: 18, color: C.darkText, fontFace: "Calibri", valign: "middle" });
  });
}

// ── SLIDE 27 · Tradition and Innovation ──────────────────────────────────────
{
  const s = darkSlide(pres);
  topBar(s, C.gold); bottomBar(s, C.gold);

  s.addText("傳統與創新 Tradition and Innovation", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.9, fontSize: 28, bold: true, color: C.gold, fontFace: "Georgia",
  });
  s.addText("The Lasting Significance of Early 17th-Century Music", {
    x: 0.4, y: 1.08, w: 9.2, h: 0.35, fontSize: 22, color: C.sand, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 1.5, y: 1.45, w: 7, h: 0.04, fill: { color: C.gold } });

  const pts = [
    "New genres: cantata, sacred concerto, oratorio, sonata,\npartita, chaconne, passacaglia, dance suite",
    "New techniques: basso continuo, concertato medium,\nground bass — all became foundational for later Baroque",
    "Recognition that different styles suit different purposes:\nstile antico for sacred tradition, stile moderno for expression",
    "Instrumental music gained independence from vocal music;\ntoccatas & ricercares compared to orations in rhetorical power",
  ];
  pts.forEach((txt, i) => {
    s.addText(txt, {
      x: 0.5, y: 1.67 + i * 0.92, w: 9, h: 0.82, fontSize: 19, color: C.lightText, fontFace: "Calibri", valign: "top",
      bullet: true, bulletColor: C.gold,
    });
  });
}

// ── SLIDE 28 · Key Terms & Listening ─────────────────────────────────────────
{
  const s = lightSlide(pres);
  topBar(s, C.olive); bottomBar(s, C.olive);

  s.addText("關鍵術語與聆聽 Key Terms & Listening", {
    x: 0.4, y: 0.15, w: 9.2, h: 0.75, fontSize: 26, bold: true, color: C.olive, fontFace: "Georgia",
  });
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 0.95, w: 9.2, h: 0.03, fill: { color: C.sand } });

  // Left column: Key Terms
  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.08, w: 4.4, h: 4.28, fill: { color: C.sand }, rectRadius: 0.08 });
  s.addText("Key Terms 關鍵術語", { x: 0.5, y: 1.15, w: 4.2, h: 0.32, fontSize: 18, bold: true, color: C.olive, fontFace: "Georgia" });
  s.addText(
    "Concertato madrigal 協奏曲風牧歌\n" +
    "Basso ostinato / ground bass 固定低音\n" +
    "Descending tetrachord 下行四度\n" +
    "Cantata 清唱套曲\n" +
    "Sacred concerto 宗教協奏曲\n" +
    "Oratorio 神劇\n" +
    "Stile antico / stile moderno\n" +
    "Toccata 觸技曲\n" +
    "Ricercare 利切卡爾 / Fugue 賦格\n" +
    "Canzona 坎佐那 / Sonata 奏鳴曲\n" +
    "Chaconne 夏康 / Passacaglia\n" +
    "Suite 組曲 / Musical figures 音樂修辭",
    { x: 0.55, y: 1.55, w: 4.15, h: 3.75, fontSize: 14, color: C.darkText, fontFace: "Calibri", valign: "top", lineSpacingMultiple: 1.2 }
  );

  // Right column: NAWM Listening
  s.addShape(pres.ShapeType.rect, { x: 5.2, y: 1.08, w: 4.4, h: 4.28, fill: { color: C.sand }, rectRadius: 0.08 });
  s.addText("NAWM Listening 聆聽清單", { x: 5.3, y: 1.15, w: 4.2, h: 0.32, fontSize: 18, bold: true, color: C.olive, fontFace: "Georgia" });
  const links = [
    ["82  Monteverdi, Ohimè dov'è il mio ben", "https://www.youtube.com/watch?v=bJfHoxZlMYw"],
    ["83  Grandi, O quam tu pulchra es", "https://www.youtube.com/watch?v=KkDj_rJCk6E"],
    ["84  Carissimi, Historia di Jephte", "https://www.youtube.com/watch?v=2s1Gf3b2wRs"],
    ["85  Schütz, Saul, was verfolgst du mich", "https://www.youtube.com/watch?v=RCDI_Jy6bEk"],
    ["86  Schütz, O lieber Herre Gott", "https://www.youtube.com/watch?v=eZZMrP9bET4"],
    ["87  Frescobaldi, Toccata No. 3", "https://www.youtube.com/watch?v=6dJmqSi_qlM"],
    ["88  Marini, Sonata IV, Op. 8", "https://www.youtube.com/watch?v=1OyyJb3WNGI"],
  ];
  links.forEach(([label, url], i) => {
    s.addText(label, {
      x: 5.35, y: 1.55 + i * 0.5, w: 4.1, h: 0.45, fontSize: 14, color: C.olive, fontFace: "Calibri", valign: "top",
      hyperlink: { url },
    });
  });
}

// ── Save ─────────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "Ch15_Chamber_Church.pptx" })
  .then(() => console.log("Ch15_Chamber_Church.pptx created"))
  .catch(e => console.error(e));
