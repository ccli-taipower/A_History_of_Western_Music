'use strict';
const pptx = require('pptxgenjs');
const pres = new pptx();
pres.layout = 'LAYOUT_16x9';

// ── Period color palettes (1-based) ───────────────────────────────────────────
const P = [
  null,
  { darkBg:"1A0F0A", panel:"2A1A10", gold:"C09050", ivory:"F5ECD5", sand:"D4B888", lightBg:"F0E8D5", lp:"E0D4C0", dt:"3A2A18" }, // 1
  { darkBg:"0F1520", panel:"182030", gold:"8090C8", ivory:"D8DCF0", sand:"A8B0D8", lightBg:"E8EAF8", lp:"D0D4F0", dt:"1A2040" }, // 2
  { darkBg:"0F1A10", panel:"182818", gold:"80B870", ivory:"D8F0D0", sand:"A8D098", lightBg:"E4F5E0", lp:"CDE8C8", dt:"182818" }, // 3
  { darkBg:"1A1408", panel:"2A2010", gold:"C8A850", ivory:"F5ECD5", sand:"D8C080", lightBg:"F5EDDC", lp:"E8E0C8", dt:"3A2A10" }, // 4
  { darkBg:"180C18", panel:"281428", gold:"B080C0", ivory:"EDD8F0", sand:"C8A0D8", lightBg:"F0E4F8", lp:"E0D0EC", dt:"281428" }, // 5
  { darkBg:"0E1A18", panel:"182820", gold:"78A890", ivory:"D0E8E4", sand:"98C8B8", lightBg:"E0F0EC", lp:"C8E4DC", dt:"183028" }, // 6
  { darkBg:"1A0C10", panel:"281418", gold:"C07888", ivory:"F0D8DE", sand:"D8A0B0", lightBg:"F5E4E8", lp:"E8D0D8", dt:"381018" }, // 7
  { darkBg:"101818", panel:"182828", gold:"70A8A8", ivory:"D0E8E8", sand:"90C0C0", lightBg:"E0F0F0", lp:"C8E4E4", dt:"183030" }, // 8
  { darkBg:"121820", panel:"1C2830", gold:"8890B0", ivory:"D8DCF0", sand:"A8ACD0", lightBg:"E8EAF8", lp:"D0D4EC", dt:"1C2838" }, // 9
];

// ── Helpers ───────────────────────────────────────────────────────────────────
function ds(C) { const s = pres.addSlide(); s.background = { color: C.darkBg }; return s; }
function ls(C) { const s = pres.addSlide(); s.background = { color: C.lightBg }; return s; }
function bars(s, C) {
  s.addShape(pres.ShapeType.rect, { x:0, y:0,    w:10, h:0.12, fill:{color:C.gold}, line:{color:C.gold} });
  s.addShape(pres.ShapeType.rect, { x:0, y:5.50, w:10, h:0.12, fill:{color:C.gold}, line:{color:C.gold} });
}
function hdr(s, C, zh, en, dark=true) {
  bars(s, C);
  s.addText(zh, { x:0.4, y:0.18, w:9.2, h:0.55, fontSize:26, bold:true, color:dark?C.gold:C.dt, fontFace:"Georgia", align:"center" });
  s.addText(en, { x:0.4, y:0.68, w:9.2, h:0.32, fontSize:14, color:dark?C.sand:C.dt, fontFace:"Georgia", align:"center", italic:true });
}
function dp(s, C, lT, lB, rT, rB) {
  s.addShape(pres.ShapeType.rect, { x:0.3, y:1.30, w:4.6, h:4.1, fill:{color:C.panel}, line:{color:C.panel} });
  s.addText("■ "+lT, { x:0.45, y:1.38, w:4.3, h:0.4, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia", margin:0 });
  s.addText(lB, { x:0.5, y:1.70, w:4.35, h:3.65, fontSize:14, color:C.ivory, fontFace:"Calibri", valign:"top", paraSpaceAfter:0 });
  s.addShape(pres.ShapeType.rect, { x:5.1, y:1.30, w:4.6, h:4.1, fill:{color:C.panel}, line:{color:C.panel} });
  s.addText("■ "+rT, { x:5.25, y:1.38, w:4.3, h:0.4, fontSize:14, bold:true, color:C.gold, fontFace:"Georgia", margin:0 });
  s.addText(rB, { x:5.3, y:1.70, w:4.35, h:3.65, fontSize:14, color:C.ivory, fontFace:"Calibri", valign:"top", paraSpaceAfter:0 });
}
function lp(s, C, lT, lB, rT, rB) {
  s.addShape(pres.ShapeType.rect, { x:0.3, y:1.30, w:4.6, h:4.1, fill:{color:C.lp}, line:{color:C.lp} });
  s.addText("■ "+lT, { x:0.45, y:1.38, w:4.3, h:0.4, fontSize:14, bold:true, color:C.dt, fontFace:"Georgia", margin:0 });
  s.addText(lB, { x:0.5, y:1.70, w:4.35, h:3.65, fontSize:14, color:C.dt, fontFace:"Calibri", valign:"top", paraSpaceAfter:0 });
  s.addShape(pres.ShapeType.rect, { x:5.1, y:1.30, w:4.6, h:4.1, fill:{color:C.lp}, line:{color:C.lp} });
  s.addText("■ "+rT, { x:5.25, y:1.38, w:4.3, h:0.4, fontSize:14, bold:true, color:C.dt, fontFace:"Georgia", margin:0 });
  s.addText(rB, { x:5.3, y:1.70, w:4.35, h:3.65, fontSize:14, color:C.dt, fontFace:"Calibri", valign:"top", paraSpaceAfter:0 });
}
function nawm(C, num, composer, work, info, guide, url) {
  const s = ds(C); hdr(s, C, num+"  "+composer, work);
  dp(s, C, "作品資訊 Work Info", info, "聆聽重點 Listening Guide", guide+"\nyoutu.be/"+url);
}
function cover(C, n, zh, en, range, chs) {
  const s = ds(C); bars(s, C);
  s.addText("第 "+n+" 時期  Period "+n, { x:0.4, y:0.90, w:9.2, h:0.45, fontSize:18, color:C.sand, fontFace:"Georgia", align:"center" });
  s.addText(zh, { x:0.4, y:1.35, w:9.2, h:0.90, fontSize:38, bold:true, color:C.gold, fontFace:"Georgia", align:"center" });
  s.addText(en, { x:0.4, y:2.25, w:9.2, h:0.55, fontSize:22, color:C.ivory, fontFace:"Georgia", align:"center", italic:true });
  s.addText(range, { x:0.4, y:2.90, w:9.2, h:0.40, fontSize:18, color:C.sand, fontFace:"Calibri", align:"center" });
  s.addText(chs,  { x:0.4, y:3.38, w:9.2, h:0.35, fontSize:14, color:C.sand, fontFace:"Calibri", align:"center" });
}

// ══════════════════════════════════════════════════════════════════════════════
// PERIOD 1 · 古代與教會音樂 (11 slides)
// ══════════════════════════════════════════════════════════════════════════════
{ const C = P[1];
cover(C, 1, "古代與教會音樂",
  "Music in Antiquity and the Early Church",
  "ca. 3000 BCE – 900 CE",
  "涵蓋 Ch01–03  ·  Chapters 1–3");

{ const s=ds(C); hdr(s,C,"時代背景","Historical Context");
dp(s,C,"政治社會背景 Political Context",
`· 古希臘城邦（Polis）文明興盛
· 伯里克利黃金時代（5c BCE）
· 希臘戲劇與音樂緊密結合
· 羅馬帝國擴張，吸納希臘文化
· 380 CE 基督教成羅馬官方宗教
· 395 CE 羅馬帝國東西分裂
· 修道院文化成教育書寫中心
· 聖本篤修道院規則（529 CE）
· 查理曼大帝加冕（800 CE）
· 卡洛林文藝復興整理聖詠
· 書寫文化：羊皮紙手稿昂貴
· 音樂 = 自由七藝 quadrivium`,
"音樂文化概況 Music Culture",
`· 音樂：哲學、數學、宇宙論三位一體
· 古希臘：八種調式體系建立
· 希臘悲劇：合唱隊是音樂中心
· 早期教會：多種禮儀傳統並存
· Ambrose 讚美詩傳統（4c）
· 格里高利整合各地聖詠（600 CE）
· 最早記譜符號：紐瑪 neumes
· 紐瑪無固定音高，依賴記憶傳唱
· Guido d'Arezzo 發明線譜（1025）
· 六聲音階 hexachord 體系建立
· 手型記憶法 Guidonian Hand
· Solmization：ut-re-mi-fa-sol-la`); }

{ const s=ds(C); hdr(s,C,"核心作曲家","Key Composers and Theorists");
dp(s,C,"古代 Antiquity",
`· Enheduanna（ca. 2300 BCE）
   蘇美爾祭司，最早有名作曲者
· Seikilos（1st c. CE）
   Epitaph，現存最完整古代樂譜
· Pythagoras（ca. 570–495 BCE）
   音程比例理論：弦長整數比
· Plato（427–347 BCE）
   調式道德論，音樂教育哲學
· Aristotle（384–322 BCE）
   音樂欣賞與倫理，《政治學》
· Boethius（480–524 CE）
   《音樂論》傳承希臘理論至中世紀`,
"教會 Church",
`· St. Ambrose（340–397 CE）
   米蘭讚美詩 Hymn 傳統奠基
· Pope Gregory I（590–604）
   整合羅馬禮儀聖詠典集
· Isidore of Seville（560–636）
   《語源學》分類音樂知識
· Hucbald（840–930）
   最早 Organum 複音論文作者
· Guido d'Arezzo（990–1050）
   五線譜、音節唱名法、教學法
· Hildegard von Bingen（1098–1179）
   修女神學家，72 首聖詠作品`); }

{ const s=ds(C); hdr(s,C,"代表體裁與形式","Genres and Forms");
dp(s,C,"禮儀唱段 Liturgical Chant",
`· Introit 進堂詠
· Kyrie 垂憐歌（Syllabic 風格）
· Gloria 光榮頌
· Gradual 升階詠（Melismatic）
· Alleluia 哈利路亞
· Tract 繼敘詠（四旬期替代）
· Sequence 繼應歌
· Antiphon 交唱詠
· Responsory 應答詠
· Psalm tone 詩篇音調
· Hymn 讚美詩（固定節律）
· Office Hours 時辰禮儀諸段`,
"記譜與理論 Notation & Theory",
`· 古希臘字母記譜法
· 紐瑪 Neumes（僅示方向）
· 導向紐瑪：高低弧度暗示音高
· 四線譜（Guido 前身）
· 五線譜 Staff notation（1025）
· C 譜號、F 譜號確定音高
· Hexachord 六聲音階體系
· Guidonian Hand 手型背誦法
· Solmization 唱名記憶
· 8 Church Modes 八種調式
· 正格 Authentic + 副格 Plagal
· Trope 擴展聖詠的插入段`); }

{ const s=ds(C); hdr(s,C,"風格特徵","Style Features");
dp(s,C,"旋律與節奏 Melody & Rhythm",
`· 單聲部 Monophony（一條旋律）
· 無固定拍號 Free rhythm
· 逐步進行為主 stepwise motion
· 小跳：三度四度為輔助
· Melisma 花唱（一音節數十音）
· Syllabic 一字一音（清楚詞義）
· Neumatic 介於兩者之間
· 旋律弧線：緩升急落形態
· 調式色彩各異，終止音固定
· 八種教會調式：D E F G 起始
· 各調式有特徵音（Leading tone）
· 無半音導音的調性感`,
"織體與功能 Texture & Function",
`· 純人聲演唱，無器樂伴奏
· 宗教禮儀功能絕對優先
· 靜穆冥想，超越時間感
· 修道院男聲演唱為主
· 問答式 Responsorial（獨唱+合唱）
· 交替式 Antiphonal（兩組交替）
· 齊唱 Unison 最常見方式
· 莊重緩慢，冥想性格
· 歌詞拉丁文，禮儀統一語言
· 禮儀年曆決定每日演唱曲目
· 聖詠集 Graduale 等典籍
· 口耳相傳與記譜記法並存`); }

nawm(C,"NAWM 1","Seikilos","Epitaph of Seikilos · ca. 1st c. CE",
`· 出土地：土耳其 Aidin 附近
· 刻於石柱 stele（大理石柱）
· 含旋律、歌詞、節奏三者齊全
· 希臘字母記譜法，清晰完整
· 弗里吉亞調式 Phrygian
· 全曲 10 個音節，短小精巧
· 節奏用希臘詩律長短音符標示
· 現存最古老完整的音樂作品
· 主題：享受人生，勿悲愁
· 旋律流暢，弧線優美
· 三個樂句對應三行詩
· 羅馬帝國時期希臘文化縮影`,
`· 注意旋律輪廓的平滑進行
· 感受弗里吉亞調式的獨特色彩
· 找出最高音落在哪個音節
· 對比現代五線譜記法差異
· 思考：古代記譜保存了什麼
· 歌詞傳達：光陰短暫，活在當下
· 注意節奏是否有規律可循
· 感受單聲部的純粹與直接
· 辨別希臘記號中的升降音
· 想像刻在石柱上的歷史感
· 與格里高利聖詠比較音響`,
"8Vkcolt-nmU");

nawm(C,"NAWM 3","Gregorian Chant","Viderunt omnes · Christmas Gradual",
`· 耶誕日彌撒 Gradual 回應詠
· 格里高利聖詠最具代表性體裁
· 單聲部、男聲合唱、無伴奏
· 方形紐瑪 square neumes 記譜
· Dorian 調式（終止音 D）
· 詩篇 98:3 拉丁文歌詞
· Respond＋Verse 兩段結構
· Melisma 花唱密集出現
· 一個音節可唱數十個音符
· 旋律片段源自古老 Psalm tone
· Melismatic 聖詠最高典範
· 耶誕節彌撒的核心儀式唱段`,
`· 注意「Viderunt」的超長花唱
· 感受無固定拍號的自由流動感
· 辨認 Respond 與 Verse 兩段
· 對比 Kyrie 的 Syllabic 風格
· 思考：花唱如何服務禮儀功能
· 感受大教堂音響的神聖空間感
· 注意旋律的波浪弧線起伏
· 感受 Dorian 調式獨特色彩
· 找出最密集的花唱音節
· 思考功能性音樂 vs. 藝術性
· 對比 Seikilos 的記譜方式`,
"uvC0NGmpuFY");

{ const s=ls(C); hdr(s,C,"重要術語 · Period 1","Key Terms: Antiquity and Early Church",false);
lp(s,C,"術語 Terms A",
`· Plainchant / Plainsong 素歌
· Gregorian Chant 格里高利聖詠
· Neumes 紐瑪記譜符號
· Church Mode 教會調式（8種）
· Dorian / Phrygian 等調式名
· Hexachord 六聲音階體系
· Melisma 一音節多音花唱
· Syllabic 一音節一音對應
· Neumatic 介於二者之間
· Monophony 單聲部音樂
· Free rhythm 自由節奏
· Quadrivium 自由七藝數學組`,
"術語 Terms B",
`· Responsorial 應答式演唱
· Antiphonal 交替式演唱
· Cantus planus 素歌（拉丁文）
· Solmization 唱名記憶法
· Guidonian Hand 手型背誦法
· Psalm tone 詩篇詠唱音調
· Gradual / Alleluia 彌撒唱段
· Trope 聖詠插入擴展段
· Organum 早期複音外加聲部
· Sequence 繼應歌
· Staff notation 線譜記法
· Authentic / Plagal 正格/副格`); }

{ const s=ls(C); hdr(s,C,"時間軸 · Period 1","Timeline: Antiquity – Early Church",false);
lp(s,C,"古代 Antiquity",
`· ca. 2300 BCE  Enheduanna，蘇美爾
· ca. 1400 BCE  Hurrian Hymn，敘利亞
· ca. 750 BCE   荷馬史詩 Iliad/Odyssey
· ca. 590 BCE   Pythagoras 音程比例論
· ca. 550 BCE   希臘悲劇黃金時代
· 458 BCE       Aeschylus 三部曲
· 347 BCE       Plato 去世，《理想國》
· ca. 1st c. CE Seikilos Epitaph 刻製
· 313 CE        君士坦丁頒布米蘭詔書
· 380 CE        基督教成羅馬官方宗教
· 395 CE        羅馬帝國東西分裂
· 480–524 CE    Boethius，《音樂論》`,
"中世紀前期 Early Medieval",
`· 529 CE  聖本篤修道院規則建立
· 590–604  Gregory I 整合禮儀聖詠
· ca. 700  羅馬聖詠傳入英格蘭
· ca. 750  西歐修道院文化確立
· 800      查理曼大帝加冕皇帝
· ca. 850  Musica enchiriadis 複音論
· ca. 900  最早 Organum 文獻記錄
· ca. 990  Guido d'Arezzo 誕生
· ca. 1025 Guido 發明五線譜系統
· 1098     Hildegard von Bingen 誕生
· ca. 1100 格里高利聖詠正典確立
· 1179     Hildegard 去世，72 首聖詠`); }

{ const s=ls(C); hdr(s,C,"考試重點 · Period 1","Exam Focus: Antiquity and Early Church",false);
lp(s,C,"必記作曲家/作品",
`· Seikilos · Epitaph（NAWM 1）
   現存最早完整音樂作品
· Gregorian Chant · Viderunt（NAWM 3）
   花唱典範，Gradual 代表
· Hildegard von Bingen
   最重要中世紀女作曲家
· Guido d'Arezzo · Ut queant laxis
   六音節唱名法來源歌曲
· Boethius · De institutione musica
   中世紀音樂理論標準教材
· 格里高利聖詠集 Graduale Romanum`,
"必懂概念",
`· 單聲部→複音的歷史演進邏輯
· 8種教會調式：Dorian/Phrygian
   Lydian/Mixolydian 各有
   正格（Authentic）副格（Plagal）
· Melisma vs. Syllabic 的區別
   與在禮儀中的不同功能
· 格里高利聖詠的禮儀分類
   彌撒唱段 vs. 時辰禮唱段
· Guido 線譜的歷史重要性
   → 音高精確傳播的革命
· 音樂 = quadrivium 的意涵`); }

{ const s=ds(C); hdr(s,C,"跨期比較","Connecting to Later Periods");
dp(s,C,"古代影響 Legacy",
`· 調式體系成複音音樂和聲基礎
· 格里高利旋律成 cantus firmus
   文藝復興彌撒的核心借用素材
· Guido 線譜 → 現代五線譜直系
· Hexachord solmization
   使用至 17 世紀才被取代
· Hildegard = 有記錄最早女作曲家
· 禮儀形式 Kyrie Gloria Credo
   延續至巴洛克安魂曲和彌撒
· Boethius 著作→中世紀大學
   音樂理論課程標準教材
· 八種調式 → 後來衍生大小調`,
"與下一時期連結 Bridge",
`· 9–10 世紀 Organum 興起
   → 中世紀複音音樂的開端
· 修道院→大學→大教堂
   → 巴黎聖母院成音樂中心
· Cantus firmus 技法建立
   → 將延續 600 年至巴洛克
· 從單一旋律 → 複數聲部
   → 對位藝術的緩慢誕生
· Trope→Sequence→Motet
   → 教堂音樂類型演化鏈
· Hildegard 的 Ordo Virtutum
   → 音樂劇場的最早雛形`); }
}

// ══════════════════════════════════════════════════════════════════════════════
// PERIOD 2 · 中世紀 (11 slides)
// ══════════════════════════════════════════════════════════════════════════════
{ const C = P[2];
cover(C, 2, "中世紀",
  "The Middle Ages",
  "ca. 900 – 1420",
  "涵蓋 Ch04–06  ·  Chapters 4–6");

{ const s=ds(C); hdr(s,C,"時代背景","Historical Context");
dp(s,C,"政治社會背景 Political Context",
`· 封建制度、莊園經濟確立
· 騎士文化與典雅愛情文學
· 十字軍東征（1096–1270）
· 哥德式大教堂建築運動
· 大學興起：巴黎（1150）、牛津
· 英法百年戰爭（1337–1453）
· 黑死病（1347–51）：三分之一人口
· 神聖羅馬帝國諸侯割據
· 教廷分裂 Schism（1378–1417）
· 拜占庭帝國末期（→1453）
· 蒙古西征衝擊歐洲（13c）
· 文藝復興的萌芽在義大利`,
"音樂文化概況 Music Culture",
`· Notre Dame 大教堂成音樂中心
· 巴黎聖母院樂派建立複音傳統
· 世俗歌謠：南法 Troubadour
· 北法 Trouvère，德語 Minnesang
· 複音音樂從教堂向世俗擴張
· Ars Antiqua 古藝術（12–13c）
· Ars Nova 新藝術革命（14c）
· 義大利 Trecento 世俗繁榮
· Ars Subtilior 極度複雜（14c末）
· 音樂手稿：Montpellier Codex
· 樂器：維愛爾、舒姆、管風琴
· 樂人從修道院走向宮廷世俗`); }

{ const s=ds(C); hdr(s,C,"核心作曲家","Key Composers");
dp(s,C,"複音與教堂 Polyphony",
`· Léonin（ca. 1135–1201）
   Magnus Liber Organi，二聲部
· Pérotin（ca. 1160–1230）
   四聲部複音，Viderunt omnes
· Philippe de Vitry（1291–1361）
   Ars Nova 論文，等分節拍理論
· Guillaume de Machaut（1300–77）
   14c 最偉大，彌撒+世俗俱全
· Johannes Ciconia（1370–1412）
   勃艮第與義大利風格橋樑
· John Dunstaple（1390–1453）
   英格蘭，三度六度協和先驅`,
"世俗與義大利 Secular/Italy",
`· Adam de la Halle（1245–1306）
   法國，Robin et Marion
· Troubadour 典型（佚名眾多）
   Countess of Dia 女遊吟詩人
· Walther von der Vogelweide
   （1170–1230）德語 Minnesang
· Jacopo da Bologna（fl.1340s）
   義大利 Trecento 先驅
· Francesco Landini（1325–1397）
   盲目管風琴師，140 首 Ballata
· Philippus de Caserta（fl.1370s）
   Ars Subtilior 最複雜代表`); }

{ const s=ds(C); hdr(s,C,"代表體裁與形式","Genres and Forms");
dp(s,C,"複音 Polyphony",
`· Organum purum（花唱上聲部）
· Discant（對節奏複音）
· Conductus（齊步進行曲）
· Motet（拉丁、法文混用）
   上方聲部有獨立歌詞
· Isorhythmic Motet（14c 主流）
   color + talea 雙重循環
· Polyphonic Mass（Machaut 首創）
   四聲部完整彌撒常規
· Hocket（鋸齒狀交替音符）`,
"世俗 Secular",
`· Troubadour / Trouvère 歌謠
· Minnesang 德語宮廷愛情詩
· Ballade（7行詩，AaB 疊歌）
· Rondeau（8行，ABaAabAB）
· Virelai（AbbaA 舞曲形）
   以上三種合稱 Formes fixes
· Madrigal（14c 義大利，2–3聲）
· Caccia 狩獵歌（卡農式）
· Ballata（舞曲，Landini 擅長）`); }

{ const s=ds(C); hdr(s,C,"風格特徵","Style Features");
dp(s,C,"和聲與節奏 Harmony & Rhythm",
`· Ars Antiqua：模仿節奏模式
   六種 Rhythmic Modes
· Mensural notation 量化記譜
   等分音符，標記時間長短
· 完全協和優先：八度、五度、四度
· 三度六度逐漸被接受協和音
· 終止式：雙導音 double leading tone
· Isorhythm：色彩＋節律雙循環
· Ars Subtilior：極複雜節奏標記
   分數拍號，超前衛記譜
· Cross-rhythm 多聲部節奏交叉
· 模進 Sequence 和聲動力`,
"織體與音色 Texture & Timbre",
`· 二至四聲部複音主流
· 各聲部可唱不同語言文字
   （多語 Polytextual Motet）
· Tenor 借自聖詠，長音持續
· 上方聲部活躍，花腔裝飾
· Hocket：聲部輪流靜止/發聲
· 器樂可替換人聲聲部
· 早期平行五八度（Organum）
· 逐漸避免平行 → 獨立進行
· 世俗聲樂：可加魯特琴伴奏`); }

nawm(C,"NAWM 25a","Machaut","Messe de Nostre Dame · Kyrie · ca. 1360s",
`· 史上第一首有名的完整複音彌撒
· 四聲部：Triplum/Motetus/Tenor/Contratenor
· Tenor 借自聖詠 Cunctipotens
· Isorhythmic 結構貫穿全曲
· Kyrie I–Christe–Kyrie II 三段
· 各段有獨立的 color/talea 設計
· 雙導音終止式貫穿全曲
· 和聲濃厚，14 世紀獨特音響
· 六月天我見」Kyrie eleison
· 表達垂憐的凝重莊嚴氣氛
· 對位嚴謹，聲部高度獨立
· 標誌複音彌撒作為獨立體裁`,
`· 注意 tenor 的緩慢長音持續
· 感受 isorhythm 的骨架結構
· 辨認 color 旋律循環出現
· 注意 hocket 的鋸齒效果
· 聽雙導音終止式的解決方式
· 思考：彌撒如何統一六個常規段
· 感受 14 世紀複音的獨特音響
· 對比格里高利聖詠的純粹
· 注意上方聲部的活躍裝飾
· 思考：為何需要 isorhythm？
· 對比文藝復興彌撒的和諧感`,
"JtFMfmG5VlY");

nawm(C,"NAWM 26","Machaut","Douce dame jolie · Virelai · ca. 1350",
`· Virelai 定型歌：AbbaA 結構
· 單聲部，可加器樂即興伴奏
· 典雅愛情題材 courtly love
· 旋律甜美流暢，弧形上揚線條
· 活潑節奏，明確舞曲性格
· 中世紀法語歌詞（古法文）
· 詩中痛苦求愛的騎士口吻
· 返唱 A 段帶來明確回歸感
· Machaut 最著名世俗歌曲
· 對比彌撒的宗教沉重感
· 可純器樂演奏的靈活性
· Virelai = 中世紀最活潑體裁`,
`· 注意旋律如何重複 AbbaA 架構
· 感受舞曲性的節奏推動力
· 對比複音 Messe 的厚重感
· 思考：同一作曲家宗教/世俗對比
· 聆聽旋律如何表達痛苦求愛
· 注意法語歌詞與旋律節律關係
· 感受 A 段返回的清晰結構感
· 對比三種 Formes fixes 的差異
· 辨認 b 段與 a 段的調性差異
· 想像宮廷宴席上的演唱場景
· 思考器樂即興的可能方式`,
"pSjXxAOkSM8");

{ const s=ls(C); hdr(s,C,"重要術語 · Period 2","Key Terms: The Middle Ages",false);
lp(s,C,"術語 Terms A",
`· Organum 奧爾加農（早期複音）
· Discant 定量複音對位法
· Motet 經文歌（多語混用）
· Isorhythm 等節奏技法
· Color 色彩（旋律循環段）
· Talea 節律（節奏循環段）
· Cantus firmus 素歌骨幹聲部
· Hocket 交替音符技法
· Formes fixes 定型歌三種
· Ars Antiqua 古藝術（12–13c）
· Ars Nova 新藝術（14c）
· Ars Subtilior 精緻藝術（14c末）`,
"術語 Terms B",
`· Mensural notation 量化記譜
· Rhythmic modes 節奏模式
· Conductus 進行曲體裁
· Troubadour 南法遊吟詩人
· Trouvère 北法遊吟詩人
· Minnesang 德語宮廷愛情詩
· Ballade / Rondeau / Virelai
· Trecento 義大利 14 世紀音樂
· Double leading tone 雙導音
· Tenor 聖詠持續聲部
· Polyphonic Mass 複音彌撒
· Motetus 經文歌中間聲部`); }

{ const s=ls(C); hdr(s,C,"時間軸 · Period 2","Timeline: The Middle Ages",false);
lp(s,C,"9–13 世紀",
`· ca. 900   Organum 最早文獻記錄
· 1096       第一次十字軍東征
· ca. 1100  Troubadour 在南法興起
· ca. 1163  Notre Dame 大教堂動工
· ca. 1170  Léonin，Magnus Liber
· ca. 1200  Pérotin 四聲部複音
· ca. 1215  Magna Carta 英國憲章
· ca. 1230  Motet 體裁確立
· ca. 1250  Adam de la Halle 活躍
· 1280s     Montpellier Codex 手稿
· 1291       十字軍東征終結`,
"14 世紀",
`· ca. 1320  Vitry，Ars Nova 論文
· ca. 1330  Machaut 生涯鼎盛
· 1337       英法百年戰爭開始
· 1340s     Jacopo 義大利 Trecento
· 1347–51  黑死病席捲歐洲
· ca. 1360  Machaut，Messe de N.D.
· ca. 1370  Ars Subtilior 出現
· ca. 1370  Landini 威尼斯盲目琴師
· 1378–1417 教廷大分裂 Schism
· 1397       Dufay 誕生
· ca. 1415  Dunstaple 英格蘭風格`); }

{ const s=ls(C); hdr(s,C,"考試重點 · Period 2","Exam Focus: The Middle Ages",false);
lp(s,C,"必記作曲家/作品",
`· Machaut · Messe de N.D.（NAWM 25a）
   史上第一首有名複音彌撒
· Machaut · Douce dame jolie（NAWM 26）
   Virelai 典範，世俗代表
· Philippe de Vitry · Ars Nova 論文
   新節奏記譜的理論基礎
· Pérotin · Viderunt omnes
   四聲部複音的頂峰
· Francesco Landini · 140 首 Ballata`,
"必懂概念",
`· Ars Antiqua vs. Ars Nova 區別
   節奏記譜是關鍵差異
· Isorhythm = color + talea
   兩個獨立循環如何運作
· Formes fixes 三種格式比較
   Ballade / Rondeau / Virelai
· Mensural notation 的革命意義
   → 精確記錄節奏長短
· 多語 Motet 的社會功能
   教育、宮廷娛樂、宗教
· 中世紀複音 vs. 文藝復興的差異`); }

{ const s=ds(C); hdr(s,C,"跨期比較","Connecting to Later Periods");
dp(s,C,"中世紀影響 Legacy",
`· Cantus firmus 技法
   → 文藝復興彌撒主要手法
· Isorhythm → 20世紀梅湘
   Turangalîla 重新運用循環
· Formes fixes 三種定型歌
   → 近代流行歌曲格式意識
· Motet → 文藝復興 motet
   語言純化，拉丁單語統一
· Tenor 聲部概念 → 複音低音
   → 後來 Basso continuo 前身
· 大學機構化 → 音樂理論學科
   → 系統性音樂教育的起點`,
"與下一時期連結 Bridge",
`· 1400年前後 Dunstaple 英格蘭風格
   三度六度 → Fauxbourdon
· 勃艮第宮廷 → 法蘭德斯樂派
   Dufay → Ockeghem → Josquin
· 複音彌撒技術傳承連鎖
   Machaut → Ockeghem → Josquin
· 人文主義興起 → 世俗化
   → 文藝復興音樂美學的轉變
· 義大利 Trecento → Quattrocento
   → 文藝復興牧歌的前驅`); }
}

// ══════════════════════════════════════════════════════════════════════════════
// PERIOD 3 · 文藝復興 (11 slides)
// ══════════════════════════════════════════════════════════════════════════════
{ const C = P[3];
cover(C, 3, "文藝復興",
  "The Renaissance",
  "ca. 1420 – 1600",
  "涵蓋 Ch07–12  ·  Chapters 7–12");

{ const s=ds(C); hdr(s,C,"時代背景","Historical Context");
dp(s,C,"政治社會背景 Political Context",
`· 人文主義 Humanism 興起
· Gutenberg 活字印刷術（1450s）
   → 樂譜大量廉價印行
· 義大利文藝復興城邦文化
· 美第奇家族（佛羅倫斯）贊助
· 宗教改革（1517 Luther）
· 天主教反改革（Trent 大公會）
· 探索時代：哥倫布（1492）
· 英格蘭都鐸王朝宮廷音樂
· 法國瓦盧瓦王朝宮廷音樂
· 教廷、宮廷、市民音樂並立
· 文藝復興宮廷：競相贊助音樂`,
"音樂文化概況 Music Culture",
`· 樂譜印刷：Petrucci（1501 Venice）
· 業餘演奏者市場大幅擴張
· 法蘭德斯樂派主宰全歐洲
· 義大利牧歌成宮廷娛樂精品
· 宗教改革：路德德語聖詠
· 日內瓦詩篇集（新教）
· 英國公禱書（Anglican）
· 器樂音樂逐漸與聲樂分離
· 仿古精神：文字意義優先
· Lute 魯特琴盛行於宮廷
· 音樂出版業商業化`); }

{ const s=ds(C); hdr(s,C,"核心作曲家","Key Composers");
dp(s,C,"法蘭德斯樂派 Franco-Flemish",
`· Guillaume Dufay（1397–1474）
   Fauxbourdon，勃艮第香頌
· Johannes Ockeghem（1410–97）
   Missa prolationum，無終卡農
· Josquin des Prez（1450–1521）
   「音符的主人」，模仿對位大師
· Heinrich Isaac（1450–1517）
   德語 Lied，神聖羅馬皇廷
· Jacob Obrecht（1457–1505）
   荷蘭，複雜對位技法
· Orlando di Lasso（1532–94）
   多語言全才，2000首以上`,
"義大利與英格蘭 Italy/England",
`· Giovanni P. da Palestrina（1525–94）
   羅馬樂派，「完美複音」典範
· Carlo Gesualdo（1560–1613）
   半音牧歌，情感極端
· Luca Marenzio（1553–99）
   牧歌精巧，Word painting 大師
· Thomas Morley（1557–1602）
   英式牧歌，嬉樂牧歌
· William Byrd（1543–1623）
   英格蘭，天主教/新教俱寫
· John Dowland（1563–1626）
   魯特琴，Flow my tears`); }

{ const s=ds(C); hdr(s,C,"代表體裁與形式","Genres and Forms");
dp(s,C,"聲樂 Vocal",
`· Mass 彌撒曲（三種類型）
   Cantus firmus Mass 素歌彌撒
   Paraphrase Mass 仿擬彌撒
   Parody Mass 引用彌撒
· Motet 拉丁語經文歌
· Madrigal 義大利語牧歌
· Chanson 法語香頌
· Lied 德語歌曲
· Anglican Anthem 英國讚美詩
· Lutheran Chorale 路德聖詠`,
"器樂 Instrumental",
`· Intabulation 聲樂作品鍵盤改編
· Fantasia 幻想曲（即興風）
· Ricercare 利切卡雷（模仿）
· Canzona 坎佐納（活潑器樂）
· Pavane / Galliard 舞曲對
· Passamezzo / Saltarello
· Lute music 魯特琴音樂
· Consort music 合奏音樂
· Variations 主題變奏
· Toccata 觸技曲（早期）`); }

{ const s=ds(C); hdr(s,C,"風格特徵","Style Features");
dp(s,C,"和聲與對位 Harmony & Counterpoint",
`· 模仿對位 Imitation 主要手法
   聲部依次進入同一動機
· 三六度協和優先（Fauxbourdon）
· 調式：從中世紀八調式走向大小調
· 半音使用：牧歌情感表達工具
· Word painting 文字畫技法
   旋律描繪詞義（如「上升」↑）
· Musica ficta 半音調整習慣
· 避免平行五八度（不協和感）
· 連音線 Syncopation 表達不安
· 平穩流動的對位線條`,
"織體與形式 Texture & Form",
`· 四至六聲部為主（文藝復興標準）
· 各聲部平等（Josquin 之後）
· Cantus firmus → Paraphrase
   → Parody（借用技法三代演進）
· 分節 Strophic 重複歌詞
· 貫穿式 Through-composed 不重複
· 彌撒常規：Kyrie-Gloria-Credo
   -Sanctus-Agnus（五段）
· 牧歌：聲部逐漸增加（5–6聲）
· 器樂：帶來獨立形式意識`); }

nawm(C,"NAWM 44","Josquin des Prez","Ave Maria...virgo serena · Motet · ca.1485",
`· 四聲部模仿對位經文歌
· Petrucci 1502 首批印行
· 各詩節對應新的模仿動機
· 模仿進入順序：S→A→T→B
· 每段動機各有獨特旋律輪廓
· 結尾段轉為同節奏 Homophony
· 拉丁祈禱文，向聖母瑪利亞
· 調式：Dorian G，莊嚴色彩
· 句法清晰：換詩節換音樂材料
· 定義性的文藝復興 Motet 典範
· Luther 讚揚 Josquin 為「音符主人」
· Petrucci 印行 = 印刷文化里程碑`,
`· 注意四聲部依次模仿進入方式
· 感受各詩節動機更換的清晰度
· 留意結尾 Homorhythm 效果
· 思考：如何兼顧文字意義與對位
· 聆聽模仿間距（時間差）的長短
· 注意 S vs. A 聲部的音域差異
· 感受 Dorian 調式的莊嚴氣氛
· 找出換詩節的音樂轉折點
· 對比 Machaut 彌撒的中世紀感
· 思考 Luther 讚揚 Josquin 的理由
· 文藝復興「清晰」美學的體現`,
"scQ5YBRpwNg");

nawm(C,"NAWM 62","Palestrina","Pope Marcellus Mass · Gloria · 1562–63",
`· 六聲部 SSATBB，羅馬樂派典範
· 為紀念 Marcellus II 教皇而作
· 符合 Council of Trent 要求
   文字清晰可辨為首要任務
· 主調 Homophonic 與對位交替
   根據歌詞性質決定織體
· 和諧流暢，避免強烈不協和音
· 模仿段：S+A/T+B 兩組呼應
· 後世稱「完美複音風格」典範
· 長達 3 頁的 Gloria 一氣呵成
· 各聲部旋律線條獨立流暢
· 成為後世對位教學的標準範本
· Zarlino《和聲的機制》理論對應`,
`· 注意文字清晰度（Syllabic 段落）
· 感受 Homophonic vs. 對位的交替
· 辨認何時 Homorhythm 出現
· 對比 Josquin Ave Maria 的線條感
· 思考：Trent 要求如何塑造風格
· 聆聽六聲部的豐富和聲色彩
· 注意各聲部彼此的平衡關係
· 感受「流暢」「清澈」的美學
· 找出最複雜的對位段落在哪
· Palestrina 作為後世標準的意義`,
"oeLIgzAe5sI");

{ const s=ls(C); hdr(s,C,"重要術語 · Period 3","Key Terms: The Renaissance",false);
lp(s,C,"術語 Terms A",
`· Imitation 模仿對位
· Cantus firmus Mass 素歌彌撒
· Paraphrase Mass 仿擬彌撒
· Parody Mass 引用彌撒
· Fauxbourdon 假低音（三六度）
· Word painting 文字畫技法
· Musica ficta 半音臨時調整
· Homophony 主調風格
· Motet 文藝復興拉丁經文歌
· Madrigal 義大利語牧歌
· Chanson 法語香頌
· Council of Trent 特倫特大公會`,
"術語 Terms B",
`· Pavane / Galliard 舞曲對
· Ricercare 利切卡雷模仿器樂
· Fantasia 幻想曲即興器樂
· Intabulation 聲樂改編鍵盤
· Consort 器樂合奏組合
· Lute tablature 魯特琴指法譜
· Lutheran Chorale 路德聖詠
· Psalm settings 詩篇配樂
· Mannerism 手法主義（後期）
· Chromaticism 半音主義（牧歌）
· Strophic / Through-composed
· Petrucci 樂譜出版商（1501）`); }

{ const s=ls(C); hdr(s,C,"時間軸 · Period 3","Timeline: The Renaissance",false);
lp(s,C,"15 世紀",
`· ca. 1420  Dufay 勃艮第風格成熟
· ca. 1420  Fauxbourdon 技法確立
· ca. 1450  Gutenberg 活字印刷術
· ca. 1470  Ockeghem Missa prolationum
· 1480s     Josquin 在 Sistine 教堂
· 1485      Josquin Ave Maria 創作
· 1492      哥倫布抵美洲
· 1494      Petrucci 開始印刷樂譜
· 1501      Petrucci 第一批樂譜集
· 1507      Josquin Missa Pange lingua`,
"16 世紀",
`· 1517      Luther 宗教改革開始
· 1520s     Chanson 在巴黎繁盛
· 1540s     義大利 Madrigal 成熟
· 1545–63  Council of Trent
· 1553      Luca Marenzio 誕生
· 1562      Palestrina Pope Marcellus
· 1563      Council of Trent 結束
· 1575      Byrd / Tallis 合集
· 1580s     英式牧歌（Morley/Byrd）
· 1597      Dowland Lachrimae 出版`); }

{ const s=ls(C); hdr(s,C,"考試重點 · Period 3","Exam Focus: The Renaissance",false);
lp(s,C,"必記作曲家/作品",
`· Josquin · Ave Maria（NAWM 44）
   模仿對位的文藝復興典範
· Palestrina · Pope Marcellus（NAWM 62）
   Trent 後複音音樂的完美典範
· Ockeghem · Missa prolationum
   嚴格卡農技法的頂峰
· Dufay · 勃艮第香頌
   Fauxbourdon 早期實踐
· Byrd · 英格蘭天主教 Motet`,
"必懂概念",
`· 三種 Mass 類型的詳細區別
   CF / Paraphrase / Parody
· Imitation polyphony 如何運作
   動機、時間差、聲部進入
· Word painting 具體例子
   「上升」「哭泣」「黑暗」
· 印刷術對音樂傳播的影響
· 宗教改革→新教 vs. 天主教音樂
   路德聖詠 vs. Palestrina 風格
· 文藝復興 vs. 中世紀複音對比`); }

{ const s=ds(C); hdr(s,C,"跨期比較","Connecting to Later Periods");
dp(s,C,"文藝復興影響 Legacy",
`· 模仿對位 → 巴洛克賦格藝術
· Word painting → 巴洛克情感論
   Affektenlehre 音樂表情規範
· 牧歌半音主義 → Gesualdo
   → 蒙台威爾第「第二實踐」
· 器樂獨立 → 巴洛克器樂體裁
   奏鳴曲、協奏曲的前身
· 教堂調式 → 大小調系統
   → 調性音樂的歷史基礎
· Palestrina 風格 → 後世對位
   教科書標準直至今日`,
"與下一時期連結 Bridge",
`· 1600 年前後歌劇誕生
   佛羅倫斯 Camerata 革命
· Basso continuo 通奏低音出現
· 旋律優先取代多聲部對位
· 宣敘調 Recitative 誕生
   → 語言自然節奏優先
· 情感論 Affektenlehre 建立
   → 音樂有責任模仿情感
· 文藝復興「廣延」→ 巴洛克「深度」`); }
}

// ══════════════════════════════════════════════════════════════════════════════
// PERIOD 4 · 早期巴洛克 (11 slides)
// ══════════════════════════════════════════════════════════════════════════════
{ const C = P[4];
cover(C, 4, "早期巴洛克",
  "The Early Baroque",
  "ca. 1600 – 1680",
  "涵蓋 Ch13–15  ·  Chapters 13–15");

{ const s=ds(C); hdr(s,C,"時代背景","Historical Context");
dp(s,C,"政治社會背景 Political Context",
`· 三十年戰爭（1618–48）宗教戰爭
· Westphalia 和約（1648）現代外交
· 科學革命：Galileo、Newton
· 絕對君主制：法國路易十四
· 凡爾賽宮廷成文化藝術中心
· 義大利城邦：美第奇等貴族贊助
· 威尼斯商業共和國贊助音樂
· 商業資本主義興起，中產階層
· 英格蘭內戰、清教徒革命
· 劇院從宮廷走向公共付費觀眾
· 耶穌會傳教藝術展演活動
· 印刷業、新聞業推動文化傳播`,
"音樂文化概況 Music Culture",
`· 歌劇誕生（1600 佛羅倫斯）
· Basso continuo 通奏低音全面普及
· 情感論 Affektenlehre 確立
· 第一實踐（舊法）vs. 第二實踐
   旋律服務文字的新法
· 裝飾音 Ornament 藝術高度發展
· 公共歌劇院開業（威尼斯 1637）
· Monody 單聲歌曲革命性出現
· 神劇 Oratorio 取代舞台展演
· 義大利風格主宰全歐洲
· 樂器改良：小提琴族群確立
· 第一批職業女歌唱家出現`); }

{ const s=ds(C); hdr(s,C,"核心作曲家","Key Composers");
dp(s,C,"義大利 Italy",
`· Giulio Caccini（1551–1618）
   Monody 單聲歌曲先驅，Le nuove musiche
· Jacopo Peri（1561–1633）
   Euridice 現存最早完整歌劇
· Claudio Monteverdi（1567–1643）
   牧歌書七冊＋L'Orfeo 歌劇
· Girolamo Frescobaldi（1583–1643）
   羅馬管風琴，Fiori musicali
· Francesco Cavalli（1602–76）
   威尼斯公共歌劇院主力
· Barbara Strozzi（1619–77）
   威尼斯，200首 Cantata 多產`,
"德義英 Germany/England",
`· Heinrich Schütz（1585–1672）
   德語巴洛克先驅，Symphoniae sacrae
· Johann Schein（1586–1630）
   路德教會，德語 Lied 融合義大利
· Samuel Scheidt（1587–1654）
   鍵盤音樂，Tabulatura nova
· Henry Purcell（1659–95）
   英格蘭，Dido and Aeneas
· Giacomo Carissimi（1605–74）
   Oratorio 神劇奠基，Jephte
· Marco da Gagliano（1582–1643）
   佛羅倫斯，Dafne 改編版`); }

{ const s=ds(C); hdr(s,C,"代表體裁與形式","Genres and Forms");
dp(s,C,"聲樂 Vocal",
`· Opera 歌劇（Favola in musica）
   佛羅倫斯 Camerata 理想
· Monody 單聲歌曲（1聲+低音）
· Oratorio 神劇（無舞台）
· Cantata 清唱劇（早期）
· Recitative 宣敘調
   secco（乾燥）/ accompagnato
· Aria 詠嘆調（早期自由形）
· Lament 悲歌（固定低音）
· Concertato 大小合唱對比`,
"器樂 Instrumental",
`· Sonata 奏鳴曲（早期形式）
· Canzona 坎佐納（源自法語香頌）
· Toccata 觸技曲（即興鍵盤）
· Fantasia / Ricercare（模仿）
· Chaconne 恰空（固定和聲）
· Passacaglia 帕薩卡利亞（固定低音）
· Variation 主題變奏
· Basso continuo 通奏低音
   鍵盤+撥弦樂器持續伴奏
· Dance suite 舞曲組曲`); }

{ const s=ds(C); hdr(s,C,"風格特徵","Style Features");
dp(s,C,"和聲與調性 Harmony & Tonality",
`· Basso continuo 通奏低音系統
· Figured bass 數字低音記法
· 大小調體系逐漸取代調式
· 不協和音可作表情之用
   第二實踐允許和聲驚喜
· Ostinato 固定低音反覆
· Affetti 情感的音樂模仿規範
· 主屬和弦功能感增強
· 半音轉調：表達悲傷、熱情
· Sequence 模進強化和聲動力`,
"織體與聲樂風格 Texture & Vocal Style",
`· 主調 Homophony 取代多聲對位
· Monody：獨奏旋律＋通奏低音
· Stile recitativo 宣敘風格
   語言節奏、音調模仿語調
· Gorgia 喉嚨裝飾（Caccini）
· Ornamental 裝飾音藝術豐富
· Terraced dynamics 階梯式強弱
· 樂器間戲劇性對比
· 樂句段落明確的起止感`); }

nawm(C,"NAWM 74","Caccini","Vedro 'l mio sol · Monody · 1602",
`· 出自《新音樂》Le nuove musiche
· 單旋律＋通奏低音（無多聲部）
· 旋律自由裝飾 gorgia 喉嚨技法
· 情感表達優先於技巧炫示
· 宣敘風格：語言節奏自然流動
· 示範朗誦調 vs. 裝飾花腔平衡
· 歌詞：愛人的眼睛像太陽
· Caccini 主張裝飾=意義非炫技
· 通奏低音以大鍵琴撥弦實現
· 早期巴洛克聲樂美學里程碑
· 打破文藝復興多聲部傳統
· 個人情感直接表達的宣言`,
`· 注意 Monody 旋律裝飾方式
· 感受宣敘風格的言語節奏感
· 思考：如何以音樂「操控情感」
· 對比文藝復興多聲部牧歌
· 聆聽 gorgia 裝飾的自由度
· 辨別朗誦調段 vs. 花腔段
· 感受通奏低音的支撐作用
· 思考：為何只剩一個旋律聲部
· 對比 Palestrina 的對位線條
· 注意拍號的自由彈性`,
"s-DaH6zpLjs");

nawm(C,"NAWM 78a","Monteverdi","L'Orfeo · 'Possente spirto' · 1607",
`· 歌劇 L'Orfeo 第三幕中心段
· Orfeo 以音樂說服冥界守門人
· Obbligato 必要性器樂對話人聲
· 第一段：小提琴二重奏伴奏
· 第二段：短號二重奏輪替
· 第三段：豎琴精彩對話
· 三段各展示不同 Obbligato 寫法
· 三種聲樂風格在一曲中示範
· Orfeo 的歌唱力量感積累升騰
· 「音樂有說服力量」的神話化
· 歌劇把戲劇與音樂融合的典範
· 蒙台威爾第 Seconda prattica 應用`,
`· 注意器樂與人聲的精心輪替對話
· 感受 Orfeo 說服能量的逐段積累
· 分辨三段不同器樂伴奏的音色
· 對比宣敘 vs. 裝飾詠嘆的段落
· 思考：歌劇如何改變聲樂期待
· 聆聽 Harp 段的特別音響效果
· 感受調性轉換帶來的情感變化
· 與 NAWM 74 Caccini 風格比較
· 注意旋律裝飾的精緻程度
· 思考「音樂說服神明」的隱喻`,
"5CZnZwb3u_g");

{ const s=ls(C); hdr(s,C,"重要術語 · Period 4","Key Terms: Early Baroque",false);
lp(s,C,"術語 Terms A",
`· Basso continuo 通奏低音
· Figured bass 數字低音記法
· Monody 單聲歌曲體裁
· Recitative 宣敘調
· Aria 詠嘆調（早期自由形）
· Opera 歌劇（Favola in musica）
· Stile rappresentativo 戲劇風格
· Affektenlehre 情感論規範
· Prima prattica 第一實踐
· Seconda prattica 第二實踐
· Gorgia 喉嚨裝飾技法
· Obbligato 義務性器樂聲部`,
"術語 Terms B",
`· Ostinato 固定音型反覆
· Chaconne 恰空（和弦固定）
· Passacaglia 帕薩卡利亞
· Lament 悲歌（下行固定低音）
· Oratorio 神劇（無舞台）
· Cantata 清唱劇（早期形式）
· Concertato 大小合唱對比
· Favola in musica 音樂寓言
· Stile antico 古式對位
· Stile moderno 現代風格
· Trillo 震音裝飾（Caccini）
· Camerata 知識分子沙龍聚會`); }

{ const s=ls(C); hdr(s,C,"時間軸 · Period 4","Timeline: Early Baroque",false);
lp(s,C,"歌劇誕生 Opera Origins",
`· ca. 1575  佛羅倫斯 Camerata 開始聚會
· 1597      Peri · Dafne（散佚，最早）
· 1600      Peri · Euridice（現存最早完整）
· 1602      Caccini · Le nuove musiche 出版
· 1607      Monteverdi · L'Orfeo 首演
· 1610      Monteverdi · Vespers of 1610
· 1613      Monteverdi 任威尼斯聖馬可樂長
· 1619      Schütz 回德國傳播義大利風格
· 1630      三十年戰爭中期，德語音樂受創
· 1637      威尼斯首座公共歌劇院開業`,
"巴洛克擴展 Baroque Expansion",
`· ca. 1640  Cavalli 成威尼斯歌劇主力
· 1649      Carissimi Oratorio Jephte
· 1651      Strozzi Cantata Op.1
· 1659      Purcell 誕生，英格蘭代表
· 1660      英格蘭王政復辟，宮廷復甦
· 1668      Buxtehude 任呂北克管風琴師
· 1678      漢堡公共歌劇院開業
· 1681      Corelli 三重奏鳴曲 Op.1
· 1689      Purcell · Dido and Aeneas
· ca. 1695  那不勒斯歌劇傳統成形`); }

{ const s=ls(C); hdr(s,C,"考試重點 · Period 4","Exam Focus: Early Baroque",false);
lp(s,C,"必記作曲家/作品",
`· Caccini · Vedro 'l mio sol（NAWM 74）
   Monody 與 gorgia 裝飾典範
· Monteverdi · L'Orfeo（NAWM 78a）
   Obbligato 器樂與歌劇典範
· Peri · Euridice
   現存最早完整歌劇（1600）
· Monteverdi · 牧歌書第 4–5 冊
   Prima/Seconda prattica 之爭
· Carissimi · Jephte
   Oratorio 體裁確立典範`,
"必懂概念",
`· Basso continuo 的功能與實現
   鍵盤+撥弦，數字低音讀法
· 第一 vs. 第二實踐的區別
   Monteverdi 自辯的理論
· 宣敘調 secco vs. accompagnato
· 歌劇誕生：Camerata 古希臘理想
· Affektenlehre 情感論的應用
   一首作品對應一種情感
· 公共歌劇院（1637）的社會意義`); }

{ const s=ds(C); hdr(s,C,"跨期比較","Connecting to Later Periods");
dp(s,C,"早期巴洛克影響 Legacy",
`· Recitative + Aria 架構
   → 晚期巴洛克歌劇標準格式
· Basso continuo 延續至
   整個巴洛克（1600–1750）
· Ostinato / Passacaglia
   → 巴哈 Chaconne for violin
· Monteverdi 手法影響
   Schütz → 德語音樂傳統
· Monody 個人聲樂表達
   → 後來獨唱詠嘆調發展
· 公共歌劇院 → 中產音樂消費
   → 現代音樂市場前身`,
"與下一時期連結 Bridge",
`· Da capo aria（ABA）形式成形
· 奏鳴曲 da chiesa / da camera 分立
· 獨奏協奏曲 Solo concerto 發展
· 各國形成對比風格
   法式 Opéra-ballet
   英式 Masque
   德式 Kirchenkantate
· 那不勒斯 Opera seria 興起
   → 歌手明星文化
· Ritornello form 利都奈羅確立`); }
}

// ══════════════════════════════════════════════════════════════════════════════
// PERIOD 5 · 晚期巴洛克 (11 slides)
// ══════════════════════════════════════════════════════════════════════════════
{ const C = P[5];
cover(C, 5, "晚期巴洛克",
  "The Late Baroque",
  "ca. 1680 – 1750",
  "涵蓋 Ch16–19  ·  Chapters 16–19");

{ const s=ds(C); hdr(s,C,"時代背景","Historical Context");
dp(s,C,"政治社會背景 Political Context",
`· 路易十四（1643–1715）太陽王
   凡爾賽宮廷文化藝術頂峰
· 英國光榮革命（1688）議會制
· 神聖羅馬帝國諸侯宮廷贊助
· 萊比錫市政府：教堂音樂制度化
· 漢薩城市商業資產階級崛起
· 報紙、咖啡館：市民社會形成
· 科學革命：Newton《原理》1687
· 啟蒙思想萌芽：理性主義
· 普魯士、奧地利爭霸
· 北美殖民地音樂生活（清教徒）
· 俄羅斯彼得大帝西化改革
· 全歐各國宮廷模仿法式宮廷`,
"音樂文化概況 Music Culture",
`· 各國民族風格對比鮮明
   法義英德分歧明顯
· 器樂地位上升至與聲樂平等
· 公開音樂會開始（倫敦 1672）
· 樂譜出版商業化，國際流通
· 大型管弦樂團架構確立
· 對位技術達到巴哈頂峰
· Collegium Musicum 大學音樂社
· 教堂清唱劇每週例行演出
· 小提琴製作黃金時代：Stradivari
· 哈普西寇德 / 管風琴鍵盤鼎盛
· 義大利風格主宰但各國有個性`); }

{ const s=ds(C); hdr(s,C,"核心作曲家","Key Composers");
dp(s,C,"義大利 Italy",
`· Arcangelo Corelli（1653–1713）
   奏鳴曲、協奏曲形式典範
· Antonio Vivaldi（1678–1741）
   500+ 協奏曲，Ritornello 典範
· Alessandro Scarlatti（1660–1725）
   那不勒斯歌劇傳統奠基
· Domenico Scarlatti（1685–1757）
   555 首鍵盤奏鳴曲，西班牙宮廷
· Antonio Caldara（1671–1736）
   維也納，神劇與歌劇
· Giovanni Bononcini（1670–1747）
   倫敦，Händel 的主要競爭者`,
"法英德 France/England/Germany",
`· Jean-Baptiste Lully（1632–87）
   法式歌劇 Tragédie en musique
· François Couperin（1668–1733）
   法式鍵盤組曲 Ordres 27本
· Henry Purcell（1659–95）
   Dido and Aeneas，英格蘭巔峰
· Dietrich Buxtehude（1637–1707）
   呂北克管風琴，影響年輕 Bach
· Georg Philipp Telemann（1681–1767）
   漢堡，產量巨大多樣
· G.F. Händel（1685–1759）神劇王
· J.S. Bach（1685–1750）對位頂峰`); }

{ const s=ds(C); hdr(s,C,"代表體裁與形式","Genres and Forms");
dp(s,C,"器樂 Instrumental",
`· Trio sonata 三重奏鳴曲（2+BC）
· Solo sonata 獨奏奏鳴曲（+BC）
   da chiesa 教會式（慢快慢快）
   da camera 室內式（舞曲組曲）
· Concerto grosso 大協奏曲
   Concertino vs. Ripieno
· Solo concerto 獨奏協奏曲
   3 樂章（快慢快），Ritornello
· French overture 法式序曲
   慢（附點）+ 快（賦格）
· Suite / Partita 舞曲組曲
· Prelude & Fugue 前奏曲＋賦格`,
"聲樂 Vocal",
`· Opera seria 正歌劇
   Da capo aria ABA 為核心
   Castratos 閹伶歌手主角
· Opera buffa 喜歌劇（萌芽）
· Oratorio 神劇（Handel，英文）
   無布景服裝，合唱核心
· Kantate 清唱劇（Bach，週日）
   獨唱+合唱+管弦+Chorale
· Passion 受難曲（Bach）
   St. Matthew / St. John
· Mass 彌撒曲：Bach h-moll Mass`); }

{ const s=ds(C); hdr(s,C,"風格特徵","Style Features");
dp(s,C,"和聲與調性 Harmony & Tonality",
`· 調性體系完全確立
   大小調取代所有教會調式
· Ritornello form 利都奈羅形式
   全奏返回段 + 獨奏插段交替
· Sequence 模進技法廣泛使用
   下行五度序列最常見
· 半音和聲精緻精確（巴哈）
· 和聲節奏：主→屬→主循環
· Well-tempered 近平均律調音
   → 全部 24 調可用
· 屬調轉調作為中段對比`,
"對位與織體 Counterpoint & Texture",
`· Fugue 賦格結構完整：
   Subject 主題→Answer 應答
   Countersubject 反主題
   Stretto 緊接→Stretta 高潮
· Invertible counterpoint 可逆對位
· Stile antico 古式（仿文藝復興）
· Basso continuo 仍廣泛使用
· Terraced dynamics 階梯式強弱
   無 crescendo（到古典才有）
· 大小調情感明確對比
· 複雜對位 vs. 主調齊奏交替`); }

nawm(C,"NAWM 103","J.S. Bach","Durch Adams Fall · Chorale Prelude BWV 637",
`· 出自 Orgelbüchlein 管風琴小冊
· 上聲部：聖詠旋律完整呈現
· 低音：大跳七度象徵亞當墮落
· 中音聲部：半音蛇的誘惑
· 次中音：下滑線象徵罪惡哀傷
· 文字畫在器樂中的完美應用
· 歌詞關於創世記亞當墮落故事
· 無歌詞但聽眾知道歌詞意義
· Bach Orgelbüchlein 46 首之一
· 短小精煉（約 2 分鐘）
· 管風琴四聲部各有象徵任務
· 展示巴洛克「器樂文字畫」技法`,
`· 注意各聲部各自的象徵意義
· 感受低音大跳的「墮落」衝擊
· 聆聽半音聲部的「蛇」誘惑感
· 辨認上聲部聖詠旋律的清晰
· 思考：無歌詞如何做「文字畫」
· 對比 Palestrina 純粹平滑對位
· 感受四聲部各自獨立的張力
· 思考：管風琴色彩如何對應象徵
· 注意整體和聲的緊張感
· 這種寫法叫做 Chorale prelude`,
"Z8Dpe0gjesg");

nawm(C,"NAWM 106","J.S. Bach","Brandenburg Concerto No.5 · BWV 1050 · 1721",
`· 義大利協奏曲形式（快慢快）
· 獨奏組：長笛、小提琴、大鍵琴
· 大鍵琴獲得首次協奏曲主角
· 第一樂章末段大鍵琴超長華彩
   約 65 小節，前所未有規模
· Ritornello form 全奏反覆段
· 六首 Brandenburg 各不同組合
· 贊助者：布蘭登堡侯爵（1721）
· 第二樂章：三獨奏室內樂親密
· 第三樂章：賦格式，振奮明快
· 開創鍵盤樂器獨奏協奏曲先河
· 典範性地示範了 Ritornello 結構`,
`· 注意 ritornello 每次返回結構
· 感受大鍵琴華彩段的衝擊力
· 辨別獨奏插段 vs. 全奏 ritornello
· 思考：大鍵琴如何從伴奏變主角
· 聆聽三獨奏在第二樂章的對話
· 注意長笛與小提琴的角色分配
· 感受第三樂章的輕快賦格性格
· 思考：協奏曲如何平衡獨奏/合奏
· 首開鍵盤協奏曲先河的歷史意義
· 對比 Vivaldi 協奏曲的簡潔形式`,
"LHjbRMIIhuM");

{ const s=ls(C); hdr(s,C,"重要術語 · Period 5","Key Terms: Late Baroque",false);
lp(s,C,"術語 Terms A",
`· Fugue 賦格
· Subject / Answer 主題/應答
· Countersubject 反主題
· Stretto 緊接（密切模仿）
· Ritornello form 利都奈羅形式
· Concerto grosso 大協奏曲
· Solo concerto 獨奏協奏曲
· Da capo aria ABA 返始詠嘆調
· Opera seria 正歌劇
· Castrato 閹伶歌手
· Chorister 清唱劇合唱隊
· Well-tempered 近平均律`,
"術語 Terms B",
`· French overture 法式序曲
· Trio sonata 三重奏鳴曲
· Da chiesa / da camera 兩種
· Chorale prelude 聖詠前奏曲
· Orgelbüchlein 管風琴小冊
· Invertible counterpoint 可逆對位
· Terraced dynamics 階梯式強弱
· Oratorio 神劇（英文版/Handel）
· Passion 受難曲（Bach）
· Kantate 巴哈清唱劇
· Ripieno 協奏大合奏聲部
· Concertino 協奏獨奏小組`); }

{ const s=ls(C); hdr(s,C,"時間軸 · Period 5","Timeline: Late Baroque",false);
lp(s,C,"1680–1720",
`· 1681  Corelli 三重奏鳴曲 Op.1
· 1687  Lully 去世，法式宮廷頂峰
· 1689  Purcell · Dido and Aeneas
· ca.1695 那不勒斯歌劇傳統確立
· 1700  Buxtehude 影響年輕 Bach
· 1709  Scarlatti 羅馬演奏會
· 1711  Händel 倫敦 Rinaldo 首演
· 1714  Stradivarius 製琴黃金期
· 1717  Bach 任 Cöthen 宮廷樂長
· 1718  Vivaldi 《四季》協奏曲`,
"1720–1750",
`· 1721  Bach · Brandenburg Concerti
· 1722  Bach · Well-Tempered Clavier I
· 1724  Bach · St. John Passion
· 1727  Bach · St. Matthew Passion
· 1733  Bach · h-moll Mass 開始
· 1735  Handel · Ariodante 歌劇
· 1741  Händel · Messiah（都柏林）
· 1742  Messiah 倫敦首演
· 1747  Bach 訪 Frederick 大帝
· 1750  J.S. Bach 去世，巴洛克終結`); }

{ const s=ls(C); hdr(s,C,"考試重點 · Period 5","Exam Focus: Late Baroque",false);
lp(s,C,"必記作曲家/作品",
`· Bach · Chorale Prelude（NAWM 103）
   BWV 637，器樂文字畫典範
· Bach · Brandenburg No.5（NAWM 106）
   大鍵琴協奏曲先河
· Händel · Messiah（NAWM 108）
   神劇合唱傳統最高峰
· Vivaldi · 《四季》協奏曲
   Ritornello form 典範
· Corelli · Trio Sonata Op.5
   奏鳴曲體裁範本`,
"必懂概念",
`· Fugue 賦格的四個組成要素
   Subject/Answer/CS/Stretto
· Ritornello form 運作方式
   全奏返回段的功能
· Da capo aria ABA 詳細結構
   A-B-A' 裝飾返回段
· 法式 vs. 義大利式風格對比
   序曲、舞曲 vs. 詠嘆調、協奏
· Bach 對位 vs. Händel 宏偉對比
· 晚期巴洛克 → 早期古典的轉變`); }

{ const s=ds(C); hdr(s,C,"跨期比較","Connecting to Later Periods");
dp(s,C,"晚期巴洛克影響 Legacy",
`· Fugue 賦格技術
   → 古典浪漫持續引用借鑑
· Ritornello form
   → 古典協奏曲雙呈示部前身
· Händel 神劇合唱傳統
   → 海頓《創世記》《四季》
   → 孟德爾頌 Elijah
· Bach 和聲語言
   → 所有後世和聲教科書基礎
· Cantata 清唱劇形式
   → 古典時期 Symphony cantata
· 大型管弦樂隊 → 古典管弦樂`,
"與下一時期連結 Bridge",
`· Galant 輕巧風格 vs. 巴洛克複雜
· C.P.E. Bach：情感風格
   Empfindsamkeit
· Sonata form 奏鳴曲式萌芽
   → 古典主要結構原則
· 公開音樂會文化興起
   → 中產階級聽眾市場
· Opera buffa 喜歌劇成熟
   → Gluck 歌劇改革
· 鋼琴（Fortepiano）漸取代大鍵琴`); }
}

// ══════════════════════════════════════════════════════════════════════════════
// PERIOD 6 · 古典時期 (11 slides)
// ══════════════════════════════════════════════════════════════════════════════
{ const C = P[6];
cover(C, 6, "古典時期",
  "The Classical Period",
  "ca. 1750 – 1810",
  "涵蓋 Ch20–24  ·  Chapters 20–24");

{ const s=ds(C); hdr(s,C,"時代背景","Historical Context");
dp(s,C,"政治社會背景 Political Context",
`· 啟蒙時代 The Enlightenment
   理性、自由、平等為核心
· 狄德羅《百科全書》（1751）
· 美國獨立宣言（1776）
· 法國大革命（1789）
· 拿破崙帝國（1799–1815）
· 市民階層崛起，宮廷贊助式微
· 維也納成為西方音樂首都
· 哈布斯堡宮廷與市民文化並存
· 英國工業革命開始（1760s）
· 美國傳教音樂生活展開
· 出版業、評論業蓬勃發展`,
"音樂文化概況 Music Culture",
`· 公開音樂會制度全面確立
· 樂譜出版市場化，鋼琴譜需求大
· 業餘演奏者市場爆增
· 鋼琴（Fortepiano）取代大鍵琴
· 維也納弦樂四重奏聚會文化
· 古典風格：清晰、均衡、優雅
· 樂評文化：Forkel, Cramer 等
· Esterházy 宮廷（海頓 30年）
· Haydn / Mozart 維也納互訪
· 音樂沙龍：市民私人聚會文化
· 四樂章交響曲形式確立`); }

{ const s=ds(C); hdr(s,C,"核心作曲家","Key Composers");
dp(s,C,"古典三大師 The Three Masters",
`· Joseph Haydn（1732–1809）
   104 首交響曲，68 首弦四重奏
   「交響曲之父」「弦四之父」
· Wolfgang Amadeus Mozart（1756–91）
   41 首交響，20+ 鋼琴協奏曲
   Don Giovanni, Le nozze di Figaro
· Ludwig van Beethoven（1770–1827）
   9 首交響，32 首鋼琴奏鳴曲
   跨越古典→浪漫的橋樑
· 三人共同在維也納活動或影響
· 三人各自建立不同個人風格`,
"其他重要人物 Other Key Figures",
`· C.P.E. Bach（1714–88）
   情感風格 Empfindsamkeit
· Christoph W. Gluck（1714–87）
   歌劇改革，回歸戲劇自然
· Johann Stamitz（1717–57）
   曼海姆樂派，管弦法革新
· Muzio Clementi（1752–1832）
   鋼琴奏鳴曲 100+ 首先驅
· Leopold Mozart（1719–87）
   小提琴教程，父親角色
· Antonio Salieri（1750–1825）
   維也納宮廷樂長，Mozart 同代`); }

{ const s=ds(C); hdr(s,C,"代表體裁與形式","Genres and Forms");
dp(s,C,"器樂 Instrumental",
`· Symphony 交響曲（四樂章）
   快（奏鳴曲式）/慢/小步曲/快
· String quartet 弦樂四重奏
   2Vn+Va+Vc，同等聲部
· Piano sonata 鋼琴奏鳴曲
   Haydn/Mozart/Beethoven 各異
· Piano concerto 鋼琴協奏曲
   雙呈示部 Double exposition
· Piano trio / Wind quintet
· Divertimento / Serenade
   輕鬆場合，多樂章娛樂曲
· Variation set 變奏曲集`,
"聲樂 Vocal",
`· Opera seria 正歌劇（Gluck 改革）
   自然 vs. 修辭：戲劇優先
· Opera buffa 喜歌劇（Mozart 頂峰）
   Ensemble finale 重唱終場
· Singspiel 德語歌唱劇
   對白+歌曲，Zauberflöte
· Oratorio 神劇（Haydn Creation）
   London / Vienna 兩版本
· Mass / Requiem（Mozart 未完成）
· Concert aria 音樂會詠嘆調
· Art song Lied（萌芽期）`); }

{ const s=ds(C); hdr(s,C,"風格特徵","Style Features");
dp(s,C,"奏鳴曲式 Sonata Form",
`· 呈示部 Exposition
   主調第一主題 → 過渡 → 屬調第二主題
· 發展部 Development
   主題動機碎片化，遠調遊走
   不穩定，增加緊張感
· 再現部 Recapitulation
   兩主題回歸主調
   第二主題調性統一
· 尾聲 Coda（Beethoven 大幅擴大）
· 「主題動機發展」為核心技術`,
"古典風格 Classical Style",
`· 清晰句法：4+4 方形對稱樂句
· Alberti bass 阿爾貝提低音伴奏
· Crescendo / Decrescendo 新技法
· 主題對比：第一（強）vs. 第二（弱）
· 半終止 Half cadence 建立期待
· 鋼琴成為主要獨奏和室內樂器
· 標準管弦：2ob+2cor+2fg+弦樂
   後加 2fl+2cl+2tp+tim
· 清晰終止，和聲功能明確`); }

nawm(C,"NAWM 121","Haydn","String Quartet Op.33/2 · 'Joke' · 1781",
`· 俄皇四重奏 Op.33 第二首
· 海頓自稱「全新手法」創作
· 四聲部民主對話，無主次
· 第三樂章稱 Scherzo 非 Minuet
· 終樂章（Rondo）：幽默「笑話」
   在聽眾預期終止處「假裝結束」
· 再次等待又是假終止，循環3次
· 最終才真正結束，全場哄笑
· Op.33 標誌古典弦四形式成熟
· 主題動機碎片化發展技術展示
· 各樂章性格對比鮮明
· 室內樂的「對話」美學典範`,
`· 注意四聲部如何平等對話
· 感受主題動機碎片的發展技術
· 等待終樂章「笑話」結尾三次
· 思考：音樂如何製造「幽默」
· 辨別哪次是假終止哪次是真的
· 感受 Scherzo vs. Minuet 節奏差
· 注意第一樂章的奏鳴曲式結構
· 聆聽各聲部的角色輪換
· 思考弦四如何模仿「對話」
· 對比文藝復興合奏音樂的差異`,
"iZo9FoajL4A");

nawm(C,"NAWM 132","Beethoven","Symphony No. 5 in c minor · Op.67 · 1808",
`· 四音命運動機：♩♩♩♩（三短一長）
· 貝多芬稱「命運在敲門」
· 全曲四個樂章由此動機生成
· 第一樂章：緊張 c 小調奏鳴曲式
· 第二樂章：優美 A♭大調變奏
· 第三樂章：陰暗 Scherzo
   低音管主題引用第一樂章動機
· 第四樂章：勝利 C 大調
   「從黑暗到光明」敘事弧線
· 循環形式貫穿四樂章
· 首演 1808 年，與第六交響同場
· 管弦編制擴大：加 piccolo/trombone`,
`· 辨認四音動機在四個樂章出現
· 感受 c 小調→C 大調的情感轉化
· 注意發展部的緊張積累過程
· 思考：貝多芬如何擴大古典形式
· 聆聽第三→第四樂章的直接連接
· 感受第四樂章的突然光明感
· 辨認動機如何分裂成更小碎片
· 思考：交響曲如何「講述故事」
· 注意銅管加入的戲劇效果
· 對比海頓「笑話」四重奏的不同`,
"I7AQeN-x_Xs");

{ const s=ls(C); hdr(s,C,"重要術語 · Period 6","Key Terms: The Classical Period",false);
lp(s,C,"術語 Terms A",
`· Sonata form 奏鳴曲式
· Exposition 呈示部
· Development 發展部
· Recapitulation 再現部
· First theme / Second theme
· Transition 過渡段
· Coda 尾聲
· Alberti bass 阿爾貝提低音
· Double exposition 雙呈示部
· Cyclic form 循環形式
· Empfindsamkeit 情感風格
· Sturm und Drang 狂飆運動`,
"術語 Terms B",
`· Singspiel 德語歌唱劇
· Opera buffa 喜歌劇
· Opera seria 正歌劇（Gluck改革）
· Ensemble finale 重唱終場
· String quartet 弦樂四重奏
· Motivic development 動機發展
· Half cadence 半終止
· Fortepiano 早期鋼琴
· Mannheim crescendo 曼海姆漸強
· Scherzo 諧謔曲（取代小步舞）
· Rondo form 迴旋曲式
· Theme and variations 主題變奏`); }

{ const s=ls(C); hdr(s,C,"時間軸 · Period 6","Timeline: The Classical Period",false);
lp(s,C,"1750–1780",
`· 1750  Bach 去世，古典時期開始
· 1751  狄德羅《百科全書》
· 1752  Guerre des bouffons 喜歌劇論戰
· 1762  Gluck · Orfeo ed Euridice
· 1769  Haydn 任 Esterházy 副樂長
· 1772  Haydn · Symphony No.45「告別」
· 1773  Haydn Op.33 前期
· 1776  美國獨立宣言
· 1778  Mozart · Piano Sonata K.331
· 1781  Haydn · Quartet Op.33`,
"1780–1810",
`· 1784  Mozart · Piano Concerto K.453
· 1786  Mozart · Don Giovanni
· 1787  Mozart · String Quintet K.516
· 1789  法國大革命爆發
· 1791  Mozart 去世，Requiem 未完
· 1795  Beethoven 維也納出道
· 1798  Beethoven · Pathétique Op.13
· 1805  Beethoven · Eroica 第三交響
· 1808  Beethoven · Sym.5 & 6 首演
· 1809  Haydn 去世`); }

{ const s=ls(C); hdr(s,C,"考試重點 · Period 6","Exam Focus: The Classical Period",false);
lp(s,C,"必記作曲家/作品",
`· Haydn · Quartet Op.33/2（NAWM 121）
   「笑話」弦四，弦四形式典範
· Mozart · Don Giovanni（NAWM 127）
   Opera buffa 重唱終場典範
· Beethoven · Sym.5（NAWM 132）
   命運動機，循環形式
· Mozart · Piano Concerto K.488
   雙呈示部協奏曲形式
· Gluck · Orfeo ed Euridice（1762）
   歌劇改革宣言作品`,
"必懂概念",
`· 奏鳴曲式三部分詳細說明
   呈示/發展/再現的調性邏輯
· 雙呈示部協奏曲形式
   管弦→獨奏的功能分工
· 四樂章交響曲各樂章功能
· Haydn / Mozart / Beethoven
   個人風格比較三人差異
· 古典風格 vs. 巴洛克風格對比
   清晰 vs. 對位，鋼琴 vs. 大鍵琴`); }

{ const s=ds(C); hdr(s,C,"跨期比較","Connecting to Later Periods");
dp(s,C,"古典影響 Legacy",
`· 奏鳴曲式 → 浪漫時期擴大
   仍為主要結構框架
· Beethoven 晚期風格 = 浪漫先兆
   → 舒伯特、布拉姆斯直接傳承
· 弦樂四重奏 → 浪漫最重要
   室內樂體裁（布拉姆斯等）
· 交響曲傳統 → 馬勒、布魯克納
   規模越來越大直至 20 世紀
· 鋼琴協奏曲 → 浪漫最重要
   獨奏體裁（Chopin/Liszt）
· 歌劇改革 → Rossini/Verdi/Wagner`,
"與下一時期連結 Bridge",
`· 1815 後浪漫主義興起
   個人情感主義取代均衡理想
· 鋼琴技術革命（Erard 1823）
   → Chopin、Liszt 鋼琴詩
· 藝術歌曲 Lied 成熟
   → 舒伯特黃金時期
· 民族主義情緒興起
   → 各國尋找音樂「民族靈魂」
· Beethoven 的「英雄」敘事弧
   → 浪漫主義「標題音樂」前驅`); }
}

// ══════════════════════════════════════════════════════════════════════════════
// PERIOD 7 · 浪漫主義 (12 slides)
// ══════════════════════════════════════════════════════════════════════════════
{ const C = P[7];
cover(C, 7, "浪漫主義",
  "The Romantic Era",
  "ca. 1815 – 1900",
  "涵蓋 Ch25–30  ·  Chapters 25–30");

{ const s=ds(C); hdr(s,C,"時代背景","Historical Context");
dp(s,C,"政治社會背景 Political Context",
`· 工業革命改變社會結構（1760–）
· 維也納會議（1815）後保守主義
· 民族主義浪潮：1848 年革命
· 德意志統一（1871）Bismarck
· 義大利統一運動 Risorgimento
· 法第二帝國，巴黎歌劇院文化
· 中產階級成為主要音樂受眾
· 鐵路使演奏家巡演成為可能
· 大型音樂廳建設（Carnegie 1891）
· 科學主義 vs. 浪漫主義的張力
· 殖民主義：「異國情調」影響音樂`,
"音樂文化概況 Music Culture",
`· 音樂評論與出版業繁盛
· 名演奏家「明星」文化（Liszt）
· 民族樂派：各國尋找音樂靈魂
· 「純音樂」vs.「標題音樂」論爭
· 保守派 Brahms vs. 前衛 Wagner
· Wagner 拜魯特樂劇院（1876）
· 音樂學 Musicology 學科萌芽
· 女性作曲家困境（Clara Schumann）
· 音樂保育社義大利/德法並行
· 樂器改良：鋼琴、長號等現代化`); }

{ const s=ds(C); hdr(s,C,"核心作曲家（早期）","Key Composers: Early Romanticism");
dp(s,C,"德奧 German-Austrian",
`· Franz Schubert（1797–1828）
   600+ Lieder，Die schöne Müllerin
· Carl Maria von Weber（1786–1826）
   Der Freischütz，德語浪漫歌劇
· Felix Mendelssohn（1809–47）
   古典形式浪漫內容，復興 Bach
· Clara Schumann（1819–96）
   鋼琴家作曲家，女性先驅典範
· Robert Schumann（1810–56）
   鋼琴、Lied、樂評，Neue Zeitschrift`,
"法義及鋼琴 French/Italian/Piano",
`· Hector Berlioz（1803–69）
   Symphonie fantastique，管弦革命
· Frédéric Chopin（1810–49）
   鋼琴詩人，波蘭靈魂，Ballade
· Franz Liszt（1811–86）
   鋼琴技巧革命，交響詩創始
· Giacomo Meyerbeer（1791–1864）
   Paris，Grand opéra 大製作
· Gioachino Rossini（1792–1868）
   義大利喜歌劇，Barber of Seville`); }

{ const s=ds(C); hdr(s,C,"核心作曲家（晚期）","Key Composers: Late Romanticism");
dp(s,C,"歌劇與民族樂派 Opera/Nationalism",
`· Giuseppe Verdi（1813–1901）
   義大利歌劇，民族精神 Risorgimento
· Richard Wagner（1813–83）
   德國樂劇，Leitmotif，Bayreuth
· Georges Bizet（1838–75）
   Carmen（法），東方情調
· Modest Mussorgsky（1839–81）
   俄國五人組，Boris Godunov
· Bedřich Smetana（1824–84）
   捷克，Má vlast「我的祖國」
· Antonín Dvořák（1841–1904）
   波希米亞，新世界交響`,
"晚期浪漫 Late Romantic",
`· Johannes Brahms（1833–97）
   古典形式浪漫內容，四首交響
· Anton Bruckner（1824–96）
   宏偉交響，天主教靈性
· Gustav Mahler（1860–1911）
   交響 = 宇宙，九首→過渡20c
· Richard Strauss（1864–1949）
   交響詩 Also sprach / Don Juan
· Hugo Wolf（1860–1903）
   德語 Lied 最後大師
· Jean Sibelius（1865–1957）
   芬蘭民族，Finlandia`); }

{ const s=ds(C); hdr(s,C,"代表體裁與形式","Genres and Forms");
dp(s,C,"器樂 Instrumental",
`· Lied 藝術歌曲（鋼琴+聲樂）
· Song cycle 連篇歌曲集
   Winterreise / Dichterliebe
· Piano miniature 鋼琴小品
   Nocturne・Ballade・Mazurka
   Impromptu・Étude・Waltz
· Program symphony 標題交響曲
   Berlioz idée fixe 固定樂念
· Symphonic poem 交響詩
   Liszt 創立，Strauss 頂峰
· Piano concerto 鋼琴協奏曲`,
"聲樂 Vocal",
`· Grand opéra 大歌劇（Paris）
· Bel canto opera 美聲歌劇
   Bellini / Donizetti
· Verismo opera 真實主義歌劇
· Music drama 樂劇（Wagner）
   Leitmotif 主導動機系統
   Gesamtkunstwerk 總體藝術
· Nationalist opera 民族歌劇
· Choral Oratorio / Mass / Requiem
· Song cycle Liederkreis`); }

{ const s=ds(C); hdr(s,C,"風格特徵","Style Features");
dp(s,C,"和聲與調性 Harmony & Tonality",
`· 半音和聲 Chromaticism 豐富
· Harmonic color 色彩優先於功能
· 遠關係轉調：第三度關係調
· Neapolitan chord 那不勒斯和弦
· Augmented sixth 增六和弦群
· Deceptive cadence 欺騙終止
· 調性模糊（Tristan 極端例子）
· Endless melody 無終旋律
   解決永遠被推遲
· 半音進行代替全音進行`,
"形式與管弦法 Form & Orchestration",
`· 形式自由化：Episodic structure
· Cyclic form 循環形式（動機貫穿）
· 管弦樂編制大幅擴大
   piccolo・contrabassoon・tuba
   tam-tam・harp・celesta
· 個人情感表達至上
· Rubato 自由節拍（演奏個性）
· 民謠引用：Nationalism 素材
· 標題音樂有文字程式
· 樂團分部細化（弦樂 18 把）`); }

nawm(C,"NAWM 136","Schubert","Winterreise · Der Lindenbaum · 1827",
`· Winterreise 24 首連篇歌曲集
· Wilhelm Müller 詩，舒伯特譜曲
· 失戀旅人冬日漫遊，孤獨敘事
· E 大調開始，E 小調轉換
· 鋼琴引子：菩提樹搖曳圖像
· 人聲：表面甜美的民謠旋律
· 歌詞說「溫暖」但調性暗示「死亡」
· 鋼琴後奏比人聲更能表達內心
· 歌曲集第五首，整體意象核心
· Schubert 數月後去世（1828）
· 鋼琴與人聲地位完全平等
· Lied 作為藝術體裁的最高成就`,
`· 注意鋼琴如何「說故事」
· 感受 E大調→E小調的情感落差
· 思考：歌詞說溫暖，音樂暗示死亡
· 聆聽鋼琴後奏如何超越人聲表達
· 注意引子的菩提葉圖像音符
· 感受大調旋律的「假裝溫暖」
· 思考：旅人為何難以離去
· 聆聽鋼琴如何從伴奏變成主體
· 對比一般藝術歌曲的從屬關係
· 想像 1828 年舒伯特的內心世界`,
"3LhO43EEPMg");

nawm(C,"NAWM 140","Chopin","Ballade No. 1 in g minor · Op.23 · 1835",
`· g 小調，單樂章，自由奏鳴曲形式
· 文學啟發：Mickiewicz 波蘭史詩
· 引子（序章）：那不勒斯六和弦
· 主題 1：g 小調，憂鬱波蘭調
· 主題 2：E♭大調抒情寬廣
· 發展：兩主題對話交織
· Coda：presto con fuoco 戲劇爆發
   悲劇性結尾，無浪漫主義幻想
· 鋼琴技巧：八度、快速音群
· 波蘭民族情感的音樂化身
· 開創鋼琴音樂「文學性」先例
· 四首 Ballade 各有文學背景`,
`· 注意主題 1 的波蘭民謠氣質
· 感受奏鳴曲式如何自由化
· 等待 Coda 的戲劇性爆發
· 思考：鋼琴如何成為「管弦樂」
· 聆聽主題 2 的甜美與主題 1 對比
· 注意發展部如何加劇緊張感
· 感受 Coda 的無可挽回悲劇感
· 辨認那不勒斯和弦的特殊效果
· 思考 Chopin 作為民族音樂象徵
· 對比 Schubert Lied 的歌唱性`,
"VmFmAvwO1pE");

nawm(C,"NAWM 154","Wagner","Tristan und Isolde · Act I Prelude · 1865",
`· 「特里斯坦和弦」F-B-D♯-G♯
· 半音進行解決被無限延宕
   → 「無終旋律」感受
· 主導動機 Leitmotif 四個：
   愛 Love・渴望 Desire
   死亡 Death・目光 Glance
· 調性邊緣，無明確功能和聲解決
· 樂劇無分曲：通貫式連續音樂
· 影響：Debussy 印象主義
   Schoenberg 走向無調性
· 序曲獨立音樂會演奏
· 附加「愛之死 Liebestod」結尾`,
`· 注意「特里斯坦和弦」反覆出現
· 感受永無解決的緊張延宕感
· 辨認四個 Leitmotif 的旋律輪廓
· 思考：不解決的和弦如何運作
· 聆聽旋律如何無止盡的往上攀升
· 感受「渴望」動機的半音色彩
· 思考「調性邊緣」的意義
· 對比 Brahms 的調性明確感
· 思考 Wagner 對後世的衝擊
· 辨認 Leitmotif 如何指涉劇情`,
"_LXa4xPULp8");

{ const s=ls(C); hdr(s,C,"重要術語 · Period 7","Key Terms: The Romantic Era",false);
lp(s,C,"術語 Terms A",
`· Lied 藝術歌曲
· Song cycle 連篇歌曲集
· Program music 標題音樂
· Absolute music 絕對音樂
· Symphonic poem 交響詩
· Leitmotif 主導動機
· Gesamtkunstwerk 總體藝術
· Music drama 樂劇（Wagner）
· Endless melody 無終旋律
· Idée fixe 固定樂念（Berlioz）
· Chromaticism 半音主義
· Rubato 自由節拍`,
"術語 Terms B",
`· Character piece 性格小品
· Nocturne / Ballade / Étude
· Nationalism 民族主義音樂
· Cyclic form 循環形式
· Grand opéra 大歌劇（Paris）
· Bel canto 美聲唱法歌劇
· Verismo 真實主義歌劇
· Neapolitan chord 那不勒斯和弦
· Augmented sixth 增六和弦
· Tristan chord 特里斯坦和弦
· Liederkreis 連篇歌曲集
· Bayreuth 拜魯特樂劇院`); }

{ const s=ls(C); hdr(s,C,"時間軸 · Period 7","Timeline: The Romantic Era",false);
lp(s,C,"1815–1860",
`· 1815  維也納會議，浪漫主義開始
· 1816  Schubert · Erlkönig
· 1821  Weber · Der Freischütz
· 1827  Schubert · Winterreise
· 1830  Berlioz · Symphonie fantastique
· 1833  Brahms 誕生
· 1835  Chopin · Ballade No.1
· 1839  Schumann · Kinderszenen
· 1848  歐洲民族主義革命浪潮
· 1851  Liszt · 交響詩形式確立
· 1854  Schumann 精神崩潰`,
"1860–1900",
`· 1859  Verdi · Un ballo in maschera
· 1865  Wagner · Tristan und Isolde
· 1866  Smetana · Prodaná nevěsta
· 1871  德意志帝國成立
· 1876  Wagner Bayreuth Ring 首演
· 1877  Brahms · Symphony No.2
· 1880  Brahms · Piano Concerto No.2
· 1888  R. Strauss · Don Juan
· 1893  Dvořák · New World Symphony
· 1896  Mahler · Symphony No.1`); }

{ const s=ls(C); hdr(s,C,"考試重點 · Period 7","Exam Focus: The Romantic Era",false);
lp(s,C,"必記作曲家/作品",
`· Schubert · Winterreise（NAWM 136）
   Der Lindenbaum，Lied 典範
· Chopin · Ballade No.1（NAWM 140）
   g 小調，鋼琴敘事曲典範
· Wagner · Tristan（NAWM 154）
   Act I Prelude，Leitmotif 典範
· Berlioz · Symphonie fantastique
   標題交響曲，idée fixe
· Brahms · Symphony No.4
   古典形式浪漫主義的橋樑`,
"必懂概念",
`· Leitmotif 的功能：
   如何指涉角色/情感/概念
· 特里斯坦和弦的和聲分析
   為何如此「不解決」
· 標題音樂 vs. 絕對音樂論爭
   Hanslick vs. Wagner 立場
· 浪漫主義民族主義特徵
   民謠引用、民族語言歌劇
· 晚期浪漫和聲如何走向邊緣
   → 直接預示 20c 無調性`); }
}

// ══════════════════════════════════════════════════════════════════════════════
// PERIOD 8 · 二十世紀前半 (11 slides)
// ══════════════════════════════════════════════════════════════════════════════
{ const C = P[8];
cover(C, 8, "二十世紀前半",
  "The Early Twentieth Century",
  "ca. 1890 – 1945",
  "涵蓋 Ch31–35  ·  Chapters 31–35");

{ const s=ds(C); hdr(s,C,"時代背景","Historical Context");
dp(s,C,"政治社會背景 Political Context",
`· 第一次世界大戰（1914–18）
· 俄國革命（1917）布爾什維克
· 威瑪共和國（1919–33）
· 大蕭條（1929–）全球經濟崩潰
· 納粹德國（1933–45）藝術管控
· 第二次世界大戰（1939–45）
· 大規模移民：歐洲作曲家赴美
   Schoenberg/Stravinsky 等
· 好萊塢電影工業崛起（1920s）
· 廣播電台（1920）改變音樂消費
· 留聲機普及，錄音文化開始
· 殖民地獨立運動全球展開`,
"音樂文化概況 Music Culture",
`· 現代主義：與浪漫主義傳統決裂
· 爵士樂（美國，非裔文化根源）
· 流行音樂工業化，Tin Pan Alley
· 百老匯音樂劇崛起（New York）
· 音樂學 Musicology 學科建立
· 新古典主義 vs. 表現主義對立
· 各國民族風格繼續分歧發展
· 音樂學院制度確立，作曲學術化
· 美學論爭：實用音樂 vs. 純藝術
· 電影配樂需求創造新市場`); }

{ const s=ds(C); hdr(s,C,"核心作曲家","Key Composers");
dp(s,C,"歐洲現代主義 European Modernism",
`· Claude Debussy（1862–1918）
   印象主義，Pelléas，調性模糊
· Arnold Schoenberg（1874–1951）
   無調性→十二音列，Pierrot lunaire
· Igor Stravinsky（1882–1971）
   節奏革命，三部芭蕾→新古典
· Béla Bartók（1881–1945）
   匈牙利民謠+現代技法
· Alban Berg（1885–1935）
   表現主義，Wozzeck/Violin Concerto
· Anton Webern（1883–1945）
   點描主義，極致精簡`,
"美洲與俗樂 Americas/Vernacular",
`· Scott Joplin（1868–1917）
   Ragtime 之王，Maple Leaf Rag
· W.C. Handy（1873–1958）
   Blues 傳播者，St. Louis Blues
· Duke Ellington（1899–1974）
   Big Band 爵士，Cotton Club
· Louis Armstrong（1901–71）
   爵士 improvisation 革命
· George Gershwin（1898–1937）
   跨越爵士與古典，Rhapsody in Blue
· Charles Ives（1874–1954）
   美國前衛先驅，polytonality`); }

{ const s=ds(C); hdr(s,C,"代表體裁與形式","Genres and Forms");
dp(s,C,"古典 Classical",
`· 交響曲（Sibelius・Bartók・Prokofiev）
· 管弦組曲：Stravinsky 三部芭蕾
   Firebird・Petrouchka・Rite
· String quartet：Bartók 6首
· Neoclassical forms 回歸古典
   Suite・Concerto grosso 重用
· Song cycle 無調性
   Schoenberg · Pierrot lunaire
· Piano étude（Bartók Mikrokosmos）
· Opera 歌劇：Wozzeck / Pelléas`,
"俗樂 Vernacular",
`· Ragtime：固定低音，切分節奏
   Scott Joplin 代表
· Blues：12-bar 12小節藍調
   I-IV-I-V-IV-I 和聲進行
· Jazz：improvisation 即興核心
   Big Band / Dixieland / Bebop
· Musical comedy 百老匯音樂劇
   Kern / Gershwin / Cole Porter
· Film music 好萊塢電影配樂
· Tin Pan Alley 流行歌曲工廠`); }

{ const s=ds(C); hdr(s,C,"風格特徵","Style Features");
dp(s,C,"現代主義手法 Modernist Techniques",
`· Atonality 無調性（無調性中心）
· Twelve-tone Serialism 十二音列
   12個半音各用一次才可重複
· Tone row 操作：P/I/R/RI 四形
· Polytonality 雙調性（Stravinsky）
· Octatonic scale 八聲音階
   半全音交替，特殊色彩
· Whole-tone scale 整音音階
   六個全音，漂浮感（Debussy）
· Sprechstimme 說唱（Schoenberg）
   介於說話與唱歌之間
· Irregular / Changing meter`,
"新古典與民族 Neoclassicism/Folk",
`· Neoclassicism 回歸巴洛克形式
   但和聲語言現代（Stravinsky）
· Folk material 民謠素材運用
   Bartók 直接採集匈牙利民謠
· Primitivism 原始主義
   《春之祭》節奏暴力
· Pandiatonicism 泛調性
   調性保留但不服從功能
· Polyrhythm 多節奏層疊
· Extended techniques 延伸技法
   弦樂撥奏/打弓/特殊泛音`); }

nawm(C,"NAWM 170","Debussy","Pelléas et Mélisande · Act IV Sc.4 · 1902",
`· 象徵主義歌劇，Maeterlinck 劇本
· 「印象主義歌劇」唯一代表作
· 朗誦調取代詠嘆調，無 aria
· 調性中心游移模糊，不解決
· 整音音階 whole-tone scale
· 平行和弦移動（organum 聯想）
· 沉默的戲劇性使用
· 旋律線完全依從法語語調
· 管弦法：弦樂羽音 sul ponticello
· 愛情場面：感傷，不誇張
· 與 Wagner 樂劇的刻意決裂
· 人聲幾乎是旁白式的存在`,
`· 感受聲樂如何接近說話節奏
· 注意整音音階的漂浮感
· 思考：印象主義 = 暗示 非描繪
· 對比 Wagner Leitmotif 的明確感
· 聆聽管弦法中弦樂弱音效果
· 注意平行和弦移動的色彩
· 感受沉默段落的戲劇張力
· 思考：為何沒有詠嘆調
· 聆聽法語語調對旋律的影響
· 思考 Debussy 與 Wagner 的對話`,
"ojMlJ7zZMiU");

nawm(C,"NAWM 172","Stravinsky","The Rite of Spring · Part I · 1913",
`· 芭蕾《春之祭》1913 巴黎首演暴動
· Nijinsky 舞蹈編排同樣激進
· 不規則拍號每小節快速切換
   2/4、3/4、3/8 不斷變換
· 重音在非強拍，完全不可預測
   故意打亂「正常」節拍感
· 雙調性：E♭+E 同時進行
   Augurs of Spring 主題
· 打擊樂組突出，節奏衝擊感強
· 八聲音階材料，民謠旋律扭曲
· 與浪漫主義的徹底決裂
· 1913 年巴黎：現代主義分水嶺`,
`· 注意拍號如何每小節改變
· 感受「錯位」重音的身體衝擊
· 辨認 Augurs 主題的雙調性
· 思考：為何 1913 年聽眾暴動
· 感受打擊樂組的主導地位
· 聆聽民謠旋律的扭曲變形方式
· 對比 Debussy 的「優雅」印象主義
· 注意不規則重音的「暴力美學」
· 思考：什麼是「原始主義」
· 感受這與浪漫主義的根本斷裂`,
"EkwqPJZe8ms");

{ const s=ls(C); hdr(s,C,"重要術語 · Period 8","Key Terms: Early Twentieth Century",false);
lp(s,C,"術語 Terms A",
`· Atonality 無調性
· Twelve-tone technique 十二音列
· Tone row 音列（12音）
· Prime / Inversion / Retrograde
· Sprechstimme 說唱技法
· Polytonality 雙調性
· Neoclassicism 新古典主義
· Octatonic scale 八聲音階
· Whole-tone scale 整音音階
· Irregular / Changing meter
· Primitivism 原始主義
· Expressionism 表現主義`,
"術語 Terms B",
`· Impressionism 印象主義
· Parallel chords 平行和弦
· Pentatonic scale 五聲音階
· Ragtime 拉格泰姆
· Blues 藍調（12小節）
· Jazz improvisation 即興
· Big Band swing 大樂團搖擺
· Musical comedy 百老匯音樂劇
· Film music 電影配樂
· Polyrhythm 多節奏層疊
· Extended techniques 延伸技法
· Tin Pan Alley 流行歌廠`); }

{ const s=ls(C); hdr(s,C,"時間軸 · Period 8","Timeline: Early Twentieth Century",false);
lp(s,C,"1890–1920",
`· 1894  Debussy · Prélude à l'après-midi
· 1897  Schoenberg · Verklärt Nacht
· 1899  Joplin · Maple Leaf Rag
· 1902  Debussy · Pelléas et Mélisande
· 1905  Schoenberg · Kammersymphonie
· 1908  Schoenberg 第一首無調性
· 1910  Stravinsky · Firebird
· 1912  Schoenberg · Pierrot lunaire
· 1913  Stravinsky · Rite of Spring 暴動
· 1914–18  第一次世界大戰`,
"1920–1945",
`· 1920s  爵士黃金年代 New Orleans
· 1923  Schoenberg 十二音列確立
· 1924  Gershwin · Rhapsody in Blue
· 1925  Berg · Wozzeck 首演
· 1927  Ellington Cotton Club 時代
· 1935  Berg · Violin Concerto
· 1935  Gershwin · Porgy and Bess
· 1937  Bartók · Music for Strings
· 1939–45  第二次世界大戰
· 1945  Bartók / Webern 雙雙去世`); }

{ const s=ls(C); hdr(s,C,"考試重點 · Period 8","Exam Focus: Early Twentieth Century",false);
lp(s,C,"必記作曲家/作品",
`· Debussy · Pelléas（NAWM 170）
   印象主義歌劇典範
· Stravinsky · Rite of Spring（NAWM 172）
   節奏革命，原始主義
· Schoenberg · Pierrot lunaire（NAWM 171）
   Sprechstimme，表現主義
· Bartók · 弦樂四重奏系列
   民謠採集+現代技法
· Joplin · Maple Leaf Rag（NAWM 169a）
   Ragtime 典範`,
"必懂概念",
`· 無調性 vs. 十二音列的區別
   自由無調性 vs. 系統化
· 音列操作四形：P/I/R/RI
· 新古典主義的「回歸」意涵
   形式回歸 + 和聲現代
· 爵士樂的非裔美國文化根源
   Blues / Ragtime / Improvisation
· Rite of Spring 為何革命性
   節奏不規律 + 雙調性 + 暴動`); }

{ const s=ds(C); hdr(s,C,"跨期比較","Connecting to Later Periods");
dp(s,C,"20世紀前半影響 Legacy",
`· 無調性 → 戰後整體序列主義
   Boulez/Stockhausen 全序列化
· 爵士即興 → 融合型前衛音樂
   Free Jazz / ECM 錄音風格
· Stravinsky 新古典 → 戰後
   新古典主義持續（各國）
· Bartók 民謠研究方法
   → 民族音樂學 Ethnomusicology
· 電影配樂：好萊塢借用浪漫
   管弦法（Korngold / Steiner）
· Cage 實驗在 1940s 開始醞釀`,
"與下一時期連結 Bridge",
`· 1945 後序列主義 vs. 極簡主義
   成為下一代核心衝突
· Cage 偶然音樂打破一切邊界
· 磁帶音樂 → 電腦音樂
· Cold War：東西方音樂對立
   Shostakovich vs. Boulez
· 反文化運動 1960s 改變美學
· 流行音樂（搖滾）衝擊古典市場`); }
}

// ══════════════════════════════════════════════════════════════════════════════
// PERIOD 9 · 二十世紀後半至當代 (11 slides)
// ══════════════════════════════════════════════════════════════════════════════
{ const C = P[9];
cover(C, 9, "二十世紀後半至當代",
  "The Late Twentieth Century to the Present",
  "ca. 1945 – 2020s",
  "涵蓋 Ch36–39  ·  Chapters 36–39");

{ const s=ds(C); hdr(s,C,"時代背景","Historical Context");
dp(s,C,"政治社會背景 Political Context",
`· 冷戰（1947–89）東西方對峙
· 馬歇爾計劃重建西歐（1948）
· 去殖民化、民族獨立浪潮
· 1960s 文化革命，反主流運動
· 越戰（1955–75）社會撕裂
· 柏林圍牆倒塌（1989）
· 網際網路興起（1990s）
· 全球化，多元文化主義
· 911 事件（2001）反恐時代
· 串流媒體（Spotify 2006）
· 社群媒體改變文化傳播`,
"音樂文化概況 Music Culture",
`· 序列主義 Serialism 統治歐美學院
· 電子音樂、磁帶音樂室 musique concrète
· 極簡主義反衝學院派高壓
· 後現代：引用、拼貼、多元並存
· 世界音樂 World music 市場崛起
· 古樂復興 HIP 運動席捲演奏界
· 搖滾、流行音樂衝擊古典市場
· 音樂廳觀眾老齡化問題
· 跨界 Crossover 打破類型邊界
· 數位錄音改變製作與消費`); }

{ const s=ds(C); hdr(s,C,"核心作曲家","Key Composers");
dp(s,C,"戰後歐洲 Postwar Europe",
`· Olivier Messiaen（1908–92）
   色彩和聲、鳥鳴、天主教靈性
· Pierre Boulez（1925–2016）
   整體序列，Structures，Le Marteau
· Karlheinz Stockhausen（1928–2007）
   電子音樂先驅，Gesang der Jünglinge
· Luciano Berio（1925–2003）
   拼貼，Sinfonia 多語引用
· Arvo Pärt（1935–）
   Tintinnabuli，新簡約
· Witold Lutosławski（1913–94）
   波蘭，Aleatoric counterpoint`,
"美洲與後現代 Americas/Postmodern",
`· John Cage（1912–92）
   偶然音樂，4'33''，準備鋼琴
· Steve Reich（1936–）
   極簡，相位技法，非洲節奏
· Philip Glass（1937–）
   重複圖案，歌劇，電影配樂
· John Adams（1947–）
   後極簡，Nixon in China
· Joan Tower（1938–）
   美國女性作曲家，管弦先驅
· Thomas Adès（1971–）
   英國，The Tempest 歌劇`); }

{ const s=ds(C); hdr(s,C,"代表體裁與形式","Genres and Forms");
dp(s,C,"學院與前衛 Academic/Avant-garde",
`· Total serialism 整體序列音樂
   音高/節奏/力度/音色全序列
· Musique concrète 具體音樂
   錄音剪接，Pierre Schaeffer
· Electronic / Computer music
   磁帶→合成器→電腦生成
· Aleatory / Chance music 偶然
   Cage 以擲硬幣決定音符
· Spectralism 頻譜音樂
   和聲從泛音列推導
· New complexity 新複雜主義
· Extended techniques 延伸技法`,
"後現代與當代 Postmodern/Contemporary",
`· Minimalism 極簡主義
   重複 / 相位 / 加法過程
· Tintinnabuli style（Pärt）
   主音三和弦 + 旋律聲部
· Postmodern collage 後現代拼貼
   Polystylism 多風格引用
· New opera 當代歌劇
   Nixon in China / Tempest
· Film music 電影配樂（John Williams）
· Musical theatre 百老匯（Sondheim）
· World music fusion 世界融合`); }

{ const s=ds(C); hdr(s,C,"風格特徵","Style Features");
dp(s,C,"序列主義 Serialism",
`· Total serialism：
   音高/節奏/力度/音色全序列化
· Point style 點描主義
   繼承 Webern 的極致簡約
· IRCAM 電腦輔助創作（巴黎）
· Complexity = 美學意識形態
   學院序列派認為複雜=高尚
· Integral serialism 整體序列
   每個音符都由數列決定
· 偶然音樂與序列主義的對立
   控制 vs. 無控制的辯證`,
"極簡與後現代 Minimalism/Postmodern",
`· Repetition 重複：靜態和聲
· Phase shifting 相位技法（Reich）
· Additive process 加法過程
· Tonal center 調性回歸傾向
· Quotation / Allusion 跨風格引用
· Intertextuality 跨文本互涉
· Ambient / Drone 環境音樂
· Spectral harmony 頻譜和聲
   以泛音列為和弦構成依據
· 跨界融合：爵士/民謠/電子/世界`); }

nawm(C,"NAWM 180","Messiaen","Quartet for End of Time · Mvt. VIII · 1941",
`· 德軍戰俘營 Stalag VIIIA 中創作
· 小提琴+大提琴+鋼琴+單簧管
   戰俘營現有樂器
· 第八樂章：極緩慢，小提琴獨奏
   「對耶穌不朽的讚歌」
· 有限移調調式 Mode of ltd. transpositions
· 色彩和聲：無功能性進行
· 宗教冥想：基督再臨的啟示
· 超越時間感的靜態美學
· 首演 1941，400 名戰俘聽眾
· Messiaen 視顏色與音符的對應`,
`· 感受極慢速的「超時間」空間感
· 注意小提琴高音域的空靈色彩
· 思考：如何在戰俘營創作偉大作品
· 聆聽和聲的非功能性色彩美
· 感受宗教冥想的神聖靜謐感
· 對比序列主義的「複雜計算」
· 注意旋律如何重複卻越來越高
· 思考信仰如何直接影響音樂語言
· 聆聽大提琴沉低的持音對比
· 想像 1941 年戰俘營的首演場景`,
"zYpBFPH7tMk");

nawm(C,"NAWM 186","Steve Reich","Piano Phase · 1967",
`· 相位音樂 Phase music 代表作
· 兩架鋼琴演奏同一 12 音符圖案
   E/F♯/B/C♯/D（E 羽調式）
· 其中一架鋼琴緩慢加速
   → 相位差產生 → 新複合節奏
· 加速至半拍錯位後再穩定
   → 再次加速 → 循環
· 非洲加納鼓樂多層節奏原理
· 與歐洲序列主義的截然對立
· 極簡主義美學：有限素材無限感知
· 可一人演奏一人錄音，或兩人現場
· 1967 年創作，與反文化運動同步`,
`· 辨認基本 12 音符圖案的音高
· 感受兩聲部如何漸漸錯開
· 注意高音旋律在錯位中浮現
· 聆聽和聲層次如何逐漸豐富
· 思考：重複 = 無聊 or 豐富？
· 感受「相位」的精確漸變過程
· 辨認每個穩定位置（共12種）
· 與 Schoenberg 十二音列美學對立
· 思考：非洲節奏如何影響美國作曲
· 感受最後回到同步的奇特感覺`,
"g0WVh1D0-6Q");

{ const s=ls(C); hdr(s,C,"重要術語 · Period 9","Key Terms: Late 20th Century to Present",false);
lp(s,C,"術語 Terms A",
`· Total serialism 整體序列
· Aleatory music 偶然音樂
· Musique concrète 具體音樂
· Electronic music 電子音樂
· Minimalism 極簡主義
· Phase music 相位音樂
· Additive process 加法過程
· Tintinnabuli 鈴聲風格（Pärt）
· Spectralism 頻譜音樂
· New complexity 新複雜主義
· Extended techniques 延伸技法
· Prepared piano 加料鋼琴（Cage）`,
"術語 Terms B",
`· Postmodernism 後現代主義
· Polystylism 多風格主義
· Quotation / Allusion 引用
· Intertextuality 跨文本互涉
· World music 世界音樂融合
· HIP 歷史演奏實踐
· Crossover 跨界音樂
· Ambient music 環境音樂
· IRCAM 電腦音樂研究院
· Fluxus 激浪派（60s 觀念藝術）
· Neo-Romanticism 新浪漫主義
· Aleatory counterpoint 偶然對位`); }

{ const s=ls(C); hdr(s,C,"時間軸 · Period 9","Timeline: Late 20th Century to Present",false);
lp(s,C,"1945–1975",
`· 1945  二戰結束，戰後重建開始
· 1941  Messiaen · Quartet（戰俘營）
· 1946  Darmstadt 前衛音樂暑校
· 1948  Schaeffer · 具體音樂誕生
· 1952  Cage · 4'33''（沉默作品）
· 1955  Boulez · Le Marteau
· 1956  Stockhausen · Gesang
· 1964  Riley · In C（早期極簡）
· 1967  Reich · Piano Phase
· 1971  Pärt 轉型 Tintinnabuli`,
"1975–2020s",
`· 1976  Glass · Einstein on the Beach
· 1987  Adams · Nixon in China
· 1989  冷戰結束，柏林圍牆倒塌
· 1997  Adès · Asyla Op.17
· 1999  Napster 改變音樂消費
· 2001  Thomas Adès 登上國際舞台
· 2006  Spotify 串流平台成立
· 2010s 古典音樂邊緣化危機
· 2020  COVID 迫使線上音樂會
· 2020s AI 輔助作曲新議題`); }

{ const s=ls(C); hdr(s,C,"考試重點 · Period 9","Exam Focus: Late 20th Century to Present",false);
lp(s,C,"必記作曲家/作品",
`· Messiaen · Quartet（NAWM 180）
   戰俘營，Mvt. VIII 讚歌
· Reich · Piano Phase（NAWM 186）
   相位音樂，極簡主義典範
· Adès · Asyla（NAWM 192）
   21c 英國，後現代融合
· Cage · 4'33''（1952）
   觀念音樂，沉默的意義
· Pärt · Cantus / Spiegel
   Tintinnabuli 鈴聲風格`,
"必懂概念",
`· Total serialism 整體序列操作
   音高/節奏/力度/音色全控制
· Aleatory = 偶然但非隨意
   框架內的自由選擇
· Minimalism 三種核心技法
   重複 / 相位 / 加法過程
· 後現代 = 多元、引用、無大敘事
   打破前衛/流行的邊界
· 世界音樂融合的倫理問題
   文化挪用 vs. 真誠交流
· 數位時代古典音樂的角色`); }

{ const s=ds(C); hdr(s,C,"全書回顧","Final Review: A History of Western Music");
dp(s,C,"九大時期速覽 I",
`1. 古代/教會（3000 BCE–900 CE）
   單聲部，調式，禮儀功能
   代表：Seikilos · Gregorian Chant
2. 中世紀（900–1420）
   複音起源，世俗/教堂並進
   代表：Machaut · Messe de N.D.
3. 文藝復興（1420–1600）
   模仿對位，印刷術，牧歌
   代表：Josquin · Palestrina
4. 早期巴洛克（1600–1680）
   歌劇誕生，通奏低音，情感論
   代表：Monteverdi · L'Orfeo
5. 晚期巴洛克（1680–1750）
   賦格頂峰，協奏曲，Bach/Handel
   代表：Bach Brandenburg · Messiah`,
"九大時期速覽 II",
`6. 古典時期（1750–1810）
   奏鳴曲式，清晰均衡，三大師
   代表：Beethoven Sym.5 · Haydn 弦四
7. 浪漫主義（1815–1900）
   半音主義，Lied，Leitmotif
   代表：Schubert · Chopin · Wagner
8. 二十世紀前半（1890–1945）
   無調性，十二音，爵士，春之祭
   代表：Schoenberg · Stravinsky
9. 二十世紀後半至當代（1945–）
   序列，極簡，後現代，全球化
   代表：Reich · Cage · Adès
— 全書共 39 章 · 901 張投影片 —`); }
}

pres.writeFile({ fileName: "Condensed_Review.pptx" }).then(() => console.log("Done: Condensed_Review.pptx"));
