'use strict';
// A3 Landscape — 百大 Q&A · 西方音樂史
// 100 Q&A distributed across 9 periods (Ch01–39), bilingual (ZH + key English terms),
// with Burkholder 10th ed. page references.
const pptx = require('pptxgenjs');
const pres = new pptx();
pres.defineLayout({ name:'A3_LAND', width:16.54, height:11.69 });
pres.layout = 'A3_LAND';

// ── Global palette ───────────────────────────────────────────────────────────
const G = {
  bg:       'F8F3E8',
  dark:     '1A1008',
  mid:      '4A3828',
  gold:     '8B6F3C',
  accent:   'B83828',   // page refs
  line:     'B8A888',
  paleLine: 'D8CEB8',
};

// ── Period data (9 blocks × 11-12 Q&A = 100) ─────────────────────────────────
// Per-period color pair: header background + header text
// qa item fields:
//   q : question (essay 申論 / term 名詞 / short 短答 — mixed ~4/4/3 per period)
//   a : concise answer (bilingual, key English terms preserved)
//   p : textbook page(s), Burkholder 10th ed.
const PERIODS = [
  // ── P1 ─ 古代與教會音樂 (Ch01–03) ─────────────────────────────────────────
  { n:1, zh:'古代與教會音樂', en:'Ancient World & Early Church',
    chs:'Ch 1–3', pages:'pp. 4–62',
    bg:'E8D4B8', dk:'5A3818', qa:[
    { q:'[申] 比較 Plato 與 Aristotle 對音樂教育觀點之差異。',
      a:'Plato 重道德—調式 ethos 塑造品格；Aristotle 重情感—音樂可娛樂亦可淨化 (catharsis)，教育與享樂並行。',
      p:'pp. 11-15' },
    { q:'[名] Ethos（情態論）',
      a:'古希臘思想：各調式具特定道德力量（Dorian 莊重／Phrygian 激昂／Lydian 柔弱），影響聽者品格。',
      p:'pp. 13-15' },
    { q:'[短] 現存最完整之古希臘樂譜為？',
      a:'Seikilos Epitaph（塞基洛斯墓誌銘，1c CE，刻於小亞細亞墓碑）。',
      p:'pp. 17-18' },
    { q:'[申] 說明 Gregorian chant 的形成與統一意義。',
      a:'羅馬與高盧 (Gallican) 禮儀合流 (8–9c)；查理曼推動整合，託名教宗 Gregory I 神化傳統，成西方禮儀聖詠之核心。',
      p:'pp. 42-47' },
    { q:'[名] Neume（紐瑪記譜）',
      a:'9c 起的無固定音高符號，示旋律走向；Guido d\'Arezzo 後發展為四線譜方形譜，現代五線譜前身。',
      p:'pp. 28-33' },
    { q:'[短] Boethius 的主要貢獻？',
      a:'著《De institutione musica》將希臘音樂理論譯入拉丁中世紀。',
      p:'pp. 34-36' },
    { q:'[申] 彌撒結構：Ordinary 與 Proper 差異？',
      a:'Ordinary 文字固定五段（Kyrie／Gloria／Credo／Sanctus／Agnus Dei）；Proper 依節期變動（Introit／Gradual／Alleluia／Offertory／Communion）。',
      p:'pp. 42-47' },
    { q:'[名] Responsorial vs Antiphonal',
      a:'前者：獨唱對會眾輪唱（如 Gradual）；後者：兩合唱隊輪唱（如 Introit）。',
      p:'pp. 47-50' },
    { q:'[短] Hildegard of Bingen 身分為何？',
      a:'12c 德國女修院長、神秘家、作曲家；代表作《Ordo virtutum》。',
      p:'pp. 60-62' },
    { q:'[申] Guido d\'Arezzo 對記譜法的三大貢獻？',
      a:'(1) 四線譜 staff；(2) Solmization 六聲音階 ut-re-mi-fa-sol-la；(3) Guidonian Hand 手型記憶法，加速唱誦教學。',
      p:'pp. 36-40' },
    { q:'[名] Trope 與 Sequence',
      a:'Trope：於既有聖詠插入新詞／新樂；Sequence：源於 Alleluia 後綴旋律加詞，獨立成體裁（Notker／Hildegard）。',
      p:'pp. 56-59' },
  ]},

  // ── P2 ─ 中世紀世俗與複音 (Ch04–06) ───────────────────────────────────────
  { n:2, zh:'中世紀世俗與複音', en:'Medieval Secular Song & Polyphony',
    chs:'Ch 4–6', pages:'pp. 63–132',
    bg:'D4D8EC', dk:'223060', qa:[
    { q:'[申] 比較 troubadour 與 trouvère。',
      a:'同為中世紀世俗詩人—作曲家。Troubadour：法國南部 Occitan 語、12c 盛期；Trouvère：法國北部 Old French 語、延續至 13c 晚；主題皆宮廷愛情 fin\'amor。',
      p:'pp. 69-75' },
    { q:'[名] Organum（早期複音）',
      a:'9c 起以聖詠 (cantus firmus) 為定旋律，上加一或多聲部；由平行 → 斜向 → 自由 → 裝飾性 organum 發展。',
      p:'pp. 81-85' },
    { q:'[短] Notre Dame 樂派代表作曲家？',
      a:'Léonin（《Magnus liber organi》2 聲部）；Pérotin（擴為 3–4 聲部）。',
      p:'pp. 86-92' },
    { q:'[申] 何謂 Ars Nova？列舉其主要創新。',
      a:'Philippe de Vitry 論著（ca.1320）—(1) Isorhythm 等節奏；(2) 二分音符 (minim) 成新時值；(3) 新記譜容許完美／不完美拍，擴節奏自由度。',
      p:'pp. 108-113' },
    { q:'[名] Isorhythm（等節奏）',
      a:'重複固定節奏模式 talea 與旋律模式 color，二者長度錯位產生結構張力；多用於 14c motet tenor。',
      p:'pp. 112-114' },
    { q:'[短] Machaut 最著名宗教作品？',
      a:'《Messe de Nostre Dame》(1360s)—現存最早由單一作曲家譜成完整彌撒 Ordinary。',
      p:'pp. 114-120' },
    { q:'[申] 說明 formes fixes 三種形式。',
      a:'Ballade (aabC) 敘事；Rondeau (ABaAabAB) 輪唱；Virelai (AbbaA) 舞曲性。由 Machaut 確立，主宰 14–15c 世俗歌曲。',
      p:'pp. 120-121' },
    { q:'[名] Motet（經文歌）',
      a:'13c 新體裁：tenor 取自聖詠，上聲部加不同歌詞（常多語並存），從禮儀邊緣走向世俗化。',
      p:'pp. 96-102' },
    { q:'[短] Ars Subtilior 的特徵？',
      a:'14c 晚期法南宮廷樂風—節奏極複雜、記譜華麗（心形／圓形樂譜），為精英聽眾而作。',
      p:'pp. 121-122' },
    { q:'[申] 14c 義大利 Trecento 與法國 Ars Nova 差異？',
      a:'義大利 Trecento：旋律抒情流暢（madrigal／caccia／ballata，Landini）；法國：節奏結構嚴密（isorhythm／formes fixes，Machaut）。',
      p:'pp. 122-128' },
    { q:'[名] Landini cadence',
      a:'14c 常見終止：上聲部 7-6-8（如 mi-re-mi 下加 do），以 Landini 為名，成 14–15c 標誌。',
      p:'pp. 126-128' },
  ]},

  // ── P3 ─ 文藝復興 (Ch07–11) ──────────────────────────────────────────────
  { n:3, zh:'文藝復興', en:'Renaissance',
    chs:'Ch 7–11', pages:'pp. 136–253',
    bg:'D4E8C8', dk:'1E4828', qa:[
    { q:'[申] 文藝復興音樂三大核心理念？',
      a:'(1) 人文主義—歌詞可感清晰；(2) 模仿對位 imitation 為基本織體；(3) 古希臘理想復興 + Zarlino 三和弦理論。',
      p:'pp. 137-154' },
    { q:'[名] Fauxbourdon（假低音）',
      a:'15c 英／勃艮第技法—上下聲部記譜、中聲部以平行四度即興填充，產生連續 6-3 和絃；Dufay 用於聖母頌歌。',
      p:'pp. 163-166' },
    { q:'[短] Dunstable 對歐陸貢獻？',
      a:'輸出 La Contenance angloise（英式風格）—三和弦音響、甜美和聲，影響 Dufay 與 Franco-Flemish 樂派。',
      p:'pp. 159-163' },
    { q:'[申] 比較 Josquin 與 Palestrina 對位風格。',
      a:'Josquin（1500 前後）：模仿嚴整、段落分明、文字音畫萌芽。Palestrina（特倫托後）：柔和線條、嚴格預備解決、整體均衡，成「古典對位」典範。',
      p:'pp. 192-203, 234-243' },
    { q:'[名] Cantus firmus Mass',
      a:'以既有旋律（聖詠或世俗歌曲）置於 tenor 作為結構骨幹之彌撒；Dufay《Se la face ay pale》彌撒為經典。',
      p:'pp. 171-175' },
    { q:'[短] Paraphrase 與 Parody Mass 差異？',
      a:'Paraphrase：取單一旋律並裝飾於各聲部；Parody (Imitation)：取整首複音作品借其和聲／動機重組。',
      p:'pp. 195-199' },
    { q:'[申] Word painting 在 madrigal 中的作用。',
      a:'以音樂具體描繪歌詞字面義（飛翔→上行／死亡→半音下行）；Weelkes、Gesualdo 極端化，成 16c 晚期義／英牧歌手法典型。',
      p:'pp. 215-227' },
    { q:'[名] Musica reservata',
      a:'16c 晚期術語—指為懂行聽眾保留之精緻音樂（Lassus 代表），強調歌詞情感與修辭手法的深度表現。',
      p:'pp. 209-215' },
    { q:'[短] Petrucci 印刷的劃時代貢獻？',
      a:'1501 威尼斯出版《Odhecaton A》—最早活字印刷複音樂譜集，推動音樂流通與樂派交流。',
      p:'pp. 156-157' },
    { q:'[申] Council of Trent 對教會音樂的影響？',
      a:'要求歌詞清晰、禁世俗旋律為基礎；Palestrina《Missa Papae Marcelli》傳說「救了複音」，確立反宗教改革音樂風範。',
      p:'pp. 234-240' },
    { q:'[名] Chorale（新教會眾讚美詩）',
      a:'Luther 為新教會眾創作之德語讚美詩，旋律簡單會眾易唱；成 Bach chorale prelude／cantata 核心素材。',
      p:'pp. 229-233' },
  ]},

  // ── P4 ─ 巴洛克早期 (Ch12–16) ────────────────────────────────────────────
  { n:4, zh:'巴洛克早期', en:'Early Baroque',
    chs:'Ch 12–16', pages:'pp. 254–370',
    bg:'E8D8B0', dk:'4A2E10', qa:[
    { q:'[申] 文藝復興 → 巴洛克的三大風格轉變？',
      a:'(1) 複音 → 數字低音 basso continuo 主調織體；(2) 均衡 → 對比（tutti/solo、強弱、快慢）；(3) 理性均衡 → 激情表達 (doctrine of affections)。',
      p:'pp. 278-296' },
    { q:'[名] Basso continuo（數字低音）',
      a:'低音線 + 數字和弦記號；由鍵盤（風琴／大鍵琴）+ 低音樂器（cello／bassoon）即興填充和聲。巴洛克核心織體。',
      p:'pp. 282-285' },
    { q:'[短] 誰寫了第一齣現存歌劇？',
      a:'Jacopo Peri《Euridice》(1600)，佛羅倫斯 Camerata 圈子產物。',
      p:'pp. 301-304' },
    { q:'[申] Monteverdi 為何是「兩種實踐」的橋樑？',
      a:'提出 prima prattica（古式—對位嚴整）與 seconda prattica（新式—歌詞主宰音樂，容許不協和）；由 Madrigali Bk.5 → Orfeo → Poppea 實踐於歌劇。',
      p:'pp. 289-295, 306-315' },
    { q:'[名] Stile concertato（協奏風格）',
      a:'對比式寫作—聲部或樂器群與獨奏者間呼應；Gabrieli《Sonata pian\'e forte》、Schütz 德式 Concerto 均採之。',
      p:'pp. 320-327' },
    { q:'[短] Monteverdi《Orfeo》(1607) 的意義？',
      a:'首部結構完整的歌劇—融 recitative／aria／合唱／器樂，奠定歌劇體裁雛形。',
      p:'pp. 310-315' },
    { q:'[申] 比較威尼斯與羅馬早期巴洛克風格。',
      a:'威尼斯（Gabrieli, Monteverdi）：多合唱 cori spezzati、奢華器樂、歌劇商業化。羅馬（Carissimi）：反宗教改革下偏 oratorio，節制、虔敬。',
      p:'pp. 320-332' },
    { q:'[名] Recitative vs Aria',
      a:'Recitative：宣敘，近語言節奏、推進劇情；Aria：詠嘆，正規拍節旋律化、抒情定格；巴洛克歌劇二者交替。',
      p:'pp. 306-315' },
    { q:'[短] 法國 tragédie lyrique 奠基者？',
      a:'Lully（1670s–80s 路易十四宮廷）—融合法國戲劇、舞蹈、合唱、節奏規整。',
      p:'pp. 340-348' },
    { q:'[申] 英國 17c 音樂戲劇特殊發展？',
      a:'Restoration (1660) 後產生 semi-opera（Purcell《Dido and Aeneas》為少數全唱歌劇），融合話劇與音樂；Purcell 綜合義法英傳統。',
      p:'pp. 355-363' },
    { q:'[名] Concerted Madrigal（協奏牧歌）',
      a:'Monteverdi Bk.7-8 代表—數字低音伴奏、獨唱／小組對比、打破 5 聲部傳統，連接牧歌與 cantata。',
      p:'pp. 317-320' },
  ]},

  // ── P5 ─ 巴洛克晚期 (Ch17–19) ────────────────────────────────────────────
  { n:5, zh:'巴洛克晚期', en:'Late Baroque',
    chs:'Ch 17–19', pages:'pp. 371–453',
    bg:'D8C0E0', dk:'3A1848', qa:[
    { q:'[申] 義大利晚期巴洛克三大體裁創新？',
      a:'(1) Corelli 確立 trio sonata 與 concerto grosso；(2) Vivaldi 發展獨奏協奏曲與 ritornello 形式；(3) A. Scarlatti 義式序曲 sinfonia／D. Scarlatti 鍵盤 sonata。',
      p:'pp. 371-401' },
    { q:'[名] Ritornello form',
      a:'巴洛克協奏曲結構—全奏主題 ritornello 於不同調多次再現，獨奏段落 episodes 穿插其間；Vivaldi 首確立。',
      p:'pp. 386-391' },
    { q:'[短] Corelli 最重要體裁？',
      a:'Trio sonata（2 高音 + 通奏低音）與 concerto grosso（concertino vs ripieno）。',
      p:'pp. 375-383' },
    { q:'[申] Vivaldi《四季》體現什麼美學？',
      a:'程式音樂—以十四行詩對應樂章；獨奏小提琴技巧華麗；展示 ritornello 形式彈性與描繪能力。',
      p:'pp. 388-391' },
    { q:'[名] Da capo aria',
      a:'ABA 三段式詠嘆調：B 段對比調性／情感，A 再現時歌手即興裝飾；義大利 opera seria 核心。',
      p:'pp. 395-399' },
    { q:'[短] 法國序曲的結構？',
      a:'慢—快—慢（附點慢段 + 賦格快段）；Lully 確立，Bach／Handel 沿用。',
      p:'pp. 402-410' },
    { q:'[申] 比較 Bach 與 Handel 的職涯與風格。',
      a:'Bach：Lutheran 德國教堂／宮廷作曲家—cantata／passion／organ 為主，對位嚴整；Handel：國際化（德→義→英）—歌劇、oratorio 為主，戲劇性強。',
      p:'pp. 424-453' },
    { q:'[名] Passion（受難曲）',
      a:'依福音書敘述基督受難，結合 recitative（福音傳道者）、aria、chorus、chorale；Bach《Matthäus-Passion》BWV 244 典範。',
      p:'pp. 446-450' },
    { q:'[短] Telemann 與 Bach 時人評價對比？',
      a:'Telemann 在世時最著名（出版量大、走 galant 風）；Bach 當時被視為過時；後世評價逆轉。',
      p:'pp. 424-428' },
    { q:'[申] 何謂 Doctrine of Affections（情感論）？',
      a:'巴洛克美學—每曲／每段表達單一明確情緒（愛／怒／悲／喜），以固定音型、節奏、調性對應之，影響 aria、fugue 等形式。',
      p:'pp. 282-285, 395-399' },
    { q:'[名] Cantata（清唱劇）',
      a:'獨唱／合唱 + 器樂伴奏，含 recitative／aria 多樂章；Bach 寫 200+ Lutheran cantata（教會年禮用）。',
      p:'pp. 442-446' },
  ]},

  // ── P6 ─ 古典時期 (Ch20–23) ──────────────────────────────────────────────
  { n:6, zh:'古典時期', en:'Classical Era',
    chs:'Ch 20–23', pages:'pp. 454–553',
    bg:'C8E0D0', dk:'1E482A', qa:[
    { q:'[申] Galant 與 Empfindsam 差異？',
      a:'Galant：義法主流—輕巧優雅、短樂句、主調織體。Empfindsam：C.P.E. Bach 德北—「敏感」突然情緒轉折、裝飾多、富戲劇性。',
      p:'pp. 454-468' },
    { q:'[名] Sonata form（奏鳴曲式）',
      a:'三部分：呈示部（主題 I 主調 → 主題 II 屬調）／發展部（調性漫遊與動機開展）／再現部（主題皆回主調）；18c 中葉確立。',
      p:'pp. 494-502' },
    { q:'[短] Opera seria 改革者？',
      a:'Gluck《Orfeo ed Euridice》(1762)—去華麗裝飾、重戲劇真實、合唱回歸。',
      p:'pp. 476-482' },
    { q:'[申] Haydn 對交響曲的貢獻？',
      a:'寫 104 首—確立四樂章結構（快—慢—minuet—終曲）、主題—動機發展；倫敦交響曲群達巔峰，「交響曲之父」之稱名副其實。',
      p:'pp. 515-528' },
    { q:'[名] Minuet and Trio',
      a:'古典交響曲／四重奏第三樂章—三拍子宮廷舞曲 + 對比中段 trio + 小步舞曲再現；Beethoven 後改為 scherzo。',
      p:'pp. 520-522' },
    { q:'[短] Mozart 三大 da Ponte 歌劇？',
      a:'Le nozze di Figaro (1786)、Don Giovanni (1787)、Così fan tutte (1790)。',
      p:'pp. 540-547' },
    { q:'[申] Haydn 與 Mozart 的弦樂四重奏風格差異？',
      a:'Haydn：單一動機開展、幽默、Op.33「俄羅斯」確立 4 聲部對等對話。Mozart：旋律豐富歌唱性、獻給 Haydn 的 K.387-465 六首展現深刻對位。',
      p:'pp. 528-535' },
    { q:'[名] Classical Concerto form',
      a:'雙呈示部—先樂團呈示、再獨奏呈示；發展部 + 再現部後有獨奏 cadenza；Mozart 鋼琴協奏曲為典範。',
      p:'pp. 503-507' },
    { q:'[短] Singspiel 是什麼？',
      a:'德語歌劇—唱段 + 口語對白，通俗性；Mozart《Zauberflöte》(1791) 代表，Beethoven《Fidelio》延續。',
      p:'pp. 483-488' },
    { q:'[申] 古典器樂各體裁的同構性為何？',
      a:'皆四樂章、第一樂章奏鳴曲式主導、主題—動機發展為共同手法；弦樂四重奏 = 抽象最純粹之交響曲，Haydn 視為「對話」。',
      p:'pp. 494-502' },
    { q:'[名] Rondo form',
      a:'A-B-A-C-A-B-A 類型—主題 refrain 反覆穿插對比段；古典末樂章常用，活潑收尾。',
      p:'pp. 499-502' },
  ]},

  // ── P7 ─ 浪漫時期 (Ch24–30) ──────────────────────────────────────────────
  { n:7, zh:'浪漫時期', en:'Romantic Era',
    chs:'Ch 24–30', pages:'pp. 554–755',
    bg:'E8C8CC', dk:'60182A', qa:[
    { q:'[申] Beethoven 為何是古典—浪漫之橋樑？',
      a:'早期延續 Haydn-Mozart；中期 (Eroica 1803) 擴張形式、加入標題性與英雄主題；晚期四重奏 Op.131 探索極端對比與賦格，啟發浪漫作曲家突破形式。',
      p:'pp. 554-579' },
    { q:'[名] Lied（藝術歌曲）',
      a:'德語獨唱藝術歌曲 + 鋼琴伴奏，詩樂結合；Schubert《Erlkönig》《Winterreise》、Schumann《Dichterliebe》、Brahms 為代表。',
      p:'pp. 580-594' },
    { q:'[短] Chopin 作品幾乎全為哪件樂器？',
      a:'鋼琴—Études、Préludes、Ballades、Nocturnes、Mazurkas、Polonaises。',
      p:'pp. 601-612' },
    { q:'[申] 標題音樂 vs 絕對音樂之爭？',
      a:'標題派（新德意志樂派 Berlioz／Liszt）主張音樂可敘事—Liszt 交響詩 symphonic poem；Brahms／Hanslick 主張音樂自足形式美；19c 下半葉美學核心辯論。',
      p:'pp. 622-645, 725-730' },
    { q:'[名] Idée fixe（固定樂念）',
      a:'Berlioz《Symphonie fantastique》中代表戀人的主題，於全曲 5 樂章變形再現；為 Wagner leitmotif 先聲。',
      p:'pp. 560-565' },
    { q:'[短] Wagner 的 Gesamtkunstwerk？',
      a:'「整體藝術作品」—詩、樂、戲、舞台合一；《Der Ring des Nibelungen》為實踐。',
      p:'pp. 684-703' },
    { q:'[申] 比較 Verdi 與 Wagner 歌劇美學。',
      a:'Verdi：延續義大利 aria／ensemble 傳統、旋律優先、戲劇緊湊。Wagner：連續樂流 (unendliche Melodie)、leitmotif 系統、樂隊戲劇主角、反傳統分曲結構。',
      p:'pp. 675-710' },
    { q:'[名] Leitmotif（主導動機）',
      a:'代表人物／物件／情感之短音型，於劇中循環變形推動敘事；Wagner 樂劇 Musikdrama 基本手法。',
      p:'pp. 692-698' },
    { q:'[短] Brahms 第一交響曲為何拖 21 年？',
      a:'Beethoven 陰影—自覺需承交響曲傳統；1876 終完成 c 小調，被稱「Beethoven 第十號」。',
      p:'pp. 636-641' },
    { q:'[申] 19c 末民族樂派興起之因與代表？',
      a:'民族自覺 + 反德奧霸權—俄五人組 (Mighty Handful: Mussorgsky《Pictures》, Rimsky-Korsakov)、捷克 (Smetana, Dvořák)、芬蘭 (Sibelius) 採民間素材、民族調式。',
      p:'pp. 731-755' },
    { q:'[名] Character piece（特性曲）',
      a:'19c 鋼琴短曲體裁，單一情感／意象—Schumann《Kinderszenen》、Chopin Nocturne、Mendelssohn《Lieder ohne Worte》；浪漫縮影。',
      p:'pp. 594-601' },
  ]},

  // ── P8 ─ 20 世紀前半 (Ch31–35) ───────────────────────────────────────────
  { n:8, zh:'20 世紀前半', en:'Early 20th Century',
    chs:'Ch 31–35', pages:'pp. 756–897',
    bg:'B8D8D8', dk:'103838', qa:[
    { q:'[申] Debussy 印象主義音樂語言特徵？',
      a:'全音階 whole-tone、五聲音階、教會調式並用；和聲脫離功能進行；音色與音響優先於結構；《Prélude à l\'après-midi d\'un faune》《La mer》《Nocturnes》為代表。',
      p:'pp. 772-780' },
    { q:'[名] Serialism / Twelve-tone',
      a:'Schoenberg (1920s)—十二半音排序為 tone row，依原型 P／逆行 R／倒影 I／逆行倒影 RI 排列；消除主音中心。',
      p:'pp. 820-830' },
    { q:'[短] Stravinsky《Le Sacre du printemps》首演？',
      a:'1913 年巴黎—因不協和、變拍、狂暴節奏引發暴動式騷亂。',
      p:'pp. 814-819' },
    { q:'[申] Bartók 如何融合民間音樂與現代技法？',
      a:'蒐集匈牙利／羅馬尼亞／斯洛伐克民歌 → 提取調式與不對稱節奏 (Bulgarian rhythms)；結合對位與軸心調性 axis tonality；《Music for Strings, Percussion and Celesta》典範。',
      p:'pp. 869-877' },
    { q:'[名] Sprechstimme（口白聲）',
      a:'介於說話與歌唱間之發聲—記音高走向但不定具體音高；Schoenberg《Pierrot lunaire》(1912) 首用。',
      p:'pp. 826-830' },
    { q:'[短] 美國實驗音樂之父？',
      a:'Charles Ives (1874-1954)—拼貼、引用、複調性、多重節奏並置；1940s 後才被廣泛發掘。',
      p:'pp. 763-769' },
    { q:'[申] 新古典主義 (Neoclassicism) 之美學主張？',
      a:'反浪漫誇張與印象模糊，回歸巴洛克／古典形式（concerto grosso、對位）與清晰織體；Stravinsky《Pulcinella》(1920) 始，Hindemith、早期 Prokofiev 跟進。',
      p:'pp. 880-887' },
    { q:'[名] Gebrauchsmusik（實用音樂）',
      a:'Hindemith 等人提倡—為業餘／社群實用之音樂，反職業專家化與新音樂過度前衛；含合唱、校園音樂作品。',
      p:'pp. 887-890' },
    { q:'[短] 爵士樂發源時代背景？',
      a:'19c 末–20c 初紐奧良—非裔美國人 blues + ragtime + 軍樂 + spirituals 融合；1920s 北傳 Chicago → NY。',
      p:'pp. 855-868' },
    { q:'[申] Socialist Realism 對 Shostakovich 的影響？',
      a:'1936《Pravda》批《Lady Macbeth of Mtsensk》後，作《Fifth Symphony》(1937) 副題「對公正批評之創造性回應」重新自塑；公開服從—私下反諷的雙重語言。',
      p:'pp. 890-897' },
    { q:'[名] Futurism / Bruitisme',
      a:'Russolo《L\'arte dei rumori》(1913) 提出「噪音音樂」—使用 intonarumori 機械噪音樂器，影響後 musique concrète。',
      p:'pp. 770-772' },
  ]},

  // ── P9 ─ 20 世紀後半至當代 (Ch36–39) ─────────────────────────────────────
  // 12 Q&A (to reach 100 total)
  { n:9, zh:'20 世紀後半至當代', en:'Postwar & Contemporary',
    chs:'Ch 36–39', pages:'pp. 898–1020',
    bg:'C8D0E0', dk:'1C2838', qa:[
    { q:'[申] Total serialism 與 Darmstadt 學派？',
      a:'戰後歐洲—Boulez、Stockhausen 將序列擴至音高／節奏／力度／音色；《Structures I》《Kreuzspiel》為代表，追求理性極致。',
      p:'pp. 898-910' },
    { q:'[名] Aleatoric / Indeterminacy',
      a:'Cage《4\'33"》《Imaginary Landscape》—以擲硬幣、《易經》、機遇決定音樂元素；演奏者與環境成作品之一部分。',
      p:'pp. 910-918' },
    { q:'[短] Cage《4\'33"》首演年代？',
      a:'1952 年 Woodstock NY，由 David Tudor 首演—全場靜默，讓環境聲成音樂。',
      p:'pp. 910-914' },
    { q:'[申] Minimalism 之興起與特徵？',
      a:'1960s 美國—La Monte Young、Riley《In C》、Reich phasing、Glass additive process；反序列主義複雜化；脈動反覆、漸進變化、調性回歸。',
      p:'pp. 940-950' },
    { q:'[名] Phasing（相位移）',
      a:'Reich 技法—兩件相同材料以極緩速度錯位，產生新節奏／和聲關係；《Piano Phase》《Clapping Music》為例。',
      p:'pp. 942-946' },
    { q:'[短] 電子音樂里程碑之作？',
      a:'Stockhausen《Gesang der Jünglinge》(1956)—錄音室合成與人聲結合，電子音樂里程碑。',
      p:'pp. 906-910' },
    { q:'[申] Postmodernism 音樂特徵？',
      a:'拒絕統一風格—引用、拼貼、風格並置（Schnittke polystylism、Rochberg 重返調性）；反 modernist 純粹性，擁抱低俗／流行。',
      p:'pp. 954-975' },
    { q:'[名] Spectralism（光譜樂派）',
      a:'1970s 法國 Grisey、Murail—以音響泛音結構 spectrum 為素材，電腦分析聲音本身；《Partiels》為代表。',
      p:'pp. 968-975' },
    { q:'[短] Pärt 之 tintinnabuli 技法？',
      a:'「鐘鳴」—三和弦持續音 + 旋律行進交織；《Spiegel im Spiegel》《Fratres》典範，新靈性主義。',
      p:'pp. 975-980' },
    { q:'[申] 21c 全球化與技術對音樂之衝擊？',
      a:'網路傳播 + 數位製作民主化；跨文化融合（Tan Dun、Golijov）；演算法作曲、AI 參與；傳統邊界消弭。',
      p:'pp. 990-1020' },
    { q:'[名] Experimental notation',
      a:'Crumb 圖形譜、Penderecki tone cluster 符號、open notation；20c 後半作曲家擴張記譜以傳達新音響。',
      p:'pp. 904-908' },
    { q:'[短] John Adams 代表歌劇？',
      a:'《Nixon in China》(1987)、《Doctor Atomic》(2005)—post-minimalist 歌劇、時事題材。',
      p:'pp. 956-962' },
  ]},
];

// ── Sanity check: 100 Q&A total ──────────────────────────────────────────────
const totalQA = PERIODS.reduce((s, p) => s + p.qa.length, 0);
if (totalQA !== 100) {
  console.error(`!!! Expected 100 Q&A, got ${totalQA}`);
  PERIODS.forEach(p => console.error(`  P${p.n} ${p.zh}: ${p.qa.length}`));
  process.exit(1);
}

// ── Build slide ──────────────────────────────────────────────────────────────
const PAGE_W = 16.54, PAGE_H = 11.69;
const s = pres.addSlide();
s.background = { color: G.bg };

// Top title bar (thin gold line top + bottom)
s.addShape(pres.ShapeType.rect, { x:0, y:0, w:PAGE_W, h:0.08, fill:{color:G.gold}, line:{color:G.gold} });

// Title (zh) — left
s.addText('西方音樂史 · 百大 Q&A', {
  x:0.25, y:0.14, w:7.5, h:0.42,
  fontSize:22, bold:true, color:G.dark, fontFace:'Georgia', valign:'middle', margin:0,
});

// Subtitle (en) — left beneath title
s.addText('100 Questions & Answers · A History of Western Music (Burkholder 10e)', {
  x:0.25, y:0.52, w:8.5, h:0.22,
  fontSize:9, italic:true, color:G.mid, fontFace:'Georgia', valign:'middle', margin:0,
});

// Legend — center-right
s.addText('[申] 申論 Essay  ·  [名] 名詞 Term  ·  [短] 短答 Short', {
  x:8.0, y:0.15, w:4.6, h:0.24,
  fontSize:9, color:G.mid, fontFace:'Calibri', align:'right', valign:'middle', margin:0,
});
s.addText('頁碼對應 Burkholder《A History of Western Music》10th ed.', {
  x:8.0, y:0.38, w:4.6, h:0.22,
  fontSize:8, color:G.mid, fontFace:'Calibri', italic:true, align:'right', valign:'middle', margin:0,
});

// Meta — right edge
s.addText('9 Periods × ~11 Q&A', {
  x:12.8, y:0.15, w:3.5, h:0.24,
  fontSize:10, bold:true, color:G.gold, fontFace:'Georgia', align:'right', valign:'middle', margin:0,
});
s.addText('Ch 1–39  ·  A3 Landscape', {
  x:12.8, y:0.38, w:3.5, h:0.22,
  fontSize:8, color:G.mid, fontFace:'Calibri', italic:true, align:'right', valign:'middle', margin:0,
});

// ── 9 Period blocks — 3×3 grid ───────────────────────────────────────────────
const GRID_TOP = 0.82;
const GRID_BOTTOM = 11.55;
const GRID_LEFT = 0.15;
const GRID_RIGHT = 16.39;
const GAP = 0.08;

const BLOCK_W = (GRID_RIGHT - GRID_LEFT - 2*GAP) / 3;   // ~5.37
const BLOCK_H = (GRID_BOTTOM - GRID_TOP - 2*GAP) / 3;   // ~3.52
const COL_X = [ GRID_LEFT,
                GRID_LEFT + BLOCK_W + GAP,
                GRID_LEFT + 2*(BLOCK_W + GAP) ];
const ROW_Y = [ GRID_TOP,
                GRID_TOP + BLOCK_H + GAP,
                GRID_TOP + 2*(BLOCK_H + GAP) ];

function addPeriodBlock(period, bx, by) {
  // Block outer border (thin)
  s.addShape(pres.ShapeType.rect, {
    x:bx, y:by, w:BLOCK_W, h:BLOCK_H,
    fill:{color:'FFFFFF'}, line:{color:period.bg, width:1.25},
  });

  // Period header band
  const HEAD_H = 0.44;
  s.addShape(pres.ShapeType.rect, {
    x:bx, y:by, w:BLOCK_W, h:HEAD_H,
    fill:{color:period.bg}, line:{color:period.bg},
  });
  // Period # + zh name (left)
  s.addText(`P${period.n}  ${period.zh}`, {
    x:bx+0.10, y:by+0.03, w:BLOCK_W*0.60, h:0.22,
    fontSize:11, bold:true, color:period.dk, fontFace:'Georgia', valign:'middle', margin:0,
  });
  // en name (left, below)
  s.addText(period.en, {
    x:bx+0.10, y:by+0.24, w:BLOCK_W*0.60, h:0.18,
    fontSize:7.5, italic:true, color:period.dk, fontFace:'Georgia', valign:'middle', margin:0,
  });
  // chapters + pages (right-aligned)
  s.addText(period.chs, {
    x:bx+BLOCK_W-2.1, y:by+0.04, w:2.0, h:0.18,
    fontSize:8.5, bold:true, color:period.dk, fontFace:'Calibri', align:'right', valign:'middle', margin:0,
  });
  s.addText(period.pages, {
    x:bx+BLOCK_W-2.1, y:by+0.23, w:2.0, h:0.18,
    fontSize:8, color:period.dk, fontFace:'Calibri', align:'right', valign:'middle', margin:0,
  });

  // Content area
  const QA_TOP = by + HEAD_H + 0.04;
  const QA_H_AREA = BLOCK_H - HEAD_H - 0.08;
  const COL1_N = Math.ceil(period.qa.length / 2);
  const COL2_N = period.qa.length - COL1_N;
  const SUB_GAP = 0.06;
  const SUB_PAD = 0.08;
  const SUB_W = (BLOCK_W - 2*SUB_PAD - SUB_GAP) / 2;
  const H1 = QA_H_AREA / COL1_N;
  const H2 = QA_H_AREA / COL2_N;

  // Vertical divider between sub-columns
  s.addShape(pres.ShapeType.line, {
    x:bx + SUB_PAD + SUB_W + SUB_GAP/2, y:QA_TOP,
    w:0, h:QA_H_AREA,
    line:{color:G.paleLine, width:0.5},
  });

  period.qa.forEach((qa, idx) => {
    const inCol1 = idx < COL1_N;
    const j = inCol1 ? idx : idx - COL1_N;
    const x = bx + SUB_PAD + (inCol1 ? 0 : SUB_W + SUB_GAP);
    const y = QA_TOP + j * (inCol1 ? H1 : H2);
    const h = (inCol1 ? H1 : H2) - 0.02;
    addQASlot(x, y, SUB_W, h, idx+1, qa);
  });
}

function addQASlot(x, y, w, h, num, qa) {
  s.addText([
    { text:`${num}. `, options:{ bold:true, color:G.accent, fontSize:6.5 } },
    { text:qa.q,       options:{ bold:true, color:G.dark,   fontSize:6.5 } },
    { text:`\n${qa.a}`, options:{            color:G.mid,    fontSize:6 } },
    { text:`  [${qa.p}]`, options:{ italic:true, color:G.accent, fontSize:5.5 } },
  ], {
    x, y, w, h,
    fontFace:'Calibri', valign:'top', margin:0,
    paraSpaceAfter:0, lineSpacingMultiple:1.05,
  });
}

PERIODS.forEach((p, i) => {
  const col = i % 3, row = Math.floor(i / 3);
  addPeriodBlock(p, COL_X[col], ROW_Y[row]);
});

// ── Footer ───────────────────────────────────────────────────────────────────
s.addShape(pres.ShapeType.rect, { x:0, y:PAGE_H-0.08, w:PAGE_W, h:0.08, fill:{color:G.gold}, line:{color:G.gold} });

// ── Save ─────────────────────────────────────────────────────────────────────
pres.writeFile({ fileName:'QA_100.pptx' }).then(f => {
  console.log('Wrote', f);
  console.log(`Total Q&A: ${totalQA}`);
  PERIODS.forEach(p => console.log(`  P${p.n} ${p.zh.padEnd(14,' ')} ${p.chs.padEnd(8)} ${p.pages.padEnd(14)} ${p.qa.length} Q&A`));
});
