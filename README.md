# A History of Western Music - Slide Generator

Programmatically generated lecture slides for **A History of Western Music, 10th Edition** using [pptxgenjs](https://github.com/gitbrent/PptxGenJS).

## Chapters

| # | Title | JS Source | PPTX |
|---|-------|-----------|------|
| 1 | Music in Antiquity | `ch01_music_in_antiquity.js` | `Ch01_Music_in_Antiquity.pptx` |
| 2 | The Christian Church | `ch02_christian_church.js` | `Ch02_Christian_Church.pptx` |
| 3 | Roman Liturgy | `ch03_roman_liturgy.js` | `Ch03_Roman_Liturgy.pptx` |
| 4 | Song and Dance | `ch04_song_dance.js` | `Ch04_Song_and_Dance.pptx` |
| 5 | Polyphony | `ch05_polyphony.js` | `Ch05_Polyphony.pptx` |
| 6 | The Fourteenth Century | `ch06_fourteenth_century.js` | `Ch06_Fourteenth_Century.pptx` |
| 7 | The Renaissance | `ch07_renaissance.js` | `Ch07_Renaissance.pptx` |
| 8 | England and Burgundy | `ch08_england_burgundy.js` | `Ch08_England_Burgundy.pptx` |
| 9 | Franco-Flemish Composers | `ch09_franco_flemish.js` | `Ch09_Franco_Flemish.pptx` |
| 10 | The Madrigal | `ch10_madrigal.js` | `Ch10_Madrigal.pptx` |
| 11 | The Reformation | `ch11_reformation.js` | `Ch11_Reformation.pptx` |
| 12 | Instrumental Music | `ch12_instrumental.js` | `Ch12_Instrumental.pptx` |
| 13 | New Styles in the 17th Century | `ch13_new_styles.js` | `Ch13_New_Styles.pptx` |
| 14 | Opera | `ch14_opera.js` | `Ch14_Opera.pptx` |
| 15 | Chamber and Church Music | `ch15_chamber_church.js` | `Ch15_Chamber_Church.pptx` |
| 16 | France, England, Spain, the New World, and Russia | `ch16_france_england.js` | `Ch16_France_England.pptx` |
| 17 | Italy and Germany in the Late 17th Century | `ch17_italy_germany.js` | `Ch17_Italy_Germany.pptx` |

## Features

- Bilingual content (English / Traditional Chinese)
- Each chapter includes key terms, NAWM listening examples with YouTube links, and further reading
- Custom color themes per chapter
- 16:9 widescreen layout

## Usage

```bash
npm install
node ch01_music_in_antiquity.js   # generates Ch01_Music_in_Antiquity.pptx
```

To generate all slides:

```bash
for f in ch*.js; do node "$f"; done
```

## Requirements

- Node.js
- pptxgenjs (installed via `npm install`)
