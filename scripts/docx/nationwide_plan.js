// 全国ダム地質分類_作業計画書（改訂版）生成スクリプト
// fullplan_p1.jsのユーティリティを使用

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ─── ユーティリティ ───────────────────────────────────────────────
const A4W = 11906, A4H = 16838, MAR = 1080, CONTENT_W = A4W - MAR * 2;

const COLORS = {
  h1bg: '1F4E79', h1fg: 'FFFFFF',
  h2bg: '2E75B6', h2fg: 'FFFFFF',
  h3bg: 'BDD7EE', h3fg: '1F4E79',
  h4bg: 'DEEAF1', h4fg: '1F4E79',
  tabhead: '2E75B6', tabheadfg: 'FFFFFF',
  row1: 'FFFFFF', row2: 'EBF3FB',
  accent1: 'C00000', accent2: '375623', accent3: '833C00',
  border: '2E75B6', lightborder: 'BDD7EE',
  note: 'FFF2CC', noteborder: 'F4B942',
  good: 'E2EFDA', warn: 'FFF2CC', risk: 'FCE4D6',
};

function HR() {
  return new Paragraph({
    paragraph: { spacing: { before: 60, after: 60 } },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: COLORS.border, space: 1 } },
    children: []
  });
}

function H1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, color: COLORS.h1fg, size: 32, font: 'Arial' })],
    shading: { fill: COLORS.h1bg, type: ShadingType.CLEAR },
    spacing: { before: 360, after: 180 },
    indent: { left: 180 },
  });
}

function H2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, color: COLORS.h2fg, size: 28, font: 'Arial' })],
    shading: { fill: COLORS.h2bg, type: ShadingType.CLEAR },
    spacing: { before: 280, after: 140 },
    indent: { left: 120 },
  });
}

function H3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, bold: true, color: COLORS.h3fg, size: 26, font: 'Arial' })],
    shading: { fill: COLORS.h3bg, type: ShadingType.CLEAR },
    spacing: { before: 200, after: 100 },
    indent: { left: 80 },
  });
}

function H4(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_4,
    children: [new TextRun({ text, bold: true, color: COLORS.h4fg, size: 24, font: 'Arial' })],
    shading: { fill: COLORS.h4bg, type: ShadingType.CLEAR },
    spacing: { before: 160, after: 80 },
    indent: { left: 60 },
  });
}

function P(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, font: 'MS Mincho', color: opts.color || '000000', bold: opts.bold || false, italics: opts.italic || false })],
    spacing: { before: opts.before || 80, after: opts.after || 80 },
    indent: opts.indent ? { left: opts.indent } : undefined,
    alignment: opts.align || AlignmentType.JUSTIFIED,
  });
}

function BUL(text, level = 0, color = null) {
  return new Paragraph({
    numbering: { reference: `bul${level}`, level: 0 },
    children: [new TextRun({ text, size: 22, font: 'MS Mincho', color: color || '000000' })],
    spacing: { before: 60, after: 60 },
  });
}

function NUM(text, level = 0) {
  return new Paragraph({
    numbering: { reference: `num${level}`, level: 0 },
    children: [new TextRun({ text, size: 22, font: 'MS Mincho' })],
    spacing: { before: 60, after: 60 },
  });
}

function PB() { return new Paragraph({ children: [new PageBreak()] }); }
function SP(n = 1) { return new Paragraph({ children: [new TextRun({ text: '' })], spacing: { before: 60 * n, after: 0 } }); }

const border = (c = COLORS.border) => ({ style: BorderStyle.SINGLE, size: 4, color: c });
const borders = (c) => ({ top: border(c), bottom: border(c), left: border(c), right: border(c) });

function mkCell(text, opts = {}) {
  const w = opts.w || Math.floor(CONTENT_W / (opts.cols || 4));
  const shading = opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined;
  const txtRun = new TextRun({
    text: String(text),
    bold: opts.bold || false,
    color: opts.color || '000000',
    size: opts.size || 20,
    font: opts.font || 'MS Mincho',
  });
  return new TableCell({
    borders: borders(opts.borderColor || COLORS.lightborder),
    width: { size: w, type: WidthType.DXA },
    shading,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    columnSpan: opts.span || 1,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [txtRun],
    })],
  });
}

function mkHeadRow(cells, cols) {
  return new TableRow({
    tableHeader: true,
    children: cells.map((c, i) =>
      mkCell(c.text || c, {
        fill: COLORS.tabhead, color: COLORS.tabheadfg,
        bold: true, size: 20, w: c.w, cols, borderColor: COLORS.border,
        align: AlignmentType.CENTER,
      })
    ),
  });
}

function mkDataRow(cells, even, cols, opts = {}) {
  return new TableRow({
    children: cells.map((c, i) => {
      const val = typeof c === 'object' ? c : { text: c };
      return mkCell(val.text || '', {
        fill: val.fill || (even ? COLORS.row2 : COLORS.row1),
        bold: val.bold || false,
        color: val.color || '000000',
        w: val.w, cols,
        align: val.align || AlignmentType.LEFT,
      });
    }),
  });
}

function Table2(headers, rows, widths = null) {
  const cols = headers.length;
  const totalW = CONTENT_W;
  const defW = Math.floor(totalW / cols);
  const hCells = headers.map((h, i) => ({
    text: typeof h === 'string' ? h : h.text,
    w: widths ? widths[i] : defW,
  }));
  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: widths || Array(cols).fill(defW),
    rows: [
      mkHeadRow(hCells, cols),
      ...rows.map((r, ri) => mkDataRow(r.map((c, ci) => ({
        text: typeof c === 'object' ? c.text : String(c),
        fill: typeof c === 'object' ? c.fill : (ri % 2 === 1 ? COLORS.row2 : COLORS.row1),
        bold: typeof c === 'object' ? c.bold : false,
        color: typeof c === 'object' ? c.color : '000000',
        w: widths ? widths[ci] : defW,
        align: typeof c === 'object' ? c.align : AlignmentType.LEFT,
      })), ri % 2 === 1, cols)),
    ],
  });
}

function NoteBox(title, lines) {
  const children = [];
  if (title) children.push(new Paragraph({
    children: [new TextRun({ text: '【' + title + '】', bold: true, size: 22, color: COLORS.accent3, font: 'MS Mincho' })],
    spacing: { before: 80, after: 60 },
  }));
  lines.forEach(l => children.push(new Paragraph({
    children: [new TextRun({ text: l, size: 21, font: 'MS Mincho' })],
    spacing: { before: 40, after: 40 },
    indent: { left: 200 },
  })));
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({
      children: [new TableCell({
        borders: { top: border(COLORS.noteborder), bottom: border(COLORS.noteborder), left: { style: BorderStyle.THICK, size: 12, color: COLORS.noteborder }, right: border(COLORS.noteborder) },
        width: { size: CONTENT_W, type: WidthType.DXA },
        shading: { fill: COLORS.note, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 200, right: 200 },
        children,
      })]
    })]
  });
}

// ─── 本文生成 ─────────────────────────────────────────────────────
const numbering = {
  config: [
    { reference: 'bul0', levels: [{ level: 0, format: LevelFormat.BULLET, text: '●', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 280 } } } }] },
    { reference: 'bul1', levels: [{ level: 0, format: LevelFormat.BULLET, text: '○', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 280 } } } }] },
    { reference: 'num0', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 560, hanging: 320 } } } }] },
  ]
};

const children = [];

// ══════════════════════════════════════════════════════════════════
// 表紙
// ══════════════════════════════════════════════════════════════════
children.push(SP(8));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: '日本全国ダム地質分類', bold: true, size: 56, font: 'Arial', color: COLORS.h1bg })],
  spacing: { before: 0, after: 200 },
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: '体系的分析・全国展開', bold: true, size: 40, font: 'Arial', color: COLORS.h2bg })],
  spacing: { before: 0, after: 200 },
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: '作業計画書（改訂版）', bold: false, size: 32, font: 'MS Mincho', color: '444444' })],
  spacing: { before: 0, after: 400 },
}));
children.push(HR());
children.push(SP(3));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: '2026年3月', size: 26, font: 'MS Mincho', color: '666666' })],
}));
children.push(SP(2));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: 'propylite.work', size: 24, font: 'Arial', color: '888888' })],
}));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// 0. 改訂にあたって
// ══════════════════════════════════════════════════════════════════
children.push(H1('0. 改訂の経緯と基本方針'));
children.push(P('本計画書は、北海道188基のダム地質分類作業として策定した「北海道ダム地質分類_全体計画書」を全国展開に向けて全面改訂したものである。改訂の主要な方針は以下のとおりである。'));
children.push(SP());
children.push(BUL('分析方法の標準化を優先する：北海道で構築した分類記号体系（Ⅰ/Ⅱ類・S/R/W/Gコード）を全国標準として確立した上で、全国展開を行う。'));
children.push(BUL('GeoNAVI Web APIを全国分析の基盤とする：産総研「20万分の1日本シームレス地質図V2」のWeb APIを活用し、ダム位置座標から地質情報を自動取得する手法を全国に展開する。'));
children.push(BUL('北海道の分析は2段階で行う：まずGeoNAVI情報のみによる基礎分析（全国共通手法）を確立し、その後にWordファイル・propylite.work情報による深化修正を重ねる。'));
children.push(BUL('全国ダム位置情報の取得コストを最小化する：国土数値情報ダムデータ（国土交通省）をベースとし、ダム便覧・各管理者公開情報で補完する。'));
children.push(SP());

children.push(NoteBox('本計画書の位置づけ', [
  '本書は全11フェーズの作業計画を定義するマスタードキュメントである。',
  '北海道分析（Ph.3〜Ph.6）は先行パイロットとして、全国分析（Ph.8〜Ph.11）の手法確立と並行して進める。',
  '「人的情報」とは、ダム管理者・設計施工関係者等から提供される設計・施工報告書・ルジオン試験記録等を指す。',
]));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// 1. 目的
// ══════════════════════════════════════════════════════════════════
children.push(H1('1. プロジェクトの目的'));
children.push(H3('1.1 最終目標'));
children.push(P('日本全国のダムを対象に、岩盤の地質生成過程・工学的強度・水理的特性を体系的に分類し、以下を実現する。'));
children.push(BUL('ダム基礎岩盤の地質的多様性を記号体系として可視化し、地域間・岩種間の比較を可能にする。'));
children.push(BUL('ダム建設未経験技術者が地質リスクを直観的に把握できる知識基盤を構築する。'));
children.push(BUL('老朽化ダムの健全性評価・補修優先度判定に資する地質的根拠を提供する。'));
children.push(BUL('国際的な岩盤分類体系（DMR・RMR・Q-system）との対応関係を整理し、日本固有の地質条件（火山岩・付加体・第四紀堆積物）を加味した独自体系を確立する。'));

children.push(H3('1.2 分類記号体系（確定版）'));
children.push(P('本プロジェクトで使用する分類記号体系を以下に示す。全国展開においてもこの体系を適用する。'));
children.push(SP());
children.push(Table2(
  ['ブロック', '記号', '段階数', '内容', 'フェーズ'],
  [
    ['① 時代大区分', 'Ⅰ・Ⅱ', '2', 'Ⅰ＝古期（先新第三紀）、Ⅱ＝新期（新第三紀以降）', 'Ph.3〜'],
    [{ text: '② 岩石種サブ', fill: COLORS.row1 }, 'a〜e', '各5', '各時代の代表岩石・地質帯', 'Ph.3〜'],
    ['③ 強度コード', 'S1〜S5・S?', '5+欠如', 'ISRM一軸圧縮強度基準', 'Ph.3〜'],
    [{ text: '④ リスク指標', fill: COLORS.row1 }, 'R1〜R6', '6', '断層・変質・膨張・冷却亀裂・溶解・未固結', 'Ph.3〜'],
    ['⑤ 透水性', 'W1〜W5・W?', '5+欠如', 'ルジオン値（Lu）基準', 'Ph.5以降'],
    [{ text: '⑥ 基礎処理難易度', fill: COLORS.row1 }, 'G1〜G4・G?', '4+欠如', 'グラウチング難易度', 'Ph.5以降'],
  ],
  [2400, 1600, 900, 3800, 1400]
));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// 2. 全体構成
// ══════════════════════════════════════════════════════════════════
children.push(H1('2. 作業フェーズ全体構成'));
children.push(P('本計画は全11フェーズ（Ph.0〜Ph.10）で構成される。北海道先行パイロット（Ph.3〜Ph.6）と全国展開（Ph.7〜Ph.10）の2トラックで進行する。'));
children.push(SP());
children.push(Table2(
  ['Ph.', 'フェーズ名', '期間目安', '対象', '主な成果物'],
  [
    [{ text: 'Ph.0', bold: true, fill: COLORS.h3bg }, '全国分析手法の確立', '〜1ヶ月', '設計作業', '全国共通分析仕様書・GeoNAVI API利用手順'],
    ['Ph.1', 'ダム位置情報の整備', '〜1ヶ月', '全国', '全国ダムリスト（緯度経度付き）'],
    [{ text: 'Ph.2', fill: COLORS.row1 }, '地質区分記号体系の確定', '〜1ヶ月', '設計作業', '分類体系確定版・国際基準対応表'],
    ['Ph.3', '北海道：GeoNAVI基礎分析', '〜1ヶ月', '北海道188基', '北海道DB（GeoNAVIベース）'],
    [{ text: 'Ph.4', fill: COLORS.row1 }, '北海道：深化修正（一次）', '〜2ヶ月', '北海道188基', 'Wordファイル・propylite.work情報追加'],
    ['Ph.5', '北海道：深化修正（二次）', '継続的', '北海道188基', '人的情報（設計・施工報告書等）による修正'],
    [{ text: 'Ph.6', fill: COLORS.row1 }, '北海道：結果考察', '〜1ヶ月', '北海道188基', '北海道ダム地質考察レポート'],
    ['Ph.7', '全国：GeoNAVI自動分析', '〜3ヶ月', '全国選定ダム', '全国DB（GeoNAVIベース）'],
    [{ text: 'Ph.8', fill: COLORS.row1 }, '全国：追加情報の整備', '継続的', '全国選定ダム', '管理者提供情報・設計資料の収集・統合'],
    ['Ph.9', '全国：人的情報修正', '継続的', '全国選定ダム', '信頼度向上版DB'],
    [{ text: 'Ph.10', fill: COLORS.row1 }, '全国：最終考察・成果物', '〜2ヶ月', '全国', '最終報告書・オープンデータ・GIS'],
  ],
  [600, 2800, 1200, 1800, 3700]
));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// Ph.0
// ══════════════════════════════════════════════════════════════════
children.push(H1('Ph.0　全国分析手法の確立'));
children.push(P('全国展開に先立ち、GeoNAVI Web APIを用いた地質情報自動取得の手法を確立し、北海道で検証する。'));

children.push(H3('Ph.0-1　産総研シームレス地質図 Web APIの活用'));
children.push(P('産総研「20万分の1日本シームレス地質図V2」は、Web API（ver.1.3.1）を公開しており、緯度・経度を指定して地質情報をJSON形式で取得できる。本プロジェクトではこのAPIを全国分析の地質情報基盤として採用する。'));
children.push(SP());
children.push(NoteBox('GeoNAVI Web API 仕様', [
  'エンドポイント：https://gbank.gsj.jp/seamless/v2/api/1.2/',
  '凡例取得：GET /legend?lang=ja → 全地質区分コードと岩相名称をJSON取得',
  '地質情報取得：GET /query?lat={緯度}&lng={経度}&datum=WGS84 → 指定点の地質区分・岩相・時代をJSON取得',
  'typeパラメータ：level4（簡略版・凡例数約400）、level8（詳細版）が指定可能',
  '利用制限：特になし（オープンデータ・CC BY 4.0相当）',
]));
children.push(SP());

children.push(H3('Ph.0-2　GeoNAVI取得情報から分類記号への変換ルール'));
children.push(P('GeoNAVI返却値（岩相名・地質時代）を本プロジェクトの分類記号（Ⅰ/Ⅱ・a〜e・S・R）に変換する対応表を作成する。この変換ルールが全国共通手法の核心部分であり、北海道での検証で精度を確認する。'));
children.push(SP());
children.push(Table2(
  ['GeoNAVI岩相分類', '時代大区分', '岩石種サブ', 'S強度目安', 'R主リスク'],
  [
    ['花崗岩・花崗閃緑岩・閃緑岩', { text: 'Ⅰ', bold: true }, 'a（深成岩）', 'S4〜S5', 'R1（断層・節理）'],
    ['緑色岩・輝緑凝灰岩・枕状溶岩', { text: 'Ⅰ', bold: true }, 'b（緑色岩）', 'S3〜S4', 'R3（変質）'],
    ['砂岩・泥岩・頁岩（中生代以前）', { text: 'Ⅰ', bold: true }, 'c（堆積岩）', 'S2〜S4', 'R1・R2'],
    ['変成岩・片麻岩・結晶片岩', { text: 'Ⅰ', bold: true }, 'd（変成岩）', 'S3〜S5', 'R1'],
    ['蛇紋岩・かんらん岩', { text: 'Ⅰ', bold: true }, 'e（超苦鉄質）', 'S1〜S5', 'R3（異質）'],
    ['安山岩・玄武岩溶岩（新第三紀〜）', { text: 'Ⅱ', bold: true }, 'a（火山岩）', 'S2〜S4', 'R1・R3'],
    ['溶結凝灰岩（第四紀）', { text: 'Ⅱ', bold: true }, 'b（溶結凝灰岩）', 'S3〜S4', 'R4（冷却亀裂）'],
    ['砂岩・泥岩（新第三紀〜）', { text: 'Ⅱ', bold: true }, 'c（新期堆積岩）', 'S2〜S3', 'R1・R2'],
    ['礫岩・砂岩（鮮新世〜更新世）', { text: 'Ⅱ', bold: true }, 'd（半固結堆積岩）', 'S1〜S3', 'R1・R2'],
    ['沖積層・段丘礫層・火砕流堆積物', { text: 'Ⅱ', bold: true }, 'e（未固結）', 'S1〜S2', 'R6（未固結）'],
  ],
  [2800, 1200, 1600, 1200, 2300]
));
children.push(SP());

children.push(H3('Ph.0-3　信頼度体系（全国共通）'));
children.push(P('全国共通の信頼度（A〜D）を以下のとおり定義する。北海道の3段階（A・B・C）を4段階に拡張する。'));
children.push(SP());
children.push(Table2(
  ['信頼度', '定義', '情報源', '目標比率'],
  [
    [{ text: 'A', bold: true, color: COLORS.accent2 }, '透水性・基礎処理実績データ確認済み', '設計・施工報告書・ルジオン試験記録', '目標10%以上'],
    [{ text: 'B', bold: true, color: COLORS.h2bg }, 'ダム固有の地質記述確認済み', '工事誌・専門文献・propylite.work等', '目標30%以上'],
    [{ text: 'C', bold: true, fill: COLORS.row1 }, 'GeoNAVI地質図＋ダム型式から推定', 'GeoNAVI API・国土数値情報', '初期分析時の主体'],
    [{ text: 'D', bold: true, color: COLORS.accent1 }, '位置情報のみで地質情報取得不可', '（データ不足）', '最小化目標'],
  ],
  [800, 3200, 3000, 1100]
));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// Ph.1
// ══════════════════════════════════════════════════════════════════
children.push(H1('Ph.1　全国ダム位置情報の整備'));
children.push(P('全国分析の前提として、各ダムの緯度・経度情報を含むマスターリストを整備する。これがGeoNAVI APIを呼び出す際の入力データとなる。'));

children.push(H3('Ph.1-1　利用可能なデータソース'));
children.push(SP());
children.push(Table2(
  ['データソース', '収録数目安', '位置情報', '取得方法', '評価'],
  [
    [{ text: '国土数値情報「ダムデータ」\n（国土交通省）', fill: COLORS.good }, '約2,700基', '緯度経度あり（GML/SHP）', '無料ダウンロード', { text: '◎最優先', bold: true, color: COLORS.accent2 }],
    ['ダム便覧（日本ダム協会）', '約3,000基', '記載なし（住所のみ）', 'Webスクレイピング', '○補完用'],
    [{ text: 'DamMaps（dammaps.jp）', fill: COLORS.row1 }, '約2,500基', '地図表示あり', 'Webスクレイピング', '○補完用'],
    ['国土交通省水資源ダム一覧', '国直轄のみ', '一部あり', '公開ページ参照', '△限定的'],
    [{ text: 'Googleマップ・地理院地図', fill: COLORS.row1 }, '個別検索', '高精度', '手動補完', '△労力大'],
  ],
  [2500, 1200, 1500, 1800, 1100]
));
children.push(SP());

children.push(H3('Ph.1-2　位置情報整備の手順'));
children.push(NUM('国土数値情報「ダムデータ」（W01）をダウンロード（GMLまたはShapeファイル形式）。'));
children.push(NUM('GMLまたはShapeファイルから緯度・経度・ダム名・河川名・都道府県・目的・型式を抽出しCSV化。'));
children.push(NUM('ダム便覧・DamMapsのデータと名称照合し、管理者区分・堤高・完成年等を補完。'));
children.push(NUM('国土数値情報に収録されていない農業ダム（農林水産省管轄）は農業水産部公開資料で補完。'));
children.push(NUM('全国マスターCSVを作成：「ダムID・ダム名・河川名・都道府県・緯度・経度・管理者区分・型式・堤高・完成年」の10列。'));
children.push(SP());
children.push(NoteBox('私（Claude）が実施できる範囲', [
  '国土数値情報GMLファイルをアップロードいただければ、Python/Pandasで自動変換・CSV作成が可能。',
  'ダム便覧Webページから名称・河川名・目的等のテキスト情報の収集補助が可能。',
  'ダム名→緯度経度のジオコーディング（Nominatim/Google Geocoding API経由）を半自動化するスクリプト作成が可能。',
]));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// Ph.2
// ══════════════════════════════════════════════════════════════════
children.push(H1('Ph.2　地質区分記号体系の確定'));
children.push(P('Ph.0で設計した変換ルールを北海道データで検証し、全国適用版として確定する。この段階で体系を固定し、以降のフェーズで変更しない。'));

children.push(H3('Ph.2-1　確定作業の内容'));
children.push(BUL('GeoNAVI APIから取得した岩相コード（約400種）すべてについてⅠ/Ⅱ・サブ記号・Sコードへの変換規則を完成させる。'));
children.push(BUL('DMR（Romana 2003）・RMR89・Q-systemとの対応表を整備し、国際的な互換性を確保する。'));
children.push(BUL('日本固有の地質条件（付加体・溶結凝灰岩・蛇紋岩・泥炭性堆積物）への特例規則を定義する。'));
children.push(BUL('W・Gコードの基準（ルジオン値・グラウチング量）を全国適用可能な形で最終化する。'));

children.push(H3('Ph.2-2　成果物'));
children.push(BUL('「GeoNAVI岩相コード→分類記号変換表」（Excel・約400行）'));
children.push(BUL('「国際基準対応表」（DMR/RMR/Q-system vs 本体系）'));
children.push(BUL('「分類記号体系確定版仕様書」（docx）'));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// Ph.3
// ══════════════════════════════════════════════════════════════════
children.push(H1('Ph.3　北海道：GeoNAVI基礎分析'));
children.push(P('北海道188基のダムについて、GeoNAVI APIのみを情報源として地質区分を行う。これは「全国共通手法のみによる分析結果」であり、Ph.4〜Ph.5の深化修正前の基準ベースラインとなる。'));
children.push(SP());

children.push(NoteBox('Ph.3の意義', [
  '現在のDB（北海道ダム地質分類DB.xlsx）はWordファイル・propylite.work情報を既に統合した状態にある。',
  'Ph.3ではこれをGeoNAVI情報のみの状態に「戻す」のではなく、GeoNAVIのみを使った場合の結果を独立したシートとして並列記録する。',
  'これにより「GeoNAVI基礎分析 → 文献修正 → 人的情報修正」という3段階の精度向上プロセスが可視化できる。',
  '全国分析ではGeoNAVI基礎分析のみの状態が出発点となるため、このプロセスの透明化は重要な意味を持つ。',
]));

children.push(SP());
children.push(H3('Ph.3-1　実施手順'));
children.push(NUM('北海道188基の緯度・経度リスト（既存KMZから抽出）を用意する。'));
children.push(NUM('GeoNAVI API（v1.2）に各ダムの座標を送信し、岩相コード・岩相名・地質時代を取得するPythonスクリプトを実行。'));
children.push(NUM('取得した岩相情報をPh.2の変換表に照らし合わせ、Ⅰ/Ⅱ・サブ・S・Rコードを自動付与。'));
children.push(NUM('結果をExcelの「GeoNAVI基礎分析」シートとして保存。信頼度はすべてC（初期値）。'));
children.push(NUM('既存DB（Wordおよびpropylite.work情報を含む）と照合し、乖離箇所をリストアップ。'));

children.push(H3('Ph.3-2　成果物'));
children.push(BUL('北海道ダム地質分類DB（改訂版）：「GeoNAVI基礎分析」シート追加'));
children.push(BUL('GeoNAVI取得結果の乖離分析レポート（乖離箇所と要因の考察）'));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// Ph.4
// ══════════════════════════════════════════════════════════════════
children.push(H1('Ph.4　北海道：深化修正（一次）— 文献・Webサイト情報'));
children.push(P('Ph.3のGeoNAVI基礎分析に対し、既存のWordファイル（国営農業用ダム地質雑感 道東・道北・道南編）およびpropylite.workの情報を追加・修正する。これが現在の「北海道ダム地質分類DB」の主要情報層に相当する。'));

children.push(H3('Ph.4-1　使用資料'));
children.push(Table2(
  ['資料名', '収録ダム数', '情報の質', '主な追加情報'],
  [
    [{ text: 'propylite.work\n（道央・道南・道北・道東編）', fill: COLORS.good }, '約60〜80基', '高', '岩種・構造・変質・透水性・施工記録'],
    ['国営農業用ダム地質雑感（道東編）', '7基', '非常に高', '詳細地質・ルジオン値・グラウト計画'],
    [{ text: '国営農業用ダム地質雑感（道北編）', fill: COLORS.row1 }, '14基', '非常に高', '詳細地質・変形係数・地すべり情報'],
    ['国営農業用ダム地質雑感（道南編）', '5基', '高', '詳細地質・透水係数・水理地質'],
  ],
  [3000, 1200, 900, 4000]
));

children.push(H3('Ph.4-2　修正内容'));
children.push(BUL('GeoNAVI基礎分析の地質区分コードを文献記述に基づき修正（例：Ⅱ-d→Ⅱ-b等）。'));
children.push(BUL('透水性コード（W）・基礎処理難易度コード（G）を文献記述から付与。'));
children.push(BUL('信頼度をB（文献確認済み）に更新。'));
children.push(BUL('リスク指標（R）を具体的な記述から精査・追加（スレーキング・熱水変質・断層等）。'));

children.push(H3('Ph.4-3　成果物（現状相当）'));
children.push(BUL('北海道ダム地質分類DB.xlsx（信頼度A=5基・B=80基・C=103基） ← 現在完成済み'));
children.push(BUL('北海道ダム地質区分.kmz（色分けマーカー付き） ← 現在完成済み'));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// Ph.5
// ══════════════════════════════════════════════════════════════════
children.push(H1('Ph.5　北海道：深化修正（二次）— 人的情報'));
children.push(P('Ph.4のDBに対し、ダム管理者・設計施工関係者等から提供される一次資料（設計報告書・施工記録・ルジオン試験記録等）を追加し、信頼度A・Bの割合を最大化する。'));

children.push(H3('Ph.5-1　収集すべき人的情報の種類'));
children.push(BUL('ダム設計報告書（地質調査編）：地質縦断図・横断図・ボーリング柱状図・室内試験結果'));
children.push(BUL('施工記録（グラウチング）：注入孔配置・ルジオン試験値・注入量・セメント/水比の記録'));
children.push(BUL('ダム完成後検査記録：漏水量観測・基礎ひずみ・揚圧力測定データ'));
children.push(BUL('老朽化点検記録：基礎岩盤変状・変質進行・溶解空洞の有無'));

children.push(H3('Ph.5-2　優先収集対象（北海道）'));
children.push(P('信頼度Cのまま残っている103基のうち、以下の条件に合うダムを優先する。'));
children.push(BUL('Ⅱ-d（半固結堆積岩）またはⅡ-e（未固結）に分類されたダム：基礎処理情報が工学的に重要。'));
children.push(BUL('堤高50m以上のコンクリートダム：設計段階での地質情報が詳細に記録されている可能性が高い。'));
children.push(BUL('老朽化対策が進行中のダム：再調査記録が新たに存在する可能性がある。'));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// Ph.6
// ══════════════════════════════════════════════════════════════════
children.push(H1('Ph.6　北海道：結果考察'));
children.push(P('Ph.3〜Ph.5で構築した北海道ダム地質分類DBを多角的に考察する。以下に考察の主要テーマとその意義を示す。'));

children.push(H3('考察①　地質帯別分布と工学的特性の比較'));
children.push(BUL('北海道を地質帯（日高帯・蝦夷帯・天北地向斜・大雪火山帯・支笏-洞爺火山帯・道東白亜系等）別に集計し、各地質帯のS・R・W・Gコード分布を統計的に比較する。'));
children.push(BUL('「付加体系（Ⅰ-b・Ⅰ-c）のダムは透水性W4〜W5が多い傾向があるか」等の仮説を検証する。'));

children.push(H3('考察②　ダム型式と地質区分の相関'));
children.push(BUL('コンクリートダム（重力式・アーチ）に採用されやすい地質区分と、フィルダムに採用されやすい地質区分を分析する。'));
children.push(BUL('Ⅱ-e（未固結）地盤に建設されたダムの特殊基礎処理技術（地下連続壁・ブランケット等）の分布を可視化する。'));

children.push(H3('考察③　建設年代と地質知識の深化'));
children.push(BUL('建設年代（1950年代〜2010年代）別に地質区分記号の分布を分析し、時代とともにより困難な地盤条件への挑戦が進んだかを検証する。'));
children.push(BUL('知識段階コード（K1〜K5）の導入による「当時の技術水準と現在の知識ギャップ」の可視化を提案する。'));

children.push(H3('考察④　老朽化リスクポテンシャル'));
children.push(BUL('Rコード（リスク指標）×建設年代×ダム型式のマトリクス分析により、優先的に再調査が必要なダムを特定するリスクランキングを作成する。'));
children.push(BUL('変質系リスク（R3：熱水変質・R2：スレーキング）と老朽化の相関を分析する。'));

children.push(H3('考察⑤　全国展開への示唆'));
children.push(BUL('北海道の地質多様性（付加体・火山岩・堆積岩・変成岩・蛇紋岩）が日本全国の地質区分の試験場として機能した点を評価する。'));
children.push(BUL('GeoNAVI基礎分析とWordファイル修正の乖離率から、全国での信頼度C→Bへの修正作業量を推計する。'));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// Ph.7
// ══════════════════════════════════════════════════════════════════
children.push(H1('Ph.7　全国：GeoNAVI自動分析 — 対象ダムの選定と分析方法'));

children.push(H3('Ph.7-1　全国ダムの規模と選定の必要性'));
children.push(P('日本全国には約3,000基のダムが存在する（国土数値情報ダムデータ 2014年版）。全基を対象とした分析は理想であるが、品質管理・考察の深度の観点から選定を行う。'));
children.push(SP());
children.push(Table2(
  ['管理者区分', '概数', '特徴'],
  [
    ['国土交通省（直轄・補助）', '約600基', '高堤高・多目的。設計資料が最も整備されている。'],
    ['農林水産省（国営・道府県営農業）', '約800基', '農業用フィルダム中心。地域に密着した地質条件。'],
    ['都道府県・市町村（治水・利水）', '約1,000基', '規模多様。資料整備状況に大きな差がある。'],
    ['電力（発電専用）', '約400基', '高山地帯・急峻地形が多い。堅硬岩盤に多い。'],
    ['水道事業体', '約200基', '都市近郊。情報公開度が比較的高い。'],
  ],
  [2500, 1000, 6600]
));

children.push(H3('Ph.7-2　選定基準の提案'));
children.push(P('以下の優先順位で選定を行い、まず約500〜800基の「第一選定群」を分析対象とすることを提案する。'));
children.push(SP());
children.push(Table2(
  ['選定基準', '優先度', '選定概数', '理由'],
  [
    [{ text: '堤高15m以上のコンクリートダム（全管理者）', fill: COLORS.good }, '★★★', '約400基', 'ダム便覧データが整備。地質調査記録が存在。'],
    ['堤高50m以上のフィルダム（全管理者）', '★★★', '約150基', '大規模基礎処理が行われている。'],
    [{ text: '老朽化対策事業中のダム（国交省公表）', fill: COLORS.row1 }, '★★★', '約100基', '再調査記録が新たに存在する可能性大。'],
    ['特殊地質ダム（石灰岩・蛇紋岩・未固結等）', '★★', '約50基', '工学的に高い学術的価値を持つ。'],
    [{ text: 'アーチダムおよびバットレスダム（全て）', fill: COLORS.row1 }, '★★', '約80基', '岩盤強度要件が最も厳しい型式。'],
    ['1960年代以前竣工のダム（全型式）', '★', '約100基', '老朽化と地質条件の複合リスク分析に重要。'],
  ],
  [3200, 800, 900, 5200]
));

children.push(H3('Ph.7-3　GeoNAVI自動取得の実施方法'));
children.push(BUL('全国マスターCSV（Ph.1で整備）の緯度・経度を使い、GeoNAVI APIをバッチ処理で呼び出す。'));
children.push(BUL('1基あたりのAPI呼び出し時間は数秒以内であり、500基で30分程度の自動処理が見込まれる。'));
children.push(BUL('返却値（岩相コード・岩相名・地質時代）をPh.2の変換表に照らして地質区分記号を自動付与。'));
children.push(BUL('全結果をExcelまたはSQLiteデータベースに格納し、GIS（QGIS）とも連携可能な形式で出力。'));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// Ph.8
// ══════════════════════════════════════════════════════════════════
children.push(H1('Ph.8　全国：追加情報の整備 — Claudeの実施可能範囲'));
children.push(P('Ph.7のGeoNAVI自動分析に対し、公開情報を活用して信頼度を向上させる。以下に私（Claude）が実施できる範囲を具体的に示す。'));

children.push(H3('Ph.8-1　Claudeが自動・半自動で実施できること'));
children.push(Table2(
  ['作業内容', '手法', '期待効果', '制約'],
  [
    [{ text: 'ダム便覧Webページの地質記述収集', fill: COLORS.good }, 'Webフェッチ＋テキスト解析', '信頼度C→B（約200〜300基）', 'ダム便覧の地質情報は記述量が少ない'],
    ['地質図幅説明書のテキスト検索', '産総研地質図カタログ参照', '周辺地質の詳細情報取得', '図幅単位のため個別ダムとの対応付けが必要'],
    [{ text: 'propylite.work相当サイトの検索', fill: COLORS.row1 }, 'Web検索＋収集', '各地方の類似Webリソース発見', '地方ごとに情報密度の差が大きい'],
    ['発電事業者の環境アセス等公開資料', '公開PDF解析', 'ルジオン値等の実測データ取得', '検索・取得の時間コストが大きい'],
    [{ text: '国会図書館デジタルコレクション', fill: COLORS.row1 }, '文献検索', '1960〜90年代施工記録の発見', '著作権・閲覧制限に注意'],
  ],
  [2800, 1800, 2400, 3100]
));

children.push(H3('Ph.8-2　人的情報でしか取得できないこと'));
children.push(BUL('未公開の設計・施工報告書（管理者保有）：ルジオン試験の全データ、グラウチング量。'));
children.push(BUL('老朽化点検記録（管理者保有）：漏水量・揚圧力・基礎変状の経年データ。'));
children.push(BUL('地方の専門技術者による口頭伝承：設計段階では記録されなかった地質上の問題・対応措置。'));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// Ph.9
// ══════════════════════════════════════════════════════════════════
children.push(H1('Ph.9　全国：人的情報による修正'));
children.push(P('Ph.8の公開情報収集に加え、ダム管理者・設計者・地質技術者等の人的ネットワークを通じて一次資料を収集し、信頼度の向上と記号体系の精度向上を図る。'));

children.push(H3('Ph.9-1　人的情報収集の優先ルート'));
children.push(BUL('各地方整備局・北海道開発局ダム管理所への情報提供依頼（公文書開示請求含む）。'));
children.push(BUL('農業水産省・道県農政局への国営農業ダム施工記録の照会。'));
children.push(BUL('電力事業者（北電・東電・関電等）の環境・技術広報資料の収集。'));
children.push(BUL('地方の建設コンサルタント・地質調査会社への協力要請。'));
children.push(BUL('ダム関連学会（日本大ダム会議・土木学会水工学委員会等）の論文・報告書収集。'));

children.push(H3('Ph.9-2　修正の優先順位'));
children.push(NUM('信頼度C（GeoNAVI推定のみ）のコンクリートダム（堤高30m以上） — 設計記録が存在する可能性が高い。'));
children.push(NUM('Ⅱ-e（未固結）に分類されたコンクリートダム — 特殊基礎処理が行われているはずであり、修正が最も重要。'));
children.push(NUM('リスクR2〜R3（スレーキング・変質）が付与されたダム — 老朽化の進行が予測され、情報収集の緊急性が高い。'));
children.push(NUM('1970年代以前竣工の高堤ダム — 老朽化対策の優先順位が高く、地質情報の更新が急務。'));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// Ph.10
// ══════════════════════════════════════════════════════════════════
children.push(H1('Ph.10　全国：最終考察と成果物取りまとめ'));

children.push(H3('考察①　日本の地質帯とダム基礎岩盤の全国分布'));
children.push(BUL('変動帯としての日本列島の地質多様性（付加体・島弧火山・内帯花崗岩・外帯変成岩）がダム基礎岩盤に与える影響を体系化する。'));
children.push(BUL('「地質帯別のリスク分布地図」を作成し、日本全国のダム建設適地・高リスク地域を可視化する。'));

children.push(H3('考察②　管理者区分・地域別の情報密度の格差'));
children.push(BUL('国交省直轄ダム（信頼度A・B比率が高い）と地方管理ダム（信頼度C比率が高い）の情報密度の格差を定量化し、優先的な情報整備が必要な地域・管理者を特定する。'));

children.push(H3('考察③　老朽化リスクの全国分布'));
children.push(BUL('Rコード（リスク指標）×建設年代×堤高の3次元マトリクスによるリスクランキングを全国ダムに適用し、補修・再調査の優先度を提示する。'));
children.push(BUL('特にR2（スレーキング系泥岩）・R3（熱水変質）・R5（石灰岩溶解）を持つ老朽ダムは最上位リスクとして分類する。'));

children.push(H3('考察④　全国展開の限界と今後の課題'));
children.push(BUL('GeoNAVI 1:20万スケールの限界：局所的な断層・岩相変化を捉えられない可能性。1:5万地質図幅との照合が望ましい事例の提示。'));
children.push(BUL('農業ダム・水道ダムの情報格差：これらの情報整備に向けた制度的提言。'));
children.push(BUL('次世代分類コード（D/E/C/V/K）の全国展開に向けた優先課題と追加データ収集計画。'));

children.push(H3('最終成果物'));
children.push(Table2(
  ['成果物名', '形式', '内容'],
  [
    [{ text: '全国ダム地質分類DB', fill: COLORS.good }, 'Excel / SQLite', '全選定ダムの分類記号・信頼度・判定根拠'],
    ['全国ダム地質区分KMZ', 'KMZ（Google Earth）', '地質区分別色分けマーカー・ポップアップ情報'],
    [{ text: '全国ダム地質分布GIS', fill: COLORS.row1 }, 'QGIS プロジェクト', 'シームレス地質図レイヤー＋ダムポイントデータ'],
    ['老朽化リスクランキング', 'Excel', 'Rコード×建設年代×堤高による優先度順位表'],
    [{ text: '地域別地質考察レポート', fill: COLORS.row1 }, 'docx', '地方ブロック別（7ブロック）の地質特性まとめ'],
    ['最終報告書', 'docx / PDF', '全フェーズの成果・考察・提言をまとめた総括文書'],
    [{ text: '教育用資料（ダム地質入門）', fill: COLORS.row1 }, 'docx / HTML', 'ダム建設未経験技術者向けの基礎知識資料'],
  ],
  [3000, 1800, 5300]
));
children.push(PB());

// ══════════════════════════════════════════════════════════════════
// 参照資料
// ══════════════════════════════════════════════════════════════════
children.push(H1('参照資料・データ収集戦略'));
children.push(H3('一次資料（無料・オープンデータ）'));
children.push(BUL('産総研 シームレス地質図V2 Web API（https://gbank.gsj.jp/seamless/v2/api/）：無料・CC BY'));
children.push(BUL('国土数値情報「ダムデータ W01」（国土交通省）：GML・SHPで全国ダム位置情報を無料配布'));
children.push(BUL('ダム便覧（日本ダム協会 http://damnet.or.jp/）：ダム名・型式・堤高・管理者情報'));
children.push(BUL('propylite.work：北海道ダム地質の一次情報（管理者照会要）'));
children.push(BUL('国営農業用ダム地質雑感（道東・道北・道南編）：25基の詳細地質情報（提供済み）'));
children.push(H3('国際基準・参考文献'));
children.push(BUL('Romana, M.（2003）DMR (Dam Mass Rating). ISRM 10th Congress — ダム基礎岩盤の最重要国際先行体系'));
children.push(BUL('Bieniawski, Z.T.（1989）Engineering Rock Mass Classifications — RMR89原典'));
children.push(BUL('ICOLD Bulletin 88 Rock Foundations for Dams・Bulletin 111 — 国際標準グラウチング指針'));
children.push(BUL('Fell et al.（2015）Geotechnical Engineering of Dams — フィルダム基礎標準テキスト'));
children.push(BUL('ICS International Chronostratigraphic Chart 2023 — 地質年代値基準'));

children.push(PB());

// ══════════════════════════════════════════════════════════════════
// フッター付きフォーマット最終化
// ══════════════════════════════════════════════════════════════════
const doc = new Document({
  numbering,
  styles: {
    default: { document: { run: { font: 'MS Mincho', size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, font: 'Arial' },
        paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: 'Arial' },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, font: 'Arial' },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
      { id: 'Heading4', name: 'Heading 4', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: 'Arial' },
        paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 3 } },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: A4W, height: A4H },
        margin: { top: MAR, right: MAR, bottom: MAR + 400, left: MAR },
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [new TextRun({ text: '日本全国ダム地質分類　体系的分析・全国展開　作業計画書（改訂版）', size: 18, color: '888888', font: 'MS Mincho' })],
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: COLORS.border } },
            spacing: { after: 100 },
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: 'propylite.work  ／  2026年3月改訂       ', size: 18, color: '888888', font: 'MS Mincho' }),
              new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '888888' }),
              new TextRun({ text: ' / ', size: 18, color: '888888' }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: '888888' }),
            ],
            alignment: AlignmentType.RIGHT,
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: COLORS.border } },
            spacing: { before: 100 },
          })
        ]
      })
    },
    children,
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/全国ダム地質分類_作業計画書_改訂版.docx', buf);
  console.log('DONE');
});
