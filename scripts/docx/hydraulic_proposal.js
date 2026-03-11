const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, PageNumber, PageBreak, Footer
} = require('docx');
const fs = require('fs');

// ─── 罫線 ───────────────────────────────────────
const bSingle = (color, size=4) => ({ style: BorderStyle.SINGLE, size, color });
const bNone = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const bdsHeader = c => { const b=bSingle(c,6); return {top:b,bottom:b,left:b,right:b}; };
const bdsNormal = { top:bSingle("CCCCCC",2), bottom:bSingle("CCCCCC",2), left:bSingle("CCCCCC",2), right:bSingle("CCCCCC",2) };
const bdsNone   = { top:bNone, bottom:bNone, left:bNone, right:bNone };

// ─── セル生成 ────────────────────────────────────
function cell(text, width, {
  isHeader=false, fill=null, center=false, bold=false,
  color="000000", size=20, small=false, xs=false,
  wrap=true, span=1, vAlign=VerticalAlign.CENTER, noBorder=false
}={}) {
  const sz = xs ? 17 : (small ? 19 : (isHeader ? 20 : size));
  return new TableCell({
    columnSpan: span,
    width: { size: width, type: WidthType.DXA },
    verticalAlign: vAlign,
    borders: noBorder ? bdsNone : (isHeader ? bdsHeader("2E5090") : bdsNormal),
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    margins: { top:100, bottom:100, left:160, right:160 },
    children: [new Paragraph({
      alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
      spacing: { before:0, after:0 },
      children: [new TextRun({
        text, font:"游明朝", size:sz,
        bold: isHeader || bold,
        color: isHeader ? "FFFFFF" : color,
      })]
    })]
  });
}

function row(cells, isHeader=false, height=null) {
  return new TableRow({
    tableHeader: isHeader,
    height: height ? { value:height, rule:"exact" } : undefined,
    children: cells.map(([t,w,o={}]) => cell(t, w, { ...o, isHeader }))
  });
}

// 区切り行（セクションヘッダー）
function sectionRow(label, fill, span=9) {
  return new TableRow({ children: [new TableCell({
    columnSpan: span,
    borders: bdsHeader("1F3864"),
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top:80, bottom:80, left:200, right:160 },
    children: [new Paragraph({ children: [
      new TextRun({ text:label, font:"游明朝", size:20, bold:true, color:"FFFFFF" })
    ]})]
  })]});
}

// ─── テキスト要素 ────────────────────────────────
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before:480, after:220 },
    border: { bottom: { style:BorderStyle.SINGLE, size:8, color:"1F3864", space:1 } },
    children: [new TextRun({ text, font:"游明朝", size:36, bold:true, color:"1F3864" })]
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before:320, after:140 },
    children: [new TextRun({ text, font:"游明朝", size:28, bold:true, color:"2E5090" })]
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before:220, after:100 },
    children: [new TextRun({ text, font:"游明朝", size:24, bold:true, color:"4472C4" })]
  });
}
function para(text, opts={}) {
  return new Paragraph({
    spacing: { before:80, after:80, line:380 },
    indent: { firstLine: opts.noIndent ? 0 : 440 },
    children: [new TextRun({ text, font:"游明朝", size:22, ...opts })]
  });
}
function note(text) {
  return new Paragraph({
    spacing: { before:60, after:60, line:340 },
    indent: { left:400 },
    children: [
      new TextRun({ text:"※ ", font:"游明朝", size:18, color:"888888" }),
      new TextRun({ text, font:"游明朝", size:18, color:"555555" })
    ]
  });
}
function sp(n=120) { return new Paragraph({ spacing:{ before:n, after:0 }, children:[] }); }
function pgBreak() { return new Paragraph({ children:[new PageBreak()] }); }
function divider(color="4472C4") {
  return new Paragraph({
    spacing: { before:120, after:120 },
    border: { bottom:{ style:BorderStyle.SINGLE, size:6, color, space:1 } },
    children: []
  });
}

// ══════════════════════════════════════════════════════
// 表1: 改訂後の記号体系 全ブロック一覧
// ══════════════════════════════════════════════════════
function tableSymbolSystem() {
  const W = [680, 1000, 1100, 1600, 4120]; // 8500
  return new Table({
    width: { size:8500, type:WidthType.DXA },
    columnWidths: W,
    rows: [
      row([
        ["ブロック",W[0],{center:true}], ["記号形式",W[1],{center:true}],
        ["段階数",W[2],{center:true}], ["指標・基準",W[3],{center:true}],
        ["内容",W[4],{center:true}]
      ], true, 38),
      row([["①時代大区分",W[0]],["Ⅰ・Ⅱ",W[1],{center:true}],["2",W[2],{center:true}],
          ["地質生成年代",W[3]],["Ⅰ＝古期（先新第三紀）　Ⅱ＝新期（新第三紀〜現在）",W[4]]],
          false, null),
      row([["②岩石種",W[0]],["a〜e",W[1],{center:true}],["各5",W[2],{center:true}],
          ["岩石種タイプ",W[3]],["各時代内の代表岩石・地質帯（別表参照）",W[4]]]),
      row([["③強度（S）",W[0],{bold:true,color:"1F3864"}],["S1〜S5 ／ S?",W[1],{center:true,bold:true,color:"1F3864"}],["5＋?",W[2],{center:true}],
          ["一軸圧縮強度（ISRM）",W[3]],["S1:>200／S2:100〜200／S3:25〜100／S4:5〜25／S5:<5  MN/m²\n情報なし＝S?",W[4],{small:true}]]),
      row([["④リスク（R）",W[0],{bold:true,color:"1F3864"}],["R1〜R6（複数可）",W[1],{center:true,bold:true,color:"1F3864"}],["6（複数可）",W[2],{center:true}],
          ["工学的リスク種別",W[3]],["R1:断層破砕帯／R2:変質変成／R3:膨張性／R4:冷却亀裂／R5:カルスト空洞／R6:未固結層",W[4],{small:true}]]),
      row([["⑤透水性（W）",W[0],{bold:true,color:"C0392B"}],["W1〜W5 ／ W?",W[1],{center:true,bold:true,color:"C0392B"}],["5＋?",W[2],{center:true,color:"C0392B",bold:true}],
          ["ルジオン値（Lu）",W[3],{bold:true}],["W1:<1／W2:1〜5／W3:5〜30／W4:30〜100／W5:>100 Lu\n情報なし＝W?　★今回新設★",W[4],{small:true,bold:true}]],
          false, null),
      row([["⑥基礎処理（G）",W[0],{bold:true,color:"C0392B"}],["G1〜G4 ／ G?",W[1],{center:true,bold:true,color:"C0392B"}],["4＋?",W[2],{center:true,color:"C0392B",bold:true}],
          ["グラウチング難易度",W[3],{bold:true}],["G1:軽微／G2:標準／G3:高難度／G4:特殊工法\n情報なし＝G?　★今回新設★",W[4],{small:true,bold:true}]]),
    ]
  });
}

// ══════════════════════════════════════════════════════
// 表2: 透水性コード W（5段階＋情報なし）
// ══════════════════════════════════════════════════════
function tablePermCode() {
  const W = [620, 1200, 1300, 1580, 1400, 2400]; // 8500
  const fills = ["D6EAF8","EAF4FB","F5FBFF","FEF9E7","FDEDEC","F5F5F5"];
  const data = [
    ["W1","極 低 透 水","< 1 Lu","< 10⁻⁸ m/s",
     "粒間浸透（事実上遮水）",
     "変成岩緻密部・花崗岩深部・粘土。カーテングラウト不要か最小限。遮水コアとして機能する場合あり"],
    ["W2","低 透 水","1〜5 Lu","10⁻⁸〜10⁻⁶ m/s",
     "微細亀裂浸透",
     "健全岩盤の標準状態。1列カーテングラウトで十分な遮水効果が得られる。Ⅰ-b（花崗岩）・Ⅱ-b（緻密溶岩）"],
    ["W3","中 透 水","5〜30 Lu","10⁻⁶〜10⁻⁵ m/s",
     "開口亀裂卓越",
     "亀裂系が主透水経路。カーテングラウト標準適用（1〜2列）。Ⅰ-c（砂岩開口亀裂）・Ⅱ-d（新第三系）"],
    ["W4","高 透 水","30〜100 Lu","10⁻⁵〜10⁻³ m/s",
     "大亀裂・断層破砕帯・冷却亀裂",
     "多段グラウト・多量注入必要。日新ダム実績値（10¹ m/day ≈ 10⁻⁴ m/s）がW4相当。Ⅱ-a（柱状節理）典型"],
    ["W5","極 高 透 水","> 100 Lu","> 10⁻³ m/s",
     "未固結層間隙・カルスト空洞",
     "グラウトが流出し効果不安定。止水矢板・コンクリート遮水壁等の補助工法が必要。Ⅱ-e・Ⅰ-e（石灰岩）"],
    ["W?","情 報 な し","—","—",
     "—",
     "ルジオン試験・透水試験の記録なし。設計・施工報告書参照で確認要。Ph.3調査優先ダムを特定するフラグ"],
  ];
  return new Table({
    width: { size:8500, type:WidthType.DXA }, columnWidths: W,
    rows: [
      row([
        ["コード",W[0],{center:true}], ["区分名",W[1],{center:true}],
        ["ルジオン値",W[2],{center:true}], ["透水係数 k",W[3],{center:true}],
        ["主な透水機構",W[4],{center:true}], ["工学的意味・北海道ダムへの対応",W[5],{center:true}]
      ], true, 42),
      ...data.map(([code, name, lu, k, mech, eng], i) => new TableRow({ children: [
        cell(code, W[0], { center:true, bold:true, color: code==="W?" ? "888888" : "C0392B", fill:fills[i] }),
        cell(name, W[1], { center:true, fill:fills[i], small:true }),
        cell(lu,   W[2], { center:true, fill:fills[i], small:true, bold:true }),
        cell(k,    W[3], { center:true, fill:fills[i], small:true }),
        cell(mech, W[4], { center:true, fill:fills[i], small:true }),
        cell(eng,  W[5], { fill:fills[i], xs:true }),
      ]}))
    ]
  });
}

// ══════════════════════════════════════════════════════
// 表3: 基礎処理コード G（4段階＋情報なし）
// ══════════════════════════════════════════════════════
function tableGroutCode() {
  const W = [620, 1100, 1580, 1500, 1200, 2500]; // 8500
  const fills = ["D6EAF8","EAF4FB","FEF9E7","FDEDEC","F5F5F5"];
  const data = [
    ["G1","軽 微",
     "< 50 kg/m（セメント）",
     "コンソリデーショングラウト（浅部のみ）",
     "W1〜W2・S1〜S2",
     "健全硬岩基礎（アーチダム相当）。グラウチングは補強目的。工期短・コスト小。豊平峡ダム相当"],
    ["G2","標 準",
     "50〜200 kg/m",
     "カーテングラウト1〜2列＋コンソリ",
     "W2〜W3・S2〜S3",
     "砂岩泥岩・安山岩・新第三系の標準。石狩川系中流ダム群の典型工法。Ⅰ-c・Ⅱ-d帯"],
    ["G3","高 難 度",
     "200〜500 kg/m（多段反復注入）",
     "多列カーテングラウト＋高圧注入・超微粒子セメント",
     "W3〜W4・S2〜S4",
     "溶結凝灰岩の柱状節理（冷却亀裂）・断層破砕帯。注入後の水圧試験反復確認要。日新・東郷・しろがね・古梅ダム典型"],
    ["G4","特 殊 工 法",
     "> 500 kg/m またはセメント以外",
     "止水矢板・コンクリート遮水壁・化学グラウト・多列カーテン併用",
     "W4〜W5・S4〜S5",
     "石灰岩カルスト（Ⅰ-e）・未固結礫層（Ⅱ-e）・蛇紋岩大破砕帯（Ⅰ-d）。グラウト流失対策必須。農業フィルダム沖積基礎"],
    ["G?","情 報 な し",
     "—",
     "設計・施工報告書参照要",
     "—",
     "グラウチング量・工法の施工記録なし。propylite.work・北海道開発局報告書・農業ダム施工記録を参照し確認"],
  ];
  return new Table({
    width: { size:8500, type:WidthType.DXA }, columnWidths: W,
    rows: [
      row([
        ["コード",W[0],{center:true}], ["難易度",W[1],{center:true}],
        ["想定注入量目安",W[2],{center:true}], ["主な工法",W[3],{center:true}],
        ["対応W・S",W[4],{center:true}], ["北海道ダムへの適用・事例",W[5],{center:true}]
      ], true, 42),
      ...data.map(([code,name,qty,method,ws,note_], i) => new TableRow({ children: [
        cell(code,   W[0], { center:true, bold:true, color: code==="G?" ? "888888" : "2E5090", fill:fills[i] }),
        cell(name,   W[1], { center:true, fill:fills[i], small:true }),
        cell(qty,    W[2], { center:true, fill:fills[i], xs:true }),
        cell(method, W[3], { fill:fills[i], xs:true }),
        cell(ws,     W[4], { center:true, fill:fills[i], small:true }),
        cell(note_,  W[5], { fill:fills[i], xs:true }),
      ]}))
    ]
  });
}

// ══════════════════════════════════════════════════════
// 表4: 強度コード S（現行＋「情報なし」追加）
// ══════════════════════════════════════════════════════
function tableStrengthCode() {
  const W = [620, 1100, 1500, 1280, 4000]; // 8500
  const fills = ["D5E8F7","E8F4FB","F5FBFF","FFF8E7","FDEBD0","F5F5F5"];
  const data = [
    ["S1","極 硬 岩","> 200 MN/m²","点荷重 > 8 MPa","花崗岩・強変成岩・強溶結凝灰岩（日新ダム相当 qu≈2,000 MN/m²）。Ⅰ-a・Ⅰ-b・Ⅱ-a（完全溶結）"],
    ["S2","硬　 岩","100〜200 MN/m²","点荷重 4〜8 MPa","堅硬砂岩・安山岩溶岩・花崗閃緑岩。Ⅰ-b・Ⅰ-c（堅硬部）・Ⅱ-b"],
    ["S3","中 硬 岩","25〜100 MN/m²","点荷重 1〜4 MPa","溶結凝灰岩（中程度）・一般砂岩・玄武岩。Ⅱ-a（中溶結）・Ⅱ-c・Ⅰ-c（平均）"],
    ["S4","軟　 岩","5〜25 MN/m²","点荷重 0.2〜1 MPa","軟質凝灰岩・泥岩・半固結礫岩。Ⅱ-c・Ⅱ-d・Ⅰ-d（変質部）"],
    ["S5","極軟岩・土質","< 5 MN/m²","点荷重 < 0.2 MPa","未固結礫層・砂・粘土（沖積〜段丘堆積物）。Ⅱ-e全般"],
    ["S?","情 報 な し","—","—","一軸圧縮強度試験データなし。設計・施工報告書参照要。今回追加"],
  ];
  return new Table({
    width: { size:8500, type:WidthType.DXA }, columnWidths: W,
    rows: [
      row([
        ["コード",W[0],{center:true}], ["区分名",W[1],{center:true}],
        ["一軸圧縮強度 qu",W[2],{center:true}], ["補助指標",W[3],{center:true}],
        ["代表岩石・北海道地質区分との対応",W[4],{center:true}]
      ], true, 42),
      ...data.map(([code,name,qu,sub,rock], i) => new TableRow({ children: [
        cell(code, W[0], { center:true, bold:true, color: code==="S?" ? "888888" : "1F3864", fill:fills[i] }),
        cell(name, W[1], { center:true, fill:fills[i], small:true }),
        cell(qu,   W[2], { center:true, fill:fills[i], small:true, bold:true }),
        cell(sub,  W[3], { center:true, fill:fills[i], small:true }),
        cell(rock, W[4], { fill:fills[i], xs:true }),
      ]}))
    ]
  });
}

// ══════════════════════════════════════════════════════
// 表5: 地質区分別 W・G 標準値（核心表）
// ══════════════════════════════════════════════════════
function tableGeoHydraulic() {
  const W = [700, 1100, 1300, 900, 1600, 800, 800, 1300]; // 8500
  const secData = [
    { label:"■ Ⅰ類（古期地質）　Prior to Neogene", fill:"1F3864", rows:[
      ["Ⅰ-a","変成岩\n（結晶片岩・片麻岩）","S1〜S2",
       "変成葉理沿いの浸透が主体。緻密部は極低透水。活断層沿いに高透水帯（W4）が局在",
       "W1〜W2\n（断層帯W4）","G1〜G2\n（断層帯G3）",
       "断層の走向・傾斜が透水性を支配。局部的高透水帯の事前探査が重要"],
      ["Ⅰ-b","花崗岩・\n花崗閃緑岩","S1",
       "急冷節理・方状節理が主透水経路。深部は極低透水。風化帯でW3に上昇",
       "W1〜W2\n（風化帯W3）","G1〜G2",
       "豊平峡ダム基礎相当。アーチダムの理想的基盤。グラウト量小"],
      ["Ⅰ-c","砂岩泥岩互層\n（タービダイト）","S2〜S3",
       "層理面・開口亀裂が主経路。砂岩層W3、泥岩層W1〜W2。層理傾斜が漏水方向を規定",
       "W2〜W3","G2〜G3",
       "美生ダム実績：開口亀裂（砂岩部卓越）→W3/G2。層理面傾斜の把握が遮水工設計の鍵"],
      ["Ⅰ-d","超苦鉄質岩・\n蛇紋岩","S1〜S4",
       "蛇紋岩化破砕帯でW4。母岩部はW2。膨潤性鉱物（クリソタイル等）で透水性が経時変動",
       "W3〜W4\n（破砕帯W4）","G3〜G4",
       "様似・幌満かんらん岩体周辺。グラウト効果の確認難。化学グラウト併用検討要"],
      ["Ⅰ-e","石灰岩・\nチャート・変成堆積岩","S2〜S3",
       "石灰岩はカルスト空洞・溶食亀裂でW5。チャートはW1。渡島古生界では両者が混在",
       "W1（チャート）\n〜W5（石灰岩）","G2〜G4",
       "カルスト空洞の事前探査（ボーリング・物理探査）が最重要。空洞充填→止水工の順序"],
    ]},
    { label:"■ Ⅱ類（新期地質）　Neogene–Quaternary", fill:"5B4FA0", rows:[
      ["Ⅱ-a","溶結凝灰岩\n（カルデラ起源）","S1〜S3",
       "柱状節理（冷却亀裂）が卓越した透水経路。鉛直・水平の大亀裂。節理開口幅でW3〜W4",
       "W3〜W4","G3",
       "日新ダム実績：10¹ m/day≈W4、冷却亀裂（鉛直・水平）。東郷・しろがね・古梅も同様パターン"],
      ["Ⅱ-b","安山岩・玄武岩\n（溶岩）","S2〜S3",
       "流理・板状節理・溶岩間堆積物が透水層形成。完全溶岩体はW2。冷却面・溶岩間でW4に上昇",
       "W2〜W3\n（冷却面W4）","G2〜G3",
       "溶岩スタック構造の層序把握が設計の鍵。中新世〜更新世溶岩帯（大雪・ニセコ等）"],
      ["Ⅱ-c","凝灰岩・\n火山礫凝灰岩","S3〜S4",
       "固結度・空隙率に依存した粒間浸透と亀裂透水の混合。固結度が高いほどW2〜W3",
       "W2〜W3","G2〜G3",
       "後志・空知・道東中新統。固結度変動が大きく、同一ダムサイトでも±1コード変動あり"],
      ["Ⅱ-d","砂岩・泥岩\n（新第三系）","S3〜S4",
       "Ⅰ-cより続成浅く空隙率高め。砂岩層で粒間浸透、層理面沿いの透水が顕著",
       "W2〜W3","G2",
       "天北地向斜・石狩低地帯縁辺の農業ダム標準地質。標準グラウトで概ね対応可"],
      ["Ⅱ-e","未固結礫層・\n段丘堆積物","S4〜S5",
       "粒間透水が卓越。礫層はW5（k>10⁻³ m/s）。フィルダムの堤体・遮水ゾーン設計に直結",
       "W4〜W5","G3〜G4",
       "石狩・名寄平野縁辺の農業フィルダム。止水矢板・コンクリート遮水壁等の補助工法必須"],
    ]},
  ];

  const hdrRow = row([
    ["区分\nコード",W[0],{center:true}], ["代表岩石種",W[1],{center:true}],
    ["強度\n（S）",W[2],{center:true}], ["透水機構・主経路",W[3],{center:true}],  // ← 列幅修正済
    ["透水性\n（W）★",W[4],{center:true}], ["基礎処理\n（G）★",W[5],{center:true}],  // ← 列インデックス修正
    ["北海道ダム工学上の要点・実績",W[6],{center:true}]  // ← 最後の列
  ], true, 44);
  // ↑ W配列と列数が合うよう修正版を下で再定義

  // 列定義を整理（7列で8500）
  const W2 = [700, 1200, 800, 1900, 850, 800, 2250];
  const hdrRow2 = row([
    ["区分\nコード",W2[0],{center:true}], ["代表岩石種",W2[1],{center:true}],
    ["強度（S）",W2[2],{center:true}], ["透水機構・主な経路",W2[3],{center:true}],
    ["透水性（W）\n★新設★",W2[4],{center:true,color:"C0392B",bold:true}],
    ["基礎処理（G）\n★新設★",W2[5],{center:true,color:"2E5090",bold:true}],
    ["北海道ダム工学上の要点・実績",W2[6],{center:true}],
  ], true, 46);

  const tableRows = [hdrRow2];
  let altIdx = 0;
  for (const sec of secData) {
    tableRows.push(sectionRow(sec.label, sec.fill, 7));
    for (const [code,rock,s,mech,w,g,tip] of sec.rows) {
      const fill = altIdx % 2 === 0 ? "F8F8F8" : "FFFFFF";
      altIdx++;
      tableRows.push(new TableRow({ children: [
        cell(code, W2[0], { center:true, bold:true, color:"1F3864", fill:"EEF4FF" }),
        cell(rock, W2[1], { center:true, fill, small:true }),
        cell(s,    W2[2], { center:true, fill, small:true }),
        cell(mech, W2[3], { fill, xs:true }),
        cell(w,    W2[4], { center:true, bold:true, color:"C0392B", fill }),
        cell(g,    W2[5], { center:true, bold:true, color:"2E5090", fill }),
        cell(tip,  W2[6], { fill, xs:true }),
      ]}));
    }
  }
  return new Table({ width:{ size:8500, type:WidthType.DXA }, columnWidths:W2, rows:tableRows });
}

// ══════════════════════════════════════════════════════
// 表6: 改訂後の完全記号 適用例
// ══════════════════════════════════════════════════════
function tableSymbolExamples() {
  const W = [3000, 700, 4800]; // 8500
  const confFill = { A:"D5F5E3", B:"D6EAF8", C:"FEF9E7" };
  const data = [
    // 信頼度A（既存データ4基）
    ["Ⅱ-a（S1/R4・W4/G3）・Ⅱ-e（S5/R6・W5/G4）","A",
     "溶結凝灰岩（極硬岩・冷却亀裂R4・高透水W4・多段グラウトG3）上に河床礫層（極軟弱・未固結R6・極高透水W5・特殊工法G4）。日新ダム相当。実績：qu≈2,000 MN/m²、透水性10¹ m/day≈W4"],
    ["Ⅱ-a（S3/R4・W4/G3）","A",
     "中程度溶結凝灰岩（中硬岩・冷却亀裂R4・高透水W4・多段グラウトG3）。東郷・聖台・新区画・古梅ダム相当。プロファイル共通：柱状節理が主要漏水経路"],
    ["Ⅱ-a（S3/R4・W4/G3）・Ⅱ-e（S5/R6・W5/G4）","A",
     "溶結凝灰岩（W4/G3）＋河床礫層（W5/G4）。しろがね・日新ダム共通パターン。河床二次堆積物の止水工設計が重要課題"],
    ["Ⅰ-c（S2/R1・W3/G2）・Ⅱ-e（S5/R6・W5/G4）","A",
     "白亜紀砂岩粘板岩タービダイト（開口亀裂R1・中透水W3・標準グラウトG2）上に未固結礫層（W5/G4）。美生ダム実績：砂岩部卓越開口亀裂・Well locked礫層"],
    // 信頼度B（推定）
    ["Ⅰ-b（S1/R1・W1/G1）","B",
     "花崗岩（極硬岩・極低透水W1・グラウト軽微G1）。豊平峡アーチダム相当の理想的基礎岩盤。グラウチングは補強目的に限定"],
    ["Ⅱ-a（S2/R4・W4/G3）","B",
     "強溶結凝灰岩（支笏カルデラ起源）。漁川・千歳川系ダム相当。柱状節理のW4は確実だが詳細Lu値は施工報告書確認要"],
    ["Ⅰ-a（S1/R1・R2・W2/G2）・Ⅱ-e（S5/R6・W5/G4）","B",
     "日高変成帯変成岩（断層リスクR1・変質R2・低透水W2・標準グラウトG2）上に礫層（W5/G4）。沙流川・新冠川系ダム典型"],
    ["Ⅰ-d（S2/R2・R3・W4/G4）・Ⅱ-e（S5/R6・W5/G4）","B",
     "幌満かんらん岩・蛇紋岩（破砕帯高透水W4・特殊工法G4）。様似・幌満川第三ダム相当。蛇紋岩膨潤R3と高透水W4の複合が設計最難関"],
    // 信頼度C（情報欠如含む）
    ["Ⅰ-c（S2/R1・W?/G?）・Ⅱ-d（S3/R1・W?/G?）","C",
     "道北・留萌帯の砂岩泥岩＋新第三系。強度・リスクは推定可だが透水試験データ未参照のためW?/G?。Ph.1で施工報告書確認要"],
    ["Ⅱ-e（S4/R6・W?/G?）","C",
     "石狩平野縁辺農業フィルダム。W4〜W5/G3〜G4が推定されるが試験記録未確認のためW?/G?。未固結基礎の農業ダムとして優先確認対象"],
  ];
  return new Table({
    width: { size:8500, type:WidthType.DXA }, columnWidths: W,
    rows: [
      row([
        ["完全記号（S・R・W・G 全ブロック）",W[0],{center:true}],
        ["信頼度",W[1],{center:true}],
        ["読み方・地質工学的意味",W[2],{center:true}]
      ], true, 38),
      ...data.map(([sym,conf,desc]) => {
        const fill = confFill[conf];
        return new TableRow({ children: [
          cell(sym,  W[0], { bold:true, color:"1F3864", fill, small:true }),
          cell(conf, W[1], { center:true, bold:true, fill }),
          cell(desc, W[2], { fill, xs:true }),
        ]});
      })
    ]
  });
}

// ══════════════════════════════════════════════════════
// ドキュメント本体
// ══════════════════════════════════════════════════════
const doc = new Document({
  styles: {
    default: { document: { run: { font:"游明朝", size:22 } } },
    paragraphStyles: [
      { id:"Heading1", name:"Heading 1", basedOn:"Normal", next:"Normal",
        run:{ size:36, bold:true, font:"游明朝", color:"1F3864" },
        paragraph:{ spacing:{ before:480, after:220 }, outlineLevel:0 } },
      { id:"Heading2", name:"Heading 2", basedOn:"Normal", next:"Normal",
        run:{ size:28, bold:true, font:"游明朝", color:"2E5090" },
        paragraph:{ spacing:{ before:320, after:140 }, outlineLevel:1 } },
      { id:"Heading3", name:"Heading 3", basedOn:"Normal", next:"Normal",
        run:{ size:24, bold:true, font:"游明朝", color:"4472C4" },
        paragraph:{ spacing:{ before:220, after:100 }, outlineLevel:2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width:11906, height:16838 },
        margin: { top:1700, right:1500, bottom:1700, left:1800 }
      }
    },
    footers: { default: new Footer({ children: [
      new Paragraph({ alignment:AlignmentType.CENTER, children: [
        new TextRun({ text:"北海道ダム地質分類システム　水理的特性コード体系（再提案）　", font:"游明朝", size:18, color:"888888" }),
        new TextRun({ children:[PageNumber.CURRENT], font:"游明朝", size:18, color:"888888" }),
        new TextRun({ text:" / ", font:"游明朝", size:18, color:"888888" }),
        new TextRun({ children:[PageNumber.TOTAL_PAGES], font:"游明朝", size:18, color:"888888" }),
      ]})
    ]})},
    children: [

      // ── 表紙 ──────────────────────────────────────
      sp(1000),
      new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:240},
        children:[new TextRun({ text:"北海道ダム地質分類システム", font:"游明朝", size:52, bold:true, color:"1F3864" })] }),
      new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:100},
        children:[new TextRun({ text:"水理的特性コード体系　再提案書", font:"游明朝", size:40, color:"2E5090" })] }),
      divider("2E5090"),
      new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:160,after:80},
        children:[new TextRun({ text:"Hydraulic Properties Coding for Dam Foundation Geology Classification", font:"Times New Roman", size:24, italics:true, color:"555555" })] }),
      sp(200),
      new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:60},
        children:[new TextRun({ text:"― 透水性コード（W）・基礎処理コード（G）の新設と強度コード（S）・リスク指標（R）への統合 ―", font:"游明朝", size:23, color:"333333" })] }),
      sp(400),
      new Paragraph({ alignment:AlignmentType.RIGHT, spacing:{before:0,after:80},
        children:[new TextRun({ text:"対象：北海道既設・建設中ダム 188基", font:"游明朝", size:20, color:"555555" })] }),
      new Paragraph({ alignment:AlignmentType.RIGHT, spacing:{before:0,after:80},
        children:[new TextRun({ text:"作成日：2026年3月", font:"游明朝", size:20, color:"555555" })] }),
      new Paragraph({ alignment:AlignmentType.RIGHT, spacing:{before:0,after:0},
        children:[new TextRun({ text:"【改訂】強度（S）へ「S?（情報なし）」を追加　★透水性（W）・基礎処理（G）を新設★", font:"游明朝", size:20, bold:true, color:"C0392B" })] }),

      pgBreak(),

      // ── 1. 改訂の趣旨 ──────────────────────────────
      h1("1. 改訂の趣旨"),
      para("前版の分類体系では、強度コード（S）とリスク指標（R）によって岩盤の力学的特性と工学的リスクを記号化した。しかし、ダム基礎の工学的評価に不可欠なもう一つの軸——水理的特性（岩盤透水性・グラウチングによる基礎処理難易度）——が体系的に表現されていなかった。"),
      para("本再提案では、強度コード（S）と完全に同格・同構造の新コードとして以下2つを新設し、既存記号の括弧内に水理ブロックとして統合する。"),
      sp(100),
      new Table({ width:{size:7500,type:WidthType.DXA}, columnWidths:[1200,1200,5100],
        rows:[
          row([["新設コード",1200,{center:true}],["略号",1200,{center:true}],["定義と段階数",5100,{center:true}]], true),
          row([["透水性コード",1200,{bold:true}],["W1〜W5・W?",1200,{center:true,bold:true,color:"C0392B"}],
               ["ルジオン値（Lu）基準の5段階。情報なし＝W?",5100]]),
          row([["基礎処理コード",1200,{bold:true}],["G1〜G4・G?",1200,{center:true,bold:true,color:"2E5090"}],
               ["グラウチング難易度の4段階。情報なし＝G?",5100]]),
        ]
      }),
      sp(120),
      para("記号の構成はSコード・Rコードの直後に「・」で区切って「W/G」ブロックを付加する。情報が存在しない場合は「W?/G?」を明記することで、「調査済みで情報がある」状態と「未調査」状態を明確に区別する。"),

      // ── 2. 改訂後の記号体系 ────────────────────────
      h1("2. 改訂後の記号体系　全ブロック一覧"),
      para("分類記号は6ブロックで構成する。①②が地質時代と岩石種、③④が力学的ブロック（従前）、⑤⑥が水理的ブロック（今回新設）。", {noIndent:true}),
      sp(120),
      tableSymbolSystem(),
      sp(100),
      para("【記号構成例】", {noIndent:true, bold:true}),
      para("　完全記号：　Ⅱ-a（S3/R4・W4/G3）", {noIndent:true}),
      para("　情報欠如：　Ⅰ-c（S2/R1・W?/G?）　← W・Gの情報なしを明記", {noIndent:true}),
      para("　複合地質：　Ⅰ-c（S2/R1・W3/G2）・Ⅱ-e（S5/R6・W5/G4）", {noIndent:true}),
      note("括弧内の「S/R」ブロックと「W/G」ブロックは「・」で区切る。力学特性と水理特性を視覚的に分離するためである。"),

      pgBreak(),

      // ── 3. 強度コード S（S?追加） ──────────────────
      h1("3. 強度コード（S）　——「情報なし（S?）」を追加"),
      para("前版との変更点は「S?（情報なし）」の追加のみ。強度試験データが存在しないダムへの対応を明示する。定義・判定基準は変更なし。", {noIndent:true}),
      sp(120),
      tableStrengthCode(),
      sp(80),
      note("基準：ISRM一軸圧縮強度分類（1979）および土木学会岩盤分類指針。点荷重試験（Is50）はqu ≈ 22×Is50 で換算。"),

      pgBreak(),

      // ── 4. 透水性コード W ──────────────────────────
      h1("4. 透水性コード（W）　——新設"),
      h2("4.1 定義"),
      para("ルジオン値（Lu）を主指標として5段階に区分する。ルジオン試験データが存在しない場合は「W?」を付与する。1 Lu ≈ 1.3×10⁻⁷ m/s（圧力 1 MPa 換算）。", {noIndent:true}),
      sp(120),
      tablePermCode(),
      sp(100),
      note("日新ダムの実績透水性「10¹ m/day 程度」はおよそ 1.2×10⁻⁴ m/s に相当し、W4（30〜100 Lu 相当帯）に位置づける。"),
      note("同一地質区分内でも亀裂開口幅・充填状態・地下水位によりLu値は1〜2オーダー変動する。単一値でなく範囲（例：W3〜W4）で表現することが実態に即している。"),
      note("「W?」の解消方法：設計・施工報告書のルジオン試験結果、グラウチング記録（注入量・返却率）を参照。propylite.work・北海道開発局・農林水産省農業ダム施工記録が主要出典。"),

      h2("4.2 工学的意義"),
      para("透水性コード（W）は以下の設計項目に直接連動する。", {noIndent:true}),
      para("カーテングラウトの列数・深度・注入圧の設計基礎値となる。W1〜W2ではカーテン省略または1列、W4〜W5では多列・多段が必要となる。"),
      para("グラウチング材料の選定（普通ポルトランドセメント → 粒調セメント → 超微粒子セメント → 化学グラウト）の判断基準となる。"),
      para("フィルダムの堤体浸潤線管理・パイピング評価における基礎漏水量の推定に活用する。"),
      para("ダム完成後の長期モニタリング計画（間隙水圧計配置・漏水量測定）の設計基礎となる。"),

      pgBreak(),

      // ── 5. 基礎処理コード G ────────────────────────
      h1("5. 基礎処理コード（G）　——新設"),
      h2("5.1 定義"),
      para("グラウチング注入量・工法の複雑さ・特殊工法の要否を4段階で区分する。施工記録がない場合は「G?」を付与する。ICOLD Bulletin 88（Rock Foundations for Dams）の分類を参考に北海道の実情に合わせて体系化した。", {noIndent:true}),
      sp(120),
      tableGroutCode(),
      sp(100),
      note("注入量目安はセメントグラウト（W/C=0.5〜1.0）での1列カーテングラウト換算。多列施工の場合は列数倍の積算となる。"),
      note("G4の「特殊工法」：化学グラウト（ウレタン・アクリル系）、鋼矢板遮水壁、地中連続壁、薬液注入工法等を含む。"),
      note("「G?」の解消方法：施工記録のグラウチング量（kg/m）・注入ステージ数・返却率・地盤改良工法の記録を参照。"),

      h2("5.2 W コードと G コードの対応関係"),
      para("透水性（W）と基礎処理難易度（G）は密接に対応するが同一ではない。亀裂の方向性・連続性・充填状況によって、同じWコードでもGコードが変わる場合がある。", {noIndent:true}),
      sp(120),
      new Table({ width:{size:8000,type:WidthType.DXA}, columnWidths:[800,800,6400],
        rows:[
          row([["W",800,{center:true}],["対応G",800,{center:true}],["判断基準・留意点",6400,{center:true}]], true, 36),
          row([["W1","800",{center:true,bold:true,color:"C0392B"}],["G1",800,{center:true,bold:true,color:"2E5090"}],
               ["極低透水。カーテングラウト原則不要。アーチダム・薄型重力式の理想基礎",6400,{small:true}]]),
          row([["W2",800,{center:true,bold:true,color:"C0392B"}],["G1〜G2",800,{center:true,bold:true,color:"2E5090"}],
               ["標準健全岩盤。1列カーテングラウトで十分。注入量少なく施工短期",6400,{small:true}]]),
          row([["W3",800,{center:true,bold:true,color:"C0392B"}],["G2",800,{center:true,bold:true,color:"2E5090"}],
               ["開口亀裂卓越。標準カーテン1〜2列。砂岩泥岩・凝灰岩の標準パターン（美生ダム実績）",6400,{small:true}]]),
          row([["W4",800,{center:true,bold:true,color:"C0392B"}],["G3",800,{center:true,bold:true,color:"2E5090"}],
               ["冷却亀裂・大亀裂。多列・多段注入。溶結凝灰岩（日新・東郷・しろがね）の典型。注入後の水圧試験反復要",6400,{small:true}]]),
          row([["W5",800,{center:true,bold:true,color:"C0392B"}],["G4",800,{center:true,bold:true,color:"2E5090"}],
               ["未固結層・カルスト空洞。グラウト流失。止水矢板・コンクリート遮水壁が必須。農業フィルダム基礎・石灰岩",6400,{small:true}]]),
          row([["W?",800,{center:true,bold:true,color:"888888"}],["G?",800,{center:true,bold:true,color:"888888"}],
               ["データなし。ルジオン試験未実施または記録未参照。Ph.1情報収集フェーズで解消すべき優先課題",6400,{small:true}]]),
        ]
      }),

      pgBreak(),

      // ── 6. 地質区分別 W・G 標準値 ──────────────────
      h1("6. 地質区分別　透水性コード（W）・基礎処理コード（G）標準値"),
      para("各地質区分（Ⅰ-a〜Ⅱ-e）についてW・Gの標準的な範囲を示す。強度コード（S）との対照で読む。個別ダムへの適用では当該ダムの試験データ・施工記録による確認が前提。", {noIndent:true}),
      sp(120),
      tableGeoHydraulic(),
      sp(80),
      note("表中のW・Gは各地質区分の「典型的な範囲」。地域差・深度差・変質度により±1段階変動することがある。"),
      note("括弧書き（例：断層帯W4、風化帯W3）は同一ダムサイト内での局部高透水帯を示す。設計では通常、最悪部を基準として採用する。"),

      pgBreak(),

      // ── 7. 適用例 ──────────────────────────────────
      h1("7. 改訂後の完全記号　適用例"),
      para("以下は全6ブロック（S・R・W・G）を含む完全記号の適用例。信頼度Aは既存データ（propylite.work・設計報告書）に基づき、BはGeoNAVI地質図・文献から推定、Cは位置・水系類推による暫定値。「W?/G?」を含む記号は Ph.1 情報収集で順次更新する。", {noIndent:true}),
      sp(120),
      tableSymbolExamples(),
      sp(80),
      note("信頼度AのW・Gコードを確定できたのは既存データ4基（日新・東郷・しろがね・美生）のみ。信頼度B（60基）はW/Gの推定範囲を付与。信頼度C（121基）はW?/G?として「要確認」を明示する。"),
      note("W?/G?を含むダムがPh.3調査の優先対象。特に G3〜G4 が推定される溶結凝灰岩帯（16基）・石灰岩帯（13基）・蛇紋岩体（2基）を最優先とする。"),

      pgBreak(),

      // ── 8. DB更新方針 ──────────────────────────────
      h1("8. データベース（Excel・KMZ）更新方針"),
      h2("8.1 Excelデータベース列追加"),
      para("現在のダム地質区分DBに以下の列を追加する。", {noIndent:true}),
      sp(80),
      new Table({ width:{size:8000,type:WidthType.DXA}, columnWidths:[2200,1400,4400],
        rows:[
          row([["追加列名",2200,{center:true}],["入力例",1400,{center:true}],["備考",4400,{center:true}]], true),
          row([["古期_透水性W",2200],["W3",1400,{center:true}],["ルジオン値（Lu）から判定。範囲表記可（W2〜W3）",4400,{small:true}]]),
          row([["古期_基礎処理G",2200],["G2",1400,{center:true}],["グラウチング記録から判定",4400,{small:true}]]),
          row([["新期_透水性W",2200],["W4",1400,{center:true}],["同上",4400,{small:true}]]),
          row([["新期_基礎処理G",2200],["G3",1400,{center:true}],["同上",4400,{small:true}]]),
          row([["記号（完全版）",2200],["Ⅱ-a（S3/R4・W4/G3）",1400,{center:true,small:true}],["S・R・W・G全ブロック統合記号",4400,{small:true}]]),
          row([["水理信頼度",2200],["A / B / C",1400,{center:true}],["W・Gコードの信頼度（強度信頼度と独立管理）",4400,{small:true}]]),
        ]
      }),
      sp(140),
      h2("8.2 W?/G?の解消フロー"),
      para("「W?/G?」を付与したダムについては、以下の優先順位で情報収集・更新を進める。", {noIndent:true}),
      para("① 設計・施工報告書のルジオン試験結果・グラウチング記録（kg/m）を参照しW・Gを確定（信頼度A相当）"),
      para("② propylite.work 掲載ページ・北海道開発局公開報告書から透水性記述を抽出（信頼度B相当）"),
      para("③ 上記なき場合は地質区分別標準値（本書表6）を暫定適用し、信頼度Cとして記録"),
      para("優先ターゲット：G3〜G4が推定されるW4〜W5帯ダム（溶結凝灰岩16基・石灰岩13基・蛇紋岩2基・未固結基礎28基）。これらはダムの安全性・長期性能に直結するため最優先で調査する。"),

      sp(300),
      divider(),
      new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:200,after:80},
        children:[new TextRun({ text:"以　上", font:"游明朝", size:22 })] }),
      new Paragraph({ alignment:AlignmentType.RIGHT, spacing:{before:60,after:0},
        children:[new TextRun({ text:"2026年3月　水理的特性コード（W・G）再提案", font:"游明朝", size:20, color:"555555" })] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/水理的特性コード体系_再提案書.docx', buf);
  console.log('完成');
}).catch(e => { console.error(e); process.exit(1); });
