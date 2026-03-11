// 日本全国ダム地質分類 作業計画書（改訂版2）
// PDFの体系を全国展開へ完全統合

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
} = require('docx');
const fs = require('fs');

// ─── ユーティリティ ───────────────────────────────────────────────
const A4W = 11906, A4H = 16838, MAR = 1000, CONTENT_W = A4W - MAR * 2;

const C = {
  h1bg:'1F4E79', h1fg:'FFFFFF',
  h2bg:'2E75B6', h2fg:'FFFFFF',
  h3bg:'BDD7EE', h3fg:'1F4E79',
  h4bg:'DEEAF1', h4fg:'1F4E79',
  th:'2E75B6',   thfg:'FFFFFF',
  r1:'FFFFFF',   r2:'EBF3FB',
  good:'E2EFDA', warn:'FFF2CC', risk:'FCE4D6',
  notebg:'FFF2CC', notebdr:'F4B942',
  bdr:'2E75B6',  lbdr:'BDD7EE',
  red:'C00000',  green:'375623', brown:'833C00',
};

const bdr = (col=C.bdr)=>({style:BorderStyle.SINGLE,size:4,color:col});
const bdrs=(col)=>({top:bdr(col),bottom:bdr(col),left:bdr(col),right:bdr(col)});

function HR(){return new Paragraph({border:{bottom:{style:BorderStyle.SINGLE,size:6,color:C.bdr,space:1}},children:[],spacing:{before:60,after:60}});}
function PB(){return new Paragraph({children:[new PageBreak()]});}
function SP(n=1){return new Paragraph({children:[new TextRun('')],spacing:{before:60*n,after:0}});}

function H1(t){return new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:t,bold:true,color:C.h1fg,size:32,font:'Arial'})],shading:{fill:C.h1bg,type:ShadingType.CLEAR},spacing:{before:360,after:180},indent:{left:180}});}
function H2(t){return new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun({text:t,bold:true,color:C.h2fg,size:28,font:'Arial'})],shading:{fill:C.h2bg,type:ShadingType.CLEAR},spacing:{before:280,after:140},indent:{left:120}});}
function H3(t){return new Paragraph({heading:HeadingLevel.HEADING_3,children:[new TextRun({text:t,bold:true,color:C.h3fg,size:26,font:'Arial'})],shading:{fill:C.h3bg,type:ShadingType.CLEAR},spacing:{before:200,after:100},indent:{left:80}});}
function H4(t){return new Paragraph({heading:HeadingLevel.HEADING_4,children:[new TextRun({text:t,bold:true,color:C.h4fg,size:24,font:'Arial'})],shading:{fill:C.h4bg,type:ShadingType.CLEAR},spacing:{before:160,after:80},indent:{left:60}});}

function P(t,opts={}){return new Paragraph({children:[new TextRun({text:t,size:22,font:'MS Mincho',color:opts.color||'000000',bold:opts.bold||false,italics:opts.italic||false})],spacing:{before:opts.before||80,after:opts.after||80},indent:opts.indent?{left:opts.indent}:undefined,alignment:opts.align||AlignmentType.JUSTIFIED});}

function BUL(t,lv=0,col=null){return new Paragraph({numbering:{reference:`bul${lv}`,level:0},children:[new TextRun({text:t,size:22,font:'MS Mincho',color:col||'000000'})],spacing:{before:60,after:60}});}
function NUM(t,lv=0){return new Paragraph({numbering:{reference:`num${lv}`,level:0},children:[new TextRun({text:t,size:22,font:'MS Mincho'})],spacing:{before:60,after:60}});}

function cell(t,opts={}){
  const w=opts.w||Math.floor(CONTENT_W/4);
  return new TableCell({
    borders:bdrs(opts.bc||C.lbdr),
    width:{size:w,type:WidthType.DXA},
    shading:opts.fill?{fill:opts.fill,type:ShadingType.CLEAR}:undefined,
    margins:{top:80,bottom:80,left:100,right:100},
    verticalAlign:VerticalAlign.CENTER,
    columnSpan:opts.span||1,
    children:[new Paragraph({alignment:opts.align||AlignmentType.LEFT,children:[new TextRun({text:String(t),bold:opts.bold||false,color:opts.color||'000000',size:opts.sz||20,font:opts.font||'MS Mincho'})]})]
  });
}

function TBL(headers,rows,widths){
  const total=CONTENT_W;
  const defW=Math.floor(total/headers.length);
  const cw=widths||Array(headers.length).fill(defW);
  const hrow=new TableRow({tableHeader:true,children:headers.map((h,i)=>cell(typeof h==='string'?h:h.text,{fill:C.th,color:C.thfg,bold:true,sz:20,w:cw[i],bc:C.bdr,align:AlignmentType.CENTER}))});
  const drows=rows.map((r,ri)=>new TableRow({children:r.map((c,ci)=>{
    const v=typeof c==='object'?c:{text:String(c)};
    return cell(v.text??String(c),{fill:v.fill||(ri%2===1?C.r2:C.r1),bold:v.bold||false,color:v.color||'000000',w:cw[ci],bc:C.lbdr,align:v.align||AlignmentType.LEFT,sz:v.sz||20});
  })}));
  return new Table({width:{size:total,type:WidthType.DXA},columnWidths:cw,rows:[hrow,...drows]});
}

function NOTE(title,items){
  const nc=[];
  if(title){
    nc.push(new Paragraph({
      children:[new TextRun({text:'【'+title+'】',bold:true,size:22,color:C.brown,font:'MS Mincho'})],
      spacing:{before:80,after:60}
    }));
  }
  items.forEach(function(l){
    nc.push(new Paragraph({
      children:[new TextRun({text:l,size:21,font:'MS Mincho'})],
      spacing:{before:40,after:40},
      indent:{left:200}
    }));
  });
  var ntcell=new TableCell({
    borders:{top:bdr(C.notebdr),bottom:bdr(C.notebdr),left:{style:BorderStyle.THICK,size:12,color:C.notebdr},right:bdr(C.notebdr)},
    width:{size:CONTENT_W,type:WidthType.DXA},
    shading:{fill:C.notebg,type:ShadingType.CLEAR},
    margins:{top:100,bottom:100,left:200,right:200},
    children:nc
  });
  return new Table({
    width:{size:CONTENT_W,type:WidthType.DXA},
    columnWidths:[CONTENT_W],
    rows:[new TableRow({children:[ntcell]})]
  });
}

// ─── 本文 ───────────────────────────────────────────────────────
const ch=[];

const numbering={config:[
  {reference:'bul0',levels:[{level:0,format:LevelFormat.BULLET,text:'●',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:480,hanging:280}}}}]},
  {reference:'bul1',levels:[{level:0,format:LevelFormat.BULLET,text:'○',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:280}}}}]},
  {reference:'num0',levels:[{level:0,format:LevelFormat.DECIMAL,text:'%1.',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:560,hanging:320}}}}]},
]};

// ═══ 表紙 ════════════════════════════════════════════════════════
ch.push(SP(8));
ch.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'日本全国ダム地質分類',bold:true,size:56,font:'Arial',color:C.h1bg})],spacing:{before:0,after:200}}));
ch.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'体系的分析・全国展開 作業計画書',bold:true,size:36,font:'Arial',color:C.h2bg})],spacing:{before:0,after:200}}));
ch.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'改訂版2（2026年3月）',size:28,font:'MS Mincho',color:'444444'})],spacing:{before:0,after:300}}));
ch.push(HR());
ch.push(SP(2));
ch.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'北海道既設ダム188基をパイロットとし、全国の既設ダムへの展開を目指す',size:24,font:'MS Mincho',color:'666666'})]}));
ch.push(SP(2));
ch.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'propylite.work',size:24,font:'Arial',color:'888888'})]}));
ch.push(PB());

// ═══ 0. 改訂の趣旨 ═══════════════════════════════════════════════
ch.push(H1('0. 改訂の趣旨と基本方針'));
ch.push(P('本計画書は「北海道ダム地質分類プロジェクト 全体計画書（2026年3月版）」を基礎として、以下の3点を反映した全面改訂版である。'));
ch.push(SP());
ch.push(NOTE('改訂の3本柱',[
  '①  北海道の全ダム（188基）を最初から分析対象とし、段階的に精度を高める。',
  '②  次世代コード（D：経時劣化 / E：地震応答 / C：流域地質 / V：気候変動 / K：知識段階 / Q：耐震照査結果 / Q：耐震照査結果）を全国共通体系に組み込む。',
  '③  GeoNAVI Web APIによる自動取得を全国展開の基盤とし、人的情報による修正を上位レイヤーとして重ねる。',
]));
ch.push(PB());

// ═══ 第Ⅰ部 分類体系（全国共通） ══════════════════════════════════
ch.push(H1('第Ⅰ部　分類体系（全国共通版）'));
ch.push(P('本章で定義する分類体系は北海道での構築・検証を経て、全国適用版として確定したものである。日本全国のすべての対象ダムにこの体系を適用する。'));

// 1. 時代区分・岩石種
ch.push(H2('1. 地質時代区分と岩石種区分（全10区分）'));
ch.push(P('分類の根幹は地質生成年代（Ma：百万年前）による時代区分と岩石種タイプの組み合わせである。Ⅰ類は先新第三紀（> 23 Ma）、Ⅱ類は新第三紀以降（< 23 Ma）。各大区分をa〜eの5サブタイプに細分する。'));
ch.push(SP());
ch.push(TBL(
  ['区分','年代範囲','代表岩石種','S目安','主R','全国地質帯・ダム工学上の意義'],
  [
    [{text:'■ Ⅰ類（古期地質） Prior to Neogene — 先新第三紀（> 23 Ma）',bold:true,fill:C.h3bg,color:C.h1bg},'','','','',{text:'',fill:C.h3bg}],
    ['Ⅰa\n変成岩帯','250〜50 Ma','結晶片岩・片麻岩・角閃岩','S1〜S2','R1・R2','日高変成帯・領家帯・三波川帯・飛騨帯。変成葉理沿い透水。断層・変質リスク大'],
    [{text:'Ⅰb\n花崗岩類',fill:C.r2},'110〜50 Ma','花崗岩・花崗閃緑岩・閃緑岩','S1','R1',{text:'領家花崗岩・中国地方・北上山地。最安定基礎岩盤。アーチダム適地',fill:C.r2}],
    ['Ⅰc\n砂岩泥岩互層','145〜23 Ma','タービダイト・チャート・頁岩','S2〜S3','R1','蝦夷層群・秩父系・三郡帯。全国最多基礎地質。層理傾斜が透水性支配'],
    [{text:'Ⅰd\n超苦鉄質岩',fill:C.r2},'150〜100 Ma','カンラン岩・蛇紋岩・斑糲岩','S1〜S4','R1〜R3',{text:'幌満・三郡・超丹波帯。蛇紋岩化・膨潤（R3）が設計最難課題',fill:C.r2}],
    ['Ⅰe\n古生界堆積岩','541〜252 Ma','石灰岩・チャート・変成堆積岩','S2〜S3','R1・R5','秋吉台・四国山地・渡島。石灰岩カルスト（R5）が遮水設計の核心'],
    [{text:'■ Ⅱ類（新期地質） Neogene–Quaternary — 新第三紀〜現在（< 23 Ma）',bold:true,fill:C.h3bg,color:C.h1bg},'','','','',{text:'',fill:C.h3bg}],
    ['Ⅱa\n溶結凝灰岩','12〜0.1 Ma','ウェルデッドタフ・火砕流堆積物','S1〜S3','R4','支笏・阿蘇・姶良等カルデラ起源。冷却亀裂（R4）が卓越した透水経路'],
    [{text:'Ⅱb\n火山岩（溶岩）',fill:C.r2},'23〜0.01 Ma','安山岩・玄武岩・デイサイト','S2〜S3','R4',{text:'大雪・富士・那須・雲仙等。溶岩スタック構造把握が設計の鍵',fill:C.r2}],
    ['Ⅱc\n火山砕屑岩','23〜1 Ma','凝灰岩・火山礫凝灰岩・集塊岩','S3〜S4','R4','東北・九州前弧盆。固結度変動±1コード'],
    [{text:'Ⅱd\n新第三系堆積岩',fill:C.r2},'23〜2.6 Ma','砂岩・泥岩・礫岩（半固結）','S3〜S4','R1',{text:'天北・石狩・内陸盆地。農業ダム標準地質。スレーキング注意',fill:C.r2}],
    ['Ⅱe\n未固結堆積物','2.6 Ma〜現在','河床礫層・段丘・沖積','S4〜S5','R6','全国の農業・灌漑ダム基礎。止水工設計が必須'],
  ],
  [900,1000,1800,700,700,5000]
));
ch.push(SP());
ch.push(P('※ 年代値はICS国際層序委員会2023年版準拠。複合地質（1ダムに複数区分）は「・」で連結表記する（例：Ⅰc・Ⅱe）。',{color:'666666',before:40,after:80}));

// 2. S・Rコード
ch.push(H2('2. 強度コード（S）・リスク指標（R）'));
ch.push(H3('2.1 強度コード（S）— ISRM一軸圧縮強度基準・5段階'));
ch.push(TBL(
  ['コード','区分名','一軸圧縮強度 qu','補助 Is50','全国代表岩石・対応区分'],
  [
    [{text:'S1',bold:true},'極硬岩','> 200 MN/m²','> 8 MPa','花崗岩・強変成岩・強溶結凝灰岩。Ⅰa・Ⅰb・Ⅱa（完全溶結）'],
    [{text:'S2',bold:true,fill:C.r2},'硬岩','100〜200','4〜8 MPa',{text:'堅硬砂岩・安山岩溶岩・花崗閃緑岩。Ⅰb・Ⅰc堅硬部・Ⅱb',fill:C.r2}],
    [{text:'S3',bold:true},'中硬岩','25〜100','1〜4 MPa','溶結凝灰岩（中程度）・一般砂岩・玄武岩。Ⅱa中溶結・Ⅱc・Ⅰc'],
    [{text:'S4',bold:true,fill:C.r2},'軟岩','5〜25','0.2〜1',{text:'軟質凝灰岩・泥岩・半固結礫岩。Ⅱc・Ⅱd・Ⅰd変質部',fill:C.r2}],
    [{text:'S5',bold:true},'極軟岩・土質','< 5 MN/m²','< 0.2','未固結礫層・砂・粘土。Ⅱe全般'],
    [{text:'S?',bold:true,fill:C.r2},'情報なし','—','—',{text:'試験データなし。設計・施工報告書参照要',fill:C.r2}],
  ],
  [700,900,1500,900,6100]
));
ch.push(SP());
ch.push(H3('2.2 リスク指標（R）— 工学的リスク種別・6種（複数付与可）'));
ch.push(TBL(
  ['コード','リスク区分','内容・全国の対象地質'],
  [
    [{text:'R1',bold:true},'断層・せん断帯','活断層・破砕帯・断層粘土。すべり面・透水路の主因。Ⅰa・Ⅰc・Ⅰd帯で顕著。中央構造線・フォッサマグナ等'],
    [{text:'R2',bold:true,fill:C.r2},'変質・変成リスク',{text:'熱水変質・接触変成・蛇紋岩化。局所的強度低下・膨潤。Ⅰa・Ⅰd。北海道ペーパーンダム等',fill:C.r2}],
    [{text:'R3',bold:true},'膨張性リスク','蛇紋岩（クリソタイル）・石膏・モンモリロナイト。経時的体積変化。Ⅰd全域'],
    [{text:'R4',bold:true,fill:C.r2},'冷却亀裂リスク',{text:'柱状節理・板状節理（溶結凝灰岩・溶岩）。高透水経路。Ⅱa・Ⅱb。支笏・阿蘇・浅間等カルデラ周辺',fill:C.r2}],
    [{text:'R5',bold:true},'溶解・空洞リスク','石灰岩カルスト空洞。グラウト流失・遮水不能。Ⅰe。秋吉・四国・沖縄・北海道渡島'],
    [{text:'R6',bold:true,fill:C.r2},'未固結層リスク',{text:'液状化・パイピング・内部侵食。フィルダム基礎の主要リスク。Ⅱe全般',fill:C.r2}],
  ],
  [700,1800,7600]
));
ch.push(PB());

// 3. W・Gコード
ch.push(H2('3. 透水性コード（W）・基礎処理コード（G）'));
ch.push(H3('3.1 透水性コード（W）— ルジオン値基準・5段階'));
ch.push(P('1 Lu ≈ 1.3×10⁻⁷ m/s（圧力1 MPa時）。W?は情報欠如の優先収集フラグ。'));
ch.push(SP());
ch.push(TBL(
  ['コード','区分名','ルジオン値','透水係数 k','透水機構','全国代表事例'],
  [
    [{text:'W1',bold:true},'極低透水','< 1 Lu','< 10⁻⁸ m/s','粒間浸透（事実上遮水）','変成岩緻密部・花崗岩深部。グラウト不要。Ⅰa・Ⅰb'],
    [{text:'W2',bold:true,fill:C.r2},'低透水','1〜5 Lu','10⁻⁸〜10⁻⁶ m/s','微細亀裂浸透',{text:'標準健全岩盤。1列カーテングラウト十分。Ⅰb・Ⅱb緻密部',fill:C.r2}],
    [{text:'W3',bold:true},'中透水','5〜30 Lu','10⁻⁶〜10⁻⁵ m/s','開口亀裂卓越','亀裂系が主経路。美生ダム実績。Ⅰc・Ⅱd'],
    [{text:'W4',bold:true,fill:C.r2},'高透水','30〜100 Lu','10⁻⁵〜10⁻³ m/s','大亀裂・断層・冷却亀裂',{text:'日新ダム実績（10¹ m/day）。Ⅱa柱状節理・断層破砕帯',fill:C.r2}],
    [{text:'W5',bold:true},'極高透水','> 100 Lu','> 10⁻³ m/s','未固結間隙・カルスト空洞','グラウト流失。止水矢板・遮水壁要。Ⅱe・Ⅰe石灰岩'],
    [{text:'W?',bold:true,fill:C.r2},'情報なし','—','—','—',{text:'試験記録なし。Ph.1で解消すべき優先課題',fill:C.r2}],
  ],
  [700,1000,1100,1500,1800,4000]
));
ch.push(SP());
ch.push(H3('3.2 基礎処理コード（G）— グラウチング難易度・4段階（ICOLD Bulletin 88準拠）'));
ch.push(TBL(
  ['コード','難易度','注入量目安','主な工法','対応W','全国代表事例'],
  [
    [{text:'G1',bold:true},'軽微','< 50 kg/m（セメント）','コンソリのみ','W1〜W2','健全硬岩。補強目的のみ。豊平峡・宮ヶ瀬アーチダム相当'],
    [{text:'G2',bold:true,fill:C.r2},'標準','50〜200 kg/m','カーテン1〜2列＋コンソリ','W2〜W3',{text:'全国中流ダム標準。砂岩泥岩・安山岩・新第三系',fill:C.r2}],
    [{text:'G3',bold:true},'高難度','200〜500 kg/m（多段反復）','多列カーテン＋高圧・超微粒子','W3〜W4','溶結凝灰岩柱状節理。日新・東郷・古梅・阿蘇系ダム典型'],
    [{text:'G4',bold:true,fill:C.r2},'特殊工法','> 500 kg/m または非セメント','止水矢板・遮水壁・化学グラウト','W4〜W5',{text:'石灰岩カルスト（Ⅰe）・未固結礫層（Ⅱe）・蛇紋岩大破砕帯',fill:C.r2}],
    [{text:'G?',bold:true},'情報なし','—','施工報告書参照要','—','グラウチング施工記録なし。優先確認対象'],
  ],
  [700,900,1800,2200,900,3600]
));
ch.push(SP());
ch.push(P('※ WコードとGコードは密接に対応するが同一ではない。亀裂の方向性・連続性・充填状態により同じWコードでもGコードが変わる（例：W4でも空洞に連通すればG4）。',{color:'666666'}));

// 4. W・G標準値一覧
ch.push(H2('4. 地質区分別 W・G 標準値'));
ch.push(P('各地質区分について透水性（W）・基礎処理難易度（G）の標準的な範囲を示す。全国ダムの初期判定に使用する。'));
ch.push(SP());
ch.push(TBL(
  ['区分','代表岩石','S','透水機構・主経路','W（標準）','G（標準）','全国工学上の要点'],
  [
    [{text:'Ⅰa',bold:true},'変成岩','S1〜2','変成葉理沿い浸透。断層帯でW4↑','W1〜2\n(断層帯W4)','G1〜2\n(断層帯G3)','断層走向・傾斜が透水性支配。局部高透水帯の事前探査重要'],
    [{text:'Ⅰb',bold:true,fill:C.r2},'花崗岩','S1','節理・方状節理が主経路。深部W1','W1〜2\n(風化帯W3)',{text:'G1〜2',fill:C.r2},{text:'理想的基礎。グラウト量最小。国内アーチダムの多くが該当',fill:C.r2}],
    [{text:'Ⅰc',bold:true},'砂岩泥岩互層','S2〜3','層理面・開口亀裂が主経路。砂岩W3・泥岩W1〜2','W2〜3','G2〜3','美生ダム実績：開口亀裂（砂岩卓越）→W3/G2。全国最多分布'],
    [{text:'Ⅰd',bold:true,fill:C.r2},'超苦鉄質岩・蛇紋岩','S1〜4','蛇紋岩化破砕帯W4。母岩W2。透水性経時変動大','W3〜4\n(破砕帯W4)',{text:'G3〜4',fill:C.r2},{text:'グラウト効果確認困難。化学グラウト検討要。経時変動に注意',fill:C.r2}],
    [{text:'Ⅰe',bold:true},'石灰岩・チャート','S2〜3','石灰岩カルストW5。チャートW1。二者混在で危険','W1(チャート)\n〜W5(石灰岩)','G2〜4','カルスト空洞の事前探査（ボーリング・物理探査）最重要'],
    [{text:'Ⅱa',bold:true,fill:C.r2},'溶結凝灰岩','S1〜3','柱状節理（冷却亀裂）卓越。鉛直・水平大亀裂','W3〜4',{text:'G3',fill:C.r2},{text:'日新実績：10¹ m/day≈W4。支笏・阿蘇・姶良系ダムに共通',fill:C.r2}],
    [{text:'Ⅱb',bold:true},'安山岩・玄武岩溶岩','S2〜3','流理・板状節理。完全溶岩体W2。冷却面W4','W2〜3\n(冷却面W4)','G2〜3','溶岩スタック構造把握が設計の鍵。大雪・富士・那須等'],
    [{text:'Ⅱc',bold:true,fill:C.r2},'凝灰岩・火山砕屑岩','S3〜4','粒間浸透＋亀裂の混合。固結度依存','W2〜3',{text:'G2〜3',fill:C.r2},{text:'固結度変動で±1コード。東北・九州前弧盆地帯',fill:C.r2}],
    [{text:'Ⅱd',bold:true},'砂岩・泥岩（新第三系）','S3〜4','砂岩層粒間浸透。層理面沿い透水','W2〜3','G2','スレーキング（R2）注意。天北・石狩・秋田・山形盆地'],
    [{text:'Ⅱe',bold:true,fill:C.r2},'未固結礫層・段丘','S4〜5','粒間透水卓越。礫層W5（k>10⁻³ m/s）','W4〜5',{text:'G3〜4',fill:C.r2},{text:'農業フィルダム基礎。止水矢板・遮水壁等補助工法必須',fill:C.r2}],
  ],
  [700,1300,600,2200,1000,900,3400]
));
ch.push(PB());

// 5. 次世代コード
ch.push(H2('5. 次世代分類コード（D・E・C・V・K）— 全国共通定義'));
ch.push(P('以下5コードは既存の国際的体系（DMR・RMR・Q-system・PWRI岩盤分類）に存在しない独創的な分類軸である。全国展開においても共通体系として適用する。'));
ch.push(SP());
ch.push(TBL(
  ['コード','名称','記号体系','定義・内容・全国ダムへの意義'],
  [
    [{text:'D',bold:true,color:C.h2bg},'経時劣化ポテンシャル','D1〜D5・D?','D1:変化極小（花崗岩・変成岩）　D2:緩慢溶解（砂岩泥岩）　D3:蛇紋岩化進行（Ⅰd）\nD4:凍結融解亀裂拡大（北日本特有・Ⅱa露出部）　D5:未固結層圧密・液状化（Ⅱe）\n100年スケールの劣化ポテンシャルを記号化。老朽化ダム優先再評価に直結'],
    [{text:'E',bold:true,color:C.h2bg,fill:C.r2},'地震応答増幅','E1〜E5・E?',{text:'E1:Vs>1500 m/s（増幅なし・岩盤）　E2:700〜1500（軽微）　E3:300〜700（中程度）\nE4:150〜300（強増幅・段丘礫）　E5:<150（液状化域）\n地質区分→Vs30の統計対応を体系化。耐震照査・老朽化ダム再評価と連動',fill:C.r2}],
    [{text:'C',bold:true,color:C.red},'貯水池・流域地質','C1〜C5・C?','C1:硬岩卓越・崩壊リスク低　C2:脆弱層あり・部分崩壊リスク　C3:カルスト・漏水リスク\nC4:活火山流域・土石流リスク（十勝・有珠・阿蘇・雲仙等）　C5:重金属・酸性水リスク（蛇紋岩・鉱山跡）\n★最独創的★ダム基礎から流域全体へ評価範囲を拡大する発想の転換'],
    [{text:'V',bold:true,color:C.h2bg,fill:C.r2},'気候変動脆弱性','V1〜V4・V?',{text:'V1:安定（影響小）　V2:融雪洪水ピーク増加→基礎水圧上昇\nV3:凍結融解サイクル変化→亀裂透水性変動（北日本特有）　V4:山岳永久凍土融解→斜面不安定\n21世紀末にかけての将来変化を記号化。長期管理計画の基盤',fill:C.r2}],
    [{text:'K',bold:true,color:C.green},'知識蓄積段階','K1〜K5・K?','K1:ボーリング＋試験＋施工記録完備　K2:設計報告書あり・施工記録一部欠落\nK3:地質図推定のみ・現地調査なし　K4:類推（位置・水系から推定）\nK5:AI推定（衛星・地質図の機械学習推論）← 将来実装\n「調査していない」ことを明示。情報収集の優先順位管理に直結'],
    [{text:'Q',bold:true,color:C.red},'耐震照査結果','Q1〜Q4・Q?','Q1:最新基準で肀震性能確認済み（地震動レベル1・2）　Q2:旧基準設計・照査済み（性能確認）\\nQ3:旧基準設計・未照査（照査実施が必要）　Q4:肀震性能不足・対策工検討中または実施中\\nQ?:肀震照査の実施状況不明\\nダム老机化対策・肀震補強計画の優先度管理に直結。Q3・Q4は最優先対策対象'],
  ],
  [700,1800,1600,6000]
));
ch.push(PB());

// 6. 完全記号体系
ch.push(H2('6. 完全記号体系とフェーズ別実装計画'));
ch.push(H3('6.1 全ブロック構造'));
ch.push(TBL(
  ['ブロック','記号形式','段階数','指標','Ph.1','Ph.2A','Ph.2B','Ph.2C'],
  [
    ['① 時代大区分','Ⅰ\\Ⅱ','2','地質生成年代','●','●','●','●'],
    ['② 岩石種サブ','a〜e','各5','岩石種タイプ','●','●','●','●'],
    ['③ 強度（S）','S1〜S5・S?','5+?','一軸圧縮強度','●','●','●','●'],
    ['④ リスク（R）','R1〜R6（複数）','6+','工学的リスク','●','●','●','●'],
    ['⑤ 透水性（W）','W1〜W5・W?','5+?','ルジオン値 Lu','—','●★','●','●'],
    ['⑥ 基礎処理（G）','G1〜G4・G?','4+?','グラウト難易度','—','●★','●','●'],
    ['⑦ 経時劣化（D）','D1〜D5・D?','5+?','劣化ポテンシャル','—','—','●★','●'],
    ['⑧ 地震応答（E）','E1〜E5・E?','5+?','Vs30・増幅率','—','—','●★','●'],
    ['⑨ 流域地質（C）','C1〜C5・C?','5+?','貯水池・流域','—','—','●★','●'],
    ['⑩ 気候変動（V）','V1〜V4・V?','4+?','将来変化','—','—','●★','●'],
    ['⑪ 知識段階（K）','K1〜K5・K?','5+?','情報充足度','—','—','—','●★'],
    ['⑫ 耐震照査結果（Q）','Q1〜Q4・Q?','4+?','照査水準','—','—','●★','●'],
  ],
  [1600,1500,900,1800,700,700,700,700]
));
ch.push(SP());
ch.push(H3('6.2 完全記号の構成と読み方'));
ch.push(P('完全記号の構造：  地質区分（S/R・W/G・D/E・C・V・K\\\\Q）'));
ch.push(P('括弧内は「力学ブロック（S/R）」・「水理ブロック（W/G）」・「将来ブロック（D/E）」・「流域（C）」・「気候（V）」・「知識（K）」・「耐震（Q）」の順に「・」で区切る。'));
ch.push(SP());
ch.push(TBL(
  ['完全記号（例）','信頼度','読み方・地質工学的意味'],
  [
    [{text:'Ⅱa（S1/R4・W4/G3）・Ⅱe（S5/R6・W5/G4）',bold:true},'A','溶結凝灰岩（極硬・冷却亀裂R4・高透水W4・多段グラウトG3）＋河床礫層（極軟・未固結R6・極高透水W5・特殊工法G4）。日新ダム相当'],
    [{text:'Ⅱa（S3/R4・W4/G3・D4/E2・C1・V3・K2\\Q3）',bold:true,fill:C.r2},'A',{text:'将来完全版。溶結凝灰岩の北海道型フル記号例。D4=凍結融解亀裂拡大・E2=軽微増幅・C1=硬岩流域・V3=凍結融解変動・K2=設計報告書あり・Q3=旧基準設計・未照査',fill:C.r2}],
    [{text:'Ⅰb（S1/R1・W1/G1）',bold:true},'B','花崗岩（極硬岩・極低透水W1・グラウト軽微G1）。宮ヶ瀬・豊平峡アーチダム相当。理想的基礎'],
    [{text:'Ⅰc（S2/R1・W?/G?）・Ⅱe（S5/R6・W5/G4）',bold:true,fill:C.r2},'C',{text:'砂岩泥岩＋未固結（水理情報欠如のW?/G?あり）。W?/G?は調査優先フラグ',fill:C.r2}],
  ],
  [3800,700,5600]
));
ch.push(PB());

// 7. 信頼度
ch.push(H2('7. 信頼度評価体系（全国共通・4段階）'));
ch.push(P('全国展開にあたり、北海道の3段階（A・B・C）を4段階（A・B・C・D）に拡張する。'));
ch.push(SP());
ch.push(TBL(
  ['ランク','定義','情報源','全国目標比率'],
  [
    [{text:'A',bold:true,color:C.green},'透水性・基礎処理実績データ確認済み','設計報告書・ルジオン試験記録・施工記録','目標10%以上'],
    [{text:'B',bold:true,color:C.h2bg,fill:C.r2},'ダム固有の地質記述確認済み',{text:'工事誌・専門文献・propylite.work等',fill:C.r2},{text:'目標30%以上',fill:C.r2}],
    [{text:'C',bold:true},'GeoNAVI地質図＋ダム型式から推定','GeoNAVI API・国土数値情報','初期分析の主体'],
    [{text:'D',bold:true,color:C.red,fill:C.r2},'位置情報のみ・地質取得不可',{text:'（データ不足）',fill:C.r2},{text:'最小化目標',fill:C.r2}],
    [{ text:'?',bold:true},'当該コードの判定情報が存在しない','—','W?/G?等の情報欠如フラグ'],
  ],
  [700,2800,3500,3100]
));
ch.push(PB());

// ═══ 第Ⅱ部 全国分析基盤 ══════════════════════════════════════════
ch.push(H1('第Ⅱ部　全国分析の基盤整備'));

// 8. GeoNAVI
ch.push(H2('8. 産総研シームレス地質図 Web API — 全国地質情報取得の基盤'));
ch.push(P('産総研「20万分の1日本シームレス地質図V2」Web API（ver.1.3.1）は、ダムの緯度・経度を指定して地質区分・岩相名・地質時代をJSON形式で自動取得できる。本プロジェクトでは全国ダムの初期地質判定（信頼度C付与）の主要手段として採用する。'));
ch.push(SP());
ch.push(NOTE('GeoNAVI Web API 仕様',[
  'エンドポイント：https://gbank.gsj.jp/seamless/v2/api/1.2/',
  '凡例取得：GET /legend?lang=ja → 全地質区分コードと岩相名称をJSONで取得',
  '地質情報取得：GET /query?lat={緯度}&lng={経度}&datum=WGS84 → 指定点の地質区分・岩相・時代をJSON取得',
  'typeパラメータ：level4（簡略版・凡例数約400）、level8（詳細版）が指定可能',
  '利用条件：オープンデータ（CC BY準拠）、APIキー不要、商用利用可',
]));
ch.push(SP());
ch.push(H3('8.1 GeoNAVI取得情報から分類記号への変換ルール'));
ch.push(P('GeoNAVI返却値（岩相名・地質時代）を本体系（Ⅰ\\Ⅱ・a〜e・S/R・W/G・K\\Q）に変換する対応表（約400行）を整備する。これが全国共通手法の核心部分である。'));
ch.push(SP());
ch.push(TBL(
  ['GeoNAVI岩相分類（代表例）','時代大区分','岩石種サブ','S強度目安','主Rコード'],
  [
    ['花崗岩・花崗閃緑岩・閃緑岩','Ⅰ','b（花崗岩類）','S1','R1'],
    ['緑色岩・輝緑凝灰岩・枕状溶岩（付加体）',{text:'Ⅰ',fill:C.r2},'b（緑色岩）','S2〜S3',{text:'R1・R3',fill:C.r2}],
    ['砂岩・泥岩・頁岩（中生代以前）','Ⅰ','c（堆積岩）','S2〜S4','R1'],
    ['変成岩・片麻岩・結晶片岩',{text:'Ⅰ',fill:C.r2},'a（変成岩）','S1〜S2',{text:'R1・R2',fill:C.r2}],
    ['蛇紋岩・カンラン岩・斑糲岩','Ⅰ','d（超苦鉄質）','S1〜S4','R1〜R3'],
    ['石灰岩・チャート（古生代）',{text:'Ⅰ',fill:C.r2},'e（古生界）','S2〜S3',{text:'R1・R5',fill:C.r2}],
    ['安山岩・玄武岩溶岩（新第三紀〜）','Ⅱ','b（火山岩）','S2〜S4','R1・R4'],
    ['溶結凝灰岩（第四紀）',{text:'Ⅱ',fill:C.r2},'a（溶結凝灰岩）','S1〜S3',{text:'R4',fill:C.r2}],
    ['砂岩・泥岩（新第三紀〜）','Ⅱ','d（新期堆積岩）','S2〜S3','R1・R2'],
    ['礫岩・砂岩（鮮新世〜更新世）',{text:'Ⅱ',fill:C.r2},'d〜e（半固結〜未固結）','S2〜S4',{text:'R1・R2',fill:C.r2}],
    ['沖積層・段丘礫層・火砕流堆積物（未固結）','Ⅱ','e（未固結）','S4〜S5','R6'],
  ],
  [3200,1200,1800,1300,2600]
));

ch.push(H2('9. 全国ダム位置情報の整備'));
ch.push(P('GeoNAVI APIを呼び出す入力として、各ダムの緯度・経度情報を含むマスターリストを整備する。'));
ch.push(SP());
ch.push(TBL(
  ['データソース','収録数目安','位置情報','取得方法','評価'],
  [
    [{text:'国土数値情報「ダムデータ W01」（国土交通省）',fill:C.good},'約2,700基','緯度経度あり（GML/SHP）','無料ダウンロード',{text:'◎ 最優先・第一候補',bold:true,color:C.green,fill:C.good}],
    ['ダム便覧（日本ダム協会）','約3,000基','住所のみ','Web収集＋ジオコーディング','○ 補完用'],
    [{text:'DamMaps（dammaps.jp）',fill:C.r2},'約2,500基','地図表示あり','スクレイピング',{text:'○ 補完用',fill:C.r2}],
    ['農林水産省 農業ダム施工記録','農業ダムのみ','一部あり','各農政局照会','△ 農業ダム専用'],
    [{text:'Google/地理院地図 手動補完',fill:C.r2},'個別','高精度','手動',{text:'△ 最終補完手段',fill:C.r2}],
  ],
  [3000,1200,1200,1800,3000]
));
ch.push(PB());

// ═══ 第Ⅲ部 フェーズ別作業計画 ══════════════════════════════════
ch.push(H1('第Ⅲ部　フェーズ別作業計画（全国共通）'));
ch.push(P('北海道（188基）を先行パイロットとして全フェーズを実施し、手法を確定した上で全国展開を行う。北海道の各フェーズ完了をもって全国同一フェーズを開始する。'));
ch.push(SP());
ch.push(TBL(
  ['Ph.','フェーズ名','期間目安','北海道','全国','主な成果物'],
  [
    [{text:'Ph.0',bold:true,fill:C.h3bg},'全国分析手法の確立','〜1ヶ月','設計','設計','GeoNAVI変換表・全国共通仕様書'],
    ['Ph.1','ダム位置情報の整備','〜1ヶ月','188基','全国ダム','全国マスターCSV（緯度経度付き）'],
    [{text:'Ph.2',fill:C.r2},'記号体系の確定','〜1ヶ月','検証','設計',{text:'変換表確定版・国際基準対応表',fill:C.r2}],
    ['Ph.3','GeoNAVI基礎分析','〜1ヶ月','全188基','全選定ダム','DB（GeoNAVIベース・信頼度C初期値）'],
    [{text:'Ph.4',fill:C.r2},'一次情報修正（文献・Web）','〜2ヶ月','全188基','全選定ダム',{text:'Wordファイル・propylite.work・W/G追加',fill:C.r2}],
    ['Ph.5','二次情報修正（人的情報）','継続的','全188基','全選定ダム','設計報告書・ルジオン記録追加'],
    [{text:'Ph.6',fill:C.r2},'結果考察（北海道）','〜1ヶ月','全188基','—',{text:'北海道ダム地質考察レポート',fill:C.r2}],
    ['Ph.7','全国対象ダムの選定','〜1ヶ月','—','全国','選定基準確定・対象ダムリスト'],
    [{text:'Ph.8',fill:C.r2},'全国公開情報収集','継続的','—','全選定',{text:'信頼度C→B（公開情報範囲）',fill:C.r2}],
    ['Ph.9','全国人的情報修正','継続的','—','全選定','信頼度向上版DB'],
    [{text:'Ph.10',fill:C.r2},'最終考察・成果物','〜2ヶ月','—','全国',{text:'最終報告書・GIS・オープンデータ',fill:C.r2}],
  ],
  [700,2400,1000,900,900,4200]
));
ch.push(PB());

// Ph.0〜2
ch.push(H2('Ph.0　全国分析手法の確立'));
ch.push(BUL('GeoNAVI API呼び出しスクリプト（Python）の作成とテスト（北海道188基で検証）。'));
ch.push(BUL('GeoNAVI岩相コード（約400種）全てについて本体系への変換規則を完成させる。'));
ch.push(BUL('変換規則の品質検証：既知の信頼度AのダムでGeoNAVI取得値と実測値を比較。'));
ch.push(BUL('DMR（Romana 2003）・RMR89・PWRI岩盤分類との対応表を整備し国際互換性を確保。'));

ch.push(H2('Ph.1　ダム位置情報の整備'));
ch.push(NUM('国土数値情報「ダムデータ W01」（GML/SHP）をダウンロード。'));
ch.push(NUM('GMLまたはSHPから：ダム名・河川名・都道府県・緯度・経度・型式・管理者・堤高・完成年を抽出しCSV化。'));
ch.push(NUM('ダム便覧・DamMapsと名称照合し、国土数値情報に収録されていないダムを補完。'));
ch.push(NUM('農業ダム（農水省管轄）は農政局公開資料で位置情報を補完。'));
ch.push(NUM('全国マスターCSV作成：「ダムID・ダム名・河川名・都道府県・緯度・経度・管理者区分・型式・堤高・完成年」10列。'));
ch.push(SP());
ch.push(NOTE('Claudeが実施できる範囲',[
  '国土数値情報GMLファイルをアップロードいただければPython/Pandasで自動変換・CSV作成が可能。',
  'ダム名→緯度経度のジオコーディングスクリプト（Nominatim API経由）の作成が可能。',
  '全国マスターCSVのGeoNAVI APIバッチ処理（500基を30分程度で自動取得）が可能。',
]));

ch.push(H2('Ph.2　記号体系の確定'));
ch.push(BUL('北海道先行分析の結果をレビューし、変換規則の誤り・不一致を修正。'));
ch.push(BUL('次世代コード（D/E/C/V/K）の全国適用基準を最終確定。'));
ch.push(BUL('日本固有の地質条件（付加体・蛇紋岩・泥炭性堆積物・カルデラ起源火砕流）への特例規則を定義。'));
ch.push(BUL('成果物：「GeoNAVI岩相コード→分類記号変換表」確定版・「国際基準対応表」・「分類体系確定版仕様書」。'));
ch.push(PB());

// Ph.3〜5
ch.push(H2('Ph.3　GeoNAVI基礎分析（北海道：全188基・全国：全選定ダム）'));
ch.push(P('ダム位置情報（緯度・経度）をGeoNAVI APIに送信し、岩相コード・岩相名・地質時代を取得。Ph.2の変換表に照らしてⅠ\\Ⅱ・サブ・S・Rコードを自動付与（W/G/K/Qは後フェーズで追加）。信頼度は全てC（K4：位置・水系から類推）。'));
ch.push(SP());
ch.push(NOTE('Ph.3の重要な役割',[
  '北海道188基のPh.3結果は「GeoNAVIのみによる分析ベースライン」として独立シートに記録する。',
  '後のPh.4（文献修正）・Ph.5（人的情報修正）との比較により、各情報層の貢献度を定量化できる。',
  '全国展開でもこの「3層構造（GeoNAVI→文献→人的情報）」を共通フォーマットとして採用する。',
]));

ch.push(H2('Ph.4　一次情報修正（文献・Webサイト情報）— 北海道：全188基・全国：全選定ダム'));
ch.push(H3('北海道での使用資料（既存資料・完成済み）'));
ch.push(TBL(
  ['資料名','収録ダム数','情報の質','主な追加情報'],
  [
    [{text:'propylite.work（道央・道南・道北・道東編）',fill:C.good},'約60〜80基','高','岩種・構造・変質・透水性・施工記録'],
    ['国営農業用ダム地質雑感（道東編）','7基','非常に高','詳細地質・ルジオン値・グラウト計画'],
    [{text:'国営農業用ダム地質雑感（道北編）',fill:C.r2},'14基','非常に高','詳細地質・変形係数・地すべり情報'],
    ['国営農業用ダム地質雑感（道南編）','5基','高','詳細地質・透水係数・水理地質'],
  ],
  [3500,1000,1000,4600]
));
ch.push(SP());
ch.push(H3('全国での使用資料（収集対象）'));
ch.push(BUL('ダム便覧 各ダムページの地質記述（Webフェッチ・テキスト解析）。'));
ch.push(BUL('産総研 地質図幅説明書（各ダム周辺20万分の1図幅の岩石記載）。'));
ch.push(BUL('発電事業者・水道事業体の環境アセスメント公開資料（ルジオン試験値等）。'));
ch.push(BUL('各地方整備局・道府県公開の施工報告書（公開分）。'));
ch.push(BUL('各地の propylite.work 相当の専門Webサイト。'));

ch.push(H2('Ph.5　二次情報修正（人的情報）— 全ダム対象・継続的'));
ch.push(H3('収集すべき人的情報の種類'));
ch.push(BUL('ダム設計報告書（地質調査編）：地質縦断図・ボーリング柱状図・室内試験結果。'));
ch.push(BUL('施工記録（グラウチング）：注入孔配置・ルジオン試験値・注入量・セメント/水比の記録。'));
ch.push(BUL('ダム完成後検査記録：漏水量観測・基礎ひずみ・揚圧力測定データ。'));
ch.push(BUL('老朽化点検記録：基礎岩盤変状・変質進行・溶解空洞の有無。'));
ch.push(BUL('地方技術者の口頭伝承：設計段階では記録されなかった地質上の問題・対応措置。'));
ch.push(SP());
ch.push(H3('優先収集対象（全国）'));
ch.push(TBL(
  ['対象地質区分','全国概数','推定W/G','目標信頼度','確認すべき情報'],
  [
    [{text:'溶結凝灰岩帯（Ⅱa）',fill:C.risk},'約60基','W4/G3推定','C→B→A','ルジオン試験値・グラウチング量（柱状節理部）'],
    ['石灰岩・古生界（Ⅰe）','約80基','W1〜W5、G2〜G4','C→B','カルスト空洞探査記録・グラウト充填記録'],
    [{text:'超苦鉄質岩・蛇紋岩（Ⅰd）',fill:C.r2},'約20基','W4/G4推定','C→B',{text:'破砕帯分布・蛇紋岩化程度・経時変動',fill:C.r2}],
    ['1920年代以前竣工','約100基（全国）','多様','C→B','建設当時の地質調査記録の発掘・劣化コードD付与'],
    [{text:'農業フィルダム（未固結基礎・Ⅱe）',fill:C.r2},'約600基（全国）','W4〜W5、G3〜G4','C→B',{text:'農林水産省農業水産部施工記録（別ルート入手要）',fill:C.r2}],
    ['日高・三波川・領家変成帯（Ⅰa）','約80基','W2/G2、断層帯W4/G3','B→A','断層破砕帯の調査記録・活断層との交差関係'],
  ],
  [2500,1000,1400,1000,4200]
));
ch.push(PB());

// Ph.6
ch.push(H2('Ph.6　結果考察（北海道188基）'));
ch.push(P('Ph.3〜Ph.5で構築した北海道DBを多角的に考察する。以下の考察テーマは全国分析（Ph.10）の雛形となる。'));
ch.push(SP());
ch.push(H3('考察①　地質帯別分布と工学的特性の比較'));
ch.push(BUL('北海道7地質帯（日高・蝦夷・天北・大雪火山・支笏-洞爺・道東白亜系・渡島古生界）別にS・R・W・Gコード分布を統計的に比較。'));
ch.push(BUL('「付加体系（Ⅰb・Ⅰc）ダムは透水性W4〜W5が多い傾向があるか」等の仮説を検証。'));
ch.push(H3('考察②　ダム型式と地質区分の相関'));
ch.push(BUL('コンクリートダム（重力式・アーチ）に採用されやすい地質区分 vs フィルダムに採用されやすい地質区分の分析。'));
ch.push(BUL('Ⅱe（未固結）地盤に建設されたダムの特殊基礎処理技術（地下連続壁・ブランケット等）の分布可視化。'));
ch.push(H3('考察③　建設年代と技術深化の関係'));
ch.push(BUL('建設年代別（1950年代〜2010年代）に地質区分・W/Gコードの分布を分析し、困難地盤への挑戦がいつ進んだかを検証。'));
ch.push(BUL('知識段階コード（K1〜K5）の導入による「当時の技術水準と現在の知識ギャップ」の可視化。'));
ch.push(H3('考察④　老朽化リスクポテンシャル（D・E・V コードの初適用）'));
ch.push(BUL('北海道188基に対してD・E・Vコードを試験的に付与し、「Rコード×Dコード×建設年代」マトリクスによる優先再評価ダムリスト（上位30基）を作成。'));
ch.push(BUL('変質系リスク（R3：熱水変質・R2：スレーキング）と老朽化の相関。D4（凍結融解）は北海道特有の重要指標。'));
ch.push(H3('考察⑤　全国展開への示唆'));
ch.push(BUL('GeoNAVI基礎分析とWordファイル修正の乖離率から、全国での信頼度C→Bへの修正作業量を推計。'));
ch.push(BUL('北海道の地質多様性が日本全国の地質区分の試験場として機能した点の評価。'));
ch.push(PB());

// Ph.7〜10
ch.push(H2('Ph.7〜Ph.10　全国展開'));
ch.push(H3('Ph.7　全国対象ダムの選定'));
ch.push(P('日本全国には約3,000基のダムが存在する（国土数値情報2014年版）。品質管理・考察の深度の観点から、まず約500〜800基の「第一選定群」を対象とする。'));
ch.push(SP());
ch.push(TBL(
  ['選定基準','優先度','全国概数','理由'],
  [
    [{text:'堤高15m以上のコンクリートダム（全管理者）',fill:C.good},{text:'★★★',bold:true,fill:C.good},'約400基','設計資料が整備。地質調査記録が存在。'],
    ['堤高50m以上のフィルダム（全管理者）',{text:'★★★',bold:true},'約150基','大規模基礎処理が行われている。'],
    [{text:'老朽化対策事業中のダム（国交省公表）',fill:C.r2},{text:'★★★',bold:true,fill:C.r2},'約100基',{text:'再調査記録が新たに存在する可能性大。',fill:C.r2}],
    ['特殊地質ダム（石灰岩・蛇紋岩・未固結等）',{text:'★★',bold:true},'約50基','工学的・学術的価値が高い。'],
    [{text:'アーチダムおよびバットレスダム（全て）',fill:C.r2},{text:'★★',bold:true,fill:C.r2},'約80基',{text:'岩盤強度要件が最も厳しい型式。',fill:C.r2}],
    ['1960年代以前竣工のダム（全型式）',{text:'★',bold:true},'約100基','老朽化と地質条件の複合リスク分析に重要。'],
  ],
  [3200,700,900,5300]
));
ch.push(SP());
ch.push(H3('Ph.8　全国：公開情報収集（Claudeの実施可能範囲）'));
ch.push(TBL(
  ['作業内容','手法','期待効果','制約'],
  [
    [{text:'ダム便覧Webページ地質記述収集',fill:C.good},'Webフェッチ＋テキスト解析','信頼度C→B（約200〜300基）','地質情報の記述量が少ない場合あり'],
    ['産総研地質図幅説明書のテキスト検索','地質図カタログ参照','周辺地質の詳細情報取得','図幅単位→個別ダムの対応付けが必要'],
    [{text:'発電事業者・水道事業体の公開資料',fill:C.r2},'公開PDF解析','ルジオン値等の実測データ取得',{text:'検索・取得に時間コストがかかる',fill:C.r2}],
    ['地方専門Webサイト収集','Web検索＋収集','各地方の類似情報発見','地方ごとに情報密度の差が大きい'],
  ],
  [2800,1800,2400,3100]
));
ch.push(SP());
ch.push(H3('Ph.9　全国：人的情報修正'));
ch.push(BUL('各地方整備局・北海道開発局ダム管理所への情報提供依頼（公文書開示請求含む）。'));
ch.push(BUL('農林水産省・道県農政局への国営農業ダム施工記録の照会。'));
ch.push(BUL('電力事業者（北電・東電・関電・九電等）の環境・技術広報資料の収集。'));
ch.push(BUL('ダム関連学会（日本大ダム会議・土木学会等）の論文・報告書収集。'));
ch.push(BUL('地方の建設コンサルタント・地質調査会社への協力要請。'));
ch.push(SP());
ch.push(H3('Ph.10　最終考察と成果物（全国版）'));
ch.push(P('北海道Ph.6の考察テーマを全国に展開し、以下の5テーマで最終考察を行う。'));
ch.push(BUL('考察①　日本の地質帯とダム基礎岩盤の全国分布：変動帯としての日本列島の地質多様性を体系化。「地質帯別リスク分布地図」作成。'));
ch.push(BUL('考察②　管理者区分・地域別の情報密度の格差：国交省直轄ダム vs 地方管理ダムの情報密度を定量化し、優先整備地域を特定。'));
ch.push(BUL('考察③　老朽化リスクの全国分布：Rコード×Dコード×建設年代×堤高の多次元マトリクスによるリスクランキング。補修・再調査優先度を全国ダムに適用。'));
ch.push(BUL('考察④　C（流域地質）・V（気候変動）の全国適用：活火山流域（C4）・酸性水リスク（C5）・凍結融解地帯（V3）の全国分布を可視化。'));
ch.push(BUL('考察⑤　AIとの協働（K5の実装）：信頼度Aのダムデータ＋GeoNAVI地質図を教師データとした機械学習モデル構築。InSAR・衛星マルチスペクトルとの統合。'));
ch.push(PB());

// ═══ 第Ⅳ部 成果物・参照資料 ═════════════════════════════════════
ch.push(H1('第Ⅳ部　成果物・参照資料'));
ch.push(H2('10. 最終成果物一覧'));
ch.push(TBL(
  ['成果物名','形式','内容・対象'],
  [
    [{text:'全国ダム地質分類DB',fill:C.good},'Excel / SQLite','全選定ダムの完全記号・信頼度・K段階・判定根拠'],
    ['北海道ダム地質区分DB（確定版）','Excel','188基の完全記号（S/R/W/G/D/E/C/V/K）'],
    [{text:'全国ダム地質区分KMZ',fill:C.r2},'KMZ（Google Earth）','地質区分別色分けマーカー・ポップアップ情報（全国）'],
    ['全国ダム地質分布GIS','QGIS プロジェクト','シームレス地質図レイヤー＋ダムポイントデータ'],
    [{text:'老朽化リスクランキング',fill:C.r2},'Excel','Rコード×Dコード×建設年代による優先度順位表'],
    ['地域別地質考察レポート（北海道版）','docx','北海道7地質帯別・5考察テーマ'],
    [{text:'地域別地質考察レポート（全国版）',fill:C.r2},'docx','全国8地方ブロック別の地質特性まとめ'],
    ['国際基準対応表','Excel','DMR/RMR/Q-system/PWRI vs 本体系変換表'],
    [{text:'各ダム個別地質カード（全国版）',fill:C.r2},'docx / HTML','標準フォーマット個別カード（選定ダム全基）'],
    ['研修用テキスト（ダム地質入門）','docx / PDF','ダム建設未経験技術者向け基礎知識'],
    [{text:'最終報告書',fill:C.r2},'docx / PDF','全フェーズの成果・考察・提言をまとめた総括文書'],
    ['オープンデータDB','CSV / JSON','公開可能形式・ライセンス確認済み'],
  ],
  [3000,1800,5300]
));
ch.push(PB());

ch.push(H2('11. 参照資料・データ収集戦略'));
ch.push(H3('11.1 一次資料（無料・オープンデータ）'));
ch.push(BUL('産総研 シームレス地質図V2 Web API（https://gbank.gsj.jp/seamless/v2/api/）：無料・CC BY。全国地質情報の基盤。'));
ch.push(BUL('国土数値情報「ダムデータ W01」（国土交通省）：GML・SHPで全国ダム位置情報を無料配布。'));
ch.push(BUL('ダム便覧（日本ダム協会 http://damnet.or.jp/）：ダム名・型式・堤高・管理者情報。'));
ch.push(BUL('propylite.work（道央・道南・道北・道東編）：北海道ダム地質の最重要一次資料（管理者照会要）。'));
ch.push(BUL('国営農業用ダム地質雑感（道東・道北・道南編）：北海道25基の詳細地質情報（提供済み）。'));
ch.push(H3('11.2 国際基準・参考文献'));
ch.push(BUL('Romana, M.（2003）DMR (Dam Mass Rating). ISRM 10th Congress — 最も近い国際先行体系'));
ch.push(BUL('Bieniawski, Z.T.（1989）Engineering Rock Mass Classifications — RMR89の原典'));
ch.push(BUL('ICOLD Bulletin 88 Rock Foundations for Dams・Bulletin 111 — 国際標準グラウチング指針'));
ch.push(BUL('Fell et al.（2015）Geotechnical Engineering of Dams — フィルダム基礎標準テキスト'));
ch.push(BUL('ICS International Chronostratigraphic Chart 2023 — 地質年代値の基準'));
ch.push(BUL('Weaver & Bruce（2007）Dam Foundation Grouting, ASCE — グラウチング技術の集大成'));
ch.push(PB());

// ═══ 第Ⅴ部 将来展望 ══════════════════════════════════════════════
ch.push(H1('第Ⅴ部　将来展望'));
ch.push(H2('12. AIとの協働（K5コードの実装）'));
ch.push(P('知識段階コード（K）のK5は「AIによる推定」を明示するフラグである。将来的には以下の技術統合を目指す。'));
ch.push(BUL('教師データ：信頼度Aのダムデータ＋GeoNAVI地質図を用いた機械学習モデルの構築。'));
ch.push(BUL('入力：ダム位置（緯度経度）・標高・水系・型式・建設年代 → 出力：地質区分記号・S/R/W/G推定コードと確信度スコア。'));
ch.push(BUL('InSAR（合成開口レーダー干渉）データとの統合：地表変位計測×地質リスク（R1・R3）の重ね合わせ。'));
ch.push(BUL('衛星マルチスペクトル画像による変質帯（R2）の自動検出。'));
ch.push(P('この「AIと人間の共同調査プラットフォーム」が実現すれば、K5ラベルが付いたAI推定コードを現地調査・報告書参照で順次K1〜K4に更新していく継続的な精度向上サイクルが確立される。'));

ch.push(H2('13. 他地域・国際展開の可能性'));
ch.push(BUL('東北・中部・九州地方への展開：火山性地質（Ⅱa・Ⅱb）・変成帯（Ⅰa）が類似する地域への適用。'));
ch.push(BUL('ICOLD（国際大ダム会議）での発表：日本から提案する新しいダム地質分類体系として国際発信。'));
ch.push(BUL('アジア地域への展開：東南アジアの新第三系（Ⅱd）・火山帯（Ⅱa・Ⅱb）への適用可能性。'));
ch.push(BUL('気候変動脆弱性コード（V）は全世界のダムに適用可能な普遍的フレームワーク。'));

ch.push(H2('14. 社会的価値'));
ch.push(BUL('老朽化ダムの科学的優先評価：DコードとVコード×建設年代のマトリクスにより、限られた予算で最もリスクの高いダムから順に再評価できる。'));
ch.push(BUL('ダム管理技術の伝承：文書化された知識基盤として、ダム建設未経験技術者への技術継承を促進する。'));
ch.push(BUL('防災・減災への貢献：洪水・地震・火山噴火等の自然災害時に、地質リスクの高いダムを即座に特定できる体制の整備。'));
ch.push(BUL('国際発信：地質年代軸・流域地質評価・AI統合という国際的に先例のない体系の確立。'));
ch.push(SP(2));
ch.push(HR());
ch.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'以　上',size:24,font:'MS Mincho',color:'444444'})],spacing:{before:200,after:100}}));
ch.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'2026年3月　日本全国ダム地質分類プロジェクト　作業計画書（改訂版2）',size:20,font:'MS Mincho',color:'888888'})],spacing:{before:60,after:0}}));

// ─── 文書組み立て ─────────────────────────────────────────────────
const doc = new Document({
  numbering,
  styles:{
    default:{document:{run:{font:'MS Mincho',size:22}}},
    paragraphStyles:[
      {id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:32,bold:true,font:'Arial'},paragraph:{spacing:{before:360,after:180},outlineLevel:0}},
      {id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:28,bold:true,font:'Arial'},paragraph:{spacing:{before:280,after:140},outlineLevel:1}},
      {id:'Heading3',name:'Heading 3',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:26,bold:true,font:'Arial'},paragraph:{spacing:{before:200,after:100},outlineLevel:2}},
      {id:'Heading4',name:'Heading 4',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:24,bold:true,font:'Arial'},paragraph:{spacing:{before:160,after:80},outlineLevel:3}},
    ]
  },
  sections:[{
    properties:{page:{size:{width:A4W,height:A4H},margin:{top:MAR,right:MAR,bottom:MAR+400,left:MAR}}},
    headers:{default:new Header({children:[new Paragraph({children:[new TextRun({text:'日本全国ダム地質分類　体系的分析・全国展開　作業計画書（改訂版2・2026年3月）',size:18,color:'888888',font:'MS Mincho'})],border:{bottom:{style:BorderStyle.SINGLE,size:4,color:C.bdr}},spacing:{after:100}})]})},
    footers:{default:new Footer({children:[new Paragraph({children:[new TextRun({text:'propylite.work  ／  2026年3月改訂       ',size:18,color:'888888',font:'MS Mincho'}),new TextRun({children:[PageNumber.CURRENT],size:18,color:'888888'}),new TextRun({text:' / ',size:18,color:'888888'}),new TextRun({children:[PageNumber.TOTAL_PAGES],size:18,color:'888888'})],alignment:AlignmentType.RIGHT,border:{top:{style:BorderStyle.SINGLE,size:4,color:C.bdr}},spacing:{before:100}})]})},
    children:ch,
  }]
});

Packer.toBuffer(doc).then(buf=>{
  fs.writeFileSync('/home/claude/全国ダム地質分類_作業計画書_改訂版2.docx',buf);
  console.log('DONE');
});
