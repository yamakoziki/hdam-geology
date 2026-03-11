// ======================================================
// 北海道ダム地質分類 全体計画書（フルプラン）メイン
// ======================================================
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, PageNumber, PageBreak, Footer, Header
} = require('docx');
const fs = require('fs');
const U = require('./fullplan_p1.js');
const {C,B,cell,row,secRow,H1,H2,H3,H4,P,NOTE,SP,PB_BREAK,DIV,PHASE_LABEL,BUL,BUL2} = U;

// ═══════════════════════════════════════════════════════════════
// 表: ロードマップ概要
// ═══════════════════════════════════════════════════════════════
function tblRoadmap() {
  const W=[900,1100,1500,1800,3200]; // 8500
  const phases = [
    ["Phase 1","〜2ヶ月","情報収集","設計・施工報告書・GeoNAVI参照","全188基のS/R/W/G信頼度A〜B化。propylite.work照合","ph1"],
    ["Phase 2","〜2ヶ月","体系確定","試行基レビュー・コード修正・国際基準対応表","記号体系確定版。DMR・RMR対応表。D/E/C/V/K各コード定義確定","ph2"],
    ["Phase 3","〜6ヶ月","全基判定","水系別一括判定・地質区分DB更新","全188基の完全記号（S/R/W/G/D/E/C）確定。KMZ更新","ph3"],
    ["Phase 4","〜3ヶ月","説明文作成","概括説明文・各ダム個別地質カード","各ダム地質カード188枚。地質帯別解説文","ph4"],
    ["Phase 5","〜2ヶ月","成果物取りまとめ","報告書・DB・GIS・教材整備","最終報告書・オープンデータDB・地質区分GIS・教育用資料","ph5"],
  ];
  return new Table({
    width:{size:8500,type:WidthType.DXA}, columnWidths:W,
    rows:[
      row([["フェーズ",W[0],{center:true}],["期間",W[1],{center:true}],
           ["テーマ",W[2],{center:true}],["主な作業",W[3],{center:true}],
           ["主な成果物",W[4],{center:true}]], true, 38),
      ...phases.map(([ph,dur,theme,work,out,colorKey])=>{
        const fill = {ph1:"D6EAF8",ph2:"FDEBD0",ph3:"D5F5E3",ph4:"FDEDEC",ph5:"EDE7F6"}[colorKey];
        return new TableRow({children:[
          cell(ph,W[0],{center:true,bold:true,color:C[colorKey]||C.navy,fill}),
          cell(dur,W[1],{center:true,fill,sm:true}),
          cell(theme,W[2],{center:true,bold:true,fill,sm:true}),
          cell(work,W[3],{fill,xs:true}),
          cell(out,W[4],{fill,xs:true}),
        ]});
      })
    ]
  });
}

// ═══════════════════════════════════════════════════════════════
// 表: 国際体系との比較
// ═══════════════════════════════════════════════════════════════
function tblInternational() {
  const W=[1400,800,1200,1000,1600,2500]; // 8500
  const data = [
    ["RMR\n（Bieniawski 1973）","連続点数\n0〜100","トンネル・斜面・基礎","強度・RQD・節理間隔・\n節理状態・地下水の6項",
     "本体系のSコードがRMRの強度項目に対応","最も普及した体系。ダム基礎への適用は少なく、DMRへ発展"],
    ["Q-system\n（Barton 1974）","連続指数\n0.001〜1000","主にトンネル","RQD・節理組数・粗さ・\n変質度・湧水・応力比",
     "RコードがQ-systemのリスク要素と部分対応","地下空洞向け。ダム基礎への直接適用は限定的"],
    ["GSI\n（Hoek 1994）","連続指数\n10〜100","岩盤強度推定","岩盤構造・節理面状態の2軸","本体系のS+Rコードに相当する情報を包含",
     "Hoek-Brown強度基準と連動。変質・破砕帯に強い"],
    ["DMR\n（Romana 2003）","連続点数\nRMRから派生","★ダム基礎専用★","RMR＋ダム安定補正係数\n（節理傾斜・方向）",
     "本体系のW/Gコードが目指す方向性に最も近い国際体系。点数系vs記号系の違い",
     "唯一のダム専用分類。ただし北海道の複合地質・年代軸は未考慮"],
    ["PWRI岩盤分類\n（土木研究所）","4区分\nCH〜D","日本国内標準\nダム・道路等","岩盤の健全性・新鮮度・\n割れ目状態",
     "SコードとCH=S1〜S2、CM=S3、CL=S4、D=S5で対応可能",
     "国内標準。地質時代・水理特性の体系化は範囲外"],
    ["本プロジェクト\n体系（新提案）","記号コード\n多ブロック","★北海道ダム専用★","地質年代×岩石種×\nS/R/W/G/(D/E/C/V/K)",
     "全体系の要素を統合しつつ地質時代軸・流域地質・劣化予測・AI連携を加えた独自体系",
     "既存体系にない独自性：①地質年代軸、②流域地質(C)、③劣化(D)、④気候変動(V)、⑤知識段階(K)"],
  ];
  return new Table({
    width:{size:8500,type:WidthType.DXA}, columnWidths:W,
    rows:[
      row([["体系名",W[0],{center:true}],["表現形式",W[1],{center:true}],["主な用途",W[2],{center:true}],
           ["主要パラメータ",W[3],{center:true}],["本体系との対応",W[4],{center:true}],
           ["特徴・限界",W[5],{center:true}]], true, 40),
      ...data.map(([nm,fmt,use,param,corr,feat],i)=>{
        const fill=i%2===0?C.row0:C.row1;
        const isOurs = i===5;
        return new TableRow({children:[
          cell(nm,W[0],{fill:isOurs?"FFF3CD":fill,bold:isOurs,color:isOurs?C.darkRed:C.navy,sm:true,center:true}),
          cell(fmt,W[1],{fill,center:true,xs:true}),
          cell(use,W[2],{fill,xs:true}),
          cell(param,W[3],{fill,xs:true}),
          cell(corr,W[4],{fill,xs:true}),
          cell(feat,W[5],{fill:isOurs?"FFF3CD":fill,xs:true}),
        ]});
      })
    ]
  });
}

// ═══════════════════════════════════════════════════════════════
// 表: Phase1 地質時代区分（完全版）
// ═══════════════════════════════════════════════════════════════
function tblGeoAgeClassification() {
  const W=[680,750,1050,1250,900,1050,2820]; // 8500
  const hdr = row([
    ["区分コード",W[0],{center:true}],["年代範囲",W[1],{center:true}],
    ["年代区分",W[2],{center:true}],["代表岩石種",W[3],{center:true}],
    ["強度目安(S)",W[4],{center:true}],["リスク指標(R)",W[5],{center:true}],
    ["北海道の地質帯・ダム地質上の意義",W[6],{center:true}]
  ], true, 40);

  const rows_=[
    secRow("■ Ⅰ類（古期地質） Prior to Neogene　— 先新第三紀（> 23 Ma）",C.navy, 7),
    [["Ⅰ-a",C.navy],"250〜50 Ma","変成岩帯","結晶片岩・片麻岩・角閃岩","S1〜S2","R1・R2","日高変成帯・神居古潭帯。高温変成による強度は高いが断層（R1）・変質帯（R2）リスク大。沙流川・日高系ダムの基礎"],
    [["Ⅰ-b",C.navy],"110〜50 Ma","花崗岩類","花崗岩・花崗閃緑岩","S1","R1","豊平川花崗岩体・増毛山地。最も安定な基礎岩盤。アーチダム適地（豊平峡）"],
    [["Ⅰ-c",C.navy],"145〜23 Ma","砂岩泥岩互層","タービダイト・チャート","S2〜S3","R1","蝦夷層群・十勝帯・天北地向斜。北海道最多の基礎地質（99基）。層理傾斜が透水性を支配"],
    [["Ⅰ-d",C.navy],"150〜100 Ma","超苦鉄質岩","カンラン岩・蛇紋岩","S1〜S4","R1・R2・R3","幌満かんらん岩体・神居古潭。蛇紋岩化（R2）・膨潤性（R3）が設計の最難課題"],
    [["Ⅰ-e",C.navy],"541〜252 Ma","古生界堆積岩","石灰岩・チャート・変成堆積岩","S2〜S3","R1・R5","渡島半島・天塩・留萌。石灰岩のカルスト空洞（R5）が遮水設計の核心問題"],
    secRow("■ Ⅱ類（新期地質） Neogene–Quaternary　— 新第三紀〜現在（< 23 Ma）","5B4FA0", 7),
    [["Ⅱ-a","5B4FA0"],"12〜0.1 Ma","溶結凝灰岩","ウェルデッドタフ・火砕流堆積物","S1〜S3","R4","支笏・洞爺・阿寒カルデラ起源。冷却亀裂（R4）が卓越した透水経路。日新・東郷・しろがね・古梅ダム"],
    [["Ⅱ-b","5B4FA0"],"23〜0.01 Ma","火山岩（溶岩）","安山岩・玄武岩・デイサイト","S2〜S3","R4","大雪・十勝・羊蹄・駒ヶ岳等。溶岩スタック構造の把握が設計鍵"],
    [["Ⅱ-c","5B4FA0"],"23〜1 Ma","火山砕屑岩","凝灰岩・火山礫凝灰岩","S3〜S4","R4","石狩低地帯周縁・道東中新統。固結度変動が大きい"],
    [["Ⅱ-d","5B4FA0"],"23〜2.6 Ma","新第三系堆積岩","砂岩・泥岩・礫岩","S3〜S4","R1","天北地向斜・石狩低地帯。農業ダム標準地質（67基）"],
    [["Ⅱ-e","5B4FA0"],"2.6 Ma〜現在","未固結堆積物","河床礫層・段丘堆積物","S4〜S5","R6","沖積〜段丘。フィルダム基礎の主体（54基）。止水工設計が要"],
  ];

  const trows=[hdr];
  let alt=0;
  for(const r of rows_){
    if(r instanceof TableRow){trows.push(r);continue;}
    const [[code,codeC],age,age_nm,rock,s,risk,note_]=r;
    const fill=alt%2===0?C.row0:C.row1; alt++;
    trows.push(new TableRow({children:[
      cell(code,W[0],{center:true,bold:true,color:codeC,fill:"EEF4FF"}),
      cell(age,W[1],{center:true,fill,xs:true}),
      cell(age_nm,W[2],{center:true,fill,sm:true}),
      cell(rock,W[3],{fill,sm:true}),
      cell(s,W[4],{center:true,fill,sm:true}),
      cell(risk,W[5],{center:true,fill,sm:true,color:C.darkRed}),
      cell(note_,W[6],{fill,xs:true}),
    ]}));
  }
  return new Table({width:{size:8500,type:WidthType.DXA},columnWidths:W,rows:trows});
}

// ═══════════════════════════════════════════════════════════════
// 表: Phase1 強度(S)・リスク(R)コード
// ═══════════════════════════════════════════════════════════════
function tblSR() {
  const W1=[620,1100,1500,1280,4000];
  const filS=["D5E8F7","E8F4FB","F5FBFF","FFF8E7","FDEBD0","F5F5F5"];
  const sData=[
    ["S1","極 硬 岩","> 200 MN/m²","点荷重 > 8 MPa","花崗岩・強変成岩・強溶結凝灰岩。日新ダム qu≈2,000 MN/m²。Ⅰ-a・Ⅰ-b・Ⅱ-a（完全溶結）"],
    ["S2","硬　 岩","100〜200","点荷重 4〜8 MPa","堅硬砂岩・安山岩溶岩・花崗閃緑岩。Ⅰ-b・Ⅰ-c（堅硬部）・Ⅱ-b"],
    ["S3","中 硬 岩","25〜100","点荷重 1〜4 MPa","溶結凝灰岩（中程度）・一般砂岩・玄武岩。Ⅱ-a（中溶結）・Ⅱ-c・Ⅰ-c"],
    ["S4","軟　 岩","5〜25","点荷重 0.2〜1","軟質凝灰岩・泥岩・半固結礫岩。Ⅱ-c・Ⅱ-d・Ⅰ-d（変質部）"],
    ["S5","極軟岩・土質","< 5 MN/m²","点荷重 < 0.2","未固結礫層・砂・粘土。Ⅱ-e全般"],
    ["S?","情 報 な し","—","—","試験データなし。設計・施工報告書参照要"],
  ];
  const tblS = new Table({
    width:{size:8500,type:WidthType.DXA},columnWidths:W1,
    rows:[
      row([["コード",W1[0],{center:true}],["区分名",W1[1],{center:true}],
           ["一軸圧縮強度 qu",W1[2],{center:true}],["補助指標 Is50",W1[3],{center:true}],
           ["代表岩石・北海道地質区分対応",W1[4],{center:true}]], true, 36),
      ...sData.map(([code,nm,qu,sub,rock],i)=>new TableRow({children:[
        cell(code,W1[0],{center:true,bold:true,color:code==="S?"?C.gray:C.navy,fill:filS[i]}),
        cell(nm,W1[1],{center:true,fill:filS[i],sm:true}),
        cell(qu,W1[2],{center:true,fill:filS[i],sm:true,bold:true}),
        cell(sub,W1[3],{center:true,fill:filS[i],sm:true}),
        cell(rock,W1[4],{fill:filS[i],xs:true}),
      ]}))
    ]
  });

  const W2=[620,1100,5780];
  const rData=[
    ["R1","断層・せん断帯リスク","活断層・破砕帯・断層粘土。すべり面形成・透水路の主因。Ⅰ-a・Ⅰ-c・Ⅰ-d帯で顕著"],
    ["R2","変質・変成リスク","熱水変質・接触変成・蛇紋岩化。局所的強度低下・膨潤。Ⅰ-a・Ⅰ-d"],
    ["R3","膨張性リスク","蛇紋岩（クリソタイル）・石膏・モンモリロナイト。経時的な体積変化。Ⅰ-d"],
    ["R4","冷却亀裂リスク","柱状節理・板状節理（溶結凝灰岩・溶岩）。高透水経路。Ⅱ-a・Ⅱ-b"],
    ["R5","溶解・空洞リスク","石灰岩カルスト空洞。グラウト流失・遮水不能。Ⅰ-e"],
    ["R6","未固結層リスク","液状化・パイピング・内部侵食。フィルダム基礎の主要リスク。Ⅱ-e"],
  ];
  const tblR = new Table({
    width:{size:8500,type:WidthType.DXA},columnWidths:W2,
    rows:[
      row([["コード",W2[0],{center:true}],["リスク区分",W2[1],{center:true}],
           ["内容・対象地質",W2[2],{center:true}]], true, 36),
      ...rData.map(([code,nm,desc],i)=>new TableRow({children:[
        cell(code,W2[0],{center:true,bold:true,color:C.darkRed,fill:i%2===0?C.row0:C.row1}),
        cell(nm,W2[1],{fill:i%2===0?C.row0:C.row1,bold:true,sm:true}),
        cell(desc,W2[2],{fill:i%2===0?C.row0:C.row1,xs:true}),
      ]}))
    ]
  });
  return [tblS, SP(100), tblR];
}

// ═══════════════════════════════════════════════════════════════
// 表: Phase1 全188基 統計サマリー
// ═══════════════════════════════════════════════════════════════
function tblStatsSummary() {
  const W=[2000,1200,800,4500]; // 8500
  const stats=[
    ["Ⅰ-c（砂岩泥岩）主体","99基（52.7%）","主体","白亜紀〜古第三紀タービダイト。石狩川系・天塩川・夕張山地が中心"],
    ["Ⅱ-d（新第三系堆積岩）","67基（35.6%）","新期","天北地向斜・石狩低地帯縁辺の農業ダム多数"],
    ["Ⅱ-e（未固結堆積物）","54基（28.7%）","表層","フィルダム基礎の主体。農業・灌漑ダムに多い"],
    ["Ⅱ-a（溶結凝灰岩）","16基（8.5%）","火山","支笏・洞爺カルデラ周辺。高透水（W4）・高難度（G3）"],
    ["Ⅱ-b（火山岩溶岩）","18基（9.6%）","火山","大雪・十勝・ニセコ・駒ヶ岳等"],
    ["Ⅰ-a（変成岩帯）","18基（9.6%）","古期","日高変成帯・神居古潭帯"],
    ["Ⅰ-e（古生界）","13基（6.9%）","古期","渡島・天塩・留萌。石灰岩カルストR5が要注意"],
    ["Ⅰ-d（超苦鉄質岩）","2基（1.1%）","古期","幌満かんらん岩体。蛇紋岩R2・R3・高難度G4"],
  ];
  return new Table({
    width:{size:8500,type:WidthType.DXA},columnWidths:W,
    rows:[
      row([["主要地質区分",W[0],{center:true}],["基数・割合",W[1],{center:true}],
           ["分類",W[2],{center:true}],["特徴・北海道ダム分布",W[3],{center:true}]], true, 36),
      ...stats.map(([gname,cnt,cls,note_],i)=>new TableRow({children:[
        cell(gname,W[0],{fill:i%2===0?C.row0:C.row1,bold:true,color:C.navy}),
        cell(cnt,W[1],{center:true,fill:i%2===0?C.row0:C.row1,bold:true}),
        cell(cls,W[2],{center:true,fill:i%2===0?C.row0:C.row1}),
        cell(note_,W[3],{fill:i%2===0?C.row0:C.row1,xs:true}),
      ]}))
    ]
  });
}

// ═══════════════════════════════════════════════════════════════
// 表: Phase2 W・G コード
// ═══════════════════════════════════════════════════════════════
function tblWG() {
  // Wコード
  const WW=[620,1200,1280,1550,1300,2550];
  const wFil=["D6EAF8","EAF4FB","F5FBFF","FEF9E7","FDEDEC","F5F5F5"];
  const wData=[
    ["W1","極 低 透 水","< 1 Lu","< 10⁻⁸ m/s","粒間浸透（事実上遮水）","変成岩緻密部・花崗岩深部。グラウト不要。Ⅰ-a・Ⅰ-b"],
    ["W2","低 透 水","1〜5 Lu","10⁻⁸〜10⁻⁶ m/s","微細亀裂浸透","標準健全岩盤。1列カーテングラウト十分。Ⅰ-b・Ⅱ-b（緻密）"],
    ["W3","中 透 水","5〜30 Lu","10⁻⁶〜10⁻⁵ m/s","開口亀裂卓越","亀裂系が主経路。美生ダム実績（開口亀裂・砂岩部卓越）。Ⅰ-c・Ⅱ-d"],
    ["W4","高 透 水","30〜100 Lu","10⁻⁵〜10⁻³ m/s","大亀裂・断層・冷却亀裂","日新ダム実績：10¹ m/day≈W4相当。Ⅱ-a柱状節理典型"],
    ["W5","極 高 透 水","> 100 Lu","> 10⁻³ m/s","未固結層間隙・カルスト空洞","グラウト流失。止水矢板・遮水壁要。Ⅱ-e・Ⅰ-e石灰岩"],
    ["W?","情 報 な し","—","—","—","試験記録なし。Ph.1で解消すべき優先課題"],
  ];
  const tblW = new Table({
    width:{size:8500,type:WidthType.DXA},columnWidths:WW,
    rows:[
      row([["コード",WW[0],{center:true}],["区分名",WW[1],{center:true}],
           ["ルジオン値",WW[2],{center:true}],["透水係数 k",WW[3],{center:true}],
           ["透水機構",WW[4],{center:true}],["北海道ダムへの対応・実績",WW[5],{center:true}]], true, 40),
      ...wData.map(([code,nm,lu,k,mech,eng],i)=>new TableRow({children:[
        cell(code,WW[0],{center:true,bold:true,color:code==="W?"?C.gray:C.red,fill:wFil[i]}),
        cell(nm,WW[1],{center:true,fill:wFil[i],sm:true}),
        cell(lu,WW[2],{center:true,fill:wFil[i],sm:true,bold:true}),
        cell(k,WW[3],{center:true,fill:wFil[i],xs:true}),
        cell(mech,WW[4],{center:true,fill:wFil[i],xs:true}),
        cell(eng,WW[5],{fill:wFil[i],xs:true}),
      ]}))
    ]
  });

  // Gコード
  const GW=[620,1100,1680,1400,1200,2500];
  const gFil=["D6EAF8","EAF4FB","FEF9E7","FDEDEC","F5F5F5"];
  const gData=[
    ["G1","軽 微","< 50 kg/m（セメント）","コンソリのみ","W1〜W2","健全硬岩基礎。補強目的のみ。豊平峡アーチダム相当"],
    ["G2","標 準","50〜200 kg/m","カーテン1〜2列＋コンソリ","W2〜W3","石狩川系中流ダム標準。砂岩泥岩・安山岩・新第三系"],
    ["G3","高 難 度","200〜500 kg/m（多段反復）","多列カーテン＋高圧・超微粒子","W3〜W4","溶結凝灰岩柱状節理。日新・東郷・しろがね・古梅典型"],
    ["G4","特 殊 工 法","> 500 kg/m またはセメント以外","止水矢板・遮水壁・化学グラウト","W4〜W5","石灰岩カルスト(Ⅰ-e)・未固結礫層(Ⅱ-e)・蛇紋岩大破砕帯"],
    ["G?","情 報 な し","—","施工報告書参照要","—","グラウチング施工記録なし。優先確認対象"],
  ];
  const tblG = new Table({
    width:{size:8500,type:WidthType.DXA},columnWidths:GW,
    rows:[
      row([["コード",GW[0],{center:true}],["難易度",GW[1],{center:true}],
           ["注入量目安",GW[2],{center:true}],["主な工法",GW[3],{center:true}],
           ["対応W",GW[4],{center:true}],["北海道ダムへの適用・事例",GW[5],{center:true}]], true, 40),
      ...gData.map(([code,nm,qty,method,w,note_],i)=>new TableRow({children:[
        cell(code,GW[0],{center:true,bold:true,color:code==="G?"?C.gray:C.darkBlue,fill:gFil[i]}),
        cell(nm,GW[1],{center:true,fill:gFil[i],sm:true}),
        cell(qty,GW[2],{center:true,fill:gFil[i],xs:true}),
        cell(method,GW[3],{fill:gFil[i],xs:true}),
        cell(w,GW[4],{center:true,fill:gFil[i],sm:true}),
        cell(note_,GW[5],{fill:gFil[i],xs:true}),
      ]}))
    ]
  });
  return [tblW, SP(100), tblG];
}

// ═══════════════════════════════════════════════════════════════
// 表: Phase2 地質区分別W・G標準値（核心表）
// ═══════════════════════════════════════════════════════════════
function tblGeoHydraulic() {
  const W=[700,1200,800,1800,850,800,2350];
  const hdr = row([
    ["区分コード",W[0],{center:true}],["代表岩石種",W[1],{center:true}],
    ["強度(S)",W[2],{center:true}],["透水機構・主経路",W[3],{center:true}],
    ["透水性(W)\n★新設",W[4],{center:true,color:C.red,bold:true}],
    ["基礎処理(G)\n★新設",W[5],{center:true,color:C.darkBlue,bold:true}],
    ["北海道ダム工学上の要点・実績",W[6],{center:true}]
  ], true, 46);

  const secs=[
    {label:"■ Ⅰ類（古期地質）　Prior to Neogene", fill:C.navy, rows:[
      ["Ⅰ-a",C.navy,"変成岩（片岩・片麻岩）","S1〜S2","変成葉理沿い浸透が主体。緻密部W1。断層帯でW4に上昇","W1〜W2\n（断層帯W4）","G1〜G2\n（断層帯G3）","断層走向・傾斜が透水性支配。局部高透水帯の事前探査重要"],
      ["Ⅰ-b",C.navy,"花崗岩・花崗閃緑岩","S1","急冷節理・方状節理が主経路。深部W1。風化帯W3","W1〜W2\n（風化帯W3）","G1〜G2","豊平峡アーチダム相当。理想的基礎。グラウト量最小"],
      ["Ⅰ-c",C.navy,"砂岩泥岩互層（タービダイト）","S2〜S3","層理面・開口亀裂が主経路。砂岩W3、泥岩W1〜W2","W2〜W3","G2〜G3","美生ダム実績：開口亀裂（砂岩卓越）→W3/G2。最多分布"],
      ["Ⅰ-d",C.navy,"超苦鉄質岩・蛇紋岩","S1〜S4","蛇紋岩化破砕帯W4。母岩W2。透水性経時変動大","W3〜W4\n（破砕帯W4）","G3〜G4","様似・幌満。グラウト効果の確認難。化学グラウト検討要"],
      ["Ⅰ-e",C.navy,"石灰岩・チャート","S2〜S3","石灰岩カルストW5。チャートW1。二者混在で危険","W1（チャート）\n〜W5（石灰岩）","G2〜G4","カルスト空洞の事前探査（ボーリング・物理探査）最重要"],
    ]},
    {label:"■ Ⅱ類（新期地質）　Neogene–Quaternary", fill:"5B4FA0", rows:[
      ["Ⅱ-a","5B4FA0","溶結凝灰岩（カルデラ起源）","S1〜S3","柱状節理（冷却亀裂）卓越。鉛直・水平大亀裂","W3〜W4","G3","日新実績：10¹ m/day≈W4。東郷・しろがね・古梅も同様"],
      ["Ⅱ-b","5B4FA0","安山岩・玄武岩（溶岩）","S2〜S3","流理・板状節理。完全溶岩体W2。冷却面W4","W2〜W3\n（冷却面W4）","G2〜G3","溶岩スタック構造把握が設計の鍵。大雪・ニセコ等"],
      ["Ⅱ-c","5B4FA0","凝灰岩・火山礫凝灰岩","S3〜S4","粒間浸透＋亀裂透水の混合。固結度依存","W2〜W3","G2〜G3","後志・空知・道東中新統。固結度変動で±1コード"],
      ["Ⅱ-d","5B4FA0","砂岩・泥岩（新第三系）","S3〜S4","砂岩層粒間浸透。層理面沿い透水","W2〜W3","G2","天北地向斜・石狩低地帯農業ダム標準地質"],
      ["Ⅱ-e","5B4FA0","未固結礫層・段丘堆積物","S4〜S5","粒間透水卓越。礫層W5（k>10⁻³ m/s）","W4〜W5","G3〜G4","農業フィルダム基礎。止水矢板・遮水壁等補助工法必須"],
    ]},
  ];

  const trows=[hdr];
  let alt=0;
  for(const s of secs){
    trows.push(secRow(s.label,s.fill,7));
    for(const [code,cc,rock,str,mech,w,g,tip] of s.rows){
      const fill=alt%2===0?C.row0:C.row1; alt++;
      trows.push(new TableRow({children:[
        cell(code,W[0],{center:true,bold:true,color:cc,fill:"EEF4FF"}),
        cell(rock,W[1],{center:true,fill,sm:true}),
        cell(str,W[2],{center:true,fill,sm:true}),
        cell(mech,W[3],{fill,xs:true}),
        cell(w,W[4],{center:true,bold:true,color:C.red,fill}),
        cell(g,W[5],{center:true,bold:true,color:C.darkBlue,fill}),
        cell(tip,W[6],{fill,xs:true}),
      ]}));
    }
  }
  return new Table({width:{size:8500,type:WidthType.DXA},columnWidths:W,rows:trows});
}

// ═══════════════════════════════════════════════════════════════
// 表: Phase2 次世代コード5提案
// ═══════════════════════════════════════════════════════════════
function tblNextGenCodes() {
  const W=[700,1000,1500,1400,3900]; // 8500
  const codes=[
    ["D","経時劣化\nポテンシャル","D1〜D5\n＋D?","国際先例なし\n北海道独自追加",
     "D1:変化極小（花崗岩・変成岩）　D2:緩慢溶解（砂岩泥岩）　D3:蛇紋岩化進行（Ⅰ-d）\nD4:凍結融解亀裂拡大（北海道特有・Ⅱ-a露出部）　D5:未固結層圧密・液状化（Ⅱ-e）\n100年スケールの劣化ポテンシャルを記号化。老朽化ダム優先再評価に直結"],
    ["E","地震応答\n増幅","E1〜E5\n＋E?","Vs30概念参照\n地震工学応用",
     "E1:Vs>1500 m/s（増幅なし、岩盤）　E2:700〜1500（軽微増幅）　E3:300〜700（中程度）\nE4:150〜300（強増幅・段丘礫）　E5:<150（液状化域）\n地質区分→Vsの統計対応を体系化。耐震再評価ガイドラインとの連動が可能"],
    ["C","貯水池・\n流域地質","C1〜C5\n＋C?","★最も独創的★\n既存体系に先例なし",
     "C1:硬岩卓越・崩壊リスク低　C2:脆弱層あり・部分崩壊リスク　C3:カルスト・漏水リスク\nC4:活火山流域・土石流リスク（十勝・有珠周辺）　C5:重金属・酸性水リスク（蛇紋岩・鉱山跡地）\nダム基礎だけでなく貯水池・流域全体の地質を評価対象とする発想の転換"],
    ["V","気候変動\n脆弱性","V1〜V4\n＋V?","気候科学との\n統合（新領域）",
     "V1:安定（気候変化の影響小）　V2:融雪洪水ピーク増加→基礎水圧上昇\nV3:凍結融解サイクル変化→亀裂透水性変動（北海道特有）　V4:山岳永久凍土融解→斜面不安定\n21世紀末にかけての将来変化ポテンシャルを記号化。長期管理計画の基盤に"],
    ["K","知識蓄積\n段階","K1〜K5\n＋K?","★最も実用的★\n既存体系に先例なし",
     "K1:ボーリング＋試験＋施工記録完備　K2:設計報告書あり・施工記録一部欠落\nK3:地質図推定のみ・現地調査なし　K4:類推（位置・水系から推定）\nK5:AI推定（衛星・地質図の機械学習推論）←将来実装\n「調査していない」ことを明示するフラグ。情報収集の優先順位管理に直結"],
  ];
  return new Table({
    width:{size:8500,type:WidthType.DXA},columnWidths:W,
    rows:[
      row([["コード",W[0],{center:true}],["名称",W[1],{center:true}],
           ["記号体系",W[2],{center:true}],["参照・根拠",W[3],{center:true}],
           ["定義・内容・北海道ダムへの意義",W[4],{center:true}]], true, 38),
      ...codes.map(([code,nm,sym,ref,desc],i)=>{
        const fills=["D6EAF8","E8F8F0","FEF9E7","FFF0F5","EDE7F6"];
        const colors=["1F3864","1E8449","C0392B","8B4513","6C3483"];
        return new TableRow({children:[
          cell(code,W[0],{center:true,bold:true,color:colors[i],fill:fills[i]}),
          cell(nm,W[1],{center:true,fill:fills[i],sm:true,bold:true}),
          cell(sym,W[2],{center:true,fill:fills[i],sm:true}),
          cell(ref,W[3],{fill:fills[i],xs:true}),
          cell(desc,W[4],{fill:fills[i],xs:true}),
        ]});
      })
    ]
  });
}

// ═══════════════════════════════════════════════════════════════
// 表: フル記号体系（全ブロック）
// ═══════════════════════════════════════════════════════════════
function tblFullSymbol() {
  const W=[800,700,1100,1100,900,900,900,900,1200]; // 8500
  return new Table({
    width:{size:8500,type:WidthType.DXA},columnWidths:W,
    rows:[
      row([
        ["ブロック",W[0],{center:true}],["記号形式",W[1],{center:true}],["段階数",W[2],{center:true}],
        ["指標",W[3],{center:true}],["フェーズ",W[4],{center:true}],
        ["Phase1\n現行確立",W[5],{center:true}],["Phase2A\nW/G追加",W[6],{center:true}],
        ["Phase2B\nD/E/C/V",W[7],{center:true}],["Phase2C\nK追加",W[8],{center:true}]
      ], true, 42),
      ...[
        ["①時代大区分","Ⅰ・Ⅱ","2","地質生成年代","基本軸","●","●","●","●"],
        ["②岩石種","a〜e","各5","岩石種タイプ","基本軸","●","●","●","●"],
        ["③強度（S）","S1〜S5/S?","5+?","一軸圧縮強度","力学","●","●（S?追加）","●","●"],
        ["④リスク（R）","R1〜R6(複数)","6+","工学的リスク","力学","●","●","●","●"],
        ["⑤透水性（W）","W1〜W5/W?","5+?","ルジオン値Lu","水理","—","●★","●","●"],
        ["⑥基礎処理（G）","G1〜G4/G?","4+?","グラウト難易度","水理","—","●★","●","●"],
        ["⑦経時劣化（D）","D1〜D5/D?","5+?","劣化ポテンシャル","長期","—","—","●★","●"],
        ["⑧地震応答（E）","E1〜E5/E?","5+?","Vs30・増幅率","地震","—","—","●★","●"],
        ["⑨流域地質（C）","C1〜C5/C?","5+?","貯水池・流域","流域","—","—","●★","●"],
        ["⑩気候変動（V）","V1〜V4/V?","4+?","将来変化","長期","—","—","●★","●"],
        ["⑪知識段階（K）","K1〜K5/K?","5+?","情報充足度","管理","—","—","—","●★"],
      ].map(([blk,sym,n,idx,cat,...phases],i)=>{
        const catFill={"基本軸":"D6EAF8","力学":"D5F5E3","水理":"FDEBD0","長期":"FEF9E7","地震":"EDE7F6","流域":"E8F8F0","管理":"F9EBEA"}[cat]||C.row0;
        return new TableRow({children:[
          cell(blk,W[0],{fill:catFill,bold:true,sm:true}),
          cell(sym,W[1],{center:true,fill:catFill,sm:true,bold:true,color:phases[0]==="●"?C.navy:C.gray}),
          cell(n,W[2],{center:true,fill:catFill,sm:true}),
          cell(idx,W[3],{fill:catFill,sm:true}),
          cell(cat,W[4],{center:true,fill:catFill,xs:true}),
          ...phases.map((v,j)=>cell(v,W[5+j],{center:true,fill:v==="●★"?"FFF3CD":(v==="●"?"D5F5E3":C.lightGray),bold:v!=="—",color:v==="●★"?C.darkRed:(v==="●"?C.darkGreen:C.gray),sm:true})),
        ]});
      })
    ]
  });
}

// ═══════════════════════════════════════════════════════════════
// 表: 完全記号 適用例
// ═══════════════════════════════════════════════════════════════
function tblCompleteSymbolExamples() {
  const W=[3000,700,4800];
  const confFill={A:"D5F5E3",B:"D6EAF8",C:"FEF9E7"};
  const examples=[
    ["Ⅱ-a（S1/R4・W4/G3）・Ⅱ-e（S5/R6・W5/G4）","A",
     "溶結凝灰岩（極硬・冷却亀裂R4・高透水W4・多段グラウトG3）＋河床礫層（極軟弱・未固結R6・極高透水W5・特殊工法G4）。日新ダム相当。実績：qu≈2,000 MN/m²、10¹ m/day"],
    ["Ⅱ-a（S3/R4・W4/G3）","A","中程度溶結凝灰岩（中硬・冷却亀裂R4・高透水W4・多段グラウトG3）。東郷・聖台・新区画・古梅ダム相当。柱状節理が主要漏水経路"],
    ["Ⅰ-c（S2/R1・W3/G2）・Ⅱ-e（S5/R6・W5/G4）","A",
     "砂岩粘板岩タービダイト（硬岩・断層R1・中透水W3・標準グラウトG2）＋未固結礫層（W5/G4）。美生ダム相当。開口亀裂砂岩卓越・Well locked礫層"],
    ["Ⅰ-b（S1/R1・W1/G1）","B","花崗岩（極硬岩・極低透水W1・グラウト軽微G1）。豊平峡アーチダム相当。理想的基礎。グラウチングは補強目的"],
    ["Ⅱ-a（S2/R4・W4/G3）","B","強溶結凝灰岩（支笏カルデラ起源）。漁川・千歳川系ダム相当。柱状節理W4は推定確実、詳細Lu値は施工報告書確認要"],
    ["Ⅰ-d（S2/R2・R3・W4/G4）・Ⅱ-e（S5/R6・W5/G4）","B",
     "幌満かんらん岩・蛇紋岩（破砕帯W4・特殊工法G4）＋礫層（W5/G4）。様似相当。蛇紋岩膨潤R3と高透水の複合が最難関"],
    ["Ⅰ-c（S2/R1・W?/G?）・Ⅱ-d（S3/R1・W?/G?）","C",
     "道北・留萌帯の砂岩泥岩＋新第三系。強度・リスクは推定可、透水試験データ未参照のためW?/G?。Ph.1で施工報告書確認要"],
    ["Ⅱ-e（S4/R6・W?/G?）","C","石狩平野縁辺農業フィルダム。W4〜W5/G3〜G4と推定されるが試験記録未確認。未固結農業ダムとして優先確認対象"],
  ];
  return new Table({
    width:{size:8500,type:WidthType.DXA},columnWidths:W,
    rows:[
      row([["完全記号（S/R・W/G ブロック）",W[0],{center:true}],["信頼度",W[1],{center:true}],
           ["読み方・地質工学的意味",W[2],{center:true}]], true, 36),
      ...examples.map(([sym,conf,desc])=>new TableRow({children:[
        cell(sym,W[0],{bold:true,color:C.navy,fill:confFill[conf],sm:true}),
        cell(conf,W[1],{center:true,bold:true,fill:confFill[conf]}),
        cell(desc,W[2],{fill:confFill[conf],xs:true}),
      ]}))
    ]
  });
}

// ═══════════════════════════════════════════════════════════════
// 表: 優先調査対象ダム
// ═══════════════════════════════════════════════════════════════
function tblPriorityDams() {
  const W=[1600,700,1000,1000,1600,2600];
  const data=[
    ["溶結凝灰岩帯（Ⅱ-a）","16基","W4/G3推定","B→A","ルジオン試験値・グラウチング量の確認","日新・東郷・しろがね・古梅（信頼度A確定済）以外の12基。propylite.work・北海道開発局報告書参照"],
    ["石灰岩・古生界（Ⅰ-e）","13基","W1〜W5\nG2〜G4推定","C→B","カルスト空洞探査記録・グラウチング記録","渡島・天塩・留萌帯。空洞充填の施工記録が最優先情報"],
    ["超苦鉄質岩・蛇紋岩（Ⅰ-d）","2基","W4/G4推定","C→B","破砕帯の分布・蛇紋岩化程度","幌満川第三・様似ダム相当。R2（変質）・R3（膨潤）と高透水の複合評価"],
    ["1920年代以前竣工","26基","多様","C→B","建設当時の地質調査記録の発掘","最も老朽化が進む群。基礎処理の再確認と劣化コード（D）付与が急務"],
    ["農業フィルダム\n（未固結基礎）","〜50基","W4〜W5\nG3〜G4","C→B","農林水産省農業水産部施工記録","Ⅱ-e基礎の多くはW?/G?状態。農業ダム施工記録は開発局とは別ルートで入手必要"],
    ["日高変成帯（Ⅰ-a）","18基","W2/G2\n断層帯W4/G3","B→A","断層破砕帯の調査記録","沙流川・静内川・新冠川系。活断層（R1）との交差関係が透水性を左右"],
  ];
  return new Table({
    width:{size:8500,type:WidthType.DXA},columnWidths:W,
    rows:[
      row([["対象地質区分",W[0],{center:true}],["基数",W[1],{center:true}],
           ["推定W/G",W[2],{center:true}],["目標信頼度",W[3],{center:true}],
           ["確認すべき情報",W[4],{center:true}],["調査方針・主要出典",W[5],{center:true}]], true, 38),
      ...data.map(([gp,cnt,wg,tgt,info,strat],i)=>new TableRow({children:[
        cell(gp,W[0],{fill:i%2===0?C.row0:C.row1,bold:true,color:C.navy,sm:true}),
        cell(cnt,W[1],{center:true,fill:i%2===0?C.row0:C.row1,bold:true}),
        cell(wg,W[2],{center:true,fill:i%2===0?C.row0:C.row1,sm:true,bold:true,color:C.darkBlue}),
        cell(tgt,W[3],{center:true,fill:i%2===0?"D5F5E3":"D6EAF8",bold:true,color:C.darkGreen}),
        cell(info,W[4],{fill:i%2===0?C.row0:C.row1,xs:true}),
        cell(strat,W[5],{fill:i%2===0?C.row0:C.row1,xs:true}),
      ]}))
    ]
  });
}

// ═══════════════════════════════════════════════════════════════
// 表: フェーズ別詳細作業計画
// ═══════════════════════════════════════════════════════════════
function tblPhaseDetail(ph, color, rows_data) {
  const W=[2000,1800,2200,2500];
  return new Table({
    width:{size:8500,type:WidthType.DXA},columnWidths:W,
    rows:[
      row([["作業項目",W[0],{center:true,hdrColor:color}],["担当・データ源",W[1],{center:true,hdrColor:color}],
           ["成果物",W[2],{center:true,hdrColor:color}],["判定基準・品質指標",W[3],{center:true,hdrColor:color}]], true, 36),
      ...rows_data.map(([item,src,out,qual],i)=>new TableRow({children:[
        cell(item,W[0],{fill:i%2===0?C.row0:C.row1,bold:true,sm:true}),
        cell(src,W[1],{fill:i%2===0?C.row0:C.row1,xs:true}),
        cell(out,W[2],{fill:i%2===0?C.row0:C.row1,xs:true}),
        cell(qual,W[3],{fill:i%2===0?C.row0:C.row1,xs:true}),
      ]}))
    ]
  });
}

// ═══════════════════════════════════════════════════════════════
// ドキュメント本体 組み立て
// ═══════════════════════════════════════════════════════════════
function buildDoc() {
  const NUMBERING = {
    config:[
      {reference:"b1",levels:[{level:0,format:LevelFormat.BULLET,text:"●",
        style:{paragraph:{indent:{left:640,hanging:320},spacing:{before:60,after:60}}}}]},
      {reference:"b2",levels:[{level:0,format:LevelFormat.BULLET,text:"○",
        style:{paragraph:{indent:{left:1000,hanging:320},spacing:{before:40,after:40}}}}]},
    ]
  };

  const children = [
    // ══════════ 表　紙 ══════════
    SP(1200),
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:280},
      children:[new TextRun({text:"北海道ダム地質分類プロジェクト",font:"游明朝",size:50,bold:true,color:C.navy})]}),
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:80},
      children:[new TextRun({text:"全体計画書（フルプラン）",font:"游明朝",size:44,bold:true,color:C.darkBlue})]}),
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:80},
      children:[new TextRun({text:"Integrated Master Plan",font:"Times New Roman",size:28,italics:true,color:"444444"})]}),
    DIV(C.darkBlue),
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:160,after:80},
      children:[new TextRun({text:"― 地質年代×工学特性×水理特性×将来リスクの統合分類体系 ―",font:"游明朝",size:24,color:"333333"})]}),
    SP(400),
    new Table({width:{size:6000,type:WidthType.DXA},columnWidths:[2500,3500],rows:[
      new TableRow({children:[
        cell("対象ダム",2500,{bold:true,fill:"EEF4FF"}),cell("北海道既設・建設中ダム 188基",3500,{fill:"EEF4FF"})]}),
      new TableRow({children:[
        cell("分類軸",2500,{bold:true}),cell("地質年代 / 岩石種 / 強度(S) / リスク(R) / 透水性(W) / 基礎処理(G) + 次世代5コード",3500,{xs:true})]}),
      new TableRow({children:[
        cell("Phase 1（確立済み）",2500,{bold:true,fill:"D6EAF8"}),cell("地質時代区分・S/Rコード・全188基暫定分類",3500,{fill:"D6EAF8",xs:true})]}),
      new TableRow({children:[
        cell("Phase 2（本書提案）",2500,{bold:true,fill:"FDEBD0"}),cell("W/Gコード（水理特性）+ D/E/C/V/Kコード（次世代）",3500,{fill:"FDEBD0",xs:true})]}),
      new TableRow({children:[
        cell("Phase 3〜5",2500,{bold:true}),cell("全基完全判定・個別カード・GIS・国際発信",3500,{xs:true})]}),
      new TableRow({children:[
        cell("作成日",2500,{bold:true,fill:"EEF4FF"}),cell("2026年3月",3500,{fill:"EEF4FF"})]}),
    ]}),
    SP(200),
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:80,after:0},
      children:[new TextRun({text:"Phase 1 基本体系確立 → Phase 2 水理・次世代コード追加 → Phase 3〜5 全基判定・成果物整備",
        font:"游明朝",size:19,color:"555555",italics:true})]}),

    PB_BREAK(),

    // ══════════ Part 0: エグゼクティブサマリー ══════════
    H1("エグゼクティブサマリー"),
    P("本プロジェクトは、北海道の既設・建設中ダム 188基を対象に、地質学的・工学的・水理的・将来リスクの4軸を統合した世界初のダム地質分類体系の構築を目指すものである。"),
    P("現在（2026年3月時点）、Phase 1として地質生成年代に基づく時代区分（Ⅰ/Ⅱ類）と岩石種（a〜e）、強度コード（S）、リスク指標（R）の基本体系が確立され、全188基の暫定分類が完了している。Phase 2では水理的特性コード（W・G）の新設と、経時劣化（D）・地震応答（E）・流域地質（C）・気候変動（V）・知識段階（K）の次世代5コードを提案する。"),
    SP(100),
    tblRoadmap(),
    SP(100),
    NOTE("Phase 2Bの次世代コード（D/E/C/V）は提案・定義段階。Phase 2Cの知識段階コード（K）は全フェーズを通じた管理ツールとして機能する。"),
    NOTE("各フェーズは独立して成果を生み出せる設計とし、途中段階でも実務利用可能な部分成果物を随時提供する。"),

    PB_BREAK(),

    // ══════════ Part I: 背景と目的 ══════════
    H1("第Ⅰ部　プロジェクトの背景と目的"),

    H2("1. 北海道ダムが直面する課題"),
    P("北海道の188基のダムは1910年代から2020年代にわたって建設されており、このうち1920年代以前に竣工したダムは26基に上る。老朽化の進行、当初設計時の地質評価手法の限界、および現代的な安全基準との乖離が課題として浮上している。"),
    P("加えて、北海道は他の地域にない複合的な地質リスクを抱えている。支笏・洞爺カルデラに起源を持つ溶結凝灰岩の冷却亀裂（柱状節理）は特異な高透水経路を形成し、蛇紋岩体の膨潤リスク、石灰岩のカルスト空洞、未固結礫層上のフィルダム等、多様な工学的難題が地域ごとに潜在する。"),
    P("しかし、これら188基の地質特性を統一的な体系で記述した公開データベースはこれまで存在しなかった。各ダムの地質情報は個別の施工報告書や学術論文に散在しており、ダム建設未経験の技術者が全体像を把握することは困難であった。"),

    H2("2. プロジェクトの目的と意義"),
    BUL("ダム基礎地質の共通性・相違点を体系的に表現し、技術者間の知識共有を促進する"),
    BUL("ダム建設未経験技術者向けの知識基盤を構築し、将来の維持管理・老朽化対応を支援する"),
    BUL("地質分類に基づく優先調査ダムリストを作成し、限られた予算・人員での効率的な安全確認を実現する"),
    BUL("国際的なダム地質評価体系（DMR等）と対応づけることで、北海道の知見を世界に発信する"),
    BUL("将来的にAI・衛星データとの統合により、自動更新可能な地質リスク管理プラットフォームへ発展させる"),
    SP(100),

    H2("3. 国際的な先行研究との比較"),
    P("ダム地質分類に関する主要な国際的体系と本プロジェクトの体系を以下に比較する。本体系の最大の独自性は、地質生成年代（絶対年代Ma）を分類軸の核心に置き、水理特性・将来リスクまで統合した点にある。", {noIndent:true}),
    SP(120),
    tblInternational(),
    SP(80),
    NOTE("DMR（Dam Mass Rating, Romana 2003）が最も本体系に近い国際的先行研究。ただし連続点数系であり、北海道特有の複合地質・地質年代軸・流域地質評価は含まれない。"),

    PB_BREAK(),

    // ══════════ Part II: Phase 1 基本分類体系 ══════════
    H1("第Ⅱ部　Phase 1 — 基本分類体系（確立済み）", C.ph1),

    H2("4. 地質時代区分と岩石種区分"),
    P("本分類体系の根幹は、地質生成年代（Ma：百万年前）を軸とした時代区分と岩石種タイプの組み合わせである。Ⅰ類は先新第三紀（> 23 Ma）の古期地質、Ⅱ類は新第三紀以降（< 23 Ma）の新期地質を表す。各大区分はさらに岩石種・地質帯によってa〜eの5サブタイプに分類される。", {noIndent:true}),
    SP(120),
    tblGeoAgeClassification(),
    SP(80),
    NOTE("年代値は国際層序委員会（ICS）2023年版準拠。北海道の地質は複数プレートの相互作用で形成されており、同一地点でも異なる時代の地質が重なる複合構造が多い。"),
    NOTE("地質年代は強度を直接規定しないが、古期ほど固結・変成が進み一般に強度高い。ただし断層・変質帯（R1・R2）では局所的弱部が発達する。"),

    H2("5. 強度コード（S）・リスク指標（R）"),
    H3("5.1 強度コード（S）— ISRM 一軸圧縮強度基準・5段階"),
    ...(()=>{const [ts,sp,tr]=tblSR(); return [ts,sp]})(),
    SP(100),
    H3("5.2 リスク指標（R）— 工学的リスク種別・6種類（複数付与可）"),
    ...(()=>{const [ts,sp,tr]=tblSR(); return [sp,tr]})(),
    SP(80),
    NOTE("Rコードは複数同時付与が可能（例：R1・R2）。括弧内では「/」で区切って表記する（例：S2/R1・R2）。"),

    H2("6. 全188基 暫定分類結果 — 統計サマリー"),
    P("2026年3月時点で全188基の暫定地質区分を完了した。以下は主要地質区分別の分布状況である。信頼度はA：7基（既存データ）、B：60基（地質図・文献推定）、C：121基（位置・水系類推）。", {noIndent:true}),
    SP(120),
    tblStatsSummary(),
    SP(80),
    NOTE("北海道最多の基礎地質はⅠ-c（砂岩泥岩互層：99基）。これは蝦夷層群タービダイトの広域分布による。2番目はⅡ-d（新第三系堆積岩：67基）で農業ダムに多い。"),
    NOTE("全188基のJSONデータ・Excel地質区分DB・地質区分付きKMZファイルは別途配布。"),

    H2("7. 信頼度評価体系"),
    P("各ダムの分類コードには信頼度ランク（A〜C・?）を付与し、情報品質を明示する。", {noIndent:true}),
    SP(80),
    new Table({width:{size:8000,type:WidthType.DXA},columnWidths:[600,1400,1800,4200],rows:[
      row([["ランク",600,{center:true}],["基数",1400,{center:true}],["情報源",1800,{center:true}],["定義・判定基準",4200,{center:true}]], true, 36),
      ...([
        ["A","7基（3.7%）","propylite.work・施工報告書","既存の透水性・岩盤試験データが存在。設計報告書・施工記録の記述に基づく"],
        ["B","60基（31.9%）","GeoNAVI地質図・文献","地質図・学術論文・一部施工記録から推定。典型的な地質帯特性を適用"],
        ["C","121基（64.4%）","位置・水系・ダム型式","ダム位置と水系・周辺地形から地質帯を類推した暫定値。要確認"],
        ["?","—","情報不足","当該コードの判定に必要な情報が存在しない。調査優先フラグ"],
      ].map(([r,cnt,src,def],i)=>row([[r,600,{center:true,bold:true,fill:["D5F5E3","D6EAF8","FEF9E7","F5F5F5"][i],color:["1E8449","2E5090","C0392B","888888"][i]}],[cnt,1400,{center:true}],[src,1800,{sm:true}],[def,4200,{xs:true}]])))
    ]}),
    SP(80),
    NOTE("Phase 1〜3を通じて全基の信頼度B以上への引き上げを目標とする。W?/G?を含む記号は「情報収集未完了」の明示的フラグとして機能する。"),

    PB_BREAK(),

    // ══════════ Part III: Phase 2 ══════════
    H1("第Ⅲ部　Phase 2 — 工学的特性の深化（新提案）", C.ph2),

    H2("8. Phase 2 の位置づけと構成"),
    P("Phase 2は2つのサブフェーズで構成される。Phase 2Aでは水理的特性コード（透水性W・基礎処理G）を新設し、Phase 1の力学特性コード（S・R）と統合する。Phase 2B・Cでは経時劣化（D）・地震応答増幅（E）・流域地質（C）・気候変動脆弱性（V）・知識蓄積段階（K）の5コードを提案する。これらはいずれも既存の国際的分類体系には存在しない独自の提案である。"),

    H2("9. Phase 2A — 透水性コード（W）・基礎処理コード（G）"),
    H3("9.1 透水性コード（W）— ルジオン値基準・5段階＋情報なし"),
    P("ルジオン値（Lu）を主指標として5段階に区分する。強度コード（S）と完全に同格・同構造で設計し、情報欠如の場合はW?を付与する。1 Lu ≈ 1.3×10⁻⁷ m/s（圧力1 MPa時）。", {noIndent:true}),
    SP(100),
    ...(()=>{const [tw,sp,tg]=tblWG(); return [tw]})(),
    SP(100),
    H3("9.2 基礎処理コード（G）— グラウチング難易度・4段階＋情報なし"),
    P("グラウチング注入量・工法の複雑さ・特殊工法の要否を4段階で区分する。ICOLD Bulletin 88（Rock Foundations for Dams）を参照基準として北海道の実情に適合させた。", {noIndent:true}),
    SP(100),
    ...(()=>{const [tw,sp,tg]=tblWG(); return [sp,tg]})(),
    SP(80),
    NOTE("WコードとGコードは密接に対応するが同一ではない。亀裂の方向性・連続性・充填状態により同じWコードでもGコードが変わる場合がある（例：W4でも空洞に連通すればG4）。"),

    H2("10. 地質区分別 W・G 標準値"),
    P("各地質区分（Ⅰ-a〜Ⅱ-e）について透水性（W）・基礎処理難易度（G）の標準的な範囲を示す。強度コード（S）との対照で読むことで、基礎工学上の見通しが一覧で把握できる。", {noIndent:true}),
    SP(120),
    tblGeoHydraulic(),
    SP(80),
    NOTE("括弧書き（例：断層帯W4、風化帯W3）は同一ダムサイト内の局部高透水帯を示す。設計では最悪部を基準に採用することが多い。"),

    PB_BREAK(),

    H2("11. Phase 2B — 次世代分類コード（D・E・C・V・K）"),
    P("以下5つのコードは、本プロジェクトが国際的な既存体系を超えて提案する独創的な分類軸である。Phase 2の後半段階（Ph.2相当）および将来の研究フェーズで実装を目指す。", {noIndent:true}),
    SP(100),
    tblNextGenCodes(),
    SP(80),
    NOTE("CコードとKコードが独自性の核心。C（流域地質）はダム基礎から流域全体への評価範囲の拡大であり、K（知識段階）は情報管理と未知領域の明示化という新しい概念。"),
    NOTE("D4（凍結融解亀裂拡大）は北海道特有のコード。年間20〜80回の凍結融解サイクルが岩盤亀裂の長期的な透水性変動に与える影響は国際的に未体系化の領域。"),

    H2("12. 完全記号体系（フルブロック）"),
    H3("12.1 全ブロック一覧とフェーズ別実装計画"),
    P("以下は分類記号の全ブロック構造と、各フェーズでの実装状況を示す。●が実装済み・計画中、●★が当該フェーズで新設するブロック。", {noIndent:true}),
    SP(100),
    tblFullSymbol(),
    SP(100),
    H3("12.2 完全記号の構成と読み方"),
    P("完全記号の構造は以下の通り。括弧内に「力学ブロック（S/R）」と「水理ブロック（W/G）」を「・」で区切って表現する。次世代ブロックは将来追加する。", {noIndent:true}),
    SP(80),
    new Table({width:{size:8500,type:WidthType.DXA},columnWidths:[4000,4500],rows:[
      row([["記号パターン",4000,{center:true}],["説明",4500,{center:true}]], true, 34),
      row([["Ⅱ-a（S3/R4・W4/G3）",4000,{bold:true,color:C.navy,sm:true}],
           ["溶結凝灰岩（中硬/冷却亀裂・高透水/多段グラウト）。Phase1+2A完全版",4500,{xs:true}]]),
      row([["Ⅰ-c（S2/R1・W?/G?）・Ⅱ-e（S5/R6・W5/G4）",4000,{bold:true,color:C.navy,sm:true}],
           ["砂岩泥岩+未固結（水理情報欠如のW?/G?あり）。W?/G?は調査優先フラグ",4500,{xs:true}]]),
      row([["Ⅰ-b（S1/R1・W1/G1）",4000,{bold:true,color:C.navy,sm:true}],
           ["花崗岩（極硬/断層リスク・極低透水/グラウト軽微）。豊平峡相当",4500,{xs:true}]]),
      row([["Ⅱ-a（S3/R4・W4/G3・D4/E2・C1・V3・K2）",4000,{bold:true,color:C.darkRed,xs:true}],
           ["将来完全版（D/E/C/V/K全ブロック）。溶結凝灰岩の北海道型フル記号例",4500,{xs:true}]]),
    ]}),
    SP(100),
    H3("12.3 完全記号 適用例（信頼度別）"),
    tblCompleteSymbolExamples(),
    SP(80),
    NOTE("信頼度AのW・Gコード確定は既存データ4基（日新・東郷・しろがね・美生）のみ。残る184基のW/GはPh.1情報収集で順次更新する。"),

    PB_BREAK(),

    // ══════════ Part IV: 作業計画 ══════════
    H1("第Ⅳ部　作業計画とフェーズ別ロードマップ", C.ph3),

    H2("13. フェーズ別詳細作業計画"),

    PHASE_LABEL(1, "情報収集・信頼度引き上げ（〜2ヶ月）", C.ph1),
    tblPhaseDetail(1, C.ph1, [
      ["propylite.work 全ダムページ精査","propylite.work（北海道の地質とダム）","各ダムの岩種・透水性・グラウチング記録シート","信頼度C→B以上への転換基数"],
      ["北海道開発局・農水省報告書収集","北海道開発局・農林水産省農業水産部","施工報告書・ルジオン試験記録","W・Gコード信頼度A確定基数"],
      ["GeoNAVI地質図照合","産総研 GeoNAVI（1:20万地質図）","各ダム周辺地質帯の確認記録","信頼度B基数の拡大"],
      ["全188基 W・G暫定コード付与","上記3ソースの統合","Excel DB更新（W/G列追加・信頼度更新）","W?/G?基数を50基以下に削減"],
    ]),
    SP(120),

    PHASE_LABEL(2, "体系確定・国際基準対応（〜2ヶ月）", C.ph2),
    tblPhaseDetail(2, C.ph2, [
      ["記号体系確定版の策定","試行基7基レビュー・専門家意見聴取","記号体系確定版ドキュメント","コードの一意性・整合性の確認"],
      ["DMR・RMR・PWRI岩盤分類との対応表","ICOLD Bulletin 88・RMR89・PWRI指針","国際基準対応表（本体系↔国際体系）","各体系との変換式または対応表"],
      ["次世代コード（D/E/C/V/K）定義確定","地震工学・気候科学・岩盤力学文献","各コードの定義書・判定フロー","コードの独自性と実用性の評価"],
      ["試行7基の完全記号（全ブロック）確定","既存データ＋体系確定版","信頼度A完全記号7基分","全ブロックの整合性確認"],
    ]),
    SP(120),

    PHASE_LABEL(3, "全188基 完全判定（〜6ヶ月）", C.ph3),
    tblPhaseDetail(3, C.ph3, [
      ["水系別一括地質判定（石狩川系88基）","Ph.1収集資料＋GeoNAVI","石狩川系全基完全記号","信頼度B以上85%以上"],
      ["水系別一括地質判定（十勝・天塩・他）","同上","十勝・天塩・他全基完全記号","同上"],
      ["優先調査対象ダム（116基）の重点確認","施工報告書・現地調査（必要に応じ）","優先対象全基のW/G信頼度A〜B化","W?/G?ゼロを目標"],
      ["地質区分付きKMZ更新","全基完全記号データ","地質区分・W/G情報付きKMZ","全188基のポップアップ情報完成"],
      ["地質区分DBの最終版作成","上記全作業の統合","Excel DB最終版（全コード・全信頼度）","全基データの品質確認"],
    ]),
    SP(120),

    PHASE_LABEL(4, "地質説明文・個別カード作成（〜3ヶ月）", C.ph4),
    tblPhaseDetail(4, C.ph4, [
      ["地質帯別概括説明文（10〜15帯）","地質区分DB・文献","地質帯別説明文（A4 1〜2頁/帯）","専門用語の平易化・図版の充実"],
      ["各ダム個別地質カード（188枚）","完全記号DB・各ダムデータ","標準フォーマットの個別カード188枚","カード間の記述一貫性の確認"],
      ["国際基準（RMR/DMR）対応ガイド","対応表（Ph.2成果）","RMR/DMR換算ガイドライン","国際論文投稿への対応可能性"],
      ["老朽化リスクマトリクス作成","地質区分×建設年代×Dコード","優先再評価ダムリスト（上位30基）","リスト根拠の明示・専門家レビュー"],
    ]),
    SP(120),

    PHASE_LABEL(5, "成果物取りまとめ・公開（〜2ヶ月）", C.ph5),
    tblPhaseDetail(5, C.ph5, [
      ["最終報告書作成","全フェーズ成果の統合","北海道ダム地質分類最終報告書","報告書の網羅性・引用可能性"],
      ["オープンデータDB整備","Excel DB最終版","公開可能形式のCSV・JSONデータ","ライセンス・プライバシー確認"],
      ["GIS・地図整備","KMZ最終版・産総研連携","地質区分付きGISデータ（Shapefile等）","GIS操作者への引き渡し確認"],
      ["教育用資料・研修テキスト","地質帯別説明文・個別カード","研修用テキスト・スライド","未経験技術者への説明可能性テスト"],
      ["国際学会・論文投稿計画","最終報告書・国際対応表","ICOLD・地盤工学会向け論文草稿","査読対応可能な品質確認"],
    ]),
    SP(100),

    H2("14. 優先調査対象ダム"),
    P("Ph.1〜Ph.3を通じて、以下の地質群に属するダムを優先的に調査・確認する。", {noIndent:true}),
    SP(100),
    tblPriorityDams(),
    SP(80),
    NOTE("優先度の高い順：溶結凝灰岩帯（G3・高難度）→石灰岩帯（G4・カルスト）→蛇紋岩体（G4・特殊）→老朽化ダム（D4〜D5・劣化）→農業フィルダム（W4〜W5・未固結）。"),

    H2("15. 参照資料・データ収集戦略"),
    H3("15.1 一次資料"),
    BUL("propylite.work：北海道の地質とダム（道央・道南・道北・道東編）— 最も重要な一次資料。個別ダムページはアクセス制限あり、管理者への照会要"),
    BUL("産総研 GeoNAVI（https://gbank.gsj.jp/geonavi/）：1:20万地質図デジタル版。各ダム周辺の地質帯確認"),
    BUL("北海道開発局各ダム設計・施工報告書：ルジオン試験値・グラウチング量の最重要データ源"),
    BUL("農林水産省農業水産部農業ダム施工記録：農業フィルダム（約56基）の透水性・基礎処理情報"),
    H3("15.2 国際基準・参考文献"),
    BUL("ICOLD Bulletin 88（Rock Foundations for Dams）・Bulletin 111：グラウチング・岩盤評価の国際標準"),
    BUL("Romana（2003）DMR (Dam Mass Rating)：最も近い国際先行体系。ISRM 10th Congress"),
    BUL("Fell et al.（2015）Geotechnical Engineering of Dams：フィルダム基礎の標準テキスト"),
    BUL("Bieniawski（1989）Engineering Rock Mass Classifications：RMR89の原典"),
    BUL("ICS International Chronostratigraphic Chart 2023：地質年代値の基準"),
    BUL("Weaver & Bruce（2007）Dam Foundation Grouting, ASCE：グラウチング技術の集大成"),

    PB_BREAK(),

    // ══════════ Part V: 将来展望 ══════════
    H1("第Ⅴ部　将来展望", C.ph5),

    H2("16. AIとの協働（K5コードの実装）"),
    P("知識段階コード（K）のK5は「AIによる推定」を明示するフラグである。将来的には以下の技術統合を目指す。"),
    BUL("教師データ：信頼度AのダムデータとGeoNAVI地質図を用いた機械学習モデルの構築"),
    BUL("入力：ダムの位置（緯度経度）・標高・水系・ダム型式・建設年代"),
    BUL("出力：地質区分記号（Ⅰ/Ⅱ-a〜e）・S/R/W/G推定コードと確信度スコア"),
    BUL("InSAR（合成開口レーダー干渉）データとの統合：地表変位計測×地質リスク（R1・R3）の重ね合わせ"),
    BUL("衛星マルチスペクトル画像による変質帯（R2）の自動検出"),
    P("この「AIと人間の共同調査プラットフォーム」が実現すれば、K5ラベルが付いたAI推定コードを現地調査・報告書参照で順次K1〜K4に更新していく継続的な精度向上サイクルが確立される。"),

    H2("17. 他地域・国際展開の可能性"),
    P("本プロジェクトの分類体系は北海道固有の地質に最適化されているが、その思想と構造は普遍的に適用可能である。"),
    BUL("東北・中部地方への展開：火山性地質（Ⅱ-a・Ⅱ-b）・変成帯（Ⅰ-a）が類似する地域への適用"),
    BUL("ICOLD（国際大ダム会議）での発表：日本から提案する新しいダム地質分類体系として国際発信"),
    BUL("アジア地域への展開：東南アジアの新第三系（Ⅱ-d）・火山帯（Ⅱ-a・Ⅱ-b）への適用可能性"),
    BUL("気候変動脆弱性コード（V）は全世界のダムに適用可能な普遍的フレームワーク"),

    H2("18. 北海道が先導する意義"),
    P("北海道は日本で最も多様な地質を持つ地域の一つであり、188基のダムはその地質的多様性を反映している。本プロジェクトが完成することで、以下の社会的価値が実現する。"),
    BUL("老朽化ダムの科学的優先評価：地質コード（D・V）×建設年代のマトリクスにより、限られた予算で最もリスクの高いダムから順に再評価できる"),
    BUL("ダム管理技術の伝承：本体系は文書化された知識基盤として、ダム建設未経験技術者への技術継承を促進する"),
    BUL("防災・減災への貢献：洪水・地震・火山噴火等の自然災害時に、地質リスクの高いダムを即座に特定できる体制の整備"),
    BUL("北海道のダム技術の国際発信：地質年代軸・流域地質評価・AI統合という国際的に先例のない体系の確立"),
    SP(200),
    DIV(),
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:200,after:80},
      children:[new TextRun({text:"以　上",font:"游明朝",size:22})]}),
    new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:60,after:0},
      children:[new TextRun({text:"2026年3月　北海道ダム地質分類プロジェクト　全体計画書",font:"游明朝",size:20,color:"555555"})]}),
  ];

  return new Document({
    styles:{
      default:{document:{run:{font:"游明朝",size:22}}},
      paragraphStyles:[
        {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",
          run:{size:38,bold:true,font:"游明朝",color:C.navy},
          paragraph:{spacing:{before:520,after:220},outlineLevel:0}},
        {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",
          run:{size:30,bold:true,font:"游明朝",color:C.darkBlue},
          paragraph:{spacing:{before:340,after:160},outlineLevel:1}},
        {id:"Heading3",name:"Heading 3",basedOn:"Normal",next:"Normal",
          run:{size:25,bold:true,font:"游明朝",color:C.blue},
          paragraph:{spacing:{before:240,after:110},outlineLevel:2}},
        {id:"Heading4",name:"Heading 4",basedOn:"Normal",next:"Normal",
          run:{size:23,bold:true,font:"游明朝",color:C.darkBlue},
          paragraph:{spacing:{before:180,after:80},outlineLevel:3}},
      ]
    },
    numbering: NUMBERING,
    sections:[{
      properties:{
        page:{size:{width:11906,height:16838},margin:{top:1700,right:1500,bottom:1700,left:1800}}
      },
      footers:{default:new Footer({children:[
        new Paragraph({alignment:AlignmentType.CENTER,children:[
          new TextRun({text:"北海道ダム地質分類プロジェクト　全体計画書（フルプラン）　",font:"游明朝",size:18,color:C.gray}),
          new TextRun({children:[PageNumber.CURRENT],font:"游明朝",size:18,color:C.gray}),
          new TextRun({text:" / ",font:"游明朝",size:18,color:C.gray}),
          new TextRun({children:[PageNumber.TOTAL_PAGES],font:"游明朝",size:18,color:C.gray}),
        ]})
      ]})},
      children
    }]
  });
}

// 実行
const doc = buildDoc();
Packer.toBuffer(doc).then(buf=>{
  fs.writeFileSync('/home/claude/北海道ダム地質分類_全体計画書_フルプラン.docx', buf);
  console.log('完成');
}).catch(e=>{console.error(e);process.exit(1);});
