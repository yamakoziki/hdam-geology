// ======================================================
// 北海道ダム地質分類 全体計画書（フルプラン）
// Part 1: ユーティリティ・スタイル定義
// ======================================================
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, PageNumber, PageBreak, Footer, Header,
  ExternalHyperlink
} = require('docx');
const fs = require('fs');

// ── 罫線定義 ──────────────────────────────────────────
const B = {
  none: { style:BorderStyle.NONE, size:0, color:"FFFFFF" },
  thin: c => ({ style:BorderStyle.SINGLE, size:2, color:c||"CCCCCC" }),
  mid:  c => ({ style:BorderStyle.SINGLE, size:4, color:c||"4472C4" }),
  thick:c => ({ style:BorderStyle.SINGLE, size:8, color:c||"1F3864" }),
};
const bds = (b) => ({top:b,bottom:b,left:b,right:b});
const BNONE = bds(B.none);
const BTHIN = bds(B.thin());
const BMID  = c => bds(B.mid(c));
const BTHICK= c => bds(B.thick(c));

// ── カラーパレット ────────────────────────────────────
const C = {
  navy:   "1F3864",  darkBlue: "2E5090", blue: "4472C4",  lightBlue:"DBE5F1",
  red:    "C0392B",  darkRed:  "8B0000", amber:"F39C12",  lightAmber:"FEF9E7",
  green:  "1E8449",  darkGreen:"145A32",
  purple: "6C3483",  darkPurple:"4A235A",
  gray:   "888888",  lightGray:"F5F5F5", midGray:"DDDDDD",
  white:  "FFFFFF",  black:    "000000",
  // Phase colors
  ph1:    "2E5090",  ph2:"8B4513",  ph3:"1E8449",  ph4:"C0392B",  ph5:"6C3483",
  // Row alternates
  row0:   "F2F7FB",  row1:"FFFFFF",
};

// ── セル ────────────────────────────────────────────
function cell(text, w, o={}) {
  const sz = o.xs?17 : o.sm?18 : o.isHdr?20 : (o.sz||20);
  return new TableCell({
    columnSpan:  o.span||1,
    width:       {size:w, type:WidthType.DXA},
    verticalAlign: o.va||VerticalAlign.CENTER,
    borders:     o.noBdr ? BNONE : (o.isHdr ? BMID(o.hdrColor||C.darkBlue) : BTHIN),
    shading:     o.fill ? {fill:o.fill, type:ShadingType.CLEAR} : undefined,
    margins:     {top:90,bottom:90,left:160,right:160},
    children: [new Paragraph({
      alignment: o.center ? AlignmentType.CENTER : (o.right ? AlignmentType.RIGHT : AlignmentType.LEFT),
      spacing:{before:0,after:0},
      children:[new TextRun({
        text, font:"游明朝", size:sz,
        bold:    o.isHdr||o.bold||false,
        italics: o.ita||false,
        color:   o.isHdr ? C.white : (o.color||C.black),
      })]
    })]
  });
}

function row(cells, isHdr=false, h=null) {
  return new TableRow({
    tableHeader:isHdr,
    height: h?{value:h,rule:"exact"}:undefined,
    children: cells.map(([t,w,o={}])=>cell(t,w,{...o,isHdr}))
  });
}

// セクション区切り行（色帯）
function secRow(label, fill, span=8, color="FFFFFF", sz=19) {
  return new TableRow({children:[new TableCell({
    columnSpan:span,
    borders: BMID(fill),
    shading:{fill,type:ShadingType.CLEAR},
    margins:{top:80,bottom:80,left:200,right:160},
    children:[new Paragraph({children:[
      new TextRun({text:label,font:"游明朝",size:sz,bold:true,color})
    ]})]
  })]});
}

// ── テキスト要素 ─────────────────────────────────────
function H1(text, color=C.navy) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing:{before:520,after:220},
    border:{bottom:{style:BorderStyle.SINGLE,size:10,color,space:1}},
    children:[new TextRun({text,font:"游明朝",size:38,bold:true,color})]
  });
}
function H2(text, color=C.darkBlue) {
  return new Paragraph({
    heading:HeadingLevel.HEADING_2,
    spacing:{before:340,after:160},
    children:[new TextRun({text,font:"游明朝",size:30,bold:true,color})]
  });
}
function H3(text, color=C.blue) {
  return new Paragraph({
    heading:HeadingLevel.HEADING_3,
    spacing:{before:240,after:110},
    children:[new TextRun({text,font:"游明朝",size:25,bold:true,color})]
  });
}
function H4(text, color=C.darkBlue) {
  return new Paragraph({
    heading:HeadingLevel.HEADING_4,
    spacing:{before:180,after:80},
    children:[new TextRun({text,font:"游明朝",size:23,bold:true,color})]
  });
}
function P(text, o={}) {
  return new Paragraph({
    spacing:{before:80,after:80,line:390},
    indent:{firstLine: o.noIndent?0:440},
    children:[new TextRun({text,font:"游明朝",size:o.sz||22,...o})]
  });
}
function PB(runs) {  // 複数TextRunの段落
  return new Paragraph({
    spacing:{before:80,after:80,line:390},
    indent:{firstLine:440},
    children:runs.map(([t,o={}])=>new TextRun({text:t,font:"游明朝",size:22,...o}))
  });
}
function NOTE(text) {
  return new Paragraph({
    spacing:{before:60,after:60,line:340},
    indent:{left:380},
    children:[
      new TextRun({text:"※ ",font:"游明朝",size:18,color:C.gray}),
      new TextRun({text,font:"游明朝",size:18,color:"444444"})
    ]
  });
}
function SP(n=120){return new Paragraph({spacing:{before:n,after:0},children:[]});}
function PB_BREAK(){return new Paragraph({children:[new PageBreak()]});}
function DIV(color=C.blue){
  return new Paragraph({
    spacing:{before:140,after:140},
    border:{bottom:{style:BorderStyle.SINGLE,size:6,color,space:1}},
    children:[]
  });
}
function PHASE_LABEL(n, label, color) {
  return new Paragraph({
    spacing:{before:300,after:140},
    shading:{fill:color,type:ShadingType.CLEAR},
    indent:{left:200,right:200},
    border:{
      top:{style:BorderStyle.SINGLE,size:6,color},
      bottom:{style:BorderStyle.SINGLE,size:6,color},
      left:{style:BorderStyle.SINGLE,size:20,color},
      right:{style:BorderStyle.SINGLE,size:6,color},
    },
    children:[
      new TextRun({text:`Phase ${n}  `, font:"游明朝",size:28,bold:true,color:C.white}),
      new TextRun({text:label,font:"游明朝",size:26,bold:true,color:C.white}),
    ]
  });
}
// ビュレット段落（箇条書き）
function BUL(text, ref="b1", lv=0) {
  return new Paragraph({
    numbering:{reference:ref,level:lv},
    spacing:{before:60,after:60,line:360},
    children:[new TextRun({text,font:"游明朝",size:22})]
  });
}
function BUL2(text) { return BUL(text,"b2",1); }

module.exports = {
  C,B,bds,BNONE,BTHIN,BMID,BTHICK,
  cell,row,secRow,H1,H2,H3,H4,P,PB,NOTE,SP,PB_BREAK,DIV,PHASE_LABEL,BUL,BUL2
};
