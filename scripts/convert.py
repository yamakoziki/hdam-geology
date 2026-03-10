#!/usr/bin/env python3
"""
北海道ダム地質区分DB - Excel → スタンドアロン HTML 生成スクリプト

使用方法:
  python3 scripts/convert.py
  python3 scripts/convert.py --input data/北海道ダム地質分類DB.xlsx
  python3 scripts/convert.py --input data/北海道ダム地質分類DB.xlsx --sheet ダム地質区分DB
  python3 scripts/convert.py --input data/北海道ダム地質分類DB.xlsx --output docs/index.html
"""

import argparse
import json
import sys
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas が必要です。")
    print("  python3 -m pip install pandas openpyxl")
    sys.exit(1)


def build_html(json_data: str, rows: int, source: str) -> str:
    return """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>北海道ダム地質区分DB 検索システム</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700&family=Share+Tech+Mono&display=swap');
:root{--bg:#0d1117;--surface:#161b22;--surface2:#21262d;--surface3:#2d333b;--border:#30363d;--accent:#58a6ff;--accent2:#3fb950;--accent3:#f78166;--accent4:#d2a8ff;--text:#e6edf3;--text2:#8b949e;--text3:#6e7681;--warn:#e3b341;--font:'Noto Sans JP',sans-serif;--mono:'Share Tech Mono',monospace}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh;font-size:13px}
.header{background:linear-gradient(135deg,#0d1117,#161b22,#0d1117);border-bottom:1px solid var(--border);padding:14px 20px;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:100}
.hi{width:34px;height:34px;background:linear-gradient(135deg,var(--accent),var(--accent4));border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0}
.ht{font-size:15px;font-weight:700;letter-spacing:.05em}.hs{font-size:10px;color:var(--text2);font-family:var(--mono)}
.hc{margin-left:auto;background:var(--surface2);border:1px solid var(--border);border-radius:20px;padding:4px 14px;font-family:var(--mono);font-size:12px;color:var(--accent2);white-space:nowrap}
.layout{display:grid;grid-template-columns:310px 1fr;height:calc(100vh - 63px)}
.sidebar{background:var(--surface);border-right:1px solid var(--border);overflow-y:auto;padding:12px;display:flex;flex-direction:column;gap:10px}
.sidebar::-webkit-scrollbar{width:4px}.sidebar::-webkit-scrollbar-thumb{background:var(--border);border-radius:2px}
.sec{border:1px solid var(--border);border-radius:8px;overflow:hidden}
.sh{background:var(--surface2);padding:7px 10px;font-size:10px;font-weight:700;color:var(--text2);letter-spacing:.1em;text-transform:uppercase;display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none}
.sh:hover{color:var(--text)}.sb{padding:10px;display:flex;flex-direction:column;gap:8px}.sb.col{display:none}
.arr{margin-left:auto;font-size:10px;transition:transform .2s}.arr.op{transform:rotate(180deg)}
.lt{display:flex;background:var(--surface3);border:1px solid var(--border);border-radius:6px;overflow:hidden}
.lb{flex:1;padding:6px;text-align:center;font-size:12px;font-weight:700;cursor:pointer;transition:all .15s;color:var(--text2);border:none;background:transparent;font-family:var(--mono)}
.lb.aa{background:var(--accent);color:var(--bg)}.lb.ao{background:var(--accent3);color:var(--bg)}
.lb:hover:not(.aa):not(.ao){background:var(--surface2);color:var(--text)}
.fl{font-size:11px;color:var(--text2);margin-bottom:3px;display:flex;align-items:center;gap:5px}
.fl .dot{width:5px;height:5px;border-radius:50%;background:var(--accent);flex-shrink:0}
.ti{width:100%;background:var(--surface3);border:1px solid var(--border);border-radius:6px;padding:5px 9px;color:var(--text);font-family:var(--font);font-size:12px;outline:none;transition:border-color .15s}
.ti:focus{border-color:var(--accent)}.ti::placeholder{color:var(--text3)}
.hint{font-size:10px;color:var(--text3);font-family:var(--mono);margin-top:2px}.hint em{color:var(--warn);font-style:normal}
.cg{display:flex;flex-wrap:wrap;gap:4px}
.ch{padding:3px 8px;border-radius:12px;border:1px solid var(--border);background:var(--surface3);color:var(--text2);cursor:pointer;font-size:11px;transition:all .12s;user-select:none;font-family:var(--mono)}
.ch:hover{border-color:var(--accent);color:var(--text)}.ch.s{background:var(--accent);border-color:var(--accent);color:var(--bg);font-weight:700}
.cg2{display:grid;grid-template-columns:1fr 1fr;gap:4px}
.cc{padding:4px 6px;border-radius:6px;border:1px solid var(--border);background:var(--surface3);color:var(--text2);cursor:pointer;font-size:11px;transition:all .12s;user-select:none;text-align:center}
.cc:hover{border-color:var(--accent2);color:var(--text)}.cc.s{background:var(--surface2);border-color:var(--accent2);color:var(--accent2);font-weight:600}
.br{display:flex;gap:6px}
.btn{flex:1;padding:8px;border-radius:6px;border:1px solid var(--border);cursor:pointer;font-size:12px;font-weight:700;font-family:var(--font);transition:all .15s}
.bts{background:linear-gradient(135deg,var(--accent),#1a73e8);border-color:var(--accent);color:#fff}
.bts:hover{opacity:.9;transform:translateY(-1px)}
.btr{background:var(--surface3);color:var(--text2)}.btr:hover{border-color:var(--accent3);color:var(--accent3)}
.main{overflow:auto;padding:14px;display:flex;flex-direction:column;gap:10px}
.main::-webkit-scrollbar{width:6px;height:6px}.main::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}
.stb{display:flex;align-items:center;gap:10px;padding:7px 14px;background:var(--surface);border:1px solid var(--border);border-radius:8px;font-family:var(--mono);font-size:12px;flex-shrink:0}
.sn{color:var(--accent2);font-size:17px;font-weight:700}.sl{color:var(--text2)}
.slg{margin-left:auto;padding:2px 10px;border-radius:10px;font-size:11px;font-weight:700}
.and{background:rgba(88,166,255,.15);border:1px solid var(--accent);color:var(--accent)}
.or{background:rgba(247,129,102,.15);border:1px solid var(--accent3);color:var(--accent3)}
.tw{border:1px solid var(--border);border-radius:8px;overflow:auto;background:var(--surface);flex:1}
table{width:100%;border-collapse:collapse;font-size:12px}
thead{position:sticky;top:0;z-index:10;background:var(--surface2)}
th{padding:8px 10px;text-align:left;font-size:10px;font-weight:700;color:var(--text2);letter-spacing:.08em;text-transform:uppercase;border-bottom:2px solid var(--border);white-space:nowrap}
td{padding:7px 10px;border-bottom:1px solid var(--border);vertical-align:top;max-width:240px;word-break:break-all}
tr:hover td{background:rgba(255,255,255,.02)}tr:last-child td{border-bottom:none}
.badge{display:inline-block;padding:1px 7px;border-radius:10px;font-size:11px;font-weight:700;font-family:var(--mono)}
.btp{background:rgba(88,166,255,.15);color:var(--accent);border:1px solid rgba(88,166,255,.3)}
.bto{background:rgba(210,168,255,.15);color:var(--accent4);border:1px solid rgba(210,168,255,.3)}
.btn2{background:rgba(63,185,80,.15);color:var(--accent2);border:1px solid rgba(63,185,80,.3)}
.bta{background:rgba(63,185,80,.2);color:var(--accent2);border:1px solid rgba(63,185,80,.4)}
.btb{background:rgba(227,179,65,.2);color:var(--warn);border:1px solid rgba(227,179,65,.4)}
.btc{background:rgba(247,129,102,.2);color:var(--accent3);border:1px solid rgba(247,129,102,.4)}
.bts2{background:rgba(88,166,255,.1);color:#79c0ff;border:1px solid rgba(88,166,255,.2)}
.btrk{background:rgba(247,129,102,.1);color:#ffa198;border:1px solid rgba(247,129,102,.2)}
.hl{background:rgba(227,179,65,.25);color:var(--warn);border-radius:2px;padding:0 2px}
.empty{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:60px 20px;color:var(--text3);gap:12px;text-align:center}
.empty .eic{font-size:48px;opacity:.3}
.rr{display:flex;align-items:center;gap:6px}
.ri{background:var(--surface3);border:1px solid var(--border);border-radius:6px;padding:5px 8px;color:var(--text);font-family:var(--mono);font-size:12px;outline:none;width:80px}
.ri:focus{border-color:var(--accent)}.rs{color:var(--text3);font-size:12px}
.cab{font-size:10px;color:var(--text3);cursor:pointer;padding:2px 6px;border-radius:4px;border:1px solid transparent;background:transparent;margin-left:auto;font-family:var(--font)}
.cab:hover{border-color:var(--border);color:var(--text2)}
.src-badge{background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:6px 10px;font-size:10px;color:var(--text3);font-family:var(--mono);display:flex;align-items:center;gap:6px}
.src-badge em{color:var(--accent2);font-style:normal}
</style>
</head>
<body>
<div class="header">
  <div class="hi">🏔</div>
  <div>
    <div class="ht">北海道ダム地質区分DB 検索システム</div>
    <div class="hs">Hokkaido Dam Geology Classification Database · """ + str(rows) + " ダム · " + source + """</div>
  </div>
  <div class="hc" id="hdrCount">全 """ + str(rows) + """ 件</div>
</div>
<div class="layout">
  <div class="sidebar" id="sidebar"></div>
  <div class="main">
    <div class="stb">
      <span class="sn" id="rNum">""" + str(rows) + """</span>
      <span class="sl">件のダムが見つかりました</span>
      <span class="slg and" id="lgBadge">AND 検索</span>
    </div>
    <div class="tw"><div id="tbl"></div></div>
  </div>
</div>
<script>
const DATA=""" + json_data + """;
const RECS=DATA.records,UNI=DATA.uniques,META=DATA.meta;
const COLS=[{k:'仮No',l:'No'},{k:'ダム名',l:'ダム名'},{k:'水系名',l:'水系名'},{k:'河川名',l:'河川名'},{k:'所在地',l:'所在地'},{k:'型式',l:'型式'},{k:'目的',l:'目的'},{k:'堤高(m)',l:'堤高(m)'},{k:'完成年度',l:'完成年'},{k:'古期_区分コード',l:'古期区分'},{k:'古期_年代名',l:'古期年代'},{k:'古期_岩石種',l:'古期岩石種'},{k:'古期_強度',l:'古期強度'},{k:'古期_リスク',l:'古期リスク'},{k:'新期_区分コード',l:'新期区分'},{k:'新期_年代名',l:'新期年代'},{k:'新期_岩石種',l:'新期岩石種'},{k:'新期_強度',l:'新期強度'},{k:'新期_リスク',l:'新期リスク'},{k:'分類記号（完全）',l:'分類記号'},{k:'信頼度',l:'信頼度'},{k:'判定根拠・参照文献',l:'判定根拠'}];
const ST={logic:'AND',f:mkF(),disp:['ダム名','水系名','所在地','型式','目的','堤高(m)','完成年度','古期_区分コード','古期_岩石種','古期_強度','古期_リスク','新期_区分コード','新期_岩石種','新期_強度','新期_リスク','分類記号（完全）','信頼度'],col:{}};
function mkF(){return{ダム名:'',水系名:'',河川名:'',所在地:'',型式:[],目的:[],堤高min:'',堤高max:'',完成年min:'',完成年max:'',古期_区分コード:[],古期_年代名:[],古期_岩石種:[],古期_強度:[],古期_リスク:[],新期_区分コード:[],新期_年代名:[],新期_岩石種:[],新期_強度:[],新期_リスク:[],分類記号:'',信頼度:[]};}
function wm(pat,str){if(!pat)return true;const p=pat.toLowerCase(),s=(str||'').toLowerCase();if(!p.includes('*'))return s.includes(p);const parts=p.split('*');let idx=0;for(let i=0;i<parts.length;i++){const pt=parts[i];if(!pt)continue;const f=s.indexOf(pt,idx);if(f===-1)return false;if(i===0&&p[0]!=='*'&&f!==0)return false;idx=f+pt.length;}if(p[p.length-1]!=='*'){const last=parts[parts.length-1];if(last&&!s.endsWith(last))return false;}return true;}
function cm(arr,val){return!arr||arr.length===0||arr.includes(val);}
function filter(){const f=ST.f;return RECS.filter(r=>{const c=[];if(f.ダム名)c.push(wm(f.ダム名,r.ダム名));if(f.水系名)c.push(wm(f.水系名,r.水系名));if(f.河川名)c.push(wm(f.河川名,r.河川名));if(f.所在地)c.push(wm(f.所在地,r.所在地));if(f.型式.length)c.push(cm(f.型式,r.型式));if(f.目的.length)c.push(cm(f.目的,r.目的));if(f.堤高min!=='')c.push(parseFloat(r['堤高(m)'])>=parseFloat(f.堤高min));if(f.堤高max!=='')c.push(parseFloat(r['堤高(m)'])<=parseFloat(f.堤高max));if(f.完成年min!=='')c.push(parseFloat(r.完成年度)>=parseFloat(f.完成年min));if(f.完成年max!=='')c.push(parseFloat(r.完成年度)<=parseFloat(f.完成年max));if(f.古期_区分コード.length)c.push(cm(f.古期_区分コード,r.古期_区分コード));if(f.古期_年代名.length)c.push(cm(f.古期_年代名,r.古期_年代名));if(f.古期_岩石種.length)c.push(cm(f.古期_岩石種,r.古期_岩石種));if(f.古期_強度.length)c.push(cm(f.古期_強度,r.古期_強度));if(f.古期_リスク.length)c.push(cm(f.古期_リスク,r.古期_リスク));if(f.新期_区分コード.length)c.push(cm(f.新期_区分コード,r.新期_区分コード));if(f.新期_年代名.length)c.push(cm(f.新期_年代名,r.新期_年代名));if(f.新期_岩石種.length)c.push(cm(f.新期_岩石種,r.新期_岩石種));if(f.新期_強度.length)c.push(cm(f.新期_強度,r.新期_強度));if(f.新期_リスク.length)c.push(cm(f.新期_リスク,r.新期_リスク));if(f.分類記号)c.push(wm(f.分類記号,r['分類記号（完全）']));if(f.信頼度.length)c.push(cm(f.信頼度,r.信頼度));if(!c.length)return true;return ST.logic==='AND'?c.every(Boolean):c.some(Boolean);});}
function hl(text,pat){if(!pat||!text)return text||'';const clean=pat.replace(/\\*/g,'');if(!clean)return text;try{const re=new RegExp('('+clean.replace(/[.*+?^${}()|[\\]\\\\]/g,'\\\\$&')+')','gi');return text.replace(re,'<span class="hl">$1</span>');}catch{return text;}}
function cell(r,k){const v=r[k]||'';if(k==='型式')return v?`<span class="badge btp">${v}</span>`:'';if(k==='古期_区分コード')return v?`<span class="badge bto">${v}</span>`:'';if(k==='新期_区分コード')return v?`<span class="badge btn2">${v}</span>`:'';if(k==='信頼度'){const c2=v==='A'?'bta':v==='B'?'btb':'btc';return v?`<span class="badge ${c2}">${v}</span>`:'';}if(k==='古期_強度'||k==='新期_強度')return v?`<span class="badge bts2">${v}</span>`:'';if(k==='古期_リスク'||k==='新期_リスク')return v?`<span class="badge btrk">${v}</span>`:'';if(k==='ダム名')return hl(v,ST.f.ダム名);if(k==='水系名')return hl(v,ST.f.水系名);if(k==='河川名')return hl(v,ST.f.河川名);if(k==='所在地')return hl(v,ST.f.所在地);if(k==='分類記号（完全）')return hl(v,ST.f.分類記号);return v;}
function renderTable(recs){const cols=COLS.filter(c=>ST.disp.includes(c.k));if(!recs.length)return'<div class="empty"><div class="eic">🔍</div><div>該当するダムが見つかりませんでした</div></div>';let h='<table><thead><tr>'+cols.map(c=>`<th>${c.l}</th>`).join('')+'</tr></thead><tbody>';recs.forEach(r=>{h+='<tr>'+cols.map(c=>`<td>${cell(r,c.k)}</td>`).join('')+'</tr>';});return h+'</tbody></table>';}
function chips(label,ukey,fkey){const vals=UNI[ukey]||[],sel=ST.f[fkey]||[];let h=`<div><div class="fl"><span class="dot"></span>${label}<button class="cab" onclick="clr('${fkey}')">クリア</button></div><div class="cg">`;vals.forEach(v=>{const s=sel.includes(v);h+=`<span class="ch ${s?'s':''}" data-fkey="${fkey}" data-val="${v.replace(/"/g,'&quot;')}" onclick="tog(this)">${v}</span>`;});return h+'</div></div>';}
function txt(label,fkey,ph,wc){const v=(ST.f[fkey]||'').replace(/"/g,'&quot;');let h=`<div><div class="fl"><span class="dot"></span>${label}</div><input class="ti" type="text" value="${v}" placeholder="${ph}" oninput="upd('${fkey}',this.value)">`;if(wc)h+='<div class="hint">ワイルドカード: <em>*</em> 使用可　例: <em>Ⅱ-b*</em>　<em>*S3*</em>　<em>*R4*</em></div>';return h+'</div>';}
function numRange(label,kMin,kMax){return`<div><div class="fl"><span class="dot"></span>${label}</div><div class="rr"><input class="ri" type="number" placeholder="最小" value="${ST.f[kMin]}" oninput="upd('${kMin}',this.value)"><span class="rs">〜</span><input class="ri" type="number" placeholder="最大" value="${ST.f[kMax]}" oninput="upd('${kMax}',this.value)"></div></div>`;}
function sec(id,title,body){const cl=ST.col[id];return`<div class="sec"><div class="sh" onclick="tsec('${id}')">${title}<span class="arr ${cl?'':'op'}" id="ar-${id}">▼</span></div><div class="sb ${cl?'col':''}" id="sb-${id}">${body}</div></div>`;}
function render(){let h='';h+=`<div class="src-badge">📄 <em>${META.source}</em> · シート: <em>${META.sheet}</em> · <em>${META.rows}</em>件</div>`;h+=`<div class="sec"><div class="sb"><div style="font-size:11px;color:var(--text2);margin-bottom:4px;font-weight:700">条件の論理演算</div><div class="lt"><button class="lb ${ST.logic==='AND'?'aa':''}" onclick="slg('AND')">AND　すべて一致</button><button class="lb ${ST.logic==='OR'?'ao':''}" onclick="slg('OR')">OR　いずれか一致</button></div></div></div>`;
let b1=txt('ダム名','ダム名','ダム名を入力...',false)+txt('水系名','水系名','水系名を入力...',false)+txt('河川名','河川名','河川名を入力...',false)+txt('所在地','所在地','市町村名...',false)+chips('型式','型式','型式')+chips('目的','目的','目的')+numRange('堤高 (m)','堤高min','堤高max')+numRange('完成年度','完成年min','完成年max');h+=sec('b1','📍 基本情報',b1);
let b2=chips('区分コード','古期_区分コード','古期_区分コード')+chips('年代名','古期_年代名','古期_年代名')+chips('岩石種','古期_岩石種','古期_岩石種')+chips('強度','古期_強度','古期_強度')+chips('リスク','古期_リスク','古期_リスク');h+=sec('b2','🪨 古期地質',b2);
let b3=chips('区分コード','新期_区分コード','新期_区分コード')+chips('年代名','新期_年代名','新期_年代名')+chips('岩石種','新期_岩石種','新期_岩石種')+chips('強度','新期_強度','新期_強度')+chips('リスク','新期_リスク','新期_リスク');h+=sec('b3','🌋 新期地質',b3);
let b4=txt('分類記号（完全）','分類記号','例: Ⅱ-b*　または　*S3*',true)+chips('信頼度','信頼度','信頼度');h+=sec('b4','📋 分類記号・信頼度',b4);
let b5='<div class="cg2">';COLS.forEach(c=>{const s=ST.disp.includes(c.k);b5+=`<span class="cc ${s?'s':''}" data-ckey="${c.k}" onclick="tcol(this)">${c.l}</span>`;});b5+='</div>';h+=sec('b5','👁 表示列の選択',b5);
h+=`<div class="br"><button class="btn bts" onclick="upd2()">🔍 検索実行</button><button class="btn btr" onclick="rst()">↺ リセット</button></div>`;
document.getElementById('sidebar').innerHTML=h;upd2();}
function upd2(){const recs=filter(),total=RECS.length;document.getElementById('rNum').textContent=recs.length;document.getElementById('hdrCount').textContent=recs.length+' / '+total+' 件';document.getElementById('lgBadge').textContent=ST.logic==='AND'?'AND 検索':'OR 検索';document.getElementById('lgBadge').className='slg '+(ST.logic==='AND'?'and':'or');document.getElementById('tbl').innerHTML=renderTable(recs);}
window.slg=l=>{ST.logic=l;document.querySelectorAll('.lb').forEach(b=>{b.classList.toggle('aa',b.textContent.startsWith('AND')&&l==='AND');b.classList.toggle('ao',b.textContent.startsWith('OR')&&l==='OR');});upd2();};
window.tog=el=>{const k=el.dataset.fkey,v=el.dataset.val;if(!ST.f[k])ST.f[k]=[];const i=ST.f[k].indexOf(v);i===-1?(ST.f[k].push(v),el.classList.add('s')):(ST.f[k].splice(i,1),el.classList.remove('s'));upd2();};
window.clr=k=>{ST.f[k]=[];render();};
window.upd=(k,v)=>{ST.f[k]=v;upd2();};
window.tcol=el=>{const k=el.dataset.ckey;const i=ST.disp.indexOf(k);i===-1?(ST.disp.push(k),el.classList.add('s')):(ST.disp.splice(i,1),el.classList.remove('s'));upd2();};
window.tsec=id=>{ST.col[id]=!ST.col[id];document.getElementById('sb-'+id).classList.toggle('col');document.getElementById('ar-'+id).classList.toggle('op');};
window.rst=()=>{ST.f=mkF();ST.logic='AND';render();};
window.upd2=upd2;
render();
</script>
</body>
</html>"""


def convert(input_path: str, sheet_name: str, output_path: str) -> None:
    print(f"読み込み中: {input_path}  シート: {sheet_name}")

    xl = pd.ExcelFile(input_path)
    if sheet_name not in xl.sheet_names:
        print(f"ERROR: シート '{sheet_name}' が見つかりません。")
        print(f"利用可能なシート: {xl.sheet_names}")
        sys.exit(1)

    df = pd.read_excel(input_path, sheet_name=sheet_name, header=0)
    print(f"  → {len(df)} 行, {len(df.columns)} 列 を読み込みました")

    col_map = {col: col.replace("\n", "_") for col in df.columns}
    df = df.rename(columns=col_map)

    records = []
    for _, row in df.iterrows():
        rec = {}
        for col in df.columns:
            val = row[col]
            rec[col] = "" if pd.isna(val) else str(val).strip()
        records.append(rec)

    cat_cols = [
        "型式", "目的",
        "古期_区分コード", "古期_年代名", "古期_岩石種", "古期_強度", "古期_リスク",
        "新期_区分コード", "新期_年代名", "新期_岩石種", "新期_強度", "新期_リスク",
        "信頼度",
    ]
    uniques = {}
    for col in cat_cols:
        if col in df.columns:
            vals = sorted(df[col].dropna().astype(str).str.strip().unique().tolist())
            uniques[col] = vals

    data = {
        "meta": {
            "source": str(Path(input_path).name),
            "sheet": sheet_name,
            "rows": len(records),
            "columns": list(df.columns),
        },
        "records": records,
        "uniques": uniques,
    }

    json_data = json.dumps(data, ensure_ascii=False)
    html = build_html(json_data, len(records), str(Path(input_path).name))

    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = out_path.stat().st_size // 1024
    print(f"出力完了: {output_path}  ({size_kb} KB)")
    print(f"  → HTMLファイル1つでそのまま使えます（サーバー不要）")


def main():
    parser = argparse.ArgumentParser(description="Excel → スタンドアロン HTML 生成スクリプト")
    parser.add_argument("--input", "-i", default="data/北海道ダム地質分類DB.xlsx")
    parser.add_argument("--sheet", "-s", default="ダム地質区分DB")
    parser.add_argument("--output", "-o", default="docs/index.html")
    args = parser.parse_args()
    convert(args.input, args.sheet, args.output)


if __name__ == "__main__":
    main()
