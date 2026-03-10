# 北海道ダム地質区分DB 検索システム

Hokkaido Dam Geology Classification Database — Search Application

## 📁 プロジェクト構成

```
dam-geology-app/
├── data/
│   └── 北海道ダム地質分類DB.xlsx   # 元データ（Excelファイル）
├── scripts/
│   └── convert.py                  # Excel → JSON 変換スクリプト
├── src/
│   ├── index.html                  # 検索アプリ本体
│   └── dam_data.json               # 変換済みデータ（自動生成）
└── README.md
```

## 🚀 使い方

### 1. 初回セットアップ

Python と必要なライブラリをインストールします。

```bash
pip install pandas openpyxl
```

### 2. Excelデータを変換する

```bash
# デフォルト設定で変換（data/ → src/dam_data.json）
python scripts/convert.py

# ファイルパスを指定して変換
python scripts/convert.py --input data/北海道ダム地質分類DB.xlsx

# オプションを詳しく指定
python scripts/convert.py \
  --input  data/北海道ダム地質分類DB.xlsx \
  --sheet  ダム地質区分DB \
  --output src/dam_data.json
```

### 3. アプリを起動する

`src/index.html` をブラウザで開くだけで動作します。

ただし、`fetch()` を使ってJSONを読み込むため、**ローカルファイルサーバー経由** での起動を推奨します。

```bash
# Python 簡易サーバー（src/ ディレクトリで実行）
cd src
python3 -m http.server 8000
# ブラウザで http://localhost:8000 を開く
```

または VS Code の **Live Server** 拡張機能を使って `src/index.html` を開いてください。

---

## 🔄 データ更新手順

Excelファイルを更新した場合は、変換スクリプトを再実行するだけです。

```bash
python scripts/convert.py
```

`src/dam_data.json` が更新され、アプリをリロードすれば最新データが反映されます。

---

## 🔍 検索機能

| 機能 | 説明 |
|------|------|
| テキスト検索 | ダム名・水系名・河川名・所在地の部分一致検索 |
| ワイルドカード `*` | 分類記号などで使用可（例：`Ⅱ-b*`、`*S3*`、`*R4*`） |
| チップ選択 | 型式・目的・強度・リスク・信頼度など複数選択可 |
| 数値範囲 | 堤高(m)・完成年度の範囲絞り込み |
| AND / OR 切替 | 複数条件の論理演算を選択 |
| 表示列選択 | 22項目から表示する列を自由に選択 |

---

## 📋 データ仕様

- **対象シート**：`ダム地質区分DB`（2枚目）
- **レコード数**：188件
- **主要カラム**：仮No、ダム名、水系名、河川名、所在地、型式、目的、堤高(m)、完成年度、古期/新期 区分コード・年代名・岩石種・強度・リスク、分類記号（完全）、信頼度、判定根拠・参照文献
