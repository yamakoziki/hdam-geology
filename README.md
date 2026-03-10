# 北海道ダム地質区分DB 検索システム

Hokkaido Dam Geology Classification Database — Search Application

🌐 **公開URL**: https://yamakoziki.github.io/hdam-geology/

---

## 📁 プロジェクト構成

```
hdam-geology/
├── data/
│   └── 北海道ダム地質分類DB.xlsx   # 元データ（Excelファイル）← ここを更新
├── scripts/
│   └── convert.py                  # Excel → HTML 生成スクリプト
├── docs/
│   └── index.html                  # 検索アプリ（自動生成・GitHub Pages公開）
└── README.md
```

---

## 🔄 データ更新手順

### 1. 必要ライブラリのインストール（初回のみ）

```bash
python3 -m pip install pandas openpyxl
```

### 2. Excelファイルを差し替える

`data/北海道ダム地質分類DB.xlsx` を最新のファイルで上書きする。

### 3. HTMLを再生成する

```bash
cd hdam-geology
python3 scripts/convert.py
```

`docs/index.html` が新しいデータで上書きされます。

### 4. GitHubにプッシュする

```bash
git add .
git commit -m "データ更新"
git push
```

数分後、公開URLに反映されます。

---

## ⚙️ オプション指定

```bash
# ファイルパス・シート名・出力先を指定
python3 scripts/convert.py \
  --input  data/北海道ダム地質分類DB.xlsx \
  --sheet  ダム地質区分DB \
  --output docs/index.html
```

---

## 🔍 検索機能

| 機能 | 説明 |
|------|------|
| テキスト検索 | ダム名・水系名・河川名・所在地の部分一致 |
| ワイルドカード `*` | 分類記号などで使用可（例：`Ⅱ-b*`、`*S3*`、`*R4*`） |
| チップ選択 | 型式・目的・強度・リスク・信頼度など複数選択可 |
| 数値範囲 | 堤高(m)・完成年度の範囲絞り込み |
| AND / OR 切替 | 複数条件の論理演算 |
| 表示列選択 | 22項目から表示列を自由に選択 |
