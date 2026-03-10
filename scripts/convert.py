#!/usr/bin/env python3
"""
北海道ダム地質区分DB - Excel → JSON 変換スクリプト
使用方法:
  python scripts/convert.py
  python scripts/convert.py --input data/北海道ダム地質分類DB.xlsx
  python scripts/convert.py --input data/北海道ダム地質分類DB.xlsx --sheet ダム地質区分DB
  python scripts/convert.py --input data/北海道ダム地質分類DB.xlsx --output src/dam_data.json
"""

import argparse
import json
import sys
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas が必要です。 pip install pandas openpyxl を実行してください。")
    sys.exit(1)


def convert(input_path: str, sheet_name: str, output_path: str) -> None:
    print(f"読み込み中: {input_path}  シート: {sheet_name}")

    xl = pd.ExcelFile(input_path)
    if sheet_name not in xl.sheet_names:
        print(f"ERROR: シート '{sheet_name}' が見つかりません。")
        print(f"利用可能なシート: {xl.sheet_names}")
        sys.exit(1)

    df = pd.read_excel(input_path, sheet_name=sheet_name, header=0)
    print(f"  → {len(df)} 行, {len(df.columns)} 列 を読み込みました")

    # キー名の改行を _ に変換（JS での扱いを容易にするため）
    col_map = {col: col.replace("\n", "_") for col in df.columns}
    df = df.rename(columns=col_map)

    # レコード変換
    records = []
    for _, row in df.iterrows():
        rec = {}
        for col in df.columns:
            val = row[col]
            rec[col] = "" if pd.isna(val) else str(val).strip()
        records.append(rec)

    # カテゴリカル列のユニーク値
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

    output = {
        "meta": {
            "source": str(Path(input_path).name),
            "sheet": sheet_name,
            "rows": len(records),
            "columns": list(df.columns),
        },
        "records": records,
        "uniques": uniques,
    }

    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    size_kb = out_path.stat().st_size // 1024
    print(f"出力完了: {output_path}  ({size_kb} KB, {len(records)} 件)")


def main():
    parser = argparse.ArgumentParser(description="Excel → JSON 変換スクリプト")
    parser.add_argument(
        "--input", "-i",
        default="data/北海道ダム地質分類DB.xlsx",
        help="入力 Excel ファイルパス (デフォルト: data/北海道ダム地質分類DB.xlsx)",
    )
    parser.add_argument(
        "--sheet", "-s",
        default="ダム地質区分DB",
        help="読み込むシート名 (デフォルト: ダム地質区分DB)",
    )
    parser.add_argument(
        "--output", "-o",
        default="src/dam_data.json",
        help="出力 JSON ファイルパス (デフォルト: src/dam_data.json)",
    )
    args = parser.parse_args()
    convert(args.input, args.sheet, args.output)


if __name__ == "__main__":
    main()
