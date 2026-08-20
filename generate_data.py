# -*- coding: utf-8 -*-
"""
取引先(ニッター)向け棚番検索アプリ用データ生成スクリプト

tanaban_master.xlsx を読み込み、ニッターコードごとに絞り込んだ
JSONファイルを data/ 配下に出力する。
生成した各社のJSONファイルには、自社の棚番データのみが入る
(他社の分は一切含まれない = 会社間の情報漏洩を防ぐ設計)。

実行方法: python generate_data.py
"""
import json
import os
import sys
from collections import defaultdict

import openpyxl

SOURCE_XLSX = r"G:\マイドライブ\AppSheetデータ\tanaban_master.xlsx"
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "data")
INDEX_FILE = os.path.join(OUTPUT_DIR, "_index.json")


def main():
    if not os.path.exists(SOURCE_XLSX):
        print(f"エラー: 元データが見つかりません: {SOURCE_XLSX}", file=sys.stderr)
        sys.exit(1)

    wb = openpyxl.load_workbook(SOURCE_XLSX, read_only=True, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(min_row=1, values_only=True))
    header = rows[0]
    expected = ("棚番", "JAN", "品番", "カラー", "品名", "ニッター", "ニッターコード")
    if header != expected:
        print(f"警告: 列構成が想定と異なります。実際: {header}", file=sys.stderr)

    by_code = defaultdict(list)
    niitaa_names = {}

    for r in rows[1:]:
        if not r or r[6] is None:
            continue
        tanaban, jan, hinban, color, hinmei, niitaa, code = r
        code = str(code).strip()
        if not code:
            continue
        by_code[code].append({
            "tanaban": tanaban,
            "jan": str(jan) if jan is not None else "",
            "hinban": str(hinban) if hinban is not None else "",
            "color": str(color) if color is not None else "",
            "hinmei": hinmei or "",
        })
        niitaa_names[code] = niitaa

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 会社ごとのファイルを書き出し
    for code, items in by_code.items():
        out_path = os.path.join(OUTPUT_DIR, f"{code}.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump({
                "code": code,
                "niitaa": niitaa_names.get(code, ""),
                "count": len(items),
                "items": items,
            }, f, ensure_ascii=False, separators=(",", ":"))

    # 索引ファイル(社数・件数の一覧。QR配布時の確認用。中身の棚番データは含まない)
    index = [
        {"code": code, "niitaa": niitaa_names.get(code, ""), "count": len(items)}
        for code, items in sorted(by_code.items())
    ]
    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)

    print(f"完了: {len(by_code)}社分のファイルを {OUTPUT_DIR} に生成しました。")


if __name__ == "__main__":
    main()
