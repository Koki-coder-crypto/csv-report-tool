#!/usr/bin/env python3
"""売上明細CSVの自動集計・レポート生成ツール（標準ライブラリのみ）。

使い方:
    python3 csv_report.py [入力CSV] [-o 出力フォルダ]
"""
import argparse
import csv
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

REQUIRED_COLUMNS = ["日付", "カテゴリ", "商品名", "数量", "単価"]
DEFAULT_INPUT = "sample_sales.csv"
DEFAULT_OUTPUT = "output"


def yen(n):
    """3桁区切りの円表記。"""
    return f"{n:,.0f}円"


def parse_row(row):
    """1行を検証して辞書に変換。不正なら ValueError。"""
    raw_date = (row["日付"] or "").strip()
    try:
        date = datetime.strptime(raw_date.replace("/", "-"), "%Y-%m-%d")
    except ValueError:
        raise ValueError(f"日付の形式が不正です: '{raw_date}'")

    category = (row["カテゴリ"] or "").strip()
    name = (row["商品名"] or "").strip()
    if not category:
        raise ValueError("カテゴリが空です")
    if not name:
        raise ValueError("商品名が空です")

    try:
        qty = int((row["数量"] or "").strip().replace(",", ""))
        price = float((row["単価"] or "").strip().replace(",", ""))
    except ValueError:
        raise ValueError(f"数量または単価が数値ではありません: '{row['数量']}', '{row['単価']}'")
    if qty <= 0:
        raise ValueError(f"数量は1以上が必要です: {qty}")
    if price < 0:
        raise ValueError(f"単価が負の値です: {price}")

    return {
        "month": date.strftime("%Y-%m"),
        "category": category,
        "name": name,
        "qty": qty,
        "price": price,
        "amount": qty * price,
    }


def load(path):
    """CSVを読み込み (有効行リスト, スキップ数) を返す。致命的エラーは SystemExit。"""
    try:
        f = open(path, newline="", encoding="utf-8-sig")
    except OSError as e:
        sys.exit(f"エラー: 入力ファイルを開けません: {e}")
    with f:
        reader = csv.DictReader(f)
        missing = [c for c in REQUIRED_COLUMNS if c not in (reader.fieldnames or [])]
        if missing:
            sys.exit(f"エラー: 必須列がありません: {', '.join(missing)}")
        rows, skipped = [], 0
        for lineno, row in enumerate(reader, start=2):  # 1行目はヘッダー
            if None in row or any(row[c] is None for c in REQUIRED_COLUMNS):
                print(f"警告: {lineno}行目をスキップ（列数が不正）", file=sys.stderr)
                skipped += 1
                continue
            try:
                rows.append(parse_row(row))
            except ValueError as e:
                print(f"警告: {lineno}行目をスキップ（{e}）", file=sys.stderr)
                skipped += 1
    return rows, skipped


def aggregate(rows, key):
    agg = defaultdict(lambda: {"count": 0, "qty": 0, "amount": 0.0})
    for r in rows:
        a = agg[r[key]]
        a["count"] += 1
        a["qty"] += r["qty"]
        a["amount"] += r["amount"]
    return dict(sorted(agg.items()))


def write_csv(path, header, data):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:  # Excelで文字化けしない
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(data)


def md_table(header, rows, align=None):
    align = align or ["---"] + ["---:"] * (len(header) - 1)
    lines = ["| " + " | ".join(header) + " |", "| " + " | ".join(align) + " |"]
    lines += ["| " + " | ".join(map(str, r)) + " |" for r in rows]
    return "\n".join(lines)


def build_report(src, rows, skipped, by_month, by_cat):
    total = sum(r["amount"] for r in rows)
    count = len(rows)
    qty_total = sum(r["qty"] for r in rows)
    months = list(by_month)
    cats = sorted(by_cat, key=lambda c: -by_cat[c]["amount"])
    best_m = max(by_month, key=lambda m: by_month[m]["amount"])
    best_c = cats[0]
    top_item = max(rows, key=lambda r: r["amount"])

    out = [
        "# 売上集計レポート", "",
        f"- 入力ファイル: `{src}`",
        f"- 作成日時: {datetime.now():%Y-%m-%d %H:%M}",
        f"- 対象期間: {months[0]} 〜 {months[-1]}", "",
        "## 1. 概要サマリー", "",
        md_table(["指標", "値"], [
            ["総売上", yen(total)],
            ["総件数（明細行数）", f"{count:,}件"],
            ["数量合計", f"{qty_total:,}個"],
            ["平均単価（売上÷数量）", yen(total / qty_total)],
            ["1件あたり平均売上", yen(total / count)],
            ["月平均売上", yen(total / len(months))],
            ["スキップした不正行", f"{skipped:,}行"],
        ], ["---", "---:"]), "",
        f"- 最も売上が大きい月は **{best_m}**（{yen(by_month[best_m]['amount'])}）です。",
        f"- 最も売上が大きいカテゴリは **{best_c}**（構成比 {by_cat[best_c]['amount'] / total * 100:.1f}%）です。",
        f"- 1件あたり最大の売上は「{top_item['name']}」"
        f"（{top_item['month']}、{yen(top_item['amount'])}）です。", "",
        "## 2. 月別集計", "",
        md_table(["月", "件数", "数量合計", "売上合計"],
                 [[m, f"{a['count']:,}", f"{a['qty']:,}", yen(a["amount"])] for m, a in by_month.items()]
                 + [["**合計**", f"**{count:,}**", f"**{qty_total:,}**", f"**{yen(total)}**"]]), "",
        "## 3. カテゴリ別集計", "",
        md_table(["カテゴリ", "件数", "数量合計", "売上合計", "売上比率"],
                 [[c, f"{by_cat[c]['count']:,}", f"{by_cat[c]['qty']:,}", yen(by_cat[c]["amount"]),
                   f"{by_cat[c]['amount'] / total * 100:.1f}%"] for c in cats]
                 + [["**合計**", f"**{count:,}**", f"**{qty_total:,}**", f"**{yen(total)}**", "**100.0%**"]]), "",
        "## 4. 月別 × カテゴリ クロス集計（売上）", "",
    ]
    cross = defaultdict(float)
    for r in rows:
        cross[(r["month"], r["category"])] += r["amount"]
    cross_rows = []
    for m in months:
        cross_rows.append([m] + [yen(cross[(m, c)]) for c in cats] + [yen(by_month[m]["amount"])])
    cross_rows.append(["**合計**"] + [f"**{yen(by_cat[c]['amount'])}**" for c in cats] + [f"**{yen(total)}**"])
    out += [md_table(["月"] + cats + ["合計"], cross_rows), ""]
    return "\n".join(out)


def main():
    p = argparse.ArgumentParser(description="売上明細CSVを月別・カテゴリ別に集計しレポートを生成します")
    p.add_argument("input", nargs="?", default=DEFAULT_INPUT, help=f"入力CSV（省略時: {DEFAULT_INPUT}）")
    p.add_argument("-o", "--output", default=DEFAULT_OUTPUT, help=f"出力フォルダ（省略時: {DEFAULT_OUTPUT}）")
    args = p.parse_args()

    rows, skipped = load(args.input)
    if not rows:
        sys.exit("エラー: 有効なデータ行がありません")

    by_month = aggregate(rows, "month")
    by_cat = aggregate(rows, "category")
    total = sum(r["amount"] for r in rows)

    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    write_csv(out_dir / "summary_by_month.csv", ["月", "件数", "数量合計", "売上合計"],
              [[m, a["count"], a["qty"], round(a["amount"])] for m, a in by_month.items()])
    write_csv(out_dir / "summary_by_category.csv", ["カテゴリ", "件数", "数量合計", "売上合計", "売上比率%"],
              [[c, a["count"], a["qty"], round(a["amount"]), f"{a['amount'] / total * 100:.1f}"]
               for c, a in sorted(by_cat.items(), key=lambda x: -x[1]["amount"])])
    (out_dir / "report.md").write_text(
        build_report(args.input, rows, skipped, by_month, by_cat), encoding="utf-8")

    print(f"完了: {len(rows)}件を集計（スキップ {skipped}件）。出力先: {out_dir}/")
    for name in ("report.md", "summary_by_month.csv", "summary_by_category.csv"):
        print(f"  - {out_dir / name}")


if __name__ == "__main__":
    main()
