#!/usr/bin/env python3
"""沖縄県「市町村の町字別住民基本台帳人口及び世帯数」を一括取得して正規化する。

出典ページ（各年1月1日現在／平成23年〜令和7年）:
  https://www.pref.okinawa.jp/kensei/shinko/1016703/1016705/1016773/1022611/1016806.html

このスクリプトは外部ネットワークにアクセスするので、ローカルまたは
GitHub Actions 上で実行してください。

    python scripts/fetch_aza_population.py            # 全年ダウンロード
    python scripts/fetch_aza_population.py --years 2024 2025
    python scripts/fetch_aza_population.py --normalize-only

出力:
  data/raw/aza/<year>.xlsx            原本
  data/generated/population_aza.json  正規化結果（字ごとの年次系列）
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
import urllib.request
from pathlib import Path

BASE = "https://www.pref.okinawa.jp/_res/projects/default_project/_page_/001/016/806/"

# 出典ページ記載のファイル名。年は「〇年1月1日現在」の西暦。
# 県がページを更新したらここに追記する（R8=2026 が出たら 2026 の行を足す）。
FILES = {
    2011: "tyouazabetsuh23_2.xls",
    2012: "tyouazabetsuh24_1.xls",
    2013: "h25tyouazabetu.xls",
    2014: "h26tyouazabetuzinkou.xlsx",
    2015: "h27tyouazabetujinkou.xlsx",
    2016: "h28tyouazabetujinkou.xlsx",
    2017: "h29chouazabetujinkou.xls",
    2018: "shicyousonnazabetejinnkou.xls",
    2019: "h31azajinkou.xls",
    2020: "r02azajinkou.xlsx",
    2021: "r03azajinkou.xlsx",
    2022: "shuuseigor04azajinkou.xlsx",
    2023: "shuuseigor05azajinkou.xlsx",
    2024: "r06azajinkou.xlsx",
    2025: "r07azajinkou.xlsx",
}

RAW = Path("data/raw/aza")
OUT = Path("data/generated/population_aza.json")

# 平成23〜25年は3月31日現在、平成26年以降は1月1日現在。基準日が違う点に注意。
REFERENCE_DATE = {2011: "03-31", 2012: "03-31", 2013: "03-31"}


def download(years):
    RAW.mkdir(parents=True, exist_ok=True)
    for y in years:
        if y not in FILES:
            print(f"  ! {y}年はFILESに未登録。出典ページを確認して追記してください。", file=sys.stderr)
            continue
        fn = FILES[y]
        dest = RAW / f"{y}{Path(fn).suffix}"
        if dest.exists():
            print(f"  = {dest} (既存)")
            continue
        url = BASE + fn
        req = urllib.request.Request(url, headers={"User-Agent": "okinawa-atlas/1.0"})
        with urllib.request.urlopen(req, timeout=60) as r, dest.open("wb") as f:
            f.write(r.read())
        print(f"  + {dest} <- {url}")


def norm_name(s: str) -> str:
    """字名の表記ゆれを畳む。全角/半角、空白、括弧注記を除去。"""
    s = unicodedata.normalize("NFKC", str(s))
    s = re.sub(r"[（(].*?[)）]", "", s)
    return re.sub(r"\s+", "", s).strip()


def read_year(path: Path, year: int):
    """1ファイルから (市町村名, 字名, 人口, 世帯数) を抽出する。

    県のExcelは年によってヘッダ位置・列構成が動くため、見出し行を
    キーワードで探し当ててから読む。想定外のレイアウトは例外にして
    黙って誤ったデータを作らないようにする。
    """
    import pandas as pd

    engine = "xlrd" if path.suffix == ".xls" else "openpyxl"
    raw = pd.read_excel(path, sheet_name=0, header=None, engine=engine)

    header_row = None
    for i in range(min(15, len(raw))):
        row = " ".join(str(v) for v in raw.iloc[i].tolist())
        if re.search(r"字|町字", row) and re.search(r"人口|総数", row):
            header_row = i
            break
    if header_row is None:
        raise ValueError(f"{path}: 見出し行を特定できませんでした。手動で確認してください。")

    df = pd.read_excel(path, sheet_name=0, header=header_row, engine=engine)
    df.columns = [norm_name(c) for c in df.columns]

    def pick(*keys):
        for c in df.columns:
            if any(k in c for k in keys):
                return c
        return None

    c_muni, c_aza = pick("市町村"), pick("字", "町字")
    c_pop, c_hh = pick("人口", "総数"), pick("世帯")
    if not (c_aza and c_pop):
        raise ValueError(f"{path}: 字/人口の列が見つかりません。列={list(df.columns)}")

    recs, cur_muni = [], None
    for _, r in df.iterrows():
        muni = norm_name(r[c_muni]) if c_muni and pd.notna(r.get(c_muni)) else None
        if muni:
            cur_muni = muni
        aza = norm_name(r[c_aza]) if pd.notna(r.get(c_aza)) else None
        if not aza or not cur_muni or aza in ("計", "合計", "総計"):
            continue
        pop = pd.to_numeric(r.get(c_pop), errors="coerce")
        if pd.isna(pop):
            continue
        recs.append({
            "municipality": cur_muni, "aza": aza, "year": year,
            "date": f"{year}-{REFERENCE_DATE.get(year, '01-01')}",
            "population": int(pop),
            "households": int(pd.to_numeric(r.get(c_hh), errors="coerce"))
                          if c_hh and pd.notna(pd.to_numeric(r.get(c_hh), errors="coerce")) else None,
        })
    return recs


def normalize(years):
    series, problems = {}, []
    for y in years:
        hit = [p for p in RAW.glob(f"{y}.*")]
        if not hit:
            problems.append(f"{y}: 原本なし（先にダウンロードしてください）")
            continue
        try:
            for rec in read_year(hit[0], y):
                key = f'{rec["municipality"]}／{rec["aza"]}'
                e = series.setdefault(key, {"municipality": rec["municipality"],
                                            "aza": rec["aza"], "trend": []})
                e["trend"].append({"year": rec["year"], "date": rec["date"],
                                   "population": rec["population"],
                                   "households": rec["households"]})
        except Exception as ex:                       # noqa: BLE001
            problems.append(f"{y}: {ex}")

    out = sorted(series.values(), key=lambda e: (e["municipality"], e["aza"]))
    for e in out:
        e["trend"].sort(key=lambda t: t["year"])

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\nwrote {OUT}  字数={len(out)}")
    if problems:
        print("\n要対応:")
        for p in problems:
            print("  -", p)
    thin = [e for e in out if len(e["trend"]) < len(years) * 0.6]
    if thin:
        print(f"\n年次が欠けている字 {len(thin)}件（名称変更・区画再編の可能性。"
              f"data/aza_crosswalk.csv で対応付けを管理してください）")
        for e in thin[:15]:
            print(f"  - {e['municipality']}／{e['aza']}  ({len(e['trend'])}年分)")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--years", nargs="*", type=int, default=sorted(FILES))
    ap.add_argument("--normalize-only", action="store_true")
    a = ap.parse_args()
    if not a.normalize_only:
        print("ダウンロード:")
        download(a.years)
    print("\n正規化:")
    normalize(a.years)
