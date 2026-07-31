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
  data/raw/aza/<year>.xlsx                 原本
  data/generated/population_aza.json       正規化結果（字ごとの年次系列）
  data/generated/aza_quality_report.json   市町村名の補正・欠けている字・失敗した年
"""
from __future__ import annotations

import argparse
import csv
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
REPORT = Path("data/generated/aza_quality_report.json")
CROSSWALK = Path("data/aza_crosswalk.csv")

# 年次間で合算してよい relation。それ以外（succeeded_by / split_into /
# merged_from / separate / 空欄）は関係の記録のみで、系列は合算しない。
MERGE_RELATIONS = {"same", "typo"}

# 丁目の漢数字 -> 算用数字。分類Bの正規化に使う。
CHOME_DIGITS = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5",
                "六": "6", "七": "7", "八": "8", "九": "9", "十": "10"}

# 平成23〜25年は3月31日現在、平成26年以降は1月1日現在。基準日が違う点に注意。
REFERENCE_DATE = {2011: "03-31", 2012: "03-31", 2013: "03-31"}

# 沖縄県の41市町村の正式名。ここに無い市町村名が出てきたら異常として落とす。
MUNICIPALITIES = frozenset({
    "那覇市", "宜野湾市", "石垣市", "浦添市", "名護市", "糸満市", "沖縄市",
    "豊見城市", "うるま市", "宮古島市", "南城市",
    "国頭村", "大宜味村", "東村", "今帰仁村", "本部町", "恩納村", "宜野座村",
    "金武町", "伊江村", "読谷村", "嘉手納町", "北谷町", "北中城村", "中城村",
    "西原町", "与那原町", "南風原町", "渡嘉敷村", "座間味村", "粟国村",
    "渡名喜村", "南大東村", "北大東村", "伊平屋村", "伊是名村", "久米島町",
    "八重瀬町", "多良間村", "竹富町", "与那国町",
})

# 県の原本Excelに実在する明白な誤字。原本は加工しないので、変換時にだけ効かせる。
# 字名の名寄せ（data/aza_crosswalk.csv）とは性質が違う——こちらは判断の余地がない
# 誤字なので別管理にする。県のExcelは印刷ページごとに市町村名を再掲するため、
# 誤記された見出し以降の字がまるごと誤った市町村に吸われる。
#   宜野湾志市: 2017年は見出し2つとも誤記、2018〜2022年は2ページ目だけ誤記
#   南大東:     2024年の見出し（「村」が欠落）
MUNI_TYPOS = {
    "宜野湾志市": "宜野湾市",
    "南大東": "南大東村",
}


class UnknownMunicipality(ValueError):
    """41市町村の正式名にも MUNI_TYPOS にも一致しない市町村名を見つけた。

    推測で寄せると全指標が静かにずれるので、その年の変換ごと失敗させる。
    """

    def __init__(self, year, path, sheet, row, value, aza):
        self.year, self.row, self.value, self.aza = year, row, value, aza
        super().__init__(
            f"{path.name} [{sheet}] 行{row}（Excel {row + 1}行目・町字名={aza}）: "
            f"市町村名 {value!r} が41市町村の正式名に一致しません。"
            f"県の原本の誤記であれば MUNI_TYPOS に追加してください。"
        )


class AzaNameCollision(ValueError):
    """年次間名寄せの正規化キーが、同一年・同一市町村内で衝突した。

    docs/aza-matching-policy.md 論点1の決定どおり、推測で片方を選ばず
    その年の変換ごと失敗させる。crosswalk（typo/same）を適用したあとの
    表記どうしが、同じ年に同じ正規化キーへ集まった場合に発生する。
    """

    def __init__(self, year, municipality, key, forms):
        self.year, self.municipality, self.key, self.forms = year, municipality, key, forms
        super().__init__(
            f"{year}年 {municipality}: 正規化キー {key!r} に複数の表記が同じ年に存在します "
            f"({', '.join(forms)})。合算すると二重計上になるため、この年の変換を失敗させます。"
            f"別の実体なら data/aza_crosswalk.csv に relation=separate で明示してください。"
        )


def load_crosswalk():
    """data/aza_crosswalk.csv を読む。# で始まる行と空行はコメントとして無視する。"""
    if not CROSSWALK.exists():
        return []
    rows = []
    with CROSSWALK.open(encoding="utf-8", newline="") as f:
        reader = csv.reader(f)
        header = None
        for raw_row in reader:
            if not raw_row or raw_row[0].startswith("#"):
                continue
            if header is None:
                header = raw_row
                continue
            row = dict(zip(header, raw_row))
            row["years"] = {int(y) for y in row.get("years", "").split(",") if y.strip()}
            rows.append(row)
    return rows


def crosswalk_rename_map(crosswalk):
    """(municipality, name_from, year) -> name_to。relation が合算対象のものだけ。"""
    m = {}
    for row in crosswalk:
        if row.get("relation") not in MERGE_RELATIONS:
            continue
        years = row["years"]
        m[(row["municipality"], row["name_from"])] = (row["name_to"], years)
    return m


def normalize_key(aza: str) -> str:
    """分類A・Bの正規化キー。先頭の「字」を落とし、丁目の漢数字を算用数字にする。"""
    s = aza[1:] if aza.startswith("字") else aza
    m = re.match(r"^(.*?)([一二三四五六七八九十]+)丁目$", s)
    if m:
        digits = CHOME_DIGITS.get(m.group(2))
        if digits:
            s = m.group(1) + digits + "丁目"
    return s


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


def _select_sheet(path: Path, engine: str):
    """町字別の実データが入っているシートを選ぶ。

    2021年以降は表紙・要約シートが先頭に来るため、シート名に「町字別」を
    含むものを優先する。見つからなければ最初のシートにフォールバックする。
    """
    if engine == "openpyxl":
        import openpyxl
        names = openpyxl.load_workbook(path, read_only=True).sheetnames
    else:
        import xlrd
        names = xlrd.open_workbook(path).sheet_names()
    candidates = [n for n in names if "町字別" in n]
    return candidates[0] if candidates else 0


def _cell(row, i):
    v = row.iloc[i] if i < len(row) else None
    import pandas as pd
    return "" if pd.isna(v) else norm_name(v)


def _ffill(vals):
    """空セルを直前の非空値で埋める（結合セルの復元）。"""
    out, last = [], ""
    for v in vals:
        if v:
            last = v
        out.append(last)
    return out


def _ffill_within_group(vals, group):
    """親グループ（group）が変わったら空文字にリセットしてからffillする。"""
    out, last, prev_group = [], "", object()
    for v, g in zip(vals, group):
        if g != prev_group:
            last = ""
        if v:
            last = v
        out.append(last)
        prev_group = g
    return out


def read_year(path: Path, year: int):
    """1ファイルから (市町村名, 字名, 人口, 世帯数) を抽出する。

    県のExcelは年によってシート構成・見出しの階層（男女・国籍別の内訳の
    有無）が異なるため、見出し行の位置と列の親子関係をそのつど検出してから
    読む。想定外のレイアウトは例外にして黙って誤ったデータを作らないように
    する。
    """
    import pandas as pd

    engine = "xlrd" if path.suffix == ".xls" else "openpyxl"
    sheet = _select_sheet(path, engine)
    raw = pd.read_excel(path, sheet_name=sheet, header=None, engine=engine)
    ncols = raw.shape[1]

    header_row = None
    for i in range(min(20, len(raw))):
        cells = [_cell(raw.iloc[i], c) for c in range(ncols)]
        if any(v in ("市町村名", "市町村") for v in cells) and \
           any(v in ("町字名", "字名", "町字") for v in cells):
            header_row = i
            break
    if header_row is None:
        raise ValueError(f"{path} [{sheet}]: 見出し行を特定できませんでした。手動で確認してください。")

    header_cells = [_cell(raw.iloc[header_row], c) for c in range(ncols)]
    muni_col = next((c for c, v in enumerate(header_cells) if v in ("市町村名", "市町村")), None)
    aza_col = next((c for c, v in enumerate(header_cells) if v in ("町字名", "字名", "町字")), None)
    if muni_col is None or aza_col is None:
        raise ValueError(f"{path} [{sheet}]: 市町村名/町字名の列が見つかりません。見出し={header_cells}")

    data_start = None
    for i in range(header_row + 1, min(header_row + 6, len(raw))):
        if _cell(raw.iloc[i], aza_col):
            data_start = i
            break
    if data_start is None:
        raise ValueError(f"{path} [{sheet}]: データ開始行を特定できませんでした。")

    n_sub = data_start - header_row - 1
    if n_sub > 2:
        raise ValueError(f"{path} [{sheet}]: 見出しの階層が想定外です（小見出し行={n_sub}）。手動で確認してください。")

    level0 = _ffill(header_cells)
    level1_raw = [_cell(raw.iloc[header_row + 1], c) for c in range(ncols)] if n_sub >= 1 else [""] * ncols
    level1 = _ffill_within_group(level1_raw, level0)
    level2 = [_cell(raw.iloc[header_row + 2], c) for c in range(ncols)] if n_sub >= 2 else [""] * ncols

    pop_col = hh_col = None
    for c in range(aza_col + 1, ncols):
        own = header_cells[c]
        if n_sub == 2:
            if level0[c] == "人口" and level1[c] == "合計" and level2[c] == "計":
                pop_col = c
            elif level0[c] == "世帯数" and level1[c] == "" and level2[c] == "":
                hh_col = c
        else:
            if level0[c] == "人口" and level1[c] == "計":
                pop_col = c
            elif level0[c] == "世帯数" and level1[c] == "計":
                hh_col = c
            elif own == "計" and level1[c] == "" and pop_col is None:
                pop_col = c
            elif own == "世帯数" and level1[c] == "":
                hh_col = c

    if pop_col is None:
        raise ValueError(f"{path} [{sheet}]: 人口(合計)列を特定できませんでした。見出し={header_cells}")

    recs, corrections, cur_muni = [], [], None
    for i in range(data_start, len(raw)):
        r = raw.iloc[i]
        muni = _cell(r, muni_col)
        if muni:
            fixed = MUNI_TYPOS.get(muni, muni)
            if fixed != muni:
                corrections.append({
                    "year": year, "file": path.name, "sheet": str(sheet),
                    "row_pandas": i, "row_excel": i + 1,
                    "found": muni, "corrected_to": fixed,
                })
            cur_muni = fixed
        aza = _cell(r, aza_col)
        if not aza or not cur_muni or aza in ("計", "合計", "総計"):
            continue
        pop = pd.to_numeric(r.iloc[pop_col], errors="coerce")
        if pd.isna(pop):
            continue
        # 検査するのはレコードを生成する行の市町村名だけ。県計行「沖縄県」や脚注行
        # （※〜）も市町村名列に現れるが、町字名か人口が空でここまで到達しない。
        if cur_muni not in MUNICIPALITIES:
            raise UnknownMunicipality(year, path, sheet, i, cur_muni, aza)
        hh = pd.to_numeric(r.iloc[hh_col], errors="coerce") if hh_col is not None else None
        recs.append({
            "municipality": cur_muni, "aza": aza, "year": year,
            "date": f"{year}-{REFERENCE_DATE.get(year, '01-01')}",
            "population": int(pop),
            "households": int(hh) if hh is not None and pd.notna(hh) else None,
        })
    return recs, corrections


def normalize(years):
    """年次間の名寄せ方針（docs/aza-matching-policy.md）を適用して系列を作る。

    1. crosswalk の relation=same/typo で該当年の表記を寄せ先に置き換える
    2. 先頭の「字」を落とし丁目の漢数字を算用数字にした正規化キーでまとめる
    3. 同一年・同一市町村で正規化キーが衝突したら、その年の変換を失敗させる
       （どちらが正しいか推測しない。合算すると二重計上になるため）
    4. 表示名はグループが最後に現れた年の表記をそのまま使う
    """
    crosswalk = load_crosswalk()
    rename_map = crosswalk_rename_map(crosswalk)

    series, problems = {}, []
    corrections, unknown, per_year = [], [], {}
    for y in years:
        hit = [p for p in RAW.glob(f"{y}.*")]
        if not hit:
            problems.append(f"{y}: 原本なし（先にダウンロードしてください）")
            continue
        try:
            recs, corr = read_year(hit[0], y)
        except UnknownMunicipality as ex:
            problems.append(f"{y}: {ex}")
            unknown.append({"year": y, "row_excel": ex.row + 1,
                            "value": ex.value, "aza": ex.aza})
            continue
        except Exception as ex:                       # noqa: BLE001
            problems.append(f"{y}: {ex}")
            continue

        # crosswalk（same/typo）は「合算するかどうか」の判断材料であって、
        # 表示名の選び方とは別。表示名は常に、その年に県が実際に書いた
        # 表記（rec["aza"]）を使う。合算用のキーだけ、読み替えたうえで
        # 分類A/Bの正規化にかける。
        renamed = []
        for rec in recs:
            name_to, apply_years = rename_map.get(
                (rec["municipality"], rec["aza"]), (None, None))
            grouping_name = rec["aza"]
            if name_to and (not apply_years or rec["year"] in apply_years):
                grouping_name = name_to
            renamed.append({**rec, "key": normalize_key(grouping_name)})

        by_muni_key = {}
        for rec in renamed:
            by_muni_key.setdefault((rec["municipality"], rec["key"]), []).append(rec)
        collisions = [(muni, key, sorted({r["aza"] for r in items}))
                      for (muni, key), items in by_muni_key.items()
                      if len({r["aza"] for r in items}) > 1]
        if collisions:
            for muni, key, forms in collisions:
                ex = AzaNameCollision(y, muni, key, forms)
                problems.append(f"{y}: {ex}")
            continue

        corrections.extend(corr)
        per_year[y] = len(recs)
        for rec in renamed:
            gkey = f'{rec["municipality"]}／{rec["key"]}'
            e = series.setdefault(gkey, {"municipality": rec["municipality"],
                                         "key": rec["key"], "trend": [], "forms": {}})
            e["trend"].append({"year": rec["year"], "date": rec["date"],
                               "population": rec["population"],
                               "households": rec["households"]})
            e["forms"][rec["year"]] = rec["aza"]

    # 1年も読めなかったときに空のJSONで既存の出力を上書きすると、画面上は
    # 「データなし」ではなく静かに全滅する。書かずに失敗として返す。
    if not per_year:
        print("\n全年の変換に失敗しました。既存の出力は上書きしません。", file=sys.stderr)
        for p in problems:
            print("  -", p, file=sys.stderr)
        return False

    out, normalized_groups = [], []
    for e in series.values():
        trend = sorted(e["trend"], key=lambda t: t["year"])
        forms = e["forms"]
        display_name = forms[max(forms)]
        out.append({"municipality": e["municipality"], "aza": display_name, "trend": trend})
        distinct = sorted(set(forms.values()))
        if len(distinct) > 1:
            normalized_groups.append({"municipality": e["municipality"],
                                      "display": display_name, "forms": distinct})
    out.sort(key=lambda e: (e["municipality"], e["aza"]))
    normalized_groups.sort(key=lambda g: (g["municipality"], g["display"]))

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\nwrote {OUT}  字数={len(out)}（名寄せ後。表記ゆれの吸収 {len(normalized_groups)}グループ）")

    thin = [e for e in out if len(e["trend"]) < len(years) * 0.6]
    munis = sorted({e["municipality"] for e in out})

    crosswalk_names = {r["name_from"] for r in crosswalk} | {r["name_to"] for r in crosswalk}
    crosswalk_candidates = [
        f'{e["municipality"]},{e["aza"]},,,,,,{len(e["trend"])}年分（要確認）'
        for e in thin if e["aza"] not in crosswalk_names
    ]

    report = {
        "counts": {
            "aza": len(out),
            "municipalities": len(munis),
            "years_loaded": len(per_year),
            "muni_name_corrections": len(corrections),
            "unknown_municipalities": len(unknown),
            "normalized_groups": len(normalized_groups),
            "thin_series": len(thin),
            "failures": len(problems),
        },
        "records_per_year": per_year,
        "municipalities": munis,
        "muni_name_corrections": corrections,
        "unknown_municipalities": unknown,
        "normalized_groups": normalized_groups,
        "failures": problems,
        "thin_series": [{"municipality": e["municipality"], "aza": e["aza"],
                         "years": [t["year"] for t in e["trend"]]} for e in thin],
        "crosswalk_candidates": crosswalk_candidates,
    }
    REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"wrote {REPORT}")

    if corrections:
        print(f"\n市町村名の誤記を補正 {len(corrections)}件（原本は変更していません）")
        for c in corrections:
            print(f"  - {c['year']}年 Excel {c['row_excel']}行目: "
                  f"{c['found']} -> {c['corrected_to']}")
    if problems:
        print("\n要対応:")
        for p in problems:
            print("  -", p)
    if thin:
        print(f"\n年次が欠けている字 {len(thin)}件（名称変更・区画再編の可能性。"
              f"data/aza_crosswalk.csv で対応付けを管理してください）")
        for e in thin[:15]:
            print(f"  - {e['municipality']}／{e['aza']}  ({len(e['trend'])}年分)")
        if crosswalk_candidates:
            print(f"\ncrosswalk 未登録の候補 {len(crosswalk_candidates)}件"
                  "（data/aza_crosswalk.csv にそのまま貼れる形）:")
            for c in crosswalk_candidates[:10]:
                print(f"  {c}")
    return True


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--years", nargs="*", type=int, default=sorted(FILES))
    ap.add_argument("--normalize-only", action="store_true")
    a = ap.parse_args()
    if not a.normalize_only:
        print("ダウンロード:")
        download(a.years)
    print("\n正規化:")
    # 失敗を終了コードで返す。CI側の「原本が無ければスキップ」がこれで効く。
    sys.exit(0 if normalize(a.years) else 1)
