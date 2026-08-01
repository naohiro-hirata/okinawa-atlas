#!/usr/bin/env python3
"""e-Stat API から市町村別の人口推移と年齢ピラミッドを取得して正規化する。

出典（すべて総務省統計局 e-Stat 経由）:
  市町村別人口推移: 国勢調査 1980〜2020年（各回・統計表IDが年ごとに異なる）
                     ＋ 令和7年国勢調査 速報集計（男女別人口。2026-05-29公表）
  市町村別ピラミッド: 令和2年国勢調査 人口等基本集計（男女，年齢5歳階級）

appId は環境変数 ESTAT_APP_ID から読む。コードにもコミットにも値を残さないこと。

    python scripts/fetch_census.py            # API取得＋正規化
    python scripts/fetch_census.py --normalize-only   # data/raw/census/ のキャッシュのみで再構築

出力:
  data/raw/census/*.json                      APIレスポンスのキャッシュ（gitignore対象）
  data/generated/population_muni.json         市町村ごとの人口推移＋ピラミッド
  data/generated/census_quality_report.json   取得できなかった市町村・年の一覧
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.parse
import urllib.request
from pathlib import Path

API_BASE = "https://api.e-stat.go.jp/rest/3.0/app/json/getStatsData"
RAW = Path("data/raw/census")
OUT = Path("data/generated/population_muni.json")
REPORT = Path("data/generated/census_quality_report.json")

# 沖縄県41市町村の現在の市町村コード（data/generated/municipalities.json と同じ体系）。
CURRENT_AREA = {
    "那覇市": "47201", "宜野湾市": "47205", "石垣市": "47207", "浦添市": "47208",
    "名護市": "47209", "糸満市": "47210", "沖縄市": "47211", "豊見城市": "47212",
    "うるま市": "47213", "宮古島市": "47214", "南城市": "47215",
    "国頭村": "47301", "大宜味村": "47302", "東村": "47303", "今帰仁村": "47306",
    "本部町": "47308", "恩納村": "47311", "宜野座村": "47313", "金武町": "47314",
    "伊江村": "47315", "読谷村": "47324", "嘉手納町": "47325", "北谷町": "47326",
    "北中城村": "47327", "中城村": "47328", "西原町": "47329", "与那原町": "47348",
    "南風原町": "47350", "渡嘉敷村": "47353", "座間味村": "47354", "粟国村": "47355",
    "渡名喜村": "47356", "南大東村": "47357", "北大東村": "47358", "伊平屋村": "47359",
    "伊是名村": "47360", "久米島町": "47361", "八重瀬町": "47362", "多良間村": "47375",
    "竹富町": "47381", "与那国町": "47382",
}
assert len(CURRENT_AREA) == 41

# 合併・境界再編で、旧市町村コードの単純合算に置き換える必要がある年。
# e-Statの各年表の沖縄県エリア一覧を実測して確認済み（1980〜2005年は旧市町村単位）。
#   宮古島市 2005-10-01合併（国勢調査基準日と同日のため2005年表は既に統合済み）
#   うるま市 2005-04-01合併（2005年表は既に統合済み）
#   久米島町 2002-04-01合併（2005年表は既に統合済み）
#   南城市   2006-01-01合併（2005年表はまだ旧町村のまま）
#   八重瀬町 2006-01-01合併（2005年表はまだ旧町村のまま）
# 豊見城市は2002-04-01に村→市（境界変更なし、単なる名称・コード変更）なので合算しない。
OLD_AREA = {
    "宮古島市": ["47206", "47371", "47372", "47373", "47374"],   # 平良市+城辺町+下地町+上野村+伊良部町
    "うるま市": ["47202", "47203", "47322", "47323"],             # 石川市+具志川市+与那城村+勝連町
    "久米島町": ["47351", "47352"],                                 # 仲里村+具志川村
    "南城市": ["47345", "47346", "47347", "47349"],               # 玉城村+知念村+佐敷町+大里村
    "八重瀬町": ["47343", "47344"],                                 # 東風平町+具志頭村
}
MERGE_YEARS = {
    "宮古島市": {1980, 1985, 1990, 1995, 2000},
    "うるま市": {1980, 1985, 1990, 1995, 2000},
    "久米島町": {1980, 1985, 1990, 1995, 2000},
    "南城市": {1980, 1985, 1990, 1995, 2000, 2005},
    "八重瀬町": {1980, 1985, 1990, 1995, 2000, 2005},
}
# 豊見城市: 境界は変わらないが、コードが村(47341)→市(47212)に変わる年。合算せず1:1で読み替える。
RENAME_AREA = {
    "豊見城市": {"old_code": "47341", "years": {1980, 1985, 1990, 1995, 2000}},
}

# 統計表ID。総人口（男女計・全域）を取り出すための固定パラメータ付き。
# e-Stat の getStatsList で「国勢調査」「市区町村」を条件に実測して特定した（2026-08-01）。
CENSUS_TABLES = {
    1980: {"id": "0000030127", "fixed": {"cdCat01": "00700", "cdCat02": "000", "cdCat03": "000"}},
    1985: {"id": "0000030447", "fixed": {"cdCat01": "000"}},
    1990: {"id": "0000031399", "fixed": {"cdCat01": "000"}},
    1995: {"id": "0000032217", "fixed": {"cdCat01": "000"}},
    2000: {"id": "0000032963", "fixed": {"cdCat01": "00700", "cdCat02": "000"}},
    2005: {"id": "0000033784", "fixed": {"cdCat01": "00700", "cdCat02": "000"}},
    2010: {"id": "0003038586", "fixed": {"cdTab": "020", "cdCat01": "00710"}},
    2015: {"id": "0003148500", "fixed": {"cdTab": "020", "cdCat01": "00710"}},
    2020: {"id": "0003445162", "fixed": {"cdCat01": "0", "cdCat02": "0", "cdCat03": "00"}},
}
R7_POP_TABLE = {"id": "0004050397", "fixed": {"cdTab": "2025_01", "cdCat01": "0"}}
R7_HH_TABLE = {"id": "0004050417", "fixed": {"cdTab": "2025_13"}}
PYRAMID_TABLE = {"id": "0003445162", "fixed": {"cdCat01": "0"}}  # cat02/cat03は開けたまま全件取る

# 令和2年国勢調査の年齢5歳階級コード。85歳以上はAGESの最終区分「85+」に合算する。
AGE_CODES = [f"{i:02d}" for i in range(1, 22)]  # 01=0-4歳 ... 21=100歳以上
AGES = ["0-4", "5-9", "10-14", "15-19", "20-24", "25-29", "30-34", "35-39", "40-44",
        "45-49", "50-54", "55-59", "60-64", "65-69", "70-74", "75-79", "80-84", "85+"]


class EstatError(RuntimeError):
    pass


def app_id():
    v = os.environ.get("ESTAT_APP_ID")
    if not v:
        raise EstatError("環境変数 ESTAT_APP_ID が設定されていません。e-Statのappidを設定してください。")
    return v


def cache_path(key):
    return RAW / f"{key}.json"


def estat_get(key, table_id, fixed, extra, use_cache_only=False):
    """1回のAPI呼び出し結果を data/raw/census/<key>.json にキャッシュしつつ返す。"""
    path = cache_path(key)
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    if use_cache_only:
        raise EstatError(f"{key}: キャッシュがありません（--normalize-only では取得できません）")

    params = {"appId": app_id(), "statsDataId": table_id, **fixed, **extra}
    url = API_BASE + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": "okinawa-atlas/1.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        body = json.loads(r.read().decode("utf-8"))

    result = body.get("GET_STATS_DATA", {}).get("RESULT", {})
    if result.get("STATUS") not in (0, 1):  # 1=正常終了だが該当データなし。0=正常
        raise EstatError(f"{key}: e-Stat APIエラー status={result.get('STATUS')} {result.get('ERROR_MSG')}")

    RAW.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(body, ensure_ascii=False), encoding="utf-8")
    return body


# e-Statの表章上の記号。DATA_INF.NOTE で「当該数値がないもの」＝実際の値が0、と
# 定義されている（総務省統計局の表記規則）。「取得できなかった」ではなく「実測値の0」
# なので、CLAUDE.mdの原則3どおり0として扱う。空文字/Noneのみ「値なし」として除外する。
ZERO_MARKERS = {"-", "***"}


def _parse_value(v):
    s = v.get("$")
    if s in (None, ""):
        return None
    if s in ZERO_MARKERS:
        return 0
    return int(s)


def values_by_area(body):
    """{area_code: population(int)} を返す。該当データなしなら空dict。"""
    data = body.get("GET_STATS_DATA", {}).get("STATISTICAL_DATA", {})
    info = data.get("DATA_INF")
    if not info:
        return {}
    vals = info["VALUE"]
    if isinstance(vals, dict):
        vals = [vals]
    out = {}
    for v in vals:
        n = _parse_value(v)
        if n is not None:
            out[v["@area"]] = n
    return out


def fetch_year_totals(year, use_cache_only):
    """1年分・全市町村（合併前は旧市町村コード）の総人口を1回のAPI呼び出しで取る。"""
    table = CENSUS_TABLES[year]
    areas = set()
    for name, code in CURRENT_AREA.items():
        if name in MERGE_YEARS and year in MERGE_YEARS[name]:
            areas.update(OLD_AREA[name])
        elif name in RENAME_AREA and year in RENAME_AREA[name]["years"]:
            areas.add(RENAME_AREA[name]["old_code"])
        else:
            areas.add(code)
    body = estat_get(f"{year}_total", table["id"], table["fixed"],
                      {"cdArea": ",".join(sorted(areas)), "cdTime": f"{year}000000"},
                      use_cache_only)
    return values_by_area(body)


def fetch_r7(use_cache_only):
    pop_body = estat_get("2025_population", R7_POP_TABLE["id"], R7_POP_TABLE["fixed"],
                          {"cdArea": ",".join(sorted(CURRENT_AREA.values())), "cdTime": "2025000000"},
                          use_cache_only)
    hh_body = estat_get("2025_households", R7_HH_TABLE["id"], R7_HH_TABLE["fixed"],
                         {"cdArea": ",".join(sorted(CURRENT_AREA.values())), "cdTime": "2025000000"},
                         use_cache_only)
    return values_by_area(pop_body), values_by_area(hh_body)


def fetch_pyramid(use_cache_only):
    """令和2年の男女×年齢5歳階級を全市町村まとめて1回で取る。"""
    body = estat_get("2020_pyramid", PYRAMID_TABLE["id"], PYRAMID_TABLE["fixed"],
                      {"cdArea": ",".join(sorted(CURRENT_AREA.values())),
                       "cdCat02": "1,2", "cdCat03": ",".join(AGE_CODES),
                       "cdTime": "2020000000"},
                      use_cache_only)
    data = body.get("GET_STATS_DATA", {}).get("STATISTICAL_DATA", {})
    info = data.get("DATA_INF")
    vals = info["VALUE"] if info else []
    if isinstance(vals, dict):
        vals = [vals]
    # (area, sex, age_code) -> population
    cells = {}
    for v in vals:
        n = _parse_value(v)
        if n is not None:
            cells[(v["@area"], v["@cat02"], v["@cat03"])] = n
    return cells


def build_pyramid_rows(area_code, cells, problems, name):
    rows = []
    missing = []
    for i, age_code in enumerate(AGE_CODES[:17]):  # 01..17 = 0-4歳..80-84歳（1区分ずつ）
        male = cells.get((area_code, "1", age_code))
        female = cells.get((area_code, "2", age_code))
        if male is None or female is None:
            missing.append(AGES[i])
            continue
        rows.append({"age": AGES[i], "male": male, "female": female})
    # 85歳以上はAGE_CODES[17:21]（85-89, 90-94, 95-99, 100歳以上）を合算
    male85 = female85 = 0
    ok85 = True
    for age_code in AGE_CODES[17:21]:
        m, f = cells.get((area_code, "1", age_code)), cells.get((area_code, "2", age_code))
        if m is None or f is None:
            ok85 = False
            break
        male85 += m
        female85 += f
    if ok85:
        rows.append({"age": "85+", "male": male85, "female": female85})
    else:
        missing.append("85+")
    if missing:
        problems.append(f"{name}: ピラミッドの年齢区分 {missing} が取得できませんでした")
        return None
    return rows


def normalize(use_cache_only):
    problems = []
    reconciled_notes = {}

    year_values = {}
    for year in CENSUS_TABLES:
        try:
            year_values[year] = fetch_year_totals(year, use_cache_only)
        except EstatError as ex:
            problems.append(str(ex))
            year_values[year] = {}

    try:
        r7_pop, r7_hh = fetch_r7(use_cache_only)
    except EstatError as ex:
        problems.append(str(ex))
        r7_pop, r7_hh = {}, {}

    try:
        pyramid_cells = fetch_pyramid(use_cache_only)
    except EstatError as ex:
        problems.append(str(ex))
        pyramid_cells = {}

    out = []
    for name, code in CURRENT_AREA.items():
        trend = []
        for year in sorted(CENSUS_TABLES):
            vals = year_values.get(year, {})
            if name in MERGE_YEARS and year in MERGE_YEARS[name]:
                codes = OLD_AREA[name]
                found = {c: vals[c] for c in codes if c in vals}
                if len(found) != len(codes):
                    missing = sorted(set(codes) - found.keys())
                    problems.append(f"{name} {year}年: 旧市町村コード {missing} が取得できませんでした（合算できないためnullにします）")
                    trend.append({"year": year, "date": f"{year}-10-01", "population": None, "boundary": None})
                    continue
                pop = sum(found.values())
                trend.append({"year": year, "date": f"{year}-10-01", "population": pop, "boundary": "reconciled"})
                reconciled_notes.setdefault(name, []).append(year)
                continue
            area_code = code
            if name in RENAME_AREA and year in RENAME_AREA[name]["years"]:
                area_code = RENAME_AREA[name]["old_code"]
            pop = vals.get(area_code)
            if pop is None:
                problems.append(f"{name} {year}年: 市町村コード {area_code} のデータが取得できませんでした")
            trend.append({"year": year, "date": f"{year}-10-01", "population": pop, "boundary": None})

        r7pop = r7_pop.get(code)
        r7hh = r7_hh.get(code)
        if r7pop is None:
            problems.append(f"{name} 2025年（速報）: データが取得できませんでした")
        trend.append({"year": 2025, "date": "2025-10-01", "population": r7pop,
                      "households": r7hh, "boundary": None, "provisional": True})

        pyramid = build_pyramid_rows(code, pyramid_cells, problems, name)

        out.append({
            "code": code, "name": name,
            "trend": trend,
            "pyramid_year": 2020,
            "pyramid": pyramid,
        })

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"wrote {OUT}  市町村数={len(out)}")

    report = {
        "counts": {
            "municipalities": len(out),
            "failures": len(problems),
            "reconciled_municipalities": len(reconciled_notes),
        },
        "reconciled_years": reconciled_notes,
        "failures": problems,
    }
    REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"wrote {REPORT}")

    if reconciled_notes:
        print("\n合併前を旧市町村の合算値に置き換えた市町村:")
        for name, years in reconciled_notes.items():
            print(f"  - {name}: {sorted(set(years))}")
    if problems:
        print("\n要対応:")
        for p in problems:
            print("  -", p)
    return len(problems) == 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--normalize-only", action="store_true",
                     help="APIを呼ばず data/raw/census/ のキャッシュのみで再構築する")
    a = ap.parse_args()
    sys.exit(0 if normalize(a.normalize_only) else 1)
