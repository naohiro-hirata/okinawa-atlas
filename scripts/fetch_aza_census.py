#!/usr/bin/env python3
"""令和2年国勢調査 小地域集計（沖縄県・町丁字等別）から字ごとの年齢構成を取り出す。

出典: 政府統計の総合窓口（e-Stat）のファイルダウンロード。
統計表「男女，年齢（5歳階級）別人口，平均年齢及び総年齢－町丁・字等」（沖縄県）
  https://www.e-stat.go.jp/stat-search/file-download?statInfId=000032163785&fileKind=1
API（getStatsList/getStatsData）には統計表IDが存在せず、CSVファイルの直接配布のみ
（Shift-JIS）。1ファイルに沖縄県41市町村ぶんが入っている。

このファイルには2つの壁がある。
  1. 階層構造: 町丁字コードは市区町村→（大字・町）→（丁目）の可変深さの木。
     丁目を持つ大字・町は、自分自身の行（小計）と丁目ごとの行の両方を含むため、
     丁目がある場合は丁目の行だけを使い、小計行は捨てないと二重計上になる。
  2. 秘匿処理: 人口が少ない町丁字は近隣に合算され、値が "X" になる。この場合は
     復元不可能なので、実測0（"-"）とは区別して「秘匿処理」のまま未対応にする。

住基Excel側（data/generated/population_aza.json）の字名とは名前空間が別なので、
丁目の漢数字を算用数字に揃えたうえで完全一致するものだけを対応付ける。
一致しないものは無理に対応付けず、未対応のまま data/aza_census_crosswalk.csv と
population_aza_age.json に残す。

    python scripts/fetch_aza_census.py            # ダウンロード＋正規化
    python scripts/fetch_aza_census.py --normalize-only   # data/raw/census_aza/ のキャッシュのみで再構築

出力:
  data/raw/census_aza/aza_age_2020.csv          原本CSV（gitignore対象）
  data/aza_census_crosswalk.csv                  住基側の字と国勢調査側の町丁字の対応表
  data/generated/population_aza_age.json         字ごとの年齢構成（対応が取れたもののみ数値あり）
  data/generated/aza_census_quality_report.json  対応状況・検証結果
"""
from __future__ import annotations

import csv
import json
import re
import sys
import urllib.request
from pathlib import Path

URL = "https://www.e-stat.go.jp/stat-search/file-download?statInfId=000032163785&fileKind=1"
RAW = Path("data/raw/census_aza/aza_age_2020.csv")
CROSSWALK = Path("data/aza_census_crosswalk.csv")
POPULATION_AZA = Path("data/generated/population_aza.json")
OUT = Path("data/generated/population_aza_age.json")
REPORT = Path("data/generated/aza_census_quality_report.json")

# AGES と揃える（app/template.html）。個々の5歳階級17区分＋85歳以上（4区分合算）。
AGES = ["0-4", "5-9", "10-14", "15-19", "20-24", "25-29", "30-34", "35-39", "40-44",
        "45-49", "50-54", "55-59", "60-64", "65-69", "70-74", "75-79", "80-84", "85+"]
# CSVの列インデックス（0始まり）。実測して確認済み（那覇市の再掲区分の合計が一致することで検証済み）。
COL_SEX, COL_MUNI_CODE, COL_AREA_CODE, COL_LEVEL = 1, 2, 3, 4
COL_HITOKU, COL_HITOKU_SAKI = 5, 6
COL_MUNI_NAME, COL_OAZA, COL_CHOME, COL_TOTAL = 9, 10, 11, 12
COL_AGE_START = 13   # 0～4歳
COL_AGE_85PLUS = [30, 31, 32, 33]  # 85～89, 90～94, 95～99, 100歳以上

# 米軍基地関連と思われ、未対応(unmatched)のうち国勢調査側に該当する行が一件も
# 見つからなかった字。(1)対象市町村の大字・町名の全件リストに同名の行が存在しない、
# (2)名前のある区域の合計が市町村計と完全に一致する（差0＝取りこぼしがない）ことを
# 実測して確認した（2026-08-01）。ただし総務省統計局の「国勢調査は米軍基地内を
# 対象外とする」という公式な説明文書までは確認できていないため、確定情報ではなく
# 推定として扱う。画面側も「確定」ではなく「可能性（推定）」と表示すること。
POSSIBLY_OUT_OF_SCOPE = {
    ("宜野湾市", "キャンプフォスター"),
    ("沖縄市", "CAMPSHIELDS"),
    ("沖縄市", "KADENAAIRBASE"),
    ("沖縄市", "キャンプフォスター"),
    ("沖縄市", "シールズ基地"),
    ("沖縄市", "シールズ基地内"),
    ("沖縄市", "ライカム基地"),
    ("沖縄市", "ライカム基地内"),
    ("沖縄市", "嘉手納基地内"),
    ("沖縄市", "知花エリア内"),
    ("沖縄市", "知花ハウジングエリア"),
    ("沖縄市", "知花ハウジングエリア内"),
    ("浦添市", "キャンプキンザー内"),
    ("北中城村", "軍施設内"),
}

# 公式な確証（総務省統計局の説明文書など）が見つかった項目だけをここに移す。
# 値は出典・確認日など。ここに移った項目は画面表示が「推定」から「確定」に変わる。
# 例: ("沖縄市", "嘉手納基地内"): "総務省統計局『○○』（2027-xx-xx確認）"
CONFIRMED_OUT_OF_SCOPE = {}

KANJI_DIGIT = {"〇": 0, "一": 1, "二": 2, "三": 3, "四": 4,
               "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
CHOME_RE = re.compile(r"^(.*?)([一二三四五六七八九十]+)丁目$")


def kanji_num_to_int(s):
    """漢数字（一〜九，十，十一〜十九，二十…）を整数に変換する。読めなければNone。"""
    if s == "十":
        return 10
    if "十" in s:
        left, _, right = s.partition("十")
        tens = KANJI_DIGIT.get(left, 1) if left else 1
        ones = KANJI_DIGIT.get(right, 0) if right else 0
        if (left and left not in KANJI_DIGIT) or (right and right not in KANJI_DIGIT):
            return None
        return tens * 10 + ones
    return KANJI_DIGIT.get(s)


def canonical_chome(name):
    """字名の末尾が漢数字丁目なら算用数字丁目に揃える。住基側・国勢調査側の両方に使う。"""
    m = CHOME_RE.match(name)
    if not m:
        return name
    n = kanji_num_to_int(m.group(2))
    if n is None:
        return name
    return f"{m.group(1)}{n}丁目"


def strip_ji_key(canonical_name):
    """先頭の「字」を落とした照合用キー（丁目表記は呼び出し側で揃え済みの前提）。

    国勢調査の大字・町名は「字」を付けたまま収録される（例：字東江上）が、住基側の
    表示名は fetch_aza_population.py の正規化で「字」が落ちていることが多い
    （例：東江上）。どちらも同じ実体を指すことが多いので、完全一致で見つからない
    ときだけこのキーで拾う。ただし「字X」と「X」が別実体として両方存在する市町村
    （那覇市の「字上之屋」と「上之屋」など）では、この関数でキーが衝突するため、
    衝突する組は完全一致でしか対応付けない（呼び出し側の stripped_groups で判定）。
    """
    return canonical_name[1:] if canonical_name.startswith("字") else canonical_name


def parse_val(s):
    """"-"=実測0、"X"=秘匿（値なし）、空欄=値なし。それ以外は整数。"""
    if s == "-":
        return 0
    if s in ("X", ""):
        return None
    return int(s)


def download():
    RAW.parent.mkdir(parents=True, exist_ok=True)
    if RAW.exists():
        print(f"  = {RAW} (既存)")
        return
    req = urllib.request.Request(URL, headers={"User-Agent": "okinawa-atlas/1.0"})
    with urllib.request.urlopen(req, timeout=60) as r, RAW.open("wb") as f:
        f.write(r.read())
    print(f"  + {RAW} <- {URL}")


def read_rows():
    with RAW.open(encoding="cp932", newline="") as f:
        rows = list(csv.reader(f))
    # 先頭4行はタイトル・注記。5行目（インデックス4）が見出し。
    header = rows[4]
    if header[COL_OAZA] != "大字・町名" or header[COL_TOTAL] != "総数":
        raise ValueError(f"{RAW}: 見出し列の並びが想定と違います。手動で確認してください。 {header}")
    out = []
    for row in rows[5:]:
        if len(row) <= COL_AGE_85PLUS[-1]:
            continue
        out.append({
            "sex": row[COL_SEX], "muni_code": row[COL_MUNI_CODE], "area_code": row[COL_AREA_CODE],
            "level": row[COL_LEVEL], "hitoku": row[COL_HITOKU], "hitoku_saki": row[COL_HITOKU_SAKI],
            "muni_name": row[COL_MUNI_NAME], "oaza": row[COL_OAZA], "chome": row[COL_CHOME],
            "total": parse_val(row[COL_TOTAL]),
            "ages": [parse_val(row[COL_AGE_START + i]) for i in range(17)],
            "age85plus": [parse_val(row[c]) for c in COL_AGE_85PLUS],
        })
    return out


def extract_leaves(rows):
    """階層から二重計上しない「葉」の町丁字だけを取り出す（分類は総数の行で決める）。

    丁目という名前で分かれている場合（例：曙→曙一〜三丁目）で、丁目の合計が
    大字・町自体の行の値と一致するときだけ、丁目ごとの行を葉として使い、
    大字・町自体の行（小計）は捨てる。合計が一致しない場合（例：字古我知の
    「内原区」は381人中62人だけを指す名前のない行政区で、字全体を尽くさない）は、
    信頼できないので子の行は使わず、大字・町自体の行を1つの葉として使う。
    同様に、名前のない基本単位区だけで細分化されている場合（例：石垣市 字登野城が
    "001000"「"001001"という無名のコードに分かれる）も、字そのものに住居表示上の
    区分がないということなので、細分化前（レベルが最も浅い行＝字全体の値）を葉として使う。
    """
    groups = {}
    for r in rows:
        if r["sex"] != "総数" or r["level"] not in ("2", "3", "4"):
            continue
        groups.setdefault((r["muni_code"], r["oaza"]), []).append(r)

    leaves, validation = [], []
    for (muni_code, oaza), items in groups.items():
        chome_rows = [r for r in items if r["chome"]]
        parent = next((r for r in items if not r["chome"]), None)
        use_children = False
        if chome_rows:
            if parent is None or parent["total"] is None or any(c["total"] is None for c in chome_rows):
                # 秘匿された子を含む場合は合計で検証できないが、丁目という名前がある
                # 以上は本物の細分化とみなして子を使う（個々の秘匿は別途処理する）。
                use_children = True
            else:
                children_sum = sum(c["total"] for c in chome_rows)
                use_children = children_sum == parent["total"]
                validation.append({
                    "municipality": (parent or chome_rows[0])["muni_name"], "oaza": oaza,
                    "parent_total": parent["total"], "children_sum": children_sum,
                    "children": [c["chome"] for c in chome_rows],
                })
        if use_children:
            leaves.extend(chome_rows)
        else:
            # 子が全体を尽くしていない（名前のない行政区など）か、そもそも子がない。
            # 最も浅い行（＝字全体の値）だけを葉として使う。
            leaves.append(min(items, key=lambda r: int(r["level"])))
    return leaves, validation


def build_pyramid(sex_rows_by_code, muni_code, area_code):
    """(muni_code, area_code) の男・女の行から年齢構成を作る。秘匿ならNoneを返す。"""
    male = sex_rows_by_code.get((muni_code, area_code, "男"))
    female = sex_rows_by_code.get((muni_code, area_code, "女"))
    if male is None or female is None:
        return None
    rows = []
    for i, age in enumerate(AGES[:17]):
        m, f = male["ages"][i], female["ages"][i]
        if m is None or f is None:
            return None
        rows.append({"age": age, "male": m, "female": f})
    m85 = sum(male["age85plus"]) if all(v is not None for v in male["age85plus"]) else None
    f85 = sum(female["age85plus"]) if all(v is not None for v in female["age85plus"]) else None
    if m85 is None or f85 is None:
        return None
    rows.append({"age": "85+", "male": m85, "female": f85})
    return rows


def normalize(use_cache_only):
    if not use_cache_only:
        download()
    elif not RAW.exists():
        print(f"{RAW} がありません（--normalize-only ではダウンロードできません）", file=sys.stderr)
        return False

    rows = read_rows()
    leaves, validation = extract_leaves(rows)
    sex_rows_by_code = {(r["muni_code"], r["area_code"], r["sex"]): r for r in rows}

    # まず「字」を残したままの完全一致キーで索引する（丁目表記だけ揃える）。国勢調査側の
    # 生の字名はこの時点では重複しないはず（同じ実体が2行になることはない）。
    raw_index = {}   # muni_name -> canonical_chome(census_name) -> leaf info
    dup_problems = []
    for leaf in leaves:
        total_row = sex_rows_by_code.get((leaf["muni_code"], leaf["area_code"], "総数"))
        suppressed = leaf["hitoku"] == "秘匿地域" or (total_row and total_row["total"] is None)
        census_name = leaf["oaza"] + leaf["chome"]
        raw_key = canonical_chome(census_name)
        info = {
            "muni_name": leaf["muni_name"], "census_name": census_name,
            "area_code": leaf["area_code"], "suppressed": suppressed,
            "hitoku_saki": leaf["hitoku_saki"],
            "pyramid": None if suppressed else build_pyramid(sex_rows_by_code, leaf["muni_code"], leaf["area_code"]),
        }
        bucket = raw_index.setdefault(leaf["muni_name"], {})
        if raw_key in bucket:
            dup_problems.append(f'{leaf["muni_name"]} {raw_key!r}: 国勢調査側で正規化キーが衝突（{bucket[raw_key]["census_name"]} / {census_name}）')
            continue
        bucket[raw_key] = info

    # 「字」を落としたキーでも引けるようにする。ただし同じ市町村内に「字X」と「X」の
    # 両方が別実体として存在する場合（那覇市の「字上之屋」と「上之屋」など）は
    # 落とすと衝突するので、そのときは字なしの照合をあきらめる（完全一致のみ残す）。
    stripped_groups = {}  # muni_name -> stripped_key -> [raw_key, ...]
    for muni_name, bucket in raw_index.items():
        g = stripped_groups.setdefault(muni_name, {})
        for raw_key in bucket:
            g.setdefault(strip_ji_key(raw_key), []).append(raw_key)

    def lookup(muni_name, aza_name):
        bucket = raw_index.get(muni_name, {})
        raw_key = canonical_chome(aza_name)
        if raw_key in bucket:
            return bucket[raw_key]
        group = stripped_groups.get(muni_name, {}).get(strip_ji_key(raw_key), [])
        if len(group) == 1:
            return bucket[group[0]]
        return None

    juki = json.loads(POPULATION_AZA.read_text(encoding="utf-8"))

    out, crosswalk_rows = [], []
    matched = suppressed_n = unmatched_n = out_of_scope_n = confirmed_n = 0
    for e in juki:
        muni, aza = e["municipality"], e["aza"]
        info = lookup(muni, aza)
        if info is None:
            if (muni, aza) in CONFIRMED_OUT_OF_SCOPE:
                status = "out_of_scope"
                reason = f"out_of_scope: {CONFIRMED_OUT_OF_SCOPE[(muni, aza)]}"
                census_name, pyramid = "", None
                confirmed_n += 1
            elif (muni, aza) in POSSIBLY_OUT_OF_SCOPE:
                status = "possibly_out_of_scope"
                reason = ("possibly_out_of_scope: 米軍基地関連の可能性（国勢調査側に該当する行がなく、"
                          "市町村計は名前のある区域の合計と一致）。総務省統計局の公式な説明は未確認のため推定")
                census_name, pyramid = "", None
                out_of_scope_n += 1
            else:
                status, reason, census_name, pyramid = "unmatched", "国勢調査(令和2年)の町丁・字と一致しません", "", None
                unmatched_n += 1
        elif info["suppressed"]:
            status = "suppressed"
            reason = f'秘匿処理（近隣の町丁・字（{info["hitoku_saki"]}）に合算されており人口が復元できません）'
            census_name, pyramid = info["census_name"], None
            suppressed_n += 1
        elif info["pyramid"] is None:
            status, reason, census_name, pyramid = "unmatched", "国勢調査側の年齢区分が一部欠けており算出できません", info["census_name"], None
            unmatched_n += 1
        else:
            status, reason, census_name, pyramid = "matched", "", info["census_name"], info["pyramid"]
            matched += 1
        out.append({"municipality": muni, "aza": aza, "status": status, "pyramid": pyramid})
        crosswalk_rows.append([muni, aza, census_name, status, reason])

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"wrote {OUT}  字数={len(out)}（一致{matched}／秘匿{suppressed_n}／未対応{unmatched_n}／"
          f"対象外の可能性(推定){out_of_scope_n}／対象外(確定){confirmed_n}）")

    with CROSSWALK.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        f.write("# 住基Excel側の字名（data/generated/population_aza.json）と\n")
        f.write("# 令和2年国勢調査 小地域集計側の町丁・字の対応表。scripts/fetch_aza_census.py が\n")
        f.write("# 丁目の漢数字/算用数字を揃えたうえでの完全一致のみで機械的に作成する（推測はしない）。\n")
        f.write("# status: matched（対応OK）/ suppressed（秘匿処理で復元不可）/\n")
        f.write("#   possibly_out_of_scope（米軍基地関連などで国勢調査の対象区域に含まれない可能性・推定）/\n")
        f.write("#   out_of_scope（同・公式に確証済み）/ unmatched（対応する町丁字が見つからない）\n")
        f.write("# possibly_out_of_scope の判定は scripts/fetch_aza_census.py の POSSIBLY_OUT_OF_SCOPE を、\n")
        f.write("# 確証が取れた場合は CONFIRMED_OUT_OF_SCOPE を編集すること（このファイルは手で編集しない）。\n")
        w.writerow(["municipality", "aza_juki", "census_name", "status", "reason"])
        for row in crosswalk_rows:
            w.writerow(row)
    print(f"wrote {CROSSWALK}")

    report = {
        "counts": {
            "aza_total": len(juki), "matched": matched, "suppressed": suppressed_n,
            "unmatched": unmatched_n, "possibly_out_of_scope": out_of_scope_n,
            "out_of_scope_confirmed": confirmed_n, "collisions": len(dup_problems),
        },
        "level_validation_sample": validation[:20],
        "level_validation_mismatches": [v for v in validation if v["parent_total"] != v["children_sum"]],
        "collisions": dup_problems,
    }
    REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"wrote {REPORT}")
    if dup_problems:
        print("\n要対応（正規化キー衝突）:")
        for p in dup_problems:
            print("  -", p)
    return True


if __name__ == "__main__":
    use_cache_only = "--normalize-only" in sys.argv[1:]
    sys.exit(0 if normalize(use_cache_only) else 1)
