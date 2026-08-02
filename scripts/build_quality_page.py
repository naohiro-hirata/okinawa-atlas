#!/usr/bin/env python3
"""data/generated/*_quality_report.json から public/quality.html を作る。

サイトの閲覧者が「この数字はどこまで確かなのか」を自分で確かめられるようにする
ための案内ページ。技術者でなくても読めるよう、各レポートが何を意味するかを
日本語で書き添え、生JSONへのリンクも併記する。

index.html と同じく fetch を使わず、数値はビルド時に埋め込む（file:// でも開ける）。
生JSONは public/reports/ にコピーして直リンクできるようにする。

使い方: python scripts/build_quality_page.py [--outfile public/quality.html]
"""
import argparse, html, json, shutil
from pathlib import Path

ap = argparse.ArgumentParser()
ap.add_argument("--datadir", default="data/generated")
ap.add_argument("--outfile", default="public/quality.html")
a = ap.parse_args()

D = Path(a.datadir)
OUT = Path(a.outfile)
REPORTS = OUT.parent / "reports"


def load(name):
    """レポートを読む。無ければ None（原本未配置でCIが落ちないように）。"""
    p = D / f"{name}.json"
    if not p.exists():
        return None
    return json.loads(p.read_text(encoding="utf-8"))


def e(s):
    return html.escape(str(s))


def n(v):
    return f"{v:,}" if isinstance(v, int) else e(v)


def stat(label, value, sub=""):
    sub = f'<div class="sub">{e(sub)}</div>' if sub else ""
    return f'<div class="stat"><div class="v">{n(value)}</div><div class="l">{e(label)}</div>{sub}</div>'


def check_row(label, actual, expected, note):
    """確認ポイント1行。期待値と違うときだけ目立たせる（判断は人がする）。"""
    ok = actual == expected
    mark = "一致" if ok else "要確認"
    cls = "ok" if ok else "warn"
    return (f'<tr class="{cls}"><td>{e(label)}</td><td class="n">{n(actual)}</td>'
            f'<td class="n">{n(expected)}</td><td><span class="mark {cls}">{mark}</span></td>'
            f'<td class="note">{e(note)}</td></tr>')


def section(anchor, title, script, source, lead, body, jsonname, present):
    if not present:
        return (f'<section id="{anchor}"><h2>{e(title)}</h2>'
                f'<p class="miss">このレポートはまだ生成されていません（原本が未配置の可能性があります）。</p></section>')
    return f"""<section id="{anchor}">
  <h2>{e(title)}</h2>
  <dl class="meta">
    <dt>生成するスクリプト</dt><dd><code>{e(script)}</code></dd>
    <dt>元データ</dt><dd>{source}</dd>
    <dt>生ファイル</dt><dd><a href="reports/{e(jsonname)}.json">{e(jsonname)}.json</a>（JSON形式）</dd>
  </dl>
  <p class="lead">{lead}</p>
  {body}
</section>"""


parts = []

# ---------------------------------------------------------------- 調査Excel
q = load("quality_report")
if q:
    c = q["counts"]
    body = f"""
  <div class="stats">
    {stat("市町村", c["municipalities"])}
    {stat("施策の回答行", c["policy_rows"])}
    {stat("空き家の回答", c["akiya_municipalities"])}
    {stat("定住促進住宅", c["housing_properties"], "物件")}
  </div>
  <h3>確認ポイント</h3>
  <table class="tbl">
    <thead><tr><th>項目</th><th class="n">現在</th><th class="n">想定</th><th>状態</th><th>意味</th></tr></thead>
    <tbody>
      {check_row("市町村コードの突合失敗", len(q["code_join_failures"]), 0,
                 "調査Excelの市町村名を41市町村のコードに結び付けられなかった行。0でなければ名寄せが必要。")}
      <tr><td>備考欄から数値を取れなかった項目</td><td class="n">{len(q["needs_manual_review"])}</td>
        <td class="n">—</td><td><span class="mark info">要確認リスト</span></td>
        <td class="note">推測で数値を作らず null（画面では「回答なし」）にしている項目。下に一覧。</td></tr>
      <tr><td>「最終確認状況」が未確認の市町村</td><td class="n">{len(q["unconfirmed_policy"])}</td>
        <td class="n">—</td><td><span class="mark info">注意して読む</span></td>
        <td class="note">令和6年度以前の回答が残っている可能性がある市町村。</td></tr>
    </tbody>
  </table>

  <h3>「最終確認状況」が未確認の市町村（{len(q["unconfirmed_policy"])}件）</h3>
  <p class="lead">これらの市町村は、令和7年度の調査票で最終確認が取れていません。
  表示されている選択式の回答（相談窓口の有無など）が、令和6年度以前のものである可能性があります。
  各市町村のページ上部にも「R7回答 未確認」と表示しています。</p>
  <p class="chips">{"".join(f'<span class="chip">{e(x)}</span>' for x in q["unconfirmed_policy"])}</p>

  <h3>備考欄から数値を取り出せなかった項目（{len(q["needs_manual_review"])}件）</h3>
  <p class="lead">自由記述の備考欄に数値らしきものがあっても、何の数値か確実に判断できないものは
  <strong>推測せず空欄のまま</strong>にしています。画面では「回答なし」と表示されます。
  以下がその全件で、元の記述をそのまま載せています。</p>
  <table class="tbl">
    <thead><tr><th>市町村</th><th>項目</th><th>備考欄の原文</th></tr></thead>
    <tbody>
      {"".join(f'<tr><td>{e(r["municipality"])}</td><td><code>{e(r["field"])}</code></td>'
               f'<td class="raw">{e(r["text"])}</td></tr>' for r in q["needs_manual_review"])}
    </tbody>
  </table>"""
    parts.append(section(
        "survey", "① 移住定住・空き家の取り組み状況調査", "scripts/parse_survey.py",
        "自社ヒアリング調査（令和7年度）",
        "市町村への聞き取り調査を集計したものです。施策の状況、空き家バンク、定住促進住宅などが含まれます。",
        body, "quality_report", True))
else:
    parts.append(section("survey", "① 移住定住・空き家の取り組み状況調査", "", "", "", "", "quality_report", False))

# ---------------------------------------------------------------- 字別人口
az = load("aza_quality_report")
if az:
    c = az["counts"]
    yrs = az["records_per_year"]
    body = f"""
  <div class="stats">
    {stat("字（名寄せ後）", c["aza"])}
    {stat("市町村", c["municipalities"])}
    {stat("読み込んだ年", c["years_loaded"], "2011〜2025年")}
    {stat("表記ゆれの統合", c["normalized_groups"], "グループ")}
  </div>
  <h3>確認ポイント</h3>
  <table class="tbl">
    <thead><tr><th>項目</th><th class="n">現在</th><th class="n">想定</th><th>状態</th><th>意味</th></tr></thead>
    <tbody>
      {check_row("変換の失敗", c["failures"], 0, "1件でもあればその年の取り込みを中止する設計。")}
      {check_row("未知の市町村名", c["unknown_municipalities"], 0,
                 "41市町村の正式名に無い名前。誤字か新表記かを人が確認するまで取り込まない。")}
      {check_row("原本の誤字を補正した件数", c["muni_name_corrections"], 8,
                 "県の原本に実在する誤字（「宜野湾志市」など）を変換時だけ直した件数。増えたら新しい誤字が出た合図。")}
      <tr><td>年数が少ない系列</td><td class="n">{c["thin_series"]}</td><td class="n">—</td>
        <td><span class="mark info">要確認リスト</span></td>
        <td class="note">区画再編で途中から現れた／消えた字。勝手につながず、別系列のままにしている。</td></tr>
    </tbody>
  </table>

  <h3>年ごとの収録字数</h3>
  <p class="lead">字の数は年によって変わります（住居表示の実施、丁目の新設など）。
  <strong>推移が不連続な字を勝手に補間したり結合したりはしていません。</strong></p>
  <table class="tbl compact">
    <thead><tr>{"".join(f"<th class='n'>{e(y)}</th>" for y in yrs)}</tr></thead>
    <tbody><tr>{"".join(f"<td class='n'>{n(v)}</td>" for v in yrs.values())}</tr></tbody>
  </table>
  <p class="chartnote">平成23〜25年（2011〜2013）は3月31日現在、平成26年（2014）以降は1月1日現在で
  基準日が異なります。グラフの横軸は年番号ではなく実際の日付で描いています。</p>

  <h3>原本の誤字を補正した箇所（{len(az["muni_name_corrections"])}件）</h3>
  <p class="lead">県が公表しているExcelそのものに市町村名の誤字があり、そのままだと
  以降の字がまるごと別の市町村に集計されてしまいます。
  <strong>原本は書き換えず、読み込むときだけ</strong>正しい名前に読み替えています。</p>
  <table class="tbl">
    <thead><tr><th>年</th><th>ファイル</th><th class="n">行</th><th>原本の表記</th><th>読み替え先</th></tr></thead>
    <tbody>
      {"".join(f'<tr><td class="n">{e(r["year"])}</td><td><code>{e(r["file"])}</code></td>'
               f'<td class="n">{e(r["row_excel"])}</td><td class="raw">{e(r["found"])}</td>'
               f'<td>{e(r["corrected_to"])}</td></tr>' for r in az["muni_name_corrections"])}
    </tbody>
  </table>"""
    parts.append(section(
        "aza", "② 字別の人口・世帯数", "scripts/fetch_aza_population.py",
        "沖縄県「市町村の町字別住民基本台帳人口及び世帯数」（平成23年〜令和7年）",
        "字ごとの人口推移のもとになっているデータです。字の名前は年によって表記が揺れるため、"
        "同じ字だと判断できたものだけをまとめています（判断の方針は <code>docs/aza-matching-policy.md</code>）。",
        body, "aza_quality_report", True))
else:
    parts.append(section("aza", "② 字別の人口・世帯数", "", "", "", "", "aza_quality_report", False))

# ---------------------------------------------------------------- 市町村人口
cs = load("census_quality_report")
if cs:
    c = cs["counts"]
    rec = cs["reconciled_years"]
    body = f"""
  <div class="stats">
    {stat("市町村", c["municipalities"])}
    {stat("取得失敗", c["failures"])}
    {stat("合算した市町村", c["reconciled_municipalities"])}
  </div>
  <h3>確認ポイント</h3>
  <table class="tbl">
    <thead><tr><th>項目</th><th class="n">現在</th><th class="n">想定</th><th>状態</th><th>意味</th></tr></thead>
    <tbody>
      {check_row("取得失敗", c["failures"], 0, "e-Stat から値を取れなかった市町村・年。")}
    </tbody>
  </table>

  <h3>合併前の年を旧市町村の合算で埋めている市町村（{len(rec)}件）</h3>
  <p class="lead">これらの市町村は、下の年について<strong>合併前の旧市町村の実測値を単純に足し合わせた値</strong>です
  （推計ではなく、公表されている実測値の合計）。市域が今と違う期間なので、
  グラフでは<strong>薄い破線</strong>にして区別しています。</p>
  <table class="tbl">
    <thead><tr><th>市町村</th><th>合算している年</th></tr></thead>
    <tbody>
      {"".join(f'<tr><td>{e(k)}</td><td class="n">{e("・".join(str(y) for y in v))}</td></tr>'
               for k, v in rec.items())}
    </tbody>
  </table>
  <p class="chartnote">豊見城市は村から市になったときにコードが変わっただけで市域は変わっていないため、
  区別せず同じ系列としてつないでいます。</p>"""
    parts.append(section(
        "census", "③ 市町村別の人口推移", "scripts/fetch_census.py",
        "国勢調査 1980〜2020年（e-Stat API、回ごとに別の統計表）＋令和7年国勢調査 速報集計",
        "市町村ページの人口推移グラフのもとになっているデータです。"
        "令和7年は速報集計のため男女別人口と世帯総数のみで、年齢別のデータはありません。",
        body, "census_quality_report", True))
else:
    parts.append(section("census", "③ 市町村別の人口推移", "", "", "", "", "census_quality_report", False))

# ------------------------------------------------------------ 字別年齢構成
ac = load("aza_census_quality_report")
if ac:
    c = ac["counts"]
    mm = ac["level_validation_mismatches"]
    tot = c["aza_total"]
    pct = c["matched"] / tot * 100 if tot else 0
    body = f"""
  <div class="stats">
    {stat("実データを表示", c["matched"], f"全{tot:,}字の{pct:.0f}%")}
    {stat("秘匿処理", c["suppressed"], "復元できない")}
    {stat("未対応", c["unmatched"], "区分が一致しない")}
    {stat("対象外の可能性", c["possibly_out_of_scope"], "推定・確定ではない")}
  </div>
  <h3>4つの状態の意味</h3>
  <p class="lead">字のページの「年齢構成」は、対応が取れた字だけが実データです。
  取れなかった字は<strong>斜線とタグを付けたサンプル値</strong>を表示しており、実際の数値ではありません。</p>
  <table class="tbl">
    <thead><tr><th>状態</th><th class="n">件数</th><th>画面の表示</th><th>意味</th></tr></thead>
    <tbody>
      <tr><td>対応済み</td><td class="n">{n(c["matched"])}</td><td>実データ（斜線なし）</td>
        <td class="note">令和2年国勢調査の小地域と対応が取れた字。</td></tr>
      <tr><td>秘匿処理</td><td class="n">{n(c["suppressed"])}</td><td><span class="tag">秘匿処理</span></td>
        <td class="note">人口が少なく近隣の字に合算されているため、元の値を復元できない。実測値の0とは別物。</td></tr>
      <tr><td>未対応</td><td class="n">{n(c["unmatched"])}</td><td><span class="tag">未対応</span></td>
        <td class="note">国勢調査の小地域区分と名前が一致しない。多くは2020年の調査より後にできた新しい丁目。</td></tr>
      <tr><td>対象外の可能性（推定）</td><td class="n">{n(c["possibly_out_of_scope"])}</td>
        <td><span class="tag">対象外の可能性（推定）</span></td>
        <td class="note">米軍基地関連とみられる区域。国勢調査側に該当する行が無く、名前のある区域の合計だけで
        市町村計と一致することは確認済みだが、<strong>総務省統計局による公式な説明は確認できておらず推定にとどまる</strong>。</td></tr>
      <tr><td>対象外（確定）</td><td class="n">{n(c["out_of_scope_confirmed"])}</td>
        <td><span class="tag">対象外（確定）</span></td>
        <td class="note">公式な確証が取れた場合のみここに移す。現在は0件。</td></tr>
    </tbody>
  </table>

  <h3>確認ポイント</h3>
  <table class="tbl">
    <thead><tr><th>項目</th><th class="n">現在</th><th class="n">想定</th><th>状態</th><th>意味</th></tr></thead>
    <tbody>
      {check_row("名前の衝突", c["collisions"], 0,
                 "同じ市町村内で同じ字名に2つの実体が当たってしまったケース。1件でもあれば対応付けを見直す。")}
      <tr><td>丁目の合計が字全体と合わない</td><td class="n">{len(mm)}</td><td class="n">—</td>
        <td><span class="mark info">要確認リスト</span></td>
        <td class="note">丁目に見える名前が字全体を尽くしていないケース。この場合は丁目ではなく字全体の行を使っている。</td></tr>
    </tbody>
  </table>

  <h3>丁目の合計が字全体と合わなかった字（{len(mm)}件）</h3>
  <p class="lead">国勢調査のデータには、字の下に「内原区」のような行政区が入っていることがあります。
  これらは字全体を分割しきっていないため、そのまま使うと人口を大幅に取りこぼします。
  合計が一致しない場合は<strong>字全体の行を使う</strong>ようにしています。</p>
  <table class="tbl">
    <thead><tr><th>市町村</th><th>字</th><th class="n">字全体</th><th class="n">下位の合計</th><th>下位の名前</th></tr></thead>
    <tbody>
      {"".join(f'<tr><td>{e(r["municipality"])}</td><td>{e(r["oaza"])}</td>'
               f'<td class="n">{n(r["parent_total"])}</td><td class="n">{n(r["children_sum"])}</td>'
               f'<td class="raw">{e("、".join(r["children"]))}</td></tr>' for r in mm)}
    </tbody>
  </table>"""
    parts.append(section(
        "aza_census", "④ 字別の年齢構成", "scripts/fetch_aza_census.py",
        "令和2年国勢調査 小地域集計（町丁・字等別／5歳階級×男女）",
        "字のページの人口ピラミッドのもとになっているデータです。"
        "住民基本台帳の字名と国勢調査の小地域名は切り方が違うため、"
        "<strong>丁目表記と「字」の有無を揃えたうえでの完全一致だけ</strong>を対応とし、"
        "一致しないものは推測せず未対応のままにしています。",
        body, "aza_census_quality_report", True))
else:
    parts.append(section("aza_census", "④ 字別の年齢構成", "", "", "", "", "aza_census_quality_report", False))

PAGE = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>データの品質と限界 — 沖縄 地域データアトラス</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&family=Roboto+Condensed:wght@400;700&display=swap" rel="stylesheet">
<style>
:root{
  --ground:#E3E1D9; --surface:#F7F6F2; --surface2:#EDEBE4;
  --ink:#151F23; --ink2:#4E5A5F; --ink3:#7B857F;
  --ai:#1D4E5F; --ai2:#2F7288; --kawara:#BE4F2D; --basho:#4C6B4A; --sand:#A79E90;
  --line:#CDC9BE; --line2:#DDD9CF;
  --jp:"Noto Sans JP",system-ui,sans-serif;
  --num:"Roboto Condensed","Noto Sans JP",sans-serif;
}
*{box-sizing:border-box}
body{margin:0;background:var(--ground);color:var(--ink);font:400 14px/1.75 var(--jp)}
a{color:var(--ai2)}
.wrap{max-width:860px;margin:0 auto;padding:0 20px 80px}
header.top{border-bottom:1px solid var(--line);background:var(--surface);padding:14px 20px 12px;margin-bottom:26px}
header.top .inner{max-width:860px;margin:0 auto;display:flex;align-items:baseline;gap:16px;flex-wrap:wrap}
.brand{font-weight:700;font-size:15px;letter-spacing:.04em}
.brand span{color:var(--ai)}
.eyebrow{font:700 10px/1 var(--num);letter-spacing:.16em;text-transform:uppercase;color:var(--ink3)}
h1{font-size:26px;letter-spacing:.02em;margin:10px 0 4px}
h2{font-size:18px;margin:0 0 4px;padding-bottom:8px;border-bottom:2px solid var(--ai)}
h3{font:700 11px/1 var(--num);letter-spacing:.14em;text-transform:uppercase;color:var(--ink2);margin:26px 0 9px}
section{background:var(--surface);border:1px solid var(--line);padding:20px 22px 26px;margin-bottom:20px}
p.lead{margin:10px 0 0;font-size:13px}
p.chartnote{font-size:11.5px;color:var(--ink3);margin-top:8px}
.intro{background:var(--surface);border:1px solid var(--line);border-left:4px solid var(--kawara);padding:16px 18px;margin-bottom:22px;font-size:13px}
.intro p{margin:0 0 8px}.intro p:last-child{margin:0}
dl.meta{margin:12px 0 0;display:grid;grid-template-columns:auto 1fr;gap:5px 14px;font-size:12px}
dl.meta dt{font:700 9.5px/1.8 var(--num);letter-spacing:.1em;text-transform:uppercase;color:var(--ink3);white-space:nowrap}
dl.meta dd{margin:0}
code{font-family:var(--num);background:var(--surface2);padding:1px 5px;border:1px solid var(--line2);font-size:.94em}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:1px;background:var(--line2);margin-top:16px}
.stat{background:var(--surface);padding:11px 13px}
.stat .v{font:700 25px/1.15 var(--num);font-variant-numeric:tabular-nums}
.stat .l{font-size:11px;color:var(--ink2);margin-top:2px}
.stat .sub{font-size:10.5px;color:var(--ink3);margin-top:1px}
table.tbl{width:100%;border-collapse:collapse;font-size:12.5px;margin-top:4px}
table.tbl th{text-align:left;font:700 9.5px/1.7 var(--num);letter-spacing:.09em;text-transform:uppercase;
  color:var(--ink3);border-bottom:1px solid var(--line);padding:6px 9px 6px 0;white-space:nowrap}
table.tbl td{border-bottom:1px solid var(--line2);padding:7px 9px 7px 0;vertical-align:top}
table.tbl td.n,table.tbl th.n{font-family:var(--num);font-variant-numeric:tabular-nums;text-align:right;padding-right:16px;white-space:nowrap}
table.tbl td.note{color:var(--ink2);font-size:11.5px;line-height:1.6}
table.tbl td.raw{color:var(--ink2);font-size:11px;line-height:1.6;word-break:break-all}
table.compact td,table.compact th{padding:5px 7px 5px 0;font-size:11.5px}
.mark{display:inline-block;padding:2px 8px;font:700 9.5px/1.5 var(--num);letter-spacing:.1em;text-transform:uppercase;white-space:nowrap}
.mark.ok{background:#E9EFE6;color:#2C4A2B;border:1px solid var(--basho)}
.mark.warn{background:#F6EAE4;color:#6E2E19;border:1px solid var(--kawara)}
.mark.info{background:var(--surface2);color:var(--ink2);border:1px solid var(--line)}
tr.warn td{background:#FBF3EF}
.tag{display:inline-block;background:var(--kawara);color:#fff;padding:2px 7px;font:700 9px/1.6 var(--num);letter-spacing:.1em;white-space:nowrap}
.chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:9px}
.chip{border:1px solid var(--line);background:var(--surface2);padding:3px 10px;font-size:12px}
.miss{color:var(--ink3);font-size:12.5px;margin-top:10px}
nav.toc{background:var(--surface2);border:1px solid var(--line);padding:13px 16px;margin-bottom:22px;font-size:13px}
nav.toc ol{margin:6px 0 0;padding-left:20px}
nav.toc li{margin:3px 0}
footer{margin-top:30px;padding-top:16px;border-top:1px solid var(--line);font-size:11.5px;color:var(--ink3)}
</style>
</head>
<body>
<header class="top"><div class="inner">
  <div class="brand">沖縄 <span>地域データアトラス</span></div>
  <a href="./">← 地図・データに戻る</a>
</div></header>
<div class="wrap">
  <div class="eyebrow">Data quality</div>
  <h1>データの品質と限界</h1>

  <div class="intro">
    <p>このサイトの数字は、公的統計と市町村への聞き取り調査をもとにしています。
    ただし<strong>すべての項目がすべての市町村について揃っているわけではありません</strong>。</p>
    <p>このページは、どこまでが確かめられた数字で、どこからが「わからない」なのかを、
    変換プログラムが自動で出したレポートそのままで公開するものです。</p>
    <p><strong>出典を確認できない数値は、それらしい値で埋めずに空欄（「回答なし」）にしています。</strong>
    実データに対応できなかった図は、斜線とタグを付けたサンプル値で表示しており、実際の数値ではありません。</p>
  </div>

  <nav class="toc">
    <strong>このページの内容</strong>
    <ol>
      <li><a href="#survey">移住定住・空き家の取り組み状況調査</a></li>
      <li><a href="#aza">字別の人口・世帯数</a></li>
      <li><a href="#census">市町村別の人口推移</a></li>
      <li><a href="#aza_census">字別の年齢構成</a></li>
    </ol>
  </nav>

__SECTIONS__

  <footer>
    <p>各レポートは変換スクリプトが実行のたびに自動生成したものです。
    生のJSONファイルは <code>public/reports/</code> に置いてあり、各節からリンクしています。
    元になった変換スクリプトとデータの取り扱い方針は
    <a href="https://github.com/naohiro-hirata/okinawa-atlas">GitHubリポジトリ</a>で公開しています。</p>
  </footer>
</div>
</body>
</html>
"""

OUT.parent.mkdir(parents=True, exist_ok=True)
OUT.write_text(PAGE.replace("__SECTIONS__", "\n".join(parts)), encoding="utf-8")

# 生JSONもそのまま置いて直リンクできるようにする
REPORTS.mkdir(parents=True, exist_ok=True)
copied = []
for name in ("quality_report", "aza_quality_report",
             "census_quality_report", "aza_census_quality_report"):
    src = D / f"{name}.json"
    if src.exists():
        shutil.copy2(src, REPORTS / f"{name}.json")
        copied.append(name)

print(f"wrote {OUT}  ({OUT.stat().st_size/1024:.0f} KB)")
print(f"copied {len(copied)} reports to {REPORTS}: {', '.join(copied)}")
