# 沖縄 地域データアトラス — プロジェクト指示

沖縄県内の地域データを2階層で見る静的サイト。GitHub Pages（無料枠・publicリポジトリ）で公開する。

- **Layer 1 / 市町村** — 人口推移、年齢5歳階級別人口ピラミッド、移住定住施策と空き家利活用の取り組み状況
- **Layer 2 / 字** — 字ごとの人口推移と年齢構成
- 市町村・字の横断検索と、境界ポリゴンによるマップ選択

## 絶対に守ること

1. **人口の数値を推測で作らない。** 出典を確認できない値は `null` にして、画面には「回答なし」または空状態を出す。もっともらしい数値を埋めるのは、このプロジェクトで最悪の失敗。
2. **サンプル値は必ず視覚的に区別する。** `.dummy` クラス（斜線）と `.dummy-tag`（「ダミー」バッジ）を外さない。実データが入ったら要素ごと消す。半端に外さない。
3. **`0` と「回答なし」を混同しない。** 花ブロックが白いのは未回答、花弁ゼロは値ゼロ。指標・凡例・集計すべてで区別する。
4. **備考欄からの数値抽出は、取れなければ `null` のまま `quality_report.json` に回す。** 正規表現を無理に緩めて誤検出を作らない。
5. **`data/raw/` の原本は加工しない。** 変換結果は `data/generated/` にのみ書く。
6. **`docs/publish-checklist.md` は人が実施するもの。** 「チェックリストを通して」と言われても
   代行せず、各項目を利用者自身が確認できる形に整えて提示する。

## コマンド

```bash
pip install openpyxl pandas xlrd

python scripts/parse_survey.py data/raw/survey/latest.xlsx   # 調査Excel → JSON
python scripts/fetch_aza_population.py                       # 県Excel取得 → 字別人口JSON
python scripts/fetch_aza_population.py --normalize-only      # 再ダウンロードなしで再変換
python scripts/build_app.py --outfile public/index.html      # 単一HTMLをビルド
python -m http.server -d public 8000                         # 確認
```

`parse_survey.py` は毎回 `data/generated/quality_report.json` を出す。作業後は必ずこれを読んで、
`needs_manual_review` と `code_join_failures` が増えていないか確認する。

`fetch_aza_population.py` は `data/generated/aza_quality_report.json` を出す。こちらは
`unknown_municipalities` と `failures` が空であること、`muni_name_corrections` が
想定どおり（現在8件）であることを確認する。

## 構成

```
data/raw/survey/latest.xlsx      毎年の調査Excel（差し替えるだけで更新）
data/raw/aza/                    県の町字別人口Excel（gitignore）
data/generated/                  正規化済みJSON（Actionsが生成）
data/aza_crosswalk.csv           字名の名寄せ辞書（手で育てる）
app/template.html                UI本体。1ファイル、外部JSライブラリなし、SVGは自前描画
scripts/                         ETLとビルド
public/                          Pages公開ディレクトリ
```

`app/template.html` の `/*__DATA__*/` に `build_app.py` がJSONを差し込む。トークンを消さない。

## データソースと落とし穴

| データ | 出典 | 注意 |
|---|---|---|
| 市町村別人口推移 | 国勢調査 1920–2020（e-Stat） | — |
| 同・最新値 | 令和7年国勢調査 人口速報集計（2026-05-29公表） | 男女別人口と世帯総数のみ。年齢別はない |
| 市町村別ピラミッド | 令和2年国勢調査 人口等基本集計 | R7確報は令和8年9月までに公表予定 |
| 字別人口・世帯数 | 沖縄県「市町村の町字別住民基本台帳人口及び世帯数」 | **平成23〜25年は3月31日現在、平成26年以降は1月1日現在。基準日が違う**。推移グラフの横軸は年番号ではなく `date` の実日付で描く |
| 字別 年齢×男女 | 令和2年国勢調査 小地域集計 | R7分は2027年以降の見込み。当面2020年が最新 |
| 字の境界 | e-Stat 統計地理情報システム `r2ka47` | **調査区ベースなので住居表示上の町丁・字と一致しないことがある** |
| 施策・空き家 | 自社ヒアリング調査（年次） | 「最終確認状況」がFALSEの行は過年度回答が残っている可能性 |

**住基Excel内の年次間の名寄せ方針は `docs/aza-matching-policy.md` で決定済み（2026-07-31）・
実装済み（Phase 3）。** 変更前に必ず読むこと。要点は4つ。(1) 先頭の「字」を落とし丁目の漢数字を
算用数字に直したキーで系列をまとめる（`fetch_aza_population.py` の `normalize_key()`）。
**同一年に同じキーの表記が2つ来たらその年の変換を失敗させる**（合算しない。`AzaNameCollision`）。
(2) 表記の例外は `data/aza_crosswalk.csv` に人が書く（`relation=same/typo` のみ合算対象）。
(3) **区画再編で切れた系列はつながない**（`succeeded_by` / `split_into` で関係だけ記録）。
(4) 宜野湾市の `大山` `野嵩` などは同名の丁目と**同じ年に併存する別実体**。統合すると
二重計上になる（`relation=separate`）。字合計と県公表の市町村計は、那覇市・宜野湾市・
久米島町の2011/2015/2020/2025年で突き合わせ済み（差0）。

`data/aza_crosswalk.csv` は上記の年次間名寄せ用。字名は住基Excelと国勢調査小地域でも
切り方が揺れるが、そちらは別ファイル `data/aza_census_crosswalk.csv` に分ける（Phase 4）。
**推移が不連続な字を勝手に補間・結合しない。** 判断が必要なものは crosswalk に載せる前に
必ず確認を取る。

住基Excelは**印刷ページごとに市町村名を再掲する**ため、見出しの誤字がそのページ以降の字を
まるごと別市町村に吸わせる。県の原本に実在する誤字（2017〜2022年の「宜野湾志市」、
2024年の「南大東」）は `fetch_aza_population.py` の `MUNI_TYPOS` で変換時にだけ補正する。
**原本は加工しない。** 41市町村の正式名に無い市町村名が出たらその年の変換を失敗させる方針なので、
新年度分で落ちたら、誤字か新表記かを確認してから `MUNI_TYPOS` に足す。

## 既知の未完了

- 字の人口推移は実データに切り替え済み（`population_aza.json`。名寄せ後1,296字、`aza_crosswalk.csv`
  に38行登録）。市町村の人口推移・年齢構成、字の年齢構成はまだサンプル値。
  `app/template.html` の `demoMuni` / `demoAzaPyramid` が生成元
- 字別人口に原因未確認の異常が1つ。**2019→2020 で県計が +1.57%**（他の年は ±0.5% 程度）、
  しかも全市町村で一様。那覇市・宜野湾市・久米島町では字合計と県公表の市町村計が
  全年一致しているため二重計上ではない。原因未確認のため注記は付けていない
- マップ未実装。`prefView()` 内の `.mapslot` が差し込み位置
- 推定空家数は41市町村中16件しか数値化できていない。基準年も混在（沖縄市の5190件はH30調査値）
- 「最終確認状況」FALSE: 糸満市・豊見城市・宜野座村・北谷町・北中城村・西原町

## UIの方針

配色・書体は決定済みで、`:root` のCSS変数から外れないこと。琉球藍 `--ai`、赤瓦 `--kawara`、
石灰岩の地色 `--ground`。花ブロック（`hanaSVG`）がこのサイトの識別要素なので、
汎用のカード並べや棒グラフ一覧に置き換えない。SVGは外部ライブラリを使わず自前で描く方針を維持する。
