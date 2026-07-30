# 沖縄 地域データアトラス

沖縄県内の地域データを2階層で見るための静的サイト。GitHub Pages の無料枠だけで動きます。

- **Layer 1｜市町村** — 人口推移、年齢5歳階級別人口ピラミッド、移住定住施策と空き家利活用の取り組み状況
- **Layer 2｜字** — 字ごとの人口推移と年齢構成
- 市町村・字を横断する検索と、境界ポリゴンによるマップ選択

## いまの実装状況

| 領域 | 状態 |
|---|---|
| 移住定住施策・空き家（令和7年度調査 全41市町村） | **実データで動作** |
| 定住促進住宅 58物件 | **実データで動作** |
| 市町村の人口推移・人口ピラミッド | サンプル値（画面上で斜線＋「ダミー」表示） |
| 字別の人口推移・年齢構成 | サンプル値（同上） |
| マップ | 未実装（枠と接続手順のみ配置） |

サンプル値は画面上で必ず斜線とタグが付きます。実データを入れると自動的に消える設計です。

## 実装の進め方

GitHubが初めての場合は [docs/setup-github.md](docs/setup-github.md)（クリック操作のみで公開まで）。
Windows で手元の環境を整えるには [docs/setup-windows.md](docs/setup-windows.md)。
フェーズごとの手順と指示テンプレートは [docs/handoff.md](docs/handoff.md)。
公開前の確認は [docs/publish-checklist.md](docs/publish-checklist.md)（人が実施）。
プロジェクト共通のルールは [CLAUDE.md](CLAUDE.md) にあり、Claude Code が毎セッション自動で読み込みます。

## セットアップ

```bash
pip install openpyxl pandas xlrd

# 1) 調査Excelを正規化（原本は data/raw/survey/latest.xlsx）
python scripts/parse_survey.py data/raw/survey/latest.xlsx

# 2) 単一HTMLをビルド
python scripts/build_app.py --outfile public/index.html

# 3) 確認（public/index.html をブラウザで直接開いてもOK）
python -m http.server -d public 8000
```

## ディレクトリ

```
data/raw/survey/latest.xlsx   毎年の調査Excel（これを差し替えるだけで更新）
data/raw/aza/                 県の町字別人口Excel（fetch スクリプトが取得）
data/generated/               正規化済みJSON（Actions が生成、コミット不要）
data/aza_crosswalk.csv        字名の名寄せ辞書（手で育てる）
scripts/parse_survey.py       調査Excel → JSON
scripts/fetch_aza_population.py  県Excel取得 → 字別人口JSON
scripts/build_app.py          テンプレート＋JSON → public/index.html
app/template.html             UI本体（1ファイル・外部JSライブラリなし）
public/                       GitHub Pages の公開ディレクトリ
```

## データソース

| 用途 | 出典 | 更新 |
|---|---|---|
| 市町村別人口推移 | 国勢調査 1920–2020（e-Stat 時系列） | 5年 |
| 同・最新値 | 令和7年国勢調査 人口速報集計（2026-05-29公表、男女別人口と世帯総数のみ） | — |
| 市町村別人口ピラミッド | 令和2年国勢調査 人口等基本集計。R7確報は令和8年9月までに公表予定 | 5年 |
| **字別人口・世帯数** | 沖縄県「市町村の町字別住民基本台帳人口及び世帯数」平成23年〜令和7年 | 年次 |
| **字別 年齢×男女** | 令和2年国勢調査 小地域集計（町丁・字等別）。R7分は2027年以降の見込み | 5年 |
| 字の境界ポリゴン | e-Stat 統計地理情報システム 境界データ `r2ka47` | 5年 |
| 移住定住・空き家の取り組み | 自社ヒアリング調査（年次） | 年次 |

字別人口の取得:

```bash
python scripts/fetch_aza_population.py              # 全年
python scripts/fetch_aza_population.py --years 2025 # 最新年だけ
```

県が新年度分を公開したら `scripts/fetch_aza_population.py` の `FILES` にファイル名を1行足します。

### 字データの注意点

国勢調査の字境界・字名は調査区をもとに作られており、住居表示上の実際の町丁・字と一致しない場合があります。
住基データ（県Excel）と国勢調査小地域の突き合わせは自動では閉じないので、対応付けを
`data/aza_crosswalk.csv` に蓄積してください。`fetch_aza_population.py` は年次が欠けている字を
実行末尾に列挙するので、そこが手当ての入口になります。

## マップの追加

```bash
# 境界データ（Shape, 緯度経度）を e-Stat 統計地理情報システムから取得し r2ka47 を展開
ogr2ogr -f GeoJSON aza.geojson r2ka47/h27ka47.shp -t_srs EPSG:4326
tippecanoe -o public/tiles/okinawa.pmtiles -Z6 -z14 --coalesce-densest-as-needed \
  --layer=aza aza.geojson
```

`public/index.html` のマップ枠に MapLibre GL JS + pmtiles プロトコルを挿し、
ズームレベルで市町村ポリゴンと字ポリゴンを切り替えます。沖縄県分のみなら PMTiles は
十数MB程度で、Pages の上限（1GB）に対して余裕があります。

## 公開

1. リポジトリを **public** で作成（GitHub Pages の無料枠は public リポジトリのみ）
2. Settings → Pages → Source を **GitHub Actions** に設定
3. `main` に push すると `.github/workflows/pages.yml` がビルドして公開

Pages は静的ホスティングなのでアクセス制限はかけられません。今回のデータは個人情報を
含まないため問題ありませんが、将来的に物件所有者などを扱う場合は、そのデータだけ
別リポジトリに分け、Cloudflare Pages + Cloudflare Access（無料枠あり）などの
認証付きホスティングに載せてください。

## データ品質

`parse_survey.py` は毎回 `data/generated/quality_report.json` を書き出します。

- `unconfirmed_policy` / `unconfirmed_akiya` — 「最終確認状況」が FALSE の市町村。
  令和7年度時点では 糸満市・豊見城市・宜野座村・北谷町・北中城村・西原町 の6件で、
  過年度の回答が残っている可能性があります。画面上も「未確認」と表示されます。
- `needs_manual_review` — 備考欄の自然文から件数を機械的に取り切れなかった箇所。
  推測で数値を作らず、人が読む対象として残します。
- `code_join_failures` — 市町村コードの紐付けに失敗した行（現状 0件）。

備考欄からは推定空家数・空き家バンク累計登録数・成約件数を拾っています。
恒久的には、調査票に数値専用の列を設けるのが確実です。現状は 16市町村ぶんしか
数値が取れておらず、指標としての比較可能性が限られます。
