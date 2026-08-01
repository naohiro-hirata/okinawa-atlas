#!/usr/bin/env bash
# 沖縄県の境界データから public/tiles/okinawa.pmtiles を作る。
#
# 元データ: e-Stat 統計地理情報システム、令和2年国勢調査
#           小地域（町丁・字等）（JGD2000）、沖縄県（r2ka47）。
#           https://www.e-stat.go.jp/gis/statmap-search?page=1&type=2&aggregateUnitForBoundary=A&toukeiCode=00200521&toukeiYear=2020&serveyId=A002005212020&datum=2000
# 市町村ポリゴンは、同じ字ポリゴンをCITYコードでdissolveして作る（別ソースを
# 使わないことで、字と市町村の境界の基準時点を完全に揃える）。
#
# 必要なツール: ogr2ogr（GDAL）、tippecanoe・tile-join（felt/tippecanoe）。
# どちらもWindowsには公式バイナリが無いため、このスクリプトは
# .github/workflows/build-tiles.yml（Ubuntu・workflow_dispatch）専用。
# ローカルで試す場合はWSL等のLinux環境で同じ手順を踏むこと。
#
#   scripts/build_tiles.sh
#
# 出力: public/tiles/okinawa.pmtiles（レイヤー: muni＝市町村、aza＝字）
#       ワークフロー側で自動コミットはしない。artifactを人が確認してから
#       手でコミットする運用（docs/handoff.md Phase 6 参照）。
set -euo pipefail
set -x

BOUNDARY_URL="https://www.e-stat.go.jp/gis/statmap-search/data?dlserveyId=A002005212020&code=47&coordSys=1&format=shape&downloadType=5&datum=2000"
WORK="$(mktemp -d)"
OUT="public/tiles/okinawa.pmtiles"
trap 'rm -rf "$WORK"' EXIT

mkdir -p public/tiles

curl -sL "$BOUNDARY_URL" -o "$WORK/r2ka47.zip"
(cd "$WORK" && unzip -o r2ka47.zip)

# 字ポリゴン（WGS84へ変換。JGD2000とはサブメートル差で実務上ほぼ同一）
ogr2ogr -f GeoJSON "$WORK/aza.geojson" "$WORK/r2ka47.shp" -t_srs EPSG:4326

# 市町村ポリゴン（PREF+CITYでdissolve。PREFは沖縄県なら常に"47"だが、
# app/template.html 側の市町村コード=PREF+CITY の組み立てと対応させるため残す）
ogr2ogr -f GeoJSON "$WORK/muni.geojson" "$WORK/r2ka47.shp" -t_srs EPSG:4326 \
  -dialect sqlite -sql "SELECT PREF, CITY, CITY_NAME, ST_Union(geometry) AS geometry FROM r2ka47 GROUP BY PREF, CITY, CITY_NAME"

# 字・市町村を別々にPMTiles化してから1本にまとめる（tile-joinはズーム範囲が
# 異なる複数タイルセットの結合を想定したツールなので、layerごとに個別の
# tippecanoe呼び出しにする方が挙動が読める）
tippecanoe -o "$WORK/aza.pmtiles" -l aza -Z0 -z14 --drop-densest-as-needed --force "$WORK/aza.geojson"
tippecanoe -o "$WORK/muni.pmtiles" -l muni -Z0 -z10 --drop-densest-as-needed --force "$WORK/muni.geojson"
tile-join -o "$OUT" --force "$WORK/aza.pmtiles" "$WORK/muni.pmtiles"

ls -la "$OUT"
echo "wrote $OUT"
