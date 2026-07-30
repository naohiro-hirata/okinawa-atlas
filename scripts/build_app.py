#!/usr/bin/env python3
"""app/template.html に data/generated/*.json を埋め込み、単一HTMLを出力する。

GitHub Pages はそのまま置くだけで動く。fetch を使わないので file:// でも開ける。
使い方: python scripts/build_app.py [--outfile public/index.html]
"""
import argparse, json
from pathlib import Path

ap = argparse.ArgumentParser()
ap.add_argument("--template", default="app/template.html")
ap.add_argument("--datadir", default="data/generated")
ap.add_argument("--outfile", default="public/index.html")
a = ap.parse_args()

d = Path(a.datadir)
payload = {k: json.loads((d / f"{k}.json").read_text(encoding="utf-8"))
           for k in ("municipalities", "policy", "housing", "akiya")}

html = Path(a.template).read_text(encoding="utf-8")
token = "/*__DATA__*/"
if token not in html:
    raise SystemExit(f"テンプレートに {token} がありません")
html = html.replace(token, json.dumps(payload, ensure_ascii=False, separators=(",", ":")))

out = Path(a.outfile); out.parent.mkdir(parents=True, exist_ok=True)
out.write_text(html, encoding="utf-8")
print(f"wrote {out}  ({out.stat().st_size/1024:.0f} KB)")
