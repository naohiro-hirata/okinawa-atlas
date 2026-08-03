#!/usr/bin/env python3
"""HTTP Rangeに対応した確認用サーバー。

python -m http.server はRangeリクエスト（206 Partial Content）に対応していないため、
public/tiles/okinawa.pmtiles を読めず地図が表示されない。pmtiles.js は
「Server returned no content-length header or content-length exceeding request.
Check that your storage backend supports HTTP Byte Serving.」で失敗する。
GitHub Pages はRangeに対応しているので、本番では問題にならない。

  python scripts/serve.py            # http://127.0.0.1:8000/
  python scripts/serve.py 8001       # ポート指定
"""
import functools
import http.server
import os
import re
import sys

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "public")


class RangeHandler(http.server.SimpleHTTPRequestHandler):
    """Rangeヘッダを解釈して206を返す。単一レンジのみ対応（pmtilesはそれで足りる）。"""

    def send_head(self):
        rng = self.headers.get("Range")
        if not rng:
            return super().send_head()
        m = re.fullmatch(r"bytes=(\d*)-(\d*)", rng.strip())
        if not m:
            return super().send_head()

        path = self.translate_path(self.path)
        if os.path.isdir(path):
            return super().send_head()
        try:
            f = open(path, "rb")
        except OSError:
            self.send_error(404, "File not found")
            return None

        size = os.fstat(f.fileno()).st_size
        start_s, end_s = m.group(1), m.group(2)
        if start_s == "":  # bytes=-N （末尾N“バイト”）
            if end_s == "":
                f.close()
                self.send_error(400, "Bad Range")
                return None
            length = min(int(end_s), size)
            start = size - length
            end = size - 1
        else:
            start = int(start_s)
            end = int(end_s) if end_s else size - 1
            end = min(end, size - 1)
        if start > end or start >= size:
            f.close()
            self.send_response(416)
            self.send_header("Content-Range", f"bytes */{size}")
            self.end_headers()
            return None

        f.seek(start)
        self.send_response(206)
        self.send_header("Content-Type", self.guess_type(path))
        self.send_header("Content-Range", f"bytes {start}-{end}/{size}")
        self.send_header("Content-Length", str(end - start + 1))
        self.send_header("Accept-Ranges", "bytes")
        self.end_headers()
        return _Slice(f, end - start + 1)

    def end_headers(self):
        # 確認中に古い内容を掴まないようにする
        self.send_header("Cache-Control", "no-store")
        super().end_headers()


class _Slice:
    """copyfile() に渡す、指定バイト数で尽きる読み出し口。"""

    def __init__(self, f, remaining):
        self.f, self.remaining = f, remaining

    def read(self, n=-1):
        if self.remaining <= 0:
            return b""
        if n is None or n < 0:
            n = self.remaining
        data = self.f.read(min(n, self.remaining))
        self.remaining -= len(data)
        return data

    def close(self):
        self.f.close()


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8000
    handler = functools.partial(RangeHandler, directory=os.path.abspath(ROOT))
    print(f"http://127.0.0.1:{port}/  （Range対応・{os.path.abspath(ROOT)} を配信）")
    http.server.ThreadingHTTPServer(("127.0.0.1", port), handler).serve_forever()
