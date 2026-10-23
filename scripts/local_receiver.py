# 本机数据接收器：浏览器页面 POST JSON 过来，存到 数据/raw/。只监听 127.0.0.1
import json, os, re, sys
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

OUT = Path(os.environ.get('JOBHUNT_HOME', './workspace')) / 'data' / 'raw'
OUT.mkdir(parents=True, exist_ok=True)

class H(BaseHTTPRequestHandler):
    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Private-Network", "true")

    def do_OPTIONS(self):
        self.send_response(204); self._cors(); self.end_headers()

    def do_GET(self):
        # 只下发 _build/serve 里的文件（网址清单），不暴露其它目录
        name = Path(self.path.strip("/")).name
        f = Path(os.environ.get('JOBHUNT_HOME', './workspace')) / 'serve' / name
        if not f.is_file():
            self.send_response(404); self._cors(); self.end_headers(); return
        b = f.read_bytes()
        self.send_response(200); self._cors(); self.send_header("Content-Type", "application/json; charset=utf-8"); self.end_headers(); self.wfile.write(b)

    def do_POST(self):
        name = re.sub(r'[\\/:*?"<>|]', "_", self.path.strip("/")) or "data.json"
        body = self.rfile.read(int(self.headers.get("Content-Length", 0)))
        json.loads(body)  # 校验是合法 JSON
        (OUT / name).write_bytes(body)
        print("saved", name, len(body), flush=True)
        self.send_response(200); self._cors(); self.end_headers(); self.wfile.write(b"ok")

HTTPServer(("127.0.0.1", 8765), H).serve_forever()
