# -*- coding: utf-8 -*-
"""网址监控测试页服务（临时）：每请求必变（秒+随机数），验证监控→变化检测→AI 分析全链路"""
import http.server
import datetime
import random

HTML = """<!doctype html><html lang="zh"><head><meta charset="utf-8"><title>监控测试页</title>
<style>body{{font-family:system-ui;max-width:640px;margin:60px auto;padding:0 20px;line-height:1.7}}
.price{{font-size:42px;color:#e4393c;font-weight:700}}
.box{{background:#f5f7fa;border-radius:10px;padding:16px 20px;margin:12px 0}}</style></head><body>
<h1>商品价格监控演示页</h1>
<div class="box">当前价格：<span class="price">¥{price}</span></div>
<div class="box">服务器时间：<b id="t">{now}</b><br>随机波动值：<b>{rand}</b></div>
<p>本页面每次刷新都会变化（时间与随机数），用于验证「网址监控 → 变化检测 → AI 分析」链路。</p>
</body></html>"""


class H(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        price = round(100 + random.random() * 50, 2)
        body = HTML.format(price=price, now=now, rand=random.randint(1000, 9999)).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *a):
        pass


if __name__ == "__main__":
    server = http.server.ThreadingHTTPServer(("127.0.0.1", 8123), H)
    print("monitor-test page: http://127.0.0.1:8123/")
    server.serve_forever()
