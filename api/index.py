"""Vercel 서버 함수: 접속할 때마다 이지비즈에서 공고를 가져와 웹페이지로 돌려준다.

- 그냥 접속하면 10분 동안은 같은 결과를 재사용해 빠르게 보여준다.
- '최신 정보 불러오기'(주소에 ?live=...)를 누르면 저장본 없이 지금 바로 다시 가져온다.
"""
import html
import os
import sys
from http.server import BaseHTTPRequestHandler
from urllib.parse import parse_qs, urlparse

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import gov_support_watch as g  # noqa: E402


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        live = "live" in parse_qs(urlparse(self.path).query)
        try:
            page, code = g.render_live(), 200
        except Exception as e:  # 이지비즈 접속 실패 등
            page = ('<!doctype html><meta charset="utf-8"><title>경기 지원사업 공고</title>'
                    '<p style="font-family:sans-serif;padding:24px">이지비즈에서 공고를 가져오지 못했습니다. '
                    f'잠시 뒤 새로고침해 주세요.<br><small>{html.escape(str(e))}</small></p>')
            code = 502
        body = page.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Cache-Control",
                         "no-store" if live or code != 200 else "public, s-maxage=600, stale-while-revalidate=3600")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)
