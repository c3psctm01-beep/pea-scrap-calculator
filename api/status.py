from http.server import BaseHTTPRequestHandler
import json
import os
import sys

API_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(API_DIR, ".."))
if API_DIR not in sys.path:
    sys.path.insert(0, API_DIR)

try:
    from index import get_available_desktop_files
except Exception:
    def get_available_desktop_files():
        return []

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        available = get_available_desktop_files()
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.end_headers()
        self.wfile.write(json.dumps({
            "status": "ok",
            "default_pdf_exists": len(available) > 0,
            "available_files": available,
            "is_vercel": True
        }, ensure_ascii=False).encode('utf-8'))

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
