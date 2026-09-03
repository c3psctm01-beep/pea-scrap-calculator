import sys
import os

# Ensure the root project directory is in python path
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from server import PEAAppHandler

class handler(PEAAppHandler):
    """
    Vercel Serverless Function Handler.
    Inherits all API routing, PDF analysis, Excel parsing, and CORS handling from PEAAppHandler.
    """
    pass

# For local standalone testing
if __name__ == "__main__":
    import http.server
    server_address = ("", 8080)
    httpd = http.server.ThreadingHTTPServer(server_address, handler)
    print("Test Vercel handler running on port 8080...")
    httpd.serve_forever()
