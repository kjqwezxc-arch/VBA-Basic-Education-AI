import http.server
import socketserver
import threading
import webbrowser
import time
import sys
from pathlib import Path

PORT = 8000
BASE_DIR = Path(__file__).parent


class QuietHandler(http.server.SimpleHTTPRequestHandler):
    """Serve files from BASE_DIR silently."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE_DIR), **kwargs)

    def log_message(self, format, *args):
        pass  # Suppress request logs


def find_free_port(start=8000):
    import socket
    for port in range(start, start + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("", port))
                return port
            except OSError:
                continue
    return start


if __name__ == "__main__":
    import socket as _socket

    port = find_free_port(PORT)

    # Get local network IP for sharing
    try:
        s = _socket.socket(_socket.AF_INET, _socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        local_ip = s.getsockname()[0]
        s.close()
    except Exception:
        local_ip = "127.0.0.1"

    local_url   = f"http://localhost:{port}/index.html"
    network_url = f"http://{local_ip}:{port}/index.html"

    # Bind on all interfaces so others on the same network can connect
    httpd = socketserver.TCPServer(("0.0.0.0", port), QuietHandler)
    httpd.allow_reuse_address = True

    def open_browser():
        time.sleep(0.8)
        webbrowser.open(local_url)

    threading.Thread(target=open_browser, daemon=True).start()

    print("=" * 56)
    print("  VBA Education Module — Local Server")
    print("=" * 56)
    print(f"  Local   : {local_url}")
    print(f"  Network : {network_url}  <-- share this")
    print()
    print("  Press Ctrl+C to stop the server.")
    print("=" * 56)

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
        httpd.server_close()
        sys.exit(0)
