#!/usr/bin/env python3
"""Event Tools — local server with real-time command relay. No dependencies."""
import http.server
import json
import os
import queue
import socket
import ssl
import subprocess
import sys
import threading
import webbrowser

DEBUG = "--debug" in sys.argv
_args = [a for a in sys.argv[1:] if a != "--debug"]
PORT = int(_args[0]) if _args else 8000
DIR = os.path.dirname(os.path.abspath(__file__))
CERT_FILE = os.path.join(DIR, ".event-tools-cert.pem")
KEY_FILE = os.path.join(DIR, ".event-tools-key.pem")
clients = []
clients_lock = threading.Lock()


def get_local_ips():
    ips = []
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ip = info[4][0]
            if not ip.startswith("127.") and ip not in ips:
                ips.append(ip)
    except Exception:
        pass
    if not ips:
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ips.append(s.getsockname()[0])
            s.close()
        except Exception:
            pass
    return ips or ["127.0.0.1"]


def broadcast(data):
    msg = f"data: {json.dumps(data)}\n\n"
    with clients_lock:
        dead = []
        for q in clients:
            try:
                q.put_nowait(msg)
            except Exception:
                dead.append(q)
        for q in dead:
            clients.remove(q)


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIR, **kwargs)

    def do_POST(self):
        if self.path == "/api/command":
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length)
            try:
                data = json.loads(body)
            except Exception:
                self.send_response(400)
                self.end_headers()
                return
            broadcast(data)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(b'{"ok":true}')
        else:
            self.send_response(404)
            self.end_headers()

    def do_GET(self):
        if self.path == "/api/events":
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream")
            self.send_header("Cache-Control", "no-cache")
            self.send_header("Connection", "keep-alive")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()

            q = queue.Queue()
            with clients_lock:
                clients.append(q)

            try:
                self.wfile.write(b"data: {\"type\":\"connected\"}\n\n")
                self.wfile.flush()
                while True:
                    try:
                        msg = q.get(timeout=15)
                        self.wfile.write(msg.encode())
                        self.wfile.flush()
                    except queue.Empty:
                        self.wfile.write(b": keepalive\n\n")
                        self.wfile.flush()
            except (BrokenPipeError, ConnectionResetError, OSError):
                pass
            finally:
                with clients_lock:
                    if q in clients:
                        clients.remove(q)
        elif self.path == "/api/files":
            assets_dir = os.path.join(DIR, "assets")
            files = []
            if os.path.isdir(assets_dir):
                for root, dirs, filenames in os.walk(assets_dir):
                    dirs.sort()
                    for f in sorted(filenames):
                        if f.startswith("."):
                            continue
                        full = os.path.join(root, f)
                        rel = os.path.relpath(full, DIR)
                        ext = os.path.splitext(f)[1].lower()
                        kind = "other"
                        if ext in (".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".bmp"):
                            kind = "image"
                        elif ext in (".mp4", ".webm", ".ogg", ".mov"):
                            kind = "video"
                        elif ext in (".pdf",):
                            kind = "pdf"
                        files.append({"name": f, "path": "/" + rel.replace(os.sep, "/"), "type": kind})
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps(files).encode())
        elif self.path == "/api/clients":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            with clients_lock:
                count = len(clients)
            self.wfile.write(json.dumps({"count": count}).encode())
        else:
            super().do_GET()

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def log_message(self, format, *args):
        if "/api/events" in args[0] or "/api/clients" in args[0]:
            return
        if not DEBUG and "/api/" in args[0]:
            return
        sys.stderr.write(f"  {args[0]}\n")


class ThreadedHTTPServer(http.server.HTTPServer):
    allow_reuse_address = True
    daemon_threads = True

    def process_request(self, request, client_address):
        t = threading.Thread(target=self.process_request_thread,
                             args=(request, client_address))
        t.daemon = True
        t.start()

    def process_request_thread(self, request, client_address):
        try:
            self.finish_request(request, client_address)
        except Exception:
            pass
        try:
            self.shutdown_request(request)
        except Exception:
            pass


def ensure_cert():
    """Generate a self-signed cert for HTTPS (needed for camera access on phones)."""
    if os.path.exists(CERT_FILE) and os.path.exists(KEY_FILE):
        return True
    try:
        subprocess.run(
            [
                "openssl", "req", "-x509", "-newkey", "rsa:2048",
                "-keyout", KEY_FILE, "-out", CERT_FILE,
                "-days", "365", "-nodes",
                "-subj", "/CN=Event Tools",
            ],
            check=True, capture_output=True,
        )
        return True
    except (FileNotFoundError, subprocess.CalledProcessError):
        return False


def main():
    ips = get_local_ips()
    primary_ip = ips[0]

    use_https = ensure_cert()
    scheme = "https" if use_https else "http"
    base = f"{scheme}://{primary_ip}:{PORT}"

    server = ThreadedHTTPServer(("0.0.0.0", PORT), Handler)

    if use_https:
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        ctx.load_cert_chain(CERT_FILE, KEY_FILE)
        server.socket = ctx.wrap_socket(server.socket, server_side=True)

    print()
    print("  ╔══════════════════════════════════════════════════════╗")
    print("  ║              Event Tools is running!                ║")
    print("  ╚══════════════════════════════════════════════════════╝")
    print()
    print(f"  Homepage:    {scheme}://localhost:{PORT}")
    print()

    if len(ips) == 1:
        print(f"  Network:     {base}")
    else:
        print("  Network IPs:")
        for ip in ips:
            print(f"    • {scheme}://{ip}:{PORT}")
    print()

    print("  ┌──────────────────────────────────────────────────────┐")
    print("  │  DISPLAY  (open on the big screen / projector)      │")
    print(f"  │  {base}/display.html")
    print("  │                                                      │")
    print("  │  CONTROLLER  (open on your phone)                   │")
    print(f"  │  {base}/controller.html")
    print("  └──────────────────────────────────────────────────────┘")
    print()
    print("  Both devices must be on the same Wi-Fi network.")
    if use_https:
        print()
        print("  ⚠  Using a self-signed certificate for camera access.")
        print("     Your browser will show a security warning — click")
        print('     "Advanced" → "Proceed" to accept it on each device.')
    print()
    print("  Press Ctrl+C to stop.")
    print()

    webbrowser.open(f"{scheme}://localhost:{PORT}")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  Stopped.")
        server.server_close()


if __name__ == "__main__":
    main()
