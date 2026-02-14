#!/usr/bin/env python3
"""Event Tools — local server with real-time command relay. No dependencies."""
import glob as globmod
import hashlib
import http.server
import json
import os
import queue
import shutil
import socket
import subprocess
import sys
import threading
import webbrowser

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8000
DIR = os.path.dirname(os.path.abspath(__file__))
PPTX_CACHE = os.path.join(DIR, "assets", ".pptx-cache")
_soffice_bin = shutil.which("soffice")
_convert_lock = threading.Lock()

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


def _pptx_fingerprint(path):
    st = os.stat(path)
    h = hashlib.md5(f"{st.st_mtime}:{st.st_size}:{path}".encode()).hexdigest()[:12]
    return h


def _convert_pptx(pptx_path):
    """Convert a .pptx to cached PNGs using LibreOffice. Returns list of image paths or []."""
    if not _soffice_bin:
        return []
    fp = _pptx_fingerprint(pptx_path)
    basename = os.path.splitext(os.path.basename(pptx_path))[0]
    cache_dir = os.path.join(PPTX_CACHE, f"{basename}_{fp}")
    manifest = os.path.join(cache_dir, "manifest.json")

    if os.path.isfile(manifest):
        with open(manifest) as f:
            return json.load(f)

    with _convert_lock:
        # Double-check after acquiring lock
        if os.path.isfile(manifest):
            with open(manifest) as f:
                return json.load(f)

        os.makedirs(cache_dir, exist_ok=True)
        try:
            subprocess.run(
                [_soffice_bin, "--headless", "--nologo", "--nofirststartwizard",
                 "--convert-to", "png", "--outdir", cache_dir, pptx_path],
                timeout=120, capture_output=True
            )
        except Exception as e:
            sys.stderr.write(f"  PPTX conversion failed: {e}\n")
            return []

        pngs = sorted(globmod.glob(os.path.join(cache_dir, "*.png")))
        if not pngs:
            sys.stderr.write(f"  PPTX conversion produced no images for {pptx_path}\n")
            return []

        # LibreOffice outputs a single file for single-slide; for multi-slide it may
        # produce multiple or a single file depending on version. Rename to ordered names.
        result = []
        for i, src in enumerate(pngs):
            dest = os.path.join(cache_dir, f"slide-{i + 1:03d}.png")
            if src != dest:
                os.rename(src, dest)
            result.append(dest)

        with open(manifest, "w") as f:
            json.dump(result, f)
        return result


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
                    # Skip the pptx cache directory
                    dirs[:] = [d for d in sorted(dirs) if d != ".pptx-cache"]
                    for f in sorted(filenames):
                        if f.startswith("."):
                            continue
                        full = os.path.join(root, f)
                        rel = os.path.relpath(full, DIR)
                        ext = os.path.splitext(f)[1].lower()

                        if ext in (".pptx",):
                            slide_images = _convert_pptx(full)
                            if slide_images:
                                pptx_name = os.path.splitext(f)[0]
                                for si, img_path in enumerate(slide_images):
                                    img_rel = os.path.relpath(img_path, DIR)
                                    files.append({
                                        "name": f"{pptx_name} — Slide {si + 1}",
                                        "path": "/" + img_rel.replace(os.sep, "/"),
                                        "type": "image",
                                        "source": f,
                                    })
                            else:
                                if not _soffice_bin:
                                    sys.stderr.write(f"  Skipping {f}: LibreOffice not found (install for PPTX support)\n")
                                files.append({"name": f, "path": "/" + rel.replace(os.sep, "/"), "type": "other"})
                            continue

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
        if "/api/events" not in args[0] and "/api/clients" not in args[0]:
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


def main():
    ips = get_local_ips()
    primary_ip = ips[0]
    base = f"http://{primary_ip}:{PORT}"

    server = ThreadedHTTPServer(("0.0.0.0", PORT), Handler)

    print()
    print("  ╔══════════════════════════════════════════════════════╗")
    print("  ║              Event Tools is running!                ║")
    print("  ╚══════════════════════════════════════════════════════╝")
    print()
    print(f"  Homepage:    http://localhost:{PORT}")
    print()

    if len(ips) == 1:
        print(f"  Network:     {base}")
    else:
        print("  Network IPs:")
        for ip in ips:
            print(f"    • http://{ip}:{PORT}")
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
    print()
    print("  Press Ctrl+C to stop.")
    print()

    webbrowser.open(f"http://localhost:{PORT}")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  Stopped.")
        server.server_close()


if __name__ == "__main__":
    main()
