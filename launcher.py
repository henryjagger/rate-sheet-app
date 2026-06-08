"""
Desktop app launcher.
Starts Flask on a free local port, then opens a native app window via pywebview.
No browser, no address bar — looks and feels like a real app.
"""
import os
import sys
import socket
import threading
import time


def resource_path(rel):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def data_dir():
    d = os.path.join(os.path.expanduser("~"), ".ratesheet")
    os.makedirs(d, exist_ok=True)
    return d


def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def start_flask(port):
    os.environ["RSG_TEMPLATE_FOLDER"] = resource_path("templates")
    os.environ["RSG_STATIC_FOLDER"]   = resource_path("static")
    os.environ["RSG_LOOKUP_PATH"]     = os.path.join(data_dir(), "institution_lookup.xlsx")
    os.environ["RSG_PASSWORDS_PATH"]  = os.path.join(data_dir(), "passwords.json")
    os.environ["RSG_STATS_PATH"]      = os.path.join(data_dir(), "stats.json")
    os.environ["RSG_BUNDLE_LOOKUP"]   = resource_path("institution_lookup.xlsx")

    from server import app
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    port = find_free_port()

    # Start Flask in background thread before creating the window
    t = threading.Thread(target=start_flask, args=(port,), daemon=True)
    t.start()
    time.sleep(1.5)

    import webview
    webview.create_window(
        title="Rate Sheet Generator",
        url=f"http://127.0.0.1:{port}",
        width=1280,
        height=820,
        min_size=(900, 600),
        text_select=True,
    )
    webview.start()
