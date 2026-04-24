#!/usr/bin/env python3
"""SVN Excel Diff — Desktop App

Wraps the Flask web server in a native window using pywebview.
"""

import sys
import threading

import webview

from server import app, FRONTEND_HTML  # noqa: F401

PORT = 9527


def start_server():
    """Run Flask in a background thread."""
    app.run(host="127.0.0.1", port=PORT, debug=False, use_reloader=False)


def main():
    # Start Flask server in background
    server_thread = threading.Thread(target=start_server, daemon=True)
    server_thread.start()

    # Create native window pointing to the Flask server
    webview.create_window(
        title="SVN Diff Viewer",
        url=f"http://127.0.0.1:{PORT}",
        width=1400,
        height=900,
        min_size=(900, 600),
        text_select=True,
    )
    webview.start()


if __name__ == "__main__":
    main()
