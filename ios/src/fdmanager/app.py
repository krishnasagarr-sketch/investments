"""Toga app shell: starts the real Flask app (flask_backend, a symlink to
the repo-root app.py) on a background thread, then shows it in a native
WebView — the same idea as the desktop .exe/.app and the Android app, with
Toga's WebView standing in for "open the system browser"."""
import os
import socket
import threading
import time
from pathlib import Path

import toga
from toga.style.pack import Pack

SERVER_URL = "http://127.0.0.1:5000"


class FDManagerApp(toga.App):
    def startup(self):
        # iOS apps are sandboxed — there's no "next to the binary" or
        # "~/Library/Application Support" the way desktop platforms have.
        # Documents is the standard, backed-up, writable place for an app's
        # own user data on iOS.
        data_dir = Path.home() / "Documents" / "FDManagerData"
        data_dir.mkdir(parents=True, exist_ok=True)
        os.environ["FDMANAGER_DATA_DIR"] = str(data_dir)

        # Imported only now — flask_backend reads FDMANAGER_DATA_DIR at
        # import time to decide where the database/session key/cache live.
        from fdmanager import flask_backend

        threading.Thread(target=flask_backend.main, daemon=True).start()
        self._wait_for_server()

        self.webview = toga.WebView(style=Pack(flex=1), url=SERVER_URL)
        self.main_window = toga.MainWindow(title=self.formal_name)
        self.main_window.content = self.webview
        self.main_window.show()

    def _wait_for_server(self, timeout=5.0):
        """Block briefly until Flask is actually listening, so the WebView
        isn't constructed with a URL that immediately fails to load."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                with socket.create_connection(("127.0.0.1", 5000), timeout=0.2):
                    return
            except OSError:
                time.sleep(0.15)


def main():
    return FDManagerApp(formal_name="FD Manager", app_id="com.fdmanager.app")
