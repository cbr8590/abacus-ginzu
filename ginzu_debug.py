"""
Debug logging for Ginzu app. Set a handler to redirect logs to the UI debug panel.
Thread-safe: handler is called from worker threads; UI updates via root.after().
"""

import threading
from datetime import datetime

_handler = None
_lock = threading.Lock()


def set_handler(handler):
    """Set the debug log handler. Call with None to clear."""
    global _handler
    with _lock:
        _handler = handler


def log(msg: str, level: str = "INFO"):
    """Log a message. If a handler is set, it receives (msg, level)."""
    ts = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    line = f"[{ts}] [{level}] {msg}"
    with _lock:
        h = _handler
    if h:
        try:
            h(line)
        except Exception:
            pass
    else:
        print(line)
