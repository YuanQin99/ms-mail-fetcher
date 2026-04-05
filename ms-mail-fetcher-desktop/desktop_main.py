import ctypes
import inspect
import logging
import os
import socket
import sys
import threading
import time
from pathlib import Path

import requests
import uvicorn
import webview
from webview.platforms.edgechromium import EdgeChrome

DESKTOP_ROOT = Path(__file__).resolve().parent
SERVER_ROOT = DESKTOP_ROOT.parent / 'ms-mail-fetcher-server'
if str(SERVER_ROOT) not in sys.path:
    sys.path.insert(0, str(SERVER_ROOT))

os.environ.setdefault('MS_MAIL_FETCHER_CONFIG', str(DESKTOP_ROOT / 'server.config.json'))
os.environ.setdefault('MS_MAIL_FETCHER_FRONTEND_DIR', str(DESKTOP_ROOT / 'template'))

from app.runtime import create_app, load_runtime_config


logger = logging.getLogger("ms_mail_fetcher.desktop")
if not logger.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )


DEFAULT_DESKTOP_HOST = "127.0.0.1"
DEFAULT_WAIT_SECONDS = 20
WINDOW_TITLE = "MS Mail Fetcher"
SINGLE_INSTANCE_MUTEX_NAME = "Local\\MS_MAIL_FETCHER_DESKTOP_SINGLE_INSTANCE"
MIN_WINDOW_WIDTH = 1100
MIN_WINDOW_HEIGHT = 760
DEFAULT_WINDOW_WIDTH = 1280
DEFAULT_WINDOW_HEIGHT = 860

_original_on_webview_ready = EdgeChrome.on_webview_ready


def _patched_on_webview_ready(self, sender, args):
    _original_on_webview_ready(self, sender, args)

    if not args.IsSuccess:
        return

    settings = sender.CoreWebView2.Settings
    settings.AreBrowserAcceleratorKeysEnabled = False
    settings.AreDefaultContextMenusEnabled = True
    settings.AreDevToolsEnabled = False
    settings.IsStatusBarEnabled = False


EdgeChrome.on_webview_ready = _patched_on_webview_ready

class _SingleInstanceGuard:
    def __init__(self, name: str):
        self._name = name
        self._handle = None

    def acquire(self) -> bool:
        kernel32 = ctypes.windll.kernel32
        self._handle = kernel32.CreateMutexW(None, False, self._name)
        if not self._handle:
            raise RuntimeError("Failed to create application mutex.")

        ERROR_ALREADY_EXISTS = 183
        last_error = kernel32.GetLastError()
        return last_error != ERROR_ALREADY_EXISTS

    def release(self) -> None:
        if self._handle:
            ctypes.windll.kernel32.CloseHandle(self._handle)
            self._handle = None


def _is_port_available(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            sock.bind((host, port))
        except OSError:
            return False
    return True


def _resolve_desktop_bind() -> tuple[str, int]:
    config, _ = load_runtime_config()
    preferred_port = int(config.get("port", 18765))
    retry_count = max(0, int(config.get("port_retry_count", 20)))

    for candidate in range(preferred_port, preferred_port + retry_count + 1):
        if _is_port_available(DEFAULT_DESKTOP_HOST, candidate):
            return DEFAULT_DESKTOP_HOST, candidate

    raise RuntimeError(
        f"No available desktop server port from {preferred_port} to {preferred_port + retry_count}."
    )


def _wait_until_ready(url: str, timeout_seconds: int = DEFAULT_WAIT_SECONDS) -> None:
    deadline = time.time() + timeout_seconds
    health_url = f"{url}/api/health"

    while time.time() < deadline:
        try:
            response = requests.get(health_url, timeout=1.5)
            if response.ok:
                return
        except requests.RequestException:
            pass
        time.sleep(0.3)

    raise RuntimeError("Backend did not become ready in time.")


def _run_server(server: uvicorn.Server) -> None:
    server.run()


def _prepare_webview_storage() -> Path:
    """Ensure WebView2 uses a stable user data directory so localStorage persists."""
    base = Path(os.getenv("LOCALAPPDATA") or Path.home()) / "ms-mail-fetcher" / "webview2"
    base.mkdir(parents=True, exist_ok=True)
    os.environ.setdefault("WEBVIEW2_USER_DATA_FOLDER", str(base))
    return base


def _sanitize_window_size(width: object, height: object) -> tuple[int, int]:
    try:
        width_val = int(width)
    except (TypeError, ValueError):
        width_val = DEFAULT_WINDOW_WIDTH

    try:
        height_val = int(height)
    except (TypeError, ValueError):
        height_val = DEFAULT_WINDOW_HEIGHT

    return max(MIN_WINDOW_WIDTH, width_val), max(MIN_WINDOW_HEIGHT, height_val)


def _load_window_size(api_url: str) -> tuple[int, int]:
    try:
        response = requests.get(f"{api_url}/api/ui/preferences", timeout=1.5)
        if not response.ok:
            return DEFAULT_WINDOW_WIDTH, DEFAULT_WINDOW_HEIGHT

        data = response.json() if response.content else {}
        if not isinstance(data, dict):
            return DEFAULT_WINDOW_WIDTH, DEFAULT_WINDOW_HEIGHT

        return _sanitize_window_size(data.get("window_width"), data.get("window_height"))
    except Exception:
        return DEFAULT_WINDOW_WIDTH, DEFAULT_WINDOW_HEIGHT


def _save_window_size(api_url: str, width: object, height: object) -> None:
    safe_width, safe_height = _sanitize_window_size(width, height)
    try:
        requests.put(
            f"{api_url}/api/ui/preferences",
            json={
                "window_width": safe_width,
                "window_height": safe_height,
            },
            timeout=1.5,
        )
    except Exception:
        logger.debug("Failed to persist desktop window size.", exc_info=True)


def _bind_window_size_persistence(
    window,
    api_url: str,
    initial_width: int,
    initial_height: int,
) -> None:
    events = getattr(window, "events", None)
    if events is None:
        return

    latest = {
        "width": initial_width,
        "height": initial_height,
    }

    def _capture_size_from_args(*args):
        if len(args) < 2:
            return
        width, height = _sanitize_window_size(args[0], args[1])
        latest["width"] = width
        latest["height"] = height

    def _persist_from_latest():
        width = latest.get("width")
        height = latest.get("height")
        _save_window_size(api_url, width, height)

    def _on_resized(*args):
        _capture_size_from_args(*args)

    def _on_closing(*args):
        _capture_size_from_args(*args)
        _persist_from_latest()

    if hasattr(events, "resized"):
        events.resized += _on_resized

    if hasattr(events, "closing"):
        events.closing += _on_closing
    elif hasattr(events, "closed"):
        events.closed += _on_closing


def main() -> None:
    guard = _SingleInstanceGuard(SINGLE_INSTANCE_MUTEX_NAME)
    if not guard.acquire():
        logger.info("Another app instance is already running. Exit.")
        return
    server = None
    server_thread = None
    main_window = None
    try:
        host, port = _resolve_desktop_bind()
        api_url = f"http://{host}:{port}"
        storage_path = _prepare_webview_storage()

        app = create_app()
        app.state.server_host = host
        app.state.server_port = port

        config = uvicorn.Config(
            app,
            host=host,
            port=port,
            reload=False,
            access_log=False,
        )
        server = uvicorn.Server(config)

        server_thread = threading.Thread(
            target=_run_server,
            args=(server,),
            daemon=True,
            name="ms-mail-fetcher-api",
        )
        server_thread.start()

        _wait_until_ready(api_url)
        logger.info("Desktop UI opening: %s", api_url)

        initial_width, initial_height = _load_window_size(api_url)

        create_window_params = inspect.signature(webview.create_window).parameters
        window_kwargs = {
            "title": WINDOW_TITLE,
            "url": api_url,
            "min_size": (MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT),
            "width": initial_width,
            "height": initial_height,
        }

        if "private_mode" in create_window_params:
            window_kwargs["private_mode"] = False
        if "storage_path" in create_window_params:
            window_kwargs["storage_path"] = str(storage_path)
        if "text_select" in create_window_params:
            window_kwargs["text_select"] = True
        main_window = webview.create_window(**window_kwargs)
        _bind_window_size_persistence(main_window, api_url, initial_width, initial_height)
        webview.start()
    finally:
        if server is not None:
            logger.info("Desktop window closed, stopping backend...")
            server.should_exit = True
        if server_thread is not None:
            server_thread.join(timeout=8)
        guard.release()


if __name__ == "__main__":
    main()

