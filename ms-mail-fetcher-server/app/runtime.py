import json
import logging
import os
import socket
import sys
import time
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI
from fastapi import HTTPException
from fastapi import Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from app.api.routes.account_types import router as account_types_router
from app.api.routes.accounts import router as accounts_router
from app.api.routes.health import router as health_router
from app.api.routes.mail import router as mail_router
from app.api.routes.ui_preferences import router as ui_preferences_router
from app.crud.account_types import ensure_default_account_types
from app.crud.accounts import migrate_inactive_accounts_to_archive_table
from app.db.database import Base, SessionLocal, engine


logger = logging.getLogger('ms_mail_fetcher')
if not logger.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s | %(levelname)s | %(message)s',
    )


CONFIG_FILE_NAME = 'server.config.json'
FRONTEND_TEMPLATE_DIR = 'template'
ENV_CONFIG_PATH = 'MS_MAIL_FETCHER_CONFIG'
ENV_FRONTEND_DIR = 'MS_MAIL_FETCHER_FRONTEND_DIR'
DEFAULT_HOST = '0.0.0.0'
DEFAULT_PORT = 18765
DEFAULT_RELOAD = False
DEFAULT_AUTO_PORT_FALLBACK = True
DEFAULT_PORT_RETRY_COUNT = 20


def _parse_bool(value, default: bool) -> bool:
    if value is None:
        return default

    if isinstance(value, bool):
        return value

    text = str(value).strip().lower()
    if text in {'1', 'true', 'yes', 'on'}:
        return True
    if text in {'0', 'false', 'no', 'off'}:
        return False

    return default


def _parse_int(value, default: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _is_port_available(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            sock.bind((host, port))
        except OSError:
            return False
    return True


def load_runtime_config() -> tuple[dict, Path]:
    env_config_path = os.getenv(ENV_CONFIG_PATH)
    if env_config_path:
        config_path = Path(env_config_path).expanduser().resolve()
    elif getattr(sys, 'frozen', False):
        config_path = Path(sys.executable).resolve().parent / CONFIG_FILE_NAME
    else:
        config_path = Path(__file__).resolve().parent.parent / CONFIG_FILE_NAME

    config = {
        'host': DEFAULT_HOST,
        'port': DEFAULT_PORT,
        'reload': DEFAULT_RELOAD,
        'auto_port_fallback': DEFAULT_AUTO_PORT_FALLBACK,
        'port_retry_count': DEFAULT_PORT_RETRY_COUNT,
    }

    if not config_path.exists():
        return config, config_path

    try:
        with config_path.open('r', encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, dict):
            config.update(data)
    except Exception as exc:
        logger.warning(f'Failed to read {config_path.name}, using defaults. reason={exc}')

    return config, config_path


def resolve_server_bind() -> tuple[str, int, bool, Path]:
    config, config_path = load_runtime_config()

    host = str(os.getenv('HOST') or config.get('host') or DEFAULT_HOST)
    preferred_port = _parse_int(os.getenv('PORT'), _parse_int(config.get('port'), DEFAULT_PORT))
    reload_enabled = _parse_bool(os.getenv('RELOAD'), _parse_bool(config.get('reload'), DEFAULT_RELOAD))

    auto_port_fallback = _parse_bool(
        config.get('auto_port_fallback'),
        DEFAULT_AUTO_PORT_FALLBACK,
    )
    retry_count = max(0, _parse_int(config.get('port_retry_count'), DEFAULT_PORT_RETRY_COUNT))

    if auto_port_fallback:
        for candidate in range(preferred_port, preferred_port + retry_count + 1):
            if _is_port_available(host, candidate):
                return host, candidate, reload_enabled, config_path

        raise RuntimeError(
            f'No available port found from {preferred_port} to {preferred_port + retry_count}.'
        )

    if not _is_port_available(host, preferred_port):
        raise RuntimeError(
            f'Port {preferred_port} is occupied and auto_port_fallback is disabled.'
        )

    return host, preferred_port, reload_enabled, config_path


def resolve_frontend_dist() -> Path | None:
    project_root = Path(__file__).resolve().parent.parent
    workspace_root = project_root.parent

    candidates = []

    env_frontend_dir = os.getenv(ENV_FRONTEND_DIR)
    if env_frontend_dir:
        candidates.append(Path(env_frontend_dir).expanduser())

    candidates.extend(
        [
            project_root / FRONTEND_TEMPLATE_DIR,
            workspace_root / 'ms-mail-fetcher-desktop' / FRONTEND_TEMPLATE_DIR,
        ]
    )

    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        candidates.append(Path(meipass) / FRONTEND_TEMPLATE_DIR)

    if getattr(sys, 'frozen', False):
        candidates.append(Path(sys.executable).resolve().parent / FRONTEND_TEMPLATE_DIR)

    for path in candidates:
        if (path / 'index.html').exists():
            return path.resolve()

    return None


def create_app() -> FastAPI:
    @asynccontextmanager
    async def lifespan(app_: FastAPI):
        Base.metadata.create_all(bind=engine)
        db = SessionLocal()
        try:
            ensure_default_account_types(db)
            moved_archived = migrate_inactive_accounts_to_archive_table(db)
            if moved_archived:
                logger.info('Migrated %s archived accounts into archived_accounts table.', moved_archived)
        finally:
            db.close()

        host = str(getattr(app_.state, 'server_host', '')) or str(os.getenv('HOST') or DEFAULT_HOST)
        port = str(getattr(app_.state, 'server_port', '')) or str(os.getenv('PORT') or DEFAULT_PORT)
        logger.info(f'Server started: http://{host}:{port}')
        yield
        logger.info('Server stopped')

    app = FastAPI(title='MS-Mail GPT Manager API', version='2.0.0', lifespan=lifespan)

    app.add_middleware(
        CORSMiddleware,
        allow_origins=['*'],
        allow_credentials=True,
        allow_methods=['*'],
        allow_headers=['*'],
    )

    @app.middleware('http')
    async def log_requests(request: Request, call_next):
        start = time.perf_counter()
        response = await call_next(request)
        duration_ms = (time.perf_counter() - start) * 1000
        client_ip = request.client.host if request.client else 'unknown'
        logger.info(
            f'{client_ip} | {request.method} {request.url.path} | {response.status_code} | {duration_ms:.2f}ms'
        )
        return response

    app.include_router(accounts_router)
    app.include_router(account_types_router)
    app.include_router(mail_router)
    app.include_router(health_router)
    app.include_router(ui_preferences_router)

    frontend_dist = resolve_frontend_dist()
    if frontend_dist:
        logger.info(f'Frontend dist loaded: {frontend_dist}')

        @app.get('/', include_in_schema=False)
        async def serve_frontend_index():
            return FileResponse(frontend_dist / 'index.html')

        @app.get('/{full_path:path}', include_in_schema=False)
        async def serve_frontend(full_path: str):
            if full_path.startswith('api/'):
                raise HTTPException(status_code=404, detail='Not Found')

            target = (frontend_dist / full_path).resolve()
            if target.is_file() and target.is_relative_to(frontend_dist.resolve()):
                return FileResponse(target)

            return FileResponse(frontend_dist / 'index.html')
    else:
        logger.warning('Frontend dist not found. API mode only.')

    return app
