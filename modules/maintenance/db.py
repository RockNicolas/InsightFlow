import os
from contextlib import contextmanager
from datetime import datetime

import psycopg2
from dotenv import load_dotenv
from psycopg2.extras import RealDictCursor

load_dotenv()


def get_database_url():
    url = os.getenv("DATABASE_URL", "").strip()
    if not url:
        raise RuntimeError("DATABASE_URL não configurada no arquivo .env")
    return url


def disconnect_db():
    """Compatível com atexit do web_app (conexões são por request)."""
    pass


@contextmanager
def db_cursor():
    conn = psycopg2.connect(get_database_url())
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            yield conn, cur
            conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def parse_display_date(value: str) -> datetime | None:
    raw = str(value or "").strip()
    if not raw:
        return None

    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def format_display_date(value) -> str | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    return str(value)
