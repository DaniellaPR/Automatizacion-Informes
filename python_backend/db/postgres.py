import threading
from typing import Optional
from contextlib import contextmanager
from psycopg2.pool import SimpleConnectionPool
from ..config import PG_CONFIG

class _SingletonMeta(type):
    _instances = {}
    _lock: threading.Lock = threading.Lock()

    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            with cls._lock:
                if cls not in cls._instances:
                    inst = super().__call__(*args, **kwargs)
                    cls._instances[cls] = inst
        return cls._instances[cls]

class PostgresPool(metaclass=_SingletonMeta):
    """Singleton que expone un pool de conexiones a PostgreSQL."""
    def __init__(self):
        self._pool: Optional[SimpleConnectionPool] = None  # <- nombre único

    def _init_pool(self):
        if self._pool is None:
            self._pool = SimpleConnectionPool(
                int(PG_CONFIG["minconn"]),
                int(PG_CONFIG["maxconn"]),
                host=PG_CONFIG["host"],
                port=int(PG_CONFIG["port"]),
                dbname=PG_CONFIG["dbname"],
                user=PG_CONFIG["user"],
                password=PG_CONFIG["password"],
                connect_timeout=5,
            )

    def get_conn(self):
        if self._pool is None:
            self._init_pool()
        return self._pool.getconn()

    def put_conn(self, conn):
        if self._pool:
            self._pool.putconn(conn)

    def closeall(self):
        if self._pool:
            self._pool.closeall()
            self._pool = None

@contextmanager
def get_cursor():
    pool = PostgresPool()
    conn = pool.get_conn()
    try:
        with conn.cursor() as cur:
            yield cur
            conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        pool.put_conn(conn)
