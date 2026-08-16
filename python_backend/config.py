from dotenv import load_dotenv
import os

load_dotenv()

PG_CONFIG = {
    "host": os.getenv("PG_HOST", "localhost"),
    "port": int(os.getenv("PG_PORT", "5432")),
    "dbname": os.getenv("PG_DB","prueba_DB"),
    "user":os.getenv("PG_USER", "postgres"),
    "password":os.getenv("PG_PASSWORD", "Mahp2005"),
    "minconn":int(os.getenv("PG_MINCONN","1")),
    "maxconn":int(os.getenv("PG_MAXCONN", "5")),
}