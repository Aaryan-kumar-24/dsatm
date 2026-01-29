import os
import psycopg2

def get_db_connection():
    db_url = os.environ.get("DATABASE_URL")

    if not db_url:
        raise RuntimeError(
            "DATABASE_URL not set. This app is designed to run on Render only."
        )

    return psycopg2.connect(db_url)
