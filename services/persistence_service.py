import json
import sqlite3
import threading
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd


class PersistenceService:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self._local = threading.local()
        self._init_db()

    def _conn(self):
        if not hasattr(self._local, "conn") or self._local.conn is None:
            conn = sqlite3.connect(self.db_path, timeout=30)
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA synchronous=NORMAL;")
            conn.execute("PRAGMA busy_timeout=30000;")
            self._local.conn = conn
        return self._local.conn

    def _init_db(self):
        conn = sqlite3.connect(self.db_path, timeout=30)
        try:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS snapshots (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    created_at TEXT NOT NULL,
                    source_file TEXT,
                    note TEXT
                )
            """)
            for table in ("snapshot_source_rows", "snapshot_raw_rows", "snapshot_missing_rows"):
                conn.execute(f"""
                    CREATE TABLE IF NOT EXISTS {table} (
                        snapshot_id INTEGER NOT NULL,
                        data_json TEXT NOT NULL
                    )
                """)
                conn.execute(f"CREATE INDEX IF NOT EXISTS idx_{table}_snapshot ON {table}(snapshot_id)")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS ktj_meta_cache (
                    ktj TEXT PRIMARY KEY,
                    payload_json TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                )
            """)
            conn.execute("CREATE TABLE IF NOT EXISTS app_settings (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
            conn.commit()
        finally:
            conn.close()

    def save_snapshot(
        self,
        source_df: pd.DataFrame,
        raw_df: pd.DataFrame,
        missing_df: pd.DataFrame,
        source_file: str = "",
        note: str = "",
    ) -> int:
        with self._conn() as conn:
            cur = conn.execute(
                "INSERT INTO snapshots (created_at, source_file, note) VALUES (?, ?, ?)",
                (datetime.now().isoformat(), source_file, note),
            )
            snapshot_id = cur.lastrowid
            for table, frame in (
                ("snapshot_source_rows", source_df),
                ("snapshot_raw_rows", raw_df),
                ("snapshot_missing_rows", missing_df),
            ):
                conn.executemany(
                    f"INSERT INTO {table} (snapshot_id, data_json) VALUES (?, ?)",
                    ((snapshot_id, json.dumps(row, ensure_ascii=False))
                     for row in frame.fillna("").to_dict(orient="records")),
                )
        return snapshot_id

    def close(self):
        conn = getattr(self._local, "conn", None)
        if conn is not None:
            conn.close()
            self._local.conn = None

    def get_setting(self, key):
        row = self._conn().execute("SELECT value FROM app_settings WHERE key=?", (key,)).fetchone()
        return row[0] if row else None

    def set_setting(self, key, value):
        with self._conn() as conn:
            conn.execute("INSERT INTO app_settings VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                         (key, value))

    def list_snapshots(self) -> List[Tuple[int, str, str, str]]:
        return self._conn().execute(
            "SELECT id, created_at, source_file, note FROM snapshots ORDER BY id DESC"
        ).fetchall()

    def delete_snapshot(self, snapshot_id: int) -> None:
        with self._conn() as conn:
            for table in ("snapshot_source_rows", "snapshot_raw_rows", "snapshot_missing_rows"):
                conn.execute(f"DELETE FROM {table} WHERE snapshot_id=?", (int(snapshot_id),))
            conn.execute("DELETE FROM snapshots WHERE id=?", (int(snapshot_id),))

    def _load_rows(self, conn, table_name: str, snapshot_id: int):
        rows = conn.execute(
            f"SELECT data_json FROM {table_name} WHERE snapshot_id=?",
            (snapshot_id,),
        ).fetchall()
        return pd.DataFrame([json.loads(row[0]) for row in rows])

    def load_snapshot(self, snapshot_id: int):
        conn = self._conn()
        return tuple(self._load_rows(conn, table, snapshot_id) for table in (
            "snapshot_source_rows", "snapshot_raw_rows", "snapshot_missing_rows",
        ))

    def load_ktj_meta(self, ktj: str) -> Optional[Dict]:
        row = self._conn().execute(
            "SELECT payload_json FROM ktj_meta_cache WHERE ktj = ?",
            (str(ktj).strip(),),
        ).fetchone()
        return json.loads(row[0]) if row and row[0] else None

    def save_ktj_meta(self, ktj: str, payload: Dict):
        with self._conn() as conn:
            conn.execute("""
                INSERT INTO ktj_meta_cache (ktj, payload_json, updated_at)
                VALUES (?, ?, ?)
                ON CONFLICT(ktj) DO UPDATE SET
                    payload_json = excluded.payload_json,
                    updated_at = excluded.updated_at
            """, (str(ktj).strip(), json.dumps(payload or {}, ensure_ascii=False), datetime.now().isoformat()))
