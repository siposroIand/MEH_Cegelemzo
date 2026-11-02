import json
import pickle
import sqlite3
import threading
import zipfile
from datetime import datetime
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Any

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
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS snapshots (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                source_file TEXT,
                note TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS snapshot_source_rows (
                snapshot_id INTEGER NOT NULL,
                data_json TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS snapshot_raw_rows (
                snapshot_id INTEGER NOT NULL,
                data_json TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS snapshot_missing_rows (
                snapshot_id INTEGER NOT NULL,
                data_json TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS ktj_meta_cache (
                ktj TEXT PRIMARY KEY,
                payload_json TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
        """)
        for table in ("snapshot_source_rows", "snapshot_raw_rows", "snapshot_missing_rows"):
            conn.execute(f"CREATE INDEX IF NOT EXISTS idx_{table}_snapshot ON {table}(snapshot_id)")
        conn.execute("CREATE TABLE IF NOT EXISTS app_settings (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
        conn.commit()
        conn.close()

    # ------------------------------------------------------------------
    # Snapshot
    # ------------------------------------------------------------------
    def save_snapshot(
        self,
        source_df: pd.DataFrame,
        raw_df: pd.DataFrame,
        missing_df: pd.DataFrame,
        source_file: str = "",
        note: str = "",
    ) -> int:
        conn = self._conn()
        with conn:
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
        conn = self._conn()
        rows = conn.execute(
            "SELECT id, created_at, source_file, note FROM snapshots ORDER BY id DESC"
        ).fetchall()
        return rows

    def delete_snapshot(self, snapshot_id: int) -> None:
        """Töröl egy snapshotot és a hozzá tartozó sorokat."""
        conn = self._conn()
        cur = conn.cursor()
        # transaction
        cur.execute("DELETE FROM snapshot_source_rows WHERE snapshot_id=?", (int(snapshot_id),))
        cur.execute("DELETE FROM snapshot_raw_rows WHERE snapshot_id=?", (int(snapshot_id),))
        cur.execute("DELETE FROM snapshot_missing_rows WHERE snapshot_id=?", (int(snapshot_id),))
        cur.execute("DELETE FROM snapshots WHERE id=?", (int(snapshot_id),))
        conn.commit()

    def _load_rows(self, conn, table_name: str, snapshot_id: int):
        rows = conn.execute(
            f"SELECT data_json FROM {table_name} WHERE snapshot_id=?",
            (snapshot_id,),
        ).fetchall()
        if not rows:
            return pd.DataFrame()
        parsed = [json.loads(r[0]) for r in rows]
        return pd.DataFrame(parsed)

    def load_snapshot(self, snapshot_id: int):
        conn = self._conn()
        source_df = self._load_rows(conn, "snapshot_source_rows", snapshot_id)
        raw_df = self._load_rows(conn, "snapshot_raw_rows", snapshot_id)
        missing_df = self._load_rows(conn, "snapshot_missing_rows", snapshot_id)
        return source_df, raw_df, missing_df

    # ------------------------------------------------------------------
    # KTJ meta cache
    # ------------------------------------------------------------------
    def load_ktj_meta(self, ktj: str) -> Optional[Dict]:
        conn = self._conn()
        row = conn.execute(
            "SELECT payload_json FROM ktj_meta_cache WHERE ktj = ?",
            (str(ktj).strip(),),
        ).fetchone()
        if not row:
            return None

        try:
            return json.loads(row[0]) if row[0] else None
        except Exception:
            return None

    def save_ktj_meta(self, ktj: str, payload: Dict):
        conn = self._conn()
        conn.execute(
            """
            INSERT INTO ktj_meta_cache (ktj, payload_json, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(ktj) DO UPDATE SET
                payload_json = excluded.payload_json,
                updated_at = excluded.updated_at
            """,
            (
                str(ktj).strip(),
                json.dumps(payload or {}, ensure_ascii=False),
                datetime.now().isoformat(),
            ),
        )
        conn.commit()

    # ------------------------------------------------------------------
    # Shared package export / import
    # ------------------------------------------------------------------
    def export_shared_package(
        self,
        package_path: str,
        source_df: pd.DataFrame,
        raw_df: pd.DataFrame,
        missing_df: pd.DataFrame,
        state: Dict[str, Any],
        ktj_meta_cache: Optional[Dict[str, Any]] = None,
    ):
        package_path = str(package_path)

        with zipfile.ZipFile(package_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("state.json", json.dumps(state or {}, ensure_ascii=False, indent=2))
            zf.writestr("ktj_meta.json", json.dumps(ktj_meta_cache or {}, ensure_ascii=False, indent=2))
            zf.writestr("source.pkl", pickle.dumps(source_df))
            zf.writestr("raw.pkl", pickle.dumps(raw_df))
            zf.writestr("missing.pkl", pickle.dumps(missing_df))

    def import_shared_package(self, package_path: str):
        package_path = str(package_path)

        with zipfile.ZipFile(package_path, "r") as zf:
            state = json.loads(zf.read("state.json").decode("utf-8")) if "state.json" in zf.namelist() else {}
            ktj_meta = json.loads(zf.read("ktj_meta.json").decode("utf-8")) if "ktj_meta.json" in zf.namelist() else {}
            source_df = pickle.loads(zf.read("source.pkl")) if "source.pkl" in zf.namelist() else pd.DataFrame()
            raw_df = pickle.loads(zf.read("raw.pkl")) if "raw.pkl" in zf.namelist() else pd.DataFrame()
            missing_df = pickle.loads(zf.read("missing.pkl")) if "missing.pkl" in zf.namelist() else pd.DataFrame()

        return source_df, raw_df, missing_df, state, ktj_meta

    # ------------------------------------------------------------------
    # Shared package merge
    # ------------------------------------------------------------------
    def _dedup_source_df(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame()
        out = df.copy()
        if "KUJ" in out.columns:
            out["KUJ"] = out["KUJ"].astype(str)
            return out.drop_duplicates(subset=["KUJ"]).reset_index(drop=True)
        if "CEGNEV" in out.columns and "KUJ" in out.columns:
            return out.drop_duplicates(subset=["CEGNEV", "KUJ"]).reset_index(drop=True)
        return out.drop_duplicates().reset_index(drop=True)

    def _dedup_missing_df(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame()
        return df.drop_duplicates().reset_index(drop=True)

    def _dedup_raw_df(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Deduplikálás biztonságos kulcsokkal.
        Nem összeadunk, hanem azonos sorokat eldobjuk.
        """
        if df is None or df.empty:
            return pd.DataFrame()

        out = df.copy()

        preferred_keys = [
            "SOURCE_KUJ",
            "CEGNEV",
            "EV",
            "ADATTIPUS",
            "ADATTIPUS_LABEL",
            "HULLADEKKOD",
            "HULLADEK_MEGNEVEZES",
            "HULLADEK_VESZELYES_10",
            "KTJ",
            "KUJ",
            "MENNYISEG",
        ]
        keys = [c for c in preferred_keys if c in out.columns]
        if keys:
            # string normalize pár oszlopra
            for c in ["SOURCE_KUJ", "KUJ", "KTJ", "HULLADEKKOD", "CEGNEV", "ADATTIPUS", "ADATTIPUS_LABEL"]:
                if c in out.columns:
                    out[c] = out[c].astype(str)
            return out.drop_duplicates(subset=keys).reset_index(drop=True)

        return out.drop_duplicates().reset_index(drop=True)

    def merge_shared_packages(
        self,
        package_paths: List[str],
        state: Optional[Dict[str, Any]] = None,
        prefer_first_state: bool = True,
        ktj_prefer_first: bool = True,
    ):
        """
        Több okirpkg beolvasása és összevonása.
        - state: ha meg van adva, ezt írjuk a merged csomagba
        - prefer_first_state: ha state nincs megadva, az első pkg state-jét tartjuk meg (különben üres/default)
        - ktj_prefer_first: ütközés esetén az első pkg meta adata marad (különben felülírjuk)
        """
        if not package_paths:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), {}, {}

        sources = []
        raws = []
        missings = []
        merged_ktj: Dict[str, Any] = {}
        chosen_state: Dict[str, Any] = {}

        for idx, p in enumerate(package_paths):
            src_df, raw_df, missing_df, st, ktj = self.import_shared_package(p)

            if idx == 0 and prefer_first_state and not state:
                chosen_state = st or {}

            sources.append(src_df if isinstance(src_df, pd.DataFrame) else pd.DataFrame())
            raws.append(raw_df if isinstance(raw_df, pd.DataFrame) else pd.DataFrame())
            missings.append(missing_df if isinstance(missing_df, pd.DataFrame) else pd.DataFrame())

            if isinstance(ktj, dict):
                for k, v in ktj.items():
                    if ktj_prefer_first:
                        if k not in merged_ktj:
                            merged_ktj[k] = v
                    else:
                        merged_ktj[k] = v

        source_df = self._dedup_source_df(pd.concat(sources, ignore_index=True) if sources else pd.DataFrame())
        raw_df = self._dedup_raw_df(pd.concat(raws, ignore_index=True) if raws else pd.DataFrame())
        missing_df = self._dedup_missing_df(pd.concat(missings, ignore_index=True) if missings else pd.DataFrame())

        final_state = state if isinstance(state, dict) else chosen_state
        if final_state is None:
            final_state = {}

        return source_df, raw_df, missing_df, final_state, merged_ktj

    def export_merged_shared_package(
        self,
        out_package_path: str,
        package_paths: List[str],
        state: Optional[Dict[str, Any]] = None,
    ) -> None:
        source_df, raw_df, missing_df, final_state, ktj_meta = self.merge_shared_packages(
            package_paths=package_paths,
            state=state,
            prefer_first_state=True,
            ktj_prefer_first=True,
        )
        self.export_shared_package(
            package_path=out_package_path,
            source_df=source_df,
            raw_df=raw_df,
            missing_df=missing_df,
            state=final_state,
            ktj_meta_cache=ktj_meta,
        )