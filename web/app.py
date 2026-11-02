import hmac
import io
import os
from pathlib import Path
import secrets
import threading
import time
import zipfile

import pandas as pd
from flask import Flask, abort, jsonify, redirect, render_template, request, send_file, session
from werkzeug.exceptions import HTTPException

from core.constants import ADATTIPUS_LABELS, CACHE_DB, PRODUCT_GROUPS_JSON
from services.cache_db import CacheDB
from services.category_service import CategoryService
from services.export_service import (
    write_dataset_excel, write_detail_excel, write_hak_matrix_excel, write_ktj_excel,
    write_missing_excel, write_summary_excel,
)
from services.fetch_service import fetch_data
from services.filter_service import Filters
from services.import_service import build_source_from_excel_by_adoszam
from services.okir_client import OkirClient
from services.persistence_service import PersistenceService
from services.report_service import csv_bytes, dataset_bytes, read_dataset, records, report_html
from services.package_service import read_package, write_package
from web.workspace import Workspace


def create_app(config=None):
    app = Flask(__name__)
    app.config.from_mapping(
        SECRET_KEY=os.environ.get("MEH_SECRET_KEY") or secrets.token_hex(32),
        WEB_PASSWORD=os.environ.get("MEH_WEB_PASSWORD"),
        DATABASE=CACHE_DB,
        MAX_CONTENT_LENGTH=20 * 1024 * 1024,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Strict",
        SESSION_COOKIE_SECURE=os.environ.get("MEH_HTTPS") == "1",
        PERMANENT_SESSION_LIFETIME=8 * 60 * 60,
    )
    if config:
        app.config.update(config)
    if not app.config["WEB_PASSWORD"] or len(app.config["WEB_PASSWORD"]) < 12:
        raise ValueError("Állíts be legalább 12 karakteres MEH_WEB_PASSWORD jelszót.")
    persistence = PersistenceService(Path(app.config["DATABASE"]))
    workspace = Workspace(persistence, CategoryService(PRODUCT_GROUPS_JSON))
    app.extensions["workspace"] = workspace
    login_lock = threading.Lock()
    attempts = []

    @app.before_request
    def protect():
        if "csrf" not in session:
            session["csrf"] = secrets.token_hex(32)
        public = request.endpoint in {"login", "static", "service_worker", "health"}
        if not public and not session.get("authenticated"):
            if request.path.startswith("/api/"):
                abort(401, description="A munkamenet lejárt. Jelentkezz be újra.")
            return redirect("/login")
        if request.method not in {"GET", "HEAD", "OPTIONS"}:
            token = request.headers.get("X-CSRF-Token") or request.form.get("csrf", "")
            if not hmac.compare_digest(token, session["csrf"]):
                abort(403, description="Érvénytelen munkamenet. Frissítsd az oldalt.")

    @app.after_request
    def headers(response):
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Referrer-Policy"] = "same-origin"
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; "
            "img-src 'self' data:; connect-src 'self'; object-src 'none'; "
            "base-uri 'none'; frame-ancestors 'none'; form-action 'self'"
        )
        if request.endpoint != "static":
            response.headers["Cache-Control"] = "no-store"
        return response

    @app.teardown_appcontext
    def close_connection(_error):
        persistence.close()

    @app.errorhandler(Exception)
    def error_response(exc):
        if isinstance(exc, HTTPException):
            code, message = exc.code, exc.description
        elif isinstance(exc, (ValueError, KeyError)):
            code, message = 400, str(exc)
        else:
            app.logger.exception("Webes kérés sikertelen")
            code, message = 500, "Szerverhiba történt. Részletek a szerver naplójában."
        if request.path.startswith("/api/"):
            return jsonify(error=message), code
        return render_template("error.html", message=message), code

    @app.get("/health")
    def health():
        return jsonify(status="ok")

    @app.route("/login", methods=["GET", "POST"])
    def login():
        error = None
        if request.method == "POST":
            with login_lock:
                now = time.monotonic()
                attempts[:] = [stamp for stamp in attempts if now - stamp < 60]
                if len(attempts) >= 10:
                    abort(429, description="Túl sok belépési kísérlet. Várj egy percet.")
                attempts.append(now)
            if hmac.compare_digest(request.form.get("password", "").encode(), app.config["WEB_PASSWORD"].encode()):
                session.clear()
                session.update(authenticated=True, csrf=secrets.token_hex(32))
                session.permanent = True
                return redirect("/")
            error = "Hibás jelszó."
        return render_template("login.html", error=error)

    @app.post("/api/logout")
    def logout():
        session.clear()
        return jsonify(ok=True)

    @app.get("/")
    def index():
        return render_template("index.html")

    @app.get("/sw.js")
    def service_worker():
        return app.send_static_file("sw.js")

    @app.get("/api/state")
    def state():
        with workspace.lock:
            dataset = workspace.dataset
            master = dataset.aggregates.master
            options = {}
            for name, column in (("companies", "CEGNEV"), ("years", "EV"), ("categories", "PRODUCT_GROUP")):
                values = master[column].dropna().unique() if column in master else []
                options[name] = sorted((int(v) for v in values), reverse=True) if name == "years" else sorted(str(v) for v in values)
            return jsonify(csrf=session["csrf"], filename=dataset.filename, options=options,
                           types=ADATTIPUS_LABELS, rows=len(dataset.raw), missing=len(dataset.missing),
                           job=dict(workspace.job), snapshots=persistence.list_snapshots())

    def filters():
        kinds = tuple(request.args.getlist("types"))
        if any(kind not in ADATTIPUS_LABELS for kind in kinds):
            raise ValueError("Érvénytelen adattípus.")
        return Filters(
            companies=tuple(request.args.getlist("companies")),
            years=tuple(int(year) for year in request.args.getlist("years")),
            categories=tuple(request.args.getlist("categories")),
            code=request.args.get("code", "")[:500], types=kinds,
        )

    @app.get("/api/view")
    def view():
        page = max(1, int(request.args.get("page", 1)))
        page_size = min(100, max(1, int(request.args.get("page_size", 50))))
        selected = filters()
        kind = request.args.get("view", "master")
        with workspace.lock:
            dataset = workspace.dataset
            master, raw, totals = workspace.filtered(selected)
            if kind == "master":
                frame = master
            elif kind == "ktj":
                frame = raw
            elif kind == "missing":
                frame = dataset.missing
            elif kind == "source":
                frame = dataset.source
                if selected.companies and not frame.empty:
                    frame = frame[frame["CEGNEV"].isin(selected.companies)]
            else:
                raise ValueError("Ismeretlen nézet.")
            total = len(frame)
            page = min(page, max(1, (total + page_size - 1) // page_size))
            frame = frame.iloc[(page - 1) * page_size:page * page_size]
            if kind == "ktj":
                frame = workspace.ktj.enrich_with_metadata(frame)
                if not frame.empty:
                    frame["PRODUCT_GROUP"] = frame["HULLADEKKOD"].map(workspace.categories.resolve_group)
            return jsonify(columns=list(frame.columns), rows=records(frame), total=total,
                           page=page, page_size=page_size, summary=totals)

    def submit(operation):
        try:
            workspace.submit(operation)
        except RuntimeError as exc:
            abort(409, description=str(exc))
        return jsonify(status="running"), 202

    @app.post("/api/import")
    def import_data():
        upload = request.files.get("file")
        if not upload or not upload.filename:
            raise ValueError("Válassz .xlsx, MÉH .json vagy .okirpkg fájlt.")
        name = Path(upload.filename.replace("\\", "/")).name
        extension = Path(name).suffix.lower()
        if extension not in {".xlsx", ".json", ".okirpkg", ".pkg", ".zip"}:
            raise ValueError("Csak .xlsx, MÉH .json és PKG fájl tölthető fel.")
        content = upload.read()
        kinds = request.form.getlist("types")
        if extension == ".xlsx" and (not kinds or any(kind not in ADATTIPUS_LABELS for kind in kinds)):
            raise ValueError("Válassz legalább egy adattípust.")
        use_cache = request.form.get("use_cache") == "true"
        refresh = request.form.get("force_refresh") == "true"

        def operation():
            metadata = {}
            if extension == ".json":
                source, raw, missing = read_dataset(content)
            elif extension in {".okirpkg", ".pkg", ".zip"}:
                source, raw, missing, _state, metadata = read_package(content)
            else:
                with zipfile.ZipFile(io.BytesIO(content)) as archive:
                    if sum(item.file_size for item in archive.infolist()) > 100 * 1024 * 1024:
                        raise ValueError("A kicsomagolt Excel fájl túl nagy (maximum 100 MB).")
                # Bound network work as well as compressed upload size.
                preview = pd.read_excel(io.BytesIO(content), nrows=2001)
                if len(preview) > 2000:
                    raise ValueError("Egyszerre legfeljebb 2000 forráscég importálható.")
                failures = []
                client = OkirClient()
                try:
                    source = build_source_from_excel_by_adoszam(
                        io.BytesIO(content), client, progress_fn=workspace.progress,
                        missing_fn=failures.append, log_fn=app.logger.info,
                    )
                finally:
                    client.session.close()
                if source.empty:
                    raw = pd.DataFrame()
                    missing = pd.DataFrame(failures)
                else:
                    result = fetch_data(source, use_cache, refresh, kinds, progress=workspace.progress,
                                        cache_path=Path(app.config["DATABASE"]))
                    raw = result.raw
                    missing = pd.concat([pd.DataFrame(failures), result.missing], ignore_index=True)
            if source.empty and raw.empty and missing.empty:
                raise ValueError("A fájl nem tartalmaz feldolgozható adatot.")
            workspace.replace(source, raw, missing, name, metadata=metadata)
            return f"Import kész: {len(raw)} adatsor, {len(missing)} kimaradt/hibás lekérés."

        return submit(operation)

    @app.post("/api/ktj")
    def load_ktj():
        with workspace.lock:
            raw = workspace.filtered(filters())[1]
        if raw.empty:
            raise ValueError("Nincs betölthető KTJ adat.")

        def operation():
            workspace.ktj.ensure_metadata_loaded(raw, workspace.progress)
            return "KTJ metaadatok betöltve."
        return submit(operation)

    @app.post("/api/snapshots")
    def save_snapshot():
        note = str((request.get_json() or {}).get("note", ""))[:200]

        def operation():
            dataset = workspace.dataset
            if dataset.source.empty and dataset.raw.empty and dataset.missing.empty:
                raise ValueError("Nincs menthető adat.")
            sid = persistence.save_snapshot(dataset.source, dataset.raw, dataset.missing, dataset.filename, note)
            persistence.set_setting("web_active_snapshot", str(sid))
            return f"Mentés elkészült: #{sid}"
        return submit(operation)

    @app.post("/api/snapshots/<int:sid>/load")
    def load_snapshot(sid):
        def operation():
            row = next((row for row in persistence.list_snapshots() if row[0] == sid), None)
            if row is None:
                raise ValueError("A mentés nem található.")
            workspace.replace(*persistence.load_snapshot(sid), row[2], persist=False)
            persistence.set_setting("web_active_snapshot", str(sid))
            return f"Mentés betöltve: #{sid}"
        return submit(operation)

    @app.delete("/api/snapshots/<int:sid>")
    def delete_snapshot(sid):
        def operation():
            if persistence.get_setting("web_active_snapshot") == str(sid):
                raise ValueError("Az aktív mentés nem törölhető. Előbb tölts be egy másik mentést.")
            if not any(row[0] == sid for row in persistence.list_snapshots()):
                raise ValueError("A mentés nem található.")
            persistence.delete_snapshot(sid)
            return "Mentés törölve."
        return submit(operation)

    @app.delete("/api/cache")
    def clear_cache():
        def operation():
            cache = CacheDB(Path(app.config["DATABASE"]))
            try:
                cache.clear_all()
            finally:
                cache.close()
            return "OKIR lekérési gyorsítótár törölve."
        return submit(operation)

    @app.get("/api/export/<kind>")
    def export(kind):
        if kind not in {
            "matrix", "summary", "detail", "ktj", "errors", "report",
            "pkg", "json", "dataset", "csv", "source-csv", "missing-csv",
        }:
            abort(404)
        selected = filters()
        with workspace.lock:
            dataset = workspace.dataset
            master, raw, _totals = workspace.filtered(selected)
            metadata = workspace.ktj.export_metadata(dataset.raw) if kind == "pkg" else {}
        if kind in {"pkg", "json", "dataset", "csv", "source-csv", "missing-csv"} and all(
            frame.empty for frame in (dataset.source, dataset.raw, dataset.missing)
        ):
            raise ValueError("Nincs exportálható adatállomány.")
        description = " | ".join(
            f"{label}: {', '.join(map(str, values))}"
            for label, values in (("Cég", selected.companies), ("Év", selected.years),
                                  ("Termékcsoport", selected.categories), ("HAK", (selected.code,) if selected.code else ()),
                                  ("Adattípus", tuple(ADATTIPUS_LABELS[k] for k in selected.types)))
            if values
        ) or "Összes adat"
        if kind == "report":
            return report_html(master, dataset.missing, description)
        buffer = io.BytesIO()
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        if kind == "pkg":
            write_package(buffer, dataset.source, dataset.raw, dataset.missing, metadata=metadata)
            filename, mime = "meh-megoszthato.okirpkg", "application/zip"
        elif kind == "json":
            buffer.write(dataset_bytes(dataset.source, dataset.raw, dataset.missing))
            filename, mime = "meh-adatok.json", "application/json"
        elif kind == "dataset":
            write_dataset_excel(buffer, dataset.source, dataset.raw, dataset.missing)
            filename = "meh-teljes-adatallomany.xlsx"
        elif kind in {"csv", "source-csv", "missing-csv"}:
            frame = raw
            if kind == "source-csv":
                frame = dataset.source
                if selected.companies and not frame.empty:
                    frame = frame[frame["CEGNEV"].isin(selected.companies)]
            elif kind == "missing-csv":
                frame = dataset.missing
            buffer.write(csv_bytes(frame))
            filename, mime = f"meh-{kind}.csv", "text/csv; charset=utf-8"
        elif kind == "matrix":
            write_hak_matrix_excel(buffer, master, description)
            filename = "meh-ceg-hak-matrix.xlsx"
        elif kind == "summary":
            write_summary_excel(buffer, master, dataset.source, workspace.categories.all_groups())
            filename = "meh-osszesito.xlsx"
        elif kind == "detail":
            write_detail_excel(buffer, master, dataset.source)
            filename = "meh-reszletes.xlsx"
        elif kind == "ktj":
            # Metadata loading is a separate, visible background operation.
            with workspace.lock:
                write_ktj_excel(buffer, raw, workspace.ktj, workspace.categories, load_metadata=False)
            filename = "meh-ktj.xlsx"
        else:
            write_missing_excel(buffer, dataset.missing)
            filename = "meh-kimaradt-lekeresek.xlsx"
        buffer.seek(0)
        return send_file(buffer, mimetype=mime, as_attachment=True, download_name=filename)

    return app
