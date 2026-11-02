import io
import json
import pickle
import zipfile

import numpy as np
import pandas as pd

from core.constants import ADATTIPUS_LABELS
from services.report_service import dataset_bytes, read_dataset

MAX_PACKAGE_BYTES = 100 * 1024 * 1024


class _LegacyTableReader(pickle.Unpickler):
    """Only reconstruct the pandas/numpy types used by the original exporter."""

    def find_class(self, module, name):
        from pandas._libs.internals import _unpickle_block
        from pandas.core.indexes.base import _new_Index
        from pandas.core.internals.managers import BlockManager
        try:
            from numpy._core import multiarray, numeric
        except ModuleNotFoundError:
            from numpy.core import multiarray, numeric

        allowed = {
            ("pandas.core.frame", "DataFrame"): pd.DataFrame,
            ("pandas.core.internals.managers", "BlockManager"): BlockManager,
            ("pandas._libs.internals", "_unpickle_block"): _unpickle_block,
            ("pandas.core.indexes.base", "_new_Index"): _new_Index,
            ("pandas.core.indexes.base", "Index"): pd.Index,
            ("pandas.core.indexes.range", "RangeIndex"): pd.RangeIndex,
            ("numpy", "ndarray"): np.ndarray,
            ("numpy", "dtype"): np.dtype,
            ("builtins", "slice"): slice,
        }
        for prefix in ("numpy.core", "numpy._core"):
            allowed[(prefix + ".multiarray", "_reconstruct")] = multiarray._reconstruct
            allowed[(prefix + ".multiarray", "scalar")] = multiarray.scalar
            allowed[(prefix + ".numeric", "_frombuffer")] = numeric._frombuffer
        if (module, name) not in allowed:
            raise ValueError(f"Nem támogatott típus a régi PKG-ben: {module}.{name}")
        return allowed[(module, name)]


def validate_extras(state, metadata):
    if not isinstance(state, dict) or not isinstance(metadata, dict):
        raise ValueError("Érvénytelen PKG állapot vagy KTJ-metaadat.")
    for key in ("current_file_path", "current_query_label", "filter_code"):
        if key in state and not isinstance(state[key], str):
            raise ValueError(f"Érvénytelen PKG állapot: {key}")
    for key in ("selected_companies", "selected_categories", "selected_types"):
        if key in state and (
            not isinstance(state[key], list) or not all(isinstance(v, str) for v in state[key])
        ):
            raise ValueError(f"Érvénytelen PKG kijelölés: {key}")
    if any(kind not in ADATTIPUS_LABELS for kind in state.get("selected_types", [])):
        raise ValueError("Érvénytelen PKG adattípus.")
    years = state.get("selected_years", [])
    if not isinstance(years, list) or not all(type(v) is int and 1900 <= v <= 2200 for v in years):
        raise ValueError("Érvénytelen PKG évszűrő.")
    if "current_tab_index" in state and type(state["current_tab_index"]) is not int:
        raise ValueError("Érvénytelen PKG nézet.")
    if "current_use_cache" in state and not isinstance(state["current_use_cache"], bool):
        raise ValueError("Érvénytelen PKG gyorsítótár-beállítás.")
    if state.get("drill_category") is not None and not isinstance(state["drill_category"], str):
        raise ValueError("Érvénytelen PKG termékcsoport.")
    if not all(isinstance(key, str) and isinstance(value, dict) for key, value in metadata.items()):
        raise ValueError("Érvénytelen PKG KTJ-metaadat.")


def write_package(target, source, raw, missing, state=None, metadata=None):
    if all(frame.empty for frame in (source, raw, missing)):
        raise ValueError("Nincs exportálható adatállomány.")
    state = {} if state is None else state
    metadata = {} if metadata is None else metadata
    validate_extras(state, metadata)
    contents = {
        "dataset.json": dataset_bytes(source, raw, missing),
        "state.json": json.dumps(state, ensure_ascii=False, allow_nan=False).encode("utf-8"),
        "ktj_meta.json": json.dumps(metadata, ensure_ascii=False, allow_nan=False).encode("utf-8"),
    }
    if sum(map(len, contents.values())) > MAX_PACKAGE_BYTES:
        raise ValueError("A PKG kicsomagolt mérete legfeljebb 100 MB lehet.")
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, content in contents.items():
            archive.writestr(name, content)


def read_package(data, *, allow_legacy=False):
    if len(data) > MAX_PACKAGE_BYTES:
        raise ValueError("A PKG fájl legfeljebb 100 MB lehet.")
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            entries = archive.infolist()
            names = [entry.filename for entry in entries]
            if len(names) != len(set(names)) or len(names) > 8:
                raise ValueError("Érvénytelen PKG fájllista.")
            if sum(entry.file_size for entry in entries) > MAX_PACKAGE_BYTES:
                raise ValueError("A PKG kicsomagolt mérete legfeljebb 100 MB lehet.")
            allowed = {"dataset.json", "state.json", "ktj_meta.json", "source.pkl", "raw.pkl", "missing.pkl"}
            if set(names) - allowed:
                raise ValueError("Ismeretlen fájl a PKG csomagban.")
            legacy = "dataset.json" not in names
            if legacy:
                if not allow_legacy:
                    raise ValueError("Régi PKG: az asztali alkalmazásban nyisd meg, majd exportáld új PKG-ként.")
                if not {"source.pkl", "raw.pkl", "missing.pkl"}.issubset(names):
                    raise ValueError("Hiányos régi PKG csomag.")
                frames = []
                for name in ("source.pkl", "raw.pkl", "missing.pkl"):
                    frame = _LegacyTableReader(io.BytesIO(archive.read(name))).load()
                    if type(frame) is not pd.DataFrame:
                        raise ValueError("A régi PKG nem adattáblát tartalmaz.")
                    frames.append(frame)
                frames = read_dataset(dataset_bytes(*frames))
            else:
                if any(name.endswith(".pkl") for name in names):
                    raise ValueError("Kevert régi és új PKG formátum.")
                frames = read_dataset(archive.read("dataset.json"))
            state = json.loads(archive.read("state.json")) if "state.json" in names else {}
            metadata = json.loads(archive.read("ktj_meta.json")) if "ktj_meta.json" in names else {}
            validate_extras(state, metadata)
            if all(frame.empty for frame in frames):
                raise ValueError("A csomag nem tartalmaz feldolgozható adatot.")
            return *frames, state, metadata
    except (zipfile.BadZipFile, pickle.UnpicklingError, EOFError) as exc:
        raise ValueError("Sérült vagy nem támogatott PKG csomag.") from exc
