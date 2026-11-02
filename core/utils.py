import json
import re
from datetime import datetime
from typing import Any, Optional

import pandas as pd


def safe_str(v: Any) -> str:
    return "" if pd.isna(v) else str(v).strip()


def normalize_whitespace(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()


def normalize_header(s: str) -> str:
    return normalize_whitespace(str(s).lower())


def format_phone(v: Any) -> str:
    s = safe_str(v)
    if not s:
        return ""

    s = s.replace("\n", ";").replace("/", ";").replace("|", ";").replace(",", ";")
    parts = [p.strip() for p in s.split(";") if p.strip()]

    normalized = []
    for part in parts:
        if part.endswith(".0"):
            part = part[:-2]
        digits = re.sub(r"\D", "", part)
        if not digits:
            continue

        if len(digits) > 15:
            chunks = re.findall(r"\d{11,12}", digits)
            if chunks:
                normalized.extend(["+" + c for c in chunks])
            else:
                normalized.append("+" + digits)
        else:
            normalized.append("+" + digits)

    seen = set()
    out = []
    for n in normalized:
        if n not in seen:
            seen.add(n)
            out.append(n)

    return "; ".join(out)


def adoszam_to_kshkod(value: Any) -> Optional[str]:
    if pd.isna(value):
        return None
    digits = re.sub(r"\D", "", str(value).strip())
    if len(digits) < 8:
        return None
    return digits[:8]


def parse_date_to_str(v: Any) -> str:
    if pd.isna(v):
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    try:
        ts = pd.to_datetime(v, errors="coerce")
        if pd.isna(ts):
            return safe_str(v)
        return ts.strftime("%Y-%m-%d")
    except Exception:
        return safe_str(v)


def split_address_parts(sz: Any):
    s = normalize_whitespace(safe_str(sz))
    if not s:
        return {
            "Vármegye": "",
            "IRSZ": "",
            "Település": "",
            "Cím": "",
        }

    parts = [p.strip() for p in s.split(",") if p.strip()]
    out = {
        "Vármegye": "",
        "IRSZ": "",
        "Település": "",
        "Cím": "",
    }

    if len(parts) > 0:
        out["Vármegye"] = parts[0]
    if len(parts) > 1:
        out["IRSZ"] = parts[1]
    if len(parts) > 2:
        out["Település"] = parts[2]
    if len(parts) > 3:
        out["Cím"] = ", ".join(parts[3:])

    return out


def try_json_load(text: str):
    try:
        return json.loads(text)
    except Exception:
        return None


def _parse_single_hu_number(value: Any):
    s = safe_str(value)
    if not s:
        return None

    s = s.replace("\xa0", " ").strip()
    s = re.sub(r"[^0-9,\.\-\s]", "", s).strip()
    if not s:
        return None

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        last = s.split(",")[-1]
        if len(last) in (1, 2):
            s = s.replace(" ", "")
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(" ", "")
            s = s.replace(",", "")
    else:
        s = s.replace(" ", "")

    try:
        return float(s)
    except Exception:
        return None


def parse_hu_numeric(series: pd.Series) -> pd.Series:
    return series.apply(_parse_single_hu_number)


def clean_int_like(value) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return ""
    try:
        f = float(s.replace(",", "."))
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return s
