from pathlib import Path
import os

APP_DIR = Path(os.environ.get("MEH_DATA_HOME", Path.cwd())).resolve()
DATA_DIR = APP_DIR / "data"
CONFIG_DIR = Path(__file__).resolve().parents[1] / "config"

DATA_DIR.mkdir(parents=True, exist_ok=True)

CACHE_DB = DATA_DIR / "okir_cache.db"
PRODUCT_GROUPS_JSON = CONFIG_DIR / "product_groups.json"

KSH_SEARCH_URL = "https://web.okir.hu/licoms/dbb-mybatis/KARLOW01/KSHKOD/{kshkod}/ORDER_BY/KUJ/DIR/ASC/"

ENDPOINTS = {
    "HULL_KELETKEZES": "https://web.okir.hu/licoms/dbb-popup/EHIRLOW19/KUJ/{kuj}/QU/HULL_KELETKEZES/",
    "HULL_ATVETEL": "https://web.okir.hu/licoms/dbb-popup/EHIRLOW21/KUJ/{kuj}/QU/HULL_ATVETEL/",
    "HULL_ATADAS": "https://web.okir.hu/licoms/dbb-popup/EHIRLOW20/KUJ/{kuj}/QU/HULL_ATADAS/",
}

ADATTIPUS_LABELS = {
    "HULL_KELETKEZES": "Keletkezés",
    "HULL_ATVETEL": "Átvétel",
    "HULL_ATADAS": "Átadás",
}
DEFAULT_MAX_WORKERS = 4
REQUEST_JITTER_MIN = 0.18
REQUEST_JITTER_MAX = 0.65