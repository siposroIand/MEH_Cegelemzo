"""Synthetic, offline benchmark: python -m tools.benchmark."""
import tempfile
from pathlib import Path
from time import perf_counter

import pandas as pd

from core.constants import PRODUCT_GROUPS_JSON
from services.aggregation_service import aggregate_data
from services.category_service import CategoryService
from services.filter_service import Filters, filter_frame
from services.persistence_service import PersistenceService
from tests.fixtures import dataset


def main():
    source, raw, missing = dataset(120)
    raw = pd.concat([raw] * 834, ignore_index=True).iloc[:100_000]
    categories = CategoryService(PRODUCT_GROUPS_JSON)
    start = perf_counter()
    aggregates = aggregate_data(raw, source, categories)
    aggregated = perf_counter()
    selected = Filters(years=(2025,), categories=("Vas",))
    for _ in range(10):
        filter_frame(raw, selected, categories)
    filtered = perf_counter()
    with tempfile.TemporaryDirectory() as directory:
        persistence = PersistenceService(Path(directory) / "benchmark.sqlite")
        try:
            before_save = perf_counter()
            sid = persistence.save_snapshot(source, raw, missing)
            saved = perf_counter()
            loaded = persistence.load_snapshot(sid)
            restored = perf_counter()
            assert len(loaded[1]) == 100_000
            assert abs(aggregates.base_unique.OSSZES_MENNYISEG.sum() - raw.MENNYISEG.sum()) < .01
            print(f"100,000 rows: aggregate={aggregated-start:.3f}s, filter_avg={(filtered-aggregated)/10:.3f}s, "
                  f"save={saved-before_save:.3f}s, restore={restored-saved:.3f}s")
        finally:
            persistence.close()


if __name__ == "__main__":
    main()
