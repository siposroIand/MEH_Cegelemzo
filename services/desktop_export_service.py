import os
from pathlib import Path
import tempfile


def atomic_export(path, writer, *args, **kwargs):
    """Keep an existing report intact if writing its replacement fails."""
    destination = Path(path)
    descriptor, temporary = tempfile.mkstemp(
        prefix=f".{destination.stem}-", suffix=destination.suffix,
        dir=destination.parent,
    )
    os.close(descriptor)
    try:
        writer(temporary, *args, **kwargs)
        os.replace(temporary, destination)
    finally:
        Path(temporary).unlink(missing_ok=True)
    return str(destination)
