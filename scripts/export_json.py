"""Exportação JSON versionada. Dados originais e derivadas permanecem distintos."""
from __future__ import annotations
import json
from pathlib import Path
from typing import Any


def export_json(payload: dict[str,Any], output: str | Path) -> Path:
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open('w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=str)
        f.write('\n')
    return output
