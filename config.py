import json
from pathlib import Path

_DEFAULTS = {"api_key": "", "top_n": 10, "report_dir": "reports"}


def load_keywords(path: str = "keywords.json") -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def load_settings(path: str = "config.json") -> dict:
    p = Path(path)
    if not p.exists():
        return dict(_DEFAULTS)
    with open(p, encoding="utf-8") as f:
        data = json.load(f)
    return {**_DEFAULTS, **data}


def save_settings(settings: dict, path: str = "config.json") -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)
