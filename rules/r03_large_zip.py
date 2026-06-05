import re
import pandas as pd


def parse_size_kb(size_str) -> int:
    if not size_str:
        return 0
    m = re.search(r"(\d+)KB", str(size_str))
    return int(m.group(1)) if m else 0


def detect_large_zips(attach: pd.DataFrame, threshold_mb: float = 10) -> pd.DataFrame:
    """Detect .zip attachments >= threshold_mb in size."""
    if attach.empty:
        return attach.iloc[0:0].copy()
    threshold_kb = int(threshold_mb * 1024)
    is_zip = attach["파일명"].fillna("").str.lower().str.endswith(".zip")
    size_kb = attach["파일크기"].apply(parse_size_kb)
    matched = attach[is_zip & (size_kb >= threshold_kb)].copy()
    matched["파일크기_KB"] = size_kb[matched.index]
    return matched.reset_index(drop=True)
