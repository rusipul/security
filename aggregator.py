from dataclasses import dataclass
import pandas as pd

_WEIGHTS = {
    "USB시도": 3,
    "설계파일": 2,
    "대용량ZIP": 1,
    "AI플랫폼": 2,
    "외부메일": 2,
    "메신저": 1,
}


@dataclass
class RuleResults:
    usb: pd.DataFrame
    design: pd.DataFrame
    large_zip: pd.DataFrame
    ai: pd.DataFrame
    email: pd.DataFrame
    messenger: pd.DataFrame


def _count_per_user(df: pd.DataFrame, count_col: str = None) -> dict[tuple[str, str], int]:
    """Count events per (이름, 부서). If count_col is given, sum that column instead of rows."""
    if df.empty or "이름" not in df.columns:
        return {}
    dept_col = df["부서"] if "부서" in df.columns else pd.Series([""] * len(df))
    counts: dict[tuple[str, str], int] = {}
    for i, (name, dept) in enumerate(zip(df["이름"], dept_col)):
        key = (str(name), str(dept))
        val = int(df[count_col].iloc[i] or 1) if count_col and count_col in df.columns else 1
        counts[key] = counts.get(key, 0) + val
    return counts


def build_risk_table(results: RuleResults, top_n: int = 10) -> pd.DataFrame:
    """Return top_n users sorted by composite risk score."""
    sources = {
        "USB시도":   _count_per_user(results.usb,      count_col="시도횟수"),
        "설계파일":  _count_per_user(results.design),
        "대용량ZIP": _count_per_user(results.large_zip),
        "AI플랫폼":  _count_per_user(results.ai),
        "외부메일":  _count_per_user(results.email,    count_col="시도횟수"),
        "메신저":    _count_per_user(results.messenger),
    }

    all_users: set[tuple[str, str]] = set()
    for counts in sources.values():
        all_users.update(counts.keys())

    rows = []
    for (name, dept) in all_users:
        row: dict = {"이름": name, "부서": dept}
        score = 0
        for col, counts in sources.items():
            n = counts.get((name, dept), 0)
            row[col] = n
            score += n * _WEIGHTS[col]
        row["위험점수"] = score
        rows.append(row)

    if not rows:
        cols = ["이름", "부서"] + list(_WEIGHTS.keys()) + ["위험점수"]
        return pd.DataFrame(columns=cols)

    df = pd.DataFrame(rows)
    df = df.sort_values("위험점수", ascending=False).head(top_n).reset_index(drop=True)
    df.insert(0, "순위", range(1, len(df) + 1))
    return df
