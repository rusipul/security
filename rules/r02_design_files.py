import pandas as pd


def detect_design_file_attachments(
    attach: pd.DataFrame,
    design_keywords: list[str],
    design_extensions: list[str],
) -> pd.DataFrame:
    """Detect attachments matching design extensions or keywords."""
    if attach.empty:
        return pd.DataFrame(columns=list(attach.columns) + ["탐지사유"])

    fname = attach["파일명"].fillna("").str.lower()

    ext_lower = [e.lower() for e in design_extensions]
    kw_lower = [k.lower() for k in design_keywords]

    ext_mask = fname.apply(lambda f: any(f.endswith(e) for e in ext_lower))
    kw_mask = fname.apply(lambda f: any(k in f for k in kw_lower))

    matched = attach[ext_mask | kw_mask].copy()

    def reason(f: str) -> str:
        f = f.lower()
        reasons = []
        hit_ext = [e for e in ext_lower if f.endswith(e)]
        hit_kw = [k for k in kw_lower if k in f]
        if hit_ext:
            reasons.append(f"확장자({hit_ext[0]})")
        if hit_kw:
            reasons.append(f"키워드({','.join(hit_kw).upper()})")
        return ", ".join(reasons)

    matched["탐지사유"] = matched["파일명"].fillna("").apply(reason)
    return matched.reset_index(drop=True)
