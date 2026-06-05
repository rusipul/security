import pandas as pd


def detect_ai_access(
    web_blocks: pd.DataFrame,
    attach_all: pd.DataFrame,
    ai_domains: list[str],
    browser_processes: list[str],
) -> pd.DataFrame:
    """Detect AI platform access; flag probable uploads."""
    if web_blocks.empty or "사이트" not in web_blocks.columns:
        return pd.DataFrame(columns=["이름", "부서", "사이트", "접속횟수", "업로드여부", "최근접속일시"])

    ai_lower = [d.lower() for d in ai_domains]
    site_col = web_blocks["사이트"].fillna("").str.lower()
    is_ai = site_col.apply(lambda s: any(d in s for d in ai_lower))
    ai_df = web_blocks[is_ai].copy()

    if ai_df.empty:
        return pd.DataFrame(columns=["이름", "부서", "사이트", "접속횟수", "업로드여부", "최근접속일시"])

    # Build set of (이름, date) with browser attachment
    browser_dates: set[tuple[str, str]] = set()
    if not attach_all.empty and "프로세스" in attach_all.columns:
        br_lower = [b.lower() for b in browser_processes]
        is_browser = attach_all["프로세스"].fillna("").str.lower().apply(
            lambda p: any(b in p for b in br_lower)
        )
        br_att = attach_all[is_browser].copy()
        if not br_att.empty and "시간" in br_att.columns:
            br_att["날짜"] = pd.to_datetime(br_att["시간"]).dt.date
            for _, row in br_att.iterrows():
                browser_dates.add((row["이름"], str(row["날짜"])))

    ai_df["날짜"] = pd.to_datetime(ai_df["시간"]).dt.date.astype(str)

    records = []
    for (name, dept, site), grp in ai_df.groupby(["이름", "부서", "사이트"]):
        has_upload = any((name, d) in browser_dates for d in grp["날짜"])
        records.append({
            "이름": name,
            "부서": dept,
            "사이트": site,
            "접속횟수": len(grp),
            "업로드여부": "업로드(추정)" if has_upload else "접속",
            "최근접속일시": grp["시간"].max(),
        })

    return pd.DataFrame(records).sort_values("접속횟수", ascending=False).reset_index(drop=True)
