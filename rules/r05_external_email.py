import pandas as pd


def detect_external_email_attempts(
    web_blocks: pd.DataFrame,
    mail_domains: list[str],
) -> pd.DataFrame:
    """Detect access attempts to external mail services."""
    if web_blocks.empty or "사이트" not in web_blocks.columns:
        return pd.DataFrame(columns=["이름", "부서", "사이트", "시도횟수", "최근시도일시"])

    mail_lower = [d.lower() for d in mail_domains]
    site_col = web_blocks["사이트"].fillna("").str.lower()
    is_mail = site_col.apply(lambda s: any(d in s for d in mail_lower))
    mail_df = web_blocks[is_mail].copy()

    if mail_df.empty:
        return pd.DataFrame(columns=["이름", "부서", "사이트", "시도횟수", "최근시도일시"])

    grp = mail_df.groupby(["이름", "부서", "사이트"], as_index=False).agg(
        시도횟수=("이름", "count"),
        최근시도일시=("시간", "max"),
    )
    return grp.sort_values("시도횟수", ascending=False).reset_index(drop=True)
