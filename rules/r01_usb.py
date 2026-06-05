import pandas as pd


def detect_usb_attempts(usb_blocked: pd.DataFrame) -> pd.DataFrame:
    """Return one row per user who attempted USB storage use."""
    usb = usb_blocked[
        usb_blocked["매체유형"].str.contains("USB", na=False, case=False)
    ].copy()
    if usb.empty:
        return pd.DataFrame(columns=["이름", "부서", "시도횟수", "최근시도일시"])
    grp = usb.groupby(["이름", "부서"], as_index=False).agg(
        시도횟수=("이름", "count"),
        최근시도일시=("시간", "max"),
    )
    return grp.sort_values("시도횟수", ascending=False).reset_index(drop=True)
