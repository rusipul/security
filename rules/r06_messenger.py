import pandas as pd

_MESSENGER_PATTERNS = [
    ("WeChat",    ["wechat"]),
    ("KakaoTalk", ["kakaotalk", "kakao"]),
]


def _identify_platform(dept: str, proc: str) -> str | None:
    text = (str(dept) + " " + str(proc)).lower()
    for name, patterns in _MESSENGER_PATTERNS:
        if any(p in text for p in patterns):
            return name
    return None


def detect_messenger_attachments(attach: pd.DataFrame) -> pd.DataFrame:
    """Detect file attachments sent via WeChat or KakaoTalk."""
    if attach.empty or "부서" not in attach.columns:
        return pd.DataFrame(columns=["이름", "부서", "플랫폼", "파일명", "파일크기", "시간"])

    dept_col = attach["부서"].fillna("")
    proc_col = attach.get("프로세스", pd.Series([""] * len(attach))).fillna("")

    platforms = [_identify_platform(d, p) for d, p in zip(dept_col, proc_col)]
    mask = [p is not None for p in platforms]
    matched = attach[mask].copy()
    matched["플랫폼"] = [p for p, m in zip(platforms, mask) if m]
    return matched.reset_index(drop=True)
