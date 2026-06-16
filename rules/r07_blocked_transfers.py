import pandas as pd


def detect_usb_file_blocks(usb_all: pd.DataFrame) -> pd.DataFrame:
    """USB 파일복사 차단 내역 — usb_all에서 분류=='차단'인 행."""
    if usb_all.empty or "분류" not in usb_all.columns:
        return pd.DataFrame(columns=["이름", "부서", "파일명", "파일크기", "시간"])
    blocked = usb_all[usb_all["분류"] == "차단"].copy()
    return blocked.reset_index(drop=True)


def detect_attachment_blocks(attach_all: pd.DataFrame) -> pd.DataFrame:
    """파일전송 차단 내역 — attach_all에서 분류=='차단'인 행 (웹/이메일/메신저 경유 차단)."""
    if attach_all.empty or "분류" not in attach_all.columns:
        return pd.DataFrame(columns=["이름", "부서", "프로세스", "파일명", "파일크기", "시간"])
    blocked = attach_all[attach_all["분류"] == "차단"].copy()
    return blocked.reset_index(drop=True)


def summarize_web_policy(web_blocks: pd.DataFrame) -> dict[str, int]:
    """웹사이트 접속 제어정책별 건수 — 차단/경고 집계."""
    if web_blocks.empty or "제어정책" not in web_blocks.columns:
        return {"차단": 0, "경고": 0}
    vc = web_blocks["제어정책"].value_counts()
    return {
        "차단": int(vc.get("차단", 0)),
        "경고": int(vc.get("경고", 0)),
    }
