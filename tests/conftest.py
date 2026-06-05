import pandas as pd
import pytest


@pytest.fixture
def attach_df():
    """Minimal 파일첨부 내역_전체파일 DataFrame."""
    return pd.DataFrame({
        "이름":   ["홍길동", "홍길동",   "이순자",  "이순자",    "김철수"],
        "부서":   ["개발팀", "개발팀",   "영업팀",  "WeChat 첨부파일", "개발팀"],
        "분류":   ["첨부",   "첨부",    "첨부",    "첨부",       "첨부"],
        "프로세스": ["OUTLOOK.EXE", "OUTLOOK.EXE", "OUTLOOK.EXE", "WeChat.exe", "chrome.exe"],
        "파일명": ["schematic_rev3.SCH", "design.tar", "report.pdf",
                   "drawings.zip", "budget.xlsx"],
        "파일크기": ["120KB(약 0 ~ 1 MB)", "5120KB(약 5 ~ 6 MB)",
                    "300KB(약 0 ~ 1 MB)", "15360KB(약 15 ~ 16 MB)",
                    "200KB(약 0 ~ 1 MB)"],
        "시간": pd.to_datetime(["2026-04-01", "2026-04-02", "2026-04-03",
                                "2026-04-04", "2026-04-05"]),
    })


@pytest.fixture
def usb_blocked_df():
    """Minimal 매체별 내역_사용차단 내역 DataFrame."""
    return pd.DataFrame({
        "이름":    ["홍길동", "홍길동", "이순자"],
        "부서":    ["개발팀", "개발팀", "영업팀"],
        "분류":    ["차단",   "차단",  "차단"],
        "매체유형": ["FDD/USB 저장장치", "FDD/USB 저장장치", "FDD/USB 저장장치"],
        "시간":    pd.to_datetime(["2026-04-10", "2026-04-15", "2026-04-20"]),
    })


@pytest.fixture
def web_blocks_df():
    """Minimal 웹사이트 접속 차단내역 DataFrame."""
    return pd.DataFrame({
        "이름":    ["홍길동", "이순자",      "김철수",           "박민준"],
        "부서":    ["개발팀", "영업팀",      "개발팀",           "인사팀"],
        "카테고리": ["포털메일", "포털메일",  "AI서비스",        "AI서비스"],
        "사이트":  ["mail.naver.com", "mail.google.com",
                   "chat.openai.com", "gemini.google.com"],
        "제어정책": ["감시", "감시", "감시", "감시"],
        "시간":    pd.to_datetime(["2026-04-01", "2026-04-02",
                                  "2026-04-03", "2026-04-04"]),
    })
