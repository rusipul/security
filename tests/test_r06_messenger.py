import pandas as pd
from rules.r06_messenger import detect_messenger_attachments


def test_detects_wechat_by_dept(attach_df):
    result = detect_messenger_attachments(attach_df)
    assert "이순자" in result["이름"].tolist()  # dept contains "WeChat"


def test_non_messenger_excluded(attach_df):
    result = detect_messenger_attachments(attach_df)
    # 홍길동 uses OUTLOOK.EXE with no WeChat dept → should NOT appear
    assert "홍길동" not in result["이름"].tolist()


def test_result_has_platform_column(attach_df):
    result = detect_messenger_attachments(attach_df)
    assert "플랫폼" in result.columns


def test_platform_identified_as_wechat(attach_df):
    result = detect_messenger_attachments(attach_df)
    row = result[result["이름"] == "이순자"].iloc[0]
    assert "WeChat" in row["플랫폼"] or "wechat" in row["플랫폼"].lower()


def test_empty_returns_empty():
    result = detect_messenger_attachments(pd.DataFrame())
    assert len(result) == 0
