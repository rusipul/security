from rules.r01_usb import detect_usb_attempts


def test_detects_all_usb_attempters(usb_blocked_df):
    result = detect_usb_attempts(usb_blocked_df)
    assert set(result["이름"]) == {"홍길동", "이순자"}


def test_counts_per_user(usb_blocked_df):
    result = detect_usb_attempts(usb_blocked_df)
    hong = result[result["이름"] == "홍길동"].iloc[0]
    assert hong["시도횟수"] == 2


def test_empty_df_returns_empty(usb_blocked_df):
    import pandas as pd
    empty = usb_blocked_df.iloc[0:0]
    result = detect_usb_attempts(empty)
    assert len(result) == 0


def test_result_columns(usb_blocked_df):
    result = detect_usb_attempts(usb_blocked_df)
    assert set(result.columns) >= {"이름", "부서", "시도횟수", "최근시도일시"}
