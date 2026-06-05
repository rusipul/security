from rules.r05_external_email import detect_external_email_attempts

MAIL_DOMAINS = ["mail.naver.com", "mail.google.com", "mail.daum.net",
                "mail.yahoo.com", "mail.kakao.com"]


def test_detects_naver_mail(web_blocks_df):
    result = detect_external_email_attempts(web_blocks_df, MAIL_DOMAINS)
    sites = result["사이트"].tolist()
    assert "mail.naver.com" in sites


def test_detects_gmail(web_blocks_df):
    result = detect_external_email_attempts(web_blocks_df, MAIL_DOMAINS)
    sites = result["사이트"].tolist()
    assert "mail.google.com" in sites


def test_ai_sites_excluded(web_blocks_df):
    result = detect_external_email_attempts(web_blocks_df, MAIL_DOMAINS)
    sites = result["사이트"].tolist()
    assert "chat.openai.com" not in sites


def test_result_columns(web_blocks_df):
    result = detect_external_email_attempts(web_blocks_df, MAIL_DOMAINS)
    assert set(result.columns) >= {"이름", "부서", "사이트", "시도횟수", "최근시도일시"}


def test_empty_returns_empty():
    import pandas as pd
    result = detect_external_email_attempts(pd.DataFrame(), MAIL_DOMAINS)
    assert len(result) == 0
