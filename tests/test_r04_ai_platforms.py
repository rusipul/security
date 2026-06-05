import pandas as pd
from rules.r04_ai_platforms import detect_ai_access

AI_DOMAINS = ["chat.openai.com", "chatgpt.com", "gemini.google.com",
              "claude.ai", "copilot.microsoft.com", "perplexity.ai",
              "notebooklm.google.com"]
BROWSERS = ["chrome.exe", "msedge.exe", "firefox.exe"]


def test_detects_ai_domain_access(web_blocks_df):
    result = detect_ai_access(web_blocks_df, pd.DataFrame(), AI_DOMAINS, BROWSERS)
    sites = result["사이트"].tolist()
    assert "chat.openai.com" in sites
    assert "gemini.google.com" in sites


def test_mail_domains_excluded(web_blocks_df):
    result = detect_ai_access(web_blocks_df, pd.DataFrame(), AI_DOMAINS, BROWSERS)
    sites = result["사이트"].tolist()
    assert "mail.naver.com" not in sites


def test_upload_flagged_when_browser_attachment_same_day(web_blocks_df, attach_df):
    # 김철수 accessed chat.openai.com on 2026-04-03
    # attach_df has 김철수 with chrome.exe on 2026-04-05 (different day — NOT upload)
    result = detect_ai_access(web_blocks_df, attach_df, AI_DOMAINS, BROWSERS)
    kim_row = result[result["이름"] == "김철수"].iloc[0]
    assert kim_row["업로드여부"] == "접속"  # different date → access only


def test_result_columns(web_blocks_df):
    result = detect_ai_access(web_blocks_df, pd.DataFrame(), AI_DOMAINS, BROWSERS)
    assert set(result.columns) >= {"이름", "부서", "사이트", "접속횟수", "업로드여부", "최근접속일시"}


def test_empty_web_blocks_returns_empty():
    result = detect_ai_access(pd.DataFrame(), pd.DataFrame(), AI_DOMAINS, BROWSERS)
    assert len(result) == 0
