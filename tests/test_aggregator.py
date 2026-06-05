import pandas as pd
from aggregator import build_risk_table, RuleResults


def _make_results():
    return RuleResults(
        usb=pd.DataFrame({"이름": ["홍길동"], "부서": ["개발팀"], "시도횟수": [3]}),
        design=pd.DataFrame({"이름": ["홍길동", "이순자"], "부서": ["개발팀", "영업팀"],
                              "파일명": ["a.SCH", "b.tar"], "탐지사유": ["키워드", "확장자"]}),
        large_zip=pd.DataFrame(),
        ai=pd.DataFrame({"이름": ["홍길동"], "부서": ["개발팀"],
                          "사이트": ["chat.openai.com"], "접속횟수": [2], "업로드여부": ["접속"]}),
        email=pd.DataFrame(),
        messenger=pd.DataFrame(),
    )


def test_top_user_is_highest_risk():
    results = _make_results()
    table = build_risk_table(results, top_n=5)
    assert table.iloc[0]["이름"] == "홍길동"


def test_risk_score_uses_weights():
    results = _make_results()
    table = build_risk_table(results, top_n=5)
    hong = table[table["이름"] == "홍길동"].iloc[0]
    # USB=3*3=9, design=1*2=2, AI=1*2=2 → total=13
    assert hong["위험점수"] == 13


def test_limited_by_top_n():
    results = _make_results()
    table = build_risk_table(results, top_n=1)
    assert len(table) == 1


def test_all_rule_columns_present():
    results = _make_results()
    table = build_risk_table(results, top_n=5)
    for col in ["USB시도", "설계파일", "대용량ZIP", "AI플랫폼", "외부메일", "메신저"]:
        assert col in table.columns
