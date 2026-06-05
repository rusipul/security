from rules.r03_large_zip import detect_large_zips, parse_size_kb


def test_parse_size_kb_standard():
    assert parse_size_kb("15360KB(약 15 ~ 16 MB)") == 15360


def test_parse_size_kb_small():
    assert parse_size_kb("608KB(약 0 ~ 1 MB)") == 608


def test_parse_size_kb_none():
    assert parse_size_kb(None) == 0


def test_detects_large_zip(attach_df):
    # drawings.zip is 15360 KB (>= 10240 KB = 10 MB)
    result = detect_large_zips(attach_df, threshold_mb=10)
    assert "drawings.zip" in result["파일명"].tolist()


def test_excludes_small_zip(attach_df):
    import pandas as pd
    extra = attach_df.copy()
    extra.loc[len(extra)] = {
        "이름": "테스트", "부서": "개발팀", "분류": "첨부",
        "프로세스": "OUTLOOK.EXE", "파일명": "small.zip",
        "파일크기": "500KB(약 0 ~ 1 MB)", "시간": pd.Timestamp("2026-04-06"),
    }
    result = detect_large_zips(extra, threshold_mb=10)
    assert "small.zip" not in result["파일명"].tolist()


def test_non_zip_excluded(attach_df):
    result = detect_large_zips(attach_df, threshold_mb=10)
    filenames = result["파일명"].tolist()
    assert "design.tar" not in filenames
