import pytest
from loader import load_workbook

SAMPLE_PATH = "오피스키퍼_정보유출방지_다믈파워반도체_202604.xlsx"

def test_load_returns_required_keys():
    sheets = load_workbook(SAMPLE_PATH)
    required = {"usb_all", "usb_blocked", "attach_all", "attach_ext",
                "sw_blocks", "proc_ctrl", "web_blocks"}
    assert required.issubset(sheets.keys())

def test_attach_all_has_expected_columns():
    sheets = load_workbook(SAMPLE_PATH)
    df = sheets["attach_all"]
    assert "이름" in df.columns
    assert "파일명" in df.columns
    assert "파일크기" in df.columns

def test_attach_all_row_count():
    sheets = load_workbook(SAMPLE_PATH)
    # sheet [6] has 32533 rows; subtract 3 header rows
    assert len(sheets["attach_all"]) > 30000

def test_web_blocks_has_site_column():
    sheets = load_workbook(SAMPLE_PATH)
    df = sheets["web_blocks"]
    assert "사이트" in df.columns
    assert len(df) > 400
