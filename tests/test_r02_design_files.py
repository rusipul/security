from rules.r02_design_files import detect_design_file_attachments

KEYWORDS = ["DB", "SCH", "REV", "PCB", "BOM", "GERBER"]
EXTENSIONS = [".tar", ".gz", ".tg", ".brd", ".sch"]


def test_detects_extension_match(attach_df):
    result = detect_design_file_attachments(attach_df, KEYWORDS, EXTENSIONS)
    filenames = result["파일명"].tolist()
    assert "design.tar" in filenames


def test_detects_keyword_match(attach_df):
    result = detect_design_file_attachments(attach_df, KEYWORDS, EXTENSIONS)
    filenames = result["파일명"].tolist()
    assert "schematic_rev3.SCH" in filenames


def test_non_design_file_excluded(attach_df):
    result = detect_design_file_attachments(attach_df, KEYWORDS, EXTENSIONS)
    filenames = result["파일명"].tolist()
    assert "report.pdf" not in filenames
    assert "budget.xlsx" not in filenames


def test_result_has_match_reason(attach_df):
    result = detect_design_file_attachments(attach_df, KEYWORDS, EXTENSIONS)
    assert "탐지사유" in result.columns
    ext_row = result[result["파일명"] == "design.tar"].iloc[0]
    assert "확장자" in ext_row["탐지사유"]


def test_empty_returns_empty(attach_df):
    empty = attach_df.iloc[0:0]
    result = detect_design_file_attachments(empty, KEYWORDS, EXTENSIONS)
    assert len(result) == 0
