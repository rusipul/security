import pandas as pd
from pathlib import Path

# (sheet_index, skip_rows, col_rename_map)
_SHEET_CONFIGS = {
    "usb_all":    (1,  3, {0:"이름", 1:"부서", 2:"시스템유형", 3:"분류", 4:"매체유형",
                            10:"장치명", 11:"파일명", 12:"파일경로", 13:"파일크기", 26:"시간"}),
    "usb_blocked":(3,  3, {0:"이름", 1:"부서", 2:"시스템유형", 3:"분류", 4:"매체유형", 26:"시간"}),
    "attach_all": (6,  3, {0:"이름", 1:"부서", 2:"시스템유형", 3:"분류", 4:"프로세스",
                            5:"제품명", 6:"회사명", 7:"파일명", 8:"파일경로", 9:"파일크기", 22:"시간"}),
    "attach_ext": (7,  3, {0:"이름", 1:"부서", 2:"시스템유형", 3:"분류", 4:"프로세스",
                            5:"제품명", 6:"회사명", 7:"파일명", 8:"파일경로", 9:"파일크기", 22:"시간"}),
    "sw_blocks":  (16, 2, {0:"이름", 1:"부서", 2:"시스템유형", 3:"카테고리",
                            4:"프로세스", 5:"제어정책", 6:"시간"}),
    "proc_ctrl":  (17, 2, {0:"이름", 1:"부서", 2:"시스템유형", 3:"서비스명",
                            4:"프로세스", 5:"시간"}),
    "web_blocks": (18, 2, {0:"이름", 1:"부서", 2:"시스템유형", 3:"카테고리",
                            4:"사이트", 5:"제어정책", 6:"시간"}),
}


def load_workbook(path: str) -> dict[str, pd.DataFrame]:
    """Read DLP Excel report and return normalized DataFrames keyed by short name."""
    wb_path = Path(path)
    result = {}
    for key, (sheet_idx, skip, col_map) in _SHEET_CONFIGS.items():
        df = pd.read_excel(
            wb_path,
            sheet_name=sheet_idx,
            header=None,
            skiprows=skip,
            engine="openpyxl",
        )
        # Keep only mapped columns; rename to semantic names
        keep = sorted(col_map.keys())
        # Some reports may have fewer columns — only keep columns that exist
        keep = [c for c in keep if c < len(df.columns)]
        df = df.iloc[:, keep].copy()
        df.columns = [col_map[c] for c in keep]
        # Drop rows where 이름 is null (empty rows at end of sheet)
        if "이름" in df.columns:
            df = df[df["이름"].notna()].reset_index(drop=True)
        result[key] = df
    return result
