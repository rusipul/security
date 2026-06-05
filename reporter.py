from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

_RED_FILL = PatternFill("solid", fgColor="FFCCCC")
_ORANGE_FILL = PatternFill("solid", fgColor="FFE5CC")
_HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_BOLD = Font(bold=True)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_THIN = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


def _auto_width(ws):
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 50)


def _write_df_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame, fill=None):
    ws = wb.create_sheet(sheet_name)
    if df.empty:
        ws.append(["데이터 없음"])
        return
    headers = list(df.columns)
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = _CENTER
    for row_data in df.itertuples(index=False):
        ws.append(list(row_data))
    if fill:
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.fill = fill
    _auto_width(ws)


def write_report(
    output_path: str,
    period: str,
    counts: dict[str, int],
    risk_table: pd.DataFrame,
    ai_summary: str,
    usb_df: pd.DataFrame,
    design_df: pd.DataFrame,
    large_zip_df: pd.DataFrame,
    ai_df: pd.DataFrame,
    email_df: pd.DataFrame,
    messenger_df: pd.DataFrame,
) -> str:
    """Write the full Excel report and return the saved path."""
    wb = Workbook()
    ws = wb.active
    ws.title = "보안위협 요약"

    # Title
    ws.merge_cells("A1:H1")
    ws["A1"] = f"DLP 보안 위협 분석 보고서 | {period}"
    ws["A1"].font = Font(size=16, bold=True)
    ws["A1"].alignment = _CENTER
    ws["A2"] = f"분석일시: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  총 탐지건수: {sum(counts.values())}건"
    ws["A2"].alignment = _CENTER
    ws.merge_cells("A2:H2")
    ws.row_dimensions[1].height = 35
    ws.row_dimensions[2].height = 22

    # Detection summary
    ws["A4"] = "탐지 유형"
    ws["B4"] = "건수"
    ws["C4"] = "위험도"
    for cell in [ws["A4"], ws["B4"], ws["C4"]]:
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = _CENTER

    risk_map = {
        "USB 시도":       ("🔴 고위험", _RED_FILL),
        "설계파일 첨부":   ("🔴 고위험", _RED_FILL),
        "대용량 ZIP 첨부": ("🟠 주의",   _ORANGE_FILL),
        "AI 플랫폼 접속":  ("🔴/🟠",    _ORANGE_FILL),
        "외부메일 시도":   ("🔴 고위험", _RED_FILL),
        "메신저 파일전송": ("🟠 주의",   _ORANGE_FILL),
    }
    count_keys = ["USB 시도", "설계파일 첨부", "대용량 ZIP 첨부",
                  "AI 플랫폼 접속", "외부메일 시도", "메신저 파일전송"]
    for i, key in enumerate(count_keys, start=5):
        label, (risk_txt, fill) = key, risk_map[key]
        ws.cell(i, 1, label).fill = fill
        ws.cell(i, 2, counts.get(key, 0)).alignment = _CENTER
        ws.cell(i, 3, risk_txt).alignment = _CENTER

    # Risk table header
    row_start = 13
    ws.cell(row_start, 1, "상위 위험 사용자").font = _BOLD
    ws.merge_cells(f"A{row_start}:H{row_start}")
    row_start += 1
    if not risk_table.empty:
        headers = list(risk_table.columns)
        for j, h in enumerate(headers, 1):
            c = ws.cell(row_start, j, h)
            c.fill = _HEADER_FILL
            c.font = _HEADER_FONT
            c.alignment = _CENTER
        for r_idx, row_data in enumerate(risk_table.itertuples(index=False), row_start + 1):
            for j, val in enumerate(row_data, 1):
                ws.cell(r_idx, j, val).alignment = _CENTER

    # AI summary
    ai_row = row_start + len(risk_table) + 3
    ws.cell(ai_row, 1, "🤖 AI 위협 요약 (Claude 분석)").font = _BOLD
    ws.merge_cells(f"A{ai_row}:H{ai_row}")
    ai_row += 1
    ws.cell(ai_row, 1, ai_summary)
    ws.merge_cells(f"A{ai_row}:H{ai_row + 8}")
    ws.cell(ai_row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[ai_row].height = 120

    _auto_width(ws)

    # Detail sheets
    _write_df_sheet(wb, "USB시도 상세",      usb_df,      _RED_FILL)
    _write_df_sheet(wb, "설계파일 상세",     design_df,   _RED_FILL)
    _write_df_sheet(wb, "대용량ZIP 상세",    large_zip_df, _ORANGE_FILL)
    _write_df_sheet(wb, "AI플랫폼 상세",     ai_df,       _ORANGE_FILL)
    _write_df_sheet(wb, "외부메일 상세",     email_df,    _RED_FILL)
    _write_df_sheet(wb, "메신저 상세",       messenger_df, _ORANGE_FILL)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
