from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 색상 팔레트 ──────────────────────────────────────────────
_NAVY       = "1F4E79"
_RED_BG     = "FFCCCC"
_ORANGE_BG  = "FFE5CC"
_YELLOW_BG  = "FFF2CC"
_GREEN_BG   = "E2EFDA"
_NAVY_BG    = "D6E4F0"
_GRAY_BG    = "F2F2F2"
_WHITE      = "FFFFFF"

_F_WHITE    = Font(color=_WHITE, bold=True, size=10)
_F_NAVY     = Font(color=_NAVY, bold=True, size=10)
_F_TITLE    = Font(color=_NAVY, bold=True, size=16)
_F_SECTION  = Font(color=_WHITE, bold=True, size=11)
_F_BODY     = Font(size=10)
_F_BOLD     = Font(bold=True, size=10)
_F_SMALL    = Font(size=9, color="595959")

_CENTER     = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT       = Alignment(horizontal="left",   vertical="center", wrap_text=True)
_RIGHT      = Alignment(horizontal="right",  vertical="center")

_THIN_SIDE  = Side(style="thin",   color="BFBFBF")
_THICK_SIDE = Side(style="medium", color=_NAVY)
_BORDER     = Border(left=_THIN_SIDE, right=_THIN_SIDE,
                     top=_THIN_SIDE,  bottom=_THIN_SIDE)
_BORDER_THICK_TOP = Border(left=_THIN_SIDE, right=_THIN_SIDE,
                            top=_THICK_SIDE, bottom=_THIN_SIDE)


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def _cell(ws, row, col, value="", font=None, fill=None, align=None, border=None):
    c = ws.cell(row, col, value)
    if font:   c.font = font
    if fill:   c.fill = fill
    if align:  c.alignment = align
    if border: c.border = border
    return c


def _hdr_row(ws, row, cols, labels, fill_color=_NAVY):
    for col, label in zip(cols, labels):
        _cell(ws, row, col, label,
              font=_F_WHITE, fill=_fill(fill_color),
              align=_CENTER, border=_BORDER)


def _section_bar(ws, row, col_start, col_end, text, fill_color=_NAVY):
    ws.merge_cells(start_row=row, start_column=col_start,
                   end_row=row, end_column=col_end)
    _cell(ws, row, col_start, text,
          font=_F_SECTION, fill=_fill(fill_color), align=_LEFT)
    ws.row_dimensions[row].height = 20


def _auto_width(ws):
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 50)


# ── 상세 시트 ────────────────────────────────────────────────

def _write_df_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame,
                    fill_color: str = _GRAY_BG):
    ws = wb.create_sheet(sheet_name)
    if df.empty:
        ws.append(["데이터 없음"])
        return
    headers = list(df.columns)
    for j, h in enumerate(headers, 1):
        _cell(ws, 1, j, h, font=_F_WHITE, fill=_fill(_NAVY), align=_CENTER)
    row_fill = _fill(fill_color)
    for r_idx, row_data in enumerate(df.itertuples(index=False), 2):
        for j, val in enumerate(row_data, 1):
            c = ws.cell(r_idx, j, val)
            c.fill = row_fill
            c.alignment = _CENTER
            c.border = _BORDER
    _auto_width(ws)


# ── 첫 장: 경영진 보고용 1페이지 요약 ────────────────────────

def _write_executive_sheet(
    ws,
    period: str,
    filename: str,
    threat_counts: dict,
    block_counts: dict,
    risk_table: pd.DataFrame,
    ai_summary: str,
):
    # 페이지 설정 (A4 세로, 1페이지)
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9   # A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToHeight = 1
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitPage = True
    ws.page_margins.left   = 0.5
    ws.page_margins.right  = 0.5
    ws.page_margins.top    = 0.6
    ws.page_margins.bottom = 0.6

    # 컬럼 폭 (A-I, 9열)
    col_widths = [22, 9, 12, 3, 22, 9, 12, 3, 26]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    r = 1

    # ── 제목 블록 ────────────────────────────
    ws.merge_cells(f"A{r}:I{r}")
    _cell(ws, r, 1, f"정보보안 현황 보고서  |  {period}",
          font=_F_TITLE, fill=_fill(_NAVY_BG), align=_CENTER)
    ws.row_dimensions[r].height = 38
    r += 1

    ws.merge_cells(f"A{r}:I{r}")
    now_str = datetime.now().strftime("%Y년 %m월 %d일  %H:%M")
    _cell(ws, r, 1, f"분석일시: {now_str}     원본파일: {filename}",
          font=_F_SMALL, fill=_fill(_GRAY_BG), align=_CENTER)
    ws.row_dimensions[r].height = 16
    r += 1

    ws.row_dimensions[r].height = 6   # 구분선
    r += 1

    # ── 섹션 1·2: 차단현황(좌) + 위협탐지(우) 나란히 ────────
    _section_bar(ws, r, 1, 3, "  ▶  DLP 차단 현황")
    _section_bar(ws, r, 5, 7, "  ▶  보안 위협 탐지 현황")
    ws.merge_cells(f"I{r}:I{r}")
    r += 1

    # 소헤더
    _hdr_row(ws, r, [1, 2, 3], ["차단 유형", "건수", "위험도"])
    _hdr_row(ws, r, [5, 6, 7], ["탐지 유형", "건수", "위험도"])
    ws.row_dimensions[r].height = 18
    r += 1

    block_rows = [
        ("USB 파일복사 차단",  "USB시도 차단",  _RED_BG,    "🔴 고위험"),
        ("파일전송 차단",      "파일전송 차단",  _RED_BG,    "🔴 고위험"),
        ("웹사이트 접속 차단", "웹사이트 차단",  _ORANGE_BG, "🟠 주의"),
        ("웹사이트 접속 경고", "웹사이트 경고",  _YELLOW_BG, "⚠️ 경고"),
    ]
    threat_rows = [
        ("USB 시도",       _RED_BG,    "🔴 고위험"),
        ("설계파일 첨부",   _RED_BG,    "🔴 고위험"),
        ("대용량 ZIP 첨부", _ORANGE_BG, "🟠 주의"),
        ("AI 플랫폼 접속",  _ORANGE_BG, "🟠 주의"),
        ("외부메일 시도",   _RED_BG,    "🔴 고위험"),
        ("메신저 파일전송", _YELLOW_BG, "🟠 주의"),
    ]

    max_rows = max(len(block_rows), len(threat_rows))
    for i in range(max_rows):
        ws.row_dimensions[r].height = 17
        if i < len(block_rows):
            label, key, bg, risk = block_rows[i]
            cnt = block_counts.get(key, block_counts.get(label, 0))
            _cell(ws, r, 1, label, font=_F_BODY, fill=_fill(bg), align=_LEFT,   border=_BORDER)
            _cell(ws, r, 2, f"{cnt:,}건", font=_F_BOLD, fill=_fill(bg), align=_CENTER, border=_BORDER)
            _cell(ws, r, 3, risk,  font=_F_BODY, fill=_fill(bg), align=_CENTER, border=_BORDER)
        if i < len(threat_rows):
            label, bg, risk = threat_rows[i]
            cnt = threat_counts.get(label, 0)
            _cell(ws, r, 5, label, font=_F_BODY, fill=_fill(bg), align=_LEFT,   border=_BORDER)
            _cell(ws, r, 6, f"{cnt:,}건", font=_F_BOLD, fill=_fill(bg), align=_CENTER, border=_BORDER)
            _cell(ws, r, 7, risk,  font=_F_BODY, fill=_fill(bg), align=_CENTER, border=_BORDER)
        r += 1

    ws.row_dimensions[r].height = 6
    r += 1

    # ── 섹션 3: 상위 위험 사용자 ───────────────────────────
    _section_bar(ws, r, 1, 9, "  ▶  상위 위험 사용자 현황")
    r += 1

    if not risk_table.empty:
        headers = list(risk_table.columns)
        for j, h in enumerate(headers, 1):
            _cell(ws, r, j, h, font=_F_WHITE, fill=_fill(_NAVY), align=_CENTER, border=_BORDER)
        ws.row_dimensions[r].height = 18
        r += 1
        for row_data in risk_table.itertuples(index=False):
            ws.row_dimensions[r].height = 16
            score = row_data[-1]   # 위험점수는 마지막 열
            row_bg = _RED_BG if score >= 10 else _ORANGE_BG if score >= 5 else _GREEN_BG
            for j, val in enumerate(row_data, 1):
                _cell(ws, r, j, val, font=_F_BODY,
                      fill=_fill(row_bg), align=_CENTER, border=_BORDER)
            r += 1
    else:
        ws.merge_cells(f"A{r}:I{r}")
        _cell(ws, r, 1, "탐지된 위험 사용자 없음", font=_F_BODY, align=_CENTER)
        r += 1

    ws.row_dimensions[r].height = 6
    r += 1

    # ── 섹션 4: AI 요약 ────────────────────────────────────
    _section_bar(ws, r, 1, 9, "  ▶  AI 보안 위협 분석  (Claude Sonnet)")
    r += 1

    ai_row = r
    ws.merge_cells(f"A{ai_row}:I{ai_row + 7}")
    c = ws.cell(ai_row, 1, ai_summary)
    c.font = Font(size=9)
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    c.fill = _fill(_GRAY_BG)
    ws.row_dimensions[ai_row].height = 110


# ── 메인 write_report ────────────────────────────────────────

def write_report(
    output_path: str,
    period: str,
    filename: str,
    threat_counts: dict[str, int],
    block_counts: dict[str, int],
    risk_table: pd.DataFrame,
    ai_summary: str,
    usb_df: pd.DataFrame,
    design_df: pd.DataFrame,
    large_zip_df: pd.DataFrame,
    ai_df: pd.DataFrame,
    email_df: pd.DataFrame,
    messenger_df: pd.DataFrame,
    usb_block_df: pd.DataFrame,
    attach_block_df: pd.DataFrame,
) -> str:
    """Write full Excel report and return saved path."""
    wb = Workbook()
    ws = wb.active
    ws.title = "보안위협 요약"

    _write_executive_sheet(
        ws, period, filename,
        threat_counts, block_counts,
        risk_table, ai_summary,
    )

    # 상세 시트 (위협 탐지)
    _write_df_sheet(wb, "USB시도 상세",       usb_df,        _RED_BG)
    _write_df_sheet(wb, "설계파일 상세",      design_df,     _RED_BG)
    _write_df_sheet(wb, "대용량ZIP 상세",     large_zip_df,  _ORANGE_BG)
    _write_df_sheet(wb, "AI플랫폼 상세",      ai_df,         _ORANGE_BG)
    _write_df_sheet(wb, "외부메일 상세",      email_df,      _RED_BG)
    _write_df_sheet(wb, "메신저 상세",        messenger_df,  _YELLOW_BG)

    # 상세 시트 (차단 내역)
    _write_df_sheet(wb, "USB복사차단 상세",   usb_block_df,  _RED_BG)
    _write_df_sheet(wb, "파일전송차단 상세",  attach_block_df, _ORANGE_BG)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
