from pathlib import Path
from loader import load_workbook
from config import load_keywords, load_settings
from rules.r01_usb import detect_usb_attempts
from rules.r02_design_files import detect_design_file_attachments
from rules.r03_large_zip import detect_large_zips
from rules.r04_ai_platforms import detect_ai_access
from rules.r05_external_email import detect_external_email_attempts
from rules.r06_messenger import detect_messenger_attachments
from aggregator import build_risk_table, RuleResults
from summarizer import generate_summary
from reporter import write_report
import re
from datetime import datetime


def _extract_period(filename: str) -> str:
    m = re.search(r"(\d{6})", filename)
    if m:
        ym = m.group(1)
        return f"{ym[:4]}년 {int(ym[4:6]):02d}월"
    return datetime.now().strftime("%Y년 %m월")


def run_analysis(
    excel_path: str,
    progress_cb=None,  # callable(float 0-1, str message)
) -> str:
    """Run full analysis pipeline. Returns path to output report."""

    def _progress(pct, msg):
        if progress_cb:
            progress_cb(pct, msg)

    kw = load_keywords()
    cfg = load_settings()

    _progress(0.05, "Excel 파일 읽는 중...")
    sheets = load_workbook(excel_path)

    _progress(0.15, "USB 시도 탐지 중...")
    usb_result = detect_usb_attempts(sheets["usb_blocked"])

    _progress(0.25, "설계파일 첨부 탐지 중...")
    design_result = detect_design_file_attachments(
        sheets["attach_all"],
        kw["design_keywords"],
        kw["design_extensions"],
    )

    _progress(0.35, "대용량 ZIP 탐지 중...")
    zip_result = detect_large_zips(sheets["attach_all"], kw["zip_size_threshold_mb"])

    _progress(0.45, "AI 플랫폼 접속 탐지 중...")
    ai_result = detect_ai_access(
        sheets["web_blocks"],
        sheets["attach_all"],
        kw["ai_domains"],
        kw["browser_processes"],
    )

    _progress(0.55, "외부 이메일 시도 탐지 중...")
    email_result = detect_external_email_attempts(sheets["web_blocks"], kw["external_mail_domains"])

    _progress(0.65, "메신저 파일전송 탐지 중...")
    messenger_result = detect_messenger_attachments(sheets["attach_all"])

    _progress(0.70, "위험도 종합 분석 중...")
    rule_results = RuleResults(
        usb=usb_result,
        design=design_result,
        large_zip=zip_result,
        ai=ai_result,
        email=email_result,
        messenger=messenger_result,
    )
    risk_table = build_risk_table(rule_results, top_n=cfg["top_n"])

    counts = {
        "USB 시도":       len(usb_result),
        "설계파일 첨부":   len(design_result),
        "대용량 ZIP 첨부": len(zip_result),
        "AI 플랫폼 접속":  len(ai_result),
        "외부메일 시도":   len(email_result),
        "메신저 파일전송": len(messenger_result),
    }

    _progress(0.80, "AI 요약 생성 중 (Claude)...")
    risk_text = risk_table.to_string(index=False) if not risk_table.empty else "없음"
    ai_summary = generate_summary(risk_text, counts, cfg["api_key"])

    _progress(0.90, "보고서 작성 중...")
    period = _extract_period(Path(excel_path).name)
    ym = re.search(r"(\d{6})", Path(excel_path).name)
    out_name = f"보안위협_{ym.group(1) if ym else 'report'}.xlsx"
    out_path = str(Path(cfg["report_dir"]) / out_name)

    write_report(
        out_path, period, counts, risk_table, ai_summary,
        usb_result, design_result, zip_result,
        ai_result, email_result, messenger_result,
    )

    _progress(1.0, f"완료! → {out_path}")
    return out_path
