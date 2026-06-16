from pathlib import Path
from loader import load_workbook
from config import load_keywords, load_settings
from rules.r01_usb import detect_usb_attempts
from rules.r02_design_files import detect_design_file_attachments
from rules.r03_large_zip import detect_large_zips
from rules.r04_ai_platforms import detect_ai_access
from rules.r05_external_email import detect_external_email_attempts
from rules.r06_messenger import detect_messenger_attachments
from rules.r07_blocked_transfers import (
    detect_usb_file_blocks,
    detect_attachment_blocks,
    summarize_web_policy,
)
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

    _progress(0.12, "USB 시도 탐지 중...")
    usb_result = detect_usb_attempts(sheets["usb_blocked"])

    _progress(0.20, "USB 파일복사 차단 탐지 중...")
    usb_block_result = detect_usb_file_blocks(sheets["usb_all"])

    _progress(0.28, "설계파일 첨부 탐지 중...")
    design_result = detect_design_file_attachments(
        sheets["attach_all"],
        kw["design_keywords"],
        kw["design_extensions"],
    )

    _progress(0.36, "대용량 ZIP 탐지 중...")
    zip_result = detect_large_zips(sheets["attach_all"], kw["zip_size_threshold_mb"])

    _progress(0.44, "AI 플랫폼 접속 탐지 중...")
    ai_result = detect_ai_access(
        sheets["web_blocks"],
        sheets["attach_all"],
        kw["ai_domains"],
        kw["browser_processes"],
    )

    _progress(0.52, "외부 이메일 시도 탐지 중...")
    email_result = detect_external_email_attempts(sheets["web_blocks"], kw["external_mail_domains"])

    _progress(0.58, "메신저 파일전송 탐지 중...")
    messenger_result = detect_messenger_attachments(sheets["attach_all"])

    _progress(0.64, "파일전송 차단 내역 탐지 중...")
    attach_block_result = detect_attachment_blocks(sheets["attach_all"])
    web_policy = summarize_web_policy(sheets["web_blocks"])

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

    threat_counts = {
        "USB 시도":        len(usb_result),
        "설계파일 첨부":    len(design_result),
        "대용량 ZIP 첨부":  len(zip_result),
        "AI 플랫폼 접속":   len(ai_result),
        "외부메일 시도":    len(email_result),
        "메신저 파일전송":  len(messenger_result),
    }
    block_counts = {
        "USB 파일복사 차단":    len(usb_block_result),
        "파일전송 차단":        len(attach_block_result),
        "웹사이트 접속 차단":   web_policy["차단"],
        "웹사이트 접속 경고":   web_policy["경고"],
    }

    _progress(0.80, "AI 요약 생성 중 (Claude)...")
    risk_text = risk_table.to_string(index=False) if not risk_table.empty else "없음"
    ai_summary = generate_summary(risk_text, {**block_counts, **threat_counts}, cfg["api_key"])

    _progress(0.90, "보고서 작성 중...")
    period = _extract_period(Path(excel_path).name)
    ym = re.search(r"(\d{6})", Path(excel_path).name)
    out_name = f"보안위협_{ym.group(1) if ym else 'report'}.xlsx"
    out_path = str(Path(cfg["report_dir"]) / out_name)

    write_report(
        output_path=out_path,
        period=period,
        filename=Path(excel_path).name,
        threat_counts=threat_counts,
        block_counts=block_counts,
        risk_table=risk_table,
        ai_summary=ai_summary,
        usb_df=usb_result,
        design_df=design_result,
        large_zip_df=zip_result,
        ai_df=ai_result,
        email_df=email_result,
        messenger_df=messenger_result,
        usb_block_df=usb_block_result,
        attach_block_df=attach_block_result,
    )

    _progress(1.0, f"완료! → {out_path}")
    return out_path
