# DLP 보안 위협 분석기 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a CustomTkinter desktop GUI that reads an OfficeKeeper DLP Excel report, applies 7 threat-detection rules, calls Claude API for AI narrative, and writes a formatted Excel summary.

**Architecture:** A loader reads the Excel workbook by sheet index (encoding-safe), passes normalized DataFrames to six rule modules, an aggregator computes per-user risk scores, a summarizer calls Claude API, and a reporter writes the output workbook. The GUI wires these together with a progress bar and settings panel.

**Tech Stack:** Python 3.11+, customtkinter, tkinterdnd2, pandas, openpyxl, anthropic

---

## Sheet Index Reference (오피스키퍼 DLP Excel)

| Index | Description | Rows | Key Columns |
|-------|-------------|------|-------------|
| 0 | 에이전트 무력화시도내역 (Print/agent log) | ~26K | c0=이름, c1=부서, c3=분류, c4=프로세스, c6=시간 |
| 1 | 매체별 내역_전체파일 (All media/USB) | ~60 | c0=이름, c1=부서, c3=분류, c4=매체유형, c11=파일명, c13=파일크기, c26=시간 |
| 3 | 매체별 내역_사용차단 내역 (USB blocked) | ~817 | c0=이름, c1=부서, c3=분류, c4=매체유형, c26=시간 |
| 6 | 파일첨부 내역_전체파일 (All attachments) | ~32K | c0=이름, c1=부서, c3=분류, c4=프로세스, c7=파일명, c9=파일크기, c22=시간 |
| 7 | 파일첨부 내역_설정된 확장자파일 (Flagged ext) | ~7.7K | same as [6] |
| 16 | 소프트웨어 실행 차단내역 (Software blocks) | ~601 | c0=이름, c1=부서, c3=카테고리, c4=프로세스, c5=제어정책, c6=시간 |
| 17 | 프로세스/서비스 제어내역 (Process control) | ~512 | c0=이름, c1=부서, c3=서비스명, c4=프로세스, c5=시간 |
| 18 | 웹사이트 접속 차단내역 (Website blocks) | ~427 | c0=이름, c1=부서, c3=카테고리, c4=사이트, c5=제어정책, c6=시간 |

**Header rows to skip per sheet:**
- Sheets 0, 16, 17, 18: skip=2 (title row + header row; data starts row 2)
- Sheets 1, 3, 6, 7: skip=3 (title + header + sub-header; data starts row 3)

---

## File Structure

```
security/
├── main.py                  # Entry point — creates App, starts mainloop
├── gui.py                   # CustomTkinter window, settings, progress
├── loader.py                # Reads Excel workbook → dict of DataFrames
├── aggregator.py            # Combines rule results → per-user risk table
├── summarizer.py            # Calls Claude API → str summary
├── reporter.py              # Writes output Excel workbook
├── rules/
│   ├── __init__.py
│   ├── r01_usb.py           # Rule 1: USB attempts
│   ├── r02_design_files.py  # Rules 2+3: design ext + keyword attachments
│   ├── r03_large_zip.py     # Rule 4: zip >= 10 MB
│   ├── r04_ai_platforms.py  # Rule 5: AI site access + upload
│   ├── r05_external_email.py # Rule 6: external mail attempts
│   └── r06_messenger.py     # Rule 7: WeChat/KakaoTalk attachments
├── keywords.json            # Configurable keywords, domains, thresholds
├── config.json              # Auto-generated: API key, settings (git-ignored)
├── requirements.txt
└── tests/
    ├── conftest.py          # Shared DataFrame fixtures
    ├── test_loader.py
    ├── test_r01_usb.py
    ├── test_r02_design_files.py
    ├── test_r03_large_zip.py
    ├── test_r04_ai_platforms.py
    ├── test_r05_external_email.py
    ├── test_r06_messenger.py
    └── test_aggregator.py
```

---

## Task 1: Project Setup

**Files:**
- Create: `requirements.txt`
- Create: `keywords.json`

- [ ] **Step 1: Create requirements.txt**

```
customtkinter==5.2.2
tkinterdnd2==0.3.0
pandas==2.2.2
openpyxl==3.1.2
anthropic==0.30.0
pytest==8.2.2
```

- [ ] **Step 2: Install dependencies**

```powershell
cd c:\Users\USER\DEV\security
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Expected: All packages install without error.

- [ ] **Step 3: Create keywords.json**

```json
{
  "design_keywords": ["DB", "SCH", "REV", "PCB", "BOM", "GERBER"],
  "design_extensions": [".tar", ".gz", ".tg", ".brd", ".sch"],
  "zip_size_threshold_mb": 10,
  "ai_domains": [
    "chat.openai.com",
    "chatgpt.com",
    "gemini.google.com",
    "claude.ai",
    "copilot.microsoft.com",
    "perplexity.ai",
    "notebooklm.google.com"
  ],
  "external_mail_domains": [
    "mail.naver.com",
    "mail.google.com",
    "mail.daum.net",
    "mail.yahoo.com",
    "mail.kakao.com"
  ],
  "browser_processes": ["chrome.exe", "msedge.exe", "firefox.exe", "iexplore.exe"]
}
```

- [ ] **Step 4: Create rules/__init__.py**

```python
```

(Empty file — just marks rules/ as a package.)

- [ ] **Step 5: Commit**

```powershell
git init
git add requirements.txt keywords.json rules/__init__.py
git commit -m "chore: project setup with dependencies and keyword config"
```

---

## Task 2: Loader

**Files:**
- Create: `loader.py`
- Create: `tests/test_loader.py`

The loader reads the workbook once and returns a dict of DataFrames keyed by a short name. It handles multi-row headers by using `skiprows` based on sheet index.

- [ ] **Step 1: Write failing test**

```python
# tests/test_loader.py
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
```

- [ ] **Step 2: Run test to confirm failure**

```powershell
pytest tests/test_loader.py -v
```

Expected: ImportError or ModuleNotFoundError.

- [ ] **Step 3: Implement loader.py**

```python
# loader.py
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
```

- [ ] **Step 4: Run tests**

```powershell
pytest tests/test_loader.py -v
```

Expected: 4 tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add loader.py tests/test_loader.py
git commit -m "feat: loader reads DLP Excel and returns normalized DataFrames"
```

---

## Task 3: Config Module

**Files:**
- Create: `config.py`
- Create: `tests/test_config.py` (minimal)

- [ ] **Step 1: Write failing test**

```python
# tests/test_config.py
import json, os
from config import load_keywords, load_settings, save_settings

def test_load_keywords_returns_design_keywords():
    kw = load_keywords("keywords.json")
    assert "DB" in kw["design_keywords"]
    assert "design_extensions" in kw
    assert kw["zip_size_threshold_mb"] == 10

def test_settings_roundtrip(tmp_path):
    cfg = tmp_path / "config.json"
    save_settings({"api_key": "sk-test", "top_n": 15}, str(cfg))
    loaded = load_settings(str(cfg))
    assert loaded["api_key"] == "sk-test"
    assert loaded["top_n"] == 15

def test_load_settings_returns_defaults_when_missing(tmp_path):
    cfg = tmp_path / "no_config.json"
    loaded = load_settings(str(cfg))
    assert loaded["top_n"] == 10
    assert loaded["api_key"] == ""
```

- [ ] **Step 2: Run to confirm failure**

```powershell
pytest tests/test_config.py -v
```

- [ ] **Step 3: Implement config.py**

```python
# config.py
import json
from pathlib import Path

_DEFAULTS = {"api_key": "", "top_n": 10, "report_dir": "reports"}


def load_keywords(path: str = "keywords.json") -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def load_settings(path: str = "config.json") -> dict:
    p = Path(path)
    if not p.exists():
        return dict(_DEFAULTS)
    with open(p, encoding="utf-8") as f:
        data = json.load(f)
    return {**_DEFAULTS, **data}


def save_settings(settings: dict, path: str = "config.json") -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)
```

- [ ] **Step 4: Run tests**

```powershell
pytest tests/test_config.py -v
```

Expected: 3 tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add config.py tests/test_config.py
git commit -m "feat: config module for keywords.json and settings"
```

---

## Task 4: Shared Test Fixtures

**Files:**
- Create: `tests/conftest.py`

These fixtures are minimal DataFrames that mirror the structure returned by loader.py.

- [ ] **Step 1: Create conftest.py**

```python
# tests/conftest.py
import pandas as pd
import pytest


@pytest.fixture
def attach_df():
    """Minimal 파일첨부 내역_전체파일 DataFrame."""
    return pd.DataFrame({
        "이름":   ["홍길동", "홍길동",   "이순자",  "이순자",    "김철수"],
        "부서":   ["개발팀", "개발팀",   "영업팀",  "WeChat 첨부파일", "개발팀"],
        "분류":   ["첨부",   "첨부",    "첨부",    "첨부",       "첨부"],
        "프로세스": ["OUTLOOK.EXE", "OUTLOOK.EXE", "OUTLOOK.EXE", "WeChat.exe", "chrome.exe"],
        "파일명": ["schematic_rev3.SCH", "design.tar", "report.pdf",
                   "drawings.zip", "budget.xlsx"],
        "파일크기": ["120KB(약 0 ~ 1 MB)", "5120KB(약 5 ~ 6 MB)",
                    "300KB(약 0 ~ 1 MB)", "15360KB(약 15 ~ 16 MB)",
                    "200KB(약 0 ~ 1 MB)"],
        "시간": pd.to_datetime(["2026-04-01", "2026-04-02", "2026-04-03",
                                "2026-04-04", "2026-04-05"]),
    })


@pytest.fixture
def usb_blocked_df():
    """Minimal 매체별 내역_사용차단 내역 DataFrame."""
    return pd.DataFrame({
        "이름":    ["홍길동", "홍길동", "이순자"],
        "부서":    ["개발팀", "개발팀", "영업팀"],
        "분류":    ["차단",   "차단",  "차단"],
        "매체유형": ["FDD/USB 저장장치", "FDD/USB 저장장치", "FDD/USB 저장장치"],
        "시간":    pd.to_datetime(["2026-04-10", "2026-04-15", "2026-04-20"]),
    })


@pytest.fixture
def web_blocks_df():
    """Minimal 웹사이트 접속 차단내역 DataFrame."""
    return pd.DataFrame({
        "이름":    ["홍길동", "이순자",      "김철수",           "박민준"],
        "부서":    ["개발팀", "영업팀",      "개발팀",           "인사팀"],
        "카테고리": ["포털메일", "포털메일",  "AI서비스",        "AI서비스"],
        "사이트":  ["mail.naver.com", "mail.google.com",
                   "chat.openai.com", "gemini.google.com"],
        "제어정책": ["감시", "감시", "감시", "감시"],
        "시간":    pd.to_datetime(["2026-04-01", "2026-04-02",
                                  "2026-04-03", "2026-04-04"]),
    })
```

- [ ] **Step 2: Verify fixtures load**

```powershell
pytest tests/conftest.py --collect-only
```

Expected: no errors, fixtures listed.

- [ ] **Step 3: Commit**

```powershell
git add tests/conftest.py
git commit -m "test: shared DataFrame fixtures for rule unit tests"
```

---

## Task 5: Rule 1 — USB Attempts

**Files:**
- Create: `rules/r01_usb.py`
- Create: `tests/test_r01_usb.py`

- [ ] **Step 1: Write failing test**

```python
# tests/test_r01_usb.py
from rules.r01_usb import detect_usb_attempts


def test_detects_all_usb_attempters(usb_blocked_df):
    result = detect_usb_attempts(usb_blocked_df)
    assert set(result["이름"]) == {"홍길동", "이순자"}


def test_counts_per_user(usb_blocked_df):
    result = detect_usb_attempts(usb_blocked_df)
    hong = result[result["이름"] == "홍길동"].iloc[0]
    assert hong["시도횟수"] == 2


def test_empty_df_returns_empty(usb_blocked_df):
    import pandas as pd
    empty = usb_blocked_df.iloc[0:0]
    result = detect_usb_attempts(empty)
    assert len(result) == 0


def test_result_columns(usb_blocked_df):
    result = detect_usb_attempts(usb_blocked_df)
    assert set(result.columns) >= {"이름", "부서", "시도횟수", "최근시도일시"}
```

- [ ] **Step 2: Run to confirm failure**

```powershell
pytest tests/test_r01_usb.py -v
```

- [ ] **Step 3: Implement rules/r01_usb.py**

```python
# rules/r01_usb.py
import pandas as pd


def detect_usb_attempts(usb_blocked: pd.DataFrame) -> pd.DataFrame:
    """Return one row per user who attempted USB storage use."""
    usb = usb_blocked[
        usb_blocked["매체유형"].str.contains("USB", na=False, case=False)
    ].copy()
    if usb.empty:
        return pd.DataFrame(columns=["이름", "부서", "시도횟수", "최근시도일시"])
    grp = usb.groupby(["이름", "부서"], as_index=False).agg(
        시도횟수=("이름", "count"),
        최근시도일시=("시간", "max"),
    )
    return grp.sort_values("시도횟수", ascending=False).reset_index(drop=True)
```

- [ ] **Step 4: Run tests**

```powershell
pytest tests/test_r01_usb.py -v
```

Expected: 4 tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add rules/r01_usb.py tests/test_r01_usb.py
git commit -m "feat: rule 1 USB attempt detection"
```

---

## Task 6: Rules 2+3 — Design File Attachments

**Files:**
- Create: `rules/r02_design_files.py`
- Create: `tests/test_r02_design_files.py`

Detects: (a) attachments with tar/gz/tg extensions, (b) filenames containing design keywords (DB, SCH, REV, etc.).

- [ ] **Step 1: Write failing test**

```python
# tests/test_r02_design_files.py
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
```

- [ ] **Step 2: Run to confirm failure**

```powershell
pytest tests/test_r02_design_files.py -v
```

- [ ] **Step 3: Implement rules/r02_design_files.py**

```python
# rules/r02_design_files.py
import pandas as pd


def detect_design_file_attachments(
    attach: pd.DataFrame,
    design_keywords: list[str],
    design_extensions: list[str],
) -> pd.DataFrame:
    """Detect attachments matching design extensions or keywords."""
    if attach.empty:
        return pd.DataFrame(columns=list(attach.columns) + ["탐지사유"])

    fname = attach["파일명"].fillna("").str.lower()

    ext_lower = [e.lower() for e in design_extensions]
    kw_lower = [k.lower() for k in design_keywords]

    ext_mask = fname.apply(lambda f: any(f.endswith(e) for e in ext_lower))
    kw_mask = fname.apply(lambda f: any(k in f for k in kw_lower))

    matched = attach[ext_mask | kw_mask].copy()

    def reason(f: str) -> str:
        f = f.lower()
        reasons = []
        hit_ext = [e for e in ext_lower if f.endswith(e)]
        hit_kw = [k for k in kw_lower if k in f]
        if hit_ext:
            reasons.append(f"확장자({hit_ext[0]})")
        if hit_kw:
            reasons.append(f"키워드({','.join(hit_kw).upper()})")
        return ", ".join(reasons)

    matched["탐지사유"] = matched["파일명"].fillna("").apply(reason)
    return matched.reset_index(drop=True)
```

- [ ] **Step 4: Run tests**

```powershell
pytest tests/test_r02_design_files.py -v
```

Expected: 5 tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add rules/r02_design_files.py tests/test_r02_design_files.py
git commit -m "feat: rules 2+3 design file attachment detection (ext + keyword)"
```

---

## Task 7: Rule 4 — Large ZIP Attachments

**Files:**
- Create: `rules/r03_large_zip.py`
- Create: `tests/test_r03_large_zip.py`

File size column format: `"15360KB(약 15 ~ 16 MB)"` — parse KB value with regex.

- [ ] **Step 1: Write failing test**

```python
# tests/test_r03_large_zip.py
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
    # No small zip in attach_df; add one and verify excluded
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
```

- [ ] **Step 2: Run to confirm failure**

```powershell
pytest tests/test_r03_large_zip.py -v
```

- [ ] **Step 3: Implement rules/r03_large_zip.py**

```python
# rules/r03_large_zip.py
import re
import pandas as pd


def parse_size_kb(size_str) -> int:
    if not size_str:
        return 0
    m = re.search(r"(\d+)KB", str(size_str))
    return int(m.group(1)) if m else 0


def detect_large_zips(attach: pd.DataFrame, threshold_mb: float = 10) -> pd.DataFrame:
    """Detect .zip attachments >= threshold_mb in size."""
    if attach.empty:
        return attach.iloc[0:0].copy()
    threshold_kb = int(threshold_mb * 1024)
    is_zip = attach["파일명"].fillna("").str.lower().str.endswith(".zip")
    size_kb = attach["파일크기"].apply(parse_size_kb)
    matched = attach[is_zip & (size_kb >= threshold_kb)].copy()
    matched["파일크기_KB"] = size_kb[matched.index]
    return matched.reset_index(drop=True)
```

- [ ] **Step 4: Run tests**

```powershell
pytest tests/test_r03_large_zip.py -v
```

Expected: 7 tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add rules/r03_large_zip.py tests/test_r03_large_zip.py
git commit -m "feat: rule 4 large ZIP attachment detection with KB parser"
```

---

## Task 8: Rule 5 — AI Platform Access + Upload Detection

**Files:**
- Create: `rules/r04_ai_platforms.py`
- Create: `tests/test_r04_ai_platforms.py`

Upload determination: user appears in web_blocks for an AI domain AND has a browser-process attachment on the same calendar date → probable upload (🔴). AI domain access only → access log (🟠).

- [ ] **Step 1: Write failing test**

```python
# tests/test_r04_ai_platforms.py
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
```

- [ ] **Step 2: Run to confirm failure**

```powershell
pytest tests/test_r04_ai_platforms.py -v
```

- [ ] **Step 3: Implement rules/r04_ai_platforms.py**

```python
# rules/r04_ai_platforms.py
import pandas as pd


def detect_ai_access(
    web_blocks: pd.DataFrame,
    attach_all: pd.DataFrame,
    ai_domains: list[str],
    browser_processes: list[str],
) -> pd.DataFrame:
    """Detect AI platform access; flag probable uploads."""
    if web_blocks.empty or "사이트" not in web_blocks.columns:
        return pd.DataFrame(columns=["이름", "부서", "사이트", "접속횟수", "업로드여부", "최근접속일시"])

    ai_lower = [d.lower() for d in ai_domains]
    site_col = web_blocks["사이트"].fillna("").str.lower()
    is_ai = site_col.apply(lambda s: any(d in s for d in ai_lower))
    ai_df = web_blocks[is_ai].copy()

    if ai_df.empty:
        return pd.DataFrame(columns=["이름", "부서", "사이트", "접속횟수", "업로드여부", "최근접속일시"])

    # Build set of (이름, date) with browser attachment
    browser_dates: set[tuple[str, str]] = set()
    if not attach_all.empty and "프로세스" in attach_all.columns:
        br_lower = [b.lower() for b in browser_processes]
        is_browser = attach_all["프로세스"].fillna("").str.lower().apply(
            lambda p: any(b in p for b in br_lower)
        )
        br_att = attach_all[is_browser].copy()
        if not br_att.empty and "시간" in br_att.columns:
            br_att["날짜"] = pd.to_datetime(br_att["시간"]).dt.date
            for _, row in br_att.iterrows():
                browser_dates.add((row["이름"], str(row["날짜"])))

    ai_df["날짜"] = pd.to_datetime(ai_df["시간"]).dt.date.astype(str)

    records = []
    for (name, dept, site), grp in ai_df.groupby(["이름", "부서", "사이트"]):
        has_upload = any((name, d) in browser_dates for d in grp["날짜"])
        records.append({
            "이름": name,
            "부서": dept,
            "사이트": site,
            "접속횟수": len(grp),
            "업로드여부": "업로드(추정)" if has_upload else "접속",
            "최근접속일시": grp["시간"].max(),
        })

    return pd.DataFrame(records).sort_values("접속횟수", ascending=False).reset_index(drop=True)
```

- [ ] **Step 4: Run tests**

```powershell
pytest tests/test_r04_ai_platforms.py -v
```

Expected: 5 tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add rules/r04_ai_platforms.py tests/test_r04_ai_platforms.py
git commit -m "feat: rule 5 AI platform access and upload detection"
```

---

## Task 9: Rule 6 — External Email Attempts

**Files:**
- Create: `rules/r05_external_email.py`
- Create: `tests/test_r05_external_email.py`

- [ ] **Step 1: Write failing test**

```python
# tests/test_r05_external_email.py
from rules.r05_external_email import detect_external_email_attempts

MAIL_DOMAINS = ["mail.naver.com", "mail.google.com", "mail.daum.net",
                "mail.yahoo.com", "mail.kakao.com"]


def test_detects_naver_mail(web_blocks_df):
    result = detect_external_email_attempts(web_blocks_df, MAIL_DOMAINS)
    sites = result["사이트"].tolist()
    assert "mail.naver.com" in sites


def test_detects_gmail(web_blocks_df):
    result = detect_external_email_attempts(web_blocks_df, MAIL_DOMAINS)
    sites = result["사이트"].tolist()
    assert "mail.google.com" in sites


def test_ai_sites_excluded(web_blocks_df):
    result = detect_external_email_attempts(web_blocks_df, MAIL_DOMAINS)
    sites = result["사이트"].tolist()
    assert "chat.openai.com" not in sites


def test_result_columns(web_blocks_df):
    result = detect_external_email_attempts(web_blocks_df, MAIL_DOMAINS)
    assert set(result.columns) >= {"이름", "부서", "사이트", "시도횟수", "최근시도일시"}


def test_empty_returns_empty():
    import pandas as pd
    result = detect_external_email_attempts(pd.DataFrame(), MAIL_DOMAINS)
    assert len(result) == 0
```

- [ ] **Step 2: Run to confirm failure**

```powershell
pytest tests/test_r05_external_email.py -v
```

- [ ] **Step 3: Implement rules/r05_external_email.py**

```python
# rules/r05_external_email.py
import pandas as pd


def detect_external_email_attempts(
    web_blocks: pd.DataFrame,
    mail_domains: list[str],
) -> pd.DataFrame:
    """Detect access attempts to external mail services."""
    if web_blocks.empty or "사이트" not in web_blocks.columns:
        return pd.DataFrame(columns=["이름", "부서", "사이트", "시도횟수", "최근시도일시"])

    mail_lower = [d.lower() for d in mail_domains]
    site_col = web_blocks["사이트"].fillna("").str.lower()
    is_mail = site_col.apply(lambda s: any(d in s for d in mail_lower))
    mail_df = web_blocks[is_mail].copy()

    if mail_df.empty:
        return pd.DataFrame(columns=["이름", "부서", "사이트", "시도횟수", "최근시도일시"])

    grp = mail_df.groupby(["이름", "부서", "사이트"], as_index=False).agg(
        시도횟수=("이름", "count"),
        최근시도일시=("시간", "max"),
    )
    return grp.sort_values("시도횟수", ascending=False).reset_index(drop=True)
```

- [ ] **Step 4: Run tests**

```powershell
pytest tests/test_r05_external_email.py -v
```

Expected: 5 tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add rules/r05_external_email.py tests/test_r05_external_email.py
git commit -m "feat: rule 6 external email access detection"
```

---

## Task 10: Rule 7 — WeChat/KakaoTalk Messenger Attachments

**Files:**
- Create: `rules/r06_messenger.py`
- Create: `tests/test_r06_messenger.py`

Detects files sent via WeChat or KakaoTalk by filtering the attachment log where the 부서 column or 프로세스 column contains messenger identifiers.

- [ ] **Step 1: Write failing test**

```python
# tests/test_r06_messenger.py
import pandas as pd
from rules.r06_messenger import detect_messenger_attachments


def test_detects_wechat_by_dept(attach_df):
    result = detect_messenger_attachments(attach_df)
    assert "이순자" in result["이름"].tolist()  # dept contains "WeChat"


def test_non_messenger_excluded(attach_df):
    result = detect_messenger_attachments(attach_df)
    # 홍길동 uses OUTLOOK.EXE with no WeChat dept → should NOT appear
    assert "홍길동" not in result["이름"].tolist()


def test_result_has_platform_column(attach_df):
    result = detect_messenger_attachments(attach_df)
    assert "플랫폼" in result.columns


def test_platform_identified_as_wechat(attach_df):
    result = detect_messenger_attachments(attach_df)
    row = result[result["이름"] == "이순자"].iloc[0]
    assert "WeChat" in row["플랫폼"] or "wechat" in row["플랫폼"].lower()


def test_empty_returns_empty():
    result = detect_messenger_attachments(pd.DataFrame())
    assert len(result) == 0
```

- [ ] **Step 2: Run to confirm failure**

```powershell
pytest tests/test_r06_messenger.py -v
```

- [ ] **Step 3: Implement rules/r06_messenger.py**

```python
# rules/r06_messenger.py
import pandas as pd

_MESSENGER_PATTERNS = [
    ("WeChat",    ["wechat"]),
    ("KakaoTalk", ["kakaotalk", "kakao"]),
]


def _identify_platform(dept: str, proc: str) -> str | None:
    text = (str(dept) + " " + str(proc)).lower()
    for name, patterns in _MESSENGER_PATTERNS:
        if any(p in text for p in patterns):
            return name
    return None


def detect_messenger_attachments(attach: pd.DataFrame) -> pd.DataFrame:
    """Detect file attachments sent via WeChat or KakaoTalk."""
    if attach.empty or "부서" not in attach.columns:
        return pd.DataFrame(columns=["이름", "부서", "플랫폼", "파일명", "파일크기", "시간"])

    dept_col = attach["부서"].fillna("")
    proc_col = attach.get("프로세스", pd.Series([""] * len(attach))).fillna("")

    platforms = [_identify_platform(d, p) for d, p in zip(dept_col, proc_col)]
    mask = [p is not None for p in platforms]
    matched = attach[mask].copy()
    matched["플랫폼"] = [p for p, m in zip(platforms, mask) if m]
    return matched.reset_index(drop=True)
```

- [ ] **Step 4: Run tests**

```powershell
pytest tests/test_r06_messenger.py -v
```

Expected: 5 tests PASS.

- [ ] **Step 5: Commit**

```powershell
git add rules/r06_messenger.py tests/test_r06_messenger.py
git commit -m "feat: rule 7 WeChat/KakaoTalk attachment detection"
```

---

## Task 11: Aggregator — Per-User Risk Scoring

**Files:**
- Create: `aggregator.py`
- Create: `tests/test_aggregator.py`

Combines all rule results into one risk table with a composite score. Weights: USB×3, design file×2, large zip×1, AI platform×2, external email×2, messenger×1.

- [ ] **Step 1: Write failing test**

```python
# tests/test_aggregator.py
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
```

- [ ] **Step 2: Run to confirm failure**

```powershell
pytest tests/test_aggregator.py -v
```

- [ ] **Step 3: Implement aggregator.py**

```python
# aggregator.py
from dataclasses import dataclass
import pandas as pd

_WEIGHTS = {
    "USB시도": 3,
    "설계파일": 2,
    "대용량ZIP": 1,
    "AI플랫폼": 2,
    "외부메일": 2,
    "메신저": 1,
}


@dataclass
class RuleResults:
    usb: pd.DataFrame
    design: pd.DataFrame
    large_zip: pd.DataFrame
    ai: pd.DataFrame
    email: pd.DataFrame
    messenger: pd.DataFrame


def _count_per_user(df: pd.DataFrame, count_col: str = None) -> dict[tuple[str, str], int]:
    """Count events per (이름, 부서). If count_col is given, sum that column instead of rows."""
    if df.empty or "이름" not in df.columns:
        return {}
    dept_col = df["부서"] if "부서" in df.columns else pd.Series([""] * len(df))
    counts: dict[tuple[str, str], int] = {}
    for i, (name, dept) in enumerate(zip(df["이름"], dept_col)):
        key = (str(name), str(dept))
        val = int(df[count_col].iloc[i] or 1) if count_col and count_col in df.columns else 1
        counts[key] = counts.get(key, 0) + val
    return counts


def build_risk_table(results: RuleResults, top_n: int = 10) -> pd.DataFrame:
    """Return top_n users sorted by composite risk score."""
    sources = {
        "USB시도":   _count_per_user(results.usb,      count_col="시도횟수"),
        "설계파일":  _count_per_user(results.design),
        "대용량ZIP": _count_per_user(results.large_zip),
        "AI플랫폼":  _count_per_user(results.ai,       count_col="접속횟수"),
        "외부메일":  _count_per_user(results.email,    count_col="시도횟수"),
        "메신저":    _count_per_user(results.messenger),
    }

    all_users: set[tuple[str, str]] = set()
    for counts in sources.values():
        all_users.update(counts.keys())

    rows = []
    for (name, dept) in all_users:
        row: dict = {"이름": name, "부서": dept}
        score = 0
        for col, counts in sources.items():
            n = counts.get((name, dept), 0)
            row[col] = n
            score += n * _WEIGHTS[col]
        row["위험점수"] = score
        rows.append(row)

    if not rows:
        cols = ["이름", "부서"] + list(_WEIGHTS.keys()) + ["위험점수"]
        return pd.DataFrame(columns=cols)

    df = pd.DataFrame(rows)
    df = df.sort_values("위험점수", ascending=False).head(top_n).reset_index(drop=True)
    df.insert(0, "순위", range(1, len(df) + 1))
    return df
```

- [ ] **Step 4: Run tests**

```powershell
pytest tests/test_aggregator.py -v
```

Expected: 4 tests PASS.

- [ ] **Step 5: Run full test suite**

```powershell
pytest tests/ -v
```

Expected: all tests PASS.

- [ ] **Step 6: Commit**

```powershell
git add aggregator.py tests/test_aggregator.py
git commit -m "feat: risk aggregator combines all rule results with weighted scoring"
```

---

## Task 12: AI Summarizer

**Files:**
- Create: `summarizer.py`

No unit test for this (requires live API key). Integration is verified via the GUI in Task 14.

- [ ] **Step 1: Create summarizer.py**

```python
# summarizer.py
import anthropic


def generate_summary(risk_table_text: str, counts: dict[str, int], api_key: str) -> str:
    """Call Claude API and return Korean-language threat narrative."""
    if not api_key:
        return "(API 키가 설정되지 않아 AI 요약을 생성하지 못했습니다.)"

    counts_text = "\n".join(f"- {k}: {v}건" for k, v in counts.items())
    prompt = f"""당신은 회사 보안 담당자의 DLP 분석을 돕는 보안 전문가입니다.
아래는 이번 달 DLP 탐지 결과입니다.

[탐지 건수 요약]
{counts_text}

[상위 위험 사용자]
{risk_table_text}

다음 세 가지를 한국어로 작성하세요:
1. 이달의 위협 총평 (2~3문장)
2. 상위 3명 사용자의 행동 패턴 코멘트 (각 1~2문장)
3. 즉시 조치 권고사항 (불릿 포인트 3~5개)

간결하고 실무적으로 작성하세요."""

    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text
```

- [ ] **Step 2: Commit**

```powershell
git add summarizer.py
git commit -m "feat: Claude API summarizer for AI threat narrative"
```

---

## Task 13: Reporter — Excel Output

**Files:**
- Create: `reporter.py`

- [ ] **Step 1: Create reporter.py**

```python
# reporter.py
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
```

- [ ] **Step 2: Quick smoke test**

```powershell
python -c "
from reporter import write_report
import pandas as pd
counts = {'USB 시도':1,'설계파일 첨부':2,'대용량 ZIP 첨부':0,'AI 플랫폼 접속':1,'외부메일 시도':3,'메신저 파일전송':4}
risk = pd.DataFrame({'순위':[1],'이름':['테스트'],'부서':['개발팀'],'위험점수':[9]})
write_report('reports/test_report.xlsx','2026년 04월',counts,risk,'테스트 요약',
    pd.DataFrame(),pd.DataFrame(),pd.DataFrame(),pd.DataFrame(),pd.DataFrame(),pd.DataFrame())
print('OK — reports/test_report.xlsx created')
"
```

Expected: `OK — reports/test_report.xlsx created`

- [ ] **Step 3: Commit**

```powershell
git add reporter.py
git commit -m "feat: Excel reporter with summary and detail sheets"
```

---

## Task 14: Analysis Pipeline Wiring

**Files:**
- Create: `pipeline.py`

Wraps loader + all rules + aggregator + summarizer into one `run_analysis()` call that the GUI calls in a background thread.

- [ ] **Step 1: Create pipeline.py**

```python
# pipeline.py
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
```

- [ ] **Step 2: Smoke test pipeline (no GUI)**

```powershell
python -c "
from pipeline import run_analysis
path = run_analysis(
    '오피스키퍼_정보유출방지_다믈파워반도체_202604.xlsx',
    progress_cb=lambda p, m: print(f'{int(p*100):3d}% {m}')
)
print('Report saved to:', path)
"
```

Expected: progress messages printed, report saved to `reports/보안위협_202604.xlsx`.

- [ ] **Step 3: Commit**

```powershell
git add pipeline.py
git commit -m "feat: analysis pipeline wires all rules, aggregator, summarizer, reporter"
```

---

## Task 15: GUI

**Files:**
- Create: `gui.py`
- Create: `main.py`

- [ ] **Step 1: Create gui.py**

```python
# gui.py
import os
import threading
import subprocess
import customtkinter as ctk
from tkinter import filedialog, messagebox
from config import load_settings, save_settings
from pipeline import run_analysis

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("🔒 DLP 보안 위협 분석기")
        self.geometry("620x560")
        self.resizable(False, False)
        self._settings = load_settings()
        self._selected_file = ctk.StringVar(value="")
        self._build_ui()

    def _build_ui(self):
        # File selection
        file_frame = ctk.CTkFrame(self, corner_radius=10)
        file_frame.pack(fill="x", padx=20, pady=(20, 8))
        ctk.CTkLabel(file_frame, text="📂 분석할 Excel 파일 선택", font=("", 14, "bold")).pack(pady=(12, 6))
        ctk.CTkButton(file_frame, text="파일 선택", command=self._pick_file).pack(pady=4)
        ctk.CTkLabel(file_frame, textvariable=self._selected_file,
                     font=("", 11), text_color="gray").pack(pady=(2, 12))

        # Settings
        cfg_frame = ctk.CTkFrame(self, corner_radius=10)
        cfg_frame.pack(fill="x", padx=20, pady=8)
        ctk.CTkLabel(cfg_frame, text="⚙️  설정", font=("", 13, "bold")).grid(
            row=0, column=0, padx=12, pady=(10, 4), sticky="w", columnspan=2)

        ctk.CTkLabel(cfg_frame, text="상위 위험 사용자 수:").grid(row=1, column=0, padx=12, sticky="w")
        self._top_n = ctk.CTkEntry(cfg_frame, width=60)
        self._top_n.insert(0, str(self._settings.get("top_n", 10)))
        self._top_n.grid(row=1, column=1, padx=12, pady=4, sticky="w")

        ctk.CTkLabel(cfg_frame, text="Claude API 키:").grid(row=2, column=0, padx=12, sticky="w")
        api_row = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        api_row.grid(row=2, column=1, padx=12, pady=4, sticky="w")
        self._api_key = ctk.CTkEntry(api_row, width=260, show="•")
        self._api_key.insert(0, self._settings.get("api_key", ""))
        self._api_key.pack(side="left")
        ctk.CTkButton(api_row, text="저장", width=50, command=self._save_settings).pack(
            side="left", padx=(6, 0))

        ctk.CTkLabel(cfg_frame, text="키워드 파일:").grid(row=3, column=0, padx=12, sticky="w")
        ctk.CTkButton(cfg_frame, text="keywords.json 편집", width=160,
                      command=lambda: os.startfile("keywords.json")).grid(
            row=3, column=1, padx=12, pady=4, sticky="w")

        ctk.CTkLabel(cfg_frame, text="보고서 저장 위치:").grid(row=4, column=0, padx=12, sticky="w")
        dir_row = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        dir_row.grid(row=4, column=1, padx=12, pady=(4, 12), sticky="w")
        self._report_dir = ctk.CTkEntry(dir_row, width=200)
        self._report_dir.insert(0, self._settings.get("report_dir", "reports"))
        self._report_dir.pack(side="left")
        ctk.CTkButton(dir_row, text="변경", width=50, command=self._pick_dir).pack(
            side="left", padx=(6, 0))

        # Analyze button
        self._analyze_btn = ctk.CTkButton(
            self, text="🔍 분석 시작", height=44, font=("", 15, "bold"),
            command=self._start_analysis)
        self._analyze_btn.pack(fill="x", padx=20, pady=12)

        # Progress
        self._progress = ctk.CTkProgressBar(self)
        self._progress.set(0)
        self._progress.pack(fill="x", padx=20, pady=(0, 4))
        self._status_label = ctk.CTkLabel(self, text="", font=("", 11), text_color="gray")
        self._status_label.pack()

        # Result
        result_frame = ctk.CTkFrame(self, corner_radius=10)
        result_frame.pack(fill="x", padx=20, pady=(8, 20))
        self._result_label = ctk.CTkLabel(result_frame, text="", font=("", 11))
        self._result_label.pack(side="left", padx=12, pady=10)
        self._open_btn = ctk.CTkButton(result_frame, text="보고서 열기", width=100,
                                        command=self._open_report, state="disabled")
        self._open_btn.pack(side="right", padx=12, pady=10)
        self._report_path = ""

    def _pick_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")])
        if path:
            self._selected_file.set(path)

    def _pick_dir(self):
        d = filedialog.askdirectory()
        if d:
            self._report_dir.delete(0, "end")
            self._report_dir.insert(0, d)

    def _save_settings(self):
        cfg = load_settings()
        try:
            cfg["top_n"] = int(self._top_n.get())
        except ValueError:
            cfg["top_n"] = 10
        cfg["api_key"] = self._api_key.get().strip()
        cfg["report_dir"] = self._report_dir.get().strip() or "reports"
        save_settings(cfg)
        messagebox.showinfo("저장 완료", "설정이 저장되었습니다.")

    def _start_analysis(self):
        path = self._selected_file.get()
        if not path:
            messagebox.showwarning("파일 없음", "분석할 Excel 파일을 선택해주세요.")
            return
        self._save_settings()
        self._analyze_btn.configure(state="disabled")
        self._open_btn.configure(state="disabled")
        self._result_label.configure(text="")
        self._progress.set(0)
        thread = threading.Thread(target=self._run_in_thread, args=(path,), daemon=True)
        thread.start()

    def _run_in_thread(self, path: str):
        try:
            def cb(pct, msg):
                self.after(0, self._progress.set, pct)
                self.after(0, self._status_label.configure, {"text": msg})

            report_path = run_analysis(path, progress_cb=cb)
            self._report_path = report_path
            self.after(0, self._on_success, report_path)
        except Exception as e:
            self.after(0, self._on_error, str(e))

    def _on_success(self, report_path: str):
        self._analyze_btn.configure(state="normal")
        self._open_btn.configure(state="normal")
        self._result_label.configure(text=f"✅ 완료! → {report_path}", text_color="green")

    def _on_error(self, msg: str):
        self._analyze_btn.configure(state="normal")
        self._progress.set(0)
        self._status_label.configure(text="")
        messagebox.showerror("분석 오류", f"분석 중 오류가 발생했습니다:\n{msg}")

    def _open_report(self):
        if self._report_path:
            os.startfile(self._report_path)
```

- [ ] **Step 2: Create main.py**

```python
# main.py
from gui import App

if __name__ == "__main__":
    app = App()
    app.mainloop()
```

- [ ] **Step 3: Run the app**

```powershell
python main.py
```

Expected: GUI window opens. Select `오피스키퍼_정보유출방지_다믈파워반도체_202604.xlsx`, enter API key, click "분석 시작". Progress bar advances. Report opens in Excel.

- [ ] **Step 4: Commit**

```powershell
git add gui.py main.py
git commit -m "feat: CustomTkinter GUI with file selection, settings, and progress bar"
```

---

## Task 16: Final Integration Smoke Test

- [ ] **Step 1: Run all tests**

```powershell
pytest tests/ -v
```

Expected: all tests PASS.

- [ ] **Step 2: End-to-end test with real file (no GUI)**

```powershell
python -c "
from pipeline import run_analysis
path = run_analysis(
    '오피스키퍼_정보유출방지_다믈파워반도체_202604.xlsx',
    progress_cb=lambda p, m: print(f'{int(p*100):3d}% {m}')
)
import openpyxl
wb = openpyxl.load_workbook(path)
print('Sheets:', wb.sheetnames)
print('Summary rows:', wb.active.max_row)
"
```

Expected: 7 sheet names printed, max_row > 20.

- [ ] **Step 3: Final commit**

```powershell
git add .
git commit -m "feat: DLP 보안 위협 분석기 complete — GUI + 7 rules + Claude AI summary"
```
