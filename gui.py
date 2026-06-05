import os
import threading
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
