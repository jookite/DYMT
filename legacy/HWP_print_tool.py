"""
한글 문서 본문/답지 분할 인쇄 도구 (통합판)

두 가지 모드를 탭으로 제공합니다.
  - 파일 목록 모드: 파일별로 본문/답지 부수를 개별 지정
  - 폴더 일괄 모드: 폴더 내 모든 .hwp/.hwpx 파일을 1부씩 일괄 인쇄

공통 로직(미주 탐색, 인쇄 실행)은 HwpPrinter 클래스로 분리되어 있어
두 탭이 동일한 처리 루틴을 공유합니다.
"""

import os
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import win32com.client
import win32print
import pythoncom


# ============================================================
# 공통 HWP 처리 로직
# ============================================================
class HwpPrinter:
    """한글 문서 열기, 미주 탐색, 본문/답지 분할 인쇄 담당."""

    def __init__(self, log_fn):
        """log_fn: 메시지를 UI로 전달할 콜백 (예: App.log)."""
        self.log = log_fn
        self.hwp = None

    def start(self):
        """HwpObject 초기화."""
        self.hwp = win32com.client.dynamic.Dispatch("HwpFrame.HwpObject.2")
        self.hwp.XHwpWindows.Item(0).Visible = True

    def quit(self):
        """HwpObject 종료."""
        if self.hwp is not None:
            try:
                self.hwp.Quit()
            except Exception:
                pass
            self.hwp = None

    # ---------- 미주 탐색 ----------
    def _note_shown_page(self, ctrl, at_end=False):
        """미주 내용의 시작(또는 끝) 위치에 표시되는 쪽 번호."""
        self.hwp.SetPosBySet(ctrl.GetAnchorPos(0))
        self.hwp.Run("NoteModify")        # 미주 내용 위치로 이동
        if at_end:
            self.hwp.Run("MoveSelListEnd")
        page = self.hwp.KeyIndicator()[3]
        if at_end:
            self.hwp.Run("Cancel")
        self.hwp.Run("CloseUpper")        # 본문으로 복귀
        return page

    COVER_KEYWORDS = ("정답", "해설", "답지")  # 답지 표지 제목에 흔한 단어

    @staticmethod
    def _squash(text):
        return "".join(str(text).split())

    def _page_text(self, page):
        """물리 쪽(1부터)의 본문 텍스트 (미주 내용은 포함되지 않음)."""
        try:
            return self._squash(self.hwp.GetPageText(page - 1, -1))
        except Exception:
            return ""

    def find_split_page_by_first_endnote(self):
        """답지가 시작되는 실제(물리) 쪽 번호를 계산 (미주 없으면 None).

        GetPageText()는 미주 내용을 제외한 본문 텍스트를 물리 쪽 단위로
        돌려주므로, 문서 끝에서부터 본문 텍스트가 남아 있는 마지막 쪽(E)을
        찾는다. E 첫머리에 '정답/해설/답지' 류 제목이 있으면 답지 표지로
        보고 E부터, 아니면 본문 마지막 쪽이므로 E+1부터를 답지로 삼는다.
        모든 쪽에 글자가 있으면(머리말 등) 표시 번호 차이 방식으로 대체.
        """
        self.log("   🔍 분할 위치(미주) 탐색 중...")

        items = []
        ctrl = self.hwp.HeadCtrl
        while ctrl:
            items.append((ctrl.CtrlID, ctrl))
            ctrl = ctrl.Next
        en_idx = [i for i, (cid, _) in enumerate(items) if cid == "en"]
        if not en_idx:
            return None
        first_en = items[en_idx[0]][1]
        last_en = items[en_idx[-1]][1]
        boundary = next((c for i, (cid, c) in enumerate(items)
                         if cid in ("secd", "nwno") and i > en_idx[-1]), None)

        try:
            total_pages = self.hwp.PageCount
            last_body = None
            for pg in range(total_pages, 0, -1):
                if self._page_text(pg):
                    last_body = pg
                    break
            if last_body is None or last_body == total_pages:
                return self._split_by_span(first_en, last_en, boundary,
                                           total_pages)

            def is_cover(pg):
                head = self._page_text(pg)[:100]
                return any(k in head for k in self.COVER_KEYWORDS)

            if is_cover(last_body):
                # 답지 앞부분(표지·빠른정답 등)이 본문 텍스트로 여러 쪽
                # 이어질 수 있어, 키워드가 나오는 동안 최대 4쪽까지 소급
                split_page = last_body
                for pg in range(last_body - 1, max(0, last_body - 4), -1):
                    if is_cover(pg):
                        split_page = pg
                    else:
                        break
                basis = "답지 표지 기준"
            else:
                split_page = last_body + 1    # 본문 다음 쪽부터
                basis = "본문 끝 기준"
            self.log(f"   📍 분할 지점: {split_page}쪽 ({basis})")
            return split_page
        except Exception as e:
            self.log(f"   ⚠️ 미주 분석 중 오류: {e}")
            return None

    def _split_by_span(self, first_en, last_en, boundary, total_pages):
        """예비 방식: (전체 쪽수 - 답지 영역 쪽수 + 1). 영역 쪽수는
        표시 번호의 차이로 구하므로 '새 쪽번호' 오프셋이 상쇄된다."""
        try:
            end_shown = self._note_shown_page(last_en, at_end=True)
            if boundary is not None:
                self.hwp.SetPosBySet(boundary.GetAnchorPos(0))
                start_shown = self.hwp.KeyIndicator()[3]
            else:
                start_shown = self._note_shown_page(first_en)
            tail_pages = end_shown - start_shown + 1
            if not 1 <= tail_pages <= total_pages:
                raise ValueError(
                    f"답지 영역 계산 이상 (표시 번호 시작 {start_shown}쪽 / "
                    f"끝 {end_shown}쪽 / 전체 {total_pages}쪽)")
            split_page = total_pages - tail_pages + 1
            self.log(f"   📍 분할 지점: {split_page}쪽 (표시 번호 기준)")
            return split_page
        except Exception as e:
            self.log(f"   ⚠️ 미주 분석 중 오류: {e}")
            return None

    # ---------- 인쇄 ----------
    def _print_range(self, print_act, print_set, printer_name, page_range,
                     copies):
        """지정 범위를 지정 부수(Copies)만큼 한 번에 인쇄."""
        print_act.GetDefault(print_set)
        print_set.SetItem("Collate", 1)  # 한 부씩 정렬
        print_set.SetItem("UsingPagenum", 0)
        print_set.SetItem("PrinterName", printer_name)
        print_set.SetItem("Range", 4)
        print_set.SetItem("RangeCustom", page_range)
        print_set.SetItem("NumCopy", copies)

        print_act.Execute(print_set)
        time.sleep(0.5)  # 인쇄 작업 전송 대기

    def execute_print(self, split_page, total_pages, printer_body,
                      printer_answer, body_copies, answer_copies):
        """본문과 답지를 각각 해당 프린터로 보낸다."""
        try:
            print_act = self.hwp.CreateAction("Print")
            print_set = print_act.CreateSet()

            # 1) 본문: 1쪽 ~ (split_page - 1)쪽
            if split_page > 1 and body_copies > 0:
                body_range = f"1-{split_page - 1}"
                self._print_range(print_act, print_set, printer_body,
                                  body_range, body_copies)
                self.log(f"   🖨️ 본문 {body_range}쪽 · {body_copies}부 "
                         f"→ {printer_body}")
            elif body_copies == 0:
                self.log("   ⏭️ 본문 0부 — 건너뜀")

            # 2) 답지: split_page쪽 ~ 마지막쪽
            if split_page <= total_pages and answer_copies > 0:
                answer_range = f"{split_page}-{total_pages}"
                self._print_range(print_act, print_set, printer_answer,
                                  answer_range, answer_copies)
                self.log(f"   🖨️ 답지 {answer_range}쪽 · {answer_copies}부 "
                         f"→ {printer_answer}")
            elif answer_copies == 0:
                self.log("   ⏭️ 답지 0부 — 건너뜀")

            return True
        except Exception as e:
            self.log(f"   ❌ 인쇄 중 오류: {e}")
            return False

    # ---------- 단일 파일 처리 ----------
    def process_file(self, file_path, body_copies, answer_copies,
                     printer_body, printer_answer, on_no_endnote=None):
        """파일 하나를 열어 미주를 찾고 본문/답지를 인쇄한다.

        on_no_endnote: 미주가 없을 때 호출되는 콜백(filename) — 선택적.
        """
        try:
            if not self.hwp.Open(file_path, "", ""):
                self.log("   ❌ 파일을 열지 못했습니다.")
                return False

            total_pages = self.hwp.PageCount
            split_page = self.find_split_page_by_first_endnote()

            if split_page is None:
                self.log("   ⏭️ 미주 없음 — 인쇄하지 않고 건너뜁니다.")
                if on_no_endnote:
                    on_no_endnote(os.path.basename(file_path))
                self.hwp.Run("FileClose")
                return False

            result = self.execute_print(split_page, total_pages, printer_body,
                                        printer_answer, body_copies,
                                        answer_copies)
            self.hwp.Run("FileClose")
            return result
        except Exception as e:
            self.log(f"   ❌ 처리 중 오류: {e}")
            try:
                self.hwp.Run("FileClose")
            except Exception:
                pass
            return False


# ============================================================
# 탭 1: 파일 목록 모드 (파일별 부수 지정)
# ============================================================
class FileListTab(tk.Frame):
    """파일을 하나씩 추가하고 각 파일의 본문/답지 부수를 지정하는 탭."""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.file_entries = []  # [{'path', 'body_var', 'answer_var', 'row'}]
        self._build_ui()

    def _build_ui(self):
        frame = tk.LabelFrame(self, text="파일 목록", padx=8, pady=8)
        frame.pack(fill="both", expand=True, padx=8, pady=8)

        top = tk.Frame(frame)
        top.pack(fill="x", pady=(0, 5))
        tk.Button(top, text="파일 추가",
                  command=self.add_files).pack(side="left")
        tk.Button(top, text="전체 비우기",
                  command=self.clear_files).pack(side="left", padx=5)
        self.lbl_count = tk.Label(top, text="총 0개")
        self.lbl_count.pack(side="right")

        # 스크롤 가능한 파일 행 영역
        canvas_wrap = tk.Frame(frame)
        canvas_wrap.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(canvas_wrap, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_wrap, orient="vertical",
                                  command=self.canvas.yview)
        self.scroll_frame = tk.Frame(self.canvas)

        self.scroll_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scroll_frame,
                                  anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.canvas.bind("<Enter>",
                         lambda e: self.canvas.bind_all("<MouseWheel>",
                                                        self._on_mousewheel))
        self.canvas.bind("<Leave>",
                         lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ---------- 파일 추가/제거 ----------
    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="한글 파일 선택",
            filetypes=[("한글 문서", "*.hwp *.hwpx"), ("모든 파일", "*.*")])
        if not paths:
            return
        existing = {e['path'] for e in self.file_entries}
        added = 0
        for path in paths:
            if path in existing:
                continue
            self._add_file_row(path)
            existing.add(path)
            added += 1
        self._update_count()
        self.app.log(f"파일 {added}개 추가 (총 {len(self.file_entries)}개)")

    def _add_file_row(self, path):
        row = tk.Frame(self.scroll_frame, bd=1, relief="solid")
        row.pack(fill="x", padx=2, pady=1)

        name = os.path.basename(path)
        display = name if len(name) <= 42 else name[:39] + "..."
        tk.Label(row, text=display, anchor="w",
                 width=38).pack(side="left", padx=4, pady=3)

        body_var = tk.StringVar(value="")
        answer_var = tk.StringVar(value="")

        tk.Label(row, text="본문").pack(side="left")
        tk.Spinbox(row, from_=0, to=999, width=4,
                   textvariable=body_var).pack(side="left", padx=(2, 8))
        tk.Label(row, text="답지").pack(side="left")
        tk.Spinbox(row, from_=0, to=999, width=4,
                   textvariable=answer_var).pack(side="left", padx=2)

        entry = {'path': path, 'body_var': body_var,
                 'answer_var': answer_var, 'row': row}
        tk.Button(row, text="✕", width=3,
                  command=lambda: self._remove_entry(entry)).pack(side="right",
                                                                  padx=4)
        self.file_entries.append(entry)

    def _remove_entry(self, entry):
        entry['row'].destroy()
        if entry in self.file_entries:
            self.file_entries.remove(entry)
        self._update_count()

    def clear_files(self):
        for entry in self.file_entries:
            entry['row'].destroy()
        self.file_entries.clear()
        self._update_count()
        self.app.log("파일 목록을 비웠습니다.")

    def _update_count(self):
        self.lbl_count.config(text=f"총 {len(self.file_entries)}개")

    # ---------- 작업 데이터 수집 ----------
    def collect_jobs(self):
        """인쇄 작업 목록을 반환. 검증 실패 시 (None, 에러메시지) 반환."""
        if not self.file_entries:
            return None, "인쇄할 파일을 추가해주세요."

        unset_files = []
        for e in self.file_entries:
            body_str = e['body_var'].get().strip()
            answer_str = e['answer_var'].get().strip()
            if body_str == "" or answer_str == "":
                unset_files.append(os.path.basename(e['path']))

        if unset_files:
            listing = "\n".join(f"• {n}" for n in unset_files[:10])
            if len(unset_files) > 10:
                listing += f"\n... 외 {len(unset_files) - 10}개"
            return None, (
                f"다음 파일의 부수가 입력되지 않았습니다:\n\n{listing}\n\n"
                f"모든 파일의 본문·답지 부수를 입력한 후 다시 시도해주세요.")

        try:
            jobs = [(e['path'],
                     max(0, int(e['body_var'].get().strip())),
                     max(0, int(e['answer_var'].get().strip())))
                    for e in self.file_entries]
        except ValueError:
            return None, "부수 값에 숫자가 아닌 문자가 포함되어 있습니다."

        return jobs, None


# ============================================================
# 탭 2: 폴더 일괄 모드 (각 파일 1부씩)
# ============================================================
class FolderTab(tk.Frame):
    """폴더 안의 모든 hwp/hwpx 파일을 본문·답지 각 1부씩 일괄 인쇄하는 탭."""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self._build_ui()

    def _build_ui(self):
        frame = tk.LabelFrame(self, text="폴더 선택", padx=10, pady=10)
        frame.pack(fill="x", padx=8, pady=8)

        row = tk.Frame(frame)
        row.pack(fill="x")
        tk.Label(row, text="작업 폴더:").pack(side="left")
        self.ent_folder = tk.Entry(row)
        self.ent_folder.pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(row, text="찾아보기",
                  command=self.browse_folder).pack(side="left")

        info = tk.LabelFrame(self, text="안내", padx=10, pady=10)
        info.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        tk.Label(info, justify="left", anchor="w",
                 text=("• 선택한 폴더 안의 모든 .hwp / .hwpx 파일을 처리합니다.\n"
                       "• 각 파일은 본문 1부, 답지 1부씩 인쇄됩니다.\n"
                       "• 파일별로 부수를 다르게 지정하려면 "
                       "'파일 목록' 탭을 사용하세요.\n"
                       "• 미주가 없는 파일은 자동으로 건너뜁니다.")
                 ).pack(anchor="w")

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.ent_folder.delete(0, tk.END)
            self.ent_folder.insert(0, folder)

    def collect_jobs(self):
        """폴더 내 파일을 (경로, 본문 1부, 답지 1부) 튜플로 수집."""
        folder = self.ent_folder.get().strip()
        if not folder:
            return None, "작업 폴더를 선택해주세요."
        if not os.path.isdir(folder):
            return None, "선택한 폴더를 찾을 수 없습니다."

        files = [f for f in os.listdir(folder)
                 if f.lower().endswith((".hwp", ".hwpx"))]
        if not files:
            return None, "폴더 안에 .hwp / .hwpx 파일이 없습니다."

        jobs = [(os.path.join(folder, f), 1, 1) for f in sorted(files)]
        return jobs, None


# ============================================================
# 메인 애플리케이션
# ============================================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("한글 문서 본문/답지 분할 인쇄 도구")
        self.root.geometry("720x780")
        self.root.minsize(640, 660)

        self.stop_flag = False
        self.current_thread = None

        self._build_printer_frame()
        self._build_tabs()
        self._build_progress_frame()
        self._build_button_frame()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    # ---------- UI ----------
    def _build_printer_frame(self):
        self.printers = self.get_printer_list()

        frame = tk.LabelFrame(self.root, text="프린터 설정", padx=10, pady=8)
        frame.pack(fill="x", padx=10, pady=(10, 5))

        tk.Label(frame, text="본문 (A3):").grid(row=0, column=0, sticky="w",
                                              pady=2)
        self.cb_printer1 = ttk.Combobox(frame, values=self.printers, width=52)
        self.cb_printer1.grid(row=0, column=1, padx=5, pady=2)

        tk.Label(frame, text="답지 (A4):").grid(row=1, column=0, sticky="w",
                                              pady=2)
        self.cb_printer2 = ttk.Combobox(frame, values=self.printers, width=52)
        self.cb_printer2.grid(row=1, column=1, padx=5, pady=2)

        try:
            default_printer = win32print.GetDefaultPrinter()
        except Exception:
            default_printer = ""
        if default_printer in self.printers:
            self.cb_printer1.set(default_printer)
            self.cb_printer2.set(default_printer)

    def _build_tabs(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)

        self.tab_files = FileListTab(self.notebook, self)
        self.tab_folder = FolderTab(self.notebook, self)

        self.notebook.add(self.tab_files, text="  파일 목록 (개별 부수)  ")
        self.notebook.add(self.tab_folder, text="  폴더 일괄 (각 1부)  ")

    def _build_progress_frame(self):
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root,
                                            variable=self.progress_var,
                                            maximum=100)
        self.progress_bar.pack(fill="x", padx=10, pady=(8, 2))

        self.lbl_status = tk.Label(self.root, text="대기 중...", fg="blue")
        self.lbl_status.pack()

        log_wrap = tk.Frame(self.root)
        log_wrap.pack(fill="x", padx=10, pady=5)
        self.txt_log = tk.Text(log_wrap, height=9, state="disabled",
                               bg="#f8f9fa", font=("Consolas", 9))
        log_scroll = ttk.Scrollbar(log_wrap, orient="vertical",
                                   command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=log_scroll.set)
        self.txt_log.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

    def _build_button_frame(self):
        frame = tk.Frame(self.root)
        frame.pack(fill="x", padx=10, pady=(0, 10))

        self.btn_start = tk.Button(frame, text="인쇄 시작",
                                   command=self.start_thread,
                                   bg="#2ecc71", fg="white", height=2,
                                   font=("맑은 고딕", 10, "bold"))
        self.btn_start.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.btn_stop = tk.Button(frame, text="중단", command=self.stop_process,
                                  bg="#e74c3c", fg="white", height=2,
                                  font=("맑은 고딕", 10, "bold"),
                                  state="disabled")
        self.btn_stop.pack(side="left", fill="x", expand=True, padx=(5, 0))

    # ---------- 유틸 ----------
    def log(self, message):
        def _log():
            self.txt_log.config(state="normal")
            self.txt_log.insert(tk.END,
                                f"[{time.strftime('%H:%M:%S')}] {message}\n")
            self.txt_log.see(tk.END)
            self.txt_log.config(state="disabled")
        self.root.after(0, _log)

    def update_progress(self, value):
        self.root.after(0, lambda: self.progress_var.set(value))

    def update_status(self, text, color="blue"):
        self.root.after(0, lambda: self.lbl_status.config(text=text,
                                                          fg=color))

    def get_printer_list(self):
        try:
            return [p[2] for p in win32print.EnumPrinters(2)]
        except Exception:
            return ["프린터를 찾을 수 없습니다"]

    # ---------- 실행 제어 ----------
    def start_thread(self):
        p1 = self.cb_printer1.get().strip()
        p2 = self.cb_printer2.get().strip()
        if not p1 or not p2:
            messagebox.showwarning("입력 오류",
                                   "본문/답지 프린터를 모두 선택해주세요.")
            return

        # 현재 활성화된 탭에서 작업 데이터 수집
        current_tab = self.notebook.nametowidget(self.notebook.select())
        jobs, err = current_tab.collect_jobs()
        if err:
            messagebox.showwarning("입력 오류", err)
            return

        self.stop_flag = False
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.current_thread = threading.Thread(
            target=self.work_process, args=(jobs, p1, p2), daemon=True)
        self.current_thread.start()

    def stop_process(self):
        self.stop_flag = True
        self.log("⚠️ 중단 요청됨 — 현재 파일까지 마치고 멈춥니다.")

    def on_closing(self):
        self.stop_flag = True
        self.root.destroy()

    # ---------- 메인 로직 ----------
    def work_process(self, jobs, printer_body, printer_answer):
        pythoncom.CoInitialize()
        printer = HwpPrinter(self.log)
        skipped_files = []

        def on_no_endnote(filename):
            skipped_files.append(filename)

        try:
            printer.start()

            total = len(jobs)
            success = 0
            for i, (path, body_copies, answer_copies) in enumerate(jobs):
                if self.stop_flag:
                    self.log("⚠️ 사용자 요청으로 중단했습니다.")
                    break

                filename = os.path.basename(path)
                self.update_progress((i / total) * 100)
                self.update_status(f"처리 중: {filename}")
                self.log(f"📄 [{i + 1}/{total}] {filename} "
                         f"(본문 {body_copies}부 / 답지 {answer_copies}부)")

                if printer.process_file(path, body_copies, answer_copies,
                                        printer_body, printer_answer,
                                        on_no_endnote=on_no_endnote):
                    success += 1

            self.update_progress(100)
            self.update_status("작업 완료", "green")

            # 완료 메시지
            msg = f"처리 완료 — 성공 {success} / 전체 {total}건"
            if skipped_files:
                preview = "\n".join(f"• {n}" for n in skipped_files[:10])
                if len(skipped_files) > 10:
                    preview += f"\n... 외 {len(skipped_files) - 10}개"
                msg += (f"\n\n미주가 없어 건너뛴 파일 "
                        f"{len(skipped_files)}개:\n{preview}")
            self.root.after(0,
                            lambda: messagebox.showinfo("완료", msg))
        except Exception as e:
            self.log(f"❌ 오류: {e}")
            self.update_status("오류 발생", "red")
        finally:
            printer.quit()
            pythoncom.CoUninitialize()
            self.root.after(0,
                            lambda: self.btn_start.config(state="normal"))
            self.root.after(0,
                            lambda: self.btn_stop.config(state="disabled"))


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()