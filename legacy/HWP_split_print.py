import os
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import win32com.client
import win32print
import pythoncom


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("한글 문서 본문/답지 분할 인쇄 도구")
        self.root.geometry("680x740")
        self.root.minsize(620, 620)

        self.stop_flag = False
        self.current_thread = None
        self.file_entries = []  # [{'path', 'body_var', 'answer_var', 'row'}]

        self._build_printer_frame()
        self._build_file_list_frame()
        self._build_progress_frame()
        self._build_button_frame()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    # ==================== UI 구성 ====================
    def _build_printer_frame(self):
        self.printers = self.get_printer_list()

        frame = tk.LabelFrame(self.root, text="프린터 설정", padx=10, pady=8)
        frame.pack(fill="x", padx=10, pady=(10, 5))

        tk.Label(frame, text="본문 (A3):").grid(row=0, column=0, sticky="w", pady=2)
        self.cb_printer1 = ttk.Combobox(frame, values=self.printers, width=48)
        self.cb_printer1.grid(row=0, column=1, padx=5, pady=2)

        tk.Label(frame, text="답지 (A4):").grid(row=1, column=0, sticky="w", pady=2)
        self.cb_printer2 = ttk.Combobox(frame, values=self.printers, width=48)
        self.cb_printer2.grid(row=1, column=1, padx=5, pady=2)

        try:
            default_printer = win32print.GetDefaultPrinter()
        except Exception:
            default_printer = ""
        if default_printer in self.printers:
            self.cb_printer1.set(default_printer)
            self.cb_printer2.set(default_printer)

    def _build_file_list_frame(self):
        frame = tk.LabelFrame(self.root, text="파일 목록", padx=8, pady=8)
        frame.pack(fill="both", expand=True, padx=10, pady=5)

        top = tk.Frame(frame)
        top.pack(fill="x", pady=(0, 5))
        tk.Button(top, text="파일 추가", command=self.add_files).pack(side="left")
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
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _build_progress_frame(self):
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var,
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

        self.btn_start = tk.Button(frame, text="인쇄 시작", command=self.start_thread,
                                   bg="#2ecc71", fg="white", height=2,
                                   font=("맑은 고딕", 10, "bold"))
        self.btn_start.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.btn_stop = tk.Button(frame, text="중단", command=self.stop_process,
                                  bg="#e74c3c", fg="white", height=2,
                                  font=("맑은 고딕", 10, "bold"), state="disabled")
        self.btn_stop.pack(side="left", fill="x", expand=True, padx=(5, 0))

    # ==================== 유틸리티 ====================
    def log(self, message):
        def _log():
            self.txt_log.config(state="normal")
            self.txt_log.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\n")
            self.txt_log.see(tk.END)
            self.txt_log.config(state="disabled")
        self.root.after(0, _log)

    def update_progress(self, value):
        self.root.after(0, lambda: self.progress_var.set(value))

    def update_status(self, text, color="blue"):
        self.root.after(0, lambda: self.lbl_status.config(text=text, fg=color))

    def get_printer_list(self):
        try:
            return [p[2] for p in win32print.EnumPrinters(2)]
        except Exception:
            return ["프린터를 찾을 수 없습니다"]

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ==================== 파일 목록 ====================
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
        self.log(f"파일 {added}개 추가 (총 {len(self.file_entries)}개)")

    def _add_file_row(self, path):
        row = tk.Frame(self.scroll_frame, bd=1, relief="solid")
        row.pack(fill="x", padx=2, pady=1)

        name = os.path.basename(path)
        display = name if len(name) <= 34 else name[:31] + "..."
        tk.Label(row, text=display, anchor="w",
                 width=32).pack(side="left", padx=4, pady=3)

        split_var = tk.BooleanVar(value=True)
        tk.Checkbutton(row, text="분할",
                       variable=split_var).pack(side="left", padx=(2, 6))

        body_var = tk.StringVar(value="")
        answer_var = tk.StringVar(value="")

        tk.Label(row, text="본문").pack(side="left")
        tk.Spinbox(row, from_=0, to=999, width=4,
                   textvariable=body_var).pack(side="left", padx=(2, 8))
        tk.Label(row, text="답지").pack(side="left")
        tk.Spinbox(row, from_=0, to=999, width=4,
                   textvariable=answer_var).pack(side="left", padx=2)

        entry = {'path': path, 'body_var': body_var,
                 'answer_var': answer_var, 'split_var': split_var, 'row': row}
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
        self.log("파일 목록을 비웠습니다.")

    def _update_count(self):
        self.lbl_count.config(text=f"총 {len(self.file_entries)}개")

    # ==================== 실행 제어 ====================
    def start_thread(self):
        p1 = self.cb_printer1.get().strip()
        p2 = self.cb_printer2.get().strip()
        if not self.file_entries:
            messagebox.showwarning("입력 오류", "인쇄할 파일을 추가해주세요.")
            return
        if not p1 or not p2:
            messagebox.showwarning("입력 오류", "본문/답지 프린터를 모두 선택해주세요.")
            return

        # 부수 미입력 검사
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
            messagebox.showwarning(
                "부수 미설정",
                f"다음 파일의 부수가 입력되지 않았습니다:\n\n{listing}\n\n"
                f"모든 파일의 본문·답지 부수를 입력한 후 다시 시도해주세요.")
            return

        # 작업 데이터를 메인 스레드에서 미리 수집 (tkinter 변수는 워커에서 접근 X)
        try:
            jobs = [(e['path'],
                     max(0, int(e['body_var'].get().strip())),
                     max(0, int(e['answer_var'].get().strip())),
                     e['split_var'].get())
                    for e in self.file_entries]
        except ValueError:
            messagebox.showwarning(
                "입력 오류",
                "부수 값에 숫자가 아닌 문자가 포함되어 있습니다. 확인 후 다시 시도해주세요.")
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

    # ==================== 메인 로직 ====================
    def work_process(self, jobs, printer_body, printer_answer):
        pythoncom.CoInitialize()
        hwp = None
        try:
            hwp = win32com.client.dynamic.Dispatch("HwpFrame.HwpObject.2")
            hwp.XHwpWindows.Item(0).Visible = True

            total = len(jobs)
            success = 0
            for i, (path, body_copies, answer_copies, do_split) in enumerate(jobs):
                if self.stop_flag:
                    self.log("⚠️ 사용자 요청으로 중단했습니다.")
                    break

                filename = os.path.basename(path)
                mode_label = "분할" if do_split else "전체"
                self.update_progress((i / total) * 100)
                self.update_status(f"처리 중: {filename}")
                self.log(f"📄 [{i + 1}/{total}] {filename} "
                         f"[{mode_label}] "
                         f"(본문 {body_copies}부 / 답지 {answer_copies}부)")

                if self.process_single_file(hwp, path, body_copies,
                                            answer_copies, do_split,
                                            printer_body, printer_answer):
                    success += 1

            self.update_progress(100)
            self.update_status("작업 완료", "green")
            self.root.after(
                0, lambda: messagebox.showinfo(
                    "완료", f"처리 완료 — 성공 {success} / 전체 {total}건"))
        except Exception as e:
            self.log(f"❌ 오류: {e}")
            self.update_status("오류 발생", "red")
        finally:
            if hwp is not None:
                try:
                    hwp.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self.btn_start.config(state="normal"))
            self.root.after(0, lambda: self.btn_stop.config(state="disabled"))

    def process_single_file(self, hwp, file_path, body_copies, answer_copies,
                            do_split, printer_body, printer_answer):
        try:
            if not hwp.Open(file_path, "", ""):
                self.log("   ❌ 파일을 열지 못했습니다.")
                return False

            total_pages = hwp.PageCount

            if do_split:
                split_page = self.find_split_page_by_first_endnote(hwp, total_pages)

                # 분할 모드인데 미주를 못 찾으면 인쇄하지 않고 안내 후 스킵
                if split_page is None:
                    filename = os.path.basename(file_path)
                    self.log("   ⏭️ 미주 없음 — 인쇄하지 않고 건너뜁니다.")
                    self.root.after(0, lambda: messagebox.showwarning(
                        "미주 없음",
                        f"'{filename}'\n\n"
                        f"분할 모드인데 미주를 찾지 못해 본문/답지를 나눌 수 없습니다.\n"
                        f"이 파일은 건너뜁니다.\n\n"
                        f"전체를 통째로 인쇄하려면 해당 파일의 '분할' 체크를 해제하세요."))
                    hwp.Run("FileClose")
                    return False

                result = self.execute_print(hwp, split_page, total_pages,
                                            printer_body, printer_answer,
                                            body_copies, answer_copies)
            else:
                # 전체 인쇄 모드 — 미주 분할 없이 전체 페이지를 양쪽 프린터로
                self.log(f"   📑 전체 인쇄 모드 ({total_pages}쪽)")
                result = self.execute_whole_print(hwp, total_pages,
                                                  printer_body, printer_answer,
                                                  body_copies, answer_copies)

            hwp.Run("FileClose")
            return result
        except Exception as e:
            self.log(f"   ❌ 처리 중 오류: {e}")
            try:
                hwp.Run("FileClose")
            except Exception:
                pass
            return False

    def find_split_page_by_first_endnote(self, hwp, total_pages):
        """문서의 첫 번째 미주를 찾아 그 페이지 번호를 반환 (없으면 None)."""
        self.log("   🔍 분할 위치(미주) 탐색 중...")
        split_page = None

        ctrl = hwp.HeadCtrl
        while ctrl:
            if ctrl.CtrlID == "en":  # 미주 컨트롤
                try:
                    hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                    hwp.Run("NoteModify")
                    info = hwp.KeyIndicator()
                    split_page = info[3]  # 페이지 번호
                    hwp.Run("CloseUpper")  # 본문으로 복귀
                    self.log(f"   📍 분할 지점: {split_page}쪽")
                    break
                except Exception as e:
                    self.log(f"   ⚠️ 미주 분석 중 오류: {e}")
            ctrl = ctrl.Next

        return split_page

    def _print_range(self, print_act, print_set, printer_name, page_range,
                     copies):
        """지정 범위를 지정 부수(Copies)만큼 한 번에 인쇄."""
        print_act.GetDefault(print_set)
        print_set.SetItem("UsingPagenum", 0)
        print_set.SetItem("PrinterName", printer_name)
        print_set.SetItem("Range", 4)
        print_set.SetItem("RangeCustom", page_range)
        print_set.SetItem("NumCopy", copies)
        print_set.SetItem("Collate", 1)  # 한 부씩 정렬(1,2,3 / 1,2,3)
        print_act.Execute(print_set)
        time.sleep(0.5)  # 인쇄 작업 전송 대기

    def execute_print(self, hwp, split_page, total_pages, printer_body,
                      printer_answer, body_copies, answer_copies):
        try:
            print_act = hwp.CreateAction("Print")
            print_set = print_act.CreateSet()

            # 1) 본문: 1쪽 ~ (split_page - 1)쪽
            if split_page > 1 and body_copies > 0:
                body_range = f"1-{split_page - 1}"
                self._print_range(print_act, print_set, printer_body,
                                  body_range, body_copies)
                self.log(f"   🖨️ 본문 {body_range}쪽 · {body_copies}부 → {printer_body}")
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

    def execute_whole_print(self, hwp, total_pages, printer_body, printer_answer,
                            body_copies, answer_copies):
        """미주 분할 없이 전체 페이지를 양쪽 프린터로 지정 부수만큼 인쇄."""
        try:
            print_act = hwp.CreateAction("Print")
            print_set = print_act.CreateSet()
            whole_range = f"1-{total_pages}"

            # 본문 프린터로 전체 N부
            if body_copies > 0:
                self._print_range(print_act, print_set, printer_body,
                                  whole_range, body_copies)
                self.log(f"   🖨️ 전체 {whole_range}쪽 · {body_copies}부 "
                         f"→ {printer_body}")
            else:
                self.log("   ⏭️ 본문 0부 — 본문 프린터 건너뜀")

            # 답지 프린터로 전체 M부
            if answer_copies > 0:
                self._print_range(print_act, print_set, printer_answer,
                                  whole_range, answer_copies)
                self.log(f"   🖨️ 전체 {whole_range}쪽 · {answer_copies}부 "
                         f"→ {printer_answer}")
            else:
                self.log("   ⏭️ 답지 0부 — 답지 프린터 건너뜀")

            return True
        except Exception as e:
            self.log(f"   ❌ 인쇄 중 오류: {e}")
            return False


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()