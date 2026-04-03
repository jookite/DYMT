import win32com.client
import os
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32print
import pythoncom

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("오답노트 자동 분할 인쇄 도구")
        self.root.geometry("600x650")

        self.stop_flag = False
        self.current_thread = None

        # --- UI 레이아웃 ---
        frame_folder = tk.Frame(root)
        frame_folder.pack(fill="x", padx=10, pady=5)
        tk.Label(frame_folder, text="작업 폴더:").pack(side="left")
        self.ent_folder = tk.Entry(frame_folder, width=45)
        self.ent_folder.pack(side="left", padx=5)
        tk.Button(frame_folder, text="찾아보기", command=self.browse_folder).pack(side="left")

        self.printers = self.get_printer_list()

        frame_printer = tk.LabelFrame(root, text="프린터 설정", padx=10, pady=10)
        frame_printer.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_printer, text="본문 (A3):").grid(row=0, column=0, sticky="w")
        self.cb_printer1 = ttk.Combobox(frame_printer, values=self.printers, width=50)
        self.cb_printer1.grid(row=0, column=1, pady=2)

        tk.Label(frame_printer, text="답지 (A4):").grid(row=1, column=0, sticky="w")
        self.cb_printer2 = ttk.Combobox(frame_printer, values=self.printers, width=50)
        self.cb_printer2.grid(row=1, column=1, pady=2)

        default_printer = win32print.GetDefaultPrinter()
        if default_printer in self.printers:
            self.cb_printer1.set(default_printer)
            self.cb_printer2.set(default_printer)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(root, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", padx=10, pady=10)

        self.lbl_status = tk.Label(root, text="대기 중...", fg="blue")
        self.lbl_status.pack()

        self.txt_log = tk.Text(root, height=18, state="disabled", bg="#f8f9fa", font=("Consolas", 9))
        self.txt_log.pack(fill="both", padx=10, pady=5, expand=True)

        btn_frame = tk.Frame(root)
        btn_frame.pack(fill="x", padx=10, pady=5)

        self.btn_start = tk.Button(btn_frame, text="인쇄 시작", command=self.start_thread,
                                   bg="#2ecc71", fg="white", height=2, font=("맑은 고딕", 10, "bold"))
        self.btn_start.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.btn_stop = tk.Button(btn_frame, text="중단", command=self.stop_process,
                                  bg="#e74c3c", fg="white", height=2, font=("맑은 고딕", 10, "bold"),
                                  state="disabled")
        self.btn_stop.pack(side="left", fill="x", expand=True, padx=(5, 0))

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    # --- 유틸리티 메서드 ---
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
        except:
            return ["프린터를 찾을 수 없습니다"]

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.ent_folder.delete(0, tk.END)
            self.ent_folder.insert(0, folder)

    def start_thread(self):
        folder = self.ent_folder.get()
        p1 = self.cb_printer1.get()
        p2 = self.cb_printer2.get()
        if not folder or not p1 or not p2:
            messagebox.showwarning("입력 오류", "모든 설정을 완료해주세요.")
            return
        self.stop_flag = False
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.current_thread = threading.Thread(target=self.work_process, args=(folder, p1, p2), daemon=True)
        self.current_thread.start()

    def stop_process(self):
        self.stop_flag = True
        self.log("⚠️ 중단 요청됨...")

    def on_closing(self):
        self.stop_flag = True
        self.root.destroy()

    # --- 메인 로직 ---
    def work_process(self, MY_FOLDER, PRINTER_NAME1, PRINTER_NAME2):
        pythoncom.CoInitialize()
        hwp = None
        try:
            hwp = win32com.client.dynamic.Dispatch("HwpFrame.HwpObject.2")
            hwp.XHwpWindows.Item(0).Visible = True

            files = [f for f in os.listdir(MY_FOLDER) if f.lower().endswith((".hwp", ".hwpx"))]
            if not files:
                self.log("⚠️ 파일이 없습니다.")
                return

            success_count = 0
            for i, filename in enumerate(files):
                if self.stop_flag: break

                self.update_progress((i / len(files)) * 100)
                self.update_status(f"처리 중: {filename}")
                self.log(f"📄 [{i+1}/{len(files)}] {filename}")

                file_path = os.path.join(MY_FOLDER, filename)
                if self.process_single_file(hwp, file_path, PRINTER_NAME1, PRINTER_NAME2):
                    success_count += 1

            self.update_progress(100)
            self.update_status("작업 완료", "green")
            messagebox.showinfo("완료", f"성공: {success_count}건")
        except Exception as e:
            self.log(f"❌ 오류: {e}")
        finally:
            if hwp: hwp.Quit()
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self.btn_start.config(state="normal"))
            self.root.after(0, lambda: self.btn_stop.config(state="disabled"))

    def process_single_file(self, hwp, file_path, p1, p2):
        try:
            if not hwp.Open(file_path, "", ""): return False
            total_pages = hwp.PageCount

            # [핵심] 1번 미주를 찾아 분할 페이지 계산
            split_page = self.find_split_page_by_first_endnote(hwp, total_pages)

            # 인쇄 실행
            res = self.execute_print(hwp, split_page, total_pages, p1, p2)
            hwp.Run("FileClose")
            return res
        except:
            return False

    def find_split_page_by_first_endnote(self, hwp, total_pages):
        """문서에서 1번 미주를 찾아 해당 페이지 번호를 반환"""
        self.log("   🔍 분할 위치 탐색 중...")
        split_page = total_pages + 1

        # 문서의 처음 컨트롤부터 정방향 탐색
        ctrl = hwp.HeadCtrl
        found_en1 = False

        while ctrl:
            if ctrl.CtrlID == "en":  # 미주 컨트롤 발견
                # 컨트롤의 속성에서 번호(Number) 확인
                try:
                    # pset = ctrl.Properties
                    # note_num = pset.Item("Number")

                    # 가장 확실한 방법: 일단 이동 후 NoteModify 진입하여 확인
                    hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                    hwp.Run("NoteModify")

                    # 미주 내부로 들어왔을 때 해당 미주의 번호를 확인하거나
                    # 정방향 탐색의 첫 번째 미주를 1번으로 간주
                    info = hwp.KeyIndicator()
                    split_page = info[3] # 페이지 번호 추출

                    hwp.Run("CloseUpper") # 본문으로 탈출

                    self.log(f"   📍 분할 지점 페이지: {split_page}쪽")
                    found_en1 = True
                    break
                except Exception as e:
                    self.log(f"   ⚠️ 컨트롤 분석 중 오류: {e}")

            ctrl = ctrl.Next

        if not found_en1:
            self.log("   ℹ️ 1번 미주를 찾지 못했습니다. 전체 페이지를 본문으로 인쇄합니다.")
            return total_pages + 1

        return split_page

    def execute_print(self, hwp, split_page, total_pages, PRINTER_NAME1, PRINTER_NAME2):
        try:
            print_act = hwp.CreateAction("Print")
            print_set = print_act.CreateSet()
            print_act.GetDefault(print_set)

            # 1. 본문 인쇄 (1쪽 ~ 1번 미주 전 페이지)
            if split_page > 1:
                body_end = split_page - 1
                body_range = f"1-{body_end}"
                print_set.SetItem("UsingPagenum", 0)
                print_set.SetItem("PrinterName", PRINTER_NAME1)
                print_set.SetItem("Range", 4)
                print_set.SetItem("RangeCustom", body_range)
                print_act.Execute(print_set)
                self.log(f"   🖨️ 본문 인쇄: {body_range}쪽 -> {PRINTER_NAME1}")
                # time.sleep(1.0) # 인쇄 작업 전송 대기

            # 2. 답지 인쇄 (1번 미주 페이지 ~ 마지막 페이지)
            if split_page <= total_pages:
                answer_range = f"{split_page}-{total_pages}"
                print_set.SetItem("UsingPagenum", 0)
                print_set.SetItem("PrinterName", PRINTER_NAME2)
                print_set.SetItem("Range", 4)
                print_set.SetItem("RangeCustom", answer_range)
                print_act.Execute(print_set)
                self.log(f"   🖨️ 답지 인쇄: {answer_range}쪽 -> {PRINTER_NAME2}")

            return True
        except Exception as e:
            self.log(f"   ❌ 인쇄 중 오류: {e}")
            return False

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()