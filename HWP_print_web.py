"""
한글 문서 본문/답지 분할 인쇄 도구 — 다중 사용자 웹 최종판 (NiceGUI)

이 스크립트를 실행하는 Windows PC(한글 + 프린터 연결)가 '인쇄 서버'가 되고,
같은 네트워크의 여러 사람이 각자 브라우저로 접속해 동시에 사용할 수 있습니다.

  실행:        python HWP_print_web.py
  이 PC에서:   http://localhost:8080
  다른 기기:   http://<이 PC의 IP>:8080  (화면 상단에 주소 표시)

구조
  - 접속자마다 자기만의 화면(파일 목록)을 가짐 — 서로 안 섞임
  - '인쇄 시작'을 누르면 작업이 공용 대기열에 들어가고,
    서버가 한 건씩 순서대로 인쇄 (한글 COM은 동시에 하나만 가능)
  - 진행 상황·로그는 모두에게 실시간 공유

기능
  - 여러 파일 업로드, 파일별 분할 여부·본문/답지 부수 지정
  - 분할 모드: 첫 미주 페이지 기준으로 본문/답지를 나눠 A3/A4 프린터로 인쇄
  - 전체 모드: 분할 없이 전체를 양쪽 프린터로 지정 부수만큼 인쇄

주의: 실제 인쇄는 서버 PC에서 일어납니다. 한글 보안 팝업이 뜨면
서버 PC 화면에서 확인해야 합니다. 사내(같은 공유기) 네트워크 전용이며
인터넷에 직접 공개하지 마세요.
"""

import os
import queue
import socket
import tempfile
import threading
import time
import uuid
from collections import deque
from pathlib import Path

import pythoncom
import win32com.client
import win32print

from nicegui import ui

UPLOAD_DIR = Path(tempfile.gettempdir()) / "hwp_print_web_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

PORT = 8080


# ============================================================
# 공통 HWP 처리 로직
# ============================================================
class HwpPrinter:
    """한글 문서 열기, 미주 탐색, 본문/답지 분할 인쇄 담당."""

    def __init__(self, log_fn):
        """log_fn: 메시지를 UI로 전달할 콜백."""
        self.log = log_fn
        self.hwp = None

    def start(self):
        """HwpObject 초기화."""
        self.hwp = win32com.client.dynamic.Dispatch("HwpFrame.HwpObject.2")
        try:
            # 보안 모듈이 레지스트리에 등록되어 있으면 파일 접근 팝업을 억제
            self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
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
    def find_split_page_by_first_endnote(self):
        """문서의 첫 번째 미주를 찾아 그 페이지 번호를 반환 (없으면 None)."""
        self.log("   🔍 분할 위치(미주) 탐색 중...")
        split_page = None

        ctrl = self.hwp.HeadCtrl
        while ctrl:
            if ctrl.CtrlID == "en":  # 미주 컨트롤
                try:
                    self.hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                    self.hwp.Run("NoteModify")
                    info = self.hwp.KeyIndicator()
                    split_page = info[3]  # 페이지 번호
                    self.hwp.Run("CloseUpper")  # 본문으로 복귀
                    self.log(f"   📍 분할 지점: {split_page}쪽")
                    break
                except Exception as e:
                    self.log(f"   ⚠️ 미주 분석 중 오류: {e}")
            ctrl = ctrl.Next

        return split_page

    # ---------- 인쇄 ----------
    def _print_range(self, print_act, print_set, printer_name, page_range,
                     copies):
        """지정 범위를 지정 부수(NumCopy)만큼 한 번에 인쇄."""
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

    def execute_whole_print(self, total_pages, printer_body, printer_answer,
                            body_copies, answer_copies):
        """미주 분할 없이 전체 페이지를 양쪽 프린터로 지정 부수만큼 인쇄."""
        try:
            print_act = self.hwp.CreateAction("Print")
            print_set = print_act.CreateSet()
            whole_range = f"1-{total_pages}"

            if body_copies > 0:
                self._print_range(print_act, print_set, printer_body,
                                  whole_range, body_copies)
                self.log(f"   🖨️ 전체 {whole_range}쪽 · {body_copies}부 "
                         f"→ {printer_body}")
            else:
                self.log("   ⏭️ 본문 0부 — 본문 프린터 건너뜀")

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

    # ---------- 단일 파일 처리 ----------
    def process_file(self, file_path, body_copies, answer_copies, do_split,
                     printer_body, printer_answer, on_no_endnote=None):
        """파일 하나를 열어 분할/전체 모드에 따라 인쇄한다."""
        try:
            if not self.hwp.Open(file_path, "", ""):
                self.log("   ❌ 파일을 열지 못했습니다.")
                return False

            total_pages = self.hwp.PageCount

            if do_split:
                split_page = self.find_split_page_by_first_endnote()

                # 분할 모드인데 미주를 못 찾으면 인쇄하지 않고 건너뜀
                if split_page is None:
                    self.log("   ⏭️ 미주 없음 — 인쇄하지 않고 건너뜁니다. "
                             "(전체 인쇄하려면 '분할'을 끄세요)")
                    if on_no_endnote:
                        on_no_endnote(os.path.basename(file_path))
                    self.hwp.Run("FileClose")
                    return False

                result = self.execute_print(split_page, total_pages,
                                            printer_body, printer_answer,
                                            body_copies, answer_copies)
            else:
                self.log(f"   📑 전체 인쇄 모드 ({total_pages}쪽)")
                result = self.execute_whole_print(total_pages, printer_body,
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
# 전역 상태 — 공용 로그 / 대기열 / 진행 상황
# ============================================================
class GlobalLog:
    """모든 접속자에게 공유되는 로그. 각 화면이 못 본 줄만 가져간다."""

    def __init__(self):
        self.lines = deque(maxlen=2000)
        self.total = 0  # 지금까지 쌓인 총 줄 수 (트리밍과 무관)

    def add(self, message):
        self.lines.append(f"[{time.strftime('%H:%M:%S')}] {message}")
        self.total += 1

    def new_lines_since(self, seen_total):
        n = self.total - seen_total
        if n <= 0:
            return []
        return list(self.lines)[-min(n, len(self.lines)):]


class AppState:
    def __init__(self):
        self.running = False
        self.stop_flag = False
        self.progress = 0.0          # 현재 작업(배치) 기준 0.0 ~ 1.0
        self.status = "대기 중..."


glog = GlobalLog()
state = AppState()
job_queue = queue.Queue()  # 항목: {'label', 'jobs', 'printer_body', 'printer_answer'}

log = glog.add


def get_printer_list():
    try:
        return [p[2] for p in win32print.EnumPrinters(2)]
    except Exception:
        return []


def get_local_ip():
    """같은 네트워크의 다른 기기에서 접속할 때 쓸 이 PC의 IP."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


# ============================================================
# 인쇄 워커 — 서버에 하나만 떠서 대기열을 순서대로 처리
# ============================================================
def worker_loop():
    pythoncom.CoInitialize()
    while True:
        batch = job_queue.get()  # 작업이 들어올 때까지 대기
        state.running = True
        state.stop_flag = False
        state.progress = 0.0

        jobs = batch["jobs"]
        label = batch["label"]
        printer = HwpPrinter(log)
        skipped_files = []

        try:
            log(f"▶️ '{label}' 작업 시작 — 파일 {len(jobs)}건 "
                f"(대기 중 {job_queue.qsize()}건)")
            printer.start()

            total = len(jobs)
            success = 0
            for i, (path, name, body_copies, answer_copies,
                    do_split) in enumerate(jobs):
                if state.stop_flag:
                    log("⚠️ 중단 요청으로 현재 작업을 멈췄습니다.")
                    break

                mode_label = "분할" if do_split else "전체"
                state.progress = i / total
                state.status = f"[{label}] 처리 중: {name}"
                log(f"📄 [{i + 1}/{total}] {name} [{mode_label}] "
                    f"(본문 {body_copies}부 / 답지 {answer_copies}부)")

                if printer.process_file(path, body_copies, answer_copies,
                                        do_split, batch["printer_body"],
                                        batch["printer_answer"],
                                        on_no_endnote=skipped_files.append):
                    success += 1

            state.progress = 1.0
            log(f"✅ '{label}' 완료 — 성공 {success} / 전체 {total}건")
            if skipped_files:
                log(f"⏭️ 미주가 없어 건너뛴 파일 {len(skipped_files)}개: "
                    + ", ".join(skipped_files))
        except Exception as e:
            log(f"❌ '{label}' 처리 중 오류: {e}")
        finally:
            printer.quit()
            for path, *_ in jobs:  # 업로드 임시 파일 정리
                try:
                    Path(path).unlink(missing_ok=True)
                except Exception:
                    pass
            state.running = False
            state.status = ("대기 중..." if job_queue.empty()
                            else f"다음 작업 준비 중 (대기 {job_queue.qsize()}건)")
            job_queue.task_done()


threading.Thread(target=worker_loop, daemon=True).start()


def stop_all():
    """현재 작업을 중단하고 대기열도 비운다."""
    drained = 0
    while True:
        try:
            batch = job_queue.get_nowait()
        except queue.Empty:
            break
        for path, *_ in batch["jobs"]:
            try:
                Path(path).unlink(missing_ok=True)
            except Exception:
                pass
        job_queue.task_done()
        drained += 1
    if state.running:
        state.stop_flag = True
    if drained or state.running:
        log(f"🛑 중단 요청 — 대기열 {drained}건 취소, "
            f"진행 중이던 작업은 현재 파일까지 마치고 멈춥니다.")


# ============================================================
# UI — 접속자(브라우저)마다 개별 화면
# ============================================================
printers = get_printer_list()
try:
    default_printer = win32print.GetDefaultPrinter()
except Exception:
    default_printer = None
if default_printer not in printers:
    default_printer = None

server_ip = get_local_ip()


@ui.page("/")
def index():
    file_entries = []      # 이 접속자만의 파일 목록
    seen = {"n": glog.total - len(glog.lines)}  # 로그는 최근 것부터 표시

    with ui.column().classes("w-full max-w-3xl mx-auto p-4 gap-3"):
        ui.label("한글 문서 본문/답지 분할 인쇄").classes("text-xl font-bold")
        ui.label(f"공용 인쇄 서버 — 다른 기기 접속: http://{server_ip}:{PORT}") \
            .classes("text-sm text-gray-500")

        # ---------- 이름 / 프린터 ----------
        with ui.card().classes("w-full"):
            with ui.row().classes("w-full items-center"):
                name_input = ui.input("이름 (로그에 표시)",
                                      placeholder="예: 홍길동") \
                    .classes("w-40").props("outlined dense")
                sel_body = ui.select(printers, label="본문 (A3) 프린터",
                                     value=default_printer) \
                    .classes("flex-1").props("outlined dense")
                sel_answer = ui.select(printers, label="답지 (A4) 프린터",
                                       value=default_printer) \
                    .classes("flex-1").props("outlined dense")

        # ---------- 파일 업로드 ----------
        with ui.card().classes("w-full"):
            ui.label("파일 목록 — 파일별로 분할 여부와 부수를 정하세요") \
                .classes("font-bold")

            def handle_upload(e):
                name = e.name
                if not name.lower().endswith((".hwp", ".hwpx")):
                    ui.notify(f"{name}: 한글 파일이 아닙니다", type="warning")
                    return
                # 여러 사용자의 같은 이름 파일이 안 겹치게 고유 접두어를 붙여 저장
                dest = UPLOAD_DIR / f"{uuid.uuid4().hex[:8]}_{name}"
                content = e.content
                data = content.read() if hasattr(content, "read") else content
                dest.write_bytes(data)
                add_file_row(dest, name)

            ui.upload(label="한글 파일 업로드 (.hwp / .hwpx — 여러 개 가능)",
                      multiple=True, auto_upload=True,
                      on_upload=handle_upload) \
                .props('accept=".hwp,.hwpx"').classes("w-full")

            with ui.row().classes("w-full items-center"):
                lbl_count = ui.label("총 0개")
                ui.space()
                ui.button("모두 1부", on_click=lambda: set_all_one()) \
                    .props("flat dense")
                ui.button("전체 비우기", on_click=lambda: clear_files()) \
                    .props("flat dense color=negative")

            files_container = ui.column().classes("w-full gap-1")

            def update_count():
                lbl_count.text = f"총 {len(file_entries)}개"

            def add_file_row(path, display_name):
                with files_container:
                    with ui.row().classes(
                            "w-full items-center border rounded-lg px-3 py-1"
                            " bg-gray-50") as row:
                        ui.label(display_name).classes("flex-1 truncate")
                        split = ui.switch("분할", value=True).props("dense")
                        body = ui.number("본문", min=0, max=999, precision=0) \
                            .classes("w-20").props("outlined dense")
                        answer = ui.number("답지", min=0, max=999,
                                           precision=0) \
                            .classes("w-20").props("outlined dense")
                        entry = {"path": path, "name": display_name,
                                 "split": split, "body": body,
                                 "answer": answer, "row": row}
                        ui.button(icon="delete",
                                  on_click=lambda entry=entry:
                                  remove_entry(entry)) \
                            .props("flat dense round color=negative")
                file_entries.append(entry)
                update_count()

            def remove_entry(entry):
                files_container.remove(entry["row"])
                if entry in file_entries:
                    file_entries.remove(entry)
                try:
                    Path(entry["path"]).unlink(missing_ok=True)
                except Exception:
                    pass
                update_count()

            def clear_files():
                for entry in list(file_entries):
                    remove_entry(entry)

            def set_all_one():
                for entry in file_entries:
                    entry["body"].value = 1
                    entry["answer"].value = 1

        # ---------- 대기열 추가 / 중단 ----------
        def submit_jobs():
            pb, pa = sel_body.value, sel_answer.value
            if not pb or not pa:
                ui.notify("본문/답지 프린터를 모두 선택해주세요.",
                          type="warning")
                return
            if not file_entries:
                ui.notify("인쇄할 파일을 업로드해주세요.", type="warning")
                return

            unset = [e["name"] for e in file_entries
                     if e["body"].value is None or e["answer"].value is None]
            if unset:
                listing = ", ".join(unset[:5])
                if len(unset) > 5:
                    listing += f" 외 {len(unset) - 5}개"
                ui.notify(f"부수가 입력되지 않은 파일: {listing}",
                          type="warning", multi_line=True)
                return

            label = (name_input.value or "").strip() or "이름없음"
            jobs = [(str(e["path"]), e["name"],
                     max(0, int(e["body"].value)),
                     max(0, int(e["answer"].value)),
                     bool(e["split"].value))
                    for e in file_entries]

            # 파일 소유권이 대기열로 넘어가므로 내 목록에서는 행만 제거
            for entry in list(file_entries):
                files_container.remove(entry["row"])
            file_entries.clear()
            update_count()

            job_queue.put({"label": label, "jobs": jobs,
                           "printer_body": pb, "printer_answer": pa})
            log(f"📥 '{label}' 작업 {len(jobs)}건 대기열 추가")
            ui.notify(f"{len(jobs)}건을 인쇄 대기열에 추가했습니다.",
                      type="positive")

        with ui.row().classes("w-full gap-3"):
            ui.button("인쇄 시작 (대기열에 추가)", on_click=submit_jobs) \
                .props("color=positive") \
                .classes("flex-1 h-12 text-base font-bold")
            ui.button("전체 중단", on_click=stop_all) \
                .props("color=negative outline") \
                .classes("w-32 h-12 text-base font-bold")

        # ---------- 공용 진행 상황 ----------
        with ui.card().classes("w-full"):
            with ui.row().classes("w-full items-center"):
                ui.label("서버 상태").classes("font-bold")
                ui.space()
                lbl_queue = ui.label("대기열 0건").classes(
                    "text-sm text-gray-500")
            progress = ui.linear_progress(value=0.0, show_value=False) \
                .classes("w-full")
            status_label = ui.label("대기 중...").classes(
                "text-sm text-blue-600")
            log_view = ui.log(max_lines=500).classes("w-full h-52 text-xs")

        def refresh():
            for line in glog.new_lines_since(seen["n"]):
                log_view.push(line)
            seen["n"] = glog.total
            progress.value = state.progress
            status_label.text = state.status
            waiting = job_queue.qsize()
            running_txt = " · 인쇄 중" if state.running else ""
            lbl_queue.text = f"대기열 {waiting}건{running_txt}"

        ui.timer(0.3, refresh)


if __name__ in {"__main__", "__mp_main__"}:
    ui.run(host="0.0.0.0", port=PORT, title="한글 분할 인쇄",
           reload=False, show=True, language="ko-KR")
