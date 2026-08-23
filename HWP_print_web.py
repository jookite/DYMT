"""
한글 문서 본문/답지 분할 인쇄 도구 — 다중 사용자 웹 최종판 (NiceGUI)

이 스크립트를 실행하는 Windows PC(한글 + 프린터 연결)가 '인쇄 서버'가 되고,
같은 네트워크의 여러 사람이 각자 브라우저로 접속해 동시에 사용할 수 있습니다.

  실행:        python HWP_print_web.py
  이 PC에서:   http://localhost:58021
  다른 기기:   http://<이 PC의 IP>:58021  (화면 상단에 주소 표시)

Lab 연동
  - Lab(https://lab.dayoungmathclass.com)의 '프린트 도우미' 메뉴가
    이 프로그램(127.0.0.1:58021)의 /api/* 창구로 직접 접속해
    파일 업로드·오답노트 ZIP 인쇄·진행 확인을 할 수 있다.
  - 허용된 출처(ALLOWED_ORIGINS)의 요청만 받는다.

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

import json
import os
import queue
import shutil
import tempfile
import threading
import time
import uuid
import zipfile
from collections import deque
from pathlib import Path

import pythoncom
import win32com.client
import win32print

from fastapi import File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from nicegui import app, ui

UPLOAD_DIR = Path(tempfile.gettempdir()) / "hwp_print_web_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

PORT = 58021          # Lab 연동용 고정 포트 (8080은 충돌이 잦아 변경)
AGENT_VERSION = "1.1.0"

# Lab(웹 관리자 페이지)이 브라우저에서 이 프로그램(127.0.0.1)으로 직접
# 요청을 보낼 수 있게 허용할 출처(Origin) 목록.
# 여기 없는 사이트가 보낸 API 요청은 거절한다 (악성 사이트의 몰래 인쇄 방지).
ALLOWED_ORIGINS = [
    "https://lab.dayoungmathclass.com",
    # Lab 로컬 개발 서버
    "http://localhost:3000", "http://127.0.0.1:3000",
    "http://localhost:3001", "http://127.0.0.1:3001",
    # 이 프로그램 자신의 화면
    f"http://localhost:{PORT}", f"http://127.0.0.1:{PORT}",
]

HWP_EXTENSIONS = (".hwp", ".hwpx")


# ============================================================
# 공통 HWP 처리 로직
# ============================================================
class FileResult:
    """파일 하나의 처리 결과. bool()은 성공 여부, 속성은 상세 정보."""

    def __init__(self, ok, split_page=None, total_pages=None,
                 skipped_no_endnote=False):
        self.ok = ok
        self.split_page = split_page
        self.total_pages = total_pages
        self.skipped_no_endnote = skipped_no_endnote

    def __bool__(self):
        return self.ok


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
        표시 쪽 번호를 쓰지 않아 '새 쪽번호' 문서에서도 안전하다.
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
        """예비 방식: 답지 영역이 문서 끝에 이어붙는 점을 이용해
        (전체 쪽수 - 답지 영역 쪽수 + 1)로 계산. 영역 쪽수는 표시 번호의
        차이로 구하므로 '새 쪽번호' 오프셋이 상쇄된다."""
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

    def find_answer_end_page(self, upto_num, total_pages):
        """upto_num번째 미주(문제)의 답이 끝나는 실제 쪽 번호.

        마지막 미주의 끝이 문서 마지막 쪽이라는 점을 기준으로,
        표시 번호의 차이만큼 문서 끝에서 되짚어 계산한다.
        문제 수 이상을 지정했거나 계산이 안 되면 None (→ 답지 전체 인쇄).
        """
        ens = []
        ctrl = self.hwp.HeadCtrl
        while ctrl:
            if ctrl.CtrlID == "en":
                ens.append(ctrl)
            ctrl = ctrl.Next
        if not ens or upto_num < 1 or upto_num >= len(ens):
            return None
        try:
            end_shown = self._note_shown_page(ens[-1], at_end=True)
            k_shown = self._note_shown_page(ens[upto_num - 1], at_end=True)
            page = total_pages - (end_shown - k_shown)
            if not 1 <= page <= total_pages:
                return None
            return page
        except Exception:
            return None

    # ---------- 인쇄 ----------
    def _print_range(self, print_act, print_set, printer_name, page_range,
                     copies, dry_run=False):
        """지정 범위를 지정 부수(NumCopy)만큼 한 번에 인쇄."""
        if dry_run:
            # 검사 모드: 분할 계산까지만 하고 실제 인쇄는 하지 않는다
            self.log(f"   🧪 (검사) {page_range}쪽 · {copies}부 "
                     f"→ {printer_name} — 실제 인쇄 안 함")
            return
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
                      printer_answer, body_copies, answer_copies,
                      answer_upto=None, dry_run=False):
        """본문과 답지를 각각 해당 프린터로 보낸다.

        answer_upto: 답지를 몇 번 문제의 답까지만 인쇄할지 (None이면 전체).
        """
        # 답지 마지막 쪽 결정 — 기본은 문서 끝, ~번까지 지정 시 해당 답 끝 쪽
        answer_last = total_pages
        upto_label = ""
        if answer_upto:
            end_pg = self.find_answer_end_page(answer_upto, total_pages)
            if end_pg is not None and end_pg >= split_page:
                answer_last = end_pg
                upto_label = f" ({answer_upto}번까지)"
            else:
                self.log(f"   ⚠️ '{answer_upto}번까지' 지정이 문제 수 이상이거나 "
                         f"계산할 수 없어 답지 전체를 인쇄합니다.")
        try:
            print_act = self.hwp.CreateAction("Print")
            print_set = print_act.CreateSet()

            # 1) 본문: 1쪽 ~ (split_page - 1)쪽
            if split_page > 1 and body_copies > 0:
                body_range = f"1-{split_page - 1}"
                self._print_range(print_act, print_set, printer_body,
                                  body_range, body_copies, dry_run=dry_run)
                self.log(f"   🖨️ 본문 {body_range}쪽 · {body_copies}부 "
                         f"→ {printer_body}")
            elif body_copies == 0:
                self.log("   ⏭️ 본문 0부 — 건너뜀")

            # 2) 답지: split_page쪽 ~ 답지 마지막 쪽
            if split_page <= answer_last and answer_copies > 0:
                answer_range = f"{split_page}-{answer_last}"
                self._print_range(print_act, print_set, printer_answer,
                                  answer_range, answer_copies,
                                  dry_run=dry_run)
                self.log(f"   🖨️ 답지 {answer_range}쪽{upto_label} · "
                         f"{answer_copies}부 → {printer_answer}")
            elif answer_copies == 0:
                self.log("   ⏭️ 답지 0부 — 건너뜀")

            return True
        except Exception as e:
            self.log(f"   ❌ 인쇄 중 오류: {e}")
            return False

    def execute_whole_print(self, total_pages, printer_body, printer_answer,
                            body_copies, answer_copies, dry_run=False):
        """미주 분할 없이 전체 페이지를 양쪽 프린터로 지정 부수만큼 인쇄."""
        try:
            print_act = self.hwp.CreateAction("Print")
            print_set = print_act.CreateSet()
            whole_range = f"1-{total_pages}"

            if body_copies > 0:
                self._print_range(print_act, print_set, printer_body,
                                  whole_range, body_copies, dry_run=dry_run)
                self.log(f"   🖨️ 전체 {whole_range}쪽 · {body_copies}부 "
                         f"→ {printer_body}")
            else:
                self.log("   ⏭️ 본문 0부 — 본문 프린터 건너뜀")

            if answer_copies > 0:
                self._print_range(print_act, print_set, printer_answer,
                                  whole_range, answer_copies, dry_run=dry_run)
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
                     printer_body, printer_answer, on_no_endnote=None,
                     answer_upto=None, dry_run=False):
        """파일 하나를 열어 분할/전체 모드에 따라 인쇄한다. FileResult 반환."""
        try:
            if not self.hwp.Open(file_path, "", ""):
                self.log("   ❌ 파일을 열지 못했습니다.")
                return FileResult(False)

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
                    return FileResult(False, total_pages=total_pages,
                                      skipped_no_endnote=True)

                ok = self.execute_print(split_page, total_pages,
                                        printer_body, printer_answer,
                                        body_copies, answer_copies,
                                        answer_upto=answer_upto,
                                        dry_run=dry_run)
                result = FileResult(ok, split_page=split_page,
                                    total_pages=total_pages)
            else:
                self.log(f"   📑 전체 인쇄 모드 ({total_pages}쪽)")
                ok = self.execute_whole_print(total_pages, printer_body,
                                              printer_answer, body_copies,
                                              answer_copies, dry_run=dry_run)
                result = FileResult(ok, total_pages=total_pages)

            self.hwp.Run("FileClose")
            return result
        except Exception as e:
            self.log(f"   ❌ 처리 중 오류: {e}")
            try:
                self.hwp.Run("FileClose")
            except Exception:
                pass
            return FileResult(False)


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


class LabJobRegistry:
    """Lab에서 접수한 인쇄 작업의 상태·로그 저장소 (스레드 안전).

    NiceGUI 화면에서 넣은 작업은 여기 등록되지 않는다 (기존 동작 유지).
    """

    def __init__(self):
        self._jobs = {}
        self._lock = threading.Lock()

    def create(self, label, files, printer_body, printer_answer, dry_run):
        job_id = uuid.uuid4().hex[:12]
        with self._lock:
            self._jobs[job_id] = {
                "id": job_id,
                "label": label,
                # queued | printing | done | failed | cancelled
                "status": "queued",
                "dry_run": dry_run,
                "printer_body": printer_body,
                "printer_answer": printer_answer,
                "created_at": time.time(),
                "finished_at": None,
                "cancel_requested": False,
                "files": files,
                "logs": [],
            }
        return job_id

    def get(self, job_id):
        with self._lock:
            return self._jobs.get(job_id)

    def log(self, job_id, line):
        with self._lock:
            job = self._jobs.get(job_id)
            if job is not None:
                job["logs"].append(f"[{time.strftime('%H:%M:%S')}] {line}")

    def update(self, job_id, **fields):
        with self._lock:
            job = self._jobs.get(job_id)
            if job is not None:
                job.update(fields)

    def update_file(self, job_id, index, **fields):
        with self._lock:
            job = self._jobs.get(job_id)
            if job is not None and 0 <= index < len(job["files"]):
                job["files"][index].update(fields)

    def snapshot(self, job_id, since_log=0):
        """API 응답용 복사본. logs는 since_log 이후 것만 담는다."""
        with self._lock:
            job = self._jobs.get(job_id)
            if job is None:
                return None
            data = {k: v for k, v in job.items() if k != "logs"}
            data["files"] = [dict(f) for f in job["files"]]
            data["logs"] = list(job["logs"][since_log:])
            data["log_total"] = len(job["logs"])
            return data


glog = GlobalLog()
state = AppState()
lab_jobs = LabJobRegistry()
job_queue = queue.Queue()  # 항목: {'label', 'jobs', 'printer_body', 'printer_answer', ...}

log = glog.add


def get_printer_list():
    try:
        return [p[2] for p in win32print.EnumPrinters(2)]
    except Exception:
        return []


# ============================================================
# 인쇄 워커 — 서버에 하나만 떠서 대기열을 순서대로 처리
# ============================================================
def _cleanup_batch(batch):
    """배치가 쓰던 업로드 임시 파일·폴더를 정리한다."""
    for path, *_ in batch["jobs"]:
        try:
            Path(path).unlink(missing_ok=True)
        except Exception:
            pass
    for d in batch.get("cleanup_dirs", []):
        shutil.rmtree(d, ignore_errors=True)


def worker_loop():
    pythoncom.CoInitialize()
    while True:
        batch = job_queue.get()  # 작업이 들어올 때까지 대기
        job_id = batch.get("job_id")  # Lab에서 접수한 작업이면 ID가 있다
        job = lab_jobs.get(job_id) if job_id else None

        # 대기 중에 취소된 Lab 작업은 실행하지 않고 정리만 한다
        if job is not None and job["cancel_requested"]:
            lab_jobs.update(job_id, status="cancelled",
                            finished_at=time.time())
            lab_jobs.log(job_id, "🚫 시작 전에 취소되었습니다.")
            log(f"🚫 '{batch['label']}' 작업이 시작 전에 취소되었습니다.")
            _cleanup_batch(batch)
            job_queue.task_done()
            continue

        state.running = True
        state.stop_flag = False
        state.progress = 0.0

        jobs = batch["jobs"]
        label = batch["label"]
        dry_run = batch.get("dry_run", False)

        def blog(message):
            """공용 로그 + (Lab 작업이면) 작업별 로그에 동시 기록."""
            log(message)
            if job_id:
                lab_jobs.log(job_id, message)

        printer = HwpPrinter(blog)
        skipped_files = []
        cancelled = False
        total = len(jobs)
        success = 0

        if job is not None:
            lab_jobs.update(job_id, status="printing")

        try:
            blog(f"▶️ '{label}' 작업 시작 — 파일 {total}건 "
                 f"(대기 중 {job_queue.qsize()}건)"
                 + (" [검사 모드 — 실제 인쇄 안 함]" if dry_run else ""))
            printer.start()

            for i, (path, name, body_copies, answer_copies, do_split,
                    answer_upto) in enumerate(jobs):
                current = lab_jobs.get(job_id) if job_id else None
                if state.stop_flag or (current is not None
                                       and current["cancel_requested"]):
                    blog("⚠️ 중단 요청으로 현재 작업을 멈췄습니다.")
                    cancelled = True
                    break

                mode_label = "분할" if do_split else "전체"
                upto_label = f" / 답지 {answer_upto}번까지" if answer_upto else ""
                state.progress = i / total
                state.status = f"[{label}] 처리 중: {name}"
                blog(f"📄 [{i + 1}/{total}] {name} [{mode_label}] "
                     f"(본문 {body_copies}부 / 답지 {answer_copies}부"
                     f"{upto_label})")
                if job_id:
                    lab_jobs.update_file(job_id, i, status="printing")

                result = printer.process_file(
                    path, body_copies, answer_copies, do_split,
                    batch["printer_body"], batch["printer_answer"],
                    on_no_endnote=skipped_files.append,
                    answer_upto=answer_upto, dry_run=dry_run)
                if result:
                    success += 1
                if job_id:
                    if result:
                        fstatus = "done"
                    elif result.skipped_no_endnote:
                        fstatus = "skipped_no_endnote"
                    else:
                        fstatus = "failed"
                    lab_jobs.update_file(job_id, i, status=fstatus,
                                         split_page=result.split_page,
                                         total_pages=result.total_pages)

            state.progress = 1.0
            blog(f"✅ '{label}' 완료 — 성공 {success} / 전체 {total}건")
            if skipped_files:
                blog(f"⏭️ 미주가 없어 건너뛴 파일 {len(skipped_files)}개: "
                     + ", ".join(skipped_files))
            if job_id:
                if cancelled:
                    final = "cancelled"
                elif success == 0 and total > 0:
                    final = "failed"
                else:
                    final = "done"
                lab_jobs.update(job_id, status=final,
                                finished_at=time.time())
                # 처리되지 못한 파일 상태 정리
                snapshot = lab_jobs.get(job_id)
                if snapshot is not None:
                    for idx, f in enumerate(snapshot["files"]):
                        if f["status"] in ("pending", "printing"):
                            lab_jobs.update_file(
                                job_id, idx,
                                status="cancelled" if cancelled else "failed")
        except Exception as e:
            blog(f"❌ '{label}' 처리 중 오류: {e}")
            if job_id:
                lab_jobs.update(job_id, status="failed",
                                finished_at=time.time())
        finally:
            printer.quit()
            _cleanup_batch(batch)
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

# '폴더째로 업로드' 위젯의 파일 선택창을 폴더 선택 모드로 바꾸는 스크립트.
# Quasar가 입력창을 다시 그려도 유지되도록 MutationObserver로 재적용한다.
_FOLDER_PICKER_JS = """
const applyFolderMode = () =>
    document.querySelectorAll('.folder-uploader input[type="file"]')
        .forEach(i => { i.setAttribute('webkitdirectory', '');
                        i.setAttribute('directory', ''); });
applyFolderMode();
new MutationObserver(applyFolderMode)
    .observe(document.body, {subtree: true, childList: true});
"""


@ui.page("/")
async def index():
    file_entries = []      # 이 접속자만의 파일 목록
    seen = {"n": glog.total - len(glog.lines)}  # 로그는 최근 것부터 표시

    with ui.column().classes("w-full max-w-3xl mx-auto p-4 gap-3"):
        ui.label("강다영수학 프린트 도우미").classes("text-xl font-bold")

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

            async def handle_upload(e):
                # NiceGUI 3.x: 업로드 파일은 e.file (FileUpload 객체)로 전달됨
                file = e.file
                name = file.name
                if not name.lower().endswith((".hwp", ".hwpx")):
                    ui.notify(f"{name}: 한글 파일이 아닙니다", type="warning")
                    return
                # 여러 사용자의 같은 이름 파일이 안 겹치게 고유 접두어를 붙여 저장
                dest = UPLOAD_DIR / f"{uuid.uuid4().hex[:8]}_{name}"
                await file.save(dest)
                add_file_row(dest, name)

            with ui.row().classes("w-full gap-2 items-start"):
                upload_files = ui.upload(
                    label="파일 업로드 (.hwp / .hwpx — 여러 개·드래그 가능)",
                    multiple=True, auto_upload=True,
                    on_upload=handle_upload) \
                    .props('accept=".hwp,.hwpx"').classes("flex-1")
                # 아래 위젯은 접속 후 JS로 폴더 선택 모드(webkitdirectory)로 전환됨
                upload_folder = ui.upload(
                    label="📁 오답노트 폴더 프린트",
                    multiple=True, auto_upload=True,
                    on_upload=handle_upload) \
                    .props('accept=".hwp,.hwpx"') \
                    .classes("flex-1 folder-uploader")

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
                        body = ui.number("본문", value=1, min=0, max=999,
                                         precision=0) \
                            .classes("w-20").props("outlined dense")
                        answer = ui.number("답지", value=1, min=0, max=999,
                                           precision=0) \
                            .classes("w-20").props("outlined dense")
                        upto = ui.number("~번까지", min=1, max=999,
                                         precision=0) \
                            .classes("w-24").props("outlined dense") \
                            .tooltip("답지를 몇 번 문제의 답까지만 인쇄할지 "
                                     "(비우면 답지 전체)")
                        entry = {"path": path, "name": display_name,
                                 "split": split, "body": body,
                                 "answer": answer, "upto": upto, "row": row}
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
                upload_files.reset()   # 업로드 위젯의 전송 내역 표시도 초기화
                upload_folder.reset()

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
                     bool(e["split"].value),
                     int(e["upto"].value) if e["upto"].value else None)
                    for e in file_entries]

            # 파일 소유권이 대기열로 넘어가므로 내 목록에서는 행만 제거
            for entry in list(file_entries):
                files_container.remove(entry["row"])
            file_entries.clear()
            update_count()
            upload_files.reset()
            upload_folder.reset()

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

    await ui.context.client.connected()
    ui.run_javascript(_FOLDER_PICKER_JS)


# ============================================================
# Lab 연동 API — Lab(https://lab.dayoungmathclass.com) 페이지가
# 브라우저에서 http://127.0.0.1:58021 로 직접 호출하는 창구.
# 브라우저는 localhost 통신을 혼합 콘텐츠 차단 예외로 허용한다.
# ============================================================
# allow_private_network: 공개 사이트(Lab)가 사설 주소(127.0.0.1)로 요청하는
# Chrome PNA 사전 요청을 허용 (Starlette가 허용 헤더를 자동으로 붙인다)
app.add_middleware(CORSMiddleware, allow_origins=ALLOWED_ORIGINS,
                   allow_methods=["*"], allow_headers=["*"],
                   allow_private_network=True)


def _check_origin(request: Request):
    """허용 목록 밖의 웹사이트가 보낸 요청을 거절한다.

    Origin 헤더가 아예 없는 요청(주소창 직접 입력, curl 등 브라우저 밖
    도구)은 이 PC 사용자의 직접 행동이므로 허용한다.
    """
    origin = request.headers.get("origin")
    if origin is not None and origin not in ALLOWED_ORIGINS:
        raise HTTPException(status_code=403, detail="허용되지 않은 출처입니다.")


def _parse_copies(value, default=1):
    try:
        return max(0, int(value))
    except (TypeError, ValueError):
        return default


def _parse_upto(value):
    try:
        n = int(value)
        return n if n >= 1 else None
    except (TypeError, ValueError):
        return None


def _validate_printers(printer_body, printer_answer):
    available = get_printer_list()
    for name in (printer_body, printer_answer):
        if name not in available:
            raise HTTPException(
                status_code=400,
                detail=f"프린터 '{name}'를 이 PC에서 찾을 수 없습니다.")


def _enqueue_lab_job(label, entries, printer_body, printer_answer,
                     dry_run, cleanup_dirs=None):
    """entries: [(path, name, body, answer, do_split, upto)] → 작업 등록."""
    files_meta = [{
        "name": name,
        "do_split": do_split,
        "body_copies": body,
        "answer_copies": answer,
        "answer_upto": upto,
        "status": "pending",
        "split_page": None,
        "total_pages": None,
    } for (_path, name, body, answer, do_split, upto) in entries]

    job_id = lab_jobs.create(label, files_meta, printer_body,
                             printer_answer, dry_run)
    job_queue.put({
        "label": label,
        "jobs": [(str(p), n, b, a, s, u) for (p, n, b, a, s, u) in entries],
        "printer_body": printer_body,
        "printer_answer": printer_answer,
        "job_id": job_id,
        "dry_run": dry_run,
        "cleanup_dirs": cleanup_dirs or [],
    })
    log(f"📥 [Lab] '{label}' 작업 {len(entries)}건 대기열 추가"
        + (" (검사 모드)" if dry_run else ""))
    return job_id


@app.get("/api/health")
def api_health(request: Request):
    _check_origin(request)
    return {
        "ok": True,
        "app": "강다영수학 프린트 도우미",
        "version": AGENT_VERSION,
        "running": state.running,
        "queue_size": job_queue.qsize(),
        "status": state.status,
    }


@app.get("/api/printers")
def api_printers(request: Request):
    _check_origin(request)
    try:
        default = win32print.GetDefaultPrinter()
    except Exception:
        default = None
    return {"printers": get_printer_list(), "default": default}


@app.post("/api/jobs")
async def api_create_job(request: Request,
                         files: list[UploadFile] = File(...),
                         items: str = Form("[]"),
                         printer_body: str = Form(...),
                         printer_answer: str = Form(...),
                         label: str = Form("Lab"),
                         dry_run: str = Form("0")):
    """파일 여러 개 + 파일별 옵션(items JSON 배열, files와 같은 순서)."""
    _check_origin(request)
    _validate_printers(printer_body, printer_answer)
    try:
        opts = json.loads(items)
        if not isinstance(opts, list):
            raise ValueError
    except ValueError:
        raise HTTPException(status_code=400,
                            detail="items는 JSON 배열이어야 합니다.")

    entries = []
    try:
        for idx, upload in enumerate(files):
            o = opts[idx] if idx < len(opts) and isinstance(opts[idx], dict) \
                else {}
            # 파일명은 items의 name을 우선 사용 (JSON은 인코딩이 항상
            # UTF-8이라, multipart 파일명 인코딩 차이로 깨질 일이 없다)
            name = os.path.basename(
                str(o.get("name") or upload.filename or f"파일{idx + 1}"))
            if not name.lower().endswith(HWP_EXTENSIONS):
                raise HTTPException(status_code=400,
                                    detail=f"'{name}'은 한글 파일(.hwp/.hwpx)이 "
                                           f"아닙니다.")
            dest = UPLOAD_DIR / f"{uuid.uuid4().hex[:8]}_{name}"
            dest.write_bytes(await upload.read())
            entries.append((dest, name,
                            _parse_copies(o.get("body_copies")),
                            _parse_copies(o.get("answer_copies")),
                            bool(o.get("do_split", True)),
                            _parse_upto(o.get("answer_upto"))))
    except HTTPException:
        for path, *_ in entries:
            Path(path).unlink(missing_ok=True)
        raise
    if not entries:
        raise HTTPException(status_code=400, detail="파일이 없습니다.")

    job_id = _enqueue_lab_job(label.strip() or "Lab", entries,
                              printer_body, printer_answer, dry_run == "1")
    return {"job_id": job_id}


@app.post("/api/jobs/zip")
async def api_create_job_zip(request: Request,
                             file: UploadFile = File(...),
                             printer_body: str = Form(...),
                             printer_answer: str = Form(...),
                             label: str = Form("Lab"),
                             do_split: str = Form("1"),
                             body_copies: str = Form("1"),
                             answer_copies: str = Form("1"),
                             answer_upto: str = Form(""),
                             dry_run: str = Form("0")):
    """ZIP 하나(예: 오답노트 묶음)를 풀어 안의 한글 파일 전체를 공통
    옵션으로 인쇄한다."""
    _check_origin(request)
    _validate_printers(printer_body, printer_answer)

    zip_path = UPLOAD_DIR / f"{uuid.uuid4().hex[:8]}_upload.zip"
    zip_path.write_bytes(await file.read())
    extract_dir = UPLOAD_DIR / f"zip_{uuid.uuid4().hex[:8]}"
    extract_dir.mkdir()

    entries = []
    try:
        with zipfile.ZipFile(zip_path) as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                raw_name = info.filename
                # UTF-8 플래그가 없는 ZIP의 한글 파일명 복원 (cp437→cp949)
                if not info.flag_bits & 0x800:
                    try:
                        raw_name = raw_name.encode("cp437").decode("cp949")
                    except (UnicodeEncodeError, UnicodeDecodeError):
                        pass
                base = os.path.basename(raw_name)
                if not base.lower().endswith(HWP_EXTENSIONS):
                    continue
                dest = extract_dir / f"{len(entries):03d}_{base}"
                with zf.open(info) as src, open(dest, "wb") as out:
                    shutil.copyfileobj(src, out)
                entries.append((dest, base,
                                _parse_copies(body_copies),
                                _parse_copies(answer_copies),
                                do_split == "1",
                                _parse_upto(answer_upto)))
    except zipfile.BadZipFile:
        shutil.rmtree(extract_dir, ignore_errors=True)
        zip_path.unlink(missing_ok=True)
        raise HTTPException(status_code=400, detail="ZIP 파일이 아닙니다.")
    finally:
        zip_path.unlink(missing_ok=True)

    if not entries:
        shutil.rmtree(extract_dir, ignore_errors=True)
        raise HTTPException(status_code=400,
                            detail="ZIP 안에 한글 파일(.hwp/.hwpx)이 없습니다.")

    entries.sort(key=lambda e: e[1])  # 파일명(학생 이름)순 인쇄
    job_id = _enqueue_lab_job(label.strip() or "Lab", entries,
                              printer_body, printer_answer, dry_run == "1",
                              cleanup_dirs=[str(extract_dir)])
    return {"job_id": job_id, "file_count": len(entries)}


@app.get("/api/jobs/{job_id}")
def api_get_job(request: Request, job_id: str, since_log: int = 0):
    _check_origin(request)
    snapshot = lab_jobs.snapshot(job_id, since_log=max(0, since_log))
    if snapshot is None:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다.")
    snapshot["queue_size"] = job_queue.qsize()
    return snapshot


@app.post("/api/jobs/{job_id}/cancel")
def api_cancel_job(request: Request, job_id: str):
    _check_origin(request)
    job = lab_jobs.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다.")
    if job["status"] in ("done", "failed", "cancelled"):
        return {"ok": True, "status": job["status"]}
    lab_jobs.update(job_id, cancel_requested=True)
    lab_jobs.log(job_id, "🛑 취소 요청됨")
    return {"ok": True, "status": job["status"]}


if __name__ in {"__main__", "__mp_main__"}:
    ui.run(host="0.0.0.0", port=PORT, title="강다영수학 프린트 도우미",
           reload=False, show=True, language="ko-KR")
