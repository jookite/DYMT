"""분할 지점 검사 도구 — 인쇄 없이 samples/ 안의 모든 HWP의 분할 지점만 계산해 출력.

사용법:  python check_split.py [폴더경로]   (기본: samples, 하위 폴더 포함)

실제 서비스와 동일한 HwpPrinter.find_split_page_by_first_endnote()를 사용하므로
여기서 나온 값이 곧 인쇄 시 분할 지점이다.
"""

import os
import sys

import pythoncom

from HWP_print_web import HwpPrinter


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else "samples"
    files = sorted(
        os.path.join(root, f)
        for root, _, names in os.walk(folder)
        for f in names
        if f.lower().endswith((".hwp", ".hwpx")))
    if not files:
        print(f"'{folder}' 폴더에 hwp/hwpx 파일이 없습니다.")
        return

    pythoncom.CoInitialize()
    printer = HwpPrinter(lambda msg: print("   " + msg.strip()))
    try:
        printer.start()
        for path in files:
            rel = os.path.relpath(path, folder)
            path = os.path.abspath(path)
            print(f"\n📄 {rel}")
            if not printer.hwp.Open(path, "", ""):
                print("   ❌ 열기 실패")
                continue
            total = printer.hwp.PageCount
            split = printer.find_split_page_by_first_endnote()
            if split is None:
                print(f"   결과: 미주 없음 (전체 {total}쪽)")
            else:
                print(f"   결과: 본문 1-{split - 1}쪽 / "
                      f"답지 {split}-{total}쪽")
            printer.hwp.Run("FileClose")
    finally:
        printer.quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
