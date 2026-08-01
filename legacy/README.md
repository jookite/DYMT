# legacy — 이전 데스크톱(tkinter) 버전

웹 버전([../HWP_print_web.py](../HWP_print_web.py))이 나오기 전까지 사용하던
단일 PC용 GUI 버전들입니다. 모두 같은 핵심 로직(첫 미주 기준 본문/답지 분할 인쇄)을
공유하며, 서버 없이 해당 PC에서 바로 실행합니다.

| 파일 | 버전 | 설명 |
|---|---|---|
| `start2zero.py` | v1 | 폴더를 지정하면 안의 모든 hwp/hwpx를 일괄 분할 인쇄 |
| `HWP_print_tool.py` | v2 | 탭 UI — 파일별 부수 지정 + 폴더 일괄 모드 |
| `HWP_split_print.py` | v3 | 파일별 분할/전체 인쇄 모드 선택 추가 |

실행 방법 (모두 동일):

```bash
python HWP_print_tool.py
```

요구 사항: Windows + 한컴오피스 한글 + `pip install pywin32`
