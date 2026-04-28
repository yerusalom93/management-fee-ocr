# Management Fee OCR

Windows 내장 OCR을 사용해 관리비 고지서 PDF에서 주요 금액을 추출하고 Excel 파일로 저장하는 로컬 실행 도구입니다.

## 가장 쉬운 사용법

1. 저장소를 내려받습니다.
2. `run_ocr.bat`을 더블클릭합니다.
3. 뜨는 창에 PDF 파일을 드래그앤드롭합니다.
4. PDF와 같은 폴더에 `<PDF파일명>_ocr.xlsx`가 생성됩니다.

처음 실행할 때는 프로젝트 폴더 안의 `.runtime`에 휴대용 Python을 자동으로 내려받고 필요한 패키지를 설치합니다.
PC에 Python이 없어도 되고, Microsoft Store도 필요하지 않습니다.

## 명령줄 사용

```powershell
.\run_ocr.ps1 -Pdf ".\sample.pdf" -OutXlsx ".\result.xlsx"
```

또는:

```bat
run_ocr.bat "sample.pdf" "result.xlsx"
```

두 번째 인자를 생략하면 PDF와 같은 폴더에 `<PDF파일명>_ocr.xlsx`로 저장합니다.

## 요구 사항

- Windows 10/11
- Windows 한국어 OCR 엔진
- PowerShell
- 처음 설치 시 인터넷 연결

Python은 별도로 설치하지 않아도 됩니다. 단, 완전 오프라인 PC에서는 처음 실행 때 런타임을 받을 수 없으므로, 인터넷이 되는 PC에서 한 번 실행해 `.runtime`이 만들어진 폴더를 통째로 옮기거나 오프라인 배포 ZIP을 만들어야 합니다.

## 오프라인 PC용 ZIP 만들기

인터넷이 되는 PC에서 아래 명령을 한 번 실행하면 Python 런타임과 패키지를 포함한 ZIP이 만들어집니다.

```powershell
.\make_offline_bundle.ps1
```

생성된 `management-fee-ocr-offline.zip`을 오프라인 PC에 옮겨 압축을 풀고 `run_ocr.bat`을 실행하면 됩니다.

직접 만들기 어렵다면 GitHub Releases에서 `management-fee-ocr-offline.zip`을 내려받아 사용할 수도 있습니다.

## 설치

보통은 `run_ocr.bat` 또는 `run_ocr.ps1`을 처음 실행할 때 자동 설치됩니다. 수동 설치가 필요하면:

```powershell
.\setup.ps1
```

기본은 휴대용 Python입니다. 이미 설치된 Python을 강제로 쓰려면:

```powershell
.\setup.ps1 -UseSystemPython -PythonExe "C:\Path\To\python.exe"
```

## 처리 흐름

1. `render_pdf_pages.py`: PDF를 고해상도 PNG로 렌더링
2. `run_windows_ocr.ps1`: Windows OCR로 전체 페이지 인식
3. `crop_ocr_regions.py`: 금액 표 영역 crop
4. `run_windows_ocr.ps1`: 표 영역 추가 OCR
5. `build_management_fee_xlsx.py`: 구조 기반 추출 후 Excel 생성

## 검증

실제 값 검증이 필요한 경우 별도 정답 JSON을 만들어 비교할 수 있습니다.

```powershell
.\.venv\Scripts\python.exe .\test_management_fee_output.py .\result.xlsx .\expected.json
```

`expected.json`은 Excel 헤더와 같은 키를 가진 배열입니다.

```json
[
  {
    "페이지": 1,
    "아파트명": "예시아파트",
    "동": "101",
    "호수": "1001",
    "TV수신료": 2500,
    "수도료": 7000,
    "주차비": null,
    "전기료": 20000,
    "납기내총액": 150000
  }
]
```

## 주의

OCR은 문서 품질과 고지서 양식에 영향을 받습니다. 이 프로젝트는 현재 샘플 양식에 맞춘 구조 기반 추출기입니다. 새 양식이 들어오면 좌표/행 구조 규칙을 추가해야 합니다.

원본 PDF, OCR JSON, Excel 결과물, 이미지 캐시는 Git에 포함하지 않도록 `.gitignore`에서 제외합니다.

## 문제 해결

엑셀 파일이 생성되지 않으면 `work\last_gui.log`를 확인하세요. GUI에서 실패하면 같은 경로를 메시지로 보여줍니다.

자주 발생하는 원인:

- Python 설치 실패
- 휴대용 Python 다운로드 실패
- 첫 실행 시 인터넷 연결 없음
- Windows 한국어 OCR 엔진 없음
- PDF가 지원 양식과 다름
