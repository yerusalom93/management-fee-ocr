# Management Fee OCR

Windows 내장 OCR을 사용해 관리비 고지서 PDF에서 주요 금액을 추출하고 Excel 파일로 저장하는 로컬 실행 도구입니다.

## 가장 쉬운 사용법

1. 저장소를 내려받습니다.
2. `run_ocr.bat`을 더블클릭합니다.
3. 뜨는 창에 PDF 파일을 드래그앤드롭합니다.
4. PDF와 같은 폴더에 `<PDF파일명>_ocr.xlsx`가 생성됩니다.

처음 실행할 때는 `.venv` 가상환경을 만들고 필요한 Python 패키지를 자동 설치합니다.

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
- Python 3.11 이상
- PowerShell

## 설치

보통은 `run_ocr.bat` 또는 `run_ocr.ps1`을 처음 실행할 때 자동 설치됩니다. 수동 설치가 필요하면:

```powershell
.\setup.ps1
```

특정 Python 실행 파일을 쓰려면:

```powershell
.\setup.ps1 -PythonExe "C:\Path\To\python.exe"
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
