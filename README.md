# Management Fee OCR

Windows 내장 OCR을 사용해 관리비 고지서 PDF에서 주요 금액을 추출하고 Excel 파일로 저장하는 로컬 실행 도구입니다.

## 특징

- 인터넷 OCR/API를 사용하지 않습니다.
- Windows 내장 OCR을 사용합니다.
- PDF 페이지 렌더링 후 전체 OCR과 표 영역 OCR을 함께 사용합니다.
- 관리비 고지서 표의 행/열 구조를 우선 사용해 숫자 추출 안정성을 높였습니다.
- 원본 PDF, OCR JSON, Excel 결과물은 Git에 포함하지 않도록 제외되어 있습니다.

## 요구 사항

- Windows 10/11
- Windows 한국어 OCR 엔진
- Python 3.11 이상
- PowerShell

처음 실행 시 `setup.ps1`이 `.venv`를 만들고 Python 패키지를 설치합니다.

## 빠른 실행

```powershell
.\run_ocr.ps1 -Pdf ".\sample.pdf" -OutXlsx ".\result.xlsx"
```

또는 더블클릭/명령 프롬프트용:

```bat
run_ocr.bat "sample.pdf" "result.xlsx"
```

처음 실행하면 의존성을 자동 설치합니다. 인터넷이 없는 PC에서는 `.venv` 또는 패키지 wheel을 별도로 준비해야 합니다.

## 수동 설치

```powershell
.\setup.ps1
```

특정 Python을 쓰려면:

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

`expected.json`은 다음처럼 Excel 헤더와 같은 키를 가진 배열입니다.

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
