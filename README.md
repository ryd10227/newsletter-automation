# Newsletter Automation

Excel에 정리한 뉴스레터 데이터를 CSV, Photoshop PSD/PNG, VDI HTML, 무투24 HTML 및 반입 ZIP으로 변환하는 Windows용 자동화 도구입니다.

## 구성

- `0_excel_macro`: 뉴스레터 DB 입력·파싱용 Excel 템플릿과 VBA
- `1_excel_to_csv`: Excel 최종 시트를 CSV와 배너 이미지로 추출
- `2_photoshop`: CSV를 PSD에 반영하고 PSD/PNG를 저장하는 Photoshop JSX
- `3_html`: VDI 뉴스레터 HTML과 무투24 HTML 생성
- `4_create_zip`: 최종 반입 ZIP 생성
- `_common`: Windows 실행 메뉴

## 시작 전 필수 준비

`run.cmd`를 실행하기 전에 아래 폴더와 파일을 반드시 준비하세요.

```text
work_history\
└─ [발행일8자리]\
   └─ input\
      ├─ newsletter_DB_[발행일8자리].xlsm 또는 .xlsx
      └─ [발행일8자리]_letter\
         ├─ newsletter_[발행일8자리].psd
         └─ newsletter_[발행일8자리].html
```

날짜 폴더, Excel, letter 폴더, PSD, HTML의 `[발행일8자리]`가 전부 같아야 합니다. Excel 파일은 `.xlsm` 또는 `.xlsx` 중 하나만 넣으세요.

## 빠른 실행

1. 위의 필수 폴더 구조를 준비합니다.
2. `run.cmd`를 실행하고 발행일 8자리를 입력합니다.
3. 메뉴에서 `1`을 입력해 CSV를 생성합니다.
4. 메뉴에서 `2`를 입력해 원본 PSD를 Photoshop으로 열고 `2_photoshop\fill.jsx`를 실행합니다.
5. 메뉴에서 `3`을 입력해 VDI HTML과 무투24 HTML을 생성합니다.
6. 메뉴에서 `4`를 입력해 반입 ZIP을 생성합니다.

자세한 파일 구조와 작업 방법은 `빠른실행.txt`와 `작업순서.txt`를 참고하세요.

## 실행 환경

- Windows
- Python 3
- Microsoft Excel
- Adobe Photoshop

`NewsletterParser_Windows.bas`와 `2_photoshop\fill.jsx`는 Windows 호환을 위해 CP949 및 CRLF 형식으로 저장되어 있습니다.

## 공개 저장소 주의사항

`work_history`의 날짜별 작업 폴더는 Git에서 제외됩니다. 실제 사업명, URL, 이미지, PSD 및 생성 결과물은 공개 저장소에 커밋하지 마세요.
