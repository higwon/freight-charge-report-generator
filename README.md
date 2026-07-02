# Freight Charge Report Generator

포워딩 운임 원본 Excel 파일을 읽어 물류코드별 AR/AP 월별 보고서와 요약 보고서를 생성하는 Windows 데스크톱 애플리케이션입니다.

## 주요 기능

- `.xls`, `.xlsx`, Excel 2003 XML Spreadsheet 형식을 지원합니다.
- 고정된 시트명이 아닌 필수 컬럼 헤더를 기준으로 Source 시트를 자동 탐지합니다.
- 원본 데이터를 첫 번째 `Source` 시트에 그대로 유지합니다.
- `AR/AP 월별 보고서`, `요약 보고서` 중 하나를 선택해 생성할 수 있습니다.
- `AR/AP 월별 보고서`는 `Overview`와 물류코드별 상세 시트를 생성합니다.
- 월별보고서 Overview에는 전체 KPI, 코드별 요약, 월별 요약, 상위 Port/Customer, 검토 대상, Top/Bottom 거래처, 차트가 포함됩니다.
- 각 코드별 시트에는 코드 전체 월별 요약과 거래처별 누적 요약이 포함됩니다.
- 수출 코드 `AE`, `OE`, `OA`는 도착항과 거래처 기준으로 집계합니다.
- 수입 코드 `AI`, `OI`는 출발항과 거래처 기준으로 집계합니다.
- 그 외 물류코드는 항구 없이 거래처 기준으로 집계합니다.
- 마진율은 정상 계산, `-`, `매출만발생`, `매입만발생`으로 구분해 표시합니다.
- Source Excel 파일을 창에 드래그 앤 드롭할 수 있습니다.
- XML 형식 원본에 잘못 저장된 `&` 문자가 있어도 자동 보정해 처리합니다.
- 진행 단계, 성공/실패 결과, 결과 파일 및 폴더 열기 기능을 제공합니다.

## 실행

```powershell
python main.py
```

사용 순서:

1. Source Excel 파일을 선택하거나 창에 드래그 앤 드롭합니다.
2. 결과를 저장할 폴더를 선택합니다.
3. 출력 파일명을 입력합니다.
4. 보고서 형식을 선택합니다. 기본 선택은 `AR/AP 월별 보고서`입니다.
5. `생성` 버튼을 누릅니다.

## AR/AP 월별 보고서

시트 순서:

```text
Source
Overview
물류코드별 시트
```

Overview에는 기존 요약 보고서 Overview의 주요 항목을 포함하고, AR/AP 검토용 항목을 추가로 제공합니다.

- 전체 매출계, 매입계, 차이, 마진율
- 코드별 매출계, 매입계, 차이, 마진율
- 월별 매출계, 매입계, 차이, 마진율
- 상위 Port, 상위 Customer
- 매출만발생, 매입만발생, 마진 음수 검토 대상
- 차이 상위/하위 거래처
- 코드별 차이, 월별 차이 차트

각 코드별 시트 구조:

```text
상단 요약
구분 | 총 누적 | 2026/01 | 2026/02 | ...
매출계
매입계
차이
마진율

상세
항구 | 거래처코드 | 거래처명 | 누적 매출계 | 누적 매입계 | 누적 차이 | 누적 마진율 | 월별 AR계/AP계/차이/마진율...
```

마진율 표시 기준:

- `AR=0`, `AP=0`: `-`
- `AR>0`, `AP=0`: `매출만발생`
- `AR=0`, `AP>0`: `매입만발생`
- 그 외: `차이 / AR`

숫자는 `#,##0.##` 형식으로 표시해 불필요한 소수점은 보이지 않게 합니다.

## 요약 보고서

- `Source`: 자동 탐지된 원본 시트의 전체 데이터
- `Overview`: Func Code별, 월별 합계와 상위 Port/Customer 및 차트
- `ALL`: AE, AI, OA, OE, OI 통합 요약
- `AE`, `AI`, `OA`, `OE`, `OI`: Func Code별 요약

시트 순서는 `Source`, `Overview`, `ALL`, 코드별 시트 순서입니다.

## 설치

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 로컬 빌드

```powershell
python -m PyInstaller FreightChargeReportGenerator.spec --noconfirm --clean
```

빌드 결과는 `dist/FreightChargeReportGenerator` 폴더에 생성됩니다.

## 릴리즈

GitHub Actions는 `main` 브랜치 push, Pull Request, 수동 실행에서 Windows ZIP 파일을 빌드합니다.

`v*` 형식의 태그를 push하면 GitHub Release와 Windows ZIP 파일을 생성합니다.

```powershell
git tag v0.5.0
git push origin v0.5.0
```
