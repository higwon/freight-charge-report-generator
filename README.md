# Freight Charge Report Generator

포워딩 운임 원본 Excel 파일을 읽어 물류코드별 AR/AP 월별 보고서를 생성하는 Windows 데스크톱 애플리케이션입니다.

## 주요 기능

- `.xls`, `.xlsx`, Excel 2003 XML Spreadsheet 형식을 지원합니다.
- 고정된 시트명이 아닌 필수 컬럼 헤더를 기준으로 Source 시트를 자동 탐지합니다.
- 원본 데이터를 첫 번째 `Source` 시트에 그대로 유지합니다.
- `AR/AP 월별 보고서`, `요약 보고서` 중 하나를 선택해 생성할 수 있습니다.
- `AR/AP 월별 보고서`는 물류코드별 시트를 만들고 월별 `AR계`, `AP계`, `차이`, `마진율`을 표시합니다.
- 수출 코드 `AE`, `OE`, `OA`는 도착항과 거래처 기준으로 집계합니다.
- 수입 코드 `AI`, `OI`는 출발항과 거래처 기준으로 집계합니다.
- 그 외 물류코드는 항구 없이 거래처 기준으로 집계합니다.
- 매출과 차이가 모두 0인 마진율은 `-`, 계산할 수 없는 마진율은 `N/A`로 표시합니다.
- Source Excel 파일을 창에 드래그 앤 드롭할 수 있습니다.
- XML 형식 원본에 잘못 저장된 `&` 문자가 있어도 자동 보정해 처리합니다.
- 진행 단계, 성공/실패 결과, 결과 파일 및 폴더 열기 기능을 제공합니다.

## 설치

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

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

## 결과 파일

### AR/AP 월별 보고서

생성 파일의 첫 번째 시트는 `Source`이며, 원본 데이터 전체가 유지됩니다.

이후 각 물류코드별 시트가 생성됩니다.

```text
수출 코드(AE/OE/OA)
도착항 | 거래처코드 | 거래처명 | 월별 AR계/AP계/차이/마진율

수입 코드(AI/OI)
출발항 | 거래처코드 | 거래처명 | 월별 AR계/AP계/차이/마진율

기타 물류코드
거래처코드 | 거래처명 | 월별 AR계/AP계/차이/마진율
```

계산 기준:

- `AR계`: `AR / AP Type = AR`인 `Loc Amt` 합계
- `AP계`: `AR / AP Type = AP`인 `Loc Amt` 합계
- `차이`: `AR계 - AP계`
- `마진율`: `차이 / AR계`
- `AR계`와 `차이`가 모두 0이면 `-`
- `AR계`가 0이고 `차이`가 0이 아니면 `N/A`

### 요약 보고서

- `Source`: 자동 탐지된 원본 시트의 전체 데이터
- `Overview`: Func Code별, 월별 합계와 상위 Port/Customer 및 차트
- `ALL`: AE, AI, OA, OE, OI 통합 요약
- `AE`, `AI`, `OA`, `OE`, `OI`: Func Code별 요약

시트 순서는 `Source`, `Overview`, `ALL`, 코드별 시트 순서입니다.

## 로컬 빌드

```powershell
python -m PyInstaller FreightChargeReportGenerator.spec --noconfirm --clean
```

빌드 결과는 `dist/FreightChargeReportGenerator` 폴더에 생성됩니다.

## 릴리즈

GitHub Actions는 `main` 브랜치 push, Pull Request, 수동 실행에서 Windows ZIP 파일을 빌드합니다.

`v*` 형식의 태그를 push하면 GitHub Release와 Windows ZIP 파일을 생성합니다.

```powershell
git tag v0.4.0
git push origin v0.4.0
```
