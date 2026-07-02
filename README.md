# Freight Charge Report Generator

포워딩 운임 원본 Excel 파일을 읽어 월, Port, 고객별 운임 요약 보고서를 생성하는 Windows 데스크톱 애플리케이션입니다.

## 주요 기능

- `.xls`, `.xlsx`, Excel 2003 XML Spreadsheet 형식을 지원합니다.
- 고정된 시트명이 아닌 필수 컬럼 헤더를 기준으로 Source 시트를 자동 탐지합니다.
- 원본 데이터를 그대로 담은 `Source` 시트를 생성합니다.
- `AE`, `AI`, `OA`, `OE`, `OI`를 통합한 `ALL` 시트를 생성합니다.
- 각 Func Code별 요약 시트를 생성합니다.
- Month → Port → Customer Name 순서로 금액을 집계합니다.
- Month와 Port 항목을 접고 펼칠 수 있는 Excel 그룹 기능을 적용합니다.
- 샘플 파일 없이 코드에 정의된 보고서 스타일을 적용합니다.
- 필요한 경우에만 소수점을 표시합니다.
- 진행 단계, 성공·실패 결과, 결과 파일 및 폴더 열기 기능을 제공합니다.
- Source Excel 파일을 창에 드래그 앤 드롭할 수 있습니다.
- XML 형식 원본의 잘못 저장된 `&` 문자를 자동 보정합니다.
- 기본 보고서와 분석 보고서를 선택해 생성할 수 있습니다.
- 작업 기록을 `logs/app.log`에 저장합니다.

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

애플리케이션 사용 순서:

1. Source Excel 파일을 선택합니다.
2. 결과를 저장할 폴더를 선택합니다.
3. `result.xlsx`와 같은 출력 파일명을 입력합니다.
4. 기본 보고서 또는 분석 보고서를 선택합니다.
5. `생성` 버튼을 누릅니다.

## 결과 파일

### 기본 보고서

생성된 파일에는 다음 시트가 포함됩니다.

- `Source`: 자동 탐지된 원본 시트의 전체 데이터
- `ALL`: AE, AI, OA, OE, OI 전체 통합 요약
- `AE`, `AI`, `OA`, `OE`, `OI`: Func Code별 요약

각 요약 시트는 아래 구조로 생성됩니다.

```text
Month
    Port
        Customer Name
```

금액은 `Loc Amt` 합계이며 `#,##0.##` 형식으로 표시됩니다.

### 분석 보고서

- `Overview`: Func Code별·월별 합계, 상위 Port/Customer 및 차트
- `ALL`: Month → Func Code → Port → Customer 구조의 통합 상세
- `AE`, `AI`, `OA`, `OE`, `OI`: Func Code별 요약
- `Source`: 자동 필터가 적용된 전체 원본 데이터

## 로컬 빌드

```powershell
python -m PyInstaller FreightChargeReportGenerator.spec --noconfirm --clean
```

빌드 결과는 `dist/FreightChargeReportGenerator` 폴더에 생성됩니다.

## 릴리즈

GitHub Actions는 `main` 브랜치 푸시, Pull Request, 수동 실행에서 Windows ZIP 파일을 빌드합니다.

`v*` 형식의 태그를 푸시하면 GitHub Release와 Windows ZIP 파일이 함께 생성됩니다.

```powershell
git tag v0.2.0
git push origin v0.2.0
```

## 프로젝트 구조

```text
main.py
requirements.txt
README.md
assets/
src/
ui/
logs/
output/
```
