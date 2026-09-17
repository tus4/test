# 설정 가이드 (SETUP) — 비공개 Excel → 공개 대시보드

비공개 저장소 **`tus4/data1`** 에 Excel 파일(`.xlsx`, `.xlsm`)을 올리면, 그 데이터로 만든 대시보드가 공개 저장소 **`tus4/test`** 의 첫 화면(README)과 GitHub Pages(<https://tus4.github.io/test/>)에 자동으로 게시되도록 설정하는 방법입니다. GitHub 기능(Actions, Pages, 저장소 시크릿, Fine-grained 개인용 액세스 토큰)만 사용하며 외부 서비스는 없습니다.

## 전체 구조

```
tus4/data1 (비공개)                                   tus4/test (공개)
├─ *.xlsx / *.xlsm (아무 폴더)                           ├─ dashboard_builder/   빌드 도구 (Python 3.11 + openpyxl)
├─ dashboard.config.json (선택: 공개 범위·표시 설정)      ├─ template/index.html  표 대시보드 템플릿
│                                                        ├─ template/market.html 시계열 대시보드 템플릿
└─ .github/workflows/publish-dashboard.yml               ├─ tests/, .github/workflows/ci.yml
   + 시크릿 DASHBOARD_PUSH_TOKEN                         ├─ data1/               data1 에 복사할 두 파일의 사본
        │                                                ├─ README.md            ← 생성된 요약 (표·차트·미리보기)
        │  push 또는 수동 실행                             └─ docs/                ← GitHub Pages 대시보드
        ▼
   GitHub Actions (data1 안에서 실행)
   1. data1 체크아웃   2. tus4/test 체크아웃(도구)   3. Excel → 파생 데이터(JSON) + README 블록 생성
   4. 변경이 있을 때만 tus4/test 에 커밋·푸시 (DASHBOARD_PUSH_TOKEN 사용)   5. GitHub Pages 가 docs/ 를 자동 재배포
```

- data1 에는 **워크플로 파일 1개 + 시크릿 1개**만 추가합니다. 도구·템플릿·테스트는 전부 이 공개 저장소에 있습니다.
- **원본 Excel 파일은 절대 공개 저장소로 복사되지 않습니다.** 공개되는 것은 파생 데이터(열 요약, 집계, 차트 데이터, 설정된 범위의 행 데이터)뿐입니다. 자세한 범위는 6단계를 보세요.
- 데이터가 바뀌지 않았으면 새 커밋을 만들지 않습니다. 워크북이 깨져 있으면 실행이 실패하고, 이미 공개된 대시보드는 그대로 유지됩니다.

## 1단계. 공개 저장소(tus4/test)에 이 파일들 올리기

`tus4/test` 는 비어 있으므로 이 폴더의 내용 전체를 **기본 브랜치**(보통 `main`) 루트에 push 합니다.

```bash
git clone https://github.com/tus4/test.git
cd test
# 이 폴더의 파일들을 복사한 뒤
git add -A
git commit -m "Add dashboard tooling"
git push origin main
```

| 파일 | 용도 |
|---|---|
| `dashboard_builder/` | Excel → 파생 데이터 · README 블록 · `docs/` 생성기와 게시(commit/push) 도구. `timeseries.py` 가 시계열 시트를 감지·정제·집계합니다 |
| `template/index.html` | 표 대시보드 페이지 템플릿 (외부 라이브러리 없음, 내려받아 `file://` 로 열어도 동작) |
| `template/market.html` | 시계열 대시보드 페이지 템플릿 (외부 라이브러리 없음; 데이터를 `fetch` 하므로 웹 서버에서 열어야 함) |
| `tests/`, `requirements-dev.txt` | pytest 테스트 (테스트 안에서 합성 워크북을 만들어 검사) |
| `.github/workflows/ci.yml` | 이 저장소의 CI (테스트 + 샘플 빌드) |
| `docs/` | 초기 상태의 대시보드 (⚠️ **합성 샘플 데이터** — 실제 데이터가 게시되면 교체됨) |
| `data1/publish-dashboard.yml`, `data1/dashboard.config.example.json` | data1 에 복사할 파일의 사본 (4단계, 6단계) |
| `requirements.txt` | `openpyxl` (pandas·Node 불필요) |

push 하면 CI(`ci.yml`)가 자동으로 실행되어 테스트가 통과하는지 확인합니다. `docs/` 에는 샘플 데이터로 만든 예시 대시보드가 들어 있어서 Pages 설정 직후부터 페이지가 열립니다.

## 2단계. 세분화된 개인용 액세스 토큰(Fine-grained PAT) 만들기

data1 의 워크플로가 tus4/test 에 push 하려면 토큰이 필요합니다. 워크플로의 기본 토큰(`GITHUB_TOKEN`)은 워크플로가 있는 저장소(data1)에만 권한이 있으므로 다른 저장소에는 쓸 수 없습니다. **공개 저장소 하나에만, 내용 쓰기 권한만** 줍니다.

1. GitHub 오른쪽 위 프로필 사진 → **Settings**
2. 왼쪽 메뉴 맨 아래 **Developer settings**
3. **Personal access tokens** → **Fine-grained tokens** → **Generate new token**
4. 다음과 같이 입력합니다.
   - **Token name**: 예) `dashboard-publish`
   - **Expiration**: 원하는 만료일 (만료되면 워크플로가 실패하므로, 만료 전에 새 토큰을 만들어 3단계의 시크릿을 갱신해야 합니다)
   - **Resource owner**: `tus4`
   - **Repository access**: **Only select repositories** → `tus4/test` **만** 선택 (data1 은 선택하지 않습니다)
   - **Permissions** → **Repository permissions** → **Contents**: **Read and write**
     (Metadata: Read-only 는 자동으로 함께 부여됩니다. 다른 권한은 필요 없습니다.)
5. **Generate token** 을 누르고 표시된 토큰 값(`github_pat_...`)을 복사합니다. (이 화면을 벗어나면 다시 볼 수 없습니다.)

## 3단계. data1 에 시크릿 추가

1. `https://github.com/tus4/data1` → **Settings**
2. 왼쪽 **Security** 영역의 **Secrets and variables** → **Actions**
3. **Secrets** 탭에서 **New repository secret**
   - **Name**: `DASHBOARD_PUSH_TOKEN` (정확히 이 이름이어야 합니다 — 워크플로 파일이 이 이름을 사용합니다)
   - **Secret**: 2단계에서 복사한 토큰 값
4. **Add secret**

## 4단계. data1 에 워크플로 파일 복사

1. 이 저장소의 `data1/publish-dashboard.yml` 을 data1 의 **`.github/workflows/publish-dashboard.yml`** 경로에 넣고 **기본 브랜치**에 커밋·푸시합니다.

   ```
   tus4/data1/
   └─ .github/
      └─ workflows/
         └─ publish-dashboard.yml
   ```

2. 필요하면 이 저장소의 `data1/dashboard.config.example.json` 을 참고해 data1 **루트**에 `dashboard.config.json` 을 만듭니다 (선택, 6단계 참고).
3. 워크플로 파일 안에서 고칠 것은 없습니다. 공개 저장소 이름이 다르면 파일 위쪽의 `PUBLIC_REPO: tus4/test` 한 줄만 수정하세요. 기본 브랜치 이름(`main`, `master` 등)은 자동으로 감지합니다.

이 워크플로는 다음 경우에 실행됩니다.

- 기본 브랜치에 push 된 커밋이 `*.xlsx` / `*.xlsm`(어느 폴더든, 확장자 대소문자 무관 — `.XLSX`, `.XLSM` 도 포함), `dashboard.config.json`, 워크플로 파일 자체 중 하나를 추가·수정·삭제했을 때. (그래서 워크플로 파일을 처음 올리는 커밋에서 바로 첫 실행이 됩니다.) Excel 임시 잠금 파일 `~$이름.xlsx` 는 무시합니다. 빌드 자체는 확장자의 대소문자를 구분하지 않으므로 `.Xlsx` 처럼 특이한 표기의 파일도 (다음 실행 때) 함께 게시됩니다.
- Actions 탭에서 수동으로 실행했을 때 (7단계).

기본 브랜치가 아닌 브랜치나 태그에 push 하면 작업이 건너뛰어집니다 (다른 브랜치의 데이터가 공개되지 않도록). 워크북을 삭제한 커밋도 실행 대상이며, 그 경우 대시보드에서 해당 파일이 사라집니다. 워크플로 파일은 기본 브랜치에 있어야 수동 실행(`Run workflow`)이 가능합니다.

## 5단계. tus4/test 에서 GitHub Pages 켜기

1. `https://github.com/tus4/test` → **Settings**
2. 왼쪽 **Code and automation** 영역의 **Pages**
3. **Build and deployment** → **Source**: **Deploy from a branch**
4. **Branch**: 기본 브랜치(`main`) 선택, 그 옆 폴더 드롭다운에서 **`/docs`** 선택 → **Save**
5. 몇 분 뒤 `https://tus4.github.io/test/` 에서 대시보드가 열립니다. (변경 사항이 반영되기까지 최대 10분 정도 걸릴 수 있습니다. 처음에는 합성 샘플 데이터가 보이며, 상단에 노란 안내가 표시됩니다.)

`docs/.nojekyll` 파일이 포함되어 있어 Pages 가 Jekyll 빌드를 건너뛰고 `docs/` 를 그대로 게시합니다. 이후에는 data1 워크플로가 `docs/` 를 push 할 때마다 GitHub 가 자동으로 재배포합니다 (별도 워크플로를 추가할 필요 없음). `docs/` 폴더를 지우면 Pages 빌드가 실패하므로 지우지 마세요. 워크플로가 이 폴더를 매번 통째로 교체합니다 (사용자 지정 도메인 파일 `docs/CNAME` 은 보존됩니다).

## 6단계. 공개 범위 제한하기 (`dashboard.config.json`) — 꼭 읽어 주세요

**공개 저장소에 push 되는 것은 전부 인터넷에 공개됩니다.** 무엇이 공개되는지 정확히 알고 필요하면 제한하세요.

### 기본 정책 (설정 파일이 없을 때)

| 항목 | 기본값 |
|---|---|
| 게시 대상 | data1 의 모든 `*.xlsx`, `*.xlsm` (잠금 파일 `~$*` 제외), 모든 **보이는** 시트. 숨김 시트는 내용도 이름도 공개하지 않고 개수만 표시합니다. |
| 행 데이터 | **공개함** (`"mode": "rows"`) — 시트당 처음 **1,000행**, **50열**까지, 텍스트 셀은 **200자**까지 |
| 통계·차트 | 열별 값 개수/합계/평균/최소/최대/날짜 범위/고유값 수, 텍스트 열의 상위 값 목록(고유값 50개 이하인 범주형 열만), 범주별 건수·합계 차트, 날짜별 추이 차트 (처음 200,000행 기준) |
| 개인정보 추정 열 | **자동 제외** — 열 이름에 `주민, 여권, 전화, 휴대폰, 핸드폰, 연락처, 이메일, 메일주소, 주소, 계좌, 카드번호, 비밀번호, 생년월일, email, phone, mobile, address, password, ssn, passport, birth` 등이 들어가거나, 값 대부분이 이메일/전화번호/주민번호 형태인 텍스트 열. 제외된 열은 **이름과 이유만** 표시됩니다. |
| 게시되지 않는 것 | 원본 Excel 파일, 수식 자체, 서식, 셀 메모, 숨김 시트, 설정으로 제외한 파일·시트의 **이름**(개수만 표시). 오류 셀(`#N/A` 등)과 `-` 같은 자리표시 값은 빈 값으로 처리됩니다. |

README 요약 블록과 대시보드 상단에 적용된 정책이 항상 표시됩니다. 행·열이 잘린 경우에도 "처음 N행만 공개" 처럼 명시됩니다.

### 설정 파일

행 데이터를 공개하고 싶지 않거나 일부만 공개하려면 data1 **루트**에 `dashboard.config.json` 을 만듭니다. 함께 제공된 `data1/dashboard.config.example.json` 을 복사해서 수정하면 됩니다. 모든 키는 선택이며, `_` 로 시작하는 키는 주석으로 무시됩니다. **알 수 없는 키나 잘못된 값이 있으면 빌드가 실패합니다** (오타 때문에 의도보다 많이 공개되는 일을 막기 위해서입니다). 이 파일을 바꾸는 커밋도 워크플로를 실행시키므로 바로 반영됩니다.

```json
{
  "mode": "aggregates",
  "exclude_files": ["backup/**"],
  "exclude_sheets": ["원본*"],
  "exclude_columns": ["*메모*", "내부*"],
  "allow_columns": ["담당자 이메일"]
}
```

| 키 | 의미 | 기본값 |
|---|---|---|
| `mode` | `"rows"`: 행 데이터 공개(상한 내), `"aggregates"`: 통계·차트만 공개하고 **행 데이터와 상위 값 목록은 공개하지 않음** | `"rows"` |
| `title` | 대시보드 제목 | `"데이터 대시보드"` |
| `include_files` / `exclude_files` | 저장소 루트 기준 경로 glob (예: `"backup/**"`, `"**/*_draft.xlsx"`) | 모든 xlsx/xlsm |
| `include_sheets` / `exclude_sheets` | 시트 이름 glob (대소문자 무시, 예: `"원본*"`) | 모든 보이는 시트 |
| `include_columns` / `exclude_columns` | 헤더 텍스트 glob, 모든 시트에 적용 (예: `"*메모*"`) | 모든 열 |
| `auto_exclude_sensitive` | 개인정보 추정 열 자동 제외 켜기/끄기 | `true` |
| `sensitive_patterns` | 개인정보 추정에 쓰는 열 이름 조각 목록을 통째로 바꿈 | 위 표의 목록 |
| `allow_columns` | 개인정보로 추정되어 자동 제외된 열을 **그래도 공개**할 때 헤더 이름(glob) 나열 | `[]` |
| `max_rows` / `max_columns` | 시트당 공개 행/열 수 상한 | `1000` / `50` |
| `max_cell_text` | 공개하는 텍스트 셀의 최대 글자 수 | `200` |
| `max_scan_rows` | 집계·차트 계산에 사용하는 행 수 상한 | `200000` |
| `preview_rows` / `preview_columns` | README 미리보기 표 크기 | `10` / `12` |
| `header_scan_rows` | 헤더 행을 찾기 위해 살펴보는 행 수 | `20` |
| `max_chart_categories` | 차트 범주 수 상한 (초과분은 "기타"로 묶음) | `12` |
| `charts` | 차트(범주별 건수·합계, 날짜별 추이) 생성 여부. 차트에는 범주 열의 값이 라벨로 들어갑니다 | `true` |
| `publish_top_values` | 텍스트 열의 상위 값 목록 공개 여부 (`rows` 모드에서만 유효) | `true` |
| `readme.charts` / `readme.preview` | README 에 차트 / 미리보기 표 포함 여부 | `true` / `true` |
| `readme.xychart` | README 막대·선 차트를 Mermaid `xychart-beta` 로 그릴지 여부 (기본은 어디서나 표시되는 표 형태) | `false` |
| `fail_if_no_workbooks` | 워크북이 하나도 없으면 실패 처리(기존 대시보드 유지) | `true` |

- 가장 안전한 설정은 `"mode": "aggregates"` 입니다. 이 모드에서도 **열 이름, 시트 이름, 파일 경로, 열 통계, 차트의 범주 라벨**은 공개되므로, 범주 열 자체가 민감하면 `exclude_columns` 로 제외하거나 `"charts": false` 로 차트를 끄세요.
- 개인정보 추정 열의 자동 제외는 보조 장치일 뿐입니다. 열 이름이 `비고`, `메모` 처럼 일반적이면 안에 든 개인정보를 걸러내지 못합니다.

## 시계열(마켓 데이터) 시트 — 자동 감지와 시계열 대시보드

`data_info.xlsx` 의 `DailyRate` 시트처럼 **날짜별 행 × 지표별 열** 구조의 시트는 (행 수 상한이 있는) 표가 아니라 **시계열 대시보드**로 게시됩니다. 시계열 시트가 하나라도 있으면 그것이 첫 화면(`docs/index.html`)이 되고, 일반 표 시트는 `docs/tables.html` 에 나옵니다. 시계열 시트가 없으면 예전처럼 표 대시보드가 첫 화면입니다.

### 감지 규칙

- 첫 셀이 `DATES`, `DATE`, `일자`, `날짜`, `기준일`(대소문자 무관) 중 하나인 헤더 행이 있고, 그 아래 행들의 A열이 **중복 없는 날짜**(오름차순 또는 내림차순)이며, 데이터 셀이 대부분 숫자여야 합니다. 텍스트 위주의 열(지역·제품·메모 등)이 있으면 날짜순으로 정렬된 표로 보고 표 대시보드로 보냅니다.
- 헤더 **위**의 `Category` / `Ticker` / `Notation` 행(있으면)은 지표의 분류·이름·코드가 되고, 헤더 **바로 아래**의 라벨 행(예: `3월이하(당일)`, `OAS`, `현재가`)은 "필드"로 함께 표시됩니다. `Category` 가 없으면 분류는 `기타` 입니다.
- 시트의 처음 60행 안에서 찾습니다. 시트 끝의 빈 행(패딩)과 A열이 비어 있는 행은 무시하며, 같은 날짜가 두 번 나오면 뒤의 행을 씁니다.
- 지표 이름은 한글 `Ticker` 라벨(예: `미국_10y`)이 있으면 그것을, 없으면 헤더(예: `한국_3M`)를 씁니다. 라벨의 만기가 헤더의 만기와 다르면(원본의 복사 실수) 헤더의 만기로 고칩니다. 같은 접두어 + 만기(`3M`, `1Y`, `10Y` …)가 3개 이상인 지표들은 자동으로 **커브**(만기 구조)로 묶입니다.

### 게시되는 것

| 파일 | 내용 |
|---|---|
| `docs/data/market.json` | 지표 목록(이름·코드·분류·단위·소수 자릿수·원본 시트), 지표별 최근값과 1일/1주/1개월/3개월/연초/1년 전 기준값, 52주·전체 최고/최저, 1년 스파크라인, 커브 스냅샷(최근/1주 전/1개월 전/1년 전), 날짜 축, 데이터 지문 |
| `docs/data/series/<id>.json` | 지표별 **전체 이력** (표 시트의 `max_rows` 상한은 적용되지 않습니다 — 전체 이력을 그리는 것이 이 대시보드의 목적입니다) |
| README 블록 | "시계열 지표" 절: 지표 수·기간·분류, 정제된 점 수, 커브 목록, `highlights` 지표의 최근값과 변화 표 |

`include_columns` / `exclude_columns` 는 지표 이름(헤더)에 그대로 적용되고, 제외된 지표는 개수와 이름만 기록됩니다. 시계열 시트는 `mode` 설정과 무관하게 전체 이력이 공개되므로, 공개하면 안 되는 지표는 `exclude_columns` 로 빼거나 `"timeseries": {"enabled": false}` 로 시계열 게시 자체를 끄세요(그러면 상한이 있는 표로 게시됩니다).

### 단위와 변화 표시

| `Category` | 값 단위 | 변화 표시 |
|---|---|---|
| `RATE`, `CREDIT`, `HEDGE` (이름에 `SP`/`OAS`/`스프레드` 가 없는 지표) | % | bp |
| 이름·필드에 `SP`, `OAS`, `스프레드` 가 있는 지표 | bp | bp |
| `FX` | 수준 | % (표에서는 절대 변화도 툴팁으로) |
| `MACRO` | 포인트 | 포인트 |
| `DEMAND` | 금액 (원본 단위 그대로, 억·조 단위로 축약 표시) | 변화 대신 **합계** (1일/5일/20일/월간/연간) |
| 그 외·없음 | 수준 | % |

변화율은 각 지표의 **최근 관측일** 기준입니다 (1일 = 직전 관측값, 1주·1개월·3개월·1년 = 그 달력 기간 전 또는 그 이전 마지막 관측값, 연초 = 전년도 마지막 관측값). 최근 관측일이 전체 기준일보다 오래된 지표는 날짜가 주황색으로 표시됩니다.

### 정제 (`clean_outliers`)

원본 파일은 바뀌지 않습니다. 게시 데이터에서만 아래 값을 **빈 값**으로 바꾸고, 그 개수를 대시보드 하단·README 블록·빌드 로그에 표시합니다 (지표별 개수와 처음 5개 날짜는 `market.json` 의 `cleaned`, `cleaned_dates` 에 있습니다).

- 시세가 없다는 뜻으로 적힌 **0**: 첫 값이 나오기 전의 0 (예: 발행 전 국고 30년), 앞뒤 값이 모두 0.5 이상인 금리·스프레드 사이의 0, 환율·지수의 0. 정책금리나 엔화 금리처럼 0 부근을 실제로 오가는 값은 그대로 둡니다.
- **하루짜리 스파이크**: 앞뒤 5영업일 값이 서로 1.5배 안에서 일관된데 그 중앙값의 5배를 넘거나 1/5 에 못 미치고 차이가 0.5 이상인 값 (예: 유로원 129,210 → 결측, 달러위안 135.73 → 결측). 추세가 실제로 바뀐 경우(앞뒤 값이 일관되지 않음)는 건드리지 않습니다.
- 수급(`DEMAND`) 지표는 정제하지 않습니다. 끄려면 `"timeseries": {"clean_outliers": false}`.

### 설정 키 (`dashboard.config.json` 의 `"timeseries"` 객체)

| 키 | 의미 | 기본값 |
|---|---|---|
| `enabled` | 시계열 감지·게시 켜기/끄기 (끄면 일반 표로 게시) | `true` |
| `highlights` | 상단 타일과 README 표에 넣을 지표를 순서대로 (헤더·Notation·이름, 대소문자 무관) | `[]` → 분류별로 돌아가며 8개 자동 선택 |
| `featured` | 개요 차트 목록. `["A", "B"]` 또는 `{"title": "제목", "series": [...], "mode": "level" \| "diff" \| "rebase"}` (차트당 최대 6개; `rebase` = 기간 초 = 100 지수화, `diff` = 기간 초 대비 변화) | `[]` → `highlights` 를 단위별로 묶어 자동 생성 |
| `clean_outliers` | 위의 정제 규칙 적용 여부 | `true` |
| `spark_points` | 스파크라인 점 수 (최근 1년) | `52` |
| `max_series` | 시트당 지표 수 상한 | `1000` |

`highlights` / `featured` 에서 찾지 못한 이름은 빌드 로그(`::warning::`)와 대시보드 상단에 경고로 표시되며 빌드는 실패하지 않습니다. 이름이 바뀐 지표는 그 목록만 고치면 됩니다.

### 페이지 구성

1. **개요** — `highlights` 타일(최근값, 1일 변화, 1주/1개월/연초/1년 변화, 1년 스파크라인)과 `featured` 개요 차트. 타일을 누르면 탐색 차트에 추가됩니다.
2. **커브** — 만기별 최근 / 1주 전 / 1개월 전 / 1년 전 곡선, 장단기 스프레드, 만기별 표.
3. **지표 탐색** — 분류·검색으로 고른 최대 6개 지표를 함께 그립니다. 원값 / 기간 초 대비 변화 / 지수화(=100) 전환, 표 보기, 선택 구간 CSV 내려받기. 단위가 다른 지표를 함께 고르면 지수화만 허용됩니다(이중 축은 쓰지 않습니다). 선택 상태는 주소창의 `#explorer?s=…` 에 남아 공유할 수 있습니다.
4. **전체 지표** — 모든 지표의 최근값·변화·52주 범위·스파크라인 표. 열 제목으로 정렬, 분류·검색으로 필터.

상단의 **차트 기간**(3M~전체)은 모든 차트에 한 번에 적용됩니다. 라이트/다크 모드는 시스템 설정을 따르며 오른쪽 위 버튼으로 바꿀 수 있습니다. 시계열 페이지는 `data/market.json` 과 지표별 파일을 `fetch` 로 읽으므로 `file://` 로 직접 열면 동작하지 않습니다 — 로컬에서는 `python -m http.server --directory docs` 로 여세요.

## 7단계. 실행

- **자동**: data1 기본 브랜치에 `.xlsx`/`.xlsm` 파일 또는 `dashboard.config.json` 을 추가·수정·삭제하고 push 하면 실행됩니다. 같은 실행 안에서 공개 저장소까지 갱신됩니다 (보통 1~2분).
- **수동**: `tus4/data1` → **Actions** 탭 → 왼쪽에서 **Publish dashboard** 선택 → 오른쪽 **Run workflow** → 브랜치 선택 → **Run workflow**. 수동 실행은 **선택한 브랜치의 데이터**를 게시하므로, 특별한 이유가 없으면 기본 브랜치를 선택하세요.
  - **dry_run** 을 켜면 빌드와 변경 내용 요약(어떤 파일이 바뀌는지)만 로그에 보여 주고 공개 저장소에는 커밋·push 하지 않습니다. 실제로 공개하기 전에 공개 범위를 확인하는 용도로 쓰세요. (시크릿은 dry_run 에서도 필요합니다.)

## 8단계. 잘 됐는지 확인

1. `tus4/data1` → **Actions** 에서 실행이 초록색(성공)인지 확인합니다. 로그의 "Build dashboard" 단계에 파일·시트별 행 수와 제외된 열이 출력되고, "Publish" 단계에 `Pushed to origin/main` 또는 `No changes to publish` 가 나옵니다.
2. `tus4/test` 에 `Update dashboard (data1@커밋해시)` 커밋이 생겼는지 확인합니다. 데이터가 바뀌지 않았으면 커밋이 없습니다 (정상).
3. `https://github.com/tus4/test` 첫 화면(README)에 방금 올린 파일 이름·시트·KPI 표·원형 차트·미리보기가 보이는지 확인합니다. README 의 `<!-- DASHBOARD:START -->` 와 `<!-- DASHBOARD:END -->` 사이만 자동으로 바뀌고, 그 바깥에 쓴 내용은 그대로 유지됩니다.
4. `tus4/test` → **Actions** 에 `pages build and deployment` 실행이 끝난 뒤 `https://tus4.github.io/test/` 를 새로고침합니다 (브라우저 캐시 때문에 강력 새로고침이 필요할 수 있음). 시계열 시트가 있으면 첫 화면 위쪽의 **기준일**이 새 데이터의 마지막 날짜로 바뀌고 타일·차트가 갱신되어 있으면 완료입니다 (표 시트만 있으면 노란 "합성 샘플" 안내가 사라지고 실제 파일이 목록에 보입니다). 표 대시보드 `docs/tables.html` 은 내려받아 브라우저에서 직접 열어도 동작합니다.

## 문제 해결

| 증상 | 원인·조치 |
|---|---|
| 첫 단계에서 `DASHBOARD_PUSH_TOKEN is not set` | 3단계 시크릿이 없거나 이름이 다릅니다. |
| tus4/test 체크아웃이 인증 오류로 실패하거나 `cannot reach branch 'main'` | 토큰 만료 여부, 저장소 선택(`tus4/test`), 권한(**Contents: Read and write**) 확인 후 토큰을 새로 만들어 시크릿 값을 갱신하세요. |
| `no workbook (*.xlsx / *.xlsm) found` | data1 에 워크북이 없거나 `include_files`/`exclude_files` 가 모두 제외했습니다. 기존 대시보드는 유지됩니다. |
| `cannot open workbook (...)` | 해당 파일이 손상되었거나 암호로 보호되었거나 Excel 파일이 아닙니다. 메시지에 파일 이름이 표시되며 기존 대시보드는 유지됩니다. 파일을 고치거나 `exclude_files` 로 제외하세요. |
| `dashboard.config.json: unknown key(s)` | 설정 파일 키 이름 오타입니다 (위 표 참고). |
| 시계열 시트가 표로 게시됨 (`state: ok`) | 헤더 첫 셀이 `DATES`/`일자` 등이 아니거나, 날짜가 중복·뒤섞여 있거나, 텍스트 열이 섞여 있습니다 (위 "감지 규칙"). 반대로 표가 시계열로 게시되면 `"timeseries": {"enabled": false}` 로 끄세요. |
| 대시보드에 "찾지 못한 지표" 경고 | `highlights` / `featured` 의 이름이 헤더·Notation 과 다릅니다. 전체 지표 표에서 코드(헤더)를 확인하세요. |
| 시계열 페이지가 "데이터를 불러오지 못했습니다" | `file://` 로 열었습니다. 웹 서버(GitHub Pages, `python -m http.server`)에서 여세요. |
| 워크플로가 "skipped" | 기본 브랜치가 아닌 브랜치에 push 했습니다. 기본 브랜치에 push 했는데도 skipped 라면 워크플로 파일의 `if:` 줄에서 `github.ref_name == github.event.repository.default_branch` 를 `github.ref_name == 'main'` 처럼 실제 기본 브랜치 이름으로 바꾸세요. |
| 실행은 성공했는데 커밋이 없음 | 데이터가 바뀌지 않아 결과가 이전과 동일합니다 (정상). |
| 수식 열이 비어 있음 | Excel 에서 저장한 파일만 수식의 계산 결과가 들어 있습니다 (다른 프로그램으로 만든 파일은 계산 결과가 없음). Excel 에서 열어 저장하면 포함됩니다. |
| 필요한 열이 사라짐 (`자동 제외한 열` 안내) | 열 이름이나 값이 개인정보로 추정되었습니다. 공개해도 되는 열이면 `allow_columns` 에 헤더 이름을 적으세요. |
| Pages 주소가 404 | 5단계 설정(브랜치 + `/docs`) 확인, `docs/index.html` 커밋 여부 확인, 배포 완료까지 최대 10분 대기. |
| README 의 차트가 그려지지 않음 | GitHub 은 Mermaid `pie` 차트를 지원합니다. `readme.xychart: true` 를 켰다면 다시 꺼 보세요. |
| 열 이름이 `(C)` 처럼 표시됨 | 헤더 셀이 비어 있는 열은 Excel 열 문자로 표시됩니다. |

## 공개 노출에 대한 주의

- `tus4/test` 는 **공개 저장소**입니다. 그곳에 push 된 모든 내용(README, `docs/data/dashboard.json` 의 행 데이터, 그리고 **커밋 이력**)은 인터넷에서 누구나 볼 수 있습니다. 한 번 push 된 데이터는 이후 커밋으로 지워도 이력에 남으므로, **처음부터** `dashboard.config.json` 으로 공개 범위를 정하고 `dry_run` 으로 미리 확인한 뒤 실행하세요.
- GitHub Pages 사이트는 저장소가 비공개여도 인터넷에 공개됩니다. 이 구성에서는 저장소 자체도 공개이므로 Pages 로 서비스되는 내용은 모두 공개 정보입니다.
- 공개되는 워크북의 파일 경로와 시트 이름, 열 이름은 공개됩니다 (제외된 파일·시트는 개수만, 제외된 열은 이름과 이유만 표시). 데이터 지문은 복원 불가능한 해시입니다.
- **시계열 시트는 모든 지표의 전체 이력**이 `docs/data/series/` 에 공개됩니다 (`max_rows` 상한 미적용). 원본 데이터의 이용 약관(재배포 제한 등)을 확인하고, 공개하면 안 되는 지표는 `exclude_columns` 로 제외하세요.
- 토큰(`DASHBOARD_PUSH_TOKEN`)은 `tus4/test` 의 내용 쓰기 권한만 가지므로, 유출되더라도 data1 은 읽을 수 없습니다. 만료일을 짧게 두고 주기적으로 갱신하는 것을 권장하며, 시크릿 외의 곳에 적어 두지 마세요.

## 로컬에서 직접 실행해 보기

```bash
pip install -r requirements.txt -r requirements-dev.txt
python -m pytest -q                                                             # 테스트
python -m dashboard_builder.build --source ../data1 --out /tmp/staging --pages-url https://tus4.github.io/test/
python -m dashboard_builder.publish --staging /tmp/staging --repo . --dry-run   # 무엇이 바뀌는지만 보기
python -m dashboard_builder.publish --staging /tmp/staging --repo . --no-push   # 커밋만 하고 푸시하지 않음
python -m dashboard_builder.sample --repo-root .                                # 샘플 docs/ + README 블록 재생성
```

`--dry-run` 을 빼면 클론의 `docs/` 와 README 블록이 실제로 갱신되므로 결과를 `git diff` 로 확인할 수 있습니다. `docs/index.html` 은 브라우저에서 파일로 바로 열어도 동작합니다.
