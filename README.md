# Excel Editor

> **macOS 전용입니다.** Windows·Linux에서는 빌드도 실행도 해 본 적이 없습니다.
> 자세한 내용은 [지원 플랫폼](#지원-플랫폼) 참고.

*A macOS-only spreadsheet editor for xlsx and CSV. Documentation is in Korean.*

데스크톱 스프레드시트 편집기. xlsx와 CSV/TSV를 열어 **값과 서식을 확인·수정하고
같은 파일에 덮어쓴다.** Numbers가 xlsx를 열 때 서식을 뭉개고 저장 시 변형하는 문제를
피하려고 만들었다.

Finder에서 `.xlsx` 를 더블클릭하면 이 앱으로 열리고, 창에 파일을 끌어다 놓아도 열린다.
최근 연 파일은 툴바의 **최근** 메뉴에 남는다. 저장할 때는 기본적으로 원본을
`이름.bak.xlsx` 로 복사해 둔다(툴바에서 끌 수 있다).

## 스택

| 영역 | 선택 | 이유 |
|---|---|---|
| 셸 | Tauri 2 (Rust) | 네이티브 파일 다이얼로그, 실제 덮어쓰기, ~10MB 바이너리 |
| 그리드 | [Univer](https://github.com/dream-num/univer) 1.0.3 (Apache-2.0) | 수식 입력줄·서식 툴바·병합셀·수식 엔진 내장 |
| xlsx I/O | [ExcelJS](https://github.com/exceljs/exceljs) 4.4.0 | 셀 단위 서식 읽기/쓰기 지원 |
| zip 검사 | fflate | 저장 시 사라질 기능을 미리 탐지 |

Univer의 공식 xlsx 변환은 별도 Docker 서버(`univer-server`)를 요구하므로,
**ExcelJS ↔ Univer 스냅샷 변환기를 직접 구현**했다 (`src/xlsx/`).

## 지원 플랫폼

**macOS 전용.** Apple Silicon·Intel 둘 다 빌드·실행을 확인했다.

Windows·Linux는 **빌드조차 시도해 본 적이 없다.** 기반 기술(Tauri, Univer, ExcelJS)은
전부 크로스 플랫폼이고 CP949 디코딩도 표준 웹 인코딩이라 원리상 돌아갈 구조지만,
확인하지 않은 것을 "된다"고 적지는 않겠다.

확실히 아는 공백이 하나 있다. **파일을 더블클릭해서 여는 경로가 macOS 전용으로
구현돼 있다** — macOS는 OS가 `RunEvent::Opened` 이벤트로 경로를 넘기고 그것만 처리한다
(`src-tauri/src/lib.rs`). Windows·Linux는 경로를 커맨드라인 인자로 넘기는데 그 처리가
없다. 그 환경에서 빌드하면 앱은 뜨지만 더블클릭한 파일이 열리지 않는다. 열기 버튼과
드래그앤드롭은 동작할 것이다.

포팅할 생각이면 거기서 시작하면 된다 — 시작 시 `std::env::args()` 를 읽어 같은 경로로
흘려보내면 되고, 앱이 이미 떠 있을 때 두 번째 파일을 여는 건 single-instance 플러그인이
추가로 필요하다.

## 필요한 것

| | 설치 |
|---|---|
| macOS | — |
| Node 20+ | [nodejs.org](https://nodejs.org) 또는 `brew install node` |
| pnpm | `corepack enable` (Node에 포함) 또는 `brew install pnpm` |
| Rust | `curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs \| sh` |
| Xcode Command Line Tools | `xcode-select --install` |

## 실행

```bash
pnpm install
pnpm tauri dev          # 데스크톱 앱 (개발 모드)
pnpm dev                # 브라우저에서 UI만 (열기=업로드, 저장=다운로드)
pnpm test               # 테스트 146개
pnpm typecheck
pnpm tauri build        # .app / .dmg 번들
```

빌드 결과물은 `src-tauri/target/release/bundle/` 에 생긴다. `.app` 을 `/Applications`
으로 옮기면 Finder에서 `.xlsx` 더블클릭 연결이 동작한다.

테스트 146개 중 2개는 실제 워크북을 지정했을 때만 돌고, 평소엔 건너뛴다:

```bash
EXCEL_EDITOR_TEST_FILE=~/어떤파일.xlsx pnpm test
```

단축키: `⌘N` 새 파일 · `⌘O` 열기 · `⌘S` 저장 · `⇧⌘S` 다른 이름으로 · `⌘F` 찾기·바꾸기

아이콘을 바꾸려면 `python3 src-tauri/icons/make-icon.py out.png` 로 1024px 소스를 만든 뒤
`pnpm tauri icon out.png` 를 돌린다. (이미징 라이브러리 없이 표준 라이브러리만 쓴다.)

## 기능

Univer 프리셋으로 **찾기·바꾸기, 열 필터, 정렬**이 리본의 `데이터` 탭에 들어 있다.
필터 범위는 xlsx의 `<autoFilter>` 와 양방향으로 오간다 — 다만 ExcelJS가 `<filterColumn>`
조건을 모델링하지 않으므로 **어떤 값이 걸러져 있었는지는 저장되지 않는다.** 필터 버튼은
돌아오고 그 뒤의 선택은 돌아오지 않는다.

## CSV / TSV

`.csv` `.tsv` `.txt` 를 열고 저장한다.

- **인코딩 자동 감지** — BOM을 먼저 보고, 없으면 엄격 UTF-8로 시도한 뒤 실패하면 CP949로
  읽는다. 순서가 중요하다: CP949는 어떤 바이트열도 받아들이므로 먼저 시도하면 진짜 UTF-8을
  조용히 깨뜨린다. 저장은 항상 UTF-8 + BOM (BOM이 없으면 Excel이 시스템 코드페이지로 읽어
  한글이 깨진다).
- **구분자 자동 감지** — `,` `\t` `;` `|` 중 모든 행을 가장 고르게 쪼개는 것을 고른다.
  값 안에 콤마가 든 탭 구분 파일도 올바르게 판정한다.
- **타입 추론은 보수적으로** — 숫자로 바꾸는 건 모호함이 없고 정확히 표현 가능할 때뿐이다.
  앞자리 0이 있으면(전화번호·우편번호·계좌번호) 텍스트로 남기고, 유효숫자 15자리를 넘는
  값(긴 ID·카드번호)도 텍스트로 남긴다. **날짜와 TRUE/FALSE는 추론하지 않는다** — 이게
  다른 도구가 데이터를 망가뜨리는 지점이다.
- 날짜 서식이 걸린 숫자는 CSV로 쓸 때 시리얼이 아니라 `2026-03-15` 로 나간다.

실측: 2.4MB·34,318행·7열 CSV에서 파싱 95ms, 쓰기 40ms, **201,708셀 전부 왕복 일치**.

## 보존되는 것

값(문자/숫자/불린/오류), 수식과 캐시된 결과, 날짜(시리얼+서식), 숫자 서식,
폰트(이름·크기·굵게·기울임·밑줄·취소선·색·위/아래 첨자), 배경색,
테두리(4변 + 대각선, 13종 선 스타일), 정렬(수평·수직·줄바꿈·회전·들여쓰기·축소맞춤),
병합셀, 열 너비, 행 높이, 숨긴 행/열, 틀 고정, 눈금선 표시, 시트 탭 색,
시트 순서와 숨김 상태, 서식 있는 리치 텍스트, 하이퍼링크.

색상은 `argb`·테마 인덱스(+tint)·레거시 56색 팔레트 세 형식을 모두 해석한다.

## 사라지는 것

차트, 피벗 테이블, 이미지·도형, 매크로(VBA), 조건부 서식, 데이터 유효성,
표(Table) 서식, 주석, 외부 링크, 슬라이서.

ExcelJS는 패키지를 자기 모델로부터 다시 써내므로 모델에 없는 것은 저장 시 사라진다.
`src/xlsx/inspect.ts` 가 파일을 열 때 zip 구성을 먼저 훑어 **어떤 기능이 사라질지 경고하고,
저장 전에 한 번 더 확인을 받는다.** 조용히 날리지는 않는다.

## 구조

```
src/
  App.tsx              툴바, 열기/저장, 최근 파일, 더티 추적, 경고 배너
  univer-worker.ts     수식 엔진 워커 (메인 스레드 차단 방지)
  fs/file.ts           Tauri dialog+fs+IPC, 브라우저 폴백
  xlsx/
    normalize.ts       ExcelJS가 못 읽는 파일 복구 (아래 참조)
    import.ts          xlsx → Univer 스냅샷
    export.ts          Univer 스냅샷 → xlsx
    style.ts           ExcelJS Style ⇄ Univer IStyleData (양방향, 한 파일)
    color.ts           argb / 테마+tint / 인덱스 팔레트 → #rrggbb
    units.ts           문자폭↔px, pt↔px, Date↔시리얼
    range.ts           A1 참조 ⇄ Univer IRange
    filter.ts          autoFilter ⇄ Univer 필터 리소스
    inspect.ts         손실될 기능 탐지
    snapshot.ts        빈 통합 문서 (파서 의존 없음)
  csv/
    encoding.ts        BOM / UTF-8 / CP949 감지, UTF-8+BOM 쓰기
    delimited.ts       RFC 4180 파서·직렬화, 구분자 감지
    convert.ts         구분자 텍스트 ⇄ Univer 스냅샷, 타입 추론

src-tauri/src/
  lib.rs               플러그인 등록, Finder 열기·드래그앤드롭 → scope 부여 후 프런트에 전달
  recent.rs            최근 파일 목록 + fs scope 재부여 + 백업 복사
```

## 구현 중 부딪힌 것들

실제 파일과 라이브러리 동작을 확인하며 우회한 지점들. 모두 테스트로 고정해 뒀다.

**`r` 속성 없는 xlsx를 ExcelJS가 못 읽는다.**
SpreadsheetML에서 `<row>`·`<c>` 의 `r`(위치) 속성은 선택이고, 없으면 문서 순서로 위치가
결정된다. Excel과 Numbers는 이를 처리하지만 ExcelJS는 `Invalid row number in model` 로
파일 전체를 거부한다. DataGrip 등 여러 내보내기 도구가 이 최소 형식을 쓴다 — 즉 **열어야 할
이유가 큰 파일이 하필 안 열린다.** `normalize.ts` 가 파싱 직전에 위치를 명시적으로 채워 넣는다.
(정규화 결과는 파서에만 전달되며 사용자 파일에 쓰이지 않는다.)

**스타일 dedup 키가 서로 다른 서식을 합쳐 버렸다.**
`JSON.stringify(style, Object.keys(style).sort())` 의 두 번째 인자 배열은 정렬 힌트가 아니라
**중첩 객체까지 재귀 적용되는 속성 허용목록**이다. 그래서 `{n:{pattern:'yyyy-mm-dd'}}` 와
`{n:{pattern:'#,##0.00'}}` 가 둘 다 `{"n":{}}` 로 직렬화돼 날짜 셀이 통화 서식을 물려받았다.
키를 재귀적으로 정렬하는 직렬화 함수로 교체했다.

**병합된 셀의 값이 복제됐다.**
ExcelJS는 병합 영역의 종속 셀에서 `cell.value` 를 읽으면 **master의 값을 돌려준다.**
그대로 가져오면 Univer 쪽에 같은 값이 영역 전체에 박히고, 병합을 풀면 드러난다.
종속 셀은 서식만 가져온다.

**셀도 높이도 없는 숨긴 행이 저장에서 사라졌다.**
ExcelJS `Row.model` 은 `height || cells.length` 가 거짓이면 `null` 을 반환해 행을 아예
직렬화하지 않는다. 숨김 플래그만 있는 행은 통째로 유실된다. 그런 행에는 기본 행 높이를
지정해 실체를 만들어 준다(보이는 결과는 동일).

**행 높이·열 너비가 저장할 때마다 조금씩 밀렸다.**
xlsx는 pt와 문자폭, Univer는 px를 쓰고 변환이 정확히 가역적이지 않다. 원본 값을 `custom`에
실어 보내고, px가 그대로면 원본 값을 다시 써서 **건드리지 않은 행/열은 반복 저장에도 고정**된다.
리치 텍스트와 하이퍼링크도 같은 방식으로, 텍스트가 변하지 않았다면 원본을 복원한다.

**ExcelJS는 `fullCalcOnLoad` 를 쓰기만 하고 읽지 않는다.**
검증을 리로드한 객체가 아니라 `xl/workbook.xml` XML에 대해 한다.

**CP949를 UTF-8보다 먼저 시도하면 안 된다.**
CP949 디코더는 어떤 바이트열이든 받아들여 절대 실패하지 않는다. 먼저 시도하면 진짜 UTF-8
파일이 조용히 깨진다. 엄격(`fatal: true`) UTF-8을 먼저 시도하고 **실패할 때만** CP949로 간다.

**1900-03-01 이전 날짜는 하루 어긋난다.**
Excel의 1900 체계에는 존재하지 않는 1900-02-29가 들어 있다. 선형 일수 계산은 Excel보다 1 크다.
ExcelJS도 같은 선형 변환을 쓰므로 앱 내부는 일관되고, 실제로 문제되는 경우는 없다. 테스트에 명시해 뒀다.

## 알려진 제약

- 리치 텍스트는 그리드에서 **평문으로 표시**된다(런별 서식 편집 불가). 텍스트를 고치지 않으면
  저장 시 원본 서식이 복원되고, 고치면 평문이 된다.
- 수식은 Univer 엔진이 계산한다. 함수 커버리지가 Excel 전체는 아니며, 저장 파일에
  `fullCalcOnLoad` 를 넣어 Excel이 열 때 다시 계산하게 한다.
- 수식 엔진 워커가 Univer 코어를 별도 청크(~7MB)로 한 번 더 번들한다. 메인 스레드
  차단을 막는 대가다. 로컬 앱이라 디스크에서 읽으므로 체감 비용은 작다.

## 파일 접근 권한

`fs` 권한에 **정적 경로 스코프를 두지 않았다**(`src-tauri/capabilities/default.json`).
`tauri-plugin-dialog` 이 사용자가 다이얼로그에서 고른 경로에 대해 런타임 fs 스코프에
`allow_file` 을 호출하고, `tauri-plugin-fs` 의 `resolve_path` 가
`fs_scope.scope.is_allowed(...) || scope.is_allowed(...)` 로 런타임 스코프를 함께 검사한다.
따라서 **사용자가 직접 고른 파일만** 읽고 쓸 수 있고, 파일시스템 전체를 여는 `**` 글롭이 필요 없다.

파일이 들어오는 경로가 셋으로 늘어나면서 이 불변식을 유지하는 게 설계의 핵심이 됐다:

- **다이얼로그** — 플러그인이 알아서 scope를 준다.
- **Finder 더블클릭 / 드래그앤드롭** — OS가 준 경로라 scope에 없다. `lib.rs` 가
  `RunEvent::Opened` 와 `WindowEvent::DragDrop` 을 받아 **Rust에서** scope를 부여하고
  최근 목록에 기록한 뒤 프런트엔드에 알린다.
- **최근 파일** — scope 부여는 재시작하면 사라진다. 그래서 최근 목록을 Rust가 소유하고,
  `grant_recent` 는 **그 목록에 있는 경로만** 다시 허용한다. 목록에 넣는 `record_recent` 는
  넣기 전에 `is_allowed` 로 검증하므로, 프런트엔드가 가짜 경로를 집어넣고 허용받는 우회가 막힌다.

백업 복사도 Rust에 있다. `이름.bak.xlsx` 는 사용자가 고른 적 없는 경로라서, 프런트엔드가
쓰려 하면 scope가 — 옳게 — 거부한다.

## 번들 크기

릴리스 빌드: `.app` 약 8.5MB, `.dmg` 약 6MB.

## 라이선스

MIT. 의존하는 [Univer](https://github.com/dream-num/univer)는 Apache-2.0,
[ExcelJS](https://github.com/exceljs/exceljs)와 [fflate](https://github.com/101arrowz/fflate)는
MIT이며, 모두 MIT 배포와 양립한다.

## 상표

Microsoft, Excel 은 Microsoft Corporation의 상표 또는 등록 상표다. **이 프로젝트는
Microsoft와 아무 관련이 없고, 승인·후원·제휴 관계가 아니다.** 이름의 "Excel" 은 이 앱이
어떤 파일을 다루는지 설명하기 위한 것이며, Microsoft의 로고나 브랜드 요소는 쓰지 않았다
(앱 아이콘은 `src-tauri/icons/make-icon.py` 로 직접 그린 것이다).

Apple, macOS, Numbers 는 Apple Inc.의 상표다. 마찬가지로 아무 관련이 없다.

*Microsoft and Excel are trademarks of Microsoft Corporation. This project is not
affiliated with, endorsed by, or sponsored by Microsoft. Apple, macOS and Numbers
are trademarks of Apple Inc.*
