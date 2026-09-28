# AGENTS.md — PowerPointLite 작업 인수인계

새 ChatGPT/Codex/다른 AI 세션이 기존 대화를 복원하지 않고 바로 개발을 이어가기 위한 기준 문서다.

## 1. 프로젝트 현재 상태와 장기 목표

저장소: `Huihuing/PowerPointLite`

현재 안정 기능은 Windows용 `.pptx/.pptm` Viewer다.

장기적으로는 기존 PPTX 기능을 유지하면서 다음 포맷을 읽고 작성/편집/저장할 수 있는 경량 문서 프로그램으로 확장하는 방향을 검토한다.

```text
PPTX
DOCX
XLSX
HWPX
HWP
ODT
ODS
ODP
```

우선순위는 반드시:

```text
PPTX Viewer 안정화
→ 법적/폰트/자산 안전 기반
→ PPTX Writer
→ 공통 OOXML/OPC 계층
→ PPTX Editor
→ DOCX
→ XLSX
→ HWPX
→ HWP
→ ODF / PDF Export
```

현재 안정 기준 버전: **1.3**

## 2. 현재 브랜치 전략

- `main` — Windows 실빌드/실행 검증을 통과한 안정 기반
- `feature/presentation-tools` — 레이저/펜/자동 진행/내부 미디어 등 발표 기능 개발
- `feature/office-foundation` — 멀티포맷 아키텍처, 폰트/라이선스, 독립 UI 기반 개발

두 feature 브랜치는 각각 안정화한 뒤 의미 단위로 main에 병합한다.

## 3. 저작권/상표/자산 최우선 규칙

먼저 [`docs/LEGAL_ASSET_POLICY.md`](docs/LEGAL_ASSET_POLICY.md)를 읽는다.

허용 방향:

- 공개/공식 파일 형식 명세 기반 자체 Reader/Writer
- OOXML 직접 생성/편집/저장
- 공식 공개 HWPX/HWP 명세 범위 자체 구현
- ODF 공개 표준 기반 구현
- 시스템 설치 폰트 런타임 사용

금지:

- Microsoft Office 실행 파일/DLL 재배포
- Hancom 실행 파일/DLL 재배포
- Office/Hancom 공식 로고/아이콘/UI 이미지 복제
- 상용 템플릿/클립아트 무단 포함
- 라이선스 불명확 폰트 TTF/OTF 번들
- 인터넷에서 가져온 타인 문서를 테스트 fixture로 커밋

새 외부 라이브러리/폰트/아이콘은 `THIRD_PARTY_NOTICES.md`와 `licenses/`를 갱신한다.

상표명은 호환성 설명에만 사실적으로 사용하고 앱 브랜딩은 독립적으로 만든다.

## 4. 폰트 규칙

`src/FontLicensing.cs`가 초기 폰트 안전 계층이다.

현재 원칙:

- Windows 설치 폰트 family를 열거해 사용 가능
- 시스템 폰트 파일을 앱에 복사하지 않음
- 문서에는 기본적으로 font family 이름만 기록
- font embedding은 기본 OFF
- OpenType/TrueType `OS/2.fsType`은 보조 판단 정보
- 실제 LICENSE 원문이 fsType보다 우선
- 불명확하면 embedding을 허용하지 않는 방향

`FontLicenseInfo`에서 확인할 항목:

```text
EmbeddingLevel
CanEmbed
CanEmbedForEditing
CanPreviewAndPrint
CanSubset
BitmapEmbeddingOnly
RawFsType
RequiresLicenseTextReview
```

향후 PPTX/DOCX/XLSX/HWPX/PDF embedding은 같은 `FontLicenseService`를 공유해야 한다.

## 5. UI 원칙

[`docs/UI_DESIGN_GUIDE.md`](docs/UI_DESIGN_GUIDE.md)를 따른다.

- Microsoft Office Ribbon이나 Hancom UI를 시각적으로 복제하지 않는다.
- 공식 제품 아이콘/로고를 사용하지 않는다.
- 자체 dark palette / flat controls 사용
- UI 기본 글꼴은 시스템 `Segoe UI` 사용 가능하나 폰트 파일은 번들하지 않는다.
- 100/125/150% DPI를 고려한다.
- TOC가 보일 때 Fit은 반드시 실제 `SplitContainer.Panel2` viewport 기준이다.

`src/UiTheme.cs`는 기존 Viewer 구조를 건드리지 않고 독립 테마를 덧씌우는 기반이다.

## 6. 빌드 제약

기본 컴파일러:

```text
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
```

fallback:

```text
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
```

따라서 다음을 피한다.

- 최신 C# 전용 문법
- .NET 8 전용 API
- 필수 NuGet 의존성
- WPF/.NET 8로 임의 전환

현재 `BUILD_EXE.cmd`는 `src\*.cs` 전체를 한 번에 컴파일한다.

Windows `BUILD_EXE.cmd` 실빌드가 최종 기준이다.

### 과거 오류 — Timer

`System.Threading`과 `System.Windows.Forms`가 함께 import될 수 있으므로 GUI 타이머는 반드시:

```csharp
System.Windows.Forms.Timer
```

로 명시한다.

### 과거 오류 — Graphics.DrawImage

.NET Framework `Graphics.DrawImage` 오버로드는 최신 런타임과 다르다.

7인자 호출에서 destination이 `RectangleF`일 경우 문제가 발생한 적이 있으므로 필요하면:

```csharp
Rectangle.Round(rect)
```

처럼 맞는 오버로드를 사용한다.

## 7. 현재 소스 구조

```text
src/
├─ Program.cs
├─ MainForm.Part01.cs ... MainForm.Part05.cs
├─ PresenterView.cs
├─ Printing.cs
├─ RecentFileStore.cs
├─ ViewerDialogs.cs
├─ InternalPptxRenderer.Part01.cs ... Part08.cs
├─ FontLicensing.cs
└─ UiTheme.cs
```

`MainForm.Part*.cs`는 `public sealed partial class MainForm : Form`이다.

`InternalPptxRenderer.Part*.cs`는 `internal static partial class InternalPptxRenderer`다.

private 필드/메서드는 partial class 전체에서 공유된다.

현재 `src\*.cs` 빌드 구조 때문에 실제 하위 디렉터리 분리는 아직 하지 않는다. 향후 빌드 스크립트를 재귀 소스 수집 방식으로 변경한 뒤 `Core/Formats/Renderer/Editor/UI` 폴더로 이동한다.

## 8. 현재 PPTX 렌더링 전략

우선순위:

```text
1. Microsoft PowerPoint Native
2. Internal OpenXML Renderer
3. LibreOffice fallback
```

Office/LibreOffice 바이너리를 앱 배포물에 포함하지 않는다.

### Internal OpenXML 처리 범위

- slide size/background
- theme color/font
- master/layout/placeholder
- text/basic formatting
- bullets/numbering
- images/crop/alpha/rotation/flip
- basic SVG subset / GDI+ EMF/WMF
- shapes/custom geometry 일부
- line/connectors/arrowheads
- groups/tables
- basic charts
- SmartArt static approximation
- hyperlinks/internal slide actions
- media extraction fallback
- gradient/pattern/basic shadow
- 일부 transition/animation fallback

## 9. Reader / Writer 분리 원칙

향후 포맷 코드는 반드시 다음 패턴으로 간다.

```text
PPTX → PptxReader → PresentationDocument → Renderer
                                      ↓
                                    Editor
                                      ↓
                                  PptxWriter → PPTX
```

같은 원칙:

```text
DocxReader / DocxWriter
XlsxReader / XlsxWriter
HwpxReader / HwpxWriter
HwpReader / HwpWriter
```

UI에서 포맷 XML을 직접 수정하지 않는다.

## 10. Macro 문서

향후:

```text
.pptm
.docm
.xlsm
```

을 열 수 있어도 VBA를 실행하지 않는다.

권장 동작:

```text
open
→ normal content edit
→ unknown/VBA parts preserve when practical
→ save
→ macro execution = disabled
```

보안상 매크로 실행은 기본 지원 대상이 아니다.

## 11. Viewer UX 반드시 유지할 동작

현재 기능:

- Drag & Drop / Recent
- thumbnails / TOC / Auto TOC
- actual right viewport Fit
- wheel previous/next
- Ctrl+wheel zoom
- arrow/PageUp/PageDown/Home/End
- F11
- search / go-to / sorter
- speaker notes / hidden slides
- F5 / Shift+F5
- Presenter View
- print full/notes/2·4·6 handout

### Fit 회귀 금지

TOC 폭을 포함한 전체 창 크기로 Fit 계산하면 슬라이드가 잘린다.

반드시 `SplitContainer.Panel2`의 실제 viewport를 사용한다.

## 12. 테스트 자산

테스트 문서는 프로젝트 자체 제작 자산만 사용한다.

금지:

- 인터넷 강의자료
- 회사/학교 문서
- 타인의 HWP
- 상용 PPT/XLSX template

새 포맷은 generator 기반 fixture를 우선한다.

예:

```text
TEST_BASIC_PPTX.pptx
TEST_BASIC_DOCX.docx
TEST_BASIC_XLSX.xlsx
TEST_BASIC_HWPX.hwpx
```

## 13. 테스트 절차

모든 수정 후 최소 확인:

1. brace/method duplication 정적 검사
2. `Timer` 모호성 검사
3. `Graphics.DrawImage` overload 검사
4. Windows `BUILD_EXE.cmd`
5. `BUILD SUCCESS`
6. `PowerPointLite.exe` 실행
7. 기본 PPTX 테스트
8. 실제 사용자 PPTX
9. TOC on/off Fit
10. wheel / Ctrl+wheel
11. F5 / Shift+F5
12. Presenter / Notes
13. Print
14. 새 기능별 fixture

이 환경에서 Windows EXE를 컴파일하지 못하면 성공했다고 주장하지 않는다.

## 14. Git 규칙

기본 브랜치: `main`

커밋 메시지는 반드시:

```text
한글 - English
```

예:

```text
폰트 라이선스 검사 기반 추가 - Add OpenType font licensing foundation
독립 UI 테마 기반 추가 - Add clean independent application theme
PPTX Writer 패키지 생성 추가 - Add initial PPTX writer package creation
```

원칙:

- 의미 단위 commit
- main에 micro-commit 남발 금지
- 작업 전 최신 main 확인
- cache/log/debug 산출물 금지
- 안정 EXE는 Windows 실빌드/실행 검증 후에만 갱신

## 15. 다음 구현 우선순위

`feature/office-foundation` 기준:

1. 법적/자산 문서 ✅
2. 독립 UI theme foundation ✅
3. OpenType fsType reader ✅ 초기 구현
4. FontLicenseService / font picker
5. OOXML 공통 package abstraction 설계
6. PPTX Writer MVP
7. Presentation internal model
8. basic Editor UI
9. DOCX
10. XLSX
11. HWPX
12. HWP

상세 로드맵은 [`docs/ARCHITECTURE_ROADMAP.md`](docs/ARCHITECTURE_ROADMAP.md)를 본다.

**새 기능보다 기존 PPTX Viewer 회귀 방지, 저작권/라이선스 안전, 실제 Windows 빌드 안정성이 우선이다.**
