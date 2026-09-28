# AGENTS.md — PowerPointLite 작업 인수인계

이 문서는 새 ChatGPT/Codex/다른 AI 세션이 과거 대화를 다시 복원하지 않고 바로 개발을 이어가기 위한 기준 문서다.

## 1. 저장소 / 현재 기준

```text
Repository: Huihuing/PowerPointLite
Stable branch: main
Active office/editor branch: feature/office-foundation
Presentation feature branch: feature/presentation-tools
```

`main`의 현재 안정 기반은 Windows용 PPTX/PPTM Viewer다.

`feature/office-foundation`에서는 다음 방향으로 확장 중이다.

```text
PPTX  Viewer + experimental Writer/Editor
DOCX  experimental Reader/Writer/Editor
XLSX  experimental Reader/Writer/Editor
HWPX  experimental Reader/Writer/Editor
HWP   experimental read-only Reader
ODF   planned
PDF   planned export
```

현재 안정 버전 표기는 **1.3 Viewer**다. feature 브랜치 기능은 Windows 실빌드와 외부 프로그램 호환 검증 전까지 안정 기능으로 표시하지 않는다.

## 2. 최우선 순서

항상 다음 순서를 지킨다.

```text
사용자 문서 손실 방지
→ 저작권/상표/폰트 라이선스 안전
→ 기존 PPTX Viewer 회귀 방지
→ Windows .NET Framework 실제 빌드 안정성
→ 새 기능
```

임의 외부 파일을 현재 내부 모델보다 좁은 형태로 다시 저장해서 내용을 날리는 기능은 추가하지 않는다.

## 3. 저작권 / 상표 / 자산 규칙

먼저 읽는다.

```text
docs/LEGAL_ASSET_POLICY.md
docs/UI_DESIGN_GUIDE.md
docs/ARCHITECTURE_ROADMAP.md
docs/EDITOR_FOUNDATION.md
docs/HWPX_FOUNDATION.md
THIRD_PARTY_NOTICES.md
licenses/README.md
```

허용 방향:

- 공개/공식 파일 형식 명세 기반 자체 Reader/Writer
- OOXML 직접 생성/편집/저장
- 공개 HWPX/HWP 명세 범위 자체 구현
- ODF 공개 표준 기반 구현
- 시스템 설치 폰트 런타임 사용

금지:

- Microsoft Office 실행파일/DLL 재배포
- Hancom 실행파일/DLL 재배포
- Office/Hancom 공식 로고/아이콘/UI 이미지 복제
- 상용 템플릿/클립아트 무단 포함
- 라이선스 불명확 TTF/OTF 번들
- 인터넷에서 가져온 타인 문서를 fixture로 커밋
- 공개 구현 소스 코드를 라이선스 검토 없이 복사

외부 라이브러리/폰트/아이콘/자산 추가 시:

```text
THIRD_PARTY_NOTICES.md
licenses/
```

를 함께 갱신한다.

## 4. 폰트 안전 계층

관련 파일:

```text
src/FontLicensing.cs
src/FontLicenseService.cs
src/FontBundlePolicy.cs
```

정책:

- Windows 설치 font family 사용 가능
- 시스템 font binary를 앱에 복사하지 않음
- 문서에는 기본적으로 family 이름만 기록
- font embedding 기본 OFF
- OpenType/TrueType `OS/2.fsType`은 보조 정보
- 실제 LICENSE 원문이 metadata보다 우선
- metadata가 허용처럼 보여도 명시적 license review가 없으면 embedding 거부
- app font bundling 권리는 fsType만으로 판정하지 않음

PPTX/DOCX/XLSX/HWPX/PDF embedding은 향후 같은 `FontLicenseService`를 사용한다.

## 5. 빌드 제약

기본 컴파일러:

```text
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
```

fallback:

```text
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
```

현재 빌드:

```text
BUILD_EXE.cmd
→ src\*.cs
→ PowerPointLite.exe
```

`.NET 8 SDK`는 요구하지 않는다.

피해야 할 것:

- 최신 C# 전용 문법
- .NET 8 전용 API
- 필수 NuGet 의존성
- 임의 WPF/.NET 8 전환

### 과거 오류 1 — Timer

`System.Threading`과 `System.Windows.Forms`가 함께 import될 수 있다.

GUI Timer는 반드시:

```csharp
System.Windows.Forms.Timer
```

로 완전 수식한다.

### 과거 오류 2 — Graphics.DrawImage

.NET Framework overload 차이로 `RectangleF` destination이 `CS1502/CS1503`을 낸 적이 있다.

필요하면:

```csharp
Rectangle.Round(rect)
```

으로 맞는 overload를 사용한다.

### 과거/예방 규칙 3 — 구형 csc 문법

복합 증가식에 바로 메서드 호출하는 등 애매한 표현은 피한다.

예:

```csharp
xml.Append(paragraphId.ToString(CultureInfo.InvariantCulture));
paragraphId++;
```

처럼 단순 문장으로 분리한다.

Windows `BUILD SUCCESS`가 최종 기준이며 이 환경에서 컴파일하지 못했으면 성공했다고 주장하지 않는다.

## 6. Git 규칙

커밋 메시지는 반드시:

```text
한글 - English
```

예:

```text
XLSX 스프레드시트 편집기 추가 - Add clean independent XLSX spreadsheet editor
HWP 5.x 읽기 전용 기반 추가 - Add public-spec HWP 5.x read-only foundation
```

원칙:

- 의미 단위 commit
- main에 micro-commit 남발 금지
- cache/log/debug output 커밋 금지
- 안정 EXE는 Windows 실빌드/실행 검증 후에만 갱신
- feature 브랜치에서 검증 전 main merge 금지

## 7. Viewer 안정 기반

주요 파일:

```text
src/Program.cs
src/MainForm.Part01.cs ... Part05.cs
src/InternalPptxRenderer.Part01.cs ... Part08.cs
src/PresenterView.cs
src/Printing.cs
src/RecentFileStore.cs
src/ViewerDialogs.cs
```

렌더링 우선순위:

```text
1. Microsoft PowerPoint Native
2. Internal OpenXML Renderer
3. LibreOffice fallback
```

반드시 유지:

- TOC 표시 상태에서도 Fit은 `SplitContainer.Panel2` 실제 viewport 기준
- slide area wheel = previous/next
- Ctrl+wheel = zoom
- thumbnail panel wheel = list scroll
- F5 / Shift+F5
- Presenter View / Notes / Print
- startup crash log

## 8. UI 방향

관련 파일:

```text
src/UiTheme.cs
src/OfficeWorkspaceDialog.cs
src/AdvancedPresentationEditorForm.cs
src/UnifiedTextDocumentEditorForm.cs
src/SpreadsheetEditorForm.cs
src/HwpReadOnlyViewerForm.cs
```

원칙:

- original dark/flat palette
- Office/Hancom Ribbon 시각 복제 금지
- 공식 아이콘/로고 사용 금지
- 문서 형식에 맞는 일반적인 canvas/grid/list 패턴은 사용 가능
- 시스템 `Segoe UI` 참조 가능, font binary 번들 금지
- DPI 100/125/150% 고려

Workspace 단축키:

```text
Ctrl+N        New Presentation
Ctrl+Alt+N    New/Open Workspace
Ctrl+Alt+O    Open supported document
Ctrl+Alt+D    New DOCX
Ctrl+Alt+X    New XLSX
Ctrl+Alt+H    New HWPX
Ctrl+Shift+E  Edit current PPTX when safe
Ctrl+Shift+I  PPTX compatibility report
```

## 9. 공통 OOXML / OPC

관련 파일:

```text
src/OpcPackage.cs
src/OpcPreservation.cs
```

현재:

- XML part read/write
- `[Content_Types].xml`
- relationships
- target part resolution
- content type mapping
- preservation 초기 기반

PPTX/DOCX/XLSX가 이 공통 구조를 공유한다.

기존 PPTX Renderer를 한 번에 rewrite하지 말고 회귀 테스트 가능한 단위로 옮긴다.

## 10. PPTX Writer / Editor

주요 model:

```text
PresentationDocument
 └─ PresentationSlide[]
     ├─ PresentationTextBox[]
     ├─ PresentationShape[]
     ├─ PresentationImage[]
     └─ PresentationTable[]
```

관련 파일:

```text
src/PresentationModel.cs
src/PresentationEditSession.cs
src/PptxWriter.cs
src/PptxEditableReader.cs
src/PptxCompatibilityAnalyzer.cs
src/AdvancedPresentationEditorForm.cs
src/AdvancedEditorThumbnailExtension.cs
src/AdvancedEditorTableExtension.cs
src/PptxWriterDiagnostics.cs
```

현재 experimental 기능:

- multiple slides
- text boxes
- text style/color/alignment
- system font picker
- shapes
- PNG/JPEG/GIF/BMP image insertion
- basic tables
- object selection / drag / resize
- keyboard nudge
- slide thumbnails
- snapshot undo/redo
- text/shape/image copy-paste
- staged save/backup
- guarded existing PPTX editing

매우 중요:

**현재 model이 표현하지 못하는 외부 PPTX를 단순화해 덮어쓰지 않는다.**

unknown/unsupported part preservation이 충분하지 않으면 편집을 거부한다.

## 11. DOCX / 공통 TextDocument

관련 파일:

```text
src/TextDocumentModel.cs
src/DocxReader.cs
src/DocxWriter.cs
src/DocxEditSafety.cs
src/TextDocumentEditSession.cs
src/UnifiedTextDocumentEditorForm.cs
src/DocxDiagnostics.cs
```

공통 model:

```text
TextDocument
 └─ DocumentParagraph[]
     └─ DocumentTextRun[]
```

현재:

- paragraph/run
- font family/size
- bold/italic/underline
- text color
- paragraph alignment
- create/read/edit/save
- project-generated DOCX safety guard

향후:

- tables/images
- headers/footers
- page layout
- lists/styles
- footnotes/comments
- unknown-part preservation

## 12. XLSX

관련 파일:

```text
src/SpreadsheetCore.cs
src/SpreadsheetEditorForm.cs
RUN_XLSX_SELFTEST.cmd
```

현재 model:

- workbook / multiple sheets
- sparse cells
- Text / Number / Boolean / Formula

현재 Editor:

- DataGridView grid
- worksheet tabs
- add/delete/move/rename sheet
- formula/input bar
- multi-cell clear
- Save / Save As

현재 formula는 저장/읽기 대상이다. 자체 계산 엔진이 완성됐다고 주장하지 않는다.

향후:

- formula calculation
- styles/number formats
- row/column sizes
- merged cells
- range copy/paste
- CSV
- charts
- comments/data validation
- XLSM preserve-only

## 13. HWPX

상세:

```text
docs/HWPX_FOUNDATION.md
src/HwpxCore.cs
src/HwpxDiagnostics.cs
RUN_HWPX_SELFTEST.cmd
```

공개 HWPX/OWPML 구조를 바탕으로 독립 구현한다.

현재 package:

```text
mimetype = application/hwp+zip
version.xml
META-INF/container.xml
META-INF/manifest.xml
Contents/content.hpf
Contents/header.xml
Contents/section0.xml
Preview/PrvText.txt
```

현재 Reader/Writer:

- paragraph/run text
- tabs/line breaks
- font references
- font size
- bold/italic/underline
- color
- paragraph alignment
- shared `TextDocument`
- unified DOCX/HWPX editor

Safety:

- current project-generated package만 제한적으로 edit
- unsupported part가 있으면 deny
- encryption metadata가 있으면 deny
- arbitrary external HWPX rewrite 금지

**HWPX structural self-test 통과와 실제 Hancom 호환은 다른 문제다.**

실제 Hancom open/save 검증 전 안정 지원으로 표시하지 않는다.

## 14. HWP 5.x read-only

관련 파일:

```text
src/CompoundFileReader.cs
src/HwpReader.cs
src/HwpReadOnlyViewerForm.cs
src/HwpDiagnostics.cs
RUN_HWP_PARSER_SELFTEST.cmd
```

현재 경로:

```text
Compound File Binary
 ↓
FAT / DIFAT / MiniFAT / Directory
 ↓
FileHeader
 ↓
BodyText/SectionN
 ↓ optional raw DEFLATE
HWP record parser
 ↓
HWPTAG_PARA_TEXT (67)
 ↓
TextDocument plain text
```

현재 `HwpReader`는:

- `HWP Document File` signature 검사
- version/flags 검사
- compressed BodyText inflate
- paragraph text record 추출
- `PrvText` fallback
- 암호/배포/DRM/인증서 암호화 문서는 거부

현재 HWP는 **읽기 전용**이다.

아직 주장하지 않는 것:

- full formatting
- tables/images
- controls/fields 완전 처리
- encrypted/distribution documents
- HWP Writer

랜덤 인터넷 HWP를 테스트 fixture로 커밋하지 않는다. 실제 문서 테스트는 사용 권한이 있는 파일만 사용한다.

## 15. 테스트

Windows 빌드 후 전체:

```bat
BUILD_EXE.cmd
RUN_ALL_FORMAT_SELFTESTS.cmd
```

개별:

```bat
RUN_WRITER_SELFTEST.cmd
RUN_DOCX_SELFTEST.cmd
RUN_XLSX_SELFTEST.cmd
RUN_HWPX_SELFTEST.cmd
RUN_HWP_PARSER_SELFTEST.cmd
```

Structural tests:

- PPTX create/read/edit/write
- DOCX create/read/edit/write
- XLSX create/read/edit/write
- HWPX create/read/edit/write package structure
- HWP FileHeader/record parser synthetic test

실제 외부 호환 검증도 별도로 한다.

```text
PPTX → PowerPoint / LibreOffice
DOCX → Word / LibreOffice
XLSX → Excel / LibreOffice
HWPX → Hancom
HWP  → rights-cleared actual HWP 5.x samples
```

이 환경에서 Windows EXE를 컴파일하지 못하면 성공했다고 주장하지 않는다.

## 16. 현재 다음 우선순위

`feature/office-foundation` 기준:

1. Windows `BUILD_EXE.cmd` 전체 compile 오류 0 확인
2. `RUN_ALL_FORMAT_SELFTESTS.cmd`
3. PPTX/DOCX/XLSX/HWPX 외부 프로그램 compatibility validation
4. HWP rights-cleared real sample validation
5. HWP table/image/control Reader 확대
6. XLSX styles / number formats / formula calculation
7. DOCX table/image/page layout
8. HWPX table/image/multiple-section / unknown-part preservation
9. multi-format recent files / visible Workspace toolbar entry
10. ODF package / ODT/ODS/ODP
11. PDF Export + FontLicenseService

상세 체크리스트는 `docs/ARCHITECTURE_ROADMAP.md`를 따른다.

**기능 개수보다 기존 Viewer 회귀 방지, 저작권/라이선스 안전, 사용자 문서 보존, 실제 Windows 빌드 검증이 우선이다.**
