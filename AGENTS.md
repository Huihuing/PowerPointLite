# AGENTS.md — PowerPointLite 작업 인수인계

이 문서는 새 ChatGPT/Codex/다른 AI 세션이 과거 대화를 다시 복원하지 않고 바로 개발을 이어가기 위한 기준 문서다.

## 1. 저장소 / 브랜치

```text
Repository: Huihuing/PowerPointLite
Stable branch: main
Active office/editor branch: feature/office-foundation
Presentation feature branch: feature/presentation-tools
```

`main`의 안정 기반은 Windows용 PPTX/PPTM Viewer다.

`feature/office-foundation`에는 현재 다음 experimental 기반이 있다.

```text
PPTX  Reader / Writer / advanced Editor / Viewer
DOCX  Reader / Writer / Editor
XLSX  Reader / Writer / Editor
HWPX  Reader / Writer / Editor
HWP   HWP 5.x read-only Reader
ODT   Reader / Writer / Editor
ODS   Reader / Writer / Editor
ODP   Reader / Writer / Editor
PDF   raster export
```

현재 안정 버전 표기는 **1.3 Viewer**다. feature 브랜치는 Windows 실빌드와 실제 Office/LibreOffice/Hancom/PDF Viewer 검증 전까지 안정판으로 표시하지 않는다.

## 2. 최우선 순서

```text
사용자 문서 손실 방지
→ 저작권/상표/폰트 라이선스 안전
→ 기존 PPTX Viewer 회귀 방지
→ Windows .NET Framework 실제 빌드 안정성
→ 새 기능
```

현재 model이 표현하지 못하는 외부 문서를 단순화해서 덮어쓰지 않는다.

## 3. 저작권 / 상표 / 자산

먼저 읽는다.

```text
docs/LEGAL_ASSET_POLICY.md
docs/UI_DESIGN_GUIDE.md
docs/ARCHITECTURE_ROADMAP.md
docs/EDITOR_FOUNDATION.md
docs/HWPX_FOUNDATION.md
docs/HWP_FOUNDATION.md
THIRD_PARTY_NOTICES.md
licenses/README.md
```

금지:

- Microsoft Office 실행파일/DLL 재배포
- Hancom 실행파일/DLL 재배포
- Office/Hancom 공식 로고/아이콘/UI 이미지 복제
- 상용 템플릿/클립아트 무단 포함
- 라이선스 불명확 TTF/OTF 번들
- 타인 실제 문서를 fixture로 커밋
- 라이선스 검토 없이 외부 구현 코드 복사

파일 형식은 공개/공식 규격을 참고해 독립 구현한다.

## 4. 폰트 정책

관련:

```text
src/FontLicensing.cs
src/FontLicenseService.cs
src/FontBundlePolicy.cs
```

원칙:

- 시스템에 설치된 font family를 런타임에 사용 가능
- 시스템 font file을 앱에 복사하지 않음
- 문서에는 기본적으로 family name만 기록
- embedding 기본 OFF
- OpenType/TrueType `OS/2.fsType`은 보조 정보
- 실제 LICENSE 원문이 metadata보다 우선
- 명시적 license review 없으면 embedding deny-by-default
- app bundling 권리는 fsType만으로 판정하지 않음

### PDF

현재 `src/PdfExport.cs`의 첫 PDF 경로는 **raster-page 방식**이다.

문서 글자를 GDI+로 페이지 bitmap에 렌더링하고 JPEG page image를 PDF에 넣는다. 따라서 원본 TTF/OTF font binary를 PDF에 자동 embedding하지 않는다.

향후 searchable/vector text PDF, subset/full font embedding을 추가할 때는 반드시 `FontLicenseService`를 거친다.

## 5. 빌드 제약

기본 compiler:

```text
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
```

fallback:

```text
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
```

빌드:

```text
BUILD_EXE.cmd
→ src\*.cs
→ PowerPointLite.exe
```

금지/주의:

- 최신 C# 전용 문법 사용 금지
- .NET 8 전용 API 사용 금지
- 필수 NuGet dependency 추가 금지
- WPF/.NET 8로 임의 전환 금지

과거 오류:

```text
Timer ambiguity
→ GUI Timer는 System.Windows.Forms.Timer 완전 수식

Graphics.DrawImage RectangleF overload
→ .NET Framework overload 확인, 필요하면 Rectangle.Round(rect)
```

Windows `BUILD SUCCESS`가 최종 기준이다. 이 환경에서 Windows compiler를 못 돌렸다면 성공했다고 말하지 않는다.

## 6. Git 규칙

커밋 메시지:

```text
한글 - English
```

의미 단위로 커밋한다.

- 검증 전 main merge 금지
- cache/log/debug output 금지
- 안정 EXE는 Windows 실빌드/실행 검증 후 갱신

## 7. Viewer 회귀 금지

Viewer 주요 파일:

```text
src/MainForm.Part01.cs ... Part05.cs
src/InternalPptxRenderer.Part01.cs ... Part08.cs
src/PresenterView.cs
src/Printing.cs
```

반드시 유지:

- Fit = `SplitContainer.Panel2` 실제 viewport
- slide area wheel = previous/next
- Ctrl+wheel = zoom
- thumbnail wheel = list scroll
- F5 / Shift+F5
- Presenter / Notes / Print
- startup crash log

## 8. UI

관련:

```text
src/UiTheme.cs
src/OfficeWorkspaceDialog.cs
src/MainFormWorkspaceToolbar.cs
src/AdvancedPresentationEditorForm.cs
src/UnifiedTextDocumentEditorForm.cs
src/SpreadsheetEditorForm.cs
src/OdsSpreadsheetEditorForm.cs
src/OdpPresentationEditorForm.cs
src/HwpReadOnlyViewerForm.cs
```

원칙:

- 프로젝트 자체 dark/flat palette
- Microsoft Office/Hancom Ribbon 시각 복제 금지
- 공식 제품 아이콘/로고 금지
- canvas/grid/list 같은 일반 UI pattern은 사용 가능
- DPI 100/125/150% 실제 점검 필요

Workspace는 MainForm 상단 `Workspace` 버튼으로 직접 노출되어 있다.

단축키:

```text
Ctrl+N         New PPTX
Ctrl+Alt+N     Workspace
Ctrl+Alt+O     Open editable document
Ctrl+Alt+D     New DOCX
Ctrl+Alt+X     New XLSX
Ctrl+Alt+H     New HWPX
Ctrl+Alt+T     New ODT
Ctrl+Alt+S     New ODS
Ctrl+Alt+P     New ODP
Ctrl+Shift+E   Edit current PPTX safely
Ctrl+Shift+I   PPTX compatibility report
Ctrl+Shift+P   Export current rendered presentation to PDF
```

## 9. 공통 package/model 구조

### OOXML/OPC

```text
src/OpcPackage.cs
src/OpcPreservation.cs
```

PPTX/DOCX/XLSX가 공통 package helper를 공유한다.

### ODF

```text
src/OdfPackage.cs
src/OdtCore.cs
src/OdsCore.cs
src/OdpCore.cs
```

ODT/ODS/ODP가 common ODF package helper를 공유한다.

### Internal models

```text
PresentationDocument
TextDocument
SpreadsheetDocument
```

UI는 가능한 한 package XML을 직접 수정하지 않고 model을 편집한다.

## 10. PPTX

관련:

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

현재:

- multiple slides
- text / system fonts / style
- shapes
- images
- basic tables
- object drag/resize
- keyboard nudge
- thumbnail preview
- undo/redo
- copy/paste
- safe staged save
- guarded existing-file editing

임의 외부 PPTX가 안전하지 않으면 edit를 거부한다.

## 11. DOCX / HWPX / ODT

공통 model:

```text
TextDocument
 └─ DocumentParagraph[]
     └─ DocumentTextRun[]
```

관련:

```text
src/TextDocumentModel.cs
src/TextDocumentEditSession.cs
src/UnifiedTextDocumentEditorForm.cs
src/DocxReader.cs / DocxWriter.cs
src/HwpxCore.cs
src/OdtCore.cs
```

`TextDocumentEditSession`의 format:

```text
Docx
Hwpx
Odt
```

project-generated/safe package만 제한적으로 edit하며, unsupported content는 deny한다.

## 12. XLSX / ODS

XLSX:

```text
src/SpreadsheetCore.cs
src/SpreadsheetEditorForm.cs
```

ODS:

```text
src/OdsCore.cs
src/OdsSpreadsheetEditorForm.cs
```

공통 `SpreadsheetDocument` model을 사용한다.

현재 셀 종류:

```text
Blank / Text / Number / Boolean / Formula
```

Formula는 저장/읽기 대상이며 자체 계산 engine이 완성됐다고 주장하지 않는다.

## 13. ODP

관련:

```text
src/OdpCore.cs
src/OdpPresentationEditorForm.cs
src/OdpDiagnostics.cs
```

`PresentationDocument` model을 재사용한다.

현재:

- text
- system font family / size / style
- image
- basic shape
- slide add/delete
- drag/resize via `AdvancedPresentationCanvas`
- Save/Save As
- project-generated edit safety

LibreOffice 실제 검증 전 experimental이다.

## 14. HWP 5.x read-only

관련:

```text
src/CompoundFileReader.cs
src/HwpReader.cs
src/HwpReadOnlyViewerForm.cs
src/HwpDiagnostics.cs
```

현재:

- CFB/OLE container
- FileHeader
- FAT / MiniFAT / Directory
- BodyText/SectionN
- raw DEFLATE compressed stream
- HWP record parser
- PARA_TEXT extraction
- preview fallback
- protected document deny
- TextDocument plain-text bridge

아직 full formatting/table/image support를 주장하지 않는다. HWP Writer는 후순위다.

## 15. PDF Export

관련:

```text
src/PdfExport.cs
src/PdfExportDiagnostics.cs
src/MainFormPdfExport.cs
RUN_PDF_SELFTEST.cmd
```

현재:

- minimal PDF 1.4 writer
- JPEG XObject pages
- presentation/text/spreadsheet raster export
- Viewer rendered slide images → PDF
- staged save + backup
- structural self-test

Raster export는 searchable text/vector export가 아니다. PDF Viewer/print 수동 검증 필요.

## 16. 테스트

Windows 빌드 후:

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
RUN_ODT_SELFTEST.cmd
RUN_ODS_SELFTEST.cmd
RUN_ODP_SELFTEST.cmd
RUN_PDF_SELFTEST.cmd
```

Structural self-test와 실제 외부 앱 호환 검증은 별개다.

수동 검증 후보:

```text
PPTX → PowerPoint / LibreOffice
DOCX → Word / LibreOffice
XLSX → Excel / LibreOffice
HWPX → Hancom
HWP → 권리 확인 실제 HWP 5.x sample
ODT/ODS/ODP → LibreOffice
PDF → 일반 PDF viewer / print
```

## 17. 다음 우선순위

1. 최신 `feature/office-foundation` Windows csc.exe compile 오류 0 만들기
2. `RUN_ALL_FORMAT_SELFTESTS.cmd` 전체 통과
3. PPTX Viewer regression
4. PPTX/DOCX/XLSX real Office/LibreOffice validation
5. HWPX Hancom validation
6. ODT/ODS/ODP LibreOffice validation
7. PDF visual/print validation
8. 발견되는 호환성 문제 수정
9. arbitrary external document preservation 확대
10. DOCX table/image, XLSX styles/formats, HWP richer read support
11. 최종 독립 brand 검토 후 public-release license 정리

**기능 수보다 문서 손실 방지, 저작권/폰트 안전, 실제 Windows build/호환성 검증이 우선이다.**
