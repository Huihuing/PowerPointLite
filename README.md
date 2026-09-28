# PowerPointLite

PowerPointLite는 Windows용 경량 문서 도구 프로젝트입니다.

`main`의 현재 안정 기반은 **PPTX/PPTM Viewer**이며, `feature/office-foundation`에서는 기존 Viewer를 유지하면서 프레젠테이션·문서·스프레드시트 Reader/Writer/Editor를 독립 구현하고 있습니다.

> 안정 기준: **1.3 Viewer**  
> Writer/Editor/멀티포맷 기능: **experimental / feature branch**

## 방향

프로젝트는 특정 상용 Office 제품의 코드나 UI 자산을 복제하지 않습니다. 공개된 문서 규격과 프로젝트 자체 내부 모델을 사용해 Reader / Writer / Editor를 구현합니다.

장기 대상:

```text
PPTX / PPTM
DOCX / DOCM
XLSX / XLSM
HWPX
HWP
ODT / ODS / ODP
PDF Export
```

매크로 문서는 향후 일반 콘텐츠를 읽고 보존할 수 있더라도 VBA/ActiveX를 자동 실행하지 않는 방향입니다.

## 현재 상태

| 포맷 | 읽기 | 생성/저장 | 편집 UI | 현재 상태 |
| --- | --- | --- | --- | --- |
| PPTX | 지원 | experimental | experimental | Viewer가 현재 안정 기반 |
| PPTM | Viewer 지원 | 보존 연구 단계 | 제한적 | 매크로 실행 안 함 |
| DOCX | experimental | experimental | experimental | 자체 round-trip 기반 |
| XLSX | experimental | experimental | experimental | 셀/시트/formula 저장 기반 |
| HWPX | experimental | experimental | experimental | 공개 구조 기반, 실제 한컴 검증 필요 |
| HWP 5.x | experimental | 없음 | 읽기 전용 | CFB + BodyText plain-text 기반 |
| ODT | experimental | experimental | experimental | 공통 TextDocument Editor 연결 |
| ODS | experimental | experimental | experimental | 독립 worksheet Editor 연결 |
| ODP | experimental | experimental | experimental | 독립 slide Editor 연결 |
| PDF Export | experimental | Export only | - | 현재 raster-page 방식 |

**experimental은 Microsoft Office, LibreOffice, Hancom 또는 다양한 실제 문서와의 호환성이 아직 충분히 검증되지 않았다는 뜻입니다.**

## PPTX Viewer

현재 안정 기반:

- `.pptx`, `.pptm` 직접 열기
- `.ppt`는 설치된 PowerPoint/LibreOffice를 사용할 수 있을 때 fallback
- Drag & Drop / Recent files
- thumbnails / TOC / Auto TOC
- 실제 오른쪽 viewport 기준 Fit
- zoom / fullscreen
- slide search / go-to / sorter
- hidden slide / speaker notes
- Presenter View
- full slide / notes / 2·4·6 handout printing
- F5 / Shift+F5 slideshow

내부 Open XML Renderer는 text, images, common shapes, tables, basic charts, SmartArt approximation, hyperlinks, gradients, 일부 transition/animation 등을 처리합니다.

렌더링 우선순위:

```text
Microsoft PowerPoint installed
        ↓
PowerPoint Native Renderer

PowerPoint unavailable
        ↓
Internal OpenXML Renderer

Legacy .ppt / selected fallback
        ↓
LibreOffice if installed
```

PowerPoint/LibreOffice 실행 파일이나 DLL을 프로그램 패키지에 포함하지 않습니다.

## Document Workspace

`feature/office-foundation`에는 프로젝트 자체 dark/flat 디자인의 Workspace가 있습니다.

메인 Viewer 상단의 **Workspace** 버튼 또는 `Ctrl+Alt+N`으로 열 수 있습니다.

현재 Workspace:

```text
PPTX
DOCX
XLSX
HWPX
ODT
ODS
ODP
Open existing
Export document to PDF
```

주요 단축키:

```text
Ctrl+N         New PPTX presentation
Ctrl+Alt+N     Workspace
Ctrl+Alt+O     Open editable document
Ctrl+Alt+D     New DOCX
Ctrl+Alt+X     New XLSX
Ctrl+Alt+H     New HWPX
Ctrl+Alt+T     New ODT
Ctrl+Alt+S     New ODS
Ctrl+Alt+P     New ODP
Ctrl+Shift+E   Edit current PPTX when safe
Ctrl+Shift+I   PPTX compatibility report
Ctrl+Shift+P   Export currently rendered presentation to PDF
```

Workspace와 Editor UI는 Microsoft Office/Hancom Ribbon, 공식 아이콘 또는 이미지 자산을 복제하지 않습니다.

### PPTX Editor

현재 experimental Editor 기반:

- multiple slides / real thumbnails
- text boxes
- system font picker
- text formatting/color/alignment
- object selection / drag / resize / keyboard nudge
- basic shapes
- image insertion
- basic tables
- undo/redo
- text/shape/image copy-paste
- staged Save / Save As
- compatibility analyzer / unsupported-content safety guard

외부 PPTX를 현재 내부 모델이 안전하게 round-trip할 수 없다고 판단하면 편집을 거부하고 원본을 덮어쓰지 않습니다.

### DOCX / HWPX / ODT Editor

세 포맷은 공통 `TextDocument` 모델과 Editor를 재사용합니다.

현재 기반:

- paragraphs / runs
- font family / size
- bold / italic / underline
- text color
- paragraph alignment
- create / read / edit / save
- project-generated document safety guard

HWPX는 실제 Hancom 호환 검증이 아직 필요합니다.

### XLSX / ODS Editor

현재 기반:

- multiple worksheets
- add/delete/move/rename sheet
- sparse cells
- text / number / boolean cells
- formula storage
- DataGridView-based independent UI
- cell input/formula bar
- Save / Save As

현재 formula는 **저장·읽기 대상**이며 자체 계산 엔진이 완성된 상태는 아닙니다.

### ODP Editor

ODP는 `PresentationDocument` 모델을 사용합니다.

현재 experimental UI:

- slide add/delete
- text boxes
- system font family / size / style
- object drag/resize
- images
- rectangle / ellipse
- Save / Save As
- project-generated ODP safety guard

ODF custom-shape mapping이 구현되기 전에는 다른 shape kind를 ODP에 조용히 단순화하지 않습니다. 현재 Editor는 Rectangle/Ellipse만 생성하도록 제한하고, Writer가 정확하게 보존할 수 없는 shape가 model에 있으면 저장을 중단합니다.

실제 LibreOffice round-trip 검증 전에는 안정 지원으로 표시하지 않습니다.

### HWP 5.x Read-only

HWP는 Writer보다 Reader를 먼저 구현합니다.

현재 experimental 경로:

```text
Compound File Binary
 ↓
FileHeader
 ↓
FAT / MiniFAT / Directory
 ↓
BodyText/SectionN
 ↓
HWP record parser
 ↓
PARA_TEXT
 ↓
TextDocument
```

현재 구현된 기반에는 CFB container, FileHeader, compressed BodyText stream, record parsing, plain paragraph text extraction, read-only Viewer가 포함됩니다.

암호/배포/DRM/인증서 암호화 문서는 해제를 시도하지 않고 거부합니다. Tables/images/full formatting 지원은 아직 주장하지 않습니다.

## PDF Export

첫 PDF Export 경로는 **페이지를 raster image로 렌더링한 뒤 PDF에 넣는 방식**입니다.

지원 기반:

- PresentationDocument → PDF
- TextDocument → PDF
- SpreadsheetDocument → PDF
- DOCX/HWPX/HWP/ODT/XLSX/ODS/ODP 파일 → PDF
- 현재 Viewer에서 렌더된 PPTX 슬라이드 → PDF (`Ctrl+Shift+P`)

PPTX는 일반 문서 export picker에서 단순 model conversion을 하지 않고, Viewer에서 실제 렌더한 결과를 PDF로 내보내는 경로를 우선합니다.

현재 raster PDF 경로는 사용자가 선택한 TTF/OTF 파일 자체를 PDF 안에 embedding하지 않습니다. 따라서 초기 Export에서 폰트 바이너리 재배포를 피할 수 있습니다.

향후 searchable/vector text PDF를 구현하면서 font subset/full embedding이 필요해질 경우 **반드시 `FontLicenseService`의 라이선스 검사를 거친 뒤** 포함합니다.

상세: [`docs/PDF_EXPORT_FOUNDATION.md`](docs/PDF_EXPORT_FOUNDATION.md)

현재 PDF 출력은 구조 self-test 외에 실제 PDF Viewer/인쇄 결과의 수동 검증이 필요합니다.

## 문서 손실 방지

Writer가 이해하지 못하는 외부 문서를 단순화해 덮어쓰지 않는 것이 기본 정책입니다.

```text
open
 ↓
safety / compatibility analysis
 ↓
현재 model로 손실 없이 표현 가능?
 ├─ YES → experimental edit path
 └─ NO  → 원본 유지 / 편집 거부
```

unknown/unsupported part preservation이 확대될수록 안전 편집 범위를 넓힙니다.

## 저작권 / 상표 / 자산 정책

프로젝트 정책:

- Microsoft Office/Hancom 실행 파일 또는 DLL 재배포 안 함
- Office/Hancom 공식 로고·아이콘·UI 이미지 복제 안 함
- 상용 템플릿/클립아트 무단 포함 안 함
- 인터넷에서 가져온 타인 문서를 테스트 fixture로 커밋하지 않음
- 공개/공식 문서 형식 규격을 바탕으로 독립 코드 작성
- 외부 라이브러리/폰트/자산은 라이선스 확인 후 사용

상세: [`docs/LEGAL_ASSET_POLICY.md`](docs/LEGAL_ASSET_POLICY.md)

Microsoft, Word, Excel and PowerPoint are trademarks of Microsoft Corporation. This project is not affiliated with or endorsed by Microsoft.

Hancom and related product names are trademarks of their respective owners. This project is not affiliated with or endorsed by Hancom.

## 폰트 정책

기본 동작은 Windows에 설치된 font family를 사용하고 문서에는 family 이름만 기록하는 것입니다.

폰트 파일 자체를 자동으로 앱/문서에 포함하지 않습니다.

관련 기반:

```text
src/FontLicensing.cs
src/FontLicenseService.cs
src/FontBundlePolicy.cs
```

- OpenType/TrueType `OS/2.fsType` 검사
- Installable / Editable / Preview&Print / Restricted 구분
- NoSubsetting / BitmapEmbeddingOnly 처리
- 실제 라이선스 원문이 metadata보다 우선
- 명시적인 license review가 없으면 embedding은 deny-by-default
- 앱 번들 권리는 fsType만으로 판단하지 않음

## UI

UI는 자체 dark/flat 디자인을 사용합니다.

- original palette
- proprietary Ribbon cloning 없음
- 공식 Microsoft/Hancom 로고·아이콘 사용 없음
- system-installed `Segoe UI` 사용 가능
- font binary는 번들하지 않음
- presentation/document/spreadsheet에 맞는 별도 canvas/editor

상세: [`docs/UI_DESIGN_GUIDE.md`](docs/UI_DESIGN_GUIDE.md)

## 빌드

요구사항:

- Windows 10/11
- Windows .NET Framework 4.x C# compiler

`.NET 8 SDK`는 필요하지 않습니다.

```bat
BUILD_EXE.cmd
```

빌드 스크립트가 찾는 컴파일러:

```text
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
```

정상 빌드 시 포터블 GUI 실행 파일:

```text
PowerPointLite.exe
```

현재 feature 브랜치는 이 환경에서 Windows `csc.exe` 실제 컴파일을 완료했다고 가정하지 않습니다. **`BUILD SUCCESS` 확인 전에는 main 안정판에 병합하지 않습니다.**

## 자체 테스트

Windows에서 빌드 + 전체 structural test를 한 번에 실행하려면:

```bat
RUN_PREMERGE_CHECKS.cmd
```

개별 실행:

```bat
BUILD_EXE.cmd
RUN_ALL_FORMAT_SELFTESTS.cmd
```

포맷별 테스트:

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

테스트 파일은 프로젝트 코드가 직접 생성합니다. 타인의 실제 문서나 상용 템플릿을 repository fixture로 사용하지 않습니다.

Structural self-test와 실제 앱 호환성은 다릅니다. 가능한 경우 다음 수동 검증이 별도로 필요합니다.

- PPTX → PowerPoint / LibreOffice
- DOCX → Word / LibreOffice
- XLSX → Excel / LibreOffice
- HWPX → Hancom
- HWP → 사용 권한이 있는 실제 HWP 5.x 샘플
- ODT / ODS / ODP → LibreOffice
- PDF → 일반 PDF Viewer와 실제 인쇄

## 주요 구조

```text
src/
├─ Viewer / rendering partials
├─ PresentationModel.cs
├─ PptxWriter.cs
├─ AdvancedPresentationEditorForm.cs
├─ OdpCore.cs / OdpPresentationEditorForm.cs
├─ TextDocumentModel.cs
├─ DocxReader.cs / DocxWriter.cs
├─ HwpxCore.cs
├─ OdtCore.cs
├─ UnifiedTextDocumentEditorForm.cs
├─ SpreadsheetCore.cs / SpreadsheetEditorForm.cs
├─ OdsCore.cs / OdsSpreadsheetEditorForm.cs
├─ HwpReader.cs / CompoundFileReader.cs
├─ HwpReadOnlyViewerForm.cs
├─ OdfPackage.cs
├─ PdfExport.cs / DocumentPdfExport.cs
├─ OpcPackage.cs / OpcPreservation.cs
├─ FontLicensing.cs / FontLicenseService.cs
├─ OfficeWorkspaceDialog.cs
└─ UiTheme.cs

docs/
├─ ARCHITECTURE_ROADMAP.md
├─ EDITOR_FOUNDATION.md
├─ HWPX_FOUNDATION.md
├─ HWP_FOUNDATION.md
├─ ODF_FOUNDATION.md
├─ PDF_EXPORT_FOUNDATION.md
├─ LEGAL_ASSET_POLICY.md
└─ UI_DESIGN_GUIDE.md
```

상세 진행 체크리스트: [`docs/ARCHITECTURE_ROADMAP.md`](docs/ARCHITECTURE_ROADMAP.md)

새 AI/ChatGPT/Codex 세션에서 개발을 이어갈 경우 [`AGENTS.md`](AGENTS.md)를 먼저 읽습니다.

## 라이선스

프로젝트 자체 코드는 루트 [`LICENSE`](LICENSE)의 **MIT License**를 사용합니다.

서드파티 폰트, 라이브러리, 아이콘, 이미지, 문서 샘플, 상표 등 외부 권리는 MIT License로 다시 허가되는 것이 아니며 각각의 원 라이선스를 따릅니다. 현재 고지 정책은 [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md)와 `licenses/`에 분리해 관리합니다.
