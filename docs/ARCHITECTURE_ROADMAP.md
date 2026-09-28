# Multi-format Architecture Roadmap

현재 안정 기준은 PPTX Viewer이며, `feature/office-foundation`에서 기존 Viewer를 깨뜨리지 않고 문서·스프레드시트·프레젠테이션 Editor로 확장한다.

## 개발 원칙

1. Reader / Writer / Editor를 분리한다.
2. UI에서 ZIP/XML을 직접 수정하지 않는다.
3. 공통 내부 모델을 사용한다.
4. 파일을 현재 모델보다 좁게 읽은 뒤 덮어써서 내용을 유실시키지 않는다.
5. 임의 외부 문서는 unknown-part preservation이 준비되기 전까지 deny-by-default다.
6. VBA/ActiveX/매크로를 실행하지 않는다.
7. Microsoft/Hancom 실행파일·DLL·로고·아이콘·템플릿을 번들하지 않는다.
8. 폰트 embedding은 실제 라이선스 확인이 없으면 허용하지 않는다.
9. 테스트 문서는 코드가 직접 생성한다.
10. Windows 실제 빌드/실행 검증 전에는 `main`으로 병합하지 않는다.

---

# Phase 0 — 기존 PPTX Viewer 유지

- [x] PPTX/PPTM Reader
- [x] Internal OpenXML Renderer
- [x] PowerPoint Native 우선 경로
- [x] LibreOffice fallback
- [x] thumbnails / TOC / Fit / zoom
- [x] slideshow / notes / hidden slide
- [x] Presenter View
- [x] printing layouts
- [x] source split / portable build
- [ ] Windows `BUILD_EXE.cmd` 최신 feature 전체 컴파일 검증
- [ ] 기존 PPTX 회귀 테스트
- [ ] `feature/presentation-tools` 안정화/병합 판단

중요 회귀 방지:

```text
Fit = SplitContainer.Panel2 실제 viewport 기준
GUI Timer = System.Windows.Forms.Timer 명시
.NET Framework csc.exe 호환 문법/API 유지
```

---

# Phase 1 — 저작권·폰트·자산 안전 기반

- [x] `docs/LEGAL_ASSET_POLICY.md`
- [x] `THIRD_PARTY_NOTICES.md`
- [x] `licenses/` 구조
- [x] 시스템 설치 font family catalog
- [x] OpenType/TrueType `OS/2.fsType` parser
- [x] Installable / Editable / Preview&Print / Restricted 구분
- [x] NoSubsetting / BitmapEmbeddingOnly 처리
- [x] `FontLicenseService`
- [x] app font bundle deny-by-default 정책
- [x] 실제 라이선스 원문이 fsType보다 우선
- [ ] 최종 독립 프로젝트 브랜드 결정
- [ ] 실제 번들 폰트가 필요해질 경우 개별 라이선스 원문 검토

현재 기본 정책은 **폰트 파일을 번들하지 않고 family 이름만 문서에 기록**하는 것이다.

---

# Phase 2 — 공통 OOXML / OPC 기반

- [x] ZIP/XML package helper
- [x] `[Content_Types].xml` helper
- [x] `.rels` helper
- [x] relationship target resolver
- [x] PPTX Writer에서 공통 OPC 사용
- [x] DOCX Writer/Reader 기반
- [x] XLSX Writer/Reader 기반
- [x] `OpcPreservation` 초기 기반
- [x] compatibility analyzer 기반
- [ ] 기존 PPTX Renderer의 중복 OPC 코드를 점진적으로 공통 계층으로 이동
- [ ] DrawingML 공통 모델 확대
- [ ] shared theme/color/font model
- [ ] unknown-part preservation 완성

---

# Phase 3 — PPTX Writer / Editor

## Writer

- [x] `PresentationDocument`
- [x] multi-slide writer
- [x] slide add/delete/reorder
- [x] text boxes
- [x] font/style/alignment/color
- [x] basic shapes
- [x] PNG/JPEG/GIF/BMP image parts
- [x] image relationships
- [x] basic DrawingML tables
- [x] safe staged save + backup
- [x] Reader → model → Writer round-trip 기반
- [x] project-generated PPTX safety guard
- [x] compatibility report
- [x] Writer self-test
- [ ] arbitrary external PPTX full preservation
- [ ] unsupported extension parts preservation 확대
- [ ] PowerPoint/LibreOffice 실제 호환 검증

## Editor

현재 주 Editor는 `AdvancedPresentationEditorForm`이다.

- [x] independent dark/flat UI
- [x] slide list
- [x] 실제 slide thumbnail preview
- [x] text object add/delete/edit
- [x] object selection
- [x] drag/resize
- [x] keyboard nudge
- [x] system font picker
- [x] text style/color/alignment
- [x] image insert
- [x] basic shapes
- [x] shape fill/line color
- [x] basic table editor
- [x] table preview overlay
- [x] snapshot undo/redo
- [x] text/shape/image copy/paste
- [x] guarded existing-PPTX edit mode
- [ ] table drag/resize를 일반 object selection과 통합
- [ ] object z-order
- [ ] multi-selection
- [ ] system clipboard structured-object interoperability
- [ ] arbitrary external PPTX safe editing

---

# Phase 4 — DOCX

공통 모델:

```text
TextDocument
 └─ DocumentParagraph[]
     └─ DocumentTextRun[]
```

현재:

- [x] `TextDocumentModel`
- [x] paragraph/run model
- [x] font family / size
- [x] bold / italic / underline
- [x] text color
- [x] paragraph alignment
- [x] DOCX Writer
- [x] DOCX Reader
- [x] project-generated DOCX safety guard
- [x] DOCX structural round-trip diagnostics
- [x] legacy DOCX editor MVP
- [x] DOCX/HWPX 공통 edit session
- [x] `UnifiedTextDocumentEditorForm`
- [ ] tables
- [ ] images
- [ ] headers/footers
- [ ] page layout
- [ ] numbered/bulleted lists
- [ ] styles hierarchy
- [ ] comments/footnotes/endnotes
- [ ] `.docm` unknown/VBA parts preserve-only
- [ ] arbitrary DOCX preservation 확대
- [ ] Word/LibreOffice 실제 호환 검증

---

# Phase 5 — XLSX

현재 `SpreadsheetCore.cs` 기반:

- [x] `SpreadsheetDocument`
- [x] multiple worksheets
- [x] add/delete/move/rename sheet
- [x] sparse cell model
- [x] text cells
- [x] number cells
- [x] boolean cells
- [x] formula storage
- [x] XLSX Writer
- [x] XLSX Reader
- [x] project-generated XLSX safety guard
- [x] create/read/edit/write diagnostics
- [x] `SpreadsheetEditorForm`
- [x] worksheet tabs
- [x] DataGridView cell editing
- [x] formula/input bar
- [x] multi-cell delete
- [x] Save / Save As
- [ ] formula calculation engine
- [ ] cell styles / fonts / colors
- [ ] number/date formats
- [ ] row height / column width
- [ ] merged cells
- [ ] copy/paste range
- [ ] CSV import/export
- [ ] charts
- [ ] comments/data validation
- [ ] `.xlsm` VBA parts preserve-only
- [ ] arbitrary XLSX preservation 확대
- [ ] Excel/LibreOffice 실제 호환 검증

Formula는 현재 **저장/읽기** 대상으로 취급하며 자체 계산 결과를 신뢰성 있게 산출한다고 주장하지 않는다.

---

# Phase 6 — HWPX

HWPX는 HWP보다 먼저 진행한다.

공개 HWPX/OWPML 구조에 기반한 독립 구현이며 한컴 바이너리나 DLL을 사용하지 않는다.

현재:

- [x] `application/hwp+zip` package detection
- [x] `version.xml`
- [x] `META-INF/container.xml`
- [x] `META-INF/manifest.xml`
- [x] `Contents/content.hpf`
- [x] `Contents/header.xml`
- [x] `Contents/section0.xml`
- [x] preview text
- [x] `HwpxReader` → `TextDocument`
- [x] `HwpxWriter` ← `TextDocument`
- [x] paragraph/run text
- [x] basic font/style reference
- [x] basic paragraph alignment
- [x] line break / tab
- [x] deny-by-default `HwpxEditSafety`
- [x] create/read/edit/save/read structural self-test
- [x] DOCX/HWPX shared document edit session
- [x] HWPX entry through unified document editor
- [ ] official schema/validator validation
- [ ] actual Hancom open/save verification
- [ ] multiple sections
- [ ] tables
- [ ] images/BinData
- [ ] lists/styles compatibility 확대
- [ ] page/section settings compatibility 확대
- [ ] unknown-part preservation
- [ ] arbitrary external HWPX safe edit mode

상세: `docs/HWPX_FOUNDATION.md`

**현재 HWPX는 experimental이다. 구조 self-test 성공만으로 한컴 호환 완료라고 표시하지 않는다.**

---

# Phase 7 — HWP Binary

HWP는 공개 HWP 5.x 파일 형식 명세 기반으로 Reader 우선 구현한다.

- [ ] OLE/Compound File container
- [ ] FileHeader
- [ ] DocInfo
- [ ] BodyText/Section stream
- [ ] record parser
- [ ] compressed stream support
- [ ] text extraction
- [ ] basic paragraph/style mapping
- [ ] table/image read
- [ ] read-only viewer/model bridge
- [ ] Writer feasibility 재검토

HWP Writer는 Reader 안정화 및 명세 범위 확인 전까지 후순위다.

---

# Phase 8 — ODF / PDF

- [ ] ODT Reader/Writer
- [ ] ODS Reader/Writer
- [ ] ODP Reader/Writer
- [ ] common ODF package layer
- [ ] PDF page-render model
- [ ] PDF text/image/vector export
- [ ] font subset/embed
- [ ] PDF font embedding에 `FontLicenseService` 강제 적용

---

# 공통 Workspace UI

- [x] original dark/flat `ApplicationTheme`
- [x] `OfficeWorkspaceDialog`
- [x] PPTX new/edit entry
- [x] DOCX new/edit entry
- [x] XLSX new/edit entry
- [x] HWPX new/edit entry
- [x] editable-file open chooser
- [ ] toolbar에서 Workspace 버튼 직접 노출
- [ ] 최근 multi-format documents
- [ ] 파일 타입별 독립 아이콘 제작
- [ ] accessibility/tab order 확대
- [ ] DPI 100/125/150% 실제 점검

현재 단축키:

```text
Ctrl+N        New Presentation
Ctrl+Alt+N    New/Open Workspace
Ctrl+Alt+O    Open editable document
Ctrl+Alt+D    New DOCX
Ctrl+Alt+X    New XLSX
Ctrl+Alt+H    New HWPX
Ctrl+Shift+E  Edit current PPTX safely
Ctrl+Shift+I  PPTX compatibility report
```

---

# 테스트 / Merge Gate

Windows에서 최소:

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
```

그 다음 수동 확인:

1. Viewer 실행
2. 기존 자체 PPTX 샘플
3. TOC on/off Fit
4. wheel / Ctrl+wheel
5. F5 / Shift+F5
6. Presenter / Notes / Print
7. PPTX Editor create/save/reopen
8. DOCX Editor create/save/reopen
9. XLSX Editor create/save/reopen
10. HWPX 구조 test
11. 가능한 경우 Microsoft Office / LibreOffice / Hancom 실제 열기 확인

외부 애플리케이션 호환 확인 전에는 해당 포맷을 안정 지원으로 표시하지 않는다.

---

# 최종 목표

```text
PPTX  read / create / edit / save / present / print
DOCX  read / create / edit / save / print
XLSX  read / create / edit / save / print
HWPX  read / create / edit / save
HWP   read first, writer later if safe
ODT / ODS / ODP
PDF export
```

항상 **기존 문서를 잃지 않는 것, 저작권/상표/폰트 라이선스를 침해하지 않는 것, 실제 빌드 검증을 거치는 것**을 기능 수보다 우선한다.
