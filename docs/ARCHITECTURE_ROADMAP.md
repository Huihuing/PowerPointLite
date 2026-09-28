# Multi-format Architecture Roadmap

현재 안정 기능은 PPTX Viewer이며, 이 문서는 기존 Viewer를 깨뜨리지 않고 향후 문서·스프레드시트·프레젠테이션 편집기로 확장하기 위한 구조 기준이다.

## 원칙

1. Reader / Writer를 분리한다.
2. UI가 파일 포맷 XML을 직접 다루지 않는다.
3. 공통 내부 모델을 통해 Renderer / Editor / Writer를 연결한다.
4. OOXML 공통 계층을 PPTX Writer보다 먼저 또는 동시에 분리한다.
5. 기존 PPTX Viewer 회귀 테스트를 통과하지 못한 대규모 refactor는 main에 병합하지 않는다.
6. 매크로 문서는 열 수 있어도 VBA를 실행하지 않는다.
7. 상표/폰트/템플릿/외부 자산은 라이선스가 명확하지 않으면 포함하지 않는다.

## 목표 구조

```text
src/
├─ Core/
│  ├─ Packaging/
│  ├─ Xml/
│  ├─ Relationships/
│  ├─ Drawing/
│  ├─ Images/
│  ├─ Fonts/
│  └─ FontLicensing/
│
├─ Models/
│  ├─ Presentation/
│  ├─ Document/
│  └─ Spreadsheet/
│
├─ Formats/
│  ├─ Presentation/PPTX/
│  ├─ Document/DOCX/
│  ├─ Document/HWPX/
│  ├─ Document/HWP/
│  ├─ Spreadsheet/XLSX/
│  └─ ODF/
│
├─ Renderer/
├─ Editor/
├─ Printing/
├─ Export/
└─ UI/
```

현재 `BUILD_EXE.cmd`가 `src\*.cs`만 컴파일하므로 실제 디렉터리 분리는 빌드 스크립트를 재귀 소스 목록 방식으로 바꾸는 단계와 함께 진행한다. 그 전에는 namespace와 파일 단위로 논리 분리를 먼저 한다.

## Reader / Writer 흐름

```text
PPTX
 ↓
PptxReader
 ↓
PresentationDocument
 ↓              ↓
Renderer      Editor
                 ↓
             PptxWriter
                 ↓
                PPTX
```

동일 패턴을 DOCX/XLSX/HWPX에 재사용한다.

## Phase 0 — PPTX Viewer 안정화

- [x] PPTX Reader / Renderer
- [x] Presentation mode
- [x] Presenter View / Notes / Printing
- [ ] Windows 실제 빌드 오류 0 유지
- [ ] 다양한 자체 테스트 PPTX 회귀 테스트
- [ ] 진행 중 presentation-tools 기능 안정화

## Phase 1 — 법적/자산/폰트 안전 기반

- [x] 자산·상표·폰트 정책 문서
- [x] THIRD_PARTY_NOTICES 유지 체계
- [x] system font catalog 초기 구현
- [x] OpenType `OS/2.fsType` reader
- [x] FontLicenseInfo / FontLicenseService 초기 구현
- [ ] bundle font allow-list 구조
- [ ] 독립적인 프로젝트 브랜딩 결정
- [ ] 실제 번들 폰트 도입 시 라이선스 원문 검토

## Phase 2 — PPTX Writer MVP

현재 `src/PptxWriter.cs`는 UI에 노출하지 않은 실험적 Writer다.

- [x] 새 presentation package 생성 초기 구현
- [x] `[Content_Types].xml`
- [x] package / presentation relationships
- [x] 자체 neutral theme/master/layout 최소 세트
- [x] 첫 슬라이드와 text box 생성
- [x] 기본 font/style XML 생성
- [ ] Presentation internal model 연결
- [ ] slide add/delete/reorder
- [ ] 기존 text box 편집
- [ ] image insert
- [ ] basic shapes
- [ ] basic table
- [ ] Save / Save As UI
- [ ] Writer 생성물을 Windows PowerPoint/PowerPointLite에서 실제 검증
- [ ] 생성 → 재열기 → 수정 → 재저장 round-trip

Writer의 기본 패키지 구조는 독립 테스트에서 일반 PPTX parser가 1-slide presentation으로 읽을 수 있는지 확인했지만, Windows `csc.exe` 실빌드와 Microsoft PowerPoint 실제 호환성 검증 전에는 안정 기능으로 간주하지 않는다.

## Phase 3 — 공통 OOXML 계층

`src/OpcPackage.cs`에 첫 공통 계층을 추가했다.

- [x] OPC package XML part helper 초기 구현
- [x] content types reader/writer helper
- [x] relationships reader/writer helper
- [x] URI/part target resolver 초기 구현
- [ ] existing PPTX Reader가 공통 OPC 계층 사용하도록 점진적 전환
- [ ] DrawingML common model
- [ ] theme/color/font common model
- [ ] media/image store
- [ ] unknown part preservation abstraction

## Phase 4 — PPTX Editor UI

- [ ] File → New/Open/Save/Save As
- [ ] edit mode / viewer mode 분리
- [ ] object selection
- [ ] text edit
- [ ] slide add/delete/reorder
- [ ] image insert
- [ ] basic shapes
- [ ] font picker
- [ ] undo/redo
- [ ] copy/paste
- [ ] dirty state / save confirmation

Editor UI는 `docs/UI_DESIGN_GUIDE.md`에 따라 독립 디자인을 사용하고 Office UI 자산을 복제하지 않는다.

## Phase 5 — DOCX

- [ ] DocxReader
- [ ] paragraph/run/styles
- [ ] tables/images
- [ ] headers/footers
- [ ] DocxWriter
- [ ] new/edit/save/print
- [ ] `.docm` macro parts preserve-only

## Phase 6 — XLSX

- [ ] XlsxReader
- [ ] workbook/sheets
- [ ] shared strings
- [ ] cells/styles/number formats
- [ ] formulas
- [ ] merged cells
- [ ] XlsxWriter
- [ ] new/edit/save/print
- [ ] `.xlsm` macro parts preserve-only

## Phase 7 — HWPX / HWP

HWPX를 먼저 구현한다.

- [ ] 최신 공식 규격 재확인
- [ ] HwpxReader
- [ ] HwpxWriter
- [ ] text/table/image/page layout
- [ ] HWP 공식 명세 버전 기록
- [ ] HwpReader
- [ ] HwpWriter는 Reader 안정화 뒤 진행

## Phase 8 — ODF / PDF

- [ ] ODT / ODS / ODP
- [ ] shared ODF package layer
- [ ] PDF export
- [ ] PDF font embedding 시 FontLicenseService 강제 사용

## 테스트 자산

테스트 파일은 자체 generator 또는 프로젝트 자체 작성 파일만 사용한다.

```text
tests/
├─ TEST_BASIC_PPTX.pptx
├─ TEST_BASIC_DOCX.docx
├─ TEST_BASIC_XLSX.xlsx
└─ TEST_BASIC_HWPX.hwpx
```

새 포맷은 최소한 다음 테스트를 갖는다.

1. Create
2. Save
3. Re-open
4. Modify
5. Save As
6. Round-trip preservation
7. Unsupported part preservation where practical

## Merge 기준

새 포맷/Editor 작업보다 기존 PPTX 기능 회귀 방지가 우선이다.

- Windows `BUILD SUCCESS`
- Viewer 실행 확인
- 기존 PPTX 샘플 열기
- Fit / TOC / slideshow / notes / print 확인
- 새 기능 자체 테스트
- 외부 자산/폰트 라이선스 확인

이 조건을 통과한 뒤 main에 병합한다.
