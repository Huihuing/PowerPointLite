# PowerPointLite

PowerPointLite는 Windows에서 프레젠테이션·문서·스프레드시트를 가볍게 읽고 다루기 위한 독립 문서 도구 프로젝트입니다.

현재 `main`은 **PPTX/PPTM Viewer**가 안정 기준이며, `feature/office-foundation`에서는 기존 Viewer를 유지하면서 Writer/Editor와 멀티포맷 기반을 실험적으로 확장하고 있습니다.

> Stable baseline: **1.3 Viewer**  
> Multi-format Writer/Editor: **experimental**

## Current Status

- 기본 UI 언어: **한국어**, 상단에서 한국어/English 전환 가능
- PPTX/PPTM Viewer 기반 동작
- Internal OpenXML renderer 표현력 확장 중
- PPTX Writer / Editor 기반 개발 중
- DOCX / XLSX / HWPX / ODT / ODS / ODP Reader·Writer·Editor 기반 개발 중
- HWP 5.x는 read-only 기반
- PDF는 현재 raster-page Export 기반
- Windows .NET Framework 전체 feature build와 실제 Office/LibreOffice/Hancom 호환 검증은 아직 필요

## Supported Formats

| Format | Read | Create / Save | Edit UI | Notes |
| --- | --- | --- | --- | --- |
| PPTX | Yes | Experimental | Experimental | Viewer가 현재 안정 기반 |
| PPTM | Viewer | Preservation research | Limited | VBA 실행 안 함 |
| DOCX | Experimental | Experimental | Experimental | 공통 TextDocument 기반 |
| XLSX | Experimental | Experimental | Experimental | 셀/시트/formula 저장 기반 |
| HWPX | Experimental | Experimental | Experimental | 실제 Hancom 검증 필요 |
| HWP 5.x | Experimental | No | Read-only | plain-text 중심 reader 기반 |
| ODT | Experimental | Experimental | Experimental | ODF text 기반 |
| ODS | Experimental | Experimental | Experimental | ODF spreadsheet 기반 |
| ODP | Experimental | Experimental | Experimental | ODF presentation 기반 |
| PDF | - | Export | - | 현재 raster-page 방식 |

`Experimental`은 structural self-test가 존재하더라도 Microsoft Office, LibreOffice, Hancom 또는 실제 다양한 문서와의 호환성이 충분히 검증되지 않았다는 뜻입니다.

## PPTX Viewer

현재 Viewer 주요 기능:

- `.pptx`, `.pptm` 열기
- 설치된 PowerPoint가 있으면 Native renderer 우선
- 설치된 LibreOffice가 있으면 Impress renderer 사용 가능
- 외부 renderer가 없으면 자체 Internal OpenXML renderer
- thumbnails / TOC / Auto TOC
- Fit / Zoom / Fullscreen
- slide search / go-to / sorter
- hidden slide / speaker notes
- Presenter View
- F5 / Shift+F5 slideshow
- full slide / notes / handout printing
- rich text / table / theme / master-layout inheritance
- image crop/rotate/flip/alpha 및 일부 color effect
- SVG basic primitive + path 렌더링
- 추가 shape/custom geometry/group transform 처리
- column/bar/line/pie/doughnut/area/scatter/bubble/radar chart 근사 렌더링
- SmartArt data-model hierarchy 근사 렌더링
- `mc:AlternateContent` 호환 fallback
- slide transition 및 internal animation timeline
  - entrance / exit / emphasis / motion 분류
  - on-click / with-previous / after-previous
  - delay / duration / repeat / auto-reverse
  - slide `advTm` / `advClick` / loop / useTimings

Internal renderer는 Microsoft PowerPoint를 복제한 엔진이 아닙니다. SmartArt 고유 layout, 모든 chart style/3D 조합, 고급 SVG filter/arc, 3D effect, 정확한 motion path/easing, 모든 animation timing edge case, Morph object matching 등은 근사 또는 fallback입니다.

상세: [`docs/PPTX_FIDELITY.md`](docs/PPTX_FIDELITY.md)

PowerPoint/LibreOffice 실행 파일이나 DLL을 프로젝트 패키지에 포함하지 않습니다.

## PPTX Editor

Editor의 interactive canvas는 개체 선택/이동/크기 조절을 빠르게 하기 위한 편집 화면이고, 최종 시각 표현의 기준은 Viewer renderer입니다.

Editor의 `미리보기 / Preview`는 현재 편집 모델을 임시 PPTX로 만든 뒤 Viewer와 동일한 Internal renderer 경로로 다시 렌더링합니다.

`PresentationTextBox`는 기존 단순 텍스트 속성과 함께 선택적으로 paragraph/run rich-text 정보를 보존할 수 있습니다.

현재 run-level 보존 범위:

- font family / size
- bold / italic / underline
- color
- baseline
- paragraph alignment / level / bullet / before-after spacing

현재 일반 TextBox 편집 UI에서 문자열 자체를 바꾸면 이전 run 경계를 억지로 새 글자에 적용하지 않고 해당 TextBox를 box-level format으로 평문화합니다. 선택 영역 단위 rich-text 편집 UI는 후속 작업입니다.

외부 PPTX는 `PptxEditableReader`가 안전한 round-trip이 가능하다고 판정한 경우에만 Editor 저장 경로를 허용합니다.

상세: [`docs/EDITOR_FOUNDATION.md`](docs/EDITOR_FOUNDATION.md)

## Document Workspace

`feature/office-foundation`에서는 프로젝트 자체 dark/flat UI의 Workspace를 제공합니다.

현재 진입 가능한 작업:

```text
New PPTX
New DOCX
New XLSX
New HWPX
New ODT
New ODS
New ODP
Open existing supported document
Export supported document to PDF
Convert supported format
```

외부 문서는 현재 내부 모델이 안전하게 표현할 수 있다고 판단될 때만 편집/변환합니다. 지원하지 않는 내용을 조용히 제거한 뒤 원본을 덮어쓰는 동작은 기본적으로 거부합니다.

## Architecture

공통 내부 모델:

```text
PresentationDocument
TextDocument
SpreadsheetDocument
```

기본 흐름:

```text
Reader
  ↓
Internal Model
  ├─ Renderer / Viewer
  └─ Editor
       ↓
     Writer
```

패키지 공통 계층:

```text
OOXML / OPC → OpcPackage / OpcPreservation
ODF          → OdfPackage
```

상세 구조와 단계별 목표는 [`docs/ARCHITECTURE_ROADMAP.md`](docs/ARCHITECTURE_ROADMAP.md)를 참고하세요.

## UI

UI는 프로젝트 자체 dark/flat 디자인을 사용합니다.

- 기본 언어 한국어
- 한국어 / English 전환
- Microsoft Office / Hancom Ribbon 시각 복제 없음
- 공식 제품 로고/아이콘 사용 없음
- presentation / document / spreadsheet 성격에 맞는 별도 editor UI
- system-installed fonts 사용
- font binary를 앱에 자동 번들하지 않음

상세: [`docs/UI_DESIGN_GUIDE.md`](docs/UI_DESIGN_GUIDE.md)

## Build

요구사항:

- Windows 10/11
- Windows .NET Framework 4.x C# compiler

`.NET 8 SDK`는 필요하지 않습니다.

```bat
RUN_SOURCE_AUDIT.cmd
BUILD_EXE.cmd
```

`RUN_SOURCE_AUDIT.cmd`는 실제 컴파일 전에 알려진 `Timer` 모호성 및 선택된 최신 C#/.NET 전용 패턴을 보수적으로 검사합니다. 최종 판정은 Windows `csc.exe` 실제 빌드입니다.

빌드 스크립트는 다음 Framework compiler를 찾습니다.

```text
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
```

정상 빌드 시:

```text
PowerPointLite.exe
```

현재 experimental feature branch는 실제 Windows `BUILD SUCCESS` 확인 전까지 안정판으로 간주하지 않습니다.

## Test

Windows에서 pre-merge 자동 검증:

```bat
RUN_PREMERGE_CHECKS.cmd
```

전체 structural format/policy self-test:

```bat
RUN_ALL_FORMAT_SELFTESTS.cmd
```

개별 테스트도 제공합니다.

```bat
RUN_WRITER_SELFTEST.cmd
RUN_ANIMATION_SELFTEST.cmd
RUN_DOCX_SELFTEST.cmd
RUN_XLSX_SELFTEST.cmd
RUN_HWPX_SELFTEST.cmd
RUN_HWP_PARSER_SELFTEST.cmd
RUN_ODT_SELFTEST.cmd
RUN_ODS_SELFTEST.cmd
RUN_ODP_SELFTEST.cmd
RUN_PDF_SELFTEST.cmd
RUN_CONVERSION_SELFTEST.cmd
RUN_FONT_LICENSE_SELFTEST.cmd
```

Writer self-test는 일반 Writer round-trip과 rich-text paragraph/run round-trip을 함께 검사합니다.

Animation self-test는 프로젝트가 생성한 PPTX에 합성 `p:timing`을 넣어 click entrance, with-previous emphasis, after-previous exit, delayed motion step 및 단계별 render state를 검사합니다.

테스트 문서는 프로젝트 코드가 직접 생성합니다. 인터넷에서 가져온 타인 문서나 상용 템플릿을 repository fixture로 사용하지 않습니다.

Structural self-test와 실제 애플리케이션 호환성은 별개입니다. 가능한 경우 다음 수동 확인이 필요합니다.

- PPTX → PowerPoint / LibreOffice, 특히 rich text / chart / SmartArt / animation timing
- DOCX → Word / LibreOffice
- XLSX → Excel / LibreOffice
- HWPX → Hancom
- HWP → 사용 권한이 있는 실제 HWP 5.x 샘플
- ODT / ODS / ODP → LibreOffice
- PDF → 일반 PDF Viewer / 실제 인쇄

## Copyright / Trademark / Fonts

프로젝트는 공개된 문서 형식 규격과 자체 구현을 중심으로 개발합니다.

하지 않는 것:

- Microsoft Office / Hancom 실행 파일 또는 DLL 재배포
- 공식 Office / Hancom 로고·아이콘·UI 이미지 복제
- 상용 템플릿/클립아트 무단 포함
- 라이선스가 확인되지 않은 상용 폰트 파일 번들
- 타인의 실제 문서를 테스트 fixture로 커밋

기본 폰트 정책:

- Windows에 설치된 font family 사용
- 문서에는 기본적으로 font family 이름만 기록
- font file 자동 번들/embedding 안 함
- OpenType/TrueType `OS/2.fsType`은 보조 정보로만 사용
- 실제 라이선스 원문이 우선
- 명시적 라이선스 검토가 없으면 embedding deny-by-default

상세:

- [`docs/LEGAL_ASSET_POLICY.md`](docs/LEGAL_ASSET_POLICY.md)
- [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md)
- [`licenses/README.md`](licenses/README.md)

Microsoft, Word, Excel and PowerPoint are trademarks of Microsoft Corporation. This project is not affiliated with or endorsed by Microsoft.

Hancom and related product names are trademarks of their respective owners. This project is not affiliated with or endorsed by Hancom.

## Documentation

Durable project knowledge is kept under `docs/`.

- [`docs/ARCHITECTURE_ROADMAP.md`](docs/ARCHITECTURE_ROADMAP.md) — architecture and format roadmap
- [`docs/EDITOR_FOUNDATION.md`](docs/EDITOR_FOUNDATION.md) — editor model and safety foundation
- [`docs/PPTX_FIDELITY.md`](docs/PPTX_FIDELITY.md) — Internal PPTX renderer / animation fidelity scope
- [`docs/HWPX_FOUNDATION.md`](docs/HWPX_FOUNDATION.md) — HWPX foundation
- [`docs/HWP_FOUNDATION.md`](docs/HWP_FOUNDATION.md) — HWP 5.x read-only foundation
- [`docs/ODF_FOUNDATION.md`](docs/ODF_FOUNDATION.md) — ODT/ODS/ODP foundation
- [`docs/PDF_EXPORT_FOUNDATION.md`](docs/PDF_EXPORT_FOUNDATION.md) — PDF export foundation
- [`docs/LEGAL_ASSET_POLICY.md`](docs/LEGAL_ASSET_POLICY.md) — copyright/trademark/font asset policy
- [`docs/UI_DESIGN_GUIDE.md`](docs/UI_DESIGN_GUIDE.md) — UI design rules

## License

Project code is licensed under the MIT License unless a file or third-party notice states otherwise.

Third-party fonts, libraries, icons, images, templates and other assets remain subject to their own licenses and notices.
