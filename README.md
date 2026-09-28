# PowerPointLite

PowerPointLite는 현재 Windows용 `.pptx` / `.pptm` 프레젠테이션 Viewer로 개발 중이며, 장기적으로는 **PPTX/DOCX/XLSX/HWPX 등 문서·스프레드시트·프레젠테이션을 읽고 작성/편집/저장할 수 있는 경량 문서 프로그램**으로 확장하는 것을 검토하고 있습니다.

현재 안정 기능은 PPTX Viewer이며, 멀티포맷 확장은 기존 Viewer를 깨뜨리지 않는 단계적 방식으로 진행합니다.

> 현재 안정 기준: **1.3**

## 현재 목표

- Microsoft PowerPoint가 없어도 일반적인 PPTX를 열 수 있는 독립 Viewer
- PowerPoint가 설치되어 있으면 네이티브 렌더러/슬라이드쇼를 선택적으로 활용
- 별도 .NET 8 SDK 없이 Windows .NET Framework `csc.exe`로 빌드 가능한 구조 유지
- 향후 Reader / Writer / Internal Model을 분리해 PPTX 편집 기능부터 단계적으로 확장
- 상용 프로그램의 코드·DLL·아이콘·템플릿을 복사하지 않는 독립 구현
- 폰트 embedding/재배포 권한을 구조적으로 검사하는 안전한 문서 엔진

## 현재 지원 기능

### PPTX Viewer

- `.pptx`, `.pptm` 직접 열기
- `.ppt`는 PowerPoint 또는 LibreOffice가 있을 때 fallback
- Drag & Drop / Recent files
- 슬라이드 썸네일 / TOC / Auto TOC
- 실제 오른쪽 viewport 기준 Fit
- 확대/축소 / F11 fullscreen
- 슬라이드 검색 / 번호 이동 / sorter
- 숨김 슬라이드
- Speaker Notes
- Presenter View
- 인쇄: full slide / notes / 2·4·6 handout

### 발표 입력

- 마우스 휠: 이전/다음
- `Ctrl + 휠`: zoom
- `← / →`, `PageUp / PageDown`, `Home / End`
- `F5`: 처음부터 슬라이드쇼
- `Shift + F5`: 현재 슬라이드부터
- `N / P`
- 숫자 + `Enter`
- `B`: black screen
- `W`: white screen
- `H`: hidden slide

### 자체 Open XML 렌더러

현재 내부 렌더러는 다음 영역을 처리합니다.

- slide size / background
- theme color/font
- master/layout/placeholder
- text / paragraph / basic formatting
- bullets / numbering
- images / crop / alpha / rotation / flip
- basic SVG subset / GDI+ EMF/WMF fallback
- shapes / custom geometry 일부
- line/connectors/arrowheads
- groups / tables
- basic charts
- SmartArt static approximation
- hyperlinks / slide actions
- media extraction fallback
- gradients / pattern fill / basic shadow
- 일부 transition / entrance animation fallback

## 렌더링 엔진

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

Microsoft 또는 LibreOffice 바이너리를 프로그램에 포함하지 않습니다.

## 향후 확장 방향

우선순위는 다음과 같습니다.

```text
1. PPTX Viewer 안정화
2. 법적/폰트/자산 안전 기반
3. PPTX Writer
4. 공통 OOXML/OPC 계층
5. PPTX Editor
6. DOCX Reader/Writer
7. XLSX Reader/Writer
8. HWPX
9. HWP
10. ODF / PDF Export
```

상세 체크리스트는 [`docs/ARCHITECTURE_ROADMAP.md`](docs/ARCHITECTURE_ROADMAP.md)를 참고하세요.

## 저작권 / 상표 / 자산 정책

이 프로젝트는 파일 포맷을 독립적으로 구현합니다.

하지 않는 것:

- Microsoft Office 실행 파일/DLL 재배포
- Hancom 실행 파일/DLL 재배포
- Office/Hancom 공식 로고·아이콘 복사
- 상용 템플릿/클립아트 무단 포함
- 라이선스 불명확 폰트 파일 번들

개발 정책은 [`docs/LEGAL_ASSET_POLICY.md`](docs/LEGAL_ASSET_POLICY.md)에 정리되어 있습니다.

Microsoft, Word, Excel and PowerPoint are trademarks of Microsoft Corporation. This project is not affiliated with or endorsed by Microsoft.

Hancom and related product names are trademarks of their respective owners. This project is not affiliated with or endorsed by Hancom.

## 폰트 정책

기본 원칙:

- Windows에 설치된 폰트를 런타임에 사용
- 앱 패키지에 시스템 폰트 파일을 복사하지 않음
- 문서에는 기본적으로 font family 이름만 기록
- 번들 폰트는 재배포 권한이 명확한 경우만 포함
- 문서/PDF embedding은 `OS/2.fsType` 메타데이터와 실제 라이선스 원문을 함께 확인

`src/FontLicensing.cs`에는 OpenType `OS/2.fsType`을 읽기 위한 초기 안전 계층이 포함되어 있습니다. 이 메타데이터는 보조 판단용이며 실제 라이선스 문구가 우선합니다.

서드파티 고지는 [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md), 라이선스 원문은 `licenses/`에서 관리합니다.

## UI 원칙

UI는 Microsoft Office 또는 Hancom UI를 복제하지 않고 자체적인 경량 데스크톱 디자인을 사용합니다.

- 독립 dark palette
- flat controls
- viewer canvas 중심 레이아웃
- 자체 제작 또는 재배포 가능한 자산만 사용
- 시스템 `Segoe UI`를 UI에 사용할 수 있으나 폰트 파일은 번들하지 않음

상세 내용: [`docs/UI_DESIGN_GUIDE.md`](docs/UI_DESIGN_GUIDE.md)

## 빌드

### 요구사항

- Windows 10/11
- Windows .NET Framework 4.x C# compiler

`.NET 8 SDK`는 필요하지 않습니다.

```bat
BUILD_EXE.cmd
```

정상 빌드 시:

```text
PowerPointLite.exe
```

빌드 스크립트는 다음을 탐색합니다.

```text
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
```

## 실행

```bat
PowerPointLite.exe
```

또는:

```bat
PowerPointLite.exe "C:\Presentations\sample.pptx"
```

## Windows 연결 프로그램

등록:

```bat
REGISTER_PPTX.cmd
```

해제:

```bat
UNREGISTER_PPTX.cmd
```

## 진단

```bat
RUN_DIAGNOSTIC.cmd
```

로그:

```text
%LOCALAPPDATA%\PptxViewer\logs\last-startup.log
```

## 저장소 구조

```text
PowerPointLite/
├─ src/
│  ├─ Program.cs
│  ├─ MainForm.Part*.cs
│  ├─ InternalPptxRenderer.Part*.cs
│  ├─ PresenterView.cs
│  ├─ Printing.cs
│  ├─ RecentFileStore.cs
│  ├─ ViewerDialogs.cs
│  ├─ FontLicensing.cs
│  └─ UiTheme.cs
├─ docs/
│  ├─ ARCHITECTURE_ROADMAP.md
│  ├─ LEGAL_ASSET_POLICY.md
│  └─ UI_DESIGN_GUIDE.md
├─ licenses/
├─ THIRD_PARTY_NOTICES.md
├─ BUILD_EXE.cmd
├─ FEATURES.txt
├─ README.md
└─ AGENTS.md
```

## 개발 상태와 테스트

상세 기능 상태는 [`FEATURES.txt`](FEATURES.txt)를 참고하세요.

AI/Codex/새 ChatGPT 세션에서 작업을 이어갈 경우 [`AGENTS.md`](AGENTS.md)를 먼저 읽습니다.

Windows 실빌드가 불가능한 환경에서는 `BUILD SUCCESS`라고 가정하지 않으며, 안정판 EXE는 실제 Windows 빌드/실행 검증 후에만 갱신합니다.

## 프로젝트 라이선스

프로젝트 코드 자체 라이선스는 아직 최종 결정되지 않았습니다. 공개 배포 전에 별도 `LICENSE`를 확정할 예정이며, 프로젝트 코드 라이선스와 번들 폰트/외부 자산 라이선스는 별도로 관리합니다.
