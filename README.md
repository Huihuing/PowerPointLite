# PowerPointLite

PowerPointLite는 Windows에서 `.pptx` / `.pptm` 프레젠테이션을 **읽기 전용으로 열고 발표하는 경량 PowerPoint Viewer**를 목표로 하는 프로젝트입니다.

Microsoft PowerPoint가 설치되어 있으면 네이티브 PowerPoint 렌더러를 우선 사용하고, 설치되어 있지 않은 환경에서는 자체 Open XML 렌더러로 PPTX를 직접 읽습니다.

> 현재 버전: **1.3**

## 목표

- PowerPoint가 없어도 일반적인 PPTX를 열 수 있는 독립 Viewer
- PowerPoint Viewer에 가까운 탐색/발표 UX
- 편집 기능보다 **보기 정확도와 발표 기능**을 우선
- 별도의 .NET 8 SDK 설치 없이 Windows에서 EXE를 빌드할 수 있는 구조 유지

## 현재 지원 기능

### 파일 및 탐색

- `.pptx`, `.pptm` 직접 열기
- `.ppt`는 PowerPoint 또는 LibreOffice가 있을 때 지원
- Drag & Drop
- 최근 파일
- 왼쪽 슬라이드 썸네일 / TOC
- TOC 자동 숨김
- 슬라이드 정렬기
- 슬라이드 검색
- 특정 슬라이드 번호로 이동
- 숨김 슬라이드 표시/건너뛰기
- 화면 맞춤(Fit)
- 확대/축소
- 전체화면

### 입력

- 마우스 휠: 이전/다음 슬라이드
- `Ctrl + 휠`: 확대/축소
- `← / →`, `PageUp / PageDown`, `Home / End`
- `F5`: 처음부터 슬라이드쇼
- `Shift + F5`: 현재 슬라이드부터 슬라이드쇼
- `N / P`: 발표 중 다음/이전
- 숫자 + `Enter`: 발표 중 특정 슬라이드 이동
- `B`: 검은 화면
- `W`: 흰 화면
- `H`: 인접 숨김 슬라이드 표시
- `F11`: 전체화면

### Presenter View

- 현재 슬라이드
- 다음 슬라이드
- 발표자 노트
- 슬라이드 제목 / 번호
- 발표 경과시간
- 이전 / 다음 / 특정 슬라이드 이동

### 인쇄

- 전체 슬라이드
- Notes Page
- 2장 Handout
- 4장 Handout
- 6장 Handout

### 자체 Open XML 렌더러

현재 내부 렌더러는 다음 요소를 처리합니다.

- 슬라이드 크기 / 비율
- 테마 색상 및 기본 테마 폰트
- Slide Master / Layout
- Placeholder 위치 상속
- 텍스트 / 기본 단락
- 굵게 / 기울임 / 밑줄
- 기본 글머리표 / 번호
- 이미지 및 Crop
- 이미지 / 도형 회전과 Flip
- 기본 도형
- 선 / 커넥터 / 화살표
- 그룹 도형
- 기본 표
- 기본 차트
- SmartArt 정적 fallback
- 하이퍼링크
- 미디어 추출/실행 fallback
- 그라데이션 / 기본 그림자
- 일부 SVG 기본 요소
- 일부 전환 효과 및 애니메이션 fallback

## 렌더링 엔진 우선순위

```text
Microsoft PowerPoint 설치됨
        ↓
PowerPoint Native Renderer

PowerPoint 없음
        ↓
Internal OpenXML Renderer

구형 .ppt 또는 일부 fallback
        ↓
LibreOffice (설치되어 있을 때)
```

PowerPoint 네이티브 엔진을 사용할 수 있는 환경에서는 호환성을 위해 항상 우선합니다.

## 빌드

### 요구사항

- Windows 10/11
- Windows에 포함된 .NET Framework 4.x C# 컴파일러

`.NET 8 SDK`는 필요하지 않습니다.

### 빌드 방법

저장소 루트에서:

```bat
BUILD_EXE.cmd
```

정상적으로 완료되면 저장소 루트에 바로 실행 가능한 포터블 EXE가 생성됩니다.

```text
PowerPointLite.exe
```

빌드 스크립트는 다음 컴파일러를 자동 탐색합니다.

```text
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
```

## 실행

```text
PowerPointLite.exe
```

또는 파일을 인자로 전달할 수 있습니다.

```bat
PowerPointLite.exe "C:\Presentations\sample.pptx"
```

PPTX 파일을 프로그램 창에 Drag & Drop할 수도 있습니다.

## Windows 연결 프로그램 등록

```bat
REGISTER_PPTX.cmd
```

등록 해제:

```bat
UNREGISTER_PPTX.cmd
```

## 진단

프로그램이 실행되지 않거나 시작 중 오류가 발생하면:

```bat
RUN_DIAGNOSTIC.cmd
```

로그 위치:

```text
%LOCALAPPDATA%\PptxViewer\logs\last-startup.log
```

로그 폴더 열기:

```bat
OPEN_LOGS.cmd
```

## 테스트 파일

저장소에는 내부 렌더러 검증용 파일이 포함됩니다.

- `tests/TEST_INTERNAL_READER.pptx`
- `tests/TEST_BASIC_FEATURES.pptx`

PowerPoint가 없는 환경에서 내부 OpenXML 렌더러 경로를 확인할 때 사용합니다.

## 저장소 구조

```text
PowerPointLite/
├─ src/
│  ├─ Program.cs
│  ├─ MainForm.Part*.cs
│  ├─ PresenterView.cs
│  ├─ Printing.cs
│  ├─ RecentFileStore.cs
│  ├─ ViewerDialogs.cs
│  └─ InternalPptxRenderer.Part*.cs
├─ tests/
│  ├─ TEST_INTERNAL_READER.pptx
│  └─ TEST_BASIC_FEATURES.pptx
├─ PowerPointLite.exe.config
├─ BUILD_EXE.cmd
├─ REGISTER_PPTX.cmd
├─ UNREGISTER_PPTX.cmd
├─ RUN_DIAGNOSTIC.cmd
├─ OPEN_LOGS.cmd
├─ FEATURES.txt
├─ README.md
└─ AGENTS.md
```

1.3부터 GitHub에서 소스를 직접 관리할 수 있도록 단일 파일을 `src/` 아래 partial class 단위로 분리했습니다. 동작 구조는 유지하면서 이후 Renderer / Viewer / Presentation / Printing 계층으로 점진적으로 정리할 수 있습니다.

## 현재 한계

PowerPoint가 없는 상태에서 아래 기능을 Microsoft PowerPoint와 100% 동일하게 재현하지는 못합니다.

- 모든 Morph 동작
- 전체 애니메이션 타임라인
- 모든 SmartArt 자동 레이아웃
- 모든 Chart subtype
- 고급 3D / Glow / Reflection
- VBA / ActiveX
- 모든 오디오/비디오 코덱 및 내장 재생
- PowerPoint 고유 렌더링의 완전한 픽셀 일치

이 기능들은 지원 범위를 계속 확대하고 있습니다.

## 개발 현황

상세 구현 상태는 [`FEATURES.txt`](FEATURES.txt)를 참고하세요.

AI 또는 새 ChatGPT 세션에서 작업을 이어갈 경우 [`AGENTS.md`](AGENTS.md)를 먼저 읽는 것을 권장합니다.

## 라이선스

현재 별도의 라이선스를 지정하지 않았습니다.

## 포터블 실행 파일 운영

`PowerPointLite.exe`는 설치 프로그램이 아니라 포터블 GUI 실행 파일입니다.

Windows에서 한 번 빌드한 뒤 저장소에 `PowerPointLite.exe`를 함께 커밋하면 다른 PC에서는:

```bat
git clone https://github.com/Huihuing/PowerPointLite.git
cd PowerPointLite
PowerPointLite.exe
```

처럼 별도의 재빌드 없이 바로 실행할 수 있습니다.

현재 `.gitignore`는 `PowerPointLite.exe`를 의도적으로 제외하지 않습니다. 따라서 private 저장소에서 최신 안정판 EXE 하나를 추적할 수 있습니다.

개발 과정에서는 모든 작은 수정마다 바이너리를 갱신하지 않고, **Windows 실제 빌드와 실행 확인이 끝난 안정 버전에서만 EXE를 교체**하는 것을 권장합니다.

## 개발 중: 발표 포인터 / 잉크

`feature/presentation-tools` 브랜치에서는 내부 슬라이드쇼용 발표 도구를 개발 중입니다.

- Arrow
- Laser Pointer
- Pen
- Highlighter
- Eraser
- 현재 슬라이드 잉크 삭제
- 전체 잉크 삭제

단축키:

```text
Ctrl+A       Arrow
Ctrl+L       Laser Pointer
Ctrl+P       Pen
Ctrl+H       Highlighter
Ctrl+E       Eraser
Ctrl+Shift+E 현재 슬라이드 잉크 삭제
```

잉크는 normalized slide coordinate로 저장되어 Fit/Zoom 변경 뒤에도 같은 위치에 표시됩니다. 이 기능은 Windows 실빌드 확인 전까지 `main`에 합치지 않습니다.

### 개발 중: 슬라이드 자동 진행

내부 슬라이드쇼는 PPTX transition의 다음 설정도 읽습니다.

- `advTm`: 설정된 시간이 지나면 자동으로 다음 슬라이드 이동
- `advClick`: 빈 슬라이드 영역의 마우스 클릭으로 진행할 수 있는지 여부

키보드의 방향키, PageUp/PageDown, N/P 같은 명시적 탐색은 `advClick=false`와 관계없이 사용할 수 있습니다.

### 개발 중: 뷰어 내부 미디어 재생

`feature/presentation-tools`에서는 PPTX에서 추출된 오디오/비디오를 클릭했을 때 먼저 PowerPointLite 내부 플레이어로 재생을 시도합니다.

- Windows `winmm.dll` / MCI 사용
- Play / Pause / Stop / Replay
- Seek bar
- 재생 시간 표시
- 비디오 surface를 별도 Viewer 창 안에 연결
- 내부 재생이 지원되지 않는 코덱이면 기존처럼 Windows 기본 앱으로 fallback

새로운 NuGet 패키지나 .NET 8 SDK 의존성은 추가하지 않습니다.
