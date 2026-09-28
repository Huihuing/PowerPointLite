# PowerPointLite

PowerPointLite는 Windows에서 `.pptx` / `.pptm` 프레젠테이션을 **읽기 전용으로 열고 발표하는 경량 PowerPoint Viewer**를 목표로 하는 프로젝트입니다.

Microsoft PowerPoint가 설치되어 있으면 네이티브 PowerPoint 렌더러를 우선 사용하고, 설치되어 있지 않은 환경에서는 자체 Open XML 렌더러로 PPTX를 직접 읽습니다.

> 현재 개발 기준 버전: **1.2.1**

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

### 발표 및 단축키

- `F5`: 처음부터 슬라이드쇼
- `Shift + F5`: 현재 슬라이드부터 슬라이드쇼
- 마우스 휠: 이전/다음 슬라이드
- `Ctrl + 휠`: 확대/축소
- `← / →`, `PageUp / PageDown`, `Home / End`
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

정상적으로 완료되면 같은 폴더에 다음 파일이 생성됩니다.

```text
PptxViewer.exe
```

빌드 스크립트는 다음 컴파일러를 자동 탐색합니다.

```text
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
```

## 실행

```text
PptxViewer.exe
```

또는 파일을 인자로 전달할 수 있습니다.

```bat
PptxViewer.exe "C:\Presentations\sample.pptx"
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

내부 렌더러 검증용 파일:

- `TEST_INTERNAL_READER.pptx`
- `TEST_BASIC_FEATURES.pptx`

PowerPoint가 없는 환경에서 내부 OpenXML 렌더러 경로를 확인할 때 사용합니다.

## 저장소 구조

```text
PowerPointLite/
├─ PptxViewer.cs
├─ PptxViewer.exe.config
├─ BUILD_EXE.cmd
├─ REGISTER_PPTX.cmd
├─ UNREGISTER_PPTX.cmd
├─ RUN_DIAGNOSTIC.cmd
├─ OPEN_LOGS.cmd
├─ FEATURES.txt
├─ TEST_INTERNAL_READER.pptx
├─ TEST_BASIC_FEATURES.pptx
├─ README.md
└─ AGENTS.md
```

현재는 빠른 호환성 개발을 위해 핵심 코드가 `PptxViewer.cs`에 집중되어 있습니다. 기능이 안정화되면 Renderer / Viewer / Presentation / Printing 계층으로 분리할 예정입니다.

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

## 개발 인수인계

새 ChatGPT 채팅, Codex 또는 다른 AI에서 작업을 이어갈 경우 [`AGENTS.md`](AGENTS.md)를 먼저 읽으세요. 빌드 제약, 현재 구조, 과거 오류와 다음 우선순위를 정리해 두었습니다.

## 라이선스

현재 별도의 라이선스를 지정하지 않았습니다.
