# AGENTS.md — PowerPointLite 작업 인수인계

이 문서는 새 ChatGPT 채팅, Codex, 다른 AI 에이전트 또는 새 개발 환경에서 **프로젝트 맥락을 다시 설명하지 않고 바로 작업을 이어가기 위한 인수인계 문서**다.

## 1. 프로젝트 목적

프로젝트명: **PowerPointLite**

저장소:

```text
https://github.com/Huihuing/PowerPointLite
```

목표는 PowerPoint 편집기가 아니라 **Windows용 경량 PPTX Viewer**다.

핵심 요구사항:

1. `.pptx/.pptm` 파일을 읽기 전용으로 정상 열 것.
2. Microsoft PowerPoint가 설치되지 않아도 기본적인 프레젠테이션을 볼 수 있을 것.
3. PowerPoint가 설치되어 있으면 네이티브 PowerPoint 엔진을 우선해 최고 호환성을 확보할 것.
4. Viewer UX는 가능한 한 Microsoft PowerPoint Viewer/Slide Show 사용감에 가깝게 만들 것.
5. `.NET 8 SDK` 설치를 요구하지 않는다.
6. Windows 기본 .NET Framework 컴파일러 `csc.exe`로 GUI EXE를 빌드할 수 있어야 한다.

현재 개발 기준 버전: **1.3**

## 2. 반드시 유지할 기술 제약

### 빌드

현재 빌드는 다음 Windows 컴파일러를 사용한다.

```text
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
```

fallback:

```text
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
```

즉 코드 수정 시 **구형 .NET Framework C# 컴파일러와 호환되는 문법**을 사용해야 한다.

피해야 할 것:

- 최신 C# 전용 문법을 무심코 추가
- .NET 8 API에만 존재하는 API
- WPF/.NET 8로 무단 전환
- NuGet 패키지를 필수 조건으로 추가

현재 프로젝트는 SDK 없이 빌드되는 것이 중요한 요구사항이다.

### Timer 주의

`src/` 소스에는 `System.Threading`과 `System.Windows.Forms`가 함께 사용된다.

따라서 GUI 타이머는 반드시 다음처럼 명시한다.

```csharp
System.Windows.Forms.Timer
```

1.2에서 `Timer`를 짧은 이름으로 썼다가 아래 컴파일 오류가 발생했다.

```text
CS0104: 'Timer' is an ambiguous reference between
System.Threading.Timer and System.Windows.Forms.Timer
```

1.2.1에서 수정했다.

### System.Drawing DrawImage 주의

.NET Framework의 `Graphics.DrawImage` 오버로드는 최신 런타임과 차이가 있다.

1.0에서 다음 패턴 때문에 `CS1502/CS1503`이 발생했다.

```csharp
g.DrawImage(clone, rect, ...);
```

`rect`가 `RectangleF`일 경우 해당 7인자 오버로드가 없다.

현재는 다음처럼 사용한다.

```csharp
Rectangle.Round(rect)
```

같은 유형의 코드를 추가할 때 반드시 .NET Framework 오버로드를 확인한다.

## 3. 현재 파일 구조

주요 소스:

```text
src/
├─ Program.cs
├─ MainForm.Part*.cs
├─ PresenterView.cs
├─ Printing.cs
├─ RecentFileStore.cs
├─ ViewerDialogs.cs
└─ InternalPptxRenderer.Part*.cs
```

1.3부터 기존 6천 줄 이상의 단일 `PptxViewer.cs`를 기능 손실 없이 partial class 파일들로 분리했다.

주요 역할:

- `Program`: GUI 시작, 전역 예외 처리, 시작 오류 로그
- `MainForm`: Viewer UI, TOC/썸네일, 키보드/마우스 이동, 슬라이드쇼, Presenter View 연동, 인쇄
- `InternalPptxRenderer`: PPTX ZIP/Open XML 직접 파싱, 내부 정적 렌더링, 메타데이터/노트/숨김 슬라이드
- `PresenterViewForm`: 현재/다음 슬라이드, 노트, 타이머
- `RecentFileStore`: 최근 파일
- `ViewerDialogs`: 간단 입력 UI

지원 코드가 안정화되면 파일을 다음 식으로 분리하는 것이 바람직하다.

```text
src/
├─ Viewer/
├─ Rendering/
├─ OpenXml/
├─ Presentation/
├─ Printing/
└─ Platform/
```

하지만 구조 분리를 이유로 기능을 깨뜨리지 말 것.

## 4. 렌더링 전략

현재 엔진 우선순위:

```text
1. Microsoft PowerPoint Native
2. Internal OpenXML Renderer
3. LibreOffice fallback
```

### PowerPoint Native

PowerPoint가 설치되어 있으면 COM을 통해 일반 슬라이드 렌더링과 실제 PowerPoint Slide Show를 우선 사용한다.

고급 애니메이션/Morph/미디어 호환성은 이 경로가 최고다.

### Internal OpenXML

PowerPoint가 없을 때 `.pptx/.pptm` ZIP 패키지를 직접 읽는다.

```text
PPTX
 ↓
ZIP
 ↓
ppt/presentation.xml
 ↓
slide relationships
 ↓
slides / layouts / masters / themes / media / charts
 ↓
System.Drawing
 ↓
PNG cache
 ↓
WinForms viewer
```

목표는 외부 Office 프로그램 없이 일반 PPTX를 최대한 정상적으로 표시하는 것이다.

### LibreOffice

현재는 주로 구형 `.ppt`와 일부 fallback 용도다. LibreOffice를 사용자 필수 설치 조건으로 만들지 않는 것이 원칙이다.

## 5. 현재 Viewer 기능

### 일반 보기

- PPTX/PPTM 열기
- PPT fallback
- Drag & Drop
- 최근 파일
- 썸네일
- TOC / Auto TOC
- Fit / 확대축소 / F11
- 검색
- 특정 슬라이드 이동
- 슬라이드 정렬기
- 발표자 노트
- 숨김 슬라이드
- 인쇄

### 발표

- F5: 처음부터
- Shift+F5: 현재부터
- N/P, 방향키, PageUp/PageDown
- 숫자+Enter
- B 검은 화면
- W 흰 화면
- H 숨김 슬라이드
- 우클릭 발표 메뉴
- Presenter View
- 발표 경과시간
- 일부 전환/애니메이션 fallback

### 인쇄

- Full slide
- Notes Page
- 2/4/6 Handout

## 6. 내부 렌더러 지원 현황

구현된 기본 영역:

- slide size / solid background
- theme color / theme font
- master/layout / inherited placeholder
- text / font formatting / bullets
- images / image crop / rotation/flip
- common shapes / custom geometry 일부
- line/connectors / groups / tables
- basic charts
- SmartArt static fallback
- hyperlinks
- media extraction/fallback
- gradients / basic shadow
- basic SVG elements
- transition fallback
- basic entrance animation fallback

아직 고급 구현이 필요한 영역:

1. PowerPoint 애니메이션 타임라인 정확도
2. Morph object interpolation
3. SmartArt 실제 자동 레이아웃
4. 모든 Chart subtype
5. SVG/EMF/WMF 렌더러 정확도
6. 뷰어 내부 오디오/비디오 재생
7. 고급 효과(Glow/Reflection/3D)
8. embedded font 처리
9. 수식/Office Math 고급 렌더링
10. PowerPoint 픽셀 단위 호환성

## 7. UI 중요 요구사항

### Fit

왼쪽 TOC가 존재하는데 전체 창 크기로 Fit을 계산해 슬라이드가 잘린 문제가 있었다. 현재는 `SplitContainer.Panel2`의 실제 Viewport를 기준으로 계산한다. **이 동작을 깨뜨리지 말 것.**

### 휠

```text
Wheel Up   → Previous
Wheel Down → Next
Ctrl+Wheel → Zoom
```

TOC 위에서는 목록 자체 스크롤이 가능해야 한다.

### TOC

- 직접 표시/숨김
- Auto TOC
- Fullscreen에서는 폭을 완전히 회수

## 8. 오류 진단

GUI 프로그램은 `WinExe`이므로 시작 예외가 콘솔에 보이지 않을 수 있다.

현재 오류 로그:

```text
%LOCALAPPDATA%\PptxViewer\logs\last-startup.log
```

도구:

```text
RUN_DIAGNOSTIC.cmd
OPEN_LOGS.cmd
```

시작 오류 처리 코드를 제거하지 말 것.

## 9. 테스트 절차

코드 수정 후 최소한 다음 순서로 확인한다.

### 정적 검사

- 괄호 구조
- 동일 메서드 중복 여부
- 이벤트 핸들러 존재 여부
- 구형 C# 문법 호환 여부
- `Timer` 모호성
- `Graphics.DrawImage` 오버로드

### Windows 실제 빌드

```bat
BUILD_EXE.cmd
```

반드시 `BUILD SUCCESS`까지 확인한다.

### 실행 테스트

1. `PowerPointLite.exe` 창이 정상적으로 뜨는지
2. `tests/TEST_INTERNAL_READER.pptx`
3. `tests/TEST_BASIC_FEATURES.pptx`
4. 실제 사용자 PPTX
5. TOC를 켠 상태에서 Fit
6. TOC 숨김 후 Fit
7. 휠 이전/다음
8. Ctrl+휠 zoom
9. F5
10. Shift+F5
11. Presenter View
12. Notes
13. Print dialog

## 10. Git 규칙

기본 브랜치:

```text
main
```

커밋 메시지는 **한글 - 영어** 형식을 사용한다.

예시:

```text
빌드 Timer 충돌 수정 - Fix ambiguous WinForms Timer references
뷰어 문서 및 인수인계 추가 - Add README and AI handoff documentation
발표자 보기 기능 추가 - Add presenter view support
```

원칙:

- 의미 단위로 커밋
- 아주 작은 수정마다 micro-commit 남발하지 않기
- 작업 시작 전 최신 `main` 상태 확인
- 작업 후 테스트 가능한 상태로 `main`에 반영
- cache와 사용자 로그는 커밋하지 않기
- `PowerPointLite.exe`는 Windows 실빌드/실행 검증이 끝난 안정판만 private 저장소에 갱신 가능

## 11. Git에 올리지 말아야 할 것

```text
*.pdb
bin/
obj/
.vs/
logs/
Cache/
*.user
```

테스트용 PPTX/PNG는 프로젝트 테스트 자산이므로 저장소에 포함 가능하다.

## 12. 다음 작업 우선순위

다른 AI가 다음 작업을 이어갈 경우 권장 순서:

1. Windows `csc.exe` 빌드 오류 0 유지
2. 뷰어 내부 미디어 재생
3. 애니메이션 timeline 확장
4. 레이저 포인터/펜
5. Presenter View 멀티모니터 배치
6. SmartArt 실제 정적 레이아웃 개선
7. chart subtype 확대
8. SVG/EMF/WMF 호환성
9. embedded fonts
10. 소스 모듈 분리

**기능 추가보다 먼저 현재 빌드/실행 상태를 깨뜨리지 않는 것이 우선이다.**

## 13. 포터블 EXE 저장소 운영

최종 사용자 실행 파일 이름은 `PowerPointLite.exe`로 통일한다.

`BUILD_EXE.cmd`는 개발/빌드 편의 도구일 뿐이며 사용자가 프로그램을 실행할 때 필요한 런처가 아니다.

```text
src\*.cs
  ↓ Windows csc.exe 빌드
PowerPointLite.exe
  ↓
GUI 직접 실행
```

private 저장소에서는 최신 안정판 `PowerPointLite.exe`를 루트에 커밋해도 된다.

규칙:

1. 모든 micro-commit마다 EXE를 교체하지 않는다.
2. Windows에서 `BUILD SUCCESS` 확인 후 갱신한다.
3. 기본 테스트 PPTX와 실제 PPTX 실행 확인 후 갱신한다.
4. 소스 커밋과 바이너리 갱신은 의미 단위로 묶는다.
5. AI 실행 환경이 Windows EXE를 직접 컴파일할 수 없으면 소스만 갱신하고 성공했다고 가장하지 않는다.

## 14. 1.3 소스 분리 규칙

GitHub 동기화를 위해 단일 소스를 partial class 기반 파일로 분리했다.

- `MainForm.Part*.cs`는 모두 `public sealed partial class MainForm : Form`
- `InternalPptxRenderer.Part*.cs`는 모두 `internal static partial class InternalPptxRenderer`
- private 필드/메서드는 partial class 전체에서 공유되므로 기존 접근 수준을 임의 변경하지 않는다.
- 새 기능은 가능한 한 관련 파일에 추가하고, 무조건 새 Part 파일을 늘리지는 않는다.
- `BUILD_EXE.cmd`는 `src\*.cs` 전체를 한 번에 컴파일한다.

## 개발 브랜치: feature/presentation-tools

내부 슬라이드쇼의 발표 포인터/잉크 기능을 개발 중이다.

관련 파일:

```text
src/Annotations.cs
src/MainForm.Part03.cs
```

설계:
- `Annotations.cs`의 `OnShown`에서 viewer overlay 이벤트를 연결
- `ProcessCmdKey`에서 PowerPoint 계열 포인터 단축키 처리
- stroke는 슬라이드 상대(normalized) 좌표로 저장
- 내부 슬라이드쇼에서만 overlay 표시
- pointer tool이 Arrow가 아니면 `OnViewerMouseClick`이 slide advance/hyperlink를 실행하지 않도록 Part03에 guard가 있음

Windows 실제 `BUILD SUCCESS`와 기본 슬라이드쇼 테스트 전에 main으로 병합하지 않는다.

### feature/presentation-tools — 자동 슬라이드 진행

관련 파일:

```text
src/SlideshowTiming.cs
src/SlideshowTimingRenderer.cs
src/MainForm.Part03.cs
```

- 100ms WinForms polling timer가 `currentIndex` 변경을 감지해 slide start time을 갱신한다.
- `advTm`은 밀리초 기준 자동 진행 시간으로 사용한다.
- `advClick=false`면 빈 영역 mouse click만 막고 명시적 keyboard navigation은 유지한다.
- timing timer는 반드시 `System.Windows.Forms.Timer` 완전한 이름을 사용한다.
- 기존 transition renderer의 `TransitionSpec` 구조는 건드리지 않아 회귀 범위를 줄였다.

### feature/presentation-tools — 내부 미디어 플레이어

관련 파일:

```text
src/MediaPlayer.cs
src/MainForm.Part03.cs
```

설계 원칙:

- 외부 라이브러리를 추가하지 않고 `winmm.dll` MCI를 사용한다.
- `PortableMediaPlayerForm.TryShow()`가 성공하면 내부에서 재생한다.
- MCI open 실패 시 기존 `Process.Start(... UseShellExecute=true)` 경로로 fallback한다.
- Timer는 반드시 `System.Windows.Forms.Timer`로 완전 수식한다.
- MP4 등 실제 코덱 지원 범위는 Windows의 MCI/코덱 환경에 따라 달라지므로 Windows 실테스트 전에는 완전 지원으로 표기하지 않는다.
