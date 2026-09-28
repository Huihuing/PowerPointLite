# AGENTS.md — PowerPointLite 작업 인수인계

새 ChatGPT/Codex/다른 AI 세션이 기존 대화를 다시 복원하지 않고 바로 개발을 이어가기 위한 기준 문서다.

## 1. 프로젝트 목표

저장소: `Huihuing/PowerPointLite`

PowerPointLite는 편집기가 아니라 **Windows용 경량 PPTX/PPTM Viewer**다.

핵심 목표:

1. `.pptx/.pptm`을 읽기 전용으로 정상 열기
2. Microsoft PowerPoint가 없어도 일반적인 프레젠테이션 표시
3. PowerPoint가 설치된 환경에서는 COM 네이티브 렌더러/슬라이드쇼 우선
4. PowerPoint Viewer에 가까운 보기/발표 UX
5. `.NET 8 SDK`를 필수 요구하지 않기
6. Windows .NET Framework `csc.exe`로 GUI EXE 빌드 가능하게 유지

현재 개발 기준 버전: **1.3**

## 2. 빌드 제약

기본 컴파일러:

```text
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
```

fallback:

```text
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
```

따라서 최신 C# 전용 문법, .NET 8 전용 API, 필수 NuGet 의존성, WPF/.NET 8 임의 전환을 피한다. Windows의 `BUILD_EXE.cmd` 실빌드가 최종 기준이다.

### 과거 컴파일 오류 — Timer

`System.Threading`과 `System.Windows.Forms`가 함께 import되어 있다. GUI 타이머는 반드시:

```csharp
System.Windows.Forms.Timer
```

처럼 명시한다. 1.2에서 `Timer`만 사용해 `CS0104`가 발생했고 1.2.1에서 수정했다.

### 과거 컴파일 오류 — Graphics.DrawImage

.NET Framework의 `Graphics.DrawImage` 오버로드는 최신 런타임과 차이가 있다. 7인자 계열에서 대상 `RectangleF`를 그대로 넘겨 `CS1502/CS1503`이 발생한 적이 있다.

현재 해당 경로는:

```csharp
Rectangle.Round(rect)
```

를 사용한다. 이미지 렌더 코드를 추가할 때 .NET Framework 오버로드를 확인한다.

## 3. 현재 소스 구조

1.3부터 기존 단일 `PptxViewer.cs`를 partial class 단위로 분리했다.

```text
src/
├─ Program.cs
├─ MainForm.Part01.cs ... MainForm.Part05.cs
├─ PresenterView.cs
├─ Printing.cs
├─ RecentFileStore.cs
├─ ViewerDialogs.cs
└─ InternalPptxRenderer.Part01.cs ... Part08.cs
```

규칙:

- `MainForm.Part*.cs`는 `public sealed partial class MainForm : Form`
- `InternalPptxRenderer.Part*.cs`는 `internal static partial class InternalPptxRenderer`
- private 필드/메서드는 partial class 전체에서 공유된다.
- 구조 정리를 이유로 기능을 깨뜨리지 않는다.
- `BUILD_EXE.cmd`는 `src\*.cs` 전체를 한 번에 컴파일한다.

주요 역할:

- `Program` — GUI 시작, 전역 예외 처리, 시작 로그
- `MainForm` — Viewer UI, TOC, 탐색, 슬라이드쇼, Presenter 연동, 인쇄
- `InternalPptxRenderer` — PPTX ZIP/Open XML 파싱 및 System.Drawing 렌더링
- `PresenterViewForm` — 현재/다음 슬라이드, 노트, 발표 타이머
- `RecentFileStore` — 최근 파일
- `ViewerDialogs` — 검색/이동 입력 UI

## 4. 렌더링 전략

우선순위:

```text
1. Microsoft PowerPoint Native
2. Internal OpenXML Renderer
3. LibreOffice fallback
```

### Native PowerPoint

PowerPoint가 설치되어 있으면 COM을 이용해 `Slide.Export` 기반 정적 렌더링과 실제 PowerPoint Slide Show를 우선 사용한다. Morph/고급 애니메이션/미디어 호환성은 이 경로가 가장 높다.

### Internal OpenXML

PowerPoint가 없을 때:

```text
PPTX ZIP
→ ppt/presentation.xml
→ slide relationships
→ slides/layouts/masters/themes/media/charts
→ System.Drawing
→ PNG cache
→ WinForms Viewer
```

현재 처리 범위:

- slide size / background / theme color/font
- master/layout/placeholder 상속
- 텍스트와 기본 서식, bullets/numbering
- 이미지, crop, alpha, rotation/flip
- SVG 기본 subset, GDI+가 지원하는 EMF/WMF
- 기본/커스텀 geometry 도형
- line/connectors/dash/arrowheads
- groups / tables
- 기본 column/bar/line/pie chart
- SmartArt 정적 approximation
- hyperlinks / internal slide actions
- media 추출 후 외부 실행 fallback
- gradients / pattern fill / 기본 shadow
- 일부 transition 및 entrance animation fallback

### LibreOffice

주로 구형 `.ppt` 또는 fallback 용도다. 사용자 필수 의존성으로 만들지 않는다.

## 5. Viewer UX 요구사항

현재 기능:

- Drag & Drop / Recent files
- thumbnails / TOC / Auto TOC
- 정확한 right viewport 기준 Fit
- wheel previous/next / Ctrl+wheel zoom
- arrow/PageUp/PageDown/Home/End
- F11 fullscreen
- search / go-to / slide sorter
- speaker notes / hidden slides
- F5 / Shift+F5
- N/P, 숫자+Enter, B/W, H
- 우클릭 slideshow menu
- Presenter View
- print: full slide / notes / 2·4·6 handout

### 반드시 유지할 Fit 동작

과거 TOC 폭을 제외하지 않고 전체 창 기준으로 Fit을 계산해 슬라이드가 잘리는 문제가 있었다. 현재는 `SplitContainer.Panel2` 실제 viewport 크기를 기준으로 계산한다. 이 동작을 깨뜨리지 않는다.

### Wheel

슬라이드 화면 위:

```text
Wheel Up   → Previous
Wheel Down → Next
Ctrl+Wheel → Zoom
```

TOC 위에서는 TOC 목록 자체가 스크롤되어야 한다.

## 6. 오류 진단

WinExe라 시작 예외가 콘솔에 보이지 않을 수 있다.

```text
%LOCALAPPDATA%\PptxViewer\logs\last-startup.log
```

도구:

```text
RUN_DIAGNOSTIC.cmd
OPEN_LOGS.cmd
```

전역 오류 처리 코드를 제거하지 않는다.

## 7. 테스트 절차

수정 후 최소 확인:

1. 정적 괄호/중복 메서드 검사
2. `Timer` 모호성 검사
3. `Graphics.DrawImage` 오버로드 검사
4. Windows에서 `BUILD_EXE.cmd`
5. 반드시 `BUILD SUCCESS` 확인
6. `PowerPointLite.exe` 실행
7. 개발 ZIP의 `TEST_INTERNAL_READER.pptx` / `TEST_BASIC_FEATURES.pptx`가 있으면 확인
8. 실제 사용자 PPTX
9. TOC on/off 각각 Fit
10. wheel / Ctrl+wheel
11. F5 / Shift+F5
12. Presenter View / Notes
13. Print dialog

이 AI 환경에서 Windows EXE를 실제 컴파일할 수 없으면 **빌드 성공했다고 주장하지 않는다.**

## 8. Git 규칙

기본 브랜치: `main`

커밋 메시지는 반드시 **한글 - English** 형식.

예:

```text
소스 구조 분리 및 포터블 빌드 정리 - Sync split source and portable build workflow
빌드 Timer 충돌 수정 - Fix ambiguous WinForms Timer references
발표자 보기 기능 추가 - Add presenter view support
```

원칙:

- 의미 단위 커밋
- main에 micro-commit 남발 금지
- 작업 전 최신 main 확인
- cache/로그/디버그 산출물 커밋 금지
- `PowerPointLite.exe`는 Windows 실빌드/실행 검증이 끝난 안정판만 private 저장소에 갱신 가능

## 9. 포터블 EXE

최종 실행 파일명:

```text
PowerPointLite.exe
```

`BUILD_EXE.cmd`는 빌드 도구일 뿐 런처가 아니다.

```text
src\*.cs
→ Windows csc.exe
→ PowerPointLite.exe
→ GUI 직접 실행
```

private repo에 안정판 EXE 하나를 보관하는 것은 허용한다. 단 실제 Windows 빌드 검증 전에는 EXE를 올리지 않는다.

## 10. 아직 남은 고급 영역

우선순위:

1. Windows 실빌드 오류 0 유지
2. 뷰어 내부 audio/video 재생
3. PowerPoint animation timeline 정확도 확대
4. 레이저 포인터/펜
5. Presenter View 멀티모니터 배치
6. SmartArt 실제 정적 레이아웃 개선
7. chart subtype 확대
8. SVG/EMF/WMF 정확도
9. embedded fonts
10. Office Math
11. Morph object interpolation
12. Glow/Reflection/3D

**새 기능 추가보다 현재 빌드/실행 상태를 깨뜨리지 않는 것이 우선이다.**

## 개발 브랜치: feature/presentation-tools

내부 슬라이드쇼의 발표 포인터/잉크 기능을 개발 중이다.

관련 파일:

```text
src/Annotations.cs
src/MainForm.Part03.cs
```

설계:

- `Annotations.cs`의 `OnShown`에서 viewer overlay 이벤트 연결
- `ProcessCmdKey`에서 PowerPoint 계열 포인터 단축키 처리
- stroke는 슬라이드 상대(normalized) 좌표로 저장
- 내부 슬라이드쇼에서만 overlay 표시
- pointer tool이 Arrow가 아니면 `OnViewerMouseClick`이 slide advance/hyperlink를 실행하지 않도록 Part03에 guard 유지

단축키:

```text
Ctrl+A       Arrow
Ctrl+L       Laser
Ctrl+P       Pen
Ctrl+H       Highlighter
Ctrl+E       Eraser
Ctrl+Shift+E 현재 슬라이드 잉크 삭제
```

Windows 실제 `BUILD SUCCESS`와 기본 슬라이드쇼 테스트 전에 main으로 병합하지 않는다.
