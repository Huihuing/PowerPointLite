# AGENTS.md — PowerPointLite 작업 인수인계

새 ChatGPT/Codex/다른 AI 세션이 기존 대화를 복원하지 않고 바로 개발을 이어가기 위한 기준 문서다.

## 1. 현재 프로젝트 상태

저장소:

```text
Huihuing/PowerPointLite
```

현재 `main`의 안정 기능은 Windows용 PPTX/PPTM Viewer다.

장기 목표는 Viewer를 유지하면서 다음 포맷을 읽고 작성/편집/저장할 수 있는 경량 문서 프로그램으로 단계적으로 확장하는 것이다.

```text
PPTX
DOCX
XLSX
HWPX
HWP
ODT
ODS
ODP
```

우선순위:

```text
PPTX Viewer 안정화
→ 법적/폰트/자산 안전 기반
→ PPTX Writer
→ 공통 OOXML/OPC 계층
→ Presentation internal model
→ PPTX Editor
→ DOCX
→ XLSX
→ HWPX
→ HWP
→ ODF / PDF Export
```

현재 안정 기준 버전은 **1.3**이며 office/editor 기능은 feature 단계다.

## 2. 브랜치 전략

```text
main
  Windows 실빌드/실행 검증을 통과한 안정 기반

feature/presentation-tools
  레이저/펜/자동진행/내부 미디어 등 발표 기능 개발

feature/office-foundation
  멀티포맷 아키텍처, 폰트/라이선스, OPC, Writer, Editor 개발
```

feature 브랜치를 Windows에서 검증하기 전 main으로 합치지 않는다.

## 3. 저작권 / 상표 / 자산 최우선 규칙

먼저 읽을 문서:

```text
docs/LEGAL_ASSET_POLICY.md
docs/UI_DESIGN_GUIDE.md
docs/ARCHITECTURE_ROADMAP.md
docs/EDITOR_FOUNDATION.md
```

허용 방향:

- 공개/공식 파일 형식 명세 기반 자체 Reader/Writer
- OOXML 직접 생성/편집/저장
- 공식 공개 HWPX/HWP 명세 범위 자체 구현
- ODF 공개 표준 기반 구현
- 시스템 설치 폰트 런타임 사용

금지:

- Microsoft Office 실행 파일/DLL 재배포
- Hancom 실행 파일/DLL 재배포
- Office/Hancom 공식 로고/아이콘/UI 이미지 복제
- 상용 템플릿/클립아트 무단 포함
- 라이선스 불명확 TTF/OTF 번들
- 인터넷에서 가져온 타인 문서를 fixture로 커밋

외부 라이브러리/폰트/아이콘을 추가하면 반드시:

```text
THIRD_PARTY_NOTICES.md
licenses/
```

를 갱신한다.

## 4. 폰트 안전 규칙

관련 파일:

```text
src/FontLicensing.cs
src/FontLicenseService.cs
```

현재 정책:

- Windows 설치 font family를 사용할 수 있다.
- 시스템 font binary를 앱에 복사하지 않는다.
- 문서에는 기본적으로 font family 이름만 기록한다.
- font embedding은 기본 OFF다.
- OpenType/TrueType `OS/2.fsType`은 보조 정보다.
- 실제 LICENSE 원문이 fsType보다 우선한다.
- metadata가 허용처럼 보여도 명시적 license review가 없으면 embedding을 기본 거부한다.
- app font bundling 권리는 fsType으로 판정하지 않는다.

향후 PPTX/DOCX/XLSX/HWPX/PDF embedding은 같은 `FontLicenseService`를 사용한다.

## 5. UI 원칙

Microsoft Office 또는 Hancom UI를 시각적으로 복제하지 않는다.

현재 기반:

```text
src/UiTheme.cs
src/PresentationEditorForm.cs
```

UI 방향:

- 자체 dark palette
- flat controls
- slide/document canvas 중심
- 공식 Office/Hancom icons/logos 사용 금지
- 시스템 `Segoe UI` 사용 가능, 폰트 파일 번들 금지
- 100/125/150% DPI 고려
- Viewer의 TOC/Fit 회귀 금지

Editor는 왼쪽 slide list / 중앙 canvas / 오른쪽 properties 구조를 사용한다.

## 6. 빌드 제약

기본 컴파일러:

```text
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
```

fallback:

```text
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
```

피해야 할 것:

- 최신 C# 전용 문법
- .NET 8 전용 API
- 필수 NuGet 의존성
- 임의의 WPF/.NET 8 전환

현재 빌드:

```text
BUILD_EXE.cmd
→ src\*.cs
→ PowerPointLite.exe
```

Windows 실제 `BUILD SUCCESS`가 최종 기준이다.

### 과거 컴파일 오류 — Timer

`System.Threading`과 `System.Windows.Forms`가 함께 import될 수 있다.

GUI Timer는 반드시:

```csharp
System.Windows.Forms.Timer
```

처럼 완전 수식한다.

### 과거 컴파일 오류 — Graphics.DrawImage

.NET Framework `Graphics.DrawImage` overload는 최신 런타임과 다를 수 있다.

7인자 destination에 `RectangleF`를 전달해 `CS1502/CS1503`이 발생한 적이 있으므로 필요하면:

```csharp
Rectangle.Round(rect)
```

을 사용한다.

## 7. 현재 Viewer 구조

주요 파일:

```text
src/Program.cs
src/MainForm.Part01.cs ... Part05.cs
src/InternalPptxRenderer.Part01.cs ... Part08.cs
src/PresenterView.cs
src/Printing.cs
src/RecentFileStore.cs
src/ViewerDialogs.cs
```

렌더링 우선순위:

```text
1. Microsoft PowerPoint Native
2. Internal OpenXML Renderer
3. LibreOffice fallback
```

Office/LibreOffice 바이너리를 앱에 포함하지 않는다.

### 반드시 유지할 Viewer 동작

- TOC를 켜도 Fit은 `SplitContainer.Panel2` 실제 viewport 기준
- slide area wheel = previous/next
- Ctrl+wheel = zoom
- TOC 위 wheel = TOC 자체 scroll
- F5 / Shift+F5
- Presenter View / Notes / Print
- startup crash logging

## 8. 공통 OOXML / OPC 기반

관련 파일:

```text
src/OpcPackage.cs
```

현재 구현:

- XML part read/write
- `[Content_Types].xml` helper
- relationships read/write
- target part resolution
- content type mapping

향후 PPTX/DOCX/XLSX가 이 공통 계층을 공유한다.

기존 PPTX Reader를 한 번에 rewrite하지 말고 테스트 가능한 단위로 옮긴다.

## 9. Presentation internal model

관련 파일:

```text
src/PresentationModel.cs
```

현재 구조:

```text
PresentationDocument
 └─ PresentationSlide[]
     └─ PresentationTextBox[]
```

`PresentationTextBox`는:

```text
Text
FontFamily
FontSizePoints
Bold
Italic
ColorHex
Alignment
X / Y / Width / Height (EMU)
```

를 가진다.

UI는 OOXML XML을 직접 수정하지 않는다.

## 10. PPTX Writer

관련 파일:

```text
src/PptxWriter.cs
src/PptxWriterDiagnostics.cs
RUN_WRITER_SELFTEST.cmd
```

현재 Writer 기능:

- new PPTX package
- content types / relationships
- 자체 neutral theme/master/layout
- multi-slide presentation
- multiple text boxes
- font family name / size / bold / italic
- alignment / text color
- slide add/delete/reorder model 결과 저장
- 임시 `.writing` 파일을 만든 뒤 destination으로 교체

Writer는 Microsoft/Hancom template을 복사하지 않는다.

Theme에 font family 이름을 참조할 수 있으나 font binary를 번들하지 않는다.

Windows 빌드 후:

```bat
RUN_WRITER_SELFTEST.cmd
```

을 실행한다.

이 테스트는 자체 생성 3-slide 문서에 대해 package 구조와 read-edit-write round trip을 검사한다.

## 11. 편집 세션 / 안전 Reader

관련 파일:

```text
src/PresentationEditSession.cs
src/PptxEditableReader.cs
src/EditorIntegration.cs
```

새 문서:

```text
Ctrl+N
```

현재 파일 편집 시도:

```text
Ctrl+Shift+E
```

매우 중요:

**아무 외부 PPTX나 Writer로 다시 저장하면 안 된다.**

현재 `PptxEditableReader`는:

- 현재 PowerPointLite Writer가 만든 문서인지
- model이 표현하지 못하는 slide content가 있는지

를 검사한다.

안전 round-trip 조건을 만족하지 않으면 Editor를 열지 않고 Viewer만 유지한다.

이 guard를 제거하지 않는다.

향후 arbitrary PPTX edit은 unknown/unsupported part preservation이 구현된 뒤에만 허용한다.

## 12. 초기 Presentation Editor UI

관련 파일:

```text
src/PresentationEditorForm.cs
```

현재 기능:

- slide list
- slide add/delete/reorder
- slide canvas
- text box selection
- text editing
- Windows installed font picker
- font size / bold / italic
- left/center/right alignment
- text color
- Save / Save As
- dirty status
- close-save confirmation

다음 작업:

- Add Text / Delete Object
- drag/resize object
- undo/redo
- copy/paste
- image insert
- basic shapes
- better thumbnails
- keyboard accessibility

Office Ribbon을 복제하지 않는다.

## 13. Macro 문서

향후:

```text
.pptm
.docm
.xlsm
```

을 열 수 있어도 VBA 실행은 기본 지원하지 않는다.

권장:

```text
open
→ normal content edit
→ unknown/VBA parts preserve when practical
→ save
→ macro execution = disabled
```

## 14. 테스트 자산

프로젝트가 직접 만든 자산만 저장소 fixture로 사용한다.

금지:

- 인터넷 강의자료
- 회사/학교 문서
- 타인 HWP
- 상용 PPT/XLSX template

예정 fixture:

```text
TEST_BASIC_PPTX.pptx
TEST_BASIC_DOCX.docx
TEST_BASIC_XLSX.xlsx
TEST_BASIC_HWPX.hwpx
```

## 15. 테스트 절차

모든 수정 후 최소 확인:

1. brace/method duplication 정적 검사
2. `Timer` 모호성 검사
3. `Graphics.DrawImage` overload 검사
4. Windows `BUILD_EXE.cmd`
5. `BUILD SUCCESS`
6. `PowerPointLite.exe` 실행
7. 기존 PPTX Viewer 테스트
8. TOC on/off Fit
9. wheel / Ctrl+wheel
10. F5 / Shift+F5
11. Presenter / Notes / Print
12. `RUN_WRITER_SELFTEST.cmd`
13. Writer output을 Viewer에서 다시 열기
14. Editor 새 문서 생성/저장/재열기
15. external PPTX safe-edit guard 확인

Windows EXE를 실제 컴파일할 수 없는 환경에서는 성공했다고 주장하지 않는다.

## 16. Git 규칙

기본 브랜치:

```text
main
```

커밋 메시지:

```text
한글 - English
```

예:

```text
프레젠테이션 내부 모델 추가 - Add editable presentation document model
다중 슬라이드 PPTX Writer 확장 - Connect presentation model to multi-slide PPTX writer
PPTX 편집용 안전 Reader 기반 추가 - Add guarded editable PPTX model reader
```

원칙:

- 의미 단위 commit
- main에 micro-commit 남발 금지
- cache/log/debug output 커밋 금지
- 안정 EXE는 Windows 실빌드/실행 검증 후 갱신

## 17. 다음 구현 우선순위

`feature/office-foundation` 기준:

1. Windows `csc.exe` compile 오류 0 확인
2. Writer self-test 실제 실행
3. Editor Add Text / delete object
4. object drag/resize
5. undo/redo command stack
6. image insert + media part abstraction
7. basic shapes
8. unknown-part preservation
9. existing PPTX editable import 범위 확대
10. DOCX Reader/Writer
11. XLSX Reader/Writer
12. HWPX
13. HWP

상세 로드맵:

```text
docs/ARCHITECTURE_ROADMAP.md
```

**기능 추가보다 기존 PPTX Viewer 회귀 방지, 저작권/라이선스 안전, 사용자 문서 손실 방지, 실제 Windows 빌드 안정성이 우선이다.**
