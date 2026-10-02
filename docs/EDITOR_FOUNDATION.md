# Presentation Editor Foundation

`feature/office-foundation`의 PPTX Editor는 기존 Viewer를 대체하지 않고 별도 창으로 동작한다.

## 진입

```text
Ctrl+N       새 프레젠테이션
Ctrl+Shift+E 현재 열려 있는 PPTX를 안전 편집 모드로 시도
```

현재 파일 편집은 아무 PPTX에나 허용하지 않는다.

`Ctrl+Shift+E`는 `PptxEditableReader`가 다음을 확인한 경우에만 Editor를 연다.

1. 현재 Writer 계열에서 생성된 문서인지
2. 편집 모델이 표현하지 못하는 요소가 없는지
3. read → model → write 과정에서 알려진 손실 위험이 없는지

조건을 만족하지 않으면 원본 파일은 수정하지 않고 Viewer 모드만 유지한다.

이 제한은 기능 부족을 숨기기 위한 것이 아니라 **사용자 문서를 손실시키지 않기 위한 의도적인 안전장치**다.

## Viewer와 Editor의 표현력 기준

Viewer의 `InternalPptxRenderer`가 PPTX 시각 표현의 기준 renderer다.

Editor canvas는 개체 선택/이동/크기 조절을 빠르게 처리하기 위한 interactive model view이므로 Viewer보다 단순한 표현을 사용할 수 있다. 따라서 **Editor canvas가 최종 PPTX 모습의 기준이 되어서는 안 된다.**

`AdvancedEditorFidelityExtension`은 Editor 상단에 `렌더러 미리보기 / Renderer Preview`를 추가한다.

```text
현재 PresentationDocument
 ↓
temporary PPTX package
 ↓
InternalPptxRenderer
 ↓
Viewer와 같은 렌더링 경로의 read-only preview
```

미리보기에서는 현재 슬라이드 선택을 따라가며, 편집 후 `미리보기 새로고침`으로 다시 렌더링한다. 임시 package와 이미지 cache는 Editor 종료 시 삭제한다.

향후 Editor 표현력을 높일 때도 Viewer renderer와 별개의 두 번째 고급 renderer를 새로 만드는 방향은 피한다. 가능한 한 Viewer renderer의 layout/text/table/image 처리 계층을 재사용해 두 경로의 시각 차이를 줄인다.

## 언어

기본 UI 언어는 한국어다.

`src/UiLocalization.cs`와 `src/MainForm.Localization.cs`가 한국어/English 전환을 담당하며 설정은 사용자 PC의 `%LOCALAPPDATA%\PowerPointLite\settings.ini`에 저장한다.

새 UI를 만들 때 영어 문자열을 하드코딩한 채 방치하지 말고 기존 localization dictionary 또는 해당 기능의 언어 갱신 경로에 연결한다.

## 현재 Editor UI

주 Editor는 `AdvancedPresentationEditorForm`이다.

- 독립 dark application chrome
- 왼쪽 slide thumbnail list
- 가운데 16:9 interactive canvas
- Viewer renderer 기반 read-only fidelity preview
- 오른쪽 object properties
- slide 추가/삭제/순서 이동
- text box 추가/삭제
- text 수정
- Windows 설치 font family 선택
- font size / bold / italic
- left / center / right alignment
- text color
- 기본 shape 추가 및 fill/line color
- PNG/JPEG/GIF/BMP 로컬 이미지 삽입
- object drag / resize
- keyboard nudge
- snapshot undo / redo
- text / shape / image copy-paste
- basic table manager/editor
- table cell text edit
- table first-row header style
- table preview overlay 및 double-click 재편집
- Save / Save As
- dirty-state 표시
- 닫을 때 저장 확인

Microsoft Office/Hancom Ribbon, 공식 아이콘, 이미지 자산을 복제하지 않는다.

## 내부 모델

```text
PresentationDocument
 └─ PresentationSlide[]
     ├─ PresentationTextBox[]
     ├─ PresentationShape[]
     ├─ PresentationImage[]
     └─ PresentationTable[]
         └─ PresentationTableCell[]
```

모델은 UI와 PPTX XML을 직접 결합하지 않는다. Editor는 모델을 수정하고 Writer가 모델을 OOXML로 직렬화한다.

폰트는 **family 이름만 문서에 기록**하며 TTF/OTF를 자동 embedding하지 않는다.

이미지는 사용자가 명시적으로 선택한 로컬 파일만 읽는다. 외부 클립아트/템플릿을 자동 다운로드하지 않는다.

## Save 구조

```text
Editor
 ↓
PresentationEditSession
 ↓
PresentationDocument
 ↓
PresentationPackageWriter
 ├─ PptxWriter
 └─ PptxTableWriter
 ↓
staged package
 ↓
backup/replace
 ↓
<destination>.pptx
```

Writer 중간 실패 시 stage 파일을 제거하고 기존 destination을 가능한 한 복원하도록 구성한다.

## Writer 자체 테스트

Windows 실제 빌드 후:

```bat
RUN_WRITER_SELFTEST.cmd
```

현재 자체 테스트는 프로젝트 코드만으로 다음을 만든다.

1. 3-slide PPTX
2. text box
3. rounded rectangle
4. 자체 생성 PNG test image
5. basic table
6. 필수 OPC part 검사
7. `PptxEditableReader` 재열기
8. text / shape / table 수정
9. 재저장
10. 두 번째 Reader로 round-trip 결과 확인

인터넷 문서나 Microsoft/Hancom 템플릿을 fixture로 사용하지 않는다.

## 현재 남은 Editor 작업

- Viewer renderer와 interactive canvas의 공통 layout/render primitive 확대
- rich text run-level 편집 모델
- chart/SmartArt/media 등 고급 요소는 Viewer fidelity를 보존하면서 단계적 편집 지원
- table을 일반 object selection과 통합해 drag/resize
- object z-order controls
- multi-selection
- system clipboard interoperability
- accessible focus order 강화
- arbitrary existing PPTX unknown/unsupported part preservation
- Windows 실제 `BUILD_EXE.cmd` / `RUN_WRITER_SELFTEST.cmd` 검증

## 외부 PPTX 편집 안전 정책

임의의 외부 PPTX를 현재 모델로 축소해 읽은 뒤 그대로 덮어쓰면 SmartArt, animation, chart extension, unknown relationship 등이 손실될 수 있다.

따라서 unknown-part preservation이 준비되기 전까지는 `PptxEditableReader`가 안전하다고 판정한 파일만 Writer 경로로 저장한다.
