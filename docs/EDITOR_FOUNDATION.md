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

`AdvancedEditorFidelityExtension`은 Editor 상단에 `미리보기 / Preview`를 추가한다.

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

PPTX 시각 표현과 animation 범위는 `docs/PPTX_FIDELITY.md`를 함께 본다.

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
     │   └─ PresentationTextParagraph[] (optional rich text)
     │       └─ PresentationTextRun[]
     ├─ PresentationShape[]
     ├─ PresentationImage[]
     └─ PresentationTable[]
         └─ PresentationTableCell[]
```

`PresentationTextRun`은 font family/size, bold, italic, underline, color, baseline을 보존한다.

`PresentationTextParagraph`는 alignment, level, bullet, before/after spacing과 run 목록을 가진다.

기존 단순 text box와의 호환을 위해 `PresentationTextBox.Text`와 box-level font 속성은 계속 존재한다. 안전 편집 대상으로 읽은 문서의 rich run은 `PptxRichTextPackage`가 추가로 읽고 Writer 저장 뒤 다시 OOXML paragraph/run으로 주입한다.

사용자가 기존 plain-text 편집 UI에서 문자열 자체를 바꾸면 이전 run 경계를 새 문자열에 억지로 적용하지 않는다. 이 경우 해당 textbox의 rich run을 비우고 box-level format으로 평문화해 **stale formatting이 다른 글자에 잘못 붙는 문제를 피한다.**

현재 property panel에서 전체 font family/size/bold/italic을 바꾸면 rich text가 있을 때 모든 run에 해당 속성을 적용하되 underline/baseline/color/bullet/paragraph spacing은 유지한다.

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
 ├─ PptxRichTextPackage
 └─ PptxTableWriter
 ↓
staged package
 ↓
backup/replace
 ↓
<destination>.pptx
```

Writer 중간 실패 시 stage 파일을 제거하고 기존 destination을 가능한 한 복원하도록 구성한다.

## Writer / animation 자체 테스트

Windows 실제 빌드 후:

```bat
RUN_WRITER_SELFTEST.cmd
RUN_ANIMATION_SELFTEST.cmd
```

Writer 자체 테스트는 프로젝트 코드만으로 text box, shape, 자체 생성 image, table이 포함된 PPTX를 만들고 read-edit-write round-trip을 검사한다. text box/shape/image는 `PresentationSlide.ObjectOrder`에 공통 layer 순서를 기록하며, Writer와 editable reader가 PPTX `spTree` 순서를 왕복 보존하는지도 함께 검증한다.

Advanced Editor는 선택 객체에 대해 **Send Back / Bring Front**를 제공한다. 이 동작은 단순 화면 표시 순서가 아니라 모델의 공통 layer order를 변경하며, Canvas paint/hit-test와 Writer 저장에 동일하게 적용된다. `Ctrl+Shift+Down` / `Ctrl+Shift+Up`으로도 실행할 수 있다.

또한 **Ctrl+클릭 multi-selection**을 지원한다. 여러 text box/shape/image를 함께 선택하면 각 객체 선택선과 전체 group bounds가 표시되고, 마우스 drag 또는 방향키로 상대 간격을 유지한 채 동시에 이동할 수 있다. 그룹이 slide 경계에 닿으면 전체 그룹 기준으로 이동량을 제한하며, Delete는 선택된 객체를 종류별 역순으로 제거해 index 변화에도 안전하게 동작한다. 다중 선택 중에는 개별 속성 편집·복사·z-order 버튼을 잠가 잘못된 단일-object 동작을 방지한다.

multi-selection 상태에서는 상단 **Align** 메뉴가 활성화된다. 2개 이상 객체에 대해 left/center/right/top/middle/bottom 정렬을 제공하고, 3개 이상에서는 horizontal/vertical distribute를 추가로 활성화한다. distribute는 양 끝 객체의 center를 고정한 채 중간 객체의 center 간격을 균등하게 재배치한다.

복사/붙여넣기는 단일 객체뿐 아니라 **현재 multi-selection 전체**를 처리한다. 선택 객체는 원본 `ObjectOrder` 순서대로 PowerPointLite 전용 versioned Clipboard payload에 기록되고, 상대 좌표·서식·이미지 bytes·상대 z-order를 유지한 채 현재 슬라이드에 붙여넣는다. 그룹이 슬라이드 경계를 벗어날 때는 객체별로 잘라 맞추지 않고 그룹 bounds에 하나의 translation을 적용해 내부 간격을 보존한다.

Windows system clipboard에도 `PowerPointLite.ObjectSelection.v1` 포맷을 함께 기록하므로 다른 PowerPointLite 창/프로세스에서도 객체 묶음을 붙여넣을 수 있다. 단일 text box는 Unicode text도 함께 내보내고 단일 image는 Bitmap도 함께 내보낸다. 반대로 외부 프로그램에서 복사한 일반 text와 Bitmap image도 각각 새 text box / image 객체로 받아들인다. 전용 payload codec은 이미지 크기와 object/run count에 상한을 두고 structural self-test에서 encode/decode와 relative z-order round-trip을 검증한다.

Animation 자체 테스트는 합성 `p:timing`을 프로젝트가 만든 PPTX에 주입해 click entrance, with-previous emphasis, after-previous exit, delayed motion step과 단계별 render state를 검사한다.

인터넷 문서나 Microsoft/Hancom 템플릿을 fixture로 사용하지 않는다.

## 현재 남은 Editor 작업

- rich text를 선택 영역 단위로 직접 편집하는 UI
- Viewer renderer와 interactive canvas의 공통 layout/render primitive 확대
- chart/SmartArt/media 등 고급 요소는 Viewer fidelity를 보존하면서 단계적 편집 지원
- table을 일반 object selection과 통합해 drag/resize
- accessible focus order 강화
- arbitrary existing PPTX unknown/unsupported part preservation
- Windows 실제 `BUILD_EXE.cmd` / `RUN_PREMERGE_CHECKS.cmd` 검증

## 외부 PPTX 편집 안전 정책

임의의 외부 PPTX를 현재 모델로 축소해 읽은 뒤 그대로 덮어쓰면 SmartArt, animation, chart extension, unknown relationship 등이 손실될 수 있다.

따라서 unknown-part preservation이 준비되기 전까지는 `PptxEditableReader`가 안전하다고 판정한 파일만 Writer 경로로 저장한다.

Viewer가 고급 chart/SmartArt/animation을 표시할 수 있다는 사실은 해당 요소를 Editor가 lossless하게 수정할 수 있다는 의미가 아니다.
