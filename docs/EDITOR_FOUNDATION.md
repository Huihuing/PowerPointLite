# Presentation Editor Foundation

`feature/office-foundation`의 초기 PPTX Editor는 기존 Viewer를 대체하지 않고 별도 창으로 동작한다.

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

## 현재 Editor UI

- 독립 dark application chrome
- 왼쪽 slide list
- 가운데 16:9 canvas
- 오른쪽 text properties
- slide 추가/삭제/순서 이동
- text 수정
- Windows 설치 font family 선택
- font size / bold / italic
- left / center / right alignment
- text color
- Save / Save As
- dirty-state 표시
- 닫을 때 저장 확인

Microsoft Office/Hancom Ribbon, 공식 아이콘, 이미지 자산을 복제하지 않는다.

## 내부 모델

```text
PresentationDocument
 └─ PresentationSlide[]
     └─ PresentationTextBox[]
```

`PresentationTextBox`는 현재 다음 정보를 가진다.

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

폰트는 **family 이름만 문서에 기록**하며 TTF/OTF를 자동 embedding하지 않는다.

## Save 구조

```text
Editor
 ↓
PresentationEditSession
 ↓
PresentationDocument
 ↓
PptxWriter
 ↓
<destination>.writing
 ↓ successful package close
atomic-style replace
 ↓
<destination>.pptx
```

Writer 중간 실패 시 가능한 한 `.writing` 파일을 제거하고 기존 destination을 건드리지 않는다.

현재는 destination 교체 전에 기존 파일을 삭제하고 `File.Move`하는 방식이므로, 향후 backup/replace 전략을 추가해 crash-safe 저장을 더 강화한다.

## Writer 자체 테스트

Windows 실제 빌드 후:

```bat
RUN_WRITER_SELFTEST.cmd
```

테스트 내용:

1. 자체 생성 3-slide PPTX 저장
2. 필수 OPC part 존재 검사
3. `presentation.xml` slide count 검사
4. `PptxEditableReader`로 다시 읽기
5. 텍스트 수정
6. 다시 PPTX 저장
7. 두 번째 Reader로 수정 내용 검증

인터넷 문서나 Microsoft/Hancom 템플릿을 fixture로 사용하지 않는다.

## 다음 Editor 작업

- Add Text / Delete Object
- object drag/resize
- undo/redo command stack
- copy/paste
- image insert
- basic shapes
- slide thumbnail preview
- keyboard object navigation
- accessible focus order
- existing arbitrary PPTX unknown-part preservation

임의의 외부 PPTX를 Writer로 다시 저장하는 기능은 unknown-part preservation이 준비되기 전까지 활성화하지 않는다.
