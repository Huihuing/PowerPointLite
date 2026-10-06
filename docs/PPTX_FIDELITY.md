# PPTX Rendering Fidelity

이 문서는 PowerPointLite의 **자체 Internal OpenXML renderer**가 어떤 PPTX 기능을 직접 해석하고, 어떤 기능을 근사하며, 어떤 경우 외부 renderer가 더 높은 표현력을 갖는지 기록한다.

우선순위는 다음과 같다.

```text
Microsoft PowerPoint (설치된 경우)
→ LibreOffice Impress (설치된 경우)
→ Internal OpenXML renderer
```

따라서 아래 내용은 PowerPoint가 설치되지 않은 포터블 환경에서 사용하는 자체 renderer의 범위다.

## 현재 직접 처리하는 영역

- slide / layout / master 시각 레이어
- theme color 및 major/minor font
- placeholder 위치와 text-style 상속
- rich text run / paragraph
  - font family / size
  - bold / italic / underline
  - color
  - baseline
  - paragraph alignment
  - bullet
  - before/after spacing
  - body margins / wrap / vertical anchor / vertical text 일부
  - normAutofit fontScale + overflow 추가 축소 근사
  - spAutoFit 텍스트 영역 확장 근사 / noAutofit
  - defTabSz / tabLst 기본 탭 정지점
  - 간단 RTL paragraph 배치
- raster image 및 crop / rotate / flip / alpha
- grayscale / bi-level / brightness / contrast 그림 효과의 GDI+ 근사
- PNG / JPEG / GIF / BMP
- EMF / WMF는 GDI+가 읽을 수 있는 경우
- SVG primitive와 path
  - rect / circle / ellipse / line / polygon / polyline / text 혼합 렌더
  - M / L / H / V / C / S / Q / T / A / Z
  - absolute / relative path
  - linearGradient / radialGradient 기본 stop interpolation
  - clipPath 기본 geometry
  - translate / scale / rotate 기본 transform
- shape
  - 기본 Office preset
  - custom geometry의 line / cubic / quadratic path
  - 추가 preset: pentagon, octagon, star5/star6, plus, chevron, homePlate, trapezoid, 방향 화살표
- gradient / picture fill / line / dash / arrowhead / basic shadow
- group coordinate transform과 중첩 group
- group rotate / horizontal flip / vertical flip의 graphics transform 근사
- table
  - row/column size
  - cell fill/border/text
  - gridSpan/rowSpan 기반 merge 일부
- charts
  - column / bar / line / pie
  - doughnut
  - area
  - scatter
  - bubble
  - radar
  - cached category/value 기반 title/legend
  - 음수 값과 0 기준선
  - valAx min/max/majorUnit/orientation 기본 반영
  - showVal/showCatName/showSerName/showPercent 데이터 레이블
  - 데이터 레이블 separator 기본 반영
  - legendPos 좌/우/상/하/우상단 배치 및 plot 영역 조정
  - 파이/도넛 category 범례 및 percent 레이블
- SmartArt / diagram
  - data model의 node / connection을 읽어 hierarchy로 표시
  - 구조를 알 수 없는 경우 기존 static approximation으로 fallback
- hyperlink / slide navigation / media extraction
- speaker notes
- hidden slide
- transition
  - cut / fade / wipe / push 계열
  - wipe / push는 좌·우·상·하 `dir` 방향을 반영
  - `spd` 및 명시적 `dur` 시간 정보를 가능한 범위에서 반영
  - unsupported transition은 fade fallback
  - Morph는 현재 smooth-fade 계열 fallback
- `mc:AlternateContent`
  - Internal renderer가 직접 이해하지 못하는 신형 개체는 호환 `Fallback`을 임시 package에서 선택해 렌더링
  - 사용자 원본 PPTX/PPTM은 수정하지 않음

## Animation timing

`AnimationTimelineRenderer`는 `p:timing` 트리를 읽어 내부 슬라이드쇼용 timeline으로 변환한다.

현재 보존/해석하는 정보:

- target shape id (`spTgt`)
- entrance / exit / emphasis / motion 분류
- `presetClass`, `presetID`, `presetSubtype`
- on-click
- with-previous
- after-previous
- start delay
- duration
- speed modifier
- repeat count
- auto reverse
- slide `advTm`
- slide `advClick`
- slideshow `loop`
- slideshow `useTimings`

실행 방식:

```text
slide load
→ entrance 대상의 초기 visibility 계산
→ click/automatic animation group 순서대로 진행
→ with-previous는 같은 group
→ after-previous는 앞 action 종료 시점 뒤에 배치
→ delay/duration을 내부 timer에 반영
→ entrance/exit는 단계별 render state에 반영
→ emphasis/motion은 현재 상태를 보존하면서 시각 근사
→ slide advTm은 animation click을 소비하지 않고 실제 다음 slide로 이동
```

`RUN_ANIMATION_SELFTEST.cmd`는 프로젝트가 자체 생성한 PPTX에 합성 `p:timing` 트리를 넣고 다음을 검사한다.

- click entrance
- with-previous emphasis
- after-previous exit
- delayed second click motion
- timeline grouping
- delay/duration retention
- 단계 0/1/2 render state 생성

외부 저작물은 fixture로 사용하지 않는다.

## 의도적인 근사 영역

Internal renderer는 Microsoft PowerPoint의 렌더링 엔진을 복제한 것이 아니다. 다음 항목은 구조와 의미를 가능한 범위에서 유지하되 화면 결과가 PowerPoint와 정확히 같지 않을 수 있다.

- SmartArt 고유 layout algorithm
- chart의 모든 axis/style/data-label/3D 조합
- theme effect style 전체
- SVG filter/mask/pattern 및 복잡한 gradientTransform/clipPathUnits 등 고급 SVG 기능
- image artistic effects
- 3D shape / bevel / material / lighting
- complex custom geometry formula/adjust handle
- animation easing / acceleration / deceleration curve
- exact motion-path interpolation
- scale / rotation / color emphasis의 프레임 단위 변화
- trigger-on-specific-object, media bookmark 등 고급 trigger
- nested sequence/parallel timing tree의 모든 PowerPoint edge case
- Morph object matching

이 경우 구조를 버리는 것보다 **보수적인 fallback 또는 시각 근사**를 선택한다.

## Editor와의 관계

Editor interactive canvas는 선택/drag/resize를 위한 빠른 편집 뷰다. 최종 시각 fidelity 기준은 `InternalPptxRenderer`다.

Editor의 `미리보기 / Preview`는 현재 편집 모델을 임시 PPTX로 저장한 뒤 Internal renderer를 사용하므로, 저장 결과에 가까운 모습을 확인하는 용도다.

`PresentationTextBox`는 plain text 속성과 함께 선택적으로 paragraph/run 모델을 가진다. 기존 안전 편집 대상으로 판정된 문서에서는 rich text run을 읽고 다시 저장할 수 있다. 사용자가 plain text 자체를 수정하면 stale run 정보를 억지로 재적용하지 않고 해당 textbox의 run-level 서식을 안전하게 평문화한다.

## 문서 안전성

렌더링 지원과 편집 안전성은 별개다.

Viewer가 chart, SmartArt, animation 등을 화면에 표시할 수 있다고 해서 해당 외부 PPTX를 현재 Editor로 안전하게 round-trip할 수 있다는 뜻은 아니다.

`PptxEditableReader`의 deny-by-default 정책은 유지한다. 현재 편집 모델이 package의 unsupported/unknown content를 손실 없이 보존한다고 확인되지 않으면 원본 파일을 Editor 저장 경로에 넣지 않는다.

## 검증 게이트

Windows에서 merge 전 다음을 실행한다.

```bat
RUN_SOURCE_AUDIT.cmd
BUILD_EXE.cmd
RUN_PREMERGE_CHECKS.cmd
```

특히 PPTX 변경 후에는 다음 수동 비교가 필요하다.

- Internal renderer와 Microsoft PowerPoint의 실제 PPTX 화면 비교
- LibreOffice Impress와의 비교
- click / with previous / after previous가 섞인 slideshow
- delay / duration / repeat / auto reverse
- slide auto advance (`advTm`)
- loop slideshow
- grouped/rotated objects
- SVG / chart / SmartArt가 포함된 rights-cleared test document

현재 비-Windows 개발 환경에서는 `csc.exe` 빌드 성공을 주장하지 않는다.
