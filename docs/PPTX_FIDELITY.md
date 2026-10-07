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
  - tint/shade
  - hue/hueOff/hueMod
  - sat/satOff/satMod
  - lum/lumOff/lumMod
  - comp/inv/gray
  - gamma/invGamma 기본 변환
- placeholder 위치와 text-style 상속
- rich text run / paragraph
  - font family / size
  - bold / italic / underline
  - color
  - baseline
  - paragraph alignment
  - bullet
  - before/after spacing
  - placeholder bodyPr margin/anchor/autofit inheritance
  - body margins / wrap / vertical anchor / vertical text 일부
  - normAutofit fontScale + overflow 추가 축소 근사
  - spAutoFit 텍스트 영역 확장 근사 / noAutofit
  - defTabSz / tabLst 탭 정지점
  - tab alignment left / center / right / decimal 기본 근사
  - 간단 RTL paragraph 배치
- raster image 및 crop / rotate / flip / alpha
- grayscale / bi-level / brightness / contrast 그림 효과의 GDI+ 근사
- DrawingML picture blur의 downsample/upscale 기반 기본 근사
- Office 2010 `a14:sharpenSoften`의 3×3 sharpen/soften 기본 근사
- Office 2010 `a14:artisticBlur` radius 기반 기본 근사
- duotone luminance 2색 매핑
- clrChange from/to 색상 remap
- PNG / JPEG / GIF / BMP
- EMF / WMF는 GDI+가 읽을 수 있는 경우
- SVG primitive와 path
  - rect / circle / ellipse / line / polygon / polyline / text 혼합 렌더
  - use href/xlink:href의 기본 geometry 참조 + x/y 이동
  - symbol/g/svg 컨테이너 참조 및 중첩 use 기본 path 펼침, 순환 참조 차단
  - symbol/svg viewBox + use width/height viewport 스케일, preserveAspectRatio none/meet/slice 기본 정렬
  - direct primitive use 참조의 명시적 fill/stroke/presentation style fallback
  - symbol/g/svg 참조 루트의 명시적 presentation style fallback, nested use 순환 차단
  - 단순 symbol/g use 참조는 subtree를 자식별 렌더링하여 서로 다른 solid fill/stroke style 보존
  - fill/stroke의 currentColor를 상속 color 속성으로 해석
  - presentation attribute/style의 inherit 키워드에서 부모 값 계속 탐색
  - SVG hex 색상 #RGB/#RGBA/#RRGGBB/#RRGGBBAA
  - rgb()/rgba() 절대 색상: 숫자·퍼센트 채널과 alpha 기본 지원
  - hsl()/hsla() 절대 색상: deg/rad/grad/turn hue와 alpha 기본 지원
  - M / L / H / V / C / S / Q / T / A / Z
  - absolute / relative path
  - linearGradient / radialGradient 기본 stop interpolation
  - clipPath 기본 geometry
  - translate / scale / rotate / matrix / skewX / skewY transform
  - gradientTransform의 matrix/translate/scale/rotate 기본 반영
  - spreadMethod pad/repeat/reflect 기본 근사
  - pattern 기본 반복 fill 및 patternTransform
  - mask의 단순 geometry/luminance clip 근사
  - clipPathUnits="objectBoundingBox" 기본 반영
  - clipPath/mask 내부 nested group/geometry transform 누적 반영
  - stroke-dasharray / stroke-dashoffset 기본 점선 패턴 반영
  - stroke-miterlimit 및 SVG 기본값 4 반영
  - fill-rule nonzero(기본) / evenodd 복합 path 채움 판정
  - clipPath 내부 clip-rule nonzero(기본) / evenodd 기본 반영
  - filter="url(#...)"의 단일 feGaussianBlur(SourceGraphic) stdDeviation 기본 근사
  - 단일 feDropShadow(SourceGraphic)의 dx/dy/stdDeviation/flood-color/flood-opacity 기본 근사
  - 단일 feOffset(SourceGraphic)의 dx/dy 이동 기본 근사
  - feOffset → feGaussianBlur → feBlend(normal/multiply/screen, SourceGraphic) 3단계 기본 합성
  - feOffset → feGaussianBlur → feComposite(over/in/out/xor/atop, SourceGraphic) 3단계 기본 합성
  - feComposite arithmetic(k1~k4)는 SourceGraphic + offset/blur chain 입력 순서에서 두 입력 union 영역 채널 수식 기본 근사
  - feOffset ↔ feGaussianBlur 2단계 SourceGraphic 체인의 result/in 연결과 최종 shifted-blur 근사
  - 단일 feColorMatrix(SourceGraphic)의 matrix 4×5 / saturate / hueRotate / luminanceToAlpha 색 변환(solid fill/stroke)
- shape
  - 기본 Office preset
  - custom geometry의 line / cubic / quadratic / arcTo path
  - avLst/gdLst guide formula 기본 계산
  - val, */, +-, +/, ?:, abs, sqrt, max/min, at2, sin/cos/tan, mod, pin, cat2/sat2
  - w/h/hc/vc/ss/ls 및 주요 cd/분할 built-in guide
  - 추가 preset: pentagon, octagon, star5/star6, plus, chevron, homePlate, trapezoid, 방향 화살표
- gradient / picture fill / line / dash / arrowhead
- outer shadow / preset shadow 기본 offset 근사
- inner shadow의 clipped edge 근사
- glow 다단계 outline 근사
- softEdge 외곽 fade 근사
- reflection mirrored alpha-gradient 근사
- sp3d bevelT/bevelB의 highlight/shade 근사
- sp3d extrusionH depth layer + scene3d camera/lightRig 방향 근사
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
  - column/bar/line/area의 catAx crossesAt 및 crosses=min/max 축 교차 위치 기본 반영
  - showVal/showCatName/showSerName/showPercent 데이터 레이블
  - dLbls numFmt 숫자 서식 기본 반영
  - 데이터 레이블 separator 기본 반영
  - dLblPos의 ctr/inBase/inEnd/outEnd 및 t/b/l/r 위치 기본 근사
  - pie/doughnut showLeaderLines 및 leaderLines spPr/ln solidFill·width·prstDash 기본 반영
  - dLbls/dLbl[idx] txPr의 solidFill 텍스트 색상 및 sz/b/i/latin typeface 기본 반영
  - dLbls/dLbl[idx] spPr의 solidFill 배경과 ln 색상·굵기·dash 박스 스타일 기본 반영
  - series dLbls/dLbl[idx]의 표시 항목·위치·numFmt·separator·delete point override 기본 반영
  - 개별 dLbl/layout/manualLayout의 x/y/w/h 및 factor/edge mode를 chart 전체 좌표 기준으로 point label box에 반영
  - chart-level dLbls와 series-level dLbls를 별도 scope로 해석
  - legendPos 좌/우/상/하/우상단 배치 및 plot 영역 조정
  - legend txPr의 solidFill·sz·b/i·latin typeface 텍스트 스타일 기본 반영
  - legendEntry/idx/delete 기반 개별 legend 항목 숨김 반영
  - chart title/legend overlay=1일 때 plot 영역을 예약하지 않고 겹쳐 배치
  - plotArea/layout/manualLayout의 x/y/w/h 및 factor/edge mode를 일반/확장 차트 plot 사각형에 기본 반영
  - plotArea manualLayout의 layoutTarget=inner/outer를 구분하고, outer에서는 축·축 레이블용 자동 inset 비율을 보존해 내부 plot을 근사
  - plotArea manualLayout은 direct-child layout scope만 사용해 dLbl 등 중첩 manualLayout과 분리
  - x/y/w/h 허용 범위를 벗어난 manualLayout은 기본 배치로 fallback하고, 유효한 배치가 차트 경계를 넘으면 크기를 유지한 채 안쪽으로 이동
  - chart title/legend layout/manualLayout의 x/y/w/h 및 factor/edge mode를 기본 위치·크기에 반영
  - chart-level title/legend를 axis title과 별도 scope로 해석해 축 제목을 차트 제목으로 오인하지 않도록 처리
  - chart title txPr의 solidFill·sz·b/i·latin typeface 텍스트 스타일 기본 반영
  - catAx/valAx title txPr의 solidFill·sz·b/i·latin typeface 축 제목 스타일 기본 반영
  - 파이/도넛 category 범례 및 percent 레이블
  - scatter/bubble의 series·value·point override data label 기본 렌더링
  - bubble data label의 showBubbleSize 및 point별 showBubbleSize override 반영
  - bubble chart도 series legend를 공통 legend renderer로 표시
  - c:ser/c:spPr series 색상 및 c:dPt point 색상 우선 반영
  - chartSpace / plotArea spPr solidFill 배경색 기본 반영
  - valAx/catAx 축 제목 기본 렌더
  - catAx/valAx spPr/ln solidFill·width·prstDash 축선 스타일 기본 반영
  - catAx/valAx delete=1이면 축선·tick·tick label·축 제목을 숨기고 value gridline 계산은 유지
  - catAx/valAx axPos를 nextTo/미지정 tick label 기본 방향과 value-axis 기본 edge에 반영
  - catAx/valAx txPr의 solidFill·sz·b/i·latin typeface tick label 스타일 기본 반영
  - catAx majorTickMark 및 valAx majorTickMark/minorTickMark의 in/out/cross 축 눈금 기본 반영
  - valAx majorGridlines spPr/ln solidFill·width·prstDash 스타일 기본 반영
  - valAx minorUnit 및 minorGridlines spPr/ln solidFill·width·prstDash 기본 반영
  - catAx tickLblPos=none 카테고리 축 레이블 숨김
  - catAx tickLblPos=high: column/area 위쪽, bar 오른쪽 레이블 배치
  - catAx tickLblSkip 기반 카테고리 레이블 표시 간격 기본 반영
  - valAx tickLblPos=none 값 축 숫자 레이블 숨김 (gridline 유지)
  - valAx tickLblPos=high: column/area 오른쪽, bar 위쪽 숫자 레이블 배치
  - valAx numFmt의 기본 소수/천단위/퍼센트 표시
  - 같은 chart kind 내부의 복수 chart-group axId를 series별 category/value axis binding으로 해석
  - primary/secondary valAx를 분리해 series별 scale을 계산하고 secondary axPos edge에 별도 axis/tick label 렌더링
  - secondary valAx에 명시된 major/minor gridlines의 선 색상·굵기·dash를 별도 scale 위치에 렌더링
  - secondary valAx title과 기본 txPr 글꼴/색상 스타일을 축 위치에 맞춰 렌더링
  - line+column mixed-kind combo chart에서 각 series의 chart-group kind를 유지해 line은 선/marker, column은 bar primitive로 함께 렌더링
  - mixed-kind series도 value-axis binding 기준으로 primary/secondary scale을 분리해 secondary column/line 값을 독립 축 범위로 처리
  - line chart series 선 굵기/prstDash 및 circle/square/diamond/triangle/x/plus marker
  - series trendline의 linear / poly(order 2~6) / movingAvg(period) 타입과 spPr/ln 색상·굵기·dash 기본 렌더링
  - linear/poly trendline의 forward/backward 기본 연장, poly는 정규화 좌표 최소제곱 fitting, movingAvg는 period 구간 평균 연결
  - series errBars의 fixedVal / percentage / cust를 지원하고 plus/minus/both 방향 기본 렌더링
  - custom errBars의 c:plus/c:minus numLit 또는 numRef>numCache point별 값을 분리해 비대칭 오차폭 반영
  - error bar는 line/column/bar와 area·scatter Y 방향에서 공통 렌더링
  - 지원하지 않는 trendline/error-bar 타입은 기존 series 렌더를 보존하고 안전하게 무시
  - bar/column grouping: clustered / stacked / percentStacked
  - gapWidth / overlap 기반 bar thickness·series overlap 근사
  - pie/doughnut firstSliceAng
  - doughnut holeSize
  - doughnut / area / radar도 공통 series 색상·point 색상·line dash·marker 스타일 재사용
- SmartArt / diagram
  - data model의 node / connection 관계 해석
  - layout definition 기반 hierarchy / process / vertical process / cycle / radial / matrix / pyramid / list / venn 배치 근사
  - layout별 connector 방향과 node shape 기본 차등 표현
  - hierarchy의 type="asst" assistant node 보존, 부모 좌우 보조 박스 배치 및 side connector 근사
  - 한쪽 assistant 공간이 부족해 반대편으로 재배치되는 경우 실제 좌/우 배정 개수 기준으로 stack index와 전체 높이를 다시 계산해 겹침을 방지
  - 동일 부모의 다중 assistant는 좌우로 번갈아 분배하고 같은 쪽 stack을 수직 중앙 정렬·비겹침 배치
  - hierarchy 일반 자식을 부모 중심별 그룹으로 묶어 sibling spacing과 행 배치를 근사
  - hierarchy parent-child 및 assistant connector를 직각 elbow routing으로 근사
  - process 같은 행은 node edge 연결, 행 전환 및 verticalProcess는 직각 elbow routing으로 근사
  - hierarchy/process/verticalProcess elbow의 전체 V-H-V 경로를 장애물과 대조하고, 중앙/분할 수평 channel이 막히면 H-V-H side-channel로 전환
  - 복잡 hierarchy에서는 장애물 top/bottom/left/right 인접 channel까지 후보로 확장하고, 모든 clear route 중 최단 거리 + 중앙 근접 점수로 deterministic 경로 선택
  - hierarchy assistant connector도 같은 full-route 검사기를 사용해 중앙 세로 channel이 막히면 다른 side-channel로 우회
  - 같은 SmartArt에서 먼저 확정된 elbow connector route와 다음 후보의 내부 교차 수를 계산하고 큰 penalty를 부여해 가능한 경우 connector crossing을 줄임
  - 같은 방향 trunk 공유는 hierarchy 의도 가능성을 고려해 crossing penalty 대상에서 제외
  - process multi-row 배치는 행별 실제 노드 수 기준 중앙 정렬 및 안정적인 row spacing 적용
  - 좁은 process 영역에서는 가용 폭 기준으로 열 수를 자동 축소해 node가 좌우 경계를 벗어나지 않도록 배치
  - verticalProcess는 높이 cap 이후 전체 node stack을 영역 중앙에 재배치
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
- on-click 조건의 특정 trigger shape id 보존 (`cond/tgtEl/spTgt`)
- with-previous
- after-previous
- start delay
- duration
- speed modifier
- repeat count
- auto reverse
- acceleration / deceleration easing
- animClr RGB color emphasis
- slide `advTm`
- slide `advClick`
- slideshow `loop`
- slideshow `useTimings`

실행 방식:

```text
slide load
→ entrance 대상의 초기 visibility 계산
→ click/automatic animation group 순서대로 진행
→ 특정 trigger shape id가 있는 click step은 해당 shape bounds를 클릭했을 때만 실행
→ PictureBox Zoom 레터박스를 제외한 실제 슬라이드 표시 영역 기준으로 클릭 좌표 정규화
→ xfrm rot가 있는 trigger shape는 중심 기준으로 클릭 좌표를 역회전한 뒤 해당 geometry 내부를 판정
→ prstGeom=ellipse trigger는 타원 방정식으로 실제 타원 내부만 hit
→ prstGeom=roundRect는 기본 corner radius 근사로 둥근 모서리 바깥 클릭 제외
→ prstGeom=diamond는 정규화 마름모 영역, triangle은 꼭짓점-밑변 선형 영역으로 정밀 hit-test
→ with-previous는 같은 group
→ after-previous는 앞 action 종료 시점 뒤에 배치
→ delay/duration을 내부 timer에 반영
→ entrance/exit는 단계별 render state에 반영
→ emphasis/motion은 현재 상태를 보존하면서 개체 레이어 단위로 시각 근사
→ motion path는 M/L/H/V/C/S/Q/T/A path를 flatten하여 경로 길이 기준으로 보간
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

- SmartArt 고유 layout algorithm의 정밀 spacing, 다중 assistant 배치 우선순위 및 공식 connector routing semantics 전체
- chart의 모든 axis/style/data-label/3D 조합
- mixed-kind combo chart의 복수 독립 category-axis 위치/레이블 parity 및 secondary axis의 crosses/crossesAt·tick style 세부 parity
- theme effect style 참조 전체 및 복합 shadow/effect 조합
- feComposite arithmetic의 filter-region 전체 k4 확장, 임의 입력 순서/복합 chain 및 feColorMatrix가 다른 primitive와 연결된 복잡한 SVG filter chain, 복잡한 mask luminance/gradient 등 고급 SVG 기능
- SVG use subtree의 복잡한 CSS cascade/selector, viewport가 있는 symbol의 자식별 확장, root filter/clip/mask 조합 전체
- artisticBlur 외의 artistic preset 등 고급 image artistic effects
- 3D material/lighting, camera projection 및 bevel/extrusion profile의 정확한 PowerPoint parity
- custom geometry의 복잡한 nested guide 의존성, 모든 preset formula/adjust-handle semantics
- exact PowerPoint easing curve parity
- PowerPoint의 origin/pathEditMode 등 세부 motion-path 좌표계 semantics
- scheme/HSL color animation과 복합 color transform
- 자유형/custom geometry 및 기타 preset path 자체의 정밀 trigger hit-test, media bookmark 등 고급 trigger
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
