# UI Design Guide

PowerPointLite와 향후 멀티포맷 편집기 UI는 특정 상용 Office 제품의 화면을 복제하지 않고, 자체적인 경량 Windows 데스크톱 UI를 사용한다.

## 목표

- 기능을 찾기 쉬운 단순한 구조
- 작은 화면에서도 슬라이드/문서 본문 영역 확보
- 일관된 간격과 색상
- 이미지 자산 없이도 깔끔한 기본 UI
- 향후 Presentation / Document / Spreadsheet 모드를 같은 디자인 시스템으로 확장

## 현재 색상 체계

`src/UiTheme.cs`가 기본 palette를 정의한다.

```text
Window          #181A1F 계열
Toolbar         #1F2228 계열
Sidebar         #1B1E23 계열
Surface         #262A31 계열
Canvas          #0F1114 계열
Primary Text    밝은 회색
Secondary Text  중간 회색
Accent          독립적인 blue accent
```

Microsoft Office의 Ribbon 색상, 공식 아이콘, 앱 로고를 복제하지 않는다.

## 레이아웃 원칙

### Viewer

```text
┌─────────────────────────────────────────────┐
│ Primary toolbar                             │
├────────────┬────────────────────────────────┤
│ Thumbnails │                                │
│ / outline  │          Document canvas       │
│            │                                │
├────────────┴────────────────────────────────┤
│ status                                      │
└─────────────────────────────────────────────┘
```

- 좌측 pane이 표시되면 Fit은 우측 실제 viewport만 사용한다.
- auto-hide pane은 canvas 폭을 회수한다.
- 기본 viewer canvas는 콘텐츠에 집중하도록 어둡고 단순하게 유지한다.

### 향후 Editor

Toolbar를 무조건 Office Ribbon처럼 복제하지 않는다.

권장 구조:

```text
File | Home | Insert | View
────────────────────────────
작은 command groups / contextual properties
────────────────────────────
content canvas
```

탭 명칭처럼 일반적인 UI 개념은 사용할 수 있지만 Microsoft의 아이콘/배치/그래픽 자산을 그대로 복사하지 않는다.

## 버튼

- Flat button
- border 최소화
- hover/pressed 상태 명확화
- 텍스트만으로 의미가 분명하면 불필요한 아이콘을 넣지 않는다.
- 아이콘을 넣을 경우 자체 제작 SVG 또는 재배포 가능한 오픈 라이선스 아이콘만 사용한다.

## 폰트

UI 기본 글꼴은 Windows 시스템의 `Segoe UI`를 우선 사용하되 앱 패키지에 해당 폰트 파일을 포함하지 않는다.

문서 font picker는 UI font와 별도이며 `FontLicensing` 계층을 따른다.

## 접근성

- 텍스트와 배경 대비 확보
- 키보드 탐색 유지
- hover만으로 기능을 숨기지 않기
- 100%/125%/150% Windows scaling에서 레이아웃이 무너지지 않도록 확인
- 색상만으로 중요한 상태를 전달하지 않기

## 새 화면 체크리스트

- [ ] 기존 palette 사용
- [ ] Office/Hancom 공식 이미지 자산 미사용
- [ ] 125% DPI 확인
- [ ] 최소 창 크기 확인
- [ ] 키보드 동작 확인
- [ ] 긴 한국어/영어 라벨 잘림 확인
- [ ] Viewer Fit 영역 침범 여부 확인
