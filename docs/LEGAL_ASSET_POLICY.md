# Legal / Asset Policy

이 문서는 PowerPointLite 및 향후 멀티포맷 문서 편집기 확장 과정에서 지켜야 할 자산·상표·폰트 안전 원칙을 정의한다.

> 이 문서는 개발 정책이며 법률 자문이 아니다. 공개 배포나 상용 배포 전에 실제 라이선스 원문과 해당 권리자의 최신 정책을 다시 확인한다.

## 1. 기본 원칙

프로그램은 파일 형식을 독립적으로 구현한다.

허용 방향:

- 공개된 파일 형식 명세를 바탕으로 자체 Reader / Writer 구현
- OOXML (`.pptx`, `.docx`, `.xlsx`) 직접 읽기·생성·수정·저장
- 공식적으로 공개된 HWPX/HWP 명세 범위에서 자체 구현
- ODF (`.odt`, `.ods`, `.odp`) 표준 기반 자체 구현
- 사용자의 시스템에 설치된 폰트를 런타임에 열거하고 렌더링에 사용

금지 방향:

- Microsoft Office 실행 파일 또는 DLL을 프로젝트 배포물에 포함
- 한컴 실행 파일 또는 DLL을 프로젝트 배포물에 포함
- Office/Hancom UI 이미지, 공식 아이콘, 로고를 추출·복제해 사용
- 상용 템플릿, 클립아트, 유료 문서 샘플을 저장소에 복사
- 라이선스를 확인하지 않은 TTF/OTF 파일을 앱에 번들

## 2. 상표/브랜드

프로젝트는 Microsoft 또는 Hancom의 공식 제품처럼 보이게 만들지 않는다.

- 독립적인 앱 이름, 아이콘, 색상 체계를 사용한다.
- Word/Excel/PowerPoint/Hancom 로고를 앱 아이콘이나 브랜딩 요소로 사용하지 않는다.
- 호환성 설명이 필요할 때 제품명을 사실 설명 용도로만 사용한다.
- 공개 배포 전 README에 비제휴 문구를 유지한다.

권장 비제휴 문구:

```text
Microsoft, Word, Excel and PowerPoint are trademarks of Microsoft Corporation.
This project is not affiliated with or endorsed by Microsoft.

Hancom and related product names are trademarks of their respective owners.
This project is not affiliated with or endorsed by Hancom.
```

## 3. 폰트 정책

### 3.1 시스템 설치 폰트

기본 동작은 Windows에 이미 설치된 폰트 family를 열거하여 사용자가 선택하도록 한다.

- 앱 패키지에 시스템 폰트 파일을 복사하지 않는다.
- 문서에는 가능한 한 font family 이름만 기록한다.
- 상대방 PC에 폰트가 없으면 fallback font를 사용한다.

### 3.2 번들 폰트

폰트를 앱과 함께 배포할 때는 반드시 다음을 확인한다.

- 앱 embedding 허용 여부
- 파일 재배포 허용 여부
- 수정 허용 여부
- 문서/PDF embedding 허용 여부
- 라이선스 고지 의무

`무료`, `상업용 무료` 문구만으로 재배포 가능하다고 판단하지 않는다.

번들 폰트를 추가할 때는:

```text
fonts/
licenses/
THIRD_PARTY_NOTICES.md
```

에 출처와 라이선스를 같이 기록한다.

### 3.3 문서 Font Embedding

향후 문서에 폰트를 포함하는 기능은 기본 OFF로 한다.

Embedding 요청 시:

1. OpenType/TrueType `OS/2.fsType` 검사
2. 번들/사용자 제공 라이선스 메타데이터 검사
3. 실제 라이선스 문구가 있으면 그 내용을 우선
4. 불명확하면 embedding을 거부하거나 사용자에게 명시적으로 경고

`fsType`은 보조 정보일 뿐 최종 법적 판정기로 취급하지 않는다.

## 4. PDF Export

PDF에 font subset/full font를 넣는 것도 embedding이다.

따라서 PDF Export는 문서 embedding과 동일한 `FontLicenseService`를 사용해야 한다.

## 5. 테스트 자산

저장소의 테스트 문서는 직접 생성한 자산만 사용한다.

허용:

- 코드 generator가 만든 테스트 PPTX/DOCX/XLSX/HWPX
- 프로젝트 개발자가 직접 작성한 단순 샘플
- 권리와 라이선스가 명확한 자체 이미지

금지:

- 인터넷에서 받은 강의자료
- 상용 Office 템플릿
- 회사/학교/타인의 실제 문서
- 유료 스프레드시트/프레젠테이션 템플릿

## 6. 서드파티 코드/라이브러리

새 라이브러리를 추가하기 전에 최소한 다음을 기록한다.

- 패키지명/프로젝트명
- 버전
- 소스 URL
- 라이선스
- 앱 배포 시 고지/소스 공개 의무

라이선스가 불명확하면 의존성을 추가하지 않는다.

## 7. 개발 체크

새 기능/자산을 추가할 때 다음 질문 중 하나라도 `모름`이면 merge하지 않는다.

- 이 코드는 직접 작성했거나 사용 허가가 명확한가?
- 이 이미지/아이콘은 직접 제작했거나 라이선스가 명확한가?
- 이 폰트 파일을 앱과 같이 재배포해도 되는가?
- 이 폰트를 문서/PDF에 embedding해도 되는가?
- 이 테스트 문서는 저장소에 배포할 권리가 있는가?
- 이 구현은 공개 명세 또는 독립 코드에 기반하는가?
