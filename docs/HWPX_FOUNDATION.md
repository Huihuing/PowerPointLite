# HWPX Foundation

`feature/office-foundation`의 HWPX 구현은 한글과컴퓨터 바이너리나 DLL을 포함하지 않고, 공개된 HWPX/OWPML 구조를 바탕으로 독립 작성한다.

## 현재 상태

실험적 기반:

- `src/HwpxCore.cs`
  - `HwpxReader`
  - `HwpxWriter`
  - `HwpxEditSafety`
- `src/HwpxDiagnostics.cs`
- `RUN_HWPX_SELFTEST.cmd`
- `PowerPointLite.exe --hwpx-selftest <path>`

현재 구현은 **구조 검증 단계**이며 Windows 한컴 실제 열기 테스트 전에는 안정 기능으로 간주하지 않는다.

## 공개 규격 기반 패키지 구조

현재 writer는 다음 최소 package를 직접 생성한다.

```text
mimetype
version.xml
META-INF/
  container.xml
  manifest.xml
Contents/
  content.hpf
  header.xml
  section0.xml
Preview/
  PrvText.txt
```

`mimetype` 값:

```text
application/hwp+zip
```

`META-INF/container.xml`의 HWPX root package는:

```text
Contents/content.hpf
```

이며 media type은:

```text
application/hwpml-package+xml
```

이다.

## Reader 흐름

```text
HWPX ZIP
 ↓
mimetype 확인
 ↓
META-INF/container.xml
 ↓
Contents/content.hpf
 ↓
manifest / spine
 ↓
Contents/header.xml
 ↓
Contents/sectionN.xml
 ↓
TextDocument
```

현재 reader는 다음을 내부 `TextDocument`로 가져온다.

- paragraph
- run text
- tab / line break
- paragraph alignment
- font family reference
- font size
- bold
- italic
- underline
- text color

## Writer 흐름

```text
TextDocument
 ↓
font/style catalog
 ↓
header.xml
 ↓
section0.xml
 ↓
content.hpf / container / manifest
 ↓
HWPX ZIP
```

시스템에 설치된 font family 이름만 문서에 기록한다. TTF/OTF 파일을 자동으로 HWPX에 포함하지 않는다.

## 편집 안전 정책

`HwpxEditSafety`는 deny-by-default다.

현재 직접 편집 가능한 것으로 취급하려면:

1. `mimetype`이 HWPX인지 확인
2. 현재 프로젝트 writer가 만든 package인지 확인
3. 지원 목록 밖의 package part가 없는지 확인
4. encryption metadata가 없는지 확인

조건을 만족하지 않는 HWPX는 원본을 다시 쓰지 않는다.

향후 unknown-part preservation이 완성되기 전에는 임의의 외부 HWPX를 단순화해 덮어쓰지 않는다.

## 자체 테스트

Windows 실빌드 후:

```bat
RUN_HWPX_SELFTEST.cmd
```

테스트 순서:

1. 프로젝트 코드가 HWPX 생성
2. 필수 ZIP part 검사
3. `application/hwp+zip` 검사
4. `HwpxEditSafety` 검사
5. `HwpxReader`로 다시 읽기
6. 텍스트 수정
7. HWPX 재저장
8. 다시 Reader로 읽어 수정 내용 확인

이 테스트는 package 구조와 자체 round-trip만 검증한다.

**한컴 프로그램에서 정상 열림/저장됨을 증명하는 테스트가 아니다.**

## 저작권 / 라이선스 원칙

- 한글과컴퓨터 실행 파일 또는 DLL을 포함하지 않는다.
- 한컴 UI/아이콘/템플릿을 복제하지 않는다.
- 테스트 문서는 코드가 직접 생성한다.
- 공개 규격의 구조/element/attribute 이름을 구현하는 것과 타 프로그램 소스 코드를 복사하는 것을 구분한다.
- 공식 오픈소스 구현은 동작 확인과 규격 이해 참고로만 사용하며, 프로젝트 코드는 별도로 작성한다.
- 외부 라이브러리를 추가할 경우 별도 라이선스 검토 및 `THIRD_PARTY_NOTICES.md` 갱신이 필요하다.

## 다음 단계

- 실제 한컴에서 writer output 열기 테스트
- schema/validator 기반 검사
- page/section property 호환성 보강
- header style table 호환성 보강
- multiple sections
- table/image reader/writer
- unknown-part preservation
- HWPX용 `TextDocumentEditorForm` 연결
- 이후 HWP binary는 공개 HWP 5.x 명세 기반 **Reader 우선**으로 별도 구현
