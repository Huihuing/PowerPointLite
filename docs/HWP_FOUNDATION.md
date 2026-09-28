# HWP 5.x Read-only Foundation

HWP binary 지원은 **Reader 우선 / Writer 보류** 원칙으로 진행한다.

프로젝트는 HWP 문서 형식의 공개 명세를 참고해 독립 구현하며 한글과컴퓨터 실행 파일이나 DLL을 앱에 포함하지 않는다.

## 현재 소스

```text
src/CompoundFileReader.cs
src/HwpReader.cs
src/HwpReadOnlyViewerForm.cs
src/HwpDiagnostics.cs
RUN_HWP_PARSER_SELFTEST.cmd
```

## 현재 처리 흐름

```text
.hwp
 ↓
Compound File Binary (CFB)
 ↓
DIFAT / FAT / MiniFAT / Directory
 ↓
/FileHeader
 ↓
version / flags
 ↓
/BodyText/SectionN
 ↓
raw DEFLATE when compressed
 ↓
32-bit HWP record headers
 ↓
HWPTAG_PARA_TEXT
 ↓
UTF-16LE plain text
 ↓
TextDocument
 ↓
read-only viewer
```

## Compound File Reader

`CompoundFileReader`는 외부 OLE/CFB 라이브러리를 필수 의존성으로 추가하지 않고 다음을 처리한다.

- CFB signature
- version 3 / 4 sector size
- DIFAT
- FAT
- MiniFAT
- Directory entries
- storage child/sibling traversal
- regular stream
- mini stream
- stream path lookup
- loop/out-of-range safety checks

초기 HWP Reader에 필요한 다음 stream을 읽을 수 있도록 설계했다.

```text
FileHeader
DocInfo
BodyText/Section0
BodyText/Section1
...
PrvText
```

현재 DocInfo는 전체 style mapping에 아직 사용하지 않는다.

## FileHeader

현재 Reader는 다음을 확인한다.

- signature `HWP Document File`
- version
- compressed flag
- password encryption flag
- distribution document flag
- DRM flag
- certificate encryption flag

보호된 문서에 대해서는 암호 해제/우회 기능을 구현하지 않는다.

```text
password encrypted   → reject
publication/distribution protection → reject in initial reader
DRM                  → reject
certificate encryption → reject
```

## BodyText record parser

HWP 5.x record header를 다음 형태로 읽는다.

```text
10 bits  tag id
10 bits  level
12 bits  size
```

size가 `0xFFF`이면 다음 32-bit 값을 확장 크기로 읽는다.

초기 text reader는 `HWPTAG_PARA_TEXT` record의 plain text를 추출한다.

control character는 보수적으로 필터링하며 tab/line break/non-breaking hyphen/space 일부를 텍스트로 변환한다.

이 단계에서는 다음을 완전 지원한다고 주장하지 않는다.

- field/control semantics
- full character shape mapping
- paragraph shape mapping
- tables
- images
- equations
- footnotes/endnotes
- headers/footers
- drawing objects
- tracked changes

## Compression

FileHeader의 compression flag가 설정되면 BodyText section stream을 `DeflateStream`으로 해제한다.

비정상/지원되지 않는 압축 stream이면 문서를 임의 해석하지 않고 오류로 중단한다.

## Read-only UI

`HwpReadOnlyViewerForm`은 현재 HWP를 plain paragraph text로 확인하기 위한 창이다.

- Save 없음
- Save As 없음
- Writer 없음
- 원본 변경 없음

전체 HWP 편집 기능이 준비되기 전까지 이 제한을 유지한다.

## 자체 테스트

```bat
RUN_HWP_PARSER_SELFTEST.cmd
```

현재 synthetic test는:

1. FileHeader signature/version/flags parser
2. synthetic HWP record header
3. `HWPTAG_PARA_TEXT` extraction
4. 일부 inline control filtering

을 검사한다.

저작권/개인정보가 불명확한 실제 HWP 파일을 repository fixture로 추가하지 않는다.

실제 HWP 검증은 사용 권한이 명확한 문서로 별도 수행한다.

## 다음 단계

1. rights-cleared HWP 5.x sample로 CFB/BodyText 실제 검증
2. DocInfo record parser
3. FACE_NAME / CHAR_SHAPE / PARA_SHAPE mapping
4. paragraph formatting
5. TABLE/control record read
6. BinData/image read
7. header/footer/footnote read
8. page layout
9. read-only renderer 확대
10. Writer는 Reader 안정화와 보존 전략 확인 뒤 feasibility 재검토

HWP Writer를 빠르게 추가하는 것보다 **기존 HWP 문서를 손상시키지 않는 Reader**를 만드는 것이 우선이다.
