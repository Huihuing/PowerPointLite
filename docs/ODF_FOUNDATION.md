# ODF Foundation

`feature/office-foundation`의 ODF 구현은 `.odt`, `.ods`, `.odp`를 공통 ODF package 계층 위에서 독립적으로 읽고 쓰기 위한 experimental 기반이다.

## 공통 원칙

- OpenDocument 공개 규격 구조를 기반으로 자체 Reader/Writer 구현
- LibreOffice 실행 파일/DLL을 앱에 번들하지 않음
- Microsoft/Hancom 자산을 사용하지 않음
- 외부 문서를 현재 model보다 단순한 형태로 덮어쓰지 않음
- project-generated package가 아니거나 unknown part가 있으면 편집 deny-by-default
- 암호화 metadata가 있는 문서는 현재 rewrite하지 않음

공통 파일:

```text
src/OdfPackage.cs
```

지원 mimetype:

```text
application/vnd.oasis.opendocument.text
application/vnd.oasis.opendocument.spreadsheet
application/vnd.oasis.opendocument.presentation
```

Writer는 ODF package의 `mimetype` entry를 no-compression으로 작성하고 manifest/content/styles/meta part를 생성한다.

## ODT

관련:

```text
src/OdtCore.cs
src/OdtDiagnostics.cs
src/TextDocumentEditSession.cs
src/UnifiedTextDocumentEditorForm.cs
RUN_ODT_SELFTEST.cmd
```

공통 `TextDocument` model에 mapping한다.

현재:

- paragraph / run
- text
- font family / size
- bold / italic / underline
- text color
- paragraph alignment
- tab / line break
- create/read/edit/save

아직 complex lists, tables, images, page layout, styles hierarchy 전체 호환을 주장하지 않는다.

## ODS

관련:

```text
src/OdsCore.cs
src/OdsDiagnostics.cs
src/OdsSpreadsheetEditorForm.cs
RUN_ODS_SELFTEST.cmd
```

공통 `SpreadsheetDocument` model에 mapping한다.

현재:

- multiple worksheets
- text / number / boolean cells
- formula storage
- add/delete/move/rename sheet
- grid editor
- input/formula bar
- create/read/edit/save

Formula는 현재 저장/읽기 대상이며 full calculation engine이 아니다.

## ODP

관련:

```text
src/OdpCore.cs
src/OdpDiagnostics.cs
src/OdpPresentationEditorForm.cs
RUN_ODP_SELFTEST.cmd
```

공통 `PresentationDocument` model에 mapping한다.

현재:

- slides
- text frames
- system font family reference
- basic text formatting
- images
- rectangle / ellipse writer-reader fidelity 기반
- drag / resize editor canvas
- create/read/edit/save

주의: ODF shape encoding을 완전히 구현하기 전에는 PPTX model의 모든 shape kind를 ODP에서 동일하게 round-trip할 수 있다고 주장하지 않는다. Unsupported shape kind는 silent downgrade하지 않는 방향으로 UI/Writer 검증을 강화해야 한다.

## Safety gate

각 포맷의 `*EditSafety`는 최소 다음을 검사한다.

```text
expected mimetype?
project-generated package?
known parts only?
encryption/unsupported metadata absent?
```

안전하지 않으면 Reader로 조사할 수 있더라도 Editor rewrite는 허용하지 않는 것이 기본 원칙이다.

## 테스트

```bat
RUN_ODT_SELFTEST.cmd
RUN_ODS_SELFTEST.cmd
RUN_ODP_SELFTEST.cmd
```

전체:

```bat
RUN_ALL_FORMAT_SELFTESTS.cmd
```

Structural self-test와 실제 LibreOffice interoperability는 별개다.

실제 안정 지원 판단 전 확인:

- LibreOffice Writer → ODT open/save
- LibreOffice Calc → ODS open/save
- LibreOffice Impress → ODP open/save
- project-generated file을 LibreOffice에서 열었다 다시 저장한 뒤 Reader robustness 확인
- Unicode/Korean text
- multiple pages/sheets/slides
- images
- unsupported feature preservation policy

테스트용 문서는 프로젝트 코드가 직접 만들며 인터넷에서 가져온 타인 문서를 repository fixture로 사용하지 않는다.
