# PDF Export Foundation

현재 `feature/office-foundation`의 첫 PDF 내보내기 구현은 **raster-page 방식**이다.

## 목적

초기 목표는 다음을 만족하는 것이다.

- PPTX/ODP presentation model을 PDF로 출력
- DOCX/HWPX/HWP/ODT text model을 PDF로 출력
- XLSX/ODS spreadsheet model을 PDF로 출력
- 별도 PDF/NuGet 라이브러리를 필수 dependency로 추가하지 않음
- 사용자의 TTF/OTF 파일을 허가 확인 없이 PDF에 넣지 않음

## 현재 구조

```text
Document model / rendered slide
        ↓
System.Drawing raster page
        ↓
JPEG stream
        ↓
RasterPdfWriter
        ↓
PDF 1.4 image page
```

관련 파일:

```text
src/PdfExport.cs
src/PdfExportDiagnostics.cs
src/MainFormPdfExport.cs
src/DocumentPdfExport.cs
RUN_PDF_SELFTEST.cmd
```

## 폰트 정책

현재 출력은 글자를 Windows/GDI+로 페이지 이미지에 렌더링한 뒤 그 결과 이미지를 PDF에 넣는다.

따라서 현재 구현은 원본 `.ttf` / `.otf` font binary를 PDF file 안에 embedding하지 않는다.

이것은 향후 font embedding 권한 검사를 생략해도 된다는 뜻이 아니다.

향후 searchable/vector text PDF를 추가하면 다음 순서를 강제한다.

```text
PDF text export request
        ↓
font binary embedding/subsetting 필요?
        ↓
FontLicenseService
        ↓
fsType metadata + actual license review
        ↓
허용된 경우에만 subset/embed
```

실제 라이선스 문구가 `fsType`보다 우선한다.

## 현재 제한

Raster PDF는 다음 한계가 있다.

- text search/copy 불가
- accessibility text layer 없음
- 큰 문서는 file size가 커질 수 있음
- vector shape가 image로 변환됨
- 인쇄 확대 시 vector PDF보다 선명도가 낮을 수 있음
- source application과 pixel-perfect pagination을 보장하지 않음

따라서 현재는 `experimental`이다.

## 문서 손실과 원본 안전

PDF Export는 source document를 수정하지 않는다.

출력 PDF는 `.writing` staging file에 먼저 작성하고 성공 후 destination으로 교체한다. 기존 destination이 있으면 backup을 사용해 교체 실패 시 복구를 시도한다.

## 자체 테스트

Windows 빌드 후:

```bat
RUN_PDF_SELFTEST.cmd
```

테스트는 프로젝트 코드가 만든 presentation/text/spreadsheet model을 사용하며 인터넷 문서나 상용 template을 사용하지 않는다.

검사 범위:

- PDF header
- non-empty output
- EOF marker
- presentation export
- text-document export
- spreadsheet export

구조 test만으로 PDF Viewer/Printer 호환 완료라고 판단하지 않는다.

추가 수동 테스트:

- Edge/Chrome/Adobe Acrobat 등 일반 PDF Viewer에서 열기
- 실제 프린터 또는 Microsoft Print to PDF preview
- 한글/영문/숫자 혼합 문서
- 긴 문서 page break
- 여러 worksheet page split
- 프레젠테이션 비율 유지

## 서드파티 자산

현재 PDF writer는 프로젝트 자체 C# 코드다.

외부 PDF 라이브러리를 향후 도입할 경우 먼저 해당 라이브러리의 라이선스와 재배포 조건을 검토하고 `THIRD_PARTY_NOTICES.md` 및 필요한 license 원문을 갱신한다.
