# Third-Party Notices

이 파일은 배포물에 포함되는 서드파티 코드, 라이브러리, 폰트, 아이콘, 기타 자산의 출처와 라이선스를 기록하기 위한 목록이다.

프로젝트 자체 코드는 루트 [`LICENSE`](LICENSE)의 MIT License를 기준으로 한다. 이 MIT License가 외부 폰트, 라이브러리, 아이콘, 이미지, 문서 샘플, 상표 등 제3자 권리까지 다시 허가하는 것은 아니다.

현재 저장소는 자체 C# 구현을 중심으로 하며 Microsoft Office 또는 Hancom의 실행 파일/DLL을 포함하지 않는다.

## 작성 규칙

새 외부 자산을 추가할 때 아래 항목을 채운다.

```text
Name:
Version:
Type: library / font / icon / image / other
Source:
License:
Bundled files:
Redistribution notes:
Modification notes:
```

라이선스가 불명확한 자산은 배포물에 포함하지 않는다.

## Current bundled third-party code/assets

현재 `feature/office-foundation` 기준으로 별도의 필수 NuGet package, Microsoft/Hancom DLL, 상용 template, 재배포용 TTF/OTF font file을 번들하지 않는다.

Windows/.NET Framework/System.Drawing/System.Windows.Forms 등은 사용자의 Windows 환경에서 제공되는 platform API를 사용한다.

## Fonts

현재 기본 저장소에는 재배포를 전제로 한 TTF/OTF 번들 폰트를 추가하지 않는다.

향후 Noto 등 오픈 폰트를 번들할 경우 해당 버전의 실제 OFL/LICENSE 원문을 `licenses/`에 함께 보관하고 이 문서에 정확한 출처를 기록한다.

PDF의 초기 raster export는 source font file을 PDF에 자동 embedding하지 않는다. 향후 vector/searchable text PDF에서 font embedding을 추가할 경우 font license review와 `FontLicenseService` 검사를 거친다.

## Product trademarks

Microsoft, Word, Excel and PowerPoint are trademarks of Microsoft Corporation. This project is not affiliated with or endorsed by Microsoft.

Hancom and related product names are trademarks of their respective owners. This project is not affiliated with or endorsed by Hancom.
