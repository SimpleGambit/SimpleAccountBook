# SimpleAccountBook
간단한 가계부 WPF 애플리케이션입니다. 엑셀에서 불러온 거래 내역을 달력에 표시하고, 월별 합계 및 자산 변화를 확인할 수 있습니다.

## 실행 방법

1. **필수 요소 설치**
   - Windows 10 이상 환경
   - [.NET 7 SDK](https://dotnet.microsoft.com/en-us/download) (WPF 데스크톱 워크로드 포함)
2. **의존성 복원 및 빌드**
   ```bash
   dotnet restore
   dotnet build
   ```
3. **애플리케이션 실행**
   ```bash
   dotnet run --project SimpleAccountBook.csproj
   ```