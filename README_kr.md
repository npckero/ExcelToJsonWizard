# JSON 변환 도구

이 도구는 엑셀 파일을 JSON 파일로 변환하고 해당 JSON 파일을 로드할 수 있는 C# 로더 클래스를 생성합니다. `config.txt` 파일을 통해 다양한 옵션을 설정할 수 있으며, Enum 정의도 처리합니다.

## 주요 기능
- 엑셀 파일을 JSON 파일로 변환
- JSON 파일을 로드할 수 있는 C# 로더 클래스 생성
- 별도의 `Enum.xlsx` 파일을 통해 Enum 정의 지원
- `config.txt` 파일을 통해 설정 가능
- 다중 시트 처리 지원
- Unity의 `Resources` 폴더를 사용할 수 있는 옵션 제공

## 사전 준비
- .NET Core 또는 .NET Framework 설치
- ClosedXML 라이브러리 설치 (엑셀 파일 읽기용)

## 설치 방법
1. 저장소를 클론합니다.
2. 선호하는 IDE에서 솔루션을 엽니다.
3. NuGet 패키지를 복원합니다.

## 설정
이 도구는 `config.txt` 파일을 사용하여 설정을 관리합니다. `config.txt` 파일과 필요한 디렉토리는 처음 실행 시 자동으로 생성됩니다.

## 사용 방법
1. **엑셀 파일 준비**:
   - 엑셀 파일을 `config.txt`에 지정된 디렉토리에 넣습니다.
   - 엑셀 파일의 형식이 올바른지 확인합니다 (아래 참조).

2. **도구 실행**:
   - 컴파일된 프로그램을 실행합니다.
   - 도구는 엑셀 파일을 읽고 JSON 파일과 C# 로더 클래스를 생성하여 지정된 디렉토리에 저장합니다.

## 엑셀 파일 형식
- 첫 번째 행은 변수 이름을 포함해야 합니다.
- 두 번째 행은 데이터 타입을 포함해야 합니다.
- 세 번째 행은 설명을 포함해야 합니다 (선택 사항, 비어 있을 경우 "No description provided."가 사용됩니다).
- 네 번째 행부터는 데이터가 포함됩니다.

## Enum 정의
- Enum 정의는 별도의 `Enum.xlsx` 파일을 통해 처리됩니다.
- `Enum.xlsx` 파일의 첫 번째 시트를 사용합니다.
- 첫 번째 행은 Enum 이름을 포함해야 하며, 두 번째 열부터는 Enum 값을 포함해야 합니다.
- Enum 정의는 설명 줄 없이 첫 번째 행부터 바로 사용됩니다.
- Enum 정의가 필요하지 않다면 `Enum.xlsx` 파일을 만들 필요가 없습니다.

## 발생할 수 있는 오류 및 해결 방법
1. **`key` 열 누락**:
   - 엑셀 파일의 첫 번째 열이 `key`인지 확인하십시오.

2. **잘못된 데이터 타입**:
   - 두 번째 행에 지정된 데이터 타입이 올바르고 지원되는지 확인하십시오 (예: `int`, `string`, `List<int>` 등).

3. **중복된 `key` 값**:
   - `key` 열에 중복된 값이 없는지 확인하십시오.

4. **Enum 타입을 찾을 수 없음**:
   - `Enum.xlsx` 파일이 존재하고 올바르게 형식화되었는지 확인하십시오.
   - 엑셀 파일의 Enum 이름이 `Enum.xlsx`의 이름과 일치하는지 확인하십시오.

5. **Resource 로딩 문제**:
   - Unity의 `Resources.Load`를 사용하는 경우 경로가 정확하고 파일이 Resources 폴더에 존재하는지 확인하십시오.


## 활용 방법
1. 유니티 Resources 폴더 사용할 때 설정법
Unity의 Resources 폴더를 사용하여 JSON 파일을 로드하려면 config.txt 파일에서 useResources 옵션을 true로 설정하고, resourcesInternalPath를 설정합니다. 

다음 단계를 따르십시오:

   1. config.txt 파일을 엽니다.
   2. 다음 항목을 설정합니다:
   - useResources=true
   - resourcesInternalPath에 JSON 파일이 저장될 Resources 폴더 내의 경로를 설정합니다. 예를 들어, resourcesInternalPath=Data/JsonFiles라고 설정하면 Resources/Data/JsonFiles 경로를 사용하게 됩니다.
   3. 엑셀 파일을 변환하여 JSON 파일을 생성합니다.
   4. 생성된 JSON 파일을 Unity 프로젝트의 Resources 폴더 내의 설정된 경로에 복사합니다.

   json 파일 생성 경로를 미리 Resources 폴더에 연결해두면 바로사용 가능 합니다.

2. 멀티 시트 사용할 때 설정법
엑셀 파일에 여러 시트가 있는 경우, 각 시트를 변환하려면 config.txt 파일에서 allowMultipleSheets 옵션을 true로 설정합니다. 

다음 단계를 따르십시오:

   1. config.txt 파일을 엽니다.
   2. 다음 항목을 설정합니다:
   - allowMultipleSheets=true
   3. 엑셀 파일을 준비합니다. 각 시트는 서로 다른 데이터 구조를 가질 수 있습니다.
   4. 도구를 실행하여 JSON 파일과 C# 로더 클래스를 생성합니다. 각 시트에 대해 별도의 JSON 파일과 로더 클래스가 생성됩니다.

   예를 들어, Example_Multiple_Sheets.xlsx 파일이 있고 Sheet1과 Sheet2가 있다면, 다음과 같은 파일들이 생성됩니다:

   - Example_Multiple_Sheets_Sheet1.json
   - Example_Multiple_Sheets_Sheet2.json
   - Example_Multiple_Sheets_Sheet1Loader.cs
   - Example_Multiple_Sheets_Sheet2Loader.cs
   이렇게 생성된 파일들을 사용하여 다양한 시트 데이터를 처리할 수 있습니다.