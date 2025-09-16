# 📧 이메일 업무 분석기

Azure OpenAI 서비스를 활용하여 이메일의 내용을 분석하고, 주요 업무를 요약 및 추출하여 생산성을 높여주는 Streamlit 웹 애플리케이션입니다.
바로가기 : [https://mvp-sunny-project-h3fvhmfbfagsg4d9.eastus2-01.azurewebsites.net]

## ✨ 주요 기능

- **이메일 자동 분석**: 텍스트를 직접 입력하거나 파일을 업로드하여 이메일 내용을 분석합니다.
- **다양한 파일 형식 지원**: `.eml`, `.msg`, `.txt` 형식의 이메일 파일을 지원합니다.
- **핵심 정보 추출**:
  - 이메일의 핵심 내용을 2-3 문장으로 요약합니다.
  - 주요 포인트, 즉시 처리 항목, 후속 조치 사항을 식별합니다.
- **업무 관리**:
  - 이메일에서 할 일(Task)을 추출하여 우선순위, 마감일, 담당자를 자동으로 제안합니다.
  - 분석된 할 일 목록을 기반으로 일정(.ics) 파일을 생성하여 Google Calendar, Outlook 등에 쉽게 추가할 수 있습니다.
- **다양한 내보내기 옵션**:
  - 분석 결과를 PDF 보고서 또는 Excel 파일로 다운로드할 수 있습니다.
- **분석 히스토리**: 이전 분석 결과를 다시 확인하고 내보낼 수 있습니다.

## ⚙️ 실행 환경 요구사항

실행에 필요한 주요 라이브러리는 다음과 같습니다.

- `streamlit`
- `openai`
- `pandas`
- `reportlab`
- `openpyxl`
- `python-dotenv` (로컬 환경에서 환경 변수 관리를 위해 권장)

## 🚀 설치 및 실행

1.  **필요한 라이브러리 설치**

    ```bash
    pip install streamlit openai pandas reportlab openpyxl python-dotenv
    ```

2.  **Azure OpenAI 서비스 설정**

    이 애플리케이션을 사용하려면 Azure OpenAI 서비스의 인증 정보가 필요합니다. 정보를 설정하는 방법은 두 가지입니다.

    **방법 1: 환경 변수 설정 (권장)**

    프로젝트 루트 디렉터리에 `.env` 파일을 생성하고 아래와 같이 Azure OpenAI 정보를 입력합니다.

    ```
    AZURE_OPENAI_ENDPOINT="<YOUR_AZURE_OPENAI_ENDPOINT>"
    AZURE_OPENAI_API_KEY="<YOUR_AZURE_OPENAI_API_KEY>"
    ```

    - `<YOUR_AZURE_OPENAI_ENDPOINT>`: Azure OpenAI 서비스의 Endpoint 주소 (예: `https://my-resource.openai.azure.com/`)
    - `<YOUR_AZURE_OPENAI_API_KEY>`: Azure OpenAI 서비스의 API 키

    **방법 2: 앱에서 직접 입력**

    애플리케이션 실행 후, 사이드바의 설정 메뉴에서 직접 Endpoint와 API Key를 입력할 수 있습니다. (앱이 실행되는 동안에만 유효합니다.)

3.  **애플리케이션 실행**

    터미널에서 아래 명령어를 실행합니다.

    ```bash
    streamlit run EmailAnalyzer_azure.py
    ```

## 📖 사용 방법

1.  **인증 정보 설정**: 앱 사이드바에서 Azure OpenAI Endpoint와 API Key를 입력하거나, 사전에 환경 변수를 설정합니다.
2.  **이메일 입력**:
    - **텍스트 입력**: '✍️ 텍스트로 입력/수정' 탭에 이메일 내용을 직접 붙여넣습니다.
    - **파일 업로드**: '📎 파일로 업로드' 탭에서 `.eml`, `.msg`, `.txt` 파일을 업로드합니다. 업로드된 내용은 자동으로 텍스트 입력 탭에 표시되어 확인 및 수정할 수 있습니다.
3.  **분석 실행**: `🔍 이메일 분석하기` 버튼을 클릭하여 분석을 시작합니다.
4.  **결과 확인**: 오른쪽에 표시되는 분석 결과(요약, 주요 포인트, 할일 목록 등)를 확인합니다.
5.  **결과 활용**: 필요에 따라 `📄 PDF 내보내기`, `📊 Excel 내보내기`, `📅 일정 내보내기` 버튼을 사용하여 결과를 다운로드하고 활용합니다.

