# MS 출결확인 자동화 프로그램

이 프로그램은 MS Teams 출석 정보와 참석 보고서를 기반으로 MS 출결확인 양식 Excel 파일을 자동으로 업데이트하는 Flask 웹 애플리케이션입니다.

## 주요 기능

1. 출석 정보(attendance_20250411_1118.csv)에서 학생들의 출석 여부를 확인하여 "중간출결" 열에 "O" 표시
2. 참석보고서(파이썬 기본 및 데이터분석2 참석 보고서 41125.csv)에서 학생들의 참여 정보를 추출하여 다음 항목 업데이트:
   - "들어온 시간" → "접속시작시간" 열에 업데이트
   - "마지막 나간 시간" → "접속종료시간" 열에 업데이트
   - "모임 참여 시간" → "활용시간" 열에 업데이트
3. 업데이트된 Excel 파일을 즉시 다운로드 가능
4. 보안 및 디스크 공간 관리를 위해 다운로드 후 1분 뒤 자동 파일 삭제

## 프로젝트 구조

```
auto_attend/
├── .git/                   # Git 버전 관리 폴더
├── .github/                # GitHub 관련 설정 폴더
│    └── workflows/         # GitHub Actions 워크플로우 폴더
│         └── main_autoattend.yml  # Azure 배포 워크플로우 파일
├── .venv/                  # 파이썬 가상 환경 (필요 시 생성)
├── results/                # 결과 파일 저장 폴더
├── static/                 # 정적 파일 폴더
│    └── results/           # 웹에서 접근 가능한 결과 파일 폴더
├── templates/              # HTML 템플릿 폴더
│    └── index.html         # 메인 페이지 HTML
├── test/                   # 테스트 관련 폴더
├── uploads/                # 업로드된 파일 임시 저장 폴더
├── .gitignore              # Git 무시 파일 설정
├── app.py                  # 메인 애플리케이션 서버 (Flask)
├── README.md               # 프로젝트 설명서
├── requirements.txt        # 의존성 패키지 목록
├── startup.txt             # Azure WebApp 시작 명령 파일
├── web.config              # Azure WebApp 구성 파일
└── wsgi.py                 # WSGI 애플리케이션 진입점
```

위 구조는 프로젝트의 전체적인 파일 및 폴더 구성을 보여줍니다. 각 파일과 폴더의 역할을 이해하면 프로젝트를 더 쉽게 관리하고 수정할 수 있습니다.

## 시스템 요구사항

- Python 3.6 이상
- Flask 2.0.1
- Werkzeug 2.0.1
- Pandas 1.3.5
- NumPy 1.21.6
- openpyxl 3.0.10
- Gunicorn 20.1.0 (Azure 배포용)

## 설치 방법

1. 필요한 Python 라이브러리 설치:

```bash
pip install -r requirements.txt
```

## 사용 방법

### 로컬 환경

1. 다음 명령어로 서버 실행:

```bash
python app.py
```

2. 웹 브라우저에서 http://localhost:5000/ 접속
3. 파일 업로드 폼에서 다음 파일들을 선택:
   - MS 출결확인 양식.xlsx
   - attendance_20250411_1118.csv (출석 정보)
   - 파이썬 기본 및 데이터분석2 참석 보고서 41125.csv
4. "업데이트 실행" 버튼 클릭
5. 업데이트 결과가 표시되며, "업데이트된 Excel 파일 다운로드" 버튼을 클릭하여 파일을 즉시 다운로드할 수 있습니다.
6. 다운로드 후 1분 뒤에 서버에서 파일이 자동으로 삭제됩니다.

### Azure 배포 환경

애플리케이션은 Azure WebApp에 배포되어 있으며, 다음 URL에서 접근 가능합니다:
https://autoattend.azurewebsites.net

## 배포 관련 파일

- **web.config**: Azure WebApp의 IIS 웹 서버 구성 파일
- **startup.txt**: Azure WebApp 시작 명령어 지정 (Gunicorn 사용)
- **wsgi.py**: WSGI 애플리케이션 진입점
- **.github/workflows/main_autoattend.yml**: GitHub Actions를 통한 자동 배포 구성

## 주의사항

1. 참석보고서 CSV 파일은 UTF-16LE 인코딩으로 되어 있어야 합니다.
2. MS 출결확인 양식 Excel 파일에는 "출결정보" 시트가 있어야 하고 다음 열이 포함되어야 합니다:
   - "성명" (D열)
   - "중간출결" (K열)
   - "접속시작시간" (H열)
   - "접속종료시간" (I열) 
   - "활용시간" (J열)
3. 출석 정보 CSV 파일에는 "이름"과 "출석여부" 열이 포함되어야 합니다.
4. 업데이트된 셀은 주황색(#FFB366)으로 표시됩니다.
5. 생성된 파일은 다운로드 후 1분 뒤에 자동으로 삭제되므로 필요한 경우 다운로드 파일을 반드시 저장해 두세요.

## 문제 해결

- 파일 업로드 오류 발생 시 인코딩 문제일 수 있습니다. 참석보고서 CSV 파일이 UTF-16LE 인코딩인지 확인하세요.
- 열 구조가 맞지 않을 경우 app.py 파일에서 해당 열 인덱스를 수정해야 할 수 있습니다.
- 다운로드 링크가 작동하지 않는 경우 results 폴더에 쓰기 권한이 있는지 확인하세요.
- Azure WebApp 배포 시 발생하는 패키지 호환성 문제는 requirements.txt 파일에 명시된 패키지 버전을 통해 해결됩니다.

## 업데이트 내역

- 2025-05-20: Azure WebApp 배포를 위해 패키지 버전 호환성 문제 해결
  - Flask 및 Werkzeug 버전 고정 (Flask==2.0.1, Werkzeug==2.0.1)
  - Pandas 및 NumPy 버전 호환성 조정 (Pandas==1.3.5, NumPy==1.21.6)
  - WSGI 서버로 Gunicorn 사용 (버전 20.1.0)
  - Azure WebApp 배포 관련 파일 추가 (web.config, wsgi.py)
