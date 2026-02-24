# 전자책 자동 생성기 - Windows 10/11 설치 및 사용 가이드

> **지원 출력 형식**: PDF · DOCX(워드) · PPTX(파워포인트) · **HWP(한글)**

---

## 준비물

| 항목 | 내용 |
|------|------|
| 운영체제 | Windows 10 또는 Windows 11 (64비트) |
| Python | 3.10 이상 (무료) |
| 인터넷 | 최초 패키지 설치 시 필요 |
| ChatGPT 계정 | AI 생성 기능 사용 시 필요 (테스트 모드는 불필요) |

---

## 설치 순서

### Step 1 — Python 설치

1. https://www.python.org/downloads/ 방문
2. **"Download Python 3.x.x"** 클릭
3. 설치 화면에서 **"Add Python to PATH"** ← 반드시 체크!
4. **"Install Now"** 클릭
5. 설치 완료 후 **PC 재시작**

### Step 2 — 앱 폴더 복사

앱 폴더 전체를 원하는 위치에 복사합니다.

```
예: C:\Users\사용자이름\Desktop\ebook-creator\
```

### Step 3 — 실행

`start_windows.bat` 파일을 **더블클릭**

→ 자동으로 패키지 설치 후 브라우저가 열립니다 (최초 실행 시 1~2분 소요)

---

## 사용 방법

### 기본 흐름

1. `start_windows.bat` 더블클릭 → 브라우저에서 `http://localhost:5000` 자동 오픈
2. **"ChatGPT 로그인"** 클릭 → ChatGPT 계정으로 로그인
3. 전자책 주제 입력
4. 원하는 출력 형식 선택:
   - ✅ PDF
   - ✅ DOCX (워드)
   - ✅ PPTX (파워포인트)
   - ✅ **HWP (한글)** ← 새로 추가된 형식
5. **"전자책 생성 시작"** 클릭
6. 완료 후 각 형식의 다운로드 버튼 클릭

### 테스트 모드 (AI 없이 즉시 생성)

- 로그인 없이 사용 가능
- 메인 화면 하단의 **"🧪 테스트 모드"** 버튼 클릭
- 약 3~5초 내 모든 형식 샘플 파일 생성

---

## 출력 파일 위치

```
ebook-creator\
  └── static\
        └── output\      ← 생성된 파일 여기에 저장
              ├── 제목.pdf
              ├── 제목.docx
              ├── 제목.pptx
              └── 제목.hwpx   ← HWP(한글) 파일
```

> **HWP 파일 열기**: 한컴오피스 또는 [한컴오피스 한글 뷰어](https://www.hancom.com/cs_center/csDownload.do) (무료) 설치 필요

---

## 앱 종료

- 검정 콘솔 창을 닫으면 서버 종료
- 또는 콘솔 창에서 `Ctrl+C`

---

## 자주 묻는 문제

| 증상 | 해결 방법 |
|------|-----------|
| `python is not recognized` 오류 | Python 재설치 → **"Add Python to PATH"** 체크 → PC 재시작 |
| 브라우저가 안 열림 | 수동으로 `http://localhost:5000` 접속 |
| 포트 5000 이미 사용 중 | `launcher.py` 에서 `port=5000` → `port=5001` 로 변경 |
| 패키지 설치 오류 | 우클릭 → **"관리자로 실행"** 으로 `start_windows.bat` 실행 |
| PDF 폰트 깨짐 | 설정 페이지에서 Windows 폰트 경로 지정 (예: `C:\Windows\Fonts\malgun.ttf`) |
| HWP 파일이 안 열림 | 한컴오피스 또는 한글 뷰어 설치 필요 (무료 다운로드 가능) |
| 로그인이 안 됨 | ChatGPT 계정 필요. 없으면 테스트 모드 사용 |
| 바이러스 경고 (.exe) | PyInstaller exe는 오탐 많음 — Python 직접 실행(방법 1) 권장 |

---

## 방법 2: 독립 실행파일(.exe) 빌드 (선택사항)

> Python이 없는 환경에 배포할 때 사용. Windows PC에서 빌드해야 합니다.

1. `start_windows.bat` 를 먼저 한 번 실행해 패키지 설치
2. `build_windows.bat` 더블클릭
3. 빌드 완료 후 `dist\전자책생성기\` 폴더 전체 배포

---

## 폴더 구조

```
ebook-creator\
  ├── start_windows.bat     ← 실행 파일 (더블클릭)
  ├── build_windows.bat     ← EXE 빌드 (선택)
  ├── app.py                ← Flask 서버
  ├── launcher.py           ← 브라우저 자동 열기
  ├── config.py             ← 설정
  ├── requirements.txt      ← Python 패키지 목록
  ├── modules\
  │     ├── hwpx_generator.py   ← HWP(한글) 생성기
  │     ├── pdf_generator.py    ← PDF / DOCX / PPTX 생성기
  │     └── ai_engine.py        ← AI 엔진
  ├── templates\            ← HTML 화면
  ├── static\               ← CSS, JS, 출력 파일
  └── venv\                 ← Python 가상환경 (자동 생성, 복사 불필요)
```

---

## 다른 PC로 이전할 때

1. `ebook-creator` 폴더 전체를 복사 (또는 ZIP 압축)
2. `venv\` 폴더는 **복사하지 않아도 됩니다** (새 PC에서 자동 재생성)
3. 새 PC에서 `start_windows.bat` 실행

---

## 보안 메모

- 로그인 토큰은 `token_data.json`에 **로컬 저장** (외부 전송 없음)
- 서버는 `127.0.0.1(내 PC)`에서만 접근 가능 — 외부 노출 없음
- `token_data.json` 파일은 타인과 공유하지 마세요
