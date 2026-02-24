@echo off
chcp 65001 >nul
title 전자책 자동 생성기

echo =====================================================
echo   전자책 자동 생성기 v2.0  (PDF/DOCX/PPTX/HWP)
echo =====================================================
echo.

:: Python 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되지 않았습니다.
    echo.
    echo  1. https://www.python.org/downloads/ 방문
    echo  2. "Download Python 3.x.x" 클릭
    echo  3. 설치 시 "Add Python to PATH" 반드시 체크
    echo  4. 설치 완료 후 PC 재시작
    echo  5. 이 파일 다시 실행
    echo.
    pause
    exit /b 1
)

echo [확인] Python 버전:
python --version
echo.

:: 가상환경 생성 (없으면)
if not exist "venv\Scripts\activate.bat" (
    echo [설정] 가상환경 생성 중... (최초 1회만 수행)
    python -m venv venv
    if errorlevel 1 (
        echo [오류] 가상환경 생성 실패.
        pause
        exit /b 1
    )
)

:: 가상환경 활성화
call venv\Scripts\activate.bat

:: 패키지 설치 / 업데이트
echo [설정] 필요한 패키지 확인 및 설치 중...
pip install -r requirements.txt --quiet --upgrade
if errorlevel 1 (
    echo [경고] 일부 패키지 설치 실패. 계속 시도합니다.
)

:: 출력 디렉토리 생성
if not exist "output" mkdir output
if not exist "static\output" mkdir static\output

:: 앱 실행
echo.
echo =====================================================
echo   서버 시작 중...
echo   잠시 후 브라우저가 자동으로 열립니다.
echo.
echo   수동 접속: http://localhost:5000
echo   종료:      이 창을 닫거나 Ctrl+C
echo =====================================================
echo.
python launcher.py

pause
