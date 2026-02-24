@echo off
chcp 65001 >nul
echo =====================================================
echo  전자책 자동 생성기 - Windows 빌드 스크립트
echo =====================================================
echo.

:: Python 설치 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되지 않았습니다.
    echo Python 3.10 이상을 설치하세요: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/4] Python 버전 확인 중...
python --version

echo.
echo [2/4] 패키지 설치 중...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [3/4] 빌드 시작...
pyinstaller ebook_creator.spec --clean --noconfirm

echo.
if exist "dist\전자책생성기\전자책생성기.exe" (
    echo [4/4] 빌드 성공!
    echo.
    echo 실행 파일 위치: dist\전자책생성기\전자책생성기.exe
    echo 또는 dist\전자책생성기\ 폴더 전체를 배포하세요.
) else (
    echo [오류] 빌드 실패. 위 오류 메시지를 확인하세요.
)

echo.
pause
