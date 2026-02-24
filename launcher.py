"""
전자책 자동 생성기 - Windows 런처
Flask 서버를 시작하고 기본 브라우저에서 자동으로 엽니다.
"""
import sys
import os
import threading
import webbrowser
import time

# PyInstaller 번들 환경에서 경로 설정
if getattr(sys, 'frozen', False):
    # PyInstaller 실행 환경
    BASE_DIR = sys._MEIPASS
    APP_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    APP_DIR = BASE_DIR

# 작업 디렉토리를 앱 디렉토리로 설정
os.chdir(APP_DIR)

# 출력 디렉토리 생성
output_dir = os.path.join(APP_DIR, 'output')
os.makedirs(output_dir, exist_ok=True)

# sys.path에 앱 디렉토리 추가
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)


def open_browser():
    """서버 시작 후 브라우저 열기"""
    time.sleep(2.0)
    webbrowser.open('http://localhost:5000')


def main():
    print("=" * 50)
    print("  전자책 자동 생성기 시작 중...")
    print("=" * 50)
    print()

    # 브라우저 자동 열기 (별도 스레드)
    browser_thread = threading.Thread(target=open_browser, daemon=True)
    browser_thread.start()

    print("  서버 주소: http://localhost:5000")
    print("  종료: Ctrl+C 또는 이 창을 닫으세요")
    print()

    # Flask 앱 import 및 실행
    from app import app
    from config import load_config

    config = load_config()
    # 출력 디렉토리를 앱 디렉토리 기준으로 재설정
    config['output_dir'] = output_dir

    app.run(host='127.0.0.1', port=5000, debug=False, use_reloader=False)


if __name__ == '__main__':
    main()
