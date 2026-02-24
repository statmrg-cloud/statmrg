"""
ChatGPT OAuth PKCE 인증 모듈
- auth.openai.com OAuth 2.1 + PKCE (S256) 흐름
- Codex 공개 클라이언트 사용
- 로컬 콜백 서버로 토큰 수신
- 자동 토큰 갱신 (만료 5분 전)
"""
import os
import json
import time
import hashlib
import base64
import secrets
import threading
import webbrowser
from urllib.parse import urlencode, parse_qs, urlparse
from http.server import HTTPServer, BaseHTTPRequestHandler

import requests

# ============================================================
# OpenAI OAuth 설정 (Codex 공개 클라이언트)
# ============================================================
AUTH_ENDPOINT = 'https://auth.openai.com/oauth/authorize'
TOKEN_ENDPOINT = 'https://auth.openai.com/oauth/token'
CLIENT_ID = 'app_EMoamEEZ73f0CkXaXp7hrann'
REDIRECT_PORT = 1455
REDIRECT_URI = f'http://localhost:{REDIRECT_PORT}/auth/callback'
SCOPES = 'openid profile email offline_access'
TOKEN_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'token_data.json')


# ============================================================
# PKCE 유틸리티
# ============================================================
def generate_code_verifier():
    """PKCE code_verifier 생성 (43~128 문자)"""
    return secrets.token_urlsafe(64)[:128]


def generate_code_challenge(verifier):
    """PKCE code_challenge 생성 (S256)"""
    digest = hashlib.sha256(verifier.encode('ascii')).digest()
    return base64.urlsafe_b64encode(digest).rstrip(b'=').decode('ascii')


# ============================================================
# 토큰 저장/로드
# ============================================================
def save_tokens(token_data):
    """토큰을 파일에 저장"""
    token_data['saved_at'] = time.time()
    with open(TOKEN_FILE, 'w') as f:
        json.dump(token_data, f)


def load_tokens():
    """저장된 토큰 로드"""
    if not os.path.exists(TOKEN_FILE):
        return None
    try:
        with open(TOKEN_FILE, 'r') as f:
            return json.load(f)
    except Exception:
        return None


def clear_tokens():
    """저장된 토큰 삭제"""
    if os.path.exists(TOKEN_FILE):
        os.remove(TOKEN_FILE)


def is_token_valid(token_data):
    """토큰이 아직 유효한지 확인 (만료 5분 전이면 갱신 필요)"""
    if not token_data or 'access_token' not in token_data:
        return False
    saved_at = token_data.get('saved_at', 0)
    expires_in = token_data.get('expires_in', 3600)
    # 만료 5분 전 체크
    return time.time() < saved_at + expires_in - 300


# ============================================================
# 토큰 갱신
# ============================================================
def refresh_access_token(refresh_token):
    """refresh_token으로 access_token 갱신"""
    try:
        resp = requests.post(TOKEN_ENDPOINT, data={
            'grant_type': 'refresh_token',
            'client_id': CLIENT_ID,
            'refresh_token': refresh_token,
        }, timeout=15)

        if resp.status_code == 200:
            token_data = resp.json()
            # refresh_token이 응답에 없으면 기존 것 유지
            if 'refresh_token' not in token_data:
                token_data['refresh_token'] = refresh_token
            save_tokens(token_data)
            return token_data
        else:
            print(f'[OAuth] 토큰 갱신 실패: {resp.status_code} {resp.text}')
            return None
    except Exception as e:
        print(f'[OAuth] 토큰 갱신 오류: {e}')
        return None


# ============================================================
# 유효한 액세스 토큰 가져오기
# ============================================================
def get_valid_access_token():
    """유효한 access_token 반환. 필요하면 자동 갱신."""
    token_data = load_tokens()
    if not token_data:
        return None

    if is_token_valid(token_data):
        return token_data['access_token']

    # 갱신 시도
    refresh_token = token_data.get('refresh_token')
    if refresh_token:
        new_data = refresh_access_token(refresh_token)
        if new_data:
            return new_data['access_token']

    return None


# ============================================================
# OAuth 로그인 플로우 (로컬 콜백 서버)
# ============================================================
class OAuthCallbackHandler(BaseHTTPRequestHandler):
    """OAuth 콜백을 처리하는 HTTP 핸들러"""

    auth_code = None
    state_expected = None

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == '/auth/callback':
            params = parse_qs(parsed.query)
            code = params.get('code', [None])[0]
            state = params.get('state', [None])[0]
            error = params.get('error', [None])[0]

            if error:
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.end_headers()
                self.wfile.write(f'''<html><body style="font-family:sans-serif;text-align:center;padding:60px;">
                <h2 style="color:#e74c3c;">로그인 실패</h2>
                <p>오류: {error}</p>
                <p>브라우저를 닫고 다시 시도해주세요.</p>
                </body></html>'''.encode('utf-8'))
                return

            if state != OAuthCallbackHandler.state_expected:
                self.send_response(400)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.end_headers()
                self.wfile.write(b'<html><body><h2>State mismatch</h2></body></html>')
                return

            OAuthCallbackHandler.auth_code = code

            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write('''<html><body style="font-family:sans-serif;text-align:center;padding:60px;">
            <h2 style="color:#6c5ce7;">ChatGPT 로그인 성공!</h2>
            <p>이 창을 닫고 전자책 생성기로 돌아가세요.</p>
            <script>setTimeout(function(){window.close();},3000);</script>
            </body></html>'''.encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        pass  # 로그 출력 안 함


def start_oauth_flow():
    """
    OAuth PKCE 로그인 플로우 시작
    1. 로컬 콜백 서버 시작
    2. 브라우저에서 OpenAI 로그인 페이지 열기
    3. 콜백으로 auth_code 수신
    4. auth_code → access_token 교환
    반환: {'success': True, 'token_data': {...}} 또는 {'success': False, 'error': '...'}
    """
    # PKCE 생성
    code_verifier = generate_code_verifier()
    code_challenge = generate_code_challenge(code_verifier)
    state = secrets.token_urlsafe(32)

    OAuthCallbackHandler.auth_code = None
    OAuthCallbackHandler.state_expected = state

    # 인증 URL 생성
    auth_params = {
        'response_type': 'code',
        'client_id': CLIENT_ID,
        'redirect_uri': REDIRECT_URI,
        'scope': SCOPES,
        'state': state,
        'code_challenge': code_challenge,
        'code_challenge_method': 'S256',
    }
    auth_url = f'{AUTH_ENDPOINT}?{urlencode(auth_params)}'

    # 로컬 콜백 서버 시작
    try:
        server = HTTPServer(('localhost', REDIRECT_PORT), OAuthCallbackHandler)
        server.timeout = 120  # 2분 대기
    except OSError as e:
        return {'success': False, 'error': f'콜백 서버 시작 실패 (포트 {REDIRECT_PORT} 사용 중?): {e}'}

    # 브라우저 열기
    print(f'[OAuth] 브라우저에서 ChatGPT 로그인 페이지를 엽니다...')
    webbrowser.open(auth_url)

    # 콜백 대기
    while OAuthCallbackHandler.auth_code is None:
        server.handle_request()
        if OAuthCallbackHandler.auth_code is None:
            # 타임아웃 확인
            break

    server.server_close()

    auth_code = OAuthCallbackHandler.auth_code
    if not auth_code:
        return {'success': False, 'error': '로그인 시간이 초과되었습니다. 다시 시도해주세요.'}

    # auth_code → access_token 교환
    try:
        resp = requests.post(TOKEN_ENDPOINT, data={
            'grant_type': 'authorization_code',
            'client_id': CLIENT_ID,
            'redirect_uri': REDIRECT_URI,
            'code': auth_code,
            'code_verifier': code_verifier,
        }, timeout=15)

        if resp.status_code == 200:
            token_data = resp.json()
            save_tokens(token_data)
            print('[OAuth] 토큰 교환 성공!')
            return {'success': True, 'token_data': token_data}
        else:
            return {'success': False, 'error': f'토큰 교환 실패: {resp.status_code} {resp.text[:200]}'}
    except Exception as e:
        return {'success': False, 'error': f'토큰 교환 오류: {e}'}


def extract_account_id(access_token):
    """JWT access_token에서 chatgpt_account_id 추출"""
    try:
        # JWT의 payload 부분 디코딩 (서명 검증 없이)
        parts = access_token.split('.')
        if len(parts) < 2:
            return None
        payload = parts[1]
        # base64 패딩 보정
        padding = 4 - len(payload) % 4
        if padding != 4:
            payload += '=' * padding
        decoded = base64.urlsafe_b64decode(payload)
        data = json.loads(decoded)
        # https://api.openai.com/auth 클레임에서 account_id 추출
        auth_info = data.get('https://api.openai.com/auth', {})
        account_id = auth_info.get('account_id')
        if account_id:
            return account_id
        # 다른 위치에서도 시도
        return data.get('account_id') or data.get('sub')
    except Exception:
        return None


def get_login_status():
    """현재 로그인 상태 반환"""
    token_data = load_tokens()
    if not token_data:
        return {'logged_in': False}

    access_token = get_valid_access_token()
    if access_token:
        return {
            'logged_in': True,
            'has_refresh': bool(token_data.get('refresh_token')),
        }
    else:
        return {'logged_in': False, 'expired': True}
