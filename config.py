import os
import sys
import json
import platform

# frozen(EXE) 모드에서는 실행파일 디렉토리에 설정 저장 (영구 보존)
# 개발 모드에서는 소스 디렉토리에 저장
if getattr(sys, 'frozen', False):
    _CONFIG_DIR = os.path.dirname(sys.executable)
else:
    _CONFIG_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(_CONFIG_DIR, 'user_config.json')

_DEFAULT_PROMPT_CHAPTER_SYSTEM = """당신은 전자책 집필 전문가입니다. 사람이 실제로 돈을 내고 살 가치가 있는 상품 수준의 글을 작성합니다.

글쓰기 원칙:
- 지식을 그냥 설명하지 말고, 결과가 보이게 정리
- 읽기 전 상태와 읽고 난 후 상태가 명확히 대비되게
- 설명용 문장은 줄이고 "그래서 뭘 하면 되는지" 실행 단위로 구성
- 한 문단은 3~5줄, 문단 사이 적절한 여백
- 독자가 바로 실행할 수 있는 구체적 행동 지침 포함
- 기계적 나열이 아닌 이야기하듯 자연스러운 흐름
- 실제 사례, 구체적 숫자, 비교 예시를 풍부하게 활용
- 독자가 "이건 내 이야기다"라고 느낄 수 있게 공감 포인트 배치
- 핵심 개념은 반복적으로 다른 표현으로 강조"""

_DEFAULT_PROMPT_TOC_RULES = """목차 설계 규칙:
1. 총 12~16개 챕터 (반드시 12개 이상)
2. 4단계 구조: 문제인식(3~4장) → 방법발견(3~4장) → 실행(4~5장) → 확신(2~3장)
3. 각 챕터 제목은 "이건 안 사면 손해다"라고 느끼게 하는 문장
4. 각 챕터마다 반드시 5~7개 소제목(sections) - 소제목이 적으면 분량 부족
5. 설명용 문장이 아닌, 결과가 보이는 실행 중심 제목"""

_DEFAULT_PROMPT_VALUE_SYSTEM = """당신은 전자책 시장 분석 전문가입니다.
반드시 아래 JSON 형식으로만 응답하세요. 다른 텍스트 없이 순수 JSON만 출력하세요."""

_DEFAULT_PROMPT_MARKETING_SYSTEM = """당신은 전자책 마케팅 전략 전문가입니다. 반드시 JSON 형식으로만 응답하세요."""

# OS별 기본 한국어 폰트 자동 선택
def _default_font():
    if platform.system() == 'Windows':
        return 'MalgunGothic'
    return 'AppleGothic'

DEFAULT_CONFIG = {
    'model': 'gpt-5-codex',
    'image_model': 'gpt-4o',
    'pdf_font': _default_font(),
    'pdf_font_size': 11,
    'pdf_heading_size': 16,
    'pdf_subheading_size': 13,
    'pdf_line_spacing': 1.6,
    'pdf_margin_top': 72,
    'pdf_margin_bottom': 72,
    'pdf_margin_left': 60,
    'pdf_margin_right': 60,
    'output_dir': os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static', 'output'),
    # 목표 페이지 수
    'target_pages_min': 100,
    'target_pages_max': 150,
    # 커스터마이즈 가능한 AI 프롬프트
    'prompt_chapter_system': _DEFAULT_PROMPT_CHAPTER_SYSTEM,
    'prompt_toc_rules': _DEFAULT_PROMPT_TOC_RULES,
    'prompt_value_system': _DEFAULT_PROMPT_VALUE_SYSTEM,
    'prompt_marketing_system': _DEFAULT_PROMPT_MARKETING_SYSTEM,
}

# Windows / Mac 기본 한국어 폰트 경로 매핑
FONT_PATHS = {
    'AppleGothic': '/System/Library/Fonts/Supplemental/AppleGothic.ttf',
    'AppleMyungjo': '/System/Library/Fonts/Supplemental/AppleMyungjo.ttf',
    'NotoSansGothic': '/System/Library/Fonts/Supplemental/NotoSansGothic-Regular.ttf',
    'MalgunGothic': 'C:/Windows/Fonts/malgun.ttf',
    'Batang': 'C:/Windows/Fonts/batang.ttc',
    'Gulim': 'C:/Windows/Fonts/gulim.ttc',
}


def load_config():
    config = DEFAULT_CONFIG.copy()
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
            config.update(saved)
        except Exception:
            pass
    return config


def save_config(updates):
    config = load_config()
    config.update(updates)
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    return config
