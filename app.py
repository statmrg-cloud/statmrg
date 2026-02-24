"""
전자책 자동 생성기 - Flask 서버
ChatGPT OAuth PKCE 인증 기반
다중 출력 형식: PDF, DOCX, PPTX
"""
import os
import json
import threading
import uuid
from flask import Flask, render_template, request, jsonify, send_from_directory
from flask_cors import CORS

from config import load_config, save_config, FONT_PATHS, DEFAULT_CONFIG, _DEFAULT_PROMPT_CHAPTER_SYSTEM, _DEFAULT_PROMPT_TOC_RULES, _DEFAULT_PROMPT_VALUE_SYSTEM, _DEFAULT_PROMPT_MARKETING_SYSTEM
from modules.ai_engine import generate_ebook
from modules.pdf_generator import EbookPDFGenerator, EbookDocxGenerator, EbookPptxGenerator
from modules.hwpx_generator import EbookHwpxGenerator
from modules.oauth import (
    start_oauth_flow, get_login_status, get_valid_access_token, clear_tokens
)

app = Flask(__name__,
            static_folder='static',
            template_folder='templates')
CORS(app)

# 진행 상태 저장소
progress_store = {}
result_store = {}


# ============================================================
# 페이지 라우트
# ============================================================
@app.route('/')
def index():
    config = load_config()
    login = get_login_status()
    return render_template('index.html', logged_in=login.get('logged_in', False), config=config)


@app.route('/settings')
def settings_page():
    config = load_config()
    fonts = list(FONT_PATHS.keys())
    login = get_login_status()
    return render_template('settings.html', config=config, fonts=fonts, logged_in=login.get('logged_in', False))


@app.route('/result/<task_id>')
def result_page(task_id):
    return render_template('result.html', task_id=task_id)


@app.route('/edit/<task_id>')
def edit_page(task_id):
    return render_template('edit.html', task_id=task_id)


# ============================================================
# OAuth API
# ============================================================
@app.route('/api/auth/status')
def auth_status():
    login = get_login_status()
    return jsonify({'success': True, **login})


@app.route('/api/auth/login', methods=['POST'])
def auth_login():
    def do_login():
        start_oauth_flow()

    thread = threading.Thread(target=do_login, daemon=True)
    thread.start()
    return jsonify({
        'success': True,
        'message': '브라우저에서 ChatGPT 로그인 페이지가 열립니다. 로그인을 완료해주세요.'
    })


@app.route('/api/auth/logout', methods=['POST'])
def auth_logout():
    clear_tokens()
    return jsonify({'success': True, 'message': '로그아웃되었습니다.'})


# ============================================================
# 설정 API
# ============================================================
@app.route('/api/config', methods=['GET'])
def get_config():
    config = load_config()
    safe = {k: v for k, v in config.items() if k != 'openai_api_key'}
    return jsonify({'success': True, 'config': safe})


@app.route('/api/config', methods=['POST'])
def update_config():
    data = request.json
    updates = {}
    allowed = [
        'model', 'image_model',
        'pdf_font', 'pdf_font_size', 'pdf_heading_size', 'pdf_subheading_size',
        'pdf_line_spacing', 'pdf_margin_top', 'pdf_margin_bottom',
        'pdf_margin_left', 'pdf_margin_right',
        'target_pages_min', 'target_pages_max',
        'prompt_chapter_system', 'prompt_toc_rules',
        'prompt_value_system', 'prompt_marketing_system',
    ]
    for key in allowed:
        if key in data:
            val = data[key]
            if key in ('pdf_font_size', 'pdf_heading_size', 'pdf_subheading_size',
                       'pdf_margin_top', 'pdf_margin_bottom', 'pdf_margin_left', 'pdf_margin_right',
                       'target_pages_min', 'target_pages_max'):
                val = int(float(val))
            elif key == 'pdf_line_spacing':
                val = float(val)
            updates[key] = val

    save_config(updates)
    return jsonify({'success': True, 'message': '설정이 저장되었습니다.'})


@app.route('/api/config/reset_prompts', methods=['POST'])
def reset_prompts():
    defaults = {
        'prompt_chapter_system': _DEFAULT_PROMPT_CHAPTER_SYSTEM,
        'prompt_toc_rules': _DEFAULT_PROMPT_TOC_RULES,
        'prompt_value_system': _DEFAULT_PROMPT_VALUE_SYSTEM,
        'prompt_marketing_system': _DEFAULT_PROMPT_MARKETING_SYSTEM,
    }
    save_config(defaults)
    return jsonify({'success': True, 'defaults': defaults})


# ============================================================
# 전자책 생성 API
# ============================================================
@app.route('/api/generate', methods=['POST'])
def start_generation():
    """전자책 생성 시작 (백그라운드)"""
    data = request.json
    topic = data.get('topic', '').strip()
    include_images = data.get('include_images', True)
    output_formats = data.get('output_formats', ['pdf'])  # 선택한 출력 형식들

    if not topic:
        return jsonify({'success': False, 'error': '주제/키워드를 입력해주세요.'})

    token = get_valid_access_token()
    if not token:
        return jsonify({'success': False, 'error': 'ChatGPT 로그인이 필요합니다. 먼저 로그인해주세요.'})

    config = load_config()
    model = config.get('model', 'gpt-5-codex')
    task_id = str(uuid.uuid4())[:8]

    progress_store[task_id] = {
        'status': 'started',
        'step': 0,
        'total_steps': 6,
        'message': '준비 중...',
        'topic': topic,
    }

    def run_generation():
        def on_progress(step, total, msg, data=None):
            progress_store[task_id] = {
                'status': 'running',
                'step': step,
                'total_steps': total,
                'message': msg,
                'topic': topic,
            }

        try:
            ebook_data = generate_ebook(model, topic, include_images, on_progress, config=config)

            if ebook_data.get('error'):
                progress_store[task_id] = {
                    'status': 'error',
                    'message': f"생성 실패: {ebook_data['error']}",
                    'topic': topic,
                }
                result_store[task_id] = ebook_data
                return

            ebook_data['generated_files'] = {}

            # PDF 생성
            if 'pdf' in output_formats:
                progress_store[task_id]['message'] = 'PDF 파일 생성 중...'
                try:
                    generator = EbookPDFGenerator(config)
                    filepath, filename = generator.generate(ebook_data)
                    ebook_data['generated_files']['pdf'] = filename
                    ebook_data['pdf_filename'] = filename  # 하위 호환
                except Exception as e:
                    print(f"[PDF 생성 오류] {e}")
                    ebook_data['generated_files']['pdf'] = None

            # DOCX 생성
            if 'docx' in output_formats:
                progress_store[task_id]['message'] = 'DOCX 파일 생성 중...'
                try:
                    gen = EbookDocxGenerator(config)
                    filepath, filename = gen.generate(ebook_data)
                    ebook_data['generated_files']['docx'] = filename
                except Exception as e:
                    print(f"[DOCX 생성 오류] {e}")
                    ebook_data['generated_files']['docx'] = None

            # PPTX 생성
            if 'pptx' in output_formats:
                progress_store[task_id]['message'] = 'PPTX 파일 생성 중...'
                try:
                    gen = EbookPptxGenerator(config)
                    filepath, filename = gen.generate(ebook_data)
                    ebook_data['generated_files']['pptx'] = filename
                except Exception as e:
                    print(f"[PPTX 생성 오류] {e}")
                    ebook_data['generated_files']['pptx'] = None

            # HWPX 생성
            if 'hwpx' in output_formats:
                progress_store[task_id]['message'] = '한글 파일(HWPX) 생성 중...'
                try:
                    gen = EbookHwpxGenerator(config)
                    filepath, filename = gen.generate(ebook_data)
                    ebook_data['generated_files']['hwpx'] = filename
                except Exception as e:
                    print(f"[HWPX 생성 오류] {e}")
                    ebook_data['generated_files']['hwpx'] = None

            result_store[task_id] = ebook_data

            progress_store[task_id] = {
                'status': 'completed',
                'step': progress_store[task_id].get('total_steps', 6),
                'total_steps': progress_store[task_id].get('total_steps', 6),
                'message': '전자책 생성 완료!',
                'topic': topic,
                'pdf_filename': ebook_data.get('pdf_filename'),
            }

        except Exception as e:
            import traceback
            traceback.print_exc()
            progress_store[task_id] = {
                'status': 'error',
                'message': f'오류 발생: {str(e)}',
                'topic': topic,
            }

    thread = threading.Thread(target=run_generation, daemon=True)
    thread.start()

    return jsonify({'success': True, 'task_id': task_id, 'message': '전자책 생성을 시작합니다.'})


@app.route('/api/progress/<task_id>')
def get_progress(task_id):
    data = progress_store.get(task_id)
    if not data:
        return jsonify({'success': False, 'error': '작업을 찾을 수 없습니다.'})
    return jsonify({'success': True, 'data': data})


@app.route('/api/result/<task_id>')
def get_result(task_id):
    data = result_store.get(task_id)
    if not data:
        return jsonify({'success': False, 'error': '결과를 찾을 수 없습니다.'})
    safe = {
        'topic': data.get('topic'),
        'analysis': data.get('analysis'),
        'book_info': data.get('book_info'),
        'marketing': data.get('marketing'),
        'cover_url': data.get('cover_url'),
        'pdf_filename': data.get('pdf_filename'),
        'generated_files': data.get('generated_files', {}),
        'chapters_count': len(data.get('chapters_content', [])),
        'error': data.get('error'),
    }
    return jsonify({'success': True, 'data': safe})


@app.route('/api/download/<filename>')
def download_file(filename):
    config = load_config()
    output_dir = config.get('output_dir', os.path.join(os.path.dirname(__file__), 'static', 'output'))
    return send_from_directory(output_dir, filename, as_attachment=True)


# ============================================================
# 편집 API - 전자책 내용 조회 및 저장
# ============================================================
@app.route('/api/edit/<task_id>', methods=['GET'])
def get_edit_content(task_id):
    """편집용 전자책 전체 콘텐츠 반환"""
    data = result_store.get(task_id)
    if not data:
        return jsonify({'success': False, 'error': '결과를 찾을 수 없습니다.'})
    content = {
        'topic': data.get('topic'),
        'book_info': data.get('book_info'),
        'chapters_content': data.get('chapters_content', []),
        'analysis': data.get('analysis'),
        'marketing': data.get('marketing'),
    }
    return jsonify({'success': True, 'data': content})


@app.route('/api/edit/<task_id>', methods=['POST'])
def save_edit_content(task_id):
    """편집된 내용 저장 후 재생성"""
    if task_id not in result_store:
        return jsonify({'success': False, 'error': '결과를 찾을 수 없습니다.'})

    updates = request.json
    data = result_store[task_id]

    # 편집 내용 반영
    if 'book_info' in updates:
        data['book_info'].update(updates['book_info'])
    if 'chapters_content' in updates:
        chapters_content = updates['chapters_content']
        for i, ch_update in enumerate(chapters_content):
            if i < len(data['chapters_content']):
                if 'content' in ch_update:
                    data['chapters_content'][i]['content'] = ch_update['content']
                if 'chapter' in ch_update:
                    data['chapters_content'][i]['chapter'].update(ch_update['chapter'])

    config = load_config()
    output_formats = updates.get('output_formats', ['pdf'])
    data['generated_files'] = {}

    # 재생성
    if 'pdf' in output_formats:
        try:
            gen = EbookPDFGenerator(config)
            filepath, filename = gen.generate(data)
            data['generated_files']['pdf'] = filename
            data['pdf_filename'] = filename
        except Exception as e:
            print(f"[재생성 PDF 오류] {e}")

    if 'docx' in output_formats:
        try:
            gen = EbookDocxGenerator(config)
            filepath, filename = gen.generate(data)
            data['generated_files']['docx'] = filename
        except Exception as e:
            print(f"[재생성 DOCX 오류] {e}")

    if 'pptx' in output_formats:
        try:
            gen = EbookPptxGenerator(config)
            filepath, filename = gen.generate(data)
            data['generated_files']['pptx'] = filename
        except Exception as e:
            print(f"[재생성 PPTX 오류] {e}")

    if 'hwpx' in output_formats:
        try:
            gen = EbookHwpxGenerator(config)
            filepath, filename = gen.generate(data)
            data['generated_files']['hwpx'] = filename
        except Exception as e:
            print(f"[재생성 HWPX 오류] {e}")

    result_store[task_id] = data
    return jsonify({
        'success': True,
        'message': '저장 및 재생성 완료',
        'generated_files': data.get('generated_files', {}),
    })


@app.route('/api/regenerate_format/<task_id>', methods=['POST'])
def regenerate_format(task_id):
    """특정 형식만 재생성"""
    if task_id not in result_store:
        return jsonify({'success': False, 'error': '결과를 찾을 수 없습니다.'})

    fmt = request.json.get('format', 'pdf')
    data = result_store[task_id]
    config = load_config()

    try:
        if fmt == 'pdf':
            gen = EbookPDFGenerator(config)
        elif fmt == 'docx':
            gen = EbookDocxGenerator(config)
        elif fmt == 'pptx':
            gen = EbookPptxGenerator(config)
        elif fmt == 'hwpx':
            gen = EbookHwpxGenerator(config)
        else:
            return jsonify({'success': False, 'error': f'지원하지 않는 형식: {fmt}'})

        filepath, filename = gen.generate(data)
        if 'generated_files' not in data:
            data['generated_files'] = {}
        data['generated_files'][fmt] = filename
        if fmt == 'pdf':
            data['pdf_filename'] = filename
        result_store[task_id] = data
        return jsonify({'success': True, 'filename': filename})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


# ============================================================
# 테스트 모드 API
# ============================================================
@app.route('/api/test_generate', methods=['POST'])
def test_generate():
    """더미 데이터로 즉시 모든 형식 생성 (AI 호출 없음)"""
    config = load_config()
    task_id = 'test_' + str(uuid.uuid4())[:6]

    ebook_data = {
        'topic': '직장인 퇴근 후 월 100만원 부수입 만들기',
        'book_info': {
            'book_title': '퇴근 후 100만원 만들기',
            'subtitle': '직장인을 위한 부수입 실전 가이드',
            'author': 'AI 전자책 생성기',
            'chapters': [
                {'chapter_num': 1, 'title': '왜 지금 부수입이 필요한가', 'phase': '문제인식'},
                {'chapter_num': 2, 'title': '나에게 맞는 부수입 방법 찾기', 'phase': '방법발견'},
                {'chapter_num': 3, 'title': '첫 달 수입 만들기: 실전 30일', 'phase': '실행'},
                {'chapter_num': 4, 'title': '수입을 안정화하고 확장하기', 'phase': '확신'},
            ],
            'reader_psychology': {
                'concerns': ['시간이 없다', '무엇부터 시작할지 모른다', '실패할까 두렵다'],
                'expectations': ['안정적인 추가 수입', '본업에 지장 없는 방법', '검증된 노하우'],
                'fears': ['사기당할까봐', '시간 낭비가 될까봐', '가족에게 부담이 될까봐'],
            },
        },
        'chapters_content': [
            {
                'chapter': {
                    'chapter_num': 1, 'title': '왜 지금 부수입이 필요한가',
                    'phase': '문제인식',
                    'before_state': '월급만으로는 부족하다는 막막함',
                    'after_state': '부수입의 필요성과 가능성을 명확히 이해',
                },
                'content': (
                    "== 월급의 한계 ==\n\n"
                    "현대 직장인의 가장 큰 고민은 '월급으로는 부족하다'는 현실입니다. "
                    "물가는 오르고 생활비는 늘지만 월급 인상률은 이를 따라가지 못합니다.\n\n"
                    "[핵심 포인트] 월급만 바라보는 삶의 위험성\n\n"
                    "- 물가 상승률 연 3~5%, 평균 임금 인상률은 2~3%\n"
                    "- 하나의 수입원에 의존하면 갑작스런 실직에 취약\n"
                    "- 부수입 없이는 노후 준비가 사실상 불가능\n\n"
                    "== 부수입이 바꾸는 삶 ==\n\n"
                    "부수입 월 100만원은 단순한 추가 금액이 아닙니다. "
                    "심리적 안정감과 선택의 자유를 가져다줍니다.\n\n"
                    "[실전 팁] 100만원의 진짜 의미\n\n"
                    "연간 1,200만원은 20년이면 2억 4천만원 이상의 차이를 만들어냅니다.\n\n"
                    "- 복리 효과: 일찍 시작할수록 결과가 크다\n"
                    "- 스킬 축적: 시간이 지날수록 더 쉬워진다\n"
                    "- 네트워크 형성: 부수입 활동을 통한 인맥 확장"
                ),
            },
            {
                'chapter': {
                    'chapter_num': 2, 'title': '나에게 맞는 부수입 방법 찾기',
                    'phase': '방법발견',
                    'before_state': '어떤 방법이 맞는지 몰라 망설임',
                    'after_state': '자신의 상황에 최적화된 방법 선택 완료',
                },
                'content': (
                    "== 디지털 부수입 ==\n\n"
                    "인터넷만 있으면 시작할 수 있는 방법들이 폭발적으로 늘었습니다.\n\n"
                    "- 전자책 출판: 전문 지식을 PDF로 판매\n"
                    "- 온라인 강의: 유데미, 클래스101 플랫폼 활용\n"
                    "- 블로그/유튜브: 광고 수익과 협찬\n"
                    "- 프리랜서: 크몽, 탈잉에서 전문 스킬 판매\n\n"
                    "== 오프라인 부수입 ==\n\n"
                    "디지털이 익숙하지 않다면 오프라인에서도 기회가 충분합니다.\n\n"
                    "- 중고 거래: 집 안 물건 정리하며 수입 창출\n"
                    "- 재능 공유: 요리, 운동, 악기 개인 레슨\n"
                    "- 단기 알바: 주말 이벤트 스태프\n\n"
                    "[핵심 포인트] 선택 기준 4가지\n\n"
                    "1. 현재 보유 스킬과 지식\n"
                    "2. 투자 가능 시간 (주당 5~10시간)\n"
                    "3. 초기 투자 비용 (0~50만원)\n"
                    "4. 빠른 수익화 가능 여부"
                ),
            },
            {
                'chapter': {
                    'chapter_num': 3, 'title': '첫 달 수입 만들기: 실전 30일',
                    'phase': '실행',
                    'before_state': '계획만 있고 실행을 못하는 상태',
                    'after_state': '첫 번째 수입을 실제로 만들어낸 경험',
                },
                'content': (
                    "== 30일 행동 계획 ==\n\n"
                    "이론보다 실행이 중요합니다. 다음 계획을 그대로 따라하면 첫 수입을 만들 수 있습니다.\n\n"
                    "[1주차] 기반 다지기\n\n"
                    "- 판매할 상품/서비스 1개 확정\n"
                    "- 플랫폼 계정 생성 및 프로필 완성\n"
                    "- 첫 포스팅 또는 상품 등록\n\n"
                    "[2주차] 첫 고객 만들기\n\n"
                    "- 지인 10명에게 소개\n"
                    "- SNS 홍보 시작\n"
                    "- 첫 피드백 수집 및 개선\n\n"
                    "[3주차] 수익 창출\n\n"
                    "- 가격 정책 최적화\n"
                    "- 반복 구매 유도 전략 실행\n"
                    "- 리뷰 수집 및 신뢰 구축\n\n"
                    "[4주차] 확장\n\n"
                    "- 성과 분석 및 개선점 파악\n"
                    "- 두 번째 상품/서비스 기획\n"
                    "- 자동화 시스템 구축 시작"
                ),
            },
            {
                'chapter': {
                    'chapter_num': 4, 'title': '수입을 안정화하고 확장하기',
                    'phase': '확신',
                    'before_state': '불규칙한 수입으로 인한 불안감',
                    'after_state': '안정적이고 성장하는 부수입 시스템 완성',
                },
                'content': (
                    "== 수입 안정화 전략 ==\n\n"
                    "첫 수입을 만들었다면 이제 안정화하고 성장시켜야 합니다.\n\n"
                    "[핵심 포인트] 패시브 인컴으로의 전환\n\n"
                    "처음에는 시간을 팔지만, 궁극적으로는 자동으로 수입이 들어오는 구조를 만들어야 합니다.\n\n"
                    "== 3단계 성장 로드맵 ==\n\n"
                    "1단계: 활성 수입 (시간 = 돈)\n"
                    "- 프리랜서, 레슨, 단기 알바 → 월 30~50만원\n\n"
                    "2단계: 반패시브 수입\n"
                    "- 전자책, 온라인 강의, 블로그 → 월 70~100만원\n\n"
                    "3단계: 패시브 수입\n"
                    "- 배당주, 부동산 수익, 로열티 → 월 100만원 이상\n\n"
                    "[실전 팁] 수입 다각화\n\n"
                    "- 최소 3개 이상의 수입원 유지\n"
                    "- 디지털 + 오프라인 조합 권장\n"
                    "- 월 1회 성과 리뷰 및 전략 수정"
                ),
            },
        ],
        'analysis': {
            'free_vs_paid': {
                'verdict': '유료 전환 강력 추천',
                'paid_conversion_points': [
                    '검증된 실전 경험과 노하우 담김',
                    '30일 행동 계획 등 즉시 실행 가능한 콘텐츠',
                    '실패 사례와 해결책 포함',
                ],
            },
            'problem_solved': {
                'time': '수개월의 시행착오를 수주로 단축',
                'money': '잘못된 투자 방지로 수십만원 절약',
                'emotion': '막막함에서 명확한 로드맵으로',
            },
            'why_pay': '이 책은 수천 시간의 실전 경험을 압축한 결과물입니다.',
            'target_reader': '부수입을 원하지만 어디서 시작할지 모르는 20~40대 직장인',
        },
        'marketing': {
            'sales_copy': '퇴근 후 2시간으로 월 100만원을 만드는 직장인들의 비밀을 공개합니다.',
            'content_topics': [
                {'topic': '월급만으로 부족한 직장인의 현실', 'hook_sentence': '"월급날이 두렵다면, 당신만 그런 게 아닙니다"'},
                {'topic': '퇴근 후 2시간 활용법', 'hook_sentence': '"하루 2시간이 1년 후 재정을 바꿉니다"'},
            ],
            'natural_distribution': {
                'blog_questions': ['직장인 부업 추천', '월급 외 수입 만들기'],
                'community_complaints': ['월급이 너무 적어요', '부업 뭐가 좋을까요'],
                'sns_consumption': ['수입 인증', '부업 성공 스토리'],
            },
            'value_summary': {
                'time_saved': '6개월 시행착오 단축',
                'money_saved': '잘못된 투자 30만원 절약',
                'mistakes_prevented': '흔한 실수 10가지 예방',
            },
        },
        'cover_url': None,
        'generated_files': {},
    }

    errors = {}
    for fmt, GenClass in [
        ('pdf',  EbookPDFGenerator),
        ('docx', EbookDocxGenerator),
        ('pptx', EbookPptxGenerator),
        ('hwpx', EbookHwpxGenerator),
    ]:
        try:
            gen = GenClass(config)
            _, filename = gen.generate(ebook_data)
            ebook_data['generated_files'][fmt] = filename
            if fmt == 'pdf':
                ebook_data['pdf_filename'] = filename
        except Exception as e:
            import traceback; traceback.print_exc()
            errors[fmt] = str(e)

    result_store[task_id] = ebook_data
    progress_store[task_id] = {
        'status': 'completed', 'step': 6, 'total_steps': 6,
        'message': '테스트 생성 완료!', 'topic': ebook_data['topic'],
    }

    return jsonify({
        'success': True,
        'task_id': task_id,
        'generated': list(ebook_data['generated_files'].keys()),
        'errors': errors,
    })


# ============================================================
# 실행
# ============================================================
if __name__ == '__main__':
    config = load_config()
    os.makedirs(config.get('output_dir', './static/output'), exist_ok=True)
    print(f"\n  전자책 자동 생성기 실행 중!")
    print(f"  http://localhost:5000")
    print(f"  ChatGPT OAuth 로그인 방식\n")
    app.run(host='0.0.0.0', port=5000, debug=False)
