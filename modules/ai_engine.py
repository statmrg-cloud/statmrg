"""
AI 전자책 생성 엔진
12개 프롬프트 원칙을 단계별로 실행하여 상품성 있는 전자책을 생성
ChatGPT OAuth 토큰 → chatgpt.com/backend-api/codex/responses (SSE stream)
"""
import json
import re
import time
import uuid
import traceback
import requests as http_requests
from modules.oauth import get_valid_access_token, extract_account_id


# ============================================================
# ChatGPT Codex Backend API
# ============================================================
CHATGPT_API_URL = 'https://chatgpt.com/backend-api/codex/responses'


def _build_headers(stream=True):
    """ChatGPT Codex 백엔드 API용 헤더"""
    token = get_valid_access_token()
    if not token:
        raise RuntimeError('ChatGPT 로그인이 필요합니다. 먼저 로그인해주세요.')

    cache_id = str(uuid.uuid4())
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
        'Accept': 'text/event-stream' if stream else 'application/json',
        'version': '0.101.0',
        'user-agent': 'codex_cli_rs/0.101.0',
        'originator': 'codex_cli_rs',
        'Conversation_id': cache_id,
        'Session_id': cache_id,
    }
    account_id = extract_account_id(token)
    if account_id:
        headers['chatgpt-account-id'] = account_id
    return headers


def _parse_sse_text(resp):
    """SSE 스트림에서 최종 텍스트 추출 (response.completed 또는 delta 누적)"""
    full_text = ''
    for raw_line in resp.iter_lines():
        # iter_lines()는 bytes를 반환할 수 있으므로 항상 str로 변환
        if isinstance(raw_line, bytes):
            line = raw_line.decode('utf-8', errors='replace')
        else:
            line = raw_line
        if not line or not line.startswith('data: '):
            continue
        payload = line[6:].strip()
        if payload == '[DONE]':
            break
        try:
            event = json.loads(payload)
            event_type = event.get('type', '')

            # response.completed → 최종 응답에서 텍스트 추출
            if event_type == 'response.completed':
                response_obj = event.get('response', event)
                text = _extract_text_from_response(response_obj)
                if text:
                    return text.strip()

            # output_text.delta → 점진적 텍스트 수집
            elif event_type == 'output_text.delta':
                full_text += event.get('delta', '')

            # response.output_text.done → 한 output_text 블록 완료
            elif event_type == 'response.output_text.done':
                t = event.get('text', '')
                if t:
                    full_text = t  # 최종 완성본으로 교체

        except json.JSONDecodeError:
            continue

    return full_text.strip() if full_text else None


def _extract_text_from_response(data):
    """Responses API 응답 객체에서 텍스트 추출"""
    # output_text 직접 필드
    if data.get('output_text'):
        return data['output_text']

    # output 배열 탐색
    for item in data.get('output', []):
        if item.get('type') == 'message':
            for c in item.get('content', []):
                if c.get('type') == 'output_text':
                    return c.get('text', '')
        if item.get('type') == 'text':
            return item.get('text', '')

    # text 직접 필드
    if 'text' in data:
        return data['text']

    return None


# ============================================================
# 핵심 API 호출 함수 (항상 stream=true)
# ============================================================
def call_gpt(client_unused, model, system_prompt, user_prompt, temperature=0.7, max_tokens=4096):
    """ChatGPT Codex 백엔드 API 호출 (SSE stream, 재시도 포함)"""

    body = {
        'model': model,
        'instructions': system_prompt,
        'stream': True,
        'store': False,
        'prompt_cache_key': str(uuid.uuid4()),
        'input': [
            {'type': 'message', 'role': 'user', 'content': user_prompt},
        ],
    }

    for attempt in range(3):
        try:
            headers = _build_headers(stream=True)
            resp = http_requests.post(
                CHATGPT_API_URL,
                headers=headers,
                json=body,
                timeout=300,
                stream=True,
            )

            if resp.status_code == 401:
                if attempt < 2:
                    time.sleep(2)
                    continue
                raise RuntimeError(f'인증 실패 (401): {resp.text[:300]}')

            if resp.status_code != 200:
                error_text = ''
                try:
                    error_text = resp.text[:500]
                except Exception:
                    error_text = f'status {resp.status_code}'
                if attempt < 2:
                    time.sleep(3)
                    continue
                raise RuntimeError(f'API 오류 ({resp.status_code}): {error_text}')

            text = _parse_sse_text(resp)
            if text:
                return text

            if attempt < 2:
                time.sleep(2)
                continue
            raise RuntimeError('응답에서 텍스트를 추출할 수 없습니다.')

        except http_requests.exceptions.Timeout:
            if attempt < 2:
                time.sleep(3)
                continue
            raise RuntimeError('API 요청 시간 초과 (300초)')
        except RuntimeError:
            raise
        except Exception as e:
            if attempt < 2:
                time.sleep(2)
                continue
            raise e

    raise RuntimeError('API 호출 실패 (최대 재시도 초과)')


def call_gpt_json(client, model, system_prompt, user_prompt, temperature=0.7, max_tokens=4096):
    """GPT API 호출 → JSON 파싱"""
    raw = call_gpt(client, model, system_prompt, user_prompt, temperature, max_tokens)
    # ```json ... ``` 블록 제거
    cleaned = raw
    if cleaned.startswith('```'):
        cleaned = cleaned.split('\n', 1)[-1] if '\n' in cleaned else cleaned[3:]
    if cleaned.endswith('```'):
        cleaned = cleaned[:-3]
    cleaned = cleaned.strip()
    if cleaned.startswith('```json'):
        cleaned = cleaned[7:]
    if cleaned.startswith('```'):
        cleaned = cleaned[3:]
    cleaned = cleaned.strip()

    # JSON 객체/배열 추출
    match = re.search(r'[\[{][\s\S]*[\]}]', cleaned)
    if match:
        cleaned = match.group(0)

    return json.loads(cleaned)


# ============================================================
# 1단계: 유료 가치 판단
# ============================================================
def step1_value_analysis(client, model, topic, config=None):
    cfg = config or {}
    system = cfg.get('prompt_value_system') or """당신은 전자책 시장 분석 전문가입니다.
반드시 아래 JSON 형식으로만 응답하세요. 다른 텍스트 없이 순수 JSON만 출력하세요."""

    user = f"""다음 주제/키워드에 대해 분석하세요: "{topic}"

JSON 형식:
{{
  "topic_refined": "정제된 주제 (구체적이고 판매 가능한 형태로)",
  "free_vs_paid": {{
    "free_consumption_risk": "무료로 소비될 확률과 그 이유 (솔직하게)",
    "paid_conversion_points": ["유료 전환 가능 포인트 1", "포인트 2", "포인트 3"],
    "verdict": "유료 가치 판단 결론 (1~2문장)"
  }},
  "problem_solved": {{
    "time": "이 전자책이 독자의 시간을 얼마나 줄여주는지 (구체적 수치/상황)",
    "money": "독자의 돈을 얼마나 아껴주는지 (구체적 수치/상황)",
    "emotion": "독자의 어떤 감정적 고통을 해결해주는지"
  }},
  "why_pay": "왜 돈 주고 사야 하는지 (구조적 설명, 2~3문장)",
  "target_reader": "이 책을 살 사람의 구체적 프로필 (1~2문장)",
  "time_saved_hours": "이 책이 절약해주는 시간 (숫자, 시간 단위)",
  "money_saved_won": "이 책이 절약해주는 비용 (숫자, 원 단위)",
  "mistakes_prevented": "이 책이 방지해주는 실수 (숫자)개"
}}"""

    return call_gpt_json(client, model, system, user, temperature=0.6, max_tokens=2000)


# ============================================================
# 2단계: 독자 심리 분석 + 목차 설계
# ============================================================
def step2_toc_design(client, model, topic, analysis, config=None):
    cfg = config or {}
    system = """당신은 베스트셀러 전자책 기획 전문가입니다.
구매자 관점에서 목차를 설계합니다. 반드시 JSON 형식으로만 응답하세요."""

    user = f"""주제: "{topic}"
분석 결과: {json.dumps(analysis, ensure_ascii=False)}

아래 JSON 형식으로 응답하세요:
{{
  "book_title": "전자책 제목 (구매 욕구를 자극하는, 30자 내외)",
  "subtitle": "부제목 (구체적 결과를 약속하는)",
  "reader_psychology": {{
    "concerns": ["구매 전 고민 1", "고민 2", "고민 3"],
    "expectations": ["기대 1", "기대 2", "기대 3"],
    "fears": ["두려움 1", "두려움 2", "두려움 3"]
  }},
  "chapters": [
    {{
      "phase": "문제인식",
      "chapter_num": 1,
      "title": "안 사면 손해라고 느끼게 하는 챕터 제목",
      "purpose": "이 챕터를 읽으면 독자가 얻는 것",
      "before_state": "읽기 전 독자 상태",
      "after_state": "읽고 난 후 독자 상태",
      "sections": ["소제목1", "소제목2", "소제목3"]
    }}
  ]
}}

목표 페이지 수: {cfg.get('target_pages_min', 100)}~{cfg.get('target_pages_max', 150)}페이지

{cfg.get('prompt_toc_rules') or '''목차 설계 규칙:
1. 총 12~16개 챕터 (반드시 12개 이상)
2. 4단계 구조: 문제인식(3~4장) → 방법발견(3~4장) → 실행(4~5장) → 확신(2~3장)
3. 각 챕터 제목은 "이건 안 사면 손해다"라고 느끼게 하는 문장
4. 각 챕터마다 반드시 5~7개 소제목(sections) - 소제목이 적으면 분량 부족
5. 설명용 문장이 아닌, 결과가 보이는 실행 중심 제목'''}"""

    return call_gpt_json(client, model, system, user, temperature=0.7, max_tokens=4000)


# ============================================================
# 3단계: 챕터별 본문 작성
# ============================================================
def step3_write_chapter(client, model, topic, book_info, chapter, chapter_idx, total_chapters, config=None):
    """챕터를 2회 호출(전반부/후반부)로 나눠 풍부한 분량 확보"""
    cfg = config or {}
    sections = chapter.get('sections', [])
    mid = max(1, len(sections) // 2)
    sections_first = sections[:mid]
    sections_second = sections[mid:]

    # 목표 페이지 수에 따라 소제목당 최소 분량 계산 (A4 1페이지 ≈ 550자)
    target_min = int(cfg.get('target_pages_min', 100))
    target_max = int(cfg.get('target_pages_max', 150))
    total_chapters_count = max(total_chapters, 1)
    # 표지+목차 페이지 제외 후 챕터당 평균 페이지
    content_pages = (target_min + target_max) // 2 - 4
    pages_per_chapter = max(6, content_pages // total_chapters_count)
    chars_per_chapter = pages_per_chapter * 550
    # 소제목당 최소 자수 (전/후반 각각)
    sections_per_half = max(1, len(sections_first))
    min_chars_per_half = max(3000, chars_per_chapter // 2)

    system = cfg.get('prompt_chapter_system') or """당신은 전자책 집필 전문가입니다. 사람이 실제로 돈을 내고 살 가치가 있는 상품 수준의 글을 작성합니다.

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

    # 1차: 도입부 + 전반부 소제목
    user1 = f"""전자책: "{book_info.get('book_title', topic)}"
현재 챕터: {chapter_idx + 1}/{total_chapters}

챕터 정보:
- 단계: {chapter.get('phase', '')}
- 제목: {chapter.get('title', '')}
- 목적: {chapter.get('purpose', '')}
- 읽기 전: {chapter.get('before_state', '')}
- 읽고 후: {chapter.get('after_state', '')}

지금 작성할 소제목들 (전반부): {json.dumps(sections_first, ensure_ascii=False)}

규칙:
1. 먼저 챕터 도입부(300~500자)를 작성하세요. 독자의 현재 상태를 공감하며 시작합니다.
2. 각 소제목 시작 시 === 소제목 === 형태로 구분
3. 소제목마다 최소 1500~2500자 분량으로 깊이있게 작성
4. 각 소제목 안에 실제 사례, 구체적 숫자, 비교 예시를 반드시 포함
5. 각 소제목마다 [핵심 포인트] 또는 [실전 팁] 박스를 1개 이상 넣으세요
6. 반드시 한국어로 작성
7. 최소 {min_chars_per_half}자 이상 작성"""

    part1 = call_gpt(client, model, system, user1, temperature=0.7, max_tokens=8000)

    # 2차: 후반부 소제목 + 요약 + 체크리스트
    user2 = f"""전자책: "{book_info.get('book_title', topic)}"
현재 챕터: {chapter_idx + 1}/{total_chapters}
챕터 제목: {chapter.get('title', '')}

앞서 작성된 전반부 내용에 이어서 후반부를 작성합니다.

지금 작성할 소제목들 (후반부): {json.dumps(sections_second, ensure_ascii=False)}

규칙:
1. 각 소제목 시작 시 === 소제목 === 형태로 구분
2. 소제목마다 최소 1500~2500자 분량으로 깊이있게 작성
3. 각 소제목 안에 실제 사례, 구체적 숫자, 비교 예시를 반드시 포함
4. 각 소제목마다 [핵심 포인트] 또는 [실전 팁] 박스를 1개 이상 넣으세요
5. 마지막에 다음을 추가하세요:
   === 핵심 요약 ===
   이 챕터에서 배운 5가지 핵심 내용을 정리 (각 2~3문장)

   === 실행 체크리스트 ===
   독자가 바로 실행할 수 있는 5~7가지 구체적 행동 목록
6. 반드시 한국어로 작성
7. 최소 {min_chars_per_half}자 이상 작성"""

    part2 = call_gpt(client, model, system, user2, temperature=0.7, max_tokens=8000)

    return part1 + '\n\n' + part2


# ============================================================
# 4단계: 자연 유통 분석 + 판매 소개문
# ============================================================
def step4_marketing(client, model, topic, analysis, book_info, config=None):
    cfg = config or {}
    system = cfg.get('prompt_marketing_system') or """당신은 전자책 마케팅 전략 전문가입니다. 반드시 JSON 형식으로만 응답하세요."""

    user = f"""전자책: "{book_info.get('book_title', topic)}"
분석: {json.dumps(analysis, ensure_ascii=False)}

JSON 형식으로 응답:
{{
  "natural_distribution": {{
    "blog_questions": ["블로그에서 이 주제 관련 자주 올라오는 질문 1", "질문 2", "질문 3"],
    "community_complaints": ["커뮤니티 불만/후기 형태 1", "형태 2", "형태 3"],
    "sns_consumption": ["SNS에서 소비되는 형태 1", "형태 2", "형태 3"]
  }},
  "content_topics": [
    {{
      "topic": "관련 콘텐츠 주제",
      "hook_sentence": "이 콘텐츠에서 전자책을 자연스럽게 연결하는 문장 예시"
    }}
  ],
  "sales_copy": "이 전자책이 왜 광고 없이도 팔릴 수 있는지 한 문단 (논리적, 담백하게, 구조 중심, 감정적 표현 배제. 상세페이지/소개글/SNS 고정글에 그대로 사용 가능하게)",
  "value_summary": {{
    "time_saved": "절약 시간 요약",
    "money_saved": "절약 비용 요약",
    "mistakes_prevented": "방지 실수 요약"
  }}
}}

content_topics는 정확히 5개를 만드세요."""

    return call_gpt_json(client, model, system, user, temperature=0.7, max_tokens=3000)


# ============================================================
# 5단계: 책 표지 이미지 생성 (이미지는 선택사항, 실패해도 계속 진행)
# ============================================================
def step5_generate_cover(client_unused, book_title, subtitle):
    """전자책 표지 생성 - Codex API는 이미지 생성을 직접 지원하지 않을 수 있으므로 실패 허용"""
    try:
        headers = _build_headers(stream=True)
        prompt = f'다음 전자책의 표지에 어울리는 이미지를 생성해주세요. 제목: "{book_title}", 부제목: "{subtitle}". 깔끔하고 현대적인 스타일.'

        body = {
            'model': 'gpt-5-codex',
            'instructions': '이미지 생성 요청을 처리해주세요.',
            'stream': True,
            'store': False,
            'prompt_cache_key': str(uuid.uuid4()),
            'input': [
                {'type': 'message', 'role': 'user', 'content': prompt},
            ],
        }

        resp = http_requests.post(CHATGPT_API_URL, headers=headers, json=body, timeout=120, stream=True)
        if resp.status_code != 200:
            print(f"[Cover] 표지 생성 실패: status={resp.status_code}")
            return None

        # SSE에서 이미지 URL 찾기
        for raw_line in resp.iter_lines():
            if isinstance(raw_line, bytes):
                line = raw_line.decode('utf-8', errors='replace')
            else:
                line = raw_line
            if not line or not line.startswith('data: '):
                continue
            payload = line[6:].strip()
            if payload == '[DONE]':
                break
            try:
                event = json.loads(payload)
                # 이미지 관련 이벤트 탐색
                if event.get('type') == 'response.completed':
                    response_obj = event.get('response', event)
                    for item in response_obj.get('output', []):
                        content = item.get('content', [])
                        if isinstance(content, list):
                            for c in content:
                                if c.get('type') == 'image':
                                    return c.get('url') or c.get('image_url')
                        if item.get('type') == 'image_generation_call':
                            return item.get('result')
            except json.JSONDecodeError:
                continue

        print("[Cover] 표지 이미지를 응답에서 찾을 수 없음")
        return None
    except Exception as e:
        print(f"[Cover] 표지 생성 실패: {e}")
        return None


# ============================================================
# 6단계: 챕터 삽입 이미지 생성
# ============================================================
def generate_chapter_image(client_unused, chapter_title, chapter_purpose):
    """챕터 이미지 생성 - 실패해도 계속 진행"""
    try:
        headers = _build_headers(stream=True)
        prompt = f'다음 챕터 내용에 어울리는 간단한 일러스트를 생성해주세요. 챕터: "{chapter_title}" - {chapter_purpose}. 미니멀 플랫 스타일.'

        body = {
            'model': 'gpt-5-codex',
            'instructions': '이미지 생성 요청을 처리해주세요.',
            'stream': True,
            'store': False,
            'prompt_cache_key': str(uuid.uuid4()),
            'input': [
                {'type': 'message', 'role': 'user', 'content': prompt},
            ],
        }

        resp = http_requests.post(CHATGPT_API_URL, headers=headers, json=body, timeout=120, stream=True)
        if resp.status_code != 200:
            return None

        for raw_line in resp.iter_lines():
            if isinstance(raw_line, bytes):
                line = raw_line.decode('utf-8', errors='replace')
            else:
                line = raw_line
            if not line or not line.startswith('data: '):
                continue
            payload = line[6:].strip()
            if payload == '[DONE]':
                break
            try:
                event = json.loads(payload)
                if event.get('type') == 'response.completed':
                    response_obj = event.get('response', event)
                    for item in response_obj.get('output', []):
                        content = item.get('content', [])
                        if isinstance(content, list):
                            for c in content:
                                if c.get('type') == 'image':
                                    return c.get('url') or c.get('image_url')
                        if item.get('type') == 'image_generation_call':
                            return item.get('result')
            except json.JSONDecodeError:
                continue

        return None
    except Exception as e:
        print(f"[Image] 챕터 이미지 생성 실패: {e}")
        return None


# ============================================================
# 전체 파이프라인 실행
# ============================================================
def generate_ebook(model, topic, include_images=True, progress_callback=None, api_key=None, config=None):
    """
    전자책 생성 전체 파이프라인
    ChatGPT OAuth 토큰 기반 - chatgpt.com/backend-api/codex/responses (SSE stream)
    """
    client = None
    result = {
        'topic': topic,
        'analysis': None,
        'book_info': None,
        'chapters_content': [],
        'marketing': None,
        'cover_url': None,
        'chapter_images': [],
        'error': None,
    }

    total_steps = 6
    current_step = 0

    def progress(msg, data=None):
        nonlocal current_step
        current_step += 1
        if progress_callback:
            progress_callback(current_step, total_steps, msg, data)

    cfg = config or {}

    try:
        # 1단계: 유료 가치 판단
        progress('주제 분석 및 유료 가치 판단 중...')
        result['analysis'] = step1_value_analysis(client, model, topic, cfg)

        # 2단계: 목차 설계
        progress('구매자 관점 목차 설계 중...')
        result['book_info'] = step2_toc_design(client, model, topic, result['analysis'], cfg)

        chapters = result['book_info'].get('chapters', [])
        total_steps = 4 + len(chapters) + (len(chapters) if include_images else 0)

        # 3단계: 챕터별 본문 작성
        for i, chapter in enumerate(chapters):
            progress(f"챕터 {i+1}/{len(chapters)} 집필 중: {chapter.get('title', '')[:30]}...")
            content = step3_write_chapter(client, model, topic, result['book_info'], chapter, i, len(chapters), cfg)
            result['chapters_content'].append({
                'chapter': chapter,
                'content': content,
            })

        # 4단계: 마케팅 분석
        progress('자연 유통 분석 및 판매 소개문 작성 중...')
        result['marketing'] = step4_marketing(client, model, topic, result['analysis'], result['book_info'], cfg)

        # 5단계: 표지 이미지
        if include_images:
            progress('전자책 표지 이미지 생성 중...')
            result['cover_url'] = step5_generate_cover(
                client,
                result['book_info'].get('book_title', topic),
                result['book_info'].get('subtitle', ''),
            )

            # 6단계: 챕터 이미지
            for i, ch_data in enumerate(result['chapters_content']):
                chapter = ch_data['chapter']
                progress(f"챕터 {i+1} 이미지 생성 중...")
                img_url = generate_chapter_image(client, chapter.get('title', ''), chapter.get('purpose', ''))
                result['chapter_images'].append(img_url)
        else:
            progress('이미지 생성 건너뜀')
            result['chapter_images'] = [None] * len(chapters)

    except Exception as e:
        result['error'] = str(e)
        traceback.print_exc()

    return result
