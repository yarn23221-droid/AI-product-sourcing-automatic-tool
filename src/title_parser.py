"""
네이버 베스트 상품명 텍스트에서 브랜드/제품명/제품번호를 뽑아내는 파서.

판매글 제목은 "브랜드 제품명 모델번호 [프로모션 문구/스펙/옵션]" 형태가 많다는 점을 이용해
규칙 기반으로 정리한다. 기본 동작은 항상 이 규칙 기반이며, ANTHROPIC_API_KEY 환경변수가
있으면 브랜드 인식 정확도가 훨씬 높은 Claude API로 더 정교하게 파싱해보고 실패 시 규칙
기반으로 되돌아간다 (demand_classifier.py와 동일한 선택적 확장 패턴).

한계: 특히 "브랜드"는 제목에 브랜드명이 안 써 있으면(예: '플레이스테이션프로'만 있고 '소니'는
없음) 규칙 기반으로는 정확히 못 맞힌다. 사람이 최종 확인해야 한다.
"""

import os
import re

from croket_client import is_model_token

# 상품명에서 지워도 되는 프로모션/배송/옵션 관련 문구 (계속 보완 필요)
NOISE_PHRASES = [
    "관부가세포함", "관세포함", "무료배송", "당일발송", "특급배송", "최다판매", "예약판매",
    "국내배송", "미국정품", "미국직구", "중국내수용", "중국내수버전", "일반버전", "단품",
    "출장지원", "슈퍼적립", "브이로그", "콤보", "세트",
]
# 제품명 끝에 붙는 사이즈 표기 (제거 대상 - 옵션은 별도 열에서 '사이즈'로 표시)
_SIZE_SUFFIXES = {"s", "m", "l", "xl", "xxl", "2xl", "3xl", "sm", "med", "lg"}
# 제품명 끝에 붙는 흔한 색상 표기 (제거 대상 - 옵션은 별도 열에서 '컬러'로 표시)
_COLOR_WORDS = {
    "블랙", "화이트", "그레이", "실버", "골드", "로즈골드", "네이비", "베이지", "브라운",
    "레드", "블루", "그린", "옐로우", "핑크", "퍼플", "민트", "카키", "아이보리",
}


def _rule_based_parse(raw_title: str) -> dict:
    text = re.sub(r"\[[^\]]*\]", " ", raw_title)  # '[PS5]' 같은 대괄호 태그 제거
    text = re.sub(r"[-–]\s*", " ", text)  # '관부가세포함 - ' 같은 하이픈 구분자 제거
    words = text.split()
    words = [w for w in words if not any(noise in w for noise in NOISE_PHRASES)]

    option_category = ""
    if words:
        last = re.sub(r"[.,]", "", words[-1])
        if last.lower() in _SIZE_SUFFIXES:
            option_category, words = "사이즈", words[:-1]
        elif last in _COLOR_WORDS:
            option_category, words = "컬러", words[:-1]

    if not words:
        return {"brand": "", "product_name": raw_title.strip(), "model_no": "", "option_category": option_category}

    brand = words[0]
    model_no = next((w for w in words[1:] if is_model_token(w)), "")
    product_words = [w for w in words if w != model_no]  # 제품명은 브랜드를 포함해 완전한 이름으로 유지

    return {
        "brand": brand,
        "product_name": " ".join(product_words),
        "model_no": model_no,
        "option_category": option_category,
    }


def _claude_parse(raw_title: str) -> dict | None:
    """(선택 확장) Claude API로 더 정확하게 브랜드/제품명/제품번호를 뽑아본다.

    ANTHROPIC_API_KEY가 없거나 호출에 실패하면 None을 반환해 규칙 기반으로 대체한다.
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        return None
    try:
        import anthropic
    except ImportError:
        return None

    prompt = (
        "다음은 온라인 쇼핑몰 상품명이다. 여기서 브랜드, 제품명(시리즈, 옵션/색상/사이즈 제외), "
        "제품번호(모델명), 옵션 카테고리(사이즈/컬러/용량 중 이 상품에 선택 옵션이 있어 보이면 "
        "해당 단어, 없으면 빈 값)를 뽑아 'brand|product_name|model_no|option_category' 형식으로만 "
        "답하라. 모르면 빈 값으로.\n"
        f"상품명: {raw_title}"
    )
    try:
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-sonnet-5", max_tokens=100,
            messages=[{"role": "user", "content": prompt}],
        )
        answer = response.content[0].text.strip()
        brand, product_name, model_no, option_category = (answer.split("|") + ["", "", "", ""])[:4]
        return {
            "brand": brand.strip(), "product_name": product_name.strip(),
            "model_no": model_no.strip(), "option_category": option_category.strip(),
        }
    except Exception:
        return None


def parse_title(raw_title: str) -> dict:
    """상품명에서 {"brand", "product_name", "model_no", "option_category"}를 뽑는다."""
    return _claude_parse(raw_title) or _rule_based_parse(raw_title)


if __name__ == "__main__":
    tests = [
        "[PS5]플레이스테이션프로 디스크",
        "DJI 오즈모 포켓 4P Osmo Pocket 4P 브이로그 콤보",
        "오닉스 북스 포크6 이북리더기 전자책 6인치 Carta1300 중국내수용 단품",
        "삼성전자 갤럭시 버즈4 프로 R640 노이즈캔슬링 ANC 블루투스 무선 이어폰",
    ]
    for t in tests:
        print(t, "->", parse_title(t))
