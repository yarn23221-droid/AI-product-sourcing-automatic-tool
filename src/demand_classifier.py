"""
직구수요 존재여부 / 발생사유 분류 모듈.

기본 동작은 키워드 규칙 기반 분류이며, 외부 API 없이 항상 동작한다.
(선택 확장) 환경변수 ANTHROPIC_API_KEY가 설정되어 있으면 Claude API로 더 정교하게 분류를 시도하고,
실패하거나 키가 없으면 자동으로 키워드 규칙으로 되돌아간다(fallback). 이 확장은 필수가 아니다.

절대 하지 않는 것: 텍스트 파일이 없거나 비어 있을 때 그럴듯한 사유를 지어내는 것.
그런 경우는 항상 "미입력"/"파일 없음(확인 필요)"로만 표시한다.
"""

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

# 기획안 5-4에 명시된 직구수요 발생사유 카테고리와, 키워드 규칙에 쓸 키워드 목록.
# 키워드는 소문자로 비교하므로 영문 키워드도 소문자로 적어 둔다.
DEMAND_CATEGORY_KEYWORDS: dict[str, list[str]] = {
    "가격": ["가격", "저렴", "비싸", "가성비", "세일", "할인", "가격경쟁력", "고가", "저가"],
    "국내 미출시/미판매": ["미출시", "미판매", "출시 안", "미발매", "정식 출시", "국내 출시", "역직구", "해외에서만"],
    "성능/기능": ["성능", "스펙", "기능", "속도", "배터리", "화질", "처리속도", "퍼포먼스", "출력", "흡입력"],
    "디자인/휴대성": ["디자인", "휴대성", "무게", "그립감", "크기", "예쁘", "이쁘", "슬림", "휴대용", "컴팩트"],
    "브랜드 신뢰도/입소문": ["입소문", "신뢰", "브랜드", "커뮤니티 추천", "유명", "평판", "화제", "인기"],
    "국내 A/S 불만": ["as", "a/s", "애프터서비스", "수리", "불만", "고객센터", "서비스센터", "as 정책"],
}

DEMAND_TEXT_DIR = Path(__file__).resolve().parent.parent / "data" / "demand_texts"

STATUS_NOT_ENTERED = "미입력"
STATUS_FILE_MISSING = "파일 없음(확인 필요)"
STATUS_EXISTS = "O"
REASON_UNCLASSIFIED = "분류 불가(수동 확인 필요)"


@dataclass
class DemandClassification:
    """직구수요 존재여부 + 발생사유 분류 결과."""
    exists: str   # "O" / "미입력" / "파일 없음(확인 필요)"
    reason: str   # 매칭된 카테고리들(콤마 구분) 또는 exists와 동일한 안내 문구


def _classify_by_keywords(text: str) -> list[str]:
    """텍스트 안에서 카테고리별 키워드를 찾아, 매칭된 카테고리 이름 목록을 반환한다."""
    lowered = text.lower()
    matched = []
    for category, keywords in DEMAND_CATEGORY_KEYWORDS.items():
        if any(keyword in lowered for keyword in keywords):
            matched.append(category)
    return matched


def _classify_with_claude(text: str) -> Optional[list[str]]:
    """(선택 확장) Claude API로 카테고리를 분류해본다.

    ANTHROPIC_API_KEY가 없거나, anthropic 패키지가 설치되어 있지 않거나, 호출에 실패하면
    None을 반환한다 - 이 경우 호출한 쪽에서 키워드 규칙으로 대체(fallback)해야 한다.
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        return None

    try:
        import anthropic  # 선택 의존성: requirements.txt에는 포함하지 않음
    except ImportError:
        return None

    categories = list(DEMAND_CATEGORY_KEYWORDS.keys())
    prompt = (
        "다음 커뮤니티 게시글 텍스트를 읽고, 해외 직구 수요가 발생하는 이유를 아래 카테고리 중에서 "
        "해당하는 것만 골라 콤마로 구분해 답하세요. 카테고리 외의 말은 하지 마세요.\n"
        f"카테고리: {', '.join(categories)}\n\n"
        f"텍스트:\n{text[:4000]}"
    )

    try:
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-sonnet-5",
            max_tokens=200,
            messages=[{"role": "user", "content": prompt}],
        )
        answer = response.content[0].text
    except Exception:
        return None

    matched = [category for category in categories if category in answer]
    return matched or None


def classify_demand_text(filename: Optional[str], demand_text_dir: Path = DEMAND_TEXT_DIR) -> DemandClassification:
    """수요텍스트파일 열의 값을 받아 (직구수요 존재여부, 발생사유)를 판단한다.

    - 파일명이 비어 있으면 "미입력"
    - 파일명은 있는데 실제 파일이 없으면 "파일 없음(확인 필요)"
    - 파일이 있으면 내용을 읽어 카테고리로 분류 ("O" + 매칭된 카테고리들)
    """
    if filename is None or str(filename).strip() == "":
        return DemandClassification(exists=STATUS_NOT_ENTERED, reason=STATUS_NOT_ENTERED)

    file_path = Path(demand_text_dir) / str(filename).strip()
    if not file_path.is_file():
        return DemandClassification(exists=STATUS_FILE_MISSING, reason=STATUS_FILE_MISSING)

    text = file_path.read_text(encoding="utf-8", errors="ignore").strip()
    if not text:
        return DemandClassification(exists=STATUS_EXISTS, reason=REASON_UNCLASSIFIED)

    matched = _classify_with_claude(text)
    if matched is None:
        matched = _classify_by_keywords(text)

    reason = ", ".join(matched) if matched else REASON_UNCLASSIFIED
    return DemandClassification(exists=STATUS_EXISTS, reason=reason)


if __name__ == "__main__":
    # 단독 실행 시 간단한 자체 테스트 (data/demand_texts/ 안의 샘플 파일을 이용)
    print("=== classify_demand_text 테스트 ===")
    print("미입력 케이스 :", classify_demand_text(None))
    print("파일없음 케이스:", classify_demand_text("이런파일없음.txt"))
    for sample_name in ["sample_01.txt", "sample_02.txt"]:
        result = classify_demand_text(sample_name)
        print(f"{sample_name} ->", result)
