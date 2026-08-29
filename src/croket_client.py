"""
크로켓(croket.co.kr) 상품 등록여부/가격/링크 자동 조회 모듈.

크로켓의 검색 API(/api_base/storeItem/search/getStoreItemList)는 브라우저에서 자바스크립트로
발급하는 세션 토큰이 있어야만 응답하기 때문에(확인 완료), requests만으로는 호출할 수 없다.
그래서 Playwright로 실제 검색 화면을 그대로 흉내낸다 - 검색창을 열고 입력해서 화면에 뜨는 API
응답을 그대로 가로채는 방식.

여러 상품을 조회할 때는 브라우저를 한 번만 띄우고 CroketClient를 재사용하는 것이 훨씬 빠르다.
"""

import re
from playwright.sync_api import sync_playwright

SEARCH_ENTRY_SELECTOR = "a.CroketBuyerTopBar_search__wrapper__TaALQ"
# 크롬이 페이지에 자동으로 끼워 넣는 구글 번역 위젯의 숨은 input과 헷갈리지 않도록,
# placeholder 텍스트로 실제 검색창만 정확히 찾는다 (일반 "input" 셀렉터는 그 위젯을 집을 수 있음)
SEARCH_INPUT_SELECTOR = "input[placeholder*='찾고']"


def _normalize_token(token: str) -> str:
    """토큰 하나의 한/영 표기 차이를 흡수한다 (예: '프로' == 'pro')."""
    token = token.lower()
    token = token.replace("프로", "pro").replace("맥스", "max").replace("미니", "mini").replace("플러스", "plus")
    return re.sub(r"[^a-z0-9가-힣]", "", token)


def _tokenize(text: str) -> set[str]:
    """상품명을 의미있는 토큰 집합으로 쪼갠다. 실제 판매글 제목은 단어 순서가 제각각이라
    (예: '오닉스 북스 포크6 이북리더기' vs '오닉스 BOOK 포크6 poke6 이북리더기'),
    전체 문자열 포함 여부가 아니라 토큰 겹침 비율로 같은 상품인지 판단한다."""
    raw_tokens = re.split(r"[\s\[\]()/,\-_]+", text)
    tokens = {_normalize_token(t) for t in raw_tokens if len(t) >= 2}
    tokens.discard("")
    return tokens


# 숫자가 붙어도 "모델 식별자"가 아니라 스펙/단위인 흔한 접미사들 (예: '6인치', '1500mAh')
_SPEC_UNIT_SUFFIXES = ("인치", "mah", "kg", "cm", "mm", "gb", "tb", "hz", "kw", "wh", "kwh", "w", "ml", "l", "ppi")


def is_model_token(token: str) -> bool:
    """'포크6', 'r640', 'carta1300'처럼 글자+숫자가 결합된, 모델을 특정하는 토큰인지 판단한다.

    '6인치', '1500mah'처럼 숫자 뒤에 흔한 단위가 붙은 토큰은 모델 식별자가 아니라 스펙이므로
    제외한다 - 그래야 화면 크기 같은 공통 스펙만 같고 모델번호가 다른 상품을 같은 걸로
    착각하지 않는다.
    """
    has_letter = any(ch.isalpha() for ch in token)
    has_digit = any(ch.isdigit() for ch in token)
    if not (has_letter and has_digit):
        return False
    return not token.endswith(_SPEC_UNIT_SUFFIXES)


def match_ratio(query: str, candidate_title: str) -> float:
    """검색어 토큰이 후보 제목에 얼마나 겹치는지 비율(0~1)을 계산한다.

    검색어에 모델 식별 토큰(예: '포크6', 'carta1300')이 있는데 후보 제목에 그 토큰이
    하나도 없으면('포크7'만 있음 등), 브랜드/스펙만 같은 다른 모델일 가능성이 커서 0을
    반환한다 - 브랜드 단어나 공통 스펙만 겹쳐서 다른 모델을 같은 상품으로 착각하지 않기 위함.
    """
    query_tokens = _tokenize(query)
    candidate_tokens = _tokenize(candidate_title)
    if not query_tokens:
        return 0.0

    model_query_tokens = {t for t in query_tokens if is_model_token(t)}
    if model_query_tokens and not (model_query_tokens & candidate_tokens):
        return 0.0

    overlap = query_tokens & candidate_tokens
    return len(overlap) / len(query_tokens)


def titles_match(query: str, candidate_title: str, min_overlap_ratio: float = 0.45, min_overlap_count: int = 2) -> bool:
    """검색어와 크로켓 검색결과 상품명이 같은 제품으로 보이는지 토큰 겹침으로 비교한다.

    완전 일치나 전체 문자열 포함을 요구하지 않고, 검색어 토큰의 상당 비율이 후보 제목에도
    등장하는지를 본다. 휴리스틱이라 오탐/누락이 있을 수 있어 사람이 최종 확인해야 한다.
    """
    overlap = _tokenize(query) & _tokenize(candidate_title)
    ratio = match_ratio(query, candidate_title)
    return len(overlap) >= min_overlap_count and ratio >= min_overlap_ratio


class CroketClient:
    """브라우저를 한 번만 띄워두고 여러 번 검색을 재사용하기 위한 클래스."""

    def __init__(self, headless: bool = True):
        self._playwright = sync_playwright().start()
        self._browser = self._playwright.chromium.launch(headless=headless)
        self._page = self._browser.new_page()
        self._page.goto("https://www.croket.co.kr/", timeout=30000)
        self._page.wait_for_timeout(1000)
        self._page.click(SEARCH_ENTRY_SELECTOR, timeout=10000)
        self._page.wait_for_timeout(1500)

    def search(self, query: str, items_limit: int = 10) -> list[dict]:
        """검색어로 크로켓 상품을 검색해 원본 결과 목록(가격/링크/배송비 포함가 등)을 반환한다."""
        captured = []

        def on_response(res):
            if "getStoreItemList" in res.url and res.request.method == "POST":
                try:
                    captured.append(res.json())
                except Exception:
                    pass

        self._page.on("response", on_response)
        try:
            box = self._page.locator(SEARCH_INPUT_SELECTOR).first
            box.click(timeout=5000)
            box.fill("")
            self._page.wait_for_timeout(200)
            box.type(query, delay=80)
            box.press("Enter")
            self._page.wait_for_timeout(4000)
        finally:
            self._page.remove_listener("response", on_response)

        # 검색 한 번에 여러 getStoreItemList 응답이 올 수 있어(신상품/중고 등), 가장 결과가 많은 걸 쓴다
        best = max(captured, key=lambda d: len(d.get("storeItemList", [])), default={})
        items = best.get("storeItemList", [])[:items_limit]
        results = []
        for item in items:
            price_with_shipping = item.get("itemPriceForBuyer", item.get("itemPrice"))
            results.append({
                "itemId": item.get("itemId"),
                "itemName": item.get("itemName"),
                "priceInclShipping": price_with_shipping,
                "soldNums": item.get("soldNums", 0),
                "linkUrl": f"https://www.croket.co.kr/item/{item.get('itemId')}",
            })
        return results

    def find_registered_product(self, product_title: str) -> dict | None:
        """상품명으로 검색해서, 같은 상품으로 판단되는 결과 중 겹침 비율이 가장 높은 것을 반환한다.

        크로켓 자체 검색엔진이 검색어가 길거나 특정 표기('포켓 4P' 등)면 결과를 0건 주는 경우가
        있어서(확인함), 원본 제목 그대로 한 번만 검색하지 않고 앞부분 단어 수를 줄여가며
        결과가 나올 때까지 재시도한다. 최종 일치 판단은 항상 원본 전체 제목 기준으로 한다.
        """
        words = product_title.split()
        query_variants = []
        for word_count in (6, 4, 3, 2):
            if len(words) >= word_count:
                variant = " ".join(words[:word_count])
                if variant not in query_variants:
                    query_variants.append(variant)
        if not query_variants:
            query_variants = [product_title]

        for query in query_variants:
            results = self.search(query)
            best, best_ratio = None, 0.0
            for item in results:
                if not titles_match(product_title, item["itemName"]):
                    continue
                ratio = match_ratio(product_title, item["itemName"])
                if ratio > best_ratio:
                    best, best_ratio = item, ratio
            if best:
                return best
        return None

    def close(self):
        self._browser.close()
        self._playwright.stop()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()


if __name__ == "__main__":
    with CroketClient() as client:
        for query in ["이어폰", "노트북"]:
            print(f"\n=== '{query}' 검색 ===")
            for item in client.search(query, items_limit=3):
                print(" -", item["itemName"], "/", item["priceInclShipping"], "원(배송비포함) /", item["linkUrl"])
