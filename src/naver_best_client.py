"""
네이버 베스트(snxbest.naver.com) 상품 후보 자동 수집 모듈.

네이버가 공개적으로 막지 않은 내부 JSON API 두 개를 그대로 호출한다 (봇 차단 없음, 확인 완료).
- 하위 카테고리 목록 API: 대분류 카테고리ID로 하위 카테고리(게임기/타이틀, PC 등) 이름+ID를 받는다.
- 순위 API: 카테고리ID + 기간(일간/주간)으로 랭킹 상품 목록을 받는다.

주의: 네이버 쇼핑 '가격비교/일반검색'(search.shopping.naver.com)은 봇 접근을 명시적으로 차단하고
있어서 이 모듈에서 다루지 않는다. 그 부분(구매/찜/리뷰수 기반 직구수요 등급, 채널별 가격비교)은
사람이 직접 조사해서 입력해야 한다.
"""

import requests

BASE_URL = "https://snxbest.naver.com/api/v1/snxbest/product"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

# 디지털/가전 대분류 카테고리ID (네이버 베스트 기준, 확인 완료)
DIGITAL_APPLIANCE_CATEGORY_ID = "50000003"


def fetch_subcategories(top_category_id: str = DIGITAL_APPLIANCE_CATEGORY_ID, period_type: str = "DAILY") -> list[dict]:
    """대분류 카테고리 밑에 있는 하위 카테고리(게임기/타이틀, PC 등) 목록을 가져온다.

    반환값: [{"categoryId": "50000088", "categoryName": "게임기/타이틀"}, ...] ("전체"는 제외)
    """
    resp = requests.get(
        f"{BASE_URL}/categories",
        params={"categoryId": top_category_id, "ageType": "ALL", "type": "DIV2",
                "sortType": "PRODUCT_CLICK", "periodType": period_type},
        headers=HEADERS, timeout=15,
    )
    resp.raise_for_status()
    data = resp.json()
    return [
        {"categoryId": item["categoryId"], "categoryName": item["categoryName"]}
        for item in data
        if item.get("categoryName") != "전체"
    ]


def fetch_rank(category_id: str, period_type: str = "DAILY", top_n: int = 20) -> list[dict]:
    """특정 (하위)카테고리의 랭킹 상품 목록을 가져온다. 기본으로 상위 top_n개만 반환한다.

    period_type: "DAILY"(일간) 또는 "WEEKLY"(주간)
    """
    resp = requests.get(
        f"{BASE_URL}/rank",
        params={"ageType": "ALL", "categoryId": category_id,
                "sortType": "PRODUCT_CLICK", "periodType": period_type},
        headers=HEADERS, timeout=15,
    )
    resp.raise_for_status()
    data = resp.json()
    products = data.get("products", [])[:top_n]
    return [
        {
            "productId": p["productId"],
            "rank": p["rank"],
            "title": p["title"],
            "linkUrl": p["linkUrl"],
            "priceValue": p["priceValue"],
            "discountPriceValue": p.get("discountPriceValue", p["priceValue"]),
            "deliveryFee": p.get("deliveryFee", "0"),
            "mallNm": p.get("mallNm", ""),
            "mallLinkUrl": p.get("mallLinkUrl", ""),
            "reviewCount": p.get("reviewCount", "0"),
        }
        for p in products
    ]


if __name__ == "__main__":
    subs = fetch_subcategories()
    print(f"디지털/가전 하위 카테고리 {len(subs)}개:")
    for s in subs[:5]:
        print(" -", s)

    rank = fetch_rank(subs[0]["categoryId"], top_n=3)
    print(f"\n'{subs[0]['categoryName']}' 일간 베스트 상위 3개:")
    for r in rank:
        print(" -", r["rank"], r["title"], r["discountPriceValue"], "원 /", r["mallNm"])
