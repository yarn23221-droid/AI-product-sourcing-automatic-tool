"""
가격 분석 로직 모듈.

외부 API를 전혀 호출하지 않는다. 입력 시트에 사람이 적어 넣은 채널별 가격/크로켓 가격만 가지고
순수 계산(최저가 찾기, 타겟가 제안)을 수행한다.

기획안 5-3의 "타겟가 제안"은 실제 원가 데이터가 없는 상태에서 만든 휴리스틱(경험적 규칙)이다.
정확한 원가 기반 가격 책정이 아니라, "시장 최저가 대비 몇 % 수준으로 갈지"에 대한 초안 제안일 뿐이며,
최종 판단은 사람이 해야 한다.
"""

from dataclasses import dataclass
from typing import Optional


# 크로켓 가격이 채널 최저가보다 이 비율(5%)을 넘게 비싸면 "가격 조정 필요"로 판단한다.
EXPENSIVE_THRESHOLD = 0.05
# 크로켓 가격이 채널 최저가보다 이 비율(5%)을 넘게 저렴하면 "이미 충분히 저렴함"으로 판단한다.
CHEAP_THRESHOLD = -0.05
# 가격 조정이 필요할 때, 최저가 대비 몇 % 위로 타겟가를 제안할지
ADJUST_MARKUP = 0.02
# 크로켓 가격 자체가 없을 때(미등록 등), 최저가 대비 몇 % 위로 초기 타겟가를 제안할지
INITIAL_MARKUP = 0.03


@dataclass
class ChannelPrice:
    """채널명 + 채널 가격 한 쌍."""
    name: Optional[str]
    price: Optional[float]


@dataclass
class TargetPriceResult:
    """타겟가 제안 결과 (제안 가격 + 사람이 읽을 코멘트)."""
    target_price: Optional[float]
    comment: str


def to_number(value) -> Optional[float]:
    """엑셀 셀에서 읽은 값을 숫자로 변환한다. 빈칸/텍스트는 None으로 처리한다(값을 지어내지 않음)."""
    if value is None:
        return None
    if isinstance(value, str):
        value = value.strip()
        if value == "":
            return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if number <= 0:
        return None
    return number


def find_lowest_channel(channels: list[ChannelPrice]) -> tuple[Optional[str], Optional[float]]:
    """채널1~3 중 이름과 가격이 모두 유효한 것들만 모아 최저가 채널명/가격을 반환한다.

    유효한 채널이 하나도 없으면 (None, None)을 반환한다 (값을 임의로 지어내지 않음).
    """
    valid_channels = []
    for channel in channels:
        name = (channel.name or "").strip() if isinstance(channel.name, str) else channel.name
        price = to_number(channel.price)
        if name and price is not None:
            valid_channels.append((name, price))

    if not valid_channels:
        return None, None

    cheapest_name, cheapest_price = min(valid_channels, key=lambda pair: pair[1])
    return cheapest_name, cheapest_price


def suggest_target_price(crocket_price_raw, lowest_price: Optional[float]) -> TargetPriceResult:
    """기획안 5-3 로직: 크로켓 가격과 채널 최저가를 비교해 타겟가와 코멘트를 만든다.

    - 채널 최저가 자체가 없으면 비교 불가 → 타겟가 없음, "확인 필요" 코멘트
    - 크로켓 가격이 없으면 → 최저가의 (1 + INITIAL_MARKUP) 을 초기 타겟가로 제안
    - 크로켓 가격이 최저가보다 5% 넘게 비싸면 → 최저가의 (1 + ADJUST_MARKUP) 을 타겟가로 제안
    - 그 외(±5% 이내 또는 이미 더 저렴)에는 → 현재가(크로켓 가격) 유지를 제안
    """
    crocket_price = to_number(crocket_price_raw)

    if lowest_price is None:
        return TargetPriceResult(
            target_price=None,
            comment="채널 가격이 입력되지 않아 비교할 수 없음 - 확인 필요",
        )

    if crocket_price is None:
        target = round(lowest_price * (1 + INITIAL_MARKUP) / 10) * 10
        return TargetPriceResult(
            target_price=target,
            comment=f"크로켓 가격 미확인 - 채널 최저가 대비 +{INITIAL_MARKUP * 100:.0f}% 수준을 초기 타겟가로 제안",
        )

    diff_ratio = (crocket_price - lowest_price) / lowest_price

    if diff_ratio > EXPENSIVE_THRESHOLD:
        target = round(lowest_price * (1 + ADJUST_MARKUP) / 10) * 10
        comment = f"시장가 대비 {diff_ratio * 100:.1f}% 비쌈 - 가격 조정 필요 (최저가 대비 +{ADJUST_MARKUP * 100:.0f}% 제안)"
        return TargetPriceResult(target_price=target, comment=comment)

    if diff_ratio < CHEAP_THRESHOLD:
        comment = f"시장가 대비 {abs(diff_ratio) * 100:.1f}% 저렴 - 현재가 유지 권장"
    else:
        comment = f"시장가와 유사한 수준(차이 {diff_ratio * 100:.1f}%) - 현재가 유지 권장"

    return TargetPriceResult(target_price=crocket_price, comment=comment)


if __name__ == "__main__":
    # 단독 실행 시 간단한 자체 테스트를 돈다 (pytest 없이도 눈으로 확인할 수 있도록).
    print("=== find_lowest_channel 테스트 ===")
    channels = [ChannelPrice("쿠팡", 259000), ChannelPrice("11번가", 265000), ChannelPrice(None, None)]
    print(find_lowest_channel(channels))  # 예상: ('쿠팡', 259000)
    print(find_lowest_channel([ChannelPrice("", ""), ChannelPrice(None, None)]))  # 예상: (None, None)

    print("\n=== suggest_target_price 테스트 ===")
    print(suggest_target_price(289000, 259000))  # 5%(12950) 넘게 비쌈 -> 조정 제안
    print(suggest_target_price(265000, 259000))  # 2.3% 비쌈 -> ±5% 이내라 유지
    print(suggest_target_price(None, 320000))    # 크로켓가 없음 -> 초기 타겟가 제안
    print(suggest_target_price(300000, None))    # 채널가 없음 -> 확인 필요
