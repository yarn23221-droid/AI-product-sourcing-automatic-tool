"""
유튜브 영상 링크에서 자막(스크립트)과 제목을 가져오는 모듈. API 키 없이 동작한다.

사용법: 이 모듈로 자막을 가져온 뒤, 그 텍스트를 사람(또는 대화 중인 Claude)이 읽고 실제
제품명을 뽑아서 discover_candidates.py의 build_manual_candidate_row()로 후보 목록에 추가한다.
(완전 자동화가 아니라, "자막 가져오기"만 자동화하고 "제품명 판단"은 사람이 확인하는 방식)
"""

import json
import re
import urllib.request

from youtube_transcript_api import YouTubeTranscriptApi


def extract_video_id(url: str) -> str:
    """유튜브 링크(watch?v=, youtu.be/, shorts/ 등 여러 형식) 에서 영상 ID를 뽑는다."""
    match = re.search(r"(?:v=|youtu\.be/|shorts/|embed/)([A-Za-z0-9_-]{11})", url)
    if not match:
        raise ValueError(f"유튜브 링크에서 영상 ID를 찾을 수 없습니다: {url}")
    return match.group(1)


def fetch_transcript(url: str, languages: tuple[str, ...] = ("ko", "en")) -> str:
    """영상 자막을 하나의 텍스트로 이어붙여 반환한다. 자막이 없으면 예외가 발생한다."""
    video_id = extract_video_id(url)
    transcript = YouTubeTranscriptApi().fetch(video_id, languages=list(languages))
    return " ".join(snippet.text for snippet in transcript)


def fetch_video_title(url: str) -> str:
    """영상 제목을 가져온다 (oEmbed, API 키 불필요)."""
    video_id = extract_video_id(url)
    oembed_url = f"https://www.youtube.com/oembed?url=https://www.youtube.com/watch?v={video_id}&format=json"
    with urllib.request.urlopen(oembed_url, timeout=10) as resp:
        data = json.load(resp)
    return data.get("title", "")


if __name__ == "__main__":
    import sys
    url = sys.argv[1] if len(sys.argv) > 1 else input("유튜브 링크: ")
    print("제목:", fetch_video_title(url))
    transcript = fetch_transcript(url)
    print(f"자막 길이: {len(transcript)}자")
    print("미리보기:", transcript[:300])
