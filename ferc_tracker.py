import json
import re
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin
from zoneinfo import ZoneInfo

import requests
from bs4 import BeautifulSoup

NEWS_URL = "https://www.ferc.gov/news-events/news/news-releases-headlines"
RELIABILITY_URL = "https://www.ferc.gov/electric-reliability"
DATA_PATH = Path("docs/data/ferc.json")
HISTORY_PATH = Path("docs/data/ferc-history.json")
MAX_ITEMS = 60

KEYWORDS = {
    "NERC / Reliability Standards": ["nerc", "reliability standard", "bulk power", "bulk-power", "electric reliability"],
    "CIP / Cybersecurity": ["cip-", "cyber", "security", "supply chain", "critical infrastructure"],
    "Extreme Weather": ["cold weather", "extreme cold", "winter", "heat", "weather", "storm"],
    "Generation / Transmission": ["generation", "generator", "transmission", "interconnection", "large load", "grid"],
    "SPP / Regional Markets": ["spp", "southwest power pool", "rto", "iso", "market"],
    "Kansas / Public Power": ["kansas", "municipal", "public power"],
}


def fetch_html(url):
    response = requests.get(url, timeout=60, headers={"User-Agent": "BPU-Regulatory-Tracker/2.0"})
    response.raise_for_status()
    return response.text


def clean(text):
    return re.sub(r"\s+", " ", text or "").strip()


def categorize(title):
    lower = title.lower()
    categories = [name for name, words in KEYWORDS.items() if any(word in lower for word in words)]
    return categories or ["Other FERC Activity"]


def relevance_score(title, categories):
    score = 0
    lower = title.lower()
    score += 4 if "nerc" in lower or "reliability standard" in lower else 0
    score += 3 if any(x in lower for x in ["cip-", "cyber", "bulk-power", "electric reliability"]) else 0
    score += 2 if any(x in lower for x in ["transmission", "generation", "interconnection", "large load", "spp"]) else 0
    score += min(len(categories), 3)
    return score


def parse_news():
    soup = BeautifulSoup(fetch_html(NEWS_URL), "html.parser")
    items = []
    seen = set()
    for link in soup.select('a[href*="/news-events/news/"]'):
        title = clean(link.get_text(" ", strip=True))
        href = urljoin(NEWS_URL, link.get("href", ""))
        if len(title) < 12 or href in seen or href.rstrip("/") == NEWS_URL.rstrip("/"):
            continue
        container = link.find_parent(["article", "div", "li"]) or link.parent
        context = clean(container.get_text(" ", strip=True)) if container else title
        date_match = re.search(r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}", context)
        type_match = re.search(r"\b(News Releases?|Headlines?|Fact Sheets?|Presentations?|Summaries|Notices?)\b", context, re.I)
        categories = categorize(title)
        items.append({
            "id": href,
            "title": title,
            "url": href,
            "date": date_match.group(0) if date_match else "Date not listed",
            "type": type_match.group(0).title() if type_match else "FERC Update",
            "categories": categories,
            "score": relevance_score(title, categories),
            "source": "FERC News Releases & Headlines",
        })
        seen.add(href)
        if len(items) >= MAX_ITEMS:
            break
    return items


def load_json(path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return default


def main():
    now_utc = datetime.now(timezone.utc)
    now_ct = now_utc.astimezone(ZoneInfo("America/Chicago"))
    previous = load_json(DATA_PATH, {})
    previous_ids = {item.get("id") for item in previous.get("items", [])}
    items = parse_news()
    new_items = [item for item in items if item["id"] not in previous_ids] if previous_ids else []
    prioritized = sorted(items, key=lambda x: (-x["score"], x["date"]))
    history = load_json(HISTORY_PATH, [])
    history_entry = {
        "date": now_ct.strftime("%b %d, %Y %I:%M %p CT"),
        "new_items": len(new_items),
        "items_checked": len(items),
        "status_text": f"{len(new_items)} new relevant item{'s' if len(new_items) != 1 else ''}",
    }
    history = [history_entry] + history
    payload = {
        "status": "ok",
        "checked_at": now_utc.isoformat(),
        "checked_at_display": now_ct.strftime("%b %d, %Y %I:%M %p"),
        "message": "Official FERC sources checked successfully.",
        "summary": {
            "items_checked": len(items),
            "new_items": len(new_items),
            "priority_items": sum(1 for item in items if item["score"] >= 4),
            "sources_checked": 2,
        },
        "source_urls": [NEWS_URL, RELIABILITY_URL],
        "new_items": sorted(new_items, key=lambda x: -x["score"]),
        "items": prioritized,
        "history": history[:30],
        "topics": list(KEYWORDS.keys()),
    }
    DATA_PATH.parent.mkdir(parents=True, exist_ok=True)
    DATA_PATH.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    HISTORY_PATH.write_text(json.dumps(history[:90], indent=2, ensure_ascii=False), encoding="utf-8")


if __name__ == "__main__":
    main()
