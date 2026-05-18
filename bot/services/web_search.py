import logging
from urllib.parse import quote_plus

import aiohttp
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

SOURCES = [
    {
        "name": "Честный ЗНАК (сообщество)",
        "base_url": "https://markirovka.ru/community/?search=",
        "domain": "markirovka.ru",
    },
    {
        "name": "Честный ЗНАК (официальный)",
        "base_url": "https://честныйзнак.рф/search/?q=",
        "domain": "честныйзнак.рф",
    },
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9",
}


async def _fetch_page(session: aiohttp.ClientSession, url: str) -> str:
    try:
        async with session.get(url, headers=HEADERS, timeout=aiohttp.ClientTimeout(total=10)) as resp:
            if resp.status == 200:
                return await resp.text()
    except Exception as e:
        logger.warning(f"Failed to fetch {url}: {e}")
    return ""


def _extract_text(html: str, max_chars: int = 3000) -> str:
    if not html:
        return ""
    soup = BeautifulSoup(html, "html.parser")

    for tag in soup(["script", "style", "nav", "footer", "header", "aside"]):
        tag.decompose()

    text = soup.get_text(separator="\n", strip=True)
    lines = [line.strip() for line in text.splitlines() if len(line.strip()) > 20]
    result = "\n".join(lines)
    return result[:max_chars]


async def _google_search(session: aiohttp.ClientSession, query: str) -> list[dict]:
    """Search using Google and extract results from relevant domains."""
    domains = "site:markirovka.ru OR site:честныйзнак.рф OR site:mdlp.crpt.ru"
    search_url = f"https://www.google.com/search?q={quote_plus(query + ' ' + domains)}&hl=ru&num=5"

    html = await _fetch_page(session, search_url)
    if not html:
        return []

    soup = BeautifulSoup(html, "html.parser")
    results = []

    for g in soup.select("div.g, div[data-hveid]"):
        link_tag = g.find("a", href=True)
        if not link_tag:
            continue
        href = link_tag["href"]
        if not href.startswith("http"):
            continue

        snippet_el = g.find("span", class_=lambda c: c and "st" not in str(c))
        snippet = g.get_text(strip=True)[:300] if not snippet_el else snippet_el.get_text(strip=True)[:300]

        results.append({"url": href, "snippet": snippet})

    return results[:5]


async def search_marking_sites(query: str) -> str:
    """Search marking-related websites and return combined context."""
    collected_texts = []

    async with aiohttp.ClientSession() as session:
        google_results = await _google_search(session, query)

        if google_results:
            for result in google_results[:3]:
                page_html = await _fetch_page(session, result["url"])
                page_text = _extract_text(page_html, max_chars=2000)
                if page_text:
                    collected_texts.append(
                        f"[Источник: {result['url']}]\n{page_text}"
                    )

        if not collected_texts:
            for source in SOURCES:
                url = source["base_url"] + quote_plus(query)
                html = await _fetch_page(session, url)
                text = _extract_text(html, max_chars=2000)
                if text:
                    collected_texts.append(
                        f"[Источник: {source['name']}]\n{text}"
                    )

    if not collected_texts:
        return ""

    return "\n\n---\n\n".join(collected_texts[:3])
