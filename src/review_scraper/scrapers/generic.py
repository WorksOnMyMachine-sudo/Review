from __future__ import annotations

from review_scraper.models import ReviewRecord
from review_scraper.scrapers.base import BaseScraper


class GenericScraper(BaseScraper):
    """通用评论页解析（amazon / walmart / target 等占位，需按站点改选择器）。"""

    site_name = "generic"

    SITE_SELECTORS: dict[str, dict[str, list[str]]] = {
        "amazon": {
            "block": ["div[data-hook='review']", "#cm_cr-review_list .review"],
            "rating": ["i[data-hook='review-star-rating'] span", "span.a-icon-alt"],
            "title": ["a[data-hook='review-title']", "span.review-title"],
            "body": ["span[data-hook='review-body'] span", "div.review-text-content span"],
            "date": ["span[data-hook='review-date']"],
            "author": ["span.a-profile-name"],
        },
        "walmart": {
            "block": ["[itemprop='review']", ".review"],
            "rating": ["[itemprop='ratingValue']", ".stars"],
            "title": [".review-title"],
            "body": [".review-text", "[itemprop='reviewBody']"],
            "date": ["[itemprop='datePublished']", ".review-date"],
            "author": [".reviewer"],
        },
        "target": {
            "block": ["[data-test='review']", ".h-padding-h-default"],
            "rating": ["[data-test='rating']"],
            "title": ["[data-test='review-title']"],
            "body": ["[data-test='review-body']"],
            "date": ["[data-test='review-date']"],
            "author": ["[data-test='review-author']"],
        },
    }

    def __init__(self, defaults, site_key: str) -> None:
        super().__init__(defaults)
        self.site_key = site_key
        self.selectors = self.SITE_SELECTORS.get(site_key, self.SITE_SELECTORS["amazon"])

    def scrape(
        self,
        *,
        url: str,
        model_id: str,
        model_name: str,
        max_pages: int = 5,
    ) -> list[ReviewRecord]:
        html = self.fetch_html(url)
        soup = self.parse_soup(html)
        blocks = []
        for selector in self.selectors["block"]:
            blocks = soup.select(selector)
            if blocks:
                break

        records: list[ReviewRecord] = []
        for block in blocks:
            records.append(
                ReviewRecord(
                    model_id=model_id,
                    model_name=model_name,
                    site=self.site_key,
                    source_url=url,
                    rating=self._pick(block, "rating"),
                    title=self._pick(block, "title"),
                    review_text=self._pick(block, "body"),
                    review_date=self._pick(block, "date"),
                    author=self._pick(block, "author"),
                )
            )
        return [r for r in records if r.review_text or r.title]

    def _pick(self, block, key: str) -> str | None:
        for selector in self.selectors.get(key, []):
            node = block.select_one(selector)
            if node:
                text = node.get_text(strip=True) or node.get("content")
                if text:
                    return str(text).strip()
        return None
