from __future__ import annotations

import json
import re
from html import unescape
from urllib.parse import parse_qs, urlencode, urlparse, urlunparse

from review_scraper.models import ReviewRecord
from review_scraper.scrapers.base import BaseScraper


class BestBuyScraper(BaseScraper):
    site_name = "bestbuy"

    REVIEW_BLOCK_SELECTORS = [
        "[data-testid='customer-review']",
        ".c-reviews-list .review-item",
        ".review-item",
        "div[itemprop='review']",
        ".BVRRDisplayContentReview",
        ".BVRRContentReview",
        ".BVRRReview",
        ".bv-content-item",
    ]
    RATING_SELECTORS = [
        "[itemprop='ratingValue']",
        ".c-ratings-reviews .c-ratings",
        ".rating",
        ".BVRRRatingNumber",
        ".bv-rating-stars-container",
        ".bv-rating",
    ]
    TITLE_SELECTORS = [
        ".review-title",
        "h3.c-review-title",
        "[itemprop='name']",
        ".BVRRReviewTitle",
        ".bv-content-title",
    ]
    BODY_SELECTORS = [
        ".review-body",
        ".c-review-body",
        "[itemprop='reviewBody']",
        ".BVRRReviewTextContainer",
        ".BVRRReviewText",
        ".bv-content-summary-body-text",
    ]
    DATE_SELECTORS = [
        "time",
        ".c-review-date",
        "[itemprop='datePublished']",
        ".BVRRReviewDate",
        ".bv-content-datetime-stamp",
    ]
    AUTHOR_SELECTORS = [
        ".c-review-author",
        ".author",
        "[itemprop='author']",
        ".BVRRNickname",
        ".bv-author",
    ]

    def __init__(self, defaults, site_name: str = "bestbuy") -> None:
        super().__init__(defaults)
        self.site_name = site_name

    def scrape(
        self,
        *,
        url: str,
        model_id: str,
        model_name: str,
        max_pages: int = 5,
    ) -> list[ReviewRecord]:
        records, endpoint_reached = self._scrape_bazaarvoice(
            url=url,
            model_id=model_id,
            model_name=model_name,
            max_pages=max_pages,
        )
        if records or endpoint_reached:
            return records

        return self._scrape_html(
            url=url,
            model_id=model_id,
            model_name=model_name,
            max_pages=max_pages,
        )

    def _scrape_bazaarvoice(
        self,
        *,
        url: str,
        model_id: str,
        model_name: str,
        max_pages: int,
    ) -> tuple[list[ReviewRecord], bool]:
        sku = self._extract_sku(url)
        if not sku:
            return [], False

        rating_filter = self._rating_filter(url)
        endpoint_reached = False
        records: list[ReviewRecord] = []
        for page in range(1, max_pages + 1):
            page_records, page_reached = self._scrape_bazaarvoice_page(
                sku=sku,
                page=page,
                rating_filter=rating_filter,
                model_id=model_id,
                model_name=model_name,
            )
            endpoint_reached = endpoint_reached or page_reached
            if not page_records:
                break
            records.extend(page_records)
            if page < max_pages:
                self.delay()
        return records, endpoint_reached

    def _scrape_bazaarvoice_page(
        self,
        *,
        sku: str,
        page: int,
        rating_filter: str | None,
        model_id: str,
        model_name: str,
    ) -> tuple[list[ReviewRecord], bool]:
        endpoint_reached = False
        for endpoint in self._bazaarvoice_urls(sku, page):
            try:
                js = self._fetch_html_with_retries(endpoint, attempts=1)
            except Exception:  # noqa: BLE001
                continue
            endpoint_reached = True

            html = self._html_from_bazaarvoice_js(js)
            if not html:
                continue

            soup = self.parse_soup(html)
            blocks = self._find_review_blocks(soup)
            if not blocks:
                continue

            records: list[ReviewRecord] = []
            for block in blocks:
                record = self._parse_block(
                    block,
                    model_id=model_id,
                    model_name=model_name,
                    source_url=endpoint,
                )
                if not record or not (record.review_text or record.title):
                    continue
                if rating_filter and record.rating and record.rating != rating_filter:
                    continue
                records.append(record)
            return records, endpoint_reached
        return [], endpoint_reached

    def _scrape_html(
        self,
        *,
        url: str,
        model_id: str,
        model_name: str,
        max_pages: int,
    ) -> list[ReviewRecord]:
        records: list[ReviewRecord] = []
        for page in range(1, max_pages + 1):
            page_url, html = self._fetch_first_html(url, page)
            soup = self.parse_soup(html)
            blocks = self._find_review_blocks(soup)
            if not blocks:
                break

            page_count = 0
            for block in blocks:
                record = self._parse_block(
                    block,
                    model_id=model_id,
                    model_name=model_name,
                    source_url=page_url,
                )
                if record and (record.review_text or record.title):
                    records.append(record)
                    page_count += 1

            if page_count == 0:
                break
            if page < max_pages:
                self.delay()
        return records

    def _bazaarvoice_urls(self, sku: str, page: int) -> list[str]:
        query = urlencode(
            {
                "format": "embeddedhtml",
                "sort": "submissionTime",
                "dir": "desc",
                "page": page,
            }
        )
        return [
            f"https://bestbuy.ugc.bazaarvoice.com/3545w/{sku}/reviews.djs?{query}",
            f"https://bestbuy.ugc.bazaarvoice.com/3545/{sku}/reviews.djs?{query}",
            f"https://bestbuy.ugc.bazaarvoice.com/3545-en_us/{sku}/reviews.djs?{query}",
        ]

    @staticmethod
    def _html_from_bazaarvoice_js(js: str) -> str:
        match = re.search(r"var\s+materials\s*=\s*(\{.*?\});", js, re.S)
        if match:
            try:
                data = json.loads(match.group(1))
                return "\n".join(BestBuyScraper._decode_escaped_html(str(value)) for value in data.values())
            except json.JSONDecodeError:
                pass

        parts = re.findall(r'"((?:\\.|[^"\\])*(?:BVRRReview|bv-content-item).*?)"', js, re.S)
        html_parts = []
        for part in parts:
            try:
                html_parts.append(json.loads(f'"{part}"'))
            except json.JSONDecodeError:
                html_parts.append(part.encode("utf-8").decode("unicode_escape", errors="ignore"))
        if html_parts:
            return "\n".join(BestBuyScraper._decode_escaped_html(part) for part in html_parts)

        return BestBuyScraper._decode_escaped_html(js)

    @staticmethod
    def _decode_escaped_html(value: str) -> str:
        return unescape(
            value.replace(r"\/", "/")
            .replace(r"\"", '"')
            .replace(r"\n", "\n")
            .replace(r"\t", "\t")
        )

    def _fetch_first_html(self, base_url: str, page: int) -> tuple[str, str]:
        last_exc: Exception | None = None
        for candidate in self._review_url_candidates(base_url):
            page_url = self._page_url(candidate, page)
            try:
                return page_url, self._fetch_html_with_retries(page_url)
            except Exception as exc:  # noqa: BLE001
                last_exc = exc
        raise last_exc or RuntimeError("Best Buy request failed")

    def _fetch_html_with_retries(self, url: str, attempts: int = 2) -> str:
        last_exc: Exception | None = None
        for attempt in range(1, attempts + 1):
            try:
                return self.fetch_html(url)
            except Exception as exc:  # noqa: BLE001
                last_exc = exc
                if attempt < attempts:
                    self.delay()
        raise last_exc or RuntimeError("Best Buy request failed")

    def _review_url_candidates(self, url: str) -> list[str]:
        candidates = []
        normalized = self._legacy_reviews_url(url)
        if normalized:
            candidates.append(normalized)
        candidates.append(url)

        out = []
        seen = set()
        for candidate in candidates:
            if candidate in seen:
                continue
            seen.add(candidate)
            out.append(candidate)
        return out

    @staticmethod
    def _legacy_reviews_url(url: str) -> str | None:
        parsed = urlparse(url)
        match = re.search(r"/product/([^/]+)/[^/]+/sku/(\d+)/reviews", parsed.path)
        if not match:
            return None
        slug, sku = match.groups()
        query = parse_qs(parsed.query)
        kept = {}
        if "ratings" in query:
            kept["ratings"] = query["ratings"]
        return urlunparse(
            parsed._replace(
                path=f"/site/reviews/{slug}/{sku}",
                query=urlencode(kept, doseq=True),
                fragment="",
            )
        )

    @staticmethod
    def _extract_sku(url: str) -> str | None:
        parsed = urlparse(url)
        for pattern in (r"/sku/(\d+)/reviews", r"/site/reviews/[^/]+/(\d+)"):
            match = re.search(pattern, parsed.path)
            if match:
                return match.group(1)
        return None

    @staticmethod
    def _rating_filter(url: str) -> str | None:
        ratings = parse_qs(urlparse(url).query).get("ratings")
        return ratings[0] if ratings else None

    def _page_url(self, base_url: str, page: int) -> str:
        if page <= 1:
            return base_url
        parsed = urlparse(base_url)
        query = parse_qs(parsed.query)
        query["page"] = [str(page)]
        return urlunparse(parsed._replace(query=urlencode(query, doseq=True)))

    def _find_review_blocks(self, soup):
        for selector in self.REVIEW_BLOCK_SELECTORS:
            blocks = soup.select(selector)
            if blocks:
                return blocks
        return []

    def _text(self, element, selectors: list[str]) -> str | None:
        for selector in selectors:
            node = element.select_one(selector)
            if node:
                text = node.get_text(strip=True) or node.get("content") or node.get("title")
                if text:
                    return str(text).strip()
        return None

    def _parse_block(self, block, *, model_id: str, model_name: str, source_url: str) -> ReviewRecord | None:
        rating_raw = self._text(block, self.RATING_SELECTORS)
        rating = self._normalize_rating(rating_raw) if rating_raw else None
        return ReviewRecord(
            model_id=model_id,
            model_name=model_name,
            site=self.site_name,
            source_url=source_url,
            rating=rating,
            title=self._text(block, self.TITLE_SELECTORS),
            review_text=self._text(block, self.BODY_SELECTORS),
            review_date=self._text(block, self.DATE_SELECTORS),
            author=self._text(block, self.AUTHOR_SELECTORS),
        )

    @staticmethod
    def _normalize_rating(raw: str) -> str:
        match = re.search(r"(\d+(?:\.\d+)?)", raw)
        return match.group(1) if match else raw
