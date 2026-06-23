from __future__ import annotations

import json
import re
from urllib.parse import parse_qs, urlencode, urlparse, urlunparse

from bs4 import BeautifulSoup

from review_scraper.config import project_root
from review_scraper.models import ReviewRecord
from review_scraper.scrapers.base import BaseScraper


class WalmartScraper(BaseScraper):
    site_name = "walmart"

    REVIEW_BLOCK_SELECTORS = [
        "[data-testid='review-card']",
        "[data-automation-id='review-card']",
        "[itemprop='review']",
        ".BVRRReview",
        ".BVRRContentReview",
        ".bv-content-item",
        "li[data-testid*='review']",
        "div[data-testid*='review']",
    ]
    TITLE_SELECTORS = [
        "[data-testid='review-title']",
        "[data-automation-id='review-title']",
        "[itemprop='name']",
        ".BVRRReviewTitle",
        ".bv-content-title",
        "h3",
    ]
    BODY_SELECTORS = [
        "[data-testid='review-text']",
        "[data-automation-id='review-text']",
        "[itemprop='reviewBody']",
        ".BVRRReviewText",
        ".BVRRReviewTextContainer",
        ".bv-content-summary-body-text",
        "p",
    ]
    DATE_SELECTORS = [
        "time",
        "[data-testid='review-date']",
        "[data-automation-id='review-date']",
        "[itemprop='datePublished']",
        ".BVRRReviewDate",
        ".bv-content-datetime-stamp",
    ]
    AUTHOR_SELECTORS = [
        "[data-testid='review-author']",
        "[data-automation-id='review-author']",
        "[itemprop='author']",
        ".BVRRNickname",
        ".bv-author",
    ]
    RATING_SELECTORS = [
        "[itemprop='ratingValue']",
        "[aria-label*='stars']",
        "[aria-label*='Stars']",
        "[data-testid='review-star-rating']",
        ".BVRRRatingNumber",
        ".bv-rating",
    ]
    OVERALL_RATING_SELECTORS = [
        "[itemprop='aggregateRating'] [itemprop='ratingValue']",
        "[data-testid='reviews-and-ratings'] [aria-label*='stars']",
        "[data-testid='product-ratings'] [aria-label*='stars']",
        "[aria-label*='average rating']",
        ".bv-average-rating",
    ]

    def __init__(self, defaults, site_name: str = "walmart") -> None:
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
        overall_rating = self._fetch_overall_rating(url)
        bazaarvoice_records = self._scrape_bazaarvoice(
            url=url,
            model_id=model_id,
            model_name=model_name,
            max_pages=max_pages,
            overall_rating=overall_rating,
        )
        if bazaarvoice_records:
            return bazaarvoice_records

        records: list[ReviewRecord] = []
        seen: set[tuple[str, str, str, str, str]] = set()
        consecutive_no_data = 0

        for page in range(1, max_pages + 1):
            page_url = self._page_url(url, page)
            html = self._fetch_review_html(page_url)
            if html is None:
                if not records:
                    return self._scrape_playwright(
                        url=url,
                        model_id=model_id,
                        model_name=model_name,
                        max_pages=max_pages,
                        overall_rating=overall_rating,
                    )
                consecutive_no_data += 1
                if consecutive_no_data >= 5:
                    break
                self.delay()
                continue
            page_records = self._parse_page(
                html,
                model_id=model_id,
                model_name=model_name,
                source_url=page_url,
                overall_rating=overall_rating,
            )
            page_records = self._dedupe(page_records, seen)
            if not page_records:
                consecutive_no_data += 1
                if consecutive_no_data >= 5:
                    break
                self.delay()
                continue
            consecutive_no_data = 0
            records.extend(page_records)
            if page < max_pages:
                self.delay()

        return records

    def _scrape_bazaarvoice(
        self,
        *,
        url: str,
        model_id: str,
        model_name: str,
        max_pages: int,
        overall_rating: str | None = None,
    ) -> list[ReviewRecord]:
        item_id = self._extract_item_id(url)
        if not item_id:
            return []

        records: list[ReviewRecord] = []
        seen: set[tuple[str, str, str, str, str]] = set()
        consecutive_empty = 0
        for page in range(1, max_pages + 1):
            endpoint = self._bazaarvoice_url(item_id, page)
            try:
                js = self._fetch_bazaarvoice_js(endpoint)
            except Exception:  # noqa: BLE001
                consecutive_empty += 1
                if consecutive_empty >= 3:
                    break
                self.delay()
                continue

            html = self._decode_escaped_html(js)
            soup = self.parse_soup(html)
            page_records = self._parse_html_reviews(
                soup,
                model_id=model_id,
                model_name=model_name,
                source_url=endpoint,
                overall_rating=overall_rating,
            )
            page_records = self._dedupe(page_records, seen)
            if not page_records:
                consecutive_empty += 1
                if consecutive_empty >= 3:
                    break
                self.delay()
                continue
            consecutive_empty = 0
            records.extend(page_records)
            if page < max_pages:
                self.delay()

        return records

    @staticmethod
    def _bazaarvoice_url(item_id: str, page: int) -> str:
        query = urlencode(
            {
                "format": "embeddedhtml",
                "sort": "submissionTime",
                "dir": "desc",
                "page": page,
            }
        )
        return f"https://walmart.ugc.bazaarvoice.com/1336/{item_id}/reviews.djs?{query}"

    def _fetch_overall_rating(self, url: str) -> str | None:
        for candidate in self._overall_rating_url_candidates(url):
            try:
                html = self._fetch_review_html(candidate)
            except Exception:  # noqa: BLE001
                continue
            if not html:
                continue
            rating = self.extract_overall_rating(
                self.parse_soup(html),
                self.OVERALL_RATING_SELECTORS,
            )
            if rating:
                return rating
        return None

    def _overall_rating_url_candidates(self, url: str) -> list[str]:
        candidates = []
        item_id = self._extract_item_id(url)
        if item_id:
            candidates.append(f"https://www.walmart.com/ip/{item_id}")
            candidates.append(f"https://www.walmart.com/reviews/product/{item_id}")
        candidates.append(url)

        out = []
        seen = set()
        for candidate in candidates:
            if candidate in seen:
                continue
            seen.add(candidate)
            out.append(candidate)
        return out

    def _fetch_bazaarvoice_js(self, url: str) -> str:
        try:
            from curl_cffi import requests

            response = requests.get(
                url,
                impersonate="chrome",
                timeout=self.defaults.timeout_seconds,
                headers={
                    "Accept": "*/*",
                    "Accept-Language": "en-US,en;q=0.9",
                    "Referer": "https://www.walmart.com/",
                },
            )
            response.raise_for_status()
            return response.text
        except ImportError:
            return self.fetch_html(url)

    @staticmethod
    def _decode_escaped_html(value: str) -> str:
        from html import unescape

        return unescape(
            value.replace(r"\/", "/")
            .replace(r"\"", '"')
            .replace(r"\n", "\n")
            .replace(r"\t", "\t")
        )

    def _scrape_playwright(
        self,
        *,
        url: str,
        model_id: str,
        model_name: str,
        max_pages: int,
        overall_rating: str | None = None,
    ) -> list[ReviewRecord]:
        try:
            from playwright.sync_api import sync_playwright
        except ImportError as exc:
            raise RuntimeError(
                "Walmart HTTP 抓取被验证页拦截，且当前环境没有 Playwright。请安装后再试：pip install playwright && playwright install chromium"
            ) from exc

        records: list[ReviewRecord] = []
        seen: set[tuple[str, str, str, str, str]] = set()
        consecutive_no_data = 0
        state_path = self._storage_state_path()
        profile_dir = self._browser_profile_dir()

        with sync_playwright() as p:
            print(f"Walmart HTTP 请求被验证页拦截，改用浏览器模式。Profile: {profile_dir}", flush=True)
            context_kwargs: dict = {
                "user_agent": self.defaults.user_agent,
                "locale": "en-US",
                "viewport": {"width": 1280, "height": 900},
                "headless": False,
                "args": ["--disable-blink-features=AutomationControlled"],
            }
            try:
                context = p.chromium.launch_persistent_context(
                    str(profile_dir),
                    channel="chrome",
                    **context_kwargs,
                )
            except Exception:
                context = p.chromium.launch_persistent_context(
                    str(profile_dir),
                    **context_kwargs,
                )
            page = context.pages[0] if context.pages else context.new_page()
            page.set_default_timeout(self.defaults.timeout_seconds * 1000)
            page.set_default_navigation_timeout(20000)

            for page_number in range(1, max_pages + 1):
                page_url = self._page_url(url, page_number)
                print(f"  Walmart 浏览器抓取第 {page_number} 页: {page_url}", flush=True)
                try:
                    page.goto(page_url, wait_until="domcontentloaded", timeout=20000)
                except Exception as exc:  # noqa: BLE001
                    print(f"  Walmart 第 {page_number} 页打开超时/失败，尝试读取当前页面: {exc}", flush=True)
                html = self._safe_page_content(page)
                if self._looks_like_robot_page(html):
                    self._wait_for_manual_verification(page)
                    context.storage_state(path=str(state_path))
                    html = self._safe_page_content(page)
                    if self._looks_like_robot_page(html):
                        if not records:
                            context.close()
                            raise RuntimeError("Walmart 验证仍未通过，无法读取评论。")
                        consecutive_no_data += 1
                        if consecutive_no_data >= 5:
                            break
                        continue

                try:
                    page.wait_for_selector(
                        ", ".join(self.REVIEW_BLOCK_SELECTORS + self.BODY_SELECTORS),
                        timeout=10000,
                    )
                except Exception:  # noqa: BLE001
                    pass

                html = self._safe_page_content(page)
                page_records = self._parse_page(
                    html,
                    model_id=model_id,
                    model_name=model_name,
                    source_url=page.url,
                    overall_rating=overall_rating,
                )
                page_records = self._dedupe(page_records, seen)
                if not page_records:
                    consecutive_no_data += 1
                    if consecutive_no_data >= 5:
                        break
                    self.delay()
                    continue
                consecutive_no_data = 0
                records.extend(page_records)
                if page_number < max_pages:
                    self.delay()

            context.storage_state(path=str(state_path))
            context.close()

        return records

    def _wait_for_manual_verification(self, page) -> None:
        print(
            "\nWalmart 返回了验证/robot 页面。\n"
            "请在打开的浏览器中完成验证，确认能看到评论页后，回到终端按【回车】继续..."
        )
        try:
            input()
        except EOFError:
            page.wait_for_timeout(30000)

    @staticmethod
    def _safe_page_content(page) -> str:
        try:
            return page.content()
        except Exception:
            return ""

    def _fetch_review_html(self, url: str) -> str | None:
        for _ in range(3):
            html = self._fetch_html_with_retries(url)
            if not self._looks_like_robot_page(html):
                return html
            self.delay()
        return None

    def _fetch_html_with_retries(self, url: str, attempts: int = 3) -> str:
        last_exc: Exception | None = None
        for attempt in range(1, attempts + 1):
            try:
                return self.fetch_html(url)
            except Exception as exc:  # noqa: BLE001
                last_exc = exc
                if attempt < attempts:
                    self.delay()
        raise last_exc or RuntimeError("Walmart request failed")

    @staticmethod
    def _storage_state_path():
        path = project_root() / "data" / "walmart_state.json"
        path.parent.mkdir(parents=True, exist_ok=True)
        return path

    @staticmethod
    def _browser_profile_dir():
        path = project_root() / "data" / "walmart_browser_profile"
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _parse_page(
        self,
        html: str,
        *,
        model_id: str,
        model_name: str,
        source_url: str,
        overall_rating: str | None = None,
    ) -> list[ReviewRecord]:
        soup = self.parse_soup(html)
        page_overall_rating = overall_rating or self.extract_overall_rating(
            soup,
            self.OVERALL_RATING_SELECTORS,
        )
        records = self._parse_json_reviews(
            soup,
            model_id=model_id,
            model_name=model_name,
            source_url=source_url,
            overall_rating=page_overall_rating,
        )
        if records:
            return records
        return self._parse_html_reviews(
            soup,
            model_id=model_id,
            model_name=model_name,
            source_url=source_url,
            overall_rating=page_overall_rating,
        )

    def _parse_json_reviews(
        self,
        soup: BeautifulSoup,
        *,
        model_id: str,
        model_name: str,
        source_url: str,
        overall_rating: str | None = None,
    ) -> list[ReviewRecord]:
        records: list[ReviewRecord] = []
        for payload in self._json_payloads(soup):
            for item in self._walk_review_objects(payload):
                record = self._record_from_json(
                    item,
                    model_id=model_id,
                    model_name=model_name,
                    source_url=source_url,
                    overall_rating=overall_rating,
                )
                if record:
                    records.append(record)
        return self._dedupe(records, set())

    def _json_payloads(self, soup: BeautifulSoup):
        for script in soup.find_all("script"):
            text = script.string or script.get_text(strip=True)
            if not text:
                continue
            text = text.strip()
            if not (text.startswith("{") or text.startswith("[")):
                continue
            try:
                yield json.loads(text)
            except json.JSONDecodeError:
                continue

    def _walk_review_objects(self, value):
        if isinstance(value, dict):
            if self._looks_like_review(value):
                yield value
            for child in value.values():
                yield from self._walk_review_objects(child)
        elif isinstance(value, list):
            for child in value:
                yield from self._walk_review_objects(child)

    @staticmethod
    def _looks_like_review(item: dict) -> bool:
        keys = {str(key).lower() for key in item}
        has_text = bool(
            keys
            & {
                "reviewtext",
                "review_text",
                "reviewbody",
                "body",
                "text",
                "comments",
                "description",
            }
        )
        has_rating = bool(keys & {"rating", "reviewrating", "ratingvalue", "stars"})
        has_review_id = bool(keys & {"reviewid", "review_id", "id"})
        return has_text and (has_rating or has_review_id)

    def _record_from_json(
        self,
        item: dict,
        *,
        model_id: str,
        model_name: str,
        source_url: str,
        overall_rating: str | None = None,
    ) -> ReviewRecord | None:
        body = self._pick_text(
            item,
            "reviewText",
            "review_text",
            "reviewBody",
            "body",
            "text",
            "comments",
            "description",
        )
        title = self._pick_text(item, "reviewTitle", "title", "summary", "name")
        if not body and not title:
            return None

        rating_raw = self._pick_value(item, "rating", "reviewRating", "ratingValue", "stars")
        author = self._pick_text(item, "userNickname", "nickname", "author", "userName", "name")
        review_date = self._pick_text(
            item,
            "submissionTime",
            "submittedDate",
            "createdDate",
            "datePublished",
            "date",
        )
        verified = self._pick_text(item, "badges", "syndicationSource", "verifiedPurchaser")
        if self._is_badge_only(body=body, title=title, rating=rating_raw, author=author, review_date=review_date):
            return None

        return ReviewRecord(
            model_id=model_id,
            model_name=model_name,
            site=self.site_name,
            source_url=source_url,
            rating=self._normalize_rating(rating_raw),
            overall_rating=overall_rating,
            title=title,
            review_text=body,
            review_date=review_date,
            author=author,
            verified_purchase=verified,
        )

    def _parse_html_reviews(
        self,
        soup: BeautifulSoup,
        *,
        model_id: str,
        model_name: str,
        source_url: str,
        overall_rating: str | None = None,
    ) -> list[ReviewRecord]:
        records: list[ReviewRecord] = []
        blocks = self._find_review_blocks(soup)
        for block in blocks:
            record = ReviewRecord(
                model_id=model_id,
                model_name=model_name,
                site=self.site_name,
                source_url=source_url,
                rating=self._normalize_rating(self._text(block, self.RATING_SELECTORS)),
                overall_rating=overall_rating,
                title=self._text(block, self.TITLE_SELECTORS),
                review_text=self._text(block, self.BODY_SELECTORS),
                review_date=self._text(block, self.DATE_SELECTORS),
                author=self._text(block, self.AUTHOR_SELECTORS),
            )
            if record.review_text or record.title:
                records.append(record)
        return records

    def _find_review_blocks(self, soup: BeautifulSoup):
        for selector in self.REVIEW_BLOCK_SELECTORS:
            blocks = soup.select(selector)
            if blocks:
                return blocks
        return []

    def _text(self, element, selectors: list[str]) -> str | None:
        for selector in selectors:
            node = element.select_one(selector)
            if not node:
                continue
            text = node.get("content") or node.get("aria-label") or node.get_text(" ", strip=True)
            if text:
                return str(text).strip()
        return None

    @staticmethod
    def _pick_value(item: dict, *keys: str):
        lowered = {str(key).lower(): value for key, value in item.items()}
        for key in keys:
            if key.lower() in lowered:
                return lowered[key.lower()]
        return None

    @classmethod
    def _pick_text(cls, item: dict, *keys: str) -> str | None:
        value = cls._pick_value(item, *keys)
        if value is None:
            return None
        if isinstance(value, dict):
            for nested_key in ("text", "name", "value", "label"):
                nested = cls._pick_value(value, nested_key)
                if nested is not None:
                    return cls._clean_text(nested)
            return None
        if isinstance(value, list):
            parts = [cls._clean_text(part) for part in value]
            parts = [part for part in parts if part]
            return ", ".join(parts) if parts else None
        return cls._clean_text(value)

    @staticmethod
    def _clean_text(value) -> str | None:
        text = str(value).strip()
        text = re.sub(r"\s+", " ", text)
        return text or None

    @staticmethod
    def _normalize_rating(raw) -> str | None:
        if raw is None:
            return None
        if isinstance(raw, dict):
            for key in ("ratingValue", "value", "rating"):
                if key in raw:
                    return WalmartScraper._normalize_rating(raw[key])
            return None
        match = re.search(r"(\d+(?:\.\d+)?)", str(raw))
        return match.group(1) if match else str(raw).strip()

    @staticmethod
    def _is_badge_only(*, body: str | None, title: str | None, rating, author: str | None, review_date: str | None) -> bool:
        text = (body or title or "").strip().lower()
        badge_texts = {
            "verified purchase",
            "walmart associate",
            "incentivized review",
            "received free product",
        }
        return bool(text in badge_texts and rating is None and not author and not review_date)

    @staticmethod
    def _looks_like_robot_page(html: str) -> bool:
        lower = html.lower()
        has_review_marker = (
            "review-card" in lower
            or "customerreviews" in lower
            or "reviewtext" in lower
            or '"reviews"' in lower
        )
        has_robot_marker = (
            "captcha" in lower
            or "robot" in lower
            or "blocked" in lower
            or "verify your identity" in lower
            or "press & hold" in lower
        )
        return has_robot_marker and not has_review_marker

    def _page_url(self, base_url: str, page: int) -> str:
        item_id = self._extract_item_id(base_url)
        if item_id:
            query = {"page": str(page), "sort": "submission-desc"}
            return f"https://www.walmart.com/reviews/product/{item_id}?{urlencode(query)}"

        parsed = urlparse(base_url)
        query = parse_qs(parsed.query)
        query["page"] = [str(page)]
        return urlunparse(parsed._replace(query=urlencode(query, doseq=True)))

    @staticmethod
    def _extract_item_id(url: str) -> str | None:
        parsed = urlparse(url)
        for pattern in (r"/reviews/product/(\d+)", r"/ip/(?:[^/]+/)?(\d+)"):
            match = re.search(pattern, parsed.path)
            if match:
                return match.group(1)
        return None

    @staticmethod
    def _dedupe(records: list[ReviewRecord], seen: set[tuple[str, str, str, str, str]]) -> list[ReviewRecord]:
        out: list[ReviewRecord] = []
        for record in records:
            key = (
                (record.rating or "").strip(),
                (record.title or "").strip().lower(),
                (record.review_text or "").strip().lower()[:200],
                (record.review_date or "").strip(),
                (record.author or "").strip().lower(),
            )
            if key in seen:
                continue
            seen.add(key)
            out.append(record)
        return out
