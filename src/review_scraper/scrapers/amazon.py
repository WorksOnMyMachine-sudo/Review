from __future__ import annotations

import re
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse, urlunparse

from review_scraper.config import project_root
from review_scraper.models import ReviewRecord
from review_scraper.scrapers.base import BaseScraper


class AmazonScraper(BaseScraper):
    site_name = "amazon"
    REVIEW_BLOCK_SELECTORS = [
        "motion-review",
        "[data-hook='review']",
        "[id^='customer_review-']",
        "#cm_cr-review_list .review",
        "div.review",
    ]
    REVIEW_CONTENT_SELECTORS = [
        "span[data-hook='review-body']",
        "a[data-hook='review-title']",
        "i[data-hook='review-star-rating']",
        "i[data-hook='cmps-review-star-rating']",
        "[id^='customer_review-']",
    ]
    NEXT_PAGE_SELECTORS = [
        "li.a-last a",
        "a[data-hook='pagination-bar-next-link']",
        ".a-pagination .a-last a",
    ]
    SHOW_MORE_SELECTORS = [
        "button:has-text('Show 10 more reviews')",
        "a:has-text('Show 10 more reviews')",
        "button:has-text('Show more reviews')",
        "a:has-text('Show more reviews')",
        "input[value*='Show 10 more']",
        "[data-hook='show-more-reviews-button']",
    ]

    def __init__(self, defaults, site_name: str = "amazon") -> None:
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
        asin = self._extract_asin(url) or model_id
        state_path = self._storage_state_path()
        if state_path and state_path.exists():
            return self._scrape_playwright(url, asin, model_id, model_name, max_pages)
        http_records = self._scrape_http(url, asin, model_id, model_name, max_pages)
        if http_records:
            return http_records
        return self._scrape_playwright(url, asin, model_id, model_name, max_pages)

    def _scrape_http(
        self,
        base_url: str,
        asin: str,
        model_id: str,
        model_name: str,
        max_pages: int,
    ) -> list[ReviewRecord]:
        records: list[ReviewRecord] = []
        seen: set[str] = set()
        for page in range(1, max_pages + 1):
            page_url = self._build_page_url(base_url, asin, page)
            try:
                html = self.fetch_html(page_url)
            except Exception:  # noqa: BLE001
                return []
            if self._is_hard_block(html):
                return []
            soup = self.parse_soup(html)
            page_records = self._records_from_soup(
                soup, model_id, model_name, page_url, seen
            )
            if not page_records:
                break
            records.extend(page_records)
            if page < max_pages:
                self.delay()
        return records

    def _scrape_playwright(
        self,
        base_url: str,
        asin: str,
        model_id: str,
        model_name: str,
        max_pages: int,
    ) -> list[ReviewRecord]:
        try:
            from playwright.sync_api import sync_playwright
        except ImportError as exc:
            raise RuntimeError(
                "需要 Playwright：pip install playwright && playwright install chromium"
            ) from exc

        state_path = self._storage_state_path()
        records: list[ReviewRecord] = []
        seen: set[str] = set()

        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=self.defaults.amazon_playwright_headless,
                args=["--disable-blink-features=AutomationControlled"],
            )
            context_kwargs: dict = {
                "user_agent": self.defaults.user_agent,
                "locale": "en-US",
                "viewport": {"width": 1280, "height": 900},
            }
            if state_path and state_path.exists():
                context_kwargs["storage_state"] = str(state_path)

            context = browser.new_context(**context_kwargs)
            page = context.new_page()
            page.set_default_timeout(self.defaults.timeout_seconds * 1000)

            current_url = self._build_page_url(base_url, asin, 1)
            page.goto(current_url, wait_until="domcontentloaded")
            for pg in range(1, max_pages + 1):
                try:
                    page.wait_for_selector(
                        ", ".join(self.REVIEW_BLOCK_SELECTORS + self.REVIEW_CONTENT_SELECTORS),
                        timeout=25000,
                    )
                except Exception:  # noqa: BLE001
                    if pg == 1:
                        html = page.content()
                        if self._is_auth_or_robot_page(html, page.url):
                            browser.close()
                            hint = (
                                "Amazon 登录态可能已过期或需要验证。\n"
                                "请在弹出的浏览器中完成登录/验证后继续。"
                            )
                            raise RuntimeError(hint) from None
                        browser.close()
                        return records
                    break

                html = page.content()
                if self._is_auth_or_robot_page(html, page.url) and pg == 1:
                    browser.close()
                    raise RuntimeError(
                        "Amazon 登录态可能已过期或需要验证。"
                    )

                before_total = len(records)
                soup = self.parse_soup(html)
                page_records = self._records_from_soup(
                    soup, model_id, model_name, page.url, seen
                )
                if page_records:
                    records.extend(page_records)

                if not page_records and pg == 1:
                    break

                if pg < max_pages:
                    if self._show_more_reviews(page):
                        self.delay()
                        continue
                    if self._go_to_next_review_page(page):
                        current_url = page.url
                        self.delay()
                        continue
                    if len(records) == before_total:
                        break
                    self.delay()
                    break

            browser.close()

        return records

    def _show_more_reviews(self, page) -> bool:
        old_count = self._review_content_count(page)
        for selector in self.SHOW_MORE_SELECTORS:
            control = page.locator(selector).first
            try:
                if not control.count() or not control.is_visible() or not control.is_enabled():
                    continue
                control.scroll_into_view_if_needed()
                control.click()
                try:
                    page.wait_for_function(
                        """oldCount => {
                            const selectors = [
                                "span[data-hook='review-body']",
                                "a[data-hook='review-title']",
                                "[id^='customer_review-']",
                                "[data-hook='review']"
                            ];
                            const total = selectors.reduce(
                                (sum, sel) => sum + document.querySelectorAll(sel).length,
                                0
                            );
                            return total > oldCount;
                        }""",
                        arg=old_count,
                        timeout=15000,
                    )
                except Exception:  # noqa: BLE001
                    page.wait_for_timeout(3000)
                return self._review_content_count(page) > old_count
            except Exception:  # noqa: BLE001
                continue
        return False

    @staticmethod
    def _review_content_count(page) -> int:
        return page.locator(
            "span[data-hook='review-body'], "
            "a[data-hook='review-title'], "
            "[id^='customer_review-'], "
            "[data-hook='review']"
        ).count()

    def _go_to_next_review_page(self, page) -> bool:
        for selector in self.NEXT_PAGE_SELECTORS:
            link = page.locator(selector).first
            try:
                if not link.count() or not link.is_visible() or not link.is_enabled():
                    continue
                old_url = page.url
                link.click()
                page.wait_for_load_state("domcontentloaded")
                try:
                    page.wait_for_url(lambda url: url != old_url, timeout=8000)
                except Exception:  # noqa: BLE001
                    page.wait_for_timeout(2000)
                return True
            except Exception:  # noqa: BLE001
                continue
        return False

    def _storage_state_path(self) -> Path | None:
        raw = self.defaults.amazon_storage_state
        if not raw:
            return None
        path = Path(raw)
        return path if path.is_absolute() else project_root() / path

    def _records_from_soup(
        self,
        soup,
        model_id: str,
        model_name: str,
        page_url: str,
        seen: set[str],
    ) -> list[ReviewRecord]:
        blocks = []
        for selector in self.REVIEW_BLOCK_SELECTORS:
            blocks = soup.select(selector)
            if blocks:
                break
        if not blocks:
            blocks = self._blocks_from_review_content(soup)
        out: list[ReviewRecord] = []
        for block in blocks:
            record = self._parse_block(
                block,
                model_id=model_id,
                model_name=model_name,
                source_url=page_url,
            )
            if not record:
                continue
            key = self._dedupe_key(block, record)
            if key in seen:
                continue
            seen.add(key)
            if record.review_text or record.title:
                out.append(record)
        return out

    @staticmethod
    def _dedupe_key(block, record: ReviewRecord) -> str:
        review_id = block.get("id")
        if review_id:
            return f"id:{review_id}"

        body = (record.review_text or "").strip()
        body_key = body[:160] if body else ""
        return "|".join(
            [
                record.author or "",
                record.review_date or "",
                record.rating or "",
                record.title or "",
                body_key,
            ]
        )

    def _blocks_from_review_content(self, soup) -> list:
        blocks = []
        seen_ids: set[int] = set()
        for node in soup.select(", ".join(self.REVIEW_CONTENT_SELECTORS)):
            block = (
                node.find_parent(attrs={"data-hook": "review"})
                or node.find_parent(id=re.compile(r"^customer_review-"))
                or node.find_parent("li")
                or node.find_parent("div", class_=re.compile(r"\breview\b"))
                or node.find_parent("div")
            )
            if not block:
                continue
            key = id(block)
            if key in seen_ids:
                continue
            seen_ids.add(key)
            blocks.append(block)
        return blocks

    @staticmethod
    def _extract_asin(url: str) -> str | None:
        match = re.search(r"/(?:product-reviews|dp)/([A-Z0-9]{10})", url, re.I)
        return match.group(1).upper() if match else None

    def _build_page_url(self, base_url: str, asin: str, page: int) -> str:
        parsed = urlparse(base_url)
        query = parse_qs(parsed.query)
        query["pageNumber"] = [str(page)]
        if "reviewerType" not in query:
            query["reviewerType"] = ["all_reviews"]
        if "filterByStar" not in query:
            query["filterByStar"] = ["all_stars"]
        if "ie" not in query:
            query["ie"] = ["UTF8"]
        return urlunparse(parsed._replace(query=urlencode(query, doseq=True)))

    @staticmethod
    def _is_hard_block(html: str) -> bool:
        if (
            "data-hook=\"review\"" in html
            or "data-hook='review'" in html
            or "data-hook=\"review-body\"" in html
            or "data-hook='review-body'" in html
            or "customer_review-" in html
        ):
            return False
        lower = html.lower()
        if "amazon sign-in" in lower:
            return True
        if "ap_captcha" in lower or "image-captcha" in lower:
            return True
        if "robot check" in lower or "not a robot" in lower:
            return True
        return False

    @staticmethod
    def _is_auth_or_robot_page(html: str, url: str = "") -> bool:
        lower = html.lower()
        url_lower = url.lower()
        if "ap/signin" in url_lower or "signin" in url_lower:
            return True
        if "amazon sign-in" in lower:
            return True
        if "ap_captcha" in lower or "image-captcha" in lower:
            return True
        if "robot check" in lower or "not a robot" in lower:
            return True
        return False

    def _parse_block(self, block, *, model_id: str, model_name: str, source_url: str) -> ReviewRecord | None:
        title_node = block.select_one("a[data-hook='review-title'] span") or block.select_one(
            "a[data-hook='review-title']"
        )
        body_node = block.select_one("span[data-hook='review-body'] span") or block.select_one(
            "span[data-hook='review-body']"
        )
        date_node = block.select_one("span[data-hook='review-date']")
        author_node = block.select_one("span.a-profile-name") or block.select_one(
            ".a-profile-name"
        )
        rating_node = block.select_one("i[data-hook='review-star-rating'] span") or block.select_one(
            "i[data-hook='cmps-review-star-rating'] span"
        ) or block.select_one("span.a-icon-alt")

        title = title_node.get_text(strip=True) if title_node else None
        body = body_node.get_text(strip=True) if body_node else None
        review_date = date_node.get_text(strip=True) if date_node else None
        author = author_node.get_text(strip=True) if author_node else None
        rating = None
        if rating_node:
            raw = rating_node.get_text(strip=True) or rating_node.get("aria-label", "")
            match = re.search(r"(\d+(?:\.\d+)?)", raw)
            rating = match.group(1) if match else raw

        verified = None
        vp = block.select_one("span[data-hook='avp-badge']")
        if vp:
            verified = vp.get_text(strip=True)

        return ReviewRecord(
            model_id=model_id,
            model_name=model_name,
            site=self.site_name,
            source_url=source_url,
            rating=rating,
            title=title,
            review_text=body,
            review_date=review_date,
            author=author,
            verified_purchase=verified,
        )
