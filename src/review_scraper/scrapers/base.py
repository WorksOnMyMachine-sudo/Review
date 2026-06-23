from __future__ import annotations

import json
import re
import time
from abc import ABC, abstractmethod

import httpx
from bs4 import BeautifulSoup

from review_scraper.config import DefaultsConfig
from review_scraper.models import ReviewRecord


class BaseScraper(ABC):
    site_name: str = "generic"

    def __init__(self, defaults: DefaultsConfig) -> None:
        self.defaults = defaults

    def fetch_html(self, url: str) -> str:
        headers = {
            "User-Agent": self.defaults.user_agent,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
            "Accept-Encoding": "gzip, deflate, br",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
        }
        with httpx.Client(
            timeout=self.defaults.timeout_seconds,
            follow_redirects=True,
            headers=headers,
        ) as client:
            response = client.get(url)
            response.raise_for_status()
            return response.text

    def parse_soup(self, html: str) -> BeautifulSoup:
        return BeautifulSoup(html, "html.parser")

    def extract_overall_rating(self, soup: BeautifulSoup, selectors: list[str] | None = None) -> str | None:
        rating = self._extract_overall_rating_from_json(soup)
        if rating:
            return rating

        for selector in selectors or []:
            for node in soup.select(selector):
                text = " ".join(
                    part
                    for part in (
                        node.get("content"),
                        node.get("aria-label"),
                        node.get("title"),
                        node.get_text(" ", strip=True),
                    )
                    if part
                )
                rating = self._normalize_overall_rating_text(text)
                if rating:
                    return rating

        return None

    def _extract_overall_rating_from_json(self, soup: BeautifulSoup) -> str | None:
        for script in soup.find_all("script"):
            text = script.string or script.get_text(strip=True)
            if not text:
                continue
            text = text.strip()
            if text.startswith("{") or text.startswith("["):
                try:
                    rating = self._walk_for_aggregate_rating(json.loads(text))
                except json.JSONDecodeError:
                    rating = None
                if rating:
                    return rating

            match = re.search(
                r'"aggregateRating"\s*:\s*\{.*?"ratingValue"\s*:\s*"?(\d+(?:\.\d+)?)"?',
                text,
                flags=re.I | re.S,
            )
            if match:
                return match.group(1)
        return None

    def _walk_for_aggregate_rating(self, value) -> str | None:
        if isinstance(value, dict):
            type_value = value.get("@type") or value.get("type")
            if str(type_value).lower() == "aggregaterating":
                rating = value.get("ratingValue") or value.get("rating") or value.get("value")
                normalized = self._normalize_overall_rating_text(str(rating)) if rating is not None else None
                if normalized:
                    return normalized
            aggregate = value.get("aggregateRating")
            if aggregate is not None:
                rating = self._walk_for_aggregate_rating(aggregate)
                if rating:
                    return rating
            for child in value.values():
                rating = self._walk_for_aggregate_rating(child)
                if rating:
                    return rating
        elif isinstance(value, list):
            for child in value:
                rating = self._walk_for_aggregate_rating(child)
                if rating:
                    return rating
        return None

    @staticmethod
    def _normalize_overall_rating_text(text: str) -> str | None:
        if not text:
            return None
        match = re.search(r"(\d+(?:\.\d+)?)\s*(?:out of\s*)?5", text, flags=re.I)
        if not match:
            match = re.search(r"(\d+(?:\.\d+)?)", text)
        if not match:
            return None
        value = float(match.group(1))
        if 0 <= value <= 5:
            return f"{value:g}"
        return None

    def delay(self) -> None:
        time.sleep(self.defaults.request_delay_seconds)

    @abstractmethod
    def scrape(
        self,
        *,
        url: str,
        model_id: str,
        model_name: str,
        max_pages: int = 5,
    ) -> list[ReviewRecord]:
        raise NotImplementedError
