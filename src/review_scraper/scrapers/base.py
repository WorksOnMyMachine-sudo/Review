from __future__ import annotations

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
