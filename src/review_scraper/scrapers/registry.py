from __future__ import annotations

from review_scraper.config import DefaultsConfig
from review_scraper.scrapers.base import BaseScraper
from review_scraper.scrapers.amazon import AmazonScraper
from review_scraper.scrapers.bestbuy import BestBuyScraper
from review_scraper.scrapers.generic import GenericScraper
from review_scraper.scrapers.walmart import WalmartScraper

_REGISTRY: dict[str, type[BaseScraper]] = {
    "bestbuy": BestBuyScraper,
    "bby": BestBuyScraper,
    "amazon": AmazonScraper,
    "walmart": WalmartScraper,
    "wmt": WalmartScraper,
}


def get_scraper(site_name: str, defaults: DefaultsConfig) -> BaseScraper:
    key = site_name.lower().strip()
    if key in _REGISTRY:
        return _REGISTRY[key](defaults)
    if key.startswith("amazon_") or key.startswith("amz_"):
        return AmazonScraper(defaults, site_name=key)
    if key.startswith("bby_") or key.startswith("bestbuy_"):
        return BestBuyScraper(defaults, site_name=key)
    if key.startswith("wmt_") or key.startswith("walmart_"):
        return WalmartScraper(defaults, site_name=key)
    if key in GenericScraper.SITE_SELECTORS:
        return GenericScraper(defaults, site_key=key)
    return GenericScraper(defaults, site_key=key)
