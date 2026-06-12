from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class ReviewRecord:
    model_id: str
    model_name: str
    site: str
    source_url: str
    rating: str | None = None
    title: str | None = None
    review_text: str | None = None
    review_date: str | None = None
    author: str | None = None
    verified_purchase: str | None = None
    scraped_at: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))

    @property
    def channel(self) -> str:
        site = self.site.lower().strip()
        if site.startswith("amazon_") or site.startswith("amz_"):
            return "AMZ"
        if site.startswith("bby_") or site.startswith("bestbuy_"):
            return "BBY"
        if site.startswith("wmt_") or site.startswith("walmart_"):
            return "WMT"
        mapping = {
            "amazon": "AMZ",
            "amz": "AMZ",
            "bestbuy": "BBY",
            "bby": "BBY",
            "walmart": "WMT",
            "wmt": "WMT",
            "csc": "CSC",
        }
        return mapping.get(site, self.site.upper())

    def as_row(self) -> dict:
        return {
            "机型ID": self.model_id,
            "机型名称": self.model_name,
            "Channel": self.channel,
            "来源站点": self.site,
            "来源URL": self.source_url,
            "评分": self.rating or "",
            "标题": self.title or "",
            "评论内容": self.review_text or "",
            "评论日期": self.review_date or "",
            "用户": self.author or "",
            "Verified": self.verified_purchase or "",
            "抓取时间": self.scraped_at,
        }
