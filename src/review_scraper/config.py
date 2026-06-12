from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml
from pydantic import BaseModel, Field


def project_root() -> Path:
    return Path(__file__).resolve().parents[2]


class SiteTarget(BaseModel):
    enabled: bool = True
    url: str = ""
    max_pages: int = 5


class ModelConfig(BaseModel):
    model_id: str
    display_name: str
    sites: dict[str, SiteTarget] = Field(default_factory=dict)


class DefaultsConfig(BaseModel):
    request_delay_seconds: float = 1.5
    timeout_seconds: int = 30
    user_agent: str = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    # Amazon：Playwright 登录态文件（复用后无需每次手动登录）
    amazon_storage_state: str | None = "data/amazon_state.json"
    amazon_playwright_headless: bool = True
    # 爬取前自动检查登录态；失效则用 .env 账号尝试无头登录
    amazon_auto_login: bool = True


class AppConfig(BaseModel):
    defaults: DefaultsConfig = Field(default_factory=DefaultsConfig)
    models: list[ModelConfig] = Field(default_factory=list)


def load_config(path: Path | None = None) -> AppConfig:
    config_path = path or (project_root() / "config" / "sites.yaml")
    raw: dict[str, Any] = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    return AppConfig.model_validate(raw)


def iter_scrape_targets(config: AppConfig) -> list[tuple[ModelConfig, str, SiteTarget]]:
    """Yield (model, site_name, site_target) for enabled sites with non-empty URL."""
    targets: list[tuple[ModelConfig, str, SiteTarget]] = []
    for model in config.models:
        for site_name, site in model.sites.items():
            if site.enabled and site.url.strip():
                targets.append((model, site_name, site))
    return targets
