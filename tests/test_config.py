from pathlib import Path

from review_scraper.config import load_config, project_root


def test_load_default_config() -> None:
    path = project_root() / "config" / "sites.yaml"
    config = load_config(path)
    assert len(config.models) >= 1
    assert "bestbuy" in config.models[0].sites
