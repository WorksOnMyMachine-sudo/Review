"""One-shot: login (from .env) + scrape Amazon reviews."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from review_scraper.amazon_auth import save_storage_state
from review_scraper.pipeline import run_scrape


def main() -> None:
    print("=== 1/2 登录并保存 Cookie（无头模式，约 12 秒等待）===")
    state = save_storage_state(headless=True, manual_fallback=False)
    print(f"登录态: {state}")

    print("\n=== 2/2 爬取评论并导出 Excel ===")
    out = run_scrape(output_filename="amazon_B0GR9NR9XV.xlsx")
    print(f"\n完成: {out}")


if __name__ == "__main__":
    main()
