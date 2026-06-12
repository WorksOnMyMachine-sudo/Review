"""
保存 Amazon 登录态（打开浏览器后手动登录，避免自动填表触发风控）。

用法（在项目根目录）:
  python scripts/save_amazon_state.py
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from review_scraper.amazon_auth import save_storage_state


def main() -> None:
    path = save_storage_state(
        headless=False,
        manual_fallback=True,
        auto_fill_credentials=False,
    )
    print(f"已保存: {path}")
    print("然后运行: python scripts/run_scraper.py")


if __name__ == "__main__":
    main()
