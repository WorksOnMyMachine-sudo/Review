"""Run configured review scrapes and print output path."""
import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from review_scraper.pipeline import run_scrape


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run configured review scrapes.")
    parser.add_argument("--model", "-m", help="只爬指定机型 model_id 或名称")
    parser.add_argument("--site", "-s", help="只爬指定站点，如 amazon / bby / wmt")
    parser.add_argument("--output", "-o", help="输出 Excel 文件名")
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    path = run_scrape(
        model_filter=args.model,
        site_filter=args.site,
        output_filename=args.output,
    )
    print("OUTPUT:", path)
