"""Scrape reviews, run AI analysis, and print final report path."""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "scripts"))

from review_scraper.pipeline import run_scrape

import analyze_reviews


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run scrape first, then AI analysis report.")
    parser.add_argument("--model", "-m", help="只爬取并分析指定机型 model_id 或名称")
    parser.add_argument("--site", "-s", help="只爬取指定站点，如 amazon / bby / wmt")
    parser.add_argument("--reviews-output", help="抓取结果 Excel 文件名，默认自动生成 reviews_时间.xlsx")
    parser.add_argument(
        "--report-output",
        "-o",
        type=Path,
        default=Path("data/output/review_issue_analysis_latest.xlsx"),
        help="最终分析报告路径",
    )
    parser.add_argument(
        "--classification-mode",
        choices=["semantic", "keyword"],
        default="semantic",
        help="semantic 使用 API 语义分类；keyword 使用本地关键词备用模式",
    )
    parser.add_argument(
        "--category-file",
        "-c",
        type=Path,
        help="Call log 分类表路径，默认 data/output/Call log分类.xlsx",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    reviews_output = args.reviews_output or f"reviews_{datetime.now():%Y%m%d_%H%M%S}.xlsx"

    print("STEP 1/2: 开始抓取网评...")
    reviews_path = run_scrape(
        model_filter=args.model,
        site_filter=args.site,
        output_filename=reviews_output,
    )
    print(f"REVIEWS_OUTPUT: {reviews_path}")

    print("\nSTEP 2/2: 开始 AI 分析并生成报告...")
    analysis_args = argparse.Namespace(
        input=Path(reviews_path),
        output=args.report_output,
        category_file=args.category_file,
        classification_mode=args.classification_mode,
        template_output=None,
    )
    analyze_reviews.run_analysis(analysis_args)
    print(f"REPORT_OUTPUT: {analyze_reviews.resolve_output(args.report_output)}")


if __name__ == "__main__":
    main()
