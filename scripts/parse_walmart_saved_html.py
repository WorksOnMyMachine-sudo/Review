"""Parse manually saved Walmart review HTML files and export reviews to Excel."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from review_scraper.config import DefaultsConfig
from review_scraper.export.excel import export_reviews_to_excel
from review_scraper.models import ReviewRecord
from review_scraper.scrapers.walmart import WalmartScraper


DEFAULT_INPUT_DIR = ROOT / "data" / "walmart_html"
DEFAULT_MODELS = ("32H40G", "43H40G")


def main() -> None:
    args = parse_args()
    input_dir = args.input if args.input.is_absolute() else ROOT / args.input
    scraper = WalmartScraper(DefaultsConfig(), site_name="wmt")

    records: list[ReviewRecord] = []
    for model_id in args.models:
        model_dir = input_dir / model_id
        if not model_dir.exists():
            safe_print(f"WARNING: 未找到目录，跳过: {model_dir}", stream=sys.stderr)
            continue

        model_records = parse_model_dir(scraper, model_dir, model_id)
        safe_print(f"{model_id}: 从 {model_dir} 解析到 {len(model_records)} 条评论")
        records.extend(model_records)

    output = export_reviews_to_excel(records, filename=args.output)
    safe_print(f"OUTPUT: {output}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Parse manually saved Walmart review HTML files."
    )
    parser.add_argument(
        "--input",
        "-i",
        type=Path,
        default=DEFAULT_INPUT_DIR,
        help=r"HTML 根目录，默认 data\walmart_html。每个机型一个子目录。",
    )
    parser.add_argument(
        "--model",
        "-m",
        dest="models",
        action="append",
        choices=DEFAULT_MODELS,
        help="只解析指定机型；可重复传。默认解析 32H40G 和 43H40G。",
    )
    parser.add_argument(
        "--output",
        "-o",
        default="reviews_wmt_saved_html.xlsx",
        help="输出 Excel 文件名，默认 reviews_wmt_saved_html.xlsx。",
    )
    args = parser.parse_args()
    if not args.models:
        args.models = list(DEFAULT_MODELS)
    return args


def parse_model_dir(scraper: WalmartScraper, model_dir: Path, model_id: str) -> list[ReviewRecord]:
    records: list[ReviewRecord] = []
    seen: set[tuple[str, str, str, str, str]] = set()
    html_files = sorted(
        [
            *model_dir.glob("*.html"),
            *model_dir.glob("*.htm"),
        ],
        key=lambda path: natural_sort_key(path.name),
    )
    for html_file in html_files:
        html = html_file.read_text(encoding="utf-8", errors="ignore")
        page_records = scraper._parse_page(  # noqa: SLF001
            html,
            model_id=model_id,
            model_name=model_id,
            source_url=str(html_file),
        )
        page_records = scraper._dedupe(page_records, seen)  # noqa: SLF001
        records.extend(page_records)
    return records


def natural_sort_key(value: str) -> list[int | str]:
    parts: list[int | str] = []
    text = ""
    for char in value:
        if char.isdigit():
            if text and not text[-1].isdigit():
                parts.append(text.lower())
                text = ""
            text += char
        else:
            if text and text[-1].isdigit():
                parts.append(int(text))
                text = ""
            text += char
    if text:
        parts.append(int(text) if text.isdigit() else text.lower())
    return parts


def safe_print(value: str, *, stream=None) -> None:
    target = stream or sys.stdout
    try:
        print(value, file=target)
    except UnicodeEncodeError:
        encoded = value.encode(target.encoding or "utf-8", errors="backslashreplace")
        print(encoded.decode(target.encoding or "utf-8"), file=target)


if __name__ == "__main__":
    main()
