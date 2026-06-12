"""Analyze low-rating reviews by model and export issue summaries to Excel."""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import Counter
from copy import copy
from datetime import datetime
from pathlib import Path

import pandas as pd
import httpx
from dotenv import load_dotenv
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "data" / "output"
DEFAULT_CATEGORY_FILE = OUTPUT_DIR / "Call log分类.xlsx"
load_dotenv(ROOT / ".env", override=True)

REQUIRED_COLUMNS = {
    "机型ID",
    "机型名称",
    "Channel",
    "来源站点",
    "来源URL",
    "评分",
    "标题",
    "评论内容",
    "评论日期",
    "用户",
}

FALLBACK_ISSUE_RULES: list[tuple[str, list[str]]] = [
    (
        "画质 / 屏幕显示",
        [
            "picture",
            "image",
            "display",
            "screen",
            "panel",
            "color",
            "brightness",
            "contrast",
            "black level",
            "backlight",
            "blooming",
            "glare",
            "reflection",
            "matte",
            "pixel",
            "dead pixel",
            "hdr",
            "qled",
        ],
    ),
    (
        "软件 / 系统 / App",
        [
            "software",
            "firmware",
            "update",
            "os",
            "google tv",
            "fire tv",
            "app",
            "apps",
            "streaming",
            "netflix",
            "youtube",
            "crash",
            "freeze",
            "lag",
            "slow",
            "bug",
            "glitch",
            "restart",
            "reboot",
        ],
    ),
    (
        "连接 / 端口 / 网络",
        [
            "wifi",
            "wi-fi",
            "internet",
            "network",
            "bluetooth",
            "hdmi",
            "arc",
            "earc",
            "usb",
            "port",
            "connect",
            "connection",
            "disconnect",
            "pairing",
            "signal",
        ],
    ),
    (
        "声音 / 音频",
        [
            "sound",
            "audio",
            "speaker",
            "volume",
            "bass",
            "dialog",
            "voice",
            "mute",
            "surround",
            "soundbar",
        ],
    ),
    (
        "遥控器 / 操作体验",
        [
            "remote",
            "button",
            "menu",
            "interface",
            "navigation",
            "control",
            "input",
            "settings",
            "voice control",
        ],
    ),
    (
        "安装 / 挂装 / 设置",
        [
            "setup",
            "install",
            "installation",
            "mount",
            "wall mount",
            "stand",
            "assemble",
            "assembly",
            "instructions",
            "manual",
            "calibration",
        ],
    ),
    (
        "质量 / 故障 / 损坏",
        [
            "defect",
            "defective",
            "broken",
            "damage",
            "damaged",
            "dead",
            "stopped working",
            "doesn't work",
            "not working",
            "failed",
            "failure",
            "issue",
            "problem",
            "returned",
            "replacement",
        ],
    ),
    (
        "配送 / 包装",
        [
            "delivery",
            "delivered",
            "shipping",
            "ship",
            "box",
            "package",
            "packaging",
            "arrived",
            "carrier",
            "fedex",
            "ups",
        ],
    ),
    (
        "客服 / 保修 / 退换货",
        [
            "customer service",
            "support",
            "warranty",
            "repair",
            "refund",
            "return",
            "exchange",
            "best buy",
            "hisense support",
        ],
    ),
    (
        "价格 / 价值感",
        [
            "price",
            "expensive",
            "cheap",
            "value",
            "worth",
            "money",
            "cost",
            "deal",
        ],
    ),
]

STOPWORDS = {
    "a",
    "an",
    "and",
    "are",
    "as",
    "be",
    "but",
    "by",
    "can",
    "could",
    "for",
    "from",
    "has",
    "have",
    "in",
    "is",
    "it",
    "no",
    "not",
    "of",
    "on",
    "or",
    "other",
    "please",
    "some",
    "the",
    "this",
    "to",
    "tv",
    "with",
}


def main() -> None:
    args = parse_args()
    input_path = resolve_input(args.input)
    output_path = resolve_output(args.output)
    category_path = resolve_category_file(args.category_file)
    issue_rules = load_issue_rules(category_path)

    reviews, dedupe_stats = load_reviews(input_path)
    write_deduped_reviews(reviews, input_path)
    low_reviews = reviews[reviews["评分_num"].le(3)].copy()
    low_reviews["问题摘要"] = low_reviews.apply(make_issue_summary, axis=1)
    if args.classification_mode == "export-template":
        template_path = resolve_template_output(args.template_output)
        write_classification_template(low_reviews, issue_rules, template_path, input_path, category_path)
        safe_print(f"INPUT: {input_path}")
        safe_print(f"DEDUPED_REVIEWS: {input_path}")
        safe_print(f"CATEGORY: {category_path or 'built-in fallback rules'}")
        safe_print("CLASSIFICATION_MODE: export-template")
        safe_print(f"TEMPLATE_OUTPUT: {template_path}")
        return

    if args.classification_mode == "semantic":
        issue_matches = pd.Series(
            classify_reviews_semantic(low_reviews, issue_rules),
            index=low_reviews.index,
        )
    else:
        issue_matches = low_reviews.apply(lambda row: classify_review_keyword(row, issue_rules), axis=1)
    low_reviews["问题分类"] = issue_matches.map(lambda item: item[0])
    low_reviews["二级问题分类"] = issue_matches.map(lambda item: item[1])
    low_reviews["问题摘要中文"] = translate_summaries(low_reviews["问题摘要"].tolist())

    write_analysis(low_reviews, output_path, input_path, dedupe_stats, issue_rules, category_path)
    safe_print(f"INPUT: {input_path}")
    safe_print(f"DEDUPED_REVIEWS: {input_path}")
    safe_print(f"CATEGORY: {category_path or 'built-in fallback rules'}")
    safe_print(f"CLASSIFICATION_MODE: {args.classification_mode}")
    safe_print(f"OUTPUT: {output_path}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Analyze <=3-star reviews and export issue summaries by model."
    )
    parser.add_argument(
        "--input",
        "-i",
        type=Path,
        help="Input review Excel. Defaults to the latest reviews_*.xlsx in data/output.",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        help="Output analysis Excel. Defaults to data/output/review_issue_analysis_<timestamp>.xlsx.",
    )
    parser.add_argument(
        "--category-file",
        "-c",
        type=Path,
        help="Call log category Excel. Defaults to data/output/Call log分类.xlsx.",
    )
    parser.add_argument(
        "--classification-mode",
        choices=["semantic", "keyword", "export-template"],
        default="semantic",
        help=(
            "semantic uses OpenAI API to understand full review meaning; "
            "keyword uses local fallback matching; export-template creates a ChatGPT upload template."
        ),
    )
    parser.add_argument(
        "--template-output",
        type=Path,
        help="Output Excel for --classification-mode export-template.",
    )
    return parser.parse_args()


def resolve_input(path: Path | None) -> Path:
    if path:
        if not path.exists():
            raise FileNotFoundError(path)
        return path

    candidates = sorted(
        (
            path
            for path in OUTPUT_DIR.glob("reviews_*.xlsx")
            if "__deduped_tmp" not in path.stem and not path.stem.endswith("_deduped")
        ),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not candidates:
        raise FileNotFoundError(f"No reviews_*.xlsx found in {OUTPUT_DIR}")
    return candidates[0]


def resolve_output(path: Path | None) -> Path:
    if path:
        return path if path.is_absolute() else ROOT / path
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return OUTPUT_DIR / f"review_issue_analysis_{stamp}.xlsx"


def resolve_template_output(path: Path | None) -> Path:
    if path:
        return path if path.is_absolute() else ROOT / path
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return OUTPUT_DIR / f"review_classification_template_{stamp}.xlsx"


def resolve_category_file(path: Path | None) -> Path | None:
    if path:
        resolved = path if path.is_absolute() else ROOT / path
        if not resolved.exists():
            raise FileNotFoundError(resolved)
        return resolved
    return DEFAULT_CATEGORY_FILE if DEFAULT_CATEGORY_FILE.exists() else None


def load_issue_rules(category_path: Path | None) -> list[tuple[str, list[str]]]:
    if not category_path:
        return FALLBACK_ISSUE_RULES

    workbook = pd.ExcelFile(category_path)
    rules: list[tuple[str, list[str]]] = []
    for sheet_name in workbook.sheet_names:
        df = pd.read_excel(category_path, sheet_name=sheet_name)
        for column in df.columns:
            category = str(column).strip()
            if not category or category.lower().startswith("unnamed"):
                continue
            phrases = [
                str(value).strip()
                for value in df[column].dropna().tolist()
                if str(value).strip()
            ]
            if phrases:
                rules.append((category, phrases))

    return rules or FALLBACK_ISSUE_RULES


def load_reviews(path: Path) -> tuple[pd.DataFrame, dict[str, int]]:
    workbook = pd.ExcelFile(path)
    frames = []
    for sheet in workbook.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        if df.empty or not REQUIRED_COLUMNS.issubset(df.columns):
            continue
        frames.append(df)

    if not frames:
        raise ValueError(f"No review sheets with expected columns found in {path}")

    data = pd.concat(frames, ignore_index=True)
    original_count = len(data)
    data = dedupe_reviews(data)
    deduped_count = len(data)
    data["评分_num"] = data["评分"].map(parse_rating)
    stats = {
        "original_count": original_count,
        "deduped_count": deduped_count,
        "removed_count": original_count - deduped_count,
    }
    return data[data["评分_num"].notna()].copy(), stats


def dedupe_reviews(df: pd.DataFrame) -> pd.DataFrame:
    dedupe_cols = [
        "机型ID",
        "Channel",
        "来源站点",
        "标题",
        "评论内容",
        "评论日期",
        "用户",
    ]
    out = df.copy()
    key_cols = []
    for col in dedupe_cols:
        key_col = f"_dedupe_{col}"
        key_cols.append(key_col)
        out[key_col] = out[col].map(normalize_dedupe_value)
    out = out.drop_duplicates(subset=key_cols, keep="first").copy()
    return out.drop(columns=key_cols)


def write_deduped_reviews(df: pd.DataFrame, output_path: Path) -> None:
    export_df = df.drop(columns=["评分_num"], errors="ignore").copy()
    temp_path = output_path.with_name(f"{output_path.stem}__deduped_tmp{output_path.suffix}")

    with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
        for model_id, group in export_df.groupby("机型ID", sort=False):
            sheet_name = safe_sheet_name(str(model_id))
            group.to_excel(writer, sheet_name=sheet_name, index=False)
        format_deduped_review_workbook(writer.book)

    try:
        temp_path.replace(output_path)
    except PermissionError:
        fallback_path = output_path.with_name(f"{output_path.stem}_deduped{output_path.suffix}")
        temp_path.replace(fallback_path)
        safe_print(
            f"WARNING: 原始 reviews 文件可能正被 Excel 打开，无法覆盖；已保存去重副本: {fallback_path}",
            stream=sys.stderr,
        )


def write_classification_template(
    df: pd.DataFrame,
    issue_rules: list[tuple[str, list[str]]],
    output_path: Path,
    source_path: Path,
    category_path: Path | None,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    template_df = df.copy()
    template_df["_sort_date"] = template_df["评论日期"].map(parse_review_date)
    template_df["评论日期"] = template_df["评论日期"].map(format_review_date)
    template_df = template_df.sort_values(
        by=["机型ID", "_sort_date"],
        ascending=[True, False],
        na_position="last",
        kind="stable",
    ).reset_index(drop=True)
    template_df.insert(0, "ReviewID", [f"R{i + 1:04d}" for i in range(len(template_df))])
    template_df["一级分类_请填写"] = ""
    template_df["二级分类_请填写"] = ""
    template_df["分类理由_可选"] = ""

    review_cols = [
        "ReviewID",
        "机型ID",
        "机型名称",
        "Channel",
        "来源站点",
        "评分",
        "评论日期",
        "标题",
        "评论内容",
        "问题摘要",
        "一级分类_请填写",
        "二级分类_请填写",
        "分类理由_可选",
    ]
    reviews_to_classify = template_df[review_cols].copy()

    taxonomy_rows = []
    for primary, secondaries in issue_rules:
        for secondary in secondaries:
            taxonomy_rows.append({"一级分类": primary, "二级分类": secondary})
    taxonomy_df = pd.DataFrame(taxonomy_rows)

    instructions = pd.DataFrame(
        [
            ["用途", "把本文件上传到 ChatGPT 网页版，让 ChatGPT 阅读完整 review 语义后填写分类。"],
            ["来源Review文件", source_path.name],
            ["Call log分类表", category_path.name if category_path else "内置兜底规则"],
            ["分类要求", "不要只按关键词匹配。请理解用户真实抱怨点，从 Call log 分类表中选择最合适的一级分类和二级分类。"],
            ["多问题处理", "如果一条评论包含多个问题，选择最能解释低评分的主要问题。"],
            ["无有效正文", "如果 review 没有可用抱怨内容，一级分类和二级分类都填 Other。"],
            ["输出要求", "请在“待分类Reviews”sheet 中填写“一级分类_请填写”和“二级分类_请填写”，不要改 ReviewID。"],
        ],
        columns=["项目", "说明"],
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        instructions.to_excel(writer, sheet_name="使用说明", index=False)
        reviews_to_classify.to_excel(writer, sheet_name="待分类Reviews", index=False)
        taxonomy_df.to_excel(writer, sheet_name="Call log分类表", index=False)
        format_classification_template_workbook(writer.book)


def format_classification_template_workbook(workbook) -> None:
    header_fill = PatternFill("solid", fgColor="00A651")
    header_font = Font(color="FFFFFF", bold=True)
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            set_cell_alignment(cell, vertical="center", wrap_text=True)
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                set_cell_alignment(cell, vertical="center", wrap_text=True)
        for column in range(1, sheet.max_column + 1):
            letter = get_column_letter(column)
            sheet.column_dimensions[letter].width = column_width(sheet, column)


def format_deduped_review_workbook(workbook) -> None:
    header_fill = PatternFill("solid", fgColor="00A651")
    header_font = Font(color="FFFFFF", bold=True)

    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            set_cell_alignment(cell, vertical="center", wrap_text=True)

        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                set_cell_alignment(cell, vertical="center", wrap_text=True)

        for column in range(1, sheet.max_column + 1):
            letter = get_column_letter(column)
            sheet.column_dimensions[letter].width = column_width(sheet, column)


def normalize_dedupe_value(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def parse_rating(value) -> float | None:
    if pd.isna(value):
        return None
    match = re.search(r"\d+(?:\.\d+)?", str(value))
    return float(match.group(0)) if match else None


def parse_review_date(value) -> pd.Timestamp | pd.NaT:
    if pd.isna(value):
        return pd.NaT
    text = str(value).replace("\xa0", " ").strip()
    match = re.search(r"on\s+(.+)$", text, re.I)
    if match:
        text = match.group(1).strip()
    parsed = pd.to_datetime(text, errors="coerce")
    return parsed if not pd.isna(parsed) else pd.NaT


def format_review_date(value) -> str:
    parsed = parse_review_date(value)
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%Y-%m-%d")


def classify_reviews_semantic(
    df: pd.DataFrame,
    issue_rules: list[tuple[str, list[str]]],
) -> list[tuple[str, str]]:
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError(
            "语义分类需要 OPENAI_API_KEY。请在 .env 中配置可用 key，"
            "或手动添加 --classification-mode keyword 使用关键词备用模式。"
        )

    model = os.getenv("REVIEW_ANALYSIS_CLASSIFY_MODEL", "gpt-4o-mini").strip()
    taxonomy = {category: phrases for category, phrases in issue_rules}
    results: list[tuple[str, str]] = []
    for batch in chunked(review_rows_for_semantic_classification(df), 12):
        results.extend(
            classify_batch_semantic(
                batch,
                taxonomy=taxonomy,
                api_key=api_key,
                model=model,
            )
        )
    return results[: len(df)]


def review_rows_for_semantic_classification(df: pd.DataFrame) -> list[dict]:
    rows = []
    for item_id, (_, row) in enumerate(df.iterrows()):
        rows.append(
            {
                "id": item_id,
                "model": str(row.get("机型名称", "")),
                "channel": str(row.get("Channel", "")),
                "rating": str(row.get("评分", "")),
                "date": str(row.get("评论日期", "")),
                "title": truncate_text(row.get("标题", ""), 300),
                "review": truncate_text(row.get("评论内容", ""), 1600),
                "summary": truncate_text(row.get("问题摘要", ""), 500),
            }
        )
    return rows


def classify_batch_semantic(
    rows: list[dict],
    *,
    taxonomy: dict[str, list[str]],
    api_key: str,
    model: str,
) -> list[tuple[str, str]]:
    payload = {
        "model": model,
        "messages": [
            {
                "role": "system",
                "content": (
                    "You classify low-rating TV product reviews into a fixed Call Log taxonomy. "
                    "Read the full review meaning and the user's actual complaint. Do not classify by simple keyword matching. "
                    "Choose the single best primary category and secondary category. "
                    "The primary must be exactly one taxonomy key. The secondary must be exactly one item under that primary. "
                    "If the review has no usable complaint text or no category fits, return primary 'Other' and secondary 'Other'. "
                    "If multiple issues appear, choose the main issue that best explains the negative rating. "
                    "Return only JSON array: [{\"id\":0,\"primary\":\"...\",\"secondary\":\"...\"}]."
                ),
            },
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "taxonomy": taxonomy,
                        "reviews": rows,
                    },
                    ensure_ascii=False,
                ),
            },
        ],
        "temperature": 0,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    with httpx.Client(timeout=90) as client:
        response = client.post(
            "https://api.openai.com/v1/chat/completions",
            headers=headers,
            json=payload,
        )
    if response.status_code == 429:
        raise RuntimeError("OpenAI API 额度不足或账单不可用，无法进行语义分类。")
    response.raise_for_status()

    content = response.json()["choices"][0]["message"]["content"].strip()
    parsed = json.loads(strip_code_fence(content))
    by_id = {
        int(item.get("id")): validate_semantic_classification(item, taxonomy)
        for item in parsed
        if isinstance(item, dict) and "id" in item
    }
    return [by_id.get(row["id"], ("Other", "Other")) for row in rows]


def validate_semantic_classification(
    item: dict,
    taxonomy: dict[str, list[str]],
) -> tuple[str, str]:
    primary = str(item.get("primary", "")).strip()
    secondary = str(item.get("secondary", "")).strip()
    if primary not in taxonomy:
        return "Other", "Other"
    if secondary not in taxonomy[primary]:
        return primary, "Other"
    return primary, secondary


def truncate_text(value, limit: int) -> str:
    if pd.isna(value):
        return ""
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


def classify_review_keyword(row: pd.Series, issue_rules: list[tuple[str, list[str]]]) -> tuple[str, str]:
    text = normalize_match_text(f"{row.get('标题', '')} {row.get('评论内容', '')}")
    scores = []
    for category, phrases in issue_rules:
        score, matched_phrase = score_category(text, phrases)
        if score:
            scores.append((score, category, matched_phrase))
    if not scores:
        return "Other", "Other"
    scores.sort(key=lambda item: (-item[0], item[1]))
    return scores[0][1], scores[0][2]


def score_category(text: str, phrases: list[str]) -> tuple[int, str]:
    best_score = 0
    best_phrase = ""
    for phrase in phrases:
        normalized = normalize_match_text(phrase)
        if not normalized:
            continue
        if normalized in text:
            score = 8 + len(normalized.split())
            if score > best_score:
                best_score = score
                best_phrase = str(phrase).strip()
            continue

        tokens = meaningful_tokens(normalized)
        if not tokens:
            continue
        matches = sum(1 for token in tokens if token in text)
        if matches >= max(1, min(2, len(tokens))):
            score = matches
            if score > best_score:
                best_score = score
                best_phrase = str(phrase).strip()
    return best_score, best_phrase


def normalize_match_text(value: str) -> str:
    text = str(value).lower()
    text = text.replace("wi-fi", "wifi")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def meaningful_tokens(value: str) -> list[str]:
    return [
        token
        for token in value.split()
        if len(token) >= 3 and token not in STOPWORDS
    ]


def make_excerpt(value, limit: int = 220) -> str:
    if pd.isna(value):
        return ""
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


def make_issue_summary(row: pd.Series) -> str:
    body = make_excerpt(row.get("评论内容", ""))
    if body:
        return body

    title = make_excerpt(row.get("标题", ""))
    if title and not re.fullmatch(r"\d+(?:\.\d+)?\s+out of 5 stars", title, flags=re.I):
        return title

    return "原始评论无正文"


def translate_summaries(values: list[str]) -> list[str]:
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        safe_print(
            "WARNING: OPENAI_API_KEY 未配置，问题摘要中文列将留空。可在 .env 中添加 OPENAI_API_KEY。",
            stream=sys.stderr,
        )
        return [""] * len(values)

    model = os.getenv("REVIEW_ANALYSIS_TRANSLATE_MODEL", "gpt-4o-mini").strip()
    translations: list[str] = []
    for batch in chunked(values, 20):
        translations.extend(translate_batch(batch, api_key=api_key, model=model))
    return translations[: len(values)]


def translate_batch(values: list[str], *, api_key: str, model: str) -> list[str]:
    if not values:
        return []

    payload = {
        "model": model,
        "messages": [
            {
                "role": "system",
                "content": (
                    "Translate product review issue summaries into concise Simplified Chinese. "
                    "Preserve product terms, app names, model names, and technical abbreviations. "
                    "Return only a JSON array of strings, with the same length and order as input."
                ),
            },
            {
                "role": "user",
                "content": json.dumps(values, ensure_ascii=False),
            },
        ],
        "temperature": 0,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    try:
        with httpx.Client(timeout=60) as client:
            response = client.post(
                "https://api.openai.com/v1/chat/completions",
                headers=headers,
                json=payload,
            )
            if response.status_code == 429:
                safe_print(
                    "WARNING: OpenAI API 额度不足或账单不可用，问题摘要中文列将留空。",
                    stream=sys.stderr,
                )
                return [""] * len(values)
            response.raise_for_status()
        content = response.json()["choices"][0]["message"]["content"].strip()
        translated = json.loads(strip_code_fence(content))
        if isinstance(translated, list) and len(translated) == len(values):
            return [str(item) for item in translated]
    except Exception as exc:  # noqa: BLE001
        safe_print(f"WARNING: translation failed for one batch: {exc}", stream=sys.stderr)

    return [""] * len(values)


def strip_code_fence(value: str) -> str:
    text = value.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.I)
        text = re.sub(r"\s*```$", "", text)
    return text.strip()


def chunked(values: list[str], size: int):
    for start in range(0, len(values), size):
        yield values[start : start + size]


def write_analysis(
    df: pd.DataFrame,
    output_path: Path,
    source_path: Path,
    dedupe_stats: dict[str, int],
    issue_rules: list[tuple[str, list[str]]],
    category_path: Path | None,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        if df.empty:
            empty = pd.DataFrame(
                [["未发现 3 星及以下评论", str(source_path)]],
                columns=["说明", "来源文件"],
            )
            empty.to_excel(writer, sheet_name="无低分评论", index=False)
            return

        overview_rows = []
        for model_id, model_df in df.groupby("机型ID", sort=False):
            total = len(model_df)
            for category, count in Counter(model_df["问题分类"]).most_common():
                overview_rows.append(
                    {
                        "机型ID": model_id,
                        "问题分类": category,
                        "问题数": count,
                        "问题占比": count / total if total else 0,
                        "低分评论总数": total,
                    }
                )
        overview = pd.DataFrame(overview_rows)
        overview_start_row = 5
        dedupe_summary = pd.DataFrame(
            [
                ["来源文件", source_path.name],
                ["分类表", category_path.name if category_path else "内置兜底规则"],
                ["原始评论数", dedupe_stats["original_count"]],
                ["去重后评论数", dedupe_stats["deduped_count"]],
                ["去掉重复评论数", dedupe_stats["removed_count"]],
            ],
            columns=["项目", "值"],
        )
        dedupe_summary.to_excel(writer, sheet_name="汇总", index=False, startrow=0)
        overview.to_excel(writer, sheet_name="汇总", index=False, startrow=overview_start_row)

        for model_id, model_df in df.groupby("机型ID", sort=False):
            write_model_sheet(writer, str(model_id), model_df, source_path, issue_rules)

        format_workbook(writer.book)


def write_model_sheet(
    writer: pd.ExcelWriter,
    model_id: str,
    model_df: pd.DataFrame,
    source_path: Path,
    issue_rules: list[tuple[str, list[str]]],
) -> None:
    total = len(model_df)
    counts = Counter(model_df["问题分类"])
    summary = pd.DataFrame(
        [
            {
                "问题分类": category,
                "问题数": count,
                "问题占比": count / total if total else 0,
                "代表关键词": representative_keywords(category, issue_rules),
            }
            for category, count in counts.most_common()
        ]
    )

    detail_cols = [
        "机型名称",
        "Channel",
        "评分",
        "评论日期",
        "问题分类",
        "问题摘要",
        "问题摘要中文",
    ]
    details = model_df.copy()
    details["_sort_date"] = details["评论日期"].map(parse_review_date)
    details["评论日期"] = details["评论日期"].map(format_review_date)
    details["问题分类"] = details["二级问题分类"]
    details = details.sort_values(
        by="_sort_date",
        ascending=False,
        na_position="last",
        kind="stable",
    )[detail_cols]

    sheet_name = safe_sheet_name(model_id)
    header = pd.DataFrame(
        [
            [f"{model_id} 低分评论问题分析"],
            [f"来源文件: {source_path.name}"],
            [f"3星及以下评论数: {total}"],
        ]
    )
    header.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=0)
    summary.to_excel(writer, sheet_name=sheet_name, index=False, startrow=4)
    details.to_excel(writer, sheet_name=sheet_name, index=False, startrow=7 + len(summary))


def representative_keywords(category: str, issue_rules: list[tuple[str, list[str]]]) -> str:
    for name, keywords in issue_rules:
        if name == category:
            return ", ".join(keywords[:8])
    return ""


def safe_sheet_name(name: str, max_len: int = 31) -> str:
    invalid = set(r"[]:*?/\\")
    cleaned = "".join(c if c not in invalid else "_" for c in name).strip() or "sheet"
    return cleaned[:max_len]


def format_workbook(workbook) -> None:
    hisense_green = "00A651"
    header_fill = PatternFill("solid", fgColor=hisense_green)
    header_font = Font(color="FFFFFF", bold=True)
    title_font = Font(size=14, bold=True, color=hisense_green)

    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A6" if sheet.title != "汇总" else "A2"
        if sheet.title != "汇总" and sheet.max_column > 1:
            sheet.merge_cells(
                start_row=1,
                start_column=1,
                end_row=1,
                end_column=sheet.max_column,
            )
            sheet.cell(row=1, column=1).alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        for row in sheet.iter_rows():
            for cell in row:
                set_cell_alignment(cell, vertical="center", wrap_text=True)
                if cell.row == 1 and cell.column == 1:
                    cell.font = title_font
                if cell.value in {
                    "问题分类",
                    "问题数",
                    "问题占比",
                    "代表关键词",
                    "机型ID",
                    "机型名称",
                    "Channel",
                    "评分",
                    "评论日期",
                    "标题",
                    "问题摘要",
                    "问题摘要中文",
                    "低分评论总数",
                }:
                    cell.fill = header_fill
                    cell.font = header_font

        if sheet.title != "汇总":
            sheet.cell(row=1, column=1).alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        for column in range(1, sheet.max_column + 1):
            letter = get_column_letter(column)
            if column == 1:
                sheet.column_dimensions[letter].width = 14 if sheet.title != "汇总" else 16
            else:
                sheet.column_dimensions[letter].width = column_width(sheet, column)

        for row in range(1, sheet.max_row + 1):
            sheet.row_dimensions[row].height = 24

        center_columns_by_header(sheet, {"问题分类", "问题数", "问题占比"})
        apply_column_number_formats(sheet)


def center_columns_by_header(sheet, headers_to_center: set[str]) -> None:
    for row in sheet.iter_rows():
        for header_cell in row:
            if header_cell.value not in headers_to_center:
                continue
            header_cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )
            for cell in sheet.iter_cols(
                min_col=header_cell.column,
                max_col=header_cell.column,
                min_row=header_cell.row + 1,
                max_row=sheet.max_row,
            ):
                for data_cell in cell:
                    set_cell_alignment(
                        data_cell,
                        horizontal="center",
                        vertical="center",
                        wrap_text=True,
                    )


def set_cell_alignment(cell, *, horizontal=None, vertical=None, wrap_text=None) -> None:
    alignment = copy(cell.alignment)
    if horizontal is not None:
        alignment.horizontal = horizontal
    if vertical is not None:
        alignment.vertical = vertical
    if wrap_text is not None:
        alignment.wrap_text = wrap_text
    cell.alignment = alignment


def apply_column_number_formats(sheet) -> None:
    for row in sheet.iter_rows():
        headers = {cell.column: cell.value for cell in row if isinstance(cell.value, str)}
        if not headers:
            continue

        data_start = row[0].row + 1
        data_end = table_data_end_row(sheet, data_start)
        for column, header in headers.items():
            if "百分比" in header or "占比" in header or "比例" in header:
                for data_cell in sheet.iter_cols(
                    min_col=column,
                    max_col=column,
                    min_row=data_start,
                    max_row=data_end,
                ):
                    for cell in data_cell:
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = "0.0%"
            elif header in {"问题数", "低分评论总数"}:
                for data_cell in sheet.iter_cols(
                    min_col=column,
                    max_col=column,
                    min_row=data_start,
                    max_row=data_end,
                ):
                    for cell in data_cell:
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = "0"
            elif header == "评分":
                for data_cell in sheet.iter_cols(
                    min_col=column,
                    max_col=column,
                    min_row=data_start,
                    max_row=data_end,
                ):
                    for cell in data_cell:
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = "0.0"


def table_data_end_row(sheet, data_start: int) -> int:
    row_idx = data_start
    while row_idx <= sheet.max_row:
        values = [sheet.cell(row=row_idx, column=col).value for col in range(1, sheet.max_column + 1)]
        if all(value is None for value in values):
            return row_idx - 1
        row_idx += 1
    return sheet.max_row


def column_width(sheet, column: int) -> int:
    values = [sheet.cell(row=row, column=column).value for row in range(1, min(sheet.max_row, 80) + 1)]
    max_len = max((len(str(value)) for value in values if value is not None), default=8)
    return max(10, min(max_len + 2, 45))


def safe_print(value: str, *, stream=None) -> None:
    target = stream or sys.stdout
    try:
        print(value, file=target)
    except UnicodeEncodeError:
        encoded = value.encode(target.encoding or "utf-8", errors="backslashreplace")
        print(encoded.decode(target.encoding or "utf-8"), file=target)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        safe_print(f"ERROR: {exc}", stream=sys.stderr)
        raise
