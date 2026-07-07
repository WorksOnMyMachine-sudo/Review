from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re

import pandas as pd

from review_scraper.config import project_root
from review_scraper.models import ReviewRecord


REVIEW_COLUMNS = [
    "机型ID",
    "机型名称",
    "尺寸",
    "Channel",
    "来源站点",
    "来源URL",
    "整体评分",
    "评分",
    "标题",
    "评论内容",
    "评论日期",
    "用户",
    "Verified",
    "抓取时间",
]

DEDUPE_COLUMNS = [
    "机型ID",
    "Channel",
    "来源站点",
    "标题",
    "评论内容",
    "评论日期",
    "用户",
]


def export_reviews_to_excel(
    records: list[ReviewRecord],
    *,
    output_dir: Path | None = None,
    filename: str | None = None,
) -> Path:
    out_dir = output_dir or (project_root() / "data" / "output")
    out_dir.mkdir(parents=True, exist_ok=True)

    if not filename:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"reviews_{stamp}.xlsx"

    output_path = out_dir / filename

    if not records:
        return export_reviews_dataframe_to_excel(pd.DataFrame(columns=REVIEW_COLUMNS), output_path)

    return export_reviews_dataframe_to_excel(records_to_dataframe(records), output_path)


def merge_reviews_to_excel(
    records: list[ReviewRecord],
    *,
    output_path: Path,
) -> tuple[Path, int, int]:
    existing = read_reviews_excel(output_path)
    incoming = records_to_dataframe(records)
    before = len(existing)
    combined = dedupe_review_dataframe(pd.concat([existing, incoming], ignore_index=True))
    export_reviews_dataframe_to_excel(combined, output_path)
    return output_path, len(combined) - before, len(combined)


def records_to_dataframe(records: list[ReviewRecord]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame(columns=REVIEW_COLUMNS)
    df = pd.DataFrame([r.as_row() for r in records])
    return normalize_review_dataframe(df)


def read_reviews_excel(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=REVIEW_COLUMNS)

    workbook = pd.ExcelFile(path)
    frames = []
    for sheet in workbook.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        if df.empty:
            continue
        if {"机型ID", "评论内容"}.issubset(df.columns):
            frames.append(df)
    if not frames:
        return pd.DataFrame(columns=REVIEW_COLUMNS)
    return normalize_review_dataframe(pd.concat(frames, ignore_index=True))


def normalize_review_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for column in REVIEW_COLUMNS:
        if column not in out.columns:
            out[column] = ""
    return out[REVIEW_COLUMNS]


def dedupe_review_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    out = normalize_review_dataframe(df)
    key_cols = []
    for column in DEDUPE_COLUMNS:
        key_col = f"_dedupe_{column}"
        key_cols.append(key_col)
        out[key_col] = out[column].map(_normalize_dedupe_value)
    out = out.drop_duplicates(subset=key_cols, keep="first").copy()
    return out.drop(columns=key_cols)


def export_reviews_dataframe_to_excel(df_all: pd.DataFrame, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df_all = normalize_review_dataframe(df_all)
    if df_all.empty:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df_all.to_excel(writer, sheet_name="无数据", index=False)
        return output_path

    df_all["_sort_date"] = df_all["评论日期"].map(_parse_review_date)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for model_id, group in df_all.groupby("机型ID", sort=False):
            sheet = _safe_sheet_name(str(model_id))
            sorted_group = group.sort_values(
                by="_sort_date",
                ascending=False,
                na_position="last",
                kind="stable",
            ).drop(columns=["_sort_date"])
            sorted_group.to_excel(writer, sheet_name=sheet, index=False)

    return output_path


def _normalize_dedupe_value(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def _parse_review_date(value: str) -> pd.Timestamp | pd.NaT:
    if not value:
        return pd.NaT

    text = str(value).replace("\xa0", " ").strip()
    if not text:
        return pd.NaT

    match = re.search(r"on\s+(.+)$", text, re.I)
    if match:
        text = match.group(1).strip()

    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return pd.NaT
    return parsed


def _safe_sheet_name(name: str, max_len: int = 31) -> str:
    invalid = set(r"[]:*?/\\")
    cleaned = "".join(c if c not in invalid else "_" for c in name).strip() or "sheet"
    return cleaned[:max_len]
