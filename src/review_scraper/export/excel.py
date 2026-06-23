from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re

import pandas as pd

from review_scraper.config import project_root
from review_scraper.models import ReviewRecord


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
        df = pd.DataFrame(
            columns=[
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
        )
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="无数据", index=False)
        return output_path

    df_all = pd.DataFrame([r.as_row() for r in records])
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
