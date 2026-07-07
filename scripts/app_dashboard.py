from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "data" / "output"
REVIEWS_PATH = OUTPUT_DIR / "reviews_incremental.xlsx"
ANALYSIS_PATH = OUTPUT_DIR / "review_issue_analysis_latest.xlsx"


st.set_page_config(
    page_title="网评分析看板",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded",
)


st.markdown(
    """
    <style>
    .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
    h1, h2, h3 { letter-spacing: 0; }
    div[data-testid="stMetric"] {
        border: 1px solid #d9e2e5;
        border-radius: 6px;
        padding: 12px 14px;
        background: #fbfcfc;
    }
    div[data-testid="stDataFrame"] { border: 1px solid #e4eaec; }
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        font-size: 1rem;
        margin-top: 0.75rem;
    }
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
        color: #607078;
    }
    section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p {
        margin-bottom: 0.2rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_reviews(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()

    workbook = pd.ExcelFile(path)
    frames: list[pd.DataFrame] = []
    for sheet_name in workbook.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet_name)
        if df.empty or "机型ID" not in df.columns:
            continue
        df["_sheet"] = sheet_name
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    out = pd.concat(frames, ignore_index=True)
    out["评分_num"] = out.get("评分", pd.Series(dtype=object)).map(parse_rating)
    out["评论日期_dt"] = pd.to_datetime(
        out.get("评论日期", pd.Series(dtype=object)).map(clean_review_date),
        errors="coerce",
    )
    out["抓取时间_dt"] = pd.to_datetime(out.get("抓取时间", pd.Series(dtype=object)), errors="coerce")
    out["周"] = format_week_label(out["评论日期_dt"])
    for column in ["机型ID", "机型名称", "尺寸", "Channel", "来源站点", "标题", "评论内容", "用户"]:
        if column not in out.columns:
            out[column] = ""
        out[column] = out[column].fillna("").astype(str)
    return out


@st.cache_data(show_spinner=False)
def load_analysis(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    if not path.exists():
        return pd.DataFrame(), pd.DataFrame()

    workbook = pd.ExcelFile(path)
    summaries: list[pd.DataFrame] = []
    details: list[pd.DataFrame] = []

    for sheet_name in workbook.sheet_names:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
        summary = extract_table(raw, required_headers={"问题分类", "问题数", "问题占比"})
        if not summary.empty:
            summary.insert(0, "机型ID", sheet_name)
            summaries.append(summary)

        detail = extract_table(raw, required_headers={"机型名称", "Channel", "评分", "评论内容"})
        if not detail.empty:
            detail.insert(0, "机型ID", sheet_name)
            detail["评分_num"] = detail.get("评分", pd.Series(dtype=object)).map(parse_rating)
            detail["评论日期_dt"] = pd.to_datetime(
                detail.get("评论日期", pd.Series(dtype=object)).map(clean_review_date),
                errors="coerce",
            )
            details.append(detail)

    summary_df = pd.concat(summaries, ignore_index=True) if summaries else pd.DataFrame()
    detail_df = pd.concat(details, ignore_index=True) if details else pd.DataFrame()
    return summary_df, detail_df


def extract_table(raw: pd.DataFrame, *, required_headers: set[str]) -> pd.DataFrame:
    for row_idx in range(len(raw)):
        values = [str(value).strip() for value in raw.iloc[row_idx].tolist() if pd.notna(value)]
        if required_headers.issubset(set(values)):
            headers = [str(value).strip() if pd.notna(value) else "" for value in raw.iloc[row_idx].tolist()]
            end_idx = row_idx + 1
            while end_idx < len(raw):
                row_values = raw.iloc[end_idx].tolist()
                if all(pd.isna(value) for value in row_values):
                    break
                if any(str(value).strip() in {"渠道评分与评论数", "低分问题分类汇总"} for value in row_values if pd.notna(value)):
                    break
                end_idx += 1
            table = raw.iloc[row_idx + 1 : end_idx].copy()
            table.columns = headers
            table = table[[column for column in table.columns if column]]
            table = table.dropna(how="all")
            return table.reset_index(drop=True)
    return pd.DataFrame()


def parse_rating(value) -> float | None:
    if pd.isna(value):
        return None
    text = str(value)
    match = pd.Series([text]).str.extract(r"(\d+(?:\.\d+)?)", expand=False).iloc[0]
    if pd.isna(match):
        return None
    value_float = float(match)
    return value_float if 0 <= value_float <= 5 else None


def clean_review_date(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("\xa0", " ").strip()
    if not text:
        return ""
    marker = "on "
    lower = text.lower()
    if marker in lower:
        return text[lower.rfind(marker) + len(marker) :].strip()
    return text


def format_week_label(values: pd.Series) -> pd.Series:
    dates = pd.to_datetime(values, errors="coerce")
    iso = dates.dt.isocalendar()
    labels = iso["year"].astype(str) + "-W" + iso["week"].astype(str).str.zfill(2)
    return labels.where(dates.notna(), "")


def rating_label(value: float) -> str:
    if pd.isna(value):
        return "无评分"
    if float(value).is_integer():
        return f"{int(value)} star"
    return f"{value:g} star"


def apply_review_filters(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    with st.sidebar:
        st.header("筛选器")
        st.caption("选择后，概览、问题分析、明细和导出都会同步变化。")

        all_total = len(out)
        all_models = out["机型ID"].nunique() if "机型ID" in out else 0
        all_channels = out["Channel"].nunique() if "Channel" in out else 0
        st.markdown(f"**当前数据**  {all_total:,} 条 / {all_models} 个机型 / {all_channels} 个渠道")

        st.divider()
        st.subheader("产品")
        models = sorted(out["机型ID"].dropna().astype(str).unique().tolist()) if "机型ID" in out else []
        selected_models = st.multiselect(
            "机型",
            models,
            default=models,
            placeholder="选择机型",
            help="默认选择全部机型。",
        )
        if selected_models:
            out = out[out["机型ID"].isin(selected_models)]

        st.subheader("渠道")
        channels = sorted(out["Channel"].dropna().astype(str).unique().tolist()) if "Channel" in out else []
        selected_channels = st.multiselect(
            "渠道",
            channels,
            default=channels,
            placeholder="选择渠道",
            help="默认选择全部渠道。",
        )
        if selected_channels:
            out = out[out["Channel"].isin(selected_channels)]

        st.subheader("评分")
        rating_values = sorted(out["评分_num"].dropna().unique().tolist()) if "评分_num" in out else []
        selected_ratings = st.multiselect(
            "评分",
            rating_values,
            default=rating_values,
            format_func=rating_label,
            placeholder="选择评分",
        )
        if selected_ratings:
            out = out[out["评分_num"].isin(selected_ratings)]

        st.subheader("时间")
        dates = out["评论日期_dt"].dropna() if "评论日期_dt" in out else pd.Series(dtype="datetime64[ns]")
        if not dates.empty:
            min_date = dates.min().date()
            max_date = dates.max().date()
            date_range = st.date_input("评论日期", value=(min_date, max_date), min_value=min_date, max_value=max_date)
            if isinstance(date_range, tuple) and len(date_range) == 2:
                start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
                out = out[out["评论日期_dt"].between(start, end + pd.Timedelta(days=1), inclusive="left")]

        st.divider()
        st.caption(f"已筛选: {len(out):,} 条")
        st.caption(f"原始文件: {REVIEWS_PATH.name}")
        st.caption(f"分析文件: {ANALYSIS_PATH.name}")
        if st.button("刷新数据", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    return out


def metric_row(df: pd.DataFrame, analysis_detail: pd.DataFrame) -> None:
    total = len(df)
    low = int(df["评分_num"].le(3).sum()) if "评分_num" in df else 0
    models = df["机型ID"].nunique() if "机型ID" in df else 0
    channels = df["Channel"].nunique() if "Channel" in df else 0
    analyzed = len(analysis_detail) if not analysis_detail.empty else 0

    cols = st.columns(5)
    cols[0].metric("评论总数", f"{total:,}")
    cols[1].metric("低分评论", f"{low:,}")
    cols[2].metric("机型数", f"{models:,}")
    cols[3].metric("渠道数", f"{channels:,}")
    cols[4].metric("已分类低分", f"{analyzed:,}")


def download_excel_button(label: str, df: pd.DataFrame, filename: str) -> None:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    st.download_button(
        label,
        data=buffer.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def main() -> None:
    st.title("网评分析看板")

    reviews = load_reviews(REVIEWS_PATH)
    summary, detail = load_analysis(ANALYSIS_PATH)

    if reviews.empty:
        st.warning("没有找到累计原始数据。请先运行 python scripts/run_full_workflow.py")
        return

    filtered = apply_review_filters(reviews)
    filtered_detail = detail
    if not detail.empty and "机型ID" in filtered:
        filtered_detail = detail[detail["机型ID"].isin(filtered["机型ID"].unique())]

    metric_row(filtered, filtered_detail)

    tab_overview, tab_issues, tab_reviews, tab_exports = st.tabs(["概览", "问题分析", "评论明细", "导出"])

    with tab_overview:
        left, right = st.columns(2)
        with left:
            st.subheader("各机型评论数")
            model_counts = filtered.groupby("机型ID").size().sort_values(ascending=False)
            st.bar_chart(model_counts)
        with right:
            st.subheader("渠道评论数")
            channel_counts = filtered.groupby("Channel").size().sort_values(ascending=False)
            st.bar_chart(channel_counts)

        trend_source = filtered[filtered["周"].astype(bool)].copy()
        model_trend = pd.pivot_table(
            trend_source,
            values="评论内容",
            index="周",
            columns="机型ID",
            aggfunc="count",
            fill_value=0,
        ).sort_index()
        st.subheader("每周评论趋势（按机型）")
        st.line_chart(model_trend)

        st.subheader("单机型趋势")
        trend_models = list(model_trend.columns)
        selected_trend_models = st.multiselect(
            "选择要单独查看的机型",
            trend_models,
            default=trend_models,
        )
        if selected_trend_models:
            trend_columns = st.columns(2)
            for index, model_id in enumerate(selected_trend_models):
                with trend_columns[index % 2]:
                    st.caption(str(model_id))
                    st.line_chart(model_trend[[model_id]])

        st.subheader("各机型问题分类占比")
        shown_summary = summary[summary["机型ID"].isin(filtered["机型ID"].unique())] if not summary.empty else pd.DataFrame()
        required_issue_columns = {"机型ID", "问题分类", "问题数"}
        if shown_summary.empty or not required_issue_columns.issubset(shown_summary.columns):
            st.info("还没有可展示的问题分类占比。运行分析后会在这里显示。")
        else:
            issue_summary = shown_summary.copy()
            issue_summary["问题数"] = pd.to_numeric(issue_summary["问题数"], errors="coerce").fillna(0)
            issue_tables = []
            for model_id, model_issues in issue_summary.groupby("机型ID", sort=True):
                model_table = (
                    model_issues.groupby("问题分类", as_index=False)["问题数"]
                    .sum()
                    .sort_values("问题数", ascending=False)
                )
                total_issues = model_table["问题数"].sum()
                model_table["问题占比"] = (
                    model_table["问题数"].div(total_issues).fillna(0).map(lambda value: f"{value:.1%}")
                    if total_issues
                    else "0.0%"
                )
                issue_tables.append((model_id, model_table))

            for row_start in range(0, len(issue_tables), 2):
                columns = st.columns(2)
                for column, (model_id, model_table) in zip(columns, issue_tables[row_start : row_start + 2]):
                    with column:
                        st.caption(str(model_id))
                        st.dataframe(model_table, use_container_width=True, hide_index=True)

    with tab_issues:
        if summary.empty and filtered_detail.empty:
            st.info("还没有最终分析数据。运行 python scripts/analyze_reviews.py 后会显示问题分类。")
        else:
            if not summary.empty:
                shown_summary = summary[summary["机型ID"].isin(filtered["机型ID"].unique())]
                issue_col = "问题分类" if "问题分类" in shown_summary.columns else None
                count_col = "问题数" if "问题数" in shown_summary.columns else None
                if issue_col and count_col:
                    issue_counts = (
                        shown_summary.groupby(issue_col)[count_col]
                        .sum()
                        .sort_values(ascending=False)
                    )
                    st.subheader("问题分类 Top")
                    st.bar_chart(issue_counts)
                st.subheader("问题分类汇总")
                st.dataframe(shown_summary, use_container_width=True, hide_index=True)

            if not filtered_detail.empty:
                st.subheader("低分评论分类明细")
                display_cols = [
                    col
                    for col in ["机型ID", "机型名称", "Channel", "评分", "评论日期", "问题分类", "分类理由", "评论内容", "问题摘要中文"]
                    if col in filtered_detail.columns
                ]
                st.dataframe(filtered_detail[display_cols], use_container_width=True, hide_index=True)

    with tab_reviews:
        st.subheader("原始评论明细")
        display_cols = [
            col
            for col in ["机型ID", "机型名称", "尺寸", "Channel", "评分", "整体评分", "评论日期", "标题", "评论内容", "用户", "来源URL"]
            if col in filtered.columns
        ]
        st.dataframe(filtered[display_cols], use_container_width=True, hide_index=True)

    with tab_exports:
        st.subheader("导出当前筛选结果")
        col1, col2 = st.columns(2)
        with col1:
            download_excel_button("导出原始评论", filtered, "filtered_reviews.xlsx")
        with col2:
            if filtered_detail.empty:
                st.button("导出分析明细", disabled=True)
            else:
                download_excel_button("导出分析明细", filtered_detail, "filtered_analysis.xlsx")


if __name__ == "__main__":
    main()
