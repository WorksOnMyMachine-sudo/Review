from __future__ import annotations

from pathlib import Path

from rich.console import Console

from review_scraper.amazon_auth import ensure_amazon_session, resolve_state_path
from review_scraper.config import AppConfig, iter_scrape_targets, load_config
from review_scraper.config import project_root
from review_scraper.export.excel import export_reviews_to_excel, merge_reviews_to_excel
from review_scraper.models import ReviewRecord
from review_scraper.scrapers.registry import get_scraper

console = Console()


def run_scrape(
    *,
    config_path: Path | None = None,
    model_filter: str | None = None,
    site_filter: str | None = None,
    output_filename: str | None = None,
    incremental: bool = False,
    incremental_path: Path | None = None,
) -> Path:
    config = load_config(config_path)
    targets = iter_scrape_targets(config)

    if model_filter:
        targets = [t for t in targets if t[0].model_id == model_filter or t[0].display_name == model_filter]
    if site_filter:
        site_key = site_filter.lower()
        targets = [
            t
            for t in targets
            if t[1].lower() == site_key
            or (site_key == "amazon" and t[1].lower().startswith(("amazon", "amz_")))
            or (site_key in {"bestbuy", "bby"} and t[1].lower().startswith(("bestbuy", "bby_")))
            or (site_key in {"walmart", "wmt"} and t[1].lower().startswith(("walmart", "wmt_")))
        ]

    if not targets:
        console.print(
            "[yellow]没有可爬取的目标。[/yellow]\n"
            "请在 config/sites.yaml 中为机型填写 url，并设置 enabled: true"
        )
        return export_reviews_to_excel([], filename=output_filename or "reviews_empty.xlsx")

    if any(site.lower().startswith(("amazon", "amz_")) for _, site, _ in targets):
        state = resolve_state_path(config.defaults.amazon_storage_state)
        ensure_amazon_session(
            state,
            auto_login=config.defaults.amazon_auto_login,
            headless=config.defaults.amazon_playwright_headless,
            force_manual_setup=False,
        )

    if incremental:
        output_path = incremental_path or (project_root() / "data" / "output" / "reviews_incremental.xlsx")
        if not output_path.is_absolute():
            output_path = project_root() / output_path
        console.print(f"[green]递进模式[/green]: 新评论会合并到 {output_path}")
    else:
        output_path = Path()

    all_records: list[ReviewRecord] = []
    for model, site_name, site_target in targets:
        console.print(
            f"[cyan]爬取[/cyan] {model.display_name} @ {site_name}\n  {site_target.url}"
        )
        scraper = get_scraper(site_name, config.defaults)
        try:
            records = _scrape_target(scraper, model, site_target)
            console.print(f"  → 获取 {len(records)} 条评论")
            if incremental:
                _, added, total = merge_reviews_to_excel(records, output_path=output_path)
                console.print(f"  → 新增 {added} 条，累计 {total} 条（已保存）")
            else:
                all_records.extend(records)
        except Exception as exc:  # noqa: BLE001
            if _is_amazon_site(site_name) and _looks_like_amazon_auth_error(exc):
                console.print(
                    "  [yellow]Amazon 登录态可能已过期或需要验证。[/yellow]\n"
                    "  将打开浏览器，请完成登录/验证后回到终端按回车保存，然后自动重试当前目标。"
                )
                try:
                    state = resolve_state_path(config.defaults.amazon_storage_state)
                    ensure_amazon_session(
                        state,
                        auto_login=False,
                        headless=False,
                        force_manual_setup=True,
                    )
                    scraper = get_scraper(site_name, config.defaults)
                    records = _scrape_target(scraper, model, site_target)
                    console.print(f"  → 重新登录后获取 {len(records)} 条评论")
                    if incremental:
                        _, added, total = merge_reviews_to_excel(records, output_path=output_path)
                        console.print(f"  → 新增 {added} 条，累计 {total} 条（已保存）")
                    else:
                        all_records.extend(records)
                    continue
                except Exception as retry_exc:  # noqa: BLE001
                    console.print(f"  [red]重新登录后仍失败[/red]: {retry_exc}")
                    continue

            console.print(f"  [red]失败[/red]: {exc}")

    if incremental:
        console.print(f"\n[green]递进更新完成[/green]: {output_path}")
        return output_path

    output = export_reviews_to_excel(all_records, filename=output_filename)
    console.print(f"\n[green]已导出 Excel[/green]: {output}")
    console.print(f"共 {len(all_records)} 条评论，按机型分 sheet 存储")
    return output


def _scrape_target(scraper, model, site_target) -> list[ReviewRecord]:
    return scraper.scrape(
        url=site_target.url.strip(),
        model_id=model.model_id,
        model_name=model.display_name,
        max_pages=site_target.max_pages,
    )


def _is_amazon_site(site_name: str) -> bool:
    return site_name.lower().startswith(("amazon", "amz_"))


def _looks_like_amazon_auth_error(exc: Exception) -> bool:
    text = str(exc).lower()
    patterns = [
        "登录态可能已过期",
        "登录",
        "验证",
        "captcha",
        "sign-in",
        "signin",
        "robot check",
        "not a robot",
        "ap/signin",
    ]
    return any(pattern in text for pattern in patterns)


def list_config(config_path: Path | None = None) -> None:
    config = load_config(config_path)
    for model in config.models:
        console.print(f"\n[bold]{model.display_name}[/bold] ({model.model_id})")
        for site_name, site in model.sites.items():
            status = "✓" if site.enabled and site.url.strip() else "○"
            url_preview = site.url.strip() or "（未填写 URL）"
            console.print(f"  {status} {site_name}: {url_preview}")
