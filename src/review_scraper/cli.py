from __future__ import annotations

from pathlib import Path

import typer
from rich.console import Console

from review_scraper.amazon_auth import ensure_amazon_session, resolve_state_path
from review_scraper.config import load_config, project_root
from review_scraper.pipeline import list_config, run_scrape

app = typer.Typer(
    name="review-scraper",
    help="爬取 BBY 等站点网评，按机型导出 Excel",
)
amazon_app = typer.Typer(help="Amazon 登录态（一次设置，长期复用）")
app.add_typer(amazon_app, name="amazon")
console = Console()


@app.command("run")
def cmd_run(
    config: Path = typer.Option(
        None,
        "--config",
        "-c",
        help="配置文件路径，默认 config/sites.yaml",
    ),
    model: str = typer.Option(None, "--model", "-m", help="只爬指定机型 model_id 或名称"),
    site: str = typer.Option(None, "--site", "-s", help="只爬指定站点，如 bestbuy"),
    output: str = typer.Option(None, "--output", "-o", help="输出 Excel 文件名"),
) -> None:
    """执行爬取并导出 Excel。"""
    run_scrape(
        config_path=config,
        model_filter=model,
        site_filter=site,
        output_filename=output,
    )


@app.command("list")
def cmd_list(
    config: Path = typer.Option(None, "--config", "-c", help="配置文件路径"),
) -> None:
    """查看已配置的机型与 URL（是否已填写）。"""
    path = config or (project_root() / "config" / "sites.yaml")
    console.print(f"配置文件: {path}\n")
    list_config(path)


@amazon_app.command("setup")
def amazon_setup(
    config: Path = typer.Option(None, "--config", "-c", help="配置文件路径"),
) -> None:
    """一次性手动登录并保存 Cookie（之后 run 可全自动）。"""
    defaults = load_config(config).defaults
    state = resolve_state_path(defaults.amazon_storage_state)
    path = ensure_amazon_session(
        state,
        auto_login=False,
        force_manual_setup=True,
    )
    console.print(f"[green]登录态已保存[/green]: {path}")
    console.print("以后直接运行: review-scraper run")


@amazon_app.command("check")
def amazon_check(
    config: Path = typer.Option(None, "--config", "-c", help="配置文件路径"),
) -> None:
    """检查当前登录态是否仍有效。"""
    from review_scraper.amazon_auth import is_session_valid

    defaults = load_config(config).defaults
    state = resolve_state_path(defaults.amazon_storage_state)
    if is_session_valid(state):
        console.print(f"[green]登录态有效[/green]: {state}")
    else:
        console.print(f"[yellow]登录态无效或不存在[/yellow]，请运行: review-scraper amazon setup")


if __name__ == "__main__":
    app()
