"""Amazon login helpers — credentials only from environment / .env (never hardcoded)."""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import Page, sync_playwright

from review_scraper.config import project_root

STATE_DEFAULT = project_root() / "data" / "amazon_state.json"
SIGNIN_URL = (
    "https://www.amazon.com/ap/signin?openid.pape.max_auth_age=0"
    "&openid.return_to=https%3A%2F%2Fwww.amazon.com%2F"
    "&openid.assoc_handle=usflex&openid.mode=checkid_setup"
)
HOME_URL = "https://www.amazon.com"


def resolve_state_path(raw: str | None) -> Path:
    if not raw:
        return STATE_DEFAULT
    path = Path(raw)
    return path if path.is_absolute() else project_root() / path


def load_amazon_credentials() -> tuple[str, str]:
    load_dotenv(project_root() / ".env")
    email = os.getenv("AMAZON_EMAIL", "").strip()
    password = os.getenv("AMAZON_PASSWORD", "").strip()
    if not email or not password:
        raise ValueError(
            "请在项目根目录 .env 中设置 AMAZON_EMAIL 与 AMAZON_PASSWORD（勿提交到 Git）"
        )
    return email, password


def _fill_if_visible(page: Page, selector: str, value: str) -> bool:
    loc = page.locator(selector).first
    if loc.count() and loc.is_visible():
        loc.fill(value)
        return True
    return False


def attempt_amazon_login(page: Page, email: str, password: str) -> None:
    """Try common Amazon sign-in form steps; may still require OTP / CAPTCHA."""
    page.goto(SIGNIN_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(1500)

    for sel in ("#ap_email", "input[name='email']", "input[type='email']"):
        if _fill_if_visible(page, sel, email):
            break
    for sel in ("#continue", "input#continue", "button:has-text('Continue')"):
        btn = page.locator(sel).first
        if btn.count() and btn.is_visible():
            btn.click()
            page.wait_for_timeout(2000)
            break

    for sel in ("#ap_password", "input[name='password']", "input[type='password']"):
        if _fill_if_visible(page, sel, password):
            break
    for sel in ("#signInSubmit", "input#signInSubmit", "button:has-text('Sign in')"):
        btn = page.locator(sel).first
        if btn.count() and btn.is_visible():
            btn.click()
            page.wait_for_timeout(3000)
            break


def is_session_valid(state_path: Path, *, headless: bool = True) -> bool:
    """Check saved cookies still represent a logged-in Amazon session."""
    if not state_path.exists():
        return False
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=headless)
            context = browser.new_context(storage_state=str(state_path), locale="en-US")
            page = context.new_page()
            page.set_default_timeout(30_000)
            page.goto(HOME_URL, wait_until="domcontentloaded")
            page.wait_for_timeout(2000)
            ok = _page_looks_logged_in(page)
            browser.close()
            return ok
    except Exception:  # noqa: BLE001
        return False


def _page_looks_logged_in(page: Page) -> bool:
    if "ap/signin" in page.url.lower():
        return False
    try:
        label = page.locator("#nav-link-accountList-nav-line-1").inner_text(timeout=5000)
        lower = label.lower()
        if "sign in" in lower:
            return False
        if lower.strip() in {"hello", ""}:
            return False
        return True
    except Exception:  # noqa: BLE001
        content = page.content().lower()
        return "nav-link-accountList" in content and "sign in" not in content[:50000]


def save_storage_state(
    *,
    state_path: Path | None = None,
    headless: bool = False,
    manual_fallback: bool = True,
    auto_fill_credentials: bool = True,
) -> Path:
    """
    Log in with .env credentials, save Playwright storage state.
    If OTP/CAPTCHA appears, complete it in the browser window when headless=False.
    """
    out = state_path or STATE_DEFAULT
    out.parent.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=headless,
            args=["--disable-blink-features=AutomationControlled"],
        )
        context = browser.new_context(locale="en-US", viewport={"width": 1280, "height": 900})
        page = context.new_page()
        page.set_default_timeout(60_000)

        if auto_fill_credentials:
            email, password = load_amazon_credentials()
            attempt_amazon_login(page, email, password)
        else:
            page.goto(HOME_URL, wait_until="domcontentloaded")

        if manual_fallback and not headless:
            print(
                "请在打开的浏览器中手动登录 Amazon。\n"
                "若出现验证码 / 二次验证，也请在浏览器中手动完成。\n"
                "确认已登录 Amazon 后，回到终端按【回车】保存登录态..."
            )
            input()
        else:
            page.wait_for_timeout(12_000)

        context.storage_state(path=str(out))
        browser.close()

    return out


def ensure_amazon_session(
    state_path: Path,
    *,
    auto_login: bool = True,
    headless: bool = True,
    force_manual_setup: bool = False,
) -> Path:
    """
    Ensure valid Amazon session file exists.

    - Valid existing file → reuse (no browser login).
    - Missing/expired + auto_login → try .env headless login.
    - Still invalid + force_manual_setup → visible browser + user confirms.
    """
    state_path.parent.mkdir(parents=True, exist_ok=True)

    if not force_manual_setup and is_session_valid(state_path, headless=headless):
        return state_path

    if auto_login and not force_manual_setup:
        print("Amazon 登录态缺失或已过期，正在用 .env 账号自动登录…")
        save_storage_state(
            state_path=state_path,
            headless=headless,
            manual_fallback=False,
            auto_fill_credentials=True,
        )
        if is_session_valid(state_path, headless=headless):
            print("自动登录成功，已更新登录态文件。")
            return state_path
        raise RuntimeError(
            "无头自动登录未通过（Amazon 可能要求验证码）。\n"
            "请在本机执行一次（只需几分钟，之后可长期自动爬取）：\n"
            "  review-scraper amazon setup\n"
            "或: python scripts/save_amazon_state.py"
        )

    print("将打开浏览器，请完成登录/验证后按回车保存（通常只需做一次）。")
    save_storage_state(
        state_path=state_path,
        headless=False,
        manual_fallback=True,
        auto_fill_credentials=False,
    )
    if not is_session_valid(state_path, headless=True):
        raise RuntimeError("登录态仍无效，请确认 Amazon 账号可正常登录后重试。")
    return state_path
