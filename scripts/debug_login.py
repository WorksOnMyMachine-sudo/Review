import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from playwright.sync_api import sync_playwright
from review_scraper.amazon_auth import attempt_amazon_login, load_amazon_credentials
from review_scraper.config import project_root

STATE = project_root() / "data" / "amazon_state.json"
REVIEW = (
    "https://www.amazon.com/product-reviews/B0GR9NR9XV/ref=cm_cr_arp_d_viewopt_sr"
    "?ie=UTF8&filterByStar=all_stars&reviewerType=all_reviews&pageNumber=1"
)
OUT = project_root() / "data" / "output"

email, password = load_amazon_credentials()
OUT.mkdir(parents=True, exist_ok=True)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    ctx_kwargs = {"locale": "en-US"}
    if STATE.exists():
        ctx_kwargs["storage_state"] = str(STATE)
    ctx = browser.new_context(**ctx_kwargs)
    page = ctx.new_page()
    if not STATE.exists():
        attempt_amazon_login(page, email, password)
        page.wait_for_timeout(8000)
        ctx.storage_state(path=str(STATE))
    page.goto(REVIEW, wait_until="domcontentloaded")
    page.wait_for_timeout(5000)
    html = page.content()
    (OUT / "_after_login_reviews.html").write_text(html, encoding="utf-8")
    n = html.count("data-hook=\"review\"")
    print("review blocks:", n)
    print("sign-in:", "sign-in" in html.lower() or "ap_signin" in html.lower())
    print("title snippet:", page.title()[:80])
    browser.close()
