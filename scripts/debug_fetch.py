import httpx
from pathlib import Path

url = (
    "https://www.amazon.com/product-reviews/B0GR9NR9XV/"
    "?pageNumber=1&reviewerType=all_reviews&filterByStar=all_stars"
)
headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}
r = httpx.get(url, headers=headers, follow_redirects=True, timeout=45)
out = Path(__file__).resolve().parents[1] / "data" / "output" / "_debug.html"
out.write_text(r.text, encoding="utf-8")
text = r.text.lower()
print("status", r.status_code, "len", len(r.text))
for k in ["captcha", "robot", "data-hook=\"review\"", "motion-review"]:
    print(k, k.replace('"', "") in text)
