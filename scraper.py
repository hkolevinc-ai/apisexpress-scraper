import argparse
import datetime as dt
import html
import json
import logging
import re
import time
from dataclasses import dataclass, asdict
from typing import Iterable, Optional
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://apisexpress.com"
DEFAULT_CATEGORY_URL = f"{BASE_URL}/produkt-kategoriya/uchenicheski-materiali/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; ApisExpressCatalogBot/1.0; +https://apisexpress.com/)",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("apisexpress-scraper")

@dataclass
class Product:
    sku: str = ""
    product_name: str = ""
    category: str = ""
    price_eur: str = ""
    image_urls: str = ""
    description: str = ""
    product_url: str = ""

class Client:
    def __init__(self, delay: float = 0.7, timeout: int = 30, retries: int = 3):
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        self.delay = delay
        self.timeout = timeout
        self.retries = retries

    def get(self, url: str, **kwargs) -> requests.Response:
        last_exc = None
        for attempt in range(1, self.retries + 1):
            try:
                r = self.session.get(url, timeout=self.timeout, **kwargs)
                if r.status_code in (429, 500, 502, 503, 504):
                    raise requests.HTTPError(f"HTTP {r.status_code}")
                r.raise_for_status()
                time.sleep(self.delay)
                return r
            except Exception as exc:
                last_exc = exc
                wait = attempt * 2
                log.warning("Fetch failed (%s/%s) %s: %s", attempt, self.retries, url, exc)
                time.sleep(wait)
        raise RuntimeError(f"Failed to fetch {url}: {last_exc}")

def clean_text(value: str | None) -> str:
    if not value:
        return ""
    value = html.unescape(value)
    value = re.sub(r"\s+", " ", value).strip()
    return value

def normalize_price(value: str | int | float | None) -> str:
    if value is None:
        return ""
    s = clean_text(str(value)).replace(",", ".")
    m = re.findall(r"\d+(?:\.\d+)?", s)
    return m[-1] if m else ""

def extract_json_ld(soup: BeautifulSoup) -> list[dict]:
    items = []
    for tag in soup.find_all("script", attrs={"type": "application/ld+json"}):
        txt = tag.get_text(strip=True)
        if not txt:
            continue
        try:
            data = json.loads(txt)
        except Exception:
            continue
        stack = data.get("@graph", []) if isinstance(data, dict) else data
        if isinstance(stack, dict):
            stack = [stack]
        for item in stack:
            if isinstance(item, dict):
                if "@graph" in item and isinstance(item["@graph"], list):
                    items.extend([x for x in item["@graph"] if isinstance(x, dict)])
                else:
                    items.append(item)
    return items

def find_product_schema(soup: BeautifulSoup) -> dict:
    for item in extract_json_ld(soup):
        typ = item.get("@type")
        if typ == "Product" or (isinstance(typ, list) and "Product" in typ):
            return item
    return {}

def image_from_schema(value) -> list[str]:
    if not value:
        return []
    if isinstance(value, str):
        return [value]
    if isinstance(value, list):
        out = []
        for x in value:
            out.extend(image_from_schema(x))
        return out
    if isinstance(value, dict):
        return [value.get("url") or value.get("contentUrl") or ""]
    return []

def unique(seq: Iterable[str]) -> list[str]:
    seen, out = set(), []
    for x in seq:
        x = clean_text(x)
        if not x:
            continue
        x = x.split()[0]
        if x.startswith("//"):
            x = "https:" + x
        if x.startswith("/"):
            x = urljoin(BASE_URL, x)
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out

def parse_product_html(page_url: str, html_text: str) -> Product:
    soup = BeautifulSoup(html_text, "lxml")
    schema = find_product_schema(soup)

    name = clean_text(schema.get("name"))
    if not name:
        h1 = soup.select_one("h1.product_title, h1.entry-title")
        name = clean_text(h1.get_text(" ") if h1 else "")

    sku = clean_text(schema.get("sku"))
    if not sku:
        sku_el = soup.select_one(".sku, [data-product_sku]")
        sku = clean_text(sku_el.get_text(" ") if sku_el and sku_el.get_text(strip=True) else sku_el.get("data-product_sku", "") if sku_el else "")

    price = ""
    offers = schema.get("offers")
    if isinstance(offers, dict):
        price = normalize_price(offers.get("price") or offers.get("lowPrice"))
    elif isinstance(offers, list):
        vals = [normalize_price(o.get("price") or o.get("lowPrice")) for o in offers if isinstance(o, dict)]
        price = next((v for v in vals if v), "")
    if not price:
        # Prefer visible EUR amounts and skip BGN mirror amounts (.amount-eur)
        amounts = []
        for el in soup.select(".summary .price .woocommerce-Price-amount, .summary .price bdi, p.price .woocommerce-Price-amount"):
            if "amount-eur" in (el.get("class") or []):
                continue
            txt = clean_text(el.get_text(" "))
            if "€" in txt or "euro" in txt.lower():
                amounts.append(normalize_price(txt))
        price = next((x for x in reversed(amounts) if x), "")

    cats = []
    for a in soup.select(".posted_in a, a[rel='tag']"):
        txt = clean_text(a.get_text(" "))
        href = a.get("href", "")
        if txt and "/produkt-kategoriya/" in href:
            cats.append(txt)
    if not cats:
        # Fallback from breadcrumb JSON-LD, excluding home/shop/product name
        for item in extract_json_ld(soup):
            if item.get("@type") == "BreadcrumbList":
                for li in item.get("itemListElement", []):
                    name_val = li.get("name") or (li.get("item") or {}).get("name")
                    if name_val and clean_text(name_val).lower() not in {"home", "начало", "магазин", name.lower()}:
                        cats.append(clean_text(name_val))
    category = " | ".join(unique(cats))

    desc = ""
    for sel in ["#tab-description", ".woocommerce-Tabs-panel--description", ".woocommerce-product-details__short-description"]:
        el = soup.select_one(sel)
        if el:
            desc = clean_text(el.get_text(" "))
            if desc:
                break
    if not desc:
        meta = soup.find("meta", attrs={"name": "description"})
        desc = clean_text(meta.get("content", "") if meta else schema.get("description", ""))

    imgs = []
    imgs.extend(image_from_schema(schema.get("image")))
    for a in soup.select(".woocommerce-product-gallery a[href], .product-images a[href]"):
        href = a.get("href", "")
        if "/wp-content/uploads/" in href:
            imgs.append(href)
    for img in soup.select(".woocommerce-product-gallery img, .product-images img, img.wp-post-image"):
        for attr in ["data-large_image", "data-src", "src"]:
            val = img.get(attr)
            if val and "/wp-content/uploads/" in val:
                imgs.append(val)
        srcset = img.get("srcset") or img.get("data-srcset")
        if srcset:
            for part in srcset.split(","):
                url = part.strip().split(" ")[0]
                if "/wp-content/uploads/" in url:
                    imgs.append(url)
    image_urls = " | ".join(unique(imgs))

    canonical = soup.find("link", rel="canonical")
    product_url = canonical.get("href") if canonical and canonical.get("href") else page_url

    return Product(sku, name, category, price, image_urls, desc, product_url)

def product_links_from_category(html_text: str, category_url: str) -> list[str]:
    soup = BeautifulSoup(html_text, "lxml")
    links = []
    selectors = [
        "h3.wd-entities-title a[href*='/produkt/']",
        "a.product-image-link[href*='/produkt/']",
        ".product-grid-item a[href*='/produkt/']",
        "a[href*='/produkt/']",
    ]
    for sel in selectors:
        for a in soup.select(sel):
            href = a.get("href", "")
            if "/produkt/" in href and "add-to-cart" not in href:
                links.append(urljoin(category_url, href).split("?")[0])
        if links:
            break
    return unique(links)

def next_page_url(html_text: str, current_url: str) -> Optional[str]:
    soup = BeautifulSoup(html_text, "lxml")
    for a in soup.select("a.next.page-numbers, .woocommerce-pagination a.next"):
        if a.get("href"):
            return urljoin(current_url, a["href"])
    return None

def category_page_url(base_category_url: str, page: int) -> str:
    if page == 1:
        return base_category_url
    return base_category_url.rstrip("/") + f"/page/{page}/"

def scrape_category_html(client: Client, category_url: str, max_pages: int = 0) -> list[str]:
    links = []
    page = 1
    current = category_url
    while True:
        if max_pages and page > max_pages:
            break
        log.info("Category page %s: %s", page, current)
        r = client.get(current)
        found = product_links_from_category(r.text, current)
        log.info("Found %s product links", len(found))
        if not found:
            break
        links.extend(found)
        nxt = next_page_url(r.text, current)
        if not nxt:
            # Woodmart sometimes has classic /page/N/ pagination but no next in parsed HTML
            test_next = category_page_url(category_url, page + 1)
            if test_next == current:
                break
            nxt = test_next
        page += 1
        current = nxt
    return unique(links)

def store_api_category_id(client: Client, slug: str) -> Optional[int]:
    url = f"{BASE_URL}/wp-json/wc/store/products/categories"
    try:
        r = client.get(url, params={"slug": slug, "per_page": 100})
        data = r.json()
        if isinstance(data, list) and data:
            return data[0].get("id")
    except Exception as exc:
        log.warning("Store categories API unavailable: %s", exc)
    return None

def scrape_store_api(client: Client, category_slug: str) -> list[Product]:
    cat_id = store_api_category_id(client, category_slug)
    if not cat_id:
        return []
    out = []
    page = 1
    while True:
        url = f"{BASE_URL}/wp-json/wc/store/products"
        log.info("Store API page %s category id %s", page, cat_id)
        r = client.get(url, params={"category": cat_id, "per_page": 100, "page": page})
        data = r.json()
        if not isinstance(data, list) or not data:
            break
        for p in data:
            prices = p.get("prices") or {}
            raw_price = prices.get("price") or prices.get("regular_price") or ""
            # Woo Store API often returns minor units; use decimals info when present.
            decimals = int(prices.get("currency_minor_unit", 2) or 2)
            price = ""
            if raw_price not in (None, ""):
                try:
                    price = f"{int(raw_price) / (10 ** decimals):.2f}"
                except Exception:
                    price = normalize_price(raw_price)
            cats = " | ".join([clean_text(c.get("name")) for c in p.get("categories", []) if c.get("name")])
            imgs = " | ".join(unique([i.get("src", "") for i in p.get("images", [])]))
            out.append(Product(
                sku=clean_text(p.get("sku", "")),
                product_name=clean_text(p.get("name", "")),
                category=cats,
                price_eur=price,
                image_urls=imgs,
                description=clean_text(BeautifulSoup(p.get("description") or p.get("short_description") or "", "lxml").get_text(" ")),
                product_url=p.get("permalink", ""),
            ))
        total_pages = int(r.headers.get("X-WP-TotalPages", "0") or 0)
        if total_pages and page >= total_pages:
            break
        page += 1
    return out

def merge_details(client: Client, products: list[Product]) -> list[Product]:
    detailed = []
    for i, p in enumerate(products, 1):
        if not p.product_url:
            detailed.append(p)
            continue
        log.info("Detail %s/%s %s", i, len(products), p.product_url)
        try:
            d = parse_product_html(p.product_url, client.get(p.product_url).text)
            # Prefer detail page for richer fields, keep API fields where detail is blank.
            detailed.append(Product(
                sku=d.sku or p.sku,
                product_name=d.product_name or p.product_name,
                category=d.category or p.category,
                price_eur=d.price_eur or p.price_eur,
                image_urls=d.image_urls or p.image_urls,
                description=d.description or p.description,
                product_url=d.product_url or p.product_url,
            ))
        except Exception as exc:
            log.warning("Detail failed for %s: %s", p.product_url, exc)
            detailed.append(p)
    return detailed

def write_xlsx(products: list[Product], output: str):
    df = pd.DataFrame([asdict(p) for p in products])
    df = df.rename(columns={
        "sku": "Код на продукт",
        "product_name": "Име на продукт",
        "category": "Категория",
        "price_eur": "Цена в евро",
        "image_urls": "URL на всички изображения",
        "description": "Описание на продукта",
        "product_url": "URL на продукта",
    })
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="products")
        ws = writer.book["products"]
        ws.freeze_panes = "A2"
        widths = [20, 45, 35, 14, 80, 100, 65]
        for idx, width in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + idx)].width = width
    log.info("Saved %s rows to %s", len(df), output)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--category-url", default=DEFAULT_CATEGORY_URL)
    parser.add_argument("--category-slug", default="uchenicheski-materiali")
    parser.add_argument("--output", default=f"apisexpress_products_{dt.date.today().isoformat()}.xlsx")
    parser.add_argument("--delay", type=float, default=0.7)
    parser.add_argument("--max-pages", type=int, default=0, help="0 = all pages")
    parser.add_argument("--html-only", action="store_true", help="Skip Woo Store API and scrape HTML pagination")
    parser.add_argument("--no-detail", action="store_true", help="Do not visit product detail pages")
    args = parser.parse_args()

    client = Client(delay=args.delay)
    products: list[Product] = []

    if not args.html_only:
        products = scrape_store_api(client, args.category_slug)
        if products:
            log.info("Store API returned %s products", len(products))

    if not products:
        links = scrape_category_html(client, args.category_url, args.max_pages)
        log.info("Total unique product links: %s", len(links))
        products = [Product(product_url=url) for url in links]

    if not args.no_detail:
        products = merge_details(client, products)

    # Deduplicate by URL/SKU.
    seen = set()
    deduped = []
    for p in products:
        key = p.product_url or p.sku or p.product_name
        if key and key not in seen:
            seen.add(key)
            deduped.append(p)
    write_xlsx(deduped, args.output)

if __name__ == "__main__":
    main()
