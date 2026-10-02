"""APIS Express complete public catalog -> eMAG preparation workbook.

This is a staging workbook, NOT an eMAG Create Products import file.
Requires the seller's category-specific eMAG creation templates for mapping.
"""
from __future__ import annotations

import html
import json
import logging
import os
import re
import threading
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from urllib.parse import urljoin, urlparse
from xml.etree import ElementTree as ET

import requests
import xlsxwriter
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE = 'https://apisexpress.com'
INDEX = f'{BASE}/sitemap.xml'
OUTPUT_DIR = Path('output')
WORKERS = max(1, min(4, int(os.getenv('WORKERS', '3'))))
MAX_IMAGES = 12
LOG = logging.getLogger('apisexpress')
LOCAL = threading.local()


def http():
    if not hasattr(LOCAL, 'session'):
        s = requests.Session()
        retry = Retry(total=4, connect=4, read=4, status=4,
                      backoff_factor=1.0, status_forcelist=[429, 500, 502, 503, 504],
                      allowed_methods={'GET'}, respect_retry_after_header=True)
        s.mount('https://', HTTPAdapter(max_retries=retry))
        s.headers.update({'User-Agent': 'Mozilla/5.0 (compatible; CatalogExport/1.0; public product metadata)',
                          'Accept-Language': 'bg-BG,bg;q=0.9',
                          'Cache-Control': 'no-cache'})
        LOCAL.session = s
    return LOCAL.session


def fetch(url):
    r = http().get(url, timeout=(15, 40))
    r.raise_for_status()
    return r.content


def xml_locations(payload):
    root = ET.fromstring(payload)
    return [el.text.strip() for el in root.iter() if el.tag.rsplit('}', 1)[-1] == 'loc' and el.text]


def discover():
    roots = xml_locations(fetch(INDEX))
    sitemaps = sorted({u for u in roots if re.search(r'/product-sitemap\d*\.xml(?:\?.*)?$', u)})
    if not sitemaps:
        raise RuntimeError('No product-sitemap XML files found. Site layout may have changed.')
    all_urls, failed = set(), []
    for sitemap in sitemaps:
        try:
            urls = {u.split('#')[0] for u in xml_locations(fetch(sitemap)) if '/produkt/' in u}
            all_urls.update(urls)
            LOG.info('%s: %s product links', sitemap, len(urls))
        except Exception as exc:
            LOG.exception('Sitemap could not be read: %s', sitemap)
            failed.append((sitemap, str(exc)))
    if not all_urls:
        raise RuntimeError('No product URLs discovered; aborting rather than exporting an empty workbook.')
    return sorted(all_urls), failed


def plain(node):
    if node is None:
        return ''
    if isinstance(node, str):
        node = BeautifulSoup(node, 'html.parser')
    return re.sub(r'\s+', ' ', html.unescape(node.get_text(' ', strip=True))).strip()


def jld(soup):
    products = []
    for script in soup.select('script[type="application/ld+json"]'):
        try:
            obj = json.loads(script.string or script.get_text())
        except (ValueError, TypeError):
            continue
        stack = [obj]
        while stack:
            cur = stack.pop()
            if isinstance(cur, list):
                stack.extend(cur)
            elif isinstance(cur, dict):
                typ = cur.get('@type', [])
                if typ == 'Product' or isinstance(typ, list) and 'Product' in typ:
                    products.append(cur)
                for val in cur.values():
                    if isinstance(val, (dict, list)):
                        stack.append(val)
    return products[0] if products else {}


def parse_eur_amount(node):
    """Read the EURO component of ONE WooCommerce price, never the BGN equivalent."""
    if not node:
        return None
    if 'amount-eur' in node.get('class', []) and 'лв' in plain(node).lower():
        return None
    value = plain(node)
    match = re.search(r'([\d\s.,]+)\s*€', value)
    if not match:
        return None
    raw = re.sub(r'\s+', '', match.group(1))
    if ',' in raw and '.' in raw:
        raw = raw.replace('.', '').replace(',', '.') if raw.rfind(',') > raw.rfind('.') else raw.replace(',', '')
    elif ',' in raw:
        raw = raw.replace(',', '.')
    try:
        return Decimal(raw)
    except InvalidOperation:
        return None


def scoped_price(root):
    """Scope to DIRECT child price of the current product's .summary-inner.

    Site has next/previous product navigation with OTHER product prices inside
    .summary-inner, so broad '.summary .price' is unsafe.
    """
    if not root:
        return None, None, 'product container missing'
    box = root.select_one('.summary-inner > p.price, .summary-inner > div.price, .entry-summary > p.price')
    if not box:
        return None, None, 'main price element missing'
    def first_amount(wrapper):
        if not wrapper:
            return None
        for amount in wrapper.select('.woocommerce-Price-amount'):
            price = parse_eur_amount(amount)
            if price is not None:
                return price
        return None
    sale = first_amount(box.select_one('ins'))
    old = first_amount(box.select_one('del'))
    if sale is not None:
        return sale, old, 'main HTML discounted'
    for a in box.select('.woocommerce-Price-amount'):
        if a.find_parent('del') or a.find_parent('ins'):
            continue
        price = parse_eur_amount(a)
        if price is not None:
            return price, old, 'main HTML regular'
    return None, old, 'price missing or variable'


def unique(seq):
    seen = set()
    for item in seq:
        if item and item not in seen:
            seen.add(item)
            yield item


def img_url(url):
    if not url:
        return ''
    url = urljoin(BASE, html.unescape(url).strip())
    p = urlparse(url)
    if p.netloc != urlparse(BASE).netloc or '/wp-content/uploads/' not in p.path:
        return ''
    # WordPress thumbnail suffix only; don't truncate arbitrary filename segments.
    path = re.sub(r'-\d+x\d+(?=\.(jpg|jpeg|png|webp|gif)$)', '', p.path, flags=re.I)
    return p._replace(path=path, query='', fragment='').geturl()


def gallery(root, product):
    out = []
    # Gallery ONLY; do not collect related products or pictures embedded in description.
    if root:
        for a in root.select('.woocommerce-product-gallery a[href], .product-images a[href]'):
            out.append(img_url(a.get('href')))
        for image in root.select('.woocommerce-product-gallery img, .product-images img'):
            for k in ('data-large_image', 'data-src', 'src'):
                out.append(img_url(image.get(k)))
    images = product.get('image', [])
    if isinstance(images, (str, dict)):
        images = [images]
    for image in images if isinstance(images, list) else []:
        out.append(img_url(image if isinstance(image, str) else image.get('url') or image.get('contentUrl')))
    return list(unique(out))[:MAX_IMAGES]


def attr_map(root):
    attrs = {}
    if root:
        for tr in root.select('.woocommerce-product-attributes tr'):
            th = tr.select_one('th')
            td = tr.select_one('td')
            if th and td:
                attrs[plain(th)] = plain(td)
    return attrs


def get_brand(prod, attrs, root):
    brand = prod.get('brand')
    if isinstance(brand, dict):
        brand = brand.get('name')
    if isinstance(brand, list):
        brand = next((v.get('name') if isinstance(v, dict) else v for v in brand if v), '')
    if brand:
        return plain(str(brand))
    for k, v in attrs.items():
        if any(x in k.casefold() for x in ('марка', 'производител', 'brand')):
            return v
    if root:
        el = root.select_one('.wd-product-brands a, .product-brands a')
        if el:
            return plain(el)
    return ''


def get_categories(root):
    if not root:
        return ''
    # On this site the final breadcrumb item is product title.
    crumb = root.select_one('nav.woocommerce-breadcrumb, .woocommerce-breadcrumb, .wd-breadcrumbs')
    links = crumb.select('a[href*="produkt-kategoriya"]') if crumb else []
    cats = [plain(a) for a in links if plain(a)]
    if not cats:
        cats = [plain(a) for a in root.select('.posted_in a[href*="produkt-kategoriya"]')]
    if not cats:
        # Some Woodmart pages render breadcrumbs outside the product container.
        # Preserve category slugs on the product element instead of leaving blank.
        slugs = [c.replace('product_cat-', '').replace('-', ' ') for c in root.get('class', [])
                 if c.startswith('product_cat-')]
        cats = slugs
    return ' > '.join(unique(cats))


def availability(root, product):
    if not root:
        return 'неизвестна', None
    explicit = root.select_one('p.stock.out-of-stock, p.stock.in-stock, .stock')
    desc = plain(explicit).lower() if explicit else ''
    offer = product.get('offers') or {}
    if isinstance(offer, list):
        offer = offer[0] if offer else {}
    schema = str(offer.get('availability', '')).lower() if isinstance(offer, dict) else ''
    if 'не е налич' in desc or 'изчерпан' in desc or 'outofstock' in schema or 'out-of-stock' in root.get('class', []):
        return 'не е наличен', 0
    qty = re.search(r'(?:(?:налични|в наличност|in stock)\s*:?\s*)(\d+)\s*(?:бр\.?|броя|pcs)?', desc)
    if qty:
        return 'наличен', int(qty.group(1))
    if 'налич' in desc or 'instock' in schema or 'instock' in root.get('class', []):
        return 'наличен', None  # no numeric stock on public site
    return 'неизвестна', None


def parse_page(markup, url):
    soup = BeautifulSoup(markup, 'html.parser')
    root = soup.select_one('.single-product-page[id^="product-"], div.single-product-content[id^="product-"]')
    if not root:
        raise ValueError('Single WooCommerce product container not found')
    prod = jld(soup)
    h1 = root.select_one('h1.product_title, h1.entry-title')
    name = plain(h1) or plain(str(prod.get('name') or ''))
    if not name:
        raise ValueError('Product title missing')
    attrs = attr_map(root)
    site_id = re.search(r'product-(\d+)', root.get('id', ''))
    site_id = site_id.group(1) if site_id else ''
    sku = plain(root.select_one('.sku')) or str(prod.get('sku') or '').strip()
    # avoid scraping a related product's SKU from HTML page text.
    gtin = next((str(prod.get(k) or '').strip() for k in ('gtin13', 'gtin12', 'gtin8', 'gtin14', 'gtin') if prod.get(k)), '')
    for label, val in attrs.items():
        if not gtin and any(word in label.casefold() for word in ('ean', 'баркод', 'gtin')):
            gtin = re.sub(r'\D', '', val)
    if gtin and (not gtin.isdecimal() or len(gtin) not in (8, 12, 13, 14)):
        gtin = ''
    # Longer on-page description is preferred to short SEO snippet.
    description = root.select_one('#tab-description, .woocommerce-Tabs-panel--description')
    desc = plain(description) or plain(root.select_one('.woocommerce-product-details__short-description'))
    if not desc:
        desc = plain(str(prod.get('description') or ''))
    price, regular, price_source = scoped_price(root)
    stock_status, stock_qty = availability(root, prod)
    imgs = gallery(root, prod)
    variant = 'product-type-variable' in root.get('class', []) or bool(root.select_one('form.variations_form'))
    notes = []
    if not sku: notes.append('Липсва SKU')
    if not gtin: notes.append('Липсва EAN/GTIN')
    if price is None: notes.append('Цената не е потвърдена от основния HTML')
    if not imgs: notes.append('Липсват изображения')
    if stock_status == 'неизвестна': notes.append('Неясна наличност')
    elif stock_qty is None: notes.append('Потвърдете числова наличност')
    if variant: notes.append('Вариативен продукт: отделните варианти изискват проверка')
    row = {'ID в изходния сайт': site_id, 'SKU': sku, 'EAN/GTIN': gtin,
           'Име': name, 'Категория APIS': get_categories(root),
           'Марка': get_brand(prod, attrs, root), 'Описание': desc,
           'Цена EUR с ДДС (сайт)': float(price) if price is not None else None,
           'Стара цена EUR с ДДС': float(regular) if regular is not None else None,
           'Източник на цена': price_source,
           'Наличност (статус)': stock_status, 'Наличност (брой)': stock_qty,
           'Вариативен продукт': 'Да' if variant else 'Не',
           'Характеристики (от сайта)': ' | '.join(f'{k}: {v}' for k, v in attrs.items()),
           'URL продукт': url, 'Проверка преди eMAG': ' | '.join(notes)}
    for i, item in enumerate(imgs, 1):
        row[f'Изображение {i}'] = item
    return row


def safe_xl(value):
    if value is None:
        return ''
    if isinstance(value, str):
        value = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', value)[:32000]
        # Excel formula injection: do not turn scraped values into formulas.
        if value.startswith(('=', '+', '-', '@')):
            value = "'" + value
    return value


def write_excel(rows, errors):
    OUTPUT_DIR.mkdir(exist_ok=True)
    filename = OUTPUT_DIR / 'APIS_full_catalog_for_eMAG_preparation.xlsx'
    columns = ['ID в изходния сайт', 'SKU', 'EAN/GTIN', 'Име', 'Категория APIS',
               'Марка', 'Описание', 'Цена EUR с ДДС (сайт)', 'Стара цена EUR с ДДС',
               'Източник на цена', 'Наличност (статус)', 'Наличност (брой)',
               'Вариативен продукт', 'Характеристики (от сайта)', 'URL продукт',
               'Проверка преди eMAG'] + [f'Изображение {i}' for i in range(1, MAX_IMAGES + 1)]
    wb = xlsxwriter.Workbook(str(filename), {'strings_to_urls': False})
    title = wb.add_format({'bold': True, 'bg_color': '#17365D', 'font_color': 'white', 'text_wrap': True, 'valign': 'vcenter'})
    body = wb.add_format({'valign': 'top'})
    number = wb.add_format({'num_format': '0.00', 'valign': 'top'})
    notice = wb.add_format({'bold': True, 'font_color': '#B91C1C', 'text_wrap': True})
    sheet = wb.add_worksheet('Каталог APIS')
    sheet.freeze_panes(1, 0)
    sheet.set_row(0, 34)
    sheet.set_column(0, 2, 19)
    sheet.set_column(3, 3, 50)
    sheet.set_column(4, 5, 27)
    sheet.set_column(6, 6, 66)
    sheet.set_column(7, 12, 21)
    sheet.set_column(13, 13, 54)
    sheet.set_column(14, 15, 58)
    sheet.set_column(16, 16 + MAX_IMAGES, 62)
    for c, name in enumerate(columns): sheet.write(0, c, name, title)
    for r, row in enumerate(rows, 1):
        for c, col in enumerate(columns):
            val = safe_xl(row.get(col))
            if val != '':
                sheet.write(r, c, val, number if col in ('Цена EUR с ДДС (сайт)', 'Стара цена EUR с ДДС') else body)
    sheet.autofilter(0, 0, max(1, len(rows)), len(columns)-1)
    checks = wb.add_worksheet('За eMAG - прочети')
    checks.set_column('A:A', 31)
    checks.set_column('B:B', 95)
    info = [
        ('ВНИМАНИЕ', 'ТОЗИ ФАЙЛ Е ЗА ПОДГОТОВКА. НЕ Е eMAG Create Products файл и не може да се импортира директно.'),
        ('Причина', 'Каченият от потребителя шаблон е Update Offer Price/Stock и съдържа идентификатори на вече съществуващи оферти.'),
        ('Следваща стъпка', 'Изтеглете от eMAG Seller → Продукти → Продуктови категории → Съдържание на продукта → Създаване на продукти: отделен шаблон за необходимите категории.'),
        ('EAN', 'Проверете истинските EAN/GTIN. Не генерирайте баркодове от ID, SKU или заглавие.'),
        ('Цени', 'Сайтът показва цени EUR с ДДС. eMAG изисква продажна цена БЕЗ ДДС. Потвърдете данъчната ставка, продажната цена и маржа.'),
        ('Количество', 'Общото „наличен“ не означава числова наличност. Потвърдете реални количества за Вашите оферти.'),
        ('Право на продажба', 'Проверете право да препродавате, наличности, изображения и продуктова документация.'),
        ('Категории', 'Категориите в APIS не са автоматично равни на eMAG категориите. Важат специфични атрибути и допустими стойности.'),
        ('Вариации', 'Родителските вариативни продукти са обозначени; отделните SKU и снимки на варианти може да изискват допълнителни данни.'),
        ('Брой продукти', len(rows)), ('Брой грешки', len(errors)),
    ]
    for r, (key, val) in enumerate(info):
        checks.write(r, 0, key, title if r == 0 else body)
        checks.write(r, 1, val, notice if r == 0 else body)
    checks.set_row(0, 40)
    category = wb.add_worksheet('Категории')
    category.set_column('A:A', 78); category.set_column('B:B', 20)
    category.write_row(0, 0, ['Категория APIS', 'Брой продукти'], title)
    counts = Counter(row['Категория APIS'] or '(неопределена)' for row in rows)
    for r, (cat, n) in enumerate(sorted(counts.items()), 1):
        category.write(r, 0, safe_xl(cat)); category.write_number(r, 1, n)
    error = wb.add_worksheet('Грешки')
    error.set_column('A:A', 78);error.set_column('B:B', 100)
    error.write_row(0, 0, ['URL', 'Грешка'], title)
    for r, (u, msg) in enumerate(errors, 1): error.write_row(r, 0, [safe_xl(u), safe_xl(msg)])
    wb.close()
    LOG.info('Exported %s (%d products, %d errors)', filename, len(rows), len(errors))
    return filename


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s',
                        handlers=[logging.StreamHandler(), logging.FileHandler(OUTPUT_DIR / 'scrape.log', encoding='utf-8')])
    urls, sitemap_errors = discover()
    LOG.info('Discovered %d unique product pages', len(urls))
    rows, errors = [], list(sitemap_errors)
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = {pool.submit(fetch, url): url for url in urls}
        for n, future in enumerate(as_completed(futures), 1):
            url = futures[future]
            try:
                row = parse_page(future.result(), url)
                rows.append(row)
            except Exception as exc:
                errors.append((url, str(exc)))
                LOG.warning('Could not extract %s: %s', url, exc)
            if n % 50 == 0 or n == len(urls):
                LOG.info('Processed %d/%d | extracted %d | errors %d', n, len(urls), len(rows), len(errors))
    if not rows:
        raise RuntimeError('All product extractions failed; see output/scrape.log')
    rows.sort(key=lambda r: (r['Категория APIS'], r['Име']))
    write_excel(rows, errors)
    with open(OUTPUT_DIR / 'coverage.txt', 'w', encoding='utf8') as f:
        f.write(f'Product URL count: {len(urls)}\nExported: {len(rows)}\nErrors: {len(errors)}\n')
        f.write('Reminder: this is NOT a directly importable eMAG new-product template.\n')
        for url, err in errors: f.write(f'{url}\t{err}\n')
    # Never silently call an incomplete scrape 'success'.
    if sitemap_errors or len(errors) > max(10, len(urls) * 0.02):
        raise RuntimeError(f'Coverage warning: {len(errors)} failures; workbook saved for review.')


if __name__ == '__main__':
    main()
