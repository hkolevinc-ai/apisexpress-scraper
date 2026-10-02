"""APIS Express -> existing eMAG offer price/stock update template.
Does not create new offers. Matches only template SKUs/EANs, keeps all other fields.
"""
import csv
import json
import logging
import re
import time
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation
from pathlib import Path
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

BASE = 'https://apisexpress.com'
TEMPLATE = Path('template/emag_template.xlsx')
OUTPUT = Path('output/emag_price_stock_update.xlsx')
REPORT = Path('output/matching_report.csv')
TIMEOUT = 30
DELAY = 0.25
HEADERS = {'User-Agent': 'Mozilla/5.0 (compatible; CatalogPriceUpdater/1.0)',
           'Accept-Language': 'bg-BG,bg;q=0.9', 'Cache-Control': 'no-cache'}
session = requests.Session()
session.headers.update(HEADERS)
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')


def fetch(url):
    for attempt in range(4):
        try:
            r = session.get(url, timeout=TIMEOUT)
            r.raise_for_status()
            return r.text
        except requests.RequestException:
            if attempt == 3:
                raise
            time.sleep(2 ** attempt)


def locs(xml):
    root = ET.fromstring(xml)
    return [x.text.strip() for x in root.iter() if x.tag.endswith('loc') and x.text]


def product_urls():
    index = locs(fetch(BASE + '/sitemap.xml'))
    maps = [u for u in index if re.search(r'/product-sitemap\d*\.xml(?:\?.*)?$', u)]
    if not maps:
        raise RuntimeError('No product sitemaps found; refusing to export unchanged data.')
    urls = []
    for sm in maps:
        entries = [u for u in locs(fetch(sm)) if '/produkt/' in u]
        logging.info('%s: %d URLs', sm, len(entries))
        urls.extend(entries)
    return list(dict.fromkeys(urls))


def text(el):
    return el.get_text(' ', strip=True) if el else ''


def normalize_id(value):
    return re.sub(r'[^\w]', '', str(value or '').casefold(), flags=re.UNICODE)


def jsonld_products(soup):
    products = []
    for script in soup.select('script[type="application/ld+json"]'):
        try:
            raw = json.loads(script.string or script.get_text())
        except (ValueError, TypeError):
            continue
        stack = [raw]
        while stack:
            obj = stack.pop()
            if isinstance(obj, list):
                stack.extend(obj)
            elif isinstance(obj, dict):
                types = obj.get('@type', [])
                if isinstance(types, str):
                    types = [types]
                if 'Product' in types:
                    products.append(obj)
                stack.extend(v for v in obj.values() if isinstance(v, (list, dict)))
    return products


def parse_money(raw):
    """Accept decimal numbers with comma/dot and thousands separators."""
    s = re.sub(r'[^0-9,.]', '', str(raw or ''))
    if not s:
        return None
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.') if s.rfind(',') > s.rfind('.') else s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        value = Decimal(s)
        return value if value > 0 else None
    except InvalidOperation:
        return None


def eur_amounts(price_el):
    result = []
    if not price_el:
        return result
    for amount in price_el.select('.woocommerce-Price-amount'):
        value = text(amount)
        # Currency must be explicitly EUR; never silently convert BGN.
        if '€' not in value and not re.search(r'\bEUR\b', value, re.I):
            continue
        price = parse_money(value)
        if price is not None:
            result.append((amount, price))
    return result


def price_from_html(soup):
    # Select the main product summary, never related products or cross-sells.
    summary = soup.select_one('.single-product .summary-inner, .single-product .entry-summary, .single-product .summary, .product-main .summary')
    if summary is None:
        return None, 'missing_main_product_summary'
    box = summary.select_one('p.price, div.price')
    if box is None:
        return None, 'missing_main_price'
    amounts = eur_amounts(box)
    sale = [p for el, p in amounts if el.find_parent('ins')]
    if sale:
        return sale[0], 'html_sale_eur'
    regular = [p for el, p in amounts if not el.find_parent('del')]
    if len(set(regular)) == 1:
        return regular[0], 'html_eur'
    return None, 'ambiguous_or_missing_eur_price'


def price_from_store_api(slug):
    try:
        response = session.get(BASE + '/wp-json/wc/store/v1/products', params={'slug': slug}, timeout=TIMEOUT)
        if response.status_code != 200:
            return None
        matches = response.json()
        if not isinstance(matches, list) or len(matches) != 1:
            return None
        prices = matches[0].get('prices', {})
        if prices.get('currency_code') != 'EUR':
            return None
        scale = Decimal(10) ** int(prices.get('currency_minor_unit', 2))
        raw = prices.get('price')
        return Decimal(str(raw)) / scale if raw not in (None, '') else None
    except (requests.RequestException, ValueError, TypeError, InvalidOperation):
        return None


def extract(soup, url):
    product = jsonld_products(soup)
    data = product[0] if product else {}
    sku_el = soup.select_one('.product_meta .sku, .sku_wrapper .sku')
    sku = text(sku_el) or str(data.get('sku') or '')
    eans = set()
    for key in ('gtin', 'gtin8', 'gtin12', 'gtin13', 'gtin14', 'ean'):
        if data.get(key):
            eans.add(normalize_id(data[key]))
    for item in soup.select('[itemprop="gtin13"], [itemprop="gtin"]'):
        eans.add(normalize_id(item.get('content') or text(item)))
    html_price, source = price_from_html(soup)
    slug = urlparse(url).path.strip('/').split('/')[-1]
    api_price = price_from_store_api(slug)
    # If two independently retrieved live sources disagree, avoid a silent bad update.
    if html_price is not None and api_price is not None and html_price != api_price:
        price, source = None, f'PRICE_CONFLICT html={html_price} api={api_price}'
    elif html_price is not None:
        price = html_price
    elif api_price is not None:
        price, source = api_price, 'store_api_eur'
    else:
        price = None
    stock = None
    stock_el = soup.select_one('.single-product .stock')
    stock_text = text(stock_el).lower()
    stock_class = ' '.join(stock_el.get('class', [])) if stock_el else ''
    if 'out-of-stock' in stock_class or 'outofstock' in stock_class or 'изчерпан' in stock_text:
        stock = 0
    elif stock_el:
        match = re.search(r'(?:налични|наличност|in stock)\s*:?\s*(\d+)', stock_text)
        if match:
            stock = int(match.group(1))
    # A generic "in stock" does NOT reveal exact stock quantity.
    return {'sku': normalize_id(sku), 'eans': eans, 'price': price, 'price_source': source,
            'stock': stock, 'url': url, 'name': text(soup.select_one('h1.product_title'))}


def main():
    wb = load_workbook(TEMPLATE)
    ws = wb['оферти']
    fields = {str(cell.value): cell.column for cell in ws[3] if cell.value}
    required = ('part_number', 'ean', 'sale_price', 'stock', 'status', 'offer_currency')
    if any(k not in fields for k in required):
        raise RuntimeError('Template columns changed: ' + str(fields))
    targets = {}
    for row in range(6, ws.max_row + 1):
        sku = normalize_id(ws.cell(row, fields['part_number']).value)
        ean = normalize_id(ws.cell(row, fields['ean']).value)
        if sku or ean:
            targets[row] = {'sku': sku, 'ean': ean, 'matches': []}
    logging.info('Existing offers: %d', len(targets))
    urls = product_urls()
    logging.info('Unique sitemap products: %d', len(urls))
    if not urls:
        raise RuntimeError('No product URLs')
    errors = []
    for i, url in enumerate(urls, 1):
        try:
            item = extract(BeautifulSoup(fetch(url), 'html.parser'), url)
            for target in targets.values():
                if (target['sku'] and target['sku'] == item['sku']) or (target['ean'] and target['ean'] in item['eans']):
                    target['matches'].append(item)
            if i % 100 == 0:
                logging.info('Checked %d/%d products', i, len(urls))
        except Exception as exc:
            errors.append((url, str(exc)))
            logging.warning('Failed %s: %s', url, exc)
        time.sleep(DELAY)
    OUTPUT.parent.mkdir(exist_ok=True)
    report = []
    for row, target in targets.items():
        matches = target['matches']
        status = 'not_found'
        if len(matches) == 1:
            item = matches[0]
            if item['price'] is not None:
                ws.cell(row, fields['sale_price']).value = float(item['price'])
                status = 'price_updated'
            else:
                status = 'price_unverified'
            if item['stock'] is not None:
                ws.cell(row, fields['stock']).value = item['stock']
                status += '_stock_updated'
            # Keep original offer status, vendor ID, PNK, VAT, other protected fields.
        elif len(matches) > 1:
            status = 'ambiguous_multiple_matches'
        report.append({'excel_row': row, 'sku': target['sku'], 'ean': target['ean'],
                       'status': status, 'matched_urls': ' | '.join(m['url'] for m in matches),
                       'price_sources': ' | '.join(m['price_source'] for m in matches)})
    wb.save(OUTPUT)
    with REPORT.open('w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=report[0].keys())
        writer.writeheader()
        writer.writerows(report)
    with Path('output/scrape_errors.csv').open('w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['url', 'error'])
        writer.writerows(errors)
    logging.info('Saved %s; matched %d/%d offers; fetch errors %d', OUTPUT,
                 sum(bool(x['matches']) for x in targets.values()), len(targets), len(errors))


if __name__ == '__main__':
    main()
