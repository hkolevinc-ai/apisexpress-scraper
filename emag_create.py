"""APIS Express complete catalog + eMAG 'Други ученически аксесоари' create-product template.

The source .xlsx template is preserved at XML package level: only Template!A6:AW... is
populated, leaving eMAG instructions, validation lists, formulas and layout intact.
DRAFT must not be uploaded without completing the mandatory/legal fields.
"""
from __future__ import annotations
import copy
import csv
import json
import logging
import os
import re
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as ET

import full_catalog as fc

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / 'templates/emag_create_school_accessories.xlsx'
CONFIG = ROOT / 'config.json'
OUTPUT = Path('output')
XL = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
N = '{' + XL + '}'
SHEET_PATH = 'xl/worksheets/sheet6.xml'  # Verified for this exact uploaded template
TEMPLATE_HEADER = ['part_number','vendor_ext_id','ean','sale_price','original_sale_price',
'vat_rate','status','offer_currency','stock','handling_time','lead_time','warranty',
'offer_properties','min_sale_price','max_sale_price','name','brand','description','url',
'source_language','main_image_url','other_image_url1','other_image_url2',
'other_image_url3','other_image_url4','other_image_url5','family_id','family_name',
'family_type','Тип продукт: [5704]','За: [9084]','Формат: [3580]','Цвят: [5401]',
'Материал: [6372]','Съдържание пакет: [6556]','Височина: [6779]',
'Ширина: [6780]','Дължина: [6862]','Тегло: [6878]',
'Брой/комплект: [7727]','Приказка /Герой: [9030]','Препоръчан за: [9840]',
'safety_information','manufacturer_name','manufacturer_address','manufacturer_email',
'responsible_name','responsible_address','responsible_email','rejected','warning','info']

# Only exact, distinctive accessory types from this eMAG category.
# Ordered from specific to general; pencils themselves do NOT match these types.
RULES = [
 ('Бутилка за вода', (r'бутилка за вода', r'шише за вода', r'water bottle')),
 ('Чанта за обяд', (r'чанта за обяд', r'lunch bag')),
 ('Kутия за обяд', (r'кутия за обяд', r'кутия за храна', r'lunch box')),
 ('Училищни етикети', (r'ученически етикет', r'етикет за тетрад', r'училищни етикет')),
 ('Училищно разписание', (r'учебно разписание', r'училищно разписание')),
 ('Удължител за молив', (r'удължител за молив',)),
 ('Защита на молив', (r'протектор за молив', r'защита на молив')),
 ('Комплект коректор', (r'комплект коректор',)),
 ('Дидактически инструмент', (r'дидактически инструмент',)),
 ('Контейнер за вода', (r'контейнер за вода',)),
 ('Модел за рисуване', (r'модел за рисуване',)),
 ('Макет за оцветяване', (r'макет за оцветяване',)),
 ('Стикер', (r'стикери', r'стикер', r'лепенки за тетрад')),
 ('Покритие', (r'подвързи', r'обложк', r'подвързия', r'покритие за тетрад')),
 ('Клипборд', (r'клипборд', r'clipboard')),
 ('Органайзер', (r'органайзер',)),
 ('Престилка', (r'престилка за рисуване', r'ученическа престилка')),
 ('Глобус', (r'глобус',)),
 ('Хербарий', (r'хербарий',)),
 ('Тефтер', (r'тефтер',)),
 ('Журнал', (r'журнал',)),
 ('Шаблон', (r'шаблон за рисуване', r'шаблон за чертане')),
 ('Значка', (r'значка',)),
 ('Брокат', (r'брокат',)),
 ('Отливка', (r'отливка',)),
]
SCHOOL_WORDS = ('ученическ', 'училищн', 'за училище', 'school', 'тетрад', 'писмен',
                'канцелар', 'детска градина', 'рисуване и хоби', 'материали за рисуване',
                'uchenicheski', 'za uchilishte', 'za detska gradina', 'risuvane')
ACCEPTED_IMAGES = {'.jpg','.jpeg','.png','.gif'}
REQUIRED_FIELDS = ('part_number','vendor_ext_id','sale_price','vat_rate','offer_currency',
                   'name','brand','description','main_image_url',
                   'Тип продукт: [5704]','За: [9084]',
                   'manufacturer_name','manufacturer_address','manufacturer_email')


def column_name(number):
    out = ''
    while number:
        number, rem = divmod(number - 1, 26)
        out = chr(65 + rem) + out
    return out


def _shared_strings(z):
    return [''.join(el.itertext()) for el in ET.fromstring(z.read('xl/sharedStrings.xml'))]


def _cell_value(el, shared):
    v = el.find(N + 'v')
    if el.get('t') == 's' and v is not None: return shared[int(v.text)]
    if el.get('t') == 'inlineStr': return ''.join(el.itertext())
    return v.text if v is not None else ''


def template_allowed():
    """Look up the actual options in 'Характеристики'; fail if template changed."""
    with ZipFile(TEMPLATE) as z:
        ss = _shared_strings(z)
        cat = ET.fromstring(z.read('xl/worksheets/sheet5.xml'))
        vals = {}
        for r in cat.findall('.//' + N + 'sheetData/' + N + 'row'):
            if int(r.get('r', '0')) < 6: continue
            for c in r.findall(N+'c'):
                ref = c.get('r','')
                match = re.match(r'^([A-Z]+)\d+$', ref)
                if not match:continue
                key = match.group(1)
                value = _cell_value(c, ss).strip()
                if value:vals.setdefault(key, set()).add(value)
        source = ET.fromstring(z.read(SHEET_PATH))
        row3 = source.find('.//' + N + 'sheetData/' + N + 'row[@r="3"]')
        names = [_cell_value(c,ss) for c in row3.findall(N+'c')]
        if names != TEMPLATE_HEADER:raise ValueError('Unexpected eMAG template version/column layout. Stop to prevent misaligned import.')
        return vals


def ean_valid(s):
    s = str(s or '').strip()
    if not s.isdecimal() or len(s) not in (8,12,13,14):return False
    digits = [int(d) for d in s]
    factor = 3
    total = 0
    for digit in reversed(digits[:-1]):
        total += digit * factor
        factor = 1 if factor == 3 else 3
    return (10 - total % 10) % 10 == digits[-1]


def classify(row, config, allowed):
    url = row['URL продукт']; sku = str(row.get('SKU') or '').strip()
    if url in config.get('exclude_product_urls',[]): return None, 'Ръчно изключен URL'
    override = config.get('product_type_by_sku',{}).get(sku,'')
    title = str(row.get('Име') or '').casefold()
    cats = str(row.get('Категория APIS') or '').casefold()
    typ = override
    matched = False
    if not typ:
        for label, patterns in RULES:
            if any(re.search(p, title) for p in patterns):
                typ = label;matched=True;break
    if typ and typ not in allowed.get('A',set()):
        return None, f'Непозволен тип продукт в шаблона: {typ}'
    school = any(k in cats or k in title for k in SCHOOL_WORDS)
    if not school and url not in config.get('include_product_urls',[]):
        return None, 'Категорията не е потвърдена като ученически аксесоари'
    if not typ:
        return '', 'Възможен аксесоар, но типът не може да се определи автоматично'
    return typ, 'Автоматично съпоставен по наименование' if matched else 'Ръчно включен'


def decimal_value(number):
    if number is None or number == '':return None
    try:return Decimal(str(number))
    except Exception:return None


def net_price(gross, vat):
    gross, vat = decimal_value(gross), decimal_value(vat)
    if gross is None or vat is None or gross <= 0:return None
    return (gross / (Decimal('1') + vat)).quantize(Decimal('0.0001'), rounding=ROUND_HALF_UP)


def _image_ok(img):
    if not img:return False
    from urllib.parse import urlparse
    return Path(urlparse(img).path).suffix.lower() in ACCEPTED_IMAGES


def build_offer(row, config, allowed):
    sku = str(row.get('SKU') or '').strip()
    typ, classification = classify(row, config, allowed)
    if typ is None: return None, classification
    issues = []
    raw_images = [row.get(f'Изображение {i}') or '' for i in range(1, fc.MAX_IMAGES+1)]
    images = [img for img in raw_images if _image_ok(img)]
    if not images and any(raw_images):issues.append('Снимките са само WEBP или неподдържан формат: необходим JPG/PNG URL')
    stock = config.get('stock_by_sku', {}).get(sku)
    if stock is None:stock = config.get('stock_by_site_id',{}).get(str(row.get('ID в изходния сайт') or ''))
    if stock is None:stock = config.get('default_stock')
    if stock is None:issues.append('Непотвърдено числово количество')
    else:
        try:
            stock = int(stock)
            if stock < 0:raise ValueError('negative')
        except (TypeError,ValueError):
            issues.append('Невалидно числово количество');stock = None
    if not config.get('vat_confirmed',False):issues.append('ДДС ставката не е потвърдена в config.json')
    vat = decimal_value(config.get('vat_rate')) if config.get('vat_confirmed',False) else None
    price = net_price(row.get('Цена EUR с ДДС (сайт)'),vat) if vat is not None else None
    old = net_price(row.get('Стара цена EUR с ДДС'),vat) if vat is not None else None
    brand = config.get('brand_by_sku',{}).get(sku) or row.get('Марка') or ''
    if not brand:issues.append('Липсва проверен бранд')
    manufacturer = config.get('manufacturer_by_sku',{}).get(sku) or config.get('manufacturer_by_brand',{}).get(brand,{})
    responsible = config.get('responsible_by_sku',{}).get(sku) or config.get('responsible_by_brand',{}).get(brand,{})
    manufacturer = manufacturer if isinstance(manufacturer,dict) else {}
    responsible = responsible if isinstance(responsible,dict) else {}
    audience = config.get('audience_by_sku',{}).get(sku,'')
    if not audience:
        title = str(row.get('Име') or '').casefold()
        if 'за момичета' in title:audience = 'Момичета'
        elif 'за момчета' in title:audience = 'Момчета'
    if audience not in allowed.get('B',set()):audience = '';issues.append('Задължително поле „За“: Момичета / Момчета — потвърдете')
    ean = config.get('ean_by_sku',{}).get(sku) or row.get('EAN/GTIN') or ''
    if not ean and ean_valid(sku):ean = sku
    if ean and not ean_valid(ean):issues.append('Невалиден EAN — премахнат');ean = ''
    if not ean:issues.append('EAN липсва — проверете дали е задължителен за артикула')
    if row.get('Вариативен продукт') == 'Да':issues.append('Родителски вариативен продукт — проверете SKU/вариации')
    if row.get('Наличност (статус)') != 'наличен':issues.append('На сайта продуктът не е потвърден като наличен')
    if row.get('Наличност (статус)') == 'не е наличен' and stock and stock > 0:
        issues.append('Конфликт: сайтът показва неналичен, а конфигурацията задава наличност')
    for key in ('name','address','email'):
        if not manufacturer.get(key):issues.append('Липсват данни за производителя: '+key)
    # Responsible person needed if manufacturer outside EU; require seller's explicit applicability decision.
    if 'eu_responsible_required' not in manufacturer:issues.append('Потвърдете необходимостта от отговорно лице в ЕС')
    elif manufacturer.get('eu_responsible_required'):
        for key in ('name','address','email'):
            if not responsible.get(key):issues.append('Липсват данни за отговорното лице в ЕС: '+key)
    safety = config.get('safety_by_sku',{}).get(sku,'')
    if not safety:issues.append('Проверете приложимата информация за безопасност')
    if not typ:issues.append('Тип продукт (задължително) не е уточнен')
    if not price:issues.append('Цена без ДДС не е потвърдена')
    if not row.get('Описание'):issues.append('Липсва описание')
    if not row.get('ID в изходния сайт'):issues.append('Липсва цифров ID на продукта')
    # Even a valid import may remain inactive if stock is 0 or site reports out of stock.
    status = 1 if stock is not None and stock > 0 and row.get('Наличност (статус)')=='наличен' else 0
    offer = {
      'part_number':sku,'vendor_ext_id':str(row.get('ID в изходния сайт') or ''),
      'ean':str(ean), 'sale_price':price, 'original_sale_price':old if old and price and old > price else '',
      'vat_rate':vat, 'status':status, 'offer_currency':'EUR', 'stock':stock,
      'name':str(row.get('Име') or ''), 'brand':brand,
      'description':str(row.get('Описание') or ''),
      # 'url' is OWN shop URL; APIS is source site, not necessarily seller-owned.
      'source_language':'Български (bg_BG)',
      'Тип продукт: [5704]':typ, 'За: [9084]':audience,
      'safety_information':safety,
      'manufacturer_name':manufacturer.get('name',''), 'manufacturer_address':manufacturer.get('address',''),
      'manufacturer_email':manufacturer.get('email',''),
      'responsible_name':responsible.get('name','') if manufacturer.get('eu_responsible_required') else '',
      'responsible_address':responsible.get('address','') if manufacturer.get('eu_responsible_required') else '',
      'responsible_email':responsible.get('email','') if manufacturer.get('eu_responsible_required') else '',
    }
    for n,img in enumerate(images[:6]): offer['main_image_url' if n==0 else f'other_image_url{n}'] = img
    for field in REQUIRED_FIELDS:
        if offer.get(field) is None or offer.get(field) == '':issues.append('Задължително поле липсва: '+field)
    if stock is None:issues.append('Липсва наличност')
    # Pricing and stock MUST be confirmed, no guessed defaults.
    status_flag = 'ГОТОВ ЗА ПРОВЕРКА' if not issues else 'НЕПЪЛЕН – НЕ КАЧВАЙ'
    audit = {'URL':row.get('URL продукт',''),'SKU':sku,'ID':offer['vendor_ext_id'],
             'Име':offer['name'],'Тип продукт':typ,'Статус':status_flag,
             'Причина за включване':classification,'Проблеми':' | '.join(dict.fromkeys(issues)),
             'Цена EUR с ДДС (APIS)':row.get('Цена EUR с ДДС (сайт)'),
             'Изчислена цена EUR без ДДС':str(price) if price else '',
             'Наличност на сайта':row.get('Наличност (статус)'),'Зададено количество':stock or 0,
             'Всички оригинални снимки':' | '.join(i for i in raw_images if i)}
    return (offer,audit),classification


def _write_cell(row_el, index, row_num, value):
    if value is None or value == '': return
    c=ET.SubElement(row_el,N+'c',r=f'{column_name(index)}{row_num}')
    if isinstance(value, (Decimal, int, float)) and not isinstance(value,bool):
        ET.SubElement(c,N+'v').text=str(value)
    else:
        text=re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]','',str(value))[:32000]
        c.set('t','inlineStr')
        ET.SubElement(ET.SubElement(c,N+'is'),N+'t').text=text


def save_emag_template(rows, output):
    if not rows:return None
    with ZipFile(TEMPLATE) as source:
        shared=_shared_strings(source)
        sheet=ET.fromstring(source.read(SHEET_PATH))
        row3=sheet.find('.//'+N+'sheetData/'+N+'row[@r="3"]')
        hdr=[_cell_value(c,shared) for c in row3.findall(N+'c')]
        if hdr != TEMPLATE_HEADER:raise ValueError('Unexpected template header — will not corrupt original format')
        data=sheet.find(N+'sheetData')
        for old in list(data):
            if int(old.get('r','0'))>=6:data.remove(old)
        # Preserve validation, defined names, other sheets, styles and all original format rules.
        for idx,values in enumerate(rows,6):
            node=ET.SubElement(data,N+'row',r=str(idx),spans='1:52')
            for c,field in enumerate(TEMPLATE_HEADER[:49],1):
                _write_cell(node,c,idx,values.get(field))
        dimension=sheet.find(N+'dimension')
        if dimension is not None: dimension.set('ref',f'A1:AZ{5+len(rows)}')
        new_xml=ET.tostring(sheet,xml_declaration=True,encoding='UTF-8',standalone=True)
        output.parent.mkdir(parents=True,exist_ok=True)
        with ZipFile(output,'w') as dest:
            for item in source.infolist():
                payload=new_xml if item.filename==SHEET_PATH else source.read(item.filename)
                dest.writestr(item,payload)
    return output


def _csv(path,rows,columns):
    with open(path,'w',newline='',encoding='utf-8-sig') as f:
        writer=csv.DictWriter(f,fieldnames=columns,extrasaction='ignore',delimiter=';')
        writer.writeheader();writer.writerows(rows)


def main():
    OUTPUT.mkdir(exist_ok=True)
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s',
                        handlers=[logging.StreamHandler(),logging.FileHandler(OUTPUT/'emag_create.log',encoding='utf8')])
    cfg=json.loads(CONFIG.read_text(encoding='utf8'))
    allowed=template_allowed()
    urls,sitemap_errors=fc.discover()
    logging.info('%d URLs discovered in product sitemap(s)',len(urls))
    rows,errors=[],list(sitemap_errors)
    with ThreadPoolExecutor(max_workers=fc.WORKERS) as pool:
        future_map={pool.submit(fc.fetch,url):url for url in urls}
        for n,future in enumerate(as_completed(future_map),1):
            url=future_map[future]
            try:rows.append(fc.parse_page(future.result(),url))
            except Exception as exc:errors.append((url,str(exc)));logging.warning('%s: %s',url,exc)
            if n%50==0:logging.info('Processed %s/%s, errors %s',n,len(urls),len(errors))
    rows.sort(key=lambda r:(r.get('Категория APIS',''),r.get('Име','')))
    fc.write_excel(rows,errors)
    candidates,ready,audits,excluded=[],[],[],[]
    for row in rows:
        result,reason=build_offer(row,cfg,allowed)
        if result is None:
            excluded.append({'URL':row.get('URL продукт',''),'Име':row.get('Име',''),
                             'Категория APIS':row.get('Категория APIS',''),'Причина':reason})
            continue
        offer,audit=result
        candidates.append(offer);audits.append(audit)
        if audit['Статус']=='ГОТОВ ЗА ПРОВЕРКА':ready.append(offer)
    save_emag_template(candidates,OUTPUT/'eMAG_school_accessories_DRAFT_DO_NOT_UPLOAD.xlsx')
    if ready:save_emag_template(ready,OUTPUT/'eMAG_school_accessories_READY_review_before_upload.xlsx')
    _csv(OUTPUT/'school_accessories_review.csv',audits,['URL','SKU','ID','Име','Тип продукт','Статус','Причина за включване','Проблеми','Цена EUR с ДДС (APIS)','Изчислена цена EUR без ДДС','Наличност на сайта','Зададено количество','Всички оригинални снимки'])
    _csv(OUTPUT/'other_categories_need_different_template.csv',excluded,['URL','Име','Категория APIS','Причина'])
    _csv(OUTPUT/'scrape_errors.csv',[{'URL':x,'Грешка':y} for x,y in errors],['URL','Грешка'])
    summary={'sitemap_urls':len(urls),'extracted':len(rows),'sitemap_and_product_errors':len(errors),
             'possible_school_accessories':len(candidates),'ready_for_seller_review':len(ready),
             'other_categories':len(excluded),'error_examples':errors[:10],
             'note':'eMAG single-category template cannot accept every APIS category; seller must validate legal fields, VAT, stock, brand, images and category.'}
    (OUTPUT/'coverage.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf8')
    logging.info('Summary: %s',summary)
    if not rows:raise RuntimeError('No products extracted; see logs')
    if sitemap_errors or len(errors)>max(10,len(urls)*0.02):
        raise RuntimeError('Partial scrape — check reports; artifacts are still uploaded by workflow')

if __name__=='__main__':main()
