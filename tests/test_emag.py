import unittest
import os
import tempfile
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as ET
import emag_create as e

class TestEMAG(unittest.TestCase):
    @classmethod
    def setUpClass(cls): cls.allowed=e.template_allowed()
    def test_template_matches_expected_headers(self):
        self.assertEqual(len(e.TEMPLATE_HEADER),52)
        self.assertIn('Бутилка за вода',self.allowed['A'])
        self.assertEqual(self.allowed['B'],{'Момичета','Момчета'})
    def test_ean_validation(self):
        self.assertTrue(e.ean_valid('4006381333931'))
        self.assertFalse(e.ean_valid('4006381333932'))
    def test_vat_rounding(self):
        self.assertEqual(str(e.net_price('1.28','0.2')),'1.0667')
    def test_child_price_and_template_preservation(self):
        offer={'part_number':'SKU-TEST','vendor_ext_id':'33920','name':'Примерен продукт',
               'offer_currency':'EUR','status':0,'stock':0,'sale_price':e.net_price('1.28','0.2'),
               'main_image_url':'https://example.com/a.jpg'}
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'draft.xlsx';e.save_emag_template([offer],path)
            with ZipFile(e.TEMPLATE) as original, ZipFile(path) as saved:
                self.assertEqual(set(original.namelist()),set(saved.namelist()))
                self.assertEqual(original.read('xl/worksheets/sheet5.xml'),saved.read('xl/worksheets/sheet5.xml'))
                xml=ET.fromstring(saved.read(e.SHEET_PATH))
                ns={'m':e.XL}
                vals={n.get('r'):''.join(n.itertext()) for n in xml.xpath('//m:row[@r="6"]/m:c',namespaces=ns)}
                self.assertEqual(vals['A6'],'SKU-TEST')
                self.assertEqual(vals['D6'],'1.0667')
                self.assertEqual(vals['U6'],'https://example.com/a.jpg')
                self.assertEqual(len(xml.xpath('//m:dataValidation',namespaces=ns)),5)
    def test_other_categories_excluded(self):
        row={'Име':'Химикал син','Категория APIS':'Писалки','URL продукт':'https://example.com/p'}
        result,reason=e.classify(row,{},self.allowed)
        self.assertIsNone(result)
    def test_no_manufacturer_or_stock_fabrication(self):
        row={'URL продукт':'https://apisexpress.com/produkt/etiketi-za-tetradki/',
          'ID в изходния сайт':'33920','SKU':'123-TEST','Име':'Училищни етикети',
          'Категория APIS':'Ученически материали','Марка':'', 'Описание':'Описание',
          'Цена EUR с ДДС (сайт)':1.28,'Стара цена EUR с ДДС':None,
          'Наличност (статус)':'наличен','Вариативен продукт':'Не',
          'Изображение 1':'https://apisexpress.com/wp-content/uploads/etiketi.jpg'}
        result,reason=e.build_offer(row,{'vat_rate':0.2,'vat_confirmed':False},self.allowed)
        offer,audit=result
        self.assertEqual(offer['manufacturer_name'],'')
        self.assertIsNone(offer['sale_price'])
        self.assertIsNone(offer['stock'])
        self.assertIn('НЕПЪЛЕН',audit['Статус'])
        self.assertEqual(offer['status'],0)

if __name__=='__main__':unittest.main()
