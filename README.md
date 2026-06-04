# ApisExpress scraper

Събира продукти от категорията **Ученически материали** в ApisExpress и генерира Excel файл с колоните:

- Код на продукт
- Име на продукт
- Категория
- Цена в евро
- URL на всички изображения
- Описание на продукта
- URL на продукта

## Автоматично изпълнение без локален компютър

### Вариант A: GitHub Actions

1. Създай private GitHub repository.
2. Качи тези файлове в repository-то.
3. Отиди в **Actions → Scrape ApisExpress catalog → Run workflow**.
4. Скриптът ще се изпълнява автоматично всеки ден по cron и ще качва `apisexpress_products.xlsx` като artifact.

### Вариант B: VPS + Docker

```bash
git clone <your-repo-url>
cd apisexpress_scraper
docker compose up --build
```

За автоматично ежедневно изпълнение на VPS:

```bash
crontab -e
```

Добави:

```cron
0 2 * * * cd /path/to/apisexpress_scraper && docker compose up --build --abort-on-container-exit
```

Файлът ще бъде записан в `./data/apisexpress_products.xlsx`.

## Ръчно стартиране

```bash
pip install -r requirements.txt
python scraper.py --output apisexpress_products.xlsx
```

## Настройки

- Само HTML scraping:

```bash
python scraper.py --html-only --output apisexpress_products.xlsx
```

- Ограничение до първите 2 страници за тест:

```bash
python scraper.py --html-only --max-pages 2 --output test.xlsx
```
