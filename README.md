# APIS Express → eMAG offer updates

This updates ONLY the existing offer rows in `template/emag_template.xlsx` (16 in the supplied file). It does not create offers for other APIS Express products; that requires a different eMAG listing template and matching account identifiers.

## GitHub setup
1. Upload **the contents** of this ZIP into your GitHub repository, preserving `.github/workflows/scrape.yml` and `template/emag_template.xlsx` directories.
2. Open **Actions → Update APIS Express eMAG offers → Run workflow**.
3. Download artifact `apisexpress-emag-update` when completed. The workbook is `emag_price_stock_update.xlsx`; check `matching_report.csv` and `scrape_errors.csv` before importing.
4. To update a different set of offers, replace `template/emag_template.xlsx` with a fresh export from your eMAG seller account.

Runs manually and weekly on Mondays at 04:00 UTC. GitHub scheduling can be delayed.

## Safety and price rules
- Match by SKU or EAN only; never assume name similarity is a sufficient match.
- Use explicitly EUR WooCommerce product summary price (discounted `ins` if present); compare with Store API if available. If values conflict, leave original price unchanged and flag in report.
- Update stock only if the site exposes a precise numeric quantity or explicitly says out of stock. Generic "in stock" is not a numeric quantity; preserve template stock in that case.
- Preserve eMAG identifiers, VAT, offer status, currency, and workbook structure.
- This script does not automatically upload to eMAG.
- Site selectors and Store API availability must be validated against current live pages; inspect matching_report before importing.
