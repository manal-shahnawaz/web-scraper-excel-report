# Competitor Price Tracker

A Python tool that scrapes product prices from an online catalogue, compares them with the previous run, and generates a formatted Excel report, so nobody has to copy-paste prices by hand.

> Built as a practice project against [books.toscrape.com](https://books.toscrape.com), a site made for scraping practice. The same approach works for competitor price monitoring, product research, or lead lists.

## What it does

- Scrapes **all pages** automatically (1,000 products) with retries and polite delays
- Cleans the data: prices become numbers, star ratings become 1-5, stock status is captured
- **Detects price changes** compared with the previous run
- Finds **top deals** (cheap and highly rated)
- Exports a formatted Excel workbook with 4 sheets:

| Sheet | Contents |
|---|---|
| Summary | Key statistics (Excel formulas) and a rating chart |
| All Books | Every product, with filters and clickable links |
| Price Changes | Old vs new price. Red = price went up, green = price went down |
| Top Deals | Products under a chosen price with 4-5 star ratings |

## Setup

```bash
git clone https://github.com/manal-shahnawaz/web-scraper-excel-report.git
cd web-scraper-excel-report
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

```bash
python price_tracker.py                    # full run (all 50 pages)
python price_tracker.py --pages 3          # quick test
python price_tracker.py --deal-price 15    # change the "top deal" price limit
python price_tracker.py --demo             # simulate price changes for a demo
```

The report is saved to `output/price_report.xlsx`.

Run it again on another day and the **Price Changes** sheet shows what moved.
`--demo` mode invents "previous prices" so you can see the feature working immediately; the report labels itself as demo data.

## Adapting it to a real website

Only `parse_books()` in `price_tracker.py` is site-specific. Change the CSS selectors there to match the new site's HTML, and the rest (pagination, comparison, Excel report) keeps working.

## Responsible scraping

Only scrape public data, read the website's terms of service, and keep a delay between requests.

## Tech stack

Python, requests, BeautifulSoup, pandas, openpyxl
