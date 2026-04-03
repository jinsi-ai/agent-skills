---
slug: smart-web-scraper
version: 1.0.1
displayName: 智能网页抓取器（Smart Web Scraper）
summary: 从任何网页提取结构化数据，支持 CSS 选择器、自动检测表格和列表，并输出 JSON/CSV 格式。
tags: clawhub
---

# Smart Web Scraper

Extract structured data from web pages into clean JSON or CSV.

## Quick Start

```bash
# Scrape a page, extract all text content
uv run --with beautifulsoup4 --with lxml python scripts/scraper.py extract "https://example.com"

# Extract specific elements with CSS selector
uv run --with beautifulsoup4 --with lxml python scripts/scraper.py extract "https://example.com/products" -s ".product-card"

# Auto-detect and extract tables
uv run --with beautifulsoup4 --with lxml python scripts/scraper.py tables "https://example.com/pricing"

# Extract all links from a page
uv run --with beautifulsoup4 --with lxml python scripts/scraper.py links "https://example.com"

# Extract structured data (title, meta, headings, links)
uv run --with beautifulsoup4 --with lxml python scripts/scraper.py structure "https://example.com"

# Output as JSON
uv run --with beautifulsoup4 --with lxml python scripts/scraper.py extract "https://example.com" -s ".item" -f json

# Output as CSV
uv run --with beautifulsoup4 --with lxml python scripts/scraper.py extract "https://example.com" -s "table tr" -f csv

# Save to file
uv run --with beautifulsoup4 --with lxml python scripts/scraper.py extract "https://example.com" -s ".product" -f json -o products.json

# Multi-page scrape (follow pagination)
uv run --with beautifulsoup4 --with lxml python scripts/scraper.py crawl "https://example.com/page/1" --pages 5 -s ".article"
```

## Commands

| Command | Args | Description |
|---------|------|-------------|
| `extract` | `<url> [-s selector] [-f format] [-o file]` | Extract content, optionally filtered by CSS selector |
| `tables` | `<url> [-f format] [-o file]` | Auto-detect and extract all HTML tables |
| `links` | `<url> [--external] [--internal]` | Extract all links (href + text) |
| `structure` | `<url>` | Extract page structure: title, meta, headings, images, links |
| `crawl` | `<url> --pages N [-s selector] [-f format] [-o file]` | Follow pagination links, extract from multiple pages |

## Output Formats

| Format | Flag | Description |
|--------|------|-------------|
| Text | `-f text` | Plain text (default) |
| JSON | `-f json` | Structured JSON array |
| CSV | `-f csv` | Comma-separated values |
| Markdown | `-f md` | Markdown-formatted |

## Examples

### Extract product listings
```bash
uv run --with beautifulsoup4 --with lxml python scripts/scraper.py extract "https://shop.example.com" -s ".product" -f json
```
Output:
```json
[
  {"text": "Widget Pro - $29.99", "tag": "div", "class": "product"},
  {"text": "Widget Max - $49.99", "tag": "div", "class": "product"}
]
```

### Extract pricing table
```bash
uv run --with beautifulsoup4 --with lxml python scripts/scraper.py tables "https://example.com/pricing" -f csv
```

### Get all external links
```bash
uv run --with beautifulsoup4 --with lxml python scripts/scraper.py links "https://example.com" --external
```

## Rate Limiting

- Default: 1 request per second (respectful crawling)
- Override with `--delay 0.5` (seconds between requests)
- Respects `robots.txt` by default (override with `--ignore-robots`)

## Notes

- Requires `beautifulsoup4` and `lxml` (auto-installed by `uv run --with`)
- Uses a standard browser User-Agent to avoid blocks
- Handles redirects, encoding detection, and error pages gracefully
- No JavaScript rendering (use for static HTML pages)
