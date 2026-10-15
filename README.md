# 🛒 Alibaba Stealth & Anti-Detect Product Scraper

An advanced, production-grade web scraping automation tool built with **Python**, **Playwright (Async API)**, and **Pandas**. Designed specifically to scrape thousands of product listings from **Alibaba.com** while bypassing aggressive anti-bot detection systems, handling human-mimicked smooth scrolling for dynamic lazy-loading content, and automating pagination with manual Captcha fallback mechanisms.

## 🚀 Key Features
- **Anti-Detect Browser Launch (`Playwright`):** Bypasses bot detection using stealth browser flags (`--disable-blink-features=AutomationControlled`), realistic user agents, customized viewports, and regional localization.
- **Smart Captcha Detection & Manual Fallback:** Continuously monitors for security verification prompts (Captchas) and pauses execution safely to allow human intervention without breaking the pipeline.
- **Human-Mimicked Smooth Scrolling:** Implements an automated human-like scrolling loop to trigger dynamic image and element lazy-loading.
- **Dynamic Pagination & Page Tracking:** Automatically evaluates pagination bars and handles multi-page navigation across hundreds of results.
- **Data Clean-up & Export:** Extracts titles, prices, MOQs, supplier details, image URLs, and product links, deduplicates them, and exports everything into a styled Excel file (`alibaba_kitchen_tools.xlsx`).

## 🛠️ Built With
- **Python 3.x**
- **Playwright (Async API)** (Stealth Browser Automation)
- **Asyncio** (Asynchronous Flow Management)
- **Pandas & OpenPyXL** (Data Deduplication, Processing & Excel Export)

## ⚙️ How to Run
1. Clone the repository:
   ```bash
   git clone [https://github.com/karamsjelnjjar-cloud/alibaba-anti-detect-scraper.git](https://github.com/karamsjelnjjar-cloud/alibaba-anti-detect-scraper.git)
