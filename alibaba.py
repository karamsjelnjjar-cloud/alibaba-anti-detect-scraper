"""
Script : Alibaba Kitchen Tools Scraper (All Pages) — Styled Version
Target : https://www.alibaba.com — Search: "ادوات مطبخ" (Kitchen Tools)
Output : alibaba_kitchen_tools.xlsx

Note   : Alibaba protection قوية — إذا ظهر Captcha سيتوقف المتصفح لحلّه يدوياً
         (لهذا HEADLESS = False افتراضياً). غيّره إلى True فقط للتجربة السريعة.
"""

import asyncio

import pandas as pd
from playwright.async_api import async_playwright


# ==========================================
# Global settings
# ==========================================
HEADLESS = False                     # Alibaba حمايته قوية — الأفضل متصفح ظاهر لحل الكابتشا
MAX_PAGES = None                     # None = جميع الصفحات تلقائياً | وضع رقم لتحديد حد أقصى
PAGE_DELAY = 3                       # استراحة بين الصفحات (ثواني)
NAV_TIMEOUT = 60000                  # مهلة تحميل الصفحة
CARD_WAIT = 30000                    # مهلة انتظار بطاقات المنتجات
CAPTCHA_WAIT = 120000                # مهلة حل الكابتشا يدوياً (ثانيتان)
SCROLL_STEP = 400                    # خطوة التمرير (بكسل)

# روابط البحث: "ادوات مطبخ" مشفّرة في الرابط
SEARCH_URL = (
    "https://www.alibaba.com/search/page?"
    "SearchScene=proSearch&pro=true"
    "&SearchText=%D8%A7%D8%AF%D9%88%D8%A7%D8%AA+%D9%85%D8%B7%D8%A8%D8%AE"
)

# المحّددات الأساسية المستخرجة من DevTools
PRODUCT_LINK_SELECTOR = "a.searchx-product-e-slider__link[href*='/product-detail/']"
PAGINATION_NUMBERS_JS = r"""
    () => {
        const nums = [];
        document.querySelectorAll("div.searchx-pagination-list span.pagination-item").forEach(s => {
            let n = parseInt((s.textContent || "").trim(), 10);
            if (isNaN(n)) {
                const m = (s.getAttribute("aria-label") || "").match(/(\d+)/);
                if (m) n = parseInt(m[1], 10);
            }
            if (!isNaN(n)) nums.push(n);
        });
        return nums;
    }
"""

all_products = []
seen_urls = set()


# ==========================================
# 1. إطلاق المتصفح بإعدادات مموّهة
# ==========================================
async def launch_browser(p):
    print("🚀 Launching browser (anti-detect mode)...")

    browser = await p.chromium.launch(
        headless=HEADLESS,
        args=[
            "--disable-blink-features=AutomationControlled",  # إخفاء آثار الأتمتة
            "--use-gl=angle",                                  # محاكاة GPU حقيقي
            "--disable-dev-shm-usage",
            "--lang=en-US",
        ],
    )
    context = await browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        viewport={"width": 1366, "height": 768},
        locale="en-US",
        timezone_id="Asia/Riyadh",
    )
    return browser, context


# ==========================================
# 2. كشف فحص Alibaba الأمني (Captcha)
# ==========================================
async def handle_security_check(page):
    content = (await page.content()).lower()

    if not any(k in content for k in ["captcha", "robot check", "verify you are", "slide to verify"]):
        return True

    print("⚠️ Alibaba security check detected!")

    if HEADLESS:
        print("❌ Captcha cannot be solved in headless mode — set HEADLESS = False")
        return False

    print("🧩 Please solve the captcha manually in the browser (120s)...")

    try:
        await page.wait_for_selector(PRODUCT_LINK_SELECTOR, timeout=CAPTCHA_WAIT)
        print("✅ Security check passed!")
        return True
    except Exception:
        print("❌ Captcha was not solved in time.")
        return False


# ==========================================
# 3. تمرير بشري لتحميل المحتوى الكسول (Lazy loading)
# ==========================================
async def human_scroll(page):
    await page.evaluate(
        """async (step) => {
            await new Promise(resolve => {
                let total = 0;
                const timer = setInterval(() => {
                    window.scrollBy(0, step);
                    total += step;
                    if (total >= document.body.scrollHeight) {
                        clearInterval(timer);
                        resolve();
                    }
                }, 250);
            });
            window.scrollTo(0, 0);
        }""",
        SCROLL_STEP,
    )
    await page.wait_for_timeout(1000)


# ==========================================
# 4. كشف العدد الكلي للصفحات من شريط الترقيم
# ==========================================
async def get_total_pages(page, quiet=False):
    try:
        nums = await page.evaluate(PAGINATION_NUMBERS_JS)

        if nums:
            total = max(nums)
            if not quiet:
                print(f"📄 Pagination detected — visible pages up to: {total}")
            return total
    except Exception as e:
        print(f"⚠️ Could not read pagination: {e}")

    return None


# ==========================================
# 5. استخراج بطاقات المنتجات من الصفحة الحالية
# ==========================================
async def extract_products(page, page_num):
    cards = await page.evaluate(
        r"""() => {
            const results = [];
            const links = [...document.querySelectorAll("a.searchx-product-e-slider__link[href*='/product-detail/']")];
            const clean = (s) => (s || "").replace(/\s+/g, " ").trim();

            for (const a of links) {
                const href = a.href || "";
                if (!href) continue;

                // ── البطاقة الأم: نصعد مستوى مستوى حتى نجد حاوية فيها سعر ──
                // (أقرب عنصر باسم searchx-product هو حاوية الصورة فقط — لا نعتمد عليه)
                let card = null;
                let node = a.parentElement;
                for (let i = 0; i < 10 && node; i++) {
                    if (node.querySelector && node.querySelector("[class*='price']")) { card = node; break; }
                    node = node.parentElement;
                }
                if (!card) { card = a.parentElement; }

                const img = a.querySelector("img") || card.querySelector("img");

                // ── العنوان: عدة مصادر بالترتيب (Lazy loading قد يؤخر img.alt) ──
                const titleAnchors = [...card.querySelectorAll("a[href*='/product-detail/']")];
                const titleLink = titleAnchors.find(x => clean(x.textContent).length > 15);
                let title = "";
                if (img && clean(img.alt).length > 5) title = img.alt;
                else if (clean(a.getAttribute("aria-label")).length > 5) title = a.getAttribute("aria-label");
                else if (titleLink) title = titleLink.textContent;
                else {
                    const t = card.querySelector("h2, h3, [class*='title'], [class*='subject']");
                    if (t) title = t.textContent;
                }

                // ── السعر / المورّد / الموقع ──
                const priceEl = card.querySelector("[class*='price']");
                const price    = priceEl ? priceEl.textContent : "";
                const supplier = (card.querySelector("[class*='company'], [class*='supplier']") || {}).textContent || "";
                const location = (card.querySelector("[class*='location'], [class*='country'], [class*='region']") || {}).textContent || "";

                // ── MOQ من نص البطاقة كاملاً ──
                const cardText = card.innerText || "";
                const moqMatch = cardText.match(/(?:MOQ|Min\.?\s*[Oo]rder|Minimum\s+[Oo]rder)[^\d\n]*(\d[\d,]*)/);

                results.push({
                    title:  clean(title),
                    price:  clean(price),
                    moq:    moqMatch ? moqMatch[1] : "",
                    supplier: clean(supplier),
                    location: clean(location),
                    url:    href,
                    image:  img ? clean(img.currentSrc || img.src || img.getAttribute("data-src") || "") : ""
                });
            }
            return results;
        }"""
    )

    added = 0

    for c in cards:
        if c["url"] in seen_urls:
            continue

        seen_urls.add(c["url"])
        all_products.append({
            "Product Name": c["title"],
            "Price":        c["price"],
            "MOQ":          c["moq"],
            "Supplier":     c["supplier"],
            "Location":     c["location"],
            "Product URL":  c["url"],
            "Image URL":    c["image"],
            "Page":         page_num,
        })
        added += 1

    print(f"  📝 Page {page_num}: extracted {len(cards)} cards ({added} new, {len(all_products)} total)")


# ==========================================
# 6. انتظار تحميل الصور والبطاقات (Lazy loading)
# ==========================================
async def wait_for_cards_ready(page, timeout=20000):
    try:
        await page.wait_for_function(
            """() => {
                const imgs = [...document.querySelectorAll("a.searchx-product-e-slider__link img")];
                return imgs.length > 0 && imgs.some(i => (i.currentSrc || i.src || '').startsWith('http'));
            }""",
            timeout=timeout,
        )
    except Exception:
        pass  # نكمل حتى لو انتهت المهلة — الأفضل جزئي من لا شيء


# ==========================================
# 7. جلب روابط المنتجات الحالية (للكشف عن تغيّر الصفحة)
# ==========================================
async def get_product_urls(page):
    return set(
        await page.evaluate(
            """() => [...new Set([...document.querySelectorAll("a[href*='/product-detail/']")].map(a => a.href))]"""
        )
    )


# ==========================================
# 8. الانتقال إلى الصفحة التالية (ضغط رقم الصفحة أو زر Next)
# ==========================================
async def goto_page(page, page_num):
    # المحاولة الأولى: الضغط على رقم الصفحة المطلوب مباشرة
    found = await page.evaluate(
        """(target) => {
            const items = [...document.querySelectorAll("div.searchx-pagination-list a")];
            for (const a of items) {
                const span = a.querySelector("span.pagination-item");
                if (span && span.textContent.trim() === target) { a.click(); return true; }
            }
            return false;
        }""",
        str(page_num),
    )

    if found:
        return True

    # المحاولة الثانية (احتياطي): زر السهم التالي (>)
    return await page.evaluate(
        r"""() => {
            const pager = document.querySelector('div.searchx-pagination');
            if (!pager) return false;

            const btns = [...pager.querySelectorAll('button, a')];
            const next = btns.find(b =>
                /next|التالي|›|»/i.test((b.getAttribute('aria-label') || '') + ' ' + b.textContent)
            );

            if (next && next.getAttribute('aria-disabled') !== 'true' && !next.disabled) {
                next.click();
                return true;
            }
            return false;
        }"""
    )


# ==========================================
# 9. انتظار تحدّث محتوى الصفحة بعد الضغط
# ==========================================
async def wait_for_page_change(page, prev_urls, timeout=25000):
    await page.wait_for_timeout(1500)

    for _ in range(timeout // 1000):
        try:
            urls = await get_product_urls(page)
            new_urls = [u for u in urls if u not in prev_urls]

            if new_urls:
                return True
        except Exception:
            pass

        await page.wait_for_timeout(1000)

    return False


# ==========================================
# 10. المرور على جميع الصفحات بالترتيب
# ==========================================
async def scrape_all_pages(context):
    page = await context.new_page()
    current = 1
    total_pages = None

    try:
        print(f"\n🔍 Loading search results: ادوات مطبخ")
        await page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
        await human_scroll(page)

        if not await handle_security_check(page):
            return

        await page.wait_for_selector(PRODUCT_LINK_SELECTOR, timeout=CARD_WAIT)
        total_pages = await get_total_pages(page)

        while True:
            if MAX_PAGES and current > MAX_PAGES:
                print(f"\n🏁 Reached MAX_PAGES limit ({MAX_PAGES}).")
                break

            print(f"\n📦 Scraping page {current} ...")
            await human_scroll(page)
            await wait_for_cards_ready(page)
            await extract_products(page, current)

            # التوقف عند الوصول لآخر صفحة معروفة
            if total_pages and current >= total_pages:
                print(f"\n🏁 Reached last page ({total_pages}). Done!")
                break

            prev_urls = await get_product_urls(page)

            if not await goto_page(page, current + 1):
                print("\n🏁 Next page button unavailable — stopping.")
                break

            if not await wait_for_page_change(page, prev_urls):
                print("\n🏁 Page content did not change — stopping.")
                break

            # تحديث العدد الكلي (نافذة أرقام الترقيم تتحرك للأمام مع التقدم)
            latest = await get_total_pages(page, quiet=True)
            if latest:
                total_pages = max(total_pages or 1, latest)

            current += 1
            await asyncio.sleep(PAGE_DELAY)

    except Exception as e:
        print(f"❌ Unexpected error: {e}")

    finally:
        await page.close()


# ==========================================
# 11. تصدير البيانات إلى Excel
# ==========================================
def export_to_excel():
    if not all_products:
        print("❌ No data extracted!")
        return

    df = pd.DataFrame(all_products)
    df = df.drop_duplicates(subset=["Product URL"])

    filename = "alibaba_kitchen_tools.xlsx"
    df.to_excel(filename, index=False)

    print(f"\n🎉 File saved successfully: {filename}")
    print(f"📊 Total products extracted: {len(df)}")


# ==========================================
# (تشخيصي) حفظ عينة من بطاقات المنتجات HTML للفحص
# ==========================================
async def dump_sample_cards(page, n=3, filename="alibaba_sample_cards.html"):
    html = await page.evaluate(
        """(n) => {
            const links = [...document.querySelectorAll("a.searchx-product-e-slider__link[href*='/product-detail/']")].slice(0, n);
            return links.map(a => {
                let node = a.parentElement;
                for (let i = 0; i < 10 && node; i++) {
                    if (node.querySelector && node.querySelector("[class*='price']")) return node.outerHTML;
                    node = node.parentElement;
                }
                return a.outerHTML;
            }).join("\n\n<!-- ===== CARD ===== -->\n\n");
        }""",
        n,
    )

    with open(filename, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"🧪 Sample cards saved to: {filename}")


# ==========================================
# 12. Main Execution
# ==========================================
async def main():
    async with async_playwright() as p:
        browser, context = await launch_browser(p)
        await scrape_all_pages(context)
        await browser.close()

    export_to_excel()


if __name__ == "__main__":
    asyncio.run(main())
