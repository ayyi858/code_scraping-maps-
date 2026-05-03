import asyncio, random, time
import pandas as pd
from playwright.async_api import async_playwright

KEYWORDS = [
    "rumah makan Makassar",
    "toko retail Makassar",
    "minimarket Gowa",
    "warung makan Bone",
    "toko kelontong Parepare",
]

async def scrape_gmaps(keyword: str, max_results: int = 50):
    results = []
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=["--disable-blink-features=AutomationControlled"]
        )
        context = await browser.new_context(
            viewport={"width": 1280, "height": 800},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/124.0.0.0 Safari/537.36",
            locale="id-ID",
        )
        page = await context.new_page()

        # FIX 1: ganti networkidle → domcontentloaded + timeout lebih panjang
        url = f"https://www.google.com/maps/search/{keyword.replace(' ', '+')}"
        try:
            await page.goto(url, wait_until="domcontentloaded", timeout=60000)
        except Exception as e:
            print(f"Gagal load {keyword}: {e}")
            await browser.close()
            return results

        # FIX 2: tunggu feed muncul dulu sebelum lanjut
        try:
            await page.wait_for_selector('div[role="feed"]', timeout=15000)
        except:
            print(f"Feed tidak muncul untuk: {keyword}")
            await browser.close()
            return results

        await asyncio.sleep(random.uniform(2, 3))

        # Scroll untuk load lebih banyak hasil
        feed = page.locator('div[role="feed"]')
        prev_count = 0
        for _ in range(15):
            await feed.evaluate("el => el.scrollBy(0, 800)")
            await asyncio.sleep(random.uniform(1.5, 2.5))
            items = page.locator('div[role="feed"] > div > div > a')
            count = await items.count()
            if count >= max_results or count == prev_count:
                break
            prev_count = count

        # Klik tiap item dan ambil detail
        items = page.locator('div[role="feed"] > div > div > a')
        count = min(await items.count(), max_results)
        print(f"Ditemukan {count} tempat untuk: {keyword}")

        for i in range(count):
            try:
                await items.nth(i).click()

                # FIX 3: tunggu panel detail terbuka
                await page.wait_for_selector('h1.DUwDvf', timeout=10000)
                await asyncio.sleep(random.uniform(1.5, 2.5))

                # Ambil data dengan fallback aman
                async def get_text(selector, fallback="N/A"):
                    try:
                        el = page.locator(selector).first
                        if await el.count() > 0:
                            return (await el.text_content() or fallback).strip()
                    except:
                        pass
                    return fallback

                nama     = await get_text('h1.DUwDvf')
                kategori = await get_text('button.DkEaL')
                rating   = await get_text('div.F7nice span[aria-hidden="true"]')
                ulasan   = await get_text('div.F7nice span[aria-label*="ulasan"]')
                alamat   = await get_text('button[data-item-id="address"] .fontBodyMedium')
                telp     = await get_text('button[data-item-id^="phone"] .fontBodyMedium')

                # Ambil koordinat dari URL
                lat, lng = "N/A", "N/A"
                current_url = page.url
                if "@" in current_url:
                    parts = current_url.split("@")[1].split(",")
                    if len(parts) >= 2:
                        lat, lng = parts[0], parts[1]

                results.append({
                    "keyword": keyword,
                    "nama": nama,
                    "kategori": kategori,
                    "rating": rating,
                    "jumlah_ulasan": ulasan.replace("(","").replace(")","").strip(),
                    "alamat": alamat,
                    "telepon": telp,
                    "latitude": lat,
                    "longitude": lng,
                })
                print(f"  [{i+1}/{count}] {nama}")

            except Exception as e:
                print(f"  Skip item {i+1}: {e}")
                continue

        await browser.close()
    return results

async def main():
    all_data = []
    for kw in KEYWORDS:
        print(f"\nScraping: {kw}")
        data = await scrape_gmaps(kw, max_results=30)
        all_data.extend(data)
        jeda = random.uniform(5, 10)
        print(f"Jeda {jeda:.0f} detik sebelum keyword berikutnya...")
        await asyncio.sleep(jeda)

    df = pd.DataFrame(all_data)
    df.drop_duplicates(subset=["nama", "alamat"], inplace=True)
    df.to_csv("data_usaha_sulsel.csv", index=False, encoding="utf-8-sig")
    print(f"\nSelesai! Total data: {len(df)} baris")
    print("Tersimpan di: data_usaha_sulsel.csv")

asyncio.run(main())