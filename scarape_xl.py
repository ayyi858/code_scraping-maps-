import asyncio, random, os
import pandas as pd
from datetime import datetime
from playwright.async_api import async_playwright

KECAMATAN = [
    "Mariso", "Mamajang", "Tamalate", "Rappocini", "Makassar",
    "Ujung Pandang", "Wajo", "Bontoala", "Ujung Tanah", "Tallo",
    "Panakkukang", "Manggala", "Biringkanaya", "Tamalanrea",
    "Kepulauan Sangkarrang"
]

KATEGORI = [
    "cafe",
    "coffee shop",
    "kedai kopi",
    "warung kopi",
    "kopi susu",
    "kopi hitam",
    "angkringan kopi",
    "espresso bar",
    "specialty coffee",
    "kafe aesthetic",
]

KEYWORDS = [f"{kat} {kec} Makassar" for kec in KECAMATAN for kat in KATEGORI]
print(f"Total keyword: {len(KEYWORDS)}")
# Output: Total keyword: 150

OUTPUT_FILE = "datarev_cafemakassar.xlsx"
CHECKPOINT_FILE = "checkpoint_progress.txt"

def load_checkpoint():
    if os.path.exists(CHECKPOINT_FILE):
        with open(CHECKPOINT_FILE) as f:
            return int(f.read().strip())
    return 0

def save_checkpoint(idx):
    with open(CHECKPOINT_FILE, "w") as f:
        f.write(str(idx))

def save_to_excel(all_data):
    df = pd.DataFrame(all_data)
    df.drop_duplicates(subset=["nama", "alamat"], inplace=True)
    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print(f"Tersimpan: {len(df)} baris → {OUTPUT_FILE}")
    return df

async def scrape_gmaps(page, keyword: str, max_results: int = 60):
    results = []
    url = f"https://www.google.com/maps/search/{keyword.replace(' ', '+')}"

    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=60000)
        await page.wait_for_selector('div[role="feed"]', timeout=15000)
    except:
        print(f"  Gagal load: {keyword}")
        return results

    await asyncio.sleep(random.uniform(2, 3))

    # Scroll untuk load semua hasil
    feed = page.locator('div[role="feed"]')
    prev_count = 0
    for _ in range(20):
        await feed.evaluate("el => el.scrollBy(0, 800)")
        await asyncio.sleep(random.uniform(1, 2))
        items = page.locator('div[role="feed"] > div > div > a')
        count = await items.count()
        if count >= max_results or count == prev_count:
            break
        prev_count = count

    items = page.locator('div[role="feed"] > div > div > a')
    count = min(await items.count(), max_results)
    print(f"  Ditemukan {count} tempat")

    async def get_text(selector, fallback="N/A"):
        try:
            el = page.locator(selector).first
            if await el.count() > 0:
                return (await el.text_content() or fallback).strip()
        except:
            pass
        return fallback

    for i in range(count):
        try:
            await items.nth(i).click()
            await page.wait_for_selector('h1.DUwDvf', timeout=10000)
            await asyncio.sleep(random.uniform(1.2, 2))

            nama     = await get_text('h1.DUwDvf')
            kategori = await get_text('button.DkEaL')
            rating   = await get_text('div.F7nice span[aria-hidden="true"]')
            ulasan   = await get_text('div.F7nice span[aria-label*="ulasan"]')
            alamat   = await get_text('button[data-item-id="address"] .fontBodyMedium')
            telp     = await get_text('button[data-item-id^="phone"] .fontBodyMedium')
            website  = await get_text('a[data-item-id="authority"] .fontBodyMedium')

            lat, lng = "N/A", "N/A"
            current_url = page.url
            if "@" in current_url:
                parts = current_url.split("@")[1].split(",")
                if len(parts) >= 2:
                    lat, lng = parts[0], parts[1]

            results.append({
                "keyword": keyword,
                "nama": nama,
                "kategori_usaha": kategori,
                "rating": rating,
                "jumlah_ulasan": ulasan.replace("(","").replace(")","").strip(),
                "alamat": alamat,
                "telepon": telp,
                "website": website,
                "latitude": lat,
                "longitude": lng,
                "kecamatan": keyword.split(" ")[-2],
                "tanggal_scraping": datetime.now().strftime("%Y-%m-%d"),
            })
            print(f"    [{i+1}/{count}] {nama}")

        except Exception as e:
            print(f"    Skip {i+1}: {e}")
            continue

    return results

async def main():
    start_idx = load_checkpoint()
    all_data = []

    # Load data sebelumnya jika ada
    if os.path.exists(OUTPUT_FILE) and start_idx > 0:
        df_existing = pd.read_excel(OUTPUT_FILE, engine="openpyxl")
        all_data = df_existing.to_dict("records")
        print(f"Lanjut dari checkpoint {start_idx}, data sebelumnya: {len(all_data)} baris")

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

        remaining = KEYWORDS[start_idx:]
        for i, kw in enumerate(remaining, start=start_idx):
            print(f"\n[{i+1}/{len(KEYWORDS)}] Scraping: {kw}")
            data = await scrape_gmaps(page, kw, max_results=60)
            all_data.extend(data)

            # Simpan tiap 5 keyword (auto-save)
            if (i + 1) % 5 == 0:
                save_to_excel(all_data)
                save_checkpoint(i + 1)
                print(f"Auto-saved. Total sementara: {len(all_data)} baris")

            jeda = random.uniform(4, 8)
            print(f"Jeda {jeda:.0f} detik...")
            await asyncio.sleep(jeda)

        await browser.close()

    # Simpan final
    df_final = save_to_excel(all_data)
    if os.path.exists(CHECKPOINT_FILE):
        os.remove(CHECKPOINT_FILE)
    print(f"\nSELESAI! Total akhir: {len(df_final)} baris unik")
    print(f"File: {OUTPUT_FILE}")

asyncio.run(main())