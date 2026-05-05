"""
SCRAPER USAHA MAKASSAR — KOORDINAT DARI URL (FIXED)
=====================================================
Fix: tunggu URL stabil sebelum ambil koordinat
"""

import asyncio
import os
import random
import re
from datetime import datetime

import pandas as pd
from playwright.async_api import async_playwright

# ─────────────────────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────────────────────
KECAMATAN = [
    "Mariso", "Mamajang", "Tamalate", "Rappocini", "Makassar",
    "Ujung Pandang", "Wajo", "Bontoala", "Ujung Tanah", "Tallo",
    "Panakkukang", "Manggala", "Biringkanaya", "Tamalanrea",
    "Kepulauan Sangkarrang",
]

KATEGORI = [
    # ── KULINER & MINUMAN ─────────────────────────────────
    "cafe", "coffee shop", "kedai kopi", "warung makan",
    "restoran", "seafood", "bakso", "mie ayam",
    "coto makassar", "konro", "pallubasa", "nasi goreng",
    "nasi kuning", "soto", "sate", "ayam geprek", "ayam bakar",
    "martabak", "pisang epe", "es pisang ijo",
    "burger", "pizza", "korean food", "japanese food",
    "chinese food", "dessert", "es krim", "minuman kekinian",
    "boba", "jus buah", "roti dan bakeri", "donat",
    "mie gacoan", "warteg", "rumah makan padang",
    # ── RETAIL & TOKO ────────────────────────────────────
    "minimarket", "supermarket", "toko kelontong",
    "toko pakaian", "toko baju muslim", "toko sepatu",
    "toko tas", "toko elektronik", "toko hp", "toko komputer",
    "toko bangunan", "toko buku", "toko mainan",
    "toko kosmetik", "toko aksesoris", "toko oleh oleh",
    "toko peralatan rumah", "toko furniture",
    # ── KESEHATAN & KECANTIKAN ───────────────────────────
    "apotek", "klinik", "klinik gigi", "klinik kecantikan",
    "rumah sakit", "puskesmas", "optik",
    "salon kecantikan", "barbershop", "spa", "nail art", "tukang cukur",
    # ── JASA UMUM ────────────────────────────────────────
    "laundry", "laundry kiloan", "percetakan", "foto studio",
    "jasa servis hp", "jasa jahit", "jasa cuci ac",
    "jasa servis elektronik", "travel agent",
    "ekspedisi pengiriman", "jasa kebersihan", "notaris", "kantor pos",
    # ── OTOMOTIF ─────────────────────────────────────────
    "bengkel motor", "bengkel mobil", "cuci motor", "cuci mobil",
    "tambal ban", "toko ban", "toko sparepart motor",
    "toko sparepart mobil", "modifikasi motor",
    # ── PENDIDIKAN & HIBURAN ─────────────────────────────
    "bimbel", "les privat", "kursus bahasa inggris",
    "kursus komputer", "taman kanak kanak",
    "gym fitness", "futsal", "kolam renang", "bioskop", "karaoke",
    "toko alat olahraga", "toko alat musik",
    # ── PROPERTI & AKOMODASI ─────────────────────────────
    "hotel", "penginapan", "kost", "guest house",
    # ── MAKANAN KHAS SULAWESI ────────────────────────────
    "ikan bakar", "kapurung", "jalangkote",
    "sup konro", "mie titi", "es pallu butung",
]

KATEGORI        = list(dict.fromkeys(KATEGORI))
KEYWORDS        = [f"{kat} {kec} Makassar" for kec in KECAMATAN for kat in KATEGORI]
OUTPUT_FILE     = "data_usaha_makassar.xlsx"
CHECKPOINT_FILE = "checkpoint_usaha.txt"
MAX_HASIL       = 60
DELAY_ANTAR_KW  = (3, 7)

print(f"Total kategori : {len(KATEGORI)}")
print(f"Total keyword  : {len(KEYWORDS)}")
print(f"Estimasi waktu : {len(KEYWORDS)*5//60} - {len(KEYWORDS)*7//60} jam\n")


# ─────────────────────────────────────────────────────────────
# UTILITAS
# ─────────────────────────────────────────────────────────────
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
    print(f"  Tersimpan: {len(df)} baris -> {OUTPUT_FILE}")
    return df

def ekstrak_kecamatan(keyword, kategori):
    kw = keyword.strip()
    if kw.endswith(" Makassar"):
        kw = kw[:-9].strip()
    if kw.lower().startswith(kategori.lower()):
        kw = kw[len(kategori):].strip()
    return kw

async def safe_text(page, selector, fallback=""):
    try:
        el = page.locator(selector).first
        if await el.count() > 0:
            return (await el.text_content() or fallback).strip()
    except Exception:
        pass
    return fallback


# ─────────────────────────────────────────────────────────────
# AMBIL KOORDINAT DARI URL — FIXED
# ─────────────────────────────────────────────────────────────
async def ambil_koordinat(page, nama_usaha="") -> tuple:
    """
    Tunggu URL stabil dengan koordinat yang benar.

    Masalah sebelumnya: URL diambil terlalu cepat, masih
    menunjuk ke lokasi search bukan lokasi usaha.

    Fix: tunggu sampai URL mengandung '/place/' DAN koordinat
    valid dalam bounding box Makassar, dengan beberapa kali retry.
    """
    MAX_RETRY  = 20
    WAIT_STEP  = 0.4  # detik per retry

    for attempt in range(MAX_RETRY):
        try:
            url = page.url

            # URL harus sudah pindah ke halaman place
            # Format: /maps/place/NamaUsaha/@lat,lng,zoom
            if "/place/" not in url and "/maps/search/" in url:
                # Masih di halaman search, belum klik atau belum navigasi
                await asyncio.sleep(WAIT_STEP)
                continue

            m = re.search(r'@(-?\d+\.\d+),(-?\d+\.\d+)', url)
            if not m:
                await asyncio.sleep(WAIT_STEP)
                continue

            lat = float(m.group(1))
            lng = float(m.group(2))

            # Validasi bounding box Kota Makassar
            if not (-5.30 <= lat <= -5.00 and 119.20 <= lng <= 119.60):
                # Koordinat tidak di Makassar, tunggu URL update
                await asyncio.sleep(WAIT_STEP)
                continue

            # Koordinat valid — kembalikan
            return round(lat, 7), round(lng, 7)

        except Exception:
            await asyncio.sleep(WAIT_STEP)
            continue

    # Gagal setelah semua retry
    return None, None


# ─────────────────────────────────────────────────────────────
# AMBIL URL FOTO
# ─────────────────────────────────────────────────────────────
async def ambil_foto_url(page) -> str:
    try:
        foto = await page.evaluate("""() => {
            const selectors = [
                'button.aoRNLd img', 'div.aoRNLd img',
                'img.DaSXdd', 'div.RZ66Rb img',
                'div.section-hero-header-image img',
            ];
            for (const sel of selectors) {
                const el = document.querySelector(sel);
                if (el && el.src && (el.src.includes('googleusercontent') || el.src.includes('ggpht')))
                    return el.src;
            }
            const imgs = document.querySelectorAll('img[src*="googleusercontent"], img[src*="ggpht"]');
            for (const img of imgs) {
                const rect = img.getBoundingClientRect();
                if (img.src && rect.width > 80 && rect.height > 60) return img.src;
            }
            for (const img of imgs) {
                if (img.src && img.src.length > 60) return img.src;
            }
            return null;
        }""")

        if foto and ("googleusercontent" in foto or "ggpht" in foto):
            return re.sub(r'=w\d+.*$', '=w800-h600-k-no', foto)

        html = await page.content()
        m = re.search(
            r'(https://lh[0-9]\.googleusercontent\.com/p/[A-Za-z0-9_\-]+=w\d+[^\s"\'<>]*)',
            html
        )
        if m:
            return re.sub(r'=w\d+.*$', '=w800-h600-k-no', m.group(1))

    except Exception:
        pass
    return ""


# ─────────────────────────────────────────────────────────────
# SCRAPE SATU KEYWORD
# ─────────────────────────────────────────────────────────────
async def scrape_keyword(page, keyword, kecamatan, kategori_key, max_results=60):
    results = []
    url = f"https://www.google.com/maps/search/{keyword.replace(' ', '+')}"

    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=60000)
        await page.wait_for_selector('div[role="feed"]', timeout=15000)
    except Exception as e:
        print(f"  Gagal load: {e}")
        return results

    await asyncio.sleep(random.uniform(2, 3))

    # Scroll feed
    feed = page.locator('div[role="feed"]')
    prev = 0
    for _ in range(25):
        await feed.evaluate("el => el.scrollBy(0, 900)")
        await asyncio.sleep(random.uniform(0.8, 1.5))
        cur = await page.locator('div[role="feed"] > div > div > a').count()
        if cur >= max_results or cur == prev:
            break
        prev = cur

    items = page.locator('div[role="feed"] > div > div > a')
    total = min(await items.count(), max_results)
    print(f"  Ditemukan {total} tempat")

    for i in range(total):
        try:
            await items.nth(i).click()

            # Tunggu panel detail muncul
            await page.wait_for_selector("h1.DUwDvf", timeout=10000)

            # Tunggu URL berubah ke /place/
            try:
                await page.wait_for_url(re.compile(r'.*/place/.*'), timeout=5000)
            except Exception:
                pass

            await asyncio.sleep(random.uniform(1.2, 1.8))

            # Ambil semua data
            nama       = await safe_text(page, "h1.DUwDvf")
            kat_usaha  = await safe_text(page, "button.DkEaL")
            rating     = await safe_text(page, "div.F7nice span[aria-hidden='true']")
            ulasan_raw = await safe_text(page, "div.F7nice span[aria-label*='ulasan']")
            alamat     = await safe_text(page, "button[data-item-id='address'] .fontBodyMedium")
            telp       = await safe_text(page, "button[data-item-id^='phone'] .fontBodyMedium")
            website    = await safe_text(page, "a[data-item-id='authority'] .fontBodyMedium")
            foto_url   = await ambil_foto_url(page)

            # Koordinat dari URL — tunggu sampai stabil & valid
            lat, lng   = await ambil_koordinat(page, nama)

            ulasan = ulasan_raw.replace("(","").replace(")","").strip()

            if not nama:
                continue

            results.append({
                "keyword"       : keyword,
                "nama"          : nama,
                "kategori_usaha": kat_usaha,
                "rating"        : rating,
                "jumlah_ulasan" : ulasan,
                "alamat"        : alamat,
                "telepon"       : telp,
                "website"       : website,
                "kecamatan"     : kecamatan,
                "latitude"      : lat,
                "longitude"     : lng,
                "foto_url"      : foto_url,
                "tanggal"       : datetime.now().strftime("%Y-%m-%d"),
            })

            ikon_f = "F" if foto_url else "-"
            ikon_k = "K" if lat else "?"
            print(f"    [{i+1}/{total}] [{ikon_f}{ikon_k}] {nama[:48]}")
            if lat:
                print(f"           lat={lat}, lng={lng}")
            else:
                print(f"           koordinat tidak dapat")

        except Exception as e:
            print(f"    Skip [{i+1}/{total}]: {e}")
            continue

    return results


# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────
async def main():
    start_idx = load_checkpoint()
    all_data  = []

    if os.path.exists(OUTPUT_FILE) and start_idx > 0:
        df_ex    = pd.read_excel(OUTPUT_FILE, engine="openpyxl")
        all_data = df_ex.to_dict("records")
        print(f"Resume checkpoint {start_idx}, data: {len(all_data)} baris\n")
    else:
        print(f"Mulai baru - {len(KEYWORDS)} keyword\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
        )
        ctx = await browser.new_context(
            viewport={"width": 1280, "height": 800},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            locale="id-ID",
        )
        page = await ctx.new_page()

        for i, kw in enumerate(KEYWORDS[start_idx:], start=start_idx):
            kat = next(
                (k for k in KATEGORI if kw.lower().startswith(k.lower())),
                KATEGORI[0]
            )
            kec = ekstrak_kecamatan(kw, kat)
            print(f"\n[{i+1}/{len(KEYWORDS)}] {kw}")

            data_kw = await scrape_keyword(page, kw, kec, kat, MAX_HASIL)
            all_data.extend(data_kw)

            if (i + 1) % 10 == 0:
                df_tmp = save_to_excel(all_data)
                save_checkpoint(i + 1)
                n_foto  = (df_tmp["foto_url"]  != "").sum()
                n_coord = df_tmp["latitude"].notna().sum()
                sisa    = len(KEYWORDS) - (i + 1)
                print(f"  Progress : {i+1}/{len(KEYWORDS)} keyword")
                print(f"  Data     : {len(df_tmp)} usaha")
                print(f"  Foto     : {n_foto} ({n_foto/len(df_tmp)*100:.1f}%)")
                print(f"  Koordinat: {n_coord} ({n_coord/len(df_tmp)*100:.1f}%)")
                print(f"  Sisa     : ~{sisa*5//60}h {sisa*5%60}m")

            jeda = random.uniform(*DELAY_ANTAR_KW)
            print(f"  Jeda {jeda:.1f}s...")
            await asyncio.sleep(jeda)

        await browser.close()

    df_final = save_to_excel(all_data)
    if os.path.exists(CHECKPOINT_FILE):
        os.remove(CHECKPOINT_FILE)

    tot     = len(df_final)
    n_foto  = (df_final["foto_url"]  != "").sum()
    n_coord = df_final["latitude"].notna().sum()

    print(f"\n{'='*55}")
    print(f"SELESAI!")
    print(f"  Total usaha  : {tot}")
    print(f"  Ada foto     : {n_foto} ({n_foto/tot*100:.1f}%)")
    print(f"  Koordinat OK : {n_coord} ({n_coord/tot*100:.1f}%)")
    print(f"  File output  : {OUTPUT_FILE}")
    print(f"{'='*55}")


if __name__ == "__main__":
    asyncio.run(main())