"""
SCRAPER USAHA MAKASSAR — LENGKAP + FOTO
=========================================
Coverage: 15 Kecamatan x 80 Kategori = 1.200 keyword

Kategori mencakup:
  - Kuliner & Minuman (23 kategori)
  - Retail & Toko (15 kategori)
  - Kesehatan & Kecantikan (10 kategori)
  - Jasa Umum (12 kategori)
  - Otomotif (8 kategori)
  - Pendidikan & Hiburan (8 kategori)
  - Properti & Akomodasi (4 kategori)

Output kolom:
  keyword | nama | kategori_usaha | rating | jumlah_ulasan
  alamat  | telepon | website | kecamatan | foto_url | tanggal
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
      "Panakkukang", "Manggala", "Biringkanaya", 
]

KATEGORI = [
    # ── KULINER & MINUMAN (23) ────────────────────────────
    "cafe",
    "coffee shop",
    "kedai kopi",
    "warung makan",
    "restoran",
    "seafood",
    "bakso",
    "mie ayam",
    "coto makassar",
    "konro",
    "pallubasa",
    "nasi goreng",
    "nasi kuning",
    "soto",
    "sate",
    "ayam geprek",
    "ayam bakar",
    "martabak",
    "pisang epe",
    "es pisang ijo",
    "burger",
    "pizza",
    "korean food",
    "japanese food",
    "chinese food",
    "dessert",
    "es krim",
    "minuman kekinian",
    "boba",
    "jus buah",
    "roti dan bakeri",
    "donat",
    "mie gacoan",
    "warteg",
    "rumah makan padang",

    # ── RETAIL & TOKO (15) ───────────────────────────────
    "minimarket",
    "supermarket",
    "toko kelontong",
    "toko pakaian",
    "toko baju muslim",
    "toko sepatu",
    "toko tas",
    "toko elektronik",
    "toko hp",
    "toko komputer",
    "toko bangunan",
    "toko buku",
    "toko mainan",
    "toko kosmetik",
    "toko aksesoris",
    "toko oleh oleh",
    "toko peralatan rumah",
    "toko furniture",

    # ── KESEHATAN & KECANTIKAN (10) ──────────────────────
    "apotek",
    "klinik",
    "klinik gigi",
    "klinik kecantikan",
    "rumah sakit",
    "puskesmas",
    "optik",
    "salon kecantikan",
    "barbershop",
    "spa",
    "nail art",
    "tukang cukur",

    # ── JASA UMUM (12) ───────────────────────────────────
    "laundry",
    "laundry kiloan",
    "percetakan",
    "foto studio",
    "jasa servis hp",
    "jasa jahit",
    "jasa cuci ac",
    "jasa servis elektronik",
    "travel agent",
    "ekspedisi pengiriman",
    "jasa kebersihan",
    "notaris",
    "kantor pos",

    # ── OTOMOTIF (8) ─────────────────────────────────────
    "bengkel motor",
    "bengkel mobil",
    "cuci motor",
    "cuci mobil",
    "tambal ban",
    "toko ban",
    "toko sparepart motor",
    "toko sparepart mobil",
    "modifikasi motor",

    # ── PENDIDIKAN & HIBURAN (8) ─────────────────────────
    "bimbel",
    "les privat",
    "kursus bahasa inggris",
    "kursus komputer",
    "taman kanak kanak",
    "gym fitness",
    "futsal",
    "kolam renang",
    "bioskop",
    "karaoke",
    "toko alat olahraga",
    "toko alat musik",

    # ── PROPERTI & AKOMODASI (4) ─────────────────────────
    "hotel",
    "penginapan",
    "kost",
    "guest house",

    # ── MAKANAN KHAS SULAWESI (bonus) ────────────────────
    "ikan bakar",
    "kapurung",
    "jalangkote",
    "sup konro",
    "mie titi",
    "es pallu butung",
]

# Hapus duplikat jaga-jaga
KATEGORI = list(dict.fromkeys(KATEGORI))

KEYWORDS        = [f"{kat} {kec} Makassar" for kec in KECAMATAN for kat in KATEGORI]
OUTPUT_FILE     = "data new/data_usaha_mar.xlsx"
CHECKPOINT_FILE = "data new/checkpoint_mar.txt"
MAX_HASIL       = 60
DELAY_ANTAR_KW  = (3, 7)

print(f"Total kategori : {len(KATEGORI)}")
print(f"Total keyword  : {len(KEYWORDS)} ({len(KECAMATAN)} kec x {len(KATEGORI)} kat)")
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
# AMBIL URL FOTO UTAMA
# ─────────────────────────────────────────────────────────────
async def ambil_foto_url(page) -> str:
    try:
        # Cara 1: img tag langsung di panel
        foto = await page.evaluate("""() => {
            const selectors = [
                'button.aoRNLd img',
                'div.aoRNLd img',
                'img.DaSXdd',
                'div.RZ66Rb img',
                'div.section-hero-header-image img',
            ];
            for (const sel of selectors) {
                const el = document.querySelector(sel);
                if (el && el.src && (el.src.includes('googleusercontent') || el.src.includes('ggpht'))) {
                    return el.src;
                }
            }
            // Fallback: semua img googleusercontent ukuran wajar
            const imgs = document.querySelectorAll(
                'img[src*="googleusercontent"], img[src*="ggpht"]'
            );
            for (const img of imgs) {
                const src = img.src || '';
                const rect = img.getBoundingClientRect();
                if (src && rect.width > 80 && rect.height > 60) return src;
            }
            // Fallback kedua: ambil yang pertama meski kecil
            for (const img of imgs) {
                if (img.src && img.src.length > 60) return img.src;
            }
            return null;
        }""")

        if foto and ("googleusercontent" in foto or "ggpht" in foto):
            # Resize ke 800x600 agar kualitas cukup bagus
            url_bersih = re.sub(r'=w\d+.*$', '=w800-h600-k-no', foto)
            return url_bersih

        # Cara 2: regex dari HTML source
        html = await page.content()
        m = re.search(
            r'(https://lh[0-9]\.googleusercontent\.com/p/[A-Za-z0-9_\-]+=w\d+[^\s"\'<>]*)',
            html
        )
        if m:
            url = m.group(1)
            url_bersih = re.sub(r'=w\d+.*$', '=w800-h600-k-no', url)
            return url_bersih

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

    # Scroll feed untuk load lebih banyak hasil
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
            await page.wait_for_selector("h1.DUwDvf", timeout=10000)
            await asyncio.sleep(random.uniform(1.2, 1.8))

            nama       = await safe_text(page, "h1.DUwDvf")
            kat_usaha  = await safe_text(page, "button.DkEaL")
            rating     = await safe_text(page, "div.F7nice span[aria-hidden='true']")
            ulasan_raw = await safe_text(page, "div.F7nice span[aria-label*='ulasan']")
            alamat     = await safe_text(page, "button[data-item-id='address'] .fontBodyMedium")
            telp       = await safe_text(page, "button[data-item-id^='phone'] .fontBodyMedium")
            website    = await safe_text(page, "a[data-item-id='authority'] .fontBodyMedium")
            foto_url   = await ambil_foto_url(page)

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
                "foto_url"      : foto_url,
                "tanggal"       : datetime.now().strftime("%Y-%m-%d"),
            })

            foto_ikon = "F" if foto_url else "-"
            print(f"    [{i+1}/{total}] [{foto_ikon}] {nama[:50]}")

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

            # Auto-save & statistik setiap 10 keyword
            if (i + 1) % 10 == 0:
                df_tmp = save_to_excel(all_data)
                save_checkpoint(i + 1)
                n_foto = (df_tmp["foto_url"] != "").sum()
                pct    = n_foto / len(df_tmp) * 100 if len(df_tmp) else 0
                sisa   = len(KEYWORDS) - (i + 1)
                print(f"  Progress : {i+1}/{len(KEYWORDS)} keyword")
                print(f"  Data     : {len(df_tmp)} usaha")
                print(f"  Foto     : {n_foto} ({pct:.1f}%)")
                print(f"  Sisa     : ~{sisa*5//60}h {sisa*5%60}m")

            jeda = random.uniform(*DELAY_ANTAR_KW)
            print(f"  Jeda {jeda:.1f}s...")
            await asyncio.sleep(jeda)

        await browser.close()

    df_final = save_to_excel(all_data)
    if os.path.exists(CHECKPOINT_FILE):
        os.remove(CHECKPOINT_FILE)

    tot    = len(df_final)
    n_foto = (df_final["foto_url"] != "").sum()

    print(f"\n{'='*55}")
    print(f"SELESAI!")
    print(f"  Total usaha  : {tot}")
    print(f"  Ada foto     : {n_foto} ({n_foto/tot*100:.1f}%)")
    print(f"  File output  : {OUTPUT_FILE}")
    print(f"{'='*55}")


if __name__ == "__main__":
    asyncio.run(main())